from __future__ import annotations

from io import BytesIO
from typing import Any
import math
import re
import unicodedata

import numpy as np
import pandas as pd


ALL_VALUES = {None, "", "Todos", "Todas"}
METRIC_COLUMNS = ["Consumo", "Valor", "Preço Unitário", "R$/L"]


def _norm(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _column_key(value: Any) -> str:
    text = unicodedata.normalize("NFKD", _norm(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _identifier(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)) and math.isfinite(float(value)):
        if float(value).is_integer():
            return str(int(value))
    return _norm(value)


def _padded_identifier(value: Any, width: int) -> str:
    identifier = _identifier(value)
    return identifier.zfill(width) if identifier.isdigit() else identifier


def _month_key(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m")
    if hasattr(value, "year") and hasattr(value, "month"):
        return f"{value.year:04d}-{value.month:02d}"
    text = str(value).strip()
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%Y-%m")
    match = re.match(
        r"(?i)^\s*(janeiro|fevereiro|março|marco|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)[/ -](\d{2}|\d{4})",
        text,
    )
    if match:
        month_numbers = {
            "janeiro": "01", "fevereiro": "02", "março": "03", "marco": "03",
            "abril": "04", "maio": "05", "junho": "06", "julho": "07",
            "agosto": "08", "setembro": "09", "outubro": "10",
            "novembro": "11", "dezembro": "12",
        }
        year = match.group(2)
        year = f"20{year}" if len(year) == 2 else year
        return f"{year}-{month_numbers[match.group(1).lower()]}"
    return text


def _display_month(key: str) -> str:
    if not key:
        return ""
    try:
        return pd.to_datetime(key + "-01").strftime("%b/%Y").capitalize()
    except Exception:
        return key


def _numeric(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)
    text = series.astype(str).str.strip()
    both = text.str.contains(",", regex=False) & text.str.contains(".", regex=False)
    text.loc[both] = text.loc[both].str.replace(".", "", regex=False)
    text = text.str.replace(",", ".", regex=False)
    return pd.to_numeric(text, errors="coerce").fillna(0.0)


def _combine_labels(code: pd.Series, description: pd.Series, fallback: str) -> pd.Series:
    codes = code.map(_identifier)
    descriptions = description.map(_norm)
    labels = []
    for item_code, item_description in zip(codes, descriptions):
        if item_code and item_description:
            labels.append(f"{item_code} — {item_description}")
        elif item_description:
            labels.append(item_description)
        elif item_code:
            labels.append(item_code)
        else:
            labels.append(fallback)
    return pd.Series(labels, index=code.index, dtype="object")


def _read_volume(xls: pd.ExcelFile, sheets: list[str]) -> tuple[pd.DataFrame, str | None]:
    volume_sheet = "Consumo" if "Consumo" in sheets else None
    volume = pd.DataFrame(columns=["Mês", "Volume Processado"])
    if not volume_sheet:
        return volume, volume_sheet
    raw = pd.read_excel(xls, sheet_name=volume_sheet)
    raw.columns = [_norm(column) for column in raw.columns]
    if {"Mês", "Volume Processado"}.issubset(raw.columns):
        volume = raw[["Mês", "Volume Processado"]].copy()
        volume["Mês"] = volume["Mês"].map(_month_key)
        volume["Volume Processado"] = _numeric(volume["Volume Processado"])
        volume = (
            volume[volume["Mês"].ne("")]
            .groupby("Mês", as_index=False)["Volume Processado"]
            .sum()
        )
    return volume, volume_sheet


def _detect_transaction_header(xls: pd.ExcelFile, sheet_name: str) -> int:
    preview = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=25)
    required = {"descrgrupo", "descrevento", "dtdigitacao", "vlrtotal"}
    for row_index, row in preview.iterrows():
        keys = {_column_key(value) for value in row.tolist() if not pd.isna(value)}
        if required.issubset(keys):
            return int(row_index)
    raise ValueError(f"Não foi possível localizar o cabeçalho transacional na aba '{sheet_name}'.")


def _read_transaction_fact(
    xls: pd.ExcelFile,
    sheet_name: str,
) -> tuple[pd.DataFrame, dict[str, int | str]]:
    header_row = _detect_transaction_header(xls, sheet_name)
    raw = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)
    raw.columns = [_norm(column) for column in raw.columns]
    column_by_key = {_column_key(column): column for column in raw.columns}

    def source(*aliases: str) -> pd.Series:
        for alias in aliases:
            column = column_by_key.get(_column_key(alias))
            if column is not None:
                return raw[column]
        return pd.Series([""] * len(raw), index=raw.index, dtype="object")

    required_aliases = {
        "Descr Grupo": ("Descr Grupo",),
        "Descr Evento": ("Descr Evento",),
        "Descr Sub Ev": ("Descr Sub Ev",),
        "Produto": ("Produto",),
        "Vlr.Total": ("Vlr.Total", "Vlr Total"),
        "DT Digitacao": ("DT Digitacao", "DT Digitação"),
        "Filial": ("Filial",),
    }
    missing = [
        label for label, aliases in required_aliases.items()
        if not any(_column_key(alias) in column_by_key for alias in aliases)
    ]
    if missing:
        raise ValueError(
            f"A aba '{sheet_name}' não contém os campos obrigatórios: " + ", ".join(missing)
        )

    reference_date = pd.to_datetime(
        source("DT Digitacao", "DT Digitação"), dayfirst=True, errors="coerce"
    )
    valid_date = reference_date.notna()
    fact = pd.DataFrame(index=raw.index)
    fact["Data Digitação"] = reference_date
    fact["Mês"] = reference_date.dt.strftime("%Y-%m")
    fact["Filial"] = source("Filial").map(_identifier).replace("", "Não informada")
    fact["Tipo de Custo"] = source("Descr Tp Cus").map(_norm).replace("", "Não informado")
    fact["Grupo"] = source("Descr Grupo").map(_norm).replace("", "Não informado")
    fact["Linha"] = source("Descr Evento").map(_norm).replace("", "Não informado")
    fact["Subevento"] = source("Descr Sub Ev").map(_norm).replace("", "Não informado")
    fact["Código Produto"] = source("Produto").map(lambda value: _padded_identifier(value, 11))
    fact["Descrição Produto"] = source("Desc Produt", "Descricao Produto").map(_norm)
    fact["Produto"] = _combine_labels(
        fact["Código Produto"], fact["Descrição Produto"], "Produto não informado"
    )
    fact["Código Fornecedor"] = source("Cod Forn/Clente", "Cod Forn Cliente").map(_identifier)
    fact["Fornecedor"] = _combine_labels(
        fact["Código Fornecedor"], source("Nome").map(_norm), "Fornecedor não informado"
    )
    fact["Documento"] = source("Documento").map(lambda value: _padded_identifier(value, 9))
    fact["Sequência"] = source("Num.Sequenc.", "Num Sequenc").map(_identifier)
    fact["Pagamento"] = _combine_labels(
        fact["Documento"], fact["Sequência"], "Pagamento sem identificação"
    )
    fact["Centro de Custo"] = source("C Custo").map(lambda value: _padded_identifier(value, 8))
    fact["Descrição Centro de Custo"] = source("Descr CC").map(_norm)
    fact["Código Fiscal"] = source("Cod. Fiscal", "Cod Fiscal").map(_identifier)
    fact["Pedido Compra"] = source("Numero PC").map(_identifier)
    fact["Observação"] = source("Observacao", "Observação").map(_norm)
    fact["Observação PC"] = source("Observacao PC", "Observação PC").map(_norm)
    fact["Usuário"] = source("Usuario", "Usuário").map(_norm)
    fact["Consumo"] = _numeric(source("Quantidade"))
    fact["Valor"] = _numeric(source("Vlr.Total", "Vlr Total"))
    fact = fact[valid_date].copy()
    fact = fact.sort_values(["Data Digitação", "Grupo", "Linha", "Subevento", "Produto"])
    quality = {
        "schema": "transacional",
        "header_row": header_row + 1,
        "source_rows": int(len(raw)),
        "valid_rows": int(len(fact)),
        "invalid_date_rows": int((~valid_date).sum()),
    }
    return fact, quality


def _read_legacy_fact(
    xls: pd.ExcelFile,
    sheet_name: str,
) -> tuple[pd.DataFrame, dict[str, int | str]]:
    raw = pd.read_excel(xls, sheet_name=sheet_name)
    raw.columns = [_norm(column) for column in raw.columns]
    required = {"Mês", "Grupo", "Linha", "Consumo", "Valor"}
    missing = required - set(raw.columns)
    if missing:
        raise ValueError(
            f"A aba '{sheet_name}' não contém as colunas obrigatórias: "
            + ", ".join(sorted(missing))
        )
    fact = raw[["Mês", "Grupo", "Linha", "Consumo", "Valor"]].copy()
    fact["Mês"] = fact["Mês"].map(_month_key)
    fact["Grupo"] = fact["Grupo"].map(_norm).replace("", "Não informado")
    fact["Linha"] = fact["Linha"].map(_norm).replace("", "Não informado")
    fact["Consumo"] = _numeric(fact["Consumo"])
    fact["Valor"] = _numeric(fact["Valor"])
    fact = fact[fact["Mês"].ne("")].copy()
    fact = fact.groupby(["Mês", "Grupo", "Linha"], as_index=False).agg(
        Consumo=("Consumo", "sum"), Valor=("Valor", "sum")
    )
    fact["Data Digitação"] = pd.to_datetime(fact["Mês"] + "-01", errors="coerce")
    fact["Filial"] = "Não informada"
    fact["Tipo de Custo"] = "Não detalhado"
    fact["Subevento"] = "Não detalhado"
    fact["Código Produto"] = ""
    fact["Descrição Produto"] = "Não detalhado"
    fact["Produto"] = "Não detalhado"
    fact["Código Fornecedor"] = ""
    fact["Fornecedor"] = "Não detalhado"
    fact["Documento"] = ""
    fact["Sequência"] = ""
    fact["Pagamento"] = "Dado compilado"
    fact["Centro de Custo"] = ""
    fact["Descrição Centro de Custo"] = ""
    fact["Código Fiscal"] = ""
    fact["Pedido Compra"] = ""
    fact["Observação"] = ""
    fact["Observação PC"] = ""
    fact["Usuário"] = ""
    quality = {
        "schema": "compilado", "header_row": 1, "source_rows": int(len(raw)),
        "valid_rows": int(len(fact)), "invalid_date_rows": 0,
    }
    return fact, quality


def _add_metrics(fact: pd.DataFrame) -> pd.DataFrame:
    fact = fact.copy()
    fact["Preço Unitário"] = np.where(
        fact["Consumo"].abs() > 1e-12, fact["Valor"] / fact["Consumo"], np.nan
    )
    fact["R$/L"] = np.where(
        fact["Volume Processado"].abs() > 1e-12,
        fact["Valor"] / fact["Volume Processado"],
        np.nan,
    )
    return fact


def _read_excel(uploaded_bytes: bytes):
    xls = pd.ExcelFile(BytesIO(uploaded_bytes), engine="openpyxl")
    sheets = xls.sheet_names
    volume, volume_sheet = _read_volume(xls, sheets)
    if "2-Custos" in sheets:
        fact_sheet = "2-Custos"
        fact, quality = _read_transaction_fact(xls, fact_sheet)
    elif "Custos em linha" in sheets:
        fact_sheet = "Custos em linha"
        fact, quality = _read_legacy_fact(xls, fact_sheet)
    else:
        candidates = [sheet for sheet in sheets if sheet != volume_sheet]
        if not candidates:
            raise ValueError("Nenhuma aba de custos foi encontrada.")
        fact_sheet = candidates[0]
        try:
            fact, quality = _read_transaction_fact(xls, fact_sheet)
        except ValueError:
            fact, quality = _read_legacy_fact(xls, fact_sheet)
    fact = fact.merge(volume, on="Mês", how="left")
    fact["Volume Processado"] = fact["Volume Processado"].fillna(0.0)
    return _add_metrics(fact), volume, fact_sheet, volume_sheet, quality


def load_data(uploaded_bytes: bytes):
    fact, volume, fact_sheet, volume_sheet, quality = _read_excel(uploaded_bytes)

    def unique_values(column: str) -> list[str]:
        return sorted(value for value in fact[column].dropna().astype(str).unique() if value)

    return {
        "fact": fact, "volume": volume, "fact_sheet": fact_sheet,
        "volume_sheet": volume_sheet, "quality": quality, "schema": quality["schema"],
        "months": unique_values("Mês"), "filiais": unique_values("Filial"),
        "groups": unique_values("Grupo"), "lines": unique_values("Linha"),
        "subevents": unique_values("Subevento"), "products": unique_values("Produto"),
        "suppliers": unique_values("Fornecedor"),
    }


def filter_fact(
    fact: pd.DataFrame,
    filial: str | None = None,
    group: str | None = None,
    line: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
    supplier: str | None = None,
) -> pd.DataFrame:
    df = fact.copy()
    filters = {
        "Filial": filial, "Grupo": group, "Linha": line,
        "Subevento": subevent, "Produto": product, "Fornecedor": supplier,
    }
    for column, selected in filters.items():
        if selected not in ALL_VALUES and column in df.columns:
            df = df[df[column].eq(selected)]
    return df


def aggregate(
    fact: pd.DataFrame,
    filial: str | None = None,
    group: str | None = None,
    line: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
) -> pd.DataFrame:
    scoped = filter_fact(fact, filial, group, line, subevent, product)
    calendar = (
        fact.groupby("Mês", as_index=False)
        .agg(Volume=("Volume Processado", "max"))
        .sort_values("Mês")
    )
    totals = scoped.groupby("Mês", as_index=False).agg(
        Consumo=("Consumo", "sum"), Valor=("Valor", "sum")
    )
    out = calendar.merge(totals, on="Mês", how="left")
    out[["Consumo", "Valor"]] = out[["Consumo", "Valor"]].fillna(0.0)
    out["Preço Unitário"] = np.where(
        out["Consumo"].abs() > 1e-12, out["Valor"] / out["Consumo"], np.nan
    )
    out["R$/L"] = np.where(
        out["Volume"].abs() > 1e-12, out["Valor"] / out["Volume"], np.nan
    )
    out["Mês Label"] = out["Mês"].map(_display_month)
    return out


def monthly_kpis(
    scoped_fact: pd.DataFrame,
    month: str,
    full_fact: pd.DataFrame | None = None,
) -> dict[str, float]:
    df = scoped_fact[scoped_fact["Mês"] == month]
    volume_source = full_fact if full_fact is not None else scoped_fact
    volume_month = volume_source[volume_source["Mês"] == month]
    cost = float(df["Valor"].sum())
    consumption = float(df["Consumo"].sum())
    volume = float(volume_month["Volume Processado"].max()) if not volume_month.empty else 0.0
    return {
        "cost": cost, "consumption": consumption, "volume": volume,
        "unit_price": cost / consumption if abs(consumption) > 1e-12 else math.nan,
        "per_liter": cost / volume if abs(volume) > 1e-12 else math.nan,
    }


def _months_for_year(fact: pd.DataFrame, reference_month: str) -> list[str]:
    months = sorted(fact["Mês"].dropna().astype(str).unique())
    match = re.match(r"^(\d{4})-\d{2}$", str(reference_month))
    if not match:
        return months
    return [month for month in months if month.startswith(match.group(1) + "-")]


def annual_average_kpis(
    scoped_fact: pd.DataFrame,
    reference_month: str,
    full_fact: pd.DataFrame | None = None,
) -> dict[str, float]:
    calendar_fact = full_fact if full_fact is not None else scoped_fact
    months = _months_for_year(calendar_fact, reference_month)
    if not months:
        return {
            "cost": 0.0, "consumption": 0.0, "volume": 0.0,
            "unit_price": math.nan, "per_liter": math.nan,
        }
    costs, consumptions, volumes = [], [], []
    for month in months:
        month_scope = scoped_fact[scoped_fact["Mês"] == month]
        month_full = calendar_fact[calendar_fact["Mês"] == month]
        costs.append(float(month_scope["Valor"].sum()))
        consumptions.append(float(month_scope["Consumo"].sum()))
        volumes.append(
            float(month_full["Volume Processado"].max()) if not month_full.empty else 0.0
        )
    cost = float(np.mean(costs))
    consumption = float(np.mean(consumptions))
    volume = float(np.mean(volumes))
    return {
        "cost": cost, "consumption": consumption, "volume": volume,
        "unit_price": cost / consumption if abs(consumption) > 1e-12 else math.nan,
        "per_liter": cost / volume if abs(volume) > 1e-12 else math.nan,
    }


def snapshot_by(
    fact: pd.DataFrame,
    month: str,
    dimensions: str | list[str],
    filial: str | None = None,
    group: str | None = None,
    line: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
    supplier: str | None = None,
) -> pd.DataFrame:
    dims = [dimensions] if isinstance(dimensions, str) else list(dimensions)
    df = fact[fact["Mês"].eq(month)]
    df = filter_fact(df, filial, group, line, subevent, product, supplier)
    columns = dims + [
        "Consumo", "Valor", "Volume", "Pagamentos", "Preço Unitário", "R$/L", "Participação %"
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)
    out = df.groupby(dims, as_index=False).agg(
        Consumo=("Consumo", "sum"), Valor=("Valor", "sum"),
        Volume=("Volume Processado", "max"), Pagamentos=("Valor", "size"),
    )
    out["Preço Unitário"] = np.where(
        out["Consumo"].abs() > 1e-12, out["Valor"] / out["Consumo"], np.nan
    )
    out["R$/L"] = np.where(
        out["Volume"].abs() > 1e-12, out["Valor"] / out["Volume"], np.nan
    )
    total = float(out["Valor"].sum())
    out["Participação %"] = np.where(
        abs(total) > 1e-12, out["Valor"] / total * 100, 0.0
    )
    return out.sort_values("Valor", ascending=False)


def annual_average_by(
    fact: pd.DataFrame,
    reference_month: str,
    dimensions: str | list[str],
    filial: str | None = None,
    group: str | None = None,
    line: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
    supplier: str | None = None,
) -> pd.DataFrame:
    dims = [dimensions] if isinstance(dimensions, str) else list(dimensions)
    months = _months_for_year(fact, reference_month)
    columns = dims + [
        "Consumo", "Valor", "Volume", "Pagamentos", "Preço Unitário", "R$/L", "Participação %"
    ]
    if not months:
        return pd.DataFrame(columns=columns)
    scoped = fact[fact["Mês"].isin(months)]
    scoped = filter_fact(scoped, filial, group, line, subevent, product, supplier)
    if scoped.empty:
        return pd.DataFrame(columns=columns)
    members = scoped[dims].drop_duplicates().copy()
    calendar = pd.DataFrame({"Mês": months})
    members["_key"] = 1
    calendar["_key"] = 1
    grid = members.merge(calendar, on="_key").drop(columns="_key")
    monthly = scoped.groupby(["Mês"] + dims, as_index=False).agg(
        Consumo=("Consumo", "sum"), Valor=("Valor", "sum"), Pagamentos=("Valor", "size")
    )
    grid = grid.merge(monthly, on=["Mês"] + dims, how="left")
    grid[["Consumo", "Valor", "Pagamentos"]] = grid[
        ["Consumo", "Valor", "Pagamentos"]
    ].fillna(0.0)
    out = grid.groupby(dims, as_index=False).agg(
        Consumo=("Consumo", "mean"), Valor=("Valor", "mean"),
        Pagamentos=("Pagamentos", "mean"),
    )
    volume = annual_average_kpis(fact, reference_month, fact)["volume"]
    out["Volume"] = volume
    out["Preço Unitário"] = np.where(
        out["Consumo"].abs() > 1e-12, out["Valor"] / out["Consumo"], np.nan
    )
    out["R$/L"] = np.where(abs(volume) > 1e-12, out["Valor"] / volume, np.nan)
    total = float(out["Valor"].sum())
    out["Participação %"] = np.where(
        abs(total) > 1e-12, out["Valor"] / total * 100, 0.0
    )
    return out.sort_values("Valor", ascending=False)


def comparison_snapshot(
    fact: pd.DataFrame,
    current_month: str,
    base_period: str,
    dimensions: str | list[str],
    annual: bool = False,
    filial: str | None = None,
    group: str | None = None,
    line: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
    supplier: str | None = None,
) -> pd.DataFrame:
    dims = [dimensions] if isinstance(dimensions, str) else list(dimensions)
    kwargs = dict(
        filial=filial, group=group, line=line, subevent=subevent,
        product=product, supplier=supplier,
    )
    current = snapshot_by(fact, current_month, dims, **kwargs)
    base = (
        annual_average_by(fact, current_month, dims, **kwargs)
        if annual else snapshot_by(fact, base_period, dims, **kwargs)
    )
    current = current[dims + METRIC_COLUMNS].rename(
        columns={metric: f"{metric} Atual" for metric in METRIC_COLUMNS}
    )
    base = base[dims + METRIC_COLUMNS].rename(
        columns={metric: f"{metric} Base" for metric in METRIC_COLUMNS}
    )
    out = current.merge(base, on=dims, how="outer")
    for metric in METRIC_COLUMNS:
        current_column = f"{metric} Atual"
        base_column = f"{metric} Base"
        out[current_column] = pd.to_numeric(out[current_column], errors="coerce").fillna(0.0)
        out[base_column] = pd.to_numeric(out[base_column], errors="coerce").fillna(0.0)
        out[f"Delta {metric}"] = out[current_column] - out[base_column]
    out["Delta Valor %"] = np.where(
        out["Valor Base"].abs() > 1e-12,
        out["Delta Valor"] / out["Valor Base"].abs() * 100,
        np.nan,
    )
    return out.sort_values("Delta Valor", ascending=False)


def group_snapshot(fact: pd.DataFrame, month: str, filial: str | None = None) -> pd.DataFrame:
    return snapshot_by(fact, month, "Grupo", filial=filial)


def line_snapshot(
    fact: pd.DataFrame,
    month: str,
    group: str | None = None,
    filial: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
) -> pd.DataFrame:
    return snapshot_by(
        fact, month, ["Grupo", "Linha"], filial=filial, group=group,
        subevent=subevent, product=product,
    )


def annual_average_by_line(
    fact: pd.DataFrame,
    reference_month: str,
    group: str | None = None,
    line: str | None = None,
    filial: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
) -> pd.DataFrame:
    return annual_average_by(
        fact, reference_month, ["Grupo", "Linha"], filial=filial,
        group=group, line=line, subevent=subevent, product=product,
    )


def transaction_detail(
    fact: pd.DataFrame,
    month: str,
    filial: str | None = None,
    group: str | None = None,
    line: str | None = None,
    subevent: str | None = None,
    product: str | None = None,
    supplier: str | None = None,
) -> pd.DataFrame:
    df = fact[fact["Mês"].eq(month)]
    df = filter_fact(df, filial, group, line, subevent, product, supplier)
    columns = [
        "Data Digitação", "Filial", "Grupo", "Linha", "Subevento",
        "Código Produto", "Descrição Produto", "Fornecedor", "Documento",
        "Centro de Custo", "Descrição Centro de Custo", "Consumo", "Valor",
        "Preço Unitário", "Pedido Compra", "Observação", "Observação PC",
    ]
    available = [column for column in columns if column in df.columns]
    return df[available].sort_values(["Valor", "Data Digitação"], ascending=[False, False])


def months_label_map(months: list[str]) -> dict[str, str]:
    return {month: _display_month(month) for month in months}
