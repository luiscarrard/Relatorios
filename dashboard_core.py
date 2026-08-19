
from __future__ import annotations

from io import BytesIO
from typing import Any
import math
import re

import pandas as pd
import numpy as np


def _norm(s: Any) -> str:
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _month_key(value: Any) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp,)):
        return value.strftime("%Y-%m")
    if hasattr(value, "year") and hasattr(value, "month"):
        return f"{value.year:04d}-{value.month:02d}"
    s = str(value).strip()
    parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%Y-%m")
    m = re.match(r"(?i)^\s*(janeiro|fevereiro|março|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)[/ -](\d{2}|\d{4})", s)
    if m:
        months = {
            "janeiro":"01","fevereiro":"02","março":"03","abril":"04","maio":"05","junho":"06",
            "julho":"07","agosto":"08","setembro":"09","outubro":"10","novembro":"11","dezembro":"12"
        }
        y = m.group(2)
        y = f"20{y}" if len(y) == 2 else y
        return f"{y}-{months[m.group(1).lower()]}"
    return s


def _display_month(key: str) -> str:
    if not key:
        return ""
    try:
        dt = pd.to_datetime(key + "-01")
        return dt.strftime("%b/%Y").capitalize()
    except Exception:
        return key


def _read_excel(uploaded_bytes: bytes):
    xls = pd.ExcelFile(BytesIO(uploaded_bytes), engine="openpyxl")
    sheets = xls.sheet_names

    fact_sheet = "Custos em linha" if "Custos em linha" in sheets else sheets[0]
    volume_sheet = "Consumo" if "Consumo" in sheets else None

    fact = pd.read_excel(xls, sheet_name=fact_sheet)
    fact.columns = [_norm(c) for c in fact.columns]

    required = {"Mês", "Grupo", "Linha", "Consumo", "Valor"}
    missing = required - set(fact.columns)
    if missing:
        raise ValueError(f"A aba '{fact_sheet}' não contém as colunas obrigatórias: {', '.join(sorted(missing))}")

    fact = fact[["Mês", "Grupo", "Linha", "Consumo", "Valor"]].copy()
    fact["Mês"] = fact["Mês"].map(_month_key)
    fact["Grupo"] = fact["Grupo"].map(_norm)
    fact["Linha"] = fact["Linha"].map(_norm)
    fact["Consumo"] = pd.to_numeric(fact["Consumo"], errors="coerce").fillna(0.0)
    fact["Valor"] = pd.to_numeric(fact["Valor"], errors="coerce").fillna(0.0)

    # Remove cabeçalhos/linhas vazias e linhas sem identificação.
    fact = fact[
        fact["Mês"].ne("") &
        fact["Grupo"].ne("") &
        fact["Linha"].ne("")
    ].copy()

    volume = pd.DataFrame(columns=["Mês", "Volume Processado"])
    if volume_sheet:
        volume = pd.read_excel(xls, sheet_name=volume_sheet)
        volume.columns = [_norm(c) for c in volume.columns]
        if "Mês" in volume.columns and "Volume Processado" in volume.columns:
            volume = volume[["Mês", "Volume Processado"]].copy()
            volume["Mês"] = volume["Mês"].map(_month_key)
            volume["Volume Processado"] = pd.to_numeric(
                volume["Volume Processado"], errors="coerce"
            ).fillna(0.0)
            volume = volume.groupby("Mês", as_index=False)["Volume Processado"].sum()

    fact = fact.groupby(
        ["Mês", "Grupo", "Linha"], as_index=False
    ).agg(Consumo=("Consumo", "sum"), Valor=("Valor", "sum"))

    # Junta volume processado.
    fact = fact.merge(volume, on="Mês", how="left")
    fact["Volume Processado"] = fact["Volume Processado"].fillna(0.0)

    return fact, volume, fact_sheet, volume_sheet


def load_data(uploaded_bytes: bytes):
    fact, volume, fact_sheet, volume_sheet = _read_excel(uploaded_bytes)

    fact["Preço Unitário"] = np.where(
        fact["Consumo"].abs() > 1e-12,
        fact["Valor"] / fact["Consumo"],
        np.nan
    )
    fact["R$/L"] = np.where(
        fact["Volume Processado"].abs() > 1e-12,
        fact["Valor"] / fact["Volume Processado"],
        np.nan
    )

    return {
        "fact": fact,
        "volume": volume,
        "fact_sheet": fact_sheet,
        "volume_sheet": volume_sheet,
        "months": sorted(fact["Mês"].unique()),
        "groups": sorted(fact["Grupo"].unique()),
        "lines": sorted(fact["Linha"].unique()),
    }


def aggregate(
    fact: pd.DataFrame,
    group: str | None = None,
    line: str | None = None,
) -> pd.DataFrame:
    df = fact.copy()
    if group and group != "Todos":
        df = df[df["Grupo"] == group]
    if line and line != "Todos":
        df = df[df["Linha"] == line]

    out = (
        df.groupby("Mês", as_index=False)
        .agg(
            Consumo=("Consumo", "sum"),
            Valor=("Valor", "sum"),
            Volume=("Volume Processado", "max"),
        )
        .sort_values("Mês")
    )
    out["Preço Unitário"] = np.where(
        out["Consumo"].abs() > 1e-12,
        out["Valor"] / out["Consumo"],
        np.nan
    )
    out["R$/L"] = np.where(
        out["Volume"].abs() > 1e-12,
        out["Valor"] / out["Volume"],
        np.nan
    )
    out["Mês Label"] = out["Mês"].map(_display_month)
    return out


def monthly_kpis(fact: pd.DataFrame, month: str) -> dict[str, float]:
    df = fact[fact["Mês"] == month]
    cost = float(df["Valor"].sum())
    consumption = float(df["Consumo"].sum())
    volume = float(df["Volume Processado"].max()) if not df.empty else 0.0
    unit_price = cost / consumption if abs(consumption) > 1e-12 else math.nan
    per_liter = cost / volume if abs(volume) > 1e-12 else math.nan
    return {
        "cost": cost,
        "consumption": consumption,
        "volume": volume,
        "unit_price": unit_price,
        "per_liter": per_liter,
    }


def _months_for_year(fact: pd.DataFrame, reference_month: str) -> list[str]:
    """Return the available months in the same year as the reference month."""
    months = sorted(fact["Mês"].dropna().astype(str).unique())
    match = re.match(r"^(\d{4})-\d{2}$", str(reference_month))
    if not match:
        return months
    year = match.group(1)
    return [month for month in months if month.startswith(f"{year}-")]


def annual_average_kpis(
    scoped_fact: pd.DataFrame,
    reference_month: str,
    full_fact: pd.DataFrame | None = None,
) -> dict[str, float]:
    """Average monthly KPIs for the selected scope and reference year."""
    calendar_fact = full_fact if full_fact is not None else scoped_fact
    months = _months_for_year(calendar_fact, reference_month)
    if not months:
        return {
            "cost": 0.0,
            "consumption": 0.0,
            "volume": 0.0,
            "unit_price": math.nan,
            "per_liter": math.nan,
        }

    monthly_costs = []
    monthly_consumption = []
    monthly_volumes = []

    for month in months:
        scoped_month = scoped_fact[scoped_fact["Mês"] == month]
        full_month = calendar_fact[calendar_fact["Mês"] == month]
        monthly_costs.append(float(scoped_month["Valor"].sum()))
        monthly_consumption.append(float(scoped_month["Consumo"].sum()))
        monthly_volumes.append(
            float(full_month["Volume Processado"].max())
            if not full_month.empty else 0.0
        )

    cost = float(np.mean(monthly_costs))
    consumption = float(np.mean(monthly_consumption))
    volume = float(np.mean(monthly_volumes))
    unit_price = cost / consumption if abs(consumption) > 1e-12 else math.nan
    per_liter = cost / volume if abs(volume) > 1e-12 else math.nan

    return {
        "cost": cost,
        "consumption": consumption,
        "volume": volume,
        "unit_price": unit_price,
        "per_liter": per_liter,
    }


def annual_average_by_line(
    fact: pd.DataFrame,
    reference_month: str,
    group: str | None = None,
    line: str | None = None,
) -> pd.DataFrame:
    """Average monthly consumption and cost by line in the reference year."""
    months = _months_for_year(fact, reference_month)
    columns = ["Grupo", "Linha", "Consumo", "Valor"]
    if not months:
        return pd.DataFrame(columns=columns)

    df = fact[fact["Mês"].isin(months)].copy()
    if group and group != "Todos":
        df = df[df["Grupo"] == group]
    if line and line != "Todos":
        df = df[df["Linha"] == line]
    if df.empty:
        return pd.DataFrame(columns=columns)

    lines = df[["Grupo", "Linha"]].drop_duplicates().copy()
    calendar = pd.DataFrame({"Mês": months})
    lines["_key"] = 1
    calendar["_key"] = 1
    grid = lines.merge(calendar, on="_key").drop(columns="_key")

    monthly = (
        df.groupby(["Mês", "Grupo", "Linha"], as_index=False)
        .agg(Consumo=("Consumo", "sum"), Valor=("Valor", "sum"))
    )
    grid = grid.merge(monthly, on=["Mês", "Grupo", "Linha"], how="left")
    grid[["Consumo", "Valor"]] = grid[["Consumo", "Valor"]].fillna(0.0)

    return (
        grid.groupby(["Grupo", "Linha"], as_index=False)
        .agg(Consumo=("Consumo", "mean"), Valor=("Valor", "mean"))
    )


def compare_months(
    fact: pd.DataFrame,
    current_month: str,
    base_month: str,
    group: str | None = None,
    line: str | None = None,
) -> dict[str, Any]:
    a = monthly_kpis(
        fact[(fact["Grupo"] == group)] if group and group != "Todos" else fact,
        current_month
    ) if not line else monthly_kpis(fact[fact["Linha"] == line], current_month)
    b = monthly_kpis(
        fact[(fact["Grupo"] == group)] if group and group != "Todos" else fact,
        base_month
    ) if not line else monthly_kpis(fact[fact["Linha"] == line], base_month)

    delta = a["cost"] - b["cost"]
    q_impact = (a["consumption"] - b["consumption"]) * (b["unit_price"] if math.isfinite(b["unit_price"]) else 0.0)
    p_impact = a["consumption"] * (
        (a["unit_price"] - b["unit_price"])
        if math.isfinite(a["unit_price"]) and math.isfinite(b["unit_price"])
        else 0.0
    )
    interaction = delta - q_impact - p_impact

    return {
        "current": a,
        "base": b,
        "delta": delta,
        "delta_pct": delta / abs(b["cost"]) if abs(b["cost"]) > 1e-12 else math.nan,
        "quantity_impact": q_impact,
        "price_impact": p_impact,
        "interaction": interaction,
    }


def group_snapshot(fact: pd.DataFrame, month: str) -> pd.DataFrame:
    df = fact[fact["Mês"] == month].copy()
    out = (
        df.groupby("Grupo", as_index=False)
        .agg(Consumo=("Consumo", "sum"), Valor=("Valor", "sum"), Volume=("Volume Processado", "max"))
    )
    out["R$/L"] = np.where(out["Volume"].abs() > 1e-12, out["Valor"] / out["Volume"], np.nan)
    out["Preço Unitário"] = np.where(out["Consumo"].abs() > 1e-12, out["Valor"] / out["Consumo"], np.nan)
    total = out["Valor"].sum()
    out["Participação %"] = np.where(abs(total) > 1e-12, out["Valor"] / total * 100, 0)
    return out.sort_values("Valor", ascending=False)


def line_snapshot(fact: pd.DataFrame, month: str, group: str | None = None) -> pd.DataFrame:
    df = fact[fact["Mês"] == month].copy()
    if group and group != "Todos":
        df = df[df["Grupo"] == group]
    out = (
        df.groupby(["Grupo", "Linha"], as_index=False)
        .agg(Consumo=("Consumo", "sum"), Valor=("Valor", "sum"), Volume=("Volume Processado", "max"))
    )
    out["R$/L"] = np.where(out["Volume"].abs() > 1e-12, out["Valor"] / out["Volume"], np.nan)
    out["Preço Unitário"] = np.where(out["Consumo"].abs() > 1e-12, out["Valor"] / out["Consumo"], np.nan)
    out["Participação %"] = np.where(
        abs(out["Valor"].sum()) > 1e-12, out["Valor"] / out["Valor"].sum() * 100, 0
    )
    return out.sort_values("Valor", ascending=False)


def delta_by_line(
    fact: pd.DataFrame,
    current_month: str,
    base_month: str,
    group: str | None = None,
    line: str | None = None,
) -> pd.DataFrame:
    cur = fact[fact["Mês"] == current_month].copy()
    base = fact[fact["Mês"] == base_month].copy()
    if group and group != "Todos":
        cur = cur[cur["Grupo"] == group]
        base = base[base["Grupo"] == group]
    if line and line != "Todos":
        cur = cur[cur["Linha"] == line]
        base = base[base["Linha"] == line]

    keys = ["Grupo", "Linha"]
    cur = cur.groupby(keys, as_index=False).agg(
        Consumo_Atual=("Consumo","sum"), Valor_Atual=("Valor","sum")
    )
    base = base.groupby(keys, as_index=False).agg(
        Consumo_Base=("Consumo","sum"), Valor_Base=("Valor","sum")
    )
    out = cur.merge(base, on=keys, how="outer").fillna(0)
    out["Delta R$"] = out["Valor_Atual"] - out["Valor_Base"]
    return out.sort_values("Delta R$", ascending=False)


def months_label_map(months):
    return {m: _display_month(m) for m in months}
