from __future__ import annotations

import math
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from dashboard_core import (
    aggregate,
    annual_average_kpis,
    comparison_snapshot,
    filter_fact,
    group_snapshot,
    line_snapshot,
    load_data,
    monthly_kpis,
    months_label_map,
    snapshot_by,
    transaction_detail,
)


st.set_page_config(
    page_title="Fechamento de Custos - Italac",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
<style>
.block-container { padding-top: 1.0rem; max-width: 1600px; }
div[data-testid="stMetric"] {
    border: 1px solid rgba(128,128,128,.18);
    padding: 10px 13px;
    border-radius: 10px;
    min-height: 112px;
}
div[data-testid="stMetricLabel"] { font-size: 0.92rem !important; }
div[data-testid="stMetricValue"] {
    font-size: 1.72rem !important;
    line-height: 1.12 !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
    word-break: normal !important;
}
div[data-testid="stMetricDelta"] { font-size: 0.86rem !important; }
div[data-testid="stDataFrame"] [role="columnheader"],
div[data-testid="stDataFrame"] [role="gridcell"] {
    justify-content: center !important;
    text-align: center !important;
}
</style>
""",
    unsafe_allow_html=True,
)

st.title("Dashboard de Fechamento de Custos - Italac")
st.caption(
    "Análise de custos por filial, grupo, evento, subevento, produto, consumo, "
    "preço e custo por litro de leite processado."
)


def fmt_number(value, decimals=1):
    if value is None or not math.isfinite(float(value)):
        return "—"
    return (
        f"{float(value):,.{decimals}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def fmt_rs(value):
    if value is None or not math.isfinite(float(value)):
        return "—"
    return f"R$ {fmt_number(value, 2)}"


def fmt_rsl(value):
    if value is None or not math.isfinite(float(value)):
        return "—"
    return f"R$ {fmt_number(value, 4)}"


def fmt_pct(value):
    if value is None or not math.isfinite(float(value)):
        return "—"
    return f"{fmt_number(value, 1)}%"


def fmt_milhoes_litros(value):
    if value is None or not math.isfinite(float(value)):
        return "—"
    return f"{fmt_number(value / 1_000_000, 2)} M L"


def center_table(df, formats=None):
    styler = df.style.set_properties(
        **{"text-align": "center", "vertical-align": "middle"}
    ).set_table_styles(
        [
            {
                "selector": "th, td",
                "props": [
                    ("text-align", "center"),
                    ("vertical-align", "middle"),
                ],
            }
        ]
    )
    if formats:
        styler = styler.format(formats, na_rep="—")
    return styler


@st.cache_data(show_spinner=False)
def cached_load_data(uploaded_bytes: bytes):
    return load_data(uploaded_bytes)


with st.sidebar:
    st.header("Entrada de dados")
    uploaded = st.file_uploader(
        "Arquivo de fechamento",
        type=["xlsx", "xlsm"],
        help=(
            "A aba de custos deve conter DT Digitacao, DT Emissao, Filial, "
            "Descr Grupo, Descr Evento, Descr Sub Ev, Produto, Quantidade e "
            "Vlr.Total. DT Emissao será usada quando DT Digitacao estiver vazia."
        ),
    )
    if not uploaded:
        st.info("Carregue o arquivo de fechamento atualizado.")
        st.stop()

try:
    data = cached_load_data(uploaded.getvalue())
except Exception as exc:
    st.error(f"Erro na leitura do arquivo: {exc}")
    st.stop()

fact = data["fact"]
months = data["months"]
month_labels = months_label_map(months)

if not months:
    st.warning(
        "Não foram encontrados pagamentos com DT Digitacao ou DT Emissao válida."
    )
    st.stop()

with st.sidebar:
    st.success(f"Aba de custos: {data['fact_sheet']}")
    if data["volume_sheet"]:
        st.success(f"Aba de produção: {data['volume_sheet']}")
    else:
        st.warning("Aba 'Consumo' não encontrada; R$/L ficará sem base de volume.")

    quality = data["quality"]
    st.caption(
        f"Registros válidos: {quality['valid_rows']} de {quality['source_rows']}"
    )
    if quality.get("fallback_date_rows", 0):
        st.info(
            f"{quality['fallback_date_rows']} registro(s) sem DT Digitacao "
            "utilizaram DT Emissao como data de referência."
        )
    if quality["invalid_date_rows"]:
        st.warning(
            f"{quality['invalid_date_rows']} registro(s) ignorado(s) por não "
            "possuírem DT Digitacao nem DT Emissao válida."
        )

    selected_month = st.selectbox(
        "Mês de análise",
        months,
        index=len(months) - 1,
        format_func=lambda value: month_labels[value],
    )

    annual_average_key = "__MEDIA_ANUAL__"
    comparison_options = [
        month for month in months if month != selected_month
    ] + [annual_average_key]
    comparison_period = st.selectbox(
        "Comparar com",
        comparison_options,
        format_func=lambda value: (
            f"Média anual ({selected_month[:4]})"
            if value == annual_average_key
            else month_labels.get(value, value)
        ),
    )

    st.divider()

    selected_filial = st.selectbox("Filial", ["Todas"] + data["filiais"])
    filial_fact = filter_fact(fact, filial=selected_filial)

    group_options = sorted(filial_fact["Grupo"].dropna().unique())
    selected_group = st.selectbox("Grupo de custo", ["Todos"] + group_options)
    group_fact = filter_fact(filial_fact, group=selected_group)

    line_options = sorted(group_fact["Linha"].dropna().unique())
    selected_line = st.selectbox(
        "Linha de custo",
        ["Todos"] + line_options,
        help="A linha de custo corresponde ao campo Descr Evento.",
    )
    line_fact = filter_fact(group_fact, line=selected_line)

    subevent_options = sorted(line_fact["Subevento"].dropna().unique())
    selected_subevent = st.selectbox(
        "Subevento", ["Todos"] + subevent_options
    )
    subevent_fact = filter_fact(line_fact, subevent=selected_subevent)

    product_options = sorted(subevent_fact["Produto"].dropna().unique())
    selected_product = st.selectbox(
        "Produto", ["Todos"] + product_options
    )

metric_labels = {
    "Real Bruto (R$)": "Valor",
    "Unidade de consumo": "Consumo",
    "Preço unitário (R$/unid.)": "Preço Unitário",
    "Reais por litro (R$/L)": "R$/L",
}
metric = st.sidebar.selectbox("Métrica dos gráficos", list(metric_labels))
metric_key = metric_labels[metric]

metric_titles = {
    "Valor": "Custo total",
    "Consumo": "Consumo",
    "Preço Unitário": "Preço unitário",
    "R$/L": "Custo por litro",
}
metric_axis_titles = {
    "Valor": "Custo total (R$)",
    "Consumo": "Consumo",
    "Preço Unitário": "Preço unitário (R$/unid.)",
    "R$/L": "Custo por litro (R$/L)",
}
metric_kpi_keys = {
    "Valor": "cost",
    "Consumo": "consumption",
    "Preço Unitário": "unit_price",
    "R$/L": "per_liter",
}


def fmt_metric_value(value):
    if metric_key == "Consumo":
        return fmt_number(value, 1)
    if metric_key == "R$/L":
        return fmt_rsl(value)
    return fmt_rs(value)


filter_kwargs = {
    "filial": selected_filial,
    "group": selected_group,
    "line": selected_line,
    "subevent": selected_subevent,
    "product": selected_product,
}
scoped_fact = filter_fact(fact, **filter_kwargs)

scope_parts = []
if selected_filial != "Todas":
    scope_parts.append(f"Filial {selected_filial}")
if selected_group != "Todos":
    scope_parts.append(selected_group)
if selected_line != "Todos":
    scope_parts.append(selected_line)
if selected_subevent != "Todos":
    scope_parts.append(selected_subevent)
if selected_product != "Todos":
    scope_parts.append(selected_product)
scope = " | ".join(scope_parts) if scope_parts else "Todos os custos"

comparison_is_annual = comparison_period == annual_average_key
comparison_label = (
    f"Média anual ({selected_month[:4]})"
    if comparison_is_annual
    else month_labels[comparison_period]
)


def breakdown_definition():
    if selected_line == "Todos":
        return "Linha", ["Grupo", "Linha"], "Evento"
    if selected_subevent == "Todos":
        return "Subevento", ["Grupo", "Linha", "Subevento"], "Subevento"
    return "Produto", ["Grupo", "Linha", "Subevento", "Produto"], "Produto"


breakdown_column, breakdown_keys, breakdown_label = breakdown_definition()

kpi = monthly_kpis(scoped_fact, selected_month, fact)
production_l = float(kpi["volume"])

st.header(f"Visão Resumida — {scope}")
st.caption(f"Mês analisado: {month_labels[selected_month]}")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Custo total", fmt_rs(kpi["cost"]))
with c2:
    st.metric("Consumo", fmt_number(kpi["consumption"], 1))
with c3:
    st.metric("Produção de leite", fmt_milhoes_litros(production_l))
with c4:
    st.metric("Custo por litro", fmt_rsl(kpi["per_liter"]))

if production_l <= 0:
    st.warning(
        f"Não há volume processado informado para {month_labels[selected_month]}; "
        "o custo por litro ficará indisponível nesse mês."
    )

main_tabs = st.tabs(
    [
        "Evolução da linha",
        "Variação por linha",
        "Comparativo",
        "Pareto",
        "Drill-down de grupos",
        "Detalhamento transacional",
    ]
)


# ==========================================================
# 1. EVOLUÇÃO
# ==========================================================
with main_tabs[0]:
    st.subheader("Evolução mensal")

    agg = aggregate(fact, **filter_kwargs)
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=agg["Mês Label"],
            y=agg[metric_key],
            mode="lines+markers",
            name=metric,
            hovertemplate="%{x}<br>%{y}<extra></extra>",
        )
    )

    annual_kpi = annual_average_kpis(scoped_fact, selected_month, fact)
    annual_metric = {
        "Valor": annual_kpi["cost"],
        "Consumo": annual_kpi["consumption"],
        "Preço Unitário": annual_kpi["unit_price"],
        "R$/L": annual_kpi["per_liter"],
    }[metric_key]
    year_mask = agg["Mês"].astype(str).str.startswith(selected_month[:4])
    annual_x = agg.loc[year_mask, "Mês Label"]
    if not annual_x.empty and math.isfinite(annual_metric):
        fig.add_trace(
            go.Scatter(
                x=annual_x,
                y=[annual_metric] * len(annual_x),
                mode="lines",
                name=f"Média anual ({selected_month[:4]})",
                line={
                    "color": "rgba(99, 110, 250, 0.55)",
                    "width": 2,
                    "dash": "dot",
                    "shape": "spline",
                },
                hovertemplate=(
                    f"Média anual: {fmt_metric_value(annual_metric)}<extra></extra>"
                ),
            )
        )

    fig.update_layout(
        title=f"{metric} — {scope}",
        xaxis_title="Mês",
        yaxis_title=metric_axis_titles[metric_key],
        hovermode="x unified",
        height=430,
        margin={"l": 15, "r": 15, "t": 55, "b": 20},
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Resumo mensal")
    monthly_table = agg[
        ["Mês Label", "Consumo", "Preço Unitário", "Valor", "R$/L", "Volume"]
    ].rename(
        columns={
            "Mês Label": "Mês",
            "Valor": "Custo total (R$)",
            "R$/L": "Custo por litro (R$/L)",
            "Volume": "Produção de leite",
        }
    )
    st.dataframe(
        center_table(
            monthly_table,
            {
                "Consumo": lambda value: fmt_number(value, 1),
                "Preço Unitário": fmt_rs,
                "Custo total (R$)": fmt_rs,
                "Custo por litro (R$/L)": fmt_rsl,
                "Produção de leite": fmt_milhoes_litros,
            },
        ),
        use_container_width=True,
        hide_index=True,
    )


# ==========================================================
# 2. VARIAÇÃO POR LINHA / CAUSAS DO DESVIO
# ==========================================================
with main_tabs[1]:
    st.subheader("Variação por linha")
    st.caption(
        f"O desvio é aberto por {breakdown_label.lower()} conforme o nível atual "
        "dos filtros. Assim é possível avançar de evento até produto."
    )

    comparison = comparison_snapshot(
        fact,
        selected_month,
        comparison_period,
        breakdown_keys,
        annual=comparison_is_annual,
        **filter_kwargs,
    )
    delta_metric_column = f"Delta {metric_key}"
    comparison = comparison.sort_values(delta_metric_column, ascending=False)
    top = pd.concat([comparison.head(10), comparison.tail(10)]).drop_duplicates(
        breakdown_keys
    )

    if top.empty:
        st.info("Não há dados para o filtro selecionado.")
    else:
        top = top.copy()
        top["Rótulo"] = top[breakdown_column].astype(str)
        if breakdown_column == "Linha" and selected_group == "Todos":
            top["Rótulo"] = top["Linha"] + " · " + top["Grupo"]

        ordered = top.sort_values(delta_metric_column)
        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=ordered[delta_metric_column],
                y=ordered["Rótulo"],
                orientation="h",
                name=f"Variação de {metric_titles[metric_key].lower()}",
                text=ordered[delta_metric_column].map(fmt_metric_value),
                textposition="outside",
                cliponaxis=False,
            )
        )
        fig.update_layout(
            title=(
                f"Variação por {breakdown_label.lower()} — "
                f"{month_labels[selected_month]} × {comparison_label}"
            ),
            xaxis_title=f"Variação — {metric_axis_titles[metric_key]}",
            yaxis_title=breakdown_label,
            height=max(480, 32 * len(ordered) + 130),
            margin={"l": 15, "r": 90, "t": 55, "b": 20},
        )
        st.plotly_chart(fig, use_container_width=True)

        comparison["Δ consumo (%)"] = (
            comparison["Consumo Atual"]
            / comparison["Consumo Base"].replace(0, pd.NA)
            - 1
        ) * 100
        comparison["Δ preço (%)"] = (
            comparison["Preço Unitário Atual"]
            / comparison["Preço Unitário Base"].replace(0, pd.NA)
            - 1
        ) * 100

        comparison_columns = breakdown_keys + [
            "Consumo Atual",
            "Consumo Base",
            "Preço Unitário Atual",
            "Preço Unitário Base",
            "Δ consumo (%)",
            "Δ preço (%)",
            "Valor Atual",
            "Valor Base",
            "Delta Valor",
            "Delta Valor %",
        ]
        comparison_table = comparison[comparison_columns].rename(
            columns={
                "Linha": "Evento",
                "Preço Unitário Atual": "Preço atual",
                "Preço Unitário Base": "Preço base",
                "Valor Atual": "Custo atual (R$)",
                "Valor Base": "Custo base (R$)",
                "Delta Valor": "Variação de custo (R$)",
                "Delta Valor %": "Variação de custo (%)",
            }
        )
        st.dataframe(
            center_table(
                comparison_table,
                {
                    "Consumo Atual": lambda value: fmt_number(value, 1),
                    "Consumo Base": lambda value: fmt_number(value, 1),
                    "Preço atual": fmt_rs,
                    "Preço base": fmt_rs,
                    "Δ consumo (%)": fmt_pct,
                    "Δ preço (%)": fmt_pct,
                    "Custo atual (R$)": fmt_rs,
                    "Custo base (R$)": fmt_rs,
                    "Variação de custo (R$)": fmt_rs,
                    "Variação de custo (%)": fmt_pct,
                },
            ),
            use_container_width=True,
            hide_index=True,
        )


# ==========================================================
# 3. COMPARATIVO
# ==========================================================
with main_tabs[2]:
    st.subheader("Comparativo de períodos")

    left = (
        annual_average_kpis(scoped_fact, selected_month, fact)
        if comparison_is_annual
        else monthly_kpis(scoped_fact, comparison_period, fact)
    )
    right = kpi
    compare_display = pd.DataFrame(
        {
            "Indicador": [
                "Custo total",
                "Consumo",
                "Produção de leite",
                "Preço unitário",
                "Custo por litro",
            ],
            comparison_label: [
                fmt_rs(left["cost"]),
                fmt_number(left["consumption"], 1),
                fmt_milhoes_litros(left["volume"]),
                fmt_rs(left["unit_price"]),
                fmt_rsl(left["per_liter"]),
            ],
            month_labels[selected_month]: [
                fmt_rs(right["cost"]),
                fmt_number(right["consumption"], 1),
                fmt_milhoes_litros(right["volume"]),
                fmt_rs(right["unit_price"]),
                fmt_rsl(right["per_liter"]),
            ],
        }
    )
    st.dataframe(
        center_table(compare_display),
        use_container_width=True,
        hide_index=True,
    )

    comparison_metric_key = metric_kpi_keys[metric_key]
    values = pd.DataFrame(
        {
            "Período": [comparison_label, month_labels[selected_month]],
            "Métrica": [left[comparison_metric_key], right[comparison_metric_key]],
        }
    )
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=values["Período"],
            y=values["Métrica"],
            text=values["Métrica"].map(fmt_metric_value),
            textposition="outside",
            cliponaxis=False,
            name=metric_titles[metric_key],
        )
    )
    fig.update_layout(
        title=f"{metric_titles[metric_key]}: referência × mês atual",
        yaxis_title=metric_axis_titles[metric_key],
        height=390,
        margin={"l": 15, "r": 30, "t": 55, "b": 20},
    )
    st.plotly_chart(fig, use_container_width=True)


# ==========================================================
# 4. PARETO
# ==========================================================
with main_tabs[3]:
    st.subheader(f"Pareto — {metric_titles[metric_key]}")
    st.caption(f"Nível de detalhamento atual: {breakdown_label}.")

    pareto = snapshot_by(
        fact,
        selected_month,
        breakdown_keys,
        **filter_kwargs,
    )
    pareto[metric_key] = pd.to_numeric(pareto[metric_key], errors="coerce")
    pareto = (
        pareto[pareto[metric_key].notna() & pareto[metric_key].ne(0)]
        .sort_values(metric_key, ascending=False)
        .head(20)
    )

    if pareto.empty:
        st.info("Não há dados para o filtro selecionado.")
    else:
        total = float(pareto[metric_key].sum())
        pareto["Participação da métrica (%)"] = (
            pareto[metric_key] / total * 100 if abs(total) > 1e-12 else 0
        )
        pareto["Acumulado %"] = (
            pareto[metric_key].cumsum() / total * 100
            if abs(total) > 1e-12
            else 0
        )
        pareto["Rótulo"] = pareto[breakdown_column].astype(str)
        if breakdown_column == "Linha" and selected_group == "Todos":
            pareto["Rótulo"] = pareto["Linha"] + " · " + pareto["Grupo"]

        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=pareto["Rótulo"],
                y=pareto[metric_key],
                name=metric_titles[metric_key],
                text=pareto[metric_key].map(fmt_metric_value),
                textposition="outside",
                cliponaxis=False,
            )
        )
        fig.add_trace(
            go.Scatter(
                x=pareto["Rótulo"],
                y=pareto["Acumulado %"],
                name="Acumulado (%)",
                yaxis="y2",
                mode="lines+markers+text",
                text=pareto["Acumulado %"].map(fmt_pct),
                textposition="top center",
            )
        )
        fig.update_layout(
            title=(
                f"Pareto de {metric_titles[metric_key].lower()} por "
                f"{breakdown_label.lower()} — {month_labels[selected_month]}"
            ),
            xaxis_title=breakdown_label,
            yaxis={"title": metric_axis_titles[metric_key]},
            yaxis2={
                "title": "Acumulado (%)",
                "overlaying": "y",
                "side": "right",
                "range": [0, 105],
            },
            height=590,
            margin={"l": 15, "r": 70, "t": 55, "b": 140},
        )
        st.plotly_chart(fig, use_container_width=True)

        metric_table_label = metric_axis_titles[metric_key]
        pareto_table = pareto[
            breakdown_keys
            + [metric_key, "Participação da métrica (%)", "Acumulado %"]
        ].rename(columns={"Linha": "Evento", metric_key: metric_table_label})
        st.dataframe(
            center_table(
                pareto_table,
                {
                    metric_table_label: fmt_metric_value,
                    "Participação da métrica (%)": fmt_pct,
                    "Acumulado %": fmt_pct,
                },
            ),
            use_container_width=True,
            hide_index=True,
        )


def formatted_snapshot_table(snapshot: pd.DataFrame, dimensions: list[str]):
    columns = dimensions + [
        "Pagamentos",
        "Consumo",
        "Valor",
        "Preço Unitário",
        "R$/L",
        "Participação %",
    ]
    table = snapshot[columns].rename(
        columns={
            "Linha": "Evento",
            "Valor": "Custo total (R$)",
            "Preço Unitário": "Preço unitário (R$/unid.)",
            "R$/L": "Custo por litro (R$/L)",
            "Participação %": "Participação (%)",
        }
    )
    return center_table(
        table,
        {
            "Pagamentos": lambda value: fmt_number(value, 0),
            "Consumo": lambda value: fmt_number(value, 1),
            "Custo total (R$)": fmt_rs,
            "Preço unitário (R$/unid.)": fmt_rs,
            "Custo por litro (R$/L)": fmt_rsl,
            "Participação (%)": fmt_pct,
        },
    )


# ==========================================================
# 5. DRILL-DOWN DE GRUPOS
# ==========================================================
with main_tabs[4]:
    st.subheader("Drill-down de grupos de custo")
    st.caption(
        "A visão de grupos respeita a filial selecionada. As tabelas seguintes "
        "aprofundam o grupo principal até evento, subevento e produto."
    )

    groups = group_snapshot(fact, selected_month, filial=selected_filial)
    if groups.empty:
        st.info("Não há dados para a filial e o mês selecionados.")
    else:
        group_plot = groups.sort_values(metric_key)
        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=group_plot[metric_key],
                y=group_plot["Grupo"],
                orientation="h",
                name=f"{metric_titles[metric_key]} por grupo",
                text=group_plot[metric_key].map(fmt_metric_value),
                textposition="outside",
                cliponaxis=False,
            )
        )
        fig.update_layout(
            title=(
                f"{metric_titles[metric_key]} por grupo — "
                f"{month_labels[selected_month]}"
            ),
            xaxis_title=metric_axis_titles[metric_key],
            yaxis_title="Grupo",
            height=520,
            margin={"l": 15, "r": 100, "t": 55, "b": 20},
        )
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(
            formatted_snapshot_table(groups, ["Grupo"]),
            use_container_width=True,
            hide_index=True,
        )

    drill_scope = (
        selected_group if selected_group != "Todos" else "todos os grupos"
    )
    st.subheader(f"Eventos — {drill_scope}")
    events = line_snapshot(
        fact,
        selected_month,
        group=selected_group,
        filial=selected_filial,
    ).head(50)
    if events.empty:
        st.info("Não há eventos para o filtro selecionado.")
    else:
        st.dataframe(
            formatted_snapshot_table(events, ["Grupo", "Linha"]),
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("Subeventos")
    subevents = snapshot_by(
        fact,
        selected_month,
        ["Grupo", "Linha", "Subevento"],
        filial=selected_filial,
        group=selected_group,
        line=selected_line,
    ).head(50)
    if subevents.empty:
        st.info("Não há subeventos para o filtro selecionado.")
    else:
        st.dataframe(
            formatted_snapshot_table(
                subevents, ["Grupo", "Linha", "Subevento"]
            ),
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("Produtos")
    products = snapshot_by(
        fact,
        selected_month,
        ["Grupo", "Linha", "Subevento", "Produto"],
        filial=selected_filial,
        group=selected_group,
        line=selected_line,
        subevent=selected_subevent,
    ).head(100)
    if products.empty:
        st.info("Não há produtos para o filtro selecionado.")
    else:
        st.dataframe(
            formatted_snapshot_table(
                products, ["Grupo", "Linha", "Subevento", "Produto"]
            ),
            use_container_width=True,
            hide_index=True,
        )


# ==========================================================
# 6. DETALHAMENTO TRANSACIONAL
# ==========================================================
with main_tabs[5]:
    st.subheader("Detalhamento transacional")
    st.caption(
        "Pagamentos individuais apurados pela DT Digitacao; quando ela estiver "
        "vazia, será utilizada a DT Emissao. Todos os filtros da barra lateral "
        "são respeitados."
    )

    supplier_scope = filter_fact(fact, **filter_kwargs)
    supplier_scope = supplier_scope[supplier_scope["Mês"].eq(selected_month)]
    supplier_options = sorted(supplier_scope["Fornecedor"].dropna().unique())
    selected_supplier = st.selectbox(
        "Fornecedor",
        ["Todos"] + supplier_options,
        key="transaction_supplier",
    )

    payments = transaction_detail(
        fact,
        selected_month,
        supplier=selected_supplier,
        **filter_kwargs,
    )

    payment_total = float(payments["Valor"].sum()) if not payments.empty else 0.0
    payment_average = payment_total / len(payments) if len(payments) else math.nan
    supplier_count = payments["Fornecedor"].nunique() if not payments.empty else 0
    product_count = (
        payments["Descrição Produto"].nunique() if not payments.empty else 0
    )

    t1, t2, t3, t4, t5 = st.columns(5)
    with t1:
        st.metric("Pagamentos", fmt_number(len(payments), 0))
    with t2:
        st.metric("Valor filtrado", fmt_rs(payment_total))
    with t3:
        st.metric("Ticket médio", fmt_rs(payment_average))
    with t4:
        st.metric("Fornecedores", fmt_number(supplier_count, 0))
    with t5:
        st.metric("Produtos", fmt_number(product_count, 0))

    if payments.empty:
        st.info("Não há pagamentos individuais para os filtros selecionados.")
    else:
        product_analysis = snapshot_by(
            fact,
            selected_month,
            "Produto",
            supplier=selected_supplier,
            **filter_kwargs,
        ).head(15)
        supplier_analysis = snapshot_by(
            fact,
            selected_month,
            "Fornecedor",
            supplier=selected_supplier,
            **filter_kwargs,
        ).head(15)

        product_column, supplier_column = st.columns(2)
        with product_column:
            product_plot = product_analysis.sort_values("Valor")
            fig = go.Figure(
                go.Bar(
                    x=product_plot["Valor"],
                    y=product_plot["Produto"],
                    orientation="h",
                    text=product_plot["Valor"].map(fmt_rs),
                    textposition="outside",
                    cliponaxis=False,
                    name="Produtos",
                )
            )
            fig.update_layout(
                title="Produtos com maior custo",
                xaxis_title="Custo total (R$)",
                yaxis_title="Produto",
                height=500,
                margin={"l": 15, "r": 80, "t": 55, "b": 20},
            )
            st.plotly_chart(fig, use_container_width=True)

        with supplier_column:
            supplier_plot = supplier_analysis.sort_values("Valor")
            fig = go.Figure(
                go.Bar(
                    x=supplier_plot["Valor"],
                    y=supplier_plot["Fornecedor"],
                    orientation="h",
                    text=supplier_plot["Valor"].map(fmt_rs),
                    textposition="outside",
                    cliponaxis=False,
                    name="Fornecedores",
                )
            )
            fig.update_layout(
                title="Fornecedores com maior custo",
                xaxis_title="Custo total (R$)",
                yaxis_title="Fornecedor",
                height=500,
                margin={"l": 15, "r": 80, "t": 55, "b": 20},
            )
            st.plotly_chart(fig, use_container_width=True)

        payment_table = payments.rename(
            columns={
                "Data Digitação": "Data de referência",
                "Linha": "Evento",
                "Consumo": "Quantidade",
                "Valor": "Valor total (R$)",
                "Preço Unitário": "Preço unitário (R$/unid.)",
            }
        )
        st.subheader("Pagamentos individuais")
        st.dataframe(
            center_table(
                payment_table,
                {
                    "Data de referência": lambda value: value.strftime("%d/%m/%Y"),
                    "Quantidade": lambda value: fmt_number(value, 2),
                    "Valor total (R$)": fmt_rs,
                    "Preço unitário (R$/unid.)": fmt_rs,
                },
            ),
            use_container_width=True,
            hide_index=True,
        )

        csv_bytes = payment_table.to_csv(
            index=False,
            sep=";",
            decimal=",",
            date_format="%d/%m/%Y",
        ).encode("utf-8-sig")
        st.download_button(
            "Baixar pagamentos filtrados (CSV)",
            data=csv_bytes,
            file_name=f"pagamentos_filtrados_{selected_month}.csv",
            mime="text/csv",
        )

st.divider()
st.caption(
    f"Base de cálculo: a aba '{data['fact_sheet']}' fornece os pagamentos "
    "individuais e o mês é definido pela DT Digitacao; quando ela estiver "
    "vazia ou inválida, será utilizada a DT Emissao. A aba 'Consumo' fornece "
    "o volume de leite. Preço unitário = Vlr.Total / Quantidade. "
    "Custo por litro = Vlr.Total / Volume processado."
)
