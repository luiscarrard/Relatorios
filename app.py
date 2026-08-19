
from __future__ import annotations

import math
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from dashboard_core import (
    load_data,
    aggregate,
    monthly_kpis,
    annual_average_kpis,
    annual_average_by_line,
    group_snapshot,
    line_snapshot,
    months_label_map,
)

st.set_page_config(
    page_title="Fechamento de Custos - Unidade de Passo Fundo",
    page_icon="📊",
    layout="wide",
)

# ---------- Visual ----------
st.markdown("""
<style>
.block-container { padding-top: 1.0rem; max-width: 1600px; }

div[data-testid="stMetric"] {
    border: 1px solid rgba(128,128,128,.18);
    padding: 10px 13px;
    border-radius: 10px;
    min-height: 112px;
}
div[data-testid="stMetricLabel"] {
    font-size: 0.92rem !important;
}
div[data-testid="stMetricValue"] {
    font-size: 1.72rem !important;
    line-height: 1.12 !important;
    white-space: normal !important;
    overflow: visible !important;
    text-overflow: clip !important;
    word-break: normal !important;
}
div[data-testid="stMetricDelta"] {
    font-size: 0.86rem !important;
}
div[data-testid="stDataFrame"] [role="columnheader"],
div[data-testid="stDataFrame"] [role="gridcell"] {
    justify-content: center !important;
    text-align: center !important;
}
</style>
""", unsafe_allow_html=True)

st.title("Dashboard de Fechamento de Custos - Unidade de Passo Fundo")
st.caption("Análise de custos por grupo, linha, consumo, preço e custo por litro de leite processado.")

# ---------- Formatting helpers ----------
def fmt_number(v, decimals=1):
    if v is None or not math.isfinite(float(v)):
        return "—"
    return f"{float(v):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_rs(v):
    if v is None or not math.isfinite(float(v)):
        return "—"
    return f"R$ {fmt_number(v, 2)}"

def fmt_rsl(v):
    if v is None or not math.isfinite(float(v)):
        return "—"
    return f"R$ {fmt_number(v, 4)}"

def fmt_pct(v):
    if v is None or not math.isfinite(float(v)):
        return "—"
    return f"{fmt_number(v, 1)}%"

def fmt_milhoes_litros(v):
    if v is None or not math.isfinite(float(v)):
        return "—"
    return f"{fmt_number(v / 1_000_000, 2)} M L"

def center_table(df, formats=None):
    styler = df.style.set_properties(
        **{"text-align": "center", "vertical-align": "middle"}
    )
    styler = styler.set_table_styles([
        {"selector": "th, td", "props": [
            ("text-align", "center"),
            ("vertical-align", "middle"),
        ]}
    ])
    if formats:
        styler = styler.format(formats, na_rep="—")
    return styler

# ---------- Input ----------
with st.sidebar:
    st.header("Entrada de dados")
    uploaded = st.file_uploader(
        "Arquivo de fechamento",
        type=["xlsx", "xlsm"],
        help="Estrutura: Mês | Grupo | Linha | Consumo | Valor."
    )

    if not uploaded:
        st.info("Carregue o arquivo de fechamento atualizado.")
        st.stop()

try:
    data = load_data(uploaded.getvalue())
except Exception as exc:
    st.error(f"Erro na leitura do arquivo: {exc}")
    st.stop()

fact = data["fact"]
months = data["months"]
month_labels = months_label_map(months)

if not months:
    st.warning("Não foram encontrados meses na aba de custos.")
    st.stop()

with st.sidebar:
    st.success(f"Aba de custos: {data['fact_sheet']}")
    if data["volume_sheet"]:
        st.success(f"Aba de produção: {data['volume_sheet']}")
    else:
        st.warning("Aba 'Consumo' não encontrada; R$/L ficará sem base de volume.")

    selected_month = st.selectbox(
        "Mês de análise",
        months,
        index=len(months) - 1,
        format_func=lambda x: month_labels[x],
    )

    annual_average_key = "__MEDIA_ANUAL__"
    comparison_options = [m for m in months if m != selected_month] + [annual_average_key]
    comparison_period = st.selectbox(
        "Comparar com",
        comparison_options,
        format_func=lambda x: (
            f"Média anual ({selected_month[:4]})"
            if x == annual_average_key else month_labels.get(x, x)
        ),
    )

    st.divider()

    selected_group = st.selectbox(
        "Grupo de custo",
        ["Todos"] + data["groups"],
    )

    lines_available = sorted(
        fact.loc[
            fact["Grupo"].eq(selected_group)
            if selected_group != "Todos"
            else fact.index.notna(),
            "Linha",
        ].unique()
    )

    selected_line = st.selectbox(
        "Linha individual",
        ["Todos"] + lines_available,
    )

metric_labels = {
    "Reais total": "Valor",
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

def fmt_metric_value(v):
    if metric_key == "Consumo":
        return fmt_number(v, 1)
    if metric_key == "R$/L":
        return fmt_rsl(v)
    return fmt_rs(v)

# ---------- Scope ----------
scoped_fact = fact.copy()
if selected_group != "Todos":
    scoped_fact = scoped_fact[scoped_fact["Grupo"].eq(selected_group)]
if selected_line != "Todos":
    scoped_fact = scoped_fact[scoped_fact["Linha"].eq(selected_line)]

scope = (
    "Todos os custos"
    if selected_line == "Todos" and selected_group == "Todos"
    else selected_group
    if selected_line == "Todos"
    else selected_line
)

kpi = monthly_kpis(scoped_fact, selected_month)
production_l = float(kpi["volume"])

st.header(f"Visão Resumida — {scope}")
st.caption(f"Mês analisado: {month_labels[selected_month]}")

# ---------- KPI cards ----------
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric("Custo total", fmt_rs(kpi["cost"]))
with c2:
    st.metric("Consumo", fmt_number(kpi["consumption"], 1))
with c3:
    st.metric("Produção de leite", fmt_milhoes_litros(production_l))
with c4:
    st.metric("Custo por litro", fmt_rsl(kpi["per_liter"]))

comparison_is_annual = comparison_period == annual_average_key
comparison_label = (
    f"Média anual ({selected_month[:4]})"
    if comparison_is_annual else month_labels[comparison_period]
)

# ---------- Navigation ----------
main_tabs = st.tabs([
    "Evolução da linha",
    "Preço × Consumo",
    "Comparativo",
    "Pareto",
    "Drill-down de grupos",
])

# ==========================================================
# 1. EVOLUTION
# ==========================================================
with main_tabs[0]:
    st.subheader("Evolução mensal")

    agg = aggregate(
        fact,
        group=selected_group,
        line=selected_line,
    )

    y_titles = {
        "Valor": "Custo total (R$)",
        "Consumo": "Consumo",
        "Preço Unitário": "Preço unitário (R$/unid.)",
        "R$/L": "Custo por litro (R$/L)",
    }

    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=agg["Mês Label"],
        y=agg[metric_key],
        mode="lines+markers",
        name=metric,
        hovertemplate="%{x}<br>%{y}<extra></extra>",
    ))

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
        fig.add_trace(go.Scatter(
            x=annual_x,
            y=[annual_metric] * len(annual_x),
            mode="lines",
            name=f"Média anual ({selected_month[:4]})",
            line=dict(
                color="rgba(99, 110, 250, 0.55)",
                width=2,
                dash="dot",
                shape="spline",
            ),
            hovertemplate=f"Média anual: {fmt_number(annual_metric, 4)}<extra></extra>",
        ))
    fig.update_layout(
        title=f"{metric} — {scope}",
        xaxis_title="Mês",
        yaxis_title=y_titles[metric_key],
        hovermode="x unified",
        height=430,
        margin=dict(l=15, r=15, t=55, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Resumo mensal")

    tbl = agg[
        ["Mês Label", "Consumo", "Preço Unitário", "Valor", "R$/L", "Volume"]
    ].rename(columns={
        "Mês Label": "Mês",
        "Valor": "Custo total (R$)",
        "R$/L": "Custo por litro (R$/L)",
        "Volume": "Produção de leite",
    })

    st.dataframe(
        center_table(
            tbl,
            {
                "Consumo": lambda v: fmt_number(v, 1),
                "Preço Unitário": lambda v: fmt_rs(v),
                "Custo total (R$)": lambda v: fmt_rs(v),
                "Custo por litro (R$/L)": lambda v: fmt_rsl(v),
                "Produção de leite": lambda v: fmt_milhoes_litros(v),
            },
        ),
        use_container_width=True,
        hide_index=True,
    )

# ==========================================================
# 2. PRICE x CONSUMPTION
# ==========================================================
with main_tabs[1]:
    st.subheader("Preço × Consumo")

    current_lines = line_snapshot(fact, selected_month, selected_group)
    if selected_line != "Todos":
        current_lines = current_lines[current_lines["Linha"] == selected_line]

    if comparison_is_annual:
        base_lines = annual_average_by_line(
            fact,
            selected_month,
            group=selected_group,
            line=selected_line,
        )
        annual_volume = annual_average_kpis(fact, selected_month, fact)["volume"]
        base_lines["Preço Unitário"] = (
            base_lines["Valor"] / base_lines["Consumo"].replace(0, pd.NA)
        )
        base_lines["R$/L"] = (
            base_lines["Valor"] / annual_volume
            if abs(annual_volume) > 1e-12 else math.nan
        )
    else:
        base_lines = line_snapshot(fact, comparison_period, selected_group)
        if selected_line != "Todos":
            base_lines = base_lines[base_lines["Linha"] == selected_line]

    current_metric = current_lines[
        ["Grupo", "Linha", metric_key]
    ].rename(columns={metric_key: "Métrica atual"})
    base_metric = base_lines[
        ["Grupo", "Linha", metric_key]
    ].rename(columns={metric_key: "Métrica base"})
    cmp_df = current_metric.merge(
        base_metric,
        on=["Grupo", "Linha"],
        how="outer",
    )
    cmp_df[["Métrica atual", "Métrica base"]] = cmp_df[
        ["Métrica atual", "Métrica base"]
    ].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    cmp_df["Delta da métrica"] = (
        cmp_df["Métrica atual"] - cmp_df["Métrica base"]
    )
    cmp_df = cmp_df.sort_values("Delta da métrica", ascending=False)

    top = pd.concat([
        cmp_df.head(10),
        cmp_df.tail(10),
    ]).drop_duplicates(["Grupo", "Linha"])

    if top.empty:
        st.info("Não há dados para o filtro selecionado.")
    else:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=top.sort_values("Delta da métrica")["Delta da métrica"],
            y=top.sort_values("Delta da métrica")["Linha"],
            orientation="h",
            name=f"Variação de {metric_titles[metric_key].lower()}",
            text=top.sort_values("Delta da métrica")["Delta da métrica"].map(fmt_metric_value),
            textposition="outside",
            cliponaxis=False,
        ))
        fig.update_layout(
            title=(
                f"Variação de {metric_titles[metric_key].lower()} por linha — "
                f"{month_labels[selected_month]} × {comparison_label}"
            ),
            xaxis_title=f"Variação — {metric_axis_titles[metric_key]}",
            yaxis_title="Linha",
            height=570,
            margin=dict(l=15, r=80, t=55, b=20),
        )
        st.plotly_chart(fig, use_container_width=True)

        current = fact[fact["Mês"] == selected_month].copy()

        if selected_group != "Todos":
            current = current[current["Grupo"] == selected_group]
        if selected_line != "Todos":
            current = current[current["Linha"] == selected_line]

        cur = current.groupby("Linha", as_index=False).agg(
            Consumo=("Consumo", "sum"),
            Valor=("Valor", "sum"),
        )
        if comparison_is_annual:
            bas = annual_average_by_line(
                fact,
                selected_month,
                group=selected_group,
                line=selected_line,
            ).groupby("Linha", as_index=False).agg(
                Consumo_Base=("Consumo", "sum"),
                Valor_Base=("Valor", "sum"),
            )
        else:
            base = fact[fact["Mês"] == comparison_period].copy()
            if selected_group != "Todos":
                base = base[base["Grupo"] == selected_group]
            if selected_line != "Todos":
                base = base[base["Linha"] == selected_line]
            bas = base.groupby("Linha", as_index=False).agg(
                Consumo_Base=("Consumo", "sum"),
                Valor_Base=("Valor", "sum"),
            )

        pc = cur.merge(bas, on="Linha", how="outer").fillna(0)
        pc["Preço Atual"] = pc["Valor"] / pc["Consumo"].replace(0, pd.NA)
        pc["Preço Base"] = pc["Valor_Base"] / pc["Consumo_Base"].replace(0, pd.NA)
        pc["Δ Consumo %"] = (
            pc["Consumo"] / pc["Consumo_Base"].replace(0, pd.NA) - 1
        ) * 100
        pc["Δ Preço %"] = (
            pc["Preço Atual"] / pc["Preço Base"].replace(0, pd.NA) - 1
        ) * 100

        table = pc.sort_values("Valor", ascending=False).rename(columns={
            "Linha": "Linha",
            "Consumo": "Consumo atual",
            "Consumo_Base": "Consumo base",
            "Preço Atual": "Preço atual",
            "Preço Base": "Preço base",
            "Δ Consumo %": "Δ consumo (%)",
            "Δ Preço %": "Δ preço (%)",
            "Valor": "Custo atual (R$)",
            "Valor_Base": "Custo base (R$)",
        })

        st.dataframe(
            center_table(
                table,
                {
                    "Consumo atual": lambda v: fmt_number(v, 1),
                    "Consumo base": lambda v: fmt_number(v, 1),
                    "Preço atual": lambda v: fmt_rs(v),
                    "Preço base": lambda v: fmt_rs(v),
                    "Δ consumo (%)": lambda v: fmt_pct(v),
                    "Δ preço (%)": lambda v: fmt_pct(v),
                    "Custo atual (R$)": lambda v: fmt_rs(v),
                    "Custo base (R$)": lambda v: fmt_rs(v),
                },
            ),
            use_container_width=True,
            hide_index=True,
        )

# ==========================================================
# 3. COMPARATIVE
# ==========================================================
with main_tabs[2]:
    st.subheader("Comparativo de períodos")

    if comparison_is_annual:
        left = annual_average_kpis(scoped_fact, selected_month, fact)
    else:
        left = monthly_kpis(scoped_fact, comparison_period)
    right = kpi

    compare_display = pd.DataFrame({
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
    })

    st.dataframe(
        center_table(compare_display),
        use_container_width=True,
        hide_index=True,
    )

    comparison_metric_key = metric_kpi_keys[metric_key]
    vals = pd.DataFrame({
        "Período": [comparison_label, month_labels[selected_month]],
        "Métrica": [left[comparison_metric_key], right[comparison_metric_key]],
    })

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=vals["Período"],
        y=vals["Métrica"],
        text=vals["Métrica"].map(fmt_metric_value),
        textposition="outside",
        cliponaxis=False,
        name=metric_titles[metric_key],
    ))
    fig.update_layout(
        title=f"{metric_titles[metric_key]}: referência × mês atual",
        yaxis_title=metric_axis_titles[metric_key],
        height=390,
        margin=dict(l=15, r=30, t=55, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)

# ==========================================================
# 4. PARETO
# ==========================================================
with main_tabs[3]:
    st.subheader(f"Pareto — {metric_titles[metric_key]}")

    pareto = line_snapshot(fact, selected_month, selected_group).copy()
    pareto[metric_key] = pd.to_numeric(pareto[metric_key], errors="coerce")
    pareto = pareto[
        pareto[metric_key].notna() & pareto[metric_key].ne(0)
    ].sort_values(metric_key, ascending=False).head(20)

    if pareto.empty:
        st.info("Não há dados para o filtro selecionado.")
    else:
        total = pareto[metric_key].sum()
        pareto["Participação da métrica (%)"] = (
            pareto[metric_key] / total * 100
            if abs(total) > 1e-12 else 0
        )
        pareto["Acumulado %"] = (
            pareto[metric_key].cumsum() / total * 100
            if abs(total) > 1e-12 else 0
        )

        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=pareto["Linha"],
            y=pareto[metric_key],
            name=metric_titles[metric_key],
            text=pareto[metric_key].map(fmt_metric_value),
            textposition="outside",
            cliponaxis=False,
        ))
        fig.add_trace(go.Scatter(
            x=pareto["Linha"],
            y=pareto["Acumulado %"],
            name="Acumulado (%)",
            yaxis="y2",
            mode="lines+markers+text",
            text=pareto["Acumulado %"].map(fmt_pct),
            textposition="top center",
        ))
        fig.update_layout(
            title=f"Pareto de {metric_titles[metric_key].lower()} — {month_labels[selected_month]}",
            xaxis_title="Linha",
            yaxis=dict(title=metric_axis_titles[metric_key]),
            yaxis2=dict(
                title="Acumulado (%)",
                overlaying="y",
                side="right",
                range=[0, 105],
            ),
            height=590,
            margin=dict(l=15, r=70, t=55, b=120),
        )
        st.plotly_chart(fig, use_container_width=True)

        metric_table_label = metric_axis_titles[metric_key]
        table = pareto[
            [
                "Grupo",
                "Linha",
                metric_key,
                "Participação da métrica (%)",
                "Acumulado %",
            ]
        ].rename(columns={metric_key: metric_table_label})

        st.dataframe(
            center_table(
                table,
                {
                    metric_table_label: fmt_metric_value,
                    "Participação da métrica (%)": lambda v: fmt_pct(v),
                    "Acumulado %": lambda v: fmt_pct(v),
                },
            ),
            use_container_width=True,
            hide_index=True,
        )

# ==========================================================
# 5. GROUP DRILL-DOWN
# ==========================================================
with main_tabs[4]:
    st.subheader("Drill-down de grupos de custo")
    st.caption(
        "Esta seção é dedicada aos grupos de custo. A seleção de linha individual da barra lateral "
        "não altera esta análise."
    )

    gs = group_snapshot(fact, selected_month)

    fig = go.Figure()
    gs_plot = gs.sort_values(metric_key)

    fig.add_trace(go.Bar(
        x=gs_plot[metric_key],
        y=gs_plot["Grupo"],
        orientation="h",
        name=f"{metric_titles[metric_key]} por grupo",
        text=gs_plot[metric_key].map(fmt_metric_value),
        textposition="outside",
        cliponaxis=False,
    ))
    fig.update_layout(
        title=f"{metric_titles[metric_key]} por grupo — {month_labels[selected_month]}",
        xaxis_title=metric_axis_titles[metric_key],
        yaxis_title="Grupo",
        height=520,
        margin=dict(l=15, r=100, t=55, b=20),
    )
    st.plotly_chart(fig, use_container_width=True)

    group_table = gs[
        [
            "Grupo",
            "Consumo",
            "Valor",
            "Preço Unitário",
            "R$/L",
            "Participação %",
        ]
    ].rename(columns={
        "Valor": "Custo total (R$)",
        "Preço Unitário": "Preço unitário (R$/unid.)",
        "R$/L": "Custo por litro (R$/L)",
        "Participação %": "Participação (%)",
    })

    st.dataframe(
        center_table(
            group_table,
            {
                "Consumo": lambda v: fmt_number(v, 1),
                "Custo total (R$)": lambda v: fmt_rs(v),
                "Preço unitário (R$/unid.)": lambda v: fmt_rs(v),
                "Custo por litro (R$/L)": lambda v: fmt_rsl(v),
                "Participação (%)": lambda v: fmt_pct(v),
            },
        ),
        use_container_width=True,
        hide_index=True,
    )

    drill_scope = (
        selected_group if selected_group != "Todos" else "todos os grupos"
    )
    st.subheader(f"Linhas — {drill_scope}")
    ls = line_snapshot(fact, selected_month, selected_group).head(30)

    line_table = ls[
        [
            "Grupo",
            "Linha",
            "Consumo",
            "Valor",
            "Preço Unitário",
            "R$/L",
            "Participação %",
        ]
    ].rename(columns={
        "Valor": "Custo total (R$)",
        "Preço Unitário": "Preço unitário (R$/unid.)",
        "R$/L": "Custo por litro (R$/L)",
        "Participação %": "Participação (%)",
    })

    st.dataframe(
        center_table(
            line_table,
            {
                "Consumo": lambda v: fmt_number(v, 1),
                "Custo total (R$)": lambda v: fmt_rs(v),
                "Preço unitário (R$/unid.)": lambda v: fmt_rs(v),
                "Custo por litro (R$/L)": lambda v: fmt_rsl(v),
                "Participação (%)": lambda v: fmt_pct(v),
            },
        ),
        use_container_width=True,
        hide_index=True,
    )

st.divider()
st.caption(
    "Base de cálculo: a aba 'Custos em linha' é a tabela de fatos; "
    "a aba 'Consumo' fornece o volume de produção de leite. "
    "Preço unitário = Valor / Consumo. Custo por litro = Valor / Volume processado."
)
