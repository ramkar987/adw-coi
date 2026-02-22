import io
import math
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from fpdf import FPDF

st.set_page_config(
    page_title="Financial Calculators | COI & Business Impact",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_EXCEL_PATH = "ADW-Demo-COI.xlsx"
SHEET_RETENTION = "COI Retention"
SHEET_NRR = "COI NRR"


# -----------------------------
# Utilities
# -----------------------------
def _safe_float(x, default=0.0) -> float:
    try:
        if x is None:
            return default
        if isinstance(x, (int, float, np.number)):
            return float(x)
        s = str(x).strip().replace(",", "")
        if s == "":
            return default
        return float(s)
    except Exception:
        return default


def fmt_currency(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_number(v: float) -> str:
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def poisson_prob_ge_2(lmbda: float) -> float:
    lmbda = max(0.0, float(lmbda))
    return 1.0 - math.exp(-lmbda) * (1.0 + lmbda)


def scenario_multipliers(scenario: str) -> dict:
    scenario = (scenario or "").lower()
    if scenario.startswith("con"):
        return {"rev_mult": 0.70, "churn_mult": 0.70, "eff_mult": 0.70}
    if scenario.startswith("agr"):
        return {"rev_mult": 1.30, "churn_mult": 1.30, "eff_mult": 1.20}
    return {"rev_mult": 1.00, "churn_mult": 1.00, "eff_mult": 1.00}


# -----------------------------
# Data loading
# -----------------------------
@st.cache_data(show_spinner=False)
def load_params_from_excel_bytes(excel_bytes: bytes, sheet_name: str) -> dict:
    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet_name, engine="openpyxl")
    df = df.rename(columns={c: c.strip() for c in df.columns})
    if "Parameter" not in df.columns or "Input" not in df.columns:
        raise ValueError(f"A aba '{sheet_name}' precisa ter colunas 'Parameter' e 'Input'.")
    out = {}
    for _, row in df.iterrows():
        k = str(row["Parameter"]).strip()
        out[k] = _safe_float(row["Input"])
    return out


@st.cache_data(show_spinner=False)
def load_excel_bytes_from_path(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def get_excel_bytes_from_sidebar() -> bytes | None:
    st.sidebar.markdown("### Dados (Excel)")
    input_mode = st.sidebar.radio(
        "Fonte dos dados",
        ["✏️ Inserir manualmente", "📂 Upload Excel"],
        index=0,
        help="Manual: use os campos interativos com valores padrão. Excel: carrega premissas do arquivo.",
    )

    if input_mode == "📂 Upload Excel":
        uploaded = st.sidebar.file_uploader("Carregar .xlsx", type=["xlsx"])
        if uploaded is not None:
            return uploaded.read()
        try:
            return load_excel_bytes_from_path(DEFAULT_EXCEL_PATH)
        except FileNotFoundError:
            st.sidebar.info("ℹ️ Arquivo local não encontrado. Usando valores padrão.")
            return None

    return None  # Modo manual: sem Excel, usa defaults hardcoded



# -----------------------------
# Calculator 1: COI (Retention / NRR)
# -----------------------------
def calc_coi_retention(inp: dict) -> dict:
    total_agents = _safe_float(inp.get("total_agents"))
    hours_per_agent = _safe_float(inp.get("hours_per_agent"))
    pct_hours_impaired = _safe_float(inp.get("pct_hours_impaired"))
    calls_per_agent = _safe_float(inp.get("calls_per_agent"))
    pct_calls_impaired = _safe_float(inp.get("pct_calls_impaired"))
    customers = _safe_float(inp.get("customers"))
    calls_per_customer = _safe_float(inp.get("calls_per_customer"))
    calls_per_at_risk_customer = _safe_float(inp.get("calls_per_at_risk_customer"))
    switching_rate_2plus = _safe_float(inp.get("switching_rate_2plus"))
    baseline_churn_rate = _safe_float(inp.get("baseline_churn_rate"))
    active_churn_share = _safe_float(inp.get("active_churn_share"))
    cx_driven_share = _safe_float(inp.get("cx_driven_share"))
    arpu = _safe_float(inp.get("arpu"))
    cost_per_hour = _safe_float(inp.get("cost_per_hour"))

    total_agent_hours = total_agents * hours_per_agent
    hours_impaired = total_agent_hours * pct_hours_impaired
    total_calls = total_agents * calls_per_agent
    calls_impaired = total_calls * pct_calls_impaired
    baseline_churned_customers = customers * baseline_churn_rate
    active_churn_customers = baseline_churned_customers * active_churn_share
    cx_driven_churns = active_churn_customers * cx_driven_share

    lmbda = calls_per_at_risk_customer * pct_calls_impaired
    prob_2plus = poisson_prob_ge_2(lmbda)

    customers_at_risk = cx_driven_churns * prob_2plus * switching_rate_2plus
    revenue_at_risk = customers_at_risk * arpu
    labor_ineff_cost = hours_impaired * cost_per_hour
    total_coi = labor_ineff_cost + revenue_at_risk

    return {
        "total_agents": total_agents,
        "hours_per_agent": hours_per_agent,
        "total_agent_hours": total_agent_hours,
        "pct_hours_impaired": pct_hours_impaired,
        "hours_impaired": hours_impaired,
        "calls_per_agent": calls_per_agent,
        "total_calls": total_calls,
        "pct_calls_impaired": pct_calls_impaired,
        "calls_impaired": calls_impaired,
        "customers": customers,
        "calls_per_customer": calls_per_customer,
        "calls_per_at_risk_customer": calls_per_at_risk_customer,
        "probability_2plus_impaired_calls": prob_2plus,
        "switching_rate_2plus": switching_rate_2plus,
        "baseline_churn_rate": baseline_churn_rate,
        "baseline_churned_customers": baseline_churned_customers,
        "active_churn_share": active_churn_share,
        "active_churn_customers": active_churn_customers,
        "cx_driven_share": cx_driven_share,
        "cx_driven_churns": cx_driven_churns,
        "customers_at_risk": customers_at_risk,
        "arpu": arpu,
        "revenue_at_risk": revenue_at_risk,
        "cost_per_hour": cost_per_hour,
        "labor_inefficiency_cost": labor_ineff_cost,
        "upsell_revenue_at_risk_shortfall": 0.0,
        "total_cost_of_inaction": total_coi,
        "mode": "Retention",
    }


def calc_coi_nrr(inp: dict) -> dict:
    total_agents = _safe_float(inp.get("total_agents"))
    hours_per_agent = _safe_float(inp.get("hours_per_agent"))
    pct_hours_impaired = _safe_float(inp.get("pct_hours_impaired"))
    calls_per_agent = _safe_float(inp.get("calls_per_agent"))
    pct_calls_impaired = _safe_float(inp.get("pct_calls_impaired"))
    retry_rate = _safe_float(inp.get("retry_rate"))
    fcr = _safe_float(inp.get("fcr"))
    csat = _safe_float(inp.get("csat"))
    ces = _safe_float(inp.get("ces"))
    ces_divisor = _safe_float(inp.get("ces_divisor"))
    customers = _safe_float(inp.get("customers"))
    baseline_churn_rate = _safe_float(inp.get("baseline_churn_rate"))
    churn_risk_multiplier = _safe_float(inp.get("churn_risk_multiplier"))
    active_churn_share = _safe_float(inp.get("active_churn_share"))
    arpu = _safe_float(inp.get("arpu"))
    upsell_to_churn_ratio = _safe_float(inp.get("upsell_to_churn_ratio"))
    cost_per_hour = _safe_float(inp.get("cost_per_hour"))

    total_agent_hours = total_agents * hours_per_agent
    impaired_hours = total_agent_hours * pct_hours_impaired
    total_calls = total_agents * calls_per_agent
    impaired_calls = total_calls * pct_calls_impaired
    retries = impaired_calls * retry_rate

    fcr_impact_factor = -retry_rate
    revised_fcr = fcr * (1.0 + fcr_impact_factor)
    csat_impact_factor = -(retry_rate * 0.90)
    revised_csat = csat * (1.0 + csat_impact_factor)
    ces_converted = ces / (ces_divisor if ces_divisor > 0 else 7.5)
    ces_impact_factor = -(retry_rate * 1.20)
    revised_ces = ces_converted * (1.0 + ces_impact_factor)

    churn_risk_factor = min(1.0, max(0.0, retry_rate * churn_risk_multiplier))
    baseline_churned_customers = customers * baseline_churn_rate
    customers_at_risk = baseline_churned_customers * churn_risk_factor
    revenue_at_risk = customers_at_risk * arpu
    active_churn_customers = baseline_churned_customers * active_churn_share
    active_churned_rev_baseline = active_churn_customers * arpu
    required_upsell_revenue = active_churned_rev_baseline * upsell_to_churn_ratio
    upsell_revenue_at_risk_shortfall = revenue_at_risk * upsell_to_churn_ratio
    labor_ineff_cost = impaired_hours * cost_per_hour
    coi = labor_ineff_cost + revenue_at_risk + upsell_revenue_at_risk_shortfall

    return {
        "total_agents": total_agents,
        "hours_per_agent": hours_per_agent,
        "total_agent_hours": total_agent_hours,
        "pct_hours_impaired": pct_hours_impaired,
        "impaired_hours": impaired_hours,
        "calls_per_agent": calls_per_agent,
        "total_calls": total_calls,
        "pct_calls_impaired": pct_calls_impaired,
        "impaired_calls": impaired_calls,
        "retry_rate": retry_rate,
        "retries": retries,
        "fcr": fcr,
        "revised_fcr": revised_fcr,
        "csat": csat,
        "revised_csat": revised_csat,
        "ces": ces,
        "ces_divisor": ces_divisor,
        "ces_converted": ces_converted,
        "revised_ces": revised_ces,
        "churn_risk_multiplier": churn_risk_multiplier,
        "churn_risk_factor": churn_risk_factor,
        "customers": customers,
        "baseline_churn_rate": baseline_churn_rate,
        "baseline_churned_customers": baseline_churned_customers,
        "active_churn_share": active_churn_share,
        "active_churn_customers": active_churn_customers,
        "arpu": arpu,
        "revenue_at_risk": revenue_at_risk,
        "upsell_to_churn_ratio": upsell_to_churn_ratio,
        "active_churned_rev_baseline": active_churned_rev_baseline,
        "required_upsell_revenue": required_upsell_revenue,
        "upsell_revenue_at_risk_shortfall": upsell_revenue_at_risk_shortfall,
        "cost_per_hour": cost_per_hour,
        "labor_inefficiency_cost": labor_ineff_cost,
        "total_cost_of_inaction": coi,
        "mode": "NRR",
    }


def render_calc_1(excel_bytes: bytes) -> None:
    st.title("Calculadora 1: Cost of Impairment (COI)")

    col_mode, col_help = st.columns([0.45, 0.55])
    with col_mode:
        coi_mode = st.radio(
            "Modelo",
            options=["Retention", "NRR"],
            horizontal=True,
            help="Use Retention para churn/revenue at risk e NRR para incluir risco de upsell.",
        )
    with col_help:
        st.info(
            "As premissas iniciais são carregadas do Excel (abas COI Retention / COI NRR) "
            "e viram inputs interativos. Os outputs são recalculados em tempo real e salvos "
            "no session_state para a Calculadora 2."
        )

if excel_bytes is not None:
    params_ret = load_params_from_excel_bytes(excel_bytes, SHEET_RETENTION)
    params_nrr = load_params_from_excel_bytes(excel_bytes, SHEET_NRR)
else:
    params_ret = {}
    params_nrr = {}

    d_total_agents = _safe_float(params_ret.get("Total Agents", 1200))
    d_hours_per_agent = _safe_float(params_ret.get("Hours per agent", 2080))
    d_pct_hours_imp = _safe_float(params_ret.get("% of hours impaired", 0.035))
    d_calls_per_agent = _safe_float(params_ret.get("Calls per agent", 25000))
    d_pct_calls_imp = _safe_float(params_ret.get("% of calls impaired*", 0.04046))
    d_customers = _safe_float(params_ret.get("Customers", 2_000_000))
    d_calls_per_customer = _safe_float(params_ret.get("Calls per customer", 15))
    d_calls_at_risk = _safe_float(params_ret.get("Calls per at risk customer**", 25.05))
    d_switch_rate = _safe_float(params_ret.get("Switching rate 2+ incidents", 0.73))
    d_base_churn = _safe_float(params_ret.get("Baseline Churn Rate", 0.15))
    d_arpu = _safe_float(params_ret.get("ARPU", 1000))
    d_cost_per_hour = _safe_float(params_ret.get("Cost per hour per agent", 20))

    d_retry_rate = _safe_float(params_nrr.get("Retry Rate*", 0.06321875))
    d_fcr = _safe_float(params_nrr.get("FCR", 0.75))
    d_csat = _safe_float(params_nrr.get("CSAT", 0.8))
    d_ces = _safe_float(params_nrr.get("CES", 5.0))
    d_base_churn_nrr = _safe_float(params_nrr.get("Baseline churn rate", d_base_churn))
    d_upsell_ratio = _safe_float(params_nrr.get("Upsell multiplier: Upsell-to-Churn Ratio", 2.2222222))
    d_cost_per_hour_nrr = _safe_float(params_nrr.get("Agent cost per hour", d_cost_per_hour))
    d_churn_risk_factor = _safe_float(params_nrr.get("Churn Risk Factor", 0.0819315))
    d_churn_risk_mult = (d_churn_risk_factor / d_retry_rate) if d_retry_rate > 0 else 1.296

    left, right = st.columns([0.58, 0.42])
    with left:
        st.subheader("Inputs")
        c1, c2, c3 = st.columns(3)

        with c1:
            total_agents = st.number_input("Total Agents", min_value=1, max_value=500_000, value=int(d_total_agents), step=10)
            hours_per_agent = st.number_input("Hours per agent (year)", min_value=1.0, max_value=10_000.0, value=float(d_hours_per_agent), step=10.0)
            pct_hours_impaired = st.slider("% of hours impaired", min_value=0.0, max_value=0.30, value=float(d_pct_hours_imp), step=0.001, format="%.3f")

        with c2:
            calls_per_agent = st.number_input("Calls per agent (year)", min_value=0.0, max_value=5_000_000.0, value=float(d_calls_per_agent), step=500.0)
            pct_calls_impaired = st.slider("% of calls impaired", min_value=0.0, max_value=0.30, value=float(d_pct_calls_imp), step=0.0005, format="%.4f")
            cost_per_hour = st.number_input(
                "Cost per hour per agent",
                min_value=0.0,
                max_value=1_000.0,
                value=float(d_cost_per_hour_nrr if coi_mode == "NRR" else d_cost_per_hour),
                step=1.0,
            )

        with c3:
            customers = st.number_input("Customers", min_value=0.0, max_value=1e9, value=float(d_customers), step=10_000.0, format="%.0f")
            baseline_churn_rate = st.slider(
                "Baseline churn rate",
                min_value=0.0,
                max_value=0.50,
                value=float(d_base_churn_nrr if coi_mode == "NRR" else d_base_churn),
                step=0.001,
                format="%.3f",
            )
            arpu = st.number_input("ARPU", min_value=0.0, max_value=1_000_000.0, value=float(d_arpu), step=10.0)

        with st.expander("Parâmetros avançados", expanded=False):
            a1, a2, a3 = st.columns(3)
            with a1:
                active_churn_share = st.slider("Active churn customers (share)", 0.0, 1.0, value=0.60, step=0.01)
                cx_driven_share = st.slider("CX-driven churns (share of active)", 0.0, 1.0, value=0.50, step=0.01)
            with a2:
                calls_per_customer = st.number_input("Calls per customer", min_value=0.0, max_value=500.0, value=float(d_calls_per_customer), step=0.5)
                calls_per_at_risk_customer = st.number_input("Calls per at-risk customer", min_value=0.0, max_value=500.0, value=float(d_calls_at_risk), step=0.5)
            with a3:
                switching_rate_2plus = st.slider("Switching rate (2+ incidents)", 0.0, 1.0, value=float(d_switch_rate), step=0.01)

        inp_common = {
            "total_agents": total_agents,
            "hours_per_agent": hours_per_agent,
            "pct_hours_impaired": pct_hours_impaired,
            "calls_per_agent": calls_per_agent,
            "pct_calls_impaired": pct_calls_impaired,
            "customers": customers,
            "baseline_churn_rate": baseline_churn_rate,
            "arpu": arpu,
            "cost_per_hour": cost_per_hour,
            "active_churn_share": active_churn_share,
            "cx_driven_share": cx_driven_share,
            "calls_per_customer": calls_per_customer,
            "calls_per_at_risk_customer": calls_per_at_risk_customer,
            "switching_rate_2plus": switching_rate_2plus,
        }

        if coi_mode == "Retention":
            outputs = calc_coi_retention(inp_common)
        else:
            with st.expander("Inputs NRR", expanded=True):
                n1, n2, n3 = st.columns(3)
                with n1:
                    retry_rate = st.slider("Retry Rate", 0.0, 0.30, value=float(d_retry_rate), step=0.0005, format="%.4f")
                    churn_risk_multiplier = st.number_input(
                        "Churn risk multiplier",
                        min_value=0.0,
                        max_value=50.0,
                        value=float(d_churn_risk_mult),
                        step=0.05,
                        help="Churn Risk Factor = min(1, Retry Rate * multiplier).",
                    )
                with n2:
                    fcr = st.slider("FCR", 0.0, 1.0, value=float(d_fcr), step=0.01)
                    csat = st.slider("CSAT", 0.0, 1.0, value=float(d_csat), step=0.01)
                with n3:
                    ces = st.number_input("CES (escala original)", min_value=0.0, max_value=10.0, value=float(d_ces), step=0.1)
                    ces_divisor = st.number_input("CES divisor (converter p/ 0-1)", min_value=1.0, max_value=20.0, value=7.5, step=0.5)
                    upsell_to_churn_ratio = st.number_input("Upsell-to-Churn Ratio", min_value=0.0, max_value=20.0, value=float(d_upsell_ratio), step=0.05)

            inp_nrr = dict(inp_common)
            inp_nrr.update({
                "retry_rate": retry_rate,
                "fcr": fcr,
                "csat": csat,
                "ces": ces,
                "ces_divisor": ces_divisor,
                "churn_risk_multiplier": churn_risk_multiplier,
                "upsell_to_churn_ratio": upsell_to_churn_ratio,
            })
            outputs = calc_coi_nrr(inp_nrr)

    with right:
        st.subheader("Outputs (tempo real)")
        m1, m2 = st.columns(2)
        with m1:
            st.metric("Labor inefficiency", fmt_currency(outputs["labor_inefficiency_cost"]))
            st.metric("Revenue at risk", fmt_currency(outputs["revenue_at_risk"]))
        with m2:
            if outputs["mode"] == "NRR":
                st.metric("Upsell shortfall risk", fmt_currency(outputs["upsell_revenue_at_risk_shortfall"]))
            st.metric("Total Cost of Inaction", fmt_currency(outputs["total_cost_of_inaction"]))

        if outputs["mode"] == "Retention":
            st.caption(f"P(2+ impaired calls) (Poisson) = {outputs['probability_2plus_impaired_calls']:.4f}")
            st.caption(f"Customers at risk = {outputs['customers_at_risk']:.0f}")
        else:
            st.caption(f"Churn Risk Factor = {outputs['churn_risk_factor']:.4f}")
            st.caption(f"Revised FCR/CSAT/CES = {outputs['revised_fcr']:.3f} / {outputs['revised_csat']:.3f} / {outputs['revised_ces']:.3f}")

        st.subheader("Gráficos (Calc 1)")

        if outputs["mode"] == "Retention":
            df_components = pd.DataFrame({
                "Componente": ["Labor inefficiency", "Revenue at risk"],
                "Valor": [outputs["labor_inefficiency_cost"], outputs["revenue_at_risk"]],
            })
        else:
            df_components = pd.DataFrame({
                "Componente": ["Labor inefficiency", "Revenue at risk", "Upsell shortfall risk"],
                "Valor": [
                    outputs["labor_inefficiency_cost"],
                    outputs["revenue_at_risk"],
                    outputs["upsell_revenue_at_risk_shortfall"],
                ],
            })

        fig1 = px.bar(df_components, x="Componente", y="Valor", text_auto=".2s", title="Decomposição do Cost of Inaction")
        fig1.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig1, use_container_width=True)

        mitigation = st.slider("Simular mitigação (%)", 0.0, 1.0, value=0.0, step=0.01, format="%.2f")
        before = outputs["total_cost_of_inaction"]
        after = before * (1.0 - mitigation)
        df_before_after = pd.DataFrame({"Estado": ["Before", "After"], "Total COI": [before, after]})

        fig2 = px.bar(df_before_after, x="Estado", y="Total COI", text_auto=".2s", title="Before vs After (simulação)")
        fig2.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    st.session_state["calc1_outputs"] = outputs
    st.session_state["calc1_components_df"] = df_components


# -----------------------------
# Calculator 2: Business Impact
# -----------------------------
def calculate_business_impact(calc1_out: dict, scenario: str, bi_inp: dict) -> dict:
    mult = scenario_multipliers(scenario)

    total_coi = float(calc1_out.get("total_cost_of_inaction", 0.0))
    labor = float(calc1_out.get("labor_inefficiency_cost", 0.0))
    revenue = float(calc1_out.get("revenue_at_risk", 0.0))
    upsell_risk = float(calc1_out.get("upsell_revenue_at_risk_shortfall", 0.0))

    churn_reduction = _safe_float(bi_inp.get("churn_reduction_pct"))
    revenue_impact = _safe_float(bi_inp.get("revenue_impact_pct"))
    labor_recovery = _safe_float(bi_inp.get("labor_recovery_pct"))
    annual_solution_cost = _safe_float(bi_inp.get("annual_solution_cost"))

    recovered_revenue = revenue * churn_reduction * mult["churn_mult"]
    recovered_upsell = upsell_risk * revenue_impact * mult["rev_mult"]
    labor_savings = labor * labor_recovery * mult["eff_mult"]

    gross_benefit = recovered_revenue + recovered_upsell + labor_savings
    net_impact = gross_benefit - annual_solution_cost
    remaining_coi = max(0.0, total_coi - gross_benefit)

    roi = (net_impact / annual_solution_cost) if annual_solution_cost > 0 else np.nan
    payback_months = (annual_solution_cost / (gross_benefit / 12.0)) if gross_benefit > 0 else np.inf

    return {
        "scenario": scenario,
        "multipliers": str(mult),
        "total_coi": total_coi,
        "labor": labor,
        "revenue": revenue,
        "upsell_risk": upsell_risk,
        "recovered_revenue": recovered_revenue,
        "recovered_upsell": recovered_upsell,
        "labor_savings": labor_savings,
        "gross_benefit": gross_benefit,
        "annual_solution_cost": annual_solution_cost,
        "net_impact": net_impact,
        "remaining_coi": remaining_coi,
        "roi": roi,
        "payback_months": payback_months,
    }


def make_waterfall(bi_out: dict) -> go.Figure:
    fig = go.Figure(go.Waterfall(
        measure=["relative", "relative", "relative", "relative", "total"],
        x=["Recovered revenue", "Recovered upsell", "Labor savings", "Solution cost", "Net impact"],
        y=[
            bi_out["recovered_revenue"],
            bi_out["recovered_upsell"],
            bi_out["labor_savings"],
            -bi_out["annual_solution_cost"],
            bi_out["net_impact"],
        ],
        connector={"line": {"color": "rgba(0,0,0,0.25)"}},
    ))
    fig.update_layout(title="Waterfall: impacto anual (Business Impact)", height=360, margin=dict(l=10, r=10, t=60, b=10))
    return fig


def df_to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            safe_name = name[:31] if name else "Sheet1"
            df.to_excel(writer, index=False, sheet_name=safe_name)
    buf.seek(0)
    return buf.read()


def build_summary_markdown(calc1_out: dict, bi_out: dict) -> str:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Resolve os valores ANTES do f-string para evitar backslash em expressão
    pb = bi_out.get("payback_months", np.inf)
    pb_str = "inf" if np.isinf(pb) else f"{pb:.1f}"

    roi = bi_out.get("roi", np.nan)
    roi_str = "n/a" if np.isnan(roi) else f"{roi:.2f}x"

    upsell_line = ""
    if calc1_out.get("mode") == "NRR":
        upsell_line = f"- Upsell shortfall risk: {fmt_currency(calc1_out.get('upsell_revenue_at_risk_shortfall', 0.0))}\n"

    lines = (
        f"# Resumo - COI & Business Impact\n"
        f"- Gerado em: {ts}\n"
        f"\n"
        f"## Calculadora 1 (COI)\n"
        f"- Mode: {calc1_out.get('mode')}\n"
        f"- Labor inefficiency: {fmt_currency(calc1_out.get('labor_inefficiency_cost', 0.0))}\n"
        f"- Revenue at risk: {fmt_currency(calc1_out.get('revenue_at_risk', 0.0))}\n"
        f"{upsell_line}"
        f"- Total Cost of Inaction: {fmt_currency(calc1_out.get('total_cost_of_inaction', 0.0))}\n"
        f"\n"
        f"## Calculadora 2 (Business Impact)\n"
        f"- Cenario: {bi_out.get('scenario')}\n"
        f"- Gross benefit: {fmt_currency(bi_out.get('gross_benefit', 0.0))}\n"
        f"- Annual solution cost: {fmt_currency(bi_out.get('annual_solution_cost', 0.0))}\n"
        f"- Net impact: {fmt_currency(bi_out.get('net_impact', 0.0))}\n"
        f"- Payback (meses): {pb_str}\n"
        f"- ROI: {roi_str}\n"
    )
    return lines



def _sanitize_pdf(text: str) -> str:
    """Converte caracteres fora do latin-1 para equivalentes ASCII."""
    replacements = {
        "∞": "inf", "×": "x", "≤": "<=", "≥": ">=",
        "\u2019": "'", "\u2018": "'", "\u201c": '"', "\u201d": '"',
        "\u2013": "-", "\u2014": "-", "R$": "BRL",
    }
    for char, rep in replacements.items():
        text = text.replace(char, rep)
    return text.encode("latin-1", errors="replace").decode("latin-1")


def markdown_to_pdf_bytes(md_text: str) -> bytes:
    pdf = FPDF()
    pdf.set_margins(left=15, top=15, right=15)
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.add_page()
    usable_w = pdf.w - pdf.l_margin - pdf.r_margin

    for line in md_text.splitlines():
        clean = _sanitize_pdf(line.replace("#", "").replace("*", "").strip())
        if not clean:
            pdf.ln(4)
            continue
        if line.lstrip().startswith("#"):
            pdf.set_font("Helvetica", style="B", size=13)
        else:
            pdf.set_font("Helvetica", size=11)
        pdf.multi_cell(usable_w, 6, clean)

    raw = pdf.output(dest="S")
    return bytes(raw) if isinstance(raw, (bytes, bytearray)) else raw.encode("latin-1", errors="replace")



def render_calc_2() -> None:
    st.title("Calculadora 2: Business Impact")

    calc1_out = st.session_state.get("calc1_outputs")
    if not calc1_out:
        st.warning("⚠️ Execute a Calculadora 1 primeiro para popular os outputs no session_state.")
        return

    top_l, top_r = st.columns([0.62, 0.38])
    with top_l:
        st.subheader("Inputs & Cenários")
        scenario = st.radio("Cenário", ["Conservador", "Moderado", "Agressivo"], horizontal=True)

        i1, i2, i3 = st.columns(3)
        with i1:
            churn_reduction_pct = st.slider("Churn reduction sobre Revenue at risk", 0.0, 1.0, value=0.25, step=0.01)
            revenue_impact_pct = st.slider("Revenue impact sobre Upsell shortfall (NRR)", 0.0, 1.0, value=0.20, step=0.01)
        with i2:
            labor_recovery_pct = st.slider("Labor recovery sobre inefficiency", 0.0, 1.0, value=0.30, step=0.01)
            annual_solution_cost = st.number_input("Custo anual da solução", min_value=0.0, max_value=1e9, value=5_000_000.0, step=50_000.0)
        with i3:
            st.caption("Base (Calc 1)")
            st.write(f"- Total COI: {fmt_currency(calc1_out.get('total_cost_of_inaction', 0.0))}")
            st.write(f"- Labor: {fmt_currency(calc1_out.get('labor_inefficiency_cost', 0.0))}")
            st.write(f"- Revenue at risk: {fmt_currency(calc1_out.get('revenue_at_risk', 0.0))}")
            if calc1_out.get("mode") == "NRR":
                st.write(f"- Upsell risk: {fmt_currency(calc1_out.get('upsell_revenue_at_risk_shortfall', 0.0))}")
            else:
                st.write("- Upsell risk: n/a (modo Retention)")

        bi_inp = dict(
            churn_reduction_pct=churn_reduction_pct,
            revenue_impact_pct=revenue_impact_pct,
            labor_recovery_pct=labor_recovery_pct,
            annual_solution_cost=annual_solution_cost,
        )
        bi_out = calculate_business_impact(calc1_out, scenario, bi_inp)
        st.session_state["calc2_outputs"] = bi_out

    with top_r:
        st.subheader("KPIs")
        k1, k2 = st.columns(2)
        with k1:
            st.metric("Gross benefit", fmt_currency(bi_out["gross_benefit"]))
            st.metric("Remaining COI", fmt_currency(bi_out["remaining_coi"]))
        with k2:
            st.metric("Net impact", fmt_currency(bi_out["net_impact"]))
            pb = bi_out["payback_months"]
            st.metric("Payback (months)", "∞" if np.isinf(pb) else f"{pb:.1f}")
        roi = bi_out["roi"]
        st.caption(f"ROI: {'n/a' if np.isnan(roi) else f'{roi:.2f}x'}")

    st.subheader("Gráficos (Calc 2)")
    cA, cB = st.columns([0.55, 0.45])
    with cA:
        st.plotly_chart(make_waterfall(bi_out), use_container_width=True)
    with cB:
        scenarios = ["Conservador", "Moderado", "Agressivo"]
        rows = [{"Scenario": s, "Net impact": calculate_business_impact(calc1_out, s, bi_inp)["net_impact"]} for s in scenarios]
        df_scen = pd.DataFrame(rows)
        fig = px.bar(df_scen, x="Scenario", y="Net impact", text_auto=".2s", title="Net impact por cenário")
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=60, b=10))
        st.plotly_chart(fig, use_container_width=True)

    months = np.arange(1, 13)
    cum_net = np.cumsum(np.repeat((bi_out["gross_benefit"] - bi_out["annual_solution_cost"]) / 12.0, 12))
    df_month = pd.DataFrame({"Month": months, "Cumulative net": cum_net})
    fig_line = px.area(df_month, x="Month", y="Cumulative net", title="Cumulativo (12 meses): benefício líquido")
    fig_line.update_layout(height=300, margin=dict(l=10, r=10, t=60, b=10))
    st.plotly_chart(fig_line, use_container_width=True)

    st.subheader("Exportação")

    calc1_components_df = st.session_state.get("calc1_components_df", pd.DataFrame())
    consolidated = pd.DataFrame([
        {"metric": "calc1_total_coi",        "value": calc1_out.get("total_cost_of_inaction", 0.0)},
        {"metric": "calc1_labor",             "value": calc1_out.get("labor_inefficiency_cost", 0.0)},
        {"metric": "calc1_revenue_risk",      "value": calc1_out.get("revenue_at_risk", 0.0)},
        {"metric": "calc1_upsell_risk",       "value": calc1_out.get("upsell_revenue_at_risk_shortfall", 0.0)},
        {"metric": "calc2_gross_benefit",     "value": bi_out.get("gross_benefit", 0.0)},
        {"metric": "calc2_solution_cost",     "value": bi_out.get("annual_solution_cost", 0.0)},
        {"metric": "calc2_net_impact",        "value": bi_out.get("net_impact", 0.0)},
        {"metric": "calc2_payback_months",    "value": bi_out.get("payback_months", np.inf)},
    ])

    csv_bytes   = consolidated.to_csv(index=False).encode("utf-8")
    excel_bytes = df_to_excel_bytes({
        "calc1_outputs":    pd.DataFrame([calc1_out]),
        "calc1_components": calc1_components_df,
        "calc2_outputs":    pd.DataFrame([bi_out]),
        "consolidated":     consolidated,
    })
    md        = build_summary_markdown(calc1_out, bi_out)
    pdf_bytes = markdown_to_pdf_bytes(md)

    e1, e2, e3, e4 = st.columns(4)
    with e1:
        st.download_button("⬇️ CSV",      data=csv_bytes,   file_name="results.csv",     mime="text/csv")
    with e2:
        st.download_button("⬇️ Excel",    data=excel_bytes, file_name="results.xlsx",    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e3:
        st.download_button("⬇️ Markdown", data=md.encode(), file_name="summary.md",      mime="text/markdown")
    with e4:
        st.download_button("⬇️ PDF",      data=pdf_bytes,   file_name="summary.pdf",     mime="application/pdf")


# -----------------------------
# Main
# -----------------------------
def main() -> None:
    st.sidebar.title("Navegação")
    page = st.sidebar.radio(
        "Escolha",
        ["Calculadora 1: Cost of Impairment", "Calculadora 2: Business Impact"],
        index=0,
    )

    excel_bytes = get_excel_bytes_from_sidebar()
    # Não bloqueia mais: modo manual funciona sem Excel

    if page.startswith("Calculadora 1"):
        render_calc_1(excel_bytes)
    else:
        render_calc_2()



if __name__ == "__main__":
    st.session_state.setdefault("calc1_outputs", None)
    st.session_state.setdefault("calc2_outputs", None)
    main()
