from __future__ import annotations

import os

import streamlit as st

from monthly_reporting_page import render_monthly_reporting_page
from quarterly_report import render_quarterly_report_page

DASHBOARD_URL = os.environ.get(
    "SARMAYACAR_DASHBOARD_URL",
    "https://sarmayacar-dashboard.vercel.app/whats-new",
).strip()


st.set_page_config(
    page_title="Sarmayacar FundOps Dashboard",
    layout="wide",
)

with st.sidebar:
    st.markdown("### Connected Dashboard")
    st.link_button(
        "Open Portfolio Dashboard",
        DASHBOARD_URL,
        use_container_width=True,
    )
    st.caption("Opens the Sarmayacar Vercel dashboard in a new tab.")
    st.divider()

    workflow = st.radio(
        "Workflow",
        ["Quarterly Report", "Monthly Reporting"],
        key="fundops_workflow",
    )
    st.caption("Only the selected workflow loads, so the app stays responsive.")


if workflow == "Quarterly Report":
    render_quarterly_report_page()
else:
    render_monthly_reporting_page()
