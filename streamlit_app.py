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

tabs = st.tabs(["Monthly Reporting", "Quarterly Report"])

with tabs[0]:
    render_monthly_reporting_page()

with tabs[1]:
    render_quarterly_report_page()
