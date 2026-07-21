from __future__ import annotations

import streamlit as st

from monthly_reporting_page import render_monthly_reporting_page


st.set_page_config(
    page_title="Sarmayacar FundOps Dashboard",
    layout="wide",
)

tabs = st.tabs(["Monthly Reporting"])

with tabs[0]:
    render_monthly_reporting_page()
