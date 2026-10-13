"""Local handoff screen using the same controls as the Quarterly Report tab."""
import json
import os
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import streamlit as st
from figma_handoff import render_figma_connection

st.set_page_config(page_title="Quarterly Report Connection", layout="centered")
st.title("Quarterly Report Connection")
pack_path = os.environ.get("FIGMA_PREVIEW_PACK")
if pack_path:
    payload = json.loads(Path(pack_path).read_text())
    st.caption(Path(pack_path).name)
else:
    upload = st.file_uploader("Quarterly report data pack", type=["json"])
    if upload is None:
        st.stop()
    payload = json.loads(upload.getvalue())
if payload.get("schema_version") != "sarmayacar.quarterly_report.v1":
    st.error("Select a Sarmayacar quarterly report data pack.")
    st.stop()
st.subheader(payload.get("report", {}).get("quarter", "Quarterly report"))
st.warning("Connection test using an existing data pack. This does not check the latest financial mappings or formatting.")
# This preview is bound to loopback by the launcher, never deployed to Streamlit Cloud.
st.session_state.setdefault("figma_access_code", os.environ["FIGMA_BRIDGE_ACCESS_CODE"])
render_figma_connection(payload)
