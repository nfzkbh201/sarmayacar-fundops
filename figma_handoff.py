"""Streamlit controls for the private Figma report connection."""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import uuid
from urllib.error import HTTPError, URLError
from urllib.parse import urlsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener

import streamlit as st

CLIENT_USER_AGENT = "Sarmayacar-FundOps/1.0"


class BridgeError(RuntimeError):
    def __init__(self, message: str, retryable: bool = True):
        super().__init__(message)
        self.retryable = retryable


class NoRedirects(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


class BridgeClient:
    def __init__(self, url: str, token: str):
        parsed = urlsplit(url)
        local = parsed.hostname in ("localhost", "127.0.0.1", "::1")
        if (parsed.scheme != "https" and not (local and parsed.scheme == "http")) or not parsed.hostname or parsed.username or parsed.password or parsed.query or parsed.fragment:
            raise BridgeError("The Figma connection needs an HTTPS service URL (HTTP is allowed on localhost).")
        self.url = url.rstrip("/")
        self.token = token

    def request(self, method: str, path: str, data=None):
        try:
            encoded = json.dumps(data, allow_nan=False).encode() if data is not None else None
        except (TypeError, ValueError):
            raise BridgeError("The report contains invalid JSON values. Regenerate it before sending.", retryable=False) from None
        request = Request(self.url + path, data=encoded, method=method,
                          headers={"Authorization": f"Bearer {self.token}", "Content-Type": "application/json",
                                   "Accept": "application/json", "User-Agent": CLIENT_USER_AGENT})
        try:
            with build_opener(NoRedirects).open(request, timeout=8) as response:
                return json.load(response)
        except HTTPError as error:
            try:
                detail = json.load(error).get("detail", "Connection request failed.")
            except (ValueError, AttributeError):
                detail = "Connection request failed."
            raise BridgeError(f"{detail} (HTTP {error.code})", retryable=error.code >= 500 or error.code == 429) from None
        except (URLError, TimeoutError, OSError, ValueError):
            raise BridgeError("Cannot reach the Figma connection service. No completion has been confirmed; retry when it is reachable.") from None


def setting(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if value:
        return value
    try:
        return str(st.secrets.get(name, "")).strip()
    except (FileNotFoundError, KeyError):
        return ""


def content_fingerprint(payload: dict) -> str:
    # A Streamlit rerun changes generated_at even when the reviewed report is identical.
    content = {k: v for k, v in payload.items() if k != "generated_at"}
    return hashlib.sha256(json.dumps(content, sort_keys=True).encode()).hexdigest()


@st.fragment(run_every="5s")
def render_transfer_status(client: BridgeClient, connection_id: str):
    try:
        connections = client.request("GET", "/v1/connections")["connections"]
        connection = next((item for item in connections if item["id"] == connection_id), None)
        if not connection:
            st.warning("This Figma connection has been revoked.")
            return
        if connection["online"]:
            st.caption("Figma plugin connected")
        else:
            st.caption("Figma plugin offline. Open it on the paired page to receive the report.")
        job = connection.get("latest_job")
        if not job:
            return
        result, quarter = job.get("result") or {}, job["quarter"]
        if job["status"] == "succeeded":
            st.success(f"{quarter}: updated {result.get('page_count', 0)} report pages in Figma.")
            if result.get("warning_count"):
                st.warning(f"{result['warning_count']} report warnings. Review them in the Figma plugin.")
        elif job["status"] == "failed":
            st.error(f"{quarter}: import failed. {result.get('message', '')}")
        elif job["status"] == "processing":
            st.info(f"{quarter}: updating the Figma report...")
        elif job["status"] == "queued":
            st.info(f"{quarter}: waiting for the paired Figma plugin.")
        else:
            st.warning(f"{quarter}: transfer {job['status']}. Send the report again.")
    except BridgeError as error:
        st.warning(str(error))


def render_figma_connection(payload: dict):
    url, token, access = (setting(name) for name in (
        "FIGMA_BRIDGE_URL", "FIGMA_BRIDGE_PUBLISHER_TOKEN", "FIGMA_BRIDGE_ACCESS_CODE"))
    if not url or not token or len(access) < 12:
        st.info("Direct Figma connection is awaiting setup. File import remains available below.")
        return
    entered = st.text_input("Figma access code", type="password", key="figma_access_code")
    if not hmac.compare_digest(entered.encode(), access.encode()):
        if entered:
            st.error("The Figma access code does not match.")
        return
    try:
        client = BridgeClient(url, token)
        connections = client.request("GET", "/v1/connections")["connections"]
        paired = [item for item in connections if item["target"]]
        with st.expander("Connect a Figma page", expanded=not paired):
            name = st.text_input("Connection name", value="Quarterly report", key="figma_connection_name")
            if st.button("Create pairing code", key="figma_pair"):
                st.session_state["figma_pairing"] = client.request("POST", "/v1/connections", {"name": name})
            pairing = st.session_state.get("figma_pairing")
            if pairing:
                st.text_input("Service URL", value=url, disabled=True)
                st.caption("Private pairing code")
                st.code(pairing["pairing_code"], language=None)
                st.caption("Enter these in the Figma plugin once, then refresh destinations.")
            if st.button("Refresh destinations", key="figma_refresh"):
                st.rerun()
        if not paired:
            return
        by_id = {item["id"]: item for item in paired}
        destination = st.selectbox("Figma destination", list(by_id), key="figma_destination",
            format_func=lambda key: f"{by_id[key]['name']} / {by_id[key]['target']['file_name']} / {by_id[key]['target']['page_name']}")
        st.caption("Updates replace the generated pages for this quarter. Other quarters and unlinked pages are kept.")
        pending_key = f"figma_pending_{destination}"
        pending = st.session_state.get(pending_key)
        fingerprint = content_fingerprint(payload)
        if pending and pending["fingerprint"] != fingerprint:
            st.warning("The reviewed report has changed. Resolve the previous transfer before sending the changed report.")
        if st.button("Retry send" if pending else "Send to Figma", type="primary", key="figma_send", width="stretch"):
            if not pending:
                pending = {"request_id": str(uuid.uuid4()), "fingerprint": fingerprint, "payload": payload}
                st.session_state[pending_key] = pending
            try:
                client.request("POST", f"/v1/connections/{destination}/jobs",
                               {"request_id": pending["request_id"], "payload": pending["payload"]})
            except BridgeError as error:
                if not error.retryable:
                    st.session_state.pop(pending_key, None)
                raise
            st.session_state.pop(pending_key, None)
        render_transfer_status(client, destination)
        with st.expander("Connection settings"):
            if st.checkbox("Revoke this connection", key=f"figma_revoke_confirm_{destination}"):
                if st.button("Revoke", key=f"figma_revoke_{destination}"):
                    client.request("DELETE", f"/v1/connections/{destination}")
                    st.session_state.pop("figma_pairing", None)
                    st.rerun()
    except BridgeError as error:
        st.error(str(error))
