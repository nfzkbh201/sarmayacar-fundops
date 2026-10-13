"""Verify an authorized hosted bridge using synthetic data and private local settings."""
import json
import sys
import time
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from figma_handoff import BridgeClient, CLIENT_USER_AGENT

settings = json.loads((ROOT / ".figma_bridge/cloudflare.json").read_text())
url = settings["service_url"]
client = BridgeClient(url, settings["publisher_token"])
with urlopen(Request(url + "/health", headers={"User-Agent": CLIENT_USER_AGENT}), timeout=20) as response:
    assert json.load(response) == {"status": "ok"}
try:
    urlopen(Request(url + "/v1/connections", headers={"User-Agent": CLIENT_USER_AGENT}), timeout=20)
except HTTPError as error:
    assert error.code == 401
else:
    raise AssertionError("Unauthenticated request was not rejected")

connection = client.request("POST", "/v1/connections", {"name": "Temporary hosted verification"})
reader = BridgeClient(url, connection["pairing_code"])
try:
    reader.request("POST", "/v1/plugin/connect", {
        "target_id": "synthetic-verification", "page_id": "synthetic-page",
        "page_name": "Synthetic verification", "file_name": "No real report",
    })
    for size in (1000, 200000):
        payload = {"schema_version": "sarmayacar.quarterly_report.v1",
                   "report": {"quarter": "Q2 2026", "months": ["Apr 26", "May 26", "Jun 26"]},
                   "companies": [{"name": "Synthetic test", "description": "x" * size}]}
        request = {"request_id": f"synthetic-{size}", "payload": payload}
        start = time.monotonic()
        job = client.request("POST", f"/v1/connections/{connection['id']}/jobs", request)
        assert client.request("POST", f"/v1/connections/{connection['id']}/jobs", request)["id"] == job["id"]
        assert reader.request("GET", "/v1/plugin/jobs")["job"]["id"] == job["id"]
        claimed = reader.request("POST", f"/v1/plugin/jobs/{job['id']}/claim", {"runner": "synthetic-runner"})
        assert claimed["payload"] == payload
        result = {"runner": "synthetic-runner", "status": "succeeded", "page_count": 1}
        for _ in range(2):
            assert reader.request("POST", f"/v1/plugin/jobs/{job['id']}/result", result)["status"] == "succeeded"
        assert reader.request("GET", "/v1/plugin/jobs")["job"] is None
        print(f"Hosted {size}-character report: send, retry, claim and acknowledgement passed in {time.monotonic() - start:.1f}s.")
finally:
    client.request("DELETE", f"/v1/connections/{connection['id']}")
print("Hosted checks passed; synthetic connection revoked. Actual Figma rendering is a separate check.")
