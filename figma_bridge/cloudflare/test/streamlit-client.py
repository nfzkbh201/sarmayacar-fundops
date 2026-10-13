"""Exercise the unchanged dashboard client against a local Cloudflare runtime."""
import os
import sys
from pathlib import Path
from urllib.parse import urlsplit

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from figma_handoff import BridgeClient

url = os.environ["FIGMA_TEST_BRIDGE_URL"]
assert urlsplit(url).hostname == "127.0.0.1"
client = BridgeClient(url, os.environ["FIGMA_TEST_PUBLISHER_TOKEN"])
connection = client.request("POST", "/v1/connections", {"name": "Python client test"})
reader = BridgeClient(url, connection["pairing_code"])
try:
    reader.request("POST", "/v1/plugin/connect", {
        "target_id": "synthetic", "page_id": "1:2", "page_name": "Test", "file_name": "Test"
    })
    payload = {"schema_version": "sarmayacar.quarterly_report.v1",
               "report": {"quarter": "Q2 2026", "months": ["Jun 26"]},
               "companies": [{"name": "Synthetic company"}]}
    job = client.request("POST", f"/v1/connections/{connection['id']}/jobs",
                         {"request_id": "python-one", "payload": payload})
    claimed = reader.request("POST", f"/v1/plugin/jobs/{job['id']}/claim", {"runner": "test"})
    assert claimed["payload"] == payload
    reader.request("POST", f"/v1/plugin/jobs/{job['id']}/result",
                   {"runner": "test", "status": "succeeded", "page_count": 1})
    listing = client.request("GET", "/v1/connections")["connections"]
    current = next(item for item in listing if item["id"] == connection["id"])
    assert current["latest_job"]["status"] == "succeeded"
    assert current["latest_job"]["result"]["page_count"] == 1
    print("Streamlit BridgeClient -> Cloudflare runtime round-trip passed.")
finally:
    client.request("DELETE", f"/v1/connections/{connection['id']}")
