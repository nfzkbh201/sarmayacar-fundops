import json
from pathlib import Path
import tempfile
import unittest

from fastapi.testclient import TestClient

from figma_bridge.service import create_app, MAX_BODY, Store


PAYLOAD = {
    "schema_version": "sarmayacar.quarterly_report.v1",
    "report": {"quarter": "Q2 2026", "months": ["Jan 26", "Feb 26", "Mar 26", "Apr 26", "May 26", "Jun 26"]},
    "companies": [{"name": "Test company", "metrics": []}],
}


class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "bridge.sqlite3"
        self.token = "test-publisher-token-" * 3
        self.client = TestClient(create_app(self.path, self.token))
        self.publisher = {"Authorization": f"Bearer {self.token}"}
        self.connection, self.reader = self.pair("Report A")

    def tearDown(self):
        self.client.close()
        self.tmp.cleanup()

    def pair(self, name):
        result = self.client.post("/v1/connections", headers=self.publisher, json={"name": name})
        self.assertEqual(result.status_code, 200)
        connection = result.json()
        reader = {"Authorization": "Bearer " + connection["pairing_code"]}
        response = self.client.post("/v1/plugin/connect", headers=reader, json={
            "target_id": name, "page_id": "1:2", "page_name": "Quarterly Report", "file_name": "Test file",
        })
        self.assertEqual(response.status_code, 200)
        return connection, reader

    def publish(self, request_id="request-1", payload=None):
        return self.client.post(f"/v1/connections/{self.connection['id']}/jobs", headers=self.publisher,
                                json={"request_id": request_id, "payload": payload or PAYLOAD})

    def claim(self, job_id, runner="runner-1", reader=None):
        return self.client.post(f"/v1/plugin/jobs/{job_id}/claim", headers=reader or self.reader, json={"runner": runner})

    def finish(self, job_id, status="succeeded", runner="runner-1"):
        return self.client.post(f"/v1/plugin/jobs/{job_id}/result", headers=self.reader,
            json={"runner": runner, "status": status, "page_count": 8, "message": "Finished", "warnings": []})

    def test_complete_roundtrip_and_retention(self):
        job = self.publish().json()
        waiting = self.client.get("/v1/plugin/jobs", headers=self.reader).json()["job"]
        self.assertEqual(waiting["id"], job["id"])
        self.assertNotIn("payload", waiting)
        self.assertEqual(self.claim(job["id"]).json()["payload"], PAYLOAD)
        self.assertEqual(self.finish(job["id"]).json()["status"], "succeeded")
        self.assertEqual(self.finish(job["id"]).status_code, 200)
        listing = self.client.get("/v1/connections", headers=self.publisher).json()["connections"][0]
        self.assertEqual(listing["latest_job"]["result"]["page_count"], 8)
        with Store(self.path).db() as db:
            self.assertIsNone(db.execute("SELECT payload FROM jobs").fetchone()[0])

    def test_auth_roles_isolation_and_secrets(self):
        self.assertEqual(self.client.get("/v1/connections").status_code, 401)
        self.assertEqual(self.client.get("/v1/connections", headers=self.reader).status_code, 401)
        self.assertEqual(self.client.get("/v1/plugin/jobs", headers=self.publisher).status_code, 401)
        _, other = self.pair("Report B")
        job = self.publish().json()
        self.assertIsNone(self.client.get("/v1/plugin/jobs", headers=other).json()["job"])
        self.assertEqual(self.claim(job["id"], reader=other).status_code, 409)
        listing = self.client.get("/v1/connections", headers=self.publisher)
        self.assertNotIn(self.connection["pairing_code"], listing.text)
        self.assertEqual(listing.headers["cache-control"], "no-store")
        self.assertNotIn(self.connection["pairing_code"].split(".")[1].encode(), self.path.read_bytes())

    def test_cannot_retarget_paired_page(self):
        response = self.client.post("/v1/plugin/connect", headers=self.reader, json={
            "target_id": "another-file", "page_id": "1:2", "page_name": "Other", "file_name": "Other",
        })
        self.assertEqual(response.status_code, 409)

    def test_publish_idempotency_and_queue(self):
        first = self.publish().json()
        self.assertEqual(self.publish().json()["id"], first["id"])
        changed = {**PAYLOAD, "companies": [{"name": "Changed"}]}
        self.assertEqual(self.publish(payload=changed).status_code, 409)
        self.assertEqual(self.publish("request-2").status_code, 409)
        self.claim(first["id"])
        self.finish(first["id"])
        self.assertEqual(self.publish("request-2").status_code, 200)

    def test_lease_recovery_and_old_runner_cannot_ack(self):
        job = self.publish().json()
        self.assertEqual(self.claim(job["id"]).status_code, 200)
        self.assertEqual(self.claim(job["id"], "runner-2").status_code, 409)
        with Store(self.path).db() as db:
            db.execute("UPDATE jobs SET lease_until=0")
        self.assertEqual(self.claim(job["id"], "runner-2").status_code, 200)
        self.assertEqual(self.finish(job["id"]).status_code, 409)
        self.assertEqual(self.finish(job["id"], runner="runner-2").status_code, 200)

    def test_expiration_and_revocation(self):
        self.publish()
        with Store(self.path).db() as db:
            db.execute("UPDATE jobs SET created=0")
        self.assertIsNone(self.client.get("/v1/plugin/jobs", headers=self.reader).json()["job"])
        self.assertEqual(self.publish("request-2").status_code, 200)
        self.client.delete(f"/v1/connections/{self.connection['id']}", headers=self.publisher)
        self.assertEqual(self.client.get("/v1/plugin/jobs", headers=self.reader).status_code, 401)
        with Store(self.path).db() as db:
            self.assertTrue(all(row[0] is None for row in db.execute("SELECT payload FROM jobs")))

    def test_validation_size_and_browser_preflight(self):
        self.assertEqual(self.publish(payload={"bad": True}).status_code, 400)
        self.assertEqual(self.publish(payload={**PAYLOAD, "report": {"quarter": "Q9 2026"}}).status_code, 400)
        response = self.client.post("/v1/connections", headers=self.publisher, content="x" * (MAX_BODY + 1))
        self.assertEqual(response.status_code, 413)
        response = self.client.options("/v1/plugin/jobs", headers={"Origin": "null",
            "Access-Control-Request-Method": "GET", "Access-Control-Request-Headers": "authorization,content-type"})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers["access-control-allow-origin"], "*")
        self.assertEqual(self.client.post("/v1/connections", headers=self.publisher, content='{"name": NaN}').status_code, 400)

    def test_persists_across_service_restart(self):
        job = self.publish().json()
        with TestClient(create_app(self.path, self.token)) as restarted:
            waiting = restarted.get("/v1/plugin/jobs", headers=self.reader).json()["job"]
        self.assertEqual(waiting["id"], job["id"])


if __name__ == "__main__":
    unittest.main()
