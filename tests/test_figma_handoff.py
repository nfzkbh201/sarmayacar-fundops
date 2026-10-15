import os
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from figma_handoff import BridgeClient, BridgeError, CLIENT_USER_AGENT, content_fingerprint

CONFIG = {"FIGMA_BRIDGE_URL": "https://bridge.example", "FIGMA_BRIDGE_PUBLISHER_TOKEN": "secret" * 10,
          "FIGMA_BRIDGE_ACCESS_CODE": "team-access-code"}
SCRIPT = '''
from figma_handoff import render_figma_connection
render_figma_connection({"schema_version": "sarmayacar.quarterly_report.v1", "report": {"quarter": "Q2 2026"}})
'''
CONNECTION = {"id": "destination", "name": "Quarterly", "target": {"file_name": "Report", "page_name": "Report"},
              "online": True, "latest_job": None}


def unlock(app, code="team-access-code"):
    app.text_input(key="figma_access_code").set_value(code)
    return app.button(key="FormSubmitter:figma_unlock-Unlock Figma").click().run()


class HandoffTests(unittest.TestCase):
    def test_client_identifies_itself_to_hosted_service(self):
        with patch("figma_handoff.build_opener") as opener, patch("figma_handoff.json.load", return_value={}):
            BridgeClient("https://bridge.example", "secret").request("GET", "/v1/connections")
        request = opener.return_value.open.call_args.args[0]
        self.assertEqual(request.get_header("User-agent"), CLIENT_USER_AGENT)
        self.assertEqual(request.get_header("Accept"), "application/json")

    def test_unconfigured_panel_needs_no_service(self):
        with patch.dict(os.environ, {key: "" for key in CONFIG}), patch("figma_handoff.setting", return_value=""):
            app = AppTest.from_string(SCRIPT).run()
        self.assertFalse(app.exception)
        self.assertIn("awaiting setup", app.info[0].value)

    def test_access_gate_prevents_network_calls(self):
        with patch.dict(os.environ, CONFIG), patch.object(BridgeClient, "request") as request:
            app = AppTest.from_string(SCRIPT).run()
            self.assertTrue(app.button(key="figma_send").disabled)
            unlock(app, "wrong")
            self.assertFalse(app.exception)
            self.assertIn("does not match", app.error[0].value)
            request.assert_not_called()

    def test_connected_send_and_confirmation(self):
        calls = []
        def request(client, method, path, data=None):
            calls.append((method, path, data))
            if method == "POST":
                CONNECTION["latest_job"] = {"id": "job", "quarter": "Q2 2026", "status": "succeeded", "result": {"page_count": 8}}
                return CONNECTION["latest_job"]
            return {"connections": [CONNECTION]}
        CONNECTION["latest_job"] = None
        with patch.dict(os.environ, CONFIG), patch.object(BridgeClient, "request", request):
            app = AppTest.from_string(SCRIPT).run()
            unlock(app)
            self.assertFalse(app.button(key="figma_send").disabled)
            self.assertFalse(any(method == "POST" for method, _, _ in calls))
            app.button(key="figma_send").click().run()
            self.assertFalse(app.exception)
            self.assertIn("updated 8 report pages", app.success[0].value)
        posts = [call for call in calls if call[0] == "POST"]
        self.assertEqual(len(posts), 1)
        self.assertEqual(posts[0][1], "/v1/connections/destination/jobs")
        self.assertEqual(posts[0][2]["payload"]["report"]["quarter"], "Q2 2026")

    def test_url_validation_and_fingerprint(self):
        for url in ("http://public.example", "file:///tmp/data", "https://bridge.example?secret=x", "https://user:secret@bridge.example"):
            with self.assertRaises(BridgeError):
                BridgeClient(url, "secret")
        BridgeClient("http://127.0.0.1:8765", "secret")
        self.assertEqual(content_fingerprint({"report": "a", "generated_at": "one"}),
                         content_fingerprint({"report": "a", "generated_at": "two"}))

    def test_invalid_json_is_a_nonretryable_error(self):
        client = BridgeClient("https://bridge.example", "secret")
        with self.assertRaises(BridgeError) as error:
            client.request("POST", "/v1/connections/destination/jobs", {"value": float("nan")})
        self.assertFalse(error.exception.retryable)

    def test_uncertain_send_retries_the_same_snapshot_and_request_id(self):
        attempts = []
        def request(client, method, path, data=None):
            if method == "POST":
                attempts.append(data)
                if len(attempts) == 1:
                    raise BridgeError("Connection interrupted")
                return {"status": "queued"}
            return {"connections": [{**CONNECTION, "latest_job": None}]}
        with patch.dict(os.environ, CONFIG), patch.object(BridgeClient, "request", request):
            app = AppTest.from_string(SCRIPT).run()
            unlock(app)
            app.button(key="figma_send").click().run()
            self.assertIn("Connection interrupted", app.error[0].value)
            app.button(key="figma_send").click().run()
            self.assertFalse(app.exception)
        self.assertEqual(len(attempts), 2)
        self.assertEqual(attempts[0], attempts[1])

    def test_unlock_persists_across_reruns_without_download_or_code_reentry(self):
        connection = {**CONNECTION, "latest_job": None}
        with patch.dict(os.environ, CONFIG), patch.object(BridgeClient, "request", return_value={"connections": [connection]}):
            app = AppTest.from_string(SCRIPT).run()
            unlock(app)
            app.run()
            self.assertFalse(app.exception)
            self.assertFalse(any(item.key == "figma_access_code" for item in app.text_input))
            self.assertFalse(app.button(key="figma_send").disabled)
            app.button(key="figma_lock").click().run()
            self.assertTrue(app.button(key="figma_send").disabled)
            self.assertTrue(any(item.key == "figma_access_code" for item in app.text_input))

    def test_changed_access_settings_require_unlock_again(self):
        with patch.dict(os.environ, CONFIG), patch.object(BridgeClient, "request", return_value={"connections": [{**CONNECTION, "latest_job": None}]}) as request:
            app = AppTest.from_string(SCRIPT).run()
            unlock(app)
            request.reset_mock()
            with patch.dict(os.environ, {"FIGMA_BRIDGE_ACCESS_CODE": "a-new-team-access-code"}):
                app.run()
            self.assertTrue(app.button(key="figma_send").disabled)
            request.assert_not_called()

    def test_unpaired_page_shows_setup_and_disabled_send(self):
        with patch.dict(os.environ, CONFIG), patch.object(BridgeClient, "request", return_value={"connections": []}):
            app = AppTest.from_string(SCRIPT).run()
            unlock(app)
            self.assertFalse(app.exception)
            self.assertTrue(app.button(key="figma_send").disabled)
            self.assertTrue(any(item.label == "Create pairing code" for item in app.button))

    def test_previous_transfer_is_not_confirmation_for_current_report(self):
        previous = {**CONNECTION, "latest_job": {"id": "old-job", "quarter": "Q2 2026", "status": "succeeded", "result": {"page_count": 25}}}
        with patch.dict(os.environ, CONFIG), patch.object(BridgeClient, "request", return_value={"connections": [previous]}):
            app = AppTest.from_string(SCRIPT).run()
            unlock(app)
            self.assertFalse(app.success)
            self.assertTrue(any("This report has not been sent yet" in item.value for item in app.caption))


if __name__ == "__main__":
    unittest.main()
