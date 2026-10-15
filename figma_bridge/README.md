# Streamlit to Figma Connection

Streamlit sends the reviewed quarterly data pack to a private service. The paired
Figma plugin receives it while open, updates generated report frames, and reports
success or failure back to Streamlit. No JSON download or paste is needed after
pairing. Closing Figma pauses delivery; this is not an unattended Figma renderer.

## Local Test

From the Monthly_Reporting directory:

```sh
.venv/bin/python -m pip install -r figma_bridge/requirements.txt
.venv/bin/python tools/run_figma_connection.py --pack "/path/to/report.json"
```

The launcher prints the local Streamlit screen and service URLs. It binds both
to `127.0.0.1`, creates private credentials under `.figma_bridge/`, and does not
publish any report automatically. That directory is excluded from Git.

1. Open the printed Streamlit URL and create a pairing code.
2. Open the report file and intended page in Figma Desktop. Run the updated local
   plugin from `figma_quarterly_report_plugin/manifest.json`.
3. Enter the local service URL and pairing code in the plugin, then Connect.
   In the plugin URL use `localhost` instead of `127.0.0.1`: Figma rejects numeric
   loopback addresses in `devAllowedDomains`. Ports 8765 and 8766 are allowed;
   add the selected localhost port to the manifest if the launcher uses another.
4. Refresh destinations in Streamlit and click Send to Figma.
5. Check that Streamlit reports completion. Send the same quarter again and
   confirm that its generated frame count does not increase.

The preview uses an existing pack to test transport; the real Quarterly Report
tab sends its current reviewed data directly. A successful transfer does not
validate the underlying financial mappings or report layout.

Unlock the handoff with the team access code once per Streamlit session using
`Unlock Figma`. `Send to Figma` is the primary action; the optional JSON download
is below it in `Download data pack` and does not trigger a report rebuild.
Portfolio totals are shown before sending. A previous transfer is labeled as
previous until the current report's own job is confirmed.

## Hosted Setup

The Cloudflare Workers + D1 implementation and deployment instructions are in
[cloudflare/README.md](cloudflare/README.md). It preserves this API and is the
selected hosting path. The Worker is deployed on the Free plan and its synthetic
transfer checks pass. Live dashboard-to-Figma acceptance remains a separate check;
the local Mac connection is not reachable from Streamlit Cloud.

For a separately managed Python host instead, deploy with HTTPS and a
persistent private disk. A Dockerfile is included:

```sh
docker build -f figma_bridge/Dockerfile -t sarmayacar-figma-bridge .
```

Set `FIGMA_BRIDGE_PUBLISHER_TOKEN` to a cryptographically random secret with at
least 32 characters in both the service environment and Streamlit secrets.
Mount a persistent volume at `/data` (writable by uid 10001). The service listens
on port 8765; the host should terminate HTTPS in front of it. Use one service
instance with this SQLite database, not independent replicas.

Configure these three values in Streamlit secrets:

```toml
FIGMA_BRIDGE_URL = "https://your-private-service.example"
FIGMA_BRIDGE_PUBLISHER_TOKEN = "REPLACE_WITH_RANDOM_SERVER_SECRET"
FIGMA_BRIDGE_ACCESS_CODE = "REPLACE_WITH_TEAM_ACCESS_CODE_AT_LEAST_12_CHARACTERS"
```

The access code gates the handoff controls even if the dashboard itself is
public. Do not share the publisher token with the plugin. The plugin gets a
separate connection token that can only receive reports and acknowledge imports
for one destination. Pair each destination once; credentials are remembered
locally in Figma client storage, not embedded in report layers or plugin source.

The plugin manifest restricts `networkAccess.allowedDomains` to the deployed
Cloudflare service origin, keeping `localhost` in `devAllowedDomains` only for
local testing. Update the manifest if the service moves to another origin.

## Updates and Recovery

- A connection is bound to the selected Figma page. Switching pages pauses it.
- Reports for the same quarter replace only frames tagged by this importer.
  Manual edits inside those generated frames are replaced. Unlinked old imports,
  other quarters, and other user-created content are left alone.
- The importer renders into temporary staging content before replacing the old
  report. A rendering failure removes the staged content and keeps the old report.
- Duplicate delivery of the same job is acknowledged without rebuilding it.
- Publishing is idempotent across a network retry. Only one unfinished report
  may be queued per connection. A crashed receiver's lease expires after two minutes.
- Report payloads are cleared after success or failure. Undelivered payloads
  expire after 24 hours and are removed on the next service request. Status and
  destination metadata remain so the app can show the outcome after a restart.
- Revoking a connection in Streamlit stops new delivery and clears queued data.
  A render already running in Figma may finish; revocation cannot undo it.
- JSON file import remains available as a fallback. Older unlinked frames are
  not automatically adopted or deleted, since their ownership cannot be verified.

## Verification

```sh
.venv/bin/python -m pip install -r figma_bridge/requirements-test.txt
.venv/bin/python -m unittest discover -s tests -p 'test_figma*.py' -v
node --test tests/figma_import.test.cjs
```

The automated checks cover authentication, destination isolation, delivery,
expiry, restart recovery, acknowledgement retries and report replacement. Node
tests use a Figma API test double; they do not replace a live Figma import test.
