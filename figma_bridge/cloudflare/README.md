# Cloudflare Figma Connection

Cloudflare Workers hosts the authenticated bridge API; D1 stores destinations,
short-lived transfer payloads and status. Streamlit still generates the report,
and the open Figma plugin still renders it. Neither application moves hosts.
This implements the same API as `figma_bridge/service.py`; the local Python
service remains available and does not need Cloudflare credentials.

## Deployment Status

The service was deployed on Workers Free on October 13, 2026:

https://sarmayacar-figma-bridge.sarmayacar-figma-cloudflare-bridge.workers.dev

Hosted synthetic checks passed for authentication, 1 KB and 200 KB reports,
publish retries, claim/result retries, and revocation. Run them from the project
root with `.venv/bin/python tools/smoke_figma_cloudflare.py`; credentials are read
from the ignored `.figma_bridge/cloudflare.json`, never printed. The dashboard
client identifies itself with `User-Agent: Sarmayacar-FundOps/1.0` because the
host rejected Python's default request identifier before reaching the Worker.
These checks do not claim that the hosted route has been tested inside Figma.

## Local Verification

Requires Node 22 or newer and the repository's Python environment for the
dashboard-client integration check. From this directory:

```sh
npm ci
npm test
npm run test:plugin
npm run check:deploy
```

Tests start Cloudflare's workerd runtime and temporary D1 databases locally.
They use synthetic reports only. The plugin tests execute the actual plugin
code against a Figma API test double, not the Figma desktop application.
The deployment dry-run bundles code without publishing it.

For an interactive local Worker, create an ignored `.dev.vars` file with a
random `FIGMA_BRIDGE_PUBLISHER_TOKEN` of at least 32 characters, then:

```sh
npx wrangler d1 migrations apply sarmayacar-figma-bridge --local
npm run dev
```

The local port is 8766, leaving the existing Python connection on 8765 alone.

## First Deployment

No Cloudflare account or infrastructure is created by installing these files.
Use the Workers Free plan. Do not enable Workers Paid, R2, or paid add-ons.
After the account owner signs in, from this directory:

```sh
npx wrangler login --scopes account:read user:read workers_scripts:write d1:write
npx wrangler whoami
npx wrangler d1 create sarmayacar-figma-bridge
npm run configure -- DATABASE_UUID_FROM_PREVIOUS_COMMAND
npx wrangler d1 migrations apply sarmayacar-figma-bridge --remote --config wrangler.production.json
npx wrangler secret put FIGMA_BRIDGE_PUBLISHER_TOKEN --config wrangler.production.json
npm run deploy
```

Create a fresh cryptographically random publisher secret of at least 32
characters in a password manager. Enter it at Wrangler's secret prompt, not as
a command-line argument, chat message, or committed file. Use the same secret
in Streamlit. The configure command creates an ignored production config and
refuses to overwrite it. The checked-in all-zero database ID is local-only;
the deploy script requires the separate production config.

Record the exact HTTPS workers.dev URL returned by deployment. In the existing
Streamlit app's private secrets, set:

```toml
FIGMA_BRIDGE_URL = "https://sarmayacar-figma-bridge.YOUR_SUBDOMAIN.workers.dev"
FIGMA_BRIDGE_PUBLISHER_TOKEN = "THE_SAME_PRIVATE_RANDOM_SECRET"
FIGMA_BRIDGE_ACCESS_CODE = "A_SEPARATE_RANDOM_TEAM_ACCESS_CODE_AT_LEAST_12_CHARACTERS"
```

These are placeholders, not usable credentials. The publisher secret is
server-only; the plugin receives its own limited pairing code. Local pairings
do not transfer to Cloudflare: create a new pairing in the hosted dashboard.
Once the actual URL is known, restrict the plugin manifest's
`networkAccess.allowedDomains` to that HTTPS origin and keep loopback URLs in
`devAllowedDomains`. Keep the production Streamlit entry point `app.py` and all
four dashboard tabs unchanged.

## Hosted Acceptance Check

1. Confirm `/health` returns `{"status":"ok"}` and unauthenticated connection
   requests return 401. Confirm the account is still on Workers Free.
2. Configure the hosted dashboard, create a pairing code, and connect the
   plugin on the intended Figma page using the deployed URL.
3. Send a synthetic report first. Check completion in both applications.
4. Send the same quarter again. Verify replacement without duplicate frames,
   and confirm unlinked frames and other quarters remain untouched.
5. Close and reopen the plugin; check reconnection. Revoke the connection and
   confirm that its old pairing code no longer works.
6. Test the current real report, inspect the actual Figma output and check
   Cloudflare's CPU/request and D1 usage metrics. Local tests do not enforce
   the Free plan's CPU allowance and do not prove hosted performance.

Do not call the hosted integration complete before these checks. Report
mapping and formatting sign-off are separate from successful delivery.

## Privacy and Recovery

- TLS protects the hosted connection. Bearer authentication is required on all
  data routes. CORS supports Figma's opaque origin; it grants no data access
  without the token. Responses use `Cache-Control: no-store`.
- Pairing secrets are stored only as SHA-256 hashes in D1. Each pairing is tied
  to one Figma document/page. Only the server publisher can queue reports.
- D1 transactions and a unique active-job index prevent concurrent transfers
  from mixing. Request IDs and renderer job IDs make retries idempotent.
- Payload chunks are below D1's 2 MB row limit. The API accepts at most 5 MB
  per request; this is a size ceiling, not a hosted Free CPU guarantee.
- Completed, failed, revoked, or expired payloads are deleted from the active
  database. Waiting transfers expire after 24 hours, checked on access and by
  hourly cleanup. Final job metadata is deleted after 30 days. Destination
  metadata and token hashes remain for reconnection/revocation.
- D1 Time Travel can retain historical data for seven days on Free. Clearing
  active payloads is not immediate erasure from Cloudflare's recovery history.
- Worker request logging/observability is disabled. Error responses do not
  contain SQL, secrets, or reports. Cloudflare still processes/stores data;
  this is not end-to-end encryption or a zero-retention service.
- Revocation cannot undo a report already downloaded/rendering inside Figma.
  JSON file import remains a fallback during outages or quota exhaustion.

The SQL migration uses one complete statement per line to work with D1's
`exec()` as well as Wrangler's migration runner, including the cleanup trigger.

## Free-Tier Budget

As checked on October 13, 2026: Workers Free allows 100,000 requests/day with
10 ms CPU per request. D1 Free includes 5 million rows read/day, 100,000 rows
written/day, 500 MB per database and 5 GB account storage. Other applications
share account limits. There is no paid overflow on Free; exhausted quotas can
interrupt delivery. Confirm current limits before deployment:

- https://developers.cloudflare.com/workers/platform/limits/
- https://developers.cloudflare.com/d1/platform/pricing/
- https://developers.cloudflare.com/d1/platform/limits/
- https://developers.cloudflare.com/d1/reference/time-travel/

The existing five-second polls from one plugin and one Streamlit status view
produce roughly 1,440 routine requests/hour while both stay open, plus sends,
claims, acknowledgements and browser preflights. The service throttles idle
heartbeat writes to about once per 20 seconds. Close idle plugin/status views
and monitor actual usage. No custom domain is required for workers.dev.
