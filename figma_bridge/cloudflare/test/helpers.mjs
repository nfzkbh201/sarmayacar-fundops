import { Miniflare, convertV4MiniflareOptions } from 'miniflare';
import { randomBytes } from 'node:crypto';
import { readFile, mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const root = fileURLToPath(new URL('../', import.meta.url));
export const sample = () => ({
  schema_version: 'sarmayacar.quarterly_report.v1',
  report: { quarter: 'Q2 2026', months: ['Apr 26', 'May 26', 'Jun 26'] },
  companies: [{ name: 'Test company', description: 'Synthetic test data only' }],
});
export const target = { target_id: 'document-one', page_id: '1:2', page_name: 'Test page', file_name: 'Test file' };

export async function startBridge({ token = randomBytes(32).toString('hex'), persist, migrate = true } = {}) {
  const directory = persist || await mkdtemp(join(tmpdir(), 'figma-cloudflare-'));
  const mf = new Miniflare(convertV4MiniflareOptions({
    modules: true, scriptPath: join(root, 'src/worker.js'),
    compatibilityDate: '2026-10-13', compatibilityFlags: ['nodejs_compat'],
    bindings: { FIGMA_BRIDGE_PUBLISHER_TOKEN: token },
    d1Databases: ['DB'], resourcePersistencePath: directory, cf: false,
    telemetry: { enabled: false },
    host: '127.0.0.1', port: 0,
  }));
  try {
    const url = String(await mf.ready).replace(/\/$/, '');
    const db = await mf.getD1Database('DB');
    if (migrate) await db.exec(await readFile(join(root, 'migrations/0001_bridge.sql'), 'utf8'));
    return {
      mf, db, token, url, directory,
      request: async (path, { method = 'GET', credential = token, data, raw } = {}) => {
        const response = await mf.dispatchFetch(url + path, {
          method, headers: { Authorization: 'Bearer ' + credential, 'Content-Type': 'application/json' },
          body: raw ?? (data === undefined ? undefined : JSON.stringify(data)),
        });
        return { status: response.status, headers: response.headers, data: await response.json() };
      },
      close: async () => { await mf.dispose(); if (!persist) await rm(directory, { recursive: true, force: true }); },
    };
  } catch (error) {
    await mf.dispose();
    if (!persist) await rm(directory, { recursive: true, force: true });
    throw error;
  }
}
