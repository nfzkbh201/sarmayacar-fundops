import { timingSafeEqual } from 'node:crypto';

const SCHEMA = 'sarmayacar.quarterly_report.v1';
const MAX_BODY = 5 * 1024 * 1024;
const RETENTION = 86400;
const LEASE = 120;
const encoder = new TextEncoder();
const headers = {
  'Cache-Control': 'no-store',
  'X-Content-Type-Options': 'nosniff',
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
  'Access-Control-Allow-Headers': 'Authorization, Content-Type',
  'Access-Control-Max-Age': '3600',
};

class HttpError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

const nowSeconds = () => Date.now() / 1000;
const json = (value, status = 200) => Response.json(value, { status, headers });
const statement = (db, sql, ...args) => db.prepare(sql).bind(...args);
const liveConnection = 'EXISTS (SELECT 1 FROM connections c WHERE c.id = jobs.connection_id AND c.revoked = 0)';

function randomToken(bytes) {
  return Buffer.from(crypto.getRandomValues(new Uint8Array(bytes))).toString('base64url');
}

async function digest(value) {
  return Buffer.from(await crypto.subtle.digest('SHA-256', encoder.encode(value))).toString('hex');
}

function equalHashes(left, right) {
  return typeof left === 'string' && left.length === right.length &&
    timingSafeEqual(encoder.encode(left), encoder.encode(right));
}

function string(data, key, limit = 200) {
  const value = data[key];
  if (typeof value !== 'string' || !value.trim() || value.length > limit) {
    throw new HttpError(400, `Invalid ${key}.`);
  }
  return value.trim();
}

async function body(request) {
  if (Number(request.headers.get('Content-Length')) > MAX_BODY) {
    throw new HttpError(413, 'Report exceeds the 5 MB transfer limit.');
  }
  const reader = request.body?.getReader();
  const chunks = [];
  let size = 0;
  if (reader) {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      size += value.byteLength;
      if (size > MAX_BODY) {
        await reader.cancel();
        throw new HttpError(413, 'Report exceeds the 5 MB transfer limit.');
      }
      chunks.push(value);
    }
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
  let value;
  try {
    value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch {
    throw new HttpError(400, 'Invalid JSON.');
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new HttpError(400, 'Expected a JSON object.');
  }
  return value;
}

function canonical(value, depth = 0) {
  if (depth > 64 || (typeof value === 'number' && !Number.isFinite(value))) {
    throw new HttpError(400, 'Report values must be finite JSON with at most 64 nesting levels.');
  }
  if (Array.isArray(value)) return value.map(item => canonical(item, depth + 1));
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.keys(value).sort().map(key => [key, canonical(value[key], depth + 1)]));
  }
  return value;
}

function reportPayload(payload) {
  if (!payload || payload.schema_version !== SCHEMA) throw new HttpError(400, 'Unsupported report schema.');
  const quarter = payload.report?.quarter;
  if (typeof quarter !== 'string' || !/^Q[1-4] 20\d{2}$/.test(quarter)) {
    throw new HttpError(400, 'Report must specify its quarter and year.');
  }
  const months = payload.report.months;
  if (!Array.isArray(months) || !months.length || !months.every(month => typeof month === 'string')) {
    throw new HttpError(400, 'Report months are missing.');
  }
  if (!['companies', 'operator_angels', 'financial_exhibits', 'portfolio_summary'].some(key => {
    const section = payload[key];
    return section && typeof section === 'object' && Object.keys(section).length > 0;
  })) throw new HttpError(400, 'Report contains no pages.');
  const encoded = JSON.stringify(canonical(payload));
  if (encoder.encode(encoded).length > MAX_BODY) throw new HttpError(413, 'Report exceeds the 5 MB transfer limit.');
  return { quarter, encoded };
}

function splitPayload(encoded) {
  const chunks = [];
  for (let start = 0; start < encoded.length;) {
    let end = Math.min(start + 256000, encoded.length);
    // Do not split a surrogate pair when D1 stores each chunk as UTF-8 text.
    const last = encoded.charCodeAt(end - 1);
    if (end < encoded.length && last >= 0xd800 && last <= 0xdbff) end--;
    chunks.push(encoded.slice(start, end));
    start = end;
  }
  return chunks;
}

function jobInfo(row) {
  if (!row) return null;
  return {
    id: row.id, connection_id: row.connection_id, quarter: row.quarter,
    status: row.status, created: row.created, updated: row.updated,
    result: row.result ? JSON.parse(row.result) : null,
  };
}

function expire(db, now, connectionId) {
  return statement(db, `UPDATE jobs SET status = 'expired', updated = ?
    WHERE connection_id = ? AND status IN ('queued', 'processing') AND created <= ?`,
  now, connectionId, now - RETENTION);
}

async function authenticate(request, env, publisher) {
  const supplied = request.headers.get('Authorization') || '';
  if (!supplied.startsWith('Bearer ') || supplied.length > 512) throw new HttpError(401, 'Invalid credentials.');
  const token = supplied.slice(7);
  if (publisher) {
    if (!equalHashes(await digest(token), await digest(env.FIGMA_BRIDGE_PUBLISHER_TOKEN))) {
      throw new HttpError(401, 'Invalid publisher credentials.');
    }
    return null;
  }
  if (!/^[\w-]{24}\.[\w-]{43}$/.test(token)) throw new HttpError(401, 'Connection expired or revoked. Pair again in Streamlit.');
  const [id, secret] = token.split('.');
  const row = await statement(env.DB, 'SELECT * FROM connections WHERE id = ? AND revoked = 0', id).first();
  if (!row || !equalHashes(row.token_hash, await digest(secret))) {
    throw new HttpError(401, 'Connection expired or revoked. Pair again in Streamlit.');
  }
  return row;
}

async function publish(request, db, connectionId) {
  const data = await body(request);
  const requestId = string(data, 'request_id', 100);
  const { quarter, encoded } = reportPayload(data.payload);
  const hash = await digest(encoded);
  const now = nowSeconds(), id = randomToken(18);
  // The batch is one transaction. Conditional inserts and a partial unique index
  // prevent competing publishers from queuing two reports or mixing their chunks.
  const results = await db.batch([
    expire(db, now, connectionId),
    statement(db, `INSERT INTO jobs(id, connection_id, request_id, quarter, payload_hash, status, created, updated)
      SELECT ?, ?, ?, ?, ?, 'queued', ?, ?
      WHERE EXISTS (SELECT 1 FROM connections WHERE id = ? AND revoked = 0 AND target IS NOT NULL)
        AND NOT EXISTS (SELECT 1 FROM jobs WHERE connection_id = ? AND status IN ('queued', 'processing'))
      ON CONFLICT(connection_id, request_id) DO NOTHING`,
    id, connectionId, requestId, quarter, hash, now, now, connectionId, connectionId),
    ...splitPayload(encoded).map((content, part) => statement(db,
      `INSERT INTO payload_chunks(job_id, part, content) SELECT ?, ?, ? WHERE EXISTS (SELECT 1 FROM jobs WHERE id = ?)`,
      id, part, content, id)),
    statement(db, 'SELECT * FROM jobs WHERE connection_id = ? AND request_id = ?', connectionId, requestId),
    statement(db, 'SELECT id FROM connections WHERE id = ? AND revoked = 0 AND target IS NOT NULL', connectionId),
  ]);
  if (!results.at(-1).results.length) throw new HttpError(409, 'Pair the Figma plugin before sending a report.');
  const job = results.at(-2).results[0];
  if (!job) throw new HttpError(409, 'A report is already waiting or importing. Wait for it to finish.');
  if (job.payload_hash !== hash) throw new HttpError(409, 'This request was already used for a different report.');
  return jobInfo(job);
}

async function route(request, env) {
  const path = new URL(request.url).pathname;
  const method = request.method;
  if (method === 'GET' && path === '/health') {
    await env.DB.prepare('SELECT id FROM connections LIMIT 1').all();
    return { status: 'ok' };
  }
  const publisherRoute = path === '/v1/connections' || /^\/v1\/connections\/[\w-]+(?:\/jobs)?$/.test(path);
  const pluginRoute = /^\/v1\/plugin\/(connect|jobs(?:\/[\w-]+\/(claim|result))?)$/.test(path);
  if (!publisherRoute && !pluginRoute) throw new HttpError(404, 'Not found.');
  const connection = await authenticate(request, env, publisherRoute);
  const db = env.DB;
  const now = nowSeconds();

  if (path === '/v1/connections') {
    if (method === 'POST') {
      const name = string(await body(request), 'name', 100);
      const id = randomToken(18), secret = randomToken(32);
      await statement(db, 'INSERT INTO connections(id, name, token_hash) VALUES (?, ?, ?)', id, name, await digest(secret)).run();
      return { id, name, pairing_code: `${id}.${secret}` };
    }
    if (method === 'GET') {
      const results = await db.batch([
        statement(db, `UPDATE jobs SET status = 'expired', updated = ? WHERE status IN ('queued', 'processing') AND created <= ?`, now, now - RETENTION),
        db.prepare(`SELECT c.id, c.name, c.target, c.last_seen,
          (SELECT json_object('id', j.id, 'connection_id', j.connection_id, 'quarter', j.quarter,
            'status', j.status, 'created', j.created, 'updated', j.updated, 'result', j.result)
           FROM jobs j WHERE j.connection_id = c.id ORDER BY j.created DESC, j.rowid DESC LIMIT 1) AS latest
          FROM connections c WHERE c.revoked = 0 ORDER BY c.name`),
      ]);
      return { connections: results[1].results.map(row => ({
        id: row.id, name: row.name, target: row.target ? JSON.parse(row.target) : null,
        last_seen: row.last_seen, online: Boolean(row.last_seen && row.last_seen > now - 45),
        latest_job: row.latest ? jobInfo(JSON.parse(row.latest)) : null,
      })) };
    }
  }

  let match = path.match(/^\/v1\/connections\/([\w-]+)(\/jobs)?$/);
  if (match) {
    if (match[2] && method === 'POST') return publish(request, db, match[1]);
    if (!match[2] && method === 'DELETE') {
      await db.batch([
        statement(db, 'UPDATE connections SET revoked = 1 WHERE id = ?', match[1]),
        statement(db, `UPDATE jobs SET status = 'cancelled', updated = ? WHERE connection_id = ? AND status IN ('queued', 'processing')`, now, match[1]),
      ]);
      return { status: 'revoked' };
    }
  }

  if (path === '/v1/plugin/connect' && method === 'POST') {
    const data = await body(request);
    const target = Object.fromEntries(['target_id', 'page_id', 'page_name', 'file_name'].map(key => [key, string(data, key)]));
    const row = await statement(db, `UPDATE connections SET target = ?, last_seen = ? WHERE id = ? AND revoked = 0
      AND (target IS NULL OR (json_extract(target, '$.target_id') = ? AND json_extract(target, '$.page_id') = ?))
      RETURNING name`, JSON.stringify(target), now, connection.id, target.target_id, target.page_id).first();
    if (!row) throw new HttpError(409, 'This connection belongs to another Figma page or was revoked. Create a new connection.');
    return { name: row.name, target };
  }

  if (path === '/v1/plugin/jobs' && method === 'GET') {
    const results = await db.batch([
      expire(db, now, connection.id),
      statement(db, `UPDATE connections SET last_seen = ? WHERE id = ? AND revoked = 0 AND (last_seen IS NULL OR last_seen < ?)`, now, connection.id, now - 20),
      statement(db, `SELECT * FROM jobs WHERE connection_id = ? AND status IN ('queued', 'processing') AND ${liveConnection} LIMIT 1`, connection.id),
    ]);
    return { job: jobInfo(results[2].results[0]) };
  }

  match = path.match(/^\/v1\/plugin\/jobs\/([\w-]+)\/(claim|result)$/);
  if (match && method === 'POST') {
    const id = match[1], data = await body(request), runner = string(data, 'runner', 100);
    if (match[2] === 'claim') {
      const results = await db.batch([
        expire(db, now, connection.id),
        statement(db, `UPDATE jobs SET status = 'processing', runner = ?, lease_until = ?, updated = ?
          WHERE id = ? AND connection_id = ? AND status IN ('queued', 'processing') AND ${liveConnection}
          AND (runner IS NULL OR runner = ? OR lease_until <= ?) RETURNING id`,
        runner, now + LEASE, now, id, connection.id, runner, now),
        statement(db, `SELECT content FROM payload_chunks WHERE job_id = ? AND EXISTS
          (SELECT 1 FROM jobs WHERE id = ? AND connection_id = ? AND runner = ? AND status = 'processing'
            AND updated = ? AND ${liveConnection}) ORDER BY part`, id, id, connection.id, runner, now),
        statement(db, 'UPDATE connections SET last_seen = ? WHERE id = ? AND revoked = 0', now, connection.id),
      ]);
      if (!results[1].results.length) throw new HttpError(409, 'Report is no longer waiting or another plugin session is importing it.');
      if (!results[2].results.length) throw new Error('Missing report payload');
      return { id, payload: JSON.parse(results[2].results.map(row => row.content).join('')) };
    }
    const status = data.status, count = data.page_count ?? 0;
    if (!['succeeded', 'failed'].includes(status)) throw new HttpError(400, 'Invalid result status.');
    if (!Number.isInteger(count) || count < 0 || count > 1000) throw new HttpError(400, 'Invalid page count.');
    const result = { page_count: count, message: String(data.message ?? '').slice(0, 1000),
      warning_count: Array.isArray(data.warnings) ? data.warnings.length : 0 };
    const results = await db.batch([
      expire(db, now, connection.id),
      statement(db, `UPDATE jobs SET status = ?, result = ?, updated = ? WHERE id = ? AND connection_id = ?
        AND runner = ? AND status = 'processing' AND ${liveConnection}`,
      status, JSON.stringify(result), now, id, connection.id, runner),
      statement(db, `SELECT * FROM jobs WHERE id = ? AND connection_id = ? AND ${liveConnection}`, id, connection.id),
    ]);
    const row = results[2].results[0];
    if (!row || row.runner !== runner || row.status !== status) throw new HttpError(409, 'Import session no longer owns this report.');
    return jobInfo(row);
  }
  throw new HttpError(405, 'Method not allowed.');
}

export default {
  async fetch(request, env) {
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers });
    if (typeof env.FIGMA_BRIDGE_PUBLISHER_TOKEN !== 'string' || env.FIGMA_BRIDGE_PUBLISHER_TOKEN.length < 32) {
      return json({ detail: 'Connection service is not configured.' }, 503);
    }
    try {
      return json(await route(request, env));
    } catch (error) {
      if (error instanceof HttpError) return json({ detail: error.message }, error.status);
      // Database errors can contain SQL parameters. Never return or log them.
      return json({ detail: 'Connection service temporarily unavailable. Please retry.' }, 503);
    }
  },
  async scheduled(_event, env) {
    const now = nowSeconds();
    await env.DB.batch([
      statement(env.DB, `UPDATE jobs SET status = 'expired', updated = ? WHERE status IN ('queued', 'processing') AND created <= ?`, now, now - RETENTION),
      statement(env.DB, `DELETE FROM jobs WHERE status NOT IN ('queued', 'processing') AND created < ?`, now - 30 * RETENTION),
    ]);
  },
};
