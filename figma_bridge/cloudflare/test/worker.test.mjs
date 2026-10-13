import { test, before, after, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { startBridge, sample, target } from './helpers.mjs';

let bridge;
before(async () => { bridge = await startBridge(); });
after(async () => { await bridge?.close(); });
beforeEach(async () => {
  await bridge.db.batch(['DELETE FROM payload_chunks', 'DELETE FROM jobs', 'DELETE FROM connections'].map(sql => bridge.db.prepare(sql)));
});
const call = async (path, data, credential = bridge.token, method = data === undefined ? 'GET' : 'POST') => {
  const response = await bridge.request(path, { method, data, credential });
  assert.equal(response.status, 200, JSON.stringify(response.data));
  return response.data;
};
async function pair(name = 'Test') {
  const c = await call('/v1/connections', { name });
  await call('/v1/plugin/connect', target, c.pairing_code);
  return c;
}
const publish = (c, request_id = 'one', payload = sample()) => call(`/v1/connections/${c.id}/jobs`, { request_id, payload });
const claim = (c, j, runner = 'runner-one') => call(`/v1/plugin/jobs/${j.id}/claim`, { runner }, c.pairing_code);
const finish = (c, j, runner = 'runner-one', status = 'succeeded') => call(`/v1/plugin/jobs/${j.id}/result`, { runner, status, page_count: 8 }, c.pairing_code);
const chunkCount = async () => (await bridge.db.prepare('SELECT COUNT(*) AS n FROM payload_chunks').first()).n;

test('full transfer, retry after lost acknowledgement, replacement and payload cleanup', async () => {
  const c = await pair();
  const j = await publish(c);
  assert.equal((await call('/v1/plugin/jobs', undefined, c.pairing_code)).job.id, j.id);
  assert.deepEqual((await claim(c, j)).payload, sample());
  await finish(c, j);
  assert.equal(await chunkCount(), 0);
  assert.equal((await finish(c, j)).status, 'succeeded');
  assert.equal((await publish(c)).id, j.id);
  const listing = (await call('/v1/connections')).connections[0];
  assert.equal(listing.online, true);
  assert.equal(listing.latest_job.result.page_count, 8);
  assert.equal(listing.latest_job.payload, undefined);
  assert.equal(listing.token_hash, undefined);
  assert.equal(listing.pairing_code, undefined);
  assert.equal((await call('/v1/plugin/jobs', undefined, c.pairing_code)).job, null);
  assert.notEqual((await publish(c, 'two')).id, j.id);
});

test('publisher authentication and destination-scoped reader credentials', async () => {
  assert.equal((await bridge.request('/v1/connections', { credential: 'wrong' })).status, 401);
  const a = await pair('a'), b = await pair('b');
  const j = await publish(a);
  assert.equal((await call('/v1/plugin/jobs', undefined, b.pairing_code)).job, null);
  assert.equal((await bridge.request(`/v1/plugin/jobs/${j.id}/claim`, { method: 'POST', credential: b.pairing_code, data: { runner: 'other' } })).status, 409);
  assert.equal((await bridge.request('/v1/connections', { credential: a.pairing_code })).status, 401);
  assert.equal((await bridge.request('/v1/plugin/jobs', { credential: bridge.token })).status, 401);
  assert.equal((await bridge.request('/v1/plugin/connect', { method: 'POST', credential: a.pairing_code, data: { ...target, page_id: 'another' } })).status, 409);
});

test('unpaired and revoked connections cannot receive reports', async () => {
  const c = await call('/v1/connections', { name: 'Unpaired' });
  assert.equal((await bridge.request(`/v1/connections/${c.id}/jobs`, { method: 'POST', data: { request_id: 'one', payload: sample() } })).status, 409);
  await call('/v1/plugin/connect', target, c.pairing_code);
  const j = await publish(c);
  await claim(c, j);
  await call(`/v1/connections/${c.id}`, undefined, bridge.token, 'DELETE');
  assert.equal(await chunkCount(), 0);
  assert.equal((await bridge.request('/v1/plugin/jobs', { credential: c.pairing_code })).status, 401);
  assert.equal((await bridge.request(`/v1/plugin/jobs/${j.id}/result`, { method: 'POST', credential: c.pairing_code, data: { runner: 'runner-one', status: 'succeeded' } })).status, 401);
});

test('concurrent publication queues one report and never interleaves payloads', async () => {
  const c = await pair();
  const results = await Promise.all(['a', 'b'].map(request_id => bridge.request(`/v1/connections/${c.id}/jobs`, {
    method: 'POST', data: { request_id, payload: { ...sample(), marker: request_id } },
  })));
  assert.deepEqual(results.map(r => r.status).sort(), [200, 409]);
  const accepted = results.find(r => r.status === 200).data;
  assert.equal((await claim(c, accepted)).payload.marker, results[0].status === 200 ? 'a' : 'b');
  assert.equal(await chunkCount(), 1);
});

test('request id is canonical, idempotent, and cannot be reused for different data', async () => {
  const c = await pair(), payload = sample();
  const j = await publish(c, 'same', payload);
  const reordered = { companies: payload.companies, report: payload.report, schema_version: payload.schema_version };
  assert.equal((await publish(c, 'same', reordered)).id, j.id);
  assert.equal(await chunkCount(), 1);
  assert.equal((await bridge.request(`/v1/connections/${c.id}/jobs`, { method: 'POST', data: { request_id: 'same', payload: { ...payload, changed: true } } })).status, 409);
});

test('simultaneous retries of the same request resolve to one job', async () => {
  const c = await pair();
  const jobs = await Promise.all([publish(c, 'same'), publish(c, 'same'), publish(c, 'same')]);
  assert.equal(new Set(jobs.map(j => j.id)).size, 1);
  assert.equal(await chunkCount(), 1);
});

test('only one importing runner owns a lease; expired lease can be reclaimed', async () => {
  const c = await pair(), j = await publish(c);
  const responses = await Promise.all(['a', 'b'].map(runner => bridge.request(`/v1/plugin/jobs/${j.id}/claim`, { method: 'POST', credential: c.pairing_code, data: { runner } })));
  assert.deepEqual(responses.map(r => r.status).sort(), [200, 409]);
  const winner = responses[0].status === 200 ? 'a' : 'b';
  await claim(c, j, winner);
  await bridge.db.prepare('UPDATE jobs SET lease_until = 0 WHERE id = ?').bind(j.id).run();
  await claim(c, j, 'new-runner');
  assert.equal((await bridge.request(`/v1/plugin/jobs/${j.id}/result`, { method: 'POST', credential: c.pairing_code, data: { runner: winner, status: 'succeeded' } })).status, 409);
  await finish(c, j, 'new-runner', 'failed');
  assert.equal(await chunkCount(), 0);
});

test('expired transfers are inaccessible and cleared before polling or acknowledgement', async () => {
  const c = await pair(), j = await publish(c);
  await claim(c, j);
  await bridge.db.prepare('UPDATE jobs SET created = 0 WHERE id = ?').bind(j.id).run();
  assert.equal((await bridge.request(`/v1/plugin/jobs/${j.id}/result`, { method: 'POST', credential: c.pairing_code, data: { runner: 'runner-one', status: 'succeeded' } })).status, 409);
  assert.equal(await chunkCount(), 0);
  assert.equal((await call('/v1/plugin/jobs', undefined, c.pairing_code)).job, null);
  assert.equal((await call('/v1/connections')).connections[0].latest_job.status, 'expired');
});

test('large Unicode reports round-trip across D1 chunk boundaries', async () => {
  const c = await pair(), payload = { ...sample(), text: 'a'.repeat(255799) + '\u{1f4c4}'.repeat(600000) };
  const j = await publish(c, 'large', payload);
  assert.ok(await chunkCount() > 2);
  assert.deepEqual((await claim(c, j)).payload, payload);
  await finish(c, j);
  assert.equal(await chunkCount(), 0);
});

test('malformed JSON, nonfinite values, invalid reports and oversized bodies fail clearly', async () => {
  const c = await pair(), path = `/v1/connections/${c.id}/jobs`;
  const nonfinite = JSON.stringify({ request_id: 'x', payload: sample() }).replace('"Synthetic test data only"', '1e400');
  for (const raw of ['{', '[]', 'null', '{"request_id":"x","payload":{}}', nonfinite]) {
    assert.equal((await bridge.request(path, { method: 'POST', raw })).status, 400);
  }
  let nested = 'deep';
  for (let i = 0; i < 65; i++) nested = { nested };
  assert.equal((await bridge.request(path, { method: 'POST', data: { request_id: 'deep', payload: { ...sample(), nested } } })).status, 400);
  assert.equal((await bridge.request(path, { method: 'POST', raw: ' '.repeat(5 * 1024 * 1024 + 1) })).status, 413);
  assert.equal(await chunkCount(), 0);
});

test('preflight supports opaque Figma origins; all responses are non-cacheable', async () => {
  const response = await bridge.mf.dispatchFetch(bridge.url + '/v1/plugin/jobs', { method: 'OPTIONS' });
  assert.equal(response.status, 204);
  assert.equal(response.headers.get('Access-Control-Allow-Origin'), '*');
  assert.match(response.headers.get('Access-Control-Allow-Headers'), /Authorization/);
  for (const path of ['/health', '/missing']) {
    const result = await bridge.request(path);
    assert.equal(result.headers.get('Cache-Control'), 'no-store');
    assert.equal(result.headers.get('X-Content-Type-Options'), 'nosniff');
  }
});

test('scheduled cleanup expires undelivered reports and deletes old metadata', async () => {
  const c = await pair(), j = await publish(c);
  await bridge.db.prepare('UPDATE jobs SET created = ? WHERE id = ?').bind(Date.now() / 1000 - 90000, j.id).run();
  const worker = await bridge.mf.getWorker();
  await worker.scheduled({ cron: '17 * * * *', scheduledTime: Date.now() });
  assert.equal(await chunkCount(), 0);
  assert.equal((await bridge.db.prepare('SELECT status FROM jobs WHERE id = ?').bind(j.id).first()).status, 'expired');
  await bridge.db.prepare('UPDATE jobs SET created = 0 WHERE id = ?').bind(j.id).run();
  await worker.scheduled({ cron: '17 * * * *', scheduledTime: Date.now() });
  assert.equal(await bridge.db.prepare('SELECT id FROM jobs WHERE id = ?').bind(j.id).first(), null);
});

test('database and authentication survive runtime restart without revealing payload in listings', async () => {
  const persist = await mkdtemp(join(tmpdir(), 'figma-restart-'));
  let service;
  try {
    service = await startBridge({ persist });
    const token = service.token;
    const c = (await service.request('/v1/connections', { method: 'POST', data: { name: 'Persist' } })).data;
    await service.close();
    service = await startBridge({ persist, token, migrate: false });
    const response = await service.request('/v1/plugin/connect', { method: 'POST', credential: c.pairing_code, data: target });
    assert.equal(response.status, 200);
  } finally { await service?.close(); await rm(persist, { recursive: true, force: true }); }
});

test('missing publisher secret fails closed', async () => {
  const service = await startBridge({ token: '' });
  try { assert.equal((await service.request('/health')).status, 503); }
  finally { await service.close(); }
});

test('database failures never expose SQL or credentials to clients', async () => {
  const service = await startBridge();
  try {
    await service.db.prepare('DROP TABLE connections').run();
    const response = await service.request('/v1/connections');
    assert.equal(response.status, 503);
    assert.deepEqual(response.data, { detail: 'Connection service temporarily unavailable. Please retry.' });
  } finally { await service.close(); }
});
