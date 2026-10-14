const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const vm = require('node:vm');
const { spawn } = require('node:child_process');
const { mkdtemp, rm } = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const net = require('node:net');
const { webcrypto } = require('node:crypto');
const TEST_UI_SESSION = 'a'.repeat(32);

test('plugin manifest keeps the hosted origin and uses localhost for development', () => {
  const manifest = JSON.parse(fs.readFileSync('figma_quarterly_report_plugin/manifest.json', 'utf8'));
  assert.deepEqual(manifest.networkAccess.allowedDomains, [
    'https://sarmayacar-figma-bridge.sarmayacar-figma-cloudflare-bridge.workers.dev',
  ]);
  // Figma rejects numeric loopback hosts in devAllowedDomains.
  assert.deepEqual(manifest.networkAccess.devAllowedDomains, [
    'http://localhost:8765', 'http://localhost:8766',
  ]);
  for (const entry of [manifest.main, manifest.ui]) {
    assert.ok(fs.existsSync(path.join('figma_quarterly_report_plugin', entry)));
  }
});

function runtime() {
  let nextId = 0;
  class Node {
    constructor(type) {
      this.type = type; this.id = String(++nextId); this.name = ''; this.children = [];
      this.parent = null; this.removed = false; this.x = this.y = 0; this.width = this.height = 100;
      this.data = {}; this.visible = true;
    }
    appendChild(node) {
      if (node.parent) node.parent.children = node.parent.children.filter(n => n !== node);
      node.parent = this; this.children.push(node);
    }
    resize(w, h) { this.width = w; this.height = h; }
    set textAutoResize(value) {
      this.autoResize = value;
      if (value !== 'HEIGHT') return;
      // Deterministic wrapping model; actual typography is checked separately.
      const capacity = Math.max(1, Math.floor(this.width / (this.fontSize * 0.52)));
      let lines = 0;
      for (const paragraph of String(this.characters || '').split('\n')) {
        let used = 0;
        lines += 1;
        for (const word of paragraph.split(/\s+/)) {
          if (used && used + 1 + word.length > capacity) { lines += 1; used = 0; }
          used += (used ? 1 : 0) + word.length;
        }
      }
      this.height = lines * this.fontSize * (this.lineHeight?.value || 135) / 100;
    }
    get textAutoResize() { return this.autoResize; }
    remove() {
      if (this.parent) this.parent.children = this.parent.children.filter(n => n !== this);
      for (const child of [...this.children]) child.remove();
      this.removed = true;
    }
    findAll(predicate) { return this.children.flatMap(n => [...(predicate(n) ? [n] : []), ...n.findAll(predicate)]); }
    findOne(predicate) { return this.findAll(predicate)[0] || null; }
    getPluginData(key) { return this.data[key] || ''; }
    setPluginData(key, value) { this.data[key] = value; }
    clone() {
      const node = new Node(this.type);
      for (const [key, value] of Object.entries(this)) {
        if (!['id', 'children', 'parent', 'removed'].includes(key)) node[key] = structuredClone(value);
      }
      for (const child of this.children) node.appendChild(child.clone());
      return node;
    }
  }
  const page = new Node('PAGE'); page.name = 'Report';
  const messages = [], storage = new Map();
  const figma = {
    currentPage: page, root: { name: 'Test report' },
    createFrame: () => new Node('FRAME'), createRectangle: () => new Node('RECTANGLE'),
    createText: () => new Node('TEXT'), showUI() { assert.equal(typeof figma.ui.onmessage, 'function'); }, loadFontAsync: async () => {},
    viewport: { scrollAndZoomIntoView() {} }, ui: { postMessage: m => messages.push(m) },
    notify() {}, on() {}, closePlugin() {},
    clientStorage: { getAsync: async k => storage.get(k), setAsync: async (k,v) => storage.set(k,v), deleteAsync: async k => storage.delete(k) },
  };
  const context = vm.createContext({ figma, __html__: '', console });
  vm.runInContext(fs.readFileSync('figma_quarterly_report_plugin/code.js', 'utf8'), context);
  return { context, page, figma, Node, messages };
}

function payload(quarter = 'Q2 2026') {
  return {
    schema_version: 'sarmayacar.quarterly_report.v1', report: { quarter, months: ['Jan 26', 'Feb 26', 'Mar 26', 'Apr 26', 'May 26', 'Jun 26'] },
    companies: [{ name: 'SimPaisa', description: 'Test description', key_updates: ['Test update'],
      metric_sections: [{ title: 'Metrics', units: "(Numbers in 000's)", rows: [
        { label: 'GTV', values: [{ actual: '123', forecast: '140' }] },
      ] }], metrics: [] }],
    portfolio_summary: { rows: [{ name: 'Trukker', section: 'Portfolio Company' }, { name: 'Taraki', section: 'Operator Angel' }] },
    operator_angels: [{ name: 'Test angel', description: 'Test', key_updates: [] }],
    financial_exhibits: { sections: [{ section: 'Balance Sheet', months: ['Jun-26'], rows: [{ label: 'Total Assets', values: [{ value: '100' }] }] }] },
  };
}

async function update(rt, data, jobId) {
  rt.context.data = data; rt.context.jobId = jobId;
  return vm.runInContext('updateReport(data, jobId)', rt.context);
}

test('receives all report sections; repeated job and changed report do not duplicate pages', async () => {
  const rt = runtime();
  const unrelated = new rt.Node('FRAME'); unrelated.name = 'Personal notes'; rt.page.appendChild(unrelated);
  const first = await update(rt, payload(), 'job-1');
  assert.equal(first.createdCount, 9);
  const ids = rt.page.children.map(n => n.id);
  assert.ok(rt.page.children.some(n => n.name === '01 Cover'));
  assert.ok(rt.page.children.some(n => n.name === '04 Portfolio'));
  assert.ok(rt.page.children.some(n => n.name.includes('Financial Exhibits')));
  assert.ok(rt.page.findOne(n => n.characters === '123'));
  await update(rt, payload(), 'job-1');
  assert.deepEqual(rt.page.children.map(n => n.id), ids);
  const changed = payload(); changed.companies[0].description = 'Changed description';
  await update(rt, changed, 'job-2');
  assert.equal(rt.page.children.length, first.createdCount + 1);
  assert.equal(unrelated.removed, false);
  assert.ok(rt.page.findOne(n => n.characters === 'Changed description'));
  assert.equal(rt.page.findOne(n => n.name === 'Company Page Template'), null);
});

const SUMMARY_GROUPS = [
  ['Trukker', 'SimPaisa', 'Abhi', 'Bykea', 'Oladoc', 'TapMad', 'Revolving Games', 'OneLoad', 'Jiye Technologies', 'Roomy'],
  ['Procheck', 'Patari', 'Dot & Line', 'Dawaai', 'Jugnu', 'Buyzilla', 'Taraki', 'Savvy Technologies', 'Delsys Technologies'],
  ['Startup Early (Holistiq)', 'Scholar Den', 'House Call', 'Orko', 'L.L.M.Bots', 'Aabshar', 'Raptr Games'],
];
const SUMMARY_NAMES = ['04 Portfolio', '05 Portfolio Continued', '06 Portfolio Operator Angels'];
const summaryFrames = rt => SUMMARY_NAMES.map(name => rt.page.children.find(n => n.name === name));
const cellText = (frame, name) => frame.findOne(n => n.type === 'TEXT' && n.name === name)?.characters;

test('portfolio layout always reserves the approved three pages and 26 rows', async () => {
  const rt = runtime(), data = payload();
  delete data.portfolio_summary;
  await update(rt, data, 'empty-summary');
  const frames = summaryFrames(rt);
  frames.forEach((frame, i) => {
    assert.ok(frame);
    assert.deepEqual(frame.findAll(n => /row_\d+\.company$/.test(n.name)).map(n => n.characters), SUMMARY_GROUPS[i]);
    assert.equal(cellText(frame, 'report.page_number'), String(i + 3).padStart(2, '0'));
    assert.equal(frame.findAll(n => /^portfolio\.header_\d+\.text$/.test(n.name)).length, 9);
    assert.equal(Boolean(frame.findOne(n => n.name === 'portfolio.operator_title')), i === 1);
    assert.equal(Boolean(frame.findOne(n => n.name === 'portfolio.note')), i === 0);
    assert.ok(frame.findAll(n => /row_\d+\.cost$/.test(n.name)).every(n => n.characters === '-'));
  });
  assert.equal(cellText(frames[1], 'portfolio.company.total_cost'), '-');
  assert.equal(cellText(frames[2], 'portfolio.operator.total_cost'), '-');
});

test('shuffled rows and aliases keep their approved positions without changing any supplied field', async () => {
  const rt = runtime(), data = payload();
  const aliases = { Abhi: 'AbhiFinance', 'Jiye Technologies': 'Jiye Tech', Taraki: 'Bolandi Technologies',
    'Savvy Technologies': 'Savvy Technologies (Inactive)', 'L.L.M.Bots': 'Truckistan' };
  data.portfolio_summary = {
    rows: SUMMARY_GROUPS.flat().map((name, i) => ({
      name: aliases[name] || name, description: `Approved description ${i}`,
      first_cash_injection: `Jan-${String(i).padStart(2, '0')}`, ownership: `${i}.5%`, sv_board_seats: 'Observer Seat',
      cost: `$${i},123`, carrying_value: `$${i},456`, unrealized_gain_loss: `$(78${i})`, realized_gain_loss: '$0',
    })).reverse(),
    portfolio_totals: { cost: '$16,449,971', carrying_value: '$21,632,938', unrealized_gain_loss: '$5,182,967', realized_gain_loss: '$0' },
    totals: { cost: '$16,924,971', carrying_value: '$22,075,199', unrealized_gain_loss: '$5,150,228', realized_gain_loss: '$0' },
  };
  const original = structuredClone(data);
  await update(rt, data, 'ordered');
  const frames = summaryFrames(rt);
  const fields = { company: 'name', description: 'description', first_cash: 'first_cash_injection', ownership: 'ownership',
    board: 'sv_board_seats', cost: 'cost', carrying: 'carrying_value', unrealized: 'unrealized_gain_loss', realized: 'realized_gain_loss' };
  frames.forEach((frame, i) => {
    const names = frame.findAll(n => /row_\d+\.company$/.test(n.name));
    assert.deepEqual(names.map(n => n.characters), SUMMARY_GROUPS[i].map(n => aliases[n] || n));
    for (const name of names) {
      const row = data.portfolio_summary.rows.find(row => row.name === name.characters);
      for (const [suffix, field] of Object.entries(fields)) {
        assert.equal(cellText(frame, name.name.replace(/company$/, suffix)), row[field]);
      }
    }
  });
  for (const [index, prefix, key] of [[1, 'company', 'portfolio_totals'], [2, 'operator', 'totals']]) {
    for (const [suffix, field] of Object.entries({ cost: 'cost', carrying: 'carrying_value', unrealized: 'unrealized_gain_loss', realized: 'realized_gain_loss' })) {
      assert.equal(cellText(frames[index], `portfolio.${prefix}.total_${suffix}`), data.portfolio_summary[key][field]);
    }
  }
  assert.deepEqual(data, original);
});

test('wrapped portfolio rows expand without overlap and keep columns aligned', async () => {
  const rt = runtime(), data = payload();
  data.portfolio_summary.rows[0].description = 'A carefully measured description that wraps over several lines while keeping the next row and all of its financial values safely below it. More approved text.';
  await update(rt, data, 'wrap');
  const frame = summaryFrames(rt)[0];
  const description = frame.findOne(n => n.name === 'portfolio.company.row_1.description');
  const next = frame.findOne(n => n.name === 'portfolio.company.row_2.company');
  assert.ok(description.height > 50);
  assert.ok(next.y >= description.y + description.height + 8);
  assert.ok(frame.findOne(n => n.name === 'portfolio.company.row_1.bg').height >= description.height + 12);
  for (const f of summaryFrames(rt)) {
    assert.ok(f.findAll(n => n.type === 'TEXT' && n.name.startsWith('portfolio.')).every(n => n.y + n.height <= 1032));
    const headers = f.findAll(n => /^portfolio\.header_\d+\.text$/.test(n.name));
    assert.ok(headers.every(n => n.y + n.height <= 220));
  }
});

test('all approved portfolio descriptions fit the fixed three-page layout', async () => {
  const rt = runtime(), data = payload();
  data.portfolio_summary.rows = JSON.parse(fs.readFileSync('tests/fixtures/portfolio_layout.json', 'utf8'));
  await update(rt, data, 'approved-descriptions');
  assert.equal(summaryFrames(rt).length, 3);
  for (const frame of summaryFrames(rt)) {
    const rows = frame.findAll(n => /row_\d+\.company$/.test(n.name));
    rows.forEach((row, i) => {
      const description = frame.findOne(n => n.name === row.name.replace(/company$/, 'description'));
      assert.ok(description.y + description.height < 1032);
      if (rows[i + 1]) assert.ok(description.y + description.height + 8 <= rows[i + 1].y);
    });
  }
});

test('unknown companies, duplicate aliases and oversized text fail without replacing the last report', async () => {
  for (const [change, error] of [
    [data => data.portfolio_summary.rows.push({ name: 'Unapproved company' }), /no position/],
    [data => data.portfolio_summary.rows.push({ name: 'TRUKKER' }), /Duplicate portfolio/],
    [data => { data.portfolio_summary.rows[0].description = 'Long description '.repeat(300); }, /footer boundary/],
  ]) {
    const rt = runtime();
    await update(rt, payload(), 'good');
    const ids = rt.page.children.map(n => n.id);
    const invalid = payload(); change(invalid);
    await assert.rejects(update(rt, invalid, 'invalid'), error);
    assert.deepEqual(rt.page.children.map(n => n.id), ids);
  }
});

test('a failed render preserves the previous report and removes staging content', async () => {
  const rt = runtime();
  await update(rt, payload(), 'good');
  const ids = rt.page.children.map(n => n.id);
  rt.figma.loadFontAsync = async () => { throw new Error('Font unavailable'); };
  await assert.rejects(update(rt, payload(), 'bad'), /Font unavailable/);
  assert.deepEqual(rt.page.children.map(n => n.id), ids);
});

test('different quarters coexist and page switching blocks an update', async () => {
  const rt = runtime();
  const first = await update(rt, payload(), 'q2');
  await update(rt, payload('Q3 2026'), 'q3');
  assert.equal(rt.page.children.length, first.createdCount * 2);
  const q2 = rt.page.children.filter(n => n.getPluginData('sarmayacar_report') === 'Q2 2026');
  const q3 = rt.page.children.filter(n => n.getPluginData('sarmayacar_report') === 'Q3 2026');
  assert.ok(Math.min(...q3.map(n => n.y)) > Math.max(...q2.map(n => n.y + n.height)));
  rt.figma.currentPage = new rt.Node('PAGE');
  await assert.rejects(update(rt, payload(), 'wrong'), /paired Figma page/);
});

test('pairing credentials stay in local client storage, not report node metadata', async () => {
  const rt = runtime();
  const connection = { url: 'https://bridge.example', code: 'id.secret', target_id: 'target' };
  await rt.figma.ui.onmessage({ type: 'ready', uiSession: TEST_UI_SESSION });
  await rt.figma.ui.onmessage({ type: 'save-connection', connection, uiSession: TEST_UI_SESSION });
  assert.equal(rt.page.getPluginData('sarmayacar_connection'), 'id');
  assert.equal(JSON.stringify(rt.page.data).includes('secret'), false);
  await rt.figma.ui.onmessage({ type: 'ready', uiSession: TEST_UI_SESSION });
  assert.equal(rt.messages.at(-1).saved.code, 'id.secret');
});

test('Figma storage errors return a connection error instead of leaving setup waiting', async () => {
  const rt = runtime();
  await rt.figma.ui.onmessage({ type: 'ready', uiSession: TEST_UI_SESSION });
  rt.figma.clientStorage.setAsync = async () => { throw new Error('Storage unavailable'); };
  await rt.figma.ui.onmessage({ type: 'save-connection', connection: { code: 'id.secret', target_id: 'target' }, uiSession: TEST_UI_SESSION });
  assert.equal(rt.messages.at(-1).type, 'connection-error');
  assert.equal(rt.messages.at(-1).error, 'Storage unavailable');
  assert.equal(rt.page.getPluginData('sarmayacar_connection'), '');
});

test('main rejects stale or missing UI sessions without saving credentials', async () => {
  const rt = runtime();
  await rt.figma.ui.onmessage(null);
  await rt.figma.ui.onmessage({ type: 'ready' });
  assert.equal(rt.messages.length, 1);
  assert.equal(rt.messages[0].type, 'plugin-boot');
  await rt.figma.ui.onmessage({ type: 'ready', uiSession: TEST_UI_SESSION });
  const count = rt.messages.length;
  await rt.figma.ui.onmessage({ type: 'ready', uiSession: 'b'.repeat(32) });
  await rt.figma.ui.onmessage({ type: 'save-connection', uiSession: 'b'.repeat(32), connection: { code: 'id.secret', target_id: 'target' } });
  assert.equal(rt.messages.length, count);
  assert.equal(rt.page.getPluginData('sarmayacar_connection'), '');
});

function uiRuntime(options = {}) {
  const elements = new Map(), messages = [], requests = [], timers = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, { value: '', textContent: '', checked: true });
    return elements.get(id);
  };
  const parent = { postMessage: envelope => messages.push(envelope) };
  let nextTimer = 0;
  const ui = vm.createContext({ document: { getElementById: element }, parent, window: {},
    crypto: options.crypto || { getRandomValues: values => webcrypto.getRandomValues(values) }, URL, AbortController,
    setTimeout: (fn, ms) => { timers.set(++nextTimer, { fn, ms }); return nextTimer; },
    clearTimeout: id => timers.delete(id),
    fetch: async (url, options) => { requests.push({ url, body: JSON.parse(options.body) }); throw new Error('Test service unavailable'); },
  });
  const source = fs.readFileSync('figma_quarterly_report_plugin/ui.html', 'utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
  assert.doesNotThrow(() => vm.runInContext(source, ui));
  const reply = { type: 'context', build: '2026.10.13.3', uiSession: messages[0]?.pluginMessage.uiSession, context: {
    file_name: 'Test', page_name: 'Report', page_id: '1:2', target_id: '',
  } };
  return { ui, element, messages, requests, timers, reply, parent };
}

test('plugin UI accepts session-matched replies from Figma relays, not unrelated messages', async () => {
  const { ui, element, messages, requests, reply } = uiRuntime();
  assert.equal(typeof element('connect').onclick, 'function');
  assert.equal(messages[0].pluginId, 'sarmayacar-quarterly-report-importer');
  assert.equal(messages[0].pluginMessage.type, 'ready');
  assert.match(reply.uiSession, /^[a-f0-9]{32}$/);
  assert.equal(element('connect').disabled, true);
  assert.doesNotThrow(() => ui.window.onmessage({ source: null, data: null }));
  ui.window.onmessage({ source: {}, data: { pluginMessage: { ...reply, uiSession: 'wrong' } } });
  ui.window.onmessage({ source: {}, data: { pluginId: 'another-plugin', pluginMessage: reply } });
  ui.window.onmessage({ source: {}, data: { pluginMessage: { ...reply, context: null } } });
  assert.equal(element('destination').textContent, '');
  element('service').value = 'http://127.0.0.1:8765';
  element('code').value = 'test.' + 'x'.repeat(43);
  await element('connect').onclick();
  assert.match(element('connection-status').textContent, /page details/i);
  assert.equal(requests.length, 0);
  ui.window.onmessage({ source: {}, data: { pluginMessage: reply } });
  assert.equal(element('destination').textContent, 'Test / Report');
  assert.equal(element('connect').disabled, false);
  assert.match(element('build').textContent, /Figma code 2026\.10\.13\.3/);
  ui.window.onmessage({ source: ui.window, data: { pluginMessage: {
    ...reply, context: { ...reply.context, page_name: 'Host relay' },
  } } });
  assert.equal(element('destination').textContent, 'Test / Report');
  await element('connect').onclick();
  assert.equal(requests.length, 1);
  assert.match(requests[0].body.target_id, /^[a-f0-9]{32}$/);
  assert.match(element('connection-status').textContent, /Test service unavailable/);
  assert.equal(element('connect').disabled, false);
  ui.crypto = {};
  await element('connect').onclick();
  assert.match(element('connection-status').textContent, /Secure random IDs are unavailable/);
  assert.equal(element('connect').disabled, false);
  assert.equal(requests.length, 1);
});

test('startup retries lost ready messages and reports a bounded diagnostic without contacting the service', () => {
  const { ui, element, messages, requests, timers } = uiRuntime();
  const runTimer = ms => {
    const [id, timer] = [...timers].find(([, t]) => t.ms === ms);
    timers.delete(id); timer.fn();
  };
  for (let i = 0; i < 3; i++) runTimer(1500);
  assert.equal(messages.length, 4);
  assert.ok(messages.every(m => m.pluginMessage.uiSession === messages[0].pluginMessage.uiSession));
  assert.equal([...timers.values()].filter(t => t.ms === 1500).length, 0);
  ui.window.onmessage({ source: {}, data: { pluginMessage: { type: 'plugin-boot', build: '2026.10.13.3' } } });
  runTimer(8000);
  assert.match(element('connection-status').textContent, /did not complete the handshake/);
  assert.equal(element('connect').disabled, true);
  assert.equal(requests.length, 0);
});

test('startup handles null-source replies and missing secure randomness explicitly', () => {
  const { ui, element, reply, timers } = uiRuntime();
  ui.window.onmessage({ source: null, data: { pluginMessage: reply } });
  assert.equal(element('destination').textContent, 'Test / Report');
  assert.equal(timers.size, 0);
  const failed = uiRuntime({ crypto: {} });
  assert.match(failed.element('connection-status').textContent, /Secure random IDs are unavailable/);
  assert.equal(failed.element('connect').disabled, true);
  assert.equal(failed.messages.length, 0);
});

test('live HTTP service to plugin UI to renderer to completion, including a lost acknowledgement', { timeout: 30000 }, async t => {
  const temp = await mkdtemp(path.join(os.tmpdir(), 'figma-bridge-test-'));
  const socket = net.createServer();
  await new Promise(resolve => socket.listen(0, '127.0.0.1', resolve));
  const port = socket.address().port;
  await new Promise(resolve => socket.close(resolve));
  const externalUrl = global.process.env.FIGMA_TEST_BRIDGE_URL;
  if (externalUrl) assert.match(externalUrl, /^http:\/\/127\.0\.0\.1:\d+$/);
  const token = externalUrl ? global.process.env.FIGMA_TEST_PUBLISHER_TOKEN : webcrypto.randomUUID() + webcrypto.randomUUID();
  const process = externalUrl ? null : spawn('.venv/bin/python', ['-m', 'uvicorn', 'figma_bridge.service:create_app', '--factory', '--host', '127.0.0.1', '--port', String(port), '--no-access-log'], {
    env: { ...global.process.env, FIGMA_BRIDGE_PUBLISHER_TOKEN: token, FIGMA_BRIDGE_DB: path.join(temp, 'bridge.sqlite3') },
    stdio: ['ignore', 'ignore', 'pipe'],
  });
  let logs = '';
  process?.stderr.on('data', chunk => { logs += chunk; });
  const timerIds = new Set();
  t.after(async () => {
    for (const id of timerIds) { clearTimeout(id); clearInterval(id); }
    if (process && process.exitCode === null) {
      process.kill();
      await new Promise(resolve => process.once('exit', resolve));
    }
    await rm(temp, { recursive: true, force: true });
  });
  const url = externalUrl || 'http://127.0.0.1:' + port;
  const waitFor = async predicate => {
    for (let i = 0; i < 200; i++) {
      if (await predicate()) return;
      await new Promise(resolve => setTimeout(resolve, 25));
    }
    throw new Error('Timed out: ' + logs);
  };
  await waitFor(async () => { try { return (await fetch(url + '/health')).ok; } catch (_) { return false; } });
  const publisher = async (route, data) => {
    const response = await fetch(url + route, { method: data ? 'POST' : 'GET',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' }, body: data ? JSON.stringify(data) : undefined });
    assert.equal(response.status, 200, await response.clone().text());
    return response.json();
  };
  const connection = await publisher('/v1/connections', { name: 'Integration test' });
  const elements = new Map();
  const element = id => {
    if (!elements.has(id)) elements.set(id, { value: '', textContent: '', hidden: false, checked: true });
    return elements.get(id);
  };
  const rt = runtime();
  let ui;
  const parent = { postMessage: ({ pluginMessage }) => {
    Promise.resolve().then(() => rt.figma.ui.onmessage(pluginMessage));
  } };
  rt.figma.ui.postMessage = message => {
    rt.messages.push(message);
    ui.window.onmessage({ source: null, data: { pluginMessage: message } });
  };
  let loseAck = true;
  ui = vm.createContext({ document: { getElementById: element }, parent, window: {},
    crypto: { getRandomValues: values => webcrypto.getRandomValues(values) }, URL, AbortController,
    fetch: async (address, options) => {
      const response = await fetch(address, options);
      if (loseAck && address.endsWith('/result')) {
        loseAck = false;
        throw new Error('Simulated lost acknowledgement after server accepted it');
      }
      return response;
    },
    setTimeout: (fn, ms) => { const id = setTimeout(fn, ms); timerIds.add(id); return id; },
    clearTimeout, setInterval: (fn, ms) => { const id = setInterval(fn, ms); timerIds.add(id); return id; }, clearInterval,
  });
  const source = fs.readFileSync('figma_quarterly_report_plugin/ui.html', 'utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
  vm.runInContext(source, ui);
  await waitFor(async () => element('destination').textContent.includes('Test report'));
  element('service').value = url;
  element('code').value = connection.pairing_code;
  await vm.runInContext('connect()', ui);
  await waitFor(async () => !vm.runInContext('inFlight', ui));
  const first = await publisher('/v1/connections/' + connection.id + '/jobs', { request_id: 'first', payload: payload() });
  await vm.runInContext('poll()', ui);
  assert.ok(element('connection-status').textContent.includes('confirmation pending'));
  const ids = rt.page.children.map(n => n.id);
  await vm.runInContext('poll()', ui);
  assert.deepEqual(rt.page.children.map(n => n.id), ids);
  const listing = await publisher('/v1/connections');
  assert.equal(listing.connections[0].latest_job.id, first.id);
  assert.equal(listing.connections[0].latest_job.status, 'succeeded');
  assert.equal(listing.connections[0].latest_job.result.page_count, 9);
  const changed = payload(); changed.companies[0].description = 'Updated through the connection';
  await publisher('/v1/connections/' + connection.id + '/jobs', { request_id: 'second', payload: changed });
  await vm.runInContext('poll()', ui);
  assert.equal(rt.page.children.length, 9);
  assert.ok(rt.page.findOne(n => n.characters === 'Updated through the connection'));
  assert.equal((await publisher('/v1/connections')).connections[0].latest_job.status, 'succeeded');
});
