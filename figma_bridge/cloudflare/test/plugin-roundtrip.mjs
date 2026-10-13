import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { startBridge } from './helpers.mjs';

const bridge = await startBridge();
async function run(command, args) {
  const child = spawn(command, args, {
    cwd: fileURLToPath(new URL('../../../', import.meta.url)),
    env: { ...process.env, FIGMA_TEST_BRIDGE_URL: bridge.url, FIGMA_TEST_PUBLISHER_TOKEN: bridge.token },
    stdio: 'inherit',
  });
  return new Promise((resolve, reject) => {
    child.once('error', reject);
    child.once('exit', code => resolve(code ?? 1));
  });
}
try {
  process.exitCode = await run('.venv/bin/python', ['figma_bridge/cloudflare/test/streamlit-client.py']);
  if (!process.exitCode) process.exitCode = await run(process.execPath, ['--test', 'tests/figma_import.test.cjs']);
} finally { await bridge.close(); }
