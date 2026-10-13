import { readFile, writeFile } from 'node:fs/promises';

const id = process.argv[2];
if (!id || !/^[a-f\d]{8}-[a-f\d]{4}-[a-f\d]{4}-[a-f\d]{4}-[a-f\d]{12}$/i.test(id) || /^0+(?:-0+)*$/.test(id)) {
  throw new Error('Pass the real database UUID printed by wrangler d1 create. No credentials are needed here.');
}
const config = JSON.parse(await readFile(new URL('../wrangler.json', import.meta.url), 'utf8'));
config.d1_databases[0].database_id = id;
await writeFile(new URL('../wrangler.production.json', import.meta.url), JSON.stringify(config, null, 2) + '\n', { flag: 'wx' });
console.log('Created wrangler.production.json. Review the database ID before applying migrations and deploying.');
