#!/usr/bin/env node
/**
 * INTEGRATION TEST ONLY -- issue #28. Runs the validation on a DISPOSABLE copy without anybody at the editor: pushes a
 * stage with a job (push-issue28-copy.mjs --job: identity, pinned original, compare-and-swap, read-back), waits for
 * the runner of Issue28Check.js to run it (a 5-minute trigger in the copy, installed once with it28InstallRunner()),
 * reads the result from the Drive metadata of its result files, and checks it (issue28-jobs.js).
 *
 *   node tools/integration-test/run-issue28-validation.mjs --target main|ahg --script <id> --stage <stage> --job <job> [--write] [--minutes 45]
 *   node tools/integration-test/run-issue28-validation.mjs --target main|ahg --script <id> --plan [--write]     (the whole PLAN)
 *   node tools/integration-test/run-issue28-validation.mjs --wait <job id>                                       (read a result again)
 *
 * Without --write nothing is pushed (the push tool's dry run). Evidence: dist/integration-test/i28-copies/<target>/jobs/.
 */
import fs from 'fs';
import os from 'os';
import path from 'path';
import { createRequire } from 'module';
import { execFileSync, spawnSync } from 'child_process';

const require = createRequire(import.meta.url);
const J = require('./issue28-jobs.js');
const argv = process.argv.slice(2);
const arg = (n) => { const i = argv.indexOf(n); return i >= 0 ? argv[i + 1] : undefined; };
const write = argv.includes('--write');
const minutes = Number(arg('--minutes') || 45);
const ROOT = execFileSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const rc = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.clasprc.json'), 'utf8'));
const cred = rc.tokens[Object.keys(rc.tokens)[0]];
let token = null;
const refresh = async () => {
  const t = await (await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: cred.client_id, client_secret: cred.client_secret, refresh_token: cred.refresh_token, grant_type: 'refresh_token' }) })).json();
  if (!t.access_token) throw new Error('the clasp token could not be refreshed');
  token = t.access_token;
};
/** Drive metadata of the result files whose name contains `text` (drive.metadata.readonly: names and descriptions only). */
const resultFiles = async (text) => {
  await refresh();
  const q = encodeURIComponent("name contains '" + text.replace(/'/g, "\\'") + "' and trashed = false");
  const r = await fetch('https://www.googleapis.com/drive/v3/files?pageSize=200&orderBy=createdTime desc&fields=files(name,description,createdTime)&q=' + q, { headers: { authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error('Drive answered ' + r.status);
  return (await r.json()).files || [];
};
const runnerOf = async (documentName) => {
  const events = (await resultFiles('IT28 result runner')).map((f) => { try { return JSON.parse(f.description || '{}'); } catch (e) { return {}; } }).filter((e) => !documentName || e.document === documentName);
  return events[0] || null;
};

async function waitFor(job, target) {
  const dir = path.join(ROOT, 'dist', 'integration-test', 'i28-copies', target || job.target || 'unknown', 'jobs');
  fs.mkdirSync(dir, { recursive: true });
  const until = Date.now() + minutes * 60000;
  process.stdout.write('Waiting for the runner (job ' + job.id + ', up to ' + minutes + ' min)');
  for (;;) {
    const files = await resultFiles('IT28 result ' + job.id);
    const err = files.find((f) => / ERROR /.test(f.name));
    if (err) { console.log('\nRUNNER ERROR: ' + err.description); return { ok: false }; }
    const result = J.assemble(job.id, files);
    if (result) {
      const failures = job.steps ? J.evaluate(job, result) : [];
      fs.writeFileSync(path.join(dir, job.id + '.result.json'), JSON.stringify({ job, result, failures }, null, 1));
      console.log('\n' + (failures.length ? 'FAILED' : 'PASSED') + ': ' + job.id + ' (' + result.release + ', ' + result.startedAt + ' .. ' + result.finishedAt + ')');
      result.steps.forEach((s, i) => console.log('  ' + (i + 1) + '. ' + s.do + (s.ok ? '' : ' ERROR ' + s.error) + (s.label ? ' ' + s.label + ' sha ' + s.contentSha + ' tabs ' + s.tabs + ' TDoc tables ' + s.tdocTables : '') +
        (s.compare ? ' ' + s.compare + (s.identical ? ' IDENTICAL' : ' ' + s.lines + ' line(s)') : '') + (s.updateMs !== undefined ? ' ' + (s.updateMs / 1000).toFixed(1) + ' s' : '') + ' [' + s.ms + ' ms]'));
      failures.forEach((f) => console.log('  ! ' + f));
      return { ok: !failures.length, result, failures };
    }
    if (Date.now() > until) { console.log('\nNo result within ' + minutes + ' minutes (is the runner installed? see it28InstallRunner).'); return { ok: false, timeout: true }; }
    process.stdout.write('.');
    await sleep(60000);
  }
}

async function one(target, script, stage, name) {
  const stamp = new Date().toISOString().replace(/[-:.]/g, '').slice(0, 15);
  const job = J.makeJob(target, name, stamp);
  const dir = path.join(ROOT, 'dist', 'integration-test', 'i28-copies', target, 'jobs');
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, job.id + '.json');
  fs.writeFileSync(file, JSON.stringify({ id: job.id, steps: job.steps }, null, 1));
  const push = spawnSync(process.execPath, [path.join(ROOT, 'tools', 'integration-test', 'push-issue28-copy.mjs'), '--target', target, '--script', script, '--stage', stage, '--job', file].concat(write ? ['--write'] : []), { encoding: 'utf8' });
  process.stdout.write(push.stdout || '');
  if (push.status !== 0) { process.stdout.write(push.stderr || ''); return { ok: false, refused: true }; }
  if (!write) return { ok: true, dryRun: true };
  return waitFor(job, target);
}

if (arg('--wait')) {
  const id = arg('--wait');
  const target = id.split('-')[0];
  const stored = path.join(ROOT, 'dist', 'integration-test', 'i28-copies', target, 'jobs', id + '.json');
  const job = fs.existsSync(stored) ? Object.assign(J.makeJob(target, id.split('-').slice(1, -1).join('-'), id.split('-').pop()), JSON.parse(fs.readFileSync(stored, 'utf8'))) : { id };
  process.exit((await waitFor(job, target)).ok ? 0 : 1);
}
const target = arg('--target');
const script = arg('--script');
if (!J.PLAN[target] || !script) { console.error('--target main|ahg and --script are required'); process.exit(2); }
const steps = argv.includes('--plan') ? J.PLAN[target] : [[arg('--stage'), arg('--job')]];
for (const [stage, name] of steps) {
  console.log('\n=== ' + target + ': stage ' + stage + ', job ' + name);
  const r = await one(target, script, stage, name);
  if (!r.ok) { console.log('Stopped at stage ' + stage + ', job ' + name + '.'); process.exit(1); }
}
console.log('\nAll done.');
