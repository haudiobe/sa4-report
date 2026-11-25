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
import zlib from 'zlib';

const require = createRequire(import.meta.url);
const J = require('./issue28-jobs.js');
const argv = process.argv.slice(2);
const arg = (n) => { const i = argv.indexOf(n); return i >= 0 ? argv[i + 1] : undefined; };
const write = argv.includes('--write');
const minutes = Number(arg('--minutes') || 45);
// The first job waits longer: until the runner is installed (it28InstallRunner, once per copy).
const firstMinutes = Number(arg('--first-minutes') || minutes);
let jobsWaited = 0;
const ROOT = execFileSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const rc = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.clasprc.json'), 'utf8'));
const cred = rc.tokens[Object.keys(rc.tokens)[0]];
let token = null;
let tokenAt = 0;
// The access token is kept for 50 minutes (it is valid for 60), not fetched again for every poll.
const refresh = async () => {
  if (token && Date.now() - tokenAt < 50 * 60000) return;
  const t = await (await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: cred.client_id, client_secret: cred.client_secret, refresh_token: cred.refresh_token, grant_type: 'refresh_token' }) })).json();
  if (!t.access_token) throw new Error('the clasp token could not be refreshed (' + (t.error || '?') + ')');
  token = t.access_token;
  tokenAt = Date.now();
};
/** Drive metadata of the result files whose name contains `text` (drive.metadata.readonly: names and descriptions only). */
// A passing Drive or network error (seen: HTTP 500) is retried; the poll goes on.
const resultFiles = async (text) => {
  const q = encodeURIComponent("name contains '" + text.replace(/'/g, "\\'") + "' and trashed = false");
  let last = '';
  for (let attempt = 1; attempt <= 6; attempt++) {
    try {
      await refresh();
      const r = await fetch('https://www.googleapis.com/drive/v3/files?pageSize=200&orderBy=createdTime desc&fields=files(name,description,createdTime)&q=' + q, { headers: { authorization: 'Bearer ' + token } });
      if (r.ok) return (await r.json()).files || [];
      last = 'Drive answered ' + r.status;
      if (r.status < 500 && r.status !== 429) throw new Error(last);
    } catch (e) {
      last = e.message;
      if (/Drive answered 4\d\d/.test(last) && !/429/.test(last)) throw e;
    }
    await sleep(attempt * 10000);
  }
  throw new Error(last + ' (6 attempts)');
};
const runnerOf = async (documentName) => {
  const events = (await resultFiles('IT28 result runner')).map((f) => { try { return JSON.parse(f.description || '{}'); } catch (e) { return {}; } }).filter((e) => !documentName || e.document === documentName);
  return events[0] || null;
};

async function waitFor(job, target) {
  const dir = path.join(ROOT, 'dist', 'integration-test', 'i28-copies', target || job.target || 'unknown', 'jobs');
  fs.mkdirSync(dir, { recursive: true });
  const limit = jobsWaited++ === 0 ? firstMinutes : minutes;
  let lastStatus = '';
  const until = Date.now() + limit * 60000;
  process.stdout.write('Waiting for the runner (job ' + job.id + ', up to ' + limit + ' min)');
  for (;;) {
    const st = (await resultFiles('IT28 status')).map((f) => { try { return JSON.parse(f.description || '{}'); } catch (e) { return {}; } }).find((s) => s.status && String(s.status).indexOf(job.id) !== -1);
    if (st && st.status !== lastStatus) { lastStatus = st.status; process.stdout.write('\n  [' + st.at.slice(11, 19) + ' UTC] ' + st.status); }
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
    if (Date.now() > until) { console.log('\nNo result within ' + limit + ' minutes (is the runner installed? see it28InstallRunner).'); return { ok: false, timeout: true }; }
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
const plan = argv.includes('--plan');
// Before a plan: the copy's TDoc list as the update will fetch it (read locally, public URL). On 2026-11-25 the Portal
// list of 86178 turned for a while into the list of another meeting (S4aP260008..0043, without its agenda TDoc);
// the report refused it, rightly. A plan does not start on such a list.
const LISTS = { main: { url: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx', must: 'S4-26' },
  main2: { url: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx', must: 'S4-26' },
  ahg: { url: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178', must: 'S4aP260149' } };
/** One file of a zip (an xlsx), from its central directory; null when not there. */
const zipFile = (buf, name) => {
  let end = buf.length - 22;
  while (end >= 0 && buf.readUInt32LE(end) !== 0x06054b50) end--;
  if (end < 0) return null;
  let at = buf.readUInt32LE(end + 16);
  for (let n = buf.readUInt16LE(end + 10); n > 0; n--) {
    if (buf.readUInt32LE(at) !== 0x02014b50) return null;
    const method = buf.readUInt16LE(at + 10), size = buf.readUInt32LE(at + 20), nameLen = buf.readUInt16LE(at + 28), extra = buf.readUInt16LE(at + 30), comment = buf.readUInt16LE(at + 32), local = buf.readUInt32LE(at + 42);
    if (buf.slice(at + 46, at + 46 + nameLen).toString() === name) {
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const data = buf.slice(start, start + size);
      return method === 8 ? zlib.inflateRawSync(data) : data;
    }
    at += 46 + nameLen + extra + comment;
  }
  return null;
};
const precheck = async (t) => {
  const l = LISTS[t];
  const r = await fetch(l.url);
  const b = Buffer.from(await r.arrayBuffer());
  let strings = null;
  try { strings = zipFile(b, 'xl/sharedStrings.xml'); } catch (e) { strings = null; }
  const ok = r.ok && !!strings && strings.toString('utf8').indexOf(l.must) !== -1;
  console.log('TDoc list precheck (' + t + '): HTTP ' + r.status + ', ' + b.length + ' bytes, ' + (ok ? 'contains ' + l.must : 'does NOT contain ' + l.must + ' -- the plan does not start'));
  return ok;
};
// --from <n>: resume the plan at its n-th job (1 = the first), after jobs that already passed.
const from = Math.max(1, Number(arg('--from') || 1));
const steps = plan ? J.PLAN[target].slice(from - 1) : [[arg('--stage'), arg('--job')]];
const pushOriginal = () => {
  const r = spawnSync(process.execPath, [path.join(ROOT, 'tools', 'integration-test', 'push-issue28-copy.mjs'), '--target', target, '--script', script, '--stage', 'original', '--write'], { encoding: 'utf8' });
  process.stdout.write((r.stdout || '') + (r.stderr || ''));
  return r.status === 0;
};
// Cleanup on any failure: the uninstall job (every trigger, document and script property the tests left), at the stage
// the copy is at; then, once it is proven clean, the copy's original script. If the runner does not answer, its
// 72-hour limit removes the trigger and the settings, and the original is pushed later.
const cleanup = async (stage) => {
  console.log('\n=== ' + target + ': CLEANUP (uninstall job at stage ' + stage + ')');
  const r = await one(target, script, stage, 'uninstall');
  if (r.ok) { console.log('Cleanup verified.'); return pushOriginal(); }
  console.log('Cleanup NOT verified: the runner removes itself and the test settings at its 72-hour limit; push the original after that.');
  return false;
};
const statePath = path.join(ROOT, 'dist', 'integration-test', 'i28-copies', target, 'state.json');
if (plan && !argv.includes('--no-precheck') && !(await precheck(target))) process.exit(3);
for (const [stage, name] of steps) {
  console.log('\n=== ' + target + ': stage ' + stage + ', job ' + name);
  const r = await one(target, script, stage, name);
  if (!r.ok) {
    console.log('Stopped at stage ' + stage + ', job ' + name + '.');
    if (write && plan && !r.dryRun) {
      const at = fs.existsSync(statePath) ? JSON.parse(fs.readFileSync(statePath, 'utf8')).stage : stage;
      if (name !== 'uninstall') await cleanup(at);
    }
    process.exit(1);
  }
}
if (write && plan) {
  console.log('\n=== ' + target + ': the original script back (cleanup verified by the uninstall job)');
  if (!pushOriginal()) { console.log('The original was NOT pushed back.'); process.exit(1); }
}
console.log('\nAll done.');
