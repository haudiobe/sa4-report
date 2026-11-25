#!/usr/bin/env node
/**
 * INTEGRATION TEST ONLY -- issue #28. Pushes one stage of the issue #28 candidates into the bound script of one of
 * the two DISPOSABLE copies the user made on 2026-11-23 -- and nowhere else. Rules: issue28-copy-rules.js.
 *
 *   node tools/integration-test/push-issue28-copy.mjs --target main|ahg --script <the copy's Script ID> --stage <stage>          (dry run)
 *   node tools/integration-test/push-issue28-copy.mjs ... --write
 *
 * Stages: base      the copy's own script as first seen + Issue28Check.js (report code unchanged)
 *         t33       T-2026.10.33 (b7299c1 + the production Release.js) + Issue28Check.js
 *         A | B | C the phase code at its PINNED commit (A; A+B; A+B+C) + Issue28Check.js; Release.js "NOT RELEASED"
 *         original  the copy's own script exactly as first seen (no Issue28Check.js): the end state
 *   --job <file.json>  also pushes Issue28Job.js (var IT28_JOB_ = <file>) for the runner of Issue28Check.js
 *                      (run-issue28-validation.mjs writes the file); refused with 'original'.
 *
 * Safeguards (each refuses; nothing is written):
 *  - identity: the script and its parent document must have the digests of the target copy, the parent must be a
 *    "DISPOSABLE" document not in the trash, and neither may be the template, a SA4#138 report, Legacy or CENTRAL;
 *  - original: on first sight the copy's script must equal the pinned original digests file by file; it is then
 *    saved (dist/integration-test/i28-copies/<target>/original, Git-ignored) and re-verified on every run;
 *  - compare-and-swap: the copy's script must equal the last state this tool recorded (state.json; the original
 *    before the first push), checked when the run starts and again, with the identity, immediately before the push;
 *  - pins: a stage's branch must still have its pinned code; the pushed code is taken from the pinned commit.
 * After a push the script is read back and compared file by file; state.json is updated only when it matches.
 */
import fs from 'fs';
import os from 'os';
import path from 'path';
import { createRequire } from 'module';
import { execFileSync } from 'child_process';

const require = createRequire(import.meta.url);
const R = require('./issue28-copy-rules.js');

const argv = process.argv.slice(2);
const arg = (n) => { const i = argv.indexOf(n); return i >= 0 ? argv[i + 1] : undefined; };
const write = argv.includes('--write');
const target = arg('--target');
const scriptId = String(arg('--script') || '').trim();
const stage = arg('--stage');
const jobPath = arg('--job');
const fail = (lines) => { console.error('\nREFUSED (nothing written):\n' + [].concat(lines).map((l) => '  - ' + l).join('\n')); process.exit(2); };
if (!R.TARGETS[target]) fail('--target must be main or ahg');
if (!R.STAGES.includes(stage)) fail('--stage must be one of ' + R.STAGES.join(', '));
if (!scriptId) fail('--script is required (the copy\'s Script ID: Extensions > Apps Script > Project Settings)');
const early = [];
if (R.REFUSED_SCRIPT_DIGESTS[R.idDigest(scriptId)]) early.push('the script is ' + R.REFUSED_SCRIPT_DIGESTS[R.idDigest(scriptId)]);
if (R.idDigest(scriptId) !== R.TARGETS[target].scriptDigest) early.push('the script is not the bound script of the ' + target + ' copy');
if (early.length) fail(early);

const ROOT = execFileSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
const git = (...a) => execFileSync('git', ['-C', ROOT, ...a], { encoding: 'utf8', maxBuffer: 1 << 28 });
const codeAt = (ref, p) => git('show', ref + ':' + p);
const pinProblem = R.pinRefusal(stage, codeAt);
if (pinProblem) fail(pinProblem);
const CHECK_REF = R.STAGE_PINS.C.branch;
const checkSource = codeAt(CHECK_REF, 'tools/integration-test/Issue28Check.js');
if (R.sourceSha(checkSource) !== R.CHECK_SHA) fail('Issue28Check.js on ' + CHECK_REF + ' is not the pinned harness (' + R.CHECK_SHA.slice(0, 12) + '): re-pin deliberately');

// --- Google APIs with the clasp user's token (never printed) ---
const rc = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.clasprc.json'), 'utf8'));
const cred = rc.tokens[Object.keys(rc.tokens)[0]];
const tok = await (await fetch('https://oauth2.googleapis.com/token', { method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' },
  body: new URLSearchParams({ client_id: cred.client_id, client_secret: cred.client_secret, refresh_token: cred.refresh_token, grant_type: 'refresh_token' }) })).json();
if (!tok.access_token) fail('the clasp token could not be refreshed');
const api = async (method, url, body) => {
  const r = await fetch(url, { method, headers: { authorization: 'Bearer ' + tok.access_token, 'content-type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
  const t = await r.text();
  if (!r.ok) throw new Error(method + ' ' + url.replace(/\?.*$/, '') + ' -> ' + r.status + ' ' + t.slice(0, 300));
  return JSON.parse(t);
};
const identity = async () => {
  const project = await api('GET', `https://script.googleapis.com/v1/projects/${scriptId}`);
  const file = project.parentId ? await api('GET', `https://www.googleapis.com/drive/v3/files/${project.parentId}?fields=id,name,trashed,modifiedTime`) : null;
  return { project, file, refusals: R.identityRefusals(target, scriptId, project, file) };
};
const content = async () => (await api('GET', `https://script.googleapis.com/v1/projects/${scriptId}/content`)).files.map((f) => ({ name: f.name, type: f.type, source: f.source }));

// --- 1. identity ---
const id1 = await identity();
if (id1.refusals.length) fail(id1.refusals);
const documentId = id1.project.parentId;

// --- 2. the original (pinned), the recorded state, the copy now ---
const BASE = path.join(ROOT, 'dist', 'integration-test', 'i28-copies');
const OUT = path.join(BASE, target);
const ext = (f) => (f.type === 'JSON' ? '.json' : '.js');
const save = (where, files) => { fs.mkdirSync(where, { recursive: true }); files.forEach((f) => fs.writeFileSync(path.join(where, f.name + ext(f)), f.source)); fs.writeFileSync(path.join(where, 'files.json'), JSON.stringify(files.map((f) => ({ name: f.name, type: f.type, sha256: R.sourceSha(f.source) })), null, 1)); };
const load = (where) => JSON.parse(fs.readFileSync(path.join(where, 'files.json'), 'utf8')).map((f) => ({ name: f.name, type: f.type, source: fs.readFileSync(path.join(where, f.name + ext(f)), 'utf8') }));
const now = await content();
const originalDir = path.join(OUT, 'original');
if (!fs.existsSync(path.join(originalDir, 'files.json'))) {
  const o = R.originalShaMap(target, now);
  if (!o.matchesPin) fail(['first sight of the ' + target + ' copy, but its script is not the pinned original:'].concat(R.drift(R.TARGETS[target].original, Object.fromEntries(now.map((f) => [f.name, R.sourceSha(f.source)])))));
  save(originalDir, now);
  console.log('First sight: the script equals the pinned original; saved to ' + originalDir);
}
const original = load(originalDir);
const orig = R.originalShaMap(target, original);
if (!orig.matchesPin) fail('the saved original (' + originalDir + ') no longer matches the pinned digests');
const statePath = path.join(OUT, 'state.json');
const state = fs.existsSync(statePath) ? JSON.parse(fs.readFileSync(statePath, 'utf8')) : null;
// The recorded state, recomputed from its saved read-back (so the comparison rules of today apply to it); the saved
// read-back must still be what was recorded for every script file.
let expected = orig.canonical;
if (state) {
  const recorded = load(path.join(OUT, state.evidence, 'readback'));
  const bad = recorded.filter((f) => f.type !== 'JSON' && state.files[f.name] !== R.sourceSha(f.source)).map((f) => f.name);
  if (bad.length || recorded.length !== Object.keys(state.files).length) fail('the saved read-back of the recorded state (' + state.evidence + ') does not match state.json: ' + (bad.join(', ') || 'file count'));
  expected = R.shaMap(recorded);
}
const d1 = R.drift(expected, R.shaMap(now));
if (d1.length) fail(['the copy\'s script is not the last recorded state (' + (state ? 'stage ' + state.stage + ' of ' + state.at : 'the original') + '):'].concat(d1));

// --- 3. what is pushed ---
const t33Dir = path.join(BASE, 'main', 'original');
const t33Release = fs.existsSync(path.join(t33Dir, 'Release.js')) ? fs.readFileSync(path.join(t33Dir, 'Release.js'), 'utf8') : null;
let jobSource = null;
if (jobPath) {
  const job = JSON.parse(fs.readFileSync(jobPath, 'utf8'));
  if (!job.id || !Array.isArray(job.steps) || !job.steps.length) fail('the job file needs an id and steps');
  jobSource = '/** INTEGRATION TEST ONLY -- the job of the issue #28 runner (run-issue28-validation.mjs). */\nvar IT28_JOB_ = ' + JSON.stringify(job, null, 1) + ';\n';
}
// Removing the harness while its runner is installed would leave a trigger whose function is gone (an error every 5 minutes).
if (stage === 'original') {
  const runner = await api('GET', 'https://www.googleapis.com/drive/v3/files?pageSize=50&orderBy=createdTime desc&fields=files(name,description,createdTime)&q=' + encodeURIComponent("name contains 'IT28 result runner' and trashed = false"));
  const mine = (runner.files || []).map((f) => { try { return JSON.parse(f.description || '{}'); } catch (e) { return {}; } }).filter((r) => r.document === id1.file.name);
  if (mine.length && mine[0].event === 'installed') fail('the runner of this copy is still installed (since ' + mine[0].at + '): run its uninstall job first');
}
let files;
try {
  files = R.stageFiles(stage, { original, checkSource, codeAt, t33Release, jobSource, stamp: new Date().toISOString(),
    codeVersionOf: (src) => (src.match(/^ \* Version: (\S+)/m) || [])[1] });
} catch (e) { fail(e.message + (t33Release ? '' : ' (the main copy\'s original is needed: push the main copy\'s base stage first)')); }
const pin = R.STAGE_PINS[stage];
console.log(`Target ${target} copy: "${id1.file.name}"; script and parent verified by digest`);
console.log('Now:   ' + now.map((f) => f.name + ' ' + R.sourceSha(f.source).slice(0, 12)).join(', ') + '  (= ' + (state ? 'recorded state, stage ' + state.stage : 'pinned original') + ')');
console.log('Push:  ' + files.map((f) => f.name + ' ' + R.sourceSha(f.source).slice(0, 12)).join(', ') + '  (stage ' + stage + (pin ? ' = ' + pin.commit.slice(0, 7) : '') + '; Issue28Check.js from ' + CHECK_REF + ')');
const harness = (f) => f.name === 'Issue28Check' || f.name === 'Issue28Job';
const reportUnchanged = (fs2) => R.drift(R.shaMap(original.filter((f) => !harness(f))), R.shaMap(fs2.filter((f) => !harness(f))));
if (stage === 'base' && reportUnchanged(files).length) fail('stage base would change report code: ' + reportUnchanged(files).join('; '));
const stamp = new Date().toISOString().replace(/[:.]/g, '').slice(0, 15) + 'Z';
const dir = path.join(OUT, stamp + '-' + stage);
const evidence = { target, documentDigest: R.idDigest(documentId), scriptDigest: R.idDigest(scriptId), documentName: id1.file.name, stage, pinnedCommit: pin ? pin.commit : null,
  checkFromRef: CHECK_REF, checkSha256: R.sourceSha(checkSource), at: new Date().toISOString(), mode: write ? 'write' : 'dry run', expectedFrom: state ? 'state ' + state.stage + ' ' + state.at : 'pinned original',
  before: R.shaMap(now), pushed: R.shaMap(files) };
if (!write) { console.log('\nDry run: nothing written. Add --write to push.'); process.exit(0); }

// --- 4. compare-and-swap immediately before the push, push, read back ---
save(path.join(dir, 'before'), now);
const id2 = await identity();
if (id2.refusals.length || id2.project.parentId !== documentId) fail(['identity changed since the run started:'].concat(id2.refusals));
const d2 = R.drift(expected, R.shaMap(await content()));
if (d2.length) fail(['the copy\'s script changed during this run:'].concat(d2));
await api('PUT', `https://script.googleapis.com/v1/projects/${scriptId}/content`, { files: files.map((f) => ({ name: f.name, type: f.type, source: f.source })) });
const back = await content();
save(path.join(dir, 'readback'), back);
const mismatch = R.drift(R.shaMap(files), R.shaMap(back));
evidence.readback = R.shaMap(back);
evidence.readbackMatches = !mismatch.length;
evidence.reportCodeEqualsOriginal = !reportUnchanged(back).length;
fs.writeFileSync(path.join(dir, 'evidence.json'), JSON.stringify(evidence, null, 1));
if (!mismatch.length) fs.writeFileSync(statePath, JSON.stringify({ stage, at: evidence.at, files: R.shaMap(back), evidence: path.relative(OUT, dir) }, null, 1));
console.log((mismatch.length ? 'WRITTEN, BUT THE READ-BACK DIFFERS (state not updated; the next run refuses):\n  ' + mismatch.join('\n  ') : 'Written and read back: every file matches.') +
  '\nReport code equals the original: ' + evidence.reportCodeEqualsOriginal + '\nEvidence: ' + path.join(dir, 'evidence.json'));
process.exit(mismatch.length ? 1 : 0);
