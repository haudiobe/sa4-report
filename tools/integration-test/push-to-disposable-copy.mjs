/**
 * INTEGRATION TEST ONLY (issue #6 / PR #7): pushes the candidate of the
 * current commit into the bound script of a DISPOSABLE COPY of the Audio
 * report -- and nowhere else.
 *
 *   node tools/integration-test/push-to-disposable-copy.mjs --script <copy's Script ID> --doc <copy's document ID>          (dry run)
 *   node tools/integration-test/push-to-disposable-copy.mjs --script <id> --doc <id> --write
 *
 * Uses clasp's stored login (clasp must be logged in); prints no token.
 *  1. refuses the original Audio report, the master template, protected projects (guard.js);
 *  2. reads the script's parent document (Apps Script API) -- it must be --doc;
 *     reads the document's Drive name -- it must start with "DISPOSABLE TEST COPY";
 *  3. backs up the script as it is (clasp pull) and records SHA-256 sums; its Code.js must be T-2026.10.22;
 *  4. builds the bundle: Code.js, HyperLink.js, ReportCreator.js of this commit, the copy's own
 *     appsscript.json (checked equal to the repository's), a candidate Release.js, IntegrationCheck.js;
 *  5. --write: clasp push (never --force), pull it back, compare every file's SHA-256.
 * Output: dist/integration-test/<timestamp>/{before,bundle,readback}, evidence.json.
 *
 * --target audio-original --backup-doc <id> --backup-script <id>   (authorized 2026-11-12, issue #6)
 *   The Audio SWG SA4#138 report itself instead of a copy -- that document and its script only
 *   (guard.js audioOriginalRefusals). Requires a backup copy of the document named "BACKUP...",
 *   made after the original's last change, whose own bound script is T-2026.10.22; and the
 *   Audio script equal to the local script backup (dist/integration-test/audio-original-backup-*).
 * --restore <dir>   push the five files of a script backup (e.g. .../before, or the local backup's
 *   script/) instead of the candidate, with the same read-back. No backup document needed.
 */
import path from 'node:path';
import os from 'node:os';
import fs from 'node:fs';
import { spawnSync, execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { pathToFileURL, fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const { pushRefusals, audioOriginalRefusals, candidateReleaseJs, sha256 } = require('./guard.js');
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, '..', '..');

const argv = process.argv.slice(2);
const value = (name) => { const i = argv.indexOf(name); return i >= 0 ? argv[i + 1] : undefined; };
const write = argv.includes('--write');
const scriptId = value('--script');
const documentId = value('--doc');
const allowOtherBase = argv.includes('--allow-other-base');
const target = value('--target');
const audioMode = target === 'audio-original';
const backupDocId = value('--backup-doc');
const backupScriptId = value('--backup-script');
const restoreDir = value('--restore') ? path.resolve(value('--restore')) : null;
if (target && !audioMode) { console.error('--target accepts only audio-original.'); process.exit(2); }

function fail(lines) { console.error('\nREFUSED:\n' + [].concat(lines).map((l) => '  - ' + l).join('\n')); process.exit(2); }

// One rule set per mode; a restore of a script backup needs no backup document.
const rules = (facts) => !audioMode ? pushRefusals(facts)
  : audioOriginalRefusals(Object.assign({ backup: restoreDir ? { documentId: 'restore' } : { documentId: backupDocId } }, facts));
let refusals = rules({ scriptId, documentId });
if (refusals.length) fail(refusals);

const git = (...a) => execFileSync('git', a, { cwd: REPO, encoding: 'utf8' }).trim();
if (git('status', '--porcelain', '--', 'Code.js', 'HyperLink.js', 'template/ReportCreator.js', 'tools/integration-test')) {
  fail('Code.js, HyperLink.js, template/ReportCreator.js or tools/integration-test have uncommitted changes: the pushed candidate must be a commit.');
}
const commit = git('rev-parse', 'HEAD');
const branch = git('rev-parse', '--abbrev-ref', 'HEAD');

// clasp's own login, in this process only.
const CLASP = path.join(process.env.APPDATA, 'npm', 'node_modules', '@google', 'clasp');
const { initAuth } = await import(pathToFileURL(path.join(CLASP, 'build', 'src', 'auth', 'auth.js')).href);
const { google } = await import(pathToFileURL(path.join(CLASP, 'node_modules', 'googleapis', 'build', 'src', 'index.js')).href);
const claspUser = Object.keys(JSON.parse(fs.readFileSync(path.join(os.homedir(), '.clasprc.json'), 'utf8')).tokens || {})[0];
const { credentials } = await initAuth({ userKey: claspUser });
if (!credentials) fail('clasp is not logged in.');

const project = (await google.script({ version: 'v1', auth: credentials }).projects.get({ scriptId })).data;
const file = (await google.drive({ version: 'v3', auth: credentials }).files.get({ fileId: documentId, supportsAllDrives: true,
  fields: 'id,name,trashed,createdTime,modifiedTime,version,mimeType' })).data;
const auditBackup = {};
if (audioMode && !restoreDir) {
  if (!backupScriptId) fail('--backup-script is required (the Script ID of the backup copy).');
  const b = (await google.drive({ version: 'v3', auth: credentials }).files.get({ fileId: backupDocId, supportsAllDrives: true,
    fields: 'id,name,trashed,createdTime,mimeType' })).data;
  const bp = (await google.script({ version: 'v1', auth: credentials }).projects.get({ scriptId: backupScriptId })).data;
  Object.assign(auditBackup, { documentId: b.id, name: b.name, mimeType: b.mimeType, trashed: b.trashed, createdTime: b.createdTime, scriptId: backupScriptId, scriptParentId: bp.parentId || null });
}
refusals = audioMode
  ? rules({ scriptId, documentId, parentId: project.parentId || null, originalModifiedTime: file.modifiedTime,
      backup: restoreDir ? { documentId: 'restore' } : auditBackup })
  : pushRefusals({ scriptId, documentId, parentId: project.parentId || null, documentName: file.name, documentTrashed: file.trashed });
if (refusals.length) fail(refusals);

const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
const OUT = path.join(REPO, 'dist', 'integration-test', stamp);
const dirs = { before: path.join(OUT, 'before'), bundle: path.join(OUT, 'bundle'), readback: path.join(OUT, 'readback') };
if (audioMode && !restoreDir) dirs.backupScript = path.join(OUT, 'backup-copy-script');
Object.values(dirs).forEach((d) => fs.mkdirSync(d, { recursive: true }));

function clasp(dir, ...args) {
  const id = dir === dirs.backupScript ? backupScriptId : scriptId;
  fs.writeFileSync(path.join(dir, '.clasp.json'), JSON.stringify({ scriptId: id, rootDir: '.' }) + '\n');
  fs.writeFileSync(path.join(dir, '.claspignore'), '**/**\n!appsscript.json\n!*.js\n');
  const r = spawnSync('clasp', ['-u', claspUser, ...args], { cwd: dir, encoding: 'utf8', shell: process.platform === 'win32' });
  const text = (r.stdout || '') + (r.stderr || '');
  if (r.status !== 0 || /error|No credentials/i.test(text)) throw new Error('clasp ' + args.join(' ') + ' failed:\n' + text);
  return text;
}
const hashes = (dir) => Object.fromEntries(fs.readdirSync(dir).filter((n) => /\.(js|json)$/.test(n) && !n.startsWith('.'))
  .sort().map((n) => [n, sha256(fs.readFileSync(path.join(dir, n)))]));

// 3. Backup of the script as it is.
clasp(dirs.before, 'pull');
const before = hashes(dirs.before);
let localBackup = null;
if (audioMode) {
  const root = path.join(REPO, 'dist', 'integration-test');
  const found = fs.readdirSync(root).filter((n) => n.startsWith('audio-original-backup-')).sort().pop();
  if (!found) fail('No local script backup (dist/integration-test/audio-original-backup-*).');
  localBackup = path.join(root, found, 'script');
  if (!restoreDir) {
    clasp(dirs.backupScript, 'pull');
    auditBackup.scriptCodeSha256 = hashes(dirs.backupScript)['Code.js'];
  }
  const matches = JSON.stringify(hashes(localBackup)) === JSON.stringify(before);
  refusals = restoreDir ? [] : rules({ scriptId, documentId, parentId: project.parentId, originalModifiedTime: file.modifiedTime, baseCodeSha256: before['Code.js'],
    allowOtherBase, scriptMatchesLocalBackup: allowOtherBase ? undefined : matches, backup: auditBackup });
} else {
  refusals = restoreDir ? pushRefusals({ scriptId, documentId, parentId: project.parentId, documentName: file.name })
    : pushRefusals({ scriptId, documentId, parentId: project.parentId, documentName: file.name, baseCodeSha256: before['Code.js'], allowOtherBase });
}
if (refusals.length) fail(refusals);

// 4. The bundle.
const manifest = fs.readFileSync(path.join(dirs.before, 'appsscript.json'), 'utf8');
const sortDeep = (v) => Array.isArray(v) ? v.map(sortDeep) : v && typeof v === 'object' ? Object.fromEntries(Object.keys(v).sort().map((k) => [k, sortDeep(v[k])])) : v;
if (JSON.stringify(sortDeep(JSON.parse(manifest))) !== JSON.stringify(sortDeep(JSON.parse(fs.readFileSync(path.join(REPO, 'appsscript.json'), 'utf8'))))) {
  fail('The copy\'s appsscript.json differs from the repository\'s (not only in formatting).');
}
const builtAt = new Date().toISOString();
const bundle = restoreDir ? Object.fromEntries(['appsscript.json', 'Code.js', 'HyperLink.js', 'Release.js', 'ReportCreator.js'].map((n) => {
  const p = path.join(restoreDir, n);
  if (!fs.existsSync(p)) fail('The restore folder has no ' + n + '.');
  return [n, fs.readFileSync(p)];
})) : {
  'Code.js': fs.readFileSync(path.join(REPO, 'Code.js')),
  'HyperLink.js': fs.readFileSync(path.join(REPO, 'HyperLink.js')),
  'ReportCreator.js': fs.readFileSync(path.join(REPO, 'template', 'ReportCreator.js')),
  'appsscript.json': Buffer.from(manifest),
  'Release.js': Buffer.from(candidateReleaseJs(fs.readFileSync(path.join(dirs.before, 'Release.js'), 'utf8'), commit, builtAt)),
  'IntegrationCheck.js': fs.readFileSync(path.join(HERE, 'IntegrationCheck.js'))
};
Object.entries(bundle).forEach(([n, b]) => fs.writeFileSync(path.join(dirs.bundle, n), b));
const planned = hashes(dirs.bundle);

const evidence = {
  action: (restoreDir ? 'restore' : 'push') + (write ? '' : ' (dry run)'), mode: audioMode ? 'audio-original' : 'disposable-copy', at: builtAt, branch, commit,
  restoreFrom: restoreDir, localScriptBackup: localBackup, backupCopy: audioMode && !restoreDir ? auditBackup : null,
  target: { scriptId, documentId, parentId: project.parentId, scriptTitle: project.title, scriptUpdateTime: project.updateTime,
    documentName: file.name, documentCreated: file.createdTime, documentModified: file.modifiedTime, documentVersion: file.version },
  before, planned, readback: null, readbackMatches: null
};

console.log('Target: script ' + scriptId + ' bound to "' + file.name + '" (' + documentId + ')');
console.log('Backup of the script as it was: ' + dirs.before);
console.log((restoreDir ? 'Restore of ' + restoreDir : 'Candidate ' + branch + ' @ ' + commit.slice(0, 7)) + ' bundle: ' + dirs.bundle);
if (audioMode && !restoreDir) console.log('Backup copy: "' + auditBackup.name + '" (' + auditBackup.documentId + '), created ' + auditBackup.createdTime + ', script ' + backupScriptId + ' (T-2026.10.22)');
Object.entries(planned).forEach(([n, h]) => console.log('  ' + n.padEnd(20) + h + (before[n] === h ? '  (unchanged)' : before[n] ? '  (changed)' : '  (new)')));

if (write) {
  // 5. Push (never --force), then read back.
  console.log(clasp(dirs.bundle, 'push').trim());
  clasp(dirs.readback, 'pull');
  evidence.readback = hashes(dirs.readback);
  evidence.readbackMatches = JSON.stringify(evidence.readback) === JSON.stringify(planned);
  console.log('Read-back ' + (evidence.readbackMatches ? 'MATCHES' : 'DOES NOT MATCH') + ' the bundle (SHA-256 of every file).');
} else {
  console.log('\nDry run: nothing was pushed. Add --write to push into the disposable copy.');
}
fs.writeFileSync(path.join(OUT, 'evidence.json'), JSON.stringify(evidence, null, 2) + '\n');
console.log('Evidence: ' + path.join(OUT, 'evidence.json'));
if (write && !evidence.readbackMatches) process.exit(1);
