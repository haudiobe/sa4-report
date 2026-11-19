/**
 * The local listing agent and the Docs folder of the meeting-room server (T-2026.10.32;
 * docs/UNDISPOSED_AND_LOCAL_DOCS.md § B): with "docsUrl" in its configuration, every run also lists
 * that folder as the group "Docs", with the checks of the groups' folders, into Docs.listing.json;
 * the report reads that file with its own checks and takes the TDocs named exactly.
 *
 *   1. the script: the Docs scan is optional, uses the same scan function, never another one;
 *   2. a real run against a local HTTP server (127.0.0.1, synthetic listings): Docs.listing.json
 *      beside the groups', valid for a report (validateAgentListing_), the exact-TDoc rule;
 *      -Group Docs alone; -Group Docs without "docsUrl" refused; without "docsUrl" no Docs file;
 *      the server failing for Docs: its status says so, the last good listing is kept, the groups
 *      are still written.
 *
 * Part 2 needs Windows (powershell.exe); elsewhere it is skipped. The agent's log goes to a temp
 * folder (LOCALAPPDATA), never to the log of an installed agent. Nothing but 127.0.0.1 is contacted.
 * Run: node tests/local-listing-agent-docs.test.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { loadCode } = require('./helpers/load-code.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`); else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const AGENT = path.join(__dirname, '..', 'tools', 'local-listing-agent', 'sa4-local-listing-agent.ps1');
const TEXT = fs.readFileSync(AGENT, 'utf8').replace(/\r/g, '');

console.log('1. the script');
check('"docsUrl" is optional: without it, nothing is scanned beyond the groups', /\$scanDocs = \[bool\]\$docsUrl/.test(TEXT), true);
check('the Docs folder goes through the same scan function as a group (same checks, same atomic files)', /Invoke-GroupScan 'Docs' \$docsUrl/.test(TEXT) && (TEXT.match(/^function Invoke-GroupScan/gm) || []).length === 1, true);
check('-Group Docs without "docsUrl": refused (exit 2)', /-Group Docs needs "docsUrl" in the configuration'; exit 2/.test(TEXT), true);

if (process.platform !== 'win32') {
  console.log('\n(part 2 skipped: it needs Windows PowerShell)');
  process.exit(failures ? 1 : 0);
}

const PAGE = (p, names) => '<html><head><title>' + p + '</title></head><body><h1>' + p + '</h1><pre><a href="../">[To Parent Directory]</a>\n' +
  names.map((n) => '11/19/2026  9:00 AM   12345 <a href="' + p + n + '">' + n + '</a>').join('\n') + '</pre></body></html>';
const DOCS_FILES = ['S4-261801.zip', 'S4-261802.zip', 'S4-261803_r1.zip', 'S4-2618040.zip'];
let docsDown = false;
const server = http.createServer((req, res) => {
  const p = decodeURIComponent(req.url);
  if (p === '/ftp/SA/SA4/Docs/') {
    if (docsDown) { res.writeHead(500, { 'Content-Type': 'text/html' }); res.end('<html>error</html>'); return; }
    res.writeHead(200, { 'Content-Type': 'text/html' }); res.end(PAGE(p, DOCS_FILES)); return;
  }
  if (p === '/ftp/SA/SA4/Inbox/Drafts/MBS/') { res.writeHead(200, { 'Content-Type': 'text/html' }); res.end(PAGE(p, ['S4-261801-r1.docx'])); return; }
  res.writeHead(404, { 'Content-Type': 'text/html' }); res.end('<html>not found</html>');
});

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'sa4-docs-'));
function run(cfg, args) {
  const cfgPath = path.join(tmp, 'agent.config.json');
  fs.writeFileSync(cfgPath, JSON.stringify(cfg));
  const once = () => new Promise((resolve) => {
    const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', AGENT, '-ConfigPath', cfgPath, '-IgnoreWindow'].concat(args || []),
      { env: Object.assign({}, process.env, { LOCALAPPDATA: tmp }) });
    let out = '';
    child.stdout.on('data', (d) => { out += d; });
    child.stderr.on('data', (d) => { out += d; });
    child.on('close', (code) => resolve({ code: code, out: out }));
  });
  // The scan mutex is shared with an installed agent on this laptop: a run it skips is tried again.
  const attempt = async (left) => {
    const r = await once();
    if (/another scan is still running; skipped/.test(r.out) && left > 0) { await new Promise((ok) => setTimeout(ok, 5000)); return attempt(left - 1); }
    return r;
  };
  return attempt(24);
}
const read = (dir, name) => { const p = path.join(dir, name); return fs.existsSync(p) ? JSON.parse(fs.readFileSync(p, 'utf8').replace(/^﻿/, '')) : null; };

server.listen(0, '127.0.0.1', async () => {
  try {
    const port = server.address().port;
    const origin = 'http://127.0.0.1:' + port;
    const DOCS = origin + '/ftp/SA/SA4/Docs/';
    const base = { baseUrl: origin + '/ftp/SA/SA4/Inbox/Drafts/', groups: ['MBS'], timeoutSeconds: 5 };

    console.log('2. a real run against a local server');
    const outA = path.join(tmp, 'a'); fs.mkdirSync(outA);
    const r1 = await run(Object.assign({}, base, { outDir: outA, docsUrl: DOCS }));
    const docs = read(outA, 'Docs.listing.json');
    if (r1.code !== 0) console.log(r1.out);
    check('a run with "docsUrl": exit 0; MBS and Docs listed', [r1.code, !!read(outA, 'MBS.listing.json'), !!docs], [0, true, true]);
    check('Docs.listing.json: group Docs, the Docs folder, the links of the page', docs && [docs.format, docs.group, docs.folder, docs.linkCount], ['sa4-local-listing', 'Docs', DOCS, 5]);

    // The report's side: the checks it makes, and the TDocs it takes.
    const { sandbox: s } = loadCode({ documentProperties: {} });
    const sha = (t) => crypto.createHash('sha256').update(String(t), 'utf8').digest('hex');
    const localBase = origin + '/ftp/SA/SA4/Inbox/Drafts/MBS/';
    const v = s.validateAgentListing_(fs.readFileSync(path.join(outA, 'Docs.listing.json'), 'utf8').replace(/^﻿/, ''), { group: 'Docs', localBase: localBase }, new Date(), sha);
    check('a report accepts it (complete, fresh, of the Docs folder of its meeting-room server)', [v.ok, v.reason], [true, '']);
    check('the Docs folder the report expects is the one listed', s.localDocsFolderOf_(localBase), DOCS);
    check('the TDocs it takes: exactly named only (not _r1, not seven digits)', Object.keys(s.localDocsPresentIn_(v.listing.payload, DOCS).present).sort(), ['S4-261801', 'S4-261802']);

    const outB = path.join(tmp, 'b'); fs.mkdirSync(outB);
    const r2 = await run(Object.assign({}, base, { outDir: outB, docsUrl: DOCS }), ['-Group', 'Docs']);
    check('-Group Docs: only the Docs folder', [r2.code, !!read(outB, 'Docs.listing.json'), !!read(outB, 'MBS.listing.json')], [0, true, false]);
    const outC = path.join(tmp, 'c'); fs.mkdirSync(outC);
    const r3 = await run(Object.assign({}, base, { outDir: outC }), ['-Group', 'Docs']);
    check('-Group Docs without "docsUrl": refused, exit 2, nothing written', [r3.code, fs.readdirSync(outC).length], [2, 0]);
    const r4 = await run(Object.assign({}, base, { outDir: outC }));
    check('without "docsUrl": as before -- the groups only, no Docs file', [r4.code, !!read(outC, 'MBS.listing.json'), !!read(outC, 'Docs.listing.json')], [0, true, false]);

    docsDown = true;
    const before = fs.readFileSync(path.join(outA, 'Docs.listing.json'), 'utf8');
    const r5 = await run(Object.assign({}, base, { outDir: outA, docsUrl: DOCS }));
    const st = read(outA, 'Docs.status.json');
    check('the Docs folder failing: exit 1, its status says why, the last good listing kept, MBS still written',
      [r5.code, st && st.ok, st && st.httpStatus, fs.readFileSync(path.join(outA, 'Docs.listing.json'), 'utf8') === before, read(outA, 'MBS.status.json').ok], [1, false, 500, true, true]);
  } catch (e) {
    failures++;
    console.log('  FAIL ' + e.stack);
  } finally {
    server.close();
    try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) { /* a temp folder */ }
    console.log(failures ? `\n${failures} check(s) FAILED` : '\nall checks passed');
    process.exit(failures ? 1 : 0);
  }
});
