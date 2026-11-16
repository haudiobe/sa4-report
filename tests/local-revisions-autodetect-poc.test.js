/**
 * PoC (research/local-revision-autodetect): the report side of the local listing agent.
 * tools/local-revisions-poc/LocalListingDriveImport.js reads the listing the agent
 * (sync-local-listings.ps1) wrote to Drive and imports it with the existing
 * saveLocalRevisionsListing(): the result must be exactly that of pasting the same page.
 *
 * The listing is synthetic, in the form the meeting-room server (IIS 10) serves:
 * absolute hrefs, a breadcrumb of folder links. Run: node tests/local-revisions-autodetect-poc.test.js
 */
const fs = require('fs');
const path = require('path');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`); else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const POC = fs.readFileSync(path.join(__dirname, '..', 'tools', 'local-revisions-poc', 'LocalListingDriveImport.js'), 'utf8');
const FOLDER = 'http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/MBS/';
const A = 'S4-269101', B = 'S4-269102';
const IIS = '<!DOCTYPE html><html><head><title>Directory Listing /ftp/SA/SA4/Inbox/Drafts/MBS/</title></head><body>' +
  ['', 'ftp/', 'ftp/SA/', 'ftp/SA/SA4/', 'ftp/SA/SA4/Inbox/', 'ftp/SA/SA4/Inbox/Drafts/'].map((p) => '<a href="http://10.10.10.10/' + p + '">x</a>').join(' / ') +
  '<table>' + [A + '_ExampleCorp.docx', A + 'r01.docx', B + '_Sample.zip', 'S4-260001_other.docx'].map((f) => '<tr><td><a href="' + FOLDER + f + '">' + f + '</a></td><td>16.11.2026 10:01</td></tr>').join('') + '</table></body></html>';
const agentFile = (fetched, html) => '<!-- sa4-local-listing folder="' + FOLDER + '" fetched="' + fetched + '" -->\n' + (html || IIS);

function cell(t) { const s = { text: t || '', links: [] }; return { getText: () => s.text, setText: (v) => { s.text = String(v); s.links = []; }, editAsText: () => ({ setLinkUrl: (a, b, u) => { if (u) s.links.push([s.text.slice(a, b + 1), u]); }, setBold() {}, setForegroundColor() {} }), _s: s }; }
function table(id) { const rows = [['TDoc', id], ['Title', 't'], ['Source', ''], ['Contact', ''], ['Agenda Item', '5'], ['E-mail Discussion', ''], ['Revisions', ''], ['Minutes', 'Minutes typed by hand'], ['Disposition', ''], ['Status', 'Agreed']].map((r) => r.map(cell)); return { getType: () => 'TABLE', getNumRows: () => rows.length, getRow: (i) => ({ getNumCells: () => 2, getCell: (k) => rows[i][k] }), getCell: (a, b) => rows[a][b] }; }

/** A disposable main MBS report with tables A and B; `drive` maps file name -> content. */
function report(drive, props) {
  const tables = [table(A), table(B)];
  const l = loadCode({ code: fs.readFileSync(CODE_JS_PATH, 'utf8') + '\n' + POC, documentProperties: Object.assign({ MEETING_TYPE: 'main', REPORT_SUFFIX: 'MBS', MEETING_FOLDER: 'TSGS4_138_Calgary', MEETING_NUMBER: '138', LOCAL_LISTING_DRIVE_FOLDER_ID: 'folder1' }, props || {}) });
  const s = l.sandbox;
  const fetched = [];
  s.DocumentApp.getActiveDocument = () => ({ getName: () => 'DISPOSABLE TEST COPY (PoC)', getId: () => 'doc', getBody: () => ({ getTables: () => tables }) });
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  s.UrlFetchApp = { fetch: (u) => { fetched.push(u); throw new Error('no fetch expected'); } };
  s.DriveApp = { getFolderById: () => ({ getFilesByName: (n) => { let done = !(n in drive); return { hasNext: () => !done, next: () => { done = true; return { getBlob: () => ({ getDataAsString: () => drive[n] }) }; } }; } }) };
  return { s, tables, docProps: l.docProps, fetched, rev: (i) => tables[i].getCell(6, 1)._s, minutes: () => tables.map((t) => t.getCell(7, 1).getText()) };
}
const now = () => new Date().toISOString();

console.log('1. the agent file');
{
  const { s } = report({});
  check('header read: folder, time, the page after it', [s.pocParseAgentListing_(agentFile('2026-11-16T22:00:00Z')).folder, s.pocParseAgentListing_(agentFile('2026-11-16T22:00:00Z')).fetchedAt, s.pocParseAgentListing_(agentFile('x')).html === IIS], [FOLDER, '2026-11-16T22:00:00Z', true]);
  check('a file without the agent header is not one', s.pocParseAgentListing_(IIS), null);
  const at = new Date('2026-11-16T22:30:00Z');
  check('fresh within 30 minutes; older, unreadable or in the future is not', ['2026-11-16T22:10:00Z', '2026-11-16T21:59:00Z', 'garbage', '2026-11-16T23:00:00Z'].map((t) => s.pocAgentListingFresh_(t, at, 30)), [true, false, false, false]);
}

console.log('2. the import: the existing code, the same result as a paste');
{
  const viaAgent = report({ 'MBS.html': agentFile(now()) });
  const res = viaAgent.s.pocImportLocalListingFromDrive();
  const viaPaste = report({});
  viaPaste.s.saveLocalRevisionsListing({ folder: FOLDER, html: IIS, text: '' });
  check('imported: 4 files, all revisions of SA4 TDocs, 3 new for this report', [res.ok, / 4 file\(s\), 4 revision\(s\) of SA4 TDocs, 3 new for this report\./.test(res.message)], [true, true]);
  check('the Revisions cells are exactly those of pasting the same page (text and links)', [viaAgent.rev(0), viaAgent.rev(1)], [viaPaste.rev(0), viaPaste.rev(1)]);
  check('A: both local files, each linked to its URL on the meeting-room server', viaAgent.rev(0).links.slice(0, 2), [[A + '_ExampleCorp.docx', FOLDER + A + '_ExampleCorp.docx'], ['[local only]', FOLDER + A + '_ExampleCorp.docx']]);
  check('minutes and statuses untouched; nothing fetched (the agent did the reading)', [viaAgent.minutes(), viaAgent.fetched], [['Minutes typed by hand', 'Minutes typed by hand'], []]);
  const again = viaAgent.s.pocImportLocalListingFromDrive();
  check('the same listing again: no new revisions', /no new revisions for this report\./.test(again.message), true);
}

console.log('3. the meeting-room network or the agent unavailable: nothing changes, and it is said');
{
  const stale = report({ 'MBS.html': agentFile('2026-11-16T08:00:00Z') });
  const r1 = stale.s.pocImportLocalListingFromDrive();
  check('a listing older than 30 minutes is not imported', [r1.ok, /^Not imported: the newest listing of MBS is from 2026-11-16T08:00:00Z/.test(r1.message), stale.rev(0).text], [false, true, '']);
  const none = report({});
  check('no listing yet: not checked', none.s.pocImportLocalListingFromDrive().message, 'Not checked: the agent has written no listing for MBS yet.');
  const foreign = report({ 'MBS.html': IIS });
  check('a file not written by the agent is not imported', foreign.s.pocImportLocalListingFromDrive().ok, false);
  const other = report({ 'MBS.html': agentFile(now()).replace(FOLDER, 'https://evil.example/MBS/') });
  check('a listing of another server is refused by the existing check', /not on the meeting-room server/.test(other.s.pocImportLocalListingFromDrive().message), true);
  const video = report({ 'MBS.html': agentFile(now()) }, { REPORT_SUFFIX: 'Video' });
  check('each report reads its own group: a Video report does not take the MBS listing', video.s.pocImportLocalListingFromDrive().message, 'Not checked: the agent has written no listing for Video yet.');
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll local revisions autodetect PoC checks passed.');
process.exit(failures ? 1 : 0);
