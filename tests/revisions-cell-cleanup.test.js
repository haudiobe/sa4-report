/**
 * Revisions-cell cleanup after T-2026.10.28 (the REAL updateRevisions_() on a fake document):
 *
 *   1. no "Revisions folders: ..." line any more; one from an earlier release goes at the next write,
 *      revision lines and their links stay;
 *   2. "No revisions available." only where it is known: every source read and none found;
 *      never for a TDoc that has revisions as TDocs (REVISION_MAP) -- an old one is removed;
 *      not newly claimed while the local listing is stale (an earlier one stays);
 *   3. Check Connections: no "not checked yet. Import Local Inbox Listing..." while the local
 *      listing agent delivers valid scans;
 *   4. all five groups, zero revisions, existing revisions, stale scans, repeat imports,
 *      text typed by hand, minutes untouched; 5. what changed (source).
 *
 * All values are synthetic. Run: node tests/revisions-cell-cleanup.test.js
 */
const fs = require('fs');
const crypto = require('crypto');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { codeWithoutRevisionsCellCleanup, CELL_CLEANUP_NEW_FUNCTIONS, CELL_CLEANUP_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`); else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const LOCAL = 'http://10.10.10.10/ftp/SA/SA4/Inbox';
const NONE = 'No revisions available.';
const OLD_FOLDERS = 'Revisions folders: Local – meeting-room server, may be newer, on-site only · Remote – 3GPP server, synchronized, accessible externally';
const A = 'S4-269101', B = 'S4-269102', C = 'S4-269103';
const sha = (t) => crypto.createHash('sha256').update(String(t), 'utf8').digest('hex');
const ago = (m) => new Date(Date.now() - m * 60000).toISOString().replace(/\.\d+Z$/, 'Z');
const GROUPS = { MBS: 'MBS', Video: 'Video', '6G': 'FS_6G_MED', RTC: 'RTC', Audio: 'Audio' };

/** An agent file of `group` listing `files`. */
function listing(group, files, fetchedAt) {
  const folder = LOCAL + '/Drafts/' + group + '/';
  const payload = files.map((f) => '<a href="' + folder + encodeURIComponent(f) + '">' + f + '</a>').join('\n') || '<a href="' + LOCAL + '/Drafts/">Drafts</a>';
  return JSON.stringify({ format: 'sa4-local-listing', version: 1, group: group, folder: folder, fetchedAt: fetchedAt || ago(2), scanId: sha(payload), agent: 'ROOM', linkCount: (payload.match(/<a\s/g) || []).length, sha256: sha(payload), payload: payload });
}

function makeCell(text) {
  const st = { text: String(text || ''), links: [], writes: 0 };
  return { getText: () => st.text, setText: (t) => { st.text = String(t); st.links = []; st.writes++; },
    editAsText: () => ({ setLinkUrl: (a, b, u) => { st.links = st.links.filter((l) => l.end < a || l.start > b); if (u) st.links.push({ start: a, end: b, url: u }); }, setBold() {}, setForegroundColor() {} }), _st: st };
}
function makeTable(id, revisions) {
  const rows = [['TDoc', id], ['Title', 't'], ['Source', ''], ['Contact', ''], ['Agenda Item', '5.1'], ['E-mail Discussion', ''], ['Revisions', revisions || ''], ['Minutes', 'Minutes typed for ' + id], ['Disposition', ''], ['Status', 'Agreed']].map((r) => r.map(makeCell));
  return { getType: () => 'TABLE', getNumRows: () => rows.length, getRow: (i) => ({ getNumCells: () => 2, getCell: (k) => rows[i][k] }), getCell: (a, b) => rows[a][b], appendTableRow() {}, removeRow() {} };
}

/** A report of `suffix` with tables A, B, C (their Revisions cells as given), the agent folder `drive`, the 3GPP folder `remote` (file names). */
function report(suffix, cells, drive, remote, props) {
  const tables = [A, B, C].map((id, i) => makeTable(id, (cells || [])[i]));
  const l = loadCode({ documentProperties: Object.assign({ MEETING_TYPE: 'main', REPORT_SUFFIX: suffix, MEETING_FOLDER: 'TSGS4_138_Calgary', MEETING_NUMBER: '138', LOCAL_LISTING_DRIVE_FOLDER_ID: 'DriveFolder0001' }, props || {}) });
  const s = l.sandbox;
  const body = { getTables: () => tables, getNumChildren: () => tables.length, getChild: (i) => tables[i], getChildIndex: (t) => tables.indexOf(t), insertTable() { throw new Error('no table may be inserted'); } };
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'doc' });
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  s.Utilities.computeDigest = (alg, text) => Array.from(crypto.createHash('sha256').update(String(text), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b));
  s.Utilities.DigestAlgorithm = { SHA_256: 'SHA_256' }; s.Utilities.Charset = { UTF_8: 'UTF_8' };
  const r = { s, tables, docProps: l.docProps, drive: drive || {}, remote: remote || [] };
  s.UrlFetchApp = { fetch: (u) => { if (/10\.10\.10\.10/.test(u)) throw new Error('never'); return { getResponseCode: () => 200, getContentText: () => '<html>' + r.remote.map((f) => '<a href="' + f + '">' + f + '</a>').join('') + '</html>' }; } };
  s.DriveApp = { getFolderById: () => ({ getFilesByName: (n) => { const raw = r.drive[n]; let done = raw === undefined; return { hasNext: () => !done, next: () => { done = true; return { getBlob: () => ({ getDataAsString: () => raw }), getLastUpdated: () => new Date() }; } }; } }) };
  r.update = () => s.updateRevisions_(s.getCollectorConfig_());
  r.cell = (i) => tables[i].getCell(6, 1)._st;
  r.texts = () => [0, 1, 2].map((i) => r.cell(i).text);
  r.writes = () => [0, 1, 2].map((i) => r.cell(i).writes);
  r.minutes = () => tables.map((t) => t.getCell(7, 1).getText());
  return r;
}

console.log('1. all five groups: own revisions, no folders line, "No revisions available." where every source has none');
Object.keys(GROUPS).forEach((suffix) => {
  const g = GROUPS[suffix];
  const r = report(suffix, [], { [g + '.listing.json']: listing(g, [A + '_' + g + '.docx']) }, [B + '_r1.docx']);
  r.update();
  check(suffix + ': A local only, B remote only, C none -- and no folders line anywhere', r.texts(), [A + '_' + g + '.docx [local only]', B + '_r1.docx', NONE]);
  check(suffix + ': the revision links stay (local to the meeting-room server, remote to 3GPP)', [r.cell(0).links[0].url, /^https:\/\/www\.3gpp\.org\/.*\/Drafts\//.test(r.cell(1).links[0].url)], [LOCAL + '/Drafts/' + g + '/' + A + '_' + g + '.docx', true]);
});

console.log('2. zero revisions');
{
  const r = report('MBS', [], { 'MBS.listing.json': listing('MBS', []) }, []);
  r.update();
  check('a fresh scan and the 3GPP folder, both without revisions: "No revisions available." -- truly empty', r.texts(), [NONE, NONE, NONE]);
  check('an empty group folder is an answer, not a failure: imported, 0 files', [JSON.parse(r.docProps.getProperty('LOCAL_LISTING_AGENT_STATE')).lastResult, JSON.parse(r.docProps.getProperty('LOCAL_LISTING_AGENT_STATE')).files], ['imported', 0]);
}

console.log('3. cells written by T-2026.10.28 (with the folders line)');
{
  const r = report('MBS', [A + '_r1.docx [local only]\n' + OLD_FOLDERS, NONE + '\n' + OLD_FOLDERS, 'Typed by hand: see the chair\'s notes\n' + OLD_FOLDERS],
    { 'MBS.listing.json': listing('MBS', [A + '_r1.docx']) }, []);
  r.update();
  check('the folders line goes; revisions, "No revisions available." where true, and typed text stay', r.texts(), [A + '_r1.docx [local only]', NONE, 'Typed by hand: see the chair\'s notes']);
  check('the revision keeps its link', r.cell(0).links.map((l) => l.url), [LOCAL + '/Drafts/MBS/' + A + '_r1.docx', LOCAL + '/Drafts/MBS/' + A + '_r1.docx']);
  check('minutes untouched', r.minutes(), ['Minutes typed for ' + A, 'Minutes typed for ' + B, 'Minutes typed for ' + C]);
}

console.log('4. TDocs that have revisions as TDocs (REVISION_MAP)');
{
  const r = report('MBS', [NONE, '', NONE], { 'MBS.listing.json': listing('MBS', []) }, [], { REVISION_MAP: JSON.stringify({ [A]: 'S4-269199', 'S4-269198': B }) });
  r.update();
  check('A (revised to another TDoc) and B (another revised into it): no "No revisions available." -- an old one removed; C: none', r.texts(), ['', '', NONE]);
}

console.log('5. a stale local listing');
{
  const r = report('MBS', [NONE, '', ''], { 'MBS.listing.json': listing('MBS', [A + '_r1.docx'], ago(45)) }, [C + '_r1.docx']);
  r.update();
  check('nothing newly claimed while the local source is unread: B stays empty; A keeps its earlier sentence; C shows its 3GPP revision', r.texts(), [NONE, '', C + '_r1.docx']);
  const fresh = report('MBS', [], { 'MBS.listing.json': listing('MBS', [A + '_r1.docx']) }, []);
  fresh.update();
  fresh.drive = { 'MBS.listing.json': listing('MBS', [A + '_r1.docx'], ago(50)) };
  fresh.update();
  check('imported revisions are kept when the listing goes stale; no sentence appears', fresh.texts(), [A + '_r1.docx [local only]', NONE, NONE]);
  const missing = report('MBS', ['', '', ''], {}, []);
  missing.update();
  check('no listing at all (agent configured): nothing claimed', missing.texts(), ['', '', '']);
  const manualOnly = report('MBS', ['', '', ''], {}, [], { LOCAL_LISTING_DRIVE_FOLDER_ID: null });
  manualOnly.update();
  check('no agent configured (manual import only): as before -- every source the report reads was read', manualOnly.texts(), [NONE, NONE, NONE]);
}

console.log('6. repeat imports');
{
  const r = report('Video', [], { 'Video.listing.json': listing('Video', [A + '_v.docx']) }, [B + '_r1.docx']);
  r.update();
  const texts = r.texts();
  const writes = r.writes();
  r.update();
  r.update();
  check('the same scan and folder twice more: the same cells, "No revisions available." not rewritten', [r.texts(), r.writes()[2]], [texts, writes[2]]);
  check('no line twice', r.texts().every((t) => new Set(t.split('\n')).size === t.split('\n').length), true);
}

console.log('7. Check Connections');
{
  const { s } = report('MBS', [], {}, []);
  const agentOk = JSON.stringify({ lastScanFetchedAt: ago(2), lastResult: 'unchanged' });
  check('a valid agent scan, no manual import: no "not checked yet. Import Local Inbox Listing..." line', s.describeLocalRevisionsStatus_(null, LOCAL, agentOk), '');
  check('the agent stale, missing or not configured: the instruction stays', [JSON.stringify({ lastScanFetchedAt: ago(40), lastResult: 'stale' }), JSON.stringify({ lastResult: 'missing' }), null].map((st) => /not checked yet\. Import Local Inbox Listing/.test(s.describeLocalRevisionsStatus_(null, LOCAL, st))), [true, true, true]);
  check('a manual import was made: its line stays', /^Local \(meeting-room server\): last listing imported 2026-11-17 10:00/.test(s.describeLocalRevisionsStatus_(JSON.stringify({ at: '2026-11-17 10:00', folder: LOCAL + '/Drafts/MBS/', listed: 3, recognized: 2, added: 1 }), LOCAL, agentOk)), true);
  const r = report('MBS', [], { 'MBS.listing.json': listing('MBS', [A + '_r1.docx']) }, []);
  r.update();
  check('after a real agent import the report\'s own state suppresses it', s.describeLocalRevisionsStatus_(null, LOCAL, r.docProps.getProperty('LOCAL_LISTING_AGENT_STATE')), '');
}

console.log('8. what changed in Code.js (source)');
{
  const now = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutRevisionsCellCleanup(now);
  check('the take-out finds every change exactly once', view.problems, []);
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (code, name) => { const at = code.indexOf('\nfunction ' + name + '('); return code.slice(at, code.indexOf('\n}\n', at)); };
  const before = names(view.code);
  check('the new functions', names(now).filter((n) => before.indexOf(n) === -1).sort(), CELL_CLEANUP_NEW_FUNCTIONS.slice().sort());
  check('the changed functions', before.filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)).sort(), CELL_CLEANUP_CHANGED_FUNCTIONS.slice().sort());
  check('untouched: revision matching, links, merge, agent processing, manual import, group isolation',
    ['parseAnchors_', 'buildRevisionAnchorIndex_', 'parseDraftAnchor_', 'normalizeRevisionKey_', 'revisionCellEntry_', 'mergeLocalRevisions_', 'storeLocalRevisions_', 'parseLocalRevisionsListing_', 'saveLocalRevisionsListing',
      'validateAgentListing_', 'readAgentListingFiles_', 'insertRevisedDocTablesAfter_'].filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll revisions-cell cleanup checks passed.');
process.exit(failures ? 1 : 0);
