/**
 * The local listing agent (after T-2026.10.27; docs/LOCAL_LISTING_AGENT.md): every update imports
 * the listing the agent wrote to Drive for the report's own group, with the existing import code.
 *
 *   1. configuration: folder id or URL; the dialog field; central-state and adoption keys
 *   2. validation: complete, checksum, link count, group, folder, server, age, clock
 *   3. import: the same cells as a paste; idempotent; minutes untouched; nothing fetched
 *   4. failures never erase: missing, stale, incomplete, foreign, Drive unreadable
 *   5. Drive: duplicate files, a newer incomplete upload, synchronization delay
 *   6. recovery, refresh, duplicate revisions, coexistence with the manual import
 *   7. through the update: before the 3GPP folder, also when that one fails
 *   8. status lines; 9. what changed in Code.js
 *
 * All values are synthetic. Run: node tests/local-listing-agent.test.js
 */
const fs = require('fs');
const crypto = require('crypto');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { codeWithoutLocalListingAgent, LOCAL_LISTING_AGENT_NEW_FUNCTIONS, LOCAL_LISTING_AGENT_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`); else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const LOCAL = 'http://10.10.10.10/ftp/SA/SA4/Inbox';
const FOLDER = LOCAL + '/Drafts/MBS/';
const A = 'S4-269101', B = 'S4-269102', C = 'S4-269103';
const MAIN = { MEETING_TYPE: 'main', REPORT_SUFFIX: 'MBS', MEETING_FOLDER: 'TSGS4_138_Calgary', MEETING_NUMBER: '138', LOCAL_LISTING_DRIVE_FOLDER_ID: 'DriveFolder0001' };
const sha = (t) => crypto.createHash('sha256').update(String(t), 'utf8').digest('hex');
const NOW = new Date('2026-11-17T10:00:00Z');
const minutesAgo = (m) => new Date(NOW.getTime() - m * 60000).toISOString().replace(/\.\d+Z$/, 'Z');

/** An agent file, as sa4-local-listing-agent.ps1 writes it (overrides for broken ones). */
function listing(files, over) {
  const folder = (over && over.folder) || FOLDER;
  const links = ['<a href="http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/">Drafts</a>'].concat(files.map((f) => '<a href="' + folder + encodeURIComponent(f) + '">' + f + '</a>'));
  const payload = links.join('\n');
  const f = Object.assign({ format: 'sa4-local-listing', version: 1, group: 'MBS', folder: folder, fetchedAt: minutesAgo(2), scanId: sha(payload), agent: 'ROOM-LAPTOP', linkCount: links.length, sha256: sha(payload), payload: payload }, over || {});
  return JSON.stringify(f, null, 2);
}

// ---- fake document and Drive
function makeCell(text) {
  const st = { text: String(text || ''), links: [], writes: 0 };
  return { getText: () => st.text, setText: (t) => { st.text = String(t); st.links = []; st.writes++; },
    editAsText: () => ({ setLinkUrl: (a, b, u) => { st.links = st.links.filter((l) => l.end < a || l.start > b); if (u) st.links.push({ start: a, end: b, url: u }); }, setBold() {}, setForegroundColor() {} }), _st: st };
}
function makeTable(id) {
  const rows = [['TDoc', id], ['Title', 't'], ['Source', ''], ['Contact', ''], ['Agenda Item', '5.1'], ['E-mail Discussion', ''], ['Revisions', ''], ['Minutes', 'Minutes typed for ' + id], ['Disposition', ''], ['Status', 'Agreed']].map((r) => r.map(makeCell));
  return { getType: () => 'TABLE', getNumRows: () => rows.length, getRow: (i) => ({ getNumCells: () => 2, getCell: (k) => rows[i][k] }), getCell: (a, b) => rows[a][b], appendTableRow() {}, removeRow() {} };
}

/** A report: the REAL Code.js; `drive` is { name: [ { raw, updated } ] } or a function that throws. */
function report(drive, props) {
  const tables = [makeTable(A), makeTable(B), makeTable(C)];
  const l = loadCode({ documentProperties: Object.assign({}, MAIN, props || {}) });
  const s = l.sandbox;
  const body = { getTables: () => tables, getNumChildren: () => tables.length, getChild: (i) => tables[i], getChildIndex: (t) => tables.indexOf(t), insertTable() { throw new Error('no table may be inserted'); } };
  const r = { s, tables, docProps: l.docProps, fetched: [], driveReads: 0, writes: [], drive: drive, remoteHtml: '<html></html>' };
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'doc', getName: () => 'MBS SWG Minutes SA4#138' });
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  s.Utilities.computeDigest = (alg, text) => Array.from(crypto.createHash('sha256').update(String(text), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b));
  s.Utilities.DigestAlgorithm = { SHA_256: 'SHA_256' };
  s.Utilities.Charset = { UTF_8: 'UTF_8' };
  s.UrlFetchApp = { fetch: (u) => { r.fetched.push(u); if (/10\.10\.10\.10/.test(u)) throw new Error('the meeting-room server must never be fetched'); if (r.remoteHtml === null) return { getResponseCode: () => 503, getContentText: () => '' }; return { getResponseCode: () => 200, getContentText: () => r.remoteHtml }; } };
  s.DriveApp = { getFolderById: (id) => {
    r.driveReads++;
    if (typeof r.drive === 'function') r.drive(id);
    return { getFilesByName: (name) => { const list = ((r.drive || {})[name] || []).slice(); let i = 0;
      return { hasNext: () => i < list.length, next: () => { const f = list[i++]; return { getBlob: () => ({ getDataAsString: () => f.raw }), getLastUpdated: () => new Date(f.updated || NOW) }; } }; } };
  } };
  const set = l.docProps.setProperty;
  l.docProps.setProperty = (k, v) => { r.writes.push(k); return set(k, v); };
  r.importNow = () => s.importAgentListing_(undefined, { now: NOW });
  r.cell = (i) => tables[i].getCell(6, 1)._st;
  r.lines = (i) => r.cell(i).text.split('\n').filter((x) => x.indexOf('Revisions folders: ') !== 0);
  r.state = () => JSON.parse(l.docProps.getProperty('LOCAL_LISTING_AGENT_STATE') || '{}');
  r.stored = (id) => JSON.parse(l.docProps.getProperty('LOCALREV_' + id) || 'null');
  r.minutes = () => tables.map((t) => t.getCell(7, 1).getText());
  return r;
}
const file = (raw, updated) => [{ raw: raw, updated: updated || NOW }];

// ====================================================================== 1
console.log('1. configuration');
{
  const { s } = report({});
  check('a folder id, a folder URL, an open?id URL; anything else is no folder', ['1AbCdEfGhIjKlMnOp', 'https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOp?usp=sharing', 'https://drive.google.com/drive/u/1/folders/1AbCdEfGhIjKlMnOp', 'https://drive.google.com/open?id=1AbCdEfGhIjKlMnOp', 'my folder', 'https://evil.example/folders/1AbCdEfGhIjKlMnOp', ''].map((v) => s.localListingFolderIdOf_(v)),
    ['1AbCdEfGhIjKlMnOp', '1AbCdEfGhIjKlMnOp', '1AbCdEfGhIjKlMnOp', '1AbCdEfGhIjKlMnOp', '', '', '']);
  const save = (props, value) => { const l = loadCode({ documentProperties: Object.assign({}, MAIN, props) }); const config = { meetingId: '60780', meetingType: 'main', reportType: 'MBS', mailingListMode: 'derived', apiTokenAction: 'keep' }; if (value !== undefined) config.localListingFolder = value; l.sandbox.persistConfigurationSettings_(config); return l.docProps.getProperty('LOCAL_LISTING_DRIVE_FOLDER_ID'); };
  check('Configure Meeting stores the id of a URL, removes it when blank, ignores an invalid value and an older dialog', [save({}, 'https://drive.google.com/drive/folders/1NewFolderId00'), save({}, ''), save({}, 'not a folder'), save({}, undefined)],
    ['1NewFolderId00', null, 'DriveFolder0001', 'DriveFolder0001']);
  check('it reaches the automatic updates: a Configure Meeting key and an adoption key', [s.CONFIG_DIALOG_MANAGED_KEYS_.indexOf('LOCAL_LISTING_DRIVE_FOLDER_ID') !== -1, s.ADDON003_ADOPTION_FIXED_KEYS_.indexOf('LOCAL_LISTING_DRIVE_FOLDER_ID') !== -1], [true, true]);
  const none = report({}, { LOCAL_LISTING_DRIVE_FOLDER_ID: null });
  check('not configured: nothing is read, nothing is written', [none.importNow(), none.driveReads, none.writes], [null, 0, []]);
}

// ====================================================================== 2
console.log('2. what may be imported');
{
  const { s } = report({});
  const v = (raw, group, now) => s.validateAgentListing_(raw, { group: group || 'MBS', localBase: LOCAL }, now || NOW, sha).reason;
  const good = listing([A + '_r1.docx']);
  check('a complete, fresh listing of the group', v(good), '');
  check('incomplete JSON (still being written or synchronized)', v(good.slice(0, good.length - 40)), 'incomplete or unreadable (not complete JSON: still being written or synchronized?)');
  check('content that does not match its checksum', v(listing([A + '_r1.docx'], { sha256: sha('something else') })), 'incomplete: its content does not match its checksum');
  check('fewer links than announced', v(listing([A + '_r1.docx'], { linkCount: 9 })), 'incomplete: 9 links announced, 2 present');
  check('another group\'s file', v(listing([A + '_r1.docx'], { group: 'Video' })), 'a listing of group Video, not of MBS');
  check('the folder of another group', v(listing([A + '_r1.docx'], { folder: LOCAL + '/Drafts/Video/' })), 'the listed folder http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/Video/ is not the folder of group MBS');
  check('another server', v(listing([A + '_r1.docx'], { folder: 'https://evil.example/Drafts/MBS/' })), 'the listed folder https://evil.example/Drafts/MBS/ is not on the meeting-room server of this report');
  check('31 minutes old: stale', /^stale: scanned .* more than 30 minutes ago/.test(v(listing([A + '_r1.docx'], { fetchedAt: minutesAgo(31) }))), true);
  check('29 minutes old: fresh', v(listing([A + '_r1.docx'], { fetchedAt: minutesAgo(29) })), '');
  check('more than 5 minutes in the future: the agent\'s clock is wrong', /^scanned in the future/.test(v(listing([A + '_r1.docx'], { fetchedAt: minutesAgo(-6) }))), true);
  check('not written by the agent', [v('{"format":"other","version":1}'), v('<html></html>')], ['not a listing written by the local listing agent', 'incomplete or unreadable (not complete JSON: still being written or synchronized?)']);
}

// ====================================================================== 3
console.log('3. the import');
{
  const files = [A + '_r1.docx', A + 'r02_ExampleCorp.docx', B + '_Sample.zip', 'S4-260001_elsewhere.docx'];
  const r = report({ 'MBS.listing.json': file(listing(files)) });
  const res = r.importNow();
  check('imported: what the listing has, for this report\'s TDocs', [res.result, res.state.message], ['imported', 'Imported the scan of ' + minutesAgo(2) + ': 4 file(s), 4 revision(s) of SA4 TDocs, 3 new, 0 no longer listed.']);
  const paste = report({});
  paste.s.saveLocalRevisionsListing({ folder: FOLDER, html: JSON.parse(listing(files)).payload, text: '' });
  check('the Revisions cells are exactly those of Import Local Inbox Listing... with the same listing', [r.cell(0).text, r.cell(0).links, r.cell(1).text], [paste.cell(0).text, paste.cell(0).links, paste.cell(1).text]);
  check('A: both local files, linked to the meeting-room server; C (not listed) untouched', [r.lines(0), r.cell(0).links[0].url, r.cell(2).writes], [[A + '_r1.docx [local only]', A + 'r02_ExampleCorp.docx [local only]'], FOLDER + A + '_r1.docx', 0]);
  check('minutes and statuses untouched; the meeting-room server never fetched', [r.minutes(), r.fetched], [['Minutes typed for ' + A, 'Minutes typed for ' + B, 'Minutes typed for ' + C], []]);
  const writes = r.writes.length, cellWrites = r.cell(0).writes;
  const again = r.importNow();
  check('the same scan again: unchanged -- only the check time is written, no revision, no cell', [again.result, r.writes.slice(writes), r.cell(0).writes], ['unchanged', ['LOCAL_LISTING_AGENT_STATE'], cellWrites]);
}

// ====================================================================== 4
console.log('4. failures never erase what was imported');
{
  const r = report({ 'MBS.listing.json': file(listing([A + '_r1.docx'])) });
  r.importNow();
  const kept = [r.stored(A), r.cell(0).text];
  const cases = [
    ['missing', () => { r.drive = {}; }, 'missing', /^Not checked: no listing of MBS/],
    ['stale', () => { r.drive = { 'MBS.listing.json': file(listing([], { fetchedAt: minutesAgo(45) })) }; }, 'stale', /^Not imported: the newest listing of MBS is stale/],
    ['incomplete', () => { r.drive = { 'MBS.listing.json': file('{"format":"sa4-local-listing","version":1,"group":"MBS","payl') }; }, 'invalid', /incomplete or unreadable/],
    ['foreign group', () => { r.drive = { 'MBS.listing.json': file(listing([], { group: 'RTC' })) }; }, 'invalid', /a listing of group RTC/],
    ['Drive unreadable', () => { r.drive = () => { throw new Error('Access denied: DriveApp'); }; }, 'unavailable', /cannot be read \(Access denied: DriveApp\)/]
  ];
  cases.forEach(([label, arrange, result, message]) => {
    arrange();
    const res = r.importNow();
    check(label + ': ' + result + ', said, and the imported revisions and the cell are kept', [res.result, message.test(res.state.message), /Previously imported revisions are kept\.$|^Not checked/.test(res.state.message), [r.stored(A), r.cell(0).text]], [result, true, true, kept]);
  });
  check('an empty but fresh listing removes nothing either (a wrong or partial page)', (() => { r.drive = { 'MBS.listing.json': file(listing([], { fetchedAt: minutesAgo(1), scanId: 'empty' })) }; r.importNow(); return [r.stored(A), r.cell(0).text]; })(), kept);
}

// ====================================================================== 5
console.log('5. Drive: duplicates and synchronization');
{
  const valid = listing([A + '_r1.docx']);
  const r = report({ 'MBS.listing.json': [{ raw: '{"format":"sa4-local-listing","version":1,"gro', updated: NOW }, { raw: valid, updated: new Date(NOW.getTime() - 120000) }] });
  check('a newer upload still incomplete: the older complete one is used', [r.importNow().result, r.lines(0)], ['imported', [A + '_r1.docx [local only]']]);
  const newer = listing([A + '_r1.docx', A + '_r2.docx'], { fetchedAt: minutesAgo(1) });
  const r2 = report({ 'MBS.listing.json': [{ raw: valid, updated: new Date(NOW.getTime() - 120000) }, { raw: newer, updated: NOW }] });
  check('two complete files (a Drive conflict copy with the same name): the newest is used', [r2.importNow().result, r2.lines(0)], ['imported', [A + '_r1.docx [local only]', A + '_r2.docx [local only]']]);
  const delayed = report({ 'MBS.listing.json': file(listing([A + '_r1.docx'], { fetchedAt: minutesAgo(40) })) });
  check('Drive has not uploaded the agent\'s newer scans: stale, nothing imported, said', [delayed.importNow().result, delayed.stored(A)], ['stale', null]);
}

// ====================================================================== 6
console.log('6. recovery, refresh, duplicates, the manual import');
{
  const r = report({ 'MBS.listing.json': file(listing([A + '_r1.docx'], { fetchedAt: minutesAgo(50) })) });
  check('stale first', r.importNow().result, 'stale');
  r.drive = { 'MBS.listing.json': file(listing([A + '_r1.docx', A + '_r1.docx', B + '_x.docx'])) };
  const res = r.importNow();
  check('then a fresh scan: imported; the same file listed twice is stored once', [res.result, Object.keys(r.stored(A)), r.state().lastResult], ['imported', [(A + '_r1.docx').toLowerCase()], 'imported']);
  r.drive = { 'MBS.listing.json': file(listing([B + '_x.docx', A + '_r3.docx'], { fetchedAt: minutesAgo(1) })) };
  const refreshed = r.importNow();
  check('a newer scan without A_r1 (renamed): removed, as with the manual import; A_r3 added', [refreshed.state.added, refreshed.state.removed, r.lines(0)], [1, 1, [A + '_r3.docx [local only]']]);
  r.s.saveLocalRevisionsListing({ folder: LOCAL + '/Drafts/MBS/Extra/', html: '<a href="' + LOCAL + '/Drafts/MBS/Extra/' + A + '_r9.docx">' + A + '_r9.docx</a>', text: '' });
  check('the manual import still works next to it (another folder: both kept)', r.lines(0), [A + '_r3.docx [local only]', A + '_r9.docx [local only]']);
  r.drive = { 'MBS.listing.json': file(listing([B + '_x.docx', A + '_r3.docx'], { fetchedAt: minutesAgo(0), scanId: 'next' })) };
  r.importNow();
  check('...and the agent\'s next scan does not remove what was imported by hand from another folder', r.lines(0), [A + '_r3.docx [local only]', A + '_r9.docx [local only]']);
}

// ====================================================================== 7
console.log('7. through the update (updateRevisions_)');
{
  const r = report({ 'MBS.listing.json': file(listing([A + '_r1.docx'], { fetchedAt: new Date(Date.now() - 60000).toISOString() })) });
  r.remoteHtml = '<html><a href="' + A + '_r1.docx">' + A + '_r1.docx</a></html>';
  r.s.updateRevisions_(r.s.getCollectorConfig_());
  check('one update: the agent listing, then the 3GPP folder -- the file in both places is one line, [local]', [r.lines(0), r.state().lastResult, r.fetched.length], [[A + '_r1.docx [local]'], 'imported', 1]);
  const failing = report({ 'MBS.listing.json': file(listing([B + '_r1.docx'], { fetchedAt: new Date(Date.now() - 60000).toISOString() })) });
  failing.remoteHtml = null;
  failing.s.updateRevisions_(failing.s.getCollectorConfig_());
  check('the 3GPP folder unreachable: the new local revision is shown anyway', failing.lines(1), [B + '_r1.docx [local only]']);
  const broken = report({}, {});
  broken.drive = () => { throw new Error('boom'); };
  check('an agent failure never stops the update', (() => { try { broken.s.updateRevisions_(broken.s.getCollectorConfig_()); return 'ran'; } catch (e) { return e.message; } })(), 'ran');
}

// ====================================================================== 8
console.log('8. status');
{
  const { s } = report({});
  check('not configured', s.describeLocalListingAgentStatus_(null, '', NOW), 'Local listing agent: not configured (Configure Meeting > Local listing folder).');
  check('configured, not checked yet', s.describeLocalListingAgentStatus_(null, 'X', NOW), 'Local listing agent: configured; not checked yet (the next update reads it).');
  const r = report({ 'MBS.listing.json': [{ raw: listing([A + '_r1.docx']), updated: new Date(NOW.getTime() - 60000) }],
    'MBS.status.json': file(JSON.stringify({ ok: false, scanOk: false, fetchedAt: minutesAgo(0), lastSuccessAt: minutesAgo(2), writtenAt: minutesAgo(2), error: 'The operation has timed out.' })) });
  r.importNow();
  check('four stages, each from its own evidence: LAN scan and local write (agent status), Drive (the file found), import (the report)',
    r.s.describeLocalListingAgentStatus_(r.docProps.getProperty('LOCAL_LISTING_AGENT_STATE'), 'X', NOW).split('\n'), [
      'Local listing agent (MBS): last check 2026-11-17 10:00 UTC (0 min ago) -- Imported the scan of 2026-11-17T09:58:00Z: 1 file(s), 1 revision(s) of SA4 TDocs, 1 new, 0 no longer listed.',
      '  1. LAN scan (agent status, as synchronized): last success 2026-11-17 09:58 UTC (2 min ago), last attempt failed: The operation has timed out.',
      '  2. local write (agent status, as synchronized): 2026-11-17 09:58 UTC (2 min ago)',
      '  3. in Google Drive: newest valid scan 2026-11-17 09:58 UTC (2 min ago), uploaded 2026-11-17 09:59 UTC (1 min ago)',
      '  4. import into this report: 2026-11-17 10:00 UTC (0 min ago), 1 file(s), 1 revision(s), 1 new, 0 removed']);
  const written = report({ 'MBS.status.json': file(JSON.stringify({ ok: true, scanOk: true, fetchedAt: minutesAgo(1), lastSuccessAt: minutesAgo(1), writtenAt: minutesAgo(1) })) });
  written.importNow();
  check('written locally but not (yet) in Drive: stage 2 says written, stage 3 says no valid scan -- a local write is not shown as an upload',
    written.s.describeLocalListingAgentStatus_(written.docProps.getProperty('LOCAL_LISTING_AGENT_STATE'), 'X', NOW).split('\n').slice(2, 5),
    ['  2. local write (agent status, as synchronized): 2026-11-17 09:59 UTC (1 min ago)', '  3. in Google Drive: no valid scan found', '  4. import into this report: none yet']);
}

// ====================================================================== 8b
console.log('8b. all five groups; new revisions during the meeting; the scheduled update');
{
  const groups = { MBS: 'MBS', Video: 'Video', '6G': 'FS_6G_MED', RTC: 'RTC', Audio: 'Audio' };
  const all = {};
  Object.keys(groups).forEach((suffix) => {
    const g = groups[suffix];
    all[g + '.listing.json'] = file(listing([A + '_' + g + '.docx'], { group: g, folder: LOCAL + '/Drafts/' + g + '/' }).replace(new RegExp(FOLDER.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), LOCAL + '/Drafts/' + g + '/'));
  });
  const fixed = {};
  Object.keys(all).forEach((name) => {
    const f = JSON.parse(all[name][0].raw);
    f.payload = f.payload.split(FOLDER).join(f.folder);
    f.sha256 = sha(f.payload); f.scanId = f.sha256;
    fixed[name] = file(JSON.stringify(f));
  });
  check('each of the five reports imports only its own group\'s listing', Object.keys(groups).map((suffix) => {
    const r = report(fixed, { REPORT_SUFFIX: suffix });
    r.importNow();
    return [suffix, r.state().group, r.lines(0)];
  }), Object.keys(groups).map((suffix) => [suffix, groups[suffix], [A + '_' + groups[suffix] + '.docx [local only]']]));
  const r = report({ 'MBS.listing.json': file(listing([A + '_r1.docx'], { fetchedAt: minutesAgo(4) })) });
  r.importNow();
  r.drive = { 'MBS.listing.json': file(listing([A + '_r1.docx', B + '_r1.docx', C + 'r02.docx'], { fetchedAt: minutesAgo(2) })) };
  const later = r.importNow();
  check('two scans later in the meeting: the new revisions appear, the earlier one stays, nothing doubled', [later.state.added, r.lines(0), r.lines(1), r.lines(2)],
    [2, [A + '_r1.docx [local only]'], [B + '_r1.docx [local only]'], [C + 'r02.docx [local only]']]);
  // After T-2026.10.29: read without Goes to Plenary, whose update reads the report once first (tests/goes-to-plenary.test.js).
  const code = require('./helpers/before-structure-fix.js').codeWithoutGoesToPlenary(require('./helpers/before-structure-fix.js').codeWithoutUndisposedDocuments(fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '')).code).code;
  const bodyOf = (name) => { const at = code.indexOf('\nfunction ' + name + '('); return code.slice(at, code.indexOf('\n}\n', at)); };
  check('the scheduled update reaches it: collectorUpdate_() runs updateRevisions_(), whose first step is the agent import (source)',
    [/updateRevisions_\(/.test(bodyOf('collectorUpdate_')), /^\nfunction updateRevisions_\(cfg, context\) \{\n(?:  \/\/[^\n]*\n)*  (?:const agent = )?importAgentListingSafely_\(context\);/.test(bodyOf('updateRevisions_'))], [true, true]);
}

// ====================================================================== 8c
console.log('8c. Check Connections: the TDoc list as the update downloads it');
{
  const { s } = report({});
  const XLSX = [0x50, 0x4b, 0x03, 0x04, 1, 2, 3];
  const answer = (map) => { s.UrlFetchApp = { fetch: (u) => { const a = map[u]; if (a instanceof Error) throw a; return { getResponseCode: () => a, getBlob: () => ({ getBytes: () => (a === 200 ? XLSX : []), getContentType: () => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }) }; } }; };
  const FTP = 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx';
  const WWW = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx';
  answer({ [FTP]: 403, [WWW]: 200 });
  check('ftp.3gpp.org refused, its www twin works: said as the update would get it (not just "HTTP 403")', s.tdocListConnectionLines_(FTP),
    ['❌ HTTP 403: ' + FTP + ' -- refused by the 3GPP server (usually temporary; the update tries again on its next run)', '✅ (alternative) Connected successfully: ' + WWW, '📊 File size: 7 bytes (Excel workbook)']);
  answer({ [WWW]: 403 });
  check('www.3gpp.org refused (Cloudflare, temporary): said so', s.tdocListConnectionLines_(WWW), ['❌ HTTP 403: ' + WWW + ' -- refused by the 3GPP server (usually temporary; the update tries again on its next run)']);
  answer({ [WWW]: 404 });
  check('a missing file (wrong number or folder) is not called temporary', s.tdocListConnectionLines_(WWW), ['❌ HTTP 404: ' + WWW]);
}

// ====================================================================== 9
console.log('9. what changed in Code.js (source)');
{
  // The overnight-pause distinction and the Revisions-cell cleanup after T-2026.10.28 came later: compared without them
  // (tests/local-listing-overnight-pause.test.js, tests/revisions-cell-cleanup.test.js).
  const full = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const later = require('./helpers/before-structure-fix.js');
  // After T-2026.10.29: and without Goes to Plenary (tests/goes-to-plenary.test.js).
  const now = later.codeWithoutRevisionsCellCleanup(later.codeWithoutOvernightPause(later.codeWithoutGoesToPlenary(later.codeWithoutUndisposedDocuments(full).code).code).code).code;
  const view = codeWithoutLocalListingAgent(full);
  check('the take-out finds every change exactly once', view.problems, []);
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (code, name) => { const at = code.indexOf('\nfunction ' + name + '('); return code.slice(at, code.indexOf('\n}\n', at)); };
  const before = names(view.code);
  check('the new functions', names(now).filter((n) => before.indexOf(n) === -1).sort(), LOCAL_LISTING_AGENT_NEW_FUNCTIONS.slice().sort());
  check('the changed functions', before.filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)).sort(), LOCAL_LISTING_AGENT_CHANGED_FUNCTIONS.slice().sort());
  check('untouched: the listing parser, the store, the merge, the cell renderer, the 3GPP fetch and matching', ['parseLocalRevisionsListing_', 'storeLocalRevisions_', 'mergeLocalRevisions_', 'renderRevisionsCellContent_', 'parseAnchors_', 'buildRevisionAnchorIndex_', 'insertRevisedDocTablesAfter_', 'continuousUpdateCore_'].filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)), []);
  check('the agent code never fetches anything', LOCAL_LISTING_AGENT_NEW_FUNCTIONS.filter((n) => /UrlFetchApp|safeFetch_/.test(bodyOf(now, n))), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll local listing agent checks passed.');
process.exit(failures ? 1 : 0);
