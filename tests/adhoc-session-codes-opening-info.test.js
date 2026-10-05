/**
 * Ad-hoc sessions -- the short codes of the sessions in the registration
 * table, and the Online information block of the opening (G2 visual review).
 *
 * Codes. A session label such as "AHG Call 1" does not fit the Session
 * column of the registration table. The column now shows a code -- A01,
 * A02, ... by the chronological order of the sessions -- and the Session
 * administration section names each session with its code and its label.
 * A code is derived every time; nothing is stored for it.
 *
 * Online information. A main-meeting report copies its opening from the
 * report template. The opening of an ad-hoc report is generated and said
 * nothing about the meeting as a whole. With sessions, a build now writes a
 * block from the report's own configuration, above Session administration.
 *
 *   1. the codes;
 *   2. the registration table: build, update, adding the column, an older table;
 *   3. where the label stays;
 *   4. the Online information block: its lines, its links, its place;
 *   5. the range of the meeting, sessions of several days;
 *   6. values the report does not have;
 *   7. kept up to date, never doubled, never created outside a build; typed text is safe;
 *   8. reports without sessions and main-meeting reports are unchanged;
 *   9. nothing is stored; a problem never fails a build.
 *
 * All names and values are synthetic.
 *
 * Run: node tests/adhoc-session-codes-opening-info.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const FX = require('./fixtures/teams-attendance-synthetic.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log(`  ok   ${name}`);
  } else {
    failures++;
    console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`);
  }
}

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const DASH = '–';
const INFO = 'Online information';
const ADMIN = 'Session administration';
const session = (id, date, endDate, start, end, label) => { const s = { id: id, label: label || '', date: date }; if (endDate) s.endDate = endDate; s.start = start || ''; s.end = end || ''; return s; };
const CALL1 = session('s1', '2026-09-22', '', '15:00', '18:00', 'AHG Call 1');
const CALL2 = session('s2', '2026-10-26', '2026-10-28', '15:00', '23:00', 'AHG Call 2');
const sessionsValue = (list, nextId) => JSON.stringify({ v: 1, nextId: nextId || list.length + 1, sessions: list });
const LIST_URL = 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916';
const DOCS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc on FS_EXAMPLE', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: DOCS, TDOC_LIST_URL: LIST_URL, AGENDA_TDOC: 'S4aA260090' };
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Topic'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
/** id, agenda item, status, uploaded (UTC): one for the first session, one for the second, one after the last, one reserved */
const TDOCS = [['S4aA269001', '4.3', 'available', '2026-09-21 08:00:00'], ['S4aA269002', '4.4', 'available', '2026-09-30 08:00:00'], ['S4aA269003', '4.4', 'available', '2026-11-02 08:00:00'], ['S4aA269004', '4.4', 'reserved', '']];
const LINES = ['Meeting name: Synthetic ad-hoc on FS_EXAMPLE', 'Start Date: September 22, 2026', 'End Date: October 28, 2026', 'Portal meeting URL: https://portal.3gpp.org/Home.aspx#/meeting?MtgId=85916',
  'TDoc List URL: https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=85916', 'Excel Docs URL: ' + LIST_URL, 'Docs Folder: ' + DOCS];

function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

/** One report: the real build, the real TDoc-list reader, the real update. `props` replaces or adds properties; a null value leaves one out. */
function report(props) {
  const merged = Object.assign({}, PROPS, { ADHOC_SESSIONS: sessionsValue([CALL1, CALL2]) }, props || {});
  Object.keys(merged).forEach((k) => { if (merged[k] === null) delete merged[k]; });
  const loaded = loadCode({ documentProperties: merged });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const r = { s: s, body: body, docProps: loaded.docProps, scriptProps: loaded.scriptProps, tdocs: TDOCS.slice(), logs: [] };
  const ui = { alerts: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' }, alert: (...args) => { ui.alerts.push(args); return 'YES'; }, showModalDialog: () => {} };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.Utilities.formatDate = (date, zone) => formatInZone(date, zone);
  s.Utilities.base64Decode = (b64) => Array.from(Buffer.from(b64, 'base64')).map((b) => (b > 127 ? b - 256 : b));
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  s.downloadMeetingAgenda_ = () => AGENDA;
  s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }) };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const values = [HEADER].concat(r.tdocs.map(([id, agenda, status, uploaded]) => [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', agenda, 'Synthetic item', status, '',
      uploaded ? new Date(uploaded.replace(' ', 'T') + 'Z') : '', '', '']));
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((row) => row.map(() => null)) }),
      getRange: (row, col, rows) => ({ getDisplayValues: () => [['Uploaded']].concat(r.tdocs.map((t) => [t[3]])).slice(row - 1, row - 1 + rows) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };

  r.build = () => { try { return s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }); } catch (e) { return { ok: false, message: e.message }; } };
  r.update = () => {
    s.insertNewTdoc_ = (b, td) => {
      const at = s.findInsertionPointForAgendaItem_(b, td.agendaItem, '');
      const cells = [['TDoc', td.row[td.tdocCol]], ['Title', td.row[td.titleCol]], ['Minutes', '']];
      if (at >= b.getNumChildren()) b.appendTable(cells); else b.insertTable(at, cells);
    };
    s.updateTdocStatus_ = () => false;
    s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
    s.checkRSSFeed_ = () => {};
    s.updateRevisions_ = () => {};
    return s.continuousUpdateCore_();
  };
  r.texts = () => body._children.map((c) => (c.getType() === 'TABLE' ? '[' + rowsOf(c)[0].join('|') + ']' : c.getText()));
  r.at = (text) => r.texts().indexOf(text);
  r.count = (text) => r.texts().filter((t) => t === text).length;
  r.registration = () => { const found = s.findRegistrationTable_(body); return found ? rowsOf(found.table) : null; };
  r.cells = () => r.registration().slice(1).map((row) => row[0] + ' ' + row[row.length - 1]);
  /** The opening sub-section: from its heading to the Registration of Documents heading. */
  r.opening = () => r.texts().slice(r.at('1.1 Opening of the session'), r.at('1.2 Registration of Documents') + 1);
  r.info = () => { const c = s.findAdhocMeetingInformationContainer_(body); return c ? r.texts().slice(c.start, c.end) : null; };
  r.admin = () => { const c = s.findAdhocOpeningContainer_(body); return c ? r.texts().slice(c.start, c.end) : null; };
  r.stored = () => JSON.stringify([r.docProps.getKeys().sort().map((k) => [k, r.docProps.getProperty(k)]), r.scriptProps.getKeys().sort().map((k) => [k, r.scriptProps.getProperty(k)])]);
  return r;
}

const S = loadCode().sandbox;

// ================================================================ 1. the codes

console.log('1. the codes');
{
  check('the first sessions: A01, A02, A03', [0, 1, 2].map((i) => S.adhocSessionCode_(i)), ['A01', 'A02', 'A03']);
  check('the ninth, the tenth, the 99th and the 100th', [8, 9, 98, 99].map((i) => S.adhocSessionCode_(i)), ['A09', 'A10', 'A99', 'A100']);
  const a = session('s1', '2026-09-22'); const b = session('s2', '2026-09-23'); const c = session('s3', '2026-09-24');
  check('by the chronological order of the sessions', S.adhocSessionCodes_([a, b, c]), { s1: 'A01', s2: 'A02', s3: 'A03' });
  check('the order in which they are given plays no part', S.adhocSessionCodes_([c, a, b]), { s1: 'A01', s2: 'A02', s3: 'A03' });
  check('nor do the ids: a session added later with an earlier date comes first', S.adhocSessionCodes_([a, b, session('s7', '2026-09-01')]), { s7: 'A01', s1: 'A02', s2: 'A03' });
  check('two sessions on one day: by their start', S.adhocSessionCodes_([session('s1', '2026-09-22', '', '15:00', '18:00'), session('s2', '2026-09-22', '', '09:00', '12:00')]), { s2: 'A01', s1: 'A02' });
  check('a session of several days has one code', S.adhocSessionCodes_([CALL2, CALL1]), { s1: 'A01', s2: 'A02' });
  check('a label plays no part: renaming a session leaves every code as it is',
    [S.adhocSessionCodes_([Object.assign({}, a, { label: 'Kick-off' }), b, c]), S.adhocSessionCodes_([Object.assign({}, a, { label: 'Zzz' }), Object.assign({}, b, { label: 'Aaa' }), c])], [{ s1: 'A01', s2: 'A02', s3: 'A03' }, { s1: 'A01', s2: 'A02', s3: 'A03' }]);
  check('no sessions: no codes', S.adhocSessionCodes_([]), {});
  check('the cell of a TDoc: one code, several in order, the dash for none',
    [S.adhocTdocSessionCellText_(['s2'], [a, b, c]), S.adhocTdocSessionCellText_(['s3', 's1'], [a, b, c]), S.adhocTdocSessionCellText_([], [a, b, c]), S.adhocTdocSessionCellText_(['s9'], [a, b, c])], ['A02', 'A01, A03', DASH, DASH]);
  check('the cell never shows a label', S.adhocTdocSessionCellText_(['s1', 's2'], [CALL1, CALL2]), 'A01, A02');
  check('the ten-session case: the tenth is A10', S.adhocTdocSessionCellText_(['s10'], Array.from({ length: 10 }, (_, i) => session('s' + (i + 1), '2026-10-' + String(i + 1).padStart(2, '0')))), 'A10');
}

// ================================================================ 2. the registration table

console.log('\n2. the registration table: build, update, adding the column, an older table');
{
  const r = report();
  check('the build completes without a note', [r.build().ok, r.build().note], [true, '']);
  check('Build Report from Scratch: the Session column has the codes, a dash for no session', [r.registration()[0], r.cells()],
    [['TDoc', 'Title', 'Source', 'Agenda Item', 'Session'], ['S4aA269001 A01', 'S4aA269002 A02', 'S4aA269003 ' + DASH, 'S4aA269004 ' + DASH]]);
  check('no label is in the registration table', JSON.stringify(r.registration()).indexOf('AHG Call') === -1, true);
  check('the widths: the Session column is narrow, the Title column wide, the TDoc column wide enough for a TDoc number', (() => { const t = r.s.findRegistrationTable_(r.body).table; return Array.from({ length: 5 }, (_, i) => t.getRow(1).getCell(i).getWidth()); })(), [89, 197, 89, 56, 37]);

  r.tdocs.push(['S4aA269005', '4.4', 'available', '2026-10-27 08:00:00']);
  r.update();
  check('Update Report Now: a new TDoc gets the code of its session', r.cells().slice(-1), ['S4aA269005 A02']);
  const before = JSON.stringify(r.registration());
  r.update(); r.update();
  check('further updates leave the table as it is', JSON.stringify(r.registration()) === before, true);

  check('a manual assignment to two sessions shows both codes', [r.s.saveAdhocTdocSessions({ clock: 'utc', tdocs: { S4aA269003: { mode: 'set', sessions: ['s1', 's2'] } } }).ok, r.cells()[2]], [true, 'S4aA269003 A01, A02']);

  // A table of an earlier candidate: the labels in the Session column.
  const older = report();
  older.build();
  const table = older.s.findRegistrationTable_(older.body).table;
  table.getRow(1).getCell(4).setText('AHG Call 1'); table.getRow(2).getCell(4).setText('AHG Call 2');
  older.update();
  check('a table that still shows labels: the next update writes the codes', older.cells(), ['S4aA269001 A01', 'S4aA269002 A02', 'S4aA269003 ' + DASH, 'S4aA269004 ' + DASH]);

  // Adding the column to a table built without it.
  const plain = report({ ADHOC_SESSIONS: null });
  plain.build();
  plain.docProps.setProperty('ADHOC_SESSIONS', sessionsValue([CALL1, CALL2]));
  check('adding the Session column to an existing table: the codes', [plain.registration()[0].length, plain.s.enableAdhocSessionColumn().changed, plain.cells()], [4, true, ['S4aA269001 A01', 'S4aA269002 A02', 'S4aA269003 ' + DASH, 'S4aA269004 ' + DASH]]);

  // Renaming and re-dating.
  const renamed = report();
  renamed.build();
  renamed.s.saveAdhocSessionsConfiguration([Object.assign({}, CALL1, { label: 'Kick-off' }), Object.assign({}, CALL2)]);
  renamed.update();
  check('renaming a session changes no cell of the table', renamed.cells(), ['S4aA269001 A01', 'S4aA269002 A02', 'S4aA269003 ' + DASH, 'S4aA269004 ' + DASH]);
  renamed.s.saveAdhocSessionsConfiguration([Object.assign({}, CALL1, { label: 'Kick-off' }), Object.assign({}, CALL2), session('', '2026-09-01', '', '09:00', '10:00', 'Preparation')]);
  renamed.update();
  check('a session added before the others takes A01; the table and the Session administration section follow together',
    [renamed.cells(), renamed.admin().filter((t) => /^A\d\d: /.test(t)).map((t) => t.split(',')[0])],
    [['S4aA269001 A02', 'S4aA269002 A03', 'S4aA269003 ' + DASH, 'S4aA269004 ' + DASH], ['A01: Preparation', 'A02: Kick-off', 'A03: AHG Call 2']]);
}

// ================================================================ 3. where the label stays

console.log('\n3. where the label stays');
{
  const r = report();
  r.build();
  check('Session administration names each session with its code and its label', r.admin(),
    [ADMIN, 'A01: AHG Call 1, September 22, 2026, 15:00–18:00', 'A02: AHG Call 2, October 26–28, 2026 (from 15:00 on the first day, until 23:00 on the last day)']);
  check('a session without a label of its own: the code and when it is', S.buildAdhocOpeningLines_([session('s1', '2026-09-22'), session('s2', '2026-09-23', '', '', '', 'Offline')], {}).map((l) => l.text), ['A01: September 22, 2026', 'A02: Offline, September 23, 2026']);
  const all = r.s.flattenTdocGroups_(r.s.downloadAndGroupTdocs_(r.s.getReportConfig_()));
  const model = r.s.buildAdhocTdocSessionsDialogModel_(r.s.makeAdhocTdocSessionResolver_(), r.s.downloadAndGroupTdocs_(r.s.getReportConfig_()), 5);
  check('Assign TDoc Sessions: the sessions by their labels, no code', [model.sessions, /A0\d/.test(JSON.stringify(model))], [[{ id: 's1', label: 'AHG Call 1' }, { id: 's2', label: 'AHG Call 2' }], false]);
  check('Configure Sessions and the opening dialog: labels, no code', [/A0\d/.test(JSON.stringify(r.s.buildAdhocSessionsDialogModel_(r.docProps.getProperty('ADHOC_SESSIONS'), 'September 22, 2026'))), r.s.buildAdhocOpeningDialogModel_().sessions.map((x) => x.when)],
    [false, ['AHG Call 1, September 22, 2026, 15:00–18:00', 'AHG Call 2, October 26–28, 2026 (from 15:00 on the first day, until 23:00 on the last day)']]);
  const b64 = FX.utf16le(FX.buildExport()).toString('base64');
  check('(an attendance import)', r.s.confirmTeamsAttendanceImport('s1', b64, r.s.previewTeamsAttendanceImport('s1', b64).token).ok, true);
  check('Attendance: the session by its label', r.texts().filter((t) => /^Session: /.test(t)), ['Session: AHG Call 1, September 22, 2026']);
  const status = r.s.adhocSessionStatusLines_().join('\n');
  check('the status summary: the sessions are counted and described, without a code', [/SESSIONS AND ATTENDANCE/.test(status), /not available/.test(status), /\bA0\d\b/.test(status)], [true, false, false]);
  check('(the TDocs of this report)', all.length, 4);
  check('errors of Configure Sessions name sessions by their labels', (() => { const out = r.s.saveAdhocSessionsConfiguration([Object.assign({}, CALL2)]); return [out.ok, /AHG Call 1/.test(out.errors.join(' ')), /A0\d/.test(out.errors.join(' '))]; })(), [false, true, false]);
}

// ================================================================ 4. the Online information block

console.log('\n4. the Online information block: its lines, its links, its place');
{
  const r = report();
  r.build();
  check('the opening sub-section: the block, then Session administration, then Registration of Documents', r.opening(),
    ['1.1 Opening of the session', INFO].concat(LINES, [ADMIN, 'A01: AHG Call 1, September 22, 2026, 15:00–18:00', 'A02: AHG Call 2, October 26–28, 2026 (from 15:00 on the first day, until 23:00 on the last day)', '1.2 Registration of Documents']));
  check('the block is these lines', r.info(), [INFO].concat(LINES));
  const paragraphs = r.body._children.slice(r.at(INFO), r.at(ADMIN));
  check('its heading is a heading; its lines are plain paragraphs, not bold', [paragraphs[0].getHeading() !== 'NORMAL', paragraphs.slice(1).map((p) => p.getHeading() === 'NORMAL' && p._bold === false).every(Boolean)], [true, true]);
  check('the four addresses are links: the address itself, and only it',
    paragraphs.slice(1).map((p) => (p._links || []).map(([start, end, url]) => [p.getText().slice(start, end + 1) === url, url === p.getText().split(': ').slice(1).join(': ')])),
    [[], [], [], [[true, true]], [[true, true]], [[true, true]], [[true, true]]]);
  check('the meeting name and the dates are not links', paragraphs.slice(1, 4).map((p) => (p._links || []).length), [0, 0, 0]);
  check('the pure lines, from the same values', S.buildAdhocMeetingInformationLines_([CALL1, CALL2], { meetingName: PROPS.MEETING_NAME, meetingId: '85916', tdocListUrl: LIST_URL, docsFolder: DOCS, draftsFolder: '' }).map((l) => l.text), LINES);
  check('a drafts folder the report is configured with is shown, last',
    S.buildAdhocMeetingInformationLines_([CALL1], { meetingName: '', meetingId: '', tdocListUrl: '', docsFolder: DOCS, draftsFolder: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/' }).map((l) => l.text).slice(-2),
    ['Docs Folder: ' + DOCS, 'Drafts Folder: https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/']);
  const withDrafts = report({ REVISIONS_URL: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/' });
  withDrafts.build();
  check('in a built report as well', withDrafts.info().slice(-1), ['Drafts Folder: https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/']);
  check('nothing of a main-meeting report is in it: no scribes, no shared minutes, no text of a report family', r.info().filter((t) => /scribe|minutes|MBS|CEST|<|>/i.test(t)), []);
  check('the block says nothing that Session administration says: no session, no time, no chair', r.info().filter((t) => /AHG Call|A0\d|15:00|Chair|Minute taker/.test(t)), []);
  check('the "<Chair> opens the session ..." line is not written', r.texts().filter((t) => /<Chair>|opens the session on/.test(t)), []);
}

// ================================================================ 5. the range of the meeting

console.log('\n5. the range of the meeting; sessions of several days');
{
  const dates = (list) => S.buildAdhocMeetingInformationLines_(list, {}).map((l) => l.text);
  check('one session of one day: start and end are that day', dates([CALL1]), ['Start Date: September 22, 2026', 'End Date: September 22, 2026']);
  check('one session of several days: its first and its last day', dates([CALL2]), ['Start Date: October 26, 2026', 'End Date: October 28, 2026']);
  check('several sessions: the first day of the first, the last day of the last', dates([CALL2, CALL1]), ['Start Date: September 22, 2026', 'End Date: October 28, 2026']);
  check('over the turn of the year', dates([session('s1', '2026-12-30', '2027-01-02'), session('s2', '2026-12-01')]), ['Start Date: December 1, 2026', 'End Date: January 2, 2027']);
  check('the times of the sessions are not part of it', dates([CALL1, CALL2]).filter((t) => /\d\d:\d\d/.test(t)), []);
  check('no sessions: no block at all', [dates([]), S.buildAdhocMeetingInformationLines_([], { meetingName: 'x', meetingId: '1', tdocListUrl: LIST_URL, docsFolder: DOCS, draftsFolder: DOCS })], [[], []]);
}

// ================================================================ 6. values the report does not have

console.log('\n6. values the report does not have are left out, never made up');
{
  const lines = (source) => S.buildAdhocMeetingInformationLines_([CALL1], source).map((l) => l.text);
  const DATES = ['Start Date: September 22, 2026', 'End Date: September 22, 2026'];
  check('nothing configured: the dates only', [lines({}), lines(null), lines({ meetingName: '', meetingId: '', tdocListUrl: '', docsFolder: '', draftsFolder: '' })], [DATES, DATES, DATES]);
  check('no placeholder stands for a missing value', lines({}).filter((t) => /<|>|TBD|unknown|n\/a|undefined|null/i.test(t)), []);
  check('no meeting id: no Portal links; the list and the folder stay', lines({ meetingName: 'M', meetingId: '', tdocListUrl: LIST_URL, docsFolder: DOCS }), ['Meeting name: M'].concat(DATES, ['Excel Docs URL: ' + LIST_URL, 'Docs Folder: ' + DOCS]));
  check('a meeting id that is not a number is not used', lines({ meetingId: '85916); DROP' }).concat(lines({ meetingId: 'abc' })), DATES.concat(DATES));
  check('a value that is not an address is not written as one', lines({ tdocListUrl: 'not a url', docsFolder: 'ftp.3gpp.org/x', draftsFolder: 'javascript:alert(1)' }), DATES);
  check('a meeting name is written as text, on one line', lines({ meetingName: '  Synthetic\n  <b>meeting</b>  ' })[0], 'Meeting name: Synthetic <b>meeting</b>');

  // From the configuration of a report.
  const source = (props) => { const r = report(props); return r.s.adhocMeetingInformationSource_(); };
  check('the meeting id is the one the TDoc list URL of the Portal names', source().meetingId, '85916');
  check('a list on the file server names no meeting: no id, so no Portal links -- the stored MEETING_ID and the defaults of a main meeting are not used',
    [source({ TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/x/Docs/TDoc_List.xlsx' }).meetingId, source({ TDOC_LIST_URL: 'https://example.org/list?meetingId=12345' }).meetingId, source({ TDOC_LIST_URL: null, MEETING_ID: null }).meetingId], ['', '', '']);
  check('a report without a meeting name, a docs folder or a list: empty values, not the defaults of a main meeting', (() => { const x = source({ MEETING_NAME: null, FTP_BASE: null, TDOC_LIST_URL: null }); return [x.meetingName, x.docsFolder, x.tdocListUrl, x.draftsFolder]; })(), ['', '', '', '']);
  const bare = report({ MEETING_NAME: null });
  bare.build();
  check('a built report without a meeting name: the block starts with the dates', bare.info().slice(0, 3), [INFO, 'Start Date: September 22, 2026', 'End Date: October 28, 2026']);
}

// ================================================================ 7. kept up to date, never doubled

console.log('\n7. kept up to date, never doubled, never created outside a build; typed text is safe');
{
  const r = report();
  r.build();
  r.build();
  check('a second build: the block once, Session administration once', [r.count(INFO), r.count(ADMIN), r.info()], [1, 1, [INFO].concat(LINES)]);
  check('each label once', LINES.map((l) => r.texts().filter((t) => t.indexOf(l.split(': ')[0] + ': ') === 0).length), [1, 1, 1, 1, 1, 1, 1]);

  // Text typed by hand directly below the block, where the opening is minuted.
  r.body.insertParagraph(r.at(ADMIN), 'The chair welcomes the participants. (typed by hand)');
  const elements = r.body._children.slice(r.at(INFO), r.at(INFO) + 1 + LINES.length);
  r.s.saveAdhocOpeningDetails({ s1: { chair: 'Alex Organizer', minuteTakers: 'Sam Rivera', note: '' }, s2: { chair: '', minuteTakers: '', note: '' } });
  check('saving opening details: the block is the very same paragraphs, and the typed text is where it was',
    [r.body._children.slice(r.at(INFO), r.at(INFO) + 1 + LINES.length).every((e, i) => e === elements[i]), r.texts()[r.at(ADMIN) - 1], r.info()], [true, 'The chair welcomes the participants. (typed by hand)', [INFO].concat(LINES)]);

  r.s.saveAdhocSessionsConfiguration([Object.assign({}, CALL1), Object.assign({}, CALL2), session('', '2026-11-10', '2026-11-12', '', '', 'AHG Call 3')]);
  check('a session added later: the End Date of the block follows at once; the block is still there once', [r.info().filter((t) => /Date/.test(t)), r.count(INFO)], [['Start Date: September 22, 2026', 'End Date: November 12, 2026'], 1]);
  check('the typed text below it is still there, once', [r.texts()[r.at(ADMIN) - 1], r.count('The chair welcomes the participants. (typed by hand)')], ['The chair welcomes the participants. (typed by hand)', 1]);
  r.update(); r.update();
  check('an update writes neither part again', [r.count(INFO), r.count(ADMIN), r.texts()[r.at(ADMIN) - 1]], [1, 1, 'The chair welcomes the participants. (typed by hand)']);

  // A report whose block was removed by hand, or that was built before the block existed.
  const without = report();
  without.build();
  const c = without.s.findAdhocMeetingInformationContainer_(without.body);
  for (let i = c.end - 1; i >= c.start; i--) without.body.removeChild(without.body.getChild(i));
  without.s.saveAdhocOpeningDetails({ s1: { chair: 'Alex Organizer', minuteTakers: '', note: '' }, s2: { chair: '', minuteTakers: '', note: '' } });
  without.s.saveAdhocSessionsConfiguration([Object.assign({}, CALL1, { label: 'Kick-off' }), Object.assign({}, CALL2)]);
  without.update();
  check('a report without the block does not get one from saving details, saving sessions or an update', [without.count(INFO), without.admin().slice(0, 2)], [0, [ADMIN, 'A01: Kick-off, September 22, 2026, 15:00–18:00']]);
  check('the render says so', without.s.renderAdhocMeetingInformation_(without.body, [{ text: 'Start Date: x', link: '' }], false), { container: 'absent' });
  without.build();
  check('Build Report from Scratch writes it', [without.count(INFO), without.at(INFO) < without.at(ADMIN)], [1, true]);

  // The render itself.
  const body = makeFakeDocumentBody(S);
  const H = S.DocumentApp.ParagraphHeading;
  body.appendParagraph('1.1 Opening of the session').setHeading(H.HEADING3);
  body.appendParagraph('1.2 Registration of Documents').setHeading(H.HEADING3);
  const texts = () => body._children.map((x) => x.getText());
  const two = [{ text: 'Start Date: September 22, 2026', link: '' }, { text: 'Docs Folder: ' + DOCS, link: DOCS }];
  check('created before Registration of Documents', [S.renderAdhocMeetingInformation_(body, two, true), texts()], [{ container: 'created' }, ['1.1 Opening of the session', INFO, 'Start Date: September 22, 2026', 'Docs Folder: ' + DOCS, '1.2 Registration of Documents']]);
  check('the same lines again: unchanged', S.renderAdhocMeetingInformation_(body, two, true), { container: 'unchanged' });
  body.insertParagraph(4, 'Typed below the block.');
  check('other lines: replaced, and text below the block is kept', [S.renderAdhocMeetingInformation_(body, two.slice(0, 1), false), texts()], [{ container: 'replaced' }, ['1.1 Opening of the session', INFO, 'Start Date: September 22, 2026', 'Typed below the block.', '1.2 Registration of Documents']]);
  check('no lines: the block is removed, the text below it stays', [S.renderAdhocMeetingInformation_(body, [], true), texts()], [{ container: 'removed' }, ['1.1 Opening of the session', 'Typed below the block.', '1.2 Registration of Documents']]);
  check('no lines and no block: nothing', S.renderAdhocMeetingInformation_(body, [], true), { container: 'none' });
  const nowhere = makeFakeDocumentBody(S);
  nowhere.appendParagraph('Some text.');
  check('a document without an opening: no place, nothing written', [S.renderAdhocMeetingInformation_(nowhere, two, true), nowhere.getNumChildren()], [{ container: 'no-place' }, 1]);
  const withAdmin = makeFakeDocumentBody(S);
  withAdmin.appendParagraph('1.1 Opening of the session').setHeading(H.HEADING3);
  withAdmin.appendParagraph('Typed above the section.');
  withAdmin.appendParagraph(ADMIN).setHeading(H.HEADING3);
  withAdmin.appendParagraph('A01: September 22, 2026');
  withAdmin.appendParagraph('1.2 Registration of Documents').setHeading(H.HEADING3);
  S.renderAdhocMeetingInformation_(withAdmin, two.slice(0, 1), true);
  check('with a Session administration section already there: directly above it', withAdmin._children.map((x) => x.getText()),
    ['1.1 Opening of the session', 'Typed above the section.', INFO, 'Start Date: September 22, 2026', ADMIN, 'A01: September 22, 2026', '1.2 Registration of Documents']);
  check('and that section is found as before', (() => { const k = S.findAdhocOpeningContainer_(withAdmin); return withAdmin._children.slice(k.start, k.end).map((x) => x.getText()); })(), [ADMIN, 'A01: September 22, 2026']);
}

// ================================================================ 8. reports without sessions; main-meeting reports

console.log('\n8. reports without sessions and main-meeting reports are unchanged');
{
  const plain = report({ ADHOC_SESSIONS: null });
  plain.build();
  check('an ad-hoc report without sessions: the one opening line, four columns, no block, no section, no code',
    [plain.opening(), plain.registration()[0], plain.count(INFO), plain.count(ADMIN), /\bA0\d\b/.test(plain.texts().join('\n') + JSON.stringify(plain.registration()))],
    [['1.1 Opening of the session', '<Chair> opens the session on September 22, 2026 at <start> CEST.', '1.2 Registration of Documents'], ['TDoc', 'Title', 'Source', 'Agenda Item'], 0, 0, false]);
  const snapshot = JSON.stringify(plain.texts()) + JSON.stringify(plain.registration());
  plain.update();
  plain.build();
  check('after an update and a rebuild: the same report', JSON.stringify(plain.texts()) + JSON.stringify(plain.registration()) === snapshot, true);

  const main = report({ MEETING_TYPE: 'main', REPORT_SUFFIX: 'Audio', AGENDA_ITEM_PREFIX: '4.' });
  const built = main.build();
  check('a main-meeting report, even with a sessions value left in it: no block, no section, four columns, no code',
    [built.ok, main.count(INFO), main.count(ADMIN), (main.registration() || [[]])[0].length <= 4, /\bA0\d\b/.test(main.texts().join('\n'))], [true, 0, 0, true, false]);
  check('the lines are for a report with sessions only (source): the block is written from the refresh of the Session administration section, which a build calls with sessions only',
    [(CODE.match(/refreshAdhocMeetingInformationSafely_\(/g) || []).length, /function refreshAdhocOpeningSection_\(body, options\) \{\n  const sessions = getAdhocSessions_\(\);\n  \/\/[^\n]*\n  refreshAdhocMeetingInformationSafely_\(body, sessions, !!\(options && options\.afterRebuild\)\);/.test(CODE),
      /function finishAdhocOpeningRebuild_\(body\) \{\n  if \(!adhocSessionsEnabled_\(\)\) return '';/.test(CODE)], [2, true, true]);
  check('no text of a report family is in the code of the block (source)', /MBS|Scribes|shared minutes/i.test(CODE.slice(CODE.indexOf('// --- Online information'), CODE.indexOf('function refreshAdhocOpeningSection_('))), false);
}

// ================================================================ 9. nothing is stored; a problem never fails a build

console.log('\n9. nothing is stored for the codes or the block; a problem never fails a build');
{
  const r = report();
  const sessionsBefore = r.docProps.getProperty('ADHOC_SESSIONS');
  r.build();
  r.tdocs.push(['S4aA269005', '4.4', 'available', '2026-10-27 08:00:00']);
  r.update();
  r.s.saveAdhocTdocSessions({ clock: 'utc', tdocs: { S4aA269003: { mode: 'set', sessions: ['s1', 's2'] } } });
  r.s.saveAdhocOpeningDetails({ s1: { chair: 'Alex Organizer', minuteTakers: '', note: '' }, s2: { chair: '', minuteTakers: '', note: '' } });
  check('after a build, an update and two saves: no code in any stored property', /\bA\d\d\b/.test(r.stored()), false);
  check('the sessions are stored exactly as they were: ids, labels, dates', r.docProps.getProperty('ADHOC_SESSIONS') === sessionsBefore, true);
  check('the manual assignment names session ids', JSON.parse(r.docProps.getProperty('ADHOC_TDOC_SESSIONS')).set, { 's1+s2': ['S4aA269003'] });
  check('nothing of the block is stored: no property holds its heading or a line of it', [/Online information|Start Date|End Date|Portal meeting URL/.test(r.stored())], [false]);
  check('a session record has no code (source and value)', [Object.keys(JSON.parse(sessionsBefore).sessions[0]), /code/i.test(CODE.slice(CODE.indexOf('\nfunction adhocSessionRecord_('), CODE.indexOf('\n}\n', CODE.indexOf('\nfunction adhocSessionRecord_('))))], [['id', 'label', 'date', 'start', 'end'], false]);
  check('the functions of the codes and of the lines write nothing (source)', ['adhocSessionCode_', 'adhocSessionCodes_', 'buildAdhocMeetingInformationLines_', 'adhocMeetingInformationSource_'].map((name) => {
    const at = CODE.indexOf('\nfunction ' + name + '('); return /setProperty|deleteProperty|UrlFetchApp|Logger/.test(CODE.slice(at, CODE.indexOf('\n}\n', at)));
  }), [false, false, false, false]);

  const failing = report();
  failing.s.renderAdhocMeetingInformation_ = () => { throw new Error('synthetic failure'); };
  const built = failing.build();
  check('the block cannot be written: the report is built, with Session administration and the Session column', [built.ok, built.note, failing.count(INFO), failing.admin().length, failing.cells()[0]], [true, '', 0, 3, 'S4aA269001 A01']);
  check('it is logged, without any value of the report', [failing.logs.filter((l) => /Online information block could not be written/.test(l)), failing.logs.filter((l) => /Online information/.test(l)).some((l) => /Synthetic ad-hoc|portal\.3gpp|ftp\.3gpp/.test(l))],
    [['Opening details: the Online information block could not be written: synthetic failure'], false]);
  check('the safe wrapper reports it and does not throw', failing.s.refreshAdhocMeetingInformationSafely_(failing.body, failing.s.getAdhocSessions_(), true), { container: 'failed' });
}

console.log(failures === 0 ? '\nAll session-code and opening-information checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
