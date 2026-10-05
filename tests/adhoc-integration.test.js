/**
 * Ad-hoc sessions -- the generated parts of the report together (release
 * audit G1): the Session administration section, the registration table with
 * its Session column, the TDoc tables, the Attendance section and the prose
 * of the closing section, through a build, updates, refreshes and a rebuild
 * of ONE report.
 *
 * The single features are tested in their own files; this one is about what
 * they do to each other:
 *
 *   1. the order of the parts after a build;
 *   2. each generated container is there once;
 *   3. writing one of them leaves the others, and all prose, alone;
 *   4. an update moves nothing, and a new TDoc table never lands inside a
 *      container -- whichever agenda item it belongs to;
 *   5. a rebuild gives the same report again, from the stored state.
 *
 * All names, TDocs and times are synthetic.
 *
 * Run: node tests/adhoc-integration.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
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

const session = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const SESSIONS = [session('s1', '2026-09-22', '15:00', '18:00'), session('s2', '2026-09-23', '', '', 'Offline'), session('s3', '2026-09-24', '15:00')];
const PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090',
  ADHOC_SESSIONS: JSON.stringify({ v: 1, nextId: 4, sessions: SESSIONS }) };
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Topic'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
/** id, agenda item, status, uploaded (UTC) */
const TDOCS = [['S4aA269001', '4.3', 'available', '2026-09-22 14:00:00'], ['S4aA269002', '4.4', 'available', '2026-09-23 08:00:00'], ['S4aA269003', '4.4', 'available', '2026-09-23 22:00:00']];
const DETAILS = { s1: { chair: 'Alex Organizer', minuteTakers: 'Sam Rivera', note: 'The agenda was approved.' }, s2: { chair: '', minuteTakers: 'Sam Rivera', note: '' }, s3: { chair: '', minuteTakers: '', note: '' } };
const OPENING = 'Session administration';
const ATTENDANCE = 'Attendance';

function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

/** One report: the real build, the real TDoc-list reader, the real update; a new TDoc table is put where the real code would put it. */
function report() {
  const loaded = loadCode({ documentProperties: PROPS });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const r = { s: s, body: body, docProps: loaded.docProps, tdocs: TDOCS.slice(), logs: [], inserted: [] };
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
  /** A complete update. The table of a new TDoc goes where insertNewTdoc_() puts it: findInsertionPointForAgendaItem_(). */
  r.update = () => {
    s.insertNewTdoc_ = (b, td) => {
      const at = s.findInsertionPointForAgendaItem_(b, td.agendaItem, '');
      const cells = [['TDoc', td.row[td.tdocCol]], ['Title', td.row[td.titleCol]], ['Minutes', '']];
      if (at >= b.getNumChildren()) b.appendTable(cells); else b.insertTable(at, cells);
      r.inserted.push(td.row[td.tdocCol] + ' @' + td.agendaItem);
    };
    s.updateTdocStatus_ = () => false;
    s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
    s.checkRSSFeed_ = () => {};
    s.updateRevisions_ = () => {};
    return s.continuousUpdateCore_();
  };
  r.importAttendance = (id) => { const b64 = FX.utf16le(FX.buildExport()).toString('base64'); return s.confirmTeamsAttendanceImport(id, b64, s.previewTeamsAttendanceImport(id, b64).token); };
  r.texts = () => body._children.map((c) => (c.getType() === 'TABLE' ? '[' + rowsOf(c)[0].join('|') + (rowsOf(c)[0][0] === 'TDoc' && rowsOf(c)[0].length === 2 ? '' : ' x' + (c.getNumRows() - 1)) + ']' : c.getText()));
  r.snapshot = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
  r.count = (text) => r.texts().filter((t) => t === text).length;
  r.at = (text) => r.texts().indexOf(text);
  r.container = (which) => { const c = which === OPENING ? s.findAdhocOpeningContainer_(body) : s.findAdhocAttendanceContainer_(body); return c ? body._children.slice(c.start, c.end) : null; };
  r.containerTexts = (which) => { const c = which === OPENING ? s.findAdhocOpeningContainer_(body) : s.findAdhocAttendanceContainer_(body); return c ? r.texts().slice(c.start, c.end) : null; };
  r.same = (elements, which) => { const now = r.container(which); return !!now && now.length === elements.length && now.every((e, i) => e === elements[i]); };
  /** The TDoc tables (two columns, first cell "TDoc") that lie inside a generated container. */
  r.tdocTablesInside = () => [OPENING, ATTENDANCE].map((which) => (r.container(which) || []).filter((e) => e.getType() === 'TABLE' && e.getRow(0).getCell(0).getText() === 'TDoc').length);
  return r;
}

// ================================================================ 1. the order of the parts

console.log('1. a built report with opening details and attendance: the order of the parts');
const r = report();
{
  check('the build completes without a note', [r.build().ok, r.build().note], [true, '']);
  check('saving opening details and importing attendance for two sessions', [r.s.saveAdhocOpeningDetails(DETAILS).ok, r.importAttendance('s1').ok, r.s.writeAdhocAttendanceValue_(r.docProps, 's2',
    JSON.stringify({ t: 'Synthetic second day', n: 2, s: '2026-09-23T09:00:00', e: null, d: null, a: null, p: [['Pat Kim', 'ExampleCorp', 'pat.kim@example.com'], ['Robin Lee', '', '']] })).changed, r.s.refreshAdhocAttendanceSection_(r.body).sessions], [true, true, true, 2]);
  // The sub-sections of the IPR item are fixed text of the build; they are left out here.
  const withoutIprText = (texts) => texts.filter((t, i) => !(i > texts.indexOf('3 IPR') && i < texts.indexOf('4 Topic')));
  check('the whole report, in order', withoutIprText(r.texts()), ['1 Opening of the meeting', '1.1 Opening of the session',
    'Online information', 'Meeting name: Synthetic ad-hoc', 'Start Date: September 22, 2026', 'End Date: September 24, 2026', 'Portal meeting URL: https://portal.3gpp.org/Home.aspx#/meeting?MtgId=85916',
    'TDoc List URL: https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=85916', 'Excel Docs URL: https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916',
    'Docs Folder: https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/',
    OPENING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera', 'The agenda was approved.', 'A02: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera', 'A03: September 24, 2026, 15:00',
    '1.2 Registration of Documents', '[TDoc|Title|Source|Agenda Item|Session x3]',
    '2 Approval of the agenda and registration of documents', '3 IPR', '4 Topic', '4.3 Performance requirements', '[TDoc|S4aA269001]', '4.4 Design constraints', '[TDoc|S4aA269002]', '[TDoc|S4aA269003]',
    '5 Close of the meeting',
    ATTENDANCE, 'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: September 22, 2026', 'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 25', 'Average attendance: 2:31:07',
    'Attendees', '[Name|Company|Email x23]', 'Statistics', 'Meeting: Synthetic second day', 'Session: Offline, September 23, 2026', 'Start: 09:00:00', 'Attendance records: 2', 'Attendees', '[Name|Company|Email x2]',
    '']);   // a document ends with a paragraph: the last one of a replaced container is emptied, not removed
  check('Session administration is in the opening sub-section, the Attendance section in the closing section; the TDoc tables are between them', [r.at('1.1 Opening of the session') < r.at(OPENING), r.at(OPENING) < r.at('1.2 Registration of Documents'),
    r.at('1.2 Registration of Documents') < r.at('[TDoc|S4aA269001]'), r.at('[TDoc|S4aA269003]') < r.at('5 Close of the meeting'), r.at('5 Close of the meeting') < r.at(ATTENDANCE)], [true, true, true, true, true]);
  check('each generated part is there once', [r.count(OPENING), r.count(ATTENDANCE), r.texts().filter((t) => /^\[TDoc\|Title\|Source\|Agenda Item/.test(t)).length, r.count('1.2 Registration of Documents')], [1, 1, 1, 1]);
  check('the Session column says what the TDoc sessions are', rowsOf(r.s.findRegistrationTable_(r.body).table).slice(1).map((row) => row[0] + ' ' + row[4]), ['S4aA269001 A01', 'S4aA269002 A02', 'S4aA269003 A03']);
  check('the build did not write the "<Chair> opens the session ..." line beside the section', r.texts().filter((t) => /<Chair>|CEST/.test(t)), []);
  check('no TDoc table is inside a container; the containers do not overlap', [r.tdocTablesInside(), r.s.findAdhocOpeningContainer_(r.body).end <= r.s.findAdhocAttendanceContainer_(r.body).start], [[0, 0], true]);
}

// ================================================================ 2. writing one part leaves the others alone

console.log('2. writing one part leaves the others, and the prose, alone');
{
  // Prose typed by hand: in the opening sub-section, in the closing section, and minutes in a TDoc table.
  r.body.insertParagraph(r.at(OPENING), 'The chair welcomes the participants. (typed by hand)');
  r.body.insertParagraph(r.at(ATTENDANCE), 'The chair thanks the participants and closes the meeting. (typed by hand)');
  r.body._children[r.at('[TDoc|S4aA269002]')].getRow(1).getCell(1).setText('Minutes typed by hand.');
  const proseIsThere = () => [r.texts()[r.at(OPENING) - 1], r.texts()[r.at(ATTENDANCE) - 1], r.texts()[r.at('5 Close of the meeting') + 1], rowsOf(r.body._children[r.at('[TDoc|S4aA269002]')])[1][1]];
  const PROSE = ['The chair welcomes the participants. (typed by hand)', 'The chair thanks the participants and closes the meeting. (typed by hand)', 'The chair thanks the participants and closes the meeting. (typed by hand)', 'Minutes typed by hand.'];
  check('prose typed above each container, and minutes in a TDoc table', proseIsThere(), PROSE);

  let opening = r.container(OPENING);
  let attendance = r.container(ATTENDANCE);
  const outside = () => { const o = r.s.findAdhocOpeningContainer_(r.body); const a = r.s.findAdhocAttendanceContainer_(r.body); return JSON.stringify(JSON.parse(r.snapshot()).filter((e, i) => !(i >= o.start && i < o.end) && !(i >= a.start && i < a.end))); };
  const rest = outside();

  r.s.refreshAdhocAttendanceSection();
  check('refreshing the Attendance section: the Session administration section is the very same elements; everything outside both containers is unchanged', [r.same(opening, OPENING), outside() === rest, proseIsThere()], [true, true, PROSE]);
  attendance = r.container(ATTENDANCE);
  check('changing the opening details: the Attendance section is the very same elements; everything outside is unchanged', [r.s.saveAdhocOpeningDetails(Object.assign({}, DETAILS, { s3: { chair: 'Alex Organizer', minuteTakers: '', note: '' } })).changed,
    r.same(attendance, ATTENDANCE), outside() === rest, proseIsThere(), r.containerTexts(OPENING).slice(-1)], [true, true, true, PROSE, ['Chair: Alex Organizer']]);
  opening = r.container(OPENING);
  check('renaming a session: both containers follow, each once; the prose and the TDoc tables are untouched', [r.s.saveAdhocSessionsConfiguration(SESSIONS.map((x) => Object.assign({}, x, { label: x.id === 's1' ? 'Kick-off' : x.label }))).ok,
    r.containerTexts(OPENING)[1], r.containerTexts(ATTENDANCE)[3], r.count(OPENING), r.count(ATTENDANCE), proseIsThere()],
    [true, 'A01: Kick-off, September 22, 2026, 15:00–18:00', 'Session: Kick-off, September 22, 2026', 1, 1, PROSE]);
  check('saving manual TDoc sessions: the Session column follows; both containers are the very same elements', (() => {
    opening = r.container(OPENING); attendance = r.container(ATTENDANCE);
    const saved = r.s.saveAdhocTdocSessions({ clock: 'utc', tdocs: { S4aA269001: { mode: 'add', sessions: ['s3'] } } });
    return [saved.ok, rowsOf(r.s.findRegistrationTable_(r.body).table)[1][4], r.same(opening, OPENING), r.same(attendance, ATTENDANCE), proseIsThere()];
  })(), [true, 'A01, A03', true, true, PROSE]);
  check('removing the attendance of one session: the other block stays; the opening section and the prose are untouched', [r.s.removeTeamsAttendanceImport('s2').ok, r.count(ATTENDANCE), r.containerTexts(ATTENDANCE).filter((t) => /^Session:/.test(t)), r.same(opening, OPENING), proseIsThere()],
    [true, 1, ['Session: Kick-off, September 22, 2026'], true, PROSE]);
  check('the status and the statistics change nothing', (() => { const before = r.snapshot(); const props = JSON.stringify(r.docProps._store); r.s.adhocSessionStatusLines_(); r.s.collectAdhocSessionStatus_(); return [r.snapshot() === before, JSON.stringify(r.docProps._store) === props]; })(), [true, true]);
}

// ================================================================ 3. updates

console.log('3. updates move nothing; a new TDoc table never lands inside a container');
{
  const opening = r.container(OPENING);
  const attendance = r.container(ATTENDANCE);
  const before = r.snapshot();
  check('two updates without a new TDoc: the report is unchanged, both containers are the very same elements', [r.update(), r.update().success, r.snapshot() === before, r.same(opening, OPENING), r.same(attendance, ATTENDANCE)], [{ success: true, error: null }, true, true, true, true]);

  // New TDocs: one in a normal item, one in the opening item, one in the opening sub-section, one in the closing item, one in an item the report has no heading for.
  r.tdocs = r.tdocs.concat([['S4aA269004', '4.4', 'available', '2026-09-24 08:00:00'], ['S4aA269005', '1', 'available', '2026-09-22 14:00:00'], ['S4aA269006', '1.1', 'available', '2026-09-22 14:00:00'],
    ['S4aA269007', '5', 'available', '2026-09-24 09:00:00'], ['S4aA269008', '6', 'available', '2026-09-24 10:00:00']]);
  check('an update with five new TDocs completes', [r.update(), r.inserted.slice().sort()], [{ success: true, error: null }, ['S4aA269004 @4.4', 'S4aA269005 @1', 'S4aA269006 @1.1', 'S4aA269007 @5', 'S4aA269008 @6']]);
  check('no TDoc table is inside the Session administration section or the Attendance section', r.tdocTablesInside(), [0, 0]);
  check('both containers are the very same elements as before the update, and still there once', [r.same(opening, OPENING), r.same(attendance, ATTENDANCE), r.count(OPENING), r.count(ATTENDANCE)], [true, true, 1, 1]);
  check('where the new tables are: at the end of their agenda item\'s section -- after a container that ends that section, never in it', [
    r.texts()[r.at('[TDoc|S4aA269004]') - 1], r.texts()[r.at('[TDoc|S4aA269004]') + 1],
    r.at('[TDoc|S4aA269006]') === r.s.findAdhocOpeningContainer_(r.body).end, r.texts()[r.at('[TDoc|S4aA269006]') + 1],
    r.texts()[r.at('[TDoc|S4aA269005]') + 1],
    r.at('[TDoc|S4aA269007]') === r.s.findAdhocAttendanceContainer_(r.body).end,
    r.texts().slice(r.at('[TDoc|S4aA269007]') + 1)],
    ['[TDoc|S4aA269003]', '5 Close of the meeting', true, '1.2 Registration of Documents', '2 Approval of the agenda and registration of documents', true, ['6', '[TDoc|S4aA269008]']]);
  check('the registration table has the new TDocs, each once, with five cells and its session', rowsOf(r.s.findRegistrationTable_(r.body).table).slice(1).map((row) => row.length + ' ' + row[0] + ' ' + row[4]).sort(),
    ['5 S4aA269001 A01, A03', '5 S4aA269002 A02', '5 S4aA269003 A03', '5 S4aA269004 A03', '5 S4aA269005 A01', '5 S4aA269006 A01', '5 S4aA269007 A03', '5 S4aA269008 A03']);
  check('the prose typed by hand is where it was', [r.texts()[r.at(OPENING) - 1], r.texts()[r.at(ATTENDANCE) - 1]], ['The chair welcomes the participants. (typed by hand)', 'The chair thanks the participants and closes the meeting. (typed by hand)']);

  // Now write the containers again: the TDoc tables next to them survive.
  const tables = () => r.texts().filter((t) => /^\[TDoc\|S4aA/.test(t));
  const all = tables();
  r.s.refreshAdhocAttendanceSection();
  r.s.saveAdhocOpeningDetails(Object.assign({}, DETAILS, { s1: { chair: 'Another Chair', minuteTakers: '', note: '' } }));
  r.importAttendance('s1');
  check('refreshing the Attendance section, changing the opening details and importing again: every TDoc table is still there, none is inside a container, each container is there once',
    [tables(), r.tdocTablesInside(), r.count(OPENING), r.count(ATTENDANCE), all.length], [all, [0, 0], 1, 1, 8]);
  check('the TDoc table after the Attendance section is still directly after it', r.at('[TDoc|S4aA269007]') === r.s.findAdhocAttendanceContainer_(r.body).end, true);
  check('removing the last attendance removes the Attendance section and nothing else', (() => {
    const snap = JSON.parse(r.snapshot());
    const c = r.s.findAdhocAttendanceContainer_(r.body);
    const expected = JSON.stringify(snap.filter((e, i) => i < c.start || i >= c.end));
    r.s.removeTeamsAttendanceImport('s1');
    return [r.snapshot() === expected, r.count(ATTENDANCE), tables().length];
  })(), [true, 0, 8]);
  check('importing again puts it back at the end of the closing section: after the TDoc table that is in that section now, before the next agenda item', (() => {
    r.importAttendance('s1');
    return [r.texts()[r.at(ATTENDANCE) - 1], r.texts()[r.s.findAdhocAttendanceContainer_(r.body).end], r.count(ATTENDANCE), r.tdocTablesInside(), r.texts()[r.at('[TDoc|S4aA269007]') - 1]];
  })(), ['[TDoc|S4aA269007]', '6', 1, [0, 0], 'The chair thanks the participants and closes the meeting. (typed by hand)']);
  const steady = r.snapshot();
  check('further updates leave the report as it is', [r.update().success, r.update().success, r.snapshot() === steady], [true, true, true]);
}

// ================================================================ 4. a rebuild

console.log('4. a rebuild gives the same report again, from the stored state');
{
  const KEYS = ['ADHOC_SESSIONS', 'ADHOC_SESSION_OPENING', 'ADHOC_SESSION_ATTENDANCE', 'ADHOC_TDOC_SESSIONS'];
  // A Company typed into the attendee table is kept through the rebuild.
  const table = r.body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getCell(0).getText() === 'Name')[0];
  table.getRow(1).getCell(1).setText('ExampleCorp (typed by hand)');
  const typedFor = table.getRow(1).getCell(0).getText();
  const stored = KEYS.map((k) => r.docProps.getProperty(k));
  const rebuilt = r.build();
  check('the rebuild completes without a note', [rebuilt.ok, rebuilt.note], [true, '']);
  check('sessions, opening details and manual TDoc sessions are stored as they were; the attendance only gained the typed Company', [KEYS.map((k, i) => (k === 'ADHOC_SESSION_ATTENDANCE' ? 'x' : r.docProps.getProperty(k) === stored[i])),
    Object.keys(r.s.readAdhocAttendance_(r.docProps).sessions), Object.keys(r.s.readAdhocAttendance_(r.docProps).companies).length], [[true, true, 'x', true], ['s1'], 1]);
  const texts = r.texts();
  check('the parts are in the same order as after the first build, each once', [r.count(OPENING), r.count(ATTENDANCE), r.at('1.1 Opening of the session') + 1 === r.at('Online information') && r.count('Online information') === 1 && r.at('Online information') < r.at(OPENING), r.at(OPENING) < r.at('1.2 Registration of Documents'),
    r.at('1.2 Registration of Documents') < r.at('5 Close of the meeting'), r.at('5 Close of the meeting') < r.at(ATTENDANCE), r.tdocTablesInside()], [1, 1, true, true, true, true, [0, 0]]);
  check('the opening details, the sessions as renamed, the manual TDoc session and the typed Company are all in the rebuilt report', [r.containerTexts(OPENING).slice(1, 3), rowsOf(r.s.findRegistrationTable_(r.body).table).filter((row) => row[0] === 'S4aA269001')[0][4],
    r.containerTexts(ATTENDANCE)[3], rowsOf(r.body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getCell(0).getText() === 'Name')[0]).filter((row) => row[0] === typedFor)[0][1]],
    [['A01: Kick-off, September 22, 2026, 15:00–18:00', 'Chair: Another Chair'], 'A01, A03', 'Session: Kick-off, September 22, 2026', 'ExampleCorp (typed by hand)']);
  check('prose typed by hand is gone, as the build says it will be; it was not stored anywhere', [texts.filter((t) => /typed by hand\)$/.test(t) && !/ExampleCorp/.test(t)), KEYS.map((k) => /welcomes the participants|closes the meeting/.test(String(r.docProps.getProperty(k))))], [[], [false, false, false, false]]);
  check('the Attendance section is the last thing of the rebuilt report, in the closing section', [r.at('5 Close of the meeting') < r.at(ATTENDANCE), r.s.findAdhocAttendanceContainer_(r.body).end === r.body.getNumChildren()], [true, true]);
  const first = r.snapshot();
  r.build();
  check('a second rebuild gives the same report, character for character', r.snapshot() === first, true);
  // The build places the TDocs of the agenda items it writes; the others (here: of the opening sub-section, the closing item and an unknown item) are added by the next update, as in every report.
  const opening = r.container(OPENING);
  const attendance = r.container(ATTENDANCE);
  r.inserted.length = 0;
  r.update();
  check('the first update after the rebuild adds the TDoc tables the build did not place -- next to the containers, never in them; the containers are the very same elements', [r.inserted.slice().sort(), r.tdocTablesInside(), r.same(opening, OPENING), r.same(attendance, ATTENDANCE),
    r.at('[TDoc|S4aA269006]') === r.s.findAdhocOpeningContainer_(r.body).end, r.at('[TDoc|S4aA269007]') === r.s.findAdhocAttendanceContainer_(r.body).end, r.count(OPENING), r.count(ATTENDANCE)],
    [['S4aA269006 @1.1', 'S4aA269007 @5', 'S4aA269008 @6'], [0, 0], true, true, true, true, 1, 1]);
  const steady = r.snapshot();
  r.update();
  check('and a further update changes nothing', r.snapshot() === steady, true);
  check('nothing of all this logged a name, an address or a note', r.logs.filter((l) => /Organizer|Rivera|Another Chair|example\.(com|org|net)|agenda was approved|typed by hand/.test(l)), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll ad-hoc integration checks passed.');
process.exitCode = failures ? 1 : 0;
