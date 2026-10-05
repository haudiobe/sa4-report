/**
 * Ad-hoc sessions, stage F -- the status of the sessions in Report Status
 * Summary, and Sessions and Attendance > Post-meeting Statistics….
 *
 *   1. The TDoc counts (pure), from the stage B assignment.
 *   2. The status model (pure): sessions, opening details, attendance.
 *   3. Report Status Summary.
 *   4. Post-meeting Statistics….
 *   5. The TDoc list cannot be read.
 *   6. Stored state that cannot be used.
 *   7. Nothing is written; the same again gives the same.
 *   8. Privacy.
 *   9. Main-meeting reports and reports without sessions; menu; source.
 *
 * All names, notes, TDocs and upload times are synthetic.
 *
 * Run: node tests/adhoc-session-status.test.js
 */

const fs = require('fs');
const vm = require('vm');
const { execFileSync } = require('child_process');
const path = require('path');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
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
const thrown = (fn) => { try { fn(); return null; } catch (e) { return e.message; } };

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
const SECTION = CODE.slice(CODE.indexOf('// AD-HOC STATUS (stage F)'), CODE.indexOf('// ARCH-009 -- MEETING-ID RESOLVER CORE'));
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.4', flavor: 'template', codeVersion: '0.0.0', gitCommit: 'abcdef0', gitTag: 'template-release/T-2026.10.4', templateDocumentId: TEMPLATE_ID };
const OPENING_KEY = 'ADHOC_SESSION_OPENING';
const ATTENDANCE_KEY = 'ADHOC_SESSION_ATTENDANCE';
const TDOC_KEY = 'ADHOC_TDOC_SESSIONS';
const HEAD = '\n\n🗓 SESSIONS AND ATTENDANCE\n';
const POINTER = '\n   Per session: Sessions and Attendance > Post-meeting Statistics…';
const RANGE = 'September 22–24, 2026';

// ------------------------------------------------------------------ sessions

const session = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const sessionsProperty = (list, nextId) => JSON.stringify({ v: 1, nextId: nextId || list.length + 1, sessions: list });
// s1: 22 Sep 15:00-18:00.  s2 "Offline": 23 Sep, no times.  s3: 24 Sep, a start and no end.
const S1 = session('s1', '2026-09-22', '15:00', '18:00');
const S2 = session('s2', '2026-09-23', '', '', 'Offline');
const S3 = session('s3', '2026-09-24', '15:00');
const SESSIONS = [S1, S2, S3];

// ------------------------------------------------------------------ a synthetic TDoc list

const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
/** id, agenda item, status, uploaded (UTC, as the Portal exports it; '' = not uploaded), revision of, revised to */
const TDOCS = [
  ['S4aA269001', '4.3', 'available', '2026-09-18 07:26:14', '', ''],            // before the first session            -> s1
  ['S4aA269002', '4.3', 'revised', '2026-09-22 14:00:00', '', 'S4aA269010'],     // 16:00 in the report's zone          -> s1
  ['S4aA269003', '4.3', 'available', '2026-09-22 16:00:00', '', ''],            // 18:00:00, exactly at the cut-off    -> s1
  ['S4aA269004', '4.4', 'available', '2026-09-22 16:00:01', '', ''],            // one second later                    -> s2
  ['S4aA269005', '4.4', 'available', '2026-09-23 21:59:59', '', ''],            // the last second of the 23rd         -> s2
  ['S4aA269006', '4.4', 'available', '2026-09-23 22:00:00', '', ''],            // the first second of the 24th        -> s3
  ['S4aA269007', '4.4', 'available', '2026-09-25 09:00:00', '', ''],            // after the last session              -> none
  ['S4aA269008', '4.4', 'reserved', '', '', ''],                                // reserved, not uploaded              -> none
  ['S4aA269009', '4.4', 'available', 'soon', '', ''],                           // not a time                          -> none
  ['S4aA269010', '4.3', 'available', '2026-09-23 08:00:00', 'S4aA269002', '']   // the revision of 9002, a day later   -> s2
];
function listValues(tdocs, header) {
  const h = header || HEADER;
  return [h].concat(tdocs.map(([id, agenda, status, uploaded, revisionOf, revisedTo]) => {
    const row = [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', agenda, 'Synthetic item', status, '18/09/2026 07:00:00',
      /^\d{4}-/.test(uploaded) ? new Date(uploaded.replace(' ', 'T') + 'Z') : uploaded, revisionOf, revisedTo];
    return h === HEADER ? row : h.map((name) => row[HEADER.indexOf(name)]);
  }));
}
function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}

// ------------------------------------------------------------------ a report

const ADHOC = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const WITH_SESSIONS = Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty(SESSIONS) });
const OPENING = { s1: { chair: 'Alex Organizer', minuteTakers: 'Sam Rivera, Casey Guest', note: 'A sensitive administrative note.' }, s2: { minuteTakers: 'Sam Rivera' }, s3: { note: 'Only a note here.' } };
const opening = (sessions) => JSON.stringify({ v: 1, sessions: sessions });

/**
 * One report in a sandbox: a document with a registration table and two TDoc
 * tables, the real TDoc-list reader, and every write recorded.
 * opts: { props, template, docId, tdocs, header, column (4 | 5 | 0), fetch }.
 */
function report(opts) {
  const o = opts || {};
  const loaded = o.template ? loadTemplateRuntime({ release: RELEASE, documentProperties: o.props || WITH_SESSIONS }) : loadCode({ documentProperties: o.props || WITH_SESSIONS });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const H2 = s.DocumentApp.ParagraphHeading.HEADING2;
  body.appendParagraph('1 Opening of the meeting').setHeading(H2);
  body.appendParagraph('The chair opens the session.');
  if (o.column !== 0) {
    const head = ['TDoc', 'Title', 'Source', 'Agenda Item'].concat(o.column === 5 ? ['Session'] : []);
    body.appendTable([head, ['S4aA269001', 'Synthetic title', 'ExampleCorp', '4.3'].concat(o.column === 5 ? ['22 Sep'] : [])]);
  }
  body.appendParagraph('4.3 Topic').setHeading(H2);
  body.appendTable([['TDoc', 'S4aA269001'], ['Title', 'Synthetic title'], ['TDoc Status', 'agreed'], ['Minutes', 'Presented and discussed.'], ['Disposition', 'agreed'], ['E-mail Discussion', 'No e-mail discussion']]);
  body.appendTable([['TDoc', 'S4aA269002'], ['Title', 'Synthetic title'], ['TDoc Status', 'available'], ['Minutes', ''], ['Disposition', ''], ['E-mail Discussion', '']]);
  body.appendParagraph('5 Close of the meeting').setHeading(H2);
  body.appendParagraph('The chair closes the meeting.');

  const r = { s: s, body: body, docProps: loaded.docProps, scriptProps: loaded.scriptProps, tdocs: o.tdocs || TDOCS, header: o.header, downloads: 0, logs: [], writes: [], reads: [], triggers: 0, userProps: 0 };
  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' } };
  ui.alert = (...args) => { ui.alerts.push(args); return 'YES'; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title: title, html: out.html }); };
  r.ui = ui;
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => o.docId || REPORT_ID, getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.Utilities.formatDate = (date, zone) => formatInZone(date, zone);
  s.Utilities.base64Decode = (b64) => Array.from(Buffer.from(b64, 'base64')).map((b) => (b > 127 ? b - 256 : b));
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.ScriptApp = { newTrigger: () => { r.triggers++; throw new Error('no trigger may be created'); }, getProjectTriggers: () => [] };
  s.PropertiesService.getUserProperties = () => { r.userProps++; return loaded.scriptProps; };
  s.UrlFetchApp = { fetch: o.fetch || (() => { r.downloads++; return { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }; }) };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const values = listValues(r.tdocs, r.header);
    const col = (r.header || HEADER).indexOf('Uploaded');
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((row) => row.map(() => null)) }),
      getRange: (row, c, rows) => ({ getDisplayValues: () => [['Uploaded']].concat(r.tdocs.map((t) => [col === -1 ? '' : t[3]])).slice(row - 1, row - 1 + rows) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };
  [loaded.docProps, loaded.scriptProps].forEach((store, i) => {
    const set = store.setProperty; const del = store.deleteProperty; const get = store.getProperty;
    store.setProperty = (k, v) => { r.writes.push('set ' + k); set(k, v); };
    store.deleteProperty = (k) => { r.writes.push('delete ' + k); del(k); };
    if (i === 0) store.getProperty = (k) => { r.reads.push(k); return get(k); };
  });

  r.model = () => s.collectAdhocSessionStatus_();
  r.lines = () => s.adhocSessionStatusLines_();
  r.status = () => { s.analyzeReportStatus(); return ui.alerts[ui.alerts.length - 1]; };
  r.statistics = () => { s.showAdhocSessionStatistics(); return ui.dialogs[ui.dialogs.length - 1]; };
  r.snapshot = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
  r.props = () => JSON.stringify([loaded.docProps._store, loaded.scriptProps._store]);
  /** Imports the synthetic Teams export for one session, through the real import. */
  r.importAttendance = (id) => { const b64 = FX.utf16le(FX.buildExport()).toString('base64'); return s.confirmTeamsAttendanceImport(id, b64, s.previewTeamsAttendanceImport(id, b64).token); };
  /** Stores attendance as the import does, from a compact value. */
  r.storeAttendance = (id, value) => s.writeAdhocAttendanceValue_(loaded.docProps, id, JSON.stringify(value));
  r.setTdocSessions = (clock, tdocs) => { loaded.docProps._store[TDOC_KEY] = s.serializeAdhocTdocSessions_({ clock: clock, tdocs: tdocs }); };
  return r;
}
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));
const overview = (model) => { const out = {}; model.overview.forEach((row) => { out[row[0]] = row[1]; }); return out; };
const notes = (model, level) => model.notes.filter((n) => n.level === level).map((n) => n.text);
const group = (model, id, title) => model.sessions.filter((x) => x.id === id)[0].groups.filter((g) => g.title === title)[0].rows;

/** Runs the statistics dialog's script against a minimal DOM; returns what it shows, as text. */
function runDialog(html) {
  const scripts = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
  const element = (tag) => { const e = { tag: tag, children: [], textContent: '', className: '' }; e.appendChild = (c) => { e.children.push(c); return c; }; return e; };
  const byId = {};
  (html.match(/id="([A-Za-z]+)"/g) || []).map((m) => m.slice(4, -1)).forEach((id) => { byId[id] = element(id); });
  let closed = 0;
  const ctx = { document: { getElementById: (id) => { if (!byId[id]) throw new Error('no element #' + id); return byId[id]; }, createElement: element }, google: { script: { host: { close: () => { closed++; } } } } };
  vm.createContext(ctx);
  vm.runInContext(scripts[0].replace(/^<script>|<\/script>$/g, ''), ctx);
  const rows = (node) => node.children.filter((c) => c.tag === 'table').reduce((all, table) => all.concat(table.children.map((tr) => tr.children.map((td) => td.textContent))), []);
  const text = (node) => [node.textContent].concat(node.children.map(text)).join('\n');
  return {
    scriptCount: scripts.length, el: byId, ctx: ctx, closed: () => closed,
    overview: rows(byId.overview),
    notes: byId.notes.children.map((c) => [c.className, c.textContent]),
    sessions: byId.sessions.children.map((box) => ({ when: box.children[0].textContent, groups: box.children.filter((c) => c.className === 'group').map((c) => c.textContent), rows: rows(box) })),
    allText: Object.keys(byId).map((id) => text(byId[id])).join('\n')
  };
}

const S = loadCode().sandbox;
const sessionsOf = (list) => { const l = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', ADHOC_SESSIONS: sessionsProperty(list, 10) } }); return l.sandbox.getAdhocSessions_(); };

// ================================================================ 1. the TDoc counts

console.log('1. the TDoc counts');
{
  const r = report();
  const all = () => r.s.flattenTdocGroups_(r.s.downloadAndGroupTdocs_(r.s.getReportConfig_()));
  const summary = () => r.s.summarizeAdhocTdocSessions_(r.s.makeAdhocTdocSessionResolver_(), all());
  const NONE = { 'after-last': 0, 'not-uploaded': 0, 'unreadable': 0, 'no-upload-time': 0, 'date-only': 0, 'stale': 0 };

  check('automatic only: every TDoc of the list once; seven with a session, three without, each for its reason', summary(), {
    total: 10, withSession: 7, severalSessions: 0, none: 3, explicitNone: 0, automatic: 7, manualAdd: 0, manualSet: 0,
    noneByReason: Object.assign({}, NONE, { 'after-last': 1, 'not-uploaded': 1, 'unreadable': 1 }), staleReferences: 0, notInList: 0,
    perSession: { s1: { assigned: 3, automatic: 3, manual: 0 }, s2: { assigned: 3, automatic: 3, manual: 0 }, s3: { assigned: 1, automatic: 1, manual: 0 } } });
  check('the counts are the ones of the Session column: the same resolver, TDoc by TDoc', (() => {
    const resolver = r.s.makeAdhocTdocSessionResolver_();
    const per = { s1: 0, s2: 0, s3: 0 };
    all().forEach((td) => resolver.effective(td).forEach((id) => { per[id]++; }));
    return [per, (SECTION.match(/assignAdhocSession_\(|adhocSessionCutoff_\(|convertUtcWallClockToReportZone_\(|\.uploaded/g) || [])];
  })(), [{ s1: 3, s2: 3, s3: 1 }, []]);
  check('a revision is a TDoc of its own: the revision of a TDoc of 22 Sep, uploaded a day later, counts in Offline', (() => {
    const resolver = r.s.makeAdhocTdocSessionResolver_();
    const byId = {}; all().forEach((td) => { byId[td.row[0]] = resolver.effective(td); });
    return [byId.S4aA269002, byId.S4aA269010];
  })(), [['s1'], ['s2']]);

  // Manual assignments: add, set, set to none, several sessions.
  r.setTdocSessions('utc', {
    S4aA269001: { mode: 'add', sessions: ['s3'] },         // s1 by upload time, and s3
    S4aA269004: { mode: 'set', sessions: ['s1'] },         // s1 instead of s2
    S4aA269005: { mode: 'set', sessions: [] },             // no session
    S4aA269007: { mode: 'set', sessions: ['s2', 's3'] },   // uploaded after the last session; put into two
    S4aA269008: { mode: 'add', sessions: ['s1'] }          // not uploaded; put into s1
  });
  const manual = summary();
  check('manual assignments: two added, three set; four TDocs left with their session by upload time', [manual.manualAdd, manual.manualSet, manual.automatic], [2, 3, 4]);
  check('eight TDocs have a session, two of them more than one; two have none, one of them set to none', [manual.withSession, manual.severalSessions, manual.none, manual.explicitNone, manual.noneByReason], [8, 2, 2, 1, Object.assign({}, NONE, { unreadable: 1 })]);
  check('per session: a TDoc counts in every session it is in, by upload time or by manual assignment', manual.perSession,
    { s1: { assigned: 5, automatic: 3, manual: 2 }, s2: { assigned: 2, automatic: 1, manual: 1 }, s3: { assigned: 3, automatic: 1, manual: 2 } });
  check('the per-session counts add up to the TDocs with a session plus the additional sessions of those in two', [5 + 2 + 3, manual.withSession + 2, manual.withSession + manual.none === manual.total], [10, 10, true]);
  check('in each session: by upload time + by manual assignment = assigned', Object.keys(manual.perSession).map((id) => manual.perSession[id].automatic + manual.perSession[id].manual === manual.perSession[id].assigned), [true, true, true]);
  check('adding the session a TDoc has by its upload time anyway keeps it "by upload time"', (() => {
    r.setTdocSessions('utc', { S4aA269002: { mode: 'add', sessions: ['s1'] } });
    const x = summary();
    return [x.perSession.s1, x.manualAdd, x.automatic, x.severalSessions];
  })(), [{ assigned: 3, automatic: 3, manual: 0 }, 1, 6, 0]);
  check('setting a TDoc to the session it has by its upload time makes it a manual one', (() => {
    r.setTdocSessions('utc', { S4aA269002: { mode: 'set', sessions: ['s1'] } });
    return summary().perSession.s1;
  })(), { assigned: 3, automatic: 2, manual: 1 });

  // The clock of the upload times.
  r.setTdocSessions('session', {});
  const compatibility = summary();
  check('with the compatibility clock the upload times are read as they stand: other sessions than with UTC', [compatibility.perSession, compatibility.withSession, compatibility.noneByReason['after-last']],
    [{ s1: { assigned: 4, automatic: 4, manual: 0 }, s2: { assigned: 3, automatic: 3, manual: 0 }, s3: { assigned: 0, automatic: 0, manual: 0 } }, 7, 1]);

  // References that cannot be resolved.
  r.setTdocSessions('utc', { S4aA269001: { mode: 'set', sessions: ['s9'] }, S4aA269002: { mode: 'add', sessions: ['s1', 's9'] }, S4aA269099: { mode: 'set', sessions: ['s1'] } });
  const stale = summary();
  check('a manual assignment that names a session that is not configured, and one for a TDoc that is not in the list', [stale.staleReferences, stale.notInList, stale.noneByReason.stale, stale.none, stale.total, stale.perSession.s1], [2, 1, 1, 4, 10, { assigned: 2, automatic: 2, manual: 0 }]);

  // Lists without usable upload times.
  const noColumn = report({ header: HEADER.filter((h) => h !== 'Uploaded') });
  const nc = noColumn.s.summarizeAdhocTdocSessions_(noColumn.s.makeAdhocTdocSessionResolver_(), noColumn.s.flattenTdocGroups_(noColumn.s.downloadAndGroupTdocs_(noColumn.s.getReportConfig_())));
  check('a TDoc list without an Uploaded column: no TDoc has a session, and that is the reason', [nc.total, nc.withSession, nc.none, nc.noneByReason['no-upload-time']], [10, 0, 10, 10]);
  const empty = report({ tdocs: [] });
  check('an empty TDoc list', (() => { const x = empty.s.summarizeAdhocTdocSessions_(empty.s.makeAdhocTdocSessionResolver_(), []); return [x.total, x.withSession, x.none, x.perSession.s1]; })(), [0, 0, 0, { assigned: 0, automatic: 0, manual: 0 }]);
}

// ================================================================ 2. the status model

const PARSED = (raw) => S.parseAdhocOpeningProperty_(raw);
const NO_ATTENDANCE = { status: 'absent', error: null, sessions: {}, raw: {}, problems: [], companies: {} };
const NO_TDOCS = { available: false, reason: 'not asked for', column: 4, clock: 'utc', storedStatus: 'absent', storedError: null, summary: null };

console.log('2. the status model: sessions, opening details, attendance');
{
  const build = (list, openingRaw, attendance, tdocs) => S.buildAdhocSessionStatusModel_(sessionsOf(list), PARSED(openingRaw || null), attendance || NO_ATTENDANCE, tdocs || NO_TDOCS);

  // Sessions.
  check('the dates of the sessions: one date, or the first and the last', [S.adhocSessionDateRangeText_(sessionsOf([S1])), S.adhocSessionDateRangeText_(sessionsOf(SESSIONS)), S.adhocSessionDateRangeText_(sessionsOf([S3, S1])),
    S.adhocSessionDateRangeText_(sessionsOf([S1, session('s2', '2026-09-22', '19:00', '20:00', 'Evening')])), S.adhocSessionDateRangeText_([])],
    ['September 22, 2026', RANGE, RANGE, 'September 22, 2026', '']);
  check('whatever order they are given in', [S.adhocSessionDateRangeText_([S3, S1, S2]), S.adhocSessionDateRangeText_([S2, S3])], [RANGE, 'September 23–24, 2026']);
  const one = build([S1]);
  check('one session with both times: one session, its date, no remark about planned times', [one.count, one.range, overview(one).Sessions, notes(one, 'info').filter((t) => /Planned/.test(t))], [1, 'September 22, 2026', '1 (September 22, 2026)', []]);
  const three = build(SESSIONS);
  check('three sessions', [three.count, overview(three).Sessions], [3, '3 (' + RANGE + ')']);
  check('sessions without a planned start or end are named, with what is missing -- as a remark, not a warning', [notes(three, 'info')[0], notes(three, 'warning').filter((t) => /Planned/.test(t))],
    ['Planned times incomplete: Offline (no start, no end), 24 Sep (no end). A session without an end takes the TDocs uploaded until the end of its day.', []]);
  check('a session with an end and no start', notes(build([session('s1', '2026-09-22', '', '18:00')]), 'info')[0].slice(0, 44), 'Planned times incomplete: 22 Sep (no start).');
  check('the sessions are in chronological order, whatever the order they are stored in', build([session('s1', '2026-09-24', '15:00', '18:00', 'Last'), session('s2', '2026-09-22', '15:00', '18:00', 'First'), session('s3', '2026-09-23', '09:00', '10:00')]).sessions.map((x) => [x.id, x.when]),
    [['s2', 'First, September 22, 2026, 15:00–18:00'], ['s3', 'September 23, 2026, 09:00–10:00'], ['s1', 'Last, September 24, 2026, 15:00–18:00']]);
  check('each session: its label, date and planned times as they are configured now', [three.sessions.map((x) => x.when), group(three, 's1', 'Planned'), group(three, 's2', 'Planned'), group(three, 's3', 'Planned')],
    [['September 22, 2026, 15:00–18:00', 'Offline, September 23, 2026', 'September 24, 2026, 15:00'], [['Date', 'September 22, 2026'], ['Start', '15:00'], ['End', '18:00']],
      [['Date', 'September 23, 2026'], ['Start', 'not planned'], ['End', 'not planned']], [['Date', 'September 24, 2026'], ['Start', '15:00'], ['End', 'not planned']]]);
  check('every session has the same four groups, in the same order', three.sessions.map((x) => x.groups.map((g) => g.title).join(' | ')).filter((t, i, list) => list.indexOf(t) === i), ['Planned | Opening | Attendance (Microsoft Teams) | TDocs']);

  // Opening details.
  const none = build(SESSIONS);
  check('no opening details: 0 of 3; Chair and minute taker(s) missing everywhere, as remarks', [overview(none)['Opening details'], none.opening, notes(none, 'info').filter((t) => /Chair|Minute/.test(t))],
    ['0 of 3 sessions', { known: true, withDetails: 0, noChair: 3, noMinuteTakers: 3, stale: 0 }, ['Chair not entered: 22 Sep, Offline, 24 Sep.', 'Minute taker(s) not entered: 22 Sep, Offline, 24 Sep.']]);
  const partial = build(SESSIONS, opening({ s2: OPENING.s2 }));
  check('one session with details: 1 of 3', [overview(partial)['Opening details'], partial.opening.withDetails, notes(partial, 'info').filter((t) => /Chair|Minute/.test(t))],
    ['1 of 3 sessions', 1, ['Chair not entered: 22 Sep, Offline, 24 Sep.', 'Minute taker(s) not entered: 22 Sep, 24 Sep.']]);
  const full = build(SESSIONS, opening(OPENING));
  check('details for every session: 3 of 3 -- a note alone counts; what is missing is still said per field', [overview(full)['Opening details'], full.opening, notes(full, 'info').filter((t) => /Chair|Minute/.test(t))],
    ['3 of 3 sessions', { known: true, withDetails: 3, noChair: 2, noMinuteTakers: 1, stale: 0 }, ['Chair not entered: Offline, 24 Sep.', 'Minute taker(s) not entered: 24 Sep.']]);
  check('Chair and minute taker(s) everywhere: no remark', notes(build(SESSIONS, opening({ s1: { chair: 'A', minuteTakers: 'B' }, s2: { chair: 'A', minuteTakers: 'B' }, s3: { chair: 'A', minuteTakers: 'B' } })), 'info').filter((t) => /Chair|Minute/.test(t)), []);
  check('per session: whether each field is filled in, and nothing of what it says', [group(full, 's1', 'Opening'), group(full, 's2', 'Opening'), group(full, 's3', 'Opening')],
    [[['Chair', 'entered'], ['Minute taker(s)', 'entered'], ['Administrative note', 'entered']], [['Chair', 'not entered'], ['Minute taker(s)', 'entered'], ['Administrative note', 'none']],
      [['Chair', 'not entered'], ['Minute taker(s)', 'not entered'], ['Administrative note', 'entered']]]);
  check('no name and no note is in the model', /Organizer|Rivera|Casey|sensitive|Only a note/.test(JSON.stringify(full)), false);
  check('missing opening details are never a warning', [notes(none, 'warning'), notes(partial, 'warning')], [['The TDoc sessions are not available: the TDoc list could not be read (not asked for).'], ['The TDoc sessions are not available: the TDoc list could not be read (not asked for).']]);
  const corrupt = build(SESSIONS, '{"v":1,');
  check('opening details that cannot be read: not available, a warning, and no guess about what is missing', [overview(corrupt)['Opening details'], corrupt.opening.known, notes(corrupt, 'warning')[0], notes(corrupt, 'info').filter((t) => /Chair|Minute/.test(t)), group(corrupt, 's1', 'Opening')],
    ['not available', false, 'The stored opening details cannot be read. The status of the opening details is not available.', [], [['Opening details', 'not available']]]);
  const future = build(SESSIONS, '{"v":2,"sessions":{"s1":{"chair":"Alex Organizer"}}}');
  check('opening details of a newer version: not available, with the versions', [overview(future)['Opening details'], notes(future, 'warning')[0], /Organizer/.test(JSON.stringify(future))],
    ['not available', 'The stored opening details have version 2; this release reads version 1. The status of the opening details is not available.', false]);
  check('the rest of the status is there all the same', [overview(corrupt).Sessions, overview(corrupt)['Attendance imported'], overview(future)['Attendance imported']], ['3 (' + RANGE + ')', '0 of 3 sessions', '0 of 3 sessions']);
  check('details stored for a session that is not configured: a warning', notes(build(SESSIONS, opening({ s1: OPENING.s1, s9: { chair: 'Nobody' } })), 'warning')[0], 'Opening details are stored for 1 session that is not configured.');

  // Attendance.
  const model = (extra) => Object.assign({ meetingTitle: 'Synthetic title', attendanceRecords: 25, start: '2026-09-22T14:48:05', end: '2026-09-22T18:11:26', durationSeconds: 12200, averageAttendanceSeconds: 9067,
    attendees: [{ name: 'Alex Organizer', company: '', email: 'alex.organizer@example.com' }, { name: 'Sam Rivera', company: 'ExampleCorp', email: '' }] }, extra || {});
  const state = (sessions, problems) => ({ status: 'ok', error: null, sessions: sessions, raw: {}, problems: problems || [], companies: {} });
  check('no attendance: 0 of 3, and the sessions without it are named -- as a remark', [overview(none)['Attendance imported'], none.attendance, notes(none, 'info').filter((t) => /Attendance/.test(t)), group(none, 's1', 'Attendance (Microsoft Teams)')],
    ['0 of 3 sessions', { known: true, imported: 0, missing: 3, damaged: 0, stale: 0, records: 0 }, ['Attendance not imported: 22 Sep, Offline, 24 Sep.'], [['Attendance', 'not imported']]]);
  const some = build(SESSIONS, null, state({ s1: model() }));
  check('attendance for one session: 1 of 3', [overview(some)['Attendance imported'], some.attendance, notes(some, 'info').filter((t) => /Attendance/.test(t))],
    ['1 of 3 sessions', { known: true, imported: 1, missing: 2, damaged: 0, stale: 0, records: 1 }, ['Attendance not imported: Offline, 24 Sep.']]);
  check('the values are the ones Teams states, as the Attendance section of the report shows them; then the number of attendees listed', group(some, 's1', 'Attendance (Microsoft Teams)'),
    [['Meeting', 'Synthetic title'], ['Start', '14:48:05'], ['End', '18:11:26'], ['Duration', '3:23:20'], ['Attendance records', '25'], ['Average attendance', '2:31:07'], ['Attendees listed', '2']]);
  check('they are the lines of the report\'s statistics block, taken from the same function', group(some, 's1', 'Attendance (Microsoft Teams)').slice(0, -1),
    S.buildAdhocAttendanceBlocks_(sessionsOf(SESSIONS), state({ s1: model() }))[0].statistics.filter((line) => line[0] !== 'Session'));
  check('Teams\' duration and average are shown as stated: 3:23:20 although end minus start is 3:23:21; records 25 although 2 attendees are listed', (() => {
    const rows = {}; group(some, 's1', 'Attendance (Microsoft Teams)').forEach((row) => { rows[row[0]] = row[1]; });
    return [rows.Duration, rows['Attendance records'], rows['Average attendance'], rows['Attendees listed']];
  })(), ['3:23:20', '25', '2:31:07', '2']);
  const sparse = build(SESSIONS, null, state({ s2: model({ attendanceRecords: null, end: null, durationSeconds: null, averageAttendanceSeconds: null }) }));
  check('a value Teams did not state is left out, never worked out', group(sparse, 's2', 'Attendance (Microsoft Teams)'), [['Meeting', 'Synthetic title'], ['Start', '14:48:05'], ['Attendees listed', '2']]);
  check('an end on another day is shown with its date', group(build(SESSIONS, null, state({ s1: model({ end: '2026-09-23T00:10:00' }) })), 's1', 'Attendance (Microsoft Teams)')[2], ['End', '2026-09-23 00:10:00']);
  const all = build(SESSIONS, null, state({ s1: model(), s2: model(), s3: model({ attendees: [] }) }));
  check('attendance for every session: 3 of 3, no remark; a session without attendees shows 0', [overview(all)['Attendance imported'], notes(all, 'info').filter((t) => /Attendance/.test(t)), group(all, 's3', 'Attendance (Microsoft Teams)').slice(-1)],
    ['3 of 3 sessions', [], [['Attendees listed', '0']]]);
  check('no attendee name, e-mail or company is in the model', /Organizer|Rivera|example\.com|ExampleCorp/.test(JSON.stringify(all)), false);
  const damaged = build(SESSIONS, null, state({ s1: model() }, ['s2']));
  check('one session whose stored attendance cannot be read: it is named in a warning; the others are counted and shown',
    [overview(damaged)['Attendance imported'], damaged.attendance, notes(damaged, 'warning')[0], notes(damaged, 'info').filter((t) => /Attendance/.test(t)), group(damaged, 's2', 'Attendance (Microsoft Teams)'), group(damaged, 's1', 'Attendance (Microsoft Teams)').length],
    ['1 of 3 sessions', { known: true, imported: 1, missing: 1, damaged: 1, stale: 0, records: 1 }, 'The stored attendance of Offline cannot be read. Import it again or remove it.', ['Attendance not imported: 24 Sep.'], [['Attendance', 'stored, but cannot be read']], 7]);
  check('attendance stored for a session that is not configured: a warning', notes(build(SESSIONS, null, state({ s1: model(), s9: model() })), 'warning')[0], 'Attendance is stored for 1 session that is not configured.');
  ['invalid', 'unsupported'].forEach((status) => {
    const x = build(SESSIONS, opening(OPENING), { status: status, error: status === 'invalid' ? 'The stored attendance index cannot be read.' : 'The stored attendance has version 2; this release reads version 1.', sessions: {}, raw: {}, problems: [], companies: {} });
    check(`an attendance index that is ${status}: not available, a warning; the opening details are still counted`, [overview(x)['Attendance imported'], x.attendance.known, notes(x, 'warning')[0], group(x, 's1', 'Attendance (Microsoft Teams)'), overview(x)['Opening details']],
      ['not available', false, (status === 'invalid' ? 'The stored attendance index cannot be read.' : 'The stored attendance has version 2; this release reads version 1.') + ' The status of the attendance is not available.', [['Attendance', 'not available']], '3 of 3 sessions']);
  });

  // The TDoc part of the model.
  const summary = { total: 10, withSession: 8, severalSessions: 2, none: 2, explicitNone: 1, automatic: 4, manualAdd: 2, manualSet: 3,
    noneByReason: { 'after-last': 0, 'not-uploaded': 0, 'unreadable': 1, 'no-upload-time': 0, 'date-only': 0, 'stale': 0 }, staleReferences: 0, notInList: 0,
    perSession: { s1: { assigned: 5, automatic: 3, manual: 2 }, s2: { assigned: 2, automatic: 1, manual: 1 }, s3: { assigned: 3, automatic: 1, manual: 2 } } };
  const tdocs = (extra, summaryExtra) => Object.assign({ available: true, reason: '', column: 5, clock: 'utc', storedStatus: 'ok', storedError: null, summary: Object.assign({}, summary, summaryExtra || {}) }, extra || {});
  const withTdocs = build(SESSIONS, null, null, tdocs());
  check('the TDoc rows of the overview', withTdocs.overview.slice(3), [['TDocs with a session', '8 of 10'], ['– by upload time', '4'], ['– with a manual assignment', '5 (2 added, 3 set)'], ['– in more than one session', '2'],
    ['TDocs without a session', '2'], ['– set to no session', '1'], ['Session column', 'in the registration table']]);
  check('per session', [group(withTdocs, 's1', 'TDocs'), group(withTdocs, 's3', 'TDocs')], [[['In this session', '5'], ['– by upload time', '3'], ['– by manual assignment', '2']], [['In this session', '3'], ['– by upload time', '1'], ['– by manual assignment', '2']]]);
  check('the Session column', [4, 5, 0, -1].map((column) => overview(build(SESSIONS, null, null, tdocs({ column: column })))['Session column']),
    ['not enabled (Assign TDoc Sessions… can add it)', 'in the registration table', 'the report has no registration table yet', 'not known']);
  check('a Session column that is not enabled is not a warning', notes(build(SESSIONS, null, null, tdocs({ column: 4 }, { noneByReason: summary.noneByReason, none: 0 })), 'warning').filter((t) => /column/i.test(t)), []);
  const reasons = (r, extra) => build(SESSIONS, null, null, tdocs(null, Object.assign({ noneByReason: Object.assign({}, summary.noneByReason, { unreadable: 0 }, r) }, extra || {})));
  check('uploaded after the last session, without a session: a warning', [notes(reasons({ 'after-last': 1 }), 'warning'), notes(reasons({ 'after-last': 3 }), 'warning')],
    [['1 TDoc was uploaded after the end of the last session and has no session.'], ['3 TDocs were uploaded after the end of the last session and have no session.']]);
  check('an upload time that cannot be read, and a list without upload times: warnings', [notes(reasons({ unreadable: 2 }), 'warning'), notes(reasons({ 'no-upload-time': 10 }), 'warning')],
    [['The upload time of 2 TDocs cannot be read; no session.'], ['The TDoc list has no upload times: 10 TDocs without a session.']]);
  check('a manual assignment that names a session that is not configured: a warning', notes(reasons({ stale: 1 }, { staleReferences: 2 }), 'warning'), ['2 manual assignments name a session that is not configured.']);
  check('not uploaded yet, a date without a time, an assignment for a TDoc not in the list: remarks, not warnings', (() => {
    const x = reasons({ 'not-uploaded': 2, 'date-only': 1 }, { notInList: 1 });
    return [notes(x, 'warning'), notes(x, 'info').slice(-3)];
  })(), [[], ['2 TDocs are not uploaded yet and have no session.', '1 TDoc has an upload date without a time, on the day of a cut-off; no session.', '1 manual assignment is for a TDoc that is not in the TDoc list now.']]);
  check('a TDoc set to no session is neither a warning nor a remark: it is what was asked for', (() => { const x = reasons({}, { explicitNone: 4, none: 4 }); return [notes(x, 'warning'), notes(x, 'info').filter((t) => /TDoc/.test(t) && !/until the end of its day/.test(t))]; })(), [[], []]);
  check('the clock is mentioned only when it is not the default', [notes(build(SESSIONS, null, null, tdocs()), 'info').filter((t) => /clock/.test(t)), notes(build(SESSIONS, null, null, tdocs({ clock: 'session' })), 'info').filter((t) => /clock/.test(t))],
    [[], ['The upload times are read on the clock of the sessions, not as UTC (compatibility setting of Assign TDoc Sessions…).']]);
  ['invalid', 'unsupported'].forEach((status) => {
    const error = status === 'invalid' ? 'The stored manual TDoc session assignments cannot be read.' : 'The stored manual TDoc session assignments have version 2; this release reads version 1.';
    const x = build(SESSIONS, null, null, tdocs({ storedStatus: status, storedError: error }));
    check(`manual assignments that are ${status}: a warning that only the sessions by upload time are counted; the counts are shown`, [notes(x, 'warning')[0], overview(x)['TDocs with a session']], [error + ' Only the sessions by upload time are counted.', '8 of 10']);
  });
  const unavailable = build(SESSIONS, opening(OPENING), state({ s1: model() }), { available: false, reason: 'Could not download TDOC list: HTTP 500', column: 5, clock: 'utc', storedStatus: 'absent', storedError: null, summary: null });
  check('TDoc sessions that are not available: one row, one warning with the reason; everything else as usual', [unavailable.overview.map((row) => row[0]), overview(unavailable)['TDoc sessions'], notes(unavailable, 'warning'), group(unavailable, 's1', 'TDocs'), unavailable.tdocs],
    [['Sessions', 'Opening details', 'Attendance imported', 'TDoc sessions', 'Session column'], 'not available', ['The TDoc sessions are not available: the TDoc list could not be read (Could not download TDOC list: HTTP 500).'], [['TDoc sessions', 'not available']], null]);

  // The block of the status summary.
  check('the block: a heading, the overview, the warnings, the remarks, and where the figures per session are', S.formatAdhocSessionStatusLines_(unavailable), [HEAD,
    '   Sessions: 3 (' + RANGE + ')', '   Opening details: 3 of 3 sessions', '   Attendance imported: 1 of 3 sessions', '   TDoc sessions: not available', '   Session column: in the registration table',
    '', '   ⚠️  The TDoc sessions are not available: the TDoc list could not be read (Could not download TDOC list: HTTP 500).',
    '', '   ℹ️  Planned times incomplete: Offline (no start, no end), 24 Sep (no end). A session without an end takes the TDocs uploaded until the end of its day.',
    '   ℹ️  Chair not entered: Offline, 24 Sep.', '   ℹ️  Minute taker(s) not entered: 24 Sep.', '   ℹ️  Attendance not imported: Offline, 24 Sep.', POINTER]);
  const clean = S.buildAdhocSessionStatusModel_(sessionsOf([S1]), PARSED(opening({ s1: { chair: 'A', minuteTakers: 'B' } })), state({ s1: model() }),
    { available: true, reason: '', column: 5, clock: 'utc', storedStatus: 'absent', storedError: null, summary: Object.assign({}, summary, { total: 3, withSession: 3, severalSessions: 0, none: 0, explicitNone: 0, automatic: 3, manualAdd: 0, manualSet: 0,
      noneByReason: Object.assign({}, summary.noneByReason, { unreadable: 0 }), perSession: { s1: { assigned: 3, automatic: 3, manual: 0 } } }) });
  check('with nothing missing and nothing wrong there are no warnings and no remarks', [clean.notes, S.formatAdhocSessionStatusLines_(clean)], [[], [HEAD, '   Sessions: 1 (September 22, 2026)', '   Opening details: 1 of 1 sessions', '   Attendance imported: 1 of 1 sessions',
    '   TDocs with a session: 3 of 3', '   – by upload time: 3', '   – with a manual assignment: 0 (0 added, 0 set)', '   – in more than one session: 0', '   TDocs without a session: 0', '   – set to no session: 0', '   Session column: in the registration table', POINTER]]);
  check('there is no grade, no score and no percentage', /%|score|grade|quality|✅|❌/i.test(JSON.stringify([unavailable, clean, withTdocs]) + S.formatAdhocSessionStatusLines_(withTdocs).join('\n')), false);
}

// ================================================================ 3. Report Status Summary

const gitShow = (spec) => { try { return execFileSync('git', ['show', spec], { cwd: path.join(__dirname, '..'), encoding: 'utf8', maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).replace(/\r/g, ''); } catch (e) { return null; } };
const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
const OLD_CODE = gitShow('template-release/T-2026.10.4:Code.js');
/** The status text of T-2026.10.4 for this report, run in the same sandbox. */
function oldStatus(r) {
  vm.runInContext(functionSource(OLD_CODE, 'analyzeReportStatus').replace('function analyzeReportStatus(', 'function analyzeReportStatusOfBaseline('), r.s);
  r.s.analyzeReportStatusOfBaseline();
  return r.ui.alerts.pop();
}
const OLD_TEXT_OF_THIS_DOCUMENT = (tables) => ['📊 REPORT STATUS ANALYSIS\n', '═══════════════════════════════════\n', `\n📈 Total Documents: ${tables}\n`, '\n📊 Status Breakdown:'].join('\n');

/** A report with opening details, attendance for two sessions and manual TDoc assignments. */
function fullReport(opts) {
  const r = report(Object.assign({ column: 5, props: Object.assign({}, WITH_SESSIONS, { [OPENING_KEY]: opening(OPENING) }) }, opts || {}));
  r.importAttendance('s1');
  r.storeAttendance('s2', { t: 'Synthetic second day', n: null, s: '2026-09-23T09:00:00', e: null, d: null, a: null, p: [['Pat Kim', 'ExampleCorp', 'pat.kim@example.com'], ['Robin Lee', '', '']] });
  r.setTdocSessions('utc', { S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269004: { mode: 'set', sessions: ['s1'] }, S4aA269005: { mode: 'set', sessions: [] },
    S4aA269007: { mode: 'set', sessions: ['s2', 's3'] }, S4aA269008: { mode: 'add', sessions: ['s1'] } });
  r.writes.length = 0;
  r.logs.length = 0;
  return r;
}
const FULL_BLOCK = [HEAD,
  '   Sessions: 3 (' + RANGE + ')', '   Opening details: 3 of 3 sessions', '   Attendance imported: 2 of 3 sessions',
  '   TDocs with a session: 8 of 10', '   – by upload time: 4', '   – with a manual assignment: 5 (2 added, 3 set)', '   – in more than one session: 2', '   TDocs without a session: 2', '   – set to no session: 1',
  '   Session column: in the registration table',
  '', '   ⚠️  The upload time of 1 TDoc cannot be read; no session.',
  '', '   ℹ️  Planned times incomplete: Offline (no start, no end), 24 Sep (no end). A session without an end takes the TDocs uploaded until the end of its day.',
  '   ℹ️  Chair not entered: Offline, 24 Sep.', '   ℹ️  Minute taker(s) not entered: 24 Sep.', '   ℹ️  Attendance not imported: 24 Sep.', POINTER];

console.log('3. Report Status Summary');
{
  const r = report();
  const alert = r.status();
  const block = r.lines();
  check('a new ad-hoc report with sessions: the block', block, [HEAD,
    '   Sessions: 3 (' + RANGE + ')', '   Opening details: 0 of 3 sessions', '   Attendance imported: 0 of 3 sessions',
    '   TDocs with a session: 7 of 10', '   – by upload time: 7', '   – with a manual assignment: 0 (0 added, 0 set)', '   – in more than one session: 0', '   TDocs without a session: 3', '   – set to no session: 0',
    '   Session column: not enabled (Assign TDoc Sessions… can add it)',
    '', '   ⚠️  1 TDoc was uploaded after the end of the last session and has no session.', '   ⚠️  The upload time of 1 TDoc cannot be read; no session.',
    '', '   ℹ️  Planned times incomplete: Offline (no start, no end), 24 Sep (no end). A session without an end takes the TDocs uploaded until the end of its day.',
    '   ℹ️  Chair not entered: 22 Sep, Offline, 24 Sep.', '   ℹ️  Minute taker(s) not entered: 22 Sep, Offline, 24 Sep.', '   ℹ️  Attendance not imported: 22 Sep, Offline, 24 Sep.',
    '   ℹ️  1 TDoc is not uploaded yet and has no session.', POINTER]);
  check('it is one alert, titled Report Status, with the status of the TDoc tables first', [r.ui.alerts.length, alert[0], alert[1].indexOf(OLD_TEXT_OF_THIS_DOCUMENT(3)), alert[2]], [1, 'Report Status', 0, 'OK']);
  check('the block is at its end', alert[1].slice(-block.join('\n').length) === block.join('\n'), true);
  if (OLD_CODE) {
    check('the alert is the one of T-2026.10.4 for this document, and then the block', alert[1] === oldStatus(r)[1] + '\n' + block.join('\n'), true);
  } else {
    console.log('  note: template-release/T-2026.10.4 is not available in this checkout; the comparisons with it are skipped.');
  }
  check('missing attendance and missing opening details are remarks; none of them is a warning', block.filter((l) => /⚠️/.test(l)).filter((l) => /Attendance|Chair|Minute|column/i.test(l)), []);

  const full = fullReport();
  check('a report with opening details, attendance and manual TDoc assignments', full.lines(), FULL_BLOCK);
  check('the same through the menu action', (() => { const a = full.status(); return a[1].slice(-FULL_BLOCK.join('\n').length) === FULL_BLOCK.join('\n'); })(), true);
  check('the Session column: enabled, not enabled, no registration table', [5, 4, 0].map((column) => report({ column: column }).lines().filter((l) => /Session column/.test(l))[0]),
    ['   Session column: in the registration table', '   Session column: not enabled (Assign TDoc Sessions… can add it)', '   Session column: the report has no registration table yet']);

  const clock = report();
  clock.setTdocSessions('session', {});
  check('the compatibility clock: the counts are the ones of that clock, and the block says which clock it is', [clock.model().tdocs.perSession, clock.lines().filter((l) => /clock/.test(l))],
    [{ s1: { assigned: 4, automatic: 4, manual: 0 }, s2: { assigned: 3, automatic: 3, manual: 0 }, s3: { assigned: 0, automatic: 0, manual: 0 } },
      ['   ℹ️  The upload times are read on the clock of the sessions, not as UTC (compatibility setting of Assign TDoc Sessions…).']]);
  check('with the default clock (UTC) nothing is said about the clock', report().lines().filter((l) => /clock|UTC/.test(l)), []);

  const one = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([S1]) }) });
  check('one session: the TDocs uploaded after its end have no session', [one.lines()[1], one.lines()[4], one.lines().filter((l) => /after the end/.test(l))],
    ['   Sessions: 1 (September 22, 2026)', '   TDocs with a session: 3 of 10', ['   ⚠️  5 TDocs were uploaded after the end of the last session and have no session.']]);
  const renamed = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-10-01', '09:00', '10:00', 'Kick-off'), session('s2', '2026-10-05', '', '', 'Wrap-up')]) }) });
  check('the current labels and dates of the sessions are used', [renamed.lines()[1], renamed.lines().filter((l) => /Planned|Chair/.test(l))],
    ['   Sessions: 2 (October 1–5, 2026)', ['   ℹ️  Planned times incomplete: Wrap-up (no start, no end). A session without an end takes the TDocs uploaded until the end of its day.', '   ℹ️  Chair not entered: Kick-off, Wrap-up.']]);
}

// ================================================================ 4. Post-meeting Statistics…

console.log('4. Post-meeting Statistics…');
{
  const r = fullReport();
  const dialog = r.statistics();
  check('one dialog, titled Post-meeting Statistics; no alert', [r.ui.dialogs.length, dialog.title, r.ui.alerts.length], [1, 'Post-meeting Statistics', 0]);
  const d = runDialog(dialog.html);
  check('all sessions: the rows of the status block', d.overview, [['Sessions', '3 (' + RANGE + ')'], ['Opening details', '3 of 3 sessions'], ['Attendance imported', '2 of 3 sessions'], ['TDocs with a session', '8 of 10'], ['– by upload time', '4'],
    ['– with a manual assignment', '5 (2 added, 3 set)'], ['– in more than one session', '2'], ['TDocs without a session', '2'], ['– set to no session', '1'], ['Session column', 'in the registration table']]);
  check('the same warnings and remarks, marked as what they are', d.notes, [['note', 'Planned times incomplete: Offline (no start, no end), 24 Sep (no end). A session without an end takes the TDocs uploaded until the end of its day.'],
    ['note', 'Chair not entered: Offline, 24 Sep.'], ['note', 'Minute taker(s) not entered: 24 Sep.'], ['note', 'Attendance not imported: 24 Sep.'], ['note warning', 'The upload time of 1 TDoc cannot be read; no session.']]);
  check('one block per session, in chronological order, with its four groups', d.sessions.map((x) => [x.when, x.groups.join(' | ')]), [['September 22, 2026, 15:00–18:00', 'Planned | Opening | Attendance (Microsoft Teams) | TDocs'],
    ['Offline, September 23, 2026', 'Planned | Opening | Attendance (Microsoft Teams) | TDocs'], ['September 24, 2026, 15:00', 'Planned | Opening | Attendance (Microsoft Teams) | TDocs']]);
  check('the first session: planned, opening, the values Teams states, and its TDocs', d.sessions[0].rows, [['Date', 'September 22, 2026'], ['Start', '15:00'], ['End', '18:00'],
    ['Chair', 'entered'], ['Minute taker(s)', 'entered'], ['Administrative note', 'entered'],
    ['Meeting', 'Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE'], ['Start', '14:48:05'], ['End', '18:11:26'], ['Duration', '3:23:20'], ['Attendance records', '25'], ['Average attendance', '2:31:07'], ['Attendees listed', '23'],
    ['In this session', '5'], ['– by upload time', '3'], ['– by manual assignment', '2']]);
  check('Teams\' values are the stated ones (duration 3h 23m 20s, average 2h 31m 7s, 25 records), not what the rows of the export would give', (() => {
    const stored = r.s.readAdhocAttendance_(r.docProps).sessions.s1;
    return [stored.durationSeconds, stored.averageAttendanceSeconds, stored.attendanceRecords, stored.attendees.length, FX.SUMMARY.meetingDuration, FX.SUMMARY.averageAttendanceTime, FX.SUMMARY.attendedParticipants];
  })(), [12200, 9067, 25, 23, '3h 23m 20s', '2h 31m 7s', '25']);
  check('the number of attendees listed is the number of rows of the attendee table of the report', (() => {
    const table = r.body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getCell(0).getText() === 'Name')[0];
    return table.getNumRows() - 1;
  })(), 23);
  check('the second session: values Teams did not state are left out', d.sessions[1].rows, [['Date', 'September 23, 2026'], ['Start', 'not planned'], ['End', 'not planned'],
    ['Chair', 'not entered'], ['Minute taker(s)', 'entered'], ['Administrative note', 'none'],
    ['Meeting', 'Synthetic second day'], ['Start', '09:00:00'], ['Attendees listed', '2'],
    ['In this session', '2'], ['– by upload time', '1'], ['– by manual assignment', '1']]);
  check('the third session: no attendance yet', d.sessions[2].rows, [['Date', 'September 24, 2026'], ['Start', '15:00'], ['End', 'not planned'],
    ['Chair', 'not entered'], ['Minute taker(s)', 'not entered'], ['Administrative note', 'entered'],
    ['Attendance', 'not imported'], ['In this session', '3'], ['– by upload time', '1'], ['– by manual assignment', '2']]);
  check('the dialog shows what the status block says: one model', d.overview.map((row) => '   ' + row[0] + ': ' + row[1]), FULL_BLOCK.slice(1, 11));

  // Read-only.
  check('the dialog has a Close button and nothing else to press; it calls no server function', [(dialog.html.match(/<button/g) || []).length, (dialog.html.match(/<button[^>]*>/) || [''])[0], /google\.script\.run/.test(dialog.html), /<input|<textarea|<select/.test(dialog.html)],
    [1, '<button type="button" id="closeBtn" onclick="google.script.host.close()">', false, false]);
  check('it says where the figures come from and that nothing is changed', /Nothing is stored or changed by opening this\. The attendance values are the ones Microsoft Teams states\./.test(dialog.html), true);
  check('one script; everything is written as text, never as HTML', [d.scriptCount, /innerHTML|insertAdjacentHTML|document\.write/.test(dialog.html), /textContent/.test(dialog.html)], [1, false, true]);
  const hostile = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-09-22', '', '', '</script><b>x')]) }) });
  hostile.storeAttendance('s1', { t: '</script><img src=x onerror=alert(1)>', n: 3, s: '2026-09-22T09:00:00', e: null, d: null, a: null, p: [] });
  const hd = runDialog(hostile.statistics().html);
  check('a label and a meeting title with markup: one script element, shown as typed', [hd.scriptCount, hd.sessions[0].when, hd.sessions[0].rows[6], hostile.ui.dialogs[0].html.indexOf('</script><b>'), hostile.ui.dialogs[0].html.indexOf('<img')],
    [1, '</script><b>x, September 22, 2026', ['Meeting', '</script><img src=x onerror=alert(1)>'], -1, -1]);

  // Where it does not open.
  const main = report({ props: Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }) });
  main.s.showAdhocSessionStatistics();
  check('a main-meeting report: a message, no dialog, no download', [main.ui.dialogs.length, main.ui.alerts.map((a) => a[1]), main.downloads], [0, ['The statistics of the sessions are available for ad-hoc reports only.'], 0]);
  const noSessions = report({ props: ADHOC });
  noSessions.s.showAdhocSessionStatistics();
  check('an ad-hoc report without sessions: where to configure them, no dialog, no download', [noSessions.ui.dialogs.length, noSessions.ui.alerts.map((a) => a[1]), noSessions.downloads],
    [0, ['The statistics are per session. Configure the sessions of this report first:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…'], 0]);
  const master = report({ template: true, docId: TEMPLATE_ID });
  check('in the master template the action refuses', [/SA4 Report Template itself/.test(thrown(() => master.s.showAdhocSessionStatistics()) || ''), master.ui.dialogs.length], [true, 0]);
}

// ================================================================ 5. the TDoc list cannot be read

console.log('5. the TDoc list cannot be read');
{
  const REST = ['   Sessions: 3 (' + RANGE + ')', '   Opening details: 3 of 3 sessions', '   Attendance imported: 2 of 3 sessions', '   TDoc sessions: not available', '   Session column: in the registration table'];
  [['the download is refused', () => ({ getResponseCode: () => 500, getBlob: () => null }), 'Could not download TDOC list: HTTP 500'],
    ['the download fails', () => { throw new Error('Address unavailable'); }, 'Address unavailable']].forEach(([what, fetch, reason]) => {
    const r = fullReport({ fetch: fetch });
    const before = [r.props(), r.snapshot()];
    const alert = r.status();
    const block = r.lines();
    check(`${what}: the status is shown; sessions, opening details and attendance are there, the TDoc sessions are not available`, [alert[0], block.slice(1, 6), block.filter((l) => /⚠️/.test(l))],
      ['Report Status', REST, ['   ⚠️  The TDoc sessions are not available: the TDoc list could not be read (' + reason + ').']]);
    check(`${what}: the status of the TDoc tables is still the first part of the alert`, alert[1].indexOf(OLD_TEXT_OF_THIS_DOCUMENT(3)), 0);
    check(`${what}: nothing was written, the report is untouched, and the log has the reason only`, [r.props() === before[0], r.snapshot() === before[1], r.writes, r.logs.filter((l) => /^Session status/.test(l)).filter((l, i, list) => list.indexOf(l) === i)],
      [true, true, [], ['Session status: the TDoc sessions are not available: ' + reason]]);
    const d = runDialog(r.statistics().html);
    check(`${what}: the statistics open as well; each session shows its attendance and "not available" for its TDocs`, [d.overview.map((row) => row[0]), d.sessions[0].rows.slice(6, 13).map((row) => row[0]), d.sessions.map((x) => x.rows.slice(-1)[0])],
      [['Sessions', 'Opening details', 'Attendance imported', 'TDoc sessions', 'Session column'], ['Meeting', 'Start', 'End', 'Duration', 'Attendance records', 'Average attendance', 'Attendees listed'],
        [['TDoc sessions', 'not available'], ['TDoc sessions', 'not available'], ['TDoc sessions', 'not available']]]);
  });
  const noUrl = fullReport();
  noUrl.s.getReportConfig_ = () => ({ TDOC_LIST_URL: '' });
  check('no TDoc list is configured: said so, nothing downloaded', [noUrl.lines().filter((l) => /⚠️/.test(l)), noUrl.downloads], [['   ⚠️  The TDoc sessions are not available: the TDoc list could not be read (no TDoc list is configured).'], 0]);
  const broken = fullReport();
  broken.s.SpreadsheetApp = { open: () => { throw new Error('Service Spreadsheets failed'); } };
  check('the downloaded list cannot be opened', broken.lines().filter((l) => /⚠️/.test(l)), ['   ⚠️  The TDoc sessions are not available: the TDoc list could not be read (Service Spreadsheets failed).']);
  const noTable = fullReport();
  noTable.s.findRegistrationTable_ = () => { throw new Error('Service Documents failed'); };
  check('the registration table cannot be looked at: the column is "not known", the rest is there', [noTable.lines().filter((l) => /Session column/.test(l)), noTable.lines()[4]], [['   Session column: not known'], '   TDocs with a session: 8 of 10']);
  const everything = fullReport();
  everything.s.buildAdhocSessionStatusModel_ = () => { throw new Error('unexpected'); };
  const alert = everything.status();
  check('even a failure of the status itself does not stop Report Status Summary', [alert[0], alert[1].indexOf(OLD_TEXT_OF_THIS_DOCUMENT(3)), alert[1].endsWith(HEAD + '\n   ⚠️  The status of the sessions is not available (unexpected).'), everything.logs.filter((l) => /^Session status/.test(l))],
    ['Report Status', 0, true, ['Session status: not available: unexpected']]);
}

// ================================================================ 6. stored state that cannot be used

console.log('6. stored state that cannot be used');
{
  // Opening details.
  const o2 = fullReport();
  o2.docProps._store[OPENING_KEY] = '{"v":2,"sessions":{}}';
  check('opening details of a newer version: not available; attendance and TDocs as before', [o2.lines().slice(2, 5), o2.lines().filter((l) => /⚠️/.test(l))[0], o2.docProps._store[OPENING_KEY]],
    [['   Opening details: not available', '   Attendance imported: 2 of 3 sessions', '   TDocs with a session: 8 of 10'], '   ⚠️  The stored opening details have version 2; this release reads version 1. The status of the opening details is not available.', '{"v":2,"sessions":{}}']);
  const o1 = fullReport();
  o1.docProps._store[OPENING_KEY] = '{"v":1,';
  check('opening details that cannot be read', [o1.lines()[2], o1.lines().filter((l) => /⚠️/.test(l))[0]], ['   Opening details: not available', '   ⚠️  The stored opening details cannot be read. The status of the opening details is not available.']);

  // Attendance: one damaged session.
  const damaged = fullReport();
  const index = JSON.parse(damaged.docProps._store[ATTENDANCE_KEY]);
  damaged.docProps._store[ATTENDANCE_KEY + '_s2_' + index.entries.s2.g + '_0'] = '{"t":"tampered"}';
  check('one session whose stored attendance is damaged: named in a warning; the other session is counted and shown', [damaged.lines()[3], damaged.lines().filter((l) => /attendance/i.test(l) && /⚠️/.test(l)), damaged.lines().filter((l) => /Attendance not imported/.test(l)),
    group(damaged.model(), 's1', 'Attendance (Microsoft Teams)').length, group(damaged.model(), 's2', 'Attendance (Microsoft Teams)')],
    ['   Attendance imported: 1 of 3 sessions', ['   ⚠️  The stored attendance of Offline cannot be read. Import it again or remove it.'], ['   ℹ️  Attendance not imported: 24 Sep.'], 7, [['Attendance', 'stored, but cannot be read']]]);
  check('the damaged value is left as it is', damaged.docProps._store[ATTENDANCE_KEY + '_s2_' + index.entries.s2.g + '_0'], '{"t":"tampered"}');
  const a2 = fullReport();
  a2.docProps._store[ATTENDANCE_KEY] = '{"v":2}';
  check('attendance of a newer version: not available; opening details and TDocs as before', [a2.lines().slice(2, 5), a2.lines().filter((l) => /⚠️/.test(l))[0]],
    [['   Opening details: 3 of 3 sessions', '   Attendance imported: not available', '   TDocs with a session: 8 of 10'], '   ⚠️  The stored attendance has version 2; this release reads version 1. The status of the attendance is not available.']);
  const a1 = fullReport();
  a1.docProps._store[ATTENDANCE_KEY] = '[]';
  check('an attendance index that cannot be read', [a1.lines()[3], a1.lines().filter((l) => /⚠️/.test(l))[0]], ['   Attendance imported: not available', '   ⚠️  The stored attendance index cannot be read. The status of the attendance is not available.']);
  const stale = fullReport();
  stale.storeAttendance('s9', { t: 'Stray', n: 1, s: '2026-09-25T09:00:00', e: null, d: null, a: null, p: [] });
  stale.docProps._store[OPENING_KEY] = opening(Object.assign({}, OPENING, { s8: { chair: 'Nobody' } }));
  check('attendance and opening details of a session that is not configured: warnings; the counts are of the configured sessions', [stale.lines().slice(2, 4), stale.lines().filter((l) => /not configured/.test(l))],
    [['   Opening details: 3 of 3 sessions', '   Attendance imported: 2 of 3 sessions'], ['   ⚠️  Opening details are stored for 1 session that is not configured.', '   ⚠️  Attendance is stored for 1 session that is not configured.']]);

  // Manual TDoc assignments.
  const t2 = fullReport();
  t2.docProps._store[TDOC_KEY] = '{"v":2,"clock":"utc","add":{},"set":{}}';
  check('manual TDoc assignments of a newer version: not reinterpreted -- the sessions by upload time are counted, with a warning; the stored value is left as it is',
    [t2.lines().slice(4, 10), t2.lines().filter((l) => /⚠️/.test(l))[0], t2.docProps._store[TDOC_KEY]],
    [['   TDocs with a session: 7 of 10', '   – by upload time: 7', '   – with a manual assignment: 0 (0 added, 0 set)', '   – in more than one session: 0', '   TDocs without a session: 3', '   – set to no session: 0'],
      '   ⚠️  The stored manual TDoc session assignments have version 2; this release reads version 1. Only the sessions by upload time are counted.', '{"v":2,"clock":"utc","add":{},"set":{}}']);
  const t1 = fullReport();
  t1.docProps._store[TDOC_KEY] = '{"v":1,"clock":"moon"}';
  check('manual TDoc assignments that cannot be read', [t1.lines()[4], t1.lines().filter((l) => /⚠️/.test(l))[0]], ['   TDocs with a session: 7 of 10', '   ⚠️  The stored manual TDoc session assignments cannot be read. Only the sessions by upload time are counted.']);
  const ts = fullReport();
  ts.setTdocSessions('utc', { S4aA269001: { mode: 'set', sessions: ['s9'] } });
  check('a manual assignment that names a session that is not configured', ts.lines().filter((l) => /not configured/.test(l)), ['   ⚠️  1 manual assignment names a session that is not configured.']);

  // Everything at once.
  const all = fullReport({ fetch: () => { throw new Error('Address unavailable'); } });
  all.docProps._store[OPENING_KEY] = '{"v":2}';
  all.docProps._store[ATTENDANCE_KEY] = '{"v":2}';
  all.docProps._store[TDOC_KEY] = '{"v":2}';
  const before = all.props();
  check('everything unusable at once: the sessions are still shown, each part says that it is not available, and nothing is changed', [all.lines().slice(1, 6), all.lines().filter((l) => /⚠️/.test(l)).length, all.props() === before, runDialog(all.statistics().html).sessions.map((x) => x.rows.slice(3).map((row) => row[1]))],
    [['   Sessions: 3 (' + RANGE + ')', '   Opening details: not available', '   Attendance imported: not available', '   TDoc sessions: not available', '   Session column: in the registration table'], 3, true,
      [['not available', 'not available', 'not available'], ['not available', 'not available', 'not available'], ['not available', 'not available', 'not available']]]);
  const thrower = fullReport();
  const get = thrower.docProps.getProperty;
  thrower.docProps.getProperty = (k) => { if (k === OPENING_KEY) throw new Error('Properties service failed'); return get(k); };
  check('a property that cannot be read at all: that part is not available, the others are', [thrower.lines().slice(2, 5), thrower.lines().filter((l) => /⚠️/.test(l))[0]],
    [['   Opening details: not available', '   Attendance imported: 2 of 3 sessions', '   TDocs with a session: 8 of 10'], '   ⚠️  The opening details could not be read (Properties service failed). The status of the opening details is not available.']);
}

// ================================================================ 7. nothing is written

console.log('7. nothing is written; the same again gives the same');
{
  const r = fullReport();
  const props = r.props();
  const snapshot = r.snapshot();
  const children = r.body._children.slice();
  const first = r.status()[1];
  const firstDialog = r.statistics().html;
  const second = r.status()[1];
  const secondDialog = r.statistics().html;
  check('Report Status Summary and Post-meeting Statistics, twice each: no property written or deleted', [r.writes, r.props() === props], [[], true]);
  check('no user property touched, no trigger created', [r.userProps, r.triggers], [0, 0]);
  check('the document is untouched: the same text, and the very same elements', [r.snapshot() === snapshot, r.body._children.length === children.length && r.body._children.every((e, i) => e === children[i])], [true, true]);
  check('the manual TDoc assignments, the attendance and the opening details are what they were', [OPENING_KEY, ATTENDANCE_KEY, TDOC_KEY].map((k) => JSON.parse(props)[0][k] === r.docProps._store[k]), [true, true, true]);
  check('the same again gives the same alert and the same dialog', [first === second, firstDialog === secondDialog], [true, true]);
  check('the TDoc list is downloaded once per view, and nothing is logged about the sessions when all can be read', [r.downloads, r.logs.filter((l) => /Session status|Opening|Attendance/.test(l))], [4, []]);
  check('no lock is taken for reading', (SECTION.match(/LockService|withAdhocAttendanceLock_/g) || []), []);
  check('the status section writes nothing (source): no property, no document element, no trigger', (SECTION.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '')
    .match(/setProperty|setProperties|deleteProperty|deleteAllProperties|writeAdhocAttendanceValue_|saveAdhoc\w+|appendParagraph|appendTable|insertParagraph|insertTable|removeChild|setText|setHeading|appendTableCell|removeAdhocBodyChild_|render\w+_\(|refresh\w+_\(|newTrigger|getUserProperties/g) || []), []);
  check('it stores nothing of its own: no property key is defined for it (source)', [/ADHOC_[A-Z_]*(STATUS|STATISTICS)[A-Z_]*_?\s*=/.test(CODE), Object.keys(r.docProps._store).filter((k) => /STAT/i.test(k))], [false, []]);
  check('the model is derived each time: a change of the sessions shows at once', (() => {
    r.docProps._store.ADHOC_SESSIONS = sessionsProperty([S1, S2]);
    return [r.lines()[1], r.lines()[3]];
  })(), ['   Sessions: 2 (September 22–23, 2026)', '   Attendance imported: 2 of 2 sessions']);
}

// ================================================================ 8. privacy

console.log('8. privacy');
{
  const r = fullReport();
  const alert = r.status()[1];
  const block = r.lines().join('\n');
  const html = r.statistics().html;
  const shown = runDialog(html).allText;
  const model = JSON.stringify(r.model());
  const everything = [block, html, shown, model, r.logs.join('\n')].join('\n');
  const attendance = r.s.readAdhocAttendance_(r.docProps);
  const attendees = [].concat(attendance.sessions.s1.attendees, attendance.sessions.s2.attendees);
  const identities = [];
  attendees.forEach((a) => { [a.name, a.email, a.company].forEach((v) => { if (v && identities.indexOf(v) === -1) identities.push(v); }); });
  check('the report has 25 stored attendees with names, addresses and companies to leak', [attendees.length, identities.length > 30, identities.indexOf('alex.organizer@example.com') !== -1, identities.indexOf('ExampleCorp') !== -1], [25, true, true, true]);
  check('no attendee name, e-mail address or company is in the block, the dialog, the model or the log', identities.filter((v) => everything.indexOf(v) !== -1), []);
  check('no e-mail address at all, and nothing of the parts of the export that are never kept', [/@/.test(everything.replace(/@media/g, '')), FX.PRIVATE_STRINGS.filter((v) => everything.indexOf(v) !== -1)], [false, []]);
  check('no join or leave time and no duration of an attendee', FX.PARTICIPANTS.map((p) => p.duration).filter((v) => v && everything.indexOf(v) !== -1), []);
  check('the Chair, the minute takers and the notes are not repeated', ['Alex Organizer', 'Sam Rivera', 'Casey Guest', 'A sensitive administrative note.', 'Only a note here.', 'sensitive'].filter((v) => everything.indexOf(v) !== -1), []);
  check('no upload time: neither as the list has it nor converted', [TDOCS.map((t) => t[3]).filter((v) => v && everything.indexOf(v) !== -1), /\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}/.test(everything), /07:26:14|09:26:14|16:00:01|18:00:01/.test(everything)], [[], false, false]);
  check('no TDoc number, title, source or contact', [/S4aA26\d{4}/.test(block + shown + model), /Synthetic title of/.test(everything)], [false, false]);
  check('the alert adds nothing about people to what the status of the TDoc tables says', identities.concat(['Alex Organizer', 'Sam Rivera']).filter((v) => alert.indexOf(v) !== -1), []);
  check('what the dialog is given: the overview, the notes and the sessions -- labels and values, nothing else', (() => {
    const json = JSON.parse((html.match(/var MODEL = (.*);\n/) || [])[1]);
    return [Object.keys(json), Object.keys(json.sessions[0]), Object.keys(json.sessions[0].groups[0])];
  })(), [['overview', 'notes', 'sessions'], ['id', 'when', 'groups'], ['title', 'rows']]);
  check('the status section logs fixed texts and error messages only (source)', (SECTION.match(/Logger\.log\((.*)\);/g) || []).filter((line) => !/^Logger\.log\('Session status: [^']*' \+ e\.message\);$/.test(line)), []);
  check('it never reads a name, an address, a company, a field of the opening details or an upload time (source)', (SECTION.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '')
    .match(/\.name\b|\.email\b|\.company\b|\.companies\b|entry\.chair \+|entry\.minuteTakers \+|entry\.note \+|\.uploaded|\.dateOnly|\.at\b|meetingTitle|titleCol|sourceCol|contactCol/g) || []), []);
  check('a failure is logged with its reason and nothing else', (() => {
    const f = fullReport({ fetch: () => { throw new Error('Address unavailable'); } });
    f.status();
    return [f.logs.filter((l) => /^Session status/.test(l)), identities.concat(['Alex Organizer', 'sensitive']).filter((v) => f.logs.join('\n').indexOf(v) !== -1)];
  })(), [['Session status: the TDoc sessions are not available: Address unavailable'], []]);
}

// ================================================================ 9. other reports; menu; source

function openMenu(props, opts) {
  const o = opts || {};
  const loaded = loadTemplateRuntime({ release: o.release === false ? null : RELEASE, documentProperties: props });
  const ui = { menus: [] };
  ui.createMenu = (name) => {
    const m = { name: name, entries: [] };
    m.addItem = (label, fn) => { m.entries.push({ label: label, fn: fn }); return m; };
    m.addSeparator = () => { m.entries.push({ sep: true }); return m; };
    m.addSubMenu = (sub) => { m.entries.push({ sub: sub }); return m; };
    m.addToUi = () => { ui.menus.push(m); };
    return m;
  };
  loaded.sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => o.docId || REPORT_ID, getBody: () => makeFakeDocumentBody(loaded.sandbox) });
  loaded.sandbox.DocumentApp.getUi = () => ui;
  loaded.sandbox.onOpen();
  const sub = ui.menus[0].entries.filter((e) => e.sub && /Sessions/.test(e.sub.name)).map((e) => e.sub)[0];
  return { loaded: loaded, sub: sub, text: JSON.stringify(ui.menus) };
}

console.log('9. main-meeting reports and reports without sessions; menu; source');
{
  const STRAY = { [OPENING_KEY]: opening(OPENING), [TDOC_KEY]: '{"v":1,"clock":"session","add":{},"set":{"":["S4aA269001"]}}', [ATTENDANCE_KEY]: '{"v":2}' };
  [['a main-meeting report (with sessions and everything else left in its properties)', Object.assign({}, WITH_SESSIONS, STRAY, { MEETING_TYPE: 'main' }), ['MEETING_TYPE']],
    ['a main-meeting report', { MEETING_TYPE: 'main', MEETING_NUMBER: '137' }, ['MEETING_TYPE']],
    ['an ad-hoc report without sessions', Object.assign({}, ADHOC, STRAY), ['MEETING_TYPE', 'ADHOC_SESSIONS']],
    ['an ad-hoc report whose sessions were all removed', Object.assign({}, ADHOC, STRAY, { ADHOC_SESSIONS: sessionsProperty([], 4) }), ['MEETING_TYPE', 'ADHOC_SESSIONS']],
    ['an ad-hoc report whose stored sessions cannot be read', Object.assign({}, ADHOC, STRAY, { ADHOC_SESSIONS: '{"v":1,' }), ['MEETING_TYPE', 'ADHOC_SESSIONS']]].forEach(([what, props, reads]) => {
    const r = report({ props: props });
    const alert = r.status();
    check(`${what}: no block; nothing is read for it but ${reads.join(' and ')}, nothing is downloaded, nothing is logged`, [r.lines(), /SESSIONS AND ATTENDANCE|Post-meeting/.test(alert[1]), r.downloads, r.logs,
      (() => { r.reads.length = 0; r.lines(); return r.reads.slice(); })(), r.model()], [[], false, 0, [], reads, null]);
    if (OLD_CODE) check(`${what}: Report Status Summary is, character for character, the one of T-2026.10.4`, [alert[1] === oldStatus(r)[1], alert[0], alert[2], r.ui.alerts.length], [true, 'Report Status', 'OK', 1]);
    check(`${what}: nothing written, the document untouched`, [r.writes, r.snapshot() === report({ props: props }).snapshot()], [[], true]);
  });
  check('the status of the TDoc tables of this document, as every release shows it', (() => { const r = report({ props: { MEETING_TYPE: 'main' } }); return r.status()[1]; })(), ['📊 REPORT STATUS ANALYSIS\n', '═══════════════════════════════════\n', '\n📈 Total Documents: 3\n',
    '\n📊 Status Breakdown:', '   ✅ Agreed:      1 (33%)', '   ✅ Noted:       0 (0%)', '   ✅ Endorsed:    0 (0%)', '   ⚠️  Revised:     0 (0%)', '   ❌ Withdrawn:   0 (0%)', '   ⏳ Available:   1 (33%)', '   ⏳ Reserved:    0 (0%)',
    '\n✅ Decided: 1/3 (33%)', '⏳ Pending: 1/3 (33%)', '\n\n⚠️  DOCUMENTS NEEDING ATTENTION:\n', '   Available: 1 documents', '\n\n❌ Missing Minutes: 2 documents', '   - Title', '   - S4aA269002',
    '\n\n❌ Missing Disposition: 2 documents', '   - Title', '   - S4aA269002', '\n\nℹ️  No Email Discussion: 3 documents'].join('\n'));

  // Menu.
  const adhoc = openMenu(Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, WITH_SESSIONS));
  check('ad-hoc report: Post-meeting Statistics… is the last item of Sessions and Attendance', [adhoc.sub.entries.length, adhoc.sub.entries[5]], [6, { label: 'Post-meeting Statistics…', fn: 'showAdhocSessionStatistics' }]);
  check('its function exists and is public; Report Status Summary is where it was', [typeof adhoc.loaded.sandbox.showAdhocSessionStatistics, /"Report Status Summary","fn":"analyzeReportStatus"/.test(adhoc.text)], ['function', true]);
  check('main-meeting report, new report, master template, CENTRAL / Legacy: no such item', [openMenu(Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' })).text, openMenu({}).text, openMenu(WITH_SESSIONS, { docId: TEMPLATE_ID }).text, openMenu(WITH_SESSIONS, { release: false }).text].map((t) => /Statistics/.test(t)), [false, false, false, false]);
  check('ReportCreator.js: one more menu item, nothing else', (CREATOR.match(/Statistics|showAdhocSessionStatistics|SessionStatus/g) || []).length, 2);

  // Source.
  const hook = functionSource(CODE, 'analyzeReportStatus');
  check('Report Status Summary calls the block once, directly before its alert', [(hook.match(/adhocSessionStatusLines_\(\)/g) || []).length, /adhocSessionStatusLines_\(\)\.forEach\(line => lines\.push\(line\)\);\n  \n  DocumentApp\.getUi\(\)\.alert\('Report Status', lines\.join\('\\n'\), DocumentApp\.getUi\(\)\.ButtonSet\.OK\);\n\}/.test(hook)], [1, true]);
  check('outside its section (and the changelog), Code.js names the status of the sessions in that one call only', (CODE.replace(SECTION, '').replace(/ \* 2\.18\.0 \(2026-10-02\)\n[\s\S]*? \* 2\.17\.4 \(2026-10-01\)\n/, ' * 2.17.4 (2026-10-01)\n').match(/[A-Za-z_]*AdhocSessionStat[A-Za-z_]*|adhocSessionStatusLines_|summarizeAdhocTdocSessions_|adhocSessionDateRangeText_/g) || []), ['adhocSessionStatusLines_']);
  check('updates, the build and adoption do not use it', ['continuousUpdateCore_', 'collectorUpdate_', 'buildSkeletonWithTdocTables', 'adoptReportDocumentForAddon_', 'runFullReportBuildCore_'].filter((name) => /SessionStatus|SessionStatistics|summarizeAdhocTdocSessions_/.test(functionSource(CODE, name) || '')), []);
  check('the TDoc sessions come from the resolver of the Session column; the attendance values from the block of the report (source)', [/makeAdhocTdocSessionResolver_\(\)/.test(SECTION), /resolver\.effective\(td\)/.test(SECTION), /buildAdhocAttendanceBlocks_\(sessions, attendance\)/.test(SECTION),
    /formatTeamsDuration_|durationSeconds|averageAttendanceSeconds|attendanceRecords/.test(SECTION.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, ''))], [true, true, true, false]);
  check('nothing is calculated from the attendance but a count (source): no sum, no mean, no percentage', /reduce\(|Math\.(round|floor|max|min)|\/ total|\* 100|pct\(/.test(SECTION.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '')), false);
  check('Code.js is version 2.19.0 (2.18.0 added the ad-hoc sessions, 2.18.1 the TDoc upload completion, 2.19.0 the status dropdowns)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.19.0');
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll ad-hoc session status checks passed.');
process.exitCode = failures ? 1 : 0;
