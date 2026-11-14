/**
 * GitHub issue #11 -- session times in a main-meeting report, with the
 * existing session implementation (Configure Sessions…, ADHOC_SESSIONS, the
 * Session administration section).
 *
 *   1. Saving times leaves the meeting type, the report type, the TDoc list
 *      and the agenda filter of a main Audio report exactly as they were.
 *   2. The times are written in the meeting's local time, across DST and
 *      Alberta's change to ABT.
 *   3. Save and reopen: what was saved is what the dialog shows again.
 *   4. The text of the report outside the Session administration section is
 *      unchanged; nothing is written that the ad-hoc feature adds.
 *   5. No TDoc is read, inserted or updated: no download, no build, no update.
 *   6. The ad-hoc report keeps its behaviour.
 *
 * All data is synthetic. Run: node tests/session-times-main-report.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { codeWithoutSessionTimes, creatorWithoutSessionTimes, SESSION_TIMES_NEW_FUNCTIONS, SESSION_TIMES_CHANGED_FUNCTIONS,
  SESSION_TIMES_CREATOR_NEW_FUNCTIONS, SESSION_TIMES_CREATOR_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');

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

const KEY = 'ADHOC_SESSIONS';
const HEADING = 'Session administration';
// A main Audio report of SA4#138 (Calgary), as Configure Meeting stores it.
const MAIN_AUDIO = {
  SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_ID: '60780', MEETING_NAME: 'SA4#138', MEETING_DATE: 'November 16, 2026',
  REPORT_SUFFIX: 'Audio', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138/Docs/TDoc_List_Meeting_SA4#138.xlsx',
  AGENDA_TDOC: 'S4-261800', FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138/Docs/',
  MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: '60780|November 16, 2026'
};
// In the stored key order: endDate (only for a session of several days) after date.
const row = (id, date, start, end, label, endDate) => { const r = { id: id, label: label || '', date: date }; if (endDate) r.endDate = endDate; r.start = start || ''; r.end = end || ''; return r; };

/** A main report in a sandbox: opening, a manual "Online information" block, Audio items with TDoc tables, the close. */
function report(props) {
  const loaded = loadCode({ documentProperties: Object.assign({}, props) });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const H = s.DocumentApp.ParagraphHeading;
  body.appendParagraph('1 Opening of the session').setHeading(H.HEADING1);
  body.appendParagraph('The chair opens the meeting on November 16, 2026 at 09:00 ABT.');
  body.appendParagraph('Online information').setHeading(H.HEADING3);
  body.appendParagraph('Meeting name: typed by hand, not generated');
  body.appendParagraph('7 Audio').setHeading(H.HEADING1);
  body.appendParagraph('7.1 IVAS').setHeading(H.HEADING2);
  body.appendTable([['TDoc', 'S4-261801'], ['Title', 'Synthetic title'], ['Minutes', 'Minutes typed by hand.']]);
  body.appendParagraph('7.2 EVS').setHeading(H.HEADING2);
  body.appendTable([['TDoc', 'S4-261802'], ['Title', 'Another synthetic title'], ['Minutes', '']]);
  body.appendParagraph('20 Close of the session').setHeading(H.HEADING1);
  body.appendParagraph('The chair closes the session.');

  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK' } };
  ui.alert = (...args) => { ui.alerts.push(args); return 'OK'; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title: title, html: out.html }); };
  const r = { s: s, body: body, ui: ui, docProps: loaded.docProps, writes: [], calls: [] };
  s.Logger = { log: () => {} };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  const set = loaded.docProps.setProperty;
  const del = loaded.docProps.deleteProperty;
  loaded.docProps.setProperty = (k, v) => { r.writes.push('set ' + k); set(k, v); };
  loaded.docProps.deleteProperty = (k) => { r.writes.push('delete ' + k); del(k); };
  // Every way to read, insert or update TDocs, and every build or update: counted, never run.
  ['downloadAndGroupTdocs_', 'continuousUpdateCore_', 'continuousUpdate', 'runFullReportBuild', 'buildSkeletonWithTdocTables', 'insertNewTdoc_',
    'updateRegisteredDocumentsTable_', 'createSummaryTable_', 'addAbstractsForTables_', 'refreshAdhocAttendanceSection_', 'refreshAdhocMeetingInformationSafely_'].forEach((name) => {
    if (typeof s[name] !== 'function') throw new Error('no function ' + name);
    s[name] = () => { r.calls.push(name); throw new Error(name + ' must not run'); };
  });
  s.UrlFetchApp = { fetch: (url) => { r.calls.push('fetch ' + url); throw new Error('no fetch'); } };
  r.texts = () => body._children.map((c) => (c.getType() === 'TABLE' ? '[table ' + c.getRow(0).getCell(1).getText() + ']' : c.getText()));
  r.container = () => { const c = s.findAdhocOpeningContainer_(body); return c ? body._children.slice(c.start, c.end).map((e) => e.getText()) : null; };
  r.withoutContainer = () => { const c = s.findAdhocOpeningContainer_(body); const t = r.texts(); return c ? t.slice(0, c.start).concat(t.slice(c.end)) : t; };
  r.save = (rows) => s.saveAdhocSessionsConfiguration(JSON.parse(JSON.stringify(rows)));
  r.props = () => Object.assign({}, loaded.docProps._store);
  return r;
}

// ================================================================ 1
console.log('1. report type, meeting type, TDoc list and agenda filter are unchanged');
{
  const r = report(MAIN_AUDIO);
  const before = r.props();
  const contextBefore = JSON.stringify(r.s.getMeetingContext_());
  const selectorBefore = r.s.getMeetingContext_().report.agendaSelector;
  const result = r.save([row('', '2026-11-16', '09:00', '17:30', 'Day 1'), row('', '2026-11-17', '09:00', '18:00')]);
  check('the save is accepted', [result.ok, result.changed, result.errors], [true, true, []]);
  check('only ADHOC_SESSIONS is written', r.writes, ['set ' + KEY]);
  const after = r.props();
  delete after[KEY];
  check('every other property is as before, byte for byte', after, before);
  check('meeting type, report type, TDoc list URL, agenda: as before', ['MEETING_TYPE', 'REPORT_SUFFIX', 'TDOC_LIST_URL', 'AGENDA_TDOC'].map((k) => r.docProps.getProperty(k)),
    ['main', 'Audio', MAIN_AUDIO.TDOC_LIST_URL, 'S4-261800']);
  check('the meeting context (report type, agenda selector, TDoc list) is the same', JSON.stringify(r.s.getMeetingContext_()), contextBefore);
  check('the Audio agenda filter: 7.x only, as before', [selectorBefore, ['7.1', '7.2.3', '8.1', '17.1', '11.1'].map((a) => r.s.agendaSelectorMatches_(r.s.getMeetingContext_().report.agendaSelector, a))],
    [{ mode: 'prefix', value: '7.' }, [true, true, false, false, false]]);
  check('no session reaches the TDoc selection, the Session column or the upload-time grouping', [r.s.getAdhocSessions_(), r.s.adhocSessionsEnabled_(), r.s.makeAdhocTdocSessionResolverSafely_ ? r.s.makeAdhocTdocSessionResolverSafely_() : null],
    [[], false, null]);
  check('the configured sessions are readable for the Session administration section', r.s.getConfiguredSessions_().map((x) => [x.id, x.date, x.start, x.end, x.displayLabel]),
    [['s1', '2026-11-16', '09:00', '17:30', 'Day 1'], ['s2', '2026-11-17', '09:00', '18:00', '17 Nov']]);
  check('no meeting-type change: a report without a meeting type is refused and nothing is written', (() => {
    const x = report(Object.assign({}, MAIN_AUDIO, { MEETING_TYPE: '' }));
    const res = x.save([row('', '2026-11-16', '09:00', '17:30')]);
    return [res.ok, res.errors, x.writes, x.texts().indexOf(HEADING)];
  })(), [false, ['Sessions are available for ad-hoc and main-meeting reports only.'], [], -1]);
}

// ================================================================ 2
console.log('2. local time of the meeting, across time zones and DST');
{
  const shown = (props, rows) => { const r = report(Object.assign({}, MAIN_AUDIO, props)); r.save(rows); return r.container(); };
  check('Calgary on 16 November 2026: ABT (Alberta stays at UTC-6), times with their zone', shown({}, [row('', '2026-11-16', '09:00', '17:30', 'Plenary'), row('', '2026-11-17', '', '12:00')]),
    [HEADING, 'A01: Plenary, November 16, 2026, 09:00–17:30 ABT', 'A02: November 17, 2026, until 12:00 ABT']);
  const BERLIN = { MEETING_ID: '60777', MEETING_DATE: 'October 23, 2026', MEETING_TIME_ZONE: 'Europe/Berlin', MEETING_START_BASIS: '60777|October 23, 2026' };
  check('Berlin over the change to winter time: each time names the zone of its own day', shown(BERLIN, [row('', '2026-10-23', '09:00', '17:00', 'AHG', '2026-10-26'), row('', '2026-10-27', '10:00', '12:00')]),
    [HEADING, 'A01: AHG, October 23–26, 2026 (from 09:00 CEST on the first day, until 17:00 CET on the last day)', 'A02: October 27, 2026, 10:00–12:00 CET']);
  check('a session through the night of the change: both zones', shown(BERLIN, [row('', '2026-10-25', '00:30', '04:00')]), [HEADING, 'A01: October 25, 2026, 00:30 CEST–04:00 CET']);
  const DENVER = { MEETING_TIME_ZONE: 'America/Denver' };
  check('Denver: MDT before 1 November 2026, MST after', shown(DENVER, [row('', '2026-10-30', '09:00', '17:00'), row('', '2026-11-02', '09:00', '17:00')]),
    [HEADING, 'A01: October 30, 2026, 09:00–17:00 MDT', 'A02: November 2, 2026, 09:00–17:00 MST']);
  check('a zone stored for another meeting is not used: the times stand without a zone', shown({ MEETING_START_BASIS: '60779|November 16, 2026' }, [row('', '2026-11-16', '09:00', '17:30')]),
    [HEADING, 'A01: November 16, 2026, 09:00–17:30']);
}

// ================================================================ 3
console.log('3. save and reopen');
{
  const r = report(MAIN_AUDIO);
  r.s.configureAdhocSessions();
  const modelOf = (html) => JSON.parse((html.match(/var MODEL = (.*);\n/) || [])[1]);
  check('the dialog opens in a main report; it says that the TDocs are not changed', [r.ui.alerts.length, r.ui.dialogs.length, /the TDocs of this report are not changed/.test(r.ui.dialogs[0].html)], [0, 1, true]);
  check('first use: one session proposed from the meeting date, nothing stored by opening', [modelOf(r.ui.dialogs[0].html).sessions, r.writes], [[row('', '2026-11-16')], []]);
  check('the times are meant in the meeting zone', /meant in the meeting time zone, America\/Edmonton \(ABT\)/.test(r.ui.dialogs[0].html), true);
  const rows = [row('', '2026-11-16', '09:00', '17:30', 'Day 1'), row('', '2026-11-18', '14:00', '10:00', 'Days 3-5', '2026-11-20')];
  check('saved', r.save(rows).ok, true);
  r.s.configureAdhocSessions();
  const saved = [row('s1', '2026-11-16', '09:00', '17:30', 'Day 1'), row('s2', '2026-11-18', '14:00', '10:00', 'Days 3-5', '2026-11-20')];
  check('reopened: the dialog shows what was saved, with the ids', modelOf(r.ui.dialogs[1].html).sessions, saved);
  const reopened = report(r.props());
  check('a new run (the document opened again) reads the same sessions', reopened.s.getConfiguredSessions_().map((x) => { const y = Object.assign({}, x); delete y.displayLabel; return y; }), saved);
  const writesBefore = r.writes.length;
  const again = r.save(modelOf(r.ui.dialogs[1].html).sessions);
  check('saving the reopened dialog unchanged writes nothing and leaves the section as it is', [again.ok, again.changed, r.writes.length === writesBefore, r.container()],
    [true, false, true, [HEADING, 'A01: Day 1, November 16, 2026, 09:00–17:30 ABT', 'A02: Days 3-5, November 18–20, 2026 (from 14:00 ABT on the first day, until 10:00 ABT on the last day)']]);
  const edited = r.save([saved[0], row('s2', '2026-11-18', '15:00', '10:00', 'Days 3-5', '2026-11-20')]);
  check('an edited time is saved and shown in place', [edited.ok, edited.changed, r.container()[2]], [true, true, 'A02: Days 3-5, November 18–20, 2026 (from 15:00 ABT on the first day, until 10:00 ABT on the last day)']);
  check('removing every session removes the section and keeps the id counter', [r.save([]).ok, r.container(), JSON.parse(r.docProps.getProperty(KEY))], [true, null, { v: 1, nextId: 3, sessions: [] }]);
}

// ================================================================ 4
console.log('4. the text of the report is unchanged outside the section');
{
  const r = report(MAIN_AUDIO);
  const before = r.texts();
  const tablesBefore = JSON.stringify(r.body._children.filter((c) => c.getType() === 'TABLE').map((t) => Array.from({ length: t.getNumRows() }, (_, i) => [t.getRow(i).getCell(0).getText(), t.getRow(i).getCell(1).getText()])));
  r.save([row('', '2026-11-16', '09:00', '17:30')]);
  check('the section is at the end of the opening, before the next heading', r.texts().slice(0, 5), ['1 Opening of the session', 'The chair opens the meeting on November 16, 2026 at 09:00 ABT.', HEADING, 'A01: November 16, 2026, 09:00–17:30 ABT', 'Online information']);
  check('everything else, in order, as before', r.withoutContainer(), before);
  check('the TDoc tables, with the minutes typed by hand: as before', JSON.stringify(r.body._children.filter((c) => c.getType() === 'TABLE').map((t) => Array.from({ length: t.getNumRows() }, (_, i) => [t.getRow(i).getCell(0).getText(), t.getRow(i).getCell(1).getText()]))), tablesBefore);
  check('the "Online information" typed by hand is not replaced (it is generated only in an ad-hoc report)', r.texts().slice(4, 6), ['Online information', 'Meeting name: typed by hand, not generated']);
  r.save([row('s1', '2026-11-16', '10:00', '17:30')]);
  check('a second save replaces the section only', [r.container(), r.withoutContainer()], [[HEADING, 'A01: November 16, 2026, 10:00–17:30 ABT'], before]);
  const noOpening = report(MAIN_AUDIO);
  noOpening.body._children.splice(0, 2);
  const textNoOpening = noOpening.texts();
  const res = noOpening.save([row('', '2026-11-16', '09:00', '17:30')]);
  check('a report without an opening section: saved, said so, the document untouched', [res.ok, /no "Opening of the session" section/.test(res.notice), noOpening.texts()], [true, true, textNoOpening]);
}

// ================================================================ 5
console.log('5. no TDoc is read, inserted or updated');
{
  const r = report(MAIN_AUDIO);
  r.s.configureAdhocSessions();
  r.save([row('', '2026-11-16', '09:00', '17:30')]);
  r.save([row('s1', '2026-11-16', '09:30', '17:30'), row('', '2026-11-17', '09:00', '12:00')]);
  r.save([]);
  check('no download, no build, no update, no TDoc table, no attendance, no Online information', r.calls, []);
  check('no alert other than the dialog', r.ui.alerts, []);
  const tables = r.body._children.filter((c) => c.getType() === 'TABLE').length;
  check('the number of TDoc tables is unchanged', tables, 2);
}

// ================================================================ 6
console.log('6. an ad-hoc report is unchanged');
{
  const adhoc = report(Object.assign({}, MAIN_AUDIO, { MEETING_TYPE: 'adhoc' }));
  const before = adhoc.texts();
  const res = adhoc.save([row('', '2026-11-16', '09:00', '17:30')]);
  check('saved; without opening details or a section, the document is not touched (as before)', [res.ok, adhoc.texts(), adhoc.calls], [true, before, []]);
  check('its sessions reach the ad-hoc readers', adhoc.s.adhocSessionsEnabled_(), true);
  adhoc.s.configureAdhocSessions();
  check('its dialog text is that of an ad-hoc series', /The meetings of this ad-hoc series\./.test(adhoc.ui.dialogs[0].html), true);
}

// ================================================================ 7
console.log('7. the change itself');
{
  const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutSessionTimes(CODE);
  const creatorView = creatorWithoutSessionTimes(CREATOR);
  check('every change is in Code.js and ReportCreator.js exactly once', view.problems.concat(creatorView.problems), []);
  const names = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((x) => x.replace(/^function\s+|\s*\($/g, ''));
  const source = (src, n) => { const at = src.indexOf('\nfunction ' + n + '('); return at === -1 ? null : src.slice(at + 1, src.indexOf('\n}\n', at) + 2); };
  check('Code.js: the new functions', names(CODE).filter((n) => names(view.code).indexOf(n) === -1).sort(), SESSION_TIMES_NEW_FUNCTIONS.slice().sort());
  check('Code.js: the changed functions', names(view.code).filter((n) => source(CODE, n) !== source(view.code, n)).sort(), SESSION_TIMES_CHANGED_FUNCTIONS.slice().sort());
  check('ReportCreator.js: the new and the changed functions', [names(CREATOR).filter((n) => names(creatorView.code).indexOf(n) === -1), names(creatorView.code).filter((n) => source(CREATOR, n) !== source(creatorView.code, n))],
    [SESSION_TIMES_CREATOR_NEW_FUNCTIONS, SESSION_TIMES_CREATOR_CHANGED_FUNCTIONS]);
  check('untouched: the TDoc download and grouping, the update, the build (only its last step, finishAdhocOpeningRebuild_(), changed), the Session column, the agenda filter', ['downloadAndGroupTdocs_', 'continuousUpdateCore_', 'buildSkeletonWithTdocTables', 'insertNewTdoc_',
    'makeAdhocTdocSessionResolver_', 'adhocSessionsEnabled_', 'isAdhocMeetingForSessions_', 'agendaSelectorMatches_', 'getMeetingContext_', 'refreshAdhocOpeningSection_'].filter((n) => source(CODE, n) !== source(view.code, n)), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nall checks passed');
process.exit(failures ? 1 : 0);
