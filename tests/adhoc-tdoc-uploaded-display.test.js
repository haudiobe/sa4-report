/**
 * Assign TDoc Sessions shows, per TDoc, when it was uploaded (G2 smoke-test
 * request): the value the automatic session is decided from, so that it can
 * be seen why a TDoc is in a session -- above all near a cut-off.
 *
 *   1. how an upload time is written;
 *   2. it is the clock the assignment uses: UTC converted to the report's
 *      time zone, in summer and in winter time, and around a cut-off;
 *   3. the dialog: the column, the hint, what is sent on Save;
 *   4. display only: nothing of it is stored, written into the report or
 *      shown in the status and the statistics;
 *   5. manual assignments are what they were.
 *
 * All TDocs and times are synthetic.
 *
 * Run: node tests/adhoc-tdoc-uploaded-display.test.js
 */

const fs = require('fs');
const vm = require('vm');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');

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
const S = loadCode().sandbox;
const session = (id, date, endDate, start, end, label) => { const s = { id: id, label: label || '', date: date }; if (endDate) s.endDate = endDate; s.start = start || ''; s.end = end || ''; return s; };
// "AHG Call 1": 22 September, 15:00-18:00.  "AHG Session 2": 26-28 October, until 23:00 on the last day.
const SESSIONS = [session('s1', '2026-09-22', '', '15:00', '18:00', 'AHG Call 1'), session('s2', '2026-10-26', '2026-10-28', '09:00', '23:00', 'AHG Session 2')];
const PROPS = { MEETING_TYPE: 'adhoc', MEETING_ID: '86178', MEETING_NAME: 'Synthetic ad-hoc', REPORT_SUFFIX: '6G', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/x/list.xlsx',
  ADHOC_SESSIONS: JSON.stringify({ v: 1, nextId: 3, sessions: SESSIONS }) };

function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}
const at = (utc) => ({ at: utc.replace(' ', 'T'), dateOnly: null, unreadable: false });
/** id, the upload as the TDoc list gives it (UTC), what the report's clock makes of it, the session */
const LIST = [
  ['S4aA269001', at('2026-09-18 07:26:14'), '18 Sep 09:26', 's1'],          // summer time: UTC+2
  ['S4aA269002', at('2026-09-22 13:42:07'), '22 Sep 15:42', 's1'],
  ['S4aA269003', at('2026-09-22 16:00:00'), '22 Sep 18:00:00', 's1'],       // exactly at the cut-off of the first session
  ['S4aA269004', at('2026-09-22 16:00:01'), '22 Sep 18:00:01', 's2'],       // one second after it
  ['S4aA269005', at('2026-10-24 21:30:00'), '24 Oct 23:30', 's2'],          // the last evening of summer time
  ['S4aA269006', at('2026-10-25 01:30:00'), '25 Oct 02:30', 's2'],          // winter time: UTC+1
  ['S4aA269007', at('2026-10-28 21:15:00'), '28 Oct 22:15', 's2'],
  ['S4aA269008', at('2026-10-28 22:00:00'), '28 Oct 23:00:00', 's2'],       // exactly at the final cut-off
  ['S4aA269009', at('2026-10-28 22:00:01'), '28 Oct 23:00:01', ''],         // one second after it
  ['S4aA269010', { at: null, dateOnly: null, unreadable: false }, DASH, ''],                  // reserved, not uploaded
  ['S4aA269011', { at: null, dateOnly: null, unreadable: true }, DASH, ''],                   // a value that is not a time
  ['S4aA269012', { at: null, dateOnly: '2026-09-21', unreadable: false }, '21 Sep', 's1'],    // a date without a time
  ['S4aA269013', undefined, DASH, '']                                                          // a list without upload times
];
const groupsOf = (list) => ({ '4.3': { tdocs: list.map(([id, uploaded]) => { const td = { row: [id, 'Synthetic title of ' + id, 'ExampleCorp'], tdocCol: 0, titleCol: 1, sourceCol: 2, revisedToCol: -1 }; if (uploaded !== undefined) td.uploaded = uploaded; return td; }) } });

function report(props, list) {
  const loaded = loadCode({ documentProperties: Object.assign({}, PROPS, props || {}) });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  body.appendParagraph('5.1.2 Registration of Documents').setHeading(s.DocumentApp.ParagraphHeading.HEADING3);
  body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item', 'Session']].concat((list || LIST).map((t) => [t[0], 'Synthetic title of ' + t[0], 'ExampleCorp', '4.3', 'old'])));
  body.appendParagraph('5.11 Close of the session').setHeading(s.DocumentApp.ParagraphHeading.HEADING2);
  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK' }, alert: (...a) => { ui.alerts.push(a); }, showModalDialog: (out, title) => { ui.dialogs.push({ title: title, html: out.html }); } };
  const r = { s: s, body: body, docProps: loaded.docProps, scriptProps: loaded.scriptProps, ui: ui, logs: [], writes: [] };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: (w) => { out.width = w; return out; }, setHeight: () => out }; return out; } };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  s.Utilities.formatDate = (date, zone) => formatInZone(date, zone);
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.downloadAndGroupTdocs_ = () => groupsOf(list || LIST);
  [loaded.docProps, loaded.scriptProps].forEach((store) => { const set = store.setProperty; store.setProperty = (k, v) => { r.writes.push(k); set(k, v); }; });
  r.model = () => s.buildAdhocTdocSessionsDialogModel_(s.makeAdhocTdocSessionResolver_(), s.downloadAndGroupTdocs_(), 5);
  r.stored = () => JSON.stringify([loaded.docProps._store, loaded.scriptProps._store]);
  r.snapshot = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? Array.from({ length: c.getNumRows() }, (_, i) => Array.from({ length: c.getRow(i).getNumCells() }, (_, k) => c.getRow(i).getCell(k).getText())) : c.getText())));
  return r;
}

/** Runs the dialog's script against a minimal DOM. */
function runDialog(html, server) {
  const scripts = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
  const element = (tag) => { const e = { tag: tag, children: [], style: {}, value: '', textContent: '', className: '', checked: false, disabled: false }; e.appendChild = (c) => { e.children.push(c); return c; }; return e; };
  const byId = {};
  (html.match(/id="([A-Za-z]+)"/g) || []).map((m) => m.slice(4, -1)).forEach((id) => { byId[id] = element(id); });
  const calls = []; let handlers = {}; let closed = 0;
  const runner = { withSuccessHandler: (fn) => { handlers.success = fn; return runner; }, withFailureHandler: (fn) => { handlers.failure = fn; return runner; } };
  runner.saveAdhocTdocSessions = (input) => { calls.push(JSON.parse(JSON.stringify(input))); const h = handlers; handlers = {}; h.success(JSON.parse(JSON.stringify(server.saveAdhocTdocSessions(input)))); };
  const ctx = { document: { getElementById: (id) => byId[id], createElement: element }, google: { script: { run: runner, host: { close: () => { closed++; } } } }, alert: () => {} };
  vm.createContext(ctx);
  vm.runInContext(scripts[0].replace(/^<script>|<\/script>$/g, ''), ctx);
  const rows = () => byId.rows.children;
  const row = (id) => rows().filter((tr) => tr.children[0].textContent === id)[0];
  return { ctx: ctx, el: byId, calls: calls, closed: () => closed, scriptCount: scripts.length,
    cells: (id) => row(id).children.map((td) => td.textContent), classes: (id) => row(id).children.map((td) => td.className),
    mode: (id) => row(id).children[4].children[0], boxes: (id) => row(id).children[5].children.map((label) => label.children[0]), result: (id) => row(id).children[6].textContent,
    visible: () => rows().filter((tr) => tr.style.display !== 'none').map((tr) => tr.children[0].textContent) };
}

// ================================================================ 1. the text

console.log('1. how an upload time is written');
{
  const text = S.adhocUploadedDisplayText_;
  check('day, month and time to the minute', [text(at('2026-09-22 15:42:07')), text(at('2026-09-03 09:05:00')), text(at('2026-10-28 22:15:59'))], ['22 Sep 15:42', '3 Sep 09:05', '28 Oct 22:15']);
  check('with the year, when asked for', [text(at('2026-09-22 15:42:07'), true), text(at('2027-01-02 08:00:00'), true)], ['22 Sep 2026 15:42', '2 Jan 2027 08:00']);
  check('with the seconds, when asked for', [text(at('2026-09-22 18:00:00'), false, true), text(at('2026-09-22 18:00:01'), false, true), text(at('2026-09-22 18:00:01'), true, true)], ['22 Sep 18:00:00', '22 Sep 18:00:01', '22 Sep 2026 18:00:01']);
  check('a date without a time is the date', [text({ at: null, dateOnly: '2026-09-21', unreadable: false }), text({ at: null, dateOnly: '2026-09-21', unreadable: false }, true, true)], ['21 Sep', '21 Sep 2026']);
  check('not uploaded, not readable, no upload time in the list, or anything that is not a time: the neutral sign, never a guess', [text({ at: null, dateOnly: null, unreadable: false }), text({ at: null, dateOnly: null, unreadable: true }), text(undefined), text(null),
    text({ at: 'soon', dateOnly: null, unreadable: false }), text({ at: '2026-09-22 15:42:07', dateOnly: null, unreadable: false }), text({ at: '2026-09-22T15:42:07', dateOnly: null, unreadable: true })], [DASH, DASH, DASH, DASH, DASH, DASH, DASH]);
  check('the neutral sign is the one the dialog uses for "no session"', [text(undefined), report().model().none], [DASH, DASH]);
  check('no time zone, no weekday and no "UTC" is written into the value', /UTC|CEST|CET|GMT|Mon|Tue|Wed|Thu|Fri|Sat|Sun/.test(LIST.map((t) => text(t[1], true, true)).join(' ')), false);
}

// ================================================================ 2. the clock of the assignment

console.log('2. it is the clock the assignment uses');
const r = report();
{
  const model = r.model();
  const shown = {}; model.tdocs.forEach((t) => { shown[t.id] = [t.uploaded, t.auto]; });
  check('each TDoc: its upload time in the report\'s time zone, and the session that follows from it', LIST.map((t) => [t[0]].concat(shown[t[0]])), LIST.map((t) => [t[0], t[2], t[3]]));
  check('summer time (UTC+2): 13:42 UTC is 15:42 in Berlin on 22 September', shown.S4aA269002[0], '22 Sep 15:42');
  check('winter time (UTC+1) after 25 October 2026: 21:15 UTC is 22:15 in Berlin on 28 October -- not 23:15', shown.S4aA269007[0], '28 Oct 22:15');
  check('across the change itself: 21:30 UTC on the 24th is 23:30 summer time, 01:30 UTC on the 25th is 02:30 winter time', [shown.S4aA269005[0], shown.S4aA269006[0]], ['24 Oct 23:30', '25 Oct 02:30']);
  check('at a cut-off the seconds are shown, because they decide: 18:00:00 is in AHG Call 1, 18:00:01 in AHG Session 2; 23:00:00 on the 28th is in AHG Session 2, 23:00:01 is after it',
    [shown.S4aA269003, shown.S4aA269004, shown.S4aA269008, shown.S4aA269009], [['22 Sep 18:00:00', 's1'], ['22 Sep 18:00:01', 's2'], ['28 Oct 23:00:00', 's2'], ['28 Oct 23:00:01', '']]);
  check('elsewhere they are not', model.tdocs.filter((t) => /:\d\d:\d\d$/.test(t.uploaded)).map((t) => t.id), ['S4aA269003', 'S4aA269004', 'S4aA269008', 'S4aA269009']);
  check('the value shown IS the value the assignment decides from: read back as a time, it gives the same session for every TDoc', (() => {
    const resolver = r.s.makeAdhocTdocSessionResolver_();
    const tds = r.s.flattenTdocGroups_(r.s.downloadAndGroupTdocs_());
    return tds.map((td) => {
      const used = resolver.uploaded(td);
      const direct = r.s.assignAdhocSession_(used, resolver.sessions).sessionId || '';
      const text = r.s.adhocUploadedDisplayText_(used, true, true);
      return direct === (resolver.automatic(td).sessionId || '') && (text === DASH || !used.at || text === formatBack(used.at));
    }).every(Boolean);
    function formatBack(at) { const m = at.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}:\d{2}:\d{2})$/); return parseInt(m[3], 10) + ' ' + ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][parseInt(m[2], 10) - 1] + ' ' + m[1] + ' ' + m[4]; }
  })(), true);
  check('the resolver\'s upload time is the converted one', [r.s.makeAdhocTdocSessionResolver_().uploaded({ uploaded: at('2026-10-28 21:15:00') }), r.s.makeAdhocTdocSessionResolver_().uploaded({ uploaded: at('2026-09-22 13:42:07') }).at, r.s.makeAdhocTdocSessionResolver_().uploaded({})],
    [{ at: '2026-10-28T22:15:00', dateOnly: null, unreadable: false }, '2026-09-22T15:42:07', undefined]);
  check('the model says which time zone that is', [model.clock, model.uploadedZone], ['utc', 'Europe/Berlin']);
  check('not uploaded, unreadable and "the list has no upload times": the neutral sign, and the reason is where it always was', model.tdocs.filter((t) => t.uploaded === DASH).map((t) => [t.id, t.why]),
    [['S4aA269010', 'not uploaded yet'], ['S4aA269011', 'upload time cannot be read'], ['S4aA269013', 'the TDoc list has no upload times']]);

  // The year.
  check('all in one year: no year', model.tdocs.some((t) => /20\d\d/.test(t.uploaded)), false);
  const twoYears = report({ ADHOC_SESSIONS: JSON.stringify({ v: 1, nextId: 3, sessions: [SESSIONS[0], session('s2', '2026-12-30', '2027-01-02', '', '', 'Turn of the year')] }) }, [['S4aA269001', at('2026-09-22 13:42:07')], ['S4aA269002', at('2027-01-01 10:00:00')]]);
  check('sessions or uploads of more than one year: the year is shown for all', twoYears.model().tdocs.map((t) => t.uploaded), ['22 Sep 2026 15:42', '1 Jan 2027 11:00']);
  const turn = report({ ADHOC_SESSIONS: JSON.stringify({ v: 1, nextId: 3, sessions: [SESSIONS[0], session('s2', '2026-12-30', '2027-01-02', '', '', 'Turn of the year')] }) }, [['S4aA269001', at('2026-09-22 13:42:07')], ['S4aA269002', at('2026-12-30 10:00:00')]]);
  check('a session that ends in the next year, all uploads in this one: the year is shown', turn.model().tdocs.map((t) => t.uploaded), ['22 Sep 2026 15:42', '30 Dec 2026 11:00']);
  check('a row of the dialog carries one upload value, the text to show: no second, raw one', [Object.keys(model.tdocs[0]).filter((k) => /upload|^at$|date|time/i.test(k)), /\d{4}-\d{2}-\d{2}/.test(JSON.stringify(model.tdocs))], [['uploaded'], false]);
  const earlier = report(null, [['S4aA269001', at('2025-12-30 10:00:00')], ['S4aA269002', at('2026-09-22 13:42:07')]]);
  check('an upload of an earlier year: the year is shown for all', earlier.model().tdocs.map((t) => t.uploaded), ['30 Dec 2025 11:00', '22 Sep 2026 15:42']);

  // The compatibility clock.
  const compat = report({ ADHOC_TDOC_SESSIONS: '{"v":1,"clock":"session","add":{},"set":{}}' });
  const cm = compat.model();
  check('with the compatibility clock the times are compared as the list has them, and shown so: 13:42 stays 13:42, and the sessions follow that clock', [cm.clock, cm.uploadedZone, cm.tdocs.slice(0, 4).map((t) => [t.uploaded, t.auto])],
    ['session', '', [['18 Sep 07:26', 's1'], ['22 Sep 13:42', 's1'], ['22 Sep 16:00', 's1'], ['22 Sep 16:00', 's1']]]);
  check('a manual assignment for a TDoc that is not in the list: the neutral sign', (() => { const x = report({ ADHOC_TDOC_SESSIONS: '{"v":1,"clock":"utc","add":{},"set":{"s1":["S4aA269099"]}}' }); return x.model().tdocs.filter((t) => t.id === 'S4aA269099').map((t) => [t.uploaded, t.why]); })(), [[DASH, 'not in the TDoc list']]);
}

// ================================================================ 3. the dialog

console.log('3. the dialog');
{
  const before = r.stored();
  const body = r.snapshot();
  r.s.assignAdhocTdocSessions();
  const html = r.ui.dialogs[0].html;
  const d = runDialog(html, r.s);
  check('opening the dialog changes nothing', [r.stored() === before, r.snapshot() === body, r.writes], [true, true, []]);
  check('the columns: Uploaded is between the title and the automatic session', (html.match(/<thead><tr>(.*?)<\/tr><\/thead>/) || [])[1], '<th>TDoc</th><th>Title</th><th>Uploaded</th><th>Automatic</th><th>Assignment</th><th>Sessions</th><th>Result</th>');
  check('a row: number, title, uploaded, automatic session; then the assignment, the sessions and the result', [d.cells('S4aA269002').slice(0, 4), d.result('S4aA269002'), d.classes('S4aA269002').slice(0, 4)],
    [['S4aA269002', 'Synthetic title of S4aA269002', '22 Sep 15:42', 'AHG Call 1'], 'AHG Call 1', ['', 'title', 'up', 'auto']]);
  check('every row shows its upload time, as the model has it', LIST.map((t) => d.cells(t[0])[2]), LIST.map((t) => t[2]));
  check('next to a cut-off the reason is visible: the two TDocs a second apart, and their sessions', [d.cells('S4aA269003').slice(2, 4), d.cells('S4aA269004').slice(2, 4)], [['22 Sep 18:00:00', 'AHG Call 1'], ['22 Sep 18:00:01', 'AHG Session 2']]);
  check('a TDoc without an upload time: the neutral sign, and the reason in the Automatic cell as before', [d.cells('S4aA269010').slice(2, 4), d.cells('S4aA269011').slice(2, 4)], [[DASH, DASH + ' (not uploaded yet)'], [DASH, DASH + ' (upload time cannot be read)']]);
  check('the hint says what the column is, in which time zone, and that it is not stored', d.el.uploadedHint.textContent, 'Uploaded is the upload time on the 3GPP Portal, shown in the time zone of this report (Europe/Berlin). It is shown here only and is not stored.');
  check('the cell is text and read-only: no input in it, written with textContent', [(html.match(/cell\(data\.uploaded, 'up'\);/g) || []).length, /innerHTML/.test(html), /td\.up \{ white-space: nowrap;/.test(html)], [1, false, true]);
  check('the layout: one narrow column more, the title column limited so that it wraps instead of pushing, the dialog a little wider', [/td\.title \{ color: #555; font-size: 12px; max-width: 250px; \}/.test(html), r.ui.dialogs[0].title, /\.setWidth\(940\)/.test(CODE)], [true, 'Assign TDoc Sessions', true]);
  check('the counts and the filter are what they were', (() => { const counts = d.el.counts.textContent; d.el.filter.value = '26900'; d.ctx.applyFilter(); const some = d.visible().length; d.el.filter.value = '22 Sep'; d.ctx.applyFilter(); const byTime = d.visible().length; d.el.filter.value = ''; d.ctx.applyFilter(); return [counts, some, byTime, d.visible().length]; })(),
    ['13 TDocs, 0 assigned by hand, 4 without a session', 9, 0, 13]);
  check('with the compatibility clock the hint says that the times are as the list has them', (() => { const c = report({ ADHOC_TDOC_SESSIONS: '{"v":1,"clock":"session","add":{},"set":{}}' }); c.s.assignAdhocTdocSessions(); const x = runDialog(c.ui.dialogs[0].html, c.s); return [x.el.uploadedHint.textContent, x.cells('S4aA269002')[2]]; })(),
    ['Uploaded is the upload time as the TDoc list has it. It is shown here only and is not stored.', '22 Sep 13:42']);
  check('a title with markup stays text, and so does everything else in the row', (() => { const h = report(null, [['S4aA269001', at('2026-09-22 13:42:07')]]); h.s.downloadAndGroupTdocs_ = () => ({ '4.3': { tdocs: [{ row: ['S4aA269001', '</script><b>x'], tdocCol: 0, titleCol: 1, revisedToCol: -1, uploaded: at('2026-09-22 13:42:07') }] } }); h.s.assignAdhocTdocSessions(); const x = runDialog(h.ui.dialogs[0].html, h.s); return [x.scriptCount, x.cells('S4aA269001').slice(1, 3)]; })(), [1, ['</script><b>x', '22 Sep 15:42']]);

  // Save.
  d.mode('S4aA269009').value = 'set';
  d.boxes('S4aA269009')[1].checked = true;
  d.mode('S4aA269009').onchange();
  d.mode('S4aA269002').value = 'add';
  d.boxes('S4aA269002')[1].checked = true;
  d.mode('S4aA269002').onchange();
  d.ctx.saveAssignments();
  check('Save sends the clock and the manual assignments -- no upload time', d.calls, [{ clock: 'utc', tdocs: { S4aA269002: { mode: 'add', sessions: ['s2'] }, S4aA269009: { mode: 'set', sessions: ['s2'] } } }]);
  check('what is stored: the property of the manual assignments, with TDoc numbers and session ids only', [r.writes, r.docProps.getProperty('ADHOC_TDOC_SESSIONS')], [['ADHOC_TDOC_SESSIONS'], '{"v":1,"clock":"utc","add":{"s2":["S4aA269002"]},"set":{"s2":["S4aA269009"]}}']);
}

// ================================================================ 4. display only

console.log('4. display only');
{
  const everything = r.stored();
  const TIMES = /\d\d:\d\d|2026-\d\d-\d\d[T ]|\b\d{1,2} (Sep|Oct) \d\d/;
  check('after opening the dialog and saving: no upload time is in any stored property', [TIMES.test(JSON.stringify(Object.keys(JSON.parse(everything)[0]).filter((k) => k !== 'ADHOC_SESSIONS').map((k) => JSON.parse(everything)[0][k]))), JSON.parse(everything)[1], /uploaded/i.test(everything)], [false, {}, false]);
  check('the report: the Session column got the session labels, and no time is anywhere in the document', (() => {
    const cells = JSON.parse(r.snapshot())[1].slice(1).map((row) => row[4]);
    return [cells.slice(0, 4), cells[8], TIMES.test(r.snapshot())];
  })(), [['A01', 'A01, A02', 'A01', 'A02'], 'A02', false]);
  check('the status summary and the statistics show no upload time', (() => {
    r.s.getReportConfig_ = () => ({ TDOC_LIST_URL: 'x' });
    const model = r.s.collectAdhocSessionStatus_();
    const text = JSON.stringify([model.overview, model.notes, model.sessions.map((x) => x.groups.filter((g) => g.title === 'TDocs'))]) + r.s.formatAdhocSessionStatusLines_(model).join('\n');
    return [/13:42|15:42|18:00:0|22:15|23:00:0|uploaded at|\b18 Sep\b/.test(text), model.tdocs.total];
  })(), [false, 13]);
  check('nothing was logged about a time', r.logs.filter((l) => /\d\d:\d\d/.test(l)), []);
  const uses = (name) => CODE.split('\n').filter((line) => line.indexOf(name) !== -1 && !/^function |^\s*(\/|\*)/.test(line)).map((l) => l.trim().slice(0, 60));
  check('the upload time of the resolver is read in one place: the dialog model (source)', [uses('resolver.uploaded(').length, uses('adhocUploadedDisplayText_(').length, (() => { const a = CODE.indexOf('\nfunction buildAdhocTdocSessionsDialogModel_('); const src = CODE.slice(a, CODE.indexOf('\n}\n', a)); return /resolver\.uploaded\(td\)/.test(src) && /adhocUploadedDisplayText_\(/.test(src); })()], [1, 1, true]);
  check('the functions that write the report, store the assignments or build the status do not use it (source)', ['createSummaryTable_', 'updateRegisteredDocumentsTable_', 'refreshRegistrationSessionColumn_', 'enableAdhocSessionColumn', 'saveAdhocTdocSessionsWith_', 'serializeAdhocTdocSessions_', 'summarizeAdhocTdocSessions_', 'buildAdhocSessionStatusModel_', 'collectAdhocSessionStatus_', 'showAdhocSessionStatistics', 'continuousUpdateCore_']
    .filter((name) => { const a = CODE.indexOf('\nfunction ' + name + '('); return /\.uploaded\b|adhocUploadedDisplayText_|uploadedZone/.test(CODE.slice(a, CODE.indexOf('\n}\n', a))); }), []);
  check('the cell of the Session column is the session labels, never a time', (() => { const resolver = r.s.makeAdhocTdocSessionResolver_(); return r.s.flattenTdocGroups_(r.s.downloadAndGroupTdocs_()).map((td) => resolver.cell(td)).filter((t) => /\d\d:\d\d/.test(t)); })(), []);
}

// ================================================================ 5. manual assignments

console.log('5. manual assignments and the automatic rule are what they were');
{
  const m = report({ ADHOC_TDOC_SESSIONS: '{"v":1,"clock":"utc","add":{"s2":["S4aA269002"]},"set":{"":["S4aA269007"],"s1":["S4aA269004"]}}' });
  const model = m.model();
  const of = (id) => model.tdocs.filter((t) => t.id === id).map((t) => [t.uploaded, t.auto, t.mode, t.sessions])[0];
  check('automatic and also, only these, and "no session": the upload time and the automatic session are shown as they are, the manual choice beside them', [of('S4aA269002'), of('S4aA269004'), of('S4aA269007'), of('S4aA269001')],
    [['22 Sep 15:42', 's1', 'add', ['s2']], ['22 Sep 18:00:01', 's2', 'set', ['s1']], ['28 Oct 22:15', 's2', 'set', []], ['18 Sep 09:26', 's1', 'auto', []]]);
  check('the effective sessions are unchanged by the display', (() => { const resolver = m.s.makeAdhocTdocSessionResolver_(); const by = {}; m.s.flattenTdocGroups_(m.s.downloadAndGroupTdocs_()).forEach((td) => { by[td.row[0]] = resolver.effective(td); }); return [by.S4aA269002, by.S4aA269004, by.S4aA269007, by.S4aA269003, by.S4aA269009]; })(),
    [['s1', 's2'], ['s1'], [], ['s1'], []]);
  check('the cut-off is inclusive, and the automatic rule is the one function it always was (source)', [(CODE.match(/if \(at <= adhocSessionCutoff_\(ordered\[i\]\)\) return ordered\[i\]\.id;/g) || []).length, (CODE.match(/const automatic = function \(td\) \{ return assignAdhocSession_\(uploaded\(td\), sessions\); \};/g) || []).length], [1, 1]);
  check('the clock selector is still under Advanced, with UTC as the normal choice', (() => { m.s.assignAdhocTdocSessions(); const html = m.ui.dialogs[0].html; return [/<details[^>]*>\s*<summary[^>]*>Advanced<\/summary>[\s\S]*<select id="clock">/.test(html), /<option value="utc">in UTC, as the 3GPP Portal records them \(normal\)<\/option>/.test(html)]; })(), [true, true]);
  check('Code.js is version 2.19.0 (2.18.0 added the ad-hoc sessions, 2.18.1 the TDoc upload completion, 2.19.0 the status dropdowns)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.19.0');
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll TDoc uploaded-display checks passed.');
process.exitCode = failures ? 1 : 0;
