/**
 * Ad-hoc sessions that span several calendar days.
 *
 * A session is one logical meeting: one id, one label, one set of opening
 * details, one place for its TDocs -- and it may last several days. Teams
 * exports one attendance report per meeting, so such a session has several
 * attendance records, one per day.
 *
 *   1. The session model: endDate, validation, order, labels, date ranges.
 *   2. Configure Sessions.
 *   3. TDoc assignment: the cut-off is on the last day.
 *   4. Teams attendance: which days are accepted, records per day.
 *   5. Rendering, formatting, refresh, rebuild.
 *   6. Company corrections across the days.
 *   7. Opening details, status and statistics.
 *   8. A session of one day is exactly what it was.
 *
 * The example of the tests: "AHG Call 1" on 22 September 2026 and
 * "AHG Session 2" from 26 to 28 October 2026. Everything else is synthetic.
 *
 * Run: node tests/adhoc-multiday-sessions.test.js
 */

const fs = require('fs');
const vm = require('vm');
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
const KEY = 'ADHOC_SESSIONS';
const ATT = 'ADHOC_SESSION_ATTENDANCE';
const S = loadCode().sandbox;

const session = (id, date, endDate, start, end, label) => { const s = { id: id, label: label || '', date: date }; if (endDate) s.endDate = endDate; s.start = start || ''; s.end = end || ''; return s; };
const CALL1 = session('s1', '2026-09-22', '', '15:00', '18:00', 'AHG Call 1');
const SESSION2 = session('s2', '2026-10-26', '2026-10-28', '09:00', '23:00', 'AHG Session 2');
const sessionsProperty = (list, nextId) => JSON.stringify({ v: 1, nextId: nextId || 10, sessions: list });
const ADHOC = { MEETING_TYPE: 'adhoc', MEETING_ID: '86178', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'September 22, 2026', REPORT_SUFFIX: '6G' };
const TWO = Object.assign({}, ADHOC, { [KEY]: sessionsProperty([CALL1, SESSION2]) });

function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}
/** The synthetic Teams export for another day, optionally with other participants. */
const exportFor = (date, participants) => {
  const [, m, d] = date.match(/^\d{4}-(\d{2})-(\d{2})$/);
  const options = participants ? { participants: participants, summary: { attendedParticipants: String(participants.length) } } : undefined;
  return FX.buildExport(options).split(FX.DAY).join(parseInt(m, 10) + '/' + parseInt(d, 10) + '/26');
};
const person = (name, email) => ({ name: name, join: '3:00:00 PM', leave: '4:00:00 PM', duration: '1h', email: email || '', upn: email ? 'upn-' + email.replace('@', '@tenant.') : '', role: 'Presenter' });
const DAY1 = [person('Alex Organizer', 'alex.organizer@example.com'), person('Sam Rivera'), person('Lee Wong', 'lee.wong@example.com')];
const DAY2 = [person('Alex Organizer', 'alex.organizer@example.com'), person('Sam Rivera'), person('Noor Late', 'noor.late@example.org'), person('Dana Twin'), person('Dana Twin')];
const DAY3 = [person('Alex Organizer', 'alex.organizer@example.com'), person('Omar Open', 'omar.open@example.com')];
const base64 = (text) => FX.utf16le(text).toString('base64');
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

/** One report in a sandbox, with an opening and a closing section. */
function report(props) {
  const loaded = loadCode({ documentProperties: props || TWO });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const H = s.DocumentApp.ParagraphHeading;
  body.appendParagraph('5.1 Opening of the session').setHeading(H.HEADING2);
  body.appendParagraph('The chair opens the session.');
  body.appendParagraph('5.1.2 Registration of Documents').setHeading(H.HEADING3);
  body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4aA269001', 'Synthetic title', 'ExampleCorp', '5.4']]);
  body.appendParagraph('5.11 Close of the session').setHeading(H.HEADING2);
  body.appendParagraph('The chair closes the session.');
  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK' }, alert: (...a) => { ui.alerts.push(a); }, showModalDialog: (out, title) => { ui.dialogs.push({ title: title, html: out.html }); } };
  const r = { s: s, body: body, docProps: loaded.docProps, ui: ui, logs: [], writes: [] };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  s.Utilities.formatDate = (date, zone) => formatInZone(date, zone);
  s.Utilities.base64Decode = (b) => Array.from(Buffer.from(b, 'base64')).map((x) => (x > 127 ? x - 256 : x));
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  const set = loaded.docProps.setProperty; const del = loaded.docProps.deleteProperty;
  loaded.docProps.setProperty = (k, v) => { r.writes.push('set ' + k); set(k, v); };
  loaded.docProps.deleteProperty = (k) => { r.writes.push('delete ' + k); del(k); };
  r.preview = (id, date, participants) => s.previewTeamsAttendanceImport(id, base64(exportFor(date, participants)));
  r.import = (id, date, participants) => { const b = base64(exportFor(date, participants)); const p = s.previewTeamsAttendanceImport(id, b); return p.ok ? s.confirmTeamsAttendanceImport(id, b, p.token) : p; };
  r.state = () => s.readAdhocAttendance_(loaded.docProps);
  r.attendanceKeys = () => Object.keys(loaded.docProps._store).filter((k) => k.indexOf(ATT) === 0).sort();
  r.records = (id) => s.adhocAttendanceRecordsOf_(r.state(), id).map((x) => [x.key, x.day, x.model ? x.model.attendees.length : null]);
  r.container = () => { const c = s.findAdhocAttendanceContainer_(body); return c ? body._children.slice(c.start, c.end) : []; };
  r.lines = () => r.container().map((e) => (e.getType() === 'TABLE' ? '[table ' + (e.getNumRows() - 1) + ']' : e.getText())).filter((t) => t !== '');
  r.tables = () => r.container().filter((e) => e.getType() === 'TABLE');
  r.snapshot = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
  return r;
}

// ================================================================ 1. the session model

console.log('1. the session model');
{
  const v = S.validateAdhocSessions_;
  const ok = (list) => { const x = v(list); return x.errors.length ? x.errors : x.sessions; };

  check('a session of one day without endDate: stored as it always was', [ok([{ id: 's1', label: '', date: '2026-09-22', start: '15:00', end: '18:00' }]), S.serializeAdhocSessions_({ nextId: 2, sessions: [{ id: 's1', label: '', date: '2026-09-22', start: '15:00', end: '18:00' }] })],
    [[{ id: 's1', label: '', date: '2026-09-22', start: '15:00', end: '18:00' }], '{"v":1,"nextId":2,"sessions":[{"id":"s1","label":"","date":"2026-09-22","start":"15:00","end":"18:00"}]}']);
  check('an endDate equal to the date, or empty: the same session of one day, no endDate stored', [ok([{ id: 's1', label: '', date: '2026-09-22', endDate: '2026-09-22', start: '', end: '' }]), ok([{ id: 's1', label: '', date: '2026-09-22', endDate: '', start: '', end: '' }]), ok([{ id: 's1', label: '', date: '2026-09-22', endDate: '  ', start: '', end: '' }])].map((x) => JSON.stringify(x)).filter((x, i, l) => l.indexOf(x) === i),
    ['[{"id":"s1","label":"","date":"2026-09-22","start":"","end":""}]']);
  check('26 to 28 October: endDate is kept, after date', [ok([SESSION2]), S.serializeAdhocSessions_({ nextId: 3, sessions: [SESSION2] })],
    [[{ id: 's2', label: 'AHG Session 2', date: '2026-10-26', endDate: '2026-10-28', start: '09:00', end: '23:00' }], '{"v":1,"nextId":3,"sessions":[{"id":"s2","label":"AHG Session 2","date":"2026-10-26","endDate":"2026-10-28","start":"09:00","end":"23:00"}]}']);
  check('it is read back as stored, and the schema version is still 1', (() => { const p = S.parseAdhocSessionsProperty_(sessionsProperty([CALL1, SESSION2], 3)); return [p.status, p.config.v, p.config.sessions]; })(), ['ok', 1, [CALL1, SESSION2]]);
  check('a range across a month and across a year', [ok([session('s1', '2026-10-30', '2026-11-01')])[0].endDate, ok([session('s1', '2026-12-31', '2027-01-02')])[0].endDate], ['2026-11-01', '2027-01-02']);
  check('an end date before the start date: refused', v([session('s1', '2026-10-28', '2026-10-26')]).errors, ['2026-10-28: the end date is before the start date.']);
  check('an end date that is not a date: refused', [v([session('s1', '2026-10-26', '2026-10-32')]).errors, v([session('s1', '2026-10-26', '28 Oct')]).errors], [['2026-10-26: the end date "2026-10-32" is not a date (YYYY-MM-DD).'], ['2026-10-26: the end date "28 Oct" is not a date (YYYY-MM-DD).']]);

  // Start and end times.
  check('on one day the planned end must be after the planned start', [v([session('s1', '2026-09-22', '', '18:00', '15:00')]).errors, v([session('s1', '2026-09-22', '2026-09-22', '18:00', '18:00')]).errors],
    [['2026-09-22: the planned end must be after the planned start.'], ['2026-09-22: the planned end must be after the planned start.']]);
  check('over several days the start is on the first day and the end on the last: an end earlier in the day than the start is fine', [v([session('s1', '2026-10-26', '2026-10-28', '15:00', '12:00')]).errors, v([session('s1', '2026-10-26', '2026-10-28', '09:00', '09:00')]).errors], [[], []]);
  check('the times are still checked as times', v([session('s1', '2026-10-26', '2026-10-28', '9', '25:00')]).errors, ['2026-10-26: the planned start "9" is not a time (HH:mm).', '2026-10-26: the planned end "25:00" is not a time (HH:mm).']);

  // Overlap.
  const A = (extra) => Object.assign(session('s1', '2026-10-26', '2026-10-28', '', '', 'First'), extra || {});
  const B = (date, endDate, extra) => Object.assign(session('s2', date, endDate || '', '', '', 'Second'), extra || {});
  check('overlapping ranges: refused -- a session inside another, one starting inside another, two starting on the same day', [v([A(), B('2026-10-27')]).errors, v([A(), B('2026-10-27', '2026-10-30')]).errors, v([A(), B('2026-10-26', '2026-10-27')]).errors, v([A(), B('2026-10-26')]).errors],
    [['First and Second overlap: their days must not overlap.'], ['First and Second overlap: their days must not overlap.'], ['First and Second overlap: their days must not overlap.'], ['First and Second overlap: their days must not overlap.']]);
  check('adjacent ranges that do not overlap: accepted, in order', [ok([B('2026-10-29', '2026-10-30'), A()]).map((x) => x.id), ok([A(), B('2026-10-29')]).map((x) => x.id), ok([B('2026-10-25'), A()]).map((x) => x.id)], [['s1', 's2'], ['s1', 's2'], ['s2', 's1']]);
  check('a session that begins on the last day of the one before: only when the planned end of the first and the planned start of the second say so', [
    v([A(), B('2026-10-28')]).errors, v([A({ end: '12:00' }), B('2026-10-28')]).errors, v([A({ end: '12:00' }), B('2026-10-28', '', { start: '11:00' })]).errors,
    ok([A({ end: '12:00' }), B('2026-10-28', '', { start: '12:00' })]).map((x) => x.id), ok([A({ end: '12:00' }), B('2026-10-28', '2026-10-30', { start: '14:00' })]).map((x) => x.id)],
    [['Second begins on the last day of First: enter the planned end of the first and the planned start of the second, so that their order is known.'],
      ['Second begins on the last day of First: enter the planned end of the first and the planned start of the second, so that their order is known.'],
      ['First: the planned end (12:00) is after the start of Second on that day (11:00).'], ['s1', 's2'], ['s1', 's2']]);
  check('two sessions of one day on the same day: the rules and messages they always had', v([session('s1', '2026-09-22', '', '', '', 'a'), session('s2', '2026-09-22', '', '', '', 'b')]).errors,
    ['Two sessions are on 22 Sep: enter a planned start for each of them, so that their order is known.']);
  check('three sessions, the middle one of several days: every neighbour is checked', v([session('s1', '2026-10-20', '', '', '', 'x'), A({ id: 's2' }), session('s3', '2026-10-28', '', '', '', 'z')]).errors.length, 1);

  // Order and ids.
  check('order: by first day, then by planned start', ok([session('s3', '2026-10-26', '2026-10-28', '', '', 'c'), session('s1', '2026-11-02', '2026-11-03', '', '', 'a'), session('s2', '2026-09-22', '', '', '', 'b')]).map((x) => x.id), ['s2', 's3', 's1']);
  const edit = S.applyAdhocSessionsEdit_(sessionsProperty([CALL1, SESSION2], 3), [Object.assign({}, SESSION2, { date: '2026-09-01', endDate: '2026-09-03' }), CALL1, { id: '', label: 'Third', date: '2026-11-10', endDate: '2026-11-12', start: '', end: '' }]);
  check('ids are stable: a session moved to earlier dates keeps its id, a new one gets the next id', [edit.ok, edit.config.sessions.map((x) => [x.id, x.date, x.endDate || '']), edit.config.nextId],
    [true, [['s2', '2026-09-01', '2026-09-03'], ['s1', '2026-09-22', ''], ['s3', '2026-11-10', '2026-11-12']], 4]);
  check('making a session of several days one of one day again removes its endDate', S.applyAdhocSessionsEdit_(sessionsProperty([SESSION2], 3), [Object.assign({}, SESSION2, { endDate: '' })]).config.sessions, [{ id: 's2', label: 'AHG Session 2', date: '2026-10-26', start: '09:00', end: '23:00' }]);

  // Labels and date ranges.
  const label = S.defaultAdhocSessionLabel_;
  check('default labels: one day, a range in a month, across months, across years', [label('2026-09-22'), label('2026-09-22', '2026-09-22'), label('2026-10-26', '2026-10-28'), label('2026-10-30', '2026-11-01'), label('2026-12-31', '2027-01-02')],
    ['22 Sep', '22 Sep', '26–28 Oct', '30 Oct – 1 Nov', '31 Dec – 2 Jan']);
  check('every default label fits the label limit', ['26–28 Oct', '30 Oct – 1 Nov', '31 Dec – 2 Jan'].map((x) => x.length <= 16), [true, true, true]);
  check('the label of the user is authoritative', [S.adhocSessionLabel_(SESSION2), S.adhocSessionLabel_(session('s2', '2026-10-26', '2026-10-28'))], ['AHG Session 2', '26–28 Oct']);
  check('two sessions whose default labels are equal are still refused', v([session('s1', '2026-10-26', '2026-10-28'), session('s2', '2027-10-26', '2027-10-28')]).errors, ['The label "26–28 Oct" is used by more than one session. Give each session its own label.']);
  const range = S.adhocDateRangeText_;
  check('THE way a range of days is written', [range('2026-09-22', '2026-09-22'), range('2026-10-26', '2026-10-28'), range('2026-10-30', '2026-11-01'), range('2026-12-31', '2027-01-02')],
    ['September 22, 2026', 'October 26–28, 2026', 'October 30 – November 1, 2026', 'December 31, 2026 – January 2, 2027']);
  check('a missing or earlier last day is the first day alone; no first day is nothing', [range('2026-09-22', ''), range('2026-09-22', '2026-09-20'), range('', '2026-09-22'), range('x', 'y')], ['September 22, 2026', 'September 22, 2026', '', '']);
  check('the days of a session', [S.adhocSessionDatesText_(CALL1), S.adhocSessionDatesText_(SESSION2), S.adhocSessionLastDay_(CALL1), S.adhocSessionLastDay_(SESSION2), S.isAdhocMultiDaySession_(CALL1), S.isAdhocMultiDaySession_(SESSION2), S.adhocSessionDayCount_(CALL1), S.adhocSessionDayCount_(SESSION2),
    S.adhocSessionDayCount_(session('s1', '2026-12-31', '2027-01-02')), S.adhocSessionDayCount_(session('s1', '2026-10-24', '2026-10-26'))],
    ['September 22, 2026', 'October 26–28, 2026', '2026-09-22', '2026-10-28', false, true, 1, 3, 3, 3]);
  check('it is one helper: no other code puts two long dates together (source)', [(CODE.match(/adhocLongDate_\([^)]*\) \+ ' – '|' – ' \+ adhocLongDate_\(/g) || []).length, (CODE.match(/adhocDateRangeText_\(/g) || []).length], [0, 3]);
}

// ================================================================ 2. Configure Sessions

/** Runs the Configure Sessions script against a minimal DOM. */
function runSessionsDialog(html, server) {
  const scripts = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
  const element = (tag) => { const e = { tag: tag, children: [], style: {}, value: '', textContent: '' }; e.appendChild = (c) => { e.children.push(c); return c; }; e.removeChild = (c) => { e.children.splice(e.children.indexOf(c), 1); return c; }; return e; };
  const byId = {};
  (html.match(/id="([A-Za-z]+)"/g) || []).map((m) => m.slice(4, -1)).forEach((id) => { byId[id] = element(id); });
  const calls = []; let closed = 0; let handlers = {};
  const runner = { withSuccessHandler: (fn) => { handlers.success = fn; return runner; }, withFailureHandler: (fn) => { handlers.failure = fn; return runner; } };
  runner.saveAdhocSessionsConfiguration = (rows) => { calls.push(JSON.parse(JSON.stringify(rows))); const h = handlers; handlers = {}; h.success(JSON.parse(JSON.stringify(server.saveAdhocSessionsConfiguration(rows)))); };
  const ctx = { document: { getElementById: (id) => byId[id], createElement: element }, google: { script: { run: runner, host: { close: () => { closed++; } } } }, alert: () => {} };
  vm.createContext(ctx);
  vm.runInContext(scripts[0].replace(/^<script>|<\/script>$/g, ''), ctx);
  const inputs = (i) => byId.rows.children[i].children.slice(0, 5).map((td) => td.children[0]);
  return { ctx: ctx, byId: byId, calls: calls, closed: () => closed, inputs: inputs, values: () => byId.rows.children.map((tr, i) => inputs(i).map((x) => x.value)) };
}

console.log('2. Configure Sessions');
{
  const r = report();
  const before = JSON.stringify(r.docProps._store);
  r.s.configureAdhocSessions();
  const html = r.ui.dialogs[0].html;
  check('the columns say what each field is', (html.match(/<thead><tr>(.*?)<\/tr><\/thead>/) || [])[1], '<th>Label</th><th>Start date</th><th>End date</th><th>Start time</th><th>End time (final cut-off)</th><th></th>');
  check('the hint explains a session of several days and what the times mean', /A session can span several days: enter its last day as End date, or leave End date empty for a session of one day\. .* the start time is on the first day, the end time is the final cut-off on the last day\./.test(html), true);
  const d = runSessionsDialog(html, r.s);
  check('the stored sessions are shown: label, start date, end date, start time, end time', d.values(), [['AHG Call 1', '2026-09-22', '', '15:00', '18:00'], ['AHG Session 2', '2026-10-26', '2026-10-28', '09:00', '23:00']]);
  check('the field types: text, date, date, time, time', d.inputs(1).map((x) => x.type), ['text', 'date', 'date', 'time', 'time']);
  check('opening the dialog and Cancel write nothing', (() => { d.inputs(0)[2].value = '2026-09-24'; d.ctx.google.script.host.close(); return [d.calls.length, JSON.stringify(r.docProps._store) === before, r.writes]; })(), [0, true, []]);

  d.inputs(0)[2].value = '';
  d.inputs(1)[2].value = '2026-10-29';
  d.ctx.saveSessions();
  check('Save sends endDate only for a session of several days; a session of one day is sent as it always was', d.calls[0], [{ id: 's1', label: 'AHG Call 1', date: '2026-09-22', start: '15:00', end: '18:00' }, { id: 's2', label: 'AHG Session 2', date: '2026-10-26', start: '09:00', end: '23:00', endDate: '2026-10-29' }]);
  check('it is stored, once, with the ids unchanged', [r.writes, JSON.parse(r.docProps.getProperty(KEY)).sessions.map((x) => [x.id, x.date, x.endDate || ''])], [['set ' + KEY], [['s1', '2026-09-22', ''], ['s2', '2026-10-26', '2026-10-29']]]);
  check('saving the same again writes nothing', (() => { const n = r.writes.length; r.s.configureAdhocSessions(); const e = runSessionsDialog(r.ui.dialogs[1].html, r.s); e.ctx.saveSessions(); return [r.writes.length === n, e.closed()]; })(), [true, 1]);
  check('an end date equal to the start date is not sent as a range', (() => { r.s.configureAdhocSessions(); const e = runSessionsDialog(r.ui.dialogs[2].html, r.s); e.inputs(0)[2].value = '2026-09-22'; return JSON.stringify(e.ctx.collect()[0]); })(), '{"id":"s1","label":"AHG Call 1","date":"2026-09-22","start":"15:00","end":"18:00"}');
  check('a refused save shows the reason and stores nothing', (() => { const n = r.writes.length; const x = r.s.saveAdhocSessionsConfiguration([Object.assign({}, SESSION2, { endDate: '2026-10-20' })]); return [x.ok, x.errors, r.writes.length === n]; })(), [false, ['AHG Session 2: the end date is before the start date.'], true]);
  check('what the session functions hand on: endDate only for a session of several days', r.s.getAdhocSessions_().map((x) => Object.keys(x).join(',')), ['id,label,date,start,end,displayLabel', 'id,label,date,endDate,start,end,displayLabel']);
}

// ================================================================ 3. TDoc assignment

console.log('3. TDoc assignment: one logical session, its cut-off on the last day');
{
  check('the cut-off: the planned end on the last day, else the end of the last day', [S.adhocSessionCutoff_(SESSION2), S.adhocSessionCutoff_(session('s2', '2026-10-26', '2026-10-28')), S.adhocSessionCutoff_(CALL1), S.adhocSessionCutoff_(session('s1', '2026-09-22'))],
    ['2026-10-28T23:00:00', '2026-10-28T23:59:59', '2026-09-22T18:00:00', '2026-09-22T23:59:59']);
  // Upload times as the Portal gives them (UTC). The report's zone is Europe/Berlin: UTC+2 until 25 October 2026, UTC+1 after.
  const r = report();
  const resolver = () => r.s.makeAdhocTdocSessionResolver_();
  const td = (id, utc) => ({ row: [id], tdocCol: 0, uploaded: utc ? { at: utc.replace(' ', 'T'), dateOnly: null, unreadable: false } : { at: null, dateOnly: null, unreadable: false } });
  const where = (utc) => { const a = resolver().automatic(td('S4aA269001', utc)); return a.sessionId || a.reason; };
  check('uploaded before the first session: the first session', where('2026-09-20 10:00:00'), 's1');
  check('exactly at the cut-off of the first session (18:00 in Berlin, 16:00 UTC): still the first; one second later: the next session', [where('2026-09-22 16:00:00'), where('2026-09-22 16:00:01')], ['s1', 's2']);
  check('uploaded between the sessions: the next one', where('2026-10-10 08:00:00'), 's2');
  check('uploaded on 26, 27 and 28 October: the one session of those days', [where('2026-10-26 10:00:00'), where('2026-10-27 12:00:00'), where('2026-10-28 08:00:00'), where('2026-10-28 21:59:59')], ['s2', 's2', 's2', 's2']);
  check('exactly at the final cut-off (23:00 in Berlin on the 28th, 22:00 UTC): in the session; one second later: after the last session', [where('2026-10-28 22:00:00'), where('2026-10-28 22:00:01')], ['s2', 'after-last']);
  check('the conversion follows the change from summer to winter time on 25 October: 21:00:01 UTC on the 28th is 22:00:01 in Berlin, not 23:00:01', [where('2026-10-28 21:00:01'), r.s.convertUtcWallClockToReportZone_('2026-10-28T21:00:01'),
    r.s.convertUtcWallClockToReportZone_('2026-10-24T21:00:01'), r.s.convertUtcWallClockToReportZone_('2026-10-25T00:30:00'), r.s.convertUtcWallClockToReportZone_('2026-10-25T01:30:00')],
    ['s2', '2026-10-28T22:00:01', '2026-10-24T23:00:01', '2026-10-25T02:30:00', '2026-10-25T02:30:00']);
  check('without a planned end the whole last day belongs to the session', (() => {
    const x = report(Object.assign({}, ADHOC, { [KEY]: sessionsProperty([CALL1, Object.assign({}, SESSION2, { end: '' })]) }));
    const at = (utc) => { const a = x.s.makeAdhocTdocSessionResolver_().automatic(td('S4aA269001', utc)); return a.sessionId || a.reason; };
    return [at('2026-10-28 22:59:59'), at('2026-10-28 23:00:00')];
  })(), ['s2', 'after-last']);
  check('a TDoc that is not uploaded has no session', where(''), 'not-uploaded');
  check('with the first day entered alone (no endDate) the 27th and 28th would be after the last session: the endDate is what makes it one session', (() => {
    const x = report(Object.assign({}, ADHOC, { [KEY]: sessionsProperty([CALL1, session('s2', '2026-10-26', '', '09:00', '23:00', 'AHG Session 2')]) }));
    const a = x.s.makeAdhocTdocSessionResolver_().automatic(td('S4aA269001', '2026-10-27 12:00:00'));
    return a.sessionId || a.reason;
  })(), 'after-last');

  // Manual assignments are what they were.
  r.docProps._store.ADHOC_TDOC_SESSIONS = r.s.serializeAdhocTdocSessions_({ clock: 'utc', tdocs: { S4aA269001: { mode: 'add', sessions: ['s2'] }, S4aA269002: { mode: 'set', sessions: ['s1'] }, S4aA269003: { mode: 'set', sessions: [] } } });
  const eff = (id, utc) => resolver().effective(td(id, utc));
  check('manual add, set and "no session" work as before', [eff('S4aA269001', '2026-09-20 10:00:00'), eff('S4aA269002', '2026-10-27 12:00:00'), eff('S4aA269003', '2026-10-27 12:00:00'), resolver().cell(td('S4aA269001', '2026-09-20 10:00:00'))],
    [['s1', 's2'], ['s1'], [], 'A01, A02']);
  check('no upload time is stored: the property holds the clock and the manual assignments only', r.docProps._store.ADHOC_TDOC_SESSIONS, '{"v":1,"clock":"utc","add":{"s2":["S4aA269001"]},"set":{"":["S4aA269003"],"s1":["S4aA269002"]}}');
  check('a stored compatibility clock is still honoured: the times are then compared as they stand', (() => {
    r.docProps._store.ADHOC_TDOC_SESSIONS = r.s.serializeAdhocTdocSessions_({ clock: 'session', tdocs: {} });
    return [where('2026-10-28 23:00:00'), where('2026-10-28 23:00:01')];
  })(), ['s2', 'after-last']);
}

// ================================================================ 4. Teams attendance

console.log('4. Teams attendance: one record per day of the session');
const r = report();
{
  // The parser's date check.
  const parse = (date, options) => { const p = S.parseTeamsAttendanceReport_(exportFor(date, DAY1), options); return p.ok ? p.attendance.summary.start.slice(0, 10) : p.error.code + ': ' + p.error.message; };
  const RANGE = { expectedDate: '2026-10-26', expectedEndDate: '2026-10-28' };
  check('the export of each day of the session is accepted', ['2026-10-26', '2026-10-27', '2026-10-28'].map((d) => parse(d, RANGE)), ['2026-10-26', '2026-10-27', '2026-10-28']);
  check('the day before and the day after are refused, and the message names the session\'s days', [parse('2026-10-25', RANGE), parse('2026-10-29', RANGE)],
    ['SESSION_DATE_MISMATCH: This attendance export is for 2026-10-25, not for the session of 2026-10-26 to 2026-10-28.', 'SESSION_DATE_MISMATCH: This attendance export is for 2026-10-29, not for the session of 2026-10-26 to 2026-10-28.']);
  check('for a session of one day the check and its message are what they were', [parse('2026-09-22', { expectedDate: '2026-09-22' }), parse('2026-09-23', { expectedDate: '2026-09-22' }), parse('2026-09-22', { expectedDate: '2026-09-22', expectedEndDate: '2026-09-22' })],
    ['2026-09-22', 'SESSION_DATE_MISMATCH: This attendance export is for 2026-09-23, not for the session on 2026-09-22.', '2026-09-22']);
  check('an end date that is not a date on or after the first day is an error of the caller, not a guess', [parse('2026-10-26', { expectedDate: '2026-10-26', expectedEndDate: '2026-10-20' }).split(':')[0], parse('2026-10-26', { expectedDate: '2026-10-26', expectedEndDate: 'x' }).split(':')[0]], ['INVALID_EXPECTED_DATE', 'INVALID_EXPECTED_DATE']);
  check('a date that can be read month-first or day-first is read as the one that is a day of the session', (() => {
    const p = S.parseTeamsAttendanceReport_(exportFor('2026-10-26', DAY1).split('10/26/26').join('11/10/26'), { expectedDate: '2026-10-09', expectedEndDate: '2026-10-12' });
    return p.ok ? [p.attendance.summary.start.slice(0, 10), p.attendance.source.dateOrder] : p.error.code;
  })(), ['2026-10-11', 'day-first']);

  // Imports.
  check('the preview names the session with its days, the day of the export, and nothing to replace', (() => { const p = r.preview('s2', '2026-10-27', DAY2); return [p.ok, p.preview.session, p.preview.date, p.preview.attendees, p.preview.replaces, p.preview.otherDays, p.preview.unchanged, r.writes]; })(),
    [true, 'AHG Session 2, October 26–28, 2026', 'October 27, 2026', 5, '', '', false, []]);
  check('26 October is imported', (() => { const x = r.import('s2', '2026-10-26', DAY1); return [x.ok, x.message]; })(), [true, 'The attendance of AHG Session 2 (October 26, 2026) was imported (3 attendees).']);
  check('28 October, then 27 October: accepted in any order', [r.import('s2', '2026-10-28', DAY3).message, r.import('s2', '2026-10-27', DAY2).message],
    ['The attendance of AHG Session 2 (October 28, 2026) was imported (2 attendees).', 'The attendance of AHG Session 2 (October 27, 2026) was imported (5 attendees).']);
  check('25 and 29 October are refused, and nothing is stored for them', (() => { const n = r.writes.length; return [r.preview('s2', '2026-10-25', DAY1), r.import('s2', '2026-10-29', DAY1), r.writes.length === n]; })(),
    [{ ok: false, error: 'This attendance export is for 2026-10-25, not for the session of 2026-10-26 to 2026-10-28.' }, { ok: false, error: 'This attendance export is for 2026-10-29, not for the session of 2026-10-26 to 2026-10-28.' }, true]);
  check('the three records are under one session, by day', r.records('s2'), [['s2d20261026', '2026-10-26', 3], ['s2d20261027', '2026-10-27', 5], ['s2d20261028', '2026-10-28', 2]]);
  check('each record has its own stored value; the index names them', [r.attendanceKeys().map((k) => k.replace(/_\d+_0$/, '')), Object.keys(JSON.parse(r.docProps.getProperty(ATT)).entries).sort(), r.s.adhocAttendanceSessionIds_(r.docProps)],
    [[ATT, ATT + '_s2d20261026', ATT + '_s2d20261027', ATT + '_s2d20261028'], ['s2d20261026', 's2d20261027', 's2d20261028'], ['s2']]);
  check('the attendance of the first session (one day) is stored as it always was: under the session id', (() => { const x = r.import('s1', '2026-09-22'); return [x.message.slice(0, 62), r.records('s1'), Object.keys(JSON.parse(r.docProps.getProperty(ATT)).entries).indexOf('s1') !== -1]; })(),
    ['The attendance of AHG Call 1 was imported (23 attendees).', [['s1', '2026-09-22', 23]], true]);

  // Importing a day again.
  const raw = r.state().raw;
  const p = r.preview('s2', '2026-10-27', DAY2.slice(0, 2));
  check('importing 27 October again: the preview says that it replaces that day, and names the other days', [p.preview.replaces, p.preview.otherDays, p.preview.unchanged],
    ['This replaces the attendance already imported for October 27, 2026 of this session (5 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE").', 'October 26, 2026; October 28, 2026', false]);
  check('it replaces 27 October only: the records of the 26th and the 28th, and of the first session, are byte for byte what they were', (() => {
    const x = r.import('s2', '2026-10-27', DAY2.slice(0, 2));
    const now = r.state().raw;
    return [x.ok, r.records('s2'), now.s2d20261026 === raw.s2d20261026, now.s2d20261028 === raw.s2d20261028, now.s1 === raw.s1, now.s2d20261027 === raw.s2d20261027];
  })(), [true, [['s2d20261026', '2026-10-26', 3], ['s2d20261027', '2026-10-27', 2], ['s2d20261028', '2026-10-28', 2]], true, true, true, false]);
  check('the same day is never stored twice', r.attendanceKeys().filter((k) => /s2d20261027/.test(k)).length, 1);
  check('importing the same file again changes nothing', (() => { const n = r.writes.length; const x = r.import('s2', '2026-10-27', DAY2.slice(0, 2)); return [x.changed, x.message, r.writes.length === n]; })(),
    [false, 'This attendance is already imported for AHG Session 2 (October 27, 2026). Nothing was changed.', true]);
  r.import('s2', '2026-10-27', DAY2);

  // The dialog.
  const model = r.s.buildTeamsAttendanceDialogModel_();
  check('the dialog is given, per session, its days and what is imported for each day -- counts and the meeting title, no attendee', [model.sessions.map((x) => [x.id, x.label, x.records.map((y) => y.day + ' ' + y.label)]), model.sessions[1].imported, model.sessions[0].imported],
    [[['s1', 'AHG Call 1, September 22, 2026', ['2026-09-22 September 22, 2026']], ['s2', 'AHG Session 2, October 26–28, 2026', ['2026-10-26 October 26, 2026', '2026-10-27 October 27, 2026', '2026-10-28 October 28, 2026']]],
      'October 26, 2026: 3 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"; October 27, 2026: 5 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"; October 28, 2026: 2 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"',
      '23 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"']);
  check('no attendee is in the dialog model', /Organizer|Rivera|Wong|example\.(com|org)/.test(JSON.stringify(model)), false);
  r.s.importTeamsAttendance();
  const html = r.ui.dialogs[r.ui.dialogs.length - 1].html;
  check('the dialog: a day to remove is chosen when a session has several, the preview shows the other days, and removal is asked for with the day (source)', [/<select id="removeDay"/.test(html), /previewRow\('Other days already imported', p\.otherDays\);/.test(html),
    /removeTeamsAttendanceImport\(s\.id, el\('removeDay'\)\.value\)/.test(html), /removeTeamsAttendanceImport\(s\.id\);/.test(html), /For a session of several days, import the report of each day/.test(html), /innerHTML/.test(html)], [true, true, true, true, true, false]);

  // Removing.
  check('removing without saying which day, when there are several: refused, nothing removed', [r.s.removeTeamsAttendanceImport('s2'), r.records('s2').length], [{ ok: false, error: 'This session has attendance for several days. Choose the day to remove.' }, 3]);
  check('removing a day that is not imported: refused', r.s.removeTeamsAttendanceImport('s2', '2026-10-25'), { ok: false, error: 'This session has no imported attendance for that day.' });
}

// ================================================================ 5. rendering

const tableLook = (t) => {
  const cells = []; for (let i = 0; i < t.getNumRows(); i++) { const row = []; for (let k = 0; k < t.getRow(i).getNumCells(); k++) row.push(t.getRow(i).getCell(k)); cells.push(row); }
  return [cells[0].map((c) => c._bold).join(','), cells.slice(1).reduce((n, row) => n + row.filter((c) => c._bold).length, 0), cells[0].map((c) => c._background).filter((v, i, l) => l.indexOf(v) === i).join(','),
    cells.every((row) => row.map((c) => c._width).join(',') === '159,122,187'), Array.from({ length: t.getNumRows() }, (_, i) => t.getRow(i)._minHeight).every((h) => h === 0), cells.every((row) => row.every((c) => c._spacing[0] === 0 && c._spacing[1] === 0))];
};
const GOOD_LOOK = ['true,true,true', 0, '#D9EAF7', true, true, true];

console.log('5. rendering, formatting, refresh and rebuild');
{
  const EXPECTED = ['Attendance',
    'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: AHG Call 1, September 22, 2026', 'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 25', 'Average attendance: 2:31:07', 'Attendees', '[table 23]',
    'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: AHG Session 2, October 26–28, 2026', 'Date: October 26, 2026', 'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 3', 'Average attendance: 2:31:07', 'Attendees', '[table 3]',
    'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: AHG Session 2, October 26–28, 2026', 'Date: October 27, 2026', 'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 5', 'Average attendance: 2:31:07', 'Attendees', '[table 5]',
    'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: AHG Session 2, October 26–28, 2026', 'Date: October 28, 2026', 'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 2', 'Average attendance: 2:31:07', 'Attendees', '[table 2]'];
  check('one Attendance section: the sessions in order, and within the session of several days its days in order -- whatever the order of the imports', r.lines(), EXPECTED);
  check('every block names its logical session with its days; a block of the session of several days also says which day it is; the block of the session of one day has no Date line', [r.lines().filter((t) => /^Session:/.test(t)).length, r.lines().filter((t) => /^Date:/.test(t)), r.lines().slice(1, 11).filter((t) => /^Date:/.test(t))],
    [4, ['Date: October 26, 2026', 'Date: October 27, 2026', 'Date: October 28, 2026'], []]);
  check('one generated container, one heading, four tables', [r.body._children.filter((e) => e.getType() === 'PARAGRAPH' && e.getText() === 'Attendance' && e.getHeading() !== 'NORMAL').length, r.tables().length], [1, 4]);
  check('all four tables have the table look: header bold and shaded, no attendee row bold, widths, compact rows', r.tables().map(tableLook), [GOOD_LOOK, GOOD_LOOK, GOOD_LOOK, GOOD_LOOK]);
  check('the tables follow a bold "Attendees" line, and the paragraphs are bold only where they are labels', [r.tables().map((t) => r.container()[r.container().indexOf(t) - 1]._bold), r.container().filter((e) => e.getType() === 'PARAGRAPH' && e._bold && e.getText() !== '').map((p) => p.getText()).filter((t, i, l) => l.indexOf(t) === i)],
    [[true, true, true, true], ['Statistics', 'Attendees']]);
  check('the e-mail cells are mailto: links in every table', r.tables().map((t) => { let n = 0; for (let i = 1; i < t.getNumRows(); i++) { const c = t.getRow(i).getCell(2); if (c.getText() && c._links.length === 1 && c._links[0][2] === 'mailto:' + c.getText()) n++; } return n; }).slice(1), [2, 2, 2]);
  check('the attendees of each day are in that day\'s table', r.tables().slice(1).map((t) => rowsOf(t).slice(1).map((row) => row[0]).join(', ')), ['Alex Organizer, Sam Rivera, Lee Wong', 'Alex Organizer, Sam Rivera, Noor Late, Dana Twin, Dana Twin', 'Alex Organizer, Omar Open']);
  check('the Session line and the Date line are built in one place (source)', [(CODE.match(/add\('Session', session\.label \? session\.label \+ ', ' \+ dates : dates\);/g) || []).length, (CODE.match(/if \(sayDay\) add\('Date', adhocLongDate_\(record\.day\)\);/g) || []).length], [1, 1]);

  const before = r.snapshot();
  r.s.refreshAdhocAttendanceSection();
  r.s.refreshAdhocAttendanceSection();
  check('Refresh Attendance Section, twice: the same document, all three days still there', [r.snapshot() === before, r.records('s2').length, r.tables().length], [true, 3, 4]);
  check('a rebuild writes all records again, in the same order, with the look', (() => {
    r.s.beginAdhocAttendanceRebuild_();
    r.body.clear();
    r.body.appendParagraph('5.11 Close of the session').setHeading(r.s.DocumentApp.ParagraphHeading.HEADING2);
    const note = r.s.finishAdhocAttendanceRebuild_(r.body, { active: true });
    return [note, r.lines(), r.tables().map(tableLook), r.records('s2').length];
  })(), ['', EXPECTED, [GOOD_LOOK, GOOD_LOOK, GOOD_LOOK, GOOD_LOOK], 3]);

  // A record that cannot be read.
  const d = report();
  d.import('s2', '2026-10-26', DAY1);
  d.import('s2', '2026-10-27', DAY2);
  const index = JSON.parse(d.docProps.getProperty(ATT));
  d.docProps._store[ATT + '_s2d20261027_' + index.entries.s2d20261027.g + '_0'] = '{"t":"tampered"}';
  check('one day whose record is damaged: named with its session and day; the other day is still read', [d.state().problems, d.records('s2'), (() => { try { d.s.refreshAdhocAttendanceSection(); return ''; } catch (e) { return e.message; } })()],
    [['s2d20261027'], [['s2d20261026', '2026-10-26', 3], ['s2d20261027', '2026-10-27', null]], 'The stored attendance of AHG Session 2 (October 27, 2026) cannot be read. Import it again or remove it; the report was not changed.']);
  check('a rebuild leaves the damaged day out, says so, and writes the other day: no empty block', (() => {
    const copy = d.snapshot();
    d.body.clear();
    d.body.appendParagraph('5.11 Close of the session').setHeading(d.s.DocumentApp.ParagraphHeading.HEADING2);
    const note = d.s.finishAdhocAttendanceRebuild_(d.body, { active: true });
    return [note, d.lines().filter((x) => /^Date:|^\[table/.test(x)), copy.length > 0];
  })(), ['\n\n⚠️ The stored attendance of AHG Session 2 (October 27, 2026) cannot be read and was left out. Import it again.', ['Date: October 26, 2026', '[table 3]'], true]);
  check('importing that day again repairs it; the other day is untouched', (() => { const raw = d.state().raw.s2d20261026; const x = d.import('s2', '2026-10-27', DAY2); return [x.ok, d.state().problems, d.records('s2').map((y) => y[2]), d.state().raw.s2d20261026 === raw]; })(), [true, [], [3, 5], true]);

  // Removing a day; the delete guard.
  check('a session with attendance for any day is not removed', (() => { const x = r.s.saveAdhocSessionsConfiguration([CALL1]); return [x.ok, x.errors[0].indexOf('AHG Session 2') !== -1 && x.errors[0].indexOf('imported attendance') !== -1, r.s.getAdhocSessions_().length]; })(), [false, true, 2]);
  check('removing one day removes that day only', (() => { const x = r.s.removeTeamsAttendanceImport('s2', '2026-10-27'); return [x.ok, x.message, r.records('s2'), r.records('s1').length, r.lines().filter((t) => /^Date:/.test(t))]; })(),
    [true, 'The attendance of AHG Session 2 (October 27, 2026) was removed.', [['s2d20261026', '2026-10-26', 3], ['s2d20261028', '2026-10-28', 2]], 1, ['Date: October 26, 2026', 'Date: October 28, 2026']]);
  check('the session still cannot be removed while a day is left', [r.s.removeTeamsAttendanceImport('s2', '2026-10-26').ok, r.s.saveAdhocSessionsConfiguration([CALL1]).ok, r.s.removeTeamsAttendanceImport('s2').message, r.s.saveAdhocSessionsConfiguration([CALL1]).ok],
    [true, false, 'The attendance of AHG Session 2 was removed.', true]);
  check('no chunk of the removed records is left behind', r.attendanceKeys().filter((k) => /s2/.test(k)), []);
}

// ================================================================ records stored before sessions could span several days

console.log('   a record stored under the session id (as before) is still read');
{
  const x = report();
  const payload = (date, people) => x.s.compactAdhocAttendance_(x.s.parseTeamsAttendanceReport_(exportFor(date, people), { expectedDate: date }).attendance);
  x.s.writeAdhocAttendanceValue_(x.docProps, 's2', payload('2026-10-27', DAY2));
  check('it belongs to the day its own Teams start time says', x.records('s2'), [['s2', '2026-10-27', 5]]);
  check('it is rendered under its session, with its day', (() => { x.s.refreshAdhocAttendanceSection(); return x.lines().slice(0, 5); })(), ['Attendance', 'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: AHG Session 2, October 26–28, 2026', 'Date: October 27, 2026']);
  check('importing that day again replaces THAT record, under its key: the day is not stored twice', (() => { const p = x.preview('s2', '2026-10-27', DAY2.slice(0, 3)); x.import('s2', '2026-10-27', DAY2.slice(0, 3)); return [p.preview.replaces.slice(0, 82), x.records('s2')]; })(),
    ['This replaces the attendance already imported for October 27, 2026 of this session', [['s2', '2026-10-27', 3]]]);
  check('another day is stored beside it', (() => { x.import('s2', '2026-10-26', DAY1); return x.records('s2'); })(), [['s2d20261026', '2026-10-26', 3], ['s2', '2026-10-27', 3]]);
  check('removing by day finds it', [x.s.removeTeamsAttendanceImport('s2', '2026-10-27').ok, x.records('s2')], [true, [['s2d20261026', '2026-10-26', 3]]]);
  check('a record under the session id that cannot be read, beside a day record: it is listed as a record that cannot be read, and can be removed without naming a day; the day record stays', (() => {
    const y = report();
    y.s.writeAdhocAttendanceValue_(y.docProps, 's2', '{"not":"attendance"}');
    y.import('s2', '2026-10-26', DAY1);
    const listed = y.s.buildTeamsAttendanceDialogModel_().sessions[1].records.map((z) => z.day + '|' + z.label + '|' + z.text);
    const removed = y.s.removeTeamsAttendanceImport('s2', '');
    return [y.state().problems.length, listed, removed.ok, y.records('s2')];
  })(), [0, ['|a record that cannot be read|stored, but it cannot be read', '2026-10-26|October 26, 2026|3 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"'], true, [['s2d20261026', '2026-10-26', 3]]]);
  check('an index entry that is neither a session id nor a day of a session is not accepted', ['s2d2026102', 's2x', 'd20261027', 's0d20261027', 's2d20261027x'].map((id) => x.s.adhocAttendanceRecordParts_(id)).concat([x.s.adhocAttendanceRecordParts_('s12d20261027'), x.s.adhocAttendanceRecordParts_('s3')]),
    [null, null, null, null, null, { sessionId: 's12', day: '2026-10-27' }, { sessionId: 's3', day: '' }]);
}

// ================================================================ privacy

console.log('   what is stored');
{
  const x = report();
  x.s.confirmTeamsAttendanceImport('s1', base64(FX.buildExport()), x.s.previewTeamsAttendanceImport('s1', base64(FX.buildExport())).token);
  x.import('s2', '2026-10-26', DAY1);
  x.import('s2', '2026-10-27', DAY2);
  const all = x.attendanceKeys().map((k) => x.docProps._store[k]).join('\n');
  check('per record: the Teams summary and [name, company, e-mail] rows, nothing else', Object.keys(x.state().raw).sort().map((k) => { const v = JSON.parse(x.state().raw[k]); return k + ' ' + Object.keys(v).join('') + ':' + v.p.every((row) => row.length === 3 && row.every((c) => typeof c === 'string')); }), ['s1 tnsedap:true', 's2d20261026 tnsedap:true', 's2d20261027 tnsedap:true']);
  check('no Participant ID, no telephone number, no join or leave time, no duration of an attendee, no role, no file name', [FX.PRIVATE_STRINGS.filter((v) => all.indexOf(v) !== -1), / [AP]M|upn-|tenant\.|Presenter|\.csv/.test(all), FX.PARTICIPANTS.filter((p) => p.duration && all.indexOf(p.duration) !== -1).length], [[], false, 0]);
  check('nothing of it is logged', x.logs.filter((l) => /Organizer|Rivera|Wong|example\./.test(l)), []);
}

// ================================================================ 6. companies

console.log('6. Company corrections across the days');
{
  const c = report();
  c.import('s2', '2026-10-26', DAY1);
  c.import('s2', '2026-10-27', DAY2);
  c.import('s2', '2026-10-28', DAY3);
  const cell = (table, name, nth) => { const rows = []; const t = c.tables()[table]; for (let i = 1; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === name) rows.push(t.getRow(i).getCell(1)); return rows[nth || 0]; };
  const companies = (name) => c.tables().map((t) => rowsOf(t).filter((row) => row[0] === name).map((row) => row[1]).join('|'));

  cell(0, 'Alex Organizer').setText('ExampleCorp');
  c.s.refreshAdhocAttendanceSection();
  check('a Company typed for an attendee with an e-mail on one day is that attendee\'s Company on every day: the e-mail identifies the person', companies('Alex Organizer'), ['ExampleCorp', 'ExampleCorp', 'ExampleCorp']);
  c.s.refreshAdhocAttendanceSection();
  c.s.refreshAdhocAttendanceSection();
  check('it stays, refresh after refresh', companies('Alex Organizer'), ['ExampleCorp', 'ExampleCorp', 'ExampleCorp']);
  check('nobody else got a Company from it', c.tables().map((t) => rowsOf(t).slice(1).filter((row) => row[0] !== 'Alex Organizer' && row[1] !== '').length), [0, 0, 0]);

  cell(1, 'Sam Rivera').setText('RiverCorp');
  c.s.refreshAdhocAttendanceSection();
  check('an attendee without an e-mail is identified by the exact name, where that name stands for one person in the table: the Company carries over to the other day he is listed', companies('Sam Rivera'), ['RiverCorp', 'RiverCorp', '']);
  cell(1, 'Dana Twin', 0).setText('TwinCorp');
  c.s.refreshAdhocAttendanceSection();
  check('two attendees of one name without e-mail: a Company typed for one of them is not stored, and not given to the other', [companies('Dana Twin'), JSON.stringify(c.state().companies).indexOf('TwinCorp')], [['', '|', ''], -1]);
  check('what is stored: one correction per identified attendee, keyed by e-mail or by name', c.state().companies, { 'email:alex.organizer@example.com': 'ExampleCorp', 'name:sam rivera': 'RiverCorp' });

  cell(2, 'Alex Organizer').setText('OtherCorp');
  c.s.refreshAdhocAttendanceSection();
  check('a later correction on another day replaces it everywhere: one person, one Company', companies('Alex Organizer'), ['OtherCorp', 'OtherCorp', 'OtherCorp']);
  check('a correction survives the re-import of one day', (() => { c.import('s2', '2026-10-27', DAY2.slice(0, 3)); return [companies('Alex Organizer'), companies('Sam Rivera')]; })(), [['OtherCorp', 'OtherCorp', 'OtherCorp'], ['RiverCorp', 'RiverCorp', '']]);
  check('removing a day keeps the corrections of the attendees who are still listed', (() => { c.s.removeTeamsAttendanceImport('s2', '2026-10-26'); return [companies('Alex Organizer'), Object.keys(c.state().companies).sort()]; })(), [['OtherCorp', 'OtherCorp'], ['email:alex.organizer@example.com', 'name:sam rivera']]);

  // A table of the same shape that is not generated.
  const e = report();
  e.body.appendParagraph('Attendees (made by hand)');
  const old = e.body.appendTable([['Name', 'Company', 'Email'], ['Alex Organizer', 'HandMadeCorp', 'alex.organizer@example.com']]);
  const first = e.import('s2', '2026-10-26', DAY1);
  e.import('s2', '2026-10-27', DAY2);
  const gen = (name) => e.tables().map((t) => rowsOf(t).filter((row) => row[0] === name).map((row) => row[1]).join('|'));
  check('a Name / Company / Email table the report had before is read once, at the first import, and pointed out; it is never changed', [/another Name \/ Company \/ Email table/.test(first.message), gen('Alex Organizer'), rowsOf(old), e.body._children.indexOf(old) !== -1],
    [true, ['HandMadeCorp', 'HandMadeCorp'], [['Name', 'Company', 'Email'], ['Alex Organizer', 'HandMadeCorp', 'alex.organizer@example.com']], true]);
  check('after that only the generated tables are read: a correction typed there is not overridden by the other table', (() => {
    const t = e.tables()[1]; for (let i = 1; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === 'Alex Organizer') t.getRow(i).getCell(1).setText('TypedCorp');
    e.s.refreshAdhocAttendanceSection(); e.s.refreshAdhocAttendanceSection(); e.s.refreshAdhocAttendanceSection();
    return [gen('Alex Organizer'), rowsOf(old)[1][1], e.s.adhocAttendeeTablesToHarvest_(e.body).length];
  })(), [['TypedCorp', 'TypedCorp'], 'HandMadeCorp', 2]);
}

// ================================================================ 7. opening, status, statistics

console.log('7. opening details, status and statistics');
{
  const when = S.adhocSessionWhenText_;
  check('when a session of several days is: its days as a range, and the times said to be on the first and the last day', [when(SESSION2), when(session('s2', '2026-10-26', '2026-10-28', '', '', 'AHG Session 2')), when(session('s2', '2026-10-26', '2026-10-28', '09:00', '')), when(session('s2', '2026-10-26', '2026-10-28', '', '17:00'))],
    ['AHG Session 2, October 26–28, 2026 (from 09:00 on the first day, until 23:00 on the last day)', 'AHG Session 2, October 26–28, 2026', 'October 26–28, 2026 (from 09:00 on the first day)', 'October 26–28, 2026 (until 17:00 on the last day)']);
  check('it never reads as one day: neither "October 28, 2026" alone nor "09:00–23:00"', [/October 28, 2026/.test(when(SESSION2)), /09:00–23:00/.test(when(SESSION2))], [false, false]);
  check('across a month and across a year', [when(session('s1', '2026-10-30', '2026-11-01', '', '', 'A')), when(session('s1', '2026-12-31', '2027-01-02', '', '', 'B'))], ['A, October 30 – November 1, 2026', 'B, December 31, 2026 – January 2, 2027']);
  check('a session of one day reads as it always did', [when(CALL1), when(session('s1', '2026-09-22', '', '15:00')), when(session('s1', '2026-09-22', '', '', '18:00')), when(session('s1', '2026-09-22'))],
    ['AHG Call 1, September 22, 2026, 15:00–18:00', 'September 22, 2026, 15:00', 'September 22, 2026, until 18:00', 'September 22, 2026']);

  const o = report();
  const saved = o.s.saveAdhocOpeningDetails({ s1: { chair: 'Alex Organizer', minuteTakers: 'Sam Rivera', note: '' }, s2: { chair: 'Alex Organizer', minuteTakers: '', note: 'Three days.' } });
  const c = o.s.findAdhocOpeningContainer_(o.body);
  check('opening details are one record per logical session: the section has one entry for the three days', [saved.ok, o.body._children.slice(c.start, c.end).map((e) => e.getText()), Object.keys(JSON.parse(o.docProps.getProperty('ADHOC_SESSION_OPENING')).sessions)],
    [true, ['Session administration', 'A01: AHG Call 1, September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera', 'A02: AHG Session 2, October 26–28, 2026 (from 09:00 on the first day, until 23:00 on the last day)', 'Chair: Alex Organizer', 'Three days.'], ['s1', 's2']]);

  // Status.
  o.import('s1', '2026-09-22', DAY1);
  o.import('s2', '2026-10-26', DAY1);
  o.import('s2', '2026-10-28', DAY3);
  const model = o.s.buildAdhocSessionStatusModel_(o.s.getAdhocSessions_(), o.s.parseAdhocOpeningProperty_(o.docProps.getProperty('ADHOC_SESSION_OPENING')), o.state(),
    { available: false, reason: 'not asked for', column: 4, clock: 'utc', storedStatus: 'absent', storedError: null, summary: null });
  const overview = {}; model.overview.forEach((row) => { overview[row[0]] = row[1]; });
  check('the status: the days of all sessions as one range, sessions with attendance, and how many attendance records there are', [overview.Sessions, overview['Attendance imported'], model.attendance], ['2 (September 22 – October 28, 2026)', '2 of 2 sessions (3 attendance records)', { known: true, imported: 2, missing: 0, damaged: 0, stale: 0, records: 3 }]);
  check('a session of several days with fewer records than days: a remark, not a warning', [model.notes.filter((n) => /imported for/.test(n.text)), model.notes.filter((n) => n.level === 'warning' && /ttendance/.test(n.text))], [[{ level: 'info', text: 'Attendance of AHG Session 2 is imported for 2 of its 3 days.' }], []]);
  const groups = (id) => { const out = {}; model.sessions.filter((x) => x.id === id)[0].groups.forEach((g) => { out[g.title] = g.rows; }); return out; };
  check('the statistics of the session of several days: its days and times, and per record the day, the values Teams states and the number of attendees', [model.sessions[1].when, groups('s2').Planned, groups('s2')['Attendance (Microsoft Teams)'].map((row) => row.join(': '))],
    ['AHG Session 2, October 26–28, 2026 (from 09:00 on the first day, until 23:00 on the last day)', [['Dates', 'October 26–28, 2026'], ['Start (first day)', '09:00'], ['End (last day)', '23:00']],
      ['Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Date: October 26, 2026', 'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 3', 'Average attendance: 2:31:07', 'Attendees listed: 3',
        'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Date: October 28, 2026', 'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 2', 'Average attendance: 2:31:07', 'Attendees listed: 2']]);
  check('the session of one day: as before', [groups('s1').Planned, groups('s1')['Attendance (Microsoft Teams)'].map((row) => row[0])], [[['Date', 'September 22, 2026'], ['Start', '15:00'], ['End', '18:00']], ['Meeting', 'Start', 'End', 'Duration', 'Attendance records', 'Average attendance', 'Attendees listed']]);
  check('status and statistics hold counts only: no attendee, no Chair, no note', /Organizer|Rivera|Wong|Omar|example\.(com|org)|Three days/.test(JSON.stringify(model) + o.s.formatAdhocSessionStatusLines_(model).join('\n')), false);
  check('all records imported: no remark about days', (() => { o.import('s2', '2026-10-27', DAY2); const m = o.s.buildAdhocSessionStatusModel_(o.s.getAdhocSessions_(), o.s.parseAdhocOpeningProperty_(null), o.state(), { available: false, reason: 'x', column: 4, clock: 'utc', storedStatus: 'absent', storedError: null, summary: null });
    return [m.overview[2][1], m.notes.filter((n) => /imported for/.test(n.text)).length]; })(), ['2 of 2 sessions (4 attendance records)', 0]);
  check('the dates of the sessions as one range', [S.adhocSessionDateRangeText_([CALL1, SESSION2]), S.adhocSessionDateRangeText_([SESSION2]), S.adhocSessionDateRangeText_([SESSION2, CALL1]), S.adhocSessionDateRangeText_([session('s1', '2026-12-30', '2027-01-02'), session('s2', '2026-12-01')])],
    ['September 22 – October 28, 2026', 'October 26–28, 2026', 'September 22 – October 28, 2026', 'December 1, 2026 – January 2, 2027']);
}

// ================================================================ range integrity

console.log('   the days of a session keep its imported attendance inside');
{
  const rows = (s2, s1) => [Object.assign({}, s1 || CALL1), Object.assign({}, s2)];
  const s2 = (date, endDate, extra) => Object.assign(session('s2', date, endDate, '09:00', '23:00', 'AHG Session 2'), extra || {});
  const WAY_OUT = 'Remove that attendance first (Sessions and Attendance > Import Teams Attendance… > Remove attendance…) or keep ';
  const state = (x) => JSON.stringify([x.docProps.getProperty(KEY), x.attendanceKeys().map((k) => x.docProps._store[k])]);
  const three = () => { const x = report(); x.import('s2', '2026-10-26', DAY1); x.import('s2', '2026-10-27', DAY2); x.import('s2', '2026-10-28', DAY3); x.writes.length = 0; return x; };

  const a = three();
  const before = state(a);
  const body = a.snapshot();
  const first = a.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-27', '2026-10-28')));
  check('attendance for 26, 27 and 28 October; 26–28 changed to 27–28: refused, naming the day that would be left outside and the way out', first,
    { ok: false, errors: ['AHG Session 2 cannot be changed to October 27–28, 2026 because attendance for October 26, 2026 is already imported. ' + WAY_OUT + 'that day within the session.'], changed: false, sessions: [] });
  check('the same records; 26–28 changed to 26–27: refused because the 28th has attendance', a.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', '2026-10-27'))).errors,
    ['AHG Session 2 cannot be changed to October 26–27, 2026 because attendance for October 28, 2026 is already imported. ' + WAY_OUT + 'that day within the session.']);
  check('several days outside: all of them, in order, in one message', [a.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-27', ''))).errors, a.s.saveAdhocSessionsConfiguration(rows(s2('2026-11-02', '2026-11-04'))).errors],
    [['AHG Session 2 cannot be changed to October 27, 2026 because attendance for October 26, 2026, October 28, 2026 is already imported. ' + WAY_OUT + 'those days within the session.'],
      ['AHG Session 2 cannot be changed to November 2–4, 2026 because attendance for October 26, 2026, October 27, 2026, October 28, 2026 is already imported. ' + WAY_OUT + 'those days within the session.']]);
  check('a refused save changes nothing: not the sessions, not the attendance, not the report; nothing is written, removed or moved, and the session is not extended', [state(a) === before, a.snapshot() === body, a.writes, a.records('s2').map((x) => x[1]), a.s.getAdhocSessions_()[1].date + '..' + a.s.getAdhocSessions_()[1].endDate],
    [true, true, [], ['2026-10-26', '2026-10-27', '2026-10-28'], '2026-10-26..2026-10-28']);
  check('the message uses the label the row has now, and the ids are untouched', [a.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-27', '2026-10-28', { label: 'Renamed' }))).errors[0].slice(0, 49), a.s.getAdhocSessions_().map((x) => x.id + ' ' + x.displayLabel)],
    ['Renamed cannot be changed to October 27–28, 2026 ', ['s1 AHG Call 1', 's2 AHG Session 2']]);
  check('two sessions that would each leave attendance outside: one message per session, in the order of the sessions', (() => {
    const x = three(); x.import('s1', '2026-09-22', DAY1);
    return x.s.saveAdhocSessionsConfiguration([Object.assign({}, CALL1, { date: '2026-09-23' }), s2('2026-10-27', '2026-10-28')]).errors.map((e) => e.split(' because ')[0]);
  })(), ['AHG Call 1 cannot be changed to September 23, 2026', 'AHG Session 2 cannot be changed to October 27–28, 2026']);

  // What stays allowed.
  const only27 = report(); only27.import('s2', '2026-10-27', DAY2);
  check('attendance for the 27th only; 26–28 changed to the 27th alone: saved', (() => { const x = only27.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-27', ''))); return [x.ok, x.changed, only27.s.getAdhocSessions_()[1].date, only27.s.getAdhocSessions_()[1].endDate, only27.records('s2')]; })(),
    [true, true, '2026-10-27', undefined, [['s2d20261027', '2026-10-27', 5]]]);
  const two = report(); two.import('s2', '2026-10-26', DAY1); two.import('s2', '2026-10-27', DAY2);
  check('attendance for the 26th and 27th; 26–28 changed to 25–29: saved (a range may grow)', (() => { const x = two.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-25', '2026-10-29'))); return [x.ok, x.changed, two.records('s2').map((y) => y[1]), two.lines().filter((t) => /^Session: AHG Session 2/.test(t))[0]]; })(),
    [true, true, ['2026-10-26', '2026-10-27'], 'Session: AHG Session 2, October 25–29, 2026']);
  check('shrinking to exactly the days that have attendance: saved', two.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', '2026-10-27'))).ok, true);
  const b = three();
  check('a label-only edit: saved', (() => { const x = b.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', '2026-10-28', { label: 'AHG 2' }))); return [x.ok, x.changed, b.s.getAdhocSessions_()[1].displayLabel]; })(), [true, true, 'AHG 2']);
  check('a time-only edit: saved', (() => { const x = b.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', '2026-10-28', { label: 'AHG 2', start: '10:00', end: '18:00' }))); return [x.ok, x.changed, b.s.getAdhocSessions_()[1].end, b.records('s2').length]; })(), [true, true, '18:00', 3]);
  check('saving unchanged sessions: nothing written', (() => { b.writes.length = 0; const x = b.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', '2026-10-28', { label: 'AHG 2', start: '10:00', end: '18:00' }))); return [x.ok, x.changed, b.writes.filter((w) => /ADHOC_SESSIONS$/.test(w))]; })(), [true, false, []]);
  check('another session can be changed and added freely while this one keeps its days', (() => { const x = b.s.saveAdhocSessionsConfiguration([Object.assign({}, CALL1, { date: '2026-09-21', endDate: '2026-09-23' }), s2('2026-10-26', '2026-10-28', { label: 'AHG 2', start: '10:00', end: '18:00' }), { id: '', label: 'Third', date: '2026-11-10', start: '', end: '' }]); return [x.ok, b.s.getAdhocSessions_().map((y) => y.id)]; })(), [true, ['s1', 's2', 's10']]);

  // A record stored under the plain session id.
  const legacy = report();
  const payload = (date, people) => legacy.s.compactAdhocAttendance_(legacy.s.parseTeamsAttendanceReport_(exportFor(date, people), { expectedDate: date }).attendance);
  legacy.s.writeAdhocAttendanceValue_(legacy.docProps, 's2', payload('2026-10-27', DAY2));
  legacy.s.writeAdhocAttendanceValue_(legacy.docProps, 's1', payload('2026-09-22', DAY1));
  check('a record stored under the session id is checked by the day its own Teams start time says', [legacy.s.adhocAttendanceDaysBySession_(legacy.state()), legacy.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-28', ''))).errors.map((e) => e.split('. ')[0]),
    legacy.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', '2026-10-28'), Object.assign({}, CALL1, { date: '2026-09-23' }))).errors.map((e) => e.split('. ')[0]), legacy.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-27', ''))).ok],
    [{ s2: ['2026-10-27'], s1: ['2026-09-22'] }, ['AHG Session 2 cannot be changed to October 28, 2026 because attendance for October 27, 2026 is already imported'], ['AHG Call 1 cannot be changed to September 23, 2026 because attendance for September 22, 2026 is already imported'], true]);
  check('a session of one day with attendance can be given more days around that day', (() => { const x = legacy.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-27', ''), Object.assign({}, CALL1, { date: '2026-09-21', endDate: '2026-09-23' }))); return [x.ok, legacy.records('s1')]; })(), [true, [['s1', '2026-09-22', 3]]]);

  // No attendance.
  const none = report();
  check('a session without attendance is edited as before: any dates', [none.s.saveAdhocSessionsConfiguration(rows(s2('2026-11-02', '2026-11-04'))).ok, none.s.saveAdhocSessionsConfiguration(rows(s2('2026-12-01', ''))).ok, none.s.getAdhocSessions_()[1].date], [true, true, '2026-12-01']);
  check('opening the dialog and Cancel still write nothing', (() => { const x = three(); const s0 = state(x); x.s.configureAdhocSessions(); const d = runSessionsDialog(x.ui.dialogs[0].html, x.s); d.inputs(1)[1].value = '2026-10-27'; d.ctx.google.script.host.close(); return [state(x) === s0, x.writes, d.calls.length]; })(), [true, [], 0]);
  check('in the dialog a refused save shows the message and keeps the dialog open', (() => { const x = three(); x.s.configureAdhocSessions(); const d = runSessionsDialog(x.ui.dialogs[0].html, x.s); d.inputs(1)[1].value = '2026-10-27'; d.ctx.saveSessions(); return [d.byId.errors.textContent.slice(0, 61), d.closed(), x.writes]; })(),
    ['AHG Session 2 cannot be changed to October 27–28, 2026 becaus', 0, []]);

  // The pure check, and what cannot be told.
  const check_ = S.adhocSessionsLeavingAttendanceOutside_;
  check('the check is inclusive at both ends, and not made when the days are not given', [check_([SESSION2], [SESSION2], { s2: ['2026-10-26', '2026-10-28'] }), check_([SESSION2], [s2('2026-10-27', '2026-10-28')], undefined), check_([SESSION2], [s2('2026-10-27', '2026-10-28')], {}), check_([SESSION2], [s2('2026-10-27', '2026-10-28')], { s9: ['2026-10-26'] })], [[], [], [], []]);
  check('attendance of a newer release (its days cannot be told): the days of an existing session are not changed; labels, times and new sessions are', [check_([CALL1, SESSION2], [CALL1, s2('2026-10-27', '2026-10-28')], null),
    check_([CALL1, SESSION2], [Object.assign({}, CALL1, { label: 'x', end: '19:00' }), SESSION2, session('s3', '2026-11-10')], null)],
    [['The stored attendance was written by a newer release, so it cannot be told which days have attendance. The dates of AHG Session 2 were not changed.'], []]);
  check('that holds for a changed last day alone, and for a session of one day given a last day', [check_([SESSION2], [s2('2026-10-26', '2026-10-27')], null).length, check_([CALL1], [Object.assign({}, CALL1, { endDate: '2026-09-23' })], null).length, check_([SESSION2], [s2('2026-10-26', '2026-10-28')], null).length], [1, 1, 0]);
  check('through the dialog\'s save with such attendance stored: refused, nothing written', (() => { const x = report(); x.docProps._store[ATT] = '{"v":2}'; const s0 = x.docProps.getProperty(KEY); const y = x.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-27', '2026-10-28'))); return [y.ok, /newer release/.test(y.errors[0]), x.docProps.getProperty(KEY) === s0, x.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', '2026-10-28', { label: 'Other' }))).ok]; })(), [false, true, true, true]);
  check('a damaged record whose key names its day is still counted for that day; one whose day cannot be told is not', (() => {
    const x = report(); x.import('s2', '2026-10-26', DAY1); x.import('s2', '2026-10-27', DAY2);
    const index = JSON.parse(x.docProps.getProperty(ATT));
    x.docProps._store[ATT + '_s2d20261027_' + index.entries.s2d20261027.g + '_0'] = '{"t":"tampered"}';
    const y = report(); y.s.writeAdhocAttendanceValue_(y.docProps, 's2', '{"not":"attendance"}');
    return [x.s.adhocAttendanceDaysBySession_(x.state()), x.s.saveAdhocSessionsConfiguration(rows(s2('2026-10-26', ''))).ok, y.s.adhocAttendanceDaysBySession_(y.state()), y.s.saveAdhocSessionsConfiguration(rows(s2('2026-11-02', ''))).ok];
  })(), [{ s2: ['2026-10-26', '2026-10-27'] }, false, { s2: [] }, true]);
  check('the delete guard is what it was: a session with attendance is not removed, with its own message', three().s.saveAdhocSessionsConfiguration([Object.assign({}, CALL1)]).errors,
    ['AHG Session 2 has imported attendance and was not removed. Remove its attendance first: Sessions and Attendance > Import Teams Attendance… > Remove attendance….']);
  check('the check reads and writes nothing itself, and removes or moves no attendance (source)', (() => {
    const at = CODE.indexOf('\nfunction adhocSessionsLeavingAttendanceOutside_('); const src = CODE.slice(at, CODE.indexOf('\nfunction saveAdhocSessionsConfiguration('));
    return (src.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '').match(/setProperty|deleteProperty|writeAdhocAttendanceValue_|getProperty|PropertiesService|DocumentApp|Logger\.log/g) || []);
  })(), []);
}

// ================================================================ 8. a session of one day

console.log('8. a session of one day is exactly what it was');
{
  const one = report(Object.assign({}, ADHOC, { [KEY]: sessionsProperty([session('s1', '2026-09-22', '', '15:00', '18:00')], 2) }));
  check('its stored form has no endDate, and saving it again writes nothing', [one.docProps.getProperty(KEY), one.s.saveAdhocSessionsConfiguration([{ id: 's1', label: '', date: '2026-09-22', start: '15:00', end: '18:00' }]).changed, one.s.saveAdhocSessionsConfiguration([{ id: 's1', label: '', date: '2026-09-22', endDate: '', start: '15:00', end: '18:00' }]).changed, one.writes],
    ['{"v":1,"nextId":2,"sessions":[{"id":"s1","label":"","date":"2026-09-22","start":"15:00","end":"18:00"}]}', false, false, []]);
  const x = one.s.confirmTeamsAttendanceImport('s1', base64(FX.buildExport()), one.s.previewTeamsAttendanceImport('s1', base64(FX.buildExport())).token);
  check('its attendance: one record under the session id, the block without a Date line, the messages without a day', [x.message, Object.keys(JSON.parse(one.docProps.getProperty(ATT)).entries), one.lines().slice(0, 5), one.lines().filter((t) => /^Date:/.test(t)), one.s.removeTeamsAttendanceImport('s1').message],
    ['The attendance of 22 Sep was imported (23 attendees).', ['s1'], ['Attendance', 'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: September 22, 2026', 'Start: 14:48:05'], [], 'The attendance of 22 Sep was removed.']);
  check('the export of another day is refused with the message it always had', one.preview('s1', '2026-09-23').error, 'This attendance export is for 2026-09-23, not for the session on 2026-09-22.');
  check('nothing here moves, renames or removes a heading of the closing section (source): the Attendance section is written, nothing else', (() => {
    const at = CODE.indexOf('\nfunction renderAdhocAttendanceSection_('); const src = CODE.slice(at, CODE.indexOf('\n}\n', at));
    return [/5\.11|Close of|HEADING[124]|setText\(/.test(src.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '')), (src.match(/removeAdhocBodyChild_\(body, body\.getChild\(i\)\)/g) || []).length];
  })(), [false, 2]);
  check('Code.js is version 2.21.1 (2.18.0 added the ad-hoc sessions, 2.18.1 the TDoc upload completion, 2.19.0 the status dropdowns, 2.20.0 their conversion in an existing report, 2.21.0 the ad-hoc meeting automation and the personal Reviewer token, 2.21.1 the dropdown status in a discussion e-mail)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.21.1');
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll multi-day session checks passed.');
process.exitCode = failures ? 1 : 0;
