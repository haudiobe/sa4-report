/**
 * Ad-hoc sessions, stage D -- importing a Teams attendance export for a
 * session, storing it, and rendering the Attendance section
 * (docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md §4, §12, §18).
 *
 *   1. Storage: chunks, generations, integrity, interrupted writes, size.
 *   2. Preview: reads only, shows what Teams states, refuses the wrong file.
 *   3. Import: after confirmation only; idempotent; replacement.
 *   4. Rendering: placement, statistics, the attendee table.
 *   5. Manual Company values.
 *   6. Sessions that change; sessions with attendance are not removed.
 *   7. Removing an import.
 *   8. Failure and recovery.
 *   9. The dialog.
 *  10. Menu and activation.
 *  11. Build from Scratch.
 *  12. Adoption.
 *  13. Privacy, and nothing else changed.
 *
 * All data is synthetic (tests/fixtures/teams-attendance-synthetic.js).
 *
 * Run: node tests/adhoc-attendance-import.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
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
const KEY = 'ADHOC_SESSION_ATTENDANCE';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.4', flavor: 'template', codeVersion: '0.0.0', gitCommit: 'abcdef0', gitTag: 'template-release/T-2026.10.4', templateDocumentId: TEMPLATE_ID };

// ------------------------------------------------------------------ data

const session = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const sessionsProperty = (list, nextId) => JSON.stringify({ v: 1, nextId: nextId || list.length + 1, sessions: list });
const THREE_SESSIONS = [session('s1', '2026-09-22', '15:00', '18:00'), session('s2', '2026-09-23', '', '', 'Offline'), session('s3', '2026-09-24')];
const ADHOC = { MEETING_TYPE: 'adhoc', MEETING_ID: '86178', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'September 22, 2026', REPORT_SUFFIX: '6G' };
const WITH_SESSIONS = Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty(THREE_SESSIONS) });

/** The synthetic export for another day, optionally with other participants / summary. */
const exportFor = (date, options) => {
  const [, m, d] = date.match(/^\d{4}-(\d{2})-(\d{2})$/);
  return FX.buildExport(options).split(FX.DAY).join(parseInt(m, 10) + '/' + parseInt(d, 10) + '/26');
};
const person = (name, email, extra) => Object.assign({ name: name, join: '3:00:00 PM', leave: '4:00:00 PM', duration: '1h', email: email || '', upn: email ? 'upn-' + email.replace('@', '@tenant.') : '', role: 'Presenter' }, extra || {});
const SMALL = [person('Alex Organizer', 'alex.organizer@example.com'), person('[ExampleCorp] Sam Rivera (Unverified)'), person('Casey Guest (Unverified)'), person('Taro Yamada (山田 太郎) (External)', 'taro.yamada@example.net')];
const smallExport = (date, participants, summary) => exportFor(date, { participants: participants || SMALL, summary: Object.assign({ attendedParticipants: String((participants || SMALL).length) }, summary || {}) });
const base64 = (text) => FX.utf16le(text).toString('base64');
const FULL_22 = base64(FX.buildExport());

// ------------------------------------------------------------------ a report

/**
 * One report document in a sandbox. opts: { props, template (true = template
 * runtime), docId, closing (false = no closing heading), after (a heading
 * after the closing section) }.
 */
function report(opts) {
  const o = opts || {};
  const loaded = o.template ? loadTemplateRuntime({ release: RELEASE, documentProperties: o.props || WITH_SESSIONS }) : loadCode({ documentProperties: o.props || WITH_SESSIONS });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const H2 = s.DocumentApp.ParagraphHeading.HEADING2;
  body.appendParagraph('5.1 Opening of the session').setHeading(H2);
  body.appendParagraph('The chair opens the session.');
  if (o.closing !== false) {
    body.appendParagraph('5.11 Close of the session').setHeading(H2);
    body.appendParagraph('The chair thanks the participants and the minute takers.');
    body.appendParagraph('The session is closed at 18:11 CEST.');
  }
  if (o.after) body.appendParagraph('Annex A Agenda').setHeading(H2);

  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' } };
  ui.alert = (...args) => { ui.alerts.push(args); return 'OK'; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title: title, html: out.html }); };
  const logs = [];
  s.Logger = { log: (m) => logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => o.docId || REPORT_ID, getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.Utilities.base64Decode = (b64) => Array.from(Buffer.from(b64, 'base64')).map((b) => (b > 127 ? b - 256 : b));
  const lock = { busy: false, tries: 0, releases: 0 };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => { lock.tries++; return !lock.busy; }, releaseLock: () => { lock.releases++; } }) };

  const writes = [];
  const set = loaded.docProps.setProperty;
  const del = loaded.docProps.deleteProperty;
  loaded.docProps.setProperty = (k, v) => { writes.push('set ' + k); set(k, v); };
  loaded.docProps.deleteProperty = (k) => { writes.push('delete ' + k); del(k); };

  const r = { s: s, body: body, ui: ui, logs: logs, lock: lock, docProps: loaded.docProps, scriptProps: loaded.scriptProps, writes: writes };
  r.snapshot = () => JSON.stringify(elements(body));
  r.propsSnapshot = () => JSON.stringify(loaded.docProps._store);
  /** Preview and confirm, as the dialog does. */
  r.import = (sessionId, b64) => {
    const preview = s.previewTeamsAttendanceImport(sessionId, b64);
    if (!preview.ok) return preview;
    return s.confirmTeamsAttendanceImport(sessionId, b64, preview.token);
  };
  return r;
}

const isTable = (c) => c.getType() === 'TABLE';
const tableRows = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));
const elements = (body) => body._children.map((c) => (isTable(c) ? ['T', tableRows(c)] : ['P', c.getText(), c.getHeading(), !!c._bold]));
const texts = (body) => body._children.map((c) => (isTable(c) ? '[table ' + (c.getNumRows() - 1) + ']' : c.getText()));
const attendeeTables = (body) => body._children.filter((c) => isTable(c) && c.getRow(0).getCell(0).getText() === 'Name');
const container = (r) => { const c = r.s.findAdhocAttendanceContainer_(r.body); return c ? r.body._children.slice(c.start, c.end) : null; };
const containerTexts = (r) => { const c = container(r); return c ? c.map((e) => (isTable(e) ? '[table ' + (e.getNumRows() - 1) + ']' : e.getText())) : null; };
const attendanceKeys = (r) => r.docProps.getKeys().filter((k) => k.indexOf(KEY) === 0).sort();
/** Types `company` into the Company cell of the row whose name is `name`, in attendee table `tableIndex`. */
function typeCompany(r, tableIndex, name, company) {
  const t = attendeeTables(r.body)[tableIndex];
  for (let i = 1; i < t.getNumRows(); i++) {
    if (t.getRow(i).getCell(0).getText() === name) { t.getRow(i).getCell(1).setText(company); return true; }
  }
  return false;
}
const companyOf = (r, tableIndex, name) => { const row = tableRows(attendeeTables(r.body)[tableIndex]).filter((x) => x[0] === name)[0]; return row ? row[1] : null; };

// ================================================================ 1. storage

console.log('1. storage');
{
  const S = loadCode().sandbox;
  const store = () => { const l = loadCode(); return { s: l.sandbox, p: l.docProps }; };
  const keys = (p) => p.getKeys().sort();

  check('hash: stable, 8 hex digits, different for different text', [S.adhocTextHash_(''), S.adhocTextHash_('a'), S.adhocTextHash_('abc') === S.adhocTextHash_('abc'), S.adhocTextHash_('abc') === S.adhocTextHash_('abd'), /^[0-9a-f]{8}$/.test(S.adhocTextHash_('山田 😀'))],
    ['811c9dc5', 'e40c292c', true, false, true]);
  check('chunks: at most the given size, nothing lost', [S.splitAdhocAttendanceChunks_('abcdefgh', 3), S.splitAdhocAttendanceChunks_('', 3), S.splitAdhocAttendanceChunks_('abc', 3)], [['abc', 'def', 'gh'], [], ['abc']]);
  const emoji = 'ab😀cd😀';
  check('chunks: a character of two code units is never cut in half', [S.splitAdhocAttendanceChunks_(emoji, 3), S.splitAdhocAttendanceChunks_(emoji, 3).join('') === emoji], [['ab', '😀c', 'd😀'], true]);

  const a = store();
  check('nothing stored: absent, nothing to read', [a.s.readAdhocAttendanceIndex_(a.p).status, a.s.readAdhocAttendance_(a.p), a.s.adhocAttendanceSessionIds_(a.p)],
    ['absent', { status: 'absent', error: null, sessions: {}, raw: {}, problems: [], companies: {} }, []]);

  const model = S.parseTeamsAttendanceReport_(FX.buildExport(), { expectedDate: '2026-09-22' }).attendance;
  const small = S.compactAdhocAttendance_(model);
  // 70 attendees with long names and addresses: more than one chunk.
  const payload = S.compactAdhocAttendance_(S.parseTeamsAttendanceReport_(smallExport('2026-09-22', Array.from({ length: 70 }, (_, k) => person('Delegate Number ' + k + ' With A Long Name', 'delegate' + k + '.with.a.long.address@department.example.org'))), { expectedDate: '2026-09-22' }).attendance);
  check('the fixture\'s attendance is one chunk; the larger one used below is three', [Math.ceil(small.length / 2500), Math.ceil(payload.length / 2500)], [1, 3]);
  check('the stored form: summary values and name / company / e-mail, nothing else', [Object.keys(JSON.parse(small)), JSON.parse(small).p[2], JSON.parse(small).p.every((row) => row.length === 3)],
    [['t', 'n', 's', 'e', 'd', 'a', 'p'], ['Sam Rivera', 'ExampleCorp', ''], true]);
  check('it is read back as what the report needs', [Object.keys(S.expandAdhocAttendance_(small)), S.expandAdhocAttendance_(small).attendees.length, S.expandAdhocAttendance_(small).attendees[0]],
    [['meetingTitle', 'attendanceRecords', 'start', 'end', 'durationSeconds', 'averageAttendanceSeconds', 'attendees'], 23, { name: 'Alex Organizer', company: '', email: 'alex.organizer@example.com' }]);
  check('not that form: null', ['', 'x', '{}', '[]', '{"t":"x"}', JSON.stringify({ t: 'x', n: 1, s: 'soon', e: null, d: 1, a: 1, p: [] }), JSON.stringify({ t: 'x', n: 1, s: '2026-09-22T15:00:00', e: null, d: 1, a: 1, p: [['a', 'b']] }),
    JSON.stringify({ t: 'x', n: '1', s: '2026-09-22T15:00:00', e: null, d: 1, a: 1, p: [] })].map(S.expandAdhocAttendance_), [null, null, null, null, null, null, null, null]);

  check('first write', [a.s.writeAdhocAttendanceValue_(a.p, 's1', payload), keys(a.p), JSON.parse(a.p.getProperty(KEY))],
    [{ changed: true }, [KEY, KEY + '_s1_1_0', KEY + '_s1_1_1', KEY + '_s1_1_2'], { v: 1, gen: 1, entries: { s1: { g: 1, n: 3, len: payload.length, hash: S.adhocTextHash_(payload) } } }]);
  check('it is read back unchanged', [a.s.readAdhocAttendanceValue_(a.p, a.s.readAdhocAttendanceIndex_(a.p).index, 's1').text === payload, Object.keys(a.s.readAdhocAttendance_(a.p).sessions), a.s.adhocAttendanceSessionIds_(a.p)], [true, ['s1'], ['s1']]);
  const before = JSON.stringify(a.p._store);
  check('writing the same value again writes nothing', [a.s.writeAdhocAttendanceValue_(a.p, 's1', payload), JSON.stringify(a.p._store) === before], [{ changed: false }, true]);

  const other = S.compactAdhocAttendance_(S.parseTeamsAttendanceReport_(smallExport('2026-09-23'), { expectedDate: '2026-09-23' }).attendance);
  a.s.writeAdhocAttendanceValue_(a.p, 's2', other);
  check('a second session: its own chunks, the first untouched', [keys(a.p), a.s.readAdhocAttendance_(a.p).raw.s1 === payload, a.s.readAdhocAttendance_(a.p).raw.s2 === other],
    [[KEY, KEY + '_s1_1_0', KEY + '_s1_1_1', KEY + '_s1_1_2', KEY + '_s2_2_0'], true, true]);
  a.s.writeAdhocAttendanceValue_(a.p, 's1', other);
  check('replacing a value: a new generation, the old chunks gone, the other session untouched', [keys(a.p), JSON.parse(a.p.getProperty(KEY)).entries.s1.g, a.s.readAdhocAttendance_(a.p).raw],
    [[KEY, KEY + '_s1_3_0', KEY + '_s2_2_0'], 3, { s1: other, s2: other }]);
  check('removing one value', [a.s.writeAdhocAttendanceValue_(a.p, 's1', null), keys(a.p), a.s.adhocAttendanceSessionIds_(a.p)], [{ changed: true }, [KEY, KEY + '_s2_2_0'], ['s2']]);
  check('removing what is not there writes nothing', a.s.writeAdhocAttendanceValue_(a.p, 's1', null), { changed: false });
  check('removing the last value removes the index and every chunk', [a.s.writeAdhocAttendanceValue_(a.p, 's2', null), keys(a.p)], [{ changed: true }, []]);

  // Integrity.
  const b = store();
  b.s.writeAdhocAttendanceValue_(b.p, 's1', payload);
  b.s.writeAdhocAttendanceValue_(b.p, 's2', other);
  b.p.deleteProperty(KEY + '_s1_1_1');
  check('a missing chunk: that session is a problem, the other is read', [b.s.readAdhocAttendance_(b.p).problems, Object.keys(b.s.readAdhocAttendance_(b.p).sessions), b.s.adhocAttendanceSessionIds_(b.p)], [['s1'], ['s2'], ['s1', 's2']]);
  const c = store();
  c.s.writeAdhocAttendanceValue_(c.p, 's1', payload);
  c.p.setProperty(KEY + '_s1_1_0', c.p.getProperty(KEY + '_s1_1_0').replace('Delegate Number 0', 'Delegate Numbex 0'));
  check('a changed chunk (same length): a problem, never used', [c.s.readAdhocAttendance_(c.p).problems, c.s.readAdhocAttendance_(c.p).sessions], [['s1'], {}]);
  const d = store();
  d.s.writeAdhocAttendanceValue_(d.p, 's1', 'not attendance at all');
  check('a value that is intact but is not attendance: a problem', d.s.readAdhocAttendance_(d.p).problems, ['s1']);
  check('a problem value is replaced by a new import', [d.s.writeAdhocAttendanceValue_(d.p, 's1', other), d.s.readAdhocAttendance_(d.p).problems], [{ changed: true }, []]);

  ['{"v":1,', '[]', '"x"', '{"v":1}', '{"v":1,"gen":1}', '{"v":1,"gen":"1","entries":{}}', '{"v":1,"gen":1,"entries":{"bad id":{"g":1,"n":1,"len":1,"hash":"x"}}}',
    '{"v":1,"gen":1,"entries":{"s1":{"g":2,"n":1,"len":1,"hash":"x"}}}', '{"v":1,"gen":1,"entries":{"s1":{"g":1,"n":"1","len":1,"hash":"x"}}}'].forEach((raw) => {
    const e = store();
    e.p.setProperty(KEY, raw);
    check(`an index that cannot be read (${raw.slice(0, 34)}): invalid, nothing used, no error thrown`, [e.s.readAdhocAttendanceIndex_(e.p).status, e.s.readAdhocAttendance_(e.p).sessions, e.s.adhocAttendanceSessionIds_(e.p)], ['invalid', {}, []]);
  });
  const f = store();
  f.p.setProperty(KEY, '{"v":1,');
  f.p.setProperty(KEY + '_s1_7_0', 'left over');
  check('a new import replaces an unreadable index, with a generation above every chunk still stored',
    [f.s.writeAdhocAttendanceValue_(f.p, 's2', other), keys(f.p), f.s.readAdhocAttendance_(f.p).raw], [{ changed: true }, [KEY, KEY + '_s2_8_0'], { s2: other }]);

  const g = store();
  g.p.setProperty(KEY, JSON.stringify({ v: 2, gen: 1, entries: {} }));
  check('a newer schema version: reported, nothing used, which sessions have attendance is unknown', [g.s.readAdhocAttendanceIndex_(g.p).status, g.s.readAdhocAttendance_(g.p).status, g.s.adhocAttendanceSessionIds_(g.p)], ['unsupported', 'unsupported', null]);
  check('and it is never overwritten', [/newer release/.test(thrown(() => g.s.writeAdhocAttendanceValue_(g.p, 's1', other)) || ''), g.p.getProperty(KEY)], [true, '{"v":2,"gen":1,"entries":{}}']);

  // Interrupted writes: a failure at any step leaves the old value valid.
  const steps = (() => { const h = store(); h.s.writeAdhocAttendanceValue_(h.p, 's1', other); let n = 0; const set = h.p.setProperty; h.p.setProperty = (k, v) => { n++; set(k, v); }; h.s.writeAdhocAttendanceValue_(h.p, 's1', payload); return n; })();
  check('replacing by a three-chunk value takes four property writes (three chunks, then the index)', steps, 4);
  for (let failAt = 1; failAt <= steps; failAt++) {
    const h = store();
    h.s.writeAdhocAttendanceValue_(h.p, 's1', other);
    let n = 0;
    const set = h.p.setProperty;
    h.p.setProperty = (k, v) => { n++; if (n === failAt) throw new Error('quota'); set(k, v); };
    const message = thrown(() => h.s.writeAdhocAttendanceValue_(h.p, 's1', payload));
    h.p.setProperty = set;
    const after = h.s.readAdhocAttendance_(h.p);
    check(`a write that fails at step ${failAt} of ${steps}: the error is raised, the old value is still the stored one`, [message, after.problems, after.raw.s1 === other], ['quota', [], true]);
    h.s.writeAdhocAttendanceValue_(h.p, 's1', payload);
    check(`  and the next write completes and leaves no chunk behind`, [h.s.readAdhocAttendance_(h.p).raw.s1 === payload, keys(h.p).length, keys(h.p).every((k) => k === KEY || /_s1_\d+_[012]$/.test(k))], [true, 4, true]);
  }
  const i = store();
  i.s.writeAdhocAttendanceValue_(i.p, 's1', other);
  const del = i.p.deleteProperty;
  i.p.deleteProperty = () => { throw new Error('cannot delete'); };
  const delMessage = thrown(() => i.s.writeAdhocAttendanceValue_(i.p, 's1', payload));
  i.p.deleteProperty = del;
  check('a failure while deleting the old chunks: the NEW value is the stored one (the index was already written)', [delMessage, i.s.readAdhocAttendance_(i.p).raw.s1 === payload], ['cannot delete', true]);

  // Size.
  const bytes = (p) => p.getKeys().map((k) => Buffer.byteLength(p.getProperty(k), 'utf8'));
  const many = (n, name) => Array.from({ length: n }, (_, k) => person(name(k), 'delegate' + k + '.with.a.long.address@department.example.org'));
  [[70, (k) => 'Delegate Number ' + k + ' With A Long Name (External)'], [400, (k) => 'Delegate ' + k], [120, (k) => '山田 太郎 デリゲート 参加者 ' + k + ' （示例公司）'], [60, (k) => '😀😀😀😀😀 ' + k]].forEach(([n, name]) => {
    const j = store();
    const big = S.compactAdhocAttendance_(S.parseTeamsAttendanceReport_(smallExport('2026-09-22', many(n, name)), { expectedDate: '2026-09-22' }).attendance);
    j.s.writeAdhocAttendanceValue_(j.p, 's1', big);
    check(`${n} attendees ("${name(0).slice(0, 12)}…"): ${keys(j.p).length - 1} chunk(s), every property under 9 KB, read back unchanged`,
      [Math.max(...bytes(j.p)) < 9000, j.s.readAdhocAttendance_(j.p).raw.s1 === big, j.s.readAdhocAttendance_(j.p).sessions.s1.attendees.length], [true, true, n]);
  });
  check('a chunk has at most 2,500 characters, so at most 7.5 KB', (CODE.match(/const ADHOC_ATTENDANCE_CHUNK_CHARS_ = (\d+);/) || [])[1], '2500');
}

// ================================================================ 2. preview

console.log('2. preview');
{
  const r = report();
  const body0 = r.snapshot();
  const props0 = r.propsSnapshot();
  let parserArgs = null;
  const realParser = r.s.parseTeamsAttendanceReport_;
  r.s.parseTeamsAttendanceReport_ = (text, options) => { parserArgs = options; return realParser(text, options); };

  const p = r.s.previewTeamsAttendanceImport('s1', FULL_22);
  check('accepted', [p.ok, p.error, typeof p.token], [true, null, 'string']);
  check('the session date is what the parser is told to expect', parserArgs, { expectedDate: '2026-09-22' });
  check('what the preview shows', p.preview, {
    session: 'September 22, 2026',
    meetingTitle: 'Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE',
    date: 'September 22, 2026', start: '14:48:05', end: '18:11:26', duration: '3:23:20',
    attendanceRecords: 25, attendees: 23, averageAttendance: '2:31:07',
    warnings: [
      'Records with the same name were kept as separate attendees: nothing in the export shows that they are one person. (rows 12, 13)',
      'Records with the same name were kept as separate attendees: nothing in the export shows that they are one person. (rows 14, 15)',
      '1 e-mail value is not a usable address and was left empty. (row 18)'
    ],
    notes: [
      '2 records of one attendee were merged. (rows 11, 22)',
      '2 records of one attendee were merged. (rows 16, 17)',
      '4 records have no complete join and leave time; they are kept. (rows 19, 20, 24, 25)',
      '2 participant rows have fewer cells than the header.',
      'Teams states an average attendance of 2:31:07; the mean of the listed durations is 1:55:35. The value Teams states is used.',
      'Teams states a meeting duration of 3:23:20; end minus start is 3:23:21. The value Teams states is used.'
    ],
    replaces: '', otherDays: '', unchanged: false
  });
  check('a preview changes no property and nothing in the document', [r.propsSnapshot() === props0, r.snapshot() === body0, r.writes], [true, true, []]);
  check('a preview takes no lock', r.lock.tries, 0);
  check('the preview and its warnings name no participant', FX.PARTICIPANTS.filter((x) => JSON.stringify(p).indexOf(x.email || 'no-such-address') !== -1 || (/^[A-Z][a-z]+ [A-Z][a-z]+$/.test(x.name) && JSON.stringify(p).indexOf(x.name) !== -1)), []);
  check('a labelled session is shown with its label', r.s.previewTeamsAttendanceImport('s2', base64(smallExport('2026-09-23'))).preview.session, 'Offline, September 23, 2026');

  // The wrong file.
  check('the export of another day: refused, with both dates', r.s.previewTeamsAttendanceImport('s2', FULL_22), { ok: false, error: 'This attendance export is for 2026-09-22, not for the session on 2026-09-23.' });
  const ambiguous = base64(FX.buildExport().split(FX.DAY).join('9/10/26'));
  check('an export whose date fits neither reading of the session date: refused', r.s.previewTeamsAttendanceImport('s1', ambiguous).error, 'This attendance export is for 2026-09-10 or 2026-10-09, not for the session on 2026-09-22.');
  check('a file that is not a Teams export: refused, with the parser\'s message', r.s.previewTeamsAttendanceImport('s1', base64('TDoc,Title\r\nS4aP260001,Agenda\r\n')),
    { ok: false, error: 'This is not a Teams attendance export: it has no "Summary" and no "Participants" section.' });
  check('broken quoting: refused', r.s.previewTeamsAttendanceImport('s1', base64(FX.buildExport().replace('Kai Invalid', '"Kai Invalid'))).error, 'The file cannot be read: its quoting is broken.');
  check('no file, an empty file', [r.s.previewTeamsAttendanceImport('s1', '').error, r.s.previewTeamsAttendanceImport('s1', null).error, r.s.previewTeamsAttendanceImport('s1', base64('')).ok], ['No file was received.', 'No file was received.', false]);
  check('a file that is too large', r.s.previewTeamsAttendanceImport('s1', 'A'.repeat(4000001)).error, 'The file is too large to be a Teams attendance export.');
  const noDecode = report();
  noDecode.s.Utilities.base64Decode = () => { throw new Error('bad input'); };
  check('bytes that cannot be decoded', noDecode.s.previewTeamsAttendanceImport('s1', '!!!'), { ok: false, error: 'The file could not be read.' });
  check('a session that is not configured', [r.s.previewTeamsAttendanceImport('s9', FULL_22).error, r.s.previewTeamsAttendanceImport('', FULL_22).error, r.s.previewTeamsAttendanceImport(null, FULL_22).ok],
    ['This session is no longer configured. Close the dialog and open it again.', 'This session is no longer configured. Close the dialog and open it again.', false]);
  check('none of the refused previews wrote anything', [r.propsSnapshot() === props0, r.snapshot() === body0, r.writes], [true, true, []]);

  // UTF-8 and comma-separated exports go the same way.
  check('a UTF-8, comma-separated export gives the same preview', JSON.stringify(r.s.previewTeamsAttendanceImport('s1', Buffer.from(FX.buildExport({ delimiter: ',' }), 'utf8').toString('base64')).preview) === JSON.stringify(p.preview), true);

  // Where attendance does not exist.
  [['a main-meeting report', Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }), 'Attendance is available for ad-hoc reports only.'],
    ['a report without a meeting type', { ADHOC_SESSIONS: WITH_SESSIONS.ADHOC_SESSIONS }, 'Attendance is available for ad-hoc reports only.'],
    ['an ad-hoc report without sessions', ADHOC, 'Attendance needs configured sessions. Use Configure Sessions first.'],
    ['an ad-hoc report whose sessions cannot be read', Object.assign({}, ADHOC, { ADHOC_SESSIONS: '{"v":1,' }), 'Attendance needs configured sessions. Use Configure Sessions first.']]
    .forEach(([what, props, message]) => {
      const x = report({ props: props });
      const before = x.propsSnapshot();
      check(`${what}: no preview, no import, nothing written`, [x.s.previewTeamsAttendanceImport('s1', FULL_22), x.s.confirmTeamsAttendanceImport('s1', FULL_22, 'x').error, x.propsSnapshot() === before, x.writes],
        [{ ok: false, error: message }, message, true, []]);
    });
}

// ================================================================ 3. import

console.log('3. import');
{
  const r = report();
  const body0 = r.snapshot();
  const props0 = r.propsSnapshot();
  const preview = r.s.previewTeamsAttendanceImport('s1', FULL_22);

  check('without the preview\'s token nothing is imported', [r.s.confirmTeamsAttendanceImport('s1', FULL_22, ''), r.s.confirmTeamsAttendanceImport('s1', FULL_22, undefined).ok, r.propsSnapshot() === props0, r.snapshot() === body0],
    [{ ok: false, error: 'Nothing was imported: the file or the session is not the one that was previewed. Preview again.' }, false, true, true]);
  check('the token of one session does not import into another', [r.s.confirmTeamsAttendanceImport('s2', base64(smallExport('2026-09-23')), preview.token).ok, r.propsSnapshot() === props0], [false, true]);
  check('the token of one file does not import another file', [r.s.confirmTeamsAttendanceImport('s1', base64(smallExport('2026-09-22')), preview.token).ok, r.propsSnapshot() === props0], [false, true]);
  check('a refused confirmation releases the lock', [r.lock.tries, r.lock.releases], [4, 4]);

  const done = r.s.confirmTeamsAttendanceImport('s1', FULL_22, preview.token);
  check('the confirmed import', done, { ok: true, error: null, changed: true, rendered: true, message: 'The attendance of 22 Sep was imported (23 attendees).' });
  check('it is stored: the index and the chunks of s1, nothing else', [attendanceKeys(r), r.s.adhocAttendanceSessionIds_(r.docProps), r.docProps.getKeys().filter((k) => !(k in JSON.parse(props0)) && k.indexOf(KEY) !== 0)],
    [[KEY, KEY + '_s1_1_0'], ['s1'], []]);
  check('the session configuration is untouched', r.docProps.getProperty('ADHOC_SESSIONS'), WITH_SESSIONS.ADHOC_SESSIONS);
  check('the section is in the report', containerTexts(r), ['Attendance', 'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: September 22, 2026',
    'Start: 14:48:05', 'End: 18:11:26', 'Duration: 3:23:20', 'Attendance records: 25', 'Average attendance: 2:31:07', 'Attendees', '[table 23]']);

  // The same file again.
  const body1 = r.snapshot();
  const props1 = r.propsSnapshot();
  const writes1 = r.writes.length;
  const again = r.s.previewTeamsAttendanceImport('s1', FULL_22);
  check('the same file again: the preview says it is already imported', [again.ok, again.preview.unchanged, again.preview.replaces], [true, true, '']);
  check('confirming it changes nothing: no property written, the document not touched',
    [r.s.confirmTeamsAttendanceImport('s1', FULL_22, again.token), r.writes.length === writes1, r.propsSnapshot() === props1, r.snapshot() === body1],
    [{ ok: true, error: null, changed: false, rendered: false, message: 'This attendance is already imported for 22 Sep. Nothing was changed.' }, true, true, true]);
  check('the same data from a differently encoded file (UTF-8, commas) is the same import', r.s.previewTeamsAttendanceImport('s1', Buffer.from(FX.buildExport({ delimiter: ',' }), 'utf8').toString('base64')).preview.unchanged, true);

  // A second session, then a replacement of the first.
  check('a second session', r.import('s2', base64(smallExport('2026-09-23'))).message, 'The attendance of Offline was imported (4 attendees).');
  const s2raw = r.s.readAdhocAttendance_(r.docProps).raw.s2;
  const replacement = base64(smallExport('2026-09-22'));
  const rp = r.s.previewTeamsAttendanceImport('s1', replacement);
  check('another file for a session that has attendance: the preview says what it replaces', [rp.preview.unchanged, rp.preview.replaces],
    [false, 'This replaces the attendance already imported for this session (23 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE").']);
  check('previewing the replacement changes nothing', r.s.readAdhocAttendance_(r.docProps).sessions.s1.attendees.length, 23);
  check('the replacement', r.s.confirmTeamsAttendanceImport('s1', replacement, rp.token).message, 'The attendance of 22 Sep was imported (4 attendees).');
  check('only that session was replaced', [r.s.readAdhocAttendance_(r.docProps).sessions.s1.attendees.length, r.s.readAdhocAttendance_(r.docProps).raw.s2 === s2raw, attendanceKeys(r).filter((k) => /_s2_/.test(k))], [4, true, [KEY + '_s2_2_0']]);
  check('the report shows both sessions, the first with its new attendance', containerTexts(r).filter((t) => /table|Session/.test(t)), ['Session: September 22, 2026', '[table 4]', 'Session: Offline, September 23, 2026', '[table 4]']);

  // A file imported for one session is not accepted for another.
  check('the file of 22 September is refused for the session of 24 September', r.import('s3', FULL_22), { ok: false, error: 'This attendance export is for 2026-09-22, not for the session on 2026-09-24.' });

  // Lock.
  const busy = report();
  const bp = busy.s.previewTeamsAttendanceImport('s1', FULL_22);
  busy.lock.busy = true;
  check('while an update of the report is running: nothing is imported', [busy.s.confirmTeamsAttendanceImport('s1', FULL_22, bp.token), attendanceKeys(busy), busy.lock.releases],
    [{ ok: false, error: 'Nothing was changed: an update of this report is running right now. Try again in a minute.' }, [], 0]);
  check('a completed import held the lock once and released it', (() => { const x = report(); x.import('s1', FULL_22); return [x.lock.tries, x.lock.releases]; })(), [1, 1]);

  // The master template.
  const master = report({ template: true, docId: TEMPLATE_ID });
  check('in the master template every attendance action refuses', ['previewTeamsAttendanceImport', 'confirmTeamsAttendanceImport', 'removeTeamsAttendanceImport', 'importTeamsAttendance', 'refreshAdhocAttendanceSection']
    .map((fn) => /SA4 Report Template itself/.test(thrown(() => master.s[fn]('s1', FULL_22, 'x')) || '')).concat([attendanceKeys(master).length]), [true, true, true, true, true, 0]);
}

// ================================================================ 4. rendering

console.log('4. rendering');
{
  const r = report();
  r.import('s1', FULL_22);
  const NORMAL = 'NORMAL';
  check('the whole document: the closing prose stays where it was, the container follows it', elements(r.body).slice(0, 6).concat([elements(r.body)[5 + 9]]), [
    ['P', '5.1 Opening of the session', 'H2', false],
    ['P', 'The chair opens the session.', NORMAL, false],
    ['P', '5.11 Close of the session', 'H2', false],
    ['P', 'The chair thanks the participants and the minute takers.', NORMAL, false],
    ['P', 'The session is closed at 18:11 CEST.', NORMAL, false],
    ['P', 'Attendance', 'H3', false],
    ['P', 'Attendees', NORMAL, true]
  ]);
  check('the container is the last thing in the closing section', [texts(r.body).length, texts(r.body)[texts(r.body).length - 1]], [16, '[table 23]']);
  check('"Statistics" and "Attendees" are bold, the values are not', container(r).filter((e) => !isTable(e)).map((e) => [e.getText().split(':')[0], !!e._bold]),
    [['Attendance', false], ['Statistics', true], ['Meeting', false], ['Session', false], ['Start', false], ['End', false], ['Duration', false], ['Attendance records', false], ['Average attendance', false], ['Attendees', true]]);

  const table = attendeeTables(r.body)[0];
  const rows = tableRows(table);
  check('one attendee table: Name | Company | Email, and nothing else', [attendeeTables(r.body).length, rows[0], rows.every((row) => row.length === 3)], [1, ['Name', 'Company', 'Email'], true]);
  check('one row per attendee, in the order of the export', rows.slice(1).map((row) => row[0]), ['Alex Organizer', 'Doe, Jane', 'Sam Rivera', 'Taro Yamada (山田 太郎)', '张 三（示例）', 'Zoë O\'Connor-Müller',
    'Pat Kim (ExampleCorp)', 'Robin Lee - ExampleCorp', 'Morgan Diaz ExampleCorp/Research', 'Dial-in participant', 'Casey Guest', 'Dana Twin', 'Dana Twin', 'Jamie Sample', 'Jamie Sample', 'Lee Wong',
    'Kai Invalid', 'Noor Late', 'Omar Open', 'Sky "Ace" Tab Name', 'Rita Rejoin', 'Quinn, Avery', 'Riley Short (ExampleCorp)']);
  check('Company: only from a leading [Company]; otherwise empty', rows.slice(1).filter((row) => row[1] !== ''), [['Sam Rivera', 'ExampleCorp', '']]);
  check('Email: the explicit address, or empty', [rows.slice(1).filter((row) => row[2]).length, rows.slice(1).filter((row) => !row[2]).map((row) => row[0])],
    [15, ['Sam Rivera', 'Robin Lee - ExampleCorp', 'Dial-in participant', 'Casey Guest', 'Dana Twin', 'Dana Twin', 'Jamie Sample', 'Kai Invalid']]);
  const links = Array.from({ length: table.getNumRows() }, (_, i) => table.getRow(i).getCell(2)._links);
  check('every address is a mailto: link over exactly its text; no other link anywhere in the table',
    [links.filter((l, i) => rows[i][2] && JSON.stringify(l) === JSON.stringify([[0, rows[i][2].length - 1, 'mailto:' + rows[i][2]]])).length, links.filter((l, i) => !rows[i][2] && l.length).length,
      Array.from({ length: table.getNumRows() }, (_, i) => table.getRow(i).getCell(0)._links.length + table.getRow(i).getCell(1)._links.length).reduce((a, b) => a + b, 0)], [15, 0, 0]);
  check('the address is shown as written', rows.filter((row) => /wong/i.test(row[2]))[0][2], 'lee.wong@example.com');
  const whole = r.snapshot();
  check('not in the report: join and leave times, individual durations, roles, markers, Participant IDs, the dial-in number',
    [/15:07:20|17:41:02|2:33:42|9222/.test(whole), /Presenter|Organizer"|\(External\)|\(Unverified\)/.test(whole.replace(/Alex Organizer/g, '')), FX.PRIVATE_STRINGS.filter((x) => whole.indexOf(x) !== -1), /\d{7,}/.test(whole)],
    [false, false, [], false]);
  check('no parser diagnostics in the report', /merged|kept as separate|not a usable address|fewer cells|mean of the listed/.test(whole), false);

  // Statistics are what Teams states.
  const model = r.s.parseTeamsAttendanceReport_(FX.buildExport(), { expectedDate: '2026-09-22' }).attendance;
  check('the report shows Teams\' average and duration, not the calculated ones (which differ)',
    [containerTexts(r).filter((t) => /^(Duration|Average attendance|Attendance records)/.test(t)), r.s.formatTeamsDuration_(model.diagnostics.calculated.meanDurationSeconds), r.s.formatTeamsDuration_(model.diagnostics.calculated.spanSeconds), model.participants.length],
    [['Duration: 3:23:20', 'Attendance records: 25', 'Average attendance: 2:31:07'], '1:55:35', '3:23:21', 23]);
  const sparse = report();
  const sparseText = smallExport('2026-09-22').replace(/^Average attendance time\t.*\r\n/m, '').replace(/^Meeting duration\t.*\r\n/m, '').replace(/^End time\t.*\r\n/m, '').replace(/^Attended participants\t.*\r\n/m, '');
  sparse.import('s1', base64(sparseText));
  check('values Teams did not state are left out, not calculated', containerTexts(sparse), ['Attendance', 'Statistics', 'Meeting: Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE', 'Session: September 22, 2026', 'Start: 14:48:05', 'Attendees', '[table 4]']);
  check('the preview of such a file shows them as missing, with warnings', (() => { const p = report().s.previewTeamsAttendanceImport('s1', base64(sparseText)).preview; return [p.end, p.duration, p.attendanceRecords, p.averageAttendance, p.warnings.length]; })(), ['', '', null, '', 4]);
  const late = report();
  late.import('s1', base64(smallExport('2026-09-22', SMALL, { startTime: '9/22/26, 11:00:00 PM', endTime: '9/23/26, 1:00:00 AM', meetingDuration: '2h' })));
  check('a meeting that ends on the next day shows the end with its date', containerTexts(late).filter((t) => /^(Start|End)/.test(t)), ['Start: 23:00:00', 'End: 2026-09-23 01:00:00']);
  check('no time zone is written', /CEST|CET|UTC|GMT/.test(JSON.stringify(containerTexts(r))), false);

  // Order: the sessions' chronological order, not the order of import.
  const order = report();
  order.import('s3', base64(smallExport('2026-09-24')));
  check('attendance for the third session only: one block, no empty block for the others', containerTexts(order).filter((t) => /^Session|table/.test(t)), ['Session: September 24, 2026', '[table 4]']);
  order.import('s1', FULL_22);
  order.import('s2', base64(smallExport('2026-09-23', SMALL.slice(0, 2))));
  check('imported third, first, second: rendered first, second, third', containerTexts(order).filter((t) => /^Session|table/.test(t)),
    ['Session: September 22, 2026', '[table 23]', 'Session: Offline, September 23, 2026', '[table 2]', 'Session: September 24, 2026', '[table 4]']);
  check('one container, one heading, three tables', [texts(order.body).filter((t) => t === 'Attendance').length, attendeeTables(order.body).length, texts(order.body).filter((t) => t === 'Statistics').length], [1, 3, 3]);
  check('the prose of the closing section is still there, once, above the container',
    [texts(order.body).slice(0, 6), texts(order.body).filter((t) => /thanks the participants|closed at 18:11/.test(t)).length], [texts(r.body).slice(0, 6), 2]);

  // Idempotent.
  const before = order.snapshot();
  order.s.refreshAdhocAttendanceSection();
  order.s.refreshAdhocAttendanceSection();
  check('rendering again and again gives the same document', order.snapshot() === before, true);

  // Generated content stays generated.
  const c = order.s.findAdhocAttendanceContainer_(order.body);
  order.body.insertParagraph(c.start + 2, 'A note typed into the generated section.');
  order.body._children[c.start + 1].setText('Statistics (edited)');
  order.s.refreshAdhocAttendanceSection();
  check('text typed into the container is not kept; the container is written as before', order.snapshot() === before, true);
  order.body._children[3].setText('The chair thanks everybody. (edited by hand)');
  order.s.refreshAdhocAttendanceSection();
  check('text typed above the container is kept', texts(order.body)[3], 'The chair thanks everybody. (edited by hand)');

  // Placement.
  const after = report({ after: true });
  after.import('s1', FULL_22);
  check('with a section after the closing one: the container ends the closing section', [texts(after.body).indexOf('Attendance'), texts(after.body)[texts(after.body).length - 1], texts(after.body).length], [5, 'Annex A Agenda', 17]);
  after.import('s2', base64(smallExport('2026-09-23')));
  check('and stays there when it grows; the following section is untouched', [texts(after.body).indexOf('Attendance'), texts(after.body)[texts(after.body).length - 1], attendeeTables(after.body).length], [5, 'Annex A Agenda', 2]);
  const noClosing = report({ closing: false });
  const result = noClosing.import('s1', FULL_22);
  check('without a closing section: at the end of the document, and the result says so',
    [texts(noClosing.body).slice(0, 3), result.message], [['5.1 Opening of the session', 'The chair opens the session.', 'Attendance'],
      'The attendance of 22 Sep was imported (23 attendees). No closing section was found; the Attendance section is at the end of the document.']);
  const twoClosings = report();
  twoClosings.body.insertParagraph(0, '2.9 Closing remarks of the first day').setHeading(twoClosings.s.DocumentApp.ParagraphHeading.HEADING2);
  twoClosings.import('s1', FULL_22);
  check('with two closing headings: under the last one', texts(twoClosings.body).slice(0, 7), ['2.9 Closing remarks of the first day', '5.1 Opening of the session', 'The chair opens the session.', '5.11 Close of the session',
    'The chair thanks the participants and the minute takers.', 'The session is closed at 18:11 CEST.', 'Attendance']);
  const plainWord = report();
  plainWord.body.insertParagraph(1, 'Attendance');
  plainWord.import('s1', FULL_22);
  check('an ordinary paragraph that reads "Attendance" is not the container', [texts(plainWord.body)[1], texts(plainWord.body).filter((t) => t === 'Attendance').length, texts(plainWord.body)[2]], ['Attendance', 2, 'The chair opens the session.']);

  // A document must end with a paragraph.
  const tail = report();
  tail.import('s1', FULL_22);
  tail.body.appendParagraph('');
  const removeChild = tail.body.removeChild;
  tail.body.removeChild = (child) => { if (tail.body.getChildIndex(child) === tail.body.getNumChildren() - 1 && !isTable(child)) throw new Error('Can\'t remove the last paragraph in a document section.'); return removeChild(child); };
  check('the last paragraph of the document is never removed: replacing the container works', [thrown(() => tail.s.refreshAdhocAttendanceSection()), texts(tail.body).slice(-2), attendeeTables(tail.body).length], [null, ['[table 23]', ''], 1]);
  check('and so does removing it', [tail.s.removeTeamsAttendanceImport('s1').ok, texts(tail.body)], [true, ['5.1 Opening of the session', 'The chair opens the session.', '5.11 Close of the session', 'The chair thanks the participants and the minute takers.', 'The session is closed at 18:11 CEST.', '']]);

  // Pure block builder.
  const state = r.s.readAdhocAttendance_(r.docProps);
  const blocks = r.s.buildAdhocAttendanceBlocks_(r.s.getAdhocSessions_(), state);
  check('blocks (pure): one per session with attendance', [blocks.length, blocks[0].sessionId, blocks[0].statistics, blocks[0].attendees.length],
    [1, 's1', [['Meeting', 'Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE'], ['Session', 'September 22, 2026'], ['Start', '14:48:05'], ['End', '18:11:26'], ['Duration', '3:23:20'], ['Attendance records', '25'], ['Average attendance', '2:31:07']], 23]);
  check('attendance of a session that is not configured is not rendered', r.s.buildAdhocAttendanceBlocks_([{ id: 's7', label: '', date: '2026-09-22', displayLabel: '22 Sep' }], state), []);
  check('a duration of zero is a value, not a missing one', r.s.buildAdhocAttendanceBlocks_(r.s.getAdhocSessions_(), { sessions: { s1: Object.assign({}, state.sessions.s1, { durationSeconds: 0, attendanceRecords: 0 }) }, companies: {} })[0].statistics.filter((l) => /Duration|records/.test(l[0])),
    [['Duration', '0:00:00'], ['Attendance records', '0']]);
}

// ================================================================ 5. manual Company values

console.log('5. manual Company values');
{
  const r = report();
  r.import('s1', FULL_22);
  check('typing into Company cells', [typeCompany(r, 0, 'Alex Organizer', 'Example Chair Corp'), typeCompany(r, 0, 'Casey Guest', 'Guest Company'), typeCompany(r, 0, 'Sam Rivera', 'Example Corporation Ltd'), typeCompany(r, 0, 'Dana Twin', 'Twin Company')], [true, true, true, true]);
  r.s.refreshAdhocAttendanceSection();
  check('they survive a refresh: by e-mail, by an unambiguous name, and over the company the parser found',
    [companyOf(r, 0, 'Alex Organizer'), companyOf(r, 0, 'Casey Guest'), companyOf(r, 0, 'Sam Rivera')], ['Example Chair Corp', 'Guest Company', 'Example Corporation Ltd']);
  check('a value typed for one of two people with the same name is not kept: it could be either', tableRows(attendeeTables(r.body)[0]).filter((row) => row[0] === 'Dana Twin').map((row) => row[1]), ['', '']);
  check('what is stored: e-mail keys and name keys, only for the values that were typed', r.s.readAdhocAttendance_(r.docProps).companies,
    { 'email:alex.organizer@example.com': 'Example Chair Corp', 'name:sam rivera': 'Example Corporation Ltd', 'name:casey guest': 'Guest Company' });
  check('the stored attendance itself is unchanged (the parser\'s company is still there)', r.s.readAdhocAttendance_(r.docProps).sessions.s1.attendees.filter((a) => a.name === 'Sam Rivera'), [{ name: 'Sam Rivera', company: 'ExampleCorp', email: '' }]);

  // They carry over to other sessions.
  const day2 = [person('Organizer, Alex (External)', 'ALEX.ORGANIZER@example.com'), person('Casey Guest (Unverified)'), person('Alex Organizer'), person('Sam Rivera (Unverified)'), person('New Person', 'new.person@example.org')];
  r.import('s2', base64(smallExport('2026-09-23', day2)));
  check('another session: the same e-mail gets the company, whatever the display name', companyOf(r, 1, 'Organizer, Alex'), 'Example Chair Corp');
  check('the same name without e-mail gets it too', [companyOf(r, 1, 'Casey Guest'), companyOf(r, 1, 'Sam Rivera')], ['Guest Company', 'Example Corporation Ltd']);
  check('a guest who merely has the name of someone with an e-mail does not get that person\'s company', companyOf(r, 1, 'Alex Organizer'), '');
  check('someone new has none', companyOf(r, 1, 'New Person'), '');

  // Re-import and replacement keep them.
  r.import('s1', base64(smallExport('2026-09-22', SMALL)));
  check('they survive the replacement of the import', [companyOf(r, 0, 'Alex Organizer'), companyOf(r, 0, 'Casey Guest'), companyOf(r, 0, 'Sam Rivera')], ['Example Chair Corp', 'Guest Company', 'Example Corporation Ltd']);

  // Changing and clearing.
  typeCompany(r, 1, 'Organizer, Alex', 'Example Chair Corporation');
  r.s.refreshAdhocAttendanceSection();
  check('a value changed in one table applies to that attendee in every table', [companyOf(r, 0, 'Alex Organizer'), companyOf(r, 1, 'Organizer, Alex')], ['Example Chair Corporation', 'Example Chair Corporation']);
  typeCompany(r, 0, 'Alex Organizer', 'One Corp');
  typeCompany(r, 1, 'Organizer, Alex', 'Another Corp');
  r.s.refreshAdhocAttendanceSection();
  check('two different new values for one attendee decide nothing: the stored value stays', [companyOf(r, 0, 'Alex Organizer'), companyOf(r, 1, 'Organizer, Alex')], ['Example Chair Corporation', 'Example Chair Corporation']);
  typeCompany(r, 0, 'Casey Guest', '');
  r.s.refreshAdhocAttendanceSection();
  check('a cell cleared in one table only: the value is still in the other table, so it stays', [companyOf(r, 0, 'Casey Guest'), companyOf(r, 1, 'Casey Guest')], ['Guest Company', 'Guest Company']);
  typeCompany(r, 0, 'Casey Guest', '');
  typeCompany(r, 1, 'Casey Guest', '');
  r.s.refreshAdhocAttendanceSection();
  check('cleared everywhere: the manual value is removed', [companyOf(r, 0, 'Casey Guest'), companyOf(r, 1, 'Casey Guest'), 'name:casey guest' in r.s.readAdhocAttendance_(r.docProps).companies], ['', '', false]);
  typeCompany(r, 0, 'Sam Rivera', '');
  typeCompany(r, 1, 'Sam Rivera', '');
  r.s.refreshAdhocAttendanceSection();
  check('clearing a manual value brings back the company the parser found; an empty cell never erases that', [companyOf(r, 0, 'Sam Rivera'), companyOf(r, 1, 'Sam Rivera')], ['ExampleCorp', '']);
  typeCompany(r, 0, 'Sam Rivera', '');
  r.s.refreshAdhocAttendanceSection();
  check('clearing the parser\'s company in the table does not remove it', companyOf(r, 0, 'Sam Rivera'), 'ExampleCorp');
  typeCompany(r, 0, 'Taro Yamada (山田 太郎)', '  示例  株式会社 \n Example KK  ');
  r.s.refreshAdhocAttendanceSection();
  check('a value in another script, with stray white space', companyOf(r, 0, 'Taro Yamada (山田 太郎)'), '示例 株式会社 Example KK');

  // An edited name or address is not a way to rename an attendee.
  const t = attendeeTables(r.body)[0];
  t.getRow(1).getCell(0).setText('Somebody Else');
  t.getRow(1).getCell(2).setText('');
  t.getRow(1).getCell(1).setText('Wrong Corp');
  r.s.refreshAdhocAttendanceSection();
  check('a row whose name and address were overwritten matches nobody: its company goes nowhere, the attendee is written as stored',
    [tableRows(attendeeTables(r.body)[0])[1], JSON.stringify(r.s.readAdhocAttendance_(r.docProps).companies).indexOf('Wrong Corp')], [['Alex Organizer', 'Example Chair Corporation', 'alex.organizer@example.com'], -1]);

  // Removing an import drops the corrections of people who are in no session any more.
  r.s.removeTeamsAttendanceImport('s2');
  check('after removing a session: corrections of attendees still present are kept', Object.keys(r.s.readAdhocAttendance_(r.docProps).companies).sort(), ['email:alex.organizer@example.com', 'email:taro.yamada@example.net']);
  r.s.removeTeamsAttendanceImport('s1');
  check('after removing the last one: nothing about anybody is stored', attendanceKeys(r), []);

  // The merge, as a pure function.
  const merge = r.s.mergeAdhocCompanyCorrections_;
  const A = (name, company, email) => ({ name: name, company: company || '', email: email || '' });
  const sessions = { s1: { attendees: [A('Pat Kim', '', 'pat.kim@example.com'), A('Sam Rivera', 'ExampleCorp'), A('Dana Twin'), A('Dana Twin')] }, s2: { attendees: [A('Pat Kim', '', 'pat.kim@example.com'), A('Dana Twin')] } };
  const H = (key, company) => ({ key: key, company: company });
  check('merge: nothing harvested keeps what is stored', merge({ 'email:pat.kim@example.com': 'Kim Corp' }, sessions, []), { 'email:pat.kim@example.com': 'Kim Corp' });
  check('merge: a typed value is stored', merge({}, sessions, [H('email:pat.kim@example.com', 'Kim Corp'), H('email:pat.kim@example.com', '')]), { 'email:pat.kim@example.com': 'Kim Corp' });
  check('merge: a value equal to what the parser found is not a correction', merge({}, sessions, [H('name:sam rivera', 'ExampleCorp')]), {});
  check('merge: a row for somebody who is in no session is ignored', merge({}, sessions, [H('email:nobody@example.com', 'Nowhere Corp'), H('name:somebody else', 'Nowhere Corp')]), {});
  check('merge: a name that is unambiguous in one session', merge({}, sessions, [H('name:dana twin', 'Twin Corp')]), { 'name:dana twin': 'Twin Corp' });
  check('merge: a stored correction of somebody who is in no session any more is dropped', merge({ 'email:gone@example.com': 'Gone Corp', 'email:pat.kim@example.com': 'Kim Corp' }, sessions, []), { 'email:pat.kim@example.com': 'Kim Corp' });
  check('merge: the inputs are not changed', (() => { const stored = { 'email:pat.kim@example.com': 'Kim Corp' }; merge(stored, sessions, [H('email:pat.kim@example.com', 'New Corp')]); return stored; })(), { 'email:pat.kim@example.com': 'Kim Corp' });
  const eff = r.s.effectiveAdhocAttendees_;
  check('applying: an ambiguous name in a session gets nothing there, an unambiguous one does',
    [eff(sessions.s1, { 'name:dana twin': 'Twin Corp', 'email:pat.kim@example.com': 'Kim Corp' }).map((a) => a.company), eff(sessions.s2, { 'name:dana twin': 'Twin Corp' }).map((a) => a.company)],
    [['Kim Corp', 'ExampleCorp', '', ''], ['', 'Twin Corp']]);
  check('a name key never applies to an attendee with an e-mail, nor an e-mail key to one without',
    eff({ attendees: [A('Pat Kim', '', 'pat.kim@example.com'), A('pat.kim@example.com')] }, { 'name:pat kim': 'X Corp', 'email:other@example.com': 'Y Corp' }).map((a) => a.company), ['', '']);

  // Two people of one name in one session, one person of that name in another.
  const twins = report();
  twins.import('s1', FULL_22);
  twins.import('s2', base64(smallExport('2026-09-23', [person('Dana Twin (Unverified)'), person('Alex Organizer', 'alex.organizer@example.com')])));
  const twinRows = () => tableRows(attendeeTables(twins.body)[0]).map((row, i) => [i, row]).filter((x) => x[1][0] === 'Dana Twin');
  attendeeTables(twins.body)[0].getRow(twinRows()[0][0]).getCell(1).setText('Twin Company');
  twins.s.refreshAdhocAttendanceSection();
  check('a value typed for one of two namesakes is not carried to the session where the name is unambiguous: it is unknown whose it was',
    [twinRows().map((x) => x[1][1]), companyOf(twins, 1, 'Dana Twin'), twins.s.readAdhocAttendance_(twins.docProps).companies], [['', ''], '', {}]);
  typeCompany(twins, 1, 'Dana Twin', 'Twin Company');
  twins.s.refreshAdhocAttendanceSection();
  check('typed where the name is unambiguous, it is kept there, and still not given to either namesake', [companyOf(twins, 1, 'Dana Twin'), twinRows().map((x) => x[1][1])], ['Twin Company', ['', '']]);

  // Harvest reads only attendee tables.
  const h = report();
  h.import('s1', base64(smallExport('2026-09-22')));
  h.body.appendTable([['Name', 'Company'], ['Alex Organizer', 'Two Column Corp']]);
  h.body.appendTable([['Name', 'Company', 'Email', 'Phone'], ['Alex Organizer', 'Four Column Corp', 'alex.organizer@example.com', '']]);
  h.body.appendTable([['TDoc', 'Title', 'Source'], ['Alex Organizer', 'Not A Company', 'alex.organizer@example.com']]);
  check('only a table headed exactly Name | Company | Email is read', h.s.harvestAdhocAttendeeCompanies_(h.body).filter((x) => x.company), [{ key: 'name:sam rivera', company: 'ExampleCorp' }]);
  check('a company is at most 200 characters', (() => { typeCompany(h, 0, 'Alex Organizer', 'x'.repeat(500)); return h.s.harvestAdhocAttendeeCompanies_(h.body)[0].company.length; })(), 200);
}

// ================================================================ 6. sessions that change

console.log('6. sessions that change');
{
  const r = report();
  r.import('s1', FULL_22);
  r.import('s3', base64(smallExport('2026-09-24')));
  const stored = r.propsSnapshot();
  const rowsOf = (list) => list.map((x) => Object.assign({}, x));

  // Renaming and re-dating keep the attendance.
  const renamed = r.s.saveAdhocSessionsConfiguration(rowsOf([session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), THREE_SESSIONS[1], THREE_SESSIONS[2]]));
  check('renaming a session: saved, the attendance is still that session\'s', [renamed.ok, renamed.changed, renamed.notice, r.s.adhocAttendanceSessionIds_(r.docProps)], [true, true, undefined, ['s1', 's3']]);
  check('the report shows the new label at once', containerTexts(r).filter((t) => /^Session/.test(t)), ['Session: Kick-off, September 22, 2026', 'Session: September 24, 2026']);
  // A session's day cannot be changed away from its imported attendance (range integrity).
  const sessionsBefore = r.docProps.getProperty('ADHOC_SESSIONS');
  const moved = r.s.saveAdhocSessionsConfiguration(rowsOf([session('s1', '2026-09-25', '', '', 'Kick-off'), THREE_SESSIONS[1], THREE_SESSIONS[2]]));
  check('changing the date of a session away from its imported attendance: refused, with the way out; nothing is changed', [moved, r.docProps.getProperty('ADHOC_SESSIONS') === sessionsBefore, r.s.readAdhocAttendance_(r.docProps).sessions.s1.attendees.length],
    [{ ok: false, errors: ['Kick-off cannot be changed to September 25, 2026 because attendance for September 22, 2026 is already imported. Remove that attendance first (Sessions and Attendance > Import Teams Attendance… > Remove attendance…) or keep that day within the session.'], changed: false, sessions: [] }, true, 23]);
  const retimed = r.s.saveAdhocSessionsConfiguration(rowsOf([session('s1', '2026-09-22', '14:00', '19:00', 'Kick-off'), THREE_SESSIONS[1], THREE_SESSIONS[2]]));
  check('changing its planned times: saved; the report keeps its order, and the times shown are Teams\'', [retimed.ok, containerTexts(r).filter((t) => /^(Session|Start)|table/.test(t))],
    [true, ['Session: Kick-off, September 22, 2026', 'Start: 14:48:05', '[table 23]', 'Session: September 24, 2026', 'Start: 14:48:05', '[table 4]']]);
  check('the stored attendance was not rewritten by any of these saves', attendanceKeys(r), [KEY, KEY + '_s1_1_0', KEY + '_s3_2_0']);
  const same = r.s.saveAdhocSessionsConfiguration(rowsOf([session('s1', '2026-09-22', '14:00', '19:00', 'Kick-off'), THREE_SESSIONS[1], THREE_SESSIONS[2]]));
  const bodyBefore = r.snapshot();
  check('saving unchanged sessions does not touch the report', [same.changed, r.snapshot() === bodyBefore], [false, true]);

  // A session with attendance is not removed.
  const before = r.propsSnapshot();
  const bodyBeforeDelete = r.snapshot();
  const refused = r.s.saveAdhocSessionsConfiguration(rowsOf([THREE_SESSIONS[1], THREE_SESSIONS[2]]));
  check('removing a session that has attendance is refused, with the way to do it', refused,
    { ok: false, errors: ['Kick-off has imported attendance and was not removed. Remove its attendance first: Sessions and Attendance > Import Teams Attendance… > Remove attendance….'], changed: false, sessions: [] });
  check('nothing was changed: not the sessions, not the attendance, not the report', [r.propsSnapshot() === before, r.snapshot() === bodyBeforeDelete], [true, true]);
  check('removing both sessions with attendance: both are named', r.s.saveAdhocSessionsConfiguration(rowsOf([THREE_SESSIONS[1]])).errors.map((e) => e.split(' has ')[0]), ['Kick-off', '24 Sep']);
  check('removing every session: refused as well', r.s.saveAdhocSessionsConfiguration([]).ok, false);
  const freed = r.s.saveAdhocSessionsConfiguration(rowsOf([session('s1', '2026-09-22', '14:00', '19:00', 'Kick-off'), THREE_SESSIONS[2]]));
  check('a session without attendance can be removed', [freed.ok, r.s.getAdhocSessions_().map((x) => x.id), r.s.adhocAttendanceSessionIds_(r.docProps)], [true, ['s1', 's3'], ['s1', 's3']]);

  // After its attendance was removed explicitly, the session can go.
  check('after removing its attendance', [r.s.removeTeamsAttendanceImport('s1').ok, r.s.saveAdhocSessionsConfiguration(rowsOf([THREE_SESSIONS[2]])).ok, r.s.getAdhocSessions_().map((x) => x.id), r.s.adhocAttendanceSessionIds_(r.docProps)], [true, true, ['s3'], ['s3']]);
  check('no attendance is left without its session', r.s.adhocAttendanceSessionIds_(r.docProps).filter((id) => r.s.getAdhocSessions_().every((x) => x.id !== id)), []);
  check('a session added afterwards does not get the id of the removed ones', r.s.saveAdhocSessionsConfiguration([THREE_SESSIONS[2], session('', '2026-10-01')]).sessions.map((x) => x.id), ['s3', 's4']);

  // The pure function, with what it is told about attendance.
  const S = loadCode().sandbox;
  const make = () => { const l = loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(THREE_SESSIONS) } }); return l; };
  const tell = (ids, rows) => { const l = make(); return [l.sandbox.saveAdhocSessionsWith_(l.docProps, true, rows, ids).ok, l.docProps.getProperty('ADHOC_SESSIONS') === sessionsProperty(THREE_SESSIONS)]; };
  check('no attendance anywhere: sessions are removed as in stage A', [tell([], rowsOf(THREE_SESSIONS.slice(1))), tell(undefined, rowsOf(THREE_SESSIONS.slice(1)))], [[true, false], [true, false]]);
  check('attendance on a session that stays: the others can be removed', tell(['s2'], rowsOf([THREE_SESSIONS[1]])), [true, false]);
  check('attendance on the session to be removed: refused, nothing written', tell(['s1'], rowsOf(THREE_SESSIONS.slice(1))), [false, true]);
  check('attendance of a newer schema version (which sessions have it is unknown): no session is removed', [tell(null, rowsOf(THREE_SESSIONS.slice(1))), make().sandbox.saveAdhocSessionsWith_(make().docProps, true, rowsOf(THREE_SESSIONS.slice(1)), null).errors[0]],
    [[false, true], 'The stored attendance was written by a newer release, so it cannot be told which sessions have attendance. No session was removed.']);
  check('and then sessions can still be edited and added', tell(null, rowsOf(THREE_SESSIONS).concat([session('', '2026-10-01')])), [true, false]);
  check('the stage A behaviour without the fourth argument is unchanged', (() => { const l = make(); return l.sandbox.saveAdhocSessionsWith_(l.docProps, true, []).ok; })(), true);
  void S;

  // The section cannot be written.
  const failing = report();
  failing.import('s1', FULL_22);
  failing.s.renderAdhocAttendanceSection_ = () => { throw new Error('document is busy'); };
  const result = failing.s.saveAdhocSessionsConfiguration(rowsOf([session('s1', '2026-09-22', '15:00', '18:00', 'Renamed'), THREE_SESSIONS[1], THREE_SESSIONS[2]]));
  check('if the section cannot be written, the sessions are saved all the same and the result says how to write it later',
    [result.ok, result.changed, result.notice, failing.s.getAdhocSessions_()[0].label], [true, true, 'The sessions were saved. The Attendance section could not be updated (document is busy). Use SA4 Report > Sessions and Attendance > Refresh Attendance Section.', 'Renamed']);
  check('the Configure Sessions dialog shows such a notice before it closes (source)', /if \(result\.notice\) alert\(result\.notice\);\n\s+google\.script\.host\.close\(\);/.test(CODE), true);
  const noAttendance = report();
  const bodyNo = noAttendance.snapshot();
  noAttendance.s.saveAdhocSessionsConfiguration(rowsOf([session('s1', '2026-09-22', '', '', 'Renamed'), THREE_SESSIONS[1], THREE_SESSIONS[2]]));
  check('saving sessions in a report without attendance does not touch the document or take a lock', [noAttendance.snapshot() === bodyNo, noAttendance.lock.tries], [true, 0]);
}

// ================================================================ 7. removing an import

console.log('7. removing an import');
{
  const r = report();
  r.import('s1', FULL_22);
  r.import('s2', base64(smallExport('2026-09-23')));
  const s2raw = r.s.readAdhocAttendance_(r.docProps).raw.s2;
  const removed = r.s.removeTeamsAttendanceImport('s1');
  check('removing one session\'s attendance', removed, { ok: true, error: null, changed: true, rendered: true, message: 'The attendance of 22 Sep was removed.' });
  check('its stored data is gone, the other session\'s is untouched', [attendanceKeys(r), r.s.readAdhocAttendance_(r.docProps).raw.s2 === s2raw], [[KEY, KEY + '_s2_2_0'], true]);
  check('its block is gone from the report, the other stays', containerTexts(r).filter((t) => /^Session|table/.test(t)), ['Session: Offline, September 23, 2026', '[table 4]']);
  check('removing it again', r.s.removeTeamsAttendanceImport('s1'), { ok: false, error: 'This session has no imported attendance.' });
  check('a session that never had any, a session that does not exist', [r.s.removeTeamsAttendanceImport('s3').error, r.s.removeTeamsAttendanceImport('s9').error, r.s.removeTeamsAttendanceImport('companies').error],
    ['This session has no imported attendance.', 'This session has no imported attendance.', 'This session has no imported attendance.']);
  const last = r.s.removeTeamsAttendanceImport('s2');
  // The last paragraph of a document is never removed (see part 4); here that paragraph was the heading, which is emptied.
  check('removing the last one: the container is gone, the closing prose is as it was, nothing is stored',
    [last.ok, texts(r.body), r.s.findAdhocAttendanceContainer_(r.body), r.body._children[5].getHeading(), attendanceKeys(r)],
    [true, ['5.1 Opening of the session', 'The chair opens the session.', '5.11 Close of the session', 'The chair thanks the participants and the minute takers.', 'The session is closed at 18:11 CEST.', ''], null, 'NORMAL', []]);
  check('the sessions themselves are still configured', r.s.getAdhocSessions_().length, 3);

  const busy = report();
  busy.import('s1', FULL_22);
  busy.lock.busy = true;
  check('while an update is running nothing is removed', [busy.s.removeTeamsAttendanceImport('s1').ok, busy.s.adhocAttendanceSessionIds_(busy.docProps)], [false, ['s1']]);
  const main = report({ props: Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }) });
  check('not in a main-meeting report', main.s.removeTeamsAttendanceImport('s1'), { ok: false, error: 'Attendance is available for ad-hoc reports only.' });
  const newer = report();
  newer.docProps.setProperty(KEY, JSON.stringify({ v: 2, gen: 1, entries: {} }));
  check('stored attendance of a newer version is not touched', [newer.s.removeTeamsAttendanceImport('s1').error, newer.r, newer.docProps.getProperty(KEY)], ['The stored attendance was written by a newer release and is not changed here.', undefined, '{"v":2,"gen":1,"entries":{}}']);
  check('nor imported into', [newer.s.previewTeamsAttendanceImport('s1', FULL_22).ok, /newer release/.test(newer.s.previewTeamsAttendanceImport('s1', FULL_22).error)], [false, true]);
}

// ================================================================ 8. failure and recovery

console.log('8. failure and recovery');
{
  // Storing fails.
  const a = report();
  const pa = a.s.previewTeamsAttendanceImport('s1', FULL_22);
  const body0 = a.snapshot();
  const set = a.docProps.setProperty;
  a.docProps.setProperty = (k) => { if (k === KEY) throw new Error('Properties storage quota exceeded'); set.apply(null, arguments); };
  const failed = a.s.confirmTeamsAttendanceImport('s1', FULL_22, pa.token);
  a.docProps.setProperty = set;
  check('storing fails: nothing is imported, the report is not touched, the lock is released',
    [failed, a.s.adhocAttendanceSessionIds_(a.docProps), a.snapshot() === body0, a.lock.releases],
    [{ ok: false, error: 'Nothing was imported: the attendance could not be stored (Properties storage quota exceeded).' }, [], true, 1]);

  // Rendering fails after storing.
  const b = report();
  const pb = b.s.previewTeamsAttendanceImport('s1', FULL_22);
  const realRender = b.s.renderAdhocAttendanceSection_;
  b.s.renderAdhocAttendanceSection_ = () => { throw new Error('Service Documents failed'); };
  const half = b.s.confirmTeamsAttendanceImport('s1', FULL_22, pb.token);
  check('rendering fails after storing: the import is kept and the result says how to write the section',
    [half, b.s.readAdhocAttendance_(b.docProps).sessions.s1.attendees.length, containerTexts(b)],
    [{ ok: true, error: null, changed: true, rendered: false, message: 'The attendance of 22 Sep was imported (23 attendees). It is stored, but the Attendance section could not be written (Service Documents failed). Use SA4 Report > Sessions and Attendance > Refresh Attendance Section.' }, 23, null]);
  check('the failure is logged without any participant data', [b.logs, b.logs.filter((l) => /@|Organizer|Rivera/.test(l))], [['Attendance: stored, but the section could not be written: Service Documents failed'], []]);
  b.s.renderAdhocAttendanceSection_ = realRender;
  check('Refresh Attendance Section then writes it from what is stored', [thrown(() => b.s.refreshAdhocAttendanceSection()), containerTexts(b).length, attendeeTables(b.body)[0].getNumRows() - 1], [null, 11, 23]);
  const c = report();
  const pc = c.s.previewTeamsAttendanceImport('s1', FULL_22);
  c.s.renderAdhocAttendanceSection_ = () => { throw new Error('Service Documents failed'); };
  c.s.confirmTeamsAttendanceImport('s1', FULL_22, pc.token);
  c.s.renderAdhocAttendanceSection_ = realRender;
  const retry = c.s.previewTeamsAttendanceImport('s1', FULL_22);
  check('importing the same file again also writes the missing section, without storing anything again',
    [retry.preview.unchanged, c.s.confirmTeamsAttendanceImport('s1', FULL_22, retry.token), containerTexts(c).length, attendanceKeys(c)],
    [true, { ok: true, error: null, changed: false, rendered: true, message: 'This attendance was already imported for 22 Sep; the Attendance section was written again.' }, 11, [KEY, KEY + '_s1_1_0']]);

  // A stored value that cannot be read.
  const d = report();
  d.import('s1', FULL_22);
  d.import('s2', base64(smallExport('2026-09-23')));
  const bodyOk = d.snapshot();
  d.docProps.deleteProperty(KEY + '_s1_1_0');
  check('a session whose stored attendance is damaged: the refresh stops before it touches the report',
    [thrown(() => d.s.refreshAdhocAttendanceSection()), d.snapshot() === bodyOk], ['The stored attendance of 22 Sep cannot be read. Import it again or remove it; the report was not changed.', true]);
  const dp = d.s.previewTeamsAttendanceImport('s1', FULL_22);
  check('importing it again is offered as a replacement', [dp.ok, dp.preview.unchanged, dp.preview.replaces], [true, false, 'This replaces the attendance stored for this session, which cannot be read.']);
  check('and repairs it', [d.s.confirmTeamsAttendanceImport('s1', FULL_22, dp.token).rendered, d.s.readAdhocAttendance_(d.docProps).problems, d.snapshot() === bodyOk], [true, [], true]);
  d.docProps.deleteProperty(d.docProps.getKeys().filter((k) => /_s1_\d+_0$/.test(k))[0]);
  check('or it can be removed', [d.s.removeTeamsAttendanceImport('s1').ok, d.s.readAdhocAttendance_(d.docProps).problems, containerTexts(d).filter((t) => /table/.test(t))], [true, [], ['[table 4]']]);

  // An index that cannot be read.
  const e = report();
  e.import('s2', base64(smallExport('2026-09-23')));
  const bodyE = e.snapshot();
  e.docProps.setProperty(KEY, '{"v":1,');
  check('an index that cannot be read: the refresh stops, the report is not touched', [thrown(() => e.s.refreshAdhocAttendanceSection()), e.snapshot() === bodyE], ['The stored attendance index cannot be read.', true]);
  check('a new import starts over: only what is imported now is stored and rendered', [e.import('s1', FULL_22).rendered, e.s.adhocAttendanceSessionIds_(e.docProps), containerTexts(e).filter((t) => /table/.test(t)), attendanceKeys(e).filter((k) => /_s2_/.test(k))],
    [true, ['s1'], ['[table 23]'], []]);

  // Refresh without anything to do, and where it does not apply.
  const f = report();
  const bodyF = f.snapshot();
  check('a refresh without any attendance changes nothing and writes nothing', [thrown(() => f.s.refreshAdhocAttendanceSection()), f.snapshot() === bodyF, f.writes], [null, true, []]);
  check('a refresh without sessions, or in a main-meeting report, is an error', [thrown(() => report({ props: ADHOC }).s.refreshAdhocAttendanceSection()), thrown(() => report({ props: Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }) }).s.refreshAdhocAttendanceSection())],
    ['Attendance needs an ad-hoc report with configured sessions. Use Configure Sessions first.', 'Attendance needs an ad-hoc report with configured sessions. Use Configure Sessions first.']);
  f.lock.busy = true;
  check('a refresh while an update is running is an error', thrown(() => f.s.refreshAdhocAttendanceSection()), 'Nothing was changed: an update of this report is running right now. Try again in a minute.');
  check('a successful refresh shows no dialog (as Update Report Now)', (() => { const x = report(); x.import('s1', FULL_22); x.s.refreshAdhocAttendanceSection(); return x.ui.alerts.length + x.ui.dialogs.length; })(), 0);
}

// ================================================================ 9. the dialog

/** Runs the import dialog's client script against a minimal DOM. `server` answers the RPCs. */
function runDialog(html, server) {
  const scripts = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
  const element = (tag) => {
    const e = { tag: tag, children: [], style: {}, value: '', textContent: '', className: '', disabled: false, files: [] };
    e.appendChild = (c) => { e.children.push(c); if (e.children.length === 1 && c.value !== undefined && !e.value) e.value = c.value; return c; };
    e.removeChild = (c) => { e.children.splice(e.children.indexOf(c), 1); return c; };
    Object.defineProperty(e, 'firstChild', { get: () => e.children[0] || null });
    return e;
  };
  const byId = {};
  const idsInHtml = (html.match(/id="([A-Za-z]+)"/g) || []).map((m) => m.slice(4, -1));
  idsInHtml.forEach((id) => { byId[id] = element(id); });
  const calls = [];
  let handlers = {};
  const runner = {
    withSuccessHandler: (fn) => { handlers.success = fn; return runner; },
    withFailureHandler: (fn) => { handlers.failure = fn; return runner; }
  };
  ['previewTeamsAttendanceImport', 'confirmTeamsAttendanceImport', 'removeTeamsAttendanceImport'].forEach((name) => {
    runner[name] = (...args) => { calls.push([name].concat(args)); const h = handlers; handlers = {}; if (server) h.success(JSON.parse(JSON.stringify(server[name](...args)))); else calls.handlers = h; };
  });
  let closed = 0;
  function FileReader() {}
  FileReader.prototype.readAsArrayBuffer = function (file) { if (file.unreadable) { this.onerror(); return; } this.result = file.buffer; this.onload(); };
  const ctx = {
    document: { getElementById: (id) => { if (!byId[id]) throw new Error('no element #' + id); return byId[id]; }, createElement: element },
    google: { script: { run: runner, host: { close: () => { closed++; } } } },
    FileReader: FileReader, btoa: (binary) => Buffer.from(binary, 'binary').toString('base64'), Uint8Array: Uint8Array
  };
  vm.createContext(ctx);
  vm.runInContext(scripts[0].replace(/^<script>|<\/script>$/g, ''), ctx);
  const file = (text, extra) => { const b = FX.utf16le(text); return Object.assign({ size: b.length, buffer: b.buffer.slice(b.byteOffset, b.byteOffset + b.length) }, extra || {}); };
  return {
    ctx: ctx, el: byId, calls: calls, scriptCount: scripts.length, closed: () => closed, file: file,
    choose: (text, extra) => { byId.file.files = [file(text, extra)]; ctx.fileChanged(); },
    select: (id) => { byId.session.value = id; ctx.sessionChanged(); },
    previewRows: () => byId.preview.children.map((tr) => tr.children.map((td) => td.textContent)),
    visible: (id) => byId[id].style.display !== 'none'
  };
}

console.log('9. the dialog');
{
  const r = report();
  r.import('s2', base64(smallExport('2026-09-23')));
  const props0 = r.propsSnapshot();
  const body0 = r.snapshot();
  r.s.importTeamsAttendance();
  check('one dialog, titled Import Teams Attendance; opening it changes nothing', [r.ui.dialogs.length, r.ui.dialogs[0].title, r.ui.alerts.length, r.propsSnapshot() === props0, r.snapshot() === body0], [1, 'Import Teams Attendance', 0, true, true]);
  check('what the dialog is given: the sessions and what is imported, nothing about any participant', r.s.buildTeamsAttendanceDialogModel_(), { notice: '', readOnly: false, sessions: [
    { id: 's1', label: 'September 22, 2026', imported: '', records: [] },
    { id: 's2', label: 'Offline, September 23, 2026', imported: '4 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"',
      records: [{ day: '2026-09-23', label: 'September 23, 2026', text: '4 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"' }] },
    { id: 's3', label: 'September 24, 2026', imported: '', records: [] }] });

  const d = runDialog(r.ui.dialogs[0].html, r.s);
  check('one script; the sessions in order, the imported one marked', [d.scriptCount, d.el.session.children.map((o) => [o.value, o.textContent])],
    [1, [['s1', 'September 22, 2026'], ['s2', 'Offline, September 23, 2026 (attendance imported)'], ['s3', 'September 24, 2026']]]);
  check('at first: the first session, no file, Preview disabled, no Import button, no Remove', [d.el.session.value, d.el.previewBtn.disabled, d.visible('importBtn'), d.visible('removeRow'), d.el.existing.textContent],
    ['s1', true, false, false, 'No attendance is imported for this session yet.']);
  check('nothing was sent by opening', d.calls, []);
  const onclick = (id) => (r.ui.dialogs[0].html.match(new RegExp('<button[^>]*id="' + id + '"[^>]*onclick="([^"]*)"')) || [])[1];
  check('what each button does: Remove only asks, and only "Yes, remove" removes (source)', ['previewBtn', 'importBtn', 'removeBtn', 'removeYes', 'removeNo', 'closeBtn'].map(onclick),
    ['previewImport()', 'confirmImport()', 'askRemove()', 'removeAttendance()', 'cancelRemove()', 'google.script.host.close()']);
  check('the Import button and the Remove question are hidden until they apply (source)', [/id="importBtn"[^>]*style="display:none"/.test(r.ui.dialogs[0].html), /id="removeConfirm" style="display:none"/.test(r.ui.dialogs[0].html), /id="previewBtn"[^>]*disabled/.test(r.ui.dialogs[0].html)], [true, true, true]);

  // Choose a file and preview.
  d.choose(FX.buildExport());
  check('with a file: Preview is enabled, still nothing sent', [d.el.previewBtn.disabled, d.calls.length], [false, 0]);
  d.ctx.previewImport();
  check('Preview sends the session and the file\'s bytes, base64-encoded', [d.calls.length, d.calls[0][0], d.calls[0][1], d.calls[0][2] === FULL_22], [1, 'previewTeamsAttendanceImport', 's1', true]);
  check('the preview: one row per value', d.previewRows(), [['Session', 'September 22, 2026'], ['Teams meeting', 'Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE'], ['Date', 'September 22, 2026'],
    ['Start', '14:48:05'], ['End', '18:11:26'], ['Duration', '3:23:20'], ['Attendance records (Teams)', '25'], ['Attendees in the table', '23'], ['Average attendance', '2:31:07'], ['Warnings', '3']]);
  check('the warnings are shown, one per line; the other notes below', [d.visible('warnings'), d.el.warnings.textContent.split('\n').length, d.el.infoNotes.textContent.split('\n').length], [true, 3, 6]);
  check('the Import button appears only now', [d.visible('importBtn'), d.el.importBtn.textContent, d.visible('replaces')], [true, 'Import', false]);
  check('still nothing is stored or written', [r.propsSnapshot() === props0, r.snapshot() === body0], [true, true]);

  // Changing the file or the session withdraws the preview.
  d.choose(smallExport('2026-09-22'));
  check('choosing another file withdraws the preview and the Import button', [d.previewRows(), d.visible('importBtn'), d.visible('warnings')], [[], false, false]);
  d.ctx.confirmImport();
  check('Import without a current preview sends nothing', d.calls.length, 1);
  d.choose(FX.buildExport());
  d.ctx.previewImport();
  d.select('s3');
  check('choosing another session withdraws it too', [d.previewRows(), d.visible('importBtn')], [[], false]);
  d.select('s1');

  // Cancel.
  d.ctx.previewImport();
  d.ctx.google.script.host.close();
  check('Cancel after a preview: nothing stored, nothing written', [d.closed(), r.propsSnapshot() === props0, r.snapshot() === body0, d.calls.filter((c) => c[0] !== 'previewTeamsAttendanceImport')], [1, true, true, []]);
  check('the Cancel button only closes the dialog (source)', (r.ui.dialogs[0].html.match(/<button[^>]*id="closeBtn"[^>]*>/) || [''])[0], '<button type="button" class="grey" id="closeBtn" onclick="google.script.host.close()">');

  // Confirm.
  d.ctx.confirmImport();
  const confirmCall = d.calls[d.calls.length - 1];
  check('Import sends the same session, the same bytes and the preview\'s token', [confirmCall[0], confirmCall[1], confirmCall[2] === FULL_22, typeof confirmCall[3], confirmCall[3].length], ['confirmTeamsAttendanceImport', 's1', true, 'string', 8]);
  check('the result is shown; the dialog is finished: everything but Close is disabled', [d.el.status.textContent, d.el.status.className, d.el.closeBtn.textContent, d.visible('importBtn'), d.el.previewBtn.disabled, d.el.session.disabled, d.el.file.disabled],
    ['The attendance of 22 Sep was imported (23 attendees).', 'done', 'Close', false, true, true, true]);
  check('and the report has it', [r.s.adhocAttendanceSessionIds_(r.docProps).sort(), attendeeTables(r.body).length], [['s1', 's2'], 2]);

  // A second dialog: unchanged, replacement, wrong file, remove.
  r.s.importTeamsAttendance();
  const e = runDialog(r.ui.dialogs[1].html, r.s);
  check('a session with attendance: what is imported, and the Remove button', [e.el.existing.textContent, e.visible('removeRow'), e.visible('removeConfirm')], ['Imported: 23 attendees, "Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE"', true, false]);
  e.choose(FX.buildExport());
  e.ctx.previewImport();
  check('the same file again: it says so, and there is no Import button', [e.el.status.textContent, e.visible('importBtn')], ['This attendance is already imported for this session. Nothing needs to be done.', false]);
  e.choose(smallExport('2026-09-22'));
  e.ctx.previewImport();
  check('another file: the replacement is announced above the preview, and the button says it', [e.visible('replaces'), /^This replaces the attendance already imported for this session \(23 attendees/.test(e.el.replaces.textContent), e.el.importBtn.textContent], [true, true, 'Replace the imported attendance']);
  e.select('s3');
  e.choose(FX.buildExport());
  e.ctx.previewImport();
  check('the wrong file for the session: the error is shown, no preview, no Import button, and no way round it', [e.el.status.textContent, e.el.status.className, e.previewRows(), e.visible('importBtn')],
    ['This attendance export is for 2026-09-22, not for the session on 2026-09-24.', 'error', [], false]);
  check('nothing in the dialog offers to import anyway (source)', /anyway|override|ignore/i.test(r.ui.dialogs[1].html), false);
  e.choose('not an export');
  e.ctx.previewImport();
  check('a file that is not a Teams export', [/^This is not a Teams attendance export/.test(e.el.status.textContent), e.visible('importBtn')], [true, false]);
  e.choose(FX.buildExport(), { size: 3000001 });
  const sent = e.calls.length;
  e.ctx.previewImport();
  check('a file that is too large is not even sent', [e.el.status.textContent, e.calls.length === sent], ['The file is too large to be a Teams attendance export.', true]);
  e.choose(FX.buildExport(), { unreadable: true });
  e.ctx.previewImport();
  check('a file the browser cannot read', [e.el.status.textContent, e.calls.length === sent, e.el.previewBtn.disabled], ['The file could not be read.', true, false]);

  // Remove: two clicks.
  e.select('s2');
  const beforeRemove = r.propsSnapshot();
  e.ctx.askRemove();
  check('Remove asks first; nothing is sent', [e.visible('removeConfirm'), e.visible('removeBtn'), e.el.removeQuestion.textContent, e.calls.filter((c) => c[0] === 'removeTeamsAttendanceImport').length, r.propsSnapshot() === beforeRemove],
    [true, false, 'Remove the imported attendance of Offline, September 23, 2026 from the report?', 0, true]);
  e.ctx.cancelRemove();
  check('No: nothing is removed', [e.visible('removeConfirm'), e.visible('removeBtn'), r.propsSnapshot() === beforeRemove], [false, true, true]);
  e.ctx.askRemove();
  e.select('s1');
  check('choosing another session withdraws the question', e.visible('removeConfirm'), false);
  e.select('s2');
  e.ctx.askRemove();
  e.ctx.removeAttendance();
  check('Yes: that session\'s attendance is removed', [e.calls[e.calls.length - 1], e.el.status.textContent, r.s.adhocAttendanceSessionIds_(r.docProps)], [['removeTeamsAttendanceImport', 's2'], 'The attendance of Offline was removed.', ['s1']]);

  // A failed request, and a result that was stored but not rendered.
  r.s.importTeamsAttendance();
  const f = runDialog(r.ui.dialogs[2].html, null);
  f.choose(FX.buildExport());
  f.ctx.previewImport();
  f.calls.handlers.failure(new Error('Server error'));
  check('a failed request: its message, and the dialog can be used again', [f.el.status.textContent, f.el.status.className, f.el.previewBtn.disabled], ['Server error', 'error', false]);
  f.ctx.previewImport();
  f.calls.handlers.success({ ok: true, token: 'abc', preview: { session: 'x', meetingTitle: 'y', date: 'z', start: '1', end: '', duration: '', attendanceRecords: null, attendees: 0, averageAttendance: '', warnings: [], notes: [], replaces: '', unchanged: false } });
  check('values that are missing are not shown as rows; a count of 0 is', f.previewRows().map((row) => row[0]), ['Session', 'Teams meeting', 'Date', 'Start', 'Attendees in the table', 'Warnings']);
  f.ctx.confirmImport();
  f.calls.handlers.success({ ok: true, changed: true, rendered: false, message: 'Stored, but the section could not be written.' });
  check('stored but not rendered: shown as a problem', [f.el.status.textContent, f.el.status.className], ['Stored, but the section could not be written.', 'error']);

  // Untrusted text stays text.
  const hostile = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-09-22', '', '', '</script><b>x')]) }) });
  hostile.import('s1', base64(smallExport('2026-09-22', [person('<img src=x onerror=alert(1)>', 'a@example.com')]).replace('Synthetic AHG Meeting ', '<script>alert(1)</script> ')));
  hostile.s.importTeamsAttendance();
  const h = runDialog(hostile.ui.dialogs[0].html, hostile.s);
  check('a label and a meeting title with markup: one script element, shown as typed', [h.scriptCount, h.el.session.children[0].textContent, hostile.ui.dialogs[0].html.indexOf('</script><b>'), hostile.ui.dialogs[0].html.indexOf('<script>alert')],
    [1, '</script><b>x, September 22, 2026 (attendance imported)', -1, -1]);
  check('the dialog script builds everything with textContent, never with innerHTML (source)', [/innerHTML|insertAdjacentHTML|document\.write/.test(hostile.ui.dialogs[0].html), /textContent/.test(hostile.ui.dialogs[0].html)], [false, true]);
  check('the name with markup is in the report as text', tableRows(attendeeTables(hostile.body)[0])[1], ['<img src=x onerror=alert(1)>', '', 'a@example.com']);

  // Where the dialog does not open.
  const main = report({ props: Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }) });
  main.s.importTeamsAttendance();
  check('a main-meeting report: a message, no dialog', [main.ui.dialogs.length, main.ui.alerts.map((a) => a[1])], [0, ['Attendance is available for ad-hoc reports only.']]);
  const noSessions = report({ props: ADHOC });
  noSessions.s.importTeamsAttendance();
  check('an ad-hoc report without sessions: no import dialog, and where to configure them',
    [noSessions.ui.dialogs.length, noSessions.ui.alerts.map((a) => a[1]), noSessions.writes], [0, ['Attendance is imported for a session. Configure the sessions of this report first:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…'], []]);
  const newer = report();
  newer.docProps.setProperty(KEY, JSON.stringify({ v: 2, gen: 1, entries: {} }));
  newer.s.importTeamsAttendance();
  const n = runDialog(newer.ui.dialogs[0].html, newer.s);
  n.choose(FX.buildExport());
  check('stored attendance of a newer version: a notice, and Preview stays disabled', [/newer|version 2/.test(n.el.notice.textContent), n.visible('notice'), n.el.previewBtn.disabled], [true, true, true]);
}

// ================================================================ 10. menu and activation

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
  const all = JSON.stringify(ui.menus);
  const sub = ui.menus[0].entries.filter((e) => e.sub && /Sessions/.test(e.sub.name)).map((e) => e.sub)[0];
  return { loaded: loaded, sub: sub, mentionsAttendance: /Attendance/.test(all), docProps: loaded.docProps };
}

console.log('10. menu and activation');
{
  const adhoc = openMenu(Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, WITH_SESSIONS));
  check('ad-hoc report: the submenu; its first three items (Assign TDoc Sessions… of stage B follows)', adhoc.sub.entries.slice(0, 3), [{ label: 'Configure Sessions…', fn: 'configureAdhocSessions' }, { label: 'Import Teams Attendance…', fn: 'importTeamsAttendance' }, { label: 'Refresh Attendance Section', fn: 'refreshAdhocAttendanceSection' }]);
  check('every item calls an existing public function', adhoc.sub.entries.filter((e) => typeof adhoc.loaded.sandbox[e.fn] !== 'function' || /_$/.test(e.fn)), []);
  check('main-meeting report, new report, master template: nothing about attendance in the menu',
    [openMenu(Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' })).mentionsAttendance, openMenu({}).mentionsAttendance, openMenu(WITH_SESSIONS, { docId: TEMPLATE_ID }).mentionsAttendance], [false, false, false]);
  check('CENTRAL / Legacy menu: nothing about attendance, with or without attendance stored', [openMenu(WITH_SESSIONS, { release: false }).mentionsAttendance, openMenu(Object.assign({ [KEY]: '{"v":1,"gen":0,"entries":{}}' }, WITH_SESSIONS), { release: false }).mentionsAttendance], [false, false]);
  check('opening a report reads but never writes', [openMenu(WITH_SESSIONS).docProps._store, openMenu(ADHOC).docProps._store], [WITH_SESSIONS, ADHOC]);
  const onOpenSource = CODE.slice(CODE.indexOf('\nfunction onOpen('), CODE.indexOf('\n}\n', CODE.indexOf('\nfunction onOpen(')));
  check('onOpen() in Code.js does not mention attendance (source)', /ttendance|ADHOC_/.test(onOpenSource), false);

  // An ad-hoc report without sessions: the items are there, and lead to Configure Sessions.
  const noSessions = report({ props: ADHOC, template: true });
  noSessions.s.importTeamsAttendance();
  check('ad-hoc without sessions: Import explains instead of opening an import that cannot work', [openMenu(ADHOC).sub.entries.length, noSessions.ui.dialogs.length, /Configure the sessions of this report first/.test(noSessions.ui.alerts[0][1])], [6, 0, true]);
  check('and Refresh is an error there', /Configure Sessions first/.test(thrown(() => noSessions.s.refreshAdhocAttendanceSection()) || ''), true);
  check('sessions without any session in them count as no sessions', [report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([], 3) }) }).s.previewTeamsAttendanceImport('s1', FULL_22).error], ['Attendance needs configured sessions. Use Configure Sessions first.']);
}

// ================================================================ 11. Build from Scratch

const FXDIR = path.join(__dirname, 'fixtures');
const LIST_85916 = JSON.parse(fs.readFileSync(path.join(FXDIR, 'meeting-85916-tdoc-list-values.json'), 'utf8')).values;
const FTP_AUDIO = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const PROPS_85916 = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio',
  FTP_BASE: FTP_AUDIO, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const AUDIO_AGENDA = (close) => [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'ULBC-MED'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints']]
  .concat(close === false ? [] : [['5', 'Close of the meeting']]).map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));

/** A report whose real buildSkeletonWithTdocTables() can be run; only network, Drive, Sheets and the template are faked. */
function buildable(props, opts) {
  const o = opts || {};
  const r = report({ props: props, template: o.template });
  const s = r.s;
  r.body.clear();
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  r.ui.alert = (...args) => { r.ui.alerts.push(args); return 'YES'; };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }) };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => LIST_85916, getRichTextValues: () => LIST_85916.map((row) => row.map(() => null)) }) }] }) };
  s.downloadMeetingAgenda_ = () => AUDIO_AGENDA(o.close);
  r.build = () => { try { return s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }); } catch (e) { return { ok: false, message: e.message }; } };
  return r;
}
const AUDIO_SESSIONS = Object.assign({}, PROPS_85916, { ADHOC_SESSIONS: sessionsProperty(THREE_SESSIONS) });

console.log('11. Build from Scratch');
{
  const r = buildable(AUDIO_SESSIONS);
  const first = r.build();
  check('a build of an ad-hoc report with sessions and no attendance: no Attendance section', [first.ok, texts(r.body).indexOf('Attendance'), attendanceKeys(r)], [true, -1, []]);
  const withoutAttendance = r.snapshot();

  r.import('s1', FULL_22);
  r.import('s2', base64(smallExport('2026-09-23')));
  check('the import puts the section at the end of the closing section of the built report', [texts(r.body)[texts(r.body).indexOf('Attendance') - 1], texts(r.body).filter((t) => t === 'Attendance').length], ['5 Close of the meeting', 1]);
  typeCompany(r, 0, 'Alex Organizer', 'Example Chair Corp');
  typeCompany(r, 1, 'Casey Guest', 'Guest Company');
  r.body.insertParagraph(r.s.findAdhocAttendanceContainer_(r.body).start, 'The chair closes the meeting. (typed by hand)');
  const stored = r.propsSnapshot();

  const rebuilt = r.build();
  check('the rebuild completes without a note', [rebuilt.ok, rebuilt.note], [true, '']);
  check('the attendance is still stored, unchanged', [r.s.readAdhocAttendance_(r.docProps).raw.s1 === JSON.parse(stored)[KEY + '_s1_1_0'], r.s.adhocAttendanceSessionIds_(r.docProps)], [true, ['s1', 's2']]);
  check('the section is written again, once, at the end of the closing section, both sessions in order',
    [texts(r.body).filter((t) => t === 'Attendance').length, texts(r.body)[texts(r.body).indexOf('Attendance') - 1], containerTexts(r).filter((t) => /^Session|table/.test(t)), texts(r.body)[texts(r.body).length - 1]],
    [1, '5 Close of the meeting', ['Session: September 22, 2026', '[table 23]', 'Session: Offline, September 23, 2026', '[table 4]'], '[table 4]']);
  check('the Company values typed before the rebuild are back, in every session', [companyOf(r, 0, 'Alex Organizer'), companyOf(r, 0, 'Casey Guest'), companyOf(r, 1, 'Alex Organizer'), companyOf(r, 1, 'Casey Guest')],
    ['Example Chair Corp', 'Guest Company', 'Example Chair Corp', 'Guest Company']);
  check('prose typed by hand is lost, as the build says', texts(r.body).filter((t) => /typed by hand/.test(t)), []);
  const afterRebuild = r.snapshot();
  r.build();
  check('a second rebuild gives the same document', r.snapshot() === afterRebuild, true);
  const c = r.s.findAdhocAttendanceContainer_(r.body);
  check('apart from the container, the rebuilt report is the report built without attendance',
    JSON.stringify(elements(r.body).slice(0, c.start).concat(elements(r.body).slice(c.end))) === withoutAttendance, true);

  // The reallocation table still survives, next to the attendance.
  r.s.saveReallocation('S4aA260092', '4.4', '4.3', 'test');
  const withReallocation = r.build();
  const reallocation = r.body._children.filter((x) => isTable(x) && x.getRow(0).getCell(1).getText() === 'Original Agenda')[0];
  check('Document Reallocations are kept as before, and the attendance with them', [withReallocation.ok, withReallocation.note, tableRows(reallocation).slice(1), attendeeTables(r.body).length], [true, '', [['S4aA260092', '4.4', '4.3', 'test']], 2]);

  // An agenda without a closing item.
  const noClose = buildable(AUDIO_SESSIONS, { close: false });
  noClose.build();
  noClose.import('s1', FULL_22);
  noClose.build();
  check('an agenda without a closing item: the section is under the closing heading the build adds', [texts(noClose.body).filter((t) => t === 'Attendance').length, /^\d+(\.\d+)* Close of the session$/.test(texts(noClose.body)[texts(noClose.body).indexOf('Attendance') - 1])], [1, true]);

  // Problems during a rebuild do not fail the build.
  const failing = buildable(AUDIO_SESSIONS);
  failing.build();
  failing.import('s1', FULL_22);
  failing.s.renderAdhocAttendanceSection_ = () => { throw new Error('Service Documents failed'); };
  const failed = failing.build();
  check('the section cannot be written: the report is built, the note says how to write it, the attendance is still stored',
    [failed.ok, failed.note, failing.s.adhocAttendanceSessionIds_(failing.docProps), texts(failing.body).indexOf('5 Close of the meeting') !== -1],
    [true, '\n\n⚠️ The Attendance section could not be written (Service Documents failed). Use SA4 Report > Sessions and Attendance > Refresh Attendance Section.', ['s1'], true]);
  const damaged = buildable(AUDIO_SESSIONS);
  damaged.build();
  damaged.import('s1', FULL_22);
  damaged.import('s2', base64(smallExport('2026-09-23')));
  damaged.docProps.deleteProperty(KEY + '_s1_1_0');
  const partly = damaged.build();
  check('one session\'s stored attendance is damaged: the other is written, and the note names the damaged one',
    [partly.ok, partly.note, containerTexts(damaged).filter((t) => /^Session/.test(t))], [true, '\n\n⚠️ The stored attendance of 22 Sep cannot be read and was left out. Import it again.', ['Session: Offline, September 23, 2026']]);
  const refusedBuild = buildable(Object.assign({}, AUDIO_SESSIONS, { TDOC_LIST_URL: '' }));
  refusedBuild.import('s1', FULL_22);
  const beforeRefusal = refusedBuild.snapshot();
  check('a build that is refused before it starts leaves the report and the attendance alone', [refusedBuild.build().ok, refusedBuild.snapshot() === beforeRefusal, refusedBuild.s.adhocAttendanceSessionIds_(refusedBuild.docProps)], [false, true, ['s1']]);

  // Where the hooks do nothing.
  const strayAttendance = (() => { const x = report(); x.import('s1', FULL_22); const out = {}; attendanceKeys(x).forEach((k) => { out[k] = x.docProps.getProperty(k); }); return out; })();
  const plain = buildable(PROPS_85916);
  plain.build();
  const noSessions = buildable(Object.assign({}, PROPS_85916, strayAttendance));
  noSessions.build();
  check('ad-hoc report without sessions (even with attendance properties left in it): the build is the build of a report without them',
    [noSessions.snapshot() === plain.snapshot(), texts(noSessions.body).indexOf('Attendance'), JSON.stringify(strayAttendance) === JSON.stringify(Object.keys(strayAttendance).reduce((o, k) => { o[k] = noSessions.docProps.getProperty(k); return o; }, {}))], [true, -1, true]);
  // Stage B: with sessions the registration table has a fifth column, Session. Without that column the documents are the same.
  const withoutSessionColumn = (snapshot) => JSON.stringify(JSON.parse(snapshot).map((e) => (e[0] === 'T' && e[1][0].length === 5 && e[1][0][4] === 'Session' ? ['T', e[1].map((row) => row.slice(0, 4))] : e)));
  // Stage E: with sessions the opening part also has the generated Online information block and Session administration section.
  const withoutOpening = (snapshot) => { const out = []; let inside = false; JSON.parse(snapshot).forEach((e) => { if (e[0] === 'P' && (e[1] === 'Session administration' || e[1] === 'Online information') && e[2] !== 'NORMAL') { inside = true; return; } if (inside && e[0] === 'P' && e[2] === 'NORMAL') return; inside = false; out.push(e); }); return JSON.stringify(out); };
  // E1: and the "<Chair> opens the session ..." line of the build is not written beside that section.
  const withoutLegacyLine = (snapshot) => JSON.stringify(JSON.parse(snapshot).filter((e) => !(e[0] === 'P' && /^<Chair> opens the session on /.test(e[1]))));
  check('ad-hoc report with sessions and no attendance: apart from the Session column and the Session administration section (in place of the "<Chair> opens ..." line), the same document as without sessions',
    [withoutAttendance === plain.snapshot(), withoutOpening(withoutSessionColumn(withoutAttendance)) === withoutLegacyLine(plain.snapshot()), withoutLegacyLine(plain.snapshot()) === plain.snapshot()], [false, true, false]);
  const reads = [];
  const quiet = buildable(PROPS_85916);
  const get = quiet.docProps.getProperty;
  quiet.docProps.getProperty = (k) => { reads.push(k); return get(k); };
  quiet.build();
  check('a report without sessions never reads an attendance property during a build', reads.filter((k) => k.indexOf(KEY) === 0), []);
  const mainReads = [];
  const main = loadCode({ documentProperties: Object.assign({ MEETING_TYPE: 'main' }, strayAttendance, { ADHOC_SESSIONS: sessionsProperty(THREE_SESSIONS) }) });
  const mainGet = main.docProps.getProperty;
  main.docProps.getProperty = (k) => { mainReads.push(k); return mainGet(k); };
  check('a main-meeting report: both hooks do nothing, and read nothing but the meeting type', [main.sandbox.beginAdhocAttendanceRebuild_(), main.sandbox.finishAdhocAttendanceRebuild_({}, null), mainReads], [null, '', ['MEETING_TYPE']]);
  check('the build calls each hook once: one before the document is cleared, one after the reallocations are restored (source)', (() => {
    const src = CODE.slice(CODE.indexOf('\nfunction buildSkeletonWithTdocTables('), CODE.indexOf('\nfunction downloadAndGroupTdocs_('));
    return [(src.match(/beginAdhocAttendanceRebuild_\(\)/g) || []).length, (src.match(/finishAdhocAttendanceRebuild_\(/g) || []).length,
      src.indexOf('beginAdhocAttendanceRebuild_()') < src.indexOf('getActiveDocument().getBody().clear()'), src.indexOf('finishAdhocAttendanceRebuild_(') > src.indexOf('saveReallocation(tdoc, r.original, r.new, r.reason)')];
  })(), [1, 1, true, true]);

  // The question before the build.
  const question = (props, attendance) => { const x = report({ props: Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, props), template: true }); if (attendance) x.import('s1', FULL_22); x.ui.alert = (...args) => { x.ui.alerts.push(args); return 'NO'; }; x.s.runFullReportBuild(); return x.ui.alerts[x.ui.alerts.length - 1][1]; };
  check('the build question of a report with attendance says that it is kept', /Only the Document Reallocations table and the imported attendance \(with its Company cells\) are kept\./.test(question(WITH_SESSIONS, true)), true);
  check('without attendance the question is the one of T-2026.10.4', [question(WITH_SESSIONS, false), question(ADHOC, false), question(Object.assign({}, ADHOC, { MEETING_TYPE: 'main' }), false)].map((t) => /Only the Document Reallocations table is kept\./.test(t) && !/attendance/.test(t)), [true, true, true]);
}

// ================================================================ 12. adoption

console.log('12. adoption');
{
  const r = report();
  r.import('s1', FULL_22);
  typeCompany(r, 0, 'Alex Organizer', 'Example Chair Corp');
  r.s.refreshAdhocAttendanceSection();
  const stored = attendanceKeys(r);
  const adopted = r.s.adoptReportDocumentForAddon_();
  check('there is attendance to adopt, in several properties', stored.length, 3);
  check('adoption copies the sessions and none of the attendance: background updates never read it', [adopted.verified, adopted.copiedKeys.filter((k) => /ADHOC/.test(k)), r.scriptProps.getKeys().filter((k) => k.indexOf('ATTENDANCE') !== -1)], [true, ['ADHOC_SESSIONS'], []]);
  check('the attendance stays in the document, untouched', attendanceKeys(r), stored);
  check('no attendance key is an adoption key or an adoption prefix', [r.s.ADDON003_ADOPTION_FIXED_KEYS_.filter((k) => /ATTENDANCE/.test(k)), r.s.ADDON003_ADOPTION_PREFIX_FAMILIES_, r.s.ADDON003_ADOPTION_PREFIX_FAMILIES_.filter((p) => KEY.indexOf(p) === 0 || 'ADHOC_SESSIONS'.indexOf(p) === 0)],
    [[], ['DISCUSS_', 'REVIS_', 'A1_EMPTY_CACHE_', 'REVIEWER_NO_SUMMARY_CACHE_'], []]);
  const plain = report({ props: ADHOC });
  check('adoption of a document without attendance copies what it copied before', plain.s.adoptReportDocumentForAddon_().copiedKeys, ['MEETING_ID', 'REPORT_SUFFIX', 'MEETING_DATE', 'MEETING_TYPE', 'MEETING_NAME']);
  check('adoption invents no attendance property anywhere', [attendanceKeys(plain), plain.scriptProps.getKeys().filter((k) => /ATTENDANCE/.test(k))], [[], []]);
  check('attendance is read from and written to the document\'s own properties, whatever the state store is (source)',
    [/function adhocAttendanceStore_\(\) \{\n  return PropertiesService\.getDocumentProperties\(\);\n\}/.test(CODE), (CODE.slice(CODE.indexOf('// AD-HOC ATTENDANCE (stage D)'), CODE.indexOf('// AD-HOC TDOC SESSIONS (stage B)')).match(/getReportStateStore_|getScriptProperties/g) || [])], [true, []]);
  check('Clear Collection Caches does not touch attendance', (() => { const x = report(); x.import('s1', FULL_22); const keys = attendanceKeys(x); x.s.clearAllCaches(); return [attendanceKeys(x).length === keys.length, keys.length]; })(), [true, 2]);
}

// ================================================================ 13. privacy, and nothing else changed

console.log('13. privacy, and nothing else changed');
{
  const r = report();
  r.import('s1', FULL_22);
  typeCompany(r, 0, 'Alex Organizer', 'Example Chair Corp');
  r.s.refreshAdhocAttendanceSection();
  r.s.importTeamsAttendance();
  const everything = r.propsSnapshot() + JSON.stringify(r.scriptProps._store) + r.snapshot() + JSON.stringify(r.ui.dialogs) + JSON.stringify(r.logs);
  check('the export does contain Participant IDs and the dial-in number', FX.PRIVATE_STRINGS.map((x) => FX.buildExport().indexOf(x) !== -1), [true, true, true]);
  check('none of them is in any property, in the report, in a dialog or in a log', FX.PRIVATE_STRINGS.filter((x) => everything.indexOf(x) !== -1), []);
  const storedText = attendanceKeys(r).map((k) => r.docProps.getProperty(k)).join('\n');
  check('stored: no join or leave time of a participant, no duration of one, no role, no marker, no diagnostics, no file',
    [/15:07:20|17:41:02|9222|firstJoin|lastLeave|durationSeconds/.test(storedText), /Presenter|\(External\)|\(Unverified\)|"Attendee"/.test(storedText), /notes|warning|DUPLICATE|rows|diagnostics|calculated/.test(storedText), /Participant ID|In-Meeting|Engagement|1\. Summary|\t/.test(storedText)],
    [false, false, false, false]);
  check('stored: exactly these properties', attendanceKeys(r), [KEY, KEY + '_companies_2_0', KEY + '_s1_1_0']);
  check('nothing was logged at all by a normal import, refresh and dialog', r.logs, []);
  check('the attendance code logs only fixed texts and error messages, never a value of a participant (source)', (() => {
    const section = CODE.slice(CODE.indexOf('// AD-HOC ATTENDANCE (stage D)'), CODE.indexOf('// AD-HOC TDOC SESSIONS (stage B)'));
    return (section.match(/Logger\.log\((.*)\);/g) || []).filter((line) => !/^Logger\.log\('Attendance: [^']*' \+ e\.message\);$/.test(line));
  })(), []);

  // The update path never looks at attendance.
  const u = report();
  u.import('s1', FULL_22);
  const propsBefore = u.propsSnapshot();
  const bodyBefore = u.snapshot();
  const events = [];
  const reads = [];
  const get = u.docProps.getProperty;
  u.docProps.getProperty = (k) => { reads.push(k); return get(k); };
  u.s.assertMeetingReadyToBuild_ = () => {};
  u.s.downloadAndGroupTdocs_ = () => { events.push('TDoc list'); return {}; };
  u.s.rearrangeRevisionTables_ = () => { events.push('revision placement'); return { moved: 0, dispositions: 0 }; };
  u.s.checkRSSFeed_ = () => { events.push('e-mail'); };
  u.s.updateRevisions_ = () => { events.push('revisions'); };
  u.s.removeRowHeightAndSpacing = () => { events.push('formatting'); };
  const result = u.s.continuousUpdateCore_();
  check('a complete update of a report with attendance: the usual steps, the report and the attendance untouched',
    [result, events, u.snapshot() === bodyBefore, u.propsSnapshot() === propsBefore], [{ success: true, error: null }, ['TDoc list', 'revision placement', 'e-mail', 'revisions'], true, true]);
  check('and it reads no attendance property', reads.filter((k) => k.indexOf(KEY) === 0), []);
  check('the attendee table is not taken for a TDoc table, a configuration table or the registration table',
    [u.s.isTDocTable_(attendeeTables(u.body)[0]), u.s.isConfigTable_(attendeeTables(u.body)[0]), u.s.isReallocationTable_(attendeeTables(u.body)[0]), u.s.isDeadlineExtensionTable_(attendeeTables(u.body)[0])], [false, false, false, false]);

  // The functions outside the feature.
  const section = CODE.slice(CODE.indexOf('// AD-HOC SESSIONS (stage A)'), CODE.indexOf('// ARCH-009 -- MEETING-ID RESOLVER CORE'));
  const outside = CODE.replace(section, '');
  check('outside the ad-hoc sections, Code.js mentions attendance in the two build hooks only',
    (outside.match(/[A-Za-z_]*[Aa]ttendance[A-Za-z_]*/g) || []).filter((w, i, list) => list.indexOf(w) === i).sort(), ['attendance', 'Attendance', 'attendanceRebuild', 'beginAdhocAttendanceRebuild_', 'finishAdhocAttendanceRebuild_'].sort());
  check('continuousUpdateCore_(), the collector and the registration table do not mention attendance (source)',
    ['continuousUpdateCore_', 'collectorUpdate_', 'updateRegisteredDocumentsTable_', 'createSummaryTable_', 'insertNewTdoc_', 'runFullReportBuildCore_'].filter((name) => {
      const at = CODE.indexOf('\nfunction ' + name + '(');
      return /ttendance/.test(CODE.slice(at, CODE.indexOf('\n}\n', at)));
    }), []);
  check('Code.js is version 2.21.1 (2.18.0 added the ad-hoc sessions, 2.18.1 the TDoc upload completion, 2.19.0 the status dropdowns, 2.20.0 their conversion in an existing report, 2.21.0 the ad-hoc meeting automation and the personal Reviewer token, 2.21.1 the dropdown status in a discussion e-mail)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.21.1');
  check('ReportCreator.js mentions attendance only in the menu and in the build question', (CREATOR.match(/^.*ttendance.*$/gm) || []).length, 6);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll ad-hoc attendance checks passed.');
process.exitCode = failures ? 1 : 0;
