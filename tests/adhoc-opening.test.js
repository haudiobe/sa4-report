/**
 * Ad-hoc sessions, stage E -- the opening details of the sessions (Chair,
 * minute takers, an administrative note) and the generated "Session
 * administration" section.
 *
 *   1. The stored details: the model and the property.
 *   2. What is rendered (pure).
 *   3. Rendering into the report: placement, replacement, what is left alone.
 *   4. The dialog.
 *   5. Sessions that change; sessions with details are not removed.
 *   6. Build from Scratch.
 *   7. Updates.
 *   8. Failure and recovery.
 *   9. Menu, adoption, privacy, and nothing else changed.
 *
 * All names and notes are synthetic.
 *
 * Run: node tests/adhoc-opening.test.js
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
const KEY = 'ADHOC_SESSION_OPENING';
const HEADING = 'Session administration';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.4', flavor: 'template', codeVersion: '0.0.0', gitCommit: 'abcdef0', gitTag: 'template-release/T-2026.10.4', templateDocumentId: TEMPLATE_ID };

const session = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const sessionsProperty = (list, nextId) => JSON.stringify({ v: 1, nextId: nextId || list.length + 1, sessions: list });
const S1 = session('s1', '2026-09-22', '15:00', '18:00');
const S2 = session('s2', '2026-09-23', '', '', 'Offline');
const S3 = session('s3', '2026-09-24', '15:00');
const SESSIONS = [S1, S2, S3];
const ADHOC = { MEETING_TYPE: 'adhoc', MEETING_ID: '86178', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'September 22, 2026', REPORT_SUFFIX: '6G' };
const WITH_SESSIONS = Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty(SESSIONS) });
const D = (chair, minuteTakers, note) => ({ chair: chair || '', minuteTakers: minuteTakers || '', note: note || '' });
const stored = (sessions) => JSON.stringify({ v: 1, sessions: sessions });
const DETAILS = { s1: D('Alex Organizer', 'Sam Rivera, Casey Guest', 'The agenda was approved.\nThe IPR reminder was read.'), s2: D('', 'Sam Rivera'), s3: D() };

/**
 * One report in a sandbox, with a body shaped like a built ad-hoc report.
 * opts: { props, template, docId, built (false: an empty document), registration (false: no "Registration of Documents" heading) }.
 */
function report(opts) {
  const o = opts || {};
  const loaded = o.template ? loadTemplateRuntime({ release: RELEASE, documentProperties: o.props || WITH_SESSIONS }) : loadCode({ documentProperties: o.props || WITH_SESSIONS });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const H2 = s.DocumentApp.ParagraphHeading.HEADING2;
  const H3 = s.DocumentApp.ParagraphHeading.HEADING3;
  if (o.built !== false) {
    body.appendParagraph('5.1 Opening of the session and registration of documents').setHeading(H2);
    body.appendParagraph('5.1.1 Opening of the session').setHeading(H3);
    body.appendParagraph('The chair opens the session and welcomes the participants.');
    body.appendParagraph('The minutes are taken in this document.');
    if (o.registration !== false) {
      body.appendParagraph('5.1.2 Registration of Documents').setHeading(H3);
      body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4aA269001', 'Synthetic title', 'ExampleCorp', '5.4']]);
    }
    body.appendParagraph('5.4 Topic').setHeading(H2);
    body.appendTable([['TDoc', 'S4aA269001'], ['Title', 'Synthetic title'], ['Minutes', 'Minutes typed by hand.']]);
    body.appendParagraph('5.11 Close of the session').setHeading(H2);
    body.appendParagraph('The chair thanks the participants and closes the session.');
  }
  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' } };
  ui.alert = (...args) => { ui.alerts.push(args); return 'YES'; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title: title, html: out.html }); };
  const r = { s: s, body: body, ui: ui, docProps: loaded.docProps, scriptProps: loaded.scriptProps, logs: [], writes: [] };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => o.docId || REPORT_ID, getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.Utilities.base64Decode = (b64) => Array.from(Buffer.from(b64, 'base64')).map((b) => (b > 127 ? b - 256 : b));
  const lock = { busy: false, tries: 0, releases: 0 };
  r.lock = lock;
  s.LockService = { getDocumentLock: () => ({ tryLock: () => { lock.tries++; return !lock.busy; }, releaseLock: () => { lock.releases++; } }) };
  const set = loaded.docProps.setProperty;
  const del = loaded.docProps.deleteProperty;
  loaded.docProps.setProperty = (k, v) => { r.writes.push('set ' + k); set(k, v); };
  loaded.docProps.deleteProperty = (k) => { r.writes.push('delete ' + k); del(k); };
  r.snapshot = () => JSON.stringify(elements(body));
  r.props = () => JSON.stringify(loaded.docProps._store);
  r.save = (input) => s.saveAdhocOpeningDetails(JSON.parse(JSON.stringify(input)));
  r.container = () => { const c = s.findAdhocOpeningContainer_(body); return c ? body._children.slice(c.start, c.end).map((e) => e.getText()) : null; };
  r.texts = () => body._children.map((c) => (c.getType() === 'TABLE' ? '[table]' : c.getText()));
  return r;
}
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));
const elements = (body) => body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading(), !!c._bold]));
/** A snapshot without the generated container. */
// Both generated parts of the opening are left out: the Online information block and the Session administration section.
const withoutContainer = (snapshot) => { const out = []; let inside = false; JSON.parse(snapshot).forEach((e) => { if (e[0] === 'P' && (e[1] === HEADING || e[1] === 'Online information') && e[2] !== 'NORMAL') { inside = true; return; } if (inside && e[0] === 'P' && e[2] === 'NORMAL') return; inside = false; out.push(e); }); return JSON.stringify(out); };

const S = loadCode().sandbox;

// ================================================================ 1. the stored details

console.log('1. the stored details');
{
  const norm = S.normalizeAdhocOpeningEntry_;
  check('outer white space is removed, the rest is kept as typed', norm({ chair: '  Alex Organizer \t', minuteTakers: ' Sam Rivera,  Casey Guest ', note: '\n  Agenda approved.  \n' }), D('Alex Organizer', 'Sam Rivera,  Casey Guest', 'Agenda approved.'));
  check('line ends are normalized; inner line breaks of the note are kept', norm({ note: 'One\r\nTwo\rThree\n\nFive' }).note, 'One\nTwo\nThree\n\nFive');
  check('a line break in a single-line field becomes a space', [norm({ chair: 'Alex\nOrganizer' }).chair, norm({ minuteTakers: 'Sam Rivera\r\nCasey Guest' }).minuteTakers], ['Alex Organizer', 'Sam Rivera Casey Guest']);
  check('other scripts and punctuation are kept', norm({ chair: 'Zoë O\'Connor-Müller (ExampleCorp)', minuteTakers: '山田 太郎; 张 三 & "Sky"', note: '<b>not markup</b> — 50% § ¶' }), D('Zoë O\'Connor-Müller (ExampleCorp)', '山田 太郎; 张 三 & "Sky"', '<b>not markup</b> — 50% § ¶'));
  check('missing fields, and values that are not text, are empty', [norm({}), norm(null), norm({ chair: 5, minuteTakers: ['x'], note: { a: 1 } })], [D(), D(), D()]);

  const ser = S.serializeAdhocOpening_;
  const parse = S.parseAdhocOpeningProperty_;
  check('one session', ser({ s1: D('Alex Organizer') }), '{"v":1,"sessions":{"s1":{"chair":"Alex Organizer"}}}');
  check('several sessions, in ascending order; empty fields and empty sessions are left out', ser({ s10: D('', 'Sam Rivera'), s2: D('Alex Organizer', '', 'Note'), s3: D() }),
    '{"v":1,"sessions":{"s2":{"chair":"Alex Organizer","note":"Note"},"s10":{"minuteTakers":"Sam Rivera"}}}');
  check('nothing to store', [ser({}), ser({ s1: D(), s2: D() })], ['', '']);
  check('it holds the session id and the three fields: no label, no date, no time', [Object.keys(JSON.parse(ser(DETAILS)).sessions), Object.keys(JSON.parse(ser(DETAILS)).sessions.s1), /22 Sep|2026|15:00|Offline/.test(ser(DETAILS))], [['s1', 's2'], ['chair', 'minuteTakers', 'note'], false]);
  check('read back', parse(ser(DETAILS)), { status: 'ok', sessions: { s1: DETAILS.s1, s2: DETAILS.s2 }, error: null });
  check('and written back identically', ser(parse(ser(DETAILS)).sessions), ser(DETAILS));
  check('nothing stored', [parse(null), parse(''), parse('  ')].map((x) => [x.status, x.sessions]), [['absent', {}], ['absent', {}], ['absent', {}]]);
  ['{"v":1,', '[]', '"x"', '{"v":1}', '{"v":1,"sessions":[]}', '{"v":1,"sessions":{"22 Sep":{"chair":"x"}}}', '{"v":1,"sessions":{"s1":"x"}}', '{"v":1,"sessions":{"s1":{"chair":5}}}', '{"v":1,"sessions":{"s0":{"chair":"x"}}}'].forEach((raw) => {
    check(`unreadable (${raw.slice(0, 40)}): invalid, no details, no error thrown`, [parse(raw).status, parse(raw).sessions, typeof parse(raw).error], ['invalid', {}, 'string']);
  });
  check('a newer schema version: unsupported, no details', [parse('{"v":2,"sessions":{}}').status, parse('{"v":2}').sessions, parse('{"sessions":{}}').status], ['unsupported', {}, 'unsupported']);
  check('it never throws', thrown(() => [{}, [], 5, true, '{', '\u0000'].forEach((raw) => parse(raw))), null);
  check('which sessions have details; unknown for a newer version', [S.adhocOpeningSessionIds_(ser(DETAILS)), S.adhocOpeningSessionIds_(null), S.adhocOpeningSessionIds_('{"v":1,'), S.adhocOpeningSessionIds_('{"v":2}')], [['s1', 's2'], [], [], null]);

  check('bytes of a text as stored', ['', 'abc', 'Zoë', '山田', '😀', 'a😀b'].map(S.utf8ByteLength_), [0, 3, 4, 6, 4, 6]);

  // Saving.
  const save = (props, input, sessions) => { const l = loadCode({ documentProperties: props || {} }); const writes = []; const set = l.docProps.setProperty; const del = l.docProps.deleteProperty;
    l.docProps.setProperty = (k, v) => { writes.push('set ' + k); set(k, v); }; l.docProps.deleteProperty = (k) => { writes.push('delete ' + k); del(k); };
    return [l.sandbox.saveAdhocOpeningWith_(l.docProps, sessions || SESSIONS, input), writes, l.docProps.getProperty(KEY)]; };
  const STORED = ser(DETAILS);
  check('a save: one property written', save({}, DETAILS), [{ ok: true, errors: [], changed: true }, ['set ' + KEY], STORED]);
  check('saving the same again writes nothing', save({ [KEY]: STORED }, DETAILS), [{ ok: true, errors: [], changed: false }, [], STORED]);
  check('the same with other outer white space and line ends is the same', save({ [KEY]: STORED }, { s1: D(' Alex Organizer ', 'Sam Rivera, Casey Guest', 'The agenda was approved.\r\nThe IPR reminder was read.\n'), s2: D('', ' Sam Rivera'), s3: D() })[0].changed, false);
  check('saving nothing where nothing is stored creates no property', save({}, { s1: D(), s2: D(), s3: D() }), [{ ok: true, errors: [], changed: false }, [], null]);
  check('clearing one session removes its entry', save({ [KEY]: STORED }, { s1: DETAILS.s1, s2: D(), s3: D() }), [{ ok: true, errors: [], changed: true }, ['set ' + KEY], ser({ s1: DETAILS.s1 })]);
  check('clearing everything removes the property', save({ [KEY]: STORED }, { s1: D(), s2: D(), s3: D() }), [{ ok: true, errors: [], changed: true }, ['delete ' + KEY], null]);
  check('a session that is not configured: refused, nothing written', save({ [KEY]: STORED }, { s9: D('Alex Organizer') }), [{ ok: false, errors: ['A session of this dialog is no longer configured. Close the dialog and open it again.'], changed: false }, [], STORED]);
  check('keyed by the session id only: a label or a date is refused', [save({}, { '22 Sep': D('x') })[0].ok, save({}, { '2026-09-22': D('x') })[0].ok, save({}, { Offline: D('x') })[0].ok], [false, false, false]);
  check('too long: refused, naming the session and the field, not the text', save({ [KEY]: STORED }, { s1: D('x'.repeat(201)), s2: D('', 'y'.repeat(501), 'z'.repeat(2001)) }),
    [{ ok: false, errors: ['22 Sep: Chair is longer than 200 characters.', 'Offline: Minute taker(s) is longer than 500 characters.', 'Offline: Administrative note is longer than 2000 characters.'], changed: false }, [], STORED]);
  check('exactly at the limits', save({}, { s1: D('x'.repeat(200), 'y'.repeat(500), 'z'.repeat(2000)) })[0], { ok: true, errors: [], changed: true });
  const five = Array.from({ length: 5 }, (_, i) => session('s' + (i + 1), '2026-10-0' + (i + 1)));
  const big = {}; five.forEach((x) => { big[x.id] = D('山'.repeat(200), '田'.repeat(500), '太'.repeat(2000)); });
  check('more than one property can hold: refused, nothing written', save({}, big, five), [{ ok: false, errors: ['The opening details are too long to store together. Shorten the notes.'], changed: false }, [], null]);
  const two = { s1: big.s1, s2: big.s2 };
  check('the limit is one of stored bytes, not of characters: 5,400 characters of three bytes each are refused', [ser(two).length < 8500, S.utf8ByteLength_(ser(two)) > 8500, save({}, two, five)[0].ok, save({}, two, five)[2]], [true, true, false, null]);
  const many = {}; const twelve = Array.from({ length: 12 }, (_, i) => session('s' + (i + 1), '2026-10-' + String(i + 1).padStart(2, '0')));
  twelve.forEach((x) => { many[x.id] = D('Zoë O\'Connor-Müller (ExampleCorp)', '山田 太郎, 张 三, Sam Rivera', 'The agenda was approved. The IPR reminder was read. The next call was confirmed.'); });
  check('twelve sessions with realistic details fit, with room to spare', [save({}, many, twelve)[0].ok, S.utf8ByteLength_(ser(many)) < 3000], [true, true]);
  check('no input', [save({}, null)[0].errors, save({}, [])[0].ok, save({}, 'x')[0].ok], [['The opening details are missing.'], false, false]);
  check('an unreadable stored value is replaced by a valid save, or removed by an empty one', [save({ [KEY]: '{"v":1,' }, DETAILS)[2], save({ [KEY]: '{"v":1,' }, { s1: D() })[1]], [STORED, ['delete ' + KEY]]);
  check('a stored value of a newer version is never replaced or removed', [save({ [KEY]: '{"v":2}' }, DETAILS), save({ [KEY]: '{"v":2}' }, { s1: D() })[2]],
    [[{ ok: false, errors: ['The stored opening details have version 2; this release reads version 1. They were written by a newer release and are not changed here.'], changed: false }, [], '{"v":2}'], '{"v":2}']);
}

// ================================================================ 2. what is rendered

console.log('2. what is rendered');
{
  const when = S.adhocSessionWhenText_;
  check('when a session is: the date, with the times that are configured', [when(S1), when(session('s1', '2026-09-22', '15:00')), when(session('s1', '2026-09-22', '', '18:00')), when(session('s1', '2026-09-22'))],
    ['September 22, 2026, 15:00–18:00', 'September 22, 2026, 15:00', 'September 22, 2026, until 18:00', 'September 22, 2026']);
  check('a session with a label of its own starts with it', [when(S2), when(session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'))], ['Offline, September 23, 2026', 'Kick-off, September 22, 2026, 15:00–18:00']);
  check('no time zone and no weekday is made up', /CEST|CET|UTC|GMT|Monday|Tuesday|Wednesday/.test(SESSIONS.map(when).join(' ')), false);

  const lines = (sessions, details) => S.buildAdhocOpeningLines_(sessions, details).map((l) => (l.bold ? '*' : '') + l.text);
  check('all fields', lines([S1], { s1: DETAILS.s1 }), ['*A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera, Casey Guest', 'The agenda was approved.', 'The IPR reminder was read.']);
  check('a field without a value is left out: no placeholder', [lines([S1], { s1: D('Alex Organizer') }), lines([S1], { s1: D('', 'Sam Rivera') }), lines([S1], { s1: D('', '', 'Only a note.') })],
    [['*A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer'], ['*A01: September 22, 2026, 15:00–18:00', 'Minute taker(s): Sam Rivera'], ['*A01: September 22, 2026, 15:00–18:00', 'Only a note.']]);
  check('a session without details: when it is, and nothing else', [lines([S1], {}), lines([S1], { s1: D() }), /TBD|<|>|n\/a|unknown|\?/i.test(lines(SESSIONS, {}).join(' '))], [['*A01: September 22, 2026, 15:00–18:00'], ['*A01: September 22, 2026, 15:00–18:00'], false]);
  check('several sessions, in the order given', lines(SESSIONS, DETAILS), ['*A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera, Casey Guest', 'The agenda was approved.', 'The IPR reminder was read.',
    '*A02: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera', '*A03: September 24, 2026, 15:00']);
  check('a note of several lines: one paragraph per line, empty lines left out', lines([S1], { s1: D('', '', 'One\n\n  Two  \n\nThree') }), ['*A01: September 22, 2026, 15:00–18:00', 'One', 'Two', 'Three']);
  check('details of a session that is not given are not rendered', lines([S1], { s9: D('Nobody') }), ['*A01: September 22, 2026, 15:00–18:00']);
  check('no sessions, no lines', lines([], DETAILS), []);
  check('the text is passed through unchanged, whatever it contains', lines([S1], { s1: D('<b>Alex</b>', 'Sam & "Casey"', '<script>alert(1)</script>') }).slice(1), ['Chair: <b>Alex</b>', 'Minute taker(s): Sam & "Casey"', '<script>alert(1)</script>']);
}

// ================================================================ 3. rendering into the report

console.log('3. rendering into the report');
{
  const r = report();
  const before = r.snapshot();
  const saved = r.save(DETAILS);
  check('saved, without a notice', saved, { ok: true, errors: [], changed: true });
  check('the container: a heading and the lines of the three sessions, in chronological order', r.container(), [HEADING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera, Casey Guest', 'The agenda was approved.', 'The IPR reminder was read.',
    'A02: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera', 'A03: September 24, 2026, 15:00']);
  check('it is at the end of the opening sub-section, directly before Registration of Documents', r.texts().slice(0, 14), ['5.1 Opening of the session and registration of documents', '5.1.1 Opening of the session',
    'The chair opens the session and welcomes the participants.', 'The minutes are taken in this document.', HEADING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera, Casey Guest',
    'The agenda was approved.', 'The IPR reminder was read.', 'A02: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera', 'A03: September 24, 2026, 15:00', '5.1.2 Registration of Documents']);
  const c = r.s.findAdhocOpeningContainer_(r.body);
  check('the heading is a heading; the session lines are bold, the rest is not', [r.body._children[c.start].getHeading(), r.body._children.slice(c.start + 1, c.end).map((e) => [e.getHeading(), !!e._bold])],
    ['H3', [['NORMAL', true], ['NORMAL', false], ['NORMAL', false], ['NORMAL', false], ['NORMAL', false], ['NORMAL', true], ['NORMAL', false], ['NORMAL', true]]]);
  check('everything else in the report is exactly what it was: the prose, the tables, the closing', withoutContainer(r.snapshot()) === before, true);
  check('one container, one heading', r.texts().filter((t) => t === HEADING).length, 1);

  // Idempotent, and no churn.
  const after = r.snapshot();
  const children = r.body._children.slice();
  check('saving the same again changes nothing: no property written, and the very same elements are still in the document', [r.save(DETAILS), r.snapshot() === after, r.body._children.every((e, i) => e === children[i])], [{ ok: true, errors: [], changed: false }, true, true]);
  check('rendering again and again gives the same document', (() => { r.s.refreshAdhocOpeningSection_(r.body); r.s.refreshAdhocOpeningSection_(r.body); return [r.snapshot() === after, r.s.refreshAdhocOpeningSection_(r.body)]; })(), [true, { container: 'unchanged' }]);

  // A change replaces the container as a whole.
  r.save({ s1: D('Zoë O\'Connor-Müller', '山田 太郎'), s2: DETAILS.s2, s3: D('Alex Organizer') });
  check('changed details: the container says the new ones, once', [r.container(), r.texts().filter((t) => t === HEADING).length], [[HEADING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Zoë O\'Connor-Müller', 'Minute taker(s): 山田 太郎',
    'A02: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera', 'A03: September 24, 2026, 15:00', 'Chair: Alex Organizer'], 1]);
  check('and the rest of the report is still what it was', withoutContainer(r.snapshot()) === before, true);

  // Generated content stays generated; prose above it is kept.
  const cc = r.s.findAdhocOpeningContainer_(r.body);
  r.body._children[cc.start + 2].setText('Chair: somebody typed over this');
  r.body.insertParagraph(cc.start + 1, 'A sentence typed into the generated section.');
  r.body._children[2].setText('The chair opens the session. (edited by hand)');
  r.save({ s1: D('Zoë O\'Connor-Müller', '山田 太郎'), s2: DETAILS.s2, s3: D('Alex Organizer') });
  check('text typed into the container is replaced, although the details did not change; text typed above it is kept', [r.container().length, r.container()[2], r.texts()[2], r.texts().filter((t) => /typed/.test(t))],
    [8, 'Chair: Zoë O\'Connor-Müller', 'The chair opens the session. (edited by hand)', []]);
  check('nothing typed into the report became stored details', /typed|edited by hand/.test(r.docProps.getProperty(KEY)), false);

  // Clearing.
  r.save({ s1: D(), s2: D(), s3: D() });
  check('all details cleared: the property is gone; the container still says when the sessions are', [r.docProps.getProperty(KEY), r.container()], [null, [HEADING, 'A01: September 22, 2026, 15:00–18:00', 'A02: Offline, September 23, 2026', 'A03: September 24, 2026, 15:00']]);

  // The container never includes a table or the next section.
  const t = report();
  t.save(DETAILS);
  const tc = t.s.findAdhocOpeningContainer_(t.body);
  check('the container ends before the next heading', t.body._children[tc.end].getText(), '5.1.2 Registration of Documents');
  const moved = report({ registration: false });
  moved.body.insertTable(4, [['TDoc', 'S4aA269002'], ['Minutes', 'Typed by hand.']]);
  moved.save(DETAILS);
  const mc = moved.s.findAdhocOpeningContainer_(moved.body);
  moved.body.insertTable(mc.end, [['TDoc', 'S4aA269003'], ['Minutes', 'More minutes.']]);
  moved.save({ s1: D('Another Chair'), s2: D(), s3: D() });
  check('a table directly after the container is not part of it and survives a replacement', [moved.body._children.filter((e) => e.getType() === 'TABLE').map((e) => e.getRow(0).getCell(1).getText()), moved.container().length], [['S4aA269002', 'S4aA269003', 'S4aA269001'], 5]);

  // Placement without the Registration heading, and without any opening.
  const noReg = report({ registration: false });
  noReg.save(DETAILS);
  check('without a Registration of Documents heading: at the end of the "Opening of the session" sub-section', [noReg.texts().indexOf(HEADING), noReg.texts()[noReg.texts().indexOf(HEADING) + 9]], [4, '5.4 Topic']);
  const onlyItem = report({ built: false });
  onlyItem.body.appendParagraph('1 Opening of the meeting').setHeading(onlyItem.s.DocumentApp.ParagraphHeading.HEADING1);
  onlyItem.body.appendParagraph('Prose of the opening.');
  onlyItem.body.appendParagraph('2 Approval of the agenda').setHeading(onlyItem.s.DocumentApp.ParagraphHeading.HEADING1);
  onlyItem.save({ s1: D('Alex Organizer') });
  check('with neither: at the end of the section of the first "Opening ..." heading', onlyItem.texts().slice(0, 4).concat(onlyItem.texts().slice(-1)), ['1 Opening of the meeting', 'Prose of the opening.', HEADING, 'A01: September 22, 2026, 15:00–18:00', '2 Approval of the agenda']);
  const empty = report({ built: false });
  const result = empty.save(DETAILS);
  check('a report without an opening section: the details are stored, nothing is written, and the result says when it will be', [result, empty.body.getNumChildren(), empty.docProps.getProperty(KEY) !== null],
    [{ ok: true, errors: [], changed: true, notice: 'The opening details were saved. This report has no opening section to put them in; they are written when the report is built.' }, 0, true]);
  const plain = report();
  plain.body.insertParagraph(3, HEADING);
  plain.save(DETAILS);
  check('an ordinary paragraph that reads "Session administration" is not the container', [plain.texts()[3], plain.texts().filter((t2) => t2 === HEADING).length, plain.texts()[5]], [HEADING, 2, HEADING]);

  // The Attendance section of stage D is not touched, and the other way round.
  const a = report();
  const b64 = FX.utf16le(FX.buildExport()).toString('base64');
  const preview = a.s.previewTeamsAttendanceImport('s1', b64);
  a.s.confirmTeamsAttendanceImport('s1', b64, preview.token);
  const withAttendance = a.snapshot();
  a.save(DETAILS);
  check('with an Attendance section: saving opening details leaves it and everything else as it was', withoutContainer(a.snapshot()) === withAttendance, true);
  const beforeAttendance = () => JSON.stringify(elements(a.body).slice(0, a.s.findAdhocAttendanceContainer_(a.body).start));
  const withBoth = beforeAttendance();
  const openingElements = a.body._children.slice(0, a.s.findAdhocOpeningContainer_(a.body).end);
  a.s.refreshAdhocAttendanceSection();
  check('refreshing the Attendance section leaves the Session administration section, and everything before the Attendance heading, as it was: the very same elements',
    [beforeAttendance() === withBoth, openingElements.every((e, i) => a.body._children[i] === e)], [true, true]);
  check('both containers are found, each by its own heading', [a.s.findAdhocOpeningContainer_(a.body).start < a.s.findAdhocAttendanceContainer_(a.body).start, a.texts().filter((x) => x === HEADING || x === 'Attendance')], [true, [HEADING, 'Attendance']]);
  // Hardening of the Attendance container (found while placing this section): a table after it is not part of it.
  a.body.appendTable([['TDoc', 'S4aA269099'], ['Minutes', 'Minutes of a TDoc that an update placed at the end of the document.']]);
  a.s.refreshAdhocAttendanceSection();
  a.s.removeTeamsAttendanceImport('s1');
  check('a TDoc table placed after the Attendance section survives its refresh and its removal', a.body._children.filter((e) => e.getType() === 'TABLE' && e.getRow(0).getCell(1).getText() === 'S4aA269099').map((e) => rowsOf(e)[1][1]),
    ['Minutes of a TDoc that an update placed at the end of the document.']);

  // The last paragraph of a document is never removed.
  const last = report({ built: false });
  last.body.appendParagraph('1.1 Opening of the session').setHeading(last.s.DocumentApp.ParagraphHeading.HEADING3);
  last.save(DETAILS);
  const removeChild = last.body.removeChild;
  last.body.removeChild = (child) => { if (last.body.getChildIndex(child) === last.body.getNumChildren() - 1) throw new Error('Can\'t remove the last paragraph in a document section.'); return removeChild(child); };
  check('a container at the very end of the document can be replaced', [last.save({ s1: D('Another Chair') }).notice, last.container().slice(0, 3)], [undefined, [HEADING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Another Chair']]);
}

// ================================================================ 4. the dialog

/** Runs the dialog's client script against a minimal DOM. */
function runDialog(html, server) {
  const scripts = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
  const element = (tag) => { const e = { tag: tag, children: [], style: {}, value: '', textContent: '', className: '', disabled: false }; e.appendChild = (c) => { e.children.push(c); return c; }; return e; };
  const byId = {};
  (html.match(/id="([A-Za-z]+)"/g) || []).map((m) => m.slice(4, -1)).forEach((id) => { byId[id] = element(id); });
  const calls = [];
  const alerts = [];
  let handlers = {};
  let closed = 0;
  const runner = { withSuccessHandler: (fn) => { handlers.success = fn; return runner; }, withFailureHandler: (fn) => { handlers.failure = fn; return runner; } };
  runner.saveAdhocOpeningDetails = (...args) => { calls.push(JSON.parse(JSON.stringify(args))); const h = handlers; handlers = {}; if (server) h.success(JSON.parse(JSON.stringify(server.saveAdhocOpeningDetails(...args)))); else calls.handlers = h; };
  const ctx = { document: { getElementById: (id) => { if (!byId[id]) throw new Error('no element #' + id); return byId[id]; }, createElement: element }, google: { script: { run: runner, host: { close: () => { closed++; } } } }, alert: (m) => alerts.push(m) };
  vm.createContext(ctx);
  vm.runInContext(scripts[0].replace(/^<script>|<\/script>$/g, ''), ctx);
  const boxes = () => byId.sessions.children;
  return {
    ctx: ctx, el: byId, calls: calls, alerts: alerts, scriptCount: scripts.length, closed: () => closed,
    /** Per session: [when, label, input, label, input, label, textarea]. */
    context: () => boxes().map((box) => box.children[0].textContent),
    labels: (i) => boxes()[i].children.filter((c) => c.tag === 'label').map((c) => c.textContent),
    inputs: (i) => boxes()[i].children.filter((c) => c.tag === 'input' || c.tag === 'textarea'),
    values: () => boxes().map((box) => box.children.filter((c) => c.tag === 'input' || c.tag === 'textarea').map((c) => c.value))
  };
}

console.log('4. the dialog');
{
  const r = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: stored({ s1: { chair: 'Alex Organizer', minuteTakers: 'Sam Rivera, Casey Guest', note: 'The agenda was approved.\nThe IPR reminder was read.' }, s2: { minuteTakers: 'Sam Rivera' } }) }) });
  const props0 = r.props();
  const body0 = r.snapshot();
  r.s.editAdhocOpeningDetails();
  check('one dialog, titled Edit Opening Details; opening it writes nothing and changes nothing', [r.ui.dialogs.length, r.ui.dialogs[0].title, r.ui.alerts.length, r.props() === props0, r.snapshot() === body0, r.writes, r.lock.tries], [1, 'Edit Opening Details', 0, true, true, [], 0]);
  const d = runDialog(r.ui.dialogs[0].html, r.s);
  check('one script; one block per session, in chronological order, each saying when the session is', [d.scriptCount, d.context()], [1, ['September 22, 2026, 15:00–18:00', 'Offline, September 23, 2026', 'September 24, 2026, 15:00']]);
  check('three fields per session, by name: two lines of text and a small text area', [d.labels(0), d.inputs(0).map((i) => i.tag + (i.type ? ':' + i.type : '')), d.inputs(0).map((i) => i.maxLength)], [['Chair', 'Minute taker(s)', 'Administrative note'], ['input:text', 'input:text', 'textarea'], [200, 500, 2000]]);
  check('the stored values are loaded', d.values(), [['Alex Organizer', 'Sam Rivera, Casey Guest', 'The agenda was approved.\nThe IPR reminder was read.'], ['', 'Sam Rivera', ''], ['', '', '']]);
  check('the session context is text, not a field', [d.el.sessions.children[0].children[0].tag, d.el.sessions.children[0].children[0].className], ['div', 'when']);
  check('the dialog says that the section is generated and where to change it', /The section "Session administration" of the report is written from these fields and from the sessions; change it here, because text typed into that section is replaced\./.test(r.ui.dialogs[0].html), true);
  check('nothing was sent by opening', d.calls, []);

  // Edit, then cancel.
  d.inputs(0)[0].value = 'Somebody Else';
  d.inputs(2)[2].value = 'A note.';
  d.ctx.google.script.host.close();
  check('Cancel after editing: nothing sent, nothing stored, the report untouched', [d.closed(), d.calls.length, r.props() === props0, r.snapshot() === body0, r.writes], [1, 0, true, true, []]);
  check('the Cancel button only closes the dialog (source)', (r.ui.dialogs[0].html.match(/<button[^>]*id="closeBtn"[^>]*>/) || [''])[0], '<button type="button" class="grey" id="closeBtn" onclick="google.script.host.close()">');
  check('the only server function the dialog calls is the save', [(r.ui.dialogs[0].html.match(/\.\s*([A-Za-z]+)\(collect\(\)\)/) || [])[1], (r.ui.dialogs[0].html.match(/google\.script\.run/g) || []).length], ['saveAdhocOpeningDetails', 1]);

  // Save.
  d.inputs(0)[0].value = '  Zoë O\'Connor-Müller ';
  d.inputs(1)[1].value = '';
  d.inputs(2)[2].value = 'Line one\nLine two';
  d.ctx.saveDetails();
  check('Save sends every session by its id with its three fields, as typed', d.calls, [[{ s1: D('  Zoë O\'Connor-Müller ', 'Sam Rivera, Casey Guest', 'The agenda was approved.\nThe IPR reminder was read.'), s2: D('', '', ''), s3: D('', '', 'Line one\nLine two') }]]);
  check('one property is written, once; the dialog closes without a message', [r.writes, d.closed(), d.alerts, r.docProps.getProperty(KEY)],
    [['set ' + KEY], 2, [], stored({ s1: { chair: 'Zoë O\'Connor-Müller', minuteTakers: 'Sam Rivera, Casey Guest', note: 'The agenda was approved.\nThe IPR reminder was read.' }, s3: { note: 'Line one\nLine two' } })]);
  check('the report shows it at once', r.container(), [HEADING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Zoë O\'Connor-Müller', 'Minute taker(s): Sam Rivera, Casey Guest', 'The agenda was approved.', 'The IPR reminder was read.',
    'A02: Offline, September 23, 2026', 'A03: September 24, 2026, 15:00', 'Line one', 'Line two']);
  check('it held the document lock for the rendering and released it', [r.lock.tries, r.lock.releases], [1, 1]);

  // Save again unchanged; then clear everything.
  r.s.editAdhocOpeningDetails();
  const e = runDialog(r.ui.dialogs[1].html, r.s);
  const writes = r.writes.length;
  const snap = r.snapshot();
  e.ctx.saveDetails();
  check('Save without a change: nothing written, the report untouched, the dialog closes', [r.writes.length === writes, r.snapshot() === snap, e.closed()], [true, true, 1]);
  r.s.editAdhocOpeningDetails();
  const f = runDialog(r.ui.dialogs[2].html, r.s);
  [0, 1, 2].forEach((i) => f.inputs(i).forEach((input) => { input.value = ''; }));
  f.ctx.saveDetails();
  check('clearing every field and saving removes the property; the section keeps the sessions', [r.docProps.getProperty(KEY), r.writes.slice(writes), r.container()], [null, ['delete ' + KEY], [HEADING, 'A01: September 22, 2026, 15:00–18:00', 'A02: Offline, September 23, 2026', 'A03: September 24, 2026, 15:00']]);

  // A refused save, a failed request, a notice.
  r.s.editAdhocOpeningDetails();
  const g = runDialog(r.ui.dialogs[3].html, null);
  g.ctx.saveDetails();
  check('Save is disabled while the request runs', g.el.saveBtn.disabled, true);
  g.calls.handlers.success({ ok: false, errors: ['First problem.', 'Second problem.'], changed: false });
  check('a refused save: the messages are shown, the dialog stays open', [g.el.errors.textContent, g.closed(), g.el.saveBtn.disabled], ['First problem.\nSecond problem.', 0, false]);
  g.ctx.saveDetails();
  g.calls.handlers.failure(new Error('Server error'));
  check('a failed request', [g.el.errors.textContent, g.closed(), g.el.saveBtn.disabled], ['Server error', 0, false]);
  g.ctx.saveDetails();
  g.calls.handlers.success({ ok: true, changed: true, notice: 'Saved, but the section could not be written.' });
  check('a save with a notice: the notice is shown, then the dialog closes', [g.alerts, g.closed()], [['Saved, but the section could not be written.'], 1]);
  r.s.editAdhocOpeningDetails();
  const long = runDialog(r.ui.dialogs[4].html, r.s);
  long.inputs(0)[0].value = 'x'.repeat(201);
  long.ctx.saveDetails();
  check('a value that is too long is refused by the server too; nothing is stored', [long.el.errors.textContent, long.closed(), r.docProps.getProperty(KEY)], ['22 Sep: Chair is longer than 200 characters.', 0, null]);
  const refusedIn = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) }) });
  const refused0 = refusedIn.snapshot();
  check('a refused save does not touch the report: no section is written from what is stored', [refusedIn.save({ s1: D('x'.repeat(201)) }).ok, refusedIn.snapshot() === refused0, refusedIn.container(), refusedIn.lock.tries], [false, true, null, 0]);

  // Untrusted text stays text.
  const hostile = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-09-22', '', '', '</script><b>x')]), [KEY]: stored({ s1: { chair: '<img src=x onerror=alert(1)>', note: '</script><script>alert(2)</script>' } }) }) });
  hostile.s.editAdhocOpeningDetails();
  const h = runDialog(hostile.ui.dialogs[0].html, hostile.s);
  check('a label, a name and a note with markup: one script element, shown as typed', [h.scriptCount, h.context()[0], h.values()[0], hostile.ui.dialogs[0].html.indexOf('</script><b>'), hostile.ui.dialogs[0].html.indexOf('<img')],
    [1, '</script><b>x, September 22, 2026', ['<img src=x onerror=alert(1)>', '', '</script><script>alert(2)</script>'], -1, -1]);
  check('the dialog script builds everything with textContent and value, never with innerHTML (source)', [/innerHTML|insertAdjacentHTML|document\.write/.test(hostile.ui.dialogs[0].html), /textContent/.test(hostile.ui.dialogs[0].html)], [false, true]);
  h.ctx.saveDetails();
  check('in the report it is text as well', hostile.container(), [HEADING, 'A01: </script><b>x, September 22, 2026', 'Chair: <img src=x onerror=alert(1)>', '</script><script>alert(2)</script>']);

  // The model.
  const model = r.s.buildAdhocOpeningDialogModel_();
  check('what the dialog is given: per session its id, when it is and the three fields; nothing else about anybody', [Object.keys(model), Object.keys(model.sessions[0]), model.fields.map((x) => x.key)], [['notice', 'readOnly', 'fields', 'sessions'], ['id', 'when', 'chair', 'minuteTakers', 'note'], ['chair', 'minuteTakers', 'note']]);

  // Unreadable and newer stored values.
  const corrupt = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: '{"v":1,' }) });
  corrupt.s.editAdhocOpeningDetails();
  const cd = runDialog(corrupt.ui.dialogs[0].html, corrupt.s);
  check('stored details that cannot be read: a notice, empty fields, saving possible; opening changed nothing', [/cannot be read\. They are not used; saving replaces them\./.test(cd.el.notice.textContent), cd.el.notice.style.display, cd.values()[0], cd.el.saveBtn.disabled, corrupt.docProps.getProperty(KEY)], [true, 'block', ['', '', ''], false, '{"v":1,']);
  const newer = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: '{"v":2,"sessions":{}}' }) });
  newer.s.editAdhocOpeningDetails();
  const nd = runDialog(newer.ui.dialogs[0].html, newer.s);
  check('a newer version: a notice, nothing can be edited or saved', [/version 2/.test(nd.el.notice.textContent), nd.el.saveBtn.disabled, nd.inputs(0).map((i) => i.disabled)], [true, true, [true, true, true]]);
  check('and the server refuses a save as well; the report is not touched', [newer.save(DETAILS).ok, newer.docProps.getProperty(KEY), newer.container()], [false, '{"v":2,"sessions":{}}', null]);

  // Where the dialog does not open.
  const main = report({ props: Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }) });
  main.s.editAdhocOpeningDetails();
  check('a main-meeting report: a message, no dialog; a save is refused; nothing is written', [main.ui.dialogs.length, main.ui.alerts.map((x) => x[1]), main.save(DETAILS), main.writes, main.container()],
    [0, ['Opening details are available for ad-hoc reports only.'], { ok: false, errors: ['Opening details are available for ad-hoc reports only.'], changed: false }, [], null]);
  const noSessions = report({ props: ADHOC });
  noSessions.s.editAdhocOpeningDetails();
  check('an ad-hoc report without sessions: where to configure them, no dialog; a save is refused; nothing is written', [noSessions.ui.dialogs.length, noSessions.ui.alerts.map((x) => x[1]), noSessions.save(DETAILS).errors, noSessions.writes, noSessions.container()],
    [0, ['Opening details are entered per session. Configure the sessions of this report first:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…'], ['Opening details are entered per session. Use Configure Sessions first.'], [], null]);
  const master = report({ template: true, docId: TEMPLATE_ID });
  check('in the master template both actions refuse', ['editAdhocOpeningDetails', 'saveAdhocOpeningDetails'].map((fn) => /SA4 Report Template itself/.test(thrown(() => master.s[fn](DETAILS)) || '')).concat([master.writes.length]), [true, true, 0]);
}

// ================================================================ 5. sessions that change

console.log('5. sessions that change');
{
  const r = report();
  r.save(DETAILS);
  const storedDetails = r.docProps.getProperty(KEY);
  const rows = (list) => list.map((x) => Object.assign({}, x));

  const renamed = r.s.saveAdhocSessionsConfiguration(rows([session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), S2, S3]));
  check('renaming a session: saved, without a notice', [renamed.ok, renamed.changed, renamed.notice], [true, true, undefined]);
  check('the section shows the new label at once; the details stay with the session', [r.container().slice(0, 3), r.docProps.getProperty(KEY) === storedDetails], [[HEADING, 'A01: Kick-off, September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer'], true]);
  r.s.saveAdhocSessionsConfiguration(rows([session('s1', '2026-09-25', '09:30', '', 'Kick-off'), S2, S3]));
  check('changing its date and times: the section follows, in the new order; the details are still that session\'s', [r.container(), r.docProps.getProperty(KEY) === storedDetails], [[HEADING, 'A01: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera', 'A02: September 24, 2026, 15:00',
    'A03: Kick-off, September 25, 2026, 09:30', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera, Casey Guest', 'The agenda was approved.', 'The IPR reminder was read.'], true]);
  const snap = r.snapshot();
  check('saving unchanged sessions does not touch the report', [r.s.saveAdhocSessionsConfiguration(rows([session('s1', '2026-09-25', '09:30', '', 'Kick-off'), S2, S3])).changed, r.snapshot() === snap], [false, true]);
  const edited = report();
  edited.save(DETAILS);
  edited.body._children[edited.s.findAdhocOpeningContainer_(edited.body).start + 2].setText('Chair: typed over by hand');
  const edited0 = edited.snapshot();
  check('nor does it write the section again: an unchanged save of the sessions is no reason to', [edited.s.saveAdhocSessionsConfiguration(rows(SESSIONS)).changed, edited.snapshot() === edited0, edited.lock.tries], [false, true, 1]);
  r.s.saveAdhocSessionsConfiguration(rows([S1, S2, S3, session('', '2026-10-01')]));
  check('a session added later appears in the section, without details', r.container().slice(-1), ['A04: October 1, 2026']);

  // A session with details is not removed.
  const before = r.props();
  const bodyBefore = r.snapshot();
  const refused = r.s.saveAdhocSessionsConfiguration(rows([S2, S3]));
  check('removing a session that has opening details is refused, with the way to do it', refused, { ok: false, errors: ['22 Sep has opening details and was not removed. Clear them first: Sessions and Attendance > Edit Opening Details….'], changed: false, sessions: [] });
  check('nothing was changed: not the sessions, not the details, not the report', [r.props() === before, r.snapshot() === bodyBefore], [true, true]);
  check('removing two sessions with details: both are named', r.s.saveAdhocSessionsConfiguration(rows([S3])).errors.map((e) => e.split(' has ')[0]), ['22 Sep', 'Offline']);
  const freed = r.s.saveAdhocSessionsConfiguration(rows([S1, S2]));
  check('sessions without details can be removed; the section follows', [freed.ok, r.s.getAdhocSessions_().map((x) => x.id), r.container().filter((t) => /2026/.test(t))], [true, ['s1', 's2'], ['A01: September 22, 2026, 15:00–18:00', 'A02: Offline, September 23, 2026']]);
  r.save({ s1: D(), s2: DETAILS.s2 });
  check('after its details were cleared, the session can go; no details are left without their session', [r.s.saveAdhocSessionsConfiguration(rows([S2])).ok, r.s.getAdhocSessions_().map((x) => x.id), r.s.adhocOpeningSessionIds_(r.docProps.getProperty(KEY)), r.container()],
    [true, ['s2'], ['s2'], [HEADING, 'A01: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera']]);
  check('a session added afterwards does not get the id of a removed one', r.s.saveAdhocSessionsConfiguration([S2, session('', '2026-10-01')]).sessions.map((x) => x.id), ['s2', 's5']);

  // The pure function, with what it is told.
  const tell = (ids, list) => { const l = loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(SESSIONS) } }); const out = l.sandbox.saveAdhocSessionsWith_(l.docProps, true, list, [], {}, ids); return [out.ok, l.docProps.getProperty('ADHOC_SESSIONS') === sessionsProperty(SESSIONS)]; };
  check('no opening details: sessions are removed as before', [tell([], rows([S2, S3])), tell(undefined, rows([S2, S3]))], [[true, false], [true, false]]);
  check('details on a session that stays: the others can be removed', tell(['s2'], rows([S2])), [true, false]);
  check('details on the session to be removed: refused, nothing written', tell(['s1'], rows([S2, S3])), [false, true]);
  check('details of a newer version (which sessions have them is unknown): no session is removed, but sessions can be edited and added',
    [tell(null, rows([S2, S3])), tell(null, rows(SESSIONS).concat([session('', '2026-10-01')])), loadCode().sandbox.saveAdhocSessionsWith_(loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(SESSIONS) } }).docProps, true, rows([S2, S3]), [], {}, null).errors[0]],
    [[false, true], [true, false], 'The opening details were written by a newer release, so it cannot be told which sessions have them. No session was removed.']);
  check('the guards of attendance and of manual TDoc assignments come first, unchanged', (() => { const l = loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(SESSIONS) } });
    return [l.sandbox.saveAdhocSessionsWith_(l.docProps, true, rows([S2, S3]), ['s1'], { s1: ['S4aA269001'] }, ['s1']).errors[0].indexOf('imported attendance') !== -1, l.sandbox.saveAdhocSessionsWith_(l.docProps, true, rows([S2, S3]), [], { s1: ['S4aA269001'] }, ['s1']).errors[0].indexOf('manual TDoc session assignments') !== -1]; })(), [true, true]);

  // Removing every session.
  const all = report();
  all.save({ s1: D(), s2: D(), s3: D() });
  check('a section without details exists', all.container().length, 4);
  check('removing every session (none has details): the section goes with them', [all.s.saveAdhocSessionsConfiguration([]).ok, all.container(), all.texts().indexOf(HEADING)], [true, null, -1]);

  // A report that has neither the section nor details is not touched by saving sessions.
  const untouched = report();
  const u0 = untouched.snapshot();
  check('saving sessions in a report without the section and without details changes nothing in the document', [untouched.s.saveAdhocSessionsConfiguration(rows([session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), S2, S3])).notice, untouched.snapshot() === u0, untouched.lock.tries], [undefined, true, 0]);
}

// ================================================================ 6. Build from Scratch

const FXDIR = path.join(__dirname, 'fixtures');
const LIST_85916 = JSON.parse(fs.readFileSync(path.join(FXDIR, 'meeting-85916-tdoc-list-values.json'), 'utf8')).values;
const FTP_AUDIO = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const PROPS_85916 = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 28, 2026',
  FTP_BASE: FTP_AUDIO, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const AUDIO_AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'ULBC-MED'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));

/** A report whose real buildSkeletonWithTdocTables() can be run; only network, Drive, Sheets and the template are faked. */
function buildable(props, opts) {
  const r = report({ props: props, template: opts && opts.template, built: false });
  const s = r.s;
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }) };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => LIST_85916, getRichTextValues: () => LIST_85916.map((row) => row.map(() => null)) }) }] }) };
  s.downloadMeetingAgenda_ = () => AUDIO_AGENDA;
  r.build = () => { try { return s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }); } catch (e) { return { ok: false, message: e.message }; } };
  return r;
}
const AUDIO_SESSIONS = Object.assign({}, PROPS_85916, { ADHOC_SESSIONS: sessionsProperty(SESSIONS) });

console.log('6. Build from Scratch');
{
  const plain = buildable(PROPS_85916);
  plain.build();

  const r = buildable(AUDIO_SESSIONS);
  const first = r.build();
  check('a build of an ad-hoc report with sessions: completes without a note', [first.ok, first.note], [true, '']);
  check('it has the section, once, at the end of the opening sub-section and before Registration of Documents; above it the Online information block', [r.texts().filter((t) => t === HEADING).length, r.texts().slice(r.texts().indexOf('1.1 Opening of the session'), r.texts().indexOf('1.2 Registration of Documents') + 1)],
    [1, ['1.1 Opening of the session', 'Online information', 'Meeting name: SA4-(AH) Audio SWG on ULBC-MED', 'Start Date: September 22, 2026', 'End Date: September 24, 2026',
      'Portal meeting URL: https://portal.3gpp.org/Home.aspx#/meeting?MtgId=85916', 'TDoc List URL: https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=85916',
      'Excel Docs URL: https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', 'Docs Folder: https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/',
      HEADING, 'A01: September 22, 2026, 15:00–18:00', 'A02: Offline, September 23, 2026', 'A03: September 24, 2026, 15:00', '1.2 Registration of Documents']]);
  check('without details it says when the sessions are, and nothing was stored for it', [r.container().length, r.docProps.getProperty(KEY)], [4, null]);
  // E1: the section is the one place that says who chaired and when.
  // After T-2026.10.8: no time zone is assumed ("CEST" was), and this report has no stored start (tests/meeting-start-opening.test.js).
  const LEGACY = '<Chair> opens the session on September 28, 2026 at <start> <time zone>.';
  check('with sessions the build does not write the "<Chair> opens the session ..." line, and nothing in its place', [r.texts().filter((t) => /<Chair>|opens the session on|<start>|CEST/.test(t)), r.texts()[r.texts().indexOf(HEADING) - 1]], [[], 'Docs Folder: https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/']);
  check('an ad-hoc report without sessions still gets that line, where it always was', [plain.texts().filter((t) => /opens the session on/.test(t)), plain.texts()[plain.texts().indexOf('1.1 Opening of the session') + 1], plain.texts()[plain.texts().indexOf('1.1 Opening of the session') + 2]],
    [[LEGACY], LEGACY, '1.2 Registration of Documents']);
  check('so does one whose stored sessions are empty or cannot be read', [sessionsProperty([]), '{"v":1,'].map((value) => { const x = buildable(Object.assign({}, PROPS_85916, { ADHOC_SESSIONS: value })); x.build(); return [x.snapshot() === plain.snapshot(), x.texts().indexOf(HEADING)]; }), [[true, -1], [true, -1]]);
  check('a main-meeting report (even with sessions left in its properties) is not affected: the line is in the ad-hoc branch only, which asks for the sessions of an ad-hoc report (source)', (() => {
    const src = CODE.slice(CODE.indexOf('\nfunction buildSkeletonWithTdocTables('), CODE.indexOf('\nfunction downloadAndGroupTdocs_('));
    const adhocBranch = src.slice(src.indexOf("if (context.meeting.type === 'adhoc') {"), src.indexOf('// Opening section for SWG reports - copy X.1 content from template'));
    const main = loadCode({ documentProperties: { MEETING_TYPE: 'main', ADHOC_SESSIONS: sessionsProperty(SESSIONS) } }).sandbox;
    return [(src.match(/buildMeetingOpeningSentence_\(/g) || []).length, /if \(!adhocSessionsEnabled_\(context\)\) \{\n\s+const meetingStart = getMeetingStartForOpening_\(\);\n\s+body\.appendParagraph\(buildMeetingOpeningSentence_\(cfg\.MEETING_DATE, meetingStart\.time, meetingStart\.zone\)\);\n\s+\}/.test(adhocBranch),
      (src.match(/adhocSessionsEnabled_\(/g) || []).length, main.adhocSessionsEnabled_(),
      src.slice(src.indexOf('    } else {\n      // Opening section for SWG reports'), src.indexOf('    // Always ensure Registration of Documents section exists')).split('\n').map((line) => line.trim())];
  })(), [1, true, 1, false, ['} else {', '// Opening section for SWG reports - copy X.1 content from template', 'const openingHeader = findHeading_(sourceBody, /^X\\.1\\s+/);', 'if (openingHeader) {',
    "copySectionContentWithReplacement_(openingHeader, body, /^X\\.2\\s+/, 'X', agendaPrefixNum);", '}', '}', '', '']]);

  r.save(DETAILS);
  r.body.insertParagraph(r.s.findAdhocOpeningContainer_(r.body).start, 'An introduction typed by hand.');
  const storedBefore = r.docProps.getProperty(KEY);
  const rebuilt = r.build();
  check('a rebuild completes; the details are still stored, unchanged', [rebuilt.ok, rebuilt.note, r.docProps.getProperty(KEY) === storedBefore], [true, '', true]);
  check('the section is written again, once, in the same place, with the details', [r.texts().filter((t) => t === HEADING).length, r.texts()[r.texts().indexOf(HEADING) - 1], r.container()],
    [1, 'Docs Folder: https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', [HEADING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera, Casey Guest', 'The agenda was approved.', 'The IPR reminder was read.',
      'A02: Offline, September 23, 2026', 'Minute taker(s): Sam Rivera', 'A03: September 24, 2026, 15:00']]);
  check('prose typed by hand is lost, as the build says; it did not become stored details', [r.texts().filter((t) => /typed by hand/.test(t)), /typed by hand/.test(r.docProps.getProperty(KEY))], [[], false]);
  const after = r.snapshot();
  r.build();
  check('a second rebuild gives the same document', r.snapshot() === after, true);
  check('apart from the section (in place of the "<Chair> opens ..." line) and the Session column, the report is the report built without sessions', (() => {
    const strip = (snapshot) => JSON.stringify(JSON.parse(withoutContainer(snapshot)).map((e) => (e[0] === 'T' && e[1][0].length === 5 && e[1][0][4] === 'Session' ? ['T', e[1].map((row) => row.slice(0, 4))] : e)));
    const withoutLegacyLine = JSON.stringify(JSON.parse(plain.snapshot()).filter((e) => !(e[0] === 'P' && e[1] === LEGACY)));
    return [strip(r.snapshot()) === withoutLegacyLine, JSON.parse(plain.snapshot()).length - JSON.parse(withoutLegacyLine).length];
  })(), [true, 1]);

  // Sessions edited between builds.
  r.s.saveAdhocSessionsConfiguration([session('s1', '2026-09-29', '10:00', '12:00', 'Wrap-up'), S2, S3].map((x) => Object.assign({}, x)));
  r.build();
  check('a rebuild uses the sessions as they are configured now; the details follow their session', r.container().slice(-5), ['A03: Wrap-up, September 29, 2026, 10:00–12:00', 'Chair: Alex Organizer', 'Minute taker(s): Sam Rivera, Casey Guest', 'The agenda was approved.', 'The IPR reminder was read.']);

  // Problems during a rebuild do not fail the build.
  const failing = buildable(Object.assign({}, AUDIO_SESSIONS, { [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) }));
  failing.s.renderAdhocOpeningSection_ = () => { throw new Error('Service Documents failed'); };
  const failed = failing.build();
  check('the section cannot be written: the report is built, the note says how to write it, the details are still stored',
    [failed.ok, failed.note, failing.docProps.getProperty(KEY), failing.texts().indexOf('5 Close of the meeting') !== -1],
    [true, '\n\n⚠️ The Session administration section could not be written (Service Documents failed). Save Edit Opening Details… again to write it.', stored({ s1: { chair: 'Alex Organizer' } }), true]);
  failing.s.renderAdhocOpeningSection_ = loadCode().sandbox.renderAdhocOpeningSection_;
  check('saving the dialog then writes it, in the opening sub-section (below the Online information block, which the build wrote)', [failing.texts().indexOf(HEADING), failing.save({ s1: D('Alex Organizer') }).notice,
    [failing.texts()[failing.texts().indexOf('1.1 Opening of the session') + 1]].concat(failing.texts().slice(failing.texts().indexOf(HEADING), failing.texts().indexOf(HEADING) + 3), [failing.texts()[failing.texts().indexOf(HEADING) + 3]])],
    [-1, undefined, ['Online information', HEADING, 'A01: September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer', 'A02: Offline, September 23, 2026']]);
  check('the failure is logged without any of the details', [failing.logs.filter((l) => /^Opening details/.test(l)), failing.logs.filter((l) => /Organizer/.test(l))], [['Opening details: the section could not be written after the rebuild: Service Documents failed'], []]);
  const corrupt = buildable(Object.assign({}, AUDIO_SESSIONS, { [KEY]: '{"v":1,' }));
  corrupt.build();
  check('stored details that cannot be read: the build writes the sessions without details and leaves the stored value alone', [corrupt.container(), corrupt.docProps.getProperty(KEY)], [[HEADING, 'A01: September 22, 2026, 15:00–18:00', 'A02: Offline, September 23, 2026', 'A03: September 24, 2026, 15:00'], '{"v":1,']);

  // Where the build does nothing new.
  const stray = buildable(Object.assign({}, PROPS_85916, { [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) }));
  stray.build();
  check('ad-hoc without sessions (even with opening details left in its properties): the build is the build of a report without them', [stray.snapshot() === plain.snapshot(), stray.texts().indexOf(HEADING), stray.docProps.getProperty(KEY)], [true, -1, stored({ s1: { chair: 'Alex Organizer' } })]);
  const reads = [];
  const quiet = buildable(PROPS_85916);
  const get = quiet.docProps.getProperty;
  quiet.docProps.getProperty = (k) => { reads.push(k); return get(k); };
  quiet.build();
  check('a report without sessions never reads the opening details during a build', reads.filter((k) => k === KEY), []);
  const main = loadCode({ documentProperties: { MEETING_TYPE: 'main', ADHOC_SESSIONS: sessionsProperty(SESSIONS), [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) } });
  const mainReads = [];
  const mainGet = main.docProps.getProperty;
  main.docProps.getProperty = (k) => { mainReads.push(k); return mainGet(k); };
  check('a main-meeting report: the hook does nothing, and reads nothing but the meeting type', [main.sandbox.finishAdhocOpeningRebuild_({}), mainReads], ['', ['MEETING_TYPE']]);
  check('the build calls the hook once, after the reallocations are restored (source)', (() => {
    const src = CODE.slice(CODE.indexOf('\nfunction buildSkeletonWithTdocTables('), CODE.indexOf('\nfunction downloadAndGroupTdocs_('));
    return [(src.match(/finishAdhocOpeningRebuild_\(/g) || []).length, src.indexOf('finishAdhocOpeningRebuild_(') > src.indexOf('saveReallocation(tdoc, r.original, r.new, r.reason)')];
  })(), [1, true]);

  // The question before the build.
  const question = (props, details) => { const x = report({ props: Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, props), template: true }); if (details) x.save(DETAILS); x.ui.alert = (...args) => { x.ui.alerts.push(args); return 'NO'; }; x.s.runFullReportBuild(); return x.ui.alerts[x.ui.alerts.length - 1][1]; };
  check('the build question of a report with opening details says that they are kept', /Only the Document Reallocations table and the opening details of the sessions are kept\./.test(question(WITH_SESSIONS, true)), true);
  check('with attendance as well, all three are named', (() => {
    const x = report({ props: Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, WITH_SESSIONS), template: true });
    x.save(DETAILS);
    const b64 = FX.utf16le(FX.buildExport()).toString('base64');
    x.s.confirmTeamsAttendanceImport('s1', b64, x.s.previewTeamsAttendanceImport('s1', b64).token);
    x.ui.alert = (...args) => { x.ui.alerts.push(args); return 'NO'; };
    x.s.runFullReportBuild();
    return /Only the Document Reallocations table, the imported attendance \(with its Company cells\) and the opening details of the sessions are kept\./.test(x.ui.alerts[x.ui.alerts.length - 1][1]);
  })(), true);
  check('without opening details the question is the one of T-2026.10.4', [question(WITH_SESSIONS, false), question(ADHOC, false), question(Object.assign({}, ADHOC, { MEETING_TYPE: 'main' }), false)].map((t) => /Only the Document Reallocations table is kept\./.test(t)), [true, true, true]);
}

// ================================================================ 7. updates

console.log('7. updates');
{
  const r = report();
  r.save(DETAILS);
  const body0 = r.snapshot();
  const props0 = r.props();
  const children = r.body._children.slice();
  const reads = [];
  const get = r.docProps.getProperty;
  r.docProps.getProperty = (k) => { reads.push(k); return get(k); };
  const events = [];
  r.s.assertMeetingReadyToBuild_ = () => {};
  r.s.downloadAndGroupTdocs_ = () => { events.push('TDoc list'); return {}; };
  r.s.rearrangeRevisionTables_ = () => { events.push('revision placement'); return { moved: 0, dispositions: 0 }; };
  r.s.checkRSSFeed_ = () => { events.push('e-mail'); };
  r.s.updateRevisions_ = () => { events.push('revisions'); };
  r.s.removeRowHeightAndSpacing = () => { events.push('formatting'); };
  const result = r.s.continuousUpdateCore_();
  r.s.continuousUpdateCore_();
  check('a complete update, twice: the usual steps, the report and the details untouched, the very same elements in the document', [result, events.slice(0, 4), r.snapshot() === body0, r.props() === props0, r.body._children.every((e, i) => e === children[i])],
    [{ success: true, error: null }, ['TDoc list', 'revision placement', 'e-mail', 'revisions'], true, true, true]);
  check('an update never reads the opening details', reads.filter((k) => k === KEY), []);
  check('one section after the updates', r.texts().filter((t) => t === HEADING).length, 1);
  check('the update, the collector and the registration table do not mention the opening (source)', ['continuousUpdateCore_', 'collectorUpdate_', 'updateRegisteredDocumentsTable_', 'createSummaryTable_', 'insertNewTdoc_', 'runFullReportBuildCore_', 'downloadAndGroupTdocs_'].filter((name) => {
    const at = CODE.indexOf('\nfunction ' + name + '(');
    return /Opening_|OPENING/.test(CODE.slice(at, CODE.indexOf('\n}\n', at)));
  }), []);
  // A new TDoc table lands where it always did: the unnumbered heading is not a section boundary for it.
  const u = report();
  u.save(DETAILS);
  check('the end of the opening agenda item\'s section is where it was before the section existed', [u.s.findAgendaSectionEndIndex_(u.body, '5.1') - u.container().length, report().s.findAgendaSectionEndIndex_(report().body, '5.1')], [6, 6]);
  check('the section is not taken for a TDoc table, the registration table, or the Attendance container', [u.s.findRegistrationTable_(u.body).columns, u.s.findAdhocAttendanceContainer_(u.body), u.s.deriveAdminAnchorsFromBuiltDocument_(u.body).registrationSection], [4, null, '5.1.2']);
}

// ================================================================ 8. failure and recovery

console.log('8. failure and recovery');
{
  // Storing fails.
  const a = report();
  const body0 = a.snapshot();
  const set = a.docProps.setProperty;
  a.docProps.setProperty = (k) => { if (k === KEY) throw new Error('Properties storage quota exceeded'); };
  const failed = a.save(DETAILS);
  a.docProps.setProperty = set;
  check('storing fails: nothing is saved and the report is not touched', [failed, a.docProps.getProperty(KEY), a.snapshot() === body0, a.lock.tries], [{ ok: false, errors: ['Nothing was saved: the opening details could not be stored (Properties storage quota exceeded).'], changed: false }, null, true, 0]);

  // Rendering fails after storing.
  const b = report();
  const realRender = b.s.renderAdhocOpeningSection_;
  b.s.renderAdhocOpeningSection_ = () => { throw new Error('Service Documents failed'); };
  const half = b.save(DETAILS);
  check('rendering fails after storing: the details are kept and the result says how to write the section', [half, b.docProps.getProperty(KEY) !== null, b.container(), b.lock.releases],
    [{ ok: true, errors: [], changed: true, notice: 'The opening details were saved. The Session administration section could not be written (Service Documents failed). Save Edit Opening Details… again to write it.' }, true, null, 1]);
  check('the failure is logged without a name or a note', [b.logs, b.logs.filter((l) => /Organizer|Rivera|agenda/.test(l))], [['Opening details: stored, but the section could not be written: Service Documents failed'], []]);
  b.s.renderAdhocOpeningSection_ = realRender;
  check('saving the dialog again, unchanged, writes the section from what is stored, and stores nothing again', [b.save(DETAILS), b.container().length, b.writes.filter((w) => w === 'set ' + KEY).length], [{ ok: true, errors: [], changed: false }, 9, 1]);

  // A section that was damaged by hand is repaired by an unchanged save.
  const cc = b.s.findAdhocOpeningContainer_(b.body);
  b.body.removeChild(b.body._children[cc.start + 3]);
  check('a section with a line missing is completed by saving again', [b.container().length, b.save(DETAILS).changed, b.container().length], [8, false, 9]);

  // The document is busy.
  const c = report();
  c.lock.busy = true;
  check('while an update is running: the details are saved, the section is not written, and the result says so', [c.save(DETAILS), c.container(), c.docProps.getProperty(KEY) !== null],
    [{ ok: true, errors: [], changed: true, notice: 'The opening details were saved. The Session administration section could not be written (Nothing was changed: an update of this report is running right now. Try again in a minute.). Save Edit Opening Details… again to write it.' }, null, true]);

  // Stored details that cannot be used leave an existing section alone.
  const d = report();
  d.save(DETAILS);
  const shown = d.snapshot();
  d.docProps.setProperty(KEY, '{"v":1,');
  check('stored details that cannot be read: a refresh does not replace what the report shows', [d.s.refreshAdhocOpeningSection_(d.body), d.snapshot() === shown], [{ container: 'skipped', error: 'The stored opening details cannot be read.' }, true]);
  const viaSessions = d.s.saveAdhocSessionsConfiguration([session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), S2, S3].map((x) => Object.assign({}, x)));
  check('nor does saving the sessions; the result says that the section was left as it is', [viaSessions.ok, viaSessions.notice, d.snapshot() === shown], [true, 'The sessions were saved. The Session administration section was left as it is (The stored opening details cannot be read.).', true]);
  check('saving the dialog replaces the unreadable value and writes the section', [d.save({ s1: D('Alex Organizer'), s2: D(), s3: D() }), d.container().slice(0, 3)], [{ ok: true, errors: [], changed: true }, [HEADING, 'A01: Kick-off, September 22, 2026, 15:00–18:00', 'Chair: Alex Organizer']]);
  d.docProps.setProperty(KEY, '{"v":2,"sessions":{}}');
  const newerShown = d.snapshot();
  check('details of a newer version: nothing is overwritten, the section is left alone', [d.save(DETAILS).ok, d.s.refreshAdhocOpeningSection_(d.body).container, d.snapshot() === newerShown, d.docProps.getProperty(KEY)], [false, 'skipped', true, '{"v":2,"sessions":{}}']);

  // A failure while the sessions are saved.
  const e = report();
  e.save(DETAILS);
  e.s.renderAdhocOpeningSection_ = () => { throw new Error('document is busy'); };
  const result = e.s.saveAdhocSessionsConfiguration([session('s1', '2026-09-22', '15:00', '18:00', 'Renamed'), S2, S3].map((x) => Object.assign({}, x)));
  check('the section cannot be written while the sessions are saved: the sessions are saved all the same, and the result says how to write it',
    [result.ok, result.changed, result.notice, e.s.getAdhocSessions_()[0].label], [true, true, 'The sessions were saved. The Session administration section could not be written (document is busy). Save Edit Opening Details… again to write it.', 'Renamed']);
}

// ================================================================ 9. menu, adoption, privacy

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
  return { loaded: loaded, sub: sub, text: JSON.stringify(ui.menus), docProps: loaded.docProps };
}

console.log('9. menu, adoption, privacy, and nothing else changed');
{
  const adhoc = openMenu(Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, WITH_SESSIONS));
  check('ad-hoc report: the submenu, with Edit Opening Details… as its fifth item', adhoc.sub.entries.slice(0, 5), [{ label: 'Configure Sessions…', fn: 'configureAdhocSessions' }, { label: 'Import Teams Attendance…', fn: 'importTeamsAttendance' },
    { label: 'Refresh Attendance Section', fn: 'refreshAdhocAttendanceSection' }, { label: 'Assign TDoc Sessions…', fn: 'assignAdhocTdocSessions' }, { label: 'Edit Opening Details…', fn: 'editAdhocOpeningDetails' }]);
  check('its function exists and is public; there is no separate refresh action for the opening', [typeof adhoc.loaded.sandbox.editAdhocOpeningDetails, /_$/.test('editAdhocOpeningDetails'), /Refresh Opening/.test(adhoc.text)], ['function', false, false]);
  check('main-meeting report, new report, master template, CENTRAL / Legacy: no such item', [openMenu(Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' })).text, openMenu({}).text, openMenu(WITH_SESSIONS, { docId: TEMPLATE_ID }).text, openMenu(WITH_SESSIONS, { release: false }).text].map((t) => /Opening Details/.test(t)), [false, false, false, false]);
  check('opening a report writes nothing', openMenu(Object.assign({}, WITH_SESSIONS, { [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) })).docProps._store, Object.assign({}, WITH_SESSIONS, { [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) }));

  // Adoption: not copied.
  const r = report();
  r.save(DETAILS);
  const adopted = r.s.adoptReportDocumentForAddon_();
  check('adoption copies the sessions and not the opening details: background updates never read them', [adopted.verified, adopted.copiedKeys.filter((k) => /ADHOC/.test(k)), r.scriptProps.getKeys().filter((k) => /OPENING/.test(k)), r.docProps.getProperty(KEY) !== null], [true, ['ADHOC_SESSIONS'], [], true]);
  check('it is neither an adoption key nor in an adoption key family', [r.s.ADDON003_ADOPTION_FIXED_KEYS_.filter((k) => /OPENING/.test(k)), r.s.ADDON003_ADOPTION_PREFIX_FAMILIES_.filter((p) => KEY.indexOf(p) === 0)], [[], []]);
  check('the details are read from and written to the document\'s own properties (source)', [/function adhocOpeningStore_\(\) \{\n  return PropertiesService\.getDocumentProperties\(\);\n\}/.test(CODE),
    (CODE.slice(CODE.indexOf('// AD-HOC OPENING (stage E)'), CODE.indexOf('// AD-HOC STATUS (stage F)')).match(/getReportStateStore_|getScriptProperties/g) || [])], [true, []]);
  check('Clear Collection Caches keeps the details', (() => { const before = r.docProps.getProperty(KEY); r.s.clearAllCaches(); return r.docProps.getProperty(KEY) === before; })(), true);
  check('a copy of the report (empty properties) has no details and keeps the section as plain text', (() => { const copy = report({ props: ADHOC }); return [copy.docProps.getProperty(KEY), copy.s.adhocOpeningUnavailable_()]; })(), [null, 'Opening details are entered per session. Use Configure Sessions first.']);

  // Privacy.
  const p = report();
  const b64 = FX.utf16le(FX.buildExport()).toString('base64');
  p.s.confirmTeamsAttendanceImport('s1', b64, p.s.previewTeamsAttendanceImport('s1', b64).token);
  p.save({ s1: D('Alex Organizer', 'Sam Rivera', 'A sensitive administrative note.'), s2: D(), s3: D() });
  p.s.editAdhocOpeningDetails();
  check('a normal save and dialog log nothing', p.logs, []);
  check('the names and the note are in the opening property, in the section and in this dialog -- and nowhere else', (() => {
    const others = Object.keys(p.docProps._store).filter((k) => k !== KEY).map((k) => p.docProps._store[k]).join('\n');
    const outside = withoutContainer(p.snapshot());
    p.s.importTeamsAttendance();
    p.s.configureAdhocSessions();
    const otherDialogs = p.ui.dialogs.filter((x) => x.title !== 'Edit Opening Details').map((x) => x.html).join('\n');
    return [/A sensitive administrative note/.test(others), /A sensitive administrative note|Chair: /.test(outside), /A sensitive administrative note|Sam Rivera/.test(otherDialogs), /A sensitive administrative note/.test(p.docProps.getProperty(KEY)), p.ui.dialogs.filter((x) => x.title !== 'Edit Opening Details').length];
  })(), [false, false, false, true, 2]);
  check('the Chair is not matched against the attendance: an attendee of that name gets no company, no e-mail is added to the Chair', (() => {
    const attendeeTable = p.body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getCell(0).getText() === 'Name')[0];
    const row = rowsOf(attendeeTable).filter((x) => x[0] === 'Alex Organizer')[0];
    return [row, p.container().filter((t) => /^Chair/.test(t)), /@/.test(p.docProps.getProperty(KEY))];
  })(), [['Alex Organizer', '', 'alex.organizer@example.com'], ['Chair: Alex Organizer'], false]);
  check('the opening code logs fixed texts and error messages only, and its messages never quote a value (source)', (() => {
    const section = CODE.slice(CODE.indexOf('// AD-HOC OPENING (stage E)'), CODE.indexOf('// AD-HOC STATUS (stage F)'));
    const logs = (section.match(/Logger\.log\((.*)\);/g) || []).filter((line) => !/^Logger\.log\('Opening details: [^']*' \+ e\.message\);$/.test(line));
    const errorLines = (section.match(/errors\.push\((.*)\);/g) || []).filter((line) => /entry\[|input\[|\.chair|\.note|minuteTakers/.test(line.replace(/entry\[field\.key\]\.length/, '')));
    return [logs, errorLines];
  })(), [[], []]);
  check('the section uses no network, no clock, and nothing of the attendance or the parser (source)', /UrlFetchApp|Date\.now|new Date\(|readAdhocAttendance_|parseTeamsAttendanceReport_|normalizeTeams|effectiveAdhocAttendees_/.test(CODE.slice(CODE.indexOf('// AD-HOC OPENING (stage E)'), CODE.indexOf('// AD-HOC STATUS (stage F)')).replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '')), false);

  // Main-meeting and session-less reports.
  [['a main-meeting report', Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main', [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) })], ['an ad-hoc report without sessions', Object.assign({}, ADHOC, { [KEY]: stored({ s1: { chair: 'Alex Organizer' } }) })]].forEach(([what, props]) => {
    const x = report({ props: props });
    const snap = x.snapshot();
    x.s.editAdhocOpeningDetails();
    x.save(DETAILS);
    check(`${what}, even with details left in its properties: nothing is rendered, nothing is written`, [x.snapshot() === snap, x.writes, x.ui.dialogs.length, x.docProps.getProperty(KEY)], [true, [], 0, stored({ s1: { chair: 'Alex Organizer' } })]);
  });
  const section = CODE.slice(CODE.indexOf('// AD-HOC SESSIONS (stage A)'), CODE.indexOf('// ARCH-009 -- MEETING-ID RESOLVER CORE'));
  check('outside the ad-hoc sections (and the changelog), Code.js mentions the opening details in the one build hook only', (CODE.replace(section, '').replace(/ \* 2\.18\.0 \(2026-10-02\)\n[\s\S]*? \* 2\.17\.4 \(2026-10-01\)\n/, ' * 2.17.4 (2026-10-01)\n').match(/[A-Za-z_]*AdhocOpening[A-Za-z_]*|ADHOC_[A-Z_]*OPENING[A-Z_]*/g) || []), ['finishAdhocOpeningRebuild_']);
  check('Code.js is version 2.22.0 (2.18.0 added the ad-hoc sessions, 2.18.1 the TDoc upload completion, 2.19.0 the status dropdowns, 2.20.0 their conversion in an existing report, 2.21.0 the ad-hoc meeting automation and the personal Reviewer token, 2.21.1 the dropdown status in a discussion e-mail, 2.22.0 the shared minutes)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.22.0');
  check('ReportCreator.js: one more menu item and the build question, nothing else', (CREATOR.match(/Opening Details|editAdhocOpeningDetails|adhocOpening|ADHOC_OPENING/g) || []).length, 5);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll ad-hoc opening checks passed.');
process.exitCode = failures ? 1 : 0;
