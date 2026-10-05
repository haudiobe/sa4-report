/**
 * Ad-hoc sessions, stage A -- the session model, its persistence and the
 * Configure Sessions dialog (docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md).
 *
 *   1. Dates, times, labels and chronological order.
 *   2. Validation.
 *   3. Reading the ADHOC_SESSIONS property: robust, never throws.
 *   4. Editing: stable ids, ids never reused.
 *   5. When sessions are active: ad-hoc only, valid configuration only.
 *   6. Saving: nothing is written unless something valid changed.
 *   7. The dialog: opening writes nothing; its client script.
 *   8. Menu visibility.
 *   9. Adoption.
 *  10. Nothing else changed: reports without sessions, main meetings, the
 *      CENTRAL / Legacy menu, a malformed property.
 *
 * All data is synthetic.
 *
 * Run: node tests/adhoc-sessions-config.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
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
const thrown = (fn) => { try { fn(); return null; } catch (e) { return e.message; } };

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
const KEY = 'ADHOC_SESSIONS';

const S = loadCode().sandbox;
const session = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const stored = (nextId, sessions) => JSON.stringify({ v: 1, nextId: nextId, sessions: sessions });
const ids = (list) => list.map((s) => s.id);

// ================================================================ 1. dates, times, labels, order

console.log('1. dates, times, labels, order');
{
  check('valid dates', ['2026-09-22', '2028-02-29', '2026-12-31'].map(S.isValidAdhocSessionDate_), [true, true, true]);
  check('invalid dates', ['', null, undefined, '2026-9-22', '22.09.2026', '2026-02-30', '2026-13-01', '2026-00-10', '2027-02-29', '2026-09-22T10:00', ' 2026-09-22']
    .map(S.isValidAdhocSessionDate_), [false, false, false, false, false, false, false, false, false, false, false]);
  check('valid times', ['00:00', '09:05', '15:00', '23:59'].map(S.isValidAdhocSessionTime_), [true, true, true, true]);
  check('invalid times', ['', null, '9:05', '24:00', '15:60', '15', '15:00:00', '3pm'].map(S.isValidAdhocSessionTime_), [false, false, false, false, false, false, false, false]);

  check('default label: day and month', ['2026-09-22', '2026-10-01', '2026-01-09', '2026-12-31'].map(S.defaultAdhocSessionLabel_), ['22 Sep', '1 Oct', '9 Jan', '31 Dec']);
  check('default label of an invalid date is empty', S.defaultAdhocSessionLabel_('2026-02-30'), '');
  check('the label shown: the user\'s own, else the default', [S.adhocSessionLabel_(session('s1', '2026-09-22')), S.adhocSessionLabel_(session('s1', '2026-09-22', '', '', 'Kick-off'))], ['22 Sep', 'Kick-off']);

  const unordered = [session('s1', '2026-09-24'), session('s2', '2026-09-22', '15:00'), session('s3', '2026-09-22', '09:00'), session('s4', '2026-09-23')];
  check('order: date, then planned start', ids(S.sortAdhocSessions_(unordered)), ['s3', 's2', 's4', 's1']);
  check('the order does not depend on the input order', ids(S.sortAdhocSessions_(unordered.slice().reverse())), ['s3', 's2', 's4', 's1']);
  check('sorting does not change its argument', ids(unordered), ['s1', 's2', 's3', 's4']);
  check('equal date and start: the id decides, numerically', ids(S.sortAdhocSessions_([session('s10', '2026-09-22'), session('s2', '2026-09-22')])), ['s2', 's10']);
}

// ================================================================ 2. validation

console.log('2. validation');
{
  const v = (list) => S.validateAdhocSessions_(list);
  const errorsOf = (list) => v(list).errors;

  check('one session without times', [errorsOf([session('s1', '2026-09-22')]), v([session('s1', '2026-09-22')]).sessions],
    [[], [session('s1', '2026-09-22')]]);
  check('an empty list is valid', v([]), { errors: [], sessions: [] });
  check('several sessions on different dates, times optional', errorsOf([session('s1', '2026-09-22', '15:00', '18:00'), session('s2', '2026-09-23'), session('s3', '2026-09-24', '15:00')]), []);
  check('a start alone, an end alone', errorsOf([session('s1', '2026-09-22', '15:00'), session('s2', '2026-09-23', '', '18:00')]), []);
  check('the result is in chronological order', ids(v([session('s1', '2026-09-24'), session('s2', '2026-09-22')]).sessions), ['s2', 's1']);
  check('values are trimmed, label white space is collapsed',
    v([{ id: ' s1 ', label: '  Call \n one ', date: ' 2026-09-22 ', start: ' 15:00 ', end: null }]).sessions, [session('s1', '2026-09-22', '15:00', '', 'Call one')]);
  check('many sessions are accepted (no fixed limit)',
    errorsOf(Array.from({ length: 30 }, (_, i) => session('s' + (i + 1), '2026-10-' + String(i + 1).padStart(2, '0')))), []);

  check('not a list', errorsOf(null), ['The session list is missing.']);
  check('not a session', errorsOf(['x']), ['Session 1 is not a session.']);
  check('missing date', errorsOf([session('s1', '')]), ['Session 1: enter a date.']);
  check('invalid date', errorsOf([session('s1', '2026-02-30')]), ['2026-02-30: "2026-02-30" is not a date (YYYY-MM-DD).']);
  check('invalid start', errorsOf([session('s1', '2026-09-22', '25:00')]), ['2026-09-22: the planned start "25:00" is not a time (HH:mm).']);
  check('invalid end', errorsOf([session('s1', '2026-09-22', '', '6pm')]), ['2026-09-22: the planned end "6pm" is not a time (HH:mm).']);
  check('end before start', errorsOf([session('s1', '2026-09-22', '15:00', '14:00')]), ['2026-09-22: the planned end must be after the planned start.']);
  check('end equal to start', errorsOf([session('s1', '2026-09-22', '15:00', '15:00')]), ['2026-09-22: the planned end must be after the planned start.']);
  check('a label longer than 16 characters', errorsOf([session('s1', '2026-09-22', '', '', 'A label that is too long')]), ['A label that is too long: the label is longer than 16 characters.']);
  check('a label of 16 characters', errorsOf([session('s1', '2026-09-22', '', '', '1234567890123456')]), []);

  check('missing id', errorsOf([session('', '2026-09-22')]), ['2026-09-22: the session id is missing or malformed.']);
  check('malformed ids', ['S1', 's0', 's01', '1', 's1a', 's-1'].map((id) => errorsOf([session(id, '2026-09-22')]).length), [1, 1, 1, 1, 1, 1]);
  check('duplicate id', errorsOf([session('s1', '2026-09-22'), session('s1', '2026-09-23')]), ['2026-09-23: the session id s1 is used twice.']);

  check('duplicate labels', errorsOf([session('s1', '2026-09-22', '', '', 'Call'), session('s2', '2026-09-23', '', '', 'call')]),
    ['The label "call" is used by more than one session. Give each session its own label.']);
  check('a label equal to another session\'s default label', errorsOf([session('s1', '2026-09-22'), session('s2', '2026-09-23', '', '', '22 Sep')]).length, 1);

  // Same day: the order must be known without guessing.
  const twoOnOneDay = (a, b) => errorsOf([session('s1', '2026-09-22', a[0], a[1], 'A'), session('s2', '2026-09-22', b[0], b[1], 'B')]);
  check('same day, no start on either', twoOnOneDay(['', ''], ['', '']), ['Two sessions are on 22 Sep: enter a planned start for each of them, so that their order is known.']);
  check('same day, a start on one only', twoOnOneDay(['09:00', '10:00'], ['', '']).length, 1);
  check('same day, the same start', twoOnOneDay(['09:00', '10:00'], ['09:00', '11:00']), ['Two sessions on 22 Sep have the same planned start (09:00).']);
  check('same day, the earlier one has no end', twoOnOneDay(['09:00', ''], ['15:00', '']), ['A: enter a planned end, because another session follows on the same day.']);
  check('same day, overlapping', twoOnOneDay(['09:00', '16:00'], ['15:00', '']), ['A: the planned end (16:00) is after the start of the next session on that day (15:00).']);
  check('same day, one after the other', twoOnOneDay(['09:00', '12:00'], ['15:00', '']), []);
  check('same day, back to back', twoOnOneDay(['09:00', '15:00'], ['15:00', '18:00']), []);
  check('same day, entered in reverse order', ids(S.validateAdhocSessions_([session('s1', '2026-09-22', '15:00', '', 'B'), session('s2', '2026-09-22', '09:00', '12:00', 'A')]).sessions), ['s2', 's1']);
  check('three on one day without times: one message, not one per pair',
    errorsOf([session('s1', '2026-09-22', '', '', 'A'), session('s2', '2026-09-22', '', '', 'B'), session('s3', '2026-09-22', '', '', 'C')]).length, 1);
  check('several problems are all reported', errorsOf([session('s1', ''), session('s2', '2026-09-23', '99:00')]).length, 2);
  check('with errors no session list is returned', v([session('s1', '')]).sessions, []);
}

// ================================================================ 3. reading the property

console.log('3. reading the property');
{
  const p = S.parseAdhocSessionsProperty_;
  check('no value', [p(null).status, p(undefined).status, p('').status, p('   ').status], ['absent', 'absent', 'absent', 'absent']);

  const one = stored(2, [session('s1', '2026-09-22', '15:00', '18:00')]);
  check('one session', p(one), { status: 'ok', config: { v: 1, nextId: 2, sessions: [session('s1', '2026-09-22', '15:00', '18:00')] }, error: null });
  check('sessions are returned in chronological order, whatever the stored order',
    ids(p(stored(4, [session('s3', '2026-09-24'), session('s1', '2026-09-22'), session('s2', '2026-09-23')])).config.sessions), ['s1', 's2', 's3']);
  check('no sessions, counter kept', p(stored(5, [])), { status: 'ok', config: { v: 1, nextId: 5, sessions: [] }, error: null });
  check('a counter above the highest id is fine (ids were removed)', p(stored(9, [session('s2', '2026-09-22')])).status, 'ok');

  check('malformed JSON', [p('{"v":1,').status, p('not json').status, p('{"v":1,').error], ['invalid', 'invalid', 'The stored session configuration is not valid JSON.']);
  check('JSON that is not an object', ['[]', '"x"', '12', 'null', 'true'].map((raw) => p(raw).status), ['invalid', 'invalid', 'invalid', 'invalid', 'invalid']);
  check('unsupported schema version', [p(JSON.stringify({ v: 2, nextId: 2, sessions: [] })).status, p(JSON.stringify({ nextId: 2, sessions: [] })).status, p(JSON.stringify({ v: '1', nextId: 2, sessions: [] })).status],
    ['unsupported', 'unsupported', 'unsupported']);
  check('its message names both versions', p(JSON.stringify({ v: 2, nextId: 2, sessions: [] })).error, 'The stored session configuration has version 2; this release reads version 1.');
  check('no session list', p(JSON.stringify({ v: 1, nextId: 2 })).status, 'invalid');
  check('an invalid session', p(stored(2, [session('s1', '2026-02-30')])).status, 'invalid');
  check('duplicate ids', p(stored(3, [session('s1', '2026-09-22'), session('s1', '2026-09-23')])).status, 'invalid');
  check('sessions that cannot be ordered', p(stored(3, [session('s1', '2026-09-22', '', '', 'A'), session('s2', '2026-09-22', '', '', 'B')])).status, 'invalid');
  check('a counter that would reuse an id', [p(stored(1, [session('s1', '2026-09-22')])).status, p(stored(2, [session('s2', '2026-09-22')])).status], ['invalid', 'invalid']);
  check('a counter that is not a whole number', ['"3"', '2.5', 'null'].map((n) => p('{"v":1,"nextId":' + n + ',"sessions":[]}').status), ['invalid', 'invalid', 'invalid']);
  check('a missing counter', p('{"v":1,"sessions":[]}').status, 'invalid');
  check('it never throws', thrown(() => [{}, [], 5, true, '{', '\u0000'].forEach((raw) => p(raw))), null);

  const config = { v: 1, nextId: 4, sessions: [session('s3', '2026-09-24'), session('s1', '2026-09-22', '15:00', '18:00', 'Call 1')] };
  check('the stored form: fixed key order, chronological', S.serializeAdhocSessions_(config),
    '{"v":1,"nextId":4,"sessions":[{"id":"s1","label":"Call 1","date":"2026-09-22","start":"15:00","end":"18:00"},{"id":"s3","label":"","date":"2026-09-24","start":"","end":""}]}');
  check('what is written is read back unchanged', p(S.serializeAdhocSessions_(config)).config, { v: 1, nextId: 4, sessions: [config.sessions[1], config.sessions[0]] });
  check('nothing but the schema fields is stored (no order, no attendance)',
    Object.keys(JSON.parse(S.serializeAdhocSessions_({ v: 1, nextId: 2, sessions: [Object.assign({ order: 1, actualStart: '14:54', displayLabel: 'x' }, session('s1', '2026-09-22'))] })).sessions[0]),
    ['id', 'label', 'date', 'start', 'end']);
}

// ================================================================ 4. editing

console.log('4. editing: stable ids');
{
  const edit = S.applyAdhocSessionsEdit_;
  const row = (id, date, start, end, label) => session(id, date, start, end, label);

  const first = edit(null, [row('', '2026-09-22', '15:00', '18:00')]);
  check('first session: id s1, counter 2', [first.ok, first.config], [true, { v: 1, nextId: 2, sessions: [session('s1', '2026-09-22', '15:00', '18:00')] }]);

  const raw1 = S.serializeAdhocSessions_(first.config);
  const second = edit(raw1, [row('s1', '2026-09-22', '15:00', '18:00'), row('', '2026-09-24'), row('', '2026-09-23')]);
  check('added sessions get the next ids, in the order of the rows', [ids(second.config.sessions), second.config.nextId,
    second.config.sessions.map((s) => s.date)], [['s1', 's3', 's2'], 4, ['2026-09-22', '2026-09-23', '2026-09-24']]);

  const raw2 = S.serializeAdhocSessions_(second.config);
  const edited = edit(raw2, [row('s1', '2026-09-25', '10:00', '12:00', 'Moved'), row('s2', '2026-09-24'), row('s3', '2026-09-23')]);
  check('editing date, times and label keeps the id', edited.config.sessions.filter((s) => s.id === 's1'), [session('s1', '2026-09-25', '10:00', '12:00', 'Moved')]);
  check('the order follows the new date; the counter is unchanged', [ids(edited.config.sessions), edited.config.nextId], [['s3', 's2', 's1'], 4]);

  const removed = edit(raw2, [row('s1', '2026-09-22'), row('s3', '2026-09-23')]);
  check('removing a session keeps the others and the counter', [ids(removed.config.sessions), removed.config.nextId], [['s1', 's3'], 4]);
  const raw3 = S.serializeAdhocSessions_(removed.config);
  const again = edit(raw3, [row('s1', '2026-09-22'), row('s3', '2026-09-23'), row('', '2026-09-24')]);
  check('a new session after a removal gets a new id, not the removed one', [ids(again.config.sessions), again.config.nextId], [['s1', 's3', 's4'], 5]);

  const emptied = edit(raw3, []);
  check('removing every session keeps the counter', emptied.config, { v: 1, nextId: 4, sessions: [] });
  const afterEmpty = edit(S.serializeAdhocSessions_(emptied.config), [row('', '2026-10-01')]);
  check('and the next session still gets a new id', [ids(afterEmpty.config.sessions), afterEmpty.config.nextId], [['s4'], 5]);

  check('a row with an id that is not stored is refused', edit(raw1, [row('s7', '2026-09-22')]),
    { ok: false, errors: ['A session of this dialog is no longer stored. Close the dialog and open Configure Sessions again.'], config: null });
  check('an id cannot be claimed in a report without sessions', edit(null, [row('s1', '2026-09-22')]).ok, false);
  check('the same id twice is refused', edit(raw2, [row('s1', '2026-09-22'), row('s1', '2026-09-23')]).ok, false);
  check('invalid rows are refused with their messages', edit(raw1, [row('s1', '2026-09-22', '18:00', '15:00')]),
    { ok: false, errors: ['2026-09-22: the planned end must be after the planned start.'], config: null });
  check('rows that are not a list', edit(raw1, null).ok, false);
  check('the edit does not change its input', raw1, S.serializeAdhocSessions_(first.config));

  // A stored value that cannot be read.
  const damaged = '{"v":1,"nextId":6,"sessions":[{"id":"s2","date":"2026-02-30"},{"id":"s5","date":';
  check('replacing a damaged configuration does not reuse the ids found in it', [S.salvageAdhocSessionNextId_(damaged), ids(edit(damaged, [row('', '2026-09-22')]).config.sessions)], [6, ['s6']]);
  check('without a counter in the text: above the highest id found', S.salvageAdhocSessionNextId_('{"sessions":[{"id":"s2"},{"id":"s9"}]'), 10);
  check('nothing to find: start at 1', [S.salvageAdhocSessionNextId_('garbage'), S.salvageAdhocSessionNextId_(null)], [1, 1]);
  check('a damaged configuration offers no ids to keep', edit(damaged, [row('s2', '2026-09-22')]).ok, false);

  const newer = JSON.stringify({ v: 2, nextId: 3, sessions: [] });
  check('a configuration of a newer schema version is never replaced', edit(newer, [row('', '2026-09-22')]),
    { ok: false, errors: ['The stored session configuration has version 2; this release reads version 1. It was written by a newer release and is not changed here.'], config: null });
}

// ================================================================ 5. when sessions are active

console.log('5. when sessions are active');
{
  const VALID = stored(3, [session('s2', '2026-09-23', '', '', 'Offline'), session('s1', '2026-09-22', '15:00', '18:00')]);
  const state = (props) => {
    const s = loadCode({ documentProperties: props }).sandbox;
    return [s.isAdhocMeetingForSessions_(), s.adhocSessionsEnabled_(), s.getAdhocSessions_().length];
  };
  check('ad-hoc with a valid configuration: active', state({ MEETING_TYPE: 'adhoc', [KEY]: VALID }), [true, true, 2]);
  check('ad-hoc without the property: not active', state({ MEETING_TYPE: 'adhoc' }), [true, false, 0]);
  check('ad-hoc, meeting type written differently', state({ MEETING_TYPE: ' AdHoc ', [KEY]: VALID }), [true, true, 2]);
  check('main meeting with the property: not active', state({ MEETING_TYPE: 'main', [KEY]: VALID }), [false, false, 0]);
  check('no meeting type (a main meeting) with the property: not active', state({ [KEY]: VALID }), [false, false, 0]);
  check('an unreadable meeting type: not active, no error', state({ MEETING_TYPE: 'electronic', [KEY]: VALID }), [false, false, 0]);
  check('ad-hoc, configuration without sessions: not active', state({ MEETING_TYPE: 'adhoc', [KEY]: stored(3, []) }), [true, false, 0]);
  check('ad-hoc, malformed JSON: not active', state({ MEETING_TYPE: 'adhoc', [KEY]: '{"v":1,' }), [true, false, 0]);
  check('ad-hoc, unsupported version: not active', state({ MEETING_TYPE: 'adhoc', [KEY]: JSON.stringify({ v: 2, nextId: 2, sessions: [session('s1', '2026-09-22')] }) }), [true, false, 0]);
  check('ad-hoc, invalid session: not active', state({ MEETING_TYPE: 'adhoc', [KEY]: stored(2, [session('s1', '2026-09-22', '18:00', '15:00')]) }), [true, false, 0]);

  const s = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', [KEY]: VALID } }).sandbox;
  check('sessions in order, with the label to show', s.getAdhocSessions_(), [
    Object.assign(session('s1', '2026-09-22', '15:00', '18:00'), { displayLabel: '22 Sep' }),
    Object.assign(session('s2', '2026-09-23', '', '', 'Offline'), { displayLabel: 'Offline' })
  ]);

  // A background run reads the document's central state, like every other setting.
  const central = loadCode({ scriptProperties: { 'SA4_STATE|DOC1|MEETING_TYPE': 'adhoc', ['SA4_STATE|DOC1|' + KEY]: VALID } }).sandbox;
  check('a background context reads the central copy', [central.adhocSessionsEnabled_({ documentId: 'DOC1', mode: 'addon-background' }), central.adhocSessionsEnabled_()], [true, false]);

  const failing = loadCode().sandbox;
  failing.PropertiesService.getDocumentProperties = () => { throw new Error('no document'); };
  check('a store that cannot be read: not active, no error', [thrown(() => failing.adhocSessionsEnabled_()), failing.getAdhocSessions_()], [null, []]);
}

// ================================================================ 6. saving

console.log('6. saving');
{
  const open = (props) => loadCode({ documentProperties: props });
  const writes = (docProps) => { const log = []; const set = docProps.setProperty; docProps.setProperty = (k, v) => { log.push(k); set(k, v); }; return log; };

  const a = open({ MEETING_TYPE: 'adhoc' });
  const aWrites = writes(a.docProps);
  const saved = a.sandbox.saveAdhocSessionsConfiguration([session('', '2026-09-22', '15:00', '18:00'), session('', '2026-09-23')]);
  check('a valid save: ok, changed, sessions with their ids', [saved.ok, saved.changed, ids(saved.sessions), saved.errors], [true, true, ['s1', 's2'], []]);
  check('exactly one property was written', aWrites, [KEY]);
  check('the stored value', a.docProps.getProperty(KEY),
    '{"v":1,"nextId":3,"sessions":[{"id":"s1","label":"","date":"2026-09-22","start":"15:00","end":"18:00"},{"id":"s2","label":"","date":"2026-09-23","start":"","end":""}]}');
  check('the report now has sessions', a.sandbox.adhocSessionsEnabled_(), true);

  const sameRows = [session('s1', '2026-09-22', '15:00', '18:00'), session('s2', '2026-09-23')];
  const repeat = a.sandbox.saveAdhocSessionsConfiguration(sameRows);
  check('saving the same again: ok, nothing written', [repeat.ok, repeat.changed, aWrites.length], [true, false, 1]);
  const reordered = a.sandbox.saveAdhocSessionsConfiguration(sameRows.slice().reverse());
  check('the same sessions in another row order: nothing written', [reordered.changed, aWrites.length], [false, 1]);

  const before = a.docProps.getProperty(KEY);
  const refused = a.sandbox.saveAdhocSessionsConfiguration([session('s1', '2026-09-22', '18:00', '15:00'), session('s2', '2026-09-23')]);
  check('an invalid save: refused with a message, nothing written, the stored value untouched',
    [refused.ok, refused.changed, refused.errors.length, aWrites.length, a.docProps.getProperty(KEY) === before], [false, false, 1, 1, true]);

  const removedAll = a.sandbox.saveAdhocSessionsConfiguration([]);
  check('removing every session: the property stays, with its counter; sessions are off',
    [removedAll.ok, removedAll.changed, a.docProps.getProperty(KEY), a.sandbox.adhocSessionsEnabled_()], [true, true, '{"v":1,"nextId":3,"sessions":[]}', false]);
  check('a session added afterwards gets a new id', ids(a.sandbox.saveAdhocSessionsConfiguration([session('', '2026-10-01')]).sessions), ['s3']);

  const fresh = open({ MEETING_TYPE: 'adhoc' });
  const freshWrites = writes(fresh.docProps);
  const none = fresh.sandbox.saveAdhocSessionsConfiguration([]);
  check('saving "no sessions" in a report that never had any creates no property', [none.ok, none.changed, freshWrites, fresh.docProps.getKeys()], [true, false, [], ['MEETING_TYPE']]);

  const main = open({ MEETING_TYPE: 'main' });
  const mainWrites = writes(main.docProps);
  check('a main-meeting report: refused, nothing written', [main.sandbox.saveAdhocSessionsConfiguration([session('', '2026-09-22')]), mainWrites],
    [{ ok: false, errors: ['Sessions are available for ad-hoc reports only.'], changed: false, sessions: [] }, []]);
  const untyped = open({});
  check('a report without a meeting type: refused, nothing written', [untyped.sandbox.saveAdhocSessionsConfiguration([session('', '2026-09-22')]).ok, untyped.docProps.getKeys()], [false, []]);

  const damaged = open({ MEETING_TYPE: 'adhoc', [KEY]: 'garbage' });
  check('a damaged value is replaced by a valid save', [damaged.sandbox.saveAdhocSessionsConfiguration([session('', '2026-09-22')]).ok, damaged.sandbox.adhocSessionsEnabled_()], [true, true]);
  const newer = open({ MEETING_TYPE: 'adhoc', [KEY]: JSON.stringify({ v: 2, sessions: [] }) });
  check('a value of a newer version is not replaced', [newer.sandbox.saveAdhocSessionsConfiguration([session('', '2026-09-22')]).ok, newer.docProps.getProperty(KEY)], [false, '{"v":2,"sessions":[]}']);

  const master = loadTemplateRuntime({ release: { releaseId: 'T-2026.10.4', flavor: 'template', templateDocumentId: 'TEMPLATEdoc0000000000000000000000000000000' }, documentProperties: { MEETING_TYPE: 'adhoc' } });
  master.sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => 'TEMPLATEdoc0000000000000000000000000000000' });
  check('in the master template the save refuses', [/SA4 Report Template itself/.test(thrown(() => master.sandbox.saveAdhocSessionsConfiguration([session('', '2026-09-22')])) || ''), master.docProps.getProperty(KEY)], [true, null]);
}

// ================================================================ 7. the dialog

/** A sandbox with a UI and HtmlService. */
function withUi(props, loader) {
  const loaded = (loader || loadCode)({ documentProperties: props });
  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK' } };
  ui.alert = (...args) => { ui.alerts.push(args); return 'OK'; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title: title, html: out.html, width: out.width, height: out.height }); };
  loaded.sandbox.DocumentApp.getUi = () => ui;
  loaded.sandbox.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: (w) => { out.width = w; return out; }, setHeight: (h) => { out.height = h; return out; } }; return out; } };
  return { s: loaded.sandbox, ui: ui, docProps: loaded.docProps };
}

/** Runs the dialog's client script against a minimal DOM; returns what it did. */
function runClient(html) {
  const scripts = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
  const element = (tag) => {
    const e = { tag: tag, children: [], style: {}, value: '', textContent: '', disabled: false };
    e.appendChild = (c) => { e.children.push(c); return c; };
    e.removeChild = (c) => { e.children.splice(e.children.indexOf(c), 1); return c; };
    return e;
  };
  const byId = {};
  ['rows', 'empty', 'addBtn', 'saveBtn', 'errors', 'notice'].forEach((id) => { byId[id] = element(id); });
  const calls = { rpc: [], closed: 0 };
  let handlers = {};
  const runner = {
    withSuccessHandler: (fn) => { handlers.success = fn; return runner; },
    withFailureHandler: (fn) => { handlers.failure = fn; return runner; },
    saveAdhocSessionsConfiguration: (rows) => { calls.rpc.push(JSON.parse(JSON.stringify(rows))); calls.handlers = handlers; handlers = {}; }
  };
  const ctx = { document: { getElementById: (id) => byId[id], createElement: element }, google: { script: { run: runner, host: { close: () => { calls.closed++; } } } } };
  vm.createContext(ctx);
  vm.runInContext(scripts[0].replace(/^<script>|<\/script>$/g, ''), ctx);
  const inputsOf = (tr) => tr.children.slice(0, 5).map((td) => td.children[0]);
  return {
    ctx: ctx, byId: byId, calls: calls, scriptCount: scripts.length,
    // A row has five inputs: label, start date, end date, start time, end time. The checks written before a session
    // could span several days address the four it had then; the end date (the third input) has its own accessor.
    rowValues: () => byId.rows.children.map((tr) => inputsOf(tr).filter((x, k) => k !== 2).map((i) => i.value)),
    inputs: (i) => inputsOf(byId.rows.children[i]).filter((x, k) => k !== 2),
    endDate: (i) => inputsOf(byId.rows.children[i])[2],
    allInputs: (i) => inputsOf(byId.rows.children[i]),
    removeButton: (i) => byId.rows.children[i].children[5].children[0]
  };
}

console.log('7. the Configure Sessions dialog');
{
  const VALID = stored(4, [session('s1', '2026-09-22', '15:00', '18:00'), session('s3', '2026-09-24', '', '', 'Wrap-up')]);

  const r = withUi({ MEETING_TYPE: 'adhoc', MEETING_DATE: 'September 22, 2026', [KEY]: VALID });
  const keysBefore = JSON.stringify(r.docProps._store);
  r.s.configureAdhocSessions();
  check('one dialog, titled Configure Sessions, no alert', [r.ui.dialogs.length, r.ui.dialogs[0].title, r.ui.alerts.length], [1, 'Configure Sessions', 0]);
  check('opening the dialog changes no property', JSON.stringify(r.docProps._store), keysBefore);

  const client = runClient(r.ui.dialogs[0].html);
  check('the dialog has one script, and it runs', client.scriptCount, 1);
  check('the stored sessions are shown, in order: label, date, start, end', client.rowValues(), [['', '2026-09-22', '15:00', '18:00'], ['Wrap-up', '2026-09-24', '', '']]);
  check('input types: text, date, time, time', client.inputs(0).map((i) => i.type), ['text', 'date', 'time', 'time']);
  check('an empty label shows where its text comes from; its length is limited', [client.inputs(0)[0].placeholder, client.inputs(0)[0].maxLength], ['from the dates', 16]);
  check('no notice for a readable configuration', [client.byId.notice.textContent, client.byId.notice.style.display], ['', undefined]);
  check('nothing was sent by rendering', [client.calls.rpc.length, client.calls.closed], [0, 0]);

  // Edit, add, remove, then save.
  client.inputs(0)[0].value = 'Call 1';
  client.inputs(0)[3].value = '18:30';
  client.ctx.addSession();
  client.inputs(2)[1].value = '2026-09-23';
  client.removeButton(1).onclick();
  check('after adding one and removing one', client.rowValues(), [['Call 1', '2026-09-22', '15:00', '18:30'], ['', '2026-09-23', '', '']]);
  check('still nothing sent', client.calls.rpc.length, 0);
  client.ctx.saveSessions();
  check('Save sends the rows: kept ids, a blank id for the new session',
    client.calls.rpc, [[session('s1', '2026-09-22', '15:00', '18:30', 'Call 1'), session('', '2026-09-23')]]);
  check('Save is disabled while the request runs', client.byId.saveBtn.disabled, true);
  client.calls.handlers.success({ ok: false, errors: ['First problem.', 'Second problem.'] });
  check('a refused save: the messages are shown, the dialog stays open, Save is enabled again',
    [client.byId.errors.textContent, client.calls.closed, client.byId.saveBtn.disabled], ['First problem.\nSecond problem.', 0, false]);
  client.ctx.saveSessions();
  client.calls.handlers.failure(new Error('Server error'));
  check('a failed request: its message is shown', [client.byId.errors.textContent, client.calls.closed], ['Server error', 0]);
  client.ctx.saveSessions();
  client.calls.handlers.success({ ok: true, changed: true });
  check('an accepted save closes the dialog', client.calls.closed, 1);

  // The rows the client sends, saved by the real server function.
  const result = r.s.saveAdhocSessionsConfiguration(client.calls.rpc[0]);
  check('end to end: the edited sessions are stored; the removed id is not reused',
    [result.ok, JSON.parse(r.docProps.getProperty(KEY))], [true, { v: 1, nextId: 5, sessions: [session('s1', '2026-09-22', '15:00', '18:30', 'Call 1'), session('s4', '2026-09-23')] }]);

  // Cancel.
  const html = r.ui.dialogs[0].html;
  check('Cancel only closes the dialog (no server call in the button)',
    (html.match(/<button[^>]*>Cancel<\/button>/) || [''])[0], '<button type="button" style="background: #666;" onclick="google.script.host.close()">Cancel</button>');
  check('the only server function the dialog calls is the save', (html.match(/\.\s*([A-Za-z]+)\(collect\(\)\)/) || [])[1], 'saveAdhocSessionsConfiguration');
  const cancelled = withUi({ MEETING_TYPE: 'adhoc', [KEY]: VALID });
  cancelled.s.configureAdhocSessions();
  const c2 = runClient(cancelled.ui.dialogs[0].html);
  c2.inputs(0)[1].value = '2027-01-01';
  c2.removeButton(1).onclick();
  c2.ctx.google.script.host.close();
  check('editing and then cancelling sends nothing and stores nothing', [c2.calls.rpc.length, cancelled.docProps.getProperty(KEY)], [0, VALID]);

  // First use.
  const first = withUi({ MEETING_TYPE: 'adhoc', MEETING_DATE: 'September 22, 2026' });
  first.s.configureAdhocSessions();
  const c3 = runClient(first.ui.dialogs[0].html);
  check('first use: one session proposed from the meeting date, without times', c3.rowValues(), [['', '2026-09-22', '', '']]);
  check('it says that nothing is stored yet', [/nothing is stored until you save/.test(c3.byId.notice.textContent), c3.byId.notice.style.display], [true, 'block']);
  check('the proposal is not stored by opening the dialog', first.docProps.getKeys().sort(), ['MEETING_DATE', 'MEETING_TYPE']);
  const noDate = withUi({ MEETING_TYPE: 'adhoc' });
  noDate.s.configureAdhocSessions();
  const c4 = runClient(noDate.ui.dialogs[0].html);
  check('first use without a meeting date: no session, and a line saying so', [c4.rowValues(), /^No sessions\./.test(c4.byId.empty.textContent)], [[], true]);
  check('the model (pure): nothing proposed from an unreadable date', S.buildAdhocSessionsDialogModel_(null, 'next week').sessions, []);

  // Unreadable and newer configurations.
  const damaged = withUi({ MEETING_TYPE: 'adhoc', MEETING_DATE: 'September 22, 2026', [KEY]: '{"v":1,' });
  damaged.s.configureAdhocSessions();
  const c5 = runClient(damaged.ui.dialogs[0].html);
  check('a damaged configuration: no sessions shown, a notice, saving possible, nothing changed by opening',
    [c5.rowValues(), /cannot be read and are not used/.test(c5.byId.notice.textContent), c5.byId.saveBtn.disabled, damaged.docProps.getProperty(KEY)], [[], true, false, '{"v":1,']);
  const newer = withUi({ MEETING_TYPE: 'adhoc', [KEY]: JSON.stringify({ v: 2, nextId: 2, sessions: [] }) });
  newer.s.configureAdhocSessions();
  const c6 = runClient(newer.ui.dialogs[0].html);
  check('a newer schema version: a notice, Add and Save disabled', [/newer release/.test(c6.byId.notice.textContent), c6.byId.addBtn.disabled, c6.byId.saveBtn.disabled], [true, true, true]);

  // Stored text cannot leave the script element.
  const hostile = withUi({ MEETING_TYPE: 'adhoc', [KEY]: stored(2, [session('s1', '2026-09-22', '', '', '</script><b>')]) });
  hostile.s.configureAdhocSessions();
  const c7 = runClient(hostile.ui.dialogs[0].html);
  check('a label with markup stays data: one script element, shown as typed', [c7.scriptCount, c7.rowValues()[0][0], hostile.ui.dialogs[0].html.indexOf('</script><b>')], [1, '</script><b>', -1]);

  // Not an ad-hoc report.
  ['main', '', 'electronic'].forEach((type) => {
    const main = withUi(type ? { MEETING_TYPE: type } : {});
    main.s.configureAdhocSessions();
    check(`meeting type "${type}": a message, no dialog, no property`, [main.ui.dialogs.length, main.ui.alerts.map((a) => a[1]), main.docProps.getKeys().length],
      [0, ['Sessions are available for ad-hoc reports only.'], type ? 1 : 0]);
  });
}

// ================================================================ 8. menu visibility

const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.4', flavor: 'template', codeVersion: '0.0.0', gitCommit: 'abcdef0', gitTag: 'template-release/T-2026.10.4', templateDocumentId: TEMPLATE_ID };

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
  if (o.before) o.before(loaded);
  loaded.sandbox.onOpen();
  const top = ui.menus[0];
  const subs = top.entries.filter((e) => e.sub).map((e) => e.sub);
  const sessions = subs.filter((m) => /Sessions/.test(m.name));
  const all = JSON.stringify(ui.menus);
  return { loaded: loaded, top: top, subNames: subs.map((m) => m.name), sessions: sessions, mentionsSessions: /Session/.test(all) };
}

console.log('8. menu visibility');
{
  const adhoc = openMenu({ SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'adhoc' });
  check('ad-hoc template report: the submenu, after Document Reallocation', adhoc.subNames, ['📄 Report', '🔀 Document Reallocation', '🗓 Sessions and Attendance', '🛠 Advanced and Repair']);
  check('its first item is Configure Sessions… (the attendance items of stage D follow)', adhoc.sessions[0].entries[0], { label: 'Configure Sessions…', fn: 'configureAdhocSessions' });
  check('its function exists and is public', [typeof adhoc.loaded.sandbox.configureAdhocSessions, /_$/.test('configureAdhocSessions')], ['function', false]);
  check('ad-hoc without any session configured: the submenu is there (sessions are opt-in through it)', openMenu({ MEETING_TYPE: 'adhoc' }).sessions.length, 1);
  check('ad-hoc, written differently', openMenu({ MEETING_TYPE: ' AdHoc ' }).sessions.length, 1);

  const VALID = stored(2, [session('s1', '2026-09-22')]);
  check('main-meeting report: no submenu, no mention of sessions', [openMenu({ SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main' }).mentionsSessions, openMenu({ MEETING_TYPE: 'main' }).subNames],
    [false, ['📄 Report', '🔀 Document Reallocation', '🛠 Advanced and Repair']]);
  check('main-meeting report that holds the property: still no submenu', openMenu({ MEETING_TYPE: 'main', [KEY]: VALID }).mentionsSessions, false);
  check('no meeting type yet (a new report before its setup): no submenu', openMenu({}).mentionsSessions, false);
  check('no meeting type, but the property: no submenu', openMenu({ [KEY]: VALID }).mentionsSessions, false);
  check('an unknown meeting type: no submenu', openMenu({ MEETING_TYPE: 'electronic' }).mentionsSessions, false);
  check('the master template: no submenu', openMenu({ MEETING_TYPE: 'adhoc' }, { docId: TEMPLATE_ID }).mentionsSessions, false);

  const unreadable = openMenu({ MEETING_TYPE: 'adhoc' }, { before: (l) => {
    const real = l.sandbox.PropertiesService.getDocumentProperties;
    let calls = 0;
    // The first read is the menu head's; the second is the meeting type.
    l.sandbox.PropertiesService.getDocumentProperties = () => { calls++; if (calls > 1) throw new Error('no access'); return real(); };
  } });
  check('if the meeting type cannot be read the menu is built without the submenu', [unreadable.mentionsSessions, unreadable.subNames.length], [false, 3]);

  // Opening never writes.
  [{ MEETING_TYPE: 'adhoc' }, { MEETING_TYPE: 'adhoc', [KEY]: 'garbage' }, { MEETING_TYPE: 'main' }, {}].forEach((props) => {
    const m = openMenu(props);
    check(`opening a report (${JSON.stringify(props).slice(0, 40)}) creates and changes no property`, m.loaded.docProps._store, props);
  });

  // CENTRAL / Legacy: Code.js without Release.js.
  check('CENTRAL / Legacy menu of an ad-hoc document: no mention of sessions',
    [openMenu({ MEETING_TYPE: 'adhoc', [KEY]: VALID }, { release: false }).mentionsSessions, openMenu({ MEETING_TYPE: 'adhoc' }, { release: false }).top.name], [false, '⚠️Scripts⚠️']);
  const onOpenSource = CODE.slice(CODE.indexOf('\nfunction onOpen('), CODE.indexOf('\n}\n', CODE.indexOf('\nfunction onOpen(')));
  check('onOpen() in Code.js does not mention sessions (source)', /Session|ADHOC_/.test(onOpenSource), false);
  check('the submenu is built in ReportCreator.js only, behind the ad-hoc check (source)',
    [/createMenu\([^)]*Sessions and Attendance/.test(CODE), /if \(isAdhocReportForMenu_\(\)\) \{\n    menu\.addSubMenu\(ui\.createMenu\('🗓 Sessions and Attendance'\)/.test(CREATOR)], [false, true]);
}

// ================================================================ 9. adoption

console.log('9. adoption (copy into the central state)');
{
  const VALID = stored(2, [session('s1', '2026-09-22')]);
  const keys = (props) => { const l = loadCode({ documentProperties: props }); return l.sandbox.reportStateKeysToAdopt_(l.docProps); };
  check('ADHOC_SESSIONS is an adoption key', S.ADDON003_ADOPTION_FIXED_KEYS_.indexOf(KEY) !== -1, true);
  check('it is copied when the document has it', keys({ MEETING_ID: '86178', [KEY]: VALID }), ['MEETING_ID', KEY]);
  check('it is not invented when the document has none', keys({ MEETING_ID: '86178' }), ['MEETING_ID']);
  check('it is not a Configure Meeting key', S.CONFIG_DIALOG_MANAGED_KEYS_.indexOf(KEY), -1);
  check('the other adoption keys are the ones of the release before', S.ADDON003_ADOPTION_FIXED_KEYS_.filter((k) => k !== KEY && k !== 'ADHOC_TDOC_SESSIONS'), [
    'MEETING_FOLDER', 'MEETING_NUMBER', 'MEETING_ID', 'FTP_BASE', 'TDOC_LIST_URL', 'REPORT_SUFFIX', 'AGENDA_ITEM_PREFIX', 'AGENDA_SOURCE_DOC_ID', 'AGENDA_TDOC',
    'AGENDA_CSV_URL', 'MEETING_DATE', 'SHOW_PREVIEW_SNIPPET', 'MEETING_TYPE', 'MEETING_NAME', 'REVISIONS_URL', 'MAILING_LIST', 'FETCH_ABSTRACTS_ON_UPDATE',
    'REVISION_MAP', 'DEADLINE_EXTENSIONS']);

  const l = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', MEETING_ID: '86178', MEETING_NAME: 'Synthetic ad-hoc', [KEY]: VALID } });
  l.sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => 'DOC1' });
  const result = l.sandbox.adoptReportDocumentForAddon_();
  check('adoption copies the configuration and verifies it', [result.verified, result.copiedKeys.indexOf(KEY) !== -1, l.scriptProps.getProperty('SA4_STATE|DOC1|' + KEY)], [true, true, VALID]);
  check('the adopted document has sessions in a background run; the source is untouched',
    [l.sandbox.adhocSessionsEnabled_({ documentId: 'DOC1', mode: 'addon-background' }), l.docProps.getProperty(KEY)], [true, VALID]);
}

// ================================================================ 10. nothing else changed

function gitShow(spec) {
  try {
    return execFileSync('git', ['show', spec], { cwd: path.join(__dirname, '..'), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    return null;
  }
}
const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
function functionSource(src, name) {
  const start = src.indexOf('\nfunction ' + name + '(');
  if (start === -1) return null;
  return src.slice(start + 1, src.indexOf('\n}\n', start) + 2);
}

console.log('10. nothing else changed');
{
  // The release before this stage: Code.js 2.17.4.
  const BASELINE = 'template-release/T-2026.10.4';
  const OLD_CODE = gitShow(BASELINE + ':Code.js');
  const OLD_CREATOR = gitShow(BASELINE + ':template/ReportCreator.js');
  if (!OLD_CODE || !OLD_CREATOR) {
    console.log('  note: ' + BASELINE + ' is not available in this checkout; the comparisons with it are skipped.');
  } else {
    const changed = (now, old) => functionNames(old).filter((name, i, list) => list.indexOf(name) === i && functionSource(now, name) !== functionSource(old, name));
    // Stage D added two lines to the build (the attendance hooks) and one sentence to its question. Stage B
    // added the Session column: the upload time in the TDoc list, the optional fifth column of the
    // registration table where it is created and where rows are appended, and its refresh in the update.
    // The TDoc upload completion (after T-2026.10.5) added its call to the update and, to the abstract sweep,
    // the TDocs the update tells it to leave out; tests/tdoc-upload-completion.test.js compares both with T-2026.10.5.
    check('Code.js: every function of T-2026.10.4 is byte-for-byte what it was, but these seven', changed(CODE, OLD_CODE),
      ['continuousUpdateCore_', 'updateRegisteredDocumentsTable_', 'createSummaryTable_', 'analyzeReportStatus', 'buildSkeletonWithTdocTables', 'downloadAndGroupTdocs_', 'addAbstractsForTables_']);
    check('ReportCreator.js: only the report menu and the build question changed', changed(CREATOR, OLD_CREATOR), ['buildTemplateReportMenu_', 'confirmTemplateBuildFromScratch_']);
    // Outside the functions: constants and comments. Without the ad-hoc sections and the two adoption keys they are unchanged.
    const outsideFunctions = (src) => functionNames(src).filter((name, i, list) => list.indexOf(name) === i).reduce((text, name) => text.replace(functionSource(src, name), ''), src);
    const start = CODE.indexOf('// AD-HOC SESSIONS (stage A) -- SESSION MODEL AND CONFIGURATION\n');
    const end = CODE.indexOf('// ARCH-009 -- MEETING-ID RESOLVER CORE\n');
    // The upload-completion section stands directly before the ad-hoc sections (its banner ends where theirs begins).
    const BAR = '// =========================================================\n';
    const uploadStart = CODE.indexOf(BAR + '// TDOC UPLOAD COMPLETION -- what an existing TDoc gets once it is uploaded\n');
    check('the TDoc upload completion is one section, directly before the ad-hoc sections', [uploadStart !== -1, uploadStart < start, CODE.slice(start - BAR.length, start)], [true, true, BAR]);
    const withoutSections = (CODE.slice(0, uploadStart) + BAR + CODE.slice(end)).replace(/  'DEADLINE_EXTENSIONS',\n[\s\S]*?  'ADHOC_TDOC_SESSIONS'\n\];/, "  'DEADLINE_EXTENSIONS'\n];");
    // The header: the version line and the changelog entries of 2.18.0 and 2.18.1 are new; the rest of it is unchanged.
    const withoutRelease = withoutSections.replace(/ \* 2\.18\.1 \(2026-10-05\)\n[\s\S]*? \* 2\.18\.0 \(2026-10-02\)\n[\s\S]*? \* 2\.17\.4 \(2026-10-01\)\n/, ' * 2.17.4 (2026-10-01)\n').replace(' * Version: 2.18.1 (2026-10-05)\n', ' * Version: 2.17.4 (2026-10-01)\n');
    check('Code.js outside its functions, without the ad-hoc sections, the upload-completion section, the two adoption keys, the version line and the changelog entries of 2.18.0 and 2.18.1, is the T-2026.10.4 file outside its functions',
      [outsideFunctions(withoutRelease) === outsideFunctions(OLD_CODE), withoutRelease === withoutSections], [true, false]);
  }

  const featureSections = CODE.slice(CODE.indexOf('// AD-HOC SESSIONS (stage A)'), CODE.indexOf('// ARCH-009 -- MEETING-ID RESOLVER CORE'));
  const sessionSection = CODE.slice(CODE.indexOf('// AD-HOC SESSIONS (stage A)'), CODE.indexOf('// TEAMS ATTENDANCE (stage C)'));
  // The changelog entry of 2.18.0 in the header names the feature; it is not code.
  const outside = CODE.replace(featureSections, '').replace(/ \* 2\.18\.1 \(2026-10-05\)\n[\s\S]*? \* 2\.17\.4 \(2026-10-01\)\n/, ' * 2.17.4 (2026-10-01)\n');
  check('outside the ad-hoc sections, Code.js names sessions only in the adoption key list, where the status summary adds its block (stage F) and where the build leaves out its "<Chair> opens ..." line (E1)',
    (outside.match(/ADHOC_SESSIONS|adhocSession[A-Za-z_]*(?:\([a-z]*\))?|AdhocSession/g) || []), ['ADHOC_SESSIONS', 'adhocSessionStatusLines_()', 'adhocSessionsEnabled_(context)']);
  check('the session section writes a property in one place only', (sessionSection.match(/\.setProperty\(/g) || []).length, 1);
  check('and touches the document in one place only: the Attendance section, after sessions with attendance were saved (stage D); it looks at the registration table (stage B) and for the opening section (stage E) once each',
    [(sessionSection.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '').match(/getBody\(\)/g) || []).length - (sessionSection.match(/find(Registration|AdhocOpening)(Table|Container)_\(DocumentApp\.getActiveDocument\(\)\.getBody\(\)\)/g) || []).length, /refreshAdhocAttendanceSection_\(DocumentApp\.getActiveDocument\(\)\.getBody\(\)\)/.test(sessionSection),
      /UrlFetchApp|appendTable|appendParagraph|insertTable|insertParagraph/.test(sessionSection.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, ''))], [1, true, false]);

  // The values every build and update starts from, with and without the property.
  const ADHOC = { MEETING_TYPE: 'adhoc', MEETING_ID: '86178', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'September 22, 2026', REPORT_SUFFIX: '6G',
    FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/x/list.xlsx', AGENDA_TDOC: 'S4aP260001' };
  const MAIN = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137_Synthetic', MEETING_NUMBER: '137', MEETING_ID: '60778', REPORT_SUFFIX: 'Audio' };
  const snapshot = (props) => {
    const s = loadCode({ documentProperties: props }).sandbox;
    return JSON.stringify([s.getReportConfig_(), s.getMeetingContext_(), s.getBuildReadiness_(), thrown(() => s.assertMeetingReadyToBuild_())]);
  };
  const VALID = stored(2, [session('s1', '2026-09-22', '15:00', '18:00')]);
  [['malformed JSON', '{"v":1,'], ['an unsupported version', '{"v":99}'], ['an invalid session', stored(2, [session('s1', 'never')])], ['a valid configuration', VALID], ['an empty value', '']]
    .forEach(([what, value]) => {
      check(`ad-hoc report, ${what}: configuration, meeting context and build readiness are what they are without the property`,
        snapshot(Object.assign({}, ADHOC, { [KEY]: value })) === snapshot(ADHOC), true);
      check(`main-meeting report, ${what}: the same`, snapshot(Object.assign({}, MAIN, { [KEY]: value })) === snapshot(MAIN), true);
    });

  // A complete update, with the real continuousUpdateCore_() and collectorUpdate_().
  const update = (props) => {
    const loaded = loadCode({ documentProperties: props });
    const s = loaded.sandbox;
    const events = [];
    const body = makeFakeDocumentBody(s);
    s.DocumentApp.getActiveDocument = () => ({ getId: () => 'DOC1', getBody: () => body });
    s.downloadAndGroupTdocs_ = () => { events.push('TDoc list'); return {}; };
    s.rearrangeRevisionTables_ = () => { events.push('revision placement'); return { moved: 0, dispositions: 0 }; };
    s.checkRSSFeed_ = () => { events.push('e-mail'); };
    s.updateRevisions_ = () => { events.push('revisions'); };
    s.removeRowHeightAndSpacing = () => { events.push('formatting'); };
    const result = s.continuousUpdateCore_();
    return JSON.stringify([result, events, body.getNumChildren(), loaded.docProps._store]);
  };
  const strip = (text) => text.replace(/,?"ADHOC_SESSIONS":"(?:[^"\\]|\\.)*"/, '');
  const reference = update(ADHOC);
  check('an update of an ad-hoc report without sessions completes', JSON.parse(reference).slice(0, 3), [{ success: true, error: null }, ['TDoc list', 'revision placement', 'e-mail', 'revisions'], 0]);
  [['malformed', 'garbage'], ['valid', VALID]].forEach(([what, value]) => {
    const withProperty = update(Object.assign({}, ADHOC, { [KEY]: value }));
    check(`an update with a ${what} ADHOC_SESSIONS: the same result, the same steps, the same document, and the property untouched`,
      [strip(withProperty) === reference, JSON.parse(withProperty)[3][KEY]], [true, value]);
  });
  check('an update creates no ADHOC_SESSIONS property', JSON.parse(reference)[3][KEY], undefined);
  check('clearing the collection caches keeps the sessions (its key prefixes, source)',
    /ADHOC/.test(functionSource(CODE, 'clearAllCaches')), false);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll ad-hoc session configuration checks passed.');
process.exitCode = failures ? 1 : 0;
