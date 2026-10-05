/**
 * Ad-hoc sessions, stage C -- the Teams attendance parser and its
 * normalized model (docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md §9-§11).
 *
 *   1. Bytes to text.
 *   2. Delimited rows: quotes, line breaks, tabs.
 *   3. Durations.
 *   4. Dates and times; month-first or day-first.
 *   5. Names, companies, e-mail addresses.
 *   6. The synthetic export, end to end: the whole model.
 *   7. What Teams states is never replaced by what is calculated.
 *   8. Privacy: the Participant ID, the dial-in number, raw rows.
 *   9. Records are not people: merging, and not merging.
 *  10. Sections, columns and line ends.
 *  11. Fatal problems.
 *  12. Dates of the whole export; the session date.
 *  13. Pure and deterministic.
 *
 * All data is synthetic (tests/fixtures/teams-attendance-synthetic.js).
 *
 * Run: node tests/teams-attendance-parser.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
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

const loaded = loadCode();
const S = loaded.sandbox;
const logged = [];
S.Logger = { log: (m) => logged.push(String(m)) };

const SESSION_DATE = '2026-09-22';
const parse = (text, options) => S.parseTeamsAttendanceReport_(text, options === undefined ? { expectedDate: SESSION_DATE } : options);
const model = (text, options) => { const r = parse(text, options); if (!r.ok) throw new Error('unexpected parser error: ' + JSON.stringify(r.error)); return r.attendance; };
const errorCode = (text, options) => { const r = parse(text, options); return r.ok ? 'ok' : r.error.code; };
const codes = (attendance) => attendance.diagnostics.notes.map((n) => n.code);
const T = (time) => (time ? SESSION_DATE + 'T' + time : null);
const TEXT = FX.buildExport();
const BASE = model(TEXT);

/** The fixture with some participant rows replaced: variant({ 0: { email: '' } }). */
const variant = (changes, options) => FX.buildExport(Object.assign({ participants: FX.PARTICIPANTS.map((p, i) => (changes[i] ? Object.assign({}, p, changes[i]) : p)) }, options || {}));
/** A small export of its own: rows are [name, join, leave, duration, email, upn]. */
const small = (rows, summary) => FX.buildExport({
  summary: Object.assign({ attendedParticipants: String(rows.length) }, summary || {}),
  participants: rows.map((r) => ({ name: r[0], join: r[1] || '', leave: r[2] || '', duration: r[3] || '', email: r[4] || '', upn: r[5] || '', role: 'Presenter' }))
});

// ================================================================ 1. bytes to text

console.log('1. bytes to text');
{
  const d = S.decodeTeamsAttendanceBytes_;
  const sample = 'Name\tÉmile 山田 😀\r\n';
  const bom = (bytes, text, enc) => Buffer.concat([Buffer.from(bytes), Buffer.from(text, enc)]);
  const be = (text) => { const b = Buffer.from(text, 'utf16le'); for (let i = 0; i + 1 < b.length; i += 2) { const x = b[i]; b[i] = b[i + 1]; b[i + 1] = x; } return b; };

  check('UTF-16 LE with a byte-order mark (the real export)', d(FX.utf16le(sample)), { text: sample, encoding: 'utf-16le' });
  check('UTF-16 BE with a byte-order mark', d(Buffer.concat([Buffer.from([0xFE, 0xFF]), be(sample)])), { text: sample, encoding: 'utf-16be' });
  check('UTF-8 with a byte-order mark', d(bom([0xEF, 0xBB, 0xBF], sample, 'utf8')), { text: sample, encoding: 'utf-8' });
  check('UTF-8 without a mark, with 2-, 3- and 4-byte characters', d(Buffer.from(sample, 'utf8')), { text: sample, encoding: 'utf-8' });
  check('UTF-16 LE without a mark is recognized by its zero bytes', d(Buffer.from('1. Summary\r\nMeeting title\tX', 'utf16le')), { text: '1. Summary\r\nMeeting title\tX', encoding: 'utf-16le' });
  check('UTF-16 BE without a mark', d(be('1. Summary\r\nMeeting title\tX')).encoding, 'utf-16be');
  check('signed bytes, as Apps Script supplies them', d(Array.from(FX.utf16le(sample)).map((b) => (b > 127 ? b - 256 : b))).text, sample);
  check('a plain array of unsigned bytes', d(Array.from(FX.utf16le(sample))).text, sample);
  check('an odd trailing byte is dropped', d(Buffer.concat([FX.utf16le('ab'), Buffer.from([0x41])])).text, 'ab');
  check('invalid UTF-8 becomes the replacement character, not an error', d(Buffer.from([0x41, 0xFF, 0xC3, 0x42])).text, 'A\uFFFD\uFFFDB');
  check('nothing', [d([]).text, d(null).text, d(Buffer.from([0xFF, 0xFE])).text], ['', '', '']);

  const whole = d(FX.utf16le(TEXT));
  check('the whole synthetic export survives the round trip', [whole.encoding, whole.text === TEXT, whole.text.length], ['utf-16le', true, TEXT.length]);
  check('and parses to the same model as the text', JSON.stringify(model(whole.text)) === JSON.stringify(BASE), true);
  check('a text that still carries the mark as a character parses the same', JSON.stringify(model('\uFEFF' + TEXT)) === JSON.stringify(BASE), true);
  const big = 'x'.repeat(70000);
  check('a text longer than one conversion block', d(FX.utf16le(big)).text === big, true);
}

// ================================================================ 2. delimited rows

console.log('2. delimited rows');
{
  const split = S.splitDelimitedRows_;
  check('tabs and CR LF', split('a\tb\r\nc\td\r\n', '\t'), [['a', 'b'], ['c', 'd']]);
  check('LF, and no line end at the end', split('a\tb\nc\td', '\t'), [['a', 'b'], ['c', 'd']]);
  check('a lone CR ends a row', split('a\rb', '\t'), [['a'], ['b']]);
  check('a quoted line break stays inside the cell', split('t\t"one\r\ntwo"\r\nx\ty', '\t'), [['t', 'one\r\ntwo'], ['x', 'y']]);
  check('a quoted tab stays inside the cell', split('"a\tb"\tc', '\t'), [['a\tb', 'c']]);
  check('"" inside quotes is one quote', split('"say ""hi"""\tx', '\t'), [['say "hi"', 'x']]);
  check('a quoted cell keeps its leading space', split('" 123 (x)"\ty', '\t'), [[' 123 (x)', 'y']]);
  check('an empty quoted cell', split('""\tb', '\t'), [['', 'b']]);
  check('a quote that is not at the start of a cell is an ordinary character', split('5\' 11" tall\tx\ta "b" c', '\t'), [['5\' 11" tall', 'x', 'a "b" c']]);
  check('empty cells and a blank line', split('a\t\tb\n\n\tc', '\t'), [['a', '', 'b'], [''], ['', 'c']]);
  check('commas, with a quoted comma', split('a,"b,c",d', ','), [['a', 'b,c', 'd']]);
  check('a tab is data in a comma-separated text', split('a\tb,c', ','), [['a\tb', 'c']]);
  check('a quoted cell that is never closed: null', [split('a\t"b\nc', '\t'), split('"', '\t')], [null, null]);
  check('text after a closing quote: null, the cell is not read leniently', [split('"a"b\tc', '\t'), split('"a" \tc', '\t'), split('"a""', '\t')], [null, null, null]);
  check('a closing quote followed by the delimiter, a line end or the end of the text', [split('"a"\tb', '\t'), split('"a"\r\nb', '\t'), split('"a"', '\t')], [[['a', 'b']], [['a'], ['b']], [['a']]]);
  // Why it must not be lenient: an unclosed quote would otherwise run on to
  // the next quoted cell and make one "name" out of several people's rows.
  check('an unclosed quote does not swallow the rows up to the next quoted cell', split('"Kai\tk@example.com\r\nNoor\tn@example.org\r\n"Sky"\ts@example.net', '\t'), null);
  check('empty text: no rows', split('', '\t'), []);

  const seen = [];
  const complete = S.scanDelimitedRows_('a\nb\nc\n"never closed', '\t', (row) => { seen.push(row[0]); return row[0] !== 'b'; });
  check('reading can be stopped; what follows is never looked at', [seen, complete], [['a', 'b'], true]);
  check('splitCsvRows_() of the agenda import is unchanged by this stage', S.splitCsvRows_('"1.1","A, b"\r\n"1.2","C"'), [['1.1', 'A, b'], ['1.2', 'C']]);
}

// ================================================================ 3. durations

console.log('3. durations');
{
  const p = S.parseTeamsDuration_;
  check('the forms Teams writes', ['3h 23m 20s', '2h 31m 7s', '1h 7s', '43m 12s', '2h 5m', '8s', '45m', '3h', '0s'].map(p), [12200, 9067, 3607, 2592, 7500, 8, 2700, 10800, 0]);
  check('spacing and case', ['3h23m20s', ' 3H 23M 20S ', '3 h 23 m 20 s'].map(p), [12200, 12200, 12200]);
  check('long durations', p('100h 1m 1s'), 360061);
  check('not a duration', ['', null, undefined, '3:23:20', '12', 'h', '3h 12', 'about 3h', '3s 2h', '3h 3h'].map(p), [null, null, null, null, null, null, null, null, null, null]);
  check('shown as H:MM:SS', [12200, 9067, 3607, 59, 0, 360061].map(S.formatTeamsDuration_), ['3:23:20', '2:31:07', '1:00:07', '0:00:59', '0:00:00', '100:01:01']);
  check('nothing to show for a missing value', [null, undefined, -1, NaN, '12'].map(S.formatTeamsDuration_), ['', '', '', '', '']);
}

// ================================================================ 4. dates and times

console.log('4. dates and times');
{
  const parts = S.readTeamsDateTimeParts_;
  const read = (value, order) => S.teamsDateTimeFromParts_(parts(value), order);
  check('the real form: 12-hour clock, two-digit year', read('9/22/26, 2:48:05 PM', 'month-first'), '2026-09-22T14:48:05');
  check('morning, noon, midnight', ['9/22/26, 9:05:00 AM', '9/22/26, 12:00:00 PM', '9/22/26, 12:00:00 AM', '9/22/26, 12:30:00 am'].map((v) => read(v, 'month-first')),
    ['2026-09-22T09:05:00', '2026-09-22T12:00:00', '2026-09-22T00:00:00', '2026-09-22T00:30:00']);
  check('the same text read day-first', read('9/10/26, 2:48:05 PM', 'day-first'), '2026-10-09T14:48:05');
  check('four-digit year, 24-hour clock, no seconds, no comma', ['22/9/2026, 14:48:05', '22/9/2026 14:54'].map((v) => read(v, 'day-first')), ['2026-09-22T14:48:05', '2026-09-22T14:54:00']);
  check('a narrow no-break space before PM', read('9/22/26, 2:48:05\u202fPM', 'month-first'), '2026-09-22T14:48:05');
  check('dots are day-first whatever the order says', [read('22.09.26, 14:48:05', 'month-first'), read('22.09.2026 14:48:05', null)], ['2026-09-22T14:48:05', '2026-09-22T14:48:05']);
  check('year-first is unambiguous', [read('2026-09-22 14:48:05', null), read('2026-09-22T14:48:05', 'day-first')], ['2026-09-22T14:48:05', '2026-09-22T14:48:05']);
  check('not a date under the given order', [read('9/22/26, 2:48:05 PM', 'day-first'), read('2/30/26, 1:00:00 PM', 'month-first'), read('31.02.26, 10:00:00', null)], [null, null, null]);
  check('not a date and time at all', ['', null, 'yesterday', '9/22/26', '2:48:05 PM', '9/22/26, 25:00:00', '9/22/26, 13:00:00 PM', '9/22/26, 0:30:00 AM', '9/22/26, 2:61:00 PM', '9-22-26, 2:48:05 PM'].map(parts),
    [null, null, null, null, null, null, null, null, null, null]);
  check('no time zone is added or assumed', /Z|[+-]\d\d:?\d\d|CEST|CET|UTC|GMT/.test(read('9/22/26, 2:48:05 PM', 'month-first').slice(10)), false);

  check('seconds between two times', [S.teamsSecondsBetween_('2026-09-22T14:48:05', '2026-09-22T18:11:26'), S.teamsSecondsBetween_('2026-09-22T23:00:00', '2026-09-23T01:00:00'), S.teamsSecondsBetween_(null, '2026-09-22T18:11:26')], [12201, 7200, null]);

  const resolve = (values, expected) => { const r = S.resolveTeamsDateOrder_(values, expected || ''); return r.error ? r.error.code : r.order; };
  check('decided by the values alone when only one reading is a date', [resolve(['9/22/26, 2:48:05 PM']), resolve(['22/9/26, 14:48:05'])], ['month-first', 'day-first']);
  check('an ambiguous start, decided by another value', [resolve(['9/10/26, 2:00:00 PM', '9/13/26, 1:00:00 AM']), resolve(['9/10/26, 14:00:00', '13/10/26, 01:00:00'])], ['month-first', 'day-first']);
  check('ambiguous throughout, no session date: an error, not a guess', resolve(['9/10/26, 2:00:00 PM', '9/10/26, 5:00:00 PM']), 'AMBIGUOUS_DATE');
  check('ambiguous throughout, decided by the session date', [resolve(['9/10/26, 2:00:00 PM'], '2026-09-10'), resolve(['9/10/26, 2:00:00 PM'], '2026-10-09')], ['month-first', 'day-first']);
  check('neither reading is the session date', resolve(['9/10/26, 2:00:00 PM'], '2026-09-11'), 'SESSION_DATE_MISMATCH');
  check('the session date wins over what other values suggest', resolve(['9/10/26, 2:00:00 PM', '9/13/26, 1:00:00 AM'], '2026-10-09'), 'day-first');
  check('day equal to month: both readings are the same date', [resolve(['5/5/26, 2:00:00 PM']), resolve(['5/5/26, 2:00:00 PM'], '2026-05-05')], ['month-first', 'month-first']);
  check('day equal to month at the start, another value differs and decides', resolve(['5/5/26, 11:00:00 PM', '5/13/26, 1:00:00 AM'], '2026-05-05'), 'month-first');
  check('day equal to month at the start, another value differs and does not decide', resolve(['5/5/26, 11:00:00 PM', '5/6/26, 1:00:00 AM'], '2026-05-05'), 'AMBIGUOUS_DATE');
  check('no slash date at all: no order is needed', [resolve(['22.09.26, 14:48:05']), resolve(['2026-09-22 14:48:05'], '2026-09-22')], [null, null]);
  check('an unambiguous start that is not the session date', [resolve(['22.09.26, 14:48:05'], '2026-09-23'), resolve(['2026-09-22 14:48:05'], '2026-09-23')], ['SESSION_DATE_MISMATCH', 'SESSION_DATE_MISMATCH']);
  check('the mismatch message names the dates, which are not private', S.resolveTeamsDateOrder_(['9/10/26, 2:00:00 PM'], '2026-09-11').error.message,
    'This attendance export is for 2026-09-10 or 2026-10-09, not for the session on 2026-09-11.');
}

// ================================================================ 5. names, companies, e-mail

console.log('5. names, companies, e-mail addresses');
{
  const n = S.normalizeTeamsDisplayName_;
  const plain = (name) => ({ name: name, company: '', external: false, unverified: false, dialIn: false });
  check('an ordinary name is unchanged', n('Alex Organizer'), plain('Alex Organizer'));
  check('surrounding and repeated white space', n('  Alex   Organizer \u00a0'), plain('Alex Organizer'));
  check('(External) is removed and recorded', n('Doe, Jane (External)'), Object.assign(plain('Doe, Jane'), { external: true }));
  check('(Unverified) is removed and recorded', n('Casey Guest (Unverified)'), Object.assign(plain('Casey Guest'), { unverified: true }));
  check('both, any case, any spacing', n('Casey Guest  (external)(UNVERIFIED) '), Object.assign(plain('Casey Guest'), { external: true, unverified: true }));
  check('a marker that is not at the end is part of the name', n('The (External) Group'), plain('The (External) Group'));
  check('other bracketed text stays: an organisation', n('Pat Kim (ExampleCorp) (External)'), Object.assign(plain('Pat Kim (ExampleCorp)'), { external: true }));
  check('other bracketed text stays: a second spelling', n('Taro Yamada (山田 太郎) (External)'), Object.assign(plain('Taro Yamada (山田 太郎)'), { external: true }));
  check('full-width brackets, another script', n('张 三（示例）'), plain('张 三（示例）'));
  check('comma, apostrophe, hyphen, diacritics', [n('O\'Neil-Smith, Zoë').name, n('Zoë O\'Connor-Müller').name], ['O\'Neil-Smith, Zoë', 'Zoë O\'Connor-Müller']);
  check('composed and decomposed forms give the same name', n('Zoe\u0308 Mu\u0308ller').name === n('Zoë Müller').name, true);
  check('quotes inside a name stay', n('Sky "Ace" Tab\tName').name, 'Sky "Ace" Tab Name');

  check('[Company] Name: the company, and the name without it', n('[ExampleCorp] Sam Rivera (Unverified)'), { name: 'Sam Rivera', company: 'ExampleCorp', external: false, unverified: true, dialIn: false });
  check('a company of several words; no space after the bracket', [n('[Example Corp. GmbH] Sam Rivera').company, n('[ExampleCorp]Sam Rivera')], ['Example Corp. GmbH', Object.assign(plain('Sam Rivera'), { company: 'ExampleCorp' })]);
  check('no company from a trailing round bracket', n('Sam Rivera (ExampleCorp)').company, '');
  check('no company from a trailing square bracket', n('Sam Rivera [ExampleCorp]'), plain('Sam Rivera [ExampleCorp]'));
  check('no company from a dash or a slash', [n('Robin Lee - ExampleCorp'), n('Morgan Diaz ExampleCorp/Research')], [plain('Robin Lee - ExampleCorp'), plain('Morgan Diaz ExampleCorp/Research')]);
  check('a bracket with nothing after it, an empty bracket, a bracket in the middle', [n('[ExampleCorp]'), n('[] Sam Rivera'), n('Sam [ExampleCorp] Rivera')],
    [plain('[ExampleCorp]'), plain('[] Sam Rivera'), plain('Sam [ExampleCorp] Rivera')]);

  const dialIn = { name: 'Dial-in participant', company: '', external: false, unverified: false, dialIn: true };
  check('a telephone number as a name', [n(' 15550100123 (Unverified)'), n('+1 555 010 0123')], [Object.assign({}, dialIn, { unverified: true }), dialIn]);
  check('written with brackets, dashes, dots, masking', ['(555) 010-0123', '+49.555.0100.123', '+1 555-***-**23 12'].map((v) => n(v).name), ['Dial-in participant', 'Dial-in participant', 'Dial-in participant']);
  check('a name that contains digits is a name', ['Agent 47', 'R2 Unit', 'Sam Rivera 2', 'Room 101 (ExampleCorp)', '3M Delegate'].map((v) => n(v).name), ['Agent 47', 'R2 Unit', 'Sam Rivera 2', 'Room 101 (ExampleCorp)', '3M Delegate']);
  check('a few digits alone are not a telephone number', [n('2026').dialIn, n('42').dialIn], [false, false]);
  check('nothing', [n('').name, n(null).name, n('   ').name, n('(Unverified)')], ['', '', '', Object.assign(plain(''), { unverified: true })]);

  const e = S.normalizeTeamsEmail_;
  check('valid addresses are kept as written, trimmed', [' alex.organizer@example.com ', 'LEE.WONG@example.com', 'a+tag@sub.example.org', 'o\'neil@example.co.uk'].map(e),
    ['alex.organizer@example.com', 'LEE.WONG@example.com', 'a+tag@sub.example.org', 'o\'neil@example.co.uk']);
  check('not one usable address', ['', null, 'kai.invalid(at)example', 'kai@example', 'kai@@example.com', 'kai @example.com', 'a@example.com, b@example.com', 'a@example.com;b@example.com',
    'Kai <kai@example.com>', 'mailto:kai@example.com', '@example.com', 'kai@', 'kai..x@example.com', 'https://example.com/u/kai'].map(e), ['', '', '', '', '', '', '', '', '', '', '', '', '', '']);

  const key = S.teamsAttendeeKey_;
  check('attendee key: the e-mail, in lower case', key({ name: 'Lee Wong', company: '', email: 'LEE.WONG@example.com' }), 'email:lee.wong@example.com');
  check('attendee key without e-mail: name and company, in lower case', [key({ name: 'Sam Rivera', company: 'ExampleCorp', email: '' }), key({ name: 'Casey Guest', company: '', email: '' })],
    ['name:sam rivera|examplecorp', 'name:casey guest|']);
}

// ================================================================ 6. the synthetic export

console.log('6. the synthetic export, end to end');
{
  const P = (name, company, email, join, leave, seconds, records) => ({ name: name, company: company, email: email, firstJoin: T(join), lastLeave: T(leave), durationSeconds: seconds, records: records || 1 });
  const result = parse(TEXT);
  check('accepted', [result.ok, result.error], [true, null]);
  check('the model has exactly these parts', [Object.keys(result), Object.keys(BASE)], [['ok', 'error', 'attendance'], ['schemaVersion', 'source', 'summary', 'participants', 'diagnostics']]);
  check('schema version and source', [BASE.schemaVersion, BASE.source], [1, { type: 'teams-attendance', delimiter: 'tab', dateOrder: 'month-first' }]);

  check('summary: what Teams states', BASE.summary, {
    meetingTitle: 'Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE',
    attendanceRecords: 25,
    start: '2026-09-22T14:48:05',
    end: '2026-09-22T18:11:26',
    durationSeconds: 12200,
    averageAttendanceSeconds: 9067,
    raw: { attendedParticipants: '25', startTime: '9/22/26, 2:48:05 PM', endTime: '9/22/26, 6:11:26 PM', meetingDuration: '3h 23m 20s', averageAttendanceTime: '2h 31m 7s' }
  });
  check('the raw summary values are the fixture\'s, character for character', BASE.summary.raw, FX.SUMMARY);

  check('participants: one entry per attendee, in order of first appearance', BASE.participants, [
    P('Alex Organizer', '', 'alex.organizer@example.com', '14:48:09', '18:11:26', 12197),
    P('Doe, Jane', '', 'jane.doe@example.org', '15:00:00', '18:00:00', 10800),
    P('Sam Rivera', 'ExampleCorp', '', '15:01:00', '17:01:00', 7200),
    P('Taro Yamada (山田 太郎)', '', 'taro.yamada@example.net', '15:00:10', '17:30:15', 9005),
    P('张 三（示例）', '', 'zhang.san@example.com', '15:05:00', '16:05:07', 3607),
    P('Zoë O\'Connor-Müller', '', 'zoe.oconnor-mueller@example.org', '15:15:00', '16:00:00', 2700),
    P('Pat Kim (ExampleCorp)', '', 'pat.kim@example.com', '15:00:00', '18:06:40', 11200),
    P('Robin Lee - ExampleCorp', '', '', '15:02:00', '18:02:00', 10800),
    P('Morgan Diaz ExampleCorp/Research', '', 'morgan.diaz@example.net', '15:00:00', '17:45:00', 9900),
    P('Dial-in participant', '', '', '15:07:20', '17:41:02', 9222),
    P('Casey Guest', '', '', '15:01:30', '18:11:20', 2592 + 8749, 2),
    P('Dana Twin', '', '', '15:00:00', '16:00:00', 3600),
    P('Dana Twin', '', '', '15:30:00', '17:00:00', 5400),
    P('Jamie Sample', '', 'jamie.sample@example.com', '15:00:00', '17:00:00', 7200),
    P('Jamie Sample', '', '', '17:10:00', '17:20:00', 600),
    P('Lee Wong', '', 'lee.wong@example.com', '15:00:00', '18:00:00', 3600 + 5400, 2),
    P('Kai Invalid', '', '', '15:00:00', '18:00:00', 10800),
    P('Noor Late', '', 'noor.late@example.org', null, '18:00:00', 1800),
    P('Omar Open', '', 'omar.open@example.com', '15:10:00', null, null),
    P('Sky "Ace" Tab Name', '', 'sky.tab@example.net', '15:00:00', '18:00:00', 10800),
    P('Rita Rejoin', '', 'rita.rejoin@example.com', '15:00:00', '18:00:00', 5400),
    P('Quinn, Avery', '', 'avery.quinn@example.org', null, null, null),
    P('Riley Short (ExampleCorp)', '', 'riley.short@example.com', null, null, null)
  ]);
  check('a participant has exactly these fields', Object.keys(BASE.participants[0]), ['name', 'company', 'email', 'firstJoin', 'lastLeave', 'durationSeconds', 'records']);

  // The same numbers, obtained from the fixture's own rows.
  const seconds = (d) => (d.match(/(\d+)([hms])/g) || []).reduce((sum, x) => sum + parseInt(x, 10) * { h: 3600, m: 60, s: 1 }[x.slice(-1)], 0);
  const durations = FX.PARTICIPANTS.filter((p) => p.duration).map((p) => seconds(p.duration));
  const total = durations.reduce((a, b) => a + b, 0);
  check('counts', BASE.diagnostics.counts, {
    participantRows: FX.PARTICIPANTS.length, rowsWithJoinTime: FX.PARTICIPANTS.filter((p) => p.join).length, attendees: 23, mergedRecords: 2,
    shortRows: FX.PARTICIPANTS.filter((p) => p.cells).length, external: FX.PARTICIPANTS.filter((p) => /\(External\)/.test(p.name)).length,
    unverified: FX.PARTICIPANTS.filter((p) => /\(Unverified\)/.test(p.name)).length, dialIn: 1, withoutEmail: 8
  });
  check('the counts, as numbers', BASE.diagnostics.counts, { participantRows: 25, rowsWithJoinTime: 22, attendees: 23, mergedRecords: 2, shortRows: 2, external: 5, unverified: 8, dialIn: 1, withoutEmail: 8 });
  check('calculated values', BASE.diagnostics.calculated, { totalDurationSeconds: total, rowsWithDuration: durations.length, meanDurationSeconds: Math.round(total / durations.length), spanSeconds: 12201 });
  check('every record is in exactly one attendee', BASE.participants.reduce((sum, p) => sum + p.records, 0), 25);
  check('the sum of the attendees\' durations is the sum of the rows\'', BASE.participants.reduce((sum, p) => sum + (p.durationSeconds || 0), 0), total);

  check('notes: code, level and rows', BASE.diagnostics.notes.map((n) => [n.level, n.code, n.rows || null]), [
    ['info', 'RECORDS_MERGED', [11, 22]],
    ['info', 'RECORDS_MERGED', [16, 17]],
    ['warning', 'DUPLICATE_CANDIDATE', [12, 13]],
    ['warning', 'DUPLICATE_CANDIDATE', [14, 15]],
    ['warning', 'INVALID_EMAIL', [18]],
    ['info', 'INCOMPLETE_TIMES', [19, 20, 24, 25]],
    ['info', 'SHORT_ROWS', null],
    ['info', 'AVERAGE_DIFFERS', null],
    ['info', 'DURATION_DIFFERS', null]
  ]);
  check('a note has a level, a code, a message and at most its rows', BASE.diagnostics.notes.filter((n) => Object.keys(n).some((k) => ['level', 'code', 'message', 'rows'].indexOf(k) === -1)), []);
}

// ================================================================ 7. source authority

console.log('7. what Teams states is never replaced by what is calculated');
{
  const c = BASE.diagnostics.calculated;
  check('the fixture is built so that they differ: records, average, duration',
    [BASE.summary.attendanceRecords !== BASE.diagnostics.counts.attendees, BASE.summary.averageAttendanceSeconds !== c.meanDurationSeconds, BASE.summary.durationSeconds !== c.spanSeconds], [true, true, true]);
  check('the summary still says what Teams wrote', [BASE.summary.attendanceRecords, BASE.summary.averageAttendanceSeconds, BASE.summary.durationSeconds], [25, 9067, 12200]);
  check('the differences are notes, and say which value is used',
    BASE.diagnostics.notes.filter((n) => /DIFFERS/.test(n.code)).map((n) => n.message), [
      'Teams states an average attendance of 2:31:07; the mean of the listed durations is 1:55:35. The value Teams states is used.',
      'Teams states a meeting duration of 3:23:20; end minus start is 3:23:21. The value Teams states is used.']);

  // Change the rows: the summary must not move.
  const fewer = model(FX.buildExport({ participants: FX.PARTICIPANTS.slice(0, 5) }));
  check('fewer participant rows: the summary is unchanged', JSON.stringify(fewer.summary) === JSON.stringify(BASE.summary), true);
  check('and the difference in the number of records is a warning', [fewer.summary.attendanceRecords, fewer.diagnostics.counts.participantRows, fewer.diagnostics.notes.filter((n) => n.code === 'RECORD_COUNT_DIFFERS')],
    [25, 5, [{ level: 'warning', code: 'RECORD_COUNT_DIFFERS', message: 'Teams states 25 attended participants; the Participants section has 5 rows.' }]]);
  const longer = model(variant({ 1: { duration: '9h' }, 2: { duration: '9h' } }));
  check('other row durations: the stated average and duration are unchanged, the calculated ones move',
    [longer.summary.averageAttendanceSeconds, longer.summary.durationSeconds, longer.diagnostics.calculated.totalDurationSeconds - c.totalDurationSeconds], [9067, 12200, (9 - 3) * 3600 + (9 - 2) * 3600]);

  // Change the summary: the rows must not move, and the summary follows Teams.
  const other = model(FX.buildExport({ summary: { attendedParticipants: '66', meetingDuration: '1h', averageAttendanceTime: '5m 1s' } }));
  check('other summary values are taken as stated, however implausible', [other.summary.attendanceRecords, other.summary.durationSeconds, other.summary.averageAttendanceSeconds], [66, 3600, 301]);
  check('and the participants and calculated values are unchanged', [JSON.stringify(other.participants) === JSON.stringify(BASE.participants), JSON.stringify(other.diagnostics.calculated) === JSON.stringify(c)], [true, true]);

  // When they agree there is nothing to note.
  const agreeing = model(small([['Alex Organizer', '3:00:00 PM', '4:00:00 PM', '1h', 'alex.organizer@example.com']],
    { startTime: FX.at('3:00:00 PM'), endTime: FX.at('4:00:00 PM'), meetingDuration: '1h', averageAttendanceTime: '1h' }));
  check('a summary that agrees with the rows: no note', codes(agreeing), []);

  // A value Teams did not supply stays missing.
  const missing = model(TEXT.replace(/^Average attendance time\t.*\r\n/m, '').replace(/^Meeting duration\t.*$/m, 'Meeting duration\tabout three hours'));
  check('a missing or unreadable summary value is null, not the calculated value',
    [missing.summary.averageAttendanceSeconds, missing.summary.durationSeconds, missing.summary.raw.averageAttendanceTime, missing.summary.raw.meetingDuration], [null, null, '', 'about three hours']);
  check('and is a warning', missing.diagnostics.notes.filter((n) => n.code === 'SUMMARY_VALUE_MISSING').map((n) => n.message),
    ['The Summary has no readable "Meeting duration".', 'The Summary has no readable "Average attendance time".']);
  check('no "DIFFERS" note is made up for a value that is missing', codes(missing).filter((code) => /DIFFERS/.test(code)), []);
  const noCount = model(TEXT.replace(/^Attended participants\t.*$/m, 'Attended participants\tmany'));
  check('an unreadable record count is null, not the number of rows', [noCount.summary.attendanceRecords, noCount.summary.raw.attendedParticipants], [null, 'many']);
}

// ================================================================ 8. privacy

console.log('8. privacy');
{
  const everything = JSON.stringify(parse(TEXT));
  check('the fixture does contain the Participant IDs and the dial-in number', FX.PRIVATE_STRINGS.map((s) => TEXT.indexOf(s) !== -1), [true, true, true]);
  check('none of them is anywhere in what the parser returns', FX.PRIVATE_STRINGS.filter((s) => everything.indexOf(s) !== -1), []);
  check('no run of seven or more digits is returned at all', /\d{7,}/.test(everything), false);
  check('no Teams marker is returned', /\((External|Unverified)\)/i.test(everything), false);
  check('the parser logged nothing', logged, []);

  // The Participant ID column has no influence on the result.
  const withoutUpn = FX.buildExport({ participants: FX.PARTICIPANTS.map((p) => Object.assign({}, p, { upn: '' })) });
  const otherUpn = FX.buildExport({ participants: FX.PARTICIPANTS.map((p, i) => Object.assign({}, p, { upn: 'other-' + i + '@tenant.example.net' })) });
  check('emptying the Participant ID column changes nothing', JSON.stringify(parse(withoutUpn)) === everything, true);
  check('replacing every Participant ID changes nothing', JSON.stringify(parse(otherUpn)) === everything, true);
  const sameUpn = model(FX.buildExport({ participants: FX.PARTICIPANTS.map((p) => Object.assign({}, p, { upn: 'shared@tenant.example.net' })) }));
  check('one Participant ID on every row merges nobody', sameUpn.participants.length, BASE.participants.length);

  // Never an e-mail fallback.
  const upnOnly = model(small([['Kai Fallback', '3:00:00 PM', '4:00:00 PM', '1h', '', 'kai.fallback@tenant.example.com'],
    ['Kai Invalid', '3:00:00 PM', '4:00:00 PM', '1h', 'not an address', 'kai.invalid@tenant.example.com']]));
  check('no e-mail, or an invalid one, with a Participant ID that looks like an address: the e-mail stays empty', upnOnly.participants.map((p) => p.email), ['', '']);
  check('the invalid value itself is not returned either', JSON.stringify(upnOnly).indexOf('not an address'), -1);

  // The dial-in number.
  const dial = BASE.participants.filter((p) => p.name === 'Dial-in participant');
  check('the dial-in record is kept, as "Dial-in participant", with its times', dial, [{ name: 'Dial-in participant', company: '', email: '', firstJoin: T('15:07:20'), lastLeave: T('17:41:02'), durationSeconds: 9222, records: 1 }]);
  const twoDialIns = model(small([['+1 555 010 0123', '3:00:00 PM', '3:30:00 PM', '30m'], ['+1 555 010 0199 (Unverified)', '4:00:00 PM', '4:30:00 PM', '30m'], ['+1 555 010 0123', '5:00:00 PM', '5:30:00 PM', '30m']]));
  check('dial-in records are never merged with each other (their numbers are gone)', [twoDialIns.participants.map((p) => [p.name, p.records]), codes(twoDialIns).filter((c) => /MERGED|DUPLICATE/.test(c))],
    [[['Dial-in participant', 1], ['Dial-in participant', 1], ['Dial-in participant', 1]], []]);
  check('and no number is returned', /555|0123|0199/.test(JSON.stringify(twoDialIns)), false);

  // Nothing that is never used is kept.
  check('no role, no engagement value, no unknown column', /Organizer|Presenter|"Attendee"|Engagement|Synthetic Future|"role"|"upn"/i.test(everything.replace('Alex Organizer', '').replace('alex.organizer', '')), false);
  check('notes carry row numbers only: no name, no address, nothing from a row',
    BASE.diagnostics.notes.filter((n) => FX.PARTICIPANTS.some((p) => { const who = S.normalizeTeamsDisplayName_(p.raw ? p.name.replace(/^"|"$/g, '') : p.name).name; return (who && n.message.indexOf(who) !== -1) || (p.email && n.message.indexOf(p.email) !== -1); })), []);
  check('no e-mail address in any note', BASE.diagnostics.notes.filter((n) => /@/.test(n.message)), []);

  // Sections 3 and 4 are never looked at.
  const marked = TEXT.replace('3. In-Meeting Activities\r\n', '3. In-Meeting Activities\r\nSECRET-ACTIVITY-ROW\tsecret.person@example.com\r\n');
  check('nothing from the sections that are not read is returned', [marked !== TEXT, JSON.stringify(parse(marked)) === everything], [true, true]);

  // Fatal errors say nothing about participants either.
  const fatal = parse(TEXT.replace('2. Participants', '2. Attendees'));
  check('a fatal error carries a code and a message only', [Object.keys(fatal), Object.keys(fatal.error), fatal.attendance], [['ok', 'error', 'attendance'], ['code', 'message'], null]);

  const section = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const parser = section.slice(section.indexOf('// TEAMS ATTENDANCE (stage C)'), section.indexOf('// AD-HOC ATTENDANCE (stage D)')).replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '');
  check('the parser never reads the Participant ID column (source)', /upn|participant id/i.test(parser), false);
  check('the parser has no log statement (source)', /Logger|console\./.test(parser), false);
}

// ================================================================ 9. records are not people

console.log('9. records are not people');
{
  const names = (attendance) => attendance.participants.map((p) => [p.name, p.records]);
  check('three numbers that are kept apart: Teams\' records, the rows with a join time, the attendees',
    [BASE.summary.attendanceRecords, BASE.diagnostics.counts.rowsWithJoinTime, BASE.diagnostics.counts.attendees], [25, 22, 23]);

  // The same explicit e-mail.
  const lee = BASE.participants.filter((p) => /wong/i.test(p.email))[0];
  check('same e-mail, different case and a different display name: one attendee, the first name and address kept', lee,
    { name: 'Lee Wong', company: '', email: 'lee.wong@example.com', firstJoin: T('15:00:00'), lastLeave: T('18:00:00'), durationSeconds: 9000, records: 2 });
  const overlappingMail = model(small([['Lee Wong', '3:00:00 PM', '5:00:00 PM', '2h', 'lee.wong@example.com'], ['Lee Wong (phone)', '3:30:00 PM', '6:00:00 PM', '2h 30m', 'lee.wong@example.com']]));
  check('same e-mail on two devices at once: still one attendee', [names(overlappingMail), overlappingMail.participants[0].firstJoin, overlappingMail.participants[0].lastLeave, overlappingMail.participants[0].durationSeconds],
    [[['Lee Wong', 2]], T('15:00:00'), T('18:00:00'), 16200]);
  const mailNoTimes = model(small([['Lee Wong', '3:00:00 PM', '5:00:00 PM', '2h', 'lee.wong@example.com'], ['Lee Wong', '', '', '', 'lee.wong@example.com']]));
  check('same e-mail, one record without times: one attendee, the known times kept', [names(mailNoTimes), mailNoTimes.participants[0].durationSeconds], [[['Lee Wong', 2]], 7200]);

  // The same name, no e-mail.
  check('a guest who left and joined again: one attendee', BASE.participants.filter((p) => p.name === 'Casey Guest').map((p) => [p.records, p.firstJoin, p.lastLeave, p.durationSeconds]),
    [[2, T('15:01:30'), T('18:11:20'), 11341]]);
  check('two guests of one name present at the same time: two attendees, and a note', [BASE.participants.filter((p) => p.name === 'Dana Twin').length,
    BASE.diagnostics.notes.filter((n) => n.code === 'DUPLICATE_CANDIDATE' && n.rows[0] === 12).length], [2, 1]);
  const three = model(small([['Casey Guest (Unverified)', '3:00:00 PM', '3:45:00 PM', '45m'], ['Casey Guest (Unverified)', '3:50:00 PM', '6:00:00 PM', '2h 10m'], ['Casey Guest  (Unverified)', '5:55:00 PM', '6:00:00 PM', '5m']]));
  check('two visits in turn and a third that overlaps the second: two attendees', [names(three), three.diagnostics.notes.filter((n) => /MERGED|DUPLICATE/.test(n.code)).map((n) => [n.code, n.rows])],
    [[['Casey Guest', 2], ['Casey Guest', 1]], [['RECORDS_MERGED', [1, 2]], ['DUPLICATE_CANDIDATE', [1, 2, 3]]]]);
  const touching = model(small([['Casey Guest', '3:00:00 PM', '4:00:00 PM', '1h'], ['Casey Guest', '4:00:00 PM', '5:00:00 PM', '1h']]));
  check('leaving and joining in the same second is not being present twice', names(touching), [['Casey Guest', 2]]);
  const inGap = model(small([['Casey Guest', '3:00:00 PM', '3:30:00 PM', '30m'], ['Casey Guest', '5:00:00 PM', '5:30:00 PM', '30m'], ['Casey Guest', '4:00:00 PM', '4:30:00 PM', '30m']]));
  check('a third visit between the first two', [names(inGap), inGap.participants[0].firstJoin, inGap.participants[0].lastLeave, inGap.participants[0].durationSeconds], [[['Casey Guest', 3]], T('15:00:00'), T('17:30:00'), 5400]);
  const noTimes = model(small([['Casey Guest', '3:00:00 PM', '4:00:00 PM', '1h'], ['Casey Guest', '', '', '']]));
  check('same name, one record without times: not enough to merge', [names(noTimes), codes(noTimes).filter((c) => c === 'DUPLICATE_CANDIDATE').length], [[['Casey Guest', 1], ['Casey Guest', 1]], 1]);
  const halfTimes = model(small([['Casey Guest', '3:00:00 PM', '4:00:00 PM', '1h'], ['Casey Guest', '5:00:00 PM', '', '']]));
  check('same name, one record without a leave time: not enough to merge', names(halfTimes), [['Casey Guest', 1], ['Casey Guest', 1]]);

  // Evidence that is not enough.
  check('same name, one with an e-mail and one without: two attendees', BASE.participants.filter((p) => p.name === 'Jamie Sample').map((p) => [p.email, p.records]), [['jamie.sample@example.com', 1], ['', 1]]);
  const twoMails = model(small([['Jamie Sample', '3:00:00 PM', '4:00:00 PM', '1h', 'jamie.sample@example.com'], ['Jamie Sample', '5:00:00 PM', '6:00:00 PM', '1h', 'j.sample@example.org']]));
  check('same name, two different e-mails: two attendees, and a note', [names(twoMails), codes(twoMails).filter((c) => c === 'DUPLICATE_CANDIDATE').length], [[['Jamie Sample', 1], ['Jamie Sample', 1]], 1]);
  const similar = model(small([['Casey Guest', '3:00:00 PM', '4:00:00 PM', '1h'], ['Casey Guest Jr', '5:00:00 PM', '6:00:00 PM', '1h'], ['Guest, Casey', '4:10:00 PM', '4:20:00 PM', '10m'], ['Casey', '4:30:00 PM', '4:40:00 PM', '10m'], ['Casey Gest', '4:45:00 PM', '4:50:00 PM', '5m']]));
  check('similar names are different people', [similar.participants.length, codes(similar).filter((c) => /MERGED|DUPLICATE/.test(c))], [5, []]);
  const companies = model(small([['[ExampleCorp] Sam Rivera', '3:00:00 PM', '4:00:00 PM', '1h'], ['[OtherCorp] Sam Rivera', '5:00:00 PM', '6:00:00 PM', '1h'], ['Sam Rivera', '4:10:00 PM', '4:20:00 PM', '10m']]));
  check('same name under different companies: different people', [companies.participants.map((p) => [p.name, p.company, p.records]), codes(companies).filter((c) => c === 'DUPLICATE_CANDIDATE').length],
    [[['Sam Rivera', 'ExampleCorp', 1], ['Sam Rivera', 'OtherCorp', 1], ['Sam Rivera', '', 1]], 1]);
  const caseOnly = model(small([['casey guest', '3:00:00 PM', '4:00:00 PM', '1h'], ['Casey GUEST', '5:00:00 PM', '6:00:00 PM', '1h']]));
  check('a name that differs in letter case only is the same name; the first spelling is kept', names(caseOnly), [['casey guest', 2]]);

  // Rejoins merged by Teams are one row and stay one attendee.
  check('a row whose duration is shorter than leave minus join (rejoins merged by Teams) is one record', BASE.participants.filter((p) => p.name === 'Rita Rejoin').map((p) => [p.records, p.durationSeconds]), [[1, 5400]]);

  // Missing timing is kept, explicitly.
  check('a record without a join time is kept, the time is null', BASE.participants.filter((p) => p.name === 'Noor Late').map((p) => [p.firstJoin, p.lastLeave, p.durationSeconds]), [[null, T('18:00:00'), 1800]]);
  check('a record without a leave time and duration is kept, both are null', BASE.participants.filter((p) => p.name === 'Omar Open').map((p) => [p.firstJoin, p.lastLeave, p.durationSeconds]), [[T('15:10:00'), null, null]]);
  check('records without any time are kept', BASE.participants.filter((p) => /Avery|Riley/.test(p.name)).map((p) => [p.firstJoin, p.lastLeave, p.durationSeconds, p.email !== '']), [[null, null, null, true], [null, null, null, true]]);
  check('they are counted as rows but not in the mean', [BASE.diagnostics.counts.participantRows, BASE.diagnostics.calculated.rowsWithDuration], [25, 22]);
  const allMissing = model(small([['Quinn, Avery', '', '', '', 'avery.quinn@example.org']]));
  check('an export in which no record has a time: totals are 0 and the mean is null, not NaN', [allMissing.diagnostics.calculated, allMissing.participants.length],
    [{ totalDurationSeconds: 0, rowsWithDuration: 0, meanDurationSeconds: null, spanSeconds: 12201 }, 1]);
}

// ================================================================ 10. sections, columns, line ends

console.log('10. sections, columns and line ends');
{
  const same = (text, options) => { const r = parse(text, options); return r.ok ? JSON.stringify(r.attendance) === JSON.stringify(BASE) : r.error.code; };

  check('LF line ends', same(FX.buildExport({ eol: '\n' })), true);
  check('mixed line ends', same(TEXT.replace('2. Participants\r\n', '2. Participants\n')), true);
  check('blank lines anywhere', same(TEXT.replace('2. Participants\r\n', '\r\n\r\n2. Participants\r\n\r\n').replace('Kai Invalid', '\r\nKai Invalid')), true);
  check('the title with its line break is one value', BASE.summary.meetingTitle, 'Synthetic AHG Meeting (September 22 - 24, 2026) on FS_EXAMPLE');
  check('a title of one line', model(FX.buildExport({ titleLines: ['One line'] })).summary.meetingTitle, 'One line');

  check('without the sections that are not read', same(FX.buildExport({ withOtherSections: false })), true);
  check('other content in those sections', same(TEXT.replace('Alex Organizer\tUnmute', 'Someone Else\tSomething New\textra\tcells')), true);
  check('a quoted field that is never closed, in a section that is not read', same(TEXT.replace('4. Meeting Engagement\r\n', '4. Meeting Engagement\r\n"never closed\t')), true);
  check('sections named differently or added after the two that are read', same(TEXT.replace('3. In-Meeting Activities', '3. Something New').replace('4. Meeting Engagement', '7. Another One') + '9. More\r\nx\ty\r\n'), true);
  check('an unknown section before the Summary is skipped', same('0. Preface\r\nsome\ttext\r\n\r\n' + TEXT), true);
  check('section titles in another letter case, with spaces around', same(TEXT.replace('1. Summary', ' 1.  SUMMARY ').replace('2. Participants', '2. participants\t\t')), true);
  check('Participants before Summary', same(TEXT.replace(/^(1\. Summary[\s\S]*?\r\n\r\n)(2\. Participants[\s\S]*?\r\n\r\n)/, '$2$1')), true);

  // Columns.
  const reorder = (text, order) => text.split('\r\n').map((line) => { const c = line.split('\t'); return order.map((i) => (i < c.length ? c[i] : '')).join('\t'); }).join('\r\n');
  const plainRows = FX.PARTICIPANTS.filter((p) => !p.raw);
  const plain = FX.buildExport({ participants: plainRows, withOtherSections: false }).split('2. Participants\r\n');
  const reordered = plain[0] + '2. Participants\r\n' + reorder(plain[1].replace(/\r\n$/, ''), [6, 4, 0, 5, 3, 2, 1]) + '\r\n';
  check('columns in another order, the unknown ones gone', JSON.stringify(model(reordered).participants) === JSON.stringify(model(plain.join('2. Participants\r\n')).participants), true);
  const sameData = (text) => { const m = model(text); return JSON.stringify([m.summary, m.participants]) === JSON.stringify([BASE.summary, BASE.participants]); };
  const wider = model(TEXT.replace('\tSynthetic Future Column', '\tSynthetic Future Column\tAnother\tAnd Another'));
  check('more unknown columns: the same summary and participants', sameData(TEXT.replace('\tSynthetic Future Column', '\tSynthetic Future Column\tAnother\tAnd Another')), true);
  check('every row is then shorter than the header, which is only a count', [wider.diagnostics.counts.shortRows, codes(wider).filter((c) => c === 'SHORT_ROWS').length], [25, 1]);
  check('header names in another letter case', same(TEXT.replace('Name\tFirst Join\tLast Leave\tIn-Meeting Duration\tEmail', 'NAME\tfirst join\tLast leave\tIN-MEETING DURATION\temail')), true);
  check('a row with more cells than the header', sameData(TEXT.replace(/(riley\.short@example\.com\triley\.short@example\.com\tPresenter)/, '$1' + '\tx'.repeat(20))), true);

  const noEmailColumn = model(TEXT.replace('\tEmail\tParticipant ID (UPN)', '\tE-Mail Address\tParticipant ID (UPN)'));
  check('without an Email column: no address at all, a warning, no fallback', [noEmailColumn.participants.filter((p) => p.email).length, noEmailColumn.diagnostics.notes.filter((n) => n.code === 'MISSING_COLUMN').map((n) => n.message)],
    [0, ['The Participants section has no "Email" column.']]);
  const noTimeColumns = model(TEXT.replace('\tFirst Join\tLast Leave\tIn-Meeting Duration', '\tA\tB\tC'));
  check('without the time columns: every participant is kept, without times', [noTimeColumns.participants.every((p) => p.firstJoin === null && p.lastLeave === null && p.durationSeconds === null), noTimeColumns.diagnostics.counts.participantRows,
    codes(noTimeColumns).filter((c) => c === 'MISSING_COLUMN').length], [true, 25, 3]);

  // Short rows.
  check('short rows are padded, not dropped', [BASE.diagnostics.counts.shortRows, BASE.participants.slice(-2).map((p) => p.email)], [2, ['avery.quinn@example.org', 'riley.short@example.com']]);
  const nameOnly = model(small([['Alex Organizer', '3:00:00 PM', '4:00:00 PM', '1h', 'alex.organizer@example.com']]).replace(/^Alex Organizer\t.*$/m, 'Alex Organizer'));
  check('a row that is only a name', nameOnly.participants, [{ name: 'Alex Organizer', company: '', email: '', firstJoin: null, lastLeave: null, durationSeconds: null, records: 1 }]);

  // Rows that cannot be used.
  const nameless = model(variant({ 4: { name: '' }, 5: { name: '(Unverified)' } }));
  check('rows without a name are left out, counted and noted', [nameless.participants.length, nameless.diagnostics.counts.participantRows, nameless.diagnostics.notes.filter((n) => n.code === 'ROW_WITHOUT_NAME')],
    [21, 25, [{ level: 'warning', code: 'ROW_WITHOUT_NAME', message: '2 participant rows have no name and were left out.', rows: [5, 6] }]]);
  const badValues = model(variant({ 1: { duration: 'a while' }, 2: { join: 'soon' } }));
  check('an unreadable duration or time: the record is kept, the value is null, and it is noted',
    [badValues.participants[1].durationSeconds, badValues.participants[2].firstJoin, badValues.participants.length, badValues.diagnostics.notes.filter((n) => /UNREADABLE/.test(n.code)).map((n) => [n.code, n.rows])],
    [null, null, 23, [['UNREADABLE_TIME', [3]], ['UNREADABLE_DURATION', [2]]]]);

  // Comma-separated.
  const comma = model(FX.buildExport({ delimiter: ',' }));
  check('a comma-separated export gives the same summary and participants', [comma.source.delimiter, JSON.stringify(comma.summary) === JSON.stringify(BASE.summary), JSON.stringify(comma.participants) === JSON.stringify(BASE.participants)], ['comma', true, true]);
  check('and the same diagnostics', JSON.stringify(comma.diagnostics) === JSON.stringify(BASE.diagnostics), true);
}

// ================================================================ 11. fatal problems

console.log('11. fatal problems');
{
  check('not text', [errorCode(null), errorCode(undefined), errorCode(42), errorCode({})], ['NOT_A_TEAMS_EXPORT', 'NOT_A_TEAMS_EXPORT', 'NOT_A_TEAMS_EXPORT', 'NOT_A_TEAMS_EXPORT']);
  check('empty', [errorCode(''), errorCode('  \r\n \r\n')], ['NOT_A_TEAMS_EXPORT', 'NOT_A_TEAMS_EXPORT']);
  check('some other file', [errorCode('TDoc,Title,Source\r\nS4aP260001,Agenda,Chair\r\n'), errorCode('Dear all,\n\nplease find attached.\n'), errorCode('{"v":1}')], ['NOT_A_TEAMS_EXPORT', 'NOT_A_TEAMS_EXPORT', 'NOT_A_TEAMS_EXPORT']);
  check('its message', parse('TDoc,Title\r\n').error.message, 'This is not a Teams attendance export: it has no "Summary" and no "Participants" section.');
  check('an export in another language (other section titles) is refused, not guessed at', errorCode(TEXT.replace('1. Summary', '1. Zusammenfassung').replace('2. Participants', '2. Teilnehmer')), 'NOT_A_TEAMS_EXPORT');

  check('no Summary section', errorCode(TEXT.replace('1. Summary\r\n', '')), 'SUMMARY_MISSING');
  check('no Participants section', [errorCode(TEXT.replace('2. Participants\r\n', '')), errorCode(TEXT.split('2. Participants')[0])], ['PARTICIPANTS_MISSING', 'PARTICIPANTS_MISSING']);
  check('a Participants section without a Name column', [errorCode(TEXT.replace('Name\tFirst Join', 'Full Name\tFirst Join')), parse(TEXT.replace('Name\tFirst Join', 'Full Name\tFirst Join')).error.message],
    ['HEADERS_UNUSABLE', 'The "Participants" section has no "Name" column.']);
  check('an empty Participants section', errorCode(FX.buildExport({ participants: [], withOtherSections: false }).replace(/^Name\t.*\r\n/m, '')), 'HEADERS_UNUSABLE');
  check('a Participants section with a header and no rows is an export without participants, not an error',
    [errorCode(FX.buildExport({ participants: [] })), model(FX.buildExport({ participants: [] })).participants], ['ok', []]);

  check('a quoted field that is never closed, in the Summary', errorCode(TEXT.replace('FS_EXAMPLE"', 'FS_EXAMPLE')), 'MALFORMED_QUOTING');
  check('a quoted field that is never closed, in the Participants', [errorCode(TEXT.replace('Kai Invalid', '"Kai Invalid')), parse(TEXT.replace('Kai Invalid', '"Kai Invalid')).error.message],
    ['MALFORMED_QUOTING', 'The file cannot be read: its quoting is broken.']);
  check('the same in an export with no other quoted cell after it', errorCode(FX.buildExport({ participants: FX.PARTICIPANTS.filter((p) => !p.raw), withOtherSections: false }).replace('Kai Invalid', '"Kai Invalid')), 'MALFORMED_QUOTING');
  check('text after a closing quote', errorCode(TEXT.replace('FS_EXAMPLE"', 'FS_EXAMPLE" (draft)')), 'MALFORMED_QUOTING');
  check('a quote that swallows the rest of the file before any section', errorCode('"' + TEXT), 'MALFORMED_QUOTING');

  check('no start time in the Summary', errorCode(TEXT.replace(/^Start time\t.*\r\n/m, '')), 'SUMMARY_START_MISSING');
  check('a start time that cannot be read', [errorCode(TEXT.replace(/^Start time\t.*$/m, 'Start time\tMonday afternoon')), errorCode(TEXT.replace(/^Start time\t.*$/m, 'Start time\t31.02.26, 14:00:00'), {})],
    ['SUMMARY_START_UNREADABLE', 'SUMMARY_START_UNREADABLE']);

  check('too large', [errorCode('1. Summary\r\n' + 'x'.repeat(2000001)), errorCode('1. Summary\r\nStart time\t' + FX.SUMMARY.startTime + '\r\n2. Participants\r\nName\r\n' + 'Sam Rivera\r\n'.repeat(5001))], ['TOO_LARGE', 'TOO_LARGE']);
  check('a session date that is not a date', [errorCode(TEXT, { expectedDate: '22.09.2026' }), errorCode(TEXT, { expectedDate: '2026-02-30' })], ['INVALID_EXPECTED_DATE', 'INVALID_EXPECTED_DATE']);

  check('nothing is thrown for any of these', thrown(() => [null, '', '"', '\t', '1. Summary', '2. Participants\r\n', '1. Summary\r\n2. Participants\r\nName\r\n', TEXT.slice(0, 200), TEXT.slice(200)].forEach((t) => parse(t, {}))), null);
  check('a fatal result never carries a partial model', [parse('1. Summary\r\nStart time\t' + FX.SUMMARY.startTime).attendance, parse(TEXT, { expectedDate: '2026-09-23' }).attendance], [null, null]);
}

// ================================================================ 12. dates of the whole export

console.log('12. dates of the whole export; the session date');
{
  /** The fixture with every date written by `rewrite(month, day, year2, time)`. */
  const redate = (rewrite, text) => (text || TEXT).replace(/(\d{1,2})\/(\d{1,2})\/(\d{2}), (\d{1,2}:\d{2}:\d{2} [AP]M)/g, (all, m, d, y, time) => rewrite(m, d, y, time));
  const clock24 = (time) => { const m = time.match(/(\d+):(\d+):(\d+) ([AP])M/); const h = (parseInt(m[1], 10) % 12) + (m[4] === 'P' ? 12 : 0); return (h < 10 ? '0' : '') + h + ':' + m[2] + ':' + m[3]; };
  const sameData = (attendance) => [JSON.stringify(attendance.participants) === JSON.stringify(BASE.participants), attendance.summary.start, attendance.summary.end];
  const EXPECTED = [true, '2026-09-22T14:48:05', '2026-09-22T18:11:26'];

  check('the export as it is, without a session date: only one reading is possible', [model(TEXT, {}).source.dateOrder, JSON.stringify(model(TEXT, {})) === JSON.stringify(BASE)], ['month-first', true]);
  check('for the session of its date', errorCode(TEXT, { expectedDate: '2026-09-22' }), 'ok');
  check('for a session of another date: refused', [errorCode(TEXT, { expectedDate: '2026-09-23' }), parse(TEXT, { expectedDate: '2026-09-23' }).error.message],
    ['SESSION_DATE_MISMATCH', 'This attendance export is for 2026-09-22, not for the session on 2026-09-23.']);

  const dayFirst = redate((m, d, y, time) => d + '/' + m + '/' + y + ', ' + clock24(time));
  check('a day-first export with a 24-hour clock: the same data', [model(dayFirst).source.dateOrder].concat(sameData(model(dayFirst))), ['day-first'].concat(EXPECTED));
  check('raw summary values stay as written', model(dayFirst).summary.raw.startTime, '22/9/26, 14:48:05');
  const german = redate((m, d, y, time) => d + '.' + (m.length < 2 ? '0' + m : m) + '.' + y + ', ' + clock24(time));
  check('dates with dots: the same data, no order needed', [model(german).source.dateOrder].concat(sameData(model(german))), [null].concat(EXPECTED));
  const iso = redate((m, d, y, time) => '20' + y + '-' + (m.length < 2 ? '0' + m : m) + '-' + d + ' ' + clock24(time));
  check('year-first dates: the same data', [model(iso).source.dateOrder].concat(sameData(model(iso))), [null].concat(EXPECTED));
  check('year-first dates for another session: refused', errorCode(iso, { expectedDate: '2026-09-21' }), 'SESSION_DATE_MISMATCH');

  // An export whose every date can be read both ways.
  const ambiguous = redate((m, d, y, time) => '9/10/' + y + ', ' + time);
  check('ambiguous everywhere, no session date: refused', [errorCode(ambiguous, {}), parse(ambiguous, {}).error.message],
    ['AMBIGUOUS_DATE', 'The dates of this export can be read month-first or day-first, and nothing in it decides which. Import it for a session with a date.']);
  check('ambiguous everywhere, session on 10 September: month-first', [model(ambiguous, { expectedDate: '2026-09-10' }).source.dateOrder, model(ambiguous, { expectedDate: '2026-09-10' }).summary.start], ['month-first', '2026-09-10T14:48:05']);
  check('ambiguous everywhere, session on 9 October: day-first', [model(ambiguous, { expectedDate: '2026-10-09' }).source.dateOrder, model(ambiguous, { expectedDate: '2026-10-09' }).summary.start], ['day-first', '2026-10-09T14:48:05']);
  check('every participant time follows the same reading', [model(ambiguous, { expectedDate: '2026-10-09' }).participants[0].firstJoin, model(ambiguous, { expectedDate: '2026-09-10' }).participants[0].firstJoin], ['2026-10-09T14:48:09', '2026-09-10T14:48:09']);
  check('ambiguous everywhere, session on another date: refused, naming both readings', [errorCode(ambiguous, { expectedDate: '2026-09-22' }), parse(ambiguous, { expectedDate: '2026-09-22' }).error.message],
    ['SESSION_DATE_MISMATCH', 'This attendance export is for 2026-09-10 or 2026-10-09, not for the session on 2026-09-22.']);

  // A meeting that runs past midnight.
  const late = small([['Alex Organizer', '', '', '1h', 'alex.organizer@example.com']], { startTime: '9/22/26, 11:00:00 PM', endTime: '9/23/26, 1:00:00 AM', meetingDuration: '2h', averageAttendanceTime: '1h' })
    .replace(/^Alex Organizer\t\t\t/m, 'Alex Organizer\t9/22/26, 11:30:00 PM\t9/23/26, 12:30:00 AM\t');
  const lateModel = model(late, { expectedDate: '2026-09-22' });
  check('past midnight: the session date is the start date; times keep their own dates', [lateModel.summary.start, lateModel.summary.end, lateModel.diagnostics.calculated.spanSeconds, lateModel.participants[0].firstJoin, lateModel.participants[0].lastLeave],
    ['2026-09-22T23:00:00', '2026-09-23T01:00:00', 7200, '2026-09-22T23:30:00', '2026-09-23T00:30:00']);

  check('no time zone anywhere in the model', /CEST|CET|UTC|GMT|Z"|[+-]\d\d:\d\d"/.test(JSON.stringify(BASE)), false);
  check('times are text in the form YYYY-MM-DDTHH:mm:ss, or null', BASE.participants.reduce((all, p) => all.concat([p.firstJoin, p.lastLeave]), [BASE.summary.start, BASE.summary.end])
    .filter((v) => v !== null && !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(v)), []);
}

// ================================================================ 13. pure and deterministic

console.log('13. pure and deterministic');
{
  const copy = TEXT.slice();
  const a = JSON.stringify(parse(TEXT));
  const b = JSON.stringify(parse(TEXT));
  check('the same input gives the same output', a === b, true);
  check('in a fresh sandbox too', JSON.stringify(loadCode().sandbox.parseTeamsAttendanceReport_(TEXT, { expectedDate: SESSION_DATE })) === a, true);
  check('the input and the options are not changed', [TEXT === copy, (() => { const o = { expectedDate: SESSION_DATE }; parse(TEXT, o); return o; })()], [true, { expectedDate: SESSION_DATE }]);
  check('the result can be stored as JSON and read back unchanged', JSON.stringify(JSON.parse(a)) === a, true);

  // No Google service: the parser runs with every service removed.
  const bare = loadCode().sandbox;
  ['PropertiesService', 'DocumentApp', 'UrlFetchApp', 'DriveApp', 'SpreadsheetApp', 'HtmlService', 'ScriptApp', 'LockService', 'Utilities', 'Session', 'Logger']
    .forEach((service) => { bare[service] = new Proxy({}, { get: () => { throw new Error(service + ' was used'); } }); });
  check('it uses no Google service at all', [thrown(() => bare.parseTeamsAttendanceReport_(bare.decodeTeamsAttendanceBytes_(FX.utf16le(TEXT)).text, { expectedDate: SESSION_DATE })),
    JSON.stringify(bare.parseTeamsAttendanceReport_(TEXT, { expectedDate: SESSION_DATE })) === a], [null, true]);
  check('no property was written and nothing was logged in this whole test', [loaded.docProps.getKeys(), loaded.scriptProps.getKeys(), logged], [[], [], []]);

  const source = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const parser = source.slice(source.indexOf('// TEAMS ATTENDANCE (stage C)'), source.indexOf('// AD-HOC ATTENDANCE (stage D)')).replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '');
  check('the parser section names no service, no clock and no random value (source)',
    /PropertiesService|DocumentApp|UrlFetchApp|DriveApp|SpreadsheetApp|HtmlService|ScriptApp|LockService|Utilities|Session\.|getReportStateStore_|Date\.now|new Date\(|Math\.random/.test(parser), false);
  // Stage D: the attendance import is the one caller, each function once.
  const importSection = source.slice(source.indexOf('// AD-HOC ATTENDANCE (stage D)'), source.indexOf('// ARCH-009 -- MEETING-ID RESOLVER CORE'));
  const outsideFeature = source.replace(source.slice(source.indexOf('// TEAMS ATTENDANCE (stage C)'), source.indexOf('// ARCH-009 -- MEETING-ID RESOLVER CORE')), '');
  check('the parser is called by the attendance import only',
    [(importSection.match(/parseTeamsAttendanceReport_\(|decodeTeamsAttendanceBytes_\(/g) || []).sort(), (outsideFeature.match(/parseTeamsAttendanceReport_|decodeTeamsAttendanceBytes_/g) || [])],
    [['decodeTeamsAttendanceBytes_(', 'parseTeamsAttendanceReport_('], []]);
  check('Code.js is version 2.18.1 (2.18.0 added the ad-hoc sessions, 2.18.1 the TDoc upload completion)', (source.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.18.1');
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll Teams attendance parser checks passed.');
process.exitCode = failures ? 1 : 0;
