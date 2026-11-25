/**
 * Issue #28, phase C (docs/ISSUE28_PHASE_C.md): AHG sessions by reservation time, carry-forward, the plenary
 * Agenda/Schedule folders, Meeting Control.
 *
 *   A. parseTdocListDateTime_(): ISO, a Date, the Portal's US form; an ambiguous day/month is not read.
 *   B. assignAdhocSessionByReservation_(): cut-off = end + 6 h as instants; +6 h exactly / +6 h + 1 s; before the first,
 *      after the last; a session of several days; a date only; DST (Europe/Amsterdam, 25 Oct 2026).
 *   C. the resolver of every caller: the rule only when set; manual overrides authoritative; the reservation times
 *      read from the list.
 *   D. carry-forward: undisposed TDocs carried into started sessions, recorded once with the time; disposed ones not
 *      further; a "set" override never; the Session cell "A01 → A02"; no second registration row.
 *   E. the plenary folders of SA4#138 (real listings and the real "Tdoc allocation rev9.docx", 2026-11-23): newest by
 *      revN (the listing times are equal), allocations of the report's TDocs, proposed/ambiguous reallocations, AHG
 *      telco candidates with a confidence, the schedule as a hint; each source fails on its own; cached.
 *   F. Meeting Control: its data on an MBS report, a reallocation accepted into Document Reallocations; sessions from
 *      candidates only in an ad-hoc report.
 *   G. the ledger: without this layer Code.js and ReportCreator.js are those of phase B, byte for byte.
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');
const { makeBlob, unzip } = require('./helpers/fake-zip.js');
const ledger = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`);
  else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const ROOT = path.join(__dirname, '..');
const RAW = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const FIX = path.join(ROOT, 'tests', 'fixtures', 'sa4-138-plenary');
const AGENDA_DIR = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Inbox/Drafts/Plenary/Agenda/';
const SCHEDULE_DIR = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Inbox/Drafts/Plenary/Schedule/';
const REV9 = AGENDA_DIR + 'SA4%23138%20Tdoc%20allocation%20rev9.docx';

function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}
function sandbox(props) {
  const { sandbox: S, docProps } = loadCode({ documentProperties: props || {} });
  S.Utilities.formatDate = (d, zone) => formatInZone(d, zone);
  S.Utilities.base64EncodeWebSafe = (s) => Buffer.from(String(s)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
  S.Utilities.unzip = unzip;
  S.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  S.Logger = { log: () => {} };
  return { S, docProps };
}
const SESSIONS = (list) => JSON.stringify({ v: 1, nextId: list.length + 1, sessions: list });
const ADHOC = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const TWO = [{ id: 's1', label: 'Call 1', date: '2026-09-22', start: '14:00', end: '16:00' }, { id: 's2', label: 'Call 2', date: '2026-09-24', start: '14:00', end: '16:00' }];

// ====================================================================== A
console.log('A. parseTdocListDateTime_()');
{
  const { S } = sandbox();
  const p = (d, raw) => { const x = S.parseTdocListDateTime_(d, raw === undefined ? '' : raw, (dt) => formatInZone(dt, 'UTC')); return x.at || x.dateOnly || (x.ambiguous ? 'ambiguous' : x.unreadable ? 'unreadable' : 'blank'); };
  check('ISO; a Date value; the Portal\'s US form with AM/PM', [p('2026-09-22 19:59:59'), p('', new Date('2026-09-22T20:00:00Z')), p('2/19/2026 3:00:31 PM'), p('9/22/2026 12:05:00 AM')],
    ['2026-09-22T19:59:59', '2026-09-22T20:00:00', '2026-02-19T15:00:31', '2026-09-22T00:05:00']);
  check('day/month without AM/PM: read only when unambiguous (19/02 is 19 February); 05/10 is not guessed', [p('19/02/2026 10:05:32'), p('05/10/2026 10:05:32'), p('05/05/2026 10:05')], ['2026-02-19T10:05:32', 'ambiguous', '2026-05-05T10:05:00']);
  check('a date only; blank; garbage', [p('9/22/2026'), p(''), p('next week')], ['2026-09-22', 'blank', 'unreadable']);
}

// ====================================================================== A2
console.log('A2. the date order of the column, on the real lists; the clock check (decisions of 2026-11-23)');
{
  const { S } = sandbox();
  const FIX = path.join(__dirname, 'fixtures', 'tdoc-list-reservation-dates');
  const col = (name) => fs.readFileSync(path.join(FIX, name), 'utf8').split(/\r?\n/).filter(Boolean);
  const v138 = col('sa4-138.txt');
  const v137 = col('sa4-137-e.txt');
  const readAll = (values) => { const order = S.tdocListDateOrder_(values); return { order: order, parsed: values.map((v) => S.parseTdocListDateTime_(v, '', () => '', order)) }; };
  const a = readAll(v138);
  const b = readAll(v137);
  const unread = (x) => x.parsed.filter((p) => !p.at).length;
  check('SA4#138 (2026-11-23, "13/10/2026 12:34:06"): the column is day/month; all 615 read, none ambiguous (without the column, 361 were refused)',
    [v138.length, a.order, unread(a), v138.map((v) => S.parseTdocListDateTime_(v, '', () => '')).filter((p) => p.ambiguous).length, a.parsed[0].at], [615, 'dmy', 0, 361, '2026-10-13T12:34:06']);
  check('SA4#137-e ("7/17/2026 12:33:23 PM"): month/day with AM/PM; all 482 read', [v137.length, b.order, unread(b), b.parsed[0].at], [482, 'mdy', 0, '2026-07-17T12:33:23']);
  check('the order of a column: no evidence, or a column that contradicts itself: none (each value on its own, 05/10 not guessed)',
    [S.tdocListDateOrder_(['05/10/2026 10:05', '01/02/2026 09:00']), S.tdocListDateOrder_(['13/10/2026 12:00:00', '7/17/2026 1:00:00 PM']), S.parseTdocListDateTime_('05/10/2026 10:05', '', () => '', '').ambiguous],
    ['', '', true]);
  // The clock check: the reservation is never after the upload on one clock.
  const T = (n, res, up) => ({ row: [n], tdocCol: 0, reserved: res ? { at: res } : null, uploaded: up ? { at: up } : null });
  const ok = S.reservationClockCheck_([T('S4aA269001', '2026-09-20T08:00:00', '2026-09-20T08:01:30'), T('S4aA269002', '2026-09-20T08:00:00', '2026-09-21T08:00:00'), T('S4aA269003', '2026-09-20T08:00:00', null)]);
  check('consistent: two checked, no inversion, the shortest gap 1.5 min', [ok.checked, ok.inversions, ok.minGapMinutes], [2, [], 1.5]);
  const bad = S.reservationClockCheck_([T('S4aA269001', '2026-09-20T09:00:00', '2026-09-20T08:01:30'), T('S4aA269002', '2026-09-20T08:00:00', '2026-09-20T07:59:00')]);
  check('uploaded an hour before the reservation as read: an inversion (the clocks disagree); a minute: within the tolerance', [bad.inversions, bad.checked], [['S4aA269001'], 2]);
}

// ====================================================================== B
console.log('B. assignAdhocSessionByReservation_(): end + 6 h, as instants');
{
  const { S } = sandbox();
  const berlin = (w) => S.sessionZoneInstantOfWall_('Europe/Berlin', w);
  const at = (iso) => ({ at: iso, dateOnly: null, unreadable: false });
  const a = (r, sessions, toInstant, clock) => { const x = S.assignAdhocSessionByReservation_(r, sessions || TWO, toInstant || berlin, clock || 'utc'); return [x.sessionId, x.reason]; };
  // Call 1 ends 16:00 CEST = 14:00Z; its cut-off is 20:00Z.
  check('reserved exactly at the cut-off (end + 6 h): that session; one second later: the next', [a(at('2026-09-22T20:00:00')), a(at('2026-09-22T20:00:01'))], [['s1', 'assigned'], ['s2', 'assigned']]);
  check('during a session and between two sessions', [a(at('2026-09-22T13:00:00')), a(at('2026-09-23T09:00:00'))], [['s1', 'assigned'], ['s2', 'assigned']]);
  check('before the first session: the first (said); after the last cut-off: none, never the last', [a(at('2026-09-01T08:00:00')), a(at('2026-09-24T20:00:01'))], [['s1', 'before-first'], [null, 'after-last']]);
  check('no reservation, unreadable, ambiguous: none with the reason', [a(null), a({ at: null, dateOnly: null, unreadable: true, ambiguous: false }), a({ at: null, dateOnly: null, unreadable: true, ambiguous: true })],
    [[null, 'not-reserved'], [null, 'unreadable'], [null, 'ambiguous']]);
  check('a date only: a session when the whole day is in one window; else none', [a({ at: null, dateOnly: '2026-09-23', unreadable: false }), a({ at: null, dateOnly: '2026-09-22', unreadable: false })], [['s2', 'assigned'], [null, 'date-only']]);
  const multi = [{ id: 's1', label: 'F2F', date: '2026-09-21', endDate: '2026-09-23', start: '09:00', end: '17:00' }, { id: 's2', label: 'Call', date: '2026-09-30', start: '14:00', end: '15:00' }];
  check('a session of several days: the cut-off is its last day\'s end + 6 h (17:00 CEST + 6 h = 21:00Z)', [a(at('2026-09-22T23:00:00'), multi), a(at('2026-09-23T21:00:00'), multi), a(at('2026-09-23T21:00:01'), multi)],
    [['s1', 'assigned'], ['s1', 'assigned'], ['s2', 'assigned']]);
  // DST: Europe/Amsterdam goes back on 25 Oct 2026 at 03:00 CEST. A session ending 01:00 CEST (23:00Z on the 24th): cut-off 05:00Z.
  const ams = (w) => S.sessionZoneInstantOfWall_('Europe/Amsterdam', w);
  const night = [{ id: 's1', label: 'Late', date: '2026-10-25', start: '00:00', end: '01:00' }, { id: 's2', label: 'Next', date: '2026-10-26', start: '10:00', end: '11:00' }];
  check('across the DST change the six hours are real time: 05:00:00Z in, 05:00:01Z out (wall-clock +6 h would be 06:00Z)', [a(at('2026-10-25T05:00:00'), night, ams), a(at('2026-10-25T05:00:01'), night, ams)], [['s1', 'assigned'], ['s2', 'assigned']]);
  check('times already on the session clock (clock "session")', a(at('2026-09-22T22:00:00'), TWO, berlin, 'session'), ['s1', 'assigned']);
}

// ====================================================================== C
console.log('C. the resolver of every caller');
{
  const td = (id, reserved, uploaded) => ({ row: [id], tdocCol: 0, reserved: reserved ? { at: reserved, dateOnly: null, unreadable: false } : null, uploaded: uploaded ? { at: uploaded, dateOnly: null, unreadable: false } : null });
  const base = sandbox(Object.assign({}, ADHOC, { ADHOC_SESSIONS: SESSIONS(TWO) })).S;
  const rb = base.makeAdhocTdocSessionResolver_(null);
  check('without the setting: the upload rule of stage B, unchanged (reserved early, uploaded after Call 1)', [rb.rule, rb.automatic(td('S4aA269001', '2026-09-20T08:00:00', '2026-09-23T08:00:00')).sessionId], ['upload', 's2']);
  const ruled = sandbox(Object.assign({}, ADHOC, { ADHOC_SESSIONS: SESSIONS(TWO), ADHOC_SESSION_RULE: 'reservation-6h',
    ADHOC_TDOC_SESSIONS: JSON.stringify({ v: 1, clock: 'utc', add: {}, set: { s2: ['S4aA269003'] } }) })).S;
  const r = ruled.makeAdhocTdocSessionResolver_(null);
  check('with "reservation-6h": by reservation (reserved before Call 1 ended + 6 h, uploaded later)', [r.rule, r.automatic(td('S4aA269001', '2026-09-20T08:00:00', '2026-09-23T08:00:00')).sessionId, r.cell(td('S4aA269001', '2026-09-20T08:00:00'))], ['reservation-6h', 's1', 'A01']);
  check('a manual "set" stays authoritative', r.cell(td('S4aA269003', '2026-09-20T08:00:00')), 'A02');
  check('no reservation time: no session (–), the reason kept', [r.cell(td('S4aA269004', null)), r.automatic(td('S4aA269004', null)).reason], ['–', 'not-reserved']);
  // The reservation column of the list.
  const sheet = { getRange: () => ({ getDisplayValues: () => [['Reservation date'], ['9/22/2026 9:59:59 PM'], ['05/10/2026 10:05'], ['']] }) };
  const read = ruled.readTdocReservationTimesForSessions_({ getSpreadsheetTimeZone: () => 'UTC' }, sheet, ['TDoc', 'Reservation date'], [['TDoc', 'Reservation date'], ['a', ''], ['b', ''], ['c', '']], null);
  check('the "Reservation date" column read like the uploads; its order from the column (AM/PM: month/day, so 05/10 is 10 May)', read.slice(1).map((x) => x.at || (x.ambiguous ? 'ambiguous' : '') || 'none'), ['2026-09-22T21:59:59', '2026-05-10T10:05:00', 'none']);
  const sheet2 = { getRange: () => ({ getDisplayValues: () => [['Reservation date'], ['05/10/2026 10:05'], ['01/02/2026 09:00']] }) };
  const read2 = ruled.readTdocReservationTimesForSessions_({ getSpreadsheetTimeZone: () => 'UTC' }, sheet2, ['TDoc', 'Reservation date'], [['TDoc', 'Reservation date'], ['a', ''], ['b', '']], null);
  check('... a column that does not show its order: ambiguous, not guessed', read2.slice(1).map((x) => (x.ambiguous ? 'ambiguous' : x.at)), ['ambiguous', 'ambiguous']);
  check('... and not read at all without the setting', base.readTdocReservationTimesForSessions_({ getSpreadsheetTimeZone: () => 'UTC' }, sheet, ['TDoc', 'Reservation date'], [[], []], null), null);
}

// ====================================================================== D
console.log('D. carry-forward');
{
  const { S } = sandbox();
  const startOf = (s) => S.sessionZoneInstantOfWall_('Europe/Berlin', s.date + 'T' + s.start + ':00');
  const THREE = TWO.concat([{ id: 's3', label: 'Call 3', date: '2026-10-01', start: '14:00', end: '16:00' }]);
  const now = Date.UTC(2026, 8, 25, 12);
  const d = (reg, carried, undisposed, override) => S.decideAdhocCarry_(reg, carried, THREE, undisposed, now, startOf, override, 'T').map((c) => c.s);
  check('undisposed, registered in Call 1: carried into Call 2 (started), not Call 3 (not yet)', d(['s1'], [], true), ['s2']);
  check('disposed: carried no further, what it has stays', [d(['s1'], [], false), d(['s1'], [{ s: 's2', at: 'x' }], false)], [[], ['s2']]);
  check('a "set" override: never carried; no registration session: none', [d(['s1'], [{ s: 's2', at: 'x' }], true, { mode: 'set', sessions: ['s1'] }), d([], [], true)], [[], []]);
  // The update's step with the real resolver.
  // Carry-forward is its own switch (decision of 2026-11-23); here with the reservation rule as well.
  const props = Object.assign({}, ADHOC, { ADHOC_SESSIONS: SESSIONS(TWO), ADHOC_SESSION_RULE: 'reservation-6h', ADHOC_CARRY_FORWARD_MODE: 'on' });
  const { S: U, docProps } = sandbox(props);
  U.reportRunSnapshot_ = () => null;
  const status = { S4aA269001: 'available', S4aA269002: 'agreed', S4aA269003: 'reserved' };
  const index = new Map(Object.keys(status).map((n) => [n, { n: n }]));
  U.lookupTdocTableInIndex_ = (ix, n) => ix.get(n) || null;
  U.provisionalStatusOfTable_ = (t) => status[t.n];
  const tds = Object.keys(status).map((n) => ({ row: [n], tdocCol: 0, reserved: { at: '2026-09-21T10:00:00', dateOnly: null, unreadable: false } }));
  const resolver = U.makeAdhocTdocSessionResolver_(null);
  check('the update: the two undisposed TDocs carried into Call 2 (both calls are over), the agreed one not', U.syncAdhocCarryForwardOnUpdate_(resolver, tds, index, null), 2);
  check('... and the same resolver shows it at once (the Session column of this update, not the next)', resolver.cell(tds[0]), 'A01 → A02');
  const stored = { tdocs: U.readAdhocCarry_({ getProperty: (k) => docProps.getProperty(k) }) };
  check('recorded with the time (provenance), once', [Object.keys(stored.tdocs).sort(), stored.tdocs.S4aA269001.length, /^\d{4}-\d{2}-\d{2}T/.test(stored.tdocs.S4aA269001[0].at), U.syncAdhocCarryForwardOnUpdate_(U.makeAdhocTdocSessionResolver_(null), tds, index, null)],
    [['S4aA269001', 'S4aA269003'], 1, true, 0]);
  const r2 = U.makeAdhocTdocSessionResolver_(null);
  check('the Session cell shows it: "A01 → A02"; the disposed TDoc "A01"', [r2.cell(tds[0]), r2.cell(tds[1]), r2.effective(tds[0])], ['A01 → A02', 'A01', ['s1', 's2']]);
  status.S4aA269001 = 'agreed';
  U.syncAdhocCarryForwardOnUpdate_(U.makeAdhocTdocSessionResolver_(null), tds, index, null);
  check('disposed later: it keeps "A01 → A02" (provenance), nothing more', U.makeAdhocTdocSessionResolver_(null).cell(tds[0]), 'A01 → A02');
  // Carry-forward off (no setting, even with the reservation rule): nothing carried or shown -- an existing ad-hoc report does not change by itself.
  const { S: V } = sandbox(Object.assign({}, ADHOC, { ADHOC_SESSIONS: SESSIONS(TWO), ADHOC_SESSION_RULE: 'reservation-6h', ADHOC_CARRY_FORWARD: docProps.getProperty('ADHOC_CARRY_FORWARD') }));
  V.reportRunSnapshot_ = () => null; V.lookupTdocTableInIndex_ = U.lookupTdocTableInIndex_; V.provisionalStatusOfTable_ = U.provisionalStatusOfTable_;
  check('carry-forward not switched on (the reservation rule alone): no carry-forward written or shown', [V.syncAdhocCarryForwardOnUpdate_(V.makeAdhocTdocSessionResolver_(null), tds, index, null), V.makeAdhocTdocSessionResolver_(null).carried(tds[2])], [0, []]);
  // Carry-forward on with the upload rule (stage B): independent of the reservation rule.
  const { S: X, docProps: xProps } = sandbox(Object.assign({}, ADHOC, { ADHOC_SESSIONS: SESSIONS(TWO), ADHOC_CARRY_FORWARD_MODE: 'on' }));
  X.reportRunSnapshot_ = () => null; X.lookupTdocTableInIndex_ = U.lookupTdocTableInIndex_; X.provisionalStatusOfTable_ = (t) => ({ S4aA269001: 'available', S4aA269002: 'agreed', S4aA269003: 'reserved' })[t.n];
  const upl = Object.keys(status).map((n) => ({ row: [n], tdocCol: 0, uploaded: { at: '2026-09-21T10:00:00', dateOnly: null, unreadable: false } }));
  const xr = X.makeAdhocTdocSessionResolver_(null);
  check('carry-forward on with the upload rule: the undisposed TDocs uploaded for Call 1 carried into Call 2, shown at once', [xr.rule, X.syncAdhocCarryForwardOnUpdate_(xr, upl, index, null), xr.cell(upl[0]), xr.cell(upl[1]), !!xProps.getProperty('ADHOC_CARRY_FORWARD')], ['upload', 2, 'A01 → A02', 'A01', true]);
  // A carried session before a session added by hand later is not shown ("A01, A03", never "A01, A03 → A02").
  const { S: W } = sandbox(Object.assign({}, ADHOC, { ADHOC_SESSION_RULE: 'reservation-6h', ADHOC_CARRY_FORWARD_MODE: 'on', ADHOC_SESSIONS: SESSIONS(TWO.concat([{ id: 's3', label: 'Call 3', date: '2026-10-01', start: '14:00', end: '16:00' }])),
    ADHOC_CARRY_FORWARD: JSON.stringify({ v: 1, tdocs: { S4aA269001: [{ s: 's2', at: 'x' }] } }), ADHOC_TDOC_SESSIONS: JSON.stringify({ v: 1, clock: 'utc', add: { s3: ['S4aA269001'] }, set: {} }) }));
  check('a carried session before a later manual one is not shown', W.makeAdhocTdocSessionResolver_(null).cell(tds[0]), 'A01, A03');
}

// ====================================================================== E
console.log('E. the plenary Agenda and Schedule folders of SA4#138');
const agendaHtml = fs.readFileSync(path.join(FIX, 'agenda-folder-2026-11-23.html.txt'), 'utf8');
const scheduleHtml = fs.readFileSync(path.join(FIX, 'schedule-folder-2026-11-23.html.txt'), 'utf8');
const rev9 = fs.readFileSync(path.join(FIX, 'SA4#138 Tdoc allocation rev9.docx'));
const doc9 = (S) => S.wordDocumentText_(unzip(makeBlob(rev9)).filter((b) => b.getName() === 'word/document.xml')[0].getDataAsString('UTF-8'));
{
  const { S } = sandbox();
  check('the folders derived from FTP_BASE (…/Docs/)', S.plenarySourceFolders_({ FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/' }, { getProperty: () => null }), { agenda: AGENDA_DIR, schedule: SCHEDULE_DIR });
  const files = S.parseFtpFolderListing_(agendaHtml, AGENDA_DIR);
  check('the listing: the ten allocation files (not the parent folders or the sort links)', [files.length, files[0].name, files.every((f) => f.url.indexOf(AGENDA_DIR) === 0)], [10, 'SA4#138 Tdoc allocation rev0.docx', true]);
  check('rev3-rev9 have the same listing time; the newest is taken by its revN: rev9', [files.filter((f) => f.listedAt === '2026-11-23T07:28').length, S.newestAllocationFile_(files).name], [7, 'SA4#138 Tdoc allocation rev9.docx']);
  check('the schedule: four files, newest listed first (a hint only)', S.scheduleFiles_(S.parseFtpFolderListing_(scheduleHtml, SCHEDULE_DIR)).map((f) => f.name.replace('S4-261859 Proposed meeting schedule for SA4#138', '…')), ['…-v2.xlsx', '…-MBSChair.xlsx', '….xlsx', '… Draft.xlsx']);
  const doc = doc9(S);
  const numbers = ['S4-261967', 'S4-261979', 'S4-261930', 'S4-262019', 'S4-262015', 'S4-262139', 'S4-262026', 'S4-261697'];
  const alloc = S.parseTdocAllocations_(doc, numbers);
  check('the allocations of the report\'s TDocs (third column; chains "2015->2139" both)', [alloc.byTdoc['S4-261967'], alloc.byTdoc['S4-261979'], alloc.byTdoc['S4-261930'], alloc.byTdoc['S4-262015'], alloc.byTdoc['S4-262139']],
    [['8.5'], ['8.6'], ['8.7'], ['10.6'], ['10.6']]);
  check('a number of another meeting ("was 1697") is no allocation; a year in the future-meetings row (20) is outside the part of an RTC report: no proposal', [alloc.byTdoc['S4-261697'] || null, S.proposedReallocations_(alloc, { 'S4-262026': '10.8' }, {})], [null, []]);
  check('only allocations within the own part of the agenda of the report: a plenary row (14) never moves an MBS TDoc', S.proposedReallocations_({ byTdoc: { 'S4-1': ['14'], 'S4-2': ['8.4', '14'] } }, { 'S4-1': '8.5', 'S4-2': '8.5' }, {}), [{ tdoc: 'S4-2', from: '8.5', to: '8.4', kind: 'proposed' }]);
  check('a TDoc reallocated by hand to another item: shown as "manual", never proposed (the row is kept)', S.proposedReallocations_({ byTdoc: { 'S4-3': ['8.6'] } }, { 'S4-3': '8.5' }, { 'S4-3': '8.7' }), [{ tdoc: 'S4-3', from: '8.5', to: '8.6', kind: 'manual', replaces: '8.7' }]);
  const report = { 'S4-261967': '8.5', 'S4-261979': '8.5', 'S4-261930': '8.7', 'S4-262019': '8.4' };
  check('proposed: 1979 8.5 -> 8.6, 2019 8.4 -> 8.6; one already reallocated there: not proposed', S.proposedReallocations_(alloc, report, { 'S4-262019': '8.6' }), [{ tdoc: 'S4-261979', from: '8.5', to: '8.6', kind: 'proposed' }]);
  check('a TDoc listed under two items: ambiguous, never proposed', S.proposedReallocations_({ byTdoc: { 'S4-1': ['8.5', '8.6'] } }, { 'S4-1': '8.5' }, {}), [{ tdoc: 'S4-1', from: '8.5', to: '8.5, 8.6', kind: 'ambiguous' }]);
  const cands = S.parseAhgCandidates_(doc, 2026);
  const find = (re) => cands.filter((c) => re.test(c.text))[0] || {};
  check('"Tuesday 8th December 2026, 15-16 CET": 8 December (never the 20 of 2026)', [find(/Tuesday 8th December 2026/).date, find(/Tuesday 8th December 2026/).start, find(/Tuesday 8th December 2026/).end], ['2026-12-08', '15:00', '16:00']);
  check('the F2F meeting 26-29 January 2027: from the 26th, several days (until the 29th), no time taken from the range', [find(/F2F Audio/).date, find(/F2F Audio/).endDate, find(/F2F Audio/).start], ['2027-01-26', '2027-01-29', '']);
  check('the MBS telco of 17 Dec: high confidence', [find(/MBS SWG AHG Telco \(Dec 17/).date, find(/MBS SWG AHG Telco \(Dec 17/).start, find(/MBS SWG AHG Telco \(Dec 17/).end, find(/MBS SWG AHG Telco \(Dec 17/).confidence], ['2026-12-17', '15:30', '17:30', 'high']);
  check('"CEST" in December: medium, said why', [find(/15:00-17:00 CEST/).confidence, /CEST does not fit/.test(find(/15:00-17:00 CEST/).why)], ['medium', true]);
  check('"9:00:10:00": the time not read, medium; 22:00-24:00 read as until 23:59', [find(/DaCAS/).confidence, find(/DaCAS/).start, find(/Jan 27, 2027, 22:00-24:00/).end], ['medium', '', '23:59']);
  check('the pending F2F meeting: medium; the empty template lines "(Date, time, Host)" are no candidates', [find(/F2F Audio SWG AH meeting/).confidence || 'none', cands.some((c) => /\(Date, time, Host\)/.test(c.text))], ['medium', false]);
}
{
  // readPlenarySources_: fetched once, cached; a failing source does not stop the other.
  const { S } = sandbox();
  const fetched = [];
  const cache = {}; S.CacheService = { getDocumentCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; } }) };
  let agendaDown = false;
  S.UrlFetchApp = { fetch: (url) => {
    fetched.push(url);
    if (url === AGENDA_DIR) return agendaDown ? { getResponseCode: () => 503 } : { getResponseCode: () => 200, getContentText: () => agendaHtml };
    if (url === SCHEDULE_DIR) return { getResponseCode: () => 200, getContentText: () => scheduleHtml };
    if (url === REV9) return { getResponseCode: () => 200, getBlob: () => makeBlob(rev9) };
    return { getResponseCode: () => 404 };
  } };
  const cfg = { FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', MEETING_DATE: 'November 16, 2026' };
  const one = S.readPlenarySources_(cfg, ['S4-261979']);
  check('the sources: rev9, its allocation of the report\'s TDoc, AHG candidates, the schedule', [one.agenda.ok, one.agenda.file, one.agenda.rev, one.agenda.allocations.byTdoc['S4-261979'], one.candidates.length > 5, one.schedule.ok, one.schedule.files.length], [true, 'SA4#138 Tdoc allocation rev9.docx', 9, ['8.6'], true, true, 4]);
  const n = fetched.length;
  const two = S.readPlenarySources_(cfg, ['S4-261979']);
  check('the second read comes from the cache (no fetch)', [fetched.length - n, two.agenda.from], [0, 'cache']);
  Object.keys(cache).forEach((k) => delete cache[k]);
  agendaDown = true;
  const down = S.readPlenarySources_(cfg, ['S4-261979']);
  check('the Agenda folder down: its error named, the Schedule still read', [down.agenda.ok, /HTTP 503/.test(down.agenda.error), down.schedule.ok], [false, true, true]);
  check('no Docs folder in FTP_BASE: said, nothing fetched', (() => { const k = fetched.length; const x = S.readPlenarySources_({ FTP_BASE: 'https://example.org/x/' }, []); return [x.agenda.ok, /set PLENARY_AGENDA_DIR_URL/.test(x.agenda.error), fetched.length - k]; })(), [false, true, 0]);
}

// ====================================================================== F
console.log('F. Meeting Control');
{
  const MAIN = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_NAME: 'SA4#138', MEETING_ID: '60780', REPORT_SUFFIX: 'MBS', MEETING_DATE: 'November 16, 2026',
    FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx', AGENDA_TDOC: 'S4-261857' };
  const AGENDA_138 = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'fixtures', 'agenda-S4-261857.google-items.json'), 'utf8')).items.map((i) => ({ number: i.number, title: i.title, level: i.level, heading: 'NORMAL', text: '' }));
  const r = templateReport({ props: MAIN, agenda: AGENDA_138, tableSeparators: true, endParagraph: true, tdocs: [
    { id: 'S4-261967', agenda: '8.5', status: 'agreed', title: 'AMD pCR', uploaded: true },
    { id: 'S4-261979', agenda: '8.5', status: 'available', title: 'Energy pCR', uploaded: true },
    { id: 'S4-261930', agenda: '8.7', status: 'reserved', title: 'QUIC eval', uploaded: false }] });
  r.s.HtmlService = { createHtmlOutput: (html) => { const out = { setWidth: () => out, setHeight: () => out, getContent: () => html }; return out; } };
  r.ui.dialogs = []; r.ui.showModalDialog = (o, t) => r.ui.dialogs.push([t, o.getContent()]);
  const cache = {}; r.s.CacheService = { getDocumentCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; } }) };
  r.s.Utilities.unzip = unzip;
  r.s.Utilities.base64EncodeWebSafe = (s) => Buffer.from(String(s)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
  const realFetch = r.s.UrlFetchApp.fetch;
  r.s.UrlFetchApp = { fetch: (url, o) => {
    if (url === AGENDA_DIR) return { getResponseCode: () => 200, getContentText: () => agendaHtml };
    if (url === SCHEDULE_DIR) return { getResponseCode: () => 200, getContentText: () => scheduleHtml };
    if (url === REV9) return { getResponseCode: () => 200, getBlob: () => makeBlob(rev9) };
    return realFetch(url, o);
  } };
  r.build(); r.body._state.dirty = false; r.update(); r.body._state.dirty = false;
  r.exec(() => r.s.openMeetingControl());
  check('Meeting Control opens at once (its data is loaded by the page, with a progress line)', [r.ui.dialogs[0][0], /Reading the report and the meeting's sources/.test(r.ui.dialogs[0][1]), /meetingControlData\(\)/.test(r.ui.dialogs[0][1])], ['Meeting Control', true, true]);
  const data = r.exec(() => r.s.meetingControlData());
  const rowOf = (n) => data.rows.filter((x) => x.tdoc === n)[0] || { tags: [] };
  check('its rows: the changed allocation of 1979 (8.5 -> 8.6, rev9), the new (reserved) 1930, the unresolved; 1967 agreed and in place: not listed',
    [data.errors, rowOf('S4-261979').tags, rowOf('S4-261930').tags.indexOf('new') !== -1, !!data.rows.filter((x) => x.tdoc === 'S4-261967')[0], data.proposals], [[], ['unresolved', 'allocation'], true, false, [{ tdoc: 'S4-261979', from: '8.5', to: '8.6', kind: 'proposed' }]]);
  check('the sources shown with their file, revision and observation time; candidates and schedule present', [data.sources.agenda.file, data.sources.agenda.rev, !!data.sources.observedAt, data.sources.candidates.length > 5, data.sources.schedule.files.length], ['SA4#138 Tdoc allocation rev9.docx', 9, true, true, 4]);
  const said = r.exec(() => r.s.acceptProposedReallocations([{ tdoc: 'S4-261979', from: '8.5', to: '8.6', source: 'SA4#138 Tdoc allocation rev9.docx' }]));
  const map = r.exec(() => r.s.getReallocationMap_());
  check('accepted: one row in Document Reallocations (8.5 -> 8.6, the source named); nothing else moved', [/Added to Document Reallocations: S4-261979 → 8.6/.test(said), map['S4-261979'].new, /rev9/.test(map['S4-261979'].reason), r.order()], [true, '8.6', true, ['S4-261979', 'S4-261967', 'S4-261930'].filter((n) => r.order().indexOf(n) !== -1).length === 3 ? r.order() : 'moved']);
  const again = r.exec(() => r.s.meetingControlData());
  check('the next view: no longer proposed', again.proposals, []);
  check('the session rule: refused in a main report', (() => { try { r.exec(() => r.s.setAdhocSessionRule('reservation-6h')); return 'set'; } catch (e) { return /ad-hoc report/.test(e.message); } })(), true);
  check('sessions from candidates: refused in a main report', (() => { try { r.exec(() => r.s.createSessionsFromCandidates([{ group: 'MBS', date: '2026-12-17', start: '15:30', end: '17:30' }])); return 'created'; } catch (e) { return /ad-hoc report only/.test(e.message); } })(), true);
  const a = templateReport({ tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'agreed', uploaded: true }] });
  a.build();
  check('the reservation rule is not switched on without a preview first', (() => { try { a.exec(() => a.s.setAdhocSessionRule('reservation-6h')); return 'set'; } catch (e) { return /Preview/.test(e.message); } })(), true);
  check('... switching back to the upload rule needs none', /by upload time/.test(a.exec(() => a.s.setAdhocSessionRule('upload'))), true);
  check('carry-forward is not switched on without its own preview; off needs none; refused in a main report',
    [(() => { try { a.exec(() => a.s.setAdhocCarryForward(true)); return 'set'; } catch (e) { return /Preview/.test(e.message); } })(), a.docProps.getProperty('ADHOC_CARRY_FORWARD_MODE'),
      /Carry-forward off/.test(a.exec(() => a.s.setAdhocCarryForward(false))), (() => { try { r.exec(() => r.s.setAdhocCarryForward(true)); return 'set'; } catch (e) { return /ad-hoc report/.test(e.message); } })()],
    [true, null, true, true]);
  // The preview of the rule checks the clocks on the list itself (decision of 2026-11-23: not switched on unverified).
  // The preview of the rule checks the clocks on the list itself (decision of 2026-11-23: not switched on unverified).
  const id = (k) => 'S4aA2690' + String(k).padStart(2, '0');
  const pad = (n) => String(n).padStart(2, '0');
  // n TDocs reserved on 20 September (a day/month column, as SA4#138), each uploaded `gap(k)` minutes later.
  const pv = (n, gap) => { const x = templateReport({ props: { ADHOC_SESSIONS: SESSIONS(TWO) }, tdocs: Array.from({ length: n }, (v, k) => ({ id: id(k + 1), agenda: '4.3', status: 'available', uploaded: true,
    reservedAt: '20/09/2026 ' + pad(8 + Math.floor(k / 4)) + ':' + pad((k % 4) * 10) + ':00', uploadedAt: new Date(Date.UTC(2026, 8, 20, 8 + Math.floor(k / 4), (k % 4) * 10) + gap(k) * 60000).toISOString().slice(0, 19).replace('T', ' ') })) }); x.build(); return x; };
  const good = pv(25, (k) => (k === 3 ? 2 : 90 + k));
  const gp = good.exec(() => good.s.previewAdhocSessionRule());
  check('preview, 25 TDocs whose times agree (shortest gap 2 min): the clock check passes, the preview is stamped', [gp.ok, gp.clock.checked, gp.clock.inversions, gp.clock.minGapMinutes, !!good.docProps.getProperty('ADHOC_SESSION_RULE_PREVIEW')], [true, 25, [], 2, true]);
  check('... the switch uses the preview up: a second switch needs a new one', [/by reservation time/.test(good.exec(() => good.s.setAdhocSessionRule('reservation-6h'))), good.docProps.getProperty('ADHOC_SESSION_RULE_PREVIEW')], [true, null]);
  const few = pv(1, () => 3);
  const fp = few.exec(() => few.s.previewAdhocSessionRule());
  check('preview, one TDoc with both times: too few to tell -- refused, nothing stamped', [fp.ok, /only 1 TDoc\(s\) of the list have both/.test(fp.error), few.docProps.getProperty('ADHOC_SESSION_RULE_PREVIEW')], [false, true, null]);
  const ahead = pv(25, (k) => (k === 5 ? -120 : 3));
  const ap = ahead.exec(() => ahead.s.previewAdhocSessionRule());
  check('preview, a TDoc uploaded 2 hours before its reservation as read (a clock ahead): refused, named', [ap.ok, /S4aA269006\) were uploaded before their reservation time/.test(ap.error)], [false, true]);
  const behind = pv(25, (k) => 360 + k);
  behind.docProps.setProperty('ADHOC_SESSION_RULE_PREVIEW', new Date().toISOString());
  const bhp = behind.exec(() => behind.s.previewAdhocSessionRule());
  check('preview, every upload 6 hours or more after its reservation (a clock behind, e.g. Calgary time): refused; an earlier preview taken back, so no switch',
    [bhp.ok, /may be behind the upload clock/.test(bhp.error), behind.docProps.getProperty('ADHOC_SESSION_RULE_PREVIEW'), (() => { try { behind.exec(() => behind.s.setAdhocSessionRule('reservation-6h')); return 'set'; } catch (e) { return /Preview/.test(e.message); } })()],
    [false, true, null, true]);
  // While the rule is on, each update checks the clock again: an alert in the report state (Meeting Control), cleared when it agrees.
  const live = pv(25, (k) => (k === 3 ? 2 : 90 + k));
  live.docProps.setProperty('ADHOC_SESSION_RULE', 'reservation-6h');
  const tdocsOf = (x, g) => x.tdocs.map((t, k) => ({ row: [t.id], tdocCol: 0, reserved: { at: '2026-09-20T' + pad(8 + Math.floor(k / 4)) + ':' + pad((k % 4) * 10) + ':00' }, uploaded: { at: new Date(Date.UTC(2026, 8, 20, 8 + Math.floor(k / 4), (k % 4) * 10) + g(k) * 60000).toISOString().slice(0, 19) } }));
  check('an update with a list that agrees: no alert', [live.exec(() => live.s.recheckReservationClockOnUpdate_(tdocsOf(live, (k) => (k === 3 ? 2 : 90)), null)), live.docProps.getProperty('ADHOC_RESERVATION_CLOCK_ALERT')], ['', null]);
  live.exec(() => live.s.recheckReservationClockOnUpdate_(tdocsOf(live, (k) => (k === 7 ? -90 : 3)), null));
  check('a regenerated list on another clock: the alert is kept and shown by Meeting Control', [/S4aA269008/.test(live.docProps.getProperty('ADHOC_RESERVATION_CLOCK_ALERT') || ''),
    live.exec(() => live.s.meetingControlData()).errors.some((x) => /^Reservation clock: .*were uploaded before their reservation/.test(x))], [true, true]);
  live.exec(() => live.s.recheckReservationClockOnUpdate_(tdocsOf(live, () => 3), null));
  check('... cleared when the list agrees again', live.docProps.getProperty('ADHOC_RESERVATION_CLOCK_ALERT'), null);
  // Carry-forward: its preview end to end, then the switch (which uses the preview up).
  const cp = good.exec(() => good.s.previewAdhocCarryForward());
  check('carry-forward preview: read-only listing (new carries and those recorded before), stamped', [cp.ok, Array.isArray(cp.changes), Array.isArray(cp.kept), cp.count >= 25, !!good.docProps.getProperty('ADHOC_CARRY_FORWARD_MODE_PREVIEW'), good.docProps.getProperty('ADHOC_CARRY_FORWARD_MODE')], [true, true, true, true, true, null]);
  good.docProps.setProperty('ADHOC_CARRY_FORWARD', JSON.stringify({ v: 1, tdocs: { S4aA269001: [{ s: 's2', at: '2026-09-24T12:00:00.000Z' }] } }));
  check('... a carry recorded while it was on before is listed as shown again', good.exec(() => good.s.previewAdhocCarryForward()).kept, [{ tdoc: 'S4aA269001', into: 'A02' }]);
  check('... the switch after it: on, the preview used up', [/Carry-forward on/.test(good.exec(() => good.s.setAdhocCarryForward(true))), good.docProps.getProperty('ADHOC_CARRY_FORWARD_MODE'), good.docProps.getProperty('ADHOC_CARRY_FORWARD_MODE_PREVIEW')], [true, 'on', null]);
  a.docProps.setProperty('ADHOC_CARRY_FORWARD_MODE_PREVIEW', new Date().toISOString());
  check('... after a preview of the last 30 minutes: on', [/Carry-forward on/.test(a.exec(() => a.s.setAdhocCarryForward(true))), a.docProps.getProperty('ADHOC_CARRY_FORWARD_MODE')], [true, 'on']);
  const refusedNoZone = a.exec(() => a.s.createSessionsFromCandidates([{ group: 'Audio', date: '2026-12-04', start: '14:00', end: '15:00', zone: '' }]));
  check('a candidate with a time but no known zone: refused, nothing created (never a guess)', [/Nothing created/.test(refusedNoZone), a.docProps.getProperty('ADHOC_SESSIONS')], [true, null]);
  const made = a.exec(() => a.s.createSessionsFromCandidates([{ group: 'Audio', date: '2026-12-04', start: '14:00', end: '15:00', zone: 'CET' }]));
  check('in an ad-hoc report: created through the writer of Configure Sessions, its CET time in the meeting zone (Europe/Berlin here)', [/Sessions saved/.test(made), JSON.parse(a.docProps.getProperty('ADHOC_SESSIONS')).sessions.map((s) => [s.label, s.date, s.start, s.end])], [true, [['Audio', '2026-12-04', '14:00', '15:00']]]);
}

{
  const { S } = sandbox();
  check('a candidate converted from its quoted zone: 15:30-17:30 CET on 17 Dec in London is 14:30-16:30; UTC as given; an unknown zone null',
    [S.convertCandidateToMeetingZone_({ date: '2026-12-17', start: '15:30', end: '17:30', zone: 'CET' }, 'Europe/London'), S.convertCandidateToMeetingZone_({ date: '2026-12-17', start: '15:30', end: '', zone: 'UTC' }, 'Europe/London').start,
      S.convertCandidateToMeetingZone_({ date: '2026-12-17', start: '15:30', end: '', zone: 'EST' }, 'Europe/London')],
    [{ date: '2026-12-17', start: '14:30', end: '16:30' }, '15:30', null]);
}

// ====================================================================== F2
console.log('F2. Meeting Control: one read of the report, no plenary folders in an ad-hoc report (2026-11-25: 281 s)');
{
  const MAIN = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_NAME: 'SA4#138', MEETING_ID: '60780', REPORT_SUFFIX: 'MBS', MEETING_DATE: 'November 16, 2026',
    FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx', AGENDA_TDOC: 'S4-261857' };
  const AGENDA_138 = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'fixtures', 'agenda-S4-261857.google-items.json'), 'utf8')).items.map((i) => ({ number: i.number, title: i.title, level: i.level, heading: 'NORMAL', text: '' }));
  const TDOCS = [
    { id: 'S4-261967', agenda: '8.5', status: 'agreed', title: 'AMD pCR', uploaded: true },
    { id: 'S4-261979', agenda: '8.5', status: 'available', title: 'Energy pCR', uploaded: true },
    { id: 'S4-261930', agenda: '8.7', status: 'reserved', title: 'QUIC eval', uploaded: false }];
  const prepare = (o) => {
    const x = templateReport(Object.assign({ props: MAIN, agenda: AGENDA_138, tableSeparators: true, endParagraph: true, tdocs: TDOCS }, o));
    const cache = {}; x.s.CacheService = { getDocumentCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; } }) };
    x.s.Utilities.base64EncodeWebSafe = (s) => Buffer.from(String(s)).toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
    x.build(); x.body._state.dirty = false; x.update(); x.body._state.dirty = false;
    // After build and update (the helper sets its own fetch there): the plenary folders answer 404, everything is recorded.
    x.fetched = [];
    const realFetch = x.s.UrlFetchApp.fetch;
    x.s.UrlFetchApp = { fetch: (url, opt) => { x.fetched.push(url); if (/\/Plenary\//.test(url)) return { getResponseCode: () => 404, getContentText: () => '' }; return realFetch(url, opt); } };
    return x;
  };
  // The three DocumentApp walks of before, counted.
  const walks = (x) => {
    let n = 0;
    const names = ['registrationRowsOfReport_', 'reportTdocScope_', 'readReportDispositions_'];
    const real = {};
    names.forEach((k) => { real[k] = x.s[k]; x.s[k] = function () { n++; return real[k].apply(this, arguments); }; });
    return () => { names.forEach((k) => { x.s[k] = real[k]; }); return n; };
  };
  const viaDocs = prepare({});
  const viaApp = prepare({ docs: false });
  const getsBefore = viaDocs.api.calls.filter((k) => k.call === 'get').length;
  const stopDocs = walks(viaDocs);
  const d1 = viaDocs.exec(() => viaDocs.s.meetingControlData());
  const docsWalks = stopDocs();
  const getsAfter = viaDocs.api.calls.filter((k) => k.call === 'get').length;
  const stopApp = walks(viaApp);
  const d2 = viaApp.exec(() => viaApp.s.meetingControlData());
  const appWalks = stopApp();
  const strip = (d) => d.rows.map((r) => [r.tdoc, r.title, r.agenda, r.status, r.plenary, r.tags.slice().sort().join(',')].join('|'));
  check('with the Docs API: at most ONE read of the report and none of the three DocumentApp walks (registration, scope, statuses)', [d1.timing.via, getsAfter - getsBefore <= 1, docsWalks, d1.rows.length > 0], ['docs-api', true, 0, true]);
  check('without it: the three DocumentApp walks, as before', [d2.timing.via, appWalks], ['documentapp', 3]);
  check('both ways the same rows (TDoc, title, agenda, status, plenary, tags), no errors', [strip(d1), d1.errors, d2.errors], [strip(d2), [], []]);
  check('... and the rows are real: the new (reserved) TDoc and the unresolved one listed', [d1.rows.some((r) => r.tdoc === 'S4-261930' && r.tags.indexOf('new') !== -1), d1.rows.some((r) => r.tags.indexOf('unresolved') !== -1)], [true, true]);
  check('the time of each part is returned (and logged)', ['read', 'tables', 'registration', 'scope', 'statuses', 'sources', 'total'].filter((k) => typeof d1.timing[k] !== 'number'), []);
  const tables = viaDocs.s.plenaryDocsTables_(viaDocs.s.plenarySnapshot_(viaDocs.s.currentReportDocument_().getId()).content);
  const fromDocs = viaDocs.s.registrationRowsOfDocsTables_(tables);
  check('registrationRowsOfDocsTables_ gives what registrationRowsOfReport_ gives from DocumentApp', [fromDocs.length > 0, JSON.stringify(fromDocs)], [true, JSON.stringify(viaDocs.exec(() => viaDocs.s.registrationRowsOfReport_(viaDocs.body)))]);
  check('a main report still reads its plenary Agenda and Schedule folders', [viaDocs.fetched.some((u) => /\/Plenary\/Agenda\//.test(u)), viaDocs.fetched.some((u) => /\/Plenary\/Schedule\//.test(u)), !d1.sources.skipped], [true, true, true]);
  const adhoc = templateReport({ tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'agreed', uploaded: true }, { id: 'S4aA269002', agenda: '4.4', status: 'reserved', uploaded: false }] });
  const acache = {}; adhoc.s.CacheService = { getDocumentCache: () => ({ get: (k) => acache[k] || null, put: (k, v) => { acache[k] = v; } }) };
  adhoc.build();
  const aFetched = [];
  const aReal = adhoc.s.UrlFetchApp.fetch;
  adhoc.s.UrlFetchApp = { fetch: (url, opt) => { aFetched.push(url); return aReal(url, opt); } };
  const d3 = adhoc.exec(() => adhoc.s.meetingControlData());
  check('an ad-hoc report: the plenary folders are not read (nothing fetched for them), and the page says why', [d3.isAdhoc, aFetched.filter((u) => /\/Plenary\/(Agenda|Schedule)\//.test(u)).length, /ad-hoc meeting: the plenary Agenda and Schedule folders/.test(d3.sources.skipped), d3.timing.sources, d3.errors], [true, 0, true, undefined, []]);
  check('... its rows are read as before (the new TDoc listed)', d3.rows.filter((r) => r.tdoc === 'S4aA269002' && r.tags.indexOf('new') !== -1).length, 1);
  check('the page shows the note instead of the sources when they are skipped', /D\.sources&&D\.sources\.skipped/.test(viaDocs.s.meetingControlHtml_()), true);
}

// ====================================================================== G
console.log('G. the ledger');
{
  const git = (spec) => { try { return execFileSync('git', ['show', spec], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, ''); } catch (e) { return null; } };
  const view = ledger.codeWithoutIssue28C(RAW);
  check('every change of the layer is in Code.js exactly once', view.problems, []);
  // Phase B pinned by its commit (8adde44: with phase A's update journal); the branch name only as a fallback.
  const b = git('8adde44:Code.js') || git('feature/i28-b-provisional-revisions:Code.js');
  if (b === null) console.log('  note: the phase B branch is not available; the byte comparison is skipped.');
  else check('without it, Code.js is that of phase B, byte for byte', view.code === b, true);
  const creator = fs.readFileSync(path.join(ROOT, 'template', 'ReportCreator.js'), 'utf8').replace(/\r/g, '');
  const cb = git('8adde44:template/ReportCreator.js') || git('feature/i28-b-provisional-revisions:template/ReportCreator.js');
  check('ReportCreator.js: one menu item; without it, that of phase B', [ledger.creatorWithoutIssue28C(creator).problems, cb === null || ledger.creatorWithoutIssue28C(creator).code === cb], [[], true]);
  check('the version is 2.30.0', (RAW.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.30.0');
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((x) => x.slice(9, -1));
  const bodyOf = (code, n) => { const at = code.indexOf('\nfunction ' + n + '('); if (at === -1) return null; return code.slice(at, code.indexOf('\n}\n', at)); };
  check('its functions are those of its section; three released functions changed by one hook each', [names(RAW).filter((n) => names(view.code).indexOf(n) === -1).sort(), names(view.code).filter((n) => bodyOf(RAW, n) !== bodyOf(view.code, n)).sort()],
    [ledger.I28C_SECTION_FUNCTIONS.slice().sort(), ledger.I28C_CHANGED_FUNCTIONS.slice().sort()]);
  check('the local listing agent: no new reference to 10.10.10.10', (RAW.match(/10\.10\.10\.10/g) || []).length, (view.code.match(/10\.10\.10\.10/g) || []).length);
  check('stage B\'s section still never reads "Reservation date" (the rule is this layer\'s, opt-in)', (() => { const a = RAW.indexOf('AD-HOC TDOC SESSIONS (stage B)'); const z = RAW.indexOf('AD-HOC OPENING (stage E)'); return RAW.slice(a, z).indexOf('Reservation date'); })(), -1);
}

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
