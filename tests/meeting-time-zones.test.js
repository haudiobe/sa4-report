/**
 * Meeting-local and selected IANA time zones (after T-2026.10.23; GitHub issue #4; docs/MEETING_TIME_ZONES.md).
 *
 *   A. how a zone is written for a date: SA4#138 Calgary (America/Edmonton) MST in November, MDT in summer;
 *      Berlin CET/CEST; the DST changes of both; zones without an abbreviation; both rule paths
 *      (Apps Script's Utilities.formatDate "Z", and the JavaScript runtime's Intl) agree;
 *   B. Meeting local from the Portal's own record: SA4#138, SA4#137 Prague, online meetings, Shanghai, a zone
 *      name or an offset that does not fit, no location -- never a guess, never the editor's/script's zone;
 *   C. the opening sentence: an IANA zone by date; a legacy "UTC+2" exactly as stored;
 *   D. sessions: one day, several days across a DST change, a session through the night of the change;
 *   E. the upload clock: UTC upload times into the meeting zone; legacy reports on the script zone as before;
 *   F. Configure Meeting: Meeting local, an explicit zone, refusals before anything is written, a Resolve
 *      keeps an explicit zone and a legacy offset, a main meeting (SA4#138) gets its zone and no start time;
 *   G. the dialog: the field, its state, Resolve selects Meeting local only when nothing is stored;
 *   H. manual edits of the opening survive an update; the update writes no opening;
 *   I. what is left of "CEST" in the code is the quoting of external deadlines, not a meeting time;
 *   J. what changed in Code.js and ReportCreator.js, against T-2026.10.23.
 *   K. regressions: Alberta (ABT) with old and new tz data, Calgary on 16 November 2026, Berlin, DST changes, CENTRAL.
 *   L. the save path after review: unchanged choices, "Not set", another meeting, another clock, CENTRAL sync.
 *
 * Fakes only: no network, no Google account. The SA4#138 and SA4#137 records are the Portal's public meeting
 * records (GetMeetings, meetings 60780 and 60779) as fetched on 2026-11-11.
 *
 * Run: node tests/meeting-time-zones.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { templateReport } = require('./helpers/template-report.js');
const { configurableReport } = require('./helpers/config-dialog.js');
const { codeWithoutMainMeetingNumber, codeWithoutMeetingTimeZones, creatorWithoutMeetingTimeZones, MEETING_TIME_ZONES_SECTION_START, MEETING_TIME_ZONES_SECTION_END, MEETING_TIME_ZONES_SECTION_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log('  ok   ' + name);
  } else {
    failures++;
    console.log('  FAIL ' + name + '\n         expected ' + e + '\n         actual   ' + a);
  }
}
const thrown = (fn) => { try { fn(); return null; } catch (e) { return e.message; } };
const plain = (v) => JSON.parse(JSON.stringify(v));

const ROOT = path.join(__dirname, '..');
const GM = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86172-85916-getmeetings.json'), 'utf8'));
const SA4_138 = { Title: '3GPPSA4#138', StartDate: '2026-11-16 09:00:00', EndDate: '2026-11-20 17:30:00', StartTimeZone: '(GMT-07:00) Mountain Time (US & Canada)', EndTimeZone: '(GMT-07:00) Mountain Time (US & Canada)', Location: 'Calgary', Country: 'CA' };
const SA4_137 = { Title: '3GPPSA4#137', StartDate: '2026-10-12 09:00:00', StartTimeZone: '(GMT+02.00)  Belgrade, Bratislava, Budapest, Ljubljana, Prague', Location: 'Prague', Country: 'CZ' };
const ZONE_KEYS = ['MEETING_START_TIME', 'MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'];

// Alberta keeps UTC-6 all year from 1 November 2026, officially ABT (MEETING_TIME_ZONE_RULE_CHANGES_): Calgary does
// not go back to MST, whatever copy of the tz database the runtime has (section K has both kinds). The Portal's
// "(GMT-07:00) Mountain Time (US & Canada)" for SA4#138 is out of date. The DST changes themselves are checked with
// America/Denver, which keeps them.
const intlOffset = (zone, iso) => { const d = new Date(iso); const p = {}; new Intl.DateTimeFormat('en-US', { timeZone: zone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(d).forEach((x) => { p[x.type] = x.value; }); return Math.round((Date.UTC(+p.year, +p.month - 1, +p.day, +p.hour % 24, +p.minute, +p.second) - d.getTime()) / 60000); };
const RUNTIME_ALBERTA = intlOffset('America/Edmonton', '2026-11-16T16:00:00Z');
const CALGARY_NOV = 'ABT';
console.log('note: this runtime\'s own rules (tzdata ' + (process.versions.tz || '?') + ') have Calgary on ' + (RUNTIME_ALBERTA === -360 ? 'UTC-6' : 'UTC-7') + ' on 16 November 2026; the report uses UTC-6 (ABT) either way.');
const SA4_138_ALBERTA = Object.assign({}, SA4_138, { StartTimeZone: '(GMT-06:00) Mountain Time (Alberta)' });
const DENVER = Object.assign({}, SA4_138, { Location: 'Denver', Country: 'US' });

/** Utilities.formatDate as Apps Script does it, for the two patterns used here; by the runtime's Intl. */
function appsScriptFormatDate(date, zone, pattern) {
  const parts = {};
  new Intl.DateTimeFormat('en-US', { timeZone: zone, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  if (pattern === 'Z') {
    const off = Math.round((Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour % 24, +parts.minute, +parts.second) - Math.floor(date.getTime() / 1000) * 1000) / 60000);
    const a = Math.abs(off);
    return (off < 0 ? '-' : '+') + String(Math.floor(a / 60)).padStart(2, '0') + String(a % 60).padStart(2, '0');
  }
  if (pattern === "yyyy-MM-dd'T'HH:mm:ss") return parts.year + '-' + parts.month + '-' + parts.day + 'T' + String(+parts.hour % 24).padStart(2, '0') + ':' + parts.minute + ':' + parts.second;
  return date.toISOString();
}
/** A sandbox; `rules`: 'apps' (Utilities.formatDate knows the zones) or 'intl' (the stub of the test loader: the Intl path). */
function sandbox(props, rules) {
  const loaded = loadCode({ documentProperties: props || {} });
  if (rules === 'apps') loaded.sandbox.Utilities.formatDate = appsScriptFormatDate;
  return loaded;
}

// ====================================================================== A
console.log('A. how a zone is written for a date');
{
  ['apps', 'intl'].forEach((rules) => {
    const s = sandbox({}, rules).sandbox;
    const show = (zone, day, time) => s.meetingTimeZoneDisplay_(zone, day, time);
    check(rules + ': Mountain Time -- MST in November, MDT in summer (Denver)', [show('America/Denver', '2026-11-16', '09:00'), show('America/Denver', '2026-11-20', '17:30'), show('America/Denver', '2026-07-15', '09:00')], ['MST', 'MST', 'MDT']);
    check(rules + ': SA4#138 Calgary -- what its clocks show: MST in January 2026, MDT in summer, ' + CALGARY_NOV + ' on 16 November 2026', [show('America/Edmonton', '2026-01-15', '09:00'), show('America/Edmonton', '2026-07-15', '09:00'), show('America/Edmonton', '2026-11-16', '09:00')], ['MST', 'MDT', CALGARY_NOV]);
    check(rules + ': Berlin -- CET in winter, CEST in summer', [show('Europe/Berlin', '2026-01-20', '10:00'), show('Europe/Berlin', '2026-08-24', '09:00')], ['CET', 'CEST']);
    check(rules + ': the DST changes in Europe (29 March, 25 October 2026): by the wall-clock time of the day', [
      show('Europe/Berlin', '2026-03-28', '12:00'), show('Europe/Berlin', '2026-03-29', '01:30'), show('Europe/Berlin', '2026-03-29', '03:30'),
      show('Europe/Berlin', '2026-10-24', '12:00'), show('Europe/Berlin', '2026-10-25', '00:30'), show('Europe/Berlin', '2026-10-25', '04:00'), show('Europe/Berlin', '2026-10-26', '09:00')],
      ['CET', 'CET', 'CEST', 'CEST', 'CEST', 'CET', 'CET']);
    check(rules + ': the DST changes of Mountain Time (8 March, 1 November 2026)', [show('America/Denver', '2026-03-07', '12:00'), show('America/Denver', '2026-03-08', '12:00'), show('America/Denver', '2026-10-31', '12:00'), show('America/Denver', '2026-11-01', '12:00')],
      ['MST', 'MDT', 'MDT', 'MST']);
    check(rules + ': a zone is written by the offset its clocks show, never by "winter" -- Calgary in January 2027', show('America/Edmonton', '2027-01-15', '09:00'), 'ABT');
    check(rules + ': zones without an abbreviation are written as their offset on that date; Sydney in its summer', [show('Asia/Shanghai', '2026-09-28', '09:00'), show('Asia/Kolkata', '2026-02-09', '09:00'), show('Asia/Seoul', '2026-05-01', '09:00'), show('Europe/Istanbul', '2026-05-01', '09:00'), show('Australia/Sydney', '2026-01-10', '09:00'), show('Australia/Sydney', '2026-07-10', '09:00'), show('UTC', '2026-07-10', '09:00')],
      ['UTC+8', 'IST', 'KST', 'UTC+3', 'AEDT', 'AEST', 'UTC']);
  });
  const s = sandbox({}, 'intl').sandbox;
  check('an offset of an earlier release is written exactly as stored, whatever the date', ['UTC+2', 'UTC-3:30', 'UTC'].map((z) => s.meetingTimeZoneDisplay_(z, '2026-12-01', '09:00')), ['UTC+2', 'UTC-3:30', 'UTC']);
  check('an IANA zone without a known date: its id, not a guessed abbreviation', [s.meetingTimeZoneDisplay_('America/Edmonton', '', '09:00'), s.meetingTimeZoneDisplay_('America/Edmonton', 'November 16', '')], ['America/Edmonton', 'America/Edmonton']);
  check('not a zone of the list, or not a zone: nothing', ['Mars/Olympus', 'CEST', 'GMT+2', '', null, 'america/edmonton'].map((z) => s.meetingTimeZoneDisplay_(z, '2026-11-16', '09:00')), ['', '', '', '', '', '']);
  check('the stored forms: an offset, or a zone of the list', [['UTC+2', 'UTC', 'America/Edmonton', 'Europe/Berlin', 'Asia/Seoul'].map(s.isValidMeetingTimeZoneSetting_), ['CEST', 'MST', 'Mars/Olympus', 'Europe/berlin', '', null].map(s.isValidMeetingTimeZoneSetting_)],
    [[true, true, true, true, true], [false, false, false, false, false, false]]);
  check('the names in the list of Configure Meeting', ['America/Denver', 'Europe/Berlin', 'Asia/Shanghai', 'Australia/Sydney', 'UTC'].map((z) => s.meetingTimeZoneOptionLabel_(z)),
    ['America/Denver (MST/MDT)', 'Europe/Berlin (CET/CEST)', 'Asia/Shanghai (UTC+8)', 'Australia/Sydney (AEST/AEDT)', 'UTC']);
  check('an offset the table has no abbreviation for is written as the offset', s.meetingTimeZoneDisplay_('America/Sao_Paulo', '2026-11-16', '09:00'), 'UTC-3');
  check('every zone of the list is known to the runtime\'s rules', Object.keys(s.MEETING_TIME_ZONES_).filter((z) => s.meetingTimeZoneOffsetMinutes_(z, Date.UTC(2026, 0, 15)) === null), []);
  check('wall-clock time to instant: 09:00 in Denver on 16 November is 16:00 UTC; in Calgary as its clocks are; 09:00 in Berlin on 24 August is 07:00 UTC',
    [new Date(s.meetingWallClockToInstant_('America/Denver', '2026-11-16', '09:00')).toISOString(), new Date(s.meetingWallClockToInstant_('America/Edmonton', '2026-11-16', '09:00')).toISOString(), new Date(s.meetingWallClockToInstant_('Europe/Berlin', '2026-08-24', '09:00')).toISOString()],
    ['2026-11-16T16:00:00.000Z', '2026-11-16T15:00:00.000Z', '2026-08-24T07:00:00.000Z']);
}

// ====================================================================== B
console.log('B. Meeting local, from the Portal\'s own record');
{
  ['apps', 'intl'].forEach((rules) => {
    const s = sandbox({}, rules).sandbox;
    s.Session.getScriptTimeZone = () => 'Asia/Tokyo';
    const local = (m) => plain(s.meetingLocalTimeZoneFromPortal_(m.StartDate, m.StartTimeZone, m.Country));
    check(rules + ': SA4#138 Calgary as the Portal has it ("(GMT-07:00) Mountain Time (US & Canada)", CA): refused -- Calgary is UTC-6 (ABT) then; the user selects the zone', local(SA4_138),
      { zone: '', reason: 'the Portal gives UTC-7 for the start, but America/Edmonton is UTC-6 (ABT) then', source: '' });
    check(rules + ': Calgary with a zone that fits (Windows\' "Mountain Time (Alberta)" where Alberta keeps UTC-6): America/Edmonton -- never the script\'s zone (Asia/Tokyo here)', local(SA4_138_ALBERTA).zone, 'America/Edmonton');
    check(rules + ': the same name in the US: America/Denver', local(DENVER).zone, 'America/Denver');
    check(rules + ': SA4#137 Prague (CZ, +02:00 on 12 October): Europe/Prague', local(SA4_137).zone, 'Europe/Prague');
    check(rules + ': an online meeting (no country): the zone of the name -- 86172, 86178', [local(GM.getMeetings86172[0]).zone, local(Object.assign({}, GM.getMeetings86172[0], { StartDate: '2026-09-22 15:00:00' })).zone], ['Europe/Brussels', 'Europe/Brussels']);
    check(rules + ': 85916 Shanghai: Asia/Shanghai', local(GM.getMeetings85916[0]).zone, 'Asia/Shanghai');
  });
  const s = sandbox({}, 'intl').sandbox;
  const local = (m) => plain(s.meetingLocalTimeZoneFromPortal_(m.StartDate, m.StartTimeZone, m.Country));
  check('an offset that does not fit the zone on that date is not taken: -06:00 for Denver in November', local(Object.assign({}, DENVER, { StartTimeZone: '(GMT-06:00) Mountain Time (US & Canada)' })),
    { zone: '', reason: 'the Portal gives UTC-6 for the start, but America/Denver is UTC-7 (MST) then', source: '' });
  check('a name the report does not know, a partial name, no name, no offset: not resolved, with the reason', [
    local(Object.assign({}, SA4_138, { StartTimeZone: '(GMT-07:00) Mountain Time' })).reason,
    local(Object.assign({}, SA4_138, { StartTimeZone: '(GMT-07:00) Chihuahua, La Paz, Mazatlan' })).zone,
    local(Object.assign({}, SA4_138, { StartTimeZone: '' })).reason,
    local(Object.assign({}, SA4_138, { StartTimeZone: 'Mountain Time (US & Canada)' })).reason,
    local(Object.assign({}, SA4_138, { StartDate: '' })).reason],
    ['the Portal\'s time zone "Mountain Time" is not one this report knows', '', 'the Portal gives no time zone for this meeting', 'the Portal gives no time zone for this meeting', 'the Portal gives no start date or UTC offset for this meeting']);
  check('a location alone names no zone: "Calgary" without the Portal\'s zone is not resolved', local({ StartDate: SA4_138.StartDate, StartTimeZone: '', Location: 'Calgary', Country: 'CA' }).zone, '');
  check('a midnight start (no time in the Portal) is checked at noon of that day', local(Object.assign({}, DENVER, { StartDate: '2026-11-16 00:00:00' })).zone, 'America/Denver');
  check('what the Portal gives MEETING_TIME_ZONE: the verified zone, else the offset as before, else nothing', [
    s.meetingTimeZoneFromPortal_(DENVER.StartDate, DENVER.StartTimeZone, 'US'), s.meetingTimeZoneFromPortal_(SA4_138.StartDate, '(GMT-07:00) Chihuahua, La Paz, Mazatlan', 'MX'), s.meetingTimeZoneFromPortal_(SA4_138.StartDate, 'Mountain Time', 'CA')],
    ['America/Denver', 'UTC-7', '']);
  check('every zone the Portal names map to is a zone of the list', s.PORTAL_TIME_ZONE_NAMES_.reduce((all, e) => all.concat([e.zone], Object.values(e.byCountry || {})), []).filter((z) => !s.isAllowedMeetingTimeZoneId_(z)), []);
  check('the preview line of Configure Meeting', [s.meetingLocalTimeZonePreview_(DENVER.StartDate, DENVER.StartTimeZone, 'US').text, s.meetingLocalTimeZonePreview_(SA4_138.StartDate, '', 'CA').text],
    ['Meeting local: America/Denver (from the Portal: Mountain Time (US & Canada), US), MST on November 16, 2026.', 'Meeting local is not available: the Portal gives no time zone for this meeting. Select a time zone.']);
  check('...for SA4#138 as the Portal has it', s.meetingLocalTimeZonePreview_(SA4_138.StartDate, SA4_138.StartTimeZone, 'CA').text,
    'Meeting local is not available: the Portal gives UTC-7 for the start, but America/Edmonton is UTC-6 (ABT) then. Select a time zone.');
}

// ====================================================================== C
console.log('C. the opening sentence');
{
  const s = sandbox({}, 'intl').sandbox;
  const sentence = (date, time, zone) => s.buildMeetingOpeningSentence_(date, time, zone);
  check('SA4#138 Calgary: 09:00 on November 16, 2026, in the zone its clocks show (' + CALGARY_NOV + ')', sentence('November 16, 2026', '09:00', 'America/Edmonton'), '<Chair> opens the session on November 16, 2026 at 09:00 ' + CALGARY_NOV + '.');
  check('Mountain Time in November and in summer: MST, MDT', [sentence('November 16, 2026', '09:00', 'America/Denver'), sentence('July 15, 2026', '09:00', 'America/Edmonton')],
    ['<Chair> opens the session on November 16, 2026 at 09:00 MST.', '<Chair> opens the session on July 15, 2026 at 09:00 MDT.']);
  check('Berlin, winter and summer', [sentence('January 20, 2026', '10:00', 'Europe/Berlin'), sentence('August 24, 2026', '09:00', 'Europe/Berlin')],
    ['<Chair> opens the session on January 20, 2026 at 10:00 CET.', '<Chair> opens the session on August 24, 2026 at 09:00 CEST.']);
  check('a legacy offset exactly as before', [sentence('October 1, 2026', '15:30', 'UTC+2'), sentence('November 5, 2026', '15:30', 'UTC+2')],
    ['<Chair> opens the session on October 1, 2026 at 15:30 UTC+2.', '<Chair> opens the session on November 5, 2026 at 15:30 UTC+2.']);
  check('no start time: the zone of that day at noon; no date: the zone\'s id; no zone: the placeholder', [sentence('November 16, 2026', '', 'America/Edmonton'), sentence('', '09:00', 'America/Edmonton'), sentence('November 16, 2026', '09:00', ''), sentence('November 16, 2026', '09:00', 'CEST')],
    ['<Chair> opens the session on November 16, 2026 at <start> ' + CALGARY_NOV + '.', '<Chair> opens the session on <meeting date> at 09:00 America/Edmonton.', '<Chair> opens the session on November 16, 2026 at 09:00 <time zone>.', '<Chair> opens the session on November 16, 2026 at 09:00 <time zone>.']);
  const stored = (extra) => sandbox(Object.assign({ MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_START_TIME: '09:00', MEETING_START_BASIS: '60780|November 16, 2026' }, extra), 'intl').sandbox.getMeetingStartForOpening_();
  check('what the build reads: an IANA zone of the meeting, a legacy offset; neither for another meeting', [plain(stored({ MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'local' })), plain(stored({ MEETING_TIME_ZONE: 'UTC-7' })), plain(stored({ MEETING_TIME_ZONE: 'America/Edmonton', MEETING_START_BASIS: '60779|November 16, 2026' }))],
    [{ time: '09:00', zone: 'America/Edmonton' }, { time: '09:00', zone: 'UTC-7' }, { time: '', zone: '' }]);
}

// ====================================================================== D
console.log('D. sessions');
{
  const s = sandbox({}, 'intl').sandbox;
  const S = (date, endDate, start, end, label) => Object.assign({ id: 's1', label: label || '', date: date, start: start || '', end: end || '' }, endDate ? { endDate: endDate } : {});
  const when = (session, zone) => s.adhocSessionWhenText_(session, zone);
  check('one day, with the meeting zone: the times and the zone of that day', [when(S('2026-09-22', '', '15:00', '18:00'), 'Europe/Berlin'), when(S('2026-11-16', '', '09:00', '17:30', 'Plenary'), 'America/Edmonton'), when(S('2026-11-16', '', '09:00'), 'America/Denver'), when(S('2026-11-16', '', '', '17:30'), 'America/Denver')],
    ['September 22, 2026, 15:00–18:00 CEST', 'Plenary, November 16, 2026, 09:00–17:30 ' + CALGARY_NOV, 'November 16, 2026, 09:00 MST', 'November 16, 2026, until 17:30 MST']);
  check('several days across the change to winter time: each time names the zone of its own day', [when(S('2026-10-23', '2026-10-26', '09:00', '17:00', 'AHG'), 'Europe/Berlin'), when(S('2026-10-30', '2026-11-02', '09:00', '17:00'), 'America/Denver')],
    ['AHG, October 23–26, 2026 (from 09:00 CEST on the first day, until 17:00 CET on the last day)', 'October 30 – November 2, 2026 (from 09:00 MDT on the first day, until 17:00 MST on the last day)']);
  check('a session through the night of the change: both zones', when(S('2026-10-25', '', '00:30', '04:00'), 'Europe/Berlin'), 'October 25, 2026, 00:30 CEST–04:00 CET');
  check('without a meeting zone (every report before, a legacy offset): exactly as before', [when(S('2026-09-22', '', '15:00', '18:00'), ''), when(S('2026-10-23', '2026-10-26', '09:00', '17:00', 'AHG')), when(S('2026-09-22', '', '15:00', '18:00'), 'UTC+2')],
    ['September 22, 2026, 15:00–18:00', 'AHG, October 23–26, 2026 (from 09:00 on the first day, until 17:00 on the last day)', 'September 22, 2026, 15:00–18:00 UTC+2']);
  check('the Session administration lines carry it', s.buildAdhocOpeningLines_([S('2026-11-16', '', '09:00', '17:30')], {}, 'America/Edmonton').map((l) => l.text), ['A01: November 16, 2026, 09:00–17:30 ' + CALGARY_NOV]);
  check('a session without times names no zone', when(S('2026-11-16', '', '', ''), 'America/Edmonton'), 'November 16, 2026');
  // Which zone the sessions are in: the meeting's IANA zone, for the meeting the report is configured for.
  const zoneOf = (props) => sandbox(props, 'intl').sandbox.meetingSessionTimeZone_();
  const BASE = { MEETING_TYPE: 'adhoc', MEETING_ID: '85916', MEETING_DATE: 'September 28, 2026', MEETING_START_BASIS: '85916|September 28, 2026' };
  check('the sessions\' zone: an IANA zone of this meeting; not a legacy offset, not another meeting\'s, not one outside the list', [
    zoneOf(Object.assign({}, BASE, { MEETING_TIME_ZONE: 'Asia/Shanghai', MEETING_TIME_ZONE_MODE: 'local' })), zoneOf(Object.assign({}, BASE, { MEETING_TIME_ZONE: 'UTC+8' })),
    zoneOf(Object.assign({}, BASE, { MEETING_TIME_ZONE: 'Asia/Shanghai', MEETING_START_BASIS: '85916|October 1, 2026' })), zoneOf(Object.assign({}, BASE, { MEETING_TIME_ZONE: 'Mars/Olympus' })), zoneOf({})],
    ['Asia/Shanghai', '', '', '', '']);
}

// ====================================================================== E
console.log('E. the upload clock');
{
  const loaded = sandbox({}, 'apps');
  const s = loaded.sandbox;
  s.Session.getScriptTimeZone = () => 'Europe/Berlin';
  check('UTC upload times into the meeting zone: Mountain Time in November (UTC-7), and over its DST change', [s.convertUtcWallClockToReportZone_('2026-11-16T16:30:00', 'America/Denver'), s.convertUtcWallClockToReportZone_('2026-11-01T07:30:00', 'America/Denver'), s.convertUtcWallClockToReportZone_('2026-11-01T08:30:00', 'America/Denver')],
    ['2026-11-16T09:30:00', '2026-11-01T01:30:00', '2026-11-01T01:30:00']);
  check('...and Calgary as its clocks are', s.convertUtcWallClockToReportZone_('2026-11-16T16:30:00', 'America/Edmonton'), '2026-11-16T10:30:00');
  check('without a meeting zone, or with anything that is not one: the script\'s zone, as before', [s.convertUtcWallClockToReportZone_('2026-09-22T13:00:00'), s.convertUtcWallClockToReportZone_('2026-09-22T13:00:00', ''), s.convertUtcWallClockToReportZone_('2026-09-22T13:00:00', 1), s.convertUtcWallClockToReportZone_('2026-09-22T13:00:00', 'UTC+2')],
    ['2026-09-22T15:00:00', '2026-09-22T15:00:00', '2026-09-22T15:00:00', '2026-09-22T15:00:00']);
  // The resolver of the Session column uses the meeting zone; a wall-clock "session" clock is never converted.
  const session = { id: 's1', label: '', date: '2026-11-16', start: '09:00', end: '12:00' };
  const props = { MEETING_TYPE: 'adhoc', MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Denver', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: '60780|November 16, 2026',
    ADHOC_SESSIONS: JSON.stringify({ v: 1, nextId: 2, sessions: [session] }) };
  const withClock = (clock) => { const l = sandbox(Object.assign({}, props, clock ? { ADHOC_TDOC_SESSIONS: JSON.stringify({ v: 1, clock: clock, add: {}, set: {} }) } : {}), 'apps'); l.sandbox.Session.getScriptTimeZone = () => 'Europe/Berlin'; return l.sandbox.makeAdhocTdocSessionResolver_(); };
  const td = (at) => ({ tdoc: 'S4-261001', uploaded: { at: at, dateOnly: null, unreadable: false } });
  const utc = withClock(null);
  check('the Session column: 18:59 UTC is 11:59 MST, in the session; 19:01 UTC is 12:01 MST, after it (in Berlin both would be 19:59/20:01, after it)', [utc.zone, utc.uploaded(td('2026-11-16T18:59:00')).at, utc.automatic(td('2026-11-16T18:59:00')).sessionId, utc.automatic(td('2026-11-16T19:01:00')).sessionId],
    ['America/Denver', '2026-11-16T11:59:00', 's1', null]);
  check('a list already on the sessions\' clock is not converted again', withClock('session').uploaded(td('2026-11-16T11:59:00')).at, '2026-11-16T11:59:00');
}

// ====================================================================== F
console.log('F. Configure Meeting');
{
  const SAVE = { meetingId: '60780', meetingType: 'main', meetingName: '3GPPSA4#138', meetingDate: 'November 16, 2026', reportType: 'MBS', meetingFolder: 'TSGS4_138_Calgary', meetingNumber: '138',
    ftpBase: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', mailingListMode: 'derived', apiTokenAction: 'keep', showPreview: true };
  const RESOLVED_138 = { meetingId: '60780', startDate: SA4_138.StartDate, timeZone: SA4_138.StartTimeZone, country: 'CA' };
  const persist = (props, changes) => {
    const l = sandbox(props, 'intl');
    l.sandbox.Session.getScriptTimeZone = () => 'Asia/Tokyo';
    const error = thrown(() => l.sandbox.persistConfigurationSettings_(Object.assign({}, SAVE, changes || {})));
    return { error: error, zone: ZONE_KEYS.map((k) => l.docProps.getProperty(k)), all: plain(l.docProps._store) };
  };
  // A record whose zone fits the rules of this runtime (see the note at the top); and SA4#138 as the Portal has it.
  const RESOLVED_OK = { meetingId: '60780', startDate: SA4_138_ALBERTA.StartDate, timeZone: SA4_138_ALBERTA.StartTimeZone, country: 'CA' };
  const asPortal = persist({}, { resolvedMeetingStart: RESOLVED_138 });
  check('SA4#138 (main) as the Portal has it: Resolve and Save store no zone -- its offset is out of date; nothing is guessed',
    [asPortal.error, asPortal.zone], [null, [null, null, null, null]]);
  check('...so SA4#138 selects its zone explicitly: America/Edmonton, no start time', persist({}, { resolvedMeetingStart: RESOLVED_138, meetingTimeZoneChoice: 'America/Edmonton' }).zone, [null, 'America/Edmonton', 'explicit', '60780|November 16, 2026']);
  check('a main meeting whose Portal zone fits: Resolve and Save store the meeting-local zone and no start time', persist({}, { resolvedMeetingStart: RESOLVED_OK }).zone, [null, 'America/Edmonton', 'local', '60780|November 16, 2026']);
  check('Meeting local chosen in the dialog with that Resolve: the same', persist({}, { resolvedMeetingStart: RESOLVED_OK, meetingTimeZoneChoice: 'local' }).zone, [null, 'America/Edmonton', 'local', '60780|November 16, 2026']);
  const before = { SA4_BOOTSTRAP_STATE: 'manual', MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026' };
  const refused = persist(before, { meetingTimeZoneChoice: 'local' });
  check('Meeting local without a known meeting-local zone: refused, and nothing at all written', [refused.error, plain(refused.all)],
    ['The meeting-local time zone is not known (Resolve the meeting first). Select a time zone instead.', before]);
  check('Meeting local when the Portal\'s zone does not fit: refused with the reason', persist(before, { meetingTimeZoneChoice: 'local', resolvedMeetingStart: Object.assign({}, RESOLVED_138, { timeZone: '(GMT-06:00) Mountain Time (US & Canada)', country: 'US' }) }).error,
    'The meeting-local time zone is not known: the Portal gives UTC-6 for the start, but America/Denver is UTC-7 (MST) then. Select a time zone instead.');
  check('a zone that is not in the list: refused, nothing written', [persist(before, { meetingTimeZoneChoice: 'Mars/Olympus' }).error, plain(persist(before, { meetingTimeZoneChoice: 'Mars/Olympus' }).all)], ['"Mars/Olympus" is not a time zone a meeting can be given.', before]);
  check('an explicit zone: stored as explicit', persist({}, { meetingTimeZoneChoice: 'America/Edmonton' }).zone, [null, 'America/Edmonton', 'explicit', '60780|November 16, 2026']);
  const explicit = { MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Denver', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: '60780|November 16, 2026' };
  check('a later Resolve and Save keeps an explicit zone; choosing Meeting local replaces it', [persist(explicit, { resolvedMeetingStart: RESOLVED_OK }).zone, persist(explicit, { resolvedMeetingStart: RESOLVED_OK, meetingTimeZoneChoice: 'local' }).zone],
    [[null, 'America/Denver', 'explicit', '60780|November 16, 2026'], [null, 'America/Edmonton', 'local', '60780|November 16, 2026']]);
  check('a stored Meeting local is kept by a Save without Resolve', persist({ MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'local', MEETING_START_BASIS: '60780|November 16, 2026' }, { meetingTimeZoneChoice: 'local' }).zone,
    [null, 'America/Edmonton', 'local', '60780|November 16, 2026']);
  // Ad-hoc 86172, with a legacy offset stored by an earlier release.
  const AD = { meetingId: '86172', meetingType: 'adhoc', meetingName: 'SA4-e (AH) MBS SWG post 137-e', meetingDate: 'October 1, 2026', reportType: 'MBS', meetingFolder: '', meetingNumber: '' };
  const LEGACY = { MEETING_ID: '86172', MEETING_TYPE: 'adhoc', MEETING_DATE: 'October 1, 2026', MEETING_START_TIME: '15:30', MEETING_TIME_ZONE: 'UTC+2', MEETING_START_BASIS: '86172|October 1, 2026' };
  const RESOLVED_86172 = { meetingId: '86172', startDate: '2026-10-01 15:30:00', timeZone: '(GMT+02.00)  Brussels, Copenhagen, Madrid, Paris', country: '' };
  check('a legacy "UTC+2" is kept by a Save, and by a Resolve that says the same as an IANA zone', [persist(LEGACY, AD).zone, persist(LEGACY, Object.assign({ resolvedMeetingStart: RESOLVED_86172 }, AD)).zone],
    [['15:30', 'UTC+2', null, '86172|October 1, 2026'], ['15:30', 'UTC+2', null, '86172|October 1, 2026']]);
  check('...a Portal that now gives another offset is taken, as a Resolve always did', persist(Object.assign({}, LEGACY, { MEETING_TIME_ZONE: 'UTC+1' }), Object.assign({ resolvedMeetingStart: RESOLVED_86172 }, AD)).zone,
    ['15:30', 'Europe/Brussels', 'local', '86172|October 1, 2026']);
  check('migration is a choice: Meeting local, or a zone', [persist(LEGACY, Object.assign({ resolvedMeetingStart: RESOLVED_86172, meetingTimeZoneChoice: 'local' }, AD)).zone, persist(LEGACY, Object.assign({ meetingTimeZoneChoice: 'Europe/Berlin' }, AD)).zone],
    [['15:30', 'Europe/Brussels', 'local', '86172|October 1, 2026'], ['15:30', 'Europe/Berlin', 'explicit', '86172|October 1, 2026']]);
  check('a zone chosen for a meeting whose stored start belongs to another one: the old start time is not revived', persist({ MEETING_ID: '86172', MEETING_DATE: 'October 1, 2026', MEETING_START_TIME: '15:30', MEETING_START_BASIS: '86999|October 1, 2026' }, Object.assign({ meetingTimeZoneChoice: 'Europe/Berlin' }, AD)).zone,
    [null, 'Europe/Berlin', 'explicit', '86172|October 1, 2026']);
  check('no meeting location: nothing is stored, nothing assumed', persist({}, { resolvedMeetingStart: Object.assign({}, RESOLVED_138, { timeZone: '' }) }).zone, [null, null, null, null]);
  // The creator: the Portal's record gives a new report its zone, verified.
  const creatorZone = (meetingType, m) => { const s = sandbox({}, 'intl').sandbox; return [meetingType === 'adhoc' ? s.meetingTimeZoneFromPortal_(m.StartDate, m.StartTimeZone, m.Country) : s.meetingLocalTimeZoneFromPortal_(m.StartDate, m.StartTimeZone, m.Country).zone]; };
  check('the creator\'s rule: a main meeting its verified zone only; an ad-hoc meeting else its offset', [creatorZone('main', SA4_138_ALBERTA), creatorZone('main', Object.assign({}, SA4_138, { StartTimeZone: '(GMT-07:00) Chihuahua, La Paz, Mazatlan' })), creatorZone('adhoc', Object.assign({}, SA4_138, { StartTimeZone: '(GMT-07:00) Chihuahua, La Paz, Mazatlan' }))],
    [['America/Edmonton'], [''], ['UTC-7']]);
  const s = sandbox({}, 'intl').sandbox;
  check('the creator\'s check: an offset or a zone of the list', ['America/Edmonton', 'UTC+2', 'Mars/Olympus', 'CEST'].map((zone) => s.isValidMeetingTimeZoneSetting_(zone)), [true, true, false, false]);
}

// ====================================================================== G
console.log('G. the dialog');
{
  const SA4_138_MEETING = [Object.assign({}, GM.getMeetings86172[0], SA4_138, { Id: 60780, Type: 'OR', MtgDocURL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/' })];
  const open = (props) => {
    const r = configurableReport({ props: props || {}, meetings: { 60780: SA4_138_MEETING } });
    r.s.Utilities.formatDate = appsScriptFormatDate;
    const page = r.open();
    const value = (id) => ((page.html.match(new RegExp('id="' + id + '" value="([^"]*)"')) || [])[1] || '');
    page.fields.meetingTimeZoneChoice = { id: 'meetingTimeZoneChoice', value: (page.html.match(/<option value="([^"]+)" selected="selected">/) || [])[1] || '' };
    page.fields.meetingTimeZoneStored = { value: value('meetingTimeZoneStored') };
    page.fields.meetingTimeZoneText = { value: value('meetingTimeZoneText').replace(/&quot;/g, '"').replace(/&amp;/g, '&') };
    page.fields.meetingTimeZoneInitial = { value: value('meetingTimeZoneInitial') };
    page.fields.meetingTimeZoneHint = { textContent: '' };
    return { r: r, page: page };
  };
  const fresh = open();
  check('the field: Not set, Meeting local, then the zones with their abbreviations; nothing stored is "Not set"', [/<select id="meetingTimeZoneChoice"/.test(fresh.page.html), /<option value="unset" selected="selected">Not set<\/option><option value="local">Meeting local<\/option>/.test(fresh.page.html),
    /<option value="America\/Denver">America\/Denver \(MST\/MDT\)<\/option>/.test(fresh.page.html), /<option value="America\/Edmonton">America\/Edmonton \(/.test(fresh.page.html), fresh.page.fields.meetingTimeZoneStored.value], [true, true, true, true, 'none']);
  fresh.page.resolve('60780');
  check('Resolve of SA4#138 as the Portal has it: Meeting local is not selected, the hint says why and asks for a zone', [fresh.page.fields.meetingTimeZoneChoice.value, fresh.page.fields.meetingTimeZoneHint.textContent],
    ['unset', 'Meeting local is not available: the Portal gives UTC-7 for the start, but America/Edmonton is UTC-6 (ABT) then. Select a time zone.']);
  fresh.page.fields.meetingTimeZoneChoice.value = 'America/Edmonton';
  fresh.page.fields.reportType.value = 'MBS';
  fresh.page.save();
  check('Save sends the choice and the country; the report has the zone', [fresh.page.sent[0].meetingTimeZoneChoice, fresh.page.sent[0].resolvedMeetingStart.country, fresh.r.props(ZONE_KEYS)],
    ['America/Edmonton', 'CA', [null, 'America/Edmonton', 'explicit', '60780|November 16, 2026']]);
  const verified = open();
  verified.r.s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify([Object.assign({}, SA4_138_MEETING[0], { StartTimeZone: SA4_138_ALBERTA.StartTimeZone })]) });
  verified.page.resolve('60780');
  check('a Portal zone that fits, with nothing stored: Resolve selects Meeting local', verified.page.fields.meetingTimeZoneChoice.value, 'local');
  const legacy = open({ MEETING_ID: '86172', MEETING_TYPE: 'adhoc', MEETING_DATE: 'October 1, 2026', MEETING_START_TIME: '15:30', MEETING_TIME_ZONE: 'UTC+2', MEETING_START_BASIS: '86172|October 1, 2026' });
  check('a legacy offset: shown as kept, "Not set" selected -- a Resolve does not change it by itself', [legacy.page.fields.meetingTimeZoneStored.value, legacy.page.fields.meetingTimeZoneChoice.value, /fixed offset UTC\+2 the Portal gave; it is kept/.test(legacy.page.fields.meetingTimeZoneText.value)], ['offset', 'unset', true]);
  const blocked = open({ MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026' });
  blocked.page.fields.meetingTimeZoneChoice.value = 'local';
  blocked.page.save();
  check('Meeting local not known in the page: Save says so and sends nothing', [blocked.page.sent.length, blocked.page.alerts], [0, ['Meeting local is not known yet: Resolve the meeting, or select a time zone.']]);
  const chosen = open({ MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: '60780|November 16, 2026' });
  const EDMONTON_LABEL = chosen.r.s.meetingTimeZoneOptionLabel_('America/Edmonton');
  check('an explicit zone is shown selected', [chosen.page.fields.meetingTimeZoneChoice.value, chosen.page.fields.meetingTimeZoneStored.value, chosen.page.fields.meetingTimeZoneText.value], ['America/Edmonton', 'explicit', 'Selected: ' + EDMONTON_LABEL + '.']);
  // Configure Sessions says which zone its times are in.
  const sessionsHint = (props) => { const r = configurableReport({ props: Object.assign({ MEETING_TYPE: 'adhoc' }, props) }); r.s.Utilities.formatDate = appsScriptFormatDate; r.s.configureAdhocSessions(); const html = r.dialogs[r.dialogs.length - 1].html; return (html.match(/Times are optional and are meant in ([^:]+):/) || [])[1]; };
  check('Configure Sessions: the times are meant in the meeting zone, else the report\'s zone as before', [sessionsHint({ MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'local', MEETING_START_BASIS: '60780|November 16, 2026' }), sessionsHint({})],
    ['the meeting time zone, ' + EDMONTON_LABEL, 'the report\'s time zone']);
}

// ====================================================================== H
console.log('H. manual edits of the opening survive an update');
{
  const t = (n, agenda) => ({ id: 'S4aA2691' + String(n).padStart(2, '0'), agenda: agenda, status: 'available', uploaded: true });
  const zoneProps = { MEETING_START_TIME: '15:00', MEETING_TIME_ZONE: 'Europe/Berlin', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: '85916|September 22, 2026' };
  const texts = (r) => r.body._children.filter((c) => c.getType() === 'PARAGRAPH').map((c) => c.getText());
  // Without sessions: the opening sentence.
  const r = templateReport({ tdocs: [t(1, '4.3')], props: zoneProps });
  r.s.Utilities.formatDate = appsScriptFormatDate;
  const b = r.build();
  check('the build writes the opening sentence with the zone of the meeting date', [b.ok, texts(r).filter((x) => /opens the session/.test(x))], [true, ['<Chair> opens the session on September 22, 2026 at 15:00 CEST.']]);
  const opening = r.body._children.filter((c) => c.getType() === 'PARAGRAPH' && /opens the session/.test(c.getText()))[0];
  r.user(() => opening.setText('Alex Organizer opened the session on September 22, 2026 at 15:05 CEST.'));
  const u = r.update();
  check('an update leaves the edited sentence as it is and writes no other', [u.success, texts(r).filter((x) => /opene?d? the session/.test(x))], [true, ['Alex Organizer opened the session on September 22, 2026 at 15:05 CEST.']]);
  // With sessions: the Session administration section, written by a build, edited by hand, an update.
  const sessions = JSON.stringify({ v: 1, nextId: 3, sessions: [{ id: 's1', label: '', date: '2026-10-23', endDate: '2026-10-26', start: '09:00', end: '17:00' }, { id: 's2', label: '', date: '2026-10-27', start: '15:00', end: '18:00' }] });
  const w = templateReport({ tdocs: [t(1, '4.3')], props: Object.assign({}, zoneProps, { ADHOC_SESSIONS: sessions }) });
  w.s.Utilities.formatDate = appsScriptFormatDate;
  w.build();
  check('a build of a multi-day series across the DST change: each session time with the zone of its day', texts(w).filter((x) => /^A0\d:/.test(x)),
    ['A01: October 23–26, 2026 (from 09:00 CEST on the first day, until 17:00 CET on the last day)', 'A02: October 27, 2026, 15:00–18:00 CET']);
  const line = w.body._children.filter((c) => c.getType() === 'PARAGRAPH' && /^A02:/.test(c.getText()))[0];
  w.user(() => line.setText('A02: October 27, 2026, 15:10–18:20 CET (started late)'));
  const before = texts(w);
  const wu = w.update();
  check('an update leaves the edited Session administration section as it is', [wu.success, texts(w).filter((x) => /^A0\d:/.test(x)), JSON.stringify(texts(w)) === JSON.stringify(before)],
    [true, ['A01: October 23–26, 2026 (from 09:00 CEST on the first day, until 17:00 CET on the last day)', 'A02: October 27, 2026, 15:10–18:20 CET (started late)'], true]);
  const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const updateSource = CODE.slice(CODE.indexOf('\nfunction continuousUpdateCore_('), CODE.indexOf('\n}\n', CODE.indexOf('\nfunction continuousUpdateCore_(')));
  check('the update writes no opening: it calls neither the opening sentence nor the Session administration section (source)', /buildMeetingOpeningSentence_|refreshAdhocOpeningSection_|finishAdhocOpeningRebuild_|renderAdhocOpeningSection_/.test(updateSource), false);
}

// ====================================================================== I
console.log('I. "CEST" in the code');
{
  const FULL_CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  // The zone table of the section names CEST as the abbreviation of +02:00; everything else is looked at.
  const a = FULL_CODE.indexOf(MEETING_TIME_ZONES_SECTION_START);
  const b = FULL_CODE.indexOf(MEETING_TIME_ZONES_SECTION_END);
  const CODE = FULL_CODE.slice(0, a) + FULL_CODE.slice(b);
  const functionAt = (index) => { const start = CODE.lastIndexOf('\nfunction ', index); return start === -1 ? '(top level)' : CODE.slice(start + 10, CODE.indexOf('(', start)); };
  const users = [];
  const re = /CEST/g;
  let m;
  while ((m = re.exec(CODE)) !== null) {
    const lineStart = CODE.lastIndexOf('\n', m.index) + 1;
    const line = CODE.slice(lineStart, CODE.indexOf('\n', m.index));
    if (/^\s*(\*|\/\/)/.test(line)) continue;
    const name = functionAt(m.index);
    if (users.indexOf(name) === -1) users.push(name);
  }
  check('"CEST" outside comments (and outside the zone table) only where deadlines quoted in e-mails are read and discussion e-mails are written -- not meeting times', users.sort(),
    ['cleanUpWrongEmailDiscussions', 'extractDeadlineFromTitle_', 'isEmailExportReserved_', 'parseEmailSubject_', 'parseExtendedDeadline_']);
  const sectionText = FULL_CODE.slice(a, b);
  check('...the meeting times themselves never name one: the opening sentence, the sessions, the clock', ['buildMeetingOpeningSentence_', 'adhocSessionWhenText_', 'convertUtcWallClockToReportZone_', 'meetingTimeZoneDisplay_'].map((n) => {
    const at = CODE.indexOf('\nfunction ' + n + '(');
    return /CEST|CET\b|'MST'/.test(CODE.slice(at, CODE.indexOf('\n}\n', at)).replace(/\/\/.*$/gm, ''));
  }), [false, false, false, false]);
  check('the e-mail collection zone is not the meeting zone: TIMEZONE is untouched by the section', /TIMEZONE\b(?!_)|getCollectorConfig_/.test(sectionText.replace(/MEETING_TIME_ZONE/g, '')), false);
}

// ====================================================================== J
console.log('J. what changed, against T-2026.10.23');
{
  const RAW = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  // The main-meeting number fix (after T-2026.10.24, tests/main-meeting-number.test.js) came later: this ledger is about the layer below it.
  const FULL = codeWithoutMainMeetingNumber(RAW).code;
  const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutMeetingTimeZones(RAW);
  const creatorView = creatorWithoutMeetingTimeZones(CREATOR);
  check('every change is in Code.js and ReportCreator.js exactly once', view.problems.concat(creatorView.problems), []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((x) => x.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(MEETING_TIME_ZONES_SECTION_START), FULL.indexOf(MEETING_TIME_ZONES_SECTION_END) + MEETING_TIME_ZONES_SECTION_END.length);
  check('the section: these functions, in this order', functionNames(section), MEETING_TIME_ZONES_SECTION_FUNCTIONS);
  check('it stands directly after getMeetingStartForOpening_()', FULL.indexOf('    return none;\n  }\n}\n\n' + MEETING_TIME_ZONES_SECTION_START) !== -1, true);
  check('Code.js is version 2.26.8 (released as T-2026.10.24; the ledger takes the header out)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.8');
  let released = null;
  let releasedCreator = null;
  try {
    released = execFileSync('git', ['show', 'template-release/T-2026.10.23:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
    releasedCreator = execFileSync('git', ['show', 'template-release/T-2026.10.23:template/ReportCreator.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    released = null;
  }
  if (released === null) {
    console.log('  note: template-release/T-2026.10.23 is not available in this checkout; the comparison with it is skipped.');
  } else {
    check('without it, Code.js and ReportCreator.js are those of T-2026.10.23, byte for byte', [view.code === released, creatorView.code === releasedCreator], [true, true]);
    check('the functions of Code.js that changed', functionNames(released).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(FULL, n) !== functionSource(released, n)).sort(),
      ['adhocSessionWhenText_', 'buildAdhocOpeningDialogModel_', 'buildAdhocOpeningLines_', 'buildAdhocSessionStatusModel_', 'buildAdhocTdocSessionsDialogModel_', 'buildMeetingOpeningSentence_', 'computeResolvedMeetingPreview_',
        'configureAdhocSessions', 'configureMeetingSettings', 'convertUtcWallClockToReportZone_', 'getMeetingStartForOpening_', 'makeAdhocTdocSessionResolver_', 'meetingStartFromResolved_', 'persistConfigurationSettings_',
        'refreshAdhocOpeningSection_', 'resolveMeetingCoreById_', 'storeMeetingStart_']);
    check('the new functions are those of the section; none removed', [functionNames(FULL).filter((n) => functionNames(released).indexOf(n) === -1).sort(), functionNames(released).filter((n) => functionNames(FULL).indexOf(n) === -1)],
      [MEETING_TIME_ZONES_SECTION_FUNCTIONS.slice().sort(), []]);
    check('the functions of ReportCreator.js that changed: the bootstrap config and its check', functionNames(releasedCreator).filter((n) => functionSource(CREATOR, n) !== functionSource(releasedCreator, n)).sort(),
      ['buildReportBootstrapPayload_', 'validateBootstrapPayload_']);
    check('untouched: the update, Close Report, the e-mail collector and its deadline parsing, the agenda, the TDoc list',
      ['continuousUpdateCore_', 'closeReportCore_', 'getCollectorConfig_', 'checkRSSFeed_', 'extractDeadlineFromSubject_', 'parseDeadlineCellText_', 'readAgendaWithFallbacks_', 'downloadAndGroupTdocs_', 'computeMeetingTimeZoneLabel_', 'isValidMeetingTimeZoneLabel_']
        .map((n) => functionSource(FULL, n) === functionSource(released, n)), Array(10).fill(true));
  }
}

// ====================================================================== K
console.log('\nK. regressions: Alberta (ABT) with every tz database, Calgary on 16 November 2026, Berlin, DST changes, CENTRAL');
{
  // Utilities.formatDate as Apps Script answers it with two generations of the tz database. "old": from before
  // Alberta's change -- Calgary goes back to UTC-7 (MST) on 1 November 2026, as Denver does. "new": tzdata 2026c
  // and later -- Calgary keeps UTC-6 and calls it MDT. Every other zone by the runtime's own rules.
  const ALBERTA_CHANGE = Date.UTC(2026, 10, 1, 8, 0, 0);
  const tzdata = (generation) => (date, zone, pattern) => {
    if (zone === 'America/Edmonton' && pattern === 'Z') {
      if (generation === 'new' && date.getTime() >= ALBERTA_CHANGE) return '-0600';
      return appsScriptFormatDate(date, 'America/Denver', 'Z');
    }
    return appsScriptFormatDate(date, zone, pattern);
  };
  const withRules = (generation, props) => {
    const loaded = loadCode({ documentProperties: props || {} });
    if (generation !== 'intl') loaded.sandbox.Utilities.formatDate = tzdata(generation);
    loaded.sandbox.Session.getScriptTimeZone = () => 'Europe/Berlin';
    return loaded;
  };
  const S = (date, endDate, start, end, label) => Object.assign({ id: 's1', label: label || '', date: date, start: start || '', end: end || '' }, endDate ? { endDate: endDate } : {});

  ['old', 'new', 'intl'].forEach((generation) => {
    const s = withRules(generation).sandbox;
    const tag = 'tzdata ' + generation + ': ';
    // Calgary, 16 November 2026 (SA4#138).
    check(tag + 'Calgary on 16 November 2026 is UTC-6 and written ABT: offset, abbreviation, 09:00 local = 15:00 UTC',
      [s.meetingTimeZoneOffsetMinutes_('America/Edmonton', Date.UTC(2026, 10, 16, 15)), s.meetingTimeZoneDisplay_('America/Edmonton', '2026-11-16', '09:00'),
        new Date(s.meetingWallClockToInstant_('America/Edmonton', '2026-11-16', '09:00')).toISOString()],
      [-360, 'ABT', '2026-11-16T15:00:00.000Z']);
    check(tag + 'SA4#138: the opening sentence, the plenary of 16-20 November, one day', [
      s.buildMeetingOpeningSentence_('November 16, 2026', '09:00', 'America/Edmonton'),
      s.adhocSessionWhenText_(S('2026-11-16', '2026-11-20', '09:00', '17:30', 'Plenary'), 'America/Edmonton'),
      s.adhocSessionWhenText_(S('2026-11-16', '', '09:00', '17:30'), 'America/Edmonton')],
    ['<Chair> opens the session on November 16, 2026 at 09:00 ABT.',
      'Plenary, November 16–20, 2026 (from 09:00 ABT on the first day, until 17:30 ABT on the last day)',
      'November 16, 2026, 09:00–17:30 ABT']);
    check(tag + 'SA4#138: UTC upload times converted once into Calgary time (UTC-6)', [
      s.convertUtcWallClockToReportZone_('2026-11-16T14:59:00', 'America/Edmonton'), s.convertUtcWallClockToReportZone_('2026-11-16T15:00:00', 'America/Edmonton'),
      s.convertUtcWallClockToReportZone_('2026-11-20T23:30:00', 'America/Edmonton')],
    ['2026-11-16T08:59:00', '2026-11-16T09:00:00', '2026-11-20T17:30:00']);
    // The night of 1 November 2026: MDT until 08:00 UTC, ABT from then on, the same UTC-6 throughout -- no hour twice.
    check(tag + '1 November 2026: 07:59 UTC is MDT, 08:00 UTC is ABT, both UTC-6', [
      s.meetingTimeZoneAbbreviationAt_('America/Edmonton', ALBERTA_CHANGE - 60000), s.meetingTimeZoneAbbreviationAt_('America/Edmonton', ALBERTA_CHANGE),
      s.meetingTimeZoneOffsetMinutes_('America/Edmonton', ALBERTA_CHANGE - 60000), s.meetingTimeZoneOffsetMinutes_('America/Edmonton', ALBERTA_CHANGE)],
    ['MDT', 'ABT', -360, -360]);
    check(tag + '1 November 2026, local times: 01:30 MDT, 02:30 ABT (no repeated hour), each at UTC-6', [
      new Date(s.meetingWallClockToInstant_('America/Edmonton', '2026-11-01', '01:30')).toISOString(), new Date(s.meetingWallClockToInstant_('America/Edmonton', '2026-11-01', '02:30')).toISOString(),
      s.meetingTimeZoneDisplay_('America/Edmonton', '2026-11-01', '01:30'), s.meetingTimeZoneDisplay_('America/Edmonton', '2026-11-01', '02:30'),
      s.convertUtcWallClockToReportZone_('2026-11-01T07:30:00', 'America/Edmonton'), s.convertUtcWallClockToReportZone_('2026-11-01T08:30:00', 'America/Edmonton')],
    ['2026-11-01T07:30:00.000Z', '2026-11-01T08:30:00.000Z', 'MDT', 'ABT', '2026-11-01T01:30:00', '2026-11-01T02:30:00']);
    check(tag + 'a session over the change: from MDT on the first day, until ABT on the last day', s.adhocSessionWhenText_(S('2026-10-30', '2026-11-02', '09:00', '17:00'), 'America/Edmonton'),
      'October 30 – November 2, 2026 (from 09:00 MDT on the first day, until 17:00 ABT on the last day)');
    check(tag + 'Calgary before the change keeps its old rules: MST in January 2026, MDT from 8 March 2026; ABT in summer 2027 (no DST any more)', [
      s.meetingTimeZoneDisplay_('America/Edmonton', '2026-01-15', '09:00'), s.meetingTimeZoneDisplay_('America/Edmonton', '2026-03-07', '12:00'),
      s.meetingTimeZoneDisplay_('America/Edmonton', '2026-03-08', '12:00'), s.meetingTimeZoneDisplay_('America/Edmonton', '2027-07-15', '09:00'),
      s.meetingTimeZoneOffsetMinutes_('America/Edmonton', Date.UTC(2027, 6, 15, 15))],
    ['MST', 'MST', 'MDT', 'ABT', -360]);
    check(tag + 'Denver is not Alberta: MST from 1 November 2026, MDT in summer 2027', [s.meetingTimeZoneDisplay_('America/Denver', '2026-11-16', '09:00'), s.meetingTimeZoneDisplay_('America/Denver', '2027-07-15', '09:00')], ['MST', 'MDT']);
    // Berlin, summer and winter, and both DST changes of 2026.
    check(tag + 'Berlin in summer (CEST, UTC+2) and in winter (CET, UTC+1): sentence and uploads', [
      s.buildMeetingOpeningSentence_('August 24, 2026', '09:00', 'Europe/Berlin'), s.buildMeetingOpeningSentence_('January 20, 2026', '10:00', 'Europe/Berlin'),
      s.convertUtcWallClockToReportZone_('2026-08-24T07:30:00', 'Europe/Berlin'), s.convertUtcWallClockToReportZone_('2026-01-20T07:30:00', 'Europe/Berlin')],
    ['<Chair> opens the session on August 24, 2026 at 09:00 CEST.', '<Chair> opens the session on January 20, 2026 at 10:00 CET.', '2026-08-24T09:30:00', '2026-01-20T08:30:00']);
    check(tag + 'Berlin, 29 March 2026: 00:59 UTC is 01:59 CET, 01:00 UTC is 03:00 CEST', [
      s.convertUtcWallClockToReportZone_('2026-03-29T00:59:00', 'Europe/Berlin'), s.convertUtcWallClockToReportZone_('2026-03-29T01:00:00', 'Europe/Berlin'),
      s.meetingTimeZoneDisplay_('Europe/Berlin', '2026-03-29', '01:59'), s.meetingTimeZoneDisplay_('Europe/Berlin', '2026-03-29', '03:00')],
    ['2026-03-29T01:59:00', '2026-03-29T03:00:00', 'CET', 'CEST']);
    check(tag + 'Berlin, 25 October 2026: 00:30 UTC is 02:30 CEST, 01:30 UTC is 02:30 CET -- each UTC time converted once', [
      s.convertUtcWallClockToReportZone_('2026-10-25T00:30:00', 'Europe/Berlin'), s.convertUtcWallClockToReportZone_('2026-10-25T01:30:00', 'Europe/Berlin')],
    ['2026-10-25T02:30:00', '2026-10-25T02:30:00']);
    check(tag + 'legacy offsets are written as stored, also for Calgary dates', [s.buildMeetingOpeningSentence_('November 16, 2026', '09:00', 'UTC+2'), s.buildMeetingOpeningSentence_('November 16, 2026', '09:00', 'UTC-7'),
      s.adhocSessionWhenText_(S('2026-11-16', '', '09:00', '17:30'), 'UTC+2')],
    ['<Chair> opens the session on November 16, 2026 at 09:00 UTC+2.', '<Chair> opens the session on November 16, 2026 at 09:00 UTC-7.', 'November 16, 2026, 09:00–17:30 UTC+2']);
  });

  const s = withRules('old').sandbox;
  check('Configure Meeting names: by the coming twelve months -- Calgary from November 2026 ABT; across the change all three; Denver MST/MDT', [
    s.meetingTimeZoneOptionLabelAt_('America/Edmonton', Date.UTC(2026, 10, 12)), s.meetingTimeZoneOptionLabelAt_('America/Edmonton', Date.UTC(2026, 0, 15)),
    s.meetingTimeZoneOptionLabelAt_('America/Denver', Date.UTC(2026, 10, 12)), s.meetingTimeZoneOptionLabelAt_('Europe/Berlin', Date.UTC(2026, 10, 12)), s.meetingTimeZoneOptionLabelAt_('UTC', Date.UTC(2026, 10, 12))],
  ['America/Edmonton (ABT)', 'America/Edmonton (MST/MDT/ABT)', 'America/Denver (MST/MDT)', 'Europe/Berlin (CET/CEST)', 'UTC']);
  check('Meeting local for SA4#138 as the Portal has it ("(GMT-07:00) Mountain Time (US & Canada)", CA): refused with old and new tz data alike',
    ['old', 'new'].map((g) => withRules(g).sandbox.meetingLocalTimeZoneFromPortal_(SA4_138.StartDate, SA4_138.StartTimeZone, 'CA').reason),
    ['the Portal gives UTC-7 for the start, but America/Edmonton is UTC-6 (ABT) then', 'the Portal gives UTC-7 for the start, but America/Edmonton is UTC-6 (ABT) then']);
  check('Meeting local with Windows\' "(UTC-06:00) Mountain Time (Alberta)": America/Edmonton, verified, with old and new tz data alike',
    ['old', 'new'].map((g) => withRules(g).sandbox.meetingLocalTimeZoneFromPortal_(SA4_138.StartDate, '(UTC-06:00) Mountain Time (Alberta)', 'CA').zone), ['America/Edmonton', 'America/Edmonton']);

  // The Session column of SA4#138 with the old tz data: the end of a session (12:00) is the cut-off; UTC-7 would move it by an hour.
  const session = { id: 's1', label: '', date: '2026-11-16', start: '09:00', end: '12:00' };
  const CALGARY_PROPS = { MEETING_TYPE: 'adhoc', MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'explicit',
    MEETING_START_BASIS: '60780|November 16, 2026', ADHOC_SESSIONS: JSON.stringify({ v: 1, nextId: 2, sessions: [session] }) };
  const td = (at) => ({ tdoc: 'S4-261001', uploaded: { at: at, dateOnly: null, unreadable: false } });
  const bound = withRules('old', CALGARY_PROPS).sandbox.makeAdhocTdocSessionResolver_();
  check('Session column, bound script, old tz data: 17:59 UTC is 11:59 ABT, in the session (until 12:00); 18:01 UTC is 12:01 ABT, after it (at UTC-7 it would be 11:01, wrongly inside)', [bound.zone, bound.uploaded(td('2026-11-16T17:59:00')).at, bound.automatic(td('2026-11-16T17:59:00')).sessionId, bound.automatic(td('2026-11-16T18:01:00')).sessionId],
    ['America/Edmonton', '2026-11-16T11:59:00', 's1', null]);

  // CENTRAL add-on: adoption copies the sessions' zone, so a background update converts uploads as the bound script does.
  const central = withRules('old', CALGARY_PROPS);
  central.sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => 'DOC138' });
  check('CENTRAL: the zone, its mode and its meeting are adoption keys (MEETING_START_TIME, read only by a build, is not)', ['MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS', 'MEETING_START_TIME'].map((k) => central.sandbox.ADDON003_ADOPTION_FIXED_KEYS_.indexOf(k) !== -1), [true, true, true, false]);
  const adopted = central.sandbox.adoptReportDocumentForAddon_();
  check('CENTRAL: adoption copies them and verifies', [adopted.verified, ['MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'].map((k) => central.scriptProps.getProperty('SA4_STATE|DOC138|' + k))],
    [true, ['America/Edmonton', 'explicit', '60780|November 16, 2026']]);
  const background = central.sandbox.makeAdhocTdocSessionResolver_({ documentId: 'DOC138', mode: 'addon-background' });
  check('CENTRAL background update: the sessions\' zone is the meeting\'s, the upload is in the session as in the bound script (not Berlin, the add-on\'s own zone)',
    [central.sandbox.meetingSessionTimeZone_({ documentId: 'DOC138', mode: 'addon-background' }), background.uploaded(td('2026-11-16T15:01:00')).at, background.automatic(td('2026-11-16T15:01:00')).sessionId],
    ['America/Edmonton', '2026-11-16T09:01:00', 's1']);
  const notAdopted = withRules('old', CALGARY_PROPS);
  check('CENTRAL background run of a document adopted before this change (no zone in the central state): the add-on\'s zone, as before',
    notAdopted.sandbox.meetingSessionTimeZone_({ documentId: 'OTHER', mode: 'addon-background' }), '');
}

// ====================================================================== L
console.log('\nL. the review of the save path: unchanged choices, "Not set", another meeting, another clock, CENTRAL sync');
{
  const SAVE = { meetingId: '60780', meetingType: 'main', meetingName: '3GPPSA4#138', meetingDate: 'November 16, 2026', reportType: 'MBS', meetingFolder: 'TSGS4_138_Calgary', meetingNumber: '138',
    ftpBase: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', mailingListMode: 'derived', apiTokenAction: 'keep', showPreview: true };
  const persist = (props, changes) => {
    const l = sandbox(props, 'intl');
    l.sandbox.Session.getScriptTimeZone = () => 'Asia/Tokyo';
    const error = thrown(() => l.sandbox.persistConfigurationSettings_(Object.assign({}, SAVE, changes || {})));
    return { error: error, zone: ZONE_KEYS.map((k) => l.docProps.getProperty(k)), l: l };
  };
  const B138 = '60780|November 16, 2026';
  const RESOLVED_138 = { meetingId: '60780', startDate: SA4_138.StartDate, timeZone: SA4_138.StartTimeZone, country: 'CA' };
  const RESOLVED_ALBERTA = { meetingId: '60780', startDate: SA4_138.StartDate, timeZone: '(GMT-06:00) Mountain Time (Alberta)', country: 'CA' };
  const EXPLICIT = { MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'Europe/Paris', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: B138 };
  const LOCAL = { MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'local', MEETING_START_BASIS: B138 };

  // "Not set" removes a selected or meeting-local zone; a start time stays with its meeting.
  check('"Not set" removes an explicit zone', persist(EXPLICIT, { meetingTimeZoneChoice: 'unset' }).zone, [null, null, null, null]);
  check('"Not set" removes a meeting-local zone, keeps the start time and its meeting', persist(Object.assign({ MEETING_START_TIME: '09:00' }, LOCAL), { meetingTimeZoneChoice: 'unset' }).zone, ['09:00', null, null, B138]);
  check('no choice sent (the select left as it opened): the explicit zone of this meeting stays', persist(EXPLICIT, {}).zone, [null, 'Europe/Paris', 'explicit', B138]);

  // Another meeting: an explicit zone does not follow the report.
  const moved = persist(EXPLICIT, { meetingId: '60779', meetingDate: 'November 9, 2026' });
  check('another meeting saved without a choice: the explicit zone of the old meeting is not the new meeting\'s (its zone is not read)',
    [moved.error, moved.l.sandbox.meetingSessionTimeZone_(), moved.l.sandbox.meetingTimeZoneDialogState_(moved.l.docProps).stored], [null, '', 'none']);
  const mainStale = persist({ MEETING_ID: '60779', MEETING_DATE: 'November 9, 2026', MEETING_TIME_ZONE: 'Europe/Paris', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: '60779|November 9, 2026' },
    { resolvedMeetingStart: RESOLVED_ALBERTA });
  check('a main report moved to a meeting whose zone verifies: Meeting local is stored, a stale explicit mode of the old meeting does not block it', mainStale.zone, [null, 'America/Edmonton', 'local', B138]);
  check('...an explicit zone for this same meeting still wins over Meeting local', persist({ MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_TIME_ZONE: 'America/Denver', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: B138 },
    { resolvedMeetingStart: RESOLVED_ALBERTA }).zone, [null, 'America/Denver', 'explicit', B138]);

  // Meeting local asked for: a Resolve of this save that does not verify is not overruled by a stored zone.
  check('Meeting local with a Resolve that does not verify: refused, although a meeting-local zone is stored', persist(LOCAL, { meetingTimeZoneChoice: 'local', resolvedMeetingStart: RESOLVED_138 }).error,
    'The meeting-local time zone is not known: the Portal gives UTC-7 for the start, but America/Edmonton is UTC-6 (ABT) then. Select a time zone instead.');
  check('Meeting local without a Resolve: the stored zone of this meeting and date counts; of another date it does not', [
    persist(LOCAL, { meetingTimeZoneChoice: 'local' }).error, persist(LOCAL, { meetingTimeZoneChoice: 'local', meetingDate: 'November 17, 2026' }).error],
  [null, 'The meeting-local time zone is not known (Resolve the meeting first). Select a time zone instead.']);

  // An explicit zone and the start time of an ad-hoc meeting: the time stays only on its own clock.
  const ADHOC = { meetingType: 'adhoc', meetingId: '60780', meetingDate: 'November 16, 2026' };
  const RESOLVED_DENVER = { meetingId: '60780', startDate: '2026-11-16 09:00:00', timeZone: '(GMT-07:00) Mountain Time (US & Canada)', country: 'US' };
  check('ad-hoc, Portal 09:00 Mountain Time (US), Europe/Berlin selected: the start time is removed (09:00 CET would be 8 hours off)',
    persist({}, Object.assign({}, ADHOC, { resolvedMeetingStart: RESOLVED_DENVER, meetingTimeZoneChoice: 'Europe/Berlin' })).zone, [null, 'Europe/Berlin', 'explicit', B138]);
  check('ad-hoc, the same, America/Denver selected: the start time stays (the same clock)',
    persist({}, Object.assign({}, ADHOC, { resolvedMeetingStart: RESOLVED_DENVER, meetingTimeZoneChoice: 'America/Denver' })).zone, ['09:00', 'America/Denver', 'explicit', B138]);
  check('ad-hoc SA4#138 as the Portal has it (GMT-07:00, CA), America/Edmonton selected: the start time stays -- the Portal\'s name is Edmonton\'s, its offset is out of date',
    persist({}, Object.assign({}, ADHOC, { resolvedMeetingStart: RESOLVED_138, meetingTimeZoneChoice: 'America/Edmonton' })).zone, ['09:00', 'America/Edmonton', 'explicit', B138]);
  check('a start time stored without a zone stays when a zone is selected', persist({ MEETING_ID: '60780', MEETING_DATE: 'November 16, 2026', MEETING_START_TIME: '09:00', MEETING_START_BASIS: B138 },
    { meetingTimeZoneChoice: 'Europe/Berlin' }).zone, ['09:00', 'Europe/Berlin', 'explicit', B138]);

  // CENTRAL: a save of an adopted document syncs the zone into its central state.
  const central = persist({}, { resolvedMeetingStart: RESOLVED_138, meetingTimeZoneChoice: 'America/Edmonton' }).l;
  check('CENTRAL: the zone keys are Configure Meeting keys too (synced on every save of an adopted document)',
    ['MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'].map((k) => central.sandbox.CONFIG_DIALOG_MANAGED_KEYS_.indexOf(k) !== -1), [true, true, true]);
  check('CENTRAL: every Configure Meeting key is an adoption key', central.sandbox.CONFIG_DIALOG_MANAGED_KEYS_.filter((k) => central.sandbox.ADDON003_ADOPTION_FIXED_KEYS_.indexOf(k) === -1), []);
  central.sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => 'DOC138', getName: () => 'Audio SWG Minutes SA4#138' });
  const synced = thrown(() => central.sandbox.syncConfigurationDialogStateToCentral_('DOC138'));
  check('CENTRAL: the save is synced into the central state, and the background update reads the zone from there',
    [synced, ['MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'].map((k) => central.scriptProps.getProperty('SA4_STATE|DOC138|' + k)),
      central.sandbox.meetingSessionTimeZone_({ documentId: 'DOC138', mode: 'addon-background' })],
    [null, ['America/Edmonton', 'explicit', B138], 'America/Edmonton']);

  // The dialog: a choice is sent only when it was changed there.
  const SA4_138_MEETING = [Object.assign({}, GM.getMeetings86172[0], SA4_138, { Id: 60780, Type: 'OR', MtgDocURL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/' })];
  const open = (props) => {
    const r = configurableReport({ props: props || {}, meetings: { 60780: SA4_138_MEETING } });
    r.s.Utilities.formatDate = appsScriptFormatDate;
    const page = r.open();
    const value = (id) => ((page.html.match(new RegExp('id="' + id + '" value="([^"]*)"')) || [])[1] || '');
    page.fields.meetingTimeZoneChoice = { id: 'meetingTimeZoneChoice', value: (page.html.match(/<option value="([^"]+)" selected="selected">/) || [])[1] || '' };
    page.fields.meetingTimeZoneStored = { value: value('meetingTimeZoneStored') };
    page.fields.meetingTimeZoneText = { value: value('meetingTimeZoneText') };
    page.fields.meetingTimeZoneInitial = { value: value('meetingTimeZoneInitial') };
    page.fields.meetingTimeZoneHint = { textContent: '' };
    page.fields.reportType.value = 'MBS';
    return { r: r, page: page };
  };
  const untouched = open(EXPLICIT);
  check('the dialog opens with what is selected: Europe/Paris, remembered as the initial choice', [untouched.page.fields.meetingTimeZoneChoice.value, untouched.page.fields.meetingTimeZoneInitial.value], ['Europe/Paris', 'Europe/Paris']);
  untouched.page.save();
  check('Save with the select untouched sends no choice; the zone stays', [untouched.page.sent.length, 'meetingTimeZoneChoice' in untouched.page.sent[0], untouched.r.props(ZONE_KEYS)], [1, false, [null, 'Europe/Paris', 'explicit', B138]]);
  const cleared = open(EXPLICIT);
  cleared.page.fields.meetingTimeZoneChoice.value = 'unset';
  cleared.page.save();
  check('"Not set" chosen in the dialog: sent as unset, the zone is removed', [cleared.page.sent[0].meetingTimeZoneChoice, cleared.r.props(ZONE_KEYS)], ['unset', [null, null, null, null]]);
  const nothing = open();
  nothing.page.save();
  check('nothing stored, "Not set" left as it is: no choice sent', [nothing.page.fields.meetingTimeZoneInitial.value, 'meetingTimeZoneChoice' in nothing.page.sent[0]], ['unset', false]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll meeting time zone checks passed.');
process.exit(failures ? 1 : 0);
