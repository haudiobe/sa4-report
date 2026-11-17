/**
 * The scheduled overnight pause of the local listing agent, told apart from a genuine synchronization
 * failure (after T-2026.10.28; docs/LOCAL_LISTING_AGENT.md, "Operating window and overnight staleness").
 *
 * The agent scans only from 07:30 to 19:30 meeting local time. A stale listing is the scheduled pause
 * (result "paused") only when it is night (or the first 15 minutes of the morning) in the meeting's
 * zone AND the agent's last scan was in the last 30 minutes before 19:30. Otherwise it stays "stale".
 *
 *   1. localListingScheduledPause_(): boundaries; ABT (America/Edmonton, UTC-6 all year from
 *      2026-11-01) and the night of that change; a zone with DST (Europe/Berlin); no zone
 *   1b. the same with OUTDATED runtime zone rules (America/Edmonton still UTC-7 in winter): Alberta's
 *      rule (MEETING_TIME_ZONE_RULE_CHANGES_) decides, and a control shows the old rules would not
 *   2. importAgentListing_() on a fake Drive: paused vs stale, nothing imported or erased either way;
 *      the morning; no meeting zone configured; Check Connections
 *   3. what changed in Code.js (source)
 *
 * All values are synthetic. Run: node tests/local-listing-overnight-pause.test.js
 */
const fs = require('fs');
const crypto = require('crypto');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { codeWithoutOvernightPause, OVERNIGHT_PAUSE_NEW_FUNCTIONS, OVERNIGHT_PAUSE_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`); else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const LOCAL = 'http://10.10.10.10/ftp/SA/SA4/Inbox';
const A = 'S4-269101';
const sha = (t) => crypto.createHash('sha256').update(String(t), 'utf8').digest('hex');
const EDMONTON = 'America/Edmonton';

function listing(files, fetchedAt) {
  const folder = LOCAL + '/Drafts/MBS/';
  const payload = files.map((f) => '<a href="' + folder + f + '">' + f + '</a>').join('\n');
  return JSON.stringify({ format: 'sa4-local-listing', version: 1, group: 'MBS', folder: folder, fetchedAt: fetchedAt, scanId: sha(payload), agent: 'ROOM', linkCount: files.length, sha256: sha(payload), payload: payload });
}

/** A MBS report with the agent folder `drive`; `zone` stored as the meeting's zone for its meeting (or none). */
function report(drive, zone) {
  const props = { MEETING_TYPE: 'main', REPORT_SUFFIX: 'MBS', MEETING_FOLDER: 'TSGS4_138_Calgary', MEETING_NUMBER: '138', MEETING_ID: '86178', MEETING_DATE: '2026-11-16', LOCAL_LISTING_DRIVE_FOLDER_ID: 'DriveFolder0001' };
  const l = loadCode({ documentProperties: props });
  const s = l.sandbox;
  if (zone) {
    l.docProps.setProperty('MEETING_TIME_ZONE', zone);
    l.docProps.setProperty('MEETING_START_BASIS', s.meetingStartBasis_('86178', '2026-11-16'));
  }
  s.Utilities.computeDigest = (alg, text) => Array.from(crypto.createHash('sha256').update(String(text), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b));
  s.Utilities.DigestAlgorithm = { SHA_256: 'SHA_256' }; s.Utilities.Charset = { UTF_8: 'UTF_8' };
  const r = { s, props: l.docProps, drive: drive || {} };
  s.DriveApp = { getFolderById: () => ({ getFilesByName: (n) => { const raw = r.drive[n]; let done = raw === undefined; return { hasNext: () => !done, next: () => { done = true; return { getBlob: () => ({ getDataAsString: () => raw }), getLastUpdated: () => new Date('2026-11-17T01:29:00Z') }; } }; } }) };
  r.importAt = (iso) => s.importAgentListing_(undefined, { now: new Date(iso), byTdoc: {} });
  r.status = (iso) => s.describeLocalListingAgentStatus_(r.props.getProperty('LOCAL_LISTING_AGENT_STATE'), 'DriveFolder0001', new Date(iso)).split('\n')[0];
  return r;
}

// ====================================================================== 1
console.log('1. localListingScheduledPause_() (Calgary: ABT = UTC-6 all year from 2026-11-01)');
{
  const { s } = report({});
  const p = (now, scan, zone) => { const r = s.localListingScheduledPause_(zone === undefined ? EDMONTON : zone, scan, new Date(now)); return r ? 'pause ' + r.lastScan + ' / ' + r.resumes : 'failure'; };
  const LAST = '2026-11-17T01:28:00Z'; // 19:28 ABT on 16 November
  check('21:00 ABT, last scan 19:28 ABT: the scheduled pause', p('2026-11-17T03:00:00Z', LAST), 'pause 2026-11-16 19:28 ABT / 07:30 ABT');
  check('23:59 ABT: pause', p('2026-11-17T05:59:00Z', LAST), 'pause 2026-11-16 19:28 ABT / 07:30 ABT');
  check('00:00 ABT (the next day): pause', p('2026-11-17T06:00:00Z', LAST), 'pause 2026-11-16 19:28 ABT / 07:30 ABT');
  check('07:44 ABT (first scan of the day still on its way): pause', p('2026-11-17T13:44:00Z', LAST), 'pause 2026-11-16 19:28 ABT / 07:30 ABT');
  check('07:45 ABT and still no new scan: failure (the agent did not resume)', p('2026-11-17T13:45:00Z', LAST), 'failure');
  check('14:00 ABT, stale: failure (daytime)', p('2026-11-16T20:00:00Z', '2026-11-16T19:00:00Z'), 'failure');
  check('21:00 ABT, last scan 15:00 ABT: failure (stopped during the day)', p('2026-11-17T03:00:00Z', '2026-11-16T21:00:00Z'), 'failure');
  check('last scan 19:00 ABT (30 minutes before the close): pause', p('2026-11-17T03:00:00Z', '2026-11-17T01:00:00Z'), 'pause 2026-11-16 19:00 ABT / 07:30 ABT');
  check('last scan 18:59 ABT: failure', p('2026-11-17T03:00:00Z', '2026-11-17T00:59:00Z'), 'failure');
  check('03:00 ABT, last scan the evening before yesterday: failure (a whole day missed)', p('2026-11-17T09:00:00Z', '2026-11-16T01:28:00Z'), 'failure');
  check('the night of Alberta\'s change: last scan 19:28 MDT on 31 Oct, 07:40 ABT on 1 Nov: pause', p('2026-11-01T13:40:00Z', '2026-11-01T01:28:00Z'), 'pause 2026-10-31 19:28 MDT / 07:30 ABT');
  check('Europe/Berlin, the evening DST starts (2027-03-28): 22:00 CEST, last scan 19:28 CEST: pause', p('2027-03-28T20:00:00Z', '2027-03-28T17:28:00Z', 'Europe/Berlin'), 'pause 2027-03-28 19:28 CEST / 07:30 CEST');
  check('Europe/Berlin, 07:40 CEST the next morning, last scan 19:28 CEST: pause', p('2027-03-29T05:40:00Z', '2027-03-28T17:28:00Z', 'Europe/Berlin'), 'pause 2027-03-28 19:28 CEST / 07:30 CEST');
  check('no meeting zone: never taken for the pause', p('2026-11-17T03:00:00Z', LAST, ''), 'failure');
  check('no scan time: failure', p('2026-11-17T03:00:00Z', 'garbage'), 'failure');
}

// ====================================================================== 1b
console.log('1b. with OUTDATED runtime zone rules (America/Edmonton still MST, UTC-7, in winter): Alberta\'s rule decides');
{
  // Apps Script's (or Node's) zone data may predate Alberta's change. Simulate that: Utilities.formatDate answers with the
  // old rules -- MDT (-0600) from the second Sunday of March to the first Sunday of November, MST (-0700) otherwise.
  const outdated = (s) => {
    s.Utilities.formatDate = (date, zone, pattern) => {
      if (pattern !== 'Z') return date.toISOString();
      if (zone !== EDMONTON) throw new Error('only America/Edmonton is simulated');
      const t = date.getTime();
      const summer = (t >= Date.UTC(2026, 2, 8, 9) && t < Date.UTC(2026, 10, 1, 8)) || (t >= Date.UTC(2027, 2, 14, 9) && t < Date.UTC(2027, 10, 7, 8));
      return summer ? '-0600' : '-0700';
    };
  };
  const { s } = report({});
  outdated(s);
  check('the simulation is old: the runtime says -0700 for 17 Nov 2026', s.Utilities.formatDate(new Date('2026-11-17T12:00:00Z'), EDMONTON, 'Z'), '-0700');
  check('the report still uses UTC-6 (ABT) for 17 Nov 2026 and 15 Jan 2027', [s.meetingTimeZoneOffsetMinutes_(EDMONTON, Date.parse('2026-11-17T12:00:00Z')), s.meetingTimeZoneOffsetMinutes_(EDMONTON, Date.parse('2027-01-15T12:00:00Z'))], [-360, -360]);
  const p = (now, scan) => { const r = s.localListingScheduledPause_(EDMONTON, scan, new Date(now)); return r ? 'pause ' + r.lastScan + ' / ' + r.resumes : 'failure'; };
  check('21:00 ABT, last scan 19:28 ABT: pause, said in ABT', p('2026-11-17T03:00:00Z', '2026-11-17T01:28:00Z'), 'pause 2026-11-16 19:28 ABT / 07:30 ABT');
  check('07:44 ABT (06:44 by the old rules): pause', p('2026-11-17T13:44:00Z', '2026-11-17T01:28:00Z'), 'pause 2026-11-16 19:28 ABT / 07:30 ABT');
  check('07:45 ABT (06:45 by the old rules, which would still be night): failure', p('2026-11-17T13:45:00Z', '2026-11-17T01:28:00Z'), 'failure');
  check('last scan 18:59 ABT (17:59 by the old rules): failure', p('2026-11-17T03:00:00Z', '2026-11-17T00:59:00Z'), 'failure');
  check('winter 2027, 19:30 ABT is 01:30 UTC: 19:29 ABT is still the window (no pause), 19:31 is the pause',
    [p('2027-01-16T01:29:00Z', '2027-01-16T01:00:00Z'), p('2027-01-16T01:31:00Z', '2027-01-16T01:00:00Z')], ['failure', 'pause 2027-01-15 19:00 ABT / 07:30 ABT']);
  check('summer 2027 (old rules: MDT, also UTC-6): pause at 21:00 ABT', p('2027-07-02T03:00:00Z', '2027-07-02T01:28:00Z'), 'pause 2027-07-01 19:28 ABT / 07:30 ABT');
  // The same with the rule table emptied: the outdated runtime then decides, and the result is wrong -- so the checks
  // above really depend on Alberta's rule, not on the runtime happening to be current.
  const { s: bare } = report({});
  outdated(bare);
  bare.MEETING_TIME_ZONE_RULE_CHANGES_ = {}; // a top-level var of Code.js: a property of the sandbox
  check('control: without Alberta\'s rule the runtime\'s -0700 is used', bare.meetingTimeZoneOffsetMinutes_(EDMONTON, Date.parse('2026-11-17T03:00:00Z')), -420);
  check('control: ... and the normal overnight pause (21:00 ABT, last scan 19:28 ABT) would be taken for a failure (the scan looks like 18:28 MST)',
    bare.localListingScheduledPause_(EDMONTON, '2026-11-17T01:28:00Z', new Date('2026-11-17T03:00:00Z')), null);
}

// ====================================================================== 2
console.log('2. importAgentListing_(): paused vs stale; nothing imported, nothing erased');
{
  const STORED = JSON.stringify([{ file: A + '_r1.docx', url: LOCAL + '/Drafts/MBS/' + A + '_r1.docx' }]);
  const night = report({ 'MBS.listing.json': listing([A + '_r1.docx', A + '_r2.docx'], '2026-11-17T01:28:00Z') }, EDMONTON);
  night.props.setProperty('LOCALREV_' + A, STORED);
  const r1 = night.importAt('2026-11-17T03:00:00Z');
  check('overnight after a full day: "paused"', r1.result, 'paused');
  check('said as the scheduled pause, with the last scan and the resumption, not as a failure',
    r1.state.message, 'Not imported: scheduled overnight pause. The local listing agent scans from 07:30 to 19:30 meeting time; its last scan of MBS was at 2026-11-16 19:28 ABT, scanning resumes at 07:30 ABT. This is not a synchronization failure. Previously imported revisions are kept.');
  check('nothing imported, nothing erased', [night.props.getProperty('LOCALREV_' + A), r1.touched], [STORED, []]);
  check('Check Connections, first line: the pause', /-- Not imported: scheduled overnight pause\./.test(night.status('2026-11-17T03:00:00Z')), true);

  const stopped = report({ 'MBS.listing.json': listing([A + '_r1.docx'], '2026-11-16T21:00:00Z') }, EDMONTON);
  stopped.props.setProperty('LOCALREV_' + A, STORED);
  const r2 = stopped.importAt('2026-11-17T03:00:00Z');
  check('overnight, but the agent stopped at 15:00: "stale", the genuine failure', [r2.result, /is stale: scanned .* \(agent stopped, meeting-room network down, or Drive not synchronized\)/.test(r2.state.message)], ['stale', true]);
  check('... also nothing imported, nothing erased', stopped.props.getProperty('LOCALREV_' + A), STORED);

  const day = report({ 'MBS.listing.json': listing([A + '_r1.docx'], '2026-11-16T19:00:00Z') }, EDMONTON);
  check('during the day: "stale"', day.importAt('2026-11-16T20:00:00Z').result, 'stale');

  const morning = report({ 'MBS.listing.json': listing([A + '_r1.docx'], '2026-11-17T01:28:00Z') }, EDMONTON);
  check('07:40 ABT, no new scan yet: still "paused"', morning.importAt('2026-11-17T13:40:00Z').result, 'paused');
  check('07:50 ABT, no new scan: "stale" (the agent did not resume)', morning.importAt('2026-11-17T13:50:00Z').result, 'stale');
  morning.drive = { 'MBS.listing.json': listing([A + '_r1.docx'], '2026-11-17T13:32:00Z') };
  check('07:32 ABT scan arrives: imported', morning.importAt('2026-11-17T13:35:00Z').result, 'imported');

  const noZone = report({ 'MBS.listing.json': listing([A + '_r1.docx'], '2026-11-17T01:28:00Z') }, null);
  check('no meeting zone configured: "stale" as before (nothing guessed)', noZone.importAt('2026-11-17T03:00:00Z').result, 'stale');
  const invalid = report({ 'MBS.listing.json': '{"format":"sa4-local-listing","version":1' }, EDMONTON);
  check('an incomplete file overnight: "invalid", never the pause', invalid.importAt('2026-11-17T03:00:00Z').result, 'invalid');
}

// ====================================================================== 3
console.log('3. what changed in Code.js (source)');
{
  const now = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutOvernightPause(now);
  check('the take-out finds every change exactly once', view.problems, []);
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (code, name) => { const at = code.indexOf('\nfunction ' + name + '('); return code.slice(at, code.indexOf('\n}\n', at)); };
  const before = names(view.code);
  check('the new functions', names(now).filter((n) => before.indexOf(n) === -1).sort(), OVERNIGHT_PAUSE_NEW_FUNCTIONS.slice().sort());
  check('the changed functions', before.filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)).sort(), OVERNIGHT_PAUSE_CHANGED_FUNCTIONS.slice().sort());
  check('untouched: validation, Drive reading, storage, rendering, the update', ['validateAgentListing_', 'readAgentListingFiles_', 'storeLocalRevisions_', 'renderStoredLocalRevisions_', 'updateRevisions_', 'describeLocalListingAgentStatus_']
    .filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)), []);
  check('a pause counts as "not read" in the update, like stale (cells keep what they say)', /const localRead = !agent \|\| agent\.result === 'imported' \|\| agent\.result === 'unchanged';/.test(bodyOf(now, 'updateRevisions_')), true);
  check('the new function is pure (no service calls)', /(DriveApp|PropertiesService|DocumentApp|UrlFetchApp)/.test(bodyOf(now, 'localListingScheduledPause_')), false);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll overnight-pause checks passed.');
process.exit(failures ? 1 : 0);
