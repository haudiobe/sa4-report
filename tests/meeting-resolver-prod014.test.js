/**
 * PROD-014 — Prevent optional source checks from blocking meeting
 * resolution.
 *
 * A live production smoke test found the Configure Meeting Settings
 * dialog's Resolve action hanging 3+ minutes with no response for meeting
 * 86178. resolveMeetingById_() makes up to four sequential network calls
 * (GetMeetings, GetiCal, TdocList.aspx, and ARCH-011's revisions/drafts
 * folder probe) with NO per-request timeout available on Apps Script's
 * UrlFetchApp -- any one of them hanging stalls the whole chain, and
 * nothing downstream can distinguish "still working" from "will never
 * return." The revisions probe's target host was already known (from
 * ARCH-011's own investigation) to sit behind a WAF that blocked a plain
 * HTTP client, making it the most plausible cause.
 *
 * This suite proves the two structural changes that address this:
 *   1. shouldFetchIcalFallback_() / the GetiCal fetch is skipped whenever
 *      GetMeetings alone already supplies everything iCal could ever
 *      substitute for (the normal-success case for every real fixture
 *      meeting) -- see tests/meeting-config-merge.test.js and
 *      tests/meeting-revisions-folder.test.js for the (already updated)
 *      revisions-folder derive-only behavior covered there.
 *   2. Core metadata resolution (Meeting ID/Name/Type/Date/FTP Base)
 *      succeeds independently of, and is never blocked or degraded by,
 *      the optional GetiCal/revisions sources being slow, failing, or
 *      throwing.
 *
 * Run: node tests/meeting-resolver-prod014.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode } = require('./helpers/load-code.js');

const FIXTURES = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-resolver-samples.json'), 'utf8'));

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

// ============================== 1. shouldFetchIcalFallback_() ==============

console.log('shouldFetchIcalFallback_() -- pure decision, deterministic');

{
  const { sandbox } = loadCode();
  const fn = sandbox.shouldFetchIcalFallback_;

  check('GetMeetings failed outright -> iCal is the only remaining source, fetch it',
    fn({ ok: false, meeting: null, error: 'network error' }), true);

  check('GetMeetings succeeded but found nothing (meeting: null) -> nothing to fall back into, skip',
    fn({ ok: true, meeting: null, error: null }), false);

  check('GetMeetings succeeded with ALL four substitutable fields present (real 86178 shape) -> skip',
    fn({ ok: true, meeting: { Title: 'x', StartDate: 'x', EndDate: 'x', Location: 'x' }, error: null }), false);

  check('GetMeetings succeeded but Title is missing -> fetch',
    fn({ ok: true, meeting: { StartDate: 'x', EndDate: 'x', Location: 'x' }, error: null }), true);
  check('GetMeetings succeeded but StartDate is missing -> fetch',
    fn({ ok: true, meeting: { Title: 'x', EndDate: 'x', Location: 'x' }, error: null }), true);
  check('GetMeetings succeeded but EndDate is missing -> fetch',
    fn({ ok: true, meeting: { Title: 'x', StartDate: 'x', Location: 'x' }, error: null }), true);
  check('GetMeetings succeeded but Location is missing -> fetch',
    fn({ ok: true, meeting: { Title: 'x', StartDate: 'x', EndDate: 'x' }, error: null }), true);

  check('null metadataParsed does not throw, defaults to "fetch" (safest default)', fn(null), true);
  check('undefined metadataParsed does not throw, defaults to "fetch"', fn(undefined), true);

  // Every real captured fixture meeting has all four fields -- confirms
  // the normal-success skip actually applies to real data, not just
  // synthetic examples.
  [
    ['86178', FIXTURES.getMeetings86178],
    ['86174', FIXTURES.getMeetings86174],
    ['85916', FIXTURES.getMeetings85916],
    ['60778', FIXTURES.getMeetings60778]
  ].forEach(([label, fixture]) => {
    check(`real ${label} GetMeetings response has all four fields -> shouldFetchIcalFallback_ is false`,
      fn({ ok: true, meeting: fixture[0], error: null }), false);
  });
}

// ==================== 2. GetiCal is skipped on the normal success path =====

console.log('resolveMeetingById_() -- GetiCal is never called when GetMeetings alone is sufficient (86178)');

{
  const { sandbox } = loadCode();
  let icalFetchCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  sandbox.fetchMeetingIcalById_ = () => { icalFetchCalled = true; return { statusCode: 200, text: FIXTURES.ical86178 }; };
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });

  const result = sandbox.resolveMeetingById_(86178);
  check('86178: GetiCal is never fetched (GetMeetings alone was sufficient)', icalFetchCalled, false);
  check('86178: core fields still resolve correctly without it', {
    name: result.meeting.name,
    type: result.meeting.type,
    startDate: result.meeting.startDate,
    endDate: result.meeting.endDate,
    location: result.meeting.location
  }, {
    name: 'SA4-e (AH) on FS_6G_MED',
    type: 'adhoc',
    startDate: '2026-09-22 15:00:00',
    endDate: FIXTURES.getMeetings86178[0].EndDate,
    location: FIXTURES.getMeetings86178[0].Location
  });
}

{
  // A meeting genuinely not found: GetiCal must not be fetched either
  // (the early "not found" return never consults it).
  const { sandbox } = loadCode();
  let icalFetchCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: '[]' });
  sandbox.fetchMeetingIcalById_ = () => { icalFetchCalled = true; return { statusCode: 200, text: '' }; };
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: '' });
  sandbox.resolveMeetingById_(1);
  check('meeting not found: GetiCal is never fetched either (its result would never be used)', icalFetchCalled, false);
}

{
  // Degraded metadata: GetMeetings succeeds but is missing StartDate --
  // GetiCal SHOULD still be fetched and its fallback still used (this
  // genuine fallback behavior must be preserved, not removed).
  const { sandbox } = loadCode();
  const degradedMeeting = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
  delete degradedMeeting[0].StartDate;
  let icalFetchCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(degradedMeeting) });
  sandbox.fetchMeetingIcalById_ = () => { icalFetchCalled = true; return { statusCode: 200, text: FIXTURES.ical86178 }; };
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });
  const result = sandbox.resolveMeetingById_(86178);
  check('degraded metadata (missing StartDate): GetiCal IS fetched (genuine fallback case)', icalFetchCalled, true);
  check('degraded metadata: startDate falls back to the iCal DTSTART, preserved fallback behavior',
    result.meeting.startDate, FIXTURES.ical86178.match(/^DTSTART:(.*)$/m)[1].trim());
}

{
  // GetMeetings itself fails outright: GetiCal must still be attempted --
  // it is the only remaining source, unaffected by this task.
  const { sandbox } = loadCode();
  let icalFetchCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => { throw new Error('simulated GetMeetings outage'); };
  sandbox.fetchMeetingIcalById_ = () => { icalFetchCalled = true; return { statusCode: 200, text: FIXTURES.ical86178 }; };
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });
  sandbox.resolveMeetingById_(86178);
  check('GetMeetings outage: GetiCal IS still attempted (only remaining source)', icalFetchCalled, true);
}

// ============== 3. Core resolution independent of optional-source failure ==

console.log('core resolution (Meeting ID/Name/Type/Date/FTP Base) succeeds independently of GetiCal/revisions failures');

{
  const { sandbox } = loadCode();
  // Force GetiCal to be attempted (degraded metadata) AND fail/throw, and
  // confirm core fields -- which don't depend on it here -- still resolve.
  const degradedMeeting = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
  delete degradedMeeting[0].Location; // only Location is missing -- name/type/date/ftpBase are unaffected either way
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(degradedMeeting) });
  sandbox.fetchMeetingIcalById_ = () => { throw new Error('simulated GetiCal outage'); };
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });
  // Revisions probe must never even be reached (PROD-014) -- throws if it is.
  sandbox.fetchRevisionsUrlCandidate_ = () => { throw new Error('must not be called (PROD-014)'); };

  const result = sandbox.resolveMeetingById_(86178);
  check('core Meeting ID resolves', result.id, 86178);
  check('core Meeting Name resolves', result.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('core Meeting Type resolves', result.meeting.type, 'adhoc');
  check('core Meeting Date (startDate) resolves', result.meeting.startDate, '2026-09-22 15:00:00');
  check('core FTP Base resolves', /\/SA4_Plenary\/Docs\/$/.test(result.sources.ftpBase), true);
  check('a GetiCal outage is recorded as a warning, not a fatal error', result.warnings.some(w => /simulated GetiCal outage/.test(w)), true);
  check('the whole call still returns a usable object despite the GetiCal outage', typeof result, 'object');
}

// ==================== 4. established core-field regression guards ==========

console.log('86178/86174/60778 still resolve their previously-established core fields (no regression from PROD-014)');

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });
  const result = sandbox.resolveMeetingById_(86178);
  check('86178: name', result.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('86178: type', result.meeting.type, 'adhoc');
  check('86178: agendaTdoc', result.documents.agendaTdoc, 'S4aP260098');
}

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86174) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86174 });
  const result = sandbox.resolveMeetingById_(86174);
  check('86174: agendaTdoc is still null -- not invented', result.documents.agendaTdoc, null);
}

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings60778) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml60778RevisionPair });
  const result = sandbox.resolveMeetingById_(60778);
  check('60778: still chooses S4-261392 (approved, current version)', result.documents.agendaTdoc, 'S4-261392');
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
