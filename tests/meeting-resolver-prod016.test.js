/**
 * PROD-016 — Separate core meeting resolution from TDoc enrichment.
 *
 * PROD-014 removed the synchronous revisions/drafts validation probe and
 * made GetiCal conditional, but a fresh live smoke test of meeting 86178
 * still hung 90+ seconds -- proof that TdocList.aspx itself (the one
 * remaining unconditional network call besides GetMeetings) was also part
 * of the problem. This suite proves the fix: resolveMeetingCoreById_()
 * makes ONLY the GetMeetings call (plus, rarely, a conditional GetiCal
 * fallback -- PROD-014, unchanged) and NEVER touches TdocList.aspx or the
 * revisions probe; agenda/TDoc discovery is now a wholly separate,
 * explicit operation (enrichMeetingFromTdocList_()). resolveMeetingById_()
 * is kept, composed from the two, with its full original behavior/shape
 * unchanged -- see tests/meeting-resolver.test.js, which continues to
 * pass against it unmodified.
 *
 * Run: node tests/meeting-resolver-prod016.test.js
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

// Throws if reached -- a stronger guard than "returns empty/404" fixtures,
// since ANY call at all (successful or not) to a network function the core
// path must never touch is itself the regression this suite exists to
// catch.
function throwIfCalled(label) {
  return () => { throw new Error(label + ' must not be called by resolveMeetingCoreById_() (PROD-016)'); };
}

// ==================== 1. resolveMeetingCoreById_() -- network isolation ===

console.log('resolveMeetingCoreById_() -- GetMeetings [+ conditional GetiCal] ONLY, never TdocList/revisions');

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  sandbox.fetchMeetingIcalById_ = throwIfCalled('fetchMeetingIcalById_'); // 86178 supplies all 4 fallback fields -- should never be reached
  sandbox.fetchMeetingTdocListById_ = throwIfCalled('fetchMeetingTdocListById_');
  sandbox.fetchRevisionsUrlCandidate_ = throwIfCalled('fetchRevisionsUrlCandidate_');

  const result = sandbox.resolveMeetingCoreById_(86178);
  check('86178 core: id resolves', result.id, 86178);
  check('86178 core: name resolves', result.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('86178 core: type resolves', result.meeting.type, 'adhoc');
  check('86178 core: startDate resolves (raw)', result.meeting.startDate, '2026-09-22 15:00:00');
  check('86178 core: endDate resolves (raw)', result.meeting.endDate, FIXTURES.getMeetings86178[0].EndDate);
  check('86178 core: location resolves', result.meeting.location, FIXTURES.getMeetings86178[0].Location);
  check('86178 core: tb/tbId resolve', [result.meeting.tb, result.meeting.tbId], [FIXTURES.getMeetings86178[0].TB, FIXTURES.getMeetings86178[0].TBId]);
  check('86178 core: ftpBase resolves', /\/SA4_Plenary\/Docs\/$/.test(result.sources.ftpBase), true);
  check('86178 core: revisionsUrlCandidate derived', result.sources.revisionsUrlCandidate,
    'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/');
  check('86178 core: sources.revisionsUrl is null (never validated)', result.sources.revisionsUrl, null);
  check('86178 core: documents is ALWAYS null (agenda/TDoc discovery is separate)', result.documents, null);
  check('86178 core: unresolved lists documents.agendaTdoc/documents.family as not-yet-discovered',
    ['documents.agendaTdoc', 'documents.family'].every(u => result.unresolved.indexOf(u) !== -1), true);
}

{
  // Degraded metadata (missing StartDate): GetiCal IS still the genuine
  // fallback and should be fetched -- proves the conditional-skip logic
  // (not a blanket "never call iCal") carried over correctly.
  const { sandbox } = loadCode();
  const degraded = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
  delete degraded[0].StartDate;
  let icalCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(degraded) });
  sandbox.fetchMeetingIcalById_ = () => { icalCalled = true; return { statusCode: 200, text: FIXTURES.ical86178 }; };
  sandbox.fetchMeetingTdocListById_ = throwIfCalled('fetchMeetingTdocListById_');
  sandbox.fetchRevisionsUrlCandidate_ = throwIfCalled('fetchRevisionsUrlCandidate_');

  const result = sandbox.resolveMeetingCoreById_(86178);
  check('degraded metadata: GetiCal fallback still fetched (unaffected by PROD-016)', icalCalled, true);
  check('degraded metadata: startDate falls back to iCal, TdocList still never touched',
    result.meeting.startDate, FIXTURES.ical86178.match(/^DTSTART:(.*)$/m)[1].trim());
}

{
  // Invalid id and "not found" cases -- never reach any network function.
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: '[]' });
  sandbox.fetchMeetingIcalById_ = throwIfCalled('fetchMeetingIcalById_');
  sandbox.fetchMeetingTdocListById_ = throwIfCalled('fetchMeetingTdocListById_');
  const notFound = sandbox.resolveMeetingCoreById_(1);
  check('meeting not found: meeting is null, documents null, no throw', [notFound.meeting, notFound.documents], [null, null]);

  const invalid = sandbox.resolveMeetingCoreById_('not-a-number');
  check('invalid id: meeting is null, no network attempted', invalid.meeting, null);
}

// ==================== 2. enrichMeetingFromTdocList_() ======================

console.log('enrichMeetingFromTdocList_() -- TdocList.aspx ONLY, merges into a clone of the core result');

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  const core = sandbox.resolveMeetingCoreById_(86178);

  let metadataCalledDuringEnrich = false;
  sandbox.fetchMeetingMetadataById_ = () => { metadataCalledDuringEnrich = true; return { statusCode: 200, text: '[]' }; };
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });

  const enriched = sandbox.enrichMeetingFromTdocList_(86178, core);
  check('enrichment does NOT re-fetch GetMeetings (core result is reused as-is)', metadataCalledDuringEnrich, false);
  check('enrichment finds S4aP260098', enriched.documents.agendaTdoc, 'S4aP260098');
  check('enrichment finds the plenary-adhoc family', enriched.documents.family, 'plenary-adhoc');
  check('enrichment preserves core fields untouched (meeting.name)', enriched.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('enrichment preserves core fields untouched (sources.ftpBase)', enriched.sources.ftpBase, core.sources.ftpBase);
  check('original coreResult object is not mutated (documents still null on it)', core.documents, null);
  check('enrichment removes documents.agendaTdoc/documents.family from unresolved once found',
    ['documents.agendaTdoc', 'documents.family'].some(u => enriched.unresolved.indexOf(u) !== -1), false);
}

{
  // 86174: real negative case -- enrichment must not invent an agenda TDoc.
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86174) });
  const core = sandbox.resolveMeetingCoreById_(86174);
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86174 });
  const enriched = sandbox.enrichMeetingFromTdocList_(86174, core);
  check('86174: enrichment leaves agendaTdoc null -- not guessed', enriched.documents.agendaTdoc, null);
  check('86174: unresolved still contains documents.agendaTdoc', enriched.unresolved.indexOf('documents.agendaTdoc') !== -1, true);
}

{
  // 60778: enrichment must still resolve the supersession correctly.
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings60778) });
  const core = sandbox.resolveMeetingCoreById_(60778);
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml60778RevisionPair });
  const enriched = sandbox.enrichMeetingFromTdocList_(60778, core);
  check('60778: enrichment selects S4-261392 (approved, current)', enriched.documents.agendaTdoc, 'S4-261392');
}

{
  // Enrichment failure cannot erase core resolution -- documents stays
  // exactly as the core result had it (null), plus a warning.
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  const core = sandbox.resolveMeetingCoreById_(86178);
  sandbox.fetchMeetingTdocListById_ = () => { throw new Error('simulated TdocList outage'); };
  const enriched = sandbox.enrichMeetingFromTdocList_(86178, core);
  check('enrichment failure: documents stays null (core untouched)', enriched.documents, null);
  check('enrichment failure: core meeting fields still present', enriched.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('enrichment failure: a warning records it', enriched.warnings.some(w => /simulated TdocList outage/.test(w)), true);
  check('enrichment failure: does not throw', typeof enriched, 'object');
}

{
  // Mismatched meetingId -- refuses to merge, does not even attempt the fetch.
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  const core = sandbox.resolveMeetingCoreById_(86178);
  sandbox.fetchMeetingTdocListById_ = throwIfCalled('fetchMeetingTdocListById_');
  const mismatched = sandbox.enrichMeetingFromTdocList_(60778, core);
  check('mismatched meetingId: documents stays null', mismatched.documents, null);
  check('mismatched meetingId: a warning explains the skip', mismatched.warnings.some(w => /does not match/.test(w)), true);
}

{
  // No meeting to enrich (core never found one) -- refuses cleanly.
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: '[]' });
  const core = sandbox.resolveMeetingCoreById_(1);
  sandbox.fetchMeetingTdocListById_ = throwIfCalled('fetchMeetingTdocListById_');
  const result = sandbox.enrichMeetingFromTdocList_(1, core);
  check('nothing to enrich: returns the core result unchanged, no fetch attempted', result.meeting, null);
}

// ============ 3. resolveMeetingById_() -- unchanged full-resolution shape =

console.log('resolveMeetingById_() -- composed core+enrichment, behavior/shape unchanged from before PROD-016');

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });
  const result = sandbox.resolveMeetingById_(86178);
  check('86178: full resolveMeetingById_() still returns core AND enrichment together',
    [result.meeting.name, result.documents.agendaTdoc], ['SA4-e (AH) on FS_6G_MED', 'S4aP260098']);
}

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: '[]' });
  let tdocCalled = false;
  sandbox.fetchMeetingTdocListById_ = () => { tdocCalled = true; return { statusCode: 200, text: '' }; };
  const result = sandbox.resolveMeetingById_(1);
  check('meeting not found: resolveMeetingById_() still short-circuits before enrichment (TdocList never called)', tdocCalled, false);
  check('meeting not found: shape matches the original early-return', result.meeting, null);
}

// ==================== 4. diagnoseMeetingResolverTiming_() `only` param =====

console.log('diagnoseMeetingResolverTiming_(meetingId, only) -- selective diagnostics without the full resolver');

{
  const { sandbox } = loadCode();
  let metadataCalled = false, icalCalled = false, tdocCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => { metadataCalled = true; return { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) }; };
  sandbox.fetchMeetingIcalById_ = () => { icalCalled = true; return { statusCode: 200, text: '' }; };
  sandbox.fetchMeetingTdocListById_ = () => { tdocCalled = true; return { statusCode: 200, text: '' }; };
  sandbox.Logger = { log: () => {} };

  const report = sandbox.diagnoseMeetingResolverTiming_(86178, ['metadata']);
  check('only:["metadata"] -- GetMeetings called', metadataCalled, true);
  check('only:["metadata"] -- GetiCal NOT called', icalCalled, false);
  check('only:["metadata"] -- TdocList NOT called', tdocCalled, false);
  check('only:["metadata"] -- report has exactly 1 entry', report.length, 1);
  check('only:["metadata"] -- entry label mentions GetMeetings', /GetMeetings/.test(report[0].label), true);
}

{
  const { sandbox } = loadCode();
  let metadataCalled = false, tdocCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => { metadataCalled = true; return { statusCode: 200, text: '[]' }; };
  sandbox.fetchMeetingTdocListById_ = () => { tdocCalled = true; return { statusCode: 200, text: '' }; };

  const report = sandbox.diagnoseMeetingResolverTiming_(86178, ['tdoc']);
  check('only:["tdoc"] -- GetMeetings NOT called', metadataCalled, false);
  check('only:["tdoc"] -- TdocList called', tdocCalled, true);
  check('only:["tdoc"] -- report has exactly 1 entry', report.length, 1);
}

{
  // Omitting `only` still times all four -- unchanged PROD-014 default.
  const { sandbox } = loadCode();
  let calls = 0;
  sandbox.fetchMeetingMetadataById_ = () => { calls++; return { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) }; };
  sandbox.fetchMeetingIcalById_ = () => { calls++; return { statusCode: 200, text: '' }; };
  sandbox.fetchMeetingTdocListById_ = () => { calls++; return { statusCode: 200, text: '' }; };
  sandbox.fetchRevisionsUrlCandidate_ = () => { calls++; return { statusCode: 200, text: '<a href="x">x</a>' }; };
  sandbox.diagnoseMeetingResolverTiming_(86178);
  check('no `only` argument -- all four sources are still probed (backward compatible default)', calls, 4);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
