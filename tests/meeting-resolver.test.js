/**
 * ARCH-009 — Meeting-ID resolver core regression coverage.
 *
 * This suite exercises resolveMeetingById_() and every pure helper it's
 * built from (parseMeetingIdInput_, parseMeetingMetadataResponse_,
 * normalizePortalMeetingTitle_, normalizePortalMeetingType_,
 * normalizeMtgDocUrlToFtpBase_, parseMeetingIcal_,
 * parseMeetingTdocListHtml_, selectAgendaCandidate_,
 * deriveTdocFamilyEvidence_) against REAL, captured raw responses from the
 * three anonymous official 3GPP/ETSI Portal endpoints discovered during
 * the ARCH-007/ARCH-008 investigations -- see
 * tests/fixtures/meeting-resolver-samples.json for exactly what was
 * captured and how.
 *
 * resolveMeetingById_() itself calls three network functions
 * (fetchMeetingMetadataById_/fetchMeetingIcalById_/
 * fetchMeetingTdocListById_) that this suite overrides on the sandbox with
 * fixture-returning stubs -- the SAME "override a global binding on the
 * loaded sandbox" pattern already used elsewhere in this project's tests
 * (e.g. tests/reviewer-api-tdoc-family.test.js stubbing
 * fetchAndAddAbstract_, tests/meeting-86178-production.test.js stubbing
 * DocumentApp.getActiveDocument). This exercises the REAL orchestration
 * logic in resolveMeetingById_(), not a reimplementation of it.
 *
 * NONE of this is wired into getMeetingContext_() or any other production
 * consumer yet -- this suite proves the resolver core in isolation only.
 *
 * Run: node tests/meeting-resolver.test.js
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

// Wires the three network functions to fixture data (or explicit
// statusCode/error overrides for failure-mode tests), then returns the
// sandbox so callers can invoke resolveMeetingById_() or any pure helper.
function loadResolverSandbox(overrides) {
  const { sandbox } = loadCode();
  const o = overrides || {};
  sandbox.fetchMeetingMetadataById_ = () => {
    if (o.metadataThrow) throw new Error(o.metadataThrow);
    return o.metadata !== undefined ? o.metadata : { statusCode: 200, text: '[]' };
  };
  sandbox.fetchMeetingIcalById_ = () => {
    if (o.icalThrow) throw new Error(o.icalThrow);
    return o.ical !== undefined ? o.ical : { statusCode: 200, text: '' };
  };
  sandbox.fetchMeetingTdocListById_ = () => {
    if (o.tdocThrow) throw new Error(o.tdocThrow);
    return o.tdoc !== undefined ? o.tdoc : { statusCode: 200, text: '' };
  };
  return sandbox;
}

// ============================================== 1. Meeting-ID validation ==

console.log('parseMeetingIdInput_() -- validation and normalization');

{
  const { sandbox } = loadCode();
  const fn = sandbox.parseMeetingIdInput_;

  check('86178 (number) is valid', fn(86178), { isValid: true, id: 86178, error: null });
  check('"86178" (string) is valid, normalized to a number', fn('86178'), { isValid: true, id: 86178, error: null });
  check('"  86178  " (padded string) is valid', fn('  86178  '), { isValid: true, id: 86178, error: null });
  check('null is rejected', fn(null).isValid, false);
  check('undefined is rejected', fn(undefined).isValid, false);
  check('"" is rejected', fn('').isValid, false);
  check('"   " (blank) is rejected', fn('   ').isValid, false);
  check('0 is rejected', fn(0).isValid, false);
  check('-86178 is rejected', fn(-86178).isValid, false);
  check('"-86178" is rejected', fn('-86178').isValid, false);
  check('86178.5 (non-integer) is rejected', fn(86178.5).isValid, false);
  check('"86178abc" (mixed string) is rejected -- digits are NOT silently extracted', fn('86178abc').isValid, false);
  check('"id: 86178" (digits embedded in text) is rejected -- NOT silently extracted', fn('id: 86178').isValid, false);
  check('"86 178" (internal space) is rejected', fn('86 178').isValid, false);
  check('{} (wrong type) is rejected', fn({}).isValid, false);
  check('[86178] (wrong type) is rejected', fn([86178]).isValid, false);
}

// ==================================== 2. GetMeetings response parsing =====

console.log('parseMeetingMetadataResponse_() -- HTTP/JSON/shape failure handling');

{
  const { sandbox } = loadCode();
  const fn = sandbox.parseMeetingMetadataResponse_;

  const real86178 = { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) };
  const real86178Result = fn(real86178);
  check('real 86178 response: ok', real86178Result.ok, true);
  check('real 86178 response: meeting.Id', real86178Result.meeting.Id, 86178);
  check('real 86178 response: meeting.Type', real86178Result.meeting.Type, 'AH');

  check('valid but unknown meeting ID ([] response): ok, meeting null (not an error)',
    fn({ statusCode: 200, text: '[]' }), { ok: true, meeting: null, error: null });

  check('non-200 HTTP status is a structured error, not a throw',
    fn({ statusCode: 500, text: 'error' }).ok, false);
  check('malformed JSON is a structured error, not a throw',
    fn({ statusCode: 200, text: 'not json{' }).ok, false);
  check('a non-array JSON body is a structured error',
    fn({ statusCode: 200, text: '{"not":"an array"}' }).ok, false);
  check('a missing fetchResult is handled without throwing',
    fn(null).ok, false);
}

// ============================================ 3. Title normalization =====

console.log('normalizePortalMeetingTitle_() -- narrow, explicit "3GPP" prefix strip only');

{
  const { sandbox } = loadCode();
  const fn = sandbox.normalizePortalMeetingTitle_;
  check('"3GPPSA4-e (AH) on FS_6G_MED" -> "SA4-e (AH) on FS_6G_MED"',
    fn('3GPPSA4-e (AH) on FS_6G_MED'), 'SA4-e (AH) on FS_6G_MED');
  check('"3GPPSA4#137-e" -> "SA4#137-e"', fn('3GPPSA4#137-e'), 'SA4#137-e');
  check('"3GPPSA4-(AH) Audio SWG on ULBC-MED" -> "SA4-(AH) Audio SWG on ULBC-MED"',
    fn('3GPPSA4-(AH) Audio SWG on ULBC-MED'), 'SA4-(AH) Audio SWG on ULBC-MED');
  check('a title with no leading "3GPP" is left untouched',
    fn('Some Other Title'), 'Some Other Title');
  check('empty/missing input does not throw', fn(''), '');
  check('null input does not throw', fn(null), '');
}

// ============================================= 4. Type normalization =====

console.log('normalizePortalMeetingType_() -- only AH/OR verified, everything else unresolved');

{
  const { sandbox } = loadCode();
  const fn = sandbox.normalizePortalMeetingType_;
  check('"AH" -> adhoc', fn('AH'), { type: 'adhoc', portalType: 'AH', recognized: true });
  check('"OR" -> main', fn('OR'), { type: 'main', portalType: 'OR', recognized: true });
  check('an unverified code is NOT guessed -- type null, recognized false',
    fn('TS'), { type: null, portalType: 'TS', recognized: false });
  check('empty/missing input is unresolved, not guessed', fn(''), { type: null, portalType: null, recognized: false });
}

// ================================= 5. FTP base normalization (all 4 IDs) ==

console.log('normalizeMtgDocUrlToFtpBase_() -- conservative slash/Docs normalization, all four real MtgDocURL values');

{
  const { sandbox } = loadCode();
  const fn = sandbox.normalizeMtgDocUrlToFtpBase_;

  // 86178: real inconsistency -- doubled slash after host, no trailing Docs/.
  check('86178 MtgDocURL (doubled slash, missing /Docs/) is normalized correctly',
    fn(FIXTURES.getMeetings86178[0].MtgDocURL).ftpBase,
    'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/');

  // 86174/85916/60778: already end in /Docs/ -- must be preserved AS-IS,
  // including their differing casing (TSG_SA vs tsg_sa), not "fixed".
  check('86174 MtgDocURL (already /Docs/, "TSG_SA" casing) is preserved unchanged',
    fn(FIXTURES.getMeetings86174[0].MtgDocURL).ftpBase, FIXTURES.getMeetings86174[0].MtgDocURL);
  check('85916 MtgDocURL (already /Docs/) is preserved unchanged',
    fn(FIXTURES.getMeetings85916[0].MtgDocURL).ftpBase, FIXTURES.getMeetings85916[0].MtgDocURL);
  check('60778 MtgDocURL (already /Docs/, "tsg_sa" casing) is preserved unchanged',
    fn(FIXTURES.getMeetings60778[0].MtgDocURL).ftpBase, FIXTURES.getMeetings60778[0].MtgDocURL);

  check('missing MtgDocURL is a structured error, not a throw', fn(null).ftpBase, null);
  check('missing MtgDocURL is a structured error, not a throw', fn('').ftpBase, null);
}

// =============================================== 6. ICS parsing ===========

console.log('parseMeetingIcal_() -- real 86178 ICS');

{
  const { sandbox } = loadCode();
  const parsed = sandbox.parseMeetingIcal_(FIXTURES.ical86178);
  check('uid', parsed.uid, '86178');
  check('summary', parsed.summary, '3GPPSA4-e (AH) on FS_6G_MED');
  check('dtstart', parsed.dtstart, '20260922T130000Z');
  check('dtend', parsed.dtend, '20260926T160000Z');
  check('location', parsed.location, 'Online,');
}

// ======================================== 7. TdocList.aspx HTML parsing ===

console.log('parseMeetingTdocListHtml_() -- real captured row fragments (86178, 60778, 86174)');

{
  const { sandbox } = loadCode();
  const rows86178 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml86178);
  check('86178: 4 real rows parsed', rows86178.length, 4);
  check('86178: row 0 is S4aP260098, the agenda TDoc, decoded/typed correctly', rows86178[0], {
    id: 'S4aP260098',
    url: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/S4aP260098.zip',
    type: 'agenda',
    title: 'Proposed agenda for SA4 AH conf. call on FS_6G_MED (22nd September, 2026)',
    source: 'SA4 Chair (Tencent)',
    status: 'available',
    forAction: 'Approval',
    meetingName: 'SA4-e (AH) on FS_6G_MED',
    agendaItem: '5.1',
    agendaTopic: 'Opening of the session and registration of documents',
    revisionOf: null,
    revisedTo: null
  });
  check('86178: agendaItemsObserved across all 4 rows', rows86178.map(r => r.agendaItem), ['5.1', '5.6.1', '5.6.1', '5.3']);
  check('86178: a real "Revision of" cross-family reference (S4aP260096 revises main-meeting S4-261834) decodes correctly',
    rows86178[2].revisionOf, 'S4-261834');

  const rows60778 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml60778RevisionPair);
  check('60778: 2 rows parsed (the revision pair)', rows60778.length, 2);
  const s375 = rows60778.find(r => r.id === 'S4-261375');
  const s392 = rows60778.find(r => r.id === 'S4-261392');
  check('60778: S4-261375 status is "revised"', s375.status, 'revised');
  check('60778: S4-261375 revisedTo points to S4-261392', s375.revisedTo, 'S4-261392');
  check('60778: S4-261375 revisionOf is null (it is the original)', s375.revisionOf, null);
  check('60778: S4-261392 status is "approved"', s392.status, 'approved');
  check('60778: S4-261392 revisionOf points to S4-261375', s392.revisionOf, 'S4-261375');

  const rows86174 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml86174);
  check('86174: 2 rows parsed', rows86174.length, 2);
  check('86174: neither row has type "agenda" (real, verified negative case)',
    rows86174.every(r => r.type !== 'agenda'), true);

  check('empty/malformed HTML yields zero rows, not a throw',
    sandbox.parseMeetingTdocListHtml_('<html><body>nothing here</body></html>'), []);
  check('null HTML does not throw', sandbox.parseMeetingTdocListHtml_(null), []);
}

// ====================================== 8. Agenda candidate selection =====

console.log('selectAgendaCandidate_() -- 60778 revision pair, 86174 none-found, ambiguity, pure logic');

{
  const { sandbox } = loadCode();

  const rows60778 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml60778RevisionPair);
  const result60778 = sandbox.selectAgendaCandidate_(rows60778);
  check('60778: S4-261392 (approved) selected, S4-261375 (revised, superseded) rejected',
    result60778.agendaTdoc, 'S4-261392');
  check('60778: not ambiguous', result60778.ambiguous, false);

  const rows86174 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml86174);
  const result86174 = sandbox.selectAgendaCandidate_(rows86174);
  check('86174: no agenda TDoc identified (real negative case) -- not guessed from title text',
    result86174.agendaTdoc, null);
  check('86174: not ambiguous, just genuinely absent', result86174.ambiguous, false);
  check('86174: unresolvedReason is populated', typeof result86174.unresolvedReason, 'string');

  const rows86178 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml86178);
  const result86178 = sandbox.selectAgendaCandidate_(rows86178);
  check('86178: S4aP260098 selected (only agenda-typed row, no revision)', result86178.agendaTdoc, 'S4aP260098');

  // Synthetic ambiguity case: two independent, non-superseding agenda rows.
  const ambiguousRows = [
    { id: 'S4-100001', type: 'agenda', status: 'available', revisionOf: null },
    { id: 'S4-100002', type: 'agenda', status: 'available', revisionOf: null }
  ];
  const ambiguousResult = sandbox.selectAgendaCandidate_(ambiguousRows);
  check('two independent agenda candidates -> ambiguous, no silent pick', ambiguousResult.agendaTdoc, null);
  check('ambiguous flag set', ambiguousResult.ambiguous, true);
  check('both candidates listed', ambiguousResult.candidates.map(c => c.id).sort(), ['S4-100001', 'S4-100002']);

  check('empty row list -> no candidates, not ambiguous', sandbox.selectAgendaCandidate_([]).agendaTdoc, null);
  check('empty row list -> not ambiguous', sandbox.selectAgendaCandidate_([]).ambiguous, false);
}

// ============================================ 9. TDoc family evidence =====

console.log('deriveTdocFamilyEvidence_() -- reuses the central SA4_TDOC_FAMILIES registry, surfaces inconsistency');

{
  const { sandbox } = loadCode();

  const rows86178 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml86178);
  const family86178 = sandbox.deriveTdocFamilyEvidence_(rows86178);
  check('86178: family is S4aP (consistent -- the revisionOf cross-reference to S4-261834 is NOT one of these rows\' own id, so it does not pollute this)',
    family86178, { family: 'plenary-adhoc', familiesSeen: { 'plenary-adhoc': 4 }, consistent: true, unrecognizedIds: [] });

  const rows86174 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml86174);
  const family86174 = sandbox.deriveTdocFamilyEvidence_(rows86174);
  check('86174: family is S4aA (audio-adhoc), consistent', family86174.family, 'audio-adhoc');
  check('86174: consistent', family86174.consistent, true);

  const rows60778 = sandbox.parseMeetingTdocListHtml_(FIXTURES.tdocListHtml60778RevisionPair);
  const family60778 = sandbox.deriveTdocFamilyEvidence_(rows60778);
  check('60778: family is main, consistent', family60778.family, 'main');

  // Synthetic inconsistent-family case.
  const mixedRows = [{ id: 'S4aP260098' }, { id: 'S4aA260090' }];
  const mixedFamily = sandbox.deriveTdocFamilyEvidence_(mixedRows);
  check('mixed families -> consistent:false, both surfaced, no silent first-pick',
    mixedFamily.consistent, false);
  check('mixed families -> familiesSeen has both keys',
    Object.keys(mixedFamily.familiesSeen).sort(), ['audio-adhoc', 'plenary-adhoc']);
  check('mixed families -> family is null, not a silent guess', mixedFamily.family, null);

  const unrecognized = sandbox.deriveTdocFamilyEvidence_([{ id: 'not-a-real-id' }]);
  check('unrecognized ids are listed, not silently dropped', unrecognized.unrecognizedIds, ['not-a-real-id']);
}

// ===================================== 10. resolveMeetingById_() -- 86178 =

console.log('resolveMeetingById_() -- full orchestration, meeting 86178');

{
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    ical: { statusCode: 200, text: FIXTURES.ical86178 },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86178 }
  });
  const result = sandbox.resolveMeetingById_(86178);

  check('id', result.id, 86178);
  check('meeting.name', result.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('meeting.portalType', result.meeting.portalType, 'AH');
  check('meeting.type', result.meeting.type, 'adhoc');
  check('meeting.group', result.meeting.group, 'SA4');
  check('meeting.tb', result.meeting.tb, '3GPP SA 4');
  check('meeting.location', result.meeting.location, 'Online');
  check('sources.ftpBase ends /SA4_Plenary/Docs/',
    /\/SA4_Plenary\/Docs\/$/.test(result.sources.ftpBase), true);
  check('documents.family', result.documents.family, 'plenary-adhoc');
  check('documents.agendaTdoc', result.documents.agendaTdoc, 'S4aP260098');
  check('documents.count (from GetMeetings.DocCount, not just the trimmed fixture row count)', result.documents.count, 32);
  check('documents.agendaItemsObserved', result.documents.agendaItemsObserved, ['5.1', '5.6.1', '5.6.1', '5.3']);
  check('mailingList is never derived -- unresolved', result.sources.mailingList, null);
  check('unresolved includes "mailingList"', result.unresolved.indexOf('mailingList') !== -1, true);
  check('unresolved does NOT include meeting.name/type/documents.agendaTdoc/sources.ftpBase (all resolved for 86178)',
    result.unresolved.some(u => ['meeting.name', 'meeting.type', 'documents.agendaTdoc', 'sources.ftpBase'].indexOf(u) !== -1),
    false);
  check('raw.meeting preserves the original, un-normalized Title', result.raw.meeting.Title, '3GPPSA4-e (AH) on FS_6G_MED');
}

// ===================================== 11. resolveMeetingById_() -- 85916 =

console.log('resolveMeetingById_() -- meeting 85916 (metadata only, no TDoc-list fixture needed for this check)');

{
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings85916) },
    tdoc: { statusCode: 200, text: '<html><body></body></html>' }
  });
  const result = sandbox.resolveMeetingById_(85916);
  check('type', result.meeting.type, 'adhoc');
  check('name', result.meeting.name, 'SA4-(AH) Audio SWG on ULBC-MED');
  check('location', result.meeting.location, 'Shanghai');
}

// ===================================== 12. resolveMeetingById_() -- 60778 =

console.log('resolveMeetingById_() -- meeting 60778 (main), agenda supersession resolved end-to-end');

{
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings60778) },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml60778RevisionPair }
  });
  const result = sandbox.resolveMeetingById_(60778);
  check('type', result.meeting.type, 'main');
  check('portalType', result.meeting.portalType, 'OR');
  check('documents.family', result.documents.family, 'main');
  check('documents.agendaTdoc is the approved, current version', result.documents.agendaTdoc, 'S4-261392');
  check('S4-261375 (superseded) is rejected as the current agenda TDoc',
    result.documents.agendaTdoc !== 'S4-261375', true);
  check('documents.count uses GetMeetings.DocCount (482), not the trimmed 2-row fixture', result.documents.count, 482);
}

// ===================================== 13. resolveMeetingById_() -- 86174 =

console.log('resolveMeetingById_() -- meeting 86174 (real negative case: no agenda TDoc)');

{
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86174) },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86174 }
  });
  const result = sandbox.resolveMeetingById_(86174);
  check('type', result.meeting.type, 'adhoc');
  check('documents.family', result.documents.family, 'audio-adhoc');
  check('documents.agendaTdoc is null -- not guessed from title text', result.documents.agendaTdoc, null);
  check('unresolved includes documents.agendaTdoc', result.unresolved.indexOf('documents.agendaTdoc') !== -1, true);
  check('warnings mention the absence explicitly', result.warnings.some(w => /agenda/i.test(w)), true);
}

// ==================================== 14. Failure / warning model =========

console.log('resolveMeetingById_() -- failure and warning behavior');

{
  // Invalid Meeting ID.
  {
    const sandbox = loadResolverSandbox({});
    const result = sandbox.resolveMeetingById_('not-a-number');
    check('invalid id: meeting is null', result.meeting, null);
    check('invalid id: unresolved contains "meeting"', result.unresolved, ['meeting']);
    check('invalid id: a warning/error string is present', result.warnings.length > 0, true);
  }

  // Valid ID, but Portal doesn't recognize it ([] response).
  {
    const sandbox = loadResolverSandbox({ metadata: { statusCode: 200, text: '[]' } });
    const result = sandbox.resolveMeetingById_(1);
    check('unknown id: meeting is null', result.meeting, null);
    check('unknown id: unresolved contains "meeting"', result.unresolved.indexOf('meeting') !== -1, true);
    check('unknown id: does not throw', typeof result, 'object');
  }

  // Network failure (thrown exception) on the primary source.
  {
    const sandbox = loadResolverSandbox({ metadataThrow: 'simulated network failure' });
    const result = sandbox.resolveMeetingById_(86178);
    check('metadata fetch throws: resolveMeetingById_ itself does not throw', typeof result, 'object');
    check('metadata fetch throws: a warning is recorded', result.warnings.some(w => /simulated network failure/.test(w)), true);
  }

  // Malformed metadata JSON.
  {
    const sandbox = loadResolverSandbox({ metadata: { statusCode: 200, text: 'not json{' } });
    const result = sandbox.resolveMeetingById_(86178);
    check('malformed metadata JSON: does not throw', typeof result, 'object');
    check('malformed metadata JSON: a warning is recorded', result.warnings.length > 0, true);
  }

  // Missing MtgDocURL.
  {
    const meetingNoFtp = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
    delete meetingNoFtp[0].MtgDocURL;
    const sandbox = loadResolverSandbox({ metadata: { statusCode: 200, text: JSON.stringify(meetingNoFtp) } });
    const result = sandbox.resolveMeetingById_(86178);
    check('missing MtgDocURL: sources.ftpBase is null', result.sources.ftpBase, null);
    check('missing MtgDocURL: unresolved contains sources.ftpBase', result.unresolved.indexOf('sources.ftpBase') !== -1, true);
  }

  // Empty TDoc list.
  {
    const sandbox = loadResolverSandbox({
      metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
      tdoc: { statusCode: 200, text: '<html><body><table id="rgTdocList"></table></body></html>' }
    });
    const result = sandbox.resolveMeetingById_(86178);
    check('empty TDoc list: agendaTdoc is null', result.documents.agendaTdoc, null);
    check('empty TDoc list: a warning is recorded', result.warnings.some(w => /no recognizable TDoc rows/.test(w)), true);
  }

  // Malformed TDoc-list HTML (network-level failure, not just empty).
  {
    const sandbox = loadResolverSandbox({
      metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
      tdoc: { statusCode: 500, text: 'error' }
    });
    const result = sandbox.resolveMeetingById_(86178);
    check('TdocList.aspx HTTP failure: does not throw, meeting metadata still resolved', result.meeting.name, 'SA4-e (AH) on FS_6G_MED');
    check('TdocList.aspx HTTP failure: a warning is recorded', result.warnings.some(w => /TdocList\.aspx/.test(w)), true);
  }

  // Unknown Portal meeting type.
  {
    const meetingUnknownType = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
    meetingUnknownType[0].Type = 'TS';
    const sandbox = loadResolverSandbox({ metadata: { statusCode: 200, text: JSON.stringify(meetingUnknownType) } });
    const result = sandbox.resolveMeetingById_(86178);
    check('unknown Type code: meeting.type is null, not guessed', result.meeting.type, null);
    check('unknown Type code: raw portalType is preserved', result.meeting.portalType, 'TS');
    check('unknown Type code: unresolved contains meeting.type', result.unresolved.indexOf('meeting.type') !== -1, true);
  }

  // Inconsistent TDoc families end-to-end.
  {
    const mixedHtml = '<html><body><table id="rgTdocList"><tbody>' +
      '<tr class="rgRow"><td></td><td align="center"><a href="https://x/S4aP260098.zip" target="_blank">S4aP260098</a></td><td align="center">agenda</td><td align="left">t</td><td align="center">s</td><td align="center">available</td><td align="center">f</td><td align="center">m</td><td><span title="topic" class="agendaItem">5.1</span></td><td><a></a></td><td></td><td></td></tr>' +
      '<tr class="rgAltRow"><td></td><td align="center"><a href="https://x/S4aA260090.zip" target="_blank">S4aA260090</a></td><td align="center">discussion</td><td align="left">t</td><td align="center">s</td><td align="center">available</td><td align="center">f</td><td align="center">m</td><td><span title="topic" class="agendaItem">5.2</span></td><td><a></a></td><td></td><td></td></tr>' +
      '</tbody></table></body></html>';
    const sandbox = loadResolverSandbox({
      metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
      tdoc: { statusCode: 200, text: mixedHtml }
    });
    const result = sandbox.resolveMeetingById_(86178);
    check('inconsistent families: documents.family is null, not a silent first-pick', result.documents.family, null);
    check('inconsistent families: unresolved contains documents.family', result.unresolved.indexOf('documents.family') !== -1, true);
    check('inconsistent families: a warning is recorded', result.warnings.some(w => /Inconsistent TDoc families/.test(w)), true);
  }

  // Multiple agenda candidates end-to-end.
  {
    const twoAgendaHtml = '<html><body><table id="rgTdocList"><tbody>' +
      '<tr class="rgRow"><td></td><td align="center"><a href="https://x/S4aP260001.zip" target="_blank">S4aP260001</a></td><td align="center">agenda</td><td align="left">t</td><td align="center">s</td><td align="center">available</td><td align="center">f</td><td align="center">m</td><td><span title="topic" class="agendaItem">5.1</span></td><td><a></a></td><td></td><td></td></tr>' +
      '<tr class="rgAltRow"><td></td><td align="center"><a href="https://x/S4aP260002.zip" target="_blank">S4aP260002</a></td><td align="center">agenda</td><td align="left">t</td><td align="center">s</td><td align="center">available</td><td align="center">f</td><td align="center">m</td><td><span title="topic" class="agendaItem">5.1</span></td><td><a></a></td><td></td><td></td></tr>' +
      '</tbody></table></body></html>';
    const sandbox = loadResolverSandbox({
      metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
      tdoc: { statusCode: 200, text: twoAgendaHtml }
    });
    const result = sandbox.resolveMeetingById_(86178);
    check('multiple agenda candidates: agendaTdoc is null, not a silent pick', result.documents.agendaTdoc, null);
    check('multiple agenda candidates: agendaAmbiguous is true', result.documents.agendaAmbiguous, true);
    check('multiple agenda candidates: both listed', result.documents.agendaCandidates.sort(), ['S4aP260001', 'S4aP260002']);
    check('multiple agenda candidates: unresolved contains documents.agendaTdoc', result.unresolved.indexOf('documents.agendaTdoc') !== -1, true);
  }
}

// ============================================ 15. no PropertiesService ====

console.log('resolveMeetingById_() -- never touches PropertiesService (source-structure)');

{
  const fs2 = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs2.readFileSync(CODE_JS_PATH, 'utf8');
  const startMatch = source.match(/^function resolveMeetingById_\(/m);
  if (!startMatch) {
    failures++;
    console.log('  FAIL could not locate "function resolveMeetingById_(" in Code.js');
  } else {
    const startIndex = startMatch.index;
    const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnRe.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnRe.exec(source);
    const body = source.slice(startIndex, next ? next.index : source.length);
    check('resolveMeetingById_() never calls PropertiesService.setProperty/setProperties',
      /PropertiesService[\s\S]*?\.set(Property|Properties)/.test(body), false);
  }
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
