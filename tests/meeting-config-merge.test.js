/**
 * ARCH-010/ARCH-012 — Meeting-ID Resolve -> Preview -> Save configuration
 * merge/save logic regression coverage.
 *
 * Exercises the three ARCH-010 pure functions (computeMeetingDateFromStartDate_,
 * computeResolvedMeetingPreview_, computeMeetingConfigReadiness_) and the
 * end-to-end save workflow (saveConfigurationSettings()) against the
 * in-memory Document Properties store returned by loadCode(), reusing the
 * same real captured 3GPP Portal fixtures as tests/meeting-resolver.test.js
 * (via resolveMeetingById_() itself, stubbing only its three network
 * functions -- the same pattern used there). No live network is used.
 *
 * ARCH-012 extends this coverage to the ARCH-011 `sources.revisionsUrl`
 * evidence now flowing through the SAME preview/save pipeline (see section
 * 2f and section 4's new (e)/(f) cases below).
 *
 * Run: node tests/meeting-config-merge.test.js
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

// Same "override the three network functions on the sandbox" pattern as
// tests/meeting-resolver.test.js -- exercises the REAL resolveMeetingById_()
// orchestration, not a reimplementation. PROD-014: fetchRevisionsUrlCandidate_
// is deliberately stubbed to throw -- resolveMeetingById_() must never call
// it any more (revisions discovery is derive-only now), so any test that
// hits this stub is a regression back to the hanging behavior PROD-014 fixed.
function resolveWithFixtures(sandbox, meetingId, overrides) {
  const o = overrides || {};
  sandbox.fetchMeetingMetadataById_ = () => o.metadata !== undefined ? o.metadata : { statusCode: 200, text: '[]' };
  sandbox.fetchMeetingIcalById_ = () => o.ical !== undefined ? o.ical : { statusCode: 200, text: '' };
  sandbox.fetchMeetingTdocListById_ = () => o.tdoc !== undefined ? o.tdoc : { statusCode: 200, text: '' };
  sandbox.fetchRevisionsUrlCandidate_ = () => {
    throw new Error('fetchRevisionsUrlCandidate_ must not be called by resolveMeetingById_() (PROD-014)');
  };
  return sandbox.resolveMeetingById_(meetingId);
}

// ============================== 1. computeMeetingDateFromStartDate_() ======

console.log('computeMeetingDateFromStartDate_() -- calendar-date-only extraction');

{
  const { sandbox } = loadCode();
  const fn = sandbox.computeMeetingDateFromStartDate_;
  check('86178 real StartDate -> "September 22, 2026" (start date, NOT end date)',
    fn('2026-09-22 15:00:00'), 'September 22, 2026');
  check('date-only string (no time) also works', fn('2026-01-05'), 'January 5, 2026');
  check('null does not throw', fn(null), null);
  check('empty string does not throw', fn(''), null);
  check('malformed string does not throw', fn('not-a-date'), null);
}

// ============================== 2. computeResolvedMeetingPreview_() ========

console.log('computeResolvedMeetingPreview_() -- resolved > existing > unresolved merge, per field');

{
  const { sandbox } = loadCode();

  // (a) 86178, full resolve, with existing MAILING_LIST + TDOC_LIST_URL that
  // the resolver never touches -- both must survive into the preview as
  // "existing", not be blanked out.
  const existing86178 = {
    MEETING_ID: null, MEETING_TYPE: null, MEETING_NAME: null, MEETING_DATE: null,
    FTP_BASE: null, AGENDA_TDOC: null,
    MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED',
    TDOC_LIST_URL: 'https://www.3gpp.org/ftp/.../TDoc_List_Meeting_SA4-e (AH) on FS_6G_MED.xlsx'
  };
  const resolved86178 = resolveWithFixtures(sandbox, 86178, {
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    ical: { statusCode: 200, text: FIXTURES.ical86178 },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86178 }
  });
  const preview86178 = sandbox.computeResolvedMeetingPreview_(existing86178, resolved86178);

  check('86178: meetingId proposed, source resolved', preview86178.meetingId, { value: '86178', source: 'resolved' });
  check('86178: meetingType adhoc, source resolved', preview86178.meetingType, { value: 'adhoc', source: 'resolved' });
  check('86178: meetingName, source resolved', preview86178.meetingName, { value: 'SA4-e (AH) on FS_6G_MED', source: 'resolved' });
  check('86178: meetingDate derived from StartDate only, "September 22, 2026", source resolved',
    preview86178.meetingDate, { value: 'September 22, 2026', source: 'resolved' });
  check('86178: ftpBase ends /SA4_Plenary/Docs/, source resolved',
    /\/SA4_Plenary\/Docs\/$/.test(preview86178.ftpBase.value) && preview86178.ftpBase.source === 'resolved', true);
  check('86178: agendaTdoc S4aP260098, source resolved', preview86178.agendaTdoc, { value: 'S4aP260098', source: 'resolved' });
  check('86178: mailingList NEVER proposed by resolver -- existing value survives untouched',
    preview86178.mailingList, { value: '3GPP_TSG_SA4_FS_6G_MED', source: 'existing' });
  check('86178: raw start/end preserved untouched (the unusual Sept 22-26 range is NOT explained/collapsed)',
    preview86178.startDateRaw, '2026-09-22 15:00:00');
  check('86178: endDateRaw preserved, distinct from startDateRaw (26th, never used for MEETING_DATE)',
    preview86178.endDateRaw !== preview86178.startDateRaw, true);

  // (b) 86174: resolver finds no agenda TDoc. An existing manual AGENDA_TDOC
  // must NOT be erased in the preview -- it must survive as "existing".
  const existing86174 = { MEETING_ID: null, MEETING_TYPE: null, MEETING_NAME: null, MEETING_DATE: null, FTP_BASE: null, AGENDA_TDOC: 'S4aA260090-manual', MAILING_LIST: null, TDOC_LIST_URL: null };
  const resolved86174 = resolveWithFixtures(sandbox, 86174, {
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86174) },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86174 }
  });
  const preview86174 = sandbox.computeResolvedMeetingPreview_(existing86174, resolved86174);
  check('86174: resolver genuinely finds no agenda TDoc', resolved86174.documents.agendaTdoc, null);
  check('86174: preview.agendaTdoc falls back to the existing manual value, NOT blanked',
    preview86174.agendaTdoc, { value: 'S4aA260090-manual', source: 'existing' });

  // (c) Unknown Portal type code: must remain unresolved in the preview,
  // must NOT overwrite a valid existing MEETING_TYPE.
  const meetingUnknownType = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
  meetingUnknownType[0].Type = 'TS';
  const existingKnownType = { MEETING_ID: null, MEETING_TYPE: 'adhoc', MEETING_NAME: null, MEETING_DATE: null, FTP_BASE: null, AGENDA_TDOC: null, MAILING_LIST: null, TDOC_LIST_URL: null };
  const resolvedUnknownType = resolveWithFixtures(sandbox, 86178, {
    metadata: { statusCode: 200, text: JSON.stringify(meetingUnknownType) }
  });
  const previewUnknownType = sandbox.computeResolvedMeetingPreview_(existingKnownType, resolvedUnknownType);
  check('unknown Portal type: preview.meetingType falls back to the valid existing "adhoc", not overwritten with blank',
    previewUnknownType.meetingType, { value: 'adhoc', source: 'existing' });
  check('unknown Portal type: raw portalType code still surfaced for the UI hint', previewUnknownType.portalType, 'TS');

  // (d) Fully unresolved + no existing value -> "unresolved" for real.
  const previewBlank = sandbox.computeResolvedMeetingPreview_({}, null);
  check('no resolution, no existing value -> mailingList genuinely unresolved',
    previewBlank.mailingList, { value: '', source: 'unresolved' });
  check('no resolution, no existing value -> meetingType genuinely unresolved',
    previewBlank.meetingType, { value: '', source: 'unresolved' });

  // (e) Ambiguous agenda candidates: none pre-filled, candidates surfaced.
  const twoAgendaHtml = '<html><body><table id="rgTdocList"><tbody>' +
    '<tr class="rgRow"><td></td><td align="center"><a href="https://x/S4aP260001.zip" target="_blank">S4aP260001</a></td><td align="center">agenda</td><td align="left">t</td><td align="center">s</td><td align="center">available</td><td align="center">f</td><td align="center">m</td><td><span title="topic" class="agendaItem">5.1</span></td><td><a></a></td><td></td><td></td></tr>' +
    '<tr class="rgAltRow"><td></td><td align="center"><a href="https://x/S4aP260002.zip" target="_blank">S4aP260002</a></td><td align="center">agenda</td><td align="left">t</td><td align="center">s</td><td align="center">available</td><td align="center">f</td><td align="center">m</td><td><span title="topic" class="agendaItem">5.1</span></td><td><a></a></td><td></td><td></td></tr>' +
    '</tbody></table></body></html>';
  const resolvedAmbiguous = resolveWithFixtures(sandbox, 86178, {
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    tdoc: { statusCode: 200, text: twoAgendaHtml }
  });
  const previewAmbiguous = sandbox.computeResolvedMeetingPreview_({}, resolvedAmbiguous);
  check('ambiguous agenda candidates: preview.agendaTdoc is NOT pre-filled', previewAmbiguous.agendaTdoc.value, '');
  check('ambiguous agenda candidates: surfaced separately for explicit user choice',
    previewAmbiguous.agendaCandidates.sort(), ['S4aP260001', 'S4aP260002']);

  // (f) PROD-014/ARCH-012: revisionsUrl merge now has FOUR states.
  // resolveMeetingById_() itself never validates any more (PROD-014), so a
  // normal Resolve on 86178 only ever derives sources.revisionsUrlCandidate
  // -- sources.revisionsUrl stays null. computeResolvedMeetingPreview_()
  // must therefore label that candidate "candidate", never "resolved".
  const resolved86178Revisions = resolveWithFixtures(sandbox, 86178, {
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    ical: { statusCode: 200, text: FIXTURES.ical86178 },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86178 }
  });
  check('86178: resolveMeetingById_() itself never validates revisionsUrl any more (PROD-014)',
    resolved86178Revisions.sources.revisionsUrl, null);
  check('86178: resolveMeetingById_() derives revisionsUrlCandidate',
    resolved86178Revisions.sources.revisionsUrlCandidate, 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/');

  const previewCandidateNoExisting = sandbox.computeResolvedMeetingPreview_({}, resolved86178Revisions);
  check('86178, no existing REVISIONS_URL: preview.revisionsUrl shows the candidate, labeled "candidate" (NOT "resolved")',
    previewCandidateNoExisting.revisionsUrl,
    { value: 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/', source: 'candidate' });

  const previewCandidateWithExisting = sandbox.computeResolvedMeetingPreview_(
    { REVISIONS_URL: 'https://old-manual-value/Inbox/Drafts/Plenary' },
    resolved86178Revisions
  );
  check('86178, existing REVISIONS_URL present: an unvalidated candidate never overrides it -- existing wins, value NOT blanked',
    previewCandidateWithExisting.revisionsUrl,
    { value: 'https://old-manual-value/Inbox/Drafts/Plenary', source: 'existing' });

  // A genuinely validated value (e.g. from a future explicit
  // validateRevisionsUrlCandidate_() action merged into resolverResult.sources)
  // must still win over both an existing value and a mere candidate --
  // computeResolvedMeetingPreview_() doesn't care how sources.revisionsUrl
  // became non-null, only that it did.
  const resolvedWithValidatedRevisions = JSON.parse(JSON.stringify(resolved86178Revisions));
  resolvedWithValidatedRevisions.sources.revisionsUrl = 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/';
  const previewValidated = sandbox.computeResolvedMeetingPreview_(
    { REVISIONS_URL: 'https://old-manual-value/Inbox/Drafts/Plenary' },
    resolvedWithValidatedRevisions
  );
  check('a genuinely validated sources.revisionsUrl is labeled "resolved" and wins over an existing value',
    previewValidated.revisionsUrl,
    { value: 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/', source: 'resolved' });

  const previewNoResolveNoExisting = sandbox.computeResolvedMeetingPreview_({}, null);
  check('no resolution, no existing REVISIONS_URL -> genuinely unresolved (optional field, never guessed)',
    previewNoResolveNoExisting.revisionsUrl, { value: '', source: 'unresolved' });
}

// ============================== 3. computeMeetingConfigReadiness_() ========

console.log('computeMeetingConfigReadiness_() -- only fields the production build actually depends on');

{
  const { sandbox } = loadCode();
  const fn = sandbox.computeMeetingConfigReadiness_;

  check('fully configured ad-hoc meeting -> ready',
    fn({ MEETING_TYPE: 'adhoc', TDOC_LIST_URL: 'x', AGENDA_TDOC: 'S4aP260098', MAILING_LIST: 'list' }).ready, true);

  const adhocMissing = fn({ MEETING_TYPE: 'adhoc', TDOC_LIST_URL: '', AGENDA_TDOC: '', MAILING_LIST: '' });
  check('ad-hoc meeting missing everything -> not ready', adhocMissing.ready, false);
  check('ad-hoc meeting missing everything -> 3 issues (TDoc List URL, Agenda TDoc, Mailing list)', adhocMissing.issues.length, 3);

  const mainMissingTdocAndAgenda = fn({ MEETING_TYPE: 'main', TDOC_LIST_URL: '', AGENDA_TDOC: '', MAILING_LIST: 'list' });
  check('main meeting missing TDOC_LIST_URL/AGENDA_TDOC -> still ready (working fallback formulas exist for main meetings)',
    mainMissingTdocAndAgenda.ready, true);

  const mainMissingMailingList = fn({ MEETING_TYPE: 'main', TDOC_LIST_URL: 'x', AGENDA_TDOC: 'x', MAILING_LIST: '' });
  check('main meeting missing MAILING_LIST -> not ready (mailing list always required, resolver can never supply it)',
    mainMissingMailingList.ready, false);
  check('main meeting missing MAILING_LIST -> exactly 1 issue', mainMissingMailingList.issues.length, 1);

  check('optional Revisions URL is never checked (not a readiness field at all)',
    fn({ MEETING_TYPE: 'adhoc', TDOC_LIST_URL: 'x', AGENDA_TDOC: 'x', MAILING_LIST: 'x', REVISIONS_URL: '' }).ready, true);
  check('ARCH-012: missing REVISIONS_URL alone does not add an issue, ready meeting stays ready',
    fn({ MEETING_TYPE: 'adhoc', TDOC_LIST_URL: 'x', AGENDA_TDOC: 'x', MAILING_LIST: 'x', REVISIONS_URL: '' }).issues.length, 0);
  check('ARCH-012: missing REVISIONS_URL on an otherwise-unready meeting does not add a 4th issue',
    fn({ MEETING_TYPE: 'adhoc', TDOC_LIST_URL: '', AGENDA_TDOC: '', MAILING_LIST: '', REVISIONS_URL: '' }).issues.length, 3);
}

// ============================== 4. saveConfigurationSettings() workflow ====

console.log('saveConfigurationSettings() -- merge-on-save semantics, skip-if-blank protection');

{
  // (a) Manual override wins over resolver proposal on Save: the dialog
  // always submits whatever is currently in the input (resolver-proposed OR
  // user-edited) -- saveConfigurationSettings() has no way to distinguish
  // them and must not need to: the LAST value in the field, submitted at
  // Save time, wins. Simulate a user who edited the resolver-proposed
  // meetingName before saving.
  {
    const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_NAME: 'Old Existing Name' } });
    sandbox.saveConfigurationSettings({
      meetingFolder: 'F', meetingNumber: '1', meetingId: '86178',
      meetingType: 'adhoc',
      meetingName: 'User-Edited Meeting Name', // user changed the resolver's proposal before Save
      meetingDate: 'September 22, 2026',
      ftpBase: 'https://example/Docs/',
      mailingList: '3GPP_TSG_SA4_FS_6G_MED',
      reportType: '6G', agendaSourceDocId: '', agendaTdoc: 'S4aP260098', tdocUrl: '', showPreview: true, apiToken: ''
    });
    check('manual override: user-edited MEETING_NAME wins over old existing value',
      docProps._store.MEETING_NAME, 'User-Edited Meeting Name');
  }

  // (b) 86174-style save: submitted agendaTdoc is blank (resolver found
  // none) -- the existing manual AGENDA_TDOC must survive the Save
  // untouched, not be erased to blank.
  {
    const { sandbox, docProps } = loadCode({ documentProperties: { AGENDA_TDOC: 'S4aA260090-manual', MAILING_LIST: 'existing-list', TDOC_LIST_URL: 'existing-url' } });
    sandbox.saveConfigurationSettings({
      meetingFolder: 'F', meetingNumber: '1', meetingId: '86174',
      meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG', meetingDate: 'June 1, 2026',
      ftpBase: 'https://example/Docs/',
      agendaTdoc: '', // blank -- resolver found nothing for 86174
      mailingList: '', // blank in the submitted form too
      reportType: '6G', agendaSourceDocId: '', tdocUrl: '', showPreview: true, apiToken: ''
    });
    check('86174 save with blank agendaTdoc: existing manual AGENDA_TDOC survives, NOT erased',
      docProps._store.AGENDA_TDOC, 'S4aA260090-manual');
    check('86174 save with blank mailingList: existing MAILING_LIST survives, NOT erased',
      docProps._store.MAILING_LIST, 'existing-list');
  }

  // (c) 86178-style full save: MEETING_ID/TYPE/NAME/DATE/FTP_BASE/AGENDA_TDOC
  // all written from resolver-proposed values; pre-existing MAILING_LIST and
  // TDOC_LIST_URL (never touched by the resolver, and left blank in this
  // submission because the dialog carried their existing values forward
  // unedited into the same fields) remain unchanged when the submitted
  // value equals what was already there, and are NOT erased even when the
  // caller passes through an unrelated blank tdocUrl override path.
  {
    const { sandbox, docProps } = loadCode({ documentProperties: { MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED', TDOC_LIST_URL: 'existing-tdoc-url.xlsx' } });
    sandbox.saveConfigurationSettings({
      meetingFolder: 'SA4_Plenary', meetingNumber: '86178', meetingId: '86178',
      meetingType: 'adhoc', meetingName: 'SA4-e (AH) on FS_6G_MED', meetingDate: 'September 22, 2026',
      ftpBase: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/',
      agendaTdoc: 'S4aP260098',
      mailingList: '3GPP_TSG_SA4_FS_6G_MED', // carried forward unedited by the dialog
      reportType: '6G', agendaSourceDocId: '', tdocUrl: 'existing-tdoc-url.xlsx', showPreview: true, apiToken: ''
    });
    check('86178 save: MEETING_ID written', docProps._store.MEETING_ID, '86178');
    check('86178 save: MEETING_TYPE written', docProps._store.MEETING_TYPE, 'adhoc');
    check('86178 save: MEETING_NAME written', docProps._store.MEETING_NAME, 'SA4-e (AH) on FS_6G_MED');
    check('86178 save: MEETING_DATE written (start date, "September 22, 2026")', docProps._store.MEETING_DATE, 'September 22, 2026');
    check('86178 save: FTP_BASE written', docProps._store.FTP_BASE, 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/');
    check('86178 save: AGENDA_TDOC written', docProps._store.AGENDA_TDOC, 'S4aP260098');
    check('86178 save: MAILING_LIST survives unchanged', docProps._store.MAILING_LIST, '3GPP_TSG_SA4_FS_6G_MED');
    check('86178 save: TDOC_LIST_URL survives unchanged', docProps._store.TDOC_LIST_URL, 'existing-tdoc-url.xlsx');
  }

  // (d) Legacy document: no MEETING_ID/MEETING_TYPE/etc. ever set, only the
  // old fields -- must remain fully editable/savable without ever going
  // through Resolve.
  {
    const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136', REPORT_SUFFIX: '6G' } });
    sandbox.saveConfigurationSettings({
      meetingFolder: 'TSGS4_137_Xian', meetingNumber: '137', meetingId: '60777',
      meetingType: '', meetingName: '', meetingDate: '', ftpBase: '', agendaTdoc: '', mailingList: '',
      reportType: '6G', agendaSourceDocId: '', tdocUrl: '', showPreview: true, apiToken: ''
    });
    check('legacy save: MEETING_FOLDER updated via the old field path', docProps._store.MEETING_FOLDER, 'TSGS4_137_Xian');
    check('legacy save: MEETING_NUMBER updated via the old field path', docProps._store.MEETING_NUMBER, '137');
    check('legacy save: no MEETING_TYPE written when blank (new ARCH-010 field never forced)', docProps._store.MEETING_TYPE, undefined);
    check('legacy save: no MEETING_NAME written when blank', docProps._store.MEETING_NAME, undefined);
    check('ARCH-012: legacy save with a blank revisionsUrl writes nothing for REVISIONS_URL', docProps._store.REVISIONS_URL, undefined);
  }

  // (e) ARCH-012: user-edited (or resolver-proposed, carried through
  // unedited) revisionsUrl wins over whatever was there before on Save.
  {
    const { sandbox, docProps } = loadCode({ documentProperties: { REVISIONS_URL: 'https://old-manual-value/Inbox/Drafts/Plenary' } });
    sandbox.saveConfigurationSettings({
      meetingFolder: 'F', meetingNumber: '1', meetingId: '86178',
      meetingType: 'adhoc', meetingName: 'SA4-e (AH) on FS_6G_MED', meetingDate: 'September 22, 2026',
      ftpBase: 'https://example/Docs/', agendaTdoc: 'S4aP260098', mailingList: '3GPP_TSG_SA4_FS_6G_MED',
      revisionsUrl: 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/', // user-reviewed/edited value submitted at Save time
      reportType: '6G', agendaSourceDocId: '', tdocUrl: '', showPreview: true, apiToken: ''
    });
    check('ARCH-012: user-edited/reviewed REVISIONS_URL wins over the old value on Save',
      docProps._store.REVISIONS_URL, 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/');
  }

  // (f) ARCH-012: a blank submitted revisionsUrl (resolver found nothing/
  // failed validation, user did not type one in) must NOT erase an
  // existing manually configured REVISIONS_URL -- the same skip-if-blank
  // protection every other ARCH-010 field already has.
  {
    const { sandbox, docProps } = loadCode({ documentProperties: { REVISIONS_URL: 'https://old-manual-value/Inbox/Drafts/Plenary' } });
    sandbox.saveConfigurationSettings({
      meetingFolder: 'F', meetingNumber: '1', meetingId: '86178',
      meetingType: 'adhoc', meetingName: 'SA4-e (AH) on FS_6G_MED', meetingDate: 'September 22, 2026',
      ftpBase: 'https://example/Docs/', agendaTdoc: 'S4aP260098', mailingList: '3GPP_TSG_SA4_FS_6G_MED',
      revisionsUrl: '', // blank -- resolver 403'd or was never run
      reportType: '6G', agendaSourceDocId: '', tdocUrl: '', showPreview: true, apiToken: ''
    });
    check('ARCH-012: blank submitted REVISIONS_URL does NOT erase the existing manual value',
      docProps._store.REVISIONS_URL, 'https://old-manual-value/Inbox/Drafts/Plenary');
  }
}

// ================== 5. Resolve without Save never mutates PropertiesService

console.log('resolveMeetingForConfigDialog_() -- never writes to PropertiesService (source-structure)');

{
  const source = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
  const startMatch = source.match(/^function resolveMeetingForConfigDialog_\(/m);
  check('resolveMeetingForConfigDialog_() exists', !!startMatch, true);
  if (startMatch) {
    const startIndex = startMatch.index;
    const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnRe.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnRe.exec(source);
    const body = source.slice(startIndex, next ? next.index : source.length);
    check('resolveMeetingForConfigDialog_() never calls PropertiesService.setProperty/setProperties',
      /PropertiesService[\s\S]*?\.set(Property|Properties)/.test(body), false);
    check('resolveMeetingForConfigDialog_() never calls deleteProperty either',
      /\.deleteProperty/.test(body), false);
  }

  // Behavioral confirmation: calling it directly against a live sandbox must
  // leave Document Properties byte-identical -- including now that a
  // normal Resolve derives a revisions/drafts candidate (PROD-014: never
  // fetched/validated synchronously) AND never touches TdocList.aspx
  // (PROD-016: resolveMeetingForConfigDialog_() now calls
  // resolveMeetingCoreById_(), not the full resolveMeetingById_()). Both
  // stubs throw if ever reached, which would be a regression back to the
  // hanging behavior these tasks fixed.
  const { sandbox, docProps } = loadCode({ documentProperties: { MAILING_LIST: 'untouched-list', MEETING_ID: '60777' } });
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  sandbox.fetchMeetingIcalById_ = () => ({ statusCode: 200, text: FIXTURES.ical86178 });
  sandbox.fetchMeetingTdocListById_ = () => {
    throw new Error('fetchMeetingTdocListById_ must not be called during a normal Resolve (PROD-016)');
  };
  sandbox.fetchRevisionsUrlCandidate_ = () => {
    throw new Error('fetchRevisionsUrlCandidate_ must not be called during a normal Resolve (PROD-014)');
  };
  const before = JSON.stringify(docProps._store);
  const result = sandbox.resolveMeetingForConfigDialog_('86178');
  const after = JSON.stringify(docProps._store);
  check('resolveMeetingForConfigDialog_() returns ok:true for a valid, resolvable id', result.ok, true);
  check('PROD-016: core-only resolve leaves preview.agendaTdoc genuinely unresolved (not yet discovered, no existing value here)',
    result.preview.agendaTdoc, { value: '', source: 'unresolved' });
  check('PROD-014: the resolve derived a revisionsUrl CANDIDATE (not validated) -- provenance is "candidate", never "resolved"',
    result.preview.revisionsUrl.source, 'candidate');
  check('resolveMeetingForConfigDialog_() does not mutate Document Properties, even on a successful resolve (candidate derivation included)', after, before);
}

// ============== 6. discoverAgendaForConfigDialog_() (PROD-016) ============

console.log('discoverAgendaForConfigDialog_() -- explicit, separate TDoc/agenda enrichment');

{
  // (a) requires a matching core result -- refuses, without ever touching
  // TdocList.aspx, if none is supplied.
  const { sandbox, docProps } = loadCode();
  sandbox.fetchMeetingTdocListById_ = () => {
    throw new Error('fetchMeetingTdocListById_ must not be called when no core result was supplied');
  };
  const before = JSON.stringify(docProps._store);
  const noCoreResult = sandbox.discoverAgendaForConfigDialog_('86178', null);
  const after = JSON.stringify(docProps._store);
  check('no core result supplied: ok:false', noCoreResult.ok, false);
  check('no core result supplied: a clear error message', typeof noCoreResult.error, 'string');
  check('no core result supplied: Document Properties untouched', after, before);
}

{
  // (b) mismatched meeting id between the input and the supplied core
  // result -- refuses rather than merging mismatched data.
  const { sandbox } = loadCode();
  sandbox.fetchMeetingTdocListById_ = () => {
    throw new Error('fetchMeetingTdocListById_ must not be called on a meetingId mismatch');
  };
  const wrongCore = { id: 60777, meeting: { name: 'Some Other Meeting' }, sources: {}, documents: null, unresolved: [], warnings: [], raw: {} };
  const mismatchResult = sandbox.discoverAgendaForConfigDialog_('86178', wrongCore);
  check('meetingId mismatch: ok:false', mismatchResult.ok, false);
}

{
  // (c) real, successful enrichment: full round trip Resolve -> Discover.
  const { sandbox, docProps } = loadCode({ documentProperties: { MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED', TDOC_LIST_URL: 'existing-tdoc-url.xlsx' } });
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: FIXTURES.tdocListHtml86178 });

  const coreResult = sandbox.resolveMeetingForConfigDialog_('86178');
  check('(c) core resolve succeeds first', coreResult.ok, true);
  check('(c) core resolve leaves agendaTdoc unresolved (not yet discovered)', coreResult.preview.agendaTdoc.source, 'unresolved');

  const before = JSON.stringify(docProps._store);
  const enrichResult = sandbox.discoverAgendaForConfigDialog_('86178', coreResult.resolved);
  const after = JSON.stringify(docProps._store);
  check('(c) discovery succeeds', enrichResult.ok, true);
  check('(c) discovery finds S4aP260098, now "resolved"', enrichResult.preview.agendaTdoc, { value: 'S4aP260098', source: 'resolved' });
  check('(c) discovery does not mutate Document Properties either', after, before);
  check('(c) discovery does not disturb fields core resolution already established (meetingName)',
    enrichResult.preview.meetingName, { value: 'SA4-e (AH) on FS_6G_MED', source: 'resolved' });
  check('(c) existing MAILING_LIST/TDOC_LIST_URL still survive after discovery',
    [enrichResult.preview.mailingList.value, enrichResult.preview.mailingList.source],
    ['3GPP_TSG_SA4_FS_6G_MED', 'existing']);
}

{
  // (d) enrichment FAILURE cannot erase/disturb what core resolution (or
  // an existing manual AGENDA_TDOC) already established.
  const { sandbox } = loadCode({ documentProperties: { AGENDA_TDOC: 'S4aA260090-manual' } });
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  const coreResult = sandbox.resolveMeetingForConfigDialog_('86178');
  check('(d) core resolve preserves the existing manual AGENDA_TDOC', coreResult.preview.agendaTdoc, { value: 'S4aA260090-manual', source: 'existing' });

  sandbox.fetchMeetingTdocListById_ = () => { throw new Error('simulated TdocList outage'); };
  const enrichResult = sandbox.discoverAgendaForConfigDialog_('86178', coreResult.resolved);
  check('(d) discovery reports ok:true even though the underlying fetch failed (degrades to a warning, matches existing resolver convention)', enrichResult.ok, true);
  check('(d) the existing manual AGENDA_TDOC still survives a failed discovery attempt',
    enrichResult.preview.agendaTdoc, { value: 'S4aA260090-manual', source: 'existing' });
  check('(d) a warning records the discovery failure', enrichResult.resolved.warnings.some(w => /simulated TdocList outage/.test(w)), true);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
