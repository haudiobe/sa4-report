/**
 * SA4-ARCH-003 / SA4-IMPL-002 — MeetingContext compatibility-layer tests.
 *
 * getMeetingContext_() is a pure, read-only normalization of whatever
 * getReportConfig_() already returns -- see the "MEETING CONTEXT" comment
 * block directly above it in Code.js for the full rationale (why
 * options.emailStartDate/fetchAbstractsOnUpdate are omitted, why
 * meeting.type is hardcoded to 'main', why report.structureProfile exists).
 *
 * These tests deliberately do NOT duplicate a second, hand-maintained set of
 * "expected" constants. Every relationship checked here compares
 * getMeetingContext_()'s output against the SAME live getReportConfig_()
 * call in the SAME sandbox instance, so a future change to
 * getReportConfig_()'s defaults (e.g. a new meeting's FTP_BASE) cannot make
 * this suite drift out of sync with reality the way a hardcoded fixture
 * could. The one place fixed values do appear is the report.structureProfile
 * classification, because that classification does not exist anywhere in
 * getReportConfig_()'s output -- it mirrors the is6G/isSWGReport branch read
 * directly from buildSkeletonWithTdocTables() (Code.js ~5350-5459).
 *
 * Also verified: getMeetingContext_() has no side effects beyond what
 * getReportConfig_() itself performs (a PropertiesService read) -- calling
 * it never writes to the property stores used by the sandbox.
 *
 * SA4-IMPL-002 adds direct tests of the pure resolveMeetingSources_() helper
 * that now produces context.sources -- see the "resolveMeetingSources_"
 * section below. These are pure-function tests with no PropertiesService
 * involvement at all (not even the sandbox's stub), proving the resolver
 * itself has no dependency on Apps Script or on getReportConfig_() -- it
 * will work identically once a real ad-hoc meeting hands it explicit
 * overrides and a numberless `derived` object, long before any production
 * code actually does that (see the SA4-IMPL-002 report for the remaining
 * gap: no DocumentProperty mechanism yet supplies those overrides).
 *
 * Run: node tests/meeting-context.test.js
 */

const { loadCode } = require('./helpers/load-code.js');

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

const REPORT_TYPES = ['6G', 'Audio', 'Video', 'MBS', 'RTC', 'Liaison', 'New'];
const SWG_REPORT_TYPES = ['Audio', 'Video', 'MBS', 'RTC'];

function expectedStructureProfile(reportType) {
  if (reportType === '6G') return 'main-6g';
  if (SWG_REPORT_TYPES.indexOf(reportType) !== -1) return 'main-swg';
  return 'main-other';
}

// --------------------------------------- relationships to getReportConfig_()

console.log('getMeetingContext_() vs. getReportConfig_() -- same live config, both directions');

REPORT_TYPES.forEach(type => {
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: type } });
  const cfg = sandbox.getReportConfig_();
  const ctx = sandbox.getMeetingContext_();

  console.log(`  -- report type: ${type} --`);

  check(`${type}: group is the constant "SA4"`, ctx.group, 'SA4');

  check(`${type}: meeting.type === 'main'`, ctx.meeting.type, 'main');
  check(`${type}: meeting.folder === config.MEETING_FOLDER`, ctx.meeting.folder, cfg.MEETING_FOLDER);
  check(`${type}: meeting.number === config.MEETING_NUMBER`, ctx.meeting.number, cfg.MEETING_NUMBER);
  check(`${type}: meeting.portalId === config.MEETING_ID`, ctx.meeting.portalId, cfg.MEETING_ID);

  check(`${type}: report.type === config.REPORT_SUFFIX`, ctx.report.type, cfg.REPORT_SUFFIX);
  check(`${type}: report.agendaPrefix === config.AGENDA_ITEM_PREFIX`, ctx.report.agendaPrefix, cfg.AGENDA_ITEM_PREFIX);
  check(`${type}: report.structureProfile matches the 6G/SWG/other classification`,
    ctx.report.structureProfile, expectedStructureProfile(type));

  check(`${type}: sources.ftpBase === config.FTP_BASE`, ctx.sources.ftpBase, cfg.FTP_BASE);
  check(`${type}: sources.tdocListUrl === config.TDOC_LIST_URL`, ctx.sources.tdocListUrl, cfg.TDOC_LIST_URL);
  check(`${type}: sources.agendaTdoc === config.AGENDA_TDOC`, ctx.sources.agendaTdoc, cfg.AGENDA_TDOC);
  check(`${type}: sources.agendaTemplateDocId === config.AGENDA_SOURCE_DOC_ID`,
    ctx.sources.agendaTemplateDocId, cfg.AGENDA_SOURCE_DOC_ID);
  check(`${type}: sources.mailingList === config.LIST_NAME`, ctx.sources.mailingList, cfg.LIST_NAME);
  check(`${type}: sources.draftsFolder === config.DRAFTS_FOLDER`, ctx.sources.draftsFolder, cfg.DRAFTS_FOLDER);
  check(`${type}: sources.revisionsUrl === config.REVISIONS_URL`, ctx.sources.revisionsUrl, cfg.REVISIONS_URL);

  check(`${type}: options.showPreviewSnippet === config.SHOW_PREVIEW_SNIPPET`,
    ctx.options.showPreviewSnippet, cfg.SHOW_PREVIEW_SNIPPET);
});

// ------------------------------------------------- omitted fields, documented

console.log('fields intentionally omitted (not derivable from getReportConfig_() without a second config source)');

{
  const { sandbox } = loadCode();
  const ctx = sandbox.getMeetingContext_();
  check('options.emailStartDate is not present (getReportConfig_() does not expose EMAIL_START_DATE)',
    Object.prototype.hasOwnProperty.call(ctx.options, 'emailStartDate'), false);
  check('options.fetchAbstractsOnUpdate is not present (only getFetchAbstractsSetting_() reads that property, a separate path)',
    Object.prototype.hasOwnProperty.call(ctx.options, 'fetchAbstractsOnUpdate'), false);
}

// --------------------------------------------------- non-default configuration

console.log('non-default configuration (explicit DocumentProperties, not just REPORT_SUFFIX)');

{
  const { sandbox } = loadCode({
    documentProperties: {
      MEETING_FOLDER: 'TSGS4_140_Example',
      MEETING_NUMBER: '140',
      MEETING_ID: '61111',
      REPORT_SUFFIX: 'MBS',
      AGENDA_TDOC: 'S4-261234',
      AGENDA_SOURCE_DOC_ID: 'abc123DocId',
      SHOW_PREVIEW_SNIPPET: 'false'
    }
  });
  const cfg = sandbox.getReportConfig_();
  const ctx = sandbox.getMeetingContext_();

  check('meeting.folder reflects a saved (non-default) MEETING_FOLDER', ctx.meeting.folder, 'TSGS4_140_Example');
  check('meeting.folder equals config.MEETING_FOLDER for the same saved value', ctx.meeting.folder, cfg.MEETING_FOLDER);
  check('sources.agendaTdoc reflects a saved (non-default) AGENDA_TDOC', ctx.sources.agendaTdoc, 'S4-261234');
  check('sources.agendaTemplateDocId reflects a saved (non-default) AGENDA_SOURCE_DOC_ID', ctx.sources.agendaTemplateDocId, 'abc123DocId');
  check('report.structureProfile is main-swg for MBS regardless of other saved values', ctx.report.structureProfile, 'main-swg');
  check('options.showPreviewSnippet reflects a saved "false" as boolean false', ctx.options.showPreviewSnippet, false);
}

// ------------------------------------------------------------------ purity

console.log('purity: getMeetingContext_() has no side effects beyond getReportConfig_() itself');

{
  const { sandbox, docProps, scriptProps } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Audio' } });

  const docKeysBefore = JSON.stringify(docProps.getKeys().slice().sort());
  const scriptKeysBefore = JSON.stringify(scriptProps.getKeys().slice().sort());

  const first = sandbox.getMeetingContext_();
  const second = sandbox.getMeetingContext_();
  const third = sandbox.getMeetingContext_();

  const docKeysAfter = JSON.stringify(docProps.getKeys().slice().sort());
  const scriptKeysAfter = JSON.stringify(scriptProps.getKeys().slice().sort());

  check('calling it 3x does not add/remove DocumentProperties keys', docKeysAfter, docKeysBefore);
  check('calling it 3x does not add/remove ScriptProperties keys', scriptKeysAfter, scriptKeysBefore);
  check('repeated calls return deeply equal objects', JSON.stringify(second), JSON.stringify(first));
  check('repeated calls return deeply equal objects (3rd call too)', JSON.stringify(third), JSON.stringify(first));
  check('repeated calls return distinct object instances (no shared mutable state)', first === second, false);
}

// ============================== SA4-IMPL-002: resolveMeetingSources_() =====

console.log('resolveMeetingSources_ -- pure precedence rule, no PropertiesService involved');

const REALISTIC_MAIN_DERIVED = {
  ftpBase: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/',
  tdocListUrl: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4%23136.xlsx',
  agendaTdoc: '',
  agendaTemplateDocId: '1qP--dusvUhNwwBtMEH4xVdxaP1c6L1hZ49geICoYV2s',
  mailingList: '3GPP_TSG_SA_WG4',
  draftsFolder: 'FS_6G_MED',
  revisionsUrl: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED'
};

{
  const { sandbox } = loadCode();

  check('no overrides -- resolved sources equal derived config exactly',
    sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, {}),
    REALISTIC_MAIN_DERIVED);

  check('no overrides object at all (undefined) -- same result as {}',
    sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, undefined),
    REALISTIC_MAIN_DERIVED);
}

{
  const { sandbox } = loadCode();
  const result = sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, { tdocListUrl: 'https://example.com/custom-list.xlsx' });

  check('one override (tdocListUrl) wins', result.tdocListUrl, 'https://example.com/custom-list.xlsx');
  check('  -> ftpBase unaffected', result.ftpBase, REALISTIC_MAIN_DERIVED.ftpBase);
  check('  -> agendaTdoc unaffected', result.agendaTdoc, REALISTIC_MAIN_DERIVED.agendaTdoc);
  check('  -> agendaTemplateDocId unaffected', result.agendaTemplateDocId, REALISTIC_MAIN_DERIVED.agendaTemplateDocId);
  check('  -> mailingList unaffected', result.mailingList, REALISTIC_MAIN_DERIVED.mailingList);
  check('  -> draftsFolder unaffected', result.draftsFolder, REALISTIC_MAIN_DERIVED.draftsFolder);
  check('  -> revisionsUrl unaffected', result.revisionsUrl, REALISTIC_MAIN_DERIVED.revisionsUrl);
}

{
  // Realistic ad-hoc-shaped override values (real URLs/identifiers from
  // SA4-ARCH-005/006 evidence) -- test-only; no production default changes.
  const { sandbox } = loadCode();
  const adhocOverrides = {
    ftpBase: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/',
    tdocListUrl: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/TDoc_List_Meeting_SA4-(AH) Audio SWG on ULBC-MED.xlsx',
    agendaTdoc: 'S4aA260090',
    revisionsUrl: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/'
  };
  const result = sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, adhocOverrides);

  check('multiple overrides: ftpBase wins independently', result.ftpBase, adhocOverrides.ftpBase);
  check('multiple overrides: tdocListUrl wins independently', result.tdocListUrl, adhocOverrides.tdocListUrl);
  check('multiple overrides: agendaTdoc wins independently', result.agendaTdoc, adhocOverrides.agendaTdoc);
  check('multiple overrides: revisionsUrl wins independently', result.revisionsUrl, adhocOverrides.revisionsUrl);
  check('multiple overrides: agendaTemplateDocId NOT overridden -- falls back to derived',
    result.agendaTemplateDocId, REALISTIC_MAIN_DERIVED.agendaTemplateDocId);
  check('multiple overrides: mailingList NOT overridden -- falls back to derived',
    result.mailingList, REALISTIC_MAIN_DERIVED.mailingList);
  check('multiple overrides: draftsFolder NOT overridden -- falls back to derived',
    result.draftsFolder, REALISTIC_MAIN_DERIVED.draftsFolder);
}

{
  const { sandbox } = loadCode();
  console.log('empty/blank/nullish override values fall back to derived (do not count as "explicit")');

  check('empty string override falls back',
    sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, { tdocListUrl: '' }).tdocListUrl,
    REALISTIC_MAIN_DERIVED.tdocListUrl);
  check('whitespace-only override falls back',
    sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, { ftpBase: '   ' }).ftpBase,
    REALISTIC_MAIN_DERIVED.ftpBase);
  check('null override falls back',
    sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, { agendaTdoc: null }).agendaTdoc,
    REALISTIC_MAIN_DERIVED.agendaTdoc);
  check('undefined override falls back',
    sandbox.resolveMeetingSources_(REALISTIC_MAIN_DERIVED, { agendaTemplateDocId: undefined }).agendaTemplateDocId,
    REALISTIC_MAIN_DERIVED.agendaTemplateDocId);
}

{
  // SA4-IMPL-002 §6: prove the RESOLVER itself needs no MEETING_NUMBER (or
  // any derived value at all) as long as every field it needs comes from
  // `overrides`. This does NOT mean getReportConfig_() or any production
  // workflow supports a numberless meeting yet -- only that this pure
  // building block already can, ahead of that wiring existing.
  const { sandbox } = loadCode();
  console.log('numberless derived config -- resolver succeeds purely from explicit overrides');

  const numberlessDerived = {
    ftpBase: undefined, tdocListUrl: undefined, agendaTdoc: '',
    agendaTemplateDocId: undefined, mailingList: undefined,
    draftsFolder: undefined, revisionsUrl: undefined
  };
  const fullAdhocOverrides = {
    ftpBase: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/',
    tdocListUrl: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/TDoc_List_Meeting_SA4-(AH) Audio SWG on ULBC-MED.xlsx',
    agendaTdoc: 'S4aA260090',
    revisionsUrl: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/'
  };
  const result = sandbox.resolveMeetingSources_(numberlessDerived, fullAdhocOverrides);

  check('numberless: ftpBase resolved purely from override', result.ftpBase, fullAdhocOverrides.ftpBase);
  check('numberless: tdocListUrl resolved purely from override', result.tdocListUrl, fullAdhocOverrides.tdocListUrl);
  check('numberless: agendaTdoc resolved purely from override', result.agendaTdoc, fullAdhocOverrides.agendaTdoc);
  check('numberless: revisionsUrl resolved purely from override', result.revisionsUrl, fullAdhocOverrides.revisionsUrl);
  check('numberless: a field with NEITHER a derived value NOR an override is left undefined (nothing invented)',
    result.agendaTemplateDocId, undefined);
  check('numberless: mailingList likewise left undefined (no property mechanism supplies it yet)',
    result.mailingList, undefined);
  check('numberless: draftsFolder likewise left undefined',
    result.draftsFolder, undefined);
}

// ---------------------------- getMeetingContext_() production wiring proof

console.log('getMeetingContext_() production wiring: resolver is used, with no overrides supplied today');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  // ADDON-002: getMeetingContext_() now takes an optional `context`
  // parameter (source-structure only -- no production caller passes one),
  // so this locator tolerates an optional argument name instead of
  // requiring an exactly-empty `()`.
  const startMatch = source.match(/^function getMeetingContext_\([A-Za-z0-9_$]*\)/m);
  if (!startMatch) {
    failures++;
    console.log('  FAIL could not locate "function getMeetingContext_(...)" in Code.js');
  } else {
    const startIndex = startMatch.index;
    const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnRe.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnRe.exec(source);
    const endIndex = next ? next.index : source.length;
    const body = source.slice(startIndex, endIndex);

    check('getMeetingContext_() body calls resolveMeetingSources_()',
      /\bresolveMeetingSources_\s*\(/.test(body), true);
  }
}

// ============================ SA4-IMPL-003: ad-hoc meeting identity =======

console.log('normalizeMeetingType_ -- backward-compatible fallback for blank/absent, strict for everything else (SA4-IMPL-003A)');

function expectThrows(name, fn, messageSubstring) {
  try {
    fn();
    failures++;
    console.log(`  FAIL ${name}\n         expected a throw, but none occurred`);
  } catch (e) {
    if (messageSubstring && e.message.indexOf(messageSubstring) === -1) {
      failures++;
      console.log(`  FAIL ${name}\n         expected message to contain ${JSON.stringify(messageSubstring)}\n         actual message: ${e.message}`);
    } else {
      console.log(`  ok   ${name}`);
    }
  }
}

{
  const { sandbox } = loadCode();

  // Backward-compatible: blank/absent -> 'main' (every document that
  // predates MEETING_TYPE, and every existing main-meeting workflow, has no
  // such property set at all).
  check('absent (undefined) -> "main"', sandbox.normalizeMeetingType_(undefined), 'main');
  check('null -> "main"', sandbox.normalizeMeetingType_(null), 'main');
  check('empty string -> "main"', sandbox.normalizeMeetingType_(''), 'main');
  check('whitespace only -> "main"', sandbox.normalizeMeetingType_('   '), 'main');
  check('"main" -> "main"', sandbox.normalizeMeetingType_('main'), 'main');
  check('"MAIN" -> "main"', sandbox.normalizeMeetingType_('MAIN'), 'main');
  check('"adhoc" -> "adhoc"', sandbox.normalizeMeetingType_('adhoc'), 'adhoc');
  check('" adhoc " (whitespace) -> "adhoc"', sandbox.normalizeMeetingType_(' adhoc '), 'adhoc');
  check('"ADHOC" -> "adhoc"', sandbox.normalizeMeetingType_('ADHOC'), 'adhoc');

  // SA4-IMPL-003A: any other non-empty value THROWS -- it must NOT silently
  // become 'main'. A configuration typo picking the wrong meeting/source
  // configuration is worse than a loud, immediate failure.
  expectThrows('"ad-hoc" (hyphenated -- a very plausible typo) throws',
    () => sandbox.normalizeMeetingType_('ad-hoc'), 'Unsupported MEETING_TYPE "ad-hoc"');
  expectThrows('"electronic" throws',
    () => sandbox.normalizeMeetingType_('electronic'), 'Unsupported MEETING_TYPE "electronic"');
  expectThrows('"BogusType" throws',
    () => sandbox.normalizeMeetingType_('BogusType'), 'Unsupported MEETING_TYPE "BogusType"');
  expectThrows('"audio" (a report type, not a meeting type -- a very plausible mix-up) throws',
    () => sandbox.normalizeMeetingType_('audio'), 'Unsupported MEETING_TYPE "audio"');
  expectThrows('error message names both supported values',
    () => sandbox.normalizeMeetingType_('BogusType'), 'Expected "main" or "adhoc"');
}

console.log('getMeetingContext_() -- meeting.type === "adhoc" only when MEETING_TYPE is explicitly set');

REPORT_TYPES.forEach(type => {
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: type } });
  const ctx = sandbox.getMeetingContext_();
  check(`${type}: MEETING_TYPE absent -> meeting.type is "main" (regression, unchanged since SA4-ARCH-003)`,
    ctx.meeting.type, 'main');
  check(`${type}: meeting.name is null for main meetings`, ctx.meeting.name, null);
});

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'BogusType' } });
  expectThrows('getMeetingContext_() propagates the throw for an unrecognized MEETING_TYPE property value',
    () => sandbox.getMeetingContext_(), 'Unsupported MEETING_TYPE "BogusType"');
}

// -------------------- first real ad-hoc context: Audio SWG AH on ULBC-MED --
//
// Values below are the VERIFIED evidence from SA4-ARCH-005/006
// (docs/ADHOC_MEETING_ANALYSIS.md): the real agenda TDoc (S4aA260090), the
// real TDoc-list filename, and the real FTP_BASE/Inbox-Drafts folder paths
// for the SA4_Audio ad-hoc series, combined with the base URL
// (https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/) stated
// once at the top of that document's SA4-ARCH-006 addendum section.

console.log('getMeetingContext_() -- first real ad-hoc context: Audio SWG AH on ULBC-MED');

{
  const ULBC_MED_PROPS = {
    MEETING_TYPE: 'adhoc',
    MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED',
    REPORT_SUFFIX: 'Audio',
    FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/',
    TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/TDoc_List_Meeting_SA4-(AH) Audio SWG on ULBC-MED.xlsx',
    AGENDA_TDOC: 'S4aA260090',
    REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/'
  };
  const { sandbox } = loadCode({ documentProperties: ULBC_MED_PROPS });
  const ctx = sandbox.getMeetingContext_();

  check('group is "SA4"', ctx.group, 'SA4');

  check('meeting.type is "adhoc"', ctx.meeting.type, 'adhoc');
  check('meeting.name is the verified ULBC-MED display name', ctx.meeting.name, ULBC_MED_PROPS.MEETING_NAME);
  check('meeting.folder is null (no fake main identifier exposed)', ctx.meeting.folder, null);
  check('meeting.number is null (no fake main identifier exposed)', ctx.meeting.number, null);
  check('meeting.portalId is null (no fake main identifier exposed; portal ID for this meeting is unverified anyway)', ctx.meeting.portalId, null);

  check('report.type is "Audio"', ctx.report.type, 'Audio');

  // SA4-IMPL-004: agenda selection is now CORRECT for this ad-hoc meeting --
  // agendaPrefix is no longer the old, wrong main-meeting "7." value; it's
  // null, and agendaSelector explicitly represents "every parsed agenda
  // item is in scope" ({mode:'all'}), matching the verified ULBC-MED
  // evidence (a dedicated single-topic ad-hoc with no SWG-plenary prefix
  // concept). This did NOT cause the title-only production consumer to
  // regress: setDocumentTitleFromTemplate_() never reads report.agendaPrefix
  // at all (see tests/title-consumer.test.js, unaffected by this task).
  check('report.agendaPrefix is null (no longer the wrong main-meeting "7." value, SA4-IMPL-004)',
    ctx.report.agendaPrefix, null);
  check('report.agendaSelector is {mode:"all"} (verified: ULBC-MED is a dedicated single-topic ad-hoc, no prefix concept applies)',
    ctx.report.agendaSelector, { mode: 'all' });

  // structureProfile is STILL the old main-meeting-shaped model -- known to
  // be semantically wrong for an ad-hoc meeting, and fixing that is
  // explicitly deferred to the future frontMatterProfile/skeleton work, NOT
  // this task (SA4-IMPL-004 is agendaSelector only). Asserted here only to
  // document and freeze the current (known-imperfect) behavior, not to
  // endorse it.
  check('report.structureProfile is STILL the OLD "main-swg" value -- KNOWN WRONG for ad-hoc, deferred, not fixed here',
    ctx.report.structureProfile, 'main-swg');

  check('sources.ftpBase is the verified SA4_Audio ad-hoc Docs folder', ctx.sources.ftpBase, ULBC_MED_PROPS.FTP_BASE);
  check('sources.tdocListUrl is the verified ULBC-MED TDoc-list URL', ctx.sources.tdocListUrl, ULBC_MED_PROPS.TDOC_LIST_URL);
  check('sources.agendaTdoc is the verified real agenda TDoc "S4aA260090"', ctx.sources.agendaTdoc, 'S4aA260090');
  check('sources.agendaTemplateDocId is undefined (no override given, no ad-hoc-appropriate default exists)',
    ctx.sources.agendaTemplateDocId, undefined);
  check('sources.mailingList PROVISIONALLY reuses the existing Audio mapping (NOT independently verified for ad-hoc traffic -- see SA4-ARCH-006)',
    ctx.sources.mailingList, '3GPP_TSG_SA_WG4_AUDIO');
  check('sources.draftsFolder is null (no ad-hoc equivalent to the main DRAFTS_FOLDERS lookup; revisionsUrl already fully represents the location)',
    ctx.sources.draftsFolder, null);
  check('sources.revisionsUrl is the verified SA4_Audio ad-hoc Inbox/Drafts/ folder', ctx.sources.revisionsUrl, ULBC_MED_PROPS.REVISIONS_URL);

  check('options.showPreviewSnippet still comes from the existing SHOW_PREVIEW_SNIPPET default', ctx.options.showPreviewSnippet, true);
}

// ---------------------------------------------- incomplete ad-hoc config ---

console.log('getMeetingContext_() -- incomplete ad-hoc configuration is represented, not rejected');

{
  // MEETING_TYPE=adhoc and NOTHING else set. Per SA4-IMPL-003 scope
  // ("prefer representation over workflow validation"), getMeetingContext_()
  // must not throw here -- no production workflow consumes ad-hoc context
  // yet, so validating completeness is left to a later task.
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } });
  const ctx = sandbox.getMeetingContext_();

  check('incomplete ad-hoc: meeting.type is still "adhoc"', ctx.meeting.type, 'adhoc');
  check('incomplete ad-hoc: meeting.name is null (MEETING_NAME missing)', ctx.meeting.name, null);
  check('incomplete ad-hoc: meeting.folder/number/portalId remain null',
    [ctx.meeting.folder, ctx.meeting.number, ctx.meeting.portalId], [null, null, null]);
  check('incomplete ad-hoc: sources.ftpBase is undefined (FTP_BASE missing -- NOT a fake main default)',
    ctx.sources.ftpBase, undefined);
  check('incomplete ad-hoc: sources.tdocListUrl is undefined (TDOC_LIST_URL missing)',
    ctx.sources.tdocListUrl, undefined);
  check('incomplete ad-hoc: sources.agendaTdoc is undefined (AGENDA_TDOC missing)',
    ctx.sources.agendaTdoc, undefined);
  check('incomplete ad-hoc: sources.revisionsUrl is undefined (REVISIONS_URL missing)',
    ctx.sources.revisionsUrl, undefined);
  check('incomplete ad-hoc: sources.mailingList is STILL populated (provisional REPORT_SUFFIX-based reuse, independent of the missing fields)',
    ctx.sources.mailingList, '3GPP_TSG_SA_WG4'); // REPORT_SUFFIX defaults to '6G' when unset
  check('incomplete ad-hoc: sources.draftsFolder is still null',
    ctx.sources.draftsFolder, null);
  check('incomplete ad-hoc: getMeetingContext_() does not throw', typeof ctx, 'object');
}

{
  // Individual missing-field cases, each with everything else present, to
  // pin down that ONLY the missing field is affected.
  const FULL_PROPS = {
    MEETING_TYPE: 'adhoc',
    MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED',
    REPORT_SUFFIX: 'Audio',
    FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/',
    TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/TDoc_List_Meeting_SA4-(AH) Audio SWG on ULBC-MED.xlsx',
    AGENDA_TDOC: 'S4aA260090',
    REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Inbox/Drafts/'
  };

  ['MEETING_NAME', 'FTP_BASE', 'TDOC_LIST_URL', 'AGENDA_TDOC', 'REVISIONS_URL'].forEach(missingKey => {
    const props = Object.assign({}, FULL_PROPS);
    delete props[missingKey];
    const { sandbox } = loadCode({ documentProperties: props });
    const ctx = sandbox.getMeetingContext_();

    const fieldMap = {
      MEETING_NAME: () => ctx.meeting.name,
      FTP_BASE: () => ctx.sources.ftpBase,
      TDOC_LIST_URL: () => ctx.sources.tdocListUrl,
      AGENDA_TDOC: () => ctx.sources.agendaTdoc,
      REVISIONS_URL: () => ctx.sources.revisionsUrl
    };
    const expectedWhenMissing = missingKey === 'MEETING_NAME' ? null : undefined;

    check(`missing only ${missingKey}: its own field is ${JSON.stringify(expectedWhenMissing)}`,
      fieldMap[missingKey](), expectedWhenMissing);
    check(`missing only ${missingKey}: meeting.type is still "adhoc" (not rejected)`,
      ctx.meeting.type, 'adhoc');
  });
}

// ================ SA4-PROD-007A: ad-hoc MAILING_LIST override ==============

console.log('getMeetingContext_() -- MAILING_LIST override reaches sources.mailingList for ad-hoc meetings');

{
  // Explicit override present: wins over the provisional cfg.LIST_NAME
  // reuse. Not "FS_6G_MED" or any other meeting-86178-specific value --
  // a generic placeholder name, proving the override mechanism itself,
  // not any particular symbolic list identifier.
  const { sandbox } = loadCode({ documentProperties: {
    MEETING_TYPE: 'adhoc',
    REPORT_SUFFIX: '6G',
    MAILING_LIST: 'SOME_EXPLICIT_LIST_NAME'
  } });
  const ctx = sandbox.getMeetingContext_();
  check('ad-hoc sources.mailingList uses the explicit MAILING_LIST override when set',
    ctx.sources.mailingList, 'SOME_EXPLICIT_LIST_NAME');
}

{
  // No override: falls back to the existing provisional cfg.LIST_NAME
  // reuse, unchanged -- same value the pre-existing "incomplete ad-hoc"
  // check above already pins down for REPORT_SUFFIX '6G'.
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G' } });
  const ctx = sandbox.getMeetingContext_();
  check('ad-hoc sources.mailingList falls back to cfg.LIST_NAME when no MAILING_LIST override is set',
    ctx.sources.mailingList, '3GPP_TSG_SA_WG4');
}

{
  // Main meetings: MAILING_LIST is a Document Property this function
  // reads (via getMeetingIdentityConfig_()), but the MAIN branch's
  // derivedSources never references identity.MAILING_LIST at all -- only
  // the ad-hoc branch's adhocOverrides does. Setting it must have zero
  // effect on a main meeting's resolved mailingList.
  const { sandbox } = loadCode({ documentProperties: {
    REPORT_SUFFIX: 'Audio',
    MAILING_LIST: 'SOME_EXPLICIT_LIST_NAME'
  } });
  const ctx = sandbox.getMeetingContext_();
  check('main-meeting sources.mailingList is UNCHANGED by a MAILING_LIST property (main branch never reads it)',
    ctx.sources.mailingList, '3GPP_TSG_SA_WG4_AUDIO');
}

// ==================== SA4-IMPL-004: main-meeting agendaSelector proof ======

console.log('getMeetingContext_() -- report.agendaSelector is the prefix-mode equivalent of report.agendaPrefix, all 7 report types');

REPORT_TYPES.forEach(type => {
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: type } });
  const ctx = sandbox.getMeetingContext_();

  check(`${type}: report.agendaPrefix is unchanged (not touched by SA4-IMPL-004)`,
    ctx.report.agendaPrefix, sandbox.getReportConfig_().AGENDA_ITEM_PREFIX);
  check(`${type}: report.agendaSelector.mode is "prefix"`, ctx.report.agendaSelector.mode, 'prefix');
  check(`${type}: report.agendaSelector.value === report.agendaPrefix`,
    ctx.report.agendaSelector.value, ctx.report.agendaPrefix);
});

// ---------------- SA4-IMPL-004/005: production filtering caller status ----
//
// SA4-IMPL-004 introduced agendaSelectorMatches_() with NO production
// caller at all. SA4-IMPL-005 migrated exactly ONE consumer --
// downloadAndGroupTdocs_() -- to call it; see tests/tdoc-agenda-filter.test.js
// for that migration's own dedicated characterization. This section now
// documents the REMAINING known non-callers (agenda-structure parsing and
// the skeleton builder, both deliberately left unmigrated) rather than
// treating "zero callers" as still true -- that would misrepresent the
// current, approved state of the codebase.

console.log('source-structure assertion: agendaSelectorMatches_() production caller status');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  // Strip the function's OWN definition and doc comment first, so the
  // string "agendaSelectorMatches_(" appearing in ITS OWN JSDoc/declaration
  // doesn't count as a "caller". Also strip normalizeAgendaSelector_'s
  // definition for the same reason (it doesn't call agendaSelectorMatches_,
  // but keeping the check symmetric and simple).
  const withoutOwnDefinition = source.replace(/function agendaSelectorMatches_\([\s\S]*?\n}\n/, '');

  function extractBody(src, fnName) {
    const startMatch = src.match(new RegExp('^function ' + fnName + '\\(', 'm'));
    if (!startMatch) return null;
    const startIndex = startMatch.index;
    const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnRe.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnRe.exec(src);
    const endIndex = next ? next.index : src.length;
    return src.slice(startIndex, endIndex);
  }

  const migratedBody = extractBody(withoutOwnDefinition, 'downloadAndGroupTdocs_');
  if (!migratedBody) {
    failures++;
    console.log('  FAIL could not locate "function downloadAndGroupTdocs_(" in Code.js');
  } else {
    check('downloadAndGroupTdocs_() DOES call agendaSelectorMatches_() (migrated by SA4-IMPL-005 -- see tests/tdoc-agenda-filter.test.js)',
      /\bagendaSelectorMatches_\s*\(/.test(migratedBody), true);
  }

  const KNOWN_NON_CALLERS = [
    'parseAgendaForReport_', 'buildSkeletonWithTdocTables', 'continuousUpdate'
  ];

  KNOWN_NON_CALLERS.forEach(fnName => {
    const body = extractBody(withoutOwnDefinition, fnName);
    if (!body) {
      failures++;
      console.log(`  FAIL could not locate "function ${fnName}(" in Code.js`);
      return;
    }
    check(`${fnName}() does NOT call agendaSelectorMatches_() (unmigrated as of SA4-IMPL-005)`,
      /\bagendaSelectorMatches_\s*\(/.test(body), false);
  });

  // And the positive half of the same proof: getMeetingContext_() only ever
  // CONSTRUCTS a selector via normalizeAgendaSelector_(), it never MATCHES
  // one -- confirm it doesn't call agendaSelectorMatches_() either.
  const gmcBody = extractBody(withoutOwnDefinition, 'getMeetingContext_');
  if (gmcBody) {
    check('getMeetingContext_() constructs selectors via normalizeAgendaSelector_() but does not call agendaSelectorMatches_() (representation only)',
      /\bagendaSelectorMatches_\s*\(/.test(gmcBody), false);
    check('getMeetingContext_() DOES call normalizeAgendaSelector_() (both main and ad-hoc branches construct a selector)',
      /\bnormalizeAgendaSelector_\s*\(/.test(gmcBody), true);
  } else {
    failures++;
    console.log('  FAIL could not locate "function getMeetingContext_()" in Code.js');
  }
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll MeetingContext tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
