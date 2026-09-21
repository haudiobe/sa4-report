/**
 * SA4-ARCH-003 — MeetingContext compatibility-layer tests.
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

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll MeetingContext tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
