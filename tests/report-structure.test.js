/**
 * SA4-ARCH-002 — Report-structure characterization tests.
 *
 * These tests do NOT build a report or touch a fake/real Google Doc. They
 * characterize the deterministic, callable-without-a-Doc decisions that
 * currently drive report structure per report type -- agenda prefix,
 * mailing list, drafts folder, and generated title -- by calling the real
 * production functions (getReportConfig_, getAgendaPrefixForReportType_,
 * generateReportTitle_) and diffing the result against the committed golden
 * snapshot in tests/fixtures/main-meeting-profile.expected.json.
 *
 * IMPORTANT LIMITATION (do not "fix" by extracting new production code):
 * whether a report type follows the "6G" skeleton (11.0.1-11.0.4 subsections)
 * or the "SWG" skeleton (X.1/X.1.2/X.2) is decided by inline booleans
 * (`is6G`, `isSWGReport`) INSIDE buildSkeletonWithTdocTables() (Code.js
 * ~5350, ~5459) -- it is not exposed as a callable function. This test suite
 * cannot exercise that branch without either building a fake DocumentApp
 * (explicitly out of scope for SA4-ARCH-002) or modifying Code.js (also out
 * of scope). The `structureBranch`/`isSWGReport` fields in the fixture were
 * therefore derived by reading the source, not by executing it, and are
 * checked here only for internal consistency (fixture vs. itself / vs. the
 * report-type list), not against live code. This gap is exactly the kind of
 * thing a future MeetingContext/profile extraction should close by making
 * the branch decision a real, callable, testable function.
 *
 * Run: node tests/report-structure.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode } = require('./helpers/load-code.js');

const FIXTURE_PATH = path.join(__dirname, 'fixtures', 'main-meeting-profile.expected.json');
const fixture = JSON.parse(fs.readFileSync(FIXTURE_PATH, 'utf8'));

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

// ---------------------------------------- per-report-type structural decisions

console.log('per-report-type structural decisions (live Code.js vs. golden fixture)');

const reportTypes = Object.keys(fixture.reportTypeProfiles);
check('fixture lists the 7 report types Code.js currently supports',
  reportTypes.sort(),
  ['6G', 'Audio', 'Liaison', 'MBS', 'New', 'RTC', 'Video'].sort());

reportTypes.forEach(type => {
  const expectedProfile = fixture.reportTypeProfiles[type];
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: type } });
  const cfg = sandbox.getReportConfig_();

  check(`${type}: agendaPrefix`, sandbox.getAgendaPrefixForReportType_(type), expectedProfile.agendaPrefix);
  check(`${type}: agendaPrefix as seen through getReportConfig_().AGENDA_ITEM_PREFIX`,
    cfg.AGENDA_ITEM_PREFIX, expectedProfile.agendaPrefix);
  check(`${type}: mailingList (getReportConfig_().LIST_NAME)`, cfg.LIST_NAME, expectedProfile.mailingList);
  check(`${type}: draftsFolder (getReportConfig_().DRAFTS_FOLDER)`, cfg.DRAFTS_FOLDER, expectedProfile.draftsFolder);
  check(`${type}: REVISIONS_URL`, cfg.REVISIONS_URL, expectedProfile.revisionsUrlPattern);
});

// ------------------------------------------------- unknown/unconfigured type

console.log('unknown/unconfigured report type fallback');

{
  const { sandbox } = loadCode();
  check('unrecognized report type -> agendaPrefix falls back to "11."',
    sandbox.getAgendaPrefixForReportType_('SomeUnknownType'),
    fixture.unknownReportTypeFallback.agendaPrefix);
}
{
  const { sandbox } = loadCode({ documentProperties: { REPORT_SUFFIX: 'SomeUnknownType' } });
  const cfg = sandbox.getReportConfig_();
  check('unrecognized report type -> AGENDA_ITEM_PREFIX', cfg.AGENDA_ITEM_PREFIX, fixture.defaultReportConfigForUnknownType.AGENDA_ITEM_PREFIX);
  check('unrecognized report type -> LIST_NAME', cfg.LIST_NAME, fixture.defaultReportConfigForUnknownType.LIST_NAME);
  check('unrecognized report type -> DRAFTS_FOLDER', cfg.DRAFTS_FOLDER, fixture.defaultReportConfigForUnknownType.DRAFTS_FOLDER);
}

// -------------------------------------------------------- default report config

console.log('default report config (brand-new Doc, no saved configuration)');

{
  const { sandbox } = loadCode();
  const cfg = sandbox.getReportConfig_();
  Object.keys(fixture.defaultReportConfig).forEach(key => {
    if (key === '_note') return;
    check(`defaultReportConfig.${key}`, cfg[key], fixture.defaultReportConfig[key]);
  });
}

// -------------------------------------------------------------------- titles

console.log('generated titles for each report type');

{
  const { sandbox } = loadCode();
  fixture.titleExamples.forEach(({ input, output }) => {
    check(`title for REPORT_SUFFIX=${input.REPORT_SUFFIX}`, sandbox.generateReportTitle_(input), output);
  });
}

// ------------------------------------------ fixture internal consistency check

console.log('fixture internal consistency (structureBranch / isSWGReport, source-derived -- see file header)');

const SWG_TYPES = ['Audio', 'Video', 'MBS', 'RTC'];
reportTypes.forEach(type => {
  const profile = fixture.reportTypeProfiles[type];
  const expectedIsSwg = SWG_TYPES.includes(type);
  check(`${type}: isSWGReport matches the SWG type list (Audio/Video/MBS/RTC)`,
    profile.isSWGReport, expectedIsSwg);
  const expectedBranch = type === '6G' ? 'plenary-6g' : (expectedIsSwg ? 'swg' : 'other');
  check(`${type}: structureBranch matches isSWGReport/6G classification`,
    profile.structureBranch, expectedBranch);
});

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll report-structure characterization checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
