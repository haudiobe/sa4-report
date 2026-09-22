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
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

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

// --------------------------- SA4-PROD-008: X.2 iprSection, tied to live source

console.log('iprSection fixture claims, tied to live Code.js source (not just fixture self-consistency)');

function extractFunctionBody(source, fnName) {
  const startMatch = source.match(new RegExp('^function ' + fnName + '\\(', 'm'));
  if (!startMatch) return null;
  const startIndex = startMatch.index;
  const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
  nextFnRe.lastIndex = startIndex + startMatch[0].length;
  const next = nextFnRe.exec(source);
  const endIndex = next ? next.index : source.length;
  return source.slice(startIndex, endIndex);
}

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const skeletonBody = extractFunctionBody(source, 'buildSkeletonWithTdocTables');

  check('buildSkeletonWithTdocTables() exists', !!skeletonBody, true);

  if (skeletonBody) {
    // The live invariant this fixture's per-profile "iprSection" claims
    // actually rest on: appendStandardIprSection_() is called exactly once,
    // inside the iprSection branch that only is6G===false profiles reach
    // (the is6G block returns early for its own openingSection/
    // registrationSection/reallocationSection/iprSection quartet BEFORE
    // that branch is reached) -- and the is6G block's OWN "X.0.4 IPR and
    // antitrust reminder" still copies template content, unchanged.
    check('buildSkeletonWithTdocTables() calls appendStandardIprSection_() (the canonical-generated path every non-plenary-6g profile below claims)',
      /appendStandardIprSection_\(/.test(skeletonBody), true);
    check('is6G\'s own nested "X.0.4 IPR and antitrust reminder" still exists and still copies template content (the plenary-6g profile\'s claimed, UNCHANGED path)',
      /\.0\.4 IPR and antitrust reminder[\s\S]{0,300}copySectionContentWithReplacement_/.test(skeletonBody), true);
    check('the is6G block still returns early for its own iprSection before the canonical-generated branch could be reached (plenary-6g never calls appendStandardIprSection_)',
      /if\s*\(is6G\s*&&\s*\([\s\S]*?item\.number === iprSection[\s\S]*?\)\)\s*\{\s*return;/.test(skeletonBody), true);
  }

  // Cross-check every profile's fixture claim against the single rule this
  // implies: plenary-6g (and only plenary-6g) keeps the old template-copy
  // path; every other profile claims the new canonical-generated path.
  reportTypes.forEach(type => {
    const profile = fixture.reportTypeProfiles[type];
    check(`${type}: fixture.iprSection.source is present`, typeof profile.iprSection?.source, 'string');
    const expectedSource = profile.structureBranch === 'plenary-6g' ? 'is6G-nested-template-copy' : 'canonical-generated';
    check(`${type}: fixture.iprSection.source matches the structureBranch-derived rule (plenary-6g keeps template-copy, everything else is canonical-generated)`,
      profile.iprSection && profile.iprSection.source, expectedSource);
    check(`${type}: fixture.iprSection.changedBySA4PROD008 matches (only plenary-6g is unchanged)`,
      profile.iprSection && profile.iprSection.changedBySA4PROD008, profile.structureBranch !== 'plenary-6g');
  });
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll report-structure characterization checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
