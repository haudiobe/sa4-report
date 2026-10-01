/**
 * DIAGNOSTIC ONLY -- TEMPLATE-002B live smoke test, "Exceeded maximum
 * execution time" in Run Full Report Build (report for meeting 86178,
 * template T-2026.10.0, Code.js 2.17.0, 2026-10-01).
 *
 * No production code is changed. Every check records what the code does
 * TODAY, so the numbers in the diagnosis are measured, not estimated. The
 * checks marked [DEFECT] describe the behaviour a fix has to change; they
 * will then need to be inverted.
 *
 * Finding: Run Full Report Build is one execution with a 6-minute limit, and
 * inside it the user has to dismiss FOUR blocking pop-ups (one after each
 * step) before the last step can even start. The work itself took well under
 * two minutes in the live log; the rest of the six minutes was spent waiting
 * in those pop-ups. The missing Reviewer token is noise: no network call is
 * made without it.
 *
 * Run: node tests/diagnose-template002b-fullbuild-timeout.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');

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

const SOURCE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
function fnBody(name) {
  const start = SOURCE.indexOf('\nfunction ' + name + '(');
  const end = SOURCE.indexOf('\n}\n', start);
  return SOURCE.slice(start, end);
}

const TDOCS = 34; // the live report: "Found 14 agenda items and 34 TDOCs"

/** A report with TDOCS built tables; the step functions run for real, only their network workers are stubbed. */
function setup(scriptProperties) {
  const loaded = loadCode({ scriptProperties: scriptProperties || {} });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  for (let i = 0; i < TDOCS; i++) body.insertTable(i, [['TDoc', 'S4aP26' + String(60 + i).padStart(4, '0')], ['Title', 't'], ['Agenda Item', '5.4']]);
  const events = [];
  const ui = {
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES', NO: 'NO' },
    alert: (title) => { events.push('ALERT ' + title); return 'YES'; }
  };
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'DOC' });
  const logs = [];
  s.Logger = { log: (m) => logs.push(String(m)) };
  let tokenReads = 0;
  const realScriptProps = s.PropertiesService.getScriptProperties();
  s.PropertiesService.getScriptProperties = () => ({
    getProperty: (k) => { if (k === 'REVIEWER_API_TOKEN') tokenReads++; return realScriptProps.getProperty(k); },
    setProperty: realScriptProps.setProperty, deleteProperty: realScriptProps.deleteProperty, getKeys: realScriptProps.getKeys
  });
  const fetches = [];
  s.UrlFetchApp = { fetch: (url) => { fetches.push(url); return { getResponseCode: () => 404, getContentText: () => '' }; } };
  // Network workers of the steps: recorded, not run.
  s.buildSkeletonWithTdocTables = () => { events.push('WORK skeleton + TDoc tables + formatting'); events.push('ALERT Done'); };
  s.getCollectorConfig_ = () => ({});
  s.checkRSSFeed_ = () => events.push('WORK e-mail collection');
  s.updateRevisions_ = () => events.push('WORK revision collection');
  s.removeRowHeightAndSpacing = () => events.push('WORK formatting');
  return { s, events, logs, fetches, tokenReads: () => tokenReads };
}

console.log('1. Run Full Report Build: the pop-ups between the steps');
{
  check('the skeleton step itself ends with a blocking "Done" pop-up whenever a UI exists (source)',
    /removeRowHeightAndSpacing\(\);\n  Logger\.log\(`Done: Built skeleton[^\n]*\n  try \{\n    DocumentApp\.getUi\(\)\.alert\('Done'/.test(fnBody('buildSkeletonWithTdocTables')), true);
  check('... and already formats the document once (source)', /\n  removeRowHeightAndSpacing\(\);\n/.test(fnBody('buildSkeletonWithTdocTables')), true);

  const r = setup();
  r.s.runFullReportBuild();
  check('[DEFECT] one execution: every step is followed by a pop-up that must be clicked before the next step starts', r.events, [
    'ALERT Run Full Report Build',
    'WORK skeleton + TDoc tables + formatting', 'ALERT Done',
    'WORK e-mail collection', 'ALERT Success',
    'WORK revision collection', 'ALERT Success',
    'ALERT Success',                       // "Abstract step completed"
    'WORK formatting',
    'ALERT Success'
  ]);
  const afterStart = r.events.slice(1);
  check('[DEFECT] four blocking pop-ups inside the 6-minute execution, before the final one',
    afterStart.filter((e) => /^ALERT/.test(e)).length - 1, 4);
  check('[DEFECT] the document is formatted twice (inside the skeleton step, and again at the end)',
    [/\n  removeRowHeightAndSpacing\(\);\n/.test(fnBody('buildSkeletonWithTdocTables')), r.events.filter((e) => e === 'WORK formatting').length], [true, 1]);
  check('the step functions are the menu\'s own single-step items, pop-up included (source)',
    ['collectEmailDiscussionOnly', 'collectRevisionsOnly', 'addAbstractsOnly'].map((f) => /DocumentApp\.getUi\(\)\.alert\('Success'/.test(fnBody(f))), [true, true, true]);
}

console.log('2. REVIEWER_API_TOKEN: what a missing token costs');
{
  const r = setup({});
  r.s.addAbstractsOnly();
  check('no token: not one network request', r.fetches, []);
  check('[NOISE] no token: the token is read once per TDoc table, and one identical line is logged each time',
    [r.tokenReads(), r.logs.filter((l) => l === 'No REVIEWER_API_TOKEN found in script properties').length], [TDOCS, TDOCS]);
  check('the abstract step still reports success with 0 abstracts', r.events, ['ALERT Success']);
  check('[NOISE] the skeleton step asks for the token once more per table, because it never sets the skip flag (source)',
    [/fetchAndAddAbstract_\(table, parsedTdoc\.raw\);/.test(fnBody('createTDocTableFromData_')), /SKIP_ABSTRACTS_DURING_TABLE_BUILD/.test(fnBody('buildSkeletonWithTdocTables')), /SKIP_ABSTRACTS_DURING_TABLE_BUILD/.test(fnBody('runFullReportBuild'))],
    [true, false, false]);

  const withToken = setup({ REVIEWER_API_TOKEN: 'TESTONLY-not-a-real-token' });
  withToken.s.addAbstractsOnly();
  check('with a token the abstract step is one Reviewer request per table -- the genuinely slow optional part', withToken.fetches.length, TDOCS);
}

console.log('3. the build is repeatable');
{
  const build = fnBody('buildSkeletonWithTdocTables');
  check('Full Build clears the document before it rebuilds (no duplicate tables on a re-run)', /getBody\(\)\.clear\(\)/.test(build), true);
  check('... after reading the saved reallocations and checking the sources', build.indexOf('assertMeetingReadyToBuild_') < build.indexOf('.clear()') && build.indexOf('getReallocationMap_') < build.indexOf('.clear()'), true);
  check('collected e-mail is re-rendered from its per-TDoc store, not only appended', /cell\.setText\(''\);/.test(fnBody('checkRSSFeed_')), true);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll diagnostic checks passed.');
process.exitCode = failures ? 1 : 0;
