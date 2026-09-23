/**
 * ADDON-005B — live acceptance test found a real gap the ADDON-004 audit
 * missed: removeEmptyParagraphs_() (called from removeRowHeightAndSpacing()
 * -> continuousUpdateCore_()'s formatting stage) called
 * getActiveDocumentBodyCounted_('removeEmptyParagraphs_') with NO context
 * argument, so it silently fell back to DocumentApp.getActiveDocument()
 * even during a background/scheduler execution -- which has no active
 * document at all. Live symptom: "Cannot read properties of null
 * (reading 'getBody')".
 *
 * This suite does NOT mock removeEmptyParagraphs_ (or any other
 * background-reachable function) -- it sets
 * DocumentApp.getActiveDocument() to literally return null (matching the
 * real background-execution condition and reproducing the exact live
 * error class) and runs continuousUpdateForDocument_() with every
 * REAL function in its call graph, proving none of them reach
 * getActiveDocument(). Only the network/xlsx-download boundary
 * (downloadAndGroupTdocs_) and the pure decision function
 * shouldReformatAfterUpdate_ (to force the formatting stage to run
 * without needing a full TDoc-table-insertion simulation) are stubbed --
 * neither touches DocumentApp.
 *
 * Run: node tests/addon005b-background-document-context.test.js
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

function makeFakeBackgroundDocument(id) {
  const body = {
    getTables: () => [],
    getParagraphs: () => [],
    clear: () => body
  };
  return { getId: () => id, getBody: () => body };
}

function setUpAdoptedEnabledDocument(sandbox, documentId, extraConfig) {
  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: documentId });
  const config = Object.assign({
    MEETING_FOLDER: 'TSGS4_999_Test',
    MEETING_NUMBER: '999',
    MEETING_ID: '99999',
    REPORT_SUFFIX: 'Audio',
    FETCH_ABSTRACTS_ON_UPDATE: 'false'
  }, extraConfig || {});
  Object.keys(config).forEach((k) => central.setProperty(k, config[k]));
  sandbox.registerReportDocument_(documentId, { enabled: true, meetingId: '99999' });
}

// ---------------------------------------------------- the live failure, reproduced

console.log('continuousUpdateForDocument_() -- full run with DocumentApp.getActiveDocument() === null and the formatting stage forced on');

{
  const { sandbox } = loadCode();
  setUpAdoptedEnabledDocument(sandbox, 'DOC_NO_ACTIVE');

  let activeDocumentCalls = 0;
  sandbox.DocumentApp.getActiveDocument = () => { activeDocumentCalls++; return null; };
  sandbox.DocumentApp.openById = (id) => makeFakeBackgroundDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => ({}); // network/xlsx boundary, unrelated to this bug
  sandbox.shouldReformatAfterUpdate_ = () => true; // force the formatting stage (removeRowHeightAndSpacing) to run, without a full TDoc-insertion simulation

  let result;
  let thrown = null;
  try {
    result = sandbox.continuousUpdateForDocument_('DOC_NO_ACTIVE');
  } catch (e) {
    thrown = e;
  }

  check('completed without throwing (the live bug threw "Cannot read properties of null")',
    thrown === null, true);
  check('result reports success', result && result.success, true);
  check('DocumentApp.getActiveDocument() was NEVER called anywhere in the reachable graph',
    activeDocumentCalls, 0);
}

// -------------------------------------- targeted: removeRowHeightAndSpacing's full real chain

console.log('removeRowHeightAndSpacing(context) -- real setTwoColumnTDocTableWidths_ + real removeEmptyParagraphs_, neither touches getActiveDocument');

{
  const { sandbox } = loadCode();
  let activeDocumentCalls = 0;
  sandbox.DocumentApp.getActiveDocument = () => { activeDocumentCalls++; return null; };

  const fakeDoc = makeFakeBackgroundDocument('DOC_FORMAT');
  const context = { document: fakeDoc, mode: 'addon-background' };

  sandbox.removeRowHeightAndSpacing(context);

  check('DocumentApp.getActiveDocument() was never called', activeDocumentCalls, 0);
}

console.log('removeEmptyParagraphs_(context) -- in isolation, uses the explicit context, never getActiveDocument');

{
  const { sandbox } = loadCode();
  let activeDocumentCalls = 0;
  sandbox.DocumentApp.getActiveDocument = () => { activeDocumentCalls++; return null; };

  let getParagraphsCalls = 0;
  const fakeBody = { getParagraphs: () => { getParagraphsCalls++; return []; } };
  const fakeDoc = { getId: () => 'DOC_EP', getBody: () => fakeBody };

  sandbox.removeEmptyParagraphs_({ document: fakeDoc, mode: 'addon-background' });

  check('used the explicit document\'s body', getParagraphsCalls, 1);
  check('DocumentApp.getActiveDocument() was never called', activeDocumentCalls, 0);
}

console.log('removeEmptyParagraphs_() -- no context (legacy call) still falls back to the active document, unchanged');

{
  const { sandbox } = loadCode();
  let getParagraphsCalls = 0;
  const fakeBody = { getParagraphs: () => { getParagraphsCalls++; return []; } };
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => fakeBody });

  sandbox.removeEmptyParagraphs_();

  check('legacy no-context call still uses the active document (unchanged behavior)', getParagraphsCalls, 1);
}

// -------------------------------------------------------- exhaustive structural audit

console.log('source-structure: every getActiveDocumentBodyCounted_() call site passes a context argument');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  // Every call site EXCEPT the function's own definition and the
  // getReportBody_() fallback comment references must pass a second
  // (context) argument -- a call with only a label string is exactly the
  // removeEmptyParagraphs_() bug this stage fixed.
  const callSitePattern = /getActiveDocumentBodyCounted_\(\s*'[^']*'\s*(,\s*context\s*)?\)/g;
  const bareCallPattern = /getActiveDocumentBodyCounted_\(\s*'[^']*'\s*\)/g;

  const lines = source.split('\n');
  const bareCallLines = [];
  lines.forEach((line, i) => {
    if (bareCallPattern.test(line) && !line.trim().startsWith('*') && !line.trim().startsWith('//')) {
      bareCallLines.push(`${i + 1}: ${line.trim()}`);
    }
    bareCallPattern.lastIndex = 0;
  });

  check('no remaining getActiveDocumentBodyCounted_(\'label\') call without a context argument',
    bareCallLines, []);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
