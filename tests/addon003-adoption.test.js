/**
 * ADDON-003 — adoptReportDocumentForAddon_(): explicit, interactive-only
 * migration of a document's report state from Document Properties
 * (Backend A) into the central, documentId-namespaced backend (Backend B),
 * followed by registry registration -- ONLY on successful verification.
 *
 * See the "ADDON-003 -- DOCUMENT ADOPTION" block in Code.js for the exact
 * property classification (required / useful-but-optional / interactive-
 * only) this is built on.
 *
 * Run: node tests/addon003-adoption.test.js
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

function expectThrows(name, fn, messageSubstring) {
  try {
    fn();
    failures++;
    console.log(`  FAIL ${name}\n         expected a throw, but none occurred`);
  } catch (e) {
    if (messageSubstring && e.message.indexOf(messageSubstring) === -1) {
      failures++;
      console.log(`  FAIL ${name}\n         expected message to contain "${messageSubstring}"\n         actual: ${e.message}`);
    } else {
      console.log(`  ok   ${name}`);
    }
  }
}

const A_REAL_MEETING_STATE = {
  MEETING_FOLDER: 'TSGS4_136_Montreal',
  MEETING_NUMBER: '136',
  MEETING_ID: '86178',
  REPORT_SUFFIX: 'Audio',
  FETCH_ABSTRACTS_ON_UPDATE: 'true',
  REVISION_MAP: JSON.stringify({ 'S4-260200': 'S4-260101' }),
  DEADLINE_EXTENSIONS: JSON.stringify({ 'S4-260101': { extended: true } }),
  'DISCUSS_S4-260101': JSON.stringify({ msg1: { author: 'A', date: '2026-09-01' } }),
  'REVIS_S4-260101': JSON.stringify({ rev1: { author: 'B', date: '2026-09-02' } }),
  'A1_EMPTY_CACHE_abc123': JSON.stringify({ ts: 1, empty: true }),
  'REVIEWER_NO_SUMMARY_CACHE_S4-260101': JSON.stringify({ ts: 1, statusCode: 404 }),
  // interactive-only -- must NOT be copied
  PARSED_AGENDA: JSON.stringify([{ number: '7.1' }]),
  PARSED_AGENDA_ALL: JSON.stringify([{ number: '7' }, { number: '7.1' }]),
  SKIP_ABSTRACTS_DURING_TABLE_BUILD: 'true'
};

function withActiveDocument(sandbox, documentId) {
  sandbox.DocumentApp.getActiveDocument = () => ({
    getId: () => documentId,
    getBody: () => ({})
  });
}

// -------------------------------------------------------------- happy path

console.log('adoptReportDocumentForAddon_() -- copies the required/optional state, skips interactive-only');

{
  const { sandbox } = loadCode({ documentProperties: A_REAL_MEETING_STATE });
  withActiveDocument(sandbox, 'DOC_ADOPT_1');

  const result = sandbox.adoptReportDocumentForAddon_();

  check('result.documentId matches the active document',
    result.documentId, 'DOC_ADOPT_1');
  check('result.verified is true',
    result.verified, true);
  check('result.mismatches is empty',
    result.mismatches, []);
  check('result.registered is true',
    result.registered, true);

  check('required fixed keys were copied',
    ['MEETING_FOLDER', 'MEETING_NUMBER', 'MEETING_ID', 'REPORT_SUFFIX', 'FETCH_ABSTRACTS_ON_UPDATE', 'REVISION_MAP', 'DEADLINE_EXTENSIONS']
      .every(k => result.copiedKeys.indexOf(k) !== -1), true);
  check('per-TDoc/per-URL cache families were copied',
    ['DISCUSS_S4-260101', 'REVIS_S4-260101', 'A1_EMPTY_CACHE_abc123', 'REVIEWER_NO_SUMMARY_CACHE_S4-260101']
      .every(k => result.copiedKeys.indexOf(k) !== -1), true);
  check('interactive-only keys were NOT copied',
    ['PARSED_AGENDA', 'PARSED_AGENDA_ALL', 'SKIP_ABSTRACTS_DURING_TABLE_BUILD']
      .some(k => result.copiedKeys.indexOf(k) !== -1), false);

  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_ADOPT_1' });
  check('copied values are byte-identical in the central backend',
    central.getProperty('REVISION_MAP'), A_REAL_MEETING_STATE.REVISION_MAP);
  check('interactive-only key is genuinely absent from the central backend',
    central.getProperty('PARSED_AGENDA'), null);
}

console.log('adoptReportDocumentForAddon_() -- source Document Properties is left completely unchanged');

{
  const { sandbox, docProps } = loadCode({ documentProperties: A_REAL_MEETING_STATE });
  withActiveDocument(sandbox, 'DOC_ADOPT_2');

  const before = JSON.stringify(docProps._store);
  sandbox.adoptReportDocumentForAddon_();
  const after = JSON.stringify(docProps._store);

  check('Document Properties store is byte-identical before/after adoption',
    after, before);
}

console.log('adoptReportDocumentForAddon_() -- registry entry reflects the adopted meeting identity');

{
  const { sandbox } = loadCode({ documentProperties: A_REAL_MEETING_STATE });
  withActiveDocument(sandbox, 'DOC_ADOPT_3');

  const result = sandbox.adoptReportDocumentForAddon_();

  check('registry entry meetingId comes from the copied MEETING_ID',
    result.registryEntry.meetingId, '86178');
  check('registry entry is retrievable via getRegisteredReportDocument_()',
    sandbox.getRegisteredReportDocument_('DOC_ADOPT_3').documentId, 'DOC_ADOPT_3');
}

// ----------------------------------------------------------------- idempotent

console.log('adoptReportDocumentForAddon_() -- repeat adoption is idempotent');

{
  const { sandbox } = loadCode({ documentProperties: A_REAL_MEETING_STATE });
  withActiveDocument(sandbox, 'DOC_ADOPT_REPEAT');

  const first = sandbox.adoptReportDocumentForAddon_();
  const second = sandbox.adoptReportDocumentForAddon_();

  check('second adoption also verifies successfully',
    second.verified, true);
  check('registeredAt is unchanged across repeat adoption',
    second.registryEntry.registeredAt, first.registryEntry.registeredAt);
  check('copiedKeys set is stable across repeat adoption',
    second.copiedKeys.slice().sort(), first.copiedKeys.slice().sort());
}

console.log('adoptReportDocumentForAddon_() -- repeat adoption picks up an interactive config change since the first run');

{
  const { sandbox, docProps } = loadCode({ documentProperties: A_REAL_MEETING_STATE });
  withActiveDocument(sandbox, 'DOC_ADOPT_UPDATE');

  sandbox.adoptReportDocumentForAddon_();
  docProps.setProperty('REPORT_SUFFIX', 'Video'); // simulate an interactive edit after first adoption
  sandbox.adoptReportDocumentForAddon_();

  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_ADOPT_UPDATE' });
  check('the central backend reflects the LATEST source value after a second adoption',
    central.getProperty('REPORT_SUFFIX'), 'Video');
}

// ----------------------------------------------------------- absent keys

console.log('adoptReportDocumentForAddon_() -- never-set optional keys are simply skipped, not invented');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  withActiveDocument(sandbox, 'DOC_ADOPT_SPARSE');

  const result = sandbox.adoptReportDocumentForAddon_();

  check('result.verified is true even with a sparse/minimal source',
    result.verified, true);
  check('REVISION_MAP (never set) is not present in copiedKeys',
    result.copiedKeys.indexOf('REVISION_MAP') === -1, true);
}

// ------------------------------------------------------------- failure path

console.log('adoptReportDocumentForAddon_() -- a verification failure does NOT register the document');

{
  const { sandbox, scriptProps } = loadCode({ documentProperties: A_REAL_MEETING_STATE });
  withActiveDocument(sandbox, 'DOC_ADOPT_FAIL');

  // Simulate a write that silently fails for exactly one key (e.g. a
  // transient Apps Script quota hiccup) by intercepting the underlying
  // Script Properties store's setProperty and dropping REVISION_MAP.
  const realSetProperty = scriptProps.setProperty;
  scriptProps.setProperty = function (key, value) {
    if (key.indexOf('REVISION_MAP') !== -1) return; // silently drop this one write
    return realSetProperty(key, value);
  };

  const result = sandbox.adoptReportDocumentForAddon_();

  check('result.verified is false',
    result.verified, false);
  check('the failed key is reported in mismatches',
    result.mismatches, ['REVISION_MAP']);
  check('result.registered is false',
    result.registered, false);
  check('no registry entry was created for a failed adoption (no half-registered document)',
    sandbox.getRegisteredReportDocument_('DOC_ADOPT_FAIL'), null);
}

// -------------------------------------------------------------- misuse guard

console.log('adoptReportDocumentForAddon_() -- refuses to run in addon-background mode (interactive-only)');

{
  const { sandbox } = loadCode({ documentProperties: A_REAL_MEETING_STATE });
  expectThrows('addon-background context is rejected',
    () => sandbox.adoptReportDocumentForAddon_({ mode: 'addon-background', documentId: 'DOC_X' }),
    'interactive-only');
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
