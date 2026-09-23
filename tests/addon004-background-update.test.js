/**
 * ADDON-004 — continuousUpdateForDocument_(): the per-document background
 * unit of work. Structurally proves it never touches
 * DocumentApp.getActiveDocument() / PropertiesService.getDocumentProperties()
 * / LockService.getDocumentLock(), reads its config/state from the central
 * backend, opens the correct document by id, and only advances
 * `lastRunAt` on success.
 *
 * continuousUpdateCore_() catches and logs its own internal per-stage
 * errors (unchanged legacy behavior -- see its header comment) but now
 * (as of the live-test follow-up fix) reports that failure back via its
 * {success, error} return value, which continuousUpdateForDocument_()
 * checks before advancing lastRunAt. Tests below that are NOT specifically
 * about a failing TDoc-list fetch stub `downloadAndGroupTdocs_` to a
 * trivial `() => ({})` -- its raw UrlFetchApp.fetch() call is NOT wrapped
 * in its own try/catch (confirmed by trace: only the OUTER
 * continuousUpdateCore_() try/catch catches it), so leaving it un-stubbed
 * against the default throwing sandbox fetch would make EVERY run fail,
 * which is correct production behavior but not what most of these tests
 * are individually isolating. The "internal per-stage failure" section
 * below deliberately does NOT stub it, to exercise exactly that path.
 *
 * Run: node tests/addon004-background-update.test.js
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

function makeFakeDocument(id, tables) {
  const body = {
    getTables: () => tables || [],
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

// ------------------------------------------------------------ misuse guards

console.log('continuousUpdateForDocument_() -- validates documentId and registration up front');

{
  const { sandbox } = loadCode();
  expectThrows('empty documentId throws', () => sandbox.continuousUpdateForDocument_(''), 'documentId is required');
  expectThrows('unregistered document throws', () => sandbox.continuousUpdateForDocument_('GHOST'), 'is not registered');
}

{
  const { sandbox } = loadCode();
  sandbox.registerReportDocument_('DOC_DISABLED', { enabled: false });
  expectThrows('registered-but-disabled document throws',
    () => sandbox.continuousUpdateForDocument_('DOC_DISABLED'), 'not enabled');
}

console.log('continuousUpdateForDocument_() -- validates BEFORE touching DocumentApp/PropertiesService at all');

{
  const { sandbox } = loadCode();
  let activeDocCalls = 0;
  let docPropsCalls = 0;
  sandbox.DocumentApp.getActiveDocument = () => { activeDocCalls++; throw new Error('should not be reached'); };
  sandbox.DocumentApp.openById = () => { throw new Error('should not be reached'); };
  const realGetDocProps = sandbox.PropertiesService.getDocumentProperties;
  sandbox.PropertiesService.getDocumentProperties = () => { docPropsCalls++; return realGetDocProps(); };

  try { sandbox.continuousUpdateForDocument_('NEVER_REGISTERED'); } catch (e) { /* expected */ }

  check('DocumentApp.getActiveDocument() was never called for an invalid documentId', activeDocCalls, 0);
  check('PropertiesService.getDocumentProperties() was never called for an invalid documentId', docPropsCalls, 0);
}

// ---------------------------------------------------------------- structural proof

console.log('continuousUpdateForDocument_() -- structurally never uses getActiveDocument/getDocumentProperties/getDocumentLock');

{
  const { sandbox } = loadCode();
  setUpAdoptedEnabledDocument(sandbox, 'DOC_STRUCT');

  let activeDocCalls = 0;
  let docPropsCalls = 0;
  let docLockCalls = 0;
  let openByIdCalls = [];

  sandbox.DocumentApp.getActiveDocument = () => { activeDocCalls++; throw new Error('must not be called from background execution'); };
  sandbox.DocumentApp.openById = (id) => { openByIdCalls.push(id); return makeFakeDocument(id); };

  const realGetDocProps = sandbox.PropertiesService.getDocumentProperties;
  sandbox.PropertiesService.getDocumentProperties = () => { docPropsCalls++; return realGetDocProps(); };

  sandbox.LockService.getDocumentLock = () => { docLockCalls++; throw new Error('must not be called from background execution'); };
  sandbox.downloadAndGroupTdocs_ = () => ({}); // isolate: this test is about structural access, not fetch behavior

  const result = sandbox.continuousUpdateForDocument_('DOC_STRUCT');

  check('completed successfully', result.success, true);
  check('DocumentApp.getActiveDocument() was NEVER called', activeDocCalls, 0);
  check('PropertiesService.getDocumentProperties() was NEVER called', docPropsCalls, 0);
  check('LockService.getDocumentLock() was NEVER called', docLockCalls, 0);
  check('the correct document was opened by id, exactly once',
    openByIdCalls, ['DOC_STRUCT']);
}

// ---------------------------------------------------------------- correct document

console.log('continuousUpdateForDocument_() -- opens the correct document, not some other one');

{
  const { sandbox } = loadCode();
  setUpAdoptedEnabledDocument(sandbox, 'DOC_A');
  setUpAdoptedEnabledDocument(sandbox, 'DOC_B');

  let openedIds = [];
  sandbox.DocumentApp.openById = (id) => { openedIds.push(id); return makeFakeDocument(id); };
  sandbox.downloadAndGroupTdocs_ = () => ({});

  sandbox.continuousUpdateForDocument_('DOC_B');

  check('only DOC_B was opened, not DOC_A', openedIds, ['DOC_B']);
}

// ---------------------------------------------------------------- config from central state

console.log('continuousUpdateForDocument_() -- all required config comes from central state, not Document Properties');

{
  const { sandbox, docProps } = loadCode({
    // A DIFFERENT value in (irrelevant, never-consulted) Document
    // Properties -- proves the run used central state, not this.
    documentProperties: { REPORT_SUFFIX: 'WRONG_VALUE_FROM_DOC_PROPS' }
  });
  setUpAdoptedEnabledDocument(sandbox, 'DOC_CFG', { REPORT_SUFFIX: 'Video' });
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => ({});

  // Spy on getReportConfig_ to capture what REPORT_SUFFIX it actually resolved.
  const realGetReportConfig = sandbox.getReportConfig_;
  let observedSuffix = null;
  sandbox.getReportConfig_ = function (context) {
    const cfg = realGetReportConfig(context);
    if (context && context.documentId === 'DOC_CFG') observedSuffix = cfg.REPORT_SUFFIX;
    return cfg;
  };

  sandbox.continuousUpdateForDocument_('DOC_CFG');

  check('the run read REPORT_SUFFIX from central state ("Video"), not Document Properties',
    observedSuffix, 'Video');
  check('Document Properties (irrelevant here) was never mutated',
    docProps.getProperty('REPORT_SUFFIX'), 'WRONG_VALUE_FROM_DOC_PROPS');
}

// ---------------------------------------------------------------- cache writes

console.log('continuousUpdateForDocument_() -- background cache/collector state writes go to the CENTRAL backend');

{
  const { sandbox } = loadCode();
  setUpAdoptedEnabledDocument(sandbox, 'DOC_CACHE');
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  // Stub out the TDoc-list xlsx download/conversion pipeline itself (its
  // own offline behavior -- raw UrlFetchApp.fetch + DriveApp + SpreadsheetApp
  // -- is exercised by other test files, not this one): this test is about
  // where rearrangeRevisionTables_()'s REVISION_MAP write lands, which runs
  // regardless of what downloadAndGroupTdocs_() returns.
  sandbox.downloadAndGroupTdocs_ = () => ({});

  sandbox.continuousUpdateForDocument_('DOC_CACHE');

  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_CACHE' });
  check('REVISION_MAP was (re)computed and written to the CENTRAL backend',
    central.getProperty('REVISION_MAP') !== null, true);
}

// ---------------------------------------------------------------- lastRunAt

console.log('continuousUpdateForDocument_() -- successful run updates lastRunAt');

{
  const { sandbox } = loadCode();
  setUpAdoptedEnabledDocument(sandbox, 'DOC_SUCCESS');
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => ({});

  const before = sandbox.getRegisteredReportDocument_('DOC_SUCCESS').lastRunAt;
  check('lastRunAt starts null', before, null);

  const result = sandbox.continuousUpdateForDocument_('DOC_SUCCESS');

  check('result reports success', result.success, true);
  const after = sandbox.getRegisteredReportDocument_('DOC_SUCCESS').lastRunAt;
  check('lastRunAt is now set', after !== null, true);
  check('result.lastRunAt matches the registry', result.lastRunAt, after);
}

console.log('continuousUpdateForDocument_() -- an INTERNAL per-stage failure (e.g. a bad TDoc-list URL) is caught inside continuousUpdateCore_(), but still does NOT update lastRunAt');

{
  // This is the exact class of failure a real TDOC_LIST_URL pointing at
  // an unresolvable host (e.g. https://example.invalid/...) produces in
  // production: downloadAndGroupTdocs_()'s raw UrlFetchApp.fetch() call
  // throws, that throw is caught by continuousUpdateCore_()'s OWN try/
  // catch (never propagates on its own), and continuousUpdateCore_()
  // returns {success:false, error}. Simulated directly here via a
  // throwing downloadAndGroupTdocs_ stub, rather than a real network
  // call, so this test has no external dependency -- the throw site
  // (inside continuousUpdateCore_()'s try block, not before it) is what
  // matters, not the specific reason.
  const { sandbox } = loadCode();
  setUpAdoptedEnabledDocument(sandbox, 'DOC_INTERNAL_FAIL');
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => { throw new Error('Could not download TDOC list: fetch failed (simulated example.invalid)'); };

  const before = sandbox.getRegisteredReportDocument_('DOC_INTERNAL_FAIL').lastRunAt;
  check('lastRunAt starts null', before, null);

  expectThrows('continuousUpdateForDocument_ throws for an internally-caught core failure (not silently "successful")',
    () => sandbox.continuousUpdateForDocument_('DOC_INTERNAL_FAIL'),
    'did not complete successfully');

  check('lastRunAt remains null after an internally-caught failure',
    sandbox.getRegisteredReportDocument_('DOC_INTERNAL_FAIL').lastRunAt, null);
}

console.log('continuousUpdateForDocument_() -- a failed run (setup-phase throw) does NOT update lastRunAt');

{
  const { sandbox } = loadCode();
  setUpAdoptedEnabledDocument(sandbox, 'DOC_FAIL');
  // Force a setup-phase failure: getReportBody_(context) throws because
  // the document cannot be opened (e.g. deleted/inaccessible file).
  sandbox.DocumentApp.openById = () => { throw new Error('Document not found (simulated)'); };

  expectThrows('continuousUpdateForDocument_ propagates the setup-phase failure',
    () => sandbox.continuousUpdateForDocument_('DOC_FAIL'), 'Document not found');

  check('lastRunAt remains null after a failed run',
    sandbox.getRegisteredReportDocument_('DOC_FAIL').lastRunAt, null);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
