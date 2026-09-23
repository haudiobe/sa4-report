/**
 * ADDON-002 — proves getReportConfig_(context)/getMeetingContext_(context)
 * produce EQUIVALENT results whether their state comes from Backend A
 * (Document Properties, today's only real backend) or Backend B (central,
 * documentId-namespaced Script Properties, not wired into production
 * anywhere yet).
 *
 * This is the integration proof that the ADDON-002 abstraction is
 * sufficient to support a future background execution path: the SAME
 * meeting configuration, read through either backend, must yield the SAME
 * derived config/context. It does not prove Backend B is populated with
 * real data anywhere today -- that migration is explicitly out of scope
 * (see the ADDON-002 report, section 3: "one logical source of truth", no
 * copying of live state).
 *
 * Run: node tests/addon002-config-parity.test.js
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

// A synthetic meeting configuration, covering every property key
// getReportConfig_()/getMeetingIdentityConfig_() read (see the ADDON-002
// report's property inventory, category "meeting configuration" +
// "report configuration" + "agenda/cache state" identity fields).
const SYNTHETIC_CONFIG = {
  MEETING_FOLDER: 'TSGS4_999_Synthetic',
  MEETING_NUMBER: '999',
  MEETING_ID: '99999',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_999_Synthetic/Docs/',
  TDOC_LIST_URL: 'https://example.invalid/TDoc_List_Meeting_SA4%23999.xlsx',
  REPORT_SUFFIX: 'Audio',
  AGENDA_ITEM_PREFIX: '7',
  AGENDA_SOURCE_DOC_ID: 'SYNTHETIC_AGENDA_DOC_ID',
  AGENDA_TDOC: 'S4-999001',
  MEETING_DATE: '2026-10-01',
  SHOW_PREVIEW_SNIPPET: 'false',
  MEETING_TYPE: 'main',
  MEETING_NAME: '',
  REVISIONS_URL: ''
};

function populateBackendB(sandbox, documentId, config) {
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: documentId });
  Object.keys(config).forEach((key) => store.setProperty(key, config[key]));
  return store;
}

console.log('getReportConfig_(context) -- Backend A vs. Backend B produce identical config for the same data');

{
  const { sandbox } = loadCode({ documentProperties: SYNTHETIC_CONFIG });
  const fromBackendA = sandbox.getReportConfig_();

  populateBackendB(sandbox, 'DOC_PARITY_1', SYNTHETIC_CONFIG);
  const fromBackendB = sandbox.getReportConfig_({ mode: 'addon-background', documentId: 'DOC_PARITY_1' });

  check('getReportConfig_() result is identical across both backends for the same underlying values',
    fromBackendB, fromBackendA);
}

console.log('getMeetingContext_(context) -- Backend A vs. Backend B produce identical context for the same data');

{
  const { sandbox } = loadCode({ documentProperties: SYNTHETIC_CONFIG });
  const fromBackendA = sandbox.getMeetingContext_();

  populateBackendB(sandbox, 'DOC_PARITY_2', SYNTHETIC_CONFIG);
  const fromBackendB = sandbox.getMeetingContext_({ mode: 'addon-background', documentId: 'DOC_PARITY_2' });

  check('getMeetingContext_() result is identical across both backends for the same underlying values',
    fromBackendB, fromBackendA);
}

console.log('getReportConfig_(context)/getMeetingContext_(context) -- parity also holds for an EMPTY (default) config');

{
  const { sandbox } = loadCode();
  const fromBackendA = { cfg: sandbox.getReportConfig_(), ctx: sandbox.getMeetingContext_() };

  // No properties ever set for this document -- every field must fall
  // back to the exact same defaults on both backends, since Backend B's
  // getProperty() returns null for a missing key exactly like real
  // Document Properties does.
  const fromBackendB = {
    cfg: sandbox.getReportConfig_({ mode: 'addon-background', documentId: 'DOC_PARITY_EMPTY' }),
    ctx: sandbox.getMeetingContext_({ mode: 'addon-background', documentId: 'DOC_PARITY_EMPTY' })
  };

  check('default (all-absent) config is identical across both backends',
    fromBackendB, fromBackendA);
}

console.log('getReportConfig_(context) -- REVIEWER_API_TOKEN stays global (Script Properties), not namespaced, under either backend');

{
  const { sandbox } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'shared-token' } });

  const cfgBackendA = sandbox.getReportConfig_();
  const cfgBackendB = sandbox.getReportConfig_({ mode: 'addon-background', documentId: 'DOC_TOKEN_TEST' });

  check('Backend A config sees the real global token',
    cfgBackendA.REVIEWER_API_TOKEN, 'shared-token');
  check('Backend B config ALSO sees the same real global token (the secret is never per-document)',
    cfgBackendB.REVIEWER_API_TOKEN, 'shared-token');
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
