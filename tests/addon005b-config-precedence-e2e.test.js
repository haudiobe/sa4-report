/**
 * ADDON-005B — end-to-end trace of the path the live acceptance test
 * exercised and questioned:
 *
 *   Configure Meeting Settings -> Document Properties -> Save
 *   -> Enable Automatic Updates (this doc) -> adoption/re-adoption
 *   -> central registry/state -> scheduler lookup
 *   -> continuousUpdateForDocument_(documentId) -> report context
 *   -> getReportConfig_() -> MeetingContext/source resolution
 *   -> downloadAndGroupTdocs_() -> actual fetch URL
 *
 * Live symptom: after saving an explicit TDOC_LIST_URL and re-running
 * "Enable Automatic Updates" to re-sync, the scheduler still fetched the
 * computed SA4#136 default URL instead.
 *
 * Root-cause trace (see the ADDON-005B report for the full writeup):
 * reproducing this EXACT sequence against current source (which already
 * includes the prior "adoption always reads real Document Properties"
 * fix, commit ae925ae) does NOT reproduce the bug -- the explicit URL
 * correctly reaches downloadAndGroupTdocs_(). The live failure is most
 * consistent with the installed test deployment not yet reflecting that
 * already-pushed fix at the moment of the live menu clicks (Phase K's
 * still-unverified update-propagation question), not a surviving code
 * defect. This suite exists as a PERMANENT regression guard for the
 * exact chain in question, regardless of that explanation -- if this
 * chain silently breaks again for any reason (a future refactor,
 * whatever), this suite fails immediately rather than only being
 * discoverable live.
 *
 * Run: node tests/addon005b-config-precedence-e2e.test.js
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

const SCRATCH_URL = 'https://example.invalid/scratch-test-tdoc-list.xlsx';

function withActiveDocument(sandbox, documentId) {
  sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => documentId, getBody: () => ({ getTables: () => [] }) });
}

function alertRecorder(sandbox) {
  sandbox.DocumentApp.getUi = () => ({ alert: () => {}, ButtonSet: { OK: 'OK' } });
}

console.log('end-to-end: Configure Meeting Settings -> Save -> Enable -> [edit] -> Enable again -> scheduler run -> downloadAndGroupTdocs_ receives the explicit URL, not the computed default');

{
  const { sandbox, docProps } = loadCode();
  const documentId = 'DOC_E2E';
  withActiveDocument(sandbox, documentId);
  alertRecorder(sandbox);
  sandbox.ScriptApp = {
    getProjectTriggers: () => [],
    newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create: () => ({ getUniqueId: () => 'u1' }) }) }) }),
    deleteTrigger: () => {}
  };

  // Step 1: "Configure Meeting Settings" -> Save, with NO explicit TDoc
  // URL yet -- mirrors saveConfigurationSettings()'s own behavior for an
  // omitted field (deletes any existing override, "will auto-detect").
  sandbox.saveConfigurationSettings({
    meetingFolder: 'TSGS4_136_Montreal',
    meetingNumber: '136',
    meetingId: '60777',
    reportType: '6G'
    // tdocUrl intentionally omitted
  });

  check('1. Document Properties has NO explicit TDoc URL yet (auto-detect)',
    docProps.getProperty('TDOC_LIST_URL'), null);

  // Step 2: "Enable Automatic Updates" -- first-ever adoption.
  sandbox.enableAutomaticUpdatesForAddon();

  const centralAfterFirstEnable = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: documentId });
  check('2. initial adoption stored the (absent) TDOC_LIST_URL centrally as absent too (no invented value)',
    centralAfterFirstEnable.getProperty('TDOC_LIST_URL'), null);

  // Step 3: "Configure Meeting Settings" -> Save, THIS time with an
  // explicit scratch TDoc URL (the live test's actual scenario).
  sandbox.saveConfigurationSettings({
    meetingFolder: 'TSGS4_136_Montreal',
    meetingNumber: '136',
    meetingId: '60777',
    reportType: '6G',
    tdocUrl: SCRATCH_URL
  });

  check('3a. Document Properties now has the explicit scratch URL',
    docProps.getProperty('TDOC_LIST_URL'), SCRATCH_URL);
  check('3b. central state does NOT yet have it (adoption not re-run yet)',
    centralAfterFirstEnable.getProperty('TDOC_LIST_URL'), null);

  // Step 4: "Enable Automatic Updates" AGAIN, specifically to re-sync.
  sandbox.enableAutomaticUpdatesForAddon();

  const centralAfterSecondEnable = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: documentId });
  check('4. re-adoption updated central state to the explicit scratch URL',
    centralAfterSecondEnable.getProperty('TDOC_LIST_URL'), SCRATCH_URL);

  // Step 5: the scheduler's per-document unit of work.
  sandbox.DocumentApp.openById = (id) => ({ getId: () => id, getBody: () => ({ getTables: () => [] }) });
  let capturedCfg = null;
  sandbox.downloadAndGroupTdocs_ = (cfg, context) => { capturedCfg = cfg; return {}; };

  const result = sandbox.continuousUpdateForDocument_(documentId);

  check('5. the run succeeded', result.success, true);
  check('6. the URL reaching downloadAndGroupTdocs_() is EXACTLY the explicit scratch URL, not the computed SA4#136 default',
    capturedCfg && capturedCfg.TDOC_LIST_URL, SCRATCH_URL);
  check('6b. specifically NOT the computed default that caused the live bug',
    capturedCfg.TDOC_LIST_URL !== 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4%23136.xlsx', true);
}

console.log('explicit URL precedence over the computed default remains intact for a document that never explicitly sets one');

{
  const { sandbox, docProps } = loadCode();
  const documentId = 'DOC_DEFAULT';
  withActiveDocument(sandbox, documentId);
  alertRecorder(sandbox);
  sandbox.ScriptApp = { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create: () => ({ getUniqueId: () => 'u2' }) }) }) }), deleteTrigger: () => {} };

  sandbox.saveConfigurationSettings({ meetingFolder: 'TSGS4_136_Montreal', meetingNumber: '136', meetingId: '60777', reportType: '6G' });
  sandbox.enableAutomaticUpdatesForAddon();

  sandbox.DocumentApp.openById = (id) => ({ getId: () => id, getBody: () => ({ getTables: () => [] }) });
  let capturedCfg = null;
  sandbox.downloadAndGroupTdocs_ = (cfg) => { capturedCfg = cfg; return {}; };
  sandbox.continuousUpdateForDocument_(documentId);

  check('with no explicit override anywhere in the chain, the computed default formula still applies correctly',
    capturedCfg.TDOC_LIST_URL,
    'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4%23136.xlsx');
}

console.log('direct check: adoptReportDocumentForAddon_() always reads Document Properties as its source, never the (possibly-already-central) addon-interactive seam');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function adoptReportDocumentForAddon_(');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  check('adoption\'s source is the literal, direct PropertiesService.getDocumentProperties() call',
    /const source = PropertiesService\.getDocumentProperties\(\);/.test(body), true);
  check('adoption\'s source is NOT routed through getReportStateStore_(context)',
    /const source = getReportStateStore_\(context\);/.test(body), false);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
