/**
 * ADDON-004 — getCentralStateUsage_(): read-only diagnostic for how much
 * of the shared Script Properties budget is actual central report state,
 * broken down per document. Never changes any state.
 *
 * Run: node tests/addon004-capacity-diagnostics.test.js
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

console.log('getCentralStateUsage_() -- empty state');

{
  const { sandbox } = loadCode();
  const usage = sandbox.getCentralStateUsage_();
  check('no keys -> zeroed-out usage report',
    usage, { totalBytesApprox: 0, keyCount: 0, documentCount: 0, perDocument: [] });
}

console.log('getCentralStateUsage_() -- counts only SA4_STATE|... keys, never the registry or global secrets');

{
  const { sandbox, scriptProps } = loadCode();
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_A' });
  store.setProperty('MEETING_ID', '123');

  sandbox.registerReportDocument_('DOC_A', {}); // registry entry -- must NOT be counted
  scriptProps.setProperty('REVIEWER_API_TOKEN', 'super-secret'); // global secret -- must NOT be counted

  const usage = sandbox.getCentralStateUsage_();
  check('only the one report-state key is counted', usage.keyCount, 1);
  check('exactly one document is represented', usage.documentCount, 1);
  check('the registry key itself does not appear in perDocument',
    usage.perDocument.every(d => d.documentId !== 'SA4_REGISTRY_INDEX' && !d.documentId.includes('REVIEWER')), true);
}

console.log('getCentralStateUsage_() -- per-document breakdown, sorted largest-first');

{
  const { sandbox } = loadCode();
  const storeSmall = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_SMALL' });
  storeSmall.setProperty('K', 'x');

  const storeBig = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_BIG' });
  storeBig.setProperty('K1', 'x'.repeat(500));
  storeBig.setProperty('K2', 'y'.repeat(500));

  const usage = sandbox.getCentralStateUsage_();

  check('two documents represented', usage.documentCount, 2);
  check('sorted largest-usage-first', usage.perDocument.map(d => d.documentId), ['DOC_BIG', 'DOC_SMALL']);
  check('DOC_BIG has 2 keys', usage.perDocument[0].keyCount, 2);
  check('DOC_SMALL has 1 key', usage.perDocument[1].keyCount, 1);
  check('total bytes is the sum of every document\'s bytes',
    usage.totalBytesApprox, usage.perDocument.reduce((sum, d) => sum + d.bytesApprox, 0));
}

console.log('getCentralStateUsage_() -- document isolation: unregistering/deleting one document does not affect another\'s reported usage');

{
  const { sandbox } = loadCode();
  sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_X' }).setProperty('K', 'v');
  sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_Y' }).setProperty('K', 'v');

  sandbox.deleteCentralReportState_('DOC_X');

  const usage = sandbox.getCentralStateUsage_();
  check('only DOC_Y remains after DOC_X\'s state is deleted',
    usage.perDocument.map(d => d.documentId), ['DOC_Y']);
}

console.log('getCentralStateUsage_() -- read-only, changes nothing');

{
  const { sandbox, scriptProps } = loadCode();
  sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_Z' }).setProperty('K', 'v');

  const before = JSON.stringify(scriptProps._store);
  sandbox.getCentralStateUsage_();
  sandbox.getCentralStateUsage_();
  const after = JSON.stringify(scriptProps._store);

  check('Script Properties store is byte-identical before/after two calls',
    after, before);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
