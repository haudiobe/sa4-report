/**
 * ADDON-003 — cleanup semantics (unregisterReportDocument_() vs. the
 * separate, explicit, destructive deleteCentralReportState_()) and a
 * capacity check proving the registry's "index + one property per
 * document" representation never needs one unbounded JSON blob.
 *
 * Run: node tests/addon003-cleanup-capacity.test.js
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

function withActiveDocument(sandbox, documentId) {
  sandbox.DocumentApp.getActiveDocument = () => ({
    getId: () => documentId,
    getBody: () => ({})
  });
}

// ---------------------------------------------------------------- cleanup

console.log('unregisterReportDocument_() -- does NOT delete central report state');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X', REVISION_MAP: '{"a":"b"}' } });
  withActiveDocument(sandbox, 'DOC_CLEAN_1');

  sandbox.adoptReportDocumentForAddon_();
  sandbox.unregisterReportDocument_('DOC_CLEAN_1');

  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_CLEAN_1' });
  check('registry entry is gone',
    sandbox.getRegisteredReportDocument_('DOC_CLEAN_1'), null);
  check('central report state SURVIVES unregistration',
    central.getProperty('REVISION_MAP'), '{"a":"b"}');
}

console.log('unregisterReportDocument_() -- does NOT touch Document Properties');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  withActiveDocument(sandbox, 'DOC_CLEAN_2');

  sandbox.adoptReportDocumentForAddon_();
  const before = JSON.stringify(docProps._store);
  sandbox.unregisterReportDocument_('DOC_CLEAN_2');
  const after = JSON.stringify(docProps._store);

  check('Document Properties untouched by unregister',
    after, before);
}

console.log('deleteCentralReportState_() -- explicit, destructive, affects only the target document');

{
  const { sandbox } = loadCode({
    documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X', REVISION_MAP: '{"a":"b"}' }
  });

  withActiveDocument(sandbox, 'DOC_DEL_A');
  sandbox.adoptReportDocumentForAddon_();

  withActiveDocument(sandbox, 'DOC_DEL_B');
  sandbox.adoptReportDocumentForAddon_();

  sandbox.deleteCentralReportState_('DOC_DEL_A');

  const storeA = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_DEL_A' });
  const storeB = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_DEL_B' });

  check('document A\'s central state is fully cleared',
    storeA.getKeys(), []);
  check('document B\'s central state is UNTOUCHED',
    storeB.getProperty('REVISION_MAP'), '{"a":"b"}');
}

console.log('deleteCentralReportState_() -- does NOT remove the registry entry (orthogonal to unregister)');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  withActiveDocument(sandbox, 'DOC_DEL_C');
  sandbox.adoptReportDocumentForAddon_();

  sandbox.deleteCentralReportState_('DOC_DEL_C');

  check('registry entry still exists after deleteCentralReportState_() alone',
    sandbox.getRegisteredReportDocument_('DOC_DEL_C') !== null, true);
}

// --------------------------------------------------------------- capacity

console.log('registry representation -- no single unbounded JSON blob, even with many documents');

{
  const { sandbox, scriptProps } = loadCode();

  const N = 60; // comfortably more than any realistic number of concurrently-tracked SA4 meeting documents
  for (let i = 0; i < N; i++) {
    sandbox.registerReportDocument_('DOC_CAP_' + i, {
      meetingId: String(90000 + i),
      meetingName: 'Synthetic Meeting ' + i,
      intervalHours: 1
    });
  }

  check(`all ${N} documents are listed`,
    sandbox.listRegisteredReportDocuments_().length, N);

  const indexValue = scriptProps.getProperty('SA4_REGISTRY_INDEX');
  const indexBytes = Buffer.byteLength(indexValue, 'utf8');
  check(`the index property alone (just documentIds) stays well under the 9KB/value limit for ${N} documents`,
    indexBytes < 9 * 1024, true);

  // The key structural property: registering MORE documents never grows
  // any EXISTING document's own entry -- each document's property is
  // sized by ITS OWN fields only, independent of how many other documents
  // are registered. Prove this by comparing one document's entry size
  // before/after registering 60 more.
  const beforeEntry = scriptProps.getProperty('SA4_REGISTRY_DOC|DOC_CAP_0');
  for (let i = N; i < N + 60; i++) {
    sandbox.registerReportDocument_('DOC_CAP_' + i, { meetingId: String(i) });
  }
  const afterEntry = scriptProps.getProperty('SA4_REGISTRY_DOC|DOC_CAP_0');

  check('DOC_CAP_0\'s own entry is byte-identical after 60 MORE documents are registered (no shared/growing blob)',
    afterEntry, beforeEntry);

  // Total script-properties footprint check: every key (index + N+60
  // per-document entries) individually well under 9KB, and the SUM well
  // under the 500KB/store ceiling for this realistic scale.
  const allKeys = scriptProps.getKeys();
  let totalBytes = 0;
  let maxSingleValueBytes = 0;
  allKeys.forEach((k) => {
    const v = scriptProps.getProperty(k) || '';
    const bytes = Buffer.byteLength(k, 'utf8') + Buffer.byteLength(v, 'utf8');
    totalBytes += bytes;
    if (bytes > maxSingleValueBytes) maxSingleValueBytes = bytes;
  });

  check('every individual registry property stays under the 9KB/value limit',
    maxSingleValueBytes < 9 * 1024, true);
  check('total registry footprint for 120 documents stays well under the 500KB/store limit',
    totalBytes < 500 * 1024, true);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
