/**
 * ADDON-004 — authoritative state selection: exactly ONE backend per
 * document, for each of the three execution cases (legacy bound,
 * add-on interactive, add-on background). See the extensive header
 * comment on getReportStateStore_() in Code.js for the exact rule.
 *
 * Run: node tests/addon004-authoritative-state.test.js
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

// -------------------------------------------------------- legacy/no-context

console.log('legacy bound / no context -> Document Properties, unconditionally');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: 'legacy-value' } });
  const store = sandbox.getReportStateStore_();
  check('reads real Document Properties',
    store.getProperty('MEETING_ID'), 'legacy-value');
  check('is literally the same object PropertiesService.getDocumentProperties() returns',
    store === sandbox.PropertiesService.getDocumentProperties(), true);
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: 'legacy-value' } });
  const store = sandbox.getReportStateStore_({ mode: 'bound' });
  check('mode="bound" behaves identically to no context',
    store === sandbox.PropertiesService.getDocumentProperties(), true);
}

// ------------------------------------------- add-on interactive, unregistered

console.log('add-on interactive, document NEVER adopted -> Document Properties (explicit, not silent data loss)');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_ID: 'never-adopted-value' } });
  withActiveDocument(sandbox, 'DOC_UNREGISTERED');

  const store = sandbox.getReportStateStore_({ mode: 'addon-interactive' });
  check('reads the real Document Properties value',
    store.getProperty('MEETING_ID'), 'never-adopted-value');
  check('is the same object as Document Properties (not an empty central store)',
    store === sandbox.PropertiesService.getDocumentProperties(), true);

  store.setProperty('NEW_KEY', 'written-interactively');
  check('a write through this store lands in real Document Properties',
    docProps.getProperty('NEW_KEY'), 'written-interactively');

  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_UNREGISTERED' });
  check('the write did NOT also land in the central backend (no dual-write, no silent switch)',
    central.getProperty('NEW_KEY'), null);
}

// ----------------------------------------------- add-on interactive, adopted

console.log('add-on interactive, document IS adopted -> central state, authoritative');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_ID: 'source-value', MEETING_FOLDER: 'X' } });
  withActiveDocument(sandbox, 'DOC_ADOPTED');
  sandbox.adoptReportDocumentForAddon_();

  const store = sandbox.getReportStateStore_({ mode: 'addon-interactive' });
  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_ADOPTED' });

  check('addon-interactive on an adopted document resolves to the SAME backend as addon-background',
    store === central || store.getProperty('MEETING_ID') === central.getProperty('MEETING_ID'), true);
  check('addon-interactive reads the value that was copied in at adoption time',
    store.getProperty('MEETING_ID'), 'source-value');

  store.setProperty('INTERACTIVE_EDIT', 'made-after-adoption');
  check('an interactive write on an adopted document lands in the CENTRAL backend',
    central.getProperty('INTERACTIVE_EDIT'), 'made-after-adoption');
  check('an interactive write on an adopted document does NOT also land in Document Properties (no dual-write)',
    docProps.getProperty('INTERACTIVE_EDIT'), null);
  check('Document Properties remains exactly as it was at adoption time (legacy/archive fallback only)',
    docProps.getProperty('MEETING_ID'), 'source-value');
}

// ------------------------------------------------------- add-on background

console.log('add-on background -> central state, always (registration status is the CALLER\'s concern, not this function\'s)');

{
  const { sandbox } = loadCode();
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_BG_ONLY' });
  store.setProperty('K', 'v');
  check('background mode writes to the central, documentId-namespaced backend',
    store.getProperty('K'), 'v');
}

// ---------------------------------------------------------- no dual-write

console.log('no dual-write, in either direction, for an adopted document across repeated interactive + background use');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  withActiveDocument(sandbox, 'DOC_NODUAL');
  sandbox.adoptReportDocumentForAddon_();

  const interactiveStore = sandbox.getReportStateStore_({ mode: 'addon-interactive' });
  const backgroundStore = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_NODUAL' });

  interactiveStore.setProperty('FROM_INTERACTIVE', 'a');
  backgroundStore.setProperty('FROM_BACKGROUND', 'b');

  check('both writes are visible from either mode (same underlying backend)',
    [interactiveStore.getProperty('FROM_BACKGROUND'), backgroundStore.getProperty('FROM_INTERACTIVE')],
    ['b', 'a']);

  check('NEITHER write touched Document Properties',
    [docProps.getProperty('FROM_INTERACTIVE'), docProps.getProperty('FROM_BACKGROUND')],
    [null, null]);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
