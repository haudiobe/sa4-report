/**
 * ADDON-002 — report state store: getReportStateStore_() and its two
 * backends (Document Properties / central Script-Properties-by-documentId).
 *
 * See the "ADDON-002 -- REPORT STATE STORE" block in Code.js for the full
 * rationale. Nothing in production calls getReportStateStore_() with a
 * context that selects Backend B yet -- this file proves the abstraction
 * itself is correct in isolation, ahead of any real caller depending on it.
 *
 * Run: node tests/addon002-state-store.test.js
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

// ============================================================ Backend A ===

console.log('getReportStateStore_() -- Backend A (Document Properties) is the default');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { FOO: 'bar' } });

  check('no context -> the exact same object PropertiesService.getDocumentProperties() returns',
    sandbox.getReportStateStore_() === sandbox.PropertiesService.getDocumentProperties(), true);
  check('reads existing Document Properties unchanged',
    sandbox.getReportStateStore_().getProperty('FOO'), 'bar');
}

['bound', 'addon-interactive', undefined].forEach((mode) => {
  const { sandbox } = loadCode({ documentProperties: { KEY: 'val' } });
  const store = sandbox.getReportStateStore_(mode === undefined ? {} : { mode: mode });
  check(`mode=${mode} -> Backend A (real Document Properties)`,
    store === sandbox.PropertiesService.getDocumentProperties(), true);
});

{
  // A context carrying document/documentId but NOT addon-background mode
  // still uses Backend A -- ADDON-001B's finding that openById() does not
  // establish a Properties context means Backend A must ignore
  // document/documentId entirely and always defer to the real, currently
  // executing document's properties.
  const { sandbox } = loadCode({ documentProperties: { KEY: 'real-doc-value' } });
  const store = sandbox.getReportStateStore_({ documentId: 'SOME_OTHER_DOC', mode: 'bound' });
  check('mode=bound with an unrelated documentId still reads the REAL current document\'s properties',
    store.getProperty('KEY'), 'real-doc-value');
}

// ============================================================ Backend B ===

console.log('getReportStateStore_() -- Backend B (central, documentId-namespaced Script Properties)');

{
  const { sandbox } = loadCode();
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_A' });

  check('get/set round-trip',
    (store.setProperty('KEY1', 'value1'), store.getProperty('KEY1')), 'value1');
  check('missing key returns null (matches real Properties behavior)',
    store.getProperty('NEVER_SET'), null);

  store.setProperty('KEY2', 'value2');
  store.deleteProperty('KEY1');
  check('deleteProperty() removes only the targeted key',
    [store.getProperty('KEY1'), store.getProperty('KEY2')], [null, 'value2']);
}

console.log('getReportStateStore_() -- Backend B: document isolation');

{
  const { sandbox } = loadCode();
  const storeA = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_A' });
  const storeB = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_B' });

  storeA.setProperty('MEETING_ID', '111');
  storeB.setProperty('MEETING_ID', '222');

  check('document A and document B do not see each other\'s values for the same key name',
    [storeA.getProperty('MEETING_ID'), storeB.getProperty('MEETING_ID')], ['111', '222']);

  storeA.setProperty('ONLY_IN_A', 'x');
  check('document B does not see a key only ever set in document A',
    storeB.getProperty('ONLY_IN_A'), null);

  check('document A\'s getKeys() does not include document B\'s keys',
    storeA.getKeys().indexOf('MEETING_ID') !== -1 && storeA.getKeys().length, 2);

  storeA.deleteProperty('MEETING_ID');
  check('deleting a key in document A does not affect document B',
    storeB.getProperty('MEETING_ID'), '222');
}

console.log('getReportStateStore_() -- Backend B: getProperties()/setProperties()');

{
  const { sandbox } = loadCode();
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_C' });

  store.setProperty('A', '1');
  store.setProperty('B', '2');
  check('getProperties() returns every key in THIS document\'s namespace, unprefixed',
    store.getProperties(), { A: '1', B: '2' });

  store.setProperties({ B: '20', C: '3' });
  check('setProperties() merges by default (deleteAllOthers omitted/false)',
    store.getProperties(), { A: '1', B: '20', C: '3' });

  store.setProperties({ D: '4' }, true);
  check('setProperties(props, true) clears every other key in this namespace first',
    store.getProperties(), { D: '4' });
}

console.log('getReportStateStore_() -- Backend B: raw values only, no hidden JSON layer');

{
  const { sandbox } = loadCode();
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_D' });

  // The abstraction stores each ORIGINAL key as its own Script Property
  // (see the ADDON-002 report's capacity analysis for why -- not one
  // combined JSON blob). It never parses or re-serializes the value it is
  // given, so a value that happens not to be valid JSON round-trips
  // exactly, the same as real Document/Script Properties -- callers that
  // need JSON (loadJsonObject_() and friends, unchanged) remain solely
  // responsible for parsing, exactly as they are for Backend A today.
  store.setProperty('NOT_JSON', 'plain string, not { valid json');
  check('a non-JSON value round-trips unchanged (store never attempts to parse it)',
    store.getProperty('NOT_JSON'), 'plain string, not { valid json');

  store.setProperty('SOME_CACHE_ENTRY', '{"ts":123,"empty":true}');
  check('a JSON-shaped value also round-trips unchanged, as an opaque string',
    store.getProperty('SOME_CACHE_ENTRY'), '{"ts":123,"empty":true}');
}

console.log('getReportStateStore_() -- Backend B: deterministic namespace, no collisions with global keys');

{
  const { sandbox, scriptProps } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'secret-token' } });
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_E' });

  store.setProperty('REVIEWER_API_TOKEN', 'should-not-touch-global');

  check('the global REVIEWER_API_TOKEN Script Property is untouched by a namespaced key of the same name',
    scriptProps.getProperty('REVIEWER_API_TOKEN'), 'secret-token');
  check('the namespaced key is stored separately, under this document\'s prefix',
    scriptProps.getKeys().some(k => k.indexOf('DOC_E') !== -1 && k.indexOf('REVIEWER_API_TOKEN') !== -1), true);

  const storeAgain = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_E' });
  check('the same documentId always maps to the same namespace (deterministic, not per-call-random)',
    storeAgain.getProperty('REVIEWER_API_TOKEN'), 'should-not-touch-global');
}

console.log('getReportStateStore_() -- Backend B: falls back to context.document.getId() when documentId is absent');

{
  const { sandbox } = loadCode();
  const fakeDoc = { getId: () => 'DOC_FROM_DOCUMENT_OBJECT' };
  const store = sandbox.getReportStateStore_({ mode: 'addon-background', document: fakeDoc });
  store.setProperty('K', 'v');

  const storeById = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_FROM_DOCUMENT_OBJECT' });
  check('a context with only `document` resolves to the same namespace as the equivalent documentId',
    storeById.getProperty('K'), 'v');
}

console.log('getReportStateStore_() -- Backend B: requires a documentId one way or another');

{
  const { sandbox } = loadCode();
  expectThrows('addon-background mode with neither documentId nor document throws clearly',
    () => sandbox.getReportStateStore_({ mode: 'addon-background' }),
    'documentId is required');
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
