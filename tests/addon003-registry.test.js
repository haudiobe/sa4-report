/**
 * ADDON-003 — central document registry: registerReportDocument_() /
 * getRegisteredReportDocument_() / listRegisteredReportDocuments_() /
 * updateRegisteredReportDocument_() / unregisterReportDocument_().
 *
 * Separate Script-Properties namespace from ADDON-002's
 * 'SA4_STATE|<documentId>|<key>' report state -- see the "ADDON-003 --
 * CENTRAL DOCUMENT REGISTRY" block in Code.js for the full rationale
 * (small index + one small property per document, not one blob).
 *
 * Nothing here creates a trigger; the registry is pure bookkeeping.
 *
 * Run: node tests/addon003-registry.test.js
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

// -------------------------------------------------------------- register

console.log('registerReportDocument_() -- register + retrieve');

{
  const { sandbox } = loadCode();
  const entry = sandbox.registerReportDocument_('DOC_1', { meetingId: '86178', meetingName: 'TSGS4_136_Montreal', intervalHours: 2 });

  check('register returns the stored entry', entry.documentId, 'DOC_1');
  check('meetingId is stored', entry.meetingId, '86178');
  check('intervalHours is stored', entry.intervalHours, 2);
  check('enabled defaults to false',
    entry.enabled, false);
  check('lastRunAt starts null (owned by the future scheduler, ADDON-004)',
    entry.lastRunAt, null);
  check('registeredAt is set (an ISO-ish string)',
    typeof entry.registeredAt === 'string' && entry.registeredAt.length > 0, true);

  const fetched = sandbox.getRegisteredReportDocument_('DOC_1');
  check('getRegisteredReportDocument_() returns the same entry',
    fetched, entry);
}

console.log('getRegisteredReportDocument_() -- unknown document returns null');

{
  const { sandbox } = loadCode();
  check('unregistered documentId -> null',
    sandbox.getRegisteredReportDocument_('NEVER_REGISTERED'), null);
  check('null/empty documentId -> null (no throw)',
    sandbox.getRegisteredReportDocument_(null), null);
}

// ------------------------------------------------------------------ list

console.log('listRegisteredReportDocuments_() -- lists every registered document');

{
  const { sandbox } = loadCode();
  sandbox.registerReportDocument_('DOC_A', { meetingId: '1' });
  sandbox.registerReportDocument_('DOC_B', { meetingId: '2' });
  sandbox.registerReportDocument_('DOC_C', { meetingId: '3' });

  const list = sandbox.listRegisteredReportDocuments_();
  check('list contains exactly the 3 registered documents',
    list.map(e => e.documentId).sort(), ['DOC_A', 'DOC_B', 'DOC_C']);
}

{
  const { sandbox } = loadCode();
  check('empty registry -> empty list, not an error',
    sandbox.listRegisteredReportDocuments_(), []);
}

// ---------------------------------------------------------------- update

console.log('updateRegisteredReportDocument_() -- patches an existing entry');

{
  const { sandbox } = loadCode();
  const original = sandbox.registerReportDocument_('DOC_U', { meetingId: '1', meetingName: 'Old Name', intervalHours: 1 });
  const updated = sandbox.updateRegisteredReportDocument_('DOC_U', { meetingName: 'New Name', enabled: true });

  check('patched field is updated', updated.meetingName, 'New Name');
  check('patched field is updated (enabled)', updated.enabled, true);
  check('un-patched field is preserved', updated.meetingId, '1');
  check('registeredAt is preserved across an update (not reset)',
    updated.registeredAt, original.registeredAt);
}

console.log('updateRegisteredReportDocument_() -- unknown document throws');

{
  const { sandbox } = loadCode();
  expectThrows('updating a never-registered document throws clearly',
    () => sandbox.updateRegisteredReportDocument_('GHOST_DOC', { enabled: true }),
    'is not registered');
}

// ------------------------------------------------------------- unregister

console.log('unregisterReportDocument_() -- removes from the registry');

{
  const { sandbox } = loadCode();
  sandbox.registerReportDocument_('DOC_R1', {});
  sandbox.registerReportDocument_('DOC_R2', {});

  sandbox.unregisterReportDocument_('DOC_R1');

  check('unregistered document is gone from getRegisteredReportDocument_()',
    sandbox.getRegisteredReportDocument_('DOC_R1'), null);
  check('unregistered document is gone from the list',
    sandbox.listRegisteredReportDocuments_().map(e => e.documentId), ['DOC_R2']);
  check('the OTHER registered document is untouched',
    sandbox.getRegisteredReportDocument_('DOC_R2') !== null, true);
}

console.log('unregisterReportDocument_() -- unregistering twice is safe (idempotent)');

{
  const { sandbox } = loadCode();
  sandbox.registerReportDocument_('DOC_R3', {});
  sandbox.unregisterReportDocument_('DOC_R3');
  // second call must not throw
  sandbox.unregisterReportDocument_('DOC_R3');
  check('still not registered after a second unregister call',
    sandbox.getRegisteredReportDocument_('DOC_R3'), null);
}

// ----------------------------------------------------- document isolation

console.log('registry -- document isolation (two documents never share fields)');

{
  const { sandbox } = loadCode();
  sandbox.registerReportDocument_('DOC_ISO_A', { meetingId: 'AAA', intervalHours: 1 });
  sandbox.registerReportDocument_('DOC_ISO_B', { meetingId: 'BBB', intervalHours: 6 });

  check('document A keeps its own meetingId',
    sandbox.getRegisteredReportDocument_('DOC_ISO_A').meetingId, 'AAA');
  check('document B keeps its own meetingId',
    sandbox.getRegisteredReportDocument_('DOC_ISO_B').meetingId, 'BBB');

  sandbox.updateRegisteredReportDocument_('DOC_ISO_A', { intervalHours: 24 });
  check('updating document A does not change document B',
    sandbox.getRegisteredReportDocument_('DOC_ISO_B').intervalHours, 6);
}

// --------------------------------------------- duplicate registration / idempotency

console.log('registerReportDocument_() -- registering the same document twice is idempotent (upsert)');

{
  const { sandbox } = loadCode();
  const first = sandbox.registerReportDocument_('DOC_DUP', { meetingId: '1', intervalHours: 2 });
  const second = sandbox.registerReportDocument_('DOC_DUP', { meetingId: '1', intervalHours: 2 });

  check('registeredAt is unchanged on a repeat registration with the same data',
    second.registeredAt, first.registeredAt);
  check('the index does not grow a duplicate entry',
    sandbox.listRegisteredReportDocuments_().filter(e => e.documentId === 'DOC_DUP').length, 1);
}

// ---------------------------------------------------- invalid interval rejected

console.log('registerReportDocument_() -- invalid intervalHours is rejected');

[15, 30, 0, -1, 3, 0.5, 'hourly'].forEach((bad) => {
  const { sandbox } = loadCode();
  expectThrows(`intervalHours=${JSON.stringify(bad)} is rejected (add-on triggers cannot run more than once per hour)`,
    () => sandbox.registerReportDocument_('DOC_BAD_INTERVAL', { intervalHours: bad }),
    'intervalHours must be one of');
});

console.log('registerReportDocument_() -- every documented allowed interval is accepted');

[1, 2, 4, 6, 12, 24].forEach((good) => {
  const { sandbox } = loadCode();
  const entry = sandbox.registerReportDocument_('DOC_GOOD_INTERVAL', { intervalHours: good });
  check(`intervalHours=${good} is accepted`, entry.intervalHours, good);
});

// ------------------------------------------------- malformed registry data

console.log('registry -- malformed stored data fails safely (never throws)');

{
  const { sandbox, scriptProps } = loadCode();
  scriptProps.setProperty('SA4_REGISTRY_INDEX', 'not valid json{{{');
  check('malformed index -> empty list, not a throw',
    sandbox.listRegisteredReportDocuments_(), []);
}

{
  const { sandbox, scriptProps } = loadCode();
  scriptProps.setProperty('SA4_REGISTRY_INDEX', JSON.stringify(['DOC_X']));
  scriptProps.setProperty('SA4_REGISTRY_DOC|DOC_X', 'not valid json{{{');
  check('malformed per-document entry -> null (treated as unregistered), not a throw',
    sandbox.getRegisteredReportDocument_('DOC_X'), null);
}

{
  const { sandbox, scriptProps } = loadCode();
  scriptProps.setProperty('SA4_REGISTRY_INDEX', JSON.stringify({ not: 'an array' }));
  check('non-array index value -> empty list, not a throw',
    sandbox.listRegisteredReportDocuments_(), []);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
