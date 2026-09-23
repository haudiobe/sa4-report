/**
 * ADDON-005 — add-on interactive UI: enableAutomaticUpdatesForAddon() /
 * disableAutomaticUpdatesForAddon() / setAutomaticUpdateIntervalForAddon()
 * (+ its RPC target) / showAddonSchedulerStatusForAddon(), and the
 * ScriptLock-based locking decision (withAddonScriptLock_()) that
 * resolves ADDON-004's open concurrency question.
 *
 * Run: node tests/addon005-interactive-addon-ui.test.js
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

function makeFakeScriptApp() {
  let triggers = [];
  let uidCounter = 0;
  function handle(t) { return { getHandlerFunction: () => t.handlerFunction, getUniqueId: () => t.uid }; }
  return {
    getProjectTriggers: () => triggers.map(handle),
    newTrigger: (h) => ({ timeBased: () => ({ everyHours: () => ({ create: () => { uidCounter++; const t = { handlerFunction: h, uid: 'u' + uidCounter }; triggers.push(t); return handle(t); } }) }) }),
    deleteTrigger: (h) => { triggers = triggers.filter(t => t.uid !== h.getUniqueId()); },
    _all: () => triggers.slice()
  };
}

function alertRecorder(sandbox) {
  const calls = [];
  sandbox.DocumentApp.getUi = () => ({
    alert: (...args) => { calls.push(args); },
    ButtonSet: { OK: 'OK' }
  });
  return calls;
}

function withActiveDocument(sandbox, documentId) {
  sandbox.DocumentApp.getActiveDocument = () => ({
    getId: () => documentId,
    getBody: () => ({})
  });
}

// -------------------------------------------------------- enable: fresh doc

console.log('enableAutomaticUpdatesForAddon() -- adopts an unregistered document, enables it, ensures the trigger');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_FRESH');
  const alerts = alertRecorder(sandbox);

  sandbox.enableAutomaticUpdatesForAddon();

  const entry = sandbox.getRegisteredReportDocument_('DOC_FRESH');
  check('document is now registered', entry !== null, true);
  check('document is enabled', entry.enabled, true);
  check('scheduler trigger was created', sandbox.getAddonSchedulerTriggerStatus_().exists, true);
  check('a success alert was shown', alerts.length, 1);
  check('Document Properties untouched by the adoption', docProps.getProperty('MEETING_ID'), '1');
}

console.log('enableAutomaticUpdatesForAddon() -- already-adopted document: re-adoption is idempotent (registeredAt preserved), just enables + ensures trigger');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_ALREADY');
  sandbox.adoptReportDocumentForAddon_({ documentId: 'DOC_ALREADY', mode: 'addon-interactive' });
  const registeredAtBefore = sandbox.getRegisteredReportDocument_('DOC_ALREADY').registeredAt;
  alertRecorder(sandbox);

  sandbox.enableAutomaticUpdatesForAddon();

  check('registeredAt is unchanged (adoptReportDocumentForAddon_() is idempotent, not a fresh registration)',
    sandbox.getRegisteredReportDocument_('DOC_ALREADY').registeredAt, registeredAtBefore);
  check('now enabled', sandbox.getRegisteredReportDocument_('DOC_ALREADY').enabled, true);
}

console.log('enableAutomaticUpdatesForAddon() -- RE-SYNCS central state from Document Properties every time it runs (not just on first adoption)');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X', TDOC_LIST_URL: 'https://example.invalid/first.xlsx' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_RESYNC');
  alertRecorder(sandbox);

  sandbox.enableAutomaticUpdatesForAddon();
  const centralBefore = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_RESYNC' });
  check('central state reflects the value present at first enable',
    centralBefore.getProperty('TDOC_LIST_URL'), 'https://example.invalid/first.xlsx');

  // Simulate an interactive config edit made AFTER adoption (e.g. via the
  // legacy "Configure Meeting Settings" dialog, which still writes
  // Document Properties -- see ADDON-004 report Section 5).
  docProps.setProperty('TDOC_LIST_URL', 'https://example.invalid/updated.xlsx');

  sandbox.enableAutomaticUpdatesForAddon(); // click "Enable" again

  const centralAfter = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_RESYNC' });
  check('central state now reflects the UPDATED Document Properties value',
    centralAfter.getProperty('TDOC_LIST_URL'), 'https://example.invalid/updated.xlsx');
}

console.log('enableAutomaticUpdatesForAddon() -- second document reuses the SAME single trigger (does not create a second)');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  const fakeScriptApp = makeFakeScriptApp();
  sandbox.ScriptApp = fakeScriptApp;
  alertRecorder(sandbox);

  withActiveDocument(sandbox, 'DOC_A');
  sandbox.enableAutomaticUpdatesForAddon();
  withActiveDocument(sandbox, 'DOC_B');
  sandbox.enableAutomaticUpdatesForAddon();

  check('both documents are registered and enabled',
    [sandbox.getRegisteredReportDocument_('DOC_A').enabled, sandbox.getRegisteredReportDocument_('DOC_B').enabled],
    [true, true]);
  check('exactly one scheduler trigger exists for both documents combined',
    fakeScriptApp._all().length, 1);
}

// ----------------------------------------------------------------- disable

console.log('disableAutomaticUpdatesForAddon() -- sets enabled=false, preserves state, never purges');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_DISABLE');
  alertRecorder(sandbox);
  sandbox.enableAutomaticUpdatesForAddon();

  const central = sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_DISABLE' });
  check('central state exists before disable', central.getProperty('MEETING_ID'), '1');

  sandbox.disableAutomaticUpdatesForAddon();

  check('enabled is now false', sandbox.getRegisteredReportDocument_('DOC_DISABLE').enabled, false);
  check('the registry entry itself still exists (not unregistered)',
    sandbox.getRegisteredReportDocument_('DOC_DISABLE') !== null, true);
  check('central report state was NOT purged',
    central.getProperty('MEETING_ID'), '1');
}

console.log('disableAutomaticUpdatesForAddon() -- never-registered document: no-op with a clear alert, no throw');

{
  const { sandbox } = loadCode();
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_NEVER');
  const alerts = alertRecorder(sandbox);

  sandbox.disableAutomaticUpdatesForAddon();

  check('an alert was shown, no throw', alerts.length, 1);
  check('still not registered', sandbox.getRegisteredReportDocument_('DOC_NEVER'), null);
}

// --------------------------------------------------------- allowed intervals

console.log('setAutomaticUpdateIntervalForAddonRpc() -- only platform-supported intervals are accepted');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_INTERVAL');
  alertRecorder(sandbox);
  sandbox.enableAutomaticUpdatesForAddon();

  [1, 2, 4, 6, 12, 24].forEach((h) => {
    const entry = sandbox.setAutomaticUpdateIntervalForAddonRpc(h);
    check(`interval=${h}h is accepted`, entry.intervalHours, h);
  });

  [15, 30, 3, 48].forEach((h) => {
    expectThrows(`interval=${h}h (unsupported by add-on triggers) is rejected`,
      () => sandbox.setAutomaticUpdateIntervalForAddonRpc(h), 'intervalHours must be one of');
  });
}

console.log('setAutomaticUpdateIntervalForAddon() dialog HTML never offers 15 or 30 minute-style options');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function setAutomaticUpdateIntervalForAddon(');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  check('dialog options come from REPORT_REGISTRY_ALLOWED_INTERVAL_HOURS_, not a hardcoded 15/30/60 list',
    /REPORT_REGISTRY_ALLOWED_INTERVAL_HOURS_\.map/.test(body), true);
  check('no literal 15-minute or 30-minute option string appears',
    /Every 15 minutes|Every 30 minutes/.test(body), false);
}

// -------------------------------------------------------------- status display

console.log('showAddonSchedulerStatusForAddon() -- reports correct interval/enabled state, no throw either way');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_STATUS');
  const alerts = alertRecorder(sandbox);

  sandbox.showAddonSchedulerStatusForAddon(); // unregistered case
  check('shows status even when never registered (no throw)', alerts.length, 1);

  sandbox.enableAutomaticUpdatesForAddon(); // also alerts once (index 1) -- status is the 3rd alert overall
  sandbox.setAutomaticUpdateIntervalForAddonRpc(6);
  sandbox.showAddonSchedulerStatusForAddon();

  const statusText = alerts[2][1];
  check('status mentions the correct interval', statusText.includes('every 6 hour'), true);
  check('status mentions enabled', statusText.includes('Enabled: Yes'), true);
}

// ----------------------------------------------------------------- locking

console.log('withAddonScriptLock_() -- used by all three interactive mutation entry points');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_LOCK');
  alertRecorder(sandbox);

  let tryLockCalls = 0;
  let releaseLockCalls = 0;
  sandbox.LockService.getScriptLock = () => ({
    tryLock: () => { tryLockCalls++; return true; },
    releaseLock: () => { releaseLockCalls++; }
  });

  sandbox.enableAutomaticUpdatesForAddon();
  check('enable acquires+releases the ScriptLock', [tryLockCalls, releaseLockCalls], [1, 1]);

  sandbox.disableAutomaticUpdatesForAddon();
  check('disable acquires+releases the ScriptLock', [tryLockCalls, releaseLockCalls], [2, 2]);

  sandbox.registerReportDocument_('DOC_LOCK', { enabled: true }); // re-enable for the interval RPC below
  sandbox.setAutomaticUpdateIntervalForAddonRpc(2);
  check('setInterval acquires+releases the ScriptLock', [tryLockCalls, releaseLockCalls], [3, 3]);
}

console.log('withAddonScriptLock_() -- lock contention surfaces as a clear error, never silent data loss');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_ID: '1', MEETING_FOLDER: 'X' } });
  sandbox.ScriptApp = makeFakeScriptApp();
  withActiveDocument(sandbox, 'DOC_CONTENDED');
  alertRecorder(sandbox);

  sandbox.LockService.getScriptLock = () => ({ tryLock: () => false, releaseLock: () => { throw new Error('should never be called'); } });

  expectThrows('enableAutomaticUpdatesForAddon() throws clearly when the scheduler (or another interactive call) holds the lock',
    () => sandbox.enableAutomaticUpdatesForAddon(), 'already in progress');
  check('the document was NOT registered as a side effect of the failed attempt',
    sandbox.getRegisteredReportDocument_('DOC_CONTENDED'), null);
}

console.log('withAddonScriptLock_() -- shares the SAME lock type the scheduler uses (getScriptLock, not getDocumentLock)');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function withAddonScriptLock_(');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  check('withAddonScriptLock_() uses LockService.getScriptLock()',
    /LockService\.getScriptLock\(\)/.test(body), true);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
