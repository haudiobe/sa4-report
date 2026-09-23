/**
 * ADDON-005 — central scheduler trigger lifecycle:
 * ensureAddonSchedulerTrigger_() / getAddonSchedulerTriggerStatus_() /
 * deleteAddonSchedulerTrigger_() / findAddonSchedulerTriggers_(), and the
 * public runAddonSchedulerTrigger() wrapper used as the actual handler
 * name (see Code.js header comment for why it's public, not
 * runAddonScheduler_() directly).
 *
 * Uses a small stateful fake ScriptApp (tracks created/deleted triggers
 * by handler name) so duplicate-prevention, status, and safe-deletion
 * behavior can be verified without a real Apps Script project.
 *
 * Run: node tests/addon005-scheduler-trigger.test.js
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

function makeFakeScriptApp() {
  let triggers = [];
  let uidCounter = 0;

  function triggerHandle(t) {
    return { getHandlerFunction: () => t.handlerFunction, getUniqueId: () => t.uid };
  }

  return {
    getProjectTriggers: () => triggers.map(triggerHandle),
    newTrigger: (handlerFunction) => ({
      timeBased: () => ({
        everyHours: (n) => ({
          create: () => {
            uidCounter++;
            const t = { handlerFunction: handlerFunction, uid: 'uid-' + uidCounter, everyHours: n };
            triggers.push(t);
            return triggerHandle(t);
          }
        })
      })
    }),
    deleteTrigger: (handle) => {
      const uid = handle.getUniqueId();
      triggers = triggers.filter((t) => t.uid !== uid);
    },
    // test-only helpers, not part of the real ScriptApp surface
    _seedForeignTrigger: (handlerFunction) => {
      uidCounter++;
      triggers.push({ handlerFunction: handlerFunction, uid: 'uid-' + uidCounter });
    },
    _all: () => triggers.slice()
  };
}

// -------------------------------------------------------------- handler name

console.log('runAddonSchedulerTrigger() -- public wrapper delegates to runAddonScheduler_()');

{
  const { sandbox } = loadCode();
  let coreCalls = 0;
  sandbox.runAddonScheduler_ = () => { coreCalls++; return { considered: 0 }; };
  const result = sandbox.runAddonSchedulerTrigger();
  check('delegates to runAddonScheduler_()', coreCalls, 1);
  check('returns its result', result, { considered: 0 });
}

console.log('the registered handler function name matches REPORT_SCHEDULER_TRIGGER_HANDLER_ ("runAddonSchedulerTrigger")');

{
  const { sandbox } = loadCode();
  check('REPORT_SCHEDULER_TRIGGER_HANDLER_ constant', sandbox.REPORT_SCHEDULER_TRIGGER_HANDLER_, 'runAddonSchedulerTrigger');
  check('the handler function actually exists under that exact name',
    typeof sandbox[sandbox.REPORT_SCHEDULER_TRIGGER_HANDLER_], 'function');
}

// -------------------------------------------------------------- ensure/create

console.log('ensureAddonSchedulerTrigger_() -- creates exactly one hourly trigger when none exists');

{
  const { sandbox } = loadCode();
  const fakeScriptApp = makeFakeScriptApp();
  sandbox.ScriptApp = fakeScriptApp;

  const result = sandbox.ensureAddonSchedulerTrigger_();

  check('reports created:true', result.created, true);
  check('exactly one trigger exists', fakeScriptApp._all().length, 1);
  check('the trigger uses the handler name, not runAddonScheduler_ directly',
    fakeScriptApp._all()[0].handlerFunction, 'runAddonSchedulerTrigger');
  check('the trigger is hourly (everyHours(1))', fakeScriptApp._all()[0].everyHours, 1);
  check('the trigger UID is stored in Script Properties',
    sandbox.PropertiesService.getScriptProperties().getProperty('SA4_SCHEDULER_TRIGGER_UID'), result.triggerUid);
}

console.log('ensureAddonSchedulerTrigger_() -- duplicate prevention: calling it again does not create a second trigger');

{
  const { sandbox } = loadCode();
  const fakeScriptApp = makeFakeScriptApp();
  sandbox.ScriptApp = fakeScriptApp;

  const first = sandbox.ensureAddonSchedulerTrigger_();
  const second = sandbox.ensureAddonSchedulerTrigger_();

  check('first call creates', first.created, true);
  check('second call does NOT create', second.created, false);
  check('still exactly one trigger', fakeScriptApp._all().length, 1);
  check('same trigger UID both times', second.triggerUid, first.triggerUid);
}

console.log('ensureAddonSchedulerTrigger_() -- self-heals by removing extra duplicates if somehow more than one exists');

{
  const { sandbox } = loadCode();
  const fakeScriptApp = makeFakeScriptApp();
  sandbox.ScriptApp = fakeScriptApp;
  // Simulate a pre-existing duplicate (e.g. from a platform hiccup) by
  // seeding two matching triggers directly, bypassing ensureAddonSchedulerTrigger_().
  fakeScriptApp._seedForeignTrigger('runAddonSchedulerTrigger');
  fakeScriptApp._seedForeignTrigger('runAddonSchedulerTrigger');

  const result = sandbox.ensureAddonSchedulerTrigger_();

  check('did not create a THIRD trigger', result.created, false);
  check('one duplicate was removed', result.duplicatesRemoved, 1);
  check('exactly one matching trigger remains', fakeScriptApp._all().length, 1);
}

// -------------------------------------------------------------- status

console.log('getAddonSchedulerTriggerStatus_() -- read-only status inspection');

{
  const { sandbox } = loadCode();
  const fakeScriptApp = makeFakeScriptApp();
  sandbox.ScriptApp = fakeScriptApp;

  const before = sandbox.getAddonSchedulerTriggerStatus_();
  check('no trigger yet', before.exists, false);
  check('count is 0', before.count, 0);

  sandbox.ensureAddonSchedulerTrigger_();
  const after = sandbox.getAddonSchedulerTriggerStatus_();
  check('trigger now exists', after.exists, true);
  check('count is 1', after.count, 1);
  check('does not itself create or delete anything (still exactly one after two status calls)',
    (sandbox.getAddonSchedulerTriggerStatus_(), fakeScriptApp._all().length), 1);
}

// -------------------------------------------------------------- deletion

console.log('deleteAddonSchedulerTrigger_() -- safe deletion, registry/state untouched');

{
  const { sandbox } = loadCode();
  const fakeScriptApp = makeFakeScriptApp();
  sandbox.ScriptApp = fakeScriptApp;
  sandbox.ensureAddonSchedulerTrigger_();

  sandbox.registerReportDocument_('DOC_SURVIVES', { enabled: true });
  sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_SURVIVES' }).setProperty('K', 'v');

  const result = sandbox.deleteAddonSchedulerTrigger_();

  check('reports one deleted', result.deleted, 1);
  check('no scheduler trigger remains', fakeScriptApp._all().length, 0);
  check('the stored trigger UID property is cleared',
    sandbox.PropertiesService.getScriptProperties().getProperty('SA4_SCHEDULER_TRIGGER_UID'), null);
  check('the registry entry SURVIVES trigger deletion',
    sandbox.getRegisteredReportDocument_('DOC_SURVIVES') !== null, true);
  check('the central report state SURVIVES trigger deletion',
    sandbox.getReportStateStore_({ mode: 'addon-background', documentId: 'DOC_SURVIVES' }).getProperty('K'), 'v');
}

// -------------------------------------------------------------- unrelated triggers

console.log('creation/deletion never touches triggers with a DIFFERENT handler function');

{
  const { sandbox } = loadCode();
  const fakeScriptApp = makeFakeScriptApp();
  sandbox.ScriptApp = fakeScriptApp;
  fakeScriptApp._seedForeignTrigger('someOtherProjectTrigger');
  fakeScriptApp._seedForeignTrigger('yetAnotherHandler');

  sandbox.ensureAddonSchedulerTrigger_();
  check('creating the scheduler trigger left the 2 unrelated triggers alone',
    fakeScriptApp._all().filter(t => t.handlerFunction !== 'runAddonSchedulerTrigger').length, 2);

  sandbox.deleteAddonSchedulerTrigger_();
  check('deleting the scheduler trigger left the 2 unrelated triggers alone',
    fakeScriptApp._all().map(t => t.handlerFunction).sort(),
    ['someOtherProjectTrigger', 'yetAnotherHandler']);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
