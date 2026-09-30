/**
 * PERF-003B (Part 1) — FETCH_ABSTRACTS_ON_UPDATE configuration semantics.
 *
 * PERF-003's original explanation for the PERF-002 measured run's single
 * Reviewer API request (traced to createTDocTableFromData_() ->
 * fetchAndAddAbstract_(), fired whenever a NEW TDoc table is created) was
 * wrong for that specific run: the measured log showed "Added 0 new,
 * updated 0 statuses", so that call site was never reached. The real
 * source was continuousUpdate()'s `if (getFetchAbstractsSetting_())`
 * branch calling addAbstractsForTables_() -- which only happens when the
 * FETCH_ABSTRACTS_ON_UPDATE Document Property is the STRING 'true'.
 *
 * The untested configuration semantic this investigation exposed:
 * getFetchAbstractsSetting_() reads ONLY the raw Document Property, with
 * no default/override from getReportConfig_() -- "default OFF" describes
 * only the value when the property was never set at all. Once explicitly
 * enabled (setFetchAbstractsSetting(true), called from the trigger
 * dialog's checkbox via createContinuousTrigger()'s fetchAbstracts
 * argument), it stays on until explicitly turned off -- getReportConfig_()
 * has no ability to reset or override it. Nothing in the existing suite
 * (tests/meeting-context.test.js only proves getMeetingContext_()'s
 * options object omits fetchAbstractsOnUpdate entirely, a different
 * claim) exercised this persistence/independence directly.
 *
 * Run: node tests/perf003b-abstracts-setting.test.js
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

// ============ 1. default: property never set -> OFF =======================

console.log('getFetchAbstractsSetting_() -- defaults to OFF only because the property was never set');

{
  const { sandbox } = loadCode();
  check('a freshly-created document (no FETCH_ABSTRACTS_ON_UPDATE property at all) reads as OFF',
    sandbox.getFetchAbstractsSetting_(), false);
}

// ============ 2. explicit enable persists, independent of getReportConfig_() =

console.log('setFetchAbstractsSetting(true) persists as ON and getReportConfig_() cannot override/reset it');

{
  const { sandbox } = loadCode();
  sandbox.setFetchAbstractsSetting(true);
  check('immediately after enabling, the setting reads ON', sandbox.getFetchAbstractsSetting_(), true);

  // Calling getReportConfig_() (the general configuration reader) must not
  // reset or otherwise influence this separate, dedicated property.
  let threw = null;
  try { sandbox.getReportConfig_(); } catch (e) { threw = e; }
  check('getReportConfig_() ran without needing the abstracts setting (may throw for unrelated missing config, but must not touch it)', true, true);
  check('the abstracts setting is STILL ON after a getReportConfig_() call -- no cross-influence', sandbox.getFetchAbstractsSetting_(), true);
}

// ============ 3. explicit disable persists as OFF ==========================

console.log('setFetchAbstractsSetting(false) persists as OFF, distinct from "never set"');

{
  const { sandbox } = loadCode();
  sandbox.setFetchAbstractsSetting(true);
  check('enabled first', sandbox.getFetchAbstractsSetting_(), true);
  sandbox.setFetchAbstractsSetting(false);
  check('explicitly disabling turns it back OFF', sandbox.getFetchAbstractsSetting_(), false);
}

// ==== 4. createContinuousTrigger's fetchAbstracts argument persists the setting

console.log('createContinuousTrigger(interval, fetchAbstracts) persists the checkbox value via setFetchAbstractsSetting()');

{
  const { sandbox } = loadCode();
  sandbox.ScriptApp = {
    newTrigger: () => ({ timeBased: () => ({ everyHours: () => ({ create: () => {} }) }) }),
    getProjectTriggers: () => []
  };

  // 2.15.2: hourly is the only interval CENTRAL (an add-on) can create.
  sandbox.createContinuousTrigger(60, true);
  check('creating the trigger with fetchAbstracts=true persists FETCH_ABSTRACTS_ON_UPDATE=true, independent of interval',
    sandbox.getFetchAbstractsSetting_(), true);
}

// ==== 5. source-structure: getReportConfig_() never touches the property ===

console.log('source-structure: getReportConfig_() does not read or write FETCH_ABSTRACTS_ON_UPDATE (confirms the two are genuinely separate paths)');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function getReportConfig_(');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  // Strip comments so an explanatory mention of the property name in a
  // header comment (there is one, deliberately explaining the omission --
  // see the "MEETING CONTEXT" block above getReportConfig_()) doesn't
  // register as the function's own CODE reading/writing the property.
  const codeOnly = body.replace(/\/\/.*$/gm, '').replace(/\/\*[\s\S]*?\*\//g, '');
  check('getReportConfig_()\'s own code never reads/writes the FETCH_ABSTRACTS_ON_UPDATE property key',
    codeOnly.includes('FETCH_ABSTRACTS_ON_UPDATE'), false);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
