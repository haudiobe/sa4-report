/**
 * 2.15.1 / 2.15.2 -- Continuous Update trigger interval.
 *
 * Live failures:
 *   2.15.0: "Every hour (Slow meeting)" -> everyMinutes(60) -> "The value
 *           you passed to everyMinutes was invalid".
 *   2.15.1: "Every 30 minutes (Recommended)" -> everyMinutes(30) failed:
 *           CENTRAL runs as an Editor add-on, and add-on time-driven
 *           triggers cannot run more often than once per hour (ADDON-001B,
 *           the same rule the add-on scheduler already follows).
 *
 * The fake TriggerBuilder below enforces the ADD-ON rules: no sub-hourly
 * clock trigger at all, and everyHours() only with Apps Script's accepted
 * values (1, 2, 4, 6, 8, 12). Every ScriptApp call is recorded.
 *
 * Run: node tests/continuous-trigger-interval.test.js
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

const ADDON_SUBHOURLY_ERROR = 'Add-ons may not create time-based triggers to run more frequently than once per hour.';

/** Loads Code.js with a recording, add-on-rule ScriptApp; `existing` = handler names of pre-existing triggers. */
function setup(existing, props) {
  const loaded = loadCode({ documentProperties: props });
  const s = loaded.sandbox;
  const calls = [];
  let triggers = (existing || []).map((h) => ({ getHandlerFunction: () => h, getTriggerSource: () => 'CLOCK', _h: h }));
  const add = (handler) => { calls.push(['create', handler]); triggers.push({ getHandlerFunction: () => handler, getTriggerSource: () => 'CLOCK', _h: handler }); };
  const clock = (handler) => ({
    everyMinutes: (n) => {
      calls.push(['everyMinutes', n]);
      if ([1, 5, 10, 15, 30].indexOf(n) === -1) throw new Error('The value you passed to everyMinutes was invalid. It must be one of 1, 5, 10, 15 or 30.');
      return { create: () => { throw new Error(ADDON_SUBHOURLY_ERROR); } };
    },
    everyHours: (n) => {
      calls.push(['everyHours', n]);
      if ([1, 2, 4, 6, 8, 12].indexOf(n) === -1) throw new Error('The value you passed to everyHours was invalid.');
      return { create: () => add(handler) };
    }
  });
  s.ScriptApp = {
    TriggerSource: { CLOCK: 'CLOCK' },
    newTrigger: (handler) => { calls.push(['newTrigger', handler]); return { timeBased: () => clock(handler) }; },
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: (t) => { calls.push(['deleteTrigger', t._h]); triggers = triggers.filter((x) => x !== t); }
  };
  return { s, calls, docProps: loaded.docProps, handlers: () => triggers.map((t) => t._h) };
}

// ----------------------------------------------------- the dialog's options

console.log('the dialog offers only intervals an add-on can create');
let offered;
{
  const { s } = setup([]);
  let html = null;
  s.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  s.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  s.manageTriggers();
  offered = [];
  html.replace(/<option value="([^"]*)"( selected)?>([^<]*)<\/option>/g, (m, v, sel, label) => { offered.push([v, !!sel, label]); return m; });
  check('options (value, selected, label): hourly only, label unchanged', offered, [['60', true, 'Every hour (Slow meeting)']]);
  check('no 15/30-minute option', /Every 15 minutes|Every 30 minutes|value="15"|value="30"/.test(html), false);
  check('Start Trigger still sends the selected value to createContinuousTrigger',
    /\.createContinuousTrigger\(parseInt\(interval\), document\.getElementById\('fetchAbstracts'\)\.checked\)/.test(html), true);
}

// ------------------------------------------- every option -> valid builder

console.log('every offered option creates a trigger under the add-on rules');
offered.forEach(([value, , label]) => {
  // exactly what the dialog passes: parseInt(<option value>)
  const env = setup([]);
  env.s.createContinuousTrigger(parseInt(value), false);
  check(`${label} -> everyHours(1)`, env.calls, [['newTrigger', 'continuousUpdate'], ['everyHours', 1], ['create', 'continuousUpdate']]);
  check(`${label}: one continuousUpdate trigger exists`, env.handlers(), ['continuousUpdate']);
});

{
  const env = setup([]);
  env.s.createContinuousTrigger(60, false);
  check('every hour never calls everyMinutes', env.calls.some((c) => c[0] === 'everyMinutes'), false);
  const str = setup([]);
  str.s.createContinuousTrigger('60', false);
  check('a digit string is accepted the same way', str.calls[1], ['everyHours', 1]);
}

// --------------------------------------- invalid input fails before ScriptApp

console.log('unsupported / malformed / missing intervals are refused before any trigger change');
[
  ['unsupported 30 (the 2.15.1 "Recommended" option)', 30], ['unsupported 15 (the 2.15.1 "Active meeting" option)', 15],
  ['unsupported 45', 45], ['unsupported 5', 5], ['unsupported 120', 120],
  ['malformed "60 minutes"', '60 minutes'], ['malformed "abc"', 'abc'], ['malformed 60.5', 60.5], ['malformed NaN (parseInt of "")', NaN],
  ['malformed true', true], ['missing undefined', undefined], ['missing null', null], ['missing ""', '']
].forEach(([name, value]) => {
  const env = setup(['continuousUpdate', 'runAddonSchedulerTrigger'], { FETCH_ABSTRACTS_ON_UPDATE: 'false' });
  let error = null;
  try { env.s.createContinuousTrigger(value, true); } catch (e) { error = e.message; }
  check(`${name}: clear application error`, /^Unsupported update interval: .* \(supported: 60 minutes\)\.$/.test(error || ''), true);
  check(`${name}: ScriptApp never called, running triggers and abstracts switch untouched`,
    [env.calls, env.handlers(), env.docProps._store.FETCH_ABSTRACTS_ON_UPDATE], [[], ['continuousUpdate', 'runAddonSchedulerTrigger'], 'false']);
});

// --------------------------------------------- existing behavior preserved

console.log('existing trigger management unchanged');
{
  // Restart: the old continuousUpdate trigger is replaced; others survive.
  const env = setup(['continuousUpdate', 'runAddonSchedulerTrigger']);
  env.s.createContinuousTrigger(60, true);
  check('the previous continuousUpdate trigger is deleted first, the add-on scheduler trigger is kept',
    [env.calls[0], env.handlers()], [['deleteTrigger', 'continuousUpdate'], ['runAddonSchedulerTrigger', 'continuousUpdate']]);
  check('the abstracts checkbox is still persisted', env.docProps._store.FETCH_ABSTRACTS_ON_UPDATE, 'true');
  const status = env.s.getTriggerStatus();
  check('getTriggerStatus() reports it active', status, { active: true, interval: 'Active' });
  env.s.deleteContinuousTrigger();
  check('Stop Trigger removes only continuousUpdate', env.handlers(), ['runAddonSchedulerTrigger']);
  check('getTriggerStatus() reports inactive afterwards', env.s.getTriggerStatus(), { active: false, interval: null });
}

{
  const env = setup([]);
  env.s.createContinuousTrigger(60);
  check('no fetchAbstracts argument leaves the switch unset (unchanged)', env.docProps._store.FETCH_ABSTRACTS_ON_UPDATE, undefined);
}

{
  // The central add-on scheduler: same rules, untouched by the fix.
  const env = setup([]);
  const fakeTrigger = (h) => ({ getHandlerFunction: () => h, getUniqueId: () => 'uid-1', _h: h });
  env.s.ScriptApp.newTrigger = (h) => ({ timeBased: () => ({ everyHours: (n) => { env.calls.push(['sched.everyHours', n]); return { create: () => fakeTrigger(h) }; } }) });
  const r = env.s.ensureAddonSchedulerTrigger_();
  check('the add-on scheduler still creates runAddonSchedulerTrigger with everyHours(1)', [r.created, env.calls], [true, [['sched.everyHours', 1]]]);
}

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const code = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  check('no everyMinutes() trigger call remains in Code.js (comments ignored)', /\.everyMinutes\(/.test(code), false);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
