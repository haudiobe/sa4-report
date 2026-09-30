/**
 * 2.15.1 -- Continuous Update trigger interval.
 *
 * Live failure: "Manage Auto-Update Trigger" -> "Every hour (Slow meeting)"
 * -> Start Trigger threw "The value you passed to everyMinutes was
 * invalid. It must be one of 1, 5, 10, 15 or 30." -- the dialog's value
 * "60" reached everyMinutes(60).
 *
 * The fake TriggerBuilder below enforces the real Apps Script argument
 * rules, and records every ScriptApp call.
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

/** Loads Code.js with a recording ScriptApp; `existing` = handler names of pre-existing triggers. */
function setup(existing, props) {
  const loaded = loadCode({ documentProperties: props });
  const s = loaded.sandbox;
  const calls = [];
  let triggers = (existing || []).map((h) => ({ getHandlerFunction: () => h, getTriggerSource: () => 'CLOCK', _h: h }));
  const clock = (handler) => ({
    everyMinutes: (n) => {
      calls.push(['everyMinutes', n]);
      if ([1, 5, 10, 15, 30].indexOf(n) === -1) throw new Error('The value you passed to everyMinutes was invalid. It must be one of 1, 5, 10, 15 or 30.');
      return { create: () => { calls.push(['create', handler]); triggers.push({ getHandlerFunction: () => handler, getTriggerSource: () => 'CLOCK', _h: handler }); } };
    },
    everyHours: (n) => {
      calls.push(['everyHours', n]);
      if ([1, 2, 4, 6, 8, 12].indexOf(n) === -1) throw new Error('The value you passed to everyHours was invalid.');
      return { create: () => { calls.push(['create', handler]); triggers.push({ getHandlerFunction: () => handler, getTriggerSource: () => 'CLOCK', _h: handler }); } };
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

console.log('the dialog offers exactly the mapped intervals, labels unchanged');
let offered;
{
  const { s } = setup([]);
  let html = null;
  s.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  s.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  s.manageTriggers();
  offered = [];
  html.replace(/<option value="([^"]*)"( selected)?>([^<]*)<\/option>/g, (m, v, sel, label) => { offered.push([v, !!sel, label]); return m; });
  check('options (value, selected, label)', offered, [
    ['15', false, 'Every 15 minutes (Active meeting)'],
    ['30', true, 'Every 30 minutes (Recommended)'],
    ['60', false, 'Every hour (Slow meeting)']
  ]);
  check('Start Trigger still sends the selected value to createContinuousTrigger',
    /\.createContinuousTrigger\(parseInt\(interval\), document\.getElementById\('fetchAbstracts'\)\.checked\)/.test(html), true);
}

// ------------------------------------------- every option -> valid builder

console.log('every offered option creates a valid Apps Script trigger');
const EXPECTED = { '15': ['everyMinutes', 15], '30': ['everyMinutes', 30], '60': ['everyHours', 1] };
offered.forEach(([value, , label]) => {
  // exactly what the dialog passes: parseInt(<option value>)
  const env = setup([]);
  env.s.createContinuousTrigger(parseInt(value), false);
  check(`${label} -> ${EXPECTED[value][0]}(${EXPECTED[value][1]})`,
    env.calls, [['newTrigger', 'continuousUpdate'], EXPECTED[value], ['create', 'continuousUpdate']]);
  check(`${label}: one continuousUpdate trigger exists`, env.handlers(), ['continuousUpdate']);
});

{
  const env = setup([]);
  env.s.createContinuousTrigger(60, false);
  check('every hour never calls everyMinutes (no everyMinutes(60))', env.calls.some((c) => c[0] === 'everyMinutes'), false);
  const str = setup([]);
  str.s.createContinuousTrigger('60', false);
  check('a digit string is accepted the same way', str.calls[1], ['everyHours', 1]);
}

// --------------------------------------- invalid input fails before ScriptApp

console.log('unsupported / malformed / missing intervals are refused before any trigger change');
[
  ['unsupported 45', 45], ['unsupported 5 (valid for Apps Script, not offered)', 5], ['unsupported 120', 120],
  ['malformed "60 minutes"', '60 minutes'], ['malformed "abc"', 'abc'], ['malformed 60.5', 60.5], ['malformed NaN (parseInt of "")', NaN],
  ['malformed true', true], ['missing undefined', undefined], ['missing null', null], ['missing ""', '']
].forEach(([name, value]) => {
  const env = setup(['continuousUpdate', 'runAddonSchedulerTrigger'], { FETCH_ABSTRACTS_ON_UPDATE: 'false' });
  let error = null;
  try { env.s.createContinuousTrigger(value, true); } catch (e) { error = e.message; }
  check(`${name}: clear application error`, /^Unsupported update interval: .* \(supported: 15, 30, 60 minutes\)\.$/.test(error || ''), true);
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
  env.s.createContinuousTrigger(30);
  check('no fetchAbstracts argument leaves the switch unset (unchanged)', env.docProps._store.FETCH_ABSTRACTS_ON_UPDATE, undefined);
}

{
  // The central add-on's own hourly scheduler trigger is a separate path.
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const code = src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  check('no everyMinutes call with a literal outside 1/5/10/15/30 anywhere in Code.js (comments ignored)',
    (code.match(/everyMinutes\(\s*(\d+)\s*\)/g) || []).filter((m) => ['1', '5', '10', '15', '30'].indexOf(m.match(/\d+/)[0]) === -1), []);
  check('the add-on scheduler still uses everyHours(1)', /ensureAddonSchedulerTrigger_[\s\S]*?\.everyHours\(1\)/.test(src), true);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
