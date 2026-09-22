/**
 * PERF-001 — Lightweight performance instrumentation for continuousUpdate().
 *
 * These tests exercise the timing/counting HELPERS themselves
 * (perfTimed_, perfTimedAccum_, perfAddTime_, perfCount_,
 * getTablesCounted_, getActiveDocumentBodyCounted_, perfResetState_,
 * perfLogSummary_) in isolation, proving:
 *   - they return the wrapped function's real result/error unchanged
 *     (functional output is never altered by instrumentation);
 *   - accumulation and counting arithmetic is correct;
 *   - state resets cleanly between runs (so one continuousUpdate()
 *     invocation's numbers can never leak into the next).
 *
 * This suite does NOT attempt to assert real timing values from
 * continuousUpdate() itself (that requires a live Apps Script execution,
 * explicitly deferred to a later controlled run per this task's own
 * instructions) -- it proves the instrumentation is correct and safe to
 * have shipped, using a fake DocumentApp body (same minimal pattern as
 * tests/revision-drafts-folder.test.js) to prove getTablesCounted_()/
 * getActiveDocumentBodyCounted_() return the exact same values the
 * uninstrumented calls would have.
 *
 * Run: node tests/perf-instrumentation.test.js
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

// ============================== 1. perfTimed_() ============================

console.log('perfTimed_() -- returns fn()\'s real result unchanged, times it, logs and accumulates');

{
  const { sandbox } = loadCode();
  sandbox.perfResetState_();

  const result = sandbox.perfTimed_('test-stage-a', () => 42);
  check('perfTimed_ returns the wrapped function\'s exact return value', result, 42);
  check('perfTimed_ recorded a non-negative accumulated time for the label',
    sandbox.PERF_TIMINGS_['test-stage-a'] >= 0, true);
}

{
  const { sandbox } = loadCode();
  sandbox.perfResetState_();
  let threw = null;
  try {
    sandbox.perfTimed_('test-stage-throws', () => { throw new Error('boom'); });
  } catch (e) {
    threw = e.message;
  }
  check('perfTimed_ re-throws the wrapped function\'s real error, never swallows it', threw, 'boom');
  check('perfTimed_ still recorded timing even though fn() threw (finally-based)',
    sandbox.PERF_TIMINGS_['test-stage-throws'] >= 0, true);
}

// ============================ 2. perfTimedAccum_() ==========================

console.log('perfTimedAccum_() -- accumulates across multiple calls under the SAME label, no per-call log line');

{
  const { sandbox } = loadCode();
  sandbox.perfResetState_();

  const r1 = sandbox.perfTimedAccum_('per-item-stage', () => 'a');
  const r2 = sandbox.perfTimedAccum_('per-item-stage', () => 'b');
  const r3 = sandbox.perfTimedAccum_('per-item-stage', () => 'c');

  check('perfTimedAccum_ returns each call\'s real result unchanged', [r1, r2, r3], ['a', 'b', 'c']);
  check('perfTimedAccum_ accumulates into a SINGLE total for the label (not three separate entries)',
    Object.keys(sandbox.PERF_TIMINGS_).filter(k => k === 'per-item-stage').length, 1);
  check('the accumulated total is non-negative', sandbox.PERF_TIMINGS_['per-item-stage'] >= 0, true);
}

// ============================== 3. perfCount_() =============================

console.log('perfCount_() -- exact counting arithmetic, default increment of 1, explicit increments');

{
  const { sandbox } = loadCode();
  sandbox.perfResetState_();

  sandbox.perfCount_('widgets');
  sandbox.perfCount_('widgets');
  sandbox.perfCount_('widgets', 3);
  check('perfCount_ defaults to +1 per call and accepts an explicit increment', sandbox.PERF_COUNTERS_['widgets'], 5);

  sandbox.perfCount_('gadgets', 10);
  check('a fresh label starts from the explicit increment, not from a stale prior value', sandbox.PERF_COUNTERS_['gadgets'], 10);
}

// ============================ 4. perfResetState_() ===========================

console.log('perfResetState_() -- fully clears both timings and counters between runs (no cross-run leakage)');

{
  const { sandbox } = loadCode();
  sandbox.perfCount_('leftover-counter', 7);
  sandbox.perfAddTime_('leftover-timing', 123);
  check('state exists before reset', [sandbox.PERF_COUNTERS_['leftover-counter'], sandbox.PERF_TIMINGS_['leftover-timing']], [7, 123]);

  sandbox.perfResetState_();
  check('perfResetState_ clears all counters', Object.keys(sandbox.PERF_COUNTERS_).length, 0);
  check('perfResetState_ clears all timings', Object.keys(sandbox.PERF_TIMINGS_).length, 0);
}

// ==================== 5. getTablesCounted_() / getActiveDocumentBodyCounted_()

console.log('getTablesCounted_() / getActiveDocumentBodyCounted_() -- return the exact same values as the uninstrumented calls, while counting');

{
  const { sandbox } = loadCode();
  sandbox.perfResetState_();

  // Minimal fake DocumentApp Body, same shape as
  // tests/revision-drafts-folder.test.js's fake -- just enough surface
  // for getTables()/getBody() to be meaningfully counted.
  const fakeTables = [{ marker: 'table-1' }, { marker: 'table-2' }];
  const fakeBody = { getTables: () => fakeTables };
  const fakeDoc = { getBody: () => fakeBody };
  sandbox.DocumentApp = { getActiveDocument: () => fakeDoc };

  const bodyResult = sandbox.getActiveDocumentBodyCounted_('unit-test-site');
  check('getActiveDocumentBodyCounted_ returns the exact same Body getBody() would have', bodyResult, fakeBody);
  check('getActiveDocumentBodyCounted_ incremented the total counter', sandbox.PERF_COUNTERS_['DocumentApp.getActiveDocument().getBody() calls (total)'], 1);
  check('getActiveDocumentBodyCounted_ incremented the per-site counter', sandbox.PERF_COUNTERS_['getActiveDocument().getBody() call site: unit-test-site'], 1);

  const tablesResult = sandbox.getTablesCounted_(fakeBody, 'unit-test-site');
  check('getTablesCounted_ returns the exact same array body.getTables() would have', tablesResult, fakeTables);
  check('getTablesCounted_ incremented the total counter', sandbox.PERF_COUNTERS_['body.getTables() calls (total)'], 1);
  check('getTablesCounted_ incremented the per-site counter', sandbox.PERF_COUNTERS_['body.getTables() call site: unit-test-site'], 1);

  // Calling it again from a DIFFERENT site must add a new counter, not
  // overwrite or merge with the first site's count.
  sandbox.getTablesCounted_(fakeBody, 'another-site');
  check('a second call site gets its OWN counter', sandbox.PERF_COUNTERS_['body.getTables() call site: another-site'], 1);
  check('the total counter reflects BOTH calls', sandbox.PERF_COUNTERS_['body.getTables() calls (total)'], 2);
  check('the first site\'s own counter is untouched by the second call', sandbox.PERF_COUNTERS_['body.getTables() call site: unit-test-site'], 1);
}

// ============================ 6. perfLogSummary_() ===========================

console.log('perfLogSummary_() -- purely a Logger.log summary, never throws, never mutates state');

{
  const { sandbox } = loadCode();
  sandbox.perfResetState_();
  sandbox.perfCount_('x', 3);
  sandbox.perfAddTime_('y', 50);

  const loggedLines = [];
  sandbox.Logger = { log: (msg) => loggedLines.push(msg) };

  let threw = null;
  try { sandbox.perfLogSummary_(); } catch (e) { threw = e; }

  check('perfLogSummary_ does not throw', threw, null);
  check('perfLogSummary_ logs exactly 2 lines (one timings summary, one counters summary)', loggedLines.length, 2);
  check('the timings line is machine-searchable with the "[PERF]" prefix', loggedLines[0].indexOf('[PERF]') === 0, true);
  check('the counters line is machine-searchable with the "[PERF]" prefix', loggedLines[1].indexOf('[PERF]') === 0, true);
  check('perfLogSummary_ does not mutate PERF_COUNTERS_/PERF_TIMINGS_', [sandbox.PERF_COUNTERS_['x'], sandbox.PERF_TIMINGS_['y']], [3, 50]);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
