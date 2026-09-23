/**
 * PERF-003B (Part 2) — concurrency protection for overlapping incremental
 * updates.
 *
 * POST-MEETING-001 (Task C) flagged, but did not fix, the fact that
 * continuousUpdate() had no LockService guard: a time-driven trigger
 * firing again while a previous run (trigger- or menu-invoked) is still
 * executing, or a manual "Update Report (During Meeting)" click landing
 * mid-run, could mutate the same document body concurrently.
 *
 * The fix: continuousUpdate() is now a thin public wrapper that acquires
 * LockService.getDocumentLock() with a short, non-blocking tryLock(5000)
 * BEFORE any work, delegates to continuousUpdateCore_() (the actual
 * logic, previously named continuousUpdate()) only on success, and always
 * releases in `finally`. updateReportIncremental() gets the same document
 * lock at its own entry, independently (neither function calls the
 * other, so there is no nested/double-lock path).
 *
 * These tests exercise the WRAPPER logic in isolation, using a
 * controllable fake LockService and spies in place of the real
 * continuousUpdateCore_()/collectorUpdate_()/removeRowHeightAndSpacing()
 * (which themselves require a real DocumentApp body -- out of scope for
 * this pure-logic sandbox, same boundary as every other test in this
 * suite; see tests/README.md). What is proven here is exactly what the
 * task requires: the lock is acquired before any work, a losing run does
 * no work at all and returns safely, and the lock is always released.
 *
 * Run: node tests/perf003b-lock.test.js
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

function makeFakeLock(tryLockResult) {
  const calls = { tryLock: [], releaseLock: 0 };
  return {
    calls,
    lock: {
      tryLock: (ms) => { calls.tryLock.push(ms); return tryLockResult; },
    }
  };
}

// ==================== 1. continuousUpdate() -- lock acquired, wins ========

console.log('continuousUpdate() -- lock acquired successfully: does the real work exactly once, then releases');

{
  const { sandbox } = loadCode();
  const fake = makeFakeLock(true);
  let releaseCalls = 0;
  fake.lock.releaseLock = () => { releaseCalls++; };
  sandbox.LockService = { getDocumentLock: () => fake.lock };

  let coreCalls = 0;
  sandbox.continuousUpdateCore_ = () => { coreCalls++; };

  sandbox.continuousUpdate();

  check('tryLock was called exactly once', fake.calls.tryLock.length, 1);
  check('tryLock was called with a short (not "minutes") timeout', fake.calls.tryLock[0] <= 60000, true);
  check('continuousUpdateCore_ (the real logic) ran exactly once', coreCalls, 1);
  check('the lock was released exactly once', releaseCalls, 1);
}

// ============= 2. continuousUpdate() -- lock busy, loses the race =========

console.log('continuousUpdate() -- lock already held elsewhere: does NO work, returns safely, does not release a lock it never held');

{
  const { sandbox } = loadCode();
  const fake = makeFakeLock(false);
  let releaseCalls = 0;
  fake.lock.releaseLock = () => { releaseCalls++; };
  sandbox.LockService = { getDocumentLock: () => fake.lock };

  let coreCalls = 0;
  sandbox.continuousUpdateCore_ = () => { coreCalls++; };

  const loggedLines = [];
  sandbox.Logger = { log: (msg) => loggedLines.push(msg) };

  let threw = null;
  try { sandbox.continuousUpdate(); } catch (e) { threw = e; }

  check('continuousUpdate() does not throw when the lock is busy', threw, null);
  check('continuousUpdateCore_ (the real logic) never ran -- no partial work before/without the lock', coreCalls, 0);
  check('a clear log line was written explaining the skip', loggedLines.some(l => /already in progress/i.test(l)), true);
  check('releaseLock was never called (a lock never acquired is never released)', releaseCalls, 0);
}

// === 3. continuousUpdate() -- lock acquired but the real work throws ======

console.log('continuousUpdate() -- lock acquired, delegated work throws: lock is still released (finally)');

{
  const { sandbox } = loadCode();
  const fake = makeFakeLock(true);
  let releaseCalls = 0;
  fake.lock.releaseLock = () => { releaseCalls++; };
  sandbox.LockService = { getDocumentLock: () => fake.lock };

  sandbox.continuousUpdateCore_ = () => { throw new Error('boom'); };

  let threw = null;
  try { sandbox.continuousUpdate(); } catch (e) { threw = e.message; }

  check('the real error propagates (not swallowed by the lock wrapper)', threw, 'boom');
  check('the lock is still released even though the delegated work threw', releaseCalls, 1);
}

// ============ 4. updateReportIncremental() -- lock acquired, wins =========

console.log('updateReportIncremental() -- lock acquired successfully: does the real work exactly once, then releases');

{
  const { sandbox } = loadCode();
  const fake = makeFakeLock(true);
  let releaseCalls = 0;
  fake.lock.releaseLock = () => { releaseCalls++; };
  sandbox.LockService = { getDocumentLock: () => fake.lock };

  let collectorCalls = 0;
  let formatCalls = 0;
  sandbox.collectorUpdate_ = () => { collectorCalls++; };
  sandbox.removeRowHeightAndSpacing = () => { formatCalls++; };

  const alerts = [];
  sandbox.DocumentApp.getUi = () => ({
    alert: (...args) => { alerts.push(args); },
    ButtonSet: { OK: 'OK' }
  });

  sandbox.updateReportIncremental();

  check('tryLock was called exactly once', fake.calls.tryLock.length, 1);
  check('collectorUpdate_ ran exactly once', collectorCalls, 1);
  check('removeRowHeightAndSpacing ran exactly once', formatCalls, 1);
  check('a success alert was shown', alerts.length, 1);
  check('the lock was released exactly once', releaseCalls, 1);
}

// ======== 5. updateReportIncremental() -- lock busy, loses the race =======

console.log('updateReportIncremental() -- lock already held elsewhere: does NO work, alerts the user, does not release a lock it never held');

{
  const { sandbox } = loadCode();
  const fake = makeFakeLock(false);
  let releaseCalls = 0;
  fake.lock.releaseLock = () => { releaseCalls++; };
  sandbox.LockService = { getDocumentLock: () => fake.lock };

  let collectorCalls = 0;
  let formatCalls = 0;
  sandbox.collectorUpdate_ = () => { collectorCalls++; };
  sandbox.removeRowHeightAndSpacing = () => { formatCalls++; };

  const alerts = [];
  sandbox.DocumentApp.getUi = () => ({
    alert: (...args) => { alerts.push(args); },
    ButtonSet: { OK: 'OK' }
  });

  sandbox.updateReportIncremental();

  check('collectorUpdate_ never ran -- no partial work before/without the lock', collectorCalls, 0);
  check('removeRowHeightAndSpacing never ran', formatCalls, 0);
  check('the user was alerted that another update is in progress (not a silent no-op)', alerts.length, 1);
  check('releaseLock was never called', releaseCalls, 0);
}

// ==== 6. source-structure: getDocumentLock() is the lock type used ========

console.log('source-structure: both entry points use the document-scoped lock (not a script- or user-scoped one)');

{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  check('LockService.getDocumentLock() appears (document-scoped, matches "only one incremental update on THIS document")',
    /LockService\.getDocumentLock\(\)/.test(source), true);
  check('LockService.getScriptLock() is NOT used (would be shared across every document using this script)',
    /LockService\.getScriptLock\(\)/.test(source), false);
  check('LockService.getUserLock() is NOT used (would not block a different user/trigger identity on the SAME document)',
    /LockService\.getUserLock\(\)/.test(source), false);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
