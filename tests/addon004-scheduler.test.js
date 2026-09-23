/**
 * ADDON-004 — the central scheduler HANDLER: runAddonScheduler_(),
 * isReportDocumentDue_(), orderDueReportDocuments_(). No trigger is
 * created anywhere -- this exercises the handler function directly, as a
 * plain function call, exactly as ADDON-004 scoped it.
 *
 * Run: node tests/addon004-scheduler.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
const vm = require('vm');

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

function setNow(sandbox, ms) {
  vm.runInContext(`Date.now = function() { return ${ms}; };`, sandbox);
}

function makeFakeDocument(id) {
  const body = { getTables: () => [], clear: () => body };
  return { getId: () => id, getBody: () => body };
}

function registerEnabled(sandbox, documentId, extra) {
  sandbox.registerReportDocument_(documentId, Object.assign({ enabled: true }, extra || {}));
}

// ==================================================== isReportDocumentDue_

console.log('isReportDocumentDue_() -- pure due calculation');

{
  const { sandbox } = loadCode();
  const now = 1000000;

  check('disabled document is never due',
    sandbox.isReportDocumentDue_({ enabled: false, lastRunAt: null, intervalHours: 1 }, now), false);
  check('never-run (lastRunAt null) enabled document is always due',
    sandbox.isReportDocumentDue_({ enabled: true, lastRunAt: null, intervalHours: 24 }, now), true);
  check('null/undefined entry is never due',
    sandbox.isReportDocumentDue_(null, now), false);

  [1, 2, 4, 6, 12, 24].forEach((hours) => {
    const intervalMs = hours * 3600 * 1000;
    const justBefore = { enabled: true, intervalHours: hours, lastRunAt: new Date(now - intervalMs + 1000).toISOString() };
    const exactlyAt = { enabled: true, intervalHours: hours, lastRunAt: new Date(now - intervalMs).toISOString() };
    const wellAfter = { enabled: true, intervalHours: hours, lastRunAt: new Date(now - intervalMs - 1000).toISOString() };

    check(`intervalHours=${hours}: not yet due (${hours}h minus 1s elapsed)`,
      sandbox.isReportDocumentDue_(justBefore, now), false);
    check(`intervalHours=${hours}: due exactly at the interval boundary`,
      sandbox.isReportDocumentDue_(exactlyAt, now), true);
    check(`intervalHours=${hours}: due well after the interval`,
      sandbox.isReportDocumentDue_(wellAfter, now), true);
  });

  check('a malformed lastRunAt fails safe as due (never permanently stuck)',
    sandbox.isReportDocumentDue_({ enabled: true, intervalHours: 1, lastRunAt: 'not-a-date' }, now), true);
}

// =============================================== orderDueReportDocuments_

console.log('orderDueReportDocuments_() -- fair ordering: never-run first, then oldest lastRunAt, stable tie-break');

{
  const { sandbox } = loadCode();

  const entries = [
    { documentId: 'C', lastRunAt: '2026-01-01T00:00:00.000Z' },
    { documentId: 'A', lastRunAt: null },
    { documentId: 'B', lastRunAt: '2026-01-01T00:00:00.000Z' }, // same timestamp as C
    { documentId: 'D', lastRunAt: '2025-12-31T00:00:00.000Z' } // oldest
  ];

  const ordered = sandbox.orderDueReportDocuments_(entries).map(e => e.documentId);
  check('never-run document (A) sorts first',
    ordered[0], 'A');
  check('oldest lastRunAt (D) sorts next',
    ordered[1], 'D');
  check('equal timestamps (B, C) tie-break by documentId, stably',
    ordered.slice(2), ['B', 'C']);
}

{
  const { sandbox } = loadCode();
  check('does not mutate its input array',
    (() => { const input = [{ documentId: 'X', lastRunAt: null }]; sandbox.orderDueReportDocuments_(input); return input; })(),
    [{ documentId: 'X', lastRunAt: null }]);
  check('empty input -> empty output, no throw',
    sandbox.orderDueReportDocuments_([]), []);
  check('undefined input -> empty output, no throw',
    sandbox.orderDueReportDocuments_(undefined), []);
}

// ===================================================== runAddonScheduler_

console.log('runAddonScheduler_() -- only enabled documents considered');

{
  const { sandbox } = loadCode();
  registerEnabled(sandbox, 'DOC_ON');
  sandbox.registerReportDocument_('DOC_OFF', { enabled: false });
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => ({});

  const summary = sandbox.runAddonScheduler_();

  check('both registered documents are "considered"', summary.considered, 2);
  check('only the enabled one is "due"', summary.due, 1);
  check('only the enabled one was actually processed',
    summary.documents.map(d => d.documentId), ['DOC_ON']);
}

console.log('runAddonScheduler_() -- not-due document is skipped, never-run document is due');

{
  const { sandbox } = loadCode();
  setNow(sandbox, 10_000_000);
  registerEnabled(sandbox, 'DOC_NEVER_RUN');
  registerEnabled(sandbox, 'DOC_RECENTLY_RUN', { lastRunAt: new Date(10_000_000 - 1000).toISOString(), intervalHours: 24 });
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => ({});

  const summary = sandbox.runAddonScheduler_();

  check('due count is exactly 1 (only the never-run document)', summary.due, 1);
  check('only the never-run document was processed',
    summary.documents.map(d => d.documentId), ['DOC_NEVER_RUN']);
}

console.log('runAddonScheduler_() -- processes due documents oldest-lastRunAt-first');

{
  const { sandbox } = loadCode();
  setNow(sandbox, 100_000_000);
  registerEnabled(sandbox, 'DOC_OLDEST', { lastRunAt: new Date(1_000_000).toISOString(), intervalHours: 1 });
  registerEnabled(sandbox, 'DOC_NEWER', { lastRunAt: new Date(2_000_000).toISOString(), intervalHours: 1 });
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => ({});

  const summary = sandbox.runAddonScheduler_();

  check('processed in oldest-lastRunAt-first order',
    summary.documents.filter(d => d.outcome === 'succeeded').map(d => d.documentId),
    ['DOC_OLDEST', 'DOC_NEWER']);
}

console.log('runAddonScheduler_() -- one document failure does not stop later due documents');

{
  const { sandbox } = loadCode();
  setNow(sandbox, 100_000_000);
  registerEnabled(sandbox, 'DOC_WILL_FAIL', { lastRunAt: new Date(1_000_000).toISOString(), intervalHours: 1 });
  registerEnabled(sandbox, 'DOC_WILL_SUCCEED', { lastRunAt: new Date(2_000_000).toISOString(), intervalHours: 1 });
  sandbox.downloadAndGroupTdocs_ = () => ({});
  sandbox.DocumentApp.openById = (id) => {
    if (id === 'DOC_WILL_FAIL') throw new Error('simulated open failure');
    return makeFakeDocument(id);
  };

  const summary = sandbox.runAddonScheduler_();

  check('one failed, one succeeded', [summary.failed, summary.succeeded], [1, 1]);
  check('the failing document is reported as failed, with an error message',
    summary.documents.find(d => d.documentId === 'DOC_WILL_FAIL').outcome, 'failed');
  check('the OTHER document still ran and succeeded despite the earlier failure',
    summary.documents.find(d => d.documentId === 'DOC_WILL_SUCCEED').outcome, 'succeeded');
  check('the failed document\'s lastRunAt is unchanged',
    sandbox.getRegisteredReportDocument_('DOC_WILL_FAIL').lastRunAt, new Date(1_000_000).toISOString());
}

console.log('runAddonScheduler_() -- acquires and releases ScriptLock');

{
  const { sandbox } = loadCode();
  let tryLockCalls = 0;
  let releaseLockCalls = 0;
  sandbox.LockService.getScriptLock = () => ({
    tryLock: (ms) => { tryLockCalls++; return true; },
    releaseLock: () => { releaseLockCalls++; }
  });
  sandbox.DocumentApp.openById = (id) => makeFakeDocument(id);
  sandbox.downloadAndGroupTdocs_ = () => ({});

  const summary = sandbox.runAddonScheduler_();

  check('tryLock was called exactly once', tryLockCalls, 1);
  check('releaseLock was called exactly once', releaseLockCalls, 1);
  check('summary reports the lock was acquired', summary.lockAcquired, true);
}

console.log('runAddonScheduler_() -- lock contention exits safely, does no work');

{
  const { sandbox } = loadCode();
  registerEnabled(sandbox, 'DOC_SHOULD_NOT_RUN');
  let releaseLockCalls = 0;
  sandbox.LockService.getScriptLock = () => ({
    tryLock: () => false,
    releaseLock: () => { releaseLockCalls++; }
  });
  let openCalls = 0;
  sandbox.DocumentApp.openById = () => { openCalls++; return makeFakeDocument('x'); };

  const summary = sandbox.runAddonScheduler_();

  check('reports the lock was NOT acquired', summary.lockAcquired, false);
  check('no documents were considered/processed', [summary.considered, summary.due, summary.documents.length], [0, 0, 0]);
  check('no document was ever opened', openCalls, 0);
  check('releaseLock is never called for a lock this run never held', releaseLockCalls, 0);
}

console.log('runAddonScheduler_() -- releases the lock even if a document throws unexpectedly outside the per-document catch');

{
  // Sanity: the lock release is in a `finally` around the WHOLE run, not
  // just the happy path -- simulate a catastrophic failure in
  // listRegisteredReportDocuments_() itself (not a per-document failure).
  const { sandbox } = loadCode();
  let releaseLockCalls = 0;
  sandbox.LockService.getScriptLock = () => ({
    tryLock: () => true,
    releaseLock: () => { releaseLockCalls++; }
  });
  sandbox.listRegisteredReportDocuments_ = () => { throw new Error('simulated catastrophic failure'); };

  try { sandbox.runAddonScheduler_(); } catch (e) { /* expected to propagate */ }

  check('the lock was still released', releaseLockCalls, 1);
}

// ===================================================== runtime budget

console.log('runAddonScheduler_() -- runtime budget defers remaining documents without failing them');

{
  const { sandbox } = loadCode();
  setNow(sandbox, 100_000_000);
  registerEnabled(sandbox, 'DOC_FIRST', { lastRunAt: new Date(1_000_000).toISOString(), intervalHours: 1 });
  registerEnabled(sandbox, 'DOC_SECOND', { lastRunAt: new Date(2_000_000).toISOString(), intervalHours: 1 });
  sandbox.downloadAndGroupTdocs_ = () => ({});

  let opened = [];
  sandbox.DocumentApp.openById = (id) => {
    opened.push(id);
    // The FIRST document's processing "consumes" the runtime budget by
    // advancing the clock past it, simulating a slow real run.
    if (id === 'DOC_FIRST') setNow(sandbox, 100_000_000 + sandbox.REPORT_SCHEDULER_RUNTIME_BUDGET_MS_ + 1);
    return makeFakeDocument(id);
  };

  const summary = sandbox.runAddonScheduler_();

  check('DOC_FIRST was processed (opened) before the budget ran out',
    opened.indexOf('DOC_FIRST') !== -1, true);
  check('DOC_SECOND was NOT opened -- deferred instead',
    opened.indexOf('DOC_SECOND') === -1, true);
  check('DOC_SECOND is reported as deferred/skipped, not failed',
    summary.documents.find(d => d.documentId === 'DOC_SECOND').outcome, 'deferred');
  check('summary.skipped counts the deferred document', summary.skipped, 1);
  check('summary.failed does NOT count the deferred document', summary.failed, 0);
  check('DOC_SECOND\'s lastRunAt is untouched (still its original value, not advanced)',
    sandbox.getRegisteredReportDocument_('DOC_SECOND').lastRunAt, new Date(2_000_000).toISOString());
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
