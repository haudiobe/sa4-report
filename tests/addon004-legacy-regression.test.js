/**
 * ADDON-004 — legacy regression: continuousUpdate() (the bound-script /
 * interactive entry point) must still use its existing DocumentLock path,
 * completely unaffected by the background scheduler's ScriptLock.
 *
 * tests/perf003b-lock.test.js already covers continuousUpdate()'s lock
 * behavior in depth (acquire/release, contention, source-structure proof
 * scoped to continuousUpdate()+updateReportIncremental() specifically,
 * updated in this same stage to exclude the scheduler's own, separate
 * getScriptLock() use). This file adds the one thing ADDON-004 introduces
 * that suite didn't anticipate: continuousUpdateCore_() now takes an
 * optional `context` parameter, and continuousUpdate() must still call it
 * with NO context at all.
 *
 * Run: node tests/addon004-legacy-regression.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

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

console.log('source-structure: continuousUpdate() calls continuousUpdateCore_() with NO arguments');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const startMatch = source.match(/^function continuousUpdate\(\)/m);
  const startIndex = startMatch ? startMatch.index : -1;
  check('continuousUpdate() (zero-arg legacy entry point) was located', startIndex >= 0, true);

  const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
  nextFnRe.lastIndex = startIndex + (startMatch ? startMatch[0].length : 0);
  const next = nextFnRe.exec(source);
  const body = source.slice(startIndex, next ? next.index : source.length);

  check('continuousUpdate() calls continuousUpdateCore_() with no arguments',
    /continuousUpdateCore_\(\);/.test(body), true);
  check('continuousUpdate() still acquires LockService.getDocumentLock() (unchanged)',
    /LockService\.getDocumentLock\(\)/.test(body), true);
  check('continuousUpdate() does NOT acquire ScriptLock',
    /LockService\.getScriptLock\(\)/.test(body), false);
}

console.log('runtime: continuousUpdate() acquires/releases DocumentLock, never touches ScriptLock');

{
  const { sandbox } = loadCode();
  let documentLockCalls = 0;
  let scriptLockCalls = 0;
  let documentTryLockCalls = 0;
  let documentReleaseCalls = 0;

  sandbox.LockService = {
    getDocumentLock: () => { documentLockCalls++; return { tryLock: () => { documentTryLockCalls++; return true; }, releaseLock: () => { documentReleaseCalls++; } }; },
    getScriptLock: () => { scriptLockCalls++; return { tryLock: () => true, releaseLock: () => {} }; }
  };
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ getTables: () => [] }), getId: () => 'LEGACY_DOC' });
  sandbox.downloadAndGroupTdocs_ = () => ({});

  sandbox.continuousUpdate();

  check('LockService.getDocumentLock() was used', documentLockCalls, 1);
  check('DocumentLock.tryLock() was called', documentTryLockCalls, 1);
  check('DocumentLock.releaseLock() was called', documentReleaseCalls, 1);
  check('LockService.getScriptLock() was NEVER called by continuousUpdate()', scriptLockCalls, 0);
}

console.log('runtime: continuousUpdate() still uses DocumentApp.getActiveDocument() and real Document Properties (no context = legacy backend)');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Video' } });
  let activeDocumentCalls = 0;
  sandbox.DocumentApp.getActiveDocument = () => { activeDocumentCalls++; return { getBody: () => ({ getTables: () => [] }), getId: () => 'LEGACY_DOC' }; };
  sandbox.downloadAndGroupTdocs_ = () => ({});

  sandbox.continuousUpdate();

  check('DocumentApp.getActiveDocument() was used (legacy path, unchanged)', activeDocumentCalls > 0, true);
  // REVISION_MAP should land in real Document Properties, not any central store.
  check('state was written to real Document Properties (legacy backend)',
    docProps.getProperty('REVISION_MAP') !== null, true);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
