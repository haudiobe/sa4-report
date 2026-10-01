/**
 * PERF-003 (Part B) — skip global formatting when nothing structural
 * changed.
 *
 * PERF-002's controlled production measurement found removeRowHeightAndSpacing()
 * costing ~16.4s of the ~73s no-change baseline: an unconditional,
 * full-document pass writing row.setMinimumHeight(0)/paragraph spacing/
 * cell widths across EVERY table, every run, regardless of whether
 * anything changed. Characterization: this pass exists because Apps
 * Script's own table/row/cell/paragraph insertion APIs default to
 * non-zero spacing/height, so a newly INSERTED table or paragraph needs
 * it -- content whose TEXT changed in an EXISTING row/cell (status
 * updates, e-mail discussion, revision cell rendering) never does.
 *
 * This suite proves the extracted decision (shouldReformatAfterUpdate_())
 * that continuousUpdate() now gates its removeRowHeightAndSpacing() call
 * with, and confirms (via source-structure) that every OTHER call site of
 * removeRowHeightAndSpacing() remains unconditional -- full report build/
 * "Update All"/manual formatting menu items are untouched.
 *
 * Run: node tests/perf003-formatting-skip.test.js
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

// ==================== 1. shouldReformatAfterUpdate_() -- pure decision ====

console.log('shouldReformatAfterUpdate_() -- pure structural-change decision');

{
  const { sandbox } = loadCode();
  const fn = sandbox.shouldReformatAfterUpdate_;

  check('1. no-change run (0 new, 0 moved, 0 inserted, 0 abstract rows) -> SKIP formatting', fn(0, 0, 0, 0), false);
  check('2. a newly inserted TDoc (newTdocsAdded=1) -> formatting REQUIRED', fn(1, 0, 0, 0), true);
  check('3. a structural revision-table movement (revisionsMoved=1) -> formatting REQUIRED', fn(0, 1, 0, 0), true);
  check('a newly inserted revision-linked table (insertRevisedDocTablesAfter_) -> formatting REQUIRED', fn(0, 0, 1, 0), true);
  check('PERF-003B: an abstract row inserted into an EXISTING table (fetchAndAddAbstract_) -> formatting REQUIRED -- the exact PERF-002 measured case (0 new, 0 moved, 0 inserted, 1 abstract row)', fn(0, 0, 0, 1), true);
  check('multiple simultaneous structural changes -> still just REQUIRED (no double-counting concern)', fn(3, 2, 1, 1), true);
  check('undefined/missing arguments do not throw and default to "no change"', fn(undefined, undefined, undefined, undefined), false);
  check('null arguments do not throw and default to "no change"', fn(null, null, null, null), false);
}

// ============ 2. source-structure: continuousUpdate() wires real inputs ===

console.log('source-structure: continuousUpdateCore_() feeds its REAL newTdocsAdded/rev.moved/counter into the decision');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  // PERF-003B (Part 2): continuousUpdate() is now a thin lock-acquiring
  // public wrapper (see tests/perf003b-lock.test.js) -- the actual logic
  // that computes/consumes these signals lives in continuousUpdateCore_().
  // ADDON-004: continuousUpdateCore_() now takes an optional `context`
  // parameter -- locator tolerates an optional argument name instead of
  // requiring an exactly-empty `()`.
  const fnStartMatch = source.match(/^function continuousUpdateCore_\([A-Za-z0-9_$]*\)/m);
  const fnStart = fnStartMatch ? fnStartMatch.index : -1;
  const nextFnMatch = fnStart >= 0 ? source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m) : null;
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = fnStart >= 0 ? source.slice(fnStart, fnEnd) : '';

  check('continuousUpdateCore_() was located in Code.js', fnStart >= 0, true);
  check('continuousUpdate() calls shouldReformatAfterUpdate_(newTdocsAdded, rev.moved, ...)',
    /shouldReformatAfterUpdate_\(\s*newTdocsAdded,\s*rev\.moved,/.test(body), true);
  check('continuousUpdate() reads the revision-linked-table-insertion counter as the third argument',
    /perfCounterValue_\('structural: new revision-linked tables inserted \(insertRevisedDocTablesAfter_\)'\)/.test(body), true);
  check('PERF-003B: continuousUpdate() reads the abstract-row-insertion counter as the fourth argument (closes the gap the PERF-002 measured run exposed)',
    /perfCounterValue_\('structural: abstract row inserted \(fetchAndAddAbstract_\)'\)/.test(body), true);
  check('the formatting call is inside an if() gated by the decision result (not called unconditionally any more here)',
    /if\s*\(structuralChangeThisRun\)\s*\{\s*perfTimed_\('formatting \(removeRowHeightAndSpacing\)'/.test(body), true);
}

// ===== 2b. source-structure: fetchAndAddAbstract_ counts its own row insert =

console.log('source-structure: fetchAndAddAbstract_() counts the structural row it inserts into an EXISTING table');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function fetchAndAddAbstract_(');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  check('fetchAndAddAbstract_ calls perfCount_ for the abstract-row-inserted label before table.insertTableRow(...)',
    /perfCount_\('structural: abstract row inserted \(fetchAndAddAbstract_\)'\);\s*\n\s*const abstractRow = table\.insertTableRow/.test(body), true);
}

// ======= 3. source-structure: every OTHER call site remains unconditional =

console.log('source-structure: full-build/"Update All"/manual formatting call sites remain UNCONDITIONAL (item 4)');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  // Every literal "removeRowHeightAndSpacing(...)" call site in the file,
  // with a few characters of context to classify it. ADDON-004:
  // removeRowHeightAndSpacing() now takes an optional `context` argument
  // -- the ONE call site inside continuousUpdateCore_() passes it
  // (`removeRowHeightAndSpacing(context)`); every other, still-interactive
  // call site is unchanged and still calls it bare
  // (`removeRowHeightAndSpacing();`). The pattern below tolerates either.
  const callSites = [];
  const pattern = /removeRowHeightAndSpacing\([A-Za-z0-9_$]*\)/g;
  let m;
  while ((m = pattern.exec(source)) !== null) {
    const lineStart = source.lastIndexOf('\n', m.index) + 1;
    const lineEnd = source.indexOf('\n', m.index);
    callSites.push(source.slice(lineStart, lineEnd === -1 ? undefined : lineEnd).trim());
  }
  // Excludes: the function's own `function removeRowHeightAndSpacing(context) {`
  // definition line, and any line that is purely a comment reference.
  const realCallSites = callSites.filter(line =>
    !line.startsWith('function removeRowHeightAndSpacing') &&
    !line.startsWith('//') && !line.startsWith('*'));

  check('at least 6 other unconditional call sites still exist besides continuousUpdate()\'s', realCallSites.length >= 6, true);
  check('exactly ONE call site is the new conditional one (inside the perfTimed_ formatting branch)',
    realCallSites.filter(line => line.includes("perfTimed_('formatting")).length, 1);
  // TEMPLATE-002C: one more call site is deliberately conditional -- the
  // skeleton build skips its own pass when Run Full Report Build asks it to,
  // because Full Build formats once after its enrichment phases
  // (tests/template002c-fullbuild-creator.test.js). Every other call site is
  // still a bare, unconditional call.
  const FULL_BUILD_SKIP = 'if (!(options && options.skipFormatting)) removeRowHeightAndSpacing();';
  check('exactly ONE call site is the Full Build option (the skeleton build)', realCallSites.filter(line => line === FULL_BUILD_SKIP).length, 1);
  check('every OTHER call site is a bare, unconditional call (not wrapped in a new if-condition)',
    realCallSites.filter(line => !line.includes("perfTimed_('formatting") && line !== FULL_BUILD_SKIP).every(line => line === 'removeRowHeightAndSpacing();'),
    true);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
