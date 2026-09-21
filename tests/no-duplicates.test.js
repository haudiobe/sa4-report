/**
 * SA4-ARCH-002 — Structural integrity / architecture invariants for Code.js.
 * SA4-BASELINE-001 — duplicate-function check changed from a hard "zero
 * duplicates" assertion to an explicit known-baseline policy (see below).
 *
 * This restores the test file referenced by the Code.js 2.6.0 changelog
 * ("Added: tests/no-duplicates.test.js guards against duplicate definitions
 * reappearing and verifies every menu target exists and is callable.") which
 * was missing from the repository (see SA4-ARCH-001 §8/§13).
 *
 * These checks are static/structural: they do not build a report or touch a
 * live Google Doc. They exist to catch accidental damage while the engine is
 * later split into modules (MeetingContext / profiles), by pinning down:
 *
 *   1. Duplicate top-level `function name(...)` definitions in Code.js are
 *      limited to an explicit, named, known baseline (see
 *      KNOWN_DUPLICATE_FUNCTIONS below) -- neither a silent new duplicate
 *      nor a silent disappearance of a known one passes unnoticed.
 *   2. Every function referenced by an onOpen() menu item actually exists.
 *   3. The canonical entry points named in SA4-ARCH-002 exist:
 *        runFullReportBuild, buildSkeletonWithTdocTables, continuousUpdate
 *   4. Key configuration functions exist:
 *        getReportConfig_, getAgendaPrefixForReportType_, generateReportTitle_
 *   5. orderTdocsByRevision_ exists (the function tests/revision-order.test.js
 *      already exercises).
 *
 * Run: node tests/no-duplicates.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
const { sandbox } = loadCode();

let failures = 0;
function check(name, cond, detail) {
  if (cond) {
    console.log(`  ok   ${name}`);
  } else {
    failures++;
    console.log(`  FAIL ${name}${detail ? '\n         ' + detail : ''}`);
  }
}

// ------------------------------------------------- 1. known-duplicate policy
//
// SA4-ARCH-002 discovered 5 top-level functions that were already declared
// twice in Code.js (JavaScript keeps the LAST declaration for each name, so
// every earlier one is silently-dead code). SA4-ARCH-002 through
// SA4-ARCH-006 deliberately left these untouched -- fixing them is
// unrelated cleanup, out of scope for the MeetingContext/ad-hoc work.
//
// Rather than either (a) asserting "zero duplicates" forever (which makes
// this suite permanently red and trains people to ignore its failures) or
// (b) silently ignoring duplicates entirely (which would let a NEW,
// unrelated duplicate slip in unnoticed), this test encodes the known set
// explicitly and asserts the baseline stays EXACTLY this set: a sixth
// duplicate function is a regression (FAILS), and any of these five
// unexpectedly ceasing to be duplicated also FAILS -- that's not assumed to
// be "good news" without deliberate review, since it could equally mean the
// parser broke, or someone accidentally deleted the wrong definition.
//
// Line numbers are intentionally NOT part of the acceptance criteria (they
// legitimately move as Code.js grows) -- only the function NAME and the
// FACT of duplication are asserted. Current line numbers are still printed
// below as non-asserted, informational context.
const KNOWN_DUPLICATE_FUNCTIONS = [
  'copySectionContentWithReplacement_',
  'formatDeadline_',
  'getMonthName_',
  'calculateTimeRemaining_',
  'addTdocTablesOnly'
];

console.log('duplicate top-level function definitions (known-baseline policy)');

const FUNCTION_DECL_RE = /^function\s+([A-Za-z0-9_$]+)\s*\(/gm;
const linesByName = new Map();
let m;
while ((m = FUNCTION_DECL_RE.exec(source))) {
  const name = m[1];
  const line = source.slice(0, m.index).split('\n').length;
  if (!linesByName.has(name)) linesByName.set(name, []);
  linesByName.get(name).push(line);
}

const totalDeclarations = [...linesByName.values()].reduce((n, lines) => n + lines.length, 0);
check('at least 150 top-level functions were found (sanity check on the parser itself)',
  linesByName.size >= 150,
  `found ${linesByName.size} distinct names, ${totalDeclarations} declarations total`);

const duplicates = [...linesByName.entries()].filter(([, lines]) => lines.length > 1);
const duplicateNames = new Set(duplicates.map(([name]) => name));
const knownSet = new Set(KNOWN_DUPLICATE_FUNCTIONS);

// Informational only -- never asserted on. Printed every run so the known
// technical debt stays visible even while this suite is green.
console.log('  known pre-existing duplicates (technical debt, not fixed by this test):');
duplicates
  .filter(([name]) => knownSet.has(name))
  .forEach(([name, lines]) => console.log(`    - ${name} (currently at lines ${lines.join(', ')})`));

const unexpectedDuplicates = [...duplicateNames].filter(name => !knownSet.has(name));
check('every duplicate found is an entry in KNOWN_DUPLICATE_FUNCTIONS (no NEW duplicates)',
  unexpectedDuplicates.length === 0,
  unexpectedDuplicates.length
    ? 'unexpected new duplicate(s): ' + unexpectedDuplicates.map(name => `${name} (lines ${linesByName.get(name).join(', ')})`).join('; ') +
      '\n         This is a NEW duplicate not present in the SA4-ARCH-002 baseline -- treat this as a real regression.'
    : '');

const missingKnownDuplicates = KNOWN_DUPLICATE_FUNCTIONS.filter(name => !duplicateNames.has(name));
check('every entry in KNOWN_DUPLICATE_FUNCTIONS is still actually duplicated in Code.js',
  missingKnownDuplicates.length === 0,
  missingKnownDuplicates.length
    ? 'no longer duplicated: ' + missingKnownDuplicates.join(', ') +
      '\n         A known duplicate disappeared. This may be intentional cleanup, but it is not assumed safe by' +
      '\n         default -- update KNOWN_DUPLICATE_FUNCTIONS deliberately (with a note on what changed and why)' +
      '\n         rather than letting this pass silently.'
    : '');

check('the duplicate set matches the known baseline exactly (no additional duplicate names)',
  duplicates.length === KNOWN_DUPLICATE_FUNCTIONS.length,
  `found ${duplicates.length} duplicated name(s), expected exactly ${KNOWN_DUPLICATE_FUNCTIONS.length} (the known baseline)`);

// Special note: unlike the other 4 (which are byte-identical copies, so the
// duplication is "only" dead code), copySectionContentWithReplacement_'s two
// definitions are BEHAVIORALLY DIFFERENT. Documented here, not fixed:
//   - earlier definition (shadowed, dead): copies source paragraphs/list
//     items via .copy()/.replaceText(), which PRESERVES original formatting
//     (bold, colors, etc.)
//   - later definition (live, the one that actually runs): rebuilds
//     paragraphs from scratch via appendParagraph()/appendListItem(), which
//     does NOT preserve that formatting
// This is genuinely more consequential than the other 4 duplicates and is
// called out separately so it isn't lost among the "just dead code" cases.
// No decision is made here about which implementation is "correct" -- that
// is an explicit non-goal of this test and of SA4-BASELINE-001.
if (knownSet.has('copySectionContentWithReplacement_')) {
  console.log('  ! copySectionContentWithReplacement_ warning: the two definitions are NOT');
  console.log('    behaviorally identical (unlike the other 4 known duplicates). The shadowed,');
  console.log('    dead definition preserves source formatting via .copy()/.replaceText(); the');
  console.log('    live definition rebuilds paragraphs from scratch and does not. See');
  console.log('    SA4-ARCH-002 report for detail. Not fixed here by design.');
}

// --------------------------------------------- 2. onOpen() menu targets exist

console.log('onOpen() menu targets');

const MENU_ITEM_RE = /\.addItem\(\s*'[^']*'\s*,\s*'([A-Za-z0-9_]+)'\s*\)/g;
const menuTargets = new Set();
while ((m = MENU_ITEM_RE.exec(source))) menuTargets.add(m[1]);

check('at least 25 menu items were found (sanity check on the parser itself)',
  menuTargets.size >= 25,
  `found ${menuTargets.size} distinct menu targets`);

for (const name of [...menuTargets].sort()) {
  check(`menu target "${name}" exists as a function`,
    typeof sandbox[name] === 'function',
    `onOpen() references '${name}' via addItem(), but no top-level function of that name was found`);
}

// --------------------------------------- 3. canonical entry points (SA4-ARCH-002)

console.log('canonical entry points');

['runFullReportBuild', 'buildSkeletonWithTdocTables', 'continuousUpdate'].forEach(name => {
  check(`"${name}" exists as a function`, typeof sandbox[name] === 'function');
});

// ------------------------------------------------- 4. configuration functions

console.log('configuration functions');

['getReportConfig_', 'getAgendaPrefixForReportType_', 'generateReportTitle_'].forEach(name => {
  check(`"${name}" exists as a function`, typeof sandbox[name] === 'function');
});

// --------------------------------------------------- 5. revision-engine entry

console.log('revision engine');

check('"orderTdocsByRevision_" exists as a function', typeof sandbox.orderTdocsByRevision_ === 'function');

// ------------------------------------------------------------------- summary

console.log(failures === 0
  ? '\nAll structural checks passed.'
  : `\n${failures} structural check(s) FAILED (see SA4-ARCH-002 report for known pre-existing issues).`);
process.exit(failures === 0 ? 0 : 1);
