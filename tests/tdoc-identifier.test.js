/**
 * SA4-IMPL-001 / SA4-IMPL-001A — SA4 TDoc identifier model tests.
 *
 * Tests the central parseSA4DocumentId_()/parseExactSA4DocumentId_() parsers
 * and their backing SA4_TDOC_FAMILIES registry (Code.js, "SA4 TDOC
 * IDENTIFIER MODEL" section) through the public/pure helpers only -- these
 * tests do not reach into the registry array directly, so they exercise the
 * same code path a real caller (normalizeTdoc_, getRevisedTo_,
 * cleanUpWrongEmailDiscussions, ...) does.
 *
 * Covers: all 6 verified families (main + 5 ad-hoc, per SA4-ARCH-005/006),
 * whitespace handling, case handling, family isolation (an unrecognized but
 * shape-plausible prefix like "S4aX" or "A4aX" must NOT be accepted), and
 * (SA4-IMPL-001A) the distinction between "extract from surrounding text"
 * (parseSA4DocumentId_, used by normalizeTdoc_ and friends) and "the whole
 * field must be exactly one identifier" (parseExactSA4DocumentId_, used by
 * cleanUpWrongEmailDiscussions).
 *
 * Run: node tests/tdoc-identifier.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { sandbox } = loadCode();

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

function valid(raw, family, familyKey, seriesCode, yearCode, sequence) {
  return { raw, isValid: true, family, familyKey, seriesCode, yearCode, sequence };
}

const INVALID_EMPTY = { raw: '', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null };

// ------------------------------------------------------- 1. verified families

console.log('parseSA4DocumentId_ -- all 6 verified families');

check('main SA4: S4-260123',
  sandbox.parseSA4DocumentId_('S4-260123'),
  valid('S4-260123', 'main', 'main', null, '26', '0123'));

check('Audio ad-hoc: S4aA260090',
  sandbox.parseSA4DocumentId_('S4aA260090'),
  valid('S4aA260090', 'adhoc', 'audio-adhoc', 'A', '26', '0090'));

check('Plenary/6G ad-hoc: S4aP260098',
  sandbox.parseSA4DocumentId_('S4aP260098'),
  valid('S4aP260098', 'adhoc', 'plenary-adhoc', 'P', '26', '0098'));

check('Video ad-hoc: S4aV200545',
  sandbox.parseSA4DocumentId_('S4aV200545'),
  valid('S4aV200545', 'adhoc', 'video-adhoc', 'V', '20', '0545'));

check('MBS ad-hoc: S4aI240064',
  sandbox.parseSA4DocumentId_('S4aI240064'),
  valid('S4aI240064', 'adhoc', 'mbs-adhoc', 'I', '24', '0064'));

check('RTC ad-hoc: A4aR260097 (note: A4 root, not S4)',
  sandbox.parseSA4DocumentId_('A4aR260097'),
  valid('A4aR260097', 'adhoc', 'rtc-adhoc', 'R', '26', '0097'));

// -------------------------------------------------------------- 2. whitespace

console.log('whitespace handling');

check('leading/trailing whitespace is trimmed, still recognized',
  sandbox.parseSA4DocumentId_(' S4aA260090 '),
  valid('S4aA260090', 'adhoc', 'audio-adhoc', 'A', '26', '0090'));

check('embedded in surrounding free text (e.g. a "Revised to" cell)',
  sandbox.parseSA4DocumentId_('revised to S4-261599'),
  valid('S4-261599', 'main', 'main', null, '26', '1599'));

// -------------------------------------------------------------------- 3. case

console.log('case behavior -- output is always the family\'s real canonical casing');

check('lowercase main input -> uppercase canonical output (existing behavior, preserved)',
  sandbox.parseSA4DocumentId_('s4-260123'),
  valid('S4-260123', 'main', 'main', null, '26', '0123'));

check('all-uppercase ad-hoc input -> mixed-case canonical output, NOT force-uppercased',
  sandbox.parseSA4DocumentId_('S4AA260090'),
  valid('S4aA260090', 'adhoc', 'audio-adhoc', 'A', '26', '0090'));

check('all-lowercase ad-hoc input -> mixed-case canonical output',
  sandbox.parseSA4DocumentId_('s4aa260090'),
  valid('S4aA260090', 'adhoc', 'audio-adhoc', 'A', '26', '0090'));

check('exact real-world casing round-trips unchanged',
  sandbox.parseSA4DocumentId_('S4aP260098'),
  valid('S4aP260098', 'adhoc', 'plenary-adhoc', 'P', '26', '0098'));

// ---------------------------------------------------------------- 4. invalid

console.log('invalid input');

check('main, too few digits (5): S4-12345', sandbox.parseSA4DocumentId_('S4-12345'),
  { raw: 'S4-12345', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null });

check('main, too many digits (7): S4-2601234 (must not truncate-match the first 6)',
  sandbox.parseSA4DocumentId_('S4-2601234'),
  { raw: 'S4-2601234', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null });

check('unknown lookalike prefix: S4aX260001 (shape-plausible, NOT a verified family)',
  sandbox.parseSA4DocumentId_('S4aX260001'),
  { raw: 'S4aX260001', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null });

check('Audio ad-hoc, too few digits (5): S4aA26009',
  sandbox.parseSA4DocumentId_('S4aA26009'),
  { raw: 'S4aA26009', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null });

check('Audio ad-hoc, too many digits (7): S4aA2600900',
  sandbox.parseSA4DocumentId_('S4aA2600900'),
  { raw: 'S4aA2600900', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null });

check('unknown lookalike A4 prefix: A4aX260001 (A4aR is verified, A4aX is not)',
  sandbox.parseSA4DocumentId_('A4aX260001'),
  { raw: 'A4aX260001', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null });

check('plain unrelated text: "foo"', sandbox.parseSA4DocumentId_('foo'),
  { raw: 'foo', isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null });

check('empty string', sandbox.parseSA4DocumentId_(''), INVALID_EMPTY);
check('null', sandbox.parseSA4DocumentId_(null), INVALID_EMPTY);
check('undefined', sandbox.parseSA4DocumentId_(undefined), INVALID_EMPTY);
check('whitespace only', sandbox.parseSA4DocumentId_('   '), INVALID_EMPTY);

// --------------------------------------------------------- 5. family isolation

console.log('family isolation (registry is authoritative -- tested through the public parser only)');

// Every verified family's own real example must be recognized...
[
  ['S4-260123', true], ['S4aA260090', true], ['S4aP260098', true],
  ['S4aV200545', true], ['S4aI240064', true], ['A4aR260097', true]
].forEach(([id, expected]) => {
  check(`${id} is recognized (known family)`, sandbox.parseSA4DocumentId_(id).isValid, expected);
});

// ...but a plausible-but-unverified prefix for EVERY letter slot must not be,
// including letters that are one edit away from a real family (guards
// against a too-loose/generic regex accidentally accepting lookalikes).
[
  'S4aB260001', 'S4aC260001', 'S4aD260001', 'S4aE260001', 'S4aF260001',
  'S4aM260001', // the "obvious but wrong" guess for MBS (real prefix is S4aI)
  'S4aR260001', // the "obvious but wrong" guess for RTC (real prefix is A4aR)
  'A4aA260001', 'A4aV260001', 'A4aP260001'
].forEach(id => {
  check(`${id} is rejected (unverified/incorrect lookalike prefix)`, sandbox.parseSA4DocumentId_(id).isValid, false);
});

// ------------------------------------------- 6. exact-match (SA4-IMPL-001A)

console.log('parseExactSA4DocumentId_ -- whole-field-only match (all 6 verified families)');

[
  ['S4-260123', 'S4-260123'],
  [' S4aA260090 ', 'S4aA260090'],
  ['S4aP260098', 'S4aP260098'],
  ['S4aV200545', 'S4aV200545'],
  ['S4aI240064', 'S4aI240064'],
  ['A4aR260097', 'A4aR260097']
].forEach(([input, expectedRaw]) => {
  const result = sandbox.parseExactSA4DocumentId_(input);
  check(`exact: ${JSON.stringify(input)} -> valid, raw="${expectedRaw}"`,
    { isValid: result.isValid, raw: result.raw },
    { isValid: true, raw: expectedRaw });
});

console.log('parseExactSA4DocumentId_ -- surrounding text is rejected (unlike parseSA4DocumentId_)');

[
  'TDoc S4-260123',
  'S4-260123 revised',
  'TDoc S4aA260090',
  'S4aA260090 revised',
  'foo A4aR260097',
  'A4aR260097 foo'
].forEach(input => {
  check(`exact: ${JSON.stringify(input)} is rejected (identifier is embedded, not the whole field)`,
    sandbox.parseExactSA4DocumentId_(input).isValid, false);
  // The same text IS accepted by the search-based parser -- proving this is
  // a real behavioral distinction between the two functions, not just two
  // names for the same thing.
  check(`  -> but parseSA4DocumentId_ (search) DOES find it in the same text`,
    sandbox.parseSA4DocumentId_(input).isValid, true);
});

check('exact: unknown lookalike family is still rejected (S4aX260001)',
  sandbox.parseExactSA4DocumentId_('S4aX260001').isValid, false);
check('exact: empty string is rejected', sandbox.parseExactSA4DocumentId_('').isValid, false);

// ------------------------------------ 7. cleanUpWrongEmailDiscussions() usage

console.log('source-structure assertion: cleanUpWrongEmailDiscussions() uses the EXACT helper');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const startMatch = source.match(/^function cleanUpWrongEmailDiscussions\(\)/m);
  if (!startMatch) {
    failures++;
    console.log('  FAIL could not locate "function cleanUpWrongEmailDiscussions()" in Code.js');
  } else {
    const startIndex = startMatch.index;
    const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnRe.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnRe.exec(source);
    const endIndex = next ? next.index : source.length;
    const rawBody = source.slice(startIndex, endIndex);

    // Strip `//` line comments first -- this function's own migration
    // comment mentions "parseSA4DocumentId_()" by name in prose (explaining
    // what it is NOT using), which would otherwise false-positive the
    // "no direct call" check below. Source-structure assertions should
    // reflect actual code, not documentation text.
    const body = rawBody.replace(/\/\/.*$/gm, '');

    check('function body calls parseExactSA4DocumentId_()',
      /\bparseExactSA4DocumentId_\s*\(/.test(body), true);
    // Strip every parseExactSA4DocumentId_(...) call out too, so a leftover
    // "parseSA4DocumentId_(" can only mean the unanchored function was
    // called directly (not a substring match against the Exact function's
    // own name, which contains "parseSA4DocumentId_" as a substring).
    const bodyWithoutExactCalls = body.replace(/parseExactSA4DocumentId_\s*\(/g, '');
    check('function body does NOT call the unanchored parseSA4DocumentId_() directly',
      /\bparseSA4DocumentId_\s*\(/.test(bodyWithoutExactCalls), false);
  }
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll TDoc-identifier tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
