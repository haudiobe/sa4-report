/**
 * SA4-IMPL-006 — Pure agenda projection tests.
 *
 * Tests projectAgendaItems_(agendaItems, selector) (Code.js, defined just
 * after agendaSelectorMatches_()) as a pure, standalone function -- no
 * Google APIs, no PropertiesService, no DocumentApp involvement.
 *
 * projectAgendaItems_() has NO production caller as of SA4-IMPL-006 -- it
 * exists as a tested abstraction, ready for a later task to migrate
 * parseAgendaForReport_() onto it. This suite proves the semantics
 * documented in the SA4-ARCH-007 report: no ancestor retention, no
 * descendant expansion beyond what prefix startsWith already provides, no
 * node synthesis, no sorting, and itemList as exact-selected-items-only.
 *
 * Run: node tests/agenda-projection.test.js
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

function expectThrows(name, fn, messageSubstring) {
  try {
    fn();
    failures++;
    console.log(`  FAIL ${name}\n         expected a throw, but none occurred`);
  } catch (e) {
    if (messageSubstring && e.message.indexOf(messageSubstring) === -1) {
      failures++;
      console.log(`  FAIL ${name}\n         expected message to contain ${JSON.stringify(messageSubstring)}\n         actual message: ${e.message}`);
    } else {
      console.log(`  ok   ${name}`);
    }
  }
}

function item(number, extra) {
  return Object.assign({ number: number, title: 'Title for ' + number, level: 1, heading: 'H2', text: '' }, extra || {});
}
function numbers(items) {
  return items.map(i => i.number);
}

// ==================================================== 1. invalid input ====

console.log('projectAgendaItems_ -- invalid agendaItems input');

expectThrows('null throws', () => sandbox.projectAgendaItems_(null, { mode: 'all' }),
  'expected an array');
expectThrows('undefined throws', () => sandbox.projectAgendaItems_(undefined, { mode: 'all' }),
  'expected an array');
expectThrows('a plain object throws', () => sandbox.projectAgendaItems_({}, { mode: 'all' }),
  'expected an array');
expectThrows('a string throws', () => sandbox.projectAgendaItems_('7.1', { mode: 'all' }),
  'expected an array');
expectThrows('a number throws', () => sandbox.projectAgendaItems_(5, { mode: 'all' }),
  'expected an array');

check('empty array -> empty array', sandbox.projectAgendaItems_([], { mode: 'all' }), []);
check('empty array with prefix selector -> empty array', sandbox.projectAgendaItems_([], { mode: 'prefix', value: '7.' }), []);

console.log('projectAgendaItems_ -- invalid selector (delegated to normalizeAgendaSelector_/agendaSelectorMatches_, not duplicated)');

expectThrows('invalid selector mode throws (via the unmodified agendaSelectorMatches_)',
  () => sandbox.projectAgendaItems_([item('7.1')], { mode: 'bogus' }),
  'unsupported mode');

// ============================================== 2. main Audio "7." prefix =

console.log('projectAgendaItems_ -- main Audio prefix "7." (real production value)');

{
  const source = ['1', '5', '6', '7', '7.1', '7.1.1', '7.2', '7.10', '8', '17'].map(n => item(n));
  const result = sandbox.projectAgendaItems_(source, { mode: 'prefix', value: '7.' });

  check('retains exactly the sub-items, in source order',
    numbers(result), ['7.1', '7.1.1', '7.2', '7.10']);
  check('the bare parent "7" is deliberately NOT retained (ARCH-007: the sole real consumer discards it anyway)',
    numbers(result).indexOf('7'), -1);
  check('unrelated top-level items (1,5,6,8) and the "17" false-prefix-collision are excluded',
    numbers(result).some(n => ['1', '5', '6', '8', '17'].indexOf(n) !== -1), false);
}

// ================================================== 3. main 6G "11." prefix

console.log('projectAgendaItems_ -- main 6G prefix "11."');

{
  const source = ['1', '10', '11', '11.1', '11.1.1', '11.2', '12', '111'].map(n => item(n));
  const result = sandbox.projectAgendaItems_(source, { mode: 'prefix', value: '11.' });

  check('retains exactly the sub-items, in source order',
    numbers(result), ['11.1', '11.1.1', '11.2']);
  check('bare "11" is NOT retained', numbers(result).indexOf('11'), -1);
  check('"111" is correctly excluded (does not start with "11.")', numbers(result).indexOf('111'), -1);
}

// ===================================================== 4. ULBC-MED "all" ==

console.log('projectAgendaItems_ -- ULBC-MED {mode:"all"}');

{
  const numbersList = ['1', '2', '3', '4', '4.1', '4.2', '4.3', '4.4', '4.5', '4.6', '4.7', '4.8', '4.9', '4.10', '5', '6'];
  const source = numbersList.map(n => item(n));
  const result = sandbox.projectAgendaItems_(source, { mode: 'all' });

  check('every real agenda item is retained, in source order', numbers(result), numbersList);
  check('projection length equals input length (no item excluded)', result.length, source.length);
}

// =============================================== 5. itemList exact-only ===
//
// Current recurring-telco evidence (SA4-ARCH-006) supports exact selected
// slots. Nested child behavior under an itemList-selected slot has not been
// observed. Use prefix mode when subtree selection is explicitly required.

console.log('projectAgendaItems_ -- itemList exact-only (evidence caveat: see comment above)');

{
  const source = ['1', '1.1', '1.5', '1.5.1', '1.5.2', '1.6', '1.6.1', '1.7', '2'].map(n => item(n));
  const result = sandbox.projectAgendaItems_(source, { mode: 'itemList', value: ['1.5', '1.6'] });

  check('retains exactly the two selected slots, nothing else', numbers(result), ['1.5', '1.6']);
  check('descendants of a selected slot are NOT pulled in (1.5.1, 1.5.2, 1.6.1 excluded)',
    numbers(result).some(n => ['1.5.1', '1.5.2', '1.6.1'].indexOf(n) !== -1), false);
  check('the parent "1" is NOT pulled in', numbers(result).indexOf('1'), -1);
  check('unrelated siblings (1.1, 1.7, 2) are excluded', numbers(result).some(n => ['1.1', '1.7', '2'].indexOf(n) !== -1), false);
}

// ============================================== 6. source-order preserved

console.log('projectAgendaItems_ -- preserves source order, never sorts');

{
  // Deliberately non-numeric source order.
  const source = ['7.2', '7.1', '7.10'].map(n => item(n));
  const result = sandbox.projectAgendaItems_(source, { mode: 'prefix', value: '7.' });
  check('output order mirrors input order exactly (7.2, 7.1, 7.10) -- NOT sorted numerically or lexically',
    numbers(result), ['7.2', '7.1', '7.10']);
}

{
  const source = ['4.10', '4.1', '4.2'].map(n => item(n));
  const result = sandbox.projectAgendaItems_(source, { mode: 'all' });
  check('"all" mode also preserves source order unchanged', numbers(result), ['4.10', '4.1', '4.2']);
}

// ============================================== 7. no ancestor synthesis ==

console.log('projectAgendaItems_ -- no ancestor/node synthesis');

{
  const source = ['7.1', '7.2'].map(n => item(n)); // no bare "7" in the source at all
  const result = sandbox.projectAgendaItems_(source, { mode: 'prefix', value: '7.' });
  check('no synthetic "7" appears -- only what was actually in the source', numbers(result), ['7.1', '7.2']);
}

// =========================================== 8. object-reference identity =

console.log('projectAgendaItems_ -- preserves original object references, does not clone/rewrite');

{
  const sentinelHeading = { thisIsTheSentinel: true };
  const original = item('7.1', { heading: sentinelHeading, text: 'some collected paragraph text' });
  const other = item('8.1');
  const source = [original, other];

  const result = sandbox.projectAgendaItems_(source, { mode: 'prefix', value: '7.' });

  check('exactly one item retained', result.length, 1);
  check('retained item is the EXACT SAME object reference as the original (===)', result[0] === original, true);
  check('the sentinel heading object reference survives untouched', result[0].heading === sentinelHeading, true);
  check('input array itself is not mutated (still has 2 items)', source.length, 2);
  check('input array item order is not mutated', source[0] === original && source[1] === other, true);
}

// ==================================================== 9. null/blank numbers

console.log('projectAgendaItems_ -- null/blank agenda numbers rely on agendaSelectorMatches_(), nothing invented');

{
  const source = [item(null), item(undefined), item(''), item('   '), item('7.1')];

  check('"all" mode: blank/null/undefined numbers are excluded (agendaSelectorMatches_ semantics, unchanged)',
    numbers(sandbox.projectAgendaItems_(source, { mode: 'all' })), ['7.1']);
  check('"prefix" mode: same blank/null/undefined entries excluded, "7.1" retained',
    numbers(sandbox.projectAgendaItems_(source, { mode: 'prefix', value: '7.' })), ['7.1']);
}

{
  // A fully null/undefined ARRAY ENTRY (not just a blank `number` field) --
  // must not throw; treated the same as a blank agenda number.
  const source = [null, item('7.1'), undefined];
  check('a null/undefined array entry itself does not throw and is simply excluded',
    numbers(sandbox.projectAgendaItems_(source, { mode: 'prefix', value: '7.' })), ['7.1']);
}

// ============================ 10. no production caller (source-structure) =

// SA4-IMPL-007 migrated parseAgendaForReport_() onto projectAgendaItems_()
// (its ZIP/AGENDA_TDOC branch only -- see tests/canonical-agenda-projection
// .test.js for that migration's own source-structure and behavior
// coverage). It is deliberately EXCLUDED from the "still has no caller"
// list below; every other function here remains unmigrated exactly as
// SA4-IMPL-006 left it.
console.log('source-structure assertion: projectAgendaItems_() has no OTHER production caller (parseAgendaForReport_ migrated by SA4-IMPL-007)');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  function extractBody(src, fnName) {
    const startMatch = src.match(new RegExp('^function ' + fnName + '\\(', 'm'));
    if (!startMatch) return null;
    const startIndex = startMatch.index;
    const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnRe.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnRe.exec(src);
    const endIndex = next ? next.index : src.length;
    return src.slice(startIndex, endIndex);
  }

  // Strip projectAgendaItems_()'s own definition/doc-comment first, so
  // mentions of its own name in its own JSDoc don't count as a "caller".
  const withoutOwnDefinition = source.replace(/function projectAgendaItems_\([\s\S]*?\n}\n/, '');

  ['parseAgendaFromZippedTdoc', 'parseAgendaDocumentById',
   'parseAgendaStructureWithText_', 'parseAgendaFromHeadings_', 'parseAgendaFromTables_',
   'buildSkeletonWithTdocTables', 'downloadAndGroupTdocs_'].forEach(fnName => {
    const body = extractBody(withoutOwnDefinition, fnName);
    if (!body) {
      failures++;
      console.log(`  FAIL could not locate "function ${fnName}(" in Code.js`);
      return;
    }
    check(`${fnName}() does NOT call projectAgendaItems_() (unmigrated, as required by SA4-IMPL-006)`,
      /\bprojectAgendaItems_\s*\(/.test(body), false);
  });

  const migrated = extractBody(withoutOwnDefinition, 'parseAgendaForReport_');
  if (!migrated) {
    failures++;
    console.log('  FAIL could not locate "function parseAgendaForReport_(" in Code.js');
  } else {
    check('parseAgendaForReport_() DOES call projectAgendaItems_() (migrated by SA4-IMPL-007; see tests/canonical-agenda-projection.test.js)',
      /\bprojectAgendaItems_\s*\(/.test(migrated), true);
  }
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll agenda-projection tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
