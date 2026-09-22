/**
 * SA4-IMPL-004 — Agenda selector model tests.
 *
 * Tests normalizeAgendaSelector_() and agendaSelectorMatches_() (Code.js,
 * defined just above getMeetingContext_()) as pure, standalone functions --
 * no Google APIs, no PropertiesService, no MeetingContext construction.
 *
 * IMPORTANT: as of SA4-IMPL-004, agendaSelectorMatches_() has NO production
 * filtering caller. downloadAndGroupTdocs_(), parseAgendaForReport_(), and
 * buildSkeletonWithTdocTables() are all untouched by this task and still
 * filter using the old agendaItem.startsWith(agendaPrefix) inline logic.
 * This suite exists to prove the NEW abstraction matches that EXISTING
 * behavior exactly for 'prefix' mode, ahead of a later migration task
 * actually wiring it in.
 *
 * Run: node tests/agenda-selector.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
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

// =========================================================== normalization

console.log('normalizeAgendaSelector_ -- valid selectors');

check('{mode:"all"} normalizes to itself', sandbox.normalizeAgendaSelector_({ mode: 'all' }), { mode: 'all' });
check('{mode:"all", value: "ignored"} still normalizes to plain {mode:"all"} (value is meaningless for this mode)',
  sandbox.normalizeAgendaSelector_({ mode: 'all', value: 'ignored' }), { mode: 'all' });

check('{mode:"prefix", value:"7."} normalizes to itself',
  sandbox.normalizeAgendaSelector_({ mode: 'prefix', value: '7.' }), { mode: 'prefix', value: '7.' });
check('{mode:"prefix", value:"  11.  "} trims the value',
  sandbox.normalizeAgendaSelector_({ mode: 'prefix', value: '  11.  ' }), { mode: 'prefix', value: '11.' });

check('{mode:"itemList", value:["1.5","1.6"]} normalizes to itself',
  sandbox.normalizeAgendaSelector_({ mode: 'itemList', value: ['1.5', '1.6'] }), { mode: 'itemList', value: ['1.5', '1.6'] });
check('{mode:"itemList", value:[" 1.5 ", "1.6"]} trims each entry',
  sandbox.normalizeAgendaSelector_({ mode: 'itemList', value: [' 1.5 ', '1.6'] }), { mode: 'itemList', value: ['1.5', '1.6'] });
check('{mode:"itemList", value:["1.5", "", "  ", "1.6"]} drops blank entries',
  sandbox.normalizeAgendaSelector_({ mode: 'itemList', value: ['1.5', '', '  ', '1.6'] }), { mode: 'itemList', value: ['1.5', '1.6'] });

console.log('normalizeAgendaSelector_ -- invalid selectors reject clearly, never silently become {mode:"all"}');

expectThrows('{mode:"bogus"} throws', () => sandbox.normalizeAgendaSelector_({ mode: 'bogus' }),
  'unsupported mode "bogus"');
expectThrows('{mode:"prefix"} (no value) throws', () => sandbox.normalizeAgendaSelector_({ mode: 'prefix' }),
  'mode "prefix" requires a non-empty string');
expectThrows('{mode:"prefix", value:""} throws', () => sandbox.normalizeAgendaSelector_({ mode: 'prefix', value: '' }),
  'mode "prefix" requires a non-empty string');
expectThrows('{mode:"prefix", value:"   "} (whitespace only) throws', () => sandbox.normalizeAgendaSelector_({ mode: 'prefix', value: '   ' }),
  'mode "prefix" requires a non-empty string');
expectThrows('{mode:"prefix", value:5} (non-string) throws', () => sandbox.normalizeAgendaSelector_({ mode: 'prefix', value: 5 }),
  'mode "prefix" requires a non-empty string');
expectThrows('{mode:"itemList"} (no value) throws', () => sandbox.normalizeAgendaSelector_({ mode: 'itemList' }),
  'mode "itemList" requires an array');
expectThrows('{mode:"itemList", value:[]} (empty array) throws', () => sandbox.normalizeAgendaSelector_({ mode: 'itemList', value: [] }),
  'at least one non-empty item');
expectThrows('{mode:"itemList", value:["", "  "]} (all-blank array) throws', () => sandbox.normalizeAgendaSelector_({ mode: 'itemList', value: ['', '  '] }),
  'at least one non-empty item');
expectThrows('{mode:"itemList", value:"1.5"} (non-array value) throws', () => sandbox.normalizeAgendaSelector_({ mode: 'itemList', value: '1.5' }),
  'mode "itemList" requires an array');
expectThrows('null selector throws', () => sandbox.normalizeAgendaSelector_(null), 'expected an object');
expectThrows('undefined selector throws', () => sandbox.normalizeAgendaSelector_(undefined), 'expected an object');
expectThrows('string selector throws', () => sandbox.normalizeAgendaSelector_('all'), 'expected an object');
expectThrows('{} (no mode at all) throws', () => sandbox.normalizeAgendaSelector_({}), 'unsupported mode');

// ======================================================= mode: 'prefix' ===
//
// These values reproduce EXACTLY the current production report-type ->
// agenda-prefix mapping (getAgendaPrefixForReportType_, Code.js) -- not
// invented. Every one of these strings is what a real 7 report types
// resolve to today.

console.log('agendaSelectorMatches_ -- mode "prefix", Audio "7." (real production value)');

{
  const sel = { mode: 'prefix', value: '7.' };
  check('"7." matches', sandbox.agendaSelectorMatches_(sel, '7.'), true);
  check('"7.1" matches', sandbox.agendaSelectorMatches_(sel, '7.1'), true);
  check('"7.1.2" matches', sandbox.agendaSelectorMatches_(sel, '7.1.2'), true);
  check('"7.10" matches', sandbox.agendaSelectorMatches_(sel, '7.10'), true);
  check('"17." does NOT match (the trailing dot in "7." already prevents this false positive)',
    sandbox.agendaSelectorMatches_(sel, '17.'), false);
  check('"70." does NOT match', sandbox.agendaSelectorMatches_(sel, '70.'), false);
}

console.log('agendaSelectorMatches_ -- mode "prefix", 6G "11." (real production value)');

{
  const sel = { mode: 'prefix', value: '11.' };
  check('"11." matches', sandbox.agendaSelectorMatches_(sel, '11.'), true);
  check('"11.1" matches', sandbox.agendaSelectorMatches_(sel, '11.1'), true);
  check('"11.2.3" matches', sandbox.agendaSelectorMatches_(sel, '11.2.3'), true);
  check('"1.1" does NOT match', sandbox.agendaSelectorMatches_(sel, '1.1'), false);
  check('"111." does NOT match', sandbox.agendaSelectorMatches_(sel, '111.'), false);
}

console.log('agendaSelectorMatches_ -- mode "prefix", whitespace/coercion safety');

{
  const sel = { mode: 'prefix', value: '7.' };
  check('leading/trailing whitespace on the agenda item is trimmed before matching',
    sandbox.agendaSelectorMatches_(sel, '  7.1  '), true);
  check('null agenda item does not match and does not throw', sandbox.agendaSelectorMatches_(sel, null), false);
  check('undefined agenda item does not match and does not throw', sandbox.agendaSelectorMatches_(sel, undefined), false);
  check('numeric agenda item is safely stringified', sandbox.agendaSelectorMatches_({ mode: 'prefix', value: '7' }, 7), true);
}

// ========================================================== mode: 'all' ==

console.log('agendaSelectorMatches_ -- mode "all", ULBC-MED-style agenda items');

{
  const sel = { mode: 'all' };
  ['1', '1.1', '4', '4.7', '6'].forEach(item => {
    check(`"${item}" matches (a real agenda item value)`, sandbox.agendaSelectorMatches_(sel, item), true);
  });

  check('null does NOT match "all" (not a real agenda item)', sandbox.agendaSelectorMatches_(sel, null), false);
  check('undefined does NOT match "all"', sandbox.agendaSelectorMatches_(sel, undefined), false);
  check('empty string does NOT match "all"', sandbox.agendaSelectorMatches_(sel, ''), false);
  check('whitespace-only string does NOT match "all"', sandbox.agendaSelectorMatches_(sel, '   '), false);
}

// ====================================================== mode: 'itemList' =
//
// Evidence-backed recurring-telco example (SA4-ARCH-006): a real Audio SWG
// ad-hoc report explicitly states it is "keeping only relevant items from
// the unique agenda for 3GPP SA4 AH telcos post-136" -- DaCAS=1.5,
// ATIAS_Ph3-MED=1.6 were both real, verified slot numbers.

console.log('agendaSelectorMatches_ -- mode "itemList", evidence-backed recurring-telco example');

{
  const sel = { mode: 'itemList', value: ['1.5', '1.6'] };
  check('"1.5" matches', sandbox.agendaSelectorMatches_(sel, '1.5'), true);
  check('"1.6" matches', sandbox.agendaSelectorMatches_(sel, '1.6'), true);
  check('"1.4" does NOT match (not in the list)', sandbox.agendaSelectorMatches_(sel, '1.4'), false);
  check('"1.7" does NOT match', sandbox.agendaSelectorMatches_(sel, '1.7'), false);
  check('"1.50" does NOT match (exact match only, not a prefix match)', sandbox.agendaSelectorMatches_(sel, '1.50'), false);
  check('"1.5.1" does NOT match (no subtree/hierarchical semantics)', sandbox.agendaSelectorMatches_(sel, '1.5.1'), false);
  check('"11.5" does NOT match', sandbox.agendaSelectorMatches_(sel, '11.5'), false);
  check('whitespace-padded agenda item still matches ("1.5" vs " 1.5 ")', sandbox.agendaSelectorMatches_(sel, ' 1.5 '), true);
}

{
  // Whitespace normalization on the SELECTOR side too.
  const sel = { mode: 'itemList', value: [' 1.5 ', '1.6'] };
  check('a whitespace-padded selector value still matches the clean agenda item',
    sandbox.agendaSelectorMatches_(sel, '1.5'), true);
}

// -------------------------------------------------------------------------

console.log(failures === 0 ? '\nAll agenda-selector tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
