/**
 * PERF-006B — accurate abstracts diagnostics.
 *
 * PERF-006's live production validation exposed a diagnostics bug: after
 * a Reviewer negative-cache HIT (zero Reviewer API requests), the log
 * still read "Fetched abstracts for 1 table(s)" -- wrong, because
 * addAbstractsForTables_()'s return value was always "candidate tables
 * processed" (this function's own loop-iteration count), never "Reviewer
 * requests actually made". A candidate can resolve via a negative-cache
 * skip without ever calling the Reviewer API (see PERF-006's own
 * isReviewerNoSummaryCached_() cache), or via a genuine fetch that
 * inserts an abstract row, or via a genuine fetch that finds nothing to
 * insert -- three different outcomes the old single count conflated.
 *
 * addAbstractsForTables_() now returns {candidatesProcessed, requestsMade,
 * cacheSkips, rowsInserted}, computed as a delta of the SAME global PERF
 * counters fetchAndAddAbstract_() already increments (no new counters, no
 * Reviewer-cache/abstract-fetch behavior change -- diagnostics only).
 *
 * Run: node tests/perf006b-abstracts-diagnostics.test.js
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

// Same minimal-surface fake DocumentApp Table pattern already proven
// against the real addAbstractsForTables_()/isTDocTable_() (which reads
// table.getCell(0,0) directly, not table.getRow(0).getCell(0)) in
// tests/reviewer-api-tdoc-family.test.js.
function makeFakeTableCell(text) {
  return { getText: () => text, editAsText: () => ({ setFontSize: () => {} }) };
}
function makeFakeTableRow(cells) {
  const _cells = cells.slice();
  return {
    getNumCells: () => _cells.length,
    getCell: (i) => makeFakeTableCell(_cells[i] === undefined ? '' : _cells[i]),
    appendTableCell: (t) => { _cells.push(t); return makeFakeTableCell(t); }
  };
}
function makeFakeTable(tdocId, hasAbstract) {
  const rowsData = [['TDoc', tdocId], ['Agenda Item', '7.1']];
  if (hasAbstract) rowsData.push(['Abstract', 'already here']);
  const _rows = rowsData.map(r => makeFakeTableRow(r));
  return {
    getType: () => 'TABLE',
    getNumRows: () => _rows.length,
    getRow: (i) => _rows[i],
    getCell: (r, c) => _rows[r].getCell(c),
    insertTableRow: (idx) => { const r = makeFakeTableRow([]); _rows.splice(idx, 0, r); return r; }
  };
}

function makeFakeBody(tables) {
  return { getTables: () => tables };
}

// ==== 1. the exact production scenario: fresh negative-cache hit =========

console.log('production scenario -- one candidate table, a FRESH negative-cache entry: zero Reviewer requests, one cache skip, zero rows inserted');

{
  const { sandbox, docProps } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'fake-token' } });
  sandbox.Logger = { log: () => {} };
  docProps.setProperty('REVIEWER_NO_SUMMARY_CACHE_S4aP260074', JSON.stringify({ ts: Date.now(), statusCode: 404 }));

  let fetchCalls = 0;
  sandbox.UrlFetchApp = { fetch: () => { fetchCalls++; return { getResponseCode: () => 200, getContentText: () => '{"text":"should not be reached"}' }; } };

  const body = makeFakeBody([makeFakeTable('S4aP260074', false)]);
  const result = sandbox.addAbstractsForTables_(body);

  check('candidatesProcessed is 1 (the old, misleading "count")', result.candidatesProcessed, 1);
  check('requestsMade is 0 -- NOT "Fetched abstracts for 1 table(s)"', result.requestsMade, 0);
  check('cacheSkips is 1', result.cacheSkips, 1);
  check('rowsInserted is 0', result.rowsInserted, 0);
  check('zero actual Reviewer fetches occurred', fetchCalls, 0);
}

// ==== 2. no candidates at all: everything zero, no false positive =========

console.log('no candidate tables at all (every table already has an Abstract cell) -- everything reports zero');

{
  const { sandbox } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'fake-token' } });
  sandbox.Logger = { log: () => {} };
  sandbox.UrlFetchApp = { fetch: () => { throw new Error('should never be called'); } };

  const body = makeFakeBody([makeFakeTable('S4-260100', true)]);
  const result = sandbox.addAbstractsForTables_(body);

  check('candidatesProcessed is 0', result.candidatesProcessed, 0);
  check('requestsMade is 0', result.requestsMade, 0);
  check('cacheSkips is 0', result.cacheSkips, 0);
  check('rowsInserted is 0', result.rowsInserted, 0);
}

// ==== 3. a real fetch that successfully inserts a row =======================

console.log('a genuine cache-miss fetch that succeeds -- one request made, one row inserted, zero cache skips');

{
  const { sandbox } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'fake-token' } });
  sandbox.Logger = { log: () => {} };
  sandbox.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ text: 'A real summary.' }) }) };

  const body = makeFakeBody([makeFakeTable('S4-260100', false)]);
  const result = sandbox.addAbstractsForTables_(body);

  check('candidatesProcessed is 1', result.candidatesProcessed, 1);
  check('requestsMade is 1', result.requestsMade, 1);
  check('cacheSkips is 0', result.cacheSkips, 0);
  check('rowsInserted is 1', result.rowsInserted, 1);
}

// ==== 4. a real fetch that gets a fresh 404 (no prior cache entry) =========

console.log('a genuine cache-miss fetch that 404s -- one request made, zero cache skips (this IS the write, not a hit), zero rows inserted');

{
  const { sandbox } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'fake-token' } });
  sandbox.Logger = { log: () => {} };
  sandbox.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 404, getContentText: () => '' }) };

  const body = makeFakeBody([makeFakeTable('S4aP260074', false)]);
  const result = sandbox.addAbstractsForTables_(body);

  check('candidatesProcessed is 1', result.candidatesProcessed, 1);
  check('requestsMade is 1 (the fetch DID happen -- the cache did not exist yet)', result.requestsMade, 1);
  check('cacheSkips is 0 (a cache WRITE is not a cache HIT)', result.cacheSkips, 0);
  check('rowsInserted is 0', result.rowsInserted, 0);
}

// ==== 5. an EARLIER-in-the-same-run fetchAndAddAbstract_() call (from a ====
// ==== different call site) is correctly excluded from this function's =====
// ==== own delta ============================================================

console.log('a fetchAndAddAbstract_() call from a DIFFERENT call site earlier in the same run does not pollute this function\'s own delta');

{
  const { sandbox } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'fake-token' } });
  sandbox.Logger = { log: () => {} };
  sandbox.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ text: 'from an earlier call site' }) }) };

  // Simulate insertNewTdoc_() -> createTDocTableFromData_() having already
  // called fetchAndAddAbstract_() earlier in the SAME run, before
  // addAbstractsForTables_() runs (exactly continuousUpdateCore_()'s real
  // ordering: new-TDoc insertion happens before the abstracts sweep).
  sandbox.fetchAndAddAbstract_(makeFakeTable('S4-260099', false), 'S4-260099');
  check('sanity: the simulated earlier call really did increment the global counters',
    sandbox.PERF_COUNTERS_['Reviewer API requests'], 1);

  const body = makeFakeBody([makeFakeTable('S4-260100', false)]);
  const result = sandbox.addAbstractsForTables_(body);

  check('this function\'s OWN delta only counts ITS OWN candidate (1), not the earlier call site\'s',
    result.requestsMade, 1);
  check('the global cumulative counter now reflects BOTH calls (2), proving the delta -- not the raw counter -- was used',
    sandbox.PERF_COUNTERS_['Reviewer API requests'], 2);
}

// ==== 6. source-structure: both callers log an accurate, non-misleading ===
// ==== message, and the old wrong wording is gone ===========================

console.log('source-structure: the old misleading "Fetched abstracts for N table(s)" wording is gone; both callers use the breakdown');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  check('the old misleading Logger.log template ("Fetched abstracts for ${...} table(s)") no longer exists as active code (only in explanatory comments)',
    /Logger\.log\(`Fetched abstracts for/.test(source), false);
  check('continuousUpdateCore_() logs requestsMade explicitly',
    /abstractsResult\.requestsMade/.test(source), true);
  check('continuousUpdateCore_() logs cacheSkips explicitly',
    /abstractsResult\.cacheSkips/.test(source), true);
  check('continuousUpdateCore_() logs rowsInserted explicitly',
    /abstractsResult\.rowsInserted/.test(source), true);
  check('addAbstractsOnly() reads result.candidatesProcessed (not a bare count)',
    /result\.candidatesProcessed/.test(source), true);
  check('addAbstractsForTables_() does not change FETCH_ABSTRACTS_ON_UPDATE or any Reviewer-cache write/skip semantics (no new perfCount_ label introduced)',
    !/function addAbstractsForTables_[\s\S]*?perfCount_\(/.test(source.slice(source.indexOf('function addAbstractsForTables_'), source.indexOf('function addAbstractsOnly'))),
    true);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
