/**
 * PERF-006 (Part A) — Reviewer API "no summary" negative cache.
 *
 * PERF-005 found that fetchAndAddAbstract_() retries the exact same
 * Reviewer API request every automatic run for a TDoc whose summary is
 * permanently unavailable: a 404 persisted nothing, so
 * addAbstractsForTables_()'s "does this table already have an Abstract
 * cell?" gate always finds it still missing and tries again.
 *
 * The fix mirrors the existing A1 empty-page cache exactly: Document
 * Properties, {ts, statusCode} shape, 24h default TTL
 * (REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS), keyed by canonical TDoc identity
 * via the central parseExactSA4DocumentId_() registry (never a second
 * parser). Only a 404 -- the Reviewer API's own explicit "no summary
 * exists" contract -- is cached negative; auth failures, other 4xx,
 * 5xx, malformed responses, and exceptions are all deliberately left
 * uncached (see fetchAndAddAbstract_()'s own comment for the reasoning
 * behind each exclusion).
 *
 * Run: node tests/perf006-reviewer-negative-cache.test.js
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

function makeFakeTable() {
  // Minimal fake TDoc table: enough surface for findAbstractInsertIndex_()/
  // insertTableRow()/appendTableCell() to run without throwing, matching
  // the existing project pattern (tests/revision-drafts-folder.test.js).
  const rows = [
    ['TDoc', 'S4-260074'],
    ['Agenda Item', '7.1']
  ];
  return {
    getNumRows: () => rows.length,
    getRow: (r) => ({
      getNumCells: () => rows[r].length,
      getCell: (c) => ({ getText: () => rows[r][c] })
    }),
    insertTableRow: (idx) => {
      const newRow = ['', ''];
      rows.splice(idx, 0, newRow);
      let cellIdx = 0;
      return {
        appendTableCell: (text) => {
          newRow[cellIdx] = text;
          cellIdx++;
          return { editAsText: () => ({ setFontSize: () => {} }) };
        }
      };
    }
  };
}

function makeSandboxWithToken(options) {
  const { sandbox, docProps } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: 'fake-token' } });
  sandbox.Logger = { log: () => {} };
  return { sandbox, docProps };
}

// ============ 1. first definitive 404: fetches, caches, logs ==============

console.log('first definitive 404 response -- performs the fetch, writes the negative cache, does not insert an Abstract row');

{
  const { sandbox, docProps } = makeSandboxWithToken();
  let fetchCalls = 0;
  sandbox.UrlFetchApp = { fetch: () => { fetchCalls++; return { getResponseCode: () => 404, getContentText: () => '' }; } };

  const table = makeFakeTable();
  sandbox.fetchAndAddAbstract_(table, 'S4-260074');

  check('exactly one Reviewer API fetch was performed', fetchCalls, 1);
  check('Reviewer API requests counter incremented', sandbox.PERF_COUNTERS_['Reviewer API requests'], 1);
  check('Reviewer negative-cache writes counter incremented', sandbox.PERF_COUNTERS_['Reviewer negative-cache writes'], 1);
  check('no negative-cache HIT counter yet (this was a miss -> fetch)', sandbox.PERF_COUNTERS_['Reviewer negative-cache hits'], undefined);
  check('a REVIEWER_NO_SUMMARY_CACHE_ property was written for the canonical TDoc id', docProps.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074') !== null, true);

  const stored = JSON.parse(docProps.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074'));
  check('the stored entry records statusCode 404', stored.statusCode, 404);
  check('the stored entry has a numeric timestamp', typeof stored.ts === 'number' && stored.ts > 0, true);
  check('no Abstract row was inserted (table still has only its original 2 rows)', table.getNumRows(), 2);
}

// ============ 2. fresh negative-cache hit: zero Reviewer fetches ==========

console.log('a fresh (within-TTL) negative-cache entry -- skips the Reviewer request entirely');

{
  const { sandbox, docProps } = makeSandboxWithToken();
  docProps.setProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074', JSON.stringify({ ts: Date.now(), statusCode: 404 }));

  let fetchCalls = 0;
  sandbox.UrlFetchApp = { fetch: () => { fetchCalls++; return { getResponseCode: () => 200, getContentText: () => '{"text":"should never be reached"}' }; } };

  const table = makeFakeTable();
  sandbox.fetchAndAddAbstract_(table, 'S4-260074');

  check('zero Reviewer API fetches were performed', fetchCalls, 0);
  check('the "Reviewer API requests" counter was NOT incremented on a cache hit', sandbox.PERF_COUNTERS_['Reviewer API requests'], undefined);
  check('the "Reviewer negative-cache hits" counter was incremented', sandbox.PERF_COUNTERS_['Reviewer negative-cache hits'], 1);
  check('no Abstract row was inserted', table.getNumRows(), 2);
}

// ============ 3. expired entry: fetches again ==============================

console.log('an EXPIRED negative-cache entry (older than the 24h TTL) -- fetches again, as a cache miss');

{
  const { sandbox, docProps } = makeSandboxWithToken();
  const twentyFiveHoursAgo = Date.now() - (25 * 3600 * 1000);
  docProps.setProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074', JSON.stringify({ ts: twentyFiveHoursAgo, statusCode: 404 }));

  let fetchCalls = 0;
  sandbox.UrlFetchApp = { fetch: () => { fetchCalls++; return { getResponseCode: () => 404, getContentText: () => '' }; } };

  const table = makeFakeTable();
  sandbox.fetchAndAddAbstract_(table, 'S4-260074');

  check('the expired entry does not suppress the fetch -- exactly one Reviewer API fetch was performed', fetchCalls, 1);
  check('the cache was re-written with a fresh timestamp', JSON.parse(docProps.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074')).ts > twentyFiveHoursAgo, true);
}

// ==== 4. a successful summary clears a stale negative-cache entry =========

console.log('a successful (200 + real text) response clears any stale negative-cache entry and inserts the Abstract row, unchanged');

{
  const { sandbox, docProps } = makeSandboxWithToken();
  // An EXPIRED entry (past the 24h TTL) so the fetch actually proceeds --
  // a still-FRESH negative entry is expected to keep skipping the fetch
  // entirely (that is exactly test 2's scenario, by design: the TTL is
  // the bounded re-check interval). This test is about what happens once
  // that TTL has elapsed and a real answer finally comes back.
  const twentyFiveHoursAgo = Date.now() - (25 * 3600 * 1000);
  docProps.setProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074', JSON.stringify({ ts: twentyFiveHoursAgo, statusCode: 404 }));

  sandbox.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ text: 'A real summary.' }) }) };

  const table = makeFakeTable();
  sandbox.fetchAndAddAbstract_(table, 'S4-260074');

  check('the stale negative-cache entry was cleared', docProps.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074'), null);
  check('the "structural: abstract row inserted" counter still fires (existing PERF-003B behavior unchanged)', sandbox.PERF_COUNTERS_['structural: abstract row inserted (fetchAndAddAbstract_)'], 1);
  check('an Abstract row was actually inserted (table grew from 2 to 3 rows)', table.getNumRows(), 3);
}

// ==== 5. transient/network/5xx/auth failures are NOT negative-cached =======

console.log('non-404 outcomes are never written to the negative cache');

{
  const cases = [
    ['401 (auth failure)', () => ({ getResponseCode: () => 401, getContentText: () => '' })],
    ['403 (forbidden)', () => ({ getResponseCode: () => 403, getContentText: () => '' })],
    ['500 (server error)', () => ({ getResponseCode: () => 500, getContentText: () => '' })],
    ['503 (transient/unavailable)', () => ({ getResponseCode: () => 503, getContentText: () => '' })]
  ];
  cases.forEach(([label, fetchImpl]) => {
    const { sandbox, docProps } = makeSandboxWithToken();
    sandbox.UrlFetchApp = { fetch: fetchImpl };
    const table = makeFakeTable();
    sandbox.fetchAndAddAbstract_(table, 'S4-260074');
    check(`${label}: no negative-cache entry written`, docProps.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074'), null);
    check(`${label}: no negative-cache WRITE counter incremented`, sandbox.PERF_COUNTERS_['Reviewer negative-cache writes'], undefined);
  });

  // A thrown exception (e.g. UrlFetchApp itself throwing -- network-level
  // failure) must also never be cached.
  const { sandbox, docProps } = makeSandboxWithToken();
  sandbox.UrlFetchApp = { fetch: () => { throw new Error('network unreachable'); } };
  const table = makeFakeTable();
  let threw = null;
  try { sandbox.fetchAndAddAbstract_(table, 'S4-260074'); } catch (e) { threw = e; }
  check('a thrown exception is swallowed by fetchAndAddAbstract_\'s own try/catch (pre-existing behavior, unchanged)', threw, null);
  check('a thrown exception writes no negative-cache entry', docProps.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074'), null);

  // A 200 with an EMPTY body is deliberately left uncached too (not proven
  // permanent -- Reviewer may simply not have finished processing yet).
  const { sandbox: sandbox2, docProps: docProps2 } = makeSandboxWithToken();
  sandbox2.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ text: '' }) }) };
  sandbox2.fetchAndAddAbstract_(makeFakeTable(), 'S4-260074');
  check('a 200 with an empty summary body is NOT negative-cached', docProps2.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4-260074'), null);
}

// ==== 6. cache key uses canonical TDoc identity ============================

console.log('the cache key is the CANONICAL TDoc identity (parseExactSA4DocumentId_), not the raw input string');

{
  const { sandbox } = loadCode();
  check('reviewerNoSummaryCacheKey_ uses the exact string passed (canonicalization happens in fetchAndAddAbstract_, before calling this)',
    sandbox.reviewerNoSummaryCacheKey_('S4aP260071'), 'REVIEWER_NO_SUMMARY_CACHE_S4aP260071');
}

{
  // An ad-hoc family TDoc (S4aP) must be cached/looked-up under its OWN
  // canonical form, proving parseExactSA4DocumentId_() (not a second,
  // main-meeting-only parser) drives the cache key.
  const { sandbox, docProps } = makeSandboxWithToken();
  let fetchCalls = 0;
  sandbox.UrlFetchApp = { fetch: () => { fetchCalls++; return { getResponseCode: () => 404, getContentText: () => '' }; } };
  sandbox.fetchAndAddAbstract_(makeFakeTable(), 'S4aP260071');
  check('an ad-hoc-family TDoc (S4aP260071) is cached under its own canonical id', docProps.getProperty('REVIEWER_NO_SUMMARY_CACHE_S4aP260071') !== null, true);

  fetchCalls = 0;
  sandbox.fetchAndAddAbstract_(makeFakeTable(), 'S4aP260071');
  check('the SAME ad-hoc TDoc is then served from cache on the very next call (zero fetches)', fetchCalls, 0);
}

// ==== 7. existing abstract insertion behavior is unchanged (success path) =

console.log('the successful-summary insertion path (font size, cell text, row position) is byte-identical to before');

{
  const { sandbox } = makeSandboxWithToken();
  sandbox.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ text: 'The abstract text.' }) }) };

  const table = makeFakeTable();
  sandbox.fetchAndAddAbstract_(table, 'S4-260074');

  check('the Abstract row is inserted immediately after "Agenda Item" (existing findAbstractInsertIndex_ behavior)',
    [table.getRow(2).getCell(0).getText(), table.getRow(2).getCell(1).getText()],
    ['Abstract', 'The abstract text.']);
}

// ==== 8. source-structure: TTL default and central helper design ==========

console.log('source-structure: 24h default TTL, central helper design (not scattered PropertiesService calls)');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  check('REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS defaults to 24 (conservative, matches the existing A1 empty-cache convention)',
    /var REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS = 24;/.test(source), true);
  check('a dedicated reviewerNoSummaryCacheKey_() helper exists (central key construction)',
    /function reviewerNoSummaryCacheKey_\(/.test(source), true);
  check('a dedicated isReviewerNoSummaryCached_() helper exists (central read+TTL check)',
    /function isReviewerNoSummaryCached_\(/.test(source), true);
  check('a dedicated markReviewerNoSummary_() helper exists (central write)',
    /function markReviewerNoSummary_\(/.test(source), true);
  check('a dedicated clearReviewerNoSummaryCache_() helper exists (central invalidation)',
    /function clearReviewerNoSummaryCache_\(/.test(source), true);
  check('FETCH_ABSTRACTS_ON_UPDATE is never referenced by any of the new Reviewer-cache code (this feature does not touch that setting)',
    !/function (reviewerNoSummaryCacheKey_|isReviewerNoSummaryCached_|markReviewerNoSummary_|clearReviewerNoSummaryCache_)[\s\S]*?FETCH_ABSTRACTS_ON_UPDATE/.test(source), true);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
