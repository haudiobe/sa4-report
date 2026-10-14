/**
 * The Reviewer "no summary" cache after T-2026.10.13.
 *
 * A 404 from the Reviewer ("no summary for this document") is remembered so
 * that an update does not ask again on every run (PERF-006,
 * tests/perf006-reviewer-negative-cache.test.js). It was remembered for 24
 * hours. The Reviewer often has no summary yet when a TDoc has just been
 * uploaded, so its abstract came most of a meeting day late, and "Update
 * Abstracts" waited too. Now:
 *
 *   1. an update waits 4 hours after a 404, then asks again;
 *   2. "Update Abstracts" asks anyway -- and an abstract it gets is inserted,
 *      the entry removed;
 *   3. a table that has an Abstract is not asked for, by either;
 *   4. other answers (5xx, 401, an empty 200) are still not remembered;
 *   5. the key, its value and its place are as before; an entry written by the
 *      earlier release is read with the new 4 hours.
 *   6. what the change is in Code.js, and nothing else.
 *
 * The REAL update and the REAL "Update Abstracts" run on the fake document
 * (tests/helpers/template-report.js), with a fake Reviewer that records every
 * request. All values are synthetic.
 *
 * Run: node tests/abstract-cache-retry.test.js
 */

const fs = require('fs');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport, CACHE } = require('./helpers/template-report.js');
const { codeWithoutAbstractCacheChange, ABSTRACT_CACHE_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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

const A = 'S4aA269001', B = 'S4aA269002', C = 'S4aA269003';
const HOUR = 3600 * 1000;
const t = (id) => ({ id: id, agenda: '4.3', status: 'available', revisedTo: '', uploaded: true });
/** A built report whose TDocs are uploaded and have no Abstract yet; automatic abstracts on. */
function report(ids, reviewer) {
  const r = templateReport({ tdocs: ids.map(t), reviewer: Object.assign({}, reviewer || {}), props: { FETCH_ABSTRACTS_ON_UPDATE: 'true' } });
  const built = r.build();
  if (!built.ok) throw new Error('fixture build failed: ' + built.message);
  ids.forEach((id) => r.docProps.deleteProperty(CACHE(id)));
  r.mark();
  return r;
}
const entry = (r, id) => { const raw = r.docProps.getProperty(CACHE(id)); return raw ? JSON.parse(raw) : null; };
const setEntry = (r, id, hoursAgo) => r.docProps.setProperty(CACHE(id), JSON.stringify({ ts: Date.now() - hoursAgo * HOUR, statusCode: 404 }));
const asked = (r) => r.requests.slice();
const abstractOf = (r, id) => r.field(id, 'Abstract');
const run = (r, fn) => { r.mark(); fn(); return { requests: asked(r), logs: r.logs.slice() }; };

// ====================================================================== 1
console.log('1. an update: a 404 is remembered for 4 hours');
{
  const r = report([A], { [A]: 404 });
  const first = run(r, () => r.update());
  check('the first update asks the Reviewer and gets 404: an entry is written', [first.requests, entry(r, A) && entry(r, A).statusCode, abstractOf(r, A)], [[A], 404, undefined]);
  const second = run(r, () => r.update());
  check('the next update does not ask, and says why', [second.requests, second.logs.some((l) => l === 'Skipping Reviewer API for ' + A + ': no summary was available on a previous run (cached, within TTL)')], [[], true]);
  check('...and counts it as a negative-cache skip', second.logs.some((l) => /^Abstracts: 1 candidate table\(s\) processed, 0 Reviewer request\(s\) made, 1 negative-cache skip\(s\), 0 abstract row\(s\) inserted$/.test(l)), true);
  setEntry(r, A, 3.9);
  check('3.9 hours after the 404: still not asked', run(r, () => r.update()).requests, []);
  setEntry(r, A, 4.1);
  r.reviewer[A] = 'The synthetic summary of ' + A + '.';
  const late = run(r, () => r.update());
  check('4.1 hours after the 404: asked again; the summary that is there now is inserted, the entry removed', [late.requests, abstractOf(r, A), entry(r, A)], [[A], 'The synthetic summary of ' + A + '.', null]);
  check('from then on the Abstract row is the answer: no further request', run(r, () => r.update()).requests, []);
}

// ====================================================================== 2
console.log('2. "Update Abstracts" asks past the cache');
{
  const r = report([A, B], { [A]: 404, [B]: 404 });
  setEntry(r, A, 0.1);
  setEntry(r, B, 0.1);
  r.reviewer[A] = 'The synthetic summary of ' + A + '.';
  // Other collection caches, as the collectors would have left them: Update Abstracts must not touch them.
  r.docProps.setProperty('DISCUSS_' + A, '["synthetic"]');
  r.docProps.setProperty('REVIS_' + A, '["synthetic"]');
  r.docProps.setProperty('A1_EMPTY_CACHE_synthetic', JSON.stringify({ ts: Date.now(), empty: true }));
  const otherCaches = () => r.docProps.getKeys().filter((k) => /^(DISCUSS_|REVIS_|A1_EMPTY_CACHE_)/.test(k)).sort().map((k) => k + '=' + r.docProps.getProperty(k));
  const othersBefore = otherCaches();
  const manual = run(r, () => r.updateAbstracts());
  check('both TDocs with a live entry are asked', manual.requests.slice().sort(), [A, B]);
  check('the other collection caches are exactly as they were', otherCaches(), othersBefore);
  check('the abstract that is there now is inserted, and its entry removed', [abstractOf(r, A), entry(r, A)], ['The synthetic summary of ' + A + '.', null]);
  check('the one that still has none is written again, with a fresh time: the updates wait 4 hours from now', [entry(r, B).statusCode, Date.now() - entry(r, B).ts < 60 * 1000, abstractOf(r, B)], [404, true, undefined]);
  check('the log says, per TDoc, that the cache is passed because this is a manual Update Abstracts',
    [A, B].map((id) => manual.logs.some((l) => l === 'Asking the Reviewer API for ' + id + ' although it had no summary on a previous run (cached, within TTL): bypassing the negative cache for a manual Update Abstracts')), [true, true]);
  check('no "Skipping" line in a manual run', manual.logs.some((l) => /^Skipping Reviewer API/.test(l)), false);
  check('the user is told how many abstracts were inserted', r.ui.alerts.slice(-1)[0].slice(0, 2), ['Success', 'Abstract step completed: 2 candidate table(s) processed, 1 abstract(s) inserted.']);
  // Only the explicit action passes the cache: the update right after it waits.
  const update = run(r, () => r.update());
  check('the update right after it does not ask for the one with a fresh entry', [update.requests, update.logs.some((l) => l === 'Skipping Reviewer API for ' + B + ': no summary was available on a previous run (cached, within TTL)')], [[], true]);
  check('no Reviewer token in any log line', r.logs.concat(manual.logs, update.logs).some((l) => /synthetic-test-token/.test(l)), false);
}

// ====================================================================== 3
console.log('3. a table that has an Abstract is not asked for');
{
  const r = report([A, C], { [A]: 'The synthetic summary of ' + A + '.', [C]: 404 });
  r.update();
  check('A got its abstract on the update', abstractOf(r, A), 'The synthetic summary of ' + A + '.');
  setEntry(r, C, 0.1);
  const manual = run(r, () => r.updateAbstracts());
  check('Update Abstracts asks for C (no Abstract, live entry), never for A (has one)', manual.requests, [C]);
  check('A\'s Abstract is unchanged', abstractOf(r, A), 'The synthetic summary of ' + A + '.');
}

// ====================================================================== 4
console.log('4. other answers are still not remembered');
{
  [[500, 'a server error'], [401, 'an authorization error'], [503, 'a transient error']].forEach(([code, label]) => {
    const r = report([A], { [A]: code });
    r.update();
    check(label + ' (' + code + ') on an update: no entry', entry(r, A), null);
    r.updateAbstracts();
    check(label + ' (' + code + ') on Update Abstracts: no entry either', entry(r, A), null);
  });
  const empty = report([A], { [A]: '' });
  empty.update();
  check('a 200 with an empty summary: no entry, no Abstract', [entry(empty, A), abstractOf(empty, A)], [null, undefined]);
}

// ====================================================================== 5
console.log('5. key, value and place as before; an entry written by the earlier release');
{
  const r = report([A], { [A]: 404 });
  r.update();
  check('the key is REVIEWER_NO_SUMMARY_CACHE_ + the canonical TDoc number, in the report\'s Document Properties', [r.s.reviewerNoSummaryCacheKey_(A), r.docProps.getKeys().filter((k) => /^REVIEWER_NO_SUMMARY_CACHE_/.test(k))], ['REVIEWER_NO_SUMMARY_CACHE_' + A, ['REVIEWER_NO_SUMMARY_CACHE_' + A]]);
  check('the value is { ts, statusCode } as before', Object.keys(entry(r, A)).sort(), ['statusCode', 'ts']);
  check('one entry per TDoc and report: another report does not see it', report([A], { [A]: 404 }).docProps.getProperty(CACHE(A)), null);
  check('the TTL is 4 hours', r.s.REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS, 4);
  // An entry of T-2026.10.13 or earlier has the same shape; its time is read with the new TTL, nothing is migrated.
  setEntry(r, A, 5);
  check('an entry written 5 hours ago by the earlier release (24 hours then) is expired now: the update asks', run(r, () => r.update()).requests, [A]);
  setEntry(r, A, 2);
  check('one written 2 hours ago still holds', run(r, () => r.update()).requests, []);
}

// ====================================================================== 6
console.log('6. what the change is');
{
  const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const before = codeWithoutAbstractCacheChange(CODE);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, '')).filter((n, i, list) => list.indexOf(n) === i);
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  check('every part of the change is in Code.js exactly once', before.problems, []);
  check('three functions changed, none added or removed', [functionNames(CODE).filter((n) => functionSource(CODE, n) !== functionSource(before.code, n)), functionNames(CODE).length === functionNames(before.code).length],
    [ABSTRACT_CACHE_CHANGED_FUNCTIONS, true]);
  check('besides, only the TTL line and its comment', CODE.split('\n').filter((l) => /REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS =/.test(l)), ['var REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS = 4;']);
  check('the cache functions themselves are unchanged', ['reviewerNoSummaryCacheKey_', 'isReviewerNoSummaryCached_', 'markReviewerNoSummary_', 'clearReviewerNoSummaryCache_', 'refreshUploadedTdocMetadata_', 'completeInsertedUploadedTdoc_', 'clearAllCaches']
    .filter((n) => functionSource(CODE, n) !== functionSource(before.code, n)), []);
  check('only Update Abstracts passes bypassNegativeCache', CODE.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '').split('\nfunction ').filter((f) => /bypassNegativeCache: true/.test(f)).map((f) => f.slice(0, f.indexOf('('))), ['addAbstractsOnly']);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll abstract-cache checks passed.');
process.exitCode = failures ? 1 : 0;
