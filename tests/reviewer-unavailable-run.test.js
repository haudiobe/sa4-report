/**
 * A Reviewer that fails is asked once per execution, not once per TDoc (after T-2026.10.24).
 *
 * A request that times out blocks for up to about a minute in Apps Script, and nothing is
 * remembered for it. Before this change every TDoc table without an Abstract was asked in
 * turn, so a hanging Reviewer could use up the six minutes of an update before its e-mail,
 * revision and formatting steps. Now:
 *
 *   1. a timeout or network error, a 401/403 and a 5xx stop further Reviewer requests in
 *      the same execution; the update goes on and finishes without abstracts;
 *   2. it is said once in the log, and nothing is written to the "no summary" cache;
 *   3. the next execution asks again, and gets the abstracts once the Reviewer answers;
 *   4. a 404 and a summary are answers about the document: they stop nothing;
 *   5. "Update Abstracts" is protected the same way.
 *
 * The REAL update and "Update Abstracts" run on the fake document
 * (tests/helpers/template-report.js). All values are synthetic.
 *
 * Run: node tests/reviewer-unavailable-run.test.js
 */

const { templateReport, CACHE } = require('./helpers/template-report.js');

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

const A = 'S4aA269101', B = 'S4aA269102', C = 'S4aA269103';
const IDS = [A, B, C];
const TIMEOUT = { throws: 'Timeout: https://reviewer.bouazizi.dev/api/v1/documents/x/summary?type=summary' };
const t = (id) => ({ id: id, agenda: '4.3', status: 'available', revisedTo: '', uploaded: true });

/** A built report whose TDocs are uploaded and have no Abstract yet; automatic abstracts on. */
function report(reviewer) {
  const r = templateReport({ tdocs: IDS.map(t), reviewer: {}, props: { FETCH_ABSTRACTS_ON_UPDATE: 'true' } });
  const built = r.build();
  if (!built.ok) throw new Error('fixture build failed: ' + built.message);
  IDS.forEach((id) => r.docProps.deleteProperty(CACHE(id)));
  r.reviewer = Object.assign(r.reviewer, reviewer);
  r.collected = 0;
  r.s.collectorUpdate_ = () => { r.collected++; return { failures: [] }; };
  r.mark();
  return r;
}
const cacheEntries = (r) => IDS.filter((id) => r.docProps.getProperty(CACHE(id)));
const abstracts = (r) => IDS.map((id) => r.field(id, 'Abstract') || null);
const unavailableLines = (r) => r.logs.filter((l) => /^The Reviewer is not available \(/.test(l));
const allFail = (answer) => ({ [A]: answer, [B]: answer, [C]: answer });

// ====================================================================== 1
console.log('1. an update: a failing Reviewer is asked once, and the update finishes');
[
  ['a timeout', TIMEOUT, 'Timeout: https://reviewer.bouazizi.dev/api/v1/documents/x/summary?type=summary'],
  ['a network error', { throws: 'Address unavailable: https://reviewer.bouazizi.dev' }, 'Address unavailable: https://reviewer.bouazizi.dev'],
  ['HTTP 503', 503, 'HTTP 503'],
  ['HTTP 500', 500, 'HTTP 500'],
  ['HTTP 401', 401, 'HTTP 401'],
  ['HTTP 403', 403, 'HTTP 403'],
].forEach(([label, answer, reason]) => {
  const r = report(allFail(answer));
  const result = r.update();
  check(`${label}: one Reviewer request in the run, not three`, r.requests.length, 1);
  check(`${label}: the update succeeds and reaches its later steps (e-mail and revisions)`, [result.success, r.collected], [true, 1]);
  check(`${label}: said once, with the reason`, unavailableLines(r), ['The Reviewer is not available (' + reason + '): it is not asked again in this run. The report is updated without further abstracts.']);
  check(`${label}: no "no summary" entry, no abstract`, [cacheEntries(r), abstracts(r)], [[], [null, null, null]]);
});

// ====================================================================== 2
console.log('2. the next execution asks again');
{
  const r = report(allFail(TIMEOUT));
  r.update();
  r.mark();
  r.update();
  check('a second update while the Reviewer still fails: again exactly one request', r.requests.length, 1);
  IDS.forEach((id) => { r.reviewer[id] = 'The synthetic summary of ' + id + '.'; });
  r.mark();
  r.update();
  check('once the Reviewer answers: every table is asked and gets its abstract', [r.requests, abstracts(r)], [IDS, IDS.map((id) => 'The synthetic summary of ' + id + '.')]);
  check('...and nothing is said about an unavailable Reviewer', unavailableLines(r), []);
}

// ====================================================================== 3
console.log('3. answers about a document stop nothing');
{
  const r = report({ [A]: 404, [B]: 'The synthetic summary of ' + B + '.', [C]: 404 });
  r.update();
  check('404, summary, 404: all three asked', r.requests, IDS);
  check('the summary is inserted; both 404s are remembered as before', [abstracts(r), cacheEntries(r)], [[null, 'The synthetic summary of ' + B + '.', null], [A, C]]);
  check('no unavailable line', unavailableLines(r), []);
}
{
  const r = report({ [A]: 404, [B]: 503, [C]: 'The synthetic summary of ' + C + '.' });
  r.update();
  check('404 then 503: the 404 is answered, the 503 stops the run before the third', [r.requests, cacheEntries(r), abstracts(r)], [[A, B], [A], [null, null, null]]);
}

// ====================================================================== 4
console.log('4. "Update Abstracts"');
{
  const r = report(allFail(TIMEOUT));
  r.updateAbstracts();
  check('one request, and its result is shown', [r.requests.length, r.ui.alerts.length > 0], [1, true]);
  IDS.forEach((id) => { r.reviewer[id] = 'The synthetic summary of ' + id + '.'; });
  r.mark();
  r.updateAbstracts();
  check('the next "Update Abstracts" asks every table again and inserts the abstracts', [r.requests, abstracts(r).every(Boolean)], [IDS, true]);
}

// ====================================================================== 5
console.log('5. the switch is per execution and in one place');
{
  const r = report({});
  r.s.markReviewerUnavailableForRun_('HTTP 502');
  check('set within an execution', r.s.REVIEWER_TOKEN_RUN_STATE_ ? !!r.s.REVIEWER_TOKEN_RUN_STATE_.unavailable : 'n/a', true);
  r.s.resetReviewerTokenRunState_();
  check('reset at the start of the next one', r.s.REVIEWER_TOKEN_RUN_STATE_ ? r.s.REVIEWER_TOKEN_RUN_STATE_.unavailable : 'n/a', '');
}

// ====================================================================== 6
console.log('6. what the change is in Code.js, and nothing else (source)');
{
  const fs = require('fs');
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const { codeWithoutReviewerRunBreaker, REVIEWER_RUN_BREAKER_NEW_FUNCTIONS, REVIEWER_RUN_BREAKER_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');
  const now = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutReviewerRunBreaker(now);
  check('the take-out finds every change exactly once', view.problems, []);
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (code, name) => { const at = code.indexOf('\nfunction ' + name + '('); return code.slice(at, code.indexOf('\n}\n', at)); };
  const before = names(view.code);
  check('the new functions', names(now).filter((n) => before.indexOf(n) === -1), REVIEWER_RUN_BREAKER_NEW_FUNCTIONS);
  check('the changed functions', before.filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)), REVIEWER_RUN_BREAKER_CHANGED_FUNCTIONS);
  const fetchSrc = bodyOf(now, 'fetchAndAddAbstract_');
  check('the withdrawn rule still comes first; the switch after the token and before the cache and the request',
    [fetchSrc.indexOf('abstractFetchBlockedBy_(') < fetchSrc.indexOf('REVIEWER_TOKEN_RUN_STATE_.unavailable'),
      fetchSrc.indexOf('getReviewerApiTokenForRun_(') < fetchSrc.indexOf('REVIEWER_TOKEN_RUN_STATE_.unavailable'),
      fetchSrc.indexOf('REVIEWER_TOKEN_RUN_STATE_.unavailable') < fetchSrc.indexOf('isReviewerNoSummaryCached_('),
      fetchSrc.indexOf('REVIEWER_TOKEN_RUN_STATE_.unavailable') < fetchSrc.indexOf('UrlFetchApp.fetch(')], [true, true, true, true]);
  check('a 404 still writes the "no summary" cache and does not trip the switch', /statusCode === 404\) \{[\s\S]*?markReviewerNoSummary_[\s\S]*?\} else \{\n      Logger\.log\(`API error/.test(fetchSrc) && !/statusCode === 404\) \{[^}]*markReviewerUnavailableForRun_/.test(fetchSrc), true);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll Reviewer-unavailable checks passed.');
process.exit(failures ? 1 : 0);
