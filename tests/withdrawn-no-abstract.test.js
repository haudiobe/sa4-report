/**
 * No abstract is asked for a withdrawn TDoc (after T-2026.10.8).
 *
 * THE RULE (abstractFetchBlockedBy_()): the abstract of a TDoc is not asked
 * for when EITHER says it is withdrawn --
 *   - the Status of its table in the report (text, or the value selected in
 *     a status dropdown);
 *   - the status the Portal's TDoc list has for it, in an operation that
 *     has the list (an update).
 * It is enforced in fetchAndAddAbstract_(), the one function that asks the
 * Contribution Reviewer for an abstract, before the token, the "no summary"
 * cache and the request -- so every path that asks goes through it:
 *
 *   path 1  createTDocTableFromData_()      the old table builder (no caller today)
 *   path 2  addAbstractsForTables_()        the sweep: an update with abstracts on,
 *                                           "Update Abstracts", the abstracts phase of a build
 *   path 3  refreshUploadedTdocMetadata_()  a TDoc of the report that is now uploaded
 *   path 4  completeInsertedUploadedTdoc_() a TDoc inserted when it is already uploaded
 *
 * For each path, with a withdrawn TDoc: no Reviewer request, no write to the
 * "no summary" cache and none removed from it. Beside each, the same case
 * with a TDoc that is not withdrawn asks exactly once -- so that a test
 * cannot pass because nothing would have been asked anyway.
 *
 * Everything runs the real build, the real update and the real sweep on the
 * fake document; the Reviewer and the Docs API are fakes.
 *
 * Run: node tests/withdrawn-no-abstract.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport, zip, CACHE } = require('./helpers/template-report.js');

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

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CODE_ONLY = CODE.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
const functionSource = (name) => { const start = CODE.indexOf('\nfunction ' + name + '('); return start === -1 ? null : CODE.slice(start + 1, CODE.indexOf('\n}\n', start) + 2); };
const S = loadCode().sandbox;
const W = 'S4aA269001', N = 'S4aA269002', O = 'S4aA269003';
const t = (id, status, uploaded, agenda) => ({ id: id, agenda: agenda || '4.3', status: status, uploaded: !!uploaded });
const stale = () => JSON.stringify({ ts: Date.now() - 2 * 3600 * 1000, statusCode: 404 });
const skipLogs = (r) => r.logs.filter((l) => /^No abstract is asked for |^Upload of \S+: withdrawn/.test(l));
/** What was done to the "no summary" cache for these TDocs. (An abstract that arrives removes the entry of its own TDoc, as it always did.) */
const cacheOf = (r, ids) => r.cacheWrites.filter((w) => ids.some((id) => w.indexOf(id) !== -1));
const ON = { FETCH_ABSTRACTS_ON_UPDATE: 'true' };

// ====================================================================
console.log('1. the rule');
{
  check('a status is withdrawn whatever its case and the space around it',
    ['withdrawn', 'Withdrawn', ' WITHDRAWN ', 'withdrawn by the source', 'Withdrawn\n'].map((v) => S.isWithdrawnStatus_(v)), [true, true, true, true, true]);
  check('no other status is, and no missing one',
    ['available', 'reserved', 'revised', 'agreed', 'noted', 'not treated', 'postponed', 'merged', '', null, undefined, 0].map((v) => S.isWithdrawnStatus_(v)), Array(12).fill(false));
  check('it is the comparison the status summary and the styling of a Status cell make: lower case, "withdrawn" anywhere in it (source)',
    [/\.trim\(\)\.toLowerCase\(\)\.indexOf\('withdrawn'\) !== -1/.test(functionSource('isWithdrawnStatus_')), /statusLower\.includes\('withdrawn'\)/.test(functionSource('analyzeReportStatus')), /v\.includes\('withdrawn'\)/.test(functionSource('styleStatusCell_'))], [true, true, true]);

  const table = (status) => ({ getNumRows: () => 2, getRow: (i) => ({ getNumCells: () => 2, getCell: (k) => ({ getText: () => [['TDoc', W], ['Status', status]][i][k] }) }) });
  check('report withdrawn, or Portal withdrawn, or both: not asked -- and which of them says so',
    [S.abstractFetchBlockedBy_(table('withdrawn'), 'available'), S.abstractFetchBlockedBy_(table('available'), 'Withdrawn'), S.abstractFetchBlockedBy_(table('withdrawn'), 'withdrawn'),
      S.abstractFetchBlockedBy_(table('Withdrawn'), undefined), S.abstractFetchBlockedBy_(table('available'), ''), S.abstractFetchBlockedBy_(table('agreed'), 'agreed'), S.abstractFetchBlockedBy_(table(''), undefined)],
    ['report', 'portal', 'portal', 'report', '', '', '']);
  check('a table without a Status row, no table, and a table that cannot be read: the Portal status alone decides; nothing throws',
    [S.abstractFetchBlockedBy_({ getNumRows: () => 0 }, ''), S.abstractFetchBlockedBy_(null, 'withdrawn'), S.abstractFetchBlockedBy_(null, 'available'), S.abstractFetchBlockedBy_({ getNumRows: () => { throw new Error('gone'); } }, 'available'),
      S.abstractFetchBlockedBy_({ getNumRows: () => { throw new Error('gone'); } }, 'withdrawn')], ['', 'portal', '', '', 'portal']);
  check('the status of a TDoc in the list: trimmed; "" without a status column',
    [S.tdocListStatus_({ row: ['x', ' withdrawn '], statusCol: 1 }), S.tdocListStatus_({ row: ['x', null], statusCol: 1 }), S.tdocListStatus_({ row: ['x', 'withdrawn'], statusCol: -1 }), S.tdocListStatus_(null),
      S.tdocListStatuses_([{ row: [W, 'withdrawn'], tdocCol: 0, statusCol: 1 }, { row: [N, 'available'], tdocCol: 0, statusCol: 1 }, { row: ['', 'withdrawn'], tdocCol: 0, statusCol: 1 }])],
    ['withdrawn', '', '', '', { [W]: 'withdrawn', [N]: 'available' }]);
}

// ====================================================================
console.log('2. one place asks the Reviewer for an abstract, and the rule is the first thing it does (source)');
{
  const fetch = functionSource('fetchAndAddAbstract_');
  const guardAt = fetch.indexOf('abstractFetchBlockedBy_(table, portalStatus)');
  check('fetchAndAddAbstract_() checks the rule before the token, the "no summary" cache and the request',
    [guardAt !== -1, guardAt < fetch.indexOf('getReviewerApiTokenForRun_('), guardAt < fetch.indexOf('isReviewerNoSummaryCached_('), guardAt < fetch.indexOf('UrlFetchApp.fetch('), guardAt < fetch.indexOf('try {')], [true, true, true, true, true]);
  check('and returns false then, having done nothing else', /if \(withdrawnBy\) \{\n    perfCount_\('abstracts not asked for: TDoc withdrawn'\);\n    Logger\.log\(.*\);\n    return false;\n  \}\n  try \{/.test(fetch), true);
  // The only requests for a summary: this one, and the two connection tests, which probe the agenda TDoc for authentication.
  const summaryRequests = CODE_ONLY.split('\nfunction ').filter((f) => /\/summary\?type=summary/.test(f)).map((f) => f.slice(0, f.indexOf('(')));
  check('the Reviewer summary is requested in fetchAndAddAbstract_() and in the two connection tests, nowhere else', summaryRequests, ['fetchAndAddAbstract_', 'testAllConnections', 'testReviewerApi']);
  const callers = CODE_ONLY.split('\nfunction ').filter((f) => /[^_]fetchAndAddAbstract_\(/.test(f.slice(f.indexOf('(')))).map((f) => f.slice(0, f.indexOf('(')));
  check('it has the four callers this test goes through, and no other', callers, ['createTDocTableFromData_', 'addAbstractsForTables_', 'refreshUploadedTdocMetadata_', 'completeInsertedUploadedTdoc_']);
  check('the three that run in an update hand it the status of the TDoc list',
    [/fetchAndAddAbstract_\(table, parsedTdoc\.raw, context, portalStatus\)/.test(functionSource('addAbstractsForTables_')), /fetchAndAddAbstract_\(table, parsed\.raw, context, portalStatus\)/.test(functionSource('refreshUploadedTdocMetadata_')),
      /fetchAndAddAbstract_\(table, parsed\.raw, context, tdocListStatus_\(tdocData\)\)/.test(functionSource('completeInsertedUploadedTdoc_')),
      /addAbstractsForTables_\(body, context, Object\.assign\(tdocsNotUploadedYet_\(allTdocs\), abstractAskedThisRun\), tdocListStatuses_\(allTdocs\)\)/.test(functionSource('continuousUpdateCore_'))], [true, true, true, true]);
  check('the upload completion asks the rule before it forgets a "no summary" answer',
    functionSource('refreshUploadedTdocMetadata_').indexOf('abstractFetchBlockedBy_(table, portalStatus)') < functionSource('refreshUploadedTdocMetadata_').indexOf('clearReviewerNoSummaryCache_('), true);
  check('no code removes an Abstract row: nothing is taken away because a TDoc is withdrawn', /removeRow\([^)]*\)[^\n]*[Aa]bstract|[Aa]bstract[^\n]*removeRow\(/.test(CODE_ONLY), false);
}

// ====================================================================
console.log('path 1: createTDocTableFromData_() (the old table builder)');
{
  const run = (status) => {
    const r = templateReport({ release: false, reviewer: { [W]: 'Summary.' } });
    r.mark();
    const table = r.exec(() => r.s.createTDocTableFromData_(r.body, [['TDoc', W], ['Title', 'T'], ['Source', 'S'], ['Contact', 'C'], ['Agenda Item', '4.3'], ['Status', status]], null, 0));
    const labels = []; for (let i = 0; i < table.getNumRows(); i++) labels.push(table.getRow(i).getCell(0).getText());
    return [r.requests.slice(), r.cacheWrites.filter((w) => /^set /.test(w)), r.tokenReads, labels.indexOf('Abstract') !== -1];
  };
  check('withdrawn: no request, no cache write, the token is not even looked up, no Abstract row', run('withdrawn'), [[], [], 0, false]);
  check('Withdrawn (as the Portal may write it): the same', run('Withdrawn'), [[], [], 0, false]);
  check('available: asked once, and the abstract is added (the path does ask)', run('available'), [[W], [], 1, true]);
}

// ====================================================================
console.log('path 2: the sweep -- text statuses');
{
  // Three TDocs without an abstract: one withdrawn in the report and in the list, one withdrawn in the report only, one not.
  const make = (props) => {
    const r = templateReport({ release: false, props: props, reviewer: { [W]: 'Summary W.', [N]: 'Summary N.', [O]: 'Summary O.' }, tdocs: [t(W, 'withdrawn', true), t(N, 'available', true), t(O, 'available', true)] });
    r.s.completeInsertedUploadedTdoc_ = () => false;              // this part is about the sweep alone
    r.exec(() => r.s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }));
    r.user(() => r.api.statusCell(O).setText('Withdrawn'));        // decided in the meeting; the list still says available
    return r;
  };
  {
    const r = make(ON);
    check('before: no table has an abstract', [W, N, O].map((id) => r.field(id, 'Abstract')), [undefined, undefined, undefined]);
    r.mark();
    const result = r.update();
    check('an update with abstracts on asks for the one TDoc that is not withdrawn', [result.success, r.requests, [W, N, O].map((id) => r.field(id, 'Abstract'))], [true, [N], [undefined, 'Summary N.', undefined]]);
    check('nothing is written to or removed from the "no summary" cache for the withdrawn ones', cacheOf(r, [W, O]), []);
    check('the log counts one candidate and one request: a withdrawn TDoc is not an attempt', r.logs.filter((l) => /^Abstracts:/.test(l)), ['Abstracts: 1 candidate table(s) processed, 1 Reviewer request(s) made, 0 negative-cache skip(s), 1 abstract row(s) inserted']);
    check('and says why the two were not asked for', skipLogs(r), ['No abstract is asked for ' + W + ': withdrawn (TDoc list)', 'No abstract is asked for ' + O + ': withdrawn (Status in the report)']);
    check('the counter of Reviewer requests is 1, and the two are counted as not asked for', [r.s.perfCounterValue_('Reviewer API requests'), r.s.perfCounterValue_('abstracts not asked for: TDoc withdrawn')], [1, 2]);
  }
  {
    const r = make({});
    r.mark();
    r.updateAbstracts();
    check('"Update Abstracts" (no TDoc list at hand): the Status in the report decides -- the same one request', [r.requests, [W, N, O].map((id) => r.field(id, 'Abstract')), cacheOf(r, [W, O])], [[N], [undefined, 'Summary N.', undefined], []]);
    check('and its message counts one candidate', r.ui.alerts.map((a) => a[1]), ['Abstract step completed: 1 candidate table(s) processed, 1 abstract(s) inserted.']);
  }
  {
    const r = templateReport({ release: false, reviewer: { [W]: 'Summary W.', [N]: 'Summary N.' }, tdocs: [t(W, 'withdrawn', true), t(N, 'available', true)] });
    r.mark();
    const built = r.build();
    check('Build Report from Scratch: its abstracts phase asks for the one that is not withdrawn', [built.ok, r.requests, r.field(W, 'Abstract'), r.field(N, 'Abstract'), cacheOf(r, [W])], [true, [N], undefined, 'Summary N.', []]);
    check('and counts one candidate', built.phases.filter((p) => p.name === 'abstracts').map((p) => p.detail), ['1 candidate table(s), 1 abstract(s) inserted']);
  }
  {
    // The Reviewer would answer "no summary" (404), which is cached for a day. For a withdrawn TDoc it is never asked, so nothing is cached.
    const r = templateReport({ release: false, props: ON, reviewer: {}, tdocs: [t(W, 'withdrawn', true), t(N, 'available', true)] });
    r.s.completeInsertedUploadedTdoc_ = () => false;
    r.exec(() => r.s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }));
    r.mark();
    r.update();
    check('a "no summary" answer is cached for the TDoc that was asked for, and none for the withdrawn one', [r.requests, r.cacheWrites, r.docProps.getProperty(CACHE(W))], [[N], ['set ' + CACHE(N)], null]);
  }
}

// ====================================================================
console.log('path 2: the sweep -- a Status that is a native dropdown');
{
  const make = (props) => {
    const r = templateReport({ props: props, token: false, reviewer: { [W]: 'Summary W.', [N]: 'Summary N.', [O]: 'Summary O.' }, tdocs: [t(W, 'withdrawn', true), t(N, 'available', true), t(O, 'available', true)] });
    r.build();                                                     // no token: the build asks for no abstract and makes the dropdowns
    r.scriptProps.setProperty('REVIEWER_API_TOKEN', 'synthetic-test-token');
    r.user(() => r.api.pick(O, 'withdrawn'));                      // decided in the meeting, in the dropdown
    return r;
  };
  {
    const r = make(ON);
    check('before: the three statuses are dropdowns, and no table has an abstract', [[W, N, O].map((id) => r.shown(id)), [W, N, O].map((id) => r.field(id, 'Abstract'))],
      [['dropdown:withdrawn', 'dropdown:available', 'dropdown:withdrawn'], [undefined, undefined, undefined]]);
    check('DocumentApp sees no text in such a cell: the rule has to read the dropdown', [W, N, O].map((id) => r.api.statusCell(id).getText()), ['', '', '']);
    r.mark();
    const result = r.update();
    check('an update with abstracts on asks for the one TDoc that is not withdrawn', [result.success, r.requests, [W, N, O].map((id) => r.field(id, 'Abstract')), cacheOf(r, [W, O])], [true, [N], [undefined, 'Summary N.', undefined], []]);
    check('the dropdowns are as they were', [W, N, O].map((id) => r.shown(id)), ['dropdown:withdrawn', 'dropdown:available', 'dropdown:withdrawn']);
  }
  {
    const r = make({});
    const reads = r.api.count('get');
    r.mark();
    r.updateAbstracts();
    check('"Update Abstracts": the value selected in the dropdown decides -- the same one request', [r.requests, [W, N, O].map((id) => r.field(id, 'Abstract')), cacheOf(r, [W, O])], [[N], [undefined, 'Summary N.', undefined], []]);
    check('the document is read through the API once for all tables, not once per table', r.api.count('get') - reads, 1);
  }
  {
    // A dropdown that cannot be read is not taken for a withdrawal: the TDoc is asked for, as before this rule.
    const r = make({});
    r.s.Docs.Documents.get = () => { throw new Error('synthetic API failure'); };
    r.mark();
    r.updateAbstracts();
    check('when the dropdowns cannot be read, their TDocs count as not withdrawn (nothing is guessed)', r.requests, [W, N, O]);
  }
}

// ====================================================================
console.log('path 3: a TDoc of the report that is uploaded now (reserved -> uploaded)');
{
  // The report has W and N from when they were reserved: no link. The Reviewer had nothing for them then, and each has a
  // "no summary" answer from that time. Now it has a summary of both.
  const make = (o) => {
    const r = templateReport({ release: o.release, token: o.dropdowns ? false : undefined, reviewer: {}, tdocs: [t(W, 'reserved', false), t(N, 'reserved', false)] });
    r.build();
    if (o.dropdowns) r.scriptProps.setProperty('REVIEWER_API_TOKEN', 'synthetic-test-token');
    r.docProps.setProperty(CACHE(W), stale());
    r.docProps.setProperty(CACHE(N), stale());
    Object.assign(r.reviewer, { [W]: 'Summary W.', [N]: 'Summary N.' });
    return r;
  };
  {
    const r = make({ release: false });
    check('before: no link, no abstract, Status reserved', [r.linkOf(W), r.linkOf(N), r.field(W, 'Abstract'), r.field(W, 'Status')], [null, null, undefined, 'reserved']);
    // The Portal: W was uploaded and then withdrawn; N was uploaded.
    Object.assign(r.tdoc(W), { status: 'withdrawn', uploaded: true });
    Object.assign(r.tdoc(N), { status: 'available', uploaded: true });
    r.mark();
    const result = r.update();
    check('reserved -> withdrawn: the status is synced and the link is added -- the completion of the upload still happens', [result.success, r.field(W, 'Status'), r.linkOf(W)], [true, 'withdrawn', zip(W)]);
    check('but no abstract is asked for it, and its "no summary" answer is neither removed nor renewed', [r.requests, r.field(W, 'Abstract'), cacheOf(r, [W]), r.docProps.getProperty(CACHE(W)) !== null], [[N], undefined, [], true]);
    check('the TDoc that is not withdrawn is completed as in T-2026.10.6: link, the old answer forgotten, one request, the abstract',
      [r.linkOf(N), cacheOf(r, [N])[0], r.docProps.getProperty(CACHE(N)), r.field(N, 'Abstract')], [zip(N), 'delete ' + CACHE(N), null, 'Summary N.']);
    check('the log counts one abstract attempt', r.logs.filter((l) => /^Upload completion:/.test(l)), ['Upload completion: 2 TDoc link(s) added, 1 abstract attempt(s), 2 registration-table link(s) added']);
    check('and says that the withdrawn one was not asked for', skipLogs(r), ['Upload of ' + W + ': withdrawn, no abstract is asked for']);
    r.mark();
    r.update();
    check('the next update asks nothing: the transition was acted on, and nothing is retried', [r.requests, r.cacheWrites, r.field(W, 'Abstract')], [[], [], undefined]);
  }
  {
    // The Status was set to withdrawn by hand; the Portal says available and uploaded.
    const r = make({ release: false });
    r.user(() => r.api.statusCell(W).setText('withdrawn'));
    Object.assign(r.tdoc(W), { status: 'available', uploaded: true });
    r.mark();
    r.update();
    check('withdrawn in the report only: the link is added, the Status stays, nothing is asked', [r.linkOf(W), r.field(W, 'Status'), r.requests, r.cacheWrites, r.field(W, 'Abstract')], [zip(W), 'withdrawn', [], [], undefined]);
  }
  {
    // The same two cases with native dropdowns.
    const r = make({ dropdowns: true });
    check('before: the statuses are dropdowns', [r.shown(W), r.shown(N)], ['dropdown:reserved', 'dropdown:reserved']);
    r.user(() => r.api.pick(W, 'withdrawn'));
    Object.assign(r.tdoc(W), { status: 'available', uploaded: true });
    Object.assign(r.tdoc(N), { status: 'withdrawn', uploaded: true });
    r.mark();
    const result = r.update();
    check('dropdown withdrawn (report), and Portal withdrawn with a dropdown: both get their link, neither is asked for',
      [result.success, r.linkOf(W), r.linkOf(N), r.requests, r.cacheWrites, r.field(W, 'Abstract'), r.field(N, 'Abstract')], [true, zip(W), zip(N), [], [], undefined, undefined]);
    check('the dropdown the user set stays; the other follows the Portal', [r.shown(W), r.shown(N)], ['dropdown:withdrawn', 'dropdown:withdrawn']);
  }
}

// ====================================================================
console.log('path 4: a TDoc that is inserted when it is already uploaded');
{
  const run = (release) => {
    const r = templateReport({ release: release, token: release === false ? undefined : false, reviewer: { [W]: 'Summary W.', [N]: 'Summary N.' }, tdocs: [t(O, 'available', false)] });
    r.build();
    if (release !== false) r.scriptProps.setProperty('REVIEWER_API_TOKEN', 'synthetic-test-token');
    r.tdocs.push(t(W, 'withdrawn', true), t(N, 'available', true));
    r.mark();
    const result = r.update();
    return [result.success, r.linkOf(W), r.requests.slice(), r.field(W, 'Abstract'), r.field(N, 'Abstract'), cacheOf(r, [W]),
      r.logs.filter((l) => /^Upload completion:/.test(l)), release === false ? r.field(W, 'Status') : r.shown(W)];
  };
  check('text statuses: the withdrawn TDoc is inserted with its link and is not asked for; the other is asked once',
    run(false), [true, zip(W), [N], undefined, 'Summary N.', [], ['Upload completion: 0 TDoc link(s) added, 1 abstract attempt(s), 0 registration-table link(s) added'], 'withdrawn']);
  check('a report with status dropdowns: the same, and the new TDoc gets its dropdown',
    run(undefined), [true, zip(W), [N], undefined, 'Summary N.', [], ['Upload completion: 0 TDoc link(s) added, 1 abstract attempt(s), 0 registration-table link(s) added'], 'dropdown:withdrawn']);
}

// ====================================================================
console.log('3. an abstract that is there stays');
{
  const r = templateReport({ release: false, props: ON, reviewer: { [W]: 'Summary W.' }, tdocs: [t(W, 'available', true), t(N, 'available', true)] });
  r.build();
  check('built: W has its abstract', r.field(W, 'Abstract'), 'Summary W.');
  r.tdoc(W).status = 'withdrawn';
  r.mark();
  r.update();
  check('available -> withdrawn: the status is synced, the abstract is kept, and nothing is asked for W again', [r.field(W, 'Status'), r.field(W, 'Abstract'), r.requests.filter((id) => id === W)], ['withdrawn', 'Summary W.', []]);
  r.mark();
  r.updateAbstracts();
  r.update();
  check('nor by "Update Abstracts" or a later update', [r.field(W, 'Abstract'), r.requests.filter((id) => id === W)], ['Summary W.', []]);
}

// ====================================================================
console.log('4. withdrawn, and active again later');
{
  const r = templateReport({ release: false, props: ON, reviewer: { [W]: 'Summary W.' }, tdocs: [t(W, 'withdrawn', true)] });
  r.build();
  r.mark();
  r.update();
  r.update();
  check('while it is withdrawn, update after update asks nothing and caches nothing', [r.requests, r.cacheWrites, r.docProps.getProperty(CACHE(W)), r.field(W, 'Abstract')], [[], [], null, undefined]);
  // The Portal says available again. The status sync does not take a Status off "withdrawn" by itself (unchanged rule).
  r.tdoc(W).status = 'available';
  r.mark();
  r.update();
  check('Portal active again, report still withdrawn: still nothing is asked (the report says withdrawn)', [r.field(W, 'Status'), r.requests], ['withdrawn', []]);
  r.user(() => r.api.statusCell(W).setText('available'));
  r.mark();
  r.update();
  check('once the report says so too, the normal sweep gets the abstract -- at once: no "no summary" answer stands in its way', [r.requests, r.field(W, 'Abstract'), r.cacheWrites.filter((w) => /^set /.test(w))], [[W], 'Summary W.', []]);
}
{
  const r = templateReport({ release: false, reviewer: { [W]: 'Summary W.' }, tdocs: [t(W, 'withdrawn', true)] });
  r.build();
  r.user(() => r.api.statusCell(W).setText('available'));
  r.mark();
  r.updateAbstracts();
  check('"Update Abstracts" gets it as well, with abstracts on update switched off', [r.requests, r.field(W, 'Abstract')], [[W], 'Summary W.']);
}

// ====================================================================
console.log('5. TDocs that are not withdrawn: as in T-2026.10.8');
{
  const statuses = ['available', 'reserved', 'revised', 'agreed', 'noted', 'postponed', 'not treated', ''];
  const r = templateReport({ release: false, reviewer: {}, tdocs: statuses.map((status, i) => t('S4aA26910' + i, status, true)) });
  r.s.completeInsertedUploadedTdoc_ = () => false;
  r.exec(() => r.s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }));
  r.mark();
  r.updateAbstracts();
  check('every other status is asked for, once each', r.requests, statuses.map((status, i) => 'S4aA26910' + i));
  check('and no line of the log speaks of a withdrawal', skipLogs(r), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll withdrawn-TDoc abstract checks passed.');
process.exitCode = failures ? 1 : 0;
