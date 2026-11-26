/**
 * Issue #28 phase A, after the first real-document run (2026-11-25): an update stopped from outside (the 6-minute
 * limit) is no longer silent. The journal (UPDATE_RUN) is found by the next update and recorded (UPDATE_INTERRUPTED);
 * a TDoc table left twice with identical content (a killed move) keeps one; one with different content is named, not
 * touched; the record is closed by an update that completes. New TDoc tables and revision moves stop at their time
 * budget and are left to the next update instead of running into the limit.
 */
const { templateReport } = require('./helpers/template-report');

let failures = 0;
function check(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) failures++;
  console.log((ok ? '  ok   ' : '  FAIL ') + name + (ok ? '' : '\n         expected ' + JSON.stringify(expected) + '\n         actual   ' + JSON.stringify(actual)));
}

const TDOCS = [
  { id: 'S4aA269001', agenda: '4.3', status: 'agreed', uploaded: true },
  { id: 'S4aA269002', agenda: '4.3', status: 'noted', uploaded: true },
  { id: 'S4aA269003', agenda: '4.4', status: 'available', uploaded: true }];
const prop = (r, k) => r.docProps.getProperty(k);
const json = (r, k) => JSON.parse(prop(r, k) || 'null');
const duplicate = (r, id, edit) => r.exec(() => {
  const t = r.table(id);
  const copy = t.copy();
  if (edit) for (let i = 0; i < copy.getNumRows(); i++) if (copy.getRow(i).getCell(0).getText() === edit.row) copy.getRow(i).getCell(1).setText(edit.text);
  r.body.insertTable(r.body.getChildIndex(t) + 1, copy);
});
const count = (r, id) => r.order().filter((n) => n === id).length;

console.log('A. an update that ends: no journal left, nothing recorded');
{
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  const res = r.update();
  check('after an update: no UPDATE_RUN, no UPDATE_INTERRUPTED, nothing named in the result', [prop(r, 'UPDATE_RUN'), prop(r, 'UPDATE_INTERRUPTED'), res.interrupted], [null, null, undefined]);
}

console.log('B. the journal follows the stages; an update that does not end leaves it');
{
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  r.exec(() => { r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null); r.s.perfTimed_('revision ordering/rearrangement (rearrangeRevisionTables_)', () => 0); });
  const run = json(r, 'UPDATE_RUN');
  check('a running update: UPDATE_RUN with its start and the stage it reached (from perfTimed_)', [!!run.startedAt, run.stage], [true, 'revision ordering/rearrangement (rearrangeRevisionTables_)']);
  // The execution ends here (killed: no finally). The next update finds the journal.
  r.exec(() => r.s.endReportUpdate_());
  const res = r.update();
  const rec = json(r, 'UPDATE_INTERRUPTED');
  check('the next update: the interruption named in its result with start and stage; closed because this update completed', [res.interrupted.stage, res.interrupted.removed, res.interrupted.differing, rec, prop(r, 'UPDATE_RUN')],
    ['revision ordering/rearrangement (rearrangeRevisionTables_)', [], [], null, null]);
  check('... and it is logged ("did not end", with its stage)', r.logs.some((l) => /Update journal: the update started .* did not end \(it was stopped during "revision ordering/.test(String(l))), true);
}

console.log('C. a TDoc table left twice by a killed move: the identical copy removed; a different one named, not touched');
{
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  duplicate(r, 'S4aA269002');
  r.docProps.setProperty('UPDATE_RUN', JSON.stringify({ startedAt: '2026-11-25T14:20:00.000Z', stage: 'revision ordering/rearrangement (rearrangeRevisionTables_)' }));
  check('before: S4aA269002 twice', count(r, 'S4aA269002'), 2);
  const res = r.update();
  check('the identical copy removed, the interruption named; the record closed by the next (complete) update', [count(r, 'S4aA269002'), res.interrupted.removed, res.interrupted.differing, !!prop(r, 'UPDATE_INTERRUPTED'), (r.update().incomplete || []).length, prop(r, 'UPDATE_INTERRUPTED')], [1, ['S4aA269002'], [], true, 0, null]);
  check('... the other TDoc tables untouched, in order', r.order(), ['S4aA269001', 'S4aA269002', 'S4aA269003']);
}
{
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  duplicate(r, 'S4aA269001', { row: 'Minutes', text: 'Typed into the copy only.' });
  r.docProps.setProperty('UPDATE_RUN', JSON.stringify({ startedAt: '2026-11-25T14:20:00.000Z', stage: 'x' }));
  const res = r.update();
  check('a copy with different content: both kept, named; the record stays open', [count(r, 'S4aA269001'), res.interrupted.differing, json(r, 'UPDATE_INTERRUPTED').differing], [2, ['S4aA269001'], ['S4aA269001']]);
  const again = r.update();
  check('... the next update names it again (the record stays until it is resolved by hand)', [again.interrupted && again.interrupted.differing, count(r, 'S4aA269001')], [['S4aA269001'], 2]);
  // Resolved by hand: the copy removed.
  const tables = r.tdocTables().filter((t) => t.getRow(0).getCell(1).getText() === 'S4aA269001');
  r.exec(() => tables[1].removeFromParent());
  const after = r.update();
  check('resolved by hand: the next update closes the record', [after.interrupted && after.interrupted.differing, prop(r, 'UPDATE_INTERRUPTED')], [[], null]);
}
{
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  duplicate(r, 'S4aA269003');
  const res = r.update();
  check('without an interruption, nothing is removed (a duplicate is only repaired after a killed update)', [count(r, 'S4aA269003'), res.interrupted], [2, undefined]);
}

console.log('D. time budgets: the rest is left to the next update, not killed');
{
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  r.tdocs.push({ id: 'S4aA269004', agenda: '4.4', status: 'available', uploaded: true });
  // The loop's own clock: 101 s pass with its first TDoc.
  let now = 1000000;
  r.s.updateBudgetNow_ = () => { const v = now; now += 101000; return v; };
  const res = r.update();
  check('past 100 s in the TDoc loop: no new TDoc table, the update "in part" with the TDoc tables named', [count(r, 'S4aA269004'), (res.incomplete || []).some((x) => /^TDoc tables from S4aA2690\d\d on/.test(x))], [0, true]);
  r.s.updateBudgetNow_ = () => Date.now();
  const next = r.update();
  check('the next update (in time) inserts it', [count(r, 'S4aA269004'), (next.incomplete || []).filter((x) => /^TDoc tables/.test(x))], [1, []]);
}
{
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  r.exec(() => { r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null); });
  let t0 = 5000000;
  r.s.updateBudgetNow_ = () => t0;
  const early = r.exec(() => [r.s.revisionMovesOverBudget_(), (t0 += 59000, r.s.revisionMovesOverBudget_())]);
  t0 += 2000;
  const late = r.exec(() => [r.s.revisionMovesOverBudget_(), r.s.revisionMovesOverBudget_(), r.s.reportUpdateDeferred_()]);
  check('revision moves: allowed for 60 s of moves; after that refused, and "revision placement" left to the next update once', [early, late[0], late[1], late[2].filter((x) => x === 'revision placement').length], [[false, false], true, true, 1]);
  r.exec(() => { r.s.updateJournalEnd_({ success: true }); r.s.endReportUpdate_(); });
  check('outside an update (no journal): no budget applies', r.exec(() => [r.s.tdocLoopOverBudget_('x'), r.s.revisionMovesOverBudget_()]), [false, false]);
}

console.log('E. controlled fixtures: a move killed half-way, and an update that converges under its budgets');
{
  // A revision filed under another agenda item: the update moves it below its original (copy, then remove).
  // At the build there is no revision yet; then the list names S4aA269005 as the revision of S4aA269001, filed under
  // 4.4 after S4aA269003: the update inserts it there (its original's table is found by the TDoc, not the list) or
  // below its original, and the re-arrangement moves it below S4aA269001 (copy, then remove) where it is not.
  const LIST = () => [
    { id: 'S4aA269001', agenda: '4.3', status: 'available', uploaded: true },
    { id: 'S4aA269002', agenda: '4.3', status: 'noted', uploaded: true },
    { id: 'S4aA269003', agenda: '4.4', status: 'available', uploaded: true }];
  const revise = (x) => {
    x.tdocs[0].status = 'revised'; x.tdocs[0].revisedTo = 'S4aA269005';
    x.tdocs.push({ id: 'S4aA269005', agenda: '4.4', status: 'agreed', isRevisionOf: 'S4aA269001', uploaded: true });
    // Placed at the end of 4.4 (as if its original had not been found), so that the update has to move it.
    x.s.findParentRevisedToTable_ = () => null;
  };
  const content = (r) => r.order().map((n) => n + ':' + r.field(n, 'Status'));
  // The control: built, the revision arrives, updated without any interruption.
  const control = templateReport({ tdocs: LIST() });
  control.build(); revise(control); control.update(); control.update();
  // The same, but the first update is killed inside its first move: the copy is inserted, the original not yet
  // removed, and no finally runs (no updateJournalEnd_).
  const r = templateReport({ tdocs: LIST() });
  r.build(); revise(r);
  const realRemove = r.s.removeTableWithItsSeparator_;
  const realEnd = r.s.updateJournalEnd_;
  let killed = false;
  r.s.removeTableWithItsSeparator_ = function () { if (!killed) { killed = true; throw new Error('Exceeded maximum execution time (simulated)'); } return realRemove.apply(this, arguments); };
  r.s.updateJournalEnd_ = function () {};
  r.update();
  r.s.removeTableWithItsSeparator_ = realRemove;
  r.s.updateJournalEnd_ = realEnd;
  const left = json(r, 'UPDATE_RUN');
  check('the killed update: the revision is in the report twice (copy and original), its journal left at the move', [killed, count(r, 'S4aA269005'), /rearrangeRevisionTables_/.test(left && left.stage)], [true, 2, true]);
  const next = r.update();
  check('the next update: the interruption found (its stage), the identical copy removed; it says it is in part, the record stays open', [/rearrangeRevisionTables_/.test(next.interrupted && next.interrupted.stage), next.interrupted && next.interrupted.removed, (next.incomplete || []).some((x) => /after the repair of the interrupted update/.test(x)), !!prop(r, 'UPDATE_INTERRUPTED'), prop(r, 'UPDATE_RUN')], [true, ['S4aA269005'], true, true, null]);
  const third = r.update();
  check('the update after it: complete, the record closed, the report exactly what an uninterrupted update gives (order, statuses -- its Status a dropdown again)', [third.incomplete || [], prop(r, 'UPDATE_INTERRUPTED'), content(r)], [[], null, content(control)]);
  const again = r.update();
  check('... a further update changes nothing', [content(r), again.interrupted], [content(control), undefined]);
}
{
  // Twelve new TDocs at once; the TDoc loop may take 100 s; here a new TDoc table "takes" 30 s and a TDoc already in
  // the report 1 s -- about three new ones per update.
  const base = [{ id: 'S4aA269001', agenda: '4.3', status: 'agreed', uploaded: true }];
  const late = [];
  for (let i = 10; i < 22; i++) late.push({ id: 'S4aA2690' + i, agenda: i % 2 ? '4.3' : '4.4', status: i % 3 ? 'available' : 'noted', uploaded: true });
  const content = (x) => x.order().map((n) => n + ':' + x.field(n, 'Status'));
  const control = templateReport({ tdocs: base.slice() });
  control.build();
  late.forEach((t) => control.tdocs.push(Object.assign({}, t)));
  control.update();
  const r = templateReport({ tdocs: base.slice() });
  r.build();
  late.forEach((t) => r.tdocs.push(Object.assign({}, t)));
  let clock = 0;
  r.s.updateBudgetNow_ = () => { clock += 1000; return clock; };
  const realInsert = r.s.insertNewTdoc_;
  r.s.insertNewTdoc_ = function () { clock += 30000; return realInsert.apply(this, arguments); };
  const runs = [];
  for (let k = 0; k < 8; k++) {
    clock = 0;
    const res = r.update();
    const partial = (res.incomplete || []).some((x) => /^TDoc tables from/.test(x));
    runs.push(partial ? 'part' : 'whole');
    if (!partial) break;
  }
  check('under the budget the update stops part-way and the next ones go on, until one is whole', [runs[0], runs[runs.length - 1], runs.length > 2], ['part', 'whole', true]);
  check('... then the report is the one a single update without budgets gives (every TDoc, order, statuses)', content(r), content(control));
  r.s.updateBudgetNow_ = () => Date.now();
  r.s.insertNewTdoc_ = realInsert;
  r.update();
  check('... and a further update changes nothing; no cursor left', [content(r), prop(r, 'UPDATE_TDOC_CURSOR')], [content(control), null]);
}

console.log('F. one clock for the whole update (after the second real-document run, 2026-11-26)');
{
  // The real collector of Code.js (the helper replaces it with a stub): taken from the source into the sandbox.
  const fs = require('fs');
  const vm = require('vm');
  const src = fs.readFileSync(require('path').join(__dirname, '..', 'Code.js'), 'utf8').split(String.fromCharCode(13)).join('');
  const fnSource = (name) => { const i = src.indexOf('\nfunction ' + name + '('); const j = src.indexOf('\n}\n', i); if (i < 0 || j < 0) throw new Error('no function ' + name); return src.slice(i + 1, j + 3); };
  const r = templateReport({ tdocs: TDOCS, });
  r.build();
  vm.runInContext(fnSource('collectorUpdate_'), r.s);
  const mailBefore = TDOCS.map((x) => r.field(x.id, 'E-mail Discussion'));
  // The clock of the budgets: the update starts at 0; the collectors begin 196 s into it.
  let clock = 0;
  r.s.updateBudgetNow_ = () => clock;
  const realCfg = r.s.getCollectorConfig_;
  r.s.getCollectorConfig_ = function () { clock = 196000; return realCfg.apply(this, arguments); };
  const before = r.fetches.length;
  const res = r.update();
  const mailFetches = r.fetches.slice(before).filter((u) => /list\.etsi\.org|wa\.exe/i.test(u));
  check('196 s into the update: no mail request is made; the e-mail discussions are left to the next update, not reported as a failure', [mailFetches.length, (res.incomplete || []).some((x) => /^e-mail discussions \(no network request 196 s into the update/.test(x)), res.collectorFailures], [0, true, undefined]);
  check('... and no E-mail Discussion cell is touched (nothing written from a part of the messages)', TDOCS.map((x) => r.field(x.id, 'E-mail Discussion')), mailBefore);
  const runLeft = json(r, 'UPDATE_RUN');
  check('... the update ended (no journal left)', runLeft, null);
  // Earlier (100 s): the request is made.
  r.s.getCollectorConfig_ = function () { clock = 100000; return realCfg.apply(this, arguments); };
  const before2 = r.fetches.length;
  const res2 = r.update();
  check('100 s into the update: the mail request is made, nothing left to the next update for it', [r.fetches.slice(before2).some((u) => /list\.etsi\.org|wa\.exe/i.test(u)), (res2.incomplete || []).some((x) => /^e-mail discussions/.test(x))], [true, false]);
  r.s.getCollectorConfig_ = realCfg;
  r.s.updateBudgetNow_ = () => Date.now();
  check('outside an update: safeFetch_ is never held back by the budget', r.exec(() => { try { r.s.safeFetch_({}, 'https://list.etsi.org/scripts/wa.exe?RSS', {}, 'test'); return 'fetched'; } catch (e) { return e.message; } }), 'fetched');
}
{
  // The stages the next update does again anyway: not started 256 s into the update; the status dropdowns still end it.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  let clock = 0;
  r.s.updateBudgetNow_ = () => clock;
  r.s.collectorUpdate_ = () => { clock = 256000; return { failures: [] }; };
  let summaries = 0;
  const realSummary = r.s.syncDispositionSummaryOnUpdate_;
  r.s.syncDispositionSummaryOnUpdate_ = function () { summaries++; return realSummary.apply(this, arguments); };
  const res = r.update();
  check('256 s into the update: the Goes to Plenary data, the summary and Undisposed Documents left to the next update', [summaries, (res.incomplete || []).includes('the Goes to Plenary data, the Disposition Summary and Undisposed Documents'), res.success], [0, true, true]);
  r.s.collectorUpdate_ = () => { clock = 200000; return { failures: [] }; };
  const res2 = r.update();
  check('200 s into the update: they run', [summaries, (res2.incomplete || []).includes('the Goes to Plenary data, the Disposition Summary and Undisposed Documents')], [1, false]);
  r.s.updateBudgetNow_ = () => Date.now();
}
{
  // The TDoc loop and the moves on the update's clock: no TDoc after 170 s, no move after 200 s, even within their own budgets.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  r.exec(() => { r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null); });
  let t0 = 1000000;
  r.s.updateBudgetNow_ = () => t0;
  r.exec(() => r.s.updateJournalStart_(null));
  const at = (ms, fn) => { t0 = 1000000 + ms; return r.exec(fn); };
  const loop = [at(160000, () => r.s.tdocLoopOverBudget_('S4aA269001')), at(171000, () => r.s.tdocLoopOverBudget_('S4aA269002'))];
  const moves = [at(195000, () => r.s.revisionMovesOverBudget_()), at(201000, () => r.s.revisionMovesOverBudget_())];
  check('the TDoc loop: a TDoc at 160 s, none at 171 s (its own clock had 11 s); the moves: one at 195 s, none at 201 s', [loop, moves, json(r, 'UPDATE_TDOC_CURSOR').tdoc], [[false, true], [false, true], 'S4aA269002']);
  r.exec(() => { r.s.updateJournalEnd_({ success: true }); r.s.endReportUpdate_(); });
  r.s.updateBudgetNow_ = () => Date.now();
}
{
  // The revisions go on where the last update stopped them; a whole pass removes the cursor.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  const order = ['S4aA269001', 'S4aA269002', 'S4aA269003'];
  const run = (fn) => r.exec(() => { r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null); const v = fn(); r.s.updateJournalEnd_({ success: true }); r.s.endReportUpdate_(); return v; });
  const first = run(() => { const o = r.s.revisionsLoopOrder_(order); r.s.revisionsLoopStopped_('S4aA269002'); return o; });
  const second = run(() => { const o = r.s.revisionsLoopOrder_(order); r.s.revisionsLoopCompleted_(); return o; });
  check('revisions stopped before S4aA269002: the next pass starts there and wraps; a whole pass removes the cursor', [first, second, prop(r, 'UPDATE_REVISIONS_CURSOR')], [order, ['S4aA269002', 'S4aA269003', 'S4aA269001'], null]);
  const third = run(() => { r.s.revisionsLoopStopped_('S4aA269003'); r.s.revisionsLoopCompleted_(); return json(r, 'UPDATE_REVISIONS_CURSOR').tdoc; });
  check('a pass that stopped is never taken as whole (the cursor stays)', [third, json(r, 'UPDATE_REVISIONS_CURSOR').tdoc], ['S4aA269003', 'S4aA269003']);
  check('outside an update the order is the report order', r.exec(() => r.s.revisionsLoopOrder_(order)), order);
}
{
  // The journal keeps where the time went; an update stopped from outside shows it to the next one.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  let clock = 0;
  r.s.updateBudgetNow_ = () => clock;
  r.exec(() => { r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null);
    clock = 85000; r.s.perfTimed_('TDoc-list fetch/download+parse (downloadAndGroupTdocs_)', () => 0);
    clock = 190000; try { r.s.updateNetworkGuard_('Revisions fetch'); } catch (e) { /* not at 190 s */ } });
  const run = json(r, 'UPDATE_RUN');
  check('UPDATE_RUN: the stage, the time since the start and the last stages with theirs (the network requests among them)', [run.stage, run.elapsedMs, run.stages.map((x) => x.stage + '@' + x.ms)],
    ['network: Revisions fetch', 190000, ['TDoc-list fetch/download+parse (downloadAndGroupTdocs_)@85000', 'network: Revisions fetch@190000']]);
  r.exec(() => r.s.endReportUpdate_());
  r.s.updateBudgetNow_ = () => Date.now();
  const res = r.update();
  check('the next update names the stopped one with its stage and its time', [res.interrupted.stage, res.interrupted.elapsedMs], ['network: Revisions fetch', 190000]);
}
{
  // The reallocations of an update come from its one Docs API read (no table walk of their own), as before.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  const labels = [];
  const realTables = r.s.getTablesCounted_;
  r.s.getTablesCounted_ = function (body, label) { labels.push(label); return realTables.apply(this, arguments); };
  r.update();
  check('an update reads its reallocations from its Docs API read: getReallocationMap_ walks no table', labels.filter((l) => /getReallocationMap_/.test(String(l))), []);
  r.s.getTablesCounted_ = realTables;
}

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
