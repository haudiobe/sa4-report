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

console.log('G. the revision pairs within the time (after Run 5, 2026-11-26)');
{
  // Six TDocs; three of them revised after the build, their revisions filed under 4.4 so that each has to be moved
  // below its original. Typed Minutes and a typed Disposition note in the originals.
  const LIST = () => [
    { id: 'S4aA269001', agenda: '4.3', status: 'available', uploaded: true },
    { id: 'S4aA269002', agenda: '4.3', status: 'available', uploaded: true },
    { id: 'S4aA269003', agenda: '4.3', status: 'available', uploaded: true },
    { id: 'S4aA269004', agenda: '4.4', status: 'noted', uploaded: true },
    { id: 'S4aA269005', agenda: '4.4', status: 'agreed', uploaded: true },
    { id: 'S4aA269006', agenda: '4.4', status: 'noted', uploaded: true }];
  const PAIRS = [['S4aA269001', 'S4aA269011'], ['S4aA269002', 'S4aA269012'], ['S4aA269003', 'S4aA269013']];
  const setCell = (x, id, label, text) => x.exec(() => { const tb = x.table(id); for (let i = 0; i < tb.getNumRows(); i++) if (tb.getRow(i).getCell(0).getText() === label) tb.getRow(i).getCell(1).setText(text); });
  const prepare = (x) => {
    x.build();
    setCell(x, 'S4aA269001', 'Minutes', 'Typed minutes of 001, kept.');
    setCell(x, 'S4aA269002', 'Disposition', 'Discussed on Tuesday');
    setCell(x, 'S4aA269005', 'Minutes', 'Typed minutes of 005, kept.');
    PAIRS.forEach(([o, rv]) => {
      const t0 = x.tdocs.find((d) => d.id === o); t0.status = 'revised'; t0.revisedTo = rv;
      x.tdocs.push({ id: rv, agenda: '4.4', status: 'agreed', isRevisionOf: o, uploaded: true });
    });
    x.s.findParentRevisedToTable_ = () => null;
  };
  const content = (x) => x.order().map((n) => [n, x.field(n, 'Status'), x.field(n, 'Disposition'), x.field(n, 'Minutes')].join(' | '));
  const control = templateReport({ tdocs: LIST() });
  prepare(control); control.update(); control.update();
  check('the control (no budget): every revision below its original, "Revised to" once, typed text kept', [
    PAIRS.every(([o, rv]) => control.order().indexOf(rv) === control.order().indexOf(o) + 1),
    control.field('S4aA269002', 'Disposition'), control.field('S4aA269001', 'Minutes')],
    [true, 'Discussed on Tuesday — Revised to S4AA269012', 'Typed minutes of 001, kept.']);

  // The same, each pair "costs" 25 s of its own clock (its disposition): the 60 s of the pairs end inside the third pair.
  const r = templateReport({ tdocs: LIST() });
  prepare(r);
  let clock = 0;
  r.s.updateBudgetNow_ = () => clock;
  const seen = [];
  const realDisposition = r.s.setDispositionRevisedTo_;
  r.s.setDispositionRevisedTo_ = function (table, child) { seen.push(child); clock += 25000; return realDisposition.apply(this, arguments); };
  const first = r.update();
  const cursor = json(r, 'UPDATE_PAIRS_CURSOR');
  check('past the budget: the update stops inside the third pair, says it is in part, keeps that pair as the next one (its move not made)', [
    (first.incomplete || []).includes('revision placement'), cursor && cursor.parent, seen.slice(), r.order().indexOf('S4aA269013') === r.order().indexOf('S4aA269003') + 1, first.success],
    [true, 'S4aA269003', ['S4aA269011', 'S4aA269012', 'S4aA269013'], false, true]);
  check('... the pairs before it are done (moved, "Revised to"), typed text kept', [
    r.order().indexOf('S4aA269011') === r.order().indexOf('S4aA269001') + 1, r.order().indexOf('S4aA269012') === r.order().indexOf('S4aA269002') + 1,
    r.field('S4aA269002', 'Disposition'), r.field('S4aA269001', 'Minutes'), r.field('S4aA269005', 'Minutes')],
    [true, true, 'Discussed on Tuesday — Revised to S4AA269012', 'Typed minutes of 001, kept.', 'Typed minutes of 005, kept.']);
  // The next update, in time: it starts at that pair.
  seen.length = 0;
  r.s.setDispositionRevisedTo_ = function (table, child) { seen.push(child); return realDisposition.apply(this, arguments); };
  clock = 0;
  const second = r.update();
  check('the next update starts at the kept pair and goes round once (no pair skipped); complete, no cursor left', [
    seen.slice(), r.logs.some((l) => /the revision pairs go on at S4aA269003/.test(l)), (second.incomplete || []).filter((x) => /revision placement/.test(x)), prop(r, 'UPDATE_PAIRS_CURSOR')],
    [['S4aA269013', 'S4aA269011', 'S4aA269012'], true, [], null]);
  check('interrupted and uninterrupted give the same report (order, statuses, dispositions, typed minutes); no table twice, "Revised to" once', [
    content(r), r.order().filter((n, i, a) => a.indexOf(n) !== i), PAIRS.map(([o]) => (r.field(o, 'Disposition').match(/Revised to/g) || []).length)],
    [content(control), [], [1, 1, 1]]);
  const again = r.update();
  check('a further update changes nothing (idempotent)', [content(r), again.incomplete || [], prop(r, 'UPDATE_PAIRS_CURSOR')], [content(control), [], null]);
  r.s.updateBudgetNow_ = () => Date.now();
  r.s.setDispositionRevisedTo_ = realDisposition;
}
{
  // The in-place check from the update's Docs API read: none of the document while the read is the report's
  // structure; the document again after a move, or when the structure differs from the read.
  const r = templateReport({ tdocs: TDOCS.concat([{ id: 'S4aA269005', agenda: '4.3', status: 'agreed', isRevisionOf: 'S4aA269001', uploaded: true }]) });
  r.tdocs[0].status = 'revised'; r.tdocs[0].revisedTo = 'S4aA269005';
  r.build();
  r.update();
  const below = r.order().indexOf('S4aA269005') === r.order().indexOf('S4aA269001') + 1;
  let asked = 0;
  const realBelow = r.s.isTableDirectlyBelow_;
  r.s.isTableDirectlyBelow_ = function () { asked++; return realBelow.apply(this, arguments); };
  r.update();
  check('a revision already below its original: known from the update\'s read, the document is not asked for positions', [below, asked], [true, 0]);
  // The structure changed in this update (as after a move or a section repair): the document is asked.
  const realStruct = r.s.repairReportStructureIfNeeded_;
  r.s.repairReportStructureIfNeeded_ = function () { realStruct.apply(this, arguments); return { sectionsMoved: 1, emptyParagraphsRemoved: 0 }; };
  asked = 0;
  r.update();
  check('... after a section repair in the same update the read is not trusted: the document is asked', asked > 0, true);
  r.s.repairReportStructureIfNeeded_ = realStruct;
  // A body that differs from the read (an element more): the document is asked.
  asked = 0;
  r.exec(() => { const realCount = r.body.getNumChildren.bind(r.body); r.body.getNumChildren = () => realCount() + 1; });
  r.update();
  check('... a body with another number of elements than the read: the document is asked', asked > 0, true);
  r.s.isTableDirectlyBelow_ = realBelow;
}

{
  // Review of 2026-11-26: the read is trusted for positions only as isTableDirectlyBelow_() would answer, and never
  // after a table was inserted or removed earlier in the update -- even when the counts happen to cancel out.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  const tbl = (n) => ({ table: {} });
  const para = (o) => ({ paragraph: Object.assign({ elements: [{ textRun: { content: '\n' } }], paragraphStyle: { namedStyleType: 'NORMAL_TEXT' } }, o || {}) });
  const ask = (between, extra) => r.exec(() => {
    r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null);
    const content = [tbl(), between, tbl()];
    const snap = { content: content, byNumber: { 'S4aA269001': { index: 0 }, 'S4aA269005': { index: 2 } }, duplicates: {} };
    const realSnap = r.s.reportRunSnapshot_;
    const realCount = r.body.getNumChildren;
    r.s.reportRunSnapshot_ = () => snap;
    r.body.getNumChildren = () => content.length;
    try { if (extra) extra(); return r.s.revisionPairInPlaceBySnapshot_(r.body, 'S4aA269001', 'S4aA269005'); }
    finally { r.s.reportRunSnapshot_ = realSnap; r.body.getNumChildren = realCount; r.s.updateJournalEnd_({ success: true }); r.s.endReportUpdate_(); }
  });
  check('in place from the read: only across empty Normal paragraphs -- not an empty heading, list item, or one with a positioned object', [
    ask(para()), ask(para({ paragraphStyle: { namedStyleType: 'HEADING_3' } })), ask(para({ bullet: { listId: 'x' } })), ask(para({ positionedObjectIds: ['kix.1'] })), ask(para({ elements: [{ textRun: { content: 'x\n' } }] }))],
    [true, false, false, false, false]);
  check('... and not after a table was inserted or removed earlier in the update, even with the same number of elements', ask(para(), () => r.s.updateStructureChanged_()), false);
  // The repair of an interrupted update removes a table: the structure changed.
  const r2 = templateReport({ tdocs: TDOCS });
  r2.build();
  duplicate(r2, 'S4aA269002');
  r2.docProps.setProperty('UPDATE_RUN', JSON.stringify({ startedAt: '2026-11-26T14:00:00.000Z', stage: 'x' }));
  let flagged = false;
  const realFlag = r2.s.updateStructureChanged_;
  r2.s.updateStructureChanged_ = function () { flagged = true; return realFlag.apply(this, arguments); };
  r2.update();
  check('the repair that removed a table copy marks the structure as changed', flagged, true);
  r2.s.updateStructureChanged_ = realFlag;
  // Resuming never splits a chain: a cursor at B (of A->B->C) resumes at A.
  const r3 = templateReport({ tdocs: TDOCS });
  r3.build();
  r3.docProps.setProperty('UPDATE_PAIRS_CURSOR', JSON.stringify({ parent: 'B' }));
  const order = r3.exec(() => { r3.s.markReportUpdateStart_(); r3.s.updateJournalStart_(null); const o = r3.s.revisionPairsOrder_(['X', 'A', 'B', 'Z'], { X: 'Y', A: 'B', B: 'C', Z: 'W' }); r3.s.updateJournalEnd_({ success: true }); r3.s.endReportUpdate_(); return o; });
  check('a cursor inside a chain (A->B, B->C): the pairs resume at the start of that chain', order, ['A', 'B', 'Z', 'X']);
}

console.log('H. every stage of the update on one timeline (diagnostics only, after Run 6, 2026-11-26)');
{
  // The same report and the same list twice: once as it runs, once with no timing at all (perfTimed_ only calls its
  // function). Revisions to move, a "Revised to", typed Minutes and Disposition: the stages have work to do.
  const LIST = () => [
    { id: 'S4aA269001', agenda: '4.3', status: 'available', uploaded: true },
    { id: 'S4aA269002', agenda: '4.3', status: 'available', uploaded: true },
    { id: 'S4aA269004', agenda: '4.4', status: 'noted', uploaded: true }];
  const setCell = (x, id, label, text) => x.exec(() => { const tb = x.table(id); for (let i = 0; i < tb.getNumRows(); i++) if (tb.getRow(i).getCell(0).getText() === label) tb.getRow(i).getCell(1).setText(text); });
  const prepare = (x) => {
    x.build();
    setCell(x, 'S4aA269001', 'Minutes', 'Typed minutes of 001, kept.');
    setCell(x, 'S4aA269002', 'Disposition', 'Discussed on Tuesday');
    const t0 = x.tdocs.find((d) => d.id === 'S4aA269002'); t0.status = 'revised'; t0.revisedTo = 'S4aA269012';
    x.tdocs.push({ id: 'S4aA269012', agenda: '4.4', status: 'agreed', isRevisionOf: 'S4aA269002', uploaded: true });
    x.tdocs.push({ id: 'S4aA269013', agenda: '4.3', status: 'reserved', uploaded: false });
    x.s.findParentRevisedToTable_ = () => null;
  };
  const view = (x) => {
    const props = x.docProps.getProperties();
    const kept = {};
    // Times of the clock (ISO times, the "no summary" cache's epoch milliseconds) are not the report's.
    Object.keys(props).sort().forEach((k) => { kept[k] = String(props[k]).replace(/\d{4}-\d\d-\d\dT[\d:.]+Z/g, '<time>').replace(/"ts":\d{13}/g, '"ts":<time>'); });
    // The whole body: every paragraph (with its heading) and every cell of every table, in order.
    const text = x.body._children.map((e) => {
      if (e.getType() !== 'TABLE') return (e.getHeading ? e.getHeading() + ' ' : '') + (e.getText ? e.getText() : e.getType());
      const rows = [];
      for (let i = 0; i < e.getNumRows(); i++) { const row = e.getRow(i); const cells = []; for (let k = 0; k < row.getNumCells(); k++) cells.push(row.getCell(k).getText()); rows.push(cells.join(' | ')); }
      return rows.join('\n');
    }).join('\n');
    return { text: text, order: x.order(), props: kept };
  };
  const strip = (res) => JSON.parse(JSON.stringify(res || {}));
  const timed = templateReport({ tdocs: LIST() });
  const plain = templateReport({ tdocs: LIST() });
  prepare(timed); prepare(plain);
  plain.s.perfTimed_ = function (label, fn) { return fn(); };
  const a1 = timed.update(), b1 = plain.update();
  const a2 = timed.update(), b2 = plain.update();
  check('with and without the timing: the same report (text, order of the tables) and the same report state', [view(timed).text === view(plain).text, view(timed).order, view(timed).props], [true, view(plain).order, view(plain).props]);
  check('... and the same results of both updates (complete, nothing left to the next update)', [strip(a1), strip(a2)], [strip(b1), strip(b2)]);
  check('... the revision below its original, "Revised to" written once, typed text kept', [timed.order().indexOf('S4aA269012') === timed.order().indexOf('S4aA269002') + 1, timed.field('S4aA269002', 'Disposition'), timed.field('S4aA269001', 'Minutes')],
    [true, 'Discussed on Tuesday — Revised to S4AA269012', 'Typed minutes of 001, kept.']);

  // The timeline of the last update: every stage named, each stage of the update once (depth 0), the nested ones marked.
  const stages = timed.exec(() => timed.s.PERF_STAGES_.slice());
  const top = stages.filter((s) => s.depth === 0).map((s) => s.stage);
  const WANTED = ['configuration/context (getReportConfig_)', 'report body (getReportBody_)', 'TDoc-list fetch/download+parse (downloadAndGroupTdocs_)',
    'meeting-room availability (applyLocalDocsAvailabilitySafely_)', 'interrupted-update check (repairAfterInterruptedUpdate_)', 'structure check (repairReportStructureIfNeeded_)',
    'existing-report/document scan (existing TDoc detection)', 'TDoc loop (new TDoc tables, statuses, upload completion)', 'registration-table links (refreshRegistrationTableLinksIfNeeded_)',
    'ad-hoc sessions (makeAdhocTdocSessionResolverSafely_)', 'registration Session column (refreshRegistrationSessionColumnSafely_)', 'title sync (syncTdocTitlesOnUpdate_)',
    'revision ordering/rearrangement (rearrangeRevisionTables_)', 'collectorUpdate_ (RSS/mail + revisions, total)', 'revision relationships (syncTdocRevisionsOnUpdate_)',
    'Goes to Plenary (syncPlenarySelectionOnUpdate_)', 'summary of the dispositions (syncDispositionSummaryOnUpdate_)', 'Undisposed Documents (syncUndisposedDocumentsOnUpdate_)', 'status dropdowns (finalizeStatusDropdowns_)'];
  check('the stages of the update on its timeline, in their order, each once (the TDoc loop, the title sync, the late stages, the status dropdowns among them)', top.filter((s) => WANTED.includes(s)), WANTED);
  check('... a stage inside another is marked as nested (depth 1 or more), not as a stage of the update', stages.filter((s) => s.depth > 0).map((s) => s.stage).filter((s) => /collector: revisions|TDoc-list: /.test(s)).length > 0 &&
    stages.filter((s) => s.depth > 0).every((s) => !WANTED.includes(s.stage)), true);
  const ordered = stages.filter((s) => s.depth === 0).slice().sort((x, y) => x.start - y.start);
  check('... the stages of the update do not overlap (each starts after the one before it ended)', ordered.every((s, i) => i === 0 || s.start >= ordered[i - 1].start + ordered[i - 1].ms), true);
  const sum = timed.exec(() => timed.s.perfStageSummary_());
  check('the summary: time in the timed stages counted once (depth 0), the total, the time outside them', [sum.timedMs === ordered.reduce((x, s) => x + s.ms, 0), sum.totalMs >= sum.timedMs, sum.untimedMs === sum.totalMs - sum.timedMs, sum.slowest.length > 0],
    [true, true, true, true]);
  check('... and one [PERF] line with it at the end of the update', timed.logs.some((l) => /\[PERF\] stage timeline: \d+ of \d+ ms in the timed stages, -?\d+ ms outside them; the slowest: /.test(String(l))), true);
}
{
  // Nested stages are not counted twice; an error in a stage leaves the depth right for the next one.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  const out = r.exec(() => {
    const s = r.s;
    s.perfResetState_(); s.perfStagesStart_();
    const busy = (ms) => { const until = Date.now() + ms; while (Date.now() < until) { /* wait */ } };
    s.perfTimed_('outer', () => { busy(15); s.perfTimed_('inner', () => busy(25)); });
    try { s.perfTimed_('failing', () => s.perfTimed_('failing inner', () => { throw new Error('boom'); })); } catch (e) { /* re-thrown, as before */ }
    s.perfTimed_('after the error', () => 0);
    s.PERF_TIMINGS_['TOTAL continuousUpdate'] = 1000000;
    return { stages: s.PERF_STAGES_.map((x) => [x.stage, x.depth]), outer: s.PERF_STAGES_.find((x) => x.stage === 'outer').ms, sum: s.perfStageSummary_() };
  });
  check('nested: the inner stage depth 1, the outer one depth 0; after an error the next stage is depth 0 again', out.stages,
    [['inner', 1], ['outer', 0], ['failing inner', 1], ['failing', 0], ['after the error', 0]]);
  check('... the time in the timed stages counts the outer stage, not the inner one again', out.sum.timedMs - out.outer < 25, true);
}
{
  // An update stopped from outside: the journal has the stages that ended (with their ends) and the one it was in
  // (without); the next update records them. The journal keeps 40 stages and stays far below a property's 9 KB.
  const r = templateReport({ tdocs: TDOCS });
  r.build();
  let inside = null;
  r.exec(() => {
    r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null);
    r.s.perfTimed_('stage one', () => 0);
    r.s.perfTimed_('stage two', () => 0);
    r.s.perfTimed_('stage three', () => { inside = JSON.parse(r.docProps.getProperty('UPDATE_RUN')); });
    // killed here: no updateJournalEnd_()
    r.s.endReportUpdate_();
  });
  const st = (inside.stages || []).map((s) => [s.stage, typeof s.ms, s.end === undefined ? 'running' : typeof s.end]);
  check('an update stopped in its third stage: the stages before it with their ends, the one it was in without', st, [['stage one', 'number', 'number'], ['stage two', 'number', 'number'], ['stage three', 'number', 'running']]);
  r.exec(() => { r.s.markReportUpdateStart_(); r.s.updateJournalStart_(null); });
  const rec = json(r, 'UPDATE_INTERRUPTED');
  check('... the next update records them (UPDATE_INTERRUPTED)', [rec.stage, rec.stages.map((s) => s.stage)], ['stage three', ['stage one', 'stage two', 'stage three']]);
  r.exec(() => {
    for (let i = 0; i < 50; i++) r.s.perfTimed_('a long stage name to fill the journal with eighty characters, number ' + String(i).padStart(4, '0') + ' ......', () => 0);
  });
  const run = prop(r, 'UPDATE_RUN');
  check('the journal keeps the last 40 stages, ends included, well below the 9 KB of a property', [JSON.parse(run).stages.length, run.length < 7000], [40, true]);
  r.exec(() => { r.s.updateJournalEnd_({ success: true }); r.s.endReportUpdate_(); });
}

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
