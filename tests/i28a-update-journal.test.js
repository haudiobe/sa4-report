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
  check('the identical copy removed, the interruption named and closed (the update completed)', [count(r, 'S4aA269002'), res.interrupted.removed, res.interrupted.differing, prop(r, 'UPDATE_INTERRUPTED')], [1, ['S4aA269002'], [], null]);
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

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
