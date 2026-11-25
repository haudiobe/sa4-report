/**
 * Issue #28: the jobs of the runner (tools/integration-test/issue28-jobs.js): their steps, the rules a result is
 * checked against (identical, manual rows kept, only the summary changed), and the assembly of a result from the
 * description chunks of its Drive files.
 */
const J = require('../tools/integration-test/issue28-jobs.js');

let failures = 0;
function check(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) failures++;
  console.log((ok ? '  ok   ' : '  FAIL ') + name + (ok ? '' : '\n         expected ' + JSON.stringify(expected) + '\n         actual   ' + JSON.stringify(actual)));
}

console.log('jobs');
const stage = J.makeJob('main', 'stage', '20261125T120000');
check('a job: unique id, its steps, its checks pointing at real steps', [stage.id, stage.steps.map((s) => s.do).join(','), stage.expect.every((e) => e.step < stage.steps.length)],
  ['main-stage-20261125T120000', 'snap,update,snap,compare,update,snap,compare,compare', true]);
check('every job has checks pointing at its steps', Object.keys(J.JOBS).every((n) => { const j = J.makeJob('x', n, 's'); return j.expect.length && j.expect.every((e) => e.step < j.steps.length); }), true);
check('an unknown job: refused', (() => { try { J.makeJob('main', 'nothing', 's'); return 'made'; } catch (e) { return /unknown job/.test(e.message); } })(), true);
check('the plan: known stages and jobs; the main copy ends with the uninstall at T-2026.10.33; no plan configures an active meeting', [Object.values(J.PLAN).every((p) => p.every(([s, n]) => ['base', 't33', 'A', 'B', 'C'].includes(s) && J.JOBS[n])), J.PLAN.main[J.PLAN.main.length - 1], J.MAIN_CONFIG.meetingNumber], [true, ['t33', 'uninstall'], '138']);

console.log('checking a result');
const ok = (o) => Object.assign({ ok: true }, o);
const cmp = (o) => ok(Object.assign({ do: 'compare', identical: false, lines: 2, first: ['release x', 'y'], counters: { rows: {}, tdocsAdded: 0, tdocsRemoved: 0, other: { changed: 0, added: 0, removed: 0, moved: 0 }, tabs: null } }, o));
const good = { steps: [ok({ do: 'snap' }), ok({ do: 'update' }), ok({ do: 'snap' }), cmp({ counters: { rows: { 'Status / text': 3, 'Title / text': 1 }, tdocsAdded: 2, tdocsRemoved: 0, other: { changed: 0, added: 1, removed: 0, moved: 0 }, tabs: null } }),
  ok({ do: 'update' }), ok({ do: 'snap' }), cmp({ identical: true, lines: 1 }), cmp({})] };
check('a good stage: passes (Status, Title and added TDocs may change; manual rows kept; the second update changes nothing)', J.evaluate(stage, good), []);
const bad = JSON.parse(JSON.stringify(good));
bad.steps[3].counters.rows['Discussion / text'] = 1;
bad.steps[3].counters.other.removed = 2;
bad.steps[6].identical = false; bad.steps[6].lines = 3;
bad.steps[4] = { do: 'update', ok: false, error: 'The report was not updated: not configured' };
const f = J.evaluate(stage, bad);
check('manual minutes changed, text removed, a second update that changes things, an update refused: each a failure', [f.length, f.some((x) => /manual rows changed: Discussion \/ text x1/.test(x)), f.some((x) => /2 paragraph\(s\) outside the TDoc tables removed/.test(x)), f.some((x) => /expected no difference, got 3/.test(x)), f.some((x) => /step 5 \(update\): The report was not updated/.test(x))], [4, true, true, true, true]);
const h = J.makeJob('main', 'headings', 's');
const hr = { steps: h.steps.map((s) => (s.do === 'compare' ? cmp({ counters: { rows: {}, tdocsAdded: 0, tdocsRemoved: 0, other: { changed: 0, added: 0, removed: 0, moved: 0 }, summary: { changed: 3 }, tabs: null } }) : ok({ do: s.do }))) };
check('Heading 2 on and off: only the summary changed: passes', J.evaluate(h, hr), []);
hr.steps[4].counters.rows['Title / formatting'] = 1;
check('... a TDoc row changed as well: a failure', J.evaluate(h, hr).filter((x) => /outside the Disposition Summary/.test(x)).length, 1);
check('no result, or a short one: a failure', [J.evaluate(stage, null), J.evaluate(stage, { steps: [] }).length > 0], [['no result'], true]);

console.log('assembling a result from Drive file descriptions');
const text = JSON.stringify({ job: 'main-x-1', steps: [{ do: 'snap', ok: true }] });
const parts = [text.slice(0, 10), text.slice(10, 25), text.slice(25)];
const files = parts.map((p, i) => ({ name: 'IT28 result main-x-1 ' + (i + 1) + '/3.txt', description: p })).reverse().concat([{ name: 'IT28 result main-x-10 1/1.txt', description: '{}' }]);
check('the parts in any order: the result; another job with a longer id is not mixed in', J.assemble('main-x-1', files), JSON.parse(text));
check('a part missing: not yet (null)', J.assemble('main-x-1', files.filter((x) => !/ 2\/3/.test(x.name))), null);
check('no part: null', J.assemble('main-x-1', []), null);

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
