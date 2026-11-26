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
check('the plans: known stages and jobs; both end with the uninstall at T-2026.10.33', [Object.values(J.PLAN).every((p) => p.every(([s, n]) => ['base', 't33', 'A', 'B', 'C'].includes(s) && J.JOBS[n])), J.PLAN.main[J.PLAN.main.length - 1], J.PLAN.ahg[J.PLAN.ahg.length - 1]], [true, ['t33', 'uninstall'], ['t33', 'uninstall']]);
check('the meetings: main = SA4#138 Calgary (portal 60780, main, 6G Media); ahg = FS_6G_MED (portal 86178, ad-hoc, 6G Media), never SA4#138', [J.MAIN_MEETING.meetingId, J.MAIN_MEETING.expectType, J.MAIN_MEETING.reportType, J.MAIN_MEETING.overrides.meetingFolder, J.AHG_MEETING.meetingId, J.AHG_MEETING.expectType, J.AHG_MEETING.reportType, J.AHG_MEETING.overrides.agendaTdoc], [60780, 'main', '6G', 'TSGS4_138_Calgary', 86178, 'adhoc', '6G', 'S4aP260149']);
check('the AHG sessions: the two calls in the meeting agendas (22 Sep; 26-28 Oct), 15:00-18:00', J.AHG_SESSIONS.map((s) => s.date + (s.endDate ? '..' + s.endDate : '') + ' ' + s.start + '-' + s.end), ['2026-09-22 15:00-18:00', '2026-10-26..2026-10-28 15:00-18:00']);
check('the main copies: the clean one (main2) is planned; its configuration marks Status dropdowns as the report has them', [J.PLAN.main2.length, J.makeJob('main2', 'configure-main', 's').steps.map((s) => s.do)], [13, ['configureMeeting', 'markStatusDropdowns', 'props']]);
check('the Meeting Control benchmark: no Portal -- the AHG settings saved as resolved (86178, ad-hoc, 6G, agenda S4aP260149), the sessions, the Status mark as the copy is, then Meeting Control twice', (() => { const st = J.makeJob('ahg', 'mc-benchmark', 's').steps; return [st.map((s) => s.do), st[0].config.meetingId, st[0].config.meetingType, st[0].config.reportType, st[0].config.agendaTdoc, st[1].rows.length, st.some((s) => s.do === 'configureMeeting')]; })(), [['configure', 'saveSessions', 'markStatusDropdowns', 'props', 'meetingControl', 'meetingControl'], '86178', 'adhoc', '6G', 'S4aP260149', 2, false]);
check('the steps that write settings stop the job when they fail (configure, markStatusDropdowns among them)', (() => { const src = require('fs').readFileSync(require('path').join(__dirname, '..', 'tools', 'integration-test', 'Issue28Check.js'), 'utf8'); const l = src.match(/var IT28_CHANGING_STEPS_ = (\[[^\]]*\])/)[1]; return ['configure', 'configureMeeting', 'markStatusDropdowns', 'setProps', 'saveSessions'].every((s) => l.includes("'" + s + "'")); })(), true);
check('main2: the settled T-2026.10.33 snapshot of run 1 (s5) checked unchanged, then marked; its rollback retries a killed update once', [J.PLAN.main2[3][1], J.makeJob('main2', 'settled-check', 's').steps.map((s) => s.do + (s.from ? ' ' + s.from : '')), J.PLAN.main2[J.PLAN.main2.length - 2]], ['settled-check', ['snap', 'compare s5', 'markSettled'], ['t33', 'rollback']]);
check('the plans settle the release the copy runs first; the stages compare with that settled snapshot', [['base', 'settle'], J.PLAN.ahg[3], J.makeJob('main2', 'settle', 's').steps.map((s) => s.do + (s.from ? ' ' + s.from : '')), J.makeJob('main2', 'stage', 's').steps[7].from], [['base', 'settle'], ['t33', 'settle'], ['snap', 'update', 'snap', 'compare previous', 'update', 'snap', 'compare previous', 'markSettled'], 'settled']);
{
  const job = J.makeJob('main2', 'settle', 's');
  const okStep = { ok: true };
  const res = (rows, extra) => ({ steps: [okStep, okStep, okStep, Object.assign({ ok: true, counters: Object.assign({ rows: rows, tdocsAdded: 2, tdocsRemoved: 0, other: { removed: 0 } }, extra || {}) }), okStep, okStep, { ok: true, identical: true }, okStep] });
  check('settle: the release reformatting a Minutes paragraph in its first update is shown, not failed', J.evaluate(job, res({ 'Minutes / formatting': 2, 'E-mail Discussion / text (with its styles)': 54 })), []);
  check('settle: typed Minutes text changed, or a TDoc table removed: failed', [J.evaluate(job, res({ 'Minutes / text (with its styles)': 1 })).length, J.evaluate(job, res({}, { tdocsRemoved: 1 })).length], [1, 1]);
  check('settle: a second update that still changes something: failed', J.evaluate(job, Object.assign(res({}), { steps: res({}).steps.map((s, i) => i === 6 ? { ok: true, identical: false, lines: 3, first: ['x'] } : s) })).length, 1);
}
{
  // The candidate stages and the rollback: no field of a typed or decision row, no TDoc table, no link lost.
  const job = J.makeJob('main2', 'stage', 's');
  const okStep = { ok: true };
  const res = (c3) => ({ steps: [okStep, okStep, okStep, Object.assign({ ok: true }, { counters: Object.assign({ rows: {}, tdocsAdded: 0, tdocsRemoved: 0, other: { removed: 0 }, counts: { links: '666 -> 666' } }, c3) }), okStep, okStep, { ok: true, identical: true }, { ok: true, counters: { rows: {}, tdocsAdded: 0, tdocsRemoved: 0, other: { removed: 0 }, counts: { links: '666 -> 666' } } }] });
  check('stage: nothing protected touched: passes', J.evaluate(job, res({})), []);
  check('stage: a Minutes paragraph setting, a Status, a Disposition format, a lost link: each fails (stop and clean up)', [
    J.evaluate(job, res({ rows: { 'Minutes / formatting': 1 } })).length, J.evaluate(job, res({ rows: { 'Status / text': 1 } })).length,
    J.evaluate(job, res({ rows: { 'Disposition / formatting': 1 } })).length, J.evaluate(job, res({ counts: { links: '840 -> 666' } })).length,
    J.evaluate(job, res({ rows: { 'Title / text': 1 }, tdocsAdded: 1, counts: { links: '666 -> 670' } })).length], [1, 1, 1, 1, 0]);
  check('the rollback has the same strict rule; phase B keeps its minutes and its links', [J.makeJob('main2', 'rollback', 's').expect.filter((e) => e.rule === 'protectedKept').length, J.makeJob('main2', 'phase-b', 's').expect.filter((e) => e.rule === 'linksKept').length], [2, 2]);
  const bd = J.makeJob('main2', 'baseline-diff', 's');
  check('baseline-diff: reads only; untouched -> before the first update the same, the second update nothing, untouched -> settled every line', [bd.steps.map((s) => s.do + ' ' + s.from + '>' + s.to + (s.all ? ' all' : '')), J.PLAN.main2[4]], [['compare s1>s3', 'compare s3>s4', 'compare s4>s5', 'compare s1>s5 all'], ['base', 'baseline-diff']]);
}
check('the static AHG settings are never SA4#138 (no meeting folder or number)', [J.AHG_CONFIG.meetingFolder, J.AHG_CONFIG.meetingNumber], ['', '']);
check('the configure jobs use exactly those meetings; every job step is of the fixed list', [J.makeJob('main', 'configure-main', 's').steps[0].meetingId, J.makeJob('ahg', 'configure-ahg', 's').steps[0].meetingId, J.makeJob('ahg', 'configure-ahg', 's').steps[1].rows.length, Object.keys(J.JOBS).every((n) => J.makeJob('x', n, 's').steps.every((st) => ['probe','props','tableHeaders','markSettled','snap','compare','update','configure','configureMeeting','markStatusDropdowns','saveSessions','setProps','candidates','expectRevision','resolveRevision','closeReportCheck','meetingControl','sessionRulePreview','sessionRuleSet','carryPreview','carrySet','uninstall'].includes(st.do)))], [60780, 86178, 2, true]);

console.log('checking a result');
const ok = (o) => Object.assign({ ok: true }, o);
const cmp = (o) => ok(Object.assign({ do: 'compare', identical: false, lines: 2, first: ['release x', 'y'], counters: { rows: {}, tdocsAdded: 0, tdocsRemoved: 0, other: { changed: 0, added: 0, removed: 0, moved: 0 }, tabs: null } }, o));
const good = { steps: [ok({ do: 'snap' }), ok({ do: 'update' }), ok({ do: 'snap' }), cmp({ counters: { rows: { 'Title / text': 1 }, tdocsAdded: 2, tdocsRemoved: 0, other: { changed: 0, added: 1, removed: 0, moved: 0 }, tabs: null } }),
  ok({ do: 'update' }), ok({ do: 'snap' }), cmp({ identical: true, lines: 1 }), cmp({})] };
check('a good stage: passes (a Title and added TDocs may change; typed and decision rows kept; the second update changes nothing)', J.evaluate(stage, good), []);
const bad = JSON.parse(JSON.stringify(good));
bad.steps[3].counters.rows['Discussion / text'] = 1;
bad.steps[3].counters.other.removed = 2;
bad.steps[6].identical = false; bad.steps[6].lines = 3;
bad.steps[4] = { do: 'update', ok: false, error: 'The report was not updated: not configured' };
const f = J.evaluate(stage, bad);
check('manual minutes changed, text removed, a second update that changes things, an update refused: each a failure', [f.length, f.some((x) => /protected rows changed: Discussion \/ text x1/.test(x)), f.some((x) => /2 paragraph\(s\) outside the TDoc tables removed/.test(x)), f.some((x) => /expected no difference, got 3/.test(x)), f.some((x) => /step 5 \(update\): The report was not updated/.test(x))], [4, true, true, true, true]);
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
