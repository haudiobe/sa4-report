/**
 * BOUNDED UPDATE (after T-2026.10.23; GitHub issue #9, integration test U5).
 *
 * An ad-hoc update of the SA4#138 copy had 280 new TDocs (392 in the list,
 * 112 in the report) and was stopped by Apps Script after 6 minutes
 * ("Exceeded maximum execution time"); nothing it did was saved. This suite
 * runs the REAL update (tests/helpers/template-report.js) with a fake clock
 * that advances with every insertion, so the time limit is reached
 * deterministically:
 *
 *   A. the pure rules;
 *   B. a small update completes as before;
 *   C. 392 TDocs, 280 new: the update stops starting TDocs after the budget,
 *      leaves the rest, says so, saves in parts (checkpoints);
 *   D. the next run goes on: no TDoc twice, every TDoc once;
 *   E. a run after that changes nothing (idempotent);
 *   F. hand-written text survives; G. status dropdowns and their values;
 *   H. an execution stopped after a checkpoint: what was saved stays, the
 *      next run adds the rest and the missing registration rows;
 *   I. optional steps are left over when time is short;
 *   J. the background path: a paused run leaves lastRunAt, the scheduler's
 *      start is the clock's start;
 *   K. the meeting-number check and the time zones (regressions);
 *   L. the ledger: without the section, Code.js and ReportCreator.js are
 *      those of 8edb1a7.
 *
 * Run: node tests/bounded-update.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { templateReport } = require('./helpers/template-report.js');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { codeWithoutBoundedUpdate, creatorWithoutBoundedUpdate, BOUNDED_UPDATE_SECTION_START, BOUNDED_UPDATE_SECTION_END, BOUNDED_UPDATE_SECTION_FUNCTIONS,
  BOUNDED_UPDATE_CHANGED_FUNCTIONS, BOUNDED_UPDATE_CREATOR_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`);
  else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const ITEMS = ['4.3', '4.4', '4.5'];
const STATUSES = ['agreed', 'noted', 'available'];
const id = (n) => 'S4aA26' + String(1000 + n);
const tdoc = (n) => ({ id: id(n), agenda: ITEMS[n % 3], status: STATUSES[n % 3], uploaded: true });

/** A report with `existing` TDocs built, the list then grown to `total`; a clock that advances `msPerInsert` per new TDoc. */
function report(existing, total, msPerInsert, props) {
  const first = [];
  for (let n = 0; n < existing; n++) first.push(tdoc(n));
  const r = templateReport({ tdocs: first, tableSeparators: true, props: props || {} });
  vm.runInContext(fs.readFileSync(REPORT_CREATOR_PATH, 'utf8'), r.s, { filename: 'ReportCreator.js' });
  r.build();
  for (let n = existing; n < total; n++) r.tdocs.push(tdoc(n));
  const clock = { now: 0, step: msPerInsert };
  r.clock = clock;
  r.s.updateNow_ = () => clock.now;
  const realInsert = r.s.insertNewTdoc_;
  r.inserts = 0;
  r.killAt = 0;
  r.s.insertNewTdoc_ = function () {
    r.inserts++;
    if (r.killAt && r.inserts === r.killAt) throw new Error('SIMULATED: Exceeded maximum execution time');
    clock.now += clock.step;
    return realInsert.apply(this, arguments);
  };
  // Every save: what the document is then (to simulate an execution whose unsaved changes are dropped).
  const doc = r.s.DocumentApp.getActiveDocument();
  const realSave = doc.saveAndClose;
  r.saved = r.body._children.slice();
  doc.saveAndClose = () => { r.saved = r.body._children.slice(); return realSave(); };
  r.run = () => { clock.now = 0; r.inserts = 0; r.mark(); return r.update(); };
  r.registered = () => {
    const t = r.body._children.filter((c) => c.getType() === 'TABLE' && r.s.registrationTableColumns_(c))[0];
    const out = [];
    for (let i = 1; i < t.getNumRows(); i++) out.push(t.getRow(i).getCell(0).getText());
    return out;
  };
  r.snapshot = () => r.body._children.map((c) => {
    if (c.getType() !== 'TABLE') return 'P:' + (c.getText ? c.getText() : '');
    const rows = [];
    for (let i = 0; i < c.getNumRows(); i++) { const row = c.getRow(i); const cells = []; for (let j = 0; j < row.getNumCells(); j++) cells.push(row.getCell(j).getText()); rows.push(cells.join('|')); }
    return 'T:' + rows.join('/');
  }).join('\n');
  return r;
}
const counts = (list) => list.reduce((m, x) => (m[x] = (m[x] || 0) + 1, m), {});
const dupes = (list) => Object.keys(counts(list)).filter((k) => counts(list)[k] > 1);

// ====================================================================== A
console.log('A. the rules (pure)');
{
  const S = loadCode().sandbox;
  check('the budgets: no new TDoc after 3 minutes, no optional step after 4 (Apps Script stops at 6)', [S.UPDATE_INSERT_BUDGET_MS_, S.UPDATE_OPTIONAL_STEPS_BUDGET_MS_], [180000, 240000]);
  let now = 1000;
  S.updateNow_ = () => now;
  const clock = S.updateClock_();
  const pause = S.newUpdatePause_();
  now += 179999;
  check('a TDoc may start just before the budget', S.updateMayInsert_(clock, pause), true);
  now += 1;
  check('...not at it', S.updateMayInsert_(clock, pause), false);
  S.noteUpdateTdocLeft_(pause, 'S4aA261001'); S.noteUpdateTdocLeft_(pause, 'S4aA261001'); S.noteUpdateTdocLeft_(pause, 'S4aA261002');
  check('a TDoc the list has twice is left over once', pause.newTdocsLeft, 2);
  now = 1000;
  check('once one is left over, none starts, whatever the time', S.updateMayInsert_(clock, pause), false);
  check('nothing left: no pause', S.updatePauseResult_(S.newUpdatePause_(), 3), null);
  S.UPDATE_EXECUTION_STARTED_AT_ = 5;
  check('the clock starts at the start of the execution when it is known (the scheduler)', S.updateClock_().startedAt, 5);
  S.UPDATE_EXECUTION_STARTED_AT_ = 0;
  check('the message: in part, how many added and left, what did not run, run again',
    S.describeUpdatePause_({ newTdocsAdded: 180, newTdocsLeft: 100, stepsLeft: ['e-mail discussions and revisions'] }),
    'The report was updated in part: the update stopped in time before the Apps Script limit of 6 minutes. 180 new TDoc(s) were added; 100 new TDoc(s) are still to be added; e-mail discussions and revisions did not run.\nNothing was lost or added twice. Run Update Report Now again to go on.');
}

// ====================================================================== B
console.log('\nB. a small update completes as before');
{
  const r = report(6, 9, 1000);
  const u = r.run();
  check('success, nothing left over, three tables more', [u.success, u.paused === undefined, r.tdocTables().length], [true, true, 9]);
  check('Update Report Now shows nothing', r.s.describeUpdateReportNowFailure_(u), '');
  check('no checkpoint (a small update saves once, in its status-dropdown step)', r.logs.filter((l) => /\[CHECKPOINT\]/.test(l)).length, 0);
  check('the e-mail/revisions step ran', r.logs.filter((l) => /left to the next run/.test(l)), []);
}

// ====================================================================== C-G
console.log('\nC. 392 TDocs, 280 new (the U5 run), 1 s per insertion');
const big = report(112, 392, 1000, { FETCH_ABSTRACTS_ON_UPDATE: 'true' });
let collectorRuns = 0;
big.s.collectorUpdate_ = () => { collectorRuns++; return { failures: [] }; };
// F: hand-written text, in the document and in a TDoc table, before the update.
big.user(() => {
  big.body.appendParagraph('IT9 opening probe -- written by hand');
  const t = big.table(id(4));
  for (let i = 0; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === 'Minutes') t.getRow(i).getCell(1).setText('Hand-written minute of ' + id(4));
});
const shownBefore = [0, 1, 2, 3].map((n) => big.shown(id(n)));
const keysBefore = big.docProps.getKeys ? big.docProps.getKeys().slice().sort() : null;
const u1 = big.run();
{
  check('the run succeeds and is paused: 180 added, 100 left, the optional steps and the upload check left over', [u1.success, u1.paused], [true, {
    newTdocsAdded: 180, newTdocsLeft: 100, stepsLeft: ['the upload check (link and abstract) of the TDocs already in the report', 'the abstract sweep', 'e-mail discussions and revisions'] }]);
  check('no TDoc was started after the budget: 180 insertions, 180 Reviewer requests (one per inserted, uploaded TDoc)', [big.inserts, big.requests.length], [180, 180]);
  check('the report has 112 + 180 TDoc tables, none twice', [big.tdocTables().length, dupes(big.order())], [292, []]);
  check('the registration table: exactly the TDocs whose tables the report has', big.registered().slice().sort(), big.order().slice().sort());
  check('the collector did not run', collectorRuns, 0);
  check('it saved in parts: checkpoints before new TDocs, about every 2000 changes', big.logs.filter((l) => /^\[CHECKPOINT\] update: before /.test(l)).length, 4);
  check('Update Report Now says so, as an error the user sees',
    /^The report was updated in part: .* 180 new TDoc\(s\) were added; 100 new TDoc\(s\) are still to be added; .*Run Update Report Now again to go on\.$/s.test(big.s.describeUpdateReportNowFailure_(u1)), true);
  check('nothing records progress apart from the document (no new document property)', big.docProps.getKeys ? big.docProps.getKeys().filter((k) => keysBefore.indexOf(k) === -1 && !/^REVIEWER_NO_SUMMARY_CACHE_/.test(k)) : [], []);
  check('the checkpoints are off after the run', big.s.LARGE_BUILD_CHECKPOINTS_.enabled, false);
}

console.log('\nD. the next run goes on');
const u2 = big.run();
{
  check('it completes: 100 added, nothing left', [u2.success, u2.paused === undefined, big.inserts], [true, true, 100]);
  check('all 392 TDocs, each exactly once, in the report and in the registration table', [big.tdocTables().length, dupes(big.order()), big.registered().length, dupes(big.registered())], [392, [], 392, []]);
  check('every TDoc of the list is in the report', big.tdocs.map((t) => t.id).filter((x) => big.order().indexOf(x) === -1), []);
  check('the optional steps ran again', collectorRuns, 1);
}

console.log('\nE. a run after that changes nothing');
{
  const before = big.snapshot();
  const u3 = big.run();
  check('success, nothing inserted, no Reviewer request for an inserted TDoc', [u3.success, u3.paused === undefined, big.inserts], [true, true, 0]);
  check('the document is as it was', big.snapshot() === before, true);
}

console.log('\nF. hand-written text survives');
check('the paragraph and the minute written by hand are there, once', [big.body._children.filter((c) => c.getType() !== 'TABLE' && c.getText() === 'IT9 opening probe -- written by hand').length, big.field(id(4), 'Minutes')],
  [1, 'Hand-written minute of ' + id(4)]);

console.log('\nG. status dropdowns');
check('the dropdowns of the TDocs that were there keep their values', [0, 1, 2, 3].map((n) => big.shown(id(n))), shownBefore);
check('the TDocs of both runs have their dropdowns with the list\'s status', [112, 113, 114, 291, 292, 391].map((n) => big.shown(id(n))),
  [112, 113, 114, 291, 292, 391].map((n) => 'dropdown:' + STATUSES[n % 3]));

// ====================================================================== H
console.log('\nH. an execution stopped after a checkpoint (its unsaved changes dropped)');
{
  const r = report(112, 392, 0);
  r.killAt = 60;
  const u = r.run();
  // What Apps Script does to an execution it stops: the changes after the last save are gone.
  r.body._children.splice(0, r.body._children.length, ...r.saved);
  check('the run failed', u.success, false);
  const kept = r.tdocTables().length - 112;
  check('what the checkpoint saved is in the report (39 TDocs)', kept, 39);
  check('...but has no registration rows yet', r.registered().length, 112);
  r.killAt = 0;
  const v = r.run();
  check('the next run completes: it adds only the 241 missing TDocs', [v.success, v.paused === undefined, r.inserts], [true, true, 241]);
  check('392 tables, none twice; every one has its registration row, once', [r.tdocTables().length, dupes(r.order()), r.registered().slice().sort(), dupes(r.registered())], [392, [], r.order().slice().sort(), []]);
}

// ====================================================================== I
console.log('\nI. optional steps when time is short');
{
  const r = report(9, 9, 0, { FETCH_ABSTRACTS_ON_UPDATE: 'true' });
  let collector = 0;
  r.s.collectorUpdate_ = () => { collector++; return { failures: [] }; };
  const realStatus = r.s.updateTdocStatus_;
  r.s.updateTdocStatus_ = function () { r.clock.now = 300000; return realStatus.apply(this, arguments); };
  const u = r.run();
  check('no new TDoc, but 5 minutes used: the abstract sweep and the collector are left to the next run', [u.success, u.paused, collector],
    [true, { newTdocsAdded: 0, newTdocsLeft: 0, stepsLeft: ['the abstract sweep', 'e-mail discussions and revisions'] }, 0]);
  r.s.updateTdocStatus_ = realStatus;
  const v = r.run();
  check('with time, they run', [v.paused === undefined, collector], [true, 1]);
}

// ====================================================================== J
console.log('\nJ. the background path');
{
  const S = loadCode().sandbox;
  let advanced = 0;
  S.getRegisteredReportDocument_ = () => ({ documentId: 'D1', enabled: true, lastRunAt: '2026-11-01T00:00:00.000Z' });
  S.updateRegisteredReportDocument_ = (docId, patch) => { advanced++; return Object.assign({ documentId: docId }, patch); };
  S.continuousUpdateCore_ = () => ({ success: true, error: null, paused: { newTdocsAdded: 5, newTdocsLeft: 7, stepsLeft: [] } });
  const p = S.continuousUpdateForDocument_('D1');
  check('a paused run does not advance lastRunAt: the report stays due', [p.success, p.paused.newTdocsLeft, p.lastRunAt, advanced], [true, 7, '2026-11-01T00:00:00.000Z', 0]);
  S.continuousUpdateCore_ = () => ({ success: true, error: null });
  const c = S.continuousUpdateForDocument_('D1');
  check('a complete run does, as before', [c.success, advanced, c.paused], [true, 1, undefined]);
  // The scheduler: the clock of each update starts at the start of the execution, and is reset after it.
  const seen = [];
  S.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  S.listRegisteredReportDocuments_ = () => [{ documentId: 'D1', enabled: true, lastRunAt: null }, { documentId: 'D2', enabled: true, lastRunAt: null }];
  S.continuousUpdateForDocument_ = (d) => { seen.push(S.UPDATE_EXECUTION_STARTED_AT_ > 0); return { durationMs: 1 }; };
  S.runAddonScheduler_();
  check('the scheduler sets the start of the execution for every update, and resets it', [seen, S.UPDATE_EXECUTION_STARTED_AT_], [[true, true], 0]);
  const SRC = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  check('checkpoints only with the report open, never in a background run (source)', /\n  if \(!context\) beginLargeBuildCheckpoints_\(\);\n/.test(SRC), true);
}

// ====================================================================== K
console.log('\nK. regressions: the meeting number, the time zones');
{
  const S = loadCode().sandbox;
  const store = (p) => ({ getProperty: (k) => (p[k] === undefined ? null : p[k]) });
  const calgary = 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/';
  check('a saved number 136 under the SA4#138 folder is still refused before anything', /The saved meeting number 136 is not the number of the saved document folder/.test(S.mainMeetingNumberMismatch_(store({ MEETING_TYPE: 'main', FTP_BASE: calgary, MEETING_NUMBER: '136' }))), true);
  check('...and 138 is accepted', S.mainMeetingNumberMismatch_(store({ MEETING_TYPE: 'main', FTP_BASE: calgary, MEETING_NUMBER: '138' })), '');
  check('Alberta from 1 November 2026: ABT (UTC-6); before it MDT', [S.meetingTimeZoneAbbreviationAt_('America/Edmonton', Date.UTC(2026, 10, 16, 15)), S.meetingTimeZoneOffsetMinutes_('America/Edmonton', Date.UTC(2026, 10, 16, 15)),
    S.meetingTimeZoneAbbreviationAt_('America/Edmonton', Date.UTC(2026, 9, 30, 15))], ['ABT', -360, 'MDT']);
  check('a standard IANA zone: Berlin CEST in August, CET in January', [S.meetingTimeZoneAbbreviationAt_('Europe/Berlin', Date.UTC(2026, 7, 24, 7)), S.meetingTimeZoneAbbreviationAt_('Europe/Berlin', Date.UTC(2026, 0, 20, 9))], ['CEST', 'CET']);
}

// ====================================================================== L
console.log('\nL. the ledger');
{
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutBoundedUpdate(FULL);
  const cview = creatorWithoutBoundedUpdate(CREATOR);
  check('every change is in Code.js and ReportCreator.js exactly once', view.problems.concat(cview.problems), []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((x) => x.replace(/^function\s+|\s*\($/g, '')).filter((n, i, all) => all.indexOf(n) === i);
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(BOUNDED_UPDATE_SECTION_START), FULL.indexOf(BOUNDED_UPDATE_SECTION_END) + BOUNDED_UPDATE_SECTION_END.length);
  check('the section: these functions', functionNames(section), BOUNDED_UPDATE_SECTION_FUNCTIONS);
  check('it stands directly after the main-meeting number', FULL.indexOf('// END OF MAIN MEETING NUMBER\n// =========================================================\n\n' + BOUNDED_UPDATE_SECTION_START) !== -1, true);
  check('of the functions there were, these changed; none removed', [functionNames(view.code).filter((n) => functionSource(FULL, n) !== functionSource(view.code, n)).sort(), functionNames(view.code).filter((n) => functionNames(FULL).indexOf(n) === -1)],
    [BOUNDED_UPDATE_CHANGED_FUNCTIONS.slice().sort(), []]);
  check('of ReportCreator.js, only the message of Update Report Now', functionNames(cview.code).filter((n) => functionSource(CREATOR, n) !== functionSource(cview.code, n)), BOUNDED_UPDATE_CREATOR_CHANGED_FUNCTIONS);
  check('Code.js is still 2.26.8 (T-2026.10.24 not released yet)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.8');
  let base = null, baseCreator = null;
  try {
    base = execFileSync('git', ['show', '8edb1a7:Code.js'], { cwd: path.join(__dirname, '..'), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
    baseCreator = execFileSync('git', ['show', '8edb1a7:template/ReportCreator.js'], { cwd: path.join(__dirname, '..'), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) { base = null; }
  if (base === null) console.log('  note: commit 8edb1a7 is not in this checkout; the comparison is skipped.');
  else check('without it, Code.js and ReportCreator.js are those of 8edb1a7 (the tested candidate), byte for byte', [view.code === base, cview.code === baseCreator], [true, true]);
}

console.log(failures ? `\n${failures} FAILED` : '\nall passed');
process.exit(failures ? 1 : 0);
