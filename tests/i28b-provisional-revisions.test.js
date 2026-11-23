/**
 * Issue #28, phase B (docs/ISSUE28_PHASE_B.md): a revision's decision recorded before 3GU registers it.
 *
 *   A. validateProvisionalInput_(): what can be recorded; a local revision file (S4-261234r01) is no successor.
 *   B. decideProvisionalTransfer_(): wait, transfer, already shown, a chain, a cycle, two successors, another than
 *      expected, a successor decided by hand (pure).
 *   C. an MBS report of SA4#138 with Status dropdowns: recorded (the original's Status unchanged), waiting (Close
 *      Report blocks), transferred exactly once when 3GU registers the revision, the original "revised"; a later
 *      manual change of the revision is never overwritten.
 *   D. conflicts: decided by hand, two successors, another than expected -- nothing written.
 *   E. Goes to Plenary: the choice and the agenda item go with the decision into a new revision's row; an existing
 *      Plenary choice is never changed (to be set by hand, Close Report says so).
 *   F. cancel; undo (by the next update, only while the revision still shows the decision).
 *   G. no record: the update reads and writes nothing more (DocumentApp calls against phase A).
 *   H. the ledger: without this layer Code.js and ReportCreator.js are those of phase A, byte for byte.
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');
const ledger = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) console.log(`  ok   ${name}`);
  else { failures++; console.log(`  FAIL ${name}\n         expected ${e}\n         actual   ${a}`); }
}

const ROOT = path.join(__dirname, '..');
const RAW = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const AGENDA_138 = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'fixtures', 'agenda-S4-261857.google-items.json'), 'utf8')).items
  .map((i) => ({ number: i.number, title: i.title, level: i.level, heading: 'NORMAL', text: '' }));
const MAIN = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_NAME: 'SA4#138', MEETING_ID: '60780', REPORT_SUFFIX: 'MBS', MEETING_DATE: 'November 16, 2026',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx',
  AGENDA_TDOC: 'S4-261857' };
const YES = '☑ Yes';
const NO = '☐ No';
const TDOCS = () => [
  { id: 'S4-265001', agenda: '8.5', status: 'available', title: '[AMD_ARCH_Ph2-MED] CR on multi-access', uploaded: true },
  { id: 'S4-265002', agenda: '8.6', status: 'available', title: '[Energy_PRO-MED] pCR', uploaded: true },
  { id: 'S4-265003', agenda: '8.4', status: 'available', title: 'CR 26.511 correction', type: 'CR', uploaded: true },
  { id: 'S4-265004', agenda: '8.7', status: 'agreed', title: 'QUIC results', uploaded: true }
];

function report(o) {
  const r = templateReport(Object.assign({ props: MAIN, agenda: AGENDA_138, tdocs: TDOCS(), columns: ['Release', 'Related WIs'], tableSeparators: true, endParagraph: true }, o || {}));
  r.s.HtmlService = { createHtmlOutput: (html) => { const out = { setWidth: () => out, setHeight: () => out, getContent: () => html }; return out; } };
  r.ui.dialogs = [];
  r.ui.showModalDialog = (output, title) => { r.ui.dialogs.push([title, output.getContent()]); };
  const up = r.update;
  r.update = function () { const out = up.apply(this, arguments); r.body._state.dirty = false; r.body._state.closed = false; return out; };
  r.save = () => { r.body._state.dirty = false; r.body._state.closed = false; };
  r.record = (input) => { const out = r.exec(() => r.s.saveExpectedRevision(input)); r.save(); return out; };
  r.change = (action, original, details) => { const out = r.exec(() => r.s.changeExpectedRevision(action, original, details)); r.save(); return out; };
  r.rec = (id) => JSON.parse(r.docProps.getProperty('PROVREV_' + id) || 'null');
  /** 3GU registers `id` as the revision of `of` (both directions of the list say it). */
  r.register = (id, of, extra) => { r.tdocs.push(Object.assign({ id: id, agenda: r.tdoc(of).agenda, status: 'reserved', title: r.tdoc(of).title, isRevisionOf: of, uploaded: false }, extra || {})); r.tdoc(of).revisedTo = id; r.tdoc(of).status = 'revised'; };
  r.closeProblems = () => { const res = r.exec(() => r.s.closeReportCore_()); return res.problems.filter((p) => /revision is expected|could not be transferred|set it by hand|undoing|by the last update/.test(p)); };
  return r;
}

// ====================================================================== A
console.log('A. what can be recorded');
{
  const { sandbox: S } = loadCode({ documentProperties: {} });
  const v = (o) => { const x = S.validateProvisionalInput_(Object.assign({ original: 'S4-265001', decision: 'agreed', plenary: 'yes', agenda: '15.4', successor: '' }, o)); return x.ok ? 'ok' : x.error; };
  check('a decision of the Status dropdown, Plenary Yes/No/blank, an agenda item, an optional expected TDoc', [v({}), v({ plenary: '' }), v({ decision: 'Noted' }), v({ successor: 'S4-265101' })], ['ok', 'ok', 'ok', 'ok']);
  check('not a decision for a revision: revised, available, withdrawn, a typo', ['revised', 'available', 'withdrawn', 'agred'].map((d) => /must be one of/.test(v({ decision: d }))), [true, true, true, true]);
  check('a local revision file is no successor: S4-265001r01, S4-265001_r1, S4-265001-r2.docx', ['S4-265001r01', 'S4-265001_r1', 'S4-265001-r2.docx'].map((s) => /is not a TDoc number/.test(v({ successor: s }))), [true, true, true]);
  check('the original must be a TDoc number; the successor cannot be the original; A.I. is a number', [/not a TDoc number/.test(v({ original: 'S4-26500' })), /cannot be the original/.test(v({ successor: 'S4-265001' })), /agenda item must be a number/.test(v({ agenda: 'AI 15' }))], [true, true, true]);
}

// ====================================================================== B
console.log('B. decideProvisionalTransfer_()');
{
  const { sandbox: S } = loadCode({ documentProperties: {} });
  const facts = (o) => { const out = {}; Object.keys(o).forEach((k) => { out[k] = Object.assign({ revisedTo: [], isRevisionOf: [] }, o[k]); }); return out; };
  const rec = (o) => Object.assign({ v: 1, original: 'A', decision: 'agreed', plenary: '', agenda: '', successor: '', state: 'pending' }, o || {});
  const st = (m) => (n) => (n in m ? { exists: true, label: m[n] } : { exists: false, label: '' });
  const d = (r, f, s) => { const x = S.decideProvisionalTransfer_(r, f, s || st({})); return [x.action, x.successor || '']; };
  check('no successor: wait; the list names one not registered yet: wait, said', [d(rec(), facts({ A: {} })), S.decideProvisionalTransfer_(rec(), facts({ A: { revisedTo: ['B'] } }), st({})).reason], [['wait', ''], 'the TDoc list names B as the revision, not registered yet']);
  check('"Revised to" or "Is revision of" alone each make the successor (registered, undecided): transfer', [d(rec(), facts({ A: { revisedTo: ['B'] }, B: {} }), st({ B: 'reserved' })), d(rec(), facts({ A: {}, B: { isRevisionOf: ['A'] } }), st({ B: 'available' }))], [['transfer', 'B'], ['transfer', 'B']]);
  check('registered but no table in the report yet: wait', d(rec(), facts({ A: { revisedTo: ['B'] }, B: {} }), st({})), ['wait', 'B']);
  check('the revision already shows the decision: done (nothing written)', d(rec(), facts({ A: { revisedTo: ['B'] }, B: {} }), st({ B: 'agreed' })), ['done', 'B']);
  check('the revision decided by hand otherwise: conflict, never written', d(rec(), facts({ A: { revisedTo: ['B'] }, B: {} }), st({ B: 'noted' })), ['conflict', 'B']);
  check('a chain A -> B -> C: the newest registered revision', d(rec(), facts({ A: { revisedTo: ['B'] }, B: { revisedTo: ['C'] }, C: {} }), st({ B: 'revised', C: 'reserved' })), ['transfer', 'C']);
  check('two successors: conflict, never a guess', d(rec(), facts({ A: { revisedTo: ['B'] }, B: {}, C: { isRevisionOf: ['A'] } }), st({ B: 'reserved', C: 'reserved' })), ['conflict', '']);
  check('another than expected: conflict; the expected one in the chain: fine', [d(rec({ successor: 'X' }), facts({ A: { revisedTo: ['B'] }, B: {} }), st({ B: 'reserved' })), d(rec({ successor: 'B' }), facts({ A: { revisedTo: ['B'] }, B: { revisedTo: ['C'] }, C: {} }), st({ B: 'revised', C: 'reserved' }))],
    [['conflict', 'B'], ['transfer', 'C']]);
  check('a cycle in the list: conflict', d(rec(), facts({ A: { revisedTo: ['B'] }, B: { revisedTo: ['A'] } }), st({ B: 'reserved' }))[0], 'conflict');
  check('the original not in the list: wait', d(rec(), facts({})), ['wait', '']);
}

// ====================================================================== C
console.log('C. recorded, waiting, transferred exactly once');
{
  const r = report();
  r.build(); r.save(); r.update();
  const said = r.record({ original: 'S4-265001', decision: 'agreed', plenary: '', agenda: '', successor: '' });
  check('recorded; the original keeps its Status (no "agreed" forced on it)', [/Recorded: the revision of S4-265001 is to be "agreed"/.test(said), r.shown('S4-265001'), r.rec('S4-265001').state], [true, 'dropdown:available', 'pending']);
  r.update();
  check('an update without a registered revision: waits, says why; the original as it was', [r.rec('S4-265001').state, r.rec('S4-265001').reason, r.shown('S4-265001')], ['pending', 'no revision is registered in 3GU yet', 'dropdown:available']);
  check('Close Report blocks while it waits, and says what to do', r.closeProblems().length === 1 && /S4-265001: a revision is expected \("agreed"\)/.test(r.closeProblems()[0]), true);
  r.register('S4-265101', 'S4-265001');
  const u = r.update();
  check('3GU registers S4-265101 as its revision: the decision written into it, the original revised; "transferring" until confirmed',
    [r.shown('S4-265101'), r.shown('S4-265001'), r.rec('S4-265001').state, r.rec('S4-265001').to, u.revisionsExpected], ['dropdown:agreed', 'dropdown:revised', 'transferring', 'S4-265101', '1 written, 0 transferred, 0 waiting for 3GU, 0 conflict(s)']);
  check('Close Report asks for the confirming update', r.closeProblems().some((p) => /written to S4-265101 by the last update/.test(p)), true);
  const u2 = r.update();
  check('the next update sees the revision show it: transferred', [r.rec('S4-265001').state, u2.revisionsExpected], ['transferred', '1 transferred, 0 waiting for 3GU, 0 conflict(s)']);
  check('Close Report no longer lists it', r.closeProblems(), []);
  r.user(() => r.api.pick('S4-265101', 'noted'));
  r.update(); r.update();
  check('exactly once: a later change of the revision by hand stays', [r.shown('S4-265101'), r.rec('S4-265001').state], ['dropdown:noted', 'transferred']);
  const log = JSON.parse(r.docProps.getProperty('PROVREV_LOG'));
  check('the log: recorded, written, transferred', log.map((e) => e.a).reverse().slice(0, 3), ['recorded', 'written', 'transferred']);
  r.exec(() => r.s.manageExpectedRevisions());
  const html = r.ui.dialogs[r.ui.dialogs.length - 1];
  check('the dialog lists it as transferred, with Undo, and the TDocs of the report to choose from', [html[0], /Transferred<\/b> → S4-265101/.test(html[1]), /Undo transfer/.test(html[1]), /<option value="S4-265002">/.test(html[1])], ['Revision Expected', true, true, true]);
  check('a TDoc without a table, a transferred record again: refused', [(() => { try { r.record({ original: 'S4-269999', decision: 'agreed' }); return 'saved'; } catch (e) { return /no TDoc table/.test(e.message); } })(),
    (() => { try { r.record({ original: 'S4-265001', decision: 'noted' }); return 'saved'; } catch (e) { return /already transferred/.test(e.message); } })()], [true, true]);
}

// ====================================================================== D
console.log('D. conflicts: nothing written');
{
  const r = report();
  r.build(); r.save(); r.update();
  r.record({ original: 'S4-265002', decision: 'agreed' });
  r.record({ original: 'S4-265003', decision: 'endorsed' });
  r.record({ original: 'S4-265001', decision: 'agreed', successor: 'S4-265199' });
  // 002: its revision registered and decided by hand in the same session before the update ran.
  r.register('S4-265102', 'S4-265002');
  r.update();
  r.user(() => r.api.pick('S4-265102', 'noted'));
  r.docProps.setProperty('PROVREV_S4-265002', JSON.stringify(Object.assign(r.rec('S4-265002'), { state: 'pending' })));
  // 003: two revisions name it.
  r.register('S4-265103', 'S4-265003'); r.tdocs.push({ id: 'S4-265113', agenda: '8.4', status: 'reserved', title: 'x', isRevisionOf: 'S4-265003', uploaded: false });
  // 001: another revision than the one expected.
  r.register('S4-265101', 'S4-265001');
  r.update();
  check('the revision decided by hand: conflict, its Status kept', [r.rec('S4-265002').state, r.shown('S4-265102')], ['conflict', 'dropdown:noted']);
  check('two revisions: conflict, neither written', [r.rec('S4-265003').state, r.shown('S4-265103'), r.shown('S4-265113'), /2 successors/.test(r.rec('S4-265003').reason)], ['conflict', 'dropdown:reserved', 'dropdown:reserved', true]);
  check('another revision than expected: conflict, not written', [r.rec('S4-265001').state, r.shown('S4-265101'), /S4-265199 was expected/.test(r.rec('S4-265001').reason)], ['conflict', 'dropdown:reserved', true]);
  check('Close Report blocks for each conflict', r.closeProblems().length, 3);
  // Resolution: a new record naming the right successor.
  r.record({ original: 'S4-265001', decision: 'agreed', successor: 'S4-265101' });
  r.update(); r.update();
  check('resolved by recording the right revision: transferred', [r.rec('S4-265001').state, r.shown('S4-265101')], ['transferred', 'dropdown:agreed']);
}

// ====================================================================== E
console.log('E. Goes to Plenary: the choice and the agenda item go with the decision');
{
  const r = report();
  r.build(); r.save(); r.update();
  r.exec(() => r.s.setUpGoesToPlenary()); r.save(); r.update();
  r.record({ original: 'S4-265001', decision: 'agreed', plenary: 'yes', agenda: '15.4' });
  r.record({ original: 'S4-265002', decision: 'agreed', plenary: 'yes', agenda: '' });
  r.register('S4-265101', 'S4-265001');
  // 002's revision exists already, with Plenary "No" in its row, before the list links the two.
  r.tdocs.push({ id: 'S4-265102', agenda: '8.6', status: 'reserved', title: 'y', uploaded: false });
  r.update();
  r.tdoc('S4-265002').revisedTo = 'S4-265102'; r.tdoc('S4-265102').isRevisionOf = 'S4-265002';
  r.update(); r.update();
  check('a new revision: agreed, Plenary Yes, A.I. 15.4 shown', [r.api.compactOf('S4-265101').status, r.api.compactOf('S4-265101').plenary, r.api.compactOf('S4-265101').ai], ['dropdown:agreed', YES, 'A.I.: 15.4']);
  check('a revision that has a Plenary choice already: its Status written, its Plenary NOT changed; to be set by hand',
    [r.api.compactOf('S4-265102').status, r.api.compactOf('S4-265102').plenary, r.rec('S4-265002').plenaryDone], ['dropdown:agreed', NO, 'manual']);
  check('Close Report says to set it by hand', r.closeProblems().some((p) => /S4-265102 has another Plenary choice than "yes".*set it by hand/.test(p)), true);
  r.user(() => r.api.pickPlenary('S4-265102', YES));
  r.update();
  check('once the rapporteur sets it, the next update clears it; Close Report no longer blocks', [r.rec('S4-265002').plenaryDone, r.closeProblems().some((p) => /S4-265102/.test(p))], ['same', false]);
  check('a plenary agenda item typed for the revision is kept (not replaced by the record)', (() => {
    const x = report(); x.build(); x.save(); x.update(); x.exec(() => x.s.setUpGoesToPlenary()); x.save(); x.update();
    x.record({ original: 'S4-265001', decision: 'agreed', plenary: 'yes', agenda: '15.4' });
    x.docProps.setProperty('PLENARY_AGENDA_MANUAL:S4-265101', '16.3');
    x.register('S4-265101', 'S4-265001'); x.update();
    return [x.docProps.getProperty('PLENARY_AGENDA_MANUAL:S4-265101'), x.api.compactOf('S4-265101').ai];
  })(), ['16.3', 'A.I.: 16.3']);
}

// ====================================================================== F
console.log('F. cancel and undo');
{
  const r = report();
  r.build(); r.save(); r.update();
  r.record({ original: 'S4-265001', decision: 'agreed' });
  r.change('cancel', 'S4-265001', { reason: 'No longer wanted by the group.' });
  r.register('S4-265101', 'S4-265001');
  r.update();
  check('cancelled while waiting: nothing transferred when the revision is registered', [r.rec('S4-265001').state, r.shown('S4-265101')], ['cancelled', 'dropdown:reserved']);
  r.record({ original: 'S4-265002', decision: 'endorsed' });
  r.register('S4-265102', 'S4-265002');
  r.update(); r.update();
  check('transferred', [r.shown('S4-265102'), r.rec('S4-265002').state], ['dropdown:endorsed', 'transferred']);
  const said = r.change('undo', 'S4-265002');
  check('undo asked: done by the next update', [/The next update sets S4-265102 back to "reserved"/.test(said), r.rec('S4-265002').state], [true, 'undo']);
  check('Close Report waits for it', r.closeProblems().some((p) => /undoing the transfer to S4-265102/.test(p)), true);
  r.update(); r.update();
  check('the next update sets the revision back (it still showed the decision), the one after confirms it', [r.shown('S4-265102'), r.rec('S4-265002').state], ['dropdown:reserved', 'undone']);
  r.record({ original: 'S4-265003', decision: 'agreed' });
  r.register('S4-265103', 'S4-265003');
  r.update(); r.update();
  r.change('undo', 'S4-265003');
  r.user(() => r.api.pick('S4-265103', 'noted'));
  r.update();
  check('undo refused when the revision was changed meanwhile: kept, said', [r.shown('S4-265103'), r.rec('S4-265003').state, /no longer shows "agreed"/.test(r.rec('S4-265003').reason)], ['dropdown:noted', 'transferred', true]);
  check('cancel/undo of the wrong state: refused', [(() => { try { r.change('cancel', 'S4-265003'); return 'done'; } catch (e) { return /only a record that waits/.test(e.message); } })(), (() => { try { r.change('undo', 'S4-265001'); return 'done'; } catch (e) { return /only a transfer can be undone/.test(e.message); } })()], [true, true]);
}

// ====================================================================== F2
console.log('F2. a revision never registered: closed explicitly, audited, information in Close Report (decision 2026-11-23)');
{
  const r = report();
  r.s.Session.getActiveUser = () => ({ getEmail: () => 'rapporteur@example.org' });
  r.build(); r.save(); r.update();
  r.record({ original: 'S4-265001', decision: 'agreed', plenary: '', agenda: '' });
  r.update();
  check('waiting: Close Report blocks on it', [r.closeProblems().some((p) => /^S4-265001: a revision is expected/.test(p)), r.exec(() => r.s.closeReportCore_()).ready], [true, false]);
  const refuse = (fn) => { try { fn(); return 'done'; } catch (e) { return e.message; } };
  check('resolving needs an outcome and a reason of 10 characters; cancelling a reason: refused otherwise, nothing changed',
    [/choose an outcome/.test(refuse(() => r.change('resolve', 'S4-265001', { outcome: '', reason: 'never came in time' }))),
      /at least 10 characters/.test(refuse(() => r.change('resolve', 'S4-265001', { outcome: 'not-submitted', reason: 'no' }))),
      /at least 10 characters/.test(refuse(() => r.change('cancel', 'S4-265001'))), r.rec('S4-265001').state], [true, true, true, 'pending']);
  const said = r.change('resolve', 'S4-265001', { outcome: 'not-submitted', reason: '  The source company did not submit the revision before the deadline.  ' });
  const rec = r.rec('S4-265001');
  check('resolved: the outcome, the reason, who, when and the state before are kept in the record and the log',
    [/Resolved without a revision/.test(said), rec.state, rec.resolution.outcome, rec.resolution.reason, rec.resolution.by, /^\d{4}-\d\d-\d\dT/.test(rec.resolution.at), /^pending/.test(rec.resolution.before),
      JSON.parse(r.docProps.getProperty('PROVREV_LOG'))[0].a, JSON.parse(r.docProps.getProperty('PROVREV_LOG'))[0].by],
    [true, 'resolved', 'not-submitted', 'The source company did not submit the revision before the deadline.', 'rapporteur@example.org', true, true, 'resolved', 'rapporteur@example.org']);
  const res = r.exec(() => r.s.closeReportCore_());
  check('Close Report: no longer blocking; listed under Information with the outcome, the reason, who and when',
    [r.closeProblems().length, res.informationNotes.filter((n) => /^S4-265001: the revision expected \("agreed"\) was closed without a transfer -- the revision was never submitted or registered: "The source company did not submit the revision before the deadline\." \(rapporteur@example\.org, \d{4}-\d\d-\d\d \d\d:\d\d UTC\)\.$/.test(n)).length, res.information >= 1],
    [0, 1, true]);
  r.register('S4-265101', 'S4-265001');
  r.update(); r.update();
  check('a revision registered after the resolution: nothing transferred, the record stays resolved', [r.shown('S4-265101'), r.rec('S4-265001').state], ['dropdown:reserved', 'resolved']);
  check('a resolved record cannot be resolved or cancelled again', /only a record that waits or is in conflict/.test(refuse(() => r.change('cancel', 'S4-265001', { reason: 'second thoughts here' }))), true);
  // Cancel with a reason (the outcome "cancelled"), from a conflict.
  r.record({ original: 'S4-265002', decision: 'endorsed', successor: 'S4-265199' });
  r.register('S4-265102', 'S4-265002');
  r.update();
  check('a conflict blocks', r.rec('S4-265002').state, 'conflict');
  r.change('cancel', 'S4-265002', { reason: 'Recorded for the wrong TDoc by mistake.' });
  check('cancelled from a conflict with a reason: information, not blocking; nothing written into the revision',
    [r.rec('S4-265002').state, r.rec('S4-265002').resolution.outcome, r.closeProblems().length, r.exec(() => r.s.closeReportCore_()).informationNotes.some((n) => /^S4-265002: .*cancelled \(recorded by mistake or no longer wanted\): "Recorded for the wrong TDoc by mistake\."/.test(n)), r.shown('S4-265102')],
    ['cancelled', 'cancelled', 0, true, 'dropdown:reserved']);
  r.exec(() => r.s.manageExpectedRevisions());
  const html = r.ui.dialogs[r.ui.dialogs.length - 1][1];
  check('the dialog: a form to close a record (outcome, required reason), the resolution shown in the list',
    [/Close a record without a transfer/.test(html), /Reason \(required/.test(html), /did not submit the revision/.test(html), /Resolved without a revision/.test(html)], [true, true, true, true]);
  // Recorded again after it was closed: allowed, and the resolution is kept in the history; logged as reopened, with the user.
  r.record({ original: 'S4-265001', decision: 'noted' });
  const reopened = r.rec('S4-265001');
  check('recorded again after its resolution: pending, the resolution kept in history, logged "reopened" by the user',
    [reopened.state, reopened.history.length, reopened.history[0].state, reopened.history[0].resolution.reason, reopened.by, JSON.parse(r.docProps.getProperty('PROVREV_LOG'))[0].a, JSON.parse(r.docProps.getProperty('PROVREV_LOG'))[0].by],
    ['pending', 1, 'resolved', 'The source company did not submit the revision before the deadline.', 'rapporteur@example.org', 'reopened', 'rapporteur@example.org']);
  check('"cancelled" is not a resolution outcome (Cancel is)', /use Cancel/.test(refuse(() => r.change('resolve', 'S4-265001', { outcome: 'cancelled', reason: 'should be refused here' }))), true);
  const noUser = report();
  noUser.s.Session.getActiveUser = () => { throw new Error('no permission'); };
  check('no user e-mail available: recorded as "dialog"', noUser.s.provisionalAuditUser_(), 'dialog');
}

// ====================================================================== R
console.log('R. the findings of the independent review of PR #30');
{
  // 1. A delegate decides the revision while the update runs: the queued write is not sent (left), the choice stays.
  const r = report();
  r.build(); r.save(); r.update();
  r.register('S4-265101', 'S4-265001'); r.update(); // the revision exists (dropdown "reserved") before the record links it
  r.tdoc('S4-265001').revisedTo = ''; r.tdoc('S4-265101').isRevisionOf = ''; r.tdoc('S4-265001').status = 'available'; r.update();
  r.record({ original: 'S4-265001', decision: 'agreed' });
  r.tdoc('S4-265001').revisedTo = 'S4-265101'; r.tdoc('S4-265101').isRevisionOf = 'S4-265001';
  const real = r.s.finalizeStatusDropdowns_;
  r.s.finalizeStatusDropdowns_ = function () { r.api.pick('S4-265101', 'noted'); r.body._state.dirty = false; return real.apply(this, arguments); };
  r.update();
  r.s.finalizeStatusDropdowns_ = real;
  check('a decision a delegate made during the update is not overwritten (the queued write is left)', r.shown('S4-265101'), 'dropdown:noted');
  r.update();
  check('... the next update records it as transferred, changed by hand afterwards, without blocking', [r.rec('S4-265001').state, /changed to "noted" after the transfer/.test(r.rec('S4-265001').note || ''), r.closeProblems()], ['transferred', true, []]);

  // 3. The write fails at the end of the run: the record stays "transferring" and the next update writes it again.
  const f = report();
  f.build(); f.save(); f.update();
  f.register('S4-265102', 'S4-265002'); f.update();
  f.tdoc('S4-265002').revisedTo = ''; f.tdoc('S4-265102').isRevisionOf = ''; f.tdoc('S4-265002').status = 'available'; f.update();
  f.record({ original: 'S4-265002', decision: 'endorsed' });
  f.tdoc('S4-265002').revisedTo = 'S4-265102'; f.tdoc('S4-265102').isRevisionOf = 'S4-265002';
  const fin = f.s.finalizeStatusDropdowns_;
  f.s.finalizeStatusDropdowns_ = function () { return { ran: false, inserted: 0, updated: 0, reason: 'simulated failure' }; };
  f.update();
  f.s.finalizeStatusDropdowns_ = fin;
  check('the write did not happen: still "reserved", the record "transferring" (not lost)', [f.shown('S4-265102'), f.rec('S4-265002').state], ['dropdown:reserved', 'transferring']);
  f.update(); f.update();
  check('... the next update writes it again, the one after confirms it', [f.shown('S4-265102'), f.rec('S4-265002').state], ['dropdown:endorsed', 'transferred']);

  // 4. Two records naming one revision with different choices: both conflicts, nothing written.
  const g = report();
  g.build(); g.save(); g.update();
  g.record({ original: 'S4-265001', decision: 'agreed' });
  g.record({ original: 'S4-265002', decision: 'noted' });
  g.tdocs.push({ id: 'S4-265109', agenda: '8.5', status: 'reserved', title: 'merged', isRevisionOf: 'S4-265001', uploaded: false });
  g.tdoc('S4-265002').revisedTo = 'S4-265109';
  g.update();
  check('two records, one revision, different choices: both conflicts, the revision untouched', [g.rec('S4-265001').state, g.rec('S4-265002').state, g.shown('S4-265109')], ['conflict', 'conflict', 'dropdown:reserved']);

  // 5. A chain whose newest revision is named but not registered: wait (never written to the middle one).
  const { sandbox: S } = loadCode({ documentProperties: {} });
  const facts = { A: { revisedTo: ['B'], isRevisionOf: [] }, B: { revisedTo: ['C'], isRevisionOf: ['A'] } };
  const d = S.decideProvisionalTransfer_({ v: 1, original: 'A', decision: 'agreed', state: 'pending' }, facts, () => ({ exists: true, label: 'reserved' }));
  check('A -> B -> C with C not registered yet: wait', [d.action, d.reason], ['wait', 'B is revised to C, not registered yet']);

  // 7, 8. "partially agreed" not offered; Undo not offered for a record whose revision showed the decision already.
  check('"partially agreed" is not offered (no option of the template dropdown)', S.validateProvisionalInput_({ original: 'S4-265001', decision: 'partially agreed' }).ok, false);
  const h = S.provisionalDialogHtml_({ 'S4-1': { original: 'S4-1', decision: 'agreed', plenary: '', agenda: '', state: 'transferred', to: 'S4-2' } }, [], [], false);
  check('a record whose revision showed the decision already: no Undo button', /Undo transfer/.test(h), false);
}

// ====================================================================== G
console.log('G. no record: nothing more read or written');
{
  const A = ledger.codeWithoutIssue28B(RAW).code;
  const run = (code) => { const m = report(Object.assign({}, code ? { code: code } : {})); m.build(); m.save(); m.update(); m.update(); const a = m.body._state.calls; const res = m.update(); return { calls: m.body._state.calls - a, res: res }; };
  const now = run(null); const before = run(A);
  check('an update without records: the same DocumentApp calls as phase A, the same result', [now.calls, now.res], [before.calls, before.res]);
}

// ====================================================================== H
console.log('H. the ledger');
{
  const git = (spec) => { try { return execFileSync('git', ['show', spec], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, ''); } catch (e) { return null; } };
  const view = ledger.codeWithoutIssue28B(RAW);
  check('every change of the layer is in Code.js exactly once', view.problems, []);
  // Phase A pinned by its commit (6f8c9df); the branch names only as a fallback.
  const a = git('6f8c9df:Code.js') || git('feature/i28-a-titles-headings-perf:Code.js') || git('origin/feature/i28-a-titles-headings-perf:Code.js');
  if (a === null) console.log('  note: the phase A branch is not available; the byte comparison is skipped.');
  else check('without it, Code.js is that of phase A, byte for byte', view.code === a, true);
  const creator = fs.readFileSync(path.join(ROOT, 'template', 'ReportCreator.js'), 'utf8').replace(/\r/g, '');
  const ca = git('6f8c9df:template/ReportCreator.js') || git('feature/i28-a-titles-headings-perf:template/ReportCreator.js') || git('origin/feature/i28-a-titles-headings-perf:template/ReportCreator.js');
  check('ReportCreator.js: one menu item; without it, that of phase A', [ledger.creatorWithoutIssue28B(creator).problems, ca === null || ledger.creatorWithoutIssue28B(creator).code === ca], [[], true]);
  check('the version is 2.29.0', (RAW.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.29.0');
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((x) => x.slice(9, -1));
  const bodyOf = (code, n) => { const at = code.indexOf('\nfunction ' + n + '('); if (at === -1) return null; return code.slice(at, code.indexOf('\n}\n', at)); };
  check('its functions are those of its section; it changes two released functions by one hook each', [names(RAW).filter((n) => names(view.code).indexOf(n) === -1).sort(), names(view.code).filter((n) => bodyOf(RAW, n) !== bodyOf(view.code, n)).sort()],
    [ledger.I28B_SECTION_FUNCTIONS.slice().sort(), ledger.I28B_CHANGED_FUNCTIONS.slice().sort()]);
}

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
