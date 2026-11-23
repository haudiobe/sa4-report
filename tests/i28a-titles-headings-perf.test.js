/**
 * Issue #28, phase A (docs/ISSUE28_PHASE_A.md): titles from the TDoc list, Heading 2 class lines in the Disposition
 * Summary, fewer DocumentApp calls in an update.
 *
 *   A. decideTitleSync_(): every decision and its boundaries (pure).
 *   B. the record, the hash, the bounded JSON (pure).
 *   C. an MBS report of SA4#138 with Status dropdowns: a title changed after the reservation is written into the
 *      TDoc table and the Registration of Documents table, its link and look kept, Minutes/Disposition untouched;
 *      the next update writes nothing; the summary follows.
 *   D. what is never written: a title typed by hand, a list that goes back to an older title, an unknown difference;
 *      the review dialog (apply, keep); the settings 'review' and 'off'.
 *   E. a report with text statuses (no report snapshot): the Title rows read at most 60 per update, the rest later;
 *      a TDoc whose list title is recorded costs nothing.
 *   F. the class lines of the Disposition Summary: Heading 2 in both forms; a summary of T-2026.10.34 restyled by the
 *      next update without a rewrite or a section break; idempotent; 'normal' keeps bold Normal; text below kept,
 *      text typed into it refused as before.
 *   G. the cost changes give the same results: the scan, the Status row, the summary check (pure equivalence), and
 *      the DocumentApp calls of an update against T-2026.10.34 on the same report.
 *   H. the ledger: without this layer Code.js and ReportCreator.js are those of T-2026.10.34, byte for byte.
 *
 * All numbers, names and titles are synthetic; the agenda is the real one of SA4#138 (S4-261857).
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
const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const T34 = ledger.codeWithoutIssue28A(FULL).code;
const AGENDA_138 = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'fixtures', 'agenda-S4-261857.google-items.json'), 'utf8')).items
  .map((i) => ({ number: i.number, title: i.title, level: i.level, heading: 'NORMAL', text: '' }));
const MAIN = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_NAME: 'SA4#138', MEETING_ID: '60780', REPORT_SUFFIX: 'MBS', MEETING_DATE: 'November 16, 2026',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx',
  AGENDA_TDOC: 'S4-261857' };
const SUMMARY = 'Disposition Summary';
const TDOCS = (n) => Array.from({ length: n }, (_, k) => {
  const num = String(k + 1).padStart(3, '0');
  return { id: 'S4-264' + num, agenda: ['8.4', '8.5', '8.6', '8.7', '8.8'][k % 5], status: ['agreed', 'noted', 'available', 'endorsed', 'reserved'][k % 5], title: 'Contribution ' + num, uploaded: true };
});
const isTable = (c) => c.getType() === 'TABLE';
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

function report(o) {
  const r = templateReport(Object.assign({ props: MAIN, agenda: AGENDA_138, tdocs: TDOCS(12), tableSeparators: true, endParagraph: true }, o || {}));
  r.s.HtmlService = { createHtmlOutput: (html) => { const out = { setWidth: () => out, setHeight: () => out, getContent: () => html }; return out; } };
  r.ui.dialogs = [];
  r.ui.showModalDialog = (output, title) => { r.ui.dialogs.push([title, output.getContent()]); };
  // Each update ends with the save of its execution, as in Google Docs.
  const up = r.update;
  r.update = function () { const out = up.apply(this, arguments); r.body._state.dirty = false; r.body._state.closed = false; return out; };
  r.save = () => { r.body._state.dirty = false; r.body._state.closed = false; };
  r.cell = (id, label) => { const t = r.table(id); for (let i = 0; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === label) return t.getRow(i).getCell(1); return null; };
  r.registration = (id) => { const t = r.body._children.filter((c) => isTable(c) && c.getNumRows() && c.getRow(0).getNumCells() >= 4 && c.getRow(0).getCell(1).getText() === 'Title' && c.getRow(0).getCell(3).getText() === 'Agenda Item')[0];
    const row = rowsOf(t).filter((x) => x[0] === id)[0]; return row ? row[1] : null; };
  r.pending = () => JSON.parse(r.docProps.getProperty('TITLE_SYNC_PENDING') || '{}');
  r.log = () => JSON.parse(r.docProps.getProperty('TITLE_SYNC_LOG') || '[]');
  r.doc = () => JSON.stringify(r.body._children.map((c) => (isTable(c) ? ['T', rowsOf(c)] : ['P', c.getHeading(), c.getText(), !!c._bold])));
  r.propsText = () => JSON.stringify(Object.keys(r.docProps._store).sort().map((k) => [k, r.docProps._store[k]]));
  r.summaryLines = () => {
    const kids = r.body._children;
    const at = kids.findIndex((c) => !isTable(c) && c.getText() === SUMMARY && c.getHeading() !== 'NORMAL');
    return at === -1 ? null : kids.slice(at + 1).filter((c) => !isTable(c) && / \(\d+\)$/.test(c.getText())).map((c) => [c.getText().replace(/ \(\d+\)$/, ''), c.getHeading(), !!c._bold]);
  };
  r.summaryTitles = () => r.body._children.filter((c) => isTable(c) && c.getNumRows() && c.getRow(0).getCell(0).getText() === 'TDoc' && c.getRow(0).getNumCells() >= 4 && c.getRow(0).getCell(1).getText() === 'Title' && c.getRow(0).getCell(3).getText() === 'Status')
    .map(rowsOf).reduce((a, rows) => a.concat(rows.slice(1).map((x) => x[0] + '=' + x[1])), []);
  return r;
}

// ====================================================================== A
console.log('A. decideTitleSync_(): the decisions');
{
  const { sandbox: S } = loadCode({ documentProperties: {} });
  const h = S.titleSyncHash_;
  const rec = (cur, prev, ack) => ({ cur: cur ? h(cur) : '', prev: (prev || []).map(h), ack: ack ? h(ack) : '' });
  const d = (cell, list, record, mode) => S.decideTitleSync_(cell, list, record || { cur: '', prev: [], ack: '' }, mode || 'auto').action;
  check('the same title, recorded: in sync; first seen: recorded', [d('Old', 'Old', rec('Old')), d('Old', 'Old')], ['in-sync', 'record']);
  check('white space does not count (collapsed, trimmed); case does (a cell changed in case only is a change by hand)', [d('  A   title ', 'A title', rec('A title')), d('A Title', 'A title', rec('A title'))], ['in-sync', 'manual']);
  check('the list changed, the report shows the title before: apply', d('Reserved title', 'Final title', rec('Reserved title')), 'apply');
  check('... but not in the setting "review"', d('Reserved title', 'Final title', rec('Reserved title'), 'review'), 'review');
  check('the report was changed by hand, the list is the same: manual (kept)', d('My better title', 'Final title', rec('Final title')), 'manual');
  check('both changed: review', d('My better title', 'Newer title', rec('Final title')), 'review');
  check('no record and a difference: review, never written', d('Some title', 'Other title'), 'review');
  check('a list that goes back to a title it had before: stale, never written', d('Final title', 'Reserved title', rec('Final title', ['Reserved title'])), 'stale');
  check('the rapporteur kept the report title against this list title: kept; a newer list title is offered again',
    [d('Mine', 'Final title', rec('Old', [], 'Final title')), d('Mine', 'Newest title', rec('Old', [], 'Final title'))], ['kept', 'review']);
  check('the same title again after "keep": recorded, the keep cleared', S.decideTitleSync_('Final title', 'Final title', rec('Old', [], 'Final title'), 'auto').record.ack, '');
  check('nothing to compare: an empty list title, a cell of more than one line', [d('Title', ''), d('Title', '   '), d('Line one\nLine two', 'Line one Line two', rec('x'))], ['skip', 'skip', 'skip']);
  const applied = S.decideTitleSync_('Reserved title', 'Final title', rec('Reserved title', ['A', 'B', 'C']), 'auto').record;
  check('after "apply" the record holds the new title, the one before first among at most three earlier ones', [applied.cur === h('Final title'), applied.prev.length, applied.prev[0] === h('Reserved title')], [true, 4, true]);
  check('... and when it is stored, three earlier titles are kept', S.titleSyncRecordOf_(S.titleSyncRecordText_(applied)).prev.length, 3);
}

// ====================================================================== B
console.log('B. the record, the hash, the bounded JSON');
{
  const { sandbox: S } = loadCode({ documentProperties: {} });
  check('the hash is 8 hex digits, stable, the same for white space differences', [S.titleSyncHash_('A  title'), S.titleSyncHash_('A title '), /^[0-9a-f]{8}$/.test(S.titleSyncHash_('x'))], [S.titleSyncHash_('A title'), S.titleSyncHash_('A title'), true]);
  check('a record reads back as it was written; an unknown value is no record', [S.titleSyncRecordOf_(S.titleSyncRecordText_({ cur: 'aaaaaaaa', prev: ['bbbbbbbb'], ack: 'cccccccc' })), S.titleSyncRecordOf_('garbage'), S.titleSyncRecordOf_(null)],
    [{ cur: 'aaaaaaaa', prev: ['bbbbbbbb'], ack: 'cccccccc' }, { cur: '', prev: [], ack: '' }, { cur: '', prev: [], ack: '' }]);
  const big = Array.from({ length: 400 }, (_, k) => ({ t: 'S4-26' + k, a: 'applied', o: 'x'.repeat(60), n: 'y'.repeat(60), at: '2026-11-23T10:00:00Z' }));
  const text = S.titleSyncBoundedJson_(big);
  check('the log is cut to one property value (newest kept)', [text.length <= 8500, JSON.parse(text)[0].t], [true, 'S4-260']);
  const pend = {}; for (let k = 0; k < 400; k++) pend['S4-26' + k] = { cell: 'c'.repeat(40), list: 'l'.repeat(40), kind: 'review', at: '2026-11-' + String(10 + (k % 10)) + 'T00:00:00Z' };
  check('the pending list is cut to one property value, the newest kept', S.titleSyncBoundedJson_(pend).length <= 8500, true);
  check('a long title is shortened for the log', [S.titleSyncShort_('z'.repeat(300)).length, S.titleSyncShort_('short')], [158, 'short']);
}

// ====================================================================== C
console.log('C. a title changed after the reservation: written, link and look kept, Minutes untouched, the summary follows');
{
  const r = report();
  r.build(); r.save();
  r.update();
  const seeded = Object.keys(r.docProps._store).filter((k) => k.indexOf('TDOC_TITLE_SEEN_') === 0).length;
  check('the first update records what it sees (one record per TDoc), writes no title, lists nothing', [seeded, r.pending()], [12, {}]);
  r.exec(() => r.s.updateDispositionSummary()); r.save();
  // A link over the whole title and Minutes typed by hand.
  r.user(() => { const c = r.cell('S4-264003', 'Title'); c.editAsText().setLinkUrl(0, c.getText().length - 1, 'https://example.org/t3'); r.cell('S4-264003', 'Minutes').setText('Presented. Agreed with changes.'); r.cell('S4-264003', 'Disposition').setText('Revised wording agreed.'); });
  r.tdoc('S4-264003').title = 'Contribution 003: the final title in 3GU';
  const u = r.update();
  check('the update writes the new title into the TDoc table and the Registration of Documents table', [r.cell('S4-264003', 'Title').getText(), r.registration('S4-264003'), u.titles],
    ['Contribution 003: the final title in 3GU', 'Contribution 003: the final title in 3GU', '1 written, 0 to review, 0 changed in the report, 0 from an older list']);
  check('the link over the whole title is on the whole new title', [r.cell('S4-264003', 'Title').editAsText().getLinkUrl(0), r.cell('S4-264003', 'Title').editAsText().getLinkUrl(39)], ['https://example.org/t3', 'https://example.org/t3']);
  check('Minutes and Disposition typed by hand are as they were', [r.cell('S4-264003', 'Minutes').getText(), r.cell('S4-264003', 'Disposition').getText()], ['Presented. Agreed with changes.', 'Revised wording agreed.']);
  check('the Disposition Summary shows the new title (it copies the TDoc table)', r.summaryTitles().indexOf('S4-264003=Contribution 003: the final title in 3GU') !== -1, true);
  check('the audit log: old, new, when, by the update', [r.log()[0].t, r.log()[0].a, r.log()[0].o, r.log()[0].n, r.log()[0].by, /^2\d{3}-/.test(r.log()[0].at)], ['S4-264003', 'applied', 'Contribution 003', 'Contribution 003: the final title in 3GU', 'update', true]);
  const doc = r.doc(); const props = r.propsText();
  const u2 = r.update();
  check('the next update writes nothing: the same document, no state written, no "titles" in the result', [r.doc() === doc, r.propsText() === props, u2.titles], [true, true, undefined]);
  // The same TDoc changes again: the newer title is written too (the record follows).
  r.tdoc('S4-264003').title = 'Contribution 003: corrected again';
  r.update();
  check('a second change of the list is written as well', [r.cell('S4-264003', 'Title').getText(), r.registration('S4-264003')], ['Contribution 003: corrected again', 'Contribution 003: corrected again']);
  // A new TDoc gets its title at insertion and is recorded; nothing listed.
  r.tdocs.push({ id: 'S4-264099', agenda: '8.6', status: 'reserved', title: 'A late contribution', uploaded: true });
  r.update();
  check('a new TDoc: its title from the insertion, recorded, nothing listed', [r.cell('S4-264099', 'Title').getText(), !!r.docProps.getProperty('TDOC_TITLE_SEEN_S4-264099'), r.pending()], ['A late contribution', true, {}]);
}

// ====================================================================== D
console.log('D. never written: by hand, an older list, an unknown difference; the dialog; the settings');
{
  const r = report();
  r.build(); r.save(); r.update();
  // By hand: the rapporteur improves a title; the list is unchanged.
  r.user(() => { const c = r.cell('S4-264004', 'Title'); c.editAsText().deleteText(0, c.getText().length - 1); c.editAsText().insertText(0, 'A better title typed by hand'); });
  r.update();
  check('a title typed by hand is kept, and listed as such', [r.cell('S4-264004', 'Title').getText(), r.pending()['S4-264004'].kind], ['A better title typed by hand', 'manual']);
  // An older list: 005 changes, then the list shows the old title again.
  r.tdoc('S4-264005').title = 'Contribution 005 (v2)';
  r.update();
  r.tdoc('S4-264005').title = 'Contribution 005';
  r.update();
  check('a list that goes back to the older title is not written; listed as an older list', [r.cell('S4-264005', 'Title').getText(), r.pending()['S4-264005'].kind], ['Contribution 005 (v2)', 'stale']);
  // Unknown: a title that differs without a record (a report of before this release).
  r.docProps.deleteProperty('TDOC_TITLE_SEEN_S4-264006');
  r.tdoc('S4-264006').title = 'Contribution 006 renamed in 3GU';
  const u = r.update();
  check('a difference without a record is listed for review, not written', [r.cell('S4-264006', 'Title').getText(), r.pending()['S4-264006'].kind, u.titles], ['Contribution 006', 'review', undefined]);
  // The dialog shows them; "Use the list title" writes 006; "Keep the report title" keeps 004.
  r.exec(() => r.s.reviewTitleChanges());
  const html = r.ui.dialogs[r.ui.dialogs.length - 1];
  check('Review Title Changes… lists the three, with both titles and why', [html[0], /S4-264004/.test(html[1]) && /S4-264005/.test(html[1]) && /S4-264006/.test(html[1]), /Contribution 006 renamed in 3GU/.test(html[1]), /Typed by hand/.test(html[1])], ['Review Title Changes', true, true, true]);
  const applied = r.exec(() => r.s.applyTitleReviewDecisions('apply', ['S4-264006'])); r.save();
  check('"Use the list title" writes it into both tables', [r.cell('S4-264006', 'Title').getText(), r.registration('S4-264006'), /Wrote the list title of S4-264006/.test(applied)], ['Contribution 006 renamed in 3GU', 'Contribution 006 renamed in 3GU', true]);
  r.exec(() => r.s.applyTitleReviewDecisions('keep', ['S4-264004'])); r.save();
  r.update();
  check('"Keep the report title": not listed again, not written', [r.cell('S4-264004', 'Title').getText(), !!r.pending()['S4-264004'], !!r.pending()['S4-264006']], ['A better title typed by hand', false, false]);
  check('the dialog refuses a title changed since the update', (() => {
    r.tdoc('S4-264007').title = 'x'; r.docProps.deleteProperty('TDOC_TITLE_SEEN_S4-264007'); r.update();
    r.user(() => r.cell('S4-264007', 'Title').editAsText().insertText(0, 'Edited: '));
    return /changed since the update/.test(r.exec(() => r.s.applyTitleReviewDecisions('apply', ['S4-264007']))) && r.cell('S4-264007', 'Title').getText() === 'Edited: Contribution 007';
  })(), true);
  // Settings.
  const off = report({ props: Object.assign({}, MAIN, { TITLE_SYNC_MODE: 'off' }) });
  off.build(); off.save(); off.update();
  off.tdoc('S4-264002').title = 'Changed';
  off.update();
  check('TITLE_SYNC_MODE off: nothing compared, nothing recorded, nothing written', [off.cell('S4-264002', 'Title').getText(), Object.keys(off.docProps._store).filter((k) => /^TDOC_TITLE_SEEN_|^TITLE_SYNC_(PENDING|LOG)/.test(k))], ['Contribution 002', []]);
  const rev = report({ props: Object.assign({}, MAIN, { TITLE_SYNC_MODE: 'review' }) });
  rev.build(); rev.save(); rev.update();
  rev.tdoc('S4-264002').title = 'Changed';
  rev.update();
  check('TITLE_SYNC_MODE review: listed, not written', [rev.cell('S4-264002', 'Title').getText(), rev.pending()['S4-264002'].kind], ['Contribution 002', 'review']);
}

// ====================================================================== E
console.log('E. a report with text statuses (no report snapshot): bounded reads, nothing read for a recorded title');
{
  const r = report({ release: false, tdocs: TDOCS(150) });
  r.build(); r.save();
  const reads = () => { let n = 0; const real = r.s.titleCellOfTdocTable_; r.s.titleCellOfTdocTable_ = function () { n++; return real.apply(this, arguments); }; return () => { r.s.titleCellOfTdocTable_ = real; return n; }; };
  let stop = reads(); r.update(); const first = stop();
  stop = reads(); r.update(); const second = stop();
  stop = reads(); r.update(); const third = stop();
  stop = reads(); r.update(); const fourth = stop();
  check('at most 60 Title rows read per update until every TDoc is recorded (150: 60, 60, 30), then none', [first, second, third, fourth], [60, 60, 30, 0]);
  r.tdoc('S4-264010').title = 'Contribution 010, final';
  stop = reads(); r.update(); const changed = stop();
  check('a changed list title: that one row read and written', [changed, r.cell('S4-264010', 'Title').getText(), r.registration('S4-264010')], [2, 'Contribution 010, final', 'Contribution 010, final']);
}

// ====================================================================== F
console.log('F. the class lines of the Disposition Summary are Heading 2');
{
  // The three-table form.
  const r = report();
  r.build(); r.save(); r.update();
  r.exec(() => r.s.updateDispositionSummary()); r.save();
  check('three-table form: the class lines are Heading 2, not bold; the summary heading is Heading 1', [r.summaryLines(), r.body._children.filter((c) => !isTable(c) && c.getText() === SUMMARY)[0].getHeading()],
    [[['Agreed, Endorsed and Approved Documents', 'H2', false], ['Other Disposed Documents', 'H2', false], ['Documents Not Yet Disposed', 'H2', false]], 'H1']);
  check('the section is found whole (its class headings do not end it): complete, nothing foreign', (() => { const s = r.s.findDispositionSummarySection_(r.body); return [s.complete, s.foreign]; })(), [true, []]);
  const doc = r.doc();
  const u = r.update();
  check('an update leaves it as it is (unchanged, no restyle)', [u.dispositionSummary, r.doc() === doc], ['unchanged', true]);
  r.user(() => r.body.appendParagraph('Text typed below the summary.'));
  r.tdoc('S4-264003').status = 'withdrawn';
  r.update();
  check('text typed below it is kept', r.body._children[r.body._children.length - 1].getText() === 'Text typed below the summary.' || r.body._children.some((c) => !isTable(c) && c.getText() === 'Text typed below the summary.'), true);

  // A summary written by T-2026.10.34 (bold Normal class lines) is restyled by the next update of this release, not rewritten.
  const t34 = report({ code: T34 });
  t34.build(); t34.save(); t34.update();
  t34.exec(() => t34.s.updateDispositionSummary()); t34.save();
  check('T-2026.10.34 wrote bold Normal class lines', t34.summaryLines().map((x) => [x[1], x[2]]), [['NORMAL', true], ['NORMAL', true], ['NORMAL', true]]);
  // The same summary in a report of this release: what T-2026.10.34 wrote (the same text; the class lines bold Normal).
  const toOldLook = (x) => x.user(() => x.body._children.forEach((c) => { if (!isTable(c) && x.s.isDispositionClassLineText_(c.getText())) { c.setHeading('NORMAL'); c.editAsText().setBold(true); } }));
  const old = report();
  old.build(); old.save(); old.update();
  old.exec(() => old.s.updateDispositionSummary()); old.save();
  toOldLook(old);
  check('... the same text as this release writes (only the look differs)', JSON.stringify(old.summaryLines().map((x) => x[0])) === JSON.stringify(t34.summaryLines().map((x) => x[0])), true);
  const texts = () => old.body._children.map((c) => (isTable(c) ? 'T' + JSON.stringify(rowsOf(c)) : c.getText()));
  const before = texts();
  const u1 = old.update();
  check('the next update restyles them to Heading 2, the text and tables of the summary as they were', [u1.dispositionSummary, old.summaryLines().map((x) => [x[1], x[2]]), JSON.stringify(texts()) === JSON.stringify(before)],
    ['unchanged', [['H2', false], ['H2', false], ['H2', false]], true]);
  const again = old.doc();
  old.update();
  check('... and the update after it changes nothing', old.doc() === again, true);

  // The four-table form (Goes to Plenary): restyled through the Docs API look step, no section break added.
  const p34 = report({ code: T34 });
  p34.build(); p34.save(); p34.update();
  p34.exec(() => p34.s.setUpGoesToPlenary()); p34.save(); p34.update();
  check('T-2026.10.34, Goes to Plenary: four bold Normal class lines, the summary in a landscape section', [p34.summaryLines().map((x) => [x[1], x[2]]), p34.api.sectionBreaks().length], [[['NORMAL', true], ['NORMAL', true], ['NORMAL', true], ['NORMAL', true]], 1]); // one break: before the summary (no text below it)
  const p = report();
  p.build(); p.save(); p.update();
  p.exec(() => p.s.setUpGoesToPlenary()); p.save(); p.update();
  toOldLook(p);
  check('... this release: the same summary and section breaks once given that look', [p.summaryLines().map((x) => [x[0], x[1], x[2]]), p.api.sectionBreaks().length], [p34.summaryLines(), p34.api.sectionBreaks().length]);
  const breaks = JSON.stringify(p.api.sectionBreaks());
  const pu = p.update();
  check('this release: the four class lines Heading 2 after one update; no section break added; the summary unchanged', [pu.dispositionSummary, p.summaryLines().map((x) => [x[1], x[2]]), JSON.stringify(p.api.sectionBreaks()) === breaks],
    ['unchanged', [['H2', false], ['H2', false], ['H2', false], ['H2', false]], true]);
  check('the landscape step sees the whole summary: nothing more to do', p.s.nextPlenaryLandscapeStep_(p.s.plenarySnapshot_('REPORTdoc000000000000000000000000000000000').content), null);
  const writes = p.api.calls.filter((c) => c.call === 'batchUpdate').length;
  p.update();
  check('the next update makes no API write', p.api.calls.filter((c) => c.call === 'batchUpdate').length, writes);
  // Text typed into the summary (between two class lines) still stops it, as before.
  p.user(() => { const kids = p.body._children; const at = kids.findIndex((c) => !isTable(c) && /^Other Dispositions – Goes to Plenary/.test(c.getText())); p.body.insertParagraph(at, 'A note typed into the summary.'); });
  p.tdoc('S4-264001').status = 'noted';
  const refused = p.exec(() => p.s.writeDispositionSummary_(p.body, { scope: p.s.reportTdocScope_(p.body) }));
  check('text typed into it, between class headings: refused, named', [refused.state, refused.foreign], ['refused', ['A note typed into the summary.']]);

  // Before a code rollback: 'normal' gives an existing Heading 2 summary the look of T-2026.10.34 again in one update.
  r.docProps.setProperty('DISPOSITION_CLASS_HEADINGS', 'normal');
  const rbBefore = r.summaryLines().map((x) => x[0]);
  r.update();
  check('rollback preparation: DISPOSITION_CLASS_HEADINGS normal restyles the Heading 2 class lines to bold Normal in one update, the text kept',
    [r.summaryLines().map((x) => [x[1], x[2]]), JSON.stringify(r.summaryLines().map((x) => x[0])) === JSON.stringify(rbBefore)], [[['NORMAL', true], ['NORMAL', true], ['NORMAL', true]], true]);
  r.docProps.deleteProperty('DISPOSITION_CLASS_HEADINGS');

  // The setting 'normal': bold Normal lines as before.
  const n = report({ props: Object.assign({}, MAIN, { DISPOSITION_CLASS_HEADINGS: 'normal' }) });
  n.build(); n.save(); n.update();
  n.exec(() => n.s.updateDispositionSummary()); n.save();
  check('DISPOSITION_CLASS_HEADINGS normal: bold Normal class lines, as in T-2026.10.34', n.summaryLines().map((x) => [x[1], x[2]]), [['NORMAL', true], ['NORMAL', true], ['NORMAL', true]]);
}

// ====================================================================== G
console.log('G. the cost changes give the same results; the calls of an update against T-2026.10.34');
{
  const r = report({ tdocs: TDOCS(20) });
  r.build(); r.save();
  const S = r.s;
  const tables = r.body.getTables();
  const existing = new Set();
  const index = S.existingTdocScanOfTables_(tables, existing);
  const oldIndex = S.buildTdocTableIndex_(tables);
  const oldExisting = new Set(); tables.forEach((t) => { if (!S.isTDocTable_(t)) return; const x = S.safeCellText_(t, 0, 1).trim(); if (x) oldExisting.add(x); });
  check('the one-pass scan: the same numbers and the same index as the two passes of before', [[...existing].sort(), [...index.keys()].sort(), [...index.keys()].every((k) => index.get(k) === oldIndex.get(k))], [[...oldExisting].sort(), [...oldIndex.keys()].sort(), true]);
  const oldFind = (table) => { for (let r2 = 0; r2 < table.getNumRows(); r2++) { const row = table.getRow(r2); if (row.getNumCells() < 2) continue; const k = row.getCell(0).getText().trim().toLowerCase(); if (k === 'tdoc status' || k === 'status') return { cell: row.getCell(1), value: S.statusPartOfCellText_(row.getCell(1).getText()).trim() }; } return null; };
  const odd = r.body.appendTable([['TDoc', 'S4-269999'], ['note'], ['Status', 'agreed​ Plenary ']]);
  const none = r.body.appendTable([['TDoc', 'S4-269998'], ['Title', 'x']]);
  check('the Status row: the same cell and value as before, for every TDoc table, a one-cell row and none', tables.concat([odd, none]).filter(S.isTDocTable_).every((t) => { const a = S.findStatusRowOfTable_(t); const b = oldFind(t); return (!a && !b) || (a && b && a.cell === b.cell && a.value === b.value); }), true);
  check('a report with the summary heading text: not absent; without: absent (one call)', [(() => { const x = report(); x.build(); return x.s.dispositionSummaryAbsentInBody_(x.body); })(), (() => { r.exec(() => r.s.updateDispositionSummary()); return r.s.dispositionSummaryAbsentInBody_(r.body); })()], [true, false]);

  // The calls of an update, this release against T-2026.10.34, on the same report of 150 TDocs (relative, not a time).
  const measure = (code, release) => {
    const m = report(Object.assign({ tdocs: TDOCS(150) }, code ? { code: code } : {}, release === false ? { release: false } : {}));
    m.build(); m.save();
    for (let k = 0; k < 4; k++) m.update(); // past the seeding of the titles (60 per update without a snapshot)
    const a = m.body._state.calls; m.update(); return { calls: m.body._state.calls - a, doc: m.doc() };
  };
  const dNow = measure(null), dOld = measure(T34);
  const tNow = measure(null, false), tOld = measure(T34, false);
  console.log('  unchanged update, 150 TDocs, DocumentApp calls -- Status dropdowns: T-2026.10.34 ' + dOld.calls + ', now ' + dNow.calls + '; text statuses: T-2026.10.34 ' + tOld.calls + ', now ' + tNow.calls);
  check('Status dropdowns: at least 3x fewer DocumentApp calls in an unchanged update; the same document', [dOld.calls / Math.max(1, dNow.calls) >= 3, dNow.doc === dOld.doc], [true, true]);
  check('text statuses: at least 20% fewer; the same document', [tNow.calls <= tOld.calls * 0.8, tNow.doc === tOld.doc], [true, true]);
}

// ====================================================================== H
console.log('H. the ledger');
{
  const base = (() => { try { return execFileSync('git', ['show', '21e0b8e:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, ''); } catch (e) { return null; } })();
  const view = ledger.codeWithoutIssue28A(FULL);
  check('every change of the layer is in Code.js exactly once', view.problems, []);
  if (base === null) console.log('  note: commit 21e0b8e is not available; the byte comparison is skipped.');
  else check('without it, Code.js is that of T-2026.10.34 (21e0b8e), byte for byte', view.code === base, true);
  const creator = fs.readFileSync(path.join(ROOT, 'template', 'ReportCreator.js'), 'utf8').replace(/\r/g, '');
  const cview = ledger.creatorWithoutIssue28A(creator);
  const cbase = (() => { try { return execFileSync('git', ['show', '21e0b8e:template/ReportCreator.js'], { cwd: ROOT, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, ''); } catch (e) { return null; } })();
  check('ReportCreator.js: one menu item; without it, that of T-2026.10.34', [cview.problems, cbase === null || cview.code === cbase], [[], true]);
  check('the version is 2.28.0', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.28.0');
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((x) => x.slice(9, -1));
  const bodyOf = (code, n) => { const at = code.indexOf('\nfunction ' + n + '('); if (at === -1) return null; const end = code.indexOf('\n}\n', at); return code.slice(at, end); };
  check('the functions it adds are those of its section; the released functions it changes are listed', [names(FULL).filter((n) => names(view.code).indexOf(n) === -1).sort(), names(view.code).filter((n) => bodyOf(FULL, n) !== bodyOf(view.code, n)).sort()],
    [ledger.I28A_SECTION_FUNCTIONS.slice().sort(), ledger.I28A_CHANGED_FUNCTIONS.slice().sort()]);
  check('the local listing agent stays as it is: no new reference to 10.10.10.10 and no LOCAL_LISTING setting written', [(FULL.match(/10\.10\.10\.10/g) || []).length, (T34.match(/10\.10\.10\.10/g) || []).length], [(T34.match(/10\.10\.10\.10/g) || []).length, (T34.match(/10\.10\.10\.10/g) || []).length]);
}

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
