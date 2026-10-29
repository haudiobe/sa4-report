/**
 * The Disposition Summary and Close Report (after T-2026.10.15) --
 * docs/DISPOSITION_SUMMARY.md.
 *
 * A report gets, when the user asks for it, a section at its end that lists
 * every TDoc of the report in one of three tables by its Status: agreed and
 * endorsed, otherwise disposed, not yet disposed. Close Report writes that
 * summary and says whether the report is ready to close: no TDoc that is not
 * yet disposed, and no TDoc the report has as revised whose revision 3GU (the
 * TDoc list) does not have.
 *
 * These tests run the REAL build, the REAL update, the REAL status dropdowns
 * (fake Docs API) and the REAL TDoc-list download (fake sheet) of
 * tests/helpers/template-report.js, on a document that keeps a paragraph
 * between two tables and its last paragraph, as Google Docs does.
 *
 *   A.  the classification of every status, and of a status it does not know
 *   B.  the summary: every TDoc once, counts, order, links, title, source, status
 *   C.  repeated: idempotent, one summary, no empty paragraphs gained,
 *       the rest of the report unchanged, and text typed into it never removed
 *   D.  Close Report: what blocks it, what it lists, what it changes (only the summary)
 *   E.  revised in the report, not in 3GU
 *   F.  where the summary stands, and that the agenda, the close and the
 *       Attendance section stay before it (ad-hoc and main meeting, with and
 *       without close and Attendance, text typed by hand at the end)
 *   G.  the MBS report (tests/fixtures/mbs-report-structure-damaged.json) --
 *       the three tables as Close Report leaves them
 *   H.  backward safety: a report without a summary is never given one;
 *       Build Report from Scratch
 *   I.  what changed in Code.js and ReportCreator.js, and nothing else
 *   J.  a summary an earlier candidate wrote (00a50c5, 4276161) is rewritten in the current form
 *
 * Run: node tests/disposition-summary-close-report.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { templateReport, zip } = require('./helpers/template-report.js');
const { codeWithoutDispositionSummary, creatorWithoutDispositionSummary, DISPOSITION_SECTION_START, DISPOSITION_SECTION_END, DISPOSITION_SECTION_FUNCTIONS, DISPOSITION_CHANGED_FUNCTIONS,
  DISPOSITION_CREATOR_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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

const ROOT = path.join(__dirname, '..');
const SUMMARY = 'Disposition Summary';
const ATTENDANCE = 'Attendance';
const TITLES = ['Agreed, Endorsed and Approved Documents', 'Other Disposed Documents', 'Documents Not Yet Disposed'];
// The fourteen options of the "Document Status" dropdown of the master template, in its order
// (docs/STATUS_DROPDOWN_RELEASE_CANDIDATE.md, the template seeding).
const OPTIONS = ['available', 'noted', 'agreed', 'revised', 'parked', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'other', 'replied', 'Plenary', 'postponed'];

const id = (n) => 'S4aA2691' + String(n).padStart(2, '0');
const t = (n, agenda, status, extra) => Object.assign({ id: id(n), agenda: agenda, status: status || 'available', uploaded: true }, extra || {});
const isHeading = (c) => c.getType() === 'PARAGRAPH' && c.getHeading() !== 'NORMAL';
const isEmpty = (c) => c.getType() === 'PARAGRAPH' && c.getHeading() === 'NORMAL' && c.getText() === '';
const isTdocTable = (c) => c.getType() === 'TABLE' && c.getNumRows() > 0 && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc';
const isSummaryTable = (c) => c.getType() === 'TABLE' && c.getNumRows() > 0 && c.getRow(0).getNumCells() === 4 && c.getRow(0).getCell(1).getText() === 'Title' && c.getRow(0).getCell(3).getText() === 'Status';
const cellLink = (cell) => cell.editAsText().getLinkUrl(0) || '';
/** The whole document, element by element: what a reader sees, and the links. */
const snapshot = (body) => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE'
  ? ['T', Array.from({ length: c.getNumRows() }, (x, r) => Array.from({ length: c.getRow(r).getNumCells() }, (y, k) => [c.getRow(r).getCell(k).getText(), c.getRow(r).getCell(k)._dropdown ? c.getRow(r).getCell(k)._dropdown.optionId : null, cellLink(c.getRow(r).getCell(k))]))]
  : ['P', c.getHeading(), c.getText(), !!c._bold])));
/**
 * The document without what the summary owns (its heading to the end of what it wrote): everything a summary must leave
 * as it is, prose typed below it included. The option selected in a dropdown is left out -- the tests pick decisions;
 * what a status shows is checked by shownAll(). So are the empty paragraphs that end the document (Google Docs keeps one
 * after a table that ends it).
 */
const outsideSummary = (r) => {
  const s = r.s.findDispositionSummarySection_(r.body);
  let all = JSON.parse(snapshot(r.body)).map((e) => (e[0] === 'T' ? ['T', e[1].map((row) => row.map((cell) => [cell[0], cell[2]]))] : e));
  if (s) all = all.slice(0, s.start).concat(all.slice(s.contentEnd));
  while (all.length && all[all.length - 1][0] === 'P' && all[all.length - 1][1] === 'NORMAL' && all[all.length - 1][2] === '') all.pop();
  return JSON.stringify(all);
};
/** Headings and TDoc numbers, in order; a summary table as '[summary n]'. */
const outline = (body) => body._children.map((c) => (isTdocTable(c) ? c.getRow(0).getCell(1).getText() : isSummaryTable(c) ? '[summary ' + (c.getNumRows() - 1) + ']' : isHeading(c) ? '# ' + c.getText() : null)).filter(Boolean);
/** The summary as it reads: [title line, [[tdoc, title, source, status, link]] | 'None.'] per class. */
function summaryOf(body) {
  const at = body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY);
  if (at === -1) return null;
  const out = [];
  for (let i = at + 1; i < body._children.length && !isHeading(body._children[i]); i++) {
    const c = body._children[i];
    if (c.getType() === 'PARAGRAPH' && /\(\d+\)$/.test(c.getText())) out.push([c.getText()]);
    else if (c.getType() === 'PARAGRAPH' && c.getText() === 'None.') out[out.length - 1].push('None.');
    else if (isSummaryTable(c)) out[out.length - 1].push(Array.from({ length: c.getNumRows() - 1 }, (x, r) => [0, 1, 2, 3].map((k) => c.getRow(r + 1).getCell(k).getText()).concat([cellLink(c.getRow(r + 1).getCell(0))])));
  }
  return out;
}
const summaryHeadings = (body) => body._children.filter((c) => isHeading(c) && c.getText() === SUMMARY).length;
const empties = (body) => body._children.filter(isEmpty).length;
/** Everything after the summary heading that is not the summary: agenda headings, TDoc tables, the Attendance heading. */
const afterSummary = (body) => {
  const at = body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY);
  return at === -1 ? ['(no summary)'] : body._children.slice(at + 1).map((c) => (isTdocTable(c) ? c.getRow(0).getCell(1).getText() : isHeading(c) ? '# ' + c.getText() : null)).filter(Boolean);
};

/** A built template report with status dropdowns; `decisions` { n: option } are picked in the dropdowns by hand. */
function built(tdocs, decisions, options) {
  const r = templateReport(Object.assign({ tdocs: tdocs, tableSeparators: true, endParagraph: true }, options || {}));
  const b = r.build();
  if (!b.ok) throw new Error('build failed: ' + b.error);
  // The build writes the TDocs of the agenda it was built from; the first update adds those of the closing item.
  const u = r.update();
  if (!u.success) throw new Error('update failed: ' + u.error);
  decide(r, decisions || {});
  return r;
}
function decide(r, decisions) {
  r.user(() => Object.keys(decisions).forEach((n) => r.api.pick(id(n), decisions[n])));
}
const updateSummary = (r) => r.exec(() => r.s.updateDispositionSummary());
const closeReport = (r) => r.exec(() => r.s.closeReport());
const lastAlert = (r) => r.ui.alerts[r.ui.alerts.length - 1];
const shownAll = (r) => r.order().map((n) => n + '=' + r.shown(n));

// ====================================================================== A
console.log('A. the classification of the statuses');
{
  const { sandbox: s } = loadCode();
  const cls = (status) => s.classifyTdocDisposition_(status).disposition;
  check('agreed, endorsed and approved: AGREED_ENDORSED_APPROVED, as three statuses', ['agreed', 'endorsed', 'approved'].map(cls), Array(3).fill('AGREED_ENDORSED_APPROVED'));
  check('approved is in the first class, and not other disposed', [cls('approved'), cls('Approved'), cls('approved') === 'DISPOSED_OTHER'], ['AGREED_ENDORSED_APPROVED', 'AGREED_ENDORSED_APPROVED', false]);
  check('noted, revised, merged, withdrawn, replied: DISPOSED_OTHER',
    ['noted', 'revised', 'merged', 'withdrawn', 'replied'].map(cls), Array(5).fill('DISPOSED_OTHER'));
  check('every option is in exactly one class: 3 + 5 + 6 = 14', ['AGREED_ENDORSED_APPROVED', 'DISPOSED_OTHER', 'UNRESOLVED'].map((k) => OPTIONS.filter((o) => cls(o) === k).length), [3, 5, 6]);
  check('available, reserved, parked, postponed, Plenary, other: UNRESOLVED', ['available', 'reserved', 'parked', 'postponed', 'Plenary', 'other'].map(cls), Array(6).fill('UNRESOLVED'));
  check('postponed: no final decision at this meeting -- not yet disposed, a known status', s.classifyTdocDisposition_('Postponed'), { label: 'postponed', known: true, disposition: 'UNRESOLVED' });
  check('every one of the fourteen options of the template is classified', OPTIONS.filter((o) => !s.classifyTdocDisposition_(o).known), []);
  const constant = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '').match(/const DISPOSITION_OF_STATUS_ = \{([\s\S]*?)\};/)[1];
  check('DISPOSITION_OF_STATUS_ names exactly the options of the template, by their label',
    (constant.match(/'([a-z ]+)':/g) || []).map((m) => m.slice(1, -2)).sort(), OPTIONS.map((o) => o.toLowerCase()).sort());
  const helper = fs.readFileSync(path.join(__dirname, 'helpers', 'template-report.js'), 'utf8');
  check('the options are those of the template definition the test report is built with', JSON.parse(helper.match(/const STATUS_NAMES = (\[[^\]]*\]);/)[1].replace(/'/g, '"')), OPTIONS);
  check('case, outer and inner whitespace do not matter; the Portal wording "Replied to" is replied',
    ['  AGREED ', 'Endorsed', 'Replied to', 'replied  to', 'PLENARY'].map((x) => s.classifyTdocDisposition_(x)),
    [{ label: 'agreed', known: true, disposition: 'AGREED_ENDORSED_APPROVED' }, { label: 'endorsed', known: true, disposition: 'AGREED_ENDORSED_APPROVED' }, { label: 'replied', known: true, disposition: 'DISPOSED_OTHER' },
      { label: 'replied', known: true, disposition: 'DISPOSED_OTHER' }, { label: 'plenary', known: true, disposition: 'UNRESOLVED' }]);
  check('a status it does not know is UNRESOLVED and not known: fail safe, never disposed',
    ['agreed, see minutes', 'not pursued', 'rejected', 'technically endorsed', '', null, undefined].map((x) => [cls(x), s.classifyTdocDisposition_(x).known]), Array(7).fill(['UNRESOLVED', false]));
  check('the three classes in the order of the summary', JSON.parse(JSON.stringify(s.buildDispositionSummaryModel_([]))).map((c) => [c.key, c.title]),
    [['AGREED_ENDORSED_APPROVED', TITLES[0]], ['DISPOSED_OTHER', TITLES[1]], ['UNRESOLVED', TITLES[2]]]);
}

// ====================================================================== B
console.log('B. the summary: every TDoc once, counts, order, links, title, source, status');
{
  // 4.3: 01 agreed, 02 revised -> 03 (agreed), 04 parked; 4.4: 05 endorsed, 06 noted (not uploaded: no link), 07 available;
  // 4.5: 08 withdrawn, 09 reserved, 10 merged; 5: 11 approved.
  const tdocs = [t(1, '4.3'), t(2, '4.3', 'revised', { revisedTo: id(3) }), t(3, '4.3'), t(4, '4.3'), t(5, '4.4'), t(6, '4.4', 'available', { uploaded: false }), t(7, '4.4'),
    t(8, '4.5'), t(9, '4.5', 'reserved', { uploaded: false }), t(10, '4.5'), t(11, '5')];
  const r = built(tdocs, { 1: 'agreed', 3: 'agreed', 4: 'parked', 5: 'endorsed', 6: 'noted', 8: 'withdrawn', 10: 'merged', 11: 'approved' });
  const order = r.order();
  check('the report: eleven TDoc tables, 03 below 02', order, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11].map(id));
  const statusSummary = () => { r.exec(() => r.s.analyzeReportStatus()); const text = lastAlert(r)[1]; r.ui.alerts.pop(); return text; };
  const statusSummaryBefore = statusSummary();
  updateSummary(r);
  check('Report Status Summary says what it said before: the summary tables are not counted as TDocs', statusSummary(), statusSummaryBefore);
  const sum = summaryOf(r.body);
  const row = (n, status, link) => [id(n), 'Synthetic title of ' + id(n), 'ExampleCorp', status, link === false ? '' : zip(id(n))];
  check('three tables with counts; the TDocs in the order of the report; title, source, status as in their tables; the TDoc linked as in its table', sum, [
    [TITLES[0] + ' (4)', [row(1, 'agreed'), row(3, 'agreed'), row(5, 'endorsed'), row(11, 'approved')]],
    [TITLES[1] + ' (4)', [row(2, 'revised'), row(6, 'noted', false), row(8, 'withdrawn'), row(10, 'merged')]],
    [TITLES[2] + ' (3)', [row(4, 'parked'), row(7, 'available'), row(9, 'reserved', false)]]]);
  const listed = sum.reduce((all, c) => all.concat(c[1].map((x) => x[0])), []);
  check('every TDoc of the report exactly once', [listed.slice().sort(), listed.length], [order.slice().sort(), order.length]);
  check('Agreed, Endorsed and Approved keep their own status', sum[0][1].map((x) => x[3]), ['agreed', 'agreed', 'endorsed', 'approved']);
  check('the approved TDoc is in the first table and in no other', sum.map((c) => c[1].some((x) => x[0] === id(11))), [true, false, false]);
  check('the three tables are mutually exclusive: no TDoc in two of them', listed.filter((n, i) => listed.indexOf(n) !== i), []);
  check('the dialog says what was done and the counts', lastAlert(r).slice(0, 2), [SUMMARY, 'Disposition Summary written at the end of the report.\n\n11 TDocs: 4 agreed/endorsed/approved, 4 other disposed, 3 not yet disposed.']);
  const at = r.body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY);
  const heading = r.body._children[at];
  const lines = r.body._children.slice(at + 1).filter((c) => c.getType() === 'PARAGRAPH' && c.getText());
  check('the heading is a level-one heading; the class lines are bold, the note is not', [heading.getHeading(), lines.map((p) => [p.getText().slice(0, 18), !!p._bold])],
    ['H1', [['Written from the S', false], [TITLES[0].slice(0, 18), true], [TITLES[1].slice(0, 18), true], [TITLES[2].slice(0, 18), true]]]);
  const table = r.body._children.filter(isSummaryTable)[0];
  check('a summary table: header row bold and shaded, the rows not bold (although the line before it is bold)',
    [[0, 1, 2, 3].map((k) => table.getRow(0).getCell(k)._bold), [0, 1, 2, 3].map((k) => table.getRow(1).getCell(k)._bold), table.getRow(0).getCell(0)._background], [[true, true, true, true], [false, false, false, false], '#D9EAF7']);
  check('a summary table is not a TDoc table to the readers of TDoc tables (its TDoc cell is "Title")', r.exec(() => r.s.detectTdocTablesInDocument_(r.body).map((f) => f.tdoc)), order);

  // A class without TDocs: "None.".
  const r2 = built([t(1, '4.3'), t(2, '4.3')], { 1: 'agreed', 2: 'noted' });
  updateSummary(r2);
  check('a class without TDocs says "None."', summaryOf(r2.body).map((c) => [c[0], c[1] === 'None.' ? 'None.' : c[1].length]), [[TITLES[0] + ' (1)', 1], [TITLES[1] + ' (1)', 1], [TITLES[2] + ' (0)', 'None.']]);

  // Statuses as text (a report without status dropdowns), a status the classification does not know, and two tables of one TDoc.
  const r3 = templateReport({ tdocs: [t(1, '4.3'), t(2, '4.3'), t(3, '4.4')], tableSeparators: true, release: false, docs: false });
  r3.build();
  const setStatus = (n, text, which) => r3.user(() => { const table = r3.tdocTables().filter((x) => x.getRow(0).getCell(1).getText() === id(n))[which || 0]; for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === 'Status') table.getRow(i).getCell(1).setText(text); });
  setStatus(1, 'Agreed');
  setStatus(2, 'agreed, see minutes');
  setStatus(3, 'noted');
  r3.user(() => { const copy = r3.table(id(3)).copy(); r3.body.insertTable(r3.body.getChildIndex(r3.table(id(3))) + 1, copy); });
  setStatus(3, 'endorsed', 1);
  updateSummary(r3);
  check('text statuses: as written; an unknown status is not yet disposed and says so; a TDoc with two tables of different statuses is listed once, not yet disposed',
    summaryOf(r3.body).map((c) => [c[0], c[1] === 'None.' ? 'None.' : c[1].map((x) => x[0] + ' ' + x[3])]),
    [[TITLES[0] + ' (1)', [id(1) + ' Agreed']], [TITLES[1] + ' (0)', 'None.'], [TITLES[2] + ' (2)', [id(2) + ' agreed, see minutes (not a known status)', id(3) + ' conflicting: noted / endorsed']]]);
}

// ====================================================================== C
console.log('C. repeated: idempotent, one summary, no empty paragraphs gained, the rest unchanged, text typed into it never removed');
{
  const tdocs = [t(1, '4.3'), t(2, '4.3'), t(3, '4.4'), t(4, '4.5'), t(5, '5')];
  const r = built(tdocs, { 1: 'agreed', 2: 'noted', 3: 'parked', 5: 'endorsed' });
  r.user(() => {
    // Minutes and a disposition typed by hand, and a sentence under the close.
    const table = r.table(id(1));
    for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === 'Minutes') table.getRow(i).getCell(1).setText('Presented by ExampleCorp. Agreed after discussion.');
    const close = r.body._children.findIndex((c) => isHeading(c) && /Close of the meeting/.test(c.getText()));
    r.body.insertParagraph(close + 1, 'The session was closed at 17:00.');
  });
  const before = outsideSummary(r);
  const statusesBefore = shownAll(r);
  updateSummary(r);
  const first = snapshot(r.body);
  const firstCount = r.body._children.length;
  const firstEmpties = empties(r.body);
  check('the first run writes it; the rest of the report is as it was', [lastAlert(r)[1].split('\n')[0], outsideSummary(r) === before], ['Disposition Summary written at the end of the report.', true]);
  for (let i = 0; i < 5; i++) updateSummary(r);
  check('five more runs: "already up to date", the document identical, element for element', [lastAlert(r)[1].split('\n')[0], snapshot(r.body) === first], ['Disposition Summary was already up to date.', true]);
  for (let i = 0; i < 5; i++) check('update ' + (i + 1) + ' completes and leaves the summary unchanged', r.update(), { success: true, error: null, dispositionSummary: 'unchanged' });
  check('after five updates: the document identical, one summary, no element and no empty paragraph more', [snapshot(r.body) === first, summaryHeadings(r.body), r.body._children.length - firstCount, empties(r.body) - firstEmpties],
    [true, 1, 0, 0]);

  // A decision changes: the summary follows, and goes on following; nothing grows.
  decide(r, { 3: 'agreed' });
  check('an update after a decision rewrites the summary', r.update(), { success: true, error: null, dispositionSummary: 'replaced' });
  check('03 moved to the agreed table', summaryOf(r.body).map((c) => c[0]), [TITLES[0] + ' (3)', TITLES[1] + ' (1)', TITLES[2] + ' (1)']);
  for (let k = 0; k < 6; k++) { decide(r, { 4: k % 2 ? 'available' : 'noted' }); r.update(); updateSummary(r); }
  check('six more changes back and forth: one summary, no element and no empty paragraph more, the rest of the report as it was',
    [summaryHeadings(r.body), r.body._children.length - firstCount, empties(r.body) - firstEmpties, outsideSummary(r) === before], [1, 0, 0, true]);
  check('no status was changed by any of it (only the picks of the test)', shownAll(r), statusesBefore.map((x) => x.replace(id(3) + '=dropdown:parked', id(3) + '=dropdown:agreed').replace(id(4) + '=dropdown:available', id(4) + '=dropdown:available')));

  // Text typed INTO the summary (between its note and its first table) is never removed: nothing is written, and the user is told.
  r.user(() => { const at = r.body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY); r.body.insertParagraph(at + 2, 'Note of the chair: 4.5 to be continued.'); });
  decide(r, { 4: 'agreed' });
  const typed = snapshot(r.body);
  updateSummary(r);
  check('typed text in the section: Update Disposition Summary writes nothing and says why',
    [snapshot(r.body) === typed, /not updated: text was typed into it \("Note of the chair: 4.5 to be continued."\), above a part of the summary/.test(lastAlert(r)[1])], [true, true]);
  check('... and so does an update, which completes', [r.update(), snapshot(r.body) === typed, r.logs.filter((l) => /^Disposition Summary: not updated/.test(l)).length > 0],
    [{ success: true, error: null, dispositionSummary: 'refused' }, true, true]);
  r.user(() => { const p = r.body._children.filter((c) => c.getType() === 'PARAGRAPH' && c.getText() === 'Note of the chair: 4.5 to be continued.')[0]; p.removeFromParent(); });
  updateSummary(r);
  check('once the text is moved out, it is written again', [lastAlert(r)[1].split('\n')[0], summaryOf(r.body)[0][0]], ['Disposition Summary updated.', TITLES[0] + ' (4)']);

  // A status that cannot be read: nothing is written.
  const r2 = built([t(1, '4.3'), t(2, '4.3')], { 1: 'agreed', 2: 'noted' });
  updateSummary(r2);
  decide(r2, { 2: 'agreed' });
  const kept = snapshot(r2.body);
  r2.api.fail.get = () => 'Service unavailable';
  updateSummary(r2);
  check('a dropdown status that cannot be read: nothing written, the user is told which', [snapshot(r2.body) === kept, lastAlert(r2)[1]],
    [true, 'The Disposition Summary was not written: the Status of ' + id(1) + ', ' + id(2) + ' could not be read. Try again in a minute.']);
}

// ====================================================================== D
console.log('D. Close Report: what blocks it, what it lists, and that it changes nothing but the summary');
{
  const clean = built([t(1, '4.3'), t(2, '4.3', 'revised', { revisedTo: id(3) }), t(3, '4.3'), t(4, '4.4'), t(5, '5')], { 1: 'agreed', 3: 'endorsed', 4: 'withdrawn', 5: 'noted' });
  const statuses = shownAll(clean);
  const tdocList = JSON.stringify(clean.tdocs);
  const writesBefore = clean.api.count('batchUpdate');
  const result = closeReport(clean);
  check('a clean report: ready to close', [result.ready, lastAlert(clean)[0]], [true, 'Report ready to close']);
  check('the dialog: counts and the summary', lastAlert(clean)[1],
    'Report ready to close.\n\n5 TDocs checked.\n2 agreed/endorsed/approved.\n3 other disposed.\n0 not yet disposed.\nRevision not reserved in 3GU: 0.\nRevision reserved but not uploaded in 3GU: 0.\n' +
    'Revision state could not be verified in 3GU: 0.\n\nDisposition Summary written at the end of the report.');
  check('it wrote the summary, and changed no status, nothing in 3GU and nothing through the Docs API', [summaryHeadings(clean.body), shownAll(clean), JSON.stringify(clean.tdocs) === tdocList, clean.api.count('batchUpdate') - writesBefore],
    [1, statuses, true, 0]);
  check('it stored nothing', Object.keys(clean.docProps._store).filter((k) => /CLOSE|DISPOSITION|CLOSED/i.test(k)), []);
  closeReport(clean);
  check('run again: ready, the summary already up to date', [lastAlert(clean)[0], lastAlert(clean)[1].split('\n').slice(-1)[0], summaryHeadings(clean.body)], ['Report ready to close', 'Disposition Summary was already up to date.', 1]);

  const approvedOnly = built([t(1, '4.3'), t(2, '4.3')], { 1: 'approved', 2: 'agreed' });
  const ares = closeReport(approvedOnly);
  check('approved does not block: ready, counted with agreed/endorsed', [ares.ready, lastAlert(approvedOnly)[1].split('\n').slice(2, 6)],
    [true, ['2 TDocs checked.', '2 agreed/endorsed/approved.', '0 other disposed.', '0 not yet disposed.']]);

  ['available', 'reserved', 'parked', 'postponed'].forEach((status) => {
    const r = built([t(1, '4.3'), t(2, '4.3')], { 1: 'agreed', 2: status });
    const res = closeReport(r);
    check(status + ' blocks; the TDoc is listed with title, source and status', [res.ready, lastAlert(r)[0], lastAlert(r)[1].split('\n').filter((l) => /^• /.test(l))],
      [false, 'Report NOT ready to close', ['• ' + id(2) + ' | Synthetic title of ' + id(2) + ' | ExampleCorp | ' + status]]);
    check(status + ': the summary is written all the same, and the status is not changed', [summaryOf(r.body)[2][0], r.shown(id(2))], [TITLES[2] + ' (1)', 'dropdown:' + status]);
  });

  const several = built([t(1, '4.3'), t(2, '4.3'), t(3, '4.4'), t(4, '4.4'), t(5, '4.5')], { 1: 'available', 2: 'Plenary', 3: 'agreed', 4: 'other', 5: 'reserved' });
  const res = closeReport(several);
  const text = lastAlert(several)[1];
  check('several not yet disposed: every one is listed, in the order of the report', [res.ready, res.unresolved.map((e) => e.tdoc), text.split('\n').filter((l) => /^• /.test(l)).map((l) => l.split(' | ')[0].slice(2))],
    [false, [id(1), id(2), id(4), id(5)], [id(1), id(2), id(4), id(5)]]);
  check('the dialog says NOT READY TO CLOSE, the counts, and what to do', [text.split('\n')[0], text.split('\n').slice(2, 9), /NOT YET DISPOSED \(4\) -- TDoc \| Title \| Source \| Current Status:/.test(text),
    text.split('\n').slice(-1)[0]],
    ['Report NOT READY TO CLOSE.', ['5 TDocs checked.', '1 agreed/endorsed/approved.', '0 other disposed.', '4 not yet disposed.', 'Revision not reserved in 3GU: 0.', 'Revision reserved but not uploaded in 3GU: 0.',
      'Revision state could not be verified in 3GU: 0.'], true,
      'Nothing was changed but the Disposition Summary. Correct the TDocs above and run Close Report again.']);

  const unknown = templateReport({ tdocs: [t(1, '4.3'), t(2, '4.3')], tableSeparators: true, release: false, docs: false });
  unknown.build();
  unknown.user(() => ['Agreed', 'Agreed with changes'].forEach((status, k) => { const table = unknown.table(id(k + 1)); for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === 'Status') table.getRow(i).getCell(1).setText(status); }));
  const u = closeReport(unknown);
  check('a status the classification does not know blocks, and is named as such', [u.ready, lastAlert(unknown)[1].split('\n').filter((l) => /^• /.test(l))],
    [false, ['• ' + id(2) + ' | Synthetic title of ' + id(2) + ' | ExampleCorp | Agreed with changes (not a known status)']]);

  // Many: the dialog shows forty and points to the summary, which has all.
  const many = built(Array.from({ length: 45 }, (x, i) => t(i + 1, '4.3')), {});
  const m = closeReport(many);
  const ml = lastAlert(many)[1].split('\n');
  check('45 not yet disposed: the dialog lists forty and points to the summary; the summary lists all',
    [m.unresolved.length, ml.filter((l) => /^• /.test(l)).length, ml.filter((l) => /^… and 5 more: see "Documents Not Yet Disposed" in the Disposition Summary\.$/.test(l)).length, summaryOf(many.body)[2][1].length], [45, 40, 1, 45]);

  // The TDoc list is read only when the report has a TDoc as revised: without one, a list that cannot be read does not matter.
  const norevision = built([t(1, '4.3')], { 1: 'agreed' });
  norevision.s.SpreadsheetApp = { open: () => { throw new Error('Spreadsheet conversion failed'); } };
  const listFetches = (r, from) => r.fetches.slice(from).filter((f) => /GenerateDocumentList/.test(f)).length;
  const nf = norevision.fetches.length;
  check('no TDoc revised in the report: the TDoc list is not read, and the report is ready', [closeReport(norevision).ready, listFetches(norevision, nf)], [true, 0]);
  // What cannot be checked blocks: the TDoc list cannot be read, with a TDoc revised in the report.
  const broken = built([t(1, '4.3', 'revised', { revisedTo: id(2) }), t(2, '4.3')], { 2: 'agreed' });
  broken.s.SpreadsheetApp = { open: () => { throw new Error('Spreadsheet conversion failed'); } };
  const bres = closeReport(broken);
  check('the TDoc list cannot be read: not ready, the summary written, the revised TDoc "could not be verified", with the reason',
    [bres.ready, summaryHeadings(broken.body), bres.revisions.map((x) => [x.tdoc, x.kind, x.reason]), lastAlert(broken)[1].split('\n').filter((l) => /^(REVISION STATE|   → )/.test(l))],
    [false, 1, [[id(1), 'revision-not-verified', 'the TDoc list could not be read: Spreadsheet conversion failed']],
      ['REVISION STATE COULD NOT BE VERIFIED IN 3GU (1) -- Original TDoc | Revision | Source | Report state | 3GU state:', '   → Check the TDoc and its revision in 3GU. (The TDoc list could not be read: Spreadsheet conversion failed.)']]);
  // A Drive copy of a report, or Resolve without Save: no saved meeting (the configuration guard). The TDoc list is not read --
  // never the built-in SA4#136 defaults -- and the revision is not verified.
  const unsaved = built([t(1, '4.3', 'revised', { revisedTo: id(2) }), t(2, '4.3')], { 2: 'agreed' });
  delete unsaved.docProps._store.MEETING_TYPE;
  delete unsaved.docProps._store.REPORT_SUFFIX;
  const fetchesBefore = unsaved.fetches.length;
  const ures = closeReport(unsaved);
  check('a report with no saved meeting: not ready, the revision not verified, and no TDoc list read',
    [ures.ready, ures.revisions.map((x) => x.kind), /^the TDoc list could not be read: /.test(ures.revisions[0].reason), listFetches(unsaved, fetchesBefore)],
    [false, ['revision-not-verified'], true, 0]);
  const unreadable = built([t(1, '4.3')], { 1: 'agreed' });
  unreadable.api.fail.get = () => 'Service unavailable';
  const nres = closeReport(unreadable);
  check('a status that cannot be read: not ready, no summary written', [nres.ready, summaryHeadings(unreadable.body), /could not be read, so the report could not be checked/.test(lastAlert(unreadable)[1])], [false, 0, true]);
  check('in the master template Close Report and Update Disposition Summary are refused', (() => {
    const m2 = built([t(1, '4.3')], {});
    m2.s.DocumentApp.getActiveDocument().getId = () => 'TEMPLATEdoc0000000000000000000000000000000';
    const errs = [];
    ['closeReport', 'updateDispositionSummary'].forEach((f) => { try { m2.exec(() => m2.s[f]()); errs.push('none'); } catch (e) { errs.push(/SA4 Report Template itself/.test(e.message)); } });
    return errs;
  })(), [true, true]);
}

// ====================================================================== E
console.log('E. revised in the report: not reserved, reserved but not uploaded, complete, not verified');
{
  const states = (r) => closeReport(r).revisions.map((x) => [x.tdoc, x.kind, x.state3gu, x.action]);
  const RESERVE = 'Author needs to reserve the revision in 3GU.';
  const UPLOAD = 'Author needs to upload the reserved revision in 3GU.';
  const VERIFY = 'Check the TDoc and its revision in 3GU.';
  // Not revised in the report, an ordinary TDoc list.
  const plain = built([t(1, '4.3'), t(2, '4.3')], { 1: 'agreed', 2: 'noted' });
  check('not revised in the report, an ordinary TDoc list: no warning', states(plain), []);

  // C. Complete: reserved (01 revised to 02 in the list) and uploaded (02 has its link).
  const proper = built([t(1, '4.3', 'revised', { revisedTo: id(2) }), t(2, '4.3')], { 2: 'agreed' });
  check('C. revision reserved and uploaded in 3GU: no warning, ready', [states(proper), closeReport(proper).ready], [[], true]);

  // A. Revised in the report (chosen in the dropdown), still available in 3GU: not reserved.
  const pending = built([t(1, '4.3'), t(2, '4.3')], { 1: 'revised', 2: 'agreed' });
  const pres = closeReport(pending);
  check('A. revised in the report, available in 3GU, no "Revised to": "Revision not reserved in 3GU.", reserve it', [pres.ready, pres.revisions.map((x) => [x.tdoc, x.kind, x.label, x.reportStatus, x.state3gu, x.action])],
    [false, [[id(1), 'revision-not-reserved', 'Revision not reserved in 3GU.', 'revised', 'available', RESERVE]]]);
  check('A. the dialog: its own list, Original TDoc | Revision | Source | Report state | 3GU state, and the action', lastAlert(pending)[1].split('\n').filter((l) => /^(• |   → |REVISION )/.test(l)),
    ['REVISION NOT RESERVED IN 3GU (1) -- Original TDoc | Revision | Source | Report state | 3GU state:', '• ' + id(1) + ' | – | ExampleCorp | revised | available', '   → ' + RESERVE]);
  check('A. a revised TDoc counts as disposed; the counts of each revision state', lastAlert(pending)[1].split('\n').slice(2, 9),
    ['2 TDocs checked.', '1 agreed/endorsed/approved.', '1 other disposed.', '0 not yet disposed.', 'Revision not reserved in 3GU: 1.', 'Revision reserved but not uploaded in 3GU: 0.', 'Revision state could not be verified in 3GU: 0.']);
  // The report knows the revision (named in its Disposition by hand): shown, but it decides nothing.
  pending.user(() => { const table = pending.table(id(1)); for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === 'Disposition') table.getRow(i).getCell(1).setText('Revised to ' + id(9) + ' (to be reserved)'); });
  check('A. a revision the report names in its Disposition is shown; the state is the same (3GU decides)', closeReport(pending).revisions.map((x) => [x.revision, x.kind]), [[id(9), 'revision-not-reserved']]);
  // A reserved original (not even uploaded itself) with no revision: not reserved either.
  const reservedOriginal = built([t(1, '4.3', 'reserved', { uploaded: false })], { 1: 'revised' });
  check('A. revised in the report, reserved in 3GU without "Revised to": not reserved', states(reservedOriginal).map((x) => [x[1], x[2]]), [['revision-not-reserved', 'reserved']]);

  // B. Reserved (Revised to 02), but 02 not uploaded: never called "not reserved".
  const notUploaded = built([t(1, '4.3', 'revised', { revisedTo: id(2) }), t(2, '4.3', 'reserved', { uploaded: false })], {});
  const nres = closeReport(notUploaded);
  check('B. revised to a revision that is reserved but not uploaded: "Revision reserved but not uploaded in 3GU.", upload it (and the revision is not yet disposed)',
    [nres.ready, nres.revisions.map((x) => [x.tdoc, x.revision, x.kind, x.label, x.action]), nres.unresolved.map((e) => e.tdoc + ' ' + e.status)],
    [false, [[id(1), id(2), 'revision-not-uploaded', 'Revision reserved but not uploaded in 3GU.', UPLOAD]], [id(2) + ' reserved']]);
  check('B. the dialog: its own list, and not the "not reserved" one', lastAlert(notUploaded)[1].split('\n').filter((l) => /^(REVISION |   → )/.test(l)),
    ['REVISION RESERVED BUT NOT UPLOADED IN 3GU (1) -- Original TDoc | Revision | Source | Report state | 3GU state:', '   → ' + UPLOAD]);
  notUploaded.tdoc(id(2)).uploaded = true;
  notUploaded.tdoc(id(2)).status = 'available';
  check('B -> C. once it is uploaded: no warning', states(notUploaded), []);

  // D. Cannot be verified: never clean.
  const noTarget = built([t(1, '4.3', 'revised')], {});
  check('D. 3GU has it as revised but names no revision: not verified', states(noTarget), [[id(1), 'revision-not-verified', 'revised, no "Revised to"', VERIFY]]);
  const otherDecision = built([t(1, '4.3', 'agreed')], { 1: 'revised' });
  check('D. 3GU has another decision for it (agreed): not verified, not "not reserved"', states(otherDecision).map((x) => [x[1], x[2]]), [['revision-not-verified', 'agreed']]);
  const lost = built([t(1, '4.3', 'revised', { revisedTo: id(2) }), t(2, '4.3')], { 2: 'agreed' });
  lost.tdocs = lost.tdocs.filter((x) => x.id !== id(2));
  check('D. its "Revised to" is not in the TDoc list (another meeting, or removed): not verified', closeReport(lost).revisions.map((x) => [x.revision, x.kind, x.reason]),
    [[id(2), 'revision-not-verified', 'the revision ' + id(2) + ' is not in the TDoc list']]);
  check('D. the dialog: its own list, with the reason', lastAlert(lost)[1].split('\n').filter((l) => /^(REVISION |   → )/.test(l)),
    ['REVISION STATE COULD NOT BE VERIFIED IN 3GU (1) -- Original TDoc | Revision | Source | Report state | 3GU state:', '   → ' + VERIFY + ' (The revision ' + id(2) + ' is not in the TDoc list.)']);

  // Several, of every state, and one the TDoc list does not have at all: each in its own list, in the order of the report.
  const multi = built([t(1, '4.3'), t(2, '4.3'), t(3, '4.4', 'revised', { revisedTo: id(5) }), t(4, '4.4'), t(5, '4.4', 'reserved', { uploaded: false }), t(6, '4.5')],
    { 1: 'revised', 2: 'revised', 4: 'revised', 5: 'agreed', 6: 'available' });
  multi.tdocs = multi.tdocs.filter((x) => x.id !== id(4));
  const mres = closeReport(multi);
  check('several: every one is reported with its own state', mres.revisions.map((x) => [x.tdoc, x.kind, x.state3gu]),
    [[id(1), 'revision-not-reserved', 'available'], [id(2), 'revision-not-reserved', 'available'], [id(3), 'revision-not-uploaded', 'revised, revised to ' + id(5) + ' (reserved, not uploaded)'], [id(4), 'revision-not-verified', 'not in the TDoc list']]);
  const mtext = lastAlert(multi)[1].split('\n');
  check('the dialog distinguishes the four: not yet disposed, not reserved, reserved but not uploaded, not verified -- counts and lists',
    [mtext.slice(5, 9), mtext.filter((l) => /^(NOT YET DISPOSED|REVISION )/.test(l)).map((l) => l.split(' -- ')[0])],
    [['1 not yet disposed.', 'Revision not reserved in 3GU: 2.', 'Revision reserved but not uploaded in 3GU: 1.', 'Revision state could not be verified in 3GU: 1.'],
      ['NOT YET DISPOSED (1)', 'REVISION NOT RESERVED IN 3GU (2)', 'REVISION RESERVED BUT NOT UPLOADED IN 3GU (1)', 'REVISION STATE COULD NOT BE VERIFIED IN 3GU (1)']]);

  // Nothing is changed: not the statuses, not the list, not a revision; nothing is reserved.
  const before = [shownAll(multi), JSON.stringify(multi.tdocs), multi.api.count('batchUpdate')];
  const fetchesBefore = multi.fetches.length;
  closeReport(multi);
  check('Close Report changes no status, nothing in the TDoc list, writes nothing through the API; it only downloads the list',
    [shownAll(multi), JSON.stringify(multi.tdocs), multi.api.count('batchUpdate'), multi.fetches.slice(fetchesBefore).map((f) => /GenerateDocumentList/.test(f))], before.concat([[true]]));

  // Pure: the same rules on rows of the list.
  const { sandbox: s } = loadCode();
  const row = (n, status, revisedTo, link) => ({ row: [id(n), status, revisedTo || ''], tdocCol: 0, statusCol: 1, revisedToCol: 2, richTextRow: [{ getLinkUrl: () => (link === false ? null : zip(id(n))) }] });
  const entry = (n, status) => Object.assign({ tdoc: id(n), title: 'T', source: 'S', status: status, revisionIds: [] }, s.classifyTdocDisposition_(status));
  const run = (entries, rows, problem) => JSON.parse(JSON.stringify(s.checkRevisionsIn3gu_(entries, rows, problem))).map((x) => [x.tdoc, x.revision, x.kind, x.state3gu, x.reason]);
  check('pure: only a TDoc the report has as revised is checked (agreed, noted, available in the report: never)',
    run([entry(1, 'agreed'), entry(2, 'noted'), entry(3, 'available')], [row(1, 'available'), row(2, 'available'), row(3, 'available')]), []);
  check('pure: C complete; A not reserved for available and reserved; D for revised without a revision, another decision, no status',
    run([entry(1, 'revised'), entry(2, 'Revised'), entry(3, 'revised'), entry(4, 'revised'), entry(5, 'revised'), entry(6, 'revised')],
      [row(1, 'revised', id(9)), row(9, 'available'), row(2, 'available'), row(3, 'reserved'), row(4, 'revised'), row(5, 'noted'), row(6, '')]),
    [[id(2), '', 'revision-not-reserved', 'available', ''], [id(3), '', 'revision-not-reserved', 'reserved', ''],
      [id(4), '', 'revision-not-verified', 'revised, no "Revised to"', 'the TDoc list has it as revised but names no revision'],
      [id(5), '', 'revision-not-verified', 'noted', 'the TDoc list has another decision for it'], [id(6), '', 'revision-not-verified', '(no status)', 'the TDoc list has no status for it']]);
  check('pure: B is the upload sign of the upload completion -- the list has no link on the revision; a link that cannot be read is D, not B',
    run([entry(1, 'revised'), entry(2, 'revised')], [row(1, 'revised', id(8)), row(8, 'available', '', false), row(2, 'revised', id(7)), Object.assign(row(7, 'reserved'), { richTextRow: null })]).map((x) => [x[0], x[1], x[2]]),
    [[id(1), id(8), 'revision-not-uploaded'], [id(2), id(7), 'revision-not-verified']]);
  check('pure: D when the list could not be read, has no "Revised to" column, lacks the TDoc or its revision, or its "Revised to" is not a TDoc number', [
    run([entry(1, 'revised')], null, 'the TDoc list could not be read: HTTP 500').map((x) => [x[2], x[4]]),
    run([entry(1, 'revised')], [Object.assign(row(1, 'available'), { revisedToCol: -1 })]).map((x) => [x[2], x[4]]),
    run([entry(1, 'revised')], [row(2, 'available')]).map((x) => [x[2], x[4]]),
    run([entry(1, 'revised')], [row(1, 'revised', id(3))]).map((x) => [x[2], x[4]]),
    run([entry(1, 'revised')], [row(1, 'revised', 'see e-mail')]).map((x) => [x[2], x[4]])],
  [[['revision-not-verified', 'the TDoc list could not be read: HTTP 500']], [['revision-not-verified', 'the TDoc list has no "Revised to" column']], [['revision-not-verified', 'the TDoc list of this report does not have it']],
    [['revision-not-verified', 'the revision ' + id(3) + ' is not in the TDoc list']], [['revision-not-verified', 'its "Revised to" is not a TDoc number']]]);
  check('pure: a TDoc with conflicting statuses is not taken for revised', run([Object.assign(entry(1, 'revised'), { conflict: true })], [row(1, 'available')]), []);
}

// ====================================================================== F
console.log('F. where the summary stands; the agenda, the close and the Attendance section stay before it');
{
  // Ad-hoc report with Attendance.
  const r = built([t(1, '4.3'), t(2, '4.4'), t(3, '5')], { 1: 'agreed', 2: 'noted', 3: 'noted' });
  r.user(() => {
    const b = r.body;
    b.appendParagraph(ATTENDANCE).setHeading('H3');
    b.appendParagraph('Statistics');
    b.appendParagraph('Attendees');
    b.appendTable([['Name', 'Company', 'Email'], ['Sam Rivera', 'ExampleCorp', 'sam.rivera@example.com']]);
    b.appendParagraph('');
  });
  updateSummary(r);
  const tail = (body) => outline(body).slice(outline(body).indexOf('# 5 Close of the meeting'));
  check('with Attendance: Close, its TDoc, Attendance, then the summary -- the end of the document', tail(r.body),
    ['# 5 Close of the meeting', id(3), '# ' + ATTENDANCE, '# ' + SUMMARY, '[summary 1]', '[summary 2]']);
  // New TDocs: of the closing item, of an item the report has no heading for, of an item before.
  r.tdocs.push(t(4, '5'), t(5, '4.6'), t(6, '4.3'));
  check('an update with three new TDocs completes and rewrites the summary', r.update(), { success: true, error: null, dispositionSummary: 'replaced' });
  check('the TDoc of the closing item is before the Attendance; the new item 4.6 before the close; nothing after the summary but the summary', [tail(r.body), afterSummary(r.body)],
    [['# 5 Close of the meeting', id(3), id(4), '# ' + ATTENDANCE, '# ' + SUMMARY, '[summary 1]', '[summary 2]', '[summary 3]'], []]);
  check('4.6 got its heading in the order of the agenda', outline(r.body).filter((x) => /^# 4\./.test(x)), ['# 4.3 Performance requirements', '# 4.4 Design constraints', '# 4.5 Other topics', '# 4.6 Synthetic item']);
  check('the summary has the new TDocs', summaryOf(r.body).map((c) => c[0]), [TITLES[0] + ' (1)', TITLES[1] + ' (2)', TITLES[2] + ' (3)']);
  // Refresh Attendance: replaced in place, before the summary.
  r.exec(() => r.s.renderAdhocAttendanceSection_(r.body, [{ statistics: [['Meeting', 'Synthetic meeting']], attendees: [{ name: 'Kim Doe', company: 'ExampleCorp', email: 'kim.doe@example.com' }] }]));
  check('Refresh Attendance replaces the section in place, before the summary', [tail(r.body), afterSummary(r.body)],
    [['# 5 Close of the meeting', id(3), id(4), '# ' + ATTENDANCE, '# ' + SUMMARY, '[summary 1]', '[summary 2]', '[summary 3]'], []]);

  // Ad-hoc report without Attendance; the Attendance is imported after the summary was written.
  const n = built([t(1, '4.3'), t(2, '5')], { 1: 'agreed', 2: 'noted' });
  updateSummary(n);
  check('without Attendance: Close, then the summary', tail(n.body), ['# 5 Close of the meeting', id(2), '# ' + SUMMARY, '[summary 1]', '[summary 1]']);
  n.exec(() => n.s.renderAdhocAttendanceSection_(n.body, [{ statistics: [['Meeting', 'Synthetic meeting']], attendees: [{ name: 'Kim Doe', company: 'ExampleCorp', email: 'kim.doe@example.com' }] }]));
  check('an Attendance section imported later goes into the closing section, before the summary', [tail(n.body), afterSummary(n.body)],
    [['# 5 Close of the meeting', id(2), '# ' + ATTENDANCE, '# ' + SUMMARY, '[summary 1]', '[summary 1]'], []]);
  n.tdocs.push(t(3, '5'));
  n.update();
  check('... and a new TDoc of the closing item before both', [tail(n.body), afterSummary(n.body)], [['# 5 Close of the meeting', id(2), id(3), '# ' + ATTENDANCE, '# ' + SUMMARY, '[summary 1]', '[summary 1]', '[summary 1]'], []]);

  // A report without a close and without Attendance: the summary ends the agenda part. (A build always writes the close
  // of its agenda; here it was deleted by hand.)
  const nc = built([t(1, '4.3'), t(2, '4.5')], { 1: 'agreed', 2: 'noted' });
  // It is the last paragraph of the document, which Google Docs does not let go: emptied instead.
  nc.user(() => { const close = nc.body._children.filter((c) => isHeading(c) && /Close of the meeting/.test(c.getText()))[0]; close.clear(); close.setHeading('NORMAL'); });
  updateSummary(nc);
  check('no close, no Attendance: the summary at the end', outline(nc.body).slice(-5), ['# 4.5 Other topics', id(2), '# ' + SUMMARY, '[summary 1]', '[summary 1]']);
  nc.tdocs.push(t(3, '4.5'), t(4, '4.6'), t(5, '7'));
  nc.update();
  check('new TDocs of the last item, of a new item and of a later item all go before the summary', [outline(nc.body).slice(outline(nc.body).indexOf('# 4.5 Other topics')), afterSummary(nc.body)],
    [['# 4.5 Other topics', id(2), id(3), '# 4.6 Synthetic item', id(4), '# 7 Synthetic item', id(5), '# ' + SUMMARY, '[summary 1]', '[summary 1]', '[summary 3]'], []]);
  nc.exec(() => nc.s.renderAdhocAttendanceSection_(nc.body, [{ statistics: [['Meeting', 'M']], attendees: [{ name: 'Kim Doe', company: 'ExampleCorp', email: 'kim.doe@example.com' }] }]));
  check('no close: an Attendance section goes before the summary too', [outline(nc.body).slice(-6, -3), afterSummary(nc.body)], [[id(5), '# ' + ATTENDANCE, '# ' + SUMMARY], []]);

  // Text typed by hand at the end: under the close it stays where it is; after a heading of its own after the summary, too.
  const h = built([t(1, '4.3'), t(2, '5')], { 1: 'agreed', 2: 'noted' });
  h.user(() => h.body.appendParagraph('Closing remarks typed by hand, after the last TDoc of the close.'));
  updateSummary(h);
  const texts = (body) => body._children.filter((c) => c.getType() === 'PARAGRAPH' && /typed by hand/.test(c.getText())).map((c) => c.getText());
  const indexOf = (body, re) => body._children.findIndex((c) => c.getType() === 'PARAGRAPH' && re.test(c.getText()));
  check('text typed at the end of the close stays where it is, before the summary', [texts(h.body), indexOf(h.body, /^Closing remarks/) < indexOf(h.body, /^Disposition Summary$/)],
    [['Closing remarks typed by hand, after the last TDoc of the close.'], true]);
  h.user(() => { h.body.appendParagraph('Annex: notes').setHeading('H1'); h.body.appendParagraph('Notes typed by hand in an annex of their own.'); });
  decide(h, { 2: 'agreed' });
  updateSummary(h);
  check('a heading of its own after the summary ends the summary section: its text is kept, the summary is updated before it',
    [lastAlert(h)[1].split('\n')[0], texts(h.body), outline(h.body).slice(-3)],
    ['Disposition Summary updated.', ['Closing remarks typed by hand, after the last TDoc of the close.', 'Notes typed by hand in an annex of their own.'], ['# ' + SUMMARY, '[summary 2]', '# Annex: notes']]);
  check('... and the summary lists both as agreed, "None." for the rest', summaryOf(h.body).map((c) => [c[0], c[1] === 'None.' ? 'None.' : c[1].length]), [[TITLES[0] + ' (2)', 2], [TITLES[1] + ' (0)', 'None.'], [TITLES[2] + ' (0)', 'None.']]);

  // Prose typed BELOW the summary, without a heading of its own: not the summary's, never touched; the summary is updated above it.
  const pr = built([t(1, '4.3'), t(2, '4.4'), t(3, '5')], { 1: 'agreed', 2: 'noted', 3: 'available' });
  updateSummary(pr);
  const PROSE = ['Action items after the meeting: ExampleCorp to circulate the updated draft.', '', 'Next session: to be announced on the list.'];
  pr.user(() => PROSE.forEach((text) => pr.body.appendParagraph(text)));
  const prose = (body) => { const at = body._children.findIndex((c) => c.getType() === 'PARAGRAPH' && c.getText() === PROSE[0]); return at === -1 ? null : body._children.slice(at, at + 3).map((c) => [c.getType(), c.getHeading(), c.getText()]); };
  const proseBefore = JSON.stringify(prose(pr.body));
  const proseOutside = outsideSummary(pr);
  check('prose typed below the summary is not the summary\'s: it ends with its last table', [pr.s.findDispositionSummarySection_(pr.body).complete, pr.s.findDispositionSummarySection_(pr.body).foreign, prose(pr.body) !== null], [true, [], true]);
  decide(pr, { 3: 'agreed' });
  updateSummary(pr);
  check('Update Disposition Summary: the summary is updated, the prose below it is unchanged and still below it',
    [lastAlert(pr)[1].split('\n')[0], summaryOf(pr.body).map((c) => c[0]), JSON.stringify(prose(pr.body)) === proseBefore, outsideSummary(pr) === proseOutside,
      pr.body._children.findIndex((c) => c.getType() === 'PARAGRAPH' && c.getText() === PROSE[0]) > pr.body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY)],
    ['Disposition Summary updated.', [TITLES[0] + ' (2)', TITLES[1] + ' (1)', TITLES[2] + ' (0)'], true, true, true]);
  decide(pr, { 2: 'available' });
  check('an update does the same', [pr.update().dispositionSummary, summaryOf(pr.body).map((c) => c[0]), JSON.stringify(prose(pr.body)) === proseBefore, outsideSummary(pr) === proseOutside],
    ['replaced', [TITLES[0] + ' (2)', TITLES[1] + ' (0)', TITLES[2] + ' (1)'], true, true]);
  for (let k = 0; k < 4; k++) { decide(pr, { 2: k % 2 ? 'available' : 'noted' }); pr.update(); updateSummary(pr); }
  check('four more changes: the prose is unchanged, nothing else changed, one summary, no empty paragraph more',
    [JSON.stringify(prose(pr.body)) === proseBefore, outsideSummary(pr) === proseOutside, summaryHeadings(pr.body)], [true, true, 1]);
  // Prose that looks like part of the summary ("None.") below it is not taken for it either: the summary ends with the table of its last class.
  pr.user(() => pr.body.appendParagraph('None.'));
  const withNone = outsideSummary(pr);
  decide(pr, { 2: 'noted' });
  updateSummary(pr);
  check('"None." typed below the summary is kept', [lastAlert(pr)[1].split('\n')[0], outsideSummary(pr) === withNone], ['Disposition Summary updated.', true]);
  // Prose typed BETWEEN two of its tables: the summary is not written, nothing is removed, the user is told.
  pr.user(() => { const tables = pr.body._children.filter(isSummaryTable); pr.body.insertParagraph(pr.body.getChildIndex(tables[0]) + 1, 'A remark typed between two tables of the summary.'); });
  decide(pr, { 2: 'available' });
  const between = snapshot(pr.body);
  updateSummary(pr);
  check('prose typed between two tables of the summary: nothing written, nothing removed, the user told what to move',
    [snapshot(pr.body) === between, /text was typed into it \("A remark typed between two tables of the summary\."\)/.test(lastAlert(pr)[1])], [true, true]);
  check('... and an update leaves it too', [pr.update().dispositionSummary, snapshot(pr.body) === between], ['refused', true]);
  // A table of the summary deleted by hand: the summary is written again whole; the prose below it is kept.
  const hd = built([t(1, '4.3'), t(2, '4.4')], { 1: 'agreed', 2: 'noted' });
  updateSummary(hd);
  hd.user(() => { hd.body.appendParagraph('Prose below the summary.'); hd.body._children.filter(isSummaryTable)[1].removeFromParent(); });
  const hdOutside = outsideSummary(hd);
  updateSummary(hd);
  check('a summary table deleted by hand: written again whole, the prose below kept', [lastAlert(hd)[1].split('\n')[0], summaryOf(hd.body).map((c) => [c[0], c[1] === 'None.' ? 'None.' : c[1].length]), outsideSummary(hd) === hdOutside],
    ['Disposition Summary updated.', [[TITLES[0] + ' (1)', 1], [TITLES[1] + ' (1)', 1], [TITLES[2] + ' (0)', 'None.']], true]);

  // A main-meeting report: the same rules (the SWG of a main meeting, its close a level-two heading).
  const { sandbox: s } = loadCode({ documentProperties: {} });
  s.styleStatusCell_ = () => {};
  s.Logger = { log: () => {} };
  s.DocumentApp.ElementType.UNSUPPORTED = 'UNSUPPORTED';
  const body = makeFakeDocumentBody(s, { tableSeparators: true });
  const tdocRow = (tdoc, agenda, status) => ({ row: [tdoc, 'Title of ' + tdoc, 'Source Co', agenda, status], richTextRow: null, tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 3, agendaTopicCol: -1, statusCol: 4, typeCol: -1, forCol: -1, revisedToCol: -1, agendaItem: agenda });
  [['9 Video SWG', 'H1'], ['9.1 Opening of the session', 'H2'], ['9.5 Item five', 'H2'], ['9.12 Close of the session', 'H2']].forEach(([text, heading]) => body.appendParagraph(text).setHeading(heading));
  body.appendParagraph('The session was closed.');
  [['S4-269001', '9.5', 'agreed'], ['S4-269002', '9.5', 'available'], ['S4-269003', '9.12', 'noted']].forEach(([n2, a, st]) => s.insertNewTdoc_(body, tdocRow(n2, a, st), null, null));
  s.writeDispositionSummary_(body);
  const mo = () => outline(body);
  check('main meeting: the summary after the close and its text', mo().slice(-6), ['# 9.12 Close of the session', 'S4-269003', '# ' + SUMMARY, '[summary 1]', '[summary 1]', '[summary 1]']);
  [['S4-269004', '9.12', 'noted'], ['S4-269005', '9.13', 'noted'], ['S4-269006', '9.6', 'noted'], ['S4-269007', '9.5', 'noted']].forEach(([n2, a, st]) => s.insertNewTdoc_(body, tdocRow(n2, a, st), null, null));
  s.writeDispositionSummary_(body);
  check('main meeting: new TDocs of the close, of 9.13, 9.6 and 9.5 all before the summary; 9.6 and 9.13 in the order of the agenda', [mo().filter((x) => x[0] === '#'), afterSummary(body)],
    [['# 9 Video SWG', '# 9.1 Opening of the session', '# 9.5 Item five', '# 9.6', '# 9.13', '# 9.12 Close of the session', '# ' + SUMMARY], []]);
  check('main meeting: the summary counts', summaryOf(body).map((c) => c[0]), [TITLES[0] + ' (1)', TITLES[1] + ' (5)', TITLES[2] + ' (1)']);
}

// ====================================================================== G
console.log('G. the MBS report: the three tables as Close Report leaves them');
{
  const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'mbs-report-structure-damaged.json'), 'utf8')).elements;
  const tdocsOfFixture = fixture.filter((e) => e.tdoc);
  const registration = fixture.filter((e) => e.registration)[0].registration;
  const revisedTo = {};
  tdocsOfFixture.forEach((e) => { if (e.revisedTo) revisedTo[e.tdoc] = e.revisedTo; });
  revisedTo.S4aI260080 = 'S4aI260091';
  // A finished meeting: the decisions as they could stand in the minutes (synthetic). Three TDocs still open (one of them
  // postponed), and one that is revised in the minutes although 3GU still has it as available (S4aI260089).
  const decisions = { S4aI260080: 'revised', S4aI260094: 'noted', S4aI260096: 'agreed', S4aI260098: 'endorsed', S4aI260099: 'agreed', S4aI260108: 'noted', S4aI260091: 'revised', S4aI260100: 'agreed',
    S4aI260089: 'revised', S4aI260082: 'revised', S4aI260102: 'revised', S4aI260105: 'agreed', S4aI260083: 'revised', S4aI260092: 'endorsed', S4aI260084: 'agreed', S4aI260085: 'noted', S4aI260086: 'merged',
    S4aI260087: 'agreed', S4aI260088: 'noted', S4aI260093: 'withdrawn', S4aI260101: 'available', S4aI260106: 'endorsed', S4aI260107: 'noted', S4aI260095: 'agreed', S4aI260097: 'postponed',
    S4aI260081: 'revised', S4aI260090: 'agreed', S4aI260103: 'parked', S4aI260104: 'noted' };
  const listStatus = (tdoc) => (revisedTo[tdoc] ? 'revised' : decisions[tdoc] === 'revised' ? 'available' : (tdocsOfFixture.filter((e) => e.tdoc === tdoc)[0] || {}).status);
  const list = registration.map(([tdoc, agenda]) => ({ id: tdoc, agenda: tdoc === 'S4aI260081' ? '3.7' : agenda, status: listStatus(tdoc), revisedTo: revisedTo[tdoc] || '', uploaded: true }));
  const r = templateReport({ tdocs: list, tableSeparators: true, props: { REPORT_SUFFIX: 'MBS' } });
  r.user(() => {
    const b = r.body;
    fixture.forEach((e) => {
      if (e.h) b.appendParagraph(e.text).setHeading(e.h);
      else if (e.p) b.appendParagraph(e.p);
      else if (e.empty) for (let i = 0; i < e.empty; i++) b.appendParagraph('');
      else if (e.tdoc) {
        b.appendTable(e.rows.map((label) => [label, label === 'TDoc' ? e.tdoc : label === 'Title' ? 'Synthetic title of ' + e.tdoc : label === 'Source' ? 'ExampleCorp' : label === 'Agenda Item' ? e.agenda
          : label === 'Status' ? decisions[e.tdoc] : label === 'Disposition' && e.revisedTo ? 'Revised to ' + e.revisedTo : label === 'Minutes' ? 'Minutes of ' + e.tdoc + ', typed by hand.' : '']));
      } else if (e.registration) b.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item']].concat(e.registration.map(([tdoc, agenda]) => [tdoc, 'Synthetic title of ' + tdoc, 'ExampleCorp', agenda])));
      else if (e.reallocations) b.appendTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason']].concat(e.reallocations));
      else if (e.attendees) b.appendTable([['Name', 'Company', 'Email']].concat(Array.from({ length: e.attendees }, (x, i) => ['Person ' + (i + 1), 'ExampleCorp', 'person' + (i + 1) + '@example.com'])));
    });
  });
  const b = r.body;
  check('the MBS report after its first update (the repair of T-2026.10.13): 29 TDocs, the close and the Attendance at the end', [r.update().success, outline(b).filter((x) => x[0] === '#').slice(-4)],
    [true, ['# 2.6 Synthetic item', '# 2.7 Synthetic item', '# 2.8 Close of the session', '# ' + ATTENDANCE]]);
  const minutesBefore = r.tdocTables().map((x) => r.field(x.getRow(0).getCell(1).getText(), 'Minutes'));
  const outsideBefore = outsideSummary(r);
  updateSummary(r);
  const res = closeReport(r);
  check('the summary is the last section: after "2.8 Close of the session" and the Attendance section; nothing after it',
    [outline(b).filter((x) => x[0] === '#').slice(-3), afterSummary(b)], [['# 2.8 Close of the session', '# ' + ATTENDANCE, '# ' + SUMMARY], []]);
  check('the report outside the summary is as it was: every minute typed by hand, the registration table, the attendees',
    [outsideBefore === outsideSummary(r), r.tdocTables().map((x) => r.field(x.getRow(0).getCell(1).getText(), 'Minutes'))], [true, minutesBefore]);
  check('29 TDocs, each once: 11 agreed/endorsed/approved, 15 other disposed, 3 not yet disposed', [res.checked, summaryOf(b).map((c) => c[0]), summaryOf(b).reduce((n, c) => n + (c[1] === 'None.' ? 0 : c[1].length), 0)],
    [29, [TITLES[0] + ' (11)', TITLES[1] + ' (15)', TITLES[2] + ' (3)'], 29]);
  check('not ready: S4aI260101 available, S4aI260097 postponed, S4aI260103 parked; S4aI260089 revised in the minutes, not reserved in 3GU; the five other revisions complete',
    [res.ready, res.unresolved.map((e) => e.tdoc + ' ' + e.status), res.revisions.map((x) => [x.tdoc, x.kind, x.state3gu, x.action])],
    [false, ['S4aI260101 available', 'S4aI260097 postponed', 'S4aI260103 parked'], [['S4aI260089', 'revision-not-reserved', 'available', 'Author needs to reserve the revision in 3GU.']]]);
  check('the postponed TDoc stands in "Documents Not Yet Disposed", with its status', summaryOf(b)[2][1].map((x) => x[0] + ' ' + x[3]), ['S4aI260101 available', 'S4aI260097 postponed', 'S4aI260103 parked']);
  // The fixture output: the three tables as they stand in the report.
  console.log('\n  --- the Disposition Summary of the MBS fixture, as written ---');
  summaryOf(b).forEach((c) => {
    console.log('  ' + c[0]);
    console.log('    TDoc       | Title                         | Source      | Status');
    (c[1] === 'None.' ? [] : c[1]).forEach((x) => console.log('    ' + x[0] + ' | ' + x[1].padEnd(29) + ' | ' + x[2].padEnd(11) + ' | ' + x[3]));
  });
  console.log('\n  --- the Close Report dialog ---\n' + lastAlert(r)[1].split('\n').map((l) => '  | ' + l).join('\n') + '\n');
  // The minutes corrected: the three decisions taken, the revision reserved and uploaded in 3GU.
  const corrected = { S4aI260101: 'noted', S4aI260097: 'noted', S4aI260103: 'agreed' };
  r.user(() => Object.keys(corrected).forEach((n2) => { const table = r.table(n2); for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === 'Status') table.getRow(i).getCell(1).setText(corrected[n2]); }));
  r.tdocs.filter((x) => x.id === 'S4aI260089')[0].status = 'revised';
  r.tdocs.filter((x) => x.id === 'S4aI260089')[0].revisedTo = 'S4aI260110';
  r.tdocs.push({ id: 'S4aI260110', agenda: '2.5', status: 'available', revisedTo: '', uploaded: true });
  r.update();
  r.user(() => { const table = r.table('S4aI260110'); for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === 'Status') table.getRow(i).getCell(1).setText('agreed'); });
  const again = closeReport(r);
  check('corrected and run again: ready to close; the final summary has 30 TDocs, none open', [again.ready, again.checked, summaryOf(b).map((c) => c[0]), lastAlert(r)[0]],
    [true, 30, [TITLES[0] + ' (13)', TITLES[1] + ' (17)', TITLES[2] + ' (0)'], 'Report ready to close']);
  check('still one summary, still the end of the report', [summaryHeadings(b), afterSummary(b), outline(b).filter((x) => x[0] === '#').slice(-3)], [1, [], ['# 2.8 Close of the session', '# ' + ATTENDANCE, '# ' + SUMMARY]]);
}

// ====================================================================== H
console.log('H. backward safety: a report without a summary is never given one; Build Report from Scratch');
{
  const r = built([t(1, '4.3'), t(2, '5')], { 1: 'agreed' });
  const before = snapshot(r.body);
  for (let i = 0; i < 4; i++) check('update ' + (i + 1) + ' of a report without a summary: its result is what it always was', r.update(), { success: true, error: null });
  check('no summary was written, nothing was logged about one, the document is as it was', [summaryHeadings(r.body), r.logs.filter((l) => /Disposition Summary/.test(l)), snapshot(r.body) === before], [0, [], true]);
  // A summary, then Build Report from Scratch: it is cleared, the result says so, and no update writes it again.
  updateSummary(r);
  const rebuilt = r.build();
  check('Build Report from Scratch clears it and says how to write it again', [rebuilt.ok, summaryHeadings(r.body), /The Disposition Summary was cleared with the report\. Use SA4 Report > Report > Update Disposition Summary to write it again\./.test(rebuilt.phases[0].detail)],
    [true, 0, true]);
  r.update();
  check('the next update does not write it again', summaryHeadings(r.body), 0);
  const plain = built([t(1, '4.3')], {}).build();
  check('a rebuild of a report that had none says nothing about it', /Disposition Summary/.test(plain.phases[0].detail), false);
  // The update path keeps a summary only where there is no background context (pinned by the ledger in I as well).
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  check('the update keeps a summary only in a run with the document open (not in a background run)', /\n    if \(!context\) \{\n      const summary = syncDispositionSummaryOnUpdate_\(body\);/.test(src), true);
}

// ====================================================================== J
console.log('J. a summary written by an earlier candidate (00a50c5, 4276161) is brought to the current form');
{
  // The 6G smoke copy, 2026-10-29: its summary was written by 00a50c5 (approved in "Other Disposed Documents", the first
  // table "Agreed and Endorsed Documents"). Update Disposition Summary of ed43e37 took that first line for text typed into
  // the summary, refused, and left the summary as it was. The fixture below is that form, literally.
  const NOTE_00A50C5 = 'Written from the Status of the TDoc tables by SA4 Report > Report > Update Disposition Summary, and kept up to date by every update. Text typed between this line and its last table stops it from being updated; text below it is kept.';
  const NOTE_4276161 = 'Written from the Status of the TDoc tables by SA4 Report > Report > Update Disposition Summary, and kept up to date by every update. Text typed into this section stops it from being updated.';
  const OLD_FIRST = 'Agreed and Endorsed Documents';
  const HEADER = ['TDoc', 'Title', 'Source', 'Status'];
  // 01 agreed, 02 approved, 03 noted, 04 available, 05 endorsed, 06 postponed (postponed: not yet disposed since 00a50c5).
  const DECISIONS = { 1: 'agreed', 2: 'approved', 3: 'noted', 4: 'available', 5: 'endorsed', 6: 'postponed' };
  const TDOCS = [t(1, '4.3'), t(2, '4.3'), t(3, '4.4'), t(4, '4.4'), t(5, '4.5'), t(6, '4.5')];
  const oldRow = (n) => [id(n), 'Synthetic title of ' + id(n), 'ExampleCorp', DECISIONS[n]];
  /**
   * Writes the summary exactly as the earlier candidate wrote it: the heading appended at the end of the document, then
   * each part inserted in its turn -- in front of the empty paragraph Google Docs keeps after a table that ends the
   * document, as renderDispositionSummary_() of that candidate did.
   */
  function writeEarlierSummary(r, note, classes) {
    r.user(() => {
      const b = r.body;
      b.appendParagraph(SUMMARY).setHeading('H1');
      let at = b.getNumChildren();
      const paragraph = (text, bold) => { const x = at >= b.getNumChildren() ? b.appendParagraph(text) : b.insertParagraph(at, text); x.setHeading('NORMAL'); x.editAsText().setBold(bold); at++; };
      paragraph(note, false);
      classes.forEach(([title, rows]) => {
        paragraph(title + ' (' + rows.length + ')', true);
        if (!rows.length) { paragraph('None.', false); return; }
        const cells = [HEADER].concat(rows);
        const table = at >= b.getNumChildren() ? b.appendTable(cells) : b.insertTable(at, cells);
        at = b.getChildIndex(table) + 1;
        rows.forEach((row, i) => table.getRow(i + 1).getCell(0).editAsText().setLinkUrl(0, row[0].length - 1, zip(row[0])));
      });
    });
  }
  const OLD_CLASSES = [[OLD_FIRST, [oldRow(1), oldRow(5)]], ['Other Disposed Documents', [oldRow(2), oldRow(3)]], ['Documents Not Yet Disposed', [oldRow(4), oldRow(6)]]];
  const fresh = () => built(TDOCS, DECISIONS);

  // The fixture is what 00a50c5 really writes (when its commit is in this checkout).
  let old = null;
  try { old = execFileSync('git', ['show', '00a50c5:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8'); } catch (e) { old = null; }
  if (old === null) {
    console.log('  note: 00a50c5 is not available in this checkout; the fixture is not compared with its real output.');
  } else {
    const real = fresh();
    const vm = require('vm');
    const ctx = vm.createContext(Object.assign({}, real.s));
    vm.runInContext(old, ctx);
    real.exec(() => ctx.writeDispositionSummary_(real.body));
    const literal = fresh();
    writeEarlierSummary(literal, NOTE_00A50C5, OLD_CLASSES);
    // Every element from the summary heading to the end of the document: type, heading, text, bold; every cell's text and link.
    const region = (x) => { const all = JSON.parse(snapshot(x.body)); const at = all.findIndex((e) => e[0] === 'P' && e[2] === SUMMARY); return JSON.stringify(all.slice(at).map((e) => (e[0] === 'T' ? ['T', e[1].map((row) => row.map((c) => [c[0], c[2]]))] : e))); };
    check('the fixture is, element for element (empty paragraphs and bold included), the summary the real 00a50c5 code writes for this report (approved in "Other Disposed")', region(literal), region(real));
  }

  const r = fresh();
  writeEarlierSummary(r, NOTE_00A50C5, OLD_CLASSES);
  const PROSE = 'Prose typed below the summary of the earlier candidate.';
  r.user(() => r.body.appendParagraph(PROSE));
  const section = r.s.findDispositionSummarySection_(r.body);
  check('the earlier summary is recognised as the summary\'s own: complete, nothing foreign', [section.complete, section.foreign, section.count], [true, [], 1]);
  const before = outsideSummary(r);
  const count = r.body._children.length;
  updateSummary(r);
  const sum = summaryOf(r.body);
  check('Update Disposition Summary: "updated", under the normal title', lastAlert(r).slice(0, 2),
    [SUMMARY, 'Disposition Summary updated.\n\n6 TDocs: 3 agreed/endorsed/approved, 1 other disposed, 2 not yet disposed.']);
  check('the new first heading is there, the old one is gone', [sum.map((c) => c[0]), r.body._children.filter((c) => c.getType() === 'PARAGRAPH' && c.getText().indexOf(OLD_FIRST) === 0).length],
    [[TITLES[0] + ' (3)', TITLES[1] + ' (1)', TITLES[2] + ' (2)'], 0]);
  check('approved moved to the first table; it is no longer in Other Disposed', [sum[0][1].map((x) => x[0] + ' ' + x[3]), sum[1][1].some((x) => x[0] === id(2))],
    [[id(1) + ' agreed', id(2) + ' approved', id(5) + ' endorsed'], false]);
  check('every other TDoc where it belongs, each once, with its link', [sum[1][1].map((x) => x[0] + ' ' + x[3]), sum[2][1].map((x) => x[0] + ' ' + x[3]), sum.reduce((a, c) => a.concat(c[1].map((x) => x[0])), []).sort(), sum[0][1].map((x) => x[4] === zip(x[0]))],
    [[id(3) + ' noted'], [id(4) + ' available', id(6) + ' postponed'], [1, 2, 3, 4, 5, 6].map(id), [true, true, true]]);
  check('one summary, rewritten in place: the prose below kept, the rest of the report as it was, no element more', [summaryHeadings(r.body), outsideSummary(r) === before, r.body._children.length - count,
    r.body._children.findIndex((c) => c.getType() === 'PARAGRAPH' && c.getText() === PROSE) > r.body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY)], [1, true, 0, true]);
  check('the note is the current one', r.body._children.filter((c) => c.getType() === 'PARAGRAPH' && /^Written from the Status/.test(c.getText())).map((c) => c.getText() === NOTE_00A50C5), [true]);
  const once = snapshot(r.body);
  updateSummary(r);
  check('a second update: "already up to date", the document identical', [lastAlert(r)[1].split('\n')[0], snapshot(r.body) === once], ['Disposition Summary was already up to date.', true]);
  check('an update keeps it so', [r.update().dispositionSummary, snapshot(r.body) === once], ['unchanged', true]);

  // The form of 4276161 (another note) is brought to the current form too.
  const r2 = fresh();
  writeEarlierSummary(r2, NOTE_4276161, OLD_CLASSES);
  updateSummary(r2);
  check('a summary of 4276161 (its note) is rewritten in the current form as well', [lastAlert(r2)[1].split('\n')[0], summaryOf(r2.body).map((c) => c[0]), summaryHeadings(r2.body)],
    ['Disposition Summary updated.', [TITLES[0] + ' (3)', TITLES[1] + ' (1)', TITLES[2] + ' (2)'], 1]);

  // Text typed INTO an earlier summary is still refused: the earlier lines are recognised, nothing broader.
  const r3 = fresh();
  writeEarlierSummary(r3, NOTE_00A50C5, OLD_CLASSES);
  r3.user(() => { const tables = r3.body._children.filter(isSummaryTable); r3.body.insertParagraph(r3.body.getChildIndex(tables[0]) + 1, 'A remark typed into the earlier summary.'); });
  const typed = snapshot(r3.body);
  updateSummary(r3);
  check('text typed into an earlier summary: refused, nothing changed, under a title that says so', [snapshot(r3.body) === typed, lastAlert(r3)[0], /text was typed into it \("A remark typed into the earlier summary\."\)/.test(lastAlert(r3)[1])],
    [true, 'Disposition Summary NOT updated', true]);
  check('a line only like an earlier one ("Agreed and Endorsed Documents" without its count) is not the summary\'s', ['Agreed and Endorsed Documents', 'Agreed and Endorsed Documents (x)', 'Agreed and Endorsed Documents (3) and more', OLD_FIRST + ' (3)'].map((x) => r3.s.isDispositionSummaryParagraphText_(x)),
    [false, false, false, true]);
  // A refusal of the current form, too, never shows under the title of a success.
  const r4 = fresh();
  updateSummary(r4);
  r4.user(() => { const tables = r4.body._children.filter(isSummaryTable); r4.body.insertParagraph(r4.body.getChildIndex(tables[0]) + 1, 'Typed.'); });
  updateSummary(r4);
  check('a refusal is titled "Disposition Summary NOT updated"', lastAlert(r4)[0], 'Disposition Summary NOT updated');
}

// ====================================================================== I
console.log('I. what changed in Code.js and template/ReportCreator.js');
{
  const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const CREATOR = fs.readFileSync(path.join(ROOT, 'template', 'ReportCreator.js'), 'utf8').replace(/\r/g, '');
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const view = codeWithoutDispositionSummary(CODE);
  const creatorView = creatorWithoutDispositionSummary(CREATOR);
  check('every change is in Code.js and ReportCreator.js exactly once', view.problems.concat(creatorView.problems), []);
  const section = CODE.slice(CODE.indexOf(DISPOSITION_SECTION_START), CODE.indexOf(DISPOSITION_SECTION_END) + DISPOSITION_SECTION_END.length);
  check('the new functions are one section, in this order', functionNames(section), DISPOSITION_SECTION_FUNCTIONS);
  check('the section stands directly after the configuration guard', CODE.indexOf('// END OF CONFIGURATION GUARD\n// =========================================================\n\n' + DISPOSITION_SECTION_START) !== -1, true);
  check('it declares its constants and nothing else at top level', section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || [],
    ['const DISPOSITION_SUMMARY_HEADING_', 'const DISPOSITION_SUMMARY_NOTE_', 'const DISPOSITION_SUMMARY_NONE_', 'const DISPOSITION_SUMMARY_TABLE_HEADER_', 'const DISPOSITION_SUMMARY_TABLE_WIDTHS_', 'const DISPOSITION_SUMMARY_MENU_HINT_',
      'const DISPOSITION_CLASSES_', 'const DISPOSITION_OF_STATUS_', 'const DISPOSITION_SUMMARY_EARLIER_NOTES_', 'const DISPOSITION_SUMMARY_EARLIER_TITLES_', 'const CLOSE_REPORT_LIST_LIMIT_',
      'const REVISION_CHECK_STATES_']);
  check('Code.js is still version 2.25.0 (not released)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.25.0');
  let released = null;
  let releasedCreator = null;
  try {
    const show = (file) => execFileSync('git', ['show', 'template-release/T-2026.10.15:' + file], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
    released = show('Code.js');
    releasedCreator = show('template/ReportCreator.js');
  } catch (e) {
    released = null;
  }
  if (released === null) {
    console.log('  note: template-release/T-2026.10.15 is not available in this checkout; the comparison with it is skipped.');
  } else {
    check('without it, Code.js is the Code.js of T-2026.10.15, byte for byte', view.code === released, true);
    check('without its two menu items, ReportCreator.js is the one of T-2026.10.15, byte for byte', creatorView.code === releasedCreator, true);
    check('the released functions of Code.js that changed are these six', functionNames(released).filter((name, i, list) => list.indexOf(name) === i).filter((name) => functionSource(CODE, name) !== functionSource(released, name)),
      DISPOSITION_CHANGED_FUNCTIONS.slice().sort((a, b) => CODE.indexOf('\nfunction ' + a + '(') - CODE.indexOf('\nfunction ' + b + '(')));
    check('the new functions are those of the section; none was removed', [functionNames(CODE).filter((name) => functionNames(released).indexOf(name) === -1).sort(), functionNames(released).filter((name) => functionNames(CODE).indexOf(name) === -1)],
      [DISPOSITION_SECTION_FUNCTIONS.slice().sort(), []]);
    check('the one function of ReportCreator.js that changed is the report menu', functionNames(releasedCreator).filter((name) => functionSource(CREATOR, name) !== functionSource(releasedCreator, name)), DISPOSITION_CREATOR_CHANGED_FUNCTIONS);
    check('HyperLink.js and the manifest are the released files', ['HyperLink.js', 'appsscript.json'].filter((file) => fs.readFileSync(path.join(ROOT, file), 'utf8').replace(/\r/g, '') !==
      execFileSync('git', ['show', 'template-release/T-2026.10.15:' + file], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '')), []);
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll disposition-summary and close-report checks passed.');
process.exitCode = failures ? 1 : 0;
