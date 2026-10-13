/**
 * The structure of a report (after T-2026.10.12) -- docs/REPORT_STRUCTURE.md.
 *
 * A report is: the agenda sections in the order of their numbers, then what
 * closes it ("Close of the session", and the Attendance section of an ad-hoc
 * report). Two defects of the update broke this in a live report:
 *
 *   - the heading of an agenda item the report did not have was appended at
 *     the END OF THE DOCUMENT -- after the close and after the attendance --
 *     and the TDocs of that item with it;
 *   - the revision placement found every revision "out of place" on every
 *     update (Google Docs keeps a paragraph between two tables, so a table is
 *     never the very next child of another), moved it again and left one
 *     empty paragraph behind each time: 1,129 of them in three runs.
 *
 * These tests run the REAL build and the REAL update on the fake document.
 * The document is made with { tableSeparators: true }: like Google Docs it
 * puts a paragraph between two tables and does not let its last paragraph be
 * removed. (The earlier tests ran on a document where two tables could touch,
 * which is why none of them saw the second defect.)
 *
 *   A.  a missing agenda item in the middle gets its heading in its place
 *   B.  the same for the item directly before the close
 *   C.  a reallocated TDoc is placed under its new agenda item
 *   D.  the order in which TDocs are discovered does not matter
 *   E.  nothing of the agenda after the Attendance section
 *   F.  nothing of the agenda after "Close of the session"
 *   G.  an update is idempotent
 *   H.  no empty paragraphs are left behind, however often it runs
 *   I.  minutes and dispositions written by hand survive every move
 *   J.  a report that is in order is not rewritten
 *   K.  the same rules in a main-meeting report
 *   L.  the ad-hoc report with sessions: tests/adhoc-integration.test.js
 *   M.  the structure of the live report that showed all this, repaired by
 *       one update (tests/fixtures/mbs-report-structure-damaged.json)
 *   N.  what the fix changed in Code.js, and nothing else
 *
 * Run: node tests/report-structure-fix.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { templateReport } = require('./helpers/template-report.js');
const { codeBeforeStructureFix, SECTION_START, SECTION_END, SECTION_FUNCTIONS, CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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
const ATTENDANCE = 'Attendance';
const item = ([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' });
/** An ad-hoc agenda like the one of the live report: 2.1 opening, 2.2 IPR, then the given items, then 2.8 Close. */
const agendaWith = (items) => [['2.1', 'Opening of the session and registration of documents'], ['2.2', 'IPR, antitrust & consensus principles reminder'], ['2.3', 'Reports/Liaisons from other groups/meetings']]
  .concat(items).concat([['2.8', 'Close of the session']]).map(item);
const t = (id, agenda, revisedTo, status) => ({ id: id, agenda: agenda, status: status || (revisedTo ? 'revised' : 'available'), revisedTo: revisedTo || '' });
const id = (n) => 'S4aA2690' + String(n).padStart(2, '0');
const report = (agendaItems, tdocs) => templateReport({ agenda: agendaWith(agendaItems), tdocs: tdocs, tableSeparators: true });

const isTdocTable = (c) => c.getType() === 'TABLE' && c.getNumRows() > 0 && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc';
const isHeading = (c) => c.getType() === 'PARAGRAPH' && c.getHeading() !== 'NORMAL';
const isEmpty = (c) => c.getType() === 'PARAGRAPH' && c.getHeading() === 'NORMAL' && c.getText() === '';
/** The document as headings and TDoc numbers, in order: ['# 2.5 ...', 'S4aA269001', ...]. */
const outline = (body) => body._children.map((c) => (isTdocTable(c) ? c.getRow(0).getCell(1).getText() : isHeading(c) ? '# ' + c.getText() : null)).filter(Boolean);
/** The agenda part and the end: the numbers of the numbered level-two headings and 'Attendance', in order. */
const sections = (body) => body._children.filter(isHeading).map((c) => c.getText()).filter((x) => /^\d+\.\d+( |$)/.test(x) || x === ATTENDANCE).map((x) => (x === ATTENDANCE ? x : x.split(' ')[0]));
/** The TDoc numbers of the tables below the heading of `agenda`, up to the next heading. */
const under = (body, agenda) => {
  const lines = outline(body);
  const at = lines.findIndex((line) => line === '# ' + agenda || line.indexOf('# ' + agenda + ' ') === 0);
  const out = [];
  for (let i = at + 1; at !== -1 && i < lines.length && lines[i][0] !== '#'; i++) out.push(lines[i]);
  return out;
};
const empties = (body) => body._children.filter(isEmpty).length;
/** The longest run of empty paragraphs. */
const longestRun = (body) => body._children.reduce((acc, c) => { const run = isEmpty(c) ? acc.run + 1 : 0; return { run: run, max: Math.max(acc.max, run) }; }, { run: 0, max: 0 }).max;
/** What stands after the element whose heading text is `text` that belongs to the agenda: TDoc tables and numbered headings. */
const agendaMaterialAfter = (body, text) => {
  const at = body._children.findIndex((c) => isHeading(c) && (c.getText() === text || c.getText().indexOf(text) !== -1));
  if (at === -1) return ['(no such heading)'];
  return body._children.slice(at + 1).map((c) => (isTdocTable(c) ? c.getRow(0).getCell(1).getText() : isHeading(c) && /^\d+(\.\d+)*( |$)/.test(c.getText()) ? '# ' + c.getText() : null)).filter(Boolean);
};
/** The Attendance section as the import writes it, at the end of the closing section -- here: the end of the document. */
function addAttendance(r) {
  r.user(() => {
    const b = r.body;
    b.appendParagraph(ATTENDANCE).setHeading('H3');
    b.appendParagraph('Statistics');
    b.appendParagraph('Meeting: Synthetic meeting');
    b.appendParagraph('Attendees');
    b.appendTable([['Name', 'Company', 'Email'], ['Sam Rivera', 'ExampleCorp', 'sam.rivera@example.com'], ['Kim Doe', 'ExampleCorp', 'kim.doe@example.com']]);
    b.appendParagraph('');
  });
}
const structureLogs = (r) => r.logs.filter((l) => /^Structure: /.test(l));
const field = (body, tdoc, label) => {
  const table = body._children.filter((c) => isTdocTable(c) && c.getRow(0).getCell(1).getText() === tdoc)[0];
  if (!table) return undefined;
  for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === label) return table.getRow(i).getCell(1).getText();
  return undefined;
};
const setField = (body, tdoc, label, text) => {
  const table = body._children.filter((c) => isTdocTable(c) && c.getRow(0).getCell(1).getText() === tdoc)[0];
  for (let i = 0; i < table.getNumRows(); i++) if (table.getRow(i).getCell(0).getText() === label) table.getRow(i).getCell(1).setText(text);
};

// ====================================================================== A
console.log('A. a missing agenda item in the middle: 2.5, 2.7, 2.8, Attendance -- and a TDoc for 2.6');
{
  const r = report([['2.5', 'Item five'], ['2.7', 'Item seven']], [t(id(1), '2.5'), t(id(2), '2.7')]);
  check('the report is built', r.build().ok, true);
  addAttendance(r);
  check('before: 2.5, 2.7, 2.8, Attendance', sections(r.body).slice(-4), ['2.5', '2.7', '2.8', ATTENDANCE]);
  r.tdocs.push(t(id(3), '2.6'));
  check('the update completes', r.update(), { success: true, error: null });
  check('after: 2.5, 2.6, 2.7, 2.8, Attendance', sections(r.body).slice(-5), ['2.5', '2.6', '2.7', '2.8', ATTENDANCE]);
  check('the TDoc is under 2.6, and 2.5 and 2.7 hold what they held', [under(r.body, '2.5'), under(r.body, '2.6'), under(r.body, '2.7')], [[id(1)], [id(3)], [id(2)]]);
  check('the new heading has the title the TDoc list gives the item, and the level of its number', r.body._children.filter((c) => isHeading(c) && /^2\.6/.test(c.getText())).map((c) => [c.getText(), c.getHeading()]), [['2.6 Synthetic item', 'H2']]);
  check('E, F: nothing of the agenda after the close or after the attendance', [agendaMaterialAfter(r.body, 'Close of the session'), agendaMaterialAfter(r.body, ATTENDANCE)], [[], []]);
}

// ====================================================================== B
console.log('B. the item directly before the close: 2.6, 2.8, Attendance -- and a TDoc for 2.7');
{
  const r = report([['2.6', 'Item six']], [t(id(1), '2.6')]);
  r.build();
  addAttendance(r);
  r.tdocs.push(t(id(2), '2.7'));
  r.update();
  check('after: 2.6, 2.7, 2.8, Attendance', sections(r.body).slice(-4), ['2.6', '2.7', '2.8', ATTENDANCE]);
  check('the TDoc is under 2.7', [under(r.body, '2.6'), under(r.body, '2.7')], [[id(1)], [id(2)]]);
  check('E, F', [agendaMaterialAfter(r.body, 'Close of the session'), agendaMaterialAfter(r.body, ATTENDANCE)], [[], []]);

  // An item numbered after the closing item has no later item to stand before: it goes to the end of the agenda part,
  // directly before the close -- never after it.
  r.tdocs.push(t(id(3), '2.9'));
  r.update();
  check('an item numbered after the closing item still stands before it', sections(r.body).slice(-5), ['2.6', '2.7', '2.9', '2.8', ATTENDANCE]);
  check('E, F', [agendaMaterialAfter(r.body, 'Close of the session'), agendaMaterialAfter(r.body, ATTENDANCE)], [[], []]);
}

// ====================================================================== C
console.log('C. a TDoc reallocated from 3.7 to 2.7 is placed under 2.7');
{
  const r = report([['2.5', 'Item five']], [t(id(1), '2.5')]);
  r.build();
  addAttendance(r);
  r.user(() => r.s.saveReallocation(id(2), '3.7', '2.7', 'Wrong SWG'));
  r.tdocs.push(t(id(2), '3.7'));
  r.tdocs.push(t(id(3), '2.7'));
  check('the update completes', r.update(), { success: true, error: null });
  check('after: 2.5, 2.7, 2.8, Attendance -- and no 3.7', sections(r.body).slice(-4), ['2.5', '2.7', '2.8', ATTENDANCE]);
  check('the reallocated TDoc is under 2.7, with the TDoc that is registered there', under(r.body, '2.7'), [id(2), id(3)]);
  check('its table names the new agenda item', field(r.body, id(2), 'Agenda Item'), '2.7');
  // The TDoc list describes item 3.7 in the row of the reallocated TDoc; the title of 2.7 comes from the TDoc registered under 2.7.
  check('the heading has the title of 2.7', r.body._children.filter((c) => isHeading(c) && /^2\.7/.test(c.getText())).map((c) => c.getText()), ['2.7 Synthetic item']);
  check('E, F', [agendaMaterialAfter(r.body, 'Close of the session'), agendaMaterialAfter(r.body, ATTENDANCE)], [[], []]);

  // Alone, the reallocated TDoc gives its new item no title: a bare number, in its place.
  const alone = report([['2.5', 'Item five']], [t(id(1), '2.5')]);
  alone.build();
  addAttendance(alone);
  alone.user(() => alone.s.saveReallocation(id(2), '3.7', '2.7', 'Wrong SWG'));
  alone.tdocs.push(t(id(2), '3.7'));
  alone.update();
  check('a reallocated TDoc alone: the heading is the number, where it belongs', [sections(alone.body).slice(-4), alone.body._children.filter((c) => isHeading(c) && /^2\.7/.test(c.getText())).map((c) => c.getText()), under(alone.body, '2.7')],
    [['2.5', '2.7', '2.8', ATTENDANCE], ['2.7'], [id(2)]]);
}

// ====================================================================== D
console.log('D. TDocs discovered in the order 2.7, 2.4, 2.6, 2.5');
{
  const r = report([], []);
  r.tdocs.push(t(id(9), '2.3'));
  r.build();
  addAttendance(r);
  r.tdocs = r.tdocs.concat([t(id(1), '2.7'), t(id(2), '2.4'), t(id(3), '2.6'), t(id(4), '2.5'), t(id(5), '2.4')]);
  check('one update', r.update(), { success: true, error: null });
  check('the sections are 2.3 ... 2.8 in the order of the agenda, then Attendance', sections(r.body).slice(-7), ['2.3', '2.4', '2.5', '2.6', '2.7', '2.8', ATTENDANCE]);
  check('each TDoc under its item', ['2.4', '2.5', '2.6', '2.7'].map((a) => under(r.body, a)), [[id(2), id(5)], [id(4)], [id(3)], [id(1)]]);
  check('each heading once', sections(r.body).filter((x, i, list) => list.indexOf(x) !== i), []);
  check('E, F', [agendaMaterialAfter(r.body, 'Close of the session'), agendaMaterialAfter(r.body, ATTENDANCE)], [[], []]);

  // The same TDocs, one update each, in that order.
  const slow = report([], [t(id(9), '2.3')]);
  slow.build();
  addAttendance(slow);
  [t(id(1), '2.7'), t(id(2), '2.4'), t(id(3), '2.6'), t(id(4), '2.5'), t(id(5), '2.4')].forEach((one) => { slow.tdocs.push(one); slow.update(); });
  check('one TDoc per update gives the same report', outline(slow.body), outline(r.body));
}

// ====================================================================== G, H
console.log('G, H. idempotence, and no empty paragraphs left behind');
{
  // Two revision chains and a revision under another agenda item than its parent.
  const [A, B, C, X, Y, P, Q, Z] = [1, 2, 3, 4, 5, 6, 7, 8].map(id);
  const r = report([['2.4', 'Item four'], ['2.5', 'Item five']],
    [t(A, '2.5', B), t(Z, '2.5'), t(X, '2.5', Y), t(B, '2.5', C), t(Y, '2.5'), t(C, '2.5'), t(P, '2.4', Q), t(Q, '2.5')]);
  r.build();
  addAttendance(r);
  r.update();
  const afterFirst = { outline: outline(r.body), children: r.body._children.length, empties: empties(r.body), tables: r.body._children.filter((c) => c.getType() === 'TABLE').length };
  check('every revision is directly below the document it revises', [under(r.body, '2.4'), under(r.body, '2.5')], [[P, Q], [A, B, C, Z, X, Y]]);
  r.mark();
  const second = r.update();
  check('G: a second update leaves every element where it is, and moves no table',
    [second.success, outline(r.body), r.body._children.length, empties(r.body), r.logs.filter((l) => /^Revisions: /.test(l)), structureLogs(r)],
    [true, afterFirst.outline, afterFirst.children, afterFirst.empties, ['Revisions: 0 moved, 0 disposition(s) filled'], []]);
  for (let i = 0; i < 40; i++) r.update();
  check('H: forty more updates: not one element more, not one empty paragraph more', [r.body._children.length, empties(r.body), outline(r.body)], [afterFirst.children, afterFirst.empties, afterFirst.outline]);
  // The bound: an empty paragraph is there only where the document needs one -- between two tables, and at its end.
  const neededSeparators = r.body._children.filter((c, i, list) => c.getType() === 'TABLE' && list.slice(i + 1).findIndex((n) => !isEmpty(n)) !== -1 && list.slice(i + 1).filter((n) => !isEmpty(n))[0].getType() === 'TABLE').length;
  check('H: there is exactly one empty paragraph between two tables, and never two in a row', [longestRun(r.body), empties(r.body) <= neededSeparators + 2], [1, true]);
  check('G: one Attendance section, each heading and each TDoc once', [sections(r.body).filter((x) => x === ATTENDANCE).length, outline(r.body).filter((x, i, list) => list.indexOf(x) !== i)], [1, []]);

  // A revision that arrives later is moved once, and leaves nothing where it stood.
  const before = empties(r.body);
  r.tdocs.filter((one) => one.id === Z)[0].revisedTo = id(20);
  r.tdocs.filter((one) => one.id === Z)[0].status = 'revised';
  r.tdocs.push(t(id(21), '2.5'));
  r.tdocs.push(t(id(20), '2.4'));
  r.update();
  const settled = { outline: outline(r.body), empties: empties(r.body) };
  check('a new revision registered under another item is below its parent', under(r.body, '2.5'), [A, B, C, Z, id(20), X, Y, id(21)]);
  r.update();
  r.update();
  check('H: and stays there, with no paragraph left behind at either place', [outline(r.body), empties(r.body), longestRun(r.body), empties(r.body) - before], [settled.outline, settled.empties, 1, 2]);
}

// ====================================================================== I
console.log('I. minutes and dispositions written by hand survive');
{
  const [A, B, X, W] = [1, 2, 3, 4].map(id);
  const r = report([['2.5', 'Item five']], [t(A, '2.5'), t(X, '2.5'), t(B, '2.5')]);
  r.build();
  addAttendance(r);
  const MINUTES = 'Presented by Sam Rivera.\nQ: is this aligned with the CR?\nA: yes, in clause 5.';
  const DISPOSITION = 'Agreed to align the terminology; a revision is expected.';
  r.user(() => {
    [A, B, X].forEach((one) => { setField(r.body, one, 'Minutes', MINUTES + ' [' + one + ']'); setField(r.body, one, 'Disposition', DISPOSITION + ' [' + one + ']'); });
    // A section the old update appended after the attendance, with a table that has minutes too, and a note typed below it.
    const b = r.body;
    b.appendParagraph('2.4').setHeading('H2');
    b.appendTable([['TDoc', W], ['Title', 'Synthetic title of ' + W], ['Agenda Item', '2.4'], ['Minutes', MINUTES + ' [' + W + ']'], ['Disposition', DISPOSITION + ' [' + W + ']'], ['Status', 'available']]);
    b.appendParagraph('A note typed by hand below the table.');
    b.appendParagraph('');
  });
  // B becomes the revision of A: it is moved below A. W's section is moved before 2.5.
  r.tdocs.filter((one) => one.id === A)[0].revisedTo = B;
  r.tdocs.filter((one) => one.id === A)[0].status = 'revised';
  r.tdocs.push(t(W, '2.4'));
  check('the update completes', r.update(), { success: true, error: null });
  check('the tables are where they belong', [under(r.body, '2.4'), under(r.body, '2.5'), sections(r.body).slice(-4)], [[W], [A, B, X], ['2.4', '2.5', '2.8', ATTENDANCE]]);
  check('the minutes of every table are exactly what was typed', [A, B, X, W].map((one) => field(r.body, one, 'Minutes')), [A, B, X, W].map((one) => MINUTES + ' [' + one + ']'));
  check('the dispositions typed by hand are kept; the revised document says what it was revised to as well',
    [A, B, X, W].map((one) => field(r.body, one, 'Disposition')),
    [DISPOSITION + ' [' + A + '] — Revised to ' + B.toUpperCase(), DISPOSITION + ' [' + B + ']', DISPOSITION + ' [' + X + ']', DISPOSITION + ' [' + W + ']']);
  const texts = r.body._children.map((c) => (c.getType() === 'PARAGRAPH' ? c.getText() : null));
  check('the note typed below the moved table went with its section, once', [texts.filter((x) => x === 'A note typed by hand below the table.').length,
    texts.indexOf('A note typed by hand below the table.') < r.body._children.findIndex((c) => isHeading(c) && /^2\.5 /.test(c.getText()))], [1, true]);
}

// ====================================================================== J
console.log('J. a report that is in order is only read');
{
  const r = report([['2.4', 'Item four'], ['2.5', 'Item five'], ['2.6', 'Item six'], ['2.7', 'Item seven']], [t(id(1), '2.4'), t(id(2), '2.5', id(3)), t(id(3), '2.5'), t(id(4), '2.6'), t(id(5), '2.7')]);
  r.build();
  addAttendance(r);
  const elements = r.body._children.slice();
  const before = outline(r.body);
  r.mark();
  check('the update completes', r.update(), { success: true, error: null });
  check('every element of the body is the very same object, in the same place', [r.body._children.length === elements.length, r.body._children.every((c, i) => c === elements[i]), outline(r.body)], [true, true, before]);
  check('nothing was repaired and nothing was moved', [structureLogs(r), r.logs.filter((l) => /^Revisions: /.test(l))], [[], ['Revisions: 0 moved, 0 disposition(s) filled']]);
  // Two empty lines typed by hand between two paragraphs, and three before a heading, are nobody's debris.
  r.user(() => { const at = r.body._children.findIndex((c) => isHeading(c) && /^2\.8 /.test(c.getText())) + 1; r.body.insertParagraph(at, 'Typed by hand.'); r.body.insertParagraph(at + 1, ''); r.body.insertParagraph(at + 2, ''); r.body.insertParagraph(at + 3, 'Typed by hand, further down.'); });
  const withBlanks = r.body._children.length;
  r.update();
  check('empty paragraphs between paragraphs typed by hand are left alone', r.body._children.length, withBlanks);
}

// ====================================================================== K
console.log('K. the same rules in a main-meeting report');
{
  const { sandbox: s } = loadCode({ documentProperties: {} });
  s.styleStatusCell_ = () => {};
  const body = makeFakeDocumentBody(s, { tableSeparators: true });
  const tdocRow = (tdoc, agenda) => ({ row: [tdoc, 'Title of ' + tdoc, 'Source Co', agenda], richTextRow: null, tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 3, agendaTopicCol: -1, statusCol: -1, typeCol: -1, forCol: -1, revisedToCol: -1, agendaItem: agenda });
  [['9 Video SWG', 'H1'], ['9.1 Opening of the session', 'H2'], ['9.1.1 Opening', 'H3'], ['9.2 IPR and antitrust reminder', 'H2'], ['9.5 Item five', 'H2'], ['9.5.1 Sub-item', 'H3'], ['9.9 Item nine', 'H2'],
    ['9.11 Any other business', 'H2'], ['9.12 Close of the session', 'H2']].forEach(([text, heading]) => body.appendParagraph(text).setHeading(heading));
  body.appendParagraph('The session was closed.');
  body.appendParagraph('');
  const numbered = () => body._children.filter(isHeading).map((c) => c.getText().split(' ')[0]);
  ['9.6', '9.10', '9.5.2', '9.3', '9.13', '9.6'].forEach((agenda, i) => s.insertNewTdoc_(body, tdocRow('S4-2690' + String(i + 1).padStart(2, '0'), agenda), null, null));
  check('9.3, 9.5.2, 9.6 and 9.10 stand in the order of their numbers (9.10 after 9.9, not after 9.1); 9.13 before the close',
    numbered(), ['9', '9.1', '9.1.1', '9.2', '9.3', '9.5', '9.5.1', '9.5.2', '9.6', '9.9', '9.10', '9.11', '9.13', '9.12']);
  check('each TDoc under its item; the second TDoc of 9.6 found the heading the first one made', ['9.3', '9.5.2', '9.6', '9.10', '9.13'].map((a) => under(body, a)),
    [['S4-269004'], ['S4-269003'], ['S4-269001', 'S4-269006'], ['S4-269002'], ['S4-269005']]);
  check('the level of a new heading is that of its number', body._children.filter(isHeading).filter((c) => ['9.6', '9.5.2'].indexOf(c.getText()) !== -1).map((c) => c.getHeading()), ['H3', 'H2']);
  check('F: nothing of the agenda after the close', agendaMaterialAfter(body, 'Close of the session'), []);
  check('compareAgendaNumbers_: 2.9 < 2.10, 2.5 < 2.5.1 < 2.6, 2 < 2.1, 10 > 9', [s.compareAgendaNumbers_('2.9', '2.10') < 0, s.compareAgendaNumbers_('2.5', '2.5.1') < 0, s.compareAgendaNumbers_('2.5.1', '2.6') < 0,
    s.compareAgendaNumbers_('2', '2.1') < 0, s.compareAgendaNumbers_('10', '9') > 0, s.compareAgendaNumbers_('2.7', '2.7')], [true, true, true, true, true, 0]);
  check('agendaNumberOfHeadingText_: a number at the start of a heading, and nothing else', ['2.7 Energy', '2.7', '11.0.1 Opening', 'Attendance', '2Multicast', '2026 was a year', ''].map(s.agendaNumberOfHeadingText_),
    ['2.7', '2.7', '11.0.1', '', '', '2026', '']);

  // A report without a closing heading and without attendance: the end of the document, as before.
  const plain = makeFakeDocumentBody(s, { tableSeparators: true });
  plain.appendParagraph('9.5 Item five').setHeading('H2');
  plain.appendParagraph('');
  s.insertNewTdoc_(plain, tdocRow('S4-269010', '9.7'), null, null);
  s.insertNewTdoc_(plain, tdocRow('S4-269011', '9.6'), null, null);
  check('without a close: a later item goes to the end, an earlier one before it', outline(plain), ['# 9.5 Item five', '# 9.6', 'S4-269011', '# 9.7', 'S4-269010']);
}

// ====================================================================== L
console.log('L. the ad-hoc report with sessions, opening details and attendance');
console.log('  note: tests/adhoc-integration.test.js runs the real build, the real imports and the real updates on it; its section 3 has the checks of this rule (a TDoc of the closing item before the Attendance section, an unknown item before the close).');

// ====================================================================== M
console.log('M. the structure of the live report, and one update');
{
  const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'mbs-report-structure-damaged.json'), 'utf8')).elements;
  const tdocsOfFixture = fixture.filter((e) => e.tdoc);
  const registration = fixture.filter((e) => e.registration)[0].registration;
  // The TDoc list of the meeting, as far as the report shows it: number, agenda item, "Revised to". The reallocated TDoc
  // is listed under the item it came from. S4aI260080 is assumed to be revised to S4aI260091 (same title, status "revised";
  // its disposition was written by hand and is not in the fixture).
  const revisedTo = {};
  tdocsOfFixture.forEach((e) => { if (e.revisedTo) revisedTo[e.tdoc] = e.revisedTo; });
  revisedTo.S4aI260080 = 'S4aI260091';
  const list = registration.map(([tdoc, agenda]) => t(tdoc, tdoc === 'S4aI260081' ? '3.7' : agenda, revisedTo[tdoc], (tdocsOfFixture.filter((e) => e.tdoc === tdoc)[0] || {}).status));
  const r = templateReport({ tdocs: list, tableSeparators: true, props: { REPORT_SUFFIX: 'MBS' } });
  // The document, element for element.
  r.user(() => {
    const b = r.body;
    fixture.forEach((e) => {
      if (e.h) b.appendParagraph(e.text).setHeading(e.h);
      else if (e.p) b.appendParagraph(e.p);
      else if (e.empty) for (let i = 0; i < e.empty; i++) b.appendParagraph('');
      else if (e.tdoc) {
        b.appendTable(e.rows.map((label) => [label, label === 'TDoc' ? e.tdoc : label === 'Agenda Item' ? e.agenda : label === 'Status' ? e.status : label === 'Disposition' && e.revisedTo ? 'Revised to ' + e.revisedTo
          : label === 'Minutes' ? 'Minutes of ' + e.tdoc + ', typed by hand.' : '']));
      } else if (e.registration) b.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item']].concat(e.registration.map(([tdoc, agenda]) => [tdoc, 'Synthetic title of ' + tdoc, 'ExampleCorp', agenda])));
      else if (e.reallocations) b.appendTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason']].concat(e.reallocations));
      else if (e.attendees) b.appendTable([['Name', 'Company', 'Email']].concat(Array.from({ length: e.attendees }, (x, i) => ['Person ' + (i + 1), 'ExampleCorp', 'person' + (i + 1) + '@example.com'])));
    });
  });
  const b = r.body;
  const allTdocs = outline(b).filter((x) => x[0] !== '#').slice().sort();
  check('the fixture is the document that was exported: 1,231 body elements less the section break, 1,157 empty paragraphs in runs of 571, 215 and 343 and 28 single ones, 32 tables',
    [b._children.length, empties(b), longestRun(b), b._children.filter((c) => c.getType() === 'TABLE').length, allTdocs.length], [1230, 1157, 571, 32, 29]);
  check('before: 2.6 and 2.7 have no heading in the agenda part; they stand after the close and the attendance, 2.7 before 2.6', sections(b), ['2.1', '2.2', '2.3', '2.4', '2.5', '2.8', ATTENDANCE, '2.7', '2.6']);
  check('before: ten TDoc tables and two headings after the Attendance section... as agenda material after the close', agendaMaterialAfter(b, 'Close of the session'),
    ['# 2.7', 'S4aI260081', 'S4aI260090', 'S4aI260103', 'S4aI260104', '# 2.6', 'S4aI260095', 'S4aI260097']);
  check('before: the reallocation of S4aI260081 is in the report', Object.keys(r.s.getReallocationMap_()), ['S4aI260081']);

  r.mark();
  check('one update completes', r.update(), { success: true, error: null });
  check('after: the sections are 2.1 ... 2.8 in the order of the agenda, and the Attendance section is the end', sections(b), ['2.1', '2.2', '2.3', '2.4', '2.5', '2.6', '2.7', '2.8', ATTENDANCE]);
  check('E, F: nothing of the agenda after the close or after the attendance', [agendaMaterialAfter(b, 'Close of the session'), agendaMaterialAfter(b, ATTENDANCE)], [[], []]);
  check('S4aI260095 and S4aI260097 are under 2.6', under(b, '2.6'), ['S4aI260095', 'S4aI260097']);
  check('S4aI260081 (reallocated from 3.7), its revision and the two others are under 2.7', under(b, '2.7'), ['S4aI260081', 'S4aI260090', 'S4aI260103', 'S4aI260104']);
  check('every TDoc table is still there, once', outline(b).filter((x) => x[0] !== '#').slice().sort(), allTdocs);
  check('2.4 and 2.5 hold what they held, in the same order (S4aI260089, an item-2.7 TDoc that stood under 2.5, is not moved by an update)',
    [under(b, '2.4'), under(b, '2.5')],
    [['S4aI260094', 'S4aI260096', 'S4aI260098', 'S4aI260099', 'S4aI260108'],
      ['S4aI260080', 'S4aI260091', 'S4aI260100', 'S4aI260089', 'S4aI260082', 'S4aI260102', 'S4aI260105', 'S4aI260083', 'S4aI260092', 'S4aI260084', 'S4aI260085', 'S4aI260086', 'S4aI260087', 'S4aI260088', 'S4aI260093', 'S4aI260101', 'S4aI260106', 'S4aI260107']]);
  check('the 1,129 empty paragraphs of the three runs are gone: of 1,230 body elements 104 are left, of 1,157 empty paragraphs 31',
    [b._children.length, empties(b), longestRun(b), structureLogs(r)],
    [104, 31, 2, ['Structure: agenda section 2.7 was after the close of the report; put back in the order of the agenda',
      'Structure: agenda section 2.6 was after the close of the report; put back in the order of the agenda',
      'Structure: 1126 empty paragraph(s) left below TDoc tables by earlier updates removed']]);
  check('between two TDoc tables there is exactly one empty paragraph, and none between a TDoc table and the heading after it', (() => {
    const kinds = b._children.map((c) => (isTdocTable(c) ? 'T' : isEmpty(c) ? 'e' : isHeading(c) ? 'H' : c.getType() === 'TABLE' ? 't' : 'p')).join('');
    return [(kinds.match(/Te+T/g) || []).filter((x) => x !== 'TeT').length, (kinds.match(/TT/g) || []).length, (kinds.match(/Te+H/g) || []).length, (kinds.match(/T(?=eT)/g) || []).length, (kinds.match(/TH/g) || []).length];
  })(), [0, 0, 0, 25, 4]);
  check('the two empty paragraphs in a row that are left are the ones in the opening section, which no TDoc table precedes: they are not touched', (() => {
    const at = b._children.findIndex((c, i, listed) => isEmpty(c) && isEmpty(listed[i + 1]));
    return [b._children[at - 1].getType(), b._children.slice(0, at).filter(isTdocTable).length];
  })(), ['PARAGRAPH', 0]);
  check('the minutes typed by hand are in every table, also in the ten that were moved', allTdocs.filter((tdoc) => field(b, tdoc, 'Minutes') !== 'Minutes of ' + tdoc + ', typed by hand.'), []);
  check('the Attendance section is one, with its attendee table of 24', [sections(b).filter((x) => x === ATTENDANCE).length, b._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getCell(0).getText() === 'Name').map((c) => c.getNumRows() - 1)], [1, [24]]);
  check('the headings an update had written as bare numbers have the titles the TDoc list gives their items', b._children.filter((c) => isHeading(c) && /^2\.[67]/.test(c.getText())).map((c) => [c.getText(), c.getHeading()]),
    [['2.6 Synthetic item', 'H2'], ['2.7 Synthetic item', 'H2']]);

  const steady = { outline: outline(b), children: b._children.length, empties: empties(b) };
  r.mark();
  r.update();
  r.update();
  check('G: two more updates change nothing, repair nothing and move nothing', [outline(b), b._children.length, empties(b), structureLogs(r), r.logs.filter((l) => /^Revisions: /.test(l))],
    [steady.outline, steady.children, steady.empties, [], ['Revisions: 0 moved, 0 disposition(s) filled', 'Revisions: 0 moved, 0 disposition(s) filled']]);
}

// ====================================================================== N
console.log('N. what the fix changed in Code.js');
{
  const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const view = codeBeforeStructureFix(CODE);
  check('every change the fix made is in Code.js exactly once', view.problems, []);
  const section = CODE.slice(CODE.indexOf(SECTION_START), CODE.indexOf(SECTION_END) + SECTION_END.length);
  check('its new functions are one section, in this order', functionNames(section), SECTION_FUNCTIONS);
  check('the section declares nothing else at top level', section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || [], []);
  check('Code.js is still version 2.22.0: this is not a release', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.22.0');
  let released = null;
  try {
    released = execFileSync('git', ['show', 'template-release/T-2026.10.12:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    released = null;
  }
  if (released === null) {
    console.log('  note: template-release/T-2026.10.12 is not available in this checkout; the comparison with it is skipped.');
  } else {
    check('without the fix, Code.js is the Code.js of T-2026.10.12, byte for byte', view.code === released, true);
    check('the released functions that changed are these six', functionNames(released).filter((name, i, list) => list.indexOf(name) === i).filter((name) => functionSource(CODE, name) !== functionSource(released, name)), CHANGED_FUNCTIONS);
    check('the new functions are those of the section; none was removed', [functionNames(CODE).filter((name) => functionNames(released).indexOf(name) === -1).sort(), functionNames(released).filter((name) => functionNames(CODE).indexOf(name) === -1)],
      [SECTION_FUNCTIONS.slice().sort(), []]);
    const creator = (ref) => execFileSync('git', ['show', ref + ':template/ReportCreator.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
    check('ReportCreator.js, HyperLink.js and the manifest are the released files',
      ['template/ReportCreator.js', 'HyperLink.js', 'appsscript.json'].filter((file) => fs.readFileSync(path.join(ROOT, file), 'utf8').replace(/\r/g, '') !==
        execFileSync('git', ['show', 'template-release/T-2026.10.12:' + file], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '')).concat(creator('template-release/T-2026.10.12') ? [] : ['?']), []);
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll report-structure checks passed.');
process.exitCode = failures ? 1 : 0;
