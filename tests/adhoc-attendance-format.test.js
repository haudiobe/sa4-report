/**
 * Ad-hoc attendance -- the look of the generated Attendance section
 * (G2 smoke-test finding: the section came out largely in bold, and the
 * attendee table was not formatted).
 *
 * Cause: in Google Docs a table that is inserted takes its text attributes
 * from the paragraph before it. The paragraph before an attendee table is
 * the bold "Attendees" line, so every cell of the table was bold. And a new
 * table only got the report's table look (row height, paragraph spacing,
 * shaded header) at the next formatting pass. The fake document of the tests
 * now models that inheritance (tests/helpers/fake-document.js).
 *
 *   1. what is bold and what is not;
 *   2. the table: header, rows, widths;
 *   3. after a refresh, a second session, a rebuild;
 *   4. a formatting failure never fails an import;
 *   5. nothing else is formatted.
 *
 * All names and values are synthetic.
 *
 * Run: node tests/adhoc-attendance-format.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const FX = require('./fixtures/teams-attendance-synthetic.js');

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

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const SESSIONS = JSON.stringify({ v: 1, nextId: 3, sessions: [{ id: 's1', label: 'AHG Call 1', date: '2026-09-22', start: '15:00', end: '18:00' }, { id: 's2', label: '', date: '2026-09-23', start: '', end: '' }] });

function report(props) {
  const loaded = loadCode({ documentProperties: Object.assign({ MEETING_TYPE: 'adhoc', ADHOC_SESSIONS: SESSIONS }, props || {}) });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const H = s.DocumentApp.ParagraphHeading;
  body.appendParagraph('5.1 Opening of the session').setHeading(H.HEADING2);
  body.appendParagraph('The chair opens the session.');
  body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4aA269001', 'Synthetic title', 'ExampleCorp', '5.4']]);
  body.appendTable([['TDoc', 'S4aA269001'], ['Title', 'Synthetic title'], ['Minutes', 'Typed by hand.']]);
  body.appendParagraph('5.11 Close of the session').setHeading(H.HEADING2);
  body.appendParagraph('The chair closes the session.');
  const r = { s: s, body: body, docProps: loaded.docProps, logs: [] };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.getUi = () => ({ alert() {}, ButtonSet: { OK: 'OK' } });
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  s.Utilities.base64Decode = (b) => Array.from(Buffer.from(b, 'base64')).map((x) => (x > 127 ? x - 256 : x));
  const b64 = FX.utf16le(FX.buildExport()).toString('base64');
  r.import = () => s.confirmTeamsAttendanceImport('s1', b64, s.previewTeamsAttendanceImport('s1', b64).token);
  r.container = () => { const c = s.findAdhocAttendanceContainer_(body); return c ? body._children.slice(c.start, c.end) : []; };
  r.paragraphs = () => r.container().filter((e) => e.getType() === 'PARAGRAPH' && e.getText() !== '');
  r.tables = () => r.container().filter((e) => e.getType() === 'TABLE');
  r.cells = (table) => { const out = []; for (let i = 0; i < table.getNumRows(); i++) { const row = []; for (let k = 0; k < table.getRow(i).getNumCells(); k++) row.push(table.getRow(i).getCell(k)); out.push(row); } return out; };
  return r;
}
/** What a table looks like: [bold header cells, bold body cells, header backgrounds, body backgrounds, widths of the columns, row heights, paragraph spacings]. */
const look = (r, table) => {
  const cells = r.cells(table);
  const distinct = (list) => list.map((v) => JSON.stringify(v)).filter((v, i, l) => l.indexOf(v) === i).map((v) => JSON.parse(v));
  return {
    headerBold: cells[0].map((c) => c._bold),
    boldBodyCells: cells.slice(1).reduce((n, row) => n + row.filter((c) => c._bold).length, 0),
    headerBackground: distinct(cells[0].map((c) => c._background)),
    bodyBackground: distinct(cells.slice(1).reduce((all, row) => all.concat(row.map((c) => c._background)), [])),
    widths: distinct(cells.map((row) => row.map((c) => c._width))),
    rowHeights: distinct(Array.from({ length: table.getNumRows() }, (_, i) => table.getRow(i)._minHeight)),
    spacing: distinct(cells.reduce((all, row) => all.concat(row.map((c) => c._spacing)), []))
  };
};
const EXPECTED_LOOK = { headerBold: [true, true, true], boldBodyCells: 0, headerBackground: ['#D9EAF7'], bodyBackground: [null], widths: [[159, 122, 187]], rowHeights: [0], spacing: [[0, 0]] };

// ================================================================ 1. bold

console.log('1. what is bold');
const r = report();
{
  check('the import succeeds', r.import().ok, true);
  check('the fake document behaves like Google Docs here: a table inserted after a bold paragraph starts bold', (() => {
    const x = report(); const p = x.body.appendParagraph('Bold line'); p.editAsText().setBold(true);
    const t = x.body.appendTable([['a', 'b']]); const u = x.body.insertTable(0, [['c']]);
    return [t.getRow(0).getCell(0)._bold, t.getRow(0).getCell(1)._bold, u.getRow(0).getCell(0)._bold];
  })(), [true, true, false]);
  check('the paragraphs of the section: the heading, then the labels "Statistics" and "Attendees" bold and every statistics line not bold', r.paragraphs().map((p) => (p.getHeading() !== 'NORMAL' ? 'H ' : (p._bold ? 'B ' : '- ')) + p.getText().split(':')[0]),
    ['H Attendance', 'B Statistics', '- Meeting', '- Session', '- Start', '- End', '- Duration', '- Attendance records', '- Average attendance', 'B Attendees']);
  const table = r.tables()[0];
  check('the attendee table follows the bold "Attendees" line', [r.container()[r.container().indexOf(table) - 1].getText(), r.container()[r.container().indexOf(table) - 1]._bold], ['Attendees', true]);
  check('no attendee cell is bold: 23 rows, 69 cells, none of them', [table.getNumRows() - 1, r.cells(table).slice(1).reduce((n, row) => n + row.length, 0), look(r, table).boldBodyCells], [23, 69, 0]);
  check('the header row is bold', look(r, table).headerBold, [true, true, true]);
  check('bold in the whole section: the two labels and the three header cells, nothing else', r.paragraphs().filter((p) => p._bold).length + r.cells(table).reduce((n, row) => n + row.filter((c) => c._bold).length, 0), 5);
}

// ================================================================ 2. the table

console.log('2. the attendee table');
{
  const table = r.tables()[0];
  check('header shaded like every table of a report, body not shaded; rows as low as their text; no paragraph spacing in any cell; the same column widths in every row', look(r, table), EXPECTED_LOOK);
  check('the widths add up to the usable page width, and the e-mail column is the widest', [159 + 122 + 187, 187 > 159 && 159 > 122], [468, true]);
  check('the header shading is the one of the report\'s formatting pass (source)', [/headerRow\.getCell\(c\)\.setBackgroundColor\('#D9EAF7'\)/.test(CODE), /const ADHOC_TABLE_HEADER_BACKGROUND_ = '#D9EAF7';/.test(CODE)], [true, true]);
  check('the e-mail links are still there, as mailto: links, on the cells that have an address', (() => {
    const rows = r.cells(table).slice(1);
    return [rows.filter((row) => row[2].getText() !== '').every((row) => row[2]._links.length === 1 && row[2]._links[0][2] === 'mailto:' + row[2].getText()), rows.filter((row) => row[2].getText() === '').every((row) => row[2]._links.length === 0),
      rows.filter((row) => row[2].getText() !== '').length > 10];
  })(), [true, true, true]);
  check('the texts of the table are what they were', [r.cells(table)[0].map((c) => c.getText()), r.cells(table)[1].map((c) => c.getText())], [['Name', 'Company', 'Email'], ['Alex Organizer', '', 'alex.organizer@example.com']]);
  const wide = report();
  wide.s.getConfig_ = () => ({ TDOC_PAGE_USABLE_WIDTH: '700' });
  wide.import();
  check('a report whose configuration table states another usable page width (as for the TDoc tables) gets the same proportions', look(wide, wide.tables()[0]).widths, [[238, 182, 280]]);
  const odd = report();
  odd.s.getConfig_ = () => ({ TDOC_PAGE_USABLE_WIDTH: 'wide' });
  odd.import();
  check('a page width that cannot be read: the default one', look(odd, odd.tables()[0]).widths, [[159, 122, 187]]);
}

// ================================================================ 3. refresh, second session, rebuild hook

console.log('3. after a refresh and with a second session');
{
  r.s.refreshAdhocAttendanceSection();
  r.s.refreshAdhocAttendanceSection();
  check('after two refreshes: one table, the same look', [r.tables().length, look(r, r.tables()[0])], [1, EXPECTED_LOOK]);
  r.s.writeAdhocAttendanceValue_(r.docProps, 's2', JSON.stringify({ t: 'Second day', n: 2, s: '2026-09-23T09:00:00', e: null, d: null, a: null, p: [['Pat Kim', 'ExampleCorp', 'pat.kim@example.com'], ['Robin Lee', '', '']] }));
  r.s.refreshAdhocAttendanceSection();
  check('with a second session: both tables have the look; the second one follows a bold "Attendees" line as well', [r.tables().length, look(r, r.tables()[0]), look(r, r.tables()[1])], [2, EXPECTED_LOOK, EXPECTED_LOOK]);
  check('the paragraphs of the second block', r.paragraphs().slice(10).map((p) => (p._bold ? 'B ' : '- ') + p.getText().split(':')[0]), ['B Statistics', '- Meeting', '- Session', '- Start', '- Attendance records', 'B Attendees']);
  check('a Company typed into a table is kept through a refresh, and the cell is still not bold', (() => {
    r.tables()[1].getRow(2).getCell(1).setText('TypedCorp');
    r.s.refreshAdhocAttendanceSection();
    const cell = r.tables()[1].getRow(2).getCell(1);
    return [cell.getText(), cell._bold];
  })(), ['TypedCorp', false]);
  check('after a rebuild (the section is written into a cleared document): the same look', (() => {
    r.s.beginAdhocAttendanceRebuild_();
    r.body.clear();
    r.body.appendParagraph('5.11 Close of the session').setHeading(r.s.DocumentApp.ParagraphHeading.HEADING2);
    r.s.finishAdhocAttendanceRebuild_(r.body, { active: true });
    return [r.tables().length, look(r, r.tables()[0]), look(r, r.tables()[1])];
  })(), [2, EXPECTED_LOOK, EXPECTED_LOOK]);
}

// ================================================================ 4. a formatting failure

console.log('4. a formatting failure never fails an import');
{
  const f = report();
  f.s.formatAdhocAttendeeTable_ = () => { throw new Error('Service Documents failed'); };
  const result = f.import();
  check('the import succeeds, the attendance is stored, the section and its table are written, the links are set', [result.ok, result.rendered, Object.keys(f.s.readAdhocAttendance_(f.docProps).sessions), f.tables().length, f.tables()[0].getNumRows() - 1,
    f.tables()[0].getRow(1).getCell(2)._links.length], [true, true, ['s1'], 1, 23, 1]);
  check('the failure is logged with its reason, and nothing else', f.logs.filter((l) => /^Attendance/.test(l)), ['Attendance: the attendee table could not be formatted: Service Documents failed']);
  const g = report();
  g.s.getConfig_ = () => { throw new Error('no configuration'); };
  g.import();
  check('a configuration that cannot be read: the default widths, no failure', [look(g, g.tables()[0]).widths, g.logs.filter((l) => /could not be formatted/.test(l))], [[[159, 122, 187]], []]);
}

// ================================================================ 5. nothing else

console.log('5. nothing else is formatted');
{
  const x = report();
  const others = x.body._children.filter((e) => e.getType() === 'TABLE');
  x.import();
  x.s.refreshAdhocAttendanceSection();
  const untouched = (table) => x.cells(table).every((row) => row.every((c) => c._bold === false && c._background === null && c._width === null && c._spacing[0] === null)) && Array.from({ length: table.getNumRows() }, (_, i) => table.getRow(i)._minHeight).every((h) => h === null);
  check('the registration table and the TDoc table of the report are not formatted by an import or a refresh', others.map(untouched), [true, true]);
  check('the paragraphs outside the section are not touched', x.body._children.filter((e) => e.getType() === 'PARAGRAPH' && x.container().indexOf(e) === -1).map((p) => [p.getText(), p._bold === undefined]),
    [['5.1 Opening of the session', true], ['The chair opens the session.', true], ['5.11 Close of the session', true], ['The chair closes the session.', true]]);
  const callers = CODE.split('\n').filter((line) => line.indexOf('formatAdhocAttendeeTable_(') !== -1 && !/^function |^\s*(\/|\*)/.test(line));
  check('the formatting is applied in one place: where the section writes an attendee table (source)', callers.map((l) => l.trim()), ['formatAdhocAttendeeTable_(table);']);
  check('it formats the table it is given and reads no attendee data (source)', (() => {
    const at = CODE.indexOf('\nfunction formatAdhocAttendeeTable_(');
    const src = CODE.slice(at, CODE.indexOf('\n}\n', at));
    return [/getText\(|\.name|\.email|\.company|Logger\.log|setProperty|removeChild|setText\(/.test(src), /getBody|getTables/.test(src)];
  })(), [false, false]);
  const main = report({ MEETING_TYPE: 'main' });
  check('a main-meeting report: an import is refused and nothing is formatted', [main.s.previewTeamsAttendanceImport('s1', FX.utf16le(FX.buildExport()).toString('base64')).ok, main.tables().length], [false, 0]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll attendance format checks passed.');
process.exitCode = failures ? 1 : 0;
