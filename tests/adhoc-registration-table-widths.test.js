/**
 * Ad-hoc sessions -- the column widths of the registration table once it has
 * the Session column (G2 smoke-test finding: with five columns the Title
 * column was far too narrow, so the table wrapped heavily and became long).
 *
 * Cause: no width was ever set on the registration table. Google Docs then
 * gives every column the same share of the page: a quarter each with the
 * four columns of every release, a fifth each with the Session column.
 *
 *   1. Build Report from Scratch with sessions: the five widths;
 *   2. a report without sessions: four columns, no width set (as before);
 *   3. Update Report Now: nothing is written again; a new row gets the widths;
 *   4. a five-column table made before the widths were set, and one that was
 *      changed by hand: the next update sets them, the one after writes nothing;
 *   5. adding the Session column to an existing table; saving assignments;
 *   6. only the registration table: the other tables, the content, the links;
 *   7. a problem with the widths never fails a build or an update.
 *
 * All names and values are synthetic.
 *
 * Run: node tests/adhoc-registration-table-widths.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
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

const WIDTHS = [89, 197, 89, 56, 37];   // 19 / 42 / 19 / 12 / 8 % of the 468 pt a report page has between its margins
const session = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const SESSIONS = JSON.stringify({ v: 1, nextId: 3, sessions: [session('s1', '2026-09-22', '15:00', '18:00'), session('s2', '2026-09-23', '15:00', '18:00')] });
const PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Topic'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
/** id, agenda item, status, uploaded (UTC) */
const TDOCS = [['S4aA269001', '4.3', 'available', '2026-09-22 14:00:00'], ['S4aA269002', '4.4', 'available', '2026-09-23 08:00:00'], ['S4aA269003', '4.4', 'available', '2026-09-23 22:00:00']];

function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));
const widthsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getWidth()));
const cellsOf = (t) => [].concat(...Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k))));
const all = (n, widths) => Array.from({ length: n }, () => widths.slice());

/** One report: the real build, the real TDoc-list reader, the real update. `sessions` false: an ad-hoc report without sessions. */
function report(sessions) {
  const loaded = loadCode({ documentProperties: Object.assign({}, PROPS, sessions === false ? {} : { ADHOC_SESSIONS: SESSIONS }) });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const r = { s: s, body: body, docProps: loaded.docProps, tdocs: TDOCS.slice(), logs: [] };
  const ui = { alerts: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' }, alert: (...args) => { ui.alerts.push(args); return 'YES'; }, showModalDialog: () => {} };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.Utilities.formatDate = (date, zone) => formatInZone(date, zone);
  s.Utilities.base64Decode = (b64) => Array.from(Buffer.from(b64, 'base64')).map((b) => (b > 127 ? b - 256 : b));
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  s.downloadMeetingAgenda_ = () => AGENDA;
  s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }) };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const values = [HEADER].concat(r.tdocs.map(([id, agenda, status, uploaded]) => [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', agenda, 'Synthetic item', status, '',
      uploaded ? new Date(uploaded.replace(' ', 'T') + 'Z') : '', '', '']));
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((row) => row.map(() => null)) }),
      getRange: (row, col, rows) => ({ getDisplayValues: () => [['Uploaded']].concat(r.tdocs.map((t) => [t[3]])).slice(row - 1, row - 1 + rows) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };

  r.build = () => { try { return s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }); } catch (e) { return { ok: false, message: e.message }; } };
  r.update = () => {
    s.insertNewTdoc_ = (b, td) => {
      const at = s.findInsertionPointForAgendaItem_(b, td.agendaItem, '');
      const cells = [['TDoc', td.row[td.tdocCol]], ['Title', td.row[td.titleCol]], ['Minutes', '']];
      if (at >= b.getNumChildren()) b.appendTable(cells); else b.insertTable(at, cells);
    };
    s.updateTdocStatus_ = () => false;
    s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
    s.checkRSSFeed_ = () => {};
    s.updateRevisions_ = () => {};
    return s.continuousUpdateCore_();
  };
  r.tables = () => body._children.filter((c) => c.getType() === 'TABLE');
  r.registration = () => r.tables().filter((t) => rowsOf(t)[0][1] === 'Title' && rowsOf(t)[0].length > 2)[0];
  r.detailTables = () => r.tables().filter((t) => rowsOf(t)[0][0] === 'TDoc' && rowsOf(t)[0].length === 2);
  /** Counts the widths written from now on, on every cell the registration table has at this moment. */
  r.countWrites = () => {
    const counter = { n: 0 };
    cellsOf(r.registration()).forEach((cell) => { if (cell._counted) { cell._counted = counter; return; } const set = cell.setWidth; cell._counted = counter; cell.setWidth = (v) => { cell._counted.n++; return set(v); }; });
    return counter;
  };
  return r;
}

// ================================================================ 1. Build Report from Scratch

console.log('1. Build Report from Scratch with sessions: the five widths');
{
  const r = report();
  check('the build completes', r.build().ok, true);
  const t = r.registration();
  check('the registration table has the Session column', rowsOf(t)[0], ['TDoc', 'Title', 'Source', 'Agenda Item', 'Session']);
  check('every row, the header too, has the widths 89 / 197 / 89 / 56 / 37 pt', widthsOf(t), all(4, WIDTHS));
  check('together they are the width of the page between its margins', WIDTHS.reduce((a, b) => a + b, 0), 468);
  check('Title is the widest column by far; Agenda Item and Session are compact; Source has the room of the TDoc column',
    [WIDTHS[1] >= 2 * WIDTHS[0], WIDTHS[1] >= 2 * WIDTHS[2], Math.max(WIDTHS[3], WIDTHS[4]) <= 60, WIDTHS[2] >= WIDTHS[0]], [true, true, true, true]);
  // A TDoc number -- ten characters such as "S4aA269001" -- has to stay on one line. With 70 pt (15 % of the page) it broke
  // before its last digit in the real document. What is protected here is the width that was chosen for it, not a
  // simulation of how Google Docs wraps: 89 pt, a good quarter more than the width that was one character short.
  check('the TDoc column is 89 pt wide, 19 % of the page: room for a ten-character TDoc number on one line',
    [widthsOf(t).map((row) => row[0]), WIDTHS[0] >= 89, WIDTHS[0] >= 70 * 1.25, rowsOf(t).slice(1).map((row) => row[0].length)], [[89, 89, 89, 89], true, true, [10, 10, 10]]);
  check('the share of the TDoc column is 19 % of the page, whatever the page width (source)',
    require('fs').readFileSync(require('./helpers/load-code.js').CODE_JS_PATH, 'utf8').indexOf('const ADHOC_REGISTRATION_TABLE_WIDTHS_ = [0.19, 0.42, 0.19, 0.12, 0.08];') !== -1, true);
  check('the shares are 19 / 42 / 19 / 12 / 8 % of the page', WIDTHS.map((w) => Math.round(w / 4.68)), [19, 42, 19, 12, 8]);
  check('the content of the table is what the build wrote', rowsOf(t).slice(1), [['S4aA269001', 'Synthetic title of S4aA269001', 'ExampleCorp', '4.3', 'A01'],
    ['S4aA269002', 'Synthetic title of S4aA269002', 'ExampleCorp', '4.4', 'A02'], ['S4aA269003', 'Synthetic title of S4aA269003', 'ExampleCorp', '4.4', '–']]);
  const again = report();
  again.build(); again.build();
  check('a second build gives the same table with the same widths', [rowsOf(again.registration()), widthsOf(again.registration())], [rowsOf(t), all(4, WIDTHS)]);
}

// ================================================================ 2. without sessions

console.log('\n2. a report without sessions: four columns and no width set, as in every release');
{
  const r = report(false);
  check('the build completes', r.build().ok, true);
  const t = r.registration();
  check('four columns', rowsOf(t)[0], ['TDoc', 'Title', 'Source', 'Agenda Item']);
  check('no width is set on any cell', widthsOf(t), all(4, [null, null, null, null]));
  const counter = r.countWrites();
  r.tdocs.push(['S4aA269004', '4.4', 'available', '2026-09-23 09:00:00']);
  r.update();
  check('an update that adds a row: five rows, still no width set anywhere, nothing written', [r.registration().getNumRows(), widthsOf(r.registration()), counter.n], [5, all(5, [null, null, null, null]), 0]);
  check('the helper leaves a four-column table alone', [r.s.applyRegistrationTableWidths_(r.registration()), widthsOf(r.registration())], [0, all(5, [null, null, null, null])]);
}

// ================================================================ 3. Update Report Now

console.log('\n3. Update Report Now: nothing is written again; a new row gets the widths');
{
  const r = report();
  r.build();
  const counter = r.countWrites();
  const before = [rowsOf(r.registration()), widthsOf(r.registration())];
  r.update(); r.update(); r.update();
  check('three updates without a new TDoc: no width is written, table and widths as they were', [counter.n, rowsOf(r.registration()), widthsOf(r.registration())], [0, before[0], before[1]]);

  r.tdocs.push(['S4aA269004', '4.4', 'available', '2026-09-23 09:00:00']);
  r.update();
  const t = r.registration();
  check('an update with a new TDoc: the new row has five cells with the widths', [t.getNumRows(), rowsOf(t)[4], widthsOf(t)[4]], [5, ['S4aA269004', 'Synthetic title of S4aA269004', 'ExampleCorp', '4.4', 'A02'], WIDTHS]);
  check('every row has the widths, and no cell that had its width was written again', [widthsOf(t), counter.n], [all(5, WIDTHS), 0]);
  const after = r.countWrites();
  r.update(); r.update();
  check('the updates after that write nothing', [after.n, widthsOf(r.registration())], [0, all(5, WIDTHS)]);
}

// ================================================================ 4. tables that do not have the widths

console.log('\n4. a five-column table without the widths, or changed by hand: set once, then left');
{
  const r = report();
  r.build();
  // As a report built by an earlier candidate has it: five columns, no width set.
  cellsOf(r.registration()).forEach((cell) => cell.setWidth(null));
  check('(the table has no widths)', widthsOf(r.registration()), all(4, [null, null, null, null, null]));
  const counter = r.countWrites();
  r.update();
  check('the next update sets all of them', [widthsOf(r.registration()), counter.n], [all(4, WIDTHS), 20]);
  r.update(); r.update();
  check('and the updates after it write nothing more', counter.n, 20);

  // Equal fifths, as Google Docs lays out a table without widths, written out.
  cellsOf(r.registration()).forEach((cell) => cell.setWidth(93.6));
  const second = r.countWrites();
  r.update();
  check('a table of five equal columns is set to the widths', [widthsOf(r.registration()), second.n], [all(4, WIDTHS), 20]);

  r.registration().getRow(2).getCell(1).setWidth(120);
  const third = r.countWrites();
  r.update();
  check('one cell that differs: that cell is set, no other', [widthsOf(r.registration()), third.n], [all(4, WIDTHS), 1]);
  const content = rowsOf(r.registration());
  r.update();
  check('no update changes the content of the table', [rowsOf(r.registration()), third.n], [content, 1]);
}

// ================================================================ 5. adding the column; saving assignments

console.log('\n5. adding the Session column to an existing table; saving assignments');
{
  const r = report(false);
  r.build();
  const before = rowsOf(r.registration());
  check('(a four-column table without widths)', [before[0].length, widthsOf(r.registration())], [4, all(4, [null, null, null, null])]);
  r.docProps.setProperty('ADHOC_SESSIONS', SESSIONS);
  const done = r.s.enableAdhocSessionColumn();
  const t = r.registration();
  check('the column is added', [done.ok, done.changed, rowsOf(t)[0]], [true, true, ['TDoc', 'Title', 'Source', 'Agenda Item', 'Session']]);
  check('the four columns that were there keep their content, in every row', rowsOf(t).map((row) => row.slice(0, 4)), before);
  check('every row has the five widths', widthsOf(t), all(4, WIDTHS));
  const counter = r.countWrites();
  check('doing it again changes nothing', [r.s.enableAdhocSessionColumn().changed, counter.n, widthsOf(r.registration())], [false, 0, all(4, WIDTHS)]);
  r.update();
  check('nor does the update after it', [counter.n, widthsOf(r.registration())], [0, all(4, WIDTHS)]);

  // Saving a manual assignment brings the column up to date -- and with it the widths of a table that lacks them.
  cellsOf(r.registration()).forEach((cell) => cell.setWidth(null));
  const saved = r.s.saveAdhocTdocSessions({ clock: 'utc', tdocs: { S4aA269001: { mode: 'set', sessions: ['s2'] } } });
  check('saving an assignment: the cell is written and the table has its widths', [saved.ok, rowsOf(r.registration())[1][4], widthsOf(r.registration())], [true, 'A02', all(4, WIDTHS)]);
}

// ================================================================ 6. only the registration table

console.log('\n6. only the registration table: the other tables, the content, the links');
{
  const r = report();
  r.build();
  check('the tables of the TDocs have no width from this (their own formatting pass sets them)', r.detailTables().map((t) => widthsOf(t).every((row) => row.every((w) => w === null))), [true, true, true]);
  const b64 = FX.utf16le(FX.buildExport()).toString('base64');
  check('(an attendance import)', r.s.confirmTeamsAttendanceImport('s1', b64, r.s.previewTeamsAttendanceImport('s1', b64).token).ok, true);
  const attendees = r.tables().filter((t) => rowsOf(t)[0][0] === 'Name')[0];
  r.update();
  check('the attendee table keeps its own widths', [widthsOf(attendees)[0], widthsOf(attendees).every((row) => JSON.stringify(row) === JSON.stringify([159, 122, 187]))], [[159, 122, 187], true]);
  check('the tables of the TDocs still have no width after an update', r.detailTables().map((t) => widthsOf(t).every((row) => row.every((w) => w === null))), [true, true, true]);

  // The helper itself.
  const S = r.s;
  const other = r.body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item', 'Notes'], ['S4aA269001', 'x', 'y', '4.3', 'z']]);
  const two = r.body.appendTable([['TDoc', 'S4aA269009'], ['Title', 'Synthetic']]);
  const three = r.body.appendTable([['Name', 'Company', 'Email'], ['Pat Kim', 'ExampleCorp', 'pat.kim@example.com']]);
  check('a table of five columns that is not the registration table, a TDoc table and a three-column table are not touched',
    [S.applyRegistrationTableWidths_(other), S.applyRegistrationTableWidths_(two), S.applyRegistrationTableWidths_(three), widthsOf(other), widthsOf(two), widthsOf(three)],
    [0, 0, 0, all(2, [null, null, null, null, null]), all(2, [null, null]), all(2, [null, null, null])]);
  check('nothing and something that is no table: 0, no exception', [S.applyRegistrationTableWidths_(null), S.applyRegistrationTableWidths_({})], [0, 0]);

  // A link on the TDoc number, the header look and the texts are not touched by the widths.
  const linked = makeFakeDocumentBody(S);
  const url = 'https://www.example.org/docs/S4aA269001.zip';
  const resolver = S.makeAdhocTdocSessionResolver_();
  const tdoc = { row: ['S4aA269001', 'Synthetic title', 'ExampleCorp'], richTextRow: [{ getLinkUrl: () => url }, null, null], agendaItem: '4.3', tdocCol: 0, titleCol: 1, sourceCol: 2,
    uploaded: { at: '2026-09-22T14:00:00', dateOnly: null, unreadable: false } };
  const made = S.createSummaryTable_(linked, [tdoc], 0, 1, 2, -1, resolver);
  check('a table made with a linked TDoc number: content, link and widths', [rowsOf(made), made.getRow(1).getCell(0)._links.length, JSON.stringify(made.getRow(1).getCell(0)._links).indexOf(url) !== -1, widthsOf(made)],
    [[['TDoc', 'Title', 'Source', 'Agenda Item', 'Session'], ['S4aA269001', 'Synthetic title', 'ExampleCorp', '4.3', 'A01']], 1, true, all(2, WIDTHS)]);
  const look = cellsOf(made).map((c) => [c._bold, c._background, JSON.stringify(c._links)]);
  made.getRow(1).getCell(1).setWidth(10);
  check('setting the widths again writes the one width and nothing else of a cell', [S.applyRegistrationTableWidths_(made), cellsOf(made).map((c) => [c._bold, c._background, JSON.stringify(c._links)]), rowsOf(made)[1]],
    [1, look, ['S4aA269001', 'Synthetic title', 'ExampleCorp', '4.3', 'A01']]);
  const plain = S.createSummaryTable_(makeFakeDocumentBody(S), [tdoc], 0, 1, 2, -1);
  check('the same table made without sessions: four columns, no width', [rowsOf(plain)[0].length, widthsOf(plain)], [4, all(2, [null, null, null, null])]);
}

// ================================================================ 7. a problem with the widths

console.log('\n7. a problem with the widths never fails a build or an update');
{
  const r = report();
  r.build();
  cellsOf(r.registration()).forEach((cell) => { cell.setWidth(null); cell.setWidth = () => { throw new Error('synthetic width failure'); }; });
  r.tdocs.push(['S4aA269004', '4.4', 'available', '2026-09-23 09:00:00']);
  let thrown = '';
  try { r.update(); } catch (e) { thrown = e.message; }
  const t = r.registration();
  check('the update completes, and the new row and its Session cell are there', [thrown, t.getNumRows(), rowsOf(t)[4]], ['', 5, ['S4aA269004', 'Synthetic title of S4aA269004', 'ExampleCorp', '4.4', 'A02']]);
  check('the problem is logged, without any content of the report', [r.logs.filter((l) => /column widths of the registration table could not be set/.test(l)).length > 0, r.logs.filter((l) => /column widths/.test(l)).some((l) => /S4aA|Synthetic title|ExampleCorp/.test(l))], [true, false]);
  check('the helper returns 0 and does not throw', r.s.applyRegistrationTableWidths_(t), 0);

  const b = report();
  const original = b.s.registrationTableColumns_;
  let calls = 0;
  b.s.registrationTableColumns_ = (table) => { calls++; return original(table); };
  b.s.getConfig_ = () => { throw new Error('synthetic config failure'); };
  check('a build when the page width cannot be read: completes, with the widths of the usual page', [b.build().ok, widthsOf(b.registration())], [true, all(4, WIDTHS)]);
  b.s.getConfig_ = () => ({ TDOC_PAGE_USABLE_WIDTH: '500' });
  b.update();
  check('a report whose configuration gives another page width: the same shares of that width', widthsOf(b.registration())[0], [95, 210, 95, 60, 40]);
}

console.log(failures === 0 ? '\nAll registration-table width checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
