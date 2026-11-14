/**
 * GitHub issue #11 -- session times in a main-meeting report through the real
 * Build from Scratch and the real update (continuousUpdateCore_()), on one
 * report in a sandbox:
 *
 *   1. a build writes the Session administration section from the stored
 *      sessions, at the end of the opening, and nothing about TDocs changes
 *      because of them (same tables, same registration rows);
 *   2. an update keeps the section and the minutes typed by hand, places a
 *      new TDoc only where it belongs and never inside the section;
 *   3. editing the times replaces only the generated lines; text typed below
 *      them stays;
 *   4. idempotence: saving, updating and building again change nothing more;
 *   5. a rebuild gives the same section again.
 *
 * All names, TDocs and times are synthetic. Audio is only the reference report type.
 *
 * Run: node tests/session-times-main-build-update.test.js
 */

const { loadCode } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');

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

const HEADING = 'Session administration';
const row = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const BASE = {
  SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_NAME: 'SA4#138', MEETING_ID: '60780', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'November 16, 2026',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138/Docs/TDoc_List_Meeting_SA4%23138.xlsx',
  AGENDA_TDOC: 'S4-261800', MEETING_TIME_ZONE: 'America/Edmonton', MEETING_TIME_ZONE_MODE: 'explicit', MEETING_START_BASIS: '60780|November 16, 2026'
};
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda'], ['7', 'Audio SWG'], ['7.1', 'Opening of the session'], ['7.2', 'Registration of documents'], ['7.5', 'IVAS'], ['7.6', 'EVS'],
  ['8', 'Video SWG'], ['8.5', 'Video topic'], ['20', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
/** id, agenda item, status, uploaded (UTC). 8.5 belongs to the Video report: an Audio report never takes it. */
const TDOCS = [['S4-261801', '7.5', 'available', '2026-11-16 14:00:00'], ['S4-261802', '7.6', 'available', '2026-11-17 08:00:00'], ['S4-261803', '8.5', 'available', '2026-11-16 15:00:00']];
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

/** One main report: the real build, the real TDoc-list reader, the real update; a new TDoc table goes where the real code puts it. */
function report(props) {
  const loaded = loadCode({ documentProperties: Object.assign({}, BASE, props || {}) });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const r = { s: s, body: body, docProps: loaded.docProps, tdocs: TDOCS.slice(), inserted: [], fetches: [] };
  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' }, alert: (...args) => { ui.alerts.push(args); return 'YES'; },
    showModalDialog: (out, title) => { ui.dialogs.push({ title: title, html: out.html }); } };
  r.ui = ui;
  s.Logger = { log: () => {} };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'REPORTdoc000000000000000000000000000000000', getBody: () => body });
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  s.downloadMeetingAgenda_ = () => AGENDA;
  s.UrlFetchApp = { fetch: (url) => { r.fetches.push(String(url)); return { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }; } };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const tdocs = r.tdocs.concat([['S4-261800', '', 'available', '']]);
    const values = [HEADER].concat(tdocs.map(([id, agenda, status, uploaded]) => [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', agenda, 'Synthetic item', status, '',
      uploaded ? new Date(uploaded.replace(' ', 'T') + 'Z') : '', '', '']));
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((x) => x.map(() => null)) }),
      getRange: (rowAt, col, rows) => ({ getDisplayValues: () => [['Uploaded']].concat(tdocs.map((t) => [t[3]])).slice(rowAt - 1, rowAt - 1 + rows) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };
  r.build = () => { try { return s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }); } catch (e) { return { ok: false, message: e.message }; } };
  r.update = () => {
    s.insertNewTdoc_ = (b, td) => {
      const at = s.findInsertionPointForAgendaItem_(b, td.agendaItem, '');
      const cells = [['TDoc', td.row[td.tdocCol]], ['Title', td.row[td.titleCol]], ['Minutes', '']];
      if (at >= b.getNumChildren()) b.appendTable(cells); else b.insertTable(at, cells);
      r.inserted.push(td.row[td.tdocCol] + ' @' + td.agendaItem);
    };
    s.updateTdocStatus_ = () => false;
    s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
    s.checkRSSFeed_ = () => {};
    s.updateRevisions_ = () => {};
    return s.continuousUpdateCore_();
  };
  r.save = (rows) => s.saveAdhocSessionsConfiguration(JSON.parse(JSON.stringify(rows)));
  r.texts = () => body._children.map((c) => (c.getType() === 'TABLE' ? '[' + rowsOf(c)[0].join('|') + (rowsOf(c)[0][0] === 'TDoc' && rowsOf(c)[0].length === 2 ? '' : ' x' + (c.getNumRows() - 1)) + ']' : c.getText()));
  r.snapshot = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
  r.section = () => { const at = r.texts().indexOf(HEADING); if (at === -1) return null; const out = [HEADING]; for (let i = at + 1; i < body.getNumChildren() && body.getChild(i).getType() !== 'TABLE' && body.getChild(i).getHeading() === 'NORMAL'; i++) out.push(body.getChild(i).getText()); return out; };
  r.tdocTables = () => body._children.filter((c) => c.getType() === 'TABLE' && rowsOf(c)[0][0] === 'TDoc' && rowsOf(c)[0].length === 2).map((t) => rowsOf(t)[0][1]);
  r.registration = () => { const t = body._children.filter((c) => c.getType() === 'TABLE' && rowsOf(c)[0].join('|') === 'TDoc|Title|Source|Agenda Item')[0]; return t ? rowsOf(t).slice(1).map((x) => x[0]) : null; };
  /** Types the minutes of a TDoc by hand. */
  r.typeMinutes = (tdoc, text) => { const t = body._children.filter((c) => c.getType() === 'TABLE' && rowsOf(c)[0][1] === tdoc)[0]; t.getRow(2).getCell(1).setText(text); };
  return r;
}

const MINUTES = 'Minutes typed by hand: the TDoc was noted.';
const OPENING_MINUTES = 'The chair welcomed the delegates (typed by hand).';
const SESSIONS = [row('', '2026-11-16', '09:00', '17:30', 'Day 1'), row('', '2026-11-17', '09:00', '18:00')];
const LINES = [HEADING, 'A01: Day 1, November 16, 2026, 09:00–17:30 ABT', 'A02: November 17, 2026, 09:00–18:00 ABT'];

// ================================================================ 0. the reference: the same report without session times
const plain = report();
const plainBuild = plain.build();

// ================================================================ 1
console.log('1. Build from Scratch with session times');
const r = report();
{
  check('the sessions are saved before the build (no section yet: an empty report has no opening)', [r.save(SESSIONS).ok, r.section()], [true, null]);
  const built = r.build();
  check('the reference build: an Audio report with the two Audio TDocs', [plainBuild.ok, plainBuild.note || '', plain.tdocTables(), plain.texts()[0], plain.texts().indexOf(HEADING)], [true, '', ['S4-261801', 'S4-261802'], '7.1 Opening of the session', -1]);
  check('the build completes, without a note', [built.ok, built.note || '', built.tdocs], [true, '', 2]);
  check('the section is written from the stored sessions', r.section(), LINES);
  const texts = r.texts();
  check('at the end of the opening, directly before Registration of Documents', texts.slice(texts.indexOf(HEADING) - 1, texts.indexOf(HEADING) + 4).map((t, i) => (i === 4 ? /Registration of Documents/.test(t) : t)).slice(1),
    LINES.concat([true]));
  check('the TDocs of the report are those of the report without session times: Audio items only', [r.tdocTables(), r.registration()], [plain.tdocTables(), plain.registration()]);
  check('no Video TDoc', r.tdocTables().indexOf('S4-261803'), -1);
  check('apart from the section, the report is the report without session times', r.texts().filter((t) => LINES.indexOf(t) === -1), plain.texts());
  check('the meeting type and the report type are as stored', [r.docProps.getProperty('MEETING_TYPE'), r.docProps.getProperty('REPORT_SUFFIX')], ['main', 'Audio']);
}

// ================================================================ 2
console.log('2. Update Report Now');
{
  r.typeMinutes('S4-261801', MINUTES);
  const at = r.texts().indexOf(LINES[2]);
  r.body.insertParagraph(at + 1, OPENING_MINUTES);
  const before = r.snapshot();
  const first = r.update();
  check('an update without new TDocs: the report is unchanged', [first.success !== false, r.snapshot() === before, r.inserted], [true, true, []]);
  r.tdocs.push(['S4-261804', '7.1', 'available', '2026-11-16 16:00:00'], ['S4-261805', '7.5', 'available', '2026-11-16 17:00:00'], ['S4-261806', '8.5', 'available', '2026-11-16 17:00:00']);
  const second = r.update();
  check('the new Audio TDocs are inserted, the Video one is not', [second.success !== false, r.inserted.slice().sort()], [true, ['S4-261804 @7.1', 'S4-261805 @7.5']]);
  check('the section is unchanged', r.section().slice(0, 3), LINES);
  check('the opening minutes typed below it are still there, directly below it', r.texts()[r.texts().indexOf(LINES[2]) + 1], OPENING_MINUTES);
  check('the minutes typed into a TDoc table are kept', rowsOf(r.body._children.filter((c) => c.getType() === 'TABLE' && rowsOf(c)[0][1] === 'S4-261801')[0])[2][1], MINUTES);
  const texts = r.texts();
  const s0 = texts.indexOf(HEADING);
  check('no TDoc table lands between the heading and the generated lines', texts.slice(s0, s0 + 3), LINES);
  check('the section exists once', texts.filter((t) => t === HEADING).length, 1);
  const after = r.snapshot();
  r.update();
  check('a second update changes nothing', r.snapshot() === after, true);
}

// ================================================================ 3
console.log('3. editing the times');
{
  const before = r.texts();
  const res = r.save([row('s1', '2026-11-16', '10:00', '17:30', 'Day 1'), row('s2', '2026-11-17', '09:00', '18:00')]);
  check('saved, without a notice', [res.ok, res.changed, res.notice || ''], [true, true, '']);
  const now = r.texts();
  check('only the first line changed', now.map((t, i) => (t === before[i] ? '=' : t)).filter((t) => t !== '='), ['A01: Day 1, November 16, 2026, 10:00–17:30 ABT']);
  check('the opening minutes typed below are kept', now[now.indexOf('A02: November 17, 2026, 09:00–18:00 ABT') + 1], OPENING_MINUTES);
  const snap = r.snapshot();
  check('saving the same again changes nothing', [r.save([row('s1', '2026-11-16', '10:00', '17:30', 'Day 1'), row('s2', '2026-11-17', '09:00', '18:00')]).changed, r.snapshot() === snap], [false, true]);
  const third = r.save([row('s1', '2026-11-16', '10:00', '17:30', 'Day 1'), row('s2', '2026-11-17', '09:00', '18:00'), row('', '2026-11-18', '09:00', '12:00')]);
  check('a session added: a line added, the minutes below stay', [third.ok, r.section().slice(1), r.tdocTables().length],
    [true, ['A01: Day 1, November 16, 2026, 10:00–17:30 ABT', 'A02: November 17, 2026, 09:00–18:00 ABT', 'A03: November 18, 2026, 09:00–12:00 ABT', OPENING_MINUTES], 4]);
  const removed = r.save([]);
  check('every session removed: heading and lines go, the minutes typed below stay', [removed.ok, r.texts().indexOf(HEADING), r.texts().indexOf(OPENING_MINUTES) !== -1], [true, -1, true]);
  r.save([row('', '2026-11-16', '10:00', '17:30', 'Day 1'), row('', '2026-11-17', '09:00', '18:00')]);
  check('saved again: the section is written again at the end of the opening', r.section().slice(1, 3), ['A01: Day 1, November 16, 2026, 10:00–17:30 ABT', 'A02: November 17, 2026, 09:00–18:00 ABT']);
}

// ================================================================ 4
console.log('4. Build from Scratch again');
{
  const first = r.build();
  check('the rebuild completes', first.ok, true);
  const snap = r.snapshot();
  check('the section is there once, from the stored sessions', [r.texts().filter((t) => t === HEADING).length, r.section()],
    [1, [HEADING, 'A01: Day 1, November 16, 2026, 10:00–17:30 ABT', 'A02: November 17, 2026, 09:00–18:00 ABT']]);
  check('the TDocs: those of the list for the Audio report', r.tdocTables().slice().sort(), ['S4-261801', 'S4-261802', 'S4-261804', 'S4-261805']);
  r.build();
  check('a second rebuild gives the same report', r.snapshot() === snap, true);
  const noSessions = report();
  noSessions.save([]);
  noSessions.build();
  check('a main report without sessions is built exactly as before', noSessions.texts(), plain.texts());
}

// ================================================================ 5
console.log('5. what is never asked for');
{
  check('no fetch other than the TDoc list (no Reviewer, no Portal) in build, update and save', r.fetches.filter((u) => u !== BASE.TDOC_LIST_URL && !/TDoc_List/.test(u)), []);
  check('the meeting and report type, the TDoc list and the agenda are as stored', ['MEETING_TYPE', 'REPORT_SUFFIX', 'TDOC_LIST_URL', 'AGENDA_TDOC'].map((k) => r.docProps.getProperty(k)),
    ['main', 'Audio', BASE.TDOC_LIST_URL, 'S4-261800']);
  check('no session reaches the TDoc selection', [r.s.getAdhocSessions_(), r.s.adhocSessionsEnabled_()], [[], false]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nall checks passed');
process.exit(failures ? 1 : 0);
