/**
 * Ad-hoc sessions, stage B -- the session(s) of a registered TDoc and the
 * Session column of the registration table
 * (docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md §6-§8, §12.1).
 *
 *   1. The "Uploaded" column of the TDoc list (stage B0).
 *   2. Automatic assignment (pure).
 *   3. Manual assignments: the model.
 *   4. Manual assignments: the stored property.
 *   5. The clock of the upload times.
 *   6. A new build: four or five columns.
 *   7. An existing report: updates, and the explicit action.
 *   8. Formal revisions.
 *   9. Sessions that change; sessions that are named are not removed.
 *  10. The dialog.
 *  11. Menu, adoption, and nothing else changed.
 *
 * The TDoc list used here is synthetic. Its "Uploaded" column has the form
 * observed in real Portal lists: a date-time cell shown as
 * "yyyy-MM-dd HH:mm:ss", without a time zone, empty for a reserved TDoc.
 *
 * Run: node tests/adhoc-tdoc-sessions.test.js
 */

const fs = require('fs');
const vm = require('vm');
const { execFileSync } = require('child_process');
const path = require('path');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
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
const thrown = (fn) => { try { fn(); return null; } catch (e) { return e.message; } };

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
const KEY = 'ADHOC_TDOC_SESSIONS';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.4', flavor: 'template', codeVersion: '0.0.0', gitCommit: 'abcdef0', gitTag: 'template-release/T-2026.10.4', templateDocumentId: TEMPLATE_ID };
const DASH = '–';

// ------------------------------------------------------------------ sessions

const session = (id, date, start, end, label) => ({ id: id, label: label || '', date: date, start: start || '', end: end || '' });
const sessionsProperty = (list, nextId) => JSON.stringify({ v: 1, nextId: nextId || list.length + 1, sessions: list });
// s1: 22 Sep 15:00-18:00 (cut-off 18:00).  s2: 23 Sep, no times (cut-off end of day).  s3: 24 Sep.
const S1 = session('s1', '2026-09-22', '15:00', '18:00');
const S2 = session('s2', '2026-09-23');
const S3 = session('s3', '2026-09-24');
const SESSIONS = [S1, S2, S3];

// ------------------------------------------------------------------ a synthetic TDoc list

const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
const UP = HEADER.indexOf('Uploaded');
/** id, agenda item, status, uploaded ('' = not uploaded), revision of, revised to */
const TDOCS = [
  // The times are UTC, as the Portal exports them; the comments give the time in the report's zone (Europe/Berlin, UTC+2 in September).
  ['S4aA269001', '4.3', 'available', '2026-09-18 07:26:14', '', ''],            // 09:26:14 on the 18th: before the first session
  ['S4aA269002', '4.3', 'revised', '2026-09-22 14:00:00', '', 'S4aA269010'],     // 16:00:00: during session 1
  ['S4aA269003', '4.3', 'available', '2026-09-22 16:00:00', '', ''],            // 18:00:00: exactly at its cut-off
  ['S4aA269004', '4.4', 'available', '2026-09-22 16:00:01', '', ''],            // 18:00:01: one second later
  ['S4aA269005', '4.4', 'available', '2026-09-23 21:59:59', '', ''],            // 23:59:59: the last second of session 2's day
  ['S4aA269006', '4.4', 'available', '2026-09-23 22:00:00', '', ''],            // 00:00:00 on the 24th: the first second of session 3's day
  ['S4aA269007', '4.4', 'available', '2026-09-25 09:00:00', '', ''],            // 11:00:00 on the 25th: after the last session
  ['S4aA269008', '4.4', 'reserved', '', '', ''],                                // reserved, not uploaded
  ['S4aA269009', '4.4', 'available', 'soon', '', ''],                           // a value that is not a time
  ['S4aA269010', '4.3', 'available', '2026-09-23 08:00:00', 'S4aA269002', '']   // 10:00:00 on the 23rd: the revision of 9002, uploaded a day later
];
// With the default clock (UTC, converted to the report's zone).
const EXPECTED_AUTO = { S4aA269001: 's1', S4aA269002: 's1', S4aA269003: 's1', S4aA269004: 's2', S4aA269005: 's2', S4aA269006: 's3', S4aA269007: '', S4aA269008: '', S4aA269009: '', S4aA269010: 's2' };

/** Rows as getValues() gives them: an uploaded time is a Date (here made as in a sheet whose zone is UTC). */
function listValues(tdocs, options) {
  const o = options || {};
  const header = o.header || HEADER;
  return [header].concat((tdocs || TDOCS).map(([id, agenda, status, uploaded, revisionOf, revisedTo]) => {
    const row = [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', agenda, 'Synthetic item', status, '18/09/2026 07:00:00',
      /^\d{4}-/.test(uploaded) ? new Date(uploaded.replace(' ', 'T') + 'Z') : uploaded, revisionOf, revisedTo];
    return header === HEADER ? row : header.map((h) => row[HEADER.indexOf(h)]);
  }));
}
/** What getDisplayValues() gives for the Uploaded column. */
const listDisplay = (tdocs) => [['Uploaded']].concat((tdocs || TDOCS).map((t) => [t[3]]));

/** "yyyy-MM-dd'T'HH:mm:ss" of a Date in a time zone, as Utilities.formatDate() does. */
function formatInZone(date, zone) {
  const parts = {};
  new Intl.DateTimeFormat('en-GB', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' })
    .formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return parts.year + '-' + parts.month + '-' + parts.day + 'T' + parts.hour + ':' + parts.minute + ':' + parts.second;
}

// ------------------------------------------------------------------ a report

const FTP_AUDIO = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const ADHOC = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: FTP_AUDIO, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const WITH_SESSIONS = Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty(SESSIONS) });
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Topic'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));

/**
 * One report in a sandbox, with the real build, the real TDoc-list reader and
 * the real update. opts: { props, template, docId, tdocs (the list), display
 * (false: the sheet cannot give display values), zone, header }.
 */
function report(opts) {
  const o = opts || {};
  const loaded = o.template ? loadTemplateRuntime({ release: RELEASE, documentProperties: o.props || WITH_SESSIONS }) : loadCode({ documentProperties: o.props || WITH_SESSIONS });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const r = { s: s, body: body, docProps: loaded.docProps, scriptProps: loaded.scriptProps, tdocs: o.tdocs || TDOCS, display: o.display, header: o.header, zone: o.zone === undefined ? 'UTC' : o.zone,
    downloads: 0, displayReads: 0, logs: [], writes: [], events: [] };
  const ui = { alerts: [], dialogs: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' } };
  ui.alert = (...args) => { ui.alerts.push(args); return 'YES'; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title: title, html: out.html }); };
  r.ui = ui;
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => o.docId || REPORT_ID, getBody: () => body });
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.Utilities.formatDate = (date, zone) => formatInZone(date, zone);
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  const lock = { busy: false, tries: 0, releases: 0 };
  r.lock = lock;
  s.LockService = { getDocumentLock: () => ({ tryLock: () => { lock.tries++; return !lock.busy; }, releaseLock: () => { lock.releases++; } }) };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => { r.events.push('formatting'); };
  s.downloadMeetingAgenda_ = () => AGENDA;
  s.UrlFetchApp = { fetch: () => { r.downloads++; return { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }; } };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const values = listValues(r.tdocs, { header: r.header });
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((row) => row.map(() => null)) }) };
    if (r.display !== false) {
      sheet.getRange = (row, col, rows, cols) => ({ getDisplayValues: () => { r.displayReads++; r.lastRange = [row, col, rows, cols]; return (r.display || listDisplay(r.tdocs)).slice(row - 1, row - 1 + rows); } });
    }
    const book = { getSheets: () => [sheet] };
    if (r.zone !== null) book.getSpreadsheetTimeZone = () => r.zone;
    return book;
  } };
  const set = loaded.docProps.setProperty;
  loaded.docProps.setProperty = (k, v) => { r.writes.push(k); set(k, v); };

  r.groups = () => s.downloadAndGroupTdocs_(s.getReportConfig_());
  r.all = () => s.flattenTdocGroups_(r.groups());
  r.build = () => { try { return s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }); } catch (e) { return { ok: false, message: e.message }; } };
  /** A complete update: the real continuousUpdateCore_(); the TDoc tables and the collector are recorded, not run. */
  r.update = () => {
    s.insertNewTdoc_ = (b, td) => { r.events.push('new ' + td.row[td.tdocCol]); };
    s.updateTdocStatus_ = () => false;
    s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
    s.checkRSSFeed_ = () => {};
    s.updateRevisions_ = () => {};
    return s.continuousUpdateCore_();
  };
  r.snapshot = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
  r.registration = () => { const found = s.findRegistrationTable_(body); return found ? rowsOf(found.table) : null; };
  r.sessionCells = () => { const rows = r.registration(); const out = {}; rows.slice(1).sort((a, b) => (a[0] < b[0] ? -1 : 1)).forEach((row) => { out[row[0]] = row[4]; }); return out; };
  r.save = (clock, tdocs) => s.saveAdhocTdocSessions({ clock: clock, tdocs: tdocs });
  return r;
}
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));
// The Session column shows the codes of the sessions (A01, A02, ...), by their chronological order.
const autoLabels = (map) => { const labels = { s1: 'A01', s2: 'A02', s3: 'A03', '': DASH }; const out = {}; Object.keys(map).forEach((k) => { out[k] = labels[map[k]]; }); return out; };

const S = loadCode().sandbox;

// ================================================================ 1. the Uploaded column

console.log('1. the "Uploaded" column of the TDoc list (stage B0)');
{
  const n = S.normalizeTdocUploadValue_;
  const at = (text) => ({ at: text, dateOnly: null, unreadable: false });
  const none = { at: null, dateOnly: null, unreadable: false };
  const bad = { at: null, dateOnly: null, unreadable: true };
  const zoneFormat = (zone) => (date) => formatInZone(date, zone);

  check('the observed form: the cell as shown, "yyyy-MM-dd HH:mm:ss"', n('2026-09-18 07:26:14', new Date('2026-09-18T07:26:14.512Z'), zoneFormat('UTC')), at('2026-09-18T07:26:14'));
  check('the shown text is used as it is: no time zone is applied to it', [n('2026-09-18 07:26:14', new Date('2026-09-18T05:26:14Z'), zoneFormat('Asia/Tokyo')), n('2026-09-18 07:26:14', null, null)], [at('2026-09-18T07:26:14'), at('2026-09-18T07:26:14')]);
  check('other spellings of the same', ['2026-09-18T07:26:14', ' 2026-09-18 07:26:14 ', '2026-09-18 07:19', '2026-09-18 07:26:14.512'].map((v) => n(v, null, null).at), ['2026-09-18T07:26:14', '2026-09-18T07:26:14', '2026-09-18T07:19:00', '2026-09-18T07:26:14']);
  check('an empty cell: not uploaded', [n('', '', null), n('', null, null), n(undefined, undefined, null), n('  ', '  ', null)], [none, none, none, none]);
  check('a text that is not a time: unreadable, not guessed', ['soon', '18/09/2026 07:26:14', '9/18/26, 7:26:14 AM', '2026-13-01 10:00:00', '2026-09-18 25:00:00', '46000.25', 'TRUE'].map((v) => n(v, v, null)), [bad, bad, bad, bad, bad, bad, bad]);
  check('a number is not read as a date', n('', 46000.25, zoneFormat('UTC')), bad);

  // The shown text is in another form: the Date is read back in the sheet's own zone.
  const date = new Date('2026-09-18T07:26:14Z');
  check('shown in another form, value a Date: read back in the zone of the sheet', [n('18/09/2026 09:26:14', date, zoneFormat('Europe/Berlin')), n('18.9.2026', date, zoneFormat('UTC')), n('', date, zoneFormat('Asia/Tokyo'))],
    [at('2026-09-18T09:26:14'), at('2026-09-18T07:26:14'), at('2026-09-18T16:26:14')]);
  check('a Date without a way to read it back, or an invalid Date: unreadable', [n('', date, null), n('', date, () => ''), n('', date, () => { throw new Error('x'); }), n('', new Date('nonsense'), zoneFormat('UTC'))], [bad, bad, bad, bad]);
  check('a Date from another script context is still a Date', n('', vm.runInNewContext('new Date("2026-09-18T07:26:14Z")'), zoneFormat('UTC')), at('2026-09-18T07:26:14'));
  check('a date without a time is kept as that, not given a time', [n('2026-09-18', null, null), n('', '2026-09-18', null)], [{ at: null, dateOnly: '2026-09-18', unreadable: false }, { at: null, dateOnly: '2026-09-18', unreadable: false }]);

  // Through the real list reader.
  const r = report();
  const all = r.all();
  check('the list is read once per call, the Uploaded column once more as shown text, for exactly its rows', [r.downloads, r.displayReads, r.lastRange], [1, 1, [1, UP + 1, TDOCS.length + 1, 1]]);
  check('every TDoc carries its upload time', all.map((td) => [td.row[0], td.uploaded.at || (td.uploaded.unreadable ? 'unreadable' : 'none')]), [
    ['S4aA269001', '2026-09-18T07:26:14'], ['S4aA269002', '2026-09-22T14:00:00'], ['S4aA269003', '2026-09-22T16:00:00'], ['S4aA269004', '2026-09-22T16:00:01'], ['S4aA269005', '2026-09-23T21:59:59'],
    ['S4aA269006', '2026-09-23T22:00:00'], ['S4aA269007', '2026-09-25T09:00:00'], ['S4aA269008', 'none'], ['S4aA269009', 'unreadable'], ['S4aA269010', '2026-09-23T08:00:00']]);
  check('it is carried as the list has it; nothing is converted while the list is read', all[1].uploaded.at, '2026-09-22T14:00:00');
  check('a reserved TDoc has none; a revision has its own, not the original\'s', [all[7].uploaded, all[9].uploaded.at, all[1].uploaded.at], [none, '2026-09-23T08:00:00', '2026-09-22T14:00:00']);
  check('nothing else of a TDoc changed: the same fields as before, plus "uploaded"', Object.keys(all[0]).sort(),
    ['agendaCol', 'agendaItem', 'agendaTopicCol', 'contactCol', 'forCol', 'revisedToCol', 'richTextRow', 'row', 'sourceCol', 'statusCol', 'tdocCol', 'titleCol', 'typeCol', 'uploaded']);

  // Where the upload time is not read.
  [['an ad-hoc report without sessions', ADHOC], ['a main-meeting report', Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137_Synthetic', MEETING_NUMBER: '137', AGENDA_ITEM_PREFIX: '4.' })],
    ['an ad-hoc report whose sessions cannot be read', Object.assign({}, ADHOC, { ADHOC_SESSIONS: '{"v":1,' })]].forEach(([what, props]) => {
    const x = report({ props: props });
    const tds = x.all();
    check(`${what}: the Uploaded column is not read and no TDoc carries an upload time`, [tds.length > 0, x.displayReads, tds.filter((td) => 'uploaded' in td).length], [true, 0, 0]);
  });

  // A sheet that cannot give display values; a sheet without a zone; a list without the column.
  const noDisplay = report({ display: false });
  check('no shown text available: the Dates are read back in the zone of the sheet', noDisplay.all().map((td) => td.uploaded.at).slice(0, 3), ['2026-09-18T07:26:14', '2026-09-22T14:00:00', '2026-09-22T16:00:00']);
  const berlin = report({ display: false, zone: 'Europe/Berlin' });
  check('the zone of the sheet decides that reading, not the zone of the script', berlin.all()[0].uploaded.at, '2026-09-18T09:26:14');
  const noZone = report({ display: false, zone: null });
  check('neither shown text nor a zone: uploaded TDocs are unreadable, reserved ones are still "not uploaded"', [noZone.all()[0].uploaded, noZone.all()[7].uploaded], [bad, none]);
  const noColumn = report({ header: HEADER.filter((h) => h !== 'Uploaded') });
  check('a list without an Uploaded column: no TDoc carries an upload time, and nothing fails', [noColumn.all().length, noColumn.all().filter((td) => 'uploaded' in td).length, noColumn.displayReads], [10, 0, 0]);
  const broken = report();
  broken.s.adhocSessionsEnabled_ = () => { throw new Error('unexpected'); };
  check('a failure while reading the upload times never fails the list: the TDocs are returned without them, and it is logged', [broken.all().length, broken.all().filter((td) => 'uploaded' in td).length, broken.logs.filter((l) => /upload times of the TDoc list could not be read: unexpected/.test(l)).length], [10, 0, 2]);
  check('the reservation date is never read as an upload time (source)', /Reservation date/.test(CODE.slice(CODE.indexOf('// AD-HOC TDOC SESSIONS (stage B)'), CODE.indexOf('// AD-HOC OPENING (stage E)'))), false);
}

// ================================================================ 2. automatic assignment

console.log('2. automatic assignment');
{
  const assign = (at, sessions) => S.assignAdhocSession_(at === undefined ? undefined : (typeof at === 'string' ? { at: at, dateOnly: null, unreadable: false } : at), sessions || SESSIONS);
  const id = (at, sessions) => assign(at, sessions).sessionId;

  check('cut-offs: the planned end, else the end of the day', SESSIONS.map(S.adhocSessionCutoff_), ['2026-09-22T18:00:00', '2026-09-23T23:59:59', '2026-09-24T23:59:59']);
  check('before the first session: the first', [id('2026-09-18T07:26:14'), id('2026-01-01T00:00:00'), id('2026-09-22T14:59:59')], ['s1', 's1', 's1']);
  check('during a session: that session', [id('2026-09-22T15:00:00'), id('2026-09-22T16:30:00'), id('2026-09-22T17:59:59')], ['s1', 's1', 's1']);
  check('exactly at the cut-off: that session', assign('2026-09-22T18:00:00'), { sessionId: 's1', reason: 'assigned' });
  check('one second after the cut-off: the next session', assign('2026-09-22T18:00:01'), { sessionId: 's2', reason: 'assigned' });
  check('between two sessions: the later one', [id('2026-09-22T21:00:00'), id('2026-09-23T03:00:00')], ['s2', 's2']);
  check('a session without an end runs to the end of its day', [id('2026-09-23T23:59:59'), id('2026-09-24T00:00:00'), id('2026-09-24T23:59:59')], ['s2', 's3', 's3']);
  check('after the last cut-off: no session, not the last one', [assign('2026-09-25T00:00:00'), assign('2027-01-01T00:00:00')], [{ sessionId: null, reason: 'after-last' }, { sessionId: null, reason: 'after-last' }]);
  check('not uploaded: no session', assign({ at: null, dateOnly: null, unreadable: false }), { sessionId: null, reason: 'not-uploaded' });
  check('an unreadable time: no session', assign({ at: null, dateOnly: null, unreadable: true }), { sessionId: null, reason: 'unreadable' });
  check('a list without upload times: no session', assign(undefined), { sessionId: null, reason: 'no-upload-time' });
  check('the sessions in any order give the same result', [id('2026-09-22T18:00:01', [S3, S1, S2]), id('2026-09-18T07:26:14', [S3, S2, S1]), id('2026-09-24T12:00:00', [S2, S3, S1])], ['s2', 's1', 's3']);
  check('the result never is a list: one session or none', TDOCS.map((t) => typeof id(t[3].replace(' ', 'T'))), TDOCS.map((t) => (/^\d{4}/.test(t[3]) && t[3] <= '2026-09-24 23:59:59' ? 'string' : 'object')));

  // Several sessions on one day.
  const morning = session('s1', '2026-09-22', '09:00', '12:00');
  const afternoon = session('s2', '2026-09-22', '15:00');
  check('two sessions on one day: before and during the first, between them, during the second, late that night',
    ['2026-09-22T08:00:00', '2026-09-22T12:00:00', '2026-09-22T12:00:01', '2026-09-22T16:00:00', '2026-09-22T23:59:59', '2026-09-23T00:00:00'].map((at) => id(at, [morning, afternoon])), ['s1', 's1', 's2', 's2', 's2', null]);

  // A date without a time.
  const dateOnly = (date, sessions) => S.assignAdhocSession_({ at: null, dateOnly: date, unreadable: false }, sessions || SESSIONS);
  check('a date without a time, on a day that lies in one session\'s window', [dateOnly('2026-09-20'), dateOnly('2026-09-23'), dateOnly('2026-09-24')],
    [{ sessionId: 's1', reason: 'assigned' }, { sessionId: 's2', reason: 'assigned' }, { sessionId: 's3', reason: 'assigned' }]);
  check('a date without a time, on the day of a cut-off within the day: no session, because the side of the cut-off is not known', dateOnly('2026-09-22'), { sessionId: null, reason: 'date-only' });
  check('a date without a time, after the last session', dateOnly('2026-09-25'), { sessionId: null, reason: 'after-last' });

  check('one session only', [id('2026-09-01T00:00:00', [S2]), id('2026-09-23T23:59:59', [S2]), id('2026-09-24T00:00:00', [S2])], ['s2', 's2', null]);
  check('the inputs are not changed', (() => { const list = [S3, S1, S2]; S.assignAdhocSession_({ at: '2026-09-22T18:00:01', dateOnly: null, unreadable: false }, list); return list.map((x) => x.id); })(), ['s3', 's1', 's2']);
  check('the assignment uses no clock, no property and no document (source)', (() => {
    const at = CODE.indexOf('\nfunction assignAdhocSession_(');
    return /new Date|Date\.now|PropertiesService|DocumentApp|getReportStateStore_|Utilities/.test(CODE.slice(at, CODE.indexOf('\n}\n', at)));
  })(), false);
}

// ================================================================ 3. manual assignments: the model

console.log('3. manual assignments: the model');
{
  const eff = (auto, override, sessions) => S.effectiveAdhocTdocSessions_(auto, override, sessions || SESSIONS);
  const text = (ids, sessions) => S.adhocTdocSessionCellText_(ids, sessions || SESSIONS);
  check('no manual assignment: the automatic session', [eff('s1'), eff('s3'), eff(null), eff('')], [['s1'], ['s3'], [], []]);
  check('"add": the automatic session and the named one', eff('s1', { mode: 'add', sessions: ['s3'] }), ['s1', 's3']);
  check('"add" several; "add" the automatic one again; "add" without an automatic session', [eff('s1', { mode: 'add', sessions: ['s3', 's2'] }), eff('s1', { mode: 'add', sessions: ['s1'] }), eff(null, { mode: 'add', sessions: ['s2'] })], [['s1', 's2', 's3'], ['s1'], ['s2']]);
  check('"add" never removes the automatic session', [eff('s2', { mode: 'add', sessions: ['s1'] }), eff('s2', { mode: 'add', sessions: ['s3'] })], [['s1', 's2'], ['s2', 's3']]);
  check('"set": the named sessions and no others', [eff('s1', { mode: 'set', sessions: ['s3'] }), eff('s1', { mode: 'set', sessions: ['s3', 's2'] }), eff(null, { mode: 'set', sessions: ['s1'] })], [['s3'], ['s2', 's3'], ['s1']]);
  check('"set" to nothing: explicitly no session', eff('s1', { mode: 'set', sessions: [] }), []);
  check('the result is in the order of the sessions, whatever the stored order', eff('s3', { mode: 'add', sessions: ['s2', 's1'] }), ['s1', 's2', 's3']);
  check('a session that is not configured is left out', [eff('s9', undefined), eff('s1', { mode: 'add', sessions: ['s9'] }), eff('s1', { mode: 'set', sessions: ['s9'] })], [[], ['s1'], []]);

  check('the cell: one label, several labels, none', [text(['s1']), text(['s1', 's3']), text(['s3', 's1']), text([])], ['A01', 'A01, A03', 'A01, A03', DASH]);
  const renamed = [session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), S2, S3];
  check('a renamed session: the same ids give the new label', [text(['s1', 's3'], renamed), eff('s1', { mode: 'add', sessions: ['s3'] }, renamed)], ['A01, A03', ['s1', 's3']]);
  const redated = [session('s1', '2026-09-30'), S2, S3];
  check('a re-dated session: the same ids, the labels in the new order', text(['s1', 's3'], redated), 'A02, A03');

  // Normalizing what a dialog sends.
  const norm = S.normalizeAdhocTdocOverrides_;
  check('accepted', norm({ S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269002: { mode: 'set', sessions: [] }, 'S4-269003': { mode: 'set', sessions: ['s10', 's2', 's2'] } }),
    { errors: [], tdocs: { S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269002: { mode: 'set', sessions: [] }, 'S4-269003': { mode: 'set', sessions: ['s2', 's10'] } } });
  check('an "add" without sessions is the same as automatic and is left out', norm({ S4aA269001: { mode: 'add', sessions: [] } }), { errors: [], tdocs: {} });
  check('nothing', norm({}), { errors: [], tdocs: {} });
  check('keyed by the TDoc number only: a title, a row number or a label is refused', [norm({ 'Synthetic title': { mode: 'set', sessions: ['s1'] } }).errors, norm({ 3: { mode: 'set', sessions: ['s1'] } }).errors, norm({ S4aA269001: { mode: 'set', sessions: ['22 Sep'] } }).errors, norm({ S4aA269001: { mode: 'set', sessions: ['2026-09-22'] } }).errors],
    [['"Synthetic title" is not a TDoc number.'], ['"3" is not a TDoc number.'], ['S4aA269001: the manual assignment cannot be read.'], ['S4aA269001: the manual assignment cannot be read.']]);
  check('an unknown mode, a missing list, not an object', [norm({ S4aA269001: { mode: 'replace', sessions: [] } }).errors.length, norm({ S4aA269001: { mode: 'add' } }).errors.length, norm({ S4aA269001: 's1' }).errors.length, norm(null).errors, norm([]).errors],
    [1, 1, 1, ['The manual assignments are missing.'], ['The manual assignments are missing.']]);
  check('with an error nothing is returned', norm({ S4aA269001: { mode: 'add', sessions: ['s3'] }, bad: { mode: 'add', sessions: ['s3'] } }).tdocs, {});
}

// ================================================================ 4. manual assignments: the stored property

console.log('4. manual assignments: the stored property');
{
  const parse = S.parseAdhocTdocSessionsProperty_;
  const ser = S.serializeAdhocTdocSessions_;
  const automatic = { clock: 'utc', tdocs: {} };
  const model = { clock: 'utc', tdocs: { S4aA269003: { mode: 'add', sessions: ['s3'] }, S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269002: { mode: 'set', sessions: [] }, S4aA269004: { mode: 'set', sessions: ['s2', 's1'] } } };
  const stored = ser(model);
  check('the stored form: version, clock, TDocs grouped by their sessions, in a fixed order', stored, '{"v":1,"clock":"utc","add":{"s3":["S4aA269001","S4aA269003"]},"set":{"":["S4aA269002"],"s1+s2":["S4aA269004"]}}');
  check('it holds session ids and TDoc numbers: no label, no date, no title, no row', /22 Sep|2026|Synthetic|title|row/i.test(stored), false);
  check('it is read back as the same model', parse(stored), { status: 'ok', model: { clock: 'utc', tdocs: { S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269003: { mode: 'add', sessions: ['s3'] }, S4aA269002: { mode: 'set', sessions: [] }, S4aA269004: { mode: 'set', sessions: ['s1', 's2'] } } }, error: null });
  check('and written back identically', ser(parse(stored).model), stored);
  check('nothing stored: automatic, on the Portal\'s clock (UTC)', [parse(null), parse(''), parse('  ')].map((x) => [x.status, x.model]), [['absent', automatic], ['absent', automatic], ['absent', automatic]]);
  check('nothing manual', ser(automatic), '{"v":1,"clock":"utc","add":{},"set":{}}');

  ['{"v":1,', '[]', '"x"', '{"v":1}', '{"v":1,"clock":"cest","add":{},"set":{}}', '{"v":1,"clock":"utc","add":[],"set":{}}', '{"v":1,"clock":"utc","add":{"22 Sep":["S4aA269001"]},"set":{}}',
    '{"v":1,"clock":"utc","add":{"s1":"S4aA269001"},"set":{}}', '{"v":1,"clock":"utc","add":{"s1":["not a tdoc"]},"set":{}}', '{"v":1,"clock":"utc","add":{"s1":["S4aA269001"]},"set":{"s2":["S4aA269001"]}}',
    '{"v":1,"clock":"utc","add":{"":["S4aA269001"]},"set":{}}', '{"v":1,"clock":"utc","add":{"s1":[5]},"set":{}}'].forEach((raw) => {
    check(`unreadable (${raw.slice(0, 44)}): invalid, the automatic model, no error thrown`, [parse(raw).status, parse(raw).model, typeof parse(raw).error], ['invalid', automatic, 'string']);
  });
  check('a newer schema version: unsupported, the automatic model', [parse('{"v":2,"clock":"utc","add":{},"set":{}}').status, parse('{"v":2}').model, parse('{"clock":"utc"}').status], ['unsupported', automatic, 'unsupported']);
  check('it never throws', thrown(() => [{}, [], 5, true, '{', '\u0000'].forEach((raw) => parse(raw))), null);

  // Saving.
  const save = (props, input, ids) => { const l = loadCode({ documentProperties: props || {} }); const writes = []; const set = l.docProps.setProperty; l.docProps.setProperty = (k, v) => { writes.push(k); set(k, v); };
    return [l.sandbox.saveAdhocTdocSessionsWith_(l.docProps, ids || ['s1', 's2', 's3'], input), writes, l.docProps.getProperty(KEY)]; };
  const INPUT = { clock: 'utc', tdocs: { S4aA269001: { mode: 'add', sessions: ['s3'] } } };
  const STORED = '{"v":1,"clock":"utc","add":{"s3":["S4aA269001"]},"set":{}}';
  check('a save: one property written', save({}, INPUT), [{ ok: true, errors: [], changed: true }, [KEY], STORED]);
  check('saving the same again writes nothing', save({ [KEY]: STORED }, INPUT), [{ ok: true, errors: [], changed: false }, [], STORED]);
  check('saving "nothing manual, the Portal\'s clock" where nothing is stored creates no property', save({}, { clock: 'utc', tdocs: {} }), [{ ok: true, errors: [], changed: false }, [], null]);
  check('the other clock alone is something to store', save({}, { clock: 'session', tdocs: {} }), [{ ok: true, errors: [], changed: true }, [KEY], '{"v":1,"clock":"session","add":{},"set":{}}']);
  const SAME_CLOCK = '{"v":1,"clock":"session","add":{"s3":["S4aA269001"]},"set":{}}';
  check('a report that chose the session clock: read back as that, never as the default', [parse(SAME_CLOCK).status, parse(SAME_CLOCK).model.clock, parse('{"v":1,"clock":"session","add":{},"set":{}}').model.clock], ['ok', 'session', 'session']);
  check('saving it again as it is writes nothing: the choice is not rewritten to the default', save({ [KEY]: SAME_CLOCK }, { clock: 'session', tdocs: { S4aA269001: { mode: 'add', sessions: ['s3'] } } }), [{ ok: true, errors: [], changed: false }, [], SAME_CLOCK]);
  check('a report that chose UTC explicitly stays UTC', [parse('{"v":1,"clock":"utc","add":{},"set":{}}').model.clock, save({ [KEY]: '{"v":1,"clock":"utc","add":{},"set":{}}' }, { clock: 'utc', tdocs: {} })], ['utc', [{ ok: true, errors: [], changed: false }, [], '{"v":1,"clock":"utc","add":{},"set":{}}']]);
  check('going back to automatic keeps the property, empty (the clock is still a choice)', save({ [KEY]: STORED }, { clock: 'utc', tdocs: {} }), [{ ok: true, errors: [], changed: true }, [KEY], '{"v":1,"clock":"utc","add":{},"set":{}}']);
  check('a session that is not configured: refused, nothing written', save({ [KEY]: STORED }, { clock: 'utc', tdocs: { S4aA269001: { mode: 'add', sessions: ['s7'] } } }),
    [{ ok: false, errors: ['S4aA269001 is assigned to a session that is not configured. Close the dialog and open it again.'], changed: false }, [], STORED]);
  check('no clock, an unknown clock: refused (the dialog always sends one)', [save({}, { tdocs: {} })[0].ok, save({}, { clock: 'cest', tdocs: {} })[0].errors, save({}, null)[0].ok], [false, ['The clock of the upload times is missing. Close the dialog and open it again.'], false]);
  check('an invalid entry: refused with its message, nothing written', save({ [KEY]: STORED }, { clock: 'utc', tdocs: { 'Synthetic title': { mode: 'set', sessions: ['s1'] } } }), [{ ok: false, errors: ['"Synthetic title" is not a TDoc number.'], changed: false }, [], STORED]);
  check('an unreadable stored value is replaced by a valid save', save({ [KEY]: '{"v":1,' }, INPUT), [{ ok: true, errors: [], changed: true }, [KEY], STORED]);
  check('a stored value of a newer version is never replaced', save({ [KEY]: '{"v":2}' }, INPUT),
    [{ ok: false, errors: ['The stored manual TDoc session assignments have version 2; this release reads version 1. They were written by a newer release and are not changed here.'], changed: false }, [], '{"v":2}']);
  const failing = loadCode({ documentProperties: { [KEY]: STORED } });
  failing.docProps.setProperty = () => { throw new Error('quota'); };
  check('a write that fails: the error is raised and the stored value is the one before', [thrown(() => failing.sandbox.saveAdhocTdocSessionsWith_(failing.docProps, ['s1', 's2', 's3'], { clock: 'utc', tdocs: {} })), failing.docProps.getProperty(KEY)], ['quota', STORED]);

  // Size.
  const many = (n, sessionsOf) => { const tdocs = {}; for (let i = 0; i < n; i++) tdocs['S4aA26' + String(1000 + i)] = sessionsOf(i); return tdocs; };
  const three = (i) => ({ mode: i % 2 ? 'add' : 'set', sessions: [['s1'], ['s2'], ['s1', 's3'], ['s2', 's3']][i % 4] });
  const big = ser({ clock: 'utc', tdocs: many(300, three) });
  check('300 manually assigned TDocs fit in one property with room to spare', [big.length < 4500, Buffer.byteLength(big, 'utf8') < 9000, Object.keys(parse(big).model.tdocs).length], [true, true, 300]);
  check('about 13 characters per TDoc, whatever the number of sessions', Math.round(big.length / 300), 13);
  const tooMany = save({}, { clock: 'utc', tdocs: many(700, three) });
  check('more than fits is refused, not cut off; nothing is written', [ser({ clock: 'utc', tdocs: many(700, three) }).length > 8000, tooMany], [true, [{ ok: false, errors: ['There are too many manual assignments to store.'], changed: false }, [], null]]);
  check('the limit is below the 9 KB of one property', (CODE.match(/const ADHOC_TDOC_SESSIONS_MAX_CHARS_ = (\d+);/) || [])[1], '8000');

  // What the session editor is told.
  const refs = S.adhocTdocSessionReferences_;
  check('which sessions are named, and by which TDocs', refs(stored), { s3: ['S4aA269001', 'S4aA269003'], s1: ['S4aA269004'], s2: ['S4aA269004'] });
  check('nothing stored, or unreadable: none; a newer version: unknown', [refs(null), refs('{"v":1,'), refs('{"v":2}')], [{}, {}, null]);
}

// ================================================================ 5. the clock

console.log('5. the clock of the upload times');
{
  const r = report();
  const resolver = () => r.s.makeAdhocTdocSessionResolver_();
  const autos = () => { const res = resolver(); const out = {}; r.all().forEach((td) => { out[td.row[0]] = res.automatic(td).sessionId || ''; }); return out; };
  check('by default the upload times are UTC and are converted to the report\'s zone before they are compared', [resolver().clock, r.docProps.getProperty(KEY), autos()], ['utc', null, EXPECTED_AUTO]);
  check('exactly at the cut-off (16:00:00 UTC is 18:00:00): that session; one second later: the next', [autos().S4aA269003, autos().S4aA269004], ['s1', 's2']);
  check('across midnight: 21:59:59 UTC is still the 23rd, 22:00:00 UTC is the 24th', [autos().S4aA269005, autos().S4aA269006], ['s2', 's3']);
  check('the reasons for "no session"', (() => { const res = resolver(); return r.all().filter((td) => !res.automatic(td).sessionId).map((td) => [td.row[0], res.automatic(td).reason]); })(),
    [['S4aA269007', 'after-last'], ['S4aA269008', 'not-uploaded'], ['S4aA269009', 'unreadable']]);

  check('conversion from UTC to the report\'s zone: summer and winter by the zone\'s own rules, across midnight',
    ['2026-09-22T16:00:00', '2026-12-01T16:00:00', '2026-09-22T22:30:00', '2026-10-25T00:30:00', '2026-10-25T01:30:00'].map(r.s.convertUtcWallClockToReportZone_),
    ['2026-09-22T18:00:00', '2026-12-01T17:00:00', '2026-09-23T00:30:00', '2026-10-25T02:30:00', '2026-10-25T02:30:00']);
  check('nothing to convert', [r.s.convertUtcWallClockToReportZone_(''), r.s.convertUtcWallClockToReportZone_('2026-09-22'), r.s.convertUtcWallClockToReportZone_(null)], ['', '', '']);
  check('no offset and no zone name is written in the code: the zone comes from the script, the rules from Apps Script (source)', (() => {
    const section = CODE.slice(CODE.indexOf('// AD-HOC TDOC SESSIONS (stage B)'), CODE.indexOf('// AD-HOC OPENING (stage E)')).replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '');
    return [/CEST|CET\b|Europe\/|\+0[12]:?00|7200|3600/.test(section), /Utilities\.formatDate\(new Date\(Date\.UTC\([^;]*Session\.getScriptTimeZone\(\)/.test(section)];
  })(), [false, true]);

  check('a report can say that its list is already on the sessions\' clock; that is stored', [r.s.saveAdhocTdocSessions({ clock: 'session', tdocs: {} }).ok, resolver().clock, JSON.parse(r.docProps.getProperty(KEY)).clock], [true, 'session', 'session']);
  // Then nothing is converted. 9004: 16:00:01 is before the cut-off of 18:00: session 1. 9006: 22:00 on the 23rd: session 2.
  check('then the times are compared as they stand', autos(), Object.assign({}, EXPECTED_AUTO, { S4aA269004: 's1', S4aA269006: 's2' }));
  check('a stored choice stays in force: it is not reinterpreted as the default', [loadCode({ documentProperties: { [KEY]: r.docProps.getProperty(KEY) } }).sandbox.parseAdhocTdocSessionsProperty_(r.docProps.getProperty(KEY)).model.clock, resolver().clock], ['session', 'session']);
  check('back to the Portal\'s clock', [r.save('utc', {}).ok, resolver().clock, autos()], [true, 'utc', EXPECTED_AUTO]);

  // Winter: the same conversion gives one hour, by the zone's rules.
  const winter = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-12-01', '15:00', '18:00'), session('s2', '2026-12-02')]) }),
    tdocs: [['S4aA269001', '4.3', 'available', '2026-12-01 16:59:59', '', ''], ['S4aA269002', '4.3', 'available', '2026-12-01 17:00:00', '', ''], ['S4aA269003', '4.3', 'available', '2026-12-01 17:00:01', '', ''],
      ['S4aA269004', '4.4', 'available', '2026-12-02 22:59:59', '', ''], ['S4aA269005', '4.4', 'available', '2026-12-02 23:00:00', '', '']] });
  const winterResolver = winter.s.makeAdhocTdocSessionResolver_();
  check('winter (UTC+1): 17:00:00 UTC is 18:00:00, the cut-off; a second later is the next session; 23:00 UTC is the next day',
    winter.all().map((td) => winterResolver.automatic(td).sessionId), ['s1', 's1', 's2', 's2', null]);
  check('summer (UTC+2), the same clock times: 17:00:00 UTC is 19:00:00, after the cut-off', (() => {
    const summer = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-07-01', '15:00', '18:00'), session('s2', '2026-07-02')]) }),
      tdocs: [['S4aA269001', '4.3', 'available', '2026-07-01 15:59:59', '', ''], ['S4aA269002', '4.3', 'available', '2026-07-01 16:00:00', '', ''], ['S4aA269003', '4.3', 'available', '2026-07-01 16:00:01', '', ''], ['S4aA269004', '4.3', 'available', '2026-07-01 17:00:00', '', '']] });
    const res = summer.s.makeAdhocTdocSessionResolver_();
    return summer.all().map((td) => res.automatic(td).sessionId);
  })(), ['s1', 's1', 's2', 's2']);
  check('the zone is the script\'s: in another zone the same list gives other sessions', (() => {
    const tokyo = report();
    tokyo.s.Session = { getScriptTimeZone: () => 'Asia/Tokyo' };
    const res = tokyo.s.makeAdhocTdocSessionResolver_();
    return [res.automatic(tokyo.all()[1]).sessionId, res.automatic(tokyo.all()[4]).sessionId];   // 14:00 UTC is 23:00 in Tokyo (after 18:00); 21:59:59 UTC on the 23rd is 06:59:59 on the 24th
  })(), ['s2', 's3']);
  const broken = report();
  broken.save('utc', {});
  broken.s.Utilities.formatDate = () => 'not a time';
  check('a conversion that does not give a time: no session for that TDoc, never an unconverted time', (() => { const res = broken.s.makeAdhocTdocSessionResolver_(); return res.automatic(broken.all()[0]); })(), { sessionId: null, reason: 'unreadable' });

  // The resolver.
  check('no resolver: main meeting, ad-hoc without sessions, sessions that cannot be read',
    [report({ props: Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }) }).s.makeAdhocTdocSessionResolver_(), report({ props: ADHOC }).s.makeAdhocTdocSessionResolver_(), report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: 'x' }) }).s.makeAdhocTdocSessionResolver_()], [null, null, null]);
  const corrupt = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: '{"v":1,' }) });
  const res = corrupt.s.makeAdhocTdocSessionResolverSafely_();
  check('manual assignments that cannot be read are ignored: the automatic sessions are used, and one line is logged', [res.status, res.clock, corrupt.all().map((td) => res.cell(td)).slice(0, 4), corrupt.logs.filter((l) => /^TDoc sessions/.test(l))],
    ['invalid', 'utc', ['A01', 'A01', 'A01', 'A02'], ['TDoc sessions: The stored manual TDoc session assignments cannot be read. Only the automatic sessions are used.']]);
  const newer = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: '{"v":2}' }) });
  check('the same for a newer version, which is left as it is', [newer.s.makeAdhocTdocSessionResolverSafely_().status, newer.s.makeAdhocTdocSessionResolverSafely_().cell(newer.all()[0]), newer.docProps.getProperty(KEY)], ['unsupported', 'A01', '{"v":2}']);
  const failing = report();
  failing.s.getAdhocSessions_ = () => { throw new Error('boom'); };
  check('a resolver that cannot be made: null and a log line, never an error', [failing.s.makeAdhocTdocSessionResolverSafely_(), failing.logs], [null, ['TDoc sessions: not available for this run: boom']]);
}

// ================================================================ 6. a new build

console.log('6. a new build: four or five columns');
{
  const r = report();
  const built = r.build();
  check('built', [built.ok, built.note], [true, '']);
  check('ad-hoc with sessions: the registration table has five columns, the fifth headed Session', r.registration()[0], ['TDoc', 'Title', 'Source', 'Agenda Item', 'Session']);
  check('one row per TDoc of the list, each with five cells', [r.registration().length - 1, r.registration().every((row) => row.length === 5)], [10, true]);
  check('the Session cells: the code of the automatic session, a dash for none', r.sessionCells(), autoLabels(EXPECTED_AUTO));
  check('the other four cells are what they always were', r.registration()[1].slice(0, 4), ['S4aA269001', 'Synthetic title of S4aA269001', 'ExampleCorp', '4.3']);
  check('no upload time is written anywhere in the report', /2026-09-18|07:26:14|18:00:01|Uploaded/.test(r.snapshot()), false);
  check('the TDoc detail tables are unchanged: no Session row', r.body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getCell(0).getText() === 'TDoc' && c.getRow(0).getNumCells() === 2).map((t) => rowsOf(t).map((row) => row[0]).join('|')).filter((x, i, l) => l.indexOf(x) === i),
    ['TDoc|Title|Source|Contact|Agenda Item|Type/For|E-mail Discussion|Revisions|Minutes|Disposition|Status']);

  // Manual assignments in a build.
  const m = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: S.serializeAdhocTdocSessions_({ clock: 'utc', tdocs: {
    S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269004: { mode: 'set', sessions: ['s1'] }, S4aA269005: { mode: 'set', sessions: [] }, S4aA269008: { mode: 'add', sessions: ['s2', 's3'] } } }) }) });
  m.build();
  check('several labels, a replacement, an explicit "no session", sessions for a TDoc that is not uploaded',
    [m.sessionCells().S4aA269001, m.sessionCells().S4aA269004, m.sessionCells().S4aA269005, m.sessionCells().S4aA269008, m.sessionCells().S4aA269002], ['A01, A03', 'A01', DASH, 'A02, A03', 'A01']);
  const labelled = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), S2, S3]) }) });
  labelled.build();
  check('the labels are the sessions\' current labels', [labelled.sessionCells().S4aA269001, labelled.sessionCells().S4aA269004], ['A01', 'A02']);

  // Four columns everywhere else.
  const noSessions = report({ props: ADHOC });
  noSessions.build();
  check('ad-hoc without sessions: four columns, as before', [noSessions.registration()[0], noSessions.registration().every((row) => row.length === 4), noSessions.displayReads], [['TDoc', 'Title', 'Source', 'Agenda Item'], true, 0]);
  const strip = (snapshot) => JSON.stringify(JSON.parse(snapshot).map((e) => (e[0] === 'T' && e[1][0].length === 5 && e[1][0][4] === 'Session' ? ['T', e[1].map((row) => row.slice(0, 4))] : e)));
  // Stage E: with sessions the opening part also has the generated Online information block and Session administration section.
  const withoutOpening = (snapshot) => { const out = []; let inside = false; JSON.parse(snapshot).forEach((e) => { if (e[0] === 'P' && (e[1] === 'Session administration' || e[1] === 'Online information') && e[2] !== 'NORMAL') { inside = true; return; } if (inside && e[0] === 'P' && e[2] === 'NORMAL') return; inside = false; out.push(e); }); return JSON.stringify(out); };
  // E1: and the "<Chair> opens the session ..." line of the build is not written beside that section.
  const withoutLegacyLine = (snapshot) => JSON.stringify(JSON.parse(snapshot).filter((e) => !(e[0] === 'P' && /^<Chair> opens the session on /.test(e[1]))));
  check('apart from that column (and the Online information and Session administration sections in place of the "<Chair> opens ..." line), the report built with sessions is the report built without',
    [r.snapshot() === noSessions.snapshot(), withoutOpening(strip(r.snapshot())) === withoutLegacyLine(noSessions.snapshot()), withoutLegacyLine(noSessions.snapshot()) === noSessions.snapshot()], [false, true, false]);
  const stray = report({ props: Object.assign({}, ADHOC, { [KEY]: '{"v":1,"clock":"utc","add":{"s1":["S4aA269001"]},"set":{}}' }) });
  stray.build();
  check('ad-hoc without sessions, with assignments left in its properties: four columns, the same report', stray.snapshot() === noSessions.snapshot(), true);
  const badSessions = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: '{"v":1,' }) });
  badSessions.build();
  check('ad-hoc with sessions that cannot be read: four columns, the same report', badSessions.snapshot() === noSessions.snapshot(), true);

  const cs = S.createSummaryTable_;
  const sandboxBody = () => makeFakeDocumentBody(S);
  const td = { row: ['S4-269001', 'T', 'Src'], tdocCol: 0, agendaItem: '7.1' };
  check('createSummaryTable_(): without a resolver four columns (every main-meeting caller); with one, five',
    [rowsOf(cs(sandboxBody(), [td], 0, 1, 2, -1)), rowsOf(cs(sandboxBody(), [td], 0, 1, 2, -1, null)), rowsOf(cs(sandboxBody(), [td], 0, 1, 2, -1, { cell: () => 'X' }))],
    [[['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4-269001', 'T', 'Src', '7.1']], [['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4-269001', 'T', 'Src', '7.1']], [['TDoc', 'Title', 'Source', 'Agenda Item', 'Session'], ['S4-269001', 'T', 'Src', '7.1', 'X']]]);
  check('the main-meeting 6G branch of the build passes no resolver (source)', (() => {
    const src = CODE.slice(CODE.indexOf('\nfunction buildSkeletonWithTdocTables('), CODE.indexOf('\nfunction downloadAndGroupTdocs_('));
    const calls = src.match(/createSummaryTable_\([^;]*\);/g) || [];
    return [calls.length, calls.filter((c) => /tdocSessions/.test(c)).length, src.indexOf(calls.filter((c) => /tdocSessions/.test(c))[0]) > src.indexOf('const emitOpeningAdminBlock')];
  })(), [2, 1, true]);
}

// ================================================================ 7. an existing report

console.log('7. an existing report: updates, and the explicit action');
{
  // A report built before sessions were configured: four columns.
  const r = report({ props: ADHOC });
  r.build();
  r.docProps.setProperty('ADHOC_SESSIONS', sessionsProperty(SESSIONS));
  const before = r.snapshot();
  const updated = r.update();
  check('an update of a four-column report with sessions configured: it completes', updated, { success: true, error: null });
  check('and the report is untouched: no column was added', [r.snapshot() === before, r.registration()[0].length], [true, 4]);
  r.tdocs = TDOCS.concat([['S4aA269011', '4.4', 'available', '2026-09-23 12:00:00', '', '']]);
  r.update();
  check('a new TDoc is appended with four cells; the table stays four-column', [r.registration().length - 1, r.registration()[11], r.registration().every((row) => row.length === 4), r.events.filter((e) => /^new/.test(e))],
    [11, ['S4aA269011', 'Synthetic title of S4aA269011', 'ExampleCorp', '4.4'], true, ['new S4aA269011']]);
  check('updating again adds nothing', (() => { const x = r.snapshot(); r.update(); return r.snapshot() === x; })(), true);
  check('saving manual assignments does not add the column either', [r.save('utc', { S4aA269001: { mode: 'add', sessions: ['s3'] } }), r.registration()[0].length], [{ ok: true, errors: [], changed: true }, 4]);
  check('nor does saving the sessions', [r.s.saveAdhocSessionsConfiguration(SESSIONS.map((x) => Object.assign({}, x, { label: x.id === 's1' ? 'Kick-off' : '' }))).notice, r.registration()[0].length], [undefined, 4]);
  r.s.saveAdhocSessionsConfiguration(SESSIONS.map((x) => Object.assign({}, x)));

  // The explicit action.
  const lockBefore = r.lock.tries;
  const beforeEnable = r.snapshot();
  const enabled = r.s.enableAdhocSessionColumn();
  check('the explicit action adds the column', enabled, { ok: true, error: null, changed: true, message: 'The Session column was added to the registration table (11 TDocs).' });
  check('header and one cell per row, with the sessions in force (the manual one included)', [r.registration()[0], r.registration().every((row) => row.length === 5), r.sessionCells()],
    [['TDoc', 'Title', 'Source', 'Agenda Item', 'Session'], true, Object.assign(autoLabels(EXPECTED_AUTO), { S4aA269001: 'A01, A03', S4aA269011: 'A02' })]);
  check('nothing else in the report changed', JSON.stringify(JSON.parse(r.snapshot()).map((e) => (e[0] === 'T' && e[1][0][4] === 'Session' ? ['T', e[1].map((row) => row.slice(0, 4))] : e))) === beforeEnable, true);
  check('it held the document lock and released it', [r.lock.tries - lockBefore, r.lock.releases === r.lock.tries], [1, true]);
  const again = r.snapshot();
  check('doing it again changes nothing', [r.s.enableAdhocSessionColumn(), r.snapshot() === again], [{ ok: true, error: null, changed: false, message: 'This report already has the Session column.' }, true]);

  // From now on updates keep the column.
  r.tdocs = r.tdocs.concat([['S4aA269012', '4.3', 'available', '2026-09-24 08:00:00', '', ''], ['S4aA269013', '4.3', 'reserved', '', '', '']]);
  r.events.length = 0;
  check('an update of the five-column report completes', r.update(), { success: true, error: null });
  check('new TDocs are appended with five cells, with their session', [r.registration().length - 1, r.registration().slice(-2), r.registration().every((row) => row.length === 5)],
    [13, [['S4aA269012', 'Synthetic title of S4aA269012', 'ExampleCorp', '4.3', 'A03'], ['S4aA269013', 'Synthetic title of S4aA269013', 'ExampleCorp', '4.3', DASH]], true]);
  check('no TDoc is in the table twice', r.registration().map((row) => row[0]).filter((id, i, l) => l.indexOf(id) !== i), []);
  r.tdocs = r.tdocs.map((t) => (t[0] === 'S4aA269013' ? ['S4aA269013', '4.3', 'available', '2026-09-22 15:00:00', '', ''] : t));
  r.update();
  check('a reserved TDoc that is uploaded gets its session at the next update, without a new TDoc being added', [r.sessionCells().S4aA269013, r.registration().length - 1], ['A01', 13]);
  const steady = r.snapshot();
  r.update();
  r.update();
  check('further updates without a change leave the report as it is', r.snapshot() === steady, true);
  check('an update never writes the manual assignments or the sessions', r.writes.filter((k) => /ADHOC_/.test(k)).length, (() => { const n = r.writes.filter((k) => /ADHOC_/.test(k)).length; r.update(); return r.writes.filter((k) => /ADHOC_/.test(k)).length === n ? n : -1; })());

  // Session cells follow the manual assignments and the sessions.
  check('saving another manual assignment shows at once in a five-column report', [r.save('utc', { S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269004: { mode: 'set', sessions: ['s1', 's3'] } }).ok, r.sessionCells().S4aA269004, r.sessionCells().S4aA269001], [true, 'A01, A03', 'A01, A03']);
  check('going back to automatic', [r.save('utc', { S4aA269001: { mode: 'add', sessions: ['s3'] } }).ok, r.sessionCells().S4aA269004], [true, 'A02']);
  const cellTyped = r.s.findRegistrationTable_(r.body).table.getRow(2).getCell(4);
  cellTyped.setText('typed by hand');
  r.update();
  check('a Session cell is generated: text typed into it is replaced by the next update', cellTyped.getText(), 'A01');
  const gone = r.s.findRegistrationTable_(r.body).table;
  gone.getRow(1).getCell(0).setText('S4aA269999');
  gone.getRow(1).getCell(4).setText('kept');
  r.update();
  check('a row whose TDoc is not in the list is left alone', gone.getRow(1).getCell(4).getText(), 'kept');
  gone.getRow(1).getCell(0).setText('S4aA269001');

  // The Session column cannot be updated.
  const f = report();
  f.build();
  f.s.refreshRegistrationSessionColumn_ = () => { throw new Error('table is locked'); };
  check('a Session column that cannot be updated does not fail the update: it is reported with the other parts', [f.update(), f.logs.filter((l) => /Session column could not be updated/.test(l)).length],
    [{ success: true, error: null, collectorFailures: [{ step: 'Session column', error: 'table is locked' }] }, 1]);
  const g = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: '{"v":1,' }) });
  g.build();
  check('manual assignments that cannot be read do not fail an update: it completes with the automatic sessions', [g.update(), g.sessionCells().S4aA269004, g.docProps.getProperty(KEY)], [{ success: true, error: null }, 'A02', '{"v":1,']);

  // Sessions removed from a report that has the column.
  const h = report();
  h.build();
  h.docProps.setProperty('ADHOC_SESSIONS', sessionsProperty([], 4));
  h.tdocs = TDOCS.concat([['S4aA269011', '4.4', 'available', '2026-09-23 12:00:00', '', '']]);
  h.update();
  check('a five-column report whose sessions were all removed still gets its new TDocs, with an empty Session cell', [h.registration().length - 1, h.registration()[11], h.registration()[1][4]], [11, ['S4aA269011', 'Synthetic title of S4aA269011', 'ExampleCorp', '4.4', ''], 'A01']);

  // Which tables are registration tables.
  const cols = (rows) => { const b = makeFakeDocumentBody(S); return S.registrationTableColumns_(b.appendTable(rows)); };
  check('recognized: the four-column table of every release, and the same with Session', [cols([['TDoc', 'Title', 'Source', 'Agenda Item']]), cols([['TDoc', 'Title', 'Source', 'Agenda Item', 'Session']]), cols([[' TDoc ', 'Title', 'Source', 'Agenda Item ', ' Session']])], [4, 5, 5]);
  check('not recognized: any other table', [cols([['TDoc', 'Title', 'Source', 'Agenda Item', 'Status']]), cols([['TDoc', 'Title', 'Source', 'Agenda Item', '']]), cols([['TDoc', 'Title', 'Source', 'Agenda Item', 'Session', 'More']]),
    cols([['TDoc', 'Title', 'Source']]), cols([['Name', 'Company', 'Email', 'Phone', 'Session']]), cols([['TDoc', 'Original Agenda', 'New Agenda', 'Reason']]), cols([['TDoc', 'S4aA269001']]), cols([['Session', 'Title', 'Source', 'Agenda Item', 'TDoc']]),
    cols([['tdoc', 'title', 'source', 'agenda item']])], [0, 0, 0, 0, 0, 0, 0, 0, 0]);
  const u = report();
  u.build();
  const other = u.body.insertTable(0, [['TDoc', 'Title', 'Source', 'Agenda Item', 'Status'], ['S4aA269001', 'x', 'y', '4.3', 'unrelated']]);
  u.tdocs = TDOCS.concat([['S4aA269011', '4.4', 'available', '2026-09-23 12:00:00', '', '']]);
  u.update();
  check('an unrelated five-column table earlier in the document is not taken for the registration table', [rowsOf(other), u.registration().length - 1, u.registration()[11][4]],
    [[['TDoc', 'Title', 'Source', 'Agenda Item', 'Status'], ['S4aA269001', 'x', 'y', '4.3', 'unrelated']], 11, 'A02']);

  // Where the explicit action does not apply.
  const empty = report();
  check('a report without a registration table', empty.s.enableAdhocSessionColumn(), { ok: false, error: 'This report has no registration table yet. Build the report first; a report built with sessions gets the Session column.' });
  [['a main-meeting report', Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }), 'TDoc sessions are available for ad-hoc reports only.'], ['an ad-hoc report without sessions', ADHOC, 'TDoc sessions need configured sessions. Use Configure Sessions first.']].forEach(([what, props, message]) => {
    const x = report({ props: ADHOC });
    x.build();
    Object.keys(props).forEach((k) => x.docProps.setProperty(k, props[k]));
    const snap = x.snapshot();
    const writes = x.writes.length;
    check(`${what}: the action and the save refuse, nothing is changed`, [x.s.enableAdhocSessionColumn(), x.save('utc', {}), x.snapshot() === snap, x.writes.length === writes], [{ ok: false, error: message }, { ok: false, errors: [message], changed: false }, true, true]);
  });
  const busy = report({ props: ADHOC });
  busy.build();
  busy.docProps.setProperty('ADHOC_SESSIONS', sessionsProperty(SESSIONS));
  busy.lock.busy = true;
  check('while an update is running the column is not added', [busy.s.enableAdhocSessionColumn().ok, busy.registration()[0].length], [false, 4]);
  const offline = report({ props: ADHOC });
  offline.build();
  offline.docProps.setProperty('ADHOC_SESSIONS', sessionsProperty(SESSIONS));
  offline.s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 503 }) };
  check('the TDoc list cannot be read: nothing is changed', [offline.s.enableAdhocSessionColumn(), offline.registration()[0].length], [{ ok: false, error: 'Nothing was changed: the TDoc list could not be read (Could not download TDOC list: HTTP 503).' }, 4]);
  const master = report({ template: true, docId: TEMPLATE_ID });
  check('in the master template every TDoc-session action refuses', ['enableAdhocSessionColumn', 'saveAdhocTdocSessions', 'assignAdhocTdocSessions'].map((fn) => /SA4 Report Template itself/.test(thrown(() => master.s[fn]({ clock: 'utc', tdocs: {} })) || '')), [true, true, true]);
}

// ================================================================ 8. formal revisions

console.log('8. formal revisions');
{
  const r = report();
  r.build();
  const rows = r.registration();
  check('the original and its revision are two rows; neither is there twice', [rows.filter((row) => row[0] === 'S4aA269002').length, rows.filter((row) => row[0] === 'S4aA269010').length], [1, 1]);
  check('each has the session of its own upload time: the original 22 Sep, the revision 23 Sep', [r.sessionCells().S4aA269002, r.sessionCells().S4aA269010], ['A01', 'A02']);
  const res = r.s.makeAdhocTdocSessionResolver_();
  const all = r.all();
  check('the revision relation plays no part: the same result without it', (() => {
    const x = report({ tdocs: TDOCS.map((t) => [t[0], t[1], t[2], t[3], '', '']) });
    x.build();
    return [x.sessionCells().S4aA269002, x.sessionCells().S4aA269010];
  })(), ['A01', 'A02']);
  check('a revision that is only reserved has no session, whatever its original has', (() => {
    const x = report({ tdocs: TDOCS.map((t) => (t[0] === 'S4aA269010' ? [t[0], t[1], 'reserved', '', t[4], t[5]] : t)) });
    x.build();
    return [x.sessionCells().S4aA269002, x.sessionCells().S4aA269010];
  })(), ['A01', DASH]);
  check('a revision uploaded before its original\'s session ended is in that session, by its own time', (() => {
    const x = report({ tdocs: TDOCS.map((t) => (t[0] === 'S4aA269010' ? [t[0], t[1], t[2], '2026-09-22 15:30:00', t[4], t[5]] : t)) });
    x.build();
    return x.sessionCells().S4aA269010;
  })(), 'A01');
  check('the build still places the revision below its original and fills the original\'s Disposition', (() => {
    const tables = r.body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc');
    const ids = tables.map((t) => t.getRow(0).getCell(1).getText());
    const original = tables[ids.indexOf('S4aA269002')];
    return [ids.indexOf('S4aA269010') - ids.indexOf('S4aA269002'), rowsOf(original).filter((row) => row[0] === 'Disposition')[0][1]];
  })(), [1, 'Revised to S4aA269010']);
  check('a manual assignment of the original does not touch the revision, and the other way round', (() => {
    r.save('utc', { S4aA269002: { mode: 'add', sessions: ['s3'] } });
    const a = [r.sessionCells().S4aA269002, r.sessionCells().S4aA269010];
    r.save('utc', { S4aA269010: { mode: 'set', sessions: ['s3'] } });
    return a.concat([r.sessionCells().S4aA269002, r.sessionCells().S4aA269010]);
  })(), ['A01, A03', 'A02', 'A01', 'A03']);
  check('nothing in the assignment code looks at revisions (source)', /[Rr]evis/.test(CODE.slice(CODE.indexOf('\nfunction assignAdhocSession_('), CODE.indexOf('\nfunction adhocSessionIdsKey_(')).replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '')), false);
  void res; void all;
}

// ================================================================ 9. sessions that change

console.log('9. sessions that change');
{
  const r = report();
  r.build();
  r.save('utc', { S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269004: { mode: 'set', sessions: ['s2'] } });
  const stored = r.docProps.getProperty(KEY);
  const rows = (list) => list.map((x) => Object.assign({}, x));

  const renamed = r.s.saveAdhocSessionsConfiguration(rows([session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), S2, S3]));
  check('renaming a session: saved; the dialog is told that the column follows with the next update', [renamed.ok, renamed.changed, renamed.notice],
    [true, true, 'The sessions were saved. The Session column of the registration table follows with the next update: SA4 Report > Report > Update Report Now.']);
  check('the manual assignments are untouched: they hold ids, not labels', r.docProps.getProperty(KEY) === stored, true);
  r.update();
  check('after the update the column shows the new label, also in the manual assignment', [r.sessionCells().S4aA269002, r.sessionCells().S4aA269001], ['A01', 'A01, A03']);

  // Session 1 is moved to the 26th, without times: it is now the last session.
  r.s.saveAdhocSessionsConfiguration(rows([session('s1', '2026-09-26', '', '', 'Kick-off'), S2, S3]));
  r.update();
  check('re-dating a session: the automatic sessions are worked out again', [r.sessionCells().S4aA269002, r.sessionCells().S4aA269003, r.sessionCells().S4aA269006, r.sessionCells().S4aA269007], ['A01', 'A01', 'A02', 'A03']);
  check('the manual assignments still name the same sessions, shown in the new order', [r.sessionCells().S4aA269001, r.sessionCells().S4aA269004, r.docProps.getProperty(KEY) === stored], ['A01, A02', 'A01', true]);
  r.s.saveAdhocSessionsConfiguration(rows([S1, S2, S3]));
  r.update();

  // A session that manual assignments name is not removed.
  const before = JSON.stringify(r.docProps._store);
  const bodyBefore = r.snapshot();
  const refused = r.s.saveAdhocSessionsConfiguration(rows([S1, S2]));
  check('removing a session that a manual assignment names is refused, naming the TDoc and the way', refused,
    { ok: false, errors: ['24 Sep is named by manual TDoc session assignments (S4aA269001) and was not removed. Change those assignments first: Sessions and Attendance > Assign TDoc Sessions….'], changed: false, sessions: [] });
  check('nothing was changed: not the sessions, not the assignments, not the report', [JSON.stringify(r.docProps._store) === before, r.snapshot() === bodyBefore], [true, true]);
  check('the same for a session named by a replacement', r.s.saveAdhocSessionsConfiguration(rows([S1, S3])).errors.map((e) => e.split(' is named')[0]), ['23 Sep']);
  const sessionOnlyAutomatic = r.s.saveAdhocSessionsConfiguration(rows([S2, S3]));
  check('a session that TDocs belong to only automatically can be removed', [sessionOnlyAutomatic.ok, r.s.getAdhocSessions_().map((x) => x.id)], [true, ['s2', 's3']]);
  r.update();
  check('its TDocs are then placed among the sessions that remain', [r.sessionCells().S4aA269002, r.sessionCells().S4aA269003, r.sessionCells().S4aA269001], ['A01', 'A01', 'A01, A02']);
  check('after the assignment that named it was changed, the session can go', [r.save('utc', { S4aA269004: { mode: 'set', sessions: ['s2'] } }).ok, r.s.saveAdhocSessionsConfiguration(rows([S2])).ok, r.s.getAdhocSessions_().map((x) => x.id)], [true, true, ['s2']]);
  check('no manual assignment names a session that is gone', Object.keys(r.s.adhocTdocSessionReferences_(r.docProps.getProperty(KEY))), ['s2']);
  check('a session added afterwards does not get the id of a removed one', r.s.saveAdhocSessionsConfiguration([S2, session('', '2026-10-01')]).sessions.map((x) => x.id), ['s2', 's4']);

  // The pure function, with what it is told.
  const tell = (refs, list) => { const l = loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(SESSIONS) } }); const out = l.sandbox.saveAdhocSessionsWith_(l.docProps, true, list, [], refs); return [out.ok, l.docProps.getProperty('ADHOC_SESSIONS') === sessionsProperty(SESSIONS)]; };
  check('no manual assignments: sessions are removed as before', [tell({}, rows([S2, S3])), tell(undefined, rows([S2, S3]))], [[true, false], [true, false]]);
  check('a manual assignment naming a session that stays: the others can be removed', tell({ s2: ['S4aA269004'] }, rows([S2])), [true, false]);
  check('manual assignments of a newer version (what they name is unknown): no session is removed, but sessions can be edited and added',
    [tell(null, rows([S2, S3])), tell(null, rows([S1, S2, S3]).concat([session('', '2026-10-01')])), loadCode().sandbox.saveAdhocSessionsWith_(loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(SESSIONS) } }).docProps, true, rows([S2, S3]), [], null).errors[0]],
    [[false, true], [true, false], 'The manual TDoc session assignments were written by a newer release, so it cannot be told which sessions they name. No session was removed.']);
  check('many TDocs naming one session: three are listed, the rest counted', (() => { const l = loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(SESSIONS) } });
    return l.sandbox.saveAdhocSessionsWith_(l.docProps, true, rows([S2, S3]), [], { s1: ['S4aA269001', 'S4aA269002', 'S4aA269003', 'S4aA269004', 'S4aA269005'] }).errors[0]; })(),
  '22 Sep is named by manual TDoc session assignments (S4aA269001, S4aA269002, S4aA269003 and 2 more) and was not removed. Change those assignments first: Sessions and Attendance > Assign TDoc Sessions….');
  check('the attendance guard of stage D is still the first one', (() => { const l = loadCode({ documentProperties: { ADHOC_SESSIONS: sessionsProperty(SESSIONS) } });
    return l.sandbox.saveAdhocSessionsWith_(l.docProps, true, rows([S2, S3]), ['s1'], { s1: ['S4aA269001'] }).errors[0].indexOf('has imported attendance'); })(), 7);

  // A four-column report gets no notice about a column it does not have.
  const four = report({ props: ADHOC });
  four.build();
  four.docProps.setProperty('ADHOC_SESSIONS', sessionsProperty(SESSIONS));
  check('saving sessions in a four-column report says nothing about a Session column', four.s.saveAdhocSessionsConfiguration(rows([session('s1', '2026-09-22', '15:00', '18:00', 'Kick-off'), S2, S3])).notice, undefined);
}

// ================================================================ 10. the dialog

/** Runs the Assign TDoc Sessions dialog's client script against a minimal DOM. */
function runDialog(html, server) {
  const scripts = html.match(/<script>([\s\S]*?)<\/script>/g) || [];
  const element = (tag) => {
    const e = { tag: tag, children: [], style: {}, value: '', textContent: '', className: '', disabled: false, checked: false };
    e.appendChild = (c) => { e.children.push(c); if (tag === 'select' && e.children.length === 1 && !e.value) e.value = c.value; return c; };
    return e;
  };
  const byId = {};
  (html.match(/id="([A-Za-z]+)"/g) || []).map((m) => m.slice(4, -1)).forEach((id) => { byId[id] = element(id); });
  const calls = [];
  const alerts = [];
  let handlers = {};
  let closed = 0;
  const runner = { withSuccessHandler: (fn) => { handlers.success = fn; return runner; }, withFailureHandler: (fn) => { handlers.failure = fn; return runner; } };
  ['saveAdhocTdocSessions', 'enableAdhocSessionColumn'].forEach((name) => {
    runner[name] = (...args) => { calls.push([name].concat(JSON.parse(JSON.stringify(args)))); const h = handlers; handlers = {}; if (server) h.success(JSON.parse(JSON.stringify(server[name](...args)))); else calls.handlers = h; };
  });
  const ctx = { document: { getElementById: (id) => { if (!byId[id]) throw new Error('no element #' + id); return byId[id]; }, createElement: element },
    google: { script: { run: runner, host: { close: () => { closed++; } } } }, alert: (m) => alerts.push(m) };
  vm.createContext(ctx);
  vm.runInContext(scripts[0].replace(/^<script>|<\/script>$/g, ''), ctx);
  const rows = () => byId.rows.children;
  const row = (id) => rows().filter((tr) => tr.children[0].textContent === id)[0];
  const d = {
    ctx: ctx, el: byId, calls: calls, alerts: alerts, scriptCount: scripts.length, closed: () => closed,
    ids: () => rows().map((tr) => tr.children[0].textContent),
    visibleIds: () => rows().filter((tr) => tr.style.display !== 'none').map((tr) => tr.children[0].textContent),
    // The cells of a row: TDoc, title, uploaded (read-only), automatic, assignment, sessions, result.
    uploaded: (id) => row(id).children[2].textContent,
    automatic: (id) => row(id).children[3].textContent,
    result: (id) => row(id).children[6].textContent,
    mode: (id) => row(id).children[4].children[0],
    boxes: (id) => row(id).children[5].children.map((label) => label.children[0]),
    setMode: (id, mode) => { const m = d.mode(id); m.value = mode; m.onchange(); },
    tick: (id, index, on) => { const b = d.boxes(id)[index]; b.checked = on !== false; b.onchange(); },
    visible: (id) => byId[id].style.display !== 'none'
  };
  return d;
}

console.log('10. the dialog');
{
  const r = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: S.serializeAdhocTdocSessions_({ clock: 'utc', tdocs: { S4aA269001: { mode: 'add', sessions: ['s3'] }, S4aA269005: { mode: 'set', sessions: [] } } }) }) });
  r.build();
  const props0 = JSON.stringify(r.docProps._store);
  const body0 = r.snapshot();
  const writes0 = r.writes.length;
  r.s.assignAdhocTdocSessions();
  check('one dialog, titled Assign TDoc Sessions; opening it reads the TDoc list and changes nothing', [r.ui.dialogs.length, r.ui.dialogs[0].title, r.ui.alerts.length, JSON.stringify(r.docProps._store) === props0, r.snapshot() === body0, r.writes.length === writes0], [1, 'Assign TDoc Sessions', 0, true, true, true]);
  const d = runDialog(r.ui.dialogs[0].html, r.s);
  check('one script; one row per TDoc, in the order of the TDoc numbers', [d.scriptCount, d.ids()], [1, TDOCS.map((t) => t[0])]);
  check('each row shows its automatic session, or why there is none', TDOCS.map((t) => d.automatic(t[0])),
    ['22 Sep', '22 Sep', '22 Sep', '23 Sep', '23 Sep', '24 Sep', DASH + ' (uploaded after the last session)', DASH + ' (not uploaded yet)', DASH + ' (upload time cannot be read)', '23 Sep']);
  check('the three kinds of assignment, by name', d.mode('S4aA269002').children.map((o) => [o.value, o.textContent]), [['auto', 'Automatic'], ['add', 'Automatic, and also'], ['set', 'Only these']]);
  check('what is stored is shown: automatic, an addition, a replacement by nothing', [d.mode('S4aA269002').value, d.mode('S4aA269001').value, d.boxes('S4aA269001').map((b) => b.checked), d.mode('S4aA269005').value, d.boxes('S4aA269005').map((b) => b.checked)],
    ['auto', 'add', [false, false, true], 'set', [false, false, false]]);
  check('the result of each row, and manual rows are marked', [d.result('S4aA269002'), d.result('S4aA269001'), d.result('S4aA269005'), d.result('S4aA269008'), ['S4aA269002', 'S4aA269001', 'S4aA269005'].map((id) => d.el.rows.children.filter((tr) => tr.children[0].textContent === id)[0].children[6].className)],
    ['22 Sep', '22 Sep, 24 Sep', DASH, DASH, ['', 'manual', 'manual']]);
  check('an automatic row cannot be ticked; a manual row can', [d.boxes('S4aA269002').map((b) => b.disabled), d.boxes('S4aA269001').map((b) => b.disabled)], [[true, true, true], [false, false, false]]);
  check('the counts', d.el.counts.textContent, '10 TDocs, 2 assigned by hand, 4 without a session');
  check('the clock is shown as stored; the report has the column, so no button to add it', [d.el.clock.value, d.el.columnState.textContent, d.visible('enableBtn')], ['utc', 'This report has the Session column.', false]);
  check('the clock is not part of the normal dialog: it is under a closed "Advanced", and there is no note about it', (() => {
    const html = r.ui.dialogs[0].html;
    const details = (html.match(/<details[^>]*>[\s\S]*?<\/details>/) || [''])[0];
    return [/id="clock"/.test(details), /<details[^>]*\bopen\b/.test(details), /<summary[^>]*>Advanced<\/summary>/.test(details), html.indexOf('id="clock"') > html.indexOf('id="rows"'), d.visible('clockNote') && !!d.el.clockNote.textContent, /time zone|UTC/.test(html.replace(details, '').replace(/<script>[\s\S]*<\/script>/, ''))];
  })(), [true, false, true, true, false, false]);
  check('nothing was sent by opening', d.calls, []);

  // Editing: the result follows, nothing is sent.
  d.setMode('S4aA269002', 'add');
  d.tick('S4aA269002', 2);
  check('"Automatic, and also" 24 Sep: the automatic session stays, the other is added', [d.result('S4aA269002'), d.boxes('S4aA269002').map((b) => b.disabled)], ['22 Sep, 24 Sep', [false, false, false]]);
  d.setMode('S4aA269002', 'set');
  check('"Only these" with the same tick: only that session', d.result('S4aA269002'), '24 Sep');
  d.tick('S4aA269002', 2, false);
  check('"Only these" with nothing ticked: no session', d.result('S4aA269002'), DASH);
  d.setMode('S4aA269002', 'auto');
  check('back to Automatic: the automatic session again, whatever is ticked', [d.result('S4aA269002'), d.boxes('S4aA269002').map((b) => b.disabled)], ['22 Sep', [true, true, true]]);
  d.setMode('S4aA269004', 'set');
  d.tick('S4aA269004', 0);
  d.tick('S4aA269004', 2);
  d.setMode('S4aA269006', 'add');
  d.setMode('S4aA269001', 'auto');
  check('the counts follow', d.el.counts.textContent, '10 TDocs, 3 assigned by hand, 4 without a session');
  check('still nothing sent, nothing stored', [d.calls.length, JSON.stringify(r.docProps._store) === props0], [0, true]);

  // Filter.
  d.el.filter.value = '26900';
  d.ctx.applyFilter();
  check('the filter narrows the rows by TDoc number', d.visibleIds(), ['S4aA269001', 'S4aA269002', 'S4aA269003', 'S4aA269004', 'S4aA269005', 'S4aA269006', 'S4aA269007', 'S4aA269008', 'S4aA269009']);
  d.el.filter.value = 'TITLE OF S4AA269010';
  d.ctx.applyFilter();
  check('and by title, in any letter case', d.visibleIds(), ['S4aA269010']);
  d.el.filter.value = '';
  d.ctx.applyFilter();
  check('an empty filter shows all', d.visibleIds().length, 10);

  // Cancel.
  d.ctx.google.script.host.close();
  check('Cancel after editing: nothing sent, nothing stored, the report untouched', [d.closed(), d.calls.length, JSON.stringify(r.docProps._store) === props0, r.snapshot() === body0], [1, 0, true, true]);
  check('the Cancel button only closes the dialog (source)', (r.ui.dialogs[0].html.match(/<button[^>]*id="closeBtn"[^>]*>/) || [''])[0], '<button type="button" class="grey" id="closeBtn" onclick="google.script.host.close()">');

  // Save.
  const writesBefore = r.writes.length;
  d.ctx.saveAssignments();
  check('Save sends the clock and only the TDocs that are not automatic; an "also" without a tick is automatic', d.calls, [['saveAdhocTdocSessions', { clock: 'utc', tdocs: { S4aA269004: { mode: 'set', sessions: ['s1', 's3'] }, S4aA269005: { mode: 'set', sessions: [] } } }]]);
  check('one property is written, once; the dialog closes', [r.writes.slice(writesBefore), r.docProps.getProperty(KEY), d.closed()], [[KEY], '{"v":1,"clock":"utc","add":{},"set":{"":["S4aA269005"],"s1+s3":["S4aA269004"]}}', 2]);
  check('the report shows it at once; the TDoc reset to Automatic has its automatic session again', [r.sessionCells().S4aA269004, r.sessionCells().S4aA269005, r.sessionCells().S4aA269001], ['A01, A03', DASH, 'A01']);

  // A refused save, a failed request, a notice.
  r.s.assignAdhocTdocSessions();
  const e = runDialog(r.ui.dialogs[1].html, null);
  e.ctx.saveAssignments();
  check('Save is disabled while the request runs', e.el.saveBtn.disabled, true);
  e.calls.handlers.success({ ok: false, errors: ['First problem.', 'Second problem.'], changed: false });
  check('a refused save: the messages are shown, the dialog stays open', [e.el.errors.textContent, e.closed(), e.el.saveBtn.disabled], ['First problem.\nSecond problem.', 0, false]);
  e.ctx.saveAssignments();
  e.calls.handlers.failure(new Error('Server error'));
  check('a failed request', [e.el.errors.textContent, e.closed(), e.el.saveBtn.disabled], ['Server error', 0, false]);
  e.ctx.saveAssignments();
  e.calls.handlers.success({ ok: true, changed: true, notice: 'Saved, but the column follows later.' });
  check('a save with a notice: the notice is shown, then the dialog closes', [e.alerts, e.closed()], [['Saved, but the column follows later.'], 1]);

  // The clock.
  r.s.assignAdhocTdocSessions();
  const c = runDialog(r.ui.dialogs[2].html, r.s);
  check('the two choices under Advanced, in words, the normal one first', /<option value="utc">in UTC, as the 3GPP Portal records them \(normal\)<\/option>\s*<option value="session">already on the same clock as the session times<\/option>/.test(r.ui.dialogs[2].html), true);
  check('and what it is for', /The Portal records upload times in UTC; they are converted to the time zone of this report automatically\. Change this only for a TDoc list that does not come from the Portal\./.test(r.ui.dialogs[2].html), true);
  check('a normal save sends the Portal\'s clock without the user choosing anything', c.el.clock.value, 'utc');
  c.el.clock.value = 'session';
  c.ctx.saveAssignments();
  check('choosing the session clock under Advanced and saving stores the choice and applies it', [c.calls[0][1].clock, JSON.parse(r.docProps.getProperty(KEY)).clock, r.sessionCells().S4aA269006], ['session', 'session', 'A02']);
  r.s.assignAdhocTdocSessions();
  const c2 = runDialog(r.ui.dialogs[3].html, r.s);
  check('the next dialog keeps the choice, shows the automatic sessions it gives, and says so outside Advanced',
    [c2.el.clock.value, c2.automatic('S4aA269004'), c2.automatic('S4aA269006'), c2.visible('clockNote'), c2.el.clockNote.textContent],
    ['session', '22 Sep', '23 Sep', true, 'This report compares the upload times as they stand, without converting them from UTC (see Advanced).']);
  c2.ctx.saveAssignments();
  check('saving that dialog without touching Advanced keeps the choice', [c2.calls[0][1].clock, JSON.parse(r.docProps.getProperty(KEY)).clock], ['session', 'session']);

  // A report without the column: the explicit action, in two steps.
  const four = report({ props: ADHOC });
  four.build();
  four.docProps.setProperty('ADHOC_SESSIONS', sessionsProperty(SESSIONS));
  four.s.assignAdhocTdocSessions();
  const f = runDialog(four.ui.dialogs[0].html, four.s);
  check('a four-column report: the dialog says so and offers the column', [f.el.columnState.textContent, f.visible('enableBtn'), f.visible('enableConfirm')], ['This report has no Session column: assignments are stored, but not shown. ', true, false]);
  const fourBody = four.snapshot();
  f.ctx.askEnable();
  check('the button only asks; nothing is sent', [f.visible('enableConfirm'), f.visible('enableBtn'), f.calls.length, four.snapshot() === fourBody], [true, false, 0, true]);
  check('the question says that the report then needs this release', /Add a Session column to the registration table of this report\? A report with this column needs this release or a later one\./.test(four.ui.dialogs[0].html), true);
  f.ctx.cancelEnable();
  check('No: nothing is changed', [f.visible('enableConfirm'), f.visible('enableBtn'), four.snapshot() === fourBody], [false, true, true]);
  f.ctx.askEnable();
  f.ctx.enableColumn();
  check('Yes: the column is added, and the dialog says so', [f.calls, f.el.status.textContent, f.el.columnState.textContent, f.visible('enableBtn'), four.registration()[0].length], [[['enableAdhocSessionColumn']], 'The Session column was added to the registration table (10 TDocs).', 'This report has the Session column.', false, 5]);
  const onclick = (id) => (four.ui.dialogs[0].html.match(new RegExp('<button[^>]*id="' + id + '"[^>]*onclick="([^"]*)"')) || [])[1];
  check('what each button does (source)', ['saveBtn', 'closeBtn', 'enableBtn', 'enableYes', 'enableNo'].map(onclick), ['saveAssignments()', 'google.script.host.close()', 'askEnable()', 'enableColumn()', 'cancelEnable()']);
  const unbuilt = report();
  unbuilt.s.assignAdhocTdocSessions();
  const u = runDialog(unbuilt.ui.dialogs[0].html, unbuilt.s);
  check('a report that is not built yet: no column to add, assignments can still be stored', [u.el.columnState.textContent, u.visible('enableBtn'), u.ids().length], ['This report has no registration table yet. ', false, 10]);

  // The model.
  const res = r.s.makeAdhocTdocSessionResolver_();
  const model = r.s.buildAdhocTdocSessionsDialogModel_(res, r.groups(), 5);
  check('what the dialog is given: sessions by id and label, the clock, and per TDoc its number, a title, the automatic session, and the manual assignment',
    [Object.keys(model), model.sessions, Object.keys(model.tdocs[0]), model.tdocs.length], [['sessions', 'clock', 'uploadedZone', 'notice', 'readOnly', 'column', 'none', 'tdocs'], [{ id: 's1', label: '22 Sep' }, { id: 's2', label: '23 Sep' }, { id: 's3', label: '24 Sep' }], ['id', 'title', 'uploaded', 'auto', 'why', 'mode', 'sessions'], 10]);
  // The upload time is given to this dialog as a short text to look at (G2); nothing else of the list is.
  check('of the upload time the dialog gets a display text only -- no raw value of the list, no ISO time -- and no contact or source', [/2026-09-|Sam Rivera|ExampleCorp|"at"|dateOnly|unreadable/.test(JSON.stringify(model)), model.clock, model.uploadedZone, model.tdocs.map((x) => x.uploaded)],
    // This report compares the times as the list has them (the compatibility clock), so they are shown as the list has them.
    [false, 'session', '', ['18 Sep 07:26', '22 Sep 14:00', '22 Sep 16:00', '22 Sep 16:00', '23 Sep 21:59', '23 Sep 22:00', '25 Sep 09:00', DASH, DASH, '23 Sep 08:00']]);
  const long = report({ tdocs: [['S4aA269001', '4.3', 'available', '2026-09-18 07:26:14', '', '']] });
  long.s.downloadAndGroupTdocs_ = () => ({ '4.3': { tdocs: [{ row: ['S4aA269001', '  A  very   long title ' + 'x'.repeat(300)], tdocCol: 0, titleCol: 1, revisedToCol: -1, uploaded: { at: '2026-09-18T07:26:14', dateOnly: null, unreadable: false } }] } });
  check('a title is shortened to 90 characters, white space collapsed', long.s.buildAdhocTdocSessionsDialogModel_(long.s.makeAdhocTdocSessionResolver_(), long.s.downloadAndGroupTdocs_(), 0).tdocs[0].title.length, 90);
  const absent = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: S.serializeAdhocTdocSessions_({ clock: 'utc', tdocs: { S4aA269999: { mode: 'set', sessions: ['s2'] } } }) }) });
  absent.s.assignAdhocTdocSessions();
  const a = runDialog(absent.ui.dialogs[0].html, absent.s);
  a.ctx.saveAssignments();
  check('a manual assignment of a TDoc that is not in the list now is shown, and a save keeps it', [a.ids().slice(-1), a.automatic('S4aA269999'), a.calls[0][1].tdocs, absent.docProps.getProperty(KEY)],
    [['S4aA269999'], DASH + ' (not in the TDoc list)', { S4aA269999: { mode: 'set', sessions: ['s2'] } }, '{"v":1,"clock":"utc","add":{},"set":{"s2":["S4aA269999"]}}']);

  // Untrusted text.
  const hostile = report({ props: Object.assign({}, ADHOC, { ADHOC_SESSIONS: sessionsProperty([session('s1', '2026-09-22', '', '', '</script><b>x')]) }) });
  const realGroups = hostile.s.downloadAndGroupTdocs_;
  hostile.s.downloadAndGroupTdocs_ = (cfg) => { const g = realGroups(cfg); Object.keys(g).forEach((k) => g[k].tdocs.forEach((td) => { td.row[td.titleCol] = '<img src=x onerror=alert(1)></script>'; })); return g; };
  hostile.s.assignAdhocTdocSessions();
  const h = runDialog(hostile.ui.dialogs[0].html, hostile.s);
  check('a title and a label with markup: one script element, shown as typed', [h.scriptCount, h.el.rows.children[0].children[1].textContent, h.automatic('S4aA269001'), hostile.ui.dialogs[0].html.indexOf('</script><b>'), hostile.ui.dialogs[0].html.indexOf('<img')],
    [1, '<img src=x onerror=alert(1)></script>', '</script><b>x', -1, -1]);
  check('the dialog script builds everything with textContent, never with innerHTML (source)', [/innerHTML|insertAdjacentHTML|document\.write/.test(hostile.ui.dialogs[0].html), /textContent/.test(hostile.ui.dialogs[0].html)], [false, true]);

  // Unreadable and newer stored values.
  const corrupt = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: '{"v":1,' }) });
  corrupt.s.assignAdhocTdocSessions();
  const cd = runDialog(corrupt.ui.dialogs[0].html, corrupt.s);
  check('manual assignments that cannot be read: a notice, every TDoc automatic, saving possible; opening changed nothing', [/cannot be read\. They are not used; saving replaces them\./.test(cd.el.notice.textContent), cd.visible('notice'), cd.el.saveBtn.disabled, cd.el.counts.textContent, corrupt.docProps.getProperty(KEY)],
    [true, true, false, '10 TDocs, 0 assigned by hand, 3 without a session', '{"v":1,']);
  const newer = report({ props: Object.assign({}, WITH_SESSIONS, { [KEY]: '{"v":2}' }) });
  newer.s.assignAdhocTdocSessions();
  const nd = runDialog(newer.ui.dialogs[0].html, newer.s);
  check('a newer version: a notice, nothing can be edited or saved', [/version 2/.test(nd.el.notice.textContent), nd.el.saveBtn.disabled, nd.el.clock.disabled, nd.mode('S4aA269001').disabled, nd.boxes('S4aA269001')[0].disabled], [true, true, true, true, true]);
  check('and the server refuses a save as well', [newer.save('utc', {}).ok, newer.docProps.getProperty(KEY)], [false, '{"v":2}']);

  // Where the dialog does not open.
  const main = report({ props: Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' }) });
  main.s.assignAdhocTdocSessions();
  check('a main-meeting report: a message, no dialog, no download', [main.ui.dialogs.length, main.ui.alerts.map((x) => x[1]), main.downloads], [0, ['TDoc sessions are available for ad-hoc reports only.'], 0]);
  const noSessions = report({ props: ADHOC });
  noSessions.s.assignAdhocTdocSessions();
  check('an ad-hoc report without sessions: where to configure them, no dialog, no download', [noSessions.ui.dialogs.length, noSessions.ui.alerts.map((x) => x[1]), noSessions.downloads],
    [0, ['Sessions are assigned to TDocs once the sessions of this report are configured:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…'], 0]);
  const offline = report();
  offline.s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 503 }) };
  offline.s.assignAdhocTdocSessions();
  check('the TDoc list cannot be read: a message, no dialog', [offline.ui.dialogs.length, offline.ui.alerts.map((x) => x[1])], [0, ['The TDoc list could not be read:\n\nCould not download TDOC list: HTTP 503']]);

  // A save whose column refresh fails.
  const stuck = report();
  stuck.build();
  stuck.s.refreshRegistrationSessionColumn_ = () => { throw new Error('table is locked'); };
  check('saved, but the column could not be updated now: the result says that the next update does it', stuck.save('utc', { S4aA269001: { mode: 'add', sessions: ['s3'] } }),
    { ok: true, errors: [], changed: true, notice: 'The assignments were saved. The Session column could not be updated now (table is locked); the next update of the report does it.' });
  const noChange = report();
  noChange.build();
  const downloads = noChange.downloads;
  check('a save that changes nothing does not read the TDoc list or touch the report', [noChange.save('utc', {}), noChange.downloads === downloads, noChange.lock.tries], [{ ok: true, errors: [], changed: false }, true, 0]);
}

// ================================================================ 11. menu, adoption, nothing else changed

function openMenu(props, opts) {
  const o = opts || {};
  const loaded = loadTemplateRuntime({ release: o.release === false ? null : RELEASE, documentProperties: props });
  const ui = { menus: [] };
  ui.createMenu = (name) => {
    const m = { name: name, entries: [] };
    m.addItem = (label, fn) => { m.entries.push({ label: label, fn: fn }); return m; };
    m.addSeparator = () => { m.entries.push({ sep: true }); return m; };
    m.addSubMenu = (sub) => { m.entries.push({ sub: sub }); return m; };
    m.addToUi = () => { ui.menus.push(m); };
    return m;
  };
  loaded.sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => o.docId || REPORT_ID, getBody: () => makeFakeDocumentBody(loaded.sandbox) });
  loaded.sandbox.DocumentApp.getUi = () => ui;
  loaded.sandbox.onOpen();
  const sub = ui.menus[0].entries.filter((e) => e.sub && /Sessions/.test(e.sub.name)).map((e) => e.sub)[0];
  return { loaded: loaded, sub: sub, text: JSON.stringify(ui.menus), docProps: loaded.docProps };
}
function gitShow(spec) {
  try {
    return execFileSync('git', ['show', spec], { cwd: path.join(__dirname, '..'), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    return null;
  }
}

console.log('11. menu, adoption, and nothing else changed');
{
  const adhoc = openMenu(Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, WITH_SESSIONS));
  check('ad-hoc report: the submenu, with Assign TDoc Sessions… as its fourth item', adhoc.sub.entries.slice(0, 4), [{ label: 'Configure Sessions…', fn: 'configureAdhocSessions' }, { label: 'Import Teams Attendance…', fn: 'importTeamsAttendance' },
    { label: 'Refresh Attendance Section', fn: 'refreshAdhocAttendanceSection' }, { label: 'Assign TDoc Sessions…', fn: 'assignAdhocTdocSessions' }]);
  check('its function exists and is public', [typeof adhoc.loaded.sandbox.assignAdhocTdocSessions, /_$/.test('assignAdhocTdocSessions')], ['function', false]);
  check('main-meeting report, new report, master template, CENTRAL / Legacy: no such item', [openMenu(Object.assign({}, WITH_SESSIONS, { MEETING_TYPE: 'main' })).text, openMenu({}).text, openMenu(WITH_SESSIONS, { docId: TEMPLATE_ID }).text, openMenu(WITH_SESSIONS, { release: false }).text].map((t) => /TDoc Sessions/.test(t)), [false, false, false, false]);
  check('opening a report writes nothing', openMenu(WITH_SESSIONS).docProps._store, WITH_SESSIONS);

  // Adoption.
  const r = report();
  r.s.saveAdhocTdocSessions({ clock: 'session', tdocs: { S4aA269001: { mode: 'add', sessions: ['s3'] } } });
  const adopted = r.s.adoptReportDocumentForAddon_();
  check('the manual assignments are adopted with the sessions: an update in the background needs both', [adopted.verified, adopted.copiedKeys.filter((k) => /ADHOC/.test(k)), r.scriptProps.getProperty('SA4_STATE|' + REPORT_ID + '|' + KEY) === r.docProps.getProperty(KEY)], [true, ['ADHOC_SESSIONS', KEY], true]);
  const background = r.s.makeAdhocTdocSessionResolver_({ documentId: REPORT_ID, mode: 'addon-background' });
  check('a background run reads them from the central copy, the stored clock included', [background.clock, Object.keys(background.overrides)], ['session', ['S4aA269001']]);
  check('a background run of a report without stored assignments uses the Portal\'s clock', (() => { const x = report(); x.s.adoptReportDocumentForAddon_(); return x.s.makeAdhocTdocSessionResolver_({ documentId: REPORT_ID, mode: 'addon-background' }).clock; })(), 'utc');
  check('a document without them: nothing is invented by adoption', (() => { const x = report({ props: ADHOC }); return [x.s.adoptReportDocumentForAddon_().copiedKeys.filter((k) => /ADHOC/.test(k)), x.scriptProps.getKeys().filter((k) => /ADHOC/.test(k))]; })(), [[], []]);
  check('it is one fixed key, not a family of keys', [r.s.ADDON003_ADOPTION_FIXED_KEYS_.filter((k) => /ADHOC/.test(k)), r.s.ADDON003_ADOPTION_PREFIX_FAMILIES_], [['ADHOC_SESSIONS', KEY], ['DISCUSS_', 'REVIS_', 'A1_EMPTY_CACHE_', 'REVIEWER_NO_SUMMARY_CACHE_']]);
  check('Clear Collection Caches keeps the manual assignments', (() => { const before = r.docProps.getProperty(KEY); r.s.clearAllCaches(); return r.docProps.getProperty(KEY) === before; })(), true);

  // A main-meeting report and an ad-hoc report without sessions: the update is what it was.
  const MAIN = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137_Synthetic', MEETING_NUMBER: '137', MEETING_ID: '60778', REPORT_SUFFIX: 'Audio', AGENDA_ITEM_PREFIX: '4.' };
  [['a main-meeting report', MAIN], ['a main-meeting report with session properties left in it', Object.assign({}, MAIN, { ADHOC_SESSIONS: sessionsProperty(SESSIONS), [KEY]: '{"v":1,"clock":"utc","add":{},"set":{}}' })], ['an ad-hoc report without sessions', ADHOC]].forEach(([what, props]) => {
    const x = report({ props: props });
    x.body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4aA269001', 'Synthetic title of S4aA269001', 'ExampleCorp', '4.3']]);
    const reads = [];
    const get = x.docProps.getProperty;
    x.docProps.getProperty = (k) => { reads.push(k); return get(k); };
    const result = x.update();
    check(`${what}: the update completes, appends four-cell rows, and never reads the upload times or the assignments`,
      [result, x.registration().length - 1, x.registration().every((row) => row.length === 4), x.displayReads, reads.filter((k) => k === KEY).length], [{ success: true, error: null }, 10, true, 0, 0]);
  });

  // Against the release before the feature.
  const OLD_CODE = gitShow('template-release/T-2026.10.4:Code.js');
  if (!OLD_CODE) {
    console.log('  note: template-release/T-2026.10.4 is not available in this checkout; the comparison with it is skipped.');
  } else {
    // The same list, the same four-column report, the old and the new update: the same registration table.
    const run = (code) => {
      const x = report({ props: ADHOC });
      if (code) vm.runInContext(code.slice(code.indexOf('\nfunction updateRegisteredDocumentsTable_('), code.indexOf('\nfunction updateAll(')), x.s);
      x.body.appendTable([['TDoc', 'Title', 'Source', 'Agenda Item'], ['S4aA269001', 'Synthetic title of S4aA269001', 'ExampleCorp', '4.3']]);
      const groups = x.groups();
      const all = [];
      Object.keys(groups).forEach((key) => groups[key].tdocs.forEach((td) => all.push(Object.assign({}, td, { agendaItem: key }))));
      x.s.updateRegisteredDocumentsTable_(x.body, all, groups);
      return x.snapshot();
    };
    check('updateRegisteredDocumentsTable_() on a four-column table gives what the T-2026.10.4 function gives', run(null) === run(OLD_CODE), true);
    // The hazard the design names: the old function does not recognize a five-column table.
    const old = report();
    old.build();
    vm.runInContext(OLD_CODE.slice(OLD_CODE.indexOf('\nfunction updateRegisteredDocumentsTable_('), OLD_CODE.indexOf('\nfunction updateAll(')), old.s);
    old.tdocs = TDOCS.concat([['S4aA269011', '4.4', 'available', '2026-09-23 12:00:00', '', '']]);
    const before = old.registration().length;
    const groups = old.groups();
    old.s.updateRegisteredDocumentsTable_(old.body, old.s.flattenTdocGroups_(groups), groups);
    check('the T-2026.10.4 function does not recognize a five-column table and appends nothing to it: a report with the column needs this release', old.registration().length === before, true);
  }

  // Source: what the feature touches.
  const section = CODE.slice(CODE.indexOf('// AD-HOC TDOC SESSIONS (stage B)'), CODE.indexOf('// AD-HOC OPENING (stage E)'));
  const code = section.replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '');
  check('a column is added in two places only: when the table is created with sessions, and by the explicit action (source)',
    [(code.match(/appendTableCell\(/g) || []).length, /function enableAdhocSessionColumn\(\)[\s\S]*appendTableCell\(texts\[r\]\)/.test(code), /appendTableCell/.test(CODE.slice(CODE.indexOf('\nfunction refreshRegistrationSessionColumn_('), CODE.indexOf('\nfunction refreshRegistrationSessionColumnSafely_(')))], [1, true, false]);
  check('the section writes one property, in one place', (code.match(/\.setProperty\(/g) || []).length, 1);
  check('it logs fixed texts and error messages only', (code.match(/Logger\.log\((.*)\);/g) || []).filter((line) => !/^Logger\.log\('TDoc sessions: [^']*' \+ (e\.message|resolver\.error \+ ' Only the automatic sessions are used\.')\);$/.test(line)), []);
  check('an update reads the TDoc list once, as before, and the assignments once', (() => { const x = report(); x.build(); const n = x.downloads; x.update(); return x.downloads - n; })(), 1);
  check('Code.js is version 2.18.0, the release that adds the ad-hoc sessions', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.18.0');
  check('ReportCreator.js: one more menu item, nothing else', (CREATOR.match(/TDoc Sessions|assignAdhocTdocSessions/g) || []).length, 2);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll ad-hoc TDoc session checks passed.');
process.exitCode = failures ? 1 : 0;
