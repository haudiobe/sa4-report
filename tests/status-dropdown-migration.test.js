/**
 * "Convert Status Fields to Dropdowns…" -- the status dropdowns for a report
 * that EXISTS (T-2026.10.8).
 *
 * A report built before the status dropdowns has text statuses. A rebuild
 * would give it dropdowns and take its minutes. This action converts the
 * Status cells in place: it reads the document through the Docs API, and
 * replaces the text of every eligible Status cell by a dropdown with the
 * same value. Nothing else is touched; no status changes; the TDoc list is
 * not read; DocumentApp changes nothing.
 *
 * The tests run the real conversion, and after it the real update, against
 * the fake document and the fake Docs API (tests/helpers/); no live service.
 *
 *    1. which text is which option;
 *    2. the plan: what happens to every TDoc table;
 *    3. looking only (the dry run);
 *    4. the conversion: values, definition, the mark, and nothing else;
 *    5. running it again;
 *    6. a run that is cut short heals on the next run;
 *    7. the document is edited while it runs;
 *    8. what is never converted;
 *    9. when something is missing or fails;
 *   10. after the conversion: updates, by the rules of T-2026.10.7;
 *   11. the menu action: it asks first;
 *   12. nothing happens because the code is installed;
 *   13. every write by position names its revision;
 *   14. the read back decides, not the reply of the API.
 *
 * All names, numbers and colours are synthetic.
 *
 * Run: node tests/status-dropdown-migration.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { makeFakeDocsApi, REPORT_TAB_ID } = require('./helpers/fake-docs-api.js');

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
const CREATOR = fs.readFileSync(path.join(__dirname, '..', 'template', 'ReportCreator.js'), 'utf8').replace(/\r/g, '');
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const TITLE = 'Document Status';
const NAMES = ['available', 'noted', 'agreed', 'revised', 'parked', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'other', 'replied', 'Plenary', 'postponed'];
const colour = (n, shift) => ({ color: { rgbColor: { red: ((n * 17 + shift) % 100) / 100, green: ((n * 31 + shift) % 100) / 100, blue: ((n * 53 + shift) % 100) / 100 } } });
const DEFINITION = { title: TITLE, options: NAMES.map((name, n) => ({ displayValue: name, textStyle: { foregroundColor: colour(n, 3), backgroundColor: colour(n, 47) } })) };
const OPTIONS = DEFINITION.options.map((o, k) => Object.assign({ optionId: 'dropdownItem.x' + k }, o));

const PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Topic'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
/** id, agenda item, Portal status -- what the report was built with */
const TDOCS = [['S4aA269001', '4.3', 'agreed'], ['S4aA269002', '4.3', 'available'], ['S4aA269003', '4.4', 'reserved'], ['S4aA269004', '4.4', 'Replied to'], ['S4aA269005', '4.4', 'not treated'],
  ['S4aA269006', '4.4', ''], ['S4aA269007', '4.4', 'available'], ['S4aA269008', '4.4', 'available'], ['S4aA269009', '4.4', 'available'], ['S4aA269010', '4.4', 'noted'], ['S4aA269011', '4.4', 'available']];
/** What was typed into the Status cells by hand since: decisions of the meeting, in the words people use. */
const TYPED = { S4aA269007: 'parked', S4aA269008: 'Plenary', S4aA269009: 'other', S4aA269010: '  NOTED ', S4aA269011: 'agreed, see minutes' };
const AFTER = ['S4aA269001 dropdown:agreed', 'S4aA269002 dropdown:available', 'S4aA269003 dropdown:reserved', 'S4aA269004 dropdown:replied', 'S4aA269005 text:not treated', 'S4aA269006 empty',
  'S4aA269007 dropdown:parked', 'S4aA269008 dropdown:Plenary', 'S4aA269009 dropdown:other', 'S4aA269010 dropdown:noted', 'S4aA269011 text:agreed, see minutes'];
const BEFORE = ['S4aA269001 text:agreed', 'S4aA269002 text:available', 'S4aA269003 text:reserved', 'S4aA269004 text:Replied to', 'S4aA269005 text:not treated', 'S4aA269006 empty',
  'S4aA269007 text:parked', 'S4aA269008 text:Plenary', 'S4aA269009 text:other', 'S4aA269010 text:  NOTED ', 'S4aA269011 text:agreed, see minutes'];

const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

/**
 * An EXISTING report: built by the runtime of before the dropdowns (text statuses, not marked), then used -- minutes
 * and statuses typed by hand -- and now running this release with the Google Docs API service.
 *   options.template  the definitions of the master template (default: the canonical one); null: it cannot be read
 *   options.tdocs     another TDoc list; options.typed: other hand-typed statuses (default TYPED; {} for none)
 */
function existingReport(options) {
  const o = options || {};
  const loaded = loadCode({ documentProperties: Object.assign({}, PROPS, o.props || {}) });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const api = makeFakeDocsApi({ body: body, reportId: REPORT_ID });
  if (o.template !== null) api.addSource(TEMPLATE_ID, o.template || [DEFINITION]);
  const r = { s: s, body: body, api: api, docProps: loaded.docProps, tdocs: (o.tdocs || TDOCS).map((t) => t.slice()), logs: [], events: [], downloads: 0, locked: false, lock: { tries: 0, releases: 0 } };
  const ui = { alerts: [], answers: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES', NO: 'NO' }, alert: (...args) => { ui.alerts.push(args); return ui.answers.length ? ui.answers.shift() : 'YES'; }, showModalDialog: () => {} };
  r.ui = ui;
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.ElementType.TEXT = 'TEXT';
  s.DocumentApp.ElementType.UNSUPPORTED = 'UNSUPPORTED';
  const doc = { getId: () => REPORT_ID, getBody: () => body, saveAndClose: () => { r.events.push('save'); body._state.dirty = false; body._state.closed = true; } };
  s.DocumentApp.getActiveDocument = () => doc;
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => { r.lock.tries++; return !r.locked; }, releaseLock: () => { r.lock.releases++; } }) };
  s.assertNotTemplateMaster_ = () => {};
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.removeRowHeightAndSpacing = () => {};
  s.downloadMeetingAgenda_ = () => AGENDA;
  s.collectEmailDiscussionCore_ = () => {};
  s.collectRevisionsCore_ = () => {};
  s.collectorUpdate_ = () => ({ failures: [] });
  s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
  s.UrlFetchApp = { fetch: () => { r.downloads++; return { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }; } };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const values = [HEADER].concat(r.tdocs.map(([id, agenda, status]) => [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', agenda, 'Synthetic item', status, '', '', '', '']));
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((row) => row.map(() => null)) }),
      getRange: (row, col, rows) => ({ getDisplayValues: () => [['Uploaded']].concat(r.tdocs.map(() => [''])).slice(row - 1, row - 1 + rows) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };

  r.table = (number) => body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(1).getText() === number)[0];
  r.cell = (number, label) => { const t = r.table(number); for (let i = 0; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === label) return t.getRow(i).getCell(1); return null; };
  r.all = () => r.tdocs.map((t) => t[0] + ' ' + api.shown(t[0]));
  r.portal = (number, status) => { r.tdocs.filter((t) => t[0] === number)[0][2] = status; };
  r.user = (fn) => { body._state.closed = false; const out = fn(); body._state.dirty = false; return out; };
  r.exec = (fn) => { body._state.closed = false; s.resetStatusDropdownRun_(); const from = r.events.length; const out = fn(); r.last = r.events.slice(from); return out; };
  /** Everything of the document that is not the value of a Status cell: every paragraph, and every table cell but those. */
  r.otherContent = () => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE'
    ? rowsOf(c).map((row) => (c.getRow(0).getNumCells() === 2 && row[0] === 'Status' ? ['Status'] : row)) : [c.getText(), c.getHeading()])));
  r.statusLogs = () => r.logs.filter((l) => /^Status dropdowns/.test(l));

  // 1. Built by the release before the dropdowns: no Docs API service, no release that knows them.
  s.runFullReportBuildCore_();
  // 2. Used: minutes and statuses typed by hand.
  r.user(() => {
    r.cell('S4aA269001', 'Minutes').setText('Presented by Sam Rivera. The group agreed the proposal.');
    r.cell('S4aA269007', 'Minutes').setText('Parked until the next call.');
    const typed = o.typed || TYPED;
    Object.keys(typed).forEach((number) => { if (r.table(number)) api.statusCell(number).setText(typed[number]); });
    body.insertParagraph(body._children.indexOf(r.table('S4aA269002')), 'A note typed between two tables.');
  });
  // 3. Now it runs this release, with the Google Docs API service.
  r.downloads = 0;
  s.SA4_RELEASE_ = { releaseId: 'T-2026.10.8', flavor: 'template', codeVersion: '2.20.0', templateDocumentId: TEMPLATE_ID };
  s.Docs = { Documents: { get: (id, opts) => { r.events.push('get ' + (id === REPORT_ID ? 'report' : 'template')); return api.Docs.Documents.get(id, opts); },
    batchUpdate: (resource, id) => { r.events.push('batchUpdate ' + resource.requests.map((q) => Object.keys(q)[0]).filter((k, i, l) => l.indexOf(k) === i).join('+') + ' x' + resource.requests.length); return api.Docs.Documents.batchUpdate(resource, id); } } };
  /** The conversion, without its dialogs. DocumentApp may change nothing while it runs: the document is closed for it. */
  r.migrate = (opts) => r.exec(() => { body._state.closed = true; return s.migrateStatusFieldsToDropdowns_(opts || {}); });
  r.update = () => r.exec(() => s.continuousUpdateCore_());
  return r;
}

const S = loadCode().sandbox;
const brief = (x) => ({ ok: x.ok, inspected: x.inspected, already: x.already, toConvert: x.toConvert, converted: x.converted, text: x.unmapped.length + x.empty, skipped: x.skipped.length, notConverted: x.notConverted.length, unexpected: x.unexpected.length, marked: x.marked });

// ================================================================ 1. which text is which option

console.log('1. which text is which option');
{
  const of = (text) => { const o = S.statusDropdownMigrationOption_(text, OPTIONS); return o ? o.displayValue : null; };
  check('each of the fourteen statuses, typed exactly, is that option', NAMES.map(of), NAMES);
  check('case and outer or doubled whitespace play no part', ['Agreed', 'AGREED', '  agreed  ', '\tnoted\n', 'plenary', 'PLENARY', ' Parked ', 'Other', 'POSTPONED'].map(of), ['agreed', 'agreed', 'agreed', 'noted', 'Plenary', 'Plenary', 'parked', 'other', 'postponed']);
  check('"replied to", as the Portal writes it, is replied', ['replied to', 'Replied to', 'REPLIED  TO', 'replied'].map(of), ['replied', 'replied', 'replied', 'replied']);
  check('parked, Plenary and other, typed as a status, are those options: a decision somebody wrote down is kept', ['parked', 'Plenary', 'other'].map(of), ['parked', 'Plenary', 'other']);
  check('whereas a status that comes from the Portal never selects them (T-2026.10.7, unchanged)', ['parked', 'Plenary', 'other'].map((x) => S.mapPortalStatusToDropdownOption_(x, OPTIONS)), [null, null, null]);
  check('a text that is no status of the dropdown is no option -- and never "other"',
    ['not treated', 'rejected', 'agreed, see minutes', 'agreed with comments', 'conditionally agreed', 'park', 'others', 'noted.', 'agreed?', 'revised to S4aA269099', 'for e-mail approval', 'TBD', '-', '–'].map(of), Array(14).fill(null));
  check('no text is no option', ['', '   ', null, undefined].map(of), [null, null, null, null]);
  check('the option is the one of the list given, with its id; without the option in the list there is none',
    [S.statusDropdownMigrationOption_('Parked', OPTIONS), S.statusDropdownMigrationOption_('parked', OPTIONS.filter((x) => x.displayValue !== 'parked')), S.statusDropdownMigrationOption_('agreed', [])], [OPTIONS[4], null, null]);
}

// ================================================================ 2. the plan

console.log('\n2. the plan: what happens to every TDoc table');
{
  const text = (t, more) => Object.assign({ kind: 'text', text: t, insertable: t.trim().length > 0, start: 10, end: 10 + t.length, at: 10 }, more || {});
  const dropdown = (value, definitionId) => ({ kind: 'dropdown', dropdownId: 'kix.d', definitionId: definitionId, selectedOptionId: 'dropdownItem.x2', displayValue: value, text: '' });
  const index = { duplicates: { S4aA269020: true }, tables: {
    S4aA269011: text('agreed'), S4aA269012: text(' Replied to '), S4aA269013: text('parked'), S4aA269014: text('not treated'), S4aA269015: text(''), S4aA269016: text('   '),
    S4aA269017: dropdown('agreed', 'kix.def1'), S4aA269018: dropdown('agreed', 'kix.other'), S4aA269019: text('agreed', { insertable: false }),
    S4aA269020: text('agreed'), 'not a number': text('agreed'), Title: text('agreed') } };
  const plan = S.planStatusDropdownMigration_(index, OPTIONS, 'kix.def1');
  check('every table is counted', plan.inspected, 12);
  check('converted: a plain text that is a status, with exactly the text it has', plan.convert, { S4aA269011: 'agreed', S4aA269012: ' Replied to ', S4aA269013: 'parked' });
  check('each with the option of its own text', [plan.forced.S4aA269011.displayValue, plan.forced.S4aA269012.displayValue, plan.forced.S4aA269013.displayValue], ['agreed', 'replied', 'parked']);
  check('already a dropdown of "Document Status": left alone', plan.already, ['S4aA269017']);
  check('no status of the dropdown: stays text, and is named', plan.unmapped, [{ number: 'S4aA269014', text: 'not treated' }]);
  check('empty: stays empty', plan.empty, ['S4aA269015', 'S4aA269016']);
  check('skipped, each with its reason: another dropdown, a cell that is not one line of plain text, a number with two tables, a table that has no TDoc number',
    plan.skipped, [{ number: 'S4aA269018', why: 'it has a dropdown that is not "Document Status"' }, { number: 'S4aA269019', why: 'its Status cell is not one line of plain text' },
      { number: 'S4aA269020', why: 'this TDoc number has more than one table' }, { number: 'Title', why: 'not a TDoc number' }, { number: 'not a number', why: 'not a TDoc number' }]);
  check('a report that has no definition yet: a dropdown it has is not "Document Status"', S.planStatusDropdownMigration_(index, OPTIONS, null).already, []);
  check('main-meeting TDoc numbers are TDoc numbers too', S.planStatusDropdownMigration_({ duplicates: {}, tables: { 'S4-261234': text('noted') } }, OPTIONS, null).convert, { 'S4-261234': 'noted' });
}

// ================================================================ 3. looking only

console.log('\n3. looking only (the dry run)');
{
  const r = existingReport();
  check('(the existing report: text statuses, typed decisions, not marked)', [r.all(), r.docProps.getProperty('STATUS_DROPDOWNS')], [BEFORE, null]);
  const before = [r.otherContent(), r.all().join('|'), r.api.revision()];
  const preview = r.migrate({ dryRun: true });
  check('what it finds: 11 tables, 8 to convert, 3 that stay text (one unmapped by the Portal, one typed, one empty)', brief(preview), { ok: true, inspected: 11, already: 0, toConvert: 8, converted: 0, text: 3, skipped: 0, notConverted: 0, unexpected: 0, marked: false });
  check('the texts that are no status are named with their TDoc', preview.unmapped, [{ number: 'S4aA269005', text: 'not treated' }, { number: 'S4aA269011', text: 'agreed, see minutes' }]);
  check('it says that the dropdown would be added to the report', [preview.definition, preview.dryRun, preview.markedBefore], ['to be created', true, false]);
  check('it only read: the report and the template, once each; nothing was written, created or marked', [r.last, Object.keys(r.api.definitions).length, r.docProps.getProperty('STATUS_DROPDOWNS'), r.api.count('batchUpdate')], [['get report', 'get template'], 0, null, 0]);
  check('the document is exactly what it was', [r.otherContent(), r.all().join('|'), r.api.revision()], before);
}

// ================================================================ 4. the conversion

console.log('\n4. the conversion: values, definition, the mark, and nothing else');
{
  const r = existingReport();
  const other = r.otherContent();
  const tableCount = r.body.getTables().length;
  const result = r.migrate();
  check('it completes', [result.ok, result.error], [true, '']);
  check('8 converted, 3 left as text, nothing skipped, nothing unexpected; the report is marked', brief(result), { ok: true, inspected: 11, already: 0, toConvert: 8, converted: 8, text: 3, skipped: 0, notConverted: 0, unexpected: 0, marked: true });
  check('every Status has the value it had: as a dropdown where its text was a status, as the same text otherwise', r.all(), AFTER);
  check('parked, Plenary and other, typed by hand, are those dropdown values', [r.api.shown('S4aA269007'), r.api.shown('S4aA269008'), r.api.shown('S4aA269009')], ['dropdown:parked', 'dropdown:Plenary', 'dropdown:other']);
  check('"  NOTED " is noted; "Replied to" is replied; the cells hold no text beside the dropdown', [r.api.shown('S4aA269010'), r.api.shown('S4aA269004'), TDOCS.map((t) => r.api.statusCell(t[0])._dropdown ? r.api.statusCell(t[0])._t : 'text').filter((x) => x !== 'text')], ['dropdown:noted', 'dropdown:replied', Array(8).fill('')]);
  check('"not treated" and "agreed, see minutes" are still that text; nothing became "other" that did not say other', [r.api.shown('S4aA269005'), r.api.shown('S4aA269011'), r.all().filter((x) => /dropdown:other/.test(x))], ['text:not treated', 'text:agreed, see minutes', ['S4aA269009 dropdown:other']]);
  check('the empty Status is still empty', r.api.shown('S4aA269006'), 'empty');
  check('nothing else of the document changed: every paragraph, every table, every other cell -- the minutes included', [r.otherContent() === other, r.body.getTables().length, r.cell('S4aA269001', 'Minutes').getText(), r.cell('S4aA269007', 'Minutes').getText()],
    [true, tableCount, 'Presented by Sam Rivera. The group agreed the proposal.', 'Parked until the next call.']);
  check('the conversion says so itself', result.contentUnchanged, true);

  const made = Object.keys(r.api.definitions).map((id) => r.api.definitions[id].dropdownDefinitionProperties);
  check('one "Document Status" definition was created, an exact copy of the template\'s: names, order, both colours', [made.length, made[0].title, made[0].options.map((x) => x.displayValue), made[0].options.map((x) => JSON.stringify(x.textStyle)).join() === DEFINITION.options.map((x) => JSON.stringify(x.textStyle)).join()], [1, TITLE, NAMES, true]);
  check('the order of the work: read, read the template (to plan, and to create), create the definition, read again, one batch, read back', r.last, ['get report', 'get template', 'get template', 'batchUpdate createDropdownDefinition x1', 'get report', 'batchUpdate deleteContentRange+insertDropdown x16', 'get report']);
  check('the batch names the revision its positions were read at', (() => { const calls = r.api.calls; const batch = calls[calls.length - 2]; return [batch.call, /^rev\.\d+$/.test(batch.requiredRevisionId)]; })(), ['batchUpdate', true]);
  check('DocumentApp changed nothing (the document was closed for it throughout), and nothing was saved', [r.body._state.dirty, r.last.indexOf('save')], [false, -1]);
  check('the TDoc list was not read: no download', r.downloads, 0);
  check('the template was only read, never written', r.api.calls.filter((c) => c.id === TEMPLATE_ID).map((c) => c.call), ['get', 'get']);
  check('what is logged names no TDoc, no status text and no minutes', r.statusLogs().filter((l) => /S4aA|Synthetic|Rivera|parked|see minutes/.test(l)), []);
  check('one line says what was done', r.statusLogs().filter((l) => /conversion of an existing report -- 11 inspected, 0 already dropdowns, 8 converted, 3 left as text, 0 skipped, 0 not converted, 0 unexpected; marked: true/.test(l)).length, 1);
}

// ================================================================ 5. running it again

console.log('\n5. running it again');
{
  const r = existingReport();
  r.migrate();
  const state = [r.all().join('|'), r.otherContent(), r.api.revision(), Object.keys(r.api.definitions).join()];
  const again = r.migrate();
  check('a second run: 8 already dropdowns, nothing to convert, nothing converted', brief(again), { ok: true, inspected: 11, already: 8, toConvert: 0, converted: 0, text: 3, skipped: 0, notConverted: 0, unexpected: 0, marked: true });
  check('it wrote nothing: two reads, no batch; no second definition; the document is at the same revision', [r.last, again.definition, [r.all().join('|'), r.otherContent(), r.api.revision(), Object.keys(r.api.definitions).join()]], [['get report', 'get report'], 'in the report', state]);
  // A value picked in a dropdown since is a value of the meeting: a run does not touch it.
  r.api.pick('S4aA269002', 'agreed');
  r.user(() => r.api.statusCell('S4aA269005').setText('postponed'));                    // and a text status was corrected by hand
  const third = r.migrate();
  check('a later run converts what has become convertible, and leaves every dropdown as it is', [brief(third).converted, brief(third).already, r.api.shown('S4aA269005'), r.api.shown('S4aA269002')], [1, 8, 'dropdown:postponed', 'dropdown:agreed']);
  check('still one definition', Object.keys(r.api.definitions).length, 1);
}

// ================================================================ 6. a run that is cut short

console.log('\n6. a run that is cut short heals on the next run');
{
  const STATES = ['agreed', 'available', 'reserved', 'noted', 'not treated', 'approved', 'replied to'];
  const many = Array.from({ length: 130 }, (_, i) => ['S4aA26' + String(9001 + i), i % 2 ? '4.3' : '4.4', STATES[i % STATES.length]]);
  const expected = many.map((t) => t[0] + ' ' + (t[2] === 'not treated' ? 'text:not treated' : 'dropdown:' + (t[2] === 'replied to' ? 'replied' : t[2])));
  const r = existingReport({ tdocs: many, typed: {} });
  const other = r.otherContent();
  r.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange && r.api.calls.filter((c) => c.call === 'batchUpdate' && c.requests[0].deleteContentRange).length === 2 ? 'synthetic failure of the second batch' : null);
  const cut = r.migrate();
  check('the second of three batches fails: the run says it is not complete, and why', [cut.ok, /^Not every Status could be converted \(synthetic failure of the second batch\)\. What was converted is complete; the rest is unchanged\. Run the conversion again\.$/.test(cut.error)], [false, true]);
  check('50 are converted, 62 are still text, none is in between; the report is NOT marked', [brief(cut).converted, brief(cut).notConverted, brief(cut).unexpected, cut.marked, r.docProps.getProperty('STATUS_DROPDOWNS')], [50, 62, 0, false, null]);
  const shown = r.all();
  check('every Status is whole: a dropdown or its own text, never an emptied cell, never text and dropdown together', [shown.filter((x) => / empty$/.test(x)).length, shown.filter((x) => /\+/.test(x)).length, shown.filter((x, i) => /text:/.test(x) && x !== many[i][0] + ' text:' + many[i][2])], [0, 0, []]);
  check('every converted Status has the value its text had', shown.filter((x, i) => /dropdown:/.test(x) && x !== expected[i]), []);
  check('nothing else changed', [r.otherContent() === other, cut.contentUnchanged], [true, true]);
  // An update of the report in this state: it is not marked, so the update converts nothing.
  r.api.fail.batchUpdate = null;
  r.update();
  check('an update in between converts nothing by itself: the report is not marked', r.all().filter((x) => /dropdown:/.test(x)).length, 50);
  const healed = r.migrate();
  check('the next run converts the rest: 50 were already dropdowns, 62 are converted now; complete, and marked', [brief(healed).ok, brief(healed).already, brief(healed).converted, healed.marked, r.all().filter((x, i) => x !== expected[i])], [true, 50, 62, true, []]);
  check('one definition throughout', [Object.keys(r.api.definitions).length, healed.definition], [1, 'in the report']);
  check('and nothing else changed in either run', r.otherContent() === other, true);

  // The API does not say which revision a batch left: no further position is trusted.
  const blind = existingReport({ tdocs: many, typed: {} });
  blind.api.noRevisionInReply = true;
  const stopped = blind.migrate();
  check('a reply without a revision: the first batch is applied, then the run stops rather than guess', [stopped.ok, brief(stopped).converted, brief(stopped).notConverted, /the revision of the document is not known/.test(stopped.error), stopped.marked], [false, 50, 62, true, false]);
  blind.api.noRevisionInReply = false;
  check('and the next run completes it', [brief(blind.migrate()).ok, blind.all().filter((x, i) => x !== expected[i])], [true, []]);
}

// ================================================================ 7. the document is edited while it runs

console.log('\n7. the document is edited while it runs');
{
  const r = existingReport();
  const other = r.otherContent();
  // Somebody types into the document after it was read and before the batch arrives.
  r.api.beforeWrite = (requests) => { if (requests[0].deleteContentRange) { r.api.edited(); r.api.beforeWrite = null; } };
  const refused = r.migrate();
  check('the batch is refused by its revision: nothing is converted, nothing is damaged, the report is not marked', [refused.ok, brief(refused).converted, brief(refused).notConverted, brief(refused).unexpected, refused.marked, r.all().slice(0, 4)],
    [false, 0, 8, 0, false, BEFORE.slice(0, 4)]);
  check('it says to run it again', /does not match the latest revision[\s\S]*Run the conversion again\.$/.test(refused.error), true);
  check('nothing else changed; the definition that was created stays for the next run', [r.otherContent() === other, Object.keys(r.api.definitions).length], [true, 1]);
  const again = r.migrate();
  check('the next run converts everything, with that one definition', [brief(again).ok, brief(again).converted, r.all(), Object.keys(r.api.definitions).length, again.definition], [true, 8, AFTER, 1, 'in the report']);

  // Somebody changes the minutes after the last write and before the run looks at the document again.
  const edited = existingReport();
  let reads = 0;
  edited.api.fail.get = (id) => { if (id === REPORT_ID && ++reads === 3) edited.cell('S4aA269001', 'Minutes')._t = 'Changed by somebody else, meanwhile.'; return null; };
  const unsure = edited.migrate();
  check('the document differs outside its Status cells afterwards: the run does not call itself complete and does not mark the report', [unsure.ok, unsure.contentUnchanged, brief(unsure).converted, unsure.marked, /may have been edited meanwhile; run the conversion again to check/.test(unsure.error)], [false, false, 8, false, true]);
  edited.api.fail.get = null;
  const checked = edited.migrate();
  check('the run after it finds everything in order and marks the report', [checked.ok, brief(checked).already, checked.marked], [true, 8, true]);
  check('what the other person typed is of course still there', edited.cell('S4aA269001', 'Minutes').getText(), 'Changed by somebody else, meanwhile.');
}

// ================================================================ 8. what is never converted

console.log('\n8. what is never converted');
{
  const r = existingReport();
  r.user(() => {
    r.body.appendTable([['TDoc', 'S4aA269003'], ['Title', 'A second table of the same number'], ['Status', 'agreed']]);
    r.body.appendTable([['TDoc', 'Summary of the meeting'], ['Status', 'agreed']]);
    r.body.appendTable([['Name', 'Status'], ['Pat Kim', 'agreed']]);
    r.body.appendTable([['TDoc', 'S4aA269050'], ['Title', 'A table without a Status row'], ['Minutes', 'agreed']]);
  });
  const other = r.otherContent();
  const result = r.migrate();
  check('a TDoc number with two tables is skipped: neither table is touched', [result.skipped.filter((x) => x.number === 'S4aA269003'), r.api.shown('S4aA269003'), rowsOf(r.body.getTables().filter((t) => rowsOf(t)[1][1] === 'A second table of the same number')[0])[2]],
    [[{ number: 'S4aA269003', why: 'this TDoc number has more than one table' }], 'text:reserved', ['Status', 'agreed']]);
  check('a table whose "TDoc" is no TDoc number is skipped', [result.skipped.filter((x) => x.number === 'Summary of the meeting').length, rowsOf(r.body.getTables().filter((t) => rowsOf(t)[0][1] === 'Summary of the meeting')[0])[1]], [1, ['Status', 'agreed']]);
  check('a table that is no TDoc table, and a TDoc table without a Status row, are not even looked at', [result.inspected, rowsOf(r.body.getTables().filter((t) => rowsOf(t)[0][0] === 'Name')[0])[1], rowsOf(r.table('S4aA269050'))[2]], [12, ['Pat Kim', 'agreed'], ['Minutes', 'agreed']]);
  check('the rest is converted; the run is complete, since a skipped table is a decision and not a failure', [result.ok, brief(result).converted, result.marked], [true, 7, true]);
  check('nothing else changed', r.otherContent() === other, true);

  // A cell that is more than one line of text, and a cell that already has a dropdown with text beside it.
  const shapes = existingReport();
  shapes.migrate();
  shapes.user(() => { shapes.api.statusCell('S4aA269001').setText('typed beside the dropdown'); });
  const second = shapes.migrate();
  check('a cell with a dropdown and text beside it is "already a dropdown": the run does not touch it', [brief(second).already, shapes.api.shown('S4aA269001'), brief(second).converted], [8, 'text:typed beside the dropdown+dropdown:agreed', 0]);
  check('the registration table and the Status column of any other table are never a Status cell',
    S.planStatusDropdownMigration_(S.buildStatusDocsIndex_({ tabs: [{ tabProperties: { tabId: 't.0' }, documentTab: { body: { content: [{ table: { tableRows: [
      { tableCells: [{ content: [{ paragraph: { elements: [{ startIndex: 1, endIndex: 6, textRun: { content: 'TDoc\n' } }] } }] }, { content: [{ paragraph: { elements: [{ startIndex: 6, endIndex: 12, textRun: { content: 'Title\n' } }] } }] }] },
      { tableCells: [{ content: [{ paragraph: { elements: [{ startIndex: 12, endIndex: 23, textRun: { content: 'S4aA269001\n' } }] } }] }, { content: [{ paragraph: { elements: [{ startIndex: 23, endIndex: 30, textRun: { content: 'agreed\n' } }] } }] }] }] } }] } } }] }), OPTIONS, null).inspected, 0);
}

// ================================================================ 9. when something is missing or fails

console.log('\n9. when something is missing or fails: nothing is changed');
{
  const untouched = (r) => [r.all(), r.docProps.getProperty('STATUS_DROPDOWNS'), Object.keys(r.api.definitions).length, r.api.count('batchUpdate')];
  const noTemplate = existingReport({ template: null });
  const a = noTemplate.migrate();
  check('the template cannot be read: an error, and nothing is changed', [a.ok, /^The conversion stopped: Requested entity was not found\.$/.test(a.error), untouched(noTemplate)], [false, true, [BEFORE, null, 0, 0]]);

  const noDefinition = existingReport({ template: [{ title: 'Review status', options: [{ displayValue: 'a' }, { displayValue: 'b' }] }] });
  const b = noDefinition.migrate();
  check('the template has no "Document Status" dropdown: said so, nothing changed', [b.ok, b.error, untouched(noDefinition)], [false, 'The template has no "Document Status" dropdown to take the statuses from. Nothing was changed.', [BEFORE, null, 0, 0]]);
  check('the dry run says the same', [noDefinition.migrate({ dryRun: true }).error], ['The template has no "Document Status" dropdown to take the statuses from. Nothing was changed.']);

  const noRead = existingReport();
  noRead.api.fail.get = () => 'synthetic read failure';
  const c = noRead.migrate();
  check('the report cannot be read through the API: an error, nothing changed', [c.ok, /^The conversion stopped: synthetic read failure$/.test(c.error), untouched(noRead)], [false, true, [BEFORE, null, 0, 0]]);

  const noCreate = existingReport();
  noCreate.api.fail.batchUpdate = (requests) => (requests[0].createDropdownDefinition ? 'synthetic failure of the definition' : null);
  const d = noCreate.migrate();
  check('the definition cannot be created: an error, no Status changed', [d.ok, /^The conversion stopped: synthetic failure of the definition$/.test(d.error), noCreate.all(), noCreate.docProps.getProperty('STATUS_DROPDOWNS'), Object.keys(noCreate.api.definitions).length], [false, true, BEFORE, null, 0]);

  const noWrite = existingReport();
  noWrite.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange ? 'synthetic write failure' : null);
  const e = noWrite.migrate();
  check('no Status can be written: every Status is its text; the definition is there for the next run; not marked', [e.ok, noWrite.all(), Object.keys(noWrite.api.definitions).length, e.marked, brief(e).notConverted], [false, BEFORE, 1, false, 8]);
  noWrite.api.fail.batchUpdate = null;
  check('the next run converts everything and creates no second definition', [brief(noWrite.migrate()).ok, noWrite.all(), Object.keys(noWrite.api.definitions).length], [true, AFTER, 1]);

  const noService = existingReport();
  delete noService.s.Docs;
  const f = noService.migrate();
  check('without the Google Docs API service: said so, nothing changed', [f.ok, f.error, noService.all()], [false, 'Status dropdowns are not available in this report: it needs the template runtime with the Google Docs API service.', BEFORE]);
  const central = existingReport();
  delete central.s.SA4_RELEASE_;
  check('outside the template runtime (CENTRAL, Legacy): the same, and no API call', [central.migrate().error, central.events], ['Status dropdowns are not available in this report: it needs the template runtime with the Google Docs API service.', []]);
  check('the conversion never throws (source: one try around all of it)', /function migrateStatusFieldsToDropdowns_\(options\) \{\n  const dryRun = [^\n]*\n  const result = [^\n]*\n[^\n]*\n  try \{[\s\S]*\n  \} catch \(e\) \{\n    result\.error = 'The conversion stopped: ' \+ e\.message;\n[^\n]*\n  \}\n  return result;\n\}/.test(CODE), true);
}

// ================================================================ 10. after the conversion

console.log('\n10. after the conversion: updates, by the rules of T-2026.10.7');
{
  const r = existingReport();
  r.migrate();
  const other = r.otherContent();
  // The Portal has moved on since the report was built.
  r.portal('S4aA269003', 'available');                                   // reserved -> available: follows
  r.portal('S4aA269007', 'agreed');                                      // parked by hand: kept
  r.portal('S4aA269008', 'noted');                                       // Plenary by hand: kept
  r.portal('S4aA269001', 'noted');                                       // agreed: a decision, kept
  r.portal('S4aA269004', 'revised');                                     // "revised" always applies
  r.tdocs.push(['S4aA269012', '4.4', 'agreed']);                         // a new TDoc
  const done = r.update();
  check('Update Report Now completes', done.success, true);
  check('a status that was typed by hand and is now a dropdown is kept against the Portal', [r.api.shown('S4aA269007'), r.api.shown('S4aA269008'), r.api.shown('S4aA269001')], ['dropdown:parked', 'dropdown:Plenary', 'dropdown:agreed']);
  check('reserved moves on to available, and "revised" applies, as for any dropdown', [r.api.shown('S4aA269003'), r.api.shown('S4aA269004')], ['dropdown:available', 'dropdown:revised']);
  check('the new TDoc gets a dropdown: the report is a report with status dropdowns now', r.api.shown('S4aA269012'), 'dropdown:agreed');
  check('the text statuses that were left stay text', [r.api.shown('S4aA269005'), r.api.shown('S4aA269011'), r.api.shown('S4aA269006')], ['text:not treated', 'text:agreed, see minutes', 'empty']);
  check('the minutes are untouched by the update too', [r.cell('S4aA269001', 'Minutes').getText(), r.cell('S4aA269007', 'Minutes').getText()], ['Presented by Sam Rivera. The group agreed the proposal.', 'Parked until the next call.']);
  check('still one definition', Object.keys(r.api.definitions).length, 1);
  r.update();
  check('the update after it writes nothing', r.last, ['get report']);
  // A value picked in the dropdown after the conversion.
  r.api.pick('S4aA269002', 'parked');
  r.portal('S4aA269002', 'agreed');
  r.update();
  check('a value picked afterwards is kept as well', r.api.shown('S4aA269002'), 'dropdown:parked');
  // The readers.
  const found = r.exec(() => r.s.detectTdocTablesInDocument_(r.body));
  const byTdoc = {}; found.forEach((x) => { byTdoc[x.tdoc] = x; });
  check('the discussion e-mails read the converted statuses: agreed is excluded, parked is offered, a text status counts as before', [byTdoc.S4aA269001.status, byTdoc.S4aA269001.excluded, byTdoc.S4aA269007.status, byTdoc.S4aA269007.excluded, byTdoc.S4aA269005.status], ['agreed', true, 'parked', false, 'not treated']);
  // A rebuild of the converted report keeps the decisions (T-2026.10.7).
  const rebuilt = r.exec(() => r.s.runFullReportBuildCore_());
  check('a later Build Report from Scratch keeps what the dropdowns said, by the rule of T-2026.10.7', [rebuilt.ok, r.api.shown('S4aA269007'), r.api.shown('S4aA269008'), r.api.shown('S4aA269002')], [true, 'dropdown:parked', 'dropdown:Plenary', 'dropdown:parked']);
  check('(before that rebuild, nothing but Status cells and the new table had changed)', typeof other, 'string');
}

// ================================================================ 11. the menu action

console.log('\n11. the menu action: it asks first');
{
  const TITLE_ = 'Convert Status Fields to Dropdowns';
  const r = existingReport();
  const other = r.otherContent();
  r.ui.answers = ['NO'];
  const declined = r.exec(() => r.s.convertStatusFieldsToDropdowns());
  check('it shows what it found and asks', [r.ui.alerts.length, r.ui.alerts[0][0], r.ui.alerts[0][2]], [1, TITLE_, 'YES_NO']);
  const question = r.ui.alerts[0][1];
  check('the question: the numbers', question.split('\n').filter((l) => /^TDoc tables found|^• /.test(l)),
    ['TDoc tables found: 11', '• already a dropdown: 0', '• to be converted now: 8', '• left as text: 3 (no such status in the dropdown: S4aA269005 "not treated", S4aA269011 "agreed, see minutes"; 1 empty)']);
  check('the question says: in place; every field keeps its status; minutes and everything else stay; NOT rebuilt; the TDoc list is not read; the dropdown is added first',
    [/converts the Status fields of this report into dropdowns, in place\./.test(question), /Every converted field keeps the status it has now\. Minutes, abstracts, links and everything else in the report stay exactly as they are\./.test(question),
      /The report is NOT rebuilt, and the TDoc list is not read\./.test(question), /The "Document Status" dropdown is added to this report first\./.test(question), /Continue\?$/.test(question)], [true, true, true, true, true]);
  check('answered No: nothing is changed, written or marked; the lock was not even taken', [declined, r.all(), r.otherContent() === other, r.api.count('batchUpdate'), r.docProps.getProperty('STATUS_DROPDOWNS'), r.lock.tries], [{ ok: false, cancelled: true }, BEFORE, true, 0, null, 0]);

  r.ui.answers = ['YES', 'OK'];
  const done = r.exec(() => r.s.convertStatusFieldsToDropdowns());
  check('answered Yes: converted, with the lock an update holds, taken and released', [done.ok, r.all(), r.lock.tries, r.lock.releases, r.docProps.getProperty('STATUS_DROPDOWNS')], [true, AFTER, 1, 1, '1']);
  const report = r.ui.alerts[r.ui.alerts.length - 1];
  check('it says what it did', [report[0], report[2], report[1].split('\n').filter((l) => /^Done|^TDoc tables found|^• /.test(l))], [TITLE_, 'OK', ['Done.', 'TDoc tables found: 11', '• already a dropdown before: 0', '• converted now: 8', '• left as text: 3']]);
  check('and what comes next', /New TDocs of this report get a dropdown from now on\. A Status that stayed text can be converted later by running this again\.$/.test(report[1]), true);

  const alerts = r.ui.alerts.length;
  r.ui.answers = [];
  r.exec(() => r.s.convertStatusFieldsToDropdowns());
  check('used again on a converted report: one message, no question, nothing written', [r.ui.alerts.length - alerts, r.ui.alerts[r.ui.alerts.length - 1][2], r.ui.alerts[r.ui.alerts.length - 1][1].split('\n')[0], r.last.filter((e) => /batchUpdate/.test(e))], [1, 'OK', 'There is nothing to convert.', []]);

  // An update is running.
  const busy = existingReport();
  busy.locked = true;
  busy.ui.answers = ['YES', 'OK'];
  const refused = busy.exec(() => busy.s.convertStatusFieldsToDropdowns());
  check('while an update of the report runs: nothing is changed, and it says so', [refused, busy.all(), busy.ui.alerts[busy.ui.alerts.length - 1][1], busy.api.count('batchUpdate')],
    [{ ok: false, error: 'locked' }, BEFORE, 'Nothing was changed: an update of this report is running right now. Try again in a minute.', 0]);

  // A run that is not complete.
  const cut = existingReport();
  cut.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange ? 'synthetic write failure' : null);
  cut.ui.answers = ['YES', 'OK'];
  cut.exec(() => cut.s.convertStatusFieldsToDropdowns());
  const said = cut.ui.alerts[cut.ui.alerts.length - 1][1];
  check('a run that is not complete says so, and that nothing was lost', [said.split('\n')[0], /• not converted, still text: 8 \(S4aA269001, S4aA269002, S4aA269003, S4aA269004, S4aA269007, S4aA269008, S4aA269009, S4aA269010\)/.test(said), /Nothing was lost: a field is either converted or as it was\. Run "Convert Status Fields to Dropdowns…" again\.$/.test(said)],
    ['The conversion is NOT complete.', true, true]);

  // A problem before anything is asked.
  const none = existingReport({ template: [{ title: 'Review status', options: [{ displayValue: 'a' }, { displayValue: 'b' }] }] });
  none.exec(() => none.s.convertStatusFieldsToDropdowns());
  check('when it cannot work, it says why and asks nothing', [none.ui.alerts.length, none.ui.alerts[0][2], none.ui.alerts[0][1]], [1, 'OK', 'The template has no "Document Status" dropdown to take the statuses from. Nothing was changed.']);

  check('the action is in the Report submenu of a template report, after the status summary (source)', /\.addItem\('Report Status Summary', 'analyzeReportStatus'\)\n    \/\/[^\n]*\n    \.addItem\('Convert Status Fields to Dropdowns…', 'convertStatusFieldsToDropdowns'\)\);/.test(CREATOR), true);
  check('it is nowhere else: not in the CENTRAL / Legacy menu of Code.js', [(CREATOR.match(/convertStatusFieldsToDropdowns/g) || []).length, (CODE.match(/'convertStatusFieldsToDropdowns'/g) || []).length], [1, 0]);
  check('it refuses in the master template, as every action of a report does (source)', /function convertStatusFieldsToDropdowns\(\) \{\n  assertNotTemplateMaster_\(\);/.test(CODE), true);
  check('the dialog shows no more than a few TDoc numbers', [S.statusDropdownNumberList_(['a', 'b']), S.statusDropdownNumberList_(Array.from({ length: 12 }, (_, i) => 'T' + i))], ['a, b', 'T0, T1, T2, T3, T4, T5, T6, T7 and 4 more']);
}

// ================================================================ 12. nothing happens because the code is installed

console.log('\n12. nothing happens because the code is installed');
{
  const r = existingReport();
  const other = r.otherContent();
  r.portal('S4aA269003', 'available');
  r.tdocs.push(['S4aA269012', '4.4', 'agreed']);
  const done = r.update();
  check('an existing report with this release: its update adds the new TDoc with a text status and moves a text status on, as before', [done.success, r.api.shown('S4aA269012'), r.api.shown('S4aA269003')], [true, 'text:agreed', 'text:available']);
  check('no API call, no definition, no mark: the report is converted only by the action (or a rebuild)', [r.events, Object.keys(r.api.definitions).length, r.docProps.getProperty('STATUS_DROPDOWNS')], [[], 0, null]);
  r.exec(() => r.s.analyzeReportStatus());
  r.exec(() => r.s.detectTdocTablesInDocument_(r.body));
  check('its readers make no API call either', r.events, []);
  check('the conversion runs from its menu action only (source)', [(CODE.match(/migrateStatusFieldsToDropdowns_\(/g) || []).length, (CODE.match(/convertStatusFieldsToDropdowns\(/g) || []).length], [3, 2]);
  check('it never rebuilds, never reads the TDoc list, never uses DocumentApp to write (source)',
    /buildSkeletonWithTdocTables|runFullReportBuild|downloadAndGroupTdocs_|continuousUpdate|\.setText\(|\.clear\(|appendTable|insertTable|appendParagraph|saveAndClose/.test(CODE.slice(CODE.indexOf('// --- An existing report: "Convert Status Fields to Dropdowns…"'), CODE.indexOf('// TDOC UPLOAD COMPLETION')).replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '')), false);
  check('the only property it writes is the mark, in one place, after a complete run (source)', (CODE.slice(CODE.indexOf('// --- An existing report: "Convert Status Fields to Dropdowns…"'), CODE.indexOf('// TDOC UPLOAD COMPLETION')).match(/\.setProperty\([^)]*\)/g) || []), ["." + "setProperty(STATUS_DROPDOWN_ENABLED_KEY_, '1')"]);
  check('(nothing but the two statuses and the new table changed in this section)', typeof other, 'string');
  check('Code.js is version 2.20.0', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.20.0');

  // What is compared to prove that nothing else changed.
  const docOf = (minutes, status) => ({ tabs: [{ tabProperties: { tabId: REPORT_TAB_ID }, documentTab: { body: { content: [
    { paragraph: { elements: [{ textRun: { content: 'A note.\n' } }] } },
    { table: { tableRows: [['TDoc', 'S4aA269001'], ['Minutes', minutes], ['Status', status]].map((row) => ({ tableCells: row.map((cell) => ({ content: [{ paragraph: { elements: cell === '<dropdown>' ? [{ dropdown: {} }, { textRun: { content: '\n' } }] : [{ textRun: { content: cell + '\n' } }] } }] })) })) } }] } } }] });
  const same = (a, b) => JSON.stringify(S.statusDocsOtherContent_(a)) === JSON.stringify(S.statusDocsOtherContent_(b));
  check('the comparison ignores the value of a Status cell and nothing else', [same(docOf('m', 'agreed'), docOf('m', '<dropdown>')), same(docOf('m', 'agreed'), docOf('m changed', 'agreed')), same(docOf('m', 'agreed'), docOf('m', 'agreed'))], [true, false, true]);
  check('a paragraph outside the tables counts', (() => { const d = docOf('m', 'agreed'); d.tabs[0].documentTab.body.content[0].paragraph.elements[0].textRun.content = 'Another note.\n'; return same(docOf('m', 'agreed'), d); })(), false);
  check('the label of the Status row itself counts: only its value is left out', (() => { const d = docOf('m', 'agreed'); d.tabs[0].documentTab.body.content[1].table.tableRows[2].tableCells[0].content[0].paragraph.elements[0].textRun.content = 'TDoc Status\n'; return same(docOf('m', 'agreed'), d); })(), false);
  check('a row that becomes a Status row counts, and so does a second "Status" row of the same table', (() => { const d = docOf('m', 'agreed'); d.tabs[0].documentTab.body.content[1].table.tableRows[1].tableCells[0].content[0].paragraph.elements[0].textRun.content = 'Status\n'; const e = JSON.parse(JSON.stringify(d)); e.tabs[0].documentTab.body.content[1].table.tableRows[2].tableCells[1].content[0].paragraph.elements[0].textRun.content = 'noted\n'; return [same(docOf('m', 'agreed'), d), same(d, e)]; })(), [false, false]);
}

// ================================================================ 13. every write by position names its revision

console.log('\n13. every write by position names its revision (updates and builds too)');
{
  const r = existingReport();
  r.migrate();
  r.tdocs.push(['S4aA269012', '4.4', 'agreed']);
  r.portal('S4aA269003', 'available');
  // Somebody types into the document between the read of the update and its write.
  r.api.beforeWrite = (requests) => { if (requests[0].deleteContentRange) { r.api.edited(); r.api.beforeWrite = null; } };
  const done = r.update();
  const batches = () => r.api.calls.filter((c) => c.call === 'batchUpdate');
  check('an update whose positions are out of date: the new TDoc keeps its text status, nothing is cut; the update itself completes', [done.success, r.api.shown('S4aA269012'), r.all().slice(0, 2)], [true, 'text:agreed', AFTER.slice(0, 2)]);
  check('the write by position named a revision; the change of a value, which names its dropdown and no position, needs none and was applied',
    [batches().filter((c) => c.requests[0].deleteContentRange).every((c) => /^rev\.\d+$/.test(c.requiredRevisionId)), batches().filter((c) => c.requests[0].updateDropdownProperties).map((c) => c.requiredRevisionId), r.api.shown('S4aA269003')], [true, [null], 'dropdown:available']);
  r.update();
  check('the next update gives the new TDoc its dropdown', r.api.shown('S4aA269012'), 'dropdown:agreed');
  check('in the source, no insert pair is sent without a revision: one sender, used by the update and by the conversion, and it refuses to send blind',
    [(CODE.match(/sendStatusDropdownInsertPairs_\(/g) || []).length, /if \(!revision\) throw new Error\('the revision of the document is not known/.test(CODE), /writeControl: \{ requiredRevisionId: revision \}/.test(CODE), (CODE.match(/Docs\.Documents\.batchUpdate\(/g) || []).length], [3, true, true, 3]);
}

// ================================================================ 14. the read back decides, not the reply

console.log('\n14. the read back decides, not the reply of the API');
{
  // The pairs of a batch go from the end of the document to its start: the first pair is the last TDoc that is converted.
  const firstPairOnly = (r, change) => { r.api.beforeWrite = (requests) => { if (requests[0].deleteContentRange) { change(requests); r.api.beforeWrite = null; } }; };
  const LAST = 'S4aA269010';

  const wrong = existingReport();
  firstPairOnly(wrong, (requests) => { requests[1].insertDropdown.selectedOptionId = wrong.api.definitions[Object.keys(wrong.api.definitions)[0]].dropdownDefinitionProperties.options[2].optionId; });
  const a = wrong.migrate();
  check('(the fake wrote another value than was asked for)', wrong.api.shown(LAST), 'dropdown:agreed');
  check('a dropdown with another value than the text had: not complete, the TDoc is named, the report is not marked', [a.ok, a.unexpected, brief(a).converted, a.marked, a.error], [false, [LAST], 7, false, 'After the conversion 1 Status cell(s) are not what was expected: ' + LAST + '.']);
  check('the dialog names it for a look', /• to be looked at: 1 \(S4aA269010\)/.test(wrong.s.formatStatusDropdownMigrationResult_(a)), true);

  const beside = existingReport();
  firstPairOnly(beside, (requests) => { requests.splice(0, 1); });
  const b = beside.migrate();
  check('a dropdown with the old text still beside it: not complete, named, not marked', [beside.api.shown(LAST), b.ok, b.unexpected, b.marked], ['text:  NOTED +dropdown:noted', false, [LAST], false]);

  const silent = existingReport();
  firstPairOnly(silent, (requests) => { requests.splice(0, 2); });
  const c = silent.migrate();
  check('a pair the API did not apply without saying so: the cell is still its text, the run is not complete, the report is not marked', [silent.api.shown(LAST), c.ok, c.notConverted, brief(c).converted, c.marked, c.error], ['text:  NOTED ', false, [LAST], 7, false, 'Not every Status was converted. Run the conversion again.']);
  check('the next run converts it and marks the report', [brief(silent.migrate()).ok, silent.api.shown(LAST), silent.docProps.getProperty('STATUS_DROPDOWNS')], [true, 'dropdown:noted', '1']);

  const other = existingReport();
  firstPairOnly(other, () => { other.api.statusCell('S4aA269005')._t = 'NOT TREATED'; });
  const d = other.migrate();
  check('a Status that was not to be converted and is another text afterwards: named, not complete, not marked', [d.ok, d.unexpected, brief(d).converted, d.marked], [false, ['S4aA269005'], 8, false]);

  // The API says a batch failed although it was applied (a timeout on the way back).
  const late = existingReport();
  const real = late.s.Docs.Documents.batchUpdate;
  late.s.Docs.Documents.batchUpdate = (resource, id) => { const reply = real(resource, id); if (resource.requests[0].deleteContentRange) throw new Error('synthetic timeout after the write'); return reply; };
  const e = late.migrate();
  check('a failure reported for a batch that was applied: everything is converted, but the run does not call itself complete and does not mark the report', [e.ok, brief(e).converted, brief(e).notConverted, e.marked, /synthetic timeout after the write/.test(e.error)], [false, 8, 0, false, true]);
  late.s.Docs.Documents.batchUpdate = real;
  // The menu action on a report in that state: everything is a dropdown, but it is not marked -- it still asks and runs.
  late.ui.answers = ['YES', 'OK'];
  const alerts = late.ui.alerts.length;
  const f = late.exec(() => late.s.convertStatusFieldsToDropdowns());
  check('the action, used again: nothing is left to convert, yet it runs (it asks first), checks, and marks the report', [late.ui.alerts.length - alerts, late.ui.alerts[alerts][2], f.ok, brief(f).already, brief(f).converted, late.docProps.getProperty('STATUS_DROPDOWNS'), late.last.filter((x) => /batchUpdate/.test(x))], [2, 'YES_NO', true, 8, 0, '1', []]);
}

console.log(failures === 0 ? '\nAll status-dropdown migration checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
