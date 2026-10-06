/**
 * The Status of a TDoc table as a native Google Docs dropdown (T-2026.10.7).
 *
 * DocumentApp cannot create, read or set a dropdown; the Google Docs API
 * can. The production code therefore writes the Status as text, as every
 * release did, and at the very end of a build or an update -- after the
 * last change DocumentApp makes -- saves the document, reads it once through
 * the API and puts a dropdown in place of the text of each new TDoc whose
 * Portal status has an option. The definition of the dropdown (names, order
 * and the two colours of the options) is read from the master template.
 *
 * These tests run the real build, the real update, the real readers and the
 * real status sync against the fake document and a fake Docs API
 * (tests/helpers/fake-docs-api.js); no live service is used.
 *
 *    1. the mapping of a Portal status to an option;
 *    2. the definition: an exact copy; the report's own one is reused;
 *    3. reading a document of the API;
 *    4. Build Report from Scratch;
 *    5. the order of the work: save, one read, then the writes; nothing after;
 *    6. Update Report Now: new TDocs, revision tables;
 *    7. the readers: status summary and discussion e-mails;
 *    8. the Portal sync: what moves on, what is kept;
 *    9. many TDocs: positions and batches; tables that were moved;
 *   10. when something is missing or fails: text statuses, as before;
 *   11. reports with text statuses: unchanged, and no API call;
 *   12. a dropdown that could not be made is made by a later update;
 *   13. Build Report from Scratch keeps what was selected in a dropdown.
 *
 * All names, numbers and colours are synthetic.
 *
 * Run: node tests/status-dropdown.test.js
 */

const fs = require('fs');
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
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const OTHER_ID = 'OTHERdoc00000000000000000000000000000000000';
const TITLE = 'Document Status';
const NAMES = ['available', 'noted', 'agreed', 'revised', 'parked', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'other', 'replied', 'Plenary', 'postponed'];
/** Synthetic colours, different for every option, foreground and background. */
const colour = (n, shift) => ({ color: { rgbColor: { red: ((n * 17 + shift) % 100) / 100, green: ((n * 31 + shift) % 100) / 100, blue: ((n * 53 + shift) % 100) / 100 } } });
const DEFINITION = { title: TITLE, options: NAMES.map((name, n) => ({ displayValue: name, textStyle: { foregroundColor: colour(n, 3), backgroundColor: colour(n, 47) } })) };
const optionsWithIds = (definition) => definition.options.map((o, k) => Object.assign({ optionId: 'dropdownItem.x' + k }, o));

const PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Topic'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
/** id, agenda item, Portal status */
const TDOCS = [['S4aA269001', '4.3', 'agreed'], ['S4aA269002', '4.3', 'available'], ['S4aA269003', '4.4', 'reserved'], ['S4aA269004', '4.4', 'Replied to'], ['S4aA269005', '4.4', 'not treated'],
  ['S4aA269006', '4.4', ''], ['S4aA269007', '4.4', 'other'], ['S4aA269008', '4.4', '  NOTED '], ['S4aA269009', '1', 'approved']];

const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));

/**
 * One report with the real build, update, readers and sync.
 *   options.docs      false: no Google Docs API service in the runtime
 *   options.release   false: not a template report (CENTRAL / Legacy)
 *   options.template  the definitions of the master template (default: the canonical one); null: the template cannot be read
 *   options.props     more document properties
 */
function report(options) {
  const o = options || {};
  const loaded = loadCode({ documentProperties: Object.assign({}, PROPS, o.props || {}) });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const api = makeFakeDocsApi({ body: body, reportId: REPORT_ID });
  if (o.template !== null) api.addSource(TEMPLATE_ID, o.template || [DEFINITION]);
  const r = { s: s, body: body, api: api, docProps: loaded.docProps, tdocs: (o.tdocs || TDOCS).map((t) => t.slice()), logs: [], events: [], saveFails: false };
  const ui = { alerts: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' }, alert: (...args) => { ui.alerts.push(args); return 'YES'; }, showModalDialog: () => {} };
  r.ui = ui;
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.ElementType.TEXT = 'TEXT';
  s.DocumentApp.ElementType.UNSUPPORTED = 'UNSUPPORTED';
  const doc = { getId: () => REPORT_ID, getBody: () => body,
    saveAndClose: () => { if (r.saveFails) throw new Error('synthetic save failure'); r.events.push('save'); body._state.dirty = false; body._state.closed = true; } };
  s.DocumentApp.getActiveDocument = () => doc;
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.Session = { getScriptTimeZone: () => 'Europe/Berlin' };
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  if (o.release !== false) s.SA4_RELEASE_ = { releaseId: 'T-2026.10.7', flavor: 'template', codeVersion: '2.19.0', templateDocumentId: TEMPLATE_ID };
  if (o.docs !== false) {
    // The calls are also listed in the order of everything else that happens to the document.
    s.Docs = { Documents: { get: (id, opts) => { r.events.push('get ' + (id === REPORT_ID ? 'report' : id === TEMPLATE_ID ? 'template' : 'other')); return api.Docs.Documents.get(id, opts); },
      batchUpdate: (resource, id) => { r.events.push('batchUpdate ' + resource.requests.map((q) => Object.keys(q)[0]).filter((k, i, l) => l.indexOf(k) === i).join('+') + ' x' + resource.requests.length); return api.Docs.Documents.batchUpdate(resource, id); } } };
  }
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.removeRowHeightAndSpacing = () => {};
  s.downloadMeetingAgenda_ = () => AGENDA;
  s.collectEmailDiscussionCore_ = () => {};
  s.collectRevisionsCore_ = () => {};
  s.collectorUpdate_ = () => ({ failures: [] });
  s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
  s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) }) };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => {
    const values = [HEADER].concat(r.tdocs.map(([id, agenda, status]) => [id, 'Synthetic title of ' + id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', agenda, 'Synthetic item', status, '', '', '', '']));
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => values.map((row) => row.map(() => null)) }),
      getRange: (row, col, rows) => ({ getDisplayValues: () => [['Uploaded']].concat(r.tdocs.map(() => [''])).slice(row - 1, row - 1 + rows) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };

  /** One execution: the document is open at its start; what happens in it is listed in r.events. */
  r.exec = (fn) => { body._state.closed = false; s.resetStatusDropdownRun_(); const from = r.events.length; const out = fn(); r.last = r.events.slice(from); return out; };
  /** Somebody edits the document by hand, between two executions; the edit is saved. */
  r.user = (fn) => { body._state.closed = false; const out = fn(); body._state.dirty = false; return out; };
  r.build = () => r.exec(() => s.runFullReportBuildCore_());
  r.update = () => r.exec(() => s.continuousUpdateCore_());
  r.shown = (number) => api.shown(number);
  r.all = () => r.tdocs.map((t) => t[0] + ' ' + api.shown(t[0]));
  r.portal = (number, status) => { r.tdocs.filter((t) => t[0] === number)[0][2] = status; };
  r.table = (number) => body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(1).getText() === number)[0];
  r.statusLogs = () => r.logs.filter((l) => /^Status dropdowns/.test(l));
  return r;
}

const S = loadCode().sandbox;

// ================================================================ 1. the mapping

console.log('1. the mapping of a Portal status to an option');
{
  const OPTIONS = optionsWithIds(DEFINITION);
  const map = (status) => { const o = S.mapPortalStatusToDropdownOption_(status, OPTIONS); return o ? o.displayValue : null; };
  check('the ten statuses the Portal sets select the option of their name',
    ['available', 'noted', 'agreed', 'revised', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'postponed'].map(map), ['available', 'noted', 'agreed', 'revised', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'postponed']);
  check('"replied" and "replied to" select replied', ['replied', 'replied to', 'Replied to', 'REPLIED  TO'].map(map), ['replied', 'replied', 'replied', 'replied']);
  check('case and outer or doubled whitespace play no part', ['Agreed', 'AGREED', '  agreed  ', '\tnoted\n', 'Approved'].map(map), ['agreed', 'agreed', 'agreed', 'noted', 'approved']);
  check('parked, Plenary and other are never selected, however they are written', ['parked', 'Parked', 'Plenary', 'plenary', 'PLENARY', 'other', 'Other'].map(map), [null, null, null, null, null, null, null]);
  check('a status without an option selects nothing -- never "other"', ['not treated', 'rejected', 'conditionally agreed', 'agreed with comments', 'revised to S4aA269099', 'for e-mail approval', 'agree', 'agreedx'].map(map), [null, null, null, null, null, null, null, null]);
  check('no status selects nothing', ['', '   ', null, undefined].map(map), [null, null, null, null]);
  check('the option returned is the one of the list given, with its id', S.mapPortalStatusToDropdownOption_('Agreed', OPTIONS), OPTIONS[2]);
  check('a list that lacks the option: nothing; no list: nothing', [S.mapPortalStatusToDropdownOption_('agreed', OPTIONS.filter((x) => x.displayValue !== 'agreed')), S.mapPortalStatusToDropdownOption_('agreed', []), S.mapPortalStatusToDropdownOption_('agreed', null)], [null, null, null]);
  check('an option the definition names in capitals is found by a status in small letters, and the other way round',
    [S.mapPortalStatusToDropdownOption_('postponed', [{ optionId: 'dropdownItem.a1', displayValue: 'Postponed' }]).optionId, S.mapPortalStatusToDropdownOption_('POSTPONED', [{ optionId: 'dropdownItem.a2', displayValue: 'postponed' }]).optionId], ['dropdownItem.a1', 'dropdownItem.a2']);
  check('the aliases and the manual-only options are these', [S.statusDropdownLabel_('  Replied   To '), (CODE.match(/const STATUS_DROPDOWN_ALIASES_ = (.*);/) || [])[1], (CODE.match(/const STATUS_DROPDOWN_MANUAL_ONLY_ = (.*);/) || [])[1]],
    ['replied to', "{ 'replied to': 'replied' }", "['parked', 'plenary', 'other']"]);
}

// ================================================================ 2. the definition

console.log('\n2. the definition: an exact copy of the template\'s; the report\'s own one is reused');
{
  const request = S.statusDropdownDefinitionRequest_({ title: TITLE, options: optionsWithIds(DEFINITION) }, REPORT_TAB_ID);
  const sent = request.createDropdownDefinition;
  check('the request: the title, and for every option, in order, its name and its two colours', [sent.tabId, sent.dropdownDefinition.dropdownDefinitionProperties.title, sent.dropdownDefinition.dropdownDefinitionProperties.options], [REPORT_TAB_ID, TITLE, DEFINITION.options]);
  check('nothing else of an option is sent: no id of the template, no other text style',
    [S.statusDropdownDefinitionRequest_({ title: TITLE, options: [{ optionId: 'dropdownItem.s0o0', displayValue: 'a', textStyle: { bold: true, fontSize: { magnitude: 9 }, foregroundColor: colour(1, 1) } }, { displayValue: 'b' }] }).createDropdownDefinition],
    [{ dropdownDefinition: { dropdownDefinitionProperties: { title: TITLE, options: [{ displayValue: 'a', textStyle: { foregroundColor: colour(1, 1) } }, { displayValue: 'b' }] } } }]);
  check('no colour is written into the code: the source of this feature names no colour value',
    /rgbColor|#[0-9A-Fa-f]{6}\b|\b(red|green|blue):/.test(CODE.slice(CODE.indexOf('// TDOC STATUS DROPDOWNS (T-2026.10.7)'), CODE.indexOf('// TDOC UPLOAD COMPLETION'))), false);
  check('nor an option name, but for the alias, the three that are never selected, and the three words of the sync rule (reserved, available, revised)',
    NAMES.filter((n) => new RegExp("'" + n + "'", 'i').test(CODE.slice(CODE.indexOf('const STATUS_DROPDOWN_TITLE_'), CODE.indexOf('// TDOC UPLOAD COMPLETION')).replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, ''))), ['available', 'revised', 'parked', 'reserved', 'other', 'replied', 'Plenary']);

  const r = report();
  r.build();
  const made = Object.keys(r.api.definitions).map((id) => r.api.definitions[id].dropdownDefinitionProperties);
  check('a build creates one definition in the report', made.length, 1);
  check('its title, option names and their order are the template\'s', [made[0].title, made[0].options.map((x) => x.displayValue)], [TITLE, NAMES]);
  check('every option has the template\'s foreground and background colour', made[0].options.map((x) => JSON.stringify(x.textStyle)), DEFINITION.options.map((x) => JSON.stringify(x.textStyle)));
  check('the option ids are the report\'s own', made[0].options.every((x) => /^dropdownItem\.d1o\d+$/.test(x.optionId)), true);
  check('the template was read once, with a field mask, and never written', [r.api.calls.filter((c) => c.id === TEMPLATE_ID).map((c) => c.call + ' ' + c.fields)], [['get tabs(documentTab(dropdownDefinitions),childTabs(documentTab(dropdownDefinitions)))']]);

  // A second build: the report has its definition.
  const calls = r.api.calls.length;
  r.build();
  check('a rebuild reuses the report\'s definition: none is created, the template is not read', [Object.keys(r.api.definitions).length, r.api.calls.slice(calls).filter((c) => c.id === TEMPLATE_ID).length, r.last.filter((e) => /createDropdownDefinition/.test(e))], [1, 0, []]);
  check('and the statuses are dropdowns of it again', [r.shown('S4aA269001'), r.api.statusCell('S4aA269001')._dropdown.definitionId], ['dropdown:agreed', Object.keys(r.api.definitions)[0]]);

  // The template changes later: an existing report keeps the definition it has.
  r.api.addSource(TEMPLATE_ID, [{ title: TITLE, options: [{ displayValue: 'agreed' }, { displayValue: 'brand new' }, { displayValue: 'noted' }] }]);
  r.build();
  check('a later change of the template does not reach a report that has its definition', [Object.keys(r.api.definitions).length, r.api.definitions[Object.keys(r.api.definitions)[0]].dropdownDefinitionProperties.options.map((x) => x.displayValue)], [1, NAMES]);

  // Which definition of the template.
  const pick = (list, inChild) => { const x = report({ template: list }); if (inChild) x.api.addSource(TEMPLATE_ID, list, true); x.build(); const d = Object.keys(x.api.definitions).map((id) => x.api.definitions[id].dropdownDefinitionProperties); return d.length ? d[0].options.map((q) => q.displayValue).join(',') : 'none'; };
  check('the template\'s definition is found by its title', pick([{ title: 'Review status', options: [{ displayValue: 'x' }, { displayValue: 'y' }] }, DEFINITION]), NAMES.join(','));
  check('another title is not taken', pick([{ title: 'Review status', options: [{ displayValue: 'agreed' }, { displayValue: 'noted' }] }, { title: 'document status', options: [{ displayValue: 'agreed' }, { displayValue: 'noted' }] }]), 'none');
  check('among several of the title: the one with the most options', pick([{ title: TITLE, options: [{ displayValue: 'agreed' }, { displayValue: 'noted' }] }, DEFINITION, { title: TITLE, options: [{ displayValue: 'agreed' }, { displayValue: 'noted' }, { displayValue: 'revised' }] }]), NAMES.join(','));
  check('a definition in a child tab of the template is found', pick([DEFINITION], true), NAMES.join(','));

  // Another source document, named by the report.
  const other = report({ props: { STATUS_DROPDOWN_SOURCE_DOC_ID: OTHER_ID } });
  other.api.addSource(OTHER_ID, [{ title: TITLE, options: [{ displayValue: 'agreed' }, { displayValue: 'available' }, { displayValue: 'reserved' }] }]);
  other.build();
  check('a report that names another source document gets that document\'s definition; the template is not read', [other.api.definitions[Object.keys(other.api.definitions)[0]].dropdownDefinitionProperties.options.map((x) => x.displayValue), other.api.count('get', TEMPLATE_ID), other.api.count('get', OTHER_ID)], [['agreed', 'available', 'reserved'], 0, 1]);
}

// ================================================================ 3. reading a document of the API

console.log('\n3. reading a document of the API');
{
  const r = report();
  r.build();
  r.user(() => {
    r.api.statusCell('S4aA269002').setText('typed beside');          // text beside a dropdown, as setText() leaves it
    r.body.appendTable([['TDoc', 'S4aA269001'], ['Status', 'second table of the same number']]);
    r.body.appendTable([['TDoc', 'S4aA269050'], ['Title', 'no status row']]);
    r.body.appendTable([['Name', 'Company', 'Email'], ['Pat Kim', 'ExampleCorp', 'pat.kim@example.com']]);
  });
  const index = S.buildStatusDocsIndex_(r.api.Docs.Documents.get(REPORT_ID, { includeTabsContent: true }));
  check('the tab, and the definition of the report by its title', [index.tabId, index.definition.title, index.definition.options.length, index.definition.id], [REPORT_TAB_ID, TITLE, 14, Object.keys(r.api.definitions)[0]]);
  check('a dropdown status: its id, its definition, the selected option and its value', (({ kind, displayValue, text, definitionId }) => [kind, displayValue, text, definitionId])(index.tables['S4aA269001']), ['dropdown', 'agreed', '', index.definition.id]);
  check('the ids are those of the document', [index.tables['S4aA269001'].dropdownId, index.tables['S4aA269001'].selectedOptionId], [r.api.statusCell('S4aA269001')._dropdown.dropdownId, r.api.statusCell('S4aA269001')._dropdown.optionId]);
  check('a text status: its text and where it is, so that a dropdown can take its place', (({ kind, text, insertable, start, end }) => [kind, text, insertable, end - start])(index.tables['S4aA269005']), ['text', 'not treated', true, 11]);
  check('an empty status cell: text, nothing to replace', (({ kind, text, insertable }) => [kind, text, insertable])(index.tables['S4aA269006']), ['text', '', false]);
  check('text beside a dropdown: it is a dropdown status', [index.tables['S4aA269002'].kind, index.tables['S4aA269002'].displayValue, index.tables['S4aA269002'].text], ['dropdown', 'available', 'typed beside']);
  check('a number that has two tables: read from the first, and marked', [index.tables['S4aA269001'].kind, index.duplicates], ['dropdown', { S4aA269001: true }]);
  check('a TDoc table without a Status row and a table that is no TDoc table are not listed', [Object.keys(index.tables).indexOf('S4aA269050'), Object.keys(index.tables).indexOf('Company'), Object.keys(index.tables).length], [-1, -1, 9]);
  check('a document without a tab, and nothing: null', [S.buildStatusDocsIndex_({ tabs: [] }), S.buildStatusDocsIndex_({}), S.buildStatusDocsIndex_(null)], [null, null, null]);
  const cell = (elements) => S.statusDocsCellEntry_({ content: [{ paragraph: { elements: elements } }] });
  check('a cell whose text is not one plain paragraph is not replaceable', [
    cell([{ startIndex: 5, endIndex: 12, textRun: { content: 'agreed\n' } }]).insertable,
    cell([{ startIndex: 5, endIndex: 8, textRun: { content: 'agr' } }, { startIndex: 8, endIndex: 12, textRun: { content: 'eed\n' } }]).insertable,
    cell([{ startIndex: 5, endIndex: 12, textRun: { content: 'agr\u000beed\n' } }]).text,
    cell([{ startIndex: 5, endIndex: 6, person: {} }, { startIndex: 6, endIndex: 13, textRun: { content: 'agreed\n' } }]).insertable,
    S.statusDocsCellEntry_({ content: [{ paragraph: { elements: [{ startIndex: 5, endIndex: 12, textRun: { content: 'agreed\n' } }] } }, { paragraph: { elements: [{ startIndex: 12, endIndex: 13, textRun: { content: '\n' } }] } }] }).insertable],
    [true, true, 'agr\u000beed', false, false]);
  check('a cell that holds something beside its one paragraph (here: a table) is not replaceable', S.statusDocsCellEntry_({ content: [{ paragraph: { elements: [{ startIndex: 5, endIndex: 12, textRun: { content: 'agreed\n' } }] } }, { table: { tableRows: [] } }] }).insertable, false);
  const foreign = report();
  foreign.api.Docs.Documents.batchUpdate({ requests: [{ createDropdownDefinition: { dropdownDefinition: { dropdownDefinitionProperties: { title: 'Review status', options: [{ displayValue: 'agreed' }, { displayValue: 'noted' }] } } } }] }, REPORT_ID);
  check('a dropdown definition of another title in the report is not the status definition', S.buildStatusDocsIndex_(foreign.api.Docs.Documents.get(REPORT_ID, { includeTabsContent: true })).definition, null);
  foreign.build();
  check('a build then creates "Document Status" beside it, and the statuses are dropdowns of that one', [Object.keys(foreign.api.definitions).map((id) => foreign.api.definitions[id].dropdownDefinitionProperties.title), foreign.api.definitions[foreign.api.statusCell('S4aA269001')._dropdown.definitionId].dropdownDefinitionProperties.title], [['Review status', TITLE], TITLE]);
  check('the range of a replaceable text is the text without its line end', (({ start, end }) => [start, end])(cell([{ startIndex: 5, endIndex: 12, textRun: { content: 'agreed\n' } }])), [5, 11]);
}

// ================================================================ 4. Build Report from Scratch

console.log('\n4. Build Report from Scratch');
{
  const r = report();
  const built = r.build();
  check('the build completes; its last phase is the status dropdowns', [built.ok, built.phases.map((p) => p.name + ' ' + p.status)], [true, ['skeleton done', 'e-mail done', 'revisions done', 'abstracts skipped', 'formatting done', 'status dropdowns done']]);
  check('every TDoc: a dropdown where the Portal status has an option, the Portal\'s own text where it has none', r.all(), ['S4aA269001 dropdown:agreed', 'S4aA269002 dropdown:available', 'S4aA269003 dropdown:reserved',
    'S4aA269004 dropdown:replied', 'S4aA269005 text:not treated', 'S4aA269006 empty', 'S4aA269007 text:other', 'S4aA269008 dropdown:noted', 'S4aA269009 dropdown:approved']);
  check('"not treated" is not turned into the option "other", and the Portal status "other" is text, not the option', [r.shown('S4aA269005'), r.shown('S4aA269007')], ['text:not treated', 'text:other']);
  check('a dropdown cell has no text left, and DocumentApp reads \'\' from it', [r.api.statusCell('S4aA269001')._t, r.api.statusCell('S4aA269001').getText(), r.s.statusCellHoldsDropdown_(r.api.statusCell('S4aA269001')), r.s.statusCellHoldsDropdown_(r.api.statusCell('S4aA269005'))], ['', '', true, false]);
  check('the registration-section TDoc (agenda item 1) has its table and its dropdown', [rowsOf(r.table('S4aA269009')).filter((row) => row[0] === 'Agenda Item').length, r.shown('S4aA269009')], [1, 'dropdown:approved']);
  check('the rest of every TDoc table is what the build wrote', rowsOf(r.table('S4aA269001')).map((row) => row.join('=')),
    ['TDoc=S4aA269001', 'Title=Synthetic title of S4aA269001', 'Source=ExampleCorp', 'Contact=Sam Rivera', 'Agenda Item=4.3', 'Type/For=discussion for Agreement', 'E-mail Discussion=', 'Revisions=', 'Minutes=', 'Disposition=', 'Status=']);
  check('the report is marked as one with status dropdowns', r.docProps.getProperty('STATUS_DROPDOWNS'), '1');
  check('the phase says what it did', built.phases[5].detail, '6 dropdown(s) inserted');
  check('what is logged names no TDoc, title or source', r.statusLogs().concat(r.logs.filter((l) => /dropdown/i.test(l))).filter((l) => /S4aA|Synthetic|ExampleCorp|Rivera/.test(l)), []);

  const again = r.build();
  check('a rebuild gives the same statuses', [again.ok, r.all()], [true, ['S4aA269001 dropdown:agreed', 'S4aA269002 dropdown:available', 'S4aA269003 dropdown:reserved',
    'S4aA269004 dropdown:replied', 'S4aA269005 text:not treated', 'S4aA269006 empty', 'S4aA269007 text:other', 'S4aA269008 dropdown:noted', 'S4aA269009 dropdown:approved']]);
  r.api.pick('S4aA269002', 'parked');
  r.build();
  check('a value picked by hand in a dropdown is still there after a rebuild (section 13)', r.shown('S4aA269002'), 'dropdown:parked');
}

// ================================================================ 5. the order of the work

console.log('\n5. the order of the work: everything DocumentApp does, one save, one read, then the writes; nothing after');
{
  const r = report();
  r.build();
  check('the build: save, read the report, read the template, create the definition, insert the dropdowns',
    r.last, ['save', 'get report', 'get template', 'batchUpdate createDropdownDefinition x1', 'batchUpdate deleteContentRange+insertDropdown x12']);
  check('the report is read once per build, with a field mask', r.api.calls.filter((c) => c.call === 'get' && c.id === REPORT_ID).map((c) => [c.includeTabsContent, /^revisionId,tabs\(tabProperties\(tabId\),documentTab\(dropdownDefinitions,body\(content\(/.test(c.fields)]), [[true, true]]);
  check('no write reached a document with unsaved changes', r.api.calls.filter((c) => c.call === 'batchUpdate').map((c) => c.dirty), [false, false]);
  check('each text is deleted and its dropdown inserted in the same batch, pair by pair', (() => { const q = r.api.calls.filter((c) => c.call === 'batchUpdate')[1].requests; return q.map((x) => Object.keys(x)[0]).join(',') === Array(6).fill('deleteContentRange,insertDropdown').join(','); })(), true);
  check('every pair is one cell: the dropdown goes where the text began', (() => { const q = r.api.calls.filter((c) => c.call === 'batchUpdate')[1].requests; const out = []; for (let i = 0; i < q.length; i += 2) out.push(q[i].deleteContentRange.range.startIndex === q[i + 1].insertDropdown.location.index); return out; })(), [true, true, true, true, true, true]);
  check('the pairs run from the end of the document to its beginning', (() => { const q = r.api.calls.filter((c) => c.call === 'batchUpdate')[1].requests; const starts = []; for (let i = 0; i < q.length; i += 2) starts.push(q[i].deleteContentRange.range.startIndex); return starts.every((x, i) => i === 0 || x < starts[i - 1]); })(), true);
  check('every request names the tab', r.api.calls.filter((c) => c.call === 'batchUpdate')[1].requests.map((x) => (x.deleteContentRange ? x.deleteContentRange.range.tabId : x.insertDropdown.location.tabId)).filter((t) => t !== REPORT_TAB_ID), []);
  check('after the save nothing changes the document through DocumentApp: it is closed, and the build did not fail', [r.body._state.closed, r.body._state.dirty], [true, false]);
  check('the step is the last thing a build and an update do (source)', [
    /phase\('status dropdowns', function \(\) \{[\s\S]*?\n    \}\);\n  \}\n\n  const totalMs = Date.now\(\) - totalStart;/.test(CODE),
    /finalizeStatusDropdowns_\('update'\);\n      if \(unmapped\.length\) result\.statusNotes = unmapped;\n    \}\n\n    Logger\.log\('=== COMPLETE ==='\);/.test(CODE),
    (CODE.match(/finalizeStatusDropdowns_\('(build|update)'\)/g) || []).length], [true, true, 3]);
  check('it never runs in a background run, which has no active document to save (source)', /if \(!context\) \{\n      const unmapped = statusDropdownRun_\(\)\.unmapped\.slice\(\);\n      finalizeStatusDropdowns_\('update'\);/.test(CODE), true);
}

// ================================================================ 6. Update Report Now

console.log('\n6. Update Report Now: new TDocs, revision tables');
{
  const r = report();
  r.build();
  const calls = r.api.calls.length;
  const done = r.update();
  check('an update without anything new: completes, with the one read the status sync needs to know what is selected -- no save, no write', [done.success, r.last, r.api.calls.length - calls], [true, ['get report'], 1]);

  r.tdocs.push(['S4aA269010', '4.4', 'agreed'], ['S4aA269011', '4.4', 'reserved'], ['S4aA269012', '4.4', 'not treated']);
  const added = r.update();
  check('new TDocs get their table, and a dropdown where the status has an option', [added.success, r.shown('S4aA269010'), r.shown('S4aA269011'), r.shown('S4aA269012')], [true, 'dropdown:agreed', 'dropdown:reserved', 'text:not treated']);
  check('the read of the sync, then one save, one read, one batch for the two of them; no definition is created, the template is not read', r.last, ['get report', 'save', 'get report', 'batchUpdate deleteContentRange+insertDropdown x4']);
  check('the statuses that were there are untouched', r.all().slice(0, 9), ['S4aA269001 dropdown:agreed', 'S4aA269002 dropdown:available', 'S4aA269003 dropdown:reserved',
    'S4aA269004 dropdown:replied', 'S4aA269005 text:not treated', 'S4aA269006 empty', 'S4aA269007 text:other', 'S4aA269008 dropdown:noted', 'S4aA269009 dropdown:approved']);
  check('the new dropdowns are of the report\'s one definition', [Object.keys(r.api.definitions).length, r.api.statusCell('S4aA269010')._dropdown.definitionId === Object.keys(r.api.definitions)[0]], [1, true]);
  r.update();
  check('the update after it writes nothing more', r.last, ['get report']);

  // A revision table, as the revision collector adds it.
  r.exec(() => {
    r.s.insertRevisedDocTablesAfter_(r.body, r.table('S4aA269001'), [{ text: 'S4aA269101_ExampleCorp.docx', link: 'https://www.example.org/docs/S4aA269101_ExampleCorp.docx' }], {});
    return r.s.finalizeStatusDropdowns_('update');
  });
  check('a revision table gets the dropdown of its status "Available"', [!!r.table('S4aA269101'), r.shown('S4aA269101')], [true, 'dropdown:available']);
  check('Update TDoc Revisions ends with the same step (source)', /function collectRevisionsOnly\(\) \{\n  resetStatusDropdownRun_\(\);\n  collectRevisionsCore_\(\);\n  \/\/[^\n]*\n  finalizeStatusDropdowns_\('update'\);\n  DocumentApp\.getUi\(\)\.alert\(/.test(CODE), true);

  // What the user changed in a new table before the end of the run is not overwritten.
  const edited = report();
  edited.build();
  edited.tdocs.push(['S4aA269010', '4.4', 'agreed']);
  edited.exec(() => {
    edited.s.insertNewTdoc_ = ((real) => function (b, td, cfg, index, context) { const out = real(b, td, cfg, index, context); edited.api.statusCell('S4aA269010').setText('agreed, see minutes'); return out; })(edited.s.insertNewTdoc_);
    return edited.s.continuousUpdateCore_();
  });
  check('a Status text that is no longer the status that was written is left alone', edited.shown('S4aA269010'), 'text:agreed, see minutes');
}

// ================================================================ 7. the readers

console.log('\n7. the readers: the status summary and the discussion e-mails');
{
  const r = report();
  r.build();
  r.api.pick('S4aA269002', 'agreed');                                  // decided in the meeting
  r.api.pick('S4aA269003', 'noted');
  const table = (number) => r.table(number);
  r.exec(() => {
    check('readTdocStatus_(): the value selected in a dropdown', ['S4aA269001', 'S4aA269002', 'S4aA269003', 'S4aA269004'].map((n) => r.s.readTdocStatus_(table(n), '')), ['agreed', 'agreed', 'noted', 'replied']);
    check('readTdocStatus_(): the text of a text status, exactly as the caller read it', [r.s.readTdocStatus_(table('S4aA269005'), 'not treated'), r.s.readTdocStatus_(table('S4aA269005'), '  not treated '), r.s.readTdocStatus_(table('S4aA269006'), '')], ['not treated', 'not treated', '']);
    check('readTdocStatus_(): text given wins over a dropdown beside it', r.s.readTdocStatus_(table('S4aA269001'), 'typed'), 'typed');
  });
  check('all of that with one read of the report, with the field mask', [r.last, /^revisionId,tabs\(/.test(r.api.calls[r.api.calls.length - 1].fields)], [['get report'], true]);

  r.exec(() => r.s.analyzeReportStatus());
  const summary = r.ui.alerts[r.ui.alerts.length - 1][1];
  const line = (label) => (summary.split('\n').filter((l) => l.indexOf(label) !== -1)[0] || '').replace(/\s+/g, ' ').trim();
  check('Report Status Summary counts the values selected in the dropdowns, the two set by hand included',
    [/Agreed: 2\b/.test(line('Agreed')), /Noted: 2\b/.test(line('Noted')), /Reserved: 0\b/.test(line('Reserved')), /Available: 0\b/.test(line('Available'))], [true, true, true, true]);
  check('it read the report once through the API and changed nothing', [r.last, r.body._state.dirty], [['get report'], false]);

  const found = r.exec(() => r.s.detectTdocTablesInDocument_(r.body));
  const byTdoc = {}; found.forEach((f) => { byTdoc[f.tdoc] = f; });
  check('Prepare Discussion E-mails: a TDoc set to agreed by hand is excluded, like one the Portal has as agreed',
    ['S4aA269001', 'S4aA269002', 'S4aA269009'].map((n) => [byTdoc[n].status, byTdoc[n].excluded, byTdoc[n].hardExclusionReason]), [['agreed', true, 'approved-agreed'], ['agreed', true, 'approved-agreed'], ['approved', true, 'approved-agreed']]);
  check('a TDoc set to noted by hand is offered; a text status counts as before', [['S4aA269003', 'S4aA269004', 'S4aA269005'].map((n) => [byTdoc[n].status, byTdoc[n].excluded])], [[['noted', false], ['replied', false], ['not treated', false]]]);
  check('one read for the whole list', r.last, ['get report']);
  r.api.pick('S4aA269003', 'reserved');
  const again = r.exec(() => r.s.detectTdocTablesInDocument_(r.body));
  check('a new execution reads again: a TDoc set to reserved is excluded as reserved', again.filter((f) => f.tdoc === 'S4aA269003').map((f) => [f.status, f.excluded, f.hardExclusionReason]), [['reserved', true, 'reserved']]);

  const tableIndex = r.body.getTables().indexOf(r.table('S4aA269002'));
  // The sender and recipient configuration is not what is tested here.
  r.s.resolveEmailExportConfiguration_ = () => ({ senderAddress: 'reporter@example.com', recipientAddress: 'list@example.org', mailingList: 'EXAMPLE_LIST' });
  const refused = r.exec(() => { try { const out = r.s.generateTdocDiscussionEmails([{ tableIndex: tableIndex }], {}); return JSON.stringify(out); } catch (e) { return 'threw: ' + e.message; } });
  check('generating an e-mail for a TDoc that is agreed in its dropdown is refused, by the status read at that moment', /S4aA269002 is already agreed and is excluded from discussion e-mail export/.test(refused), true);

  // When the API cannot be read.
  r.api.fail.get = () => 'synthetic read failure';
  const blind = r.exec(() => r.s.detectTdocTablesInDocument_(r.body));
  check('the API cannot be read: the list is still made; a dropdown status counts as unknown, a text status as before', [blind.length, blind.filter((f) => f.tdoc === 'S4aA269001')[0].status, blind.filter((f) => f.tdoc === 'S4aA269005')[0].status], [9, '', 'not treated']);
  check('and it is logged once', r.statusLogs().filter((l) => /could not be read \(synthetic read failure\)/.test(l)).length, 1);
  r.api.fail.get = null;
}

// ================================================================ 8. the Portal sync

console.log('\n8. the Portal sync: what moves on, what is kept');
{
  const r = report();
  r.build();
  // The Portal moves on.
  r.portal('S4aA269003', 'available');                                  // reserved -> available: follows
  r.portal('S4aA269002', 'agreed');                                     // available -> agreed: follows
  r.portal('S4aA269001', 'noted');                                      // agreed -> noted: a decision is not replaced
  r.portal('S4aA269004', 'revised');                                    // replied -> revised: "revised" always applies
  const done = r.update();
  check('reserved moves on to available, available to agreed; a decision is kept; "revised" is applied', [done.success, r.shown('S4aA269003'), r.shown('S4aA269002'), r.shown('S4aA269001'), r.shown('S4aA269004')],
    [true, 'dropdown:available', 'dropdown:agreed', 'dropdown:agreed', 'dropdown:revised']);
  check('one read at the start, then save, one read, one batch of three updates', r.last, ['get report', 'save', 'get report', 'batchUpdate updateDropdownProperties x3']);
  check('a dropdown is set by its id, with the field mask of the one value', r.api.calls[r.api.calls.length - 1].requests.map((q) => [Object.keys(q.updateDropdownProperties).sort().join(','), q.updateDropdownProperties.fields]),
    Array(3).fill(['dropdownId,dropdownProperties,fields,tabId', 'selectedOptionId']));
  check('the cells still have no text: a dropdown is never written with setText', ['S4aA269001', 'S4aA269002', 'S4aA269003', 'S4aA269004'].map((n) => r.api.statusCell(n)._t), ['', '', '', '']);
  r.update();
  check('the next update has nothing to do: one read, no write', r.last, ['get report']);

  // A decision made in the Minutes.
  r.api.pick('S4aA269003', 'parked');                                   // available -> parked by hand
  r.portal('S4aA269003', 'agreed');
  r.update();
  check('a status set by hand is kept when the Portal says something else', [r.shown('S4aA269003'), r.last], ['dropdown:parked', ['get report']]);
  r.api.pick('S4aA269008', 'available');                                // noted -> available by hand
  r.portal('S4aA269008', 'withdrawn');
  r.update();
  check('a dropdown that stands at available follows the Portal, whoever set it', r.shown('S4aA269008'), 'dropdown:withdrawn');

  // A Portal status without an option, for a TDoc that has a dropdown.
  r.portal('S4aA269002', 'available');
  r.api.pick('S4aA269002', 'available');
  r.portal('S4aA269002', 'not treated');
  const unmapped = r.update();
  check('a Portal status that has no option: the dropdown and its value stay; nothing is written; it is not replaced by text', [r.shown('S4aA269002'), r.last, r.api.statusCell('S4aA269002')._t], ['dropdown:available', ['get report'], '']);
  check('it is reported with the result of the update and logged', [unmapped.success, unmapped.statusNotes, r.statusLogs().filter((l) => /has no option in the Status dropdown/.test(l)).length > 0],
    [true, ['S4aA269002: the Portal status "not treated" has no option in the Status dropdown; "available" is kept.'], true]);
  r.portal('S4aA269001', 'not treated');
  const kept = r.update();
  check('where the rule would not have changed the status anyway, nothing is reported for it', (kept.statusNotes || []).filter((n) => /S4aA269001/.test(n)), []);
  r.portal('S4aA269002', 'other');
  r.update();
  check('the Portal status "other" does not select the option other', r.shown('S4aA269002'), 'dropdown:available');

  // Text statuses in the same report keep the rule they had.
  r.portal('S4aA269005', 'agreed');                                     // "not treated" is neither reserved nor available: kept
  r.portal('S4aA269006', 'available');                                  // empty: kept, as before
  r.update();
  check('a text status follows the rule as before: "not treated" stays', [r.shown('S4aA269005'), r.shown('S4aA269006')], ['text:not treated', 'empty']);
  r.portal('S4aA269005', 'revised');
  r.update();
  check('and "revised" is written into it, as before; being a status with an option, it is a dropdown by the end of that update', [r.shown('S4aA269005'), r.last], ['dropdown:revised', ['get report', 'save', 'get report', 'batchUpdate deleteContentRange+insertDropdown x2']]);

  // Text typed beside a dropdown.
  r.user(() => r.api.statusCell('S4aA269009').setText('see minutes'));
  r.portal('S4aA269009', 'revised');
  r.update();
  check('a cell with a dropdown is never written as text, even with text typed beside it: the dropdown is set', r.shown('S4aA269009'), 'text:see minutes+dropdown:revised');

  // The rule itself, on the value of a dropdown.
  const rule = report();
  rule.build();
  const sync = (number, selected, portal) => { rule.api.pick(number, selected); rule.portal(number, portal); rule.update(); return rule.shown(number).replace('dropdown:', ''); };
  check('from reserved or available the Portal status is taken; from anything else only "revised"', [
    sync('S4aA269001', 'reserved', 'agreed'), sync('S4aA269001', 'available', 'noted'), sync('S4aA269001', 'noted', 'agreed'), sync('S4aA269001', 'agreed', 'available'),
    sync('S4aA269001', 'approved', 'revised'), sync('S4aA269001', 'parked', 'revised'), sync('S4aA269001', 'Plenary', 'withdrawn'), sync('S4aA269001', 'other', 'agreed')],
    ['agreed', 'noted', 'noted', 'agreed', 'revised', 'revised', 'Plenary', 'other']);
  check('a Portal status equal to the selected value writes nothing', (() => { rule.api.pick('S4aA269002', 'agreed'); rule.portal('S4aA269002', 'Agreed'); rule.update(); return rule.last; })(), ['get report']);

  // The old sheet-based path.
  const legacy = report();
  legacy.build();
  // "revised" is the status that path writes whatever the cell says.
  const sheet = { getDataRange: () => ({ getValues: () => [['TDoc', 'S4aA269003'], ['TDoc Status', 'revised']] }) };
  legacy.api.statusCell('S4aA269003').clear = function () { this.setText(''); return this; };
  legacy.api.statusCell('S4aA269005').clear = function () { this.setText(''); return this; };
  legacy.exec(() => legacy.s.updateStatusWithStrictRules_(sheet, legacy.table('S4aA269003')));
  check('the old sheet path does not write into a Status cell that holds a dropdown, even the status it always writes', [legacy.shown('S4aA269003'), legacy.body._state.dirty], ['dropdown:reserved', false]);
  legacy.exec(() => legacy.s.updateStatusWithStrictRules_(sheet, legacy.table('S4aA269005')));
  check('and writes it into a text status as it always did', legacy.shown('S4aA269005'), 'text:revised');
}

// ================================================================ 9. many TDocs; moved tables

console.log('\n9. many TDocs: positions and batches; tables that were moved');
{
  const STATES = ['agreed', 'available', 'reserved', 'noted', 'not treated', 'approved', 'replied to'];
  const many = Array.from({ length: 130 }, (_, i) => ['S4aA26' + String(9001 + i), i % 2 ? '4.3' : '4.4', STATES[i % STATES.length]]);
  const r = report({ tdocs: many });
  const built = r.build();
  const expected = many.map((t) => t[0] + ' ' + (t[2] === 'not treated' ? 'text:not treated' : 'dropdown:' + (t[2] === 'replied to' ? 'replied' : t[2])));
  check('130 TDocs: every one has exactly the status of its own row', [built.ok, r.all().filter((x, i) => x !== expected[i])], [true, []]);
  const batches = r.api.calls.filter((c) => c.call === 'batchUpdate' && c.requests[0].deleteContentRange);
  check('the 112 dropdowns are sent in three batches of at most fifty', batches.map((b) => b.requests.length / 2), [50, 50, 12]);
  check('one save and one read of the report for all of them', [r.last.filter((e) => e === 'save').length, r.last.filter((e) => e === 'get report').length], [1, 1]);
  check('across the batches the positions only go down: a batch never moves what a later one writes', (() => { const starts = []; batches.forEach((b) => b.requests.forEach((q) => { if (q.deleteContentRange) starts.push(q.deleteContentRange.range.startIndex); })); return starts.every((x, i) => i === 0 || x < starts[i - 1]); })(), true);
  check('no title, number or other cell of any table was changed by it', many.map((t) => rowsOf(r.table(t[0])).slice(0, 5).map((row) => row[1]).join('|')).filter((x, i) => x !== [many[i][0], 'Synthetic title of ' + many[i][0], 'ExampleCorp', 'Sam Rivera', many[i][1]].join('|')), []);

  // The plan itself.
  const index = S.buildStatusDocsIndex_(r.api.Docs.Documents.get(REPORT_ID, { includeTabsContent: true }));
  const plan = S.planStatusDropdownRequests_(index, index.definition, { S4aA269005: 'not treated', S4aA269001: 'agreed', S4aA269999: 'agreed' }, { S4aA269001: 'noted', S4aA269005: 'agreed', S4aA269002: 'available', S4aA269003: 'not treated' });
  check('the plan: a dropdown is updated by number; a text status and an unknown number are left; an equal value needs no request',
    [plan.updated, plan.updateRequests.length, plan.insertPairs.map((p) => p.number), plan.left.sort()], [['S4aA269001'], 1, [], ['S4aA269001', 'S4aA269003', 'S4aA269005', 'S4aA269999']]);
  const twice = report({ docs: false });
  twice.build();
  twice.user(() => twice.body.appendTable([['TDoc', 'S4aA269001'], ['Status', 'agreed']]));
  const twiceIndex = S.buildStatusDocsIndex_(twice.api.Docs.Documents.get(REPORT_ID, { includeTabsContent: true }));
  check('a number that has two tables gets no dropdown; the others do', S.planStatusDropdownRequests_(twiceIndex, { id: 'kix.x', title: TITLE, options: optionsWithIds(DEFINITION) }, { S4aA269001: 'agreed', S4aA269002: 'available' }, {}).insertPairs.map((x) => x.number), ['S4aA269002']);
  check('without a definition nothing is inserted', S.planStatusDropdownRequests_(index, null, { S4aA269005: 'not treated' }, {}).insertPairs, []);

  // A batch that is refused is refused as a whole.
  const partial = report({ tdocs: many });
  partial.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange && requests.length === 100 && partial.api.calls.filter((c) => c.call === 'batchUpdate' && c.requests[0].deleteContentRange).length === 2 ? 'synthetic failure of the second batch' : null);
  const half = partial.build();
  const shown = partial.all();
  check('the second of three batches fails: the build still completes', [half.ok, half.phases[5].status], [true, 'done']);
  check('every Status is whole: a dropdown or its text, never an emptied cell', [shown.filter((x) => / empty$/.test(x)).length, shown.filter((x) => /dropdown:/.test(x)).length, shown.filter((x) => /text:/.test(x)).length], [0, 50, 80]);
  check('the texts that are left are the statuses the Portal has', shown.filter((x, i) => /text:/.test(x) && x !== many[i][0] + ' text:' + many[i][2]), []);
  check('it is logged', partial.statusLogs().filter((l) => /not completed \(synthetic failure of the second batch\)/.test(l)).length, 1);

  // Tables that are moved during the run.
  const moved = report();
  moved.build();
  moved.portal('S4aA269003', 'available');
  moved.exec(() => {
    moved.s.rearrangeRevisionTables_ = () => { const copy = moved.s.moveTableAfter_(moved.body, moved.table('S4aA269003'), moved.table('S4aA269001')); return { moved: copy ? 1 : 0, dispositions: 0 }; };
    return moved.s.continuousUpdateCore_();
  });
  const order = moved.body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2).map((t) => t.getRow(0).getCell(1).getText());
  check('a table that was moved in the same run, after its status was decided: it is where it was moved to, and has the new status', [order.indexOf('S4aA269003') === order.indexOf('S4aA269001') + 1, moved.shown('S4aA269003')], [true, 'dropdown:available']);
  check('the dropdown of the moved table is a new one (a copy), and it is the one that was set', [/^kix\.copy/.test(moved.api.statusCell('S4aA269003')._dropdown.dropdownId), moved.api.calls[moved.api.calls.length - 1].requests[0].updateDropdownProperties.dropdownId === moved.api.statusCell('S4aA269003')._dropdown.dropdownId], [true, true]);
  // A reallocation moves the table of a TDoc under another agenda item.
  moved.api.pick('S4aA269002', 'parked');
  const headingOf = (number) => { const at = moved.body._children.indexOf(moved.table(number)); for (let i = at; i >= 0; i--) { const c = moved.body._children[i]; if (c.getType() !== 'TABLE' && c.getHeading() !== 'NORMAL') return c.getText(); } return ''; };
  const beforeHeading = headingOf('S4aA269002');
  moved.exec(() => moved.s.applyDocumentReallocationsToBody_(moved.body, { S4aA269002: { original: '4.3', new: '4.4', reason: 'synthetic' } }));
  check('a reallocated table is under its new agenda item and keeps its dropdown with the value set by hand',
    [beforeHeading, headingOf('S4aA269002'), moved.shown('S4aA269002')], ['4.3 Performance requirements', '4.4 Design constraints', 'dropdown:parked']);
  const copy = moved.exec(() => moved.s.moveTableAfter_(moved.body, moved.table('S4aA269001'), moved.table('S4aA269008')));
  check('a moved table keeps a working dropdown with its value; the readers find it by its number', [!!copy, moved.shown('S4aA269001'), moved.exec(() => moved.s.readTdocStatus_(moved.table('S4aA269001'), ''))], [true, 'dropdown:agreed', 'agreed']);
  moved.portal('S4aA269004', 'revised');
  moved.update();
  check('and the update after a move sets the right dropdowns', [moved.shown('S4aA269004'), moved.shown('S4aA269001')], ['dropdown:revised', 'dropdown:agreed']);
}

// ================================================================ 10. when something is missing or fails

console.log('\n10. when something is missing or fails: text statuses, as before');
{
  const TEXT = ['S4aA269001 text:agreed', 'S4aA269002 text:available', 'S4aA269003 text:reserved', 'S4aA269004 text:Replied to', 'S4aA269005 text:not treated', 'S4aA269006 empty', 'S4aA269007 text:other', 'S4aA269008 text:  NOTED ', 'S4aA269009 text:approved'];
  const outcome = (r) => { const built = r.build(); return [built.ok, built.phases.map((p) => p.name), r.all(), r.docProps.getProperty('STATUS_DROPDOWNS')]; };
  const FIVE = ['skeleton', 'e-mail', 'revisions', 'abstracts', 'formatting'];
  const SIX = FIVE.concat(['status dropdowns']);

  const noDocs = report({ docs: false });
  check('no Google Docs API service: the five phases of every release, text statuses, no save', [outcome(noDocs), noDocs.events], [[true, FIVE, TEXT, null], []]);
  check('without the service a new TDoc is not even remembered', (() => { noDocs.s.noteStatusDropdownCandidate_('S4aA269099', 'agreed'); return Object.keys(noDocs.s.statusDropdownRun_().inserts); })(), []);
  const central = report({ release: false });
  check('not a template report (CENTRAL, Legacy), even with the service: the same, and no API call', [outcome(central), central.events], [[true, FIVE, TEXT, null], []]);

  const noTemplate = report({ template: null });
  check('the template cannot be read: text statuses; the build completes; the report is not marked', outcome(noTemplate), [true, SIX, TEXT, null]);
  check('it is logged, and nothing was written', [noTemplate.statusLogs().filter((l) => /not completed \(Requested entity was not found\.\)/.test(l)).length, noTemplate.api.count('batchUpdate')], [1, 0]);

  const noDefinition = report({ template: [{ title: 'Review status', options: [{ displayValue: 'a' }, { displayValue: 'b' }] }] });
  check('the template has no "Document Status" dropdown: text statuses', outcome(noDefinition), [true, SIX, TEXT, null]);
  check('the phase says why', noDefinition.exec(() => noDefinition.s.runFullReportBuildCore_()).phases[5].detail, '0 dropdown(s) inserted -- no "Document Status" dropdown is defined in the template');

  const createFails = report();
  createFails.api.fail.batchUpdate = (requests) => (requests[0].createDropdownDefinition ? 'synthetic failure of the definition' : null);
  check('the definition cannot be created: text statuses', outcome(createFails), [true, SIX, TEXT, null]);

  const insertFails = report();
  insertFails.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange ? 'synthetic failure of the inserts' : null);
  check('the dropdowns cannot be inserted: every status is its text, none is emptied; the report has its definition and is marked, so that a later update makes them (section 12)', outcome(insertFails), [true, SIX, TEXT, '1']);
  insertFails.api.fail.batchUpdate = null;
  check('the build after it makes the dropdowns, with the definition that was created', [insertFails.build().ok, insertFails.shown('S4aA269001'), Object.keys(insertFails.api.definitions).length, insertFails.docProps.getProperty('STATUS_DROPDOWNS')], [true, 'dropdown:agreed', 1, '1']);

  const readFails = report();
  readFails.api.fail.get = (id) => (id === REPORT_ID ? 'synthetic read failure' : null);
  check('the report cannot be read through the API: text statuses', outcome(readFails), [true, SIX, TEXT, null]);
  check('the read was tried with the mask and then without, and then given up', readFails.api.calls.map((c) => c.call + (c.fields ? ' masked' : ' whole')), ['get masked', 'get whole']);

  const maskRefused = report();
  maskRefused.api.fail.get = (id, opts) => (opts && opts.fields ? 'synthetic: invalid field mask' : null);
  check('only the field mask is refused: the document is read whole and the dropdowns are made', [maskRefused.build().ok, maskRefused.shown('S4aA269001'), maskRefused.docProps.getProperty('STATUS_DROPDOWNS')], [true, 'dropdown:agreed', '1']);

  const saveFails = report();
  saveFails.saveFails = true;
  check('the document cannot be saved: text statuses, no API call', [outcome(saveFails), saveFails.api.calls.length], [[true, SIX, TEXT, null], 0]);

  // An update of a report with dropdowns, when the API fails.
  const r = report();
  r.build();
  r.tdocs.push(['S4aA269010', '4.4', 'agreed']);
  r.portal('S4aA269003', 'available');
  r.api.fail.batchUpdate = () => 'synthetic write failure';
  const done = r.update();
  check('an update whose writes fail: it completes; the new TDoc has its text status; the dropdown keeps its value', [done.success, r.shown('S4aA269010'), r.shown('S4aA269003')], [true, 'text:agreed', 'dropdown:reserved']);
  r.api.fail.batchUpdate = null;
  r.update();
  check('the next update sets the dropdown, and makes the dropdown of the new TDoc that could not be made', [r.shown('S4aA269003'), r.shown('S4aA269010')], ['dropdown:available', 'dropdown:agreed']);
  r.api.fail.get = () => 'synthetic read failure';
  r.portal('S4aA269002', 'agreed');
  const blind = r.update();
  check('an update that cannot read the dropdowns: it completes and leaves every dropdown as it is', [blind.success, r.shown('S4aA269002'), r.api.statusCell('S4aA269002')._t], [true, 'dropdown:available', '']);
  r.api.fail.get = null;

  check('the step never throws: every failure above ended in a completed build or update (source: one try around all of it)',
    /function finalizeStatusDropdowns_\(mode\) \{\n  const result = [^\n]*\n  const run = statusDropdownRun_\(\);\n  try \{[\s\S]*\n  \} catch \(e\) \{\n    result\.reason = e\.message;\n[^\n]*\n  \} finally \{\n    resetStatusDropdownRun_\(\);\n  \}\n  return result;\n\}/.test(CODE), true);
}

// ================================================================ 11. reports with text statuses

console.log('\n11. reports with text statuses: unchanged, and no API call');
{
  // A report built before this release: text statuses, no mark.
  const old = report({ docs: false });
  old.build();
  const before = old.all();
  old.s.Docs = { Documents: { get: (id, opts) => { old.events.push('get'); return old.api.Docs.Documents.get(id, opts); }, batchUpdate: (resource, id) => { old.events.push('batchUpdate'); return old.api.Docs.Documents.batchUpdate(resource, id); } } };
  old.tdocs.push(['S4aA269010', '4.4', 'agreed']);
  old.portal('S4aA269003', 'available');
  const done = old.update();
  check('its update with this release: the new TDoc has a text status, the sync writes text as it did', [done.success, old.shown('S4aA269010'), old.shown('S4aA269003'), old.all().slice(0, 2)], [true, 'text:agreed', 'text:available', before.slice(0, 2)]);
  check('no API call, no save, no mark: nothing converts it', [old.events, old.docProps.getProperty('STATUS_DROPDOWNS'), old.body._state.closed], [[], null, false]);
  old.exec(() => old.s.analyzeReportStatus());
  old.exec(() => old.s.detectTdocTablesInDocument_(old.body));
  check('its status summary and its e-mail list make no API call either', old.events, []);
  check('the mark is set by a build, and by a complete conversion of an existing report (tests/status-dropdown-migration.test.js) -- nowhere else (source)',
    [(CODE.match(/setProperty\(STATUS_DROPDOWN_ENABLED_KEY_/g) || []).length, /if \(definition && mode === 'build'\) props\.setProperty\(STATUS_DROPDOWN_ENABLED_KEY_, '1'\);/.test(CODE), /if \(result\.ok\) \{ props\.setProperty\(STATUS_DROPDOWN_ENABLED_KEY_, '1'\); result\.marked = true; \}/.test(CODE)], [2, true, true]);

  // What a text report produces is what the previous release produced.
  const withService = report();
  const without = report({ docs: false });
  withService.s.finalizeStatusDropdowns_ = () => ({ ran: false, inserted: 0, updated: 0, reason: 'switched off for this comparison' });
  withService.build(); without.build();
  const snapshot = (x) => JSON.stringify(x.body._children.map((c) => (c.getType() === 'TABLE' ? ['T', rowsOf(c)] : ['P', c.getText(), c.getHeading()])));
  check('before the last step, the document of a build is the document every release built', snapshot(withService) === snapshot(without), true);

  check('the manifest: the Google Docs API service beside Drive, and no new scope', (() => { const m = JSON.parse(fs.readFileSync(require('path').join(__dirname, '..', 'appsscript.json'), 'utf8')); return [m.dependencies.enabledAdvancedServices, m.oauthScopes.length, m.oauthScopes.filter((x) => /documents$/.test(x)).length, m.timeZone, m.runtimeVersion]; })(),
    [[{ userSymbol: 'Drive', serviceId: 'drive', version: 'v3' }, { userSymbol: 'Docs', serviceId: 'docs', version: 'v1' }], 6, 1, 'Europe/Berlin', 'V8']);
  check('the feature stores one property of its own and reads one more (source)', [(CODE.match(/const STATUS_DROPDOWN_[A-Z_]*KEY_ = '[A-Z_]+';/g) || []).length, /STATUS_DROPDOWN/.test(JSON.stringify(S.ADDON003_ADOPTION_FIXED_KEYS_))], [2, false]);
  check('Code.js is version 2.21.1 (the status dropdowns came with 2.19.0)', (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.21.1');
}

// ================================================================ 12. a dropdown that could not be made

console.log('\n12. a dropdown that could not be made is made by a later update');
{
  const BUILT = ['S4aA269001 dropdown:agreed', 'S4aA269002 dropdown:available', 'S4aA269003 dropdown:reserved', 'S4aA269004 dropdown:replied', 'S4aA269005 text:not treated', 'S4aA269006 empty', 'S4aA269007 text:other', 'S4aA269008 dropdown:noted', 'S4aA269009 dropdown:approved'];
  const r = report();
  r.build();
  r.tdocs.push(['S4aA269010', '4.4', 'agreed'], ['S4aA269011', '4.4', 'not treated']);
  r.api.fail.batchUpdate = () => 'synthetic write failure';
  const failed = r.update();
  check('the insert for a new TDoc fails: the update completes, and the TDoc has its status as text', [failed.success, r.shown('S4aA269010'), r.shown('S4aA269011'), r.statusLogs().filter((l) => /not completed \(synthetic write failure\)/.test(l)).length], [true, 'text:agreed', 'text:not treated', 1]);
  check('the report is as usable as before: the other statuses are untouched, and the readers read the text', [r.all().slice(0, 9), r.exec(() => r.s.readTdocStatus_(r.table('S4aA269010'), 'agreed'))], [BUILT, 'agreed']);
  r.api.fail.batchUpdate = null;
  const healed = r.update();
  check('the next update makes the dropdown, with the value the text had', [healed.success, r.shown('S4aA269010'), r.api.statusCell('S4aA269010')._t], [true, 'dropdown:agreed', '']);
  check('a status without an option is not converted', r.shown('S4aA269011'), 'text:not treated');
  check('it took the read of the sync, one save, one read and one batch of one pair; no definition was created', [r.last, Object.keys(r.api.definitions).length], [['get report', 'save', 'get report', 'batchUpdate deleteContentRange+insertDropdown x2'], 1]);
  check('no dropdown that was there was written', r.all().slice(0, 9), BUILT);
  r.update();
  check('the update after it has nothing to do: one read, no save, no write', [r.last, r.shown('S4aA269010')], [['get report'], 'dropdown:agreed']);

  // A value picked meanwhile is not overwritten by the healing.
  r.api.pick('S4aA269010', 'parked');
  r.update();
  check('a dropdown is never a candidate, whatever it says', [r.shown('S4aA269010'), r.last], ['dropdown:parked', ['get report']]);

  // The first build: every insert fails.
  const first = report();
  first.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange ? 'synthetic failure of the inserts' : null);
  first.build();
  check('a first build whose inserts all fail: text statuses, one definition, and the report is marked', [first.all().filter((x) => /dropdown/.test(x)), Object.keys(first.api.definitions).length, first.docProps.getProperty('STATUS_DROPDOWNS')], [[], 1, '1']);
  first.api.fail.batchUpdate = null;
  const later = first.update();
  check('the next update makes every dropdown the build could not make', [later.success, first.all()], [true, BUILT]);
  check('with one read to find them (no status is a dropdown yet, so the sync read nothing), one save, one read, one batch; the definition is the one the build created',
    [first.last, Object.keys(first.api.definitions).length], [['get report', 'save', 'get report', 'batchUpdate deleteContentRange+insertDropdown x12'], 1]);
  first.update();
  check('and then nothing more', first.last, ['get report']);

  // Text typed by hand.
  first.user(() => { first.api.statusCell('S4aA269005').setText('agreed'); first.api.statusCell('S4aA269007').setText('agreed, with comments'); });
  first.update();
  check('a status typed as text in a dropdown report becomes a dropdown when it is exactly a status with an option, and stays text otherwise', [first.shown('S4aA269005'), first.shown('S4aA269007')], ['dropdown:agreed', 'text:agreed, with comments']);

  // What is a candidate.
  const index = S.buildStatusDocsIndex_(r.api.Docs.Documents.get(REPORT_ID, { includeTabsContent: true }));
  const entry = (kind, text, more) => Object.assign({ kind: kind, text: text, insertable: kind === 'text' && text.length > 0, start: 10, end: 10 + text.length, at: 10 }, more || {});
  const definition = { id: 'kix.x', title: TITLE, options: optionsWithIds(DEFINITION) };
  const candidates = S.statusDropdownHealCandidates_({ duplicates: { T7: true }, tables: {
    T1: entry('text', 'agreed'), T2: entry('text', ' Replied to '), T3: entry('text', 'not treated'), T4: entry('text', ''), T5: entry('text', 'other'), T6: entry('dropdown', '', { displayValue: 'agreed' }),
    T7: entry('text', 'agreed'), T8: entry('text', 'agreed', { insertable: false }), T9: entry('dropdown', 'agreed', { displayValue: 'noted' }), T10: entry('text', 'Plenary') } }, definition);
  check('a candidate: a plain text status with an option, in a table that occurs once -- not an unmapped, empty or manual-only status, not a dropdown, not text beside one, not a cell of another shape',
    candidates, { T1: 'agreed', T2: ' Replied to ' });
  check('without a definition, or without a read, there is none', [S.statusDropdownHealCandidates_(index, null), S.statusDropdownHealCandidates_(null, definition)], [{}, {}]);

  // A report that is not marked is never read for this.
  const plain = report({ docs: false });
  plain.build();
  plain.s.Docs = { Documents: { get: (id, opts) => { plain.events.push('get'); return plain.api.Docs.Documents.get(id, opts); }, batchUpdate: (resource, id) => { plain.events.push('batchUpdate'); return plain.api.Docs.Documents.batchUpdate(resource, id); } } };
  plain.update(); plain.update();
  check('a report with text statuses that is not marked: its updates look for nothing and make no API call, although every status "has an option"', [plain.events, plain.all()[0]], [[], 'S4aA269001 text:agreed']);

  // The read that looks for candidates fails.
  const blind = report();
  blind.api.fail.batchUpdate = (requests) => (requests[0].deleteContentRange ? 'synthetic failure of the inserts' : null);
  blind.build();
  blind.api.fail.batchUpdate = null;
  blind.api.fail.get = () => 'synthetic read failure';
  const done = blind.update();
  check('when even that read fails, the update completes and the statuses stay text', [done.success, blind.all().filter((x) => /dropdown/.test(x))], [true, []]);
  blind.api.fail.get = null;
  blind.update();
  check('and the update after it heals the report', blind.all(), BUILT);
}

// ================================================================ 13. Build Report from Scratch keeps what was selected

console.log('\n13. Build Report from Scratch keeps what was selected in a dropdown');
{
  const r = report();
  r.build();
  r.api.pick('S4aA269002', 'parked');                                   // available -> parked, by hand
  r.api.pick('S4aA269001', 'noted');                                    // agreed -> noted, by hand
  r.api.pick('S4aA269009', 'Plenary');                                  // approved -> Plenary, by hand; this TDoc will be gone
  r.api.pick('S4aA269008', 'agreed');                                   // noted -> agreed, by hand; the Portal will say revised
  r.portal('S4aA269003', 'available');                                  // nobody touched it: reserved, and the Portal moved on
  r.portal('S4aA269008', 'revised');
  r.portal('S4aA269004', 'withdrawn');                                  // nobody touched it: replied (from the Portal), and the Portal moved on
  r.tdocs.splice(r.tdocs.findIndex((t) => t[0] === 'S4aA269009'), 1);  // removed from the TDoc list
  r.tdocs.push(['S4aA269010', '4.4', 'agreed']);                        // new in the TDoc list
  const rebuilt = r.build();
  check('the rebuild completes', [rebuilt.ok, rebuilt.phases[5].status], [true, 'done']);
  check('a value set by hand is the value after the rebuild', [r.shown('S4aA269002'), r.shown('S4aA269001')], ['dropdown:parked', 'dropdown:noted']);
  check('a dropdown that stood at reserved follows the Portal, as in an update', r.shown('S4aA269003'), 'dropdown:available');
  check('a decision the Portal had set is kept against a later Portal status, as in an update', r.shown('S4aA269004'), 'dropdown:replied');
  check('a "revised" Portal status applies, as in an update', r.shown('S4aA269008'), 'dropdown:revised');
  check('a new TDoc has the option of its Portal status', r.shown('S4aA269010'), 'dropdown:agreed');
  check('a TDoc that is no longer in the list is not in the report, whatever its dropdown said', [r.shown('S4aA269009'), !!r.table('S4aA269009')], ['none', false]);
  check('the text statuses are the Portal\'s, as after every rebuild', [r.shown('S4aA269005'), r.shown('S4aA269006'), r.shown('S4aA269007')], ['text:not treated', 'empty', 'text:other']);
  check('one read before the document is cleared, then save, one read, one batch; no definition is created', [r.last, Object.keys(r.api.definitions).length], [['get report', 'save', 'get report', 'batchUpdate deleteContentRange+insertDropdown x12'], 1]);
  check('the values are kept by TDoc number: every table has the status that belongs to its own number', r.tdocs.map((t) => rowsOf(r.table(t[0]))[0][1] + ' ' + r.shown(t[0])),
    ['S4aA269001 dropdown:noted', 'S4aA269002 dropdown:parked', 'S4aA269003 dropdown:available', 'S4aA269004 dropdown:replied', 'S4aA269005 text:not treated', 'S4aA269006 empty', 'S4aA269007 text:other', 'S4aA269008 dropdown:revised', 'S4aA269010 dropdown:agreed']);
  r.build();
  check('a second rebuild gives the same', [r.shown('S4aA269002'), r.shown('S4aA269001'), r.shown('S4aA269003')], ['dropdown:parked', 'dropdown:noted', 'dropdown:available']);

  // The Portal status has no option, or is empty, where a value was set by hand.
  r.portal('S4aA269001', 'not treated');
  r.portal('S4aA269002', '');
  r.build();
  check('a value set by hand is kept also where the Portal status has no option, or is empty: the dropdown is made with it', [r.shown('S4aA269001'), r.shown('S4aA269002'), r.api.statusCell('S4aA269001')._t], ['dropdown:noted', 'dropdown:parked', '']);
  check('the dropdown of the empty cell was inserted without anything to delete', r.api.calls[r.api.calls.length - 1].requests.filter((q) => q.deleteContentRange).length, r.api.calls[r.api.calls.length - 1].requests.filter((q) => q.insertDropdown).length - 1);

  // The rule.
  const OPTIONS = optionsWithIds(DEFINITION);
  const choice = (previous, portal) => { const o = S.statusDropdownRebuildChoice_(previous, portal, OPTIONS); return o ? o.displayValue : null; };
  check('kept: a decision, including parked, Plenary and other, whatever the Portal says -- unless it says revised',
    [choice('parked', 'available'), choice('Plenary', 'agreed'), choice('other', 'noted'), choice('agreed', 'available'), choice('noted', 'not treated'), choice('agreed', ''), choice(' Agreed ', 'noted')],
    ['parked', 'Plenary', 'other', 'agreed', 'noted', 'agreed', 'agreed']);
  check('not kept: reserved and available (they follow the Portal), and anything when the Portal says revised', [choice('reserved', 'agreed'), choice('available', 'noted'), choice('available', ''), choice('agreed', 'revised'), choice('parked', 'Revised')], [null, null, null, null, null]);
  check('a previous value that is no option of the definition is not kept, and not guessed', [choice('replied to', 'agreed'), choice('agree', 'noted'), choice('', 'agreed'), choice(null, 'agreed'), choice('decided', 'available'), S.statusDropdownRebuildChoice_('agreed', 'noted', [])], [null, null, null, null, null, null]);
  check('a value is looked up by its own name only: no alias, and nothing maps to other', [S.statusDropdownOptionByName_('replied to', OPTIONS), S.statusDropdownOptionByName_('not treated', OPTIONS), S.statusDropdownOptionByName_('PLENARY', OPTIONS).displayValue], [null, null, 'Plenary']);

  // A previous value that the definition no longer has.
  const changed = report();
  changed.build();
  changed.api.pick('S4aA269001', 'noted');
  const def = changed.api.definitions[Object.keys(changed.api.definitions)[0]].dropdownDefinitionProperties;
  def.options.filter((x) => x.displayValue === 'noted')[0].displayValue = 'taken note of';      // the option was renamed in the report by hand
  changed.user(() => { changed.s.captureStatusDropdownValuesForRebuild_ = ((real) => function () { real(); changed.s.statusDropdownRun_().previous.S4aA269001 = 'no such status'; })(changed.s.captureStatusDropdownValuesForRebuild_); });
  const safe = changed.build();
  check('a previous value that cannot be represented: the rebuilt report is the normal one, with the Portal status', [safe.ok, changed.shown('S4aA269001'), changed.all().filter((x) => / empty$/.test(x)).length], [true, 'dropdown:agreed', 1]);

  // A registration-section TDoc, and the Portal status it has at the rebuild.
  const registration = report();
  registration.build();
  registration.api.pick('S4aA269009', 'noted');
  registration.api.pick('S4aA269001', 'noted');
  registration.portal('S4aA269009', 'revised');
  registration.build();
  check('the Portal status of a registration-section TDoc counts at the rebuild as that of any other: "revised" applies', [registration.shown('S4aA269009'), registration.shown('S4aA269001')], ['dropdown:revised', 'dropdown:noted']);

  // A text status of a dropdown report is not carried over.
  const typed = report();
  typed.build();
  typed.user(() => typed.api.statusCell('S4aA269005').setText('agreed'));   // a status with an option, typed as text
  typed.build();
  check('a text status is written from the Portal by a rebuild, as it always was', typed.shown('S4aA269005'), 'text:not treated');

  // Failures.
  const noRead = report();
  noRead.build();
  noRead.api.pick('S4aA269002', 'parked');
  let reads = 0;
  noRead.api.fail.get = (id) => (id === REPORT_ID && ++reads <= 2 ? 'synthetic read failure' : null);   // the read before the clear, masked and whole
  const built = noRead.build();
  check('the values cannot be read before the rebuild: the build completes with the statuses of the Portal, and says so in the log', [built.ok, noRead.shown('S4aA269002'), noRead.statusLogs().filter((l) => /could not be read \(synthetic read failure\)/.test(l)).length > 0], [true, 'dropdown:available', true]);

  const broken = report();
  broken.build();
  broken.api.pick('S4aA269002', 'parked');
  broken.s.statusDropdownIndex_ = ((real) => function () { if (broken.breakRead) { broken.breakRead = false; throw new Error('synthetic failure inside the read'); } return real(); })(broken.s.statusDropdownIndex_);
  broken.breakRead = true;
  const survived = broken.build();
  check('whatever goes wrong while the values are read before a rebuild, the build is not failed by it', [survived.ok, broken.shown('S4aA269002'), broken.statusLogs().filter((l) => l.indexOf('could not be read before the rebuild (synthetic failure inside the read)') !== -1).length], [true, 'dropdown:available', 1]);

  const noWrite = report();
  noWrite.build();
  noWrite.api.pick('S4aA269002', 'parked');
  noWrite.api.fail.batchUpdate = () => 'synthetic write failure';
  const text = noWrite.build();
  check('the dropdowns cannot be made after the rebuild: every status is the Portal\'s text, and none is emptied', [text.ok, noWrite.shown('S4aA269002'), noWrite.all().filter((x) => /dropdown/.test(x)).length], [true, 'text:available', 0]);

  // Where nothing is kept, because nothing is read.
  const old = report({ docs: false });
  old.build();
  old.s.Docs = { Documents: { get: (id, opts) => { old.events.push('get'); return old.api.Docs.Documents.get(id, opts); }, batchUpdate: (resource, id) => { old.events.push('batchUpdate'); return old.api.Docs.Documents.batchUpdate(resource, id); } } };
  old.exec(() => old.s.captureStatusDropdownValuesForRebuild_());
  check('a report that is not marked is not read before a rebuild', [old.events, old.s.statusDropdownRun_().previous], [[], null]);
  const central = report({ release: false });
  central.build(); central.build();
  check('CENTRAL and Legacy: nothing is read, nothing is kept, nothing changes', [central.events, central.all()[0]], [[], 'S4aA269001 text:agreed']);
  check('the values are read in one place: before the build clears the document (source)', [(CODE.match(/captureStatusDropdownValuesForRebuild_\(\)/g) || []).length,
    /  captureStatusDropdownValuesForRebuild_\(\);\n\n  \/\/ Step 2: Clear document, set title\n  const body = DocumentApp\.getActiveDocument\(\)\.getBody\(\)\.clear\(\);/.test(CODE)], [2, true]);
  check('nothing of it is stored: the feature writes its one property only, in its two places (source)', (CODE.slice(CODE.indexOf('// TDOC STATUS DROPDOWNS (T-2026.10.7)'), CODE.indexOf('// TDOC UPLOAD COMPLETION')).match(/\.setProperty\(([A-Z_a-z]*)/g) || []), ['.setProperty(STATUS_DROPDOWN_ENABLED_KEY_', '.setProperty(STATUS_DROPDOWN_ENABLED_KEY_']);
}

console.log(failures === 0 ? '\nAll status-dropdown checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
