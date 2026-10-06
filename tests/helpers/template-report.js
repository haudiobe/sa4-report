/**
 * One synthetic ad-hoc report on which the REAL Build Report from Scratch,
 * the REAL update, the REAL abstract sweep and the REAL status sync run,
 * against the fake document (fake-document.js) and the fake Docs API
 * (fake-docs-api.js). The same arrangement as tests/status-dropdown.test.js,
 * as a helper, with three things more:
 *   - a TDoc of the TDoc list can have the hyperlink that marks it as
 *     uploaded, and a "Revised to";
 *   - the Contribution Reviewer is a fake that records every request;
 *   - the revision placement is NOT replaced (it is what some tests are about).
 *
 * Only the network collectors (e-mail, revisions folder) and the formatter
 * are replaced. All names and numbers are synthetic.
 */

const { loadCode } = require('./load-code.js');
const { loadTemplateRuntime } = require('./load-template.js');
const { makeFakeDocumentBody } = require('./fake-document.js');
const { makeFakeDocsApi } = require('./fake-docs-api.js');

const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const STATUS_NAMES = ['available', 'noted', 'agreed', 'revised', 'parked', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'other', 'replied', 'Plenary', 'postponed'];
const colour = (n, shift) => ({ color: { rgbColor: { red: ((n * 17 + shift) % 100) / 100, green: ((n * 31 + shift) % 100) / 100, blue: ((n * 53 + shift) % 100) / 100 } } });
const STATUS_DEFINITION = { title: 'Document Status', options: STATUS_NAMES.map((name, n) => ({ displayValue: name, textStyle: { foregroundColor: colour(n, 3), backgroundColor: colour(n, 47) } })) };

const DOCS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio', MEETING_DATE: 'September 22, 2026',
  FTP_BASE: DOCS, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'], ['4', 'Topic'], ['4.3', 'Performance requirements'],
  ['4.4', 'Design constraints'], ['4.5', 'Other topics'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
const HEADER = ['TDoc', 'Title', 'Source', 'Contact', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Reservation date', 'Uploaded', 'Is revision of', 'Revised to'];
const zip = (id) => DOCS + id + '.zip';
const CACHE = (id) => 'REVIEWER_NO_SUMMARY_CACHE_' + id;

/**
 * options:
 *   tdocs     [{ id, agenda, status, revisedTo, uploaded }] -- the TDoc list; `uploaded: true` gives the TDoc its hyperlink
 *   props     more document properties (null removes one)
 *   token     false: no Reviewer token
 *   reviewer  { id: text | 404 | 500 } -- what the Reviewer answers (default 404)
 *   release   false: not a template report (so: text statuses, no Docs API use)
 *   docs      false: no Google Docs API service
 *   agenda    another agenda
 *   drive     a fake Drive of tests/helpers/fake-drive-files.js: the Drive of the user the report runs as
 */
function templateReport(options) {
  const o = options || {};
  const props = Object.assign({}, PROPS, o.props || {});
  Object.keys(props).forEach((k) => { if (props[k] === null) delete props[k]; });
  // With a Drive (the personal settings of the user) the report is a whole template bundle: Code.js and ReportCreator.js.
  const loaded = (o.drive ? loadTemplateRuntime : loadCode)({ documentProperties: props, scriptProperties: o.token === false ? {} : { REVIEWER_API_TOKEN: 'synthetic-test-token' } });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const api = makeFakeDocsApi({ body: body, reportId: REPORT_ID });
  api.addSource(TEMPLATE_ID, [STATUS_DEFINITION]);
  const r = { s: s, body: body, api: api, docProps: loaded.docProps, scriptProps: loaded.scriptProps, logs: [], requests: [], fetches: [], reviewer: Object.assign({}, o.reviewer || {}),
    tdocs: (o.tdocs || []).map((t) => Object.assign({}, t)) };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.ElementType.TEXT = 'TEXT';
  s.DocumentApp.ElementType.UNSUPPORTED = 'UNSUPPORTED';
  const doc = { getId: () => REPORT_ID, getBody: () => body, saveAndClose: () => { body._state.dirty = false; body._state.closed = true; } };
  const ui = { alerts: [], ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' }, alert: (...args) => { ui.alerts.push(args); return 'YES'; }, showModalDialog: () => {} };
  r.ui = ui;
  s.DocumentApp.getActiveDocument = () => doc;
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ui;
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  if (o.release !== false) s.SA4_RELEASE_ = { releaseId: 'T-2026.10.8', flavor: 'template', codeVersion: '2.20.0', templateDocumentId: TEMPLATE_ID };
  if (o.docs !== false) s.Docs = api.Docs;
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.removeRowHeightAndSpacing = () => {};
  s.styleStatusCell_ = () => {};
  s.downloadMeetingAgenda_ = () => o.agenda || AGENDA;
  s.collectEmailDiscussionCore_ = () => {};
  s.collectRevisionsCore_ = () => {};
  s.collectorUpdate_ = () => ({ failures: [] });
  s.UrlFetchApp = { fetch: (url) => {
    r.fetches.push(String(url));
    const m = String(url).match(/^https:\/\/reviewer\.bouazizi\.dev\/api\/v1\/documents\/([^/]+)\/summary\?type=summary$/);
    if (m) {
      r.requests.push(m[1]);
      const answer = Object.prototype.hasOwnProperty.call(r.reviewer, m[1]) ? r.reviewer[m[1]] : 404;
      if (typeof answer === 'number') return { getResponseCode: () => answer, getContentText: () => '' };
      return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ text: answer }) };
    }
    return { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) };
  } };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  if (o.drive) o.drive.install(s);
  r.drive = o.drive || null;
  s.SpreadsheetApp = { open: () => {
    const values = [HEADER].concat(r.tdocs.map((t) => [t.id, 'Synthetic title of ' + t.id, 'ExampleCorp', 'Sam Rivera', 'discussion', 'Agreement', t.agenda, 'Synthetic item', t.status || '', '', '', '', t.revisedTo || '']));
    const noLink = { getLinkUrl: () => null };
    const rich = [HEADER.map(() => noLink)].concat(r.tdocs.map((t) => HEADER.map((h, col) => (col === 0 && t.uploaded ? { getLinkUrl: () => zip(t.id) } : noLink))));
    const sheet = { getDataRange: () => ({ getValues: () => values, getRichTextValues: () => rich }),
      getRange: (row, col, rows) => ({ getDisplayValues: () => [['Uploaded']].concat(r.tdocs.map(() => [''])).slice(row - 1, row - 1 + rows) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };

  // The Reviewer token: how often it is looked up.
  r.tokenReads = 0;
  const realScriptProps = s.PropertiesService.getScriptProperties();
  s.PropertiesService.getScriptProperties = () => Object.assign({}, realScriptProps, {
    getProperty: (k) => { if (k === 'REVIEWER_API_TOKEN') r.tokenReads++; return realScriptProps.getProperty(k); }
  });
  // The "no summary" cache: every write and every removal.
  r.cacheWrites = [];
  const store = loaded.docProps;
  const realSet = store.setProperty;
  const realDelete = store.deleteProperty;
  store.setProperty = (k, v) => { if (/^REVIEWER_NO_SUMMARY_CACHE_/.test(k)) r.cacheWrites.push('set ' + k); return realSet(k, v); };
  store.deleteProperty = (k) => { if (/^REVIEWER_NO_SUMMARY_CACHE_/.test(k)) r.cacheWrites.push('delete ' + k); return realDelete(k); };

  /** One execution, as Apps Script starts it: the document open, nothing remembered. */
  r.exec = (fn) => { body._state.closed = false; s.resetStatusDropdownRun_(); s.resetReviewerTokenRunState_(); return fn(); };
  /** Somebody edits the document by hand, between two executions; the edit is saved. */
  r.user = (fn) => { body._state.closed = false; const out = fn(); body._state.dirty = false; return out; };
  r.build = () => r.exec(() => s.runFullReportBuildCore_());
  r.update = () => r.exec(() => s.continuousUpdateCore_());
  r.updateAbstracts = () => r.exec(() => s.addAbstractsOnly());
  /** What happened since `mark()`: Reviewer requests, cache writes, token look-ups. */
  r.mark = () => { r.requests.length = 0; r.cacheWrites.length = 0; r.tokenReads = 0; r.logs.length = 0; };
  r.tdoc = (id) => r.tdocs.filter((t) => t.id === id)[0];
  r.tdocTables = () => body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc');
  r.table = (id) => r.tdocTables().filter((t) => t.getRow(0).getCell(1).getText() === id)[0];
  r.field = (id, label) => { const t = r.table(id); if (!t) return undefined; for (let i = 0; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === label) return t.getRow(i).getCell(1).getText(); return undefined; };
  r.linkOf = (id) => { const t = r.table(id); return t ? t.getRow(0).getCell(1).editAsText().getLinkUrl(0) : undefined; };
  r.shown = (id) => api.shown(id);
  /** The TDoc numbers of the TDoc tables, in the order of the document. */
  r.order = () => r.tdocTables().map((t) => t.getRow(0).getCell(1).getText());
  /** The document as headings and TDoc numbers, in order: ['# 4.3 ...', 'S4aA269001', ...]. */
  r.outline = () => body._children.map((c) => {
    if (c.getType() === 'TABLE') return c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc' ? c.getRow(0).getCell(1).getText() : null;
    return c.getHeading && c.getHeading() !== 'NORMAL' ? '# ' + c.getText() : null;
  }).filter(Boolean);
  return r;
}

module.exports = { templateReport, zip, CACHE, PROPS, AGENDA, REPORT_ID, TEMPLATE_ID, DOCS };
