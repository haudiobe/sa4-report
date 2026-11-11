/**
 * Close Report's external check: the Portal's TDoc list page as fallback evidence (after T-2026.10.18) --
 * docs/DISPOSITION_SUMMARY.md § Close Report, "The external check".
 *
 * 6G, 2026-10-30: the Excel export of the meeting's TDoc list (GenerateDocumentList.aspx?meetingId=86178) served the 36 TDocs
 * of earlier calls (S4aP260008..043) while the Portal's TDoc list page (TdocList.aspx?meetingId=86178) showed the 116 of the
 * meeting. The identity guard refused the export; Close Report could verify none of the 39 revised TDocs.
 *
 * Close Report alone may now read that page -- only after the configured list could not be had or was refused by its
 * identity -- and only through the same guard, the same filtering and the same revision check. Nothing else reads it: not
 * the update, not the build, not the Disposition Summary.
 *
 *   A.  the parser on the real page of 86178 (tests/fixtures/meeting-86178-tdoclist-page-2026-10-30.html.txt): every row,
 *       the disabled rows, every number of "Revised To"; and what it refuses (fail closed)
 *   B.  the real case: the wrong Excel export (tests/fixtures/meeting-86178-tdoc-list-wrong-2026-10-30.json) and the real
 *       page, on a report built from the meeting's 116 TDocs -- the fallback verifies the revisions (083, 140, 074)
 *   C.  the conditions: Excel valid -> Excel, no page; Excel refused -> page; Excel unavailable -> page; both wrong -> not
 *       ready, zero mutation; a page with a plausible count but no agenda TDoc -> refused
 *   D.  MBS (no agenda TDoc): the anchor and the report bootstrap decide the page as they decide a list; nothing stored
 *   E.  not for the update, not for the summary: Update Report Now refuses the wrong list and never fetches the page;
 *       Update Disposition Summary fetches nothing
 *   L.  what changed in Code.js: against T-2026.10.18, one section and three functions of Close Report
 *
 * Run: node tests/close-report-portal-page-fallback.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');
const { codeWithoutAgendaWordConversion, codeWithoutCloseReportPageFallback, PAGE_FALLBACK_SECTION_START, PAGE_FALLBACK_SECTION_END, PAGE_FALLBACK_SECTION_FUNCTIONS, PAGE_FALLBACK_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log('  ok   ' + name);
  } else {
    failures++;
    console.log('  FAIL ' + name + '\n         expected ' + e + '\n         actual   ' + a);
  }
}

const ROOT = path.join(__dirname, '..');
// The captured page, as .html.txt: the release file-set guard admits no tracked .html file (tests/release-file-set.test.js).
const PAGE_86178 = fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86178-tdoclist-page-2026-10-30.html.txt'), 'utf8');
const WRONG_EXCEL_86178 = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86178-tdoc-list-wrong-2026-10-30.json'), 'utf8')).values;
const SUMMARY = 'Disposition Summary';
const AGENDA = 'S4aA260090';
const EXCEL = 'Excel';
const PAGE = 'Portal HTML fallback';
const id = (n) => 'S4aA2691' + String(n).padStart(2, '0');
const t = (n, agenda, status, extra) => Object.assign({ id: id(n), agenda: agenda, status: status || 'available', uploaded: true }, extra || {});
const cellLink = (cell) => cell.editAsText().getLinkUrl(0) || '';
const snapshot = (body) => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE'
  ? ['T', Array.from({ length: c.getNumRows() }, (x, r) => Array.from({ length: c.getRow(r).getNumCells() }, (y, k) => [c.getRow(r).getCell(k).getText(), c.getRow(r).getCell(k)._dropdown ? c.getRow(r).getCell(k)._dropdown.optionId : null, cellLink(c.getRow(r).getCell(k))]))]
  : ['P', c.getHeading(), c.getText()])));
const props = (r) => JSON.stringify(Object.keys(r.docProps._store).sort().map((k) => [k, r.docProps._store[k]]));
const listFetches = (r, from) => r.fetches.slice(from).filter((f) => /GenerateDocumentList/.test(f)).length;
const pageFetches = (r, from) => r.fetches.slice(from).filter((f) => /TdocList\.aspx/.test(f)).length;
const lastAlert = (r) => r.ui.alerts[r.ui.alerts.length - 1];
const revisionsOf = (res) => res.revisions.map((x) => [x.tdoc, x.kind, x.revision]);

/** Close Report shows its result in an HTML dialog; it is recorded with the alerts. */
function withDialogs(r) {
  r.s.HtmlService = { createHtmlOutput: (html) => { const o = { setWidth: () => o, setHeight: () => o, getContent: () => html }; return o; } };
  r.ui.showModalDialog = (output, title) => { r.ui.alerts.push([title, r.logs[r.logs.length - 1], 'dialog', output.getContent()]); };
  return r;
}
/** A built template report (ad-hoc, agenda TDoc S4aA260090) with status dropdowns. */
function built(tdocs, decisions, options) {
  const r = withDialogs(templateReport(Object.assign({ tdocs: tdocs, tableSeparators: true, endParagraph: true }, options || {})));
  const b = r.build();
  if (!b.ok) throw new Error('build failed: ' + (b.error || b.message));
  const u = r.update();
  if (!u.success) throw new Error('update failed: ' + u.error);
  r.user(() => Object.keys(decisions || {}).forEach((n) => r.api.pick(id(n), decisions[n])));
  return r;
}
const closeReport = (r) => r.exec(() => r.s.closeReport());
const updateSummary = (r) => r.exec(() => r.s.updateDispositionSummary());

/**
 * The Portal's TDoc list page for `rows` [{ id, status, agenda, topic, revisedTo: [numbers], revisionOf, uploaded, type }],
 * in the structure of the real page: the pager with "N contributions found, displaying 1 to N", the grid rgTdocList_ctl00,
 * one rgRow/rgAltRow of twelve cells per TDoc, the TDoc as a link to its document or as a disabled link, the agenda item in
 * span.agendaItem, every revision of "Revised To" as a link. `o.found` / `o.to` override the pager's numbers.
 */
function pageFor(rows, o) {
  o = o || {};
  const n = rows.length;
  const esc = (s) => String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
  const docLink = (x) => 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Synthetic/Docs/' + x + '.zip';
  const popup = (x) => '<a href="javascript:OpenContributionDetailsPopup(&#39;https://portal.3gpp.org/ngppapp/CreateTDoc.aspx?mode=view&amp;contributionUid=' + x + '&#39;, &#39;' + x + '&#39;);">' + x + '</a>';
  const tr = rows.map((r, i) => '<tr class="' + (i % 2 ? 'rgAltRow' : 'rgRow') + '" id="rgTdocList_ctl00__' + i + '">' +
    '<td><img id="rgTdocList_ctl00_ctl' + i + '_imgContributionId" title="Details of Contribution" src="../Images/details.png" /></td>' +
    '<td>' + (r.uploaded === false ? '<a disabled="disabled" target="_blank">' + r.id + '</a>' : '<a href="' + docLink(r.id) + '" target="_blank">' + r.id + '</a>') + '</td>' +
    '<td>' + esc(r.type || 'pCR') + '</td><td>' + esc(r.title || 'Synthetic title of ' + r.id) + '</td><td>ExampleCorp</td><td>' + esc(r.status || 'available') + '</td><td>Agreement</td>' +
    '<td>' + esc(o.meeting || 'Synthetic ad-hoc') + '</td>' +
    '<td><div> <span id="rgTdocList_ctl00_ctl' + i + '_lblMeetingAllocation" title="' + esc(r.topic || 'Synthetic item') + '" class="agendaItem">' + esc(r.agenda) + '</span></div></td>' +
    '<td>' + (r.revisionOf ? popup(r.revisionOf) : '<a></a>') + '</td>' +
    '<td>' + ((r.revisedTo || []).map(popup).join(' , ') || '') + '</td><td></td></tr>').join('\n');
  return '<!DOCTYPE html><html><head><title>TDoc List</title></head><body><form method="post" action="./TdocList.aspx?meetingId=' + (o.meetingId || '85916') + '" id="form1">' +
    '<input type="hidden" name="__VIEWSTATE" id="__VIEWSTATE" value="x" />' +
    '<div id="pnlTDocList"><div id="upDownloadTdocs"><a id="hlDownloadTdocs" href="GenerateDocumentList.aspx?meetingId=' + (o.meetingId || '85916') + '" target="_blank">Download Tdocs to Excel</a></div>' +
    '<div id="rgTdocList" class="RadGrid RadGrid_Default"><table class="rgMasterTable rgClipCells" border="0" id="rgTdocList_ctl00_TopPager" width="100%"><thead><tr class="' + (o.pagerShown ? 'rgPager' : 'hideDataPager rgPager') + '"><td colspan="12">' +
    '<table summary="Data pager"><tr><td><div class="rgWrap rgNumPart"><a class="rgCurrentPage" href="#"><span>1</span></a></div>' +
    '<div class="rgWrap rgInfoPart"> ' + (o.found === undefined ? n : o.found) + ' contributions found, displaying 1 to ' + (o.to === undefined ? n : o.to) + ' </div></td></tr></table></td></tr></thead></table>' +
    '<table class="rgMasterTable rgClipCells rgClipCells" border="0" id="rgTdocList_ctl00" width="100%"><thead><tr class="rgHeader"><th class="rgHeader">&nbsp;</th><th>Tdoc</th><th>Type</th><th>Title</th><th>Source</th><th>Status</th><th>For</th><th>Meeting</th><th>Agenda item</th><th>Revision of</th><th>Revised To</th><th>Extra info</th></tr></thead>' +
    '<tbody>' + tr + '</tbody></table></div></div></form></body></html>';
}
/** The Portal serves `html` (or an HTTP status) for TdocList.aspx; everything else as before. */
function servePage(r, html) {
  const fetch = r.s.UrlFetchApp.fetch;
  r.s.UrlFetchApp.fetch = (url, o) => {
    if (!/TdocList\.aspx/.test(String(url))) return fetch(url, o);
    r.fetches.push(String(url));
    if (typeof html === 'number') return { getResponseCode: () => html, getContentText: () => 'Service Unavailable' };
    return { getResponseCode: () => 200, getContentText: () => html };
  };
}
/** The Portal returns another meeting's Excel export: no row for the agenda TDoc, other TDocs. */
function wrongExcel(r, rows) {
  r.listAgenda = null;
  r.tdocs = rows || [t(41, '4.3'), t(42, '4.7', 'reserved', { uploaded: false })];
}
function excelBroken(r) {
  r.s.SpreadsheetApp = { open: () => { throw new Error('Spreadsheet conversion failed'); } };
}
/** The page of the meeting's own list: `tdocs` of the harness, plus the agenda TDoc, as the real page has them. */
const pageOfList = (tdocs, agendaTdoc, o) => pageFor((agendaTdoc ? [{ id: agendaTdoc, type: 'agenda', agenda: '2', status: 'available' }] : []).concat(tdocs.map((x) => ({
  id: x.id, agenda: x.agenda, status: x.status, uploaded: x.uploaded !== false, revisedTo: x.revisedTo ? [x.revisedTo] : [] }))), o);

// The meeting of sections C and E: 4.3: 01 agreed, 02 revised -> 03 (agreed), 4.4: 04 noted, 05 revised -> 06 (reserved, not uploaded).
const MEETING = () => [t(1, '4.3'), t(2, '4.3', 'revised', { revisedTo: id(3) }), t(3, '4.3'), t(4, '4.4'), t(5, '4.4', 'revised', { revisedTo: id(6) }), t(6, '4.4', 'reserved', { uploaded: false })];
const DECISIONS = { 1: 'agreed', 3: 'agreed', 4: 'noted' };
const CLEAN = () => [t(1, '4.3'), t(2, '4.3', 'revised', { revisedTo: id(3) }), t(3, '4.3')];
const CLEAN_DECISIONS = { 1: 'agreed', 3: 'agreed' };

// ====================================================================== A
console.log('A. the parser on the real page of 86178, and what it refuses');
{
  const { sandbox: s } = loadCode();
  const p = s.parsePortalTdocListPage_(PAGE_86178);
  const by = {};
  p.rows.forEach((r) => { by[r.tdoc] = r; });
  const all = Array.from({ length: 116 }, (x, i) => 'S4aP' + (260067 + i));
  check('the real page: one complete list of 116 contributions, 116 rows, S4aP260067..182, each once', [p.ok, p.error, p.total, p.rows.length, all.filter((n) => !by[n])], [true, '', 116, 116, []]);
  check('the TDocs the Portal shows as reserved -- disabled links -- are rows too: 10, among them 099, 172, 173, 174, 175',
    [p.rows.filter((r) => !r.uploaded).length, ['S4aP260099', 'S4aP260172', 'S4aP260173', 'S4aP260174', 'S4aP260175'].map((n) => by[n] ? by[n].status + '/' + (by[n].uploaded ? 'uploaded' : 'not uploaded') : 'missing')],
    [10, ['withdrawn/not uploaded', 'withdrawn/not uploaded', 'reserved/not uploaded', 'withdrawn/not uploaded', 'reserved/not uploaded']]);
  check('"Revised To" is the whole cell: S4aP260140 names 174 AND 175; 074 names 099 AND 148', [by.S4aP260140.revisedTo, by.S4aP260074.revisedTo], ['S4aP260174 , S4aP260175', 'S4aP260099 , S4aP260148']);
  check('S4aP260098, the agenda TDoc: an agenda-typed row, uploaded, item 5.1', [by.S4aP260098.type, by.S4aP260098.uploaded, by.S4aP260098.agendaItem, /S4aP260098\.zip$/.test(by.S4aP260098.url)], ['agenda', true, '5.1', true]);
  check('every row has an agenda item and a status; the statuses are 3GU\'s', [p.rows.filter((r) => !r.agendaItem || !r.status).length, p.rows.reduce((a, r) => (a[r.status] = (a[r.status] || 0) + 1, a), {})],
    [0, { available: 55, reserved: 7, withdrawn: 3, agreed: 5, approved: 1, revised: 38, noted: 6, merged: 1 }]);
  const rows = s.portalTdocListPageRows_(p);
  check('as raw rows: the header of the Excel export\'s names, 116 rows; the link cell known for every TDoc (a link, or none for a reserved one)',
    [rows.data[0], rows.data.length - 1, rows.richTextValues[rows.data.findIndex((x) => x[0] === 'S4aP260148')][0].getLinkUrl(), rows.richTextValues[rows.data.findIndex((x) => x[0] === 'S4aP260175')][0].getLinkUrl()],
    [['TDoc', 'Title', 'Source', 'Type', 'For', 'Agenda item', 'Agenda item description', 'TDoc Status', 'Is revision of', 'Revised to', 'Uploaded'], 116, 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/S4aP260148.zip', null]);
  const cfg = { AGENDA_TDOC: 'S4aP260098', MEETING_ID: '86178', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178' };
  check('the SAME identity guard decides it: the page has the row of S4aP260098 (agenda); the wrong export has not',
    [s.evaluateTdocListIdentity_(rows.data, cfg, null, [], 'read'), /no row for S4aP260098, the agenda TDoc/.test(s.evaluateTdocListIdentity_(WRONG_EXCEL_86178, cfg, null, [], 'read').error)],
    [{ ok: true, via: 'agenda', anchor: null, error: '' }, true]);
  check('the wrong export: 36 TDocs, S4aP260008..043, none of the meeting\'s', (() => { const c = WRONG_EXCEL_86178[0].indexOf('TDoc'); const ids = WRONG_EXCEL_86178.slice(1).map((x) => x[c]).filter(Boolean).sort(); return [ids.length, ids[0], ids[ids.length - 1], ids.filter((n) => all.indexOf(n) !== -1).length]; })(), [36, 'S4aP260008', 'S4aP260043', 0]);

  // Fail closed: nothing is read from a page that is not shown to be one complete list.
  const fail = (html) => { const q = s.parsePortalTdocListPage_(html); return q.ok ? 'ACCEPTED' : q.error; };
  const rowsOf = (n) => Array.from({ length: n }, (x, i) => ({ id: n > 99 ? 'S4aA27' + String(i + 1).padStart(4, '0') : id(i + 1), agenda: '4.3' }));
  check('a login page, an empty answer, a page without the grid: refused', [fail('<html><body><form action="Login.aspx">Sign in</form></body></html>'), fail(''), fail('<html><body><p>An error occurred.</p></body></html>')],
    ['the page asks for a login', 'the page has no TDoc list grid (a login or error page, or another layout)', 'the page has no TDoc list grid (a login or error page, or another layout)']);
  check('a grid without the pager\'s count, or with none found: refused', [fail(pageFor(rowsOf(3)).replace(/class="rgWrap rgInfoPart">[^<]*</, 'class="rgWrap rgInfoPart"><')), fail(pageFor([], { found: 0, to: 0 }))],
    ['the page does not say how many contributions it found', 'the page lists no contribution']);
  check('a paged or truncated list: refused -- shown 1 to 50 of 116; 116 found but 115 rows', [fail(pageFor(rowsOf(50), { found: 116, to: 50, pagerShown: true })), fail(PAGE_86178.replace(/<tr[^>]*class="rgAltRow"[^>]*>[\s\S]*?<\/tr>/, ''))],
    ['the page shows 1 to 50 of 116 contributions: not the whole list', 'the page says 116 contributions but shows 115 rows']);
  check('a TDoc twice, a row without the twelve cells, a cell that is no TDoc number, a TDoc cell that is neither a link nor a disabled link: refused',
    [fail(pageFor(rowsOf(3).concat([{ id: id(2), agenda: '4.4' }]))), fail(pageFor(rowsOf(3)).replace(/<td><\/td><\/tr>/, '</tr>')), fail(pageFor(rowsOf(3)).replace('>' + id(2) + '</a>', '>S4aA26910</a>')),
      fail(pageFor(rowsOf(3)).replace(/<a href="[^"]*" target="_blank">(S4aA269102)<\/a>/, '$1'))],
    ['the page lists S4aA269102 twice', 'row 1 has 11 cells instead of 12', 'row 2 has no TDoc number ("S4aA26910")', 'the TDoc cell of S4aA269102 is neither a link nor a disabled link']);
  check('a count alone is no identity: the synthetic page of 116 other TDocs parses, and the guard refuses it for want of the agenda TDoc',
    (() => { const q = s.parsePortalTdocListPage_(pageFor(rowsOf(116))); return [q.ok, q.total, /no row for S4aP260098/.test(s.evaluateTdocListIdentity_(s.portalTdocListPageRows_(q).data, cfg, null, [], 'read').error)]; })(), [true, 116, true]);
  check('the agenda-discovery parser is not this one, and is unchanged in what it does: 106 rows of the real page (it skips the disabled links), the first revision only',
    (() => { const old = s.parseMeetingTdocListHtml_(PAGE_86178); return [old.length, (old.filter((x) => x.id === 'S4aP260140')[0] || {}).revisedTo]; })(), [106, 'S4aP260174']);
}

// ====================================================================== B
console.log('B. the real case: the wrong export, the real page, a report of the meeting\'s 116 TDocs');
{
  // The report is built from the page's own rows as its list (the shape the harness's fake sheet takes), with the Portal's
  // statuses (the update synchronises them), under an agenda of the meeting's items. 099 is then reallocated to removed, as
  // in live 6G (its table removed by hand here; its registration row stays).
  const { sandbox: pure } = loadCode();
  const parsed = pure.parsePortalTdocListPage_(PAGE_86178);
  const items = [['5', 'FS_6G_MED'], ['5.1', 'Opening of the session and registration of documents'], ['5.3', 'Reports/Liaisons from other groups/meetings'], ['5.4', 'FS_6G_MED - General and working documents'],
    ['5.5', 'FS_6G_MED - WT#1'], ['5.6', 'FS_6G_MED - WT#2'], ['5.6.1', 'WT#2.1'], ['5.6.2', 'WT#2.2'], ['5.6.3', 'WT#2.3'], ['5.7', 'WT#3'], ['5.8', 'WT#4'], ['5.9', 'WT#5'], ['5.10', 'Other issues'], ['5.11', 'Close of the session']]
    .map(([number, title]) => ({ number: number, title: title, level: number.split('.').length, heading: 'NORMAL', text: '' }));
  const list = parsed.rows.map((r) => ({ id: r.tdoc, agenda: r.agendaItem, status: r.status, uploaded: r.uploaded, revisedTo: r.revisedTo || '', type: r.type }));
  const r = withDialogs(templateReport({ tdocs: list, tableSeparators: true, endParagraph: true, agenda: items,
    props: { MEETING_ID: '86178', MEETING_NAME: 'SA4-e (AH) on FS_6G_MED', REPORT_SUFFIX: '6G Media', AGENDA_TDOC: 'S4aP260098', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178' } }));
  const b = r.build();
  if (!b.ok) throw new Error('build failed: ' + (b.error || b.message));
  const u = r.update();
  if (!u.success) throw new Error('update failed: ' + u.error);
  r.user(() => {
    r.body.appendParagraph('Document Reallocations');
    r.body.appendTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason'], ['S4aP260099', '5.4', 'removed', 'Reserved document not used']]);
    r.body.removeChild(r.table('S4aP260099'));
  });
  const revisedInReport = r.order().filter((n) => /revised/.test(r.shown(n)));
  const scope = r.exec(() => r.s.reportTdocScope_(r.body));
  check('the report as live 6G: 115 TDoc tables, 116 registered, 099 reallocated to removed, scope 115; 38 revised', [r.order().length, Object.keys(scope.numbers).length, scope.numbers.S4aP260099, revisedInReport.length],
    [115, 115, undefined, 38]);

  // The Portal now: the export is the wrong list; the page is the real one.
  r.s.SpreadsheetApp = { open: () => {
    const noLink = { getLinkUrl: () => null };
    const sheet = { getDataRange: () => ({ getValues: () => WRONG_EXCEL_86178, getRichTextValues: () => WRONG_EXCEL_86178.map((row) => row.map(() => noLink)) }),
      getRange: () => ({ getDisplayValues: () => WRONG_EXCEL_86178.map(() => ['']) }) };
    return { getSheets: () => [sheet], getSpreadsheetTimeZone: () => 'UTC' };
  } };
  servePage(r, PAGE_86178);
  const doc = snapshot(r.body);
  const store = props(r);
  const f = r.fetches.length;
  const res = closeReport(r);
  check('Close Report: the export refused, the page read once, shown to be the meeting\'s by S4aP260098, and used -- external verification completed, from the Portal HTML fallback',
    [listFetches(r, f), pageFetches(r, f), res.externalVerified, res.externalSource, res.checked, r.logs.some((l) => /^TDoc list identity: REFUSED -- .*no row for S4aP260098/.test(l)),
      r.logs.some((l) => /^External TDoc verification source: Portal HTML fallback \(TdocList\.aspx\?meetingId=86178, 116 contributions, 115 TDocs of the meeting\)/.test(l))],
    [1, 1, true, PAGE, 115, true, true]);
  check('no revised TDoc is "TDoc list not read": every one is judged against the page', res.revisions.filter((x) => x.state3gu === 'TDoc list not read').length, 0);
  const of = (n) => revisionsOf(res).filter((x) => x[0] === n).map((x) => x.slice(1));
  check('083 -> 172 withdrawn: NOT RESERVED (no live revision)', [of('S4aP260083'), (res.revisions.filter((x) => x.tdoc === 'S4aP260083')[0] || {}).reason], [[['revision-not-reserved', 'S4aP260172']], 'No live revision is reserved in 3GU; revision S4aP260172 is withdrawn']);
  check('140 -> 174 withdrawn, 175 reserved not uploaded: RESERVED BUT NOT UPLOADED, naming 175', of('S4aP260140'), [['revision-not-uploaded', 'S4aP260175']]);
  check('074 -> 099 (reallocated to removed: not in the list) and 148 agreed, uploaded: complete -- no finding', of('S4aP260074'), []);
  check('the summary and the TDocs not yet disposed are the report\'s; the dialog says NOT READY and names the source',
    [res.ready, res.summaryState, lastAlert(r)[0], /External verification \(TDoc list of this meeting, 3GU\): completed\. Source: the Portal TDoc list page \(TdocList\.aspx\); the configured TDoc list could not be used\./.test(lastAlert(r)[3])],
    [false, 'created', 'Report NOT ready to close', true]);
  check('nothing but the summary was written; no property changed (no anchor: the meeting has an agenda TDoc, and a reader stores nothing)',
    [props(r) === store, r.exec(() => r.s.findDispositionSummarySection_(r.body)) !== null, (() => { const s2 = r.exec(() => r.s.findDispositionSummarySection_(r.body)); const all = JSON.parse(snapshot(r.body)); return JSON.stringify(all.slice(0, s2.start)) === JSON.stringify(JSON.parse(doc).slice(0, s2.start)); })()],
    [true, true, true]);
  check('run again: the same result, the summary unchanged', (() => { const again = closeReport(r); return [again.externalSource, again.summaryState, revisionsOf(again).length === revisionsOf(res).length]; })(), [PAGE, 'unchanged', true]);
  // The two other entries never read the page.
  const f2 = r.fetches.length;
  updateSummary(r);
  check('Update Disposition Summary on the same report: nothing fetched, the summary already up to date', [r.fetches.length - f2, lastAlert(r)[1].split('\n')[0]], [0, 'Disposition Summary was already up to date.']);
  const doc2 = snapshot(r.body);
  const u2 = r.update();
  check('Update Report Now on the same report: the wrong export refused before anything, the page never fetched, the document identical',
    [u2.success, u2.notUpdated, /no row for S4aP260098/.test(u2.error), pageFetches(r, f2), snapshot(r.body) === doc2, props(r) === store], [false, true, true, 0, true, true]);
}

// ====================================================================== C
console.log('C. the conditions: Excel valid; Excel refused; Excel unavailable; both wrong; a plausible count without the agenda TDoc');
{
  // C1. Excel valid: used as before; the page is never asked for.
  const a = built(MEETING(), DECISIONS);
  servePage(a, pageOfList(MEETING(), AGENDA));
  let f = a.fetches.length;
  let res = closeReport(a);
  check('Excel valid: source Excel, no page fetched; the dialog line as before', [res.externalVerified, res.externalSource, listFetches(a, f), pageFetches(a, f), revisionsOf(res), /3GU\): completed\.<\/div>/.test(lastAlert(a)[3])],
    [true, EXCEL, 1, 0, [[id(5), 'revision-not-uploaded', id(6)]], true]);
  check('... logged', a.logs.some((l) => l === 'External TDoc verification source: Excel (the configured TDoc list)'), true);

  // C2. Excel refused (another meeting's list): the page, through the same guard.
  wrongExcel(a);
  servePage(a, pageOfList(MEETING(), AGENDA));
  f = a.fetches.length;
  res = closeReport(a);
  check('Excel refused by its identity, the page the meeting\'s: list once, page once, verified from the page, the same revision findings',
    [listFetches(a, f), pageFetches(a, f), res.externalVerified, res.externalSource, revisionsOf(res), a.logs.some((l) => /^External TDoc verification source: Portal HTML fallback \(TdocList\.aspx\?meetingId=85916, 7 contributions, 7 TDocs of the meeting\) -- the TDoc list of this meeting could not be read: The TDoc list was not used/.test(l))],
    [1, 1, true, PAGE, [[id(5), 'revision-not-uploaded', id(6)]], true]);
  const clean = built(CLEAN(), CLEAN_DECISIONS);
  wrongExcel(clean);
  servePage(clean, pageOfList(CLEAN(), AGENDA));
  check('... and a clean report is READY on the page\'s evidence', [closeReport(clean).ready, closeReport(clean).externalSource], [true, PAGE]);

  // C3. Excel unavailable (conversion failure; HTTP failure): the page.
  const c = built(CLEAN(), CLEAN_DECISIONS);
  excelBroken(c);
  servePage(c, pageOfList(CLEAN(), AGENDA));
  res = closeReport(c);
  check('the export cannot be converted, the page the meeting\'s: READY, verified from the page', [res.ready, res.externalVerified, res.externalSource], [true, true, PAGE]);
  const d = built(CLEAN(), CLEAN_DECISIONS);
  const fetch = d.s.UrlFetchApp.fetch;
  d.s.UrlFetchApp.fetch = (url, o) => (/GenerateDocumentList/.test(String(url)) ? (d.fetches.push(String(url)), { getResponseCode: () => 503, getContentText: () => '' }) : fetch(url, o));
  servePage(d, pageOfList(CLEAN(), AGENDA));
  res = closeReport(d);
  check('the export HTTP 503, the page the meeting\'s: READY, verified from the page', [res.ready, res.externalSource], [true, PAGE]);

  // C4. Both wrong: not ready, both reasons, zero mutation but the summary.
  const e = built(MEETING(), DECISIONS);
  updateSummary(e);
  wrongExcel(e);
  servePage(e, pageOfList([t(41, '4.3'), t(42, '4.7')], null, { meetingId: '85000' }));
  const doc = snapshot(e.body);
  const store = props(e);
  res = closeReport(e);
  check('Excel wrong and page wrong: NOT READY, external verification not completed, both reasons given, every revised TDoc not verified, nothing changed',
    [res.ready, res.externalVerified, res.externalSource, /^External verification could not be completed: the TDoc list of this meeting could not be read: The TDoc list was not used: .*no row for S4aA260090.*; the Portal TDoc list page \(TdocList\.aspx\) could not be used either: The TDoc list was not used: .*no row for S4aA260090/.test(res.problems[0] || ''),
      revisionsOf(res).map((x) => x[1]), snapshot(e.body) === doc, props(e) === store],
    [false, false, '', true, ['revision-not-verified', 'revision-not-verified'], true, true]);
  // The page unavailable, or not a list.
  servePage(e, 503);
  check('the page HTTP 503: NOT READY, "HTTP 503" named', [closeReport(e).ready, /could not be used either: HTTP 503\./.test(closeReport(e).problems[0])], [false, true]);
  servePage(e, '<html><body><form action="Login.aspx">Sign in</form></body></html>');
  check('the page a login page: NOT READY, said so', /could not be used either: the page asks for a login\./.test(closeReport(e).problems[0]), true);
  servePage(e, pageOfList(MEETING(), AGENDA, { found: 7, to: 5, pagerShown: true }));
  check('the page paged: NOT READY, said so', /could not be used either: the page shows 1 to 5 of 7 contributions: not the whole list\./.test(closeReport(e).problems[0]), true);

  // C5. A plausible count without the agenda TDoc: the page has the report's six TDocs and more, but not S4aA260090.
  servePage(e, pageOfList(MEETING().concat([t(7, '4.5')]), null));
  res = closeReport(e);
  check('a page with every TDoc of the report, 7 contributions, but no row for the agenda TDoc: refused by the guard, NOT READY', [res.ready, res.externalVerified, /could not be used either: The TDoc list was not used: .*no row for S4aA260090/.test(res.problems[0])], [false, false, true]);
  check('a wrong page with TDocs in common with the report is refused too (the zero-overlap guard is not the identity)', (() => { servePage(e, pageOfList([t(1, '4.3', 'withdrawn'), t(4, '4.4', 'agreed'), t(41, '4.3')], null)); const x = closeReport(e); return [x.ready, x.externalVerified]; })(), [false, false]);
}

// ====================================================================== D
console.log('D. MBS, no agenda TDoc: the anchor and the report bootstrap decide the page as they decide a list; nothing stored');
{
  const withoutAgendaTdoc = (r) => {
    r.docProps.deleteProperty('AGENDA_TDOC');
    r.docProps.setProperty('AGENDA_CSV_URL', r.s.deriveAdhocAgendaCsvCandidate_(r.docProps.getProperty('FTP_BASE')).url);
    r.listAgenda = null;
  };
  const anchorOf = (r) => { const v = r.docProps.getProperty('TDOC_LIST_ANCHOR'); return v ? JSON.parse(v) : null; };
  // D1. The anchor recorded by the first update; a wrong export without its row refused; the page with it accepted via the anchor.
  const r = built(CLEAN(), CLEAN_DECISIONS);
  withoutAgendaTdoc(r);
  check('the first update without an agenda TDoc records the anchor', [r.update().success, (anchorOf(r) || {}).tdoc], [true, id(1)]);
  const store = props(r);
  r.tdocs = [t(2, '4.3', 'withdrawn'), t(41, '4.3')];
  servePage(r, pageOfList(CLEAN(), null));
  let res = closeReport(r);
  check('the export without the anchor refused; the page with the anchor\'s row accepted (anchor); READY; nothing stored', [res.ready, res.externalSource, props(r) === store, r.logs.filter((l) => /^TDoc list identity: the list of this meeting \(anchor\)$/.test(l)).length > 0], [true, PAGE, true, true]);
  servePage(r, pageOfList([t(2, '4.3'), t(3, '4.3'), t(41, '4.3')], null));
  res = closeReport(r);
  check('a page without the anchor\'s row: refused for want of the anchor, NOT READY', [res.ready, /could not be used either: .*no row for S4aA269101, a TDoc of the list this meeting had before/.test(res.problems[0])], [false, true]);
  // D2. No anchor yet: the page must account for every TDoc table of the report; a reader records nothing.
  const q = built(CLEAN(), CLEAN_DECISIONS);
  withoutAgendaTdoc(q);
  const before = props(q);
  q.tdocs = [t(41, '4.3')];
  servePage(q, pageOfList(CLEAN(), null));
  res = closeReport(q);
  check('no anchor yet, the page has every TDoc table of the report: accepted (report), READY, and the anchor NOT recorded by a reader', [res.ready, res.externalSource, anchorOf(q), props(q) === before], [true, PAGE, null, true]);
  servePage(q, pageOfList([t(1, '4.3'), t(2, '4.3')], null));
  res = closeReport(q);
  check('no anchor yet, the page lacks one of the report\'s TDocs: refused, NOT READY', [res.ready, /could not be used either: .*1 of the 3 TDocs of the report's TDoc tables are not in it \(S4aA269103\)/.test(res.problems[0])], [false, true]);
}

// ====================================================================== E
console.log('E. not for the update, not for the summary');
{
  const r = built(MEETING(), DECISIONS);
  updateSummary(r);
  wrongExcel(r);
  servePage(r, pageOfList(MEETING(), AGENDA));
  const doc = snapshot(r.body);
  const store = props(r);
  let f = r.fetches.length;
  const u = r.update();
  check('Update Report Now, the export refused: not updated, the page never fetched although it would pass, the document and its properties identical',
    [u.success, u.notUpdated, /no row for S4aA260090/.test(u.error), listFetches(r, f), pageFetches(r, f), snapshot(r.body) === doc, props(r) === store], [false, true, true, 1, 0, true, true]);
  excelBroken(r);
  f = r.fetches.length;
  const u2 = r.update();
  check('Update Report Now, the export unavailable: failed, the page never fetched, nothing changed', [u2.success, pageFetches(r, f), snapshot(r.body) === doc], [false, 0, true]);
  f = r.fetches.length;
  r.user(() => r.api.pick(id(4), 'agreed'));
  updateSummary(r);
  check('Update Disposition Summary: nothing fetched at all, the summary updated from the report', [r.fetches.length - f, lastAlert(r)[1].split('\n')[0]], [0, 'Disposition Summary updated.']);
  // The source without its comments: the calls, not what the comments say about them.
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '').replace(/^\s*(\/\/|\*|\/\*\*).*$/gm, '');
  const fn = (name) => { const at = src.indexOf('\nfunction ' + name + '('); return src.slice(at, src.indexOf('\n}\n', at)); };
  check('in the source: readPortalTdocListPageScope_ is called by readMeetingTdocScope_ only, and readMeetingTdocScope_ by closeReportCore_ only',
    [(src.match(/readPortalTdocListPageScope_\(/g) || []).length, /readPortalTdocListPageScope_\(cfg, excelProblem\)/.test(fn('readMeetingTdocScope_')), (src.match(/readMeetingTdocScope_\(\);/g) || []).length, /= readMeetingTdocScope_\(\);/.test(fn('closeReportCore_'))],
    [2, true, 1, true]);
  check('PORTAL_TDOC_LIST_PAGE_URL_ is fetched in one place, and the update, the build and the legacy import name neither it nor the parser',
    [(src.match(/PORTAL_TDOC_LIST_PAGE_URL_/g) || []).length, ['continuousUpdateCore_', 'downloadAndGroupTdocs_', 'downloadAndProcessFromWeb', 'buildSkeletonWithTdocTables', 'updateDispositionSummary', 'writeDispositionSummary_', 'syncDispositionSummaryOnUpdate_']
      .filter((name) => /PORTAL_TDOC_LIST_PAGE|parsePortalTdocListPage_|readPortalTdocListPageScope_|TdocList\.aspx/.test(fn(name)))], [2, []]);
}

// ====================================================================== L
console.log('L. what changed in Code.js: against T-2026.10.18, one section and three functions of Close Report');
{
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutCloseReportPageFallback(FULL);
  // After T-2026.10.19: the agenda Word conversion came later and has its own ledger (tests/agenda-word-conversion.test.js);
  // the comparisons below are made without it.
  const CODE = codeWithoutAgendaWordConversion(FULL).code;
  check('every change is in Code.js exactly once', view.problems, []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(PAGE_FALLBACK_SECTION_START), FULL.indexOf(PAGE_FALLBACK_SECTION_END) + PAGE_FALLBACK_SECTION_END.length);
  check('the section: these functions, in this order', functionNames(section), PAGE_FALLBACK_SECTION_FUNCTIONS);
  check('it stands directly after the TDoc list identity section', FULL.indexOf('// END OF TDOC LIST IDENTITY\n// =========================================================\n\n' + PAGE_FALLBACK_SECTION_START) !== -1, true);
  check('it declares its constants and nothing else at top level', section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || [],
    ['const PORTAL_TDOC_LIST_PAGE_URL_', 'const PORTAL_TDOC_LIST_PAGE_CELLS_', 'const PORTAL_TDOC_LIST_PAGE_HEADER_', 'const EXTERNAL_SOURCE_EXCEL_', 'const EXTERNAL_SOURCE_PORTAL_PAGE_']);
  let released = null;
  try {
    released = execFileSync('git', ['show', 'template-release/T-2026.10.18:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    released = null;
  }
  if (released === null) {
    console.log('  note: template-release/T-2026.10.18 is not available in this checkout; the comparison with it is skipped.');
  } else {
    check('without it, Code.js is the Code.js of T-2026.10.18, byte for byte', view.code === released, true);
    check('the functions that changed: readMeetingTdocScope_ (the fallback), closeReportCore_ (the source), externalVerificationText_ (its line)',
      functionNames(released).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(CODE, n) !== functionSource(released, n)).sort(), PAGE_FALLBACK_CHANGED_FUNCTIONS.slice().sort());
    check('the new functions are those of the section; none removed', [functionNames(CODE).filter((n) => functionNames(released).indexOf(n) === -1).sort(), functionNames(released).filter((n) => functionNames(CODE).indexOf(n) === -1)],
      [PAGE_FALLBACK_SECTION_FUNCTIONS.slice().sort(), []]);
    const identity = (src) => src.slice(src.indexOf('// TDOC LIST IDENTITY (after T-2026.10.16)'), src.indexOf('// END OF TDOC LIST IDENTITY'));
    check('untouched: the identity section; the update, the build, the legacy import; the summary and its scope; the revision check; the agenda-discovery parser',
      [identity(FULL) === identity(released)].concat(['continuousUpdateCore_', 'downloadAndGroupTdocs_', 'downloadAndProcessFromWeb', 'buildSkeletonWithTdocTables', 'updateDispositionSummary', 'writeDispositionSummary_', 'syncDispositionSummaryOnUpdate_',
        'reportTdocScope_', 'meetingTdocScope_', 'checkRevisionsIn3gu_', 'parseMeetingTdocListHtml_', 'extractTdocListRowCells_', 'assertTdocListIdentity_', 'evaluateTdocListIdentity_'].map((n) => functionSource(CODE, n) === functionSource(released, n))), Array(15).fill(true));
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll Close Report Portal-page fallback checks passed.');
process.exitCode = failures ? 1 : 0;
