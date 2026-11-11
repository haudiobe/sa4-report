/**
 * The TDoc-list download after T-2026.10.21 (MBS SA4#138, 2026-11-11).
 *
 * With T-2026.10.21, Full Report Build of the SA4#138 MBS report read the
 * agenda (by the Docs API) and stopped at "Could not download TDOC list:
 * HTTP 403". The configured list is on ftp.3gpp.org, which answers Apps
 * Script with 403; the same path under https://www.3gpp.org/ftp/ answers
 * HTTP 200 with the workbook (504b0304). The build fetched only the
 * configured address and took any answer below HTTP 400 as the workbook.
 * Now downloadTdocListXlsx_() asks the configured address, then (for
 * ftp.3gpp.org only) its www.3gpp.org/ftp twin, and accepts only HTTP 2xx
 * with the XLSX signature.
 *
 *   A. the configured address works: one request, the report as T-2026.10.21 builds it;
 *   B. it answers 403, the alternative works: the build succeeds, the same report
 *      (T-2026.10.21 fails with the production error); the configured address stays
 *      the list's identity;
 *   C. it answers HTTP 200 with an HTML page: refused, the alternative is asked;
 *   D. both fail: one error naming each; nothing converted, nothing written;
 *   E. the alternative answers with something that is not a workbook (HTML, empty,
 *      other bytes): refused;
 *   F. addresses that are not on ftp.3gpp.org: no alternative, the errors of before;
 *   G. the candidates, pure;
 *   H. diagnoseTdocListDownload(): every address, the build's check, nothing written;
 *   I. the update and Close Report download the same way;
 *   J. what changed in Code.js: one section and one function, against T-2026.10.21.
 *
 * Fakes only: no network, no Google account. Run: node tests/tdoc-list-download.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');
const { codeWithoutTdocListDownload, TDOC_LIST_DOWNLOAD_SECTION_START, TDOC_LIST_DOWNLOAD_SECTION_END, TDOC_LIST_DOWNLOAD_SECTION_FUNCTIONS, TDOC_LIST_DOWNLOAD_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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
const LIST_PATH = 'tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx';
const FTP_URL = 'https://ftp.3gpp.org/' + LIST_PATH;
const WWW_URL = 'https://www.3gpp.org/ftp/' + LIST_PATH;
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const SECRET = 'SyntheticAccessDeniedMarker';
const HTML = Array.from(Buffer.from('<!DOCTYPE html><html><body>' + SECRET + '</body></html>'));

const XLSX = { status: 200, bytes: [0x50, 0x4B, 0x03, 0x04, 0x14, 0x00], type: XLSX_TYPE };
const FORBIDDEN = { status: 403, bytes: HTML, type: 'text/html' };
const HTML_200 = { status: 200, bytes: HTML, type: 'text/html' };

let RELEASED = null;
try {
  RELEASED = execFileSync('git', ['show', 'template-release/T-2026.10.21:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8');
} catch (e) {
  RELEASED = null;
}

const t = (n, agenda) => ({ id: 'S4aA2691' + String(n).padStart(2, '0'), agenda: agenda, status: 'available', uploaded: true });
const TDOCS = [t(1, '4.3'), t(2, '4.3'), t(3, '4.4'), t(4, '4.5')];
const snapshot = (body) => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE'
  ? ['T', Array.from({ length: c.getNumRows() }, (x, r) => Array.from({ length: c.getRow(r).getNumCells() }, (y, k) => c.getRow(r).getCell(k).getText()))]
  : ['P', c.getText()])));

/**
 * A template report whose TDoc list is configured at `url`; `answers` maps an
 * address to { status, bytes, type } or 'throw'. Other fetches go to the helper.
 */
function report(answers, options) {
  const o = options || {};
  const r = templateReport({ tdocs: TDOCS, code: o.code, props: Object.assign({ TDOC_LIST_URL: o.url || FTP_URL }, o.props || {}) });
  const helperFetch = r.s.UrlFetchApp.fetch;
  r.listFetches = [];
  r.converted = 0;
  r.s.UrlFetchApp = { fetch: (url, params) => {
    if (!Object.prototype.hasOwnProperty.call(answers, url)) return helperFetch(url, params);
    r.listFetches.push(url);
    const a = answers[url];
    if (a === 'throw') throw new Error('Address unavailable: ' + url);
    return { getResponseCode: () => a.status, getBlob: () => ({ getBytes: () => a.bytes.slice(), getContentType: () => a.type, setName() { return this; } }) };
  } };
  const createFile = r.s.DriveApp.createFile;
  r.s.DriveApp.createFile = (blob) => { r.converted++; return createFile(blob); };
  return r;
}
const listLogs = (r) => r.logs.filter((l) => l.indexOf('[TDOC LIST]') === 0);
const leaks = (r) => r.logs.filter((l) => l.indexOf(SECRET) !== -1).length;
const buildError = (b) => (b.ok ? null : String(b.error || b.message || ''));

// ====================================================================== A
console.log('A. the configured address works');
{
  const r = report({ [FTP_URL]: XLSX, [WWW_URL]: XLSX });
  const b = r.build();
  check('the build succeeds, with the TDocs of the list', [b.ok, r.order()], [true, TDOCS.map((x) => x.id)]);
  check('one request, to the configured address; the alternative not asked; converted once', [r.listFetches, r.converted], [[FTP_URL], 1]);
  check('nothing logged about the download', listLogs(r), []);
  if (RELEASED) {
    const before = report({ [FTP_URL]: XLSX }, { code: RELEASED });
    const bb = before.build();
    check('the same report as T-2026.10.21 builds from the same list (the TDoc reading unchanged)', [bb.ok, snapshot(r.body) === snapshot(before.body)], [true, true]);
  }
}

// ====================================================================== B
console.log('B. the configured address answers 403, the alternative works');
{
  const r = report({ [FTP_URL]: FORBIDDEN, [WWW_URL]: XLSX });
  const b = r.build();
  check('the build succeeds, with the TDocs of the list', [b.ok, r.order()], [true, TDOCS.map((x) => x.id)]);
  check('the configured address first, then the alternative; converted once', [r.listFetches, r.converted], [[FTP_URL, WWW_URL], 1]);
  check('the log: the failure, then the alternative used (no content)', listLogs(r),
    ['[TDOC LIST] ' + FTP_URL + ': HTTP 403', '[TDOC LIST] downloaded from ' + WWW_URL + ' (6 bytes) after the configured address failed']);
  check('the configured address is unchanged and stays the list\'s identity', [r.docProps.getProperty('TDOC_LIST_URL'), JSON.stringify(r.docProps._store).indexOf('www.3gpp.org/ftp/tsg_sa') === -1], [FTP_URL, true]);
  const a = report({ [FTP_URL]: XLSX });
  a.build();
  check('the report equals the one built from the configured address', snapshot(r.body) === snapshot(a.body), true);
  check('the update downloads the same way', (() => { const n = r.listFetches.length; const u = r.update(); return [u.success, r.listFetches.slice(n)]; })(), [true, [FTP_URL, WWW_URL]]);
  if (RELEASED) {
    const before = report({ [FTP_URL]: FORBIDDEN, [WWW_URL]: XLSX }, { code: RELEASED });
    const bb = before.build();
    check('T-2026.10.21 fails here with the production error, asking only the configured address', [/Could not download TDOC list: HTTP 403/.test(buildError(bb)), before.listFetches], [true, [FTP_URL]]);
  }
}

// ====================================================================== C
console.log('C. the configured address answers HTTP 200 with an HTML page');
{
  const r = report({ [FTP_URL]: HTML_200, [WWW_URL]: XLSX });
  const b = r.build();
  check('refused, the alternative is asked and used', [b.ok, r.listFetches, r.converted], [true, [FTP_URL, WWW_URL], 1]);
  check('the log names it as not a workbook, with size, first bytes and type -- not its text', [listLogs(r)[0], leaks(r)],
    ['[TDOC LIST] ' + FTP_URL + ': HTTP 200 but not an Excel workbook (' + HTML.length + ' bytes starting 3c21444f, text/html)', 0]);
  if (RELEASED) {
    const before = report({ [FTP_URL]: HTML_200 }, { code: RELEASED });
    before.build();
    check('T-2026.10.21 took the HTML page as the workbook and converted it', [before.listFetches, before.converted], [[FTP_URL], 1]);
  }
  const both = report({ [FTP_URL]: HTML_200, [WWW_URL]: HTML_200 });
  const bb = both.build();
  check('both HTML: the build fails, nothing converted', [bb.ok, both.converted], [false, 0]);
}

// ====================================================================== D
console.log('D. both addresses fail');
{
  const r = report({ [FTP_URL]: FORBIDDEN, [WWW_URL]: { status: 404, bytes: HTML, type: 'text/html' } });
  const before = snapshot(r.body);
  const b = r.build();
  check('the build fails with one error naming each address and its reason', [b.ok, buildError(b).indexOf('Could not download TDOC list: HTTP 403 (' + FTP_URL + '); HTTP 404 (' + WWW_URL + ')') !== -1], [false, true]);
  check('nothing converted, no TDoc table written', [r.converted, r.order(), snapshot(r.body) === before || r.order().length === 0], [0, [], true]);
  const thrown = report({ [FTP_URL]: 'throw', [WWW_URL]: 'throw' });
  const tb = thrown.build();
  check('both fetches throw: each named', /Could not download TDOC list: fetch failed \(Address unavailable: https:\/\/ftp\.3gpp\.org\/.*\); fetch failed \(Address unavailable: https:\/\/www\.3gpp\.org\/ftp\//.test(buildError(tb)), true);
  check('no page content in the log', [leaks(r), leaks(thrown)], [0, 0]);
}

// ====================================================================== E
console.log('E. the alternative answers with something that is not a workbook');
{
  const cases = [
    ['an HTML page with HTTP 200', HTML_200, 'HTTP 200 but not an Excel workbook (' + HTML.length + ' bytes starting 3c21444f, text/html)'],
    ['an empty body', { status: 200, bytes: [], type: XLSX_TYPE }, 'HTTP 200 with an empty body'],
    ['other bytes (a legacy .xls)', { status: 200, bytes: [0xD0, 0xCF, 0x11, 0xE0, 0xA1], type: 'application/vnd.ms-excel' }, 'HTTP 200 but not an Excel workbook (5 bytes starting d0cf11e0, application/vnd.ms-excel)'],
    ['three bytes of a signature', { status: 200, bytes: [0x50, 0x4B, 0x03], type: XLSX_TYPE }, 'HTTP 200 but not an Excel workbook (3 bytes starting 504b03, ' + XLSX_TYPE + ')']
  ];
  check('each refused, the build fails naming both reasons, nothing converted', cases.map(([name, answer, reason]) => {
    const r = report({ [FTP_URL]: FORBIDDEN, [WWW_URL]: answer });
    const b = r.build();
    return [name, b.ok, buildError(b).indexOf('Could not download TDOC list: HTTP 403 (' + FTP_URL + '); ' + reason + ' (' + WWW_URL + ')') !== -1, r.converted];
  }), cases.map(([name]) => [name, false, true, 0]));
  const s = report({}).s;
  check('a redirect (3xx) is not a workbook either', (() => {
    s.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 302, getBlob: () => null }) };
    return s.tdocListDownloadAttempt_('https://example.invalid/list.xlsx').problem;
  })(), 'HTTP 302');
}

// ====================================================================== F
console.log('F. addresses not on ftp.3gpp.org: no alternative, the errors of before');
{
  const www = report({ [WWW_URL]: FORBIDDEN }, { url: WWW_URL });
  check('www.3gpp.org/ftp: one request, "HTTP 403" as before', [www.build().ok, www.listFetches, www.s.tdocListDownloadCandidates_(WWW_URL)], [false, [WWW_URL], [WWW_URL]]);
  check('...its error text is the one of before', (() => { try { www.s.downloadTdocListXlsx_(WWW_URL); return null; } catch (e) { return e.message; } })(), 'Could not download TDOC list: HTTP 403');
  const PORTAL = 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916';
  const portal = report({ [PORTAL]: 'throw' }, { url: PORTAL });
  check('the Portal list: a fetch that throws gives its own error, as before', (() => { try { portal.s.downloadTdocListXlsx_(PORTAL); return null; } catch (e) { return e.message; } })(), 'Address unavailable: ' + PORTAL);
  const ok = report({ [PORTAL]: XLSX }, { url: PORTAL });
  check('the Portal list that works: built, one request', [ok.build().ok, ok.listFetches], [true, [PORTAL]]);
}

// ====================================================================== G
console.log('G. the candidates');
{
  const s = report({}).s;
  const c = (u) => s.tdocListDownloadCandidates_(u);
  check('ftp.3gpp.org (https, http, any case, a doubled slash): the same path under https://www.3gpp.org/ftp/', [
    c(FTP_URL), c('http://ftp.3gpp.org/' + LIST_PATH), c('https://FTP.3GPP.ORG/' + LIST_PATH), c('https://ftp.3gpp.org//' + LIST_PATH), c('  ' + FTP_URL + ' ')
  ], [[FTP_URL, WWW_URL], ['http://ftp.3gpp.org/' + LIST_PATH, WWW_URL], ['https://FTP.3GPP.ORG/' + LIST_PATH, WWW_URL], ['https://ftp.3gpp.org//' + LIST_PATH, WWW_URL], [FTP_URL, WWW_URL]]);
  check('other hosts, look-alikes and nothing: no alternative', [
    c(WWW_URL), c('https://ftp.3gpp.org.example.org/' + LIST_PATH), c('https://ftp.3gpp.org@example.org/' + LIST_PATH), c('https://example.org/ftp.3gpp.org/' + LIST_PATH), c(''), c(null)
  ], [[WWW_URL], ['https://ftp.3gpp.org.example.org/' + LIST_PATH], ['https://ftp.3gpp.org@example.org/' + LIST_PATH], ['https://example.org/ftp.3gpp.org/' + LIST_PATH], [], []]);
}

// ====================================================================== H
console.log('H. diagnoseTdocListDownload()');
{
  const r = report({ [FTP_URL]: FORBIDDEN, [WWW_URL]: XLSX });
  const store = JSON.stringify(r.docProps._store);
  const text = r.exec(() => r.s.diagnoseTdocListDownload());
  check('every address, with the build\'s check', text,
    'TDoc list download: configured ' + FTP_URL + ': refused -- HTTP 403; alternative ' + WWW_URL + ': accepted -- HTTP 200, 6 bytes starting 504b0304.');
  check('nothing written: no property, no file, no document change; the log has the same line', [JSON.stringify(r.docProps._store) === store, r.converted, r.order(), listLogs(r)], [true, 0, [], ['[TDOC LIST] diagnosis: ' + text]]);
  const html = report({ [FTP_URL]: HTML_200, [WWW_URL]: HTML_200 });
  check('HTML with HTTP 200 is refused by the diagnosis too; no page text', [/configured .*: refused -- HTTP 200 but not an Excel workbook .*alternative .*: refused -- HTTP 200 but not an Excel workbook/.test(html.exec(() => html.s.diagnoseTdocListDownload())), leaks(html)], [true, 0]);
  const derived = report({}, { props: { TDOC_LIST_URL: null } });
  check('no address saved: the one the build derives from the document folder is diagnosed (and its twin)', /^TDoc list download: configured https:\/\/ftp\.3gpp\.org\/TSG_SA\/WG4_CODEC\/3GPP_SA4_AHOC_MTGs\/SA4_Audio\/Docs\/TDoc_List_Meeting_SA4%23136\.xlsx: accepted .*; alternative https:\/\/www\.3gpp\.org\/ftp\/TSG_SA\//.test(derived.exec(() => derived.s.diagnoseTdocListDownload())), true);
  const none = report({});
  none.s.getReportConfig_ = () => ({ TDOC_LIST_URL: '' });
  const fetched = [];
  none.s.UrlFetchApp = { fetch: (url) => { fetched.push(url); throw new Error('not expected'); } };
  check('no TDoc list at all: nothing fetched', [none.exec(() => none.s.diagnoseTdocListDownload()), fetched], ['No TDoc list configured.', []]);
}

// ====================================================================== I
console.log('I. Close Report downloads the same way');
{
  const r = report({ [FTP_URL]: XLSX, [WWW_URL]: XLSX });
  r.build(); r.update();
  r.s.HtmlService = { createHtmlOutput: (html) => { const o = { setWidth: () => o, setHeight: () => o, getContent: () => html }; return o; } };
  r.ui.showModalDialog = () => {};
  const answers = { [FTP_URL]: FORBIDDEN, [WWW_URL]: XLSX };
  r.s.UrlFetchApp = { fetch: ((inner) => (url, p) => (url === FTP_URL || url === WWW_URL ? (r.listFetches.push(url), { getResponseCode: () => answers[url].status, getBlob: () => ({ getBytes: () => answers[url].bytes.slice(), getContentType: () => answers[url].type, setName() { return this; } }) }) : inner(url, p)))(r.s.UrlFetchApp.fetch) };
  const n = r.listFetches.length;
  const c = r.exec(() => r.s.closeReportCore_());
  check('the external check reads the list from the alternative', [c.externalVerified, r.listFetches.slice(n)], [true, [FTP_URL, WWW_URL]]);
}

// ====================================================================== J
console.log('J. what changed in Code.js: against T-2026.10.21, one section and one function');
{
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutTdocListDownload(FULL);
  check('every change is in Code.js exactly once', view.problems, []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(TDOC_LIST_DOWNLOAD_SECTION_START), FULL.indexOf(TDOC_LIST_DOWNLOAD_SECTION_END) + TDOC_LIST_DOWNLOAD_SECTION_END.length);
  check('the section: these functions, in this order', functionNames(section), TDOC_LIST_DOWNLOAD_SECTION_FUNCTIONS);
  check('it stands directly after the AGENDA READERS section', FULL.indexOf('// END OF AGENDA READERS\n// =========================================================\n\n' + TDOC_LIST_DOWNLOAD_SECTION_START) !== -1, true);
  check('it declares its constants and nothing else at top level', section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || [], ['var TDOC_LIST_FTP_HOST_RE_', 'var TDOC_LIST_WWW_FTP_BASE_']);
  check('the version is unchanged (a candidate)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.5');
  if (RELEASED === null) {
    console.log('  note: template-release/T-2026.10.21 is not available in this checkout; the comparison with it is skipped.');
  } else {
    const released = RELEASED.replace(/\r/g, '');
    check('without it, Code.js is the Code.js of T-2026.10.21, byte for byte', view.code === released, true);
    check('the function that changed: downloadAndGroupTdocs_ (the download)',
      functionNames(released).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(FULL, n) !== functionSource(released, n)).sort(), TDOC_LIST_DOWNLOAD_CHANGED_FUNCTIONS.slice().sort());
    check('the new functions are those of the section; none removed', [functionNames(FULL).filter((n) => functionNames(released).indexOf(n) === -1).sort(), functionNames(released).filter((n) => functionNames(FULL).indexOf(n) === -1)],
      [TDOC_LIST_DOWNLOAD_SECTION_FUNCTIONS.slice().sort(), []]);
    check('untouched: the TDoc reading and grouping helpers, the identity, the agenda readers, the build, the update, Close Report, the other list readers',
      ['assertTdocListIdentity_', 'evaluateTdocListIdentity_', 'readTdocUploadTimesForSessions_', 'buildSkeletonWithTdocTables', 'continuousUpdateCore_', 'closeReportCore_', 'readAgendaWithFallbacks_', 'parseAgendaStructureWithText_',
        'readTdocListWorkbook_', 'isXlsxSignature_', 'downloadAndProcessFromWeb', 'testTdocListUrl', 'getReportConfig_', 'buildRevisionMapFromTdocList_', 'readPortalTdocListPageScope_']
        .map((n) => functionSource(FULL, n) === functionSource(released, n)), Array(15).fill(true));
    const body = (src) => { const f = functionSource(src, 'downloadAndGroupTdocs_'); return f.slice(f.indexOf("blob.setName('TDoc_List_Temp.xlsx');")); };
    check('downloadAndGroupTdocs_: from the blob on (conversion, columns, identity, grouping), unchanged', body(FULL) === body(released), true);
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll TDoc-list download checks passed.');
process.exit(failures ? 1 : 0);
