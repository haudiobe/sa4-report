/**
 * The local revisions folder (after T-2026.10.25).
 *
 * During a physical SA4 meeting, revised documents can appear first on the
 * meeting-room server (http://10.10.10.10/ftp/SA/SA4/Inbox, LAN only). The
 * 3GPP drafts folder is synchronized from it later. The report links to both:
 *
 *   1. LOCAL_REVISIONS_URL: absent -> the default for a main meeting that is
 *      not an e-meeting; 'none' -> off; a URL -> that folder;
 *   2. Configure Meeting has the field and stores, removes or refuses it;
 *      it is a central-state key like REVISIONS_URL;
 *   3. the Revisions cell of a TDoc gets one last line with both folders,
 *      each linked; the revisions themselves are found and linked as before,
 *      in the 3GPP folder only. The local server is never fetched;
 *   4. "No revisions available." gets the line too; text typed by hand is
 *      left alone; an unchanged cell is not written again;
 *   5. the Online information block names both folders, "local" and "remote";
 *   6. without a local folder (ad-hoc, e-meeting, 'none') nothing changes.
 *
 * All values are synthetic. Run: node tests/local-revisions-folder.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { codeWithoutLocalRevisions, LOCAL_REVISIONS_NEW_FUNCTIONS, LOCAL_REVISIONS_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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

const LOCAL = 'http://10.10.10.10/ftp/SA/SA4/Inbox';
const FOLDERS_LINE = 'Revisions folders: Local \u2013 meeting-room server, may be newer, on-site only \u00b7 Remote \u2013 3GPP server, synchronized, accessible externally';
const MAIN = { MEETING_TYPE: 'main', REPORT_SUFFIX: '6G', MEETING_FOLDER: 'TSGS4_136_Montreal' };

// ---- a minimal fake cell/table/body, the surface updateRevisions_() uses (as in tests/revision-drafts-folder.test.js)
function makeCell(text) {
  const state = { text: String(text || ''), links: [], writes: 0 };
  return {
    getText: () => state.text,
    setText: (t) => { state.text = String(t || ''); state.links = []; state.writes++; },
    editAsText: () => ({
      getText: () => state.text,
      setLinkUrl: (start, end, url) => { state.links = state.links.filter((l) => l.end < start || l.start > end); if (url) state.links.push({ start, end, url }); },
      setForegroundColor: () => {}, setBold: () => {}
    }),
    _state: state
  };
}
function makeRow(texts) {
  const cells = texts.map(makeCell);
  return { getNumCells: () => cells.length, getCell: (i) => cells[i], appendTableCell: (t) => { const c = makeCell(t); cells.push(c); return c; } };
}
function makeTable(rows) {
  const r = rows.map(makeRow);
  return { getType: () => 'TABLE', getNumRows: () => r.length, getRow: (i) => r[i], getCell: (a, b) => r[a].getCell(b),
    appendTableRow: () => { const row = makeRow([]); r.push(row); return row; }, removeRow: (i) => { r.splice(i, 1); } };
}
function tdocTable(tdoc, revisions) {
  return makeTable([['TDoc', tdoc], ['Title', 'A title'], ['Source', 'A source'], ['Contact', ''], ['Agenda Item', '5.1'],
    ['E-mail Discussion', ''], ['Revisions', revisions || ''], ['Minutes', 'Typed minutes'], ['Disposition', ''], ['Status', 'Available']]);
}
function body(tables) {
  const children = tables.slice();
  return { getNumChildren: () => children.length, getChild: (i) => children[i], getChildIndex: (c) => children.indexOf(c),
    getTables: () => children.filter((c) => c.getType() === 'TABLE'), insertTable: (i, data) => { const t = makeTable(data); children.splice(i, 0, t); return t; } };
}
const cellOf = (t) => t.getCell(6, 1)._state;
const linked = (cell) => cell.links.map((l) => [cell.text.slice(l.start, l.end + 1), l.url]);

/** One update of the revisions of `tables`, with the 3GPP folder listing `html`. Returns { fetched, sandbox, docProps }. */
function runRevisions(props, tables, html) {
  const loaded = loadCode({ documentProperties: props });
  const doc = body(tables);
  loaded.sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => doc });
  const fetched = [];
  loaded.sandbox.UrlFetchApp = { fetch: (url) => { fetched.push(url); return { getResponseCode: () => 200, getContentText: () => html }; } };
  loaded.sandbox.updateRevisions_(loaded.sandbox.getCollectorConfig_());
  return Object.assign({ fetched: fetched }, loaded);
}

// ====================================================================== 1
console.log('1. which local folder a report has');
{
  const { sandbox: s } = loadCode({});
  const r = (stored, type, ftp) => s.resolveLocalRevisionsUrl_(stored, type, ftp);
  check('a main meeting without a setting: the meeting-room server', r(null, 'main', 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/'), { url: LOCAL, source: 'default' });
  check('a main e-meeting ("...-e" folder): none', r(null, 'main', 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_130-e/Docs/'), { url: '', source: 'none' });
  check('an ad-hoc meeting without a setting: none', r(null, 'adhoc', 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/Ad-hoc_meetings/x/Docs/'), { url: '', source: 'none' });
  check('"none" (any case): off, also for a main meeting', [r('none', 'main', ''), r('NONE', 'main', '')], [{ url: '', source: 'off' }, { url: '', source: 'off' }]);
  check('a configured folder, also for an ad-hoc meeting', r(' http://192.168.1.5/Inbox/ ', 'adhoc', ''), { url: 'http://192.168.1.5/Inbox/', source: 'configured' });
  check('a value that is not an http(s) folder is never linked', [r('ftp://10.10.10.10/Inbox', 'main', ''), r('javascript:alert(1)', 'main', ''), r('http://a b', 'main', '')].map((x) => x.url), ['', '', '']);
  check('a main report with nothing stored at all reads the default', loadCode({ documentProperties: MAIN }).sandbox.localRevisionsUrlFor_(), LOCAL);
  check('...an ad-hoc one none', loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } }).sandbox.localRevisionsUrlFor_(), '');
}

// ====================================================================== 2
console.log('2. Configure Meeting keeps it');
{
  const save = (props, value) => {
    const loaded = loadCode({ documentProperties: props });
    const config = { meetingId: '60777', meetingType: 'main', reportType: '6G', mailingListMode: 'derived', apiTokenAction: 'keep' };
    if (value !== undefined) config.localRevisionsUrl = value;
    loaded.sandbox.persistConfigurationSettings_(config);
    return loaded.docProps.getProperty('LOCAL_REVISIONS_URL');
  };
  check('a folder is stored (trimmed)', save(MAIN, ' http://10.10.10.11/Inbox '), 'http://10.10.10.11/Inbox');
  check('"none" is stored as none', save(MAIN, 'None'), 'none');
  check('blank removes it: the default applies again', save(Object.assign({ LOCAL_REVISIONS_URL: 'none' }, MAIN), ''), null);
  check('a value that is not a folder is not stored; the one before stays', save(Object.assign({ LOCAL_REVISIONS_URL: 'http://10.10.10.12/Inbox' }, MAIN), 'file:///c:/x'), 'http://10.10.10.12/Inbox');
  check('a save without the field (an older dialog) keeps it', save(Object.assign({ LOCAL_REVISIONS_URL: 'http://10.10.10.12/Inbox' }, MAIN), undefined), 'http://10.10.10.12/Inbox');
  const { sandbox: s } = loadCode({});
  check('it reaches the automatic updates: a Configure Meeting key and an adoption key', [s.CONFIG_DIALOG_MANAGED_KEYS_.indexOf('LOCAL_REVISIONS_URL') !== -1, s.ADDON003_ADOPTION_FIXED_KEYS_.indexOf('LOCAL_REVISIONS_URL') !== -1], [true, true]);

  const loaded = loadCode({ documentProperties: Object.assign({ LOCAL_REVISIONS_URL: 'http://10.10.10.13/Inbox' }, MAIN) });
  let html = null;
  loaded.sandbox.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  loaded.sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  loaded.sandbox.configureMeetingSettings();
  check('the dialog shows the stored value, with the default as placeholder', [/<input type="text" id="localRevisionsUrl" value="http:\/\/10\.10\.10\.13\/Inbox" placeholder="http:\/\/10\.10\.10\.10\/ftp\/SA\/SA4\/Inbox">/.test(html)], [true]);
  check('...and sends it with a save', /localRevisionsUrl: document\.getElementById\('localRevisionsUrl'\)/.test(html), true);
}

// ====================================================================== 3
console.log('3. the Revisions cell of a main report');
{
  const A = tdocTable('S4-260123'), B = tdocTable('S4-260124'), C = tdocTable('S4-260125', 'Typed by hand: see the chair\'s notes');
  const listing = '<html><body><a href="S4-260123_r1.docx">S4-260123_r1.docx</a> <a href="S4-260123_r2.docx">S4-260123_r2.docx</a></body></html>';
  const run = runRevisions(MAIN, [A, B, C], listing);
  const remote = run.fetched[0];
  check('only the 3GPP folder is fetched, never the local server', [run.fetched.length, run.fetched.some((u) => u.indexOf('10.10.10.10') !== -1)], [1, false]);
  check('the revisions as before, then the line of both folders', cellOf(A).text, 'S4-260123_r1.docx\nS4-260123_r2.docx\n' + FOLDERS_LINE);
  check('each revision is linked directly in the 3GPP folder; each folder by its label',
    linked(cellOf(A)), [['S4-260123_r1.docx', remote + 'S4-260123_r1.docx'], ['S4-260123_r2.docx', remote + 'S4-260123_r2.docx'],
      ['Local \u2013 meeting-room server, may be newer, on-site only', LOCAL], ['Remote \u2013 3GPP server, synchronized, accessible externally', remote]]);
  check('a TDoc without revisions: "No revisions available." and the folders (the local one may have one already)', [cellOf(B).text, linked(cellOf(B)).map((l) => l[1])], ['No revisions available.\n' + FOLDERS_LINE, [LOCAL, remote]]);
  check('text typed by hand is left alone', [cellOf(C).text, cellOf(C).writes], ['Typed by hand: see the chair\'s notes', 0]);
  check('the Minutes are not touched', [A, B, C].map((t) => t.getCell(7, 1).getText()), ['Typed minutes', 'Typed minutes', 'Typed minutes']);
  check('no revision is taken from the line of the folders (no TDoc number in it)', run.sandbox.findSA4DocumentIdsInText_(FOLDERS_LINE), []);

  // the next update: the same cell is not written again when nothing changed
  const writesB = cellOf(B).writes;
  const again = runRevisions(MAIN, [B], listing);
  check('an unchanged "No revisions available." cell is not written again', [cellOf(B).writes, again.fetched.length], [writesB, 1]);
}

// ====================================================================== 4
console.log('4. without a local folder nothing changes');
{
  const listing = '<html><body><a href="S4-260123_r1.docx">S4-260123_r1.docx</a></body></html>';
  [['an ad-hoc report', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: 'https://www.3gpp.org/ftp/x/Inbox/Drafts/Y/' }],
    ['a main report set to "none"', Object.assign({ LOCAL_REVISIONS_URL: 'none' }, MAIN)],
    ['a main e-meeting', { MEETING_TYPE: 'main', REPORT_SUFFIX: '6G', MEETING_FOLDER: 'TSGS4_130-e' }]].forEach(([label, props]) => {
    const A = tdocTable('S4-260123'), B = tdocTable('S4-260124'), D = tdocTable('S4-260125', 'No revisions available.');
    runRevisions(props, [A, B, D], listing);
    check(label + ': the cells as before', [cellOf(A).text, cellOf(B).text, cellOf(D).text, cellOf(D).writes], ['S4-260123_r1.docx', 'No revisions available.', 'No revisions available.', 0]);
  });
  const E = tdocTable('S4-260124', 'No revisions available.\n' + FOLDERS_LINE);
  runRevisions(Object.assign({ LOCAL_REVISIONS_URL: 'none' }, MAIN), [E], '<html></html>');
  check('switched off later: the line is removed again', [cellOf(E).text, cellOf(E).links], ['No revisions available.', []]);
}

// ====================================================================== 5
console.log('5. the Online information block');
{
  const { sandbox: s } = loadCode({});
  const SESSIONS = [{ id: 's1', code: 'S1', date: '2026-11-16', endDate: '', startTime: '', endTime: '', label: '' }];
  const src = { meetingName: 'SA4#136', meetingId: '', tdocListUrl: '', docsFolder: 'https://www.3gpp.org/ftp/x/Docs/', draftsFolder: 'https://www.3gpp.org/ftp/x/Inbox/Drafts/Y' };
  const lines = (source) => s.buildAdhocMeetingInformationLines_(SESSIONS, source).filter((l) => /^Drafts Folder/.test(l.text));
  check('without a local folder: the one line as before', lines(src), [{ text: 'Drafts Folder: https://www.3gpp.org/ftp/x/Inbox/Drafts/Y', link: 'https://www.3gpp.org/ftp/x/Inbox/Drafts/Y' }]);
  check('with one: local first, then remote, both linked and labelled', lines(Object.assign({ localDraftsFolder: LOCAL }, src)), [
    { text: 'Drafts Folder (local \u2013 meeting-room server, may be newer, on-site only): ' + LOCAL, link: LOCAL },
    { text: 'Drafts Folder (remote \u2013 3GPP server, synchronized, accessible externally): https://www.3gpp.org/ftp/x/Inbox/Drafts/Y', link: 'https://www.3gpp.org/ftp/x/Inbox/Drafts/Y' }]);
  // The block is found by its heading and the lines below it that begin with one of its labels.
  const para = (text, heading) => ({ getType: () => 'PARAGRAPH', asParagraph: () => ({ getText: () => text, getHeading: () => heading || 'NORMAL' }) });
  const blockOf = (texts) => { const children = [para('Online information', 'H3')].concat(texts.map((t) => para(t)), [para('Typed below the block')]);
    return s.findAdhocMeetingInformationContainer_({ getNumChildren: () => children.length, getChild: (i) => children[i] }); };
  const both = s.buildAdhocMeetingInformationLines_(SESSIONS, Object.assign({ localDraftsFolder: LOCAL }, src)).map((l) => l.text);
  const old = s.buildAdhocMeetingInformationLines_(SESSIONS, src).map((l) => l.text);
  check('the block with both folders is recognized up to its last line, and text typed below is not part of it', blockOf(both), { start: 0, end: both.length + 1 });
  check('an older block (one Drafts Folder line) is still recognized, so it is replaced and not doubled', blockOf(old), { start: 0, end: old.length + 1 });
  const main = loadCode({ documentProperties: Object.assign({ MEETING_NAME: 'SA4#136' }, MAIN) }).sandbox;
  const mainSource = main.adhocMeetingInformationSource_();
  check('a main report: the local folder, and the 3GPP folder its updates read', [mainSource.localDraftsFolder, mainSource.draftsFolder], [LOCAL, main.getReportConfig_().REVISIONS_URL]);
  const adhoc = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REVISIONS_URL: 'https://www.3gpp.org/ftp/x/Inbox/Drafts/Y' } }).sandbox.adhocMeetingInformationSource_();
  check('an ad-hoc report: as before', [adhoc.localDraftsFolder, adhoc.draftsFolder], ['', 'https://www.3gpp.org/ftp/x/Inbox/Drafts/Y']);
}

// ====================================================================== 6
console.log('6. what the change is in Code.js, and nothing else (source)');
{
  const now = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutLocalRevisions(now);
  check('the take-out finds every change exactly once', view.problems, []);
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (code, name) => { const at = code.indexOf('\nfunction ' + name + '('); return code.slice(at, code.indexOf('\n}\n', at)); };
  const before = names(view.code);
  check('the new functions', names(now).filter((n) => before.indexOf(n) === -1).sort(), LOCAL_REVISIONS_NEW_FUNCTIONS.slice().sort());
  check('the changed functions', before.filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)).sort(), LOCAL_REVISIONS_CHANGED_FUNCTIONS.slice().sort());
  check('untouched: revision detection and placement, TDoc selection, the update, statuses',
    ['parseAnchors_', 'buildRevisionAnchorIndex_', 'parseDraftAnchor_', 'normalizeRevisionKey_', 'insertRevisedDocTablesAfter_', 'collectRevisionsCore_', 'downloadAndGroupTdocs_',
      'continuousUpdateCore_', 'insertNewTdoc_', 'updateStatusWithStrictRules_', 'applyTdocStatusUpdate_', 'getMeetingContext_'].filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)), []);
  check('the local server is never fetched: no UrlFetchApp in the new functions', LOCAL_REVISIONS_NEW_FUNCTIONS.filter((n) => /UrlFetchApp|safeFetch_/.test(bodyOf(now, n))), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll local revisions folder checks passed.');
process.exit(failures ? 1 : 0);
