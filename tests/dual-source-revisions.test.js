/**
 * Issue #16: revisions from two sources -- the 3GPP drafts folder (fetched by
 * every update, as before) and the meeting-room server (LAN only, never
 * reachable from Apps Script: its listing is pasted in by someone in the
 * room, "Import Local Inbox Listing…").
 *
 * Fixture-based, on the REAL Code.js (loadCode + a fake document):
 *   1. parsing a pasted listing: HTML links, absolute/relative/rooted hrefs,
 *      text only, shortened names, foreign links, a folder of another server;
 *   2. merging: remote only (unchanged), local only, both, conflicting file
 *      names, duplicates;
 *   3. the import: only TDocs that have a table, only adds, writes only the
 *      cells it changed, nothing else in the report;
 *   4. the update afterwards: stale remote, remote catching up, the 3GPP
 *      folder unreachable (nothing lost), repeat runs (no change), no
 *      revision table from a local-only file;
 *   5. offline / malformed / denied: nothing changes;
 *   6. status texts: not checked / no new revisions / new ones;
 *   7. source: only these functions changed; detection untouched.
 *
 * All values are synthetic. Run: node tests/dual-source-revisions.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { codeWithoutDualRevisions, DUAL_REVISIONS_NEW_FUNCTIONS, DUAL_REVISIONS_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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
const FOLDER = LOCAL + '/Drafts/MBS/';
const MAIN = { MEETING_TYPE: 'main', REPORT_SUFFIX: '6G', MEETING_FOLDER: 'TSGS4_138_Calgary' };
const FOLDERS_PREFIX = 'Revisions folders: ';

// ---- fake document (the surface the revisions code uses)
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
    appendTableRow: () => { const row = makeRow([]); r.push(row); return row; }, removeRow: (i) => { r.splice(i, 1); }, setBorderWidth: () => {} };
}
function tdocTable(tdoc, revisions) {
  return makeTable([['TDoc', tdoc], ['Title', 'A title'], ['Source', 'A source'], ['Contact', ''], ['Agenda Item', '5.1'],
    ['E-mail Discussion', ''], ['Revisions', revisions || ''], ['Minutes', 'Typed minutes of ' + tdoc], ['Disposition', 'agreed'], ['Status', 'Agreed']]);
}
function makeBody(tables) {
  const children = tables.slice();
  return { _children: children, getNumChildren: () => children.length, getChild: (i) => children[i], getChildIndex: (c) => children.indexOf(c),
    getTables: () => children.filter((c) => c.getType() === 'TABLE'),
    insertTable: (i) => { const t = makeTable([]); children.splice(i, 0, t); return t; } };
}
const cellOf = (t) => t.getCell(6, 1)._state;
const linked = (cell) => cell.links.map((l) => [cell.text.slice(l.start, l.end + 1), l.url]);
const linesOf = (t) => cellOf(t).text.split('\n').filter((l) => l.indexOf(FOLDERS_PREFIX) !== 0);
const snapshot = (tables) => tables.map((t) => [0, 1, 2, 3, 4, 5, 7, 8, 9].map((r) => t.getCell(r, 1).getText()));

/** A report: the REAL Code.js on a fake document with these tables; `remoteHtml` is the 3GPP folder (null: unreachable). */
function report(props, tables, remoteHtml) {
  const loaded = loadCode({ documentProperties: Object.assign({}, MAIN, props || {}) });
  const s = loaded.sandbox;
  const body = makeBody(tables);
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'synthetic-doc' });
  s.LockService = { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }), getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  const fetched = [];
  const r = { s, body, docProps: loaded.docProps, fetched, remoteHtml };
  s.UrlFetchApp = { fetch: (url) => {
    fetched.push(url);
    if (/10\.10\.10\.10/.test(url)) throw new Error('the local server must never be fetched: ' + url);
    if (r.remoteHtml === null) return { getResponseCode: () => 503, getContentText: () => '' };
    return { getResponseCode: () => 200, getContentText: () => r.remoteHtml };
  } };
  r.update = () => s.updateRevisions_(s.getCollectorConfig_());
  r.paste = (payload) => s.saveLocalRevisionsListing(Object.assign({ folder: FOLDER, html: '', text: '' }, payload));
  r.remote = () => s.getCollectorConfig_().REVISIONS_URL.replace(/\/$/, '') + '/';
  return r;
}
const listing = (files, base) => '<html><body><h1>Index of /ftp/SA/SA4/Inbox/Drafts/MBS</h1><a href="?C=N;O=D">Name</a> <a href="../">Parent Directory</a>\n' +
  files.map((f) => '<a href="' + (base === undefined ? '' : base) + encodeURIComponent(f) + '">' + f + '</a>').join('\n') + '</body></html>';

const A = 'S4-260123', B = 'S4-260124', C = 'S4-260125';

// ====================================================================== 1
console.log('1. reading a pasted listing');
{
  const { sandbox: s } = loadCode({});
  const p = (payload) => s.parseLocalRevisionsListing_(Object.assign({ folder: FOLDER }, payload), LOCAL);
  const rel = p({ html: listing([A + '_r1.docx', B + '_r2 Nokia.zip']) });
  check('HTML links, relative: resolved against the listed folder, exactly as listed', rel.anchors.map((a) => a.href).filter((h) => /S4-/.test(h)),
    [FOLDER + A + '_r1.docx', FOLDER + encodeURIComponent(B + '_r2 Nokia.zip')]);
  check('rooted and absolute hrefs stay on the meeting-room server', p({ html: '<a href="/ftp/SA/SA4/Inbox/' + A + '_r1.docx">' + A + '_r1.docx</a><a href="http://10.10.10.10/x/' + B + '.zip">' + B + '.zip</a>' }).anchors.map((a) => a.href),
    ['http://10.10.10.10/ftp/SA/SA4/Inbox/' + A + '_r1.docx', 'http://10.10.10.10/x/' + B + '.zip']);
  check('single-quoted hrefs (IIS-style listings) are read too', p({ html: "<A HREF='" + A + "_r3.docx'>" + A + '_r3.docx</A>' }).anchors.map((a) => a.href), [FOLDER + A + '_r3.docx']);
  check('a link to another server, a protocol-relative one or javascript: is dropped (never a foreign link in the report)',
    p({ html: '<a href="https://evil.example/' + A + '_r1.docx">' + A + '_r1.docx</a><a href="//evil.example/' + A + '.docx">' + A + '.docx</a><a href="javascript:alert(1)">' + A + '_r9.docx</a>' }),
    { ok: false, error: 'The pasted text has no file list. Nothing was changed.', folder: FOLDER, anchors: [], listed: 0, truncated: 0 });
  const text = p({ text: 'Name    Last modified    Size\nParent Directory\n' + A + '_r1.docx    2026-11-16 10:01   20K\n' + B + '_r2_with_a_very_long_name_that_is_cut..>  2026-11-16 10:02  1M\n' });
  check('text only (a browser that copies no HTML): complete file names, in the listed folder', text.anchors.map((a) => a.href), [FOLDER + A + '_r1.docx']);
  check('...a name the listing cut short ("..>") is not used, and counted', text.truncated, 1);
  check('a folder on another server is refused', p({ folder: 'https://www.3gpp.org/ftp/x/', html: listing([A + '_r1.docx']) }).error, 'The listed folder is not on the meeting-room server of this report (http://10.10.10.10).');
  check('no local folder configured: refused', s.parseLocalRevisionsListing_({ folder: FOLDER, html: listing([A + '.docx']) }, '').error, 'This report has no Local Revisions URL (Configure Meeting).');
  check('an error page pasted instead of a listing: nothing recognized', p({ html: '<html><body>This site can\u2019t be reached. 10.10.10.10 took too long to respond.</body></html>', text: 'This site can\u2019t be reached' }).ok, false);
}

// ====================================================================== 2
console.log('2. merging the two sources');
{
  const { sandbox: s } = loadCode({});
  const R = 'https://www.3gpp.org/ftp/x/Inbox/Drafts/MBS/';
  const remote = [{ text: A + '_r1.docx', link: R + A + '_r1.docx' }, { text: A + '_r2.docx', link: R + A + '_r2.docx' }];
  check('remote only: the very same list (nothing changes without local ones)', s.mergeLocalRevisions_(remote, {}) === remote, true);
  const merged = s.mergeLocalRevisions_(remote, {
    a: { text: A + '_R1.DOCX', link: FOLDER + A + '_R1.DOCX' },          // the same file, other case
    b: { text: A + '_r3.docx', link: FOLDER + A + '_r3.docx' },          // local only (remote not synchronized yet)
    c: { text: A + '_r2%20.docx', link: FOLDER + A + '_r2%20.docx' },    // a different name: a different file
    d: { text: A + '_r3.docx', link: FOLDER + 'other/' + A + '_r3.docx' } // a duplicate of b: once
  });
  check('both: one line with both links; local only: its local link only; conflicting names stay apart; duplicates once', merged, [
    { text: A + '_r1.docx', link: R + A + '_r1.docx', localLink: FOLDER + A + '_R1.DOCX' },
    { text: A + '_r2.docx', link: R + A + '_r2.docx' },
    { text: A + '_r2%20.docx', link: '', localLink: FOLDER + A + '_r2%20.docx' },
    { text: A + '_r3.docx', link: '', localLink: FOLDER + A + '_r3.docx' }]);
  check('the remote list itself is not changed by a merge', remote, [{ text: A + '_r1.docx', link: R + A + '_r1.docx' }, { text: A + '_r2.docx', link: R + A + '_r2.docx' }]);
  check('a stored local entry without an http link is ignored', s.mergeLocalRevisions_(remote, { x: { text: A + '_r5.docx', link: 'file:///c:/x' } }) === remote, true);
}

// ====================================================================== 3
console.log('3. the import');
{
  const tA = tdocTable(A), tB = tdocTable(B), tC = tdocTable(C);
  const r = report({}, [tA, tB, tC], listing([A + '_r1.docx']));
  r.update();                                       // the 3GPP folder has A_r1 (as before)
  const remote = r.remote();
  const before = snapshot([tA, tB, tC]);
  const writesC = cellOf(tC).writes;
  const res = r.paste({ html: listing([A + '_r1.docx', A + '_r2.docx', B + '_r1.docx', 'S4-269999_r1.docx', 'README.txt']) });
  // 6 links: the parent folder and five files (the sort link "?C=N;O=D" is not one); 4 are revisions of SA4 TDocs.
  check('the result says what was found and what is new', [res.ok, res.message.split('\n')[0].replace(/imported .*? from /, 'imported <time> from ')],
    [true, 'Local (meeting-room server): last listing imported <time> from ' + FOLDER + ': 6 file(s), 4 revision(s) of SA4 TDocs, 3 new for this report.']);
  check('A: the 3GPP revision now also local; the new local one first shown here', linesOf(tA), [A + '_r1.docx [local]', A + '_r2.docx [local only]']);
  check('...each name linked: 3GPP where it is there, else local; the markers to the local files', linked(cellOf(tA)).slice(0, 4), [
    [A + '_r1.docx', remote + A + '_r1.docx'], ['[local]', FOLDER + A + '_r1.docx'], [A + '_r2.docx', FOLDER + A + '_r2.docx'], ['[local only]', FOLDER + A + '_r2.docx']]);
  check('B: a local-only revision where the 3GPP folder had none', linesOf(tB), [B + '_r1.docx [local only]']);
  check('C (nothing local) is not written; a TDoc without a table is not imported', [cellOf(tC).writes === writesC, r.docProps.getProperty('LOCALREV_S4-269999')], [true, null]);
  check('nothing else in the report changed (TDoc, title, minutes, disposition, status...)', snapshot([tA, tB, tC]), before);
  check('no table was added', r.body.getTables().length, 3);
  check('nothing was fetched by the import, and the local server never', r.fetched.length, 1);

  const writes = [tA, tB].map((t) => cellOf(t).writes);
  const again = r.paste({ html: listing([A + '_r1.docx', A + '_r2.docx', B + '_r1.docx']) });
  check('the same listing again: nothing new, no cell written', [/no new revisions for this report\./.test(again.message), [tA, tB].map((t) => cellOf(t).writes)], [true, writes]);
}

// ====================================================================== 4
console.log('4. the updates afterwards');
{
  const tA = tdocTable(A), tB = tdocTable(B);
  const r = report({}, [tA, tB], listing([A + '_r1.docx']));
  r.update();
  r.paste({ html: listing([A + '_r2.docx', B + '_r1.docx']) });
  r.update();                                       // the 3GPP folder is stale: still only A_r1
  check('stale remote: the update keeps the local-only revisions', [linesOf(tA), linesOf(tB)], [[A + '_r1.docx', A + '_r2.docx [local only]'], [B + '_r1.docx [local only]']]);
  check('...and adds no revision table for a local-only file (no new imports)', r.body.getTables().length, 2);
  const textA = cellOf(tA).text;
  r.update();
  check('repeat run: the same cell', cellOf(tA).text, textA);
  r.remoteHtml = listing([A + '_r1.docx', A + '_r2.docx']);
  r.update();                                       // the 3GPP folder caught up
  check('remote caught up: the revision is now linked to 3GPP, and still marked local', [linesOf(tA), linked(cellOf(tA)).filter((l) => l[0] === A + '_r2.docx')[0][1]], [[A + '_r1.docx', A + '_r2.docx [local]'], r.remote() + A + '_r2.docx']);
  const kept = cellOf(tA).text;
  r.remoteHtml = null;                              // the 3GPP folder unreachable
  r.update();
  check('3GPP folder unreachable: the cells stay as they are', cellOf(tA).text, kept);
  check('the tables of the 3GPP revisions are placed as before (A_r1/A_r2 are revisions of A itself: none)', r.body.getTables().length, 2);
}

// ====================================================================== 5
console.log('5. offline, malformed, denied: nothing changes');
{
  const tA = tdocTable(A);
  const r = report({}, [tA], listing([A + '_r1.docx']));
  r.update();
  const text = cellOf(tA).text;
  const results = [
    r.paste({ html: '', text: '' }),
    r.paste({ html: '<html><body>ERR_ADDRESS_UNREACHABLE</body></html>', text: 'This site can\u2019t be reached' }),
    r.paste({ folder: 'https://www.3gpp.org/ftp/', html: listing([A + '_r9.docx']) }),
    r.paste({ folder: 'not a url', html: listing([A + '_r9.docx']) })
  ];
  check('each is refused with a reason', results.map((x) => x.ok), [false, false, false, false]);
  check('...the cell, the stored revisions and the status are untouched', [cellOf(tA).text, r.docProps.getProperty('LOCALREV_' + A), r.docProps.getProperty('LOCAL_REVISIONS_LAST_IMPORT')], [text, null, null]);
  const off = report({ LOCAL_REVISIONS_URL: 'none' }, [tdocTable(A)], listing([]));
  check('a report without a local folder refuses an import', off.paste({ html: listing([A + '_r1.docx']) }).message, 'This report has no Local Revisions URL (Configure Meeting).');
  const busy = report({}, [tdocTable(A)], listing([]));
  busy.s.LockService = { getDocumentLock: () => ({ tryLock: () => false, releaseLock: () => {} }) };
  check('the report is being updated: nothing changed, try again', busy.paste({ html: listing([A + '_r1.docx']) }).message, 'The report is being updated. Try again in a minute; nothing was changed.');
  const broken = report({ ['LOCALREV_' + A]: '{not json' }, [tdocTable(A)], listing([A + '_r1.docx']));
  broken.update();
  check('a damaged stored entry: the update renders the 3GPP revisions as before', linesOf(broken.body.getTables()[0]), [A + '_r1.docx']);
}

// ====================================================================== 6
console.log('6. what the report says about the local source');
{
  const { sandbox: s } = loadCode({});
  check('not checked yet', s.describeLocalRevisionsStatus_(null, LOCAL), 'Local (meeting-room server): not checked yet. Import Local Inbox Listing… on the meeting-room network.');
  check('no local folder', s.describeLocalRevisionsStatus_(null, ''), 'Local (meeting-room server): none configured.');
  check('checked, no new revisions', s.describeLocalRevisionsStatus_(JSON.stringify({ at: '2026-11-16 10:05', folder: FOLDER, listed: 12, recognized: 3, added: 0 }), LOCAL),
    'Local (meeting-room server): last listing imported 2026-11-16 10:05 from ' + FOLDER + ': 12 file(s), 3 revision(s) of SA4 TDocs, no new revisions for this report.');
}

// ====================================================================== 7
console.log('7. what the change is in Code.js (source)');
{
  const now = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutDualRevisions(now);
  check('the take-out finds every change exactly once', view.problems, []);
  const names = (code) => (code.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (code, name) => { const at = code.indexOf('\nfunction ' + name + '('); return code.slice(at, code.indexOf('\n}\n', at)); };
  const before = names(view.code);
  check('the new functions', names(now).filter((n) => before.indexOf(n) === -1).sort(), DUAL_REVISIONS_NEW_FUNCTIONS.slice().sort());
  check('the changed functions', before.filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)).sort(), DUAL_REVISIONS_CHANGED_FUNCTIONS.slice().sort());
  check('untouched: 3GPP detection and parsing, revision tables, TDoc selection, the update, statuses, e-mails, abstracts',
    ['parseAnchors_', 'buildRevisionAnchorIndex_', 'parseDraftAnchor_', 'normalizeRevisionKey_', 'toAbsoluteUrl_', 'insertRevisedDocTablesAfter_', 'collectRevisionsCore_', 'downloadAndGroupTdocs_',
      'continuousUpdateCore_', 'insertNewTdoc_', 'updateStatusWithStrictRules_', 'applyTdocStatusUpdate_', 'getMeetingContext_', 'fetchAndAddAbstract_', 'collectorUpdate_'].filter((n) => bodyOf(now, n) !== bodyOf(view.code, n)), []);
  check('no new function fetches anything', DUAL_REVISIONS_NEW_FUNCTIONS.filter((n) => /UrlFetchApp|safeFetch_/.test(bodyOf(now, n))), []);
  const upd = bodyOf(now, 'updateRevisions_');
  check('updateRevisions_(): the 3GPP revisions (not the local ones) still go to the revision tables', /insertRevisedDocTablesAfter_\(body, t, ordered, cfg\);/.test(upd), true);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll dual-source revision checks passed.');
process.exit(failures ? 1 : 0);
