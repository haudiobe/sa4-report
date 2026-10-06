/**
 * TDoc upload completion -- an EXISTING TDoc that entered the report before
 * it was uploaded gets what the upload brings, once the TDoc list shows the
 * upload: the hyperlink of its TDoc number (in its table and in its row of
 * the registration table) and one attempt at its abstract.
 *
 * The defect, as found in the live 6G report (meeting 86178) on 2026-10-05:
 *
 *   S4aP260099  still reserved; no link in the list
 *   S4aP260100  reserved 29/09, uploaded 30/09   table: no link, Status "available"
 *   S4aP260101  inserted after its upload        table: link,    Status "available"
 *   S4aP260102  reserved 10:09, uploaded 10:32   table: no link, Status "available"
 *
 * The update wrote the link only when it inserted a table, so 100 and 102
 * never got one. The abstract comes from the Contribution Reviewer, which
 * answers 404 for a TDoc that is not uploaded; that answer is cached for 24
 * hours and blocked the abstract after the upload, and with
 * FETCH_ABSTRACTS_ON_UPDATE off the update never asked at all. The Status
 * was correct in the report and its logic is not part of this.
 *
 * Everything here runs the real continuousUpdateCore_() -- the function
 * "Update Report Now" and the automatic update run -- on a report whose
 * tables record every write, with the Reviewer replaced by a fake.
 *
 * Run: node tests/tdoc-upload-completion.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
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

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
function functionSource(code, name) {
  const start = code.indexOf('\nfunction ' + name + '(');
  if (start === -1) return null;
  return code.slice(start + 1, code.indexOf('\n}\n', start) + 2);
}
const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
function gitShow(spec) {
  try {
    return execFileSync('git', ['show', spec], { cwd: path.join(__dirname, '..'), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    return null;
  }
}
const withoutComments = (src) => src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');

const DOCS = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/';
const zip = (id) => DOCS + id + '.zip';
const OTHER_LINK = 'https://example.org/somewhere/else.docx';

/**
 * A table of the report. Every cell keeps one link per character (as Google
 * Docs does), and every write to the table is recorded in `_writes`, so a
 * test can show that nothing but what it names was touched.
 */
function makeTable(sandbox, rowsData) {
  const writes = [];
  const rows = [];
  const makeCell = (where, value) => {
    const c = { _t: String(value), _links: String(value).split('').map(() => null) };
    const te = {
      getText: () => c._t,
      getLinkUrl: (i) => (i === undefined ? null : c._links[i] || null),
      setLinkUrl: (start, end, url) => { writes.push(['setLinkUrl', where(), start, end, url]); for (let i = start; i <= end; i++) c._links[i] = url; return te; },
      setBold: () => { writes.push(['setBold', where()]); return te; },
      setForegroundColor: () => { writes.push(['setForegroundColor', where()]); return te; },
      setFontSize: (n) => { writes.push(['setFontSize', where(), n]); return te; }
    };
    c.getText = () => c._t;
    c.setText = (v) => { writes.push(['setText', where(), String(v)]); c._t = String(v); c._links = c._t.split('').map(() => null); return c; };
    c.editAsText = () => te;
    return c;
  };
  const makeRow = (texts) => {
    const cells = [];
    const row = {
      getNumCells: () => cells.length, getCell: (i) => cells[i],
      appendTableCell: (t) => { const i = cells.length; const c = makeCell(() => (cells[0] ? cells[0]._t : '?') + '[' + i + ']', t === undefined ? '' : t); cells.push(c); return c; }
    };
    (texts || []).forEach((t) => row.appendTableCell(t));
    return row;
  };
  rowsData.forEach((r) => rows.push(makeRow(r)));
  const table = {
    _writes: writes,
    getType: () => sandbox.DocumentApp.ElementType.TABLE, asTable: () => table,
    getNumRows: () => rows.length, getRow: (i) => rows[i], getCell: (r, c) => rows[r].getCell(c),
    insertTableRow: (i) => { writes.push(['insertTableRow', i]); const r = makeRow([]); rows.splice(i, 0, r); return r; },
    appendTableRow: () => { writes.push(['appendTableRow']); const r = makeRow([]); rows.push(r); return r; },
    removeRow: (i) => { writes.push(['removeRow', i]); rows.splice(i, 1); }
  };
  return table;
}
const setLink = (cell, url) => { cell._links = cell._t.split('').map(() => url); };

/** A TDoc table as the update inserts it; `o`: { link, abstract, minutes, disposition, tdocCellText }. */
function makeTdocTable(sandbox, id, status, o) {
  o = o || {};
  const rows = [['TDoc', o.tdocCellText === undefined ? id : o.tdocCellText], ['Title', 'Title of ' + id], ['Source', 'Source Co'], ['Contact', 'A Person'],
    ['Agenda Item', '5.4'], ['Type/For', 'pCR for Agreement'], ['E-mail Discussion', ''], ['Revisions', ''],
    ['Minutes', o.minutes === undefined ? 'Presented by A Person.\nQuestion from B: why?\tAnswer: because.' : o.minutes],
    ['Disposition', o.disposition || ''], ['Status', status]];
  if (o.abstract !== undefined) rows.splice(5, 0, ['Abstract', o.abstract]);
  const table = makeTable(sandbox, rows);
  if (o.link) setLink(table.getCell(0, 1), o.link);
  return table;
}
/** The registration table: `entries` are [id, link or null]. */
function makeRegistrationTable(sandbox, entries, withSession) {
  const header = ['TDoc', 'Title', 'Source', 'Agenda Item'].concat(withSession ? ['Session'] : []);
  const table = makeTable(sandbox, [header].concat(entries.map(([id]) => [id, 'Title of ' + id, 'Source Co', '5.4'].concat(withSession ? ['A01'] : []))));
  entries.forEach(([, link], i) => { if (link) setLink(table.getCell(i + 1, 0), link); });
  return table;
}

const linksOf = (cell) => cell._links;
const linkOf = (table) => { const all = linksOf(table.getCell(0, 1)).filter(Boolean); return all.length ? all[0] : null; };
const field = (table, label) => { for (let r = 0; r < table.getNumRows(); r++) if (table.getCell(r, 0).getText() === label) return table.getCell(r, 1).getText(); return undefined; };
const labels = (table) => { const out = []; for (let r = 0; r < table.getNumRows(); r++) out.push(table.getCell(r, 0).getText()); return out; };
const textOf = (table) => { const out = []; for (let r = 0; r < table.getNumRows(); r++) { const row = []; for (let c = 0; c < table.getRow(r).getNumCells(); c++) row.push(table.getCell(r, c).getText()); out.push(row); } return out; };
/** The fields of a TDoc table the update must never write, as they are. */
const PROTECTED = ['TDoc', 'Title', 'Source', 'Contact', 'Agenda Item', 'Type/For', 'Minutes', 'Disposition', 'Status'];
const protectedOf = (table) => PROTECTED.map((label) => [label, field(table, label)]);
const linkWrite = (id, url) => ['setLinkUrl', 'TDoc[1]', 0, id.length - 1, url];
const abstractWrites = (text) => [['insertTableRow', 5], ['setFontSize', 'Abstract[1]', 8]];

/** One TDoc of the TDoc list, as downloadAndGroupTdocs_() hands it on: `link` is the hyperlink of its TDoc cell, null when it has none. */
function listRow(id, status, link) {
  const noLink = { getLinkUrl: () => null };
  return {
    row: [id, 'CHANGED title of ' + id, 'CHANGED source', '9.9', status, 'CHANGED contact', 'CR', 'Approval'],
    richTextRow: [{ getLinkUrl: () => link || null }, noLink, noLink, noLink, noLink, noLink, noLink, noLink],
    tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: 5, agendaCol: 3, agendaTopicCol: -1, statusCol: 4, typeCol: 6, forCol: 7, revisedToCol: -1
  };
}

const ADHOC_6G = { MEETING_TYPE: 'adhoc', MEETING_ID: '86178', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'September 22, 2026', REPORT_SUFFIX: '6G',
  FTP_BASE: DOCS, TDOC_LIST_URL: 'https://www.3gpp.org/ftp/x/list.xlsx', AGENDA_TDOC: 'S4aP260098' };
const ADHOC_MBS = { MEETING_TYPE: 'adhoc', MEETING_ID: '86172', MEETING_NAME: 'Synthetic MBS ad-hoc', MEETING_DATE: 'October 1, 2026', REPORT_SUFFIX: 'MBS',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/x/list.xlsx', AGENDA_TDOC: 'S4aI260001' };
const MAIN = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137_Synthetic', MEETING_NUMBER: '137', MEETING_ID: '60778', REPORT_SUFFIX: 'Audio' };

const CACHE = (id) => 'REVIEWER_NO_SUMMARY_CACHE_' + id;
const HOUR = 3600 * 1000;

/**
 * A report with the given tables, updated by the REAL continuousUpdateCore_().
 * Only the TDoc list download, the stages after the TDoc loop and the
 * Reviewer are replaced. `o`: { props, token, reviewer: { id: text | 404 | 500 | 'throw' } }.
 */
function makeReport(o, tables) {
  const loaded = loadCode({ documentProperties: o.props || ADHOC_6G, scriptProperties: o.token === null ? {} : { REVIEWER_API_TOKEN: 'synthetic-test-token' } });
  const s = loaded.sandbox;
  const logs = [];
  s.Logger.log = (m) => logs.push(String(m));
  const body = makeFakeDocumentBody(s);
  const built = tables(s);
  built.forEach((t) => body._children.push(t));
  s.DocumentApp.getActiveDocument = () => ({ getId: () => 'DOC1', getBody: () => body });
  let list = [];
  s.downloadAndGroupTdocs_ = () => ({ '5.4': { tdocs: list } });
  s.rearrangeRevisionTables_ = () => ({ moved: 0, dispositions: 0 });
  s.checkRSSFeed_ = () => {};
  s.updateRevisions_ = () => {};
  let formatted = 0;
  s.removeRowHeightAndSpacing = () => { formatted++; };
  const requests = [];
  const reviewer = o.reviewer || {};
  s.UrlFetchApp.fetch = (url) => {
    const m = String(url).match(/^https:\/\/reviewer\.bouazizi\.dev\/api\/v1\/documents\/([^/]+)\/summary\?type=summary$/);
    if (!m) throw new Error('unexpected fetch: ' + url);
    requests.push(m[1]);
    const answer = Object.prototype.hasOwnProperty.call(reviewer, m[1]) ? reviewer[m[1]] : 404;
    if (answer === 'throw') throw new Error('Address unavailable');
    if (typeof answer === 'number') return { getResponseCode: () => answer, getContentText: () => '' };
    return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ text: answer }) };
  };
  return {
    s, body, tables: built, logs, requests, reviewer, props: loaded.docProps,
    update: (rows) => { list = rows; s.resetReviewerTokenRunState_(); return s.continuousUpdateCore_(); },
    formatted: () => formatted
  };
}
const OK = { success: true, error: null };

const LIST_2026_10_05 = () => [
  listRow('S4aP260099', 'reserved', null),
  listRow('S4aP260100', 'available', zip('S4aP260100')),
  listRow('S4aP260101', 'available', zip('S4aP260101')),
  listRow('S4aP260102', 'available', zip('S4aP260102'))
];

console.log('1. the live 6G report of 2026-10-05: abstracts on update switched OFF; the Reviewer has a summary of 100 and none of 102');
{
  const r = makeReport({ reviewer: { S4aP260100: 'Summary of 100.', S4aP260102: 404 } }, (s) => [
    makeRegistrationTable(s, [['S4aP260099', null], ['S4aP260100', null], ['S4aP260101', zip('S4aP260101')], ['S4aP260102', null]]),
    makeTdocTable(s, 'S4aP260099', 'reserved'),
    makeTdocTable(s, 'S4aP260100', 'available'),
    makeTdocTable(s, 'S4aP260101', 'available', { link: zip('S4aP260101') }),
    makeTdocTable(s, 'S4aP260102', 'available')
  ]);
  const [reg, t099, t100, t101, t102] = r.tables;
  // "No summary" answers the Reviewer gave while 100 and 102 were reserved, and one for 101 and 099.
  const stale = JSON.stringify({ ts: Date.now() - 2 * HOUR, statusCode: 404 });
  ['S4aP260099', 'S4aP260100', 'S4aP260101', 'S4aP260102'].forEach((id) => r.props.setProperty(CACHE(id), stale));
  const protectedBefore = r.tables.slice(1).map(protectedOf);
  const regBefore = textOf(reg);
  check('before: abstracts on update are off, 100 and 102 have no link, 101 has one, nobody has an abstract',
    [r.s.getFetchAbstractsSetting_(), linkOf(t100), linkOf(t101), linkOf(t102), r.tables.slice(1).map((t) => field(t, 'Abstract'))],
    [false, null, zip('S4aP260101'), null, [undefined, undefined, undefined, undefined]]);

  const result = r.update(LIST_2026_10_05());
  check('the update completes', result, OK);

  check('S4aP260100: the link of the list is on the whole TDoc number', linksOf(t100.getCell(0, 1)), 'S4aP260100'.split('').map(() => zip('S4aP260100')));
  check('S4aP260100: the Abstract the Reviewer has is in the table, after Agenda Item, although the setting is off', [labels(t100).slice(4, 7), field(t100, 'Abstract')], [['Agenda Item', 'Abstract', 'Type/For'], 'Summary of 100.']);
  check('S4aP260100: the link and the abstract row are the only writes', t100._writes, [linkWrite('S4aP260100', zip('S4aP260100'))].concat(abstractWrites()));
  check('S4aP260100: the "no summary" answer from before the upload is gone', r.props.getProperty(CACHE('S4aP260100')), null);

  check('S4aP260102: the link of the list is on the whole TDoc number', linksOf(t102.getCell(0, 1)), 'S4aP260102'.split('').map(() => zip('S4aP260102')));
  check('S4aP260102: the Reviewer has no summary: no Abstract, and the link is the only write', [field(t102, 'Abstract'), t102._writes], [undefined, [linkWrite('S4aP260102', zip('S4aP260102'))]]);
  const cached102 = JSON.parse(r.props.getProperty(CACHE('S4aP260102')));
  check('S4aP260102: the old answer was replaced by the new one: cached now, as any "no summary" is', [cached102.statusCode, Date.now() - cached102.ts < 60000], [404, true]);

  check('S4aP260101, already uploaded when it was inserted: not a transition. Its link stays, nothing is written, the Reviewer is not asked, its cache entry stays',
    [linkOf(t101), t101._writes, r.props.getProperty(CACHE('S4aP260101'))], [zip('S4aP260101'), [], stale]);
  check('S4aP260099, still reserved: no link, no abstract, nothing written, its cache entry untouched',
    [linkOf(t099), field(t099, 'Abstract'), t099._writes, r.props.getProperty(CACHE('S4aP260099'))], [null, undefined, [], stale]);
  check('the Reviewer was asked for 100 and 102, once each, and for nothing else', r.requests, ['S4aP260100', 'S4aP260102']);

  check('TDoc, Title, Source, Contact, Agenda Item, Type/For, Minutes, Disposition and Status of every table are what they were (the list has other values for five of them)',
    r.tables.slice(1).map(protectedOf), protectedBefore);
  check('Minutes are byte-identical', r.tables.slice(1).map((t) => field(t, 'Minutes')), protectedBefore.map((p) => p[6][1]));

  check('registration table: 100 and 102 got their link, 099 has none, 101 kept its own',
    [1, 2, 3, 4].map((row) => linksOf(reg.getCell(row, 0)).filter(Boolean)[0] || null), [null, zip('S4aP260100'), zip('S4aP260101'), zip('S4aP260102')]);
  check('registration table: two links are the only writes; no row was added, removed or moved and no text changed',
    [reg._writes, textOf(reg)], [[['setLinkUrl', 'S4aP260100[0]', 0, 9, zip('S4aP260100')], ['setLinkUrl', 'S4aP260102[0]', 0, 9, zip('S4aP260102')]], regBefore]);
  check('the update says what it completed', r.logs.filter((m) => /^Upload completion/.test(m)), ['Upload completion: 2 TDoc link(s) added, 2 abstract attempt(s), 2 registration-table link(s) added']);
  check('the general abstract sweep did not run (the setting is off)', r.logs.filter((m) => /Abstract fetching is disabled/.test(m)).length, 1);
  check('an abstract row was inserted, so the formatting pass ran', r.formatted(), 1);
  check('no table was added or removed', r.body.getTables().length, 5);

  console.log('   the update run again');
  const writes = r.tables.map((t) => t._writes.length);
  const cacheAfterFirst = r.props.getProperty(CACHE('S4aP260102'));
  check('the second update completes', r.update(LIST_2026_10_05()), OK);
  check('it writes to no table', r.tables.map((t) => t._writes.length), writes);
  check('it does not ask the Reviewer: 102 is not a transition any more, and its new cache entry is left alone', [r.requests, r.props.getProperty(CACHE('S4aP260102'))], [['S4aP260100', 'S4aP260102'], cacheAfterFirst]);
  check('one Abstract row in 100, not two', labels(t100).filter((l) => l === 'Abstract').length, 1);
  check('it reports nothing completed, and no formatting pass was needed', [r.logs.filter((m) => /^Upload completion/.test(m))[1], r.formatted()], ['Upload completion: 0 TDoc link(s) added, 0 abstract attempt(s), 0 registration-table link(s) added', 1]);

  console.log('   and a third time, with the Reviewer now having the summary of 102');
  r.reviewer.S4aP260102 = 'Summary of 102.';
  r.update(LIST_2026_10_05());
  check('with the setting off the update still does not ask: the upload gave one attempt, not one per update', [r.requests.length, field(t102, 'Abstract')], [2, undefined]);
  check('"Update Abstracts" follows the cache as always: within its 24 hours 102 is not asked for', (() => { r.s.addAbstractsForTables_(r.body); return r.requests.filter((id) => id === 'S4aP260102').length; })(), 1);
  r.props.setProperty(CACHE('S4aP260102'), JSON.stringify({ ts: Date.now() - 25 * HOUR, statusCode: 404 }));
  r.s.addAbstractsForTables_(r.body);
  check('and after them it is, and the abstract arrives', [r.requests.filter((id) => id === 'S4aP260102').length, field(t102, 'Abstract')], [2, 'Summary of 102.']);
}

console.log('2. the link arrives together with the status change: reserved -> available');
{
  const r = makeReport({ reviewer: { S4aP260099: 'Summary of 099.' } }, (s) => [makeTdocTable(s, 'S4aP260099', 'reserved')]);
  r.update([listRow('S4aP260099', 'available', zip('S4aP260099'))]);
  const t = r.tables[0];
  check('Status available, link and abstract, in one update', [field(t, 'Status'), linkOf(t), field(t, 'Abstract')], ['available', zip('S4aP260099'), 'Summary of 099.']);
  check('the status was written by the status logic, before and apart from the completion', t._writes.filter((w) => w[0] === 'setText' || w[0] === 'setLinkUrl'),
    [['setText', 'Status[1]', 'available'], linkWrite('S4aP260099', zip('S4aP260099'))]);
}

console.log('3. a Status decided by hand, with Minutes and Disposition');
// "withdrawn" was in this list until T-2026.10.8. No abstract is asked for a withdrawn TDoc now: it gets its link only (below,
// and tests/withdrawn-no-abstract.test.js).
['agreed', 'noted', 'endorsed', 'Agreed with changes', 'postponed', ''].forEach((manual) => {
  const r = makeReport({ reviewer: { S4aP260100: 'Summary.' } }, (s) => [makeTdocTable(s, 'S4aP260100', manual, { minutes: 'Discussed at length.\n\n  Offline until Thursday.  ', disposition: 'Agreed, see the chair\'s notes' })]);
  const t = r.tables[0];
  const before = protectedOf(t);
  const result = r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]);
  check(`Status ${JSON.stringify(manual)}: link and abstract are added; Status, Minutes, Disposition and the rest are untouched`,
    [result, linkOf(t), field(t, 'Abstract'), protectedOf(t), t._writes], [OK, zip('S4aP260100'), 'Summary.', before, [linkWrite('S4aP260100', zip('S4aP260100'))].concat(abstractWrites())]);
});

{
  const r = makeReport({ reviewer: { S4aP260100: 'Summary.' } }, (s) => [makeTdocTable(s, 'S4aP260100', 'withdrawn', { minutes: 'Discussed at length.\n\n  Offline until Thursday.  ', disposition: 'Agreed, see the chair\'s notes' })]);
  const t = r.tables[0];
  const before = protectedOf(t);
  const result = r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]);
  check('Status "withdrawn": the link is added and no abstract is asked for; Status, Minutes, Disposition and the rest are untouched',
    [result, linkOf(t), field(t, 'Abstract'), protectedOf(t), t._writes, r.requests], [OK, zip('S4aP260100'), undefined, before, [linkWrite('S4aP260100', zip('S4aP260100'))], []]);
}

console.log('4. a link the report already has is never replaced, and is not a transition');
{
  const r = makeReport({ reviewer: { S4aP260100: 'Summary.', S4aP260101: 'Summary.', S4aP260102: 'Summary.' } }, (s) => [
    makeRegistrationTable(s, [['S4aP260100', OTHER_LINK], ['S4aP260101', null], ['S4aP260102', null]], true),
    makeTdocTable(s, 'S4aP260100', 'available', { link: OTHER_LINK }),
    makeTdocTable(s, 'S4aP260102', 'reserved', { link: OTHER_LINK }),
    makeTdocTable(s, 'S4aP260101', 'available')
  ]);
  const [reg, t100, t102, partial] = r.tables;
  // A link on part of the TDoc number only is a link too; the same in the registration table.
  partial.getCell(0, 1)._links[3] = OTHER_LINK;
  reg.getCell(2, 0)._links[9] = OTHER_LINK;
  const stale = JSON.stringify({ ts: Date.now() - HOUR, statusCode: 404 });
  r.props.setProperty(CACHE('S4aP260100'), stale);

  check('the update completes', r.update(LIST_2026_10_05().slice(1)), OK);
  check('a different link stays, whole and unchanged; nothing is written', [linksOf(t100.getCell(0, 1)), t100._writes], ['S4aP260100'.split('').map(() => OTHER_LINK), []]);
  check('the same when the status changes in that update: the status is written, the link is not', [linkOf(t102), field(t102, 'Status'), t102._writes.filter((w) => w[0] === 'setLinkUrl')], [OTHER_LINK, 'available', []]);
  check('a link on part of the TDoc number is left as it is', [linksOf(partial.getCell(0, 1)).filter(Boolean), partial._writes], [[OTHER_LINK], []]);
  check('none of them is a transition: the Reviewer is not asked, no abstract appears, the cache entry stays', [r.requests, r.tables.slice(1).map((t) => field(t, 'Abstract')), r.props.getProperty(CACHE('S4aP260100'))], [[], [undefined, undefined, undefined], stale]);
  check('registration table (five columns): the different link and the partial link stay, the missing one is added, nothing else is written',
    [linksOf(reg.getCell(1, 0)).filter(Boolean)[0], linksOf(reg.getCell(2, 0)).filter(Boolean), reg._writes], [OTHER_LINK, [OTHER_LINK], [['setLinkUrl', 'S4aP260102[0]', 0, 9, zip('S4aP260102')]]]);
}

console.log('5. an Abstract the table already has is never rewritten or doubled');
{
  const r = makeReport({ reviewer: { S4aP260100: 'Summary from the Reviewer.' } }, (s) => [makeTdocTable(s, 'S4aP260100', 'available', { abstract: 'An abstract, edited by hand.' })]);
  const t = r.tables[0];
  r.props.setProperty(CACHE('S4aP260100'), JSON.stringify({ ts: Date.now() - HOUR, statusCode: 404 }));
  r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]);
  check('the link is added; the Abstract is the one that was there, once; the Reviewer is not asked',
    [linkOf(t), field(t, 'Abstract'), labels(t).filter((l) => l === 'Abstract').length, r.requests, t._writes], [zip('S4aP260100'), 'An abstract, edited by hand.', 1, [], [linkWrite('S4aP260100', zip('S4aP260100'))]]);
  check('the stale "no summary" answer is cleared all the same', r.props.getProperty(CACHE('S4aP260100')), null);
}

console.log('6. when the Reviewer cannot answer, the update still completes');
[
  ['answers 404', 404, true],
  ['answers 500', 500, false],
  ['cannot be reached', 'throw', false],
  ['answers 200 with an empty text', '', false]
].forEach(([what, answer, cached]) => {
  const r = makeReport({ reviewer: { S4aP260100: answer } }, (s) => [makeTdocTable(s, 'S4aP260100', 'available')]);
  const t = r.tables[0];
  const first = r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]);
  const second = r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]);
  check(`the Reviewer ${what}: both updates succeed, the link is there, no Abstract, one request in all, ${cached ? 'cached as "no summary"' : 'nothing cached'}`,
    [first, second, linkOf(t), field(t, 'Abstract'), r.requests, r.props.getProperty(CACHE('S4aP260100')) !== null, t._writes], [OK, OK, zip('S4aP260100'), undefined, ['S4aP260100'], cached, [linkWrite('S4aP260100', zip('S4aP260100'))]]);
});
{
  const r = makeReport({ token: null }, (s) => [makeTdocTable(s, 'S4aP260100', 'available')]);
  check('no Reviewer token: the update succeeds, the link is added, no request is made',
    [r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]), linkOf(r.tables[0]), r.requests], [OK, zip('S4aP260100'), []]);
}
{
  const r = makeReport({ reviewer: { S4aP260100: 'Summary.' } }, (s) => [makeTdocTable(s, 'S4aP260100', 'available'), makeTdocTable(s, 'S4aP260102', 'available')]);
  r.tables[0].getCell(0, 1).editAsText = () => { throw new Error('Service unavailable: Docs'); };
  const result = r.update(LIST_2026_10_05().filter((td) => /100|102/.test(td.row[0])));
  check('a link that cannot be set: the update succeeds and goes on to the next TDoc; the Reviewer is not asked for the one that failed',
    [result, linkOf(r.tables[0]), linkOf(r.tables[1]), r.requests, r.logs.filter((m) => /could not be completed/.test(m))],
    [OK, null, zip('S4aP260102'), ['S4aP260102'], ['Upload of S4aP260100: could not be completed: Service unavailable: Docs']]);
}

console.log('7. abstracts on update switched ON: the general sweep leaves out what is not uploaded');
{
  const on = Object.assign({}, ADHOC_6G, { FETCH_ABSTRACTS_ON_UPDATE: 'true' });
  const r = makeReport({ props: on, reviewer: { S4aP260101: 'Summary of 101.', S4aP260050: 'Summary of 050.' } }, (s) => [
    makeTdocTable(s, 'S4aP260099', 'reserved'),
    makeTdocTable(s, 'S4aP260101', 'available', { link: zip('S4aP260101') }),
    makeTdocTable(s, 'S4aP260050', 'noted', { link: zip('S4aP260050') })
  ]);
  const [t099, t101, t050] = r.tables;
  check('the update completes', r.update(LIST_2026_10_05().filter((td) => /099|101/.test(td.row[0]))), OK);
  check('the reserved TDoc is not sent to the Reviewer, and no "no summary" entry is created for it', [r.requests.indexOf('S4aP260099'), r.props.getProperty(CACHE('S4aP260099')), t099._writes], [-1, null, []]);
  check('an uploaded TDoc without an abstract is, as before; so is a table whose TDoc is not in the list', [r.requests, field(t101, 'Abstract'), field(t050, 'Abstract')], [['S4aP260101', 'S4aP260050'], 'Summary of 101.', 'Summary of 050.']);

  console.log('   the reserved TDoc is uploaded');
  r.reviewer.S4aP260099 = 'Summary of 099.';
  r.update([listRow('S4aP260099', 'available', zip('S4aP260099'))]);
  check('link and abstract arrive in that update, with one request and one Abstract row', [linkOf(t099), field(t099, 'Abstract'), r.requests.filter((id) => id === 'S4aP260099').length, labels(t099).filter((l) => l === 'Abstract').length], [zip('S4aP260099'), 'Summary of 099.', 1, 1]);
}
{
  const on = Object.assign({}, ADHOC_6G, { FETCH_ABSTRACTS_ON_UPDATE: 'true' });
  const r = makeReport({ props: on }, (s) => [makeTdocTable(s, 'S4aP260099', 'reserved')]);
  r.update([Object.assign(listRow('S4aP260099', 'reserved', null), { richTextRow: null })]);
  check('a list whose links could not be read names nobody "not uploaded": the sweep asks as it did before', r.requests, ['S4aP260099']);
}
{
  const r = makeReport({}, (s) => [makeTdocTable(s, 'S4aP260099', 'reserved')]);
  r.s.resetReviewerTokenRunState_();
  r.s.addAbstractsForTables_(r.body);
  check('"Update Abstracts" and Full Build (no list at hand) ask for every table without an abstract, as before', r.requests, ['S4aP260099']);
  check('tdocsNotUploadedYet_(): only what the list is known to have no link for',
    r.s.tdocsNotUploadedYet_(LIST_2026_10_05().concat([Object.assign(listRow('S4aP260103', 'reserved', null), { richTextRow: null })])), { S4aP260099: true });
}

console.log('8. the same through continuousUpdateCore_() in the other report flavours');
[
  ['a 6G ad-hoc report', ADHOC_6G, 'S4aP260100', zip('S4aP260100')],
  ['an MBS ad-hoc report', ADHOC_MBS, 'S4aI260123', 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/S4aI260123.zip'],
  ['an Audio ad-hoc TDoc', ADHOC_6G, 'S4aA260090', 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/S4aA260090.zip'],
  ['a main-meeting report', MAIN, 'S4-261234', 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Synthetic/Docs/S4-261234.zip']
].forEach(([what, props, id, url]) => {
  const r = makeReport({ props: props, reviewer: { [id]: 'Summary of ' + id }, }, (s) => [makeRegistrationTable(s, [[id, null]]), makeTdocTable(s, id, 'available')]);
  const [reg, t] = r.tables;
  r.props.setProperty(CACHE(id), JSON.stringify({ ts: Date.now() - HOUR, statusCode: 404 }));
  const before = protectedOf(t);
  const first = r.update([listRow(id, 'available', url)]);
  const second = r.update([listRow(id, 'available', url)]);
  check(`${what} (${id}): both updates complete; link, registration link and abstract are added once; one request; the rest untouched`,
    [first, second, linkOf(t), linksOf(reg.getCell(1, 0))[0], field(t, 'Abstract'), r.requests, protectedOf(t), t._writes, reg._writes.length],
    [OK, OK, url, url, 'Summary of ' + id, [id], before, [linkWrite(id, url)].concat(abstractWrites()), 1]);
});

console.log('8b. a TDoc that is already uploaded when it is inserted (the S4aP260101 case)');
{
  /** The table the update inserted for `id` (a table of the fake document: a link is recorded as [start, end, url]). */
  const inserted = (r, id) => r.body.getTables().filter((t) => t.getCell(0, 0).getText() === 'TDoc' && t.getCell(0, 1).getText() === id);
  const insertedLinks = (t) => t.getCell(0, 1)._links;
  const row101 = () => [listRow('S4aP260101', 'available', zip('S4aP260101'))];

  console.log('   abstracts on update switched OFF, the Reviewer has the summary');
  {
    const r = makeReport({ reviewer: { S4aP260101: 'Summary of 101.' } }, () => []);
    check('before: the report has no table, and abstracts on update are off', [r.body.getTables().length, r.s.getFetchAbstractsSetting_()], [0, false]);
    check('the update completes', r.update(row101()), OK);
    const t = inserted(r, 'S4aP260101')[0];
    check('the table is inserted with the link of the list', [inserted(r, 'S4aP260101').length, insertedLinks(t)], [1, [[0, 9, zip('S4aP260101')]]]);
    check('and with its Abstract, after Agenda Item; the other rows are those of the insertion', [labels(t), field(t, 'Abstract'), field(t, 'Status'), field(t, 'Minutes')],
      [['TDoc', 'Title', 'Source', 'Contact', 'Agenda Item', 'Abstract', 'Type/For', 'E-mail Discussion', 'Revisions', 'Minutes', 'Disposition', 'Status'], 'Summary of 101.', 'available', '']);
    check('one request', r.requests, ['S4aP260101']);
    check('the update says so', r.logs.filter((m) => /^Upload completion/.test(m)), ['Upload completion: 0 TDoc link(s) added, 1 abstract attempt(s), 0 registration-table link(s) added']);
    t.getCell(9, 1).setText('Minutes written after the insertion.');
    check('a second update completes, asks nothing, inserts nothing and leaves the Minutes', [r.update(row101()), r.requests, inserted(r, 'S4aP260101').length, labels(t).filter((l) => l === 'Abstract').length, field(t, 'Minutes'), insertedLinks(t)],
      [OK, ['S4aP260101'], 1, 1, 'Minutes written after the insertion.', [[0, 9, zip('S4aP260101')]]]);
  }

  console.log('   the Reviewer has no summary');
  {
    const r = makeReport({ reviewer: { S4aP260101: 404 } }, () => []);
    check('the update completes', r.update(row101()), OK);
    const t = inserted(r, 'S4aP260101')[0];
    const cached = JSON.parse(r.props.getProperty(CACHE('S4aP260101')));
    check('the table is inserted as always, with its link and without an Abstract', [inserted(r, 'S4aP260101').length, insertedLinks(t), field(t, 'Abstract'), labels(t).length], [1, [[0, 9, zip('S4aP260101')]], undefined, 11]);
    check('"no summary" is cached, as any is', [cached.statusCode, Date.now() - cached.ts < 60000], [404, true]);
    const entry = r.props.getProperty(CACHE('S4aP260101'));
    r.reviewer.S4aP260101 = 'Summary of 101.';
    check('the next updates do not ask again, and leave the cache entry', [r.update(row101()), r.update(row101()), r.requests, field(t, 'Abstract'), r.props.getProperty(CACHE('S4aP260101'))], [OK, OK, ['S4aP260101'], undefined, entry]);
  }

  console.log('   a TDoc that is inserted while reserved');
  {
    const r = makeReport({ reviewer: { S4aP260099: 'Summary of 099.' } }, () => []);
    check('the update completes', r.update([listRow('S4aP260099', 'reserved', null)]), OK);
    const t = inserted(r, 'S4aP260099')[0];
    check('it is inserted without a link and without an Abstract; the Reviewer is not asked and nothing is cached',
      [inserted(r, 'S4aP260099').length, insertedLinks(t), field(t, 'Abstract'), r.requests, r.props.getProperty(CACHE('S4aP260099'))], [1, [], undefined, [], null]);
    check('the update reports no attempt', r.logs.filter((m) => /^Upload completion/.test(m)), ['Upload completion: 0 TDoc link(s) added, 0 abstract attempt(s), 0 registration-table link(s) added']);
  }

  console.log('   a "no summary" answer that is already cached is respected');
  {
    const r = makeReport({ reviewer: { S4aP260101: 'Summary of 101.' } }, () => []);
    const entry = JSON.stringify({ ts: Date.now() - HOUR, statusCode: 404 });
    r.props.setProperty(CACHE('S4aP260101'), entry);
    r.update(row101());
    check('no request, no Abstract, the entry stays (the insertion is not a reason to forget it)', [r.requests, field(inserted(r, 'S4aP260101')[0], 'Abstract'), r.props.getProperty(CACHE('S4aP260101'))], [[], undefined, entry]);
  }

  console.log('   abstracts on update switched ON: no second request in the same update');
  const on = Object.assign({}, ADHOC_6G, { FETCH_ABSTRACTS_ON_UPDATE: 'true' });
  [['has the summary', 'Summary of 101.', 'Summary of 101.'], ['answers 404', 404, undefined], ['answers 500 (nothing cached)', 500, undefined],
    ['cannot be reached (nothing cached)', 'throw', undefined], ['answers 200 with an empty text (nothing cached)', '', undefined]].forEach(([what, answer, abstract]) => {
    const r = makeReport({ props: on, reviewer: { S4aP260101: answer } }, () => []);
    const result = r.update(row101());
    check(`inserted as uploaded, the Reviewer ${what}: one request in that update`, [result, r.requests, field(inserted(r, 'S4aP260101')[0], 'Abstract')], [OK, ['S4aP260101'], abstract]);
  });
  [['answers 500 (nothing cached)', 500], ['cannot be reached (nothing cached)', 'throw']].forEach(([what, answer]) => {
    const r = makeReport({ props: on, reviewer: { S4aP260100: answer } }, (s) => [makeTdocTable(s, 'S4aP260100', 'available')]);
    const result = r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]);
    check(`an existing table at its transition, the Reviewer ${what}: one request in that update too`, [result, r.requests], [OK, ['S4aP260100']]);
    r.update([listRow('S4aP260100', 'available', zip('S4aP260100'))]);
    check('the next update asks through the ordinary sweep (the setting is on), once', r.requests, ['S4aP260100', 'S4aP260100']);
  });

  console.log('   both ways in one update: 100 and 102 at their transition, 101 inserted as uploaded, 099 inserted as reserved');
  {
    const r = makeReport({ reviewer: { S4aP260100: 'Summary of 100.', S4aP260101: 'Summary of 101.', S4aP260102: 404 } }, (s) => [
      makeTdocTable(s, 'S4aP260100', 'available'),
      makeTdocTable(s, 'S4aP260102', 'available')
    ]);
    const [t100, t102] = r.tables;
    const before = [protectedOf(t100), protectedOf(t102)];
    check('the update completes', r.update(LIST_2026_10_05()), OK);
    check('100 and 102 are completed exactly as when they are alone', [t100._writes, t102._writes, field(t100, 'Abstract'), field(t102, 'Abstract'), [protectedOf(t100), protectedOf(t102)]],
      [[linkWrite('S4aP260100', zip('S4aP260100'))].concat(abstractWrites()), [linkWrite('S4aP260102', zip('S4aP260102'))], 'Summary of 100.', undefined, before]);
    check('101 is inserted with link and Abstract, 099 without either', [insertedLinks(inserted(r, 'S4aP260101')[0]), field(inserted(r, 'S4aP260101')[0], 'Abstract'), insertedLinks(inserted(r, 'S4aP260099')[0]), field(inserted(r, 'S4aP260099')[0], 'Abstract')],
      [[[0, 9, zip('S4aP260101')]], 'Summary of 101.', [], undefined]);
    check('three requests, in the order of the list; none for 099', r.requests, ['S4aP260100', 'S4aP260101', 'S4aP260102']);
    check('the update says so', r.logs.filter((m) => /^Upload completion/.test(m)), ['Upload completion: 2 TDoc link(s) added, 3 abstract attempt(s), 0 registration-table link(s) added']);
    check('a second update asks nothing', [r.update(LIST_2026_10_05()), r.requests.length], [OK, 3]);
  }
}

console.log('9. the helpers');
{
  const s = loadCode({ documentProperties: ADHOC_6G }).sandbox;
  const cellOf = (text, link) => { const t = makeTable(s, [['x', text]]); if (link) t.getCell(0, 1)._links[0] = link; return { cell: t.getCell(0, 1), writes: t._writes }; };
  const plain = cellOf('S4aP260100');
  check('addMissingTdocLink_(): sets the link and says so; the second time it does nothing', [s.addMissingTdocLink_(plain.cell, 'S4aP260100', zip('S4aP260100')), s.addMissingTdocLink_(plain.cell, 'S4aP260100', OTHER_LINK), plain.writes.length], [true, false, 1]);
  const padded = cellOf(' S4aP260100 ');
  s.addMissingTdocLink_(padded.cell, 'S4aP260100', zip('S4aP260100'));
  check('spaces around the number: the link is on the number, not on the spaces', padded.writes, [['setLinkUrl', 'x[1]', 1, 10, zip('S4aP260100')]]);
  const more = cellOf('S4aP260100 (late)');
  check('a cell that holds more than the number, another number, or no url: nothing is done',
    [s.addMissingTdocLink_(more.cell, 'S4aP260100', zip('S4aP260100')), s.addMissingTdocLink_(cellOf('S4aP260101').cell, 'S4aP260100', zip('S4aP260100')), s.addMissingTdocLink_(cellOf('S4aP260100').cell, 'S4aP260100', ''), more.writes], [false, false, false, []]);
  check('tdocListLinkUrl_(): the link of the list cell, read as for a new table; \'\' when there is none or it cannot be read',
    [s.tdocListLinkUrl_(listRow('X', 'available', zip('X'))), s.tdocListLinkUrl_(listRow('X', 'reserved', null)), s.tdocListLinkUrl_({ richTextRow: null, tdocCol: 0 }), s.tdocListLinkUrl_({ richTextRow: [{}], tdocCol: 0 }), s.tdocListLinkUrl_(null)],
    [zip('X'), '', '', '', '']);
  const none = makeTdocTable(s, 'S4aP260100', 'available');
  const body = { getTables: () => [none] };
  check('refreshUploadedTdocMetadata_(): a TDoc without a table, and a list without a link: nothing done',
    [s.refreshUploadedTdocMetadata_(body, 'S4aP260555', listRow('S4aP260555', 'available', zip('S4aP260555')), s.buildTdocTableIndex_([none])),
      s.refreshUploadedTdocMetadata_(body, 'S4aP260100', listRow('S4aP260100', 'reserved', null), s.buildTdocTableIndex_([none])), none._writes],
    [{ linkAdded: false, abstractAttempted: false }, { linkAdded: false, abstractAttempted: false }, []]);
  check('refreshRegistrationTableLinks_(): a report without a registration table: 0, no error', s.refreshRegistrationTableLinks_({ getTables: () => [none] }, LIST_2026_10_05()), 0);
}

console.log('10. source: where the completion happens, and what was left alone');
{
  const core = functionSource(CODE, 'continuousUpdateCore_');
  check('continuousUpdateCore_() completes an upload in its existing-TDoc branch, once, after the status update',
    [(core.match(/refreshUploadedTdocMetadata_\(/g) || []).length,
      /updateTdocStatus_\(body, tdocNumber, tdocData, tdocTableIndex\)\)\) \{\s*statusUpdated\+\+;\s*\}[\s\S]{0,400}?refreshUploadedTdocMetadata_\(body, tdocNumber, tdocData, tdocTableIndex, context\)\);[\s\S]{0,400}?\n      \}\n    \}\);/.test(core)], [1, true]);
  check('and asks for the abstract of a TDoc inserted as uploaded in its new-TDoc branch, once, after the insertion',
    [(core.match(/completeInsertedUploadedTdoc_\(/g) || []).length, (withoutComments(CODE).match(/completeInsertedUploadedTdoc_\(/g) || []).length,
      /insertNewTdoc_\(body, tdocData, cfg, tdocTableIndex, context\)\);[\s\S]{0,700}?newTdocsAdded\+\+;[\s\S]{0,400}?completeInsertedUploadedTdoc_\(body, tdocNumber, tdocData, tdocTableIndex, context\)\)\) \{[\s\S]{0,160}?\} else \{/.test(core)], [1, 2, true]);
  check('nothing else calls it, and the registration links are completed once per update', [(withoutComments(CODE).match(/refreshUploadedTdocMetadata_\(/g) || []).length, (withoutComments(CODE).match(/refreshRegistrationTableLinks_\(/g) || []).length], [2, 2]);
  check('the changelog has the entry of 2.18.1, the release of the upload completion (Code.js is 2.21.0 now)', [(CODE.match(/^ \* Version: (\d+\.\d+\.\d+) \((\d{4}-\d{2}-\d{2})\)/m) || []).slice(1), /\n \* 2\.18\.1 \(2026-10-05\)\n \*   - Fixed: a TDoc that entered the report while it was only reserved\n/.test(CODE)], [['2.21.0', '2026-10-06'], true]);
  // The status dropdowns (T-2026.10.7) came after the upload completion. What they added to released functions, taken out again:
  const DROPDOWN_HOOKS = [
    "  // T-2026.10.7: a Status that is a dropdown is never written as text; the same rule is applied to its selected value.\n  if (statusCellHoldsDropdown_(statusInfo.cell)) return applyTdocStatusUpdateToDropdown_(table, tdocNumber, newStatus);\n",
    "  // T-2026.10.7: its Status is text now; the end of the update makes it a dropdown where the report has them.\n  noteStatusDropdownCandidate_(row[tdocData.tdocCol], tdocData.statusCol >= 0 ? row[tdocData.statusCol] : '');\n",
    "      // T-2026.10.7: in a report with status dropdowns this text becomes one at the end of the run, if it has an option.\n      noteStatusDropdownCandidate_(tdocNumber, newStatus);\n"];
  const withoutDropdownHooks = (source) => DROPDOWN_HOOKS.reduce((text, hook) => (text === null ? text : text.replace(hook, '')), source);
  // The release this hotfix is made on: T-2026.10.5, Code.js 2.18.0.
  const RELEASED = gitShow('template-release/T-2026.10.5:Code.js');
  if (!RELEASED) {
    console.log('  note: template-release/T-2026.10.5 is not available in this checkout; the comparisons with it are skipped.');
  } else {
    // After T-2026.10.8 the Reviewer request has one guard at its head (no abstract for a withdrawn TDoc) and, for it, a
    // fourth parameter. Taken out again, the function is the released one.
    const WITHDRAWN_GUARD = /  \/\/ Nothing is asked for a withdrawn TDoc \(abstractFetchBlockedBy_\(\)\): this\n[\s\S]*?\n    return false;\n  \}\n/;
    const withoutWithdrawnGuard = (source) => (source === null ? source : source.replace('function fetchAndAddAbstract_(table, tdocNumber, context, portalStatus) {\n', 'function fetchAndAddAbstract_(table, tdocNumber, context) {\n').replace(WITHDRAWN_GUARD, ''));
    check('the guard of fetchAndAddAbstract_() is there, once, and is the first thing it does',
      [(functionSource(CODE, 'fetchAndAddAbstract_').match(WITHDRAWN_GUARD) || []).length, functionSource(CODE, 'fetchAndAddAbstract_').indexOf('  // Nothing is asked for a withdrawn TDoc') === 'function fetchAndAddAbstract_(table, tdocNumber, context, portalStatus) {\n'.length], [1, true]);
    const same = (names) => names.filter((name) => functionSource(CODE, name) === null || withoutWithdrawnGuard(withoutDropdownHooks(functionSource(CODE, name))) !== functionSource(RELEASED, name));
    check('the status logic is the released one, byte for byte (but for the dropdown branch of T-2026.10.7)',
      same(['applyTdocStatusUpdate_', 'updateTdocStatus_', 'normalizeStatus_', 'findStatusInDocTable_', 'styleStatusCell_']), []);
    check('so are the Reviewer request (but for its guard for withdrawn TDocs), its parsing, the Abstract row and the "no summary" cache',
      same(['fetchAndAddAbstract_', 'findAbstractInsertIndex_', 'isReviewerNoSummaryCached_', 'markReviewerNoSummary_', 'clearReviewerNoSummaryCache_', 'getFetchAbstractsSetting_']), []);
    check('and the insertion of a new TDoc (but for the note of T-2026.10.7), the registration table and the revision placement',
      same(['insertNewTdoc_', 'insertTDocTableAtIndex_', 'updateRegisteredDocumentsTable_', 'findRegistrationTable_', 'rearrangeRevisionTables_', 'setDispositionRevisedTo_', 'collectorUpdate_']), []);
    const released = functionNames(RELEASED);
    // The eleven other functions are those the status dropdowns changed (tests/status-dropdown.test.js, tests/adhoc-sessions-config.test.js).
    const DROPDOWN_FUNCTIONS = ['insertNewTdoc_', 'applyTdocStatusUpdate_', 'updateStatusWithStrictRules_', 'insertRevisedDocTablesAfter_', 'analyzeReportStatus', 'runFullReportBuildCore_', 'buildSkeletonWithTdocTables',
      'appendTdocDetailTable_', 'collectRevisionsOnly', 'detectTdocTablesInDocument_', 'generateTdocDiscussionEmails'];
    // After T-2026.10.8 eleven more changed, for other reasons than the upload completion (tests/after-10-8-ledger.test.js).
    const AFTER_10_8_FUNCTIONS = ['fetchAndAddAbstract_', 'persistConfigurationSettings_', 'computeResolvedMeetingPreview_', 'configureMeetingSettings', 'saveConfigurationSettings',
      'getReportConfig_', 'testAllConnections', 'testReviewerApi', 'validateConfiguration', 'getReviewerApiTokenForRun_', 'resetReviewerTokenRunState_'];
    check('of the released functions, the upload completion changed exactly two: the update and the abstract sweep',
      released.filter((name, i, list) => list.indexOf(name) === i && functionSource(CODE, name) !== functionSource(RELEASED, name) && DROPDOWN_FUNCTIONS.indexOf(name) === -1 && AFTER_10_8_FUNCTIONS.indexOf(name) === -1), ['continuousUpdateCore_', 'addAbstractsForTables_']);
    check('and those eleven did change', AFTER_10_8_FUNCTIONS.filter((name) => functionSource(CODE, name) === functionSource(RELEASED, name)), []);
    // The functions of the section after T-2026.10.8 (below) are not of the upload completion.
    const after108Section = CODE.slice(CODE.indexOf('// =========================================================\n// AFTER T-2026.10.8 -- WITHDRAWN TDOCS AND ABSTRACTS; THE START OF AN AD-HOC MEETING\n'), CODE.indexOf('// =========================================================\n// TDOC STATUS DROPDOWNS'));
    const AFTER_10_8_NEW = functionNames(after108Section);
    check('the section after T-2026.10.8 has its fourteen functions', AFTER_10_8_NEW.length, 14);
    check('the new functions are the seven of the upload completion; none was removed',
      [functionNames(CODE).filter((name) => released.indexOf(name) === -1 && !/[sS]tatusD(ropdown|ocs)|^readTdocStatus_$|^statusCellHoldsDropdown_$|^mapPortalStatusToDropdownOption_$|^applyTdocStatusUpdateToDropdown_$|^(convert|migrate)StatusFieldsToDropdowns_?$/.test(name) && AFTER_10_8_NEW.indexOf(name) === -1).sort(), released.filter((name) => functionNames(CODE).indexOf(name) === -1)],
      [['addMissingTdocLink_', 'completeInsertedUploadedTdoc_', 'refreshRegistrationTableLinks_', 'refreshUploadedTdocMetadata_', 'tdocListLinkIsKnown_', 'tdocListLinkUrl_', 'tdocsNotUploadedYet_'].sort(), []]);
    // The section of the status dropdowns stands directly before the upload-completion section; both are left out.
    // So is the section of the work after T-2026.10.8, which stands directly before that of the status dropdowns.
    const withoutSection = CODE.replace(CODE.slice(CODE.indexOf('// =========================================================\n// AFTER T-2026.10.8 -- WITHDRAWN TDOCS AND ABSTRACTS; THE START OF AN AD-HOC MEETING\n'), CODE.indexOf('// =========================================================\n// AD-HOC SESSIONS (stage A)')), '');
    const outsideFunctions = (src) => functionNames(src).filter((name, i, list) => list.indexOf(name) === i).reduce((text, name) => text.replace(functionSource(src, name), ''), src);
    // The header: the version line and the changelog entry of 2.18.1 are new; the rest of it is unchanged.
    const withoutRelease = withoutSection.replace(/ \* 2\.21\.0 \(2026-10-06\)\n[\s\S]*? \* 2\.18\.0 \(2026-10-02\)\n/, ' * 2.18.0 (2026-10-02)\n').replace(' * Version: 2.21.0 (2026-10-06)\n', ' * Version: 2.18.0 (2026-10-02)\n');
    check('outside its functions, the upload-completion, status-dropdown and after-T-2026.10.8 sections, the version line and the changelog entries since 2.18.0, Code.js is the released file',
      [outsideFunctions(withoutRelease) === outsideFunctions(RELEASED), withoutRelease === withoutSection], [true, false]);
  }
  const section = CODE.slice(CODE.indexOf('// TDOC UPLOAD COMPLETION'), CODE.indexOf('// AD-HOC SESSIONS (stage A)'));
  const sectionCode = withoutComments(section);
  // After T-2026.10.8 it asks ONE thing about a status: whether the TDoc is withdrawn (abstractFetchBlockedBy_(), with the
  // status of the TDoc list). Those three names taken out, it still names no Status -- and it writes none (next check).
  const withdrawnRuleUses = (sectionCode.match(/tdocListStatus_\(tdocData\)|abstractFetchBlockedBy_\(table, portalStatus\)|\bportalStatus\b/g) || []).length;
  check('the completion code asks whether the TDoc is withdrawn, through the one rule, in its two ways', withdrawnRuleUses, 5);
  check('otherwise the completion code does not read or write a Status, Minutes, Disposition, Title, Source, Contact or Type/For',
    /status|minutes|disposition|title|source|contact|type\/for/i.test(sectionCode.replace(/tdocListStatus_\(tdocData\)|abstractFetchBlockedBy_\(table, portalStatus\)|\bportalStatus\b/g, '')), false);
  check('it sets a link, and nothing else in a table: no setText, no row, no style', /setText|appendTableRow|insertTableRow|removeRow|appendTableCell|setBold|setForegroundColor|setFontSize|styleStatusCell_/.test(sectionCode), false);
  check('it has no Reviewer request or parser of its own: the abstract goes through fetchAndAddAbstract_() only', [/UrlFetchApp|reviewer\.bouazizi|JSON\.parse/.test(sectionCode), (sectionCode.match(/fetchAndAddAbstract_\(/g) || []).length], [false, 2]);
  check('it does not look at FETCH_ABSTRACTS_ON_UPDATE, and stores nothing', /getFetchAbstractsSetting_|FETCH_ABSTRACTS_ON_UPDATE|setProperty/.test(sectionCode), false);
  check('the sweep leaves out a TDoc only when the update names it', /if \(leaveOut && leaveOut\[tdocNumber\] === true\) return;/.test(functionSource(CODE, 'addAbstractsForTables_')), true);
  check('"Update Abstracts" and Full Build call the sweep without a list, as before', (CODE.match(/addAbstractsForTables_\(\);/g) || []).length, 2);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll TDoc upload completion checks passed.');
process.exitCode = failures ? 1 : 0;
