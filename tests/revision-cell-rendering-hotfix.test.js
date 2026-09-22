/**
 * HOTFIX-001 — Deterministic revision cell rendering.
 *
 * Live production failure: after running Collect Revisions against the
 * real meeting 86178 report, Revisions cells for tables with multiple real
 * draft files (S4aP260068, S4aP260075 -- both confirmed to have multiple
 * real contributor-tagged draft variants) showed repeated/concatenated
 * filenames and broken hyperlink ranges.
 *
 * Root cause: the ORIGINAL rendering loop called `cell.editAsText()` once,
 * then repeatedly called `te.appendText(line + '\n')` per revision,
 * re-querying `te.getText().length` for each new offset and calling
 * `te.setLinkUrl()` immediately after each append -- interleaving live
 * text MUTATION with offset CALCULATION and hyperlink APPLICATION against
 * a Text reference that kept changing underneath those calls. This suite
 * exists because tests/revision-drafts-folder.test.js's fake DocumentApp
 * was too well-behaved (a simple, purely-additive string buffer) to
 * exercise this failure mode: its assertions used loose
 * `.split('\n').filter(Boolean)` comparisons, which tolerate missing
 * separators, extra characters, and duplicate/garbled text as long as
 * SOME lines happen to still split out correctly. This suite instead
 * asserts the cell's exact, byte-for-byte final text and exact hyperlink
 * offsets -- the level of precision the live bug actually violated.
 *
 * The fix (renderRevisionsCellContent_(), Code.js): build the ENTIRE final
 * cell text as one plain JavaScript string FIRST, replace the cell's
 * contents with exactly one cell.setText(fullText) call, and ONLY THEN
 * compute each line's start/end offsets via plain string arithmetic
 * against that already-final, now-stable text before applying
 * setLinkUrl() -- no further text mutation happens after setText(), so no
 * offset can ever be invalidated by a later append.
 *
 * Run: node tests/revision-cell-rendering-hotfix.test.js
 */

const { loadCode } = require('./helpers/load-code.js');

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

// ---- fake Table/Cell, same minimal surface as tests/revision-drafts-folder.test.js

function makeFakeTextEditor(state) {
  return {
    getText: () => state.text,
    appendText: (s) => { state.text += s; },
    setLinkUrl: (start, end, url) => { state.links.push({ start, end, url }); },
    setForegroundColor: () => {},
    setBold: () => {}
  };
}

function makeFakeCell(initialText) {
  const state = { text: String(initialText === undefined || initialText === null ? '' : initialText), links: [] };
  return {
    getText: () => state.text,
    setText: (t) => { state.text = String(t || ''); state.links = []; },
    editAsText: () => makeFakeTextEditor(state),
    _state: state
  };
}

function makeFakeRow(cellTexts) {
  const cells = cellTexts.map(t => makeFakeCell(t));
  return {
    getNumCells: () => cells.length,
    getCell: (i) => cells[i],
    appendTableCell: (t) => { const c = makeFakeCell(t); cells.push(c); return c; },
    _cells: cells
  };
}

function makeFakeTable(rows) {
  const _rows = rows.map(r => makeFakeRow(r));
  return {
    getType: () => 'TABLE',
    getNumRows: () => _rows.length,
    getRow: (i) => _rows[i],
    getCell: (r, c) => _rows[r].getCell(c),
    appendTableRow: () => { const row = makeFakeRow([]); _rows.push(row); return row; },
    removeRow: (i) => { _rows.splice(i, 1); },
    _rows
  };
}

function makeFakeBody(children) {
  const _children = children.slice();
  return {
    _children,
    getNumChildren: () => _children.length,
    getChild: (i) => _children[i],
    getChildIndex: (child) => _children.indexOf(child),
    getTables: () => _children.filter(c => c.getType() === 'TABLE'),
    insertTable: (idx, data) => { const t = makeFakeTable(data); _children.splice(idx, 0, t); return t; }
  };
}

function makeTdocTable(tdoc, agendaItem, revisionsCellInitialText) {
  return makeFakeTable([
    ['TDoc', tdoc],
    ['Title', 'Some title'],
    ['Source', 'Some source'],
    ['Contact', ''],
    ['Agenda Item', agendaItem || '5.1'],
    ['E-mail Discussion', ''],
    ['Revisions', revisionsCellInitialText || ''],
    ['Minutes', ''],
    ['Disposition', ''],
    ['Status', 'Available']
  ]);
}

function stubUrlFetch(sandbox, urlToHtml) {
  sandbox.UrlFetchApp = {
    fetch: (url) => {
      const html = urlToHtml[url];
      if (html === undefined) throw new Error('Unexpected fetch URL in test: ' + url);
      return { getResponseCode: () => 200, getContentText: () => html };
    }
  };
}

const AD_HOC_URL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/';

// ============================ 1. renderRevisionsCellContent_() directly ====

console.log('renderRevisionsCellContent_() -- exact text and exact hyperlink offsets, direct unit test');

{
  const { sandbox } = loadCode();
  const cell = makeFakeCell('');
  const ordered = [
    { text: 'S4aP260071_BBC.docx', link: 'https://x/drafts/S4aP260071_BBC.docx' },
    { text: 'S4aP260071_QCOM.docx', link: 'https://x/drafts/S4aP260071_QCOM.docx' }
  ];
  sandbox.renderRevisionsCellContent_(cell, ordered);

  check('exact final text for two drafts (byte-for-byte, not just split-and-compare)',
    cell.getText(), 'S4aP260071_BBC.docx\nS4aP260071_QCOM.docx');
  check('exactly 2 logical lines', cell.getText().split('\n').length, 2);
  check('no concatenation: the first filename does not run directly into the second',
    /docx[A-Za-z]/.test(cell.getText()), false);

  check('exactly 2 hyperlink ranges were applied', cell._state.links.length, 2);
  const [link1, link2] = cell._state.links;
  check('link 1 covers exactly "S4aP260071_BBC.docx" (offsets 0..19 inclusive)',
    [link1.start, link1.end, link1.url],
    [0, 'S4aP260071_BBC.docx'.length - 1, 'https://x/drafts/S4aP260071_BBC.docx']);
  check('link 1 range, applied to the actual final text, reproduces exactly the first filename',
    cell.getText().slice(link1.start, link1.end + 1), 'S4aP260071_BBC.docx');
  const expectedLink2Start = 'S4aP260071_BBC.docx'.length + 1; // + the '\n'
  check('link 2 starts exactly after the separator (no off-by-one)', link2.start, expectedLink2Start);
  check('link 2 range, applied to the actual final text, reproduces exactly the second filename',
    cell.getText().slice(link2.start, link2.end + 1), 'S4aP260071_QCOM.docx');
  check('link 1 and link 2 ranges do not overlap', link1.end < link2.start, true);
}

// =========== 2. a case resembling the live S4aP260068/S4aP260075 failure ===

console.log('renderRevisionsCellContent_() -- multiple real-shaped contributor-tagged drafts (S4aP260068-style), exact text');

{
  const { sandbox } = loadCode();
  const cell = makeFakeCell('');
  // Real evidence (ARCH-011/PROD-017 investigation): S4aP260068 has
  // multiple real contributor-tagged variants.
  const ordered = [
    { text: 'S4aP260068_docomo.docx', link: 'https://x/drafts/S4aP260068_docomo.docx' },
    { text: 'S4aP260068_huawei.docx', link: 'https://x/drafts/S4aP260068_huawei.docx' },
    { text: 'S4aP260068_huawei_v2.docx', link: 'https://x/drafts/S4aP260068_huawei_v2.docx' }
  ];
  sandbox.renderRevisionsCellContent_(cell, ordered);

  check('exact final text for three drafts, sorted, one per line, no duplication',
    cell.getText(),
    'S4aP260068_docomo.docx\nS4aP260068_huawei.docx\nS4aP260068_huawei_v2.docx');
  check('exactly 3 logical lines', cell.getText().split('\n').length, 3);
  check('each filename occurs exactly once in the final text',
    ordered.every(item => cell.getText().split(item.text).length - 1 === 1), true);
  check('exactly 3 hyperlinks, each range reproduces its own filename exactly',
    cell._state.links.map(l => cell.getText().slice(l.start, l.end + 1)),
    ordered.map(o => o.text));
  check('each hyperlink URL matches its own filename\'s intended target',
    cell._state.links.map(l => l.url), ordered.map(o => o.link));
}

// ==================== 3. already-corrupted cell is repaired, not appended ===

console.log('renderRevisionsCellContent_() -- a pre-existing corrupted cell is fully REPLACED, never appended to');

{
  const { sandbox } = loadCode();
  // Simulates exactly the reported live symptom: repeated/concatenated
  // filenames with no clean separation, left over from the old buggy loop.
  const corrupted = makeFakeCell('S4aP260068_docomo.docxS4aP260068_docomo.docxS4aP260068_huawei.docx');
  const ordered = [
    { text: 'S4aP260068_docomo.docx', link: 'https://x/drafts/S4aP260068_docomo.docx' },
    { text: 'S4aP260068_huawei.docx', link: 'https://x/drafts/S4aP260068_huawei.docx' }
  ];
  sandbox.renderRevisionsCellContent_(corrupted, ordered);

  check('the corrupted starting content is entirely gone -- the cell is fully replaced, not appended to',
    corrupted.getText(), 'S4aP260068_docomo.docx\nS4aP260068_huawei.docx');
  check('no leftover duplicate substring from the corrupted state remains',
    corrupted.getText().split('S4aP260068_docomo.docx').length - 1, 1);
}

// ==================== 4. second Collect Revisions run is idempotent =========

console.log('updateRevisions_() -- running Collect Revisions twice produces byte-identical cell content');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_URL } });
  const parentTable = makeTdocTable('S4aP260068', '5.6.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };

  const html = '<html><body>' +
    '<a href="S4aP260068_docomo.docx">S4aP260068_docomo.docx</a>' +
    '<a href="S4aP260068_huawei.docx">S4aP260068_huawei.docx</a>' +
    '</body></html>';
  stubUrlFetch(sandbox, { [AD_HOC_URL.replace(/\/$/, '') + '/']: html });

  sandbox.updateRevisions_(sandbox.getCollectorConfig_());
  const firstRunText = parentTable.getCell(6, 1).getText();
  const firstRunLinks = JSON.stringify(parentTable.getCell(6, 1)._state.links);

  check('first run: exact expected text', firstRunText, 'S4aP260068_docomo.docx\nS4aP260068_huawei.docx');

  sandbox.updateRevisions_(sandbox.getCollectorConfig_());
  const secondRunText = parentTable.getCell(6, 1).getText();
  const secondRunLinks = JSON.stringify(parentTable.getCell(6, 1)._state.links);

  check('second run: text is byte-identical to the first run', secondRunText, firstRunText);
  check('second run: hyperlink ranges are identical to the first run', secondRunLinks, firstRunLinks);
  check('second run: still exactly 2 logical lines (no growth from re-running)', secondRunText.split('\n').length, 2);
}

// ============= 5. a REVIS_ store carrying pre-existing (already-corrupted-run)
//                 entries still renders clean, deduplicated output ==========

console.log('updateRevisions_() -- an existing REVIS_ store from a prior (buggy) run is cleaned on the next Collect Revisions, no manual cleanup required');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_URL } });
  // Simulate a REVIS_ store that already has the two real entries (as the
  // OLD, buggy run would have discovered and persisted them -- the store
  // itself was never the corruption source, only the rendering was) plus
  // the cell already showing corrupted, concatenated text from that old run.
  docProps._store['REVIS_S4aP260068'] = JSON.stringify({
    [AD_HOC_URL + 'S4aP260068_docomo.docx']: { text: 'S4aP260068_docomo.docx', link: AD_HOC_URL + 'S4aP260068_docomo.docx' },
    [AD_HOC_URL + 'S4aP260068_huawei.docx']: { text: 'S4aP260068_huawei.docx', link: AD_HOC_URL + 'S4aP260068_huawei.docx' }
  });

  const parentTable = makeTdocTable('S4aP260068', '5.6.1', 'S4aP260068_docomo.docxS4aP260068_docomo.docxS4aP260068_huawei.docx');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };

  // This run's live directory fetch finds nothing NEW -- the two entries
  // above are already in the (correct, un-corrupted) REVIS_ store from a
  // prior successful discovery; only the CELL TEXT was left corrupted by
  // the old buggy rendering loop. Proves the repair comes from re-
  // rendering the store, not from re-discovering the files.
  stubUrlFetch(sandbox, { [AD_HOC_URL.replace(/\/$/, '') + '/']: '<html><body></body></html>' });

  check('BEFORE Collect Revisions: the cell shows the corrupted, concatenated text (simulating the live bug)',
    parentTable.getCell(6, 1).getText().indexOf('docxS4aP') !== -1, true);

  sandbox.updateRevisions_(sandbox.getCollectorConfig_());

  check('AFTER Collect Revisions: the corrupted cell is automatically repaired -- the user did not have to clean it manually',
    parentTable.getCell(6, 1).getText(), 'S4aP260068_docomo.docx\nS4aP260068_huawei.docx');
}

// ================ 6. no duplicate rendering from persisted REVIS_ entries ==

console.log('normalizeRevisionKey_() -- store-level dedup, no duplicate rendering even from encoding-variant keys');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_URL } });
  // A store with the SAME logical file stored under two different literal
  // URL-encoding keys from separate historical runs (e.g. a space encoded
  // as "%20" one run, a literal space captured a different way another
  // run) -- must still render as ONE line, not two.
  docProps._store['REVIS_S4aP260071'] = JSON.stringify({
    'https://x/drafts/S4aP260071%20QCOM.docx': { text: 'S4aP260071 QCOM.docx', link: 'https://x/drafts/S4aP260071%20QCOM.docx' },
    'https://x/drafts/S4aP260071 QCOM.docx': { text: 'S4aP260071 QCOM.docx', link: 'https://x/drafts/S4aP260071%20QCOM.docx' }
  });

  const parentTable = makeTdocTable('S4aP260071', '5.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  stubUrlFetch(sandbox, { [AD_HOC_URL.replace(/\/$/, '') + '/']: '<html><body></body></html>' }); // no new anchors this run -- purely testing store-level dedup

  sandbox.updateRevisions_(sandbox.getCollectorConfig_());

  const lines = parentTable.getCell(6, 1).getText().split('\n').filter(Boolean);
  check('two encoding-variant store entries for the SAME URL render as exactly ONE line', lines.length, 1);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
