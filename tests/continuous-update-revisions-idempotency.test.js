/**
 * POST-MEETING-001 (Task B) — proves the HOTFIX-001 deterministic revision
 * rendering fix also applies through the AUTOMATIC update path
 * (continuousUpdate() -> collectorUpdate_() -> updateRevisions_()), not
 * just the manual "Collect Revisions" menu item
 * (collectRevisionsOnly() -> updateRevisions_()).
 *
 * B1 (source-level trace, confirmed by reading Code.js):
 *
 *   Manual:    "Collect Revisions" menu -> collectRevisionsOnly()
 *                -> updateRevisions_(getCollectorConfig_())
 *   Automatic: continuousUpdate() -> (TDoc download/insert/status-update/
 *                summary-table/rearrangeRevisionTables_/optional abstracts,
 *                all unrelated to revisions-cell rendering, see this
 *                task's report for the full continuousUpdate() risk
 *                characterization) -> collectorUpdate_()
 *                -> checkRSSFeed_(cfg) [unrelated, email discussion]
 *                -> updateRevisions_(cfg)
 *
 * Both updateRevisions_() call sites are THE SAME function, calling THE
 * SAME renderRevisionsCellContent_() (HOTFIX-001) -- there is only one
 * implementation, not two independently-maintained copies. This suite
 * proves the automatic path's specific entry point (collectorUpdate_(),
 * which is what continuousUpdate() itself calls) still produces
 * deterministic, idempotent, non-duplicated Revisions cells across
 * repeated invocations -- exactly what an automatic trigger firing
 * repeatedly over the course of a live meeting would do.
 *
 * continuousUpdate() itself is NOT executed end-to-end here (it requires
 * a much heavier DocumentApp/TDoc-list-download simulation touching TDoc
 * grouping/report-skeleton territory, explicitly out of scope for this
 * task) -- collectorUpdate_() is the correct, minimal, safe boundary: it
 * is the exact shared function both the manual and automatic paths
 * converge on for revisions.
 *
 * Run: node tests/continuous-update-revisions-idempotency.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode } = require('./helpers/load-code.js');

const DRAFTS_FIXTURE = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86178-drafts-folder.json'), 'utf8'));

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

// ---- same minimal fake DocumentApp surface as tests/revision-drafts-folder.test.js

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

const AD_HOC_URL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/';

// ============ 1. collectorUpdate_() -- both real drafts, repeated calls ====

console.log('collectorUpdate_() (the function continuousUpdate() itself calls) -- repeated automatic-style invocations stay deterministic');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_URL } });

  // Isolate this test to revisions -- checkRSSFeed_ is an entirely
  // separate, unrelated concern (email discussion), stubbed to a no-op so
  // this test cannot be affected by, or accidentally exercise, that logic.
  sandbox.checkRSSFeed_ = () => {};

  const parentTable = makeTdocTable('S4aP260071', '5.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  sandbox.UrlFetchApp = {
    fetch: (url) => {
      if (url === AD_HOC_URL) return { getResponseCode: () => 200, getContentText: () => DRAFTS_FIXTURE.html };
      throw new Error('Unexpected fetch URL in test: ' + url);
    }
  };

  // Simulates an automatic trigger firing repeatedly over the course of
  // the live meeting -- e.g. every few minutes.
  sandbox.collectorUpdate_();
  const afterRun1 = parentTable.getCell(6, 1).getText();
  const linksAfterRun1 = JSON.stringify(parentTable.getCell(6, 1)._state.links);
  check('run 1 (automatic path): both real drafts discovered, exact text',
    afterRun1, 'S4aP260071_BBC.docx\nS4aP260071_QCOM.docx');

  sandbox.collectorUpdate_();
  const afterRun2 = parentTable.getCell(6, 1).getText();
  const linksAfterRun2 = JSON.stringify(parentTable.getCell(6, 1)._state.links);
  check('run 2 (automatic path): text is byte-identical to run 1 (no duplication from repeated automatic firing)',
    afterRun2, afterRun1);
  check('run 2: hyperlink ranges identical to run 1', linksAfterRun2, linksAfterRun1);

  sandbox.collectorUpdate_();
  const afterRun3 = parentTable.getCell(6, 1).getText();
  check('run 3 (automatic path): still byte-identical (idempotent across 3+ repeated automatic firings)',
    afterRun3, afterRun1);
  check('run 3: still exactly 2 logical lines, no growth', afterRun3.split('\n').length, 2);
}

// ==== 2. an already-populated cell (from a PRIOR automatic run) stays clean

console.log('collectorUpdate_() -- an already-populated Revisions cell (from a prior automatic run) is not duplicated by the next one');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_URL } });
  sandbox.checkRSSFeed_ = () => {};

  // The cell already shows the correct, clean result of a PRIOR automatic
  // run -- exactly the steady-state condition a real periodic trigger
  // would be in for most of its firings.
  const parentTable = makeTdocTable('S4aP260071', '5.1', 'S4aP260071_BBC.docx\nS4aP260071_QCOM.docx');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  sandbox.UrlFetchApp = {
    fetch: (url) => {
      if (url === AD_HOC_URL) return { getResponseCode: () => 200, getContentText: () => DRAFTS_FIXTURE.html };
      throw new Error('Unexpected fetch URL in test: ' + url);
    }
  };

  sandbox.collectorUpdate_();
  check('the already-clean cell remains exactly the same after another automatic run (no duplication)',
    parentTable.getCell(6, 1).getText(), 'S4aP260071_BBC.docx\nS4aP260071_QCOM.docx');
}

// ===================== 3. REVIS_ store does not grow across repeated runs ==

console.log('REVIS_<tdoc> store does not grow indefinitely across repeated automatic updates');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', REVISIONS_URL: AD_HOC_URL } });
  sandbox.checkRSSFeed_ = () => {};

  const parentTable = makeTdocTable('S4aP260071', '5.1');
  const body = makeFakeBody([parentTable]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  sandbox.UrlFetchApp = {
    fetch: (url) => {
      if (url === AD_HOC_URL) return { getResponseCode: () => 200, getContentText: () => DRAFTS_FIXTURE.html };
      throw new Error('Unexpected fetch URL in test: ' + url);
    }
  };

  for (let i = 0; i < 5; i++) sandbox.collectorUpdate_();

  const store = JSON.parse(docProps._store['REVIS_S4aP260071']);
  check('after 5 repeated automatic runs, the store still has exactly 2 entries (no unbounded growth)',
    Object.keys(store).length, 2);
  check('the rendered cell is still exactly the 2 real drafts after 5 runs',
    parentTable.getCell(6, 1).getText(), 'S4aP260071_BBC.docx\nS4aP260071_QCOM.docx');
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
