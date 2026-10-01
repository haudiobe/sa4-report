/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/bugfix-legacy-002.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-bugfix-legacy-002.test.js
 */

/**
 * BUGFIX-LEGACY-002 -- duplicate TDoc detail tables from the legacy
 * "Build Initial Report" pathway.
 *
 * Root cause (see Code.js's own BUGFIX-LEGACY-002 comments at
 * processWebDownloadedSheet_()/continuousUpdateCore_()): the menu item
 * "📝 Legacy: Build Initial Report" (buildInitialReport() ->
 * downloadAndProcessFromWeb() -> processWebDownloadedSheet_()) does not
 * clear the document first and, unlike continuousUpdateCore_(), never
 * checked whether a table for a given TDoc number already existed before
 * inserting a new one. Running it more than once therefore inserted a
 * complete duplicate table (and a duplicate "Registered Documents"
 * heading) for every TDoc it processed -- reproduced here directly,
 * without any live document or network access.
 *
 * This suite first demonstrates the bug is real and reproducible (see the
 * "BEFORE the fix" section below, kept for the record), then verifies the
 * fix restores the required idempotency invariant: the same TDoc
 * processed twice must still result in exactly one physical table.
 *
 * Run: node tests/bugfix-legacy-002.test.js
 */

const { loadCode } = require('./helpers/legacy-load-code.js');

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

// ------------------------------------------------------------ fakes

function makeFakeDocumentBody(sandbox) {
  const PARAGRAPH = sandbox.DocumentApp.ElementType.PARAGRAPH;
  const TABLE = sandbox.DocumentApp.ElementType.TABLE;
  const NORMAL = sandbox.DocumentApp.ParagraphHeading.NORMAL;
  const children = [];

  function makeParagraph(text, heading) {
    const p = {
      _text: text, _heading: heading || NORMAL,
      getType: () => PARAGRAPH, asParagraph: () => p,
      getText: () => p._text,
      getHeading: () => p._heading,
      setHeading: (h) => { p._heading = h; return p; },
      setGlyphType: () => p
    };
    return p;
  }
  function makeTable(rowsData) {
    const rows = [];
    function makeRow(cellTexts) {
      const cells = cellTexts.map((t) => ({ _t: String(t), getText: () => String(t), editAsText: () => ({ getText: () => String(t), setLinkUrl() {}, setBold() {}, setItalic() {}, setUnderline() {}, setForegroundColor() {} }) }));
      return {
        getNumCells: () => cells.length, getCell: (i) => cells[i],
        appendTableCell: (t) => { const c = { _t: String(t), getText: () => String(t), editAsText: () => ({ getText: () => String(t), setLinkUrl() {}, setBold() {}, setItalic() {}, setUnderline() {}, setForegroundColor() {} }) }; cells.push(c); return c; }
      };
    }
    (rowsData || []).forEach((r) => rows.push(makeRow(r)));
    const t = {
      getType: () => TABLE, asTable: () => t,
      getNumRows: () => rows.length, getRow: (i) => rows[i],
      getCell: (r, c) => rows[r].getCell(c),
      appendTableRow: () => { const r = makeRow([]); rows.push(r); return r; },
      removeRow: (i) => { rows.splice(i, 1); }
    };
    return t;
  }

  const body = {
    _children: children,
    clear: () => { children.length = 0; return body; },
    getNumChildren: () => children.length,
    getChild: (i) => children[i],
    getChildIndex: (c) => children.indexOf(c),
    getTables: () => children.filter((c) => c.getType() === TABLE),
    getParagraphs: () => children.filter((c) => c.getType() === PARAGRAPH),
    appendParagraph: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendListItem: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendTable: () => { const t = makeTable([]); children.push(t); return t; },
    insertParagraph: (idx, text) => { const p = makeParagraph(text); children.splice(idx, 0, p); return p; },
    insertTable: (idx, data) => { const t = makeTable(data); children.splice(idx, 0, t); return t; }
  };
  return body;
}

// A minimal fake of the Sheet object processWebDownloadedSheet_() reads
// via sheet.getDataRange().getValues()/getRichTextValues() -- the exact
// header row this codebase's own downloadAndGroupTdocs_()/
// processWebDownloadedSheet_() both expect.
function makeFakeSheet(rows) {
  const header = ['TDoc', 'Title', 'Source', 'Contact', 'Agenda item', 'Agenda Topic', 'TDoc Status', 'Type', 'For', 'Revised to'];
  const data = [header].concat(rows);
  return {
    getDataRange: () => ({
      getValues: () => data,
      getRichTextValues: () => data.map(() => header.map(() => null))
    })
  };
}

function tdocRow(tdoc, title, source, agendaItem, status) {
  return [tdoc, title, source, '', agendaItem, '', status || '', '', '', ''];
}

// The exact live-observed duplicate: same TDoc/Title/Source/Agenda Item.
const S4AP260099_ROW = tdocRow('S4aP260099', 'Draft TR 26.870 FS_6G_MED v0.6.2', 'VODAFONE Group Plc', '5.4', 'Revised');

function countTdocTables(sandbox, body, tdocNumber) {
  return body.getTables().filter((t) => sandbox.isTDocTable_(t) && sandbox.safeCellText_(t, 0, 1).trim() === tdocNumber).length;
}

// ------------------------------------------------------------ reproduction + fix verification

console.log('processWebDownloadedSheet_() -- BEFORE the fix, this reproduced the live bug (kept here for the record; see the "idempotency" checks below for the actual pass/fail signal against current code)');

{
  const { sandbox } = loadCode({ documentProperties: { AGENDA_ITEM_PREFIX: '5.', REPORT_SUFFIX: 'Audio' } });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body });
  const sheet = makeFakeSheet([S4AP260099_ROW]);

  // Simulates the user running "📝 Legacy: Build Initial Report" (which
  // calls this exact function) a first, then a second time, against the
  // SAME already-built document -- exactly the real workflow that
  // produced three live copies of S4aP260099.
  sandbox.processWebDownloadedSheet_(sheet);
  const afterFirstRun = countTdocTables(sandbox, body, 'S4aP260099');
  sandbox.processWebDownloadedSheet_(sheet);
  const afterSecondRun = countTdocTables(sandbox, body, 'S4aP260099');

  check('BUGFIX-LEGACY-002 root cause reproduced and fixed: one run creates exactly one table', afterFirstRun, 1);
  check('BUGFIX-LEGACY-002 root cause reproduced and fixed: a second run on the same document does NOT duplicate it (idempotency invariant)', afterSecondRun, 1);
}

console.log('processWebDownloadedSheet_() -- idempotency: repeated processing of the same base TDoc');

{
  const { sandbox } = loadCode({ documentProperties: { AGENDA_ITEM_PREFIX: '5.', REPORT_SUFFIX: 'Audio' } });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body });
  const sheet = makeFakeSheet([S4AP260099_ROW]);

  sandbox.processWebDownloadedSheet_(sheet);
  sandbox.processWebDownloadedSheet_(sheet);
  sandbox.processWebDownloadedSheet_(sheet); // a THIRD run, matching the exact live symptom (3 copies observed)
  check('three separate runs still leave exactly ONE table for the TDoc (never three)', countTdocTables(sandbox, body, 'S4aP260099'), 1);
}

console.log('processWebDownloadedSheet_() -- a genuinely different TDoc still creates normally on a later run');

{
  const { sandbox } = loadCode({ documentProperties: { AGENDA_ITEM_PREFIX: '5.', REPORT_SUFFIX: 'Audio' } });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body });

  sandbox.processWebDownloadedSheet_(makeFakeSheet([S4AP260099_ROW]));
  sandbox.processWebDownloadedSheet_(makeFakeSheet([S4AP260099_ROW, tdocRow('S4aP260074', 'Another contribution', 'Nokia', '5.4', '')]));

  check('the already-processed TDoc is still exactly one table', countTdocTables(sandbox, body, 'S4aP260099'), 1);
  check('the genuinely new TDoc introduced on the second run IS created', countTdocTables(sandbox, body, 'S4aP260074'), 1);
}

console.log('processWebDownloadedSheet_() -- a duplicate row WITHIN the same source list/run is also not inserted twice');

{
  // Covers the other mechanism investigated: the source TDoc list itself
  // (whatever the reason) containing the same TDoc number more than once
  // in a SINGLE fetch/run.
  const { sandbox } = loadCode({ documentProperties: { AGENDA_ITEM_PREFIX: '5.', REPORT_SUFFIX: 'Audio' } });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body });
  const sheet = makeFakeSheet([S4AP260099_ROW, S4AP260099_ROW]); // the SAME TDoc twice in one fetch

  sandbox.processWebDownloadedSheet_(sheet);
  check('a within-run duplicate source row still results in exactly one table', countTdocTables(sandbox, body, 'S4aP260099'), 1);
}

// ------------------------------------------------------------ continuousUpdateCore_ hardening

console.log('continuousUpdateCore_() -- the same within-run duplicate-source-row protection is now also present in the modern pathway');

{
  const { sandbox } = loadCode({ documentProperties: {
    MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic AH', REPORT_SUFFIX: 'Audio', AGENDA_TDOC: 'S4aA000001',
    FTP_BASE: 'https://example.invalid/Docs/', TDOC_LIST_URL: 'https://example.invalid/list.xlsx'
  } });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  sandbox.downloadAndGroupTdocs_ = () => ({
    // Two DIFFERENT agenda-item groups both citing the SAME TDoc number,
    // the exact shape a duplicated source row could produce.
    '1': { tdocs: [{ row: ['S4aP260099', 'Draft TR', 'VODAFONE', ''], tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 3, statusCol: -1, typeCol: -1, forCol: -1, revisedToCol: -1, richTextRow: null }] },
    '1.1': { tdocs: [{ row: ['S4aP260099', 'Draft TR', 'VODAFONE', ''], tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 3, statusCol: -1, typeCol: -1, forCol: -1, revisedToCol: -1, richTextRow: null }] }
  });
  sandbox.styleStatusCell_ = () => {};
  sandbox.removeRowHeightAndSpacing = () => {};
  sandbox.getFetchAbstractsSetting_ = () => false;
  sandbox.collectorUpdate_ = () => {};
  const logs = [];
  sandbox.Logger = { log: (m) => logs.push(String(m)) };
  // This legacy continuousUpdateCore_() (unlike the central add-on's
  // ADDON-004-enhanced version) does not return a {success,error} result --
  // it swallows internal failures via its own catch/Logger.log, matching
  // BUGFIX-LEGACY-001's own established test pattern.
  let thrown = null;
  try { sandbox.continuousUpdateCore_(); } catch (e) { thrown = e; }
  check('continuousUpdateCore_ completes without throwing', thrown, null);
  check('no internal ERROR was logged', logs.some((l) => /^ERROR:/.test(l)), false);
  check('a TDoc appearing twice in the SAME run\'s downloaded groups still results in exactly one table', countTdocTables(sandbox, body, 'S4aP260099'), 1);
}

// ------------------------------------------------------------ main/ad-hoc regression

console.log('regression: main-meeting behavior is unaffected by the idempotency guard');

{
  const { sandbox } = loadCode({ documentProperties: { AGENDA_ITEM_PREFIX: '5.', REPORT_SUFFIX: 'Audio' } });
  const body = makeFakeDocumentBody(sandbox);
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body });
  sandbox.processWebDownloadedSheet_(makeFakeSheet([tdocRow('S4-260123', 'Main meeting contribution', 'Ericsson', '5.1', '')]));
  check('a normal, previously-unseen main-series TDoc is still created on the first run', countTdocTables(sandbox, body, 'S4-260123'), 1);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
