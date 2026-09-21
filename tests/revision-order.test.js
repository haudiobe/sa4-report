/**
 * Tests for the revision placement logic added in Code.js 2.5.0.
 *
 * Code.js targets Google Apps Script, so it is loaded into a VM sandbox with
 * minimal stubs for the Apps Script globals. The functions under test are pure
 * JavaScript, so they run unchanged.
 *
 * Run: node tests/revision-order.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---------------------------------------------------------------- load Code.js

const sandbox = {
  Logger: { log: () => {} },
  PropertiesService: {
    getDocumentProperties: () => ({
      getProperty: () => null,
      setProperty: () => {},
      deleteProperty: () => {},
      getKeys: () => []
    }),
    getScriptProperties: () => ({ getProperty: () => null, setProperty: () => {} })
  },
  DocumentApp: {
    getActiveDocument: () => { throw new Error('not used in these tests'); },
    ParagraphHeading: { NORMAL: 'NORMAL', TITLE: 'TITLE', HEADING1: 'H1', HEADING2: 'H2', HEADING3: 'H3' },
    ElementType: { PARAGRAPH: 'PARAGRAPH', TABLE: 'TABLE', LIST_ITEM: 'LIST_ITEM', INLINE_IMAGE: 'IMG' }
  },
  Session: { getScriptTimeZone: () => 'Europe/Berlin' },
  Utilities: {},
  UrlFetchApp: {},
  DriveApp: {},
  SpreadsheetApp: {},
  HtmlService: {},
  ScriptApp: {},
  MimeType: {},
  XmlService: {},
  console
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8'), sandbox);

const {
  orderTdocsByRevision_,
  orderRevisionChains_,
  getRevisedTo_,
  tdocNumberOf_,
  setDispositionRevisedTo_
} = sandbox;

// ------------------------------------------------------------------- harnesses

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

// A TDOC-list row shaped like downloadAndGroupTdocs_ produces:
// column 0 = TDoc, column 1 = "Revised to".
function row(tdoc, revisedTo, revisedToCol = 1) {
  return { row: [tdoc, revisedTo || ''], tdocCol: 0, revisedToCol };
}
const numbers = tdocs => tdocs.map(t => tdocNumberOf_(t));

// A minimal stand-in for a Google Docs table of label/value rows.
function fakeTable(rows) {
  return {
    rows,
    getNumRows: () => rows.length,
    getRow: r => ({
      getNumCells: () => 2,
      getCell: c => ({
        getText: () => rows[r][c],
        setText: v => { rows[r][c] = v; }
      })
    })
  };
}

// ----------------------------------------------------------- getRevisedTo_

console.log('getRevisedTo_');
check('reads a TDOC number', getRevisedTo_(row('S4-261480', 'S4-261599')), 'S4-261599');
check('tolerates surrounding text', getRevisedTo_(row('S4-261480', 'revised to S4-261599')), 'S4-261599');
check('lower case is normalized', getRevisedTo_(row('S4-261480', 's4-261599')), 'S4-261599');
check('empty cell -> ""', getRevisedTo_(row('S4-261480', '')), '');
check('absent column -> ""', getRevisedTo_(row('S4-261480', 'S4-261599', -1)), '');
check('non-TDOC text -> ""', getRevisedTo_(row('S4-261480', 'n/a')), '');

// ------------------------------------------------------- orderTdocsByRevision_

console.log('orderTdocsByRevision_');

check('revision follows its parent even when listed first',
  numbers(orderTdocsByRevision_([
    row('S4-261599', ''),            // the revision, listed first
    row('S4-261480', 'S4-261599')    // the document it revises
  ])),
  ['S4-261480', 'S4-261599']);

check('chain A -> B -> C is emitted in order',
  numbers(orderTdocsByRevision_([
    row('S4-000003', ''),
    row('S4-000002', 'S4-000003'),
    row('S4-000001', 'S4-000002')
  ])),
  ['S4-000001', 'S4-000002', 'S4-000003']);

check('unrelated documents keep their relative order',
  numbers(orderTdocsByRevision_([
    row('S4-000010', ''),
    row('S4-000020', ''),
    row('S4-000030', '')
  ])),
  ['S4-000010', 'S4-000020', 'S4-000030']);

check('parent stays put when its revision is in another agenda item',
  numbers(orderTdocsByRevision_([
    row('S4-000010', 'S4-999999'),   // target not in this group
    row('S4-000020', '')
  ])),
  ['S4-000010', 'S4-000020']);

check('two independent chains interleave correctly',
  numbers(orderTdocsByRevision_([
    row('S4-000002', ''),            // revision of 1
    row('S4-000004', ''),            // revision of 3
    row('S4-000001', 'S4-000002'),
    row('S4-000003', 'S4-000004')
  ])),
  ['S4-000001', 'S4-000002', 'S4-000003', 'S4-000004']);

check('a cycle drops nothing',
  numbers(orderTdocsByRevision_([
    row('S4-000001', 'S4-000002'),
    row('S4-000002', 'S4-000001')
  ])).sort(),
  ['S4-000001', 'S4-000002']);

check('missing "Revised to" column leaves the order untouched',
  numbers(orderTdocsByRevision_([
    row('S4-000002', 'S4-000001', -1),
    row('S4-000001', '', -1)
  ])),
  ['S4-000002', 'S4-000001']);

check('no document is ever lost',
  orderTdocsByRevision_([
    row('S4-000003', ''),
    row('S4-000002', 'S4-000003'),
    row('S4-000001', 'S4-000002'),
    row('S4-000009', '')
  ]).length,
  4);

// ------------------------------------------------------- orderRevisionChains_

console.log('orderRevisionChains_');
check('chain roots come first', orderRevisionChains_({ 'S4-000002': 'S4-000003', 'S4-000001': 'S4-000002' }),
  ['S4-000001', 'S4-000002']);
check('single relationship', orderRevisionChains_({ 'S4-000001': 'S4-000002' }), ['S4-000001']);
check('empty map', orderRevisionChains_({}), []);
check('cycle is still returned', orderRevisionChains_({ 'S4-000001': 'S4-000002', 'S4-000002': 'S4-000001' }).sort(),
  ['S4-000001', 'S4-000002']);

// ---------------------------------------------------- setDispositionRevisedTo_

console.log('setDispositionRevisedTo_');

let t = fakeTable([['TDoc', 'S4-261480'], ['Disposition', '']]);
check('empty Disposition is filled', setDispositionRevisedTo_(t, 'S4-261599'), true);
check('  -> cell text', t.rows[1][1], 'Revised to S4-261599');

t = fakeTable([['Disposition', 'Noted']]);
check('hand-written text is preserved', setDispositionRevisedTo_(t, 'S4-261599'), true);
check('  -> marker appended', t.rows[0][1], 'Noted — Revised to S4-261599');

t = fakeTable([['Disposition', 'Revised to S4-261599']]);
check('already correct -> no change', setDispositionRevisedTo_(t, 'S4-261599'), false);
check('  -> cell untouched', t.rows[0][1], 'Revised to S4-261599');

t = fakeTable([['Disposition', 'revised to something else']]);
check('existing revision mention is not doubled', setDispositionRevisedTo_(t, 'S4-261599'), false);
check('  -> cell untouched', t.rows[0][1], 'revised to something else');

t = fakeTable([['Disposition:', '']]);
check('"Disposition:" label variant is found', setDispositionRevisedTo_(t, 'S4-261599'), true);

t = fakeTable([['TDoc', 'S4-261480'], ['Status', 'Available']]);
check('no Disposition row -> false', setDispositionRevisedTo_(t, 'S4-261599'), false);

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll tests passed.' : `\n${failures} test(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);