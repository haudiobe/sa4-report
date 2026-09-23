/**
 * SA4-PROD-006 — Reviewer API abstract-fetching gates migrated from a
 * hardcoded main-meeting-only /^S4-\d{6}$/ regex to the centralized,
 * registered-family identifier check (parseExactSA4DocumentId_() /
 * SA4_TDOC_FAMILIES), so ad-hoc TDocs (e.g. "S4aP260098") can reach
 * fetchAndAddAbstract_() the same way a main-meeting TDoc always could.
 *
 * Two call sites were migrated:
 *   - createTDocTableFromData_() (Code.js) -- gates the abstract fetch
 *     performed while a TDoc table is first built.
 *   - addAbstractsForTables_() (Code.js) -- gates the batch "fill in any
 *     missing abstract" pass over every existing TDoc table.
 *
 * fetchAndAddAbstract_() itself makes a real UrlFetchApp.fetch() call, so
 * this suite does not execute it -- instead, per this project's established
 * pattern (see tests/README.md and the header comments in
 * tests/canonical-agenda-projection.test.js /
 * tests/meeting-86178-production.test.js), it monkey-patches
 * fetchAndAddAbstract_() itself to a recording stub and executes the REAL,
 * unmodified createTDocTableFromData_()/addAbstractsForTables_() against
 * small hand-built fake Body/Table objects -- enough to prove the actual
 * gating decision and the actual argument passed, without needing a fake
 * UrlFetchApp or DocumentApp beyond what these two functions need.
 *
 * Run: node tests/reviewer-api-tdoc-family.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

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

function extractFunctionBody(source, fnName) {
  const startMatch = source.match(new RegExp('^function ' + fnName + '\\(', 'm'));
  if (!startMatch) return null;
  const startIndex = startMatch.index;
  const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
  nextFnRe.lastIndex = startIndex + startMatch[0].length;
  const next = nextFnRe.exec(source);
  const endIndex = next ? next.index : source.length;
  return source.slice(startIndex, endIndex);
}

function stripComments(src) {
  return src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
}

// Minimal fake Table/Row/Body -- just enough for createTDocTableFromData_()
// and addAbstractsForTables_() (appendTable/appendTableRow/appendTableCell,
// getRow/getCell/getNumRows/getNumCells/getCell(row,col), getTables()).
function makeFakeTableCell(text) {
  return { getText: () => text, editAsText: () => ({ setLinkUrl: () => {} }) };
}
function makeFakeTableRow(cells) {
  const _cells = cells.slice();
  return {
    getNumCells: () => _cells.length,
    getCell: (i) => makeFakeTableCell(_cells[i] === undefined ? '' : _cells[i]),
    appendTableCell: (t) => { _cells.push(t); return makeFakeTableCell(t); }
  };
}
function makeFakeTable(rows) {
  const _rows = (rows || []).map(r => makeFakeTableRow(r));
  return {
    getType: () => 'TABLE',
    getNumRows: () => _rows.length,
    getRow: (i) => _rows[i],
    getCell: (r, c) => _rows[r].getCell(c),
    appendTableRow: () => { const r = makeFakeTableRow([]); _rows.push(r); return r; }
  };
}
function makeFakeBody(tables) {
  return { getTables: () => tables };
}

// ================================================== 1. source-structure ===

console.log('source-structure: both Reviewer gates migrated to parseExactSA4DocumentId_(), no new regex introduced');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  const createBody = extractFunctionBody(source, 'createTDocTableFromData_');
  if (!createBody) {
    failures++;
    console.log('  FAIL could not locate "function createTDocTableFromData_(" in Code.js');
  } else {
    check('createTDocTableFromData_() now calls parseExactSA4DocumentId_()',
      /\bparseExactSA4DocumentId_\s*\(/.test(createBody), true);
    check('createTDocTableFromData_() no longer contains the old hardcoded /^S4-\\d{6}$/ gate',
      /\/\^S4-\\d\{6\}\$\//.test(stripComments(createBody)), false);
    check('createTDocTableFromData_() passes the PARSED canonical id (parsedTdoc.raw) to fetchAndAddAbstract_(), not the raw cell text',
      /fetchAndAddAbstract_\(table,\s*parsedTdoc\.raw\)/.test(createBody), true);
  }

  const addBody = extractFunctionBody(source, 'addAbstractsForTables_');
  if (!addBody) {
    failures++;
    console.log('  FAIL could not locate "function addAbstractsForTables_(" in Code.js');
  } else {
    check('addAbstractsForTables_() now calls parseExactSA4DocumentId_()',
      /\bparseExactSA4DocumentId_\s*\(/.test(addBody), true);
    check('addAbstractsForTables_() no longer contains the old hardcoded /^S4-\\d{6}$/ gate',
      /\/\^S4-\\d\{6\}\$\//.test(stripComments(addBody)), false);
    // ADDON-004: this call site now also threads an optional trailing
    // `context` argument through to fetchAndAddAbstract_() -- tolerate it
    // rather than requiring an exact 2-argument call.
    check('addAbstractsForTables_() passes the PARSED canonical id (parsedTdoc.raw) to fetchAndAddAbstract_()',
      /fetchAndAddAbstract_\(table,\s*parsedTdoc\.raw[,)]/.test(addBody), true);
  }

  // Every OTHER /^S4-\d{6}$/-shaped regex in Code.js is deliberately
  // unrelated to Reviewer (TDOC_ID_REGEX default, the legacy web-import
  // dispatcher, etc) and untouched by this task -- not asserted here.

  const testReviewerBody = extractFunctionBody(source, 'testReviewerApi');
  if (!testReviewerBody) {
    failures++;
    console.log('  FAIL could not locate "function testReviewerApi(" in Code.js');
  } else {
    check('testReviewerApi() now probes cfg.AGENDA_TDOC via parseExactSA4DocumentId_() when valid',
      /parseExactSA4DocumentId_\(cfg\.AGENDA_TDOC\)/.test(testReviewerBody), true);
    check('testReviewerApi() still falls back to the original "S4-260001" probe when no valid AGENDA_TDOC is configured',
      /'S4-260001'/.test(testReviewerBody), true);
  }

  const testAllBody = extractFunctionBody(source, 'testAllConnections');
  if (!testAllBody) {
    failures++;
    console.log('  FAIL could not locate "function testAllConnections(" in Code.js');
  } else {
    check('testAllConnections() Reviewer API test now probes cfg.AGENDA_TDOC via parseExactSA4DocumentId_() when valid',
      /parseExactSA4DocumentId_\(cfg\.AGENDA_TDOC\)/.test(testAllBody), true);
  }
}

// ============================== 2. identifier family acceptance/rejection =

console.log('parseExactSA4DocumentId_() -- exactly the six registered SA4_TDOC_FAMILIES, the same abstraction both gates now use');

{
  const { sandbox } = loadCode();
  const accepted = ['S4-260123', 'S4aA260090', 'S4aP260098', 'S4aV200545', 'S4aI240064', 'A4aR260097'];
  accepted.forEach(id => {
    check(`"${id}" is accepted (isValid)`, sandbox.parseExactSA4DocumentId_(id).isValid, true);
  });

  const rejected = ['S4-26012', 'S4-2601234', 'S4aX260090', 'A4-260097', 'S4a260090', 'not-a-tdoc', '', 'S4aP26009'];
  rejected.forEach(id => {
    check(`"${id}" is rejected (not isValid)`, sandbox.parseExactSA4DocumentId_(id).isValid, false);
  });
}

// ==== 3. createTDocTableFromData_() actually reaches fetchAndAddAbstract_() =

console.log('createTDocTableFromData_() -- real execution, fetchAndAddAbstract_() stubbed, proving S4aP260098 (and every registered family) is reached');

{
  const { sandbox } = loadCode();
  const fakeBody = { appendTable: () => makeFakeTable([]) };

  const cases = [
    { id: 'S4-260123', shouldReach: true },
    { id: 'S4aA260090', shouldReach: true },
    { id: 'S4aP260098', shouldReach: true },
    { id: 'S4aV200545', shouldReach: true },
    { id: 'S4aI240064', shouldReach: true },
    { id: 'A4aR260097', shouldReach: true },
    { id: 'not-a-tdoc', shouldReach: false },
    { id: 'S4aX260090', shouldReach: false }
  ];

  cases.forEach(({ id, shouldReach }) => {
    let calledWith = null;
    sandbox.fetchAndAddAbstract_ = (table, tdocNumber) => { calledWith = tdocNumber; };
    const data = [['TDoc', id], ['Title', 'Some title']];
    sandbox.createTDocTableFromData_(fakeBody, data, null, 1);
    if (shouldReach) {
      check(`createTDocTableFromData_("${id}") reaches fetchAndAddAbstract_() with the canonical id`, calledWith, id);
    } else {
      check(`createTDocTableFromData_("${id}") does NOT reach fetchAndAddAbstract_()`, calledWith, null);
    }
  });

  // SKIP_ABSTRACTS_DURING_TABLE_BUILD must still be honored, unchanged,
  // even for a now-newly-accepted ad-hoc id.
  {
    const { sandbox: sandbox2, docProps } = loadCode({ documentProperties: { SKIP_ABSTRACTS_DURING_TABLE_BUILD: 'true' } });
    let called = false;
    sandbox2.fetchAndAddAbstract_ = () => { called = true; };
    sandbox2.createTDocTableFromData_(fakeBody, [['TDoc', 'S4aP260098'], ['Title', 'x']], null, 1);
    check('SKIP_ABSTRACTS_DURING_TABLE_BUILD=true still suppresses the fetch, even for a valid ad-hoc id', called, false);
  }
}

// ======= 4. addAbstractsForTables_() actually reaches fetchAndAddAbstract_() =

console.log('addAbstractsForTables_() -- real execution, fetchAndAddAbstract_() stubbed, proving S4aP260098 is reached in the batch pass too');

{
  const { sandbox } = loadCode();

  function tdocTable(tdocId, hasAbstract) {
    const rows = [['TDoc', tdocId], ['Title', 'Some title']];
    if (hasAbstract) rows.push(['Abstract', 'already here']);
    return makeFakeTable(rows);
  }

  const s4apTable = tdocTable('S4aP260098', false);
  const s4Table = tdocTable('S4-260123', false);
  const invalidTable = tdocTable('not-a-tdoc', false);
  const alreadyHasAbstractTable = tdocTable('S4aA260090', true);
  const notATdocTableAtAll = makeFakeTable([['Something', 'else']]);

  const fakeBody = makeFakeBody([s4apTable, s4Table, invalidTable, alreadyHasAbstractTable, notATdocTableAtAll]);

  const calledFor = [];
  sandbox.fetchAndAddAbstract_ = (table, tdocNumber) => { calledFor.push(tdocNumber); };

  // PERF-006B: addAbstractsForTables_() now returns a {candidatesProcessed,
  // requestsMade, cacheSkips, rowsInserted} breakdown, not a bare count --
  // see its own header comment (Code.js) for why "candidates processed"
  // is not the same thing as "Reviewer requests actually made".
  const result = sandbox.addAbstractsForTables_(fakeBody);

  check('addAbstractsForTables_() reaches fetchAndAddAbstract_() for S4aP260098', calledFor.indexOf('S4aP260098') !== -1, true);
  check('addAbstractsForTables_() reaches fetchAndAddAbstract_() for the main-meeting S4-260123 too', calledFor.indexOf('S4-260123') !== -1, true);
  check('addAbstractsForTables_() skips the invalid/unregistered id', calledFor.indexOf('not-a-tdoc') === -1, true);
  check('addAbstractsForTables_() skips a table that already has an Abstract row', calledFor.indexOf('S4aA260090') === -1, true);
  check('addAbstractsForTables_() only counted the tables it actually processed (candidatesProcessed)', result.candidatesProcessed, 2);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
