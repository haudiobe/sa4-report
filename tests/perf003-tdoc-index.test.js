/**
 * PERF-003 (Part A) — eliminate repeated TDoc table scans.
 *
 * PERF-002's controlled production measurement found updateTdocStatus_()
 * alone calling body.getTables() 32 times for a 32-TDoc, zero-change run
 * (one fresh full-document table scan per EXISTING TDoc). This suite
 * proves the fix -- a per-run canonical-identity index
 * (buildTdocTableIndex_(), Map<canonicalTdocId, Table>, built from a
 * SINGLE getTablesCounted_() call and keyed via the existing central SA4
 * registry, parseExactSA4DocumentId_ -- never a second/parallel TDoc-regex
 * implementation) -- is correct and behavior-preserving:
 *
 *   - updateTdocStatus_()/findTdocTable_() produce the EXACT same result
 *     with the index as without it (index is a pure performance path, not
 *     a semantic change);
 *   - omitting the index preserves the ORIGINAL scan-every-call behavior
 *     exactly, for full backward compatibility;
 *   - the index is correctly extended when insertNewTdoc_() inserts a new
 *     table (so a LATER new TDoc in the same run that revises this one
 *     still finds it, without a full rebuild);
 *   - the index is correctly updated (not left dangling) when
 *     rearrangeRevisionTables_() actually moves a table via
 *     moveTableAfter_().
 *
 * Run: node tests/perf003-tdoc-index.test.js
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

// ---- minimal fake DocumentApp surface (same pattern as
// tests/meeting-86178-production.test.js / tests/revision-drafts-folder.test.js)

function makeFakeCell(initialText) {
  const state = { text: String(initialText === undefined || initialText === null ? '' : initialText) };
  return {
    getText: () => state.text,
    setText: (t) => { state.text = String(t || ''); },
    editAsText: () => ({ getText: () => state.text, setLinkUrl: () => {}, setForegroundColor: () => {}, setBold: () => {} }),
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

function makeFakeTable(rows, id) {
  let _rows = rows.map(r => makeFakeRow(r));
  let _parent = null;
  const self = {
    getType: () => 'TABLE',
    getNumRows: () => _rows.length,
    getRow: (i) => _rows[i],
    getCell: (r, c) => _rows[r].getCell(c),
    appendTableRow: () => { const row = makeFakeRow([]); _rows.push(row); return row; },
    removeRow: (i) => { _rows.splice(i, 1); },
    copy: () => makeFakeTable(_rows.map(r => r._cells.map(c => c.getText())), (id || 'table') + '-copy'),
    removeFromParent: () => { if (_parent) _parent._remove(self); },
    _setParent: (p) => { _parent = p; },
    _id: id,
    _rows
  };
  return self;
}

function makeFakeParagraph(text, heading) {
  const self = {
    _text: text,
    _heading: heading || 'NORMAL',
    getType: () => 'PARAGRAPH',
    asParagraph: () => self,
    getText: () => self._text,
    getHeading: () => self._heading,
    setHeading: (h) => { self._heading = h; return self; }
  };
  return self;
}

function makeFakeBody(children) {
  const _children = children.slice();
  const body = {
    _children,
    getNumChildren: () => _children.length,
    getChild: (i) => _children[i],
    getChildIndex: (child) => _children.indexOf(child),
    getTables: () => _children.filter(c => c.getType() === 'TABLE'),
    insertTable: (idx, dataOrTable) => {
      let t;
      if (dataOrTable && typeof dataOrTable.getType === 'function') {
        t = dataOrTable; // already a Table (e.g. from table.copy())
      } else {
        t = makeFakeTable(dataOrTable || [['', '']]);
      }
      t._setParent(body);
      if (idx === undefined || idx === null || idx > _children.length) idx = _children.length;
      _children.splice(idx, 0, t);
      return t;
    },
    appendParagraph: (text) => { const p = makeFakeParagraph(text); _children.push(p); return p; },
    _remove: (child) => { const i = _children.indexOf(child); if (i !== -1) _children.splice(i, 1); }
  };
  _children.forEach(c => { if (c.getType() === 'TABLE') c._setParent(body); });
  return body;
}

function makeTdocTable(tdoc, agendaItem, status) {
  return makeFakeTable([
    ['TDoc', tdoc],
    ['Title', 'Some title'],
    ['Source', 'Some source'],
    ['Contact', ''],
    ['Agenda Item', agendaItem || '5.1'],
    ['Type/For', ''],
    ['E-mail Discussion', ''],
    ['Revisions', ''],
    ['Minutes', ''],
    ['Disposition', ''],
    ['Status', status || 'Available']
  ], tdoc);
}

function makeTdocRow(tdoc, status) {
  return {
    row: ['', status || 'Available'],
    statusCol: 1,
    tdocCol: 0
  };
}

// ============================== 1. buildTdocTableIndex_() ==================

console.log('buildTdocTableIndex_() -- canonical-identity index from one scan, using the central SA4 registry');

{
  const { sandbox } = loadCode();
  const t1 = makeTdocTable('S4-260123');
  const t2 = makeTdocTable('S4aP260071');
  const notATdoc = makeFakeTable([['Key', 'Value'], ['SomeSetting', 'x']]);
  const malformed = makeFakeTable([['TDoc', 'not-a-real-tdoc-id'], ['Status', '']]);

  const index = sandbox.buildTdocTableIndex_([t1, notATdoc, t2, malformed]);
  check('index has exactly 2 entries (only real, valid SA4 TDoc tables)', index.size, 2);
  check('S4-260123 maps to its own table object', index.get('S4-260123'), t1);
  check('S4aP260071 maps to its own table object', index.get('S4aP260071'), t2);
  check('a non-TDoc table (Key/Value config table) is never indexed', index.has('SomeSetting'), false);
}

// ============================ 2. lookupTdocTableInIndex_() =================

console.log('lookupTdocTableInIndex_() -- O(1) canonical lookup, never throws, never guesses');

{
  const { sandbox } = loadCode();
  const t1 = makeTdocTable('S4aA260090');
  const index = sandbox.buildTdocTableIndex_([t1]);

  check('exact canonical id found', sandbox.lookupTdocTableInIndex_(index, 'S4aA260090'), t1);
  check('an id not in the index returns null', sandbox.lookupTdocTableInIndex_(index, 'S4aA999999'), null);
  check('an unrecognized/invalid id returns null, does not throw', sandbox.lookupTdocTableInIndex_(index, 'not-a-tdoc'), null);
  check('empty/null input returns null, does not throw', sandbox.lookupTdocTableInIndex_(index, null), null);
}

// ==================== 3. findTdocTable_() -- indexed vs fallback ===========

console.log('findTdocTable_() -- indexed lookup and full-scan fallback return the EXACT SAME result');

{
  const { sandbox } = loadCode();
  const t1 = makeTdocTable('S4-260123');
  const t2 = makeTdocTable('S4aP260071');
  const body = makeFakeBody([t1, t2]);
  const index = sandbox.buildTdocTableIndex_(body.getTables());

  check('WITH index: finds the right table, without calling body.getTables() again',
    sandbox.findTdocTable_(body, 'S4aP260071', index), t2);

  let scanCalled = false;
  const spyBody = Object.assign({}, body, { getTables: () => { scanCalled = true; return body.getTables(); } });
  sandbox.findTdocTable_(spyBody, 'S4aP260071', index);
  check('WITH index: body.getTables() is never called (no fresh scan)', scanCalled, false);

  check('WITHOUT index: falls back to the original scan-based behavior, same result',
    sandbox.findTdocTable_(body, 'S4aP260071'), t2);
  check('WITHOUT index: a genuinely absent TDoc still returns null', sandbox.findTdocTable_(body, 'S4-999999'), null);
  check('WITH index: a genuinely absent TDoc still returns null', sandbox.findTdocTable_(body, 'S4-999999', index), null);
}

// ================== 4. updateTdocStatus_() -- indexed vs fallback ==========

console.log('updateTdocStatus_() -- indexed path produces the EXACT same result as the original full scan');

{
  const { sandbox } = loadCode();
  const t1 = makeTdocTable('S4-260123', '5.1', 'Reserved');
  const body = makeFakeBody([t1]);
  const index = sandbox.buildTdocTableIndex_(body.getTables());

  const tdocData = makeTdocRow('S4-260123', 'Available');
  const result = sandbox.updateTdocStatus_(body, 'S4-260123', tdocData, index);
  check('WITH index: status update succeeds (Reserved -> Available is a valid transition)', result, true);
  check('WITH index: the cell text was actually updated', t1.getCell(10, 1).getText(), 'Available');
}

{
  // Same scenario, no index -- must produce the identical result.
  const { sandbox } = loadCode();
  const t1 = makeTdocTable('S4-260123', '5.1', 'Reserved');
  const body = makeFakeBody([t1]);
  const tdocData = makeTdocRow('S4-260123', 'Available');

  const result = sandbox.updateTdocStatus_(body, 'S4-260123', tdocData);
  check('WITHOUT index: identical result to the indexed path', result, true);
  check('WITHOUT index: the cell text was actually updated', t1.getCell(10, 1).getText(), 'Available');
}

{
  // No status change needed -- both paths must return false and leave the
  // cell untouched, matching the ORIGINAL function's exact semantics.
  const { sandbox } = loadCode();
  const t1 = makeTdocTable('S4-260123', '5.1', 'Agreed');
  const body = makeFakeBody([t1]);
  const index = sandbox.buildTdocTableIndex_(body.getTables());
  const tdocData = makeTdocRow('S4-260123', 'Noted'); // Agreed is neither reserved/available nor "revised"

  const resultIndexed = sandbox.updateTdocStatus_(body, 'S4-260123', tdocData, index);
  check('WITH index: status "Agreed" is not upgraded to "Noted" (matches original business rule)', resultIndexed, false);
  check('WITH index: cell text is untouched', t1.getCell(10, 1).getText(), 'Agreed');
}

{
  // Correctness safety net: if the indexed candidate's raw text does NOT
  // exactly match the requested tdocNumber (a deliberately constructed,
  // otherwise-unreachable edge case), updateTdocStatus_() must fall back
  // to the original full scan rather than silently using a bad candidate.
  const { sandbox } = loadCode();
  const t1 = makeTdocTable('S4-260123', '5.1', 'Reserved');
  const body = makeFakeBody([t1]);
  const badIndex = new Map();
  badIndex.set('S4-999999', t1); // deliberately wrong key, simulating a stale/mismatched index
  const tdocData = makeTdocRow('S4-260123', 'Available');

  const result = sandbox.updateTdocStatus_(body, 'S4-260123', tdocData, badIndex);
  check('a mismatched index entry safely falls back to the full scan and still finds the real table', result, true);
}

// ================ 5. insertNewTdoc_() keeps the index up to date ===========

console.log('insertNewTdoc_() -- extends the index with the newly inserted table, so a later lookup in the SAME run finds it');

{
  const { sandbox } = loadCode();
  const parent = makeTdocTable('S4-260100', '5.1');
  const body = makeFakeBody([
    makeFakeParagraph('5.1 Opening', 'HEADING2'),
    parent,
    makeFakeParagraph('5.2 Next section', 'HEADING2')
  ]);
  const index = sandbox.buildTdocTableIndex_(body.getTables());
  check('sanity: index does not yet contain the new TDoc', index.has('S4-260101'), false);

  const tdocData = {
    row: ['S4-260101', 'New title', 'Some source', '', '5.1', '', ''],
    tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 4, typeCol: -1, forCol: -1, statusCol: -1, revisedToCol: -1,
    agendaItem: '5.1',
    richTextRow: null
  };
  sandbox.insertNewTdoc_(body, tdocData, {}, index);

  check('the index now contains the newly inserted TDoc', index.has('S4-260101'), true);
  check('the indexed table is really in the document (findTdocTable_ agrees)',
    sandbox.findTdocTable_(body, 'S4-260101', index) === index.get('S4-260101'), true);

  // A SECOND new TDoc, in the SAME run, that revises the one just inserted
  // above -- must find it via the index without any rescan, proving the
  // index update actually took effect mid-run.
  const props = sandbox.PropertiesService.getDocumentProperties();
  props.setProperty('REVISION_MAP', JSON.stringify({ 'S4-260101': 'S4-260102' }));

  let scanCalled = false;
  const spyBody = Object.assign({}, body, { getTables: () => { scanCalled = true; return body.getTables(); } });
  const tdocData2 = {
    row: ['S4-260102', 'Revision title', 'Some source', '', '5.1', '', ''],
    tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 4, typeCol: -1, forCol: -1, statusCol: -1, revisedToCol: -1,
    agendaItem: '5.1',
    richTextRow: null
  };
  sandbox.insertNewTdoc_(spyBody, tdocData2, {}, index);
  check('the revision of a just-inserted (same-run) TDoc is placed via the index, no rescan needed', scanCalled, false);

  const newIdx = body.getChildIndex(index.get('S4-260101'));
  const revIdx = body.getChildIndex(index.get('S4-260102'));
  check('the revision table (S4-260102) was placed directly after its just-inserted parent (S4-260101)', revIdx, newIdx + 1);
}

// ============ 6. rearrangeRevisionTables_() keeps the index up to date ====

console.log('rearrangeRevisionTables_() -- updates the index to the NEW table object after moveTableAfter_() actually moves one');

{
  const { sandbox } = loadCode({ documentProperties: { TDOC_LIST_URL: '' } });
  const parent = makeTdocTable('S4-260100', '5.1');
  const child = makeTdocTable('S4-260101', '5.1');
  const unrelated = makeFakeParagraph('filler', 'NORMAL');
  // child starts BEFORE parent -- moveTableAfter_() must actually move it.
  const body = makeFakeBody([child, unrelated, parent]);
  sandbox.DocumentApp = { getActiveDocument: () => ({ getBody: () => body }) };
  const index = sandbox.buildTdocTableIndex_(body.getTables());
  const originalChildRef = index.get('S4-260101');

  // buildRevisionMapFromTdocList_() builds map[from]=to from EACH row's own
  // tdocNumberOf_()/getRevisedTo_() -- so the ORIGINAL (S4-260100) row is
  // the one whose OWN "Revised to" column names the newer document.
  const groups = {
    '5.1': {
      tdocs: [
        { row: ['S4-260100', '', '', '', '5.1', '', '', 'S4-260101'], tdocCol: 0, agendaCol: 4, revisedToCol: 7 },
        { row: ['S4-260101', '', '', '', '5.1', '', '', ''], tdocCol: 0, agendaCol: 4, revisedToCol: 7 }
      ]
    }
  };

  const stats = sandbox.rearrangeRevisionTables_({ TDOC_LIST_URL: '' }, groups, index);
  check('exactly one revision pair was found and moved', [stats.total, stats.moved], [1, 1]);

  const updatedRef = index.get('S4-260101');
  check('the index entry for the moved TDoc no longer points at the ORIGINAL (now-removed) table object',
    updatedRef !== originalChildRef, true);
  check('the index entry for the moved TDoc points at a table that is actually still in the document',
    body.getTables().indexOf(updatedRef) !== -1, true);
  check('the moved table sits directly after its parent in the document',
    body.getChildIndex(updatedRef), body.getChildIndex(index.get('S4-260100')) + 1);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
