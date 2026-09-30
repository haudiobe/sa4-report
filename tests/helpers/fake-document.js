/**
 * Fake Google Docs body for tests that run the REAL
 * buildSkeletonWithTdocTables() (moved unchanged from
 * tests/addon007a-adhoc-admin-anchors.test.js so ADDON-008A can reuse it).
 */

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
      const cells = cellTexts.map(t => ({ _t: String(t), getText: () => String(t), editAsText: () => ({ setLinkUrl() {} }) }));
      return {
        getNumCells: () => cells.length, getCell: (i) => cells[i],
        appendTableCell: (t) => { const c = { _t: String(t), getText: () => String(t), editAsText: () => ({ setLinkUrl() {} }) }; cells.push(c); return c; }
      };
    }
    (rowsData || []).forEach(r => rows.push(makeRow(r)));
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
    getTables: () => children.filter(c => c.getType() === TABLE),
    getParagraphs: () => children.filter(c => c.getType() === PARAGRAPH),
    appendParagraph: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendListItem: (text) => { const p = makeParagraph(text); children.push(p); return p; },
    appendTable: () => { const t = makeTable([]); children.push(t); return t; },
    insertParagraph: (idx, text) => { const p = makeParagraph(text); children.splice(idx, 0, p); return p; },
    insertTable: (idx, data) => { const t = makeTable(data); children.splice(idx, 0, t); return t; }
  };
  body._headingTexts = () => children.filter(c => c.getType() === PARAGRAPH && c.getHeading() !== NORMAL).map(c => c.getText());
  return body;
}

function tdoc(id, agendaItem) {
  return {
    row: [id, 'Title of ' + id, 'Source Co', agendaItem],
    richTextRow: null,
    tdocCol: 0, titleCol: 1, sourceCol: 2, contactCol: -1, agendaCol: 3,
    agendaTopicCol: -1, statusCol: -1, typeCol: -1, forCol: -1, revisedToCol: -1
  };
}

module.exports = { makeFakeDocumentBody, tdoc };
