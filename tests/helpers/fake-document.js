/**
 * Fake Google Docs body for tests that run the REAL
 * buildSkeletonWithTdocTables() (moved unchanged from
 * tests/addon007a-adhoc-admin-anchors.test.js so ADDON-008A can reuse it).
 *
 * ADDON-008A1: mirrors the real Apps Script hierarchy where reallocation
 * depends on it -- the body is a BODY_SECTION (not castable to a table),
 * Body.removeChild() returns the BODY (not the removed child),
 * insertTable() only accepts a DETACHED table (e.g. from Table.copy()),
 * and elements know their parent.
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
      getParent: () => (children.indexOf(p) !== -1 ? body : null),
      getText: () => p._text,
      getHeading: () => p._heading,
      setHeading: (h) => { p._heading = h; return p; },
      setGlyphType: () => p
    };
    return p;
  }
  function makeTable(rowsData) {
    const rows = [];
    function makeCell(text) {
      const c = { _t: String(text), getText: () => c._t, setText: (v) => { c._t = String(v); return c; } };
      // ADDON-008A2: a minimal Text element over the cell's text (the e-mail
      // collector appends and styles text in place); styling is a no-op.
      const te = {
        getText: () => c._t,
        appendText: (v) => { c._t += String(v); return te; },
        setLinkUrl: () => te, setFontSize: () => te, setForegroundColor: () => te, setBold: () => te
      };
      c.editAsText = () => te;
      return c;
    }
    function makeRow(cellTexts) {
      const cells = cellTexts.map(makeCell);
      return {
        getNumCells: () => cells.length, getCell: (i) => cells[i],
        appendTableCell: (t) => { const c = makeCell(t); cells.push(c); return c; }
      };
    }
    (rowsData || []).forEach(r => rows.push(makeRow(r)));
    const t = {
      getType: () => TABLE, asTable: () => t,
      getNumRows: () => rows.length, getRow: (i) => rows[i],
      getCell: (r, c) => rows[r].getCell(c),
      appendTableRow: () => { const r = makeRow([]); rows.push(r); return r; },
      insertTableRow: (i) => { const r = makeRow([]); rows.splice(i, 0, r); return r; },
      removeRow: (i) => { rows.splice(i, 1); },
      getParent: () => (children.indexOf(t) !== -1 ? body : null),
      removeFromParent: () => { const i = children.indexOf(t); if (i !== -1) children.splice(i, 1); return t; },
      copy: () => makeTable(rows.map(r => { const out = []; for (let i = 0; i < r.getNumCells(); i++) out.push(r.getCell(i).getText()); return out; }))
    };
    return t;
  }

  const body = {
    _children: children,
    getType: () => 'BODY_SECTION',
    asTable: () => { throw new Error("BODY_SECTION can't be cast to TABLE."); },
    asParagraph: () => { throw new Error("BODY_SECTION can't be cast to PARAGRAPH."); },
    removeChild: (c) => { const i = children.indexOf(c); if (i === -1) throw new Error('Element is not a child of this body.'); children.splice(i, 1); return body; },
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
    insertTable: (idx, data) => {
      let t;
      if (data && typeof data.getType === 'function') {
        if (children.indexOf(data) !== -1) throw new Error('Element must be detached.');
        t = data;
      } else {
        t = makeTable(data);
      }
      children.splice(idx, 0, t);
      return t;
    }
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
