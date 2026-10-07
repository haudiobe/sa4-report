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
 *
 * Ad-hoc attendance (stage D), additions only: a paragraph's text can be
 * set and its bold state is recorded (_bold); a link set on a cell's text is
 * recorded (_links: [start, end, url]); appendTable() takes cell texts like
 * insertTable().
 *
 * Attendee-table formatting (G2 smoke finding), additions only: a cell
 * records its bold state, background and width (_bold, _background,
 * _width), holds one paragraph whose spacing is recorded (_spacing), and a
 * row records its minimum height (_minHeight). As in Google Docs, a table
 * that is inserted or appended with cell texts TAKES ITS TEXT ATTRIBUTES
 * FROM THE PARAGRAPH BEFORE IT: its cells start bold when that paragraph is
 * bold. (That inheritance is what made the generated attendee table bold.)
 *
 * Status dropdowns (T-2026.10.7), additions only. A cell can hold a native
 * dropdown (_dropdown: { dropdownId, definitionId, optionId }), which only
 * the fake Docs API (fake-docs-api.js) creates and sets. As in Google Docs:
 * DocumentApp sees it as an UNSUPPORTED child of the cell's paragraph and
 * not in getText(); setText() on such a cell writes text and leaves the
 * dropdown; a copy of a table carries its dropdowns, each with a NEW id.
 * The body knows whether it has unsaved changes (_dirty) and, once the test
 * says the document was saved and closed (_closed), refuses every change.
 *
 * Discussion e-mails (T-2026.10.10), addition only: the paragraph of a cell
 * gives its text as a Text element (editAsText()), the one of the cell --
 * the e-mail export reads a cell paragraph by paragraph. A dropdown is not
 * in it, as in Google Docs.
 *
 * Shared minutes (after T-2026.10.10), addition only: the link at one
 * character of a body paragraph can be read (getLinkUrl()), as recorded by
 * setLinkUrl() -- the last one set wins, and null takes a link away.
 */

function makeFakeDocumentBody(sandbox) {
  const PARAGRAPH = sandbox.DocumentApp.ElementType.PARAGRAPH;
  const TABLE = sandbox.DocumentApp.ElementType.TABLE;
  const NORMAL = sandbox.DocumentApp.ParagraphHeading.NORMAL;
  const children = [];
  // Unsaved changes, and the state after saveAndClose(): see the header.
  const state = { dirty: false, closed: false, nextDropdown: 0 };
  // onFirstTouch: called before the first unsaved change, so that the fake Docs API can keep what the saved document says.
  const touch = () => { if (state.closed) throw new Error('The document is closed: it was saved and closed earlier in this execution.'); if (!state.dirty && state.onFirstTouch) state.onFirstTouch(); state.dirty = true; };

  function makeParagraph(text, heading) {
    const p = {
      _text: text, _heading: heading || NORMAL,
      getType: () => PARAGRAPH, asParagraph: () => p,
      getParent: () => (children.indexOf(p) !== -1 ? body : null),
      getText: () => p._text,
      getHeading: () => p._heading,
      setHeading: (h) => { p._heading = h; return p; },
      setGlyphType: () => p,
      setText: (v) => { p._text = String(v); return p; },
      // A link set on a paragraph's text is recorded (_links: [start, end, url]).
      editAsText: () => ({ setBold: (b) => { p._bold = !!b; }, setLinkUrl: (...args) => { (p._links = p._links || []).push(args); },
        getLinkUrl: (i) => { const links = p._links || []; for (let k = links.length - 1; k >= 0; k--) if (i >= links[k][0] && i <= links[k][1]) return links[k][2] || null; return null; } })
    };
    return p;
  }
  function makeTable(rowsData, inheritedBold) {
    const rows = [];
    function makeCell(text) {
      const c = { _t: String(text), _links: [], _bold: !!inheritedBold, _background: null, _width: null, _spacing: [null, null], _dropdown: null,
        getText: () => c._t, setText: (v) => { touch(); c._t = String(v); return c; },
        setBackgroundColor: (v) => { c._background = v; return c; }, setWidth: (v) => { c._width = v; return c; }, getWidth: () => c._width };
      const para = { getType: () => PARAGRAPH, asParagraph: () => para, getText: () => c._t,
        setSpacingBefore: (v) => { c._spacing[0] = v; return para; }, setSpacingAfter: (v) => { c._spacing[1] = v; return para; },
        // The children of the paragraph: its text, then a dropdown -- which DocumentApp knows as UNSUPPORTED only.
        getNumChildren: () => (c._t ? 1 : 0) + (c._dropdown ? 1 : 0),
        getChild: (k) => ({ getType: () => (c._t && k === 0 ? 'TEXT' : 'UNSUPPORTED') }),
        // Addition only (T-2026.10.10): the text of the paragraph, for the e-mail export.
        editAsText: () => te };
      c.getNumChildren = () => 1;
      c.getChild = () => para;
      // ADDON-008A2: a minimal Text element over the cell's text (the e-mail
      // collector appends and styles text in place); styling is a no-op.
      const te = {
        getText: () => c._t,
        appendText: (v) => { c._t += String(v); return te; },
        setLinkUrl: (...args) => { c._links.push(args); return te; },
        // Addition only (after T-2026.10.8): the link at one character, as recorded by setLinkUrl() -- the last one set wins.
        getLinkUrl: (i) => { for (let k = c._links.length - 1; k >= 0; k--) if (i >= c._links[k][0] && i <= c._links[k][1]) return c._links[k][2] || null; return null; },
        setFontSize: () => te, setForegroundColor: () => te, setBold: (b) => { c._bold = !!b; return te; }
      };
      c.editAsText = () => te;
      return c;
    }
    function makeRow(cellTexts) {
      const cells = cellTexts.map(makeCell);
      const row = {
        _minHeight: null, setMinimumHeight: (v) => { row._minHeight = v; return row; },
        getNumCells: () => cells.length, getCell: (i) => cells[i],
        appendTableCell: (t) => { touch(); const c = makeCell(t); cells.push(c); return c; }
      };
      return row;
    }
    (rowsData || []).forEach(r => rows.push(makeRow(r)));
    const t = {
      getType: () => TABLE, asTable: () => t,
      getNumRows: () => rows.length, getRow: (i) => rows[i],
      getCell: (r, c) => rows[r].getCell(c),
      appendTableRow: () => { touch(); const r = makeRow([]); rows.push(r); return r; },
      insertTableRow: (i) => { const r = makeRow([]); rows.splice(i, 0, r); return r; },
      removeRow: (i) => { rows.splice(i, 1); },
      getParent: () => (children.indexOf(t) !== -1 ? body : null),
      removeFromParent: () => { touch(); const i = children.indexOf(t); if (i !== -1) children.splice(i, 1); return t; },
      copy: () => {
        const made = makeTable(rows.map(r => { const out = []; for (let i = 0; i < r.getNumCells(); i++) out.push(r.getCell(i).getText()); return out; }));
        // A copy within the document keeps a working dropdown; it is a new dropdown with its own id.
        rows.forEach((r, ri) => { for (let i = 0; i < r.getNumCells(); i++) { const d = r.getCell(i)._dropdown; if (d) made.getRow(ri).getCell(i)._dropdown = { dropdownId: 'kix.copy' + (++state.nextDropdown), definitionId: d.definitionId, optionId: d.optionId }; } });
        return made;
      }
    };
    return t;
  }

  /** Whether the element before position `idx` is a bold paragraph: what a table inserted there starts with. */
  const boldBefore = (idx) => { const prev = children[idx - 1]; return !!(prev && prev.getType() === PARAGRAPH && prev._bold); };

  const body = {
    _children: children,
    getType: () => 'BODY_SECTION',
    asTable: () => { throw new Error("BODY_SECTION can't be cast to TABLE."); },
    asParagraph: () => { throw new Error("BODY_SECTION can't be cast to PARAGRAPH."); },
    removeChild: (c) => { touch(); const i = children.indexOf(c); if (i === -1) throw new Error('Element is not a child of this body.'); children.splice(i, 1); return body; },
    clear: () => { touch(); children.length = 0; return body; },
    getNumChildren: () => children.length,
    getChild: (i) => children[i],
    getChildIndex: (c) => children.indexOf(c),
    getTables: () => children.filter(c => c.getType() === TABLE),
    getParagraphs: () => children.filter(c => c.getType() === PARAGRAPH),
    appendParagraph: (text) => { touch(); const p = makeParagraph(text); children.push(p); return p; },
    appendListItem: (text) => { touch(); const p = makeParagraph(text); children.push(p); return p; },
    appendTable: (data) => { touch(); const t = makeTable(Array.isArray(data) ? data : [], boldBefore(children.length)); children.push(t); return t; },
    insertParagraph: (idx, text) => { touch(); const p = makeParagraph(text); children.splice(idx, 0, p); return p; },
    insertTable: (idx, data) => {
      touch();
      let t;
      if (data && typeof data.getType === 'function') {
        if (children.indexOf(data) !== -1) throw new Error('Element must be detached.');
        t = data;
      } else {
        t = makeTable(data, boldBefore(idx));
      }
      children.splice(idx, 0, t);
      return t;
    }
  };
  body._state = state;
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
