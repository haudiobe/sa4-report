/**
 * Fake Google Docs API (the `Docs` advanced service) for the status-dropdown
 * tests (T-2026.10.7). No live service is needed for the regression suite.
 *
 * It serves ONE report document, which is the fake DocumentApp body of
 * tests/helpers/fake-document.js, and any number of read-only source
 * documents (the master template). What it models is what the production
 * code relies on, as it was observed in the live probes:
 *
 *   - Documents.get() describes the report as the API does: tabs, the
 *     dropdown definitions of the tab, and the body with start and end
 *     indexes. A cell's text is a textRun that ends with its line end; a
 *     dropdown is one element of length 1 with dropdownId and
 *     dropdownProperties { dropdownDefinitionId, selectedOptionId,
 *     displayValue }.
 *   - Documents.batchUpdate() knows createDropdownDefinition,
 *     deleteContentRange, insertDropdown and updateDropdownProperties, and
 *     (after T-2026.10.29) updateDropdownDefinitionProperties with the field
 *     mask "options": the options are the complete new list, an option with
 *     the optionId of an existing one keeps it, and an option that is
 *     removed while a dropdown selects it needs a replacement in
 *     selectedOptionIdReplacements (else the request is refused, as the API
 *     answers 400); nothing else. Requests are applied in order, each against the indexes
 *     AS THEY ARE AFTER the requests before it. A batch is applied as a
 *     whole or not at all.
 *   - After T-2026.10.29 (additions only): a cell can mix text and several
 *     dropdowns in one paragraph (fake-document.js `_parts`), rendered in
 *     order; insertText, deleteContentRange and insertDropdown work at any
 *     position inside one cell (a range that would delete a dropdown is
 *     refused); a body element can be preceded by a section break
 *     (`_sectionBreakBefore`), made by insertSectionBreak (NEXT_PAGE; the
 *     newline Google Docs puts before it is not modelled) and styled by
 *     updateSectionStyle (flipPageOrientation only; the first section of
 *     the document is never written). Paragraphs carry namedStyleType.
 *   - Round 3 (additions only): text runs carry textStyle { bold, link }
 *     (cut where a link begins or ends), table cells their startIndex and
 *     endIndex, tables tableStyle.tableColumnProperties, the tab its
 *     documentStyle (A4); the formatting requests updateTableColumnProperties,
 *     updateTableRowStyle, updateTableCellStyle, updateTextStyle (bold, link)
 *     and updateParagraphStyle (spacing) work inside one table or one cell.
 *   - A read sees the SAVED document: while the document has unsaved
 *     DocumentApp changes, Documents.get() describes it as it was before
 *     the first of them -- a table added in this execution is not there.
 *   - The document has a revision. It changes with every write of the API,
 *     with every DocumentApp change that is saved, and when a user edits the
 *     document (pick(), edited()). Documents.get() names the revision it
 *     describes; a batchUpdate with writeControl.requiredRevisionId is
 *     refused unless that is the revision the document is at, and every
 *     reply names the revision the document is at afterwards.
 *   - A write is refused while the document has unsaved DocumentApp changes:
 *     the API works on the saved document, so a write before the save would
 *     be a write to other positions than the code believes.
 *
 * `calls` lists every call; `fail` lets a test make calls fail.
 */

const REPORT_TAB_ID = 't.0';
/** After T-2026.10.29: the TDoc number of a TDoc cell that may read "S4-26XXXX (revision of S4-26YYYY)". */
const numberOfTdocCell = (text) => String(text).replace(/\s*\(revision of [^)]*\)\s*$/, '').trim();

function makeFakeDocsApi(options) {
  const body = options.body;
  const reportId = options.reportId;
  const sources = {};
  const definitions = {};                    // of the report: id -> { dropdownDefinitionId, dropdownDefinitionProperties }
  const counters = { definition: 0, dropdown: 0, revision: 1 };
  const revisionOf = (n) => 'rev.' + n;
  const api = { calls: [], fail: { get: null, batchUpdate: null }, definitions: definitions, tabId: REPORT_TAB_ID };
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const refuse = (message) => { const e = new Error(message); e.fakeDocsApi = true; return e; };

  const optionOf = (dropdown) => {
    const def = definitions[dropdown.definitionId];
    const options = def ? def.dropdownDefinitionProperties.options : [];
    return options.filter((o) => o.optionId === dropdown.optionId)[0] || null;
  };

  /** Round 3: the runs of a cell's text from index `at` -- cut where a link begins or ends -- each with its bold and link. */
  function styledRuns(cell, text, at, newline) {
    const linkAt = (k) => { for (let n = cell._links.length - 1; n >= 0; n--) if (k >= cell._links[n][0] && k <= cell._links[n][1]) return cell._links[n][2] || null; return null; };
    const runs = [];
    let from = 0;
    for (let k = 1; k <= text.length; k++) if (k === text.length || linkAt(k) !== linkAt(k - 1)) { runs.push([from, k]); from = k; }
    if (!runs.length) runs.push([0, 0]);
    return runs.map(([a, b], n) => {
      const last = n === runs.length - 1 && newline;
      const style = { bold: !!cell._bold };
      const url = text.length ? linkAt(a) : null;
      if (url) style.link = { url: url };
      return { startIndex: at + a, endIndex: at + b + (last ? 1 : 0), textRun: { content: text.slice(a, b) + (last ? '\n' : ''), textStyle: style } };
    });
  }

  /** The report as the API describes it, and where each table cell is: [{ cell, start, textStart, textEnd }]. */
  // Round 3: a render is kept while no request moves a position (style requests do not); `fresh` drops it.
  let kept = null;
  const fresh = () => { kept = null; };
  function render() {
    if (kept) return kept;
    kept = renderNow();
    return kept;
  }
  function renderNow() {
    const content = [];
    const cells = [];
    let i = 1;
    body._children.forEach((child) => {
      // After T-2026.10.29: a section break before the element.
      if (child._sectionBreakBefore) {
        content.push({ startIndex: i, endIndex: i + 1, sectionBreak: { sectionStyle: { sectionType: 'NEXT_PAGE', flipPageOrientation: !!child._sectionBreakBefore.flip } }, _child: child });
        i += 1;
      }
      if (child.getType() === 'TABLE') {
        const element = { startIndex: i, table: { tableRows: [] } };
        i += 1;
        for (let r = 0; r < child.getNumRows(); r++) {
          const row = { tableCells: [] };
          i += 1;
          for (let c = 0; c < child.getRow(r).getNumCells(); c++) {
            const cell = child.getRow(r).getCell(c);
            i += 1;
            const elements = [];
            const start = i;
            const text = cell._t;
            const dropdown = cell._dropdown;
            if (cell._parts) {
              // After T-2026.10.29: text and dropdowns in their order, then the line end.
              const positions = [];
              cell._parts.forEach((p) => {
                if (p.d) {
                  const option = optionOf(p.d);
                  elements.push({ startIndex: i, endIndex: i + 1, dropdown: { dropdownId: p.d.dropdownId,
                    dropdownProperties: { dropdownDefinitionId: p.d.definitionId, selectedOptionId: p.d.optionId, displayValue: option ? option.displayValue : '' } } });
                  positions.push({ part: p, at: i, length: 1 });
                  i += 1;
                } else {
                  elements.push({ startIndex: i, endIndex: i + p.t.length, textRun: { content: p.t } });
                  positions.push({ part: p, at: i, length: p.t.length });
                  i += p.t.length;
                }
              });
              elements.push({ startIndex: i, endIndex: i + 1, textRun: { content: '\n' } });
              cells.push({ cell: cell, start: start, textStart: null, textEnd: null, contentEnd: i, table: child, r: r, c: c });
              i += 1;
              row.tableCells.push({ startIndex: start - 1, endIndex: i, content: [{ startIndex: start, endIndex: i, paragraph: { elements: elements } }] });
              continue;
            }
            if (text && !dropdown) {
              styledRuns(cell, text, i, true).forEach((el) => elements.push(el));
              i += text.length + 1;
            } else {
              if (text) { elements.push({ startIndex: i, endIndex: i + text.length, textRun: { content: text } }); i += text.length; }
              if (dropdown) {
                const option = optionOf(dropdown);
                elements.push({ startIndex: i, endIndex: i + 1, dropdown: { dropdownId: dropdown.dropdownId,
                  dropdownProperties: { dropdownDefinitionId: dropdown.definitionId, selectedOptionId: dropdown.optionId, displayValue: option ? option.displayValue : '' } } });
                i += 1;
              }
              elements.push({ startIndex: i, endIndex: i + 1, textRun: { content: '\n' } });
              i += 1;
            }
            cells.push({ cell: cell, start: start, textStart: start, textEnd: start + text.length, contentEnd: i - 1, table: child, r: r, c: c });
            row.tableCells.push({ startIndex: start - 1, endIndex: i, content: [{ startIndex: start, endIndex: i, paragraph: { elements: elements } }] });
          }
          element.table.tableRows.push(row);
        }
        i += 1;
        element.endIndex = i;
        element._child = child;
        // Round 3: the widths of the columns, as the first row has them.
        element.table.tableStyle = { tableColumnProperties: Array.from({ length: child.getNumRows() ? child.getRow(0).getNumCells() : 0 }, (x, k) => {
          const w = child.getRow(0).getCell(k)._width;
          return typeof w === 'number' ? { width: { magnitude: w, unit: 'PT' }, widthType: 'FIXED_WIDTH' } : { widthType: 'EVENLY_DISTRIBUTED' };
        }) };
        content.push(element);
      } else {
        const text = child.getText();
        content.push({ startIndex: i, endIndex: i + text.length + 1, _child: child, paragraph: { paragraphStyle: { namedStyleType: child.getHeading && child.getHeading() !== 'NORMAL' ? 'HEADING_1' : 'NORMAL_TEXT' },
          elements: [{ startIndex: i, endIndex: i + text.length + 1, textRun: { content: text + '\n', textStyle: { bold: !!child._bold } } }] } });
        i += text.length + 1;
      }
    });
    return { content: content, cells: cells };
  }

  const allCells = () => render().cells.map((x) => x.cell);

  // The saved document: kept the moment before DocumentApp makes its first unsaved change.
  // The unsaved changes of DocumentApp will be another revision once they are saved; until then the API is at the old one.
  let saved = null;
  const publicContent = () => render().content.map((e) => { const o = Object.assign({}, e); delete o._child; return o; });
  body._state.onFirstTouch = () => { fresh(); const was = body._state.paused; body._state.paused = true; try { saved = clone({ definitions: definitions, content: publicContent(), revisionId: revisionOf(counters.revision) }); } finally { body._state.paused = was; fresh(); } counters.revision++; };
  const view = () => (body._state.dirty && saved ? saved : { definitions: definitions, content: publicContent(), revisionId: revisionOf(counters.revision) });

  /** After T-2026.10.29: the cell whose content (its line end aside) holds index `at` (inclusive at its end), with its parts and their positions. */
  function cellAt(at) {
    const hit = render().cells.filter((x) => at >= x.start && at <= x.contentEnd)[0];
    if (!hit) return null;
    const parts = hit.cell._allParts();
    let pos = hit.start;
    const placed = parts.map((p) => { const out = { part: p, at: pos, length: p.d ? 1 : p.t.length }; pos += out.length; return out; });
    return { hit: hit, parts: parts, placed: placed };
  }
  function setParts(cell, parts) { cell._parts = parts; cell._sync(); }

  function applyRequest(request, n) {
    const kind = Object.keys(request)[0];
    if (['deleteContentRange', 'insertText', 'insertDropdown', 'insertSectionBreak', 'updateSectionStyle', 'updateDropdownProperties', 'updateDropdownDefinitionProperties', 'createDropdownDefinition'].indexOf(kind) !== -1) fresh();
    const r = request[kind];
    const bad = (message) => refuse('Invalid requests[' + n + '].' + kind + ': ' + message);
    const checkTab = (tabId) => { if (tabId !== undefined && tabId !== REPORT_TAB_ID) throw bad('The tab ' + tabId + ' does not exist.'); };

    if (kind === 'createDropdownDefinition') {
      checkTab(r.tabId);
      const props = r.dropdownDefinition && r.dropdownDefinition.dropdownDefinitionProperties;
      if (!props || !Array.isArray(props.options) || props.options.length < 2 || props.options.length > 50) throw bad('A dropdown definition must have at least 2 options and at most 50 options.');
      props.options.forEach((o) => {
        Object.keys(o).forEach((k) => { if (['displayValue', 'optionId', 'textStyle'].indexOf(k) === -1) throw bad('Unknown field ' + k + '.'); });
        Object.keys(o.textStyle || {}).forEach((k) => { if (k !== 'foregroundColor' && k !== 'backgroundColor') throw bad('Only the foreground_color and background_color properties are supported.'); });
      });
      const id = 'kix.def' + (++counters.definition);
      definitions[id] = { dropdownDefinitionId: id, dropdownDefinitionProperties: { title: props.title,
        options: props.options.map((o, k) => Object.assign({ optionId: 'dropdownItem.d' + counters.definition + 'o' + k }, clone(o))) } };
      return { createDropdownDefinition: { dropdownDefinition: clone(definitions[id]) } };
    }

    if (kind === 'deleteContentRange') {
      const range = r.range || {};
      checkTab(range.tabId);
      const at = range.startIndex < range.endIndex ? cellAt(range.startIndex) : null;
      if (!at || range.endIndex > at.hit.contentEnd) throw bad('The range ' + range.startIndex + '..' + range.endIndex + ' is not text inside one table cell.');
      if (at.placed.some((x) => x.part.d && x.at >= range.startIndex && x.at < range.endIndex)) throw bad('FAKE DOCS API: the range ' + range.startIndex + '..' + range.endIndex + ' would delete a dropdown.');
      at.placed.forEach((x) => {
        if (x.part.d) return;
        const a = Math.max(range.startIndex, x.at) - x.at;
        const b = Math.min(range.endIndex, x.at + x.length) - x.at;
        if (b > a) x.part.t = x.part.t.slice(0, a) + x.part.t.slice(b);
      });
      setParts(at.hit.cell, at.parts);
      return {};
    }

    if (kind === 'insertText') {
      const location = r.location || {};
      checkTab(location.tabId);
      const at = cellAt(location.index);
      if (!at) throw bad('Index ' + location.index + ' is not inside a table cell (the fake writes text into table cells only).');
      if (!r.text) throw bad('The text is empty.');
      let done = false;
      at.placed.forEach((x) => {
        if (done || x.part.d || location.index < x.at || location.index > x.at + x.length) return;
        x.part.t = x.part.t.slice(0, location.index - x.at) + r.text + x.part.t.slice(location.index - x.at);
        done = true;
      });
      if (!done) {
        const k = at.placed.filter((x) => x.at < location.index).length;
        at.parts.splice(k, 0, { t: r.text });
      }
      setParts(at.hit.cell, at.parts);
      return {};
    }

    if (kind === 'insertDropdown') {
      const location = r.location || {};
      checkTab(location.tabId);
      const def = definitions[r.dropdownDefinitionId];
      if (!def) throw bad('The dropdown definition ' + r.dropdownDefinitionId + ' does not exist.');
      const options = def.dropdownDefinitionProperties.options;
      const optionId = r.selectedOptionId === undefined ? options[0].optionId : r.selectedOptionId;
      if (!options.some((o) => o.optionId === optionId)) throw bad('The option ' + optionId + ' is not an option of the dropdown definition.');
      const at = cellAt(location.index);
      if (!at) throw bad('Index ' + location.index + ' is not the start of a table cell.');
      // As before for a cell's start: a cell that has a dropdown there, or one without the compact Status row, gets no second one.
      if (location.index === at.hit.start && at.hit.cell._dropdown && (at.parts.length && at.parts[0].d || at.hit.cell._t.indexOf('\u2003') === -1)) throw bad('The cell already has a dropdown.');
      const made = { dropdownId: 'kix.dd' + (++counters.dropdown), definitionId: r.dropdownDefinitionId, optionId: optionId };
      // As before: a dropdown at the start of a plain cell is the cell's one dropdown, shown after any text left in it.
      if (location.index === at.hit.start && !at.hit.cell._parts && !at.hit.cell._dropdown && at.hit.cell._t.indexOf(' ') === -1) {
        at.hit.cell._dropdown = made;
        return { insertDropdown: { dropdown: { dropdownId: made.dropdownId } } };
      }
      let k = 0;
      let split = false;
      at.placed.forEach((x, n) => {
        if (x.at + x.length <= location.index) { k = n + 1; return; }
        if (!split && !x.part.d && x.at < location.index && location.index < x.at + x.length) {
          const cut = location.index - x.at;
          at.parts.splice(n, 1, { t: x.part.t.slice(0, cut) }, { t: x.part.t.slice(cut) });
          k = n + 1;
          split = true;
        }
      });
      at.parts.splice(k, 0, { d: made });
      setParts(at.hit.cell, at.parts);
      return { insertDropdown: { dropdown: { dropdownId: made.dropdownId } } };
    }

    if (kind === 'updateDropdownProperties') {
      checkTab(r.tabId);
      // After T-2026.10.29: the dropdown may be any one of a cell's.
      let found = null;
      allCells().forEach((c) => { if (!found) found = (c._parts ? c._parts.map((p) => p.d).filter(Boolean) : [c._dropdown].filter(Boolean)).filter((d) => d.dropdownId === r.dropdownId)[0] || null; });
      if (!found) throw bad('The dropdown ' + r.dropdownId + ' does not exist.');
      if (!/(^|,)\s*(selectedOptionId|\*)\s*(,|$)/.test(String(r.fields || ''))) throw bad('At least one field must be specified.');
      const def = definitions[found.definitionId];
      const optionId = r.dropdownProperties && r.dropdownProperties.selectedOptionId;
      if (!def || !def.dropdownDefinitionProperties.options.some((o) => o.optionId === optionId)) throw bad('The option ' + optionId + ' is not an option of the dropdown definition.');
      found.optionId = optionId;
      return {};
    }

    if (kind === 'insertSectionBreak') {
      const location = r.location || {};
      checkTab(location.tabId);
      if (r.sectionType !== 'NEXT_PAGE' && r.sectionType !== 'CONTINUOUS') throw bad('A section type is required.');
      const target = render().content.filter((e) => e._child && e.startIndex === location.index && !e.sectionBreak)[0];
      if (!target) throw bad('Index ' + location.index + ' is not the start of a body element (the fake inserts section breaks there only).');
      if (target._child._sectionBreakBefore) throw bad('FAKE DOCS API: there is a section break before that element already.');
      target._child._sectionBreakBefore = { flip: false, type: r.sectionType };
      return {};
    }

    if (kind === 'updateSectionStyle') {
      const range = r.range || {};
      checkTab(range.tabId);
      if (!/(^|,)\s*flipPageOrientation\s*(,|$)/.test(String(r.fields || ''))) throw bad('The fake updates flipPageOrientation only.');
      const breaks = render().content.filter((e) => e.sectionBreak && e.startIndex <= range.startIndex);
      if (!breaks.length) throw bad('FAKE DOCS API: the first section of the document is never written.');
      breaks[breaks.length - 1]._child._sectionBreakBefore.flip = !!(r.sectionStyle && r.sectionStyle.flipPageOrientation);
      return {};
    }

    if (kind === 'updateDropdownDefinitionProperties') {
      checkTab(r.tabId);
      const def = definitions[r.dropdownDefinitionId];
      if (!def) throw bad('The dropdown definition ' + r.dropdownDefinitionId + ' does not exist.');
      if (!/(^|,)\s*options\s*(,|$)/.test(String(r.fields || ''))) throw bad('The field mask must name options.');
      const wanted = (r.dropdownDefinitionProperties && r.dropdownDefinitionProperties.options) || [];
      if (wanted.length < 2 || wanted.length > 50) throw bad('A dropdown definition must have at least 2 options and at most 50 options.');
      const old = def.dropdownDefinitionProperties.options;
      const made = wanted.map((o, k) => {
        Object.keys(o).forEach((key) => { if (['displayValue', 'optionId', 'textStyle'].indexOf(key) === -1) throw bad('Unknown field ' + key + '.'); });
        if (o.optionId && !old.some((x) => x.optionId === o.optionId)) throw bad('The option ' + o.optionId + ' is not an option of the dropdown definition.');
        return Object.assign(clone(o), { optionId: o.optionId || 'dropdownItem.u' + (++counters.definition) + 'o' + k });
      });
      const kept = made.map((o) => o.optionId);
      const replacements = r.selectedOptionIdReplacements || {};
      allCells().forEach((c) => {
        (c._parts ? c._parts.map((p) => p.d).filter(Boolean) : [c._dropdown].filter(Boolean)).forEach((d) => {
          if (d.definitionId !== r.dropdownDefinitionId || kept.indexOf(d.optionId) !== -1) return;
          const to = replacements[d.optionId];
          if (!to || kept.indexOf(to) === -1) throw bad('The option ' + d.optionId + ' is selected in a dropdown and has no replacement.');
          d.optionId = to;
        });
      });
      def.dropdownDefinitionProperties.options = made;
      return {};
    }

    // Round 3: the formatting requests of the summary's tables.
    const tableAt = (location) => { const e = render().content.filter((x) => x.table && x.startIndex === (location || {}).index)[0]; if (!e) throw bad('There is no table at index ' + (location || {}).index + '.'); return e._child; };
    const cellOfRange = (range) => {
      const hit = render().cells.filter((x) => range.startIndex >= x.start && range.endIndex <= x.contentEnd + 1 && range.startIndex < range.endIndex)[0];
      if (!hit) throw bad('The range ' + range.startIndex + '..' + range.endIndex + ' is not inside one table cell (the fake formats inside one cell only).');
      return hit;
    };
    if (kind === 'updateTableColumnProperties') {
      checkTab((r.tableStartLocation || {}).tabId);
      const t = tableAt(r.tableStartLocation);
      const w = r.tableColumnProperties && r.tableColumnProperties.width && r.tableColumnProperties.width.magnitude;
      if (typeof w !== 'number') throw bad('A width is required.');
      (r.columnIndices || []).forEach((k) => { for (let row = 0; row < t.getNumRows(); row++) if (t.getRow(row).getCell(k)) t.getRow(row).getCell(k)._width = w; });
      return {};
    }
    if (kind === 'updateTableRowStyle') {
      checkTab((r.tableStartLocation || {}).tabId);
      const t = tableAt(r.tableStartLocation);
      (r.rowIndices || []).forEach((k) => { if (t.getRow(k)) t.getRow(k)._minHeight = r.tableRowStyle.minRowHeight.magnitude; });
      return {};
    }
    if (kind === 'updateTableCellStyle') {
      const location = r.tableRange && r.tableRange.tableCellLocation;
      checkTab((location && location.tableStartLocation || {}).tabId);
      const t = tableAt(location && location.tableStartLocation);
      const rgb = r.tableCellStyle.backgroundColor.color.rgbColor;
      const hex = '#' + [rgb.red, rgb.green, rgb.blue].map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
      for (let row = location.rowIndex; row < location.rowIndex + r.tableRange.rowSpan; row++) for (let k = location.columnIndex; k < location.columnIndex + r.tableRange.columnSpan; k++) t.getRow(row).getCell(k)._background = hex;
      return {};
    }
    if (kind === 'updateTextStyle') {
      checkTab((r.range || {}).tabId);
      const hit = cellOfRange(r.range || {});
      const fields = String(r.fields || '');
      if (/(^|,)\s*bold\s*(,|$)/.test(fields)) hit.cell._bold = !!r.textStyle.bold;
      if (/(^|,)\s*link\s*(,|$)/.test(fields)) {
        if (hit.textStart === null) throw bad('FAKE DOCS API: links are set in cells of text only.');
        hit.cell._links.push([r.range.startIndex - hit.start, r.range.endIndex - hit.start - 1, r.textStyle.link ? r.textStyle.link.url : null]);
      }
      return {};
    }
    if (kind === 'updateParagraphStyle') {
      checkTab((r.range || {}).tabId);
      const hit = cellOfRange(r.range || {});
      hit.cell._spacing = [r.paragraphStyle.spaceAbove.magnitude, r.paragraphStyle.spaceBelow.magnitude];
      return {};
    }

    throw bad('This request is not supported by the fake Docs API.');
  }

  let getNow = null;
  let batchNow = null;
  api.Docs = { Documents: {
    get: (id, opts) => { const was = body._state.paused; body._state.paused = true; fresh(); try { return getNow(id, opts); } finally { body._state.paused = was; fresh(); } },
    _get: (id, opts) => {
      api.calls.push({ call: 'get', id: id, fields: (opts && opts.fields) || null, includeTabsContent: !!(opts && opts.includeTabsContent) });
      if (api.fail.get) { const message = api.fail.get(id, opts, api.calls.length); if (message) throw refuse(message); }
      if (id === reportId) {
        const v = view();
        return clone({ revisionId: v.revisionId, tabs: [{ tabProperties: { tabId: REPORT_TAB_ID }, documentTab: { dropdownDefinitions: v.definitions,
          documentStyle: { pageSize: { height: { magnitude: 841.89, unit: 'PT' }, width: { magnitude: 595.28, unit: 'PT' } }, marginLeft: { magnitude: 72, unit: 'PT' }, marginRight: { magnitude: 72, unit: 'PT' } },
          body: { content: v.content } } }] });
      }
      if (sources[id]) return clone(sources[id]);
      throw refuse('Requested entity was not found.');
    },
    batchUpdate: (resource, id) => { const was = body._state.paused; body._state.paused = true; fresh(); try { return batchNow(resource, id); } finally { body._state.paused = was; fresh(); } },
    _batchUpdate: (resource, id) => {
      const requests = (resource && resource.requests) || [];
      api.calls.push({ call: 'batchUpdate', id: id, requests: clone(requests), dirty: body._state.dirty });
      if (id !== reportId) throw refuse('The fake Docs API writes to the report only: ' + id);
      if (api.fail.batchUpdate) { const message = api.fail.batchUpdate(requests, api.calls.length); if (message) throw refuse(message); }
      if (body._state.dirty) throw refuse('FAKE DOCS API: the document has unsaved DocumentApp changes; a write now would go to other positions than the caller read.');
      if (api.beforeWrite) api.beforeWrite(requests, api.calls.length);          // a test lets somebody edit the document at this very moment
      const required = resource && resource.writeControl && resource.writeControl.requiredRevisionId;
      api.calls[api.calls.length - 1].requiredRevisionId = required || null;
      if (required && required !== revisionOf(counters.revision)) throw refuse('Invalid requests: The required revision ID \'' + required + '\' does not match the latest revision.');
      // All or nothing.
      const before = { cells: allCells().map((c) => ({ c: c, t: c._t, d: c._dropdown ? Object.assign({}, c._dropdown) : null, parts: c._parts ? clone(c._parts) : null, links: c._links.slice(), bold: c._bold, width: c._width, background: c._background, spacing: c._spacing.slice() })), definitions: clone(definitions), counters: Object.assign({}, counters),
        breaks: body._children.map((ch) => [ch, ch._sectionBreakBefore ? Object.assign({}, ch._sectionBreakBefore) : undefined]) };
      try {
        const replies = requests.map((request, n) => applyRequest(request, n));
        counters.revision++;
        return { replies: replies, writeControl: api.noRevisionInReply ? undefined : { requiredRevisionId: revisionOf(counters.revision) } };
      } catch (e) {
        before.cells.forEach((x) => { x.c._t = x.t; x.c._dropdown = x.d; x.c._parts = x.parts; x.c._links = x.links; x.c._bold = x.bold; x.c._width = x.width; x.c._background = x.background; x.c._spacing = x.spacing; });
        before.breaks.forEach(([ch, b]) => { if (b) ch._sectionBreakBefore = b; else delete ch._sectionBreakBefore; });
        Object.keys(definitions).forEach((k) => delete definitions[k]);
        Object.assign(definitions, before.definitions);
        Object.assign(counters, before.counters);
        throw e;
      }
    }
  } };

  getNow = api.Docs.Documents._get;
  batchNow = api.Docs.Documents._batchUpdate;

  /** A read-only document with one dropdown definition per entry of `list` ({ title, options: [{ displayValue, textStyle }] }), in a tab or a child tab. */
  api.addSource = (id, list, inChildTab) => {
    const defs = {};
    list.forEach((props, n) => {
      defs['kix.src' + n] = { dropdownDefinitionId: 'kix.src' + n, dropdownDefinitionProperties: { title: props.title,
        options: props.options.map((o, k) => Object.assign({ optionId: 'dropdownItem.s' + n + 'o' + k }, clone(o))) } };
    });
    const tab = { tabProperties: { tabId: 't.0' }, documentTab: { dropdownDefinitions: defs } };
    sources[id] = { tabs: inChildTab ? [{ tabProperties: { tabId: 't.0' }, documentTab: {}, childTabs: [Object.assign({}, tab, { tabProperties: { tabId: 't.1' } })] }] : [tab] };
  };

  /** The Status cell of the TDoc table of `number`, as the fake DocumentApp has it (the first such table), or null. */
  api.statusCell = (number) => {
    const tables = body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc' && numberOfTdocCell(c.getRow(0).getCell(1).getText()) === number);
    if (!tables.length) return null;
    for (let r = 0; r < tables[0].getNumRows(); r++) if (tables[0].getRow(r).getCell(0).getText() === 'Status') return tables[0].getRow(r).getCell(1);
    return null;
  };
  /** What a reader of the document sees in the Status cell of `number`: 'text:agreed', 'dropdown:agreed', 'text:agreed+dropdown:noted', 'none'. */
  api.shown = (number) => {
    const cell = api.statusCell(number);
    if (!cell) return 'none';
    if (cell._parts || cell._t.indexOf('\u200B\u2003Plenary') !== -1) return api.compactOf(number).status;
    const parts = [];
    if (cell._t) parts.push('text:' + cell._t);
    if (cell._dropdown) { const option = optionOf(cell._dropdown); parts.push('dropdown:' + (option ? option.displayValue : '?')); }
    return parts.join('+') || 'empty';
  };
  /** A user picks another option in the dropdown of `number` (this is not a DocumentApp change: nothing becomes unsaved). */
  api.pick = (number, displayValue) => {
    const cell = api.statusCell(number);
    // After T-2026.10.29: in a compact Status row, the status dropdown is the one before " Plenary".
    const d = cell._parts ? (cell._parts[0] && cell._parts[0].d) : cell._dropdown;
    const def = definitions[d.definitionId];
    d.optionId = def.dropdownDefinitionProperties.options.filter((o) => o.displayValue === displayValue)[0].optionId;
    counters.revision++;
  };
  /**
   * After T-2026.10.29: what a reader sees in the compact Status row of `number` --
   * { status: 'dropdown:agreed' | 'text:...' | 'empty', plenary: '☐ No' | '☑ Yes' | null, ai: 'A.I.: 15.4' | null, parts }.
   */
  api.compactOf = (number) => {
    const cell = api.statusCell(number);
    if (!cell) return null;
    const parts = cell._allParts();
    const k = parts.findIndex((p) => !p.d && p.t.indexOf('\u200B\u2003Plenary') !== -1);
    const head = k === -1 ? parts : parts.slice(0, k).concat(parts[k].t.indexOf('\u200B\u2003Plenary') > 0 ? [{ t: parts[k].t.slice(0, parts[k].t.indexOf('\u200B\u2003Plenary')) }] : []);
    const tail = k === -1 ? [] : [{ t: parts[k].t.slice(parts[k].t.indexOf('\u200B\u2003Plenary')) }].concat(parts.slice(k + 1));
    const describe = (list) => list.map((p) => (p.d ? 'dropdown:' + ((optionOf(p.d) || {}).displayValue || '?') : 'text:' + p.t)).join('+') || 'empty';
    const plenaryDropdown = tail.filter((p) => p.d)[0];
    const tailText = tail.filter((p) => !p.d).map((p) => p.t).join('');
    const ai = tailText.match(/\u2003(A\.I\.: .*)$/);
    return { status: describe(head), plenary: plenaryDropdown ? (optionOf(plenaryDropdown.d) || {}).displayValue || '?' : null, ai: ai ? ai[1] : null,
      parts: parts.map((p) => (p.d ? '[' + ((optionOf(p.d) || {}).displayValue || '?') + ']' : p.t.replace(/\u200B/g, ''))).join('') };
  };
  /** After T-2026.10.29: a user picks a value in the Plenary dropdown of the compact Status row of `number`. */
  api.pickPlenary = (number, displayValue) => {
    const cell = api.statusCell(number);
    const parts = cell._allParts();
    const k = parts.findIndex((p) => !p.d && p.t.indexOf('\u200B\u2003Plenary') !== -1);
    const d = parts.slice(k + 1).filter((p) => p.d)[0].d;
    const live = (cell._parts || []).map((p) => p.d).filter(Boolean).filter((x) => x.dropdownId === d.dropdownId)[0] || cell._dropdown;
    live.optionId = definitions[live.definitionId].dropdownDefinitionProperties.options.filter((o) => o.displayValue === displayValue)[0].optionId;
    counters.revision++;
  };
  /** After T-2026.10.29: the section breaks of the report, in order: [{ before: text of the element after it, flip }]. */
  api.sectionBreaks = () => body._children.filter((ch) => ch._sectionBreakBefore).map((ch) => ({ before: ch.getType() === 'TABLE' ? '<table>' : ch.getText(), flip: !!ch._sectionBreakBefore.flip }));
  /** After T-2026.10.29: the value cell beside `label` in the TDoc table of `number` (the first such table), or null. */
  api.cellOf = (number, label) => {
    const tables = body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc' && numberOfTdocCell(c.getRow(0).getCell(1).getText()) === number);
    if (!tables.length) return null;
    for (let r = 0; r < tables[0].getNumRows(); r++) if (tables[0].getRow(r).getCell(0).getText() === label) return tables[0].getRow(r).getCell(1);
    return null;
  };
  /** What a reader sees in the cell beside `label` of `number`, in the form of shown(). */
  api.shownIn = (number, label) => {
    const cell = api.cellOf(number, label);
    if (!cell) return 'none';
    const parts = [];
    if (cell._t) parts.push('text:' + cell._t);
    if (cell._dropdown) { const option = optionOf(cell._dropdown); parts.push('dropdown:' + (option ? option.displayValue : '?')); }
    return parts.join('+') || 'empty';
  };
  /** A user picks an option in the dropdown beside `label` of `number`. */
  api.pickIn = (number, label, displayValue) => {
    const cell = api.cellOf(number, label);
    const def = definitions[cell._dropdown.definitionId];
    cell._dropdown.optionId = def.dropdownDefinitionProperties.options.filter((o) => o.displayValue === displayValue)[0].optionId;
    counters.revision++;
  };
  /** The report's definition of that title: { id, options: [displayValue] }, or null. */
  api.definitionTitled = (title) => {
    const id = Object.keys(definitions).filter((k) => definitions[k].dropdownDefinitionProperties.title === title)[0];
    return id ? { id: id, options: definitions[id].dropdownDefinitionProperties.options.map((o) => o.displayValue) } : null;
  };
  /** Somebody edited the document outside this execution (the test changed the fake document itself): it is at a new revision. */
  api.edited = () => { counters.revision++; };
  api.revision = () => revisionOf(counters.revision);
  api.count = (call, id) => api.calls.filter((c) => c.call === call && (id === undefined || c.id === id)).length;
  return api;
}

module.exports = { makeFakeDocsApi, REPORT_TAB_ID };
