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
 *     nothing else. Requests are applied in order, each against the indexes
 *     AS THEY ARE AFTER the requests before it. A batch is applied as a
 *     whole or not at all.
 *   - A read sees the SAVED document: while the document has unsaved
 *     DocumentApp changes, Documents.get() describes it as it was before
 *     the first of them -- a table added in this execution is not there.
 *   - A write is refused while the document has unsaved DocumentApp changes:
 *     the API works on the saved document, so a write before the save would
 *     be a write to other positions than the code believes.
 *
 * `calls` lists every call; `fail` lets a test make calls fail.
 */

const REPORT_TAB_ID = 't.0';

function makeFakeDocsApi(options) {
  const body = options.body;
  const reportId = options.reportId;
  const sources = {};
  const definitions = {};                    // of the report: id -> { dropdownDefinitionId, dropdownDefinitionProperties }
  const counters = { definition: 0, dropdown: 0 };
  const api = { calls: [], fail: { get: null, batchUpdate: null }, definitions: definitions, tabId: REPORT_TAB_ID };
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const refuse = (message) => { const e = new Error(message); e.fakeDocsApi = true; return e; };

  const optionOf = (dropdown) => {
    const def = definitions[dropdown.definitionId];
    const options = def ? def.dropdownDefinitionProperties.options : [];
    return options.filter((o) => o.optionId === dropdown.optionId)[0] || null;
  };

  /** The report as the API describes it, and where each table cell is: [{ cell, start, textStart, textEnd }]. */
  function render() {
    const content = [];
    const cells = [];
    let i = 1;
    body._children.forEach((child) => {
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
            if (text && !dropdown) {
              elements.push({ startIndex: i, endIndex: i + text.length + 1, textRun: { content: text + '\n' } });
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
            cells.push({ cell: cell, start: start, textStart: start, textEnd: start + text.length });
            row.tableCells.push({ content: [{ startIndex: start, endIndex: i, paragraph: { elements: elements } }] });
          }
          element.table.tableRows.push(row);
        }
        i += 1;
        element.endIndex = i;
        content.push(element);
      } else {
        const text = child.getText();
        content.push({ startIndex: i, endIndex: i + text.length + 1, paragraph: { elements: [{ startIndex: i, endIndex: i + text.length + 1, textRun: { content: text + '\n' } }] } });
        i += text.length + 1;
      }
    });
    return { content: content, cells: cells };
  }

  const allCells = () => render().cells.map((x) => x.cell);

  // The saved document: kept the moment before DocumentApp makes its first unsaved change.
  let saved = null;
  body._state.onFirstTouch = () => { saved = clone({ definitions: definitions, content: render().content }); };
  const view = () => (body._state.dirty && saved ? saved : { definitions: definitions, content: render().content });

  function applyRequest(request, n) {
    const kind = Object.keys(request)[0];
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
      const hit = render().cells.filter((x) => range.startIndex >= x.textStart && range.endIndex <= x.textEnd && range.startIndex < range.endIndex)[0];
      if (!hit) throw bad('The range ' + range.startIndex + '..' + range.endIndex + ' is not text inside one table cell.');
      hit.cell._t = hit.cell._t.slice(0, range.startIndex - hit.textStart) + hit.cell._t.slice(range.endIndex - hit.textStart);
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
      const hit = render().cells.filter((x) => x.start === location.index)[0];
      if (!hit) throw bad('Index ' + location.index + ' is not the start of a table cell.');
      if (hit.cell._dropdown) throw bad('The cell already has a dropdown.');
      hit.cell._dropdown = { dropdownId: 'kix.dd' + (++counters.dropdown), definitionId: r.dropdownDefinitionId, optionId: optionId };
      return { insertDropdown: { dropdown: { dropdownId: hit.cell._dropdown.dropdownId } } };
    }

    if (kind === 'updateDropdownProperties') {
      checkTab(r.tabId);
      const cell = allCells().filter((c) => c._dropdown && c._dropdown.dropdownId === r.dropdownId)[0];
      if (!cell) throw bad('The dropdown ' + r.dropdownId + ' does not exist.');
      if (!/(^|,)\s*(selectedOptionId|\*)\s*(,|$)/.test(String(r.fields || ''))) throw bad('At least one field must be specified.');
      const def = definitions[cell._dropdown.definitionId];
      const optionId = r.dropdownProperties && r.dropdownProperties.selectedOptionId;
      if (!def || !def.dropdownDefinitionProperties.options.some((o) => o.optionId === optionId)) throw bad('The option ' + optionId + ' is not an option of the dropdown definition.');
      cell._dropdown.optionId = optionId;
      return {};
    }

    throw bad('This request is not supported by the fake Docs API.');
  }

  api.Docs = { Documents: {
    get: (id, opts) => {
      api.calls.push({ call: 'get', id: id, fields: (opts && opts.fields) || null, includeTabsContent: !!(opts && opts.includeTabsContent) });
      if (api.fail.get) { const message = api.fail.get(id, opts, api.calls.length); if (message) throw refuse(message); }
      if (id === reportId) {
        const v = view();
        return clone({ revisionId: 'rev' + api.calls.length, tabs: [{ tabProperties: { tabId: REPORT_TAB_ID }, documentTab: { dropdownDefinitions: v.definitions, body: { content: v.content } } }] });
      }
      if (sources[id]) return clone(sources[id]);
      throw refuse('Requested entity was not found.');
    },
    batchUpdate: (resource, id) => {
      const requests = (resource && resource.requests) || [];
      api.calls.push({ call: 'batchUpdate', id: id, requests: clone(requests), dirty: body._state.dirty });
      if (id !== reportId) throw refuse('The fake Docs API writes to the report only: ' + id);
      if (api.fail.batchUpdate) { const message = api.fail.batchUpdate(requests, api.calls.length); if (message) throw refuse(message); }
      if (body._state.dirty) throw refuse('FAKE DOCS API: the document has unsaved DocumentApp changes; a write now would go to other positions than the caller read.');
      // All or nothing.
      const before = { cells: allCells().map((c) => ({ c: c, t: c._t, d: c._dropdown ? Object.assign({}, c._dropdown) : null })), definitions: clone(definitions), counters: Object.assign({}, counters) };
      try {
        return { replies: requests.map((request, n) => applyRequest(request, n)) };
      } catch (e) {
        before.cells.forEach((x) => { x.c._t = x.t; x.c._dropdown = x.d; });
        Object.keys(definitions).forEach((k) => delete definitions[k]);
        Object.assign(definitions, before.definitions);
        Object.assign(counters, before.counters);
        throw e;
      }
    }
  } };

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
    const tables = body._children.filter((c) => c.getType() === 'TABLE' && c.getRow(0).getNumCells() === 2 && c.getRow(0).getCell(0).getText() === 'TDoc' && c.getRow(0).getCell(1).getText() === number);
    if (!tables.length) return null;
    for (let r = 0; r < tables[0].getNumRows(); r++) if (tables[0].getRow(r).getCell(0).getText() === 'Status') return tables[0].getRow(r).getCell(1);
    return null;
  };
  /** What a reader of the document sees in the Status cell of `number`: 'text:agreed', 'dropdown:agreed', 'text:agreed+dropdown:noted', 'none'. */
  api.shown = (number) => {
    const cell = api.statusCell(number);
    if (!cell) return 'none';
    const parts = [];
    if (cell._t) parts.push('text:' + cell._t);
    if (cell._dropdown) { const option = optionOf(cell._dropdown); parts.push('dropdown:' + (option ? option.displayValue : '?')); }
    return parts.join('+') || 'empty';
  };
  /** A user picks another option in the dropdown of `number` (this is not a DocumentApp change: nothing becomes unsaved). */
  api.pick = (number, displayValue) => {
    const cell = api.statusCell(number);
    const def = definitions[cell._dropdown.definitionId];
    cell._dropdown.optionId = def.dropdownDefinitionProperties.options.filter((o) => o.displayValue === displayValue)[0].optionId;
  };
  api.count = (call, id) => api.calls.filter((c) => c.call === call && (id === undefined || c.id === id)).length;
  return api;
}

module.exports = { makeFakeDocsApi, REPORT_TAB_ID };
