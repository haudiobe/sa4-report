/**
 * INTEGRATION TEST ONLY -- issue #6 / PR #7 (large-build checkpoints).
 *
 * Pushed ONLY into the bound script of a DISPOSABLE COPY of a report, next to
 * the candidate bundle, by tools/integration-test/push-to-disposable-copy.mjs.
 * It is not part of any release bundle (tools/template-release.js never
 * includes it) and must never be pushed to the master template or a report
 * in use.
 *
 * Every function here only READS the document (one Docs API read each). The
 * only thing written is a snapshot of that read into this copy's own
 * document properties (keys IT6_*), so that two runs can be compared.
 *
 * Run from the Apps Script editor of the copy:
 *   it6Inspect()        what the report holds now (logged and returned)
 *   it6Snapshot()       it6Inspect() and keep the result as the next snapshot
 *   it6CompareLast()    the last two snapshots side by side
 *   it6ClearSnapshots() forget all snapshots
 *
 * Reuses the Status reading of Code.js (buildStatusDocsIndex_,
 * statusDocsCellText_, statusDocsOtherContent_, planStatusDropdownRepair_),
 * so it reads a report exactly as the build and the repair do.
 */

// Documents this file refuses to look at, even read-only: the original Audio
// report and the master template (it exists for disposable copies only).
var IT6_REFUSED_DOCUMENT_IDS_ = [
  '1zawfIcbkzCJG-h2HYtnbQSEsm6MCd9OGhYzMgAFYENs', // Audio SWG Minutes SA4#138 (original)
  '11XNUHU3TpAr8BaS7gyFQlwiKEzQFvnvzvGu4ms8lo2s'  // SA4 Report Template (master)
];
var IT6_SNAPSHOT_PREFIX_ = 'IT6_SNAP_';
var IT6_SNAPSHOT_COUNT_KEY_ = 'IT6_SNAP_COUNT';

/** Pure: does a TDoc table's row read "Abstract" in its first cell? */
function it6IsAbstractRow_(cells) {
  return cells.length >= 2 && statusDocsCellText_(cells[0]).trim().toLowerCase() === 'abstract';
}

/**
 * Pure: a summary of one Docs API read of a report.
 * `document`: Docs.Documents.get(...) with STATUS_DOCS_MIGRATION_FIELDS_;
 * `marked`: the STATUS_DROPDOWNS document property; `templateOptions`: the
 * template's dropdown options, used when the report has no definition.
 */
function it6Summarize_(document, marked, templateOptions) {
  const index = buildStatusDocsIndex_(document);
  if (!index) return { ok: false, error: 'The document could not be read.' };
  const tab = document.tabs[0].documentTab;
  const content = (tab.body && tab.body.content) || [];
  let tables = 0;
  let tdocTableCount = 0;
  const withAbstract = {};
  content.forEach(function (element) {
    if (!element.table) return;
    tables++;
    const rows = element.table.tableRows || [];
    const head = (rows[0] && rows[0].tableCells) || [];
    if (head.length < 2 || statusDocsCellText_(head[0]).trim() !== 'TDoc') return;
    tdocTableCount++;
    const number = statusDocsCellText_(head[1]).trim();
    if (rows.some(function (row) { return it6IsAbstractRow_(row.tableCells || []); })) withAbstract[number] = true;
  });

  const definition = index.definition;
  const options = definition ? definition.options : (templateOptions || []);
  const optionName = {};
  options.forEach(function (o) { if (o && o.optionId) optionName[o.optionId] = o.displayValue; });
  const plan = planStatusDropdownRepair_(index, options);

  const statuses = {};
  let dropdowns = 0;
  let foreignDropdowns = 0;
  let texts = 0;
  let empty = 0;
  Object.keys(index.tables).sort().forEach(function (number) {
    const e = index.tables[number];
    if (e.kind === 'dropdown') {
      dropdowns++;
      if (!definition || e.definitionId !== definition.id) foreignDropdowns++;
      statuses[number] = 'D:' + (optionName[e.selectedOptionId] || e.displayValue || '?');
    } else if (String(e.text || '').trim() === '') {
      empty++;
      statuses[number] = 'E:';
    } else {
      texts++;
      statuses[number] = 'T:' + String(e.text).trim().slice(0, 24);
    }
  });

  const other = statusDocsOtherContent_(document);
  return {
    ok: true,
    revisionId: index.revisionId,
    tables: tables,
    tdocTableElements: tdocTableCount,
    tdocNumbers: Object.keys(index.tables).length,
    duplicates: Object.keys(index.duplicates).sort(),
    abstracts: Object.keys(withAbstract).length,
    definition: definition ? 'present (' + definition.options.length + ' options)' : 'none',
    dropdowns: dropdowns,
    dropdownsOfAnotherDefinition: foreignDropdowns,
    textStatuses: texts,
    emptyStatuses: empty,
    stillToConvert: Object.keys(plan.convert).length,
    leftAsTextNoOption: plan.unmapped.length,
    marked: marked === '1',
    otherContentLines: other.length,
    otherContentDigest: it6Digest_(other.join('\n')),
    statuses: statuses
  };
}

/** Pure: a short, stable digest of a text (FNV-1a, 32 bit, hex): enough to tell "unchanged" from "changed". */
function it6Digest_(text) {
  let h = 0x811c9dc5;
  const s = String(text);
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return ('00000000' + h.toString(16)).slice(-8);
}

/** Pure: what changed from snapshot `a` to snapshot `b`. */
function it6Diff_(a, b) {
  const out = { tablesAdded: [], tablesRemoved: [], statusChanged: [], counts: {}, otherContentUnchanged: a.otherContentDigest === b.otherContentDigest };
  Object.keys(b.statuses).forEach(function (n) { if (!Object.prototype.hasOwnProperty.call(a.statuses, n)) out.tablesAdded.push(n); });
  Object.keys(a.statuses).forEach(function (n) {
    if (!Object.prototype.hasOwnProperty.call(b.statuses, n)) out.tablesRemoved.push(n);
    else if (a.statuses[n] !== b.statuses[n]) out.statusChanged.push(n + ': ' + a.statuses[n] + ' -> ' + b.statuses[n]);
  });
  ['tables', 'tdocTableElements', 'tdocNumbers', 'abstracts', 'dropdowns', 'textStatuses', 'emptyStatuses', 'stillToConvert', 'otherContentLines'].forEach(function (k) {
    if (a[k] !== b[k]) out.counts[k] = a[k] + ' -> ' + b[k];
  });
  if (a.duplicates.join(',') !== b.duplicates.join(',')) out.counts.duplicates = a.duplicates.join(',') + ' -> ' + b.duplicates.join(',');
  return out;
}

/** Pure: one summary as log text (without the per-TDoc statuses). */
function it6Format_(s) {
  if (!s.ok) return '[IT6] ' + s.error;
  const copy = Object.assign({}, s);
  delete copy.statuses;
  return '[IT6] ' + JSON.stringify(copy);
}

function it6AssertDisposable_() {
  const id = DocumentApp.getActiveDocument().getId();
  if (IT6_REFUSED_DOCUMENT_IDS_.indexOf(id) !== -1) {
    throw new Error('IntegrationCheck.js is for disposable copies only; this is a protected document (' + id + '). Nothing was read.');
  }
  return id;
}

function it6Read_() {
  const id = it6AssertDisposable_();
  const document = statusDocsGet_(id, STATUS_DOCS_MIGRATION_FIELDS_);
  let templateOptions = null;
  try {
    const source = readStatusDropdownSourceDefinition_();
    templateOptions = source ? source.options : null;
  } catch (e) {
    // Without the template's options a report without a definition just shows "none".
  }
  return it6Summarize_(document, PropertiesService.getDocumentProperties().getProperty(STATUS_DROPDOWN_ENABLED_KEY_), templateOptions);
}

/** Editor: what the report holds now. */
function it6Inspect() {
  const s = it6Read_();
  Logger.log(it6Format_(s));
  if (s.ok) Logger.log('[IT6] statuses: ' + JSON.stringify(s.statuses));
  return s;
}

/** Editor: it6Inspect(), kept as the next snapshot. */
function it6Snapshot() {
  const s = it6Inspect();
  if (!s.ok) return s;
  const props = PropertiesService.getDocumentProperties();
  const n = Number(props.getProperty(IT6_SNAPSHOT_COUNT_KEY_) || '0') + 1;
  s.takenAt = new Date().toISOString();
  props.setProperty(IT6_SNAPSHOT_PREFIX_ + n, JSON.stringify(s));
  props.setProperty(IT6_SNAPSHOT_COUNT_KEY_, String(n));
  Logger.log('[IT6] kept as snapshot #' + n + ' (' + s.takenAt + ')');
  return s;
}

/** Editor: the last two snapshots compared. */
function it6CompareLast() {
  it6AssertDisposable_();
  const props = PropertiesService.getDocumentProperties();
  const n = Number(props.getProperty(IT6_SNAPSHOT_COUNT_KEY_) || '0');
  if (n < 2) { Logger.log('[IT6] fewer than two snapshots'); return null; }
  const a = JSON.parse(props.getProperty(IT6_SNAPSHOT_PREFIX_ + (n - 1)));
  const b = JSON.parse(props.getProperty(IT6_SNAPSHOT_PREFIX_ + n));
  const d = it6Diff_(a, b);
  Logger.log('[IT6] snapshot #' + (n - 1) + ' (' + a.takenAt + ') -> #' + n + ' (' + b.takenAt + '): ' + JSON.stringify(d));
  return d;
}

/** Editor: forget every snapshot. */
function it6ClearSnapshots() {
  it6AssertDisposable_();
  const props = PropertiesService.getDocumentProperties();
  props.getKeys().forEach(function (k) { if (k.indexOf('IT6_') === 0) props.deleteProperty(k); });
  Logger.log('[IT6] snapshots cleared');
}
