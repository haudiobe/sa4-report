/**
 * INTEGRATION TEST ONLY -- issue #6 / PR #7 (large-build checkpoints).
 *
 * Pushed ONLY into the bound script of a DISPOSABLE COPY of a report (or, with
 * --target audio-original, the Audio SWG SA4#138 report itself), next to
 * the candidate bundle, by tools/integration-test/push-to-disposable-copy.mjs.
 * It is not part of any release bundle (tools/template-release.js never
 * includes it) and must never be pushed to the master template or a report
 * in use.
 *
 * Every function here only READS the document (one Docs API read each). The
 * only thing written is a snapshot of that read into this copy's own
 * document properties (keys IT6_*), so that two runs can be compared.
 *
 * Run from the Apps Script editor of the copy (issue #11: it11Inspect(), it11Snapshot(), it11CompareLast() at the end):
 *   it6Inspect()        what the report holds now (logged and returned)
 *   it6Snapshot()       it6Inspect() and keep the result as the next snapshot
 *   it6CompareLast()    the last two snapshots side by side
 *   it6ClearSnapshots() forget all snapshots
 *
 * Reuses the Status reading of Code.js (buildStatusDocsIndex_,
 * statusDocsCellText_, statusDocsOtherContent_, planStatusDropdownRepair_),
 * so it reads a report exactly as the build and the repair do.
 */

// Documents this file refuses to look at, even read-only: the master template and the original Audio report
// (its issue #6 test is over; from issue #9 on, only disposable copies). Where this file may go at all is decided
// by push-to-disposable-copy.mjs.
var IT6_REFUSED_DOCUMENT_IDS_ = [
  '11XNUHU3TpAr8BaS7gyFQlwiKEzQFvnvzvGu4ms8lo2s', // SA4 Report Template (master)
  '1zawfIcbkzCJG-h2HYtnbQSEsm6MCd9OGhYzMgAFYENs'  // Audio SWG Minutes SA4#138 (original)
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

// =========================================================
// Issue #9 (T-2026.10.24): the meeting time zones in this Apps Script runtime. Read-only.
// =========================================================

/** Pure: the lines of it9TimeZoneCheck() from the functions of the runtime; `facts` holds what was read. */
function it9TimeZoneLines_(facts) {
  const edmonton = 'America/Edmonton';
  const calgary = Date.UTC(2026, 10, 16, 15, 0, 0);
  const change = Date.UTC(2026, 10, 1, 8, 30, 0);
  const one = { id: 'it9a', label: '', date: '2026-11-16', start: '09:00', end: '17:30' };
  const many = { id: 'it9b', label: '', date: '2026-10-30', endDate: '2026-11-02', start: '09:00', end: '17:00' };
  return [
    '[IT9] release: ' + facts.release,
    '[IT9] stored: ' + JSON.stringify(facts.stored),
    '[IT9] sessions zone of this report: ' + (facts.sessionsZone || '(none: the script zone ' + facts.scriptZone + ')'),
    '[IT9] Apps Script tz data, America/Edmonton at 2026-11-16 15:00 UTC: ' + facts.rawCalgary + '; at 2026-11-01 08:30 UTC: ' + facts.rawChange,
    '[IT9] report rule, 2026-11-16 15:00 UTC: offset ' + meetingTimeZoneOffsetMinutes_(edmonton, calgary) + ' min, written ' + meetingTimeZoneAbbreviationAt_(edmonton, calgary),
    '[IT9] 09:00 on 2026-11-16 in Calgary is ' + new Date(meetingWallClockToInstant_(edmonton, '2026-11-16', '09:00')).toISOString(),
    '[IT9] opening sentence: ' + buildMeetingOpeningSentence_('November 16, 2026', '09:00', edmonton),
    '[IT9] session, one day: ' + adhocSessionWhenText_(one, edmonton),
    '[IT9] session, over 1 November: ' + adhocSessionWhenText_(many, edmonton),
    '[IT9] uploads (UTC -> Calgary): 2026-11-16T17:59:00 -> ' + convertUtcWallClockToReportZone_('2026-11-16T17:59:00', edmonton) +
      ', 2026-11-16T18:01:00 -> ' + convertUtcWallClockToReportZone_('2026-11-16T18:01:00', edmonton) +
      ', 2026-11-01T08:30:00 -> ' + convertUtcWallClockToReportZone_('2026-11-01T08:30:00', edmonton) + ' (' + meetingTimeZoneAbbreviationAt_(edmonton, change) + ')',
    '[IT9] Berlin: ' + meetingTimeZoneDisplay_('Europe/Berlin', '2026-08-24', '09:00') + ' on 24 August, ' + meetingTimeZoneDisplay_('Europe/Berlin', '2026-01-20', '10:00') + ' on 20 January; ' +
      '25 October 00:30/01:30 UTC -> ' + convertUtcWallClockToReportZone_('2026-10-25T00:30:00', 'Europe/Berlin') + ' / ' + convertUtcWallClockToReportZone_('2026-10-25T01:30:00', 'Europe/Berlin'),
    '[IT9] Configure Meeting label: ' + meetingTimeZoneOptionLabel_(edmonton),
    '[IT9] report text naming a zone: ' + JSON.stringify(facts.paragraphs)
  ];
}

/** Editor: the meeting time zones as this runtime computes them, and the report's text that names a zone. Writes nothing. */
function it9TimeZoneCheck() {
  it6AssertDisposable_();
  const props = PropertiesService.getDocumentProperties();
  const stored = {};
  ['MEETING_ID', 'MEETING_TYPE', 'MEETING_DATE', 'MEETING_START_TIME', 'MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'].forEach(function (k) { stored[k] = props.getProperty(k); });
  const paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs()
    .map(function (p) { return p.getText(); })
    .filter(function (t) { return /opens the session|\b(ABT|MST|MDT|CET|CEST|UTC[+-]\d)/.test(t); })
    .slice(0, 20);
  const lines = it9TimeZoneLines_({
    release: typeof SA4_RELEASE_ !== 'undefined' ? SA4_RELEASE_.releaseId + ' / ' + SA4_RELEASE_.codeVersion + ' / ' + SA4_RELEASE_.gitCommit : '(no Release.js)',
    stored: stored,
    sessionsZone: meetingSessionTimeZone_(),
    scriptZone: Session.getScriptTimeZone(),
    rawCalgary: Utilities.formatDate(new Date(Date.UTC(2026, 10, 16, 15, 0, 0)), 'America/Edmonton', 'Z zzz'),
    rawChange: Utilities.formatDate(new Date(Date.UTC(2026, 10, 1, 8, 30, 0)), 'America/Edmonton', 'Z zzz'),
    paragraphs: paragraphs
  });
  lines.forEach(function (l) { Logger.log(l); });
  return lines;
}

// =========================================================
// Issue #11 (T-2026.10.25): session times in a main report. Read-only, except the snapshots (keys IT11_*).
// =========================================================

var IT11_SNAPSHOT_PREFIX_ = 'IT11_SNAP_';
var IT11_SNAPSHOT_COUNT_KEY_ = 'IT11_SNAP_COUNT';
// What a session-time save must never change.
var IT11_CONFIG_KEYS_ = ['MEETING_TYPE', 'REPORT_SUFFIX', 'TDOC_LIST_URL', 'AGENDA_TDOC', 'AGENDA_ITEM_PREFIX', 'AGENDA_SOURCE_DOC_ID', 'MEETING_ID', 'MEETING_FOLDER', 'MEETING_NUMBER',
  'FTP_BASE', 'MEETING_DATE', 'MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'];

/**
 * Pure: the summary of a report for issue #11. `facts`: { config, selector, sessions, elements }, where
 * `elements` are the body's children in order -- { type: 'PARAGRAPH', text, heading } or { type: 'TABLE', rows }.
 * The Session administration section (its heading and the "A01: ..." lines below it) is read apart; everything
 * else is digested: each TDoc table on its own (its minutes included), the rest as one text.
 */
function it11Summarize_(facts) {
  const elements = facts.elements || [];
  let sectionAt = -1;
  for (let i = 0; i < elements.length; i++) {
    if (elements[i].type === 'PARAGRAPH' && elements[i].heading !== 'NORMAL' && String(elements[i].text).trim() === 'Session administration') { sectionAt = i; break; }
  }
  let sectionEnd = sectionAt;
  if (sectionAt !== -1) {
    sectionEnd = sectionAt + 1;
    while (sectionEnd < elements.length && elements[sectionEnd].type === 'PARAGRAPH' && elements[sectionEnd].heading === 'NORMAL' && /^A\d{2,}: /.test(elements[sectionEnd].text)) sectionEnd++;
  }
  const section = sectionAt === -1 ? null : elements.slice(sectionAt, sectionEnd).map(function (e) { return e.text; });
  const tdocTables = {};
  const order = [];
  const other = [];
  let sections = 0;
  elements.forEach(function (e, i) {
    if (e.type === 'PARAGRAPH' && e.heading !== 'NORMAL' && String(e.text).trim() === 'Session administration') sections++;
    if (i >= sectionAt && i < sectionEnd && sectionAt !== -1) return;
    if (e.type === 'TABLE' && e.rows.length && e.rows[0].length === 2 && e.rows[0][0] === 'TDoc') {
      const n = String(e.rows[0][1]).trim();
      order.push(n);
      tdocTables[n] = it6Digest_(JSON.stringify(e.rows));
      return;
    }
    other.push(e.type === 'TABLE' ? 'T' + JSON.stringify(e.rows) : 'P' + e.heading + '|' + e.text);
  });
  return {
    ok: true, config: facts.config, selector: facts.selector, sessions: facts.sessions, section: section, sectionHeadings: sections,
    sectionAt: sectionAt, before: sectionAt > 0 ? elements[sectionAt - 1].text : null, after: sectionEnd < elements.length && sectionAt !== -1 ? (elements[sectionEnd].text || '[table]') : null,
    tdocTables: order.length, tdocOrder: it6Digest_(order.join(',')), tdocs: tdocTables, otherDigest: it6Digest_(other.join('\n')), otherLines: other.length
  };
}

/** Pure: what changed from snapshot `a` to snapshot `b` (issue #11). */
function it11Diff_(a, b) {
  const out = { configChanged: [], selectorUnchanged: JSON.stringify(a.selector) === JSON.stringify(b.selector), tdocsAdded: [], tdocsRemoved: [], tdocTablesChanged: [],
    tdocOrderUnchanged: a.tdocOrder === b.tdocOrder, otherTextUnchanged: a.otherDigest === b.otherDigest, sectionBefore: a.section, sectionAfter: b.section, sessionsChanged: a.sessions !== b.sessions };
  Object.keys(Object.assign({}, a.config, b.config)).forEach(function (k) { if ((a.config || {})[k] !== (b.config || {})[k]) out.configChanged.push(k); });
  Object.keys(b.tdocs).forEach(function (n) { if (!Object.prototype.hasOwnProperty.call(a.tdocs, n)) out.tdocsAdded.push(n); });
  Object.keys(a.tdocs).forEach(function (n) {
    if (!Object.prototype.hasOwnProperty.call(b.tdocs, n)) out.tdocsRemoved.push(n);
    else if (a.tdocs[n] !== b.tdocs[n]) out.tdocTablesChanged.push(n);
  });
  return out;
}

/** The facts of the open copy: its configuration, its agenda selector, its stored sessions and its body. */
function it11Read_() {
  it6AssertDisposable_();
  const props = PropertiesService.getDocumentProperties();
  const config = {};
  IT11_CONFIG_KEYS_.forEach(function (k) { config[k] = props.getProperty(k); });
  let selector = null;
  try { selector = getMeetingContext_().report.agendaSelector; } catch (e) { selector = 'unreadable: ' + e.message; }
  const body = DocumentApp.getActiveDocument().getBody();
  const elements = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const c = body.getChild(i);
    if (c.getType() === DocumentApp.ElementType.TABLE) {
      const t = c.asTable();
      const rows = [];
      for (let r = 0; r < t.getNumRows(); r++) { const row = []; for (let k = 0; k < t.getRow(r).getNumCells(); k++) row.push(t.getRow(r).getCell(k).getText()); rows.push(row); }
      elements.push({ type: 'TABLE', rows: rows });
    } else if (c.getType() === DocumentApp.ElementType.PARAGRAPH) {
      elements.push({ type: 'PARAGRAPH', text: c.asParagraph().getText(), heading: String(c.asParagraph().getHeading()) });
    } else {
      elements.push({ type: 'PARAGRAPH', text: '[' + c.getType() + ']' + (c.getText ? c.getText() : ''), heading: 'NORMAL' });
    }
  }
  return it11Summarize_({ config: config, selector: selector, sessions: props.getProperty('ADHOC_SESSIONS'), elements: elements });
}

/** Editor: what the copy holds now for issue #11 (logged and returned). Writes nothing. */
function it11Inspect() {
  const s = it11Read_();
  const shown = Object.assign({}, s);
  delete shown.tdocs;
  Logger.log('[IT11] ' + JSON.stringify(shown));
  return s;
}

/** Editor: it11Inspect(), kept as the next snapshot in this copy's own document properties. */
function it11Snapshot() {
  const s = it11Inspect();
  const props = PropertiesService.getDocumentProperties();
  const n = Number(props.getProperty(IT11_SNAPSHOT_COUNT_KEY_) || '0') + 1;
  s.takenAt = new Date().toISOString();
  props.setProperty(IT11_SNAPSHOT_PREFIX_ + n, JSON.stringify(s));
  props.setProperty(IT11_SNAPSHOT_COUNT_KEY_, String(n));
  Logger.log('[IT11] kept as snapshot #' + n + ' (' + s.takenAt + ')');
  return s;
}

/** Editor: the last two snapshots compared. */
function it11CompareLast() {
  it6AssertDisposable_();
  const props = PropertiesService.getDocumentProperties();
  const n = Number(props.getProperty(IT11_SNAPSHOT_COUNT_KEY_) || '0');
  if (n < 2) { Logger.log('[IT11] fewer than two snapshots'); return null; }
  const a = JSON.parse(props.getProperty(IT11_SNAPSHOT_PREFIX_ + (n - 1)));
  const b = JSON.parse(props.getProperty(IT11_SNAPSHOT_PREFIX_ + n));
  const d = it11Diff_(a, b);
  Logger.log('[IT11] snapshot #' + (n - 1) + ' (' + a.takenAt + ') -> #' + n + ' (' + b.takenAt + '): ' + JSON.stringify(d));
  return d;
}
