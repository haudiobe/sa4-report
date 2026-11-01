/**
 * The report-local scope of the Disposition Summary (after T-2026.10.17) --
 * docs/DISPOSITION_SUMMARY.md § Which TDocs.
 *
 * The summary is derived from the report: the TDoc tables whose number the
 * report has registered (its "Registration of Documents" table), less those its
 * Document Reallocations remove (reportTdocScope_()). No TDoc list is read for
 * it, so a Portal list that cannot be read, or is not the meeting's, neither
 * stops it nor gets into it. Close Report writes the same summary and then
 * checks the TDoc list (3GU) on top: without that check it is never ready.
 * The update keeps its TDoc list identity (T-2026.10.17): a wrong list is
 * refused before any change.
 *
 * Every scenario runs on the REAL build, update and status dropdowns of
 * tests/helpers/template-report.js -- once with this Code.js, once with the
 * Code.js of T-2026.10.17 (from git) -- so the tests show which of them the
 * released code fails and that this code passes them all.
 *
 *   S.  the scenarios (both implementations)
 *   O.  the released code fails exactly the scenarios this change is for
 *   L.  what changed in Code.js: against T-2026.10.17, only the summary's scope and Close Report
 *
 * Run: node tests/disposition-report-local-scope.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');
const { codeWithoutCloseReportPageFallback, codeWithoutReportLocalScope, REPORT_LOCAL_SCOPE_NEW_FUNCTIONS, REPORT_LOCAL_SCOPE_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

let failures = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) {
    console.log('  ok   ' + name);
  } else {
    failures++;
    console.log('  FAIL ' + name + '\n         expected ' + e + '\n         actual   ' + a);
  }
}

const ROOT = path.join(__dirname, '..');
const RELEASED_TAG = 'template-release/T-2026.10.17';
let released = null;
try {
  released = execFileSync('git', ['show', RELEASED_TAG + ':Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8');
} catch (e) {
  released = null;
}

const SUMMARY = 'Disposition Summary';
const NOT_UPDATED = 'Disposition Summary NOT updated';
const TITLES = ['Agreed, Endorsed and Approved Documents', 'Other Disposed Documents', 'Documents Not Yet Disposed'];
const AGENDA = 'S4aA260090';
const SECOND_AGENDA = 'S4aA260149';
const id = (n) => 'S4aA2691' + String(n).padStart(2, '0');
const t = (n, agenda, status, extra) => Object.assign({ id: id(n), agenda: agenda, status: status || 'available', uploaded: true }, extra || {});
const isHeading = (c) => c.getType() === 'PARAGRAPH' && c.getHeading() !== 'NORMAL';
const isSummaryTable = (c) => c.getType() === 'TABLE' && c.getNumRows() > 0 && c.getRow(0).getNumCells() === 4 && c.getRow(0).getCell(1).getText() === 'Title' && c.getRow(0).getCell(3).getText() === 'Status';
const cellLink = (cell) => cell.editAsText().getLinkUrl(0) || '';
/** The whole document, element by element: text, selected dropdown option, link. */
const snapshot = (body) => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE'
  ? ['T', Array.from({ length: c.getNumRows() }, (x, r) => Array.from({ length: c.getRow(r).getNumCells() }, (y, k) => [c.getRow(r).getCell(k).getText(), c.getRow(r).getCell(k)._dropdown ? c.getRow(r).getCell(k)._dropdown.optionId : null, cellLink(c.getRow(r).getCell(k))]))]
  : ['P', c.getHeading(), c.getText()])));
const props = (r) => JSON.stringify(Object.keys(r.docProps._store).sort().map((k) => [k, r.docProps._store[k]]));
const listFetches = (r, from) => r.fetches.slice(from).filter((f) => /GenerateDocumentList/.test(f)).length;
const lastAlert = (r) => r.ui.alerts[r.ui.alerts.length - 1];
/** The summary as it reads: [title line, ['<tdoc> <status>'] | 'None.'] per class; null without one. */
function summaryOf(body) {
  const at = body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY);
  if (at === -1) return null;
  const out = [];
  for (let i = at + 1; i < body._children.length && !isHeading(body._children[i]); i++) {
    const c = body._children[i];
    if (c.getType() === 'PARAGRAPH' && /\(\d+\)$/.test(c.getText())) out.push([c.getText()]);
    else if (c.getType() === 'PARAGRAPH' && c.getText() === 'None.' && out.length && out[out.length - 1].length === 1) out[out.length - 1].push('None.');
    else if (isSummaryTable(c)) out[out.length - 1].push(Array.from({ length: c.getNumRows() - 1 }, (x, r) => c.getRow(r + 1).getCell(0).getText() + ' ' + c.getRow(r + 1).getCell(3).getText()));
  }
  return out;
}
/** The document without the summary (its heading to the end of what it wrote). */
const outsideSummary = (r) => {
  const at = r.body._children.findIndex((c) => isHeading(c) && c.getText() === SUMMARY);
  const all = JSON.parse(snapshot(r.body));
  if (at === -1) return JSON.stringify(all);
  let end = at + 1;
  while (end < r.body._children.length && !isHeading(r.body._children[end]) && (isSummaryTable(r.body._children[end]) || r.body._children[end].getType() === 'PARAGRAPH' &&
    (r.body._children[end].getText() === '' || r.body._children[end].getText() === 'None.' || /^Written from the Status/.test(r.body._children[end].getText()) || /\(\d+\)$/.test(r.body._children[end].getText())))) end++;
  return JSON.stringify(all.slice(0, at).concat(all.slice(end)));
};

/** Close Report shows its result in an HTML dialog; it is recorded with the alerts. */
function withDialogs(r) {
  r.s.HtmlService = { createHtmlOutput: (html) => { const o = { setWidth: () => o, setHeight: () => o, getContent: () => html }; return o; } };
  r.ui.showModalDialog = (output, title) => { r.ui.alerts.push([title, r.logs[r.logs.length - 1], 'dialog', output.getContent()]); };
  return r;
}
/** A built template report (ad-hoc, agenda TDoc S4aA260090) with status dropdowns, on `code` (null: this Code.js). */
function built(code, tdocs, decisions, options) {
  const r = withDialogs(templateReport(Object.assign({ tdocs: tdocs, tableSeparators: true, endParagraph: true, code: code || undefined }, options || {})));
  const b = r.build();
  if (!b.ok) throw new Error('build failed: ' + (b.error || b.message));
  const u = r.update();
  if (!u.success) throw new Error('update failed: ' + u.error);
  decide(r, decisions || {});
  return r;
}
function decide(r, decisions) {
  r.user(() => Object.keys(decisions).forEach((n) => r.api.pick(id(n), decisions[n])));
}
const updateSummary = (r) => r.exec(() => r.s.updateDispositionSummary());
const closeReport = (r) => r.exec(() => r.s.closeReport());
/** The Portal does not answer for the TDoc list (HTTP 503); everything else as before. */
function portalUnavailable(r) {
  const fetch = r.s.UrlFetchApp.fetch;
  r.s.UrlFetchApp.fetch = (url, o) => {
    if (!/GenerateDocumentList/.test(String(url))) return fetch(url, o);
    r.fetches.push(String(url));
    return { getResponseCode: () => 503, getContentText: () => 'Service Unavailable' };
  };
}
/** The Portal returns the list of another meeting: no row for the agenda TDoc; some of the report's TDocs (other statuses), and others. */
function wrongList(r, overlap) {
  r.listAgenda = null;
  r.tdocs = (overlap || []).map(([n, agenda, status]) => t(n, agenda, status)).concat([t(41, '4.3'), t(42, '4.7', 'reserved', { uploaded: false })]);
}
const tablesOf = (r) => r.order();
/** MBS: an ad-hoc meeting whose agenda is its agenda.csv -- no agenda TDoc, so the TDoc list identity is the anchor. */
function withoutAgendaTdoc(r) {
  r.docProps.deleteProperty('AGENDA_TDOC');
  r.docProps.setProperty('AGENDA_CSV_URL', r.s.deriveAdhocAgendaCsvCandidate_(r.docProps.getProperty('FTP_BASE')).url);
  r.listAgenda = null;
}
/** The document without its empty paragraphs (Google Docs keeps one at the end, so a summary written again follows it). */
const content = (r) => JSON.stringify(JSON.parse(snapshot(r.body)).filter((e) => !(e[0] === 'P' && e[1] === 'NORMAL' && e[2] === '')));

// The meeting of most scenarios: 4.3: 01 agreed, 02 revised to 03 (03 agreed), 4.4: 04 noted, 05 available, 4.5: 06 postponed.
const MEETING = () => [t(1, '4.3'), t(2, '4.3', 'revised', { revisedTo: id(3) }), t(3, '4.3'), t(4, '4.4'), t(5, '4.4'), t(6, '4.5')];
const DECISIONS = { 1: 'agreed', 2: 'revised', 3: 'agreed', 4: 'noted', 6: 'postponed' };
const MEETING_SUMMARY = [[TITLES[0] + ' (2)', [id(1) + ' agreed', id(3) + ' agreed']], [TITLES[1] + ' (2)', [id(2) + ' revised', id(4) + ' noted']],
  [TITLES[2] + ' (2)', [id(5) + ' available', id(6) + ' postponed']]];
// All disposed: 01 agreed, 02 noted, 03 endorsed.
const CLEAN = () => [t(1, '4.3'), t(2, '4.3'), t(3, '4.4')];
const CLEAN_DECISIONS = { 1: 'agreed', 2: 'noted', 3: 'endorsed' };

/**
 * The scenarios. Each takes the Code.js to run (null: this one) and returns [[name, actual, expected], ...]; a scenario
 * passes when every one of its checks does.
 */
const SCENARIOS = [
  ['1. a normal report, the TDoc list of the meeting', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    const out = [['Update Disposition Summary: the TDocs of the report, in three tables', [lastAlert(r)[0], summaryOf(r.body)], [SUMMARY, MEETING_SUMMARY]]];
    const c = built(code, CLEAN(), CLEAN_DECISIONS);
    const res = closeReport(c);
    out.push(['Close Report on a report with every TDoc disposed: ready', [res.ready, lastAlert(c)[0]], [true, 'Report ready to close']]);
    return out;
  }],
  ['1a. Update Disposition Summary is report-local: it reads no TDoc list', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    const f = r.fetches.length;
    updateSummary(r);
    return [['no TDoc list fetched, the summary written', [listFetches(r, f), lastAlert(r)[0]], [0, SUMMARY]]];
  }],
  ['2. a normal report, the Portal unavailable', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    portalUnavailable(r);
    const before = outsideSummary(r);
    updateSummary(r);
    const out = [['Update Disposition Summary: written from the report, nothing else changed', [lastAlert(r)[0], summaryOf(r.body), outsideSummary(r) === before], [SUMMARY, MEETING_SUMMARY, true]]];
    const res = closeReport(r);
    out.push(['Close Report: NOT ready, external verification not completed, the summary the report\'s', [res.ready, res.externalVerified, res.summaryState, /^External verification could not be completed: .*HTTP 503/.test(res.problems.join('\n'))],
      [false, false, 'unchanged', true]]);
    out.push(['the revised TDoc is listed as not verified, never as complete', res.revisions.map((x) => [x.tdoc, x.kind]), [[id(2), 'revision-not-verified']]]);
    return out;
  }],
  ['3. a normal report, the Portal returns the list of another meeting', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    wrongList(r, []);
    const before = outsideSummary(r);
    updateSummary(r);
    const out = [['Update Disposition Summary: written from the report, the wrong list not used', [lastAlert(r)[0], summaryOf(r.body), outsideSummary(r) === before], [SUMMARY, MEETING_SUMMARY, true]]];
    const res = closeReport(r);
    out.push(['Close Report: NOT ready, the list refused by its identity, said so', [res.ready, res.externalVerified, /^External verification could not be completed: .*could not be shown to be the list of this meeting \(it has no row for S4aA260090/.test(res.problems.join('\n'))],
      [false, false, true]]);
    return out;
  }],
  ['4. a continuing report: two calls, two agenda TDocs, both calls in the summary', (code) => {
    // First call (agenda S4aA260090): 01 agreed, 02 parked. Second call (agenda S4aA260149, configured since): 11 noted, 12 available.
    const r = built(code, [t(1, '4.3'), t(2, '4.4')], { 1: 'agreed', 2: 'parked' });
    r.tdocs.push(t(11, '4.3'), t(12, '4.5'));
    r.docProps.setProperty('AGENDA_TDOC', SECOND_AGENDA);
    r.listAgenda = SECOND_AGENDA;
    const u = r.update();
    decide(r, { 11: 'noted' });
    const all = [[TITLES[0] + ' (1)', [id(1) + ' agreed']], [TITLES[1] + ' (1)', [id(11) + ' noted']], [TITLES[2] + ' (2)', [id(2) + ' parked', id(12) + ' available']]];
    updateSummary(r);
    const out = [['the second call registered by its update; the summary has both calls -- not only the latest agenda', [u.success, summaryOf(r.body)], [true, all]]];
    const res = closeReport(r);
    out.push(['Close Report checks both calls: the parked TDoc of the first call blocks too', [res.ready, res.externalVerified, res.unresolved.map((e) => e.tdoc)], [false, true, [id(2), id(12)]]]);
    portalUnavailable(r);
    decide(r, { 12: 'agreed' });
    updateSummary(r);
    out.push(['the Portal unavailable: the summary still the report\'s, both calls', [lastAlert(r)[0], summaryOf(r.body)[0]], [SUMMARY, [TITLES[0] + ' (2)', [id(1) + ' agreed', id(12) + ' agreed']]]]);
    return out;
  }],
  ['5. a wrong Portal list that overlaps the report', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    // Another meeting's list: 01 and 05 in it, with other statuses, and 41, 42 the report never had.
    wrongList(r, [[1, '4.3', 'withdrawn'], [5, '4.4', 'agreed']]);
    const doc = snapshot(r.body);
    const store = props(r);
    const u = r.update();
    const out = [['Update Report Now: refused before anything; the document and its properties identical', [u.success, u.notUpdated, snapshot(r.body) === doc, props(r) === store], [false, true, true, true]]];
    const statuses = () => [1, 2, 3, 4, 5, 6].map((n) => r.shown(id(n)));
    const statusesBefore = statuses();
    const before = outsideSummary(r);
    updateSummary(r);
    out.push(['Update Disposition Summary: the report\'s TDocs and statuses only -- not 41, 42, not the list\'s statuses', [lastAlert(r)[0], summaryOf(r.body)], [SUMMARY, MEETING_SUMMARY]]);
    out.push(['... and nothing but the summary changed: no table, no status, no property', [tablesOf(r), statuses(), outsideSummary(r) === before, props(r) === store], [[1, 2, 3, 4, 5, 6].map(id), statusesBefore, true, true]]);
    const res = closeReport(r);
    out.push(['Close Report: NOT ready, not verified', [res.ready, res.externalVerified], [false, false]]);
    return out;
  }],
  ['6. an MBS-style report without AGENDA_TDOC (the anchor)', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    withoutAgendaTdoc(r);
    const first = r.update();
    const anchor = JSON.parse(r.docProps.getProperty('TDOC_LIST_ANCHOR') || 'null');
    const out = [['the first update without an agenda TDoc records the anchor from the report', [first.success, anchor && anchor.tdoc], [true, id(1)]]];
    portalUnavailable(r);
    updateSummary(r);
    out.push(['the Portal unavailable: Update Disposition Summary written from the report', [lastAlert(r)[0], summaryOf(r.body)], [SUMMARY, MEETING_SUMMARY]]);
    const c1 = closeReport(r);
    out.push(['... Close Report NOT ready, not verified', [c1.ready, c1.externalVerified], [false, false]]);
    return out;
  }],
  ['6a. an MBS-style report without AGENDA_TDOC, a wrong list that overlaps it without the anchor', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    withoutAgendaTdoc(r);
    r.update();
    // Another meeting's list: 02 and 05 of the report (not 01, the anchor), and 41, 42.
    r.tdocs = [t(2, '4.3', 'withdrawn'), t(5, '4.4', 'agreed'), t(41, '4.3'), t(42, '4.7')];
    const doc = snapshot(r.body);
    const store = props(r);
    const u = r.update();
    const out = [['Update Report Now: refused (no row for the anchor), nothing changed', [u.success, u.notUpdated, /no row for S4aA269101, a TDoc of the list this meeting had before/.test(u.error), snapshot(r.body) === doc, props(r) === store],
      [false, true, true, true, true]]];
    updateSummary(r);
    out.push(['Update Disposition Summary: the report\'s TDocs only', [lastAlert(r)[0], summaryOf(r.body)], [SUMMARY, MEETING_SUMMARY]]);
    const res = closeReport(r);
    out.push(['Close Report: NOT ready, the list refused', [res.ready, res.externalVerified, /could not be shown to be the list of this meeting/.test(res.problems.join('\n'))], [false, false, true]]);
    return out;
  }],
  ['7. an existing summary is updated idempotently', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    const first = snapshot(r.body);
    updateSummary(r);
    updateSummary(r);
    const out = [['run again twice: "already up to date", the document identical', [lastAlert(r)[1].split('\n')[0], snapshot(r.body) === first], ['Disposition Summary was already up to date.', true]]];
    decide(r, { 5: 'agreed' });
    updateSummary(r);
    const second = snapshot(r.body);
    out.push(['a decision: the summary updated in place, one summary', [lastAlert(r)[1].split('\n')[0], summaryOf(r.body).map((c) => c[0]), r.body._children.filter((c) => isHeading(c) && c.getText() === SUMMARY).length],
      ['Disposition Summary updated.', [TITLES[0] + ' (3)', TITLES[1] + ' (2)', TITLES[2] + ' (1)'], 1]]);
    updateSummary(r);
    out.push(['and then unchanged again', snapshot(r.body) === second, true]);
    out.push(['an update keeps it as it is', [r.update().dispositionSummary, snapshot(r.body) === second], ['unchanged', true]]);
    return out;
  }],
  ['8. a deleted summary is written again', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    const written = content(r);
    const without = (() => {
      const s = r.exec(() => r.s.findDispositionSummarySection_(r.body));
      r.user(() => { for (let i = s.contentEnd - 1; i >= s.start; i--) r.body.removeChild(r.body._children[i]); });
      return summaryOf(r.body);
    })();
    updateSummary(r);
    const out = [['deleted by hand, then Update Disposition Summary: written again at the end, the same', [without, lastAlert(r)[1].split('\n')[0], summaryOf(r.body), content(r) === written],
      [null, 'Disposition Summary written at the end of the report.', MEETING_SUMMARY, true]]];
    r.user(() => { const table = r.body._children.filter(isSummaryTable)[1]; r.body.removeChild(table); });
    updateSummary(r);
    out.push(['one of its tables deleted by hand: written again whole', [lastAlert(r)[1].split('\n')[0], summaryOf(r.body)], ['Disposition Summary updated.', MEETING_SUMMARY]]);
    return out;
  }],
  ['9. text typed below the summary is kept', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    r.user(() => { r.body.appendParagraph('Annex A: notes of the chair').setHeading('HEADING1'); r.body.appendParagraph('Kept as written.'); r.body.appendParagraph('None.'); });
    const tail = () => r.body._children.filter((c) => c.getType() === 'PARAGRAPH' && c.getText()).map((c) => c.getText()).slice(-3);
    decide(r, { 5: 'noted', 6: 'agreed' });
    updateSummary(r);
    return [['a decision, then Update Disposition Summary: the summary updated, the text below it as it was', [lastAlert(r)[1].split('\n')[0], summaryOf(r.body).map((c) => c[0]), tail()],
      ['Disposition Summary updated.', [TITLES[0] + ' (3)', TITLES[1] + ' (3)', TITLES[2] + ' (0)'], ['Annex A: notes of the chair', 'Kept as written.', 'None.']]]];
  }],
  ['10. text typed into the summary: refused, nothing changed', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    r.user(() => { const table = r.body._children.filter(isSummaryTable)[0]; r.body.insertParagraph(r.body.getChildIndex(table) + 1, 'Typed into the summary by hand.'); });
    decide(r, { 5: 'agreed' });
    const doc = snapshot(r.body);
    updateSummary(r);
    const out = [['Update Disposition Summary: "NOT updated", the text named, the document identical', [lastAlert(r)[0], /Typed into the summary by hand/.test(lastAlert(r)[1]), snapshot(r.body) === doc], [NOT_UPDATED, true, true]]];
    const res = closeReport(r);
    out.push(['Close Report: NOT ready, the document identical', [res.ready, res.summaryState, snapshot(r.body) === doc], [false, 'refused', true]]);
    return out;
  }],
  ['11. Update Report Now still refuses a wrong Portal list, with zero mutation', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    wrongList(r, [[1, '4.3', 'withdrawn'], [4, '4.4', 'agreed']]);
    const doc = snapshot(r.body);
    const store = props(r);
    r.mark();
    const f = r.fetches.length;
    const u = r.update();
    return [['refused before anything: "not updated", the identity reason; document, properties, statuses identical; no other request',
      [u.success, u.notUpdated, /could not be shown to be the list of this meeting \(it has no row for S4aA260090/.test(u.error), snapshot(r.body) === doc, props(r) === store, listFetches(r, f), r.requests.length],
      [false, true, true, true, true, 1, 0]]];
  }],
  ['12. Close Report writes the summary but is not ready when the external check fails', (code) => {
    const r = built(code, CLEAN(), CLEAN_DECISIONS);
    portalUnavailable(r);
    const before = outsideSummary(r);
    const res = closeReport(r);
    const out = [['every TDoc disposed, the Portal unavailable: the summary created from the report, NOT ready, said so in the dialog',
      [res.ready, res.summaryState, summaryOf(r.body).map((c) => c[0]), outsideSummary(r) === before, lastAlert(r)[0], /External verification \(TDoc list of this meeting, 3GU\): NOT completed\./.test(lastAlert(r)[3])],
      [false, 'created', [TITLES[0] + ' (2)', TITLES[1] + ' (1)', TITLES[2] + ' (0)'], true, 'Report NOT ready to close', true]]];
    return out;
  }],
  ['12a. ... and ready once the external check passes', (code) => {
    const r = built(code, CLEAN(), CLEAN_DECISIONS);
    wrongList(r, []);
    const bad = closeReport(r);
    r.listAgenda = AGENDA;
    r.tdocs = CLEAN();
    const good = closeReport(r);
    return [['the wrong list: NOT ready; the meeting\'s list: ready, the summary unchanged', [bad.ready, bad.summaryState, good.ready, good.summaryState, good.externalVerified], [false, 'created', true, 'unchanged', true]]];
  }],
  ['13. no Registration of Documents table: fail closed, never every table', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    const reg = r.exec(() => r.s.findRegistrationTable_(r.body));
    r.user(() => r.body.removeChild(reg.table));
    decide(r, { 5: 'agreed' });
    const doc = snapshot(r.body);
    updateSummary(r);
    const out = [['Update Disposition Summary: "NOT updated", the reason, the document identical', [lastAlert(r)[0], /this report has no Registration of Documents table/.test(lastAlert(r)[1]), snapshot(r.body) === doc], [NOT_UPDATED, true, true]]];
    const sync = r.exec(() => r.s.syncDispositionSummaryOnUpdate_(r.body));
    out.push(['the sync of an update: the summary left as it is', [sync.state, snapshot(r.body) === doc], ['no-scope', true]]);
    const res = closeReport(r);
    out.push(['Close Report: NOT ready, nothing written', [res.ready, res.summaryState, snapshot(r.body) === doc], [false, 'no-scope', true]]);
    return out;
  }],
  ['14. tables the report did not register, and a TDoc reallocated to Removed, are left out', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    // 31: a table pasted by hand from another report (never registered). 06: reallocated to Removed.
    r.user(() => { const copy = r.table(id(5)).copy(); copy.getRow(0).getCell(1).setText(id(31)); r.body.insertTable(r.body.getChildIndex(r.table(id(5))) + 1, copy); });
    r.user(() => { r.body.appendParagraph('Document Reallocations'); r.body.appendTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason'], [id(6), '4.5', 'Removed', 'Synthetic']]); });
    updateSummary(r);
    // (The wording of the line differs between the releases; the count is what is compared.)
    return [['the summary: the registered TDocs less 06; 31 and 06 counted as left out', [summaryOf(r.body)[2], lastAlert(r)[1].split('\n').slice(-1)[0].replace(/^(\d+ TDoc tables) .*( left out\.)$/, '$1$2')],
      [[TITLES[2] + ' (1)', [id(5) + ' available']], '2 TDoc tables left out.']]];
  }],
  ['15. a registration table that has none of the report\'s TDoc tables: no scope, no summary of nothing', (code) => {
    const r = built(code, MEETING(), DECISIONS);
    updateSummary(r);
    const reg = r.exec(() => r.s.findRegistrationTable_(r.body));
    r.user(() => { for (let k = 1; k < reg.table.getNumRows(); k++) reg.table.getRow(k).getCell(0).setText(id(70 + k)); });
    const doc = snapshot(r.body);
    updateSummary(r);
    return [['"NOT updated", the reason; the summary of the report left as it was', [lastAlert(r)[0], /none of the 6 TDocs of the report's TDoc tables is registered in this report/.test(lastAlert(r)[1]), snapshot(r.body) === doc],
      [NOT_UPDATED, true, true]]];
  }]
];

/** Runs every scenario on `code`: { name: [failed check names] } ([] = passed; a throw is a failure). */
function runAll(code, report) {
  const out = {};
  SCENARIOS.forEach(([name, fn]) => {
    let results;
    try {
      results = fn(code);
    } catch (e) {
      results = [['(threw) ' + e.message, 'threw', 'no throw']];
    }
    if (report) console.log(name);
    out[name] = results.filter(([check_, actual, expected]) => {
      if (report) check(check_, actual, expected);
      return JSON.stringify(actual) !== JSON.stringify(expected);
    }).map((x) => x[0]);
  });
  return out;
}

// ====================================================================== S
console.log('S. the scenarios, on this Code.js');
const now = runAll(null, true);
check('every scenario passes on this Code.js', Object.keys(now).filter((n) => now[n].length), []);

// ====================================================================== O
console.log('O. the released code (T-2026.10.17) against the same scenarios');
if (released === null) {
  console.log('  note: ' + RELEASED_TAG + ' is not available in this checkout; the comparison with it is skipped.');
} else {
  const before = runAll(released, false);
  // The released code reads the TDoc list for the summary and needs it: it fails every scenario in which the list is not the
  // meeting's or cannot be read, and the report-local ones (no list fetched; no registration table; a registration table that
  // has none of the tables). It passes those this change must leave as they were.
  check('the released code fails exactly the scenarios this change is for', Object.keys(before).filter((n) => before[n].length), [
    '1a. Update Disposition Summary is report-local: it reads no TDoc list',
    '2. a normal report, the Portal unavailable',
    '3. a normal report, the Portal returns the list of another meeting',
    '4. a continuing report: two calls, two agenda TDocs, both calls in the summary',
    '5. a wrong Portal list that overlaps the report',
    '6. an MBS-style report without AGENDA_TDOC (the anchor)',
    '6a. an MBS-style report without AGENDA_TDOC, a wrong list that overlaps it without the anchor',
    '12. Close Report writes the summary but is not ready when the external check fails',
    '12a. ... and ready once the external check passes',
    '13. no Registration of Documents table: fail closed, never every table',
    '15. a registration table that has none of the report\'s TDoc tables: no scope, no summary of nothing']);
  check('... and passes the ones it must keep: the normal report, idempotence, rebuild, prose below, text typed into it, the update\'s refusal, left out',
    ['1. a normal report, the TDoc list of the meeting', '7. an existing summary is updated idempotently', '8. a deleted summary is written again', '9. text typed below the summary is kept',
      '10. text typed into the summary: refused, nothing changed', '11. Update Report Now still refuses a wrong Portal list, with zero mutation',
      '14. tables the report did not register, and a TDoc reallocated to Removed, are left out'].map((n) => [n, before[n]]),
    ['1. a normal report, the TDoc list of the meeting', '7. an existing summary is updated idempotently', '8. a deleted summary is written again', '9. text typed below the summary is kept',
      '10. text typed into the summary: refused, nothing changed', '11. Update Report Now still refuses a wrong Portal list, with zero mutation',
      '14. tables the report did not register, and a TDoc reallocated to Removed, are left out'].map((n) => [n, []]));
  // What the released code says in the case that started this: "the TDoc list of this meeting could not be read".
  const r = built(released, MEETING(), DECISIONS);
  wrongList(r, []);
  updateSummary(r);
  check('the released code, the wrong list: "Disposition Summary NOT updated ... could not be shown to be the list of this meeting" (the 6G incident)',
    [lastAlert(r)[0], /^The Disposition Summary was not written: the TDoc list of this meeting could not be read: The TDoc list was not used: it could not be shown to be the list of this meeting \(it has no row for S4aA260090/.test(lastAlert(r)[1])],
    [NOT_UPDATED, true]);
}

// ====================================================================== L
console.log('L. what changed in Code.js: against T-2026.10.17, only the summary\'s scope and Close Report');
{
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const view = codeWithoutReportLocalScope(FULL);
  // After T-2026.10.18: the Portal-page fallback of Close Report came later and has its own ledger (tests/close-report-portal-page-fallback.test.js);
  // the comparisons below are made without it.
  const CODE = codeWithoutCloseReportPageFallback(FULL).code;
  check('every change is in Code.js exactly once', view.problems, []);
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  check('Code.js is version 2.26.3 (2.26.2 released this change, T-2026.10.18; 2.26.3 the Portal-page fallback of Close Report; the ledgers take the headers out)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.3');
  if (released === null) {
    console.log('  note: ' + RELEASED_TAG + ' is not available in this checkout; the comparison with it is skipped.');
  } else {
    const before = released.replace(/\r/g, '');
    check('without it, Code.js is the Code.js of T-2026.10.17, byte for byte', view.code === before, true);
    check('the functions that changed: the summary\'s callers and Close Report, and the one call in continuousUpdateCore_',
      functionNames(before).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(CODE, n) !== functionSource(before, n)).sort(), REPORT_LOCAL_SCOPE_CHANGED_FUNCTIONS.slice().sort());
    check('the new functions: reportTdocScope_ and externalVerificationText_; none removed',
      [functionNames(CODE).filter((n) => functionNames(before).indexOf(n) === -1).sort(), functionNames(before).filter((n) => functionNames(CODE).indexOf(n) === -1)], [REPORT_LOCAL_SCOPE_NEW_FUNCTIONS.slice().sort(), []]);
    const identity = (src) => src.slice(src.indexOf('// TDOC LIST IDENTITY (after T-2026.10.16)'), src.indexOf('// END OF TDOC LIST IDENTITY'));
    check('the TDoc list identity is untouched: its section, and every place a TDoc list is acquired',
      [identity(CODE) === identity(before)].concat(['downloadAndGroupTdocs_', 'downloadAndProcessFromWeb', 'buildSkeletonWithTdocTables', 'readMeetingTdocScope_', 'meetingTdocScope_', 'checkRevisionsIn3gu_', 'classifyTdocDisposition_',
        'readReportDispositions_', 'renderDispositionSummary_', 'findDispositionSummarySection_'].map((n) => functionSource(CODE, n) === functionSource(before, n))), Array(11).fill(true));
    check('the classification is untouched: DISPOSITION_OF_STATUS_ and the classes', [/const DISPOSITION_OF_STATUS_ = \{[\s\S]*?\};/, /const DISPOSITION_CLASSES_ = \[[\s\S]*?\];/].map((re) => CODE.match(re)[0] === before.match(re)[0]), [true, true]);
    const core = (src) => functionSource(src, 'continuousUpdateCore_').replace(/\n    \/\/ After T-2026\.10\.1[67]: for the TDocs[\s\S]*?const summary = syncDispositionSummaryOnUpdate_\([^)]*\);/, '');
    check('continuousUpdateCore_: only the call of the summary sync changed', core(CODE) === core(before), true);
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll report-local scope checks passed.');
process.exitCode = failures ? 1 : 0;
