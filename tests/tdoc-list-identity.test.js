/**
 * The TDoc list identity (after T-2026.10.16) -- docs/TDOC_LIST_IDENTITY.md.
 *
 * 6G, 2026-10-30: the Portal list of the configured meeting ID (86178) gave, for hours, a well-formed workbook with the
 * 36 TDocs of earlier calls of the series (S4aP260008..043) instead of the 116 of this one (S4aP260067..182); nothing
 * in it names its meeting, and every update inserted the 36 into the report. A downloaded list is now used only when
 * it is shown to be the meeting's: the saved agenda TDoc has a row in it, else the anchor of a list accepted before
 * does, else -- the first time -- every TDoc table of the report is in it. Otherwise nothing is changed.
 *
 * These tests run the REAL build, update, Disposition Summary and Close Report of tests/helpers/template-report.js
 * (real TDoc-list reading through a fake sheet).
 *
 *   A.  the contamination case: report X already holds TDocs of meeting Y, the Portal returns Y's list (overlap NOT
 *       zero) -- every path refuses and changes nothing
 *   B.  the agenda TDoc: accepted with its row in any state, refused without it, reallocations do not count
 *   C.  the anchor (no valid agenda TDoc): bootstrap from the report, continuity, refusals, meeting and URL changes
 *   D.  the rules, pure
 *   E.  what changed in Code.js
 *
 * Run: node tests/tdoc-list-identity.test.js
 */

const fs = require('fs');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');
const { codeWithoutReportLocalScope, codeWithoutListIdentity, IDENTITY_SECTION_START, IDENTITY_SECTION_END, IDENTITY_SECTION_FUNCTIONS, IDENTITY_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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

const AGENDA = 'S4aA260090';
const id = (n) => 'S4aA2691' + String(n).padStart(2, '0');
const t = (n, agenda, status, extra) => Object.assign({ id: id(n), agenda: agenda, status: status || 'available', uploaded: true }, extra || {});
const cellLink = (cell) => cell.editAsText().getLinkUrl(0) || '';
/** The whole document, element by element: text, selected dropdown option, link. */
const snapshot = (body) => JSON.stringify(body._children.map((c) => (c.getType() === 'TABLE'
  ? ['T', Array.from({ length: c.getNumRows() }, (x, r) => Array.from({ length: c.getRow(r).getNumCells() }, (y, k) => [c.getRow(r).getCell(k).getText(), c.getRow(r).getCell(k)._dropdown ? c.getRow(r).getCell(k)._dropdown.optionId : null, cellLink(c.getRow(r).getCell(k))]))]
  : ['P', c.getHeading(), c.getText()])));
const props = (r) => JSON.stringify(Object.keys(r.docProps._store).sort().map((k) => [k, r.docProps._store[k]]));
const statusOf = (r, n) => r.field(id(n), 'Status');
const listFetches = (r, from) => r.fetches.slice(from).filter((f) => /GenerateDocumentList/.test(f)).length;
const lastAlert = (r) => r.ui.alerts[r.ui.alerts.length - 1];

/** Close Report shows its result in an HTML dialog; it is recorded with the alerts. */
function withDialogs(r) {
  r.s.HtmlService = { createHtmlOutput: (html) => { const o = { setWidth: () => o, setHeight: () => o, getContent: () => html }; return o; } };
  r.ui.showModalDialog = (output, title) => { r.ui.alerts.push([title, r.logs[r.logs.length - 1], 'dialog']); };
  return r;
}
/** A built template report (ad-hoc, agenda TDoc S4aA260090) with status dropdowns; `decisions` picked by hand. */
function built(tdocs, decisions, options) {
  const r = withDialogs(templateReport(Object.assign({ tdocs: tdocs, tableSeparators: true, endParagraph: true }, options || {})));
  const b = r.build();
  if (!b.ok) throw new Error('build failed: ' + (b.error || b.message));
  const u = r.update();
  if (!u.success) throw new Error('update failed: ' + u.error);
  r.user(() => Object.keys(decisions || {}).forEach((n) => r.api.pick(id(n), decisions[n])));
  return r;
}
/** A main-meeting report (no agenda TDoc) whose TDoc tables were written before: [n, status] each, under 4.3. */
function mainReport(existing, list, extraProps) {
  const r = withDialogs(templateReport({ tdocs: list, tableSeparators: true, endParagraph: true,
    props: Object.assign({ MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137_Synthetic', MEETING_NUMBER: '137', MEETING_ID: '60778', AGENDA_ITEM_PREFIX: '4.', AGENDA_TDOC: null }, extraProps || {}) }));
  r.user(() => {
    r.body.appendParagraph('4.3 Performance requirements').setHeading('HEADING2');
    existing.forEach(([n, status]) => r.body.appendTable([['TDoc', id(n)], ['Title', 'Synthetic title of ' + id(n)], ['Source', 'ExampleCorp'], ['Agenda Item', '4.3'], ['Status', status]]));
  });
  return r;
}
const anchorOf = (r) => { const v = r.docProps.getProperty('TDOC_LIST_ANCHOR'); return v ? JSON.parse(v) : null; };

// ====================================================================== A
console.log('A. the contamination case: report X holds TDocs of meeting Y; the Portal returns Y\'s list');
{
  // Meeting X: agenda TDoc S4aA260090. Its report already holds 21 and 22, TDocs of meeting Y (as live 6G holds the 36).
  const r = built([t(1, '4.3'), t(2, '4.3', 'revised', { revisedTo: id(3) }), t(3, '4.3'), t(4, '4.4'), t(21, '4.4'), t(22, '4.5')], { 1: 'agreed', 4: 'noted' });
  r.exec(() => r.s.updateDispositionSummary());
  r.exec(() => r.s.buildRevisionMapFromTdocList_(r.s.getReportConfig_(), r.s.downloadAndGroupTdocs_(r.s.getReportConfig_())));
  const doc = snapshot(r.body);
  const store = props(r);
  const map = r.docProps.getProperty('REVISION_MAP');
  const order = r.order();
  const statuses = [1, 2, 3, 4, 21, 22].map((n) => statusOf(r, n));
  // The Portal returns meeting Y's list: well-formed, not empty, no row for X's agenda TDoc, 21 and 22 in it (with other
  // statuses), and TDocs X never had -- under an agenda item of X's report and one it does not have.
  r.listAgenda = null;
  r.tdocs = [t(21, '4.4', 'withdrawn'), t(22, '4.5', 'agreed'), t(23, '4.3'), t(24, '4.7', 'reserved', { uploaded: false })];
  const y = r.exec(() => r.s.readReportDispositions_(r.body, { scope: r.s.meetingTdocScope_(r.tdocs.map((x) => ({ row: [x.id], tdocCol: 0 }))) }));
  check('the setting: Y\'s list overlaps the report (21, 22) -- the zero-overlap guard of the summary would not catch it', y.entries.map((e) => e.tdoc), [id(21), id(22)]);

  r.mark();
  const f = r.fetches.length;
  const u = r.update();
  check('the update: refused, "not updated", the reason', [u.success, u.notUpdated, /could not be shown to be the list of this meeting \(it has no row for S4aA260090, the agenda TDoc of this meeting/.test(u.error)], [false, true, true]);
  // (The Session column of a report with sessions: tests/adhoc-tdoc-sessions.test.js, "the list of another meeting".)
  check('... the document identical: no Y table, no agenda heading, no registration row, no revision moved, no Disposition cell, no summary change', snapshot(r.body) === doc, true);
  check('... no table added or moved; no status synchronised from Y (21 stays as it was, not withdrawn; 22 not agreed)', [r.order(), [1, 2, 3, 4, 21, 22].map((n) => statusOf(r, n))], [order, statuses]);
  check('... the document properties identical: REVISION_MAP, the abstract cache, everything', [props(r) === store, r.docProps.getProperty('REVISION_MAP') === map], [true, true]);
  check('... the list read once, nothing else asked (no abstract request)', [listFetches(r, f), r.requests.length], [1, 0]);

  const f2 = r.fetches.length;
  // After T-2026.10.17 the summary is the report's (its registration table): Update Disposition Summary reads no TDoc list, so
  // Y's list neither stops it nor gets into it. Close Report reads the list for its external check, refuses it, and is not ready.
  r.exec(() => r.s.updateDispositionSummary());
  check('Update Disposition Summary: from the report -- already up to date, the document identical; Y\'s list not read', [lastAlert(r)[0], snapshot(r.body) === doc, lastAlert(r)[1].split('\n')[0], listFetches(r, f2)],
    ['Disposition Summary', true, 'Disposition Summary was already up to date.', 0]);
  const c = r.exec(() => r.s.closeReportCore_());
  check('Close Report: NOT READY, external verification not completed (Y\'s list refused), the revised TDoc not verified, nothing written',
    [c.ready, c.externalVerified, c.summaryState, c.checked, c.revisions.map((x) => [x.tdoc, x.kind]), snapshot(r.body) === doc, /^External verification could not be completed: .*could not be shown to be the list of this meeting/.test(c.problems[0] || '')],
    [false, false, 'unchanged', 6, [[id(2), 'revision-not-verified']], true, true]);
  // After T-2026.10.18: with the list refused, Close Report tries the Portal's TDoc list page (TdocList.aspx) through the same
  // guard (tests/close-report-portal-page-fallback.test.js); here that page is not served, so nothing is verified.
  check('... the list read once, by Close Report, then the Portal page tried once; nothing else fetched', [listFetches(r, f2), r.fetches.slice(f2).filter((f) => /TdocList\.aspx/.test(f)).length, r.fetches.length - f2], [1, 1, 2]);

  const b = r.build();
  check('Build Report from Scratch: refused before the document is cleared', [b.ok, snapshot(r.body) === doc, props(r) === store], [false, true, true]);
  check('everything together: the document and its properties as they were', [snapshot(r.body) === doc, props(r) === store], [true, true]);

  // The list of meeting X again: the update goes on.
  r.listAgenda = AGENDA;
  r.tdocs = [t(1, '4.3'), t(2, '4.3', 'revised', { revisedTo: id(3) }), t(3, '4.3'), t(4, '4.4'), t(5, '4.4'), t(21, '4.4'), t(22, '4.5')];
  check('X\'s list again: accepted, the new TDoc 05 added', [r.update().success, r.order().indexOf(id(5)) !== -1], [true, true]);
}

// ====================================================================== B
console.log('B. the agenda TDoc: a row in any state; reallocations do not count');
{
  const r = built([t(1, '4.3'), t(2, '4.4')]);
  check('the meeting\'s list (it has the agenda TDoc): accepted', [r.update().success, r.logs.some((l) => /TDoc list identity: the list of this meeting \(agenda\)/.test(l))], [true, true]);
  check('... no anchor is written: the agenda TDoc is the identity', r.docProps.getProperty('TDOC_LIST_ANCHOR'), null);
  [['reserved, not uploaded', { status: 'reserved', uploaded: false }], ['withdrawn', { status: 'withdrawn' }], ['revised', { status: 'revised' }]].forEach(([what, row]) => {
    r.tdocs = r.tdocs.filter((x) => x.id !== AGENDA).concat([Object.assign({ id: AGENDA, agenda: '' }, row)]);
    check('its row ' + what + ': accepted', r.update().success, true);
  });
  // The agenda TDoc reallocated to Removed in the report: the identity is the raw list, before the reallocations.
  r.user(() => { r.body.appendParagraph('Document Reallocations'); r.body.appendTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason'], [AGENDA, '2', 'Removed', 'Synthetic']]); });
  check('the agenda TDoc reallocated to Removed: still accepted (the raw list has its row)', r.update().success, true);
  const doc = snapshot(r.body);
  r.tdocs = r.tdocs.filter((x) => x.id !== AGENDA);
  r.listAgenda = null;
  const u = r.update();
  check('a list without its row (everything else the same): refused, nothing changed', [u.success, u.notUpdated, snapshot(r.body) === doc], [false, true, true]);
  // A list reallocating TDocs of another list into the report does not make it the meeting's.
  r.tdocs = [t(1, '4.3'), t(2, '4.4'), t(9, '4.4')];
  r.user(() => { r.body.appendParagraph('Document Reallocations'); r.body.appendTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason'], [id(9), '4.4', '4.3', 'Synthetic']]); });
  check('... also with a reallocation of one of its TDocs in the report', [r.update().success, r.order().indexOf(id(9))], [false, -1]);
}

// ====================================================================== C
console.log('C. the anchor: a report without an agenda TDoc (MBS: agenda.csv; main meetings)');
{
  // C1. Bootstrap: every TDoc table of the report is in the list -> accepted, the anchor recorded after the check.
  const r = mainReport([[1, 'agreed'], [2, 'available']], [t(1, '4.3'), t(2, '4.3'), t(3, '4.3')]);
  const before = props(r);
  const scope = r.exec(() => r.s.readMeetingTdocScope_());
  check('a reader (Close Report, the summary) checks the same way and stores nothing', [scope.ok, props(r) === before], [true, true]);
  const u1 = r.update();
  check('the first update: the report accounts for the list -- accepted, 03 added', [u1.success, r.order()], [true, [id(1), id(2), id(3)]]);
  const cfg = r.exec(() => r.s.getReportConfig_());
  check('... the anchor: a TDoc of the report and the list, with the meeting ID and the TDoc list URL',
    (({ tdoc, meetingId, tdocListUrl }) => [tdoc, meetingId, tdocListUrl])(anchorOf(r)), [id(1), '60778', cfg.TDOC_LIST_URL]);
  check('... logged', r.logs.some((l) => l === 'TDoc list identity: the list of this meeting (report; anchor ' + id(1) + ' recorded)'), true);
  const recorded = r.docProps.getProperty('TDOC_LIST_ANCHOR');

  // C2. Continuity.
  r.tdocs = [t(1, '4.3', 'withdrawn'), t(2, '4.3', 'noted'), t(3, '4.3'), t(4, '4.3'), t(5, '4.4')];
  check('a later list, appended to, statuses changed, the anchor withdrawn: accepted; the anchor is not rewritten',
    [r.update().success, r.order().length, r.docProps.getProperty('TDOC_LIST_ANCHOR') === recorded, r.logs.some((l) => l === 'TDoc list identity: the list of this meeting (anchor)')], [true, 5, true, true]);
  r.user(() => { r.body.appendParagraph('Document Reallocations'); r.body.appendTable([['TDoc', 'Original Agenda', 'New Agenda', 'Reason'], [id(1), '4.3', 'Removed', 'Synthetic']]); });
  check('the anchor reallocated to Removed in the report: accepted (the raw list has its row)', r.update().success, true);

  // C3. Refusals.
  const doc = snapshot(r.body);
  const store = props(r);
  r.tdocs = [t(31, '4.3'), t(32, '4.4')];
  const d = r.update();
  check('a disjoint list: refused, nothing changed', [d.success, d.notUpdated, snapshot(r.body) === doc, props(r) === store, /it has no row for S4aA269101, a TDoc of the list this meeting had before/.test(d.error)], [false, true, true, true, true]);
  r.tdocs = [t(2, '4.3'), t(3, '4.3'), t(4, '4.3'), t(5, '4.4'), t(31, '4.4')];
  const p = r.update();
  check('a list that overlaps the report (02..05) but has no row for the anchor: refused, nothing changed', [p.success, snapshot(r.body) === doc, props(r) === store], [false, true, true]);

  // C4. Meeting Setup changes the meeting or the URL: the anchor is no anchor; the report is the evidence again.
  r.docProps.setProperty('MEETING_ID', '60779');
  r.tdocs = [t(1, '4.3'), t(2, '4.3'), t(3, '4.3'), t(4, '4.3'), t(5, '4.4'), t(6, '4.4')];
  check('another meeting ID: the old anchor is not used; the report accounts for the list -- accepted, a new anchor for the new meeting',
    [r.update().success, anchorOf(r).meetingId, anchorOf(r).tdoc], [true, '60779', id(1)]);
  r.docProps.setProperty('TDOC_LIST_URL', 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=60779');
  r.tdocs = [t(2, '4.3'), t(3, '4.3'), t(4, '4.3'), t(5, '4.4'), t(6, '4.4')];
  const url = r.update();
  check('another TDoc list URL: the old anchor is not used; a list that does not account for the report (01 missing) is refused',
    [url.success, /1 of the 6 TDocs of the report's TDoc tables are not in it \(S4aA269101\)/.test(url.error), anchorOf(r).tdocListUrl !== 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=60779'], [false, true, true]);

  // C5. A report that does not account for the list (contaminated, or tables of another meeting) is never a bootstrap.
  const mixed = mainReport([[1, 'agreed'], [2, 'available'], [21, 'available']], [t(1, '4.3'), t(2, '4.3'), t(3, '4.3')]);
  const mixedDoc = snapshot(mixed.body);
  const m1 = mixed.update();
  check('no anchor, the report holds a TDoc the list does not have (21): refused, nothing changed, no anchor', [m1.success, snapshot(mixed.body) === mixedDoc, anchorOf(mixed)], [false, true, null]);
  mixed.tdocs = [t(21, '4.3'), t(22, '4.3')];
  check('... and the other list (21, 22): refused too -- nothing is guessed', [mixed.update().success, snapshot(mixed.body) === mixedDoc, anchorOf(mixed)], [false, true, null]);
  const empty = mainReport([], [t(1, '4.3'), t(2, '4.3')]);
  const e1 = empty.update();
  check('no anchor, no TDoc table: refused (the update has no evidence; Build Report from Scratch takes the list)', [e1.success, /Build Report from Scratch takes the list/.test(e1.error), anchorOf(empty)], [false, true, null]);

  // C6. An agenda TDoc that is not a TDoc number is no identity: the anchor applies.
  const manual = mainReport([[1, 'agreed']], [t(1, '4.3'), t(2, '4.3')], { AGENDA_TDOC: 'S4aA260090-manual' });
  check('AGENDA_TDOC "S4aA260090-manual": not an identity TDoc -- the report is the evidence, the anchor recorded', [manual.update().success, anchorOf(manual).tdoc], [true, id(1)]);
}

// ====================================================================== D
console.log('D. the rules, pure');
{
  const { sandbox: s } = loadCode();
  const H = ['TDoc', 'Title', 'Uploaded'];
  const list = (rows) => [H].concat(rows.map(([n, up]) => [n, 'x', up || '']));
  const cfg = { MEETING_ID: '86178', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86178', AGENDA_TDOC: 'S4aP260098' };
  const E = (data, c, saved, report, use) => JSON.parse(JSON.stringify(s.evaluateTdocListIdentity_(data, c, saved || null, report || [], use || 'read')));
  // The two lists of 2026-10-30, as far as identity goes.
  const right = list([['S4aP260067'], ['S4aP260098', '2026-09-18'], ['S4aP260149']]);
  const wrong = list([['S4aP260008'], ['S4aP260020', '2026-02-19'], ['S4aP260043']]);
  check('6G, the right list (has S4aP260098): accepted by its agenda TDoc, no anchor', (({ ok, via, anchor }) => [ok, via, anchor])(E(right, cfg)), [true, 'agenda', null]);
  check('6G, the wrong list: refused, whatever the report holds', [E(wrong, cfg).ok, E(wrong, cfg, null, ['S4aP260008', 'S4aP260020'], 'update').ok], [false, false]);
  check('an empty list: refused', E([H], cfg).error, 'The TDoc list was not used: it could not be shown to be the list of this meeting (it has no TDoc)');
  const noAgenda = Object.assign({}, cfg, { AGENDA_TDOC: '' });
  const saved = JSON.stringify({ tdoc: 'S4aP260067', meetingId: '86178', tdocListUrl: cfg.TDOC_LIST_URL });
  check('an anchor: its row there -> accepted; not there -> refused', [E(right, noAgenda, saved).ok, E(wrong, noAgenda, saved).ok], [true, false]);
  check('an anchor of another meeting ID or URL is no anchor', [s.tdocListAnchorFor_(saved, Object.assign({}, noAgenda, { MEETING_ID: '86179' })), s.tdocListAnchorFor_(saved, Object.assign({}, noAgenda, { TDOC_LIST_URL: 'https://x' }))], [null, null]);
  check('a damaged anchor is no anchor', [s.tdocListAnchorFor_('{', noAgenda), s.tdocListAnchorFor_(JSON.stringify({ tdoc: 'nonsense', meetingId: '86178', tdocListUrl: cfg.TDOC_LIST_URL }), noAgenda)], [null, null]);
  check('bootstrap: the report within the list -> accepted, an anchor of the report (an uploaded one first) to write on an update',
    (({ ok, via, anchor }) => [ok, via, anchor && anchor.tdoc])(E(right, noAgenda, null, ['S4aP260149', 'S4aP260098'], 'update')), [true, 'report', 'S4aP260098']);
  check('... a reader writes none', E(right, noAgenda, null, ['S4aP260149'], 'read').anchor, null);
  check('bootstrap refused: one TDoc of the report not in the list; no TDoc table at all', [E(right, noAgenda, null, ['S4aP260149', 'S4aP260008'], 'update').ok, E(right, noAgenda, null, [], 'update').ok], [false, false]);
  check('a build without an identity takes the list: an anchor of the list (an uploaded one first)', (({ ok, via, anchor }) => [ok, via, anchor.tdoc])(E(right, noAgenda, null, [], 'build')), [true, 'build', 'S4aP260098']);
  check('a build with an agenda TDoc or an anchor is checked like any other use', [E(wrong, cfg, null, [], 'build').ok, E(wrong, noAgenda, saved, [], 'build').ok], [false, false]);
  check('the TDoc numbers are canonical; text that is no TDoc number is not one', Object.keys(s.tdocListRawNumbers_([H, [' S4aP260067 '], ['S4aP260067'], ['Total'], [''], ['S4-261234']], 0)), ['S4aP260067', 'S4-261234']);
}

// ====================================================================== E
console.log('E. what changed in Code.js');
{
  // The report-local scope of the Disposition Summary came after it (tests/disposition-report-local-scope.test.js): this is
  // Code.js without that, the Code.js of T-2026.10.17.
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const CODE = codeWithoutReportLocalScope(FULL).code;
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const view = codeWithoutListIdentity(FULL);
  check('the change is in Code.js exactly once: one section, four call sites', view.problems, []);
  const section = CODE.slice(CODE.indexOf(IDENTITY_SECTION_START), CODE.indexOf(IDENTITY_SECTION_END) + IDENTITY_SECTION_END.length);
  check('the section: these functions, in this order', functionNames(section), IDENTITY_SECTION_FUNCTIONS);
  check('it stands directly after the Disposition Summary section', CODE.indexOf('// END OF DISPOSITION SUMMARY AND CLOSE REPORT\n// =========================================================\n\n' + IDENTITY_SECTION_START) !== -1, true);
  let before = null;
  try {
    before = execFileSync('git', ['show', 'c48ba943945dd782ed95c00f0e695d150e6b52d3:Code.js'], { cwd: require('path').dirname(CODE_JS_PATH), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    before = null;
  }
  if (before === null) {
    console.log('  note: the candidate before it (c48ba94) is not in this checkout; the comparison with it is skipped.');
  } else {
    check('without it, Code.js is the candidate before it (c48ba94), byte for byte', view.code === before, true);
    check('the functions that changed outside the section: the four places a TDoc list is acquired', functionNames(before).filter((n, i, all) => all.indexOf(n) === i).filter((n) => functionSource(CODE, n) !== functionSource(before, n)).sort(),
      IDENTITY_CHANGED_FUNCTIONS.slice().sort());
    check('the new functions are those of the section; none was removed', [functionNames(CODE).filter((n) => functionNames(before).indexOf(n) === -1).sort(), functionNames(before).filter((n) => functionNames(CODE).indexOf(n) === -1)],
      [IDENTITY_SECTION_FUNCTIONS.slice().sort(), []]);
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll TDoc list identity checks passed.');
process.exitCode = failures ? 1 : 0;
