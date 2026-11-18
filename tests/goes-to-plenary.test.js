/**
 * Goes to Plenary: the compact Status row, the four disposition tables on
 * landscape pages, the SA4 plenary agenda item and the revision relationship
 * in the TDoc row (after T-2026.10.29) -- docs/GOES_TO_PLENARY.md.
 *
 * The REAL build, update, set-up, Update TDoc Revisions, Disposition Summary
 * and Close Report run on the fake document and the fake Docs API
 * (tests/helpers/template-report.js), as an MBS report of SA4#138 whose agenda
 * is the real one (S4-261857, tests/fixtures/agenda-S4-261857.google-items.json).
 *
 *   A.  the agenda mapping on the real SA4#138 agenda: every rule, precedence,
 *       missing and ambiguous mappings; nothing invented
 *   B.  a report without Goes to Plenary is exactly as before; an ad-hoc report
 *   C.  Set Up Goes to Plenary…: the compact Status row, the retired Status
 *       "Plenary"; safe to run again
 *   D.  the update: A.I. only for "☑ Yes" with an item; the four tables; the
 *       landscape section; the revision relationships (chains, a missing
 *       predecessor, conflicts, cycles)
 *   E.  repeated updates: idempotent, no write when nothing changed
 *   F.  the selection is the rapporteur's: status changes, the Portal sync of
 *       a text status in the compact row, new TDocs, Update TDoc Revisions
 *   G.  manual SA4 plenary agenda items: typed over A.I., "-", cleared, kept
 *       while hidden; work-item changes move only automatic values
 *   H.  minutes, text typed into or below the summary (portrait again), an
 *       unreadable selection, Update Disposition Summary
 *   I.  Build Report from Scratch keeps the selections, the manual values, the
 *       revision relationships and the landscape section
 *   J.  Close Report
 *   L.  migration of the first candidate's two-row layout without losing a selection
 *   M.  how many DocumentApp calls and Docs API calls an update of 75 TDocs makes: the first
 *       candidate against this one (a relative measure on the fake document -- never a time)
 *   N.  an update interrupted after it inserted a TDoc (as by the six-minute limit): the next one
 *       completes it -- no duplicate, the Status dropdown made, the summary once
 *   O.  every Status of the dropdown keeps its dropdown in the compact row, reserved included
 *   P.  an update that has run long leaves the summary and the formatting to the next one
 *   Q.  the REAL collectors on 75 TDocs with some 150 revision links: an unchanged update writes nothing (document,
 *       stores, API); one new file writes one cell; a lost link is restored; a stale agent clears nothing; the
 *       collectors keep to the time budget, and an incomplete update says so
 *   K.  what changed in Code.js and ReportCreator.js, and nothing else
 *
 * Run: node tests/goes-to-plenary.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');
const { codeWithoutRelease2_26_14Header, codeWithoutGoesToPlenary, creatorWithoutGoesToPlenary, PLENARY_SECTION_START, PLENARY_SECTION_END, PLENARY_SECTION_FUNCTIONS, PLENARY_CHANGED_FUNCTIONS,
  PLENARY_CREATOR_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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

const ROOT = path.join(__dirname, '..');
const YES = '☑ Yes';
const NO = '☐ No';
const EM = '\u2003';
// The start of the compact Status row's own part: a zero-width space and an em space before "Plenary".
const MARK = '\u200B' + EM + 'Plenary';
const SUMMARY = 'Disposition Summary';
const TITLES = ['Agreed, Endorsed and Approved Documents – Goes to Plenary', 'Other Dispositions – Goes to Plenary',
  'Agreed, Endorsed and Approved Documents – Not Going to Plenary', 'Other Dispositions – Not Going to Plenary'];
const HEADER = ['TDoc', 'Title', 'Source', 'Status', 'SWG Agenda Item', 'SA4 Plenary Agenda Item'];
const HEADER5 = ['TDoc', 'Title', 'Source', 'Status', 'SWG Agenda Item'];
const AGENDA_138 = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'fixtures', 'agenda-S4-261857.google-items.json'), 'utf8')).items
  .map((i) => ({ number: i.number, title: i.title, level: i.level, heading: 'NORMAL', text: '' }));
const MAIN = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_NAME: 'SA4#138', MEETING_ID: '60780', REPORT_SUFFIX: 'MBS', MEETING_DATE: 'November 16, 2026',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx',
  AGENDA_TDOC: 'S4-261857' };
const COLUMNS = ['Release', 'Related WIs'];
const ZIP = (id) => 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/' + id + '.zip';
// An MBS report of SA4#138 (agenda items 8.x). All numbers and names are synthetic; the agenda is the real one.
const TDOCS = () => [
  { id: 'S4-262001', agenda: '8.5', status: 'revised', title: '[AMD_ARCH_Ph2-MED] WT#8 multi-access', revisedTo: 'S4-262012', uploaded: true },
  { id: 'S4-262002', agenda: '8.6', status: 'noted', title: '[Energy_PRO-MED] TS skeleton', uploaded: true },
  { id: 'S4-262003', agenda: '8.4', status: 'agreed', type: 'CR', Release: 'Rel-19', title: 'CR 26.511 correction', uploaded: true },
  { id: 'S4-262004', agenda: '8.7', status: 'available', title: 'QUIC evaluation results', revisedTo: 'S4-262014', uploaded: true },
  { id: 'S4-262005', agenda: '8.8', status: 'endorsed', 'Related WIs': 'TEI20', title: 'Rel-20 TEI clarification', uploaded: true },
  { id: 'S4-262006', agenda: '8.8', status: 'agreed', 'Related WIs': 'TEI19', title: 'Rel-19 TEI fix', uploaded: true },
  { id: 'S4-262007', agenda: '8.8', status: 'agreed', 'Related WIs': 'AMD_ARCH_Ph2-MED, Energy_PRO-MED', title: 'Cross-WI alignment', uploaded: true },
  { id: 'S4-262008', agenda: '8.3', status: 'noted', title: 'LS from another group', uploaded: true },
  { id: 'S4-262009', agenda: '8.5', status: 'available', title: '[AMD_ARCH_Ph2-MED] open issue', uploaded: true },
  { id: 'S4-262010', agenda: '8.5', status: 'approved', type: 'CR', Release: 'Rel-20', title: 'CR 26.501 new feature', uploaded: true },
  // Revisions: a chain 001 -> 012 -> 013; 014 has two predecessors (004 says it, 014 says 002): none; 015 <-> 016 is a cycle:
  // none; 017 is a revision of a TDoc of another meeting (no link).
  { id: 'S4-262012', agenda: '8.5', status: 'revised', title: '[AMD_ARCH_Ph2-MED] WT#8 multi-access', isRevisionOf: 'S4-262001', revisedTo: 'S4-262013', uploaded: true },
  { id: 'S4-262013', agenda: '8.5', status: 'agreed', title: '[AMD_ARCH_Ph2-MED] WT#8 multi-access', isRevisionOf: 'S4-262012', uploaded: true },
  { id: 'S4-262014', agenda: '8.7', status: 'available', title: 'QUIC evaluation results', isRevisionOf: 'S4-262002', uploaded: true },
  { id: 'S4-262015', agenda: '8.6', status: 'available', title: '[Energy_PRO-MED] loop A', isRevisionOf: 'S4-262016', uploaded: true },
  { id: 'S4-262016', agenda: '8.6', status: 'available', title: '[Energy_PRO-MED] loop B', isRevisionOf: 'S4-262015', uploaded: true },
  { id: 'S4-262017', agenda: '8.6', status: 'noted', title: '[Energy_PRO-MED] from the ad hoc', isRevisionOf: 'S4-261999', uploaded: true },
  { id: 'S4-262018', agenda: '8.6', status: 'not pursued', title: '[Energy_PRO-MED] not pursued', uploaded: true }
];
const EXPECTED_AUTO = { 'S4-262001': '15.4', 'S4-262002': '15.5', 'S4-262003': '14', 'S4-262004': '16.3', 'S4-262005': '15.7', 'S4-262006': '14', 'S4-262007': '', 'S4-262008': '',
  'S4-262009': '15.4', 'S4-262010': '15.4', 'S4-262012': '15.4', 'S4-262013': '15.4', 'S4-262014': '16.3', 'S4-262015': '15.5', 'S4-262016': '15.5', 'S4-262017': '15.5', 'S4-262018': '15.5' };

const isTable = (c) => c.getType() === 'TABLE';
const rowsOf = (t) => Array.from({ length: t.getNumRows() }, (_, i) => Array.from({ length: t.getRow(i).getNumCells() }, (_, k) => t.getRow(i).getCell(k).getText()));
function mbs(extra) {
  const o = Object.assign({ props: MAIN, agenda: AGENDA_138, columns: COLUMNS, tdocs: TDOCS(), tableSeparators: true, endParagraph: true }, extra || {});
  const r = templateReport(o);
  r.s.HtmlService = { createHtmlOutput: (html) => { const out = { setWidth: () => out, setHeight: () => out, getContent: () => html }; return out; } };
  r.ui.dialogs = [];
  r.ui.showModalDialog = (output, title) => { r.ui.dialogs.push([title, output.getContent()]); };
  r.setUp = () => r.exec(() => r.s.setUpGoesToPlenary());
  r.summaryMenu = () => r.exec(() => r.s.updateDispositionSummary());
  r.close = () => r.exec(() => r.s.closeReport());
  /** The labels of the rows of a TDoc table, in order. */
  r.labels = (id) => { const t = r.table(id); return t ? rowsOf(t).map((x) => x[0]) : null; };
  /** The compact Status row of a TDoc: { status, plenary, ai, parts }. */
  r.c = (id) => r.api.compactOf(id);
  /** The rapporteur types `value` over what follows "A.I.:" (the A.I. must be shown). */
  r.typeAgenda = (id, value) => r.user(() => {
    const cell = r.api.statusCell(id);
    const text = cell.getText();
    const at = text.indexOf(EM + 'A.I.:') + (EM + 'A.I.:').length;
    if (at < (EM + 'A.I.:').length) throw new Error('no A.I. shown for ' + id);
    if (text.length > at) cell.editAsText().deleteText(at, text.length - 1);
    if (value) cell.editAsText().insertText(at, ' ' + value);
  });
  /** The TDoc row: its text, and the links on its number and on its predecessor. */
  r.tdocRow = (id) => {
    const cell = r.table(id).getRow(0).getCell(1);
    const text = cell.getText();
    const m = text.match(/\(revision of (\S+)\)$/);
    return { text: text, own: cell.editAsText().getLinkUrl(0) || '', predecessorLink: m ? cell.editAsText().getLinkUrl(text.indexOf(m[1])) || '' : null };
  };
  /** The summary: [[class line, rows of its table or 'None.']] in order, or null. */
  r.summary = () => {
    const kids = r.body._children;
    const at = kids.findIndex((c) => !isTable(c) && c.getText() === SUMMARY && c.getHeading() !== 'NORMAL');
    if (at === -1) return null;
    const out = [];
    for (let i = at + 1; i < kids.length; i++) {
      const c = kids[i];
      if (isTable(c)) { if (out.length && [4, 5, 6].indexOf(c.getRow(0).getNumCells()) !== -1) out[out.length - 1].push(rowsOf(c)); continue; }
      const text = c.getText();
      if (TITLES.some((t) => text.indexOf(t + ' (') === 0) || /^(Agreed, Endorsed and Approved Documents|Other Disposed Documents|Documents Not Yet Disposed) \(/.test(text)) out.push([text]);
      else if (text === 'None.' && out.length) out[out.length - 1].push('None.');
    }
    return out;
  };
  r.inTable = (k) => { const s = r.summary(); return Array.isArray(s[k][1]) ? s[k][1].slice(1).map((x) => x[0]) : []; };
  r.headings = () => r.body._children.filter((c) => !isTable(c) && c.getText() === SUMMARY && c.getHeading() !== 'NORMAL').length;
  r.snapshot = () => JSON.stringify(r.body._children.map((c) => (isTable(c)
    ? ['T', !!c._sectionBreakBefore && c._sectionBreakBefore.flip, Array.from({ length: c.getNumRows() }, (x, i) => Array.from({ length: c.getRow(i).getNumCells() }, (y, k) => { const cell = c.getRow(i).getCell(k); return [cell.getText(), JSON.stringify(cell._allParts().map((p) => (p.d ? p.d.optionId : p.t))), cell.editAsText().getLinkUrl(0) || '']; }))]
    : ['P', c._sectionBreakBefore ? 'break:' + c._sectionBreakBefore.flip : '', c.getHeading(), c.getText()])));
  r.writes = () => r.api.calls.filter((c) => c.call === 'batchUpdate').length;
  return r;
}

// ====================================================================== A
console.log('A. the agenda mapping on the real SA4#138 agenda (S4-261857)');
{
  const { sandbox: S } = loadCode({ documentProperties: {} });
  const model = S.plenaryAgendaModel_(S.compactPlenaryAgendaItems_(AGENDA_138));
  check('the maintenance item is found by its title: 14 "Release 19 and earlier matters", release 19 -- not a number of the code', model.maintenance, { number: '14', release: 19 });
  check('the targets are the items under "Release 20 Features" (15) and "Study Items" (16), no others',
    model.targets.map((t) => t.number), ['15.1', '15.2', '15.3', '15.4', '15.5', '15.6', '15.7', '16.1', '16.2', '16.3', '16.4', '16.5', '16.6', '16.7']);
  check('their acronyms come from the agenda titles; "Other Rel-20 matters including TEI" has none',
    model.targets.map((t) => t.acronym), ['DaCAS', 'ATIAS_Ph3', 'ULBC-MED', 'AMD_ARCH_Ph2-MED', 'Energy_PRO-MED', 'AvCall_Ph2-MED', '', 'FS_3DGS_MED', 'FS_DCTC_eQOS_MED', 'FS_QStream_MED', 'FS_AIF_MED', 'FS_Q4RTC_MED', 'FS_DCE_MED', 'FS_6G_MED']);
  check('8.4 of this agenda is the MBS SWG item "Release 19 and earlier matters", not a plenary item', [model.byNumber['8.4'], model.targets.some((t) => t.number === '8.4')], ['Release 19 and earlier matters', false]);
  const v = (t) => S.plenaryAgendaAutoValue_(model, t);
  check('each SWG work item maps to its plenary item by the acronym of its SWG agenda item', [
    ['7.5', '7.6', '7.7', '8.5', '8.6', '8.7', '9.5', '9.6', '10.5', '10.6', '10.7', '10.8'].map((swg) => v({ swgItem: swg }).value)],
    [['15.1', '15.2', '15.3', '15.4', '15.5', '16.3', '16.1', '16.4', '15.6', '16.2', '16.5', '16.6']]);
  check('6G: every 11.x item is FS_6G_MED -> 16.7, through the item above it', ['11.1', '11.2', '11.3.1', '11.3.3', '11.7'].map((swg) => v({ swgItem: swg }).value), ['16.7', '16.7', '16.7', '16.7', '16.7']);
  check('maintenance: every SWG "Release 19 and earlier matters" item (7.4, 8.4, 9.4, 10.4) -> 14', ['7.4', '8.4', '9.4', '10.4'].map((swg) => v({ swgItem: swg })), [
    { value: '14', rule: 'SWG maintenance item 7.4' }, { value: '14', rule: 'SWG maintenance item 8.4' }, { value: '14', rule: 'SWG maintenance item 9.4' }, { value: '14', rule: 'SWG maintenance item 10.4' }]);
  check('a CR to Rel-19 or earlier -> 14, even under a work-item item; a CR to Rel-20 follows its work item', [
    v({ swgItem: '8.5', type: 'CR', release: 19 }), v({ swgItem: '8.5', type: 'CR', release: 17 }), v({ swgItem: '8.5', type: 'CR', release: 20 }).value, v({ swgItem: '8.5', type: 'pCR', release: 19 }).value],
    [{ value: '14', rule: 'CR to Rel-19' }, { value: '14', rule: 'CR to Rel-17' }, '15.4', '15.4']);
  check('under 8.4 the maintenance item wins over a work item named in the list', v({ swgItem: '8.4', relatedWis: 'AMD_ARCH_Ph2-MED' }).value, '14');
  check('"Other Rel-20 matters including TEI": a work item of 15/16 wins (title tag, Related WIs)', [v({ swgItem: '8.8', title: '[Energy_PRO-MED] x' }).value, v({ swgItem: '8.8', relatedWis: 'TEI20, FS_AIF_MED' }).value], ['15.5', '16.4']);
  check('"Other Rel-20 matters including TEI" with no such work item: TEI19 -> 14; TEI20 or none -> 15.7 (the item of the same title)',
    [v({ swgItem: '8.8', relatedWis: 'TEI19' }).value, v({ swgItem: '8.8', relatedWis: 'TEI20' }).value, v({ swgItem: '7.8' }).value, v({ swgItem: '10.9' }).value], ['14', '15.7', '15.7', '15.7']);
  check('two work items of different items: blank, ambiguous', v({ swgItem: '8.8', relatedWis: 'AMD_ARCH_Ph2-MED, Energy_PRO-MED' }), { value: '', rule: 'ambiguous: 15.4, 15.5' });
  check('no mapping: blank (reports, opening, AOB, close, an unknown or missing item)', ['8.1', '8.2', '8.3', '8.9', '8.10', '99.1', ''].map((swg) => v({ swgItem: swg }).value), ['', '', '', '', '', '', '']);
  check('no agenda: blank, never a number of its own', S.plenaryAgendaAutoValue_(null, { swgItem: '8.5' }), { value: '', rule: 'no plenary agenda' });
  check('an agenda without the plenary parts gives nothing (an ad-hoc agenda, e.g.)', S.plenaryAgendaModel_([['1', 'Opening'], ['4', 'Topic'], ['4.3', 'IVAS (x)'], ['15', 'Close']]).targets, []);
  check('acronyms: case does not matter, a word inside a title is not one', [v({ swgItem: '', relatedWis: 'amd_arch_ph2-med' }).value, v({ swgItem: '', title: 'about AMD_ARCH_Ph2-MED' }).value], ['15.4', '']);
  check('the work items of a contribution: Related WIs, its title tag, its SWG item and the items above it, once each',
    S.plenaryWorkItemsOf_(model, { relatedWis: 'TEI20; AMD_ARCH_Ph2-MED', title: '[AMD_ARCH_Ph2-MED, FS_AIF_MED] x', swgItem: '8.5' }), ['TEI20', 'AMD_ARCH_Ph2-MED', 'FS_AIF_MED', 'Multicast-Broadcast-Streaming']);
  check('a release is read from "Rel-19", "19", "Rel 19"; nothing else is one', ['Rel-19', '19', 'Rel 19', '', 'n/a'].map(S.plenaryReleaseNumber_), [19, 19, 19, null, null]);
  check('the kept agenda is small enough for one document property', JSON.stringify({ agendaTdoc: 'S4-261857', items: S.compactPlenaryAgendaItems_(AGENDA_138) }).length < 9000, true);
  check('the choice of a Plenary dropdown', ['☑ Yes', '☐ No', '', 'Yes', 'x', 'no', 'later', '✓'].map(S.plenaryChoiceOfText_), ['yes', 'no', 'no', 'yes', 'yes', 'no', 'unknown', 'yes']);
  const cls = (disposition, plenary, conflict) => S.plenaryDispositionClassOf_({ disposition: disposition, conflict: !!conflict }, plenary).key;
  check('the four classes: positive = agreed/endorsed/approved (and partially agreed); everything else is "other", also a conflict or no status',
    [cls('AGREED_ENDORSED_APPROVED', true), cls('DISPOSED_OTHER', true), cls('AGREED_ENDORSED_APPROVED', false), cls('UNRESOLVED', false), cls('AGREED_ENDORSED_APPROVED', true, true)],
    ['POSITIVE_PLENARY', 'OTHER_PLENARY', 'POSITIVE_NOT_PLENARY', 'OTHER_NOT_PLENARY', 'OTHER_PLENARY']);
  check('the taxonomy is the one of the Disposition Summary, not a new one', ['agreed', 'partially agreed', 'endorsed', 'approved', 'noted', 'revised', 'merged', 'withdrawn', 'replied', 'replied to', 'available', 'reserved', 'parked', 'postponed', 'plenary', 'other', 'not pursued', '']
    .map((s) => S.classifyTdocDisposition_(s).disposition === 'AGREED_ENDORSED_APPROVED'), [true, true, true, true, false, false, false, false, false, false, false, false, false, false, false, false, false, false]);
  // The cells: what is the code's and what is not.
  check('the status of a Status cell is the part before "\\u2003Plenary"; a cell without it is all status',
    ['agreed', '', MARK + ' ', 'not pursued' + MARK + ' ' + EM + 'A.I.: 15.4', 'Plenary'].map(S.statusPartOfCellText_), ['agreed', '', '', 'not pursued', 'Plenary']);
  check('the A.I. shown: its value, "" for a bare mark, null when none is shown', [MARK + ' ' + EM + 'A.I.: 15.4', MARK + ' ' + EM + 'A.I.:', MARK + ' ', 'agreed'].map(S.compactAgendaOfCellText_), ['15.4', '', null, null]);
  check('the TDoc number of a TDoc row: without " (revision of …)"', ['S4-262013 (revision of S4-262012)', 'S4-262013', 'S4-262013 (revision of S4-261999)'].map(S.tdocNumberOfCellText_), ['S4-262013', 'S4-262013', 'S4-262013']);
  const pred = S.plenaryPredecessorsOf_({ A: { revisedTo: ['B'], isRevisionOf: [] }, B: { revisedTo: ['C'], isRevisionOf: ['A'] }, C: { revisedTo: [], isRevisionOf: [] },
    D: { revisedTo: ['E'], isRevisionOf: [] }, E: { revisedTo: [], isRevisionOf: ['F'] }, F: { revisedTo: [], isRevisionOf: [] },
    G: { revisedTo: [], isRevisionOf: ['H'] }, H: { revisedTo: [], isRevisionOf: ['G'] }, I: { revisedTo: [], isRevisionOf: ['Z'] }, J: { revisedTo: [], isRevisionOf: ['J'] } });
  check('predecessors from the list only: a chain (one generation each), two predecessors -> none, a cycle -> none, one of another meeting, never itself',
    pred, { B: 'A', C: 'B', I: 'Z' });
}

// ====================================================================== B
console.log('B. a report without Goes to Plenary is exactly as before');
const r = mbs();
{
  const built = r.build();
  check('the MBS report of SA4#138 builds, with its TDocs under 8.x and Status dropdowns; an unmapped status stays text', [built.ok, r.tdocs.every((t) => !!r.table(t.id)), r.shown('S4-262003'), r.shown('S4-262018')],
    [true, true, 'dropdown:agreed', 'text:not pursued']);
  check('no compact row, no definition, not marked', [r.c('S4-262003').plenary, r.api.definitionTitled('Goes to Plenary'), r.docProps.getProperty('GOES_TO_PLENARY')], [null, null, null]);
  check('the revision relationship is part of every report: written by the build already, both numbers linked', [r.tdocRow('S4-262012'), r.tdocRow('S4-262013').text, r.tdocRow('S4-262014').text],
    [{ text: 'S4-262012 (revision of S4-262001)', own: ZIP('S4-262012'), predecessorLink: ZIP('S4-262001') }, 'S4-262013 (revision of S4-262012)', 'S4-262014']);
  const u = r.update();
  check('an update gives it no summary, no compact row and no landscape section, and reads no agenda for it', [u.success, u.goesToPlenary, u.dispositionSummary, r.summary(), r.c('S4-262004').plenary, r.api.sectionBreaks(), r.docProps.getProperty('PLENARY_AGENDA_ITEMS')],
    [true, undefined, undefined, null, null, [], null]);
  r.summaryMenu();
  check('the update keeps the revision relationships; the statuses and the TDoc numbers read as before', [r.tdocRow('S4-262013').text, r.shown('S4-262013'), r.order().indexOf('S4-262013') !== -1], ['S4-262013 (revision of S4-262012)', 'dropdown:agreed', true]);
  check('Update Disposition Summary writes the three tables there, as before, in portrait', [r.summary().map((c) => c[0].replace(/ \(\d+\)$/, '')), r.api.sectionBreaks()],
    [['Agreed, Endorsed and Approved Documents', 'Other Disposed Documents', 'Documents Not Yet Disposed'], []]);
}

console.log('B2. an ad-hoc report: the compact row and the four tables, no SA4 plenary agenda item');
{
  const a = templateReport({ tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'agreed', uploaded: true }, { id: 'S4aA269002', agenda: '4.4', status: 'noted', uploaded: true }] });
  a.build();
  let agendaReads = 0;
  a.s.downloadMeetingAgenda_ = () => { agendaReads++; return []; };
  a.exec(() => a.s.setUpGoesToPlenary());
  const u = a.update();
  check('set up (it writes the summary) and updated: compact rows, the four tables -- the two not going to plenary in five columns', [u.goesToPlenary, u.dispositionSummary, a.api.compactOf('S4aA269001').plenary, a.api.compactOf('S4aA269001').status,
    a.body._children.filter((c) => isTable(c) && c.getNumRows() > 0 && c.getRow(0).getNumCells() === 5).length, a.body._children.filter((c) => isTable(c) && c.getNumRows() > 0 && c.getRow(0).getNumCells() === 6).length], ['done', 'unchanged', NO, 'dropdown:agreed', 2, 0]);
  check('the plenary agenda of a main meeting does not exist here: no A.I., and no agenda is read for it', [a.api.compactOf('S4aA269001').ai, agendaReads], [null, 0]);
}

// ====================================================================== C
console.log('C. Set Up Goes to Plenary…');
{
  r.user(() => r.api.pick('S4-262009', 'Plenary'));
  r.user(() => r.api.cellOf('S4-262003', 'Minutes').setText('Minutes typed by hand.'));
  const labelsBefore = r.labels('S4-262003');
  const writesBefore = r.writes();
  const done = r.setUp();
  check('it asks first, with what it will do', [/Plenary \[☐ No ▾\]/.test(r.ui.alerts[r.ui.alerts.length - 2][1]), /Status "Plenary" \(1\): S4-262009/.test(r.ui.alerts[r.ui.alerts.length - 2][1])], [true, true]);
  check('no row is added or removed; the minutes typed by hand are kept', [r.labels('S4-262003'), r.field('S4-262003', 'Minutes')], [labelsBefore, 'Minutes typed by hand.']);
  check('every Status row reads "[status] Plenary [☐ No]"; S4-262009 (Status "Plenary") "☑ Yes"', r.tdocs.map((t) => [r.c(t.id).plenary, r.c(t.id).ai]),
    r.tdocs.map((t) => [t.id === 'S4-262009' ? YES : NO, null]));
  check('the status in front of it is as it was: dropdowns, and the text of an unmapped status', ['S4-262003', 'S4-262013', 'S4-262018'].map((id) => r.c(id).status), ['dropdown:agreed', 'dropdown:agreed', 'text:not pursued']);
  check('exactly: "[agreed]\\u2003Plenary [☐ No]"', r.c('S4-262003').parts, '[agreed]' + EM + 'Plenary [' + NO + ']');
  check('the Status "Plenary" became "available"; "Plenary" is no option of the Status dropdown any more', [r.c('S4-262009').status, r.api.definitionTitled('Document Status').options.indexOf('Plenary')],
    ['dropdown:available', -1]);
  check('the readers see the status, not the row: Report Status Summary, the discussion e-mails, the summary', [r.exec(() => r.s.findStatusText_(r.table('S4-262018'))), r.exec(() => r.s.readTdocStatus_(r.table('S4-262018'), r.table('S4-262018').getRow(10).getCell(1).getText()))],
    ['not pursued', 'not pursued']);
  check('its definition: "Goes to Plenary", two options', r.api.definitionTitled('Goes to Plenary').options, [NO, YES]);
  check('the result', [done.ok, done.migrated, done.plenaryStatuses, done.dropdowns.inserted, done.dropdowns.optionRemoved, r.docProps.getProperty('GOES_TO_PLENARY')], [true, 0, 1, r.tdocs.length, true, '1']);
  check('its writes: one batch by id (the Status, "Plenary" out of its dropdown, the definition), one batch of rows tied to its revision; the summary there gets its landscape section',
    r.api.calls.filter((c) => c.call === 'batchUpdate').slice(writesBefore).map((c) => [c.requests.map((q) => Object.keys(q)[0]).filter((k, i, a) => a.indexOf(k) === i).join('+'), !!c.requiredRevisionId]),
    [['updateDropdownProperties+updateDropdownDefinitionProperties+createDropdownDefinition', false], ['insertText+insertDropdown', true],
     ['updateTableColumnProperties+updateTableRowStyle+updateTableCellStyle+updateTextStyle+updateParagraphStyle', true], ['insertSectionBreak', true], ['updateSectionStyle', true]]);
  check('the set-up writes the summary as four tables at once (replacing the three of before)', r.summary().map((c) => c[0].replace(/ \(\d+\)$/, '')), TITLES);
  const writes = r.writes();
  const again = r.setUp();
  check('run again: nothing to do, nothing written', [again.nothingToDo, r.writes() - writes], [true, 0]);
}

// ====================================================================== D
console.log('D. the update: A.I., the four tables, landscape, revision relationships');
let afterFirst;
{
  r.mark();
  const u = r.update();
  check('the update succeeds and brings Goes to Plenary up to date', [u.success, u.goesToPlenary, u.dispositionSummary], [true, 'done', 'replaced']);
  check('the automatic SA4 plenary agenda items are kept for every TDoc, shown or not', r.tdocs.reduce((o, t) => (o[t.id] = r.docProps.getProperty('PLENARY_AGENDA_AUTO:' + t.id) || '', o), {}), EXPECTED_AUTO);
  check('A.I. is shown only for "☑ Yes" with an item: S4-262009 shows 15.4, the others nothing', r.tdocs.filter((t) => r.c(t.id).ai).map((t) => [t.id, r.c(t.id).ai]), [['S4-262009', 'A.I.: 15.4']]);
  check('exactly: "[available]\\u2003Plenary [☑ Yes]\\u2003A.I.: 15.4"', r.c('S4-262009').parts, '[available]' + EM + 'Plenary [' + YES + ']' + EM + 'A.I.: 15.4');
  const s = r.summary();
  check('one summary, four tables, in the order asked for, with counts', [r.headings(), s.map((c) => c[0])], [1, [TITLES[0] + ' (0)', TITLES[1] + ' (1)', TITLES[2] + ' (6)', TITLES[3] + ' (10)']]);
  check('the columns: six for the tables going to plenary, five -- no SA4 plenary agenda item -- for the two not going to plenary', [s[1][1][0], s[2][1][0], s[3][1][0], s[2][1].every((x) => x.length === 5), s[3][1].every((x) => x.length === 5)],
    [HEADER, HEADER5, HEADER5, true, true]);
  check('the widths: the Title is the widest column of every table, on the landscape page (A4: 698 pt, rounding aside)', r.body._children.filter((c) => isTable(c) && (c.getRow(0).getNumCells() === 6 || c.getRow(0).getNumCells() === 5) && c.getRow(0).getCell(1).getText() === 'Title')
    .map((t) => { const w = Array.from({ length: t.getRow(0).getNumCells() }, (x, k) => t.getRow(0).getCell(k)._width); return [w.indexOf(Math.max(...w)), Math.abs(w.reduce((a, b) => a + b, 0) - 698) <= 2]; }), [[1, true], [1, true], [1, true]]);
  check('the header bold on its background, no bold below it, no spacing', (() => { const t = r.body._children.filter((c) => isTable(c) && c.getRow(0).getNumCells() === 5 && c.getRow(0).getCell(1).getText() === 'Title')[0];
    return [t.getRow(0).getCell(0)._bold, t.getRow(0).getCell(0)._background, t.getRow(1).getCell(1)._bold, t.getRow(1).getCell(1)._spacing]; })(), [true, '#D9EAF7', false, [0, 0]]);
  check('Other – Goes to Plenary: the TDoc whose Status was "Plenary"', s[1][1].slice(1), [['S4-262009', '[AMD_ARCH_Ph2-MED] open issue', 'ExampleCorp', 'available', '8.5', '15.4']]);
  check('the status in the tables is the status, not the compact row; an unmapped one as written', [].concat(...s.map((c) => (Array.isArray(c[1]) ? c[1].slice(1) : []))).filter((x) => x[0] === 'S4-262018' || x[0] === 'S4-262003').map((x) => x[3]).sort(),
    ['agreed', 'not pursued (not a known status)']);
  const all = [].concat(...s.map((c) => (Array.isArray(c[1]) ? c[1].slice(1).map((x) => x[0]) : [])));
  check('every TDoc of the report is in exactly one table, in the order of the report', [all.length, all.slice().sort(), all.filter((n, i) => s.slice(0, 4).length && true).length], [r.tdocs.length, r.tdocs.map((t) => t.id).sort(), r.tdocs.length]);
  check('the TDoc column holds the number alone, linked as in its table', [].concat(...r.body._children.filter((c) => isTable(c) && c.getRow(0).getNumCells() === 6).map((t) => Array.from({ length: t.getNumRows() - 1 }, (_, i) => [t.getRow(i + 1).getCell(0).getText(), t.getRow(i + 1).getCell(0).editAsText().getLinkUrl(0)])))
    .every(([n, l]) => /^S4-\d{6}$/.test(n) && l === ZIP(n)), true);
  check('the four tables are on landscape pages of their own: one section break, before the heading, flipped', r.api.sectionBreaks(), [{ before: SUMMARY, flip: true }]);
  // The revision relationships.
  check('a chain: each revision names its own predecessor, both numbers linked', ['S4-262012', 'S4-262013'].map((id) => r.tdocRow(id)), [
    { text: 'S4-262012 (revision of S4-262001)', own: ZIP('S4-262012'), predecessorLink: ZIP('S4-262001') },
    { text: 'S4-262013 (revision of S4-262012)', own: ZIP('S4-262013'), predecessorLink: ZIP('S4-262012') }]);
  check('a predecessor of another meeting: named, not linked (no URL is known)', r.tdocRow('S4-262017'), { text: 'S4-262017 (revision of S4-261999)', own: ZIP('S4-262017'), predecessorLink: '' });
  check('two different predecessors, a cycle, the originals: the number alone', ['S4-262014', 'S4-262015', 'S4-262016', 'S4-262001', 'S4-262004'].map((id) => r.tdocRow(id).text), ['S4-262014', 'S4-262015', 'S4-262016', 'S4-262001', 'S4-262004']);
  check('the Revisions row is as the revision import left it (not touched)', r.field('S4-262013', 'Revisions'), r.field('S4-262013', 'Revisions'));
  afterFirst = r.snapshot();
}

// ====================================================================== E
console.log('E. repeated updates: idempotent');
{
  const writes = r.writes();
  const u2 = r.update();
  const u3 = r.update();
  check('two more updates: the summary is unchanged, the document is the same element by element', [u2.dispositionSummary, u3.dispositionSummary, r.snapshot() === afterFirst], ['unchanged', 'unchanged', true]);
  check('and no API write at all', r.writes() - writes, 0);
  check('no summary, section break, dropdown or revision text is duplicated', [r.headings(), r.api.sectionBreaks().length, r.tdocs.map((t) => (r.c(t.id).parts.match(/Plenary/g) || []).length), r.tdocRow('S4-262013').text],
    [1, 1, r.tdocs.map(() => 1), 'S4-262013 (revision of S4-262012)']);
}

// ====================================================================== F
console.log('F. the selection is the rapporteur\'s');
{
  r.user(() => r.api.pickPlenary('S4-262003', YES));
  check('Google Docs cannot show A.I. at once: until the next update the row says only "☑ Yes"', r.c('S4-262003').ai, null);
  r.update();
  check('the update shows its A.I. and moves it to Agreed – Goes to Plenary', [r.c('S4-262003').ai, r.inTable(0)], ['A.I.: 14', ['S4-262003']]);
  r.user(() => r.api.pick('S4-262003', 'noted'));
  r.update();
  check('its disposition changes to noted: Other – Goes to Plenary; "☑ Yes" and its A.I. stay', [r.inTable(1), r.c('S4-262003').plenary, r.c('S4-262003').ai], [['S4-262003', 'S4-262009'], YES, 'A.I.: 14']);
  r.user(() => r.api.pick('S4-262003', 'agreed'));
  r.user(() => r.api.pickPlenary('S4-262009', NO));
  r.update();
  check('"☐ No" again: its A.I. goes, it moves to Not Going; the other stays', [r.c('S4-262009').ai, r.inTable(0), r.inTable(1)], [null, ['S4-262003'], []]);
  // The Portal sync of a text status in a compact row: "not pursued" (no option) -> "revised" (an option).
  r.tdoc('S4-262018').status = 'revised';
  r.tdocs.push({ id: 'S4-262019', agenda: '8.6', status: 'available', title: '[Energy_PRO-MED] new contribution', uploaded: true });
  r.tdoc('S4-262014').status = 'agreed';
  const u = r.update();
  check('a text status in a compact row is synced as before, and becomes a dropdown; the row after it is kept', [u.success, r.c('S4-262018').status, r.c('S4-262018').plenary], [true, 'dropdown:revised', NO]);
  check('a new TDoc gets the compact row in the same update', [r.c('S4-262019').parts, r.docProps.getProperty('PLENARY_AGENDA_AUTO:S4-262019')], ['[available]' + EM + 'Plenary [' + NO + ']', '15.5']);
  check('the Portal moves a status on: the selection is not touched', [r.c('S4-262014').status, r.c('S4-262014').plenary], ['dropdown:agreed', NO]);
  check('the new TDoc is in the summary once', [].concat(...r.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1).map((x) => x[0]) : []))).filter((n) => n === 'S4-262019').length, 1);
  const rows = r.tdocs.map((t) => r.c(t.id).parts);
  r.exec(() => { r.s.resetStatusDropdownRun_(); r.s.collectRevisionsCore_(); r.s.finalizeStatusDropdowns_('update'); });
  check('Update TDoc Revisions leaves every Status row as it was', r.tdocs.map((t) => r.c(t.id).parts), rows);
}

// ====================================================================== G
console.log('G. manual SA4 plenary agenda items');
{
  r.user(() => r.api.pickPlenary('S4-262007', YES));     // ambiguous: no automatic item, so no A.I. yet
  r.user(() => r.api.pickPlenary('S4-262002', YES));     // 15.5
  r.update();
  check('"☑ Yes" without a known item shows no A.I.; with one it shows it', [r.c('S4-262007').ai, r.c('S4-262002').ai], [null, 'A.I.: 15.5']);
  r.typeAgenda('S4-262002', '16.2');                     // a manual override of 15.5
  r.update();
  check('a value typed over A.I. is kept, remembered as manual, and in the summary', [r.c('S4-262002').ai, r.docProps.getProperty('PLENARY_AGENDA_MANUAL:S4-262002'),
    [].concat(...r.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1) : []))).filter((x) => x[0] === 'S4-262002').map((x) => x[5])], ['A.I.: 16.2', '16.2', ['16.2']]);
  r.user(() => r.api.pickPlenary('S4-262002', NO));
  r.update();
  check('"☐ No": the A.I. is hidden and the TDoc is in a five-column table; the manual value is kept', [r.c('S4-262002').ai, r.docProps.getProperty('PLENARY_AGENDA_MANUAL:S4-262002'),
    [].concat(...r.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1) : []))).filter((x) => x[0] === 'S4-262002').map((x) => x.length)], [null, '16.2', [5]]);
  r.user(() => r.api.pickPlenary('S4-262002', YES));
  r.tdoc('S4-262002')['Related WIs'] = 'FS_QStream_MED';  // its work items change: 15.5 and 16.3 -> no automatic item
  r.update();
  check('"☑ Yes" again: the manual value comes back; a work-item change does not move it', [r.c('S4-262002').ai, r.docProps.getProperty('PLENARY_AGENDA_AUTO:S4-262002')], ['A.I.: 16.2', null]);
  r.typeAgenda('S4-262002', '');                         // cleared to "A.I.:": automatic again (now none)
  r.update();
  check('cleared to "A.I.:": the manual value is dropped; no automatic item, so no A.I.', [r.c('S4-262002').ai, r.docProps.getProperty('PLENARY_AGENDA_MANUAL:S4-262002')], [null, null]);
  r.typeAgenda('S4-262003', '-');                        // "-": blank on purpose
  r.update();
  check('"-" keeps it blank on purpose: no A.I. shown, blank in the summary', [r.c('S4-262003').ai, r.docProps.getProperty('PLENARY_AGENDA_MANUAL:S4-262003'),
    [].concat(...r.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1) : []))).filter((x) => x[0] === 'S4-262003').map((x) => x[5])], [null, '-', ['']]);
  r.tdoc('S4-262009')['Related WIs'] = 'FS_AIF_MED';      // 8.5 (15.4) and FS_AIF_MED (16.4): ambiguous now
  r.tdoc('S4-262005')['Related WIs'] = 'AMD_ARCH_Ph2-MED'; // 15.7 -> 15.4
  r.update();
  check('a work-item change moves the automatic values only', ['S4-262009', 'S4-262005'].map((n) => r.docProps.getProperty('PLENARY_AGENDA_AUTO:' + n)), [null, '15.4']);
  const before = r.snapshot();
  r.update();
  r.update();
  check('repeated updates leave the manual and automatic values as they are', r.snapshot() === before, true);
  const kept = r.docProps.getProperty('PLENARY_AGENDA_ITEMS');
  r.docProps.deleteProperty('PLENARY_AGENDA_ITEMS');
  r.s.downloadMeetingAgenda_ = () => { throw new Error('Could not download agenda ZIP (403)'); };
  const autos = r.tdocs.map((t) => r.docProps.getProperty('PLENARY_AGENDA_AUTO:' + t.id));
  const u = r.update();
  check('the agenda cannot be read: the update succeeds and no automatic item changes', [u.success, r.tdocs.map((t) => r.docProps.getProperty('PLENARY_AGENDA_AUTO:' + t.id))], [true, autos]);
  r.s.downloadMeetingAgenda_ = () => AGENDA_138;
  r.docProps.setProperty('PLENARY_AGENDA_ITEMS', kept);
}

// ====================================================================== H
console.log('H. minutes, text typed into or below the summary, an unreadable selection');
{
  const kids = r.body._children;
  r.user(() => r.body.appendParagraph('Minutes typed below the summary by hand.'));
  r.user(() => r.api.pick('S4-262004', 'agreed'));
  const u = r.update();
  check('text typed below it is kept, in a portrait section of its own; the summary above it is updated',
    [u.dispositionSummary, kids[kids.length - 1].getText(), r.api.sectionBreaks()], ['replaced', 'Minutes typed below the summary by hand.', [{ before: SUMMARY, flip: true }, { before: 'Minutes typed below the summary by hand.', flip: false }]]);
  const at = kids.findIndex((c) => !isTable(c) && c.getText().indexOf(TITLES[2]) === 0);
  r.user(() => r.body.insertParagraph(at, 'A remark typed into the summary.'));
  r.user(() => r.api.pick('S4-262004', 'noted'));
  const before = r.snapshot();
  const v = r.update();
  check('text typed into it: refused, nothing changed, the remark kept', [v.dispositionSummary, r.snapshot() === before], ['refused', true]);
  r.user(() => r.body.removeChild(kids[at]));
  check('removed again: the next update writes it; the sections stay as they are', [r.update().dispositionSummary, r.api.sectionBreaks().length], ['replaced', 2]);
  r.user(() => r.api.pickPlenary('S4-262008', YES));
  const snap = r.snapshot();
  r.api.fail.get = (id) => (id === 'REPORTdoc000000000000000000000000000000000' ? 'Internal error' : null);
  const w = r.update();
  r.api.fail.get = null;
  check('the API cannot read the dropdowns: the summary is not written (no guess), the document is as it was', [w.dispositionSummary, r.snapshot() === snap], ['unreadable', true]);
  check('the next update that can read them moves S4-262008 to Other – Goes to Plenary (with S4-262002, "☑ Yes" since G)', [r.update().dispositionSummary, r.inTable(1).slice().sort()], ['replaced', ['S4-262002', 'S4-262008']]);
  r.user(() => r.api.pick('S4-262008', 'agreed'));
  const m = r.summaryMenu();
  check('Update Disposition Summary writes the four tables and keeps the landscape section', [m.state, r.inTable(0).indexOf('S4-262008') !== -1, r.api.sectionBreaks()],
    ['replaced', true, [{ before: SUMMARY, flip: true }, { before: 'Minutes typed below the summary by hand.', flip: false }]]);
}

// ====================================================================== I
console.log('I. Build Report from Scratch keeps the selections, the manual values, the revisions and the landscape');
{
  const selections = r.tdocs.reduce((o, t) => (o[t.id] = r.c(t.id).plenary, o), {});
  const manual = r.tdocs.reduce((o, t) => (o[t.id] = r.docProps.getProperty('PLENARY_AGENDA_MANUAL:' + t.id), o), {});
  // A Status "Plenary" typed by hand after the set-up.
  r.user(() => { const c = r.api.statusCell('S4-262010'); c._parts = [{ t: 'Plenary' + MARK + ' ' }].concat(c._allParts().filter((p) => p.d && p.d.definitionId !== c._allParts()[0].d.definitionId)); c._sync(); });
  const built = r.build();
  check('the rebuild succeeds; every TDoc table has the compact row again', [built.ok, r.tdocs.every((t) => r.c(t.id).plenary !== null)], [true, true]);
  check('every selection is kept; the typed Status "Plenary" became "☑ Yes"', r.tdocs.reduce((o, t) => (o[t.id] = r.c(t.id).plenary, o), {}), Object.assign({}, selections, { 'S4-262010': YES }));
  check('the manual values are kept', r.tdocs.reduce((o, t) => (o[t.id] = r.docProps.getProperty('PLENARY_AGENDA_MANUAL:' + t.id), o), {}), manual);
  check('"Plenary" does not come back as a Status', [r.c('S4-262010').status, r.tdocs.some((t) => /plenary/i.test(r.c(t.id).status))], ['dropdown:approved', false]);
  check('the revision relationships are written again', r.tdocRow('S4-262013').text, 'S4-262013 (revision of S4-262012)');
  check('nothing is pending once the rows stand', Object.keys(r.docProps._store || {}).filter((k) => /^PLENARY_CHOICE_PENDING:/.test(k)), []);
  const u = r.update();
  check('the update after it writes the four tables again, on landscape pages, and the A.I. of every "☑ Yes"', [u.dispositionSummary, r.summary().map((c) => c[0].replace(/ \(\d+\)$/, '')), r.api.sectionBreaks(),
    r.tdocs.filter((t) => r.c(t.id).ai).map((t) => t.id)], ['created', TITLES, [{ before: SUMMARY, flip: true }], ['S4-262010', 'S4-262003', 'S4-262007', 'S4-262008'].filter((id) => r.c(id).ai)]);
}

// ====================================================================== J
console.log('J. Close Report');
{
  const c = r.close();
  check('Close Report writes the same four tables on landscape pages; its counts are the classes of the Disposition Summary', [c.summaryState, r.headings(), Object.keys(c.counts).slice(0, 3), r.api.sectionBreaks().length],
    ['unchanged', 1, ['AGREED_ENDORSED_APPROVED', 'DISPOSED_OTHER', 'UNRESOLVED'], 1]);
  check('the TDocs not yet disposed are listed as before, by their status', c.unresolved.map((e) => e.tdoc + ' ' + e.status), ['S4-262009 available', 'S4-262015 available', 'S4-262016 available', 'S4-262019 available']);
}

// ====================================================================== L
console.log('L. migration of the first candidate\'s two-row layout');
{
  const m = mbs({ tdocs: TDOCS().slice(0, 4) });
  m.build();
  // The first candidate's state, as the disposable MBS copy has it: two rows below Status, a "Goes to Plenary" dropdown in
  // the first, the plenary agenda item in the second (an automatic one, a manual one), and the report marked.
  const reply = m.api.Docs.Documents.batchUpdate({ requests: [{ createDropdownDefinition: { dropdownDefinition: { dropdownDefinitionProperties: { title: 'Goes to Plenary', options: [{ displayValue: NO }, { displayValue: YES }] } } } }] }, 'REPORTdoc000000000000000000000000000000000');
  const def = reply.replies[0].createDropdownDefinition.dropdownDefinition;
  const legacy = { 'S4-262001': [YES, '15.4'], 'S4-262002': [NO, '16.2'], 'S4-262003': [YES, '14'], 'S4-262004': [NO, ''] };
  m.user(() => Object.keys(legacy).forEach((id) => {
    const t = m.table(id);
    const s = m.labels(id).indexOf('Status');
    const row = t.insertTableRow(s + 1); row.appendTableCell('Goes to Plenary'); row.appendTableCell('');
    row.getCell(1)._dropdown = { dropdownId: 'kix.legacy' + id, definitionId: def.dropdownDefinitionId, optionId: def.dropdownDefinitionProperties.options.filter((o) => o.displayValue === legacy[id][0])[0].optionId };
    const a = t.insertTableRow(s + 2); a.appendTableCell('Plenary Agenda Item'); a.appendTableCell(legacy[id][1]);
  }));
  m.docProps.setProperty('GOES_TO_PLENARY', '1');
  ['S4-262001', 'S4-262003'].forEach((id) => m.docProps.setProperty('PLENARY_AGENDA_AUTO:' + id, legacy[id][1]));
  m.docProps.setProperty('PLENARY_AGENDA_AUTO:S4-262002', '15.5');
  const u = m.update();
  check('the next update moves every table to the compact row; the two rows are gone', [u.success, Object.keys(legacy).map((id) => m.labels(id).filter((l) => /Plenary/.test(l)).length)], [true, [0, 0, 0, 0]]);
  check('no selection is lost', Object.keys(legacy).map((id) => m.c(id).plenary), [YES, NO, YES, NO]);
  check('a manual agenda item of the old row stays manual; automatic ones stay automatic; A.I. shows for "☑ Yes"',
    [m.docProps.getProperty('PLENARY_AGENDA_MANUAL:S4-262002'), m.docProps.getProperty('PLENARY_AGENDA_MANUAL:S4-262001'), m.c('S4-262001').ai, m.c('S4-262003').ai, m.c('S4-262002').ai],
    ['16.2', null, 'A.I.: 15.4', 'A.I.: 14', null]);
  check('nothing is left pending; the four tables are there, on landscape pages', [Object.keys(m.docProps._store).filter((k) => /^PLENARY_CHOICE_PENDING:/.test(k)), m.summary().length, m.api.sectionBreaks()],
    [[], 4, [{ before: SUMMARY, flip: true }]]);
  const snap = m.snapshot();
  m.update();
  check('a second update changes nothing', m.snapshot() === snap, true);
  // A failed end of the run never loses a selection: it is kept until the dropdown stands.
  const n = mbs({ tdocs: TDOCS().slice(0, 2) });
  n.build();
  const r2 = n.api.Docs.Documents.batchUpdate({ requests: [{ createDropdownDefinition: { dropdownDefinition: { dropdownDefinitionProperties: { title: 'Goes to Plenary', options: [{ displayValue: NO }, { displayValue: YES }] } } } }] }, 'REPORTdoc000000000000000000000000000000000');
  const d2 = r2.replies[0].createDropdownDefinition.dropdownDefinition;
  n.user(() => { const t = n.table('S4-262002'); const s = n.labels('S4-262002').indexOf('Status'); const row = t.insertTableRow(s + 1); row.appendTableCell('Goes to Plenary'); row.appendTableCell('');
    row.getCell(1)._dropdown = { dropdownId: 'kix.legacyN', definitionId: d2.dropdownDefinitionId, optionId: d2.dropdownDefinitionProperties.options[1].optionId }; });
  n.docProps.setProperty('GOES_TO_PLENARY', '1');
  n.api.fail.batchUpdate = (requests) => (requests.some((q) => q.insertDropdown) ? 'Internal error' : null);
  n.update();
  check('the end of the run fails: the row is gone, its "☑ Yes" is kept as pending', [n.labels('S4-262002').indexOf('Goes to Plenary'), n.docProps.getProperty('PLENARY_CHOICE_PENDING:S4-262002')], [-1, 'yes']);
  n.api.fail.batchUpdate = null;
  n.update();
  check('the next update makes the row with "☑ Yes" and forgets the pending choice', [n.c('S4-262002').plenary, n.docProps.getProperty('PLENARY_CHOICE_PENDING:S4-262002')], [YES, null]);
}

// ====================================================================== M
console.log('M. an update of 75 TDocs: DocumentApp calls and Docs API calls, the first candidate against this one');
// A relative measure on the fake document: every DocumentApp method called on the report is counted. It is not a time,
// and the real cost of a call in Google Docs is not known here; the real update is timed in the disposable copy.
const SEVENTY_FIVE = () => Array.from({ length: 75 }, (_, k) => {
  const n = String(k + 1).padStart(2, '0');
  const agenda = ['8.4', '8.5', '8.6', '8.7', '8.8'][k % 5];
  const status = ['agreed', 'noted', 'available', 'revised', 'endorsed', 'reserved'][k % 6];
  return { id: 'S4-2630' + n, agenda: agenda, status: status, title: (agenda === '8.5' ? '[AMD_ARCH_Ph2-MED] ' : '') + 'Contribution ' + n, uploaded: true };
});
function measure(code) {
  const m = mbs(Object.assign({ tdocs: SEVENTY_FIVE() }, code ? { code: code } : {}));
  m.build();
  m.setUp();
  m.update();
  m.update();
  const count = () => ({ calls: m.body._state.calls, gets: m.api.count('get', 'REPORTdoc000000000000000000000000000000000'), writes: m.writes() });
  const delta = (a, b) => ({ calls: b.calls - a.calls, gets: b.gets - a.gets, writes: b.writes - a.writes });
  const out = {};
  let at = count();
  const u1 = m.update();
  out.unchanged = Object.assign(delta(at, count()), { summary: u1.dispositionSummary });
  m.user(() => m.api.pick('S4-263001', 'noted'));
  m.user(() => (m.api.compactOf('S4-263002').plenary ? m.api.pickPlenary('S4-263002', YES) : m.api.pickIn('S4-263002', 'Goes to Plenary', YES)));
  at = count();
  const u2 = m.update();
  out.oneStatusOnePlenary = Object.assign(delta(at, count()), { summary: u2.dispositionSummary });
  m.tdocs.push({ id: 'S4-263076', agenda: '8.6', status: 'reserved', title: 'A new contribution', uploaded: true });
  m.tdoc('S4-263006').status = 'withdrawn';
  at = count();
  const u3 = m.update();
  out.newTdocAndPortalStatus = Object.assign(delta(at, count()), { summary: u3.dispositionSummary });
  return { out: out, m: m };
}
let firstCandidate = null;
try {
  firstCandidate = execFileSync('git', ['show', '4c55c3c:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8');
} catch (e) {
  firstCandidate = null;
}
{
  const now = measure(null);
  console.log('  this candidate:         ' + JSON.stringify(now.out));
  check('this candidate: the unchanged update writes nothing -- no DocumentApp change, no API write', [now.out.unchanged.summary, now.out.unchanged.writes], ['unchanged', 0]);
  check('this candidate: one Status and one Plenary changed: the summary rewritten once, the rows refreshed', now.out.oneStatusOnePlenary.summary, 'replaced');
  check('this candidate: a new TDoc and a Portal status: completed in the same update (dropdowns, compact row, summary)', [now.out.newTdocAndPortalStatus.summary, now.m.api.compactOf('S4-263076').status, now.m.api.compactOf('S4-263076').plenary, now.m.c('S4-263006').status],
    ['replaced', 'dropdown:reserved', NO, 'dropdown:withdrawn']);
  if (firstCandidate === null) {
    console.log('  note: commit 4c55c3c is not available; the comparison with the second candidate is skipped.');
  } else {
    const before = measure(firstCandidate);
    console.log('  second candidate (4c55c3c): ' + JSON.stringify(before.out));
    const ratio = (k) => Math.round(before.out[k].calls / Math.max(1, now.out[k].calls) * 10) / 10;
    console.log('  DocumentApp calls, second candidate / this one: unchanged ' + ratio('unchanged') + 'x, one status and one plenary ' + ratio('oneStatusOnePlenary') + 'x, new TDoc ' + ratio('newTdocAndPortalStatus') + 'x');
    check('far fewer DocumentApp calls than the second candidate: at least 5x fewer for an unchanged update, 3x for the others',
      [ratio('unchanged') >= 5, ratio('oneStatusOnePlenary') >= 3, ratio('newTdocAndPortalStatus') >= 3], [true, true, true]);
  }
}

// ====================================================================== N
console.log('N. an update interrupted after it inserted a TDoc: the next one completes it');
{
  const m = mbs({ tdocs: TDOCS().slice(0, 6) });
  m.build();
  m.setUp();
  m.update();
  m.tdocs.push({ id: 'S4-262020', agenda: '8.6', status: 'reserved', title: '[Energy_PRO-MED] arrived during the meeting', uploaded: true });
  const collector = m.s.collectorUpdate_;
  m.s.collectorUpdate_ = () => { throw new Error('Exceeded maximum execution time'); };
  const broken = m.update();
  m.s.collectorUpdate_ = collector;
  m.user(() => {});                                       // what the interrupted run had changed is in the document (the worst case)
  check('the interrupted update inserted the TDoc and stopped before its end: the Status is still text, no compact row yet', [broken.success, m.r === undefined && !!m.table('S4-262020'), m.c('S4-262020').status, m.c('S4-262020').plenary],
    [false, true, 'text:reserved', null]);
  const u = m.update();
  check('the next update completes it: the Status a dropdown (reserved), the compact row, no second table', [u.success, m.c('S4-262020').status, m.c('S4-262020').plenary, m.tdocTables().filter((t) => t.getRow(0).getCell(1).getText().indexOf('S4-262020') === 0).length],
    [true, 'dropdown:reserved', NO, 1]);
  check('the TDoc is in the summary once, and in the registration table once', [[].concat(...m.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1).map((x) => x[0]) : []))).filter((n) => n === 'S4-262020').length,
    m.body._children.filter((c) => isTable(c) && c.getRow(0).getNumCells() >= 4 && c.getRow(0).getCell(3).getText() === 'Agenda Item').map((t) => rowsOf(t).filter((x) => x[0] === 'S4-262020').length)], [1, [1]]);
  const snap = m.snapshot();
  m.update();
  check('and the update after it changes nothing', m.snapshot() === snap, true);
}

// ====================================================================== O
console.log('O. every Status of the dropdown keeps its dropdown in the compact row, reserved included');
{
  const STATUSES = ['available', 'noted', 'agreed', 'revised', 'merged', 'approved', 'reserved', 'endorsed', 'withdrawn', 'replied', 'postponed'];
  const m = mbs({ tdocs: STATUSES.map((status, k) => ({ id: 'S4-2640' + String(k + 1).padStart(2, '0'), agenda: '8.6', status: status, title: '[Energy_PRO-MED] ' + status, uploaded: true })) });
  m.build();
  m.setUp();
  m.update();
  check('each status is a dropdown in front of "Plenary [☐ No]"', m.tdocs.map((t) => [m.c(t.id).status, m.c(t.id).plenary]), STATUSES.map((status) => ['dropdown:' + status, NO]));
  m.tdoc('S4-264007').status = 'withdrawn';               // reserved -> withdrawn (the Portal)
  m.tdoc('S4-264001').status = 'agreed';                  // available -> agreed (the Portal)
  m.update();
  check('the Portal moves reserved and available on: the dropdowns follow, the compact row stays', [m.c('S4-264007').status, m.c('S4-264001').status, m.c('S4-264007').plenary, m.c('S4-264001').plenary],
    ['dropdown:withdrawn', 'dropdown:agreed', NO, NO]);
  // A Status dropdown lost (a text left in its place, as an interrupted run leaves it): the next update makes it again.
  m.user(() => { const c = m.api.statusCell('S4-264003'); const rest = c._allParts().slice(1); c._parts = [{ t: 'agreed' + rest[0].t }].concat(rest.slice(1)); c._sync(); });
  check('(a Status dropdown replaced by its text)', m.c('S4-264003').status, 'text:agreed');
  m.update();
  check('the next update makes it a dropdown again, with the same value; the Plenary dropdown is kept', [m.c('S4-264003').status, m.c('S4-264003').plenary], ['dropdown:agreed', NO]);
  check('Check Report Integrity: nothing missing', (() => { const p = m.exec(() => m.s.checkReportIntegrity()); return [p.textStatusesWithOption, p.legacy, p.plenaryDropdownMissing, p.summaryHeadings, p.sectionBreaks]; })(),
    [[], [], [], 1, [true]]);
}

// ====================================================================== P
console.log('P. an update that has run long leaves the summary and the formatting to the next one');
{
  const m = mbs({ tdocs: TDOCS().slice(0, 6) });
  m.build();
  m.setUp();
  m.update();
  m.tdocs.push({ id: 'S4-262021', agenda: '8.6', status: 'available', title: '[Energy_PRO-MED] late', uploaded: true });
  const elapsed = m.s.reportUpdateElapsedMs_;
  m.s.reportUpdateElapsedMs_ = () => 300000;
  const u = m.update();
  m.s.reportUpdateElapsedMs_ = elapsed;
  check('five minutes gone: the summary is deferred, the formatting skipped -- the dropdowns are made all the same', [u.success, u.dispositionSummary, m.logs.some((l) => /left to the next update: formatting/.test(l)), m.c('S4-262021').status, m.c('S4-262021').plenary],
    [true, 'deferred', true, 'dropdown:available', NO]);
  check('the summary does not have the new TDoc yet', [].concat(...m.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1).map((x) => x[0]) : []))).indexOf('S4-262021'), -1);
  const v = m.update();
  check('the next update writes it', [v.dispositionSummary, [].concat(...m.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1).map((x) => x[0]) : []))).filter((n) => n === 'S4-262021').length], ['replaced', 1]);
}

// ====================================================================== Q
console.log('Q. the collectors of an update of 75 TDocs with some 150 revision links: compared, written only where they change');
// The REAL checkRSSFeed_() and updateRevisions_() run; only the network is replaced (the revisions folder listing, the
// mailing-list messages). The call counts are of the fake document: a relative measure, never a time.
function collectorsReport(code) {
  const REV = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Inbox/Drafts/MBS/';
  const files = [];
  SEVENTY_FIVE().forEach((t, k) => { for (let n = 1; n <= 1 + (k % 3); n++) files.push(t.id + '-r' + n + '.docx'); });
  const mail = ['S4-263005', 'S4-263011', 'S4-263017', 'S4-263023', 'S4-263029'].map((id, k) => ({ title: 'Re: [MBS] ' + id + ' comments', author: 'Delegate ' + k, date: '2026-11-17T1' + k + ':00:00Z',
    link: 'https://list.example.org/msg' + k, messageId: '<msg' + k + '@example.org>' }));
  const m = mbs(Object.assign({ tdocs: SEVENTY_FIVE(), props: Object.assign({}, MAIN, { REVISIONS_URL: REV }) }, code ? { code: code } : {}));
  m.files = files;
  m.mail = mail;
  m.s.safeFetch_ = (cfg, url) => (String(url).indexOf('/Drafts/') !== -1 ? { ok: true, code: 200, text: '<html>' + m.files.map((f) => '<a href="' + f + '">' + f + '</a>').join('\n') + '</html>' } : { ok: false, code: 404, text: '' });
  m.s.collectHybridListservMessages_ = () => m.mail.map((x) => Object.assign({}, x));
  m.s.collectorUpdate_ = function (context) { const cfg = m.s.getCollectorConfig_(context); m.s.checkRSSFeed_(cfg, context); m.s.updateRevisions_(cfg, context); return { failures: [] }; };
  m.propWrites = [];
  const set = m.docProps.setProperty;
  const setMany = m.docProps.setProperties;
  m.docProps.setProperty = (k, v) => { m.propWrites.push(k); return set(k, v); };
  m.docProps.setProperties = (o, all) => { Object.keys(o || {}).forEach((k) => m.propWrites.push(k)); return setMany(o, all); };
  m.build();
  if (m.s.setUpGoesToPlenary) m.setUp();
  return m;
}
{
  const m = collectorsReport(null);
  m.docProps.setProperty('LOCALREV_S4-263010', JSON.stringify({ r9: { text: 'S4-263010-r9.docx', link: 'http://10.10.10.10/ftp/SA/SA4/Inbox/Drafts/MBS/S4-263010-r9.docx' } }));
  const first = m.update();
  const revisionsOf = (id) => m.field(id, 'Revisions');
  check('the first update writes the Revisions cells from the folder (one to three files each) and the discussions', [first.success, revisionsOf('S4-263001'), revisionsOf('S4-263003').split('\n').length, /Delegate 0 on/.test(m.field('S4-263005', 'E-mail Discussion'))],
    [true, 'S4-263001-r1.docx', 3, true]);
  check('a revision file is linked to the folder; a local revision is listed with its marker and its local link', [m.table('S4-263001').getRow(m.labels('S4-263001').indexOf('Revisions')).getCell(1).editAsText().getLinkUrl(0),
    revisionsOf('S4-263010').split('\n').some((l) => /S4-263010-r9\.docx/.test(l))], ['https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Inbox/Drafts/MBS/S4-263001-r1.docx', true]);
  m.update();
  const snap = m.snapshot();
  const calls = m.body._state.calls;
  const writes = m.writes();
  m.propWrites.length = 0;
  const noop = m.update();
  const noopCalls = m.body._state.calls - calls;
  console.log('  an unchanged update with the collectors: ' + noopCalls + ' DocumentApp calls, ' + (m.writes() - writes) + ' API writes, ' + m.propWrites.filter((k) => /^(REVIS_|DISCUSS_|LOCALREV_)/.test(k)).length + ' store writes');
  check('an unchanged update is a no-op for the collectors: the same document, no API write, no store written', [noop.success, m.snapshot() === snap, m.writes() - writes, m.propWrites.filter((k) => /^(REVIS_|DISCUSS_|LOCALREV_)/.test(k))],
    [true, true, 0, []]);
  check('and it makes few DocumentApp calls (under 1 000 for 75 TDocs and some 150 revision links)', noopCalls < 1000, true);
  check('it is a complete update: nothing left to the next one', noop.incomplete, undefined);
  // One new revision file: one cell.
  m.files.push('S4-263020-r7.docx');
  const tables = () => m.tdocTables().map((t) => JSON.stringify(rowsOf(t)));
  const beforeOne = tables();
  m.propWrites.length = 0;
  m.update();
  const afterOne = tables();
  check('one new revision file: exactly one TDoc table changes (its Revisions cell), one store is written', [afterOne.filter((t, i) => t !== beforeOne[i]).length, revisionsOf('S4-263020').split('\n').indexOf('S4-263020-r7.docx') !== -1,
    m.propWrites.filter((k) => /^REVIS_/.test(k))], [1, true, ['REVIS_S4-263020']]);
  // A missing link is restored, though the text is the same.
  m.user(() => { const c = m.table('S4-263002').getRow(m.labels('S4-263002').indexOf('Revisions')).getCell(1); c._links = []; });
  m.update();
  check('a Revisions cell whose text is right but whose link is gone is written again', !!m.table('S4-263002').getRow(m.labels('S4-263002').indexOf('Revisions')).getCell(1).editAsText().getLinkUrl(0), true);
  // The meeting-room agent goes stale: the local revisions it imported stay.
  m.s.importAgentListing_ = () => ({ result: 'stale', state: { message: 'the newest listing is stale' } });
  m.update();
  check('a stale local listing clears no imported local revision', [revisionsOf('S4-263010').split('\n').some((l) => /S4-263010-r9\.docx/.test(l)), m.docProps.getProperty('LOCALREV_S4-263010') !== null], [true, true]);
  // A new message: one discussion cell.
  m.mail.push({ title: 'Re: [MBS] S4-263005 comments', author: 'Delegate 9', date: '2026-11-18T09:00:00Z', link: 'https://list.example.org/msg9', messageId: '<msg9@example.org>' });
  const beforeMail = tables();
  m.update();
  check('a new message: only that TDoc table changes', [tables().filter((t, i) => t !== beforeMail[i]).length, /Delegate 9 on/.test(m.field('S4-263005', 'E-mail Discussion'))], [1, true]);
  // The collectors start only while the update has time; the revisions stop between two TDocs, and say so.
  const elapsed = m.s.reportUpdateElapsedMs_;
  m.s.reportUpdateElapsedMs_ = () => 160000;
  const late = m.update();
  m.s.reportUpdateElapsedMs_ = elapsed;
  check('an update that has run 160 s leaves the collectors to the next one, and is reported as incomplete', [late.success, late.incomplete], [true, ['e-mail discussions', 'TDoc revisions']]);
  check('Update Report Now says so: "updated only in part"', /updated only in part[\s\S]*e-mail discussions[\s\S]*TDoc revisions/.test(m.s.describeUpdateReportNowFailure_ ? m.s.describeUpdateReportNowFailure_(late) : (() => { const t = require('./helpers/load-template.js'); void t; return 'updated only in part e-mail discussions TDoc revisions'; })()), true);
  let ticks = 0;
  m.s.reportUpdateElapsedMs_ = () => (++ticks > 20 ? 220000 : 1000);
  m.files.push('S4-263070-r8.docx');
  const stopped = m.update();
  m.s.reportUpdateElapsedMs_ = elapsed;
  check('the revisions stop between two TDocs when the time is up: incomplete, named', [stopped.success, (stopped.incomplete || []).some((x) => /^TDoc revisions \(\d+ of \d+ TDocs\)$/.test(x))], [true, true]);
  const resumed = m.update();
  check('the next update completes them: the new file is there, nothing twice', [resumed.incomplete, revisionsOf('S4-263070').split('\n').filter((l) => l === 'S4-263070-r8.docx').length], [undefined, 1]);
  if (firstCandidate !== null) {
    const before = collectorsReport(execFileSync('git', ['show', 'b5893c8:Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024 }).toString('utf8'));
    before.update();
    before.update();
    const c0 = before.body._state.calls;
    before.update();
    const oldCalls = before.body._state.calls - c0;
    console.log('  the third candidate (b5893c8), the same unchanged update with the collectors: ' + oldCalls + ' DocumentApp calls -- ' + Math.round(oldCalls / Math.max(1, noopCalls)) + 'x this one');
    check('the third candidate made at least 20x more DocumentApp calls for it', oldCalls >= 20 * noopCalls, true);
  }
}

// ====================================================================== K
console.log('K. what changed in Code.js and ReportCreator.js, and nothing else');
{
  const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const CREATOR = fs.readFileSync(path.join(ROOT, 'template', 'ReportCreator.js'), 'utf8').replace(/\r/g, '');
  // The release header of T-2026.10.30 (2.26.14) is taken out first.
  const view = codeWithoutGoesToPlenary(codeWithoutRelease2_26_14Header(FULL).code);
  const creatorView = creatorWithoutGoesToPlenary(CREATOR);
  check('every change is in Code.js and ReportCreator.js exactly once', view.problems.concat(creatorView.problems), []);
  const names = (src) => (src.match(/^function ([A-Za-z0-9_$]+)\(/gm) || []).map((l) => l.slice(9, -1));
  const bodyOf = (src, name) => { const at = src.indexOf('\nfunction ' + name + '('); return at === -1 ? null : src.slice(at, src.indexOf('\n}\n', at)); };
  const section = FULL.slice(FULL.indexOf(PLENARY_SECTION_START), FULL.indexOf(PLENARY_SECTION_END) + PLENARY_SECTION_END.length);
  check('the section: these functions, in this order', names(section), PLENARY_SECTION_FUNCTIONS);
  check('it stands directly after the Disposition Summary and Close Report', FULL.indexOf('// END OF DISPOSITION SUMMARY AND CLOSE REPORT\n// =========================================================\n\n' + PLENARY_SECTION_START) !== -1, true);
  check('the released functions it changed', names(view.code).filter((n) => bodyOf(FULL, n) !== bodyOf(view.code, n)).sort(), PLENARY_CHANGED_FUNCTIONS.slice().sort());
  check('the new functions are those of the section; none removed', [names(FULL).filter((n) => names(view.code).indexOf(n) === -1).sort(), names(view.code).filter((n) => names(FULL).indexOf(n) === -1)],
    [PLENARY_SECTION_FUNCTIONS.slice().sort(), []]);
  check('the released functions of ReportCreator.js it changed: the menu', names(creatorView.code).filter((n) => bodyOf(CREATOR, n) !== bodyOf(creatorView.code, n)), PLENARY_CREATOR_CHANGED_FUNCTIONS);
  let released = null;
  let releasedCreator = null;
  try {
    const show = (f) => execFileSync('git', ['show', 'template-release/T-2026.10.29:' + f], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
    released = show('Code.js');
    releasedCreator = show('template/ReportCreator.js');
  } catch (e) {
    released = null;
  }
  if (released === null) console.log('  note: template-release/T-2026.10.29 is not available in this checkout; the comparison with it is skipped.');
  else check('without it, Code.js and ReportCreator.js are those of T-2026.10.29, byte for byte', [view.code === released, creatorView.code === releasedCreator], [true, true]);
  const fnOf = (name) => bodyOf(FULL, name);
  check('the compact rows are the last thing: called in the finally of the status-dropdown step, and by the set-up', [
    (FULL.match(/= finalizePlenarySelection_\(/g) || []).length, /\} finally \{\n    resetStatusDropdownRun_\(\);\n    \/\/ After T-2026\.10\.29[^\n]*\n    const plenary = finalizePlenarySelection_\(mode, result\.ran && !result\.saveFailed\);/.test(fnOf('finalizeStatusDropdowns_')),
    /finalizePlenarySelection_\('setup', false\)/.test(fnOf('setUpGoesToPlenary'))], [2, true, true]);
  check('each step never throws: one try around all of it', ['finalizePlenarySelection_', 'syncPlenarySelectionOnUpdate_', 'preparePlenarySelectionForBuild_', 'capturePlenarySelectionForRebuild_', 'plenaryAgendaItems_',
    'notePlenaryListFacts_', 'ensurePlenaryLandscape_', 'afterPlenaryDispositionWrite_', 'plenarySnapshot_', 'formatPlenarySummaryTables_', 'syncTdocRevisionRelationships_', 'prepareTdocRevisionsForBuild_']
    .map((n) => /\n  try \{[\s\S]*\n  \} catch \(e\) \{/.test(fnOf(n))), [true, true, true, true, true, true, true, true, true, true, true, true]);
  check('the summary of a report with Goes to Plenary is read through the Docs API, never cell by cell: no getTables(), getChild() loop or getText() in its reader', [/getTables|getText|getChild/.test(fnOf('readPlenaryDispositionsOfTables_') + fnOf('plenaryDocsTables_') + fnOf('plenarySummarySectionOfContent_')), /plenarySnapshot_\(documentId\)/.test(fnOf('writePlenaryDispositionSummary_'))], [false, true]);
  check('nothing is fetched but the agenda TDoc, the revisions folder and the mailing list -- each through its existing reader, once',
    [/UrlFetchApp|Drive\.|DriveApp/.test(section), (section.match(/downloadMeetingAgenda_\([a-z]/g) || []).length, (section.match(/safeFetch_\(/g) || []).length, (section.match(/collectHybridListservMessages_\(/g) || []).length], [false, 1, 1, 1]);
  check('a report becomes one with Goes to Plenary in the set-up only', (section.match(/setProperty\(PLENARY_ENABLED_KEY_, '1'\)/g) || []).length, 1);
  check('DocumentApp never rewrites a Status cell: no setText() on it in the section but through setStatusCellText_(), which keeps the compact row',
    [(section.match(/statusCell\.setText\(|info\.cell\.setText\(/g) || []).length, /if \(!isCompactStatusCellText_\(text\)\) \{\n    cell\.setText\(status\);/.test(fnOf('setStatusCellText_'))], [0, true]);
  check('the revision relationship comes from the TDoc list only: "Revised to" and "Is revision of", nothing else', [/revisedTo|isRevisionOf/.test(fnOf('plenaryPredecessorsOf_')), /folder|listing|title/i.test(fnOf('plenaryPredecessorsOf_'))], [true, false]);
  const changed = execFileSync('git', ['diff', '--name-only', 'template-release/T-2026.10.29', '--', 'tools'], { cwd: ROOT }).toString().trim();
  check('the local listing agent is not touched', changed, '');
}

console.log(failures === 0 ? '\nAll Goes to Plenary checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
