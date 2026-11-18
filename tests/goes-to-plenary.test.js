/**
 * Goes to Plenary, the four disposition tables and the SA4 plenary agenda
 * item (after T-2026.10.29) -- docs/GOES_TO_PLENARY.md.
 *
 * The REAL build, update, set-up, Disposition Summary and Close Report run on
 * the fake document and the fake Docs API (tests/helpers/template-report.js),
 * as an MBS report of SA4#138 whose agenda is the real one (S4-261857,
 * tests/fixtures/agenda-S4-261857.google-items.json).
 *
 *   A.  the agenda mapping on the real SA4#138 agenda: every rule, precedence,
 *       missing and ambiguous mappings; nothing invented
 *   B.  a report without Goes to Plenary is exactly as before
 *   C.  Set Up Goes to Plenary…: the rows, the dropdowns, the retired Status
 *       "Plenary"; safe to run again
 *   D.  the update: the four tables, each TDoc once, the columns, the links
 *   E.  repeated updates: idempotent, no write when nothing changed
 *   F.  the selection is the rapporteur's: kept across status changes, updates,
 *       new TDocs and revision imports; never cleared
 *   G.  manual SA4 plenary agenda items: kept; work-item changes move only
 *       automatic values; "-" and an emptied cell
 *   H.  minutes and text typed into or below the summary; a three-table
 *       summary is rewritten in place; an unreadable selection writes nothing
 *   I.  Build Report from Scratch keeps the selections and the manual values
 *   J.  Close Report
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
const SUMMARY = 'Disposition Summary';
const TITLES = ['Agreed, Endorsed and Approved Documents – Goes to Plenary', 'Other Dispositions – Goes to Plenary',
  'Agreed, Endorsed and Approved Documents – Not Going to Plenary', 'Other Dispositions – Not Going to Plenary'];
const HEADER = ['TDoc', 'Title', 'Source', 'Status', 'SWG Agenda Item', 'SA4 Plenary Agenda Item'];
const AGENDA_138 = JSON.parse(fs.readFileSync(path.join(ROOT, 'tests', 'fixtures', 'agenda-S4-261857.google-items.json'), 'utf8')).items
  .map((i) => ({ number: i.number, title: i.title, level: i.level, heading: 'NORMAL', text: '' }));
const MAIN = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_NAME: 'SA4#138', MEETING_ID: '60780', REPORT_SUFFIX: 'MBS', MEETING_DATE: 'November 16, 2026',
  FTP_BASE: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/', TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23138.xlsx',
  AGENDA_TDOC: 'S4-261857' };
const COLUMNS = ['Release', 'Related WIs'];
// An MBS report of SA4#138 (agenda items 8.x). All numbers and names are synthetic; the agenda is the real one.
const TDOCS = () => [
  { id: 'S4-262001', agenda: '8.5', status: 'agreed', title: '[AMD_ARCH_Ph2-MED] WT#8 multi-access', uploaded: true },
  { id: 'S4-262002', agenda: '8.6', status: 'noted', title: '[Energy_PRO-MED] TS skeleton', uploaded: true },
  { id: 'S4-262003', agenda: '8.4', status: 'agreed', type: 'CR', Release: 'Rel-19', title: 'CR 26.511 correction', uploaded: true },
  { id: 'S4-262004', agenda: '8.7', status: 'available', title: 'QUIC evaluation results', uploaded: true },
  { id: 'S4-262005', agenda: '8.8', status: 'endorsed', 'Related WIs': 'TEI20', title: 'Rel-20 TEI clarification', uploaded: true },
  { id: 'S4-262006', agenda: '8.8', status: 'agreed', 'Related WIs': 'TEI19', title: 'Rel-19 TEI fix', uploaded: true },
  { id: 'S4-262007', agenda: '8.8', status: 'agreed', 'Related WIs': 'AMD_ARCH_Ph2-MED, Energy_PRO-MED', title: 'Cross-WI alignment', uploaded: true },
  { id: 'S4-262008', agenda: '8.3', status: 'noted', title: 'LS from another group', uploaded: true },
  { id: 'S4-262009', agenda: '8.5', status: 'available', title: '[AMD_ARCH_Ph2-MED] open issue', uploaded: true },
  { id: 'S4-262010', agenda: '8.5', status: 'approved', type: 'CR', Release: 'Rel-20', title: 'CR 26.501 new feature', uploaded: true }
];
const EXPECTED_AUTO = { 'S4-262001': '15.4', 'S4-262002': '15.5', 'S4-262003': '14', 'S4-262004': '16.3', 'S4-262005': '15.7', 'S4-262006': '14', 'S4-262007': '', 'S4-262008': '',
  'S4-262009': '15.4', 'S4-262010': '15.4' };

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
  r.close = () => r.exec(() => r.s.closeReportCore_());
  /** The labels of the rows of a TDoc table, in order. */
  r.labels = (id) => { const t = r.table(id); return t ? rowsOf(t).map((x) => x[0]) : null; };
  /** The summary: [[class line, rows of its table or 'None.']] in order, or null. */
  r.summary = () => {
    const kids = r.body._children;
    const at = kids.findIndex((c) => !isTable(c) && c.getText() === SUMMARY && c.getHeading() !== 'NORMAL');
    if (at === -1) return null;
    const out = [];
    for (let i = at + 1; i < kids.length; i++) {
      const c = kids[i];
      if (isTable(c)) { if (out.length) out[out.length - 1].push(rowsOf(c)); continue; }
      const text = c.getText();
      if (TITLES.some((t) => text.indexOf(t + ' (') === 0) || /^(Agreed, Endorsed and Approved Documents|Other Disposed Documents|Documents Not Yet Disposed) \(/.test(text)) out.push([text]);
      else if (text === 'None.' && out.length) out[out.length - 1].push('None.');
    }
    return out;
  };
  r.headings = () => r.body._children.filter((c) => !isTable(c) && c.getText() === SUMMARY && c.getHeading() !== 'NORMAL').length;
  r.snapshot = () => JSON.stringify(r.body._children.map((c) => (isTable(c)
    ? ['T', Array.from({ length: c.getNumRows() }, (x, i) => Array.from({ length: c.getRow(i).getNumCells() }, (y, k) => { const cell = c.getRow(i).getCell(k); return [cell.getText(), cell._dropdown ? cell._dropdown.optionId : null, cell.editAsText().getLinkUrl(0) || '']; }))]
    : ['P', c.getHeading(), c.getText()])));
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
  // Every SWG's work items map to their 15.x / 16.x item through the acronym of the SWG agenda item.
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
  check('the choice of a "Goes to Plenary" cell', ['☑ Yes', '☐ No', '', 'Yes', 'x', 'no', 'later', '✓'].map(S.plenaryChoiceOfText_), ['yes', 'no', 'no', 'yes', 'yes', 'no', 'unknown', 'yes']);
  const cls = (disposition, plenary, conflict) => S.plenaryDispositionClassOf_({ disposition: disposition, conflict: !!conflict }, plenary).key;
  check('the four classes: positive = agreed/endorsed/approved (and partially agreed); everything else is "other", also a conflict or no status',
    [cls('AGREED_ENDORSED_APPROVED', true), cls('DISPOSED_OTHER', true), cls('AGREED_ENDORSED_APPROVED', false), cls('UNRESOLVED', false), cls('AGREED_ENDORSED_APPROVED', true, true)],
    ['POSITIVE_PLENARY', 'OTHER_PLENARY', 'POSITIVE_NOT_PLENARY', 'OTHER_NOT_PLENARY', 'OTHER_PLENARY']);
  check('the taxonomy is the one of the Disposition Summary, not a new one', ['agreed', 'partially agreed', 'endorsed', 'approved', 'noted', 'revised', 'merged', 'withdrawn', 'replied', 'replied to', 'available', 'reserved', 'parked', 'postponed', 'plenary', 'other', 'not pursued', '']
    .map((s) => S.classifyTdocDisposition_(s).disposition === 'AGREED_ENDORSED_APPROVED'), [true, true, true, true, false, false, false, false, false, false, false, false, false, false, false, false, false, false]);
}

// ====================================================================== B
console.log('B. a report without Goes to Plenary is exactly as before');
const r = mbs();
{
  const built = r.build();
  check('the MBS report of SA4#138 builds, with its ten TDocs under 8.x and Status dropdowns', [built.ok, r.order(), r.shown('S4-262001')],
    [true, ['S4-262008', 'S4-262003', 'S4-262001', 'S4-262009', 'S4-262010', 'S4-262002', 'S4-262004', 'S4-262005', 'S4-262006', 'S4-262007'], 'dropdown:agreed']);
  check('no Goes to Plenary rows, no definition, not marked', [r.labels('S4-262001').indexOf('Goes to Plenary'), r.api.definitionTitled('Goes to Plenary'), r.docProps.getProperty('GOES_TO_PLENARY')], [-1, null, null]);
  const u = r.update();
  check('an update gives it no summary and no rows, and reads no agenda for it', [u.success, u.goesToPlenary, u.dispositionSummary, r.summary(), r.labels('S4-262004').indexOf('Goes to Plenary'), r.docProps.getProperty('PLENARY_AGENDA_ITEMS')],
    [true, undefined, undefined, null, -1, null]);
  // The three-table summary of T-2026.10.29, written by the menu: it is what H rewrites later.
  r.summaryMenu();
  check('Update Disposition Summary writes the three tables there, as before', r.summary().map((c) => c[0]), ['Agreed, Endorsed and Approved Documents (6)', 'Other Disposed Documents (2)', 'Documents Not Yet Disposed (2)']);
}

console.log('B2. an ad-hoc report: the rows and the four tables, no SA4 plenary agenda item');
{
  const a = templateReport({ tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'agreed', uploaded: true }, { id: 'S4aA269002', agenda: '4.4', status: 'noted', uploaded: true }] });
  a.build();
  let agendaReads = 0;
  a.s.downloadMeetingAgenda_ = () => { agendaReads++; return []; };
  a.exec(() => a.s.setUpGoesToPlenary());
  const u = a.update();
  check('set up and updated: rows, dropdowns, the four tables', [u.goesToPlenary, u.dispositionSummary, a.api.shownIn('S4aA269001', 'Goes to Plenary'), a.body._children.filter((c) => isTable(c) && c.getNumRows() > 0 && c.getRow(0).getNumCells() === 6).length],
    ['done', 'created', 'dropdown:' + NO, 2]);
  check('the plenary agenda of a main meeting does not exist here: blank, and no agenda is read for it', [a.field('S4aA269001', 'Plenary Agenda Item'), agendaReads], ['', 0]);
}

// ====================================================================== C
console.log('C. Set Up Goes to Plenary…');
{
  // Before the set-up: S4-262009 was set to the manual Status "Plenary", and somebody typed minutes into S4-262001.
  r.user(() => r.api.pick('S4-262009', 'Plenary'));
  r.user(() => r.api.cellOf('S4-262001', 'Minutes').setText('Minutes typed by hand.'));
  const writesBefore = r.writes();
  const done = r.setUp();
  check('it asks first, with what it will do', [/Goes to Plenary/.test(r.ui.alerts[r.ui.alerts.length - 2][1]), /Status "Plenary" \(1\): S4-262009/.test(r.ui.alerts[r.ui.alerts.length - 2][1])], [true, true]);
  check('every TDoc table gets the two rows right below its Status', r.tdocs.map((t) => r.labels(t.id).slice(-3)), r.tdocs.map(() => ['Status', 'Goes to Plenary', 'Plenary Agenda Item']));
  check('the rows above it are untouched; the minutes typed by hand are kept', [r.labels('S4-262001').slice(0, -2), r.field('S4-262001', 'Minutes')],
    [['TDoc', 'Title', 'Source', 'Contact', 'Agenda Item', 'Type/For', 'E-mail Discussion', 'Revisions', 'Minutes', 'Disposition', 'Status'], 'Minutes typed by hand.']);
  check('Goes to Plenary is a dropdown: "☐ No" for every TDoc, "☑ Yes" for the one whose Status was "Plenary"',
    r.tdocs.map((t) => r.api.shownIn(t.id, 'Goes to Plenary')), r.tdocs.map((t) => (t.id === 'S4-262009' ? 'dropdown:' + YES : 'dropdown:' + NO)));
  check('its definition: "Goes to Plenary", two options', r.api.definitionTitled('Goes to Plenary').options, [NO, YES]);
  check('the Status "Plenary" became "available"; "Plenary" is no option of the Status dropdown any more; every other status is as it was',
    [r.shown('S4-262009'), r.api.definitionTitled('Document Status').options.indexOf('Plenary'), r.tdocs.filter((t) => t.id !== 'S4-262009').map((t) => r.shown(t.id))],
    ['dropdown:available', -1, r.tdocs.filter((t) => t.id !== 'S4-262009').map((t) => 'dropdown:' + t.status)]);
  check('the Plenary Agenda Item cells are empty until the next update', r.tdocs.map((t) => r.field(t.id, 'Plenary Agenda Item')), r.tdocs.map(() => ''));
  check('the result', [done.ok, done.rows, done.plenaryStatuses, done.dropdowns.inserted, done.dropdowns.optionRemoved, r.docProps.getProperty('GOES_TO_PLENARY')], [true, 20, 1, 10, true, '1']);
  check('its writes: one batch by id (the Status, "Plenary" out of its dropdown), the definition, one batch of inserts tied to its revision',
    r.api.calls.filter((c) => c.call === 'batchUpdate').slice(writesBefore).map((c) => [c.requests.map((q) => Object.keys(q)[0]).filter((k, i, a) => a.indexOf(k) === i).join('+'), !!c.requiredRevisionId]),
    [['updateDropdownProperties+updateDropdownDefinitionProperties', false], ['createDropdownDefinition', false], ['deleteContentRange+insertDropdown', true]]);
  const again = r.setUp();
  check('run again: nothing to do, nothing written', [again.nothingToDo, r.writes() - writesBefore], [true, 3]);
}

// ====================================================================== D
console.log('D. the update: the four tables');
let afterFirst;
{
  r.mark();
  const u = r.update();
  check('the update succeeds and brings Goes to Plenary up to date', [u.success, u.goesToPlenary, u.dispositionSummary], [true, 'done', 'replaced']);
  check('the SA4 plenary agenda items are filled in from the work items and the agenda (S4-261857)', r.tdocs.reduce((o, t) => (o[t.id] = r.field(t.id, 'Plenary Agenda Item'), o), {}), EXPECTED_AUTO);
  check('the agenda was read once and is kept for its agenda TDoc', [JSON.parse(r.docProps.getProperty('PLENARY_AGENDA_ITEMS')).agendaTdoc, r.docProps.getProperty('PLENARY_AGENDA_AUTO:S4-262001')], ['S4-261857', '15.4']);
  const s = r.summary();
  check('one summary, four tables, in the order asked for, with counts', [r.headings(), s.map((c) => c[0])], [1, [TITLES[0] + ' (0)', TITLES[1] + ' (1)', TITLES[2] + ' (6)', TITLES[3] + ' (3)']]);
  check('an empty table says "None."', s[0][1], 'None.');
  check('the columns', s[1][1][0], HEADER);
  check('Other – Goes to Plenary: the TDoc whose Status was "Plenary"', s[1][1].slice(1), [['S4-262009', '[AMD_ARCH_Ph2-MED] open issue', 'ExampleCorp', 'available', '8.5', '15.4']]);
  check('Agreed/endorsed/approved – Not Going: in the order of the report, with status, SWG and plenary agenda item', s[2][1].slice(1).map((x) => [x[0], x[3], x[4], x[5]]), [
    ['S4-262003', 'agreed', '8.4', '14'], ['S4-262001', 'agreed', '8.5', '15.4'], ['S4-262010', 'approved', '8.5', '15.4'], ['S4-262005', 'endorsed', '8.8', '15.7'],
    ['S4-262006', 'agreed', '8.8', '14'], ['S4-262007', 'agreed', '8.8', '']]);
  check('Other – Not Going', s[3][1].slice(1).map((x) => [x[0], x[3], x[4], x[5]]), [['S4-262008', 'noted', '8.3', ''], ['S4-262002', 'noted', '8.6', '15.5'], ['S4-262004', 'available', '8.7', '16.3']]);
  const all = [].concat(...s.map((c) => (Array.isArray(c[1]) ? c[1].slice(1).map((x) => x[0]) : [])));
  check('every TDoc of the report is in exactly one table', [all.length, all.slice().sort()], [10, r.tdocs.map((t) => t.id).sort()]);
  const links = r.body._children.filter((c) => isTable(c) && c.getRow(0).getNumCells() === 6).map((t) => Array.from({ length: t.getNumRows() - 1 }, (_, i) => t.getRow(i + 1).getCell(0).editAsText().getLinkUrl(0)));
  check('the TDoc number keeps the link of its TDoc table', [].concat(...links).every((l, i) => /\.zip$/.test(l)), true);
  check('the summary stands at the end of the report, as before', r.body._children.slice(-1)[0].getType() === 'TABLE' || r.body._children.slice(-2)[0].getType() === 'TABLE', true);
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
  check('no table and no row is duplicated', [r.headings(), r.tdocs.map((t) => r.labels(t.id).filter((l) => l === 'Goes to Plenary' || l === 'Plenary Agenda Item').length)], [1, r.tdocs.map(() => 2)]);
}

// ====================================================================== F
console.log('F. the selection is the rapporteur\'s');
{
  r.user(() => r.api.pickIn('S4-262001', 'Goes to Plenary', YES));
  r.update();
  check('"☑ Yes" on an agreed TDoc moves it to Agreed – Goes to Plenary', r.summary()[0][1].slice(1).map((x) => x[0]), ['S4-262001']);
  r.user(() => r.api.pick('S4-262001', 'noted'));
  r.update();
  check('its disposition changes to noted: it moves to Other – Goes to Plenary; the selection stays "☑ Yes"', [r.summary()[1][1].slice(1).map((x) => x[0]), r.api.shownIn('S4-262001', 'Goes to Plenary')],
    [['S4-262001', 'S4-262009'], 'dropdown:' + YES]);
  r.user(() => r.api.pick('S4-262001', 'agreed'));
  r.update();
  check('back to agreed: back in Agreed – Goes to Plenary, still "☑ Yes"', [r.summary()[0][1].slice(1).map((x) => x[0]), r.api.shownIn('S4-262001', 'Goes to Plenary')], [['S4-262001'], 'dropdown:' + YES]);
  // A new TDoc, and a status the Portal moves on (available -> agreed): neither touches a selection.
  r.tdocs.push({ id: 'S4-262011', agenda: '8.6', status: 'available', title: '[Energy_PRO-MED] new contribution', uploaded: true });
  r.tdoc('S4-262009').status = 'agreed';
  const u = r.update();
  check('a new TDoc gets the rows, "☐ No" as a dropdown and its plenary agenda item, in the same update', [u.success, r.labels('S4-262011').slice(-3), r.api.shownIn('S4-262011', 'Goes to Plenary'), r.field('S4-262011', 'Plenary Agenda Item')],
    [true, ['Status', 'Goes to Plenary', 'Plenary Agenda Item'], 'dropdown:' + NO, '15.5']);
  check('the Portal moved S4-262009 on to agreed; its "☑ Yes" stays: Agreed – Goes to Plenary', [r.shown('S4-262009'), r.api.shownIn('S4-262009', 'Goes to Plenary'), r.summary()[0][1].slice(1).map((x) => x[0])],
    ['dropdown:agreed', 'dropdown:' + YES, ['S4-262001', 'S4-262009']]);
  check('the new TDoc is in the summary once', [].concat(...r.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1).map((x) => x[0]) : []))).filter((n) => n === 'S4-262011').length, 1);
  const selections = r.tdocs.map((t) => r.api.shownIn(t.id, 'Goes to Plenary'));
  r.exec(() => { r.s.resetStatusDropdownRun_(); r.s.collectRevisionsCore_(); r.s.finalizeStatusDropdowns_('update'); });
  check('Update TDoc Revisions leaves every selection as it was', r.tdocs.map((t) => r.api.shownIn(t.id, 'Goes to Plenary')), selections);
}

// ====================================================================== G
console.log('G. manual SA4 plenary agenda items');
{
  r.user(() => r.api.cellOf('S4-262002', 'Plenary Agenda Item').setText('16.2'));     // a manual override of 15.5
  r.user(() => r.api.cellOf('S4-262007', 'Plenary Agenda Item').setText('15.4'));     // blank (ambiguous) completed by hand
  r.user(() => r.api.cellOf('S4-262003', 'Plenary Agenda Item').setText('-'));        // kept blank on purpose
  r.update();
  check('a manual value is kept by the update and shown in the summary; "-" is shown blank', [r.field('S4-262002', 'Plenary Agenda Item'), r.field('S4-262007', 'Plenary Agenda Item'), r.field('S4-262003', 'Plenary Agenda Item'),
    [].concat(...r.summary().map((c) => (Array.isArray(c[1]) ? c[1].slice(1) : []))).filter((x) => ['S4-262002', 'S4-262007', 'S4-262003'].indexOf(x[0]) !== -1).map((x) => [x[0], x[5]])],
    ['16.2', '15.4', '-', [['S4-262003', ''], ['S4-262007', '15.4'], ['S4-262002', '16.2']]]);
  // The work items change in the TDoc list: an automatic value follows, a manual one does not.
  r.tdoc('S4-262005')['Related WIs'] = 'AMD_ARCH_Ph2-MED';                  // automatic: 15.7 -> 15.4
  r.tdoc('S4-262002')['Related WIs'] = 'FS_QStream_MED';                    // manual 16.2 stays
  r.tdoc('S4-262003').Release = 'Rel-20';                                   // "-" stays (manual)
  r.tdoc('S4-262006')['Related WIs'] = '';                                  // TEI19 gone: 14 -> 15.7 (the item of "Other Rel-20 matters including TEI")
  r.update();
  check('a work-item change moves the automatic values only', ['S4-262005', 'S4-262002', 'S4-262003', 'S4-262006'].map((n) => r.field(n, 'Plenary Agenda Item')), ['15.4', '16.2', '-', '15.7']);
  check('what the update wrote is what it remembers; a manual value is not remembered as automatic', ['S4-262005', 'S4-262006', 'S4-262002'].map((n) => r.docProps.getProperty('PLENARY_AGENDA_AUTO:' + n)), ['15.4', '15.7', '15.5']);
  r.user(() => r.api.cellOf('S4-262002', 'Plenary Agenda Item').setText(''));
  r.update();
  check('a manual value removed by hand: the cell is automatic again (now ambiguous: 15.5 and 16.3 -> blank)', [r.field('S4-262002', 'Plenary Agenda Item'), r.docProps.getProperty('PLENARY_AGENDA_AUTO:S4-262002')], ['', null]);
  r.user(() => r.api.cellOf('S4-262002', 'Plenary Agenda Item').setText('16.3'));
  r.update();                                                               // the summary takes the new value
  const before = r.snapshot();
  r.update();
  r.update();
  check('repeated updates leave the manual values as they are', [r.snapshot() === before, ['S4-262002', 'S4-262007', 'S4-262003'].map((n) => r.field(n, 'Plenary Agenda Item'))], [true, ['16.3', '15.4', '-']]);
  // An agenda that cannot be had changes nothing that is automatic.
  const kept = r.docProps.getProperty('PLENARY_AGENDA_ITEMS');
  r.docProps.deleteProperty('PLENARY_AGENDA_ITEMS');
  r.s.downloadMeetingAgenda_ = () => { throw new Error('Could not download agenda ZIP (403)'); };
  const values = r.tdocs.map((t) => r.field(t.id, 'Plenary Agenda Item'));
  const u = r.update();
  check('the agenda cannot be read: the update succeeds and no plenary agenda item changes', [u.success, r.tdocs.map((t) => r.field(t.id, 'Plenary Agenda Item'))], [true, values]);
  r.s.downloadMeetingAgenda_ = () => AGENDA_138;
  r.docProps.setProperty('PLENARY_AGENDA_ITEMS', kept);
}

// ====================================================================== H
console.log('H. minutes, text typed into or below the summary, an unreadable selection');
{
  const kids = r.body._children;
  // Minutes typed below the summary are not the summary's.
  r.user(() => r.body.appendParagraph('Minutes typed below the summary by hand.'));
  r.user(() => r.api.pick('S4-262004', 'agreed'));
  const u = r.update();
  check('text typed below it is kept; the summary above it is updated', [u.dispositionSummary, kids[kids.length - 1].getText(), r.summary()[2][0]], ['replaced', 'Minutes typed below the summary by hand.', TITLES[2] + ' (6)']);
  // Text typed INTO it: nothing is written, nothing removed.
  const at = kids.findIndex((c) => !isTable(c) && c.getText().indexOf(TITLES[2]) === 0);
  r.user(() => r.body.insertParagraph(at, 'A remark typed into the summary.'));
  r.user(() => r.api.pick('S4-262004', 'noted'));
  const before = r.snapshot();
  const v = r.update();
  check('text typed into it: refused, nothing changed, the remark kept', [v.dispositionSummary, r.snapshot() === before], ['refused', true]);
  r.user(() => r.body.removeChild(kids[at]));
  check('removed again: the next update writes it', r.update().dispositionSummary, 'replaced');
  // A selection that cannot be read: nothing is written.
  r.user(() => r.api.pickIn('S4-262008', 'Goes to Plenary', YES));
  const snap = r.snapshot();
  r.api.fail.get = (id) => (id === 'REPORTdoc000000000000000000000000000000000' ? 'Internal error' : null);
  const w = r.update();
  r.api.fail.get = null;
  check('the API cannot read the dropdowns: the summary is not written (no guess), the document is as it was', [w.dispositionSummary, r.snapshot() === snap], ['unreadable', true]);
  check('the next update that can read them moves S4-262008 to Other – Goes to Plenary', [r.update().dispositionSummary, r.summary()[1][1].slice(1).map((x) => x[0])], ['replaced', ['S4-262008']]);
}

// ====================================================================== I
console.log('I. Build Report from Scratch keeps the selections and the manual values');
{
  const selections = r.tdocs.reduce((o, t) => (o[t.id] = r.api.shownIn(t.id, 'Goes to Plenary'), o), {});
  const agenda = r.tdocs.reduce((o, t) => (o[t.id] = r.field(t.id, 'Plenary Agenda Item'), o), {});
  // A Status "Plenary" picked by somebody after the set-up (the option is gone; a text typed by hand).
  r.user(() => { const c = r.api.statusCell('S4-262010'); c._dropdown = null; c._t = 'Plenary'; });
  const built = r.build();
  check('the rebuild succeeds; every TDoc table has the rows again', [built.ok, r.tdocs.map((t) => r.labels(t.id).slice(-2))], [true, r.tdocs.map(() => ['Goes to Plenary', 'Plenary Agenda Item'])]);
  const expected = Object.assign({}, selections, { 'S4-262010': 'dropdown:' + YES });
  check('every "☑ Yes" is kept; the typed Status "Plenary" became "☑ Yes"', r.tdocs.reduce((o, t) => (o[t.id] = r.api.shownIn(t.id, 'Goes to Plenary'), o), {}), expected);
  check('the manual plenary agenda items are kept, the automatic ones computed again', r.tdocs.reduce((o, t) => (o[t.id] = r.field(t.id, 'Plenary Agenda Item'), o), {}), agenda);
  check('"Plenary" does not come back as a Status', [r.shown('S4-262010'), r.tdocs.some((t) => /plenary/i.test(r.shown(t.id)))], ['dropdown:approved', false]);
  const u = r.update();
  check('the update after it writes the four tables again', [u.dispositionSummary, r.summary().map((c) => c[0].replace(/ \(\d+\)$/, ''))], ['created', TITLES]);
}

// ====================================================================== J
console.log('J. Close Report');
{
  const c = r.close();
  check('Close Report writes the same four tables; its counts are the classes of the Disposition Summary', [c.summaryState, r.headings(), Object.keys(c.counts).slice(0, 3)],
    ['unchanged', 1, ['AGREED_ENDORSED_APPROVED', 'DISPOSED_OTHER', 'UNRESOLVED']]);
  check('the TDocs not yet disposed are listed as before', c.unresolved.map((e) => e.tdoc), ['S4-262011']);
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
  check('the Goes to Plenary dropdowns are the last thing: called in the finally of the status-dropdown step, and by the set-up', [
    (FULL.match(/= finalizePlenarySelection_\(/g) || []).length, /\} finally \{\n    resetStatusDropdownRun_\(\);\n    \/\/ After T-2026\.10\.29[^\n]*\n    const plenary = finalizePlenarySelection_\(mode, result\.ran && !result\.saveFailed\);/.test(fnOf('finalizeStatusDropdowns_')),
    /finalizePlenarySelection_\('setup', false\)/.test(fnOf('setUpGoesToPlenary'))], [2, true, true]);
  check('each step never throws: one try around all of it', ['finalizePlenarySelection_', 'syncPlenarySelectionOnUpdate_', 'preparePlenarySelectionForBuild_', 'capturePlenarySelectionForRebuild_', 'plenaryAgendaItems_', 'notePlenaryListFacts_']
    .map((n) => /\n  try \{[\s\S]*\n  \} catch \(e\) \{/.test(fnOf(n))), [true, true, true, true, true, true]);
  check('nothing is fetched but the agenda TDoc of the meeting, through the existing reader', [/UrlFetchApp|safeFetch_|Drive\.|DriveApp/.test(section), (section.match(/downloadMeetingAgenda_\(agendaTdoc, cfg\.FTP_BASE\)/g) || []).length, (section.match(/downloadMeetingAgenda_\([a-z]/g) || []).length], [false, 1, 1]);
  check('a report becomes one with Goes to Plenary in the set-up only', (section.match(/setProperty\(PLENARY_ENABLED_KEY_, '1'\)/g) || []).length, 1);
  check('nothing in the section clears a selection: the only value it writes into Goes to Plenary is "☑ Yes" (or "☐ No" in a new row)',
    [(section.match(/setText\(PLENARY_YES_\)/g) || []).length, /setText\(PLENARY_NO_\)/.test(section)], [1, false]);
  const changed = execFileSync('git', ['diff', '--name-only', 'template-release/T-2026.10.29', '--', 'tools'], { cwd: ROOT }).toString().trim();
  check('the local listing agent is not touched', changed, '');
}

console.log(failures === 0 ? '\nAll Goes to Plenary checks passed.' : `\n${failures} check(s) FAILED.`);
process.exit(failures === 0 ? 0 : 1);
