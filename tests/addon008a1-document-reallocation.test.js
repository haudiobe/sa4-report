/**
 * ADDON-008A1 -- "Apply Document Reallocations" on a built report.
 *
 * Live failure (86172, first report built from the ADDON-008A agenda.csv
 * fallback): reallocating S4aI260081 from 3.7 (Portal allocation) to 2.7
 * threw "BODY_SECTION can't be cast to TABLE.". Root cause: the move did
 * `body.insertTable(i, body.removeChild(table).asTable())`, but
 * Body.removeChild() returns the BODY -- so every real move failed after
 * the table had already been removed. Unchanged since the initial commit and
 * previously untested; the fake body now mirrors that Apps Script behavior
 * (tests/helpers/fake-document.js), so these tests fail on the old code.
 *
 * Run: node tests/addon008a1-document-reallocation.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode } = require('./helpers/load-code.js');
const { makeFakeDocumentBody, tdoc } = require('./helpers/fake-document.js');

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

const FX = path.join(__dirname, 'fixtures');
const CSV_86172 = fs.readFileSync(path.join(FX, 'meeting-86172-agenda.csv'), 'utf8');
const LIST_86172 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-tdoc-list-values.json'), 'utf8')).values;
const FTP_MBS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/';
const PROPS_86172 = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) MBS SWG post 137-e', MEETING_ID: '86172', REPORT_SUFFIX: 'MBS',
  FTP_BASE: FTP_MBS, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172',
  AGENDA_CSV_URL: FTP_MBS.replace('Docs/', 'Agenda/agenda.csv') };

function wireDocument(sandbox, body) {
  const alerts = [];
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  sandbox.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(sandbox) });
  sandbox.DocumentApp.getUi = () => ({
    alert: (title, msg) => { alerts.push([title, msg]); return 'YES'; },
    Button: { YES: 'YES' }, ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }
  });
  sandbox.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  return alerts;
}

/** The real 86172 report, built through the ADDON-008A agenda.csv path. */
function build86172() {
  const { sandbox: s } = loadCode({ documentProperties: PROPS_86172 });
  const body = makeFakeDocumentBody(s);
  const alerts = wireDocument(s, body);
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  s.UrlFetchApp = { fetch: (u) => /agenda\.csv$/.test(u)
    ? { getResponseCode: () => 200, getContentText: () => CSV_86172 }
    : { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) } };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => LIST_86172 }) }] }) };
  const groups = {};
  const h = LIST_86172[0];
  LIST_86172.slice(1).forEach(r => { const k = r[h.indexOf('Agenda item')]; (groups[k] = groups[k] || { tdocs: [] }).tdocs.push(tdoc(r[0], k)); });
  s.downloadAndGroupTdocs_ = () => groups;
  s.buildSkeletonWithTdocTables();
  s.ensureReallocationTable_();
  return { s, body, alerts };
}

const isTdocTable = (c, id) => c.getType() === 'TABLE' && (() => { try { return c.getCell(0, 0).getText() === 'TDoc' && (!id || c.getCell(0, 1).getText() === id); } catch (e) { return false; } })();
const tablesFor = (body, id) => body._children.filter(c => isTdocTable(c, id));
function agendaRow(table) {
  for (let r = 0; r < table.getNumRows(); r++) if (table.getRow(r).getCell(0).getText() === 'Agenda Item') return table.getRow(r).getCell(1).getText();
  return null;
}
function sectionOf(body, element) {
  const i = body._children.indexOf(element);
  const h = body._children.slice(0, i).reverse().find(c => c.getType() === 'PARAGRAPH' && c.getHeading() !== 'NORMAL');
  return h ? h.getText() : null;
}
/** Every TDoc detail table: [id, section heading, agenda row]. */
function placements(body) {
  return body._children.filter(c => isTdocTable(c) && !/^(Title|Original Agenda)$/.test(c.getCell(0, 1).getText()))
    .map(t => [t.getCell(0, 1).getText(), sectionOf(body, t), agendaRow(t)]);
}
function snapshot(body) {
  return JSON.stringify(body._children.map(c => c.getType() === 'TABLE'
    ? ['T', Array.from({ length: c.getNumRows() }, (_, r) => Array.from({ length: c.getRow(r).getNumCells() }, (_, k) => c.getRow(r).getCell(k).getText()))]
    : ['P', c.getText(), c.getHeading()]));
}

// ============================================== the live 86172 structure ==

console.log('86172 (ADDON-008A CSV build): reallocate S4aI260081 3.7 -> 2.7 via the real menu path');
{
  const { s, body, alerts } = build86172();
  const before = placements(body);
  check('precondition: S4aI260081 is built under 3.7 (raw Portal allocation)',
    before.find(p => p[0] === 'S4aI260081'), ['S4aI260081', '3.7 Other Rel-20 matters including TEI', '3.7']);
  check('precondition: 2.7 already holds S4aI260089', before.filter(p => p[1].indexOf('2.7 ') === 0).map(p => p[0]), ['S4aI260089']);

  s.saveReallocation('S4aI260081', '3.7', '2.7', 'Wrongly allocated in the Portal');
  let error = null;
  try { s.applyDocumentReallocations(); } catch (e) { error = e.message; }
  check('no error (was: BODY_SECTION can\'t be cast to TABLE.)', error, null);

  const after = placements(body);
  check('S4aI260081 is now under 2.7 with Agenda Item 2.7',
    after.find(p => p[0] === 'S4aI260081'), ['S4aI260081', '2.7 Other Rel-20 matters including TEI', '2.7']);
  check('exactly one S4aI260081 table (no duplicate, nothing lost)', tablesFor(body, 'S4aI260081').length, 1);
  check('placed at the END of 2.7, after the existing S4aI260089 table',
    after.filter(p => p[1].indexOf('2.7 ') === 0).map(p => p[0]), ['S4aI260089', 'S4aI260081']);
  check('3.7 has no TDoc table left', after.filter(p => p[1].indexOf('3.7 ') === 0), []);
  check('every other TDoc untouched (same section, same agenda row)',
    after.filter(p => p[0] !== 'S4aI260081'), before.filter(p => p[0] !== 'S4aI260081'));
  const summary = alerts[alerts.length - 1][1];
  check('summary alert: 1 updated, 1 moved, 0 removed',
    [/Updated agenda items: 1\b/.test(summary), /Moved tables: 1\b/.test(summary), /Removed tables: 0\b/.test(summary)], [true, true, true]);

  // Section 3 consequence: apply never adds or removes agenda headings.
  check('section 3 headings remain (as empty agenda sections)',
    body._headingTexts().filter(t => /^3\./.test(t)).map(t => t.split(' ')[0]), ['3.0', '3.1', '3.2', '3.3', '3.4', '3.5', '3.6', '3.7', '3.8']);

  const moved = snapshot(body);
  s.applyDocumentReallocations();
  check('re-applying is a no-op on placement (idempotent)', snapshot(body), moved);
}

console.log('destination heading directly in the body with no TDoc table yet');
{
  const { s, body } = build86172();
  s.saveReallocation('S4aI260081', '3.7', '2.6', '');
  s.applyDocumentReallocations();
  const t = tablesFor(body, 'S4aI260081')[0];
  const i = body._children.indexOf(t);
  check('placed directly under the 2.6 heading', body._children[i - 1].getText().indexOf('2.6 ') === 0, true);
  check('before the 2.7 heading', body._children[i + 1].getText().indexOf('2.7 ') === 0, true);
  check('agenda row 2.6', agendaRow(t), '2.6');
}

console.log('invalid destination: nothing is changed, the error names TDoc / old / new / reason');
{
  const { s, body } = build86172();
  s.saveReallocation('S4aI260081', '3.7', '2.99', '');
  const before = snapshot(body);
  let error = null;
  try { s.applyDocumentReallocations(); } catch (e) { error = e.message; }
  check('refused with a useful message', /No reallocations were applied[\s\S]*S4aI260081 \(3\.7 → 2\.99\): the report has no agenda item 2\.99 heading\./.test(error), true);
  check('document unchanged (no heading appended, table and agenda row intact)', snapshot(body), before);
}

console.log('one invalid entry aborts the whole run before any change');
{
  const { s, body } = build86172();
  s.saveReallocation('S4aI260081', '3.7', '2.7', '');
  s.saveReallocation('S4aI260080', '2.5', 'soon', '');
  const before = snapshot(body);
  let error = null;
  try { s.applyDocumentReallocations(); } catch (e) { error = e.message; }
  check('refused', /"soon" is not an agenda item number/.test(error), true);
  check('the valid 3.7 -> 2.7 entry was NOT half-applied', snapshot(body), before);
}

console.log('missing source: skipped cleanly and reported');
{
  const { s, body, alerts } = build86172();
  s.saveReallocation('S4aI269999', '2.5', '2.7', '');
  const before = snapshot(body);
  let error = null;
  try { s.applyDocumentReallocations(); } catch (e) { error = e.message; }
  check('no error, document unchanged', [error, snapshot(body)], [null, before]);
  check('reported as not in the report', /Not in this report \(skipped\): S4aI269999/.test(alerts[alerts.length - 1][1]), true);
}

console.log('removal targets and relabel-in-place keep their existing meaning');
{
  const { s, body } = build86172();
  s.saveReallocation('S4aI260081', '3.7', 'withdrawn', '');
  s.applyDocumentReallocations();
  check('"withdrawn" removes the table', tablesFor(body, 'S4aI260081').length, 0);

  const b2 = build86172();
  b2.s.saveReallocation('S4aI260089', '2.7', '2.7', '');
  const idx = b2.body._children.indexOf(tablesFor(b2.body, 'S4aI260089')[0]);
  b2.s.applyDocumentReallocations();
  check('already in its section: relabelled only, not moved', b2.body._children.indexOf(tablesFor(b2.body, 'S4aI260089')[0]), idx);
}

console.log('two tables for the same TDoc: refused before any change');
{
  const { s, body } = build86172();
  const dup = tablesFor(body, 'S4aI260080')[0].copy();
  body.insertTable(body._children.length, dup);
  s.saveReallocation('S4aI260080', '2.5', '2.7', '');
  const before = snapshot(body);
  let error = null;
  try { s.applyDocumentReallocations(); } catch (e) { error = e.message; }
  check('refused, unchanged', [/contains 2 tables for this TDoc/.test(error), snapshot(body)], [true, before]);
}

// ========================================== legacy / main-style structure ==

console.log('main-style report (7.x headings, detail tables): forward and backward moves');
{
  const { sandbox: s } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Audio' } });
  const body = makeFakeDocumentBody(s);
  wireDocument(s, body);
  const heading = (t) => body.appendParagraph(t).setHeading('H3');
  const detail = (id, item) => { const t = body.appendTable(); [['TDoc', id], ['Title', 'T ' + id], ['Agenda Item', item]].forEach(r => { const row = t.appendTableRow(); r.forEach(c => row.appendTableCell(c)); }); };
  heading('7.1.3 Document Reallocations');
  body.insertTable(body._children.length, [['TDoc', 'Original Agenda', 'New Agenda', 'Reason'], ['S4-260001', '7.3', '7.5', ''], ['S4-260004', '7.5', '7.3', '']]);
  heading('7.3 Topic A'); detail('S4-260001', '7.3'); detail('S4-260002', '7.3');
  heading('7.4 Topic B'); detail('S4-260003', '7.4');
  heading('7.5 Topic C'); detail('S4-260004', '7.5');
  heading('7.6 Close of the session');
  s.applyDocumentReallocations();
  check('S4-260001 moved forward to the end of 7.5; S4-260004 moved back to the end of 7.3; others untouched', placements(body), [
    ['S4-260002', '7.3 Topic A', '7.3'], ['S4-260004', '7.3 Topic A', '7.3'],
    ['S4-260003', '7.4 Topic B', '7.4'],
    ['S4-260001', '7.5 Topic C', '7.5']
  ]);
}

// ======================================== shared lookup stays compatible ==

console.log('findInsertionPointForAgendaItem_ (continuous update) is unchanged; the new lookup never mutates');
{
  const { sandbox: s } = loadCode();
  const body = makeFakeDocumentBody(s);
  body.appendParagraph('2.5 A').setHeading('H3');
  body.appendParagraph('2.7 B').setHeading('H3');
  check('lookup: missing heading -> -1, nothing added', [s.findAgendaSectionEndIndex_(body, '2.9'), body._children.length], [-1, 2]);
  check('lookup: section end of 2.5 is the 2.7 heading', s.findAgendaSectionEndIndex_(body, '2.5'), 1);
  check('insertion point for an existing item is the same', s.findInsertionPointForAgendaItem_(body, '2.5', ''), 1);
  check('insertion point for a missing item still appends its heading (existing behavior)',
    [s.findInsertionPointForAgendaItem_(body, '2.9', 'New'), body._children[2].getText(), body._children[2].getHeading()], [3, '2.9 New', 'H2']);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
