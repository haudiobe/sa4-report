/**
 * ADDON-008A1b -- saved Document Reallocations survive a full rebuild and
 * drive the rebuilt report.
 *
 * Old path: buildSkeletonWithTdocTables() cleared the document, THEN
 * downloadAndGroupTdocs_() read the reallocation table (already gone), so a
 * rebuild used raw Portal allocations, the agenda.csv sections came from the
 * raw allocations too, and the reallocation table itself was never
 * recreated (the 6G branch rebuilt it from the cleared, empty body).
 *
 * Now the build reads the reallocations first, uses the EFFECTIVE allocation
 * (reallocation, else source) for agenda projection and TDoc placement,
 * refuses unusable destinations before clearing, and restores the table.
 *
 * Real case: 86172, S4aI260081 is 3.7 in the Portal, reallocated to 2.7.
 * Everything below runs the REAL build and the REAL downloadAndGroupTdocs_();
 * only network / Drive / Sheets and the template are faked.
 *
 * Run: node tests/addon008a1b-reallocation-rebuild.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');

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
const LIST_85916 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-85916-tdoc-list-values.json'), 'utf8')).values;
const FTP_MBS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/';
const PROPS_86172 = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) MBS SWG post 137-e', MEETING_ID: '86172', REPORT_SUFFIX: 'MBS',
  FTP_BASE: FTP_MBS, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172',
  AGENDA_CSV_URL: FTP_MBS.replace('Docs/', 'Agenda/agenda.csv') };
const FTP_AUDIO = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const PROPS_85916 = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio',
  FTP_BASE: FTP_AUDIO, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916', AGENDA_TDOC: 'S4aA260090' };
const AUDIO_AGENDA = [['1', 'Opening of the meeting'], ['2', 'Approval of the agenda and registration of documents'], ['3', 'IPR'],
  ['4', 'ULBC-MED'], ['4.2', 'Project plan/timeline'], ['4.3', 'Performance requirements'], ['4.4', 'Design constraints'],
  ['4.5', 'Qualification/Selection rules'], ['4.6', 'Deliverables'], ['4.7', 'Processing plan'], ['4.8', 'Test plan'],
  ['4.9', 'RTP header compression & system integration'], ['4.10', 'Other'], ['5', 'Close of the meeting']]
  .map(([number, title]) => ({ number, title, level: number.split('.').length, heading: 'NORMAL', text: '' }));

/** One report document; build() can be run repeatedly on it (a rebuild). */
function makeReport(props, options) {
  const opts = options || {};
  const loaded = loadCode({ documentProperties: props });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const alerts = [];
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ({ alert: (t, m) => { alerts.push([t, m]); return 'YES'; }, Button: { YES: 'YES' }, ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' } });
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  const values = () => opts.values || LIST_86172;
  s.UrlFetchApp = { fetch: (u) => /agenda\.csv$/.test(u)
    ? { getResponseCode: () => 200, getContentText: () => CSV_86172 }
    : { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) } };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => values(), getRichTextValues: () => values().map(r => r.map(() => null)) }) }] }) };
  if (opts.tdocAgenda) s.downloadMeetingAgenda_ = () => opts.tdocAgenda;
  function build() {
    try { s.buildSkeletonWithTdocTables(); return null; } catch (e) { return e.message; }
  }
  return { s, body, alerts, docProps: loaded.docProps, build };
}

const isTable = (c) => c.getType() === 'TABLE';
const cell = (t, r, k) => { try { return t.getRow(r).getCell(k).getText(); } catch (e) { return null; } };
const detailTables = (body) => body._children.filter(c => isTable(c) && cell(c, 0, 0) === 'TDoc' && /^S4/.test(cell(c, 0, 1) || ''));
function agendaRow(t) { for (let r = 0; r < t.getNumRows(); r++) if (cell(t, r, 0) === 'Agenda Item') return cell(t, r, 1); return null; }
function sectionOf(body, el) {
  const i = body._children.indexOf(el);
  const h = body._children.slice(0, i).reverse().find(c => c.getType() === 'PARAGRAPH' && c.getHeading() !== 'NORMAL');
  return h ? h.getText().split(' ')[0] : null;
}
const placements = (body) => detailTables(body).map(t => [cell(t, 0, 1), sectionOf(body, t), agendaRow(t)]);
const where = (body, id) => placements(body).filter(p => p[0] === id);
const topSections = (body) => Array.from(new Set(body._headingTexts().map(h => h.split(' ')[0]).filter(n => /^\d/.test(n)).map(n => n.split('.')[0])));
function reallocationRows(body) {
  const t = body._children.find(c => isTable(c) && cell(c, 0, 1) === 'Original Agenda');
  return t ? Array.from({ length: t.getNumRows() - 1 }, (_, i) => [0, 1, 2, 3].map(k => cell(t, i + 1, k))) : null;
}
function snapshot(body) {
  return JSON.stringify(body._children.map(c => isTable(c)
    ? ['T', Array.from({ length: c.getNumRows() }, (_, r) => Array.from({ length: c.getRow(r).getNumCells() }, (_, k) => cell(c, r, k)))]
    : ['P', c.getText(), c.getHeading()]));
}
/** Build once, record reallocations through the real dialog save path, rebuild. */
function rebuildWith(entries, props, options) {
  const rep = makeReport(props || PROPS_86172, options);
  const first = rep.build();
  const initial = placements(rep.body);
  entries.forEach(([tdoc, from, to]) => rep.s.saveReallocation(tdoc, from, to, 'test'));
  const beforeRebuild = snapshot(rep.body);
  const error = rep.build();
  return Object.assign(rep, { first, initial, error, beforeRebuild });
}

// =============================================== the exact 86172 rebuild ==

console.log('86172: raw 3.7, saved reallocation 3.7 -> 2.7, full rebuild');
{
  const r = rebuildWith([['S4aI260081', '3.7', '2.7']]);
  check('first (raw) build ok, S4aI260081 under 3.7, sections 2 and 3', [r.first, r.initial.find(p => p[0] === 'S4aI260081'), 'x'].slice(0, 2),
    [null, ['S4aI260081', '3.7', '3.7']]);
  check('rebuild ok', r.error, null);
  check('S4aI260081 exactly once, under 2.7, Agenda Item 2.7', where(r.body, 'S4aI260081'), [['S4aI260081', '2.7', '2.7']]);
  check('section 3 is no longer projected (S4aI260081 was its only TDoc)', topSections(r.body), ['2']);
  check('the MBS 2.x agenda is complete (2.0 ... 2.8)',
    r.body._headingTexts().filter(h => /^2\.\d+ /.test(h)).map(h => h.split(' ')[0]), ['2.0', '2.1', '2.2', '2.3', '2.4', '2.5', '2.6', '2.7', '2.8']);
  check('unrelated TDocs unchanged', placements(r.body).filter(p => p[0] !== 'S4aI260081'), r.initial.filter(p => p[0] !== 'S4aI260081'));
  check('2.7 now holds S4aI260089 and S4aI260081', placements(r.body).filter(p => p[1] === '2.7').map(p => p[0]).sort(), ['S4aI260081', 'S4aI260089']);
  check('the Document Reallocations table is restored with the saved entry', reallocationRows(r.body), [['S4aI260081', '3.7', '2.7', 'test']]);
  check('recorded agenda (PARSED_AGENDA) is the projected MBS section only', JSON.parse(r.docProps._store.PARSED_AGENDA).length, 9);

  const rebuilt = snapshot(r.body);
  const again = r.build();
  check('a second rebuild gives the identical report (reallocation persists)', [again, snapshot(r.body)], [null, rebuilt]);

  let applyError = null;
  try { r.s.applyDocumentReallocations(); } catch (e) { applyError = e.message; }
  check('Apply Document Reallocations afterwards: no error, nothing changes', [applyError, snapshot(r.body)], [null, rebuilt]);
  check('... and it reports nothing moved', /Moved tables: 0\b/.test(r.alerts[r.alerts.length - 1][1]), true);
}

console.log('section projection follows effective allocations generically');
{
  const extra = LIST_86172.concat([['S4aI260090', 'Other video work', 'X', 'discussion', 'Discussion', '27', '3.7', 'Other Rel-20 matters including TEI', 'available', '', '']]);
  const two = rebuildWith([['S4aI260081', '3.7', '2.7']], PROPS_86172, { values: extra });
  check('two TDocs in section 3, one moved: section 3 stays', [two.error, topSections(two.body)], [null, ['2', '3']]);
  check('... the remaining one stays under 3.7', where(two.body, 'S4aI260090'), [['S4aI260090', '3.7', '3.7']]);

  const out = ['S4aI260080', 'S4aI260082', 'S4aI260083', 'S4aI260084', 'S4aI260085', 'S4aI260086', 'S4aI260087', 'S4aI260088', 'S4aI260089']
    .map(id => [id, '2.x', '3.7']);
  const moved = rebuildWith(out);
  check('every TDoc moved out of section 2: section 2 disappears, section 3 stays', [moved.error, topSections(moved.body)], [null, ['3']]);

  const other = rebuildWith([['S4aI260080', '2.5', '4.8']]);
  check('destination in another valid top-level section projects that section', [other.error, topSections(other.body)], [null, ['2', '3', '4']]);
  check('... with the TDoc under 4.8', where(other.body, 'S4aI260080'), [['S4aI260080', '4.8', '4.8']]);
}

console.log('removed / withdrawn / n/a: not placed, and never the reason for a section');
{
  ['withdrawn', 'removed', 'N/A'].forEach(value => {
    const r = rebuildWith([['S4aI260081', '3.7', value]]);
    check(`"${value}": rebuild ok, TDoc not in the report, section 3 not projected`,
      [r.error, where(r.body, 'S4aI260081'), topSections(r.body)], [null, [], ['2']]);
  });
}

console.log('unusable destinations refuse the build BEFORE the document is cleared');
{
  const bad = rebuildWith([['S4aI260081', '3.7', '2.99']]);
  check('not in the agenda: refused with TDoc / source / destination / reason',
    /report was not changed[\s\S]*S4aI260081 \(3\.7 → 2\.99\): this meeting's agenda has no item 2\.99\./.test(bad.error), true);
  check('... document untouched', snapshot(bad.body), bad.beforeRebuild);
  const word = rebuildWith([['S4aI260081', '3.7', 'later']]);
  check('not a number or removal word: refused, document untouched',
    [/"later" is not an agenda item number or removed\/withdrawn\/n\/a\./.test(word.error), snapshot(word.body) === word.beforeRebuild], [true, true]);
}

console.log('no reallocations: identical to ADDON-008A');
{
  const r = makeReport(PROPS_86172);
  check('build ok, sections 2 and 3 from raw Portal data', [r.build(), topSections(r.body)], [null, ['2', '3']]);
  check('S4aI260081 under 3.7', where(r.body, 'S4aI260081'), [['S4aI260081', '3.7', '3.7']]);
  check('no reallocation table is created', reallocationRows(r.body), null);
}

// ================================================ agenda-TDoc path (Audio) ==

console.log('85916 (agenda TDoc S4aA260090): unchanged without reallocations; reallocations honoured on rebuild');
{
  const plain = makeReport(PROPS_85916, { values: LIST_85916, tdocAgenda: AUDIO_AGENDA });
  const plainError = plain.build();
  const plainPlacements = placements(plain.body);
  check('build ok, every TDoc placed at its source item (007A: opening-item TDocs under 1.4 Documents)',
    [plainError, plainPlacements.every(p => p[1] === p[2] || (p[2] === '1' && p[1] === '1.4'))], [null, true]);
  check('no agenda.csv involved', Object.keys(plain.docProps._store).indexOf('AGENDA_CSV_URL'), -1);

  const moved = rebuildWith([['S4aA260092', '4.4', '4.3']], PROPS_85916, { values: LIST_85916, tdocAgenda: AUDIO_AGENDA });
  check('reallocated within the agenda: placed under 4.3', [moved.error, where(moved.body, 'S4aA260092')], [null, [['S4aA260092', '4.3', '4.3']]]);
  check('everything else as before', placements(moved.body).filter(p => p[0] !== 'S4aA260092').map(p => p[1]),
    plainPlacements.filter(p => p[0] !== 'S4aA260092').map(p => p[1]));

  const bad = rebuildWith([['S4aA260092', '4.4', '9.9']], PROPS_85916, { values: LIST_85916, tdocAgenda: AUDIO_AGENDA });
  check('destination not in the Audio agenda: refused before clearing', [/S4aA260092 \(4\.4 → 9\.9\): this meeting's agenda has no item 9\.9\./.test(bad.error), snapshot(bad.body) === bad.beforeRebuild], [true, true]);
}

// ================================================================ main ====

console.log('main meetings: moving a TDoc to another SWG section stays allowed');
{
  const { sandbox: s } = loadCode();
  check('main: another SWG\'s item is not a build problem', s.findReallocationBuildProblems_({ 'S4-260001': { original: '7.3', new: '9.3', reason: '' } }, null, null), []);
  check('main: a non-number is still refused', s.findReallocationBuildProblems_({ 'S4-260001': { original: '7.3', new: 'later', reason: '' } }, null, null).length, 1);
}

// ========================================= one shared interpretation =====

console.log('one interpretation of the saved value, everywhere');
{
  const { sandbox: s } = loadCode();
  check('interpretReallocationTarget_', ['2.7', ' 2.7 ', 'Withdrawn', 'n/a', 'REMOVED', 'later', ''].map(v => s.interpretReallocationTarget_(v).kind),
    ['move', 'move', 'remove', 'remove', 'remove', 'invalid', 'invalid']);
  check('effective allocations never change the source value',
    s.computeEffectiveTdocAllocations_(LIST_86172, { S4aI260081: { original: '3.7', new: '2.7' }, S4aI260080: { original: '2.5', new: 'withdrawn' } })
      .filter(a => a.reallocated).map(a => [a.tdoc, a.source, a.effective, a.removed]),
    [['S4aI260080', '2.5', null, true], ['S4aI260081', '3.7', '2.7', false]]);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
