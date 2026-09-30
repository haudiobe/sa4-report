/**
 * ADDON-008A -- robust ad-hoc agenda + TDoc-list discovery.
 *
 * Real cases (captured live 2026-09-30, see tests/fixtures/):
 *   - 86172 "SA4-e (AH) MBS SWG post 137-e": no agenda-typed TDoc; the
 *     series agenda.csv (53 headerless "number","title" rows, combined SA4
 *     agenda) matches the meeting's TDoc list (2.5 / 2.7 / 3.7).
 *   - 85916 "SA4-(AH) Audio SWG on ULBC-MED": agenda TDoc S4aA260090; the
 *     series agenda.csv is 0 bytes.
 *
 * Agenda structure precedence (ad-hoc only): saved AGENDA_TDOC >
 * discovered agenda TDoc > validated agenda.csv > readiness error. The
 * meeting-ID Portal document list is discovered; a saved TDOC_LIST_URL wins.
 * Only the network/Drive/Sheets boundaries are faked.
 *
 * Run: node tests/addon008a-adhoc-agenda-sources.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
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
const CSV_85916 = fs.readFileSync(path.join(FX, 'meeting-85916-agenda.csv'), 'utf8');
const LIST_86172 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-tdoc-list-values.json'), 'utf8')).values;
const LIST_85916 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-85916-tdoc-list-values.json'), 'utf8')).values;
const GM = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-85916-getmeetings.json'), 'utf8'));

const FTP_MBS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/';
const CSV_URL_MBS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Agenda/agenda.csv';
const FTP_AUDIO = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Docs/';
const CSV_URL_AUDIO = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/Agenda/agenda.csv';
const LIST_URL_86172 = 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172';
const LIST_URL_85916 = 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916';
const XLSX_BYTES = [0x50, 0x4B, 0x03, 0x04, 0x14];

// TdocList.aspx rows as the real parser produced them on 2026-09-30.
const ROWS_86172 = [
  { id: 'S4aI260089', type: 'discussion', revisionOf: null, agendaItem: '2.7' },
  { id: 'S4aI260083', type: 'CR', revisionOf: null, agendaItem: '2.5' },
  { id: 'S4aI260082', type: 'CR', revisionOf: null, agendaItem: '2.5' },
  { id: 'S4aI260081', type: 'draft TS', revisionOf: null, agendaItem: '3.7' },
  { id: 'S4aI260080', type: 'CR', revisionOf: null, agendaItem: '2.5' }
];
const ROWS_85916 = [
  { id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '2' },
  { id: 'S4aA260092', type: 'discussion', revisionOf: null, agendaItem: '4.4' }
];

/**
 * Fakes UrlFetchApp/DriveApp/SpreadsheetApp. `csv` / `list`: { status, text }
 * / { status, bytes, values }, or 'throw'. Records every fetched URL.
 */
function installNetwork(sandbox, spec) {
  const calls = [];
  sandbox.UrlFetchApp = {
    fetch: (url) => {
      calls.push(url);
      const isCsv = /\/Agenda\/agenda\.csv$/.test(url);
      const s = isCsv ? spec.csv : spec.list;
      if (!s || s === 'throw') throw new Error('network down');
      if (isCsv) return { getResponseCode: () => s.status, getContentText: () => s.text };
      return { getResponseCode: () => s.status, getBlob: () => ({ getBytes: () => s.bytes, setName() { return this; } }) };
    }
  };
  let current = null;
  sandbox.DriveApp = { createFile: () => { current = (spec.list || {}).values; return { setTrashed() {} }; } };
  sandbox.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => current }) }] }) };
  return calls;
}
const csvCalls = (calls) => calls.filter(u => /agenda\.csv$/.test(u));

// ================================================================ pure ====

console.log('agenda.csv candidate: only the ad-hoc series layout on the 3GPP FTP');
{
  const { sandbox: s } = loadCode();
  check('86172 folder -> <series>/Agenda/agenda.csv', s.deriveAdhocAgendaCsvCandidate_(FTP_MBS).url, CSV_URL_MBS);
  check('www.3gpp.org/ftp host variant is accepted',
    s.deriveAdhocAgendaCsvCandidate_('https://www.3gpp.org/ftp/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/').url,
    'https://www.3gpp.org/ftp/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Agenda/agenda.csv');
  check('main-meeting folder has no candidate', s.deriveAdhocAgendaCsvCandidate_('https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/').url, null);
  check('plain http is rejected', s.deriveAdhocAgendaCsvCandidate_(FTP_MBS.replace('https:', 'http:')).url, null);
  check('a foreign host is rejected', s.deriveAdhocAgendaCsvCandidate_('https://evil.example/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/').url, null);
  check('a ".." series is rejected', s.deriveAdhocAgendaCsvCandidate_('https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/../Docs/').url, null);
  check('blank folder has no candidate', s.deriveAdhocAgendaCsvCandidate_('').url, null);
  check('TDoc list candidate is meeting-ID based', s.buildAdhocTdocListCandidateUrl_('86172'), LIST_URL_86172);
  check('no TDoc list candidate without a valid meeting ID', s.buildAdhocTdocListCandidateUrl_('86172x'), null);
}

console.log('parseAgendaCsv_: the real 86172 file maps onto the existing agenda item shape');
{
  const { sandbox: s } = loadCode();
  const r = s.parseAgendaCsv_(CSV_86172);
  check('parses', [r.ok, r.items.length], [true, 53]);
  check('first row', r.items[0], { number: '1.0', title: 'Audio SWG', level: 2, heading: 'NORMAL', text: '' });
  check('2.5 title preserved', r.items.find(i => i.number === '2.5').title, 'AMD_ARCH_Ph2-MED (Architectural Updates for Advanced Media Delivery Phase 2)');
  check('hierarchy preserved (5.6.1 is level 3)', r.items.find(i => i.number === '5.6.1').level, 3);
  check('file order preserved', r.items.slice(10, 13).map(i => i.number), ['2.0', '2.1', '2.2']);
}

console.log('parseAgendaCsv_: unusable files are rejected as a whole');
{
  const { sandbox: s } = loadCode();
  check('the real 0-byte Audio file is empty', [s.parseAgendaCsv_(CSV_85916).ok, s.parseAgendaCsv_(CSV_85916).reason], [false, 'agenda.csv is empty.']);
  check('whitespace only is empty', s.parseAgendaCsv_('\r\n \r\n').ok, false);
  check('a header row is malformed', s.parseAgendaCsv_('"Number","Title"\n"2.1","Opening"\n').ok, false);
  check('a non-numbered row is malformed', s.parseAgendaCsv_('"2.1","Opening"\n"x","Oops"\n').ok, false);
  check('a single-column row is malformed', s.parseAgendaCsv_('"2.1"\n').ok, false);
  check('an extra non-empty column is malformed', s.parseAgendaCsv_('"2.1","Opening","extra"\n').ok, false);
  check('unbalanced quotes are malformed', s.parseAgendaCsv_('"2.1","Opening\n').ok, false);
  check('duplicate agenda numbers are rejected', s.parseAgendaCsv_('"2.1","A"\n"2.1","B"\n').ok, false);
  check('HTML is rejected', s.parseAgendaCsv_('<html><body>Error</body></html>').ok, false);
  const q = s.parseAgendaCsv_('﻿"2.1","Opening, registration"\r\n"2.2","Say ""hi"""\r\n');
  check('BOM, CRLF, quoted commas and "" escapes', [q.ok, q.items.map(i => i.title)], [true, ['Opening, registration', 'Say "hi"']]);
}

console.log('TDoc list values: importer columns, plausibility and agenda pairs');
{
  const { sandbox: s } = loadCode();
  check('86172 list is valid (10 TDocs)', s.validateTdocListWorkbookValues_(LIST_86172), { ok: true, tdocCount: 10, reason: null });
  check('85916 list is valid (95 TDocs)', s.validateTdocListWorkbookValues_(LIST_85916).tdocCount, 95);
  check('missing Agenda item column is rejected', s.validateTdocListWorkbookValues_([['TDoc', 'Title'], ['S4aI260080', 'x']]).ok, false);
  check('no recognizable SA4 TDoc rows is rejected', s.validateTdocListWorkbookValues_([['TDoc', 'Agenda item'], ['hello', '1']]).ok, false);
  check('86172 pairs are its real agenda items', s.extractTdocListAgendaPairs_(LIST_86172).pairs, [
    { number: '2.5', title: 'AMD_ARCH_Ph2-MED (Architectural Updates for Advanced Media Delivery Phase 2)' },
    { number: '3.7', title: 'Other Rel-20 matters including TEI' },
    { number: '2.7', title: 'Other Rel-20 matters including TEI' }
  ]);
  const conflicting = [LIST_86172[0], ['S4aI260080', '', '', '', '', '1', '2.5', 'One title'], ['S4aI260081', '', '', '', '', '1', '2.5', 'Another title']];
  check('a list describing one item two ways is not usable evidence', s.extractTdocListAgendaPairs_(conflicting).ok, false);
  check('XLSX signature', [s.isXlsxSignature_(XLSX_BYTES), s.isXlsxSignature_([0x3C, 0x68, 0x74, 0x6D]), s.isXlsxSignature_([])], [true, false, false]);
}

console.log('validateAgendaCsvForMeeting_: the series CSV must match THIS meeting');
{
  const { sandbox: s } = loadCode();
  const items = s.parseAgendaCsv_(CSV_86172).items;
  const v = s.validateAgendaCsvForMeeting_(items, s.extractTdocListAgendaPairs_(LIST_86172));
  check('86172 validates; sections come from its TDocs (2 and 3, not a hard-coded MBS=2)', [v.ok, v.sections, v.matchedCount], [true, ['2', '3'], 3]);

  const unrelated = s.parseAgendaCsv_('"9.0","Other SWG"\n"9.1","Opening of the session"\n').items;
  const m = s.validateAgendaCsvForMeeting_(unrelated, s.extractTdocListAgendaPairs_(LIST_86172));
  check('numbered rows without this meeting\'s items are rejected', [m.ok, /does not contain agenda item\(s\) 2\.5, 3\.7, 2\.7/.test(m.reason)], [false, true]);

  const renamed = s.parseAgendaCsv_(CSV_86172.replace('"2.5","AMD_ARCH_Ph2-MED', '"2.5","Something else')).items;
  const c = s.validateAgendaCsvForMeeting_(renamed, s.extractTdocListAgendaPairs_(LIST_86172));
  check('a conflicting title is rejected', [c.ok, /titles differ .* 2\.5/.test(c.reason)], [false, true]);

  const spacing = s.parseAgendaCsv_(CSV_86172.replace('"2.7","Other Rel-20 matters including TEI"', '"2.7","  other  Rel-20 matters including TEI "')).items;
  check('whitespace/case differences are tolerated', s.validateAgendaCsvForMeeting_(spacing, s.extractTdocListAgendaPairs_(LIST_86172)).ok, true);

  const noPairs = [LIST_86172[0], ['S4aI260080', 't', 's', 'CR', 'Agreement', '', '', '', 'available', '', '']];
  const n = s.validateAgendaCsvForMeeting_(items, s.extractTdocListAgendaPairs_(noPairs));
  check('no agenda pairs to validate against -> not accepted automatically', [n.ok, /cannot be checked/.test(n.reason)], [false, true]);

  const sel = s.selectAdhocCsvAgendaItems_(items, v.sections);
  check('selection keeps only the sections this meeting uses, in CSV order', [sel.length, sel[0].number, sel[sel.length - 1].number], [18, '2.0', '3.8']);
  check('no other subgroup section leaks in', sel.some(i => /^[145]\./.test(i.number)), false);
  check('selection returns the parsed item objects (existing projection)', sel[0] === items.find(i => i.number === '2.0'), true);
}

// =========================================================== readiness ====

console.log('readiness: an agenda TDoc OR a validated agenda.csv satisfies the agenda rule');
{
  const { sandbox: s } = loadCode();
  const base = { meetingType: 'adhoc', meetingId: '86172', meetingName: 'SA4-e (AH) MBS SWG post 137-e', ftpBase: FTP_MBS,
    reportFamily: 'MBS', tdocListUrl: LIST_URL_86172, mailingList: '' };
  const rules = vm.runInContext('MEETING_READINESS_RULES_', s);
  const evaluate = (f) => s.evaluateMeetingReadiness_(f, rules);
  check('agenda.csv only -> ready', evaluate(Object.assign({}, base, { agendaCsvUrl: CSV_URL_MBS })).ready, true);
  check('agenda TDoc only -> ready', evaluate(Object.assign({}, base, { agendaTdoc: 'S4aA260090' })).ready, true);
  const neither = evaluate(base);
  check('neither -> the same 007B3 issue and message', neither.issues.map(i => [i.code, i.message]), [['AGENDA_TDOC_REQUIRED', 'Discover or select the agenda document.']]);
  check('other ad-hoc requirements are unchanged (TDoc list still required)',
    evaluate(Object.assign({}, base, { agendaCsvUrl: CSV_URL_MBS, tdocListUrl: '' })).issues.map(i => i.code), ['TDOC_LIST_REQUIRED']);
}

const ADHOC_86172_PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) MBS SWG post 137-e', MEETING_ID: '86172', REPORT_SUFFIX: 'MBS',
  FTP_BASE: FTP_MBS, TDOC_LIST_URL: LIST_URL_86172, AGENDA_CSV_URL: CSV_URL_MBS };

console.log('build readiness: the saved agenda.csv must be this meeting\'s own candidate (server-side)');
{
  check('86172 with its own agenda.csv -> build-ready', loadCode({ documentProperties: ADHOC_86172_PROPS }).sandbox.getBuildReadiness_().ready, true);
  const forged = Object.assign({}, ADHOC_86172_PROPS, { AGENDA_CSV_URL: 'https://evil.example/agenda.csv' });
  check('a forged agenda.csv URL does not count', loadCode({ documentProperties: forged }).sandbox.getBuildReadiness_().issues.map(i => i.code), ['AGENDA_TDOC_REQUIRED']);
  const otherSeries = Object.assign({}, ADHOC_86172_PROPS, { AGENDA_CSV_URL: CSV_URL_AUDIO });
  check('another series\' agenda.csv does not count', loadCode({ documentProperties: otherSeries }).sandbox.getBuildReadiness_().ready, false);
  const main = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136', REPORT_SUFFIX: 'Audio' };
  check('main meetings are never build-blocked (unchanged)', loadCode({ documentProperties: main }).sandbox.getBuildReadiness_().ready, true);
}

// ============================================================ discover ====

function discover(props, rows, spec, meetingKey) {
  const loaded = loadCode({ documentProperties: props || {} });
  const s = loaded.sandbox;
  s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(GM[meetingKey || 'getMeetings86172']) });
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => rows;
  const calls = installNetwork(s, spec);
  const id = meetingKey === 'getMeetings85916' ? '85916' : '86172';
  const core = s.resolveMeetingForConfigDialog_(id);
  const callsAfterResolve = calls.length;
  const result = JSON.parse(JSON.stringify(s.discoverAgendaForConfigDialog_(id, JSON.parse(JSON.stringify(core.resolved)))));
  return { result, calls, callsAfterResolve, sandbox: s, docProps: loaded.docProps };
}
const OK_86172 = { csv: { status: 200, text: CSV_86172 }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_86172 } };
const OK_85916 = { csv: { status: 200, text: CSV_85916 }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_85916 } };

console.log('Discover 86172: no agenda TDoc, meeting-specific list + validated agenda.csv');
{
  const d = discover({}, ROWS_86172, OK_86172);
  const r = d.result;
  check('Resolve alone makes no new network call', d.callsAfterResolve, 0);
  check('discover ok', r.ok, true);
  check('Portal document list discovered', r.resolved.adhocSources.tdocList, { url: LIST_URL_86172, ok: true, tdocCount: 10, reason: null });
  check('agenda.csv validated', [r.resolved.adhocSources.agendaCsv.ok, r.resolved.adhocSources.agendaCsv.url, r.resolved.adhocSources.agendaCsv.sections],
    [true, CSV_URL_MBS, ['2', '3']]);
  check('"no agenda TDoc" is no longer reported as a problem', r.resolved.warnings.indexOf('No TDoc with type "agenda" was found.'), -1);
  check('preview: TDoc list found automatically', r.preview.tdocListUrl, { value: LIST_URL_86172, source: 'resolved' });
  check('preview: agenda.csv found automatically', r.preview.agendaCsvUrl, { value: CSV_URL_MBS, source: 'resolved' });
  check('preview: agenda TDoc still unresolved (nothing invented)', r.preview.agendaTdoc, { value: '', source: 'unresolved' });
  check('agenda source is described positively',
    r.preview.agendaSourceStatus, 'Agenda structure: validated agenda.csv fallback (no agenda TDoc yet) — 18 of 53 items, section(s) 2, 3.');
  check('TDoc list provenance', r.preview.tdocListStatus, 'Meeting-specific Portal document list (10 TDocs).');
  check('007B2 inference unchanged: MBS, confident', [r.preview.familyInference.family, r.preview.familyInference.confidence], ['MBS', 'confident']);
  check('remote agenda titles are not sent to the client', /Gaussian/.test(JSON.stringify(r)), false);
  check('nothing was saved by Discover', d.docProps.getKeys(), []);
}

console.log('Discover 85916: the agenda TDoc wins; the 0-byte agenda.csv is never consulted');
{
  const d = discover({}, ROWS_85916, OK_85916, 'getMeetings85916');
  const r = d.result;
  check('agenda TDoc S4aA260090', r.preview.agendaTdoc, { value: 'S4aA260090', source: 'resolved' });
  check('agenda.csv not fetched', csvCalls(d.calls), []);
  check('agenda.csv not checked', [r.resolved.adhocSources.agendaCsv.checked, r.preview.agendaCsvUrl.value], [false, '']);
  check('agenda source text', r.preview.agendaSourceStatus, 'Agenda TDoc: S4aA260090');
  check('document list discovered (95 TDocs)', [r.preview.tdocListUrl.value, r.resolved.adhocSources.tdocList.tdocCount], [LIST_URL_85916, 95]);
  check('family still Audio', r.preview.familyInference.family, 'Audio');
}

console.log('agenda TDoc + usable agenda.csv -> the TDoc wins, the CSV is not used');
{
  const rows = ROWS_86172.concat([{ id: 'S4aI260090', type: 'agenda', revisionOf: null, agendaItem: '2.1' }]);
  const d = discover({}, rows, OK_86172);
  check('agenda TDoc resolved', d.result.preview.agendaTdoc.value, 'S4aI260090');
  check('agenda.csv not fetched', csvCalls(d.calls), []);
  check('agenda source is the TDoc', d.result.preview.agendaSourceStatus, 'Agenda TDoc: S4aI260090');
}

console.log('a saved agenda TDoc wins over the CSV fallback');
{
  const d = discover({ AGENDA_TDOC: 'S4aI260099' }, ROWS_86172, OK_86172);
  check('saved agenda TDoc kept', d.result.preview.agendaTdoc, { value: 'S4aI260099', source: 'existing' });
  check('agenda.csv not fetched', csvCalls(d.calls), []);
}

console.log('ambiguous agenda TDocs are not silently replaced by the CSV');
{
  const rows = ROWS_86172.concat([
    { id: 'S4aI260090', type: 'agenda', revisionOf: null, agendaItem: '2.1' },
    { id: 'S4aI260091', type: 'agenda', revisionOf: null, agendaItem: '2.1' }
  ]);
  const d = discover({}, rows, OK_86172);
  check('candidates surfaced, CSV not fetched', [d.result.preview.agendaCandidates, csvCalls(d.calls)], [['S4aI260090', 'S4aI260091'], []]);
}

console.log('failures degrade to reasons; Discover itself still succeeds');
{
  const thrown = discover({}, ROWS_86172, { csv: 'throw', list: OK_86172.list }).result;
  check('CSV fetch failure', [thrown.ok, thrown.resolved.adhocSources.agendaCsv.ok, /could not be fetched/.test(thrown.resolved.adhocSources.agendaCsv.reason)], [true, false, true]);
  check('the no-agenda warning stays when no fallback validates', thrown.resolved.warnings.indexOf('No TDoc with type "agenda" was found.') !== -1, true);
  check('agenda source explains why', /agenda\.csv fallback is not usable: agenda\.csv could not be fetched/.test(thrown.preview.agendaSourceStatus), true);
  check('TDoc list still discovered', thrown.preview.tdocListUrl.source, 'resolved');

  const notFound = discover({}, ROWS_86172, { csv: { status: 404, text: '' }, list: OK_86172.list }).result;
  check('CSV HTTP 404', notFound.resolved.adhocSources.agendaCsv.reason, 'agenda.csv returned HTTP 404.');

  const empty = discover({}, ROWS_86172, { csv: { status: 200, text: '' }, list: OK_86172.list }).result;
  check('0-byte CSV is a failed fallback, not an error', [empty.ok, empty.resolved.adhocSources.agendaCsv.reason, empty.preview.agendaCsvUrl.value], [true, 'agenda.csv is empty.', '']);

  const html = discover({}, ROWS_86172, { csv: OK_86172.csv, list: { status: 200, bytes: [0x3C, 0x21, 0x44, 0x4F], values: null } }).result;
  check('non-XLSX document list is rejected', [html.resolved.adhocSources.tdocList.ok, html.resolved.adhocSources.tdocList.reason, html.preview.tdocListUrl.value],
    [false, 'The document list is not an Excel workbook.', '']);
  check('without a valid list the CSV cannot be validated', /could not be checked/.test(html.resolved.adhocSources.agendaCsv.reason), true);

  const err = discover({}, ROWS_86172, { csv: OK_86172.csv, list: { status: 500, bytes: [], values: null } }).result;
  check('document list HTTP 500', err.resolved.adhocSources.tdocList.reason, 'The document list returned HTTP 500.');
  check('status text names the failure', /not usable: The document list returned HTTP 500\./.test(err.preview.tdocListStatus), true);
}

console.log('a saved TDoc list URL always wins');
{
  const d = discover({ TDOC_LIST_URL: 'https://example.invalid/manual.xlsx' }, ROWS_86172, OK_86172);
  check('preview keeps the saved URL', d.result.preview.tdocListUrl, { value: 'https://example.invalid/manual.xlsx', source: 'existing' });
}

console.log('main meetings: no ad-hoc discovery at all');
{
  const mainMeeting = JSON.parse(JSON.stringify(GM.getMeetings86172));
  mainMeeting[0].Type = 'TSG';
  mainMeeting[0].MtgDocURL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137/Docs/';
  const loaded = loadCode();
  const s = loaded.sandbox;
  s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(mainMeeting) });
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => ROWS_86172;
  const calls = installNetwork(s, OK_86172);
  const core = s.resolveMeetingForConfigDialog_('86172');
  if (core.resolved.meeting.type !== 'adhoc') {
    const r = s.discoverAgendaForConfigDialog_('86172', core.resolved);
    check('no extra fetches, no ad-hoc sources', [calls, r.resolved.adhocSources === undefined], [[], true]);
    check('no TDoc list proposed for a main meeting', r.preview.tdocListUrl, { value: '', source: 'unresolved' });
  } else {
    check('fixture sanity: meeting is not ad-hoc', core.resolved.meeting.type, 'main');
  }
}

// ================================================================ save ====

console.log('Save: only this meeting\'s own agenda.csv is stored (server-side)');
{
  const saveWith = (props, extra) => {
    const loaded = loadCode({ documentProperties: props || {} });
    loaded.sandbox.saveConfigurationSettings(Object.assign({
      meetingId: '86172', meetingType: 'adhoc', meetingName: 'SA4-e (AH) MBS SWG post 137-e', meetingDate: 'October 1, 2026',
      ftpBase: FTP_MBS, reportType: 'MBS', agendaTdoc: '', tdocUrl: LIST_URL_86172, showPreview: true, apiTokenAction: 'keep'
    }, extra));
    return loaded.docProps._store;
  };
  check('valid candidate stored', saveWith({}, { agendaCsvUrl: CSV_URL_MBS }).AGENDA_CSV_URL, CSV_URL_MBS);
  check('client cannot store a foreign URL', saveWith({}, { agendaCsvUrl: 'https://evil.example/agenda.csv' }).AGENDA_CSV_URL, undefined);
  check('client cannot store another series\' CSV', saveWith({}, { agendaCsvUrl: CSV_URL_AUDIO }).AGENDA_CSV_URL, undefined);
  check('client cannot claim "validated"', saveWith({}, { agendaCsvUrl: CSV_URL_MBS + '?validated=true' }).AGENDA_CSV_URL, undefined);
  check('not stored for a main meeting', saveWith({}, { meetingType: 'main', agendaCsvUrl: CSV_URL_MBS }).AGENDA_CSV_URL, undefined);
  check('blank keeps an existing value (skip-if-blank)', saveWith({ AGENDA_CSV_URL: CSV_URL_MBS }, { agendaCsvUrl: '' }).AGENDA_CSV_URL, CSV_URL_MBS);
  const { sandbox: s } = loadCode();
  check('AGENDA_CSV_URL is dialog-managed and adopted (central copy)',
    [s.CONFIG_DIALOG_MANAGED_KEYS_.indexOf('AGENDA_CSV_URL') !== -1, s.ADDON003_ADOPTION_FIXED_KEYS_.indexOf('AGENDA_CSV_URL') !== -1], [true, true]);
}

// ============================================================== dialog ====

const IDS = ['meetingId', 'agendaStructure', 'tdocUrlHint', 'mailingListHint', 'mailingListReset', 'familyStatus', 'mainMeetingFields', 'meetingType', 'reportType',
  'familyInfo', 'readinessRules', 'readinessStatus', 'meetingSummary', 'resolveStatus', 'resolveBtn', 'discoverStatus', 'discoverBtn', 'portalTypeHint',
  'dateRangeHint', 'agendaCandidates', 'meetingName', 'meetingDate', 'ftpBase', 'ftpBaseBadge', 'agendaTdoc', 'agendaTdocBadge', 'mailingList', 'revisionsUrl',
  'revisionsUrlBadge', 'apiToken', 'clearApiToken', 'meetingFolder', 'meetingNumber', 'agendaSourceDocId', 'tdocUrl', 'showPreview',
  'agendaSourceStatus', 'agendaCsvUrl'];

function openDialog(props, rows, spec) {
  const loaded = loadCode({ documentProperties: props || {} });
  const s = loaded.sandbox;
  let html = null;
  s.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  s.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  s.configureMeetingSettings();
  s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(GM.getMeetings86172) });
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => rows;
  installNetwork(s, spec);

  const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')).replace(/\\\\/g, '\\');
  const registry = {};
  IDS.forEach((id) => { registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }; });
  const value = (id) => { const m = html.match(new RegExp('id="' + id + '" value="([^"]*)"')); return m ? m[1].replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&') : ''; };
  ['familyInfo', 'readinessRules', 'meetingId', 'meetingType', 'meetingName', 'mailingList', 'ftpBase', 'agendaTdoc', 'tdocUrl', 'meetingFolder', 'meetingNumber', 'agendaCsvUrl']
    .forEach((id) => { registry[id].value = value(id); });
  let sent = null;
  const runner = () => {
    let ok = null;
    const r = {
      withSuccessHandler: (fn) => { ok = fn; return r; },
      withFailureHandler: () => r,
      resolveMeetingForConfigDialog: (id) => ok(JSON.parse(JSON.stringify(s.resolveMeetingForConfigDialog(id)))),
      discoverAgendaForConfigDialog: (id, core) => ok(JSON.parse(JSON.stringify(s.discoverAgendaForConfigDialog(id, core)))),
      saveConfigurationSettings: (config) => { sent = config; ok(); }
    };
    return r;
  };
  const globals = {
    document: { getElementById: (id) => registry[id] || null },
    google: { script: { get run() { return runner(); }, host: { close: () => {} } } },
    alert: () => {},
    console: { log: () => {} }
  };
  vm.createContext(globals);
  vm.runInContext(script, globals);
  return { registry, globals, sandbox: s, docProps: loaded.docProps, html, sent: () => sent };
}

console.log('dialog 86172: Resolve -> Discover -> Ready to build -> Save');
{
  const d = openDialog({}, ROWS_86172, OK_86172);
  d.registry.meetingId.value = '86172';
  d.globals.resolveMeeting();
  check('after Resolve: agenda and TDoc list outstanding', /agenda document/.test(d.registry.readinessStatus.innerHTML) && /TDoc list URL/.test(d.registry.readinessStatus.innerHTML), true);
  d.globals.discoverAgendaTdocs();
  check('TDoc list field filled with the meeting-specific list', d.registry.tdocUrl.value, LIST_URL_86172);
  check('its hint names the source', d.registry.tdocUrlHint.textContent, 'Meeting-specific Portal document list (found automatically).');
  check('agenda.csv carried in the hidden field', d.registry.agendaCsvUrl.value, CSV_URL_MBS);
  check('agenda source shown positively', /validated agenda\.csv fallback/.test(d.registry.agendaSourceStatus.textContent), true);
  check('no "no agenda TDoc" alarm in the status line', /No TDoc with type/.test(d.registry.discoverStatus.textContent), false);
  check('family MBS selected by inference', d.registry.reportType.value, 'MBS');
  check('Ready to build', [/Ready to build/.test(d.registry.readinessStatus.innerHTML), d.registry.readinessStatus.className], [true, 'ready']);
  check('agenda TDoc field left empty (nothing invented)', d.registry.agendaTdoc.value, '');

  d.globals.saveConfig();
  const sent = d.sent();
  check('Save sends the CSV and the list', [sent.agendaCsvUrl, sent.tdocUrl, sent.agendaTdoc], [CSV_URL_MBS, LIST_URL_86172, '']);
  d.sandbox.saveConfigurationSettings(sent);
  const store = d.docProps._store;
  check('saved: agenda.csv, list, family, FTP base; no agenda TDoc', [store.AGENDA_CSV_URL, store.TDOC_LIST_URL, store.REPORT_SUFFIX, store.FTP_BASE, store.AGENDA_TDOC],
    [CSV_URL_MBS, LIST_URL_86172, 'MBS', FTP_MBS, undefined]);
  check('saved configuration is build-ready (server rules)', d.sandbox.getBuildReadiness_().ready, true);
}

console.log('dialog: a typed TDoc list URL is never replaced by Discover');
{
  const d = openDialog({}, ROWS_86172, OK_86172);
  d.registry.meetingId.value = '86172';
  d.globals.resolveMeeting();
  d.registry.tdocUrl.value = 'https://example.invalid/typed.xlsx';
  d.globals.discoverAgendaTdocs();
  check('typed value kept', d.registry.tdocUrl.value, 'https://example.invalid/typed.xlsx');
}

console.log('dialog: without discovery the ad-hoc hint is the 007B1 wording');
{
  const d = openDialog({ MEETING_TYPE: 'adhoc' }, ROWS_86172, OK_86172);
  d.globals.updateDependentUi();
  check('asks to paste the list', /Paste the meeting's TDoc list URL\./.test(d.registry.tdocUrlHint.textContent), true);
}

// =============================================================== build ====

function runBuild(props, spec, opts) {
  const loaded = loadCode({ documentProperties: props });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  body.appendParagraph('EXISTING REPORT CONTENT');
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.getUi = () => ({ alert() {}, ButtonSet: { OK: 'OK' } });
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  s.removeRowHeightAndSpacing = () => {};
  const calls = installNetwork(s, spec);
  const agendaDownloads = [];
  s.downloadMeetingAgenda_ = (t) => { agendaDownloads.push(t); return (opts && opts.tdocAgenda) || []; };
  const values = (opts && opts.groupsFrom) || LIST_86172;
  const groups = {};
  const h = values[0];
  values.slice(1).forEach(r => { const k = r[h.indexOf('Agenda item')]; (groups[k] = groups[k] || { tdocs: [] }).tdocs.push(tdoc(r[0], k)); });
  s.downloadAndGroupTdocs_ = () => groups;
  let error = null;
  try { s.buildSkeletonWithTdocTables(); } catch (e) { error = e.message; }
  return { body, calls, agendaDownloads, error, docProps: loaded.docProps };
}

console.log('build 86172 from the validated agenda.csv through the normal skeleton path');
{
  const b = runBuild(ADHOC_86172_PROPS, OK_86172);
  const h = b.body._headingTexts();
  check('no error', b.error, null);
  check('the MBS section headings come from the CSV', ['2.0 Multicast-Broadcast-Streaming (MBS) SWG', '2.5 AMD_ARCH_Ph2-MED (Architectural Updates for Advanced Media Delivery Phase 2)',
    '2.7 Other Rel-20 matters including TEI', '2.8 Close of the session'].every(t => h.indexOf(t) !== -1), true);
  check('3.7 is present because a TDoc of this meeting is filed there', h.indexOf('3.7 Other Rel-20 matters including TEI') !== -1, true);
  check('no other subgroup section (1.x Audio, 4.x RTC, 5.x 6G)', h.some(t => /^[145]\.\d/.test(t)), false);
  check('007A anchors: opening/registration under 2.1, IPR under 2.2 (not 1.1 Audio)',
    ['2.1.1 Opening of the session', '2.1.2 Registration of Documents'].every(t => h.indexOf(t) !== -1) && h.indexOf('2.2 IPR, antitrust & consensus principles reminder') !== -1, true);
  check('old content replaced', b.body._children.some(c => c.getText && c.getText() === 'EXISTING REPORT CONTENT'), false);
  check('no agenda TDoc download attempted', b.agendaDownloads, []);
  check('parsed agenda recorded as before (PARSED_AGENDA)', JSON.parse(b.docProps._store.PARSED_AGENDA).length, 18);
}

console.log('build fails BEFORE touching the document when the CSV fallback no longer validates');
{
  const intact = (b) => b.body._children.length === 1 && b.body._children[0].getText() === 'EXISTING REPORT CONTENT';
  const empty = runBuild(ADHOC_86172_PROPS, { csv: { status: 200, text: '' }, list: OK_86172.list });
  check('0-byte CSV at build: refused, document intact', [/Cannot build report yet[\s\S]*agenda\.csv is empty\./.test(empty.error), intact(empty)], [true, true]);
  const changed = runBuild(ADHOC_86172_PROPS, { csv: { status: 200, text: CSV_86172.replace('"2.5","AMD_ARCH_Ph2-MED', '"2.5","Renamed') }, list: OK_86172.list });
  check('conflicting CSV at build: refused, document intact', [/titles differ/.test(changed.error), intact(changed)], [true, true]);
  const listDown = runBuild(ADHOC_86172_PROPS, { csv: OK_86172.csv, list: 'throw' });
  check('TDoc list unavailable at build: refused, document intact', [/could not be fetched/.test(listDown.error), intact(listDown)], [true, true]);
  const forged = runBuild(Object.assign({}, ADHOC_86172_PROPS, { AGENDA_CSV_URL: CSV_URL_AUDIO }), OK_86172);
  check('another series\' CSV: refused by readiness, nothing fetched, document intact', [/Cannot build report yet/.test(forged.error), forged.calls, intact(forged)], [true, [], true]);
}

console.log('build 85916: the agenda TDoc is used even when an agenda.csv URL is also saved');
{
  const props = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED', MEETING_ID: '85916', REPORT_SUFFIX: 'Audio',
    FTP_BASE: FTP_AUDIO, TDOC_LIST_URL: LIST_URL_85916, AGENDA_TDOC: 'S4aA260090', AGENDA_CSV_URL: CSV_URL_AUDIO };
  const agenda = [
    { number: '1', title: 'Opening of the meeting', level: 1, heading: 'NORMAL', text: '' },
    { number: '2', title: 'Approval of the agenda and registration of documents', level: 1, heading: 'NORMAL', text: '' },
    { number: '4', title: 'ULBC-MED', level: 1, heading: 'NORMAL', text: '' },
    { number: '4.4', title: 'Design constraints', level: 2, heading: 'NORMAL', text: '' }
  ];
  const b = runBuild(props, OK_85916, { tdocAgenda: agenda, groupsFrom: LIST_85916 });
  check('no error', b.error, null);
  check('agenda TDoc S4aA260090 downloaded', b.agendaDownloads, ['S4aA260090']);
  check('agenda.csv never fetched, no extra list download', b.calls, []);
  check('agenda TDoc structure rendered', b.body._headingTexts().indexOf('4.4 Design constraints') !== -1, true);
}

console.log('main meetings: the CSV path never applies');
{
  const { sandbox: s } = loadCode({ documentProperties: { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136', AGENDA_CSV_URL: CSV_URL_MBS, FTP_BASE: FTP_MBS } });
  check('prepareAdhocCsvAgendaForBuild_ returns null', s.prepareAdhocCsvAgendaForBuild_(s.getReportConfig_()), null);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
