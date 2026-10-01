/**
 * SA4 Report Template -- report creation from the bound template and the
 * first-run bootstrap (template/ReportCreator.js).
 *
 * Every meeting case runs the REAL Code.js discovery pipeline
 * (resolveMeetingCoreById_ -> enrichMeetingFromTdocList_ ->
 * applyAdhocSourceDiscovery_ -> computeResolvedMeetingPreview_) against the
 * captured 86172/85916 fixtures, and the first run persists through the REAL
 * persistConfigurationSettings_() and is checked with the REAL
 * getBuildReadiness_(). Drive/Docs/ScriptApp are injected fakes that behave
 * as TEMPLATE-001 found live (2026-10-01): the copy is a fresh project with
 * empty property stores and no triggers, and can read its description.
 *
 * Run: node tests/template-bootstrap.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadTemplateRuntime } = require('./helpers/load-template.js');
const { loadCode } = require('./helpers/load-code.js');

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
const plain = (v) => JSON.parse(JSON.stringify(v));

const FX = path.join(__dirname, 'fixtures');
const CSV_86172 = fs.readFileSync(path.join(FX, 'meeting-86172-agenda.csv'), 'utf8');
const CSV_85916 = fs.readFileSync(path.join(FX, 'meeting-85916-agenda.csv'), 'utf8');
const LIST_86172 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-tdoc-list-values.json'), 'utf8')).values;
const LIST_85916 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-85916-tdoc-list-values.json'), 'utf8')).values;
const GM = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-85916-getmeetings.json'), 'utf8'));
const XLSX_BYTES = [0x50, 0x4B, 0x03, 0x04, 0x14];

const TEMPLATE_DOC_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const NEW_DOC_ID = 'NEWREPORTdoc00000000000000000000000000000';
const RELEASE = {
  releaseId: 'T-2026.10.0', flavor: 'template', gitCommit: 'd7c23bc0000000000000000000000000000000aa',
  gitTag: 'template-release/T-2026.10.0', templateDocumentId: TEMPLATE_DOC_ID
};

const ROWS_86172 = [
  { id: 'S4aI260089', type: 'discussion', revisionOf: null, agendaItem: '2.7' },
  { id: 'S4aI260083', type: 'CR', revisionOf: null, agendaItem: '2.5' },
  { id: 'S4aI260081', type: 'draft TS', revisionOf: null, agendaItem: '3.7' }
];
const ROWS_85916 = [
  { id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '2' },
  { id: 'S4aA260092', type: 'discussion', revisionOf: null, agendaItem: '4.4' }
];

/** A template runtime whose network returns one captured meeting. */
function creatorFor(getMeetings, rows, spec) {
  const loaded = loadTemplateRuntime({ release: RELEASE });
  const s = loaded.sandbox;
  s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(getMeetings) });
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => rows;
  s.UrlFetchApp = {
    fetch: (url) => {
      const isCsv = /\/Agenda\/agenda\.csv$/.test(url);
      const r = isCsv ? spec.csv : spec.list;
      if (!r) throw new Error('network down');
      if (isCsv) return { getResponseCode: () => r.status, getContentText: () => r.text };
      return { getResponseCode: () => r.status, getBlob: () => ({ getBytes: () => r.bytes, setName() { return this; } }) };
    }
  };
  let current = null;
  s.DriveApp = { createFile: () => { current = (spec.list || {}).values; return { setTrashed() {} }; } };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => current }) }] }) };
  return s;
}

/** Fake Drive for the creator side; records every call. */
function fakeCreatorDeps(s, overrides) {
  const drive = { calls: [], descriptions: {} };
  const deps = Object.assign({
    release: RELEASE,
    activeDocumentId: () => TEMPLATE_DOC_ID,
    nowIso: () => '2026-09-30T22:00:00.000Z',
    copyTemplate: (title) => { drive.calls.push(['copy', title]); return { id: NEW_DOC_ID, url: 'https://docs.google.com/document/d/' + NEW_DOC_ID + '/edit' }; },
    setDescription: (id, text) => { drive.calls.push(['setDescription', id]); drive.descriptions[id] = text; },
    trashFile: (id) => { drive.calls.push(['trash', id]); }
  }, overrides || {});
  return { deps, drive };
}

/** The NEW report's own runtime: a fresh sandbox = a fresh bound project with empty stores. */
function firstRunFor(description, overrides) {
  const loaded = loadTemplateRuntime({ release: RELEASE, documentProperties: (overrides && overrides.props) || {} });
  const s = loaded.sandbox;
  const state = { description };
  const deps = Object.assign({
    release: RELEASE,
    activeDocumentId: () => NEW_DOC_ID,
    scriptId: () => 'NEWSCRIPTid000000000000000000000000000000000000000',
    nowIso: () => '2026-09-30T22:05:00.000Z',
    documentProperties: loaded.docProps,
    getOwnDescription: () => state.description,
    setOwnDescription: (t) => { state.description = t; },
    persistConfiguration: (config) => s.persistConfigurationSettings_(config),
    buildReadiness: () => plain(s.getBuildReadiness_())
  }, (overrides && overrides.deps) || {});
  return { s, deps, state, docProps: loaded.docProps };
}

function createAndSetUp(getMeetings, rows, spec, meetingId, choices) {
  const s = creatorFor(getMeetings, rows, spec);
  const preview = plain(s.previewNewReportFromTemplate(meetingId, choices));
  const { deps, drive } = fakeCreatorDeps(s);
  const created = plain(s.createReportFromTemplateWith_(deps, { preview: s.computeResolvedMeetingPreview_({}, preview.resolved), choices }));
  const report = firstRunFor(drive.descriptions[NEW_DOC_ID]);
  const setup = plain(report.s.finishReportSetupWith_(report.deps));
  return { preview, created, drive, report, setup };
}

// ============================================================ MBS / agenda.csv

console.log('86172 MBS ad-hoc: agenda.csv, discovered list, family from name + TDocs -> zero questions');
{
  const spec = { csv: { status: 200, text: CSV_86172 }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_86172 } };
  const r = createAndSetUp(GM.getMeetings86172, ROWS_86172, spec, '86172', {});
  check('preview needs nothing from Thomas', [r.preview.ok, r.preview.errors, r.preview.pending], [true, [], []]);
  check('family detected confidently', r.preview.notes, ['Report family MBS detected (confident, from meeting-name + tdoc-family + ftp-path).']);
  check('title is the one the build would give', r.preview.title, 'MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e');
  check('exactly one copy + one description write', r.drive.calls, [['copy', r.preview.title], ['setDescription', NEW_DOC_ID]]);
  check('creator returns the new report link', [r.created.ok, r.created.documentId], [true, NEW_DOC_ID]);
  check('first run stores the configuration', r.setup.status, 'configured');
  const p = r.report.docProps._store;
  check('stored as ordinary Document Properties',
    [p.MEETING_ID, p.MEETING_TYPE, p.REPORT_SUFFIX, p.TDOC_LIST_URL, p.AGENDA_CSV_URL, p.FTP_BASE],
    ['86172', 'adhoc', 'MBS', 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172',
      'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Agenda/agenda.csv',
      'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/']);
  check('mailing list left to family derivation (no override stored)', p.MAILING_LIST, undefined);
  check('e-mail collection starts at the meeting start date', p.EMAIL_START_DATE, '2026-10-01');
  check('no historical SA4#136 folder/number for an ad-hoc meeting', [p.MEETING_FOLDER, p.MEETING_NUMBER], [undefined, undefined]);
  check('build readiness is green (real getBuildReadiness_)', [r.setup.readiness.ready, r.setup.readiness.issues], [true, []]);
  check('release provenance recorded',
    [p.SA4_CREATED_FROM_RELEASE, p.SA4_SETUP_RELEASE, p.SA4_BOOTSTRAP_STATE, p.SA4_BOUND_SCRIPT_ID.slice(0, 9)],
    ['T-2026.10.0', 'T-2026.10.0', 'done', 'NEWSCRIPT']);
  check('payload removed from the description, provenance line left',
    [r.report.state.description.indexOf('SA4-BOOTSTRAP-V1:'), /^SA4 report for meeting 86172 \(MBS\)/.test(r.report.state.description)], [-1, true]);
  check('the Reviewer token is never touched', r.report.s.PropertiesService.getScriptProperties().getKeys(), []);

  const again = plain(r.report.s.finishReportSetupWith_(r.report.deps));
  check('second run is a no-op', again.status, 'already-done');
}

// ============================================================ Audio / agenda TDoc

console.log('85916 Audio ad-hoc: agenda TDoc wins over the 0-byte agenda.csv');
{
  const spec = { csv: { status: 200, text: CSV_85916 }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_85916 } };
  const r = createAndSetUp(GM.getMeetings85916, ROWS_85916, spec, '85916', {});
  check('no questions', [r.preview.ok, r.preview.pending], [true, []]);
  const p = r.report.docProps._store;
  check('agenda TDoc stored, no agenda.csv', [p.AGENDA_TDOC, p.AGENDA_CSV_URL, p.REPORT_SUFFIX], ['S4aA260090', undefined, 'Audio']);
  check('ready to build', r.setup.readiness.ready, true);
}

// ============================================================ missing sources -> pending, not refused

console.log('ad-hoc before the agenda exists: report is created, first run lists what is pending');
{
  const spec = { csv: { status: 200, text: '' }, list: { status: 404, bytes: [], values: [] } };
  const rows = [{ id: 'S4aI260089', type: 'discussion', revisionOf: null, agendaItem: '2.7' }];
  const r = createAndSetUp(GM.getMeetings86172, rows, spec, '86172', {});
  check('creation allowed', r.created.ok, true);
  check('pending items named up front',
    r.preview.pending, ['Discover or select the agenda document.', 'Paste the meeting\'s TDoc list URL.']);
  check('first run: configured but not ready, same messages from the real build guard',
    [r.setup.status, r.setup.readiness.ready, r.setup.readiness.issues.map((i) => i.code)],
    ['configured', false, ['AGENDA_TDOC_REQUIRED', 'TDOC_LIST_REQUIRED']]);
}

// ============================================================ family questions

console.log('ambiguous / conflicting family: exactly one question');
{
  const odd = plain(GM.getMeetings86172);
  odd[0].Title = '3GPPSA4-e (AH) joint Audio and Video session';
  const spec = { csv: { status: 200, text: CSV_86172 }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_86172 } };
  const s = creatorFor(odd, ROWS_86172, spec);
  const refused = plain(s.previewNewReportFromTemplate('86172', {}));
  check('no guess: refused with the family question', [refused.ok, refused.errors], [false, ['Choose the report family (the evidence conflicts).']]);
  const answered = plain(s.previewNewReportFromTemplate('86172', { reportFamily: 'Video' }));
  check('an explicit choice resolves it', [answered.ok, answered.title], [true, 'Video SWG Minutes – SA4-e (AH) joint Audio and Video session']);
  const bogus = plain(s.previewNewReportFromTemplate('86172', { reportFamily: 'Plenary' }));
  check('an unknown family is refused', bogus.errors[0], 'Unknown report family "Plenary".');
}

console.log('6G / plenary-style ad-hoc: family from the FS_6G_MED name (S4aP TDocs never decide)');
{
  const sixg = plain(GM.getMeetings86172);
  sixg[0].Id = 86178; sixg[0].Title = '3GPPSA4-e (AH) on FS_6G_MED';
  sixg[0].MtgDocURL = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/';
  const rows = [{ id: 'S4aP260074', type: 'agenda', revisionOf: null, agendaItem: '1' }];
  const spec = { csv: { status: 200, text: '' }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_86172 } };
  const s = creatorFor(sixg, rows, spec);
  const r = plain(s.previewNewReportFromTemplate('86178', {}));
  check('6G suggested from the name', [r.ok, r.notes], [true, ['Report family 6G detected (suggested, from meeting-name).']]);
  check('title keeps the Legacy naming', r.title, '6G Media Minutes – SA4-e (AH) on FS_6G_MED');
}

// ============================================================ main meeting

console.log('main meeting: family is always asked, folder/number derived from the FTP folder');
{
  const main = plain(GM.getMeetings86172);
  main[0].Type = 'OR'; main[0].Title = '3GPPSA4#137'; main[0].MtgDocURL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Xian/Docs/';
  const s = creatorFor(main, [], { csv: null, list: null });
  const noFamily = plain(s.previewNewReportFromTemplate('86172', {}));
  const mainType = noFamily.resolved.meeting.type;
  if (mainType !== 'main') {
    check('fixture sanity: portal type OR resolves to main', mainType, 'main');
  } else {
    check('family question', noFamily.errors, ['Choose the report family -- a main meeting has one report per family.']);
    const r = plain(s.previewNewReportFromTemplate('86172', { reportFamily: 'Audio' }));
    check('then no further questions', [r.ok, r.pending], [true, []]);
    check('title SA4#137 (not the portal id)', r.title, 'Audio SWG Minutes SA4#137');
    const { deps, drive } = fakeCreatorDeps(s);
    s.createReportFromTemplateWith_(deps, { preview: s.computeResolvedMeetingPreview_({}, r.resolved), choices: { reportFamily: 'Audio' } });
    const report = firstRunFor(drive.descriptions[NEW_DOC_ID]);
    report.s.finishReportSetupWith_(report.deps);
    const p = report.docProps._store;
    check('folder + number stored (never the SA4#136 defaults)', [p.MEETING_FOLDER, p.MEETING_NUMBER, p.MEETING_TYPE], ['TSGS4_137_Xian', '137', 'main']);
    const cfg = report.s.getReportConfig_();
    check('runtime derives the main TDoc list from them',
      cfg.TDOC_LIST_URL, 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Xian/Docs/TDoc_List_Meeting_SA4%23137.xlsx');
  }
}

console.log('main folder derivation (pure)');
{
  const { sandbox: s } = loadTemplateRuntime();
  check('Xian', plain(s.deriveMainMeetingFolderFromFtpBase_('https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Xian/Docs/')),
    { folder: 'TSGS4_137_Xian', number: '137', error: null });
  check('e-meeting keeps -e', plain(s.deriveMainMeetingFolderFromFtpBase_('https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137-e/Docs/')),
    { folder: 'TSGS4_137-e', number: '137-e', error: null });
  check('ad-hoc series folder is not a main folder',
    s.deriveMainMeetingFolderFromFtpBase_('https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/').folder, null);
}

console.log('unresolvable meeting: nothing is copied');
{
  const s = creatorFor([], [], { csv: null, list: null });
  const r = plain(s.previewNewReportFromTemplate('99999', {}));
  check('refused', r.ok, false);
  const bad = plain(s.previewNewReportFromTemplate('abc', {}));
  check('invalid id refused before any fetch', bad.ok, false);
}

// ============================================================ guards

console.log('creator guards');
{
  const s = creatorFor(GM.getMeetings86172, ROWS_86172, { csv: { status: 200, text: CSV_86172 }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_86172 } });
  const preview = plain(s.previewNewReportFromTemplate('86172', {}));
  const pv = s.computeResolvedMeetingPreview_({}, preview.resolved);

  const fromReport = fakeCreatorDeps(s, { activeDocumentId: () => NEW_DOC_ID });
  check('refused outside the master template, nothing copied',
    [s.createReportFromTemplateWith_(fromReport.deps, { preview: pv, choices: {} }).ok, fromReport.drive.calls], [false, []]);

  const noRelease = fakeCreatorDeps(s, { release: null });
  check('refused without Release.js (e.g. a CENTRAL/Legacy copy)',
    [s.createReportFromTemplateWith_(noRelease.deps, { preview: pv, choices: {} }).ok, noRelease.drive.calls], [false, []]);

  const failing = fakeCreatorDeps(s, { setDescription: () => { throw new Error('quota'); } });
  const f = plain(s.createReportFromTemplateWith_(failing.deps, { preview: pv, choices: {} }));
  check('a copy that cannot be prepared goes to the trash', [f.ok, failing.drive.calls.map((c) => c[0])], [false, ['copy', 'trash']]);

  const evil = plain(pv);
  evil.tdocListUrl = { value: 'https://evil.example/list.xlsx', source: 'existing' };
  const e = fakeCreatorDeps(s);
  const er = plain(s.createReportFromTemplateWith_(e.deps, { preview: evil, choices: {} }));
  check('non-3GPP URL refused before copying', [er.ok, e.drive.calls], [false, []]);
}

console.log('first-run guards');
{
  const s = creatorFor(GM.getMeetings86172, ROWS_86172, { csv: { status: 200, text: CSV_86172 }, list: { status: 200, bytes: XLSX_BYTES, values: LIST_86172 } });
  const preview = plain(s.previewNewReportFromTemplate('86172', {}));
  const { deps, drive } = fakeCreatorDeps(s);
  s.createReportFromTemplateWith_(deps, { preview: s.computeResolvedMeetingPreview_({}, preview.resolved), choices: {} });
  const description = drive.descriptions[NEW_DOC_ID];

  const inTemplate = firstRunFor(description, { deps: { activeDocumentId: () => TEMPLATE_DOC_ID } });
  check('never configures the template itself',
    [inTemplate.s.finishReportSetupWith_(inTemplate.deps).status, inTemplate.docProps.getKeys()], ['refused', []]);

  const copyOfCopy = firstRunFor(description, { deps: { activeDocumentId: () => 'SOMEOTHERdoc000000000000000000000000000000' } });
  const cc = plain(copyOfCopy.s.finishReportSetupWith_(copyOfCopy.deps));
  check('a copy of a not-yet-set-up report is refused (payload is for another doc)',
    [cc.status, cc.errors, copyOfCopy.docProps.getKeys()],
    ['refused', ['The setup information was created for a different document. Create the report again from the template.'], []]);

  const tampered = description.replace('"meetingType":"adhoc"', '"meetingType":"adhoc","apiToken":"x"');
  const t = firstRunFor(tampered);
  check('unexpected payload fields refused', plain(t.s.finishReportSetupWith_(t.deps)).errors, ['Unexpected setup field "apiToken".']);

  const damaged = firstRunFor(description.slice(0, description.length - 5));
  check('damaged payload refused', plain(damaged.s.finishReportSetupWith_(damaged.deps)).errors, ['The setup information in the file description is damaged.']);

  const legacy = firstRunFor('', { props: { MEETING_ID: '86178', MEETING_TYPE: 'adhoc' } });
  check('an existing/migrated report without payload keeps its configuration; only the "checked once" marker is added',
    [legacy.s.finishReportSetupWith_(legacy.deps).status, plain(legacy.docProps._store)],
    ['not-a-new-report', { MEETING_ID: '86178', MEETING_TYPE: 'adhoc', SA4_BOOTSTRAP_STATE: 'manual' }]);
  check('... and is not looked at again', legacy.s.finishReportSetupWith_(Object.assign({}, legacy.deps, { getOwnDescription: () => { throw new Error('read again'); } })).status, 'not-a-new-report');

  const blank = firstRunFor('');
  check('an unconfigured document without setup information: nothing to finish, no configuration written',
    [blank.s.finishReportSetupWith_(blank.deps).status, plain(blank.docProps._store)], ['no-setup-info', { SA4_BOOTSTRAP_STATE: 'manual' }]);

  const other = firstRunFor(description, { props: { MEETING_ID: '12345' } });
  check('never overwrites a different meeting', plain(other.s.finishReportSetupWith_(other.deps)).errors,
    ['This document is already configured for meeting 12345.']);
}

// ============================================================ template role + intervals

console.log('template role and release metadata');
{
  const { sandbox: s } = loadTemplateRuntime({ release: RELEASE });
  check('template id -> template', s.templateDocumentRole_(TEMPLATE_DOC_ID, RELEASE), 'template');
  check('any other id -> report', s.templateDocumentRole_(NEW_DOC_ID, RELEASE), 'report');
  check('release metadata valid', plain(s.validateTemplateRelease_(RELEASE)), []);
  check('Release.js is picked up at call time', s.templateRuntimeRelease_().releaseId, 'T-2026.10.0');
  const { sandbox: none } = loadTemplateRuntime();
  check('without Release.js the runtime reports no release', none.templateRuntimeRelease_(), null);
}

// ============================================================ isolation from Code.js

console.log('ReportCreator.js does not redefine anything in Code.js');
{
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
  const creator = fs.readFileSync(path.join(__dirname, '..', 'template', 'ReportCreator.js'), 'utf8');
  const names = (src) => new Set([...src.matchAll(/^(?:function|var|const|let)\s+([A-Za-z0-9_$]+)/gm)].map((m) => m[1]));
  const codeNames = names(code);
  check('no shared top-level names', [...names(creator)].filter((n) => codeNames.has(n)), []);
  check('no Node-only code in the Apps Script file', /\brequire\s*\(/.test(creator), false);
  const { sandbox: c1 } = loadCode();
  const { sandbox: c2 } = loadTemplateRuntime({ release: RELEASE });
  check('Code.js behaves identically with ReportCreator.js loaded (sample: title + readiness)',
    [c2.generateReportTitle_({ REPORT_SUFFIX: 'MBS', meetingLabel: 'X' }), plain(c2.getBuildReadiness_())],
    [c1.generateReportTitle_({ REPORT_SUFFIX: 'MBS', meetingLabel: 'X' }), plain(c1.getBuildReadiness_())]);
}

console.log(failures ? `\n${failures} FAILURE(S)` : '\nall ok');
process.exitCode = failures ? 1 : 0;
