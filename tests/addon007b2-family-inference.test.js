/**
 * ADDON-007B2 -- report-family auto-detection + derived mailing list.
 *
 * Pure inference (table driven, generic), meeting-type scoping, the dialog's
 * client behaviour (Resolve / Discover / manual override) run against the
 * REAL client script and real server functions, mailing-list
 * derived-vs-override save semantics, and central-state sync.
 *
 * Run: node tests/addon007b2-family-inference.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

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

const FIXTURES = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-resolver-samples.json'), 'utf8'));
const { sandbox: pure } = loadCode();
const infer = (evidence) => JSON.parse(JSON.stringify(pure.inferReportFamily_(evidence)));

// ------------------------------------------------------------ pure inference

console.log('inferReportFamily_ -- generic, table driven');

const NAME_ONLY = [
  ['Audio', 'SA4-(AH) Audio SWG on ULBC-MED', 'Audio'],
  ['Video', 'SA4-(AH) Video SWG on something', 'Video'],
  ['MBS', 'SA4-(AH) MBS SWG', 'MBS'],
  ['RTC', 'SA4-(AH) RTC SWG', 'RTC'],
  ['6G (FS_6G_MED)', 'SA4-e (AH) on FS_6G_MED', '6G'],
  ['6G (plain)', 'SA4 ad hoc on 6G media', '6G']
];
NAME_ONLY.forEach(([label, name, family]) => {
  const r = infer({ meetingType: 'adhoc', meetingName: name });
  check(`name only ${label} -> suggested ${family}`, [r.family, r.confidence, r.sources], [family, 'suggested', ['meeting-name']]);
});

const TDOC_ONLY = [
  ['S4aA', 'audio-adhoc', 'Audio'],
  ['S4aV', 'video-adhoc', 'Video'],
  ['S4aI', 'mbs-adhoc', 'MBS'],
  ['A4aR', 'rtc-adhoc', 'RTC']
];
TDOC_ONLY.forEach(([prefix, key, family]) => {
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc', tdocFamilyKeys: { [key]: 3 } });
  check(`${prefix} documents only -> suggested ${family}`, [r.family, r.confidence, r.sources, r.tdocPrefixes], [family, 'suggested', ['tdoc-family'], [prefix]]);
});

NAME_ONLY.slice(0, 4).forEach(([label, name, family], i) => {
  const r = infer({ meetingType: 'adhoc', meetingName: name, tdocFamilyKeys: { [TDOC_ONLY[i][1]]: 2 } });
  check(`name + documents agree (${label}) -> confident`, [r.family, r.confidence, r.sources], [family, 'confident', ['meeting-name', 'tdoc-family']]);
});

{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-e (AH) on FS_6G_MED', tdocFamilyKeys: { 'plenary-adhoc': 5 } });
  check('6G name + S4aP documents: S4aP adds nothing, stays suggested 6G', [r.family, r.confidence], ['6G', 'suggested']);
}

{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc', tdocFamilyKeys: { 'plenary-adhoc': 5 } });
  check('S4aP alone -> unresolved (no family, never 6G/Liaison/New)', [r.family, r.confidence, r.reason], [null, 'unresolved', 'ambiguous-tdoc-series']);
}

{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc' });
  check('no evidence -> unresolved', [r.family, r.confidence, r.reason], [null, 'unresolved', 'no-evidence']);
}

{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG', tdocFamilyKeys: { 'video-adhoc': 4 } });
  check('name Audio + S4aV documents -> conflict, nothing chosen', [r.family, r.confidence], [null, 'conflict']);
  const flipped = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Video SWG', tdocFamilyKeys: { 'audio-adhoc': 4 } });
  check('conflict result does not depend on which source is processed first', [flipped.family, flipped.confidence], [null, 'conflict']);
}

{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio and Video SWG' });
  check('multiple incompatible name tokens -> conflict', [r.family, r.confidence], [null, 'conflict']);
  const docs = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc', tdocFamilyKeys: { 'audio-adhoc': 1, 'video-adhoc': 1 } });
  check('mixed strong TDoc series -> conflict', [docs.family, docs.confidence], [null, 'conflict']);
}

console.log('inferReportFamily_ -- conservative matching');

['Audiobook workshop', 'Rtcweb', 'Newsletter Liaison New items', 'Multimbs2', 'SA4 16G'].forEach((name) => {
  check(`no substring false positive: "${name}"`, infer({ meetingType: 'adhoc', meetingName: name }).family, null);
});

{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG', ftpBase: 'https://example.invalid/ftp/Audio/Docs/' });
  check('FTP path agreeing with the strong evidence is listed as supporting', r.sources, ['meeting-name', 'ftp-path']);
  const conflicting = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG', ftpBase: 'https://example.invalid/ftp/Video/Docs/' });
  check('FTP path never overrides stronger evidence', [conflicting.family, conflicting.confidence, conflicting.sources], ['Audio', 'suggested', ['meeting-name']]);
  const alone = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc', ftpBase: 'https://example.invalid/ftp/Audio/Docs/' });
  check('FTP path alone selects nothing', alone.family, null);
}

console.log('inferReportFamily_ -- supported families come from existing constants');

{
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const body = src.slice(src.indexOf('function inferReportFamily_('), src.indexOf('function buildMeetingSummary_('));
  check('no meeting-specific or ULBC-specific runtime rule', /85916|ULBC/i.test(body), false);
  check('no hard-coded S4aA/S4aV/S4aI/A4aR mapping in the helper (the registry is reused)', /S4aA|S4aV|S4aI|A4aR/.test(body), false);
  const registry = JSON.parse(vm.runInContext('JSON.stringify(SA4_TDOC_FAMILIES)', pure)).filter((f) => f.reportFamily).map((f) => [f.prefix, f.reportFamily]);
  check('canonical registry maps the four SWG series and leaves S4aP/S4- unmapped',
    registry, [['S4aA', 'Audio'], ['S4aV', 'Video'], ['S4aI', 'MBS'], ['A4aR', 'RTC']]);
}

// ------------------------------------------------------------ meeting type

console.log('meeting type scoping');

{
  const r = infer({ meetingType: 'main', meetingName: 'SA4#137 Audio and Video', tdocFamilyKeys: { 'audio-adhoc': 1 } });
  check('main meeting + Audio-like meeting-level evidence -> no auto-selection', [r.applicable, r.family, r.confidence], [false, null, 'none']);
  check('a not-yet-typed meeting is not inferred either', infer({ meetingName: 'Audio' }).applicable, false);
  check('ad-hoc -> inference allowed', infer({ meetingType: 'adhoc', meetingName: 'Audio' }).applicable, true);

  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'main', MEETING_NAME: 'SA4#137 Audio' } });
  const preview = sandbox.computeResolvedMeetingPreview_({ MEETING_TYPE: 'main', MEETING_NAME: 'SA4#137 Audio' }, null);
  check('preview of a main meeting carries a non-applicable inference', preview.familyInference.applicable, false);
}

// ------------------------------------------------------------ dialog client

const IDS = ['meetingId', 'agendaStructure', 'tdocUrlHint', 'mailingListHint', 'mailingListReset', 'familyStatus', 'mainMeetingFields', 'meetingType', 'reportType',
  'familyInfo', 'meetingSummary', 'resolveStatus', 'resolveBtn', 'discoverStatus', 'discoverBtn', 'portalTypeHint', 'dateRangeHint', 'agendaCandidates',
  'meetingName', 'meetingDate', 'ftpBase', 'ftpBaseBadge', 'agendaTdoc', 'agendaTdocBadge', 'mailingList', 'revisionsUrl', 'revisionsUrlBadge', 'apiToken',
  'clearApiToken', 'meetingFolder', 'meetingNumber', 'agendaSourceDocId', 'tdocUrl', 'showPreview'];

// Renders the real dialog for a document with `documentProperties`, and runs
// its real client script with google.script.run wired to the real server
// functions of the same sandbox (network stubbed with the captured fixtures).
function openDialog(documentProperties, meetingFixture, tdocRows) {
  const loaded = loadCode({ documentProperties: documentProperties || {} });
  const sandbox = loaded.sandbox;
  let html = null;
  sandbox.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.configureMeetingSettings();

  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(meetingFixture) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  sandbox.parseMeetingTdocListHtml_ = () => tdocRows || [];

  const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')).replace(/\\\\/g, '\\');
  const registry = {};
  IDS.forEach((id) => { registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }; });
  const value = (id) => { const m = html.match(new RegExp('id="' + id + '" value="([^"]*)"')); return m ? m[1].replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&') : ''; };
  ['familyInfo', 'meetingId', 'meetingType', 'meetingName', 'mailingList', 'ftpBase', 'agendaTdoc'].forEach((id) => { registry[id].value = value(id); });
  const selected = html.match(/<option value="([A-Za-z0-9]+)" selected>/);
  registry.reportType.value = selected ? selected[1] : '';

  let sent = null;
  const runner = () => {
    let ok = null;
    const r = {
      withSuccessHandler: (fn) => { ok = fn; return r; },
      withFailureHandler: () => r,
      resolveMeetingForConfigDialog: (id) => ok(JSON.parse(JSON.stringify(sandbox.resolveMeetingForConfigDialog(id)))),
      discoverAgendaForConfigDialog: (id, core) => ok(JSON.parse(JSON.stringify(sandbox.discoverAgendaForConfigDialog(id, core)))),
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
  return { registry, globals, sandbox, loaded, sentConfig: () => sent, html };
}

// mimics the user picking an option in the select
function pickFamily(d, family) { d.registry.reportType.value = family; d.globals.onFamilyChanged(); }
function typeMailing(d, text) { d.registry.mailingList.value = text; d.globals.onMailingListEdited(); }

const M85916 = FIXTURES.getMeetings85916;
const M86178 = FIXTURES.getMeetings86178;
const AUDIO_ROWS = [
  { id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '' },
  { id: 'S4aA260091', type: 'pCR', revisionOf: null, agendaItem: '4' }
];

console.log('client: Resolve on an Audio ad-hoc proposes Audio immediately (before Discover)');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  check('precondition: fresh dialog has no family and no mailing list', [d.registry.reportType.value, d.registry.mailingList.value], ['', '']);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Report Family becomes Audio', d.registry.reportType.value, 'Audio');
  check('status says it was suggested from the meeting name', d.registry.familyStatus.textContent, 'Audio — suggested from meeting name');
  check('mailing list immediately shows the derived Audio default', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_AUDIO');
  check('agenda structure stays separate and never says Audio 7.x', d.registry.agendaStructure.textContent, 'Ad-hoc — use complete discovered agenda');
  check('the derived value is not offered as an override reset', d.registry.mailingListReset.style.display, 'none');

  console.log('client: Discover adds S4aA evidence and confirms it');
  d.globals.discoverAgendaTdocs();
  check('family unchanged', d.registry.reportType.value, 'Audio');
  check('status is confirmed by meeting name and S4aA documents', d.registry.familyStatus.textContent, 'Audio — confirmed by meeting name and S4aA documents');
  check('agenda structure still not tied to the family', /7\.x|Audio/.test(d.registry.agendaStructure.textContent), false);
}

console.log('client: 6G ad-hoc title suggests 6G and derives its list');

{
  const d = openDialog({}, M86178, []);
  d.registry.meetingId.value = '86178';
  d.globals.resolveMeeting();
  check('FS_6G_MED title -> 6G', d.registry.reportType.value, '6G');
  check('status', d.registry.familyStatus.textContent, '6G — suggested from meeting name');
}

console.log('client: manual selection is never overwritten');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  d.registry.meetingId.value = '85916';
  pickFamily(d, 'Video');
  d.globals.resolveMeeting();
  check('Resolve does not replace the manual choice', d.registry.reportType.value, 'Video');
  d.globals.discoverAgendaTdocs();
  check('Discover does not replace the manual choice', d.registry.reportType.value, 'Video');
  check('the disagreement is surfaced, not resolved', d.registry.familyStatus.textContent, 'Video — selected (meeting information suggests Audio)');
  check('mailing list follows the chosen family', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_VIDEO');
}

console.log('client: a saved family (existing explicit REPORT_SUFFIX) wins over inference');

{
  const d = openDialog({ REPORT_SUFFIX: 'Video', MEETING_TYPE: 'adhoc' }, M85916, AUDIO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('saved Video survives an Audio-named ad-hoc', d.registry.reportType.value, 'Video');
}

console.log('client: conflicts are surfaced and never silently resolved');

{
  const VIDEO_ROWS = [{ id: 'S4aV260010', type: 'agenda', revisionOf: null, agendaItem: '' }];
  const d = openDialog({}, M85916, VIDEO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Resolve first suggests Audio', d.registry.reportType.value, 'Audio');
  d.globals.discoverAgendaTdocs();
  check('Discover finds S4aV documents: the suggestion is withdrawn, not switched', d.registry.reportType.value, '');
  check('the conflict is stated', d.registry.familyStatus.textContent, 'Needs input — conflicting family information found');
  check('the derived mailing list is withdrawn with it', d.registry.mailingList.value, '');
  pickFamily(d, 'Video');
  check('the user can then choose; the choice stands', [d.registry.reportType.value, d.registry.familyStatus.textContent], ['Video', '']);
}

console.log('client: unresolved evidence asks for input');

{
  const noHint = JSON.parse(JSON.stringify(M85916));
  noHint[0].Title = 'SA4-(AH) ad hoc session';
  const d = openDialog({}, noHint, [{ id: 'S4aP260001', type: 'agenda', revisionOf: null, agendaItem: '' }]);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  d.globals.discoverAgendaTdocs();
  check('S4aP alone leaves the family unselected', d.registry.reportType.value, '');
  check('status asks for input', d.registry.familyStatus.textContent, 'Needs input — family could not be determined');
}

console.log('client: main meeting never auto-selects a family');

{
  const d = openDialog({}, FIXTURES.getMeetings60778, []);
  d.registry.meetingId.value = '60778';
  d.globals.resolveMeeting();
  check('type is main', d.registry.meetingType.value, 'main');
  check('family stays blank', d.registry.reportType.value, '');
  check('no inference status shown', d.registry.familyStatus.textContent, '');
  check('agenda structure asks for a family', d.registry.agendaStructure.textContent, 'Main meeting — select a report family.');
  pickFamily(d, 'Audio');
  check('a manual Audio on main keeps the section wording', d.registry.agendaStructure.textContent, 'Main meeting — Audio section 7.x');
}

// ------------------------------------------------------------ mailing list

console.log('mailing list: derived vs explicit override (client)');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  pickFamily(d, 'Audio');
  check('Audio becomes effective -> derived Audio list shown', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_AUDIO');
  pickFamily(d, 'Video');
  check('derived list follows the family change', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_VIDEO');

  typeMailing(d, 'MY_OWN_LIST');
  check('an edit marks the value as overridden', /Custom value/.test(d.registry.mailingListHint.textContent), true);
  check('the reset link is offered', d.registry.mailingListReset.style.display, '');
  pickFamily(d, 'MBS');
  check('the explicit override survives a family change', d.registry.mailingList.value, 'MY_OWN_LIST');
  check('the hint names the default it overrides', /3GPP_TSG_SA_WG4_MBS/.test(d.registry.mailingListHint.textContent), true);

  d.globals.resetMailingList();
  check('reset returns to the family-derived default', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_MBS');
  check('and derivation resumes on the next family change', (pickFamily(d, 'RTC'), d.registry.mailingList.value), '3GPP_TSG_SA_WG4_RTC');
}

console.log('mailing list: save payload (client)');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  pickFamily(d, 'Audio');
  d.globals.saveConfig();
  check('derived default is sent as mode "derived" with no value', [d.sentConfig().mailingListMode, d.sentConfig().mailingList], ['derived', '']);
  typeMailing(d, 'MY_OWN_LIST');
  d.globals.saveConfig();
  check('explicit edit is sent as an override', [d.sentConfig().mailingListMode, d.sentConfig().mailingList], ['override', 'MY_OWN_LIST']);
  d.globals.resetMailingList();
  d.globals.saveConfig();
  check('reset is sent as derived again', [d.sentConfig().mailingListMode, d.sentConfig().mailingList], ['derived', '']);
  typeMailing(d, '   ');
  d.globals.saveConfig();
  check('a blanked override field counts as derived', d.sentConfig().mailingListMode, 'derived');
}

console.log('mailing list: an existing saved override is shown as an override and survives Resolve');

{
  const d = openDialog({ MAILING_LIST: 'LEGACY_LIST', REPORT_SUFFIX: 'Audio', MEETING_TYPE: 'adhoc' }, M85916, AUDIO_ROWS);
  check('saved list shown', d.registry.mailingList.value, 'LEGACY_LIST');
  check('marked as custom', /Custom value/.test(d.registry.mailingListHint.textContent), true);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Resolve keeps the saved override', d.registry.mailingList.value, 'LEGACY_LIST');
  d.globals.saveConfig();
  check('saving untouched keeps it an override', [d.sentConfig().mailingListMode, d.sentConfig().mailingList], ['override', 'LEGACY_LIST']);
  d.globals.resetMailingList();
  d.globals.resolveMeeting();
  check('a reset done in the dialog is not undone by a later Resolve', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_AUDIO');
}

// ------------------------------------------------------------ save semantics

console.log('save semantics (server)');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: '', mailingListMode: 'derived', showPreview: true });
  check('family persisted', docProps._store.REPORT_SUFFIX, 'Audio');
  check('a derived mailing list is NOT persisted as an override', docProps._store.MAILING_LIST, undefined);
  check('runtime still derives the Audio list for the persisted family', sandbox.getReportConfig_().LIST_NAME, '3GPP_TSG_SA_WG4_AUDIO');
}

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: ' MY_LIST ', mailingListMode: 'override', showPreview: true });
  check('an explicit override is persisted (trimmed)', docProps._store.MAILING_LIST, 'MY_LIST');
  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: '', mailingListMode: 'derived', showPreview: true });
  check('reset deletes the explicit MAILING_LIST so derivation resumes', docProps._store.MAILING_LIST, undefined);
  check('the property is truly absent', Object.prototype.hasOwnProperty.call(docProps._store, 'MAILING_LIST'), false);
}

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MAILING_LIST: 'LEGACY_LIST' } });
  sandbox.saveConfigurationSettings({ meetingId: '1', reportType: 'Audio', mailingList: '', showPreview: true });
  check('legacy caller (no mode) with a blank list keeps the old skip-if-blank behaviour', docProps._store.MAILING_LIST, 'LEGACY_LIST');
  sandbox.saveConfigurationSettings({ meetingId: '1', reportType: 'Audio', mailingList: 'NEW_LEGACY', showPreview: true });
  check('legacy caller with a value still sets it', docProps._store.MAILING_LIST, 'NEW_LEGACY');
}

console.log('existing documents: nothing is rewritten or deleted before Save');

{
  const before = { REPORT_SUFFIX: 'Audio', MAILING_LIST: '3GPP_TSG_SA_WG4_AUDIO', MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED' };
  const d = openDialog(before, M85916, AUDIO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  d.globals.discoverAgendaTdocs();
  check('dialog open + Resolve + Discover leave every legacy property untouched (even one equal to the derived value)',
    JSON.parse(JSON.stringify(d.loaded.docProps._store)), before);
  check('the explicit REPORT_SUFFIX is what the dialog shows', d.registry.reportType.value, 'Audio');
}

console.log('registered central document: set/delete reach central state through the B1 sync');

{
  const DOC = 'DOC_B2_REGISTERED';
  const BACKGROUND = { documentId: DOC, mode: 'addon-background' };
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Video', MAILING_LIST: 'OLD_OVERRIDE' } });
  sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => DOC, getBody: () => ({ getTables: () => [] }) });
  sandbox.adoptReportDocumentForAddon_({ documentId: DOC, mode: 'addon-interactive' });
  sandbox.registerReportDocument_(DOC, { enabled: true, intervalHours: 2 });
  const central = sandbox.getReportStateStore_(BACKGROUND);
  check('precondition: adopted central copy has the old override', central.getProperty('MAILING_LIST'), 'OLD_OVERRIDE');

  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: '', mailingListMode: 'derived', showPreview: true });
  check('derived save deletes the override centrally too (no stale central value)', central.getProperty('MAILING_LIST'), null);
  check('background config derives the Audio list from the central family', sandbox.getReportConfig_(BACKGROUND).LIST_NAME, '3GPP_TSG_SA_WG4_AUDIO');
  check('background identity has no explicit override', sandbox.getMeetingIdentityConfig_(BACKGROUND).MAILING_LIST, '');
  check('central family updated', central.getProperty('REPORT_SUFFIX'), 'Audio');

  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: 'EXPLICIT', mailingListMode: 'override', showPreview: true });
  check('an explicit override is set centrally', central.getProperty('MAILING_LIST'), 'EXPLICIT');
  check('background identity sees it', sandbox.getMeetingIdentityConfig_(BACKGROUND).MAILING_LIST, 'EXPLICIT');
  check('Document Properties and central agree', docProps._store.MAILING_LIST, central.getProperty('MAILING_LIST'));
}

// ------------------------------------------------------------ regression

console.log('regression');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingType: 'main', meetingId: '', meetingFolder: '', meetingNumber: '', reportType: '6G', showPreview: true });
  check('main-meeting save keeps the historical defaults', [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID], ['TSGS4_136_Montreal', '136', '60777']);
  check('main-meeting save with no mailing list mode writes no override', docProps._store.MAILING_LIST, undefined);

  const golden = fs.existsSync(path.join(__dirname, 'fixtures', 'main-meeting-profile.expected.json'));
  check('golden report-structure fixture is still present (asserted by report-structure.test.js)', golden, true);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
