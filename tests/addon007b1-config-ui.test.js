/**
 * ADDON-007B1 -- Meeting Configuration dialog cleanup + save/state consistency.
 *
 * Covers: fresh-document defaults are not presented or persisted, ad-hoc vs
 * main save semantics, family/agenda-structure wording, removal of the dead
 * Email Start Date field, Reviewer token rendering/logging/replace/clear,
 * and (the invariant) Save on a registered central-add-on document making
 * the central state used by background runs immediately consistent.
 *
 * Run: node tests/addon007b1-config-ui.test.js
 */

const fs = require('fs');
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

const REAL_LOOKING_TOKEN = 'crv1_TESTONLY_not_a_real_token_0123456789';
const NEW_TOKEN = 'crv1_TESTONLY_replacement_token_9876543210';

// Renders the real dialog by capturing what configureMeetingSettings() hands
// to HtmlService, without any real Apps Script UI.
function renderDialog(options) {
  const loaded = loadCode(options);
  let captured = null;
  loaded.sandbox.HtmlService = {
    createHtmlOutput: (html) => {
      captured = html;
      const out = { setWidth: () => out, setHeight: () => out };
      return out;
    }
  };
  loaded.sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  loaded.sandbox.configureMeetingSettings();
  return Object.assign({ html: captured }, loaded);
}

function withActiveDocument(sandbox, documentId) {
  sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => documentId, getBody: () => ({ getTables: () => [] }) });
}

function captureLogs(sandbox) {
  const lines = [];
  sandbox.Logger = { log: (m) => lines.push(String(m)) };
  return lines;
}

// Runs the dialog's REAL client script against a minimal fake DOM.
function runClientScript(html, values) {
  const scriptText = html.slice(html.indexOf('<script>') + '<script>'.length, html.lastIndexOf('</script>'));
  const ids = ['meetingId', 'agendaStructure', 'tdocUrlHint', 'mailingListHint', 'mainMeetingFields', 'meetingType', 'reportType', 'familyInfo',
    'meetingSummary', 'resolveStatus', 'resolveBtn', 'discoverStatus', 'discoverBtn', 'portalTypeHint', 'dateRangeHint', 'agendaCandidates',
    'meetingName', 'meetingDate', 'ftpBase', 'agendaTdoc', 'mailingList', 'revisionsUrl', 'apiToken', 'clearApiToken', 'meetingFolder',
    'meetingNumber', 'agendaSourceDocId', 'tdocUrl', 'showPreview'];
  const registry = {};
  ids.forEach((id) => { registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }; });
  // the hidden familyInfo input carries the server-rendered JSON (HTML-unescaped by a real browser)
  const m = html.match(/id="familyInfo" value="([^"]*)"/);
  registry.familyInfo.value = m[1].replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  Object.keys(values || {}).forEach((k) => { registry[k].value = values[k]; });
  const globals = {
    document: { getElementById: (id) => registry[id] || null },
    google: { script: { run: { withSuccessHandler: () => ({ withFailureHandler: () => ({}) }) }, host: { close: () => {} } } },
    alert: () => {},
    console: { log: () => {} }
  };
  vm.createContext(globals);
  vm.runInContext(scriptText.replace(/\\\\/g, '\\'), globals);
  return { registry, globals };
}

// ------------------------------------------------------------ 1. fresh doc

console.log('fresh central-add-on document: no SA4#136 / fallback-ID / template default is presented as configuration');

{
  const { html } = renderDialog();
  check('no Montreal folder text anywhere in the rendered dialog', /Montreal/.test(html), false);
  check('no "136" meeting-number value', /id="meetingNumber" value="136"/.test(html), false);
  check('Meeting Folder value is blank', /id="meetingFolder" value=""/.test(html), true);
  check('Meeting Number value is blank', /id="meetingNumber" value=""/.test(html), true);
  check('no fallback portal meeting ID 60777', /60777/.test(html), false);
  check('Meeting ID input is blank', /id="meetingId" value=""/.test(html), true);
  check('the shared template default is not presented as a prefilled value', /id="agendaSourceDocId" value=""/.test(html), true);
  check('summary says nothing is resolved yet', /No meeting resolved yet\./.test(html), true);
}

console.log('legacy runtime fallbacks are untouched (getReportConfig_ on a fresh store)');

{
  const { sandbox } = loadCode();
  const cfg = sandbox.getReportConfig_();
  check('MEETING_FOLDER fallback intact', cfg.MEETING_FOLDER, 'TSGS4_136_Montreal');
  check('MEETING_NUMBER fallback intact', cfg.MEETING_NUMBER, '136');
  check('MEETING_ID fallback intact', cfg.MEETING_ID, '60777');
  check('REPORT_SUFFIX fallback intact', cfg.REPORT_SUFFIX, '6G');
  check('template fallback intact', cfg.AGENDA_SOURCE_DOC_ID, '1qP--dusvUhNwwBtMEH4xVdxaP1c6L1hZ49geICoYV2s');
}

// ------------------------------------------------------------ 2. family unset

console.log('fresh document: Report Family is unselected, not an implicit 6G');

{
  const { html } = renderDialog();
  check('the blank placeholder option is the selected one', /<option value="" selected>/.test(html), true);
  check('no real family option is pre-selected', /<option value="[A-Za-z0-9]+" selected>/.test(html), false);
}

console.log('a configured family is shown as the real selection (legacy 6G document keeps 6G)');

{
  const { html } = renderDialog({ documentProperties: { REPORT_SUFFIX: '6G' } });
  check('6G is selected when it is actually configured', /<option value="6G" selected>6G<\/option>/.test(html), true);
  check('the blank placeholder is then not selected', /<option value="" selected>/.test(html), false);
}

// ------------------------------------------------------------ 5. wording

console.log('family label / agenda structure wording (family and agenda structure are separate concepts)');

{
  const { html } = renderDialog();
  check('Audio is labelled just "Audio"', /<option value="Audio">Audio<\/option>/.test(html), true);
  check('no option label carries an "(Agenda N.x)" suffix', /Agenda \d+\.x/.test(html), false);
  const visibleText = html.replace(/<script>[\s\S]*<\/script>/, '').replace(/<style>[\s\S]*<\/style>/, '');
  check('no developer/migration wording is visible in the dialog', /Not migrated|never auto-resolved|PROD-0|ARCH-0|ADDON-0|resolver in this task|legacy \/ main-meeting/.test(visibleText), false);
  const clientScript = html.slice(html.indexOf('<script>'));
  check('no developer wording in any client-side message string', /'[^']*(Not migrated|never auto-resolved|does NOT look this up)[^']*'/.test(clientScript), false);

  const info = JSON.parse(JSON.stringify(loadCode().sandbox.buildReportFamilyInfo_()));
  check('family info exposes the label, main section and default mailing list from the existing lookups',
    info.Audio, { label: 'Audio', section: '7.', mailingList: '3GPP_TSG_SA_WG4_AUDIO' });

  const adhoc = runClientScript(html, { meetingType: 'adhoc', reportType: 'Audio', meetingId: '85916' });
  adhoc.globals.updateDependentUi();
  check('ad-hoc + Audio: agenda structure is the discovered-agenda wording',
    adhoc.registry.agendaStructure.textContent, 'Ad-hoc — use complete discovered agenda');
  check('ad-hoc: Meeting Folder/Number block hidden', adhoc.registry.mainMeetingFields.style.display, 'none');
  check('ad-hoc: TDoc list hint asks the user to paste the URL',
    /Paste the meeting's TDoc list URL\./.test(adhoc.registry.tdocUrlHint.textContent), true);
  check('ad-hoc structure text never mentions "7.x"', /7\.x|section 7/.test(adhoc.registry.agendaStructure.textContent), false);

  const main = runClientScript(html, { meetingType: 'main', reportType: 'Audio', meetingId: '60777' });
  main.globals.updateDependentUi();
  check('main + Audio: describes the family section', main.registry.agendaStructure.textContent, 'Main meeting — Audio section 7.x');
  check('main: Meeting Folder/Number block shown', main.registry.mainMeetingFields.style.display, '');
  check('mailing-list hint names the derived default for the selected family',
    /3GPP_TSG_SA_WG4_AUDIO/.test(main.registry.mailingListHint.textContent), true);

  const unresolved = runClientScript(html, { meetingType: '', reportType: '', meetingId: '' });
  unresolved.globals.updateDependentUi();
  check('unresolved fresh state does not claim main or ad-hoc',
    unresolved.registry.agendaStructure.textContent, 'Resolve the meeting to see how the agenda is used.');
}

console.log('an existing ad-hoc document renders with the main-only fields hidden and a real summary');

{
  const { html } = renderDialog({ documentProperties: {
    MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED', MEETING_ID: '85916', MEETING_DATE: 'September 28, 2026'
  } });
  check('main-only block starts hidden for ad-hoc', /id="mainMeetingFields" style="display:none"/.test(html), true);
  check('summary shows name and Ad-hoc meeting',
    /SA4-\(AH\) Audio SWG on ULBC-MED[^<]*Ad-hoc meeting/.test(html), true);
}

console.log('an existing legacy (main) document still shows its own stored folder/number under Advanced');

{
  const { html } = renderDialog({ documentProperties: { MEETING_FOLDER: 'TSGS4_137_Xian', MEETING_NUMBER: '137', MEETING_ID: '12345' } });
  check('stored folder shown', /id="meetingFolder" value="TSGS4_137_Xian"/.test(html), true);
  check('stored number shown', /id="meetingNumber" value="137"/.test(html), true);
  check('main block visible (not hidden)', /id="mainMeetingFields" style="display:none"/.test(html), false);
}

// ------------------------------------------------------------ 6. email date

console.log('Email Collection Start Date is absent from the Meeting Configuration dialog');

{
  const { html } = renderDialog();
  check('no email start date input', /emailStartDate|Email Collection Start Date/.test(html), false);
  check('the dialog script never submits it', /emailStartDate/.test(html), false);
}

// ------------------------------------------------------------ 3. ad-hoc save

console.log('ad-hoc save does not persist fallback folder/number/portal ID');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({
    meetingFolder: '', meetingNumber: '', meetingId: '', meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG',
    reportType: 'Audio', agendaSourceDocId: '', tdocUrl: '', showPreview: true
  });
  check('no MEETING_FOLDER written', docProps._store.MEETING_FOLDER, undefined);
  check('no MEETING_NUMBER written', docProps._store.MEETING_NUMBER, undefined);
  check('no fallback MEETING_ID written', docProps._store.MEETING_ID, undefined);
  check('the user-chosen family is written', docProps._store.REPORT_SUFFIX, 'Audio');
  check('type and name are written', [docProps._store.MEETING_TYPE, docProps._store.MEETING_NAME], ['adhoc', 'SA4-(AH) Audio SWG']);
  check('the shared runtime template default is kept (not presented, but still applied)',
    docProps._store.AGENDA_SOURCE_DOC_ID, '1qP--dusvUhNwwBtMEH4xVdxaP1c6L1hZ49geICoYV2s');
}

console.log('ad-hoc save with a resolved Meeting ID persists it; existing type decides when the payload has none');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } });
  sandbox.saveConfigurationSettings({ meetingId: '85916', reportType: 'Audio', showPreview: true });
  check('MEETING_ID written when present', docProps._store.MEETING_ID, '85916');
  check('folder/number still not invented when the type comes from the saved property',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER], [undefined, undefined]);
}

console.log('a blank family never overwrites or invents one');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Video' } });
  sandbox.saveConfigurationSettings({ meetingId: '85916', meetingType: 'adhoc', reportType: '', showPreview: true });
  check('existing family survives a blank submission', docProps._store.REPORT_SUFFIX, 'Video');
  const fresh = loadCode();
  fresh.sandbox.saveConfigurationSettings({ meetingId: '85916', meetingType: 'adhoc', reportType: '', showPreview: true });
  check('a blank family on a fresh document writes nothing (no implicit 6G)', fresh.docProps._store.REPORT_SUFFIX, undefined);
}

// ------------------------------------------------------------ 4. main/legacy

console.log('main / legacy save behaviour remains compatible');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({
    meetingFolder: '', meetingNumber: '', meetingId: '', meetingType: 'main', reportType: '6G', showPreview: true
  });
  check('explicit main + blank fields keeps the historical defaults (identical to the runtime fallbacks)',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID], ['TSGS4_136_Montreal', '136', '60777']);
}

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136', REPORT_SUFFIX: '6G' } });
  sandbox.saveConfigurationSettings({
    meetingFolder: 'TSGS4_137_Xian', meetingNumber: '137', meetingId: '12345', reportType: 'Audio', tdocUrl: 'https://example.invalid/list.xlsx', showPreview: false
  });
  check('legacy document (no meeting type ever set) still updates folder/number/ID from the submitted values',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID], ['TSGS4_137_Xian', '137', '12345']);
  check('legacy save writes family, TDoc URL and preview flag as before',
    [docProps._store.REPORT_SUFFIX, docProps._store.TDOC_LIST_URL, docProps._store.SHOW_PREVIEW_SNIPPET], ['Audio', 'https://example.invalid/list.xlsx', 'false']);
  check('legacy save writes nothing centrally (no registry, no central state)',
    Object.keys(sandbox.PropertiesService.getScriptProperties()._store).filter((k) => k.indexOf('SA4_') === 0), []);
}

// ------------------------------------------------------------ 7. token

console.log('Reviewer API token is never rendered');

{
  const { html } = renderDialog({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  check('the saved token value is not in the rendered HTML', html.indexOf(REAL_LOOKING_TOKEN) === -1, true);
  check('a configured-status line is shown instead', /Reviewer API token configured/.test(html), true);
  check('the replacement input starts empty', /id="apiToken" value=""/.test(html), true);
}

{
  const { html } = renderDialog();
  check('no token: status says none is configured', /No Reviewer API token configured/.test(html), true);
}

console.log('Reviewer API token is never logged');

{
  const { sandbox } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  const logs = captureLogs(sandbox);
  sandbox.saveConfigurationSettings({ meetingId: '85916', meetingType: 'adhoc', reportType: 'Audio', apiTokenAction: 'replace', apiToken: NEW_TOKEN, showPreview: true });
  const all = logs.join('\n');
  check('the new token is not in any log line', all.indexOf(NEW_TOKEN) === -1, true);
  check('the previously saved token is not in any log line', all.indexOf(REAL_LOOKING_TOKEN) === -1, true);
  check('a save log line still exists', /Configuration saved/.test(all), true);
  check('the token field is marked redacted', /\[redacted\]/.test(all), true);
}

{
  const { sandbox } = loadCode();
  const redacted = sandbox.redactConfigForLog_({ apiToken: NEW_TOKEN, someSecret: 'x', reviewerPassword: 'y', meetingId: '1', apiTokenAction: 'replace' });
  check('every secret-like key is redacted, ordinary keys and the action are not',
    [redacted.apiToken, redacted.someSecret, redacted.reviewerPassword, redacted.meetingId, redacted.apiTokenAction], ['[redacted]', '[redacted]', '[redacted]', '1', 'replace']);
}

console.log('token keep / replace / clear semantics');

{
  const keep = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  keep.sandbox.saveConfigurationSettings({ apiTokenAction: 'keep', apiToken: '', showPreview: true });
  check('keep leaves the saved token', keep.scriptProps._store.REVIEWER_API_TOKEN, REAL_LOOKING_TOKEN);

  const blank = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  blank.sandbox.saveConfigurationSettings({ apiToken: '', showPreview: true });
  check('a legacy-style blank submission still keeps the token', blank.scriptProps._store.REVIEWER_API_TOKEN, REAL_LOOKING_TOKEN);

  const replace = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  replace.sandbox.saveConfigurationSettings({ apiTokenAction: 'replace', apiToken: NEW_TOKEN, showPreview: true });
  check('replace stores the new token', replace.scriptProps._store.REVIEWER_API_TOKEN, NEW_TOKEN);

  const legacy = loadCode();
  legacy.sandbox.saveConfigurationSettings({ apiToken: NEW_TOKEN, showPreview: true });
  check('a caller sending only apiToken (no action) still sets it', legacy.scriptProps._store.REVIEWER_API_TOKEN, NEW_TOKEN);

  const emptyReplace = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  emptyReplace.sandbox.saveConfigurationSettings({ apiTokenAction: 'replace', apiToken: '   ', showPreview: true });
  check('replace with a blank value does not wipe the token', emptyReplace.scriptProps._store.REVIEWER_API_TOKEN, REAL_LOOKING_TOKEN);

  const clear = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  clear.sandbox.saveConfigurationSettings({ apiTokenAction: 'clear', apiToken: '', showPreview: true });
  check('clear removes the token', clear.scriptProps._store.REVIEWER_API_TOKEN, undefined);
  check('after clearing, getReportConfig_ reports no token', clear.sandbox.getReportConfig_().REVIEWER_API_TOKEN, '');
}

{
  const { html } = renderDialog();
  const s = runClientScript(html, {});
  // drive saveConfig() and inspect the payload sent to the server
  let sent = null;
  s.globals.google.script.run = {
    withSuccessHandler: () => ({ withFailureHandler: () => ({ saveConfigurationSettings: (c) => { sent = c; } }) })
  };
  s.registry.apiToken.value = NEW_TOKEN;
  s.globals.saveConfig();
  check('client sends replace + the new token when one is typed', [sent.apiTokenAction, sent.apiToken], ['replace', NEW_TOKEN]);
  s.registry.clearApiToken.checked = true;
  s.globals.saveConfig();
  check('client sends clear (and no token) when the clear box is ticked', [sent.apiTokenAction, sent.apiToken], ['clear', '']);
  s.registry.clearApiToken.checked = false;
  s.registry.apiToken.value = '';
  s.globals.saveConfig();
  check('client sends keep when nothing was changed', sent.apiTokenAction, 'keep');
  check('client payload has no email start date', 'emailStartDate' in sent, false);
}

// ------------------------------------------------------------ 9. registered doc

console.log('registered/adopted document: Save immediately updates the authoritative central state');

const DOC = 'DOC_B1_REGISTERED';
const INTERACTIVE = { documentId: DOC, mode: 'addon-interactive' };
const BACKGROUND = { documentId: DOC, mode: 'addon-background' };

{
  const { sandbox, docProps, scriptProps } = loadCode({ documentProperties: {
    MEETING_TYPE: 'adhoc', MEETING_NAME: 'Old Name', MEETING_ID: '85916', REPORT_SUFFIX: 'Video',
    TDOC_LIST_URL: 'https://example.invalid/old-list.xlsx', MAILING_LIST: 'OLD_LIST', SHOW_PREVIEW_SNIPPET: 'true',
    REVISION_MAP: '{"a":"b"}'
  } });
  withActiveDocument(sandbox, DOC);
  sandbox.adoptReportDocumentForAddon_(INTERACTIVE);
  sandbox.registerReportDocument_(DOC, { enabled: true, intervalHours: 2 });
  const central = sandbox.getReportStateStore_(BACKGROUND);
  central.setProperty('DISCUSS_S4aA260090', '{"cache":1}');

  check('precondition: adopted copy has the OLD family', central.getProperty('REPORT_SUFFIX'), 'Video');
  check('precondition: background config sees the old TDoc list', sandbox.getReportConfig_(BACKGROUND).TDOC_LIST_URL, 'https://example.invalid/old-list.xlsx');

  sandbox.saveConfigurationSettings({
    meetingId: '85916', meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG on ULBC-MED', meetingDate: 'September 28, 2026',
    ftpBase: 'https://example.invalid/Docs/', agendaTdoc: 'S4aA260090', reportType: 'Audio',
    tdocUrl: 'https://example.invalid/new-list.xlsx', mailingList: 'NEW_LIST', showPreview: false, apiTokenAction: 'keep'
  });

  check('addon-interactive config sees the new family (no re-Enable)', sandbox.getReportConfig_(INTERACTIVE).REPORT_SUFFIX, 'Audio');
  check('addon-background config sees the new family', sandbox.getReportConfig_(BACKGROUND).REPORT_SUFFIX, 'Audio');
  check('background config sees the new TDoc list URL', sandbox.getReportConfig_(BACKGROUND).TDOC_LIST_URL, 'https://example.invalid/new-list.xlsx');
  check('background config sees the new agenda TDoc', sandbox.getReportConfig_(BACKGROUND).AGENDA_TDOC, 'S4aA260090');
  check('background config sees the preview flag change', sandbox.getReportConfig_(BACKGROUND).SHOW_PREVIEW_SNIPPET, false);
  check('background meeting identity sees the new mailing list override',
    sandbox.getMeetingIdentityConfig_(BACKGROUND).MAILING_LIST, 'NEW_LIST');
  check('background meeting identity sees the new name', sandbox.getMeetingIdentityConfig_(BACKGROUND).MEETING_NAME, 'SA4-(AH) Audio SWG on ULBC-MED');

  check('Document Properties (legacy source) were also saved and retained', [docProps._store.REPORT_SUFFIX, docProps._store.MAILING_LIST], ['Audio', 'NEW_LIST']);
  check('unrelated central state (cache) is untouched', central.getProperty('DISCUSS_S4aA260090'), '{"cache":1}');
  check('unrelated central state (adopted REVISION_MAP) is untouched', central.getProperty('REVISION_MAP'), '{"a":"b"}');

  const entry = sandbox.getRegisteredReportDocument_(DOC);
  check('registration flags are preserved (still enabled, same interval)', [entry.enabled, entry.intervalHours], [true, 2]);
  check('registry shows the real meeting name, not a stale folder', entry.meetingName, 'SA4-(AH) Audio SWG on ULBC-MED');
  check('central state keeps its own namespace (keys live under SA4_STATE|<documentId>|)',
    Object.keys(scriptProps._store).some((k) => k === 'SA4_STATE|' + DOC + '|REPORT_SUFFIX'), true);

  // clearing an override must clear it centrally too
  sandbox.saveConfigurationSettings({ meetingId: '85916', meetingType: 'adhoc', reportType: 'Audio', tdocUrl: '', showPreview: true });
  check('clearing the TDoc list in the dialog removes it from central state as well', central.getProperty('TDOC_LIST_URL'), null);
  check('and the ad-hoc identity no longer exposes it to background runs', sandbox.getMeetingIdentityConfig_(BACKGROUND).TDOC_LIST_URL, '');
}

console.log('registered document: a lock failure leaves BOTH copies untouched');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Video' } });
  withActiveDocument(sandbox, DOC);
  sandbox.adoptReportDocumentForAddon_(INTERACTIVE);
  sandbox.LockService = { getScriptLock: () => ({ tryLock: () => false, releaseLock: () => {} }) };

  let thrown = null;
  try { sandbox.saveConfigurationSettings({ meetingType: 'adhoc', reportType: 'Audio', showPreview: true }); } catch (e) { thrown = e; }
  check('Save reports the contention', thrown !== null && /already in progress/.test(thrown.message), true);
  check('Document Properties unchanged', docProps._store.REPORT_SUFFIX, 'Video');
  check('central state unchanged', sandbox.getReportStateStore_(BACKGROUND).getProperty('REPORT_SUFFIX'), 'Video');
}

// ------------------------------------------------------------ 10. unregistered

console.log('unregistered / legacy document: Save uses Document Properties only');

{
  const { sandbox, docProps, scriptProps } = loadCode({ documentProperties: { REPORT_SUFFIX: 'Video' } });
  withActiveDocument(sandbox, 'DOC_NEVER_ADOPTED');
  sandbox.saveConfigurationSettings({ meetingId: '12345', meetingType: 'main', reportType: 'Audio', showPreview: true });
  check('Document Properties updated', docProps._store.REPORT_SUFFIX, 'Audio');
  check('no central state was created', Object.keys(scriptProps._store).filter((k) => k.indexOf('SA4_') === 0), []);
  check('no registry entry was created', sandbox.getRegisteredReportDocument_('DOC_NEVER_ADOPTED'), null);
  check('an interactive-context read of an unregistered document still uses Document Properties',
    sandbox.getReportConfig_({ documentId: 'DOC_NEVER_ADOPTED', mode: 'addon-interactive' }).REPORT_SUFFIX, 'Audio');
}

console.log('no active document (bound/legacy sandbox): Save still works unchanged');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingId: '12345', reportType: 'Audio', showPreview: true });
  check('saved without an active document', docProps._store.REPORT_SUFFIX, 'Audio');
}

// ------------------------------------------------------------ structure

console.log('structural guarantees');

{
  const { sandbox } = loadCode();
  const adoption = sandbox.ADDON003_ADOPTION_FIXED_KEYS_;
  check('every dialog-managed key is also an adoption key (sync never invents a new central key)',
    sandbox.CONFIG_DIALOG_MANAGED_KEYS_.filter((k) => adoption.indexOf(k) === -1), []);

  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const start = source.indexOf('function persistConfigurationSettings_(');
  const body = source.slice(start, source.indexOf('function syncConfigurationDialogStateToCentral_('));
  check('the save path no longer logs the raw config object', /JSON\.stringify\(config\)/.test(body), false);
  check('getReportConfig_ legacy fallbacks are still present in source',
    /props\.getProperty\('MEETING_FOLDER'\) \|\| 'TSGS4_136_Montreal'/.test(source) && /props\.getProperty\('REPORT_SUFFIX'\) \|\| '6G'/.test(source), true);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
