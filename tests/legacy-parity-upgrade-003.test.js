/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/legacy-upgrade-003.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-upgrade-003.test.js
 */

/**
 * LEGACY-UPGRADE-003 -- report-family inference + derived mailing list,
 * ported from the central add-on's ADDON-007B2, adapted to this legacy
 * bound script's own (pre-B1) Configure Meeting dialog.
 *
 * Covers: generic table-driven inference (confident/suggested/unresolved/
 * conflict), explicit/saved family always wins over a suggestion, derived
 * vs. explicit-override mailing list with a reset path, no accidental
 * Document Property writes from merely opening the dialog or from Resolve/
 * Discover, and regression coverage for EMAIL_START_DATE,
 * LEGACY-UPGRADE-002's token handling and stale-default fix, and
 * BUGFIX-LEGACY-001.
 *
 * This harness is isolated -- it lives only in this git checkout, is not
 * referenced by appsscript.json, and is excluded from the Apps Script push
 * payload by .claspignore.
 *
 * Run: node tests/legacy-upgrade-003.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode, CODE_JS_PATH } = require('./helpers/legacy-load-code.js');

// TEMPLATE-002A harness adaptation: Legacy's historical commits exist only in
// the sa4-report-legacy repository. Blocks that load one (to show Legacy's own
// before/after, or that a Legacy stage touched only certain functions) are
// skipped here; they say nothing about the current Code.js.
const LEGACY_GIT_HISTORY = false;
function skipLegacyHistory(what) { console.log('  skip ' + what + ' [needs Legacy git history]'); }

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

{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG on ULBC-MED' });
  check('name only Audio -> suggested', [r.family, r.confidence, r.sources], ['Audio', 'suggested', ['meeting-name']]);
}
{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc', tdocFamilyKeys: { 'video-adhoc': 3 } });
  check('S4aV documents only -> suggested Video', [r.family, r.confidence, r.sources, r.tdocPrefixes], ['Video', 'suggested', ['tdoc-family'], ['S4aV']]);
}
{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG', tdocFamilyKeys: { 'audio-adhoc': 2 } });
  check('name + documents agree (Audio) -> confident', [r.family, r.confidence, r.sources], ['Audio', 'confident', ['meeting-name', 'tdoc-family']]);
}
{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-e (AH) on FS_6G_MED' });
  check('FS_6G_MED title -> suggested 6G', [r.family, r.confidence], ['6G', 'suggested']);
}
{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc', tdocFamilyKeys: { 'plenary-adhoc': 5 } });
  check('S4aP alone -> unresolved (never 6G/Liaison/New)', [r.family, r.confidence, r.reason], [null, 'unresolved', 'ambiguous-tdoc-series']);
}
{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4 ad hoc' });
  check('no evidence at all -> unresolved', [r.family, r.confidence, r.reason], [null, 'unresolved', 'no-evidence']);
}
{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG', tdocFamilyKeys: { 'video-adhoc': 4 } });
  check('name Audio + S4aV documents -> conflict, nothing chosen', [r.family, r.confidence], [null, 'conflict']);
  const flipped = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Video SWG', tdocFamilyKeys: { 'audio-adhoc': 4 } });
  check('conflict does not depend on which source is processed first', [flipped.family, flipped.confidence], [null, 'conflict']);
}
{
  const r = infer({ meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio and Video SWG' });
  check('multiple incompatible name tokens -> conflict', [r.family, r.confidence], [null, 'conflict']);
}
{
  const r = infer({ meetingType: 'main', meetingName: 'SA4#137 Audio and Video', tdocFamilyKeys: { 'audio-adhoc': 1 } });
  check('main meeting: never inferred, regardless of evidence', [r.applicable, r.family, r.confidence], [false, null, 'none']);
}
['Audiobook workshop', 'Rtcweb', 'Newsletter Liaison New items'].forEach((name) => {
  check(`no substring false positive: "${name}"`, infer({ meetingType: 'adhoc', meetingName: name }).family, null);
});

console.log('supported families come from the existing registry, not a second universe');

{
  const registry = JSON.parse(vm.runInContext('JSON.stringify(SA4_TDOC_FAMILIES)', pure));
  check('canonical registry maps the four SWG series and leaves S4aP/S4- unmapped',
    registry.filter((f) => f.reportFamily).map((f) => [f.prefix, f.reportFamily]),
    [['S4aA', 'Audio'], ['S4aV', 'Video'], ['S4aI', 'MBS'], ['A4aR', 'RTC']]);
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const body = src.slice(src.indexOf('function inferReportFamily_('), src.indexOf('function buildReportFamilyInfo_('));
  check('no meeting-specific (85916/86178/ULBC) constant in the inference code', /85916|86178|ULBC/i.test(body), false);
}

// ------------------------------------------------------------ dialog client

const IDS = ['meetingId', 'meetingName', 'meetingType', 'meetingDate', 'ftpBase', 'agendaTdoc', 'mailingList', 'mailingListHint', 'mailingListReset',
  'revisionsUrl', 'meetingFolder', 'meetingNumber', 'reportType', 'familyStatus', 'agendaSourceDocId', 'tdocUrl', 'showPreview', 'emailStartDate',
  'apiToken', 'clearApiToken', 'resolveStatus', 'resolveBtn', 'discoverStatus', 'discoverBtn', 'portalTypeHint', 'dateRangeHint', 'agendaCandidates',
  'agendaTdocBadge', 'revisionsUrlBadge', 'readinessBox',
  // LEGACY-UPGRADE-006H addition (restored field -- the dialog script reads this unconditionally in saveConfig()).
  'discussionEmailSender'];

// Renders the real dialog and runs its real client <script> with
// google.script.run wired to the REAL server functions of the same
// sandbox (network stubbed with fixtures).
function openDialog(documentProperties, meetingFixture, tdocRows) {
  const loaded = loadCode({ documentProperties: documentProperties || {} });
  const sandbox = loaded.sandbox;
  let html = null;
  sandbox.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.configureMeetingSettings();

  if (meetingFixture) sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(meetingFixture) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  sandbox.parseMeetingTdocListHtml_ = () => tdocRows || [];

  const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')).replace(/\\\\/g, '\\');
  const registry = {};
  IDS.forEach((id) => { registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }; });
  const value = (id) => { const m = html.match(new RegExp('id="' + id + '" value="([^"]*)"')); return m ? m[1].replace(/&quot;/g, '"') : ''; };
  // TEMPLATE-002A harness adaptation: the CENTRAL dialog carries the
  // per-family defaults in a hidden "familyInfo" input (Legacy inlined them
  // in the script), so the fake DOM has to hand that input's value over too.
  registry.familyInfo = { id: 'familyInfo', value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' };
  ['familyInfo', 'meetingId', 'meetingType', 'meetingName', 'mailingList', 'ftpBase', 'agendaTdoc', 'meetingFolder', 'meetingNumber'].forEach((id) => { registry[id].value = value(id); });
  const selected = html.match(/<option value="([A-Za-z0-9]+)" selected>/);
  registry.reportType.value = selected ? selected[1] : '';

  let sent = null;
  const alerts = [];
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
    // TEMPLATE-002A harness adaptation: Legacy's fake DOM listed the element
    // ids of Legacy's dialog. The CENTRAL dialog has more elements, so an
    // unknown id gets an element on demand instead of null.
    document: { getElementById: (id) => registry[id] || (registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }) },
    google: { script: { get run() { return runner(); }, host: { close: () => {} } } },
    alert: (m) => alerts.push(m),
    console: { log: () => {} }
  };
  vm.createContext(globals);
  vm.runInContext(script, globals);
  return { registry, globals, sandbox, alerts, html, sentConfig: () => sent, loaded };
}
const pickFamily = (d, f) => { d.registry.reportType.value = f; d.globals.onFamilyChanged(); };
const typeMailing = (d, text) => { d.registry.mailingList.value = text; d.globals.onMailingListEdited(); };

const M85916 = FIXTURES.getMeetings85916;
const AUDIO_ROWS = [{ id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '' }, { id: 'S4aA260091', type: 'pCR', revisionOf: null, agendaItem: '4' }];

console.log('client: Resolve on an Audio ad-hoc proposes Audio and derives its mailing list');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  check('precondition: fresh dialog has no family and no mailing list', [d.registry.reportType.value, d.registry.mailingList.value], ['', '']);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Report Family becomes Audio', d.registry.reportType.value, 'Audio');
  check('mailing list shows the derived Audio default', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_AUDIO');
  check('the reset link is not shown for a derived (non-overridden) value', d.registry.mailingListReset.style.display, 'none');

  d.globals.discoverAgendaTdocs();
  check('Discover with S4aA documents confirms Audio (still selected)', d.registry.reportType.value, 'Audio');
}

console.log('client: an explicit/saved family always wins, never silently overwritten');

{
  const d = openDialog({ REPORT_SUFFIX: 'Video', MEETING_TYPE: 'adhoc' }, M85916, AUDIO_ROWS);
  check('a saved family is shown on open', d.registry.reportType.value, 'Video');
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Resolve does not replace the saved family (Audio-named meeting, saved Video)', d.registry.reportType.value, 'Video');
  d.globals.discoverAgendaTdocs();
  check('Discover (S4aA documents) does not replace it either', d.registry.reportType.value, 'Video');
}

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('an automatic suggestion is in place first', d.registry.reportType.value, 'Audio');
  pickFamily(d, 'RTC');
  check('a manual choice overrides the suggestion', d.registry.reportType.value, 'RTC');
  d.globals.discoverAgendaTdocs();
  check('a later Discover never overwrites the manual choice', d.registry.reportType.value, 'RTC');
}

console.log('client: conflicting evidence withdraws a suggestion, never silently switches');

{
  const VIDEO_ROWS = [{ id: 'S4aV260010', type: 'agenda', revisionOf: null, agendaItem: '' }];
  const d = openDialog({}, M85916, VIDEO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Resolve first suggests Audio (from the meeting name)', d.registry.reportType.value, 'Audio');
  d.globals.discoverAgendaTdocs();
  check('Discover finds S4aV documents: the suggestion is withdrawn, not switched', d.registry.reportType.value, '');
  check('the derived mailing list is withdrawn with it', d.registry.mailingList.value, '');
}

// ------------------------------------------------------------ mailing list

console.log('mailing list: derived default follows the family (client)');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  pickFamily(d, 'Audio');
  check('Audio -> derived Audio list', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_AUDIO');
  pickFamily(d, 'Video');
  check('derived list follows a family change', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_VIDEO');
}

console.log('mailing list: an explicit override survives a family change, and can be reset');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  pickFamily(d, 'Audio');
  typeMailing(d, 'MY_OWN_LIST');
  check('an edit marks the value as overridden', /Custom value/.test(d.registry.mailingListHint.textContent), true);
  check('the reset link appears', d.registry.mailingListReset.style.display, '');
  pickFamily(d, 'MBS');
  check('the explicit override survives a family change', d.registry.mailingList.value, 'MY_OWN_LIST');
  check('the hint names the default it overrides', /3GPP_TSG_SA_WG4_MBS/.test(d.registry.mailingListHint.textContent), true);

  d.globals.resetMailingList();
  check('reset returns to the family-derived default', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_MBS');
  check('the reset link is hidden again', d.registry.mailingListReset.style.display, 'none');
  pickFamily(d, 'RTC');
  check('derivation resumes on the next family change after a reset', d.registry.mailingList.value, '3GPP_TSG_SA_WG4_RTC');
}

console.log('mailing list: an existing saved override is shown and preserved across Resolve');

{
  const d = openDialog({ MAILING_LIST: 'LEGACY_LIST', REPORT_SUFFIX: 'Audio', MEETING_TYPE: 'adhoc' }, M85916, AUDIO_ROWS);
  check('saved list shown on open', d.registry.mailingList.value, 'LEGACY_LIST');
  check('marked as custom on open', /Custom value/.test(d.registry.mailingListHint.textContent), true);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Resolve keeps the saved override', d.registry.mailingList.value, 'LEGACY_LIST');
}

// ------------------------------------------------------------ save semantics

console.log('save: derived vs explicit override');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  pickFamily(d, 'Audio');
  d.globals.saveConfig();
  check('a derived value is sent as mode "derived" with no value', [d.sentConfig().mailingListMode, d.sentConfig().mailingList], ['derived', '']);

  typeMailing(d, 'MY_OWN_LIST');
  d.globals.saveConfig();
  check('an explicit edit is sent as an override', [d.sentConfig().mailingListMode, d.sentConfig().mailingList], ['override', 'MY_OWN_LIST']);

  d.globals.resetMailingList();
  d.globals.saveConfig();
  check('reset is sent as derived again', [d.sentConfig().mailingListMode, d.sentConfig().mailingList], ['derived', '']);
}

console.log('save (server): derived is never persisted as an override; explicit override is persisted; reset deletes it');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: '', mailingListMode: 'derived', showPreview: true });
  check('family persisted', docProps._store.REPORT_SUFFIX, 'Audio');
  check('a derived mailing list is NOT persisted as an override', docProps._store.MAILING_LIST, undefined);

  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: ' MY_LIST ', mailingListMode: 'override', showPreview: true });
  check('an explicit override is persisted (trimmed)', docProps._store.MAILING_LIST, 'MY_LIST');

  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: '', mailingListMode: 'derived', showPreview: true });
  check('reset deletes the explicit MAILING_LIST so derivation resumes', docProps._store.MAILING_LIST, undefined);
  check('the property is truly absent, not merely blank', Object.prototype.hasOwnProperty.call(docProps._store, 'MAILING_LIST'), false);
}

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MAILING_LIST: 'LEGACY_LIST' } });
  sandbox.saveConfigurationSettings({ meetingId: '1', reportType: 'Audio', mailingList: '', showPreview: true });
  check('legacy caller (no mode) with a blank list keeps the old skip-if-blank behaviour', docProps._store.MAILING_LIST, 'LEGACY_LIST');
}

// ------------------------------------------------------------ no accidental persistence

console.log('no accidental REPORT_SUFFIX/MAILING_LIST persistence merely from opening/Resolve/Discover');

{
  const before = { REPORT_SUFFIX: 'Audio', MAILING_LIST: '3GPP_TSG_SA_WG4_AUDIO', MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-(AH) Audio SWG on ULBC-MED' };
  const d = openDialog(before, M85916, AUDIO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  d.globals.discoverAgendaTdocs();
  check('dialog open + Resolve + Discover leave every legacy property untouched (even one equal to the derived value)',
    JSON.parse(JSON.stringify(d.loaded.docProps._store)), before);
}

{
  const before = {};
  const d = openDialog(before, M85916, AUDIO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  d.globals.discoverAgendaTdocs();
  check('on a fresh document, open + Resolve + Discover write nothing at all (no implicit REPORT_SUFFIX/MAILING_LIST)',
    Object.keys(d.loaded.docProps._store), []);
}

// ------------------------------------------------------------ regression

console.log('EMAIL_START_DATE: no dialog field in the template (intentional difference)');

// INTENTIONAL DIFFERENCE (CENTRAL ADDON-007B1, commit 0510461): the
// Meeting Configuration dialog has no Email Collection Start Date field
// and Save does not write EMAIL_START_DATE. Legacy keeps the field. The
// collector honours a stored EMAIL_START_DATE either way (BUGFIX-LEGACY-003,
// tests/legacy-parity-adhoc-email-collection.test.js).
{
  const d2 = openDialog({ EMAIL_START_DATE: '2026-09-01' }, null, []);
  check('[template] the dialog has no Email Collection Start Date field', /Email Collection Start Date|id="emailStartDate"/.test(d2.html), false);
  const { sandbox, docProps } = loadCode({ documentProperties: { EMAIL_START_DATE: '2026-09-01' } });
  sandbox.saveConfigurationSettings({ emailStartDate: '2026-09-20', showPreview: true });
  check('[template] Save neither writes nor erases EMAIL_START_DATE', docProps._store.EMAIL_START_DATE, '2026-09-01');
}

console.log('regression: LEGACY-UPGRADE-002 token handling unchanged');

{
  const REAL_TOKEN = 'crv1_TESTONLY_not_a_real_token_0123456789';
  const d = openDialog({}, null, []);
  check('no token: status says none configured', /No Reviewer API token configured/.test(d.html), true);
  const withToken = openDialog({ scriptProperties: { REVIEWER_API_TOKEN: REAL_TOKEN } }, null, []);
  check('a configured token never appears in the rendered HTML', withToken.html.indexOf(REAL_TOKEN) === -1, true);

  const { sandbox, scriptProps } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_TOKEN } });
  sandbox.saveConfigurationSettings({ apiTokenAction: 'keep', apiToken: '', showPreview: true });
  check('keep still leaves the token untouched', scriptProps._store.REVIEWER_API_TOKEN, REAL_TOKEN);
  sandbox.saveConfigurationSettings({ apiTokenAction: 'clear', apiToken: '', showPreview: true });
  check('clear still removes it', scriptProps._store.REVIEWER_API_TOKEN, undefined);
}

console.log('regression: LEGACY-UPGRADE-002 stale-default fix unchanged');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingFolder: '', meetingNumber: '', meetingId: '', meetingType: 'adhoc', reportType: '', showPreview: true });
  check('ad-hoc blank folder/number/ID/family still invents nothing',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID, docProps._store.REPORT_SUFFIX],
    [undefined, undefined, undefined, undefined]);

  const main = loadCode();
  main.sandbox.saveConfigurationSettings({ meetingFolder: '', meetingNumber: '', meetingId: '', meetingType: 'main', reportType: '', showPreview: true });
  check('explicit main still keeps the historical folder/number/ID defaults',
    [main.docProps._store.MEETING_FOLDER, main.docProps._store.MEETING_NUMBER, main.docProps._store.MEETING_ID],
    ['TSGS4_136_Montreal', '136', '60777']);
}

console.log('regression: BUGFIX-LEGACY-001 unchanged');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'main', MEETING_FOLDER: 'F', MEETING_NUMBER: '1' } });
  // LEGACY-0099 (as ADDON-008A1b): the build's first document access is now
  // reading the saved Document Reallocations (before anything is cleared),
  // so reaching the document means reaching either that read or clear().
  const reachDocument = () => { throw new Error('REACHED_BODY_CLEAR'); };
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ clear: reachDocument, getTables: reachDocument }) });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('main meetings remain unblocked', thrown && thrown.message, 'REACHED_BODY_CLEAR');
}
{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } });
  const calls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => { calls.push('getBody'); return { clear: () => calls.push('clear') }; } });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('incomplete ad-hoc build still fails before touching the document', [thrown !== null, calls], [true, []]);
  check('error message unchanged', thrown && /Cannot build report yet/.test(thrown.message), true);
}

// ------------------------------------------------------------ structural

console.log('structural: no report/build/content mutation logic touched');

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('LEGACY-UPGRADE-003 stage scope (6 checks)'); else {
  const { execFileSync } = require('child_process');
  const POST_002_COMMIT = '9739085';
  // LEGACY-UPGRADE-003's own commit (not the live/current file) -- this
  // keeps "exactly these new functions" a fact about the 003 stage itself,
  // so it stays valid once LEGACY-UPGRADE-004+ add their own functions on
  // top (same pattern LEGACY-UPGRADE-002's own structural check uses).
  const POST_003_COMMIT = 'b97b46c';
  const REPO_ROOT = path.join(__dirname, '..');
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const post002Source = execFileSync('git', ['show', `${POST_002_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  const post003Source = execFileSync('git', ['show', `${POST_003_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  const names = (s) => new Set((s.match(/^function ([A-Za-z0-9_]+)\(/gm) || []).map((m) => m.replace(/^function |\($/g, '')));
  const before002 = names(post002Source);
  const after = names(src);
  const removed = [...before002].filter((n) => !after.has(n));
  check('no function present as of LEGACY-UPGRADE-002 has since been removed', removed, []);
  const addedByThisStage = [...names(post003Source)].filter((n) => !before002.has(n));
  check('LEGACY-UPGRADE-003 itself added exactly its own new functions (family inference + info)', addedByThisStage.sort(),
    ['buildReportFamilyInfo_', 'inferReportFamily_', 'reportFamiliesNamedIn_', 'tokenizeFamilyEvidenceText_']);

  check('assertMeetingReadyToBuild_ untouched (still present, unmodified signature)', /function assertMeetingReadyToBuild_\(\) \{/.test(src), true);
  check('buildSkeletonWithTdocTables still calls the guard first', /function buildSkeletonWithTdocTables\(\) \{\s*\/\/ BUGFIX-LEGACY-001/.test(src), true);
  check('no ADDON-007A admin-anchor functions were present as of LEGACY-UPGRADE-003 itself (that is LEGACY-UPGRADE-004\'s scope)',
    /getAdministrativeAgendaAnchors_|deriveAdminAnchorsFromBuiltDocument_/.test(post003Source), false);
  check('no central registry/adoption/scheduler infrastructure introduced',
    /adoptReportDocumentForAddon_|getReportStateStore_|runAddonScheduler_/.test(src), false);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
