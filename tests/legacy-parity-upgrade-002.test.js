/**
 * TEMPLATE-002A LEGACY PARITY GATE -- imported from sa4-report-legacy
 * tests/legacy-upgrade-002.test.js at Legacy master b95e06e (deployed source 4fe295a).
 * It runs against THIS repository's Code.js. Changes to the Legacy file are
 * limited to, and marked as:
 *   "TEMPLATE-002A harness adaptation"  test mechanics only
 *   "INTENTIONAL DIFFERENCE" / "[template]"  the CENTRAL/template behaviour
 *                                       that deliberately differs from Legacy
 * See docs/TEMPLATE-002A_LEGACY_PARITY.md. The original header follows.
 *
 * Run: node tests/legacy-parity-upgrade-002.test.js
 */

/**
 * LEGACY-UPGRADE-002 -- configuration/security cleanup for the legacy
 * bound Apps Script project (REDACTED-LIVE-6G-SCRIPT-ID).
 *
 * Covers: the Reviewer API token is never rendered into dialog HTML and
 * never logged unredacted, keep/replace/clear token semantics, ad-hoc
 * stale-default-write correction (folder/number/ID/family), main-meeting
 * compatibility, EMAIL_START_DATE preserved exactly, and the
 * BUGFIX-LEGACY-001 build guard unchanged.
 *
 * This harness is isolated -- it lives only in this git checkout, is not
 * referenced by appsscript.json, and is never part of what clasp pushes.
 *
 * Where a test demonstrates a behavior that is INTENTIONALLY different
 * from before this stage, it also runs the same scenario against the git
 * baseline commit (the fresh pull right after BUGFIX-LEGACY-001) so the
 * "before" and "after" are both visible, not just asserted.
 *
 * Run: node tests/legacy-upgrade-002.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/legacy-load-code.js');

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

const REPO_ROOT = path.join(__dirname, '..');
const BASELINE_COMMIT = '248e1c3';
const REAL_LOOKING_TOKEN = 'crv1_TESTONLY_not_a_real_token_0123456789';
const NEW_TOKEN = 'crv1_TESTONLY_replacement_token_9876543210';

// TEMPLATE-002A harness adaptation: Legacy's historical commits exist only in
// the sa4-report-legacy repository. Blocks that load one (to show Legacy's own
// before/after, or that a Legacy stage touched only certain functions) are
// skipped here; they say nothing about the current Code.js.
const LEGACY_GIT_HISTORY = false;
function skipLegacyHistory(what) { console.log('  skip ' + what + ' [needs Legacy git history]'); }

// ---- loads the BASELINE (pre-LEGACY-UPGRADE-002) Code.js from git, into
// ---- the exact same kind of sandbox load-code.js builds. Used only to
// ---- characterize what is intentionally changing.
function loadBaselineCode(options) {
  const opts = options || {};
  const source = execFileSync('git', ['show', `${BASELINE_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  function store(initial) {
    const s = Object.assign({}, initial || {});
    return { getProperty: (k) => (k in s ? s[k] : null), setProperty: (k, v) => { s[k] = v; }, deleteProperty: (k) => { delete s[k]; }, _store: s };
  }
  const docProps = store(opts.documentProperties);
  const scriptProps = store(opts.scriptProperties);
  const sandbox = {
    Logger: { log: () => {} },
    PropertiesService: { getDocumentProperties: () => docProps, getScriptProperties: () => scriptProps },
    DocumentApp: { getActiveDocument: () => { throw new Error('n/a'); }, getUi: () => { throw new Error('n/a'); }, openById: () => { throw new Error('n/a'); }, ParagraphHeading: {}, ElementType: {} },
    Session: { getScriptTimeZone: () => 'Europe/Berlin' },
    Utilities: {}, UrlFetchApp: { fetch: () => { throw new Error('n/a'); } }, DriveApp: {}, SpreadsheetApp: {}, HtmlService: {}, ScriptApp: {}, MimeType: {}, XmlService: {},
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }), getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    console
  };
  vm.createContext(sandbox);
  vm.runInContext(source, sandbox, { filename: 'Code.js@' + BASELINE_COMMIT });
  return { sandbox, docProps, scriptProps };
}

// Renders the real dialog by capturing what configureMeetingSettings() hands
// to HtmlService.
function renderDialog(loadFn, options) {
  const loaded = loadFn(options);
  let captured = null;
  loaded.sandbox.HtmlService = { createHtmlOutput: (html) => { captured = html; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  loaded.sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  loaded.sandbox.configureMeetingSettings();
  return Object.assign({ html: captured }, loaded);
}

// Runs the dialog's real client <script> against a minimal fake DOM.
function extractClientScript(html) {
  const start = html.indexOf('<script>') + '<script>'.length;
  const end = html.lastIndexOf('</script>');
  return html.slice(start, end).replace(/\\\\/g, '\\');
}

function runClientScript(html, values) {
  const scriptText = extractClientScript(html);
  const ids = ['meetingId', 'meetingName', 'meetingType', 'meetingDate', 'ftpBase', 'agendaTdoc', 'mailingList', 'revisionsUrl',
    'meetingFolder', 'meetingNumber', 'reportType', 'agendaSourceDocId', 'tdocUrl', 'showPreview', 'emailStartDate',
    'apiToken', 'clearApiToken', 'resolveStatus', 'resolveBtn', 'discoverStatus', 'discoverBtn', 'portalTypeHint',
    'dateRangeHint', 'agendaCandidates',
    // LEGACY-UPGRADE-003 additions (dialog script now references these unconditionally on load).
    'mailingListHint', 'mailingListReset', 'familyStatus',
    // LEGACY-UPGRADE-006H addition (restored field -- the dialog script reads this unconditionally in saveConfig()).
    'discussionEmailSender'];
  const registry = {};
  ids.forEach((id) => { registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }; });
  Object.keys(values || {}).forEach((k) => {
    if (typeof values[k] === 'boolean') registry[k].checked = values[k];
    else registry[k].value = values[k];
  });
  let sent = null;
  const globals = {
    // TEMPLATE-002A harness adaptation: Legacy's fake DOM listed the element
    // ids of Legacy's dialog. The CENTRAL dialog has more elements, so an
    // unknown id gets an element on demand instead of null.
    document: { getElementById: (id) => registry[id] || (registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }) },
    google: { script: { run: { withSuccessHandler: () => ({ withFailureHandler: () => ({ saveConfigurationSettings: (c) => { sent = c; } }) }) }, host: { close: () => {} } } },
    alert: () => {},
    console: { log: () => {} }
  };
  vm.createContext(globals);
  vm.runInContext(scriptText, globals);
  return { registry, globals, sentConfig: () => sent };
}

// ============================================================ 1/token HTML

console.log('Reviewer API token is never rendered into the dialog HTML');

{
  const { html } = renderDialog(loadCode, { scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  check('the saved token value does not appear anywhere in the rendered HTML', html.indexOf(REAL_LOOKING_TOKEN) === -1, true);
  check('a status line says a token is configured', /Reviewer API token configured/.test(html), true);
  check('the replacement input itself starts blank', /id="apiToken" value=""/.test(html), true);
  check('a Clear-token checkbox exists', /id="clearApiToken"/.test(html), true);
}

{
  const { html } = renderDialog(loadCode, {});
  check('no token configured: status says so', /No Reviewer API token configured/.test(html), true);
}

console.log('BEFORE (baseline): the token WAS rendered directly into the HTML -- this is the exact bug being fixed');

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('baseline token rendering (1 check)'); else {
  const { html } = renderDialog(loadBaselineCode, { scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  check('baseline exposes the real token value in the rendered HTML', html.indexOf(REAL_LOOKING_TOKEN) !== -1, true);
}

// ============================================================ keep/replace/clear

console.log('token keep / replace / clear semantics (server)');

{
  const keep = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  keep.sandbox.saveConfigurationSettings({ apiTokenAction: 'keep', apiToken: '', showPreview: true });
  check('keep leaves the saved token', keep.scriptProps._store.REVIEWER_API_TOKEN, REAL_LOOKING_TOKEN);

  const legacyBlank = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  legacyBlank.sandbox.saveConfigurationSettings({ apiToken: '', showPreview: true });
  check('a caller sending only a blank apiToken (no action) still keeps the token -- unrelated Save preserves it', legacyBlank.scriptProps._store.REVIEWER_API_TOKEN, REAL_LOOKING_TOKEN);

  const replace = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  replace.sandbox.saveConfigurationSettings({ apiTokenAction: 'replace', apiToken: NEW_TOKEN, showPreview: true });
  check('replace stores the new token', replace.scriptProps._store.REVIEWER_API_TOKEN, NEW_TOKEN);

  const blankReplace = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  blankReplace.sandbox.saveConfigurationSettings({ apiTokenAction: 'replace', apiToken: '   ', showPreview: true });
  check('a blank replacement input alone does NOT clear the token', blankReplace.scriptProps._store.REVIEWER_API_TOKEN, REAL_LOOKING_TOKEN);

  const clear = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  clear.sandbox.saveConfigurationSettings({ apiTokenAction: 'clear', apiToken: '', showPreview: true });
  check('explicit clear removes the token', clear.scriptProps._store.REVIEWER_API_TOKEN, undefined);
}

console.log('token keep / replace / clear semantics (client -> payload sent to server)');

{
  const { html } = renderDialog(loadCode, {});
  const s = runClientScript(html, {});
  s.registry.apiToken.value = NEW_TOKEN;
  s.globals.saveConfig();
  check('typing a new token sends replace + the new token', [s.sentConfig().apiTokenAction, s.sentConfig().apiToken], ['replace', NEW_TOKEN]);

  s.registry.clearApiToken.checked = true;
  s.globals.saveConfig();
  check('ticking Clear sends clear (and no token), even if a token was typed', [s.sentConfig().apiTokenAction, s.sentConfig().apiToken], ['clear', '']);

  s.registry.clearApiToken.checked = false;
  s.registry.apiToken.value = '';
  s.globals.saveConfig();
  check('leaving both untouched sends keep', s.sentConfig().apiTokenAction, 'keep');
}

// ============================================================ 2/log redaction

console.log('configuration logging never emits a raw token/secret/password value');

{
  const { sandbox } = loadCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  const logs = [];
  sandbox.Logger = { log: (m) => logs.push(String(m)) };
  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', apiTokenAction: 'replace', apiToken: NEW_TOKEN, showPreview: true });
  const all = logs.join('\n');
  check('the new token is not in any log line', all.indexOf(NEW_TOKEN) === -1, true);
  check('the previously saved token is not in any log line', all.indexOf(REAL_LOOKING_TOKEN) === -1, true);
  check('a save log line still exists (useful diagnostic logging kept)', /Configuration saved/.test(all), true);
  check('the redacted token field is shown as a marker', /\[redacted\]/.test(all), true);
}

console.log('BEFORE (baseline): the raw token WAS logged via JSON.stringify(config) -- this is the exact bug being fixed');

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('baseline token logging (1 check)'); else {
  const { sandbox } = loadBaselineCode({ scriptProperties: { REVIEWER_API_TOKEN: REAL_LOOKING_TOKEN } });
  const logs = [];
  sandbox.Logger = { log: (m) => logs.push(String(m)) };
  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', apiToken: NEW_TOKEN, showPreview: true });
  check('baseline logs the raw new token in plain text', logs.join('\n').indexOf(NEW_TOKEN) !== -1, true);
}

console.log('redactConfigForLog_() -- generic token/secret/password-like keys');

{
  const { sandbox } = loadCode();
  const redacted = sandbox.redactConfigForLog_({ apiToken: NEW_TOKEN, someSecret: 'x', reviewerPassword: 'y', anApiKey: 'z', meetingId: '1', apiTokenAction: 'replace' });
  check('every secret-like key is redacted; ordinary keys and the action itself are not',
    [redacted.apiToken, redacted.someSecret, redacted.reviewerPassword, redacted.anApiKey, redacted.meetingId, redacted.apiTokenAction],
    ['[redacted]', '[redacted]', '[redacted]', '[redacted]', '1', 'replace']);
  check('an empty secret-like value redacts to an empty marker, not "[redacted]" (never implies a real value existed)',
    sandbox.redactConfigForLog_({ apiToken: '' }).apiToken, '');
}

// ============================================================ 3/stale defaults

console.log('ad-hoc save no longer invents historical main-meeting defaults');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({
    meetingFolder: '', meetingNumber: '', meetingId: '', meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG',
    reportType: '', agendaSourceDocId: '', tdocUrl: '', showPreview: true
  });
  check('no MEETING_FOLDER (TSGS4_136_Montreal) written for a blank ad-hoc folder', docProps._store.MEETING_FOLDER, undefined);
  check('no MEETING_NUMBER (136) written for a blank ad-hoc number', docProps._store.MEETING_NUMBER, undefined);
  check('no MEETING_ID (60777) written for a blank ad-hoc ID', docProps._store.MEETING_ID, undefined);
  check('no REPORT_SUFFIX (6G) written for a blank ad-hoc family', docProps._store.REPORT_SUFFIX, undefined);
  check('type and name are still written', [docProps._store.MEETING_TYPE, docProps._store.MEETING_NAME], ['adhoc', 'SA4-(AH) Audio SWG']);
}

{
  // meeting type comes from the ALREADY-SAVED property, not the submitted config
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Audio' } });
  sandbox.saveConfigurationSettings({ meetingId: '85916', reportType: '', apiTokenAction: 'replace', apiToken: NEW_TOKEN, showPreview: true });
  check('a Save that omits meetingType is still classified ad-hoc from the saved property', docProps._store.MEETING_TYPE, 'adhoc');
  check('an existing family is not erased by a blank submission', docProps._store.REPORT_SUFFIX, 'Audio');
  check('folder/number are still not invented; the submitted (non-blank) meeting ID IS written',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID], [undefined, undefined, '85916']);
}

console.log('existing values are never deleted merely because they already exist');

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', MEETING_FOLDER: 'SomeFolderSetManually' } });
  sandbox.saveConfigurationSettings({ meetingId: '85916', meetingType: 'adhoc', meetingFolder: '', reportType: 'Audio', showPreview: true });
  check('a blank submitted folder does not delete an existing manually-set one', docProps._store.MEETING_FOLDER, 'SomeFolderSetManually');
}

console.log('BEFORE (baseline): an ad-hoc save DID invent TSGS4_136_Montreal/136/60777/6G -- this is the exact bug being fixed');

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('baseline invented defaults (1 check)'); else {
  const { sandbox, docProps } = loadBaselineCode();
  sandbox.saveConfigurationSettings({
    meetingFolder: '', meetingNumber: '', meetingId: '', meetingType: 'adhoc', meetingName: 'SA4-(AH) Audio SWG',
    reportType: '', agendaSourceDocId: '', tdocUrl: '', showPreview: true
  });
  check('baseline invents MEETING_FOLDER/NUMBER/ID/REPORT_SUFFIX even for an ad-hoc meeting',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID, docProps._store.REPORT_SUFFIX],
    ['TSGS4_136_Montreal', '136', '60777', '6G']);
}

console.log('main-meeting save behavior remains fully compatible');

{
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingFolder: '', meetingNumber: '', meetingId: '', meetingType: 'main', reportType: '', showPreview: true });
  check('explicit main + blank folder/number/ID keeps the historical defaults, byte-identical to before this stage',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID],
    ['TSGS4_136_Montreal', '136', '60777']);
  // REPORT_SUFFIX is never invented for ANY meeting type, including main --
  // getReportConfig_()'s own '6G' runtime fallback (untouched) covers a
  // genuinely-blank REPORT_SUFFIX. This is a deliberate, small behavior
  // difference from the historical folder/number/ID defaults (which ARE
  // still invented for an explicit main meeting) -- report family has
  // always been a user choice, never a main-meeting-specific invented value.
  check('a blank family is not invented even for an explicit main meeting', docProps._store.REPORT_SUFFIX, undefined);
}

{
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136', REPORT_SUFFIX: '6G' } });
  sandbox.saveConfigurationSettings({
    meetingFolder: 'TSGS4_137_Xian', meetingNumber: '137', meetingId: '12345', meetingType: 'main', reportType: 'Audio', tdocUrl: 'https://example.invalid/list.xlsx', showPreview: false
  });
  check('a legacy main document still updates folder/number/ID/family from submitted values',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID, docProps._store.REPORT_SUFFIX],
    ['TSGS4_137_Xian', '137', '12345', 'Audio']);
}

{
  // A document with MEETING_TYPE never saved AND no meetingType submitted
  // either: effectiveType is '' (neither 'main' nor 'adhoc'), so this falls
  // to the skip-if-blank branch -- matching the SAME effectiveType rule the
  // central add-on's own B1 implementation uses (submittedType || saved
  // MEETING_TYPE, with no third "assume main" fallback). This is REALLY
  // never main-meeting-invented behavior for a document that has truly
  // never been typed at all; a genuinely working main-meeting document
  // already has MEETING_TYPE='main' or a real folder/number saved from a
  // previous Save (see the next case), so this only matters for a document
  // that has never once been saved through this dialog.
  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingFolder: '', meetingNumber: '', meetingId: '', reportType: '', showPreview: true });
  check('an entirely untyped, never-saved document does not have historical defaults invented either (skip-if-blank, matching the B1 effectiveType rule)',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID],
    [undefined, undefined, undefined]);
}

{
  // The realistic legacy-compatibility case: MEETING_TYPE explicitly 'main'
  // was already saved from an earlier Save (every real, previously-used
  // main-meeting document has this, since saveConfigurationSettings has
  // always written it when submitted) -- effectiveType falls back to it
  // and historical defaults are still written exactly as before.
  const { sandbox, docProps } = loadCode({ documentProperties: { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_136_Montreal' } });
  sandbox.saveConfigurationSettings({ meetingFolder: '', meetingNumber: '', meetingId: '', reportType: '', showPreview: true });
  check('a previously-saved main document (MEETING_TYPE already stored) keeps historical defaults on a Save that omits meetingType',
    [docProps._store.MEETING_FOLDER, docProps._store.MEETING_NUMBER, docProps._store.MEETING_ID],
    ['TSGS4_136_Montreal', '136', '60777']);
}

// ============================================================ 4/EMAIL_START_DATE

console.log('EMAIL_START_DATE: no dialog field in the template (intentional difference)');

// INTENTIONAL DIFFERENCE (CENTRAL ADDON-007B1, commit 0510461): the
// Meeting Configuration dialog has no Email Collection Start Date field
// and Save does not write EMAIL_START_DATE. Legacy keeps the field. The
// collector honours a stored EMAIL_START_DATE either way (BUGFIX-LEGACY-003,
// tests/legacy-parity-adhoc-email-collection.test.js).
// TEMPLATE-002B: these suites load the add-on runtime (no Release.js). The
// template runtime restores the field (decision 2026-10-01), covered by
// tests/template002b-runtime.test.js.
{
  const { html } = renderDialog(loadCode, { documentProperties: { EMAIL_START_DATE: '2026-09-01' } });
  check('[add-on runtime] the dialog has no Email Collection Start Date field', /Email Collection Start Date|id="emailStartDate"/.test(html), false);
  const { sandbox, docProps } = loadCode({ documentProperties: { EMAIL_START_DATE: '2026-09-01' } });
  sandbox.saveConfigurationSettings({ emailStartDate: '2026-09-15', showPreview: true });
  check('[add-on runtime] Save neither writes nor erases EMAIL_START_DATE', docProps._store.EMAIL_START_DATE, '2026-09-01');
}

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('collector config byte-identical to the Legacy baseline (2 checks)'); else {
  // NOTE (pre-existing, not something this stage touches or fixes):
  // EMAIL_START_DATE as saved by the dialog lives in Document Properties,
  // but the actual collector config (getCollectorConfig_()) reads its
  // EMAIL_START_DATE from a SEPARATE "Collector Config" table in the
  // document body (readCollectorConfigTable_()), falling back to the
  // hardcoded '2026-08-21' only when that table has no such row. This
  // disconnect already existed before LEGACY-UPGRADE-002 (confirmed
  // identical against the baseline below) -- fixing it is out of this
  // stage's configuration/security scope, so this test only proves the
  // plumbing is BYTE-IDENTICAL to before, not that it is correct.
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const baselineSource = execFileSync('git', ['show', `${BASELINE_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  function fnBody(src, name) {
    const start = src.indexOf('function ' + name + '(');
    const nextMatch = src.slice(start + 1).match(/\nfunction [A-Za-z0-9_]+\(/);
    const end = nextMatch ? start + 1 + nextMatch.index : src.length;
    return src.slice(start, end).replace(/\r\n/g, '\n').trim();
  }
  // The ad-hoc e-mail collection fix deliberately extends
  // getCollectorConfig_() from its EMAIL_START_DATE line onwards (saved
  // start date, meeting-specific mailing list -- covered by
  // tests/legacy-adhoc-email-collection.test.js). Everything before that
  // line is still byte-identical to the pre-stage baseline.
  const upToStartDate = (body) => body.slice(0, body.indexOf('  if (!cfg.EMAIL_START_DATE)'));
  check('getCollectorConfig_() is byte-identical to the pre-stage baseline up to its EMAIL_START_DATE line',
    [upToStartDate(fnBody(source, 'getCollectorConfig_')), fnBody(source, 'getCollectorConfig_').indexOf('  if (!cfg.EMAIL_START_DATE)') > 0],
    [upToStartDate(fnBody(baselineSource, 'getCollectorConfig_')), true]);
  check('readCollectorConfigTable_() is byte-identical to the pre-stage baseline', fnBody(source, 'readCollectorConfigTable_'), fnBody(baselineSource, 'readCollectorConfigTable_'));
}

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  // TEMPLATE-002B: the code that writes EMAIL_START_DATE now exists, for the
  // template runtime only (decision 2026-10-01). What still holds here is
  // that it sits behind the template-runtime switch.
  check('[add-on runtime] EMAIL_START_DATE is written only behind the template-runtime switch',
    [/const emailStartDate = templateRuntimeRelease_\(\) \? String\(config\.emailStartDate/.test(source),
      (source.match(/setProperty\('EMAIL_START_DATE'/g) || []).length], [true, 1]);
  check('the collector still reads cfg.EMAIL_START_DATE (unchanged usage sites)',
    (source.match(/cfg\.EMAIL_START_DATE/g) || []).length >= 3, true);
}

// ============================================================ 5/BUGFIX-LEGACY-001 unchanged

console.log('BUGFIX-LEGACY-001 ad-hoc readiness guard is unchanged');

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'main', MEETING_FOLDER: 'F', MEETING_NUMBER: '1' } });
  // LEGACY-0099 (as ADDON-008A1b): the build's first document access is now
  // reading the saved Document Reallocations (before anything is cleared),
  // so reaching the document means reaching either that read or clear().
  const reachDocument = () => { throw new Error('REACHED_BODY_CLEAR'); };
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ clear: reachDocument, getTables: reachDocument }) });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('main meetings remain unblocked (guard is a no-op, reaches the document)', thrown && thrown.message, 'REACHED_BODY_CLEAR');
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } });
  const calls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => { calls.push('getBody'); return { clear: () => calls.push('clear') }; } });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('incomplete ad-hoc build still fails before touching the document', [thrown !== null, calls], [true, []]);
  check('error message is unchanged', thrown && /Cannot build report yet/.test(thrown.message), true);
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc' } });
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({}) });
  const logs = [];
  sandbox.Logger = { log: (m) => logs.push(String(m)) };
  // continuousUpdateCore_ swallows internal failures via its existing catch
  // -- confirm this stage did not change that.
  sandbox.continuousUpdateCore_();
  check('continuousUpdateCore_ still swallows the readiness error via its own catch (no throw escapes)',
    logs.some((l) => /ERROR:.*Cannot build report yet/.test(l)), true);
}

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('assertMeetingReadyToBuild_ byte-identical to the Legacy baseline (1 check)'); else {
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const baselineSource = execFileSync('git', ['show', `${BASELINE_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  function fnBody(src, name) {
    const start = src.indexOf('function ' + name + '(');
    const nextMatch = src.slice(start + 1).match(/\nfunction [A-Za-z0-9_]+\(/);
    const end = nextMatch ? start + 1 + nextMatch.index : src.length;
    return src.slice(start, end).replace(/\r\n/g, '\n').trim();
  }
  check('assertMeetingReadyToBuild_() source is byte-identical to the pre-stage baseline (untouched by this stage)',
    fnBody(source, 'assertMeetingReadyToBuild_'), fnBody(baselineSource, 'assertMeetingReadyToBuild_'));
}

// ============================================================ structural: diff scope

console.log('structural: this stage touched only the expected configuration/token area');

if (!LEGACY_GIT_HISTORY) skipLegacyHistory('LEGACY-UPGRADE-002 stage scope (2 checks)'); else {
  // Compared against the LEGACY-UPGRADE-002 commit ITSELF (not the
  // pre-002 baseline) -- this makes the "exactly these functions" claim
  // scoped to this stage alone, so it stays valid/re-runnable once later
  // stages (LEGACY-UPGRADE-003+) add their own functions on top.
  const POST_002_COMMIT = '9739085';
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const post002Source = execFileSync('git', ['show', `${POST_002_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  const names = (src) => new Set((src.match(/^function ([A-Za-z0-9_]+)\(/gm) || []).map((m) => m.replace(/^function |\($/g, '')));
  const before = names(post002Source);
  const after = names(source);
  const removed = [...before].filter((n) => !after.has(n));
  check('no function present as of LEGACY-UPGRADE-002 has since been removed', removed, []);
  const baselineSource = execFileSync('git', ['show', `${BASELINE_COMMIT}:Code.js`], { cwd: REPO_ROOT, maxBuffer: 1024 * 1024 * 64 }).toString('utf8');
  const addedSince002Baseline = [...names(post002Source)].filter((n) => !names(baselineSource).has(n));
  check('LEGACY-UPGRADE-002 itself added exactly its two new helper functions', addedSince002Baseline.sort(), ['redactConfigForLog_', 'setPropertyIfPresent_']);
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
