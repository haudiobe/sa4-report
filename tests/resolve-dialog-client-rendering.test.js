/**
 * POST-MEETING-001 (Task A3) — proves/disproves that the Configure Meeting
 * Settings dialog's client-side JavaScript can silently leave "Resolving
 * from 3GPP…" on screen even though the server already returned
 * successfully, because of a client-side rendering exception or missing
 * DOM element rather than an actual network hang.
 *
 * This is NOT a source-structure regex check -- it extracts the REAL
 * <script> block that configureMeetingSettings() embeds in its HTML
 * output, runs it in a Node vm sandbox against a minimal fake DOM/
 * google.script.run, and feeds it REAL preview objects produced by the
 * REAL, unmodified server-side computeResolvedMeetingPreview_()/
 * resolveMeetingCoreById_() (via the same Node sandbox load-code.js uses
 * elsewhere) -- so both halves of the actual Resolve round trip are
 * exercised together, not reimplemented separately.
 *
 * The fake DOM's element registry is built from an EXPLICIT, exact list of
 * the ids the real HTML actually defines (verified directly against
 * Code.js below) -- including the known gap that meetingName/meetingType/
 * meetingDate/ftpBase/mailingList have NO corresponding "...Badge" <span>
 * element in the HTML (only agendaTdocBadge and revisionsUrlBadge exist).
 * getElementById() for a missing id correctly returns null here, exactly
 * like a real browser, so this test can prove whether setFieldWithBadge()'s
 * `if (badgeHost)` guard actually protects every call site or not.
 *
 * Run: node tests/resolve-dialog-client-rendering.test.js
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

// ------------------------------------------------- extract the real script

function extractConfigureMeetingSettingsScript() {
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function configureMeetingSettings()');
  if (fnStart === -1) throw new Error('could not locate configureMeetingSettings() in Code.js');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  const scriptStart = body.indexOf('<script>');
  const scriptEnd = body.lastIndexOf('</script>');
  if (scriptStart === -1 || scriptEnd === -1) throw new Error('could not locate <script>...</script> inside configureMeetingSettings()');
  let scriptText = body.slice(scriptStart + '<script>'.length, scriptEnd);

  // The script is embedded inside a JS template literal with no ${...}
  // interpolations of its own (verified) -- every `\\` the SOURCE FILE
  // contains here is a template-literal-escaped literal backslash (e.g.
  // "\\u2026" in Code.js becomes the two characters … in the actual
  // browser-side script, which the BROWSER then interprets as a real
  // unicode escape). Un-escape exactly that, so this test runs the exact
  // same text a real browser would receive and execute.
  scriptText = scriptText.replace(/\\\\/g, '\\');
  return scriptText;
}

// --------------------------------------------------------- fake DOM/runner

// The EXACT set of element ids the real HTML defines (cross-checked
// against Code.js by hand while writing this test) -- deliberately does
// NOT include "meetingNameBadge"/"meetingTypeBadge"/"meetingDateBadge"/
// "ftpBaseBadge"/"mailingListBadge", which the script REFERENCES but the
// HTML never actually defines a <span> for.
const REAL_ELEMENT_IDS = [
  'meetingId', 'resolveBtn', 'resolveStatus', 'agendaCandidates',
  'portalTypeHint', 'dateRangeHint',
  'meetingName', 'meetingType', 'meetingDate', 'ftpBase',
  'agendaTdoc', 'agendaTdocBadge', 'discoverBtn', 'discoverStatus',
  'mailingList', 'revisionsUrl', 'revisionsUrlBadge',
  'meetingFolder', 'meetingNumber', 'reportType', 'agendaSourceDocId',
  'tdocUrl', 'showPreview', 'apiToken'
];

function makeFakeElement(id) {
  return { id, value: '', textContent: '', className: '', disabled: false, innerHTML: '', checked: false };
}

function makeFakeDocument() {
  const registry = {};
  REAL_ELEMENT_IDS.forEach(id => { registry[id] = makeFakeElement(id); });
  return { getElementById: (id) => (Object.prototype.hasOwnProperty.call(registry, id) ? registry[id] : null), _registry: registry };
}

// Mimics google.script.run's chained withSuccessHandler/withFailureHandler,
// then invokes whichever handler applies once the named "server function"
// is called -- letting each test control exactly what the "server"
// returned, using REAL result objects computed by the real server-side
// functions below.
function makeFakeRunner(response) {
  let successHandler = null;
  let failureHandler = null;
  const runner = {
    withSuccessHandler: (fn) => { successHandler = fn; return runner; },
    withFailureHandler: (fn) => { failureHandler = fn; return runner; }
  };
  function invoke() {
    if (response.throwError !== undefined) {
      if (failureHandler) failureHandler(response.throwError);
      return;
    }
    if (successHandler) successHandler(response.result);
  }
  runner.resolveMeetingForConfigDialog_ = invoke;
  runner.discoverAgendaForConfigDialog_ = invoke;
  runner.saveConfigurationSettings = invoke;
  return runner;
}

function loadDialogSandbox(runnerResponse) {
  const scriptText = extractConfigureMeetingSettingsScript();
  const fakeDocument = makeFakeDocument();
  const sandboxGlobals = {
    document: fakeDocument,
    google: { script: { run: makeFakeRunner(runnerResponse || { result: {} }), host: { close: () => {} } } },
    alert: () => {},
    console
  };
  vm.createContext(sandboxGlobals);
  vm.runInContext(scriptText, sandboxGlobals, { filename: 'configureMeetingSettings-dialog-script.js' });
  return { sandboxGlobals, fakeDocument };
}

// ======================== 1. real preview objects, real client rendering ==

console.log('resolveMeeting() -- real server preview objects rendered by the real client script, no exception, status resets');

{
  // A REAL resolveMeetingForConfigDialog_() result for 86178, produced by
  // the actual, unmodified server-side functions -- not a hand-built mock.
  const { sandbox: serverSandbox } = loadCode({ documentProperties: { MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED', TDOC_LIST_URL: 'existing-url.xlsx', AGENDA_TDOC: 'S4aP260098' } });
  const fs2 = require('fs');
  const fixtures = JSON.parse(fs2.readFileSync(require('path').join(__dirname, 'fixtures', 'meeting-resolver-samples.json'), 'utf8'));
  serverSandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(fixtures.getMeetings86178) });
  const realResult = serverSandbox.resolveMeetingForConfigDialog_('86178');
  check('sanity: the real server result is ok:true for 86178', realResult.ok, true);

  const { sandboxGlobals, fakeDocument } = loadDialogSandbox({ result: realResult });

  let threw = null;
  try {
    sandboxGlobals.resolveMeeting();
  } catch (e) {
    threw = e;
  }

  check('resolveMeeting() success handler does not throw when rendering a REAL 86178 preview', threw, null);
  check('status text is NOT stuck on "Resolving from 3GPP…" after a successful resolve',
    fakeDocument._registry.resolveStatus.textContent.indexOf('Resolving from 3GPP') === -1, true);
  check('status text ends up showing the "Resolved" message (since 86178 core-resolves cleanly, no warnings)',
    fakeDocument._registry.resolveStatus.textContent.indexOf('Resolved.') !== -1, true);
  check('resolve button is re-enabled', fakeDocument._registry.resolveBtn.disabled, false);
  check('meetingName input reflects the real resolved value', fakeDocument._registry.meetingName.value, 'SA4-e (AH) on FS_6G_MED');
  check('meetingType input reflects the real resolved value', fakeDocument._registry.meetingType.value, 'adhoc');
  check('ftpBase input reflects the real resolved value', /\/SA4_Plenary\/Docs\/$/.test(fakeDocument._registry.ftpBase.value), true);
  check('agendaTdoc input keeps the EXISTING manual value (core resolve never looks this up)', fakeDocument._registry.agendaTdoc.value, 'S4aP260098');
  check('mailingList input keeps the EXISTING manual value', fakeDocument._registry.mailingList.value, '3GPP_TSG_SA4_FS_6G_MED');
  check('revisionsUrl input shows the derived candidate', /\/inbox\/drafts\/$/.test(fakeDocument._registry.revisionsUrl.value), true);

  // The known gap: badge spans that do not exist in the HTML must not
  // throw (guarded), and are simply left unset -- proving the guard works
  // for every one of them, not just the two that DO have real spans.
  check('non-existent "...Badge" elements are correctly null in this fake DOM (matches the real HTML gap)',
    ['meetingNameBadge', 'meetingTypeBadge', 'meetingDateBadge', 'ftpBaseBadge', 'mailingListBadge'].every(id => fakeDocument.getElementById(id) === null),
    true);
  check('agendaTdocBadge (which DOES exist in the real HTML) was actually updated', fakeDocument._registry.agendaTdocBadge.innerHTML !== '', true);
  check('revisionsUrlBadge (which DOES exist in the real HTML) was actually updated', fakeDocument._registry.revisionsUrlBadge.innerHTML !== '', true);
}

// ============ 2. a genuinely unresolved/degraded preview still renders =====

console.log('resolveMeeting() -- an unresolved/degraded real preview (no existing config at all) still renders without throwing');

{
  const { sandbox: serverSandbox } = loadCode(); // no Document Properties at all
  const fs2 = require('fs');
  const fixtures = JSON.parse(fs2.readFileSync(require('path').join(__dirname, 'fixtures', 'meeting-resolver-samples.json'), 'utf8'));
  serverSandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(fixtures.getMeetings86174) }); // 86174: no agenda TDoc, no existing config
  const realResult = serverSandbox.resolveMeetingForConfigDialog_('86174');
  check('sanity: real server result is ok:true for 86174 (meeting found, even though agenda/mailing list are unresolved)', realResult.ok, true);

  const { sandboxGlobals, fakeDocument } = loadDialogSandbox({ result: realResult });
  let threw = null;
  try { sandboxGlobals.resolveMeeting(); } catch (e) { threw = e; }

  check('does not throw even with a fully-blank existing configuration', threw, null);
  check('status is not stuck on "Resolving from 3GPP…"', fakeDocument._registry.resolveStatus.textContent.indexOf('Resolving from 3GPP') === -1, true);
  check('agendaTdoc input is left blank (genuinely unresolved, never guessed)', fakeDocument._registry.agendaTdoc.value, '');
  check('mailingList input is left blank (never auto-resolved)', fakeDocument._registry.mailingList.value, '');
}

// ==================== 3. an invalid Meeting ID (ok:false) still renders ====

console.log('resolveMeeting() -- a validation failure (ok:false) result still resets the UI without throwing');

{
  const { sandbox: serverSandbox } = loadCode();
  const realResult = serverSandbox.resolveMeetingForConfigDialog_('not-a-number');
  check('sanity: real server result is ok:false for an invalid id', realResult.ok, false);

  const { sandboxGlobals, fakeDocument } = loadDialogSandbox({ result: realResult });
  let threw = null;
  try { sandboxGlobals.resolveMeeting(); } catch (e) { threw = e; }

  check('does not throw on an ok:false result', threw, null);
  check('status shows the error, not stuck on "Resolving from 3GPP…"', fakeDocument._registry.resolveStatus.textContent.indexOf('Resolving from 3GPP') === -1, true);
  check('resolve button is re-enabled even on a validation failure', fakeDocument._registry.resolveBtn.disabled, false);
}

// ================ 4. google.script.run itself failing (network/server error)

console.log('resolveMeeting() -- withFailureHandler path (e.g. a genuine server-side exception) resets the UI');

{
  const { sandboxGlobals, fakeDocument } = loadDialogSandbox({ throwError: 'simulated Apps Script server error' });
  let threw = null;
  try { sandboxGlobals.resolveMeeting(); } catch (e) { threw = e; }

  check('the failure handler itself does not throw', threw, null);
  check('status shows the failure, not stuck on "Resolving from 3GPP…"', fakeDocument._registry.resolveStatus.textContent.indexOf('Resolving from 3GPP') === -1, true);
  check('status mentions the failure reason', fakeDocument._registry.resolveStatus.textContent.indexOf('simulated Apps Script server error') !== -1, true);
  check('resolve button is re-enabled after a failure', fakeDocument._registry.resolveBtn.disabled, false);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
