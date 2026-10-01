/**
 * TEMPLATE-002A -- behaviour that intentionally differs from the accepted
 * Legacy script, in one place. Each block names the decision behind it.
 * The Legacy-side counterpart of every block is marked "INTENTIONAL
 * DIFFERENCE" in tests/legacy-parity-*.test.js; see
 * docs/TEMPLATE-002A_LEGACY_PARITY.md.
 *
 * Also covers the two places where a ported Legacy fix had to follow this
 * repository's execution-context design (state store instead of Document
 * Properties).
 *
 * Run: node tests/template002a-intentional-differences.test.js
 */

const fs = require('fs');
const vm = require('vm');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
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

const SOURCE = fs.readFileSync(CODE_JS_PATH, 'utf8');
const stripComments = (src) => src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
const CODE_ONLY = stripComments(SOURCE);

// ---------------------------------------------------------------- 1
console.log('1. outgoing discussion e-mail introduction is meeting-neutral (decision 2026-10-01)');
{
  const { sandbox: s } = loadCode();
  const mbs = s.buildEmailExportDefaultIntroText_('MBS');
  const sixG = s.buildEmailExportDefaultIntroText_('FS_6G_MED');
  check('the introduction names the meeting\'s own tag', [/As discussed during the MBS AHG/.test(mbs), /As discussed during the FS_6G_MED AHG/.test(sixG)], [true, true]);
  check('an MBS report\'s introduction never mentions FS_6G_MED', /FS_6G_MED|6G/.test(mbs), false);
  check('"the upcoming meeting", never Legacy\'s "the October meeting"', [/for the upcoming meeting\./.test(mbs), /October/.test(mbs)], [true, false]);
  check('no month name or FS_6G_MED literal in the default wording itself',
    /October|FS_6G_MED/.test(vm.runInContext('EMAIL_EXPORT_DEFAULT_INTRO_TEXT_TEMPLATE_', s)), false);
  check('no prefilled discussion deadline (Legacy: 2026-10-15 15:00) anywhere in the code', /2026-10-15/.test(CODE_ONLY), false);
}

// ---------------------------------------------------------------- 2
console.log('2. subject tag is derived from the meeting\'s mailing list, never hard-coded (decision 2026-10-01)');
{
  const { sandbox: s } = loadCode();
  check('MBS list -> MBS', s.deriveEmailExportListTag_('3gpp_tsg_sa_wg4_mbs@list.etsi.org', 'MBS'), 'MBS');
  check('Audio list -> AUDIO', s.deriveEmailExportListTag_('3gpp_tsg_sa_wg4_audio@list.etsi.org', 'Audio'), 'AUDIO');
  check('the 6G list still yields Legacy\'s FS_6G_MED', s.deriveEmailExportListTag_('3gpp_tsg_sa4_fs_6g_med@list.etsi.org', 'FS_6G_MED'), 'FS_6G_MED');
  check('the subject carries the tag it is given',
    s.buildEmailExportSubject_('S4aI260082', 'Title', '2.5', '26-10-15-1500CEST', 'MBS'), '[MBS,2.5,26-10-15-1500CEST][S4aI260082] Discussion: Title');
  let refused = null;
  try { s.buildEmailExportSubject_('S4aI260082', 'Title', '2.5', '26-10-15-1500CEST'); } catch (e) { refused = e.message; }
  check('without a tag no subject is built (no built-in default)', refused, 'buildEmailExportSubject_: a list tag is required.');
  check('Legacy\'s constant EMAIL_EXPORT_LIST_TAG_ does not exist', /EMAIL_EXPORT_LIST_TAG_/.test(CODE_ONLY), false);
  check('no quoted FS_6G_MED tag literal is passed to the exporter', /\[FS_6G_MED|'FS_6G_MED,|"FS_6G_MED,/.test(CODE_ONLY), false);
}

// ---------------------------------------------------------------- 3
// TEMPLATE-002B: true for the add-on runtime this file loads (no Release.js).
// The template runtime restores the field (decision 2026-10-01); see
// tests/template002b-runtime.test.js.
console.log('3. no Email Collection Start Date field in the dialog outside the template runtime (CENTRAL ADDON-007B1)');
{
  const loaded = loadCode({ documentProperties: { EMAIL_START_DATE: '2026-09-15' } });
  const s = loaded.sandbox;
  let html = null;
  s.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  s.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  s.configureMeetingSettings();
  check('the dialog offers no start-date input', /emailStartDate|Email Collection Start Date/.test(html), false);
  s.saveConfigurationSettings({ emailStartDate: '2026-12-01', showPreview: true });
  check('Save leaves a stored EMAIL_START_DATE exactly as it is', loaded.docProps._store.EMAIL_START_DATE, '2026-09-15');
}

// ---------------------------------------------------------------- 4
console.log('4. Continuous Update interval: hourly works, unoffered intervals are refused first (CENTRAL 2.15.1/2.15.2)');
{
  const { sandbox: s } = loadCode();
  const calls = [];
  s.ScriptApp = {
    getProjectTriggers: () => [{ getHandlerFunction: () => 'continuousUpdate' }],
    deleteTrigger: () => calls.push('delete'),
    newTrigger: (fn) => ({ timeBased: () => ({
      everyMinutes: (n) => { calls.push('everyMinutes(' + n + ')'); throw new Error('The value you passed to everyMinutes was invalid.'); },
      everyHours: (n) => ({ create: () => calls.push('everyHours(' + n + ') ' + fn) })
    }) })
  };
  s.createContinuousTrigger(60, false);
  check('"every hour" is everyHours(1), not Legacy\'s everyMinutes(60)', calls, ['delete', 'everyHours(1) continuousUpdate']);
  calls.length = 0;
  let err = null;
  try { s.createContinuousTrigger(30, false); } catch (e) { err = e.message; }
  check('30 minutes is refused before the running trigger is touched (the template runtime offers it again in TEMPLATE-002B)',
    [/^Unsupported update interval: 30/.test(err), calls], [true, []]);
}

// ---------------------------------------------------------------- 5
console.log('5. ported Legacy fixes follow this repository\'s execution context');
{
  // Interactive / bound: Document Properties, exactly as in Legacy.
  const bound = loadCode({ documentProperties: { EMAIL_START_DATE: '2026-09-15' } });
  bound.sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => makeFakeDocumentBody(bound.sandbox), getId: () => 'DOC' });
  check('bound: the stored start date is the collector\'s start date', bound.sandbox.getCollectorConfig_().EMAIL_START_DATE, '2026-09-15');

  // Background: the adopted document's central state, not Document Properties.
  const bg = loadCode({
    documentProperties: { EMAIL_START_DATE: '2026-01-01' },
    scriptProperties: { 'SA4_STATE|DOC|EMAIL_START_DATE': '2026-09-20' }
  });
  check('background: read from the document\'s central state',
    bg.sandbox.resolveCollectorStartDate_({ mode: 'addon-background', documentId: 'DOC' }), '2026-09-20');
  const bgNone = loadCode({ documentProperties: { EMAIL_START_DATE: '2026-01-01' } });
  check('background without central value: the default, never another document\'s property',
    bgNone.sandbox.resolveCollectorStartDate_({ mode: 'addon-background', documentId: 'DOC' }), '2026-08-21');

  const logs = [];
  const bad = loadCode({ documentProperties: { EMAIL_START_DATE: '15.09.2026' } });
  bad.sandbox.Logger = { log: (m) => logs.push(String(m)) };
  check('a malformed stored date falls back to the default and is logged',
    [bad.sandbox.resolveCollectorStartDate_(), logs.some((l) => /Ignoring EMAIL_START_DATE "15\.09\.2026"/.test(l))], ['2026-08-21', true]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll checks passed.');
process.exitCode = failures ? 1 : 0;
