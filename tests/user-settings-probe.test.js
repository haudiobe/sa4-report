/**
 * The personal-settings probe (template/probe/UserSettingsProbe.gs), run
 * against the fake Drive.
 *
 * The probe is the one live check the personal Reviewer token needs: report A
 * creates a sentinel settings file; a DIFFERENT report B finds and reads it;
 * so does a one-shot trigger of B. Here its logic is run with fakes, so that
 * what is later run live is known to do what it says -- and to be safe:
 *
 *   1. A creates, B reads, B's trigger reads: PASS each time;
 *   2. the sentinel is never shown, logged or stored by the probe;
 *   3. it never touches settings that are not its own sentinel;
 *   4. what it reports when something is wrong;
 *   5. it is not part of any release bundle.
 *
 * Nothing live is used; this test does not run the probe against Google.
 *
 * Run: node tests/user-settings-probe.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const { loadTemplateRuntime } = require('./helpers/load-template.js');
const { makeFakeDrive } = require('./helpers/fake-drive-files.js');
const { RELEASE } = require('./helpers/config-dialog.js');

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

const PROBE_PATH = path.join(__dirname, '..', 'template', 'probe', 'UserSettingsProbe.gs');
const PROBE = fs.readFileSync(PROBE_PATH, 'utf8');
const SENTINEL = 'sa4-user-settings-probe-sentinel-not-a-token';
const SETTINGS = (token) => JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: token });

/** One scratch report: its own script project (id), the Drive of the user, and the probe pasted in. */
function scratchReport(scriptId, drive) {
  const loaded = loadTemplateRuntime({ release: RELEASE, documentProperties: {} });
  const s = loaded.sandbox;
  drive.install(s);
  const r = { s: s, docProps: loaded.docProps, logs: [], alerts: [], triggers: [], drive: drive };
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getUi = () => ({ alert: (title, text) => { r.alerts.push([title, text]); return 'OK'; }, ButtonSet: { OK: 'OK' } });
  s.Utilities = Object.assign({}, s.Utilities, { DigestAlgorithm: { SHA_256: 'SHA_256' }, Charset: { UTF_8: 'UTF_8' },
    computeDigest: (algorithm, text) => Array.from(crypto.createHash('sha256').update(String(text), 'utf8').digest()).map((b) => (b > 127 ? b - 256 : b)) });
  s.ScriptApp = {
    getScriptId: () => scriptId,
    getProjectTriggers: () => r.triggers.map((t) => ({ getHandlerFunction: () => t.handler, _t: t })),
    deleteTrigger: (t) => { r.triggers = r.triggers.filter((x) => x !== t._t); },
    newTrigger: (handler) => ({ timeBased: () => ({ after: (ms) => ({ create: () => { r.triggers.push({ handler: handler, after: ms }); } }) }) })
  };
  vm.runInContext(PROBE, s, { filename: 'UserSettingsProbe.gs' });
  /** The trigger fires: its handler runs without a user interface. */
  r.fire = () => { const ui = s.DocumentApp.getUi; s.DocumentApp.getUi = () => { throw new Error('no UI in a trigger'); }; try { r.triggers.slice().forEach((t) => s[t.handler]()); } finally { s.DocumentApp.getUi = ui; } };
  r.everythingShown = () => JSON.stringify([r.logs, r.alerts, r.docProps._store]);
  return r;
}

console.log('1. A creates, B reads, B\'s trigger reads');
{
  const drive = makeFakeDrive();
  const a = scratchReport('SCRIPT-A', drive);
  a.s.userSettingsProbeA_Create();
  check('A: the sentinel file is created, private and marked, and found again', [a.alerts[0][0], drive.live().map((f) => [f.content, f.shared, f.ownedByMe, f.properties, f.mimeType])],
    ['Probe A: PASS', [[SETTINGS(SENTINEL), false, true, { sa4ReportUserSettings: '1' }, 'application/json']]]);
  const b = scratchReport('SCRIPT-B', drive);
  check('B is another script project with nothing stored in it', [b.s.ScriptApp.getScriptId() !== a.s.ScriptApp.getScriptId(), b.docProps.getKeys()], [true, []]);
  b.s.userSettingsProbeB_Read();
  check('B, from the menu: finds exactly one file and it holds the sentinel', [b.alerts[0][0], /status: ok, usable files: 1, private: 1, shared: 0/.test(b.alerts[0][1]), /holds the sentinel: true/.test(b.alerts[0][1])], ['Probe B (menu): PASS', true, true]);
  b.s.userSettingsProbeB_InstallTrigger();
  check('B: a one-shot trigger is installed, a minute ahead', b.triggers, [{ handler: 'userSettingsProbeTriggered', after: 60000 }]);
  b.s.userSettingsProbeB_ShowTriggerResult();
  check('before it ran there is no result', b.alerts[b.alerts.length - 1][0], 'Probe B (trigger): NO RESULT YET');
  b.fire();
  check('the trigger ran without a user interface, recorded its result and removed itself', [b.triggers, JSON.parse(b.docProps.getProperty('USER_SETTINGS_PROBE_TRIGGER_RESULT')).verdict, JSON.parse(b.docProps.getProperty('USER_SETTINGS_PROBE_TRIGGER_RESULT')).observation],
    [[], 'PASS', { status: 'ok', usableFiles: 1, privateFiles: 1, sharedFiles: 0, matches: true }]);
  b.s.userSettingsProbeB_ShowTriggerResult();
  check('B, from the trigger: PASS', b.alerts[b.alerts.length - 1][0], 'Probe B (trigger): PASS');

  console.log('2. the sentinel is never shown, logged or stored');
  check('nothing A or B showed, logged or stored contains the sentinel', [a.everythingShown().indexOf(SENTINEL), b.everythingShown().indexOf(SENTINEL)], [-1, -1]);
  check('the probe compares hashes, and prints neither the content nor a hash (source)',
    [/hash\(f\.token\) === hash\(USER_SETTINGS_PROBE_SENTINEL_\)/.test(PROBE), /Logger\.log\([^)]*token/.test(PROBE), /alert\([^)]*token/.test(PROBE), /\.token\b[^=]*\+|'\s*\+\s*found\.token/.test(PROBE)], [true, false, false, false]);

  b.s.userSettingsProbeCleanup();
  check('cleanup: the sentinel file goes to the trash, the trigger result is removed', [b.alerts[b.alerts.length - 1][0], drive.live().length, drive.files.map((f) => f.trashed), b.docProps.getKeys()], ['Probe cleanup: done', 0, [true], []]);
  b.s.userSettingsProbeCleanup();
  check('a second cleanup has nothing to do', b.alerts[b.alerts.length - 1][0], 'Probe cleanup: nothing to remove');
}

console.log('3. it never touches settings that are not its own sentinel');
{
  const REAL = 'a-real-personal-token-1234';
  const drive = makeFakeDrive();
  drive.addSettings(REAL, { id: 'real' });
  const a = scratchReport('SCRIPT-A', drive);
  a.s.userSettingsProbeA_Create();
  check('an account that has a personal token: step A does not run and changes nothing', [a.alerts[0][0], drive.files.map((f) => [f.id, f.content, f.trashed]), drive.calls.map((c) => c[0]).filter((n) => n !== 'list' && n !== 'read')],
    ['Probe A: NOT RUN', [['real', SETTINGS(REAL), false]], []]);
  a.s.userSettingsProbeCleanup();
  check('and cleanup does not remove it', [a.alerts[a.alerts.length - 1][0], drive.files.map((f) => f.trashed)], ['Probe cleanup: NOT DONE', [false]]);
  check('the real token is not shown by the probe either', a.everythingShown().indexOf(REAL), -1);
  const b = scratchReport('SCRIPT-B', drive);
  b.s.userSettingsProbeB_Read();
  check('B reports a file that does not hold the sentinel as a failure of the probe', b.alerts[0][0], 'Probe B (menu): FAIL: one settings file was found, but it does not hold the sentinel');
  const shared = makeFakeDrive();
  shared.addSettings(REAL, { shared: true });
  const c = scratchReport('SCRIPT-C', shared);
  c.s.userSettingsProbeA_Create();
  c.s.userSettingsProbeCleanup();
  check('with a shared settings file: neither step A nor cleanup does anything', [c.alerts.map((x) => x[0]), shared.files.map((f) => [f.shared, f.trashed]), shared.files.length], [['Probe A: NOT RUN', 'Probe cleanup: NOT DONE'], [[true, false]], 1]);
}

console.log('4. what it reports when something is wrong');
{
  const none = scratchReport('SCRIPT-B', makeFakeDrive());
  none.s.userSettingsProbeB_Read();
  check('no file', none.alerts[0][0], 'Probe B (menu): FAIL: no settings file was found');
  const two = makeFakeDrive();
  two.addSettings(SENTINEL);
  two.addSettings(SENTINEL);
  const twice = scratchReport('SCRIPT-B', two);
  twice.s.userSettingsProbeB_Read();
  check('two files', twice.alerts[0][0], 'Probe B (menu): FAIL: status "ambiguous" (usable files: 2, private: 2, shared: 0)');
  const down = makeFakeDrive();
  down.addSettings(SENTINEL);
  down.failList = true;
  const dead = scratchReport('SCRIPT-B', down);
  dead.s.userSettingsProbeB_InstallTrigger();
  dead.fire();
  check('Drive cannot be read from the trigger: recorded as a failure, and the trigger still removes itself',
    [JSON.parse(dead.docProps.getProperty('USER_SETTINGS_PROBE_TRIGGER_RESULT')).verdict, dead.triggers], ['FAIL: status "unreadable" (usable files: 0, private: 0, shared: 0)', []]);
  const v = none.s.userSettingsProbeVerdict_;
  check('the verdict', [v({ status: 'ok', usableFiles: 1, matches: true }), v({ status: 'ok', usableFiles: 1, matches: false }), v({ status: 'none', usableFiles: 0 }), v(null)],
    ['PASS', 'FAIL: one settings file was found, but it does not hold the sentinel', 'FAIL: no settings file was found', 'FAIL: status "undefined" (usable files: undefined, private: undefined, shared: undefined)']);
}

console.log('5. it is not part of any release bundle');
{
  const tool = fs.readFileSync(path.join(__dirname, '..', 'tools', 'template-release.js'), 'utf8');
  const claspignore = fs.readFileSync(path.join(__dirname, '..', '.claspignore'), 'utf8');
  check('the release tool and the push filter do not name the probe or its folder', [/UserSettingsProbe|template\/probe|probe\//.test(tool), /UserSettingsProbe|probe/.test(claspignore)], [false, false]);
  check('it uses the production functions, and has no Drive call of its own (source)',
    [/userSettingsDrive_\(\)/.test(PROBE), /writePersonalReviewerTokenWith_\(drive, USER_SETTINGS_PROBE_SENTINEL_\)/.test(PROBE), /inspectPersonalReviewerSettingsWith_\(/.test(PROBE), /Drive\.Files|DriveApp|UrlFetchApp/.test(PROBE.replace(/\/\*[\s\S]*?\*\//g, ''))], [true, true, true, false]);
  check('it makes no request to the Reviewer and uses no real token (source)', /reviewer\.bouazizi|REVIEWER_API_TOKEN|X-API-Key/.test(PROBE), false);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll personal-settings probe checks passed.');
process.exitCode = failures ? 1 : 0;
