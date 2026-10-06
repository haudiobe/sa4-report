/*******************************
 * PERSONAL-SETTINGS PROBE -- one live check of the personal Reviewer token
 * store, with a harmless sentinel instead of a token.
 *
 * NOT part of any release bundle (the release tool takes named files only,
 * and none from template/probe/). It is pasted, by hand, into the bound
 * script of SCRATCH reports that carry the candidate code (Code.js,
 * ReportCreator.js, Release.js), and removed with them afterwards.
 *
 * WHAT ONLY A LIVE RUN CAN SHOW. The automated tests run against a fake
 * Drive. Three things about the real one are assumptions until seen:
 *   A. a report can create the marked, private settings file;
 *   B. a DIFFERENT report -- a fresh copy, another script project -- finds
 *      and reads that file by its marker;
 *   C. so does a time-driven trigger created in that other report.
 *
 * HOW (about ten minutes, in a Google account that has NO personal Reviewer
 * token yet -- step A refuses otherwise, so a real token is never touched):
 *   1. Scratch report A:  run  userSettingsProbeA_Create
 *   2. Scratch report B (a fresh copy of the scratch template, never opened
 *      before):           run  userSettingsProbeB_Read            -> expect PASS
 *   3. In B:              run  userSettingsProbeB_InstallTrigger
 *      wait two minutes,  run  userSettingsProbeB_ShowTriggerResult -> expect PASS
 *   4. In B (or A):       run  userSettingsProbeCleanup
 *   5. Trash both scratch reports.
 * Run the functions from the Apps Script editor (function list); each shows
 * its result in a dialog and in the execution log.
 *
 * While the sentinel file exists, template reports of this account would
 * take the sentinel for a personal Reviewer token (and the Reviewer would
 * refuse it). Hence a scratch account or a quiet moment, and step 4.
 *
 * NOTHING SECRET IS INVOLVED, and the probe still never shows what it read:
 * it reports the status, counts, and whether a SHA-256 of what it read
 * equals the SHA-256 of the sentinel.
 *
 * It calls the production functions themselves (userSettingsDrive_(),
 * writePersonalReviewerTokenWith_(), inspectPersonalReviewerSettingsWith_(),
 * clearPersonalReviewerTokenWith_()), so what it proves is the code that
 * will run.
 *******************************/

var USER_SETTINGS_PROBE_SENTINEL_ = 'sa4-user-settings-probe-sentinel-not-a-token';
var USER_SETTINGS_PROBE_RESULT_KEY_ = 'USER_SETTINGS_PROBE_TRIGGER_RESULT';
var USER_SETTINGS_PROBE_HANDLER_ = 'userSettingsProbeTriggered';

/** SHA-256 of a text, as hex. */
function userSettingsProbeHash_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

/**
 * Pure (but for the hash function it is given): what was found, without the
 * content -- { status, usableFiles, privateFiles, sharedFiles, matches }.
 */
function userSettingsProbeObservation_(found, hash) {
  const f = found || {};
  return {
    status: String(f.status || ''),
    usableFiles: (f.validIds || []).length,
    privateFiles: (f.privateIds || []).length,
    sharedFiles: Number(f.sharedCount || 0),
    matches: f.status === 'ok' && hash(f.token) === hash(USER_SETTINGS_PROBE_SENTINEL_)
  };
}

/** Pure: PASS, or what is wrong. */
function userSettingsProbeVerdict_(observation) {
  const o = observation || {};
  if (o.status === 'ok' && o.usableFiles === 1 && o.matches === true) return 'PASS';
  if (o.status === 'ok') return 'FAIL: one settings file was found, but it does not hold the sentinel';
  if (o.status === 'none') return 'FAIL: no settings file was found';
  return 'FAIL: status "' + o.status + '" (usable files: ' + o.usableFiles + ', private: ' + o.privateFiles + ', shared: ' + o.sharedFiles + ')';
}

function userSettingsProbeObserve_() {
  return userSettingsProbeObservation_(inspectPersonalReviewerSettingsWith_(userSettingsDrive_()), userSettingsProbeHash_);
}

function userSettingsProbeSay_(title, lines) {
  const text = lines.join('\n');
  Logger.log('[USER-SETTINGS-PROBE] ' + title + ' -- ' + lines.join(' | '));
  try {
    DocumentApp.getUi().alert(title, text, DocumentApp.getUi().ButtonSet.OK);
  } catch (e) {
    // no dialog in this context (a trigger): the log has it
  }
  return text;
}

/** Step 1, in scratch report A: creates the sentinel settings file. Refuses when the account has any settings file already. */
function userSettingsProbeA_Create() {
  const drive = userSettingsDrive_();
  const before = inspectPersonalReviewerSettingsWith_(drive);
  if (before.status !== 'none') {
    return userSettingsProbeSay_('Probe A: NOT RUN', ['This account already has personal settings (status "' + before.status + '").',
      'The probe does not touch them. Use an account without a personal Reviewer token.']);
  }
  const written = writePersonalReviewerTokenWith_(drive, USER_SETTINGS_PROBE_SENTINEL_);
  const after = userSettingsProbeObserve_();
  return userSettingsProbeSay_('Probe A: ' + (written.ok && userSettingsProbeVerdict_(after) === 'PASS' ? 'PASS' : 'FAIL'), [
    'written and confirmed: ' + written.ok + (written.ok ? '' : ' (' + written.reason + ')'),
    'found again by its marker: ' + userSettingsProbeVerdict_(after),
    'script: ' + ScriptApp.getScriptId(),
    'Now run userSettingsProbeB_Read in a DIFFERENT, fresh scratch report.']);
}

/** Step 2, in scratch report B: finds and reads the file A created. */
function userSettingsProbeB_Read() {
  const seen = userSettingsProbeObserve_();
  return userSettingsProbeSay_('Probe B (menu): ' + userSettingsProbeVerdict_(seen), [
    'status: ' + seen.status + ', usable files: ' + seen.usableFiles + ', private: ' + seen.privateFiles + ', shared: ' + seen.sharedFiles,
    'holds the sentinel: ' + seen.matches,
    'script: ' + ScriptApp.getScriptId() + ' (must differ from the script of report A)']);
}

/** Step 3, in scratch report B: a one-shot trigger, a minute from now, that does the same under the trigger's identity. */
function userSettingsProbeB_InstallTrigger() {
  userSettingsProbeRemoveTriggers_();
  PropertiesService.getDocumentProperties().deleteProperty(USER_SETTINGS_PROBE_RESULT_KEY_);
  ScriptApp.newTrigger(USER_SETTINGS_PROBE_HANDLER_).timeBased().after(60 * 1000).create();
  return userSettingsProbeSay_('Probe B (trigger): installed', ['It runs once, in about a minute.', 'Then run userSettingsProbeB_ShowTriggerResult.']);
}

/** The trigger's handler: records what it saw (never the content) and removes itself. */
function userSettingsProbeTriggered() {
  let result;
  try {
    const seen = userSettingsProbeObserve_();
    result = { verdict: userSettingsProbeVerdict_(seen), observation: seen, at: new Date().toISOString() };
  } catch (e) {
    result = { verdict: 'FAIL: the trigger could not look (' + e.message + ')', observation: null, at: new Date().toISOString() };
  }
  PropertiesService.getDocumentProperties().setProperty(USER_SETTINGS_PROBE_RESULT_KEY_, JSON.stringify(result));
  Logger.log('[USER-SETTINGS-PROBE] trigger -- ' + result.verdict);
  userSettingsProbeRemoveTriggers_();
}

/** Step 3, afterwards: what the trigger saw. */
function userSettingsProbeB_ShowTriggerResult() {
  const raw = PropertiesService.getDocumentProperties().getProperty(USER_SETTINGS_PROBE_RESULT_KEY_);
  if (!raw) return userSettingsProbeSay_('Probe B (trigger): NO RESULT YET', ['The trigger has not run (or was not installed). Wait a minute and try again.']);
  const result = JSON.parse(raw);
  const o = result.observation || {};
  return userSettingsProbeSay_('Probe B (trigger): ' + result.verdict, ['ran at ' + result.at,
    'status: ' + o.status + ', usable files: ' + o.usableFiles + ', private: ' + o.privateFiles + ', shared: ' + o.sharedFiles, 'holds the sentinel: ' + o.matches]);
}

function userSettingsProbeRemoveTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === USER_SETTINGS_PROBE_HANDLER_) ScriptApp.deleteTrigger(t);
  });
}

/** Step 4: moves the sentinel file to the trash -- only when the one settings file of the account IS the sentinel. */
function userSettingsProbeCleanup() {
  userSettingsProbeRemoveTriggers_();
  PropertiesService.getDocumentProperties().deleteProperty(USER_SETTINGS_PROBE_RESULT_KEY_);
  const drive = userSettingsDrive_();
  const found = inspectPersonalReviewerSettingsWith_(drive);
  if (found.status === 'none') return userSettingsProbeSay_('Probe cleanup: nothing to remove', ['There is no settings file.']);
  if (!(found.status === 'ok' && found.token === USER_SETTINGS_PROBE_SENTINEL_)) {
    return userSettingsProbeSay_('Probe cleanup: NOT DONE', ['The settings of this account are not the probe sentinel (status "' + found.status + '"). Nothing was removed.']);
  }
  const cleared = clearPersonalReviewerTokenWith_(drive);
  return userSettingsProbeSay_('Probe cleanup: ' + (cleared.ok ? 'done' : 'FAILED'), ['settings files moved to the trash: ' + cleared.trashed + (cleared.ok ? '' : ' (' + cleared.reason + ')')]);
}
