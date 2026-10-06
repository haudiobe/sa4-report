/**
 * What changed in the production files after T-2026.10.8 -- and that nothing
 * else did.
 *
 * The work after T-2026.10.8 (not released yet):
 *   - no abstract is asked for a withdrawn TDoc        tests/withdrawn-no-abstract.test.js
 *   - the opening sentence of an ad-hoc report         tests/meeting-start-opening.test.js
 *   - the drafts folder of a new ad-hoc report         tests/adhoc-revisions-folder-setup.test.js
 *   - revision placement at the end of a build         tests/build-revision-placement.test.js
 *   - the personal Reviewer API token                  tests/personal-reviewer-token.test.js
 *
 * This file compares Code.js and template/ReportCreator.js with the release
 * tag template-release/T-2026.10.8:
 *   1. the functions of Code.js that changed are exactly sixteen, and each is
 *      the released function again once the lines named here are taken out;
 *   2. the new functions of Code.js are the fourteen of one section, and the
 *      rest of the file outside its functions is the released file;
 *   3. ReportCreator.js: which functions changed, which are new;
 *   4. the other files of the bundle are the released ones; the version
 *      line and one changelog entry are those of the candidate for
 *      T-2026.10.9 (Code.js 2.21.0).
 *
 * Without the tag in the checkout the comparisons are skipped, and said so.
 *
 * Run: node tests/after-10-8-ledger.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');

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

const ROOT = path.join(__dirname, '..');
const read = (file) => fs.readFileSync(path.join(ROOT, file), 'utf8').replace(/\r/g, '');
const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
const TAG = 'template-release/T-2026.10.8';
function gitShow(file) {
  try {
    return execFileSync('git', ['show', TAG + ':' + file], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    return null;
  }
}
const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, '')).filter((name, i, list) => list.indexOf(name) === i);
const changed = (now, old) => functionNames(old).filter((name) => functionSource(now, name) !== functionSource(old, name));
const added = (now, old) => functionNames(now).filter((name) => functionNames(old).indexOf(name) === -1).sort();
const removed = (now, old) => functionNames(old).filter((name) => functionNames(now).indexOf(name) === -1);
const outsideFunctions = (src) => functionNames(src).reduce((text, name) => text.replace(functionSource(src, name), ''), src);

const BAR = '// =========================================================\n';
const SECTION = BAR + '// AFTER T-2026.10.8 -- WITHDRAWN TDOCS AND ABSTRACTS; THE START OF AN AD-HOC MEETING\n';
const NEXT_SECTION = BAR + '// TDOC STATUS DROPDOWNS (T-2026.10.7)\n';

/**
 * Per changed function of Code.js: [what it has now, what stood there in
 * T-2026.10.8] -- every pair exactly once. Taking them back gives the
 * released function, byte for byte.
 */
const CODE_CHANGES = {
  // No abstract for a withdrawn TDoc: the guard, and the status of the TDoc list for it.
  fetchAndAddAbstract_: [
    ['function fetchAndAddAbstract_(table, tdocNumber, context, portalStatus) {\n', 'function fetchAndAddAbstract_(table, tdocNumber, context) {\n'],
    ['  // Nothing is asked for a withdrawn TDoc (abstractFetchBlockedBy_()): this\n' +
     '  // is checked here, before the token, the "no summary" cache and the\n' +
     '  // request, so that no caller can ask for one. `portalStatus` is the status\n' +
     '  // of the TDoc in the TDoc list, from a caller that has it. Returns false\n' +
     '  // when the TDoc is withdrawn and nothing was done -- and nothing otherwise.\n' +
     '  const withdrawnBy = abstractFetchBlockedBy_(table, portalStatus);\n' +
     '  if (withdrawnBy) {\n' +
     "    perfCount_('abstracts not asked for: TDoc withdrawn');\n" +
     "    Logger.log(`No abstract is asked for ${tdocNumber}: withdrawn (${withdrawnBy === 'portal' ? 'TDoc list' : 'Status in the report'})`);\n" +
     '    return false;\n' +
     '  }\n', '']
  ],
  addAbstractsForTables_: [
    ['function addAbstractsForTables_(body, context, leaveOut, portalStatuses) {\n', 'function addAbstractsForTables_(body, context, leaveOut) {\n'],
    ['  // `portalStatuses`, { number: status }, is the TDoc list of an update\n' +
     '  // (tdocListStatuses_()). A withdrawn TDoc -- by the Status in the report,\n' +
     '  // or by that list -- is no candidate: abstractFetchBlockedBy_().\n', ''],
    ["    const portalStatus = portalStatuses && Object.prototype.hasOwnProperty.call(portalStatuses, tdocNumber) ? portalStatuses[tdocNumber] : '';\n" +
     '\n' +
     '    // false: the TDoc is withdrawn and nothing was asked -- no candidate.\n' +
     "    const asked = perfTimedAccum_('Reviewer API abstract fetch (fetchAndAddAbstract_, accumulated)', () => fetchAndAddAbstract_(table, parsedTdoc.raw, context, portalStatus));\n" +
     '    if (asked !== false) candidatesProcessed++;\n',
     '\n' +
     "    perfTimedAccum_('Reviewer API abstract fetch (fetchAndAddAbstract_, accumulated)', () => fetchAndAddAbstract_(table, parsedTdoc.raw, context));\n" +
     '    candidatesProcessed++;\n']
  ],
  continuousUpdateCore_: [
    ['      // TDoc that was asked for above, in this run. A withdrawn TDoc is no\n' +
     "      // candidate either (the list's statuses are passed for that).\n", '      // TDoc that was asked for above, in this run.\n'],
    ['Object.assign(tdocsNotUploadedYet_(allTdocs), abstractAskedThisRun), tdocListStatuses_(allTdocs)));\n', 'Object.assign(tdocsNotUploadedYet_(allTdocs), abstractAskedThisRun)));\n']
  ],
  refreshUploadedTdocMetadata_: [
    ['    // A withdrawn TDoc gets its link and nothing else: no abstract is asked\n' +
     '    // for, and its "no summary" answer is not touched.\n' +
     '    const portalStatus = tdocListStatus_(tdocData);\n' +
     '    if (abstractFetchBlockedBy_(table, portalStatus)) {\n' +
     "      perfCount_('abstracts not asked for: TDoc withdrawn');\n" +
     '      Logger.log(`Upload of ${tdocNumber}: withdrawn, no abstract is asked for`);\n' +
     '      return done;\n' +
     '    }\n', ''],
    ['    done.abstractAttempted = fetchAndAddAbstract_(table, parsed.raw, context, portalStatus) !== false;\n',
     '    done.abstractAttempted = true;\n    fetchAndAddAbstract_(table, parsed.raw, context);\n']
  ],
  completeInsertedUploadedTdoc_: [
    ['    return fetchAndAddAbstract_(table, parsed.raw, context, tdocListStatus_(tdocData)) !== false;\n', '    fetchAndAddAbstract_(table, parsed.raw, context);\n    return true;\n']
  ],
  // The start of an ad-hoc meeting.
  computeResolvedMeetingPreview_: [
    ["    // The Portal's time zone of the start, as it is given (see computeMeetingTimeZoneLabel_()).\n" +
     '    startTimeZoneRaw: resolvedMeeting && resolvedMeeting.timeZone ? resolvedMeeting.timeZone : null,\n', '']
  ],
  persistConfigurationSettings_: [
    ['  // The start of the meeting (time of day, time zone), from the Portal. It\n' +
     '  // comes with the configuration in one of two ways, and both store the\n' +
     '  // same thing (storeMeetingStart_()):\n' +
     '  //   - the creator of a report sends the two values (meetingStartTime,\n' +
     '  //     meetingTimeZone);\n' +
     '  //   - Configure Meeting sends what its Resolve got from the Portal for the\n' +
     '  //     meeting that is saved (resolvedMeetingStart) -- the values are worked\n' +
     '  //     out here, by the functions the creator uses.\n' +
     '  // A save that carries neither leaves what is stored alone.\n' +
     "  const meetingStartTime = String(config.meetingStartTime || '').trim();\n" +
     "  const meetingTimeZone = String(config.meetingTimeZone || '').trim();\n" +
     '  if (meetingStartTime || meetingTimeZone) {\n' +
     '    storeMeetingStart_(docProps, meetingStartTime, meetingTimeZone);\n' +
     '  } else {\n' +
     "    const resolvedStart = meetingStartFromResolved_(config.resolvedMeetingStart, docProps.getProperty('MEETING_ID'), docProps.getProperty('MEETING_TYPE'), docProps.getProperty('MEETING_DATE'));\n" +
     '    if (resolvedStart) storeMeetingStart_(docProps, resolvedStart.time, resolvedStart.zone);\n' +
     '  }\n' +
     '\n', '']
  ],
  // Configure Meeting: the start its Resolve found goes with its Save; the Save hands the personal Reviewer token to the
  // template (before it stores anything) and says what became of it and of the drafts folder.
  saveConfigurationSettings: [
    ["    // Template runtime only: what a save does besides (ReportCreator.js) -- the personal Reviewer token, taken out of\n    // the configuration before it is stored, and the drafts folder afterwards. Nothing elsewhere.\n    const templateSave = templateRuntimeRelease_() ? beforeTemplateConfigurationSaved_(config) : null;\n    persistConfigurationSettings_(templateSave ? templateSave.config : config);\n    return templateSave ? afterTemplateConfigurationSaved_(templateSave) : undefined;\n",
     "    persistConfigurationSettings_(config);\n    return;\n"]
  ],
  configureMeetingSettings: [
    ['      // The start of the meeting Resolve found, as the Portal gives it, and\n' +
     '      // the meeting it is of. Sent with Save only for that same meeting id.\n' +
     '      let resolvedMeetingStart = null;\n', ''],
    ["            resolvedMeetingStart = { meetingId: String(result.resolved.id), startDate: result.preview.startDateRaw || '', timeZone: result.preview.startTimeZoneRaw || '' };\n", ''],
    ['        // The start Resolve found goes with the save of the meeting it was found for (see persistConfigurationSettings_()).\n' +
     "        if (resolvedMeetingStart && resolvedMeetingStart.meetingId === String(config.meetingId || '').trim()) {\n" +
     '          config.resolvedMeetingStart = resolvedMeetingStart;\n' +
     '        }\n', ''],
    ['          .withSuccessHandler((saved) => {\n            // Saved is not the same as ready to build.\n', '          .withSuccessHandler(() => {\n            // Saved is not the same as ready to build.\n'],
    ["              : '') + (saved && saved.note ? '\\\\n\\\\n' + saved.note : ''));\n", "              : ''));\n"],
    // The personal Reviewer token: its part of the dialog comes from the template, and the page says what is to become of it.
    ["  const tokenConfigured = Boolean(resolveReviewerApiToken_(true).token);\n  // A template report: the token of this report and the personal one (templateReviewerTokenDialogParts_(), ReportCreator.js).\n  // Words and controls only; no token goes into the dialog. Everywhere else the dialog is as it was.\n  const tokenUi = templateRuntimeRelease_() ? templateReviewerTokenDialogParts_() : null;\n",
     "  const tokenConfigured = Boolean(PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN'));\n"],
    ["      <div class=\"hint\" id=\"tokenStatus\">${tokenUi ? esc(tokenUi.status) : (tokenConfigured ? 'Reviewer API token configured' : 'No Reviewer API token configured')}</div>\n",
     "      <div class=\"hint\" id=\"tokenStatus\">${tokenConfigured ? 'Reviewer API token configured' : 'No Reviewer API token configured'}</div>\n"],
    ["        ${tokenUi ? tokenUi.fieldsHtml : `<label>Replace Reviewer API Token:</label>\n",
     "        <label>Replace Reviewer API Token:</label>\n"],
    ["The saved token is never displayed.</div>`}\n",
     "The saved token is never displayed.</div>\n"],
    ["        // A template report has controls for a personal token as well (see beforeTemplateConfigurationSaved_(),\n        // ReportCreator.js). Where they do not exist, nothing here is ticked and all is as it was.\n        const ticked = function (id) { const el = document.getElementById(id); return !!(el && el.checked); };\n        const tokenForAllReports = ticked('apiTokenTargetPersonal');\n",
     ""],
    ["        } else if (newToken && newToken.trim() && !tokenForAllReports) {\n",
     "        } else if (newToken && newToken.trim()) {\n"],
    ["        // The personal token: one thing at a time -- a new one, the one of this report, or its removal.\n        const personalTokenWishes = [ticked('clearPersonalApiToken') ? 'clear' : '', ticked('promoteApiToken') ? 'promote' : '',\n          tokenForAllReports && newToken && newToken.trim() ? 'replace' : ''].filter(Boolean);\n        if (personalTokenWishes.length > 1) {\n          alert('Choose one thing for your personal token: a new token, the token of this report, or removing it.');\n          return;\n        }\n        if (personalTokenWishes.length === 1) {\n          config.personalTokenAction = personalTokenWishes[0];\n          if (personalTokenWishes[0] === 'replace') config.personalToken = newToken;\n        }\n",
     ""]
  ],
  // The personal Reviewer token: one resolver, asked only when the Reviewer is; every reader of the token goes through it.
  getReportConfig_: [
    ["  // ========================================\n  const REVIEWER_API_BASE = ",
     "  // ========================================\n  const REVIEWER_API_TOKEN = getReviewerApiTokenForRun_(); // TEMPLATE-002C: one read per execution\n  const REVIEWER_API_BASE = "],
    ["  const reportConfig = {\n    // Meeting identification\n",
     "  return {\n    // Meeting identification\n"],
    ["    // API\n    REVIEWER_API_BASE,\n",
     "    // API\n    REVIEWER_API_TOKEN,\n    REVIEWER_API_BASE,\n"],
    ["  // The token in use is still to be had as REVIEWER_API_TOKEN, but it is no longer looked up for every configuration:\n  // only when it is asked for (resolveReviewerApiToken_()). It is not enumerable, so a copy or a print of the\n  // configuration neither looks it up nor contains it.\n  Object.defineProperty(reportConfig, 'REVIEWER_API_TOKEN', { enumerable: false, get: function () { return getReviewerApiTokenForRun_(); } });\n  return reportConfig;\n",
     ""]
  ],
  getReviewerApiTokenForRun_: [
    ["  return resolveReviewerApiToken_().token;\n",
     "  if (!REVIEWER_TOKEN_RUN_STATE_.read) {\n    REVIEWER_TOKEN_RUN_STATE_.token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN') || '';\n    REVIEWER_TOKEN_RUN_STATE_.read = true;\n  }\n  return REVIEWER_TOKEN_RUN_STATE_.token;\n"]
  ],
  resetReviewerTokenRunState_: [
    ["  REVIEWER_TOKEN_RUN_STATE_.resolved = null;\n",
     ""]
  ],
  testAllConnections: [
    ["    const token = resolveReviewerApiToken_().token;\n",
     "    const token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN');\n"]
  ],
  testReviewerApi: [
    ["  const token = resolveReviewerApiToken_().token;\n",
     "  const token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN');\n"]
  ],
  validateConfiguration: [
    ["  // The token in use (of this report, or -- in a template report -- the personal one)\n  const apiToken = resolveReviewerApiToken_().token;\n",
     "  // Check script properties\n  const scriptProps = PropertiesService.getScriptProperties();\n  const apiToken = scriptProps.getProperty('REVIEWER_API_TOKEN');\n"]
  ],
  // The opening sentence, and the revision placement at the end of the build.
  buildSkeletonWithTdocTables: [
    ['      // The chair stays a literal, editable placeholder to fill in by hand.\n',
     '      // No CHAIR_NAME/START_TIME property is introduced -- those remain\n      // literal, editable placeholders for the chair to fill in by hand.\n'],
    ['      // instead of inventing or assuming any specific date. The start time\n' +
     '      // and the time zone are the ones the Portal gave when the report was\n' +
     '      // created (getMeetingStartForOpening_()); without them the sentence\n' +
     '      // has "<start>" and "<time zone>" -- no zone is assumed.\n', '      // instead of inventing or assuming any specific date.\n'],
    ["      body.appendParagraph(`${openingSubSection} Opening of the session`).setHeading(DocumentApp.ParagraphHeading.HEADING3);\n" +
     '      // Ad-hoc opening (stage E)',
     "      body.appendParagraph(`${openingSubSection} Opening of the session`).setHeading(DocumentApp.ParagraphHeading.HEADING3);\n" +
     "      const meetingDateText = (cfg.MEETING_DATE || '').trim() || '<meeting date>';\n" +
     '      // Ad-hoc opening (stage E)'],
    ['        const meetingStart = getMeetingStartForOpening_();\n' +
     '        body.appendParagraph(buildMeetingOpeningSentence_(cfg.MEETING_DATE, meetingStart.time, meetingStart.zone));\n',
     '        body.appendParagraph(`<Chair> opens the session on ${meetingDateText} at <start> CEST.`);\n'],
    ['  // Revision placement: the loops above put a revision below the document\n' +
     '  // it revises only when both are under the same agenda item\n' +
     '  // (orderTdocsByRevision_()). The pass an update ends with is run here too,\n' +
     '  // on the TDoc list this build already downloaded, so a new report has\n' +
     '  // every revision directly below its parent -- also across agenda items --\n' +
     '  // and does not wait for its first update to get there.\n' +
     "  let revisionPlacementNote = '';\n" +
     '  try {\n' +
     '    const placed = rearrangeRevisionTables_(cfg, { all: { tdocs: allTdocs } });\n' +
     '    Logger.log(`Revisions: ${placed.moved} moved, ${placed.dispositions} disposition(s) filled`);\n' +
     '  } catch (e) {\n' +
     "    Logger.log('Revision placement failed: ' + e.message);\n" +
     "    revisionPlacementNote = '\\n\\n⚠️ Revisions could not be placed below the documents they revise (' + e.message + '). The next update of the report places them.';\n" +
     '  }\n' +
     '\n', ''],
    ['  let reallocationRestoreNote = revisionPlacementNote;\n', "  let reallocationRestoreNote = '';\n"]
  ]
};
const SECTION_FUNCTIONS = ['isWithdrawnStatus_', 'abstractFetchBlockedBy_', 'tdocListStatus_', 'tdocListStatuses_', 'resolveReviewerApiToken_', 'computeMeetingStartTimeFromStartDate_', 'computeMeetingTimeZoneLabel_',
  'isValidMeetingStartTime_', 'isValidMeetingTimeZoneLabel_', 'buildMeetingOpeningSentence_', 'meetingStartBasis_', 'storeMeetingStart_', 'meetingStartFromResolved_', 'getMeetingStartForOpening_'];

const OLD_CODE = gitShow('Code.js');
const OLD_CREATOR = gitShow('template/ReportCreator.js');

console.log('1. Code.js: the functions that changed');
if (!OLD_CODE) {
  console.log('  note: ' + TAG + ' is not available in this checkout; the comparisons with it are skipped.');
} else {
  check('exactly these sixteen functions of T-2026.10.8 changed', changed(CODE, OLD_CODE).sort(), Object.keys(CODE_CHANGES).sort());
  Object.keys(CODE_CHANGES).forEach((name) => {
    const now = functionSource(CODE, name);
    const counts = CODE_CHANGES[name].map((pair) => now.split(pair[0]).length - 1);
    const back = CODE_CHANGES[name].reduce((text, pair) => text.replace(pair[0], () => pair[1]), now);
    check(name + '(): the lines named here are in it once each, and without them it is the released function', [counts, back === functionSource(OLD_CODE, name)], [CODE_CHANGES[name].map(() => 1), true]);
  });
  check('no function of T-2026.10.8 was removed', removed(CODE, OLD_CODE), []);
}

console.log('2. Code.js: the new functions and the rest of the file');
{
  const at = CODE.indexOf(SECTION);
  const end = CODE.indexOf(NEXT_SECTION);
  check('the work after T-2026.10.8 is one section, directly before the status dropdowns', [at !== -1, CODE.indexOf(SECTION, at + 1), at < end], [true, -1, true]);
  const section = CODE.slice(at, end);
  check('it has these fourteen functions, in this order', functionNames(section), SECTION_FUNCTIONS);
  if (OLD_CODE) {
    check('they are the new functions of Code.js, all of them', added(CODE, OLD_CODE), SECTION_FUNCTIONS.slice().sort());
    // The header: the version line and the changelog entry of 2.21.0 are new; the rest of it is unchanged.
    const withoutSection = CODE.slice(0, at) + CODE.slice(end);
    const withoutRelease = withoutSection.replace(/ \* 2\.21\.0 \(2026-10-06\)\n[\s\S]*? \* 2\.20\.0 \(2026-10-06\)\n/, ' * 2.20.0 (2026-10-06)\n').replace(' * Version: 2.21.0 (2026-10-06)\n', ' * Version: 2.20.0 (2026-10-06)\n');
    check('outside its functions, that section, the version line and the changelog entry of 2.21.0, Code.js is the released file: no constant and no comment differs',
      [outsideFunctions(withoutRelease) === outsideFunctions(OLD_CODE), withoutRelease === withoutSection], [true, false]);
  }
  check('the section declares three constants, the keys of the stored start, and nothing else at top level',
    (section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm) || []), ['const MEETING_START_TIME_KEY_', 'const MEETING_TIME_ZONE_KEY_', 'const MEETING_START_BASIS_KEY_']);
}

console.log('3. template/ReportCreator.js');
if (!OLD_CREATOR) {
  console.log('  note: ' + TAG + ' is not available in this checkout; the comparisons with it are skipped.');
} else {
  // The setup information (payload and its validation), the first run and what it reports, About This Report.
  check('these functions of T-2026.10.8 changed, and no other', changed(CREATOR, OLD_CREATOR),
    ['buildReportBootstrapPayload_', 'validateBootstrapPayload_', 'finishReportSetupWith_', 'describeTemplateRuntime_', 'liveTemplateDeps_', 'ensureReportBootstrapped_', 'finishReportSetup', 'showTemplateInfo']);
  // The drafts folder (five), and the personal Reviewer token (sixteen).
  check('twenty-one functions are new; none was removed', [added(CREATOR, OLD_CREATOR), removed(CREATOR, OLD_CREATOR)],
    [['afterTemplateConfigurationSaved_', 'confirmAdhocRevisionsFolderWith_', 'currentRevisionsCandidate_', 'describeRevisionsFolderOutcome_', 'retryAdhocRevisionsFolderWith_',
      'applyPersonalReviewerTokenPlanSafely_', 'applyPersonalReviewerTokenPlanWith_', 'beforeTemplateConfigurationSaved_', 'classifyUserSettingsFile_', 'clearPersonalReviewerTokenWith_', 'describeReviewerTokenStatus_',
      'inspectPersonalReviewerSettingsWith_', 'isUsableReviewerTokenValue_', 'isVerifiedUserSettingsFile_', 'parseUserSettingsText_', 'personalTokenProblemForLog_', 'readPersonalReviewerTokenSafely_', 'serializeUserSettings_',
      'templateReviewerTokenDialogParts_', 'userSettingsDrive_', 'writePersonalReviewerTokenWith_'].sort(), []]);
  const constant = (src, name) => { const start = src.indexOf('\nvar ' + name + ' = '); return src.slice(start + 1, src.indexOf(';\n', start) + 1); };
  const constants = (src) => (src.match(/^var [A-Za-z0-9_$]+/gm) || []).map((m) => m.slice(4));
  check('of the constants it had, two changed: the keys of the setup information (two more) and the state keys (one more)',
    [constants(OLD_CREATOR).filter((name) => constant(CREATOR, name) !== constant(OLD_CREATOR, name)), constants(OLD_CREATOR).filter((name) => constants(CREATOR).indexOf(name) === -1)], [['TEMPLATE_BOOTSTRAP_CONFIG_KEYS_', 'TEMPLATE_STATE_KEYS_'], []]);
  check('the new constants are those of the personal settings file',
    constants(CREATOR).filter((name) => constants(OLD_CREATOR).indexOf(name) === -1),
    ['USER_SETTINGS_SCHEMA_', 'USER_SETTINGS_MARKER_KEY_', 'USER_SETTINGS_MARKER_VALUE_', 'USER_SETTINGS_FILE_NAME_', 'USER_SETTINGS_MIME_', 'USER_SETTINGS_MAX_BYTES_', 'USER_SETTINGS_TOKEN_MAX_CHARS_', 'USER_SETTINGS_QUERY_', 'USER_SETTINGS_FILE_FIELDS_']);
  check('the menus are the released ones', ['buildTemplateMasterMenu_', 'buildTemplateReportMenu_', 'addTemplateReportMenuHead_', 'addTemplateReportMenuTail_'].filter((name) => functionSource(CREATOR, name) !== functionSource(OLD_CREATOR, name)), []);
}

console.log('4. the rest of the bundle, and the version');
{
  const others = ['HyperLink.js', 'appsscript.json', 'colab_notebook.html', 'colab_notebook_shared.html', 'tools/template-release.js', '.claspignore'];
  if (OLD_CODE) check('the other files of the bundle, the release tool and the push filter are the released ones', others.filter((file) => read(file) !== gitShow(file)), []);
  check('Code.js is version 2.21.0, and the first changelog entry is that of 2.21.0, directly above the one of 2.20.0',
    [(CODE.match(/^ \* Version: (\d+\.\d+\.\d+) \((\d{4}-\d{2}-\d{2})\)/m) || []).slice(1), (CODE.match(/^ \* (\d+\.\d+\.\d+) \(\d{4}-\d{2}-\d{2}\)$/gm) || []).slice(0, 2)], [['2.21.0', '2026-10-06'], [' * 2.21.0 (2026-10-06)', ' * 2.20.0 (2026-10-06)']]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll after-T-2026.10.8 ledger checks passed.');
process.exitCode = failures ? 1 : 0;
