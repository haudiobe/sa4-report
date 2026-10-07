/**
 * TEMPLATE-003 stage 1 -- the "SA4 Report" menu of the template runtime.
 *
 *   1. The exact menu trees: master template, report before setup, report
 *      after setup. No LEGACY submenu, none of the hidden operations.
 *   2. What was hidden is still in the code; nothing was removed.
 *   3. The CENTRAL add-on / Legacy menu (no Release.js) is exactly what it
 *      was in release T-2026.10.3.
 *   4. A: "Build Report from Scratch" is runFullReportBuild(); its one
 *      question warns that the content is replaced; no UI call afterwards.
 *   5. B: "Update Report Now" runs the complete update and fails visibly;
 *      the trigger's continuousUpdate() behaves as before.
 *   6. The partial updates and Automatic Updates call their existing functions.
 *
 * Run: node tests/template003-menu.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
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
const thrown = (fn) => { try { fn(); return null; } catch (e) { return e.message; } };

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');

const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const RELEASE = {
  releaseId: 'T-2026.10.4', flavor: 'template', codeVersion: (CODE.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1],
  gitCommit: 'abcdef0000000000000000000000000000000000', gitTag: 'template-release/T-2026.10.4', templateDocumentId: TEMPLATE_ID
};
/** The accepted live release: the reference for "unchanged". */
const BASELINE = 'template-release/T-2026.10.3';
const SET_UP = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'adhoc', MEETING_ID: '86178', REPORT_SUFFIX: '6G' };

// ------------------------------------------------------------------ helpers

function fakeUi(events) {
  const ui = { menus: [], alerts: [], dialogs: [], alertResponse: 'YES', ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES', NO: 'NO', OK: 'OK' } };
  ui.createMenu = (name) => {
    const m = { name, entries: [] };
    m.addItem = (label, fn) => { m.entries.push({ label, fn }); return m; };
    m.addSeparator = () => { m.entries.push({ sep: true }); return m; };
    m.addSubMenu = (sub) => { m.entries.push({ sub }); return m; };
    m.addToUi = () => { ui.menus.push(m); };
    return m;
  };
  ui.alert = (...args) => { ui.alerts.push(args); if (events) events.push('ALERT ' + args[0]); return ui.alertResponse; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title, html: out.html }); };
  return ui;
}

/** The installed menus as indented lines: "label -> function", "---" for a separator. */
function tree(ui) {
  const out = [];
  const walk = (m, depth) => {
    out.push('  '.repeat(depth) + m.name);
    m.entries.forEach((e) => {
      if (e.sub) walk(e.sub, depth + 1);
      else out.push('  '.repeat(depth + 1) + (e.sep ? '---' : e.label + ' -> ' + e.fn));
    });
  };
  ui.menus.forEach((m) => walk(m, 0));
  return out;
}
const targets = (ui) => tree(ui).filter((l) => / -> /.test(l)).map((l) => l.split(' -> ')[1]);

/** A template-runtime sandbox bound to one document. opts: { docId, props, release (false = none) }. */
function runtime(opts) {
  const o = opts || {};
  const loaded = loadTemplateRuntime({ release: o.release === false ? null : RELEASE, documentProperties: o.props || {} });
  const s = loaded.sandbox;
  const events = [];
  const logs = [];
  const ui = fakeUi(events);
  const body = makeFakeDocumentBody(s);
  const docId = o.docId || REPORT_ID;
  s.Logger = { log: (m) => logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => docId, getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  return { s, ui, events, logs, body, docProps: loaded.docProps };
}

/** A git object as text, or null when this checkout cannot provide it. */
function gitShow(spec) {
  try {
    return execFileSync('git', ['show', spec], { cwd: path.join(__dirname, '..'), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, '');
  } catch (e) {
    return null;
  }
}
const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, ''));
/** The source of one top-level function (up to its closing brace in column 0). */
function functionSource(src, name) {
  const start = src.indexOf('\nfunction ' + name + '(');
  if (start === -1) return null;
  return src.slice(start + 1, src.indexOf('\n}\n', start) + 2);
}

const OLD_CODE = gitShow(BASELINE + ':Code.js');
const OLD_CREATOR = gitShow(BASELINE + ':template/ReportCreator.js');
const haveBaseline = !!(OLD_CODE && OLD_CREATOR);
if (!haveBaseline) console.log('  note: ' + BASELINE + ' is not available in this checkout; the comparisons with it are skipped.');

// ================================================================ 1. menu trees

const MASTER_TREE = [
  'SA4 Report',
  '  🆕 Create New SA4 Report -> showCreateReportDialog',
  '  ---',
  '  ℹ️ Template Release Info -> showTemplateInfo'
];
const REPORT_TREE_AFTER_SETUP = [
  'SA4 Report',
  '  📄 Report',
  '    Build Report from Scratch… -> runFullReportBuild',
  '    Update Report Now -> updateReportNow',
  '    ---',
  '    Update E-mail Discussions -> collectEmailDiscussionOnly',
  '    Update TDoc Revisions -> collectRevisionsOnly',
  '    Update Abstracts -> addAbstractsOnly',
  '    ---',
  '    Report Status Summary -> analyzeReportStatus',
  '    Convert Status Fields to Dropdowns… -> convertStatusFieldsToDropdowns',
  '  💬 Prepare Discussion E-mails… -> prepareTdocDiscussionEmails',
  '  🔀 Document Reallocation',
  '    Add or Change a Reallocation… -> addDocumentReallocation',
  '    Show Reallocations -> viewAllReallocations',
  '    Apply Reallocations to Report… -> applyDocumentReallocations',
  '    Remove All Reallocations… -> clearAllReallocations',
  '  🔄 Automatic Updates… -> manageTriggers',
  '  ⚙️ Configure Meeting… -> configureMeetingSettings',
  '  ---',
  '  🛠 Advanced and Repair',
  '    Check Connections -> testAllConnections',
  '    Format Report -> removeRowHeightAndSpacing',
  '    Remove Duplicate E-mail Entries… -> removeDuplicateEmailEntries',
  '    Remove Wrong E-mail Matches… -> cleanUpWrongEmailDiscussions',
  '    Clear Collection Caches… -> clearAllCaches',
  '  ℹ️ About This Report -> showTemplateInfo'
];
// Ad-hoc sessions (stage A): an ad-hoc report has one more submenu, after
// Document Reallocation. A report that is not ad-hoc keeps the tree above.
const SESSIONS_SUBMENU = ['  🗓 Sessions and Attendance', '    Configure Sessions… -> configureAdhocSessions',
  '    Import Teams Attendance… -> importTeamsAttendance', '    Refresh Attendance Section -> refreshAdhocAttendanceSection',
  '    Assign TDoc Sessions… -> assignAdhocTdocSessions', '    Edit Opening Details… -> editAdhocOpeningDetails', '    Post-meeting Statistics… -> showAdhocSessionStatistics'];
const REALLOCATION_END = REPORT_TREE_AFTER_SETUP.indexOf('    Remove All Reallocations… -> clearAllReallocations') + 1;
const REPORT_TREE_ADHOC = REPORT_TREE_AFTER_SETUP.slice(0, REALLOCATION_END).concat(SESSIONS_SUBMENU, REPORT_TREE_AFTER_SETUP.slice(REALLOCATION_END));
const SET_UP_MAIN = { SA4_BOOTSTRAP_STATE: 'done', MEETING_TYPE: 'main', MEETING_ID: '60777', REPORT_SUFFIX: '6G' };
const REPORT_TREE_BEFORE_SETUP = [REPORT_TREE_AFTER_SETUP[0], '  🚀 Finish Report Setup -> finishReportSetup', '  ---'].concat(REPORT_TREE_AFTER_SETUP.slice(1));

/** Operations that left the template report menu and must stay in the code. */
const HIDDEN = ['buildInitialReport', 'updateAll', 'updateReportIncremental', 'buildSkeletonWithTdocTables', 'parseAgendaDocument',
  'autoCreateReportStructure', 'rearrangeRevisionTables', 'createConfigurationTables', 'testTdocListUrl', 'testReviewerApi',
  'testEmailFeeds', 'testRevisionsFolder', 'validateConfiguration', 'rewritePortalLinksInDoc_', 'fixColumnWidths', 'continuousUpdate'];

console.log('1. master template menu');
{
  const r = runtime({ docId: TEMPLATE_ID });
  r.s.onOpen();
  check('exact tree', tree(r.ui), MASTER_TREE);
  check('one menu, no report operation', [r.ui.menus.length, targets(r.ui)], [1, ['showCreateReportDialog', 'showTemplateInfo']]);
}

console.log('1. report menu before setup');
let reportTargets;
{
  const r = runtime({ docId: REPORT_ID });
  r.s.onOpen();
  check('exact tree', tree(r.ui), REPORT_TREE_BEFORE_SETUP);
  check('one menu', r.ui.menus.length, 1);
}

console.log('1. report menu after setup');
{
  const r = runtime({ docId: REPORT_ID, props: SET_UP });
  r.s.onOpen();
  const t = tree(r.ui);
  reportTargets = targets(r.ui);
  check('exact tree (ad-hoc report: with the Sessions and Attendance submenu)', t, REPORT_TREE_ADHOC);
  check('the menu is not called Scripts', [r.ui.menus[0].name, t.some((l) => /Scripts/.test(l))], ['SA4 Report', false]);
  check('no LEGACY submenu and no Legacy item', t.filter((l) => /legacy/i.test(l)), []);
  check('none of the hidden operations is offered', HIDDEN.filter((f) => reportTargets.indexOf(f) !== -1), []);
  check('Configure Meeting is offered once', reportTargets.filter((f) => f === 'configureMeetingSettings').length, 1);
  check('every item calls an existing public function (a menu cannot call a private one)',
    reportTargets.filter((f) => typeof r.s[f] !== 'function' || /_$/.test(f)), []);
  check('no label is used twice', t.filter((l, i) => / -> /.test(l) && t.indexOf(l) !== i), []);
  check('ad-hoc report: 26 items after setup, 4 submenus', [reportTargets.length, r.ui.menus[0].entries.filter((e) => e.sub).map((e) => e.sub.name)],
    [26, ['📄 Report', '🔀 Document Reallocation', '🗓 Sessions and Attendance', '🛠 Advanced and Repair']]);
  const main = runtime({ docId: REPORT_ID, props: SET_UP_MAIN });
  main.s.onOpen();
  check('main-meeting report: exact tree, without the Sessions and Attendance submenu', tree(main.ui), REPORT_TREE_AFTER_SETUP);
  check('main-meeting report: 20 items after setup, 3 submenus', [targets(main.ui).length, main.ui.menus[0].entries.filter((e) => e.sub).map((e) => e.sub.name)],
    [20, ['📄 Report', '🔀 Document Reallocation', '🛠 Advanced and Repair']]);
  check('a document configured by hand (state "manual") gets the same menu',
    (() => { const m = runtime({ docId: REPORT_ID, props: { SA4_BOOTSTRAP_STATE: 'manual' } }); m.s.onOpen(); return tree(m.ui); })(), REPORT_TREE_AFTER_SETUP);
}

console.log('1. texts of the template runtime name the new menu');
{
  const r = runtime({ docId: REPORT_ID, props: SET_UP });
  check('no template text mentions the old menu or its submenus', /Scripts⚠️|Scripts menu|INITIAL SETUP|REPORT OPERATIONS|Run Full Report Build\./.test(CREATOR.replace(/^\s*(\/\/|\*).*$/gm, '')), false);
  const about = r.s.describeTemplateRuntime_({ release: RELEASE, documentId: REPORT_ID, scriptId: 'S', bootstrapState: null, meetingId: null, continuousInterval: '30 minutes' }).join('\n');
  check('About This Report: the build by its new name, "Automatic updates"', [/Build Report from Scratch/.test(about), /Automatic updates: every 30 minutes/.test(about), /Continuous Update/.test(about)], [true, true, false]);
  // A document without the creator's setup information (an empty file description).
  const plainDoc = runtime({ docId: REPORT_ID });
  plainDoc.s.DriveApp = { getFileById: () => ({ getDescription: () => '', setDescription: () => {} }) };
  plainDoc.s.ScriptApp = { getScriptId: () => 'SCRIPT' };
  plainDoc.s.finishReportSetup();
  check('Finish Report Setup without setup information points to Configure Meeting in the new menu',
    /Use SA4 Report > ⚙️ Configure Meeting…\./.test(plainDoc.ui.alerts[plainDoc.ui.alerts.length - 1][1]), true);
  check('its "ready" result names the build in the new menu (source)', CREATOR.indexOf("'✅ Ready to build: ' + TEMPLATE_MENU_NAME_ + ' > 📄 Report > Build Report from Scratch….'") !== -1, true);
  check('the setup note written into a new report names the new menu (source)', CREATOR.indexOf("'. Open it and run ' + TEMPLATE_MENU_NAME_ + ' > 🚀 Finish Report Setup.\\n'") !== -1, true);
}

// ================================================================ 2. nothing removed

console.log('2. hidden operations are still in the code');
{
  const r = runtime({ docId: REPORT_ID, props: SET_UP });
  check('every hidden operation is still a function', HIDDEN.filter((f) => typeof r.s[f] !== 'function'), []);
  const addon = loadCode().sandbox;
  check('and in the add-on runtime (Code.js alone) too', HIDDEN.filter((f) => typeof addon[f] !== 'function'), []);
  if (haveBaseline) {
    const before = functionNames(OLD_CODE).concat(functionNames(OLD_CREATOR));
    const now = functionNames(CODE).concat(functionNames(CREATOR));
    check('no function of T-2026.10.3 was removed', before.filter((f) => now.indexOf(f) === -1), []);
    // TEMPLATE-003 added four functions, all in ReportCreator.js. The ad-hoc
    // session configuration (stage A) added its own, listed here by name.
    const SESSION_FUNCTIONS_CODE = ['adhocSessionLabel_', 'adhocSessionsEnabled_', 'applyAdhocSessionsEdit_', 'buildAdhocSessionsDialogModel_',
      'compareAdhocSessions_', 'configureAdhocSessions', 'defaultAdhocSessionLabel_', 'getAdhocSessions_', 'isAdhocMeetingForSessions_',
      'isValidAdhocSessionDate_', 'isValidAdhocSessionTime_', 'parseAdhocSessionsProperty_', 'salvageAdhocSessionNextId_',
      'saveAdhocSessionsConfiguration', 'saveAdhocSessionsWith_', 'serializeAdhocSessions_', 'sortAdhocSessions_', 'validateAdhocSessions_'];
    const SESSION_FUNCTIONS_CREATOR = ['isAdhocReportForMenu_'];
    // The Teams attendance parser (stage C): pure functions, called by nothing yet.
    const TEAMS_PARSER_FUNCTIONS_CODE = ['decodeTeamsAttendanceBytes_', 'formatTeamsDuration_', 'normalizeTeamsDisplayName_', 'normalizeTeamsEmail_',
      'parseTeamsAttendanceReport_', 'parseTeamsDuration_', 'readTeamsDateTimeParts_', 'resolveTeamsDateOrder_', 'scanDelimitedRows_', 'splitDelimitedRows_',
      'teamsAttendeeKey_', 'teamsDateTimeFromParts_', 'teamsSecondsBetween_', 'teamsTwoDigits_'];
    // Attendance import, persistence and rendering (stage D).
    const ATTENDANCE_FUNCTIONS_CODE = ['adhocAttendanceChunkKey_', 'adhocAttendanceSessionIds_', 'adhocAttendanceStore_', 'adhocAttendeeCompanyKey_', 'adhocAttendeeTablesToHarvest_', 'countAdhocAttendeeTablesElsewhere_', 'formatAdhocAttendeeTable_', 'adhocLongDate_',
      'adhocTextHash_', 'ambiguousAdhocAttendeeNames_', 'beginAdhocAttendanceRebuild_', 'buildAdhocAttendanceBlocks_', 'buildTeamsAttendanceDialogModel_',
      'compactAdhocAttendance_', 'confirmTeamsAttendanceImport', 'effectiveAdhocAttendees_', 'expandAdhocAttendance_', 'findAdhocAttendanceContainer_',
      'findAdhocAttendanceInsertIndex_', 'finishAdhocAttendanceRebuild_', 'harvestAdhocAttendeeCompanies_', 'importTeamsAttendance', 'isAdhocAttendeeTable_',
      'mergeAdhocCompanyCorrections_', 'planTeamsAttendanceImport_', 'previewTeamsAttendanceImport', 'readAdhocAttendanceIndex_', 'readAdhocAttendanceValue_',
      'readAdhocAttendance_', 'refreshAdhocAttendanceSection', 'refreshAdhocAttendanceSection_', 'removeAdhocBodyChild_', 'removeTeamsAttendanceImport',
      'renderAdhocAttendanceSection_', 'renderAfterAdhocAttendanceChange_', 'splitAdhocAttendanceChunks_', 'teamsAttendanceTextFromBase64_',
      'withAdhocAttendanceLock_', 'writeAdhocAttendanceValue_'];
    // TDoc session assignment and the Session column (stage B).
    const TDOC_SESSION_FUNCTIONS_CODE = ['adhocSessionCutoff_', 'adhocSessionIdsKey_', 'adhocTdocSessionCellText_', 'adhocTdocSessionReferences_', 'adhocTdocSessionsUnavailable_',
      'assignAdhocSession_', 'assignAdhocTdocSessions', 'buildAdhocTdocSessionsDialogModel_', 'convertUtcWallClockToReportZone_', 'effectiveAdhocTdocSessions_',
      'enableAdhocSessionColumn', 'findRegistrationTable_', 'flattenTdocGroups_', 'makeAdhocTdocSessionResolverSafely_', 'makeAdhocTdocSessionResolver_',
      'normalizeAdhocTdocOverrides_', 'normalizeTdocUploadValue_', 'parseAdhocTdocSessionsProperty_', 'readTdocUploadTimesForSessions_',
      'refreshRegistrationSessionColumnSafely_', 'refreshRegistrationSessionColumn_', 'refreshSessionColumnFromList_', 'registrationTableColumns_',
      'saveAdhocTdocSessions', 'saveAdhocTdocSessionsWith_', 'serializeAdhocTdocSessions_'];
    // Opening details of the sessions (stage E).
    const OPENING_FUNCTIONS_CODE = ['adhocOpeningSessionIds_', 'adhocOpeningStore_', 'adhocOpeningUnavailable_', 'adhocSessionWhenText_', 'buildAdhocOpeningDialogModel_',
      'buildAdhocOpeningLines_', 'editAdhocOpeningDetails', 'findAdhocOpeningContainer_', 'findAdhocOpeningInsertIndex_', 'finishAdhocOpeningRebuild_',
      'normalizeAdhocOpeningEntry_', 'parseAdhocOpeningProperty_', 'refreshAdhocOpeningSection_', 'renderAdhocOpeningAfterChange_', 'renderAdhocOpeningSection_',
      'saveAdhocOpeningDetails', 'saveAdhocOpeningWith_', 'serializeAdhocOpening_', 'utf8ByteLength_'];
    // Sessions of several days: date ranges and attendance records per day.
    const MULTIDAY_FUNCTIONS_CODE = ['adhocAttendanceRecordKey_', 'adhocAttendanceRecordParts_', 'adhocAttendanceRecordsOf_', 'adhocDateRangeText_', 'adhocSessionDatesText_',
      'adhocSessionDayCount_', 'adhocSessionLastDay_', 'adhocSessionRecord_', 'isAdhocMultiDaySession_',
      // Range integrity: a session's days keep its imported attendance inside.
      'adhocAttendanceDaysBySession_', 'adhocSessionsLeavingAttendanceOutside_',
      // Assign TDoc Sessions shows the upload time (display only).
      'adhocUploadedDisplayText_',
      // The registration table with the Session column gets explicit column widths.
      'applyRegistrationTableWidths_',
      // The codes of the sessions (A01, A02, ...) for the Session column, and the Online information block of the opening.
      'adhocSessionCode_', 'adhocSessionCodes_', 'adhocMeetingInformationSource_', 'buildAdhocMeetingInformationLines_', 'findAdhocMeetingInformationContainer_',
      'refreshAdhocMeetingInformationSafely_', 'renderAdhocMeetingInformation_'];
    // Status summary and statistics of the sessions (stage F).
    const STATUS_FUNCTIONS_CODE = ['adhocSessionDateRangeText_', 'adhocSessionStatusLines_', 'buildAdhocSessionStatusModel_', 'collectAdhocSessionStatus_',
      'formatAdhocSessionStatusLines_', 'showAdhocSessionStatistics', 'summarizeAdhocTdocSessions_'];
    // TDoc upload completion: an existing TDoc gets its link and one attempt at its abstract once it is uploaded.
    const UPLOAD_COMPLETION_FUNCTIONS_CODE = ['addMissingTdocLink_', 'completeInsertedUploadedTdoc_', 'refreshRegistrationTableLinks_', 'refreshUploadedTdocMetadata_', 'tdocListLinkIsKnown_',
      'tdocListLinkUrl_', 'tdocsNotUploadedYet_'];
    // The Status of a TDoc table as a native dropdown (T-2026.10.7).
    // T-2026.10.8: the conversion of an existing report.
    const STATUS_DROPDOWN_FUNCTIONS_CODE = ['convertStatusFieldsToDropdowns', 'formatStatusDropdownMigrationPreview_', 'formatStatusDropdownMigrationResult_', 'migrateStatusFieldsToDropdowns_', 'planStatusDropdownMigration_',
      'sendStatusDropdownInsertPairs_', 'statusDocsOtherContent_', 'statusDropdownMigrationOption_', 'statusDropdownNumberList_', 'captureStatusDropdownValuesForRebuild_', 'statusDropdownHealCandidates_', 'statusDropdownOptionByName_', 'statusDropdownRebuildChoice_',
      'applyTdocStatusUpdateToDropdown_', 'buildStatusDocsIndex_', 'createStatusDropdownDefinition_', 'finalizeStatusDropdowns_', 'mapPortalStatusToDropdownOption_',
      'noteStatusDropdownCandidate_', 'planStatusDropdownRequests_', 'readStatusDropdownSourceDefinition_', 'readTdocStatus_', 'resetStatusDropdownRun_', 'statusCellHoldsDropdown_', 'statusDocsCellEntry_',
      'statusDocsCellText_', 'statusDocsGet_', 'statusDropdownDefinitionRequest_', 'statusDropdownEntryOfTable_', 'statusDropdownIndex_', 'statusDropdownLabel_', 'statusDropdownRun_', 'statusDropdownsAvailable_'];
    // After T-2026.10.8: no abstract for a withdrawn TDoc; the start time and time zone of an ad-hoc meeting in its opening
    // sentence; the drafts folder of a new ad-hoc report, confirmed at setup (ReportCreator.js).
    const AFTER_10_8_FUNCTIONS_CODE = ['abstractFetchBlockedBy_', 'isWithdrawnStatus_', 'tdocListStatus_', 'tdocListStatuses_',
      'buildMeetingOpeningSentence_', 'computeMeetingStartTimeFromStartDate_', 'computeMeetingTimeZoneLabel_', 'getMeetingStartForOpening_', 'isValidMeetingStartTime_',
      'isValidMeetingTimeZoneLabel_', 'meetingStartBasis_', 'meetingStartFromResolved_', 'storeMeetingStart_',
      // the personal Reviewer API token: the one resolver
      'resolveReviewerApiToken_'];
    const AFTER_10_8_FUNCTIONS_CREATOR = ['confirmAdhocRevisionsFolderWith_', 'describeRevisionsFolderOutcome_', 'retryAdhocRevisionsFolderWith_', 'afterTemplateConfigurationSaved_', 'currentRevisionsCandidate_',
      // the personal Reviewer API token: the private settings file in the user's Drive
      'applyPersonalReviewerTokenPlanSafely_', 'applyPersonalReviewerTokenPlanWith_', 'beforeTemplateConfigurationSaved_', 'classifyUserSettingsFile_', 'clearPersonalReviewerTokenWith_', 'describeReviewerTokenStatus_',
      'inspectPersonalReviewerSettingsWith_', 'isUsableReviewerTokenValue_', 'isVerifiedUserSettingsFile_', 'parseUserSettingsText_', 'personalTokenProblemForLog_', 'readPersonalReviewerTokenSafely_', 'serializeUserSettings_',
      'templateReviewerTokenDialogParts_', 'userSettingsDrive_', 'writePersonalReviewerTokenWith_',
      // After T-2026.10.10 (Code.js 2.22.0): the shared minutes, one section at the end of ReportCreator.js (tests/shared-minutes.test.js)
      'sharedMinutesUrl_', 'sharedMinutesName_', 'sharedMinutesHtml_', 'sharedMinutesAppProperties_', 'sharedMinutesQuery_', 'isSharedMinutesFile_', 'hasAnyoneWriterPermission_', 'serializeSharedMinutes_',
      'readSharedMinutes_', 'sharedMinutesDrive_', 'currentSharedMinutesWith_', 'writeSharedMinutesLinkWith_', 'createSharedMinutesWith_', 'verifySharedMinutesWith_', 'forgetSharedMinutesWith_', 'sharedMinutesTextHolder_',
      'findSharedMinutesLine_', 'sharedMinutesLineLink_', 'findSharedMinutesInsertIndex_', 'renderSharedMinutesLink_', 'removeSharedMinutesLine_', 'finishSharedMinutesRebuild_', 'sharedMinutesView_', 'currentSharedMinutesView_',
      'sharedMinutesDialogParts_', 'liveSharedMinutesDeps_', 'runSharedMinutesAction_', 'createSharedMinutes', 'verifySharedMinutes', 'forgetSharedMinutes'];
    check('the new functions are exactly these: TEMPLATE-003 (ReportCreator.js), the ad-hoc sessions and attendance, the TDoc upload completion, the status dropdowns, the work after T-2026.10.8, and the shared minutes',
      [now.filter((f) => before.indexOf(f) === -1).sort(), functionNames(CODE).filter((f) => functionNames(OLD_CODE).indexOf(f) === -1).sort()],
      [['buildTemplateReportMenu_', 'confirmTemplateBuildFromScratch_', 'describeUpdateReportNowFailure_', 'updateReportNow']
        .concat(SESSION_FUNCTIONS_CODE, SESSION_FUNCTIONS_CREATOR, TEAMS_PARSER_FUNCTIONS_CODE, ATTENDANCE_FUNCTIONS_CODE, TDOC_SESSION_FUNCTIONS_CODE, OPENING_FUNCTIONS_CODE, STATUS_FUNCTIONS_CODE, MULTIDAY_FUNCTIONS_CODE, UPLOAD_COMPLETION_FUNCTIONS_CODE, STATUS_DROPDOWN_FUNCTIONS_CODE, AFTER_10_8_FUNCTIONS_CODE, AFTER_10_8_FUNCTIONS_CREATOR).sort(),
        SESSION_FUNCTIONS_CODE.concat(TEAMS_PARSER_FUNCTIONS_CODE, ATTENDANCE_FUNCTIONS_CODE, TDOC_SESSION_FUNCTIONS_CODE, OPENING_FUNCTIONS_CODE, STATUS_FUNCTIONS_CODE, MULTIDAY_FUNCTIONS_CODE, UPLOAD_COMPLETION_FUNCTIONS_CODE, STATUS_DROPDOWN_FUNCTIONS_CODE, AFTER_10_8_FUNCTIONS_CODE).sort()]);
  }
}

// ================================================================ 3. CENTRAL / Legacy menu

const ADDON_TREE = [
  '⚠️Scripts⚠️',
  '  📝 INITIAL SETUP',
  '    ⚙️ Configure Meeting Settings -> configureMeetingSettings',
  '    🧪 Test All Connections -> testAllConnections',
  '    📋 Create Configuration Tables -> createConfigurationTables',
  '  🚀 REPORT OPERATIONS',
  '    ▶️ Run Full Report Build -> runFullReportBuild',
  '    ---',
  '    0️⃣ Configure Meeting -> configureMeetingSettings',
  '    1️⃣+2️⃣ Build Skeleton + TDOC Tables -> buildSkeletonWithTdocTables',
  '    3️⃣ Collect E-mail Discussion -> collectEmailDiscussionOnly',
  '    4️⃣ Collect Revisions -> collectRevisionsOnly',
  '    5️⃣ Add Abstracts -> addAbstractsOnly',
  '    ---',
  '    🔄 Continuous Update (New TDOCs + Status) -> continuousUpdate',
  '    ⏰ Manage Auto-Update Trigger -> manageTriggers',
  '    ---',
  '    📝 Legacy: Build Initial Report -> buildInitialReport',
  '    🔄 Update Report (During Meeting) -> updateReportIncremental',
  '    📊 Analyze Report Status -> analyzeReportStatus',
  '  📋 DOCUMENT MANAGEMENT',
  '    ➕ Add Document Reallocation -> addDocumentReallocation',
  '    📊 View All Reallocations -> viewAllReallocations',
  '    🗑️ Clear All Reallocations -> clearAllReallocations',
  '    🔄 Apply Document Reallocations -> applyDocumentReallocations',
  '    🔀 Re-arrange Revision Tables -> rearrangeRevisionTables',
  '    ---',
  '    📄 Parse Agenda Document -> parseAgendaDocument',
  '    🏗️ Auto-Create Report Structure -> autoCreateReportStructure',
  '    ---',
  '    🧹 Clean Up Wrong Email Discussions -> cleanUpWrongEmailDiscussions',
  '    🔧 Remove Duplicate Email Entries -> removeDuplicateEmailEntries',
  '  🔧 TOOLS & DIAGNOSTICS',
  '    🧪 Test TDOC List URL -> testTdocListUrl',
  '    🧪 Test Reviewer API -> testReviewerApi',
  '    🧪 Test Email Feeds -> testEmailFeeds',
  '    🧪 Test Revisions Folder -> testRevisionsFolder',
  '    ✅ Validate Configuration -> validateConfiguration',
  '    🗑️ Clear All Caches -> clearAllCaches',
  '  🎨 FORMATTING & FIXES',
  '    🎨 Format Document -> removeRowHeightAndSpacing',
  '    🔗 Fix Links (portal→FTP) -> rewritePortalLinksInDoc_',
  '    📏 Fix Column Widths -> fixColumnWidths',
  '  📧 EMAIL EXPORT',
  '    Prepare TDoc Discussion E-mails -> prepareTdocDiscussionEmails',
  '  ☁️ CENTRAL ADD-ON (hourly)',
  '    ▶️ Enable Automatic Updates (this doc) -> enableAutomaticUpdatesForAddon',
  '    ⏹️ Disable Automatic Updates (this doc) -> disableAutomaticUpdatesForAddon',
  '    ⏱️ Set Update Interval (this doc) -> setAutomaticUpdateIntervalForAddon',
  '    📊 Show Add-on Status (this doc) -> showAddonSchedulerStatusForAddon',
  '  ---',
  '  ⚠️ Legacy: Update All -> updateAll'
];

console.log('3. the CENTRAL add-on / Legacy menu (no Release.js) is unchanged');
{
  // Code.js alone, as the CENTRAL add-on project and a Legacy-style bound copy have it.
  const addon = loadCode().sandbox;
  const ui = fakeUi();
  addon.DocumentApp.getUi = () => ui;
  addon.onOpen();
  check('exact tree, as before TEMPLATE-003', tree(ui), ADDON_TREE);
  ui.menus.length = 0;
  addon.onInstall({});
  check('onInstall() builds the same menu', tree(ui), ADDON_TREE);

  // Code.js + ReportCreator.js without Release.js: still not the template runtime.
  const noRelease = runtime({ release: false });
  noRelease.s.onOpen();
  check('with ReportCreator.js but without Release.js: the same menu', tree(noRelease.ui), ADDON_TREE);

  if (haveBaseline) {
    const base = loadCode().sandbox;
    const old = { console };
    ['Logger', 'PropertiesService', 'DocumentApp', 'Session', 'Utilities', 'UrlFetchApp', 'DriveApp', 'SpreadsheetApp', 'HtmlService', 'ScriptApp', 'MimeType', 'XmlService', 'LockService']
      .forEach((k) => { old[k] = base[k]; });
    vm.createContext(old);
    vm.runInContext(OLD_CODE, old, { filename: 'Code.js@T-2026.10.3' });
    const oldUi = fakeUi();
    old.DocumentApp.getUi = () => oldUi;
    old.onOpen();
    check('identical to the menu the T-2026.10.3 Code.js builds without Release.js', tree(oldUi), ADDON_TREE);
  }
  const onOpen = functionSource(CODE, 'onOpen');
  check('onOpen(): the template runtime returns before the first line of this menu',
    [onOpen.indexOf('if (templateRuntimeRelease_()) {') !== -1, onOpen.indexOf('if (templateRuntimeRelease_()) {') < onOpen.indexOf("ui.createMenu('⚠️Scripts⚠️')"),
      /templateRuntime\b(?!Release_)/.test(onOpen.slice(onOpen.indexOf("ui.createMenu('⚠️Scripts⚠️')")))],
    [true, true, false]);
}

// ================================================================ 4. A: Build Report from Scratch

/** A set-up template report whose build core is recorded instead of run. */
function buildReport(opts) {
  const o = opts || {};
  const r = runtime({ docId: REPORT_ID, props: SET_UP, release: o.release });
  r.s.runFullReportBuildCore_ = () => {
    r.events.push('WORK build');
    return o.fail ? { ok: false, failedPhase: 'e-mail', error: 'RSS is down', phases: [{ name: 'skeleton', status: 'done', ms: 1 }, { name: 'e-mail', status: 'failed', ms: 1, detail: 'RSS is down' }], totalMs: 2 }
      : { ok: true, phases: [], totalMs: 1 };
  };
  if (o.uiFailsAfterWork) {
    const alert = r.ui.alert;
    r.ui.alert = (...args) => {
      if (r.events.some((e) => /^WORK/.test(e))) { r.events.push('ALERT ATTEMPT ' + args[0]); throw new Error('Service Documents failed while accessing document with id FAKE_DOC.'); }
      return alert(...args);
    };
  }
  return r;
}

console.log('4. Build Report from Scratch (A)');
{
  check('the menu item calls runFullReportBuild', REPORT_TREE_AFTER_SETUP.filter((l) => /Build Report from Scratch/.test(l)), ['    Build Report from Scratch… -> runFullReportBuild']);

  const r = buildReport();
  const out = r.s.runFullReportBuild();
  const [title, text, buttons] = r.ui.alerts[0];
  check('one question, before any work; nothing after the work', r.events, ['ALERT Build Report from Scratch', 'WORK build']);
  check('a YES/NO question', buttons, 'YES_NO');
  check('it says the content of the document is replaced', [/replaces the content of this document/.test(text), /Everything that is in the document now is removed/.test(text)], [true, true]);
  check('it says minutes and other content entered by hand will be lost', /Meeting minutes and any other content entered by hand will be lost/.test(text), true);
  check('it names what is kept, and the alternative that loses nothing', [/Only the Document Reallocations table is kept/.test(text), /Update Report Now/.test(text)], [true, true]);
  check('it says there is no completion message', /no completion message/.test(text), true);
  check('it ends with the question', /Replace the content of this document and build the report from scratch\?$/.test(text), true);
  check('a completed build ends normally and returns its result', out && out.ok, true);
  check('title', title, 'Build Report from Scratch');

  const no = buildReport();
  no.ui.alertResponse = 'NO';
  check('No: nothing is built, no further dialog, no error', [thrown(() => no.s.runFullReportBuild()), no.events], [null, ['ALERT Build Report from Scratch']]);

  const uiDown = buildReport({ uiFailsAfterWork: true });
  check('no UI call is attempted after the build (a UI that fails then cannot fail the build)',
    [thrown(() => uiDown.s.runFullReportBuild()), uiDown.events], [null, ['ALERT Build Report from Scratch', 'WORK build']]);

  const failed = buildReport({ fail: true });
  const message = thrown(() => failed.s.runFullReportBuild());
  check('a failed phase still throws (execution "Failed"), without a pop-up', [/^Full report build failed: RSS is down/.test(message || ''), failed.events], [true, ['ALERT Build Report from Scratch', 'WORK build']]);

  const wrapper = functionSource(CODE, 'runFullReportBuild').replace(/^\s*\/\/.*$/gm, '');
  check('runFullReportBuild(): nothing but the result handling after the build (source)',
    /ui\.|getUi|DocumentApp|saveAndClose/.test(wrapper.slice(wrapper.indexOf('runFullReportBuildCore_()'))), false);
  check('the confirmation is the only UI call of confirmTemplateBuildFromScratch_()', (functionSource(CREATOR, 'confirmTemplateBuildFromScratch_').match(/ui\.alert\(/g) || []).length, 1);

  // The clearing the warning speaks of is real: the skeleton step empties the body.
  check('the build does clear the document body (source)', /getActiveDocument\(\)\.getBody\(\)\.clear\(\)/.test(functionSource(CODE, 'buildSkeletonWithTdocTables')), true);

  const addon = buildReport({ release: false });
  addon.s.runFullReportBuild();
  check('add-on / Legacy runtime: the question is the old one, word for word',
    [addon.ui.alerts[0][0], /^This will run all steps:\n\n1\+2\) Build skeleton from agenda \+ insert TDOC tables\n3\) Collect e-mail discussion\n4\) Collect revisions\n5\) Add abstracts\n\nContinue\?$/.test(addon.ui.alerts[0][1])],
    ['Run Full Report Build', true]);
}

// ================================================================ 5. B: Update Report Now

/**
 * A set-up report with the REAL continuousUpdateCore_() and collectorUpdate_();
 * only the network / document work below them is recorded. opts: { failList,
 * failEmail, failRevisions, lockBusy, release }.
 */
function updateReport(opts) {
  const o = opts || {};
  const r = runtime({ docId: o.docId || REPORT_ID, props: SET_UP, release: o.release });
  const lock = { tries: 0, releases: 0 };
  r.lock = lock;
  r.s.LockService = { getDocumentLock: () => ({ tryLock: () => { lock.tries++; return !o.lockBusy; }, releaseLock: () => { lock.releases++; } }) };
  r.s.assertMeetingReadyToBuild_ = () => {};
  r.s.downloadAndGroupTdocs_ = () => { r.events.push('WORK TDoc list'); if (o.failList) throw new Error('TDoc list not reachable'); return {}; };
  r.s.rearrangeRevisionTables_ = () => { r.events.push('WORK revision placement'); return { moved: 0, dispositions: 0 }; };
  r.s.checkRSSFeed_ = () => { r.events.push('WORK e-mail'); if (o.failEmail) throw new Error('RSS is down'); };
  r.s.updateRevisions_ = () => { r.events.push('WORK revisions'); if (o.failRevisions) throw new Error('drafts folder not found'); };
  r.s.removeRowHeightAndSpacing = () => r.events.push('WORK formatting');
  r.s.updateReportIncremental = () => r.events.push('CALLED updateReportIncremental');
  return r;
}
const COMPLETE_UPDATE = ['WORK TDoc list', 'WORK revision placement', 'WORK e-mail', 'WORK revisions'];

console.log('5. Update Report Now (B): the complete update');
{
  check('the menu item calls updateReportNow', REPORT_TREE_AFTER_SETUP.filter((l) => /Update Report Now/.test(l)), ['    Update Report Now -> updateReportNow']);
  check('nothing in the template menu calls updateReportIncremental or continuousUpdate', reportTargets.filter((f) => /updateReportIncremental|^continuousUpdate$/.test(f)), []);

  const r = updateReport();
  let out;
  const message = thrown(() => { out = r.s.updateReportNow(); });
  check('a completed update ends normally', [message, out && out.success], [null, true]);
  check('it ran the complete update: TDoc list, revision placement, e-mail, revisions', r.events, COMPLETE_UPDATE);
  check('no dialog at all: none before, none after the document was changed', r.ui.alerts.length + r.ui.dialogs.length, 0);
  check('it did not go through updateReportIncremental', r.events.indexOf('CALLED updateReportIncremental'), -1);
  check('it held the document lock for the run and released it', [r.lock.tries, r.lock.releases], [1, 1]);

  // The same work as the function the timer runs.
  const timer = updateReport();
  timer.s.continuousUpdate();
  check('the same work as continuousUpdate()', timer.events, r.events);
  const src = functionSource(CREATOR, 'updateReportNow');
  check('updateReportNow(): continuousUpdateCore_() once, no UI call (source)',
    [(src.match(/continuousUpdateCore_\(\)/g) || []).length, /getUi|ui\.|\.alert\(|updateReportIncremental|collectorUpdate_/.test(src.replace(/^\s*\/\/.*$/gm, ''))], [1, false]);
  // The runtime never saves or flushes the document itself (decided with
  // 2.17.3: no saveAndClose()). The TEMPLATE-002C suite asserts this for
  // Code.js; this is the same rule for the template-only file, which now
  // holds a document-changing entry point of its own. Carried over from
  // the T-2026.10.2 post-build diagnostic (d1dcb6c), whose other checks
  // described the pop-up that 2.17.3 removed.
  check('ReportCreator.js never saves or flushes the document explicitly (source)',
    /saveAndClose\s*\(|\.flush\s*\(/.test(CREATOR.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')), false);
}

console.log('5. Update Report Now (B): failures are visible');
{
  const list = updateReport({ failList: true });
  const m1 = thrown(() => list.s.updateReportNow());
  check('the update itself fails: the menu function throws', [/^The report update failed: TDoc list not reachable/.test(m1 || ''), list.ui.alerts.length], [true, 0]);
  check('the lock is released after a failure', [list.lock.tries, list.lock.releases], [1, 1]);
  check('the failure is logged as well', list.logs.filter((l) => /^Update Report Now failed: The report update failed: TDoc list not reachable/.test(l)).length, 1);

  const email = updateReport({ failEmail: true });
  const m2 = thrown(() => email.s.updateReportNow());
  check('e-mail collection fails: thrown after the rest of the update ran', [/^The report update finished, but one part failed:\n• e-mail discussions: RSS is down/.test(m2 || ''), email.events], [true, COMPLETE_UPDATE]);

  const both = updateReport({ failEmail: true, failRevisions: true });
  const m3 = thrown(() => both.s.updateReportNow()) || '';
  check('both collector steps fail: both are named', [/2 parts failed/.test(m3), /• e-mail discussions: RSS is down/.test(m3), /• TDoc revisions: drafts folder not found/.test(m3)], [true, true, true]);
  check('no pop-up is used for any failure', email.ui.alerts.length + both.ui.alerts.length, 0);

  const busy = updateReport({ lockBusy: true });
  const m4 = thrown(() => busy.s.updateReportNow());
  check('another update is running: nothing is done, and that is an error too', [/^The report was not updated: another update of this report is running/.test(m4 || ''), busy.events, busy.lock.releases], [true, [], 0]);

  const master = updateReport({ docId: TEMPLATE_ID });
  check('in the master template it refuses before taking the lock', [/SA4 Report Template itself/.test(thrown(() => master.s.updateReportNow()) || ''), master.events, master.lock.tries], [true, [], 0]);

  const d = updateReport().s.describeUpdateReportNowFailure_;
  check('describeUpdateReportNowFailure_ (pure)', [d({ success: true, error: null }), d({ success: true, error: null, collectorFailures: [] }),
    /^The report update failed: no result was returned/.test(d(undefined)), /^The report update failed: x/.test(d({ success: false, error: 'x' }))], ['', '', true, true]);
}

console.log('5. the automatic trigger behaves as before');
{
  // continuousUpdate() -- the function the timer calls -- still swallows and logs.
  const list = updateReport({ failList: true });
  check('continuousUpdate(): a failed update ends normally (logged, not thrown)',
    [thrown(() => list.s.continuousUpdate()), list.logs.some((l) => l === 'ERROR: TDoc list not reachable'), list.ui.alerts.length], [null, true, 0]);
  const email = updateReport({ failEmail: true });
  check('continuousUpdate(): a failed e-mail collection ends normally', [thrown(() => email.s.continuousUpdate()), email.events], [null, COMPLETE_UPDATE]);
  const busy = updateReport({ lockBusy: true });
  check('continuousUpdate(): a busy lock is skipped silently', [thrown(() => busy.s.continuousUpdate()), busy.events], [null, []]);
  check('continuousUpdate() returns nothing', updateReport().s.continuousUpdate(), undefined);

  const ok = updateReport().s.continuousUpdateCore_();
  check('continuousUpdateCore_(): the result of a clean run is what it was', ok, { success: true, error: null });
  const failed = updateReport({ failEmail: true }).s.continuousUpdateCore_();
  check('continuousUpdateCore_(): a collector failure is added to the result, success stays true',
    failed, { success: true, error: null, collectorFailures: [{ step: 'e-mail discussions', error: 'RSS is down' }] });
  check('collectorUpdate_(): one step failing does not stop the other', updateReport({ failEmail: true }).s.collectorUpdate_(), { failures: [{ step: 'e-mail discussions', error: 'RSS is down' }] });

  // The add-on scheduler and the other callers of the collector do not read the new values.
  check('continuousUpdateForDocument_() reads success / error only',
    [/coreResult\.success/.test(functionSource(CODE, 'continuousUpdateForDocument_')), /collectorFailures/.test(functionSource(CODE, 'continuousUpdateForDocument_'))], [true, false]);
  check('collectorFailures is read nowhere in Code.js; only Update Report Now reads it',
    [(CODE.replace(/^\s*(\/\/|\*).*$/gm, '').match(/collectorFailures/g) || []).length, /result\.collectorFailures/.test(functionSource(CREATOR, 'describeUpdateReportNowFailure_'))], [1, true]);
  check('every other caller of collectorUpdate_() ignores its result',
    ['updateAll', 'updateAllFromWeb', 'buildInitialReport', 'updateReportIncremental'].map((n) => /^\s*collectorUpdate_\(\);$/m.test(functionSource(CODE, n))), [true, true, true, true]);
  check('collectorUpdate_(): a clean run returns no failures', updateReport().s.collectorUpdate_(), { failures: [] });
  check('updateReportNow exists in the template runtime only (not in Code.js)',
    [typeof loadCode().sandbox.updateReportNow, typeof updateReport().s.updateReportNow, /function updateReportNow\(/.test(CODE)], ['undefined', 'function', false]);

  check('the trigger is still created for continuousUpdate, never for updateReportNow',
    [/ScriptApp\.newTrigger\('continuousUpdate'\)/.test(CODE), /ScriptApp\.newTrigger\('continuousUpdate'\)/.test(CREATOR), /newTrigger\('updateReportNow'\)/.test(CODE + CREATOR)], [true, true, false]);
  if (haveBaseline) {
    ['continuousUpdate', 'continuousUpdateForDocument_', 'runAddonScheduler_', 'runAddonSchedulerTrigger', 'createContinuousTrigger', 'deleteContinuousTrigger',
      'updateReportIncremental', 'updateAll', 'updateAllFromWeb', 'buildInitialReport']
      .forEach((name) => check(`${name}() is byte-for-byte the T-2026.10.3 function`, functionSource(CODE, name) === functionSource(OLD_CODE, name), true));
    check('createTemplateContinuousTrigger_() is byte-for-byte the T-2026.10.3 function',
      functionSource(CREATOR, 'createTemplateContinuousTrigger_') === functionSource(OLD_CREATOR, 'createTemplateContinuousTrigger_'), true);
  }
}

// ================================================================ 6. partial updates, Automatic Updates

console.log('6. the partial updates and Automatic Updates call their existing functions');
{
  const map = {};
  REPORT_TREE_AFTER_SETUP.filter((l) => / -> /.test(l)).forEach((l) => { const [label, fn] = l.trim().split(' -> '); map[label] = fn; });
  check('Update E-mail Discussions / TDoc Revisions / Abstracts',
    [map['Update E-mail Discussions'], map['Update TDoc Revisions'], map['Update Abstracts']], ['collectEmailDiscussionOnly', 'collectRevisionsOnly', 'addAbstractsOnly']);
  check('Automatic Updates, Configure Meeting, Prepare Discussion E-mails',
    [map['🔄 Automatic Updates…'], map['⚙️ Configure Meeting…'], map['💬 Prepare Discussion E-mails…']], ['manageTriggers', 'configureMeetingSettings', 'prepareTdocDiscussionEmails']);

  const r = runtime({ docId: REPORT_ID, props: SET_UP });
  r.s.checkRSSFeed_ = () => r.events.push('WORK e-mail');
  r.s.updateRevisions_ = () => r.events.push('WORK revisions');
  r.s.addAbstractsForTables_ = () => { r.events.push('WORK abstracts'); return { candidatesProcessed: 0, requestsMade: 0, cacheSkips: 0, rowsInserted: 0 }; };
  r.s.collectEmailDiscussionOnly();
  r.s.collectRevisionsOnly();
  r.s.addAbstractsOnly();
  check('each runs its own step once and reports as before', r.events,
    ['WORK e-mail', 'ALERT Success', 'WORK revisions', 'ALERT Success', 'WORK abstracts', 'ALERT Success']);
  check('their messages are the existing ones', r.ui.alerts.map((a) => a[1].replace(/:.*$/, '')),
    ['E-mail discussion collection completed.', 'Revision collection completed.', 'Abstract step completed']);

  r.s.ScriptApp = { TriggerSource: { CLOCK: 'CLOCK' }, getProjectTriggers: () => [] };
  r.s.HtmlService = { createHtmlOutput: (html) => { const out = { html, setWidth: () => out, setHeight: () => out }; return out; } };
  r.s.manageTriggers();
  check('Automatic Updates opens the existing trigger dialog with 15 / 30 / 60 minutes',
    [r.ui.dialogs.length, r.ui.dialogs[0].title, ['15', '30', '60'].map((m) => r.ui.dialogs[0].html.indexOf('<option value="' + m + '"') !== -1)],
    [1, 'Manage Continuous Update Trigger', [true, true, true]]);

  // The abstracts hint is shared with the add-on: it names no menu item of either menu.
  const NEW_HINT = 'Abstracts can also be updated manually at any time.';
  const OLD_HINT = 'The menu step "5️⃣ Add Abstracts" always works regardless of this setting.';
  const addon = runtime({ release: false });
  addon.s.ScriptApp = r.s.ScriptApp;
  addon.s.HtmlService = r.s.HtmlService;
  addon.s.manageTriggers();
  [['template report', r.ui.dialogs[0].html], ['add-on', addon.ui.dialogs[0].html]].forEach(([where, html]) => {
    check(`trigger dialog (${where}): neutral abstracts hint, no menu item named`,
      [html.indexOf(NEW_HINT) !== -1, /5️⃣|Add Abstracts|menu step|Update Abstracts/.test(html)], [true, false]);
  });
  check('no user-facing text of Code.js says "The menu step" any more', /The menu step/.test(CODE.replace(/^\s*(\/\/|\*).*$/gm, '')), false);
  check('title and Start / Stop wording of the dialog are unchanged',
    [/<h2>⏰ Continuous Update Trigger<\/h2>/.test(r.ui.dialogs[0].html), /▶️ Start Trigger/.test(r.ui.dialogs[0].html), /⏹️ Stop Trigger/.test(functionSource(CODE, 'manageTriggers'))], [true, true, true]);
  if (haveBaseline) {
    check('manageTriggers() differs from T-2026.10.3 in that one sentence only',
      [functionSource(CODE, 'manageTriggers') === functionSource(OLD_CODE, 'manageTriggers'), functionSource(CODE, 'manageTriggers').replace(NEW_HINT, OLD_HINT) === functionSource(OLD_CODE, 'manageTriggers')], [false, true]);
  }
  if (haveBaseline) {
    const STATUS_HOOK = '\n  // Ad-hoc sessions (stage F): sessions, opening details, attendance and\n  // TDoc sessions. No lines -- and nothing read -- unless this is an ad-hoc\n' +
      '  // report with sessions.\n  adhocSessionStatusLines_().forEach(line => lines.push(line));\n';
    // T-2026.10.7: the status of a table is read through readTdocStatus_(), which adds the value of a dropdown to the text read as before.
    const STATUS_READ = "    // T-2026.10.7: the text of the Status cell as always, else the value selected in its dropdown.\n    const status = readTdocStatus_(t, findStatusText_(t));\n";
    check('analyzeReportStatus() differs from T-2026.10.3 in the one call that adds the block of an ad-hoc report with sessions (stage F) and in how it reads a status (T-2026.10.7)',
      [functionSource(CODE, 'analyzeReportStatus') === functionSource(OLD_CODE, 'analyzeReportStatus'),
        functionSource(CODE, 'analyzeReportStatus').replace(STATUS_HOOK, '').replace(STATUS_READ, '    const status = findStatusText_(t);\n') === functionSource(OLD_CODE, 'analyzeReportStatus')], [false, true]);
    // T-2026.10.7: the two functions that end with the status-dropdown step.
    check('collectRevisionsOnly() differs from T-2026.10.3 in the status-dropdown step around the collection only',
      [functionSource(CODE, 'collectRevisionsOnly') === functionSource(OLD_CODE, 'collectRevisionsOnly'),
        functionSource(CODE, 'collectRevisionsOnly').replace('  resetStatusDropdownRun_();\n', '').replace("  // T-2026.10.7: revision tables added by this run get their dropdown (see finalizeStatusDropdowns_()).\n  finalizeStatusDropdowns_('update');\n", '') === functionSource(OLD_CODE, 'collectRevisionsOnly')], [false, true]);
    check('runFullReportBuildCore_() differs from T-2026.10.3 in the status-dropdown phase at its end only',
      [functionSource(CODE, 'runFullReportBuildCore_') === functionSource(OLD_CODE, 'runFullReportBuildCore_'),
        functionSource(CODE, 'runFullReportBuildCore_').replace('  resetStatusDropdownRun_();\n', '').replace(/  \/\/ T-2026\.10\.7: the last thing that touches the document\.[\s\S]*?\n  if \(statusDropdownsAvailable_\(\)\) \{\n[\s\S]*?\n    \}\);\n  \}\n/, '') === functionSource(OLD_CODE, 'runFullReportBuildCore_')], [false, true]);
    // After T-2026.10.8 the Save of Configure Meeting carries the start its Resolve found: configureMeetingSettings() was the
    // T-2026.10.3 function up to T-2026.10.8, and tests/after-10-8-ledger.test.js names every line that differs from that release.
    const CODE_10_8 = gitShow('template-release/T-2026.10.8:Code.js');
    if (CODE_10_8) {
      check('configureMeetingSettings() was byte-for-byte the T-2026.10.3 function in T-2026.10.8, and differs from it now',
        [functionSource(CODE_10_8, 'configureMeetingSettings') === functionSource(OLD_CODE, 'configureMeetingSettings'), functionSource(CODE, 'configureMeetingSettings') === functionSource(OLD_CODE, 'configureMeetingSettings')], [true, false]);
    }
    // "Check Connections" reads the Reviewer token through the one resolver now (the personal token): one line, and the rest is the T-2026.10.3 function.
    check('testAllConnections() differs from T-2026.10.3 in the line that gets the Reviewer token only',
      [functionSource(CODE, 'testAllConnections') === functionSource(OLD_CODE, 'testAllConnections'),
        functionSource(CODE, 'testAllConnections').replace('    const token = resolveReviewerApiToken_().token;\n', "    const token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN');\n") === functionSource(OLD_CODE, 'testAllConnections')], [false, true]);
    ['collectEmailDiscussionOnly', 'addAbstractsOnly', 'prepareTdocDiscussionEmails',
      'addDocumentReallocation', 'viewAllReallocations', 'applyDocumentReallocations', 'clearAllReallocations', 'removeRowHeightAndSpacing',
      'removeDuplicateEmailEntries', 'cleanUpWrongEmailDiscussions', 'clearAllCaches']
      .forEach((name) => check(`${name}() is byte-for-byte the T-2026.10.3 function`, functionSource(CODE, name) === functionSource(OLD_CODE, name), true));
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll TEMPLATE-003 menu checks passed.');
process.exitCode = failures ? 1 : 0;
