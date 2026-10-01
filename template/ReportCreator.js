/*******************************
 * SA4 Report Template -- report creator, first run, template menus and
 * the bound Continuous Update trigger (TEMPLATE-002B).
 *
 * See docs/SA4_REPORT_TEMPLATE_ARCHITECTURE.md and
 * docs/TEMPLATE-002B_RELEASE_CANDIDATE.md.
 *
 * This file is meant to be pushed ONLY into the bound script of the master
 * "SA4 Report Template" Google Doc, next to an unchanged Code.js and a
 * generated Release.js (tools/template-release.js). It is never part of the
 * CENTRAL or Legacy push payload: both .claspignore files whitelist named
 * files only, and this file lives under template/.
 *
 * Two roles, decided by the id of the document the script runs in:
 *   - the master template itself (id === SA4_RELEASE_.templateDocumentId):
 *     "Create New SA4 Report" copies the template and hands the copy a
 *     bootstrap payload. Report operations are refused here.
 *   - a report created from the template: "Finish Report Setup" reads the
 *     payload, stores it as ordinary Document Properties through the same
 *     persistConfigurationSettings_() the Configure Meeting dialog uses,
 *     and records which release the report was created with.
 *
 * Why a Drive description and not Document Properties: properties belong
 * to one script project, and a copied document gets a NEW bound project
 * whose property stores start empty. The creator (template project) cannot
 * write the copy's Document Properties. The Drive file description is plain
 * file metadata both projects can read and write with the drive scope the
 * manifest already requests.
 *
 * Pure functions (no Apps Script service calls) are the tested core; the
 * *With_(deps) functions take every service as an injected dependency so the
 * control flow is tested with fakes; the public menu/RPC functions at the
 * bottom are the only code that touches live services.
 *
 * TEMPLATE-001 (live, 2026-10-01, GO): a DriveApp.makeCopy() copy keeps the
 * bound script as its own project (own Script ID), starts with empty
 * property stores and no triggers, can read the description the creator
 * wrote, and can create its own everyMinutes(30) trigger. The copy call in
 * liveTemplateDeps_() is the one that was tested. The entry points
 * themselves have not run live yet (first live smoke test).
 *
 * Code.js stays the one runtime source. It reaches into this file only
 * through templateRuntimeRelease_() being non-null, i.e. only when the
 * generated Release.js is present -- and a release bundle always contains
 * Release.js and this file together.
 *******************************/

var TEMPLATE_BOOTSTRAP_SCHEMA_ = 'sa4-report-bootstrap/1';
var TEMPLATE_BOOTSTRAP_MARKER_ = 'SA4-BOOTSTRAP-V1:';
var TEMPLATE_BOOTSTRAP_MAX_CHARS_ = 4000;

// Configuration keys a bootstrap payload may carry -- the Configure Meeting
// dialog's own config-object names (see saveConfigurationSettings()), so the
// payload is persisted by exactly the code path a manual Save uses.
var TEMPLATE_BOOTSTRAP_CONFIG_KEYS_ = [
  'meetingId', 'meetingType', 'meetingName', 'meetingDate', 'ftpBase',
  'meetingFolder', 'meetingNumber', 'reportType', 'agendaTdoc', 'agendaCsvUrl',
  'tdocUrl', 'revisionsUrl', 'mailingList', 'mailingListMode', 'emailStartDate'
];
var TEMPLATE_BOOTSTRAP_URL_KEYS_ = ['ftpBase', 'agendaCsvUrl', 'tdocUrl', 'revisionsUrl'];
var TEMPLATE_BOOTSTRAP_ALLOWED_HOSTS_ = ['www.3gpp.org', 'ftp.3gpp.org', 'portal.3gpp.org'];

// Document Properties the first run records (never read by Code.js).
var TEMPLATE_STATE_KEYS_ = {
  state: 'SA4_BOOTSTRAP_STATE',
  createdFromRelease: 'SA4_CREATED_FROM_RELEASE',
  createdAt: 'SA4_CREATED_AT',
  setupRelease: 'SA4_SETUP_RELEASE',
  scriptId: 'SA4_BOUND_SCRIPT_ID'
};

// ------------------------------------------------------------------
// Release metadata
// ------------------------------------------------------------------

function validateTemplateRelease_(release) {
  const errors = [];
  if (!release || typeof release !== 'object') return ['No template release metadata (Release.js) is present.'];
  if (!/^T-\d{4}\.\d{2}\.\d+$/.test(String(release.releaseId || ''))) errors.push('releaseId must look like T-2026.10.0.');
  if (release.flavor !== 'template') errors.push('flavor must be "template".');
  if (!isPlausibleDriveId_(release.templateDocumentId)) errors.push('templateDocumentId is missing or malformed.');
  if (!/^[0-9a-f]{7,40}$/.test(String(release.gitCommit || ''))) errors.push('gitCommit is missing or malformed.');
  return errors;
}

function isPlausibleDriveId_(id) {
  return /^[A-Za-z0-9_-]{25,100}$/.test(String(id === null || id === undefined ? '' : id));
}

// ------------------------------------------------------------------
// Continuous Update trigger of the bound template runtime
// ------------------------------------------------------------------

// A bound script is not an add-on, so everyMinutes(15|30) is allowed
// (ClockTriggerBuilder: n must be 1, 5, 10, 15 or 30; the hourly-only rule
// of the CENTRAL add-on does not apply). Every hour is everyHours(1) --
// never everyMinutes(60), the Legacy defect. 30 minutes is the default: a
// consumer account has 90 minutes of trigger runtime per day across all
// scripts, so 15 minutes is for meeting days only.
var TEMPLATE_CONTINUOUS_TRIGGER_INTERVALS_ = [
  { minutes: 15, label: 'Every 15 minutes (Active meeting)', everyMinutes: 15 },
  { minutes: 30, label: 'Every 30 minutes (Recommended)', everyMinutes: 30, selected: true },
  { minutes: 60, label: 'Every hour (Slow meeting)', everyHours: 1 }
];

function templateContinuousTriggerIntervals_() {
  return TEMPLATE_CONTINUOUS_TRIGGER_INTERVALS_;
}

/**
 * Installs the Continuous Update trigger of a template report. `interval`
 * is an entry of the table above, already validated by
 * createContinuousTrigger() in Code.js -- an interval that is not offered
 * never gets here. The new trigger is created FIRST; only when that
 * succeeded are the previous continuousUpdate triggers removed, so a failed
 * request (quota, authorization) leaves the running trigger as it was.
 * Repeating the call ends with exactly one trigger. Triggers of any other
 * handler are never touched.
 */
function createTemplateContinuousTrigger_(interval, fetchAbstracts) {
  assertNotTemplateMaster_();

  const previous = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'continuousUpdate';
  });

  const builder = ScriptApp.newTrigger('continuousUpdate').timeBased();
  if (interval.everyMinutes) {
    builder.everyMinutes(interval.everyMinutes).create();
  } else {
    builder.everyHours(interval.everyHours).create();
  }

  previous.forEach(function (t) { ScriptApp.deleteTrigger(t); });

  if (fetchAbstracts !== undefined && fetchAbstracts !== null) {
    setFetchAbstractsSetting(fetchAbstracts);
  }
  PropertiesService.getDocumentProperties().setProperty(CONTINUOUS_UPDATE_INTERVAL_KEY_, String(interval.minutes));
  Logger.log('Trigger created: every ' + interval.minutes + ' minutes (abstracts: ' + getFetchAbstractsSetting_() + ')');
}

// ------------------------------------------------------------------
// Meeting -> bootstrap payload (pure)
// ------------------------------------------------------------------

/**
 * Main meetings: MEETING_FOLDER/MEETING_NUMBER from the resolved FTP folder
 * (…/WG4_CODEC/TSGS4_137_Xian/Docs/ -> TSGS4_137_Xian, 137;
 * …/TSGS4_137-e/Docs/ -> TSGS4_137-e, 137-e). Without them
 * persistConfigurationSettings_() would write the historical SA4#136
 * defaults, so an underivable folder is an error, never a guess.
 */
function deriveMainMeetingFolderFromFtpBase_(ftpBase) {
  const m = String(ftpBase || '').trim().match(/\/WG4_CODEC\/(TSGS4_(\d+(?:-e)?)(?:[_-][A-Za-z0-9_.-]*)?)\/Docs\/$/i);
  if (!m) return { folder: null, number: null, error: 'The main-meeting folder could not be derived from the document folder "' + ftpBase + '".' };
  // "137-e" is kept: the TDoc list is TDoc_List_Meeting_SA4%23137-e.xlsx.
  return { folder: m[1], number: m[2], error: null };
}

function trimmed_(v) {
  return String(v === null || v === undefined ? '' : v).trim();
}

/**
 * Merges a computeResolvedMeetingPreview_() result with Thomas's explicit
 * choices into a bootstrap payload.
 *
 * choices: { reportFamily, agendaTdoc, tdocListUrl, mailingList } -- all
 * optional; a non-blank choice always wins over a discovered value.
 *
 * Creation is refused only when the report cannot be named or identified
 * (meeting unresolved, no family). Everything else the readiness rules
 * need (agenda not yet published, no TDoc list yet) becomes `pending`:
 * the report is created and first run shows what is still missing, because
 * reports are routinely created before the agenda exists.
 */
function buildReportBootstrapPayload_(preview, choices, release, nowIso) {
  const p = preview || {};
  const c = choices || {};
  const v = function (name) { return p[name] && p[name].value ? trimmed_(p[name].value) : ''; };
  const errors = [];
  const notes = [];

  const meetingType = v('meetingType').toLowerCase();
  const meetingId = v('meetingId');
  if (!meetingId || !/^\d+$/.test(meetingId)) errors.push('Resolve the 3GPP meeting first.');
  if (meetingType !== 'adhoc' && meetingType !== 'main') errors.push('The meeting type could not be determined (' + (meetingType || 'blank') + ').');

  // Report family: explicit choice > confident/suggested inference. A main
  // meeting holds several families, so it always needs the choice.
  let family = trimmed_(c.reportFamily);
  const inference = p.familyInference || {};
  if (family && !Object.prototype.hasOwnProperty.call(REPORT_FAMILY_LABELS_, family)) {
    errors.push('Unknown report family "' + family + '".');
    family = '';
  }
  if (!family && meetingType === 'adhoc' && (inference.confidence === 'confident' || inference.confidence === 'suggested')) {
    family = inference.family;
    notes.push('Report family ' + family + ' detected (' + inference.confidence + ', from ' + (inference.sources || []).join(' + ') + ').');
  }
  if (!family) {
    errors.push(meetingType === 'main'
      ? 'Choose the report family -- a main meeting has one report per family.'
      : 'Choose the report family (' + (inference.confidence === 'conflict' ? 'the evidence conflicts' : 'nothing identifies it') + ').');
  }

  const config = {
    meetingId: meetingId,
    meetingType: meetingType,
    meetingName: v('meetingName'),
    meetingDate: v('meetingDate'),
    ftpBase: v('ftpBase'),
    meetingFolder: '',
    meetingNumber: '',
    reportType: family,
    agendaTdoc: trimmed_(c.agendaTdoc) || v('agendaTdoc'),
    agendaCsvUrl: v('agendaCsvUrl'),
    tdocUrl: trimmed_(c.tdocListUrl) || v('tdocListUrl'),
    revisionsUrl: p.revisionsUrl && p.revisionsUrl.source !== 'candidate' ? v('revisionsUrl') : '',
    mailingList: trimmed_(c.mailingList),
    mailingListMode: trimmed_(c.mailingList) ? 'override' : 'derived',
    // Decision 2026-10-01: a new report collects e-mail from the meeting
    // start date; the user can move it in Configure Meeting.
    emailStartDate: meetingStartDateIso_(p.startDateRaw) || meetingStartDateIso_(v('meetingDate'))
  };

  if (meetingType === 'main') {
    const main = deriveMainMeetingFolderFromFtpBase_(config.ftpBase);
    if (main.error) errors.push(main.error);
    config.meetingFolder = main.folder || '';
    config.meetingNumber = main.number || '';
    // A main report reads the series agenda/TDoc list by folder + number.
    config.agendaCsvUrl = '';
  }

  const readiness = evaluateMeetingReadiness_({
    meetingType: meetingType,
    meetingId: config.meetingId,
    meetingName: config.meetingName,
    ftpBase: config.ftpBase,
    reportFamily: config.reportType,
    agendaTdoc: config.agendaTdoc,
    agendaCsvUrl: config.agendaCsvUrl,
    tdocListUrl: config.tdocUrl,
    mailingList: config.mailingList || (family ? (MAILING_LISTS[family] || LIST_NAME_LOCK) : ''),
    meetingFolder: config.meetingFolder,
    meetingNumber: config.meetingNumber
  }, MEETING_READINESS_RULES_, {});
  const pending = readiness.issues
    .filter(function (i) { return i.code !== 'MEETING_NOT_RESOLVED' && i.code !== 'REPORT_FAMILY_REQUIRED'; })
    .map(function (i) { return i.message; });

  const payload = {
    schema: TEMPLATE_BOOTSTRAP_SCHEMA_,
    createdAt: nowIso,
    templateRelease: release ? release.releaseId : null,
    templateDocumentId: release ? release.templateDocumentId : null,
    targetDocumentId: null,
    config: config,
    pending: pending
  };
  const urlErrors = validateBootstrapUrls_(config);
  return {
    ok: errors.length === 0 && urlErrors.length === 0,
    errors: errors.concat(urlErrors),
    notes: notes,
    pending: pending,
    title: errors.length === 0 ? proposeReportTitle_(config) : null,
    payload: payload
  };
}

function validateBootstrapUrls_(config) {
  const errors = [];
  TEMPLATE_BOOTSTRAP_URL_KEYS_.forEach(function (key) {
    const url = trimmed_(config[key]);
    if (!url) return;
    const m = url.match(/^https:\/\/([^/?#:]+)(?:[/?#]|$)/i);
    if (!m || TEMPLATE_BOOTSTRAP_ALLOWED_HOSTS_.indexOf(m[1].toLowerCase()) === -1) {
      errors.push(key + ' is not an https URL on a 3GPP host: ' + url);
    }
  });
  return errors;
}

/**
 * The title the build itself will give the report (generateReportTitle_()),
 * so the file is named correctly from the start and the build's rename is a
 * no-op. Ad-hoc: "<family topic> Minutes – <meeting name>"; main:
 * "<family topic> Minutes SA4#<number>".
 */
function proposeReportTitle_(config) {
  if (config.meetingType === 'adhoc') {
    return generateReportTitle_({ REPORT_SUFFIX: config.reportType, meetingLabel: config.meetingName });
  }
  const tdocListUrl = config.tdocUrl ||
    (config.ftpBase && config.meetingNumber ? config.ftpBase + 'TDoc_List_Meeting_SA4%23' + config.meetingNumber + '.xlsx' : '');
  return generateReportTitle_({ REPORT_SUFFIX: config.reportType, TDOC_LIST_URL: tdocListUrl, MEETING_ID: config.meetingId });
}

// ------------------------------------------------------------------
// Payload transport: the copy's Drive file description (pure)
// ------------------------------------------------------------------

/** Human-readable first line, then the marker and the JSON payload. */
function serializeBootstrapDescription_(payload) {
  const cfg = payload.config || {};
  const text = 'SA4 report being set up from template ' + payload.templateRelease + ' for meeting ' + cfg.meetingId +
    '. Open it and run ⚠️Scripts⚠️ > 🚀 Finish Report Setup.\n' + TEMPLATE_BOOTSTRAP_MARKER_ + JSON.stringify(payload);
  if (text.length > TEMPLATE_BOOTSTRAP_MAX_CHARS_) throw new Error('Bootstrap payload too large (' + text.length + ' characters).');
  return text;
}

function parseBootstrapDescription_(text) {
  const s = String(text === null || text === undefined ? '' : text);
  if (s.length > TEMPLATE_BOOTSTRAP_MAX_CHARS_) return { payload: null, error: 'The setup information is too large.' };
  const at = s.indexOf(TEMPLATE_BOOTSTRAP_MARKER_);
  if (at === -1) return { payload: null, error: null };
  try {
    return { payload: JSON.parse(s.slice(at + TEMPLATE_BOOTSTRAP_MARKER_.length)), error: null };
  } catch (e) {
    return { payload: null, error: 'The setup information in the file description is damaged.' };
  }
}

/**
 * Remote-origin data rules: only the known schema, only whitelisted string
 * config keys, only 3GPP https URLs, and only for THIS document -- a copy of
 * a not-yet-set-up report (whose description would travel with it) is
 * refused instead of silently configuring a second report.
 */
function validateBootstrapPayload_(payload, documentId) {
  const errors = [];
  if (!payload || typeof payload !== 'object') return ['No setup information found.'];
  if (payload.schema !== TEMPLATE_BOOTSTRAP_SCHEMA_) errors.push('Unsupported setup information (' + payload.schema + ').');
  if (payload.targetDocumentId !== documentId) {
    errors.push('The setup information was created for a different document. Create the report again from the template.');
  }
  const cfg = payload.config;
  if (!cfg || typeof cfg !== 'object') {
    errors.push('The setup information has no meeting configuration.');
    return errors;
  }
  Object.keys(cfg).forEach(function (key) {
    if (TEMPLATE_BOOTSTRAP_CONFIG_KEYS_.indexOf(key) === -1) errors.push('Unexpected setup field "' + key + '".');
    else if (typeof cfg[key] !== 'string') errors.push('Setup field "' + key + '" is not text.');
  });
  if (!/^\d+$/.test(String(cfg.meetingId || ''))) errors.push('The meeting ID is missing.');
  if (cfg.meetingType !== 'adhoc' && cfg.meetingType !== 'main') errors.push('The meeting type is missing.');
  if (!Object.prototype.hasOwnProperty.call(REPORT_FAMILY_LABELS_, cfg.reportType)) errors.push('The report family is missing.');
  if (cfg.mailingListMode !== 'derived' && cfg.mailingListMode !== 'override') errors.push('Invalid mailing-list mode.');
  if (cfg.emailStartDate && !isValidCollectorStartDate_(cfg.emailStartDate)) errors.push('The e-mail collection start date is not a date.');
  return errors.concat(validateBootstrapUrls_(cfg));
}

/** The dialog config object persistConfigurationSettings_() expects. */
function bootstrapPayloadToConfig_(payload) {
  const cfg = payload.config;
  const config = {};
  TEMPLATE_BOOTSTRAP_CONFIG_KEYS_.forEach(function (key) { config[key] = cfg[key] || ''; });
  config.showPreview = true;          // the dialog's default for a new document
  config.agendaSourceDocId = '';      // -> the default preamble source document
  config.apiTokenAction = 'keep';     // never touches the Reviewer token
  config.apiToken = '';
  return config;
}

function describeBootstrapProvenance_(payload, setupRelease, nowIso) {
  const cfg = payload.config;
  return 'SA4 report for meeting ' + cfg.meetingId + ' (' + cfg.reportType + '), created ' + payload.createdAt +
    ' from template ' + payload.templateRelease + '; set up ' + nowIso + ' with ' + setupRelease + '.';
}

// ------------------------------------------------------------------
// Control flow with injected services (tested with fakes)
// ------------------------------------------------------------------

/**
 * deps: { release, activeDocumentId(), nowIso(), copyTemplate(title) ->
 * {id, url}, setDescription(id, text), trashFile(id) }.
 * request: { preview, choices } -- preview is recomputed server-side by the
 * caller; the client never supplies payload fields directly.
 */
function createReportFromTemplateWith_(deps, request) {
  const release = deps.release;
  const releaseErrors = validateTemplateRelease_(release);
  if (releaseErrors.length) return { ok: false, errors: releaseErrors };
  if (templateDocumentRole_(deps.activeDocumentId(), release) !== 'template') {
    return { ok: false, errors: ['New reports are created from the SA4 Report Template document only.'] };
  }
  const built = buildReportBootstrapPayload_(request.preview, request.choices, release, deps.nowIso());
  if (!built.ok) return { ok: false, errors: built.errors, notes: built.notes };

  const copy = deps.copyTemplate(built.title);
  const payload = JSON.parse(JSON.stringify(built.payload));
  payload.targetDocumentId = copy.id;
  try {
    deps.setDescription(copy.id, serializeBootstrapDescription_(payload));
  } catch (e) {
    // A copy without setup information would open as an unconfigured
    // report; move it to the trash (recoverable) and report the failure.
    deps.trashFile(copy.id);
    return { ok: false, errors: ['The report was copied but could not be prepared (' + e.message + '); the copy was moved to the trash.'] };
  }
  return { ok: true, documentId: copy.id, url: copy.url, title: built.title, pending: built.pending, notes: built.notes };
}

/**
 * First run in a newly created report. Idempotent: a second run, or a run in
 * a report that was configured some other way, changes nothing.
 *
 * deps: { release, activeDocumentId(), scriptId(), nowIso(),
 * documentProperties, getOwnDescription(), setOwnDescription(text),
 * persistConfiguration(config), buildReadiness() -> readiness }.
 */
function finishReportSetupWith_(deps) {
  const release = deps.release;
  const documentId = deps.activeDocumentId();
  if (templateDocumentRole_(documentId, release) === 'template') {
    return { status: 'refused', errors: ['This is the SA4 Report Template itself -- create a report from it instead.'] };
  }
  const props = deps.documentProperties;
  const state = props.getProperty(TEMPLATE_STATE_KEYS_.state);
  if (state === 'done') return { status: 'already-done', readiness: deps.buildReadiness() };
  if (state === 'manual') return { status: 'not-a-new-report', readiness: deps.buildReadiness() };

  const parsed = parseBootstrapDescription_(deps.getOwnDescription());
  if (parsed.error) return { status: 'refused', errors: [parsed.error] };
  if (!parsed.payload) {
    // An ordinary document that merely contains the code: nothing to set
    // up. It is configured by hand (Configure Meeting), as before. The
    // marker only records that this was checked once, so the check (a Drive
    // read) and the "Finish Report Setup" menu item do not recur.
    props.setProperty(TEMPLATE_STATE_KEYS_.state, 'manual');
    if (props.getProperty('MEETING_ID')) return { status: 'not-a-new-report', readiness: deps.buildReadiness() };
    return { status: 'no-setup-info', errors: ['No setup information found. Create reports from the SA4 Report Template, or use Configure Meeting.'] };
  }
  const errors = validateBootstrapPayload_(parsed.payload, documentId);
  if (props.getProperty('MEETING_ID') && props.getProperty('MEETING_ID') !== parsed.payload.config.meetingId) {
    errors.push('This document is already configured for meeting ' + props.getProperty('MEETING_ID') + '.');
  }
  if (errors.length) return { status: 'refused', errors: errors };

  // Properties first, "done" marker last: an interrupted run is simply
  // repeated (every write is an idempotent overwrite of the same value).
  const config = bootstrapPayloadToConfig_(parsed.payload);
  // An e-mail collection start date already stored in this document is an
  // explicit value: the proposed meeting start date never replaces it.
  if (isValidCollectorStartDate_(props.getProperty('EMAIL_START_DATE'))) config.emailStartDate = '';
  deps.persistConfiguration(config);
  const now = deps.nowIso();
  const setupRelease = release ? release.releaseId : 'unknown';
  props.setProperty(TEMPLATE_STATE_KEYS_.createdFromRelease, String(parsed.payload.templateRelease));
  props.setProperty(TEMPLATE_STATE_KEYS_.createdAt, String(parsed.payload.createdAt));
  props.setProperty(TEMPLATE_STATE_KEYS_.setupRelease, setupRelease);
  props.setProperty(TEMPLATE_STATE_KEYS_.scriptId, String(deps.scriptId()));
  props.setProperty(TEMPLATE_STATE_KEYS_.state, 'done');

  const warnings = [];
  try {
    deps.setOwnDescription(describeBootstrapProvenance_(parsed.payload, setupRelease, now));
  } catch (e) {
    warnings.push('The file description could not be tidied up (' + e.message + '); this is harmless.');
  }

  return { status: 'configured', readiness: deps.buildReadiness(), pending: parsed.payload.pending || [], warnings: warnings };
}

// ------------------------------------------------------------------
// Which kind of document is this? (pure)
// ------------------------------------------------------------------

/**
 * 'master-template'   the template itself (document id from Release.js)
 * 'created-report'    set up from the creator's bootstrap information
 * 'ordinary-document' checked once, no bootstrap information: a document
 *                     that merely contains the code, configured by hand
 * 'not-set-up'        not looked at yet (a new copy before its first run,
 *                     or an ordinary document before its first menu action)
 */
function templateDocumentOrigin_(documentId, release, bootstrapState) {
  if (templateDocumentRole_(documentId, release) === 'template') return 'master-template';
  if (bootstrapState === 'done') return 'created-report';
  if (bootstrapState === 'manual') return 'ordinary-document';
  return 'not-set-up';
}

/** Lines for "Template Release Info" / "About This Report" (pure). */
function describeTemplateRuntime_(facts) {
  const release = facts.release || {};
  const origin = templateDocumentOrigin_(facts.documentId, facts.release, facts.bootstrapState);
  const lines = [];
  if (origin === 'master-template') {
    lines.push('SA4 Report Template (master document)');
    lines.push('Reports are created from here. Nothing is built or configured in this document.');
  } else if (origin === 'created-report') {
    lines.push('Report created from the SA4 Report Template');
    lines.push('Created ' + (facts.createdAt || '?') + ' from release ' + (facts.createdFromRelease || '?'));
  } else if (origin === 'ordinary-document') {
    lines.push('Report (not created by the template creator; configured by hand)');
  } else {
    lines.push('Report, not set up yet: use Finish Report Setup, Configure Meeting or Run Full Report Build');
  }
  lines.push('');
  lines.push('Script release: ' + (release.releaseId || '?') + ' (Code.js ' + (release.codeVersion || '?') + ', commit ' + String(release.gitCommit || '?').slice(0, 7) + ')');
  if (origin !== 'master-template') {
    lines.push('Meeting ID: ' + (facts.meetingId || 'not configured'));
    lines.push('Continuous Update: ' + (facts.continuousInterval ? 'every ' + facts.continuousInterval : 'off'));
  }
  lines.push('Script ID: ' + (facts.scriptId || '?'));
  lines.push('');
  lines.push('This document keeps this script release until someone deliberately updates it.');
  return lines;
}

// ------------------------------------------------------------------
// Menus (called from onOpen() in Code.js, template runtime only)
// ------------------------------------------------------------------

/** The master template: creation and release information, nothing else. */
function buildTemplateMasterMenu_(ui) {
  ui.createMenu('⚠️Scripts⚠️')
    .addItem('🆕 Create New SA4 Report', 'showCreateReportDialog')
    .addSeparator()
    .addItem('ℹ️ Template Release Info', 'showTemplateInfo')
    .addToUi();
}

/**
 * A report: "Finish Report Setup" leads the menu until the first run has
 * looked at the document once. onOpen() is a simple trigger, so only the
 * document's own properties are read here (no authorization needed); if
 * even that fails the item is simply shown.
 */
function addTemplateReportMenuHead_(menu) {
  let checked = false;
  try {
    checked = !!PropertiesService.getDocumentProperties().getProperty(TEMPLATE_STATE_KEYS_.state);
  } catch (e) {
    checked = false;
  }
  if (!checked) {
    menu.addItem('🚀 Finish Report Setup', 'finishReportSetup');
    menu.addSeparator();
  }
}

function addTemplateReportMenuTail_(menu) {
  menu.addItem('ℹ️ About This Report', 'showTemplateInfo');
}

// ------------------------------------------------------------------
// Live entry points. Copy semantics were verified by TEMPLATE-001; these
// functions themselves first run live in the TEMPLATE-002B smoke test.
// Public names (no trailing "_") because menus and google.script.run can
// only call public functions (RESOLVER-HOTFIX).
// ------------------------------------------------------------------

function liveTemplateDeps_() {
  const doc = DocumentApp.getActiveDocument();
  return {
    release: templateRuntimeRelease_(),
    activeDocumentId: function () { return doc.getId(); },
    scriptId: function () { return ScriptApp.getScriptId(); },
    nowIso: function () { return new Date().toISOString(); },
    documentProperties: PropertiesService.getDocumentProperties(),
    getOwnDescription: function () { return DriveApp.getFileById(doc.getId()).getDescription(); },
    setOwnDescription: function (text) { DriveApp.getFileById(doc.getId()).setDescription(text); },
    // The copy mechanism TEMPLATE-001 verified live: DriveApp only, same
    // folder, description written after the copy.
    copyTemplate: function (title) {
      const template = DriveApp.getFileById(doc.getId());
      const parents = template.getParents();
      const copy = parents.hasNext() ? template.makeCopy(title, parents.next()) : template.makeCopy(title);
      return { id: copy.getId(), url: copy.getUrl() };
    },
    setDescription: function (id, text) { DriveApp.getFileById(id).setDescription(text); },
    trashFile: function (id) { DriveApp.getFileById(id).setTrashed(true); },
    persistConfiguration: function (config) { persistConfigurationSettings_(config); },
    buildReadiness: function () { return getBuildReadiness_(); }
  };
}

/** Only what computeResolvedMeetingPreview_() reads -- the raw Portal responses stay on the server. */
function slimResolvedMeeting_(resolved) {
  const r = resolved || {};
  return {
    id: r.id, meeting: r.meeting, sources: r.sources, documents: r.documents,
    adhocSources: r.adhocSources, warnings: r.warnings, unresolved: r.unresolved
  };
}

/** What the creator dialog shows for a resolved meeting (pure). */
function describeNewReportProposal_(resolved, choices, release, nowIso) {
  const preview = computeResolvedMeetingPreview_({}, resolved);
  const built = buildReportBootstrapPayload_(preview, choices, release, nowIso);
  const cfg = built.payload.config;
  const family = cfg.reportType;
  const lines = [
    ['Meeting', preview.summary || ''],
    ['Report family', family ? REPORT_FAMILY_LABELS_[family] : 'please choose'],
    ['Mailing list', cfg.mailingList || (family ? (MAILING_LISTS[family] || LIST_NAME_LOCK) + ' (from the family)' : '')],
    ['Document folder', cfg.ftpBase || 'not found'],
    ['TDoc list', preview.tdocListStatus || cfg.tdocUrl || (cfg.meetingType === 'main' ? 'derived from the meeting folder' : 'not found yet')],
    ['Agenda', preview.agendaSourceStatus || cfg.agendaTdoc || (cfg.meetingType === 'main' ? 'from the meeting agenda' : 'not found yet')],
    ['E-mail collection from', cfg.emailStartDate || 'default']
  ];
  return {
    ok: built.ok, errors: built.errors, notes: built.notes, pending: built.pending, title: built.title,
    family: family || '', lines: lines, resolved: slimResolvedMeeting_(resolved)
  };
}

/**
 * Creator step 1 (RPC): resolve + discover the meeting and propose a report,
 * without writing anything anywhere. Reuses the Configure Meeting pipeline
 * unchanged: resolveMeetingCoreById_ -> enrichMeetingFromTdocList_ ->
 * applyAdhocSourceDiscovery_ -> computeResolvedMeetingPreview_.
 */
function previewNewReportFromTemplate(meetingIdInput, choices) {
  const idResult = parseMeetingIdInput_(meetingIdInput);
  if (!idResult.isValid) return { ok: false, errors: [idResult.error] };
  const core = resolveMeetingCoreById_(idResult.id);
  if (!core.meeting) return { ok: false, errors: (core.warnings && core.warnings.length) ? core.warnings : ['The meeting could not be resolved.'] };
  const enriched = enrichMeetingFromTdocList_(idResult.id, core);
  applyAdhocSourceDiscovery_(enriched, idResult.id, {});
  return describeNewReportProposal_(enriched, choices, templateRuntimeRelease_(), new Date().toISOString());
}

/** Creator step 1b (RPC): the same proposal for another choice, without any network call. */
function previewNewReportFromResolved(meetingIdInput, resolved, choices) {
  const idResult = parseMeetingIdInput_(meetingIdInput);
  if (!idResult.isValid || !resolved || resolved.id !== idResult.id) {
    return { ok: false, errors: ['Look up the meeting again.'] };
  }
  return describeNewReportProposal_(resolved, choices, templateRuntimeRelease_(), new Date().toISOString());
}

/**
 * Creator step 2 (RPC): create the report. The preview is recomputed here
 * from the resolved meeting the dialog holds (the pure merge, no refetch --
 * the trust model of discoverAgendaForConfigDialog_()), and every payload
 * field is re-validated before anything is copied.
 */
function createNewReportFromTemplate(meetingIdInput, resolved, choices) {
  const idResult = parseMeetingIdInput_(meetingIdInput);
  if (!idResult.isValid || !resolved || resolved.id !== idResult.id) {
    return { ok: false, errors: ['Look up the meeting again before creating the report.'] };
  }
  const preview = computeResolvedMeetingPreview_({}, resolved);
  return createReportFromTemplateWith_(liveTemplateDeps_(), { preview: preview, choices: choices });
}

/** Menu (master template): the creator dialog. */
function showCreateReportDialog() {
  const release = templateRuntimeRelease_();
  if (!release || !isTemplateMasterDocument_()) {
    DocumentApp.getUi().alert('Create New SA4 Report', 'New reports are created from the SA4 Report Template document only.', DocumentApp.getUi().ButtonSet.OK);
    return;
  }
  const familyOptions = '<option value="">(detect automatically)</option>' + Object.keys(REPORT_FAMILY_LABELS_).map(function (family) {
    return '<option value="' + family + '">' + REPORT_FAMILY_LABELS_[family] + '</option>';
  }).join('');
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      label { display: block; margin-top: 12px; font-weight: bold; }
      input, select { padding: 7px; box-sizing: border-box; }
      button { padding: 8px 16px; border: none; cursor: pointer; color: white; background: #4285f4; }
      button:disabled { background: #b0b0b0; cursor: default; }
      table { border-collapse: collapse; margin-top: 12px; width: 100%; }
      td { padding: 3px 8px 3px 0; vertical-align: top; }
      td.k { color: #555; white-space: nowrap; }
      .release { color: #777; font-size: 11px; }
      .err { color: #b00020; margin-top: 10px; white-space: pre-wrap; }
      .warn { color: #8a6d00; margin-top: 10px; white-space: pre-wrap; }
      .note { color: #2e7d32; margin-top: 6px; }
      #result { margin-top: 14px; }
      #title { font-weight: bold; margin-top: 10px; }
    </style>
    <div class="release">Template release ${release.releaseId}</div>
    <label>3GPP Meeting ID</label>
    <input type="text" id="meetingId" placeholder="e.g. 86172" style="width: 55%">
    <button id="lookupBtn" onclick="lookUp()">Look up</button>
    <div id="status"></div>
    <table id="details"></table>
    <div id="familyRow" style="display:none">
      <label>Report family</label>
      <select id="family" onchange="familyChanged()">${familyOptions}</select>
    </div>
    <div id="title"></div>
    <div style="margin-top: 16px">
      <button id="createBtn" onclick="createReport()" disabled>Create Report</button>
      <button style="background:#6c757d" onclick="google.script.host.close()">Close</button>
    </div>
    <div id="result"></div>
    <script>
      var resolvedMeeting = null;
      function el(id) { return document.getElementById(id); }
      function esc(v) {
        return String(v === null || v === undefined ? '' : v)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      }
      function choices() {
        var family = el('family').value;
        return family ? { reportFamily: family } : {};
      }
      function showProposal(r) {
        el('lookupBtn').disabled = false;
        if (r.resolved) resolvedMeeting = r.resolved;
        var html = '';
        (r.lines || []).forEach(function (line) { html += '<tr><td class="k">' + esc(line[0]) + '</td><td>' + esc(line[1]) + '</td></tr>'; });
        el('details').innerHTML = html;
        var status = '';
        if (r.errors && r.errors.length) status += '<div class="err">' + esc(r.errors.join(String.fromCharCode(10))) + '</div>';
        if (r.pending && r.pending.length) {
          status += '<div class="warn">The report can be created now. Still to do in Configure Meeting later:' + String.fromCharCode(10) +
            esc(r.pending.map(function (p) { return String.fromCharCode(8226) + ' ' + p; }).join(String.fromCharCode(10))) + '</div>';
        }
        (r.notes || []).forEach(function (n) { status += '<div class="note">' + esc(n) + '</div>'; });
        el('status').innerHTML = status;
        el('familyRow').style.display = resolvedMeeting ? '' : 'none';
        if (r.family && !el('family').value) el('family').value = r.family;
        el('title').textContent = r.title ? 'New report: ' + r.title : '';
        el('createBtn').disabled = !r.ok;
      }
      function showFailure(error) {
        el('lookupBtn').disabled = false;
        el('status').innerHTML = '<div class="err">' + esc(error && error.message ? error.message : error) + '</div>';
      }
      function lookUp() {
        resolvedMeeting = null;
        el('family').value = '';
        el('createBtn').disabled = true;
        el('lookupBtn').disabled = true;
        el('result').innerHTML = '';
        el('status').textContent = 'Looking up the meeting (this can take a minute)...';
        google.script.run.withSuccessHandler(showProposal).withFailureHandler(showFailure)
          .previewNewReportFromTemplate(el('meetingId').value, {});
      }
      function familyChanged() {
        if (!resolvedMeeting) return;
        google.script.run.withSuccessHandler(showProposal).withFailureHandler(showFailure)
          .previewNewReportFromResolved(el('meetingId').value, resolvedMeeting, choices());
      }
      function createReport() {
        el('createBtn').disabled = true;
        el('status').textContent = 'Creating the report...';
        google.script.run.withSuccessHandler(function (r) {
          if (!r.ok) {
            el('status').innerHTML = '<div class="err">' + esc((r.errors || []).join(String.fromCharCode(10))) + '</div>';
            el('createBtn').disabled = false;
            return;
          }
          el('status').textContent = '';
          el('result').innerHTML = '<div class="note">Report created: ' + esc(r.title) + '</div>' +
            '<p><a href="' + esc(r.url) + '" target="_blank">Open the new report</a></p>' +
            '<p>In the report: Scripts menu, then Run Full Report Build (or Configure Meeting first). ' +
            'Google asks for permission once for the new report; allow it and click the item again.</p>';
        }).withFailureHandler(showFailure)
          .createNewReportFromTemplate(el('meetingId').value, resolvedMeeting, choices());
      }
    </script>
  `).setWidth(620).setHeight(600);
  DocumentApp.getUi().showModalDialog(html, 'Create New SA4 Report');
}

/**
 * Stores the creator's setup information the first time a report is used.
 * Called by Configure Meeting and Run Full Report Build (Code.js), so the
 * first thing the user clicks in a new report also finishes its setup.
 * Does nothing once the document has been looked at; throws only when
 * setup information exists but cannot be used.
 */
function ensureReportBootstrapped_() {
  if (PropertiesService.getDocumentProperties().getProperty(TEMPLATE_STATE_KEYS_.state)) return null;
  const result = finishReportSetupWith_(liveTemplateDeps_());
  if (result.status === 'refused') {
    throw new Error('This report could not be set up:\n' + result.errors.join('\n'));
  }
  return result;
}

/** Menu (report): first run in a new report, with a summary. */
function finishReportSetup() {
  const ui = DocumentApp.getUi();
  const result = finishReportSetupWith_(liveTemplateDeps_());
  if (result.status === 'refused') {
    ui.alert('Report setup', '❌ ' + result.errors.join('\n'), ui.ButtonSet.OK);
    return result;
  }
  if (result.status === 'no-setup-info') {
    ui.alert('Report setup', 'This document was not created with "Create New SA4 Report", so there is nothing to finish.\n\n' +
      'Use ⚠️Scripts⚠️ > 📝 INITIAL SETUP > ⚙️ Configure Meeting Settings.', ui.ButtonSet.OK);
    return result;
  }
  const r = result.readiness || { ready: false, issues: [] };
  ui.alert('Report setup',
    (result.status === 'configured' ? '✅ Meeting configuration stored.' : 'ℹ️ This report is already set up.') + '\n\n' +
    (r.ready ? '✅ Ready to build: ⚠️Scripts⚠️ > 🚀 REPORT OPERATIONS > ▶️ Run Full Report Build.'
      : '⚠️ Not ready to build yet:\n' + r.issues.map(function (i) { return '• ' + i.message; }).join('\n') +
        '\n\nOpen Configure Meeting to discover the missing sources.') +
    ((result.warnings || []).length ? '\n\n' + result.warnings.join('\n') : ''),
    ui.ButtonSet.OK);
  return result;
}

/** Menu: "Template Release Info" (master) / "About This Report" (report). */
function showTemplateInfo() {
  const ui = DocumentApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const stored = describeStoredContinuousInterval_(props);
  const active = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'continuousUpdate'; });
  const lines = describeTemplateRuntime_({
    release: templateRuntimeRelease_(),
    documentId: DocumentApp.getActiveDocument().getId(),
    scriptId: ScriptApp.getScriptId(),
    bootstrapState: props.getProperty(TEMPLATE_STATE_KEYS_.state),
    createdAt: props.getProperty(TEMPLATE_STATE_KEYS_.createdAt),
    createdFromRelease: props.getProperty(TEMPLATE_STATE_KEYS_.createdFromRelease),
    meetingId: props.getProperty('MEETING_ID'),
    continuousInterval: active ? (stored ? stored.label : 'an unknown interval') : null
  });
  ui.alert(isTemplateMasterDocument_() ? 'Template Release Info' : 'About This Report', lines.join('\n'), ui.ButtonSet.OK);
}
