/*******************************
 * SA4 Report Template -- Report Creator + first-run bootstrap
 * PROTOTYPE (design/report-template-architecture) -- NOT DEPLOYED.
 *
 * See docs/SA4_REPORT_TEMPLATE_ARCHITECTURE.md.
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
 * bottom are the only code that touches live services, and they are
 * UNVERIFIED until the TEMPLATE-001 copy experiment has passed.
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
  'tdocUrl', 'revisionsUrl', 'mailingList', 'mailingListMode'
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

// A bound script (not an add-on) may use everyMinutes(1|5|10|15|30)
// (ClockTriggerBuilder docs); the hourly-only rule of CENTRAL 2.15.2 is an
// add-on rule. This is the 2.15.1 table, restored for the template runtime.
var TEMPLATE_CONTINUOUS_TRIGGER_INTERVALS_ = [
  { minutes: 15, label: 'Every 15 minutes (Active meeting)', everyMinutes: 15 },
  { minutes: 30, label: 'Every 30 minutes (Recommended)', everyMinutes: 30, selected: true },
  { minutes: 60, label: 'Every hour (Slow meeting)', everyHours: 1 }
];

// ------------------------------------------------------------------
// Release metadata
// ------------------------------------------------------------------

/**
 * The release this script copy was built from, or null outside the template
 * runtime. SA4_RELEASE_ is defined by the generated Release.js, which exists
 * only in template release bundles -- checked at call time, so file load
 * order never matters.
 */
function getTemplateRelease_() {
  return typeof SA4_RELEASE_ !== 'undefined' && SA4_RELEASE_ ? SA4_RELEASE_ : null;
}

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

/** 'template' for the master template, 'report' otherwise (incl. no release). */
function templateDocumentRole_(documentId, release) {
  return release && documentId && documentId === release.templateDocumentId ? 'template' : 'report';
}

/**
 * Refuses report operations (build, update, triggers, configuration) in the
 * master template, so the template can never acquire configuration, content
 * or a trigger that a later copy might inherit. TEMPLATE-002 calls this from
 * assertMeetingReadyToBuild_(), createContinuousTrigger() and
 * saveConfigurationSettings().
 */
function assertNotTemplateDocument_(documentId, release) {
  if (templateDocumentRole_(documentId, release) === 'template') {
    throw new Error('This is the SA4 Report Template itself. Use "Create New SA4 Report" -- ' +
      'reports are never built in the template.');
  }
}

/** The runtime's Continuous Update intervals: sub-hourly only in the template runtime. */
function continuousTriggerIntervalsForRuntime_(release, centralIntervals) {
  return release && release.flavor === 'template' ? TEMPLATE_CONTINUOUS_TRIGGER_INTERVALS_ : centralIntervals;
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
    mailingListMode: trimmed_(c.mailingList) ? 'override' : 'derived'
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
 * persistConfiguration(config), buildReadiness() -> readiness,
 * projectTriggerHandlers() -> [names] }.
 */
function finishReportSetupWith_(deps) {
  const release = deps.release;
  const documentId = deps.activeDocumentId();
  if (templateDocumentRole_(documentId, release) === 'template') {
    return { status: 'refused', errors: ['This is the SA4 Report Template itself -- create a report from it instead.'] };
  }
  const props = deps.documentProperties;
  if (props.getProperty(TEMPLATE_STATE_KEYS_.state) === 'done') {
    return { status: 'already-done', readiness: deps.buildReadiness() };
  }

  const parsed = parseBootstrapDescription_(deps.getOwnDescription());
  if (parsed.error) return { status: 'refused', errors: [parsed.error] };
  if (!parsed.payload) {
    if (props.getProperty('MEETING_ID')) return { status: 'not-a-new-report', readiness: deps.buildReadiness() };
    return { status: 'refused', errors: ['No setup information found. Create reports from the SA4 Report Template, or use Configure Meeting.'] };
  }
  const errors = validateBootstrapPayload_(parsed.payload, documentId);
  if (props.getProperty('MEETING_ID') && props.getProperty('MEETING_ID') !== parsed.payload.config.meetingId) {
    errors.push('This document is already configured for meeting ' + props.getProperty('MEETING_ID') + '.');
  }
  if (errors.length) return { status: 'refused', errors: errors };

  // Properties first, "done" marker last: an interrupted run is simply
  // repeated (every write is an idempotent overwrite of the same value).
  deps.persistConfiguration(bootstrapPayloadToConfig_(parsed.payload));
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
  // Installable triggers belong to one script project, so a fresh copy
  // should have none. Report (never delete) anything unexpected.
  const handlers = deps.projectTriggerHandlers();
  if (handlers.length) warnings.push('Unexpected triggers already exist in this report: ' + handlers.join(', ') + '.');

  return { status: 'configured', readiness: deps.buildReadiness(), pending: parsed.payload.pending || [], warnings: warnings };
}

// ------------------------------------------------------------------
// Live entry points -- UNVERIFIED until TEMPLATE-001 has passed.
// Public names (no trailing "_") because menus and google.script.run can
// only call public functions (RESOLVER-HOTFIX).
// ------------------------------------------------------------------

function liveTemplateDeps_() {
  const doc = DocumentApp.getActiveDocument();
  return {
    release: getTemplateRelease_(),
    activeDocumentId: function () { return doc.getId(); },
    scriptId: function () { return ScriptApp.getScriptId(); },
    nowIso: function () { return new Date().toISOString(); },
    documentProperties: PropertiesService.getDocumentProperties(),
    getOwnDescription: function () { return DriveApp.getFileById(doc.getId()).getDescription(); },
    setOwnDescription: function (text) { DriveApp.getFileById(doc.getId()).setDescription(text); },
    copyTemplate: function (title) {
      const template = DriveApp.getFileById(doc.getId());
      const parents = template.getParents();
      const copy = parents.hasNext() ? template.makeCopy(title, parents.next()) : template.makeCopy(title);
      return { id: copy.getId(), url: copy.getUrl() };
    },
    setDescription: function (id, text) { DriveApp.getFileById(id).setDescription(text); },
    trashFile: function (id) { DriveApp.getFileById(id).setTrashed(true); },
    persistConfiguration: function (config) { persistConfigurationSettings_(config); },
    buildReadiness: function () { return getBuildReadiness_(); },
    projectTriggerHandlers: function () {
      return ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
    }
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
  if (!core.meeting) return { ok: false, errors: core.warnings || ['The meeting could not be resolved.'] };
  const enriched = enrichMeetingFromTdocList_(idResult.id, core);
  applyAdhocSourceDiscovery_(enriched, idResult.id, {});
  const preview = computeResolvedMeetingPreview_({}, enriched);
  const built = buildReportBootstrapPayload_(preview, choices, getTemplateRelease_(), new Date().toISOString());
  return { ok: built.ok, errors: built.errors, notes: built.notes, pending: built.pending, title: built.title, preview: preview, resolved: enriched };
}

/**
 * Creator step 2 (RPC): create the report. The preview is recomputed here
 * from the resolved meeting the dialog already holds (the pure merge, no
 * refetch -- same trust model as discoverAgendaForConfigDialog_()), and
 * every payload field is re-validated before anything is copied.
 */
function createNewReportFromTemplate(meetingIdInput, resolved, choices) {
  const idResult = parseMeetingIdInput_(meetingIdInput);
  if (!idResult.isValid || !resolved || resolved.id !== idResult.id) {
    return { ok: false, errors: ['Resolve the meeting again before creating the report.'] };
  }
  const preview = computeResolvedMeetingPreview_({}, resolved);
  return createReportFromTemplateWith_(liveTemplateDeps_(), { preview: preview, choices: choices });
}

/** Menu: first run in a new report. */
function finishReportSetup() {
  const result = finishReportSetupWith_(liveTemplateDeps_());
  const ui = DocumentApp.getUi();
  if (result.status === 'refused') {
    ui.alert('Report setup', '❌ ' + result.errors.join('\n'), ui.ButtonSet.OK);
    return result;
  }
  const r = result.readiness || { ready: false, issues: [] };
  ui.alert('Report setup',
    (result.status === 'configured' ? '✅ Meeting configuration stored.' : 'ℹ️ This report is already set up.') + '\n\n' +
    (r.ready ? '✅ Ready to build: ⚠️Scripts⚠️ > 🚀 REPORT OPERATIONS > ▶️ Run Full Report Build.'
      : '⚠️ Not ready to build yet:\n' + r.issues.map(function (i) { return '• ' + i.message; }).join('\n') +
        '\n\nOpen Configure Meeting later to discover the missing sources.') +
    ((result.warnings || []).length ? '\n\n' + result.warnings.join('\n') : ''),
    ui.ButtonSet.OK);
  return result;
}
