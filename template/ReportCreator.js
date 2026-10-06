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
  'tdocUrl', 'revisionsUrl', 'mailingList', 'mailingListMode', 'emailStartDate',
  // The start of an ad-hoc meeting as the Portal gives it: "HH:MM" and "UTC+2" ('' when it gives none).
  'meetingStartTime', 'meetingTimeZone'
];
var TEMPLATE_BOOTSTRAP_URL_KEYS_ = ['ftpBase', 'agendaCsvUrl', 'tdocUrl', 'revisionsUrl'];
var TEMPLATE_BOOTSTRAP_ALLOWED_HOSTS_ = ['www.3gpp.org', 'ftp.3gpp.org', 'portal.3gpp.org'];

// Document Properties the first run records (never read by Code.js).
var TEMPLATE_STATE_KEYS_ = {
  state: 'SA4_BOOTSTRAP_STATE',
  createdFromRelease: 'SA4_CREATED_FROM_RELEASE',
  createdAt: 'SA4_CREATED_AT',
  setupRelease: 'SA4_SETUP_RELEASE',
  scriptId: 'SA4_BOUND_SCRIPT_ID',
  // The drafts folder derived for an ad-hoc report that could not be confirmed at setup (never used as REVISIONS_URL).
  revisionsCandidate: 'SA4_REVISIONS_URL_CANDIDATE'
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

  // Mailing list (TEMPLATE-002C): the list derived from the report family,
  // unless the user names another one. An override is kept as the plain list
  // name -- the same forms the collector accepts ("<list>" or
  // "<list>@list.etsi.org", normalizeEtsiListName_()). Naming the derived
  // list again is no override. Nothing here knows any particular meeting.
  const familyList = family ? (MAILING_LISTS[family] || LIST_NAME_LOCK) : '';
  const requestedList = trimmed_(c.mailingList);
  let listOverride = '';
  if (requestedList) {
    const normalizedList = normalizeEtsiListName_(requestedList);
    if (!normalizedList) {
      errors.push('Mailing list "' + requestedList.replace(/[\r\n]/g, ' ') + '" is not a valid ETSI list name. ' +
        'Use the list name (for example 3GPP_TSG_SA_WG4_MBS) or its ...@list.etsi.org address.');
    } else if (normalizedList.toLowerCase() !== familyList.toLowerCase()) {
      if (meetingType === 'main') {
        errors.push('A main-meeting report always reads its report family\'s list (' + familyList + '); it cannot be overridden.');
      } else {
        listOverride = normalizedList;
      }
    }
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
    mailingList: listOverride,
    mailingListMode: listOverride ? 'override' : 'derived',
    // Decision 2026-10-01: a new report collects e-mail from the meeting
    // start date; the user can move it in Configure Meeting.
    emailStartDate: meetingStartDateIso_(p.startDateRaw) || meetingStartDateIso_(v('meetingDate')),
    // The opening sentence of an ad-hoc report: the time of day and the time
    // zone of the Portal's StartDate, '' where the Portal has none. A main
    // report copies its opening from the report template and has no use for them.
    meetingStartTime: meetingType === 'adhoc' ? computeMeetingStartTimeFromStartDate_(p.startDateRaw) : '',
    meetingTimeZone: meetingType === 'adhoc' ? computeMeetingTimeZoneLabel_(p.startTimeZoneRaw) : ''
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
    mailingList: config.mailingList || familyList,
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
    mailingList: { effective: listOverride || familyList, derived: familyList, overridden: !!listOverride },
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
    '. Open it and run ' + TEMPLATE_MENU_NAME_ + ' > 🚀 Finish Report Setup.\n' + TEMPLATE_BOOTSTRAP_MARKER_ + JSON.stringify(payload);
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
  if (cfg.mailingList && normalizeEtsiListName_(cfg.mailingList) !== cfg.mailingList) errors.push('The mailing list is not a valid ETSI list name.');
  if ((cfg.mailingListMode === 'override') !== !!cfg.mailingList) errors.push('The mailing-list override is inconsistent.');
  if (cfg.emailStartDate && !isValidCollectorStartDate_(cfg.emailStartDate)) errors.push('The e-mail collection start date is not a date.');
  if (cfg.meetingStartTime && !isValidMeetingStartTime_(cfg.meetingStartTime)) errors.push('The meeting start time is not a time.');
  if (cfg.meetingTimeZone && !isValidMeetingTimeZoneLabel_(cfg.meetingTimeZone)) errors.push('The meeting time zone is not a UTC offset.');
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
 * The drafts/revisions folder of a new AD-HOC report (REVISIONS_URL: where
 * the revisions of its TDocs are found, and where the discussion e-mails
 * send them).
 *
 * An ad-hoc series has its own folder next to its document folder:
 * <meeting root>/Docs/ -> <meeting root>/inbox/drafts/
 * (deriveRevisionsUrlCandidate_(), from the Portal's MtgDocURL). That is a
 * derivation, not a fact, so it is asked for ONCE here and stored only when
 * the answer is a folder listing (validateRevisionsUrlCandidate_(): HTTP 200
 * and a listing; a redirect, 403, another status, an empty page or a failed
 * request is no confirmation). It is asked here and not when the meeting is
 * looked up, so that a slow answer of that server can never hold up the
 * lookup again (PROD-014).
 *
 * TWO WAYS a folder becomes REVISIONS_URL, on purpose:
 *   - automatically (here, and after a Save of Configure Meeting): only a
 *     derived folder that this request confirmed;
 *   - by the user: Configure Meeting shows the derived folder as "suggested --
 *     please review", and a Save with it (or with any typed address) in the
 *     field stores it WITHOUT a request. That is the manual way for when the
 *     3GPP server refuses scripted requests. It is not to be "tidied" into
 *     the automatic rule (tests/configure-meeting-setup-parity.test.js, 10).
 *
 * A folder that is already configured is never replaced, and nothing is
 * asked for it. A main-meeting report derives its folder by formula and is
 * not touched. Only an https address on a 3GPP host is asked.
 *
 * Returns { status, url, reason }:
 *   'not-applicable' a main meeting          'kept'      one is configured
 *   'stored'         confirmed and stored    'unconfirmed' not confirmed: the
 *   'no-candidate'   none could be derived                 candidate is kept
 *   'not-checked'    this runtime cannot ask               for the user to see
 * An unconfirmed candidate is recorded under its own key
 * (TEMPLATE_STATE_KEYS_.revisionsCandidate), never as REVISIONS_URL.
 *
 * deps: { documentProperties, validateRevisionsCandidate(url) -> { ok, reason } }.
 */
function confirmAdhocRevisionsFolderWith_(deps, cfg) {
  const props = deps.documentProperties;
  if (!cfg || cfg.meetingType !== 'adhoc') return { status: 'not-applicable', url: '', reason: '' };
  if (trimmed_(props.getProperty('REVISIONS_URL'))) return { status: 'kept', url: trimmed_(props.getProperty('REVISIONS_URL')), reason: '' };

  const derived = deriveRevisionsUrlCandidate_(cfg.ftpBase);
  const candidate = trimmed_(derived.revisionsUrlCandidate);
  if (!candidate) return { status: 'no-candidate', url: '', reason: trimmed_(derived.error) };
  if (validateBootstrapUrls_({ revisionsUrl: candidate }).length) {
    return { status: 'no-candidate', url: '', reason: 'The derived folder is not an https address on a 3GPP host.' };
  }
  if (typeof deps.validateRevisionsCandidate !== 'function') return { status: 'not-checked', url: candidate, reason: '' };

  let checked;
  try {
    checked = deps.validateRevisionsCandidate(candidate);
  } catch (e) {
    checked = { ok: false, reason: 'The request failed: ' + e.message };
  }
  if (checked && checked.ok === true) {
    props.setProperty('REVISIONS_URL', candidate);
    props.deleteProperty(TEMPLATE_STATE_KEYS_.revisionsCandidate);
    return { status: 'stored', url: candidate, reason: '' };
  }
  props.setProperty(TEMPLATE_STATE_KEYS_.revisionsCandidate, candidate);
  return { status: 'unconfirmed', url: candidate, reason: trimmed_(checked && checked.reason) || 'no answer' };
}

/**
 * A later, explicit chance for the drafts folder of an ad-hoc report: a save
 * of Configure Meeting. While the report has no REVISIONS_URL, the folder
 * derived from the document folder IT IS CONFIGURED WITH NOW is asked for
 * again -- one request, as at setup -- and stored when the answer is a
 * listing. So one failed request at setup does not leave the report without
 * its folder for good, and nothing is asked on a routine operation: an
 * update never calls this.
 *
 * A folder that is configured is never asked for or replaced. The candidate
 * is always derived anew: a report that was given another meeting gets the
 * candidate of that meeting, and a record of an earlier one is removed
 * unless it is the candidate that was just not confirmed.
 *
 * deps: { documentProperties, validateRevisionsCandidate(url) -> { ok, reason } }.
 */
function retryAdhocRevisionsFolderWith_(deps) {
  const props = deps.documentProperties;
  const outcome = confirmAdhocRevisionsFolderWith_(deps, {
    meetingType: trimmed_(props.getProperty('MEETING_TYPE')).toLowerCase(),
    ftpBase: trimmed_(props.getProperty('FTP_BASE'))
  });
  if (outcome.status !== 'unconfirmed') props.deleteProperty(TEMPLATE_STATE_KEYS_.revisionsCandidate);
  return outcome;
}

/**
 * Called by saveConfigurationSettings() (Code.js) in the template runtime,
 * after the configuration was stored. `tokenPlan` is what
 * beforeTemplateConfigurationSaved_() returned: the personal Reviewer token
 * is dealt with first, then the drafts folder. Returns { note } for the
 * message of the dialog ('' when there is nothing to say). Never throws: the
 * configuration is saved whatever becomes of this.
 */
function afterTemplateConfigurationSaved_(tokenPlan) {
  const notes = [];
  const tokenNote = applyPersonalReviewerTokenPlanSafely_(tokenPlan);
  if (tokenNote) {
    Logger.log('Configure Meeting: ' + tokenNote);
    notes.push(tokenNote);
  }
  try {
    const outcome = retryAdhocRevisionsFolderWith_({
      documentProperties: PropertiesService.getDocumentProperties(),
      validateRevisionsCandidate: function (url) { return validateRevisionsUrlCandidate_(url); }
    });
    const note = describeRevisionsFolderOutcome_(outcome);
    if (note) {
      Logger.log('Configure Meeting: ' + note);
      notes.push(note);
    }
  } catch (e) {
    Logger.log('Configure Meeting: the drafts folder could not be looked at: ' + e.message);
  }
  return { note: notes.join('\n\n') };
}

/** Pure: the recorded candidate, when it is the one of the document folder the report has now; '' otherwise (a record of an earlier meeting is not shown). */
function currentRevisionsCandidate_(recorded, ftpBase) {
  const candidate = trimmed_(recorded);
  return candidate && candidate === trimmed_(deriveRevisionsUrlCandidate_(ftpBase).revisionsUrlCandidate) ? candidate : '';
}

/** What the user is told about the drafts folder after setup: one text, '' when there is nothing to say (pure). */
function describeRevisionsFolderOutcome_(outcome) {
  const o = outcome || {};
  if (o.status === 'stored') return '✅ Drafts folder found and stored: ' + o.url;
  if (o.status === 'unconfirmed') {
    return '⚠️ Drafts folder still to set up. The folder derived for this meeting could not be confirmed (' + o.reason + '):\n' + o.url +
      '\nOpen it in a browser; if it is the drafts folder, enter it in ' + TEMPLATE_MENU_NAME_ + ' > ⚙️ Configure Meeting… (Revisions / Drafts URL). ' +
      'Until then no revisions are collected and no discussion e-mails can be prepared.';
  }
  if (o.status === 'no-candidate') {
    return '⚠️ Drafts folder still to set up: none could be derived (' + o.reason + '). Enter it in ' + TEMPLATE_MENU_NAME_ + ' > ⚙️ Configure Meeting… (Revisions / Drafts URL).';
  }
  return '';
}

/**
 * First run in a newly created report. Idempotent: a second run, or a run in
 * a report that was configured some other way, changes nothing.
 *
 * deps: { release, activeDocumentId(), scriptId(), nowIso(),
 * documentProperties, getOwnDescription(), setOwnDescription(text),
 * persistConfiguration(config), buildReadiness() -> readiness,
 * validateRevisionsCandidate(url) -> { ok, reason } }.
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

  // Last, after the "done" marker: whatever becomes of this one request, the
  // report is set up and the request is not made a second time.
  let revisionsFolder;
  try {
    revisionsFolder = confirmAdhocRevisionsFolderWith_(deps, parsed.payload.config);
  } catch (e) {
    revisionsFolder = { status: 'failed', url: '', reason: e.message };
  }

  return { status: 'configured', readiness: deps.buildReadiness(), pending: parsed.payload.pending || [], warnings: warnings, revisionsFolder: revisionsFolder };
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
    lines.push('Report, not set up yet: use Finish Report Setup, Configure Meeting or Build Report from Scratch');
  }
  lines.push('');
  lines.push('Script release: ' + (release.releaseId || '?') + ' (Code.js ' + (release.codeVersion || '?') + ', commit ' + String(release.gitCommit || '?').slice(0, 7) + ')');
  if (origin !== 'master-template') {
    lines.push('Meeting ID: ' + (facts.meetingId || 'not configured'));
    lines.push('Automatic updates: ' + (facts.continuousInterval ? 'every ' + facts.continuousInterval : 'off'));
    // An ad-hoc report whose drafts folder could not be confirmed at setup, until one is configured.
    if (facts.revisionsCandidate && !facts.revisionsUrl) {
      lines.push('Drafts folder: not set up. Derived, not confirmed: ' + facts.revisionsCandidate);
      lines.push('  Check it and enter it in Configure Meeting (Revisions / Drafts URL).');
    }
  }
  lines.push('Script ID: ' + (facts.scriptId || '?'));
  lines.push('');
  lines.push('This document keeps this script release until someone deliberately updates it.');
  return lines;
}

// ------------------------------------------------------------------
// Menus (called from onOpen() in Code.js, template runtime only)
// ------------------------------------------------------------------

/** TEMPLATE-003: the name of the menu in the master template and in every report. */
var TEMPLATE_MENU_NAME_ = 'SA4 Report';

/** The master template: creation and release information, nothing else. */
function buildTemplateMasterMenu_(ui) {
  ui.createMenu(TEMPLATE_MENU_NAME_)
    .addItem('🆕 Create New SA4 Report', 'showCreateReportDialog')
    .addSeparator()
    .addItem('ℹ️ Template Release Info', 'showTemplateInfo')
    .addToUi();
}

/**
 * TEMPLATE-003: the menu of a report created from the template, ordered by
 * what a rapporteur does with a report. Every item calls an existing
 * function without arguments; only "Update Report Now" has a wrapper of its
 * own (updateReportNow() below). "…" marks an item that opens a dialog or
 * asks before it acts.
 *
 * Not shown here, and still in Code.js for the CENTRAL add-on, for Legacy
 * copies and for reports created from earlier releases: the old import
 * (buildInitialReport, updateAll), updateReportIncremental,
 * buildSkeletonWithTdocTables, parseAgendaDocument, autoCreateReportStructure,
 * rearrangeRevisionTables, createConfigurationTables, the single connection
 * tests, validateConfiguration, rewritePortalLinksInDoc_, fixColumnWidths.
 */
function buildTemplateReportMenu_(ui) {
  const menu = ui.createMenu(TEMPLATE_MENU_NAME_);
  addTemplateReportMenuHead_(menu);

  menu.addSubMenu(ui.createMenu('📄 Report')
    .addItem('Build Report from Scratch…', 'runFullReportBuild')
    .addItem('Update Report Now', 'updateReportNow')
    .addSeparator()
    .addItem('Update E-mail Discussions', 'collectEmailDiscussionOnly')
    .addItem('Update TDoc Revisions', 'collectRevisionsOnly')
    .addItem('Update Abstracts', 'addAbstractsOnly')
    .addSeparator()
    .addItem('Report Status Summary', 'analyzeReportStatus')
    // T-2026.10.8: an existing report gets its status dropdowns without a rebuild.
    .addItem('Convert Status Fields to Dropdowns…', 'convertStatusFieldsToDropdowns'));

  menu.addItem('💬 Prepare Discussion E-mails…', 'prepareTdocDiscussionEmails');

  menu.addSubMenu(ui.createMenu('🔀 Document Reallocation')
    .addItem('Add or Change a Reallocation…', 'addDocumentReallocation')
    .addItem('Show Reallocations', 'viewAllReallocations')
    .addItem('Apply Reallocations to Report…', 'applyDocumentReallocations')
    .addItem('Remove All Reallocations…', 'clearAllReallocations'));

  // Ad-hoc sessions: only an ad-hoc report has this submenu. A main-meeting
  // report's menu is unchanged.
  if (isAdhocReportForMenu_()) {
    menu.addSubMenu(ui.createMenu('🗓 Sessions and Attendance')
      .addItem('Configure Sessions…', 'configureAdhocSessions')
      .addItem('Import Teams Attendance…', 'importTeamsAttendance')
      .addItem('Refresh Attendance Section', 'refreshAdhocAttendanceSection')
      .addItem('Assign TDoc Sessions…', 'assignAdhocTdocSessions')
      .addItem('Edit Opening Details…', 'editAdhocOpeningDetails')
      .addItem('Post-meeting Statistics…', 'showAdhocSessionStatistics'));
  }

  menu.addItem('🔄 Automatic Updates…', 'manageTriggers');
  menu.addItem('⚙️ Configure Meeting…', 'configureMeetingSettings');
  menu.addSeparator();

  menu.addSubMenu(ui.createMenu('🛠 Advanced and Repair')
    .addItem('Check Connections', 'testAllConnections')
    .addItem('Format Report', 'removeRowHeightAndSpacing')
    .addItem('Remove Duplicate E-mail Entries…', 'removeDuplicateEmailEntries')
    .addItem('Remove Wrong E-mail Matches…', 'cleanUpWrongEmailDiscussions')
    .addItem('Clear Collection Caches…', 'clearAllCaches'));

  addTemplateReportMenuTail_(menu);
  menu.addToUi();
}

/**
 * A report (buildTemplateReportMenu_()): "Finish Report Setup" leads the menu until the first run has
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

/**
 * Whether this document is configured as an ad-hoc report. Like the check
 * above, it reads only the document's own properties (onOpen() is a simple
 * trigger); a report that becomes ad-hoc gets the submenu the next time the
 * document is opened. If the property cannot be read, the submenu is not shown.
 */
function isAdhocReportForMenu_() {
  try {
    return String(PropertiesService.getDocumentProperties().getProperty('MEETING_TYPE') || '').trim().toLowerCase() === 'adhoc';
  } catch (e) {
    return false;
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
    buildReadiness: function () { return getBuildReadiness_(); },
    // One request, no redirect followed (fetchRevisionsUrlCandidate_()); a failure is an answer, not an error.
    validateRevisionsCandidate: function (url) { return validateRevisionsUrlCandidate_(url); }
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
    ['Document folder', cfg.ftpBase || 'not found'],
    ['TDoc list', preview.tdocListStatus || cfg.tdocUrl || (cfg.meetingType === 'main' ? 'derived from the meeting folder' : 'not found yet')],
    ['Agenda', preview.agendaSourceStatus || cfg.agendaTdoc || (cfg.meetingType === 'main' ? 'from the meeting agenda' : 'not found yet')],
    ['E-mail collection from', cfg.emailStartDate || 'default']
  ];
  return {
    ok: built.ok, errors: built.errors, notes: built.notes, pending: built.pending, title: built.title,
    family: family || '', lines: lines, resolved: slimResolvedMeeting_(resolved),
    // The list this report will read, where it comes from, and whether the
    // creator may change it (a main-meeting report always reads its family list).
    mailingList: built.mailingList, mailingListEditable: cfg.meetingType === 'adhoc'
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
    <div id="mailingRow" style="display:none">
      <label>Mailing list</label>
      <input type="text" id="mailingList" style="width: 100%" oninput="mailingTouched = true" onchange="mailingChanged()">
      <div class="release" id="mailingHint"></div>
    </div>
    <div id="title"></div>
    <div style="margin-top: 16px">
      <button id="createBtn" onclick="createReport()" disabled>Create Report</button>
      <button style="background:#6c757d" onclick="google.script.host.close()">Close</button>
    </div>
    <div id="result"></div>
    <script>
      var resolvedMeeting = null;
      var mailingTouched = false;
      function el(id) { return document.getElementById(id); }
      function esc(v) {
        return String(v === null || v === undefined ? '' : v)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      }
      function choices() {
        var c = {};
        if (el('family').value) c.reportFamily = el('family').value;
        // Only a list the user typed is sent; otherwise the report follows its family.
        if (mailingTouched && el('mailingList').value.trim()) c.mailingList = el('mailingList').value.trim();
        return c;
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
        el('mailingRow').style.display = resolvedMeeting ? '' : 'none';
        if (r.mailingList) {
          if (!mailingTouched) el('mailingList').value = r.mailingList.effective;
          el('mailingList').disabled = !r.mailingListEditable;
          el('mailingHint').textContent = !r.mailingListEditable
            ? 'A main-meeting report always reads the list of its report family.'
            : (r.mailingList.overridden
              ? 'Your list will be used instead of the family default ' + r.mailingList.derived + '.'
              : 'Derived from the report family. Replace it if this meeting uses another list.');
        }
        el('title').textContent = r.title ? 'New report: ' + r.title : '';
        el('createBtn').disabled = !r.ok;
      }
      function showFailure(error) {
        el('lookupBtn').disabled = false;
        el('status').innerHTML = '<div class="err">' + esc(error && error.message ? error.message : error) + '</div>';
      }
      function lookUp() {
        resolvedMeeting = null;
        mailingTouched = false;
        el('mailingList').value = '';
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
      function mailingChanged() {
        // An emptied field goes back to the list derived from the family.
        if (!el('mailingList').value.trim()) mailingTouched = false;
        familyChanged();
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
            '<p>In the report: SA4 Report menu, then Report, then Build Report from Scratch (or Configure Meeting first). ' +
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
  const drafts = describeRevisionsFolderOutcome_(result.revisionsFolder);
  if (drafts) Logger.log('Report setup: ' + drafts);
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
      'Use ' + TEMPLATE_MENU_NAME_ + ' > ⚙️ Configure Meeting….', ui.ButtonSet.OK);
    return result;
  }
  const r = result.readiness || { ready: false, issues: [] };
  ui.alert('Report setup',
    (result.status === 'configured' ? '✅ Meeting configuration stored.' : 'ℹ️ This report is already set up.') + '\n\n' +
    (r.ready ? '✅ Ready to build: ' + TEMPLATE_MENU_NAME_ + ' > 📄 Report > Build Report from Scratch….'
      : '⚠️ Not ready to build yet:\n' + r.issues.map(function (i) { return '• ' + i.message; }).join('\n') +
        '\n\nOpen Configure Meeting to discover the missing sources.') +
    (describeRevisionsFolderOutcome_(result.revisionsFolder) ? '\n\n' + describeRevisionsFolderOutcome_(result.revisionsFolder) : '') +
    ((result.warnings || []).length ? '\n\n' + result.warnings.join('\n') : ''),
    ui.ButtonSet.OK);
  return result;
}

/**
 * TEMPLATE-003 (A): the one question before "Build Report from Scratch".
 * Called by runFullReportBuild() in Code.js, before anything is changed;
 * returns the button that was clicked. The build clears the document body
 * and writes a new report; only the Document Reallocations table is read
 * first and written back.
 */
function confirmTemplateBuildFromScratch_(ui) {
  // Ad-hoc attendance (stage D): imported attendance is stored outside the
  // document and is written again by the build, so the question says so.
  // Ad-hoc opening (stage E): the same for the opening details of the sessions.
  let kept = 'Only the Document Reallocations table is kept.';
  try {
    const also = [];
    if (adhocSessionsEnabled_()) {
      if ((adhocAttendanceSessionIds_(adhocAttendanceStore_()) || []).length) also.push('the imported attendance (with its Company cells)');
      if ((adhocOpeningSessionIds_(adhocOpeningStore_().getProperty(ADHOC_OPENING_KEY_)) || []).length) also.push('the opening details of the sessions');
    }
    if (also.length) {
      const all = ['the Document Reallocations table'].concat(also);
      kept = 'Only ' + all.slice(0, -1).join(', ') + ' and ' + all[all.length - 1] + ' are kept.';
    }
  } catch (e) {
    // The question is asked as for a report without attendance.
  }
  return ui.alert(
    'Build Report from Scratch',
    'This replaces the content of this document with a newly generated report.\n\n' +
    '⚠️ Everything that is in the document now is removed first. Meeting minutes and any other ' +
    'content entered by hand will be lost. ' + kept + '\n\n' +
    'To bring an existing report up to date without losing anything, answer No and use ' +
    'Report > Update Report Now.\n\n' +
    'The build then runs without further questions:\n' +
    '• report structure and TDoc tables\n' +
    '• e-mail discussions\n' +
    '• TDoc revisions\n' +
    '• abstracts (only if a Reviewer API token is configured)\n' +
    '• formatting\n\n' +
    'There is no completion message: the build has finished when "Running script" disappears.\n\n' +
    'Replace the content of this document and build the report from scratch?',
    ui.ButtonSet.YES_NO
  );
}

/**
 * TEMPLATE-003 (B), menu: "Update Report Now" -- the complete update, the
 * same work the Automatic Updates timer does (continuousUpdateCore_()): new
 * TDocs, statuses, revision placement, e-mail discussions, revisions,
 * abstracts if switched on.
 *
 * The timer keeps calling continuousUpdate(), which logs a failure and ends
 * normally. A person who clicked the item must see a failure, so this
 * wrapper throws instead: Docs shows the error and the execution is
 * "Failed". As in Full Build (Code.js 2.17.3) there is no UI call after the
 * document was changed: a completed update simply ends.
 *
 * It takes the same document lock as continuousUpdate(), so it never runs
 * together with an automatic update; when one is running, nothing is done
 * and that is reported as an error too.
 */
function updateReportNow() {
  assertNotTemplateMaster_();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('The report was not updated: another update of this report is running right now. Try again in a minute.');
  }
  let result;
  try {
    result = continuousUpdateCore_();
  } finally {
    lock.releaseLock();
  }
  const problem = describeUpdateReportNowFailure_(result);
  if (problem) {
    Logger.log('Update Report Now failed: ' + problem);
    throw new Error(problem);
  }
  return result;
}

/** The error text for a result of continuousUpdateCore_(), or '' when the update completed (pure). */
function describeUpdateReportNowFailure_(result) {
  if (!result || result.success !== true) {
    return 'The report update failed: ' + ((result && result.error) || 'no result was returned') +
      '\nThe report may have been updated in part. Nothing is lost; run Update Report Now again once the cause is fixed.';
  }
  const failures = result.collectorFailures || [];
  if (failures.length) {
    return 'The report update finished, but ' + (failures.length === 1 ? 'one part' : failures.length + ' parts') + ' failed:\n' +
      failures.map(function (f) { return '• ' + f.step + ': ' + f.error; }).join('\n') +
      '\nNew TDocs and statuses were updated. Run Update Report Now again later.';
  }
  return '';
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
    continuousInterval: active ? (stored ? stored.label : 'an unknown interval') : null,
    revisionsUrl: props.getProperty('REVISIONS_URL'),
    revisionsCandidate: currentRevisionsCandidate_(props.getProperty(TEMPLATE_STATE_KEYS_.revisionsCandidate), props.getProperty('FTP_BASE'))
  });
  ui.alert(isTemplateMasterDocument_() ? 'Template Release Info' : 'About This Report', lines.join('\n'), ui.ButtonSet.OK);
}

// ------------------------------------------------------------------
// The personal Reviewer API token: one private settings file in the
// user's own Google Drive (after T-2026.10.8)
// ------------------------------------------------------------------
//
// WHY. The Reviewer API token of a report is a Script Property of the
// report's own script project. Every report created from the template is a
// new project with empty property stores, so the token had to be entered
// again in every report -- and every editor of a report can read its Script
// Properties.
//
// WHAT. A user can keep the token once, for all their template reports, in
// one small JSON file in their own Google Drive:
//
//   { "schema": "sa4-report-user-settings/1", "reviewerApiToken": "<token>" }
//
// A token stored in a report still comes first (resolveReviewerApiToken_(),
// Code.js): report -> personal -> none. Nothing is ever copied from a report
// into the personal file without the user asking for it.
//
// HOW A NEW REPORT FINDS IT. Not by name, and not by anything stored in a
// report (a new report has nothing stored). The file carries a public Drive
// custom file property, sa4ReportUserSettings = 1, and is found by a query
// for that property among the files the user OWNS. (Drive's appProperties
// are private to one script project and would not be seen by the next
// report; User Properties are per script project as well.)
//
// WHAT IS TRUSTED. Only a file that is owned by the user, shared with
// nobody, not in the trash, of type application/json, at most 2 KB, and whose
// content is exactly the schema above. Each of these is checked here on
// every read, whatever the query returned.
//   - no such file:               no personal token;
//   - exactly one:                its token is used;
//   - more than one:              NONE is used. A secret is never picked
//                                 from several candidates; Configure Meeting
//                                 says so and the user resolves it;
//   - a marked file that is shared, or whose content is something else: not
//     used, and never changed or removed by a lookup.
//
// WHO. Every execution reads the Drive of the account it runs as: the person
// at the keyboard for a menu action, and -- for automatic updates -- the
// person who switched them on, whoever else has a personal token.
//
// The token is never logged, never written into the document or its
// properties, and never sent to the browser. The text of a JSON parse error
// is never logged either: it can quote what it failed to read.

var USER_SETTINGS_SCHEMA_ = 'sa4-report-user-settings/1';
var USER_SETTINGS_MARKER_KEY_ = 'sa4ReportUserSettings';
var USER_SETTINGS_MARKER_VALUE_ = '1';
var USER_SETTINGS_FILE_NAME_ = 'SA4 Report – private settings (do not share).json';
var USER_SETTINGS_MIME_ = 'application/json';
var USER_SETTINGS_MAX_BYTES_ = 2048;
var USER_SETTINGS_TOKEN_MAX_CHARS_ = 512;
// Owner-restricted, and by the marker -- never by the name of the file.
var USER_SETTINGS_QUERY_ = "'me' in owners and trashed = false and mimeType = 'application/json' and properties has { key='sa4ReportUserSettings' and value='1' }";
var USER_SETTINGS_FILE_FIELDS_ = 'id,name,mimeType,size,ownedByMe,shared,trashed,properties';

/**
 * Pure: a token as it can be stored and sent -- text without space around
 * it, 1 to 512 characters, no line break or other control character (it is
 * sent as an HTTP header). Nothing else about its form is assumed: a report
 * stores whatever was entered, trimmed, and so does this.
 */
function isUsableReviewerTokenValue_(value) {
  return typeof value === 'string' && value.length >= 1 && value.length <= USER_SETTINGS_TOKEN_MAX_CHARS_ &&
    value === value.trim() && !/[\u0000-\u001F\u007F]/.test(value);
}

/** Pure: the content of the settings file for a token. */
function serializeUserSettings_(token) {
  return JSON.stringify({ schema: USER_SETTINGS_SCHEMA_, reviewerApiToken: token });
}

/**
 * Pure: the token in the content of a settings file, or '' when the content
 * is anything but exactly the schema. Never throws, and gives no reason:
 * what could not be read is not repeated anywhere.
 */
function parseUserSettingsText_(text) {
  if (typeof text !== 'string' || utf8ByteLength_(text) > USER_SETTINGS_MAX_BYTES_) return '';
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (e) {
    return '';
  }
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return '';
  const keys = Object.keys(parsed).sort();
  if (keys.length !== 2 || keys[0] !== 'reviewerApiToken' || keys[1] !== 'schema') return '';
  if (parsed.schema !== USER_SETTINGS_SCHEMA_) return '';
  return isUsableReviewerTokenValue_(parsed.reviewerApiToken) ? parsed.reviewerApiToken : '';
}

/**
 * Pure: what a file of Drive is to the personal settings, from its metadata.
 *   'private'  owned by the user, shared with nobody, not in the trash,
 *              application/json, and marked: it may be read;
 *   'shared'   the same, but shared (or not known to be unshared): never read;
 *   'other'    anything else -- not the user's own, in the trash, another
 *              type, not marked: as if it were not there.
 * Every value must be exactly what is required; a missing one does not pass.
 */
function classifyUserSettingsFile_(meta) {
  if (!meta || typeof meta !== 'object' || !meta.id) return 'other';
  if (meta.ownedByMe !== true || meta.trashed !== false || meta.mimeType !== USER_SETTINGS_MIME_) return 'other';
  if (!meta.properties || meta.properties[USER_SETTINGS_MARKER_KEY_] !== USER_SETTINGS_MARKER_VALUE_) return 'other';
  return meta.shared === false ? 'private' : 'shared';
}

/**
 * The Drive calls of the personal settings, in one place (the functions
 * below take them as `drive`, so that they are tested with a fake Drive).
 */
function userSettingsDrive_() {
  return {
    list: function () {
      const found = Drive.Files.list({ q: USER_SETTINGS_QUERY_, fields: 'files(' + USER_SETTINGS_FILE_FIELDS_ + ')', pageSize: 50, spaces: 'drive' });
      return (found && found.files) || [];
    },
    meta: function (id) { return Drive.Files.get(id, { fields: USER_SETTINGS_FILE_FIELDS_ }); },
    read: function (id) { return DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'); },
    create: function (text) {
      const properties = {};
      properties[USER_SETTINGS_MARKER_KEY_] = USER_SETTINGS_MARKER_VALUE_;
      const made = Drive.Files.create({ name: USER_SETTINGS_FILE_NAME_, mimeType: USER_SETTINGS_MIME_, properties: properties },
        Utilities.newBlob(text, USER_SETTINGS_MIME_, USER_SETTINGS_FILE_NAME_), { fields: 'id' });
      return made.id;
    },
    write: function (id, text) { DriveApp.getFileById(id).setContent(text); },
    trash: function (id) { DriveApp.getFileById(id).setTrashed(true); }
  };
}

/**
 * Looks at the user's personal settings. Returns
 *   { status, token, validIds, privateIds, sharedCount }
 * status: 'ok' (exactly one usable file; `token` is its token), 'none',
 * 'ambiguous' (more than one usable file: no token), 'shared' (none usable,
 * and a marked file is shared), 'corrupt' (none usable, and a private marked
 * file has other content), 'unreadable' (Drive could not be read: no token).
 * `token` is '' unless the status is 'ok'. Never throws. What is logged names
 * no content of any file.
 */
function inspectPersonalReviewerSettingsWith_(drive) {
  const out = { status: 'none', token: '', validIds: [], privateIds: [], sharedCount: 0 };
  let files;
  try {
    files = drive.list() || [];
  } catch (e) {
    Logger.log('Personal settings: Google Drive could not be searched (' + e.message + ').');
    out.status = 'unreadable';
    return out;
  }
  const tokens = [];
  let unreadable = false;
  files.forEach(function (meta) {
    const kind = classifyUserSettingsFile_(meta);
    if (kind === 'shared') out.sharedCount++;
    if (kind !== 'private') return;
    out.privateIds.push(meta.id);
    if (Number(meta.size) > USER_SETTINGS_MAX_BYTES_) return;       // too large to be the settings: not read
    let text;
    try {
      text = drive.read(meta.id);
    } catch (e) {
      Logger.log('Personal settings: a settings file could not be read.');
      unreadable = true;
      return;
    }
    const token = parseUserSettingsText_(text);
    if (token) { out.validIds.push(meta.id); tokens.push(token); }
  });
  if (unreadable) out.status = 'unreadable';
  else if (out.validIds.length === 1) { out.status = 'ok'; out.token = tokens[0]; }
  else if (out.validIds.length > 1) out.status = 'ambiguous';
  else if (out.sharedCount > 0) out.status = 'shared';
  else if (out.privateIds.length > 0) out.status = 'corrupt';
  return out;
}

/** Pure: why a personal token is not in use, in a few fixed words for the log ('' when there is nothing to say). */
function personalTokenProblemForLog_(status) {
  if (status === 'ambiguous') return 'more than one private settings file exists, so none is used';
  if (status === 'shared') return 'the settings file is shared';
  if (status === 'corrupt') return 'the settings file has other content than expected';
  if (status === 'unreadable') return 'Google Drive could not be read';
  return '';
}

/**
 * The personal token for resolveReviewerApiToken_() (Code.js): { token,
 * problem } -- `problem` is '' or the fixed words above. Never throws: with
 * anything wrong there is no personal token, and the report works without.
 */
function readPersonalReviewerTokenSafely_() {
  try {
    const found = inspectPersonalReviewerSettingsWith_(userSettingsDrive_());
    return { token: found.status === 'ok' ? found.token : '', problem: personalTokenProblemForLog_(found.status) };
  } catch (e) {
    return { token: '', problem: 'Google Drive could not be read' };
  }
}

/** Pure: whether the file of this metadata and content is a private settings file holding exactly this token. */
function isVerifiedUserSettingsFile_(meta, text, token) {
  return classifyUserSettingsFile_(meta) === 'private' && parseUserSettingsText_(text) === token && token !== '';
}

/**
 * Stores `token` as the personal token. Returns { ok, reason, sharedLeft }.
 *   - no private settings file: one is created;
 *   - one usable file: it is rewritten;
 *   - one private marked file with other content, and no usable one: it is
 *     rewritten (it is the user's own, private, and marked as this);
 *   - several: nothing is written -- which one would be a guess, and another
 *     one would be a duplicate.
 * A marked file that is shared is left exactly as it is, and never written to.
 * Success is reported only after the written file was read back by its id
 * and found private and holding the token.
 */
function writePersonalReviewerTokenWith_(drive, token) {
  const fail = function (reason) { return { ok: false, reason: reason, sharedLeft: 0 }; };
  if (!isUsableReviewerTokenValue_(token)) return fail('The token is empty, longer than ' + USER_SETTINGS_TOKEN_MAX_CHARS_ + ' characters, or contains a line break.');
  const text = serializeUserSettings_(token);
  const before = inspectPersonalReviewerSettingsWith_(drive);
  if (before.status === 'unreadable') return fail('Google Drive could not be read.');
  if (before.validIds.length > 1 || (before.validIds.length === 0 && before.privateIds.length > 1)) {
    return fail('Your Google Drive has more than one private settings file ("' + USER_SETTINGS_FILE_NAME_ + '"). Remove the personal token first, or delete the extra files in Drive.');
  }
  let id = before.validIds.length === 1 ? before.validIds[0] : (before.privateIds.length === 1 ? before.privateIds[0] : null);
  try {
    if (id) drive.write(id, text);
    else id = drive.create(text);
  } catch (e) {
    Logger.log('Personal settings: the settings file could not be written (' + e.message + ').');
    return fail('The settings file could not be written to Google Drive.');
  }
  let verified = false;
  try {
    verified = !!id && isVerifiedUserSettingsFile_(drive.meta(id), drive.read(id), token);
  } catch (e) {
    Logger.log('Personal settings: the written settings file could not be read back.');
  }
  if (!verified) return fail('The settings file was written but could not be confirmed as private and complete. The personal token is not counted as stored.');
  return { ok: true, reason: '', sharedLeft: before.sharedCount };
}

/**
 * "Remove my personal token": every PRIVATE marked settings file is moved to
 * the trash (recoverable) -- also when there are several, which is how that
 * situation is cleared. A shared one is not touched. Returns { ok, trashed,
 * sharedLeft, reason }.
 */
function clearPersonalReviewerTokenWith_(drive) {
  const before = inspectPersonalReviewerSettingsWith_(drive);
  if (before.status === 'unreadable') return { ok: false, trashed: 0, sharedLeft: 0, reason: 'Google Drive could not be read.' };
  let trashed = 0;
  try {
    before.privateIds.forEach(function (id) { drive.trash(id); trashed++; });
  } catch (e) {
    Logger.log('Personal settings: a settings file could not be moved to the trash (' + e.message + ').');
    return { ok: false, trashed: trashed, sharedLeft: before.sharedCount, reason: 'A settings file could not be moved to the trash.' };
  }
  return { ok: true, trashed: trashed, sharedLeft: before.sharedCount, reason: '' };
}

/** Pure: what Configure Meeting says about the token in use. `hasLocal`: the report has a token of its own; `personal`: { status, sharedCount }. */
function describeReviewerTokenStatus_(hasLocal, personal) {
  const p = personal || { status: 'none', sharedCount: 0 };
  const name = '"' + USER_SETTINGS_FILE_NAME_ + '"';
  let problem = '';
  if (p.status === 'ambiguous') {
    problem = 'Your Google Drive has more than one private settings file (' + name + '), so none of them is used. ' +
      'Tick "Remove my personal token" to move them all to the trash and save the token again, or delete the extra files in Drive.';
  } else if (p.status === 'shared') {
    problem = 'Your personal settings file (' + name + ') is shared with others, so its token is not used. Treat that token as disclosed: delete the file in Drive, get a new token and save it.';
  } else if (p.status === 'corrupt') {
    problem = 'Your personal settings file (' + name + ') has other content than expected and is not used. Save the token again to replace it.';
  } else if (p.status === 'unreadable') {
    problem = 'Your personal token could not be checked: Google Drive could not be read.';
  } else if (p.status === 'ok' && p.sharedCount > 0) {
    problem = 'Another settings file in your Drive (' + name + ') is shared and is not used; delete it.';
  }
  let text;
  if (hasLocal) {
    text = p.status === 'ok'
      ? 'Using the Reviewer API token stored in this report. Your personal token is not used here.'
      : 'Using the Reviewer API token stored in this report. Editors of this report can see it.';
  } else {
    text = p.status === 'ok'
      ? 'Using your personal Reviewer API token (kept privately in your Google Drive, for all your reports).'
      : 'No Reviewer API token. Abstracts are skipped.';
  }
  return problem ? text + ' ' + problem : text;
}

/**
 * The token part of Configure Meeting in a template report: { status,
 * fieldsHtml }. Only words and controls -- no token, and nothing derived
 * from one. Called by configureMeetingSettings() (Code.js). Never throws.
 */
function templateReviewerTokenDialogParts_() {
  let hasLocal = false;
  let personal = { status: 'unreadable', sharedCount: 0 };
  try {
    hasLocal = resolveReviewerApiToken_(true).source === 'report';
    const found = inspectPersonalReviewerSettingsWith_(userSettingsDrive_());
    personal = { status: found.status, sharedCount: found.sharedCount };
  } catch (e) {
    Logger.log('Configure Meeting: the personal Reviewer token could not be looked at.');
  }
  const fieldsHtml =
    '<label>New Reviewer API token:</label>\n' +
    '        <input type="password" id="apiToken" value="" placeholder="Enter a token to store it" autocomplete="off">\n' +
    '        <label><input type="radio" name="apiTokenTarget" id="apiTokenTargetPersonal" checked> Store it for all my reports (privately, in my Google Drive)</label>\n' +
    '        <label><input type="radio" name="apiTokenTarget" id="apiTokenTargetReport"> Store it in this report only (its editors can see it)</label>\n' +
    (hasLocal ? '        <label><input type="checkbox" id="promoteApiToken"> Save the token of this report as my personal token</label>\n' : '') +
    '        <label><input type="checkbox" id="clearApiToken"> Remove the token of this report</label>\n' +
    '        <label><input type="checkbox" id="clearPersonalApiToken"> Remove my personal token</label>\n' +
    '        <div class="hint">Used to fetch AI summaries and abstracts. A stored token is never displayed. A token stored in this report is used before the personal one. ' +
    'Automatic updates run as the person who switched them on, and use the tokens of that person.</div>';
  return { status: describeReviewerTokenStatus_(hasLocal, personal), fieldsHtml: fieldsHtml };
}

/**
 * Called by saveConfigurationSettings() (Code.js) in the template runtime,
 * BEFORE the configuration is stored. Takes what the dialog sent about the
 * personal token out of the configuration and returns the plan for it:
 *   { config, personalAction, personalToken, clearLocalAfter }
 * `config` is what is stored as before (it never contains the personal
 * token). When the personal token is to be written AND the token of this
 * report is to be removed, the removal is taken out of `config` and done
 * afterwards, once the personal file is written and confirmed
 * (afterTemplateConfigurationSaved_()): a report never loses its working
 * token to a write that did not succeed.
 * Throws -- before anything is stored -- for an action that does not exist
 * and for a new personal token that cannot be used.
 */
function beforeTemplateConfigurationSaved_(config) {
  const copy = Object.assign({}, config || {});
  const action = String(copy.personalTokenAction === undefined || copy.personalTokenAction === null || copy.personalTokenAction === '' ? 'keep' : copy.personalTokenAction);
  const token = action === 'replace' ? String(copy.personalToken === undefined || copy.personalToken === null ? '' : copy.personalToken).trim() : '';
  delete copy.personalTokenAction;
  delete copy.personalToken;
  if (['keep', 'replace', 'clear', 'promote'].indexOf(action) === -1) throw new Error('Unknown action for the personal Reviewer token.');
  if (action === 'replace' && !isUsableReviewerTokenValue_(token)) {
    throw new Error('The Reviewer token cannot be stored: it is empty, longer than ' + USER_SETTINGS_TOKEN_MAX_CHARS_ + ' characters, or contains a line break.');
  }
  const plan = { config: copy, personalAction: action, personalToken: token, clearLocalAfter: false };
  if ((action === 'replace' || action === 'promote') && copy.apiTokenAction === 'clear') {
    copy.apiTokenAction = 'keep';
    plan.clearLocalAfter = true;
  }
  return plan;
}

/**
 * Carries out the plan for the personal token. Returns the text for the
 * message of the dialog ('' when nothing was asked for). Never throws.
 *
 * deps: { drive, localToken() -> the token stored in this report or '',
 * removeLocalToken() }.
 */
function applyPersonalReviewerTokenPlanWith_(deps, plan) {
  if (!plan || plan.personalAction === 'keep' || !plan.personalAction) return '';
  try {
    if (plan.personalAction === 'clear') {
      const cleared = clearPersonalReviewerTokenWith_(deps.drive);
      if (!cleared.ok) return '⚠️ Your personal Reviewer token was NOT removed: ' + cleared.reason;
      return (cleared.trashed ? '✅ Your personal Reviewer token was removed (' + cleared.trashed + ' settings file' + (cleared.trashed === 1 ? '' : 's') + ' moved to the trash of your Google Drive).'
        : 'ℹ️ There was no personal Reviewer token to remove.') +
        (cleared.sharedLeft ? ' A settings file that is shared was left as it is; delete it in Drive.' : '');
    }
    let token = plan.personalToken;
    if (plan.personalAction === 'promote') {
      token = deps.localToken();
      if (!token) return '⚠️ Nothing was saved as your personal Reviewer token: this report has no token of its own.';
    }
    const written = writePersonalReviewerTokenWith_(deps.drive, token);
    if (!written.ok) {
      return '⚠️ Your personal Reviewer token was NOT saved: ' + written.reason + (plan.clearLocalAfter ? ' The token of this report was kept.' : '');
    }
    let note = '✅ Your personal Reviewer token is saved privately in your Google Drive ("' + USER_SETTINGS_FILE_NAME_ + '") and is used in all your template reports.';
    if (plan.clearLocalAfter) {
      deps.removeLocalToken();
      note += ' The token of this report was removed.';
    }
    if (written.sharedLeft) note += ' Another settings file in your Drive is shared and is not used; delete it.';
    return note;
  } catch (e) {
    Logger.log('Configure Meeting: the personal Reviewer token could not be handled.');
    return '⚠️ Your personal Reviewer token could not be changed. The token of this report is as it was.';
  }
}

/** The same with the live services. Never throws. */
function applyPersonalReviewerTokenPlanSafely_(plan) {
  try {
    const note = applyPersonalReviewerTokenPlanWith_({
      drive: userSettingsDrive_(),
      // Through the one resolver, after the save: the token of this report, when it has one.
      localToken: function () { resetReviewerTokenRunState_(); return resolveReviewerApiToken_(true).token; },
      removeLocalToken: function () { PropertiesService.getScriptProperties().deleteProperty('REVIEWER_API_TOKEN'); }
    }, plan);
    resetReviewerTokenRunState_();
    return note;
  } catch (e) {
    Logger.log('Configure Meeting: the personal Reviewer token could not be handled.');
    return '';
  }
}
