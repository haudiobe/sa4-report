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

// ------------------------------------------------------------------
// Shared Minutes: a separate Google Doc that anyone with the link can
// edit, created for one report and linked from it
// ------------------------------------------------------------------
//
// WHAT. The report of a meeting says where its shared minutes are ("Link to
// the shared minutes:"). Until now that document was made and shared by
// hand, and its address typed into the report. Configure Meeting has a
// button for it now: Create Shared Minutes makes one Google Doc -- its title
// and one line naming the meeting, nothing else --, gives ANYONE WITH THE
// LINK the right to EDIT it, and writes the link into the report.
//
// WHEN. Only when that button is pressed. Opening a report, opening or
// saving Configure Meeting, an update (by hand or by a trigger), a build and
// the discussion e-mails create nothing, share nothing and ask Drive
// nothing about it: a build writes the link again from what is stored, and
// nothing else here runs in them.
//
// WHAT IS STORED. One Document Property of the report, SHARED_DOCUMENT:
//
//   { "v": 1, "fileId", "url", "name", "reportDocumentId", "meetingId",
//     "createdAt", "permissionVerifiedAt" }
//
// It is used only in the document it names and for the meeting it names. A
// copy of a report starts without properties (TEMPLATE-001); a value that
// names another document is ignored all the same. permissionVerifiedAt is
// null until the permission "anyone, writer" was READ BACK from Drive; the
// link is written into the report only then.
//
// ONE DOCUMENT, ALSO AFTER A CRASH. The document is created with Drive
// appProperties that say what it is (sa4Purpose, sa4ReportId, sa4MeetingId)
// -- in the request that creates it, so an unmarked one never exists. They
// are private to the script project that wrote them, i.e. to this report: a
// copy of the report cannot see them. Before anything is created Drive is
// asked for a document so marked:
//   - exactly one:   it is the document. It is recorded and used;
//   - more than one: NOTHING is created, shared or recorded. The user is
//                    shown them and resolves it;
//   - none:          one is created -- unless a creation was started less
//                    than two minutes ago (SHARED_DOCUMENT_PENDING): what
//                    Drive finds can lag behind what it has, so the answer
//                    "none" is not trusted that soon.
// The property is stored directly after the creation, before the document
// is shared; pressing the button again goes on from wherever it stopped.
//
// Like the personal settings above: pure functions, *With_(deps) functions
// with every service injected, and the public functions the dialog calls.

var SHARED_MINUTES_KEY_ = 'SHARED_DOCUMENT';
var SHARED_MINUTES_PENDING_KEY_ = 'SHARED_DOCUMENT_PENDING';
var SHARED_MINUTES_SCHEMA_VERSION_ = 1;
var SHARED_MINUTES_LABEL_ = 'Link to the shared minutes:';
var SHARED_MINUTES_PURPOSE_ = 'shared-minutes';
var SHARED_MINUTES_MIME_ = 'application/vnd.google-apps.document';
var SHARED_MINUTES_PENDING_MS_ = 120000;
var SHARED_MINUTES_FILE_FIELDS_ = 'id,name,mimeType,trashed,appProperties,createdTime';

/** Pure: the address of the document, in the one form that is stored and linked. */
function sharedMinutesUrl_(fileId) {
  return 'https://docs.google.com/document/d/' + fileId + '/edit';
}

/**
 * Pure: the name of the document, from the title the report is given
 * (generateReportTitle_()): its first "Minutes" becomes "Shared Minutes" --
 * "Video SWG Shared Minutes SA4#137-e". The name says what the document is;
 * it is never what the document is found by.
 */
function sharedMinutesName_(reportTitle) {
  const title = trimmed_(reportTitle);
  if (/\bMinutes\b/.test(title)) return title.replace(/\bMinutes\b/, 'Shared Minutes');
  return title ? 'Shared Minutes – ' + title : 'Shared Minutes';
}

/** Pure: what the new document says -- its title, and one line naming the meeting when the report knows it. */
function sharedMinutesHtml_(name, meetingName, meetingDate) {
  const esc = function (v) { return trimmed_(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const meeting = [esc(meetingName), esc(meetingDate)].filter(Boolean).join(', ');
  return '<html><head><meta charset="utf-8"></head><body><h1>' + esc(name) + '</h1>' + (meeting ? '<p>' + meeting + '</p>' : '') + '</body></html>';
}

/** Pure: the marker of the document of one report and one meeting. */
function sharedMinutesAppProperties_(documentId, meetingId) {
  return { sa4Purpose: SHARED_MINUTES_PURPOSE_, sa4ReportId: String(documentId), sa4MeetingId: String(meetingId) };
}

/** Pure: the one query for the marked documents of a report. The id is a Drive id: it has nothing to escape. */
function sharedMinutesQuery_(documentId) {
  return "appProperties has { key='sa4ReportId' and value='" + documentId + "' } and appProperties has { key='sa4Purpose' and value='" + SHARED_MINUTES_PURPOSE_ + "' }" +
    " and mimeType = '" + SHARED_MINUTES_MIME_ + "' and trashed = false";
}

/** Pure: whether a file, as Drive describes it, is the shared minutes of this report and meeting. Checked for every file, whatever a query returned. */
function isSharedMinutesFile_(meta, documentId, meetingId) {
  if (!meta || typeof meta !== 'object' || !isPlausibleDriveId_(meta.id)) return false;
  if (meta.mimeType !== SHARED_MINUTES_MIME_ || meta.trashed !== false) return false;
  const marker = meta.appProperties;
  return !!marker && marker.sa4Purpose === SHARED_MINUTES_PURPOSE_ && marker.sa4ReportId === String(documentId) && marker.sa4MeetingId === String(meetingId);
}

/** Pure: whether the permissions of a file, as Drive lists them, give anyone with the link the right to edit. */
function hasAnyoneWriterPermission_(permissions) {
  return Array.isArray(permissions) && permissions.some(function (p) { return !!p && p.type === 'anyone' && p.role === 'writer'; });
}

/** Pure: the one stored form, fields in a fixed order. */
function serializeSharedMinutes_(record) {
  return JSON.stringify({
    v: SHARED_MINUTES_SCHEMA_VERSION_, fileId: record.fileId, url: record.url, name: record.name, reportDocumentId: record.reportDocumentId,
    meetingId: record.meetingId, createdAt: record.createdAt, permissionVerifiedAt: record.permissionVerifiedAt || null
  });
}

/**
 * Pure: reads a SHARED_DOCUMENT value for the document and the meeting it is
 * asked for. { status, record }:
 *   'absent'        nothing stored;
 *   'invalid'       not what this code stores -- not used;
 *   'unsupported'   written by a newer release -- not used, not changed;
 *   'other-report'  names another document (a copied value) -- not used;
 *   'other-meeting' of this document, for another meeting -- not used;
 *   'ok'            the shared minutes of this report.
 * `record` is given for the last two. Never throws.
 */
function readSharedMinutes_(raw, documentId, meetingId) {
  const none = function (status) { return { status: status, record: null }; };
  if (raw === null || raw === undefined || trimmed_(raw) === '') return none('absent');
  let data;
  try { data = JSON.parse(String(raw)); } catch (e) { data = null; }
  if (!data || typeof data !== 'object' || Array.isArray(data)) return none('invalid');
  if (data.v !== SHARED_MINUTES_SCHEMA_VERSION_) return none(typeof data.v === 'number' && data.v > SHARED_MINUTES_SCHEMA_VERSION_ ? 'unsupported' : 'invalid');
  const text = function (v) { return typeof v === 'string' && v.trim() !== '' && v.length <= 500; };
  if (!isPlausibleDriveId_(data.fileId) || data.url !== sharedMinutesUrl_(data.fileId) || !text(data.name) || !text(data.reportDocumentId) ||
      !/^\d+$/.test(String(data.meetingId === undefined ? '' : data.meetingId)) || typeof data.meetingId !== 'string' || !text(data.createdAt) ||
      !(data.permissionVerifiedAt === null || text(data.permissionVerifiedAt))) return none('invalid');
  const record = { fileId: data.fileId, url: data.url, name: data.name, reportDocumentId: data.reportDocumentId, meetingId: data.meetingId, createdAt: data.createdAt, permissionVerifiedAt: data.permissionVerifiedAt };
  if (record.reportDocumentId !== String(documentId)) return none('other-report');
  if (record.meetingId !== trimmed_(meetingId)) return { status: 'other-meeting', record: record };
  return { status: 'ok', record: record };
}

/**
 * The Drive calls of the shared minutes, in one place (the functions below
 * take them as `drive`, so that they are tested with a fake Drive). The only
 * calls in the project that create a document for other people or change
 * who can open a file -- and only createSharedMinutesWith_() makes them.
 */
function sharedMinutesDrive_() {
  return {
    search: function (documentId) {
      const found = Drive.Files.list({ q: sharedMinutesQuery_(documentId), fields: 'files(' + SHARED_MINUTES_FILE_FIELDS_ + ')', pageSize: 20, includeItemsFromAllDrives: true, supportsAllDrives: true });
      return (found && found.files) || [];
    },
    meta: function (id) { return Drive.Files.get(id, { fields: SHARED_MINUTES_FILE_FIELDS_, supportsAllDrives: true }); },
    /** The folder the report is in: { id, name, writable }, or null when it has none the user can see. */
    reportFolder: function (documentId) {
      const report = Drive.Files.get(documentId, { fields: 'parents', supportsAllDrives: true });
      const parentId = report && report.parents && report.parents[0];
      if (!parentId) return null;
      const folder = Drive.Files.get(parentId, { fields: 'id,name,capabilities(canAddChildren)', supportsAllDrives: true });
      return { id: folder.id, name: folder.name, writable: !!(folder.capabilities && folder.capabilities.canAddChildren) };
    },
    /** One request: name, place, marker and content together. Without a folder the document goes to the user's My Drive. */
    create: function (name, folderId, appProperties, html) {
      const resource = { name: name, mimeType: SHARED_MINUTES_MIME_, appProperties: appProperties };
      if (folderId) resource.parents = [folderId];
      return Drive.Files.create(resource, Utilities.newBlob(html, 'text/html', name + '.html'), { fields: SHARED_MINUTES_FILE_FIELDS_, supportsAllDrives: true });
    },
    permissions: function (id) {
      const all = [];
      let pageToken = null;
      for (let page = 0; page < 20; page++) {
        const args = { fields: 'nextPageToken,permissions(id,type,role)', pageSize: 100, supportsAllDrives: true };
        if (pageToken) args.pageToken = pageToken;
        const found = Drive.Permissions.list(id, args) || {};
        (found.permissions || []).forEach(function (p) { all.push(p); });
        pageToken = found.nextPageToken || null;
        if (!pageToken) break;
      }
      return all;
    },
    /** Anyone with the link can edit. What Drive answers is not looked at: the permission is read back with permissions(). */
    shareWithAnyoneAsWriter: function (id) {
      Drive.Permissions.create({ type: 'anyone', role: 'writer', allowFileDiscovery: false }, id, { fields: 'id', supportsAllDrives: true });
    }
  };
}

/** What is stored for this document and its meeting now: { documentId, meetingId, status, record }. */
function currentSharedMinutesWith_(deps) {
  const documentId = deps.activeDocumentId();
  const props = deps.documentProperties;
  const meetingId = trimmed_(props.getProperty('MEETING_ID'));
  const stored = readSharedMinutes_(props.getProperty(SHARED_MINUTES_KEY_), documentId, meetingId);
  return { documentId: documentId, meetingId: meetingId, status: stored.status, record: stored.record };
}

/** Writes the link of a verified record into the report. Returns '' or what to tell the user; never throws. */
function writeSharedMinutesLinkWith_(deps, record) {
  try {
    const rendered = renderSharedMinutesLink_(deps.body(), record);
    return rendered.line === 'no-place' ? 'The report has no opening section yet: the link is written when the report is built.' : '';
  } catch (e) {
    Logger.log('Shared minutes: the link could not be written into the report: ' + e.message);
    return 'The link could not be written into the report (' + e.message + '). Press Verify to write it, or build the report.';
  }
}

/**
 * Create Shared Minutes -- the one action that creates and shares. It is
 * also the action that finishes what an earlier one began: every step looks
 * at what is there first, so pressing the button again never makes a second
 * document and never shares twice.
 *
 * deps: { release, activeDocumentId(), nowIso(), nowMs(), documentProperties,
 * drive (sharedMinutesDrive_()), describe() -> { name, meetingName,
 * meetingDate }, body() }.
 * Returns { ok, state, message, created, adopted, record }; `state` is
 * 'complete' | 'sharing-failed' | 'ambiguous' | 'refused'. Never throws.
 */
function createSharedMinutesWith_(deps) {
  const refused = function (message, state) { return { ok: false, state: state || 'refused', message: message, created: false, adopted: false, record: null }; };
  if (templateDocumentRole_(deps.activeDocumentId(), deps.release) === 'template') return refused('This is the SA4 Report Template itself: shared minutes belong to a report.');
  const props = deps.documentProperties;
  const drive = deps.drive;
  const current = currentSharedMinutesWith_(deps);
  const documentId = current.documentId;
  const meetingId = current.meetingId;
  if (!isPlausibleDriveId_(documentId)) return refused('Nothing was created: this document could not be identified.');
  if (!/^\d+$/.test(meetingId)) return refused('Nothing was created: this report has no meeting yet. Resolve the meeting and save the configuration first.');
  if (current.status === 'unsupported') return refused('Nothing was created: the shared minutes of this report were recorded by a newer release and are not changed here.');
  if (current.status === 'other-meeting') {
    return refused('Nothing was created: this report has the shared minutes of meeting ' + current.record.meetingId + ' recorded ("' + current.record.name + '"), and is configured for meeting ' + meetingId +
      ' now. Use "Forget" first; the document of meeting ' + current.record.meetingId + ' is not changed by it.');
  }

  let record = current.status === 'ok' ? current.record : null;
  let created = false;
  let adopted = false;
  let placeNote = '';

  if (record) {
    // Recorded already: it must still be the document that was created.
    let meta;
    try {
      meta = drive.meta(record.fileId);
    } catch (e) {
      return refused('Nothing was created or shared: the recorded Shared Minutes document "' + record.name + '" could not be opened (' + e.message + '). If it was deleted, use "Forget" and create a new one.');
    }
    if (!isSharedMinutesFile_(meta, documentId, meetingId)) {
      return refused('Nothing was created or shared: the recorded Shared Minutes document "' + record.name + '" is in the trash, or is not the document that was created for this report. Restore it, or use "Forget" and create a new one.');
    }
  } else {
    // Not recorded. Is there one already -- from an action that did not get as far as recording it?
    let found;
    try {
      found = (drive.search(documentId) || []).filter(function (meta) { return isSharedMinutesFile_(meta, documentId, meetingId); });
    } catch (e) {
      return refused('Nothing was created: Google Drive could not be searched for an existing Shared Minutes document (' + e.message + '). Try again.');
    }
    if (found.length > 1) {
      return refused('Nothing was created, shared or recorded: Google Drive has ' + found.length + ' Shared Minutes documents for this report and meeting, and none is picked for you:\n' +
        found.map(function (meta) { return '• ' + meta.name + ' — ' + sharedMinutesUrl_(meta.id); }).join('\n') +
        '\nMove all but one to the trash, then press the button again.', 'ambiguous');
    }
    let meta;
    if (found.length === 1) {
      meta = found[0];
      adopted = true;
    } else {
      const now = deps.nowMs();
      let pendingSince = NaN;
      try { pendingSince = Number(JSON.parse(String(props.getProperty(SHARED_MINUTES_PENDING_KEY_))).startedAt); } catch (e) { pendingSince = NaN; }
      if (isFinite(pendingSince) && now >= pendingSince && now - pendingSince < SHARED_MINUTES_PENDING_MS_) {
        return refused('Nothing was created: a creation was started ' + Math.max(1, Math.round((now - pendingSince) / 1000)) + ' seconds ago, and Google Drive may not show its document yet. ' +
          'Press the button again in ' + Math.ceil((SHARED_MINUTES_PENDING_MS_ - (now - pendingSince)) / 1000) + ' seconds: the document is used if it exists, and created if it does not.');
      }
      let description;
      try {
        description = deps.describe();
        props.setProperty(SHARED_MINUTES_PENDING_KEY_, JSON.stringify({ startedAt: now }));
      } catch (e) {
        return refused('Nothing was created: the creation could not be prepared (' + e.message + ').');
      }
      // The folder of the report when the user can add to it; the user's My Drive otherwise. Never a reason to fail.
      let folder = null;
      try { folder = drive.reportFolder(documentId); } catch (e) { folder = null; }
      const folderId = folder && folder.writable ? folder.id : '';
      if (!folderId) placeNote = 'It is in your My Drive: the folder of the report could not be used.';
      try {
        meta = drive.create(description.name, folderId, sharedMinutesAppProperties_(documentId, meetingId), sharedMinutesHtml_(description.name, description.meetingName, description.meetingDate));
      } catch (e) {
        return refused('The Shared Minutes document could not be created (' + e.message + '). Nothing was shared or recorded. Press the button again in two minutes: should Google Drive have created it after all, it is found and used.');
      }
      if (!meta || !isPlausibleDriveId_(meta.id)) return refused('Google Drive did not say which document it created. Nothing was shared or recorded. Press the button again in two minutes: the document is found and used if it exists.');
      created = true;
    }
    record = { fileId: meta.id, url: sharedMinutesUrl_(meta.id), name: trimmed_(meta.name) || 'Shared Minutes', reportDocumentId: documentId, meetingId: meetingId,
      createdAt: created ? deps.nowIso() : (trimmed_(meta.createdTime) || deps.nowIso()), permissionVerifiedAt: null };
    // Recorded now, before it is shared: whatever happens next, the report knows its document.
    try {
      props.setProperty(SHARED_MINUTES_KEY_, serializeSharedMinutes_(record));
    } catch (e) {
      return refused('The Shared Minutes document "' + record.name + '" exists, but could not be recorded in this report (' + e.message + '), and was not shared. Press the button again: it is found and used, not created a second time.');
    }
  }

  // Anyone with the link can edit -- added only when it is not there, and believed only when it was read back.
  let shareError = '';
  let permissions = null;
  try {
    permissions = drive.permissions(record.fileId);
    if (!hasAnyoneWriterPermission_(permissions)) {
      try { drive.shareWithAnyoneAsWriter(record.fileId); } catch (e) { shareError = e.message; }
      permissions = drive.permissions(record.fileId);
    }
  } catch (e) {
    permissions = null;
    shareError = shareError || e.message;
  }
  if (!hasAnyoneWriterPermission_(permissions)) {
    if (record.permissionVerifiedAt) {
      record.permissionVerifiedAt = null;
      try { props.setProperty(SHARED_MINUTES_KEY_, serializeSharedMinutes_(record)); } catch (e) { Logger.log('Shared minutes: the record could not be updated: ' + e.message); }
    }
    return { ok: false, state: 'sharing-failed', created: created, adopted: adopted, record: record,
      message: 'SHARING FAILED. The Shared Minutes document "' + record.name + '" ' + (created ? 'was created' : 'exists') + ' and is recorded in this report, but "anyone with the link can edit" could not be ' +
        (shareError ? 'set (' + shareError + ')' : 'confirmed') + '. It is NOT shared, and no link was written into the report. Press "Finish sharing" to try again.' + (placeNote ? ' ' + placeNote : '') };
  }
  record.permissionVerifiedAt = deps.nowIso();
  try {
    props.setProperty(SHARED_MINUTES_KEY_, serializeSharedMinutes_(record));
  } catch (e) {
    return { ok: false, state: 'sharing-failed', created: created, adopted: adopted, record: record,
      message: 'The Shared Minutes document "' + record.name + '" is shared, but that could not be recorded in this report (' + e.message + '). No link was written. Press the button again.' };
  }
  try { props.deleteProperty(SHARED_MINUTES_PENDING_KEY_); } catch (e) { Logger.log('Shared minutes: the creation marker could not be removed: ' + e.message); }

  const linkNote = writeSharedMinutesLinkWith_(deps, record);
  const what = created ? 'The Shared Minutes document "' + record.name + '" was created.'
    : (adopted ? 'The Shared Minutes document "' + record.name + '" existed already for this report and is used; no second one was created.'
      : 'The Shared Minutes document "' + record.name + '" is set up already; nothing was created.');
  return { ok: true, state: 'complete', created: created, adopted: adopted, record: record,
    message: what + ' Anyone with the link can edit it.' + (linkNote ? ' ' + linkNote : (created || adopted ? ' The link is in the report.' : '')) + (placeNote ? ' ' + placeNote : '') };
}

/**
 * Verify: asks Drive whether the recorded document is there and whether
 * anyone with the link can edit it, and records the answer. It never
 * creates a document and never changes who can open one. A document that is
 * verified has its link written into the report (it may not be there yet);
 * a link that is in the report is not taken out.
 * Returns { ok, message }. Never throws.
 */
function verifySharedMinutesWith_(deps) {
  const props = deps.documentProperties;
  const current = currentSharedMinutesWith_(deps);
  if (current.status !== 'ok') return { ok: false, message: 'There is no Shared Minutes document of this report to verify.' };
  const record = current.record;
  const unverified = function (message) {
    if (record.permissionVerifiedAt) {
      record.permissionVerifiedAt = null;
      try { props.setProperty(SHARED_MINUTES_KEY_, serializeSharedMinutes_(record)); } catch (e) { Logger.log('Shared minutes: the record could not be updated: ' + e.message); }
    }
    return { ok: false, message: message };
  };
  let meta;
  let permissions;
  try {
    meta = deps.drive.meta(record.fileId);
  } catch (e) {
    return { ok: false, message: 'Not verified: the Shared Minutes document "' + record.name + '" could not be opened (' + e.message + '). Nothing was changed.' };
  }
  if (!isSharedMinutesFile_(meta, current.documentId, current.meetingId)) {
    return unverified('NOT VERIFIED: the Shared Minutes document "' + record.name + '" is in the trash, or is not the document that was created for this report. Nothing was changed in Drive.');
  }
  try {
    permissions = deps.drive.permissions(record.fileId);
  } catch (e) {
    return { ok: false, message: 'Not verified: who can open "' + record.name + '" could not be read (' + e.message + '). Nothing was changed.' };
  }
  if (!hasAnyoneWriterPermission_(permissions)) {
    return unverified('NOT VERIFIED: "' + record.name + '" is not shared so that anyone with the link can edit. Nothing was changed in Drive; press "Finish sharing" to share it.');
  }
  record.permissionVerifiedAt = deps.nowIso();
  try {
    props.setProperty(SHARED_MINUTES_KEY_, serializeSharedMinutes_(record));
  } catch (e) {
    return { ok: false, message: 'The sharing is as it should be, but that could not be recorded (' + e.message + ').' };
  }
  const linkNote = writeSharedMinutesLinkWith_(deps, record);
  return { ok: true, message: 'Verified: "' + record.name + '" exists, and anyone with the link can edit it.' + (linkNote ? ' ' + linkNote : '') };
}

/**
 * Forget: the report no longer names a Shared Minutes document. Drive is
 * not asked anything and the document is not changed -- it exists and is
 * shared as before. The link line is taken out of the report only when it
 * is the link of the forgotten document.
 * Returns { ok, message }. Never throws.
 */
function forgetSharedMinutesWith_(deps) {
  const props = deps.documentProperties;
  const current = currentSharedMinutesWith_(deps);
  if (current.status === 'absent') return { ok: false, message: 'This report has no Shared Minutes document recorded.' };
  if (current.status === 'unsupported') return { ok: false, message: 'The shared minutes of this report were recorded by a newer release and are not changed here.' };
  try {
    props.deleteProperty(SHARED_MINUTES_KEY_);
    props.deleteProperty(SHARED_MINUTES_PENDING_KEY_);
  } catch (e) {
    return { ok: false, message: 'Nothing was changed: the record could not be removed (' + e.message + ').' };
  }
  let lineNote = '';
  if (current.record) {
    try {
      if (removeSharedMinutesLine_(deps.body(), current.record.url) === 'removed') lineNote = ' Its link was taken out of the report.';
    } catch (e) {
      Logger.log('Shared minutes: the link could not be taken out of the report: ' + e.message);
      lineNote = ' Its link could not be taken out of the report (' + e.message + '); remove the line by hand.';
    }
  }
  return { ok: true, message: 'This report no longer names a Shared Minutes document.' + lineNote +
    (current.record ? ' The document "' + current.record.name + '" itself was not changed: it exists and is shared as before.' : '') +
    (current.status === 'ok' ? ' Create Shared Minutes finds and uses it again, unless it is in the trash.' : '') };
}

// --- The link in the report: one line, "Link to the shared minutes: <name>".

/** The paragraph or list item a body child is, when it is ordinary text; null for a heading, a table and anything else. */
function sharedMinutesTextHolder_(child) {
  const type = child.getType();
  if (type === DocumentApp.ElementType.PARAGRAPH) {
    const p = child.asParagraph();
    return p.getHeading() === DocumentApp.ParagraphHeading.NORMAL ? p : null;
  }
  if (type === DocumentApp.ElementType.LIST_ITEM && typeof child.asListItem === 'function') return child.asListItem();
  return null;
}

/**
 * The line of the report that begins with the label: { index, holder }, or
 * null. In a main-meeting report it is the line the build copies from the
 * meeting report template; elsewhere it is the line written here.
 */
function findSharedMinutesLine_(body) {
  const label = SHARED_MINUTES_LABEL_.replace(/:$/, '').toLowerCase();
  const count = body.getNumChildren();
  for (let i = 0; i < count; i++) {
    const holder = sharedMinutesTextHolder_(body.getChild(i));
    if (holder && String(holder.getText()).trim().toLowerCase().indexOf(label) === 0) return { index: i, holder: holder };
  }
  return null;
}

/** The address a line links to (its first link), '' when it has none or that cannot be told. */
function sharedMinutesLineLink_(holder) {
  try {
    const text = holder.editAsText();
    if (typeof text.getLinkUrl !== 'function') return '';
    const length = Math.min(String(holder.getText()).length, 600);
    for (let i = 0; i < length; i++) {
      const url = text.getLinkUrl(i);
      if (url) return String(url);
    }
  } catch (e) {
    Logger.log('Shared minutes: the link of the line could not be read: ' + e.message);
  }
  return '';
}

/**
 * Where a new line goes: at the end of the Online information block of an
 * ad-hoc report with sessions (directly below its last line -- the block
 * itself is not touched, and is replaced without it); else directly below
 * the "Opening of the session" heading, else below the first "Opening ..."
 * heading. -1 when the report has none of them.
 */
function findSharedMinutesInsertIndex_(body) {
  const information = findAdhocMeetingInformationContainer_(body);
  if (information) return information.end;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const count = body.getNumChildren();
  let opening = -1;
  for (let i = 0; i < count; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH || child.asParagraph().getHeading() === NORMAL) continue;
    const text = child.asParagraph().getText().trim();
    if (/^(?:\d+(?:\.\d+)*\s+)?opening of the session\s*$/i.test(text)) return i + 1;
    if (opening === -1 && /^(?:\d+(?:\.\d+)*\s+)?opening\b/i.test(text)) opening = i;
  }
  return opening === -1 ? -1 : opening + 1;
}

/**
 * Writes the line for `link` ({ name, url }) and nothing else. A line that
 * says exactly this and links there is left alone; another line with the
 * label is rewritten where it stands; without one a new line is written.
 * Returns { line: 'created' | 'replaced' | 'unchanged' | 'no-place' }.
 */
function renderSharedMinutesLink_(body, link) {
  const text = SHARED_MINUTES_LABEL_ + ' ' + link.name;
  const existing = findSharedMinutesLine_(body);
  let holder;
  if (existing) {
    holder = existing.holder;
    if (holder.getText() === text && sharedMinutesLineLink_(holder) === link.url) return { line: 'unchanged' };
    holder.setText(text);
  } else {
    const index = findSharedMinutesInsertIndex_(body);
    if (index === -1) return { line: 'no-place' };
    holder = index >= body.getNumChildren() ? body.appendParagraph(text) : body.insertParagraph(index, text);
    holder.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  }
  const start = text.length - link.name.length;
  const t = holder.editAsText();
  if (!existing) t.setBold(false);
  // The label is plain text; the name is the link. (The line may have been a link as a whole before.)
  if (existing) { try { t.setLinkUrl(0, start - 1, null); } catch (e) { Logger.log('Shared minutes: an old link of the label could not be removed: ' + e.message); } }
  t.setLinkUrl(start, text.length - 1, link.url);
  return { line: existing ? 'replaced' : 'created' };
}

/** Takes the line out when it links to `url`. A line that links elsewhere, or nowhere, is somebody's text and stays. Returns 'removed' | 'kept' | 'none'. */
function removeSharedMinutesLine_(body, url) {
  const existing = findSharedMinutesLine_(body);
  if (!existing) return 'none';
  if (!url || sharedMinutesLineLink_(existing.holder) !== url) return 'kept';
  removeAdhocBodyChild_(body, body.getChild(existing.index));
  return 'removed';
}

/**
 * Build Report from Scratch, at the end (called by Code.js in the template
 * runtime): the link is written again from what is stored. Drive is not
 * asked anything, and without a verified record of this report nothing in
 * the document is looked at. Returns a note for the build result, '' when
 * all is well.
 */
function finishSharedMinutesRebuild_(body) {
  try {
    const current = currentSharedMinutesWith_({
      activeDocumentId: function () { return DocumentApp.getActiveDocument().getId(); },
      documentProperties: PropertiesService.getDocumentProperties()
    });
    if (current.status !== 'ok' || !current.record.permissionVerifiedAt) return '';
    const rendered = renderSharedMinutesLink_(body, current.record);
    return rendered.line === 'no-place' ? '\n\n⚠️ The link to the shared minutes could not be written: the report has no opening section.' : '';
  } catch (e) {
    Logger.log('Shared minutes: the link could not be written after the rebuild: ' + e.message);
    return '\n\n⚠️ The link to the shared minutes could not be written (' + e.message + '). Use Configure Meeting > Verify to write it.';
  }
}

// --- Configure Meeting: the Shared Minutes section.

/**
 * Pure: what the section shows -- { lines: [{ text, strong }], warning,
 * createLabel, createDisabled, openUrl, canVerify, canForget }. `stale`:
 * the report shows a link in its shared-minutes line that is not the link
 * of the recorded document (a copied report shows the link of the report it
 * was copied from).
 */
function sharedMinutesView_(current, stale) {
  const view = { lines: [], warning: '', createLabel: '', createDisabled: false, openUrl: '', canVerify: false, canForget: false };
  const line = function (text, strong) { view.lines.push({ text: text, strong: !!strong }); };
  const record = current.record;
  if (current.status === 'ok') {
    line(record.name, true);
    view.openUrl = record.url;
    view.canVerify = true;
    view.canForget = true;
    if (record.permissionVerifiedAt) {
      line('Anyone with the link can edit ✓ (verified ' + String(record.permissionVerifiedAt).slice(0, 10) + ')');
    } else {
      line('NOT shared: "anyone with the link can edit" is not confirmed. The link is not in the report.');
      view.createLabel = 'Finish sharing';
    }
  } else if (current.status === 'other-meeting') {
    line('Status: recorded for meeting ' + record.meetingId + ' ("' + record.name + '"); this report is configured for meeting ' + (current.meetingId || '(none)') + ' now. It is not used.');
    view.openUrl = record.url;
    view.canForget = true;
    view.createLabel = 'Create Shared Minutes';
    view.createDisabled = true;
  } else if (current.status === 'unsupported') {
    line('Status: recorded by a newer release; not shown and not changed here.');
  } else {
    line('Status: Not created');
    view.createLabel = 'Create Shared Minutes';
    if (!/^\d+$/.test(current.meetingId)) {
      view.createDisabled = true;
      line('Resolve the meeting and save the configuration first.');
    }
    if (current.status === 'invalid' || current.status === 'other-report') view.canForget = true;
  }
  if (stale) {
    view.warning = 'This report shows a link to shared minutes that is not the one recorded for it' + (current.status === 'ok' ? '' : ' (a copied report shows the link of the report it was copied from)') +
      '. It is replaced when the Shared Minutes of this report are created or verified; Build Report from Scratch does not write it again.';
  }
  return view;
}

/** The view of the active document, from what is stored and what the report shows. Drive is not asked. Never throws. */
function currentSharedMinutesView_() {
  let current = { documentId: '', meetingId: '', status: 'absent', record: null };
  let stale = false;
  try {
    const doc = DocumentApp.getActiveDocument();
    current = currentSharedMinutesWith_({ activeDocumentId: function () { return doc.getId(); }, documentProperties: PropertiesService.getDocumentProperties() });
    const shown = findSharedMinutesLine_(doc.getBody());
    const link = shown ? sharedMinutesLineLink_(shown.holder) : '';
    stale = !!link && link !== (current.status === 'ok' ? current.record.url : '');
  } catch (e) {
    Logger.log('Configure Meeting: the shared minutes could not be looked at: ' + e.message);
  }
  return sharedMinutesView_(current, stale);
}

/**
 * Called by configureMeetingSettings() (Code.js) in the template runtime:
 * the section and the script of its buttons. Words and controls only --
 * opening the dialog asks Drive nothing and changes nothing. The buttons
 * call createSharedMinutes(), verifySharedMinutes() and
 * forgetSharedMinutes(); Save Configuration does not.
 */
function sharedMinutesDialogParts_() {
  const attr = function (v) { return String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const sectionHtml =
    '<div class="section">\n' +
    '      <h3>Shared Minutes</h3>\n' +
    '      <div id="sharedMinutesBox"></div>\n' +
    '      <div class="hint">A separate Google Doc for the minutes takers. Create Shared Minutes makes it, lets anyone with the link edit it, and writes its link into the report. ' +
    'Nothing is created or shared by saving this dialog, by an update or by a build.</div>\n' +
    '      <input type="hidden" id="sharedMinutesView" value="' + attr(JSON.stringify(currentSharedMinutesView_())) + '">\n' +
    '    </div>';
  const script = [
    'var sharedMinutesViewNow = null;',
    'function sharedMinutesShow(view, message, busy) {',
    '  var box = document.getElementById("sharedMinutesBox");',
    '  if (!box) return;',
    '  if (view) sharedMinutesViewNow = view;',
    '  view = sharedMinutesViewNow;',
    '  if (!view) return;',
    '  var off = busy ? " disabled" : "";',
    '  var html = "";',
    '  view.lines.forEach(function (line) { html += "<div" + (line.strong ? " style=\'font-weight:bold\'" : "") + ">" + escapeHtml(line.text) + "</div>"; });',
    '  if (view.warning) html += "<div class=\'hint\' style=\'color:#a94442\'>" + escapeHtml(view.warning) + "</div>";',
    '  if (message) html += "<div id=\'sharedMinutesMessage\' style=\'white-space:pre-wrap;font-size:12px;margin-top:6px\'>" + escapeHtml(message) + "</div>";',
    '  if (view.createLabel) html += "<button type=\'button\' onclick=\'sharedMinutesCreate()\'" + (view.createDisabled ? " disabled" : off) + ">" + escapeHtml(view.createLabel) + "</button> ";',
    '  if (view.openUrl) html += "<a href=\'" + escapeHtml(view.openUrl) + "\' target=\'_blank\'>Open Shared Minutes</a> ";',
    '  if (view.canVerify) html += "<button type=\'button\' style=\'background:#666\' onclick=\'sharedMinutesVerify()\'" + off + ">Verify</button> ";',
    '  if (view.canForget) html += "<button type=\'button\' style=\'background:#666\' onclick=\'sharedMinutesForget()\'" + off + ">Forget</button>";',
    '  box.innerHTML = html;',
    '}',
    'function sharedMinutesRun(name, busyText) {',
    '  sharedMinutesShow(null, busyText, true);',
    '  google.script.run',
    '    .withSuccessHandler(function (result) { sharedMinutesShow(result && result.view, result && result.message, false); })',
    '    .withFailureHandler(function (error) { sharedMinutesShow(null, "Failed: " + error, false); })',
    '    [name]();',
    '}',
    'function sharedMinutesCreate() { sharedMinutesRun("createSharedMinutes", "Working. This can take a moment; do not close the dialog."); }',
    'function sharedMinutesVerify() { sharedMinutesRun("verifySharedMinutes", "Asking Google Drive."); }',
    'function sharedMinutesForget() {',
    '  if (typeof confirm === "function" && !confirm("Forget the Shared Minutes document of this report? The document itself is not changed.")) return;',
    '  sharedMinutesRun("forgetSharedMinutes", "Working.");',
    '}',
    // The first thing the dialog script does; escapeHtml() is a function of that script, declared further down.
    'try { sharedMinutesShow(JSON.parse(document.getElementById("sharedMinutesView").value), "", false); } catch (e) {}'
  ].join('\n      ');
  return { sectionHtml: sectionHtml, script: script };
}

// --- The live entry points (the buttons of the dialog).

function liveSharedMinutesDeps_() {
  const doc = DocumentApp.getActiveDocument();
  const props = PropertiesService.getDocumentProperties();
  return {
    release: templateRuntimeRelease_(),
    activeDocumentId: function () { return doc.getId(); },
    nowIso: function () { return new Date().toISOString(); },
    nowMs: function () { return Date.now(); },
    documentProperties: props,
    drive: sharedMinutesDrive_(),
    // The name follows the title a build gives the report (setDocumentTitleFromTemplate_(), Code.js) -- from the configuration, not from the file name.
    describe: function () {
      const context = getMeetingContext_();
      const title = generateReportTitle_({
        REPORT_SUFFIX: context.report.type,
        TDOC_LIST_URL: context.sources.tdocListUrl,
        MEETING_ID: context.meeting.portalId,
        meetingLabel: context.meeting.type === 'adhoc' ? context.meeting.name : undefined
      });
      return { name: sharedMinutesName_(title), meetingName: props.getProperty('MEETING_NAME'), meetingDate: props.getProperty('MEETING_DATE') };
    },
    body: function () { return doc.getBody(); }
  };
}

/**
 * One action of the section, under the document lock an update holds: two
 * presses, or a press during an update, run one after the other. Returns
 * what the action returned, with the view the section shows afterwards.
 * Never throws.
 */
function runSharedMinutesAction_(action) {
  let result;
  try {
    assertNotTemplateMaster_();
    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(30000)) {
      result = { ok: false, message: 'Nothing was done: another action on this report is running right now. Try again in a minute.' };
    } else {
      try {
        result = action(liveSharedMinutesDeps_());
      } finally {
        lock.releaseLock();
      }
    }
  } catch (e) {
    Logger.log('Shared minutes: the action failed: ' + e.message);
    result = { ok: false, message: 'The action stopped (' + e.message + '). Press the button again: it goes on from what is there.' };
  }
  result.view = currentSharedMinutesView_();
  return result;
}

/** RPC of Configure Meeting: Create Shared Minutes / Finish sharing. The only entry point that creates or shares. */
function createSharedMinutes() {
  return runSharedMinutesAction_(createSharedMinutesWith_);
}

/** RPC of Configure Meeting: Verify. */
function verifySharedMinutes() {
  return runSharedMinutesAction_(verifySharedMinutesWith_);
}

/** RPC of Configure Meeting: Forget. */
function forgetSharedMinutes() {
  return runSharedMinutesAction_(forgetSharedMinutesWith_);
}
