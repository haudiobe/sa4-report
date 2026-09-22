/**
 * POST-MEETING-001 (Task A1/A2) / RESOLVER-HOTFIX — proves the EXACT
 * network footprint of a normal UI Resolve click, end to end from the
 * dialog's own client-side call target through to the server function it
 * actually invokes.
 *
 * RESOLVER-HOTFIX: live evidence (Apps Script Executions, after clicking
 * Resolve for meeting 86178) showed ZERO executions of
 * resolveMeetingForConfigDialog_ -- only the unrelated
 * configureMeetingSettings() call that renders the dialog itself. Root
 * cause: Apps Script's google.script.run can only invoke a PUBLIC
 * top-level function; by Apps Script convention/platform behavior, a
 * function name ending in "_" is private -- hidden from the IDE's
 * function selector, not assignable as a trigger handler, and NOT
 * invokable via google.script.run from client HTML. This is exactly the
 * convention this whole codebase already uses for internal-only helpers
 * (hence the pervasive "_" suffix throughout Code.js) -- and, cross-
 * checked directly against every OTHER google.script.run target in this
 * file (createContinuousTrigger, setFetchAbstractsSetting,
 * deleteContinuousTrigger, saveConfigurationSettings, saveReallocation,
 * parseAgendaDocumentById, parseAgendaFromZippedTdoc), every one of them
 * is a public, non-underscore name -- resolveMeetingForConfigDialog_ and
 * discoverAgendaForConfigDialog_ were the only two RPC targets that broke
 * this pattern, and the only two known to be broken live.
 *
 * Fix: two thin public wrappers, resolveMeetingForConfigDialog() and
 * discoverAgendaForConfigDialog() (no trailing underscore), each
 * delegating to its identically-named "_" internal implementation. The
 * dialog's client script now calls the public wrappers; the internal
 * functions, their names, and every test against them directly are
 * otherwise unchanged.
 *
 * Task A1's call graph (verified directly against Code.js while writing
 * this suite, cross-checked with tests/resolve-dialog-client-rendering.test.js's
 * extraction of the real <script> block):
 *
 *   [Resolve button, onclick="resolveMeeting()"]
 *     -> client resolveMeeting() (configureMeetingSettings()'s <script>)
 *          reads #meetingId.value, sets #resolveBtn.disabled=true,
 *          #resolveStatus.textContent = "Resolving from 3GPP…"
 *     -> google.script.run.withSuccessHandler(...).withFailureHandler(...)
 *          .resolveMeetingForConfigDialog(meetingId)      <-- THE public server RPC target
 *     -> SERVER: resolveMeetingForConfigDialog(meetingIdInput)
 *          -> resolveMeetingForConfigDialog_(meetingIdInput)  <-- internal, unchanged
 *          - PropertiesService.getDocumentProperties() reads (9 properties,
 *            no network)
 *          - parseMeetingIdInput_(meetingIdInput)                (pure)
 *          - resolveMeetingCoreById_(idResult.id)                <-- NOT
 *            the full resolveMeetingById_()
 *              - parseMeetingIdInput_ again (pure)
 *              - fetchMeetingMetadataById_(id)                   NETWORK #1 (POST GetMeetings)
 *              - parseMeetingMetadataResponse_ (pure)
 *              - shouldFetchIcalFallback_(metadataParsed)         (pure)
 *              - IF (and only if) shouldFetchIcalFallback_ says so:
 *                  fetchMeetingIcalById_(id)                     NETWORK #2 (conditional, GetiCal)
 *                  parseMeetingIcal_ (pure)
 *              - normalizePortalMeetingType_ (pure)
 *              - normalizeMtgDocUrlToFtpBase_ (pure)
 *              - deriveRevisionsUrlCandidate_ (pure, NO network --
 *                derives a candidate string only)
 *              - normalizePortalMeetingTitle_ (pure)
 *          - computeResolvedMeetingPreview_(existing, resolved)     (pure)
 *     -> client success/failure handler (see
 *        tests/resolve-dialog-client-rendering.test.js for that half)
 *
 * Network calls reachable from this path: fetchMeetingMetadataById_
 * (always), fetchMeetingIcalById_ (conditionally). NEVER reachable:
 * fetchMeetingTdocListById_, fetchRevisionsUrlCandidate_,
 * validateRevisionsUrlCandidate_ (the last is never called from ANY
 * production code path at all -- it exists only for a future explicit
 * "Validate" action). This suite proves all of that against the ACTUAL
 * public RPC entry points the dialog's google.script.run calls name --
 * not merely resolveMeetingCoreById_() in isolation.
 *
 * Run: node tests/resolve-dialog-network-isolation.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

const FIXTURES = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-resolver-samples.json'), 'utf8'));

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

function throwIfCalled(label) {
  return () => { throw new Error(label + ' must not be called from a normal Resolve (POST-MEETING-001 Task A2)'); };
}

// ===== 1. confirm the dialog's own client script names the PUBLIC wrapper =

console.log('source-structure: the dialog\'s Resolve/Discover buttons target the PUBLIC (no trailing "_") wrapper functions');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function configureMeetingSettings()');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  check('the Resolve button\'s onclick handler is resolveMeeting()',
    /id="resolveBtn"[^>]*onclick="resolveMeeting\(\)"/.test(body), true);
  check('client resolveMeeting() calls google.script.run...resolveMeetingForConfigDialog(meetingId) -- the PUBLIC wrapper',
    /\.resolveMeetingForConfigDialog\(meetingId\)/.test(body), true);
  check('client resolveMeeting() does NOT call the private "_" name directly any more',
    /\.resolveMeetingForConfigDialog_\(/.test(body), false);
  check('client discoverAgendaTdocs() calls google.script.run...discoverAgendaForConfigDialog(...) -- the PUBLIC wrapper',
    /\.discoverAgendaForConfigDialog\(meetingId,\s*lastResolvedCore\)/.test(body), true);
  check('client discoverAgendaTdocs() does NOT call the private "_" name directly any more',
    /\.discoverAgendaForConfigDialog_\(/.test(body), false);
  check('client resolveMeeting() does NOT call the full resolveMeetingById_ or any TdocList/revisions function by name',
    /\.(resolveMeetingById_|fetchMeetingTdocListById_|fetchRevisionsUrlCandidate_|validateRevisionsUrlCandidate_)\(/.test(body), false);
}

// ===== 1b. no google.script.run target anywhere in the dialog ends in "_" ==

console.log('RESOLVER-HOTFIX: every google.script.run target used by the Meeting Configuration dialog is a public (non-"_" ) top-level function');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function configureMeetingSettings()');
  const saveFnStart = source.indexOf('function saveConfigurationSettings(');
  const nextFnMatch = source.slice(saveFnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? saveFnStart + 1 + nextFnMatch.index : source.length;
  // configureMeetingSettings() through saveConfigurationSettings() covers
  // the WHOLE Meeting Configuration dialog: its own HTML/script (Resolve,
  // Discover Agenda/TDocs, Save) plus the server function Save calls.
  const dialogRegion = source.slice(fnStart, fnEnd);

  // Every terminal google.script.run call in this dialog: the line
  // immediately after a `.withFailureHandler(...)` close is always the
  // actual RPC dispatch in this codebase's established style (see every
  // google.script.run chain in Code.js) -- excludes withSuccessHandler/
  // withFailureHandler themselves and unrelated chained calls like
  // .setWidth()/.setHeight() on HtmlOutput, which never follow
  // withFailureHandler.
  const rpcCallPattern = /\.withFailureHandler\([\s\S]*?\)\s*\n\s*\.([A-Za-z_][A-Za-z0-9_]*)\(/g;
  const rpcTargets = [];
  let m;
  while ((m = rpcCallPattern.exec(dialogRegion)) !== null) rpcTargets.push(m[1]);

  check('exactly the three known Meeting Configuration RPC targets are found, each exactly once',
    rpcTargets.sort(), ['discoverAgendaForConfigDialog', 'resolveMeetingForConfigDialog', 'saveConfigurationSettings'].sort());
  check('none of them end with "_" (all are public, callable via google.script.run)',
    rpcTargets.every(n => !n.endsWith('_')), true);
}

// ===== 2. resolveMeetingForConfigDialog() -- the actual PUBLIC RPC target =

console.log('resolveMeetingForConfigDialog() (public wrapper) -- exact network isolation, the REAL function the dialog invokes');

{
  const { sandbox } = loadCode();
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });
  sandbox.fetchMeetingIcalById_ = throwIfCalled('fetchMeetingIcalById_'); // 86178 supplies all 4 fallback fields -- never needed
  sandbox.fetchMeetingTdocListById_ = throwIfCalled('fetchMeetingTdocListById_');
  sandbox.fetchRevisionsUrlCandidate_ = throwIfCalled('fetchRevisionsUrlCandidate_');
  sandbox.validateRevisionsUrlCandidate_ = throwIfCalled('validateRevisionsUrlCandidate_');

  let threw = null;
  let result;
  try {
    result = sandbox.resolveMeetingForConfigDialog('86178');
  } catch (e) {
    threw = e;
  }

  check('resolveMeetingForConfigDialog(\'86178\') (the public wrapper, exactly as google.script.run calls it) does not throw', threw, null);
  check('result.ok is true', result && result.ok, true);
  check('result.resolved.meeting.name resolved correctly (proves the real path ran, not a no-op)',
    result && result.resolved && result.resolved.meeting && result.resolved.meeting.name, 'SA4-e (AH) on FS_6G_MED');
  check('result.resolved.documents is null (core-only -- confirms TdocList truly was not reached)',
    result && result.resolved && result.resolved.documents, null);
}

// ===== 2b. the public wrapper delegates to, and only to, the "_" internal =

console.log('resolveMeetingForConfigDialog() delegates exactly to resolveMeetingForConfigDialog_()');

{
  const { sandbox } = loadCode();
  let internalCalledWith = null;
  const originalInternal = sandbox.resolveMeetingForConfigDialog_;
  sandbox.resolveMeetingForConfigDialog_ = function (arg) {
    internalCalledWith = arg;
    return originalInternal(arg);
  };
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) });

  const publicResult = sandbox.resolveMeetingForConfigDialog('86178');
  check('the public wrapper called the internal "_" function with the same argument', internalCalledWith, '86178');
  check('the public wrapper returns exactly what the internal function returned (pure delegation, no extra logic)',
    publicResult, originalInternal('86178'));
}

{
  const { sandbox } = loadCode();
  let internalCalledWith = null;
  const originalInternal = sandbox.discoverAgendaForConfigDialog_;
  sandbox.discoverAgendaForConfigDialog_ = function (id, core) {
    internalCalledWith = [id, core];
    return originalInternal(id, core);
  };

  sandbox.discoverAgendaForConfigDialog('86178', null);
  check('discoverAgendaForConfigDialog() delegates to discoverAgendaForConfigDialog_() with the same arguments',
    internalCalledWith, ['86178', null]);
}

// ===== 3. same proof for a meeting requiring the conditional iCal fallback

console.log('resolveMeetingForConfigDialog() -- iCal IS called only when genuinely needed, TdocList/revisions probes still never reached');

{
  const { sandbox } = loadCode();
  const degraded = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
  delete degraded[0].Location; // forces shouldFetchIcalFallback_() to true
  let icalCalled = false;
  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(degraded) });
  sandbox.fetchMeetingIcalById_ = () => { icalCalled = true; return { statusCode: 200, text: FIXTURES.ical86178 }; };
  sandbox.fetchMeetingTdocListById_ = throwIfCalled('fetchMeetingTdocListById_');
  sandbox.fetchRevisionsUrlCandidate_ = throwIfCalled('fetchRevisionsUrlCandidate_');
  sandbox.validateRevisionsUrlCandidate_ = throwIfCalled('validateRevisionsUrlCandidate_');

  const result = sandbox.resolveMeetingForConfigDialog('86178');
  check('degraded metadata: the conditional GetiCal fallback IS reached (genuine fallback case)', icalCalled, true);
  check('degraded metadata: still ok:true, still never reaches TdocList/revisions', result.ok, true);
}

// ===== 4. validateRevisionsUrlCandidate_ has NO callers anywhere in Code.js

console.log('source-structure: validateRevisionsUrlCandidate_() is called from NO production code path at all');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const definitionMatch = source.match(/^function validateRevisionsUrlCandidate_\(/m);
  check('validateRevisionsUrlCandidate_() is still defined (reserved for a future explicit action)', !!definitionMatch, true);

  // Every occurrence of the identifier that is NOT its own `function` line
  // must be inside a comment (an actual call site would be
  // "validateRevisionsUrlCandidate_(" immediately following non-comment
  // code) -- there must be zero real call sites.
  const callSitePattern = /(?<!function )validateRevisionsUrlCandidate_\(/g;
  const lines = source.split('\n');
  const nonCommentCallSites = lines.filter(line => {
    if (!callSitePattern.test(line)) return false;
    callSitePattern.lastIndex = 0;
    const trimmed = line.trim();
    return !trimmed.startsWith('//') && !trimmed.startsWith('*') && !trimmed.startsWith('/*');
  });
  check('no non-comment call site of validateRevisionsUrlCandidate_() exists outside its own definition',
    nonCommentCallSites.length, 0);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
