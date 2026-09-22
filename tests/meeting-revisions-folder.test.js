/**
 * ARCH-011 — Meeting-ID resolver: revisions/drafts folder discovery.
 *
 * Extends the committed ARCH-009 resolver core (resolveMeetingById_() and
 * friends) with a fourth, DERIVED endpoint: a candidate revisions/drafts
 * folder one level up from the resolved FTP Docs/ directory
 * (<meeting root>/inbox/drafts/), which is only ever reported as
 * `sources.revisionsUrl` when it has been independently validated (HTTP
 * 200 + credible directory/file listing evidence), never merely because it
 * was derived.
 *
 * NOT integrated into getMeetingContext_(), the configuration dialog, or
 * Document Properties in this task -- resolveMeetingById_() remains dead
 * code from the rest of the app's point of view (see ARCH-009's own header
 * comment in Code.js), this only extends what it discovers.
 *
 * A note on fixtures: unlike tests/meeting-resolver.test.js's GetMeetings/
 * GetiCal/TdocList.aspx fixtures (which are byte-for-byte real captured
 * responses), the "directory listing" HTML bodies used below are
 * SYNTHETIC -- written to match the real, observed shape of 3GPP's FTP
 * directory pages (an Apache-style autoindex with an "Index of" title and
 * `<a href="...">` entries, confirmed by a live, browser-rendered fetch of
 * https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/
 * during this task, which showed real files such as
 * "3GPP_TR_26870-0.5.1-6G-MED rm.docx" and "S4aP260096_stabilization_huawei.docx")
 * but not literal captured bytes, since a direct (non-browser) HTTP probe
 * of that same URL from this environment returned HTTP 403 from 3GPP's
 * WAF/bot-challenge -- the same 403 a plain UrlFetchApp-style client may
 * see in practice (see this task's report for the full finding). These
 * tests therefore exercise the VALIDATION LOGIC against realistic inputs,
 * not a claim about what any specific live fetch will return today.
 *
 * Run: node tests/meeting-revisions-folder.test.js
 */

const fs = require('fs');
const path = require('path');
const { loadCode } = require('./helpers/load-code.js');

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

// Synthetic (see file header), but shaped exactly like the real, confirmed
// Apache-style autoindex 3GPP's FTP server serves for a valid directory.
const SYNTHETIC_INDEX_HTML_86178 = `<!DOCTYPE html><html><head><title>Index of /ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/</title></head>
<body><h1>Index of /ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/</h1>
<table><tr><th>Name</th><th>Last modified</th><th>Size</th></tr>
<tr><td><a href="../">../</a></td><td></td><td>-</td></tr>
<tr><td><a href="S4aP260096_stabilization_huawei.docx">S4aP260096_stabilization_huawei.docx</a></td><td>2026-09-20 10:12</td><td>142K</td></tr>
<tr><td><a href="3GPP_TR_26870-0.5.1-6G-MED%20rm.docx">3GPP_TR_26870-0.5.1-6G-MED rm.docx</a></td><td>2026-09-21 08:03</td><td>1.1M</td></tr>
</table></body></html>`;

// A "folder of subfolders" listing -- the shape observed for the main
// meeting (60778) candidate, which lands on the SHARED per-series
// Inbox/Drafts/ parent (containing one subfolder per SWG, matching this
// project's existing DRAFTS_FOLDERS convention in getReportConfig_())
// rather than a single meeting-specific drafts folder directly. Still a
// credible, valid directory listing by this task's validation rule.
const SYNTHETIC_INDEX_HTML_60778 = `<!DOCTYPE html><html><head><title>Index of /ftp/tsg_sa/WG4_CODEC/TSGS4_137-e/inbox/drafts/</title></head>
<body><h1>Index of /ftp/tsg_sa/WG4_CODEC/TSGS4_137-e/inbox/drafts/</h1>
<table><tr><th>Name</th><th>Last modified</th></tr>
<tr><td><a href="Audio/">Audio/</a></td><td>2026-08-30 09:00</td></tr>
<tr><td><a href="FS_6G_MED/">FS_6G_MED/</a></td><td>2026-08-30 09:00</td></tr>
<tr><td><a href="MBS/">MBS/</a></td><td>2026-08-30 09:00</td></tr>
<tr><td><a href="Plenary/">Plenary/</a></td><td>2026-08-30 09:00</td></tr>
<tr><td><a href="RTC/">RTC/</a></td><td>2026-08-30 09:00</td></tr>
<tr><td><a href="Video/">Video/</a></td><td>2026-08-30 09:00</td></tr>
</table></body></html>`;

// A representative WAF/bot-challenge block page -- HTTP 403, HTML body,
// but NOT a directory listing. This is the real, observed shape a plain
// (non-browser) HTTP client currently gets from www.3gpp.org's FTP paths
// (see file header) -- distinct from a clean 404.
const WAF_CHALLENGE_BODY = '<!DOCTYPE html><html><head><title>Access Denied</title></head><body><div class="message-container">Please verify you are human to continue.</div></body></html>';

// ============================== 1. deriveRevisionsUrlCandidate_() ==========

console.log('deriveRevisionsUrlCandidate_() -- pure derivation from an already-normalized ftpBase');

{
  const { sandbox } = loadCode();
  const fn = sandbox.deriveRevisionsUrlCandidate_;

  check('/SA4_Plenary/Docs/ -> /SA4_Plenary/inbox/drafts/ (the known real 86178 relationship)',
    fn('https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/').revisionsUrlCandidate,
    'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/');

  check('86174/85916 real ftpBase (TSG_SA casing preserved) -> .../SA4_Audio/inbox/drafts/',
    fn(FIXTURES.getMeetings86174[0].MtgDocURL).revisionsUrlCandidate,
    'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/');

  check('60778 real ftpBase (tsg_sa casing preserved) -> .../TSGS4_137-e/inbox/drafts/',
    fn(FIXTURES.getMeetings60778[0].MtgDocURL).revisionsUrlCandidate,
    'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_137-e/inbox/drafts/');

  check('already-normalized, single-trailing-slash input is handled unchanged',
    fn('https://ftp.3gpp.org/a/b/Docs/').revisionsUrlCandidate,
    'https://ftp.3gpp.org/a/b/inbox/drafts/');

  check('missing trailing slash on Docs is normalized before deriving',
    fn('https://ftp.3gpp.org/a/b/Docs').revisionsUrlCandidate,
    'https://ftp.3gpp.org/a/b/inbox/drafts/');

  check('a doubled internal slash is normalized before deriving (mirrors normalizeMtgDocUrlToFtpBase_\'s own 86178 case)',
    fn('https://ftp.3gpp.org/a//b/Docs/').revisionsUrlCandidate,
    'https://ftp.3gpp.org/a/b/inbox/drafts/');

  check('lowercase "docs/" is still recognized (case-insensitive match)',
    fn('https://ftp.3gpp.org/a/b/docs/').revisionsUrlCandidate,
    'https://ftp.3gpp.org/a/b/inbox/drafts/');

  const missingRoot = fn(null);
  check('missing ftpBase -> candidate null', missingRoot.revisionsUrlCandidate, null);
  check('missing ftpBase -> error explains why', typeof missingRoot.error, 'string');

  const blankRoot = fn('');
  check('blank ftpBase -> candidate null', blankRoot.revisionsUrlCandidate, null);

  const invalidRoot = fn('https://ftp.3gpp.org/some/other/structure/');
  check('ftpBase not ending in a recognizable Docs/ segment -> candidate null (never guessed)',
    invalidRoot.revisionsUrlCandidate, null);
  check('invalid meeting root -> error explains why', typeof invalidRoot.error, 'string');
}

// ============================ 2. validateRevisionsUrlCandidateResponse_() ==

console.log('validateRevisionsUrlCandidateResponse_() -- accept/reject a fetched candidate');

{
  const { sandbox } = loadCode();
  const fn = sandbox.validateRevisionsUrlCandidateResponse_;

  check('HTTP 200 + "Index of" autoindex title -> ok',
    fn({ statusCode: 200, text: SYNTHETIC_INDEX_HTML_86178 }).ok, true);

  check('HTTP 200 + at least one <a href=...> entry (no "Index of" title) -> ok',
    fn({ statusCode: 200, text: '<html><body><a href="file.docx">file.docx</a></body></html>' }).ok, true);

  const notFound = fn({ statusCode: 404, text: 'Not Found' });
  check('HTTP 404 -> not ok', notFound.ok, false);
  check('HTTP 404 -> reason mentions the status', /404/.test(notFound.reason), true);

  const forbidden = fn({ statusCode: 403, text: WAF_CHALLENGE_BODY });
  check('HTTP 403 (WAF/bot-challenge page) -> not ok, even though it has a body', forbidden.ok, false);

  const redirect = fn({ statusCode: 301, text: '' });
  check('HTTP 301 (redirect status, since fetchRevisionsUrlCandidate_ uses followRedirects:false) -> not ok', redirect.ok, false);

  const emptyBody = fn({ statusCode: 200, text: '' });
  check('HTTP 200 but empty body -> not ok (not credible evidence)', emptyBody.ok, false);

  const noListing = fn({ statusCode: 200, text: '<html><body>Welcome to 3GPP</body></html>' });
  check('HTTP 200 but body has no "Index of"/<a href> evidence -> not ok', noListing.ok, false);

  check('missing fetchResult does not throw', fn(null).ok, false);
  check('missing fetchResult -> reason present', typeof fn(null).reason, 'string');
}

// ==================== 3. resolveMeetingById_() end-to-end (86178) ==========

console.log('resolveMeetingById_() -- revisions folder discovery wired end-to-end, meeting 86178');

function loadResolverSandbox(overrides) {
  const { sandbox } = loadCode();
  const o = overrides || {};
  sandbox.fetchMeetingMetadataById_ = () => o.metadata !== undefined ? o.metadata : { statusCode: 200, text: '[]' };
  sandbox.fetchMeetingIcalById_ = () => o.ical !== undefined ? o.ical : { statusCode: 200, text: '' };
  sandbox.fetchMeetingTdocListById_ = () => o.tdoc !== undefined ? o.tdoc : { statusCode: 200, text: '' };
  sandbox.fetchRevisionsUrlCandidate_ = (url) => {
    if (o.revisionsThrow) throw new Error(o.revisionsThrow);
    return o.revisions !== undefined ? o.revisions(url) : { statusCode: 404, text: 'Not Found' };
  };
  return sandbox;
}

{
  // (a) successful validation
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    ical: { statusCode: 200, text: FIXTURES.ical86178 },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86178 },
    revisions: () => ({ statusCode: 200, text: SYNTHETIC_INDEX_HTML_86178 })
  });
  const result = sandbox.resolveMeetingById_(86178);
  check('86178: sources.revisionsUrl resolved to the derived candidate',
    result.sources.revisionsUrl, 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/');
  check('86178: unresolved does NOT contain sources.revisionsUrl when validated',
    result.unresolved.indexOf('sources.revisionsUrl') !== -1, false);
  check('86178: meeting/ftpBase/agendaTdoc still resolve normally alongside it',
    [result.meeting.name, result.sources.ftpBase, result.documents.agendaTdoc].every(Boolean), true);
}

{
  // (b) candidate derived, but inaccessible (404) -- must NOT be fatal.
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    revisions: () => ({ statusCode: 404, text: 'Not Found' })
  });
  const result = sandbox.resolveMeetingById_(86178);
  check('86178, inaccessible candidate: sources.revisionsUrl is null, not the unvalidated candidate',
    result.sources.revisionsUrl, null);
  check('86178, inaccessible candidate: unresolved contains sources.revisionsUrl',
    result.unresolved.indexOf('sources.revisionsUrl') !== -1, true);
  check('86178, inaccessible candidate: a warning names the candidate URL and the reason',
    result.warnings.some(w => /inbox\/drafts/.test(w) && /404/.test(w)), true);
  check('86178, inaccessible candidate: does not throw, meeting metadata still resolves',
    result.meeting.name, 'SA4-e (AH) on FS_6G_MED');
}

{
  // (c) WAF/bot-challenge 403 -- the realistic failure mode found during
  // this task's live investigation (see file header).
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    revisions: () => ({ statusCode: 403, text: WAF_CHALLENGE_BODY })
  });
  const result = sandbox.resolveMeetingById_(86178);
  check('86178, WAF-blocked candidate: sources.revisionsUrl is null, not silently accepted',
    result.sources.revisionsUrl, null);
  check('86178, WAF-blocked candidate: unresolved contains sources.revisionsUrl',
    result.unresolved.indexOf('sources.revisionsUrl') !== -1, true);
}

{
  // (d) the revisions-folder probe itself throws (network failure) --
  // resolveMeetingById_() overall must still not throw or fail.
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86178) },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86178 },
    revisionsThrow: 'simulated network failure'
  });
  const result = sandbox.resolveMeetingById_(86178);
  check('86178, revisions probe throws: resolveMeetingById_ itself does not throw', typeof result, 'object');
  check('86178, revisions probe throws: sources.revisionsUrl is null', result.sources.revisionsUrl, null);
  check('86178, revisions probe throws: a warning records the failure',
    result.warnings.some(w => /simulated network failure/.test(w)), true);
  check('86178, revisions probe throws: meeting/ftpBase/documents still resolved normally',
    [result.meeting.name, result.sources.ftpBase, result.documents.agendaTdoc].every(Boolean), true);
}

{
  // (e) no ftpBase at all (MtgDocURL missing) -- derivation is skipped
  // entirely, not attempted against a garbage URL.
  const meetingNoFtp = JSON.parse(JSON.stringify(FIXTURES.getMeetings86178));
  delete meetingNoFtp[0].MtgDocURL;
  let revisionsFetchCalled = false;
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(meetingNoFtp) },
    revisions: () => { revisionsFetchCalled = true; return { statusCode: 200, text: SYNTHETIC_INDEX_HTML_86178 }; }
  });
  const result = sandbox.resolveMeetingById_(86178);
  check('missing MtgDocURL: sources.revisionsUrl is null', result.sources.revisionsUrl, null);
  check('missing MtgDocURL: unresolved contains sources.revisionsUrl', result.unresolved.indexOf('sources.revisionsUrl') !== -1, true);
  check('missing MtgDocURL: the revisions-folder probe is never even attempted', revisionsFetchCalled, false);
}

// ============ 4. resolveMeetingById_() -- 86174, 85916, 60778 (candidates)

console.log('resolveMeetingById_() -- revisions folder candidate derivation for 86174/85916/60778');

{
  // 86174 and 85916 share the same real MtgDocURL (SA4_Audio/Docs/, the
  // shared, unshared-per-series ad hoc folder -- see Code.js's own
  // pre-existing comment on this). Both derive to, and here validate
  // against, the SAME candidate -- this is a real, observed relationship,
  // not a meeting-type guess.
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings86174) },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml86174 },
    revisions: (url) => {
      check('86174: the exact candidate probed is .../SA4_Audio/inbox/drafts/',
        url, 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/');
      return { statusCode: 200, text: '<html><head><title>Index of /TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/</title></head><body><a href="report.docx">report.docx</a></body></html>' };
    }
  });
  const result = sandbox.resolveMeetingById_(86174);
  check('86174: sources.revisionsUrl validated', result.sources.revisionsUrl,
    'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/');
  check('86174: agendaTdoc still genuinely unresolved (real negative case, unaffected by this change)',
    result.documents.agendaTdoc, null);
}

{
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings85916) },
    revisions: () => ({ statusCode: 200, text: '<html><head><title>Index of /TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/</title></head><body><a href="report.docx">report.docx</a></body></html>' })
  });
  const result = sandbox.resolveMeetingById_(85916);
  check('85916: sources.revisionsUrl validated (same shared ad-hoc-series folder as 86174)',
    result.sources.revisionsUrl, 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/');
}

{
  // 60778 (main/plenary): the candidate is real and validates, but lands
  // on the SHARED per-series Inbox/Drafts/ parent (one subfolder per SWG),
  // not a single meeting-specific folder directly -- recorded as observed
  // evidence, not treated as equivalent to the ad-hoc case. Deciding how a
  // future consumer should use this (e.g. still requiring the existing
  // REPORT_SUFFIX-based DRAFTS_FOLDERS subfolder pick) is explicitly out
  // of scope for this task.
  const sandbox = loadResolverSandbox({
    metadata: { statusCode: 200, text: JSON.stringify(FIXTURES.getMeetings60778) },
    tdoc: { statusCode: 200, text: FIXTURES.tdocListHtml60778RevisionPair },
    revisions: () => ({ statusCode: 200, text: SYNTHETIC_INDEX_HTML_60778 })
  });
  const result = sandbox.resolveMeetingById_(60778);
  check('60778: sources.revisionsUrl validated', result.sources.revisionsUrl,
    'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_137-e/inbox/drafts/');
  check('60778: documents.agendaTdoc still resolves as before, unaffected by this change',
    result.documents.agendaTdoc, 'S4-261392');
}

// ===================== 5. no PropertiesService, no new production wiring ==

console.log('source-structure: revisions-folder discovery stays additive, dead code from production\'s point of view');

{
  const { CODE_JS_PATH } = require('./helpers/load-code.js');
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  function extractFunctionBody(fnName) {
    const startMatch = source.match(new RegExp('^function ' + fnName + '\\(', 'm'));
    if (!startMatch) return null;
    const startIndex = startMatch.index;
    const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
    nextFnRe.lastIndex = startIndex + startMatch[0].length;
    const next = nextFnRe.exec(source);
    return source.slice(startIndex, next ? next.index : source.length);
  }

  const resolveBody = extractFunctionBody('resolveMeetingById_');
  check('resolveMeetingById_() exists', !!resolveBody, true);
  if (resolveBody) {
    check('resolveMeetingById_() never calls PropertiesService.setProperty/setProperties',
      /PropertiesService[\s\S]*?\.set(Property|Properties)/.test(resolveBody), false);
  }

  check('getMeetingContext_() source is unchanged by this task (still no reference to deriveRevisionsUrlCandidate_/fetchRevisionsUrlCandidate_)',
    /deriveRevisionsUrlCandidate_|fetchRevisionsUrlCandidate_/.test(extractFunctionBody('getMeetingContext_') || ''), false);

  check('configureMeetingSettings() is unchanged by this task (no reference to the new revisions-folder functions)',
    /deriveRevisionsUrlCandidate_|fetchRevisionsUrlCandidate_/.test(extractFunctionBody('configureMeetingSettings') || ''), false);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
