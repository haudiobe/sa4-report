/**
 * The drafts / revisions folder of a new ad-hoc report (after T-2026.10.8).
 *
 * REVISIONS_URL is where the revisions of the TDocs of a report are found
 * and where its discussion e-mails send them. A main-meeting report derives
 * it by formula. An ad-hoc series has its own folder,
 *
 *   <meeting root>/Docs/  ->  <meeting root>/inbox/drafts/
 *
 * which the meeting lookup derives from the Portal's MtgDocURL but does not
 * ask for (asking that server made the lookup hang -- PROD-014), and which
 * the creator therefore did not store. A new ad-hoc report had no drafts
 * folder until somebody typed it into Configure Meeting.
 *
 * Finish Report Setup now asks for the derived folder ONCE, after the report
 * is set up, and stores it only when the answer is a folder listing.
 *
 *   1. confirmed: stored as REVISIONS_URL, and used;
 *   2. not confirmed (403, a redirect, another status, an empty page, a
 *      page that is no listing, a request that fails or times out):
 *      nothing is stored as fact; the candidate is shown to the user;
 *   3. a folder that is configured is never replaced, and nothing is asked;
 *   4. a main-meeting report is not touched;
 *   5. a Video ad-hoc series (SA4_VIDEO) -- the same rule, nothing Video-specific;
 *   6. only an https address on a 3GPP host is asked;
 *   7. once, after the setup, and not from the meeting lookup.
 *
 * Portal answers are the captured ones of tests/fixtures; everything else is synthetic.
 *
 * Run: node tests/adhoc-revisions-folder-setup.test.js
 */

const fs = require('fs');
const path = require('path');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { createReport, firstRunFor, payloadOf, descriptionWith, plain, RELEASE, NEW_DOC_ID } = require('./helpers/template-creator.js');

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

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
const withoutComments = (src) => src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
const GM = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-86172-85916-getmeetings.json'), 'utf8'));
const ROWS_MBS = [{ id: 'S4aI260089', type: 'discussion', revisionOf: null, agendaItem: '2.7' }];
const MBS_DOCS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/';
const MBS_DRAFTS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/inbox/drafts/';
const CANDIDATE_KEY = 'SA4_REVISIONS_URL_CANDIDATE';
const LISTING = '<html><head><title>Index of /TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/inbox/drafts</title></head><body><a href="S4aI260089_r1.docx">S4aI260089_r1.docx</a></body></html>';

/**
 * The first run of a new report for 86172, with the REAL check of the
 * candidate (validateRevisionsUrlCandidate_()) over a fake network:
 * `answer` is { status, text } or 'throw'. Returns what happened.
 */
function setUp(answer, o) {
  const created = (o && o.created) || createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  const requests = [];
  const seen = [];
  const run = firstRunFor((o && o.description) || created.description, { props: o && o.props });
  run.s.UrlFetchApp = { fetch: (url, options) => {
    requests.push([url, plain(options || {})]);
    seen.push(run.docProps.getProperty('SA4_BOOTSTRAP_STATE'));
    if (answer === 'throw') throw new Error('Timeout: ' + url);
    return { getResponseCode: () => answer.status, getContentText: () => answer.text || '' };
  } };
  run.deps.validateRevisionsCandidate = (url) => run.s.validateRevisionsUrlCandidate_(url);
  const setup = run.run();
  return { run: run, setup: setup, requests: requests, stateWhenAsked: seen, props: run.docProps, created: created,
    folder: setup.revisionsFolder, revisionsUrl: run.docProps.getProperty('REVISIONS_URL'), candidate: run.docProps.getProperty(CANDIDATE_KEY) };
}

// ====================================================================
console.log('1. the derived folder is confirmed');
{
  const r = setUp({ status: 200, text: LISTING });
  check('the report is set up', [r.setup.status, r.props.getProperty('SA4_BOOTSTRAP_STATE'), r.props.getProperty('FTP_BASE')], ['configured', 'done', MBS_DOCS]);
  check('the creator still does not put an unconfirmed folder into the setup information', payloadOf(r.created.s, r.created.description).config.revisionsUrl, '');
  check('the folder derived from the document folder of the Portal is asked for, once, without following a redirect', r.requests, [[MBS_DRAFTS, { muteHttpExceptions: true, followRedirects: false }]]);
  check('it is stored as REVISIONS_URL', [r.revisionsUrl, r.folder], [MBS_DRAFTS, { status: 'stored', url: MBS_DRAFTS, reason: '' }]);
  check('nothing is kept as "not confirmed"', r.candidate, null);
  check('the report uses it: revision collection and the discussion e-mails read it from here', [r.run.s.getMeetingContext_().sources.revisionsUrl, r.run.s.resolveEmailExportRevisionUploadUrl_()], [MBS_DRAFTS, MBS_DRAFTS]);
  check('the user is told', r.run.s.describeRevisionsFolderOutcome_(r.folder), '✅ Drafts folder found and stored: ' + MBS_DRAFTS);
  check('About This Report has nothing to add', r.run.s.describeTemplateRuntime_({ release: RELEASE, documentId: NEW_DOC_ID, bootstrapState: 'done', meetingId: '86172', revisionsUrl: r.revisionsUrl, revisionsCandidate: r.candidate })
    .filter((line) => /Drafts folder|Revisions \/ Drafts URL/.test(line)), []);
}

// ====================================================================
console.log('2. the derived folder is not confirmed');
[
  ['403 (the server refuses the request)', { status: 403, text: 'Forbidden' }, 'HTTP 403'],
  ['a redirect (not followed: where it leads is not known)', { status: 302, text: '' }, 'HTTP 302'],
  ['a redirect to the same folder with another spelling', { status: 301, text: LISTING }, 'HTTP 301'],
  ['404 (there is no such folder)', { status: 404, text: 'Not Found' }, 'HTTP 404'],
  ['500', { status: 500, text: '' }, 'HTTP 500'],
  ['an empty page', { status: 200, text: '   ' }, 'Empty response body.'],
  ['a page that is no folder listing (a bot challenge)', { status: 200, text: '<html><body>Please enable JavaScript to continue.</body></html>' }, 'Response did not look like a directory/file listing (no "Index of" title, no <a href> entries).'],
  ['a request that fails or times out', 'throw', 'Revisions/drafts folder candidate request failed: Timeout: ' + MBS_DRAFTS]
].forEach(([name, answer, reason]) => {
  const r = setUp(answer);
  check(name + ': the report is set up all the same, and no REVISIONS_URL is stored',
    [r.setup.status, r.props.getProperty('SA4_BOOTSTRAP_STATE'), r.revisionsUrl, r.requests.length], ['configured', 'done', null, 1]);
  check(name + ': the candidate is kept apart, as not confirmed, with the reason', [r.candidate, r.folder], [MBS_DRAFTS, { status: 'unconfirmed', url: MBS_DRAFTS, reason: reason }]);
});
{
  const r = setUp({ status: 403, text: 'Forbidden' });
  check('the report does not use an unconfirmed folder: it has none', [r.run.s.getMeetingContext_().sources.revisionsUrl || '', r.run.s.resolveEmailExportRevisionUploadUrl_()], ['', '']);
  const told = r.run.s.describeRevisionsFolderOutcome_(r.folder);
  check('the user is told that it is still to set up, with the candidate to look at and where to enter it',
    [/^⚠️ Drafts folder still to set up\./.test(told), told.indexOf('(HTTP 403)') !== -1, told.indexOf('\n' + MBS_DRAFTS + '\n') !== -1, /Configure Meeting… \(Revisions \/ Drafts URL\)/.test(told),
      /no revisions are collected and no discussion e-mails can be prepared/.test(told)], [true, true, true, true, true]);
  const about = (facts) => r.run.s.describeTemplateRuntime_(Object.assign({ release: RELEASE, documentId: NEW_DOC_ID, bootstrapState: 'done', meetingId: '86172' }, facts)).filter((line) => /Drafts folder|Revisions \/ Drafts URL/.test(line));
  check('About This Report says so for as long as no folder is configured',
    [about({ revisionsUrl: r.revisionsUrl, revisionsCandidate: r.candidate }), about({ revisionsUrl: MBS_DRAFTS, revisionsCandidate: r.candidate }), about({})],
    [['Drafts folder: not set up. Derived, not confirmed: ' + MBS_DRAFTS, '  Check it and enter it in Configure Meeting (Revisions / Drafts URL).'], [], []]);
  // The user checks the folder and enters it in Configure Meeting: the ordinary save.
  r.run.s.persistConfigurationSettings_({ meetingId: '86172', meetingType: 'adhoc', reportType: 'MBS', mailingListMode: 'derived', apiTokenAction: 'keep', revisionsUrl: MBS_DRAFTS, tdocUrl: r.props.getProperty('TDOC_LIST_URL') });
  check('entered in Configure Meeting, it is the folder of the report', [r.props.getProperty('REVISIONS_URL'), r.run.s.getMeetingContext_().sources.revisionsUrl], [MBS_DRAFTS, MBS_DRAFTS]);
}
{
  // A runtime that cannot ask (no such dependency): nothing is stored, nothing is claimed.
  const created = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  const run = firstRunFor(created.description);
  const setup = run.run();
  check('without a way to ask, nothing is stored and nothing is recorded', [setup.status, setup.revisionsFolder, run.docProps.getProperty('REVISIONS_URL'), run.docProps.getProperty(CANDIDATE_KEY)],
    ['configured', { status: 'not-checked', url: MBS_DRAFTS, reason: '' }, null, null]);
  // A check that throws instead of answering.
  const thrower = firstRunFor(created.description, { deps: { validateRevisionsCandidate: () => { throw new Error('synthetic'); } } });
  const out = thrower.run();
  check('a check that throws is an unconfirmed folder, not a failed setup', [out.status, out.revisionsFolder, thrower.docProps.getProperty('REVISIONS_URL'), thrower.docProps.getProperty('SA4_BOOTSTRAP_STATE')],
    ['configured', { status: 'unconfirmed', url: MBS_DRAFTS, reason: 'The request failed: synthetic' }, null, 'done']);
  // An answer that is not a plain "ok: true".
  const vague = [{ ok: 'yes' }, { ok: 1 }, {}, null, undefined, 'ok'].map((answer) => { const x = firstRunFor(created.description, { deps: { validateRevisionsCandidate: () => answer } }); x.run(); return x.docProps.getProperty('REVISIONS_URL'); });
  check('only the answer "ok: true" stores it', vague, [null, null, null, null, null, null]);
}

// ====================================================================
console.log('3. a folder that is configured is never replaced');
{
  const OWN = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Inbox/Drafts/OwnChoice/';
  const r = setUp({ status: 200, text: LISTING }, { props: { REVISIONS_URL: OWN } });
  check('a REVISIONS_URL the document already has stays, and nothing is asked', [r.setup.status, r.revisionsUrl, r.requests, r.candidate, r.folder], ['configured', OWN, [], null, { status: 'kept', url: OWN, reason: '' }]);
  check('nothing is said about it', r.run.s.describeRevisionsFolderOutcome_(r.folder), '');
  // Setup information that carries a folder (one that was confirmed): it is stored by the ordinary save, and nothing is asked.
  const created = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  const payload = payloadOf(created.s, created.description);
  payload.config.revisionsUrl = OWN;
  const carried = setUp({ status: 200, text: LISTING }, { created: created, description: descriptionWith(created.s, payload) });
  check('a folder that comes with the setup information is stored as it is, and nothing is asked', [carried.revisionsUrl, carried.requests, carried.folder.status], [OWN, [], 'kept']);
  // A candidate recorded earlier is removed once a folder is confirmed.
  const again = setUp({ status: 200, text: LISTING }, { props: { [CANDIDATE_KEY]: MBS_DRAFTS } });
  check('a confirmation removes an earlier "not confirmed" record', [again.revisionsUrl, again.candidate], [MBS_DRAFTS, null]);
}

// ====================================================================
console.log('4. a main-meeting report is not touched');
{
  const mainMeeting = [Object.assign({}, GM.getMeetings86172[0], { Id: 60778, Type: 'OR', Title: '3GPPSA4#137', MtgDocURL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Synthetic/Docs/' })];
  const created = createReport(mainMeeting, [], '60778', { reportFamily: 'Video' });
  const r = setUp({ status: 200, text: LISTING }, { created: created });
  check('it is set up as a main-meeting report', [r.created.created.ok, r.setup.status, r.props.getProperty('MEETING_TYPE'), r.props.getProperty('MEETING_FOLDER')], [true, 'configured', 'main', 'TSGS4_137_Synthetic']);
  check('nothing is asked, nothing is stored, nothing is recorded', [r.requests, r.revisionsUrl, r.candidate, r.folder], [[], null, null, { status: 'not-applicable', url: '', reason: '' }]);
  check('its drafts folder is the one the formula gives, as before', r.run.s.getMeetingContext_().sources.revisionsUrl, 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Synthetic/Inbox/Drafts/Video');
  check('nothing is said about it', r.run.s.describeRevisionsFolderOutcome_(r.folder), '');
}

// ====================================================================
console.log('5. a Video ad-hoc series');
{
  const VIDEO_DOCS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_VIDEO/Docs/';
  const VIDEO_DRAFTS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_VIDEO/inbox/drafts/';
  const videoMeeting = [Object.assign({}, GM.getMeetings86172[0], { Id: 86999, Title: '3GPPSA4-e (AH) Video SWG post 137-e', ShortTitle: '3GPPSA4-e (AH) Video SWG post 137-e', StartDate: '2026-10-06 15:00:00', MtgDocURL: VIDEO_DOCS })];
  const created = createReport(videoMeeting, [{ id: 'S4aV260010', type: 'discussion', revisionOf: null, agendaItem: '2.1' }], '86999', {});
  check('a Video ad-hoc report is created (its family from the meeting name, its TDocs and its folder)', [created.created.ok, payloadOf(created.s, created.description).config.reportType, payloadOf(created.s, created.description).config.ftpBase], [true, 'Video', VIDEO_DOCS]);
  const r = setUp({ status: 200, text: LISTING }, { created: created });
  check('the folder of the Video series is asked for and stored', [r.requests.map((q) => q[0]), r.revisionsUrl, r.run.s.getMeetingContext_().sources.revisionsUrl], [[VIDEO_DRAFTS], VIDEO_DRAFTS, VIDEO_DRAFTS]);
  const refused = setUp({ status: 403, text: '' }, { created: created });
  check('and when it cannot be confirmed, it is shown as the candidate', [refused.revisionsUrl, refused.candidate], [null, VIDEO_DRAFTS]);
  check('nothing in the code names a Video folder: the rule is the one of every ad-hoc series (source)',
    [/SA4_VIDEO|SA4_MBS|SA4_Audio|SA4_Plenary/.test(withoutComments(functionSource(CREATOR, 'confirmAdhocRevisionsFolderWith_'))), /SA4_VIDEO/.test(withoutComments(CREATOR)), /SA4_VIDEO/.test(withoutComments(CODE)),
      /deriveRevisionsUrlCandidate_\(cfg\.ftpBase\)/.test(functionSource(CREATOR, 'confirmAdhocRevisionsFolderWith_'))], [false, false, false, true]);
  // The other captured series, through the same rule.
  const audio = setUp({ status: 200, text: LISTING }, { created: createReport(GM.getMeetings85916, [{ id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '2' }], '85916', {}) });
  check('the Audio series 85916 gets its own', audio.revisionsUrl, 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Audio/inbox/drafts/');
}

// ====================================================================
console.log('6. only an https address on a 3GPP host is asked');
{
  const run = firstRunFor('');
  const asked = [];
  const confirm = (cfg, props) => {
    const store = firstRunFor('', { props: props }).docProps;
    const out = plain(run.s.confirmAdhocRevisionsFolderWith_({ documentProperties: store, validateRevisionsCandidate: (url) => { asked.push(url); return { ok: true, reason: null }; } }, cfg));
    return [out.status, store.getProperty('REVISIONS_URL'), store.getProperty(CANDIDATE_KEY)];
  };
  check('another host, http, an address in disguise: nothing is asked and nothing is stored',
    [confirm({ meetingType: 'adhoc', ftpBase: 'https://example.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/' }),
      confirm({ meetingType: 'adhoc', ftpBase: 'http://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/' }),
      confirm({ meetingType: 'adhoc', ftpBase: 'https://ftp.3gpp.org.example.org/x/Docs/' }),
      confirm({ meetingType: 'adhoc', ftpBase: 'https://ftp.3gpp.org@example.org/x/Docs/' }),
      confirm({ meetingType: 'adhoc', ftpBase: 'ftp://ftp.3gpp.org/x/Docs/' })],
    Array(5).fill(['no-candidate', null, null]));
  check('a document folder that is no "Docs/" folder, and none: no candidate',
    [confirm({ meetingType: 'adhoc', ftpBase: 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/' }), confirm({ meetingType: 'adhoc', ftpBase: '' }), confirm({ meetingType: 'adhoc' })], Array(3).fill(['no-candidate', null, null]));
  check('a main meeting, a meeting of no type, and no configuration: not applicable', [confirm({ meetingType: 'main', ftpBase: MBS_DOCS }), confirm({ ftpBase: MBS_DOCS }), confirm(null)], Array(3).fill(['not-applicable', null, null]));
  check('up to here nothing was asked at all', asked, []);
  check('the three 3GPP hosts are asked',
    [confirm({ meetingType: 'adhoc', ftpBase: MBS_DOCS }), confirm({ meetingType: 'adhoc', ftpBase: MBS_DOCS.replace('ftp.3gpp.org', 'www.3gpp.org') }), confirm({ meetingType: 'adhoc', ftpBase: MBS_DOCS.replace('ftp.3gpp.org', 'portal.3gpp.org') })].map((x) => x[0]).concat([asked.length]),
    ['stored', 'stored', 'stored', 3]);
  check('a document folder with a doubled slash (as the Portal gave it for 86178) derives a clean address', (() => { asked.length = 0; confirm({ meetingType: 'adhoc', ftpBase: 'https://ftp.3gpp.org//tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/' }); return asked; })(),
    ['https://ftp.3gpp.org/tsg_sa/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/inbox/drafts/']);
  const tell = run.s.describeRevisionsFolderOutcome_;
  check('what the user is told in the other cases', [tell({ status: 'no-candidate', url: '', reason: 'ftpBase is missing.' }), tell({ status: 'kept', url: MBS_DRAFTS }), tell({ status: 'not-applicable' }), tell({ status: 'not-checked', url: MBS_DRAFTS }), tell(null), tell(undefined)],
    ['⚠️ Drafts folder still to set up: none could be derived (ftpBase is missing.). Enter it in SA4 Report > ⚙️ Configure Meeting… (Revisions / Drafts URL).', '', '', '', '', '']);
}

// ====================================================================
console.log('7. once, after the setup, and not from the meeting lookup');
{
  const r = setUp({ status: 403, text: '' });
  check('when the folder is asked for, the report is already set up: a request that never returns cannot leave it half configured', r.stateWhenAsked, ['done']);
  const before = r.requests.length;
  const second = r.run.run();
  const third = r.run.run();
  check('running the setup again asks nothing', [second.status, third.status, r.requests.length - before, second.revisionsFolder === undefined], ['already-done', 'already-done', 0, true]);
  const finish = functionSource(CREATOR, 'finishReportSetupWith_');
  check('in the first run it is the last step, after the "done" marker (source)',
    [finish.indexOf('confirmAdhocRevisionsFolderWith_(') > finish.indexOf("props.setProperty(TEMPLATE_STATE_KEYS_.state, 'done')"), (finish.match(/confirmAdhocRevisionsFolderWith_\(/g) || []).length,
      finish.indexOf("if (state === 'done') return") < finish.indexOf('confirmAdhocRevisionsFolderWith_(')], [true, 1, true]);
  check('the live setup asks through validateRevisionsUrlCandidate_() -- one request, no redirect followed (source)',
    [/validateRevisionsCandidate: function \(url\) \{ return validateRevisionsUrlCandidate_\(url\); \}/.test(functionSource(CREATOR, 'liveTemplateDeps_')),
      /UrlFetchApp\.fetch\(url, \{ muteHttpExceptions: true, followRedirects: false \}\)/.test(functionSource(CODE, 'fetchRevisionsUrlCandidate_')), (functionSource(CODE, 'validateRevisionsUrlCandidate_').match(/fetchRevisionsUrlCandidate_\(/g) || []).length], [true, true, 1]);
  // The lookup of a meeting (Configure Meeting's Resolve, and the creator's "Look up") still asks no drafts folder.
  const callers = (src) => withoutComments(src).split('\nfunction ').filter((f) => /validateRevisionsUrlCandidate_\(|fetchRevisionsUrlCandidate_\(/.test(f.slice(f.indexOf('(')))).map((f) => f.slice(0, f.indexOf('(')));
  // (The step after a Save of Configure Meeting is the second place: tests/configure-meeting-setup-parity.test.js.)
  check('the folder is asked for from the first run and from the step after a Save of Configure Meeting only: the meeting lookup still derives it without a request (source)',
    [callers(CODE), callers(CREATOR)], [['validateRevisionsUrlCandidate_', 'diagnoseMeetingResolverTiming_'], ['afterTemplateConfigurationSaved_', 'liveTemplateDeps_']]);
  const lookedUp = createReport(GM.getMeetings86172, ROWS_MBS, '86172', {});
  check('creating the report asked nothing either: its network was down, and the report was created', [lookedUp.created.ok, payloadOf(lookedUp.s, lookedUp.description).config.revisionsUrl], [true, '']);
  check('the candidate is recorded under its own key, which no report code reads as a folder (source)',
    [(withoutComments(CODE).match(/SA4_REVISIONS_URL_CANDIDATE|revisionsCandidate/g) || []).length, /revisionsCandidate: 'SA4_REVISIONS_URL_CANDIDATE'/.test(CREATOR)], [0, true]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll ad-hoc revisions folder setup checks passed.');
process.exitCode = failures ? 1 : 0;
