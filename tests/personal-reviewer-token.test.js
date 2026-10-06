/**
 * The personal Reviewer API token (after T-2026.10.8).
 *
 * The Reviewer API token of a report is a Script Property of the report's
 * own script project. Every report created from the template is a new
 * project, so the token had to be entered again in each. A user can now keep
 * it once, for all their template reports, in one private settings file in
 * their own Google Drive. A token stored in a report still comes first:
 *
 *     the report's own token  ->  the personal token  ->  none
 *
 * The file is found by a marker (a Drive custom file property) among the
 * files the user owns -- never by its name -- and is trusted only when it is
 * owned by the user, shared with nobody, not in the trash, application/json,
 * at most 2 KB, and exactly { schema, reviewerApiToken }. With MORE THAN ONE
 * such file none is used: a secret is not picked from several candidates.
 *
 *    1. the pure rules: a usable token, the file content, the file metadata;
 *    2. precedence; no Drive query when the report has its own token;
 *    3. zero, one, several files; every way a file is not trusted;
 *    4. Drive fails: no personal token, and nothing else is affected;
 *    5. looked up once per execution, and only when the Reviewer is asked;
 *    6. one place reads the token of a report (source);
 *    7. Configure Meeting: what it says;
 *    8. Configure Meeting: every action;
 *    9. write, read back, and only then remove the report's token;
 *   10. several files, a shared file, a file with other content: writing;
 *   11. a stored token is in no HTML, no answer to the browser, no log;
 *   12. an update with a personal token; whose token an update uses;
 *   13. CENTRAL / Legacy: nothing changed.
 *
 * Google Drive is the fake of tests/helpers/fake-drive-files.js. All tokens
 * are synthetic strings.
 *
 * Run: node tests/personal-reviewer-token.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { makeFakeDrive, EXPECTED_QUERY, NAME } = require('./helpers/fake-drive-files.js');
const { configurableReport, RELEASE } = require('./helpers/config-dialog.js');
const { templateReport } = require('./helpers/template-report.js');

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
const plain = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)));

const LOCAL = 'LOCAL-token-1111-aaaa';
const PERSONAL = 'PERSONAL-token-2222-bbbb';
const NEW = 'NEW-token-3333-cccc';
const OTHER = 'OTHER-token-4444-dddd';
const SETTINGS = (token) => JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: token });
const MBS_DOCS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/';
// A configured ad-hoc report that has its drafts folder (so that a Save has nothing else to say).
const PROPS = { MEETING_ID: '86172', MEETING_TYPE: 'adhoc', MEETING_NAME: 'Synthetic ad-hoc', MEETING_DATE: 'October 1, 2026', REPORT_SUFFIX: 'MBS', FTP_BASE: MBS_DOCS,
  REVISIONS_URL: MBS_DOCS.replace('Docs/', 'inbox/drafts/'), TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172' };

/** A script project with Code.js and (in a template report) ReportCreator.js, the given token of its own, and the Drive of its user. */
function runtime(o) {
  const opts = o || {};
  const loaded = opts.release === false ? loadCode({ documentProperties: PROPS, scriptProperties: opts.local ? { REVIEWER_API_TOKEN: opts.local } : {} })
    : loadTemplateRuntime({ release: RELEASE, documentProperties: PROPS, scriptProperties: opts.local ? { REVIEWER_API_TOKEN: opts.local } : {} });
  const s = loaded.sandbox;
  if (opts.release === false) s.SA4_RELEASE_ = undefined;
  const drive = opts.drive || makeFakeDrive();
  drive.install(s);
  const logs = [];
  s.Logger = { log: (m) => logs.push(String(m)) };
  return { s: s, drive: drive, logs: logs, scriptProps: loaded.scriptProps, docProps: loaded.docProps,
    /** One execution: nothing remembered from the one before. */
    resolve: () => { s.resetReviewerTokenRunState_(); return plain(s.resolveReviewerApiToken_()); } };
}
/** The Drive of a user with these files; `files` are arguments of add(). */
function driveWith(files, options) {
  const drive = makeFakeDrive(options);
  (files || []).forEach((f) => drive.add(f));
  return drive;
}
const NONE = { token: '', source: 'none', problem: '' };

// =====================================================================
console.log('1. the pure rules');
{
  const s = runtime().s;
  check('a usable token: text of 1 to 512 characters without space around it and without a line break or control character',
    [['a', 'crv1_AbC-123.xyz', 'x'.repeat(512), 'two words'].map((v) => s.isUsableReviewerTokenValue_(v)),
      ['', ' x', 'x ', 'x'.repeat(513), 'a\nb', 'a\rb', 'a\tb', 'a\u0000b', '\u007F', null, undefined, 12345, {}, ['x']].map((v) => s.isUsableReviewerTokenValue_(v))],
    [[true, true, true, true], Array(14).fill(false)]);
  check('the content written for a token is exactly the schema, and is read back as that token', [s.serializeUserSettings_(PERSONAL), s.parseUserSettingsText_(s.serializeUserSettings_(PERSONAL))], [SETTINGS(PERSONAL), PERSONAL]);
  check('content that is anything else gives no token',
    ['', '   ', 'not json', '{', '[]', 'null', '"' + PERSONAL + '"', '[' + SETTINGS(PERSONAL) + ']',
      JSON.stringify({ schema: 'sa4-report-user-settings/2', reviewerApiToken: PERSONAL }), JSON.stringify({ schema: 'SA4-report-user-settings/1', reviewerApiToken: PERSONAL }),
      JSON.stringify({ reviewerApiToken: PERSONAL }), JSON.stringify({ schema: 'sa4-report-user-settings/1' }), JSON.stringify({ schema: 'sa4-report-user-settings/1', token: PERSONAL }),
      JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: PERSONAL, extra: 1 }), JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: '' }),
      JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: ' ' + PERSONAL }), JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: 'a\nb' }),
      JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: 12345 }), JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: 'x'.repeat(513) }),
      SETTINGS(PERSONAL) + ' '.repeat(2048)].map((text) => s.parseUserSettingsText_(text)).concat([s.parseUserSettingsText_(null), s.parseUserSettingsText_(undefined), s.parseUserSettingsText_({ schema: 1 })]),
    Array(23).fill(''));
  const META = { id: 'f1', mimeType: 'application/json', ownedByMe: true, shared: false, trashed: false, properties: { sa4ReportUserSettings: '1' } };
  const with_ = (changes) => s.classifyUserSettingsFile_(Object.assign({}, META, changes));
  check('a file is private only when every one of these is exactly so: owned by the user, not shared, not in the trash, application/json, marked',
    [with_({}), with_({ shared: true }), with_({ shared: undefined }), with_({ shared: 'false' }), with_({ ownedByMe: false }), with_({ ownedByMe: undefined }), with_({ ownedByMe: 'true' }), with_({ trashed: true }), with_({ trashed: undefined }),
      with_({ mimeType: 'text/plain' }), with_({ mimeType: 'application/vnd.google-apps.document' }), with_({ properties: undefined }), with_({ properties: {} }), with_({ properties: { sa4ReportUserSettings: '2' } }),
      with_({ properties: { sa4ReportUserSettings: 1 } }), with_({ id: '' }), s.classifyUserSettingsFile_(null), s.classifyUserSettingsFile_('f1')],
    ['private', 'shared', 'shared', 'shared', 'other', 'other', 'other', 'other', 'other', 'other', 'other', 'other', 'other', 'other', 'other', 'other', 'other', 'other']);
  check('the file is looked for by its marker among the files the user owns -- never by its name (source)',
    [s.USER_SETTINGS_QUERY_, /name\s*(=|contains)/.test(s.USER_SETTINGS_QUERY_), /getFilesByName|searchFiles|title\s*=/.test(withoutComments(CREATOR)), /appProperties/.test(withoutComments(CREATOR)),
      /getUserProperties/.test(withoutComments(CREATOR)) || /getUserProperties/.test(withoutComments(CODE))],
    [EXPECTED_QUERY, false, false, false, false]);
  check('its name is for the reader only', [s.USER_SETTINGS_FILE_NAME_, s.USER_SETTINGS_SCHEMA_, s.USER_SETTINGS_MAX_BYTES_], [NAME, 'sa4-report-user-settings/1', 2048]);
}

// =====================================================================
console.log('2. precedence');
{
  const none = runtime();
  check('neither: no token', [none.resolve(), none.drive.count('list')], [NONE, 1]);
  const personal = runtime({ drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  check('a personal token only: it is used', personal.resolve(), { token: PERSONAL, source: 'user', problem: '' });
  const local = runtime({ local: LOCAL });
  check('a token of the report only: it is used, and Drive is not asked', [local.resolve(), local.drive.calls], [{ token: LOCAL, source: 'report', problem: '' }, []]);
  const both = runtime({ local: LOCAL, drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  check('both: the token of the report, and Drive is not asked at all', [both.resolve(), both.drive.calls], [{ token: LOCAL, source: 'report', problem: '' }, []]);
  check('getReviewerApiTokenForRun_() is that token', [both.s.getReviewerApiTokenForRun_(), personal.s.getReviewerApiTokenForRun_(), none.s.getReviewerApiTokenForRun_()], [LOCAL, PERSONAL, '']);
  both.scriptProps.deleteProperty('REVIEWER_API_TOKEN');
  check('the report loses its token: the personal one is used from the next execution on', both.resolve(), { token: PERSONAL, source: 'user', problem: '' });
  // The report's token is used as it is stored (as before): nothing about its form is checked.
  const odd = runtime({ local: 'a token with spaces', drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  check('a token of the report is used as it is stored, whatever its form (as before)', odd.resolve().token, 'a token with spaces');
}

// =====================================================================
console.log('3. zero, one, several files; every way a file is not trusted');
{
  const resolveWith = (files, options) => { const r = runtime({ drive: driveWith(files, options) }); const out = r.resolve(); return [out.token, out.source, out.problem, r.drive.count('read')]; };
  check('no file', resolveWith([]), ['', 'none', '', 0]);
  check('exactly one usable file', resolveWith([{ content: SETTINGS(PERSONAL) }]), [PERSONAL, 'user', '', 1]);
  check('two usable files: NONE is used', resolveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER) }]), ['', 'none', 'more than one private settings file exists, so none is used', 2]);
  check('two usable files with the same token: still none -- which file is the one is not known', resolveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(PERSONAL) }]),
    ['', 'none', 'more than one private settings file exists, so none is used', 2]);
  check('three', resolveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER) }, { content: SETTINGS(NEW) }])[0], '');
  check('one usable file beside one that is shared: the usable one; the shared one is not even read', resolveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER), shared: true }]), [PERSONAL, 'user', '', 1]);
  check('one usable file beside one with other content: the usable one', resolveWith([{ content: SETTINGS(PERSONAL) }, { content: 'something else' }]), [PERSONAL, 'user', '', 2]);

  // Each of these is the ONLY marked file. The fake returns it whatever the query says (leaky), so that it is this code that rejects it.
  const alone = (file) => resolveWith([Object.assign({ content: SETTINGS(PERSONAL) }, file)], { leaky: true });
  check('owned by somebody else (shared with the user): not used, not read', alone({ ownedByMe: false }), ['', 'none', '', 0]);
  check('shared by the user with somebody: not used, not read, and said so', alone({ shared: true }), ['', 'none', 'the settings file is shared', 0]);
  check('in the trash: not used, not read', alone({ trashed: true }), ['', 'none', '', 0]);
  check('another type of file: not used, not read', [alone({ mimeType: 'text/plain' }), alone({ mimeType: 'application/vnd.google-apps.document' })], [['', 'none', '', 0], ['', 'none', '', 0]]);
  check('without the marker (a file that merely has the name): not used, not read', [alone({ properties: undefined }), alone({ properties: { other: '1' } })], [['', 'none', '', 0], ['', 'none', '', 0]]);
  check('larger than 2 KB: not used, not read', alone({ content: SETTINGS(PERSONAL) + ' '.repeat(2048) }), ['', 'none', 'the settings file has other content than expected', 0]);
  check('another schema', alone({ content: JSON.stringify({ schema: 'sa4-report-user-settings/2', reviewerApiToken: PERSONAL }) }), ['', 'none', 'the settings file has other content than expected', 1]);
  check('a field more than the schema', alone({ content: JSON.stringify({ schema: 'sa4-report-user-settings/1', reviewerApiToken: PERSONAL, note: 'x' }) })[0], '');
  check('a token that cannot be used (empty, with a line break, too long)',
    [alone({ content: SETTINGS('') })[0], alone({ content: SETTINGS('a\nb') })[0], alone({ content: SETTINGS('x'.repeat(513)) })[0]], ['', '', '']);
  check('content that is no JSON', alone({ content: '{ "schema": "sa4-report-user-settings/1", "reviewerApiToken": "' + PERSONAL }), ['', 'none', 'the settings file has other content than expected', 1]);
  // Without the leak: the query itself already leaves out what is not the user's own, in the trash, of another type or unmarked.
  check('the query leaves those out as well: a foreign, a trashed, an unmarked and a text file beside the user\'s own -- one candidate',
    resolveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER), ownedByMe: false }, { content: SETTINGS(OTHER), trashed: true }, { content: SETTINGS(OTHER), properties: undefined }, { content: SETTINGS(OTHER), mimeType: 'text/plain' }]),
    [PERSONAL, 'user', '', 1]);
  check('and with the leak the same files are still rejected one by one', resolveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER), ownedByMe: false }, { content: SETTINGS(OTHER), trashed: true },
    { content: SETTINGS(OTHER), properties: undefined }, { content: SETTINGS(OTHER), mimeType: 'text/plain' }], { leaky: true }), [PERSONAL, 'user', '', 1]);
  // A lookup never changes Drive.
  const r = runtime({ drive: driveWith([{ content: SETTINGS(PERSONAL), shared: true }, { content: 'junk' }, { content: SETTINGS(OTHER) }, { content: SETTINGS(NEW) }]) });
  const before = JSON.stringify(r.drive.files);
  r.resolve();
  check('a lookup writes nothing, creates nothing and moves nothing to the trash', [JSON.stringify(r.drive.files) === before, r.drive.calls.map((c) => c[0]).filter((n) => n !== 'list' && n !== 'read')], [true, []]);
}

// =====================================================================
console.log('4. Drive fails');
{
  const listFails = driveWith([{ content: SETTINGS(PERSONAL) }]);
  listFails.failList = true;
  const a = runtime({ drive: listFails });
  check('the search fails: no personal token, a reason in fixed words, nothing thrown', a.resolve(), { token: '', source: 'none', problem: 'Google Drive could not be read' });
  const readFails = driveWith([{ content: SETTINGS(PERSONAL) }]);
  readFails.failRead = true;
  const b = runtime({ drive: readFails });
  check('the file cannot be read: the same', b.resolve(), { token: '', source: 'none', problem: 'Google Drive could not be read' });
  const oneOfTwo = driveWith([{ id: 'good', content: SETTINGS(PERSONAL) }, { id: 'bad', content: SETTINGS(OTHER) }]);
  oneOfTwo.failRead = (id) => id === 'bad';
  check('one of two files cannot be read: no token (the unread one may be a second usable file)', runtime({ drive: oneOfTwo }).resolve().token, '');
  const noDrive = runtime();
  noDrive.s.Drive = undefined;
  check('no Drive service in the project at all: no personal token, nothing thrown', noDrive.resolve(), { token: '', source: 'none', problem: 'Google Drive could not be read' });
  const localAnyway = driveWith([]);
  localAnyway.failList = true;
  const c = runtime({ local: LOCAL, drive: localAnyway });
  check('with a token of its own a report does not notice any of it', [c.resolve(), c.drive.calls], [{ token: LOCAL, source: 'report', problem: '' }, []]);
  // Without ReportCreator.js beside a release (not a real bundle): the resolver still does not throw.
  const bare = loadCode({ documentProperties: PROPS });
  bare.sandbox.SA4_RELEASE_ = RELEASE;
  check('the resolver itself never throws', plain(bare.sandbox.resolveReviewerApiToken_()), { token: '', source: 'none', problem: 'the personal settings could not be looked up' });
  // What a failure writes into the log: fixed words, and nothing of any file.
  const SECRET_JUNK = '{"schema":"sa4-report-user-settings/1","reviewerApiToken":"JUNK-SECRET-9999" BROKEN';
  const junk = runtime({ drive: driveWith([{ content: SECRET_JUNK }]) });
  junk.resolve();
  check('content that cannot be parsed is not repeated in the log -- not even by the parser\'s own message',
    [junk.logs, junk.logs.some((l) => /JUNK-SECRET|BROKEN|Unexpected|JSON|position/i.test(l))], [['The personal Reviewer token is not used: the settings file has other content than expected.'], false]);
  check('no code logs the message of a JSON parse error (source)', /catch \(e\) \{\n    return '';\n  \}/.test(functionSource(CREATOR, 'parseUserSettingsText_')), true);
  check('and each failure is logged in fixed words', [a.logs, b.logs].map((l) => l.filter((x) => /personal Reviewer token/.test(x))),
    [['The personal Reviewer token is not used: Google Drive could not be read.'], ['The personal Reviewer token is not used: Google Drive could not be read.']]);
}

// =====================================================================
console.log('5. looked up once per execution, and only when the Reviewer is asked');
{
  const r = runtime({ drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  r.s.resetReviewerTokenRunState_();
  for (let i = 0; i < 25; i++) { r.s.getReviewerApiTokenForRun_(); r.s.resolveReviewerApiToken_(); }
  check('fifty requests for the token in one execution: one search, one read', [r.drive.count('list'), r.drive.count('read')], [1, 1]);
  r.s.resetReviewerTokenRunState_();
  r.s.getReviewerApiTokenForRun_();
  check('the next execution looks again, once', [r.drive.count('list'), r.drive.count('read')], [2, 2]);
  const empty = runtime();
  empty.s.resetReviewerTokenRunState_();
  for (let i = 0; i < 10; i++) empty.s.getReviewerApiTokenForRun_();
  check('"no personal token" is remembered for the execution as well: one search, not ten', empty.drive.count('list'), 1);

  // The configuration of a report is read by nearly everything; it must not look the token up.
  const cfgRun = runtime({ drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  cfgRun.s.resetReviewerTokenRunState_();
  const cfg = cfgRun.s.getReportConfig_();
  cfgRun.s.getMeetingContext_();
  const copy = Object.assign({}, cfg);
  const printed = JSON.stringify(cfg) + JSON.stringify(copy) + Object.keys(cfg).join();
  check('reading, copying and printing the configuration asks Drive nothing, and has no token in it', [cfgRun.drive.calls, /token/i.test(printed), printed.indexOf(PERSONAL)], [[], false, -1]);
  check('the token can still be had from it by name, and is looked up then', [cfg.REVIEWER_API_TOKEN, cfgRun.drive.count('list')], [PERSONAL, 1]);
  check('getReportConfig_() no longer looks the token up for every configuration (source)',
    [/getReviewerApiTokenForRun_\(\); \/\/ TEMPLATE-002C/.test(CODE), /Object\.defineProperty\(reportConfig, 'REVIEWER_API_TOKEN', \{ enumerable: false, get: function \(\) \{ return getReviewerApiTokenForRun_\(\); \} \}\);/.test(functionSource(CODE, 'getReportConfig_'))], [false, true]);
}
{
  // The real update. Abstracts off, nothing uploaded since: the Reviewer is not asked, so no lookup.
  const drive = driveWith([{ content: SETTINGS(PERSONAL) }]);
  const report = templateReport({ token: false, drive: drive, tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'available', uploaded: true }, { id: 'S4aA269002', agenda: '4.3', status: 'reserved' }] });
  report.s.completeInsertedUploadedTdoc_ = () => false;
  report.exec(() => report.s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }));
  drive.calls.length = 0;
  const off = report.update();
  check('an update with abstracts off: completes, and asks Drive nothing for a token', [off.success, drive.calls], [true, []]);
  // Abstracts on: two tables without an abstract (one not uploaded, so left out) -> one request, one lookup.
  report.docProps.setProperty('FETCH_ABSTRACTS_ON_UPDATE', 'true');
  report.docProps.setProperty('REVIEWER_NO_SUMMARY_CACHE_x', '');
  const sentKeys = [];
  const realFetch = report.s.UrlFetchApp.fetch;
  report.s.UrlFetchApp.fetch = (url, options) => { if (/reviewer\./.test(url)) sentKeys.push(options && options.headers && options.headers['X-API-Key']); return realFetch(url, options); };
  report.reviewer.S4aA269001 = 'Summary one.';
  drive.calls.length = 0;
  const on = report.update();
  check('an update with abstracts on: the abstract is fetched with the personal token', [on.success, report.requests, sentKeys, report.field('S4aA269001', 'Abstract')], [true, ['S4aA269001'], [PERSONAL], 'Summary one.']);
  check('for that, one search and one read -- however many tables', [drive.count('list'), drive.count('read'), drive.calls.map((c) => c[0]).filter((n) => n !== 'list' && n !== 'read')], [1, 1, []]);
}
{
  // Many tables, one lookup; and a build.
  const drive = driveWith([{ content: SETTINGS(PERSONAL) }]);
  const tdocs = Array.from({ length: 6 }, (_, i) => ({ id: 'S4aA26900' + (i + 1), agenda: '4.3', status: 'available', uploaded: true }));
  const report = templateReport({ token: false, drive: drive, tdocs: tdocs });
  const built = report.build();
  check('Build Report from Scratch with a personal token: its abstracts phase runs, on one lookup for six tables',
    [built.ok, built.phases.filter((p) => p.name === 'abstracts').map((p) => p.status), report.requests.length, drive.count('list'), drive.count('read')], [true, ['done'], 6, 1, 1]);
  // A report where only withdrawn TDocs lack an abstract: the Reviewer is not asked, so the token is not looked up either.
  const drive2 = driveWith([{ content: SETTINGS(PERSONAL) }]);
  const withdrawn = templateReport({ token: false, drive: drive2, props: { FETCH_ABSTRACTS_ON_UPDATE: 'true' }, tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'withdrawn', uploaded: true }] });
  withdrawn.s.completeInsertedUploadedTdoc_ = () => false;
  withdrawn.exec(() => withdrawn.s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }));
  drive2.calls.length = 0;
  withdrawn.update();
  check('an update whose only candidate is withdrawn: no Reviewer request and no token lookup', [withdrawn.requests, drive2.calls], [[], []]);
}

// =====================================================================
console.log('6. one place reads the token of a report');
{
  const all = withoutComments(CODE) + withoutComments(CREATOR);
  const reads = (all.match(/getProperty\(\s*['"]REVIEWER_API_TOKEN['"]\s*\)/g) || []).length;
  check('getProperty(\'REVIEWER_API_TOKEN\') occurs once in the production code', reads, 1);
  check('and that is in resolveReviewerApiToken_()', (withoutComments(functionSource(CODE, 'resolveReviewerApiToken_')).match(/getProperty\('REVIEWER_API_TOKEN'\)/g) || []).length, 1);
  check('the name of the property occurs elsewhere only where it is written or removed, and in one message',
    withoutComments(all).split('\n').filter((line) => /REVIEWER_API_TOKEN/.test(line) && !/getProperty\('REVIEWER_API_TOKEN'\)/.test(line)).map((line) => line.trim()),
    ["Object.defineProperty(reportConfig, 'REVIEWER_API_TOKEN', { enumerable: false, get: function () { return getReviewerApiTokenForRun_(); } });",
      "Logger.log('No REVIEWER_API_TOKEN found in script properties');", "scriptProps.deleteProperty('REVIEWER_API_TOKEN');", "scriptProps.setProperty('REVIEWER_API_TOKEN', config.apiToken.trim());",
      "removeLocalToken: function () { PropertiesService.getScriptProperties().deleteProperty('REVIEWER_API_TOKEN'); }"]);
  const viaResolver = (name) => /resolveReviewerApiToken_\((true)?\)/.test(withoutComments(functionSource(CODE, name)));
  check('every reader asks the resolver: the run accessor, the dialog, the two connection tests, the configuration check',
    ['getReviewerApiTokenForRun_', 'configureMeetingSettings', 'testAllConnections', 'testReviewerApi', 'validateConfiguration'].map(viaResolver), [true, true, true, true, true]);
  check('and what asks the Reviewer for an abstract asks the run accessor', [/getReviewerApiTokenForRun_\(\)/.test(functionSource(CODE, 'fetchAndAddAbstract_')), /getReviewerApiTokenForRun_\(\)/.test(functionSource(CODE, 'runFullReportBuildCore_'))], [true, true]);
  // The diagnostics with a personal token.
  const r = configurableReport({ props: PROPS, drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  const keys = [];
  r.s.UrlFetchApp = { fetch: (url, options) => { if (/reviewer\./.test(url)) keys.push(options.headers['X-API-Key']); return { getResponseCode: () => 200, getContentText: () => '{}' }; } };
  r.s.resetReviewerTokenRunState_();
  r.s.testReviewerApi();
  check('"Test Reviewer API" uses the personal token of a report without its own', [keys, r.alerts.some((a) => /not configured/i.test(a.join(' ')))], [[PERSONAL], false]);
}

// =====================================================================
console.log('7. Configure Meeting: what it says');
const STATUS = (page) => (page.html.match(/id="tokenStatus">([^<]*)</) || [])[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&');
function dialog(o) {
  const opts = o || {};
  const r = configurableReport({ props: PROPS, scriptProps: opts.local ? { REVIEWER_API_TOKEN: opts.local } : {}, drive: opts.drive || makeFakeDrive(), release: opts.release });
  r.saved = [];
  const realSave = r.s.saveConfigurationSettings;
  r.s.saveConfigurationSettings = (config) => { const out = realSave(config); r.saved.push(out === undefined ? null : plain(out)); return out; };
  r.s.resetReviewerTokenRunState_();
  return r;
}
{
  const texts = {
    none: STATUS(dialog().open()),
    personal: STATUS(dialog({ drive: driveWith([{ content: SETTINGS(PERSONAL) }]) }).open()),
    local: STATUS(dialog({ local: LOCAL }).open()),
    both: STATUS(dialog({ local: LOCAL, drive: driveWith([{ content: SETTINGS(PERSONAL) }]) }).open())
  };
  check('no token', texts.none, 'No Reviewer API token. Abstracts are skipped.');
  check('a personal token', texts.personal, 'Using your personal Reviewer API token (kept privately in your Google Drive, for all your reports).');
  check('a token of the report', texts.local, 'Using the Reviewer API token stored in this report. Editors of this report can see it.');
  check('both: the report\'s is in use', texts.both, 'Using the Reviewer API token stored in this report. Your personal token is not used here.');
  const two = driveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER) }]);
  check('several private settings files: none is used, and what to do',
    STATUS(dialog({ drive: two }).open()), 'No Reviewer API token. Abstracts are skipped. Your Google Drive has more than one private settings file ("' + NAME + '"), so none of them is used. ' +
    'Tick "Remove my personal token" to move them all to the trash and save the token again, or delete the extra files in Drive.');
  check('the same beside a token of the report (which goes on working)',
    /^Using the Reviewer API token stored in this report\. Editors of this report can see it\. Your Google Drive has more than one private settings file/.test(STATUS(dialog({ local: LOCAL, drive: driveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER) }]) }).open())), true);
  check('the settings file is shared', STATUS(dialog({ drive: driveWith([{ content: SETTINGS(PERSONAL), shared: true }]) }).open()),
    'No Reviewer API token. Abstracts are skipped. Your personal settings file ("' + NAME + '") is shared with others, so its token is not used. Treat that token as disclosed: delete the file in Drive, get a new token and save it.');
  check('the settings file has other content', STATUS(dialog({ drive: driveWith([{ content: 'junk' }]) }).open()),
    'No Reviewer API token. Abstracts are skipped. Your personal settings file ("' + NAME + '") has other content than expected and is not used. Save the token again to replace it.');
  check('the text is put into the page as text (escaped), like every other value of the dialog',
    [/id="tokenStatus">[^<]*\(&quot;SA4 Report – private settings \(do not share\)\.json&quot;\)/.test(dialog({ drive: driveWith([{ content: 'junk' }]) }).open().html), /id="tokenStatus">[^<]*"SA4 Report/.test(dialog({ drive: driveWith([{ content: 'junk' }]) }).open().html)], [true, false]);
  const down = makeFakeDrive();
  down.failList = true;
  check('Drive cannot be read: the dialog still opens, and says so', STATUS(dialog({ drive: down }).open()), 'No Reviewer API token. Abstracts are skipped. Your personal token could not be checked: Google Drive could not be read.');
  check('a usable file beside a shared one', STATUS(dialog({ drive: driveWith([{ content: SETTINGS(PERSONAL) }, { content: SETTINGS(OTHER), shared: true }]) }).open()),
    'Using your personal Reviewer API token (kept privately in your Google Drive, for all your reports). Another settings file in your Drive ("' + NAME + '") is shared and is not used; delete it.');

  const page = dialog({ local: LOCAL, drive: driveWith([{ content: SETTINGS(PERSONAL) }]) }).open();
  check('the controls: a token field, where it goes (personal by default), removing either, and saving the report\'s token as the personal one',
    ['apiToken', 'apiTokenTargetPersonal', 'apiTokenTargetReport', 'promoteApiToken', 'clearApiToken', 'clearPersonalApiToken'].map((id) => [new RegExp('id="' + id + '"').test(page.html), !!(page.fields[id] && page.fields[id].checked)]),
    [[true, false], [true, true], [true, false], [true, false], [true, false], [true, false]]);
  check('"save the token of this report as my personal token" is offered only when the report has one',
    [/id="promoteApiToken"/.test(dialog({ drive: driveWith([{ content: SETTINGS(PERSONAL) }]) }).open().html), /id="promoteApiToken"/.test(dialog().open().html), /id="promoteApiToken"/.test(dialog({ local: LOCAL }).open().html)], [false, false, true]);
  check('it says whose tokens an automatic update uses', /Automatic updates run as the person who switched them on, and use the tokens of that person\./.test(page.html), true);
  check('opening the dialog looks at Drive once, and changes nothing there', (() => { const d = driveWith([{ content: SETTINGS(PERSONAL) }]); const before = JSON.stringify(d.files); dialog({ drive: d }).open(); return [d.count('list'), JSON.stringify(d.files) === before, d.calls.map((c) => c[0]).filter((n) => n !== 'list' && n !== 'read')]; })(), [1, true, []]);
}

// =====================================================================
console.log('8. Configure Meeting: every action');
const LOCAL_OF = (r) => r.scriptProps.getProperty('REVIEWER_API_TOKEN');
const FILES_OF = (drive) => drive.live().map((f) => [f.content, f.shared]);
{
  // A new token, for all my reports (the default).
  const r = dialog();
  const page = r.open();
  page.fields.apiToken.value = '  ' + NEW + '  ';
  page.save();
  check('a new token, "for all my reports": one private settings file with exactly the schema; nothing is stored in the report',
    [FILES_OF(r.drive), r.drive.files.map((f) => [f.name, f.mimeType, f.ownedByMe, f.properties]), LOCAL_OF(r)], [[[SETTINGS(NEW), false]], [[NAME, 'application/json', true, { sa4ReportUserSettings: '1' }]], null]);
  check('the browser sent it as the personal token, and nothing for the report', [page.sent[0].apiTokenAction, page.sent[0].apiToken, page.sent[0].personalTokenAction], ['keep', '', 'replace']);
  check('the message of the Save says where it is', [page.alerts.length, page.alerts[0].indexOf('✅ Your personal Reviewer token is saved privately in your Google Drive ("' + NAME + '") and is used in all your template reports.') !== -1], [1, true]);
  check('the report uses it from then on', r.resolve ? null : (() => { r.s.resetReviewerTokenRunState_(); return plain(r.s.resolveReviewerApiToken_()); })(), { token: NEW, source: 'user', problem: '' });
  // A second report of the same user: a fresh project, the same Drive.
  const second = dialog({ drive: r.drive });
  check('ANOTHER report of the same user finds it, with nothing stored in that report', [plain(second.s.resolveReviewerApiToken_()), second.scriptProps.getKeys(), STATUS(second.open())],
    [{ token: NEW, source: 'user', problem: '' }, [], 'Using your personal Reviewer API token (kept privately in your Google Drive, for all your reports).']);
}
{
  // A new token, for this report only.
  const r = dialog({ drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  const page = r.open();
  page.fields.apiToken.value = NEW;
  page.fields.apiTokenTargetPersonal.checked = false;
  page.fields.apiTokenTargetReport.checked = true;
  page.save();
  check('a new token, "this report only": stored in the report; the personal file is not touched',
    [LOCAL_OF(r), FILES_OF(r.drive), r.drive.calls.map((c) => c[0]).filter((n) => n !== 'list' && n !== 'read'), page.sent[0].apiTokenAction, page.sent[0].personalTokenAction],
    [NEW, [[SETTINGS(PERSONAL), false]], [], 'replace', undefined]);
  check('nothing is said about the personal token', /personal Reviewer token/.test(page.alerts[0]), false);
}
{
  // Replacing the personal token.
  const drive = driveWith([{ id: 'mine', content: SETTINGS(PERSONAL) }]);
  const r = dialog({ drive: drive });
  const page = r.open();
  page.fields.apiToken.value = NEW;
  page.save();
  check('replacing the personal token rewrites the one file; no second file appears', [drive.files.map((f) => [f.id, f.content, f.trashed]), drive.count('create'), drive.count('write')], [[['mine', SETTINGS(NEW), false]], 0, 1]);
}
{
  // Removing the personal token.
  const drive = driveWith([{ id: 'mine', content: SETTINGS(PERSONAL) }]);
  const r = dialog({ local: LOCAL, drive: drive });
  const page = r.open();
  page.fields.clearPersonalApiToken.checked = true;
  page.save();
  check('"Remove my personal token": the file goes to the trash (it can be restored); the token of the report stays',
    [drive.files.map((f) => [f.id, f.trashed]), LOCAL_OF(r), page.sent[0].personalTokenAction, /✅ Your personal Reviewer token was removed \(1 settings file moved to the trash of your Google Drive\)\./.test(page.alerts[0])], [[['mine', true]], LOCAL, 'clear', true]);
  const nothing = dialog();
  const nothingPage = nothing.open();
  nothingPage.fields.clearPersonalApiToken.checked = true;
  nothingPage.save();
  check('with none to remove, it says so', /ℹ️ There was no personal Reviewer token to remove\./.test(nothingPage.alerts[0]), true);
}
{
  // Removing the token of the report.
  const drive = driveWith([{ content: SETTINGS(PERSONAL) }]);
  const r = dialog({ local: LOCAL, drive: drive });
  const page = r.open();
  page.fields.clearApiToken.checked = true;
  page.save();
  r.s.resetReviewerTokenRunState_();
  check('"Remove the token of this report": removed at once; the personal one is then in use, untouched',
    [LOCAL_OF(r), FILES_OF(drive), plain(r.s.resolveReviewerApiToken_()), page.sent[0].apiTokenAction, page.sent[0].personalTokenAction], [null, [[SETTINGS(PERSONAL), false]], { token: PERSONAL, source: 'user', problem: '' }, 'clear', undefined]);
}
{
  // Saving the token of the report as the personal one: on the server only.
  const r = dialog({ local: LOCAL });
  const page = r.open();
  page.fields.promoteApiToken.checked = true;
  page.save();
  check('"Save the token of this report as my personal token": the personal file holds it; the report keeps its own',
    [FILES_OF(r.drive), LOCAL_OF(r), page.sent[0].personalTokenAction, page.sent[0].apiTokenAction], [[[SETTINGS(LOCAL), false]], LOCAL, 'promote', 'keep']);
  check('the token itself never went to the browser or came from it: not in the page, not in what was sent, not in the answer',
    [page.html.indexOf(LOCAL), JSON.stringify(page.sent).indexOf(LOCAL), JSON.stringify(r.saved).indexOf(LOCAL), page.alerts.join('\n').indexOf(LOCAL)], [-1, -1, -1, -1]);
  const none = dialog();
  none.s.saveConfigurationSettings(Object.assign({ meetingId: '86172', meetingType: 'adhoc', reportType: 'MBS', mailingListMode: 'derived', apiTokenAction: 'keep' }, { personalTokenAction: 'promote' }));
  check('asked for a report without a token of its own: nothing is written, and it says so', [none.drive.live().length, none.saved[0].note], [0, '⚠️ Nothing was saved as your personal Reviewer token: this report has no token of its own.']);
}
{
  // One thing at a time.
  const r = dialog({ local: LOCAL, drive: driveWith([{ content: SETTINGS(PERSONAL) }]) });
  const page = r.open();
  page.fields.apiToken.value = NEW;
  page.fields.clearPersonalApiToken.checked = true;
  page.save();
  check('a new personal token AND its removal: the page asks to choose, and sends nothing', [page.alerts, page.sent.length, FILES_OF(r.drive), LOCAL_OF(r)],
    [['Choose one thing for your personal token: a new token, the token of this report, or removing it.'], 0, [[SETTINGS(PERSONAL), false]], LOCAL]);
  page.fields.clearPersonalApiToken.checked = false;
  page.fields.promoteApiToken.checked = true;
  page.save();
  check('a new personal token AND the report\'s token as the personal one: the same', [page.alerts.length, page.sent.length], [2, 0]);
  page.fields.apiToken.value = '';
  page.fields.clearPersonalApiToken.checked = true;
  page.save();
  check('saving the report\'s token as personal AND removing the personal one: the same', [page.alerts.length, page.sent.length], [3, 0]);
}
{
  // What the server refuses, before anything is stored.
  const r = dialog({ local: LOCAL });
  const tryIt = (extra) => { try { r.s.saveConfigurationSettings(Object.assign({ meetingId: '86999', meetingType: 'adhoc', reportType: 'MBS', mailingListMode: 'derived', apiTokenAction: 'clear' }, extra)); return null; } catch (e) { return e.message; } };
  check('an action that does not exist, and a token that cannot be used, are refused',
    [tryIt({ personalTokenAction: 'publish' }), tryIt({ personalTokenAction: 'replace', personalToken: '' }), tryIt({ personalTokenAction: 'replace', personalToken: 'x'.repeat(513) }), tryIt({ personalTokenAction: 'replace', personalToken: 'a\nb' })],
    ['Unknown action for the personal Reviewer token.', 'The Reviewer token cannot be stored: it is empty, longer than 512 characters, or contains a line break.',
      'The Reviewer token cannot be stored: it is empty, longer than 512 characters, or contains a line break.', 'The Reviewer token cannot be stored: it is empty, longer than 512 characters, or contains a line break.']);
  check('and then nothing at all was stored: not the configuration, not a file, and the report\'s token is still there', [r.docProps.getProperty('MEETING_ID'), r.drive.files.length, LOCAL_OF(r)], ['86172', 0, LOCAL]);
  const plan = r.s.beforeTemplateConfigurationSaved_({ meetingId: '1', apiTokenAction: 'keep', personalTokenAction: 'replace', personalToken: ' ' + NEW + ' ' });
  check('what is stored as configuration never contains the personal token', [JSON.stringify(plan.config).indexOf(NEW), Object.keys(plan.config).sort(), plan.personalToken === NEW], [-1, ['apiTokenAction', 'meetingId'], true]);
}

// =====================================================================
console.log('9. write, read back, and only then remove the token of the report');
{
  /** A Save that makes the report's token the personal one and removes it from the report; `prepare(drive)` breaks Drive first. */
  const promoteAndRemove = (prepare, typed) => {
    const drive = makeFakeDrive();
    const r = dialog({ local: LOCAL, drive: drive });
    const order = [];
    const realDelete = r.scriptProps.deleteProperty;
    r.scriptProps.deleteProperty = (k) => { if (k === 'REVIEWER_API_TOKEN') order.push('remove the token of the report'); return realDelete(k); };
    const page = r.open();
    if (prepare) prepare(drive);
    drive.calls.length = 0;
    if (typed) page.fields.apiToken.value = typed; else page.fields.promoteApiToken.checked = true;
    page.fields.clearApiToken.checked = true;
    const realPush = drive.calls.push.bind(drive.calls);
    drive.calls.push = (c) => { if (['create', 'write', 'meta', 'read'].indexOf(c[0]) !== -1) order.push(c[0]); return realPush(c); };
    page.save();
    return { r: r, drive: drive, page: page, order: order, local: LOCAL_OF(r), files: FILES_OF(drive), alert: page.alerts[0] };
  };
  const ok = promoteAndRemove();
  check('it is done: the personal file holds the token, and the report has none', [ok.files, ok.local], [[[SETTINGS(LOCAL), false]], null]);
  check('in this order: the file is written, read back by its id (metadata and content), and only then is the report\'s token removed', ok.order.slice(-4), ['create', 'meta', 'read', 'remove the token of the report']);
  check('the message says both', [/✅ Your personal Reviewer token is saved privately/.test(ok.alert), /The token of this report was removed\./.test(ok.alert)], [true, true]);
  check('the report then works with the personal token', (() => { ok.r.s.resetReviewerTokenRunState_(); return plain(ok.r.s.resolveReviewerApiToken_()); })(), { token: LOCAL, source: 'user', problem: '' });

  const cases = [
    ['the file cannot be created', (d) => { d.failCreate = true; }, 'The settings file could not be written to Google Drive.'],
    ['Drive cannot be searched', (d) => { d.failList = true; }, 'Google Drive could not be read.'],
    ['the write reports success but nothing is stored', (d) => { d.dropWrites = true; }, 'The settings file was written but could not be confirmed as private and complete. The personal token is not counted as stored.'],
    ['the new file comes out shared', (d) => { d.createShared = true; }, 'The settings file was written but could not be confirmed as private and complete. The personal token is not counted as stored.'],
    ['the written file cannot be read back', (d) => { d.failRead = true; }, 'The settings file was written but could not be confirmed as private and complete. The personal token is not counted as stored.'],
    ['its metadata cannot be read back', (d) => { d.failMeta = true; }, 'The settings file was written but could not be confirmed as private and complete. The personal token is not counted as stored.']
  ];
  cases.forEach(([name, prepare, reason]) => {
    const failed = promoteAndRemove(prepare);
    check(name + ': the report KEEPS its token, and the Save says so', [failed.local, failed.order.indexOf('remove the token of the report'), failed.alert.indexOf('⚠️ Your personal Reviewer token was NOT saved: ' + reason + ' The token of this report was kept.') !== -1, /Configuration saved\./.test(failed.alert)],
      [LOCAL, -1, true, true]);
    failed.r.s.resetReviewerTokenRunState_();
    check(name + ': and it goes on working with it', failed.r.s.resolveReviewerApiToken_().token, LOCAL);
  });
  const typed = promoteAndRemove(null, NEW);
  check('the same order for a new personal token typed together with "remove the token of this report"', [typed.files, typed.local, typed.order.slice(-4)], [[[SETTINGS(NEW), false]], null, ['create', 'meta', 'read', 'remove the token of the report']]);
  const typedFailed = promoteAndRemove((d) => { d.failCreate = true; }, NEW);
  check('and when that write fails, the report keeps its token as well', [typedFailed.local, typedFailed.files, /The token of this report was kept\./.test(typedFailed.alert)], [LOCAL, [], true]);
  // Removing the report's token without touching the personal one is not deferred: it is what was always done.
  check('the removal is held back only when a personal token is written in the same Save (source)',
    /if \(\(action === 'replace' \|\| action === 'promote'\) && copy\.apiTokenAction === 'clear'\) \{\n    copy\.apiTokenAction = 'keep';\n    plan\.clearLocalAfter = true;\n  \}/.test(functionSource(CREATOR, 'beforeTemplateConfigurationSaved_')), true);
}

// =====================================================================
console.log('10. several files, a shared file, a file with other content: writing');
{
  const save = (drive, o) => { const r = dialog({ drive: drive, local: o && o.local }); const page = r.open(); drive.calls.length = 0; if (o && o.clear) page.fields.clearPersonalApiToken.checked = true; else page.fields.apiToken.value = NEW; page.save(); return { r: r, page: page, alert: page.alerts[0] }; };
  const two = driveWith([{ id: 'a', content: SETTINGS(PERSONAL) }, { id: 'b', content: SETTINGS(OTHER) }]);
  const refused = save(two);
  check('several usable files, a new token: NOTHING is written -- neither of them, and no third file',
    [two.files.map((f) => [f.id, f.content, f.trashed]), two.calls.map((c) => c[0]).filter((n) => n !== 'list' && n !== 'read'), /⚠️ Your personal Reviewer token was NOT saved: Your Google Drive has more than one private settings file/.test(refused.alert)],
    [[['a', SETTINGS(PERSONAL), false], ['b', SETTINGS(OTHER), false]], [], true]);
  const cleared = save(two, { clear: true });
  check('"Remove my personal token" is how the user clears that: all of them go to the trash', [two.files.map((f) => [f.id, f.trashed]), /\(2 settings files moved to the trash/.test(cleared.alert)], [[['a', true], ['b', true]], true]);
  save(two);
  check('and a token saved after that is in one new file', [two.live().map((f) => f.content), two.live().length], [[SETTINGS(NEW)], 1]);

  const twoJunk = driveWith([{ id: 'a', content: 'junk' }, { id: 'b', content: 'other junk' }]);
  check('several private marked files of other content: nothing is written either',
    [/NOT saved: Your Google Drive has more than one private settings file/.test(save(twoJunk).alert), twoJunk.files.map((f) => f.content), twoJunk.count('create') + twoJunk.count('write')], [true, ['junk', 'other junk'], 0]);

  const oneJunk = driveWith([{ id: 'a', content: 'junk' }]);
  save(oneJunk);
  check('one private marked file of other content: it is the user\'s own and private, and is rewritten -- no second file', [oneJunk.files.map((f) => [f.id, f.content]), oneJunk.count('create')], [[['a', SETTINGS(NEW)]], 0]);

  const shared = driveWith([{ id: 'leaked', content: SETTINGS(PERSONAL), shared: true }]);
  const beside = save(shared);
  check('a shared settings file: it is not written to, not unshared and not removed; the token goes into a new private file, and the Save says the shared one is still there',
    [shared.files.map((f) => [f.id, f.content, f.shared, f.trashed]).slice(0, 1), shared.live().length, shared.calls.filter((c) => c[1] === 'leaked').map((c) => c[0]), / Another settings file in your Drive is shared and is not used; delete it\./.test(beside.alert)],
    [[['leaked', SETTINGS(PERSONAL), true, false]], 2, [], true]);
  beside.r.s.resetReviewerTokenRunState_();
  check('after that there is exactly one usable file, so no ambiguity was created', plain(beside.r.s.resolveReviewerApiToken_()), { token: NEW, source: 'user', problem: '' });
  const sharedCleared = save(driveWith([{ id: 'leaked', content: SETTINGS(PERSONAL), shared: true }, { id: 'mine', content: SETTINGS(OTHER) }]), { clear: true });
  check('removing the personal token leaves a shared file as it is, and says so', [sharedCleared.r.drive.files.map((f) => [f.id, f.trashed]), /A settings file that is shared was left as it is; delete it in Drive\./.test(sharedCleared.alert)], [[['leaked', false], ['mine', true]], true]);
  check('no code changes who a file is shared with (source)', /addEditor|addViewer|removeEditor|removeViewer|setSharing|Permissions\.|\.permissions/.test(withoutComments(CREATOR)), false);
}

// =====================================================================
console.log('11. a stored token is in no HTML, no answer to the browser, no log');
{
  const drive = driveWith([{ content: SETTINGS(PERSONAL) }]);
  const r = dialog({ local: LOCAL, drive: drive });
  const page = r.open();
  page.resolve && null;
  page.fields.promoteApiToken.checked = true;
  page.save();
  const second = r.open();
  second.fields.clearApiToken.checked = true;
  second.save();
  const third = r.open();
  third.fields.apiToken.value = NEW;
  third.save();
  r.s.resetReviewerTokenRunState_();
  r.s.resolveReviewerApiToken_();
  const leaks = (text) => [LOCAL, PERSONAL, NEW].filter((t) => String(text).indexOf(t) !== -1);
  check('three Saves later: no stored token is in any page that was rendered', leaks([page.html, second.html, third.html, r.dialogs.map((d) => d.html).join('')].join('')), []);
  check('none is in any answer the server gave to the browser', leaks(JSON.stringify(r.saved)), []);
  check('none is in any message shown, and none in the log (a typed one is masked there, as always)', [leaks(page.alerts.concat(second.alerts, third.alerts).join('\n')), leaks(r.logs.join('\n'))], [[], []]);
  check('none is in the document\'s own properties', leaks(JSON.stringify(r.docProps._store)), []);
  check('the only token the browser ever held is the one typed into it', [leaks(JSON.stringify(page.sent) + JSON.stringify(second.sent)), leaks(JSON.stringify(third.sent))], [[], [NEW]]);
  // The status text is made from two flags, and nothing of a token can be in it.
  check('what the dialog shows about the token is made of the state alone (source)',
    [/\.token\b/.test(withoutComments(functionSource(CREATOR, 'describeReviewerTokenStatus_'))) || !/^function describeReviewerTokenStatus_\(hasLocal, personal\) \{/.test(functionSource(CREATOR, 'describeReviewerTokenStatus_')),
      /\.token\b/.test(withoutComments(functionSource(CREATOR, 'templateReviewerTokenDialogParts_'))),
      /personal = \{ status: found\.status, sharedCount: found\.sharedCount \};/.test(functionSource(CREATOR, 'templateReviewerTokenDialogParts_'))], [false, false, true]);
  check('the personal token is never written into the report: no property of the report is set by this code (source)',
    /setProperty\(/.test(withoutComments(CREATOR.slice(CREATOR.indexOf('// The personal Reviewer API token: one private settings file')))), false);
}

// =====================================================================
console.log('12. whose token an update uses');
{
  // Two people edit one report. X switched automatic updates on, so the update runs as X: with X's Drive.
  const keyUsedWith = (drive, local) => {
    const report = templateReport({ token: false, drive: drive, props: { FETCH_ABSTRACTS_ON_UPDATE: 'true' }, tdocs: [{ id: 'S4aA269001', agenda: '4.3', status: 'available', uploaded: true }] });
    if (local) report.scriptProps.setProperty('REVIEWER_API_TOKEN', local);
    report.s.completeInsertedUploadedTdoc_ = () => false;
    report.exec(() => report.s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true }));
    const keys = [];
    const realFetch = report.s.UrlFetchApp.fetch;
    report.s.UrlFetchApp.fetch = (url, options) => { if (/reviewer\./.test(url)) keys.push(options.headers['X-API-Key']); return realFetch(url, options); };
    const out = report.update();
    return [out.success, keys];
  };
  check('run as X (who has a personal token): X\'s token', keyUsedWith(driveWith([{ content: SETTINGS('TOKEN-OF-X') }])), [true, ['TOKEN-OF-X']]);
  check('run as Y (who has another): Y\'s token -- an execution reads the Drive of the account it runs as', keyUsedWith(driveWith([{ content: SETTINGS('TOKEN-OF-Y') }])), [true, ['TOKEN-OF-Y']]);
  check('run as somebody without one: no request, and the update completes', keyUsedWith(makeFakeDrive()), [true, []]);
  check('a token stored in the report is used by everybody, before any personal one', keyUsedWith(driveWith([{ content: SETTINGS('TOKEN-OF-X') }]), LOCAL), [true, [LOCAL]]);
  const broken = driveWith([{ content: SETTINGS('TOKEN-OF-X') }, { content: SETTINGS('TOKEN-OF-X-2') }]);
  check('with several settings files the update completes without a token and without a request', keyUsedWith(broken), [true, []]);
  const down = makeFakeDrive();
  down.failList = true;
  check('with Drive down the update completes just the same', keyUsedWith(down), [true, []]);
  check('and with Drive down a report with its own token is not affected at all', keyUsedWith(down, LOCAL), [true, [LOCAL]]);
}

// =====================================================================
console.log('13. CENTRAL / Legacy (no template release): nothing changed');
{
  const drive = driveWith([{ content: SETTINGS(PERSONAL) }]);
  const central = runtime({ release: false, drive: drive });
  check('a personal settings file is not looked for and not used', [central.resolve(), drive.calls], [NONE, []]);
  const withLocal = runtime({ release: false, local: LOCAL, drive: drive });
  check('the token of the project is used as before', [withLocal.resolve(), drive.calls], [{ token: LOCAL, source: 'report', problem: '' }, []]);
  const page = dialog({ release: false, local: LOCAL, drive: drive }).open();
  check('the dialog has the token controls it had, and none of the new ones',
    [/<label>Replace Reviewer API Token:<\/label>\s*<input type="password" id="apiToken" value="" placeholder="Enter a new token to replace the current one" autocomplete="off">/.test(page.html),
      /Remove the saved Reviewer API token/.test(page.html), /<div class="hint">Used to fetch AI summaries and abstracts\. The saved token is never displayed\.<\/div>/.test(page.html),
      STATUS(page), /apiTokenTarget|clearPersonalApiToken|promoteApiToken|personal/i.test(page.html.slice(0, page.html.indexOf('<script>')))], [true, true, true, 'Reviewer API token configured', false]);
  check('without a token it says what it said', STATUS(dialog({ release: false, drive: drive }).open()), 'No Reviewer API token configured');
  const r = dialog({ release: false, drive: drive });
  const p = r.open();
  p.fields.apiToken.value = NEW;
  p.save();
  check('a typed token goes into the project, as before; Drive is never touched', [LOCAL_OF(r), p.sent[0].apiTokenAction, p.sent[0].apiToken, p.sent[0].personalTokenAction, drive.calls, r.saved], [NEW, 'replace', NEW, undefined, [], [null]]);
  // The dialog of the code before this work, for the same document: the same page, byte for byte.
  let before = null;
  try { before = execFileSync('git', ['show', 'template-release/T-2026.10.8:Code.js'], { cwd: path.join(__dirname, '..'), maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8'); } catch (e) { before = null; }
  if (!before) {
    console.log('  note: template-release/T-2026.10.8 is not available in this checkout; the comparison with it is skipped.');
  } else {
    const htmlOf = (source) => {
      const loaded = loadCode({ documentProperties: PROPS, scriptProperties: { REVIEWER_API_TOKEN: LOCAL } });
      const sandbox = { Logger: { log: () => {} }, PropertiesService: loaded.sandbox.PropertiesService, Session: loaded.sandbox.Session, Utilities: loaded.sandbox.Utilities, LockService: loaded.sandbox.LockService, console: console };
      let html = null;
      sandbox.DocumentApp = { getUi: () => ({ showModalDialog: (out) => { html = out.html; } }), getActiveDocument: () => ({ getId: () => 'DOC' }), ParagraphHeading: {}, ElementType: {} };
      sandbox.HtmlService = { createHtmlOutput: (h) => { const out = { html: h, setWidth: () => out, setHeight: () => out }; return out; } };
      vm.createContext(sandbox);
      vm.runInContext(source, sandbox, { filename: 'Code.js' });
      sandbox.configureMeetingSettings();
      return html;
    };
    const now = htmlOf(fs.readFileSync(CODE_JS_PATH, 'utf8'));
    const then = htmlOf(before);
    // The one difference of the page since T-2026.10.8 that is not about the token: the start Resolve found goes with Save (Batch 1).
    const withoutStart = (html) => html.replace(/\r/g, '')
      .replace("      // The start of the meeting Resolve found, as the Portal gives it, and\n      // the meeting it is of. Sent with Save only for that same meeting id.\n      let resolvedMeetingStart = null;\n", '')
      .replace("            resolvedMeetingStart = { meetingId: String(result.resolved.id), startDate: result.preview.startDateRaw || '', timeZone: result.preview.startTimeZoneRaw || '' };\n", '')
      .replace("        // The start Resolve found goes with the save of the meeting it was found for (see persistConfigurationSettings_()).\n        if (resolvedMeetingStart && resolvedMeetingStart.meetingId === String(config.meetingId || '').trim()) {\n          config.resolvedMeetingStart = resolvedMeetingStart;\n        }\n", '')
      .replace('.withSuccessHandler((saved) => {', '.withSuccessHandler(() => {').replace(" + (saved && saved.note ? '\\n\\n' + saved.note : '')", '');
    const markup = (html) => html.slice(0, html.indexOf('<script>'));
    check('the markup of the CENTRAL / Legacy dialog is byte for byte that of T-2026.10.8', [markup(now.replace(/\r/g, '')) === markup(then.replace(/\r/g, '')), markup(now).length > 3000], [true, true]);
    const scriptNow = withoutStart(now).slice(now.indexOf('<script>'));
    check('its script differs from T-2026.10.8 only by the lines that look for the personal-token controls, which it does not have',
      scriptNow.replace("        // A template report has controls for a personal token as well (see beforeTemplateConfigurationSaved_(),\n        // ReportCreator.js). Where they do not exist, nothing here is ticked and all is as it was.\n        const ticked = function (id) { const el = document.getElementById(id); return !!(el && el.checked); };\n        const tokenForAllReports = ticked('apiTokenTargetPersonal');\n", '')
        .replace('} else if (newToken && newToken.trim() && !tokenForAllReports) {', '} else if (newToken && newToken.trim()) {')
        .replace(/        \/\/ The personal token: one thing at a time[\s\S]*?          if \(personalTokenWishes\[0\] === 'replace'\) config\.personalToken = newToken;\n        \}\n/, '') === then.replace(/\r/g, '').slice(then.replace(/\r/g, '').indexOf('<script>')), true);
  }
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll personal Reviewer token checks passed.');
process.exitCode = failures ? 1 : 0;
