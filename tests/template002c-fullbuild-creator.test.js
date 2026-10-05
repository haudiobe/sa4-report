/**
 * TEMPLATE-002C -- regression tests for what the first live template report
 * (86178, template T-2026.10.0) exposed, and for the decisions taken with it:
 *
 *   1. Run Full Report Build: one confirmation, no pop-up between phases and
 *      (2.17.3) none at the end -- a completed build makes no UI call after
 *      changing the document; one formatting pass; the optional Reviewer token
 *      read once; phase timing; a failed phase stops the build and throws.
 *   2. The single-step menu items keep their own completion pop-ups.
 *   3. Create New SA4 Report: optional Mailing List override.
 *   4. Generated discussion e-mails: Reply-To = the effective mailing list.
 *   5. Discussion subject: no list/family tag; "[<agenda item>][<deadline>][<TDoc>]".
 *
 * The behaviour before the fix was measured by the diagnostic test on branch
 * diagnose/template-002b-fullbuild-timeout (a3021a7): four pop-ups inside
 * the build, two formatting passes, one token read and log line per TDoc.
 *
 * Run: node tests/template002c-fullbuild-creator.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { installExportStubs, blobText } = require('./helpers/email-export-stubs.js');

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
const plain = (v) => JSON.parse(JSON.stringify(v));

const FX = path.join(__dirname, 'fixtures');
const CSV_86172 = fs.readFileSync(path.join(FX, 'meeting-86172-agenda.csv'), 'utf8');
const LIST_86172 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-tdoc-list-values.json'), 'utf8')).values;
const GM = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-85916-getmeetings.json'), 'utf8'));
const FTP_MBS = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_MBS/Docs/';
const PROPS_86172 = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4-e (AH) MBS SWG post 137-e', MEETING_ID: '86172', REPORT_SUFFIX: 'MBS',
  FTP_BASE: FTP_MBS, TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172',
  AGENDA_CSV_URL: FTP_MBS.replace('Docs/', 'Agenda/agenda.csv') };

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CODE_ONLY = CODE.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');

// =====================================================================
// 1. FULL BUILD
// =====================================================================

/**
 * A report for 86172 in which the REAL skeleton build, the REAL step
 * functions and the REAL Reviewer handling run. Only the network collectors
 * (e-mail, revisions) and the formatter are recorded instead of run.
 */
function fullBuildReport(options) {
  const o = options || {};
  const loaded = loadCode({ documentProperties: o.props || PROPS_86172, scriptProperties: o.token ? { REVIEWER_API_TOKEN: o.token } : {} });
  const s = loaded.sandbox;
  const body = makeFakeDocumentBody(s);
  const events = [];
  const alerts = [];
  const logs = [];
  s.Logger = { log: (m) => logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => body, getId: () => 'FAKE_DOC' });
  s.DocumentApp.openById = () => ({ getBody: () => makeFakeDocumentBody(s) });
  s.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  s.DocumentApp.getUi = () => ({
    alert: (title, message) => {
      // o.uiFailsAfterWork: as live on T-2026.10.2 -- once the build has
      // changed the document, a ui.alert() throws instead of appearing.
      if (o.uiFailsAfterWork && events.some((e) => /^WORK/.test(e))) {
        events.push('ALERT ATTEMPT ' + title);
        throw new Error('Service Documents failed while accessing document with id FAKE_DOC.');
      }
      alerts.push([title, message]); events.push('ALERT ' + title); return o.decline ? 'NO' : 'YES';
    },
    Button: { YES: 'YES', NO: 'NO' }, ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }
  });
  s.setDocumentTitleFromTemplate_ = () => {};
  s.findHeading_ = () => null;
  s.styleStatusCell_ = () => {};
  const reviewerRequests = [];
  s.UrlFetchApp = { fetch: (u) => {
    if (/reviewer\./.test(u)) { reviewerRequests.push(u); events.push('REVIEWER request'); return { getResponseCode: () => 404, getContentText: () => '' }; }
    return /agenda\.csv$/.test(u)
      ? { getResponseCode: () => 200, getContentText: () => CSV_86172 }
      : { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 3, 4], setName() { return this; } }) };
  } };
  s.DriveApp = { createFile: () => ({ setTrashed() {} }) };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => LIST_86172, getRichTextValues: () => LIST_86172.map((r) => r.map(() => null)) }) }] }) };
  // Recorded instead of run:
  s.removeRowHeightAndSpacing = () => events.push('WORK formatting');
  s.checkRSSFeed_ = () => { events.push('WORK e-mail'); if (o.failEmail) throw new Error('RSS is down'); };
  s.updateRevisions_ = () => events.push('WORK revisions');
  // Count how often the token is looked up.
  let tokenReads = 0;
  const realScriptProps = s.PropertiesService.getScriptProperties();
  s.PropertiesService.getScriptProperties = () => Object.assign({}, realScriptProps, {
    getProperty: (k) => { if (k === 'REVIEWER_API_TOKEN') tokenReads++; return realScriptProps.getProperty(k); }
  });
  const tdocTables = () => body.getTables().filter((t) => { try { return t.getRow(0).getCell(0).getText() === 'TDoc' && /^S4/.test(t.getRow(0).getCell(1).getText()); } catch (e) { return false; } });
  return { s, body, events, alerts, logs, reviewerRequests, tokenReads: () => tokenReads, tdocTables, docProps: loaded.docProps };
}
const fullBuildLog = (r) => r.logs.filter((l) => /^\[FULLBUILD\]/.test(l));
/** Runs the menu function; returns what it returned and what it threw. */
function runMenu(r) { try { return { returned: r.s.runFullReportBuild(), thrown: null }; } catch (e) { return { returned: undefined, thrown: e.message }; } }

console.log('1. Run Full Report Build: one confirmation, then no UI call at all');
{
  const r = fullBuildReport();
  const run = runMenu(r);
  check('dialogs: exactly the confirmation; no pop-up between phases and none at the end', r.alerts.map((a) => a[0]), ['Run Full Report Build']);
  check('the confirmation is unchanged', [/^This will run all steps:/.test(r.alerts[0][1]), /Continue\?$/.test(r.alerts[0][1])], [true, true]);
  check('nothing suspends the build: confirmation, then only work', r.events, ['ALERT Run Full Report Build', 'WORK e-mail', 'WORK revisions', 'WORK formatting']);
  check('a completed build ends normally (no error thrown)', run.thrown, null);
  check('the skeleton was built once, with the real TDoc tables', [r.logs.filter((l) => /^Done: Built skeleton/.test(l)).length, r.tdocTables().length > 0], [1, true]);
  check('e-mail core once, revisions core once', [r.events.filter((e) => e === 'WORK e-mail').length, r.events.filter((e) => e === 'WORK revisions').length], [1, 1]);
  check('phases start in this order', fullBuildLog(r).filter((l) => /: start$/.test(l)).map((l) => l.replace(/^\[FULLBUILD\] |: start$/g, '')),
    ['skeleton', 'e-mail', 'revisions', 'abstracts', 'formatting']);
  check('the total line is the last thing logged', /^\[FULLBUILD\] total: \d+ ms/.test(r.logs[r.logs.length - 1]), true);
  const summary = r.s.formatFullBuildSummary_(run.returned);
  check('the result still describes every phase and the total (available to the log and to tests, not shown in a pop-up)',
    [run.returned.ok, /^Full report build completed\./.test(summary), /Skeleton \+ TDoc tables: done \(\d+ agenda items, \d+ TDocs\)/.test(summary), /E-mail discussion: done/.test(summary),
      /Revisions: done/.test(summary), /Abstracts: skipped \(no Reviewer API token configured\)/.test(summary), /Formatting: done/.test(summary), /Total: \d+ s$/.test(summary)],
    [true, true, true, true, true, true, true, true]);
}

console.log('1. the live failure (T-2026.10.2): a UI that fails once the document has changed');
{
  const r = fullBuildReport({ uiFailsAfterWork: true });
  const run = runMenu(r);
  check('the completed build is not turned into a failure', [run.thrown, run.returned && run.returned.ok], [null, true]);
  check('because no UI call is even attempted after the work', r.events.filter((e) => /^ALERT ATTEMPT/.test(e)), []);
  check('all five phases ran and the total was logged', [fullBuildLog(r).filter((l) => /: (done|skipped) in \d+ ms/.test(l)).length, /^\[FULLBUILD\] total/.test(r.logs[r.logs.length - 1])], [5, true]);
  const wrapper = CODE.slice(CODE.indexOf('function runFullReportBuild()'), CODE.indexOf('function runFullReportBuildCore_()')).replace(/^\s*\/\/.*$/gm, '');
  check('after runFullReportBuildCore_() the menu function makes no UI or document call (source)',
    /ui\.|getUi|DocumentApp|DriveApp|saveAndClose/.test(wrapper.slice(wrapper.indexOf('runFullReportBuildCore_()'))), false);
  // T-2026.10.7: one save, in the step that ends a build or an update of a report with status dropdowns -- after the last
  // change of the document, and never before a UI call. A runtime without the Google Docs API service has the five phases above.
  const finalize = CODE_ONLY.slice(CODE_ONLY.indexOf('function finalizeStatusDropdowns_('), CODE_ONLY.indexOf('\n}\n', CODE_ONLY.indexOf('function finalizeStatusDropdowns_(')));
  check('the build still never flushes, and saves in one place only: the status-dropdown step at its very end; it still runs in one execution',
    [/\.flush\s*\(/.test(CODE_ONLY), (CODE_ONLY.match(/saveAndClose\s*\(/g) || []).length, (finalize.match(/saveAndClose\s*\(/g) || []).length, (CODE.match(/phase\('/g) || []).length,
      /if \(statusDropdownsAvailable_\(\)\) \{\n    phase\('status dropdowns'/.test(CODE)], [false, 1, 1, 6, true]);
}

console.log('1. declining the confirmation runs nothing');
{
  const r = fullBuildReport({ decline: true });
  const run = runMenu(r);
  check('one question, no work, no further dialog, no error', [r.events, r.body.getNumChildren(), fullBuildLog(r), run.thrown], [['ALERT Run Full Report Build'], 0, [], null]);
}

console.log('1. formatting: once, after the enrichment phases');
{
  const r = fullBuildReport();
  r.s.runFullReportBuild();
  check('exactly one full formatting pass', r.events.filter((e) => e === 'WORK formatting').length, 1);
  check('it is the last piece of work, after e-mail and revisions', r.events.filter((e) => /^WORK/.test(e)), ['WORK e-mail', 'WORK revisions', 'WORK formatting']);
  check('the formatter treats the whole document from scratch, so an earlier pass adds nothing a later one does not redo (source)',
    [/function removeRowHeightAndSpacing\(context\) \{\n  const body = getReportBody_\(context\);\n  const tables = getTablesCounted_\(body, 'removeRowHeightAndSpacing'\);/.test(CODE),
      /for \(let i = 0; i < tables\.length; i\+\+\)/.test(CODE.slice(CODE.indexOf('function removeRowHeightAndSpacing(context)'), CODE.indexOf('function removeRowHeightAndSpacing(context)') + 1500))],
    [true, true]);
}

console.log('1. Reviewer token absent: checked once, said once, no request');
{
  const r = fullBuildReport();
  r.s.runFullReportBuild();
  check('the token is looked up once in the whole build', r.tokenReads(), 1);
  check('zero Reviewer requests', r.reviewerRequests, []);
  check('its absence is logged once (the skipped phase), never per TDoc',
    [r.logs.filter((l) => /no Reviewer API token configured/.test(l) && /^\[FULLBUILD\] abstracts: skipped/.test(l)).length,
      r.logs.filter((l) => l === 'No REVIEWER_API_TOKEN found in script properties').length], [1, 0]);
  check('no abstract row was added', r.tdocTables().some((t) => { for (let i = 0; i < t.getNumRows(); i++) if (t.getRow(i).getCell(0).getText() === 'Abstract') return true; return false; }), false);
}

console.log('1. Reviewer token present: the abstracts phase does the work, once per table');
{
  const r = fullBuildReport({ token: 'TESTONLY-not-a-real-token' });
  r.s.runFullReportBuild();
  const n = r.tdocTables().length;
  check('one Reviewer request per TDoc table -- not two', [r.reviewerRequests.length, n > 0], [n, true]);
  check('the token is still read once', r.tokenReads(), 1);
  check('every request happens in the abstracts phase (after e-mail and revisions, before formatting)',
    [r.events.indexOf('REVIEWER request') > r.events.indexOf('WORK revisions'), r.events.lastIndexOf('REVIEWER request') < r.events.indexOf('WORK formatting')], [true, true]);
  check('the phase reports what it did', fullBuildLog(r).some((l) => new RegExp('^\\[FULLBUILD\\] abstracts: done in \\d+ ms -- ' + n + ' candidate table\\(s\\), 0 abstract\\(s\\) inserted$').test(l)), true);
  check('the token never appears in a log line or in the summary', [r.logs.join('\n').indexOf('TESTONLY'), r.s.formatFullBuildSummary_(r.s.runFullReportBuildCore_()).indexOf('TESTONLY')], [-1, -1]);
}

console.log('1. phase timing in the log');
{
  const r = fullBuildReport();
  r.s.runFullReportBuild();
  const lines = fullBuildLog(r);
  ['skeleton', 'e-mail', 'revisions', 'abstracts', 'formatting'].forEach((name) => {
    check(`${name}: a start line and a result line with its duration`,
      [lines.indexOf('[FULLBUILD] ' + name + ': start') !== -1, lines.some((l) => new RegExp('^\\[FULLBUILD\\] ' + name + ': (done|skipped) in \\d+ ms').test(l))], [true, true]);
  });
  check('one total line naming every phase', lines.filter((l) => /^\[FULLBUILD\] total: \d+ ms -- skeleton done \d+ ms, e-mail done \d+ ms, revisions done \d+ ms, abstracts skipped \d+ ms, formatting done \d+ ms$/.test(l)).length, 1);
  check('concise: two lines per phase plus the total', lines.length, 11);
}

console.log('1. a failed phase stops the build and stays a visible failure');
{
  const r = fullBuildReport({ failEmail: true });
  const run = runMenu(r);
  check('later phases are not started, and no pop-up is used to report it', r.events, ['ALERT Run Full Report Build', 'WORK e-mail']);
  check('the failure propagates: the menu function throws, so Docs shows it and the execution is "Failed"',
    [run.thrown !== null, /^Full report build failed: RSS is down/.test(run.thrown || '')], [true, true]);
  check('the error names the failed phase and what did not run, once',
    [/E-mail discussion: FAILED/.test(run.thrown), /Not run: Revisions, Abstracts, Formatting\./.test(run.thrown), (run.thrown.match(/RSS is down/g) || []).length], [true, true, 1]);
  check('it is logged too', [r.logs.filter((l) => l === 'Full report build failed: e-mail: RSS is down').length,
    fullBuildLog(r).filter((l) => !/total/.test(l)).slice(-1)[0].replace(/in \d+ ms/, 'in N ms')], [1, '[FULLBUILD] e-mail: failed in N ms -- RSS is down']);
  check('the skeleton that was already built stays', r.tdocTables().length > 0, true);

  const uiDown = fullBuildReport({ failEmail: true, uiFailsAfterWork: true });
  const run2 = runMenu(uiDown);
  check('a real failure is reported even when the UI cannot show anything', [/^Full report build failed: RSS is down/.test(run2.thrown || ''), uiDown.events.filter((e) => /^ALERT ATTEMPT/.test(e))], [true, []]);

  ['skeleton', 'revisions', 'formatting'].forEach((name) => {
    const f = fullBuildReport();
    if (name === 'skeleton') f.s.buildSkeletonWithTdocTables = () => { throw new Error('boom in ' + name); };
    if (name === 'revisions') f.s.updateRevisions_ = () => { throw new Error('boom in ' + name); };
    if (name === 'formatting') f.s.removeRowHeightAndSpacing = () => { throw new Error('boom in ' + name); };
    const res = runMenu(f);
    check(`an error in the ${name} phase fails the build`, [/^Full report build failed: boom in /.test(res.thrown || ''), f.alerts.map((a) => a[0])], [true, ['Run Full Report Build']]);
  });
  const abs = fullBuildReport({ token: 'TESTONLY-not-a-real-token' });
  abs.s.addAbstractsForTables_ = () => { throw new Error('boom in abstracts'); };
  check('an error in the abstracts phase fails the build', /^Full report build failed: boom in abstracts/.test(runMenu(abs).thrown || ''), true);

  const notReady = fullBuildReport({ props: { MEETING_TYPE: 'adhoc', MEETING_ID: '86172' } });
  const nr = runMenu(notReady);
  check('a build that is not ready fails in the skeleton phase, before the document is touched',
    [notReady.alerts.map((a) => a[0]), /^Full report build failed: Cannot build report yet\./.test(nr.thrown || ''), notReady.body.getNumChildren(), notReady.events.filter((e) => /^WORK/.test(e))],
    [['Run Full Report Build'], true, 0, []]);
}

console.log('1. the summary text (pure)');
{
  const { sandbox: s } = loadCode();
  const text = s.formatFullBuildSummary_({ ok: true, totalMs: 62400, phases: [
    { name: 'skeleton', status: 'done', detail: '14 agenda items, 34 TDocs', ms: 17200 }, { name: 'e-mail', status: 'done', detail: '', ms: 27000 },
    { name: 'revisions', status: 'done', detail: '', ms: 3100 }, { name: 'abstracts', status: 'skipped', detail: 'no Reviewer API token configured', ms: 2 },
    { name: 'formatting', status: 'done', detail: '', ms: 15000 }] });
  check('readable summary', text.split('\n'), ['Full report build completed.', '',
    '✓ Skeleton + TDoc tables: done (14 agenda items, 34 TDocs), 17 s', '✓ E-mail discussion: done, 27 s', '✓ Revisions: done, 3 s',
    '– Abstracts: skipped (no Reviewer API token configured), 0 s', '✓ Formatting: done, 15 s', '', 'Total: 62 s']);
}

// =====================================================================
// 2. SINGLE-STEP MENU ITEMS KEEP THEIR POP-UPS
// =====================================================================

console.log('2. single-step menu items: their own completion pop-ups remain');
{
  const r = fullBuildReport();
  r.s.buildSkeletonWithTdocTables();
  check('Build Skeleton + TDOC Tables: its "Done" pop-up, and it still formats',
    [r.alerts.map((a) => a[0]), /^Built skeleton with \d+ agenda items and \d+ TDOCs\./.test(r.alerts[0][1]), r.events], [['Done'], true, ['WORK formatting', 'ALERT Done']]);
  r.alerts.length = 0;
  r.s.collectEmailDiscussionOnly();
  check('Collect E-mail Discussion', r.alerts, [['Success', 'E-mail discussion collection completed.']]);
  r.alerts.length = 0;
  r.s.collectRevisionsOnly();
  check('Collect Revisions', r.alerts, [['Success', 'Revision collection completed.']]);
  r.alerts.length = 0;
  const n = r.tdocTables().length;
  r.s.addAbstractsOnly();
  check('Add Abstracts: always runs, reports its counts', r.alerts, [['Success', `Abstract step completed: ${n} candidate table(s) processed, 0 abstract(s) inserted.`]]);
  check('... and without a token it says so once, not once per table', r.logs.filter((l) => l === 'No REVIEWER_API_TOKEN found in script properties').length, 1);
}

console.log('2. the token is re-read after Configure Meeting changes it');
{
  const loaded = loadCode();
  const s = loaded.sandbox;
  check('no token', s.getReviewerApiTokenForRun_(), '');
  s.saveConfigurationSettings({ apiTokenAction: 'replace', apiToken: 'TESTONLY-new', showPreview: true });
  check('a saved token is seen at once', s.getReviewerApiTokenForRun_(), 'TESTONLY-new');
  s.saveConfigurationSettings({ apiTokenAction: 'clear', showPreview: true });
  check('a cleared token too', s.getReviewerApiTokenForRun_(), '');
}

// =====================================================================
// 3. CREATOR: OPTIONAL MAILING LIST
// =====================================================================

const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.3', flavor: 'template', codeVersion: '2.17.3', gitCommit: 'abcdef0000000000000000000000000000000000', templateDocumentId: TEMPLATE_ID };
// The 6G ad-hoc of the live test, as the Portal describes it (captured 86172 record, re-labelled).
const MEETING_86178 = (() => { const m = plain(GM.getMeetings86172); m[0].Id = 86178; m[0].Title = '3GPPSA4-e (AH) on FS_6G_MED';
  m[0].MtgDocURL = 'https://ftp.3gpp.org/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Docs/'; return m; })();
const MAIN_137 = (() => { const m = plain(GM.getMeetings86172); m[0].Type = 'OR'; m[0].Title = '3GPPSA4#137';
  m[0].MtgDocURL = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137_Xian/Docs/'; return m; })();

function fakeDrive() {
  const drive = { files: {}, calls: [], seq: 0 };
  drive.add = (id, name) => { drive.files[id] = { id, name, description: '' }; return drive.files[id]; };
  drive.file = (id) => {
    const f = drive.files[id];
    return { getId: () => f.id, getName: () => f.name, getUrl: () => 'https://docs.google.com/document/d/' + f.id + '/edit',
      getDescription: () => f.description, setDescription: (t) => { f.description = t; }, setTrashed: () => {},
      getParents: () => ({ hasNext: () => true, next: () => 'FOLDER' }),
      makeCopy: (name) => { drive.calls.push(['makeCopy', name]); return drive.file(drive.add('COPYdoc' + String(++drive.seq).padStart(30, '0'), name).id); } };
  };
  return drive;
}

function templateRuntime(docId, drive, meeting, props) {
  const loaded = loadTemplateRuntime({ release: RELEASE, documentProperties: props || {} });
  const s = loaded.sandbox;
  if (!drive.files[docId]) drive.add(docId, 'doc');
  const body = makeFakeDocumentBody(s);
  const dialogs = [];
  s.DocumentApp.getActiveDocument = () => ({ getId: () => docId, getBody: () => body });
  s.DocumentApp.getUi = () => ({ alert: () => 'YES', showModalDialog: (out, title) => dialogs.push({ title, html: out.html }), ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES' } });
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.ScriptApp = { getScriptId: () => 'SCRIPT', getProjectTriggers: () => [] };
  s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(meeting) });
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => [{ id: 'S4aP260074', type: 'agenda', revisionOf: null, agendaItem: '1' }];
  let sheet = null;
  s.UrlFetchApp = { fetch: (url) => (/\/Agenda\/agenda\.csv$/.test(url) ? { getResponseCode: () => 200, getContentText: () => '' }
    : { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 0x03, 0x04, 0x14], setName() { return this; } }) }) };
  s.DriveApp = { getFileById: drive.file, createFile: () => { sheet = LIST_86172; return { setTrashed() {} }; } };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => sheet }) }] }) };
  return { s, dialogs, docProps: loaded.docProps, body };
}

/** The creator dialog with its real client script and the real server functions behind google.script.run. */
function openCreator(meeting) {
  const drive = fakeDrive();
  drive.add(TEMPLATE_ID, 'SA4 Report Template');
  const master = templateRuntime(TEMPLATE_ID, drive, meeting);
  master.s.showCreateReportDialog();
  const html = master.dialogs[0].html;
  const registry = {};
  const el = (id) => registry[id] || (registry[id] = { id, value: '', textContent: '', innerHTML: '', disabled: false, style: {} });
  const rpc = ['previewNewReportFromTemplate', 'previewNewReportFromResolved', 'createNewReportFromTemplate'];
  const run = () => { let ok = () => {}; const r = { withSuccessHandler: (fn) => { ok = fn; return r; }, withFailureHandler: () => r };
    rpc.forEach((name) => { r[name] = (...args) => ok(plain(master.s[name](...args.map((a) => plain(a))))); }); return r; };
  const globals = { document: { getElementById: el }, google: { script: { get run() { return run(); }, host: { close: () => {} } } }, console: { log: () => {} } };
  vm.createContext(globals);
  vm.runInContext(html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')), globals);
  const lookUp = (id) => { el('meetingId').value = id; globals.lookUp(); };
  const typeList = (text) => { el('mailingList').value = text; globals.mailingTouched = true; vm.runInContext('mailingTouched = true', globals); globals.mailingChanged(); };
  const create = () => { globals.createReport(); const id = Object.keys(drive.files).find((k) => /^COPY/.test(k)); return id; };
  /** The new report: a fresh project (empty properties), set up by its first use. */
  const openReport = (copyId) => { const rep = templateRuntime(copyId, drive, meeting); rep.s.finishReportSetupWith_(rep.s.liveTemplateDeps_()); return rep; };
  return { master, drive, html, el, globals, lookUp, typeList, create, openReport };
}
const collectorList = (rep) => rep.s.getCollectorConfig_().LIST_NAME;
function configureDialogList(rep) {
  rep.s.configureMeetingSettings();
  const html = rep.dialogs[rep.dialogs.length - 1].html;
  return (html.match(/id="mailingList" value="([^"]*)"/) || [])[1];
}

/** One discussion e-mail generated in a report created by the creator; returns its text. */
function exportFromReport(rep, copyId) {
  installExportStubs(rep.s);
  rep.docProps.setProperty('DISCUSSION_EMAIL_SENDER', 'reporter@example.com');
  rep.docProps.setProperty('REVISIONS_URL', 'https://www.3gpp.org/ftp/TSG_SA/WG4_CODEC/3GPP_SA4_AHOC_MTGs/SA4_Plenary/Inbox/Drafts/');
  const files = [];
  const exportTables = [exportTable('S4aP260091', 'Title', '5.6.1')];
  rep.s.DocumentApp.getActiveDocument = () => ({ getId: () => copyId, getBody: () => ({ getTables: () => exportTables }) });
  rep.s.DriveApp = { getFoldersByName: () => ({ hasNext: () => false }), createFolder: () => ({ createFile: (b) => { files.push(b); return { getUrl: () => 'u' }; } }) };
  const result = rep.s.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: { date: '2026-10-15', time: '15:00', tz: 'CEST' } }], null);
  return { ok: result.ok, eml: blobText(files.find((f) => /\.eml$/.test(f.getName()))) };
}
const LIST_6G = '3GPP_TSG_SA4_FS_6G_MED';

console.log('3. the family -> mailing list table (Code.js 2.17.2: 6G has its own list)');
{
  const listOf = (family, type) => loadCode({ documentProperties: { REPORT_SUFFIX: family, MEETING_TYPE: type } }).sandbox.getMeetingContext_().sources.mailingList;
  check('6G derives exactly 3GPP_TSG_SA4_FS_6G_MED, ad-hoc and main', [listOf('6G', 'adhoc'), listOf('6G', 'main')], [LIST_6G, LIST_6G]);
  check('Audio / Video / MBS / RTC are unchanged', ['Audio', 'Video', 'MBS', 'RTC'].map((f) => listOf(f, 'adhoc')),
    ['3GPP_TSG_SA_WG4_AUDIO', '3GPP_TSG_SA_WG4_VIDEO', '3GPP_TSG_SA_WG4_MBS', '3GPP_TSG_SA_WG4_RTC']);
  check('the main-meeting families Liaison / New still use the general SA4 list', ['Liaison', 'New'].map((f) => listOf(f, 'main')), ['3GPP_TSG_SA_WG4', '3GPP_TSG_SA_WG4']);
  const { sandbox: s } = loadCode();
  check('Configure Meeting offers the same default (one table, no second mapping)', plain(s.buildReportFamilyInfo_())['6G'].mailingList, LIST_6G);
  const table = CODE.slice(CODE.indexOf('const MAILING_LISTS = {'), CODE.indexOf('};', CODE.indexOf('const MAILING_LISTS = {')));
  check('the mapping is by family, with no meeting id in it', [/'6G': '3GPP_TSG_SA4_FS_6G_MED'/.test(table), /\d{5}/.test(table)], [true, false]);
  check('no code path is keyed on meeting 86178 (its only mention in code is the pre-existing default of a manual timing diagnostic)',
    (CODE_ONLY.match(/86178/g) || []).length, 1);
}

console.log('3. creator: the derived mailing list is shown and can be replaced');
{
  const c = openCreator(MEETING_86178);
  check('the field is in the dialog, hidden until a meeting is looked up', /id="mailingRow" style="display:none"[\s\S]*id="mailingList"/.test(c.html), true);
  c.lookUp('86178');
  check('86178: lookup shows the 6G family list, editable', [c.el('family').value, c.el('mailingList').value, c.el('mailingList').disabled, c.el('mailingRow').style.display],
    ['6G', LIST_6G, false, '']);
  check('the hint says where it comes from', c.el('mailingHint').textContent, 'Derived from the report family. Replace it if this meeting uses another list.');
  check('nothing to answer: Create is enabled', c.el('createBtn').disabled, false);

  const copyId = c.create();
  const payload = JSON.parse(c.drive.files[copyId].description.split('SA4-BOOTSTRAP-V1:')[1]);
  const rep = c.openReport(copyId);
  check('left unchanged: no override is carried or stored; the report derives the same list',
    [payload.config.mailingList, payload.config.mailingListMode, rep.docProps._store.MAILING_LIST, rep.s.getMeetingContext_().sources.mailingList],
    ['', 'derived', undefined, LIST_6G]);
  const cfg = rep.s.getCollectorConfig_();
  check('collector RSS reads that list', [cfg.LIST_NAME, cfg.RSS_URL_V2], [LIST_6G, 'https://list.etsi.org/scripts/wa.exe?RSS&L=' + LIST_6G + '&v=2.0&LIMIT=2000']);
  const a1 = rep.s.buildArchiveIndexUrlsByDaysBack_(cfg.LIST_NAME, 14, cfg);
  check('collector A1 archive reads that list', [a1.length > 0, a1.every((u) => new RegExp('&L=' + LIST_6G + '$').test(u))], [true, true]);
  check('Configure Meeting shows it as the derived default (no saved override)',
    [configureDialogList(rep), plain(rep.s.buildReportFamilyInfo_())['6G'].mailingList], ['', LIST_6G]);
  const x = exportFromReport(rep, copyId);
  check('generated e-mail: To and Reply-To are the 6G list',
    [x.ok, emlHeader(x.eml, 'To'), emlHeader(x.eml, 'Reply-To')], [true, '3gpp_tsg_sa4_fs_6g_med@list.etsi.org', LIST_6G + '@list.etsi.org']);

  // The same family default for any 6G meeting -- nothing is keyed on the meeting id.
  const other = plain(MEETING_86178); other[0].Id = 90001; other[0].Title = '3GPPSA4-e (AH2) on FS_6G_MED';
  const c2 = openCreator(other);
  c2.lookUp('90001');
  check('another 6G meeting (different id) gets the same derived list', [c2.el('family').value, c2.el('mailingList').value], ['6G', LIST_6G]);
}

console.log('3. creator: an explicit override still wins');
const OVERRIDE = '3GPP_TSG_SA4_SOME_OTHER_WI';
{
  const c = openCreator(MEETING_86178);
  c.lookUp('86178');
  c.typeList(OVERRIDE);
  check('the proposal accepts the override and explains it',
    [c.el('createBtn').disabled, c.el('mailingList').value, c.el('mailingHint').textContent],
    [false, OVERRIDE, 'Your list will be used instead of the family default ' + LIST_6G + '.']);
  const copyId = c.create();
  const payload = JSON.parse(c.drive.files[copyId].description.split('SA4-BOOTSTRAP-V1:')[1]);
  check('carried in the setup information as an override', [payload.config.mailingList, payload.config.mailingListMode], [OVERRIDE, 'override']);
  const rep = c.openReport(copyId);
  check('stored in the report\'s normal MAILING_LIST property -- no second key',
    [rep.docProps._store.MAILING_LIST, rep.docProps.getKeys().filter((k) => /MAILING/.test(k))], [OVERRIDE, ['MAILING_LIST']]);
  check('the collector reads that list', collectorList(rep), OVERRIDE);
  check('Configure Meeting shows the same value', configureDialogList(rep), OVERRIDE);
  check('the e-mail collection start date decision is unaffected', rep.docProps._store.EMAIL_START_DATE, '2026-10-01');
  const x = exportFromReport(rep, copyId);
  check('the creator\'s override flows through to Reply-To', [x.ok, emlHeader(x.eml, 'Reply-To'), emlHeader(x.eml, 'From')],
    [true, OVERRIDE + '@list.etsi.org', 'reporter@example.com']);
}

console.log('3. creator: accepted forms, and what is refused');
{
  const c = openCreator(MEETING_86178);
  c.lookUp('86178');
  c.typeList('3gpp_tsg_sa_wg4@LIST.ETSI.ORG');
  check('the reflector-address form is accepted', c.el('createBtn').disabled, false);
  const rep = c.openReport(c.create());
  check('... and stored as the plain list name, exactly as the collector normalizes it',
    [rep.docProps._store.MAILING_LIST, rep.s.normalizeEtsiListName_('3gpp_tsg_sa_wg4@LIST.ETSI.ORG'), collectorList(rep)],
    ['3gpp_tsg_sa_wg4', '3gpp_tsg_sa_wg4', '3gpp_tsg_sa_wg4']);
}
{
  const c = openCreator(MEETING_86178);
  c.lookUp('86178');
  ['not a list!', 'someone@example.org', 'https://list.etsi.org/x', 'A&L=OTHER'].forEach((bad) => {
    c.typeList(bad);
    check(`invalid "${bad}": explained, Create disabled`, [c.el('createBtn').disabled, /is not a valid ETSI list name/.test(c.el('status').innerHTML)], [true, true]);
  });
  const resolved = plain(c.master.s.previewNewReportFromTemplate('86178', {})).resolved;
  const refused = plain(c.master.s.createNewReportFromTemplate('86178', resolved, { mailingList: 'not a list!' }));
  check('the server refuses too, before anything is copied', [refused.ok, c.drive.calls], [false, []]);

  c.typeList('3gpp_tsg_sa4_fs_6g_med');
  check('typing the derived list again (any case) is no override', [c.el('createBtn').disabled, c.el('mailingHint').textContent],
    [false, 'Derived from the report family. Replace it if this meeting uses another list.']);
  c.typeList('');
  check('an emptied field goes back to the derived list', c.el('mailingList').value, LIST_6G);
  c.typeList(OVERRIDE);
  c.el('family').value = 'Audio';
  c.globals.familyChanged();
  check('a typed list survives a family change', c.el('mailingList').value, OVERRIDE);
}
{
  const c = openCreator(MAIN_137);
  c.lookUp('86172');
  c.el('family').value = 'Audio';
  c.globals.familyChanged();
  check('main meeting: the family list is shown and cannot be changed in the creator',
    [c.el('mailingList').value, c.el('mailingList').disabled, c.el('mailingHint').textContent],
    ['3GPP_TSG_SA_WG4_AUDIO', true, 'A main-meeting report always reads the list of its report family.']);
  const resolved = plain(c.master.s.previewNewReportFromTemplate('86172', { reportFamily: 'Audio' })).resolved;
  const forced = plain(c.master.s.previewNewReportFromResolved('86172', resolved, { reportFamily: 'Audio', mailingList: 'SOME_OTHER_LIST' }));
  check('a main-meeting override is refused, not silently ignored', [forced.ok, /cannot be overridden/.test(forced.errors.join(' '))], [false, true]);
}
{
  check('no meeting-specific code: neither 86178 nor FS_6G_MED in the creator', /86178|FS_6G_MED/.test(CREATOR), false);
  const first = templateRuntime('REPORTdoc000000000000000000000000000000000', fakeDrive(), MEETING_86178);
  check('a tampered override in the setup information is refused by the new report',
    first.s.validateBootstrapPayload_({ schema: 'sa4-report-bootstrap/1', targetDocumentId: 'D',
      config: { meetingId: '1', meetingType: 'adhoc', reportType: 'MBS', mailingList: 'x@example.org', mailingListMode: 'override' } }, 'D'),
    ['The mailing list is not a valid ETSI list name.']);
}

// =====================================================================
// 4. REPLY-TO  and  5. SUBJECT
// =====================================================================

const SENDER = 'reporter@example.com';
const DEADLINE = { date: '2026-10-15', time: '15:00', tz: 'CEST' };
const REVISIONS = 'https://www.3gpp.org/ftp/x/Inbox/Drafts/';

function exportTable(id, title, agendaItem) {
  const cell = (text) => {
    const t = { getText: () => String(text), getTextAttributeIndices: () => [0], isBold: () => false, isItalic: () => false, isUnderline: () => false, getLinkUrl: () => null };
    const p = { getType: () => 'PARAGRAPH', asParagraph: () => p, editAsText: () => t, getText: () => String(text) };
    return { getText: () => String(text), getNumChildren: () => 1, getChild: () => p };
  };
  const rows = [['TDoc', id], ['Title', title], ['Source', 'Qualcomm'], ['Agenda Item', agendaItem], ['Type/For', 'CR / Agreement'],
    ['E-mail Discussion', 'No e-mail discussion.'], ['Revisions', ''], ['Minutes', ''], ['Disposition', ''], ['Status', 'Revised']]
    .map((r) => { const cells = r.map(cell); return { getNumCells: () => cells.length, getCell: (i) => cells[i] }; });
  return { getType: () => 'TABLE', getNumRows: () => rows.length, getRow: (i) => rows[i], getCell: (r, c) => rows[r].getCell(c) };
}
function emlHeader(eml, name) { const m = new RegExp('^' + name + ': (.*)$', 'm').exec(eml.split('\r\n\r\n')[0]); return m ? m[1] : null; }

/** Generates one discussion e-mail from a report with `props`; returns the result and the files written. */
function exportOne(props, id, title, agendaItem) {
  const loaded = loadCode({ documentProperties: Object.assign({ DISCUSSION_EMAIL_SENDER: SENDER, REVISIONS_URL: REVISIONS }, props) });
  const s = installExportStubs(loaded.sandbox);
  const tables = [exportTable(id, title, agendaItem)];
  s.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ getTables: () => tables }), getId: () => 'FAKE_DOC' });
  const files = [];
  s.DriveApp = { getFoldersByName: () => ({ hasNext: () => false }), createFolder: () => ({ createFile: (b) => { files.push(b); return { getUrl: () => 'u' }; } }) };
  let result, error = null;
  try { result = s.generateTdocDiscussionEmails([{ tableIndex: 0, deadline: DEADLINE }], null); } catch (e) { error = e.message; }
  const emlBlob = files.find((b) => /\.eml$/.test(b.getName()));
  return { s, result, error, files, eml: emlBlob ? blobText(emlBlob) : null, name: emlBlob ? emlBlob.getName() : null };
}

console.log('4. Reply-To is the effective mailing list; From is unchanged');
[
  ['MBS ad-hoc (family list)', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'MBS' }, 'S4aI260082', '2.5', '3GPP_TSG_SA_WG4_MBS@list.etsi.org', '3gpp_tsg_sa_wg4_mbs@list.etsi.org'],
  ['Audio ad-hoc (family list)', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Audio' }, 'S4aA260090', '3.1', '3GPP_TSG_SA_WG4_AUDIO@list.etsi.org', '3gpp_tsg_sa_wg4_audio@list.etsi.org'],
  ['Video ad-hoc (family list)', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Video' }, 'S4aV260012', '3.1', '3GPP_TSG_SA_WG4_VIDEO@list.etsi.org', '3gpp_tsg_sa_wg4_video@list.etsi.org'],
  ['RTC ad-hoc (family list)', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'RTC' }, 'A4aR260005', '3.1', '3GPP_TSG_SA_WG4_RTC@list.etsi.org', '3gpp_tsg_sa_wg4_rtc@list.etsi.org'],
  ['6G ad-hoc (family list, 2.17.2: the 6G list)', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G' }, 'S4aP260091', '5.6.1', '3GPP_TSG_SA4_FS_6G_MED@list.etsi.org', '3gpp_tsg_sa4_fs_6g_med@list.etsi.org'],
  ['6G main meeting (always the family list)', { MEETING_TYPE: 'main', REPORT_SUFFIX: '6G', MAILING_LIST: 'IGNORED_FOR_MAIN' }, 'S4-260500', '9.1', '3GPP_TSG_SA4_FS_6G_MED@list.etsi.org', '3gpp_tsg_sa4_fs_6g_med@list.etsi.org'],
  ['6G ad-hoc, another Mailing List as a plain name (override wins)', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', MAILING_LIST: '3GPP_TSG_SA_WG4' }, 'S4aP260091', '5.6.1', '3GPP_TSG_SA_WG4@list.etsi.org', '3gpp_tsg_sa_wg4@list.etsi.org'],
  ['6G ad-hoc, another Mailing List as an address (override wins)', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', MAILING_LIST: '3GPP_TSG_SA_WG4@list.etsi.org' }, 'S4aP260091', '5.6.1', '3GPP_TSG_SA_WG4@list.etsi.org', '3gpp_tsg_sa_wg4@list.etsi.org'],
  ['main meeting, Audio (always the family list)', { MEETING_TYPE: 'main', REPORT_SUFFIX: 'Audio', MAILING_LIST: 'IGNORED_FOR_MAIN' }, 'S4-261234', '7.3', '3GPP_TSG_SA_WG4_AUDIO@list.etsi.org', '3gpp_tsg_sa_wg4_audio@list.etsi.org']
].forEach(([name, props, id, agenda, replyTo, to]) => {
  const x = exportOne(props, id, 'Title', agenda);
  check(`${name}: From, Reply-To, To`, [x.result && x.result.ok, emlHeader(x.eml, 'From'), emlHeader(x.eml, 'Reply-To'), emlHeader(x.eml, 'To')], [true, SENDER, replyTo, to]);
});
{
  const x = exportOne({ MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'MBS' }, 'S4aI260082', 'Title', '2.5');
  check('Reply-To is never the sender', emlHeader(x.eml, 'Reply-To') === SENDER, false);
  check('header order, CRLF, quoted-printable and the ZIP are as before, with Reply-To after To',
    [x.eml.split('\r\n\r\n')[0].split('\r\n').map((l) => l.split(':')[0]), /\r\n\r\n/.test(x.eml), /[^\r]\n/.test(x.eml), emlHeader(x.eml, 'Content-Transfer-Encoding'), x.files.filter((b) => /\.zip$/.test(b.getName())).length],
    [['MIME-Version', 'X-Unsent', 'From', 'To', 'Reply-To', 'Subject', 'Content-Type', 'Content-Transfer-Encoding'], true, false, 'quoted-printable', 1]);

  const invalid = exportOne({ MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'MBS', MAILING_LIST: 'a.b@list.etsi.org' }, 'S4aI260082', 'Title', '2.5');
  check('a list the collector\'s rule rejects refuses the export: no file, no fallback to the sender',
    [/Could not derive the Reply-To address from the Mailing List "a\.b@list\.etsi\.org"/.test(invalid.error || (invalid.result && invalid.result.error) || ''), invalid.files], [true, []]);
  const s = x.s;
  check('Reply-To derivation follows normalizeEtsiListName_()',
    ['3GPP_TSG_SA4_FS_6G_MED', '3gpp_tsg_sa4_fs_6g_med@list.etsi.org', 'bad list', 'x@example.org', 'A\r\nBcc: x', ''].map((v) => s.deriveEmailExportReplyToFromMailingList_(v)),
    ['3GPP_TSG_SA4_FS_6G_MED@list.etsi.org', '3gpp_tsg_sa4_fs_6g_med@list.etsi.org', '', '', '', '']);
  check('the pure builder writes Reply-To only when it is given, and never invents it from From',
    [/^Reply-To:/m.test(s.buildEmlContent_({ from: SENDER, to: 'l@list.etsi.org', subject: 'x' }, '<p>b</p>')),
      /^Reply-To: L@list\.etsi\.org$/m.test(s.buildEmlContent_({ from: SENDER, to: 'l@list.etsi.org', replyTo: 'L@list.etsi.org', subject: 'x' }, '<p>b</p>'))], [false, true]);
  const replyFn = CODE.slice(CODE.indexOf('function deriveEmailExportReplyToFromMailingList_'), CODE.indexOf('function deriveEmailExportReplyToFromMailingList_') + 300);
  check('no meeting-specific code in the Reply-To derivation', /FS_6G_MED|86178|6G/.test(replyFn), false);
}

console.log('5. subject: [<agenda item>][<deadline>][<TDoc>] Discussion: <title>, no family or list label');
[
  ['MBS', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'MBS' }, 'S4aI260082', '2.5', 'MBS'],
  ['6G', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: '6G', MAILING_LIST: '3GPP_TSG_SA4_FS_6G_MED' }, 'S4aP260091', '5.6.1', 'FS_6G_MED'],
  ['Audio', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Audio' }, 'S4aA260090', '3.1', 'AUDIO'],
  ['Video', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Video' }, 'S4aV260012', '3.1', 'VIDEO'],
  ['RTC', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'RTC' }, 'A4aR260005', '3.1', 'RTC'],
  ['main SA4 (Audio)', { MEETING_TYPE: 'main', REPORT_SUFFIX: 'Audio' }, 'S4-261234', '7.3', 'AUDIO'],
  ['a long work-item list name', { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'MBS', MAILING_LIST: '3GPP_TSG_SA4_AMD_ARCH_PH2_MED_VERY_LONG_WORK_ITEM' }, 'S4aI260082', '2.5', 'AMD_ARCH_PH2_MED_VERY_LONG_WORK_ITEM']
].forEach(([name, props, id, agenda, tag]) => {
  const x = exportOne(props, id, 'A plain title', agenda);
  const subject = emlHeader(x.eml, 'Subject');
  check(`${name}: exact subject`, subject, `[${agenda}][26-10-15-1500CEST][${id}] Discussion: A plain title`);
  check(`${name}: neither the tag "${tag}" nor the family appears in the subject (it still names the file)`,
    [subject.indexOf(tag), new RegExp(props.REPORT_SUFFIX + '\\b', 'i').test(subject.replace(id, '')), x.name], [-1, false, tag + '_' + id + '.eml']);
});
{
  const { sandbox: s } = loadCode();
  const build = (...a) => s.buildEmailExportSubject_(...a);
  check('agenda item 2.5 appears as [2.5]', build('S4aI260082', 'T', '2.5', '26-10-15-1500CEST').slice(0, 5), '[2.5]');
  check('agenda item 5.10 is kept textually', build('S4aP260068', 'T', '5.10', '26-10-15-1500CEST'), '[5.10][26-10-15-1500CEST][S4aP260068] Discussion: T');
  check('the deadline token and the body deadline are formatted exactly as before',
    [s.formatEmailExportDeadlineForSubjectToken_({ date: '2026-10-15', time: '15:00', tz: 'CEST' }), s.formatEmailExportDeadline_({ date: '2026-10-15', time: '15:00', tz: 'CEST' })],
    ['26-10-15-1500CEST', '26-10-15 15:00 CEST']);
  check('TDoc bracket and "Discussion: <title>" are exactly as before', build('S4aI260082', '26501-CR0124-B "Slicing"', '2.5', '26-10-15-1500CEST').slice('[2.5][26-10-15-1500CEST]'.length),
    '[S4aI260082] Discussion: 26501-CR0124-B "Slicing"');
  check('missing parts are omitted, never invented',
    [build('S4aI260082', 'T', '', '26-10-15-1500CEST'), build('S4aI260082', 'T', '2.5', ''), build('S4aI260082', 'T', '', ''), build('S4aI260082', '', '2.5', '26-10-15-1500CEST')],
    ['[26-10-15-1500CEST][S4aI260082] Discussion: T', '[2.5][S4aI260082] Discussion: T', '[S4aI260082] Discussion: T', '[2.5][26-10-15-1500CEST][S4aI260082] Discussion']);
  check('a title that itself starts with a bracketed work item is left alone', build('S4aP260069', '[FS_6G_MED] pCR', '5.4', '26-10-15-1500CEST'),
    '[5.4][26-10-15-1500CEST][S4aP260069] Discussion: [FS_6G_MED] pCR');
  check('the collector associates replies to the new form and to e-mails already sent in the old form',
    ['Re: [2.5][26-10-15-1500CEST][S4aI260082] Discussion: T', 'RE: [MBS,2.5,26-10-15-1500CEST][S4aI260082] Discussion: T', 'AW: [FS_6G_MED,5.6.1,26-10-15-1500CEST][S4aP260091] Discussion: T']
      .map((subj) => s.findSA4DocumentIdsInText_(s.stripReplyPrefixes_(subj))), [['S4aI260082'], ['S4aI260082'], ['S4aP260091']]);
  check('the agenda-item bracket is not mistaken for a document', s.findSA4DocumentIdsInText_('[2.5][26-10-15-1500CEST] no document here'), []);
  const fn = CODE.slice(CODE.indexOf('function buildEmailExportSubject_('), CODE.indexOf('function buildEmailExportSubject_(') + 500);
  check('the subject builder takes no tag and names no family or meeting', [/listTag|FS_6G_MED|MBS/.test(fn), /function buildEmailExportSubject_\(tdocNumber, title, agendaItem, deadlineText\) \{/.test(fn)], [false, true]);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll checks passed.');
process.exitCode = failures ? 1 : 0;
