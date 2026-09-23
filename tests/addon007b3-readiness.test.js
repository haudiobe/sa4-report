/**
 * ADDON-007B3 -- live configuration readiness + safe build boundary.
 *
 * One canonical declarative rule set (MEETING_READINESS_RULES_) evaluated by
 * one evaluator (evaluateMeetingReadiness_): shown live in the dialog (the
 * evaluator's own source is shipped to the browser), and enforced at the
 * ad-hoc build entry points, including the central background path.
 *
 * Run: node tests/addon007b3-readiness.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

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

const FIXTURES = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'meeting-resolver-samples.json'), 'utf8'));
const { sandbox: pure } = loadCode();
const RULES = JSON.parse(vm.runInContext('JSON.stringify(MEETING_READINESS_RULES_)', pure));
const evaluate = (fields, options) => JSON.parse(JSON.stringify(pure.evaluateMeetingReadiness_(fields, RULES, options)));
const codes = (list) => list.map((i) => i.code);

const COMPLETE_ADHOC = {
  meetingType: 'adhoc', meetingId: '12345', meetingName: 'SA4 ad hoc', ftpBase: 'https://example.invalid/Docs/', reportFamily: 'Audio',
  agendaTdoc: 'S4aA000001', tdocListUrl: 'https://example.invalid/list.xlsx', mailingList: 'SOME_LIST'
};
const without = (fields, ...keys) => { const c = Object.assign({}, fields); keys.forEach((k) => { c[k] = ''; }); return c; };

// ------------------------------------------------------------ evaluator

console.log('evaluateMeetingReadiness_ -- ad-hoc');

{
  const r = evaluate(COMPLETE_ADHOC);
  check('complete ad-hoc configuration -> ready, no issues, no warnings', [r.ready, r.issues, r.warnings], [true, [], []]);

  const unresolved = evaluate({});
  check('unresolved meeting (no type, nothing configured) -> blocked with the resolve message',
    [unresolved.ready, unresolved.issues.map((i) => i.message)], [false, ['Resolve the 3GPP meeting first.', 'Select a report family.']]);
  check('ad-hoc without a meeting id -> blocked', codes(evaluate(without(COMPLETE_ADHOC, 'meetingId')).issues), ['MEETING_NOT_RESOLVED']);
  check('missing family -> blocked', codes(evaluate(without(COMPLETE_ADHOC, 'reportFamily')).issues), ['REPORT_FAMILY_REQUIRED']);
  check('missing agenda -> blocked', codes(evaluate(without(COMPLETE_ADHOC, 'agendaTdoc')).issues), ['AGENDA_TDOC_REQUIRED']);
  check('missing TDoc list -> blocked', codes(evaluate(without(COMPLETE_ADHOC, 'tdocListUrl')).issues), ['TDOC_LIST_REQUIRED']);
  check('missing FTP base -> blocked', codes(evaluate(without(COMPLETE_ADHOC, 'ftpBase')).issues), ['FTP_BASE_REQUIRED']);
  check('missing meeting name -> blocked', codes(evaluate(without(COMPLETE_ADHOC, 'meetingName')).issues), ['MEETING_NAME_REQUIRED']);
  check('whitespace-only counts as missing', codes(evaluate(Object.assign({}, COMPLETE_ADHOC, { tdocListUrl: '   ' })).issues), ['TDOC_LIST_REQUIRED']);

  const mailing = evaluate(without(COMPLETE_ADHOC, 'mailingList'));
  check('missing mailing list only -> ready with a warning', [mailing.ready, codes(mailing.warnings)], [true, ['MAILING_LIST_MISSING']]);
  check('a revisions URL is not part of the rules at all -> still ready', evaluate(Object.assign({}, COMPLETE_ADHOC, { revisionsUrl: '' })).ready, true);
  check('no rule mentions the Reviewer token / email preview / revisions', RULES.some((r) => /token|preview|revisions/i.test(JSON.stringify(r))), false);

  const all = evaluate({ meetingType: 'adhoc' });
  check('several missing items are all listed, in a stable order', codes(all.issues),
    ['MEETING_NOT_RESOLVED', 'REPORT_FAMILY_REQUIRED', 'MEETING_NAME_REQUIRED', 'AGENDA_TDOC_REQUIRED', 'TDOC_LIST_REQUIRED', 'FTP_BASE_REQUIRED']);
}

console.log('evaluateMeetingReadiness_ -- main is evaluated separately');

{
  const MAIN = { meetingType: 'main', reportFamily: 'Audio', meetingFolder: 'TSGS4_137_Xian', meetingNumber: '137', mailingList: 'L' };
  check('complete main -> ready without any ad-hoc-only field', evaluate(MAIN).ready, true);
  check('main does not need agenda TDoc, TDoc list, FTP base or meeting name', codes(evaluate(MAIN).issues), []);
  check('main missing family -> blocked', codes(evaluate(without(MAIN, 'reportFamily')).issues), ['REPORT_FAMILY_REQUIRED']);
  check('main with neither folder nor resolved FTP base -> blocked', codes(evaluate(without(MAIN, 'meetingFolder')).issues), ['MAIN_MEETING_FOLDER_REQUIRED']);
  check('a resolved FTP base stands in for the folder', evaluate(Object.assign(without(MAIN, 'meetingFolder'), { ftpBase: 'https://x/Docs/' })).ready, true);
  check('an explicit TDoc list stands in for the number', evaluate(Object.assign(without(MAIN, 'meetingNumber'), { tdocListUrl: 'https://x/l.xlsx' })).ready, true);
  check('a legacy document with no type but a folder is evaluated as main', evaluate({ reportFamily: '6G', meetingFolder: 'F', meetingNumber: '1' }).meetingType, 'main');
}

console.log('evaluateMeetingReadiness_ -- build scope (forBuild)');

{
  check('ad-hoc build: the sources with no runtime default are enforced', codes(evaluate({ meetingType: 'adhoc' }, { forBuild: true }).issues),
    ['REPORT_FAMILY_REQUIRED', 'AGENDA_TDOC_REQUIRED', 'TDOC_LIST_REQUIRED', 'FTP_BASE_REQUIRED']);
  check('ad-hoc build: meeting id/name and the mailing list never block a build',
    evaluate(without(COMPLETE_ADHOC, 'meetingId', 'meetingName', 'mailingList'), { forBuild: true }).ready, true);
  check('main build: nothing is enforced (every legacy fallback keeps working)', evaluate({ meetingType: 'main' }, { forBuild: true }).ready, true);
  check('main build with an empty configuration is still not blocked', evaluate({ meetingType: 'main', reportFamily: '', meetingFolder: '' }, { forBuild: true }).issues, []);
}

console.log('user-facing wording');

{
  const text = JSON.stringify(RULES.map((r) => r.message));
  check('no internal property / function / task names in any message',
    /REPORT_SUFFIX|TDOC_LIST_URL|FTP_BASE|AGENDA_TDOC|Document Propert|ADDON|PROD|ARCH|resolver|_/.test(text), false);
  const messages = Object.fromEntries(RULES.map((r) => [r.code, r.message]));
  check('required messages', [messages.REPORT_FAMILY_REQUIRED, messages.AGENDA_TDOC_REQUIRED, messages.TDOC_LIST_REQUIRED, messages.MEETING_NOT_RESOLVED, messages.MAILING_LIST_MISSING],
    ['Select a report family.', 'Discover or select the agenda document.', "Paste the meeting's TDoc list URL.", 'Resolve the 3GPP meeting first.',
      'No mailing list is configured. Email collection will be unavailable.']);
}

// ------------------------------------------------------------ live dialog

const IDS = ['meetingId', 'agendaStructure', 'tdocUrlHint', 'mailingListHint', 'mailingListReset', 'familyStatus', 'mainMeetingFields', 'meetingType', 'reportType',
  'familyInfo', 'readinessRules', 'readinessStatus', 'meetingSummary', 'resolveStatus', 'resolveBtn', 'discoverStatus', 'discoverBtn', 'portalTypeHint',
  'dateRangeHint', 'agendaCandidates', 'meetingName', 'meetingDate', 'ftpBase', 'ftpBaseBadge', 'agendaTdoc', 'agendaTdocBadge', 'mailingList', 'revisionsUrl',
  'revisionsUrlBadge', 'apiToken', 'clearApiToken', 'meetingFolder', 'meetingNumber', 'agendaSourceDocId', 'tdocUrl', 'showPreview'];

function openDialog(documentProperties, meetingFixture, tdocRows) {
  const loaded = loadCode({ documentProperties: documentProperties || {} });
  const sandbox = loaded.sandbox;
  let html = null;
  sandbox.HtmlService = { createHtmlOutput: (h) => { html = h; const o = { setWidth: () => o, setHeight: () => o }; return o; } };
  sandbox.DocumentApp.getUi = () => ({ showModalDialog: () => {} });
  sandbox.configureMeetingSettings();

  sandbox.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(meetingFixture) });
  sandbox.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  sandbox.parseMeetingTdocListHtml_ = () => tdocRows || [];

  const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')).replace(/\\\\/g, '\\');
  const registry = {};
  IDS.forEach((id) => { registry[id] = { id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' }; });
  const value = (id) => { const m = html.match(new RegExp('id="' + id + '" value="([^"]*)"')); return m ? m[1].replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&') : ''; };
  ['familyInfo', 'readinessRules', 'meetingId', 'meetingType', 'meetingName', 'mailingList', 'ftpBase', 'agendaTdoc', 'tdocUrl', 'meetingFolder', 'meetingNumber']
    .forEach((id) => { registry[id].value = value(id); });
  const selected = html.match(/<option value="([A-Za-z0-9]+)" selected>/);
  registry.reportType.value = selected ? selected[1] : '';

  let sent = null;
  const alerts = [];
  const runner = () => {
    let ok = null;
    const r = {
      withSuccessHandler: (fn) => { ok = fn; return r; },
      withFailureHandler: () => r,
      resolveMeetingForConfigDialog: (id) => ok(JSON.parse(JSON.stringify(sandbox.resolveMeetingForConfigDialog(id)))),
      discoverAgendaForConfigDialog: (id, core) => ok(JSON.parse(JSON.stringify(sandbox.discoverAgendaForConfigDialog(id, core)))),
      saveConfigurationSettings: (config) => { sent = config; ok(); }
    };
    return r;
  };
  const globals = {
    document: { getElementById: (id) => registry[id] || null },
    google: { script: { get run() { return runner(); }, host: { close: () => {} } } },
    alert: (m) => alerts.push(m),
    console: { log: () => {} }
  };
  vm.createContext(globals);
  vm.runInContext(script, globals);
  return { registry, globals, sandbox, alerts, html, status: () => registry.readinessStatus.innerHTML, sentConfig: () => sent };
}
const pickFamily = (d, f) => { d.registry.reportType.value = f; d.globals.onFamilyChanged(); };
const typeInto = (d, id, v) => { d.registry[id].value = v; d.globals.updateDependentUi(); };

const M85916 = FIXTURES.getMeetings85916;
const AUDIO_ROWS = [{ id: 'S4aA260090', type: 'agenda', revisionOf: null, agendaItem: '' }, { id: 'S4aA260091', type: 'pCR', revisionOf: null, agendaItem: '4' }];

console.log('live dialog');

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  check('the status block sits before the Save button', d.html.indexOf('id="readinessStatus"') < d.html.indexOf('Save Configuration'), true);
  check('the shipped evaluator is the server function (single implementation)', d.html.indexOf(pure.evaluateMeetingReadiness_.toString()) !== -1, true);
  check('initial incomplete state shows Needs attention', /Needs attention/.test(d.status()) && d.registry.readinessStatus.className === 'attention', true);
  check('initial state lists what to do', /Resolve the 3GPP meeting first\./.test(d.status()) && /Select a report family\./.test(d.status()), true);

  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  check('Resolve updates readiness (identity/family/folder satisfied, agenda + TDoc list outstanding)',
    [/Resolve the 3GPP meeting first/.test(d.status()), /Select a report family/.test(d.status()), /agenda document/.test(d.status()), /TDoc list URL/.test(d.status())],
    [false, false, true, true]);

  d.globals.discoverAgendaTdocs();
  check('Discover updates readiness (agenda found)', [/agenda document/.test(d.status()), /TDoc list URL/.test(d.status())], [false, true]);

  typeInto(d, 'tdocUrl', 'https://example.invalid/TDoc_List.xlsx');
  check('editing the TDoc list URL flips to Ready to build without reopening', [/Ready to build/.test(d.status()), d.registry.readinessStatus.className], [true, 'ready']);
  check('ready wording', /All required meeting sources are configured\./.test(d.status()), true);
  check('no warning while the mailing list is derived from the family', /Warnings/.test(d.status()), false);

  pickFamily(d, '');
  check('changing the family back to blank updates readiness', /Select a report family/.test(d.status()) && /Needs attention/.test(d.status()), true);
  pickFamily(d, 'Audio');
  check('choosing it again restores ready', /Ready to build/.test(d.status()), true);

  typeInto(d, 'agendaTdoc', '');
  check('clearing the agenda TDoc updates readiness', /agenda document/.test(d.status()), true);
  typeInto(d, 'agendaTdoc', 'S4aA260090');
  typeInto(d, 'ftpBase', '');
  check('an Advanced edit (document folder) updates readiness', /document folder/.test(d.status()), true);
  typeInto(d, 'ftpBase', 'https://example.invalid/Docs/');
  check('and restoring it is Ready again', /Ready to build/.test(d.status()), true);

  d.registry.mailingList.value = '';
  d.globals.onMailingListEdited();
  check('a blank mailing list override is a warning, not a blocker',
    [/Ready to build/.test(d.status()), /Warnings:/.test(d.status()), /No mailing list is configured/.test(d.status())], [true, true, true]);
  d.globals.resetMailingList();
  check('resetting to the derived default clears the warning', /Warnings/.test(d.status()), false);

  d.globals.saveConfig();
  check('Save on a ready configuration only says saved', d.alerts[d.alerts.length - 1], '✅ Configuration saved.');
}

{
  const d = openDialog({}, M85916, AUDIO_ROWS);
  d.registry.meetingId.value = '85916';
  d.globals.resolveMeeting();
  d.globals.saveConfig();
  const msg = d.alerts[d.alerts.length - 1];
  check('Save of an incomplete configuration is allowed and does not imply readiness',
    [/Configuration saved/.test(msg), /not ready to build yet/.test(msg), /TDoc list URL/.test(msg)], [true, true, true]);
  check('the dialog still says Needs attention afterwards', /Needs attention/.test(d.status()), true);
  check('the save payload was sent (partial configuration is saveable)', d.sentConfig().reportType, 'Audio');
}

console.log('live dialog -- main');

{
  const d = openDialog({ MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137_Xian', MEETING_NUMBER: '137', REPORT_SUFFIX: 'Audio', MEETING_ID: '12345' });
  check('a saved main meeting is ready without any ad-hoc-only field', /Ready to build/.test(d.status()), true);
  pickFamily(d, '');
  check('main: family is still required', /Select a report family/.test(d.status()), true);
}

// ------------------------------------------------------------ build boundary

console.log('build boundary');

const ADHOC_COMPLETE_PROPS = { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4 ad hoc', REPORT_SUFFIX: 'Audio', AGENDA_TDOC: 'S4aA000001',
  FTP_BASE: 'https://example.invalid/Docs/', TDOC_LIST_URL: 'https://example.invalid/list.xlsx' };

function bodyRecorder(calls) {
  const body = new Proxy({}, { get: (_t, name) => (...args) => { calls.push(String(name)); return name === 'getTables' ? [] : body; } });
  return body;
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', MEETING_NAME: 'SA4 ad hoc', REPORT_SUFFIX: 'Audio' } });
  const calls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => bodyRecorder(calls), getId: () => 'D' });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('incomplete ad-hoc build fails', thrown !== null, true);
  check('and fails before the document is touched', calls, []);
  check('the error is actionable: says what is missing and where to fix it',
    [/Cannot build report yet/.test(thrown.message), /TDoc list URL/.test(thrown.message), /agenda document/.test(thrown.message), /Configure Meeting/.test(thrown.message)], [true, true, true, true]);
  check('no internal names leak into the error', /TDOC_LIST_URL|AGENDA_TDOC|REPORT_SUFFIX|FTP_BASE/.test(thrown.message), false);
}

{
  const { sandbox } = loadCode({ documentProperties: Object.assign({}, ADHOC_COMPLETE_PROPS, { TDOC_LIST_URL: '' }) });
  let thrown = null;
  try { sandbox.assertMeetingReadyToBuild_(); } catch (e) { thrown = e; }
  check('a missing TDoc list alone names exactly that', thrown && thrown.message.split('\n').filter((l) => l.startsWith('•')), ["• Paste the meeting's TDoc list URL."]);
}

{
  const { sandbox } = loadCode({ documentProperties: ADHOC_COMPLETE_PROPS });
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ clear: () => { throw new Error('REACHED_BODY_CLEAR'); } }), getId: () => 'D' });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('complete ad-hoc build proceeds past the guard (reaches the document)', thrown && thrown.message, 'REACHED_BODY_CLEAR');
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_137_Xian', MEETING_NUMBER: '137' } });
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ clear: () => { throw new Error('REACHED_BODY_CLEAR'); } }), getId: () => 'D' });
  let thrown = null;
  try { sandbox.buildSkeletonWithTdocTables(); } catch (e) { thrown = e; }
  check('main build is unchanged (no readiness guard, reaches the document)', thrown && thrown.message, 'REACHED_BODY_CLEAR');

  const legacy = loadCode();
  legacy.sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({ clear: () => { throw new Error('REACHED_BODY_CLEAR'); } }), getId: () => 'D' });
  let t2 = null;
  try { legacy.sandbox.buildSkeletonWithTdocTables(); } catch (e) { t2 = e; }
  check('a legacy document with no configuration at all still relies on its fallbacks', t2 && t2.message, 'REACHED_BODY_CLEAR');
}

// ------------------------------------------------------------ central / background

console.log('central add-on: registered document + background updates');

const DOC = 'DOC_B3_REGISTERED';
const BACKGROUND = { documentId: DOC, mode: 'addon-background' };

function registeredDocument(props) {
  const loaded = loadCode({ documentProperties: props });
  const { sandbox } = loaded;
  const calls = [];
  sandbox.DocumentApp.getActiveDocument = () => ({ getId: () => DOC, getBody: () => bodyRecorder(calls) });
  sandbox.DocumentApp.openById = () => ({ getBody: () => bodyRecorder(calls) });
  sandbox.adoptReportDocumentForAddon_({ documentId: DOC, mode: 'addon-interactive' });
  sandbox.registerReportDocument_(DOC, { enabled: true, intervalHours: 2 });
  return Object.assign(loaded, { calls });
}

{
  const { sandbox } = registeredDocument(ADHOC_COMPLETE_PROPS);
  check('registered document: readiness is read from the saved central state', sandbox.getBuildReadiness_(BACKGROUND).ready, true);
  sandbox.getReportStateStore_(BACKGROUND).deleteProperty('TDOC_LIST_URL');
  check('and follows the central copy, not the document property (which is still set)', codes(sandbox.getBuildReadiness_(BACKGROUND).issues), ['TDOC_LIST_REQUIRED']);
}

{
  const { sandbox, calls } = registeredDocument(Object.assign({}, ADHOC_COMPLETE_PROPS, { TDOC_LIST_URL: '' }));
  let thrown = null;
  try { sandbox.continuousUpdateForDocument_(DOC); } catch (e) { thrown = e; }
  check('incomplete ad-hoc background update fails safely', [thrown !== null, thrown && /Cannot build report yet/.test(thrown.message)], [true, true]);
  check('before any document mutation', calls.filter((c) => /clear|insert|append|remove|set|replace/i.test(c)), []);
  check('a failed background update does not advance lastRunAt', sandbox.getRegisteredReportDocument_(DOC).lastRunAt, null);
}

{
  const { sandbox } = registeredDocument(ADHOC_COMPLETE_PROPS);
  sandbox.downloadAndGroupTdocs_ = () => { throw new Error('REACHED_TDOC_DOWNLOAD'); };
  let thrown = null;
  try { sandbox.continuousUpdateForDocument_(DOC); } catch (e) { thrown = e; }
  check('a complete ad-hoc document gets past the guard in the background', thrown && /REACHED_TDOC_DOWNLOAD/.test(thrown.message), true);
  check('(and, that run failing, lastRunAt still did not advance)', sandbox.getRegisteredReportDocument_(DOC).lastRunAt, null);
}

{
  const { sandbox } = loadCode({ documentProperties: { MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'Audio' } });
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => bodyRecorder([]), getId: () => 'D' });
  const result = sandbox.continuousUpdateCore_();
  check('interactive update keeps its swallow-and-report semantics; the guard reports through the same result',
    [result.success, /Cannot build report yet/.test(result.error)], [false, true]);
}

// ------------------------------------------------------------ regression

console.log('regression');

{
  const src = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const start = src.indexOf('const MEETING_READINESS_RULES_');
  const block = src.slice(start, src.indexOf('function buildMeetingSummary_('));
  check('no meeting-specific constants in the readiness code', /85916|ULBC|86178/.test(block), false);
  check('the old parallel readiness function is gone (one rule set)', /function computeMeetingConfigReadiness_/.test(src), false);

  const { sandbox, docProps } = loadCode();
  sandbox.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '85916', reportType: 'Audio', mailingList: '', mailingListMode: 'derived', apiTokenAction: 'replace', apiToken: 'crv1_TESTONLY_x', showPreview: true });
  check('B2 derived mailing-list semantics intact (no explicit override persisted)', docProps._store.MAILING_LIST, undefined);
  check('B1 token handling intact', sandbox.PropertiesService.getScriptProperties()._store.REVIEWER_API_TOKEN, 'crv1_TESTONLY_x');
  check('B2 inference intact', JSON.parse(JSON.stringify(sandbox.inferReportFamily_({ meetingType: 'adhoc', meetingName: 'SA4 Audio SWG' }))).family, 'Audio');
}

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
