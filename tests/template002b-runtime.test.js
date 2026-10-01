/**
 * TEMPLATE-002B -- the SA4 Report Template runtime: Code.js + ReportCreator.js
 * + Release.js loaded together, as in the bound script of the master template
 * and of every report copied from it.
 *
 * Covers the menus (master / created report / add-on unchanged), the template
 * guard, the creator dialog end to end (lookup -> DriveApp.makeCopy ->
 * setup information -> first use of the copy), Email Collection Start Date,
 * and the 15/30/60-minute Continuous Update trigger.
 *
 * Google services are fakes that behave as TEMPLATE-001 found live: a copy
 * is a new project with empty property stores and no triggers, and it can
 * read the description the creator wrote.
 *
 * Run: node tests/template002b-runtime.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadTemplateRuntime, REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');
const { makeFakeDocumentBody } = require('./helpers/fake-document.js');
const { planRelease } = require('../tools/template-release.js');

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
const thrown = (fn) => { try { fn(); return null; } catch (e) { return e.message; } };

const FX = path.join(__dirname, 'fixtures');
const CSV_86172 = fs.readFileSync(path.join(FX, 'meeting-86172-agenda.csv'), 'utf8');
const LIST_86172 = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-tdoc-list-values.json'), 'utf8')).values;
const GM = JSON.parse(fs.readFileSync(path.join(FX, 'meeting-86172-85916-getmeetings.json'), 'utf8'));
const ROWS_86172 = [
  { id: 'S4aI260089', type: 'discussion', revisionOf: null, agendaItem: '2.7' },
  { id: 'S4aI260083', type: 'CR', revisionOf: null, agendaItem: '2.5' },
  { id: 'S4aI260081', type: 'draft TS', revisionOf: null, agendaItem: '3.7' }
];

const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const RELEASE = {
  releaseId: 'T-2026.10.0', flavor: 'template', codeVersion: '2.17.0', gitCommit: 'abcdef0000000000000000000000000000000000',
  gitTag: 'template-release/T-2026.10.0', templateDocumentId: TEMPLATE_ID
};

// ------------------------------------------------------------------ fakes

function fakeUi() {
  const ui = { menus: [], alerts: [], dialogs: [], alertResponse: 'OK', ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES', NO: 'NO', OK: 'OK' } };
  ui.createMenu = (name) => {
    const m = { name, entries: [] };
    m.addItem = (label, fn) => { m.entries.push({ label, fn }); return m; };
    m.addSeparator = () => { m.entries.push({ sep: true }); return m; };
    m.addSubMenu = (sub) => { m.entries.push({ sub }); return m; };
    m.addToUi = () => { ui.menus.push(m); };
    return m;
  };
  ui.alert = (...args) => { ui.alerts.push(args); return ui.alertResponse; };
  ui.showModalDialog = (out, title) => { ui.dialogs.push({ title, html: out.html }); };
  return ui;
}

/** One Drive for a whole scenario: files with a description, copies made by makeCopy(). */
function fakeDrive() {
  const drive = { files: {}, calls: [], seq: 0 };
  drive.add = (id, name) => { drive.files[id] = { id, name, description: '', trashed: false }; return drive.files[id]; };
  drive.file = (id) => {
    const f = drive.files[id];
    if (!f) throw new Error('No such file ' + id);
    return {
      getId: () => f.id, getName: () => f.name, getUrl: () => 'https://docs.google.com/document/d/' + f.id + '/edit',
      getDescription: () => f.description, setDescription: (t) => { drive.calls.push(['setDescription', f.id]); f.description = t; },
      setTrashed: (v) => { f.trashed = v; },
      getParents: () => ({ hasNext: () => true, next: () => 'FOLDER' }),
      makeCopy: (name, folder) => { drive.calls.push(['makeCopy', name, folder]); return drive.file(drive.add('COPYdoc' + String(++drive.seq).padStart(30, '0'), name).id); }
    };
  };
  return drive;
}

/**
 * A template-runtime sandbox bound to one document. opts: { docId, props,
 * release (false = no Release.js), drive, meeting, triggers }.
 */
function runtime(opts) {
  const o = opts || {};
  const loaded = loadTemplateRuntime({ release: o.release === false ? null : (o.release || RELEASE), documentProperties: o.props || {} });
  const s = loaded.sandbox;
  const ui = fakeUi();
  const drive = o.drive || fakeDrive();
  const docId = o.docId || REPORT_ID;
  if (!drive.files[docId]) drive.add(docId, 'doc');
  const body = makeFakeDocumentBody(s);
  s.DocumentApp.getActiveDocument = () => ({ getId: () => docId, getBody: () => body });
  s.DocumentApp.getUi = () => ui;
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html, setWidth: () => out, setHeight: () => out }; return out; } };
  // Network: one captured meeting (the real discovery pipeline runs on it).
  let sheetValues = null;
  s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(o.meeting || GM.getMeetings86172) });
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => ROWS_86172;
  s.UrlFetchApp = { fetch: (url) => (/\/Agenda\/agenda\.csv$/.test(url)
    ? { getResponseCode: () => 200, getContentText: () => CSV_86172 }
    : { getResponseCode: () => 200, getBlob: () => ({ getBytes: () => [0x50, 0x4B, 0x03, 0x04, 0x14], setName() { return this; } }) }) };
  s.DriveApp = { getFileById: drive.file, createFile: () => { sheetValues = LIST_86172; return { setTrashed() {} }; } };
  s.SpreadsheetApp = { open: () => ({ getSheets: () => [{ getDataRange: () => ({ getValues: () => sheetValues }) }] }) };
  const triggers = boundScriptApp(o.triggers || []);
  s.ScriptApp = triggers.api;
  return { s, ui, drive, docId, docProps: loaded.docProps, scriptProps: loaded.scriptProps, triggers };
}

/** ScriptApp of a bound script: everyMinutes(1|5|10|15|30) and everyHours(n) are allowed. */
function boundScriptApp(existingHandlers) {
  const state = { triggers: existingHandlers.map((h, i) => ({ h, id: 'old' + i })), calls: [], failCreate: false, seq: 0 };
  const builder = (h, kind, n) => ({
    create: () => {
      state.calls.push(['create', h, kind + '(' + n + ')']);
      if (state.failCreate) throw new Error('Service error: quota');
      state.triggers.push({ h, id: 'new' + (++state.seq), kind, n });
    }
  });
  state.api = {
    TriggerSource: { CLOCK: 'CLOCK' },
    getScriptId: () => 'SCRIPTid00000000000000000000000000000000000000000000',
    getProjectTriggers: () => state.triggers.map((t) => ({ getHandlerFunction: () => t.h, getTriggerSource: () => 'CLOCK', getUniqueId: () => t.id, _t: t })),
    deleteTrigger: (t) => { state.calls.push(['delete', t._t.h, t._t.id]); state.triggers = state.triggers.filter((x) => x !== t._t); },
    newTrigger: (h) => ({ timeBased: () => ({
      everyMinutes: (n) => {
        if ([1, 5, 10, 15, 30].indexOf(n) === -1) throw new Error('The value you passed to everyMinutes was invalid.');
        return builder(h, 'everyMinutes', n);
      },
      everyHours: (n) => builder(h, 'everyHours', n)
    }) })
  };
  state.handlers = () => state.triggers.map((t) => t.h);
  return state;
}

/** [{ path, fn }] for every item of the installed menu, submenus flattened. */
function menuItems(ui) {
  const out = [];
  const walk = (m, prefix) => m.entries.forEach((e) => {
    if (e.sub) walk(e.sub, prefix + e.sub.name + ' > ');
    else if (!e.sep) out.push({ path: prefix + e.label, fn: e.fn });
  });
  if (ui.menus.length) walk(ui.menus[ui.menus.length - 1], '');
  return out;
}
const submenus = (ui) => ui.menus[ui.menus.length - 1].entries.filter((e) => e.sub).map((e) => e.sub.name);

/** Runs the first <script> of a dialog against a small fake DOM and google.script.run. */
function runDialogScript(html, rpc, seedIds) {
  const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')).replace(/\\\\/g, '\\');
  const registry = {};
  const element = (id) => ({ id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' });
  const attr = (id, name) => { const m = html.match(new RegExp('id="' + id + '"[^>]*\\s' + name + '="([^"]*)"')); return m ? m[1].replace(/&quot;/g, '"') : null; };
  (seedIds || []).forEach((id) => { registry[id] = element(id); const v = attr(id, 'value'); if (v !== null) registry[id].value = v; });
  if (/id="collectionStartField"/.test(html)) {
    const f = element('collectionStartField');
    f.value = attr('collectionStartField', 'value') || '';
    f.touched = false;
    f.getAttribute = (n) => attr('collectionStartField', n);
    registry.collectionStartField = f;
  }
  const selected = html.match(/<option value="([A-Za-z0-9]+)" selected>/);
  if (registry.reportType && selected) registry.reportType.value = selected[1];
  const alerts = [];
  let sent = null;
  const run = () => {
    let ok = () => {};
    const r = { withSuccessHandler: (fn) => { ok = fn; return r; }, withFailureHandler: () => r };
    Object.keys(rpc).forEach((name) => { r[name] = (...args) => { if (name === 'saveConfigurationSettings') sent = args[0]; ok(plain(rpc[name](...args.map((a) => (a === undefined ? a : plain(a)))) === undefined ? null : rpc[name].last)); }; });
    return r;
  };
  const globals = {
    document: { getElementById: (id) => registry[id] || (registry[id] = element(id)) },
    google: { script: { get run() { return run(); }, host: { close: () => {} } } },
    alert: (m) => alerts.push(m), console: { log: () => {} }
  };
  vm.createContext(globals);
  vm.runInContext(script, globals);
  return { globals, registry, alerts, sent: () => sent };
}
/** Wraps a server function so the fake google.script.run can hand its result to the success handler. */
function serverCall(fn) {
  const wrapped = (...args) => { wrapped.last = plain(fn(...args)); return wrapped.last === undefined ? undefined : true; };
  wrapped.last = null;
  return wrapped;
}

// ================================================================ 1. menus

console.log('1. master-template menu: creation and release information only');
{
  const r = runtime({ docId: TEMPLATE_ID });
  r.s.onOpen();
  check('one menu, two items', menuItems(r.ui).map((i) => i.fn), ['showCreateReportDialog', 'showTemplateInfo']);
  check('labels', menuItems(r.ui).map((i) => i.path), ['🆕 Create New SA4 Report', 'ℹ️ Template Release Info']);
  check('no report operation, no submenu at all', submenus(r.ui), []);
  check('every target exists', menuItems(r.ui).every((i) => typeof r.s[i.fn] === 'function'), true);
}

console.log('1. created-report menu: the normal report operations, no CENTRAL submenu');
let addonMenuFns;
{
  const addon = runtime({ release: false });
  addon.s.onOpen();
  addonMenuFns = menuItems(addon.ui).map((i) => i.fn);
  check('add-on runtime (no Release.js) menu is unchanged: submenus',
    submenus(addon.ui), ['📝 INITIAL SETUP', '🚀 REPORT OPERATIONS', '📋 DOCUMENT MANAGEMENT', '🔧 TOOLS & DIAGNOSTICS', '🎨 FORMATTING & FIXES', '📧 EMAIL EXPORT', '☁️ CENTRAL ADD-ON (hourly)']);
  check('add-on runtime: no template items', addonMenuFns.filter((f) => /finishReportSetup|showTemplateInfo|showCreateReportDialog/.test(f)), []);

  const r = runtime({ docId: REPORT_ID });
  r.s.onOpen();
  const items = menuItems(r.ui);
  check('submenus: the accepted structure without CENTRAL, plus LEGACY',
    submenus(r.ui), ['📝 INITIAL SETUP', '🚀 REPORT OPERATIONS', '📋 DOCUMENT MANAGEMENT', '🔧 TOOLS & DIAGNOSTICS', '🎨 FORMATTING & FIXES', '📧 EMAIL EXPORT', '🗄️ LEGACY (old workflow)']);
  check('CENTRAL add-on operations absent', items.filter((i) => /ForAddon$|CENTRAL/.test(i.fn + i.path)), []);
  const addonOnly = ['enableAutomaticUpdatesForAddon', 'disableAutomaticUpdatesForAddon', 'setAutomaticUpdateIntervalForAddon', 'showAddonSchedulerStatusForAddon'];
  check('no existing report operation was dropped',
    addonMenuFns.filter((f) => addonOnly.indexOf(f) === -1 && !items.some((i) => i.fn === f)), []);
  check('first item of a new report: Finish Report Setup', items[0], { path: '🚀 Finish Report Setup', fn: 'finishReportSetup' });
  check('Run Full Report Build is the first report operation',
    items.filter((i) => /^🚀 REPORT OPERATIONS > /.test(i.path))[0], { path: '🚀 REPORT OPERATIONS > ▶️ Run Full Report Build', fn: 'runFullReportBuild' });
  check('the old "Legacy: Build Initial Report" is only in the LEGACY submenu',
    items.filter((i) => i.fn === 'buildInitialReport').map((i) => i.path), ['🗄️ LEGACY (old workflow) > 📝 Legacy: Build Initial Report']);
  check('About This Report is the last item', items[items.length - 1], { path: 'ℹ️ About This Report', fn: 'showTemplateInfo' });
  check('every target exists', items.filter((i) => typeof r.s[i.fn] !== 'function'), []);

  const done = runtime({ docId: REPORT_ID, props: { SA4_BOOTSTRAP_STATE: 'done' } });
  done.s.onOpen();
  check('after setup the Finish Report Setup item is gone', menuItems(done.ui).some((i) => i.fn === 'finishReportSetup'), false);
}

// ================================================================ 2. guard

console.log('2. template guard: report operations are refused in the master template');
{
  const r = runtime({ docId: TEMPLATE_ID, triggers: [] });
  const refused = /^This is the SA4 Report Template itself/;
  check('Run Full Report Build', refused.test(thrown(() => r.s.runFullReportBuild())), true);
  check('Build Skeleton + TDoc Tables (build guard)', refused.test(thrown(() => r.s.buildSkeletonWithTdocTables())), true);
  check('Configure Meeting', refused.test(thrown(() => r.s.configureMeetingSettings())), true);
  check('Save configuration', refused.test(thrown(() => r.s.saveConfigurationSettings({ meetingId: '86172', showPreview: true }))), true);
  check('Continuous Update trigger', refused.test(thrown(() => r.s.createContinuousTrigger(30, false))), true);
  const logs = [];
  r.s.Logger = { log: (m) => logs.push(String(m)) };
  r.s.continuousUpdate();
  check('Continuous Update itself stops at the guard', logs.some((l) => /SA4 Report Template itself/.test(l)), true);
  check('nothing was written, shown or scheduled', [r.docProps.getKeys(), r.ui.dialogs, r.triggers.calls, r.drive.calls], [[], [], [], []]);
  check('Finish Report Setup refuses too', r.s.finishReportSetupWith_(r.s.liveTemplateDeps_()).status, 'refused');
}

console.log('2. the three kinds of document are told apart by document id, not by inherited properties');
{
  const master = runtime({ docId: TEMPLATE_ID });
  const report = runtime({ docId: REPORT_ID });
  const addon = runtime({ docId: TEMPLATE_ID, release: false });
  check('master / report / same id without Release.js',
    [master.s.isTemplateMasterDocument_(), report.s.isTemplateMasterDocument_(), addon.s.isTemplateMasterDocument_()], [true, false, false]);
  check('a copy that inherited nothing is a report', report.s.templateDocumentOrigin_(REPORT_ID, RELEASE, null), 'not-set-up');
  check('origins', [
    report.s.templateDocumentOrigin_(TEMPLATE_ID, RELEASE, 'done'), report.s.templateDocumentOrigin_(REPORT_ID, RELEASE, 'done'),
    report.s.templateDocumentOrigin_(REPORT_ID, RELEASE, 'manual')], ['master-template', 'created-report', 'ordinary-document']);
  check('a background (CENTRAL) context never consults the template guard',
    /SA4 Report Template itself/.test(thrown(() => master.s.assertMeetingReadyToBuild_({ mode: 'addon-background', documentId: 'X' })) || ''), false);
}

// ================================================================ 3. creator end to end

console.log('3. Create New SA4 Report: dialog -> lookup -> DriveApp.makeCopy -> setup information');
let scenario;
{
  const drive = fakeDrive();
  drive.add(TEMPLATE_ID, 'SA4 Report Template');
  const master = runtime({ docId: TEMPLATE_ID, drive });
  master.s.showCreateReportDialog();
  const dialog = master.ui.dialogs[0];
  check('dialog shown with the release it runs', [dialog.title, /Template release T-2026\.10\.0/.test(dialog.html)], ['Create New SA4 Report', true]);
  // TEMPLATE-002C added the optional Mailing list field; it stays hidden until a meeting is looked up.
  check('only a meeting ID is asked for up front',
    [(dialog.html.match(/<input /g) || []).length, /id="mailingRow" style="display:none"/.test(dialog.html), /id="familyRow" style="display:none"/.test(dialog.html)], [2, true, true]);

  const d = runDialogScript(dialog.html, {
    previewNewReportFromTemplate: serverCall(master.s.previewNewReportFromTemplate),
    previewNewReportFromResolved: serverCall(master.s.previewNewReportFromResolved),
    createNewReportFromTemplate: serverCall(master.s.createNewReportFromTemplate)
  }, ['meetingId', 'family', 'createBtn', 'lookupBtn', 'details', 'status', 'familyRow', 'title', 'result']);
  check('Create is disabled until a meeting is looked up', /id="createBtn"[^>]*disabled/.test(dialog.html), true);

  d.registry.meetingId.value = '86172';
  d.globals.lookUp();
  check('lookup: family detected, nothing to answer', [d.registry.family.value, d.registry.createBtn.disabled, /class="err"/.test(d.registry.status.innerHTML)], ['MBS', false, false]);
  check('proposal names the title', d.registry.title.textContent, 'New report: MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e');
  check('proposal shows the e-mail collection start = meeting start date', /E-mail collection from<\/td><td>2026-10-01</.test(d.registry.details.innerHTML), true);
  check('nothing was copied by looking up', drive.calls, []);

  d.registry.family.value = 'Video';
  d.globals.familyChanged();
  check('choosing another family re-proposes without a new lookup', d.registry.title.textContent, 'New report: Video SWG Minutes – SA4-e (AH) MBS SWG post 137-e');
  d.registry.family.value = 'MBS';
  d.globals.familyChanged();

  d.globals.createReport();
  const copyId = Object.keys(drive.files).find((id) => id !== TEMPLATE_ID);
  check('exactly one DriveApp.makeCopy (title, same folder), then one description write',
    drive.calls, [['makeCopy', 'MBS SWG Minutes – SA4-e (AH) MBS SWG post 137-e', 'FOLDER'], ['setDescription', copyId]]);
  check('the dialog links to the new report', d.registry.result.innerHTML.indexOf('href="https://docs.google.com/document/d/' + copyId + '/edit"') !== -1, true);
  const payload = JSON.parse(drive.files[copyId].description.split('SA4-BOOTSTRAP-V1:')[1]);
  check('setup information: only the bootstrap fields, bound to the copy',
    [payload.targetDocumentId, payload.templateRelease, payload.config.meetingId, payload.config.reportType, payload.config.meetingType, payload.config.emailStartDate],
    [copyId, 'T-2026.10.0', '86172', 'MBS', 'adhoc', '2026-10-01']);
  check('the master template itself is untouched', [master.docProps.getKeys(), drive.files[TEMPLATE_ID].description], [[], '']);
  scenario = { drive, copyId };
}

console.log('3. the creator works only in the master template');
{
  const drive = fakeDrive();
  const report = runtime({ docId: REPORT_ID, drive });
  report.s.showCreateReportDialog();
  check('in a report: a message, no dialog', [report.ui.dialogs.length, /SA4 Report Template document only/.test(report.ui.alerts[0][1])], [0, true]);
  const core = plain(report.s.previewNewReportFromTemplate('86172', {}));
  const res = plain(report.s.createNewReportFromTemplate('86172', core.resolved, {}));
  check('the RPC refuses too, nothing copied', [res.ok, res.errors, drive.calls], [false, ['New reports are created from the SA4 Report Template document only.'], []]);
  const stale = plain(report.s.createNewReportFromTemplate('85916', core.resolved, {}));
  check('a meeting ID that does not match the looked-up meeting is refused', stale.ok, false);
}

console.log('3. the new report: its first use stores the setup information, then the normal workflow');
{
  const { drive, copyId } = scenario;
  // A copy is a new project: empty property stores, no triggers (TEMPLATE-001).
  const report = runtime({ docId: copyId, drive });
  report.s.onOpen();
  check('menu offers Finish Report Setup and Run Full Report Build',
    ['finishReportSetup', 'runFullReportBuild'].map((f) => menuItems(report.ui).some((i) => i.fn === f)), [true, true]);

  report.ui.alertResponse = 'NO'; // decline the "Continue?" question: only the setup runs
  report.s.runFullReportBuild();
  const p = report.docProps._store;
  check('clicking Run Full Report Build first stored the configuration',
    [p.MEETING_ID, p.MEETING_TYPE, p.REPORT_SUFFIX, p.EMAIL_START_DATE, p.SA4_BOOTSTRAP_STATE, p.SA4_CREATED_FROM_RELEASE],
    ['86172', 'adhoc', 'MBS', '2026-10-01', 'done', 'T-2026.10.0']);
  check('ready to build (the real build guard)', plain(report.s.getBuildReadiness_()).ready, true);
  check('setup information removed from the description', drive.files[copyId].description.indexOf('SA4-BOOTSTRAP-V1:'), -1);
  check('no trigger appeared by itself', report.triggers.handlers(), []);

  const before = JSON.stringify(p);
  report.s.configureMeetingSettings();
  check('Configure Meeting opens pre-filled and changes nothing',
    [report.ui.dialogs[0].title, /id="meetingId" value="86172"/.test(report.ui.dialogs[0].html), JSON.stringify(report.docProps._store) === before], ['Meeting Configuration', true, true]);

  const info = report.s.describeTemplateRuntime_({ release: RELEASE, documentId: copyId, scriptId: 'S', bootstrapState: 'done', createdAt: 'x', createdFromRelease: 'T-2026.10.0', meetingId: '86172', continuousInterval: null });
  check('About This Report: origin, release, pinned', [info[0], /T-2026\.10\.0 \(Code\.js 2\.17\.0, commit abcdef0\)/.test(info.join('\n')), /keeps this script release/.test(info.join('\n'))],
    ['Report created from the SA4 Report Template', true, true]);
}

console.log('3. an ordinary document that merely contains the code is configured by hand, as before');
{
  const drive = fakeDrive();
  const ordinary = runtime({ docId: REPORT_ID, drive });
  ordinary.s.configureMeetingSettings();
  check('Configure Meeting opens; the document is marked as checked once, nothing else',
    [ordinary.ui.dialogs.length, plain(ordinary.docProps._store)], [1, { SA4_BOOTSTRAP_STATE: 'manual' }]);
  ordinary.s.saveConfigurationSettings({ meetingId: '86172', meetingType: 'adhoc', reportType: 'MBS', showPreview: true });
  check('manual configuration is saved normally', ordinary.docProps._store.MEETING_ID, '86172');
  check('its origin', ordinary.s.templateDocumentOrigin_(REPORT_ID, RELEASE, ordinary.docProps._store.SA4_BOOTSTRAP_STATE), 'ordinary-document');
}

// ================================================================ 4. Email Collection Start Date

const CONFIG_IDS = ['meetingId', 'meetingType', 'meetingName', 'meetingDate', 'ftpBase', 'agendaTdoc', 'mailingList', 'meetingFolder', 'meetingNumber',
  'reportType', 'familyInfo', 'readinessRules', 'tdocUrl', 'revisionsUrl', 'agendaSourceDocId', 'showPreview', 'apiToken', 'clearApiToken', 'discussionEmailSender'];
function openConfig(props, release) {
  const r = runtime({ docId: REPORT_ID, props: Object.assign({ SA4_BOOTSTRAP_STATE: 'manual' }, props || {}), release });
  r.s.configureMeetingSettings();
  const html = r.ui.dialogs[0].html;
  const d = runDialogScript(html, {
    resolveMeetingForConfigDialog: serverCall(r.s.resolveMeetingForConfigDialog),
    discoverAgendaForConfigDialog: serverCall(r.s.discoverAgendaForConfigDialog),
    saveConfigurationSettings: serverCall((config) => { r.s.saveConfigurationSettings(config); return null; })
  }, CONFIG_IDS);
  return Object.assign(r, { html, d });
}
const startField = (c) => c.d.registry.collectionStartField;

console.log('4. Email Collection Start Date: a normal Configure Meeting field in the template runtime');
{
  const c = openConfig({});
  check('the field is in the dialog, in Options', /<h3>3\. Options<\/h3>\s*<label>Email Collection Start Date:<\/label><input type="date"/.test(c.html), true);
  check('new document, meeting not resolved yet: empty, not explicit', [startField(c).value, startField(c).getAttribute('data-explicit')], ['', 'false']);

  // (1) new meeting -> defaults to the meeting start date
  c.d.registry.meetingId.value = '86172';
  c.d.globals.resolveMeeting();
  check('(1) Resolve proposes the meeting start date', startField(c).value, '2026-10-01');
  c.d.globals.saveConfig();
  check('(1) saved through the existing EMAIL_START_DATE property', c.docProps._store.EMAIL_START_DATE, '2026-10-01');
  check('(1) no second property for the same thing', c.docProps.getKeys().filter((k) => /START/i.test(k)), ['EMAIL_START_DATE']);
  check('(1) the collector uses it', c.s.getCollectorConfig_().EMAIL_START_DATE, '2026-10-01');
}
{
  // (2) an explicit stored value is preserved
  const c = openConfig({ EMAIL_START_DATE: '2026-09-15', MEETING_ID: '86172', MEETING_DATE: 'October 1, 2026' });
  check('(2) the stored value is shown as the saved value', [startField(c).value, startField(c).getAttribute('data-explicit'), /Saved value\./.test(c.html)], ['2026-09-15', 'true', true]);
  c.d.registry.meetingId.value = '86172';
  c.d.globals.resolveMeeting();
  check('(2) Resolve does not replace it with the meeting start date', startField(c).value, '2026-09-15');
  c.d.globals.saveConfig();
  check('(2) Save keeps it', c.docProps._store.EMAIL_START_DATE, '2026-09-15');

  // ... also when the creator's setup information arrives in a document that already has one
  const drive = fakeDrive();
  const master = runtime({ docId: TEMPLATE_ID, drive: (drive.add(TEMPLATE_ID, 'T'), drive) });
  const proposal = plain(master.s.previewNewReportFromTemplate('86172', {}));
  const made = plain(master.s.createNewReportFromTemplate('86172', proposal.resolved, {}));
  const copy = runtime({ docId: made.documentId, drive, props: { EMAIL_START_DATE: '2026-09-15' } });
  copy.s.finishReportSetupWith_(copy.s.liveTemplateDeps_());
  check('(2) first run never overwrites an explicit stored date', [copy.docProps._store.MEETING_ID, copy.docProps._store.EMAIL_START_DATE], ['86172', '2026-09-15']);
}
{
  // (3) a user-edited value is persisted and used by the collector
  const c = openConfig({ MEETING_ID: '86172', MEETING_TYPE: 'adhoc', MEETING_DATE: 'October 1, 2026', REPORT_SUFFIX: 'MBS' });
  check('(3) proposal before editing: the meeting start date', startField(c).value, '2026-10-01');
  startField(c).value = '2026-09-20';
  startField(c).touched = true;
  c.d.registry.meetingId.value = '86172';
  c.d.globals.resolveMeeting();
  check('(3) a typed date survives Resolve', startField(c).value, '2026-09-20');
  c.d.globals.saveConfig();
  check('(3) the payload carries it under the dialog\'s config key', c.d.sent().emailStartDate, '2026-09-20');
  check('(3) persisted', c.docProps._store.EMAIL_START_DATE, '2026-09-20');
  check('(3) used by the collector (earlier than the meeting)', c.s.getCollectorConfig_().EMAIL_START_DATE, '2026-09-20');
  startField(c).value = '2026-10-05';
  c.d.globals.saveConfig();
  check('(3) ... and it can be moved later', c.s.getCollectorConfig_().EMAIL_START_DATE, '2026-10-05');
}
{
  // (4) an existing report without the property -> safe fallback to the meeting start date
  const c = openConfig({ MEETING_ID: '86172', MEETING_TYPE: 'adhoc', MEETING_DATE: 'October 1, 2026', REPORT_SUFFIX: 'MBS' });
  check('(4) dialog proposes the meeting start date, marked as the default',
    [startField(c).value, startField(c).getAttribute('data-explicit'), /Default: the meeting start date\./.test(c.html)], ['2026-10-01', 'false', true]);
  check('(4) the collector falls back to the meeting start date', c.s.getCollectorConfig_().EMAIL_START_DATE, '2026-10-01');
  check('(4) nothing was written just by opening the dialog', c.docProps._store.EMAIL_START_DATE, undefined);

  const noDate = openConfig({ MEETING_ID: '86172', MEETING_TYPE: 'adhoc', REPORT_SUFFIX: 'MBS' });
  check('(4) no meeting date known: empty field, historical default 2026-08-21', [startField(noDate).value, noDate.s.getCollectorConfig_().EMAIL_START_DATE], ['', '2026-08-21']);
  const odd = openConfig({ MEETING_ID: '1', MEETING_DATE: '28-30 September 2026' });
  check('(4) an unreadable meeting date is never guessed', [startField(odd).value, odd.s.getCollectorConfig_().EMAIL_START_DATE], ['', '2026-08-21']);
  const table = openConfig({ MEETING_DATE: 'October 1, 2026' });
  makeTableOverride(table);
  check('(4) a Collector Configuration table row still wins', table.s.getCollectorConfig_().EMAIL_START_DATE, '2026-09-01');
}
function makeTableOverride(c) {
  const body = c.s.DocumentApp.getActiveDocument().getBody();
  body.insertTable(0, [['Key', 'Value'], ['EMAIL_START_DATE', '2026-09-01']]);
}

console.log('4. validation: the collector\'s own rule, before anything is written');
{
  const r = runtime({ docId: REPORT_ID, props: { EMAIL_START_DATE: '2026-09-15', MEETING_ID: '1' } });
  const before = JSON.stringify(r.docProps._store);
  ['15.09.2026', '2026-9-5', 'tomorrow', '2026-13-40'].forEach((bad) => {
    check(`"${bad}" is refused, nothing changed`,
      [thrown(() => r.s.saveConfigurationSettings({ meetingId: '2', emailStartDate: bad, showPreview: true })), JSON.stringify(r.docProps._store) === before],
      ['Email Collection Start Date must be a date in the form YYYY-MM-DD.', true]);
  });
  r.s.saveConfigurationSettings({ meetingId: '1', emailStartDate: '', showPreview: true });
  check('a blank value never erases a stored date', r.docProps._store.EMAIL_START_DATE, '2026-09-15');
  check('dialog and collector share one rule',
    ['2026-10-01', '2026-13-40', '15.09.2026', ''].map((v) => r.s.isValidCollectorStartDate_(v)), [true, false, false, false]);
  check('meeting date -> ISO', ['October 1, 2026', 'September 28, 2026', '2026-10-01 15:30:00', '28-30 September 2026', 'Octember 1, 2026', '', null]
    .map((v) => r.s.meetingStartDateIso_(v)), ['2026-10-01', '2026-09-28', '2026-10-01', '', '', '', '']);
}

console.log('4. outside the template runtime nothing changed (CENTRAL ADDON-007B1, TEMPLATE-002A)');
{
  const c = openConfig({ MEETING_ID: '86172', MEETING_TYPE: 'adhoc', MEETING_DATE: 'October 1, 2026', REPORT_SUFFIX: 'MBS' }, false);
  check('no field, no mention in the dialog', /emailStartDate|Email Collection Start Date|collectionStartField"/.test(c.html), false);
  c.d.globals.saveConfig();
  check('the payload has no start date', 'emailStartDate' in c.d.sent(), false);
  c.s.saveConfigurationSettings({ meetingId: '86172', emailStartDate: '2026-09-20', showPreview: true });
  check('Save does not write it', c.docProps._store.EMAIL_START_DATE, undefined);
  check('the collector keeps the 2026-08-21 fallback', c.s.getCollectorConfig_().EMAIL_START_DATE, '2026-08-21');
}

// ================================================================ 5. triggers

function triggerRuntime(existing, props) {
  return runtime({ docId: REPORT_ID, triggers: existing || [], props: Object.assign({ SA4_BOOTSTRAP_STATE: 'done' }, props || {}) });
}
function triggerDialog(r) {
  r.s.manageTriggers();
  const html = r.ui.dialogs[r.ui.dialogs.length - 1].html;
  const options = [];
  html.replace(/<option value="([^"]*)"( selected)?>([^<]*)<\/option>/g, (m, v, sel, label) => { options.push([v, !!sel, label]); return m; });
  return { html, options };
}

console.log('5. Continuous Update in the template runtime: 15 / 30 / 60 minutes');
{
  const r = triggerRuntime();
  const dlg = triggerDialog(r);
  check('the dialog offers 15, 30 (default) and 60 minutes', dlg.options,
    [['15', false, 'Every 15 minutes (Active meeting)'], ['30', true, 'Every 30 minutes (Recommended)'], ['60', false, 'Every hour (Slow meeting)']]);
  check('status: inactive', /Status:<\/strong> ❌ Inactive/.test(dlg.html), true);

  [[15, 'everyMinutes(15)'], [30, 'everyMinutes(30)'], [60, 'everyHours(1)']].forEach(([minutes, call]) => {
    const t = triggerRuntime();
    t.s.createContinuousTrigger(minutes, false);
    check(`${minutes} minutes -> ${call}, handler continuousUpdate`, t.triggers.calls, [['create', 'continuousUpdate', call]]);
    check(`${minutes} minutes: interval recorded and reported`,
      [t.docProps._store.CONTINUOUS_UPDATE_INTERVAL_MINUTES, plain(t.s.getTriggerStatus())], [String(minutes), { active: true, interval: minutes + ' minutes' }]);
  });
  const hour = triggerRuntime();
  hour.s.createContinuousTrigger('60', false);
  check('every hour never uses everyMinutes(60) (the Legacy defect)', hour.triggers.calls.some((c) => /everyMinutes/.test(c[2])), false);
  check('every offered interval uses a value ClockTriggerBuilder accepts',
    plain(hour.s.templateContinuousTriggerIntervals_()).every((o) => (o.everyMinutes ? [1, 5, 10, 15, 30].indexOf(o.everyMinutes) !== -1 : o.everyHours >= 1)), true);
}

console.log('5. the installed interval is shown');
{
  const r = triggerRuntime();
  r.s.createContinuousTrigger(15, true);
  const dlg = triggerDialog(r);
  check('status names the interval', /Status:<\/strong> ✅ Active<br>\s*<strong>Running every:<\/strong> 15 minutes/.test(dlg.html), true);
  check('the running interval is preselected', dlg.options.filter((o) => o[1]).map((o) => o[0]), ['15']);
  check('the abstracts switch is stored with it', r.docProps._store.FETCH_ABSTRACTS_ON_UPDATE, 'true');
}

console.log('5. replacement: the new trigger exists before the old one is removed; others are preserved');
{
  const r = triggerRuntime(['continuousUpdate', 'someOtherHandler', 'template001TriggerTick']);
  r.s.createContinuousTrigger(30, false);
  check('create first, then delete the previous continuousUpdate trigger only',
    r.triggers.calls, [['create', 'continuousUpdate', 'everyMinutes(30)'], ['delete', 'continuousUpdate', 'old0']]);
  check('exactly one continuousUpdate trigger; unrelated triggers untouched', r.triggers.handlers(), ['someOtherHandler', 'template001TriggerTick', 'continuousUpdate']);

  r.triggers.calls.length = 0;
  r.s.createContinuousTrigger(30, false);
  r.s.createContinuousTrigger(30, false);
  check('repeating the request is idempotent: still exactly one', r.triggers.handlers().filter((h) => h === 'continuousUpdate').length, 1);
}

console.log('5. a failed or invalid request leaves the running trigger alone');
{
  const r = triggerRuntime(['continuousUpdate', 'someOtherHandler'], { CONTINUOUS_UPDATE_INTERVAL_MINUTES: '30' });
  r.triggers.failCreate = true;
  check('creation fails (quota): the error is reported', thrown(() => r.s.createContinuousTrigger(15, true)), 'Service error: quota');
  check('... the working trigger was not deleted, the recorded interval and abstracts switch are unchanged',
    [r.triggers.handlers(), r.triggers.calls.filter((c) => c[0] === 'delete'), r.docProps._store.CONTINUOUS_UPDATE_INTERVAL_MINUTES, r.docProps._store.FETCH_ABSTRACTS_ON_UPDATE],
    [['continuousUpdate', 'someOtherHandler'], [], '30', undefined]);

  r.triggers.failCreate = false;
  r.triggers.calls.length = 0;
  [45, 5, 120, '60 minutes', 'abc', NaN, null, undefined, ''].forEach((bad) => {
    const msg = thrown(() => r.s.createContinuousTrigger(bad, true));
    check(`interval ${JSON.stringify(bad === undefined ? 'undefined' : bad)} is refused before ScriptApp is touched`,
      [/^Unsupported update interval: .* \(supported: 15, 30, 60 minutes\)\.$/.test(msg || ''), r.triggers.calls, r.triggers.handlers()], [true, [], ['continuousUpdate', 'someOtherHandler']]);
  });
}

console.log('5. disabling');
{
  const r = triggerRuntime(['someOtherHandler']);
  r.s.createContinuousTrigger(30, false);
  r.s.deleteContinuousTrigger();
  check('Stop Trigger removes only continuousUpdate and forgets the interval',
    [r.triggers.handlers(), r.docProps._store.CONTINUOUS_UPDATE_INTERVAL_MINUTES, plain(r.s.getTriggerStatus())], [['someOtherHandler'], undefined, { active: false, interval: null }]);
  r.s.deleteContinuousTrigger();
  check('stopping twice is harmless', r.triggers.handlers(), ['someOtherHandler']);
}

console.log('5. a new report has no trigger until the user starts one');
{
  const fresh = triggerRuntime([], { SA4_BOOTSTRAP_STATE: null });
  fresh.s.onOpen();
  check('opening a report creates nothing', [fresh.triggers.calls, fresh.triggers.handlers()], [[], []]);
}

console.log('5. the add-on runtime keeps its hourly-only rule');
{
  const addon = runtime({ release: false, triggers: ['continuousUpdate'] });
  check('30 minutes is still refused without Release.js', /\(supported: 60 minutes\)\.$/.test(thrown(() => addon.s.createContinuousTrigger(30, false)) || ''), true);
  const code = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  check('Code.js itself contains no sub-hourly trigger call; the installer is in ReportCreator.js',
    [/\.everyMinutes\(/.test(code), /\.everyMinutes\(/.test(fs.readFileSync(REPORT_CREATOR_PATH, 'utf8'))], [false, true]);
}

// ================================================================ 6. release bundle

console.log('6. release bundle: one Code.js, no probe code, version recorded');
{
  const ROOT = path.join(__dirname, '..');
  const plan = planRelease({
    releaseId: 'T-2026.10.0',
    target: { templateDocumentId: TEMPLATE_ID, templateScriptId: 'TEMPLATEscript00000000000000000000000000000000000000000' },
    git: { commit: 'a'.repeat(40), dirty: false, tagsAtHead: ['template-release/T-2026.10.0'] },
    readFile: (rel) => fs.readFileSync(path.join(ROOT, rel)), builtAt: '2026-10-01T10:00:00.000Z'
  });
  check('bundle plans cleanly', [plan.ok, plan.errors], [true, []]);
  check('exactly one Code.js, byte-identical to the repository file',
    [plan.files.filter((f) => /Code/.test(f.name)).map((f) => f.name), plan.files.find((f) => f.name === 'Code.js').content.equals(fs.readFileSync(CODE_JS_PATH))], [['Code.js'], true]);
  const codeVersion = (fs.readFileSync(CODE_JS_PATH, 'utf8').match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1];
  check('Release.js records the Code.js version', [/^2\.\d+\.\d+$/.test(codeVersion), plan.meta.codeVersion, String(plan.files.find((f) => f.name === 'Release.js').content).indexOf('"codeVersion": "' + codeVersion + '"') !== -1], [true, codeVersion, true]);
  check('TEMPLATE-001 probe code is not shipped',
    [plan.files.some((f) => /probe|Template001/i.test(f.name + f.source)), plan.files.some((f) => /template001|TEMPLATE001_/.test(String(f.content)))], [false, false]);

  // The three files of a bundle in one global scope, Release.js last: load
  // order must not matter, because the release is read at call time.
  const ctx = loadCode().sandbox;
  vm.runInContext(fs.readFileSync(REPORT_CREATOR_PATH, 'utf8'), ctx, { filename: 'ReportCreator.js' });
  check('before Release.js is loaded the runtime is the add-on runtime', ctx.templateRuntimeRelease_(), null);
  vm.runInContext(String(plan.files.find((f) => f.name === 'Release.js').content), ctx, { filename: 'Release.js' });
  check('after it, the template runtime', [ctx.templateRuntimeRelease_().releaseId, plain(ctx.validateTemplateRelease_(ctx.templateRuntimeRelease_()))], ['T-2026.10.0', []]);

  const creator = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8');
  const targets = [...creator.matchAll(/\.addItem\(\s*'[^']*'\s*,\s*'([A-Za-z0-9_]+)'\s*\)/g)].map((m) => m[1]);
  check('every menu target in ReportCreator.js exists', [targets.length > 0, targets.filter((t) => typeof ctx[t] !== 'function')], [true, []]);
  const names = (src) => new Set([...src.matchAll(/^(?:function|var|const|let)\s+([A-Za-z0-9_$]+)/gm)].map((m) => m[1]));
  const codeNames = names(fs.readFileSync(CODE_JS_PATH, 'utf8'));
  check('ReportCreator.js redefines nothing of Code.js', [...names(creator)].filter((n) => codeNames.has(n)), []);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll checks passed.');
process.exitCode = failures ? 1 : 0;
