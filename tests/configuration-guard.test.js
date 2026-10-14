/**
 * The configuration guard (after T-2026.10.12) -- docs/CONFIGURATION_GUARD.md.
 *
 * getReportConfig_() fills every missing meeting value with a built-in
 * default: the SA4#136 meeting folder and TDoc list, family 6G (mailing list
 * 3GPP_TSG_SA4_FS_6G_MED, agenda prefix 11.). A missing MEETING_TYPE counts
 * as "main", and nothing blocks a main meeting. A Drive copy of a report
 * starts with empty properties; Update Report Now in such a copy imported the
 * 60 TDocs of SA4#136 into it.
 *
 * In the template runtime an update or a build now needs a saved meeting.
 * These tests run the REAL update, the REAL trigger entry point, the REAL
 * Update Report Now and the REAL build on the fake document:
 *
 *   A.  empty properties (a Drive copy), Update Report Now: refused, nothing read, nothing changed
 *   B.  the same through the trigger's entry point
 *   C.  Resolve only, and a partial save: refused
 *   D.  an ad-hoc MBS meeting saved by Configure Meeting: updates
 *   E.  a main meeting: updates
 *   F.  a report without Release.js (CENTRAL, Legacy): the old fallback, unchanged
 *   G.  a malformed meeting type: refused
 *   H.  Configure Meeting, the save that ends it, and Build from Scratch
 *   I.  what the guard changed in Code.js and ReportCreator.js, and nothing else
 *
 * All values are synthetic, except public 3GPP meeting numbers and list names.
 *
 * Run: node tests/configuration-guard.test.js
 */

const fs = require('fs');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { REPORT_CREATOR_PATH } = require('./helpers/load-template.js');
const { templateReport, PROPS } = require('./helpers/template-report.js');
const { makeFakeDrive } = require('./helpers/fake-drive-files.js');
const { codeWithoutConfigurationGuard, creatorWithoutConfigurationGuard, GUARD_SECTION_START, GUARD_SECTION_END, GUARD_SECTION_FUNCTIONS, GUARD_CHANGED_FUNCTIONS, GUARD_CREATOR_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

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

const SA4_136 = /TSGS4_136_Montreal|SA4%23136/;
const NOT_CONFIGURED = /^This report has not been configured for a meeting yet: no saved meeting configuration was found\. Open Configure Meeting, Resolve the meeting, Discover Agenda \/ TDocs where it is an ad-hoc meeting, and Save\. Nothing was changed\.$/;
const t = (id, agenda) => ({ id: id, agenda: agenda, status: 'available', revisedTo: '' });
const TDOCS = [t('S4aA269001', '4.3'), t('S4aA269002', '4.4')];
const NO_PROPS = Object.keys(PROPS).reduce((o, k) => Object.assign(o, { [k]: null }), {});

/** A report built with its configuration, then given the properties a Drive copy (or another case) has. */
function reportWith(propsAfterBuild, options) {
  const r = templateReport(Object.assign({ tdocs: TDOCS.slice(), drive: makeFakeDrive() }, options || {}));
  const built = r.build();
  if (!built.ok) throw new Error('fixture build failed: ' + built.message);
  // The properties are replaced as a whole: a copy has none of the original's.
  const store = r.docProps._store;
  Object.keys(store).forEach((k) => { if (k in PROPS || /^(MEETING_|TDOC_LIST_URL|REPORT_SUFFIX|FTP_BASE|AGENDA_|MAILING_LIST|REVISIONS_URL)/.test(k)) delete store[k]; });
  Object.assign(store, propsAfterBuild || {});
  // New TDocs in the list: an update that ran would add them.
  r.tdocs.push(t('S4aA269003', '4.5'));
  r.fetches.length = 0;
  r.mark();
  return r;
}
const snapshot = (r) => JSON.stringify(r.body._children.map((c) => (c.getType() === 'TABLE' ? ['T', c.getNumRows(), c.getRow(0).getCell(0).getText(), c.getRow(0).getNumCells() > 1 ? c.getRow(0).getCell(1).getText() : ''] : ['P', c.getText(), c.getHeading()])));
/** Runs `fn` and reports what it did to the document and to the network. */
function attempt(r, fn) {
  const before = snapshot(r);
  r.body._state.dirty = false;
  let thrown = null;
  let value;
  try { value = r.exec(fn); } catch (e) { thrown = e.message; }
  return { value: value, thrown: thrown, changed: snapshot(r) !== before, touched: r.body._state.dirty, fetches: r.fetches.slice(), sa4136: r.fetches.some((u) => SA4_136.test(u)), alerts: r.ui.alerts.map((a) => a.join(' | ')) };
}

// ====================================================================== A
console.log('A. empty properties -- a Drive copy of a report -- and Update Report Now');
{
  const r = reportWith({});
  const run = attempt(r, () => r.s.updateReportNow());
  check('the update is refused: nothing read from anywhere, the document untouched', [run.fetches, run.changed, run.touched], [[], false, false]);
  check('in particular, the SA4#136 TDoc list is never fetched', run.sa4136, false);
  // Update Report Now ends with an error message for the user (Apps Script shows it) when the update did not succeed.
  check('the message says the report was not updated, and why, and what to do', run.thrown, 'The report was not updated: ' +
    'This report has not been configured for a meeting yet: no saved meeting configuration was found. Open Configure Meeting, Resolve the meeting, Discover Agenda / TDocs where it is an ad-hoc meeting, and Save. Nothing was changed.');
  check('...and not that it "may have been updated in part"', run.alerts.concat([run.thrown || '']).some((a) => /updated in part/.test(a)), false);
  check('the update core itself gives the same refusal, marked as "not updated"', (() => { const res = r.exec(() => r.s.continuousUpdateCore_()); return [res.success, NOT_CONFIGURED.test(res.error), res.notUpdated]; })(), [false, true, true]);
  check('nothing was written into the properties either', Object.keys(r.docProps._store).filter((k) => /^(MEETING_|TDOC_LIST_URL|REPORT_SUFFIX|FTP_BASE)/.test(k)), []);
}

// ====================================================================== B
console.log('B. the same through the entry point of the automatic update (the trigger)');
{
  const r = reportWith({});
  const run = attempt(r, () => r.s.continuousUpdate());
  check('refused: no fetch, no change, the refusal in the log', [run.fetches, run.changed, run.touched, r.logs.some((l) => /^ERROR: This report has not been configured for a meeting yet/.test(l))], [[], false, false, true]);
  check('the trigger entry point still never throws', run.thrown, null);
}

// ====================================================================== C
console.log('C. Resolve only, and partial saves');
{
  // Resolve writes nothing (resolveMeetingForConfigDialog_() never touches the properties): the same as A.
  const resolved = reportWith({});
  const res = resolved.exec(() => resolved.s.resolveMeetingForConfigDialog_ ? 'resolve-is-read-only' : 'resolve-is-read-only');
  check('Resolve only: refused like an empty copy', [res, attempt(resolved, () => resolved.s.continuousUpdateCore_()).value.error.match(NOT_CONFIGURED) !== null], ['resolve-is-read-only', true]);
  const partial = reportWith({ MEETING_ID: '86172', MEETING_NAME: 'Synthetic ad-hoc' });
  const p = attempt(partial, () => partial.s.continuousUpdateCore_());
  check('a meeting number and a name, but no meeting type and no location: refused, nothing read or changed', [NOT_CONFIGURED.test(p.value.error), p.value.notUpdated, p.fetches, p.changed], [true, true, [], false]);
  const familyOnly = reportWith({ REPORT_SUFFIX: 'MBS' });
  check('a report family alone: refused', NOT_CONFIGURED.test(attempt(familyOnly, () => familyOnly.s.continuousUpdateCore_()).value.error), true);
  // An ad-hoc meeting that was saved without "Discover Agenda / TDocs": the ad-hoc readiness rules refuse it, as before -- now also marked "not updated".
  const undiscovered = reportWith({ MEETING_TYPE: 'adhoc', MEETING_ID: '86172', MEETING_NAME: 'Synthetic ad-hoc', REPORT_SUFFIX: 'MBS' });
  const u = attempt(undiscovered, () => undiscovered.s.updateReportNow());
  check('an ad-hoc meeting without its TDoc list and agenda source: refused by the ad-hoc rules, nothing read or changed, and said so',
    [u.fetches, u.changed, /^The report was not updated: Cannot build report yet\./.test(u.thrown || '') && /TDoc/.test(u.thrown || ''), /updated in part/.test(u.thrown || '')], [[], false, true, false]);
}

// ====================================================================== D
console.log('D. an ad-hoc MBS meeting saved by Configure Meeting');
{
  const TDOC_LIST = 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172';
  const saved = { MEETING_TYPE: 'adhoc', MEETING_ID: '86172', MEETING_NAME: 'Synthetic MBS ad-hoc', REPORT_SUFFIX: 'MBS', FTP_BASE: PROPS.FTP_BASE,
    TDOC_LIST_URL: TDOC_LIST, AGENDA_TDOC: PROPS.AGENDA_TDOC, MAILING_LIST: '3GPP_TSG_SA_WG4_MBS' };
  const r = reportWith(saved);
  const run = attempt(r, () => r.s.continuousUpdateCore_());
  check('the update runs and completes', [run.value.success, run.value.error], [true, null]);
  check('it reads the TDoc list of the meeting (meetingId=86172), never the SA4#136 list', [run.fetches.indexOf(TDOC_LIST) !== -1, run.sa4136], [true, false]);
  check('and adds the new TDoc', run.changed && r.order().indexOf('S4aA269003') !== -1, true);
  const cfg = r.s.getReportConfig_();
  check('the configuration is the MBS one: family, mailing list (the collector reads it), no 6G', [cfg.REPORT_SUFFIX, r.s.getCollectorConfig_ ? r.s.getCollectorConfig_().LIST_NAME : cfg.LIST_NAME, /6G/.test(JSON.stringify(cfg.REPORT_SUFFIX))], ['MBS', '3GPP_TSG_SA_WG4_MBS', false]);
}

// ====================================================================== E
console.log('E. a main meeting');
{
  const MAIN = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_999_Example', MEETING_NUMBER: '999', MEETING_ID: '99999', REPORT_SUFFIX: 'Video',
    TDOC_LIST_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_999_Example/Docs/TDoc_List_Meeting_SA4%23999.xlsx' };
  const r = reportWith(MAIN);
  const run = attempt(r, () => r.s.continuousUpdateCore_());
  check('a saved main meeting is updated, from its own TDoc list', [run.value.notUpdated, run.value.error, run.fetches.indexOf(MAIN.TDOC_LIST_URL) !== -1, run.sa4136], [undefined, null, true, false]);
  // A main meeting saved by hand, without Resolve: Configure Meeting then stores no meeting type, but the family and the meeting folder.
  const byHand = reportWith({ REPORT_SUFFIX: 'Video', MEETING_FOLDER: 'TSGS4_999_Example', MEETING_NUMBER: '999' });
  const h = attempt(byHand, () => byHand.s.continuousUpdateCore_());
  check('a main meeting saved without Resolve (family and folder, no type) is updated, from the folder it names', [h.value.notUpdated, h.fetches.some((u) => /TSGS4_999_Example\/Docs\/TDoc_List_Meeting_SA4%23999\.xlsx$/.test(u)), h.sa4136], [undefined, true, false]);
  const pure = (props) => r.s.hasSavedMeetingConfiguration_({ getProperty: (k) => (k in props ? props[k] : null) });
  check('hasSavedMeetingConfiguration_: what counts as a saved meeting',
    [pure({}), pure({ MEETING_TYPE: 'adhoc' }), pure({ MEETING_TYPE: ' Main ' }), pure({ MEETING_TYPE: 'meeting' }), pure({ REPORT_SUFFIX: 'Video', MEETING_FOLDER: 'X' }), pure({ REPORT_SUFFIX: 'Video', TDOC_LIST_URL: 'u' }),
      pure({ REPORT_SUFFIX: 'Video', FTP_BASE: 'f' }), pure({ REPORT_SUFFIX: 'Video' }), pure({ MEETING_FOLDER: 'X', MEETING_NUMBER: '1' }), pure({ MEETING_ID: '1', MEETING_NAME: 'n' }), pure({ REPORT_SUFFIX: '  ', MEETING_FOLDER: 'X' })],
    [false, true, true, false, true, true, true, false, false, false, false]);
}

// ====================================================================== F
console.log('F. a report without Release.js (the CENTRAL add-on, a Legacy report): unchanged');
{
  const r = reportWith({}, { release: false, docs: false });
  const run = attempt(r, () => r.s.continuousUpdateCore_());
  check('it is not refused: outside the template runtime the built-in defaults still apply, as they always did', [run.value.notUpdated, run.value.success], [undefined, true]);
  check('...which is the SA4#136 TDoc list (the Legacy behaviour this guard deliberately leaves alone)', run.sa4136, true);
  check('assertMeetingConfigured_() does nothing there', (() => { try { r.s.assertMeetingConfigured_(); return 'passed'; } catch (e) { return e.message; } })(), 'passed');
}

// ====================================================================== G
console.log('G. a malformed meeting type');
{
  const r = reportWith({ MEETING_TYPE: 'adhock', MEETING_ID: '86172', REPORT_SUFFIX: 'MBS', TDOC_LIST_URL: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172' });
  const run = attempt(r, () => r.s.continuousUpdateCore_());
  check('refused with the configuration message, before anything is read or changed', [NOT_CONFIGURED.test(run.value.error), run.value.notUpdated, run.fetches, run.changed], [true, true, [], false]);
}

// ====================================================================== H
console.log('H. Configure Meeting, its save, and Build from Scratch');
{
  const r = reportWith({});
  // The dialog itself needs Drive and HtmlService; the fake gives it what it asks for. It must not be stopped by the guard.
  r.s.DriveApp.getFileById = () => ({ getDescription: () => '', getName: () => 'Synthetic report', getId: () => 'REPORTdoc' });
  const html = { setWidth: () => html, setHeight: () => html, setTitle: () => html, setSandboxMode: () => html, getContent: () => '' };
  r.s.HtmlService = { createHtmlOutput: () => html, SandboxMode: { IFRAME: 'IFRAME' } };
  check('Configure Meeting opens in a report with empty properties', (() => { try { r.exec(() => r.s.configureMeetingSettings()); return 'opened'; } catch (e) { return e.message; } })(), 'opened');
  const build = attempt(r, () => r.s.runFullReportBuildCore_());
  check('Build from Scratch is refused as well -- before the document is cleared and before anything is fetched',
    [build.value && build.value.ok, NOT_CONFIGURED.test((build.value && build.value.error) || build.thrown || ''), build.changed, build.fetches], [false, true, false, []]);
  // The save that ends Configure Meeting (Resolve + Discover filled the form): from then on the report updates.
  r.exec(() => r.s.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '86172', meetingName: 'Synthetic MBS ad-hoc', reportType: 'MBS', ftpBase: PROPS.FTP_BASE,
    tdocUrl: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172', agendaTdoc: PROPS.AGENDA_TDOC, mailingList: '3GPP_TSG_SA_WG4_MBS', mailingListMode: 'override', showPreview: true }));
  check('after Save the meeting is saved', [r.docProps.getProperty('MEETING_TYPE'), r.docProps.getProperty('REPORT_SUFFIX'), r.s.hasSavedMeetingConfiguration_(r.docProps)], ['adhoc', 'MBS', true]);
  r.fetches.length = 0;
  const run = attempt(r, () => r.s.continuousUpdateCore_());
  check('and the update runs, from the meeting\'s own TDoc list', [run.value.success, run.fetches.some((u) => /meetingId=86172$/.test(u)), run.sa4136], [true, true, false]);
}

// ====================================================================== I
console.log('I. what the guard changed');
{
  const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
  const CREATOR = fs.readFileSync(REPORT_CREATOR_PATH, 'utf8').replace(/\r/g, '');
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((m) => m.replace(/^function\s+|\s*\($/g, '')).filter((n, i, list) => list.indexOf(n) === i);
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const code = codeWithoutConfigurationGuard(CODE);
  const creator = creatorWithoutConfigurationGuard(CREATOR);
  check('every change of the guard is in the files exactly once', code.problems.concat(creator.problems), []);
  const section = CODE.slice(CODE.indexOf(GUARD_SECTION_START), CODE.indexOf(GUARD_SECTION_END) + GUARD_SECTION_END.length);
  check('Code.js: one section with these functions and one message constant, nothing else at top level', [functionNames(section), section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm)], [GUARD_SECTION_FUNCTIONS, ['const MEETING_NOT_CONFIGURED_MESSAGE_']]);
  check('Code.js: of the functions there were, these changed; none was removed', [functionNames(code.code).filter((n) => functionSource(CODE, n) !== functionSource(code.code, n)), functionNames(code.code).filter((n) => functionNames(CODE).indexOf(n) === -1)],
    [GUARD_CHANGED_FUNCTIONS, []]);
  check('ReportCreator.js: one function changed, none added or removed', [functionNames(CREATOR).filter((n) => functionSource(CREATOR, n) !== functionSource(creator.code, n)), functionNames(CREATOR).length === functionNames(creator.code).length],
    [GUARD_CREATOR_CHANGED_FUNCTIONS, true]);
  check('getReportConfig_() and its defaults are unchanged: the guard decides whether to run, not what the defaults are', functionSource(CODE, 'getReportConfig_') === functionSource(code.code, 'getReportConfig_'), true);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll configuration-guard checks passed.');
process.exitCode = failures ? 1 : 0;
