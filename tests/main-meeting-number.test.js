/**
 * The main-meeting number (after T-2026.10.23) -- docs/MAIN_MEETING_NUMBER.md, GitHub issue #9.
 *
 * In the integration test of T-2026.10.24 (a copy of Audio SWG SA4#138,
 * meeting 60780), Configure Meeting was saved after Resolve. Resolve fills
 * the document folder (FTP_BASE) of SA4#138 but not the Meeting Folder and
 * Meeting Number fields, and the save of a main meeting wrote a blank one as
 * the historical default, TSGS4_136_Montreal and 136. Update Report Now then
 * looked for .../TSGS4_138_Calgary/Docs/TDoc_List_Meeting_SA4%23136.xlsx --
 * a file that does not exist; 3GPP answers HTTP 403 for it -- and failed
 * with "may have been updated in part", although nothing had been changed.
 *
 *   A.  the meeting folder and number of a document folder (pure)
 *   B.  the save of issue #9: with 3dd4263 (2.26.8 before the fix) the SA4#136 list, now the SA4#138 list
 *   C.  a folder or number that contradicts the document folder: refused, nothing written
 *   D.  unchanged: no document folder (SA4#136 defaults), ad-hoc, a typed TDoc list URL
 *   E.  a report saved by the candidate: refused before any download and before the report is touched
 *   F.  getReportConfig_(): no number saved
 *   G.  a failed download: "not updated", the report untouched; a later failure still says "in part"
 *   H.  what changed in Code.js, and nothing else
 *
 * All values are synthetic, except public 3GPP meeting numbers, folders and list names.
 *
 * Run: node tests/main-meeting-number.test.js
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport, PROPS } = require('./helpers/template-report.js');
const { makeFakeDrive } = require('./helpers/fake-drive-files.js');
const { codeWithoutSessionTimes, codeWithoutMainMeetingNumber, MAIN_MEETING_NUMBER_SECTION_START, MAIN_MEETING_NUMBER_SECTION_END, MAIN_MEETING_NUMBER_SECTION_FUNCTIONS, MAIN_MEETING_NUMBER_CHANGED_FUNCTIONS } = require('./helpers/before-structure-fix.js');

const ROOT = path.join(__dirname, '..');
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

const FULL = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const CANDIDATE = codeWithoutMainMeetingNumber(FULL).code;

// SA4#138 as the Portal gives it (MtgDocURL) and as Resolve puts it into the form.
const CALGARY_FTP = 'https://ftp.3gpp.org/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/';
const CALGARY_WWW = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_138_Calgary/Docs/';
const LIST_138 = CALGARY_FTP + 'TDoc_List_Meeting_SA4%23138.xlsx';
const LIST_136_IN_138 = CALGARY_FTP + 'TDoc_List_Meeting_SA4%23136.xlsx';
const SA4_136 = /SA4%23136|TSGS4_136_Montreal/;
const t = (id, agenda) => ({ id: id, agenda: agenda, status: 'available', revisedTo: '' });
const TDOCS = [t('S4-260001', '4.3'), t('S4-260002', '4.4')];
const MEETING_KEYS = /^(MEETING_|TDOC_LIST_URL|REPORT_SUFFIX|FTP_BASE|AGENDA_|MAILING_LIST|REVISIONS_URL)/;

/** 3GPP's answer for a file that is not in a folder: HTTP 403 (www.3gpp.org/ftp, checked 2026-11-13). */
function answer403For(r, re) {
  const fetch = r.s.UrlFetchApp.fetch;
  r.s.UrlFetchApp.fetch = (url, options) => {
    if (re.test(String(url))) { r.fetches.push(String(url)); return { getResponseCode: () => 403, getBlob: () => ({ getBytes: () => [0x3c, 0x21], getContentType: () => 'text/html' }) }; }
    return fetch(url, options);
  };
}

/** A report built, then given these meeting properties (and nothing else of the build's). */
function reportWith(props, options) {
  const r = templateReport(Object.assign({ tdocs: TDOCS.slice() }, options || {}));
  const built = r.build();
  if (!built.ok) throw new Error('fixture build failed: ' + built.message);
  const store = r.docProps._store;
  Object.keys(store).forEach((k) => { if (k in PROPS || MEETING_KEYS.test(k)) delete store[k]; });
  Object.assign(store, props || {});
  r.tdocs.push(t('S4-260003', '4.5'));
  r.fetches.length = 0;
  r.mark();
  return r;
}
const snapshot = (r) => JSON.stringify(r.body._children.map((c) => (c.getType() === 'TABLE' ? ['T', c.getNumRows(), c.getRow(0).getCell(0).getText(), c.getRow(0).getNumCells() > 1 ? c.getRow(0).getCell(1).getText() : ''] : ['P', c.getText(), c.getHeading()])));
function attempt(r, fn) {
  const before = snapshot(r);
  r.body._state.dirty = false;
  let thrown = null;
  let value;
  try { value = r.exec(fn); } catch (e) { thrown = e.message; }
  return { value: value, thrown: thrown, changed: snapshot(r) !== before, touched: r.body._state.dirty, fetches: r.fetches.slice() };
}
const meetingProps = (r) => Object.keys(r.docProps._store).filter((k) => MEETING_KEYS.test(k)).sort().reduce((o, k) => Object.assign(o, { [k]: r.docProps._store[k] }), {});

/** The form as Resolve leaves it for meeting 60780 (SA4#138): type main, the document folder, no folder, no number, no list. */
const SAVE_138 = { meetingType: 'main', meetingId: '60780', meetingName: '3GPP TSG SA WG4 #138', meetingDate: 'November 16, 2026', ftpBase: CALGARY_FTP,
  meetingFolder: '', meetingNumber: '', tdocUrl: '', reportType: 'Audio', agendaTdoc: '', showPreview: true, mailingList: '', mailingListMode: 'derived', apiTokenAction: 'keep' };

// ====================================================================== A
console.log('A. the meeting folder and number of a document folder');
{
  const r = templateReport({ release: false });
  const of = (u) => r.s.mainMeetingFolderOfFtpBase_(u);
  check('SA4#138, as the Portal gives it and under www.3gpp.org/ftp', [of(CALGARY_FTP), of(CALGARY_WWW)], [{ folder: 'TSGS4_138_Calgary', number: '138' }, { folder: 'TSGS4_138_Calgary', number: '138' }]);
  check('SA4#136, an e-meeting (its number keeps "-e": TDoc_List_Meeting_SA4%23137-e.xlsx), a folder without a place, no trailing slash',
    [of('https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/'), of('https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137-e/Docs/'), of('https://www.3gpp.org/ftp/TSG_SA/WG4_CODEC/TSGS4_139/Docs'),
      r.s.mainMeetingNumberOfFtpBase_('https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137-e/Docs/')],
    [{ folder: 'TSGS4_136_Montreal', number: '136' }, { folder: 'TSGS4_137-e', number: '137-e' }, { folder: 'TSGS4_139', number: '139' }, '137-e']);
  check('no meeting in it: an ad-hoc series folder, the Inbox, nothing', [of(PROPS.FTP_BASE), of(CALGARY_FTP.replace('/Docs/', '/Inbox/')), of(''), of(null), r.s.mainMeetingNumberOfFtpBase_('')], [null, null, null, null, '']);
}

// ====================================================================== B
console.log('B. the save of issue #9: SA4#138 after Resolve, Meeting Folder and Number left blank');
{
  const before = templateReport({ release: false, code: CANDIDATE, props: Object.keys(PROPS).reduce((o, k) => Object.assign(o, { [k]: null }), {}) });
  before.exec(() => before.s.saveConfigurationSettings(Object.assign({}, SAVE_138)));
  check('the candidate (2.26.8) saved SA4#136\'s folder and number next to SA4#138\'s document folder',
    [before.docProps.getProperty('MEETING_FOLDER'), before.docProps.getProperty('MEETING_NUMBER'), before.docProps.getProperty('FTP_BASE'), before.docProps.getProperty('TDOC_LIST_URL')],
    ['TSGS4_136_Montreal', '136', CALGARY_FTP, null]);
  check('...so its TDoc list was TDoc_List_Meeting_SA4%23136.xlsx in the SA4#138 folder -- the first address of issue #9', before.s.getReportConfig_().TDOC_LIST_URL, LIST_136_IN_138);

  const now = templateReport({ release: false, props: Object.keys(PROPS).reduce((o, k) => Object.assign(o, { [k]: null }), {}) });
  now.exec(() => now.s.saveConfigurationSettings(Object.assign({}, SAVE_138)));
  check('now the folder and number are those of the document folder', [now.docProps.getProperty('MEETING_FOLDER'), now.docProps.getProperty('MEETING_NUMBER'), now.docProps.getProperty('FTP_BASE')],
    ['TSGS4_138_Calgary', '138', CALGARY_FTP]);
  check('...and the TDoc list is SA4#138\'s, the file the folder has', now.s.getReportConfig_().TDOC_LIST_URL, LIST_138);
  check('a folder and number typed as the document folder has them are saved as typed', (() => {
    const r = templateReport({ release: false, props: { MEETING_TYPE: null, TDOC_LIST_URL: null } });
    r.exec(() => r.s.saveConfigurationSettings(Object.assign({}, SAVE_138, { meetingFolder: 'TSGS4_138_Calgary', meetingNumber: '138' })));
    return [r.docProps.getProperty('MEETING_FOLDER'), r.docProps.getProperty('MEETING_NUMBER')];
  })(), ['TSGS4_138_Calgary', '138']);

  // The whole run, in a template report: Save, then Update Report Now.
  const r = reportWith({}, { drive: makeFakeDrive() });
  answer403For(r, /SA4%23136\.xlsx$/);
  r.exec(() => r.s.saveConfigurationSettings(Object.assign({}, SAVE_138)));
  r.fetches.length = 0;
  const run = attempt(r, () => r.s.updateReportNow());
  check('Update Report Now after that save: completes, from the SA4#138 list, never an SA4#136 one', [run.thrown, run.fetches.indexOf(LIST_138) !== -1, run.fetches.some((u) => SA4_136.test(u))], [null, true, false]);
}

// ====================================================================== C
console.log('C. a folder or number that contradicts the document folder');
{
  // What the dialog shows in a report the candidate saved: the SA4#136 defaults it wrote.
  const stale = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136', MEETING_ID: '60780', FTP_BASE: CALGARY_FTP, REPORT_SUFFIX: 'Audio', TDOC_LIST_URL: null };
  const r = templateReport({ release: false, props: stale });
  const before = JSON.stringify(meetingProps(r));
  let thrown = null;
  try { r.exec(() => r.s.saveConfigurationSettings(Object.assign({}, SAVE_138, { meetingFolder: 'TSGS4_136_Montreal', meetingNumber: '136' }))); } catch (e) { thrown = e.message; }
  check('the save is refused, saying what does not fit and what to do', thrown,
    'Meeting Folder "TSGS4_136_Montreal" and Meeting Number "136" do not belong to the meeting\'s document folder ' + CALGARY_FTP + ' (TSGS4_138_Calgary, 138): ' +
    'the TDoc list would be looked for under a wrong name. Correct them, or clear both to take them from the document folder. Nothing was saved.');
  check('...and nothing was written', JSON.stringify(meetingProps(r)), before);
  let numberOnly = null;
  try { r.exec(() => r.s.saveConfigurationSettings(Object.assign({}, SAVE_138, { meetingNumber: '136' }))); } catch (e) { numberOnly = e.message; }
  check('a wrong number alone is refused too', /^Meeting Folder "" and Meeting Number "136" do not belong/.test(numberOnly || ''), true);
  r.exec(() => r.s.saveConfigurationSettings(Object.assign({}, SAVE_138)));
  check('cleared, both are taken from the document folder', [r.docProps.getProperty('MEETING_FOLDER'), r.docProps.getProperty('MEETING_NUMBER')], ['TSGS4_138_Calgary', '138']);
  const e = templateReport({ release: false, props: { MEETING_TYPE: null, TDOC_LIST_URL: null } });
  e.exec(() => e.s.saveConfigurationSettings(Object.assign({}, SAVE_138, { ftpBase: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_137-e/Docs/', meetingNumber: '137-E' })));
  check('compared without case: "137-E" for TSGS4_137-e is saved as typed', e.docProps.getProperty('MEETING_NUMBER'), '137-E');
}

// ====================================================================== D
console.log('D. unchanged: no document folder, ad-hoc, a typed TDoc list URL');
{
  const none = Object.keys(PROPS).reduce((o, k) => Object.assign(o, { [k]: null }), {});
  const legacy = (code) => {
    const r = templateReport({ release: false, code: code, props: none });
    r.exec(() => r.s.saveConfigurationSettings(Object.assign({}, SAVE_138, { ftpBase: '', meetingId: '' })));
    return meetingProps(r);
  };
  check('a main meeting without a document folder: the SA4#136 defaults, as the candidate wrote them', legacy(undefined), legacy(CANDIDATE));
  check('...which are', (({ MEETING_FOLDER, MEETING_NUMBER, MEETING_ID }) => [MEETING_FOLDER, MEETING_NUMBER, MEETING_ID])(legacy(undefined)), ['TSGS4_136_Montreal', '136', '60777']);
  const adhoc = (code) => {
    const r = templateReport({ release: false, code: code, props: none });
    r.exec(() => r.s.saveConfigurationSettings({ meetingType: 'adhoc', meetingId: '86172', meetingName: 'Synthetic ad-hoc', reportType: 'MBS', ftpBase: PROPS.FTP_BASE,
      tdocUrl: 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=86172', agendaTdoc: PROPS.AGENDA_TDOC, showPreview: true, apiTokenAction: 'keep' }));
    return meetingProps(r);
  };
  check('an ad-hoc meeting: saved exactly as the candidate saved it (no folder, no number)', [adhoc(undefined), 'MEETING_NUMBER' in adhoc(undefined)], [adhoc(CANDIDATE), false]);
  const typedList = templateReport({ release: false, props: { MEETING_TYPE: null, TDOC_LIST_URL: null } });
  typedList.exec(() => typedList.s.saveConfigurationSettings(Object.assign({}, SAVE_138, { meetingNumber: '136', tdocUrl: LIST_138 })));
  check('a typed TDoc list URL is the list: a number that does not fit the folder is not refused then', [typedList.docProps.getProperty('MEETING_NUMBER'), typedList.s.getReportConfig_().TDOC_LIST_URL], ['136', LIST_138]);
}

// ====================================================================== E
console.log('E. a report saved by the candidate (the copy of issue #9): refused before any download');
{
  const SAVED = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_136_Montreal', MEETING_NUMBER: '136', MEETING_ID: '60780', MEETING_NAME: '3GPP TSG SA WG4 #138', FTP_BASE: CALGARY_FTP, REPORT_SUFFIX: 'Audio' };
  const r = reportWith(SAVED, { drive: makeFakeDrive() });
  answer403For(r, /SA4%23136\.xlsx$/);
  const run = attempt(r, () => r.s.updateReportNow());
  check('Update Report Now: refused with the reason, nothing fetched, the report untouched', [run.thrown, run.fetches, run.changed, run.touched], [
    'The report was not updated: The saved meeting number 136 is not the number of the saved document folder ' + CALGARY_FTP + ' (138), so the TDoc list would be looked for as TDoc_List_Meeting_SA4#136.xlsx. ' +
    'Open Configure Meeting and correct the Meeting Number. Nothing was changed.', [], false, false]);
  const build = attempt(r, () => r.s.runFullReportBuildCore_());
  check('Build from Scratch: refused before the report is cleared, nothing fetched', [build.value && build.value.ok, /^The saved meeting number 136 is not the number/.test((build.value && build.value.error) || build.thrown || ''), build.changed, build.fetches], [false, true, false, []]);
  const bg = attempt(r, () => r.s.continuousUpdate());
  check('the automatic update: refused, logged, never thrown', [bg.thrown, bg.fetches, bg.changed, r.logs.some((l) => /^ERROR: The saved meeting number 136/.test(l))], [null, [], false, true]);
  // The same report with the candidate: what the copy saw (both addresses 403), and nothing changed by it either.
  const old = reportWith(SAVED, { release: false, code: CANDIDATE });
  answer403For(old, /SA4%23136\.xlsx$/);
  const was = attempt(old, () => old.s.continuousUpdateCore_());
  check('with the candidate: the SA4#136 list in the SA4#138 folder, then its www twin, both HTTP 403 -- and the report untouched',
    [was.value.success, was.fetches, was.value.error, was.changed, was.touched],
    [false, [LIST_136_IN_138, CALGARY_WWW + 'TDoc_List_Meeting_SA4%23136.xlsx'],
      'Could not download TDOC list: HTTP 403 (' + LIST_136_IN_138 + '); HTTP 403 (' + CALGARY_WWW + 'TDoc_List_Meeting_SA4%23136.xlsx)', false, false]);
  check('...but it was not marked "not updated" (Update Report Now said "may have been updated in part")', was.value.notUpdated, undefined);
  const fine = reportWith(Object.assign({}, SAVED, { MEETING_FOLDER: 'TSGS4_138_Calgary', MEETING_NUMBER: '138' }), { drive: makeFakeDrive() });
  const ok = attempt(fine, () => fine.s.updateReportNow());
  check('number and folder that fit: updated from TDoc_List_Meeting_SA4%23138.xlsx', [ok.thrown, ok.fetches.indexOf(LIST_138) !== -1], [null, true]);
  const listed = reportWith(Object.assign({}, SAVED, { TDOC_LIST_URL: LIST_138 }), { drive: makeFakeDrive() });
  check('a saved TDoc list URL is the list: not refused for the number', attempt(listed, () => listed.s.updateReportNow()).thrown, null);
  const adhoc = reportWith(Object.assign({}, PROPS, { MEETING_NUMBER: '136', FTP_BASE: CALGARY_FTP }), { drive: makeFakeDrive() });
  check('an ad-hoc meeting is never refused for it', adhoc.s.mainMeetingNumberMismatch_(adhoc.docProps), '');
}

// ====================================================================== F
console.log('F. getReportConfig_(): no number saved');
{
  const cfg = (props, code) => templateReport({ release: false, code: code, props: Object.assign(Object.keys(PROPS).reduce((o, k) => Object.assign(o, { [k]: null }), {}), props) }).s.getReportConfig_();
  check('a saved document folder of SA4#138: number 138 and its list', [cfg({ FTP_BASE: CALGARY_FTP }).MEETING_NUMBER, cfg({ FTP_BASE: CALGARY_FTP }).TDOC_LIST_URL], ['138', LIST_138]);
  check('the candidate gave 136 there', cfg({ FTP_BASE: CALGARY_FTP }, CANDIDATE).TDOC_LIST_URL, LIST_136_IN_138);
  check('nothing saved, an ad-hoc folder, a saved number: as before', [cfg({}).TDOC_LIST_URL, cfg({ FTP_BASE: PROPS.FTP_BASE }).MEETING_NUMBER, cfg({ FTP_BASE: CALGARY_FTP, MEETING_NUMBER: '137' }).MEETING_NUMBER],
    [cfg({}, CANDIDATE).TDOC_LIST_URL, '136', '137']);
}

// ====================================================================== G
console.log('G. a failed download, and a failure after it');
{
  const MAIN = { MEETING_TYPE: 'main', MEETING_FOLDER: 'TSGS4_138_Calgary', MEETING_NUMBER: '138', MEETING_ID: '60780', FTP_BASE: CALGARY_FTP, REPORT_SUFFIX: 'Audio' };
  const r = reportWith(MAIN, { drive: makeFakeDrive() });
  answer403For(r, /SA4%23138\.xlsx$/);
  const run = attempt(r, () => r.s.updateReportNow());
  check('the list cannot be downloaded: "not updated", both addresses named, nothing changed', [/^The report was not updated: Could not download TDOC list: HTTP 403 \(https:\/\/ftp\.3gpp\.org\/.*\); HTTP 403 \(https:\/\/www\.3gpp\.org\/ftp\//.test(run.thrown || ''),
    /in part/.test(run.thrown || ''), run.changed, run.touched], [true, false, false, false]);
  const later = reportWith(MAIN, { drive: makeFakeDrive() });
  later.s.buildTdocTableIndex_ = () => { throw new Error('synthetic failure while inserting'); };  // after the download, before the TDocs are written
  const l = attempt(later, () => later.s.updateReportNow());
  check('a failure after the download still says "may have been updated in part"', [/^The report update failed: synthetic failure while inserting\nThe report may have been updated in part\./.test(l.thrown || '')], [true]);
}

// ====================================================================== H
console.log('H. what changed in Code.js');
{
  const view = codeWithoutMainMeetingNumber(FULL);
  check('every change is in Code.js exactly once', view.problems, []);
  // The session times of issue #11 (T-2026.10.25) came later; the functions are compared without them (tests/session-times-main-report.test.js).
  const FULL_T = codeWithoutSessionTimes(FULL).code;
  const functionNames = (src) => (src.match(/^function\s+[A-Za-z0-9_$]+\s*\(/gm) || []).map((x) => x.replace(/^function\s+|\s*\($/g, '')).filter((n, i, all) => all.indexOf(n) === i);
  const functionSource = (src, name) => { const start = src.indexOf('\nfunction ' + name + '('); return start === -1 ? null : src.slice(start + 1, src.indexOf('\n}\n', start) + 2); };
  const section = FULL.slice(FULL.indexOf(MAIN_MEETING_NUMBER_SECTION_START), FULL.indexOf(MAIN_MEETING_NUMBER_SECTION_END) + MAIN_MEETING_NUMBER_SECTION_END.length);
  check('the section: these functions, nothing else at top level', [functionNames(section), section.match(/^(?:const|let|var) [A-Za-z0-9_$]+/gm)], [MAIN_MEETING_NUMBER_SECTION_FUNCTIONS, null]);
  check('it stands directly after the meeting time zones', FULL.indexOf('// END OF MEETING TIME ZONES\n// =========================================================\n\n' + MAIN_MEETING_NUMBER_SECTION_START) !== -1, true);
  check('of the functions there were, these changed; none removed', [functionNames(view.code).filter((n) => functionSource(FULL_T, n) !== functionSource(view.code, n)).sort(), functionNames(view.code).filter((n) => functionNames(FULL_T).indexOf(n) === -1)],
    [MAIN_MEETING_NUMBER_CHANGED_FUNCTIONS.slice().sort(), []]);
  check('Code.js is version 2.26.9 (2.26.9 the session times in a main report; released as T-2026.10.24 with the meeting time zones; the ledger takes its changelog entry out)', (FULL.match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], '2.26.9');
  let candidate = null;
  let from = '';
  // 3dd4263: the release commit of the meeting time zones, the candidate the fix was made on (8a47910 changed no Code.js).
  for (const ref of ['3dd4263']) {
    try { candidate = execFileSync('git', ['show', ref + ':Code.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').replace(/\r/g, ''); from = ref; break; } catch (e) { candidate = null; }
  }
  if (candidate === null) console.log('  note: commit 3dd4263 is not in this checkout; the comparison is skipped.');
  else check('without it, Code.js is that of ' + from + ', byte for byte', view.code === candidate, true);
  check('ReportCreator.js, HyperLink.js and appsscript.json are not touched by it (source)', ['template/ReportCreator.js', 'HyperLink.js'].map((f) => /mainMeeting(FolderOf|NumberOf|IdentityOfSave|NumberMismatch)/.test(fs.readFileSync(path.join(ROOT, f), 'utf8'))), [false, false]);
}

console.log(failures === 0 ? '\nALL PASSED' : '\n' + failures + ' FAILED');
process.exit(failures === 0 ? 0 : 1);
