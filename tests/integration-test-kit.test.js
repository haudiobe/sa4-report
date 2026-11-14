/**
 * The disposable-copy integration-test kit of issue #6 / PR #7
 * (tools/integration-test/): the push guard, the candidate Release.js, and
 * the read-only report summary of IntegrationCheck.js. The kit never reaches
 * a release bundle.
 *
 *   node tests/integration-test-kit.test.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { loadCode } = require('./helpers/load-code');
const { pushRefusals, audioOriginalRefusals, candidateReleaseJs, codeVersionOf, AUDIO_BASE_CODE_SHA256, COPY_BASE_CODE_SHA256S, DISPOSABLE_NAME_PREFIX } = require('../tools/integration-test/guard.js');
const { RELEASE_FILES } = require('../tools/template-release.js');

const ROOT = path.join(__dirname, '..');
let passed = 0;
let failed = 0;
function check(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a === e) { passed++; console.log('  ok   ' + name); } else { failed++; console.log('  FAIL ' + name + '\n       expected ' + e + '\n       actual   ' + a); }
}
function truthy(name, value) { check(name, !!value, true); }

const AUDIO_DOC = '1zawfIcbkzCJG-h2HYtnbQSEsm6MCd9OGhYzMgAFYENs';
const AUDIO_SCRIPT = '1TJAxpoSF2t6RS4FdO9TuCkHeM4eyaI7bDxcZk8v8GkaexZJLiOvKtTN1';
const TEMPLATE_DOC = '11XNUHU3TpAr8BaS7gyFQlwiKEzQFvnvzvGu4ms8lo2s';
const TEMPLATE_SCRIPT = '17xPMiY9PTft9KblONpetvZCA0Z7N_cwbG28qNMmYh3_QbPFDJhvtG8Pr';
const COPY_DOC = '1copyDocumentAAAAAAAAAAAAAAAAAAAAAAAAAAAAA';
const COPY_SCRIPT = '1copyScriptBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB';
const GOOD = { scriptId: COPY_SCRIPT, documentId: COPY_DOC, parentId: COPY_DOC, documentName: DISPOSABLE_NAME_PREFIX + ' (issue #6) - Audio', documentTrashed: false, baseCodeSha256: AUDIO_BASE_CODE_SHA256 };

console.log('A. push guard');
check('a disposable copy bound to its own script may be pushed', pushRefusals(GOOD), []);
truthy('the original Audio script is refused', pushRefusals(Object.assign({}, GOOD, { scriptId: AUDIO_SCRIPT })).some((r) => /original Audio/.test(r)));
truthy('the original Audio document is refused', pushRefusals(Object.assign({}, GOOD, { documentId: AUDIO_DOC, parentId: AUDIO_DOC })).some((r) => /original Audio/.test(r)));
truthy('the master template script is refused', pushRefusals(Object.assign({}, GOOD, { scriptId: TEMPLATE_SCRIPT })).some((r) => /Template/.test(r)));
truthy('the master template document is refused', pushRefusals(Object.assign({}, GOOD, { documentId: TEMPLATE_DOC, parentId: TEMPLATE_DOC })).some((r) => /Template/.test(r)));
truthy('a script bound to another document is refused', pushRefusals(Object.assign({}, GOOD, { parentId: '1someOtherReportCCCCCCCCCCCCCCCCCCCCCCCCCCC' })).some((r) => /not to --doc/.test(r)));
truthy('a script bound to the original, named as the copy, is refused', pushRefusals(Object.assign({}, GOOD, { parentId: AUDIO_DOC })).some((r) => /original Audio/.test(r)));
truthy('a standalone script is refused', pushRefusals(Object.assign({}, GOOD, { parentId: null })).some((r) => /not bound/.test(r)));
truthy('a document not named as a disposable copy is refused', pushRefusals(Object.assign({}, GOOD, { documentName: 'Copy of Audio SWG Minutes SA4#138' })).some((r) => /Rename the copy/.test(r)));
truthy('a trashed document is refused', pushRefusals(Object.assign({}, GOOD, { documentTrashed: true })).some((r) => /trash/.test(r)));
truthy('a script that does not run a release Audio ran is refused', pushRefusals(Object.assign({}, GOOD, { baseCodeSha256: 'ab'.repeat(32) })).some((r) => /not a release the Audio report ran/.test(r)));
check('... unless --allow-other-base', pushRefusals(Object.assign({}, GOOD, { baseCodeSha256: 'ab'.repeat(32), allowOtherBase: true })), []);
truthy('missing IDs are refused', pushRefusals({}).length === 2);
truthy('without --target audio-original the Audio report stays refused', pushRefusals(Object.assign({}, GOOD, { scriptId: AUDIO_SCRIPT, documentId: AUDIO_DOC, parentId: AUDIO_DOC, documentName: DISPOSABLE_NAME_PREFIX })).length > 0);

console.log('A2. --target audio-original (authorized 2026-11-12)');
const BACKUP_DOC = '1backupDocumentDDDDDDDDDDDDDDDDDDDDDDDDDDDD';
const BACKUP_SCRIPT = '1backupScriptEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEE';
const AGOOD = { scriptId: AUDIO_SCRIPT, documentId: AUDIO_DOC, parentId: AUDIO_DOC, originalModifiedTime: '2026-11-11T22:29:47.307Z', baseCodeSha256: AUDIO_BASE_CODE_SHA256, scriptMatchesLocalBackup: true,
  backup: { documentId: BACKUP_DOC, name: 'BACKUP before issue #6 test - Audio SWG Minutes SA4#138', mimeType: 'application/vnd.google-apps.document', trashed: false,
    createdTime: '2026-11-12T09:00:00.000Z', scriptParentId: BACKUP_DOC, scriptCodeSha256: AUDIO_BASE_CODE_SHA256 } };
const withBackup = (b) => Object.assign({}, AGOOD, { backup: Object.assign({}, AGOOD.backup, b) });
check('the Audio report with a verified backup may be pushed', audioOriginalRefusals(AGOOD), []);
truthy('the master template is refused', audioOriginalRefusals(Object.assign({}, AGOOD, { scriptId: TEMPLATE_SCRIPT, documentId: TEMPLATE_DOC, parentId: TEMPLATE_DOC })).some((r) => /permits only/.test(r)));
truthy('another report is refused', audioOriginalRefusals(Object.assign({}, AGOOD, { scriptId: COPY_SCRIPT, documentId: COPY_DOC, parentId: COPY_DOC })).some((r) => /permits only/.test(r)));
truthy('the Audio script with another document is refused', audioOriginalRefusals(Object.assign({}, AGOOD, { documentId: COPY_DOC })).some((r) => /permits only/.test(r)));
truthy('a script the API does not bind to Audio is refused', audioOriginalRefusals(Object.assign({}, AGOOD, { parentId: COPY_DOC })).some((r) => /parent/.test(r)));
truthy('no backup: refused', audioOriginalRefusals(Object.assign({}, AGOOD, { backup: null })).some((r) => /--backup-doc is required/.test(r)));
truthy('the original as its own backup: refused', audioOriginalRefusals(withBackup({ documentId: AUDIO_DOC })).some((r) => /original itself/.test(r)));
truthy('a backup older than the last change: refused', audioOriginalRefusals(withBackup({ createdTime: '2026-11-11T22:00:00.000Z' })).some((r) => /older than/.test(r)));
truthy('a backup not named BACKUP: refused', audioOriginalRefusals(withBackup({ name: 'Copy of Audio SWG Minutes SA4#138' })).some((r) => /must start with "BACKUP"/.test(r)));
truthy('a trashed backup: refused', audioOriginalRefusals(withBackup({ trashed: true })).some((r) => /trash/.test(r)));
truthy('a backup that is not a Doc: refused', audioOriginalRefusals(withBackup({ mimeType: 'application/pdf' })).some((r) => /not a Google Doc/.test(r)));
truthy('a backup script bound elsewhere: refused', audioOriginalRefusals(withBackup({ scriptParentId: AUDIO_DOC })).some((r) => /not bound to the backup/.test(r)));
truthy('a backup without the T-2026.10.22 script: refused', audioOriginalRefusals(withBackup({ scriptCodeSha256: 'cd'.repeat(32) })).some((r) => /not T-2026.10.22/.test(r)));
truthy('an Audio script changed since the local backup: refused', audioOriginalRefusals(Object.assign({}, AGOOD, { scriptMatchesLocalBackup: false })).some((r) => /differs from the local script backup/.test(r)));
truthy('an Audio script that is not T-2026.10.22: refused', audioOriginalRefusals(Object.assign({}, AGOOD, { baseCodeSha256: 'ab'.repeat(32) })).some((r) => /not T-2026.10.22/.test(r)));
check('... unless --allow-other-base (a second push over the candidate)', audioOriginalRefusals(Object.assign({}, AGOOD, { baseCodeSha256: 'ab'.repeat(32), allowOtherBase: true })), []);

console.log('B. candidate Release.js');
const audioRelease = '/**\n * GENERATED by tools/template-release.js -- do not edit.\n */\nvar SA4_RELEASE_ = ' + JSON.stringify({
  releaseId: 'T-2026.10.22', flavor: 'template', codeVersion: '2.26.6', gitCommit: 'ae1d1ca985111e702347ae17617687ef0a79d7c7',
  gitTag: 'template-release/T-2026.10.22', builtAt: '2026-11-11T20:24:56.324Z', templateDocumentId: TEMPLATE_DOC, templateScriptId: TEMPLATE_SCRIPT
}, null, 2) + ';\n';
const commit = 'e014f5d8f8db6b31c2d19e87511fb266f9b88b05';
const CAND = { codeVersion: '2.26.8', release: 'T-2026.10.24', branch: 'fix/bounded-update' };
const rel = candidateReleaseJs(audioRelease, commit, '2026-11-12T00:00:00.000Z', CAND);
const box = { console };
vm.createContext(box);
vm.runInContext(rel + '\n' + fs.readFileSync(path.join(ROOT, 'template', 'ReportCreator.js'), 'utf8') + '\n;this.__r = SA4_RELEASE_;', box);
check('the runtime accepts it (validateTemplateRelease_)', box.validateTemplateRelease_(box.__r), []);
check('it names the candidate commit', box.__r.gitCommit, commit);
check('it keeps the base release and the template IDs', [box.__r.releaseId, box.__r.templateDocumentId, box.__r.templateScriptId], ['T-2026.10.22', TEMPLATE_DOC, TEMPLATE_SCRIPT]);
truthy('it says it is not a release', /NOT RELEASED/.test(box.__r.codeVersion) && /NOT RELEASED/.test(box.__r.gitTag));
const roleBox = loadCode().sandbox;
check('the copy is a report, not the master (Code.js templateDocumentRole_)', [roleBox.templateDocumentRole_(COPY_DOC, box.__r), roleBox.templateDocumentRole_(TEMPLATE_DOC, box.__r)], ['report', 'template']);
let threw = '';
try { candidateReleaseJs(audioRelease, 'e014f5d', 'x', CAND); } catch (e) { threw = e.message; }
truthy('a short commit is refused', /full SHA-1/.test(threw));
// Issue #9: the release line of the diagnostics read "T-2026.10.23 / 2.26.7 + issue #6 candidate (NOT RELEASED) + issue #6 candidate (NOT RELEASED)".
check('it names the candidate, not issue #6: its Code.js version and release, its branch', [box.__r.codeVersion, box.__r.gitTag],
  ['2.26.8 (candidate of T-2026.10.24, NOT RELEASED)', 'candidate fix/bounded-update (NOT RELEASED; disposable-copy integration test)']);
check('it keeps the release the copy was made from beside it', [box.__r.baseReleaseId, box.__r.baseCodeVersion], ['T-2026.10.22', '2.26.6']);
const second = candidateReleaseJs(rel, 'f'.repeat(40), '2026-11-13T00:00:00.000Z', Object.assign({}, CAND, { branch: 'fix/other' }));
const box2 = { console };
vm.createContext(box2);
vm.runInContext(second + '\n;this.__r = SA4_RELEASE_;', box2);
check('a second push over a candidate writes the version anew -- nothing appended twice -- and keeps the base', [box2.__r.codeVersion, box2.__r.gitTag, box2.__r.baseReleaseId, box2.__r.baseCodeVersion, box2.__r.releaseId],
  ['2.26.8 (candidate of T-2026.10.24, NOT RELEASED)', 'candidate fix/other (NOT RELEASED; disposable-copy integration test)', 'T-2026.10.22', '2.26.6', 'T-2026.10.22']);
// The copy of issue #9 as it is: its Release.js written by the old guard (twice).
const oldTwice = '/** old */\nvar SA4_RELEASE_ = ' + JSON.stringify({ releaseId: 'T-2026.10.23', flavor: 'template', codeVersion: '2.26.7 + issue #6 candidate (NOT RELEASED) + issue #6 candidate (NOT RELEASED)',
  gitCommit: '8edb1a730582f873b53270cf9c7393bb3962ce03', gitTag: 'candidate fix/audio-large-build-checkpoint (NOT RELEASED; disposable-copy integration test)', builtAt: 'x', templateDocumentId: TEMPLATE_DOC, templateScriptId: TEMPLATE_SCRIPT }) + ';\n';
const box3 = { console };
vm.createContext(box3);
vm.runInContext(candidateReleaseJs(oldTwice, commit, 'y', CAND) + '\n;this.__r = SA4_RELEASE_;', box3);
check('over the copy of issue #9 (old Release.js, suffix twice): base T-2026.10.23 / 2.26.7, the suffix gone', [box3.__r.codeVersion, box3.__r.baseReleaseId, box3.__r.baseCodeVersion],
  ['2.26.8 (candidate of T-2026.10.24, NOT RELEASED)', 'T-2026.10.23', '2.26.7']);
['codeVersion', 'release', 'branch'].forEach((k) => {
  let refused = '';
  try { candidateReleaseJs(audioRelease, commit, 'x', Object.assign({}, CAND, { [k]: '' })); } catch (e) { refused = e.message; }
  truthy('a candidate without its ' + k + ' is refused', /candidate needs/.test(refused));
});
check('the Code.js version is read from the version line', [codeVersionOf(fs.readFileSync(path.join(ROOT, 'Code.js'), 'utf8')), codeVersionOf('no version')], ['2.26.8', '']);

console.log('C. IntegrationCheck.js summary (read-only)');
const { sandbox } = loadCode();
vm.runInContext(fs.readFileSync(path.join(ROOT, 'tools', 'integration-test', 'IntegrationCheck.js'), 'utf8'), sandbox, { filename: 'IntegrationCheck.js' });
const para = (text) => ({ paragraph: { elements: [{ startIndex: 1, endIndex: 1 + text.length + 1, textRun: { content: text + '\n' } }] } });
const cell = (text) => ({ content: [para(text)] });
const dropCell = (optionId, displayValue) => ({ content: [{ paragraph: { elements: [{ startIndex: 1, endIndex: 2, dropdown: { dropdownId: 'd' + optionId, dropdownProperties: { dropdownDefinitionId: 'def1', selectedOptionId: optionId, displayValue } } }, { textRun: { content: '\n' } }] } }] });
const row = (...cells) => ({ tableCells: cells });
const tdoc = (number, status, abstract) => ({ table: { tableRows: [row(cell('TDoc'), cell(number)), row(cell('Title'), cell('t ' + number)), row(cell('Status'), status)]
  .concat(abstract ? [row(cell('Abstract'), cell('a ' + number))] : []) } });
const OPTIONS = [{ optionId: 'o1', displayValue: 'agreed' }, { optionId: 'o2', displayValue: 'noted' }, { optionId: 'o3', displayValue: 'parked' }];
const doc = (content, withDefinition) => ({ revisionId: 'r1', tabs: [{ tabProperties: { tabId: 't.0' }, documentTab: {
  dropdownDefinitions: withDefinition ? { def1: { dropdownDefinitionProperties: { title: 'Document Status', options: OPTIONS } } } : {},
  body: { content: [para('Report')].concat(content) } } }] });

const built = doc([tdoc('S4-261001', dropCell('o1', 'agreed'), true), tdoc('S4-261002', cell('noted'), true), tdoc('S4-261003', cell(''), false),
  tdoc('S4-261004', cell('withdrawn'), false), { table: { tableRows: [row(cell('Name'), cell('Company'))] } }], true);
const s = sandbox.it6Summarize_(built, '1', null);
check('counts', [s.tables, s.tdocTableElements, s.tdocNumbers, s.abstracts, s.dropdowns, s.textStatuses, s.emptyStatuses], [5, 4, 4, 2, 1, 2, 1]);
check('a mappable text status is still to be converted; one without an option is not', [s.stillToConvert, s.leftAsTextNoOption], [1, 1]);
check('definition and mark', [s.definition, s.marked, s.duplicates], ['present (3 options)', true, []]);
check('per-TDoc statuses', s.statuses, { 'S4-261001': 'D:agreed', 'S4-261002': 'T:noted', 'S4-261003': 'E:', 'S4-261004': 'T:withdrawn' });

const doubled = doc([tdoc('S4-261001', cell('agreed'), true), tdoc('S4-261001', cell('agreed'), true)], false);
const d0 = sandbox.it6Summarize_(doubled, null, OPTIONS);
check('a doubled TDoc table is reported', [d0.duplicates, d0.tdocTableElements, d0.tdocNumbers, d0.definition, d0.marked], [['S4-261001'], 2, 1, 'none', false]);

// A status chosen by hand, then a rebuild that keeps it -- and one that does not.
const handChosen = doc([tdoc('S4-261001', dropCell('o3', 'parked'), true), tdoc('S4-261002', dropCell('o2', 'noted'), true)], true);
const kept = doc([tdoc('S4-261001', dropCell('o3', 'parked'), true), tdoc('S4-261002', dropCell('o2', 'noted'), true)], true);
const lost = doc([tdoc('S4-261001', dropCell('o1', 'agreed'), true), tdoc('S4-261002', dropCell('o2', 'noted'), true), tdoc('S4-261009', cell('agreed'), false)], true);
const a = sandbox.it6Summarize_(handChosen, '1', null);
check('nothing changed: empty diff', JSON.parse(JSON.stringify(sandbox.it6Diff_(a, sandbox.it6Summarize_(kept, '1', null)))), { tablesAdded: [], tablesRemoved: [], statusChanged: [], counts: {}, otherContentUnchanged: true });
const dl = JSON.parse(JSON.stringify(sandbox.it6Diff_(a, sandbox.it6Summarize_(lost, '1', null))));
check('a lost hand-chosen value and an added table are reported', [dl.statusChanged, dl.tablesAdded, dl.otherContentUnchanged], [['S4-261001: D:parked -> D:agreed'], ['S4-261009'], false]);
check('the Status value alone does not change the other-content digest', sandbox.it6Summarize_(doc([tdoc('S4-261001', cell('agreed'), true)], true), null, null).otherContentDigest,
  sandbox.it6Summarize_(doc([tdoc('S4-261001', dropCell('o1', 'agreed'), true)], true), null, null).otherContentDigest);
const minutes = doc([tdoc('S4-261001', dropCell('o1', 'agreed'), true), para('Minutes typed by hand')], true);
truthy('typed minutes change the other-content digest', sandbox.it6Summarize_(minutes, null, null).otherContentDigest !== sandbox.it6Summarize_(doc([tdoc('S4-261001', dropCell('o1', 'agreed'), true)], true), null, null).otherContentDigest);
check('the log line has no per-TDoc list', /statuses/.test(sandbox.it6Format_(s)), false);

console.log('D. the kit stays out of releases');
check('RELEASE_FILES never names the kit', RELEASE_FILES.filter((f) => /integration-test|IntegrationCheck/.test(JSON.stringify(f))), []);
const code = fs.readFileSync(path.join(ROOT, 'Code.js'), 'utf8') + fs.readFileSync(path.join(ROOT, 'template', 'ReportCreator.js'), 'utf8');
check('no runtime function or variable starts with it6/IT6 (no clash in the copy)', (code.match(/^(?:function|var|const|let)\s+(it6|IT6)\w*/gm) || []), []);
const kitSrc = fs.readFileSync(path.join(ROOT, 'tools', 'integration-test', 'IntegrationCheck.js'), 'utf8');
check('IntegrationCheck.js makes no DocumentApp change and calls no Docs batchUpdate', /batchUpdate|saveAndClose|\.(setText|appendParagraph|insertTable|clear|removeChild|appendTable)\(/.test(kitSrc), false);
truthy('IntegrationCheck.js refuses the master template', kitSrc.includes(TEMPLATE_DOC));

// ====================================================================== E (issue #9)
console.log('E. issue #9: a copy of Audio as it runs T-2026.10.23; the time-zone check');
{
  const T23 = '3f9e719eb43c3abb3fb0721b6b8fa6390abe8ee3ade0851fdd1575d3cbaf27a8';
  check('a fresh copy may start from T-2026.10.22 or T-2026.10.23', [pushRefusals(Object.assign({}, GOOD, { baseCodeSha256: AUDIO_BASE_CODE_SHA256 })), pushRefusals(Object.assign({}, GOOD, { baseCodeSha256: T23 }))], [[], []]);
  check('T-2026.10.23 is the Code.js of its release bundle', COPY_BASE_CODE_SHA256S[T23], 'T-2026.10.23');
  truthy('any other base is refused unless --allow-other-base', pushRefusals(Object.assign({}, GOOD, { baseCodeSha256: 'ab'.repeat(32) })).some((r) => /not a fresh copy/.test(r)) &&
    pushRefusals(Object.assign({}, GOOD, { baseCodeSha256: 'ab'.repeat(32), allowOtherBase: true })).length === 0);
  truthy('the Audio original stays refused in the copy mode', pushRefusals(Object.assign({}, GOOD, { scriptId: AUDIO_SCRIPT, documentId: AUDIO_DOC, parentId: AUDIO_DOC, baseCodeSha256: T23 })).some((r) => /original Audio/.test(r)));
  const check9 = fs.readFileSync(path.join(__dirname, '..', 'tools', 'integration-test', 'IntegrationCheck.js'), 'utf8');
  truthy('the check file refuses the template and the Audio original', /11XNUHU3TpAr8BaS7gyFQlwiKEzQFvnvzvGu4ms8lo2s/.test(check9.slice(0, check9.indexOf('IT6_SNAPSHOT_PREFIX_'))) && /1zawfIcbkzCJG-h2HYtnbQSEsm6MCd9OGhYzMgAFYENs/.test(check9.slice(0, check9.indexOf('IT6_SNAPSHOT_PREFIX_'))));

  const l = loadCode({});
  const vm = require('vm');
  vm.runInContext(check9, l.sandbox);
  const lines = l.sandbox.it9TimeZoneLines_({ release: 'T', stored: {}, sessionsZone: '', scriptZone: 'Europe/Berlin', rawCalgary: '-0700 MST', rawChange: '-0700 MST', paragraphs: [] });
  check('it9TimeZoneCheck: what it reports for Calgary (by the report rule, whatever the tz data says)', lines.filter((x) => /report rule|09:00 on|opening sentence|session, |uploads/.test(x)), [
    '[IT9] report rule, 2026-11-16 15:00 UTC: offset -360 min, written ABT',
    '[IT9] 09:00 on 2026-11-16 in Calgary is 2026-11-16T15:00:00.000Z',
    '[IT9] opening sentence: <Chair> opens the session on November 16, 2026 at 09:00 ABT.',
    '[IT9] session, one day: November 16, 2026, 09:00–17:30 ABT',
    '[IT9] session, over 1 November: October 30 – November 2, 2026 (from 09:00 MDT on the first day, until 17:00 ABT on the last day)',
    '[IT9] uploads (UTC -> Calgary): 2026-11-16T17:59:00 -> 2026-11-16T11:59:00, 2026-11-16T18:01:00 -> 2026-11-16T12:01:00, 2026-11-01T08:30:00 -> 2026-11-01T02:30:00 (ABT)']);
  check('the release line names a candidate as one, with the release of the copy beside it; a release as before; none', [
    l.sandbox.it9ReleaseLine_({ releaseId: 'T-2026.10.23', codeVersion: '2.26.8 (candidate of T-2026.10.24, NOT RELEASED)', gitCommit: 'abc', gitTag: 'candidate fix/bounded-update (NOT RELEASED; disposable-copy integration test)', baseReleaseId: 'T-2026.10.23', baseCodeVersion: '2.26.7' }),
    l.sandbox.it9ReleaseLine_({ releaseId: 'T-2026.10.24', codeVersion: '2.26.8', gitCommit: 'def' }),
    l.sandbox.it9ReleaseLine_(null)], [
    'candidate Code.js 2.26.8 (candidate of T-2026.10.24, NOT RELEASED) / abc / candidate fix/bounded-update (NOT RELEASED; disposable-copy integration test) -- on a copy of T-2026.10.23 (Code.js 2.26.7)',
    'T-2026.10.24 / 2.26.8 / def', '(no Release.js)']);
  truthy('it9TimeZoneCheck writes nothing (source)', !/setProperty|deleteProperty|setText|appendParagraph|insert|clear\(|saveAndClose/.test(check9.slice(check9.indexOf('// Issue #9'))));
}

console.log('\n' + passed + ' passed, ' + failed + ' failed');
if (failed) { console.log('SOME CHECKS FAILED'); process.exit(1); }
console.log('All integration-test kit checks passed.');
