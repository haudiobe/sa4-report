/**
 * Issue #28: the safeguards of tools/integration-test/push-issue28-copy.mjs (pure rules in issue28-copy-rules.js).
 * Identity by digest (the template, the five SA4#138 reports, Legacy and CENTRAL refused), the pinned original,
 * compare-and-swap, pinned stage commits, the stage files (base and original never change report code).
 */
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const R = require('../tools/integration-test/issue28-copy-rules.js');

let failures = 0;
function check(name, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) failures++;
  console.log((ok ? '  ok   ' : '  FAIL ') + name + (ok ? '' : '\n         expected ' + JSON.stringify(expected) + '\n         actual   ' + JSON.stringify(actual)));
}
const ROOT = path.join(__dirname, '..');

// A synthetic target (the real IDs are not written in this public repository).
const DOC = '1TestDocument0000000000000000000000000000000';
const SCRIPT = '1TestScript00000000000000000000000000000000000000000000000';
const files0 = [
  { name: 'appsscript', type: 'JSON', source: '{"timeZone":"Europe/Berlin"}' },
  { name: 'Code', type: 'SERVER_JS', source: '// code v1\n' },
  { name: 'HyperLink', type: 'SERVER_JS', source: '// hl\n' },
  { name: 'Release', type: 'SERVER_JS', source: 'var SA4_RELEASE_ = {"releaseId":"T-2026.10.33","codeVersion":"2.27.1"};\n' },
  { name: 'ReportCreator', type: 'SERVER_JS', source: '// rc\n' }
];
R.TARGETS.test = { documentDigest: R.idDigest(DOC), scriptDigest: R.idDigest(SCRIPT), original: Object.fromEntries(files0.map((f) => [f.name, R.sourceSha(f.source)])) };
const project = (o) => Object.assign({ scriptId: SCRIPT, parentId: DOC }, o);
const file = (o) => Object.assign({ id: DOC, name: 'DISPOSABLE Issue28 - test', trashed: false }, o);

console.log('identity');
check('the right script bound to the right DISPOSABLE copy: allowed', R.identityRefusals('test', SCRIPT, project(), file()), []);
check('another script: refused', R.identityRefusals('test', '1Other', project({ scriptId: '1Other' }), file()), ['the script is not the bound script of the test copy']);
check('bound to another document: refused', R.identityRefusals('test', SCRIPT, project({ parentId: '1OtherDoc' }), file({ id: '1OtherDoc' })), ['the script is not bound to the test copy']);
check('not named DISPOSABLE, or in the trash: refused', [R.identityRefusals('test', SCRIPT, project(), file({ name: '6G Media report' })), R.identityRefusals('test', SCRIPT, project(), file({ trashed: true }))],
  [['the document name does not start with "DISPOSABLE"'], ['the document is in the trash']]);
check('the API returned another project, or Drive another file: refused', [R.identityRefusals('test', SCRIPT, project({ scriptId: 'x' }), file()), R.identityRefusals('test', SCRIPT, project(), file({ id: 'y' }))],
  [['the Apps Script API did not return this project'], ['Drive did not return the parent document']]);
// The protected tables by digest: a digest of the table refuses even when the target expects it (belt and braces).
const fakeProd = '1FakeProductionScript';
R.REFUSED_SCRIPT_DIGESTS[R.idDigest(fakeProd)] = 'Video SA4#138 (test)';
R.REFUSED_DOCUMENT_DIGESTS[R.idDigest('1FakeProdDoc')] = 'MBS SA4#138 (test)';
check('a protected script or parent (by digest): refused by name', R.identityRefusals('test', fakeProd, project({ scriptId: fakeProd, parentId: '1FakeProdDoc' }), file({ id: '1FakeProdDoc' })).filter((r) => /SA4#138/.test(r)),
  ['the script is Video SA4#138 (test)', 'the script is bound to MBS SA4#138 (test)']);
const clasp = path.join(ROOT, '.clasp.json');
if (fs.existsSync(clasp)) {
  const legacy = JSON.parse(fs.readFileSync(clasp, 'utf8').replace(/^﻿/, '')).scriptId;
  check('the Legacy bound script (repository .clasp.json; Legacy/CENTRAL digests of template-release.js): refused', R.identityRefusals('test', legacy, project({ scriptId: legacy }), file()).some((r) => /Legacy bound script/.test(r)), true);
}
check('the protected tables: 7 documents (template, five reports, Legacy 6G AH) and 6 scripts', [Object.keys(R.REFUSED_DOCUMENT_DIGESTS).length - 1, Object.keys(R.REFUSED_SCRIPT_DIGESTS).length - 1], [7, 6]);

console.log('the original, compare-and-swap');
check('the copy as first seen equals the pin: accepted; one byte else: not', [R.originalShaMap('test', files0).matchesPin, R.originalShaMap('test', files0.map((f) => (f.name === 'Code' ? Object.assign({}, f, { source: '// code v2\n' }) : f))).matchesPin], [true, false]);
check('... a file more or less: not', [R.originalShaMap('test', files0.slice(1)).matchesPin, R.originalShaMap('test', files0.concat([{ name: 'X', type: 'SERVER_JS', source: '' }])).matchesPin], [false, false]);
const exp = R.shaMap(files0);
check('no drift: []', R.drift(exp, R.shaMap(files0)), []);
const changed = R.drift(exp, R.shaMap(files0.map((f) => (f.name === 'Code' ? Object.assign({}, f, { source: 'edited in the editor' }) : f)).concat([{ name: 'Extra', type: 'SERVER_JS', source: '' }])));
check('an edit in the editor and a new file: both named', [changed.length, /^Code: /.test(changed[0]), changed[1]], [2, true, 'Extra: unexpected file']);
check('appsscript.json compared by content, not formatting; CRLF ignored', R.drift(exp, R.shaMap(files0.map((f) => (f.name === 'appsscript' ? Object.assign({}, f, { source: '{\n  "timeZone": "Europe/Berlin"\n}' }) : f.name === 'HyperLink' ? Object.assign({}, f, { source: '// hl\r\n' }) : f)))), []);

const manifest = (src) => [{ name: 'appsscript', type: 'JSON', source: src }];
check('appsscript.json re-saved by the editor with keys in another order (seen 2026-11-25): no drift',
  R.drift(R.shaMap(manifest('{"deps":[{"userSymbol":"Docs","serviceId":"docs","version":"v1"}],"timeZone":"Europe/Berlin"}')), R.shaMap(manifest('{"timeZone":"Europe/Berlin","deps":[{"userSymbol":"Docs","version":"v1","serviceId":"docs"}]}'))), []);
check('... a real change of the manifest (a scope added, an array reordered): drift',
  [R.drift(R.shaMap(manifest('{"oauthScopes":["a"]}')), R.shaMap(manifest('{"oauthScopes":["a","b"]}'))).length, R.drift(R.shaMap(manifest('{"s":["a","b"]}')), R.shaMap(manifest('{"s":["b","a"]}'))).length], [1, 1]);

console.log('stage files');
const CHECK = '// Issue28Check\n';
const T33_RELEASE = 'var SA4_RELEASE_ = {"releaseId":"T-2026.10.33","codeVersion":"2.27.1","gitCommit":"b7299c1"};\n';
const codeAt = (ref, p) => '// ' + p + ' at ' + ref + '\n';
const o = (x) => Object.assign({ original: files0, checkSource: CHECK, codeAt, t33Release: null, stamp: '2026-11-25T00:00:00Z', codeVersionOf: () => '9.9.9' }, x);
const base = R.stageFiles('base', o());
check('base: the original report files unchanged + Issue28Check', [base.map((f) => f.name), R.drift(R.shaMap(files0), R.shaMap(base.filter((f) => f.name !== 'Issue28Check')))], [files0.map((f) => f.name).concat(['Issue28Check']), []]);
check('original: exactly the original, Issue28Check removed', R.drift(R.shaMap(files0), R.shaMap(R.stageFiles('original', o({ original: base })))), []);
check('t33: refused with a Release.js that is not the production one', (() => { try { R.stageFiles('t33', o({ t33Release: T33_RELEASE })); return 'built'; } catch (e) { return /not the production one/.test(e.message); } })(), true);
check('the production T-2026.10.33 Release.js digest is the main copy\'s pinned original', R.TARGETS.main.original.Release, R.RELEASE_T33_SHA);
const a = (() => { try { return R.stageFiles('A', o({ t33Release: T33_RELEASE })); } catch (e) { return e.message; } })();
check('A: code from the PINNED commit, the copy\'s manifest, Release.js marked NOT RELEASED with the pin', [a.map((f) => f.name), a[1].source, /NOT RELEASED/.test(a[4].source) && a[4].source.includes(R.STAGE_PINS.A.commit), a[0] === files0[0]],
  [['appsscript', 'Code', 'HyperLink', 'ReportCreator', 'Release', 'Issue28Check'], '// Code.js at ' + R.STAGE_PINS.A.commit + '\n', true, true]);
const withJob = R.stageFiles('A', o({ t33Release: T33_RELEASE, jobSource: 'var IT28_JOB_ = {};' }));
check('a job goes with a stage as Issue28Job (after the harness); never with the original', [withJob.map((f) => f.name).slice(-2), (() => { try { R.stageFiles('original', o({ jobSource: 'x' })); return 'built'; } catch (e) { return /cannot go with the original/.test(e.message); } })(),
  R.stageFiles('base', o({ jobSource: 'var IT28_JOB_ = {};' })).map((f) => f.name).slice(-2), R.stageFiles('original', o({ original: files0.concat([{ name: 'Issue28Job', type: 'SERVER_JS', source: 'x' }]) })).map((f) => f.name).includes('Issue28Job')], [['Issue28Check', 'Issue28Job'], true, ['Issue28Check', 'Issue28Job'], false]);
check('a file the bundle does not know: refused', (() => { try { R.stageFiles('B', o({ t33Release: T33_RELEASE, original: files0.concat([{ name: 'Mine', type: 'SERVER_JS', source: '' }]) })); return 'built'; } catch (e) { return /does not know: Mine/.test(e.message); } })(), true);

console.log('pinned commits');
check('a branch still at its pinned code: allowed', R.pinRefusal('A', (ref, p) => '// ' + p), '');
check('a branch whose code moved: refused, files named', /has moved from its pinned commit 6f8c9df \(Code\.js differ\)/.test(R.pinRefusal('A', (ref, p) => (ref === R.STAGE_PINS.A.branch && p === 'Code.js' ? 'new' : '// ' + p))), true);
check('t33 has no branch: nothing to move', R.pinRefusal('t33', () => { throw new Error('not called'); }), '');
const git = (...x) => execFileSync('git', ['-C', ROOT, ...x], { encoding: 'utf8', maxBuffer: 1 << 28 });
const realAt = (ref, p) => git('show', ref + ':' + p);
const version = (c) => (realAt(c, 'Code.js').match(/^ \* Version: (\S+)/m) || [])[1];
check('the pins in this repository: T-2026.10.33 = 2.27.1, A = 2.28.0, B = 2.29.0, C = 2.30.0', ['t33', 'A', 'B', 'C'].map((s) => version(R.STAGE_PINS[s].commit)), ['2.27.1', '2.28.0', '2.29.0', '2.30.0']);
let branchesHere = true;
try { ['A', 'B', 'C'].forEach((s) => git('rev-parse', '--verify', '--quiet', R.STAGE_PINS[s].branch)); } catch (e) { branchesHere = false; }
if (branchesHere) check('the three phase branches are at their pinned code', ['A', 'B', 'C'].map((s) => R.pinRefusal(s, realAt)), ['', '', '']);

console.log('public repository');
const kit = ['tools/integration-test/push-issue28-copy.mjs', 'tools/integration-test/issue28-copy-rules.js', 'tools/integration-test/Issue28Check.js', 'tests/issue28-copy-rules.test.js',
  'tests/issue28-integration-check.test.js', 'docs/ISSUE28_INTEGRATION_TEST.md'].map((f) => fs.readFileSync(path.join(ROOT, f), 'utf8')).join('\n');
delete R.TARGETS.test; // the synthetic target of this test
check('no real ID of the copies, the protected reports, Legacy or CENTRAL anywhere in the kit (found by digest)', R.plainIdsIn(kit), []);
check('... and the scan finds one: a token with a protected digest is reported (by its digest, not its text)', (() => { R.REFUSED_SCRIPT_DIGESTS[R.idDigest('1FakeLeakedScriptId000000000000000000x7')] = 'leak (test)'; return R.plainIdsIn('x 1FakeLeakedScriptId000000000000000000x7 y'); })(), ['leak (test)']);
check('the harness pushed into the copies is pinned by digest; the branch still has it', R.sourceSha(git('show', R.STAGE_PINS.C.branch + ':tools/integration-test/Issue28Check.js')) === R.CHECK_SHA || R.sourceSha(fs.readFileSync(path.join(ROOT, 'tools', 'integration-test', 'Issue28Check.js'), 'utf8')) === R.CHECK_SHA, true);

if (failures) { console.log(`\n${failures} check(s) FAILED`); process.exit(1); }
console.log('\nAll checks passed.');
