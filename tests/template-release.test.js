/**
 * TEMPLATE prototype -- release bundle guardrails (tools/template-release.js).
 * Pure planRelease(): no git, no files written, no clasp.
 *
 * Run: node tests/template-release.test.js
 */

const fs = require('fs');
const path = require('path');
const {
  planRelease, validateTarget, RELEASE_FILES,
  PROTECTED_SCRIPT_ID_DIGESTS, PROTECTED_PROJECT_LABELS, protectionForScriptIds, protectedProject
} = require('../tools/template-release.js');

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

const ROOT = path.join(__dirname, '..');
const readRepo = (rel) => fs.readFileSync(path.join(ROOT, rel));
const COMMIT = 'd7c23bc' + '0'.repeat(33);
const TARGET = {
  templateDocumentId: 'TEMPLATEdoc0000000000000000000000000000000',
  templateScriptId: 'TEMPLATEscript00000000000000000000000000000000000000000'
};
const GIT_OK = { commit: COMMIT, dirty: false, tagsAtHead: ['template-release/T-2026.10.0'] };
const plan = (over) => planRelease(Object.assign({
  releaseId: 'T-2026.10.0', target: TARGET, git: GIT_OK, readFile: readRepo, builtAt: '2026-10-01T08:00:00.000Z'
}, over || {}));

console.log('a clean, tagged release is planned');
{
  const p = plan();
  check('ok', [p.ok, p.errors], [true, []]);
  check('exactly the whitelisted runtime files + Release.js',
    p.files.map((f) => f.name), ['appsscript.json', 'Code.js', 'HyperLink.js', 'colab_notebook.html', 'colab_notebook_shared.html', 'ReportCreator.js', 'Release.js']);
  check('Code.js is bundled byte-identical (the tested CENTRAL runtime)',
    p.files.find((f) => f.name === 'Code.js').content.equals(readRepo('Code.js')), true);
  check('no test file is ever bundled', p.files.some((f) => /test|helpers|fixtures/.test(f.source)), false);
  check('.clasp.json targets the template script only', JSON.parse(p.claspJson), { scriptId: TARGET.templateScriptId, rootDir: '.' });
  check('.claspignore whitelists exactly the bundle', p.claspIgnore.trim().split('\n'),
    ['**/**', '!appsscript.json', '!Code.js', '!HyperLink.js', '!colab_notebook.html', '!colab_notebook_shared.html', '!ReportCreator.js', '!Release.js']);
  const release = p.files.find((f) => f.name === 'Release.js').content;
  check('Release.js records release, commit, tag and template ids',
    [/"releaseId": "T-2026.10.0"/.test(release), release.indexOf(COMMIT) !== -1, /"gitTag": "template-release\/T-2026.10.0"/.test(release), release.indexOf(TARGET.templateDocumentId) !== -1],
    [true, true, true, true]);
  check('manifest hashes every file', p.manifest.files.every((f) => /^[0-9a-f]{64}$/.test(f.sha256)), true);
}

console.log('the Legacy and CENTRAL projects can never be a target');
{
  // The tool lists the protected projects by the SHA-256 digest of their
  // Script ID, so the real IDs are not in this repository. The refusal is
  // tested with synthetic IDs and a digest table made from them.
  const SYNTHETIC = { legacy: 'LEGACYscript0000000000000000000000000000000000000000000', central: 'CENTRALscript000000000000000000000000000000000000000000' };
  const protection = protectionForScriptIds(SYNTHETIC);
  Object.keys(SYNTHETIC).forEach((key) => {
    const p = plan({ target: Object.assign({}, TARGET, { templateScriptId: SYNTHETIC[key] }), protection });
    check('refused: ' + PROTECTED_PROJECT_LABELS[key], [p.ok, p.errors.some((e) => /refusing/.test(e))], [false, true]);
  });
  check('the built-in list holds two SHA-256 digests and no Script ID',
    [Object.keys(PROTECTED_SCRIPT_ID_DIGESTS).sort(), Object.keys(PROTECTED_SCRIPT_ID_DIGESTS).every((k) => /^[0-9a-f]{64}$/.test(PROTECTED_SCRIPT_ID_DIGESTS[k]))], [['central', 'legacy'], true]);
  check('an ordinary Script ID is not protected', [protectedProject(TARGET.templateScriptId), plan().ok], [null, true]);
  const localClasp = (() => { try { return JSON.parse(fs.readFileSync(path.join(ROOT, '.clasp.json'), 'utf8').replace(/^﻿/, '')).scriptId; } catch (e) { return null; } })();
  // The built-in digests against the real IDs, wherever the local clasp files exist.
  if (localClasp) check('this checkout\'s own .clasp.json target is protected', !!protectedProject(localClasp), true);
  const centralClasp = (() => { try { return JSON.parse(fs.readFileSync(path.join(ROOT, '..', 'sa4-report-central-addon', '.clasp.json'), 'utf8').replace(/^\uFEFF/, '')).scriptId; } catch (e) { return null; } })();
  if (centralClasp) check('the CENTRAL push folder\'s .clasp.json target is protected', protectedProject(centralClasp), PROTECTED_PROJECT_LABELS.central);
  check('missing target file refused', validateTarget(null).length > 0, true);
  check('placeholder ids from the example file refused',
    validateTarget(JSON.parse(fs.readFileSync(path.join(ROOT, 'template', 'template-target.example.json'), 'utf8'))).length, 2);
}

console.log('git state guards');
{
  check('dirty tree refused', plan({ git: Object.assign({}, GIT_OK, { dirty: true }) }).errors, ['The working tree is not clean -- commit or stash first.']);
  check('untagged HEAD refused', plan({ git: Object.assign({}, GIT_OK, { tagsAtHead: [] }) }).errors, ['HEAD is not tagged template-release/T-2026.10.0.']);
  check('malformed release id refused', plan({ releaseId: '2.15.2', git: Object.assign({}, GIT_OK, { tagsAtHead: ['template-release/2.15.2'] }) }).errors,
    ['--release must look like T-2026.10.0.']);
}

console.log('content guards');
{
  const withRequire = (rel) => (rel === 'template/ReportCreator.js' ? Buffer.from('var x = require("fs");') : readRepo(rel));
  check('Node-only require() in a bundled script refused', plan({ readFile: withRequire }).errors,
    ['ReportCreator.js calls require() -- Node-only code would break onOpen in Apps Script.']);
  const addonManifest = (rel) => (rel === 'appsscript.json' ? Buffer.from(JSON.stringify({ addOns: {} })) : readRepo(rel));
  check('an add-on manifest refused', plan({ readFile: addonManifest }).errors,
    ['appsscript.json declares addOns -- the template runtime is a bound script, not an add-on.']);
  check('a missing runtime file refused', plan({ readFile: (rel) => { if (rel === 'HyperLink.js') throw new Error('x'); return readRepo(rel); } }).errors,
    ['Missing release file HyperLink.js.']);
  check('the whitelist contains no directory outside template/ and the root', RELEASE_FILES.every(([src]) => !/\//.test(src) || src.indexOf('template/') === 0), true);
}

console.log('CENTRAL/Legacy push payloads cannot pick up template files');
{
  const ignore = fs.readFileSync(path.join(ROOT, '.claspignore'), 'utf8').replace(/^﻿/, '');
  check('repo .claspignore whitelists named files only, none under template/', [/^\*\*\/\*\*$/m.test(ignore), /template/.test(ignore)], [true, false]);
}

console.log(failures ? `\n${failures} FAILURE(S)` : '\nall ok');
process.exitCode = failures ? 1 : 0;
