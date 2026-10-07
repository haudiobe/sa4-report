/**
 * Existing-report migration -- the adoption mode of tools/template-release.js.
 *
 * Adoption prepares the UNCHANGED payload of an already built release for
 * another bound script project. These checks cover the wrong-target
 * protection and the "nothing is rebuilt" guarantee. Pure planAdoption():
 * no git, no files written, no clasp.
 *
 * Run: node tests/template-adopt.test.js
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const {
  planRelease, planAdoption, adoptionTargetRefusal, adoptionBanner, RELEASE_FILES,
  PROTECTED_SCRIPT_ID_DIGESTS, PROTECTED_PROJECT_LABELS, protectionForScriptIds, protectedProjectKey
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
const sha = (c) => crypto.createHash('sha256').update(c).digest('hex');
const COMMIT = '962fb7d' + '0'.repeat(33);
const RELEASE_ID = 'T-2026.10.4';
const TEMPLATE_SCRIPT = 'TEMPLATEscript00000000000000000000000000000000000000000';
const REPORT_SCRIPT = 'REPORTscript0000000000000000000000000000000000000000000';
// The tool lists the protected projects by the SHA-256 digest of their
// Script ID; the real IDs are not in this repository. These tests use
// synthetic IDs and a digest table made from them.
const LEGACY_SCRIPT_ID = 'LEGACYscript0000000000000000000000000000000000000000000';
const CENTRAL_SCRIPT_ID = 'CENTRALscript000000000000000000000000000000000000000000';
const PROTECTION = protectionForScriptIds({ legacy: LEGACY_SCRIPT_ID, central: CENTRAL_SCRIPT_ID });
// The files of a release built from this checkout. The files of an older
// release are those of its own tag: tests/release-file-set.test.js.
const FIVE = ['appsscript.json', 'Code.js', 'HyperLink.js', 'ReportCreator.js', 'Release.js'];

/** A release bundle exactly as the normal release mode writes it (in memory). */
function builtBundle() {
  const p = planRelease({
    releaseId: RELEASE_ID,
    target: { templateDocumentId: 'TEMPLATEdoc0000000000000000000000000000000', templateScriptId: TEMPLATE_SCRIPT },
    git: { commit: COMMIT, dirty: false, tagsAtHead: ['template-release/' + RELEASE_ID] },
    readFile: readRepo,
    builtAt: '2026-10-01T12:01:57.805Z'
  });
  if (!p.ok) throw new Error('fixture release refused: ' + p.errors.join('; '));
  const files = {};
  p.files.forEach((f) => { files[f.name] = Buffer.from(f.content); });
  return { files, manifest: JSON.parse(JSON.stringify(p.manifest)) };
}
const GIT = { tagCommit: COMMIT, showAtTag: (rel) => readRepo(rel) };
const adopt = (over) => planAdoption(Object.assign({
  releaseId: RELEASE_ID, targetScriptId: REPORT_SCRIPT, label: 'rehearsal-6g', allowLegacyTarget: false, bundle: builtBundle(), git: GIT, protection: PROTECTION
}, over || {}));

console.log('an adoption of a built release for another script is planned');
{
  const bundle = builtBundle();
  const p = adopt({ bundle });
  check('ok', [p.ok, p.errors], [true, []]);
  check('exactly the five release files, in release order', p.files.map((f) => f.name), FIVE);
  check('every file is the bundle\'s file, byte for byte', p.files.map((f) => Buffer.compare(Buffer.from(f.content), bundle.files[f.name])), [0, 0, 0, 0, 0]);
  check('Release.js is not regenerated: same bytes, same build time',
    [p.files[4].name, sha(p.files[4].content) === bundle.manifest.files[4].sha256, /"builtAt": "2026-10-01T12:01:57\.805Z"/.test(String(p.files[4].content))], ['Release.js', true, true]);
  check('the report will identify itself as the release', [p.release.releaseId, p.release.codeVersion, p.release.gitCommit, p.release.flavor],
    [RELEASE_ID, (String(readRepo('Code.js')).match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1], COMMIT, 'template']);
  check('Release.js still names the MASTER template, so the adopted document is a report', p.release.templateScriptId, TEMPLATE_SCRIPT);
  check('.clasp.json targets the given script only', JSON.parse(p.claspJson), { scriptId: REPORT_SCRIPT, rootDir: '.' });
  check('.claspignore uploads the five files and nothing else (not .clasp.json, not the adoption record)',
    p.claspIgnore.trim().split('\n'), ['**/**'].concat(FIVE.map((n) => '!' + n)));
  check('the adoption record names action, release and target', [p.record.action, p.record.releaseId, p.record.targetScriptId, p.record.label, p.record.legacyTarget],
    ['adopt', RELEASE_ID, REPORT_SCRIPT, 'rehearsal-6g', false]);
}

console.log('wrong-target protection');
{
  check('no target: refused', adopt({ targetScriptId: null }).errors, ['--target-script is required: adoption never guesses a target.']);
  check('empty target: refused', adopt({ targetScriptId: '' }).ok, false);
  check('malformed target: refused', adopt({ targetScriptId: 'not a script id' }).errors, ['--target-script is not a Script ID.']);
  check('too short to be a Script ID: refused', adoptionTargetRefusal('short', TEMPLATE_SCRIPT, false), '--target-script is not a Script ID.');

  const central = adopt({ targetScriptId: CENTRAL_SCRIPT_ID });
  check('CENTRAL: refused', [central.ok, central.errors.length, /CENTRAL editor add-on -- refusing/.test(central.errors[0])], [false, 1, true]);
  check('CENTRAL: refused even with the Legacy flag', adopt({ targetScriptId: CENTRAL_SCRIPT_ID, allowLegacyTarget: true }).ok, false);

  const legacy = adopt({ targetScriptId: LEGACY_SCRIPT_ID });
  check('Legacy: refused by default', [legacy.ok, legacy.errors.length, /Legacy bound script .* refusing.*--allow-legacy-target/.test(legacy.errors[0])], [false, 1, true]);
  const legacyAllowed = adopt({ targetScriptId: LEGACY_SCRIPT_ID, allowLegacyTarget: true });
  check('Legacy: possible only with the explicit flag, and marked', [legacyAllowed.ok, legacyAllowed.legacyTarget, legacyAllowed.record.legacyTarget], [true, true, true]);

  const master = adopt({ targetScriptId: TEMPLATE_SCRIPT });
  check('the master template\'s own script: refused in adoption mode', [master.ok, /master template's own script -- refusing/.test(master.errors[0])], [false, true]);
  check('the master template: refused with the Legacy flag too', adopt({ targetScriptId: TEMPLATE_SCRIPT, allowLegacyTarget: true }).ok, false);

  check('the built-in protection names the same two projects, by SHA-256 digest only',
    [Object.keys(PROTECTED_SCRIPT_ID_DIGESTS).sort(), Object.keys(PROTECTED_PROJECT_LABELS).sort(), Object.keys(PROTECTED_SCRIPT_ID_DIGESTS).every((k) => /^[0-9a-f]{64}$/.test(PROTECTED_SCRIPT_ID_DIGESTS[k]))],
    [['central', 'legacy'], ['central', 'legacy'], true]);
  check('the synthetic IDs are protected only through the test table, not by the built-in one',
    [protectedProjectKey(LEGACY_SCRIPT_ID), protectedProjectKey(CENTRAL_SCRIPT_ID), protectedProjectKey(LEGACY_SCRIPT_ID, PROTECTION), protectedProjectKey(CENTRAL_SCRIPT_ID, PROTECTION)],
    [null, null, 'legacy', 'central']);
  check('adoptionTargetRefusal: an ordinary report script is accepted', adoptionTargetRefusal(REPORT_SCRIPT, TEMPLATE_SCRIPT, false), '');
  check('the flag allows nothing but the Legacy script', [adoptionTargetRefusal(REPORT_SCRIPT, TEMPLATE_SCRIPT, true), adoptionTargetRefusal(CENTRAL_SCRIPT_ID, TEMPLATE_SCRIPT, true, PROTECTION) !== ''], ['', true]);
}

console.log('the target is printed before any deployment command');
{
  const lines = adoptionBanner(adopt(), 'C:\\out').join('\n');
  check('action, release identity and target Script ID are shown',
    [/ACTION : ADOPT an existing report/.test(lines), lines.indexOf('TARGET : Script ID ' + REPORT_SCRIPT) !== -1, /T-2026\.10\.4 \(Code\.js \d+\.\d+\.\d+, commit 962fb7d\)/.test(lines), /NOT the master template and NOT a new release/.test(lines)],
    [true, true, true, true]);
  check('the target appears before the clasp command', lines.indexOf('TARGET') < lines.indexOf('clasp push'), true);
  check('the command is a plain push; no force option is offered anywhere', [/clasp push$/m.test(lines), /--force/.test(lines)], [true, false]);
  check('a manifest overwrite is called out as a separate decision', /manifest, stop: that needs a separate decision/.test(lines), true);
  check('an ordinary target carries no Legacy warning', /LEGACY/.test(lines), false);
  const legacyLines = adoptionBanner(adopt({ targetScriptId: LEGACY_SCRIPT_ID, allowLegacyTarget: true }), 'C:\\out').join('\n');
  check('the Legacy target is announced as such', /THE TARGET IS THE LIVE LEGACY BOUND SCRIPT/.test(legacyLines), true);
  const source = fs.readFileSync(path.join(ROOT, 'tools', 'template-release.js'), 'utf8');
  check('the tool has no force option and never runs clasp', [/argv\.includes\('--force'\)|'--force'/.test(source), /execFileSync\(\s*'clasp/.test(source)], [false, false]);
}

console.log('nothing is rebuilt: the bundle must be the tagged release');
{
  check('no bundle: refused (adoption never rebuilds a release)', adopt({ bundle: null }).errors,
    ['The release bundle for T-2026.10.4 was not found. Adoption copies an existing bundle; it never rebuilds a release.']);
  check('no tag: refused', adopt({ git: { tagCommit: null, showAtTag: null } }).errors, ['Tag template-release/T-2026.10.4 does not exist in this repository.']);
  check('malformed release id: refused', adopt({ releaseId: '2.17.4' }).ok, false);
  check('missing label: refused', adopt({ label: null }).ok, false);
  check('a label that could leave the output folder: refused', adopt({ label: '../x' }).ok, false);

  const edited = builtBundle();
  edited.files['Code.js'] = Buffer.concat([edited.files['Code.js'], Buffer.from('\n// edited')]);
  check('a file changed after the build: refused', adopt({ bundle: edited }).errors.some((e) => /Code\.js in the bundle does not match its manifest hash/.test(e)), true);

  const reReleased = builtBundle();
  reReleased.files['Release.js'] = Buffer.from(String(reReleased.files['Release.js']).replace('T-2026.10.4', 'T-2026.10.5'));
  reReleased.manifest.files.find((f) => f.name === 'Release.js').sha256 = sha(reReleased.files['Release.js']);
  check('a Release.js with another release id: refused', adopt({ bundle: reReleased }).errors.some((e) => /Release\.js says T-2026\.10\.5, not T-2026\.10\.4/.test(e)), true);

  check('a bundle built from another commit than the tag: refused',
    adopt({ git: { tagCommit: 'a'.repeat(40), showAtTag: GIT.showAtTag } }).errors.some((e) => /Release\.js was built from .* but template-release\/T-2026\.10\.4 is a{40}/.test(e)), true);

  const otherSource = (rel) => (rel === 'template/ReportCreator.js' ? Buffer.from('// something else') : readRepo(rel));
  check('a source file that is not what the tag holds: refused', adopt({ git: { tagCommit: COMMIT, showAtTag: otherSource } }).errors,
    ['ReportCreator.js in the bundle differs from template/ReportCreator.js at template-release/T-2026.10.4.']);
  const lfOnly = (rel) => Buffer.from(String(readRepo(rel)).replace(/\r\n/g, '\n'));
  check('line endings alone are not a difference (git stores LF)', adopt({ git: { tagCommit: COMMIT, showAtTag: lfOnly } }).ok, true);

  const extra = builtBundle();
  extra.files['Template001Probe.gs'] = Buffer.from('x');
  check('an extra file in the bundle: refused', adopt({ bundle: extra }).errors, ['The bundle contains an unexpected file: Template001Probe.gs.']);
  const missing = builtBundle();
  delete missing.files['HyperLink.js'];
  check('a missing file in the bundle: refused', adopt({ bundle: missing }).errors, ['The bundle has no HyperLink.js.']);
  check('the release whitelist is four sources plus Release.js', RELEASE_FILES.map((f) => f[1]).concat(['Release.js']), FIVE);
}

console.log('the normal release mode is unchanged');
{
  const base = { releaseId: RELEASE_ID, git: { commit: COMMIT, dirty: false, tagsAtHead: ['template-release/' + RELEASE_ID] }, readFile: readRepo, builtAt: 'x', protection: PROTECTION };
  const target = (id) => ({ templateDocumentId: 'TEMPLATEdoc0000000000000000000000000000000', templateScriptId: id });
  check('it still refuses Legacy and CENTRAL as the template target',
    [LEGACY_SCRIPT_ID, CENTRAL_SCRIPT_ID].map((id) => planRelease(Object.assign({ target: target(id) }, base)).ok), [false, false]);
  check('and still accepts the template script', planRelease(Object.assign({ target: target(TEMPLATE_SCRIPT) }, base)).ok, true);
}

console.log('the built-in protection covers the real projects (checked only where their local clasp files exist)');
{
  const scriptIdOf = (p) => { try { return JSON.parse(fs.readFileSync(p, 'utf8').replace(/^\uFEFF/, '')).scriptId; } catch (e) { return null; } };
  const ownId = scriptIdOf(path.join(ROOT, '.clasp.json'));
  if (ownId) {
    check('this checkout\'s own .clasp.json target is refused without the flag',
      [!!protectedProjectKey(ownId), adoptionTargetRefusal(ownId, TEMPLATE_SCRIPT, false) !== ''], [true, true]);
  }
  const centralId = scriptIdOf(path.join(ROOT, '..', 'sa4-report-central-addon', '.clasp.json'));
  if (centralId) {
    check('the CENTRAL push folder\'s target is refused, with the Legacy flag too',
      [protectedProjectKey(centralId), adoptionTargetRefusal(centralId, TEMPLATE_SCRIPT, true) !== ''], ['central', true]);
  }
  if (!ownId && !centralId) console.log('  note: no local clasp files in this checkout; the built-in digests are not exercised here.');
}

console.log(failures ? `\n${failures} FAILURE(S)` : '\nall ok');
process.exitCode = failures ? 1 : 0;
