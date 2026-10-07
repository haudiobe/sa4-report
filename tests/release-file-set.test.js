/**
 * The production file set, and that every release keeps its own.
 *
 * A release built from this checkout is five files: appsscript.json, Code.js,
 * HyperLink.js, ReportCreator.js and the generated Release.js. The releases
 * up to T-2026.10.11 are seven: they also carry colab_notebook.html and
 * colab_notebook_shared.html, two saved pages of the Google Colab web
 * application that no code ever loaded (docs/PRODUCTION_FILE_SET.md).
 *
 * The adoption mode of tools/template-release.js takes the file list of a
 * release from the release's own tag -- the RELEASE_FILES list of the tool as
 * the tag holds it -- and never from today's list. These checks cover both
 * generations:
 *   1. the current release: five files, in the plan, the manifest and the
 *      push filter;
 *   2. the list is read from a tool source, strictly, without executing it;
 *   3. a five-file release is adopted as five files;
 *   4. a seven-file release is adopted as seven files, with today's tool;
 *   5. an extra or a missing file is refused in both generations, and one
 *      generation's bundle is not accepted for the other's tag;
 *   6. no production file names the two pages any more;
 *   7. the real tags of this repository: T-2026.10.11 and the earlier ones
 *      declare seven files and still hold both pages; where the built bundle
 *      of T-2026.10.11 exists locally, it is still adoptable.
 *
 * 1-6 need no git. Without the tags in the checkout 7 is skipped, and said so.
 *
 * Run: node tests/release-file-set.test.js
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const {
  planRelease, planAdoption, buildReleaseJs, buildClaspIgnore, RELEASE_FILES,
  releaseFilesFromToolSource, taggedReleaseSources, RELEASE_TOOL_PATH, protectionForScriptIds
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
const FIVE = ['appsscript.json', 'Code.js', 'HyperLink.js', 'ReportCreator.js', 'Release.js'];
const SEVEN = ['appsscript.json', 'Code.js', 'HyperLink.js', 'colab_notebook.html', 'colab_notebook_shared.html', 'ReportCreator.js', 'Release.js'];
const COLAB = ['colab_notebook.html', 'colab_notebook_shared.html'];
const COMMIT = 'f1e5e70' + '0'.repeat(33);
const TEMPLATE_SCRIPT = 'TEMPLATEscript00000000000000000000000000000000000000000';
const TEMPLATE_DOC = 'TEMPLATEdoc0000000000000000000000000000000';
const REPORT_SCRIPT = 'REPORTscript0000000000000000000000000000000000000000000';
const PROTECTION = protectionForScriptIds({ legacy: 'LEGACYscript0000000000000000000000000000000000000000000', central: 'CENTRALscript000000000000000000000000000000000000000000' });

// A seven-file release as the tool of the time built it. Its tag holds the
// tool with the seven-file list and the two pages; neither is in this checkout.
const SEVEN_LIST_SOURCE = [
  '// Repository path -> published file name. Nothing else is ever bundled.',
  'const RELEASE_FILES = [',
  "  ['appsscript.json', 'appsscript.json'],",
  "  ['Code.js', 'Code.js'],",
  "  ['HyperLink.js', 'HyperLink.js'],",
  "  ['colab_notebook.html', 'colab_notebook.html'],",
  "  ['colab_notebook_shared.html', 'colab_notebook_shared.html'],",
  "  ['template/ReportCreator.js', 'ReportCreator.js']",
  '];',
  ''
].join('\n');
const OLD_TAG_FILES = {
  'colab_notebook.html': Buffer.from('<!DOCTYPE html><html><head><title>Google Colab</title></head></html>\n'),
  'colab_notebook_shared.html': Buffer.from('<!DOCTYPE html><html><head><title>Google Colab</title></head><body></body></html>\n')
};
OLD_TAG_FILES[RELEASE_TOOL_PATH] = Buffer.from(SEVEN_LIST_SOURCE);
const showAtOldTag = (rel) => (Object.prototype.hasOwnProperty.call(OLD_TAG_FILES, rel) ? OLD_TAG_FILES[rel] : readRepo(rel));
const showAtNewTag = (rel) => readRepo(rel);

/** A bundle as the release mode writes it, for the file list a tag declares (in memory). */
function bundleOf(releaseId, showAtTag, commit) {
  const tagged = taggedReleaseSources(showAtTag, 'template-release/' + releaseId);
  if (!tagged.ok) throw new Error('fixture: ' + tagged.errors.join('; '));
  const codeVersion = (String(showAtTag('Code.js')).slice(0, 400).match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1];
  const meta = { releaseId, flavor: 'template', codeVersion, gitCommit: commit || COMMIT, gitTag: 'template-release/' + releaseId, builtAt: '2026-10-07T10:00:00.000Z', templateDocumentId: TEMPLATE_DOC, templateScriptId: TEMPLATE_SCRIPT };
  const list = tagged.files.map((f) => ({ name: f.name, source: f.source, content: Buffer.from(f.content) }));
  list.push({ name: 'Release.js', source: '(generated)', content: Buffer.from(buildReleaseJs(meta)) });
  const files = {};
  list.forEach((f) => { files[f.name] = f.content; });
  return { files, manifest: { releaseId, codeVersion, gitCommit: meta.gitCommit, templateScriptId: TEMPLATE_SCRIPT, files: list.map((f) => ({ name: f.name, source: f.source, sha256: sha(f.content) })) } };
}
const adopt = (releaseId, showAtTag, over) => planAdoption(Object.assign({
  releaseId, targetScriptId: REPORT_SCRIPT, label: 'rollback', allowLegacyTarget: false,
  bundle: bundleOf(releaseId, showAtTag), git: { tagCommit: COMMIT, showAtTag }, protection: PROTECTION
}, over || {}));
const OLD = 'T-2026.10.11';
const NEW = 'T-2026.10.12';

console.log('1. the release built from this checkout is five files');
{
  const p = planRelease({
    releaseId: NEW, target: { templateDocumentId: TEMPLATE_DOC, templateScriptId: TEMPLATE_SCRIPT },
    git: { commit: COMMIT, dirty: false, tagsAtHead: ['template-release/' + NEW] }, readFile: readRepo, builtAt: '2026-10-07T10:00:00.000Z'
  });
  check('planned', [p.ok, p.errors], [true, []]);
  check('the plan has exactly the five production files', p.files.map((f) => f.name), FIVE);
  check('the manifest lists exactly those five', p.manifest.files.map((f) => f.name), FIVE);
  check('the push filter admits exactly those five', p.claspIgnore, '**/**\n!appsscript.json\n!Code.js\n!HyperLink.js\n!ReportCreator.js\n!Release.js\n');
  check('neither Colab page is in the plan, the manifest or the push filter', [p.files, p.manifest.files].map((list) => list.filter((f) => /colab/i.test(f.name + ' ' + f.source)).length).concat([/colab/i.test(p.claspIgnore)]), [0, 0, false]);
  check('RELEASE_FILES is the four sources', RELEASE_FILES, [['appsscript.json', 'appsscript.json'], ['Code.js', 'Code.js'], ['HyperLink.js', 'HyperLink.js'], ['template/ReportCreator.js', 'ReportCreator.js']]);
}

console.log('2. the file list of a release is read from the tool source its tag holds');
{
  check('this checkout\'s tool describes itself: the list read from its source is RELEASE_FILES', releaseFilesFromToolSource(readRepo(RELEASE_TOOL_PATH)), RELEASE_FILES);
  check('the same with CRLF line ends', releaseFilesFromToolSource(String(readRepo(RELEASE_TOOL_PATH)).replace(/\r\n/g, '\n').replace(/\n/g, '\r\n')), RELEASE_FILES);
  check('the seven-file tool gives its six sources, in order', releaseFilesFromToolSource(SEVEN_LIST_SOURCE).map((f) => f[1]), SEVEN.slice(0, 6));
  check('no list, an empty list, an expression in the list, a path out of the repository, a second Release.js, a name twice: not a list',
    ['var x = 1;\n', 'const RELEASE_FILES = [\n];\n', "const RELEASE_FILES = [\n  ['Code.js', 'Code.js'],\n  other()\n];\n", "const RELEASE_FILES = [\n  ['../x.js', 'x.js']\n];\n",
      "const RELEASE_FILES = [\n  ['a.js', 'Release.js']\n];\n", "const RELEASE_FILES = [\n  ['a.js', 'Code.js'],\n  ['b.js', 'Code.js']\n];\n", null].map(releaseFilesFromToolSource),
    [null, null, null, null, null, null, null]);
  const old = taggedReleaseSources(showAtOldTag, 'template-release/' + OLD);
  check('a seven-file tag reconstructs its six sources, both Colab pages among them, and names Release.js as the seventh',
    [old.ok, old.published, old.files.map((f) => f.name), COLAB.map((name) => String(old.files.find((f) => f.name === name).content) === String(OLD_TAG_FILES[name]))],
    [true, SEVEN, SEVEN.slice(0, 6), [true, true]]);
  const now = taggedReleaseSources(showAtNewTag, 'template-release/' + NEW);
  check('a five-file tag reconstructs its four sources', [now.ok, now.published, now.files.map((f) => f.source)], [true, FIVE, RELEASE_FILES.map((f) => f[0])]);
  check('a tag without a readable list gives no file set at all',
    [taggedReleaseSources((rel) => { if (rel === RELEASE_TOOL_PATH) throw new Error('x'); return readRepo(rel); }, 'template-release/' + OLD), taggedReleaseSources(() => Buffer.from('nothing'), 'template-release/' + OLD)].map((r) => [r.ok, r.published, r.files.length, r.errors]),
    [1, 2].map(() => [false, null, 0, ['Could not read the release file list of template-release/T-2026.10.11 (RELEASE_FILES in tools/template-release.js at the tag).']]));
  check('a file the tag lists but does not hold is reported', taggedReleaseSources((rel) => { if (rel === 'colab_notebook.html') throw new Error('x'); return showAtOldTag(rel); }, 'template-release/' + OLD).errors,
    ['Could not read colab_notebook.html at template-release/T-2026.10.11.']);
}

console.log('3. a five-file release is adopted as five files');
{
  const p = adopt(NEW, showAtNewTag);
  check('planned', [p.ok, p.errors], [true, []]);
  check('exactly the five files, in release order', [p.files.map((f) => f.name), p.record.files.map((f) => f.name)], [FIVE, FIVE]);
  check('the push filter admits exactly those five', p.claspIgnore, buildClaspIgnore(FIVE));
}

console.log('4. a seven-file release is still adopted as seven files');
{
  const bundle = bundleOf(OLD, showAtOldTag);
  const p = adopt(OLD, showAtOldTag, { bundle });
  check('planned, by a tool whose own list is five files', [p.ok, p.errors, RELEASE_FILES.length + 1], [true, [], 5]);
  check('exactly the seven files, in release order', [p.files.map((f) => f.name), p.record.files.map((f) => f.name)], [SEVEN, SEVEN]);
  check('both Colab pages are in the payload, byte for byte', COLAB.map((name) => Buffer.compare(Buffer.from(p.files.find((f) => f.name === name).content), OLD_TAG_FILES[name])), [0, 0]);
  check('the push filter admits exactly those seven', p.claspIgnore, buildClaspIgnore(SEVEN));
  check('Release.js is the bundle\'s, not a new one', sha(p.files[6].content), bundle.manifest.files[6].sha256);
  check('the record names the release', [p.record.action, p.record.releaseId], ['adopt', OLD]);
}

console.log('5. an unexpected or a missing file is refused, in both generations');
{
  [[NEW, showAtNewTag, FIVE], [OLD, showAtOldTag, SEVEN]].forEach(([releaseId, showAtTag, names]) => {
    const n = names.length + ' files';
    const extra = bundleOf(releaseId, showAtTag);
    extra.files['Template001Probe.gs'] = Buffer.from('x');
    check(n + ': an extra file in the bundle', adopt(releaseId, showAtTag, { bundle: extra }).errors, ['The bundle contains an unexpected file: Template001Probe.gs.']);
    const missing = bundleOf(releaseId, showAtTag);
    delete missing.files['HyperLink.js'];
    check(n + ': a file missing from the bundle', adopt(releaseId, showAtTag, { bundle: missing }).errors, ['The bundle has no HyperLink.js.']);
    const unlisted = bundleOf(releaseId, showAtTag);
    unlisted.manifest.files = unlisted.manifest.files.filter((f) => f.name !== 'HyperLink.js');
    check(n + ': a file missing from the manifest', adopt(releaseId, showAtTag, { bundle: unlisted }).errors,
      ['The bundle manifest does not list exactly the ' + names.length + ' files of template-release/' + releaseId + ' (' + names.join(', ') + ').']);
    const changed = (rel) => (rel === 'HyperLink.js' ? Buffer.from('// other') : showAtTag(rel));
    check(n + ': a source that is not the tag\'s', adopt(releaseId, showAtTag, { git: { tagCommit: COMMIT, showAtTag: changed } }).errors, ['HyperLink.js in the bundle differs from HyperLink.js at template-release/' + releaseId + '.']);
  });

  const withColab = bundleOf(NEW, showAtNewTag);
  withColab.files['colab_notebook.html'] = OLD_TAG_FILES['colab_notebook.html'];
  check('a Colab page in a five-file release: refused as unexpected', adopt(NEW, showAtNewTag, { bundle: withColab }).errors, ['The bundle contains an unexpected file: colab_notebook.html.']);

  const oneGone = bundleOf(OLD, showAtOldTag);
  delete oneGone.files['colab_notebook_shared.html'];
  check('a seven-file release without one of its Colab pages: refused as incomplete', adopt(OLD, showAtOldTag, { bundle: oneGone }).errors, ['The bundle has no colab_notebook_shared.html.']);
  const pageEdited = (rel) => (rel === 'colab_notebook.html' ? Buffer.from('<html>other</html>') : showAtOldTag(rel));
  check('a Colab page that is not the tag\'s: refused', adopt(OLD, showAtOldTag, { git: { tagCommit: COMMIT, showAtTag: pageEdited } }).errors,
    ['colab_notebook.html in the bundle differs from colab_notebook.html at template-release/T-2026.10.11.']);

  // The regression this file exists for: a seven-file release must not be
  // measured with today's five-file list, nor the other way round.
  const fiveForOld = bundleOf(OLD, showAtNewTag);
  const cut = adopt(OLD, showAtOldTag, { bundle: fiveForOld });
  check('a five-file bundle under a seven-file tag: refused, and both pages are named as missing',
    [cut.ok, cut.errors.filter((e) => /^The bundle has no colab_notebook/.test(e)).length, cut.errors.some((e) => /does not list exactly the 7 files/.test(e))], [false, 2, true]);
  const sevenForNew = bundleOf(NEW, showAtOldTag);
  const grown = adopt(NEW, showAtNewTag, { bundle: sevenForNew });
  check('a seven-file bundle under a five-file tag: refused, and both pages are named as unexpected',
    [grown.ok, grown.errors.filter((e) => /^The bundle contains an unexpected file: colab_notebook/.test(e)).length, grown.errors.some((e) => /does not list exactly the 5 files/.test(e))], [false, 2, true]);

  const noList = adopt(OLD, showAtOldTag, { git: { tagCommit: COMMIT, showAtTag: (rel) => (rel === RELEASE_TOOL_PATH ? Buffer.from('// no list') : showAtOldTag(rel)) } });
  check('a tag that does not say which files it has: nothing is planned, and today\'s list is not used instead',
    [noList.ok, noList.files.length, noList.claspIgnore, noList.errors],
    [false, 0, null, ['Could not read the release file list of template-release/T-2026.10.11 (RELEASE_FILES in tools/template-release.js at the tag).']]);
  const source = String(readRepo(RELEASE_TOOL_PATH)).replace(/\r/g, '');
  const adoptionSource = source.slice(source.indexOf('\nfunction planAdoption('), source.indexOf('\n}\n', source.indexOf('\nfunction planAdoption(')));
  check('planAdoption() does not read today\'s RELEASE_FILES in its code', /RELEASE_FILES/.test(adoptionSource.replace(/^\s*(\/\/|\*).*$/gm, '')), false);
}

console.log('6. no production file names the two pages');
{
  const production = ['appsscript.json', 'Code.js', 'HyperLink.js', 'template/ReportCreator.js', '.claspignore', 'tools/template-release.js'];
  check('neither page is in this checkout', COLAB.map((name) => fs.existsSync(path.join(ROOT, name))), [false, false]);
  check('no production source, push filter or release list names them',
    production.filter((rel) => /colab_notebook/.test(rel === 'tools/template-release.js' ? JSON.stringify(RELEASE_FILES) : String(readRepo(rel)))), []);
  check('no production source loads an HTML file by name', ['Code.js', 'HyperLink.js', 'template/ReportCreator.js'].filter((rel) => /FromFile\s*\(|createTemplate\s*\(/.test(String(readRepo(rel)))), []);
  const tracked = (() => { try { return execFileSync('git', ['ls-files', '-z'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }).toString('utf8').split('\0').filter(Boolean); } catch (e) { return null; } })();
  if (tracked) {
    const isText = (rel) => /\.(js|gs|json|md|html|txt|ps1|csv)$/.test(rel) || /(^|\/)\.[a-z]+$/.test(rel);
    const key = new RegExp('AIza' + '[0-9A-Za-z_-]{35}');
    check('no tracked file carries a Google browser key any more', tracked.filter((rel) => isText(rel) && fs.existsSync(path.join(ROOT, rel)) && key.test(String(readRepo(rel)))), []);
    check('no tracked file is an HTML file', tracked.filter((rel) => /\.html?$/i.test(rel) && fs.existsSync(path.join(ROOT, rel))), []);
  } else {
    console.log('  note: git is not available here; the tracked files are not searched.');
  }
}

console.log('7. the real release tags keep their own file set');
{
  const git = (args) => { try { return execFileSync('git', args, { cwd: ROOT, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'ignore'] }); } catch (e) { return null; } };
  const tagsOut = git(['tag', '-l', 'template-release/T-2026.10.*']);
  const tags = tagsOut ? tagsOut.toString('utf8').split(/\r?\n/).filter(Boolean) : [];
  const upTo11 = tags.filter((tag) => Number(tag.split('.').pop()) <= 11);
  if (!upTo11.length) {
    console.log('  note: the release tags are not available in this checkout; the comparisons with them are skipped.');
  } else {
    const show = (tag) => (rel) => { const out = git(['show', tag + ':' + rel]); if (out === null) throw new Error('not at the tag'); return out; };
    check('every release tag up to T-2026.10.11 declares the seven files and holds both Colab pages',
      upTo11.filter((tag) => { const t = taggedReleaseSources(show(tag), tag); return !(t.ok && JSON.stringify(t.published) === JSON.stringify(SEVEN) && COLAB.every((name) => /<title>Google Colab<\/title>/.test(String(t.files.find((f) => f.name === name).content)))); }), []);

    const tag = 'template-release/' + OLD;
    const tagCommitOut = git(['rev-parse', '--verify', '--quiet', tag + '^{commit}']);
    if (!tagCommitOut) {
      console.log('  note: ' + tag + ' is not available in this checkout; its checks are skipped.');
    } else {
      const tagCommit = tagCommitOut.toString('utf8').trim();
      const rebuilt = bundleOf(OLD, show(tag), tagCommit);
      const p = planAdoption({ releaseId: OLD, targetScriptId: REPORT_SCRIPT, label: 'rollback', bundle: rebuilt, git: { tagCommit, showAtTag: show(tag) }, protection: PROTECTION });
      check(OLD + ': its sources, taken from the tag alone, make a seven-file payload that is accepted', [p.ok, p.errors, p.files.map((f) => f.name)], [true, [], SEVEN]);
      check(OLD + ': the two pages of that payload are the tag\'s, about 93 KB each', COLAB.map((name) => p.files.find((f) => f.name === name).content.length > 90000), [true, true]);
      check(OLD + ': the tag holds Code.js 2.22.0', p.release.codeVersion, '2.22.0');

      // The bundle that was built and deployed, where this machine has it
      // (dist/ is not in Git): still accepted, file for file.
      const dir = path.join(ROOT, 'dist', 'template-release', OLD);
      if (fs.existsSync(path.join(dir, 'RELEASE-MANIFEST.json'))) {
        const files = {};
        fs.readdirSync(dir).forEach((name) => { if (['.clasp.json', '.claspignore', 'RELEASE-MANIFEST.json'].indexOf(name) === -1) files[name] = fs.readFileSync(path.join(dir, name)); });
        const built = planAdoption({ releaseId: OLD, targetScriptId: REPORT_SCRIPT, label: 'rollback', bundle: { files, manifest: JSON.parse(fs.readFileSync(path.join(dir, 'RELEASE-MANIFEST.json'), 'utf8')) }, git: { tagCommit, showAtTag: show(tag) }, protection: PROTECTION });
        check(OLD + ': the bundle built at the time is still adoptable, as seven files', [built.ok, built.errors, built.files.map((f) => f.name)], [true, [], SEVEN]);
      } else {
        console.log('  note: the built bundle of ' + OLD + ' is not on this machine; that check is skipped.');
      }
    }
  }
}

console.log(failures ? `\n${failures} FAILURE(S)` : '\nall ok');
process.exitCode = failures ? 1 : 0;
