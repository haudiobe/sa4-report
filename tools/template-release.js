#!/usr/bin/env node
/**
 * SA4 Report Template -- release bundle builder.
 * See docs/SA4_REPORT_TEMPLATE_ARCHITECTURE.md §10 and
 * docs/TEMPLATE-002B_RELEASE_CANDIDATE.md.
 *
 * Assembles the exact files for the master template's bound script into
 * dist/template-release/<releaseId>/ together with a generated Release.js,
 * a .clasp.json for the TEMPLATE script only, a whitelisting .claspignore
 * and RELEASE-MANIFEST.json (sha256 per file).
 *
 * It never runs clasp and never talks to Google. It prints the one command
 * to run from the bundle directory after review.
 *
 *   node tools/template-release.js --release T-2026.10.0            (dry run)
 *   node tools/template-release.js --release T-2026.10.0 --write    (bundle)
 *
 * Adoption mode (existing-report migration): prepares the UNCHANGED payload
 * of a release that was already built, for ANOTHER bound script project --
 * an existing report that is to run the template runtime. Nothing is
 * rebuilt: the release's files are copied byte for byte from the release
 * bundle (Release.js included, so the report identifies itself as that
 * release), after they were checked against the bundle's manifest and against
 * the release tag. Only the deployment-local .clasp.json differs, and it is
 * not part of the uploaded payload.
 *
 * Which files a release consists of is read from the release's own tag (the
 * RELEASE_FILES list of this tool as it was at the tag), never from the list
 * below: T-2026.10.11 and earlier are seven files, later releases are five.
 * See docs/PRODUCTION_FILE_SET.md.
 *
 *   node tools/template-release.js --adopt T-2026.10.4 --target-script <id> --label <name>            (dry run)
 *   node tools/template-release.js --adopt T-2026.10.4 --target-script <id> --label <name> --write
 *
 * Adoption guardrails: the target Script ID must be given explicitly; the
 * CENTRAL add-on is always refused; the master template's own script is
 * refused (that is a normal release, not an adoption); the Legacy bound
 * script is refused unless --allow-legacy-target is given. There is no
 * force option here: a clasp manifest overwrite stays a separate decision.
 *
 * Guardrails (each one is a real past or near-miss incident):
 *   - the target Script ID must come from template/template-target.json and
 *     must not be the Legacy or CENTRAL project (sa4-report/.clasp.json
 *     itself points at Legacy today);
 *   - only whitelisted runtime files are bundled -- never tests/ (LEGACY-
 *     UPGRADE-002: leaked test files with require() broke onOpen);
 *   - the working tree must be clean and HEAD must carry the tag
 *     template-release/<releaseId>, so every template state maps to Git;
 *   - Release.js records releaseId + commit, and the report's first run
 *     copies it into Document Properties.
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const REPO_ROOT = path.join(__dirname, '..');

// Projects that must never receive a template push. They are listed by the
// SHA-256 digest of their Script ID, not by the ID itself: this file is
// public, the IDs are not. A target is protected when its digest matches.
const PROTECTED_PROJECT_LABELS = {
  legacy: 'Legacy bound script (6G Media Minutes, meeting 86178)',
  central: 'CENTRAL editor add-on'
};
const PROTECTED_SCRIPT_ID_DIGESTS = {
  legacy: '0b5b6122b8c4d661e5c9ce65542441ca724c93880a8c5f15e371cbbc576c556b',
  central: 'a536eb96d5a5ec45b1275a76b56afde4cbadd3c083ec44126da2ef6107c6268d'
};
const scriptIdDigest = (id) => crypto.createHash('sha256').update(String(id === null || id === undefined ? '' : id)).digest('hex');

/** A digest table for the given Script IDs -- for the tests, which use synthetic IDs. */
function protectionForScriptIds(ids) {
  return { legacy: scriptIdDigest(ids.legacy), central: scriptIdDigest(ids.central) };
}

/** 'legacy', 'central' or null. `protection` defaults to the built-in digests. */
function protectedProjectKey(id, protection) {
  const table = protection || PROTECTED_SCRIPT_ID_DIGESTS;
  const idDigest = scriptIdDigest(id);
  return Object.keys(table).find((key) => table[key] === idDigest) || null;
}

/** The name of the protected project a Script ID belongs to, or null. */
function protectedProject(id, protection) {
  const key = protectedProjectKey(id, protection);
  return key ? PROTECTED_PROJECT_LABELS[key] : null;
}

// Repository path -> published file name. Nothing else is ever bundled.
// This is the list of the release built from THIS commit. An existing release
// is described by the list at its own tag: releaseFilesFromToolSource() reads
// it from there, so keep one ['source', 'name'] pair per line.
const RELEASE_FILES = [
  ['appsscript.json', 'appsscript.json'],
  ['Code.js', 'Code.js'],
  ['HyperLink.js', 'HyperLink.js'],
  ['template/ReportCreator.js', 'ReportCreator.js']
];
const RELEASE_TOOL_PATH = 'tools/template-release.js';
const GENERATED_RELEASE_FILE = 'Release.js';

/**
 * The RELEASE_FILES list written in a copy of this tool's source -- the copy
 * a release tag holds -- or null when it cannot be read with certainty. The
 * source is parsed, never executed. Pure.
 */
function releaseFilesFromToolSource(source) {
  const block = String(source === null || source === undefined ? '' : source).replace(/\r\n/g, '\n').match(/^const RELEASE_FILES = \[\n([\s\S]*?)^\];/m);
  if (!block) return null;
  const list = [];
  const lines = block[1].split('\n').map((line) => line.trim()).filter(Boolean);
  for (let i = 0; i < lines.length; i++) {
    const pair = lines[i].match(/^\['([A-Za-z0-9_.\/-]+)', '([A-Za-z0-9_.-]+)'\],?$/);
    if (!pair || /(^|\/)\.\.?(\/|$)/.test(pair[1]) || pair[1].charAt(0) === '/' || pair[2] === GENERATED_RELEASE_FILE) return null;
    if (list.some((entry) => entry[0] === pair[1] || entry[1] === pair[2])) return null;
    list.push([pair[1], pair[2]]);
  }
  return list.length ? list : null;
}

/**
 * The source payload of an existing release, from its tag alone: the file
 * list the tag declares and each of those files as the tag holds it.
 * showAtTag(repoPath) returns the content or throws. Release.js is not among
 * them: it is generated at build time and exists only in the bundle. Pure.
 */
function taggedReleaseSources(showAtTag, tag) {
  let list = null;
  try { list = releaseFilesFromToolSource(showAtTag(RELEASE_TOOL_PATH)); } catch (e) { /* reported below */ }
  if (!list) {
    return { ok: false, releaseFiles: null, published: null, files: [], errors: ['Could not read the release file list of ' + tag + ' (RELEASE_FILES in ' + RELEASE_TOOL_PATH + ' at the tag).'] };
  }
  const errors = [];
  const files = [];
  list.forEach(([src, name]) => {
    try { files.push({ name, source: src, content: showAtTag(src) }); } catch (e) { errors.push('Could not read ' + src + ' at ' + tag + '.'); }
  });
  return { ok: errors.length === 0, releaseFiles: list, published: list.map(([, name]) => name).concat([GENERATED_RELEASE_FILE]), files, errors };
}

const RELEASE_ID_RE = /^T-\d{4}\.\d{2}\.\d+$/;
const SCRIPT_ID_RE = /^[A-Za-z0-9_-]{40,80}$/;
const DRIVE_ID_RE = /^[A-Za-z0-9_-]{25,100}$/;

function validateTarget(target, protection) {
  const errors = [];
  if (!target || typeof target !== 'object') return ['template/template-target.json is missing or not an object.'];
  if (!SCRIPT_ID_RE.test(String(target.templateScriptId || ''))) errors.push('templateScriptId is missing or malformed.');
  if (!DRIVE_ID_RE.test(String(target.templateDocumentId || ''))) errors.push('templateDocumentId is missing or malformed.');
  const protectedAs = protectedProject(target.templateScriptId, protection);
  if (protectedAs) {
    errors.push('templateScriptId is the ' + protectedAs + ' -- refusing.');
  }
  return errors;
}

function buildReleaseJs(meta) {
  return '/**\n * GENERATED by tools/template-release.js -- do not edit.\n' +
    ' * Present only in SA4 Report Template release bundles.\n */\n' +
    'var SA4_RELEASE_ = ' + JSON.stringify(meta, null, 2) + ';\n';
}

function buildClaspIgnore(publishedNames) {
  return ['**/**'].concat(publishedNames.map((n) => '!' + n)).join('\n') + '\n';
}

/**
 * Pure release plan. git: { commit, dirty, tagsAtHead }. readFile(relPath)
 * returns the file content (Buffer or string) or throws.
 */
function planRelease({ releaseId, target, git, readFile, builtAt, protection }) {
  const errors = [];
  if (!RELEASE_ID_RE.test(String(releaseId || ''))) errors.push('--release must look like T-2026.10.0.');
  errors.push(...validateTarget(target, protection));
  if (!git || !/^[0-9a-f]{40}$/.test(String(git.commit || ''))) errors.push('Could not read the HEAD commit.');
  if (git && git.dirty) errors.push('The working tree is not clean -- commit or stash first.');
  const tag = 'template-release/' + releaseId;
  if (git && (git.tagsAtHead || []).indexOf(tag) === -1) errors.push('HEAD is not tagged ' + tag + '.');

  const files = [];
  RELEASE_FILES.forEach(([src, name]) => {
    let content;
    try { content = readFile(src); } catch (e) { errors.push('Missing release file ' + src + '.'); return; }
    files.push({ name, source: src, content });
  });
  const manifestFile = files.find((f) => f.name === 'appsscript.json');
  if (manifestFile) {
    try {
      const manifest = JSON.parse(String(manifestFile.content));
      if (manifest.addOns) errors.push('appsscript.json declares addOns -- the template runtime is a bound script, not an add-on.');
    } catch (e) {
      errors.push('appsscript.json is not valid JSON.');
    }
  }
  files.forEach((f) => {
    if (/\brequire\s*\(/.test(String(f.content)) && /\.js$/.test(f.name)) {
      errors.push(f.name + ' calls require() -- Node-only code would break onOpen in Apps Script.');
    }
  });

  // The Code.js version of this snapshot, from its own header line.
  const codeFile = files.find((f) => f.name === 'Code.js');
  const codeVersion = codeFile ? (String(codeFile.content).slice(0, 400).match(/^ \* Version: (\d+\.\d+\.\d+)/m) || [])[1] || null : null;
  if (codeFile && !codeVersion) errors.push('Code.js has no "Version: x.y.z" header line.');

  const meta = {
    releaseId,
    flavor: 'template',
    codeVersion,
    gitCommit: git ? git.commit : null,
    gitTag: tag,
    builtAt,
    templateDocumentId: target ? target.templateDocumentId : null,
    templateScriptId: target ? target.templateScriptId : null
  };
  files.push({ name: GENERATED_RELEASE_FILE, source: '(generated)', content: buildReleaseJs(meta) });

  const published = files.map((f) => f.name);
  return {
    ok: errors.length === 0,
    errors,
    meta,
    files,
    claspJson: target ? JSON.stringify({ scriptId: target.templateScriptId, rootDir: '.' }, null, 2) + '\n' : null,
    claspIgnore: buildClaspIgnore(published),
    manifest: {
      releaseId,
      codeVersion,
      gitCommit: meta.gitCommit,
      templateScriptId: meta.templateScriptId,
      files: files.map((f) => ({ name: f.name, source: f.source, sha256: crypto.createHash('sha256').update(f.content).digest('hex') }))
    }
  };
}

// ------------------------------------------------------------------
// Adoption mode: an existing release bundle for another script project
// ------------------------------------------------------------------

const ADOPT_LABEL_RE = /^[A-Za-z0-9][A-Za-z0-9._-]{0,39}$/;
const sha256 = (content) => crypto.createHash('sha256').update(content).digest('hex');
const lf = (content) => String(content).replace(/\r\n/g, '\n');

/**
 * Why a Script ID may not receive an adopted payload, or '' when it may.
 * `templateScriptId` is the master template's own script, from the
 * release's Release.js. Pure.
 */
function adoptionTargetRefusal(targetScriptId, templateScriptId, allowLegacyTarget, protection) {
  const id = String(targetScriptId || '');
  if (!id) return '--target-script is required: adoption never guesses a target.';
  if (!SCRIPT_ID_RE.test(id)) return '--target-script is not a Script ID.';
  const project = protectedProjectKey(id, protection);
  if (project === 'central') return 'The target is the ' + PROTECTED_PROJECT_LABELS.central + ' -- refusing. It is never an adoption target.';
  if (id === templateScriptId) {
    return 'The target is the master template\'s own script -- refusing. Use the normal release (--release), not --adopt.';
  }
  if (project === 'legacy' && !allowLegacyTarget) {
    return 'The target is the ' + PROTECTED_PROJECT_LABELS.legacy + ' -- refusing. ' +
      'Replacing it needs the explicit --allow-legacy-target flag and a separate authorisation.';
  }
  return '';
}

/**
 * Pure adoption plan. bundle: { files: { name: Buffer }, manifest } -- an
 * existing release bundle as read from disk (null when it does not exist).
 * git: { tagCommit, showAtTag(repoPath) -> content or throws }.
 * protection: optional digest table (tests); the built-in one otherwise.
 *
 * The expected files are those of the release's own tag (see
 * taggedReleaseSources()), not today's RELEASE_FILES: a seven-file release
 * stays seven files, and nothing is adopted from a tag that does not say
 * which files it has.
 */
function planAdoption({ releaseId, targetScriptId, label, allowLegacyTarget, bundle, git, protection }) {
  const errors = [];
  if (!RELEASE_ID_RE.test(String(releaseId || ''))) errors.push('--adopt must name a release like T-2026.10.4.');
  if (!ADOPT_LABEL_RE.test(String(label || ''))) errors.push('--label is required (letters, digits, ".", "_", "-"; it names the output folder).');
  const tag = 'template-release/' + releaseId;
  const tagExists = !!git && /^[0-9a-f]{40}$/.test(String(git.tagCommit || ''));
  if (!tagExists) errors.push('Tag ' + tag + ' does not exist in this repository.');

  // What the release consists of, from its tag. Without it nothing below
  // can be checked, and nothing is planned.
  const tagged = tagExists && git.showAtTag ? taggedReleaseSources(git.showAtTag, tag) : null;
  if (tagged && !tagged.releaseFiles) errors.push(...tagged.errors);
  const published = tagged && tagged.published ? tagged.published : null;

  let release = null;
  const files = [];
  if (!bundle || !bundle.files || !bundle.manifest) {
    errors.push('The release bundle for ' + releaseId + ' was not found. Adoption copies an existing bundle; it never rebuilds a release.');
  } else if (published) {
    const listed = (bundle.manifest.files || []).map((f) => f.name);
    if (JSON.stringify(listed) !== JSON.stringify(published)) {
      errors.push('The bundle manifest does not list exactly the ' + published.length + ' files of ' + tag + ' (' + published.join(', ') + ').');
    }
    if (bundle.manifest.releaseId !== releaseId) errors.push('The bundle manifest is for ' + bundle.manifest.releaseId + ', not ' + releaseId + '.');
    published.forEach((name) => {
      const content = bundle.files[name];
      if (content === undefined || content === null) { errors.push('The bundle has no ' + name + '.'); return; }
      const entry = (bundle.manifest.files || []).find((f) => f.name === name);
      if (entry && entry.sha256 !== sha256(content)) errors.push(name + ' in the bundle does not match its manifest hash -- the bundle was changed after it was built.');
      files.push({ name, content, sha256: sha256(content) });
    });
    Object.keys(bundle.files).forEach((name) => {
      if (published.indexOf(name) === -1) errors.push('The bundle contains an unexpected file: ' + name + '.');
    });

    // Release.js is the release identity: it must be the tagged release.
    const releaseJs = bundle.files[GENERATED_RELEASE_FILE];
    if (releaseJs) {
      try {
        release = JSON.parse(String(releaseJs).slice(String(releaseJs).indexOf('{'), String(releaseJs).lastIndexOf('}') + 1));
      } catch (e) {
        errors.push('Release.js in the bundle cannot be read.');
      }
    }
    if (release) {
      if (release.releaseId !== releaseId) errors.push('Release.js says ' + release.releaseId + ', not ' + releaseId + '.');
      if (release.flavor !== 'template') errors.push('Release.js is not a template release.');
      if (git && git.tagCommit && release.gitCommit !== git.tagCommit) errors.push('Release.js was built from ' + release.gitCommit + ', but ' + tag + ' is ' + git.tagCommit + '.');
    }
    // The source files must be what the tag holds (line endings aside: the
    // bundle carries the checkout's, git stores LF).
    errors.push(...tagged.errors);
    tagged.files.forEach((f) => {
      if (!bundle.files[f.name]) return;
      if (lf(f.content) !== lf(bundle.files[f.name])) errors.push(f.name + ' in the bundle differs from ' + f.source + ' at ' + tag + '.');
    });
  }

  const refusal = adoptionTargetRefusal(targetScriptId, release ? release.templateScriptId : null, allowLegacyTarget, protection);
  if (refusal) errors.push(refusal);

  const legacyTarget = protectedProjectKey(targetScriptId, protection) === 'legacy';
  return {
    ok: errors.length === 0,
    errors,
    release,
    files,
    legacyTarget,
    // Deployment-local only: the whitelist below keeps both files (and the
    // adoption record) out of the upload.
    claspJson: JSON.stringify({ scriptId: String(targetScriptId || ''), rootDir: '.' }, null, 2) + '\n',
    claspIgnore: published ? buildClaspIgnore(published) : null,
    record: {
      action: 'adopt',
      releaseId,
      codeVersion: release ? release.codeVersion : null,
      gitCommit: release ? release.gitCommit : null,
      targetScriptId: String(targetScriptId || ''),
      label,
      legacyTarget,
      files: files.map((f) => ({ name: f.name, sha256: f.sha256 }))
    }
  };
}

/** The lines printed before the deployment command (pure). */
function adoptionBanner(plan, outDir) {
  const r = plan.record;
  const lines = [
    '============================================================',
    '  ACTION : ADOPT an existing report -- push the UNCHANGED',
    '           payload of ' + r.releaseId + ' (Code.js ' + r.codeVersion + ', commit ' + String(r.gitCommit).slice(0, 7) + ')',
    '  TARGET : Script ID ' + r.targetScriptId,
    '  LABEL  : ' + r.label,
    '  This is NOT the master template and NOT a new release.'
  ];
  if (plan.legacyTarget) {
    lines.push('  !!! THE TARGET IS THE LIVE LEGACY BOUND SCRIPT (--allow-legacy-target) !!!');
    lines.push('  !!! Its current code is replaced. Rollback: the Legacy repository.      !!!');
  }
  lines.push('============================================================');
  if (outDir) {
    lines.push('Check the Script ID above against the report\'s Apps Script project settings, then:');
    lines.push('  cd "' + outDir + '" && clasp push');
    lines.push('If clasp asks to overwrite the manifest, stop: that needs a separate decision. This tool never forces a push.');
  }
  return lines;
}

function readBundle(dir) {
  if (!fs.existsSync(path.join(dir, 'RELEASE-MANIFEST.json'))) return null;
  const files = {};
  fs.readdirSync(dir).forEach((name) => {
    if (name === '.clasp.json' || name === '.claspignore' || name === 'RELEASE-MANIFEST.json') return;
    files[name] = fs.readFileSync(path.join(dir, name));
  });
  return { files, manifest: JSON.parse(fs.readFileSync(path.join(dir, 'RELEASE-MANIFEST.json'), 'utf8')) };
}

function adoptMain(argv) {
  const value = (flag) => { const i = argv.indexOf(flag); return i !== -1 ? argv[i + 1] : null; };
  const releaseId = value('--adopt');
  const label = value('--label');
  const tag = 'template-release/' + releaseId;
  const git = { tagCommit: null, showAtTag: null };
  try {
    git.tagCommit = execFileSync('git', ['rev-parse', '--verify', '--quiet', tag + '^{commit}'], { cwd: REPO_ROOT, encoding: 'utf8' }).trim();
    git.showAtTag = (rel) => execFileSync('git', ['show', tag + ':' + rel], { cwd: REPO_ROOT, maxBuffer: 64 * 1024 * 1024 });
  } catch (e) { /* reported by planAdoption */ }

  const plan = planAdoption({
    releaseId,
    targetScriptId: value('--target-script'),
    label,
    allowLegacyTarget: argv.includes('--allow-legacy-target'),
    bundle: RELEASE_ID_RE.test(String(releaseId || '')) ? readBundle(path.join(REPO_ROOT, 'dist', 'template-release', releaseId)) : null,
    git
  });
  if (!plan.ok) {
    console.error('Adoption refused:\n' + plan.errors.map((e) => '  - ' + e).join('\n'));
    return 1;
  }
  plan.files.forEach((f) => console.log('  ' + f.name.padEnd(28) + ' ' + f.sha256));
  const outDir = path.join(REPO_ROOT, 'dist', 'template-adopt', releaseId, label);
  if (!argv.includes('--write')) {
    console.log('\n' + adoptionBanner(plan, null).join('\n'));
    console.log('\nDry run OK. Re-run with --write to create ' + path.relative(REPO_ROOT, outDir));
    return 0;
  }
  if (fs.existsSync(outDir)) {
    console.error('\n' + outDir + ' already exists -- an adoption folder is never overwritten.');
    return 1;
  }
  fs.mkdirSync(outDir, { recursive: true });
  plan.files.forEach((f) => fs.writeFileSync(path.join(outDir, f.name), f.content));
  fs.writeFileSync(path.join(outDir, '.clasp.json'), plan.claspJson);
  fs.writeFileSync(path.join(outDir, '.claspignore'), plan.claspIgnore);
  fs.writeFileSync(path.join(outDir, 'ADOPTION.json'), JSON.stringify(plan.record, null, 2) + '\n');
  console.log('\n' + adoptionBanner(plan, outDir).join('\n'));
  return 0;
}

function readGit() {
  const run = (args) => execFileSync('git', args, { cwd: REPO_ROOT, encoding: 'utf8' }).trim();
  return {
    commit: run(['rev-parse', 'HEAD']),
    dirty: run(['status', '--porcelain']) !== '',
    tagsAtHead: run(['tag', '--points-at', 'HEAD']).split(/\r?\n/).filter(Boolean)
  };
}

function main(argv) {
  if (argv.includes('--adopt')) {
    if (argv.includes('--release')) {
      console.error('Use either --release (the master template) or --adopt (an existing report), not both.');
      return 1;
    }
    return adoptMain(argv);
  }
  const args = { write: argv.includes('--write') };
  const i = argv.indexOf('--release');
  args.release = i !== -1 ? argv[i + 1] : null;

  const targetPath = path.join(REPO_ROOT, 'template', 'template-target.json');
  let target = null;
  try { target = JSON.parse(fs.readFileSync(targetPath, 'utf8')); } catch (e) { /* reported by validateTarget */ }

  const plan = planRelease({
    releaseId: args.release,
    target,
    git: readGit(),
    readFile: (rel) => fs.readFileSync(path.join(REPO_ROOT, rel)),
    builtAt: new Date().toISOString()
  });

  plan.files.forEach((f) => console.log('  ' + f.name.padEnd(28) + ' <- ' + f.source));
  if (!plan.ok) {
    console.error('\nRelease refused:\n' + plan.errors.map((e) => '  - ' + e).join('\n'));
    return 1;
  }
  const outDir = path.join(REPO_ROOT, 'dist', 'template-release', args.release);
  if (!args.write) {
    console.log('\nDry run OK. Re-run with --write to create ' + path.relative(REPO_ROOT, outDir));
    return 0;
  }
  if (fs.existsSync(outDir)) {
    console.error('\n' + outDir + ' already exists -- a release bundle is never overwritten.');
    return 1;
  }
  fs.mkdirSync(outDir, { recursive: true });
  plan.files.forEach((f) => fs.writeFileSync(path.join(outDir, f.name), f.content));
  fs.writeFileSync(path.join(outDir, '.clasp.json'), plan.claspJson);
  fs.writeFileSync(path.join(outDir, '.claspignore'), plan.claspIgnore);
  fs.writeFileSync(path.join(outDir, 'RELEASE-MANIFEST.json'), JSON.stringify(plan.manifest, null, 2) + '\n');
  console.log('\nBundle written to ' + outDir);
  console.log('Review it, then push to the TEMPLATE project only:');
  console.log('  cd "' + outDir + '" && clasp push');
  return 0;
}

if (require.main === module) process.exitCode = main(process.argv.slice(2));

module.exports = {
  planRelease, validateTarget, buildReleaseJs, buildClaspIgnore, RELEASE_FILES,
  PROTECTED_SCRIPT_ID_DIGESTS, PROTECTED_PROJECT_LABELS, protectionForScriptIds, protectedProject, protectedProjectKey,
  planAdoption, adoptionTargetRefusal, adoptionBanner,
  releaseFilesFromToolSource, taggedReleaseSources, RELEASE_TOOL_PATH
};
