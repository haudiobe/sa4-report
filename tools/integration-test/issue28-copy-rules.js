/**
 * INTEGRATION TEST ONLY -- issue #28. The pure rules of push-issue28-copy.mjs (tested by
 * tests/issue28-copy-rules.test.js). No network, no file system.
 *
 * This repository is public: documents and scripts are known here by the SHA-256 digest of their ID only
 * (as tools/template-release.js does for Legacy and CENTRAL). An ID given on the command line or returned by
 * the Apps Script API is accepted only when its digest is the expected one.
 */
const crypto = require('crypto');
const { protectedProjectKey, PROTECTED_PROJECT_LABELS } = require('../template-release.js');

const idDigest = (id) => crypto.createHash('sha256').update(String(id === null || id === undefined ? '' : id)).digest('hex');
/** The digest of a script file's source, line ends ignored (the API returns LF). */
const sourceSha = (s) => crypto.createHash('sha256').update(String(s).replace(/\r\n/g, '\n'), 'utf8').digest('hex');

// The two DISPOSABLE copies the user made on 2026-11-23 and their bound scripts; their script files as first read
// (2026-11-25 07:44 UTC, read-only, before anything was pushed). main: a copy of 6G Media SA4#138, T-2026.10.33.
// ahg: a copy of the Legacy 6G AH report, T-2026.10.19 (its own adoption of 2026-11-01).
const TARGETS = {
  // The clean copy (2026-11-25 16:18 UTC) of the 6G Media SA4#138 report; its bound script is named by the user
  // (no API finds it). Named 2026-11-26; read-only first read the same day: bound to this document (parentId), only
  // its HEAD deployment, no versions, every file equal to the production T-2026.10.33 bundle below. It replaces the
  // retired main copy.
  main2: {
    documentDigest: '00e2719b841d479c10de1325d55a0e783aa915777f0770f01f54501cd7df7ef0',
    scriptDigest: '71371d644e36741e4d74cc3e0585591a9d7e7449a5bdf9e27c8ec150ac24f09b',
    original: {
      appsscript: '4c2cc5ea159da448851aa22363457c1cf1cdd272a980bff45ecad8e5cf651082',
      Code: '99bf89d44592bd8920b41ebd69d81d5b9bc5cbaada6da2200d97391e26eeac23',
      HyperLink: 'df0e1b10d6860f705b6ed26ac86453ed3ee0e81a680f19128241d10ca92ce9fa',
      Release: '550bb7c5017613e30c8f5f02c59bf6fcbf425f7e7e7512403ac10006828e321b',
      ReportCreator: '7fe51e5e2203a638af1c7c6d07a0d901cc711eb877d23b796bd4705594dd07b0'
    }
  },
  main: {
    // Changed by the updates killed on 2026-11-25 (139 -> 140 TDoc tables): kept as evidence, never written again.
    retired: 'the first main copy, changed by the killed updates of 2026-11-25 and kept as evidence',
    documentDigest: 'e8f43b8da78448e92f220dbc3f6ab8092a0c55a8d901923d807a8bbdc9fce1b1',
    scriptDigest: 'e9cc31be0635e92b41ac76c114e0959fcd2dcebf9086b0f6a38bab88c38a8c20',
    original: {
      appsscript: '4c2cc5ea159da448851aa22363457c1cf1cdd272a980bff45ecad8e5cf651082',
      Code: '99bf89d44592bd8920b41ebd69d81d5b9bc5cbaada6da2200d97391e26eeac23',
      HyperLink: 'df0e1b10d6860f705b6ed26ac86453ed3ee0e81a680f19128241d10ca92ce9fa',
      Release: '550bb7c5017613e30c8f5f02c59bf6fcbf425f7e7e7512403ac10006828e321b',
      ReportCreator: '7fe51e5e2203a638af1c7c6d07a0d901cc711eb877d23b796bd4705594dd07b0'
    }
  },
  ahg: {
    documentDigest: 'd98a4d5c8b6a5d5f8b77d47048cef7347c7fdcddc6df5e62bf180d8c430091ac',
    scriptDigest: 'dd248145eb40a0704af6aeb6cfe716594824bb01dee2b9f46199e9a3a0403fcb',
    original: {
      appsscript: '4c2cc5ea159da448851aa22363457c1cf1cdd272a980bff45ecad8e5cf651082',
      Code: 'e8c5b8b52ee5c90da0ea76252a334e67fb568c830c9f549413039056b7278f44',
      HyperLink: 'df0e1b10d6860f705b6ed26ac86453ed3ee0e81a680f19128241d10ca92ce9fa',
      Release: 'b90ba87e4cb3e58ccc2ae0c8a61edac7e48300c7c5edc9eeda0cf45bfca277a8',
      ReportCreator: 'b4a5c2ea5f7e2214a1db9a94336f3fe93b3126701b83a78dda368087d5b68741'
    }
  }
};

// Never a target, whatever else is true (digests of the IDs).
const REFUSED_DOCUMENT_DIGESTS = {
  '52d82baa6f520368257374a29e6cfaed042b885500d863326faab02db8cca5cc': 'the SA4 Report Template (master)',
  '04c6f4603a9188d1a1c2c6c59071a47a981424511c597eac6ebe9692e9c1bac4': 'MBS SA4#138',
  '12a5d26658048eb1c104f580187ac76613aed8338d5db89c96b2b1e87fcc4996': 'Video SA4#138',
  'f0718379ccee0d1750ba2cbf97e5b2670e03895e87754284a8687f6afff726e3': '6G Media SA4#138',
  'efa7506097faaa0b4fb66f31c8c549cc502fb2dbefa621a8acbfa68fbea0855d': 'RTC SA4#138',
  '37d155fa5e383c2fa487000f43e232de7db26a8e0bddaaa66484c80ef41505da': 'Audio SA4#138',
  'e985bf59891602c65b9b3ab38329b5640c4db1da93768f412709f8c8fd4a1550': 'the Legacy 6G AH report'
};
const REFUSED_SCRIPT_DIGESTS = {
  'dc40c0e2411c80fd0a2d4526402efc8c4ed32777e1a9c2fc8f07a20f09fee516': 'the bound script of the SA4 Report Template (master)',
  '04a03e9024efb73bd3482c8f6de1f913862a52bc4d54a796fed94cde249fc343': 'MBS SA4#138',
  '819d4717c5a7fa3c5b6eba8d95776978d3837a6d47141048f386caaf9ff6d799': 'Video SA4#138',
  '7b715e7148ea541b5d387b2853b51ce9b823ebcd03d471763d86447f1ae0c20c': '6G Media SA4#138',
  '09f7c044aaa1ca2c9d4a5a3c8ac8a83b5e1278561ea6e73b7e1d67eb76d4da06': 'RTC SA4#138',
  '453dd1dfbad4e3059c990b6ea2ef23387b0ba764be2afb9d56f2dd360f5f5689': 'Audio SA4#138'
};

// The code of each stage, by commit: a branch that moved is not pushed (its code must equal the pinned commit's).
const STAGE_PINS = {
  t33: { commit: 'b7299c12d12befe089070379ba5f50541d5dd929', branch: null, label: 'T-2026.10.33 (production)' },
  A: { commit: '03d7aa4222f64d197773b1bfc528f97d009214d6', branch: 'feature/i28-a-titles-headings-perf', label: 'phase A' },
  B: { commit: '8e2d0f54ac41edd5143d9cf1fc19beb6b8cb4856', branch: 'feature/i28-b-provisional-revisions', label: 'phases A+B' },
  C: { commit: '1ea0c06d2d263e610d38c2ea7980d282a5e377a8', branch: 'feature/i28-c-ahg-sessions-agenda', label: 'phases A+B+C' }
};
const CODE_PATHS = { Code: 'Code.js', HyperLink: 'HyperLink.js', ReportCreator: 'template/ReportCreator.js' };
// Release.js of T-2026.10.33 as the five production reports read it back on 2026-11-20 (= the main copy's original).
const RELEASE_T33_SHA = '550bb7c5017613e30c8f5f02c59bf6fcbf425f7e7e7512403ac10006828e321b';
const STAGES = ['base', 't33', 'A', 'B', 'C', 'original'];
// The pinned Issue28Check.js (re-pinned deliberately 2026-11-25: fingerprint v2 -- Docs REST read with a field mask, canonical,
// compared by TDoc number and row, snapshots in Drive and verified; review fixes; all tabs; the runner; fixed steps and cleanup; the update result read; no retry loops, status). Before: ef08027f..., 799c1c56..., 9389a540..., a6fe0b9d..., 7a5d4932..., 01016ee9..., 1b5e45cf..., 26f9a482... A different harness is refused.
const CHECK_SHA = '1208d3eda960497ae67facba329ae98375085f26cb5af91171a116e0ef3ae7ec';

/** Every reason why this script/document may not be written as `target` ([] when it may). Pure. */
function identityRefusals(target, scriptId, project, file) {
  const out = [];
  const t = TARGETS[target];
  if (!t) return ['unknown target ' + target];
  if (t.retired) return ['the target is retired: ' + t.retired];
  if (!t.scriptDigest) return ['the bound script of the ' + target + ' copy is not known yet (its Script ID is needed once)'];
  const sd = idDigest(scriptId);
  if (REFUSED_SCRIPT_DIGESTS[sd]) out.push('the script is ' + REFUSED_SCRIPT_DIGESTS[sd]);
  const pk = protectedProjectKey(scriptId);
  if (pk) out.push('the script is the ' + PROTECTED_PROJECT_LABELS[pk]);
  if (sd !== t.scriptDigest) out.push('the script is not the bound script of the ' + target + ' copy');
  if (!project || project.scriptId !== scriptId) out.push('the Apps Script API did not return this project');
  const parent = project && project.parentId;
  const pd = idDigest(parent);
  if (REFUSED_DOCUMENT_DIGESTS[pd]) out.push('the script is bound to ' + REFUSED_DOCUMENT_DIGESTS[pd]);
  if (pd !== t.documentDigest) out.push('the script is not bound to the ' + target + ' copy');
  if (!file || file.id !== parent) out.push('Drive did not return the parent document');
  else {
    if (!String(file.name || '').startsWith('DISPOSABLE')) out.push('the document name does not start with "DISPOSABLE"');
    if (file.trashed) out.push('the document is in the trash');
  }
  return out;
}

/** Pure: canonical JSON (keys sorted at every level): the editor re-saves appsscript.json with keys in another order. */
const canonicalJson = (v) => (Array.isArray(v) ? '[' + v.map(canonicalJson).join(',') + ']'
  : v && typeof v === 'object' ? '{' + Object.keys(v).sort().map((k) => JSON.stringify(k) + ':' + canonicalJson(v[k])).join(',') + '}' : JSON.stringify(v));
/** {name: sha} of API files; a JSON file by its canonical content (key order is not a change, any value is). */
const shaMap = (files) => {
  const m = {};
  (files || []).forEach((f) => { m[f.name] = f.type === 'JSON' ? sourceSha(canonicalJson(JSON.parse(f.source))) : sourceSha(f.source); });
  return m;
};
/** The pinned original as {name: sha} in shaMap terms (appsscript.json compared by content). */
function originalShaMap(target, originalFiles) {
  // The pin is of the raw source; the original files themselves (verified against it) give the canonical form.
  const pin = TARGETS[target].original;
  const raw = {};
  (originalFiles || []).forEach((f) => { raw[f.name] = sourceSha(f.source); });
  return { matchesPin: sameMap(raw, pin), canonical: shaMap(originalFiles) };
}
const sameMap = (a, b) => { const ka = Object.keys(a).sort(); const kb = Object.keys(b).sort(); return JSON.stringify(ka) === JSON.stringify(kb) && ka.every((k) => a[k] === b[k]); };
/** Pure: how `now` differs from `expected` (both {name: sha}): [] when equal. */
function drift(expected, now) {
  const out = [];
  Object.keys(expected).forEach((k) => { if (!(k in now)) out.push(k + ': missing'); else if (now[k] !== expected[k]) out.push(k + ': ' + expected[k].slice(0, 12) + ' expected, ' + now[k].slice(0, 12) + ' found'); });
  Object.keys(now).forEach((k) => { if (!(k in expected)) out.push(k + ': unexpected file'); });
  return out;
}

/**
 * Pure: the files of a stage. o: { original: [files], checkSource, codeAt(commit, path), t33Release: source, codeVersionOf(source), stamp }.
 * The copy's appsscript.json is always kept.
 */
function stageFiles(stage, o) {
  const check = { name: 'Issue28Check', type: 'SERVER_JS', source: o.checkSource };
  // The runner's job (tools/integration-test/issue28-jobs.js) goes with the harness, never with the original.
  if (o.jobSource && stage === 'original') throw new Error('a job cannot go with the original (no harness, no runner)');
  const job = o.jobSource ? [{ name: 'Issue28Job', type: 'SERVER_JS', source: o.jobSource }] : [];
  const original = o.original.filter((f) => f.name !== 'Issue28Check' && f.name !== 'Issue28Job');
  if (stage === 'original') return original;
  if (stage === 'base') return original.concat([check], job);
  const pin = STAGE_PINS[stage];
  if (!pin) throw new Error('unknown stage ' + stage);
  const manifest = original.find((f) => f.name === 'appsscript');
  const code = Object.keys(CODE_PATHS).map((n) => ({ name: n, type: 'SERVER_JS', source: o.codeAt(pin.commit, CODE_PATHS[n]) }));
  let release;
  if (stage === 't33') {
    if (!o.t33Release || sourceSha(o.t33Release) !== RELEASE_T33_SHA) throw new Error('the T-2026.10.33 Release.js is not the production one (digest)');
    release = o.t33Release;
  } else {
    const base = /var SA4_RELEASE_ = (\{[\s\S]*?\});/.exec(o.t33Release || '');
    if (!base) throw new Error('no T-2026.10.33 Release.js to build the candidate release from');
    const b = JSON.parse(base[1]);
    const rec = Object.assign({}, b, { codeVersion: o.codeVersionOf(code[0].source) + ' (issue #28 ' + pin.label + ' candidate, NOT RELEASED)', gitCommit: pin.commit,
      gitTag: 'candidate ' + pin.branch + ' (NOT RELEASED; disposable-copy integration test)', builtAt: o.stamp });
    release = '/**\n * INTEGRATION TEST ONLY -- tools/integration-test/push-issue28-copy.mjs, into a DISPOSABLE COPY.\n * Base release ' + b.releaseId + ' with the issue #28 ' + pin.label + ' candidate.\n */\nvar SA4_RELEASE_ = ' + JSON.stringify(rec, null, 2) + ';\n';
  }
  const extra = original.filter((f) => !['appsscript', 'Release'].concat(Object.keys(CODE_PATHS)).includes(f.name));
  if (extra.length) throw new Error('the copy has files the bundle does not know: ' + extra.map((f) => f.name).join(', '));
  return [manifest].concat(code, [{ name: 'Release', type: 'SERVER_JS', source: release }, check], job);
}

/** Pure: why a branch is not at its pin ('' when its code files equal the pinned commit's). codeAt(ref, path). */
function pinRefusal(stage, codeAt) {
  const pin = STAGE_PINS[stage];
  if (!pin || !pin.branch) return '';
  const moved = Object.values(CODE_PATHS).filter((p) => sourceSha(codeAt(pin.branch, p)) !== sourceSha(codeAt(pin.commit, p)));
  return moved.length ? pin.branch + ' has moved from its pinned commit ' + pin.commit.slice(0, 7) + ' (' + moved.join(', ') + ' differ): re-pin deliberately or push the pinned code' : '';
}

/** Pure: the protected or copy IDs written in plain text in `text` (by digest), as their labels. */
function plainIdsIn(text) {
  const labels = Object.assign({}, REFUSED_DOCUMENT_DIGESTS, REFUSED_SCRIPT_DIGESTS);
  Object.keys(TARGETS).forEach((k) => { labels[TARGETS[k].documentDigest] = 'the ' + k + ' copy'; labels[TARGETS[k].scriptDigest] = 'the ' + k + ' copy\'s script'; });
  const out = [];
  new Set(String(text).match(/[-\w]{25,}/g) || []).forEach((tok) => {
    const d = idDigest(tok);
    const label = labels[d] || (protectedProjectKey(tok) ? PROTECTED_PROJECT_LABELS[protectedProjectKey(tok)] : '');
    if (label) out.push(label);
  });
  return out;
}

module.exports = { CHECK_SHA, plainIdsIn, canonicalJson, TARGETS, REFUSED_DOCUMENT_DIGESTS, REFUSED_SCRIPT_DIGESTS, STAGE_PINS, STAGES, RELEASE_T33_SHA, CODE_PATHS,
  idDigest, sourceSha, identityRefusals, shaMap, originalShaMap, drift, stageFiles, pinRefusal };
