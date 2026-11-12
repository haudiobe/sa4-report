/**
 * INTEGRATION TEST ONLY (issue #6 / PR #7): the pure rules for pushing a
 * candidate into the bound script of a DISPOSABLE COPY of a report. Used by
 * push-to-disposable-copy.mjs; tested by tests/integration-test-kit.test.js.
 */
const crypto = require('crypto');
const { protectedProject } = require('../template-release.js');

// Never targets, whatever else is true.
const REFUSED_DOCUMENTS = {
  '1zawfIcbkzCJG-h2HYtnbQSEsm6MCd9OGhYzMgAFYENs': 'the original Audio SWG Minutes SA4#138 report',
  '11XNUHU3TpAr8BaS7gyFQlwiKEzQFvnvzvGu4ms8lo2s': 'the SA4 Report Template (master)'
};
const REFUSED_SCRIPTS = {
  '1TJAxpoSF2t6RS4FdO9TuCkHeM4eyaI7bDxcZk8v8GkaexZJLiOvKtTN1': 'the bound script of the original Audio report',
  '17xPMiY9PTft9KblONpetvZCA0Z7N_cwbG28qNMmYh3_QbPFDJhvtG8Pr': 'the bound script of the SA4 Report Template (master)',
  '1jZOioQS9_dTU8uTwD4OrOQBHFwN6tX_T9f7xr76nxbwm9RQj5ySoznXA': 'the project of the repository .clasp.json'
};
// The name the copy must carry, so that nobody takes it for a report.
const DISPOSABLE_NAME_PREFIX = 'DISPOSABLE TEST COPY';
// Code.js of T-2026.10.22, the release the Audio report runs (RELEASE-MANIFEST.json): a copy of Audio starts from it.
const AUDIO_BASE_CODE_SHA256 = '240017f562e60e798cd3c22909c13cddf41c0eea5954e70865cb7d7d14756135';

const sha256 = (buf) => crypto.createHash('sha256').update(buf).digest('hex');

/**
 * Every reason the push must not happen ([] when it may).
 * facts: { scriptId, documentId, parentId, documentName, documentTrashed, baseCodeSha256, allowOtherBase }
 */
function pushRefusals(facts) {
  const f = facts || {};
  const out = [];
  const scriptId = String(f.scriptId || '').trim();
  const documentId = String(f.documentId || '').trim();
  if (!scriptId) out.push('--script is required.');
  if (!documentId) out.push('--doc is required.');
  if (REFUSED_SCRIPTS[scriptId]) out.push('The script is ' + REFUSED_SCRIPTS[scriptId] + ' -- refusing.');
  const prot = scriptId ? protectedProject(scriptId) : null;
  if (prot) out.push('The script is the ' + prot + ' -- refusing.');
  if (REFUSED_DOCUMENTS[documentId]) out.push('The document is ' + REFUSED_DOCUMENTS[documentId] + ' -- refusing.');
  if (f.parentId !== undefined) {
    if (!f.parentId) out.push('The script is not bound to a document.');
    else if (REFUSED_DOCUMENTS[f.parentId]) out.push('The script is bound to ' + REFUSED_DOCUMENTS[f.parentId] + ' -- refusing.');
    else if (f.parentId !== documentId) out.push('The script is bound to ' + f.parentId + ', not to --doc ' + documentId + '.');
  }
  if (f.documentName !== undefined && String(f.documentName || '').indexOf(DISPOSABLE_NAME_PREFIX) !== 0) {
    out.push('The document name must start with "' + DISPOSABLE_NAME_PREFIX + '" (it is "' + f.documentName + '"). Rename the copy first.');
  }
  if (f.documentTrashed) out.push('The document is in the trash.');
  if (f.baseCodeSha256 !== undefined && f.baseCodeSha256 !== AUDIO_BASE_CODE_SHA256 && !f.allowOtherBase) {
    out.push('The script\'s Code.js is not T-2026.10.22 (' + (f.baseCodeSha256 || 'missing') + '): this is not a fresh copy of the Audio report. ' +
      '(--allow-other-base to push anyway, e.g. into a copy that already has the candidate.)');
  }
  return out;
}

/**
 * The Release.js of the candidate on the copy: the copy's own, with the
 * candidate's commit. releaseId and the template IDs stay (the runtime
 * validates releaseId as T-YYYY.MM.N and uses templateDocumentId to tell the
 * master from a report); codeVersion and gitTag say it is not a release.
 */
function candidateReleaseJs(baseReleaseJs, commit, builtAt) {
  const m = /var SA4_RELEASE_ = (\{[\s\S]*?\});/.exec(String(baseReleaseJs || ''));
  if (!m) throw new Error('The copy has no Release.js with SA4_RELEASE_.');
  const base = JSON.parse(m[1]);
  if (!/^[0-9a-f]{40}$/.test(String(commit || ''))) throw new Error('The candidate commit must be a full SHA-1.');
  const release = Object.assign({}, base, {
    codeVersion: base.codeVersion + ' + issue #6 candidate (NOT RELEASED)',
    gitCommit: commit,
    gitTag: 'candidate fix/audio-large-build-checkpoint (NOT RELEASED; disposable-copy integration test)',
    builtAt: builtAt
  });
  return '/**\n * INTEGRATION TEST ONLY -- written by tools/integration-test/push-to-disposable-copy.mjs\n' +
    ' * into a DISPOSABLE COPY. Base release ' + base.releaseId + ' with the candidate of issue #6.\n */\n' +
    'var SA4_RELEASE_ = ' + JSON.stringify(release, null, 2) + ';\n';
}

module.exports = { pushRefusals, candidateReleaseJs, sha256, REFUSED_DOCUMENTS, REFUSED_SCRIPTS, DISPOSABLE_NAME_PREFIX, AUDIO_BASE_CODE_SHA256 };
