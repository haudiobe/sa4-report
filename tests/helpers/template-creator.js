/**
 * The two sides of "Create New SA4 Report", as tests/template-bootstrap.test.js
 * arranges them, as a helper: the creator (the master template's runtime,
 * with a captured Portal answer) and the first run in the new report (a
 * fresh sandbox = a fresh bound project with empty property stores).
 *
 * Nothing live is used. All ids are synthetic.
 */

const { loadTemplateRuntime } = require('./load-template.js');

const TEMPLATE_DOC_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const NEW_DOC_ID = 'NEWREPORTdoc00000000000000000000000000000';
const RELEASE = {
  releaseId: 'T-2026.10.8', flavor: 'template', gitCommit: 'd7c23bc0000000000000000000000000000000aa',
  gitTag: 'template-release/T-2026.10.8', templateDocumentId: TEMPLATE_DOC_ID
};
const plain = (v) => JSON.parse(JSON.stringify(v));

/** The master template's runtime; the Portal answers with `getMeetings`, the TDoc list page has `rows`. Lists and agenda files are not reachable. */
function creatorFor(getMeetings, rows) {
  const s = loadTemplateRuntime({ release: RELEASE }).sandbox;
  s.fetchMeetingMetadataById_ = () => ({ statusCode: 200, text: JSON.stringify(getMeetings) });
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => rows || [];
  s.UrlFetchApp = { fetch: () => { throw new Error('network down'); } };
  return s;
}

/** Creates the report: returns { preview, created, description } -- `description` is what the creator wrote on the new file. */
function createReport(getMeetings, rows, meetingId, choices) {
  const s = creatorFor(getMeetings, rows);
  const preview = plain(s.previewNewReportFromTemplate(meetingId, choices || {}));
  const descriptions = {};
  const deps = {
    release: RELEASE, activeDocumentId: () => TEMPLATE_DOC_ID, nowIso: () => '2026-10-06T08:00:00.000Z',
    copyTemplate: () => ({ id: NEW_DOC_ID, url: 'https://docs.google.com/document/d/' + NEW_DOC_ID + '/edit' }),
    setDescription: (id, text) => { descriptions[id] = text; }, trashFile: () => {}
  };
  const created = plain(s.createReportFromTemplateWith_(deps, { preview: s.computeResolvedMeetingPreview_({}, preview.resolved), choices: choices || {} }));
  return { s: s, preview: preview, created: created, description: descriptions[NEW_DOC_ID] };
}

/**
 * The new report's own runtime. `o`: { props (document properties it has
 * already), deps (replacing or adding dependencies of the first run) }.
 */
function firstRunFor(description, o) {
  const loaded = loadTemplateRuntime({ release: RELEASE, documentProperties: (o && o.props) || {} });
  const s = loaded.sandbox;
  const state = { description: description };
  const deps = Object.assign({
    release: RELEASE,
    activeDocumentId: () => NEW_DOC_ID,
    scriptId: () => 'NEWSCRIPTid000000000000000000000000000000000000000',
    nowIso: () => '2026-10-06T08:05:00.000Z',
    documentProperties: loaded.docProps,
    getOwnDescription: () => state.description,
    setOwnDescription: (text) => { state.description = text; },
    persistConfiguration: (config) => s.persistConfigurationSettings_(config),
    buildReadiness: () => plain(s.getBuildReadiness_())
  }, (o && o.deps) || {});
  return { s: s, deps: deps, state: state, docProps: loaded.docProps, run: () => plain(s.finishReportSetupWith_(deps)) };
}

/** The setup information of a description, and a description with other setup information. */
function payloadOf(s, description) {
  return plain(s.parseBootstrapDescription_(description).payload);
}
function descriptionWith(s, payload) {
  return s.serializeBootstrapDescription_(payload);
}

module.exports = { creatorFor, createReport, firstRunFor, payloadOf, descriptionWith, plain, RELEASE, TEMPLATE_DOC_ID, NEW_DOC_ID };
