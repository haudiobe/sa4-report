/**
 * A template report in which the REAL Configure Meeting runs: the real
 * dialog is opened, its real client script is executed against a minimal
 * page, and its buttons call the real server functions -- the arrangement
 * of tests/template002b-runtime.test.js, as a helper.
 *
 * The Portal answers from `meetings` ({ id: [meeting] }, the form of
 * GetMeetings); every other request goes to `network(url, options)`, which
 * returns { status, text } or throws. Every request is recorded.
 * All ids are synthetic.
 */

const vm = require('vm');
const { loadTemplateRuntime } = require('./load-template.js');
const { makeFakeDocumentBody } = require('./fake-document.js');

const REPORT_ID = 'REPORTdoc000000000000000000000000000000000';
const TEMPLATE_ID = 'TEMPLATEdoc0000000000000000000000000000000';
const RELEASE = { releaseId: 'T-2026.10.8', flavor: 'template', gitCommit: 'd7c23bc0000000000000000000000000000000aa', gitTag: 'template-release/T-2026.10.8', templateDocumentId: TEMPLATE_ID };
const plain = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)));
const FIELD_IDS = ['meetingId', 'meetingType', 'meetingName', 'meetingDate', 'ftpBase', 'agendaTdoc', 'mailingList', 'meetingFolder', 'meetingNumber',
  'reportType', 'familyInfo', 'readinessRules', 'tdocUrl', 'revisionsUrl', 'agendaSourceDocId', 'showPreview', 'apiToken', 'clearApiToken', 'discussionEmailSender'];

/**
 * options: { props, scriptProps, meetings, rows (the TDocs of the Portal's TDoc list page), network, release (false: not a
 * template report), drive (a fake Drive of tests/helpers/fake-drive-files.js: installed as the Drive of the user) }
 * The document is one that was looked at already (SA4_BOOTSTRAP_STATE), unless props say otherwise.
 */
function configurableReport(options) {
  const o = options || {};
  const loaded = loadTemplateRuntime({ release: o.release === false ? null : RELEASE, documentProperties: Object.assign({ SA4_BOOTSTRAP_STATE: 'manual' }, o.props || {}), scriptProperties: o.scriptProps || {} });
  const s = loaded.sandbox;
  if (o.drive) o.drive.install(s);
  const r = { s: s, docProps: loaded.docProps, scriptProps: loaded.scriptProps, drive: o.drive || null, portalRequests: [], requests: [], dialogs: [], alerts: [], logs: [] };
  const body = makeFakeDocumentBody(s);
  s.Logger = { log: (m) => r.logs.push(String(m)) };
  s.DocumentApp.getActiveDocument = () => ({ getId: () => REPORT_ID, getBody: () => body });
  s.DocumentApp.getUi = () => ({ alert: (...args) => { r.alerts.push(args); return 'OK'; }, showModalDialog: (out, title) => r.dialogs.push({ html: out.html, title: title }),
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { YES: 'YES', NO: 'NO' } });
  s.HtmlService = { createHtmlOutput: (html) => { const out = { html: html, setWidth: () => out, setHeight: () => out }; return out; } };
  s.fetchMeetingMetadataById_ = (id) => { r.portalRequests.push(String(id)); return { statusCode: 200, text: JSON.stringify((o.meetings || {})[String(id)] || []) }; };
  s.fetchMeetingTdocListById_ = () => ({ statusCode: 200, text: 'stub' });
  s.parseMeetingTdocListHtml_ = () => o.rows || [];
  s.UrlFetchApp = { fetch: (url, opts) => {
    r.requests.push(String(url));
    const answer = o.network ? o.network(String(url), opts || {}) : null;
    if (!answer) throw new Error('network down: ' + url);
    return { getResponseCode: () => answer.status, getContentText: () => answer.text || '' };
  } };

  /** Opens Configure Meeting and runs its client script. Returns the page: fields, buttons, what was sent and shown. */
  r.open = () => {
    s.configureMeetingSettings();
    const html = r.dialogs[r.dialogs.length - 1].html;
    const script = html.slice(html.indexOf('<script>') + 8, html.lastIndexOf('</script>')).replace(/\\\\/g, '\\');
    const fields = {};
    const element = (id) => ({ id: id, value: '', textContent: '', innerHTML: '', checked: false, disabled: false, style: {}, className: '' });
    const attr = (id, name) => { const m = html.match(new RegExp('id="' + id + '"[^>]*\\s' + name + '="([^"]*)"')); return m ? m[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&') : null; };
    FIELD_IDS.forEach((id) => { fields[id] = element(id); const v = attr(id, 'value'); if (v !== null) fields[id].value = v; });
    if (/id="collectionStartField"/.test(html)) {
      const f = element('collectionStartField');
      f.value = attr('collectionStartField', 'value') || '';
      f.touched = false;
      f.getAttribute = (n) => attr('collectionStartField', n);
      fields.collectionStartField = f;
    }
    // The Shared Minutes section of a template report: what it shows when the dialog opens.
    if (/id="sharedMinutesView"/.test(html)) {
      fields.sharedMinutesView = element('sharedMinutesView');
      fields.sharedMinutesView.value = (attr('sharedMinutesView', 'value') || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
    }
    const selectedFamily = html.match(/<option value="([A-Za-z0-9]+)" selected>/);
    if (selectedFamily) fields.reportType.value = selectedFamily[1];
    const selectedType = html.match(/id="meetingType"[\s\S]*?<option value="(main|adhoc)" selected>/);
    if (selectedType) fields.meetingType.value = selectedType[1];
    // Check boxes and radio buttons that the page has: present, and ticked as the HTML says.
    (html.match(/<input type="(?:checkbox|radio)"[^>]*>/g) || []).forEach((tag) => {
      const id = (tag.match(/\sid="([^"]+)"/) || [])[1];
      if (!id) return;
      if (!fields[id]) fields[id] = element(id);
      fields[id].checked = /\schecked(?=[\s>])/.test(tag);
      fields[id].inPage = true;
    });
    const page = { html: html, fields: fields, alerts: [], sent: [], called: [], closed: false };
    const rpc = { resolveMeetingForConfigDialog: s.resolveMeetingForConfigDialog, discoverAgendaForConfigDialog: s.discoverAgendaForConfigDialog, saveConfigurationSettings: s.saveConfigurationSettings };
    // The buttons of the Shared Minutes section, where the bundle has them. Every call is recorded (page.called).
    ['createSharedMinutes', 'verifySharedMinutes', 'forgetSharedMinutes'].forEach((name) => { if (typeof s[name] === 'function') rpc[name] = s[name]; });
    const run = () => {
      let ok = () => {};
      let failed = () => {};
      const call = { withSuccessHandler: (fn) => { ok = fn; return call; }, withFailureHandler: (fn) => { failed = fn; return call; } };
      Object.keys(rpc).forEach((name) => {
        call[name] = (...args) => {
          if (name === 'saveConfigurationSettings') page.sent.push(plain(args[0]));
          page.called.push(name);
          let out;
          try { out = rpc[name](...args.map(plain)); } catch (e) { failed(e.message); return; }
          ok(out === undefined ? null : plain(out));
        };
      });
      return call;
    };
    const globals = {
      // As in a browser: a control the page does not have is null (the fields the dialog script always expects exist).
      document: { getElementById: (id) => fields[id] || (/^(apiTokenTarget|clearPersonalApiToken|promoteApiToken)/.test(id) ? null : (fields[id] = element(id))) },
      google: { script: { get run() { return run(); }, host: { close: () => { page.closed = true; } } } },
      alert: (m) => page.alerts.push(String(m)), console: { log: () => {} }
    };
    vm.createContext(globals);
    vm.runInContext(script, globals);
    page.resolve = (meetingId) => { fields.meetingId.value = String(meetingId); globals.resolveMeeting(); return page; };
    page.save = () => { globals.saveConfig(); return page; };
    /** Calls a function of the dialog script, as a button of the page does. */
    page.press = (name) => { globals[name](); return page; };
    return page;
  };
  r.props = (keys) => keys.map((k) => loaded.docProps.getProperty(k));
  return r;
}

module.exports = { configurableReport, RELEASE, REPORT_ID, TEMPLATE_ID, plain };
