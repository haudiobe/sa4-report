/**
 * TEMPLATE-001 -- copy-semantics probe (the ONE manual Google experiment).
 * See docs/SA4_REPORT_TEMPLATE_ARCHITECTURE.md §19.
 *
 * Question: can a Google Doc with a bound Apps Script be copied -- by hand
 * and by script -- so that the copy is an independent document with its own
 * usable bound script?
 *
 * HOW TO RUN (nothing else is needed; no services to add, no manifest edit):
 *   1. Create a new, empty scratch Google Doc.
 *   2. Extensions > Apps Script: replace everything with this file, save.
 *   3. Reload the Doc. A "Report" menu appears.
 *   4. Report > Initialize TEMPLATE-001   (allow the permission prompt, then
 *      click the item again)
 *   5. Report > Create Programmatic Copy  (open the link it shows)
 *   6. File > Make a copy                 (the manual copy)
 *   7. In EACH of the two copies: Report > Install 30-Minute Test Trigger
 *      (allow the permission prompt if asked, then click again).
 *   8. Copy the text of the three status boxes (original + two copies) and
 *      send them back. "Report > TEMPLATE-001 Status" shows a box again.
 *
 * Later (any time after ~35 minutes, optional): TEMPLATE-001 Status in a
 * copy shows whether the test trigger ran. The trigger removes itself after
 * its first run. When finished: Report > Clean Up Test State in each of the
 * three documents, then move them to the trash.
 *
 * SAFETY: touches only the document it runs in and the one copy it creates.
 * Every property it writes starts with TEMPLATE001_. Its only trigger calls
 * template001TriggerTick(), which records a timestamp and deletes itself.
 * It never sends mail, never fetches a URL, never lists other files.
 */

var T001 = {
  DOC_MARKER: 'TEMPLATE001_DOCUMENT_MARKER',
  SCRIPT_MARKER: 'TEMPLATE001_SCRIPT_MARKER',
  USER_MARKER: 'TEMPLATE001_USER_MARKER',
  WRITE_TEST: 'TEMPLATE001_WRITE_TEST',
  FIRST_SEEN: 'TEMPLATE001_FIRST_SEEN_TRIGGERS',
  TRIGGER_INFO: 'TEMPLATE001_TRIGGER_INFO',
  LAST_RUN: 'TEMPLATE001_LAST_TRIGGER_RUN',
  TRIGGER_DOCPROP: 'TEMPLATE001_TRIGGER_DOCPROP_TEST',
  KEY_PREFIX: 'TEMPLATE001_',
  HANDLER: 'template001TriggerTick',
  BODY_MARKER: 'TEMPLATE-001 ORIGINAL',
  COPY_MARKER: 'TEMPLATE001-PROGRAMMATIC-COPY',
  ORIGINAL_DESCRIPTION: 'TEMPLATE001 description of the ORIGINAL'
};

function onOpen() {
  DocumentApp.getUi().createMenu('Report')
    .addItem('Initialize TEMPLATE-001 (original only)', 'template001Initialize')
    .addItem('Create Programmatic Copy (original only)', 'template001CreateProgrammaticCopy')
    .addSeparator()
    .addItem('TEMPLATE-001 Status', 'template001Status')
    .addItem('Install 30-Minute Test Trigger', 'template001InstallTrigger')
    .addSeparator()
    .addItem('Clean Up Test State', 'template001CleanUp')
    .addToUi();
}

// ------------------------------------------------------------------
// Menu actions
// ------------------------------------------------------------------

/** Original only: writes the three markers, the body marker, and one trigger a copy could inherit. */
function template001Initialize() {
  const obs = template001Collect_();
  if (obs.origin && obs.origin.doc !== obs.docId) {
    template001ShowText_('TEMPLATE-001', 'This document is a COPY. Initialize runs in the original only.\n' +
      'Here, use: Report > Install 30-Minute Test Trigger.');
    return;
  }
  const stamp = template001Stamp_(obs.docId, obs.scriptId, { note: 'set by Initialize in the original' });
  PropertiesService.getDocumentProperties().setProperty(T001.DOC_MARKER, stamp);
  PropertiesService.getScriptProperties().setProperty(T001.SCRIPT_MARKER, stamp);
  PropertiesService.getUserProperties().setProperty(T001.USER_MARKER, stamp);
  if (!obs.origin) {
    // The body is the one thing every copy certainly carries: it tells a
    // copy who the original was.
    DocumentApp.getActiveDocument().getBody()
      .appendParagraph(T001.BODY_MARKER + ' doc=' + obs.docId + ' script=' + obs.scriptId);
  }
  DriveApp.getFileById(obs.docId).setDescription(T001.ORIGINAL_DESCRIPTION);
  template001CreateTrigger_(obs);
  template001ShowStatus_('ORIGINAL INITIALIZED. Next: Create Programmatic Copy, then File > Make a copy.');
}

/**
 * The exact copy mechanism of template/ReportCreator.js (copyTemplate +
 * setDescription): DriveApp only, same folder, description written AFTER
 * the copy. Nothing is injected into the copy's script.
 */
function template001CreateProgrammaticCopy() {
  const obs = template001Collect_();
  if (!obs.origin || obs.origin.doc !== obs.docId) {
    template001ShowText_('TEMPLATE-001', 'Run this in the ORIGINAL, after Report > Initialize TEMPLATE-001.');
    return;
  }
  const template = DriveApp.getFileById(obs.docId);
  const title = template.getName() + ' (PROGRAMMATIC COPY)';
  const parents = template.getParents();
  const copy = parents.hasNext() ? template.makeCopy(title, parents.next()) : template.makeCopy(title);
  let descriptionRightAfterCopy = null;
  try { descriptionRightAfterCopy = copy.getDescription(); } catch (e) { /* informational only */ }
  DriveApp.getFileById(copy.getId()).setDescription(
    T001.COPY_MARKER + ' target=' + copy.getId() + ' source=' + obs.docId +
    ' descriptionRightAfterCopy=' + JSON.stringify(descriptionRightAfterCopy));

  const url = copy.getUrl();
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<p><b>Programmatic copy created.</b></p>' +
    '<p><a href="' + template001Escape_(url) + '" target="_blank">Open the programmatic copy</a></p>' +
    '<p>In the copy: reload if the <b>Report</b> menu is missing, then<br>' +
    'Report &gt; Install 30-Minute Test Trigger.</p>' +
    '<p style="color:#666;font-size:12px">' + template001Escape_(url) + '</p></div>';
  DocumentApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(230), 'TEMPLATE-001');
}

function template001Status() {
  template001ShowStatus_('');
}

function template001InstallTrigger() {
  const obs = template001Collect_();
  const info = template001CreateTrigger_(obs);
  template001ShowStatus_(info.ok ? 'TRIGGER CREATED SUCCESSFULLY' : 'TRIGGER CREATION FAILED: ' + info.error);
}

/** Removes every probe trigger and every TEMPLATE001_ property of this document/project. */
function template001CleanUp() {
  const removed = template001DeleteProbeTriggers_();
  let keys = 0;
  [PropertiesService.getDocumentProperties(), PropertiesService.getScriptProperties(), PropertiesService.getUserProperties()]
    .forEach(function (store) {
      store.getKeys().forEach(function (key) {
        if (key.indexOf(T001.KEY_PREFIX) === 0) { store.deleteProperty(key); keys++; }
      });
    });
  template001ShowText_('TEMPLATE-001 clean-up', 'Removed ' + removed + ' test trigger(s) and ' + keys +
    ' TEMPLATE001_ properties here.\n\nRun this in each of the three documents, then move them to the trash.');
}

/**
 * The 30-minute test trigger. Records what a time-driven run of a bound
 * script can see (the container document, its Document Properties), then
 * removes itself so nothing keeps running.
 */
function template001TriggerTick() {
  const scriptProps = PropertiesService.getScriptProperties();
  const scriptId = ScriptApp.getScriptId();
  const run = { activeDoc: null, docPropsWritable: false };
  try {
    const doc = DocumentApp.getActiveDocument();
    run.activeDoc = doc ? doc.getId() : null;
  } catch (e) {
    run.activeDocError = String(e && e.message || e);
  }
  try {
    PropertiesService.getDocumentProperties().setProperty(T001.TRIGGER_DOCPROP, new Date().toISOString());
    run.docPropsWritable = true;
  } catch (e) {
    run.docPropsError = String(e && e.message || e);
  }
  const installed = template001ParseStamp_(scriptProps.getProperty(T001.TRIGGER_INFO));
  scriptProps.setProperty(T001.LAST_RUN,
    template001Stamp_(run.activeDoc || (installed ? installed.doc : null), scriptId, run));
  template001DeleteProbeTriggers_();
}

// ------------------------------------------------------------------
// Observation (service calls) and evaluation (pure)
// ------------------------------------------------------------------

function template001Stamp_(docId, scriptId, extra) {
  const stamp = { doc: docId, script: scriptId, at: new Date().toISOString() };
  Object.keys(extra || {}).forEach(function (k) { stamp[k] = extra[k]; });
  return JSON.stringify(stamp);
}

function template001ParseStamp_(raw) {
  if (raw === null || raw === undefined || raw === '') return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : { unreadable: String(raw) };
  } catch (e) {
    return { unreadable: String(raw) };
  }
}

/** A stamped value written by THIS document's own project, else null. */
function template001Own_(raw, docId, scriptId) {
  const stamp = template001ParseStamp_(raw);
  return stamp && stamp.doc === docId && stamp.script === scriptId ? stamp : null;
}

function template001ProbeTriggers_() {
  return ScriptApp.getProjectTriggers().map(function (t) {
    return { handler: t.getHandlerFunction(), id: t.getUniqueId() };
  });
}

function template001DeleteProbeTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === T001.HANDLER) { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

function template001CreateTrigger_(obs) {
  let info;
  try {
    template001DeleteProbeTriggers_();
    const trigger = ScriptApp.newTrigger(T001.HANDLER).timeBased().everyMinutes(30).create();
    info = { ok: true, handler: T001.HANDLER, id: trigger.getUniqueId(), every: 'everyMinutes(30)' };
  } catch (e) {
    info = { ok: false, handler: T001.HANDLER, error: String(e && e.message || e) };
  }
  PropertiesService.getScriptProperties().setProperty(T001.TRIGGER_INFO, template001Stamp_(obs.docId, obs.scriptId, info));
  return info;
}

/**
 * Everything the status box reports. The first call in a document also
 * records which triggers existed BEFORE the probe created any here -- that
 * is the "inherited triggers" answer for a copy.
 */
function template001Collect_() {
  const doc = DocumentApp.getActiveDocument();
  const docId = doc.getId();
  const scriptId = ScriptApp.getScriptId();
  const docProps = PropertiesService.getDocumentProperties();
  const scriptProps = PropertiesService.getScriptProperties();
  const obs = { docId: docId, scriptId: scriptId, name: doc.getName(), checkedAt: new Date().toISOString() };

  const m = doc.getBody().getText().match(/TEMPLATE-001 ORIGINAL doc=(\S+) script=(\S+)/);
  obs.origin = m ? { doc: m[1], script: m[2] } : null;

  try {
    obs.description = { ok: true, text: String(DriveApp.getFileById(docId).getDescription() || '') };
  } catch (e) {
    obs.description = { ok: false, error: String(e && e.message || e) };
  }

  obs.docMarker = docProps.getProperty(T001.DOC_MARKER);
  obs.scriptMarker = scriptProps.getProperty(T001.SCRIPT_MARKER);
  obs.userMarker = PropertiesService.getUserProperties().getProperty(T001.USER_MARKER);

  try {
    const value = template001Stamp_(docId, scriptId, {});
    docProps.setProperty(T001.WRITE_TEST, value);
    obs.writeTest = { ok: docProps.getProperty(T001.WRITE_TEST) === value };
  } catch (e) {
    obs.writeTest = { ok: false, error: String(e && e.message || e) };
  }

  obs.triggersNow = template001ProbeTriggers_();
  let firstSeen = template001Own_(scriptProps.getProperty(T001.FIRST_SEEN), docId, scriptId);
  if (!firstSeen) {
    scriptProps.setProperty(T001.FIRST_SEEN, template001Stamp_(docId, scriptId, { triggers: obs.triggersNow }));
    firstSeen = template001Own_(scriptProps.getProperty(T001.FIRST_SEEN), docId, scriptId);
  }
  obs.inheritedTriggers = firstSeen ? firstSeen.triggers : null;
  obs.triggerInfo = template001Own_(scriptProps.getProperty(T001.TRIGGER_INFO), docId, scriptId);
  const lastRun = template001ParseStamp_(scriptProps.getProperty(T001.LAST_RUN));
  obs.lastRun = lastRun && lastRun.script === scriptId ? lastRun : null;
  return obs;
}

function template001Role_(obs) {
  if (!obs.origin) return 'NOT INITIALIZED';
  if (obs.origin.doc === obs.docId) return 'ORIGINAL';
  // The creator's description names its target; the title is the fallback
  // signal in case the description turns out not to be readable.
  return template001HasCreatorDescription_(obs) || String(obs.name || '').indexOf('(PROGRAMMATIC COPY)') !== -1
    ? 'PROGRAMMATIC COPY' : 'MANUAL COPY';
}

function template001HasCreatorDescription_(obs) {
  return !!(obs.description && obs.description.ok &&
    obs.description.text.indexOf(T001.COPY_MARKER + ' target=' + obs.docId) !== -1);
}

/** MISSING / SET IN THIS DOCUMENT / COPIED FROM ANOTHER DOCUMENT. */
function template001MarkerState_(raw, obs) {
  const stamp = template001ParseStamp_(raw);
  if (!stamp) return 'MISSING';
  if (stamp.doc === obs.docId && stamp.script === obs.scriptId) return 'SET IN THIS DOCUMENT';
  return 'PRESENT / COPIED (written in doc ' + stamp.doc + ')';
}

/**
 * Pure: observations -> classified lines + overall result.
 * PASS / FAIL decide GO; EXPECTED is the predicted, harmless outcome;
 * WARNING is an outcome the design can absorb but must be told about.
 */
function template001Evaluate_(obs) {
  const role = template001Role_(obs);
  const isCopy = role === 'PROGRAMMATIC COPY' || role === 'MANUAL COPY';
  const lines = [];
  const add = function (verdict, text) { lines.push({ verdict: verdict, text: text }); };

  add(isCopy ? 'PASS' : 'INFO', 'Bound script present, Report menu works (this box proves it)');

  const origin = obs.origin;
  add(!isCopy ? 'INFO' : (obs.docId !== origin.doc ? 'PASS' : 'FAIL'),
    'Document ID: ' + obs.docId + (isCopy ? ' (original: ' + origin.doc + ')' : ''));
  add(!isCopy ? 'INFO' : (obs.scriptId !== origin.script ? 'PASS' : 'FAIL'),
    'Script ID: ' + obs.scriptId + (isCopy ? (obs.scriptId !== origin.script ? ' -- own project' : ' -- SAME project as the original') +
      ' (original: ' + origin.script + ')' : ''));
  add('INFO', 'Template marker in body: ' + (origin ? 'original doc ' + origin.doc : 'MISSING (not initialized)'));

  [['Document Property marker', obs.docMarker], ['Script Property marker', obs.scriptMarker], ['User Property marker', obs.userMarker]]
    .forEach(function (pair) {
      const state = template001MarkerState_(pair[1], obs);
      add(!isCopy ? 'INFO' : (state === 'MISSING' ? 'EXPECTED' : 'WARNING'), pair[0] + ': ' + state);
    });

  add(obs.writeTest.ok ? (isCopy ? 'PASS' : 'INFO') : 'FAIL',
    'Document Properties usable here (new value written + read back): ' +
    (obs.writeTest.ok ? 'YES' : 'NO' + (obs.writeTest.error ? ' -- ' + obs.writeTest.error : '')));

  if (isCopy) {
    const inherited = obs.inheritedTriggers || [];
    add(inherited.length === 0 ? 'PASS' : 'FAIL',
      'Triggers inherited from the original: ' + (inherited.length === 0 ? 'NONE'
        : inherited.map(function (t) { return t.handler + ' [' + t.id + ']'; }).join(', ')));
  }

  if (!obs.description.ok) {
    add(isCopy ? 'WARNING' : 'INFO', 'Drive description: NOT READABLE -- ' + obs.description.error);
  } else if (role === 'PROGRAMMATIC COPY') {
    add(template001HasCreatorDescription_(obs) ? 'PASS' : 'WARNING',
      'Drive description written by the creator: ' + (template001HasCreatorDescription_(obs) ? 'READABLE here -- ' : 'NOT FOUND here -- ') +
      (obs.description.text || 'EMPTY'));
  } else if (role === 'MANUAL COPY') {
    add('INFO', 'Drive description: ' + (obs.description.text
      ? (obs.description.text === T001.ORIGINAL_DESCRIPTION ? 'COPIED from the original' : obs.description.text) : 'EMPTY (not copied)'));
  } else {
    add('INFO', 'Drive description: ' + (obs.description.text || 'EMPTY'));
  }

  const info = obs.triggerInfo;
  if (!info) {
    add(isCopy ? 'TODO' : 'INFO', '30-minute trigger: NOT TESTED YET -- Report > Install 30-Minute Test Trigger');
  } else if (info.ok) {
    add(isCopy ? 'PASS' : 'INFO', '30-minute trigger: CREATED -- handler ' + info.handler + ', id ' + info.id + ', created ' + info.at);
  } else {
    add('FAIL', '30-minute trigger: CREATION FAILED -- ' + info.error);
  }
  add('INFO', 'Installed test triggers now: ' + (obs.triggersNow.length
    ? obs.triggersNow.map(function (t) { return t.handler + ' [' + t.id + ']'; }).join(', ')
    : 'NONE' + (obs.lastRun ? ' (the test trigger removed itself after running)' : '')));

  const run = obs.lastRun;
  if (!run) {
    add(info && info.ok ? 'EXPECTED' : 'INFO', 'Last trigger execution: NOT YET (check again ~35 minutes after creation)');
  } else {
    const sawThisDoc = run.activeDoc === obs.docId;
    add(!isCopy ? 'INFO' : (sawThisDoc && run.docPropsWritable ? 'PASS' : 'WARNING'),
      'Last trigger execution: ' + run.at + ' -- saw document ' + (run.activeDoc || 'NONE') +
      (sawThisDoc ? ' (this one)' : ' (NOT this one)') + ', Document Properties writable: ' + (run.docPropsWritable ? 'YES' : 'NO'));
  }

  const count = function (v) { return lines.filter(function (l) { return l.verdict === v; }).length; };
  let result;
  if (role === 'NOT INITIALIZED') {
    result = 'NOT INITIALIZED -- run Report > Initialize TEMPLATE-001 in the original first';
  } else if (role === 'ORIGINAL') {
    result = 'ORIGINAL -- baseline only, no verdict';
  } else if (count('FAIL') > 0) {
    result = (role === 'PROGRAMMATIC COPY' ? 'NO-GO' : 'MANUAL COPY PROBLEM') + ' (' + count('FAIL') + ' FAIL)';
  } else if (count('TODO') > 0) {
    result = 'INCOMPLETE -- install the 30-minute test trigger';
  } else {
    result = (role === 'PROGRAMMATIC COPY' ? 'GO' : 'MANUAL COPY OK') + (count('WARNING') ? ' (with ' + count('WARNING') + ' WARNING)' : '');
  }
  return { role: role, lines: lines, result: result };
}

function template001FormatStatus_(obs, headline) {
  const ev = template001Evaluate_(obs);
  const pad = function (v) { return ('[' + v + ']' + '          ').slice(0, 11); };
  return ['TEMPLATE-001 STATUS -- ' + ev.role, 'Checked: ' + obs.checkedAt]
    .concat(headline ? ['>> ' + headline] : [])
    .concat([''])
    .concat(ev.lines.map(function (l) { return pad(l.verdict) + l.text; }))
    .concat(['', 'RESULT: ' + ev.result])
    .join('\n');
}

// ------------------------------------------------------------------
// Dialogs
// ------------------------------------------------------------------

function template001Escape_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function template001ShowText_(title, text) {
  const html = '<textarea readonly onclick="this.select()" style="width:100%;height:92%;box-sizing:border-box;' +
    'font-family:Consolas,monospace;font-size:12px;white-space:pre">' + template001Escape_(text) + '</textarea>' +
    '<div style="font-family:Arial,sans-serif;font-size:11px;color:#666">Click the text to select all, then copy.</div>';
  DocumentApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(760).setHeight(430), title);
}

function template001ShowStatus_(headline) {
  template001ShowText_('TEMPLATE-001 Status', template001FormatStatus_(template001Collect_(), headline));
}
