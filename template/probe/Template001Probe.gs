/**
 * TEMPLATE-001 -- copy-semantics probe (the ONE manual experiment).
 * See docs/SA4_REPORT_TEMPLATE_ARCHITECTURE.md §19 and §20.
 *
 * Paste this whole file into the bound script of a NEW, EMPTY scratch
 * Google Doc (Extensions > Apps Script). It touches only that scratch doc
 * and the copies it makes. It never opens, lists or changes any other file,
 * never sends mail, and its only trigger calls probeTick(), which just
 * appends a timestamp to a Document Property.
 *
 * Before step 1, in the script editor add Services > Drive API (v3), so the
 * report shows whether the advanced-service setting travels with a copy.
 *
 * Steps (in the ORIGINAL scratch doc, after reloading it):
 *   1. 🧪 Probe > 1. Mark original        (authorize when asked)
 *   2. 🧪 Probe > 2. Copy via DriveApp    (makes "…(DriveApp copy)")
 *   3. File > Make a copy                 (makes "Copy of …", by hand)
 * Then in EACH copy:
 *   4. reload, check the 🧪 Probe menu appears WITHOUT authorizing (onOpen)
 *   5. 🧪 Probe > 3. Report               (note: authorization prompt? y/n)
 *   6. 🧪 Probe > 4. Start 30-min trigger, wait ≥ 35 min, 3. Report again
 *   7. 🧪 Probe > 5. Clean up (deletes this copy's probe trigger)
 * Finally run 5. Clean up in the original, and trash all probe docs.
 */

function onOpen() {
  DocumentApp.getUi().createMenu('🧪 Probe')
    .addItem('1. Mark original', 'probeMarkOriginal')
    .addItem('2. Copy via DriveApp', 'probeCopyViaDriveApp')
    .addItem('3. Report', 'probeReport')
    .addItem('4. Start 30-min trigger', 'probeStartTrigger')
    .addItem('5. Clean up', 'probeCleanUp')
    .addToUi();
}

function probeMarkOriginal() {
  const id = DocumentApp.getActiveDocument().getId();
  PropertiesService.getDocumentProperties().setProperty('PROBE_DOC', 'set in original ' + id);
  PropertiesService.getScriptProperties().setProperty('PROBE_SCRIPT', 'set in original script ' + ScriptApp.getScriptId());
  PropertiesService.getUserProperties().setProperty('PROBE_USER', 'set in original script ' + ScriptApp.getScriptId());
  DriveApp.getFileById(id).setDescription('PROBE description set in original ' + id);
  probeStartTrigger();
  DocumentApp.getUi().alert('Original marked (properties, description, 30-min trigger).\nScript ID: ' + ScriptApp.getScriptId());
}

function probeCopyViaDriveApp() {
  const file = DriveApp.getFileById(DocumentApp.getActiveDocument().getId());
  const copy = file.makeCopy(file.getName() + ' (DriveApp copy)');
  DocumentApp.getUi().alert('DriveApp copy created:\n' + copy.getUrl() +
    '\n\nCopy description right after makeCopy: ' + JSON.stringify(copy.getDescription()));
}

function probeReport() {
  const docId = DocumentApp.getActiveDocument().getId();
  const lines = [
    'Doc ID: ' + docId,
    'Script ID: ' + ScriptApp.getScriptId(),
    'Document property PROBE_DOC: ' + JSON.stringify(PropertiesService.getDocumentProperties().getProperty('PROBE_DOC')),
    'Script property PROBE_SCRIPT: ' + JSON.stringify(PropertiesService.getScriptProperties().getProperty('PROBE_SCRIPT')),
    'User property PROBE_USER: ' + JSON.stringify(PropertiesService.getUserProperties().getProperty('PROBE_USER')),
    'Drive description: ' + JSON.stringify(DriveApp.getFileById(docId).getDescription()),
    'Triggers in this project: ' + JSON.stringify(ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); })),
    'Trigger ticks recorded here: ' + (PropertiesService.getDocumentProperties().getProperty('PROBE_TICKS') || '(none)'),
    'Drive advanced service present: ' + (typeof Drive !== 'undefined')
  ];
  Logger.log(lines.join('\n'));
  DocumentApp.getUi().alert(lines.join('\n'));
}

function probeStartTrigger() {
  probeDeleteTriggers_();
  ScriptApp.newTrigger('probeTick').timeBased().everyMinutes(30).create();
}

function probeTick() {
  const props = PropertiesService.getDocumentProperties();
  const doc = DocumentApp.getActiveDocument();
  props.setProperty('PROBE_TICKS', (props.getProperty('PROBE_TICKS') || '') + ' ' + new Date().toISOString() +
    '@' + (doc ? doc.getId().slice(0, 6) : 'no-active-doc'));
}

function probeCleanUp() {
  const n = probeDeleteTriggers_();
  DocumentApp.getUi().alert('Deleted ' + n + ' probe trigger(s) in this project.');
}

function probeDeleteTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'probeTick') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}
