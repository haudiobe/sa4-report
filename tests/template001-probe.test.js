/**
 * TEMPLATE-001 -- the copy-semantics probe (template/probe/Template001Probe.gs).
 *
 * The probe is run by hand in real Google Docs; this test only makes sure
 * its self-diagnosis is right before Thomas spends time on it. A small fake
 * "Google" holds documents and script projects, and copies a document under
 * a chosen copy semantics (the expected one, and each bad outcome), so every
 * verdict line and the GO / NO-GO result are exercised.
 *
 * Run: node tests/template001-probe.test.js
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

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

const PROBE_PATH = path.join(__dirname, '..', 'template', 'probe', 'Template001Probe.gs');
const PROBE_SRC = fs.readFileSync(PROBE_PATH, 'utf8');
const CREATOR_SRC = fs.readFileSync(path.join(__dirname, '..', 'template', 'ReportCreator.js'), 'utf8');

function makeStore(initial) {
  const store = Object.assign({}, initial || {});
  return {
    getProperty: (k) => (Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null),
    setProperty: (k, v) => { store[k] = String(v); },
    deleteProperty: (k) => { delete store[k]; },
    getKeys: () => Object.keys(store),
    _store: store
  };
}

/** Fake Google: documents, bound projects, and a configurable copy. */
function makeWorld(semantics) {
  const sem = Object.assign({ copyProperties: false, copyTriggers: false, sameScript: false, triggerCreateFails: false, descriptionUnreadable: false }, semantics || {});
  const world = { docs: {}, projects: {}, seq: 0, sem };
  const newProject = (id) => { world.projects[id] = { id, scriptProps: makeStore(), userProps: makeStore(), triggers: [] }; return world.projects[id]; };

  world.createDoc = (id, name) => {
    newProject('script-' + id);
    world.docs[id] = { id, name, body: [], description: '', docProps: makeStore(), scriptId: 'script-' + id, dialogs: [], menu: null };
    return world.docs[id];
  };

  // What Drive does when a document with a bound script is copied.
  world.copyDoc = (srcId, name) => {
    const src = world.docs[srcId];
    const srcProject = world.projects[src.scriptId];
    const id = 'doc-copy-' + (++world.seq);
    const scriptId = sem.sameScript ? src.scriptId : 'script-' + id;
    if (!sem.sameScript) {
      const p = newProject(scriptId);
      if (sem.copyProperties) {
        Object.assign(p.scriptProps._store, srcProject.scriptProps._store);
        Object.assign(p.userProps._store, srcProject.userProps._store);
      }
      if (sem.copyTriggers) p.triggers = srcProject.triggers.map((t) => ({ handler: t.handler, id: t.id + '-inherited' }));
    }
    world.docs[id] = {
      id, name, body: src.body.slice(), description: sem.copyDescription ? src.description : '',
      docProps: makeStore(sem.copyProperties ? src.docProps._store : {}), scriptId, dialogs: [], menu: null
    };
    return world.docs[id];
  };

  /** The bound script of one document, freshly loaded (one "execution"). */
  world.open = (docId, inTrigger) => {
    const doc = world.docs[docId];
    const project = world.projects[doc.scriptId];
    const fileFor = (id) => {
      const d = world.docs[id];
      return {
        getId: () => d.id, getName: () => d.name, getUrl: () => 'https://docs.google.com/document/d/' + d.id + '/edit',
        getDescription: () => { if (sem.descriptionUnreadable && docId !== 'doc-original') throw new Error('no access'); return d.description; },
        setDescription: (t) => { d.description = t; },
        getParents: () => ({ hasNext: () => true, next: () => 'folder' }),
        makeCopy: (name, folder) => { world.lastCopyArgs = [name, folder]; return fileFor(world.copyDoc(id, name).id); }
      };
    };
    const sandbox = {
      DocumentApp: {
        getActiveDocument: () => ({
          getId: () => doc.id, getName: () => doc.name,
          getBody: () => ({ getText: () => doc.body.join('\n'), appendParagraph: (t) => { doc.body.push(t); } })
        }),
        getUi: () => {
          if (inTrigger) throw new Error('no UI in a trigger');
          const menu = { items: [], addItem(label, fn) { this.items.push([label, fn]); return this; }, addSeparator() { return this; }, addToUi() { doc.menu = this; } };
          return { createMenu: (name) => { menu.name = name; return menu; }, showModalDialog: (out, title) => { doc.dialogs.push({ title, html: out.html }); } };
        }
      },
      PropertiesService: { getDocumentProperties: () => doc.docProps, getScriptProperties: () => project.scriptProps, getUserProperties: () => project.userProps },
      ScriptApp: {
        getScriptId: () => project.id,
        getProjectTriggers: () => project.triggers.map((t) => ({ getHandlerFunction: () => t.handler, getUniqueId: () => t.id, _t: t })),
        deleteTrigger: (t) => { project.triggers = project.triggers.filter((x) => x !== t._t); },
        newTrigger: (handler) => ({ timeBased: () => ({ everyMinutes: (n) => ({ create: () => {
          if (sem.triggerCreateFails && docId !== 'doc-original') throw new Error('trigger refused');
          if ([1, 5, 10, 15, 30].indexOf(n) === -1) throw new Error('invalid everyMinutes');
          const t = { handler, id: 'trg-' + (++world.seq), minutes: n };
          project.triggers.push(t);
          return { getUniqueId: () => t.id };
        } }) }) })
      },
      DriveApp: { getFileById: fileFor },
      HtmlService: { createHtmlOutput: (html) => ({ html, setWidth() { return this; }, setHeight() { return this; } }) },
      console
    };
    vm.createContext(sandbox);
    vm.runInContext(PROBE_SRC, sandbox, { filename: 'Template001Probe.gs' });
    return sandbox;
  };

  world.statusText = (docId) => {
    const html = world.docs[docId].dialogs[world.docs[docId].dialogs.length - 1].html;
    return html.replace(/^.*?<textarea[^>]*>/, '').replace(/<\/textarea>.*$/, '').replace(/&quot;/g, '"').replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&amp;/g, '&');
  };
  world.result = (docId) => (world.statusText(docId).match(/^RESULT: (.*)$/m) || [])[1];
  world.verdicts = (docId) => world.statusText(docId).split('\n').filter((l) => /^\[/.test(l))
    .map((l) => [l.slice(1, l.indexOf(']')), l.slice(11).split(/:| --/)[0]]);
  world.verdictOf = (docId, prefix) => (world.verdicts(docId).find((v) => v[1].indexOf(prefix) === 0) || [null])[0];
  return world;
}

/** Thomas's checklist up to and including the two copies. */
function runExperiment(semantics) {
  const world = makeWorld(semantics);
  world.createDoc('doc-original', 'T001 scratch');
  world.open('doc-original').onOpen();
  world.open('doc-original').template001Initialize();
  world.open('doc-original').template001CreateProgrammaticCopy();
  const programmatic = Object.keys(world.docs).find((id) => id !== 'doc-original');
  const manual = world.copyDoc('doc-original', 'Copy of T001 scratch').id;
  return { world, programmatic, manual };
}

// =============================================================== original

console.log('original: menu, markers, baseline status');
{
  const { world } = runExperiment();
  const o = world.docs['doc-original'];
  check('menu is "Report" with the five actions', [o.menu.name, o.menu.items.map((i) => i[1])],
    ['Report', ['template001Initialize', 'template001CreateProgrammaticCopy', 'template001Status', 'template001InstallTrigger', 'template001CleanUp']]);
  check('every menu handler exists as a public function', o.menu.items.every((i) => typeof world.open('doc-original')[i[1]] === 'function'), true);
  check('three harmless markers, one per store',
    [Object.keys(o.docProps._store).sort(), Object.keys(world.projects[o.scriptId].userProps._store)],
    [['TEMPLATE001_DOCUMENT_MARKER', 'TEMPLATE001_WRITE_TEST'], ['TEMPLATE001_USER_MARKER']]);
  check('body carries the original\'s ids', o.body, ['TEMPLATE-001 ORIGINAL doc=doc-original script=script-doc-original']);
  check('original owns one 30-minute trigger a copy could inherit',
    world.projects[o.scriptId].triggers.map((t) => [t.handler, t.minutes]), [['template001TriggerTick', 30]]);
  world.open('doc-original').template001Status();
  check('original has no verdict', world.result('doc-original'), 'ORIGINAL -- baseline only, no verdict');
  check('original markers read as set here', world.statusText('doc-original').match(/SET IN THIS DOCUMENT/g).length, 3);
}

// =============================================================== expected semantics

console.log('programmatic copy, expected Google behaviour: INCOMPLETE -> GO');
{
  const { world, programmatic } = runExperiment();
  check('copy made with the ReportCreator mechanism (name, same folder)', world.lastCopyArgs, ['T001 scratch (PROGRAMMATIC COPY)', 'folder']);
  check('the creator dialog shows the copy\'s link', /href="https:\/\/docs\.google\.com\/document\/d\/doc-copy-\d+\/edit"/.test(world.docs['doc-original'].dialogs.slice(-1)[0].html), true);

  world.open(programmatic).template001Status();
  check('role detected', world.statusText(programmatic).split('\n')[0], 'TEMPLATE-001 STATUS -- PROGRAMMATIC COPY');
  check('before the trigger test: INCOMPLETE, not GO', world.result(programmatic), 'INCOMPLETE -- install the 30-minute test trigger');

  world.open(programmatic).template001InstallTrigger();
  check('headline', /^>> TRIGGER CREATED SUCCESSFULLY$/m.test(world.statusText(programmatic)), true);
  check('verdict per line', world.verdicts(programmatic), [
    ['PASS', 'Bound script present, Report menu works (this box proves it)'],
    ['PASS', 'Document ID'],
    ['PASS', 'Script ID'],
    ['INFO', 'Template marker in body'],
    ['EXPECTED', 'Document Property marker'],
    ['EXPECTED', 'Script Property marker'],
    ['EXPECTED', 'User Property marker'],
    ['PASS', 'Document Properties usable here (new value written + read back)'],
    ['PASS', 'Triggers inherited from the original'],
    ['PASS', 'Drive description written by the creator'],
    ['PASS', '30-minute trigger'],
    ['INFO', 'Installed test triggers now'],
    ['EXPECTED', 'Last trigger execution']
  ]);
  check('GO without waiting for the trigger to fire', world.result(programmatic), 'GO');
  check('markers reported MISSING', world.statusText(programmatic).match(/marker: MISSING/g).length, 3);
  check('the copy\'s trigger is its own, the original\'s is untouched',
    [world.projects[world.docs[programmatic].scriptId].triggers.length, world.projects['script-doc-original'].triggers.length], [1, 1]);
}

console.log('manual copy, expected behaviour');
{
  const { world, manual } = runExperiment();
  world.open(manual).template001InstallTrigger();
  check('role detected', world.statusText(manual).split('\n')[0], 'TEMPLATE-001 STATUS -- MANUAL COPY');
  check('result', world.result(manual), 'MANUAL COPY OK');
  check('description not copied is reported, not judged', world.verdictOf(manual, 'Drive description'), 'INFO');
}

console.log('the trigger later fires in the copy and removes itself');
{
  const { world, programmatic } = runExperiment();
  world.open(programmatic).template001InstallTrigger();
  world.open(programmatic, true).template001TriggerTick();
  world.open(programmatic).template001Status();
  check('execution recorded against this document', world.verdictOf(programmatic, 'Last trigger execution'), 'PASS');
  check('saw this document, Document Properties writable',
    /saw document doc-copy-\d+ \(this one\), Document Properties writable: YES/.test(world.statusText(programmatic)), true);
  check('test trigger gone after its first run', world.projects[world.docs[programmatic].scriptId].triggers, []);
  check('still GO', world.result(programmatic), 'GO');
  check('the original\'s trigger is not touched by the copy\'s run', world.projects['script-doc-original'].triggers.length, 1);
}

// =============================================================== other outcomes

console.log('properties copied: WARNING, still GO');
{
  const { world, programmatic } = runExperiment({ copyProperties: true });
  world.open(programmatic).template001InstallTrigger();
  check('three warnings', ['Document', 'Script', 'User'].map((k) => world.verdictOf(programmatic, k + ' Property marker')), ['WARNING', 'WARNING', 'WARNING']);
  check('copied values name the original', world.statusText(programmatic).match(/PRESENT \/ COPIED \(written in doc doc-original\)/g).length, 3);
  check('a copied trigger record is not mistaken for this copy\'s own', world.verdictOf(programmatic, '30-minute trigger'), 'PASS');
  check('result', world.result(programmatic), 'GO (with 3 WARNING)');
}

console.log('triggers inherited: FAIL -> NO-GO');
{
  const { world, programmatic } = runExperiment({ copyTriggers: true });
  world.open(programmatic).template001InstallTrigger();
  check('inherited trigger named', /Triggers inherited from the original: template001TriggerTick \[trg-\d+-inherited\]/.test(world.statusText(programmatic)), true);
  check('still reported after the probe replaced it with its own', world.verdictOf(programmatic, 'Triggers inherited'), 'FAIL');
  check('result', world.result(programmatic), 'NO-GO (1 FAIL)');
}

console.log('copy shares the original\'s script project: FAIL -> NO-GO');
{
  const { world, programmatic } = runExperiment({ sameScript: true });
  world.open(programmatic).template001InstallTrigger();
  check('script id line fails', world.verdictOf(programmatic, 'Script ID'), 'FAIL');
  check('result is NO-GO', /^NO-GO/.test(world.result(programmatic)), true);
}

console.log('30-minute trigger cannot be created: FAIL -> NO-GO');
{
  const { world, programmatic } = runExperiment({ triggerCreateFails: true });
  world.open(programmatic).template001InstallTrigger();
  check('headline carries the error', /^>> TRIGGER CREATION FAILED: trigger refused$/m.test(world.statusText(programmatic)), true);
  check('result', world.result(programmatic), 'NO-GO (1 FAIL)');
}

console.log('description not readable in the copy: WARNING only, role from the title');
{
  const { world, programmatic } = runExperiment({ descriptionUnreadable: true });
  world.open(programmatic).template001InstallTrigger();
  check('still recognised as the programmatic copy', world.statusText(programmatic).split('\n')[0], 'TEMPLATE-001 STATUS -- PROGRAMMATIC COPY');
  check('result', world.result(programmatic), 'GO (with 1 WARNING)');
}

// =============================================================== guards + cleanup

console.log('guards and clean-up');
{
  const { world, programmatic } = runExperiment();
  const before = Object.keys(world.docs).length;
  world.open(programmatic).template001Initialize();
  check('Initialize refuses to run in a copy', [Object.keys(world.docs[programmatic].docProps._store).indexOf('TEMPLATE001_DOCUMENT_MARKER'), /is a COPY/.test(world.statusText(programmatic))], [-1, true]);
  world.open(programmatic).template001CreateProgrammaticCopy();
  check('a copy cannot create further copies', Object.keys(world.docs).length, before);

  const fresh = makeWorld();
  fresh.createDoc('doc-fresh', 'fresh');
  fresh.open('doc-fresh').template001Status();
  check('an uninitialized document says so', /^NOT INITIALIZED/.test(fresh.result('doc-fresh')), true);

  world.open(programmatic).template001InstallTrigger();
  world.docs[programmatic].docProps.setProperty('MEETING_ID', '86172');
  world.open(programmatic).template001CleanUp();
  const p = world.projects[world.docs[programmatic].scriptId];
  check('clean-up removes the test trigger and only TEMPLATE001_ properties',
    [p.triggers, p.scriptProps.getKeys(), world.docs[programmatic].docProps.getKeys()], [[], [], ['MEETING_ID']]);
}

console.log('the probe itself');
{
  const keys = [...PROBE_SRC.matchAll(/'(TEMPLATE001_[A-Z_]*)'/g)].map((m) => m[1]);
  check('every property key is a TEMPLATE001_ key', [keys.length > 5, [...PROBE_SRC.matchAll(/(?:set|get|delete)Property\('([^']+)'/g)].length], [true, 0]);
  check('no Advanced Drive service, no network, no mail', [/\bDrive\.[A-Z]/.test(PROBE_SRC), /UrlFetchApp|MailApp|GmailApp/.test(PROBE_SRC)], [false, false]);
  check('nothing but DriveApp.getFileById is asked of Drive', [...PROBE_SRC.matchAll(/DriveApp\.(\w+)/g)].every((m) => m[1] === 'getFileById'), true);
  const mechanism = /parents\.hasNext\(\) \? template\.makeCopy\(title, parents\.next\(\)\) : template\.makeCopy\(title\)/;
  check('same copy call as template/ReportCreator.js', [mechanism.test(PROBE_SRC), mechanism.test(CREATOR_SRC)], [true, true]);
  check('one top-level name prefix (no clash with anything pasted later)',
    [...PROBE_SRC.matchAll(/^(?:function|var)\s+([A-Za-z0-9_]+)/gm)].map((m) => m[1]).filter((n) => !/^(template001|T001$|onOpen$)/.test(n)), []);
}

// T001_SHOW=1 prints the status box of a successful programmatic copy.
if (process.env.T001_SHOW) {
  const { world, programmatic } = runExperiment();
  world.open(programmatic).template001InstallTrigger();
  console.log('\n' + world.statusText(programmatic));
}

console.log(failures ? `\n${failures} FAILURE(S)` : '\nall ok');
process.exitCode = failures ? 1 : 0;
