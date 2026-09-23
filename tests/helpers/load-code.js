/**
 * Shared VM-sandbox loader for testing pure/near-pure functions in Code.js
 * outside Apps Script.
 *
 * Code.js targets Google Apps Script, so it is loaded into a VM sandbox with
 * minimal stubs for the Apps Script globals (same approach as
 * tests/revision-order.test.js). The functions under test are pure or
 * near-pure JavaScript, so they run unchanged.
 *
 * This module only loads Code.js; it does not modify or wrap any of its
 * functions. Each call returns a fresh sandbox, so tests cannot leak state
 * (e.g. Document/Script Properties) into one another.
 *
 * NOTE: this intentionally does NOT provide a working DocumentApp. Any
 * function that calls DocumentApp.getActiveDocument() (or otherwise touches
 * a real document body) will throw when invoked through this loader. That is
 * by design for this baseline: see tests/README.md for what is and is not
 * covered.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const CODE_JS_PATH = path.join(__dirname, '..', '..', 'Code.js');

/**
 * In-memory stand-in for PropertiesService.getDocumentProperties() /
 * getScriptProperties(). Starts empty (every getProperty() returns null),
 * matching a freshly-created Doc with no configuration saved yet -- this is
 * what exercises the DEFAULT values baked into Code.js.
 */
function makePropertyStore(initial) {
  const store = Object.assign({}, initial || {});
  return {
    getProperty: (key) => (Object.prototype.hasOwnProperty.call(store, key) ? store[key] : null),
    setProperty: (key, value) => { store[key] = value; },
    deleteProperty: (key) => { delete store[key]; },
    getKeys: () => Object.keys(store),
    _store: store
  };
}

function loadCode(options) {
  const opts = options || {};
  const docProps = makePropertyStore(opts.documentProperties);
  const scriptProps = makePropertyStore(opts.scriptProperties);

  const sandbox = {
    Logger: { log: () => {} },
    PropertiesService: {
      getDocumentProperties: () => docProps,
      getScriptProperties: () => scriptProps
    },
    DocumentApp: {
      getActiveDocument: () => { throw new Error('DocumentApp.getActiveDocument() is not available in the pure-logic test sandbox'); },
      getUi: () => { throw new Error('DocumentApp.getUi() is not available in the pure-logic test sandbox'); },
      openById: () => { throw new Error('DocumentApp.openById() is not available in the pure-logic test sandbox'); },
      ParagraphHeading: { NORMAL: 'NORMAL', TITLE: 'TITLE', HEADING1: 'H1', HEADING2: 'H2', HEADING3: 'H3' },
      ElementType: { PARAGRAPH: 'PARAGRAPH', TABLE: 'TABLE', LIST_ITEM: 'LIST_ITEM', INLINE_IMAGE: 'IMG' }
    },
    Session: { getScriptTimeZone: () => 'Europe/Berlin' },
    Utilities: {
      formatDate: (date, tz, fmt) => date.toISOString(),
      unzip: () => { throw new Error('Utilities.unzip() is not available in the pure-logic test sandbox'); }
    },
    UrlFetchApp: { fetch: () => { throw new Error('UrlFetchApp.fetch() is not available in the pure-logic test sandbox'); } },
    DriveApp: {},
    SpreadsheetApp: {},
    HtmlService: {},
    ScriptApp: {},
    MimeType: {},
    XmlService: {},
    // PERF-003B: default fake lock always succeeds and no-ops release, so
    // any code path that happens to call LockService in a test that isn't
    // specifically exercising lock behavior doesn't crash. Tests that DO
    // exercise lock behavior (tryLock failure, release-on-error, etc.)
    // override sandbox.LockService with their own controllable fake.
    LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    console
  };

  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(CODE_JS_PATH, 'utf8'), sandbox, { filename: 'Code.js' });

  return { sandbox, docProps, scriptProps };
}

module.exports = { loadCode, CODE_JS_PATH };
