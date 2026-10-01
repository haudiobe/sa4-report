/**
 * TEMPLATE-002A: the Legacy repository's tests/helpers/load-code.js at
 * b95e06e, unchanged below. Used only by tests/legacy-parity-*.test.js. It
 * loads THIS repository's Code.js (CODE_JS_PATH is relative) with the richer
 * Utilities/Session fakes the Legacy suites expect.
 */

/**
 * Shared VM-sandbox loader for testing pure/near-pure functions in this
 * LEGACY project's Code.js outside Apps Script.
 *
 * This is the central add-on repo's tests/helpers/load-code.js, adapted
 * only in CODE_JS_PATH (points at THIS repo's Code.js) -- the loading
 * approach itself needs no legacy-specific change: it is already generic
 * (no getReportStateStore_/registry/scheduler assumptions), which is
 * exactly why it can be reused unmodified for a bound-script codebase that
 * has none of that central-only infrastructure.
 *
 * This module only loads Code.js; it does not modify or wrap any of its
 * functions. Each call returns a fresh sandbox, so tests cannot leak state
 * (e.g. Document/Script Properties) into one another.
 *
 * NOTE: this intentionally does NOT provide a working DocumentApp. Any
 * function that calls DocumentApp.getActiveDocument() (or otherwise touches
 * a real document body) will throw when invoked through this loader.
 *
 * This test harness is NOT part of the Apps Script deployment payload --
 * see the project's .claspignore, which excludes everything except the
 * script/HTML source files actually pushed.
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const CODE_JS_PATH = path.join(__dirname, '..', '..', 'Code.js');

/**
 * In-memory stand-in for PropertiesService.getDocumentProperties() /
 * getScriptProperties(). Starts empty (every getProperty() returns null),
 * matching a freshly-created Doc with no configuration saved yet.
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
    Session: {
      getScriptTimeZone: () => 'Europe/Berlin',
      // LEGACY-UPGRADE-006: real Apps Script sometimes has no active user
      // in a given execution context -- default here matches that (empty
      // string, never a fabricated address); a test overrides this to
      // exercise the "From" header being present.
      getActiveUser: () => ({ getEmail: () => '' })
    },
    // LEGACY-UPGRADE-006: newBlob()/base64Encode() are genuinely
    // implemented (via Node's Buffer) rather than stubbed, so the
    // quoted-printable/MIME pure functions can be tested for real, exactly
    // like they will run in Apps Script (both are byte-for-byte UTF-8
    // operations with no other platform dependency).
    Utilities: {
      // LEGACY-UPGRADE-006D: a genuine, minimal token-based formatter
      // (yyyy/MM/dd/HH/mm/ss/MMM -- every token this codebase's own
      // Utilities.formatDate() calls actually use), using the Date
      // object's own UTC fields for determinism regardless of the
      // machine/timezone running the tests. `tz` is intentionally unused
      // (a plain Date carries no timezone of its own to convert from) --
      // every caller in this codebase only needs deterministic digit
      // formatting for test purposes, not a real timezone conversion.
      formatDate: (date, tz, fmt) => {
        const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const pad = (n) => String(n).padStart(2, '0');
        const y = date.getUTCFullYear(), mo = date.getUTCMonth(), d = date.getUTCDate();
        const h = date.getUTCHours(), mi = date.getUTCMinutes(), s = date.getUTCSeconds();
        return String(fmt)
          .replace(/yyyy/g, String(y))
          .replace(/MMM/g, MONTHS[mo])
          .replace(/MM/g, pad(mo + 1))
          .replace(/dd/g, pad(d))
          .replace(/HH/g, pad(h))
          .replace(/mm/g, pad(mi))
          .replace(/ss/g, pad(s));
      },
      unzip: () => { throw new Error('Utilities.unzip() is not available in the pure-logic test sandbox'); },
      // LEGACY-UPGRADE-006D: a lightweight fake, NOT a real ZIP encoder --
      // this sandbox has no need to produce genuine ZIP bytes (nothing
      // decompresses it), only to let a test verify WHICH blobs/filenames
      // production code actually asked to be archived. Records each input
      // blob's own name (via its real getName()) on a `_zipEntryNames`
      // property (test-introspection only; not a real Apps Script Blob
      // property) so a test can assert on ZIP membership without parsing
      // binary ZIP data.
      zip: (blobs, name) => {
        const entryNames = (blobs || []).map((b) => (typeof b.getName === 'function' ? b.getName() : ''));
        let blobName = name || 'archive.zip';
        return {
          getBytes: () => [],
          getName: () => blobName,
          setName: (n) => { blobName = n; },
          getContentType: () => 'application/zip',
          _zipEntryNames: entryNames
        };
      },
      newBlob: (data, contentType, name) => {
        // Apps Script's real Blob.getBytes() returns SIGNED bytes
        // (-128..127) -- matched here so production code's own
        // "if (b < 0) b += 256" correction is exercised by tests exactly
        // as it runs for real, not silently skipped.
        const bytes = Array.from(Buffer.from(String(data === undefined ? '' : data), 'utf8')).map((b) => (b > 127 ? b - 256 : b));
        let blobName = name || '';
        return {
          getBytes: () => bytes.slice(),
          getName: () => blobName,
          setName: (n) => { blobName = n; return this; },
          getContentType: () => contentType || 'application/octet-stream'
        };
      },
      base64Encode: (input) => {
        const buf = Array.isArray(input) ? Buffer.from(input.map((b) => (b < 0 ? b + 256 : b))) : Buffer.from(String(input), 'utf8');
        return buf.toString('base64');
      },
      // LEGACY-UPGRADE-006B (Part F diagnostic): genuinely implemented via
      // Node's crypto module (MD5 only -- the one algorithm this codebase's
      // diagnostic actually uses), returning a signed-byte array exactly
      // like real Utilities.computeDigest() does, so it round-trips through
      // the same base64Encode() above just like production.
      computeDigest: (algorithm, value) => {
        if (algorithm !== 'MD5') throw new Error('Utilities.computeDigest() sandbox stub only implements MD5');
        const crypto = require('crypto');
        const bytes = Array.from(crypto.createHash('md5').update(String(value === undefined ? '' : value), 'utf8').digest());
        return bytes.map((b) => (b > 127 ? b - 256 : b));
      },
      DigestAlgorithm: { MD5: 'MD5' }
    },
    UrlFetchApp: { fetch: () => { throw new Error('UrlFetchApp.fetch() is not available in the pure-logic test sandbox'); } },
    DriveApp: {},
    SpreadsheetApp: {},
    HtmlService: {},
    ScriptApp: {},
    MimeType: {},
    XmlService: {},
    LockService: {
      getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }),
      getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} })
    },
    console
  };

  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(CODE_JS_PATH, 'utf8'), sandbox, { filename: 'Code.js' });

  return { sandbox, docProps, scriptProps };
}

module.exports = { loadCode, CODE_JS_PATH };
