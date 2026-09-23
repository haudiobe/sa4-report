/**
 * ADDON-002 — report execution context: getReportDocument_() /
 * getReportBody_() / getReportDocumentId_().
 *
 * These are the smallest possible seam between "business logic wants the
 * report Document/Body" and "where that Document actually comes from" --
 * see the "ADDON-002 -- REPORT EXECUTION CONTEXT" block in Code.js for the
 * full rationale. Nothing in production calls these with an explicit
 * context yet; what's verified here is the CONTRACT the future background
 * execution path (ADDON-004+) will rely on:
 *
 *   - no context                      -> DocumentApp.getActiveDocument()
 *   - context.document supplied       -> that exact object, reused as-is
 *   - context.documentId supplied     -> DocumentApp.openById(documentId),
 *                                        called AT MOST ONCE per context
 *                                        even across many calls
 *
 * Run: node tests/addon002-context.test.js
 */

const { loadCode } = require('./helpers/load-code.js');

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

function expectThrows(name, fn, messageSubstring) {
  try {
    fn();
    failures++;
    console.log(`  FAIL ${name}\n         expected a throw, but none occurred`);
  } catch (e) {
    if (messageSubstring && e.message.indexOf(messageSubstring) === -1) {
      failures++;
      console.log(`  FAIL ${name}\n         expected message to contain "${messageSubstring}"\n         actual: ${e.message}`);
    } else {
      console.log(`  ok   ${name}`);
    }
  }
}

// -------------------------------------------------------------- no context

console.log('getReportDocument_()/getReportBody_()/getReportDocumentId_() -- no context: active document');

{
  const { sandbox } = loadCode();
  const fakeBody = { marker: 'active-body' };
  const fakeDoc = { getBody: () => fakeBody, getId: () => 'ACTIVE_DOC_ID' };
  sandbox.DocumentApp.getActiveDocument = () => fakeDoc;

  check('getReportDocument_() with no context returns the active document',
    sandbox.getReportDocument_(), fakeDoc);
  check('getReportBody_() with no context returns the active document\'s body',
    sandbox.getReportBody_(), fakeBody);
  check('getReportDocumentId_() with no context returns the active document\'s id',
    sandbox.getReportDocumentId_(), 'ACTIVE_DOC_ID');
}

{
  const { sandbox } = loadCode();
  const fakeDoc = { getBody: () => ({}), getId: () => 'X' };
  sandbox.DocumentApp.getActiveDocument = () => fakeDoc;

  check('getReportDocument_(undefined) is identical to no-arg call',
    sandbox.getReportDocument_(undefined), fakeDoc);
  check('getReportDocument_(null) is identical to no-arg call',
    sandbox.getReportDocument_(null), fakeDoc);
}

// --------------------------------------------------- explicit document object

console.log('getReportDocument_() -- explicit context.document is reused as-is, no DocumentApp call');

{
  const { sandbox } = loadCode();
  let activeDocumentCalls = 0;
  sandbox.DocumentApp.getActiveDocument = () => { activeDocumentCalls++; throw new Error('should not be called'); };
  sandbox.DocumentApp.openById = () => { throw new Error('should not be called'); };

  const explicitDoc = { getBody: () => ({ marker: 'explicit-body' }), getId: () => 'EXPLICIT_ID' };
  const context = { document: explicitDoc };

  check('getReportDocument_(context) returns context.document unchanged',
    sandbox.getReportDocument_(context), explicitDoc);
  check('getReportBody_(context) returns context.document.getBody()',
    sandbox.getReportBody_(context), { marker: 'explicit-body' });
  check('getReportDocumentId_(context) returns context.document.getId()',
    sandbox.getReportDocumentId_(context), 'EXPLICIT_ID');
  check('DocumentApp.getActiveDocument() was never called',
    activeDocumentCalls, 0);
}

// ------------------------------------------------------- explicit documentId

console.log('getReportDocument_() -- explicit context.documentId resolves via DocumentApp.openById()');

{
  const { sandbox } = loadCode();
  let openByIdCalls = [];
  const fakeDoc = { getBody: () => ({ marker: 'opened-body' }), getId: () => 'DOC_123' };
  sandbox.DocumentApp.openById = (id) => { openByIdCalls.push(id); return fakeDoc; };
  sandbox.DocumentApp.getActiveDocument = () => { throw new Error('should not be called'); };

  const context = { documentId: 'DOC_123' };

  check('getReportDocument_(context) returns the document DocumentApp.openById() resolved',
    sandbox.getReportDocument_(context), fakeDoc);
  check('openById() was called with the exact documentId',
    openByIdCalls, ['DOC_123']);
  check('getReportDocumentId_(context) returns context.documentId directly (no re-resolve needed)',
    sandbox.getReportDocumentId_(context), 'DOC_123');
}

console.log('getReportDocument_() -- repeated calls with the same context do not reopen the document');

{
  const { sandbox } = loadCode();
  let openByIdCalls = 0;
  const fakeDoc = { getBody: () => ({}), getId: () => 'DOC_ONCE' };
  sandbox.DocumentApp.openById = () => { openByIdCalls++; return fakeDoc; };

  const context = { documentId: 'DOC_ONCE' };

  sandbox.getReportDocument_(context);
  sandbox.getReportBody_(context);
  sandbox.getReportDocument_(context);
  sandbox.getReportDocumentId_(context);
  sandbox.getReportBody_(context);

  check('DocumentApp.openById() was called exactly once across 5 calls sharing one context',
    openByIdCalls, 1);
}

console.log('getReportDocument_() -- two separate contexts for the same documentId each open independently');

{
  const { sandbox } = loadCode();
  let openByIdCalls = 0;
  sandbox.DocumentApp.openById = () => { openByIdCalls++; return { getBody: () => ({}), getId: () => 'DOC_SEP' }; };

  sandbox.getReportDocument_({ documentId: 'DOC_SEP' });
  sandbox.getReportDocument_({ documentId: 'DOC_SEP' });

  check('a fresh context object is not cached across contexts (caching is per-context, not global)',
    openByIdCalls, 2);
}

// ------------------------------------------------------- documentId resolution

console.log('getReportDocumentId_() -- resolution priority');

{
  const { sandbox } = loadCode();
  sandbox.DocumentApp.getActiveDocument = () => ({ getBody: () => ({}), getId: () => 'FROM_ACTIVE' });
  check('no context -> id comes from the active document',
    sandbox.getReportDocumentId_(), 'FROM_ACTIVE');
}

{
  const { sandbox } = loadCode();
  sandbox.DocumentApp.openById = () => { throw new Error('should not be called: documentId already known'); };
  check('context.documentId is returned directly, without resolving a Document at all',
    sandbox.getReportDocumentId_({ documentId: 'RAW_ID' }), 'RAW_ID');
}

{
  const { sandbox } = loadCode();
  const explicitDoc = { getBody: () => ({}), getId: () => 'FROM_DOCUMENT_OBJECT' };
  check('context.document (no documentId) -> id comes from document.getId()',
    sandbox.getReportDocumentId_({ document: explicitDoc }), 'FROM_DOCUMENT_OBJECT');
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
