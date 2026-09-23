/**
 * SA4-PROD-008 — Standard X.2 IPR / antitrust / consensus principles
 * section, generated directly (not copied from any template document) for
 * EVERY report type, via the new appendStandardIprSection_(body,
 * agendaPrefixNum) helper.
 *
 * appendStandardIprSection_() is a small, self-contained DocumentApp
 * consumer (appendParagraph/appendListItem/setHeading/setGlyphType only --
 * no DocumentApp.openById, no network, no UrlFetchApp) -- unlike the full
 * buildSkeletonWithTdocTables() orchestration (explicitly out of scope for
 * execution per SA4-ARCH-002/tests/report-structure.test.js's precedent),
 * this function CAN be executed directly against a small hand-built fake
 * Body, and this suite does so to prove the real generated content and
 * structure, not just its source shape.
 *
 * Run: node tests/standard-ipr-section.test.js
 */

const fs = require('fs');
const { loadCode, CODE_JS_PATH } = require('./helpers/load-code.js');

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

function extractFunctionBody(source, fnName) {
  const startMatch = source.match(new RegExp('^function ' + fnName + '\\(', 'm'));
  if (!startMatch) return null;
  const startIndex = startMatch.index;
  const nextFnRe = /^function\s+[A-Za-z0-9_$]+\s*\(/gm;
  nextFnRe.lastIndex = startIndex + startMatch[0].length;
  const next = nextFnRe.exec(source);
  const endIndex = next ? next.index : source.length;
  return source.slice(startIndex, endIndex);
}

function stripComments(src) {
  return src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '');
}

// load-code.js's shared sandbox does not define DocumentApp.GlyphType (no
// other existing production function under test needed it before this
// task). appendStandardIprSection_() is the first to call
// .setGlyphType(DocumentApp.GlyphType.BULLET) -- patched on locally here
// rather than touching the shared test helper, since no other suite needs
// it.
function loadCodeWithGlyphType(options) {
  const result = loadCode(options);
  result.sandbox.DocumentApp.GlyphType = { BULLET: 'BULLET' };
  return result;
}

// Minimal fake Body -- just appendParagraph()/appendListItem(), each
// returning an element with setHeading()/setGlyphType() (both no-ops that
// simply record what was set), plus getText()/getType() so the test can
// inspect what was produced.
function makeFakeBody() {
  const children = [];
  function makeParagraph(text) {
    const el = {
      _kind: 'PARAGRAPH',
      _heading: 'NORMAL',
      getType: () => 'PARAGRAPH',
      getText: () => text,
      getHeading: () => el._heading,
      setHeading: (h) => { el._heading = h; return el; }
    };
    children.push(el);
    return el;
  }
  function makeListItem(text) {
    const el = {
      _kind: 'LIST_ITEM',
      _glyph: null,
      getType: () => 'LIST_ITEM',
      getText: () => text,
      getGlyphType: () => el._glyph,
      setGlyphType: (g) => { el._glyph = g; return el; }
    };
    children.push(el);
    return el;
  }
  return {
    children,
    appendParagraph: (text) => makeParagraph(text),
    appendListItem: (text) => makeListItem(text)
  };
}

// ================================================== 1. source-structure ===

console.log('source-structure: appendStandardIprSection_() exists, is prefix-independent, and iprSection now calls it unconditionally');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');

  const iprFnBody = extractFunctionBody(source, 'appendStandardIprSection_');
  if (!iprFnBody) {
    failures++;
    console.log('  FAIL could not locate "function appendStandardIprSection_(" in Code.js');
  } else {
    check('appendStandardIprSection_() never hardcodes "5.2" (it is parameterized by agendaPrefixNum)',
      /['"]5\.2/.test(stripComments(iprFnBody)), false);
    check('appendStandardIprSection_() never creates a bare "X.2 " heading itself (only X.2.1-X.2.4 -- the caller\'s own agenda heading remains the anchor)',
      /`\$\{agendaPrefixNum\}\.2\s/.test(iprFnBody), false);
    check('appendStandardIprSection_() creates X.2.1 through X.2.4', [
      // ADDON-007A: children are `${base}.${off + N}`; base defaults to
      // `${agendaPrefixNum}.2` and off to 0 -- identical output for main meetings.
      /\$\{base\}\.\$\{off \+ 1\} Introduction/.test(iprFnBody),
      /\$\{base\}\.\$\{off \+ 2\} Call for IPRs/.test(iprFnBody),
      /\$\{base\}\.\$\{off \+ 3\} Statement regarding competition law/.test(iprFnBody),
      /\$\{base\}\.\$\{off \+ 4\} Consensus principles reminder/.test(iprFnBody)
    ], [true, true, true, true]);
    check('appendStandardIprSection_() uses two appendListItem() calls (the IPR invitation bullets)',
      (iprFnBody.match(/appendListItem\(/g) || []).length, 2);
  }

  const skeletonBody = extractFunctionBody(source, 'buildSkeletonWithTdocTables');
  if (!skeletonBody) {
    failures++;
    console.log('  FAIL could not locate "function buildSkeletonWithTdocTables(" in Code.js');
  } else {
    check('buildSkeletonWithTdocTables() calls appendStandardIprSection_(body, agendaPrefixNum) in the iprSection branch',
      /appendStandardIprSection_\(body,\s*agendaPrefixNum,/.test(skeletonBody), true);
    check('the old isSWGReport branching for the IPR section is gone (no report-type-specific IPR handling remains)',
      /isSWGReport/.test(stripComments(skeletonBody)), false);
    // Registration-boundary extraction (SA4-PROD-007), reallocation
    // staging, and the AOB/close/registration branches are untouched --
    // still present exactly as before.
    check('the registration-boundary extraction (isBeforeRegistrationBoundary_) is still present, unchanged',
      /isBeforeRegistrationBoundary_\(/.test(skeletonBody), true);
    check('the X.1.4 Documents section (appendTdocDetailTable_) is still present, unchanged',
      /appendTdocDetailTable_\(body,\s*tdocData,\s*tdocData\.row\[tdocData\.agendaCol\]\)/.test(skeletonBody), true);
    check('the AOB branch (title-based, after iprSection) is still present, unchanged',
      /any other business/.test(skeletonBody), true);
    check('the regular "TDOC tables" else branch (for real 5.3+ items) is still present, unchanged',
      /Regular agenda item - insert TDOC tables/.test(skeletonBody), true);
  }

  // Reviewer API / reallocation / mailing-list source resolution are
  // untouched by this task -- not re-asserted here in detail (covered by
  // their own dedicated suites, all still green), just confirmed absent
  // from this diff's own functions.
  const reallocBody = extractFunctionBody(source, 'ensureReallocationTable_');
  check('ensureReallocationTable_() does not reference appendStandardIprSection_ (unrelated, untouched)',
    reallocBody ? /appendStandardIprSection_/.test(reallocBody) : null, false);
}

// ======================================== 2. real execution -- meeting 86178

console.log('appendStandardIprSection_() -- real execution, prefix "5" (meeting 86178)');

{
  const { sandbox } = loadCodeWithGlyphType();
  const body = makeFakeBody();
  sandbox.appendStandardIprSection_(body, '5');

  const headings = body.children.filter(c => c._kind === 'PARAGRAPH' && c._heading !== 'NORMAL').map(c => c.getText());
  check('exactly four subsection headings are created: 5.2.1-5.2.4',
    headings, ['5.2.1 Introduction', '5.2.2 Call for IPRs', '5.2.3 Statement regarding competition law', '5.2.4 Consensus principles reminder']);

  check('no bare "5.2 " heading (without a third segment) was created -- the agenda\'s own heading remains the sole X.2 anchor',
    body.children.some(c => c._kind === 'PARAGRAPH' && /^5\.2\s/.test(c.getText()) && !/^5\.2\.\d/.test(c.getText())),
    false);

  const allTexts = body.children.map(c => c.getText());
  check('the Introduction body text is present verbatim',
    allTexts.indexOf('The chair read the antitrust and IPR clause at the opening of the meeting session.') !== -1, true);
  check('the Call for IPRs opening paragraph is present verbatim (curly quotes, apostrophe preserved)',
    allTexts.indexOf('“I draw your attention to your obligations under the 3GPP Partner Organizations’ IPR policies. Every Individual Member organization is obliged to declare to the Partner Organization or Organizations of which it is a member any IPR owned by the Individual Member or any other organization which is or is likely to become essential to the work of 3GPP.') !== -1,
    true);
  check('the "Delegates are asked to take note" lead-in is present',
    allTexts.indexOf('Delegates are asked to take note that they are thereby invited:') !== -1, true);

  const listItems = body.children.filter(c => c._kind === 'LIST_ITEM').map(c => c.getText());
  check('exactly two IPR-invitation bullet items are present, verbatim, in order',
    listItems, [
      'to investigate whether their organization or any other organization owns IPRs which were, or were likely to become Essential in respect of the work of 3GPP.',
      'to notify their respective Organizational Partners of all potential IPRs, e.g., for ETSI, by means of the IPR Information Statement and the Licensing declaration forms"'
    ]);
  check('both bullet items were given a glyph type (rendered as an actual bulleted list, not plain paragraphs)',
    body.children.filter(c => c._kind === 'LIST_ITEM').every(c => c._glyph !== null), true);

  check('the competition-law statement (X.2.3) is present verbatim, closing with a curly right quote',
    allTexts.indexOf('Furthermore, I would like to remind you that timely submission of work items in advance of TSG/WG/SWG meetings is important to allow for full and fair consideration of such matters.”') !== -1,
    true);
  check('the consensus-principles statement (X.2.4) is present verbatim, closing with a curly right quote',
    allTexts.some(t => t.indexOf('3GPP endeavours to reach consensus on all decisions') !== -1 && t.trim().endsWith('”')),
    true);
}

// =========================== 3. prefix independence -- main-meeting prefixes

console.log('appendStandardIprSection_() -- prefix independence: main 6G ("11") and main Audio ("7")');

{
  const { sandbox } = loadCodeWithGlyphType();

  const body11 = makeFakeBody();
  sandbox.appendStandardIprSection_(body11, '11');
  const headings11 = body11.children.filter(c => c._kind === 'PARAGRAPH' && c._heading !== 'NORMAL').map(c => c.getText());
  check('main 6G (prefix "11") produces 11.2.1-11.2.4',
    headings11, ['11.2.1 Introduction', '11.2.2 Call for IPRs', '11.2.3 Statement regarding competition law', '11.2.4 Consensus principles reminder']);

  const body7 = makeFakeBody();
  sandbox.appendStandardIprSection_(body7, '7');
  const headings7 = body7.children.filter(c => c._kind === 'PARAGRAPH' && c._heading !== 'NORMAL').map(c => c.getText());
  check('main Audio (prefix "7") produces 7.2.1-7.2.4',
    headings7, ['7.2.1 Introduction', '7.2.2 Call for IPRs', '7.2.3 Statement regarding competition law', '7.2.4 Consensus principles reminder']);

  // The quoted boilerplate body text itself must be prefix-independent --
  // identical wording regardless of which meeting's prefix is used.
  const bodyTexts5 = (() => {
    const b = makeFakeBody();
    sandbox.appendStandardIprSection_(b, '5');
    return b.children.map(c => c.getText());
  })();
  const bodyTexts11 = body11.children.map(c => c.getText());
  const nonHeadingTexts5 = bodyTexts5.filter(t => !/^5\.2\.\d/.test(t));
  const nonHeadingTexts11 = bodyTexts11.filter(t => !/^11\.2\.\d/.test(t));
  check('the quoted/plain boilerplate body text is byte-identical across different prefixes',
    nonHeadingTexts5, nonHeadingTexts11);
}

// ========================================================== summary =======

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
