/**
 * SA4-ARCH-007 — Agenda structure projection characterization.
 *
 * This is NOT a regression-baseline suite in the usual sense: it exists to
 * pin down, with actual execution (not just reading the source), several
 * non-obvious current behaviors of the agenda-structure parsers
 * (parseAgendaFromHeadings_, parseAgendaFromTables_,
 * parseAgendaStructureWithText_) that the SA4-ARCH-007 report relies on as
 * evidence. Nothing in Code.js was changed to make these pass -- they
 * capture behavior exactly as found.
 *
 * Uses small, purpose-built fakes for a DocumentApp Body (paragraphs-only,
 * or a single agenda table) -- NOT a general DocumentApp mock. Sized to
 * exactly what these two parser functions call (getNumChildren/getChild/
 * asParagraph/getText/getHeading for the heading parser; getTables/
 * getNumRows/getCell/getRow for the table parser).
 *
 * Per SA4-ARCH-007 instructions: characterization only, NOT committed.
 *
 * Run: node tests/agenda-structure-characterization.test.js
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

function makeFakeParagraphBody(sandbox, paragraphs) {
  return {
    getNumChildren: () => paragraphs.length,
    getChild: (i) => ({
      getType: () => sandbox.DocumentApp.ElementType.PARAGRAPH,
      asParagraph: () => ({
        getText: () => paragraphs[i].text,
        getHeading: () => paragraphs[i].heading || sandbox.DocumentApp.ParagraphHeading.NORMAL
      })
    }),
    getTables: () => []
  };
}

function makeFakeTableBody(rows) {
  const table = {
    getNumRows: () => rows.length,
    getCell: (r, c) => ({ getText: () => rows[r][c] }),
    getRow: (r) => ({
      getNumCells: () => rows[r].length,
      getCell: (c) => ({ getText: () => rows[r][c] })
    })
  };
  return { getTables: () => [table], getNumChildren: () => 0, getChild: () => null };
}

const MAIN_LIKE_PARAGRAPHS = (sandbox) => [
  { text: '7 Audio SWG', heading: sandbox.DocumentApp.ParagraphHeading.HEADING2 },
  { text: '7.1 Some Topic', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 },
  { text: '7.2 Another Topic', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 },
  { text: '8 MBS SWG', heading: sandbox.DocumentApp.ParagraphHeading.HEADING2 },
  { text: '8.1 MBS Topic', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 }
];

// ================================================================= finding 1
//
// parseAgendaFromHeadings_(body, prefix) returns [] for ANY non-empty
// prefix, regardless of document content -- a state-machine bug (inSection
// can never become true, because the same source line can never match both
// subItemRegex, which requires the prefix's dot + at least one digit, AND
// parentRegex, which requires exactly the bare number with nothing after
// it). This does not affect real production usage because the one caller
// that matters (buildSkeletonWithTdocTables -> parseAgendaForReport_, via
// the AGENDA_TDOC/ZIP path) always calls it with an EMPTY prefix
// (parseAgendaWordBlob_ -> parseAgendaStructureWithText_(doc.getBody(), ''))
// and filters separately afterward. It DOES affect the "template Google
// Doc" fallback branch of parseAgendaForReport_ (no AGENDA_TDOC configured)
// and parseAgendaDocumentById()'s Doc-ID branch, both of which pass the
// real prefix directly into this buggy path.

console.log('finding 1: parseAgendaFromHeadings_ with a non-empty prefix returns [] (state-machine bug)');

{
  const { sandbox } = loadCode();
  const body = makeFakeParagraphBody(sandbox, MAIN_LIKE_PARAGRAPHS(sandbox));

  check('parseAgendaFromHeadings_(body, "7.") returns [] despite matching content existing',
    sandbox.parseAgendaFromHeadings_(body, '7.'), []);
  check('parseAgendaStructureWithText_(body, "7.") ALSO returns [] (falls through to parseAgendaFromTables_, which finds no table either)',
    sandbox.parseAgendaStructureWithText_(body, '7.'), []);
}

console.log('finding 1b: the SAME body, same parser, with an EMPTY prefix (the real production ZIP-path call shape) works correctly and is unfiltered');

{
  const { sandbox } = loadCode();
  const body = makeFakeParagraphBody(sandbox, MAIN_LIKE_PARAGRAPHS(sandbox));
  const result = sandbox.parseAgendaFromHeadings_(body, '');

  check('returns all 5 items, unfiltered (filtering happens later, in parseAgendaForReport_ itself)',
    result.map(i => i.number), ['7', '7.1', '7.2', '8', '8.1']);
}

// ================================================================= finding 2
//
// parseAgendaFromTables_ (the OTHER half of parseAgendaStructureWithText_'s
// fallback chain) does NOT have the same bug -- its filtering is a simple
// per-row conditional, not a stateful "have we entered the section yet"
// flag, so it correctly includes the bare parent alongside descendants.

console.log('finding 2: parseAgendaFromTables_ (table-sourced agendas) does NOT have the heading-parser bug');

{
  const { sandbox } = loadCode();
  const rows = [['A.I.#', 'Agenda Item'], ['7', 'Audio SWG'], ['7.1', 'Some Topic'], ['7.2', 'Another Topic'], ['8', 'MBS SWG']];
  const result = sandbox.parseAgendaFromTables_(makeFakeTableBody(rows), '7.');

  check('correctly includes the bare parent "7" plus its descendants, excludes "8"',
    result.map(i => i.number), ['7', '7.1', '7.2']);
}

// ================================================================= finding 3
//
// Neither parser ever synthesizes a missing ancestor. If the bare parent
// isn't present as its own line/row in the source, it simply never appears
// in the output -- confirmed with both parsers.

console.log('finding 3: missing ancestors are never synthesized');

{
  const { sandbox } = loadCode();
  const paragraphsNoParent = [
    { text: '7.1 Some Topic', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 },
    { text: '7.2 Another Topic', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 }
  ];
  const headingResult = sandbox.parseAgendaFromHeadings_(makeFakeParagraphBody(sandbox, paragraphsNoParent), '');
  check('parseAgendaFromHeadings_: "7" is absent from source -> absent from output (not synthesized)',
    headingResult.map(i => i.number), ['7.1', '7.2']);

  const rowsNoParent = [['A.I.#', 'Agenda Item'], ['7.1', 'Some Topic'], ['7.2', 'Another Topic']];
  const tableResult = sandbox.parseAgendaFromTables_(makeFakeTableBody(rowsNoParent), '7.');
  check('parseAgendaFromTables_: same -- "7" absent from source -> absent from output',
    tableResult.map(i => i.number), ['7.1', '7.2']);
}

// ================================================================= finding 4
//
// Output order is exactly source-document order (top-to-bottom paragraph/
// row order), never re-sorted lexically or numerically.

console.log('finding 4: output order is source-document order, not re-sorted');

{
  const { sandbox } = loadCode();
  // Deliberately out-of-numeric-order in the source, to distinguish
  // "preserves source order" from "happens to look sorted".
  const outOfOrderParagraphs = [
    { text: '7.2 Second In Source', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 },
    { text: '7.1 First In Source Position But Higher Would-Be-Sorted-First Number', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 },
    { text: '7.10 Third In Source', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 }
  ];
  const result = sandbox.parseAgendaFromHeadings_(makeFakeParagraphBody(sandbox, outOfOrderParagraphs), '');
  check('output order mirrors source order exactly (7.2 before 7.1 before 7.10), not lexical/numeric sort',
    result.map(i => i.number), ['7.2', '7.1', '7.10']);
}

// ================================================================= finding 5
//
// Parsed agenda-item object shape: {number, title, level, heading, text}.
// No explicit parent/children/depth fields -- hierarchy exists ONLY as the
// dot-separated structure of the `number` string; `level` is a cached
// (dot-count + 1) derived value, not an independent hierarchy pointer.

console.log('finding 5: parsed agenda-item object shape');

{
  const { sandbox } = loadCode();
  const body = makeFakeParagraphBody(sandbox, [
    { text: '7.1.2 Deep Item', heading: sandbox.DocumentApp.ParagraphHeading.HEADING3 }
  ]);
  const result = sandbox.parseAgendaFromHeadings_(body, '');
  check('object has exactly the fields {number, title, level, heading, text}, no parent/children/depth',
    Object.keys(result[0]).sort(), ['heading', 'level', 'number', 'text', 'title']);
  check('level is derived from dot-count + 1 (7.1.2 has 2 dots -> level 3)', result[0].level, 3);
}

// ------------------------------------------------------------------- summary

console.log(failures === 0 ? '\nAll agenda-structure characterization findings confirmed.' : `\n${failures} check(s) did NOT match expectations.`);
process.exit(failures === 0 ? 0 : 1);
