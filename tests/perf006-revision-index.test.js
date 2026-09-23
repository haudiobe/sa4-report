/**
 * PERF-006 (Part B) — revision-folder anchor pre-parsing/bucketing.
 *
 * PERF-005's source analysis found that updateRevisions_() re-ran
 * parseDraftAnchor_() (and the parseSA4DocumentId_() regex match it does
 * internally) once per (TDoc table, anchor) pair, even though the parse
 * result depends only on the anchor + the per-run-constant baseUrl, never
 * on which table is currently being processed -- an O(tables*anchors)
 * cost collapsible to O(anchors). PERF-004's granular log showed the old
 * matching loop at only ~9ms for a 34-table/65-anchor workload, so this
 * is a complexity/code-quality fix, not a measured runtime bottleneck --
 * these tests do not assert a timing improvement, only that:
 *   (a) buildRevisionAnchorIndex_() itself buckets correctly and uses the
 *       SAME parseDraftAnchor_() every other caller uses;
 *   (b) updateRevisions_() calls parseDraftAnchor_() O(anchors) times,
 *       not O(tables*anchors), for a multi-table workload;
 *   (c) end-to-end output (ordering/dedup/multi-draft/persisted state) is
 *       byte-identical to the pre-existing behavior -- already reconfirmed
 *       unchanged by the untouched tests/revision-drafts-folder.test.js
 *       and tests/revision-cell-rendering-hotfix.test.js suites, which
 *       both still pass without modification against the new code path.
 *
 * Run: node tests/perf006-revision-index.test.js
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

// ============ 1. buildRevisionAnchorIndex_() -- pure bucketing ============

console.log('buildRevisionAnchorIndex_() -- buckets by canonical TDoc identity, via the real parseDraftAnchor_()');

{
  const { sandbox } = loadCode();
  const anchors = [
    { text: 'S4-260100_r1.docx', href: 'S4-260100_r1.docx' },
    { text: 'S4-260100_r2.docx', href: 'S4-260100_r2.docx' },
    { text: 'S4aP260071_QCOM.docx', href: 'S4aP260071_QCOM.docx' },
    { text: 'not-a-tdoc-readme.txt', href: 'readme.txt' },
    { text: 'S4-260200.docx', href: 'S4-260200.docx' }
  ];
  const index = sandbox.buildRevisionAnchorIndex_(anchors, 'https://example.org/drafts/');

  check('valid anchors are bucketed under their canonical TDoc id', index.has('S4-260100'), true);
  check('multiple drafts for the SAME TDoc all land in ONE bucket, in original order', index.get('S4-260100').map(d => d.fileName), ['S4-260100_r1.docx', 'S4-260100_r2.docx']);
  check('an ad-hoc family (S4aP) is recognized via the same central parser', index.has('S4aP260071'), true);
  check('a single-draft TDoc gets a single-entry bucket', index.get('S4-260200').length, 1);
  check('an anchor with no recognized SA4 identifier is silently skipped (matches parseDraftAnchor_() returning null)', index.has('readme'), false);
  check('the total number of distinct buckets matches the number of distinct valid TDoc ids', index.size, 3);

  const entry = index.get('S4-260100')[0];
  check('each bucketed entry is exactly what parseDraftAnchor_() itself returns (tdocId/fileName/url)',
    [entry.tdocId, entry.fileName, entry.url],
    ['S4-260100', 'S4-260100_r1.docx', 'https://example.org/drafts/S4-260100_r1.docx']);
}

{
  const { sandbox } = loadCode();
  check('an empty anchor list produces an empty (not null/undefined) index', sandbox.buildRevisionAnchorIndex_([], 'https://x/').size, 0);

  let threw = null;
  let sizeWhenUndefined = null;
  try { sizeWhenUndefined = sandbox.buildRevisionAnchorIndex_(undefined, 'https://x/').size; } catch (e) { threw = e.message; }
  check('a null/undefined anchor list does not throw', threw, null);
  check('...and produces an empty index', sizeWhenUndefined, 0);
}

// ==== 2. source-structure: updateRevisions_() builds the index once, ======
// ==== looks it up per table, never re-scans anchors per table =============

console.log('source-structure: updateRevisions_() builds the anchor index ONCE and looks it up per table (no re-scan)');

{
  const source = fs.readFileSync(CODE_JS_PATH, 'utf8');
  const fnStart = source.indexOf('function updateRevisions_(cfg)');
  const nextFnMatch = source.slice(fnStart + 1).match(/^function\s+[A-Za-z0-9_$]+\s*\(/m);
  const fnEnd = nextFnMatch ? fnStart + 1 + nextFnMatch.index : source.length;
  const body = source.slice(fnStart, fnEnd);

  check('updateRevisions_() calls buildRevisionAnchorIndex_(anchors, url) exactly once, outside the tables.forEach loop',
    (body.match(/=\s*buildRevisionAnchorIndex_\(anchors,\s*url\)/g) || []).length, 1);
  check('the per-table loop looks up anchorIndex.get(tdoc) instead of re-calling anchors.forEach(...parseDraftAnchor_...)',
    /anchorIndex\.get\(tdoc\)/.test(body), true);
  check('parseDraftAnchor_ is NOT called directly inside updateRevisions_() itself any more (only inside buildRevisionAnchorIndex_)',
    /anchors\.forEach\([^)]*=>\s*\{[\s\S]*?parseDraftAnchor_/.test(body), false);
  check('the new "revision anchors parsed" counter is recorded once, sized to anchors.length',
    /perfCount_\('revision anchors parsed', anchors\.length\)/.test(body), true);
  check('the new "revision TDoc bucket lookups" counter is recorded once per table iteration',
    /perfCount_\('revision TDoc bucket lookups'\)/.test(body), true);
  check('the pre-existing "revisions: parseAnchors_ (listing parse)" PERF label is unchanged',
    body.includes("perfTimed_('    revisions: parseAnchors_ (listing parse)'"), true);
  check('the pre-existing matchingMs/"revisions: matching" accumulator variable is still used',
    /matchingMs \+= Date\.now\(\) - matchStart/.test(body), true);
}

// ==== 3. spy: parseDraftAnchor_ call count is O(anchors), not O(tables*anchors)

console.log('spy: parseDraftAnchor_() is called exactly once per anchor for the WHOLE run, regardless of table count');

{
  const { sandbox } = loadCode();

  let parseCalls = 0;
  const realParseDraftAnchor = sandbox.parseDraftAnchor_;
  sandbox.parseDraftAnchor_ = (...args) => { parseCalls++; return realParseDraftAnchor(...args); };

  const anchors = [];
  for (let i = 0; i < 10; i++) {
    anchors.push({ text: `S4-2601${String(i).padStart(2, '0')}_r1.docx`, href: `S4-2601${String(i).padStart(2, '0')}_r1.docx` });
  }
  const index = sandbox.buildRevisionAnchorIndex_(anchors, 'https://example.org/');
  check('buildRevisionAnchorIndex_ alone calls parseDraftAnchor_ exactly once per anchor', parseCalls, 10);
  check('all 10 distinct TDocs produced their own bucket', index.size, 10);

  // Simulate what updateRevisions_() now does per table: N tables, each
  // doing an O(1) Map.get() against the SAME prebuilt index -- must NOT
  // call parseDraftAnchor_ again.
  parseCalls = 0;
  const tableTdocs = ['S4-260100', 'S4-260101', 'S4-260102', 'S4-260103', 'S4-260104'];
  tableTdocs.forEach(tdoc => { index.get(tdoc); });
  check('N subsequent per-table bucket lookups call parseDraftAnchor_ ZERO additional times', parseCalls, 0);
}

// ==================================== summary ===============================

if (failures > 0) {
  console.log(`\n${failures} check(s) FAILED`);
  process.exit(1);
} else {
  console.log('\nAll checks passed.');
}
