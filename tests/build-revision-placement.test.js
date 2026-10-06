/**
 * Revision placement after Build Report from Scratch (after T-2026.10.8).
 *
 * The rule of the report: the table of a revision stands directly below the
 * table of the document it revises, and a chain A -> B -> C stands in that
 * order. An update has always ended with the pass that establishes this
 * (rearrangeRevisionTables_()). A build ordered revisions only INSIDE one
 * agenda item (orderTdocsByRevision_()), so a revision registered under
 * another agenda item than its parent stood apart until the first update.
 *
 * The build now ends its table writing with the same pass, on the TDoc list
 * it has already downloaded. These tests run the real build and the real
 * update on the fake document:
 *
 *   1. A under one agenda item, B (its revision) under another;
 *   2. a chain A -> B -> C over three agenda items;
 *   3. what was right before is as it was (same agenda item; no revisions);
 *   4. a build leaves what an update leaves: the update after it moves nothing;
 *   5. reallocations: a rebuild with one, and applying one followed by an update;
 *   6. no second download of the TDoc list; a failure does not fail the build;
 *   7. the Revisions cell is not part of this.
 *
 * Run: node tests/build-revision-placement.test.js
 */

const fs = require('fs');
const { CODE_JS_PATH } = require('./helpers/load-code.js');
const { templateReport } = require('./helpers/template-report.js');

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

const CODE = fs.readFileSync(CODE_JS_PATH, 'utf8').replace(/\r/g, '');
const functionSource = (name) => { const start = CODE.indexOf('\nfunction ' + name + '('); return start === -1 ? null : CODE.slice(start + 1, CODE.indexOf('\n}\n', start) + 2); };
const A = 'S4aA269001', B = 'S4aA269002', C = 'S4aA269003', X = 'S4aA269004', Y = 'S4aA269005', Z = 'S4aA269006';
const t = (id, agenda, revisedTo, status) => ({ id: id, agenda: agenda, status: status || 'available', revisedTo: revisedTo || '' });
/** The TDoc numbers of the tables below the heading of `agenda`, in order. */
const under = (r, agenda) => {
  const outline = r.outline();
  const at = outline.findIndex((line) => line.indexOf('# ' + agenda + ' ') === 0);
  const out = [];
  for (let i = at + 1; i < outline.length && outline[i][0] !== '#'; i++) out.push(outline[i]);
  return out;
};
const moved = (r) => r.logs.filter((l) => /^Revisions: \d+ moved/.test(l));
const TDOC_LIST = 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=85916';

console.log('1. A under 4.3, its revision B under 4.4');
{
  const r = templateReport({ tdocs: [t(X, '4.3'), t(A, '4.3', B, 'revised'), t(Y, '4.3'), t(Z, '4.4'), t(B, '4.4')] });
  const built = r.build();
  check('the build completes', [built.ok, built.error], [true, null]);
  check('B is directly below A', r.order().indexOf(B) - r.order().indexOf(A), 1);
  check('under the agenda item of A, in its place; the other tables keep their order', [under(r, '4.3'), under(r, '4.4')], [[X, A, B, Y], [Z]]);
  check('every table is there once', r.order().slice().sort(), [A, B, X, Y, Z].sort());
  check('B keeps the agenda item the TDoc list gives it (only its place changes)', r.field(B, 'Agenda Item'), '4.4');
  check('A says what it was revised to, once', r.field(A, 'Disposition'), 'Revised to ' + B);
  check('the build says what it moved', moved(r), ['Revisions: 1 moved, 0 disposition(s) filled']);
  check('the revision relationships are stored as the update stores them', JSON.parse(r.docProps.getProperty('REVISION_MAP')), { [A]: B });
}

console.log('1. the same with the revision registered BEFORE its parent in the list, and under an earlier agenda item');
{
  const r = templateReport({ tdocs: [t(B, '4.3'), t(X, '4.3'), t(A, '4.5', B, 'revised'), t(Y, '4.5')] });
  r.build();
  check('B is directly below A, under the agenda item of A', [under(r, '4.3'), under(r, '4.5')], [[X], [A, B, Y]]);
}

console.log('2. a chain A -> B -> C over three agenda items');
{
  const r = templateReport({ tdocs: [t(C, '4.5'), t(X, '4.5'), t(B, '4.4', C, 'revised'), t(Y, '4.4'), t(A, '4.3', B, 'revised'), t(Z, '4.3')] });
  r.build();
  check('A, B, C stand one below the other, in that order', [under(r, '4.3'), under(r, '4.4'), under(r, '4.5')], [[A, B, C, Z], [Y], [X]]);
  check('each says what it was revised to', [r.field(A, 'Disposition'), r.field(B, 'Disposition'), r.field(C, 'Disposition')], ['Revised to ' + B, 'Revised to ' + C, '']);
  check('two tables were moved', moved(r), ['Revisions: 2 moved, 0 disposition(s) filled']);
}

console.log('2. two chains that cross agenda items do not mix');
{
  const r = templateReport({ tdocs: [t(A, '4.3', B, 'revised'), t(X, '4.3', Y, 'revised'), t(B, '4.4'), t(Y, '4.4', Z, 'revised'), t(Z, '4.5')] });
  r.build();
  check('each chain is contiguous', [under(r, '4.3'), under(r, '4.4'), under(r, '4.5')], [[A, B, X, Y, Z], [], []]);
}

console.log('2. a parent among the documents of the opening item (the "Documents" sub-section), its revision under a normal agenda item');
{
  // The build takes such TDocs out of its groups to write them in that sub-section; the placement still knows them.
  const r = templateReport({ tdocs: [t(A, '1', B, 'revised'), t(X, '1'), t(Y, '4.3'), t(B, '4.3')] });
  r.build();
  check('the revision is directly below its parent there, as an update places it', [under(r, '1.4'), under(r, '4.3'), moved(r)], [[A, B, X], [Y], ['Revisions: 1 moved, 0 disposition(s) filled']]);
}

console.log('3. what was right before is as it was');
{
  const same = templateReport({ tdocs: [t(X, '4.3'), t(B, '4.3'), t(A, '4.3', B, 'revised'), t(Y, '4.4')] });
  same.build();
  check('a revision under the same agenda item is below its parent, as before, and nothing had to be moved', [under(same, '4.3'), under(same, '4.4'), moved(same)], [[X, A, B, Y].filter((id) => id !== Y), [Y], ['Revisions: 0 moved, 0 disposition(s) filled']]);
  const none = templateReport({ tdocs: [t(X, '4.3'), t(A, '4.3'), t(Y, '4.4'), t(B, '4.4')] });
  none.build();
  check('a report without revisions is in the order of the TDoc list; nothing is moved', [under(none, '4.3'), under(none, '4.4'), moved(none)], [[X, A], [Y, B], ['Revisions: 0 moved, 0 disposition(s) filled']]);
  const outside = templateReport({ tdocs: [t(X, '4.3'), t(A, '4.3', 'S4aA269999', 'revised')] });
  outside.build();
  check('a revision that is not in this report changes nothing', [under(outside, '4.3'), outside.field(A, 'Disposition')], [[X, A], 'Revised to S4aA269999']);
}

console.log('4. a build leaves what an update leaves');
[
  ['A and B over two agenda items', [t(X, '4.3'), t(A, '4.3', B, 'revised'), t(Y, '4.3'), t(Z, '4.4'), t(B, '4.4')]],
  ['a chain over three', [t(C, '4.5'), t(X, '4.5'), t(B, '4.4', C, 'revised'), t(Y, '4.4'), t(A, '4.3', B, 'revised'), t(Z, '4.3')]],
  ['two chains', [t(A, '4.3', B, 'revised'), t(X, '4.3', Y, 'revised'), t(B, '4.4'), t(Y, '4.4', Z, 'revised'), t(Z, '4.5')]],
  ['a parent in the Documents sub-section', [t(A, '1', B, 'revised'), t(X, '1'), t(Y, '4.3'), t(B, '4.3')]]
].forEach(([name, tdocs]) => {
  const r = templateReport({ tdocs: tdocs });
  r.build();
  const afterBuild = r.outline();
  r.mark();
  const updated = r.update();
  check(name + ': the update after the build moves nothing and leaves the document as it is', [updated.success, moved(r), r.outline()], [true, ['Revisions: 0 moved, 0 disposition(s) filled'], afterBuild]);
});

console.log('5. reallocations');
{
  // A rebuild with a saved reallocation: A goes to 4.5, its revision B is registered under 4.4.
  const r = templateReport({ tdocs: [t(X, '4.3'), t(A, '4.3', B, 'revised'), t(Z, '4.4'), t(B, '4.4'), t(Y, '4.5')] });
  r.build();
  r.user(() => r.s.saveReallocation(A, '4.3', '4.5', 'synthetic'));
  const rebuilt = r.build();
  check('a rebuild with the reallocation completes and keeps it', [rebuilt.ok, Object.keys(r.s.getReallocationMap_())], [true, [A]]);
  check('A is under its new agenda item, and B directly below it', [under(r, '4.3'), under(r, '4.4'), under(r, '4.5')], [[X], [Z], [A, B, Y]]);
  check('A shows its new agenda item', r.field(A, 'Agenda Item'), '4.5');
  const afterBuild = r.outline();
  r.mark();
  r.update();
  check('the update after it changes nothing', [moved(r), r.outline()], [['Revisions: 0 moved, 0 disposition(s) filled'], afterBuild]);
}
{
  // Applying a reallocation to a built report moves the one table; the update then brings its revision, as it always did.
  const r = templateReport({ tdocs: [t(X, '4.3'), t(A, '4.3', B, 'revised'), t(Z, '4.4'), t(B, '4.4'), t(Y, '4.5')] });
  r.build();
  check('built: B below A under 4.3', under(r, '4.3'), [X, A, B]);
  r.user(() => { r.s.saveReallocation(A, '4.3', '4.5', 'synthetic'); r.s.applyDocumentReallocationsToBody_(r.body, r.s.getReallocationMap_()); });
  check('applying the reallocation moves A alone (unchanged behaviour)', [under(r, '4.3'), under(r, '4.5')], [[X, B], [Y, A]]);
  r.mark();
  r.update();
  check('the update puts B below A again (unchanged behaviour)', [under(r, '4.3'), under(r, '4.5'), moved(r)], [[X], [Y, A, B], ['Revisions: 1 moved, 0 disposition(s) filled']]);
}

console.log('6. cost and failure');
{
  const r = templateReport({ tdocs: [t(X, '4.3'), t(A, '4.3', B, 'revised'), t(Z, '4.4'), t(B, '4.4')] });
  r.build();
  check('the build downloads the TDoc list once: the placement uses the list it has', r.fetches.filter((u) => u === TDOC_LIST).length, 1);
  const build = functionSource('buildSkeletonWithTdocTables');
  check('the build calls the pass of the update, once, with the TDocs it downloaded (source)',
    [(build.match(/rearrangeRevisionTables_\(/g) || []).length, /rearrangeRevisionTables_\(cfg, \{ all: \{ tdocs: allTdocs \} \}\);/.test(build), (build.match(/downloadAndGroupTdocs_\(/g) || []).length], [1, true, 1]);
  check('it is the pass itself: no second ordering was written for the build', [(CODE.match(/\nfunction rearrangeRevisionTables_\(/g) || []).length, (CODE.match(/\nfunction moveTableAfter_\(/g) || []).length, /moveTableAfter_\(|orderRevisionChains_\(/.test(build)], [1, 1, false]);
  check('it comes after the last TDoc table is written and before the reallocations, the attendance and the formatting',
    [build.indexOf('rearrangeRevisionTables_(') > build.lastIndexOf('orderTdocsByRevision_('), build.indexOf('rearrangeRevisionTables_(') < build.indexOf('saveReallocation('),
      build.indexOf('rearrangeRevisionTables_(') < build.indexOf('finishAdhocAttendanceRebuild_('), build.indexOf('rearrangeRevisionTables_(') < build.indexOf('removeRowHeightAndSpacing()')], [true, true, true, true]);
  check('the update is as it was: it still ends its TDoc work with the pass', /rearrangeRevisionTables_\(cfg, tdocGroups, tdocTableIndex, context\)/.test(functionSource('continuousUpdateCore_')), true);
}
{
  const r = templateReport({ tdocs: [t(X, '4.3'), t(A, '4.3', B, 'revised'), t(Z, '4.4'), t(B, '4.4')] });
  r.s.rearrangeRevisionTables_ = () => { throw new Error('synthetic placement failure'); };
  const built = r.s.buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true });
  check('a placement that fails does not fail the build: the report is built, and the result says what was not done',
    [built.ok, /Revisions could not be placed below the documents they revise \(synthetic placement failure\)\. The next update of the report places them\./.test(built.note), r.order().slice().sort()], [true, true, [A, B, X, Z].sort()]);
}

console.log('7. the Revisions cell is not part of this');
{
  check('renderRevisionsCellContent_() still writes one revision per line (source; its own tests: revision-cell-rendering-hotfix)',
    /const fullText = entries\.map\(e => e\.line\)\.join\('\\n'\);\n  cell\.setText\(fullText\);/.test(functionSource('renderRevisionsCellContent_')), true);
  const r = templateReport({ tdocs: [t(A, '4.3', B, 'revised'), t(B, '4.4')] });
  r.build();
  check('the build does not write the Revisions cells', [r.field(A, 'Revisions'), r.field(B, 'Revisions')], ['', '']);
}

console.log(failures ? `\n${failures} check(s) FAILED` : '\nAll build revision placement checks passed.');
process.exitCode = failures ? 1 : 0;
