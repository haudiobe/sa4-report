/**
 * Code.js as it was before the work after T-2026.10.12: the report-structure
 * fix and the configuration guard.
 *
 * Several test files hold a ledger of what changed in Code.js since a
 * release: which functions, and that nothing else did. The structure fix
 * changed six released functions and added one section; the configuration
 * guard changed one and added one section. Those ledgers are about their own
 * feature; they go on comparing with their release through this view, which
 * takes both out again -- every change named here, each found exactly once.
 * What each changed is pinned in its own test (tests/report-structure-fix.test.js,
 * tests/configuration-guard.test.js): without both, Code.js is, byte for
 * byte, the Code.js of T-2026.10.12.
 *
 * codeBeforeStructureFix(code) -> { code, problems }: both taken out.
 * codeWithoutConfigurationGuard(code) -> { code, problems }: the guard only.
 * `problems` names every change that was not found exactly once (then the
 * view is not to be trusted and the ledger that uses it fails).
 */

const BAR = '// =========================================================\n';
const SECTION_START = BAR + '// REPORT STRUCTURE (after T-2026.10.12) -- THE ORDER OF THE AGENDA SECTIONS; NO PARAGRAPHS LEFT BEHIND\n' + BAR;
const SECTION_END = BAR + '// END OF REPORT STRUCTURE\n' + BAR + '\n';

/** The functions of the section, in the order they stand in it. */
const SECTION_FUNCTIONS = ['isEmptyBodyParagraph_', 'isTableDirectlyBelow_', 'collapseEmptyParagraphsAfter_', 'removeSeparatorParagraph_', 'removeTableWithItsSeparator_', 'makeAgendaTitleResolver_', 'agendaNumberOfHeadingText_',
  'compareAgendaNumbers_', 'isClosingHeadingText_', 'findReportTerminalBoundary_', 'findOrderedAgendaInsertIndex_', 'findAgendaHeadingIndexBefore_', 'findMisplacedAgendaSections_',
  'moveMisplacedAgendaSection_', 'scanReportStructure_', 'repairReportStructure_'];

/** The released functions the fix changed. */
const CHANGED_FUNCTIONS = ['continuousUpdateCore_', 'insertNewTdoc_', 'findInsertionPointForAgendaItem_', 'findAgendaSectionEndIndex_', 'moveTableAfter_', 'downloadAndGroupTdocs_'];

/** [what the fix wrote, what the release had], each in Code.js exactly once. */
const CHANGES = [
  // continuousUpdateCore_(): the repair before the document is read ...
  ["    // The structure of the report, before anything is looked up in it: agenda\n" +
   "    // sections that an earlier release appended after the close of the report\n" +
   "    // go back into the order of the agenda, and the empty paragraphs the\n" +
   "    // revision placement left below TDoc tables are removed. A report that\n" +
   "    // needs neither is only read. Never fails the update.\n" +
   "    const agendaTitleOf = makeAgendaTitleResolver_(allTdocs, context);\n" +
   "    let structure = { sectionsMoved: 0, emptyParagraphsRemoved: 0 };\n" +
   "    try {\n" +
   "      structure = repairReportStructure_(body, agendaTitleOf);\n" +
   "    } catch (e) {\n" +
   "      Logger.log('Structure: the structure of the report could not be checked: ' + e.message);\n" +
   "    }\n\n", ''],
  // ... the title of the agenda item of a new TDoc ...
  ["        // The title of its agenda item, for a heading the report does not have yet.\n" +
   "        tdocData.agendaTitle = agendaTitleOf(tdocData.agendaItem);\n", ''],
  // ... and a moved section is a reason to format, like a new table.
  ["      newTdocsAdded + structure.sectionsMoved,\n", "      newTdocsAdded,\n"],
  // insertNewTdoc_()
  ["  if (insertIdx < 0) insertIdx = findInsertionPointForAgendaItem_(body, agendaItem, tdocData.agendaTitle || '');\n",
   "  if (insertIdx < 0) insertIdx = findInsertionPointForAgendaItem_(body, agendaItem, '');\n"],
  // findInsertionPointForAgendaItem_()
  ["  // Heading not found: create it IN ITS PLACE in the order of the agenda and\n" +
   "  // return the position after it. It used to be appended at the end of the\n" +
   "  // document -- after \"Close of the session\" and after the Attendance\n" +
   "  // section, whatever stood there.\n",
   "  // Heading not found: create it and return position after it\n"],
  ["  const at = findOrderedAgendaInsertIndex_(body, agendaItem);\n" +
   "  const newHeading = at >= body.getNumChildren() ? body.appendParagraph(headingText) : body.insertParagraph(at, headingText);\n",
   "  const newHeading = body.appendParagraph(headingText);\n"],
  ["  else newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING3);\n\n  return body.getChildIndex(newHeading) + 1;\n}\n",
   "  else newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING3);\n  \n  return body.getNumChildren();\n}\n"],
  // findAgendaSectionEndIndex_()
  ["      // The Attendance section closes the report: nothing of an agenda item\n" +
   "      // goes after it, also not of the item it stands in.\n" +
   "      if (nextText === ADHOC_ATTENDANCE_HEADING_) return j;\n", ''],
  // moveTableAfter_()
  ["  // \"In the right place\" is: the next table below `afterTable`, with nothing\n" +
   "  // but empty paragraphs between them. Google Docs keeps a paragraph between\n" +
   "  // two tables, so a table can never be the very next child of another one:\n" +
   "  // comparing child indices (as this did) found every revision out of place\n" +
   "  // on every update, moved it again, and left one empty paragraph behind\n" +
   "  // each time -- hundreds of them after some days of updates.\n" +
   "  if (isTableDirectlyBelow_(body, table, afterTable)) return null;\n",
   "  const targetIndex = body.getChildIndex(afterTable) + 1;\n" +
   "  const currentIndex = body.getChildIndex(table);\n" +
   "  if (currentIndex === targetIndex) return null; // already in the right place\n"],
  ["  body.insertTable(body.getChildIndex(afterTable) + 1, copy);\n" +
   "  removeTableWithItsSeparator_(body, table);\n" +
   "  // Whatever Google Docs added around the copy: one empty paragraph at most\n" +
   "  // stays between two tables.\n" +
   "  collapseEmptyParagraphsAfter_(body, afterTable);\n" +
   "  collapseEmptyParagraphsAfter_(body, copy);\n",
   "  body.insertTable(targetIndex, copy);\n" +
   "  table.removeFromParent();\n"],
  // downloadAndGroupTdocs_()
  ["    const agendaDescriptionCol = headers.indexOf('Agenda item description');\n", ''],
  ["        revisedToCol, agendaDescriptionCol\n", "        revisedToCol\n"]
];

const GUARD_SECTION_START = BAR + '// CONFIGURATION GUARD (after T-2026.10.12) -- NO UPDATE OR BUILD WITHOUT A SAVED MEETING\n' + BAR;
const GUARD_SECTION_END = BAR + '// END OF CONFIGURATION GUARD\n' + BAR + '\n';
const GUARD_SECTION_FUNCTIONS = ['hasSavedMeetingConfiguration_', 'assertMeetingConfigured_'];
const GUARD_CHANGED_FUNCTIONS = ['continuousUpdateCore_', 'assertMeetingReadyToBuild_'];
const GUARD_CHANGES = [
  ["    try {\n" +
   "      assertMeetingReadyToBuild_(context);\n" +
   "    } catch (e) {\n" +
   "      // After T-2026.10.12: the caller can say that nothing was changed.\n" +
   "      result.notUpdated = true;\n" +
   "      throw e;\n" +
   "    }\n", "    assertMeetingReadyToBuild_(context);\n"],
  ["    result = Object.assign({ success: false, error: e.message }, result.notUpdated ? { notUpdated: true } : {});\n",
   "    result = { success: false, error: e.message };\n"],
  ["  // After T-2026.10.12: a template report with no saved meeting is not a main meeting by default.\n" +
   "  assertMeetingConfigured_(context);\n", '']
];

function takeOut(code, start, end, changes, name) {
  let text = String(code).replace(/\r/g, '');
  const problems = [];
  if (start !== null) {
    const a = text.indexOf(start);
    const b = text.indexOf(end);
    if (a === -1 || b === -1 || b < a || text.indexOf(start, a + 1) !== -1) {
      problems.push('the ' + name + ' section is not in Code.js exactly once');
    } else {
      text = text.slice(0, a) + text.slice(b + end.length);
    }
  }
  changes.forEach(([now, before], i) => {
    const at = text.indexOf(now);
    if (at === -1 || text.indexOf(now, at + 1) !== -1) { problems.push(name + ' change ' + (i + 1) + ' is not in Code.js exactly once: ' + JSON.stringify(now.slice(0, 60))); return; }
    text = text.slice(0, at) + before + text.slice(at + now.length);
  });
  return { code: text, problems: problems };
}

/** The one change of the guard in template/ReportCreator.js: Update Report Now says the report was not changed. */
const GUARD_CREATOR_CHANGED_FUNCTIONS = ['describeUpdateReportNowFailure_'];
const GUARD_CREATOR_CHANGE = "  // After T-2026.10.12: refused before the report was touched (no saved meeting, or an ad-hoc meeting that is not ready).\n" +
   "  if (result && result.success !== true && result.notUpdated) return 'The report was not updated: ' + (result.error || 'it is not ready to be updated.');\n";

/** template/ReportCreator.js without the guard's change: { code, problems }. */
function creatorWithoutConfigurationGuard(code) {
  const text = String(code).replace(/\r/g, '');
  const at = text.indexOf(GUARD_CREATOR_CHANGE);
  if (at === -1 || text.indexOf(GUARD_CREATOR_CHANGE, at + 1) !== -1) return { code: text, problems: ['the CONFIGURATION GUARD change of ReportCreator.js is not in it exactly once'] };
  return { code: text.slice(0, at) + text.slice(at + GUARD_CREATOR_CHANGE.length), problems: [] };
}

function codeWithoutConfigurationGuard(code) {
  return takeOut(code, GUARD_SECTION_START, GUARD_SECTION_END, GUARD_CHANGES, 'CONFIGURATION GUARD');
}

/** The release of both (T-2026.10.13): the version line and the changelog entry of 2.23.0. */
const RELEASE_ENTRY = " * 2.23.0 (2026-10-14)\n *   - Fixed: the revision placement of every update took each revision for\n *     out of place (Google Docs keeps a paragraph between two tables), moved\n *     it again and left one empty paragraph behind each time -- over a\n *     thousand in a live report. A revision is in place when it is the next\n *     table below its parent; a moved table takes its separator with it.\n *   - Fixed: the heading of an agenda item a report did not have was\n *     appended at the end of the document, after \"Close of the session\" and\n *     the Attendance section, with its TDocs. It is now written in the order\n *     of the agenda, with the title of the item where known; a TDoc of the\n *     closing item goes before the Attendance section.\n *   - Added: an update first repairs what earlier releases left: agenda\n *     sections after the close go back to their place, and a run of empty\n *     paragraphs below a TDoc table becomes one. A report in order is only\n *     read. (REPORT STRUCTURE section; docs/REPORT_STRUCTURE.md.)\n *   - Fixed (template reports): a report with no saved meeting -- a Drive\n *     copy, or Resolve without Save -- is no longer updated or built on the\n *     built-in SA4#136 / 6G defaults. It is refused before anything is read\n *     or changed, with what to do. CENTRAL and Legacy are unchanged.\n *     (CONFIGURATION GUARD section; docs/CONFIGURATION_GUARD.md.)\n";
const RELEASE_CHANGES = [
  [' * Version: 2.23.0 (2026-10-14)\n', ' * Version: 2.22.0 (2026-10-07)\n'],
  [' * CHANGELOG\n' + RELEASE_ENTRY, ' * CHANGELOG\n']
];

/** Code.js without the release header of 2.23.0 only: { code, problems }. */
function codeWithoutReleaseHeader(code) {
  return takeOut(code, null, null, RELEASE_CHANGES, 'RELEASE HEADER');
}

/**
 * The abstract-cache change after T-2026.10.13: the Reviewer "no summary" cache
 * lasts 4 hours, and "Update Abstracts" asks past it. [now, before] pairs.
 */
const ABSTRACT_CACHE_CHANGED_FUNCTIONS = ['fetchAndAddAbstract_', 'addAbstractsForTables_', 'addAbstractsOnly'];
const ABSTRACT_CACHE_CHANGES = [
 [" * Version: 2.24.0 (2026-10-14)\n", " * Version: 2.23.0 (2026-10-14)\n"],
 [" * CHANGELOG\n * 2.24.0 (2026-10-14)\n *   - Changed: the Reviewer's \"no summary\" answer (404) is remembered for\n *     4 hours instead of 24 (REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS). An update\n *     asks again after 4 hours, so an abstract the Reviewer produces soon\n *     after an upload no longer comes most of a meeting day late. Only a\n *     404 is remembered, as before; existing entries keep their key and are\n *     read with the new 4 hours.\n *   - Changed: \"Update Abstracts\" asks the Reviewer for every table without\n *     an Abstract even when a \"no summary\" answer is still remembered, and\n *     logs that it bypasses the negative cache. Updates do not.\n", " * CHANGELOG\n"],
 [
  "//\n// After T-2026.10.13: 4 hours, not 24. The Reviewer often has no summary yet\n// when a TDoc has just been uploaded; with 24 hours its abstract came most\n// of a meeting day late. An hourly update now asks again after 4 hours. The\n// explicit \"Update Abstracts\" does not wait at all (bypassNegativeCache).\nvar REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS = 4;",
  "var REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS = 24;"
 ],
 [
  "function fetchAndAddAbstract_(table, tdocNumber, context, portalStatus, options) {",
  "function fetchAndAddAbstract_(table, tdocNumber, context, portalStatus) {"
 ],
 [
  "    if (cacheId && isReviewerNoSummaryCached_(cacheId, context)) {\n      // After T-2026.10.13: the explicit \"Update Abstracts\" asks anyway\n      // (options.bypassNegativeCache); every other caller waits out the TTL.\n      if (!(options && options.bypassNegativeCache)) {\n        perfCount_('Reviewer negative-cache hits');\n        Logger.log(`Skipping Reviewer API for ${tdocNumber}: no summary was available on a previous run (cached, within TTL)`);\n        return;\n      }\n      perfCount_('Reviewer negative-cache bypasses (Update Abstracts)');\n      Logger.log(`Asking the Reviewer API for ${tdocNumber} although it had no summary on a previous run (cached, within TTL): bypassing the negative cache for a manual Update Abstracts`);\n    }",
  "    if (cacheId && isReviewerNoSummaryCached_(cacheId, context)) {\n      perfCount_('Reviewer negative-cache hits');\n      Logger.log(`Skipping Reviewer API for ${tdocNumber}: no summary was available on a previous run (cached, within TTL)`);\n      return;\n    }"
 ],
 [
  "function addAbstractsForTables_(body, context, leaveOut, portalStatuses, options) {",
  "function addAbstractsForTables_(body, context, leaveOut, portalStatuses) {"
 ],
 [
  "  // `options` goes to fetchAndAddAbstract_() unchanged: only \"Update\n  // Abstracts\" passes { bypassNegativeCache: true } (after T-2026.10.13).\n",
  ""
 ],
 [
  "() => fetchAndAddAbstract_(table, parsedTdoc.raw, context, portalStatus, options));",
  "() => fetchAndAddAbstract_(table, parsedTdoc.raw, context, portalStatus));"
 ],
 [
  "  // After T-2026.10.13: an explicit request, so a TDoc that had no summary on\n  // an earlier run is asked again now -- the negative cache is not waited out.\n  const result = addAbstractsForTables_(null, undefined, null, null, { bypassNegativeCache: true });",
  "  const result = addAbstractsForTables_();"
 ]
];

/** Code.js without the abstract-cache change: { code, problems }. */
function codeWithoutAbstractCacheChange(code) {
  return takeOut(code, null, null, ABSTRACT_CACHE_CHANGES, 'ABSTRACT CACHE');
}

/**
 * The attendance-rendering fix after T-2026.10.14: removeAdhocBodyChild_()
 * no longer calls setText('') on an empty last paragraph (Google Docs: "Cannot
 * insert an empty text element."). [now, before] pairs.
 * tests/adhoc-attendance-empty-paragraph.test.js pins the behaviour.
 */
const ATTENDANCE_EMPTY_TEXT_CHANGED_FUNCTIONS = ['removeAdhocBodyChild_'];
const ATTENDANCE_EMPTY_TEXT_CHANGES = [
 [
  "/**\n * Removes one child of the body. A document must end with a paragraph, so\n * the last paragraph is emptied instead of removed. Google Docs cannot insert\n * an empty text: setText('') throws on a paragraph that has no text -- and the\n * empty paragraph after a table at the end of the document is exactly that --\n * so an empty paragraph is left as it is and any other is cleared.\n */",
  "/**\n * Removes one child of the body. A document must end with a paragraph, so\n * the last paragraph is emptied instead of removed.\n */"
 ],
 [
  "    const paragraph = child.asParagraph();\n    if (paragraph.getText() !== '') paragraph.clear();\n    paragraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);\n",
  "    child.asParagraph().setText('');\n    child.asParagraph().setHeading(DocumentApp.ParagraphHeading.NORMAL);\n"
 ]
];

/** Code.js without the attendance-rendering fix: { code, problems }. */
function codeWithoutAttendanceEmptyTextFix(code) {
  return takeOut(code, null, null, ATTENDANCE_EMPTY_TEXT_CHANGES, 'ATTENDANCE EMPTY TEXT');
}

function codeBeforeStructureFix(code) {
  const attendanceless = codeWithoutAttendanceEmptyTextFix(code);
  const abstractless = codeWithoutAbstractCacheChange(attendanceless.code);
  const unreleased = codeWithoutReleaseHeader(abstractless.code);
  const guardless = codeWithoutConfigurationGuard(unreleased.code);
  const view = takeOut(guardless.code, SECTION_START, SECTION_END, CHANGES, 'REPORT STRUCTURE');
  return { code: view.code, problems: attendanceless.problems.concat(abstractless.problems, unreleased.problems, guardless.problems, view.problems) };
}

module.exports = { codeBeforeStructureFix, codeWithoutConfigurationGuard, codeWithoutReleaseHeader, codeWithoutAbstractCacheChange, ABSTRACT_CACHE_CHANGED_FUNCTIONS, codeWithoutAttendanceEmptyTextFix, ATTENDANCE_EMPTY_TEXT_CHANGED_FUNCTIONS, creatorWithoutConfigurationGuard, GUARD_CREATOR_CHANGED_FUNCTIONS, SECTION_START, SECTION_END, SECTION_FUNCTIONS, CHANGED_FUNCTIONS,
  GUARD_SECTION_START, GUARD_SECTION_END, GUARD_SECTION_FUNCTIONS, GUARD_CHANGED_FUNCTIONS };
