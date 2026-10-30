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

/**
 * template/ReportCreator.js without the guard's change: { code, problems }.
 * The two menu items of the Disposition Summary (after T-2026.10.15), which
 * came later, are taken out first.
 */
function creatorWithoutConfigurationGuard(code) {
  const later = creatorWithoutDispositionSummary(code);
  const text = later.code;
  const at = text.indexOf(GUARD_CREATOR_CHANGE);
  if (at === -1 || text.indexOf(GUARD_CREATOR_CHANGE, at + 1) !== -1) return { code: text, problems: later.problems.concat(['the CONFIGURATION GUARD change of ReportCreator.js is not in it exactly once']) };
  return { code: text.slice(0, at) + text.slice(at + GUARD_CREATOR_CHANGE.length), problems: later.problems };
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
  [" * Version: 2.25.0 (2026-10-16)\n", " * Version: 2.24.0 (2026-10-14)\n"],
  [" * CHANGELOG\n * 2.25.0 (2026-10-16)\n *   - Fixed: Refresh Attendance, and every import that replaced an existing\n *     Attendance section, failed with \"Cannot insert an empty text element.\"\n *     when the section ended the report. Google Docs keeps an empty paragraph\n *     after the last table; removing the section tried to empty it with\n *     setText(''), which Google Docs refuses on a paragraph without text.\n *     An empty last paragraph is now left as it is, any other is cleared.\n *     Stored attendance is unchanged; Refresh Attendance writes the section\n *     again from it.\n", " * CHANGELOG\n"],
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

/**
 * The Disposition Summary and Close Report after T-2026.10.15: one new
 * section, and six released functions that keep the agenda, the close and
 * the Attendance section before the summary, keep a summary up to date on an
 * update, leave its tables out of Report Status Summary, and say after a
 * rebuild that it was cleared. [now, before] pairs.
 * tests/disposition-summary-close-report.test.js pins the behaviour.
 */
const DISPOSITION_SECTION_START = BAR + '// DISPOSITION SUMMARY AND CLOSE REPORT (after T-2026.10.15)\n' + BAR;
const DISPOSITION_SECTION_END = BAR + '// END OF DISPOSITION SUMMARY AND CLOSE REPORT\n' + BAR + '\n';
const DISPOSITION_SECTION_FUNCTIONS = ['classifyTdocDisposition_', 'isDispositionSummaryHeadingText_', 'tdocCellLinkUrl_', 'meetingTdocScope_', 'readMeetingTdocScope_', 'readReportDispositions_', 'buildDispositionSummaryModel_', 'dispositionSummaryStatusText_',
  'dispositionSummaryContent_', 'isDispositionSummaryParagraphText_', 'isDispositionSummaryTable_', 'findDispositionSummarySection_', 'readDispositionSummaryContent_', 'formatDispositionSummaryTable_',
  'renderDispositionSummary_', 'writeDispositionSummary_', 'syncDispositionSummaryOnUpdate_', 'dispositionSummaryRebuildNote_', 'dispositionCountsText_', 'isRevisedInReport_', 'checkRevisionsIn3gu_', 'closeReportCore_',
  'dispositionDialogTitle_', 'describeDispositionSummaryState_', 'otherMeetingsText_', 'formatCloseReportResult_', 'formatDispositionSummaryResult_', 'withDispositionLock_', 'updateDispositionSummary', 'closeReport', 'closeReportResultHtml_'];
const DISPOSITION_CHANGED_FUNCTIONS = ['continuousUpdateCore_', 'findReportTerminalBoundary_', 'findAgendaSectionEndIndex_', 'analyzeReportStatus', 'buildSkeletonWithTdocTables', 'findAdhocAttendanceInsertIndex_'];
const DISPOSITION_CHANGES = [
  // The release of it (T-2026.10.16): the version line and the changelog entry of 2.26.0.
  [" * Version: 2.26.0 (2026-10-30)\n", " * Version: 2.25.0 (2026-10-16)\n"],
  [" * CHANGELOG\n * 2.26.0 (2026-10-30)\n *   - Added (template reports): Report > Update Disposition Summary writes a\n *     managed section at the end of the report -- after the close and the\n *     Attendance section -- with every TDoc once in one of three tables by\n *     its Status: Agreed, Endorsed and Approved (agreed, partially agreed,\n *     endorsed, approved); Other Disposed (noted, revised, merged,\n *     withdrawn, replied); Not Yet Disposed (available, reserved, parked,\n *     postponed, Plenary, other, and any status it does not know). Updates\n *     keep a summary that is there up to date and never add one; text typed\n *     below it is kept, text typed into it stops it from being written.\n *   - Added: Report > Close Report checks the report and says whether it is\n *     ready to close: no TDoc not yet disposed, and for every TDoc revised\n *     in the report a complete revision in 3GU (every revision of \"Revised\n *     to\" read; withdrawn ones aside): not reserved, reserved but not\n *     uploaded, and could not verify are reported apart, with what to do.\n *     It writes the summary and nothing else, stores nothing, and shows the\n *     whole result in a dialog the execution does not wait on.\n *     (DISPOSITION SUMMARY AND CLOSE REPORT section;\n *     docs/DISPOSITION_SUMMARY.md.)\n",
   " * CHANGELOG\n"],
  // continuousUpdateCore_(): a summary that is there is kept up to date.
  // (After T-2026.10.16: for the TDocs of the meeting -- the list the update has read.)
  ["\n    // After T-2026.10.15: a Disposition Summary that is there is kept up to\n" +
   "    // date; a report without one is not given one. Not in a background run\n" +
   "    // (no status dropdowns can be read there). Never fails the update.\n" +
   "    // After T-2026.10.16: for the TDocs of the meeting -- the list read above.\n" +
   "    if (!context) {\n" +
   "      const summary = syncDispositionSummaryOnUpdate_(body, allTdocs);\n" +
   "      if (summary.state !== 'absent') result.dispositionSummary = summary.state;\n" +
   "    }\n", ''],
  // findReportTerminalBoundary_(): without a close and an Attendance section, the summary ends the agenda part.
  ["  let attendance = -1;\n  let summary = -1;\n", "  let attendance = -1;\n"],
  ["    else if (summary === -1 && isDispositionSummaryHeadingText_(text)) summary = i;\n", ''],
  ["  // After T-2026.10.15: without either, the Disposition Summary ends the agenda part.\n" +
   "  if (summary !== -1) return { index: summary, closeNumber: '' };\n", ''],
  // findAgendaSectionEndIndex_(): a TDoc of the closing item goes before the summary.
  ["      // After T-2026.10.15: nor after the Disposition Summary.\n" +
   "      if (isDispositionSummaryHeadingText_(nextText)) return j;\n", ''],
  // analyzeReportStatus(): the summary tables are not counted as TDocs.
  ["    // After T-2026.10.15: the tables of the Disposition Summary list TDocs; they are not TDoc tables.\n" +
   "    if (isDispositionSummaryTable_(t)) return;\n", ''],
  // buildSkeletonWithTdocTables(): the rebuild says that the summary was cleared.
  ["  // After T-2026.10.15: the Disposition Summary is not rebuilt; the result says how to write it again.\n" +
   "  const dispositionSummaryNote = dispositionSummaryRebuildNote_(DocumentApp.getActiveDocument().getBody());\n", ''],
  ["  reallocationRestoreNote += dispositionSummaryNote;\n", ''],
  // findAdhocAttendanceInsertIndex_(): without a close, a new Attendance section goes before the summary.
  ["  if (closing === -1) {\n" +
   "    // After T-2026.10.15: never after the Disposition Summary.\n" +
   "    const summary = findDispositionSummarySection_(body);\n" +
   "    return { index: summary ? summary.start : count, inClosingSection: false };\n" +
   "  }\n", "  if (closing === -1) return { index: count, inClosingSection: false };\n"]
];

/** Code.js without the Disposition Summary and Close Report: { code, problems }. */
/**
 * The TDoc list identity (after T-2026.10.16): a downloaded TDoc list is used
 * only when it is shown to be the meeting's. One section after the Disposition
 * Summary, and its calls in the four places a TDoc list is acquired. It came
 * after the Disposition Summary, so codeWithoutDispositionSummary() takes it
 * out first. codeWithoutListIdentity(code) -> { code, problems }.
 */
const IDENTITY_SECTION_START = BAR + '// TDOC LIST IDENTITY (after T-2026.10.16)\n' + BAR;
const IDENTITY_SECTION_END = BAR + '// END OF TDOC LIST IDENTITY\n' + BAR + '\n';
const IDENTITY_SECTION_FUNCTIONS = ['tdocListRawNumbers_', 'tdocListAnchorFor_', 'chooseTdocListAnchor_', 'evaluateTdocListIdentity_', 'assertTdocListIdentity_'];
const IDENTITY_CHANGED_FUNCTIONS = ['continuousUpdateCore_', 'downloadAndProcessFromWeb', 'buildSkeletonWithTdocTables', 'downloadAndGroupTdocs_'];
const IDENTITY_CHANGES = [
  // The release of it and of the meeting scope of the Disposition Summary (T-2026.10.17): the version line and the
  // changelog entry of 2.26.1.
  [" * Version: 2.26.1 (2026-10-30)\n", " * Version: 2.26.0 (2026-10-30)\n"],
  [" * CHANGELOG\n * 2.26.1 (2026-10-30)\n *   - Fixed (template reports): the Disposition Summary and Close Report\n *     read only the TDocs of this meeting -- its TDoc list as an update\n *     reads it (agenda selector, Document Reallocations) -- not every TDoc\n *     table of the document; tables of other meetings are left out,\n *     counted and not touched. Without that list (not readable, empty, or\n *     with none of the report's TDocs) nothing is written and Close Report\n *     is not ready. (docs/DISPOSITION_SUMMARY.md.)\n *   - Fixed: a downloaded TDoc list is used only when it is shown to be the\n *     list of this meeting: the saved agenda TDoc has a row in it, else a\n *     TDoc of a list accepted before (TDOC_LIST_ANCHOR), else -- the first\n *     time -- every TDoc table of the report is in it. Otherwise the\n *     update, the build, the summary, Close Report and the session actions\n *     change nothing. (6G, 2026-10-30: the Portal list of the meeting ID\n *     gave the TDocs of earlier calls, and updates inserted them.) (TDOC\n *     LIST IDENTITY section; docs/TDOC_LIST_IDENTITY.md.)\n",
   " * CHANGELOG\n"],
  // continuousUpdateCore_(): a refused list is a run that changed nothing.
  ["    // After T-2026.10.16: a list that is not the meeting's changes nothing.\n" +
   "    let tdocGroups;\n" +
   "    try {\n" +
   "      tdocGroups = perfTimed_('TDoc-list fetch/download+parse (downloadAndGroupTdocs_)', () => downloadAndGroupTdocs_(cfg, context, undefined, 'update'));\n" +
   "    } catch (e) {\n" +
   "      if (e && e.tdocListIdentity) result.notUpdated = true;\n" +
   "      throw e;\n" +
   "    }\n",
   "    const tdocGroups = perfTimed_('TDoc-list fetch/download+parse (downloadAndGroupTdocs_)', () => downloadAndGroupTdocs_(cfg, context));\n"],
  // downloadAndProcessFromWeb()
  ["    // After T-2026.10.16: the list of this meeting, or nothing.\n" +
   "    try {\n" +
   "      assertTdocListIdentity_(sheet.getDataRange().getValues(), cfg, undefined, 'read');\n" +
   "    } catch (e) {\n" +
   "      tempFile.setTrashed(true);\n" +
   "      throw e;\n" +
   "    }\n    \n", ''],
  // buildSkeletonWithTdocTables()
  ["  // After T-2026.10.16: refused here, before the document is cleared, when it is not the meeting's list.\n" +
   "  const tdocGroups = downloadAndGroupTdocs_(cfg, undefined, savedReallocations, 'build');\n",
   "  const tdocGroups = downloadAndGroupTdocs_(cfg, undefined, savedReallocations);\n"],
  // downloadAndGroupTdocs_(): its header, its signature and the check.
  [" * After T-2026.10.16: only a list shown to be the meeting's (assertTdocListIdentity_(); `use` 'update' |\n" +
   " * 'build' | 'read', default 'read'); else it throws before anything is grouped.\n */\n" +
   "function downloadAndGroupTdocs_(cfg, context, reallocationsSnapshot, use) {",
   " */\nfunction downloadAndGroupTdocs_(cfg, context, reallocationsSnapshot) {"],
  ["    // After T-2026.10.16: the list of this meeting, on its raw rows, or nothing.\n" +
   "    assertTdocListIdentity_(data, cfg, context, use);\n", '']
];

/**
 * The report-local scope of the Disposition Summary (after T-2026.10.17): the
 * summary reads the TDocs the report has registered (reportTdocScope_()), not
 * the TDoc list; Close Report checks the list on top and is not ready without
 * it. Inside the Disposition Summary section, and its call in
 * continuousUpdateCore_. It came after the TDoc list identity, so
 * codeWithoutListIdentity() takes it out first. [now, before] pairs, one per
 * hunk; together they give the Code.js of T-2026.10.17, byte for byte.
 * tests/disposition-report-local-scope.test.js pins the behaviour.
 */
const REPORT_LOCAL_SCOPE_CHANGES = [
  ["    // (no status dropdowns can be read there). Never fails the update.\n    // After T-2026.10.17: for the TDocs the report has registered -- the\n    // TDocs of the list read above among them.\n    if (!context) {\n      const summary = syncDispositionSummaryOnUpdate_(body);\n      if (summary.state !== 'absent') result.dispositionSummary = summary.state;\n",
   "    // (no status dropdowns can be read there). Never fails the update.\n    // After T-2026.10.16: for the TDocs of the meeting -- the list read above.\n    if (!context) {\n      const summary = syncDispositionSummaryOnUpdate_(body, allTdocs);\n      if (summary.state !== 'absent') result.dispositionSummary = summary.state;\n"],
  ["//\n// WHICH TDOCS (after T-2026.10.17): those the REPORT has registered -- a row\n// in its own \"Registration of Documents\" table, and no reallocation to\n// removed/withdrawn/n/a in its own Document Reallocations table\n// (reportTdocScope_()). Every build writes that table from its list, every\n// update appends the TDocs it inserts, nothing removes a row: it is the\n// report's record of what it took on, for every call of a continuing report.\n// The summary is report-local: no TDoc list is downloaded for it, so a Portal\n// list that cannot be read or verified does not stop it. Without a\n// registration table there is no scope -- never every table of the document.\n// Close Report checks the TDoc list (3GU) on top, and is not ready without it.\n//\n// WHERE: a new summary is appended at the end of the document -- after\n",
   "//\n// WHERE: a new summary is appended at the end of the document -- after\n"],
  ["//   2. a TDoc the report has as revised (its Status; nothing else) does not\n//      have a complete revision in the TDoc list (3GU) -- the list is read\n//      every time, and a list that cannot be read or shown to be the\n//      meeting's makes the report not ready:\n//        not reserved   no \"Revised to\" and the list still has it as\n",
   "//   2. a TDoc the report has as revised (its Status; nothing else) does not\n//      have a complete revision in the TDoc list (3GU), which reads the TDoc\n//      list only when the report has such a TDoc:\n//        not reserved   no \"Revised to\" and the list still has it as\n"],
  ["/**\n * Pure: the TDocs of the configured meeting's TDoc list, as the update reads\n * them (Close Report's external check since T-2026.10.17; the summary reads\n * reportTdocScope_()): the rows of its TDoc list\n * (downloadAndGroupTdocs_(): the meeting's list, its agenda selector, its\n",
   "/**\n * Pure: THE scope of the summary and of Close Report -- the TDocs of the\n * configured meeting, as the update reads them: the rows of its TDoc list\n * (downloadAndGroupTdocs_(): the meeting's list, its agenda selector, its\n"],
  ["/**\n * THE scope of the summary (after T-2026.10.17), from the report alone: the\n * TDocs of its \"Registration of Documents\" table(s) (registrationTableColumns_()),\n * less those its Document Reallocations remove. { ok, allTdocs: [], numbers:\n * { canonical number: true }, error }, the form of meetingTdocScope_().\n * Nothing is downloaded and nothing is changed. Not ok without a registration\n * table or a TDoc in it.\n */\nfunction reportTdocScope_(body) {\n  const numbers = {};\n  const removed = {};\n  let registrations = 0;\n  body.getTables().forEach(function (table) {\n    if (registrationTableColumns_(table)) {\n      registrations++;\n      for (let r = 1; r < table.getNumRows(); r++) {\n        const parsed = parseExactSA4DocumentId_(table.getRow(r).getCell(0).getText().trim());\n        if (parsed.isValid) numbers[parsed.raw] = true;\n      }\n    } else if (isReallocationTable_(table)) {\n      for (let r = 1; r < table.getNumRows(); r++) {\n        const row = table.getRow(r);\n        if (row.getNumCells() < 3) continue;\n        const parsed = parseExactSA4DocumentId_(row.getCell(0).getText().trim());\n        if (parsed.isValid && interpretReallocationTarget_(row.getCell(2).getText()).kind === 'remove') removed[parsed.raw] = true;\n      }\n    }\n  });\n  Object.keys(removed).forEach(function (n) { delete numbers[n]; });\n  const ok = Object.keys(numbers).length > 0;\n  const error = ok ? '' : registrations ? 'the Registration of Documents table of this report has no TDoc' : 'this report has no Registration of Documents table';\n  return { ok: ok, allTdocs: [], numbers: numbers, error: error };\n}\n\n/**\n * Reads the TDoc list of the configured meeting (meetingTdocScope_()): the\n * configuration guard first -- never the built-in defaults -- then the\n * meeting's TDoc list, once, shown to be the meeting's (TDOC LIST\n * IDENTITY). Close Report's external check. Never throws: { ok: false,\n * error } when it cannot be had.\n */\n",
   "/**\n * Reads the scope of the configured meeting (meetingTdocScope_()): the\n * configuration guard first -- never the built-in defaults -- then the\n * meeting's TDoc list, once. Never throws: { ok: false, error } when it\n * cannot be had.\n */\n"],
  ["/**\n * The TDocs of the report that are in options.scope (required; from\n * reportTdocScope_()), and their dispositions, in the order of the\n * report: { entries, unreadable, otherMeetings }. One entry per TDoc number:\n",
   "/**\n * The TDocs of the report that are TDocs of the meeting (options.scope, from\n * meetingTdocScope_(); required), and their dispositions, in the order of the\n * report: { entries, unreadable, otherMeetings }. One entry per TDoc number:\n"],
  [" * `otherMeetings`: the TDoc numbers of the tables left out because the\n * scope does not have them (not registered in the report, removed by a\n * reallocation); their tables are not read further. `revisionIds` (only\n * with options.revisions): the other TDoc numbers its Disposition and\n",
   " * `otherMeetings`: the TDoc numbers of the tables left out because the\n * meeting's list does not have them (earlier calls of the series, removed by\n * a reallocation); their tables are not read further. `revisionIds` (only\n * with options.revisions): the other TDoc numbers its Disposition and\n"],
  ["  const read = readReportDispositions_(body, options);\n  // A scope that has none of the TDocs of the report's TDoc tables does not\n  // describe this report (6G smoke copy, 2026-10-30: the Portal list of the\n",
   "  const read = readReportDispositions_(body, options);\n  // A list that has none of the TDocs of the report's TDoc tables does not\n  // describe this report (6G smoke copy, 2026-10-30: the Portal list of the\n"],
  ["  if (!read.entries.length && read.otherMeetings.length) {\n    return { state: 'no-scope', error: 'none of the ' + read.otherMeetings.length + ' TDocs of the report\\'s TDoc tables is registered in this report (' +\n      Object.keys(scope.numbers).length + ' TDocs in its Registration of Documents table): it does not describe these TDoc tables', read: read, model: buildDispositionSummaryModel_([]), foreign: [] };\n  }\n",
   "  if (!read.entries.length && read.otherMeetings.length) {\n    return { state: 'no-scope', error: 'none of the ' + read.otherMeetings.length + ' TDocs of the report\\'s TDoc tables is in the TDoc list of this meeting (' +\n      Object.keys(scope.numbers).length + ' TDocs): the list does not describe this report', read: read, model: buildDispositionSummaryModel_([]), foreign: [] };\n  }\n"],
  ["/**\n * The update: a summary that is there is brought up to date, for the TDocs the\n * report has registered (reportTdocScope_(), the same scope as the menu; the\n * update has registered the TDocs of its list just before). A report without\n * one is not given one; without a scope the summary is left as it is. Never\n * throws. Returns { state }.\n */\nfunction syncDispositionSummaryOnUpdate_(body) {\n  try {\n    if (!findDispositionSummarySection_(body)) return { state: 'absent' };\n    const done = writeDispositionSummary_(body, { scope: reportTdocScope_(body) });\n    if (done.state === 'no-scope') Logger.log('Disposition Summary: not updated -- ' + done.error + '.');\n",
   "/**\n * The update: a summary that is there is brought up to date, for the TDocs of\n * the meeting -- `allTdocs`, the TDoc list the update has just read; nothing\n * is downloaded again. A report without one is not given one; with no TDoc\n * in the list the summary is left as it is. Never throws. Returns { state }.\n */\nfunction syncDispositionSummaryOnUpdate_(body, allTdocs) {\n  try {\n    if (!findDispositionSummarySection_(body)) return { state: 'absent' };\n    const done = writeDispositionSummary_(body, { scope: meetingTdocScope_(allTdocs) });\n    if (done.state === 'no-scope') Logger.log('Disposition Summary: not updated -- ' + done.error + '.');\n"],
  ["    else if (done.state === 'refused') Logger.log('Disposition Summary: not updated -- text was typed into it, above a part of it: ' + done.foreign.join(' | '));\n    else Logger.log('Disposition Summary: ' + done.state + ' (' + dispositionCountsText_(done.model) + '; ' + done.read.otherMeetings.length + ' TDoc table(s) not registered in this report left out)');\n    return { state: done.state };\n",
   "    else if (done.state === 'refused') Logger.log('Disposition Summary: not updated -- text was typed into it, above a part of it: ' + done.foreign.join(' | '));\n    else Logger.log('Disposition Summary: ' + done.state + ' (' + dispositionCountsText_(done.model) + '; ' + done.read.otherMeetings.length + ' TDoc table(s) of other meetings left out)');\n    return { state: done.state };\n"],
  [" * Close Report, without the dialog: writes the summary and runs the checks.\n * 1. the summary and the TDocs not yet disposed, from the report alone\n *    (reportTdocScope_());\n * 2. the external check: the TDoc list of the meeting (3GU), shown to be the\n *    meeting's, and the revisions of the TDocs of 1. against it. Required:\n *    without it the report is never ready, and the result says so.\n * Returns { ready, checked, counts, unresolved, revisions, problems,\n * summaryState, otherMeetings, externalVerified }. Never changes a status, a\n * revision, 3GU or anything stored; the summary is the only thing written.\n */\n",
   " * Close Report, without the dialog: writes the summary and runs the checks.\n * Returns { ready, checked, counts, unresolved, revisions, problems,\n * summaryState }. Never changes a status, a revision, 3GU or anything\n * stored; the summary is the only thing written.\n */\n"],
  ["  const body = DocumentApp.getActiveDocument().getBody();\n  const result = { ready: false, checked: 0, counts: {}, unresolved: [], revisions: [], problems: [], summaryState: '', otherMeetings: 0, externalVerified: false };\n  // ONE report scope for the summary, the TDocs not yet disposed and the revision check.\n  const summary = writeDispositionSummary_(body, { revisions: true, scope: reportTdocScope_(body) });\n  result.summaryState = summary.state;\n",
   "  const body = DocumentApp.getActiveDocument().getBody();\n  const result = { ready: false, checked: 0, counts: {}, unresolved: [], revisions: [], problems: [], summaryState: '', otherMeetings: 0 };\n  // ONE scope for everything below: the summary, the TDocs not yet disposed, the revision check.\n  const scope = readMeetingTdocScope_();\n  const summary = writeDispositionSummary_(body, { revisions: true, scope: scope });\n  result.summaryState = summary.state;\n"],
  ["  if (summary.state === 'no-scope') {\n    result.problems.push('The report could not be checked and the Disposition Summary was not written: ' + summary.error + '. Check the Registration of Documents table of the report and run Close Report again.');\n  } else if (summary.state === 'unreadable') {\n",
   "  if (summary.state === 'no-scope') {\n    result.problems.push('The report could not be checked and the Disposition Summary was not written: ' + summary.error + '. Run Close Report again once the TDoc list can be read.');\n  } else if (summary.state === 'unreadable') {\n"],
  ["\n  // The external check, always: the TDoc list of the meeting, read once (the configuration guard, the identity of the list).\n  // A list that cannot be read or shown to be the meeting's: every revised TDoc is \"could not be verified\", and not ready.\n  const list = readMeetingTdocScope_();\n  result.externalVerified = list.ok;\n  if (!list.ok) {\n    result.problems.push('External verification could not be completed: ' + list.error + '. The Disposition Summary is written from the report; the revisions were not checked in 3GU. Run Close Report again once the TDoc list of this meeting can be read.');\n  }\n  result.revisions = checkRevisionsIn3gu_(summary.read.entries, list.ok ? list.allTdocs : null, list.error);\n  REVISION_CHECK_STATES_.forEach(function (s) { result.counts[s.kind] = result.revisions.filter(function (r) { return r.kind === s.kind; }).length; });\n\n  result.ready = result.externalVerified && summary.state !== 'unreadable' && summary.state !== 'no-scope' && result.problems.length === 0 && result.unresolved.length === 0 && result.revisions.length === 0;\n  Logger.log('Close Report: ' + (result.ready ? 'ready to close' : 'NOT ready to close') + ' -- ' + result.checked + ' TDocs, ' + result.unresolved.length + ' not yet disposed, ' +\n    REVISION_CHECK_STATES_.map(function (s) { return result.counts[s.kind] + ' ' + s.short; }).join(', ') + ', ' + result.problems.length + ' problem(s); summary ' + result.summaryState +\n    '; external verification ' + (result.externalVerified ? 'completed' : 'NOT completed'));\n  return result;\n",
   "\n  // The revision check: the TDocs of the meeting only (the entries above), against the same TDoc list -- not read again.\n  if (scope.ok) result.revisions = checkRevisionsIn3gu_(summary.read.entries, scope.allTdocs, '');\n  REVISION_CHECK_STATES_.forEach(function (s) { result.counts[s.kind] = result.revisions.filter(function (r) { return r.kind === s.kind; }).length; });\n\n  result.ready = summary.state !== 'unreadable' && summary.state !== 'no-scope' && result.problems.length === 0 && result.unresolved.length === 0 && result.revisions.length === 0;\n  Logger.log('Close Report: ' + (result.ready ? 'ready to close' : 'NOT ready to close') + ' -- ' + result.checked + ' TDocs, ' + result.unresolved.length + ' not yet disposed, ' +\n    REVISION_CHECK_STATES_.map(function (s) { return result.counts[s.kind] + ' ' + s.short; }).join(', ') + ', ' + result.problems.length + ' problem(s); summary ' + result.summaryState);\n  return result;\n"],
  ["\n/** Pure: the line that says how many TDoc tables outside the scope were left out ('' for none). */\nfunction otherMeetingsText_(count) {\n  return count ? count + ' TDoc table' + (count === 1 ? '' : 's') + ' not registered in this report (not in its Registration of Documents table, or reallocated to Removed) left out.' : '';\n}\n\n/** Pure: the line of the Close Report dialog that says whether the external check was done. */\nfunction externalVerificationText_(result) {\n  return result.externalVerified ? 'External verification (TDoc list of this meeting, 3GU): completed.' : 'External verification (TDoc list of this meeting, 3GU): NOT completed.';\n}\n",
   "\n/** Pure: the line that says how many TDoc tables of other meetings were left out ('' for none). */\nfunction otherMeetingsText_(count) {\n  return count ? count + ' TDoc table' + (count === 1 ? '' : 's') + ' of other meetings (not in this meeting\\'s TDoc list) left out.' : '';\n}\n"],
  ["  if (result.otherMeetings) lines.push(otherMeetingsText_(result.otherMeetings));\n  lines.push(externalVerificationText_(result));\n  if (result.unresolved.length) {\n",
   "  if (result.otherMeetings) lines.push(otherMeetingsText_(result.otherMeetings));\n  if (result.unresolved.length) {\n"],
  ["function formatDispositionSummaryResult_(done) {\n  if (done.state === 'no-scope') return 'The Disposition Summary was not written: ' + done.error + '. Nothing was changed.\\n\\nThe summary lists the TDocs registered in this report (its Registration of Documents table); without them it is not written.';\n  if (done.state === 'unreadable') return 'The Disposition Summary was not written: the Status of ' + done.read.unreadable.join(', ') + ' could not be read. Try again in a minute.';\n",
   "function formatDispositionSummaryResult_(done) {\n  if (done.state === 'no-scope') return 'The Disposition Summary was not written: ' + done.error + '. Nothing was changed.\\n\\nThe summary lists the TDocs of the configured meeting only; without its TDoc list it is not written.';\n  if (done.state === 'unreadable') return 'The Disposition Summary was not written: the Status of ' + done.read.unreadable.join(', ') + ' could not be read. Try again in a minute.';\n"],
  ["  const done = withDispositionLock_('Update Disposition Summary', function () {\n    // The TDocs the report has registered: from the report alone, no TDoc list is read.\n    const body = DocumentApp.getActiveDocument().getBody();\n    return writeDispositionSummary_(body, { scope: reportTdocScope_(body) });\n  });\n",
   "  const done = withDispositionLock_('Update Disposition Summary', function () {\n    // The TDocs of the configured meeting: its TDoc list, read once.\n    return writeDispositionSummary_(DocumentApp.getActiveDocument().getBody(), { scope: readMeetingTdocScope_() });\n  });\n"],
  ["  if (result.otherMeetings) out.push('<div>' + esc(otherMeetingsText_(result.otherMeetings)) + '</div>');\n  out.push('<div class=\"' + (result.externalVerified ? 'ready' : 'notready') + '\">' + esc(externalVerificationText_(result)) + '</div>');\n  if (result.unresolved.length) {\n",
   "  if (result.otherMeetings) out.push('<div>' + esc(otherMeetingsText_(result.otherMeetings)) + '</div>');\n  if (result.unresolved.length) {\n"]
];
const REPORT_LOCAL_SCOPE_NEW_FUNCTIONS = ['reportTdocScope_', 'externalVerificationText_'];
const REPORT_LOCAL_SCOPE_CHANGED_FUNCTIONS = ['continuousUpdateCore_', 'writeDispositionSummary_', 'syncDispositionSummaryOnUpdate_', 'closeReportCore_', 'otherMeetingsText_', 'formatCloseReportResult_',
  'formatDispositionSummaryResult_', 'updateDispositionSummary', 'closeReportResultHtml_'];

/** Code.js without the report-local scope: the Code.js of T-2026.10.17. { code, problems }. */
function codeWithoutReportLocalScope(code) {
  return takeOut(code, null, null, REPORT_LOCAL_SCOPE_CHANGES, 'REPORT LOCAL SCOPE');
}

function codeWithoutListIdentity(code) {
  const later = codeWithoutReportLocalScope(code);
  const view = takeOut(later.code, IDENTITY_SECTION_START, IDENTITY_SECTION_END, IDENTITY_CHANGES, 'TDOC LIST IDENTITY');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

function codeWithoutDispositionSummary(code) {
  const later = codeWithoutListIdentity(code);
  const view = takeOut(later.code, DISPOSITION_SECTION_START, DISPOSITION_SECTION_END, DISPOSITION_CHANGES, 'DISPOSITION SUMMARY');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** The one change of the Disposition Summary in template/ReportCreator.js: its two menu items, a group of their own. */
const DISPOSITION_CREATOR_CHANGED_FUNCTIONS = ['buildTemplateReportMenu_'];
const DISPOSITION_CREATOR_CHANGE = "    // After T-2026.10.15: the Disposition Summary at the end of the report, and the checks before it is closed.\n" +
   "    .addItem('Update Disposition Summary', 'updateDispositionSummary')\n" +
   "    .addItem('Close Report', 'closeReport')\n" +
   "    .addSeparator()\n";

/** template/ReportCreator.js without the two menu items of the Disposition Summary: { code, problems }. */
function creatorWithoutDispositionSummary(code) {
  const text = String(code).replace(/\r/g, '');
  const at = text.indexOf(DISPOSITION_CREATOR_CHANGE);
  if (at === -1 || text.indexOf(DISPOSITION_CREATOR_CHANGE, at + 1) !== -1) return { code: text, problems: ['the DISPOSITION SUMMARY change of ReportCreator.js is not in it exactly once'] };
  return { code: text.slice(0, at) + text.slice(at + DISPOSITION_CREATOR_CHANGE.length), problems: [] };
}

function codeBeforeStructureFix(code) {
  const dispositionless = codeWithoutDispositionSummary(code);
  const attendanceless = codeWithoutAttendanceEmptyTextFix(dispositionless.code);
  const abstractless = codeWithoutAbstractCacheChange(attendanceless.code);
  const unreleased = codeWithoutReleaseHeader(abstractless.code);
  const guardless = codeWithoutConfigurationGuard(unreleased.code);
  const view = takeOut(guardless.code, SECTION_START, SECTION_END, CHANGES, 'REPORT STRUCTURE');
  return { code: view.code, problems: dispositionless.problems.concat(attendanceless.problems, abstractless.problems, unreleased.problems, guardless.problems, view.problems) };
}

module.exports = { codeWithoutReportLocalScope, REPORT_LOCAL_SCOPE_NEW_FUNCTIONS, REPORT_LOCAL_SCOPE_CHANGED_FUNCTIONS, codeWithoutListIdentity, IDENTITY_SECTION_START, IDENTITY_SECTION_END, IDENTITY_SECTION_FUNCTIONS, IDENTITY_CHANGED_FUNCTIONS, codeWithoutDispositionSummary, creatorWithoutDispositionSummary, DISPOSITION_SECTION_START, DISPOSITION_SECTION_END, DISPOSITION_SECTION_FUNCTIONS, DISPOSITION_CHANGED_FUNCTIONS, DISPOSITION_CREATOR_CHANGED_FUNCTIONS,
  codeBeforeStructureFix,codeWithoutConfigurationGuard, codeWithoutReleaseHeader, codeWithoutAbstractCacheChange, ABSTRACT_CACHE_CHANGED_FUNCTIONS, codeWithoutAttendanceEmptyTextFix, ATTENDANCE_EMPTY_TEXT_CHANGED_FUNCTIONS, creatorWithoutConfigurationGuard, GUARD_CREATOR_CHANGED_FUNCTIONS, SECTION_START, SECTION_END, SECTION_FUNCTIONS, CHANGED_FUNCTIONS,
  GUARD_SECTION_START, GUARD_SECTION_END, GUARD_SECTION_FUNCTIONS, GUARD_CHANGED_FUNCTIONS };
