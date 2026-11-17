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
  // The release of it (T-2026.10.18): the version line and the changelog entry of 2.26.2.
  [" * Version: 2.26.2 (2026-10-30)\n",
   " * Version: 2.26.1 (2026-10-30)\n"],
  [" * CHANGELOG\n * 2.26.2 (2026-10-30)\n *   - Changed (template reports): the Disposition Summary is report-local.\n *     It lists the TDoc tables whose TDoc the report has registered (its\n *     Registration of Documents table), less those its Document\n *     Reallocations remove; no TDoc list is read for it, so a Portal list\n *     that cannot be read or is not the meeting's neither stops it nor gets\n *     into it. Without a registration table nothing is written.\n *   - Changed: Close Report writes that summary, then always checks the\n *     TDoc list of the meeting (3GU) for the revisions; without that check\n *     it is never ready and says \"External verification ... NOT\n *     completed\". The update and its TDoc list identity are unchanged.\n *     (docs/DISPOSITION_SUMMARY.md § Which TDocs.)\n",
   " * CHANGELOG\n"],
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

/**
 * Close Report's Portal-page fallback (after T-2026.10.18): when the configured
 * TDoc list cannot be had or is refused by its identity, Close Report -- and
 * nothing else -- reads the Portal's TDoc list page through the same guard.
 * One section after the TDoc list identity, and three functions of the
 * Disposition Summary section (readMeetingTdocScope_, closeReportCore_,
 * externalVerificationText_). It came after the report-local scope, so
 * codeWithoutReportLocalScope() takes it out first. [now, before] pairs; with
 * the section they give the Code.js of T-2026.10.18 byte for byte.
 * tests/close-report-portal-page-fallback.test.js pins the behaviour.
 */
const PAGE_FALLBACK_SECTION_START = BAR + '// CLOSE REPORT: THE PORTAL TDOC LIST PAGE AS FALLBACK EVIDENCE (after T-2026.10.18)\n' + BAR;
const PAGE_FALLBACK_SECTION_END = BAR + '// END OF CLOSE REPORT PORTAL PAGE FALLBACK\n' + BAR + '\n';
const PAGE_FALLBACK_SECTION_FUNCTIONS = ['portalTdocListPageCellText_', 'parsePortalTdocListPage_', 'portalTdocListPageRows_', 'portalTdocListPageRecords_', 'readPortalTdocListPageScope_'];
const PAGE_FALLBACK_CHANGED_FUNCTIONS = ['readMeetingTdocScope_', 'closeReportCore_', 'externalVerificationText_'];
const PAGE_FALLBACK_CHANGES = [
  // The release of it (T-2026.10.19): the version line and the changelog entry of 2.26.3.
  [" * Version: 2.26.3 (2026-11-01)\n",
   " * Version: 2.26.2 (2026-10-30)\n"],
  [" * CHANGELOG\n * 2.26.3 (2026-11-01)\n *   - Added (template reports): Close Report's external check has a second\n *     source of evidence. When the configured TDoc list (the Portal's Excel\n *     export) cannot be had or is not shown to be the meeting's, Close\n *     Report -- and nothing else -- reads the Portal's TDoc list page\n *     (TdocList.aspx?meetingId=...), through the SAME identity guard, the\n *     same agenda selector and Document Reallocations, and the same\n *     revision check. The page is used only when it is one complete list\n *     (every row, the reserved TDocs as disabled links, every number of\n *     \"Revised To\"); the two sources are never combined; the dialog and the\n *     log name the source. The update, the build, the legacy import and the\n *     Disposition Summary are unchanged. (6G, 2026-10-30: the export served\n *     the TDocs of earlier calls while the page showed the meeting's 116.)\n *     (CLOSE REPORT: THE PORTAL TDOC LIST PAGE AS FALLBACK EVIDENCE\n *     section; docs/DISPOSITION_SUMMARY.md.)\n",
   " * CHANGELOG\n"],
  ["function readMeetingTdocScope_() {\n  let cfg;\n  try {\n    assertMeetingConfigured_();\n    cfg = getReportConfig_();\n    if (!cfg.TDOC_LIST_URL) throw new Error('this report has no TDoc list configured.');\n  } catch (e) {\n    return { ok: false, allTdocs: [], numbers: {}, error: 'the TDoc list of this meeting could not be read: ' + e.message, source: '' };\n  }\n  let excelProblem;\n  try {\n    const groups = downloadAndGroupTdocs_(cfg);\n",
   "function readMeetingTdocScope_() {\n  try {\n    assertMeetingConfigured_();\n    const cfg = getReportConfig_();\n    if (!cfg.TDOC_LIST_URL) throw new Error('this report has no TDoc list configured.');\n    const groups = downloadAndGroupTdocs_(cfg);\n"],
  ["    Object.keys(groups).forEach(function (key) { groups[key].tdocs.forEach(function (td) { allTdocs.push(td); }); });\n    const scope = meetingTdocScope_(allTdocs);\n    scope.source = EXTERNAL_SOURCE_EXCEL_;\n    Logger.log('External TDoc verification source: Excel (the configured TDoc list)');\n    return scope;\n  } catch (e) {\n    excelProblem = 'the TDoc list of this meeting could not be read: ' + e.message;\n  }\n  // After T-2026.10.18: the configured list could not be had, or is not shown to be the meeting's -- the Portal's TDoc list\n  // page is tried, through the same identity guard (CLOSE REPORT PORTAL PAGE FALLBACK section). Close Report only.\n  return readPortalTdocListPageScope_(cfg, excelProblem);\n}\n",
   "    Object.keys(groups).forEach(function (key) { groups[key].tdocs.forEach(function (td) { allTdocs.push(td); }); });\n    return meetingTdocScope_(allTdocs);\n  } catch (e) {\n    return { ok: false, allTdocs: [], numbers: {}, error: 'the TDoc list of this meeting could not be read: ' + e.message };\n  }\n}\n"],
  ["  result.externalVerified = list.ok;\n  result.externalSource = list.ok ? list.source || '' : '';\n  if (!list.ok) {\n",
   "  result.externalVerified = list.ok;\n  if (!list.ok) {\n"],
  ["    REVISION_CHECK_STATES_.map(function (s) { return result.counts[s.kind] + ' ' + s.short; }).join(', ') + ', ' + result.problems.length + ' problem(s); summary ' + result.summaryState +\n    '; external verification ' + (result.externalVerified ? 'completed (' + result.externalSource + ')' : 'NOT completed'));\n  return result;\n",
   "    REVISION_CHECK_STATES_.map(function (s) { return result.counts[s.kind] + ' ' + s.short; }).join(', ') + ', ' + result.problems.length + ' problem(s); summary ' + result.summaryState +\n    '; external verification ' + (result.externalVerified ? 'completed' : 'NOT completed'));\n  return result;\n"],
  ["function externalVerificationText_(result) {\n  if (!result.externalVerified) return 'External verification (TDoc list of this meeting, 3GU): NOT completed.';\n  // After T-2026.10.18: the source is named when it is not the configured list.\n  return 'External verification (TDoc list of this meeting, 3GU): completed.' +\n    (result.externalSource === EXTERNAL_SOURCE_PORTAL_PAGE_ ? ' Source: the Portal TDoc list page (TdocList.aspx); the configured TDoc list could not be used.' : '');\n}\n",
   "function externalVerificationText_(result) {\n  return result.externalVerified ? 'External verification (TDoc list of this meeting, 3GU): completed.' : 'External verification (TDoc list of this meeting, 3GU): NOT completed.';\n}\n"]
];

/**
 * The agenda Word conversion (after T-2026.10.19, MBS SA4#138): the agenda
 * TDoc's Word file is sent to Drive with a Word content type, the result is
 * checked to be a Google Doc, it is opened with bounded retries, every step
 * is logged and the temporary document is always trashed. One section after
 * convertWordBlobToGoogleDoc_(), and two functions changed
 * (parseAgendaWordBlob_, convertWordBlobToGoogleDoc_). It came after the
 * Portal-page fallback, so codeWithoutCloseReportPageFallback() takes it out
 * first. [now, before] pairs; with the section they give the Code.js of
 * T-2026.10.19 byte for byte. tests/agenda-word-conversion.test.js pins the
 * behaviour.
 */
const AGENDA_CONVERSION_SECTION_START = BAR + '// AGENDA WORD CONVERSION (after T-2026.10.19)\n' + BAR;
const AGENDA_CONVERSION_SECTION_END = BAR + '// END OF AGENDA WORD CONVERSION\n' + BAR + '\n';
const AGENDA_CONVERSION_SECTION_FUNCTIONS = ['agendaWordContentType_', 'agendaErrorForLog_', 'convertAgendaWordBlob_', 'openConvertedAgendaBody_', 'trashConvertedAgenda_', 'diagnoseAgendaConversion'];
const AGENDA_CONVERSION_CHANGED_FUNCTIONS = ['parseAgendaWordBlob_', 'convertWordBlobToGoogleDoc_'];
const AGENDA_CONVERSION_CHANGES = [
  // The release of it (T-2026.10.20): the version line and the changelog entry of 2.26.4.
  [" * Version: 2.26.4 (2026-11-11)\n",
   " * Version: 2.26.3 (2026-11-01)\n"],
  [" * CHANGELOG\n * 2.26.4 (2026-11-11)\n *   - Fixed: the agenda TDoc's Word file is converted reliably. It goes to\n *     Drive with the content type of a Word file (by its extension), the\n *     result must be a native Google Doc (otherwise refused, naming its\n *     type), and it is opened with bounded retries (0/1/2/4 s). Every step\n *     is logged as [AGENDA] with the document ID, every failure with its\n *     original error and stack; the temporary document is always trashed.\n *     diagnoseAgendaConversion(), run from the editor, does the same steps\n *     for the configured agenda TDoc and writes nothing. How the agenda is\n *     parsed is unchanged. (MBS SA4#138, 2026-11-11: two Full Builds failed\n *     at DocumentApp.openById on the converted agenda S4-261857.) (AGENDA\n *     WORD CONVERSION section; docs/AGENDA_WORD_CONVERSION.md.)\n",
   " * CHANGELOG\n"],
  ["  // After T-2026.10.19: opened with bounded retries, trashed in every case (AGENDA WORD CONVERSION section).\n" +
   "  try {\n    return parseAgendaStructureWithText_(openConvertedAgendaBody_(convertedDocId), '');\n  } finally {\n    trashConvertedAgenda_(convertedDocId);\n  }\n",
   "  try {\n    const doc = DocumentApp.openById(convertedDocId);\n    return parseAgendaStructureWithText_(doc.getBody(), '');\n  } finally {\n" +
   "    try { DriveApp.getFileById(convertedDocId).setTrashed(true); } catch (e) { }\n  }\n"],
  ["  // After T-2026.10.19: a Word content type, a checked result, every step logged (AGENDA WORD CONVERSION section).\n" +
   "  return convertAgendaWordBlob_(wordBlob);\n",
   "  const resource = {\n    title: 'TEMP_Agenda_' + new Date().getTime(),\n    mimeType: MimeType.GOOGLE_DOCS\n  };\n\n" +
   "  // The 'insert' method is from an older version of the Drive API.\n  // The modern equivalent is 'create'.\n" +
   "  const file = Drive.Files.create(resource, wordBlob);\n  if (!file || !file.id) throw new Error('Drive conversion did not return a file ID.');\n  return file.id;\n"]
];

/**
 * The agenda readers (after T-2026.10.20, MBS SA4#138): the converted agenda
 * is read by DocumentApp, else the Docs API, else the Word file itself, each
 * through the same Body view and the unchanged parser. One section after the
 * AGENDA WORD CONVERSION section, and two functions changed
 * (parseAgendaWordBlob_ reads through the chain; diagnoseAgendaConversion
 * names the reader). It came after the agenda Word conversion, so
 * codeWithoutAgendaWordConversion() takes it out first. [now, before] pairs;
 * with the section they give the Code.js of T-2026.10.20 byte for byte.
 * tests/agenda-readers.test.js pins the behaviour.
 */
const AGENDA_READERS_SECTION_START = BAR + '// AGENDA READERS (after T-2026.10.20)\n' + BAR;
const AGENDA_READERS_SECTION_END = BAR + '// END OF AGENDA READERS\n' + BAR + '\n';
const AGENDA_READERS_SECTION_FUNCTIONS = ['readAgendaWithFallbacks_', 'readAgendaWith_', 'agendaItemsProblem_', 'readConvertedAgendaWithDocsApi_', 'docsApiAgendaBody_', 'docsApiAgendaBlocks_', 'wordFileAgendaBody_',
  'wordFileAgendaBlocks_', 'agendaWordStyles_', 'parseAgendaXml_', 'agendaXmlPrefixFor_', 'agendaXmlText_', 'agendaBodyView_', 'diagnoseAgendaReaders'];
const AGENDA_READERS_CHANGED_FUNCTIONS = ['parseAgendaWordBlob_', 'diagnoseAgendaConversion'];
const AGENDA_READERS_CHANGES = [
  // The release of it (T-2026.10.21): the version line and the changelog entry of 2.26.5.
  [" * Version: 2.26.5 (2026-11-11)\n",
   " * Version: 2.26.4 (2026-11-11)\n"],
  [" * CHANGELOG\n * 2.26.5 (2026-11-11)\n *   - Fixed: the agenda is read even when DocumentApp cannot open the\n *     Google Doc converted from it. It is read by DocumentApp, else by the\n *     Docs API on the same converted document, else from the Word file\n *     itself; all three feed the unchanged parser. A reader whose agenda\n *     is empty or malformed has failed, and the next one is asked. The log\n *     names the reader (counts only, never agenda text); if none reads it,\n *     the error names each. diagnoseAgendaReaders(), run from the editor,\n *     runs all three on the configured agenda and writes nothing. (MBS\n *     SA4#138, 2026-11-11: DocumentApp.openById failed 4 of 4 times on the\n *     converted agenda S4-261857 with T-2026.10.20.) (AGENDA READERS\n *     section; docs/AGENDA_WORD_CONVERSION.md.)\n",
   " * CHANGELOG\n"],
  ["  // After T-2026.10.19: opened with bounded retries, trashed in every case (AGENDA WORD CONVERSION section).\n" +
   "  // After T-2026.10.20: read by DocumentApp, else the Docs API, else the Word file itself (AGENDA READERS section).\n" +
   "  try {\n    return readAgendaWithFallbacks_(convertedDocId, wordBlob.getBytes());\n",
   "  // After T-2026.10.19: opened with bounded retries, trashed in every case (AGENDA WORD CONVERSION section).\n" +
   "  try {\n    return parseAgendaStructureWithText_(openConvertedAgendaBody_(convertedDocId), '');\n"],
  ["    // After T-2026.10.20: the reader that read it is named (AGENDA READERS section).\n" +
   "    const text = 'Agenda ' + cfg.AGENDA_TDOC + ': converted, opened and parsed -- ' + items.length + ' agenda item(s)' +\n" +
   "      (AGENDA_LAST_READER_ ? ', read by ' + AGENDA_LAST_READER_ : '') + '.';\n",
   "    const text = 'Agenda ' + cfg.AGENDA_TDOC + ': converted, opened and parsed -- ' + items.length + ' agenda item(s).';\n"]
];

const LARGE_BUILD_SECTION_START = BAR + '// LARGE BUILD CHECKPOINTS (after T-2026.10.22)\n' + BAR;
const LARGE_BUILD_SECTION_END = BAR + '// END OF LARGE BUILD CHECKPOINTS\n' + BAR + '\n';
const LARGE_BUILD_SECTION_FUNCTIONS = ['currentReportDocument_', 'beginLargeBuildCheckpoints_', 'endLargeBuildCheckpoints_', 'largeBuildCount_', 'reopenReportDocument_', 'largeBuildCheckpoint_',
  'largeBuildPhaseCheckpoint_', 'largeBuildStep_', 'largeBuildTablesIfDue_', 'planStatusDropdownRepair_', 'repairStatusDropdowns_', 'formatStatusDropdownRepair_', 'previewStatusDropdownRepair', 'repairStatusDropdownsAfterBuild'];
// The released functions it changed beyond the document access below.
const LARGE_BUILD_CHANGED_FUNCTIONS = ['runFullReportBuildCore_', 'buildSkeletonWithTdocTables', 'addAbstractsForTables_', 'removeRowHeightAndSpacing', 'setTwoColumnTDocTableWidths_', 'finalizeStatusDropdowns_'];
// Every code use of DocumentApp.getActiveDocument() outside the section became currentReportDocument_(): this many.
const LARGE_BUILD_DOCUMENT_ACCESS_COUNT = 44;
// After T-2026.10.22 (GitHub issue #6): the changes outside the section, [now, before], in the order they were made.
const LARGE_BUILD_CHANGES = [
  // The release of it (T-2026.10.23): the version line and the changelog entry of 2.26.7.
  [" * Version: 2.26.7 (2026-11-12)\n",
   " * Version: 2.26.6 (2026-11-11)\n"],
  [" * CHANGELOG\n * 2.26.7 (2026-11-12)\n *   - Fixed: a large Full Report Build is saved in parts. When about 2000\n *     DocumentApp changes are pending, the build saves the report and\n *     opens it again (saveAndClose + openById, as Google's message asks)\n *     and goes on with the reopened document; every later access takes\n *     the current document, never an element of the closed one. A small\n *     build saves once, as before. A save that fails fails the build.\n *     previewStatusDropdownRepair() and repairStatusDropdownsAfterBuild(),\n *     run from the editor, finish the Status dropdowns of a built report\n *     whose last phase did not complete: only the Status cells, never a\n *     table twice, a report with doubled TDoc tables refused. (Audio SWG\n *     SA4#138, 2026-11-11: Full Report Build of 111 TDocs ended in \"Too\n *     many changes applied before saving document\".) (LARGE BUILD\n *     CHECKPOINTS section; docs/LARGE_BUILD_CHECKPOINTS.md.)\n",
   " * CHANGELOG\n"],
  ["      entry.detail = outcome.detail || '';\n      // After T-2026.10.22: in a large build, what the phase changed is saved before the next one starts (LARGE BUILD CHECKPOINTS section).\n      if (name !== 'status dropdowns') largeBuildPhaseCheckpoint_('after ' + name);\n",
   "      entry.detail = outcome.detail || '';\n"],
  ["  resetStatusDropdownRun_();\n  // After T-2026.10.22: a large build is saved in parts (LARGE BUILD CHECKPOINTS section).\n  beginLargeBuildCheckpoints_();\n\n  phase('skeleton', function () {\n",
   "  resetStatusDropdownRun_();\n\n  phase('skeleton', function () {\n"],
  ["      // After T-2026.10.22: a save that failed is a failure of the build, not a skipped step (LARGE BUILD CHECKPOINTS section).\n      if (done.saveFailed) throw new Error(done.reason);\n", ''],
  ["  const checkpoints = endLargeBuildCheckpoints_();\n  const totalMs = Date.now() - totalStart;\n", "  const totalMs = Date.now() - totalStart;\n"],
  [".join(', ') + (checkpoints ? ' -- ' + checkpoints + ' checkpoint(s)' : ''));\n", ".join(', '));\n"],
  ["  // After T-2026.10.22: `let` -- after a checkpoint the body is the reopened document's (LARGE BUILD CHECKPOINTS section).\n  let body = currentReportDocument_().getBody().clear();\n",
   "  const body = currentReportDocument_().getBody().clear();\n"],
  ["          body = largeBuildStep_(body, 'skeleton: TDoc tables', LARGE_BUILD_UNITS_.tdocTable);\n", ''],
  ["      largeBuildCount_(allTdocs.length * LARGE_BUILD_UNITS_.registrationRow);\n", ''],
  ["  // After T-2026.10.22: in a Full Build the tables are taken again after a checkpoint (LARGE BUILD CHECKPOINTS section).\n  let tables = getTablesCounted_(body, 'addAbstractsForTables_');\n  tables.slice().forEach((_, tableIndex) => {\n    tables = largeBuildTablesIfDue_(tables, 'abstracts');\n    const table = tables[tableIndex];\n",
   "  getTablesCounted_(body, 'addAbstractsForTables_').forEach(table => {\n"],
  ["    if (asked !== false) largeBuildCount_(LARGE_BUILD_UNITS_.abstractRow);\n", ''],
  ["  // After T-2026.10.22: `let` -- in a Full Build the tables are taken again after a checkpoint (LARGE BUILD CHECKPOINTS section).\n  let tables = getTablesCounted_(body, 'removeRowHeightAndSpacing');\n",
   "  const tables = getTablesCounted_(body, 'removeRowHeightAndSpacing');\n"],
  ["  // After T-2026.10.22: the widths may have made a checkpoint; the tables of the document that is open now (LARGE BUILD CHECKPOINTS section).\n  tables = largeBuildTablesIfDue_(tables, 'formatting', true);\n", ''],
  ["    tables = largeBuildTablesIfDue_(tables, 'formatting');\n    const table = tables[i];\n    largeBuildCount_(table.getNumRows() * LARGE_BUILD_UNITS_.formattedRow);\n",
   "    const table = tables[i];\n"],
  ["  // After T-2026.10.22: in a Full Build the tables are taken again after a checkpoint (LARGE BUILD CHECKPOINTS section).\n  let tables = getTablesCounted_(body, 'setTwoColumnTDocTableWidths_');\n\n  tables.slice().forEach((_, tableIndex) => {\n    tables = largeBuildTablesIfDue_(tables, 'formatting: widths');\n    const t = tables[tableIndex];\n",
   "  const tables = getTablesCounted_(body, 'setTwoColumnTDocTableWidths_');\n\n  tables.forEach(t => {\n"],
  ["    largeBuildCount_(t.getNumRows() * LARGE_BUILD_UNITS_.widthRow);\n", ''],
  ["    // After T-2026.10.22: the caller is told that the save failed (LARGE BUILD CHECKPOINTS section).\n    try {\n      doc.saveAndClose();\n    } catch (e) {\n      result.saveFailed = true;\n      throw new Error('the document could not be saved (' + e.message + ')');\n    }\n",
   "    doc.saveAndClose();\n"]
];

/** Code.js without the large-build checkpoints: the Code.js of T-2026.10.22. { code, problems }. */
function codeWithoutLargeBuildCheckpoints(code) {
  // The meeting time zones (after T-2026.10.23) came later: taken out first.
  const later = codeWithoutMeetingTimeZones(code);
  const view = takeOut(later.code, LARGE_BUILD_SECTION_START, LARGE_BUILD_SECTION_END, LARGE_BUILD_CHANGES, 'LARGE BUILD CHECKPOINTS');
  // The document access: every currentReportDocument_() left is one that was DocumentApp.getActiveDocument().
  const uses = (view.code.match(/currentReportDocument_\(\)/g) || []).length;
  if (uses !== LARGE_BUILD_DOCUMENT_ACCESS_COUNT) view.problems.push('LARGE BUILD CHECKPOINTS: ' + uses + ' uses of currentReportDocument_() outside the section, not ' + LARGE_BUILD_DOCUMENT_ACCESS_COUNT);
  return { code: view.code.split('currentReportDocument_()').join('DocumentApp.getActiveDocument()'), problems: later.problems.concat(view.problems) };
}

const MAIN_MEETING_NUMBER_SECTION_START = BAR + '// MAIN MEETING NUMBER (after T-2026.10.24)\n' + BAR;
const MAIN_MEETING_NUMBER_SECTION_END = BAR + '// END OF MAIN MEETING NUMBER\n' + BAR + '\n';
const MAIN_MEETING_NUMBER_SECTION_FUNCTIONS = ['mainMeetingFolderOfFtpBase_', 'mainMeetingNumberOfFtpBase_', 'mainMeetingIdentityOfSave_', 'mainMeetingNumberMismatch_'];
const MAIN_MEETING_NUMBER_CHANGED_FUNCTIONS = ['getReportConfig_', 'continuousUpdateCore_', 'persistConfigurationSettings_', 'assertMeetingReadyToBuild_'];
// After T-2026.10.24 (GitHub issue #9): the changes outside the section, [now, before], in the order they stand in Code.js.
const MAIN_MEETING_NUMBER_CHANGES = [
  ["  // After T-2026.10.24: a number that is not saved is the document folder's before it is SA4#136's (MAIN MEETING NUMBER section).\n  const MEETING_NUMBER = props.getProperty('MEETING_NUMBER') || mainMeetingNumberOfFtpBase_(props.getProperty('FTP_BASE')) || '136';\n",
   "  const MEETING_NUMBER = props.getProperty('MEETING_NUMBER') || '136';\n"],
  ["      // After T-2026.10.24: the list is fetched before the report is touched; a failed download changed nothing either (MAIN MEETING NUMBER section).\n      result.notUpdated = true;\n",
   "      if (e && e.tdocListIdentity) result.notUpdated = true;\n"],
  ["  const keptMeetingTimeZone = keptMeetingTimeZoneBeforeSave_(docProps);\n  // After T-2026.10.24: a main meeting's folder and number go with its document folder (MAIN MEETING NUMBER section).\n  const mainMeetingIdentity = mainMeetingIdentityOfSave_(config, docProps);\n  if (mainMeetingIdentity.problem) throw new Error(mainMeetingIdentity.problem);\n",
   "  const keptMeetingTimeZone = keptMeetingTimeZoneBeforeSave_(docProps);\n"],
  ["  if (effectiveType === 'main') {\n    // After T-2026.10.24: taken from the document folder where blank; SA4#136's only without one (MAIN MEETING NUMBER section).\n    setPropertyIfPresent_(docProps, 'MEETING_FOLDER', mainMeetingIdentity.folder);\n    setPropertyIfPresent_(docProps, 'MEETING_NUMBER', mainMeetingIdentity.number);\n",
   "  if (effectiveType === 'main') {\n    docProps.setProperty('MEETING_FOLDER', config.meetingFolder || 'TSGS4_136_Montreal');\n    docProps.setProperty('MEETING_NUMBER', config.meetingNumber || '136');\n"],
  ["  assertMeetingConfigured_(context);\n  // After T-2026.10.24: a main meeting whose number is not its document folder's (MAIN MEETING NUMBER section).\n  const meetingNumberProblem = mainMeetingNumberMismatch_(getReportStateStore_(context));\n  if (meetingNumberProblem) throw new Error(meetingNumberProblem);\n  const readiness = getBuildReadiness_(context);\n",
   "  assertMeetingConfigured_(context);\n  const readiness = getBuildReadiness_(context);\n"]
];

// After T-2026.10.28: the scheduled overnight pause of the local listing agent (07:30-19:30 meeting time) is told apart
// from a genuine synchronization failure in the report. The newest layer: taken out first.
const OVERNIGHT_PAUSE_NEW_FUNCTIONS = ['localListingScheduledPause_'];
const OVERNIGHT_PAUSE_CHANGED_FUNCTIONS = ['importAgentListing_'];
const OVERNIGHT_PAUSE_CHANGES = [
  ["  if (!chosen && /^stale/.test(firstReason)) {\n    // After T-2026.10.28: the agent's scheduled overnight pause is not a synchronization failure (localListingScheduledPause_()).\n    let newest = null;\n    try { newest = JSON.parse(read.files[0].raw).fetchedAt; } catch (e) { newest = null; }\n    const pause = localListingScheduledPause_(meetingSessionTimeZone_(context), newest, now);\n    if (pause) return finish('paused', 'Not imported: scheduled overnight pause. The local listing agent scans from 07:30 to 19:30 meeting time; its last scan of ' + group +\n      ' was at ' + pause.lastScan + ', scanning resumes at ' + pause.resumes + '. This is not a synchronization failure. Previously imported revisions are kept.');\n  }\n  if (!chosen) return finish(/^stale/",
   "  if (!chosen) return finish(/^stale/"],
  ["/**\n * After T-2026.10.28 (pure): whether a stale listing of the local listing\n * agent is its scheduled overnight pause rather than a failure. The agent\n * scans only from 07:30 to 19:30 meeting local time (docs/LOCAL_LISTING_AGENT.md).\n * It is the pause when `now` is outside that window, or in its first 15\n * minutes (the first scan of the day is still on its way), AND the newest\n * listing was scanned at most 30 minutes before the window closed: the agent\n * worked until the end. A listing that stopped earlier, and any staleness\n * during the day, stay failures. `zone` is the meeting's IANA zone\n * (meetingSessionTimeZone_()); without one nothing is taken for the pause.\n * Returns null, or { lastScan, resumes } as wall-clock texts of the zone\n * (\"2026-11-17 19:28 ABT\", \"07:30 ABT\").\n */\nfunction localListingScheduledPause_(zone, fetchedAtIso, now) {\n  const START = '07:30', END = '19:30', GRACE_MINUTES = 15, LAST_SCAN_MINUTES = 30;\n  const t = Date.parse(fetchedAtIso);\n  if (!zone || isNaN(t) || !now || isNaN(now.getTime())) return null;\n  const wall = meetingZoneWallClockOfInstant_(zone, now.getTime());\n  if (!wall) return null;\n  const minutes = function (hhmm) { return parseInt(hhmm.slice(0, 2), 10) * 60 + parseInt(hhmm.slice(3, 5), 10); };\n  const m = minutes(wall.slice(11, 16));\n  let closedOn;\n  if (m >= minutes(END)) closedOn = wall.slice(0, 10);\n  else if (m < minutes(START) + GRACE_MINUTES) closedOn = new Date(Date.parse(wall.slice(0, 10) + 'T00:00:00Z') - 86400000).toISOString().slice(0, 10);\n  else return null;\n  const closedAt = meetingWallClockToInstant_(zone, closedOn, END);\n  if (closedAt === null || t < closedAt - LAST_SCAN_MINUTES * 60000) return null;\n  const scan = meetingZoneWallClockOfInstant_(zone, t);\n  return { lastScan: scan.slice(0, 10) + ' ' + scan.slice(11, 16) + ' ' + meetingTimeZoneAbbreviationAt_(zone, t), resumes: START + ' ' + meetingTimeZoneAbbreviationAt_(zone, now.getTime()) };\n}\n\n/** importAgentListing_() for an update: never throws",
   "/** importAgentListing_() for an update: never throws"]
];

/** Code.js without the overnight-pause distinction after T-2026.10.28: { code, problems }. */
function codeWithoutOvernightPause(code) {
  return takeOut(code, null, null, OVERNIGHT_PAUSE_CHANGES, 'OVERNIGHT PAUSE');
}

// After T-2026.10.28: the Revisions cell without the "Revisions folders" line; "No revisions available." only where
// it is known; Check Connections without the manual-import instruction while the agent delivers. The newest layer.
const CELL_CLEANUP_NEW_FUNCTIONS = ['revisionNoneMode_', 'revisedIntoSet_'];
const CELL_CLEANUP_CHANGED_FUNCTIONS = ['renderRevisionsCellContent_', 'updateRevisions_', 'renderStoredLocalRevisions_', 'describeLocalRevisionsStatus_', 'testAllConnections', 'testRevisionsFolder', 'importAgentListing_'];
const CELL_CLEANUP_CHANGES = [
  ["    .filter(e => e.line);\n  // After T-2026.10.28: no \"Revisions folders: ...\" line any more (the report links the folders elsewhere);\n  // one written by an earlier release goes at this cell's next write.\n  const folderLine = null;\n\n",
   "    .filter(e => e.line);\n  // After T-2026.10.25: a last line linking both revisions folders, when the\n  // report has a local (meeting-room) folder. Without one, nothing changes.\n  const folderLine = buildRevisionFoldersLine_(folders);\n\n"],
  ["  if (!entries.length) {\n    // After T-2026.10.28: \"No revisions available.\" only where it is known (options.none, revisionNoneMode_()):\n    // 'write' -- every source was read and has none; 'remove' -- the TDoc has revisions as TDocs (REVISION_MAP), so\n    // the sentence would be wrong; 'keep' -- a source could not be read (the local listing is stale or missing):\n    // nothing is claimed, an earlier text stays.\n    const mode = (options && options.none) || 'write';\n    const current = cell.getText();\n",
   "  if (!entries.length) {\n    const current = cell.getText();\n"],
  ["    const current = cell.getText();\n    const lines = current.split('\\n');\n    const generated = function (line) { return isLocalRevisionCellLine_(line) || line.indexOf(REVISION_FOLDERS_PREFIX_) === 0 || line.trim() === REVISIONS_NONE_TEXT_; };\n    const typed = lines.filter(function (line) { return line.trim() && !generated(line); }).join('\\n').trim();\n    // Text typed by hand is left alone; only lines this code wrote go (a local revision no longer listed, a folders line).\n    if (typed) {\n      const kept = lines.filter(function (line) { return !isLocalRevisionCellLine_(line) && line.indexOf(REVISION_FOLDERS_PREFIX_) !== 0; }).join('\\n').trim();\n      if (kept !== current.trim()) cell.setText(kept);\n      return;\n",
   "    const current = cell.getText();\n    if (!folderLine && current.indexOf(REVISION_FOLDERS_PREFIX_) === -1) {\n      if (!current.trim()) cell.setText(REVISIONS_NONE_TEXT_);\n      return;\n"],
  ["    }\n    const hadNone = lines.some(function (line) { return line.trim() === REVISIONS_NONE_TEXT_; });\n    let wanted;\n    if (mode === 'write') wanted = REVISIONS_NONE_TEXT_;\n    else if (mode === 'remove') wanted = '';\n    else wanted = hadNone ? REVISIONS_NONE_TEXT_ : '';\n    if (current === wanted || (!current.trim() && !wanted)) return;\n    cell.setText(wanted);\n",
   "    }\n    // Issue #16: a local revision no longer listed leaves the cell; whatever else is in it stays.\n    const withoutLocal = current.split('\\n').filter(line => !isLocalRevisionCellLine_(line)).join('\\n');\n    if (withoutLocal !== current) {\n      const rest = stripRevisionFoldersLine_(withoutLocal).trim();\n      const keep = rest && rest !== REVISIONS_NONE_TEXT_ ? rest : REVISIONS_NONE_TEXT_;\n      const text = folderLine ? keep + '\\n' + folderLine.text : keep;\n      cell.setText(text);\n      if (folderLine) linkRevisionFoldersLine_(cell.editAsText(), keep.length + 1, folderLine);\n      return;\n    }\n    // Text that is not ours (typed by hand) is left alone, as before.\n    const own = stripRevisionFoldersLine_(current).trim();\n    if (own && own !== REVISIONS_NONE_TEXT_) return;\n    const wanted = folderLine ? REVISIONS_NONE_TEXT_ + '\\n' + folderLine.text : REVISIONS_NONE_TEXT_;\n    if (current === wanted) return;\n    cell.setText(wanted);\n"],
  ["    cell.setText(wanted);\n    return;\n",
   "    cell.setText(wanted);\n    if (folderLine) linkRevisionFoldersLine_(cell.editAsText(), REVISIONS_NONE_TEXT_.length + 1, folderLine);\n    return;\n"],
  ["  manual.forEach(line => { offset += line.length + 1; });\n}\n\n/**\n * After T-2026.10.28 (pure): whether a Revisions cell without revision lines\n * may say \"No revisions available.\" -- 'remove' when the TDoc has revisions\n * as TDocs (`revisionMap`, REVISION_MAP: revised to another, or another\n * revised to it); 'keep' when a source could not be read (`sourcesRead`\n * false); otherwise 'write'.\n */\nfunction revisionNoneMode_(tdoc, revisionMap, revisedInto, sourcesRead) {\n  if ((revisionMap && revisionMap[tdoc]) || (revisedInto && revisedInto[tdoc])) return 'remove';\n  return sourcesRead ? 'write' : 'keep';\n}\n\n/** Pure: the TDocs some other TDoc was revised into (the values of REVISION_MAP), as { tdoc: true }. */\nfunction revisedIntoSet_(revisionMap) {\n  const set = {};\n  Object.keys(revisionMap || {}).forEach(function (k) { const v = String(revisionMap[k] || '').trim(); if (v) set[v] = true; });\n  return set;\n}\n",
   "  manual.forEach(line => { offset += line.length + 1; });\n  if (folderLine) linkRevisionFoldersLine_(te, offset, folderLine);\n}\n"],
  ["    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n    const map = loadJsonObject_(props.getProperty('REVISION_MAP'));\n    renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders,\n      { force: true, none: revisionNoneMode_(tdoc, map, revisedIntoSet_(map), true) });\n  });\n",
   "    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n    renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders, { force: true });\n  });\n"],
  ["  const parsed = parseLocalRevisionsListing_({ folder: chosen.folder, html: chosen.payload, text: '' }, localBase);\n  // After T-2026.10.28: a complete, fresh scan of the group's folder with no file in it is an answer -- the folder is\n  // empty -- not a failure. Nothing is stored or removed (storeLocalRevisions_() never removes on an empty listing).\n  if (!parsed.ok && !parsed.anchors.length && !parsed.truncated) {\n    return finish('imported', 'Imported the scan of ' + chosen.fetchedAt + ': the folder of ' + group + ' has no files.',\n      { lastScanId: chosen.scanId, lastImportAt: now.toISOString(), files: 0, recognized: 0, added: 0, removed: 0 });\n  }\n  if (!parsed.ok) return finish('invalid', 'Not imported: ' + parsed.error + ' Previously imported revisions are kept.');\n",
   "  const parsed = parseLocalRevisionsListing_({ folder: chosen.folder, html: chosen.payload, text: '' }, localBase);\n  if (!parsed.ok) return finish('invalid', 'Not imported: ' + parsed.error + ' Previously imported revisions are kept.');\n"],
  ["/** Pure: what the report says about the meeting-room server: not checked yet, or when a listing was last imported. */\nfunction describeLocalRevisionsStatus_(lastImportJson, localBase, agentStateJson) {\n  if (!localBase) return 'Local (meeting-room server): none configured.';\n",
   "/** Pure: what the report says about the meeting-room server: not checked yet, or when a listing was last imported. */\nfunction describeLocalRevisionsStatus_(lastImportJson, localBase) {\n  if (!localBase) return 'Local (meeting-room server): none configured.';\n"],
  ["  const last = loadJsonObject_(lastImportJson);\n  if (!last.at) {\n    // After T-2026.10.28: with a valid scan of the local listing agent there is nothing to do by hand.\n    const agent = loadJsonObject_(agentStateJson);\n    if (agent.lastScanFetchedAt && (agent.lastResult === 'imported' || agent.lastResult === 'unchanged')) return '';\n    return 'Local (meeting-room server): not checked yet. Import Local Inbox Listing… on the meeting-room network.';\n  }\n  return 'Local (meeting-room server): last listing imported ' + last.at + ' from ' + last.folder + ': ' + last.listed + ' file(s), ' +\n",
   "  const last = loadJsonObject_(lastImportJson);\n  if (!last.at) return 'Local (meeting-room server): not checked yet. Import Local Inbox Listing… on the meeting-room network.';\n  return 'Local (meeting-room server): last listing imported ' + last.at + ' from ' + last.folder + ': ' + last.listed + ' file(s), ' +\n"],
  ["  // Never throws, and without a configured folder it does nothing at all.\n  const agent = importAgentListingSafely_(context);\n  // After T-2026.10.28: with an agent folder, the local source counts as read only when its listing was valid now.\n  const localRead = !agent || agent.result === 'imported' || agent.result === 'unchanged';\n  const baseUrl = String(getMeetingContext_(context).sources.revisionsUrl || '').trim();\n",
   "  // Never throws, and without a configured folder it does nothing at all.\n  importAgentListingSafely_(context);\n  const baseUrl = String(getMeetingContext_(context).sources.revisionsUrl || '').trim();\n"],
  ["  perfCount_('revision anchors parsed', anchors.length);\n  // After T-2026.10.28: TDocs with revisions as TDocs never say \"No revisions available.\" (revisionNoneMode_()).\n  const revisionMap = loadJsonObject_(props.getProperty('REVISION_MAP'));\n  const revisedInto = revisedIntoSet_(revisionMap);\n  // After T-2026.10.25: both folders, for the line under the revisions (the local one is never fetched).\n",
   "  perfCount_('revision anchors parsed', anchors.length);\n  // After T-2026.10.25: both folders, for the line under the revisions (the local one is never fetched).\n"],
  ["    // from here. Only the 3GPP ones below get a table of their own, as before.\n    renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders,\n      { none: revisionNoneMode_(tdoc, revisionMap, revisedInto, localRead) });\n    renderingMs += Date.now() - renderStart;\n",
   "    // from here. Only the 3GPP ones below get a table of their own, as before.\n    renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders);\n    renderingMs += Date.now() - renderStart;\n"],
  ["  // Issue #16: when a listing of it was last imported.\n  const localRevisionsStatus = localRevisions ? describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions, getReportStateStore_().getProperty(LOCAL_LISTING_STATE_KEY_)) : '';\n  if (localRevisionsStatus) results.push('   ' + localRevisionsStatus);\n  // After T-2026.10.27: the local listing agent (LOCAL LISTING AGENT section).\n",
   "  // Issue #16: when a listing of it was last imported.\n  if (localRevisions) results.push('   ' + describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions));\n  // After T-2026.10.27: the local listing agent (LOCAL LISTING AGENT section).\n"],
  ["        // Issue #16: when a listing of it was last imported.\n        '\\n' + describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions, getReportStateStore_().getProperty(LOCAL_LISTING_STATE_KEY_)) : ''),\n      ui.ButtonSet.OK\n",
   "        // Issue #16: when a listing of it was last imported.\n        '\\n' + describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions) : ''),\n      ui.ButtonSet.OK\n"]
];

/** Code.js without the Revisions-cell cleanup after T-2026.10.28: { code, problems }. */
function codeWithoutRevisionsCellCleanup(code) {
  return takeOut(code, null, null, CELL_CLEANUP_CHANGES, 'REVISIONS CELL CLEANUP');
}

// After T-2026.10.27: the local listing agent -- every update imports the agent's Drive listing of the report's group.
// The newest layer: taken out first.
const LOCAL_LISTING_AGENT_NEW_FUNCTIONS = ['tdocTablesByNumber_', 'renderStoredLocalRevisions_', 'localListingFolderIdOf_', 'localListingSha256Hex_', 'validateAgentListing_', 'readAgentListingFiles_',
  'importAgentListing_', 'importAgentListingSafely_', 'describeLocalListingAgentStatus_', 'tdocListConnectionLines_'];
const LOCAL_LISTING_AGENT_CHANGED_FUNCTIONS = ['saveLocalRevisionsListing', 'updateRevisions_', 'configureMeetingSettings', 'persistConfigurationSettings_', 'testAllConnections'];
const LOCAL_LISTING_AGENT_CHANGES = [
  ["  'LOCAL_REVISIONS_URL',\n  // After T-2026.10.27: the Drive folder of the local listing agent (LOCAL LISTING AGENT).\n  'LOCAL_LISTING_DRIVE_FOLDER_ID',\n  'FETCH_ABSTRACTS_ON_UPDATE',\n",
   "  'LOCAL_REVISIONS_URL',\n  'FETCH_ABSTRACTS_ON_UPDATE',\n"],
  ["\n/** The TDoc tables of a report body: { <TDoc number>: table }, the first table of each number. */\nfunction tdocTablesByNumber_(body) {\n  const byTdoc = {};\n  body.getTables().filter(isTDocTable_).forEach(function (t) {\n    const id = parseExactSA4DocumentId_(String(safeCellText_(t, 0, 1) || '').trim());\n    if (id.isValid && !byTdoc[id.raw]) byTdoc[id.raw] = t;\n  });\n  return byTdoc;\n}\n\n/** Writes the Revisions cells of `touched` TDocs from what is stored for both folders (nothing is fetched). */\nfunction renderStoredLocalRevisions_(props, byTdoc, touched, localBase, context) {\n  if (!touched.length) return;\n  const remoteBase = String(getMeetingContext_(context).sources.revisionsUrl || '').trim();\n  const folders = { local: localBase, remote: remoteBase ? remoteBase.replace(/\\/$/, '') + '/' : '' };\n  touched.forEach(function (tdoc) {\n    const t = byTdoc[tdoc];\n    if (!t) return;\n    const ordered = orderedStoredRevisions_(loadJsonObject_(props.getProperty('REVIS_' + tdoc)));\n    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n    renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders, { force: true });\n  });\n}\n\n// ---------------------------------------------------------------------------\n// LOCAL LISTING AGENT (after T-2026.10.27)\n// The meeting-room server cannot be read from Google (docs/LOCAL_REVISIONS_\n// AUTODETECT.md). One Windows laptop on the meeting-room network runs the\n// local listing agent (tools/local-listing-agent/): every two minutes it reads\n// the folder listing of each working group and writes <group>.listing.json\n// into a Drive folder that Google Drive for desktop synchronizes. Every update\n// of a report (updateRevisions_()) reads its own group's file from that folder\n// (LOCAL_LISTING_DRIVE_FOLDER_ID) and imports it with the code of Import Local\n// Inbox Listing... -- the same parser, the same per-folder refresh, the same\n// merge and rendering. The manual import stays as it is.\n//\n// The file (written by the agent, atomically: temp file + rename):\n//   { format: 'sa4-local-listing', version: 1, group, folder, fetchedAt (ISO\n//     UTC), scanId, agent, linkCount, sha256 (hex, of `payload` in UTF-8),\n//     payload: '<a href=\"...\">name</a>\\n...' -- the listing's links only }\n// A file is imported only when it is complete (parses, payload matches its\n// sha256 and linkCount), is this report's group, lists a folder of this\n// group on the report's meeting-room server, and was fetched at most 30\n// minutes ago (and not in the future). Anything else -- no file, a stale one\n// (agent stopped, network down, Drive not yet synchronized: a local write is\n// no proof of an upload), an incomplete or foreign one -- imports nothing\n// and erases nothing; it is recorded in LOCAL_LISTING_AGENT_STATE and shown\n// in Check Connections. The same scan twice is imported once (scanId).\n// ---------------------------------------------------------------------------\nconst LOCAL_LISTING_FOLDER_KEY_ = 'LOCAL_LISTING_DRIVE_FOLDER_ID';\nconst LOCAL_LISTING_STATE_KEY_ = 'LOCAL_LISTING_AGENT_STATE';\nconst LOCAL_LISTING_FORMAT_ = 'sa4-local-listing';\nconst LOCAL_LISTING_MAX_AGE_MINUTES_ = 30;\nconst LOCAL_LISTING_MAX_FUTURE_MINUTES_ = 5;\n\n/** Pure: a Drive folder id from an id or a folder URL; '' when it is neither. */\nfunction localListingFolderIdOf_(value) {\n  const v = String(value || '').trim();\n  const fromUrl = /^https:\\/\\/drive\\.google\\.com\\/(?:drive\\/(?:u\\/\\d+\\/)?folders\\/|open\\?id=)([A-Za-z0-9_-]{10,})/.exec(v);\n  if (fromUrl) return fromUrl[1];\n  return /^[A-Za-z0-9_-]{10,}$/.test(v) ? v : '';\n}\n\n/** The SHA-256 of a text's UTF-8 bytes, in hex. */\nfunction localListingSha256Hex_(text) {\n  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);\n  return bytes.map(function (b) { const v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');\n}\n\n/**\n * Pure: whether `raw` (the text of a <group>.listing.json) may be imported\n * into a report of group `expect.group` whose meeting-room server is\n * `expect.localBase`, at `now` (Date). `sha256` is the digest function.\n * Returns { ok, reason, listing: { group, folder, fetchedAt, scanId, payload } }.\n */\nfunction validateAgentListing_(raw, expect, now, sha256) {\n  const no = function (reason) { return { ok: false, reason: reason, listing: null }; };\n  let f;\n  try { f = JSON.parse(String(raw || '')); } catch (e) { return no('incomplete or unreadable (not complete JSON: still being written or synchronized?)'); }\n  if (!f || f.format !== LOCAL_LISTING_FORMAT_ || f.version !== 1) return no('not a listing written by the local listing agent');\n  if (String(f.group || '') !== String(expect.group || '')) return no('a listing of group ' + f.group + ', not of ' + expect.group);\n  const origin = (String(expect.localBase || '').match(/^(https?:\\/\\/[^\\/?#]+)/i) || [])[1];\n  const folder = String(f.folder || '');\n  if (!origin || folder.toLowerCase().indexOf(origin.toLowerCase() + '/') !== 0) return no('the listed folder ' + folder + ' is not on the meeting-room server of this report');\n  if (!new RegExp('/' + String(expect.group).replace(/[.*+?^${}()|[\\]\\\\]/g, '\\\\$&') + '/?$', 'i').test(folder)) return no('the listed folder ' + folder + ' is not the folder of group ' + expect.group);\n  const t = Date.parse(f.fetchedAt);\n  if (isNaN(t)) return no('no valid scan time');\n  if (t > now.getTime() + LOCAL_LISTING_MAX_FUTURE_MINUTES_ * 60000) return no('scanned in the future (' + f.fetchedAt + '): the agent laptop\\'s clock is wrong');\n  if (now.getTime() - t > LOCAL_LISTING_MAX_AGE_MINUTES_ * 60000) return no('stale: scanned ' + f.fetchedAt + ', more than ' + LOCAL_LISTING_MAX_AGE_MINUTES_ + ' minutes ago (agent stopped, meeting-room network down, or Drive not synchronized)');\n  const payload = String(f.payload || '');\n  if (typeof f.sha256 !== 'string' || sha256(payload) !== f.sha256.toLowerCase()) return no('incomplete: its content does not match its checksum');\n  if ((payload.match(/<a\\s/gi) || []).length !== Number(f.linkCount)) return no('incomplete: ' + f.linkCount + ' links announced, ' + (payload.match(/<a\\s/gi) || []).length + ' present');\n  return { ok: true, reason: '', listing: { group: f.group, folder: folder, fetchedAt: f.fetchedAt, scanId: String(f.scanId || f.sha256), payload: payload } };\n}\n\n/** The candidate files of a group in the agent's folder, newest first: [{ raw, updated }]; also the agent's status file. */\nfunction readAgentListingFiles_(folderId, group) {\n  const folder = DriveApp.getFolderById(folderId);\n  const list = [];\n  const files = folder.getFilesByName(group + '.listing.json');\n  while (files.hasNext()) { const file = files.next(); list.push({ raw: file.getBlob().getDataAsString('UTF-8'), updated: file.getLastUpdated() }); }\n  list.sort(function (a, b) { return b.updated.getTime() - a.updated.getTime(); });\n  let status = null;\n  const st = folder.getFilesByName(group + '.status.json');\n  if (st.hasNext()) { try { status = JSON.parse(st.next().getBlob().getDataAsString('UTF-8')); } catch (e) { status = null; } }\n  return { files: list, status: status };\n}\n\n/**\n * Imports the agent's listing of this report's group, if it may be. Returns\n * null when no folder is configured; otherwise { result, touched, state }.\n * Imports nothing and erases nothing unless a complete, fresh listing of\n * this group and server is found; the state says why.\n */\nfunction importAgentListing_(context, deps) {\n  const props = getReportStateStore_(context);\n  const folderId = localListingFolderIdOf_(props.getProperty(LOCAL_LISTING_FOLDER_KEY_));\n  if (!folderId) return null;\n  const now = (deps && deps.now) || new Date();\n  const previous = loadJsonObject_(props.getProperty(LOCAL_LISTING_STATE_KEY_));\n  const state = Object.assign({}, previous, { lastCheckAt: now.toISOString() });\n  const finish = function (result, message, extra) {\n    Object.assign(state, extra || {}, { lastResult: result, message: message });\n    props.setProperty(LOCAL_LISTING_STATE_KEY_, JSON.stringify(state));\n    return { result: result, touched: (extra && extra.touchedList) || [], state: state };\n  };\n  const group = getReportConfig_(context).DRAFTS_FOLDER;\n  const localBase = localRevisionsUrlFor_(context);\n  state.group = group;\n  if (!localBase) return finish('off', 'This report has no Local Revisions URL; the agent\\'s listings are not imported.');\n  let read;\n  try { read = readAgentListingFiles_(folderId, group); } catch (e) { return finish('unavailable', 'The Drive folder of the local listing agent cannot be read (' + e.message + '). Previously imported revisions are kept.'); }\n  if (read.status) {\n    state.agent = { lastSuccessAt: read.status.lastSuccessAt || null, lastAttemptAt: read.status.fetchedAt || null, lastError: read.status.ok ? null : (read.status.error || 'failed'),\n      lastWrittenAt: read.status.writtenAt || null };\n  }\n  if (!read.files.length) return finish('missing', 'Not checked: no listing of ' + group + ' in the agent\\'s Drive folder yet. Previously imported revisions are kept.');\n  let chosen = null;\n  let firstReason = '';\n  read.files.forEach(function (file) {\n    if (chosen) return;\n    const v = validateAgentListing_(file.raw, { group: group, localBase: localBase }, now, localListingSha256Hex_);\n    if (v.ok) { chosen = v.listing; chosen.driveUpdatedAt = file.updated ? file.updated.toISOString() : null; } else if (!firstReason) firstReason = v.reason;\n  });\n  if (!chosen) return finish(/^stale/.test(firstReason) ? 'stale' : 'invalid', 'Not imported: the newest listing of ' + group + ' is ' + firstReason + '. Previously imported revisions are kept.');\n  state.lastScanFetchedAt = chosen.fetchedAt;\n  state.lastScanDriveUpdatedAt = chosen.driveUpdatedAt;\n  if (previous.lastScanId === chosen.scanId) return finish('unchanged', 'Scan of ' + chosen.fetchedAt + ' already imported.');\n  const parsed = parseLocalRevisionsListing_({ folder: chosen.folder, html: chosen.payload, text: '' }, localBase);\n  if (!parsed.ok) return finish('invalid', 'Not imported: ' + parsed.error + ' Previously imported revisions are kept.');\n  const byTdoc = (deps && deps.byTdoc) || tdocTablesByNumber_(getActiveDocumentBodyCounted_('importAgentListing_', context));\n  const stored = storeLocalRevisions_(props, parsed, Object.keys(byTdoc));\n  renderStoredLocalRevisions_(props, byTdoc, stored.touched, localBase, context);\n  return finish('imported', 'Imported the scan of ' + chosen.fetchedAt + ': ' + parsed.listed + ' file(s), ' + stored.recognized + ' revision(s) of SA4 TDocs, ' +\n    stored.added + ' new, ' + stored.removed + ' no longer listed.',\n    { lastScanId: chosen.scanId, lastImportAt: now.toISOString(), files: parsed.listed, recognized: stored.recognized, added: stored.added, removed: stored.removed, touchedList: stored.touched });\n}\n\n/** importAgentListing_() for an update: never throws (a failure is logged and the update goes on). */\nfunction importAgentListingSafely_(context) {\n  try {\n    const r = importAgentListing_(context);\n    if (r) Logger.log('Local listing agent: ' + r.result + ' -- ' + r.state.message);\n    return r;\n  } catch (e) {\n    Logger.log('Local listing agent: not imported (' + e.message + ')');\n    return null;\n  }\n}\n\n/** Pure: the status lines of the local listing agent for this report. */\nfunction describeLocalListingAgentStatus_(stateJson, folderId, now) {\n  if (!folderId) return 'Local listing agent: not configured (Configure Meeting > Local listing folder).';\n  const s = loadJsonObject_(stateJson);\n  if (!s.lastCheckAt) return 'Local listing agent: configured; not checked yet (the next update reads it).';\n  const ago = function (iso) { const t = Date.parse(iso); if (isNaN(t)) return '?'; const m = Math.round((now.getTime() - t) / 60000); return iso.replace('T', ' ').replace(/:\\d\\d(\\.\\d+)?Z$/, ' UTC') + ' (' + m + ' min ago)'; };\n  // Four stages, each from its own evidence: the agent's status file says what the laptop did (as far as Drive has\n  // synchronized it -- a local write is no proof of an upload); the listing found in Drive proves the upload; the\n  // report's own state proves the import.\n  const lines = ['Local listing agent (' + (s.group || '?') + '): last check ' + ago(s.lastCheckAt) + ' -- ' + (s.message || s.lastResult || '')];\n  lines.push('  1. LAN scan (agent status, as synchronized): ' + (s.agent ? 'last success ' + (s.agent.lastSuccessAt ? ago(s.agent.lastSuccessAt) : 'never') + (s.agent.lastError ? ', last attempt failed: ' + s.agent.lastError : '') : 'no status file in Drive'));\n  lines.push('  2. local write (agent status, as synchronized): ' + (s.agent && s.agent.lastWrittenAt ? ago(s.agent.lastWrittenAt) : 'not reported'));\n  lines.push('  3. in Google Drive: ' + (s.lastScanFetchedAt ? 'newest valid scan ' + ago(s.lastScanFetchedAt) + (s.lastScanDriveUpdatedAt ? ', uploaded ' + ago(s.lastScanDriveUpdatedAt) : '') : 'no valid scan found'));\n  lines.push('  4. import into this report: ' + (s.lastImportAt ? ago(s.lastImportAt) + ', ' + (s.files || 0) + ' file(s), ' + (s.recognized || 0) + ' revision(s), ' + (s.added || 0) + ' new, ' + (s.removed || 0) + ' removed' : 'none yet'));\n  return lines.join('\\n');\n}\n\n/** Pure: what the report says about the meeting-room server: not checked yet, or when a listing was last imported. */\n",
   "\n/** Pure: what the report says about the meeting-room server: not checked yet, or when a listing was last imported. */\n"],
  ["    const props = getReportStateStore_();\n    const byTdoc = tdocTablesByNumber_(currentReportDocument_().getBody());\n    const stored = storeLocalRevisions_(props, parsed, Object.keys(byTdoc));\n",
   "    const props = getReportStateStore_();\n    const body = currentReportDocument_().getBody();\n    const tables = body.getTables().filter(isTDocTable_);\n    const byTdoc = {};\n    tables.forEach(function (t) {\n      const id = parseExactSA4DocumentId_(String(safeCellText_(t, 0, 1) || '').trim());\n      if (id.isValid && !byTdoc[id.raw]) byTdoc[id.raw] = t;\n    });\n    const stored = storeLocalRevisions_(props, parsed, Object.keys(byTdoc));\n"],
  ["    const stored = storeLocalRevisions_(props, parsed, Object.keys(byTdoc));\n    renderStoredLocalRevisions_(props, byTdoc, stored.touched, localBase);\n    const at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');\n",
   "    const stored = storeLocalRevisions_(props, parsed, Object.keys(byTdoc));\n    const remoteBase = String(getMeetingContext_().sources.revisionsUrl || '').trim();\n    const folders = { local: localBase, remote: remoteBase ? remoteBase.replace(/\\/$/, '') + '/' : '' };\n    stored.touched.forEach(function (tdoc) {\n      const t = byTdoc[tdoc];\n      const ordered = orderedStoredRevisions_(loadJsonObject_(props.getProperty('REVIS_' + tdoc)));\n      const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n      renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders, { force: true });\n    });\n    const at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');\n"],
  ["function updateRevisions_(cfg, context) {\n  // After T-2026.10.27: the listing the local listing agent wrote to Drive, first (LOCAL LISTING AGENT section).\n  // Never throws, and without a configured folder it does nothing at all.\n  importAgentListingSafely_(context);\n  const baseUrl = String(getMeetingContext_(context).sources.revisionsUrl || '').trim();\n",
   "function updateRevisions_(cfg, context) {\n  const baseUrl = String(getMeetingContext_(context).sources.revisionsUrl || '').trim();\n"],
  ["      <div class=\"hint\">On-site only and possibly newer than the 3GPP folder above, which the updates keep reading. The report only links to it, never reads it. Blank: ${esc(LOCAL_REVISIONS_URL_DEFAULT_)} for a main meeting that is not an e-meeting; \"none\": no local link.</div>\n\n      <label>Local listing folder (Google Drive, written by the local listing agent):</label>\n      <input type=\"text\" id=\"localListingFolder\" value=\"${esc(props.getProperty(LOCAL_LISTING_FOLDER_KEY_) || '')}\" placeholder=\"Drive folder ID or https://drive.google.com/drive/folders/...\">\n      <div class=\"hint\">Optional. Every update then imports this report's group listing from that folder (at most 30 minutes old), like Import Local Inbox Listing... Blank: no automatic import.</div>\n    </div>\n",
   "      <div class=\"hint\">On-site only and possibly newer than the 3GPP folder above, which the updates keep reading. The report only links to it, never reads it. Blank: ${esc(LOCAL_REVISIONS_URL_DEFAULT_)} for a main meeting that is not an e-meeting; \"none\": no local link.</div>\n    </div>\n"],
  ["          localRevisionsUrl: document.getElementById('localRevisionsUrl') ? document.getElementById('localRevisionsUrl').value : undefined,\n          localListingFolder: document.getElementById('localListingFolder') ? document.getElementById('localListingFolder').value : undefined,\n          reportType: document.getElementById('reportType').value,\n",
   "          localRevisionsUrl: document.getElementById('localRevisionsUrl') ? document.getElementById('localRevisionsUrl').value : undefined,\n          reportType: document.getElementById('reportType').value,\n"],
  ["  'LOCAL_REVISIONS_URL',\n  // After T-2026.10.27: the Drive folder of the local listing agent (LOCAL LISTING AGENT).\n  'LOCAL_LISTING_DRIVE_FOLDER_ID',\n  // After T-2026.10.23: the meeting time zone a save stores -- synced into an adopted document's central state, which\n",
   "  'LOCAL_REVISIONS_URL',\n  // After T-2026.10.23: the meeting time zone a save stores -- synced into an adopted document's central state, which\n"],
  ["    else Logger.log('Configure Meeting: the Local Revisions URL was not stored (not an http or https address): ' + localRevisions);\n  }\n  // After T-2026.10.27 (LOCAL LISTING AGENT): only a dialog that has the field changes it. Blank removes it; an id or a folder URL is stored as the id.\n  if (config.localListingFolder !== undefined && config.localListingFolder !== null) {\n    const listingFolder = String(config.localListingFolder).trim();\n    if (!listingFolder) docProps.deleteProperty(LOCAL_LISTING_FOLDER_KEY_);\n    else if (localListingFolderIdOf_(listingFolder)) docProps.setProperty(LOCAL_LISTING_FOLDER_KEY_, localListingFolderIdOf_(listingFolder));\n    else Logger.log('Configure Meeting: the Local listing folder was not stored (not a Drive folder id or URL): ' + listingFolder);\n  }\n",
   "    else Logger.log('Configure Meeting: the Local Revisions URL was not stored (not an http or https address): ' + localRevisions);\n  }\n"],
  [" */\n/**\n * After T-2026.10.27: what Check Connections says about the TDoc list -- the\n * same addresses and checks as downloadTdocListXlsx_(), one line per address\n * tried. A refusal (HTTP 403/429/503) from 3GPP's front end is usually\n * temporary: it is said so, and that the update tries again on its next run.\n */\nfunction tdocListConnectionLines_(url) {\n  const lines = [];\n  const candidates = tdocListDownloadCandidates_(url);\n  for (let i = 0; i < candidates.length; i++) {\n    const attempt = tdocListDownloadAttempt_(candidates[i]);\n    const label = i === 0 ? '' : '(alternative) ';\n    if (attempt.ok) {\n      lines.push('✅ ' + label + 'Connected successfully: ' + attempt.url);\n      lines.push('📊 File size: ' + attempt.bytes + ' bytes (Excel workbook)');\n      return lines;\n    }\n    const refused = /^HTTP (403|429|503)\\b/.test(attempt.problem);\n    lines.push('❌ ' + label + attempt.problem + ': ' + attempt.url + (refused ? ' -- refused by the 3GPP server (usually temporary; the update tries again on its next run)' : ''));\n  }\n  return lines;\n}\n\nfunction testAllConnections() {\n",
   " */\nfunction testAllConnections() {\n"],
  ["    } else {\n      // After T-2026.10.27: checked as the update downloads it -- the configured address, then its\n      // www.3gpp.org/ftp twin (TDOC LIST DOWNLOAD section) -- so this says what an update will get.\n      tdocListConnectionLines_(url).forEach(function (line) { results.push('   ' + line); });\n    }\n",
   "    } else {\n      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });\n      if (response.getResponseCode() === 200) {\n        results.push('   ✅ Connected successfully');\n        results.push('   📊 File size: ' + response.getBlob().getBytes().length + ' bytes');\n      } else {\n        results.push('   ❌ HTTP ' + response.getResponseCode());\n      }\n    }\n"],
  ["  if (localRevisions) results.push('   ' + describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions));\n  // After T-2026.10.27: the local listing agent (LOCAL LISTING AGENT section).\n  if (localRevisions) results.push('   ' + describeLocalListingAgentStatus_(getReportStateStore_().getProperty(LOCAL_LISTING_STATE_KEY_), localListingFolderIdOf_(getReportStateStore_().getProperty(LOCAL_LISTING_FOLDER_KEY_)), new Date()).split('\\n').join('\\n   '));\n\n",
   "  if (localRevisions) results.push('   ' + describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions));\n\n"]
];

/** Code.js without the local listing agent: { code, problems }. */
// The release of T-2026.10.28: the version line and the changelog entry of 2.26.12, taken out with the outermost layer.
const RELEASE_2_26_12_CHANGES = [
  [" * Version: 2.26.12 (2026-11-17)\n", " * Version: 2.26.11 (2026-11-16)\n"],
  [" * CHANGELOG\n * 2.26.12 (2026-11-17)\n *   (T-2026.10.28; on T-2026.10.27 / 2.26.11.)\n *   - Added: the local listing agent. A Windows PowerShell agent on one\n *     laptop in the meeting room reads the listings of MBS, Video,\n *     FS_6G_MED, RTC and Audio every two minutes and writes them (links\n *     only, with a checksum) into a Google Drive folder; every update of a\n *     report imports its own group's listing from there (Configure Meeting\n *     > Local listing folder) with the code of Import Local Inbox Listing...\n *     -- only complete, fresh (30 min), same-group, same-server listings;\n *     a missing, stale or broken one imports nothing and erases nothing.\n *     Check Connections shows scan, write, Drive and import apart.\n *     (tools/local-listing-agent/; docs/LOCAL_LISTING_AGENT.md.)\n *   - Fixed: Check Connections tries the TDoc list like the update (the\n *     configured address, then its www.3gpp.org/ftp twin) and says that a\n *     403 from 3GPP is a refusal, usually temporary.\n", " * CHANGELOG\n"]
];

function codeWithoutLocalListingAgent(code) {
  // The overnight-pause distinction and the Revisions-cell cleanup after T-2026.10.28 came later still: taken out first
  // (newest first); then the release header of T-2026.10.28 (2.26.12).
  const overnight = codeWithoutOvernightPause(code);
  const cleanup = codeWithoutRevisionsCellCleanup(overnight.code);
  const header = takeOut(cleanup.code, null, null, RELEASE_2_26_12_CHANGES, 'RELEASE 2.26.12 HEADER');
  const view = takeOut(header.code, null, null, LOCAL_LISTING_AGENT_CHANGES, 'LOCAL LISTING AGENT');
  return { code: view.code, problems: overnight.problems.concat(cleanup.problems, header.problems, view.problems) };
}

// After the issue #11 live test (2026-11-16): the Session administration section of a main report keeps text typed into
// its lines, and a label that is the code is not written twice. The newest layer: taken out first.
const SESSION_SECTION_NEW_FUNCTIONS = ['withoutOwnCodeLabel_', 'mainSessionLines_'];
const SESSION_SECTION_CHANGED_FUNCTIONS = ['saveAdhocSessionsConfiguration', 'renderMainSessionAdministration_', 'renderMainSessionAdministrationIn_',
  'findMainSessionAdministration_', 'renderMainSessionAdministrationSection_'];
const SESSION_SECTION_CHANGES = [
  ["  const adhoc = meetingType === 'adhoc';\n  // The sessions the section was written from, before this save: their lines are recognized in it (issue #11).\n  // Read outside the lock: if two saves overlap, these can be the older ones. A line that then matches none is\n  // kept as it is and the save says so -- nothing is lost, the next save writes it.\n  let sessionsBefore = null;\n  try { sessionsBefore = adhoc ? null : getConfiguredSessions_(); } catch (e) { sessionsBefore = null; }\n  const result = saveAdhocSessionsWith_(getReportStateStore_(), meetingType !== '', rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays);\n",
   "  const adhoc = meetingType === 'adhoc';\n  const result = saveAdhocSessionsWith_(getReportStateStore_(), meetingType !== '', rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays);\n"],
  ["  if (result.ok && !adhoc) {\n    const notice = renderMainSessionAdministration_(sessionsBefore);\n    if (notice) result.notice = notice;\n",
   "  if (result.ok && !adhoc) {\n    const notice = renderMainSessionAdministration_();\n    if (notice) result.notice = notice;\n"],
  [" */\nfunction renderMainSessionAdministration_(sessionsBefore) {\n  const outcome = withAdhocAttendanceLock_(function () {\n",
   " */\nfunction renderMainSessionAdministration_() {\n  const outcome = withAdhocAttendanceLock_(function () {\n"],
  ["    try {\n      return { ok: true, rendered: renderMainSessionAdministrationIn_(currentReportDocument_().getBody(), sessionsBefore) };\n    } catch (e) {\n",
   "    try {\n      return { ok: true, rendered: renderMainSessionAdministrationIn_(currentReportDocument_().getBody()) };\n    } catch (e) {\n"],
  ["  if (outcome.rendered.container === 'no-place') return 'The sessions were saved. This report has no \"Opening of the session\" section to show them in.';\n  if (outcome.rendered.kept) return 'The sessions were saved. ' + outcome.rendered.kept + ' line(s) of the Session administration section were changed by hand and were kept as they are: please check the section.';\n  return '';\n",
   "  if (outcome.rendered.container === 'no-place') return 'The sessions were saved. This report has no \"Opening of the session\" section to show them in.';\n  return '';\n"],
  ["\n/**\n * Issue #11: the section of a main-meeting report, from its configured\n * sessions, into `body`. `sessionsBefore` are the sessions the section was\n * written from (a save passes those before it); without them, the ones now.\n * Returns what renderMainSessionAdministrationSection_() returns.\n */\nfunction renderMainSessionAdministrationIn_(body, sessionsBefore) {\n  const zone = meetingSessionTimeZone_();\n  const sessions = getConfiguredSessions_();\n  const lines = mainSessionLines_(sessions, zone);\n  const written = sessionsBefore || sessions;\n  return renderMainSessionAdministrationSection_(body, lines, mainSessionLines_(written, zone).concat(mainSessionLines_(written, zone, { legacyLabels: true })));\n}\n",
   "\n/** Issue #11: the section of a main-meeting report, from its configured sessions, into `body`. Returns what renderMainSessionAdministrationSection_() returns. */\nfunction renderMainSessionAdministrationIn_(body) {\n  return renderMainSessionAdministrationSection_(body, buildAdhocOpeningLines_(getConfiguredSessions_(), {}, meetingSessionTimeZone_()));\n}\n"],
  ["/**\n * Pure (issue #11): one line per session of a main report -- [{ id, text,\n * bold }], in the order of `sessions`; a label that is the code is not\n * written (withoutOwnCodeLabel_()). `options.legacyLabels`: as T-2026.10.26\n * wrote them, with such a label written twice (\"A01: A01, ...\").\n */\nfunction mainSessionLines_(sessions, zone, options) {\n  const list = sessions || [];\n  const codes = adhocSessionCodes_(list);\n  // Main reports only: an ad-hoc report writes its lines (buildAdhocOpeningLines_()) as before.\n  return list.map(function (session) {\n    const shown = options && options.legacyLabels ? session : withoutOwnCodeLabel_(session, codes[session.id]);\n    return { id: session.id, text: codes[session.id] + ': ' + adhocSessionWhenText_(shown, zone), bold: true };\n  });\n}\n\n/**\n * Issue #11: the section of a main-meeting report -- { start, end, lines }:\n * its heading (start), the end of its part of the opening (end, exclusive:\n * the next heading or table), and the paragraphs in it that begin with a\n * session code (\"A01: ...\", `lines`, indexes). Text typed between or below\n * them is part of the opening but never one of `lines`. null without one.\n */\n",
   "/**\n * Issue #11: the generated part of a main-meeting report's section --\n * { start, end } (end exclusive): its heading and the \"A01: ...\" lines\n * directly below it -- or null. A main report has no opening details, so\n * every generated line is such a line; text typed below them (minutes of\n * the opening) is not part of it and is never replaced or removed.\n */\n"],
  ["  if (!found) return null;\n  const lines = [];\n  for (let i = found.start + 1; i < found.end; i++) if (/^A\\d{2,}: /.test(body.getChild(i).asParagraph().getText())) lines.push(i);\n  return { start: found.start, end: found.end, lines: lines };\n}\n",
   "  if (!found) return null;\n  let end = found.start + 1;\n  while (end < found.end && /^A\\d{2,}: /.test(body.getChild(end).asParagraph().getText())) end++;\n  return { start: found.start, end: end };\n}\n"],
  ["/**\n * Issue #11: writes `lines` (mainSessionLines_()) as the section of a\n * main-meeting report, and never removes what was typed by hand.\n * `written` are the lines the section was written from (previous sessions,\n * also in the form of T-2026.10.26). Each session line in the section is\n * recognized by its beginning, which is one of them: only that beginning is\n * replaced, text typed after it in the same paragraph stays, with its\n * formatting. A line that is none of them (changed by hand) is kept as it is\n * and counted (`kept`). A session no longer there loses its line, but not\n * what was typed into it. A new session gets a paragraph after the line of\n * the session before it. Text typed between or below the lines is never\n * touched. Without lines the heading goes too. A section that already says\n * this is left alone. Returns { container: 'created' | 'replaced' |\n * 'unchanged' | 'removed' | 'none' | 'no-place', kept }.\n */\n",
   "/**\n * Issue #11: writes `lines` (buildAdhocOpeningLines_()) as the section of a\n * main-meeting report. Only the generated part is replaced; one that already\n * says this is left alone. Without lines the heading and the generated lines\n * are removed, and text typed below them stays. A new section goes where the\n * ad-hoc one goes (renderAdhocOpeningSection_()). Returns { container:\n * 'created' | 'replaced' | 'unchanged' | 'removed' | 'none' | 'no-place' }.\n */\n"],
  [" */\nfunction renderMainSessionAdministrationSection_(body, lines, written) {\n  const existing = findMainSessionAdministration_(body);\n",
   " */\nfunction renderMainSessionAdministrationSection_(body, lines) {\n  const existing = findMainSessionAdministration_(body);\n"],
  ["  if (!existing) return renderAdhocOpeningSection_(body, lines);\n  const heading = body.getChild(existing.start);\n  const known = (written || []).concat(lines);\n  const paras = existing.lines.map(function (i, k) { const p = body.getChild(i).asParagraph(); return { p: p, text: p.getText(), id: '', prefix: '', inBlock: i === existing.start + 1 + k }; });\n  const taken = {};\n  paras.forEach(function (para) {\n    let best = null;\n    known.forEach(function (k) {\n      if (!k.text || taken[k.id]) return;\n      const fits = para.text === k.text || (para.text.indexOf(k.text) === 0 && !/[0-9A-Za-z]/.test(para.text.charAt(k.text.length)));\n      if (fits && (!best || k.text.length > best.text.length)) best = k;\n    });\n    if (best) { taken[best.id] = para; para.id = best.id; para.prefix = best.text; }\n  });\n  // Counted (and said) only in the block directly under the heading: \"A01: discussed X\" typed further below is minutes.\n  const kept = paras.filter(function (para) { return !para.id && para.inBlock; }).length;\n  const wanted = {};\n  lines.forEach(function (line) { wanted[line.id] = true; });\n  let changed = false;\n  // A session no longer there: its line goes; text typed into it stays.\n  paras.forEach(function (para) {\n    if (!para.id || wanted[para.id]) return;\n    if (para.text === para.prefix) { removeAdhocBodyChild_(body, para.p); para.removed = true; }\n    else para.p.editAsText().deleteText(0, para.prefix.length - 1);\n    changed = true;\n  });\n  if (lines.length === 0) {\n",
   "  if (!existing) return renderAdhocOpeningSection_(body, lines);\n  if (lines.length === 0) {\n"],
  ["  if (lines.length === 0) {\n    removeAdhocBodyChild_(body, heading);\n    return { container: 'removed', kept: kept };\n  }\n",
   "  if (lines.length === 0) {\n    for (let i = existing.end - 1; i >= existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));\n    return { container: 'removed' };\n  }\n"],
  ["  }\n  let after = heading;\n  lines.forEach(function (line) {\n",
   "  }\n  const shown = [];\n  for (let i = existing.start + 1; i < existing.end; i++) shown.push(body.getChild(i).asParagraph().getText());\n  if (shown.length === lines.length && shown.every(function (text, i) { return text === lines[i].text; })) return { container: 'unchanged' };\n  for (let i = existing.end - 1; i > existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));\n  let at = existing.start + 1;\n  lines.forEach(function (line) {\n"],
  ["  lines.forEach(function (line) {\n    const para = taken[line.id];\n    if (para && !para.removed) {\n      if (para.prefix !== line.text) {\n        if (para.text === para.prefix) {\n          para.p.setText(line.text);\n          para.p.editAsText().setBold(line.bold);\n        } else {\n          const te = para.p.editAsText();\n          te.deleteText(0, para.prefix.length - 1);\n          te.insertText(0, line.text);\n          te.setBold(0, line.text.length - 1, line.bold);\n        }\n        changed = true;\n      }\n      if (body.getChildIndex(para.p) > body.getChildIndex(after)) after = para.p;\n      return;\n    }\n    const at = body.getChildIndex(after) + 1;\n    // The text is passed as text: nothing in it is interpreted.\n",
   "  lines.forEach(function (line) {\n    // The text is passed as text: nothing in it is interpreted.\n"],
  ["    p.editAsText().setBold(line.bold);\n    after = p;\n    changed = true;\n  });\n",
   "    p.editAsText().setBold(line.bold);\n    at++;\n  });\n"],
  ["  });\n  return { container: changed ? 'replaced' : 'unchanged', kept: kept };\n}\n",
   "  });\n  return { container: 'replaced' };\n}\n"],
  ["  return lines;\n}\n\n/**\n * Pure (issue #11): the session as it is written after its code. A label\n * that is the code itself (\"A01\", any case) would read \"A01: A01, ...\": it\n * is left out. The code is given by the order of the sessions; the label is\n * a name of one's own (\"Plenary\", \"Day 1\").\n */\nfunction withoutOwnCodeLabel_(session, code) {\n  return String((session && session.label) || '').trim().toUpperCase() === String(code || '').toUpperCase() ? Object.assign({}, session, { label: '' }) : session;\n}\n",
   "  return lines;\n}\n"]
];

/** Code.js without the session-section fix after the issue #11 live test: { code, problems }. */
// The release of T-2026.10.27: the version line and the changelog entry of 2.26.11, taken out with the outermost layer.
const RELEASE_2_26_11_CHANGES = [
  [" * Version: 2.26.11 (2026-11-16)\n", " * Version: 2.26.7 (2026-11-12)\n"],
  [" * CHANGELOG\n * 2.26.11 (2026-11-16)\n *   (T-2026.10.27; supersedes the candidate T-2026.10.26 / 2.26.10, which\n *   was never deployed: it lost text typed into a session line.)\n *   - Added: meeting-local and selected IANA time zones (MEETING_TIME_ZONE,\n *     MEETING_TIME_ZONE_MODE); America/Edmonton written ABT, UTC-6 all year\n *     from 1 November 2026. (GitHub issue #4, PR #5.)\n *   - Added: session times in a main-meeting report (Sessions > Configure\n *     Sessions...), shown in Session administration; Build from Scratch\n *     writes the section again and keeps minutes typed below it. Meeting\n *     type, TDocs and agenda filter unchanged. (Issue #11, PR #12, #13.)\n *   - Fixed: saving session times never removes text typed into a session\n *     line (only the generated beginning is replaced); a label that is the\n *     code is not written twice (\"A01: A01\"). (Issue #11, PR #21.)\n *   - Fixed: a main meeting's blank Meeting Folder/Number are taken from\n *     its document folder, never saved as SA4#136; a saved number that\n *     contradicts the folder is refused before any download; a failed\n *     TDoc-list download says \"not updated\". (Issue #9, PR #18.)\n *   - Added: the meeting-room revisions folder (LOCAL_REVISIONS_URL,\n *     default http://10.10.10.10/ftp/SA/SA4/Inbox for a main meeting that\n *     is not an e-meeting) is linked next to the 3GPP drafts folder; never\n *     fetched. (PR #15.)\n *   - Added: Import Local Inbox Listing... -- revisions listed on the\n *     meeting-room server (a listing pasted from a browser on its LAN)\n *     are shown with the 3GPP ones ([local], [local only]); 3GPP\n *     detection and revision tables unchanged. (Issue #16, PR #17;\n *     docs/LOCAL_REVISIONS.md.)\n", " * CHANGELOG\n"]
];

function codeWithoutSessionSectionFix(code) {
  // The local listing agent (T-2026.10.28) came later: taken out first; then the release header of T-2026.10.27 (2.26.11).
  const later = codeWithoutLocalListingAgent(code);
  const header = takeOut(later.code, null, null, RELEASE_2_26_11_CHANGES, 'RELEASE 2.26.11 HEADER');
  const view = takeOut(header.code, null, null, SESSION_SECTION_CHANGES, 'SESSION SECTION FIX');
  return { code: view.code, problems: later.problems.concat(header.problems, view.problems) };
}

/** Code.js without the main-meeting number fix: the Code.js of the T-2026.10.24 candidate (2.26.8). { code, problems }. */
function codeWithoutMainMeetingNumber(code) {
  // The session-section fix after the issue #11 live test came later: taken out first.
  const later = codeWithoutSessionSectionFix(code);
  const view = takeOut(later.code, MAIN_MEETING_NUMBER_SECTION_START, MAIN_MEETING_NUMBER_SECTION_END, MAIN_MEETING_NUMBER_CHANGES, 'MAIN MEETING NUMBER');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

const MEETING_TIME_ZONES_SECTION_START = BAR + '// MEETING TIME ZONES (after T-2026.10.23)\n' + BAR;
const MEETING_TIME_ZONES_SECTION_END = BAR + '// END OF MEETING TIME ZONES\n' + BAR + '\n';
const MEETING_TIME_ZONES_SECTION_FUNCTIONS = ['isAllowedMeetingTimeZoneId_', 'isValidMeetingTimeZoneSetting_', 'utcOffsetLabelMinutes_', 'formatUtcOffsetLabel_', 'meetingTimeZoneRuleChange_', 'meetingTimeZoneOffsetMinutes_',
  'meetingWallClockToInstant_', 'meetingZoneWallClockOfInstant_', 'meetingTimeZoneShownAt_', 'meetingTimeZoneAbbreviationAt_', 'meetingTimeZoneDisplay_', 'meetingTimeZoneOptionLabel_', 'meetingTimeZoneOptionLabelAt_', 'portalNamedTimeZone_', 'meetingLocalTimeZoneFromPortal_', 'meetingTimeZoneFromPortal_', 'meetingLocalTimeZonePreview_',
  'meetingSessionTimeZone_', 'storeMeetingTimeZone_', 'keptMeetingTimeZoneBeforeSave_', 'meetingLocalTimeZoneOfSave_', 'meetingTimeZoneChoiceProblem_', 'dropMeetingStartOfAnotherClock_', 'applyMeetingTimeZoneChoice_', 'meetingTimeZoneDialogState_'];
// After T-2026.10.23 (GitHub issue #4): the changes outside the section, [now, before], in the order they were made.
const MEETING_TIME_ZONES_CHANGES = [
  ["function buildAdhocSessionStatusModel_(sessions, opening, attendance, tdocs) {\n  const total = sessions.length;\n  // After T-2026.10.23: the sessions' zone, read once (MEETING TIME ZONES section).\n  const sessionZone = meetingSessionTimeZone_();\n",
   "function buildAdhocSessionStatusModel_(sessions, opening, attendance, tdocs) {\n  const total = sessions.length;\n"],
  ["function buildAdhocOpeningDialogModel_() {\n  const stored = parseAdhocOpeningProperty_(adhocOpeningStore_().getProperty(ADHOC_OPENING_KEY_));\n  // After T-2026.10.23: the sessions' zone, read once (MEETING TIME ZONES section).\n  const sessionZone = meetingSessionTimeZone_();\n",
   "function buildAdhocOpeningDialogModel_() {\n  const stored = parseAdhocOpeningProperty_(adhocOpeningStore_().getProperty(ADHOC_OPENING_KEY_));\n"],
  // Configure Meeting syncs the zone into an adopted document's central state.
  ["  'MEETING_TYPE', 'MEETING_NAME', 'MEETING_DATE', 'FTP_BASE', 'MAILING_LIST', 'REVISIONS_URL',\n  // After T-2026.10.23: the meeting time zone a save stores -- synced into an adopted document's central state, which\n  // every background update reads for the Session column (MEETING TIME ZONES section).\n  'MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'\n];\n\nfunction getActiveDocumentIdSafely_() {",
   "  'MEETING_TYPE', 'MEETING_NAME', 'MEETING_DATE', 'FTP_BASE', 'MAILING_LIST', 'REVISIONS_URL'\n];\n\nfunction getActiveDocumentIdSafely_() {"],
  // The adoption into the CENTRAL add-on's state copies the sessions' zone (read by every update).
  ["  'ADHOC_TDOC_SESSIONS',\n  // After T-2026.10.23: the meeting's IANA zone, its mode and the meeting it was stored for -- the zone into which\n  // every update converts the TDoc list's UTC upload times for the Session column (MEETING TIME ZONES section).\n  'MEETING_TIME_ZONE', 'MEETING_TIME_ZONE_MODE', 'MEETING_START_BASIS'\n",
   "  'ADHOC_TDOC_SESSIONS'\n"],
  ["    return { time: isValidMeetingStartTime_(time) ? time : '', zone: isValidMeetingTimeZoneSetting_(zone) ? zone : '' };\n",
   "    return { time: isValidMeetingStartTime_(time) ? time : '', zone: isValidMeetingTimeZoneLabel_(zone) ? zone : '' };\n"],
  ["  // After T-2026.10.23: an IANA zone is written as its abbreviation on the meeting date (MEETING TIME ZONES section).\n  const zone = meetingTimeZoneDisplay_(timeZoneLabel, meetingStartDateIso_(meetingDateText), isValidMeetingStartTime_(startTime) ? startTime : '') || '<time zone>';\n",
   "  const zone = isValidMeetingTimeZoneLabel_(timeZoneLabel) ? timeZoneLabel : '<time zone>';\n"],
  ["  const hasZone = isValidMeetingTimeZoneSetting_(zone);\n",
   "  const hasZone = isValidMeetingTimeZoneLabel_(zone);\n"],
  ["  else docProps.deleteProperty(MEETING_START_BASIS_KEY_);\n  // After T-2026.10.23: an IANA zone from the Portal is the meeting-local zone; an offset has no mode (MEETING TIME ZONES section).\n  if (hasZone && isAllowedMeetingTimeZoneId_(zone) && !isValidMeetingTimeZoneLabel_(zone)) docProps.setProperty(MEETING_TIME_ZONE_MODE_KEY_, 'local');\n  else docProps.deleteProperty(MEETING_TIME_ZONE_MODE_KEY_);\n}\n",
   "  else docProps.deleteProperty(MEETING_START_BASIS_KEY_);\n}\n"],
  ["  // After T-2026.10.23: the verified meeting-local zone, else the offset as before (MEETING TIME ZONES section).\n  return { time: computeMeetingStartTimeFromStartDate_(resolved.startDate), zone: meetingTimeZoneFromPortal_(resolved.startDate, resolved.timeZone, resolved.country) };\n",
   "  return { time: computeMeetingStartTimeFromStartDate_(resolved.startDate), zone: computeMeetingTimeZoneLabel_(resolved.timeZone) };\n"],
  ["function persistConfigurationSettings_(config) {\n  const docProps = PropertiesService.getDocumentProperties();\n  const scriptProps = PropertiesService.getScriptProperties();\n  // After T-2026.10.23: the time zone a save asks for is checked before anything is written; a zone selected\n  // for the meeting, or an offset stored before, is kept by a Resolve (MEETING TIME ZONES section).\n  const meetingTimeZoneProblem = meetingTimeZoneChoiceProblem_(config, docProps);\n  if (meetingTimeZoneProblem) throw new Error(meetingTimeZoneProblem);\n  const keptMeetingTimeZone = keptMeetingTimeZoneBeforeSave_(docProps);\n",
   "function persistConfigurationSettings_(config) {\n  const docProps = PropertiesService.getDocumentProperties();\n  const scriptProps = PropertiesService.getScriptProperties();\n"],
  ["    if (resolvedStart) storeMeetingStart_(docProps, resolvedStart.time, resolvedStart.zone);\n  }\n  applyMeetingTimeZoneChoice_(docProps, config, keptMeetingTimeZone);\n",
   "    if (resolvedStart) storeMeetingStart_(docProps, resolvedStart.time, resolvedStart.zone);\n  }\n"],
  ["      location: m.Location !== undefined ? m.Location : (ical ? ical.location : null),\n      // After T-2026.10.23: the meeting's country, which names its local time zone (MEETING TIME ZONES section).\n      country: m.Country !== undefined ? m.Country : null\n",
   "      location: m.Location !== undefined ? m.Location : (ical ? ical.location : null)\n"],
  ["    location: resolvedMeeting && resolvedMeeting.location ? resolvedMeeting.location : null,\n    // After T-2026.10.23 (MEETING TIME ZONES section): the country, and the meeting-local zone the Portal's record gives.\n    country: resolvedMeeting && resolvedMeeting.country ? resolvedMeeting.country : null,\n    meetingLocalTimeZone: resolvedMeeting ? meetingLocalTimeZonePreview_(resolvedMeeting.startDate, resolvedMeeting.timeZone, resolvedMeeting.country) : null\n  };\n",
   "    location: resolvedMeeting && resolvedMeeting.location ? resolvedMeeting.location : null\n  };\n"],
  ["  const initialPreview = computeResolvedMeetingPreview_(existingForPreview, null);\n  // After T-2026.10.23: the meeting time zone -- Meeting local, or a zone selected (MEETING TIME ZONES section).\n  const meetingTimeZoneState = meetingTimeZoneDialogState_(props);\n  const meetingTimeZoneOptions = ['<option value=\"unset\"' + (meetingTimeZoneState.choice === '' ? ' selected=\"selected\"' : '') + '>Not set</option>',\n    '<option value=\"local\"' + (meetingTimeZoneState.choice === 'local' ? ' selected=\"selected\"' : '') + '>Meeting local</option>']\n    .concat(Object.keys(MEETING_TIME_ZONES_).map(function (zone) {\n      return '<option value=\"' + zone + '\"' + (meetingTimeZoneState.choice === zone ? ' selected=\"selected\"' : '') + '>' + meetingTimeZoneOptionLabel_(zone) + '</option>';\n    })).join('');\n",
   "  const initialPreview = computeResolvedMeetingPreview_(existingForPreview, null);\n"],
  ["      <div id=\"meetingSummary\">${esc(initialPreview.summary)}</div>\n      <label>Meeting Time Zone:</label>\n      <select id=\"meetingTimeZoneChoice\" onchange=\"updateMeetingTimeZoneHint()\">${meetingTimeZoneOptions}</select>\n      <input type=\"hidden\" id=\"meetingTimeZoneStored\" value=\"${esc(meetingTimeZoneState.stored)}\">\n      <input type=\"hidden\" id=\"meetingTimeZoneText\" value=\"${esc(meetingTimeZoneState.text)}\">\n      <input type=\"hidden\" id=\"meetingTimeZoneInitial\" value=\"${esc(meetingTimeZoneState.choice || 'unset')}\">\n      <div class=\"hint\" id=\"meetingTimeZoneHint\">${esc(meetingTimeZoneState.text)}</div>\n",
   "      <div id=\"meetingSummary\">${esc(initialPreview.summary)}</div>\n"],
  ["            resolvedMeetingStart = { meetingId: String(result.resolved.id), startDate: result.preview.startDateRaw || '', timeZone: result.preview.startTimeZoneRaw || '', country: result.preview.country || '' };\n            applyPreview(result.preview);\n            onMeetingLocalTimeZone(result.preview.meetingLocalTimeZone);\n",
   "            resolvedMeetingStart = { meetingId: String(result.resolved.id), startDate: result.preview.startDateRaw || '', timeZone: result.preview.startTimeZoneRaw || '' };\n            applyPreview(result.preview);\n"],
  ["      // After T-2026.10.23: Meeting local, as Resolve found it; selected by itself only when no zone is stored.\n      let resolvedLocalTimeZone = null;\n      function onMeetingLocalTimeZone(local) {\n        resolvedLocalTimeZone = local || null;\n        const select = document.getElementById('meetingTimeZoneChoice');\n        if (!select) return;\n        if (local && local.zone && select.value === 'unset' && fieldValue('meetingTimeZoneStored') === 'none') select.value = 'local';\n        updateMeetingTimeZoneHint();\n      }\n      function fieldValue(id) { const el = document.getElementById(id); return el ? String(el.value || '') : ''; }\n      function updateMeetingTimeZoneHint() {\n        const select = document.getElementById('meetingTimeZoneChoice');\n        const hint = document.getElementById('meetingTimeZoneHint');\n        if (!select || !hint) return;\n        if (select.value !== 'unset' && select.value !== 'local') hint.textContent = 'Selected: ' + select.value + '. The meeting\\\\u2019s times are written in this zone.';\n        else hint.textContent = resolvedLocalTimeZone ? resolvedLocalTimeZone.text : fieldValue('meetingTimeZoneText');\n      }\n\n      function resolveMeeting() {\n",
   "      function resolveMeeting() {\n"],
  ["        // After T-2026.10.23: the meeting time zone, sent only when it was changed here (\"Not set\" as 'unset'); a choice\n        // left as the dialog opened is decided by the server for the meeting saved; Meeting local only when it is known\n        // (MEETING TIME ZONES section).\n        const timeZoneSelect = document.getElementById('meetingTimeZoneChoice');\n        if (timeZoneSelect && timeZoneSelect.value && timeZoneSelect.value !== (fieldValue('meetingTimeZoneInitial') || 'unset')) {\n          if (timeZoneSelect.value === 'local' && !(resolvedLocalTimeZone && resolvedLocalTimeZone.zone)) {\n            alert(resolvedLocalTimeZone ? resolvedLocalTimeZone.text : 'Meeting local is not known yet: Resolve the meeting, or select a time zone.');\n            return;\n          }\n          config.meetingTimeZoneChoice = timeZoneSelect.value;\n        }\n        // The start Resolve found goes with the save of the meeting it was found for (see persistConfigurationSettings_()).\n",
   "        // The start Resolve found goes with the save of the meeting it was found for (see persistConfigurationSettings_()).\n"],
  ["function adhocSessionWhenText_(session, zone) {\n  // After T-2026.10.23: with a meeting zone, each time names its zone on its own day (MEETING TIME ZONES section).\n  const at = function (time, day) { const shown = zone && time ? meetingTimeZoneDisplay_(zone, day, time) : ''; return shown ? time + ' ' + shown : time; };\n  const prefix = session.label ? session.label + ', ' : '';\n  if (isAdhocMultiDaySession_(session)) {\n    const parts = [];\n    if (session.start) parts.push('from ' + at(session.start, session.date) + ' on the first day');\n    if (session.end) parts.push('until ' + at(session.end, adhocSessionLastDay_(session)) + ' on the last day');\n    return prefix + adhocSessionDatesText_(session) + (parts.length ? ' (' + parts.join(', ') + ')' : '');\n  }\n  const date = adhocLongDate_(session.date);\n  const startZone = zone && session.start ? meetingTimeZoneDisplay_(zone, session.date, session.start) : '';\n  const endZone = zone && session.end ? meetingTimeZoneDisplay_(zone, session.date, session.end) : '';\n  const time = session.start && session.end\n    ? ', ' + (startZone === endZone ? session.start + '–' + session.end + (endZone ? ' ' + endZone : '') : at(session.start, session.date) + '–' + at(session.end, session.date))\n    : (session.start ? ', ' + at(session.start, session.date) : (session.end ? ', until ' + at(session.end, session.date) : ''));\n  return prefix + date + time;\n}\n",
   "function adhocSessionWhenText_(session) {\n  const prefix = session.label ? session.label + ', ' : '';\n  if (isAdhocMultiDaySession_(session)) {\n    const parts = [];\n    if (session.start) parts.push('from ' + session.start + ' on the first day');\n    if (session.end) parts.push('until ' + session.end + ' on the last day');\n    return prefix + adhocSessionDatesText_(session) + (parts.length ? ' (' + parts.join(', ') + ')' : '');\n  }\n  const date = adhocLongDate_(session.date);\n  const time = session.start && session.end ? ', ' + session.start + '–' + session.end\n    : (session.start ? ', ' + session.start : (session.end ? ', until ' + session.end : ''));\n  return prefix + date + time;\n}\n"],
  ["function buildAdhocOpeningLines_(sessions, details, zone) {\n",
   "function buildAdhocOpeningLines_(sessions, details) {\n"],
  ["    lines.push({ text: codes[session.id] + ': ' + adhocSessionWhenText_(session, zone), bold: true });\n",
   "    lines.push({ text: codes[session.id] + ': ' + adhocSessionWhenText_(session), bold: true });\n"],
  ["  return renderAdhocOpeningSection_(body, buildAdhocOpeningLines_(sessions, stored.sessions, meetingSessionTimeZone_()));\n",
   "  return renderAdhocOpeningSection_(body, buildAdhocOpeningLines_(sessions, stored.sessions));\n"],
  ["      return { id: s.id, when: adhocSessionWhenText_(s, sessionZone), chair: entry.chair || ''",
   "      return { id: s.id, when: adhocSessionWhenText_(s), chair: entry.chair || ''"],
  ["      return { id: s.id, when: adhocSessionWhenText_(s, sessionZone), groups: groups };\n",
   "      return { id: s.id, when: adhocSessionWhenText_(s), groups: groups };\n"],
  ["function convertUtcWallClockToReportZone_(at, zone) {\n",
   "function convertUtcWallClockToReportZone_(at) {\n"],
  ["  // After T-2026.10.23: into the meeting's IANA zone by its rules -- rule changes newer than Apps Script's tz database included (MEETING TIME ZONES section).\n  const instant = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);\n  const text = isAllowedMeetingTimeZoneId_(zone) ? meetingZoneWallClockOfInstant_(zone, instant) : Utilities.formatDate(new Date(instant), Session.getScriptTimeZone(), \"yyyy-MM-dd'T'HH:mm:ss\");\n",
   "  const text = Utilities.formatDate(new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6])), Session.getScriptTimeZone(), \"yyyy-MM-dd'T'HH:mm:ss\");\n"],
  ["  const model = stored.model;\n  // After T-2026.10.23: the sessions' zone -- the meeting's IANA zone, else the script's as before (MEETING TIME ZONES section).\n  const sessionZone = meetingSessionTimeZone_(context);\n  const onSessionClock = function (upload) {\n",
   "  const model = stored.model;\n  const onSessionClock = function (upload) {\n"],
  ["    const at = convertUtcWallClockToReportZone_(upload.at, sessionZone);\n",
   "    const at = convertUtcWallClockToReportZone_(upload.at);\n"],
  ["    sessions: sessions, zone: sessionZone, status: stored.status, error: stored.error, clock: model.clock, overrides: model.tdocs,\n",
   "    sessions: sessions, status: stored.status, error: stored.error, clock: model.clock, overrides: model.tdocs,\n"],
  ["    uploadedZone: resolver.clock === 'utc' ? String(resolver.zone || Session.getScriptTimeZone() || '') : '',\n",
   "    uploadedZone: resolver.clock === 'utc' ? String(Session.getScriptTimeZone() || '') : '',\n"],
  ["  const model = buildAdhocSessionsDialogModel_(props.getProperty(ADHOC_SESSIONS_KEY_), props.getProperty('MEETING_DATE'));\n  // After T-2026.10.23: the zone the times are meant in (MEETING TIME ZONES section).\n  const sessionsZone = meetingSessionTimeZone_();\n  const sessionsZoneText = sessionsZone ? 'the meeting time zone, ' + meetingTimeZoneOptionLabel_(sessionsZone) : 'the report\\'s time zone';\n",
   "  const model = buildAdhocSessionsDialogModel_(props.getProperty(ADHOC_SESSIONS_KEY_), props.getProperty('MEETING_DATE'));\n"],
  ["Times are optional and are meant in ${sessionsZoneText}:",
   "Times are optional and are meant in the report's time zone:"]
];
const MEETING_TIME_ZONES_CREATOR_CHANGES = [
  ["    // After T-2026.10.23 (MEETING TIME ZONES section of Code.js): the verified meeting-local zone, for any\n    // meeting; else, for an ad-hoc meeting, the offset as before.\n    meetingTimeZone: meetingType === 'adhoc' ? meetingTimeZoneFromPortal_(p.startDateRaw, p.startTimeZoneRaw, p.country)\n      : meetingLocalTimeZoneFromPortal_(p.startDateRaw, p.startTimeZoneRaw, p.country).zone\n",
   "    meetingTimeZone: meetingType === 'adhoc' ? computeMeetingTimeZoneLabel_(p.startTimeZoneRaw) : ''\n"],
  ["  if (cfg.meetingTimeZone && !isValidMeetingTimeZoneSetting_(cfg.meetingTimeZone)) errors.push('The meeting time zone is neither a UTC offset nor a supported time zone.');\n",
   "  if (cfg.meetingTimeZone && !isValidMeetingTimeZoneLabel_(cfg.meetingTimeZone)) errors.push('The meeting time zone is not a UTC offset.');\n"]
];

// Issue #16: revisions listed on the meeting-room server (an imported listing), next to those of the 3GPP folder.
// The changes, [now, before], in the order of Code.js; the newest layer, taken out first.
const DUAL_REVISIONS_NEW_FUNCTIONS = ['isLocalRevisionCellLine_', 'manualRevisionCellLines_', 'localListingFileName_', 'revisionCellEntry_', 'revisionFileKey_', 'mergeLocalRevisions_', 'orderedStoredRevisions_', 'resolveLocalListingHref_', 'parseLocalRevisionsListing_',
  'storeLocalRevisions_', 'describeLocalRevisionsStatus_', 'importLocalRevisionsListing', 'saveLocalRevisionsListing'];
const DUAL_REVISIONS_CHANGED_FUNCTIONS = ['renderRevisionsCellContent_', 'updateRevisions_', 'testAllConnections', 'testRevisionsFolder'];
const DUAL_REVISIONS_CHANGES = [
  ["var ADDON003_ADOPTION_PREFIX_FAMILIES_ = [\n  'DISCUSS_', 'REVIS_', 'A1_EMPTY_CACHE_', 'REVIEWER_NO_SUMMARY_CACHE_',\n  // Issue #16: the revisions listed on the meeting-room server, per TDoc (LOCAL REVISIONS LISTING).\n  'LOCALREV_'\n];\n",
   "var ADDON003_ADOPTION_PREFIX_FAMILIES_ = [\n  'DISCUSS_', 'REVIS_', 'A1_EMPTY_CACHE_', 'REVIEWER_NO_SUMMARY_CACHE_'\n];\n"],
  [" */\nfunction renderRevisionsCellContent_(cell, orderedRevisions, folders, options) {\n  // Issue #16: a revision also seen on the meeting-room server carries\n  // `localLink` and gets a marker after its name; without one, as before.\n  const entries = orderedRevisions\n",
   " */\nfunction renderRevisionsCellContent_(cell, orderedRevisions, folders) {\n  const entries = orderedRevisions\n"],
  ["  const entries = orderedRevisions\n    .map(item => revisionCellEntry_(item))\n    .filter(e => e.line);\n",
   "  const entries = orderedRevisions\n    .map(item => ({ line: String(item.text || '').trim(), link: item.link }))\n    .filter(e => e.line);\n"],
  ["      if (!current.trim()) cell.setText(REVISIONS_NONE_TEXT_);\n      return;\n    }\n    // Issue #16: a local revision no longer listed leaves the cell; whatever else is in it stays.\n    const withoutLocal = current.split('\\n').filter(line => !isLocalRevisionCellLine_(line)).join('\\n');\n    if (withoutLocal !== current) {\n      const rest = stripRevisionFoldersLine_(withoutLocal).trim();\n      const keep = rest && rest !== REVISIONS_NONE_TEXT_ ? rest : REVISIONS_NONE_TEXT_;\n      const text = folderLine ? keep + '\\n' + folderLine.text : keep;\n      cell.setText(text);\n      if (folderLine) linkRevisionFoldersLine_(cell.editAsText(), keep.length + 1, folderLine);\n      return;\n",
   "      if (!current.trim()) cell.setText(REVISIONS_NONE_TEXT_);\n      return;\n"],
  ["\n  // Issue #16: a line typed into the cell by hand -- one that is not a\n  // revision line, \"No revisions available.\" or the folders line -- is kept,\n  // below the revisions. A cell that already says exactly this is not written\n  // again (options.force: its links changed).\n  const current = cell.getText();\n  const manual = manualRevisionCellLines_(current, entries);\n  const fullText = entries.map(e => e.line).join('\\n') + (manual.length ? '\\n' + manual.join('\\n') : '') + (folderLine ? '\\n' + folderLine.text : '');\n  if (current === fullText && !(options && options.force)) return;\n  cell.setText(fullText);\n",
   "\n  const fullText = entries.map(e => e.line).join('\\n') + (folderLine ? '\\n' + folderLine.text : '');\n  cell.setText(fullText);\n"],
  ["    const end = start + e.line.length - 1; // setLinkUrl's endOffsetInclusive\n    if (e.link) { try { te.setLinkUrl(start, e.marker ? start + e.name.length - 1 : end, e.link); } catch (err) { } }\n    if (e.marker) { try { te.setLinkUrl(end - e.marker.length + 2, end, e.localLink); } catch (err) { } }\n    offset = end + 2; // skip past this line's '\\n' separator\n",
   "    const end = start + e.line.length - 1; // setLinkUrl's endOffsetInclusive\n    if (e.link) { try { te.setLinkUrl(start, end, e.link); } catch (err) { } }\n    offset = end + 2; // skip past this line's '\\n' separator\n"],
  ["  });\n  manual.forEach(line => { offset += line.length + 1; });\n  if (folderLine) linkRevisionFoldersLine_(te, offset, folderLine);\n",
   "  });\n  if (folderLine) linkRevisionFoldersLine_(te, offset, folderLine);\n"],
  ["  if (folderLine) linkRevisionFoldersLine_(te, offset, folderLine);\n}\n\n/** Pure (issue #16): whether a line of a Revisions cell is a local revision line written by renderRevisionsCellContent_(). */\nfunction isLocalRevisionCellLine_(line) {\n  const t = String(line || '').trim();\n  if (!LOCAL_REVISION_MARKER_PATTERN_.test(t)) return false;\n  const name = t.replace(LOCAL_REVISION_MARKER_PATTERN_, '');\n  return /\\.[A-Za-z0-9]{1,5}$/.test(name) && parseSA4DocumentId_(name).isValid;\n}\n\n/**\n * Issue #16 (pure): the lines of a Revisions cell that were typed by hand.\n * Not one: a line of `entries`, a revision line written before (a file name\n * with an SA4 TDoc number and an extension, with or without a local marker),\n * \"No revisions available.\", the folders line, an empty line.\n */\nfunction manualRevisionCellLines_(text, entries) {\n  const shown = {};\n  (entries || []).forEach(e => { shown[e.line] = true; shown[e.name] = true; });\n  return String(text || '').split('\\n').filter(line => {\n    const t = line.trim();\n    if (!t || t === REVISIONS_NONE_TEXT_ || line.indexOf(REVISION_FOLDERS_PREFIX_) === 0 || shown[t]) return false;\n    const name = t.replace(LOCAL_REVISION_MARKER_PATTERN_, '');\n    if (shown[name]) return false;\n    return !(/\\.[A-Za-z0-9]{1,5}$/.test(name) && parseSA4DocumentId_(name).isValid);\n  });\n}\n\n/**\n * Issue #16 (pure): one line of a Revisions cell. A revision from the 3GPP\n * folder only is its file name, linked there -- as before. One also listed\n * on the meeting-room server gets \" [local]\", linked to its file there; one\n * listed there only gets \" [local only]\" and both parts link to the local\n * file. `marker` is the text after the name, including its leading space.\n */\nfunction revisionCellEntry_(item) {\n  const name = String((item && item.text) || '').trim();\n  const localLink = String((item && item.localLink) || '');\n  if (!name || !localLink) return { line: name, name: name, link: item && item.link, marker: '', localLink: '' };\n  // A listing copied as text only gives no file link: the local folder is linked, and the marker says so.\n  const marker = item.localFolderOnly ? (item.link ? LOCAL_REVISION_MARKERS_.bothFolder : LOCAL_REVISION_MARKERS_.localOnlyFolder)\n    : (item.link ? LOCAL_REVISION_MARKERS_.both : LOCAL_REVISION_MARKERS_.localOnly);\n  return { line: name + marker, name: name, link: item.link || localLink, marker: marker, localLink: localLink };\n}\n",
   "  if (folderLine) linkRevisionFoldersLine_(te, offset, folderLine);\n}\n"],
  ["\n// ---------------------------------------------------------------------------\n// LOCAL REVISIONS LISTING (issue #16)\n// The meeting-room server (LOCAL_REVISIONS_URL, e.g. http://10.10.10.10/...)\n// is on the meeting-room LAN only. Apps Script runs at Google and cannot\n// reach it, and the report's dialogs cannot either: they are https pages of\n// googleusercontent.com, so a request to an http address on a private\n// network is blocked (mixed content, Private Network Access, no CORS).\n// What can reach it is the browser of someone in the room, as an ordinary\n// page. So the listing comes from there: open the local folder in the\n// browser, select all, copy, and paste it into \"Import Local Inbox\n// Listing…\". Nothing is installed, nothing on the LAN is exposed.\n//\n// What is kept: per TDoc of this report, the files of the listing that are\n// revisions of it (the same parser as the 3GPP folder: parseDraftAnchor_()),\n// with the link the listing gave -- never a path built from the 3GPP folder\n// -- in LOCALREV_<tdoc>, which only grows (like REVIS_<tdoc>). A file of\n// another origin than the listed folder is dropped. TDocs that have no\n// table are not imported. The 3GPP folder, its fetch and the tables of\n// revisions are not changed by any of this; a listing that cannot be read\n// changes nothing.\n// ---------------------------------------------------------------------------\nconst LOCAL_REVISIONS_KEY_PREFIX_ = 'LOCALREV_';\nconst LOCAL_REVISIONS_LAST_IMPORT_KEY_ = 'LOCAL_REVISIONS_LAST_IMPORT';\nconst LOCAL_REVISION_MARKERS_ = { both: ' [local]', localOnly: ' [local only]', bothFolder: ' [local folder]', localOnlyFolder: ' [local only, see folder]' };\nconst LOCAL_REVISION_MARKER_PATTERN_ = / \\[local(?: only)?(?: folder|, see folder)?\\]$/;\n\n/** Pure: the identity of a file in either folder -- its decoded name, case ignored. */\nfunction revisionFileKey_(name) {\n  let n = String(name || '').trim();\n  try { n = decodeURIComponent(n); } catch (e) { }\n  return n.toLowerCase();\n}\n\n/**\n * Pure: the revisions of the 3GPP folder (`ordered`, as updateRevisions_()\n * orders them) with the ones of the meeting-room server (`localStore`,\n * LOCALREV_<tdoc>) -- the same file once, with `localLink`; a file listed\n * locally only with `link` '' and `localLink`. In the same order. Without\n * local ones, `ordered` itself.\n */\nfunction mergeLocalRevisions_(ordered, localStore) {\n  const locals = Object.keys(localStore || {}).map(function (k) { return localStore[k]; })\n    .filter(function (l) { return l && String(l.text || '').trim() && /^https?:\\/\\//i.test(String(l.link || '')); });\n  if (!locals.length) return ordered;\n  const byKey = {};\n  const merged = ordered.map(function (item) {\n    const copy = { text: item.text, link: item.link };\n    const key = revisionFileKey_(item.text);\n    if (key && !byKey[key]) byKey[key] = copy;\n    return copy;\n  });\n  locals.forEach(function (l) {\n    const key = revisionFileKey_(l.text);\n    if (byKey[key]) {\n      if (!byKey[key].localLink) { byKey[key].localLink = l.link; if (l.folderOnly) byKey[key].localFolderOnly = true; }\n      return;\n    }\n    const item = { text: String(l.text).trim(), link: '', localLink: l.link };\n    if (l.folderOnly) item.localFolderOnly = true;\n    byKey[key] = item;\n    merged.push(item);\n  });\n  return merged.sort(function (x, y) {\n    const byText = String(x.text || '').localeCompare(String(y.text || ''));\n    return byText !== 0 ? byText : String(x.link || x.localLink || '').localeCompare(String(y.link || y.localLink || ''));\n  });\n}\n\n/** Pure: the revisions of REVIS_<tdoc>, deduplicated and ordered exactly as updateRevisions_() does. */\nfunction orderedStoredRevisions_(store) {\n  const deduped = {};\n  Object.values(store || {}).forEach(item => {\n    const key = normalizeRevisionKey_(item);\n    if (!key) return;\n    if (!deduped[key]) deduped[key] = item;\n  });\n  return Object.values(deduped).sort((x, y) => {\n    const byText = String(x.text || '').localeCompare(String(y.text || ''));\n    return byText !== 0 ? byText : String(x.link || '').localeCompare(String(y.link || ''));\n  });\n}\n\n/** Pure: `href` of a listing of `folder` as an absolute URL on the same server, or '' (another server, or not a file link). */\nfunction resolveLocalListingHref_(folder, href) {\n  const h = String(href || '').trim().replace(/&amp;/g, '&');\n  const origin = (String(folder).match(/^(https?:\\/\\/[^\\/?#]+)/i) || [])[1];\n  if (!h || !origin || /^(?:javascript|data|mailto):/i.test(h) || h.charAt(0) === '?' || h.charAt(0) === '#') return '';\n  let url;\n  if (/^https?:\\/\\//i.test(h)) url = h;\n  else if (h.indexOf('//') === 0) return '';\n  else if (h.charAt(0) === '/') url = origin + h;\n  else url = String(folder).replace(/[?#].*$/, '').replace(/\\/?$/, '/') + h;\n  return url.toLowerCase().indexOf(origin.toLowerCase() + '/') === 0 ? url : '';\n}\n\n/**\n * Pure: what a pasted listing says. `payload` { folder, html, text } is\n * what the dialog sends: the address of the listed folder, and the\n * clipboard as HTML (its links) and as text. Returns { ok, error, folder,\n * anchors: [{ href (absolute), text }], listed, truncated }. Text alone (a\n * browser that copies no HTML) gives file names; one cut short by the\n * listing (\"..>\") is not used. `localBase` is the report's local folder:\n * only a folder on its server is accepted.\n */\nfunction parseLocalRevisionsListing_(payload, localBase) {\n  const folder = String((payload && payload.folder) || '').trim();\n  const fail = function (error) { return { ok: false, error: error, folder: folder, anchors: [], listed: 0, truncated: 0 }; };\n  const baseOrigin = (String(localBase || '').match(/^(https?:\\/\\/[^\\/?#]+)/i) || [])[1];\n  if (!baseOrigin) return fail('This report has no Local Revisions URL (Configure Meeting).');\n  if (!isLocalRevisionsUrlValue_(folder)) return fail('The address of the listed folder is not an http or https address.');\n  if (folder.toLowerCase().indexOf(baseOrigin.toLowerCase() + '/') !== 0 && folder.toLowerCase() !== baseOrigin.toLowerCase()) {\n    return fail('The listed folder is not on the meeting-room server of this report (' + baseOrigin + ').');\n  }\n  const html = String((payload && payload.html) || '');\n  const text = String((payload && payload.text) || '');\n  if (html.length + text.length > 2000000) return fail('The pasted listing is too large.');\n  const anchors = [];\n  let truncated = 0;\n  parseAnchors_(html.replace(/href='([^']*)'/gi, 'href=\"$1\"')).forEach(function (a) {\n    const url = resolveLocalListingHref_(folder, a.href);\n    if (!url) return;\n    // The name is the one of the linked file -- a listing may show it shortened (\"S4-2601..>\"). A folder link has none.\n    const name = localListingFileName_(url) || (/\\/$/.test(url.replace(/[?#].*$/, '')) ? '' : a.text);\n    if (name) anchors.push({ href: url, text: name });\n  });\n  if (!anchors.length && text) {\n    text.split(/[\\r\\n\\t]+|\\s{2,}/).forEach(function (token) {\n      const name = token.trim();\n      if (!name) return;\n      if (/\\.\\.>$|\\u2026$/.test(name)) { if (parseSA4DocumentId_(name).isValid) truncated++; return; }\n      if (!/^[^\\/\\\\?#<>\"]+\\.[A-Za-z0-9]{1,5}$/.test(name)) return;\n      // Text gives no link: the file is linked by its folder (which the user opened), never by a path made up here.\n      anchors.push({ href: folder.replace(/\\/?$/, '/'), text: name, folderOnly: true });\n    });\n  }\n  if (!anchors.length) return fail(truncated ? 'The pasted text has only shortened file names. Copy the page again, or use another browser.' : 'The pasted text has no file list. Nothing was changed.');\n  return { ok: true, error: '', folder: folder, anchors: anchors, listed: anchors.length, truncated: truncated };\n}\n\n/** Pure: the decoded last path segment of a file URL; '' for a folder (ending in \"/\"). */\nfunction localListingFileName_(url) {\n  const path = String(url || '').replace(/[?#].*$/, '');\n  if (/\\/$/.test(path)) return '';\n  const segment = path.slice(path.lastIndexOf('/') + 1);\n  try { return decodeURIComponent(segment).trim(); } catch (e) { return segment.trim(); }\n}\n\n/**\n * Stores the revisions of a parsed listing for the TDocs in `reportTdocs`\n * (TDoc numbers that have a table). A listing is the current state of its\n * folder: an entry from the same folder that the listing no longer has is\n * removed (renamed, withdrawn), one from another folder is kept; the same\n * file twice is one entry. A listing without a single revision removes\n * nothing (a wrong or partial paste). A direct file link is never replaced\n * by a folder link (a text-only paste of the same folder). Returns {\n * touched: [tdoc], added, removed, recognized }.\n */\nfunction storeLocalRevisions_(props, parsed, reportTdocs) {\n  const folderKey = String(parsed.folder || '').toLowerCase().replace(/\\/?$/, '/');\n  const fromFolder = function (e) { return String((e && e.folder) || '').toLowerCase().replace(/\\/?$/, '/') === folderKey; };\n  const files = {};\n  parsed.anchors.forEach(function (a) { files[a.href + '\\u0000' + a.text] = a; });\n  const index = buildRevisionAnchorIndex_(parsed.anchors.map(function (a) { return { href: a.href, text: a.text }; }), parsed.folder);\n  const wanted = {};\n  (reportTdocs || []).forEach(function (t) { wanted[t] = true; });\n  let recognized = 0;\n  index.forEach(function (drafts) { recognized += drafts.length; });\n  const folderOnly = {};\n  parsed.anchors.forEach(function (a) { if (a.folderOnly) folderOnly[revisionFileKey_(a.text)] = true; });\n\n  // The TDocs with stored entries (one read of the keys), and the ones of this listing.\n  const tdocs = {};\n  (typeof props.getKeys === 'function' ? props.getKeys() : []).forEach(function (k) {\n    if (k.indexOf(LOCAL_REVISIONS_KEY_PREFIX_) === 0 && wanted[k.slice(LOCAL_REVISIONS_KEY_PREFIX_.length)]) tdocs[k.slice(LOCAL_REVISIONS_KEY_PREFIX_.length)] = true;\n  });\n  index.forEach(function (drafts, tdoc) { if (wanted[tdoc]) tdocs[tdoc] = true; });\n\n  let added = 0, removed = 0;\n  const touched = [];\n  Object.keys(tdocs).sort().forEach(function (tdoc) {\n    const key = LOCAL_REVISIONS_KEY_PREFIX_ + tdoc;\n    const store = loadJsonObject_(props.getProperty(key));\n    const listed = {};\n    (index.get(tdoc) || []).forEach(function (d) {\n      const id = revisionFileKey_(d.fileName);\n      if (!id || listed[id]) return;\n      listed[id] = { text: d.fileName, link: d.url, folder: parsed.folder };\n      if (folderOnly[id]) listed[id].folderOnly = true;\n    });\n    let changed = false;\n    if (recognized > 0) {\n      Object.keys(store).forEach(function (id) {\n        if (fromFolder(store[id]) && !listed[id]) { delete store[id]; removed++; changed = true; }\n      });\n    }\n    Object.keys(listed).forEach(function (id) {\n      const before = store[id];\n      const next = listed[id];\n      if (before && before.link && !before.folderOnly && next.folderOnly) return;  // keep the file link\n      if (before && before.link === next.link && !!before.folderOnly === !!next.folderOnly && fromFolder(before)) return;\n      if (!before) added++;\n      store[id] = next;\n      changed = true;\n    });\n    if (!changed) return;\n    if (Object.keys(store).length) props.setProperty(key, JSON.stringify(store)); else props.deleteProperty(key);\n    touched.push(tdoc);\n  });\n  return { touched: touched, added: added, removed: removed, recognized: recognized };\n}\n\n/** Pure: what the report says about the meeting-room server: not checked yet, or when a listing was last imported. */\nfunction describeLocalRevisionsStatus_(lastImportJson, localBase) {\n  if (!localBase) return 'Local (meeting-room server): none configured.';\n  const last = loadJsonObject_(lastImportJson);\n  if (!last.at) return 'Local (meeting-room server): not checked yet. Import Local Inbox Listing… on the meeting-room network.';\n  return 'Local (meeting-room server): last listing imported ' + last.at + ' from ' + last.folder + ': ' + last.listed + ' file(s), ' +\n    last.recognized + ' revision(s) of SA4 TDocs, ' + (last.added ? last.added + ' new for this report' : 'no new revisions for this report') +\n    (last.removed ? ', ' + last.removed + ' no longer listed (removed)' : '') + '.';\n}\n\n/** Menu: Import Local Inbox Listing… -- the dialog that takes a pasted listing of the meeting-room server. */\nfunction importLocalRevisionsListing() {\n  const ui = DocumentApp.getUi();\n  const localBase = localRevisionsUrlFor_();\n  if (!localBase) {\n    ui.alert('Import Local Inbox Listing', 'This report has no Local Revisions URL. Set one in Configure Meeting first.', ui.ButtonSet.OK);\n    return;\n  }\n  const esc = function (v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\"/g, '&quot;'); };\n  const status = describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localBase);\n  const html = `<!DOCTYPE html><html><head><base target=\"_top\"><style>\n    body { font-family: Arial, sans-serif; font-size: 13px; margin: 12px; }\n    #paste { border: 1px dashed #888; min-height: 90px; max-height: 160px; overflow: auto; padding: 6px; margin: 6px 0; color: #555; }\n    input[type=text] { width: 100%; box-sizing: border-box; }\n    .hint { color: #555; font-size: 12px; } #result { margin-top: 8px; white-space: pre-wrap; }\n  </style></head><body>\n    <div class=\"hint\">${esc(status)}</div>\n    <ol>\n      <li>On the meeting-room network, open <a href=\"${esc(localBase)}\" target=\"_blank\" rel=\"noopener\">${esc(localBase)}</a> (or the folder of this group under it) in this browser.</li>\n      <li>On that page: Ctrl+A, Ctrl+C. Copy its address too.</li>\n      <li>Paste the address below, click into the box and press Ctrl+V. Then Import.</li>\n    </ol>\n    <label>Address of the listed folder:</label>\n    <input type=\"text\" id=\"folder\" value=\"${esc(localBase)}\">\n    <div id=\"paste\" contenteditable=\"true\">Paste the page here (Ctrl+V)</div>\n    <button id=\"go\" onclick=\"go()\" disabled>Import</button> <button onclick=\"google.script.host.close()\">Close</button>\n    <div id=\"result\"></div>\n    <script>\n      var clip = { html: '', text: '' };\n      document.getElementById('paste').addEventListener('paste', function (e) {\n        e.preventDefault();\n        clip.html = e.clipboardData.getData('text/html') || '';\n        clip.text = e.clipboardData.getData('text/plain') || '';\n        var links = (clip.html.match(/<a[\\\\s>]/gi) || []).length;\n        this.textContent = 'Pasted: ' + (links ? links + ' link(s)' : clip.text.split(/\\\\n/).length + ' line(s) of text');\n        document.getElementById('go').disabled = !(clip.html || clip.text);\n      });\n      function go() {\n        document.getElementById('go').disabled = true;\n        document.getElementById('result').textContent = 'Importing…';\n        google.script.run\n          .withSuccessHandler(function (r) { document.getElementById('result').textContent = r.message; document.getElementById('go').disabled = false; })\n          .withFailureHandler(function (e) { document.getElementById('result').textContent = 'Not imported: ' + e.message; document.getElementById('go').disabled = false; })\n          .saveLocalRevisionsListing({ folder: document.getElementById('folder').value, html: clip.html, text: clip.text });\n      }\n    </script></body></html>`;\n  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(470), 'Import Local Inbox Listing');\n}\n\n/**\n * The dialog's Import. Stores the revisions the listing has for this\n * report's TDocs and writes the Revisions cells of those TDocs (from what\n * is stored for both folders -- nothing is fetched). Returns { ok, message }.\n */\nfunction saveLocalRevisionsListing(payload) {\n  const localBase = localRevisionsUrlFor_();\n  const parsed = parseLocalRevisionsListing_(payload, localBase);\n  if (!parsed.ok) return { ok: false, message: parsed.error };\n  const lock = LockService.getDocumentLock();\n  if (!lock.tryLock(30000)) return { ok: false, message: 'The report is being updated. Try again in a minute; nothing was changed.' };\n  try {\n    const props = getReportStateStore_();\n    const body = currentReportDocument_().getBody();\n    const tables = body.getTables().filter(isTDocTable_);\n    const byTdoc = {};\n    tables.forEach(function (t) {\n      const id = parseExactSA4DocumentId_(String(safeCellText_(t, 0, 1) || '').trim());\n      if (id.isValid && !byTdoc[id.raw]) byTdoc[id.raw] = t;\n    });\n    const stored = storeLocalRevisions_(props, parsed, Object.keys(byTdoc));\n    const remoteBase = String(getMeetingContext_().sources.revisionsUrl || '').trim();\n    const folders = { local: localBase, remote: remoteBase ? remoteBase.replace(/\\/$/, '') + '/' : '' };\n    stored.touched.forEach(function (tdoc) {\n      const t = byTdoc[tdoc];\n      const ordered = orderedStoredRevisions_(loadJsonObject_(props.getProperty('REVIS_' + tdoc)));\n      const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n      renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders, { force: true });\n    });\n    const at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');\n    const last = { at: at, folder: parsed.folder, listed: parsed.listed, recognized: stored.recognized, added: stored.added, removed: stored.removed };\n    props.setProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_, JSON.stringify(last));\n    Logger.log('Local revisions: ' + JSON.stringify(last) + ', TDocs: ' + stored.touched.join(', '));\n    return { ok: true, message: describeLocalRevisionsStatus_(JSON.stringify(last), localBase) +\n      (stored.touched.length ? '\\nUpdated: ' + stored.touched.join(', ') : '') +\n      (parsed.truncated ? '\\n' + parsed.truncated + ' shortened file name(s) were not used.' : '') +\n      (parsed.anchors.some(function (a) { return a.folderOnly; }) ? '\\nThe browser copied text only: these revisions link to the local folder, not to their files. Paste the page again from Chrome or Edge for file links.' : '') };\n  } finally {\n    lock.releaseLock();\n  }\n}\n\nfunction updateRevisions_(cfg, context) {\n",
   "\nfunction updateRevisions_(cfg, context) {\n"],
  ["    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n    // Issue #16: with the revisions listed on the meeting-room server (LOCALREV_<tdoc>), which are never fetched\n    // from here. Only the 3GPP ones below get a table of their own, as before.\n    renderRevisionsCellContent_(cell, mergeLocalRevisions_(ordered, loadJsonObject_(props.getProperty(LOCAL_REVISIONS_KEY_PREFIX_ + tdoc))), folders);\n    renderingMs += Date.now() - renderStart;\n",
   "    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n    renderRevisionsCellContent_(cell, ordered, folders);\n    renderingMs += Date.now() - renderStart;\n"],
  ["  if (localRevisions) results.push('   🏢 Local (meeting-room server, on-site only, not checked from here): ' + localRevisions);\n  // Issue #16: when a listing of it was last imported.\n  if (localRevisions) results.push('   ' + describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions));\n\n",
   "  if (localRevisions) results.push('   🏢 Local (meeting-room server, on-site only, not checked from here): ' + localRevisions);\n\n"],
  ["      `Files found: ~${linkCount}` +\n      (localRevisions ? `\\n\\nLocal (meeting-room server, on-site only, not checked from here): ${localRevisions}` +\n        // Issue #16: when a listing of it was last imported.\n        '\\n' + describeLocalRevisionsStatus_(getReportStateStore_().getProperty(LOCAL_REVISIONS_LAST_IMPORT_KEY_), localRevisions) : ''),\n      ui.ButtonSet.OK\n",
   "      `Files found: ~${linkCount}` +\n      (localRevisions ? `\\n\\nLocal (meeting-room server, on-site only, not checked from here): ${localRevisions}` : ''),\n      ui.ButtonSet.OK\n"]
];
const DUAL_REVISIONS_CREATOR_CHANGED_FUNCTIONS = ['buildTemplateReportMenu_'];
const DUAL_REVISIONS_CREATOR_CHANGES = [
  ["    .addItem('Update TDoc Revisions', 'collectRevisionsOnly')\n    .addItem('Import Local Inbox Listing…', 'importLocalRevisionsListing')\n    .addItem('Update Abstracts', 'addAbstractsOnly')\n",
   "    .addItem('Update TDoc Revisions', 'collectRevisionsOnly')\n    .addItem('Update Abstracts', 'addAbstractsOnly')\n"]
];

/** Code.js without the imported listings of the meeting-room server (issue #16): { code, problems }. */
function codeWithoutDualRevisions(code) {
  // The main-meeting number fix (issue #9, brought to master after issue #16) came later: taken out first.
  const later = codeWithoutMainMeetingNumber(code);
  const view = takeOut(later.code, null, null, DUAL_REVISIONS_CHANGES, 'DUAL REVISIONS');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** template/ReportCreator.js without them: { code, problems }. */
function creatorWithoutDualRevisions(code) {
  return takeOut(code, null, null, DUAL_REVISIONS_CREATOR_CHANGES, 'DUAL REVISIONS (ReportCreator.js)');
}

// The local revisions folder (after T-2026.10.25): a link to the meeting-room server next to the 3GPP drafts folder.
// The changes, [now, before], in the order of Code.js; the newest layer, taken out first.
const LOCAL_REVISIONS_NEW_FUNCTIONS = ['resolveLocalRevisionsUrl_', 'isLocalRevisionsUrlValue_', 'localRevisionsUrlFor_', 'buildRevisionFoldersLine_', 'stripRevisionFoldersLine_', 'linkRevisionFoldersLine_'];
const LOCAL_REVISIONS_CHANGED_FUNCTIONS = ['renderRevisionsCellContent_', 'updateRevisions_', 'configureMeetingSettings', 'persistConfigurationSettings_', 'testAllConnections', 'testRevisionsFolder', 'buildAdhocMeetingInformationLines_', 'adhocMeetingInformationSource_'];
const LOCAL_REVISIONS_CHANGES = [
  ["  'MEETING_TYPE', 'MEETING_NAME', 'REVISIONS_URL', 'MAILING_LIST',\n  // After T-2026.10.25: the local (meeting-room) revisions folder (LOCAL REVISIONS FOLDER).\n  'LOCAL_REVISIONS_URL',\n  'FETCH_ABSTRACTS_ON_UPDATE',\n",
   "  'MEETING_TYPE', 'MEETING_NAME', 'REVISIONS_URL', 'MAILING_LIST',\n  'FETCH_ABSTRACTS_ON_UPDATE',\n"],
  [" */\nfunction renderRevisionsCellContent_(cell, orderedRevisions, folders) {\n  const entries = orderedRevisions\n",
   " */\nfunction renderRevisionsCellContent_(cell, orderedRevisions) {\n  const entries = orderedRevisions\n"],
  ["    .filter(e => e.line);\n  // After T-2026.10.25: a last line linking both revisions folders, when the\n  // report has a local (meeting-room) folder. Without one, nothing changes.\n  const folderLine = buildRevisionFoldersLine_(folders);\n\n",
   "    .filter(e => e.line);\n\n"],
  ["  if (!entries.length) {\n    const current = cell.getText();\n    if (!folderLine && current.indexOf(REVISION_FOLDERS_PREFIX_) === -1) {\n      if (!current.trim()) cell.setText(REVISIONS_NONE_TEXT_);\n      return;\n    }\n    // Text that is not ours (typed by hand) is left alone, as before.\n    const own = stripRevisionFoldersLine_(current).trim();\n    if (own && own !== REVISIONS_NONE_TEXT_) return;\n    const wanted = folderLine ? REVISIONS_NONE_TEXT_ + '\\n' + folderLine.text : REVISIONS_NONE_TEXT_;\n    if (current === wanted) return;\n    cell.setText(wanted);\n    if (folderLine) linkRevisionFoldersLine_(cell.editAsText(), REVISIONS_NONE_TEXT_.length + 1, folderLine);\n    return;\n",
   "  if (!entries.length) {\n    if (!cell.getText().trim()) cell.setText('No revisions available.');\n    return;\n"],
  ["\n  const fullText = entries.map(e => e.line).join('\\n') + (folderLine ? '\\n' + folderLine.text : '');\n  cell.setText(fullText);\n",
   "\n  const fullText = entries.map(e => e.line).join('\\n');\n  cell.setText(fullText);\n"],
  ["  });\n  if (folderLine) linkRevisionFoldersLine_(te, offset, folderLine);\n}\n\n// ---------------------------------------------------------------------------\n// LOCAL REVISIONS FOLDER (after T-2026.10.25)\n// During a physical SA4 meeting, revised documents can appear first on the\n// meeting-room server (LAN only); the 3GPP folder (REVISIONS_URL) is\n// synchronized from it later. The report links to both folders. The local\n// one is only ever a link: it is never fetched (a private address cannot be\n// reached from Google), and revisions are found, as before, in the 3GPP\n// folder alone. No direct link to a file on the local server is made: its\n// path there cannot be known reliably, so the folder is linked instead.\n//\n// LOCAL_REVISIONS_URL (Document Property, set in Configure Meeting):\n//   absent  -- the default below for a main meeting that is not an\n//              e-meeting (meeting folder \"...-e\"); no local folder otherwise;\n//   'none'  -- no local folder;\n//   a URL   -- that folder (http or https).\n// ---------------------------------------------------------------------------\nconst LOCAL_REVISIONS_URL_DEFAULT_ = 'http://10.10.10.10/ftp/SA/SA4/Inbox';\nconst REVISIONS_NONE_TEXT_ = 'No revisions available.';\nconst REVISION_FOLDERS_PREFIX_ = 'Revisions folders: ';\nconst REVISION_FOLDER_LABELS_ = {\n  local: 'Local \\u2013 meeting-room server, may be newer, on-site only',\n  remote: 'Remote \\u2013 3GPP server, synchronized, accessible externally'\n};\n\n/**\n * Pure: the local revisions folder of a report -- { url, source }, source\n * 'configured' | 'default' | 'off' | 'none'. `stored` is the raw\n * LOCAL_REVISIONS_URL, `meetingType` 'main' | 'adhoc', `ftpBase` the\n * meeting's Docs folder (an e-meeting has no meeting room).\n */\nfunction resolveLocalRevisionsUrl_(stored, meetingType, ftpBase) {\n  const value = String(stored === null || stored === undefined ? '' : stored).trim();\n  if (/^none$/i.test(value)) return { url: '', source: 'off' };\n  if (isLocalRevisionsUrlValue_(value)) return { url: value, source: 'configured' };\n  if (value) return { url: '', source: 'none' }; // not a URL: never linked\n  if (meetingType !== 'main') return { url: '', source: 'none' };\n  if (/[-_]e\\/(?:Docs\\/?)?$/i.test(String(ftpBase || '').trim())) return { url: '', source: 'none' };\n  return { url: LOCAL_REVISIONS_URL_DEFAULT_, source: 'default' };\n}\n\n/** Pure: whether a value can be stored as LOCAL_REVISIONS_URL as a folder (http or https, nothing else). */\nfunction isLocalRevisionsUrlValue_(value) {\n  return /^https?:\\/\\/[^\\s\"<>]+$/i.test(String(value || '').trim());\n}\n\n/** The local revisions folder of this report, from its stored configuration ('' when it has none). Never fetched. */\nfunction localRevisionsUrlFor_(context) {\n  try {\n    const props = getReportStateStore_(context);\n    const identity = getMeetingIdentityConfig_(context);\n    let ftpBase = identity.FTP_BASE;\n    if (!ftpBase && identity.MEETING_TYPE === 'main') ftpBase = String(getReportConfig_(context).FTP_BASE || '');\n    return resolveLocalRevisionsUrl_(props.getProperty('LOCAL_REVISIONS_URL'), identity.MEETING_TYPE, ftpBase).url;\n  } catch (e) {\n    Logger.log('Local revisions folder: not read (' + e.message + ')');\n    return '';\n  }\n}\n\n/**\n * Pure: the line naming both folders -- { text, links: [{ start, end, url }] }\n * (offsets within the line, end inclusive) -- or null without a local folder.\n */\nfunction buildRevisionFoldersLine_(folders) {\n  const local = String((folders && folders.local) || '').trim();\n  const remote = String((folders && folders.remote) || '').trim();\n  if (!local) return null;\n  const parts = [{ label: REVISION_FOLDER_LABELS_.local, url: local }];\n  if (remote) parts.push({ label: REVISION_FOLDER_LABELS_.remote, url: remote });\n  let text = REVISION_FOLDERS_PREFIX_;\n  const links = [];\n  parts.forEach(function (part, i) {\n    if (i) text += ' \\u00b7 ';\n    links.push({ start: text.length, end: text.length + part.label.length - 1, url: part.url });\n    text += part.label;\n  });\n  return { text: text, links: links };\n}\n\n/** Pure: a cell's text without the line written by buildRevisionFoldersLine_(). */\nfunction stripRevisionFoldersLine_(text) {\n  return String(text || '').split('\\n').filter(function (line) { return line.indexOf(REVISION_FOLDERS_PREFIX_) !== 0; }).join('\\n');\n}\n\n/** Links the folders line that starts at `offset` of the text, and only its two labels. */\nfunction linkRevisionFoldersLine_(te, offset, folderLine) {\n  try {\n    te.setLinkUrl(offset, offset + folderLine.text.length - 1, null);\n    folderLine.links.forEach(function (l) { te.setLinkUrl(offset + l.start, offset + l.end, l.url); });\n  } catch (err) { }\n}\n",
   "  });\n}\n"],
  ["  perfCount_('revision anchors parsed', anchors.length);\n  // After T-2026.10.25: both folders, for the line under the revisions (the local one is never fetched).\n  const folders = { local: localRevisionsUrlFor_(context), remote: url };\n\n",
   "  perfCount_('revision anchors parsed', anchors.length);\n\n"],
  ["    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n    renderRevisionsCellContent_(cell, ordered, folders);\n    renderingMs += Date.now() - renderStart;\n",
   "    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);\n    renderRevisionsCellContent_(cell, ordered);\n    renderingMs += Date.now() - renderStart;\n"],
  ["      <div class=\"hint\">Suggested drafts/revisions folder. Review if needed.</div>\n\n      <label>Local Revisions URL (meeting-room server):</label>\n      <input type=\"text\" id=\"localRevisionsUrl\" value=\"${esc(props.getProperty('LOCAL_REVISIONS_URL') || '')}\" placeholder=\"${esc(LOCAL_REVISIONS_URL_DEFAULT_)}\">\n      <div class=\"hint\">On-site only and possibly newer than the 3GPP folder above, which the updates keep reading. The report only links to it, never reads it. Blank: ${esc(LOCAL_REVISIONS_URL_DEFAULT_)} for a main meeting that is not an e-meeting; \"none\": no local link.</div>\n    </div>\n",
   "      <div class=\"hint\">Suggested drafts/revisions folder. Review if needed.</div>\n    </div>\n"],
  ["          revisionsUrl: document.getElementById('revisionsUrl').value,\n          localRevisionsUrl: document.getElementById('localRevisionsUrl') ? document.getElementById('localRevisionsUrl').value : undefined,\n          reportType: document.getElementById('reportType').value,\n",
   "          revisionsUrl: document.getElementById('revisionsUrl').value,\n          reportType: document.getElementById('reportType').value,\n"],
  ["  'MEETING_TYPE', 'MEETING_NAME', 'MEETING_DATE', 'FTP_BASE', 'MAILING_LIST', 'REVISIONS_URL',\n  // After T-2026.10.25: the local (meeting-room) revisions folder, read by every update (LOCAL REVISIONS FOLDER).\n  'LOCAL_REVISIONS_URL',\n  // After T-2026.10.23: the meeting time zone a save stores -- synced into an adopted document's central state, which\n",
   "  'MEETING_TYPE', 'MEETING_NAME', 'MEETING_DATE', 'FTP_BASE', 'MAILING_LIST', 'REVISIONS_URL',\n  // After T-2026.10.23: the meeting time zone a save stores -- synced into an adopted document's central state, which\n"],
  ["    docProps.setProperty('REVISIONS_URL', config.revisionsUrl.trim());\n  }\n  // After T-2026.10.25 (LOCAL REVISIONS FOLDER): only a dialog that has the field changes it. Blank removes it (the\n  // default applies again), \"none\" switches it off, a folder URL is stored; anything else is not stored.\n  if (config.localRevisionsUrl !== undefined && config.localRevisionsUrl !== null) {\n    const localRevisions = String(config.localRevisionsUrl).trim();\n    if (!localRevisions) docProps.deleteProperty('LOCAL_REVISIONS_URL');\n    else if (/^none$/i.test(localRevisions)) docProps.setProperty('LOCAL_REVISIONS_URL', 'none');\n    else if (isLocalRevisionsUrlValue_(localRevisions)) docProps.setProperty('LOCAL_REVISIONS_URL', localRevisions);\n    else Logger.log('Configure Meeting: the Local Revisions URL was not stored (not an http or https address): ' + localRevisions);\n  }\n",
   "    docProps.setProperty('REVISIONS_URL', config.revisionsUrl.trim());\n  }\n"],
  ["  }\n  // After T-2026.10.25: the local folder is named, never fetched (meeting-room LAN only).\n  const localRevisions = localRevisionsUrlFor_();\n  if (localRevisions) results.push('   🏢 Local (meeting-room server, on-site only, not checked from here): ' + localRevisions);\n\n  results.push('\\n\\n✅ Connection test complete!');\n",
   "  }\n  \n  results.push('\\n\\n✅ Connection test complete!');\n"],
  ["    \n    // After T-2026.10.25: the local folder is named, never fetched (meeting-room LAN only).\n    const localRevisions = localRevisionsUrlFor_();\n    ui.alert(\n",
   "    \n    ui.alert(\n"],
  ["      `URL: ${url}\\n` +\n      `Files found: ~${linkCount}` +\n      (localRevisions ? `\\n\\nLocal (meeting-room server, on-site only, not checked from here): ${localRevisions}` : ''),\n      ui.ButtonSet.OK\n",
   "      `URL: ${url}\\n` +\n      `Files found: ~${linkCount}`,\n      ui.ButtonSet.OK\n"],
  ["// The labels of its lines, in the order they are written. A line of the block is a paragraph that begins with one of them.\nconst ADHOC_MEETING_INFO_LABELS_ = ['Meeting name', 'Start Date', 'End Date', 'Portal meeting URL', 'TDoc List URL', 'Excel Docs URL', 'Docs Folder', 'Drafts Folder',\n  // After T-2026.10.25: with a local (meeting-room) revisions folder, the drafts folder is named twice (LOCAL REVISIONS FOLDER).\n  'Drafts Folder (local \\u2013 meeting-room server, may be newer, on-site only)',\n  'Drafts Folder (remote \\u2013 3GPP server, synchronized, accessible externally)'];\n\n",
   "// The labels of its lines, in the order they are written. A line of the block is a paragraph that begins with one of them.\nconst ADHOC_MEETING_INFO_LABELS_ = ['Meeting name', 'Start Date', 'End Date', 'Portal meeting URL', 'TDoc List URL', 'Excel Docs URL', 'Docs Folder', 'Drafts Folder'];\n\n"],
  ["  add(L[6], url(source && source.docsFolder), true);\n  // After T-2026.10.25: a local drafts folder comes first and the 3GPP one is named \"remote\"; without one, as before.\n  const localDrafts = url(source && source.localDraftsFolder);\n  if (localDrafts) add(L[8], localDrafts, true);\n  add(localDrafts ? L[9] : L[7], url(source && source.draftsFolder), true);\n  return lines;\n",
   "  add(L[6], url(source && source.docsFolder), true);\n  add(L[7], url(source && source.draftsFolder), true);\n  return lines;\n"],
  ["  const named = String(identity.TDOC_LIST_URL || '').match(/^https:\\/\\/portal\\.3gpp\\.org\\/[^?#\\s]*\\?(?:[^#\\s]*&)?meetingId=(\\d+)(?:&|$)/i);\n  // After T-2026.10.25: a local drafts folder (LOCAL REVISIONS FOLDER) is shown with the 3GPP folder the updates\n  // read -- for a main meeting without a stored one, the folder derived for it.\n  const localDraftsFolder = localRevisionsUrlFor_();\n  const draftsFolder = identity.REVISIONS_URL || (localDraftsFolder && identity.MEETING_TYPE === 'main' ? String(getReportConfig_().REVISIONS_URL || '') : '');\n  return { meetingName: identity.MEETING_NAME, meetingId: named ? named[1] : '', tdocListUrl: identity.TDOC_LIST_URL, docsFolder: identity.FTP_BASE, draftsFolder: draftsFolder,\n    localDraftsFolder: localDraftsFolder };\n}\n",
   "  const named = String(identity.TDOC_LIST_URL || '').match(/^https:\\/\\/portal\\.3gpp\\.org\\/[^?#\\s]*\\?(?:[^#\\s]*&)?meetingId=(\\d+)(?:&|$)/i);\n  return { meetingName: identity.MEETING_NAME, meetingId: named ? named[1] : '', tdocListUrl: identity.TDOC_LIST_URL, docsFolder: identity.FTP_BASE, draftsFolder: identity.REVISIONS_URL };\n}\n"]
];

/** Code.js without the local revisions folder: { code, problems }. */
function codeWithoutLocalRevisions(code) {
  // The listings of issue #16 came later: taken out first.
  const later = codeWithoutDualRevisions(code);
  const view = takeOut(later.code, null, null, LOCAL_REVISIONS_CHANGES, 'LOCAL REVISIONS FOLDER');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

// GitHub issue #11 (after T-2026.10.23): session times in a main-meeting report. The changes, [now, before],
// in the order of Code.js; taken out before the meeting time zones, which came earlier.
const SESSION_TIMES_NEW_FUNCTIONS = ['sessionTimesMeetingType_', 'getConfiguredSessions_', 'renderMainSessionAdministration_', 'renderMainSessionAdministrationIn_', 'findMainSessionAdministration_', 'renderMainSessionAdministrationSection_'];
const SESSION_TIMES_CHANGED_FUNCTIONS = ['getAdhocSessions_', 'saveAdhocSessionsWith_', 'saveAdhocSessionsConfiguration', 'configureAdhocSessions', 'finishAdhocOpeningRebuild_'];
const SESSION_TIMES_CHANGES = [
  ["\nconst ADHOC_SESSIONS_KEY_ = 'ADHOC_SESSIONS';\n// Issue #11: sessions need a report configured as an ad-hoc or a main-meeting report.\nconst SESSION_TIMES_UNAVAILABLE_ = 'Sessions are available for ad-hoc and main-meeting reports only.';\nconst ADHOC_SESSIONS_SCHEMA_VERSION_ = 1;\nconst ADHOC_SESSION_LABEL_MAX_LENGTH_ = 16;\n",
   "\nconst ADHOC_SESSIONS_KEY_ = 'ADHOC_SESSIONS';\nconst ADHOC_SESSIONS_SCHEMA_VERSION_ = 1;\nconst ADHOC_SESSION_LABEL_MAX_LENGTH_ = 16;\n"],
  ["function getAdhocSessions_(context) {\n  if (!isAdhocMeetingForSessions_(context)) return [];\n  return getConfiguredSessions_(context);\n}\n\n/**\n * Issue #11: the meeting type under which session times can be edited --\n * 'adhoc' or 'main', as stored; '' for a report without a meeting type or\n * with one that cannot be read. Never throws.\n */\nfunction sessionTimesMeetingType_(context) {\n  try {\n    const raw = getReportStateStore_(context).getProperty('MEETING_TYPE');\n    return String(raw || '').trim() ? normalizeMeetingType_(raw) : '';\n  } catch (e) {\n    return '';\n  }\n}\n\n/**\n * The stored sessions in chronological order, each with its `displayLabel`,\n * whatever the meeting type; [] when none can be read. Never throws. Only the\n * Configure Sessions dialog and the Session administration section of a\n * main-meeting report read sessions this way (issue #11); everything that\n * selects, groups or places TDocs reads getAdhocSessions_(), which stays\n * [] outside an ad-hoc report.\n */\nfunction getConfiguredSessions_(context) {\n  let stored;\n  try {\n",
   "function getAdhocSessions_(context) {\n  if (!isAdhocMeetingForSessions_(context)) return [];\n  let stored;\n  try {\n"],
  [" */\nfunction saveAdhocSessionsWith_(store, isAdhoc, rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays) {\n  if (!isAdhoc) return { ok: false, errors: [SESSION_TIMES_UNAVAILABLE_], changed: false, sessions: [] };\n  const raw = store.getProperty(ADHOC_SESSIONS_KEY_);\n  const edit = applyAdhocSessionsEdit_(raw, rows);\n",
   " */\nfunction saveAdhocSessionsWith_(store, isAdhoc, rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays) {\n  if (!isAdhoc) return { ok: false, errors: ['Sessions are available for ad-hoc reports only.'], changed: false, sessions: [] };\n  const raw = store.getProperty(ADHOC_SESSIONS_KEY_);\n  const edit = applyAdhocSessionsEdit_(raw, rows);\n"],
  ["    openingSessionIds = null;\n  }\n  // Issue #11: a main-meeting report can have session times too. Only\n  // ADHOC_SESSIONS is written; the meeting type stays as stored.\n  const meetingType = sessionTimesMeetingType_();\n  const adhoc = meetingType === 'adhoc';\n  const result = saveAdhocSessionsWith_(getReportStateStore_(), meetingType !== '', rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays);\n\n  // Issue #11: in a main-meeting report the times are shown in the Session\n  // administration section, and nothing else is written: no Online\n  // information, no attendance, no Session column, no TDoc.\n  if (result.ok && !adhoc) {\n    const notice = renderMainSessionAdministration_();\n    if (notice) result.notice = notice;\n    return result;\n  }\n\n  // Stage E: the Session administration section says when each session is,\n",
   "    openingSessionIds = null;\n  }\n  const result = saveAdhocSessionsWith_(getReportStateStore_(), isAdhocMeetingForSessions_(), rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays);\n\n  // Stage E: the Session administration section says when each session is,\n"],
  ["}\n\n/**\n * Issue #11: writes the Session administration section of a main-meeting\n * report from its configured sessions -- their codes, dates and times in the\n * meeting time zone -- and nothing else. A section that already says this is\n * left alone; without sessions it is removed. Text outside the section is\n * never touched. Returns '' or a notice for the dialog.\n */\nfunction renderMainSessionAdministration_() {\n  const outcome = withAdhocAttendanceLock_(function () {\n    try {\n      return { ok: true, rendered: renderMainSessionAdministrationIn_(currentReportDocument_().getBody()) };\n    } catch (e) {\n      Logger.log('Session times: stored, but the Session administration section could not be written: ' + e.message);\n      return { ok: false, error: e.message };\n    }\n  });\n  if (!outcome.ok) return 'The sessions were saved. The Session administration section could not be written (' + outcome.error + '). Save the sessions again to write it.';\n  if (outcome.rendered.container === 'no-place') return 'The sessions were saved. This report has no \"Opening of the session\" section to show them in.';\n  return '';\n}\n\n/** Issue #11: the section of a main-meeting report, from its configured sessions, into `body`. Returns what renderMainSessionAdministrationSection_() returns. */\nfunction renderMainSessionAdministrationIn_(body) {\n  return renderMainSessionAdministrationSection_(body, buildAdhocOpeningLines_(getConfiguredSessions_(), {}, meetingSessionTimeZone_()));\n}\n\n/**\n * Issue #11: the generated part of a main-meeting report's section --\n * { start, end } (end exclusive): its heading and the \"A01: ...\" lines\n * directly below it -- or null. A main report has no opening details, so\n * every generated line is such a line; text typed below them (minutes of\n * the opening) is not part of it and is never replaced or removed.\n */\nfunction findMainSessionAdministration_(body) {\n  const found = findAdhocOpeningContainer_(body);\n  if (!found) return null;\n  let end = found.start + 1;\n  while (end < found.end && /^A\\d{2,}: /.test(body.getChild(end).asParagraph().getText())) end++;\n  return { start: found.start, end: end };\n}\n\n/**\n * Issue #11: writes `lines` (buildAdhocOpeningLines_()) as the section of a\n * main-meeting report. Only the generated part is replaced; one that already\n * says this is left alone. Without lines the heading and the generated lines\n * are removed, and text typed below them stays. A new section goes where the\n * ad-hoc one goes (renderAdhocOpeningSection_()). Returns { container:\n * 'created' | 'replaced' | 'unchanged' | 'removed' | 'none' | 'no-place' }.\n */\nfunction renderMainSessionAdministrationSection_(body, lines) {\n  const existing = findMainSessionAdministration_(body);\n  if (!existing) return renderAdhocOpeningSection_(body, lines);\n  if (lines.length === 0) {\n    for (let i = existing.end - 1; i >= existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));\n    return { container: 'removed' };\n  }\n  const shown = [];\n  for (let i = existing.start + 1; i < existing.end; i++) shown.push(body.getChild(i).asParagraph().getText());\n  if (shown.length === lines.length && shown.every(function (text, i) { return text === lines[i].text; })) return { container: 'unchanged' };\n  for (let i = existing.end - 1; i > existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));\n  let at = existing.start + 1;\n  lines.forEach(function (line) {\n    // The text is passed as text: nothing in it is interpreted.\n    const p = at >= body.getNumChildren() ? body.appendParagraph(line.text) : body.insertParagraph(at, line.text);\n    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);\n    p.editAsText().setBold(line.bold);\n    at++;\n  });\n  return { container: 'replaced' };\n}\n\n/**\n * What the Configure Sessions dialog shows (pure; reads nothing but its\n",
   "}\n\n/**\n * What the Configure Sessions dialog shows (pure; reads nothing but its\n"],
  ["  assertNotTemplateMaster_();\n  const ui = DocumentApp.getUi();\n  // Issue #11: an ad-hoc or a main-meeting report; the meeting type is only read.\n  const meetingType = sessionTimesMeetingType_();\n  if (!meetingType) {\n    ui.alert('Configure Sessions', SESSION_TIMES_UNAVAILABLE_, ui.ButtonSet.OK);\n    return;\n  }\n",
   "  assertNotTemplateMaster_();\n  const ui = DocumentApp.getUi();\n  if (!isAdhocMeetingForSessions_()) {\n    ui.alert('Configure Sessions', 'Sessions are available for ad-hoc reports only.', ui.ButtonSet.OK);\n    return;\n  }\n"],
  ["\n    <h2>Sessions</h2>\n    <div class=\"hint\">${meetingType === 'adhoc' ? 'The meetings of this ad-hoc series.' : 'The sessions of this meeting. They are shown in the Session administration section; the TDocs of this report are not changed.'} A session can span several days: enter its last day as End date, or leave End date empty for a session of one day. Times are optional and are meant in ${sessionsZoneText}: the start time is on the first day, the end time is the final cut-off on the last day. An empty label is shown as the day and month of the dates.</div>\n    <div id=\"notice\"></div>\n\n",
   "\n    <h2>Sessions</h2>\n    <div class=\"hint\">The meetings of this ad-hoc series. A session can span several days: enter its last day as End date, or leave End date empty for a session of one day. Times are optional and are meant in ${sessionsZoneText}: the start time is on the first day, the end time is the final cut-off on the last day. An empty label is shown as the day and month of the dates.</div>\n    <div id=\"notice\"></div>\n\n"],
  ["/** Build from Scratch, at the end: writes the container again. Returns a note for the build result, '' when all is well. */\nfunction finishAdhocOpeningRebuild_(body) {\n  // Issue #11: a main-meeting report with session times gets its section again; without sessions nothing is done.\n  if (sessionTimesMeetingType_() === 'main') {\n    if (getConfiguredSessions_().length === 0) return '';\n    try {\n      const rendered = renderMainSessionAdministrationIn_(body);\n      return rendered.container === 'no-place' ? '\\n\\n⚠️ The Session administration section was not written: the report has no \"Opening of the session\" section.' : '';\n    } catch (e) {\n      Logger.log('Opening details: the Session administration section of a main-meeting report could not be written after the rebuild: ' + e.message);\n      return '\\n\\n⚠️ The Session administration section could not be written (' + e.message + '). Save Configure Sessions… again to write it.';\n    }\n  }\n  if (!adhocSessionsEnabled_()) return '';\n  try {\n",
   "/** Build from Scratch, at the end: writes the container again. Returns a note for the build result, '' when all is well. */\nfunction finishAdhocOpeningRebuild_(body) {\n  if (!adhocSessionsEnabled_()) return '';\n  try {\n"]
];
const SESSION_TIMES_CREATOR_NEW_FUNCTIONS = ['isMainReportForMenu_'];
const SESSION_TIMES_CREATOR_CHANGED_FUNCTIONS = ['buildTemplateReportMenu_'];
const SESSION_TIMES_CREATOR_CHANGES = [
  ["      .addItem('Edit Opening Details…', 'editAdhocOpeningDetails')\n      .addItem('Post-meeting Statistics…', 'showAdhocSessionStatistics'));\n  } else if (isMainReportForMenu_()) {\n    // Issue #11: a main-meeting report can have session times, and nothing else of the ad-hoc sessions.\n    menu.addSubMenu(ui.createMenu('🗓 Sessions')\n      .addItem('Configure Sessions…', 'configureAdhocSessions'));\n  }\n\n",
   "      .addItem('Edit Opening Details…', 'editAdhocOpeningDetails')\n      .addItem('Post-meeting Statistics…', 'showAdhocSessionStatistics'));\n  }\n\n"],
  ["}\n\n/** Issue #11: whether this document is configured as a main-meeting report; read as isAdhocReportForMenu_() reads. */\nfunction isMainReportForMenu_() {\n  try {\n    return String(PropertiesService.getDocumentProperties().getProperty('MEETING_TYPE') || '').trim().toLowerCase() === 'main';\n  } catch (e) {\n    return false;\n  }\n}\n\nfunction addTemplateReportMenuTail_(menu) {\n  menu.addItem('ℹ️ About This Report', 'showTemplateInfo');\n",
   "}\n\nfunction addTemplateReportMenuTail_(menu) {\n  menu.addItem('ℹ️ About This Report', 'showTemplateInfo');\n"]
];

// The follow-up of issue #11 (PR #13: Build from Scratch keeps the section, minutes typed below it are kept), merged
// after the main-meeting number fix (PR #18). Its changes are also in SESSION_TIMES_CHANGES (the session times as a
// whole); these pairs take only them out, for a comparison with the master before them (tests/main-meeting-number.test.js).
const SESSION_TIMES_BUILD_CHANGES = [
  ["    try {\n      return { ok: true, rendered: renderMainSessionAdministrationIn_(currentReportDocument_().getBody()) };\n    } catch (e) {\n",
   "    try {\n      const lines = buildAdhocOpeningLines_(getConfiguredSessions_(), {}, meetingSessionTimeZone_());\n      return { ok: true, rendered: renderAdhocOpeningSection_(currentReportDocument_().getBody(), lines) };\n    } catch (e) {\n"],
  ["  return '';\n}\n\n/** Issue #11: the section of a main-meeting report, from its configured sessions, into `body`. Returns what renderMainSessionAdministrationSection_() returns. */\nfunction renderMainSessionAdministrationIn_(body) {\n  return renderMainSessionAdministrationSection_(body, buildAdhocOpeningLines_(getConfiguredSessions_(), {}, meetingSessionTimeZone_()));\n}\n\n/**\n * Issue #11: the generated part of a main-meeting report's section --\n * { start, end } (end exclusive): its heading and the \"A01: ...\" lines\n * directly below it -- or null. A main report has no opening details, so\n * every generated line is such a line; text typed below them (minutes of\n * the opening) is not part of it and is never replaced or removed.\n */\nfunction findMainSessionAdministration_(body) {\n  const found = findAdhocOpeningContainer_(body);\n  if (!found) return null;\n  let end = found.start + 1;\n  while (end < found.end && /^A\\d{2,}: /.test(body.getChild(end).asParagraph().getText())) end++;\n  return { start: found.start, end: end };\n}\n\n/**\n * Issue #11: writes `lines` (buildAdhocOpeningLines_()) as the section of a\n * main-meeting report. Only the generated part is replaced; one that already\n * says this is left alone. Without lines the heading and the generated lines\n * are removed, and text typed below them stays. A new section goes where the\n * ad-hoc one goes (renderAdhocOpeningSection_()). Returns { container:\n * 'created' | 'replaced' | 'unchanged' | 'removed' | 'none' | 'no-place' }.\n */\nfunction renderMainSessionAdministrationSection_(body, lines) {\n  const existing = findMainSessionAdministration_(body);\n  if (!existing) return renderAdhocOpeningSection_(body, lines);\n  if (lines.length === 0) {\n    for (let i = existing.end - 1; i >= existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));\n    return { container: 'removed' };\n  }\n  const shown = [];\n  for (let i = existing.start + 1; i < existing.end; i++) shown.push(body.getChild(i).asParagraph().getText());\n  if (shown.length === lines.length && shown.every(function (text, i) { return text === lines[i].text; })) return { container: 'unchanged' };\n  for (let i = existing.end - 1; i > existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));\n  let at = existing.start + 1;\n  lines.forEach(function (line) {\n    // The text is passed as text: nothing in it is interpreted.\n    const p = at >= body.getNumChildren() ? body.appendParagraph(line.text) : body.insertParagraph(at, line.text);\n    p.setHeading(DocumentApp.ParagraphHeading.NORMAL);\n    p.editAsText().setBold(line.bold);\n    at++;\n  });\n  return { container: 'replaced' };\n}\n",
   "  return '';\n}\n"],
  ["function finishAdhocOpeningRebuild_(body) {\n  // Issue #11: a main-meeting report with session times gets its section again; without sessions nothing is done.\n  if (sessionTimesMeetingType_() === 'main') {\n    if (getConfiguredSessions_().length === 0) return '';\n    try {\n      const rendered = renderMainSessionAdministrationIn_(body);\n      return rendered.container === 'no-place' ? '\\n\\n⚠️ The Session administration section was not written: the report has no \"Opening of the session\" section.' : '';\n    } catch (e) {\n      Logger.log('Opening details: the Session administration section of a main-meeting report could not be written after the rebuild: ' + e.message);\n      return '\\n\\n⚠️ The Session administration section could not be written (' + e.message + '). Save Configure Sessions… again to write it.';\n    }\n  }\n  if (!adhocSessionsEnabled_()) return '';\n",
   "function finishAdhocOpeningRebuild_(body) {\n  if (!adhocSessionsEnabled_()) return '';\n"]
];

/** Code.js without the follow-up of issue #11 (PR #13) alone: { code, problems }. */
function codeWithoutSessionTimesBuild(code) {
  return takeOut(code, null, null, SESSION_TIMES_BUILD_CHANGES, 'SESSION TIMES BUILD (PR #13)');
}

/** Code.js without the session times of issue #11: { code, problems }. */
function codeWithoutSessionTimes(code) {
  // The local revisions folder came later: taken out first.
  const later = codeWithoutLocalRevisions(code);
  const view = takeOut(later.code, null, null, SESSION_TIMES_CHANGES, 'SESSION TIMES');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** template/ReportCreator.js without the session times of issue #11: { code, problems }. */
function creatorWithoutSessionTimes(code) {
  // The menu item of issue #16 came later: taken out first.
  const later = creatorWithoutDualRevisions(code);
  const view = takeOut(later.code, null, null, SESSION_TIMES_CREATOR_CHANGES, 'SESSION TIMES (ReportCreator.js)');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** Code.js without the meeting time zones: the Code.js of T-2026.10.23. { code, problems }. */
function codeWithoutMeetingTimeZones(code) {
  // The session times of issue #11 came later: taken out first.
  const later = codeWithoutSessionTimes(code);
  const view = takeOut(later.code, MEETING_TIME_ZONES_SECTION_START, MEETING_TIME_ZONES_SECTION_END, MEETING_TIME_ZONES_CHANGES, 'MEETING TIME ZONES');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** template/ReportCreator.js without the meeting time zones: { code, problems }. */
function creatorWithoutMeetingTimeZones(code) {
  // The session times of issue #11 came later: taken out first.
  const later = creatorWithoutSessionTimes(code);
  const view = takeOut(later.code, null, null, MEETING_TIME_ZONES_CREATOR_CHANGES, 'MEETING TIME ZONES (ReportCreator.js)');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

const TDOC_LIST_DOWNLOAD_SECTION_START = BAR + '// TDOC LIST DOWNLOAD (after T-2026.10.21)\n' + BAR;
const TDOC_LIST_DOWNLOAD_SECTION_END = BAR + '// END OF TDOC LIST DOWNLOAD\n' + BAR + '\n';
const TDOC_LIST_DOWNLOAD_SECTION_FUNCTIONS = ['tdocListDownloadCandidates_', 'tdocListDownloadAttempt_', 'downloadTdocListXlsx_', 'diagnoseTdocListDownload'];
const TDOC_LIST_DOWNLOAD_CHANGED_FUNCTIONS = ['downloadAndGroupTdocs_'];
const TDOC_LIST_DOWNLOAD_CHANGES = [
  // The release of it (T-2026.10.22): the version line and the changelog entry of 2.26.6.
  [" * Version: 2.26.6 (2026-11-11)\n",
   " * Version: 2.26.5 (2026-11-11)\n"],
  [" * CHANGELOG\n * 2.26.6 (2026-11-11)\n *   - Fixed: the TDoc list is downloaded when its configured address on\n *     ftp.3gpp.org is refused. The configured address is asked first;\n *     only if it fails, and only for ftp.3gpp.org, the same path under\n *     https://www.3gpp.org/ftp/. An answer counts only with HTTP 2xx, a\n *     non-empty body and the XLSX signature, so an HTML page sent with\n *     HTTP 200 is refused. The configured address is unchanged and stays\n *     the list's identity; how the list is read is unchanged.\n *     diagnoseTdocListDownload(), run from the editor, asks every address\n *     with the same check and writes nothing. (MBS SA4#138, 2026-11-11:\n *     Full Report Build stopped at \"Could not download TDOC list: HTTP\n *     403\".) (TDOC LIST DOWNLOAD section; docs/TDOC_LIST_DOWNLOAD.md.)\n",
   " * CHANGELOG\n"],
  ["  // After T-2026.10.21: checked as an XLSX, the www.3gpp.org/ftp twin of an ftp.3gpp.org address as fallback (TDOC LIST DOWNLOAD section).\n" +
   "  const blob = perfTimed_('  TDoc-list: UrlFetchApp.fetch (download xlsx)', () => downloadTdocListXlsx_(meetingUrl)).blob;\n",
   "  const response = perfTimed_('  TDoc-list: UrlFetchApp.fetch (download xlsx)', () => UrlFetchApp.fetch(meetingUrl, { muteHttpExceptions: true }));\n" +
   "  if (response.getResponseCode() >= 400) {\n    throw new Error('Could not download TDOC list: HTTP ' + response.getResponseCode());\n  }\n\n" +
   "  const blob = response.getBlob();\n"]
];

/** Code.js without the TDoc-list download fallback: the Code.js of T-2026.10.21. { code, problems }. */
function codeWithoutTdocListDownload(code) {
  // The large-build checkpoints (after T-2026.10.22) came later: taken out first.
  const later = codeWithoutLargeBuildCheckpoints(code);
  const view = takeOut(later.code, TDOC_LIST_DOWNLOAD_SECTION_START, TDOC_LIST_DOWNLOAD_SECTION_END, TDOC_LIST_DOWNLOAD_CHANGES, 'TDOC LIST DOWNLOAD');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** Code.js without the agenda readers: the Code.js of T-2026.10.20. { code, problems }. */
function codeWithoutAgendaReaders(code) {
  const later = codeWithoutTdocListDownload(code);
  const view = takeOut(later.code, AGENDA_READERS_SECTION_START, AGENDA_READERS_SECTION_END, AGENDA_READERS_CHANGES, 'AGENDA READERS');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** Code.js without the agenda Word conversion: the Code.js of T-2026.10.19. { code, problems }. */
function codeWithoutAgendaWordConversion(code) {
  const later = codeWithoutAgendaReaders(code);
  const view = takeOut(later.code, AGENDA_CONVERSION_SECTION_START, AGENDA_CONVERSION_SECTION_END, AGENDA_CONVERSION_CHANGES, 'AGENDA WORD CONVERSION');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** Code.js without the Portal-page fallback: the Code.js of T-2026.10.18. { code, problems }. */
function codeWithoutCloseReportPageFallback(code) {
  const later = codeWithoutAgendaWordConversion(code);
  const view = takeOut(later.code, PAGE_FALLBACK_SECTION_START, PAGE_FALLBACK_SECTION_END, PAGE_FALLBACK_CHANGES, 'CLOSE REPORT PAGE FALLBACK');
  return { code: view.code, problems: later.problems.concat(view.problems) };
}

/** Code.js without the report-local scope: the Code.js of T-2026.10.17. { code, problems }. */
function codeWithoutReportLocalScope(code) {
  const later = codeWithoutCloseReportPageFallback(code);
  const view = takeOut(later.code, null, null, REPORT_LOCAL_SCOPE_CHANGES, 'REPORT LOCAL SCOPE');
  return { code: view.code, problems: later.problems.concat(view.problems) };
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
  // The meeting time zones (after T-2026.10.22) came later: taken out first.
  const later = creatorWithoutMeetingTimeZones(code);
  const text = later.code;
  const at = text.indexOf(DISPOSITION_CREATOR_CHANGE);
  if (at === -1 || text.indexOf(DISPOSITION_CREATOR_CHANGE, at + 1) !== -1) return { code: text, problems: later.problems.concat(['the DISPOSITION SUMMARY change of ReportCreator.js is not in it exactly once']) };
  return { code: text.slice(0, at) + text.slice(at + DISPOSITION_CREATOR_CHANGE.length), problems: later.problems };
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

module.exports = { codeWithoutOvernightPause, OVERNIGHT_PAUSE_NEW_FUNCTIONS, OVERNIGHT_PAUSE_CHANGED_FUNCTIONS, codeWithoutRevisionsCellCleanup, CELL_CLEANUP_NEW_FUNCTIONS, CELL_CLEANUP_CHANGED_FUNCTIONS, codeWithoutLocalListingAgent, LOCAL_LISTING_AGENT_NEW_FUNCTIONS, LOCAL_LISTING_AGENT_CHANGED_FUNCTIONS, codeWithoutSessionSectionFix, SESSION_SECTION_NEW_FUNCTIONS, SESSION_SECTION_CHANGED_FUNCTIONS, codeWithoutSessionTimesBuild, codeWithoutMainMeetingNumber, MAIN_MEETING_NUMBER_SECTION_START, MAIN_MEETING_NUMBER_SECTION_END, MAIN_MEETING_NUMBER_SECTION_FUNCTIONS, MAIN_MEETING_NUMBER_CHANGED_FUNCTIONS, codeWithoutDualRevisions, creatorWithoutDualRevisions, DUAL_REVISIONS_NEW_FUNCTIONS, DUAL_REVISIONS_CHANGED_FUNCTIONS, DUAL_REVISIONS_CREATOR_CHANGED_FUNCTIONS, codeWithoutLocalRevisions, LOCAL_REVISIONS_NEW_FUNCTIONS, LOCAL_REVISIONS_CHANGED_FUNCTIONS, codeWithoutSessionTimes, creatorWithoutSessionTimes, SESSION_TIMES_NEW_FUNCTIONS, SESSION_TIMES_CHANGED_FUNCTIONS, SESSION_TIMES_CREATOR_NEW_FUNCTIONS, SESSION_TIMES_CREATOR_CHANGED_FUNCTIONS, codeWithoutMeetingTimeZones, creatorWithoutMeetingTimeZones, MEETING_TIME_ZONES_SECTION_START, MEETING_TIME_ZONES_SECTION_END, MEETING_TIME_ZONES_SECTION_FUNCTIONS, codeWithoutLargeBuildCheckpoints, LARGE_BUILD_SECTION_START, LARGE_BUILD_SECTION_END, LARGE_BUILD_SECTION_FUNCTIONS, LARGE_BUILD_CHANGED_FUNCTIONS, LARGE_BUILD_DOCUMENT_ACCESS_COUNT, codeWithoutTdocListDownload, TDOC_LIST_DOWNLOAD_SECTION_START, TDOC_LIST_DOWNLOAD_SECTION_END, TDOC_LIST_DOWNLOAD_SECTION_FUNCTIONS, TDOC_LIST_DOWNLOAD_CHANGED_FUNCTIONS, codeWithoutAgendaReaders, AGENDA_READERS_SECTION_START, AGENDA_READERS_SECTION_END, AGENDA_READERS_SECTION_FUNCTIONS, AGENDA_READERS_CHANGED_FUNCTIONS, codeWithoutAgendaWordConversion,AGENDA_CONVERSION_SECTION_START, AGENDA_CONVERSION_SECTION_END, AGENDA_CONVERSION_SECTION_FUNCTIONS, AGENDA_CONVERSION_CHANGED_FUNCTIONS, codeWithoutCloseReportPageFallback,PAGE_FALLBACK_SECTION_START, PAGE_FALLBACK_SECTION_END, PAGE_FALLBACK_SECTION_FUNCTIONS, PAGE_FALLBACK_CHANGED_FUNCTIONS, codeWithoutReportLocalScope, REPORT_LOCAL_SCOPE_NEW_FUNCTIONS, REPORT_LOCAL_SCOPE_CHANGED_FUNCTIONS, codeWithoutListIdentity, IDENTITY_SECTION_START, IDENTITY_SECTION_END, IDENTITY_SECTION_FUNCTIONS, IDENTITY_CHANGED_FUNCTIONS, codeWithoutDispositionSummary, creatorWithoutDispositionSummary, DISPOSITION_SECTION_START, DISPOSITION_SECTION_END, DISPOSITION_SECTION_FUNCTIONS, DISPOSITION_CHANGED_FUNCTIONS, DISPOSITION_CREATOR_CHANGED_FUNCTIONS,
  codeBeforeStructureFix,codeWithoutConfigurationGuard, codeWithoutReleaseHeader, codeWithoutAbstractCacheChange, ABSTRACT_CACHE_CHANGED_FUNCTIONS, codeWithoutAttendanceEmptyTextFix, ATTENDANCE_EMPTY_TEXT_CHANGED_FUNCTIONS, creatorWithoutConfigurationGuard, GUARD_CREATOR_CHANGED_FUNCTIONS, SECTION_START, SECTION_END, SECTION_FUNCTIONS, CHANGED_FUNCTIONS,
  GUARD_SECTION_START, GUARD_SECTION_END, GUARD_SECTION_FUNCTIONS, GUARD_CHANGED_FUNCTIONS };
