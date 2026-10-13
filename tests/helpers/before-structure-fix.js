/**
 * Code.js as it was before the report-structure fix (after T-2026.10.12).
 *
 * Several test files hold a ledger of what changed in Code.js since a
 * release: which functions, and that nothing else did. The structure fix
 * changed seven released functions and added one section. Those ledgers are
 * about their own feature; they go on comparing with their release through
 * this view, which takes the structure fix out again -- every change named
 * here, each found exactly once. What the fix itself changed is pinned in
 * tests/report-structure.fix.test.js: this view is, byte for byte, the
 * Code.js of T-2026.10.12.
 *
 * codeBeforeStructureFix(code) -> { code, problems }: `problems` names every
 * change that was not found exactly once (then the view is not to be trusted
 * and the ledger that uses it fails).
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

function codeBeforeStructureFix(code) {
  let text = String(code).replace(/\r/g, '');
  const problems = [];
  const start = text.indexOf(SECTION_START);
  const end = text.indexOf(SECTION_END);
  if (start === -1 || end === -1 || end < start || text.indexOf(SECTION_START, start + 1) !== -1) {
    problems.push('the REPORT STRUCTURE section is not in Code.js exactly once');
  } else {
    text = text.slice(0, start) + text.slice(end + SECTION_END.length);
  }
  CHANGES.forEach(([now, before], i) => {
    const at = text.indexOf(now);
    if (at === -1 || text.indexOf(now, at + 1) !== -1) { problems.push('change ' + (i + 1) + ' is not in Code.js exactly once: ' + JSON.stringify(now.slice(0, 60))); return; }
    text = text.slice(0, at) + before + text.slice(at + now.length);
  });
  return { code: text, problems: problems };
}

module.exports = { codeBeforeStructureFix, SECTION_START, SECTION_END, SECTION_FUNCTIONS, CHANGED_FUNCTIONS };
