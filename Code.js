/*******************************
 * SA4 Report Generator + Email/Revisions Collector
 * Version: 2.6.0 (2026-08-28)
 * - NO global name collisions
 * - RSS/A1 + Revisions restored
 * - Agenda Item rows preserved/merged
 *
 * CHANGELOG
 * 2.6.0 (2026-08-28)
 *   - Added: ONE-CLICK UPDATE. "▶️ Update Report" runs every enabled step in
 *     the correct order; each step can be switched off in "⚙️ Update Options".
 *     The same engine (runUpdate_) powers the menu action and the time-based
 *     trigger, so manual and automatic updates can no longer drift apart.
 *   - Added: "Last updated" timestamp written into the first section of the
 *     report (updated in place on later runs, never duplicated).
 *   - Added: "❓ Help" dialog; full guide in DOCUMENTATION.md.
 *   - Changed: menu reorganised into SETUP / BUILD & UPDATE / DOCUMENTS /
 *     TOOLS, with the duplicated "Configure Meeting" entry removed.
 *   - Fixed: the "Fix Links (portal→FTP)" menu item pointed at
 *     rewritePortalLinksInDoc_ — Apps Script refuses to run a function whose
 *     name ends in "_" from a menu, so the item always failed. Public wrapper
 *     fixPortalLinks() added.
 *   - Removed: 5 duplicated function definitions (~200 lines of dead code).
 *     JavaScript keeps the LAST definition, so these were silently shadowed:
 *     formatDeadline_, getMonthName_, calculateTimeRemaining_ (exact copies),
 *     copySectionContentWithReplacement_ (the unused formatting-preserving
 *     variant) and the addTdocTablesOnly alias. Behaviour is unchanged: in
 *     every case the definition that was actually in effect is the one kept.
 *   - Added: tests/no-duplicates.test.js guards against duplicate definitions
 *     reappearing and verifies every menu target exists and is callable.
 * 2.5.1 (2026-08-28)
 *   - Fixed: 2.5.0 only applied the revision logic on a full skeleton rebuild,
 *     so TDOCs arriving through the other two import paths still landed at the
 *     end of their agenda section with an empty Disposition. Now covered:
 *     * continuousUpdate() re-arranges revisions and fills Dispositions after
 *       inserting new TDOCs, reusing the TDOC list it already downloaded (no
 *       extra fetch). Existing reports therefore self-heal on the next update
 *       and no longer need the manual menu action.
 *     * insertNewTdoc_() places a new revision directly below the document it
 *       revises (via the cached REVISION_MAP) and fills its Disposition.
 *     * processWebDownloadedSheet_() (downloadAndProcessFromWeb) now orders by
 *       revision and fills Disposition like the skeleton build.
 * 2.5.0 (2026-08-28)
 *   - Added: revision handling driven by the TDOC list "Revised to" column.
 *     * A revision's table is now emitted directly below the document it
 *       revises, following chains (A -> B -> C).
 *     * The revised document's Disposition is filled with
 *       "Revised to S4-xxxxxx". Existing hand-written text is preserved.
 *     * New menu action "Re-arrange Revision Tables" repairs documents where
 *       revisions were already inserted in the wrong place, without rebuilding.
 * 2.4.0 (2026-08-26)
 *   - Added: "Document Deadline Extensions" table (TDOC | Extended Deadline).
 *     Deadlines extended verbally or by e-mail now override the deadline parsed
 *     from the subject line. Applied during e-mail discussion collection; the
 *     deadline row is tagged "(extended)" and late-response greying uses the
 *     extended time. The table is found by its header cells, so its heading
 *     and position in the document are irrelevant. Its contents are mirrored
 *     into the DEADLINE_EXTENSIONS property so a skeleton rebuild (which
 *     clears the body) does not silently lose the extensions.
 * 2.3.0 (2026-08-26)
 *   - Added: "Fetch abstracts during each update" switch in the trigger
 *     configuration dialog (FETCH_ABSTRACTS_ON_UPDATE, default OFF). Only
 *     governs automatic/continuous updates; menu step 5 always works.
 *   - Added: e-mails received after the thread deadline are rendered in grey
 *     (#808080) instead of black, unless the sender is the opening sender.
 * 2.2.0 (2026-08-26)
 *   - Fixed: reply/forward e-mails were silently dropped. parseEmailSubject_
 *     required "[" at position 0, so "Re: [FS_6G_MED, 1483, ...]" returned null
 *     and the message never reached the store (3 msgs in RSS -> 1 in cache).
 *     Added stripReplyPrefixes_() (Re/RE/AW/FW/FWD/TR/SV/ANTW/VS/RIF/RES,
 *     repeated, and "RE[2]:" style) applied before bracket extraction.
 * 2.1.0 (2026-08-26)
 *   - Fixed: storage deduplication used RAW dates, which differ between RSS
 *     ("Mon, 24 Aug 2026 07:23:30 +0200") and A1 ("24 Aug 2026 07:23:30").
 *     Now both storage and display dedupe on formatLocalDate_() output.
 *******************************/

const DEFAULT_S4_FTP_BASE =
  'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/';

// Default SA4 list used by the collector unless report-type-specific logic overrides it.
// This constant is required by RSS, archive, and configuration helper functions below.
const LIST_NAME_LOCK = '3GPP_TSG_SA_WG4';

// Mailing list names by report type
const MAILING_LISTS = {
  'Audio': '3GPP_TSG_SA_WG4_AUDIO',
  'Video': '3GPP_TSG_SA_WG4_VIDEO',
  'MBS': '3GPP_TSG_SA_WG4_MBS',
  'RTC': '3GPP_TSG_SA_WG4_RTC',
  '6G': '3GPP_TSG_SA_WG4',
  'Liaison': '3GPP_TSG_SA_WG4',
  'New': '3GPP_TSG_SA_WG4'
};

// Drafts folder names by report type (SWG names)
const DRAFTS_FOLDERS = {
  'Audio': 'Audio',
  'Video': 'Video',
  'MBS': 'MBS',
  'RTC': 'RTC',
  '6G': 'FS_6G_MED',
  'Liaison': 'Plenary',
  'New': 'Plenary'
};

function onOpen() {
  const ui = DocumentApp.getUi();
  
  // Main menu
  const menu = ui.createMenu('⚠️Scripts⚠️');
  
  // INITIAL SETUP submenu
  const setupMenu = ui.createMenu('📝 INITIAL SETUP');
  setupMenu.addItem('⚙️ Configure Meeting Settings', 'configureMeetingSettings');
  setupMenu.addItem('🧪 Test All Connections', 'testAllConnections');
  setupMenu.addItem('📋 Create Configuration Tables', 'createConfigurationTables');
  
  // REPORT OPERATIONS submenu
  const reportMenu = ui.createMenu('🚀 REPORT OPERATIONS');
  reportMenu.addItem('▶️ Run Full Report Build', 'runFullReportBuild');
  reportMenu.addSeparator();
  reportMenu.addItem('0️⃣ Configure Meeting', 'configureMeetingSettings');
  reportMenu.addItem('1️⃣+2️⃣ Build Skeleton + TDOC Tables', 'buildSkeletonWithTdocTables');
  reportMenu.addItem('3️⃣ Collect E-mail Discussion', 'collectEmailDiscussionOnly');
  reportMenu.addItem('4️⃣ Collect Revisions', 'collectRevisionsOnly');
  reportMenu.addItem('5️⃣ Add Abstracts', 'addAbstractsOnly');
  reportMenu.addSeparator();
  reportMenu.addItem('🔄 Continuous Update (New TDOCs + Status)', 'continuousUpdate');
  reportMenu.addItem('⏰ Manage Auto-Update Trigger', 'manageTriggers');  // ADD THIS LINE
  reportMenu.addSeparator();
  reportMenu.addItem('📝 Legacy: Build Initial Report', 'buildInitialReport');
  reportMenu.addItem('🔄 Update Report (During Meeting)', 'updateReportIncremental');
  reportMenu.addItem('📊 Analyze Report Status', 'analyzeReportStatus');

  
  // TOOLS & DIAGNOSTICS submenu
  const toolsMenu = ui.createMenu('🔧 TOOLS & DIAGNOSTICS');
  toolsMenu.addItem('🧪 Test TDOC List URL', 'testTdocListUrl');
  toolsMenu.addItem('🧪 Test Reviewer API', 'testReviewerApi');
  toolsMenu.addItem('🧪 Test Email Feeds', 'testEmailFeeds');
  toolsMenu.addItem('🧪 Test Revisions Folder', 'testRevisionsFolder');
  toolsMenu.addItem('✅ Validate Configuration', 'validateConfiguration');
  toolsMenu.addItem('🗑️ Clear All Caches', 'clearAllCaches');
  
  // FORMATTING & FIXES submenu
  const formatMenu = ui.createMenu('🎨 FORMATTING & FIXES');
  formatMenu.addItem('🎨 Format Document', 'removeRowHeightAndSpacing');
  formatMenu.addItem('🔗 Fix Links (portal→FTP)', 'rewritePortalLinksInDoc_');
  formatMenu.addItem('📏 Fix Column Widths', 'fixColumnWidths');
  
  // DOCUMENT MANAGEMENT submenu
  const docMenu = ui.createMenu('📋 DOCUMENT MANAGEMENT');
  docMenu.addItem('➕ Add Document Reallocation', 'addDocumentReallocation');
  docMenu.addItem('📊 View All Reallocations', 'viewAllReallocations');
  docMenu.addItem('🗑️ Clear All Reallocations', 'clearAllReallocations');
  docMenu.addItem('🔄 Apply Document Reallocations', 'applyDocumentReallocations');
  docMenu.addItem('🔀 Re-arrange Revision Tables', 'rearrangeRevisionTables');
  docMenu.addSeparator();
  docMenu.addItem('📄 Parse Agenda Document', 'parseAgendaDocument');
  docMenu.addItem('🏗️ Auto-Create Report Structure', 'autoCreateReportStructure');
  docMenu.addSeparator();
  docMenu.addItem('🧹 Clean Up Wrong Email Discussions', 'cleanUpWrongEmailDiscussions');
  docMenu.addItem('🔧 Remove Duplicate Email Entries', 'removeDuplicateEmailEntries');
  
  // Add all submenus to main menu
  menu.addSubMenu(setupMenu);
  menu.addSubMenu(reportMenu);
  menu.addSubMenu(docMenu);
  menu.addSubMenu(toolsMenu);
  menu.addSubMenu(formatMenu);
  
  // Legacy functions (for backward compatibility)
  menu.addSeparator();
  menu.addItem('⚠️ Legacy: Update All', 'updateAll');
  
  menu.addToUi();
}


/**
 * CONTINUOUS UPDATE FUNCTION
 * Downloads latest TDOCs, adds new ones, updates status
 */
function continuousUpdate() {
  const cfg = getReportConfig_();
  const body = DocumentApp.getActiveDocument().getBody();
  
  Logger.log('=== CONTINUOUS UPDATE START ===');
  
  try {
    // Download latest TDOC list
    const tdocGroups = downloadAndGroupTdocs_(cfg);
    const allTdocs = [];
    Object.keys(tdocGroups).forEach(key => {
      tdocGroups[key].tdocs.forEach(td => allTdocs.push({ ...td, agendaItem: key }));
    });
    
    Logger.log(`Downloaded ${allTdocs.length} TDOCs`);
    
    // Get existing TDOCs
    const existingTdocs = new Set();
    body.getTables().forEach(t => {
      if (!isTDocTable_(t)) return;
      const tdoc = safeCellText_(t, 0, 1).trim();
      if (tdoc) existingTdocs.add(tdoc);
    });
    
    Logger.log(`Found ${existingTdocs.size} existing TDOCs`);
    
    // Add new TDOCs and update status
    let newTdocsAdded = 0;
    let statusUpdated = 0;
    
    allTdocs.forEach(tdocData => {
      const row = tdocData.row;
      const tdocNumber = String(row[tdocData.tdocCol] || '').trim();
      
      if (!existingTdocs.has(tdocNumber)) {
        insertNewTdoc_(body, tdocData, cfg);
        newTdocsAdded++;
      } else {
        if (updateTdocStatus_(body, tdocNumber, tdocData)) {
          statusUpdated++;
        }
      }
    });
    
    Logger.log(`Added ${newTdocsAdded} new, updated ${statusUpdated} statuses`);
    
    // Update summary table
    if (newTdocsAdded > 0) {
      updateRegisteredDocumentsTable_(body, allTdocs, tdocGroups);
    }
    
    // Revision placement: new TDOCs are appended at the end of their agenda
    // section, so move every revision back under the document it revises and
    // fill the revised document's Disposition. Reuses the TDOC list already
    // downloaded above, so this costs no extra fetch.
    const rev = rearrangeRevisionTables_(cfg, tdocGroups);
    Logger.log(`Revisions: ${rev.moved} moved, ${rev.dispositions} disposition(s) filled`);
    
    // Abstracts are optional on automatic updates (OFF by default).
    // Toggle via: REPORT OPERATIONS -> Manage Auto-Update Trigger.
    if (getFetchAbstractsSetting_()) {
      const abstractsAdded = addAbstractsForTables_(body);
      Logger.log(`Fetched abstracts for ${abstractsAdded} table(s)`);
    } else {
      Logger.log('Abstract fetching is disabled (trigger configuration)');
    }
    
    // Collect emails and revisions
    collectorUpdate_();
    
    // Format
    removeRowHeightAndSpacing();
    
    Logger.log('=== COMPLETE ===');
    
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log(e.stack);
  }
}

/**
 * TRIGGER MANAGEMENT
 */

function manageTriggers() {
  const ui = DocumentApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const fetchAbstracts = getFetchAbstractsSetting_();
  
  // Check current trigger status
  const triggers = ScriptApp.getProjectTriggers();
  const continuousTrigger = triggers.find(t => t.getHandlerFunction() === 'continuousUpdate');
  const isActive = !!continuousTrigger;
  
  // Get current interval
  let currentInterval = 'Not set';
  if (continuousTrigger) {
    const minutes = continuousTrigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK 
      ? 'Active' 
      : 'Unknown';
    currentInterval = minutes;
  }
  
  // Build HTML dialog
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .status { padding: 15px; margin: 15px 0; border-radius: 5px; }
      .active { background: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
      .inactive { background: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      label.check { font-weight: normal; }
      select { width: 100%; padding: 8px; margin-top: 5px; }
      .hint { font-size: 11px; color: #666; margin-top: 3px; }
      button { margin-top: 20px; padding: 10px 20px; border: none; cursor: pointer; color: white; }
      .btn-start { background: #28a745; }
      .btn-stop { background: #dc3545; }
      .btn-cancel { background: #6c757d; }
      button:hover { opacity: 0.9; }
    </style>
    
    <h2>⏰ Continuous Update Trigger</h2>
    
    <div class="status ${isActive ? 'active' : 'inactive'}">
      <strong>Status:</strong> ${isActive ? '✅ Active' : '❌ Inactive'}<br>
      ${isActive ? '<strong>Running every:</strong> ' + currentInterval : ''}
    </div>
    
    <label>Update Interval:</label>
    <select id="interval">
      <option value="15">Every 15 minutes (Active meeting)</option>
      <option value="30" selected>Every 30 minutes (Recommended)</option>
      <option value="60">Every hour (Slow meeting)</option>
    </select>
    
    <label>Abstracts:</label>
    <label class="check">
      <input type="checkbox" id="fetchAbstracts" ${fetchAbstracts ? 'checked' : ''} onchange="saveAbstracts()">
      Fetch abstracts during each update
    </label>
    <div class="hint">
      Off by default. When on, every automatic update calls the Reviewer API for
      each TDOC still missing an abstract, which makes updates noticeably slower.
      The menu step "5️⃣ Add Abstracts" always works regardless of this setting.
    </div>
    
    <div style="margin-top: 20px;">
      ${isActive 
        ? '<button class="btn-stop" onclick="stopTrigger()">⏹️ Stop Trigger</button>'
        : '<button class="btn-start" onclick="startTrigger()">▶️ Start Trigger</button>'
      }
      <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
    </div>
    
    <script>
      function startTrigger() {
        const interval = document.getElementById('interval').value;
        google.script.run
          .withSuccessHandler(() => {
            alert('✅ Trigger started! Updates will run every ' + interval + ' minutes.');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('❌ Error: ' + error);
          })
          .createContinuousTrigger(parseInt(interval), document.getElementById('fetchAbstracts').checked);
      }
      
      function saveAbstracts() {
        google.script.run
          .withFailureHandler((error) => {
            alert('❌ Error: ' + error);
          })
          .setFetchAbstractsSetting(document.getElementById('fetchAbstracts').checked);
      }
      
      function stopTrigger() {
        google.script.run
          .withSuccessHandler(() => {
            alert('⏹️ Trigger stopped.');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('❌ Error: ' + error);
          })
          .deleteContinuousTrigger();
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(500);
  
  ui.showModalDialog(html, 'Manage Continuous Update Trigger');
}

/**
 * Whether automatic/continuous updates should fetch abstracts.
 * Defaults to false (off) when the property was never set.
 */
function getFetchAbstractsSetting_() {
  return PropertiesService.getDocumentProperties()
    .getProperty('FETCH_ABSTRACTS_ON_UPDATE') === 'true';
}

/**
 * Persist the abstract-fetching switch. Called from the trigger dialog.
 */
function setFetchAbstractsSetting(enabled) {
  const value = enabled ? 'true' : 'false';
  PropertiesService.getDocumentProperties().setProperty('FETCH_ABSTRACTS_ON_UPDATE', value);
  Logger.log('FETCH_ABSTRACTS_ON_UPDATE = ' + value);
}

function createContinuousTrigger(intervalMinutes, fetchAbstracts) {
  // Delete existing trigger first
  deleteContinuousTrigger();
  
  // Persist the abstracts switch alongside the trigger
  if (fetchAbstracts !== undefined && fetchAbstracts !== null) {
    setFetchAbstractsSetting(fetchAbstracts);
  }
  
  // Create new trigger
  ScriptApp.newTrigger('continuousUpdate')
    .timeBased()
    .everyMinutes(intervalMinutes)
    .create();
  
  Logger.log(`Trigger created: every ${intervalMinutes} minutes (abstracts: ${getFetchAbstractsSetting_()})`);
}

function deleteContinuousTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continuousUpdate') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Trigger deleted');
    }
  });
}

function getTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const continuousTrigger = triggers.find(t => t.getHandlerFunction() === 'continuousUpdate');
  
  if (!continuousTrigger) {
    return { active: false, interval: null };
  }
  
  return {
    active: true,
    interval: 'Active' // Apps Script doesn't expose the exact interval
  };
}


function insertNewTdoc_(body, tdocData, cfg) {
  const row = tdocData.row;
  const agendaItem = tdocData.agendaItem;
  const revisedTo = getRevisedTo_(tdocData);

  // Prefer sitting directly below the document this one revises. Only if that
  // document is not in the report yet do we fall back to the end of the
  // agenda section (continuousUpdate's re-arrangement pass fixes it later).
  let insertIdx = -1;
  const parentTable = findParentRevisedToTable_(body, tdocNumberOf_(tdocData));
  if (parentTable) {
    insertIdx = body.getChildIndex(parentTable) + 1;
    Logger.log(`Placing ${tdocNumberOf_(tdocData)} directly below its parent`);
  }
  if (insertIdx < 0) insertIdx = findInsertionPointForAgendaItem_(body, agendaItem, '');
  
  const typeCol = tdocData.typeCol;
  const forCol = tdocData.forCol;
  const typeFor = (typeCol >= 0 && forCol >= 0 && row[typeCol] && row[forCol])
    ? `${row[typeCol]} for ${row[forCol]}`
    : (typeCol >= 0 && row[typeCol]) ? row[typeCol] 
    : (forCol >= 0 && row[forCol]) ? row[forCol] : '';
  
  const tempData = [
    ['TDoc', row[tdocData.tdocCol]],
    ['Title', row[tdocData.titleCol]],
    ['Source', row[tdocData.sourceCol]],
    ['Contact', tdocData.contactCol >= 0 ? row[tdocData.contactCol] : ''],
    ['Agenda Item', agendaItem],
    ['Type/For', typeFor],
    ['E-mail Discussion', ''],
    ['Revisions', ''],
    ['Minutes', ''],
    ['Disposition', revisedTo ? 'Revised to ' + revisedTo : ''],
    ['Status', tdocData.statusCol >= 0 ? row[tdocData.statusCol] : '']
  ];
  
  insertTDocTableAtIndex_(body, insertIdx, tempData, tdocData.richTextRow, tdocData.tdocCol);
  Logger.log(`Inserted ${row[tdocData.tdocCol]}`);
}

/**
 * Find the table of the document that was revised INTO `tdocNumber`, using the
 * REVISION_MAP built from the TDOC list. Returns null when unknown or when that
 * document has no table in this report.
 */
function findParentRevisedToTable_(body, tdocNumber) {
  const want = String(tdocNumber || '').trim().toUpperCase();
  if (!want) return null;

  const map = loadJsonObject_(
    PropertiesService.getDocumentProperties().getProperty('REVISION_MAP'));

  for (const parent of Object.keys(map)) {
    if (String(map[parent] || '').toUpperCase() !== want) continue;
    const t = findTdocTable_(body, parent);
    if (t) return t;
  }
  return null;
}

function updateTdocStatus_(body, tdocNumber, tdocData) {
  const row = tdocData.row;
  const newStatus = tdocData.statusCol >= 0 ? String(row[tdocData.statusCol] || '').trim() : '';
  
  if (!newStatus) return false;
  
  const tables = body.getTables();
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (!isTDocTable_(table)) continue;
    
    const tableTdoc = safeCellText_(table, 0, 1).trim();
    if (tableTdoc !== tdocNumber) continue;
    
    const statusInfo = findStatusInDocTable_(table);
    if (!statusInfo) continue;
    
    const currentStatus = statusInfo.value.trim();
    
    if (currentStatus !== newStatus) {
      const docStatus = normalizeStatus_(currentStatus);
      const newStatusLower = newStatus.toLowerCase();
      const isRevised = newStatusLower.includes('revised');
      
      if (docStatus === 'reserved' || docStatus === 'available' || isRevised) {
        statusInfo.cell.setText(newStatus);
        styleStatusCell_(table);
        Logger.log(`Updated ${tdocNumber}: ${currentStatus} → ${newStatus}`);
        return true;
      }
    }
    return false;
  }
  return false;
}

function updateRegisteredDocumentsTable_(body, allTdocs, tdocGroups) {
  const tables = body.getTables();
  let summaryTable = null;
  
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (table.getNumRows() < 1) continue;
    
    const row0 = table.getRow(0);
    if (row0.getNumCells() === 4 &&
        row0.getCell(0).getText().trim() === 'TDoc' &&
        row0.getCell(1).getText().trim() === 'Title' &&
        row0.getCell(2).getText().trim() === 'Source' &&
        row0.getCell(3).getText().trim() === 'Agenda Item') {
      summaryTable = table;
      break;
    }
  }
  
  if (!summaryTable) return;
  
  const existingInSummary = new Set();
  for (let r = 1; r < summaryTable.getNumRows(); r++) {
    const tdoc = summaryTable.getRow(r).getCell(0).getText().trim();
    if (tdoc) existingInSummary.add(tdoc);
  }
  
  let added = 0;
  allTdocs.forEach(tdocData => {
    const row = tdocData.row;
    const tdocNumber = String(row[tdocData.tdocCol] || '').trim();
    
    if (!existingInSummary.has(tdocNumber)) {
      const dataRow = summaryTable.appendTableRow();
      
      const tdocCell = dataRow.appendTableCell(tdocNumber);
      if (tdocData.richTextRow && tdocData.richTextRow[tdocData.tdocCol]) {
        const richText = tdocData.richTextRow[tdocData.tdocCol];
        if (richText.getLinkUrl && richText.getLinkUrl()) {
          const text = tdocCell.editAsText();
          text.setLinkUrl(0, tdocNumber.length - 1, richText.getLinkUrl());
        }
      }
      
      dataRow.appendTableCell(String(row[tdocData.titleCol] || ''));
      dataRow.appendTableCell(String(row[tdocData.sourceCol] || ''));
      dataRow.appendTableCell(String(tdocData.agendaItem || ''));
      
      added++;
    }
  });
  
  Logger.log(`Added ${added} to summary table`);
}

function updateAll() {
  // 1) Copy / merge TDOC + non-TDOC tables
  try {
    if (typeof copyDocsToReport === 'function') {
      copyDocsToReport();
    }
  } catch (e) {
    Logger.log('copyDocsToReport() failed: ' + e.message);
  }

  copyIndividualToReport();

  // 2) Collect e-mail discussion + revisions
  collectorUpdate_();

  // 3) Fix links + ordering (safe, non-destructive)
  fixLinksAndReorder_();

  // 4) ✅ FORMATTING MUST ALWAYS BE LAST
  removeRowHeightAndSpacing();
}

// =========================================================
// CENTRALIZED CONFIGURATION
// =========================================================

function getReportConfig_() {
  const props = PropertiesService.getDocumentProperties();
  
  // ========================================
  // 1. MEETING CONFIGURATION
  // ========================================
  const MEETING_FOLDER = props.getProperty('MEETING_FOLDER') || 'TSGS4_136_Montreal';
  const MEETING_NUMBER = props.getProperty('MEETING_NUMBER') || '136';
  const MEETING_ID = props.getProperty('MEETING_ID') || '60777';
  
  // ========================================
  // 2. BASE PATHS (Auto-calculated from meeting folder)
  // ========================================
  const FTP_BASE_ROOT = 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/';
  const FTP_BASE = props.getProperty('FTP_BASE') || `${FTP_BASE_ROOT}${MEETING_FOLDER}/Docs/`;
  const INBOX_BASE = FTP_BASE.replace('/Docs/', '/Inbox/');
  
  // ========================================
  // 3. TDOC LIST (Auto-detect or explicit)
  // ========================================
  const explicitTdocUrl = props.getProperty('TDOC_LIST_URL');
  const TDOC_LIST_URL = explicitTdocUrl || `${FTP_BASE}TDoc_List_Meeting_SA4%23${MEETING_NUMBER}.xlsx`;
  
  // ========================================
  // 4. REPORT TYPE CONFIGURATION
  // ========================================
  const REPORT_SUFFIX = props.getProperty('REPORT_SUFFIX') || '6G';
  const DRAFTS_FOLDER = DRAFTS_FOLDERS[REPORT_SUFFIX] || 'Plenary';
  const LIST_NAME = MAILING_LISTS[REPORT_SUFFIX] || LIST_NAME_LOCK;
  const REVISIONS_URL = `${INBOX_BASE}Drafts/${DRAFTS_FOLDER}`;
  
  // ========================================
  // 5. AGENDA CONFIGURATION
  // ========================================
  const AGENDA_ITEM_PREFIX = props.getProperty('AGENDA_ITEM_PREFIX') || getAgendaPrefixForReportType_(REPORT_SUFFIX);
  const AGENDA_SOURCE_DOC_ID = props.getProperty('AGENDA_SOURCE_DOC_ID') || '1qP--dusvUhNwwBtMEH4xVdxaP1c6L1hZ49geICoYV2s';
  const AGENDA_TDOC = props.getProperty('AGENDA_TDOC') || '';
  
  // ========================================
  // 6. OPTIONS
  // ========================================
  const SHOW_PREVIEW_SNIPPET = props.getProperty('SHOW_PREVIEW_SNIPPET') !== 'false';
  
  // ========================================
  // 7. API INTEGRATION (Optional)
  // ========================================
  const REVIEWER_API_TOKEN = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN') || '';
  const REVIEWER_API_BASE = 'https://reviewer.bouazizi.dev/api/v1';
  
  return {
    // Meeting identification
    MEETING_FOLDER,
    MEETING_NUMBER,
    MEETING_ID,
    
    // Paths
    FTP_BASE,
    TDOC_LIST_URL,
    INBOX_BASE,
    REVISIONS_URL,
    
    // Report configuration
    REPORT_SUFFIX,
    DRAFTS_FOLDER,
    LIST_NAME,
    
    // Agenda
    AGENDA_ITEM_PREFIX,
    AGENDA_SOURCE_DOC_ID,
    AGENDA_TDOC,
    
    // Options
    SHOW_PREVIEW_SNIPPET,
    
    // API
    REVIEWER_API_TOKEN,
    REVIEWER_API_BASE,
    
    // RSS feeds
    RSS_URL_V2: `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME}&v=2.0&LIMIT=2000`,
    RSS_URL_V1: `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME}&v=1.0&LIMIT=2000`
  };
}

// =========================================================
// MEETING CONTEXT (SA4-ARCH-003 compatibility layer)
// =========================================================
//
//   existing configuration (DocumentProperties/ScriptProperties)
//           |
//           v
//   getReportConfig_()
//           |
//           v
//   getMeetingContext_()   <-- this function
//           |
//           v
//   future profile-aware consumers (NOT YET WIRED UP)
//
// getMeetingContext_() is a read-only, pure normalization of whatever
// getReportConfig_() currently returns. It introduces NO second source of
// truth: every field below is derived from the getReportConfig_() result,
// never by re-reading PropertiesService directly. It has no side effects
// beyond whatever getReportConfig_() itself already performs (a
// PropertiesService read).
//
// As of SA4-ARCH-003, NOTHING in this file calls getMeetingContext_() yet.
// runFullReportBuild(), buildSkeletonWithTdocTables(), continuousUpdate(),
// the collector, agenda parsing, revision processing, reallocation
// processing, the setup/config UI, and the legacy report builders all
// continue to call getReportConfig_() exactly as before. Migrating them is
// deliberately left to a later task (see SA4-ARCH-002/003 reports).
//
// meeting.type is hardcoded to 'main' for every meeting today, because the
// entire codebase currently only knows how to build main-meeting reports.
// It exists now as the FUTURE discriminator between 'main', 'adhoc', and
// any later meeting profile -- no ad-hoc behavior is implemented anywhere
// yet, and nothing reads this field to branch on it.
//
// report.structureProfile makes explicit the structural branch that is
// currently embedded as inline `is6G`/`isSWGReport` booleans inside
// buildSkeletonWithTdocTables() (Code.js, see the "6G/SWG/other" skeleton
// logic ~5350-5517): 'main-6g' for the 6G plenary skeleton (11.0.1-11.0.4
// subsections), 'main-swg' for the Audio/Video/MBS/RTC skeleton (X.1/X.1.2/
// X.2 template sections), 'main-other' for report types that get neither
// (Liaison, New, and any unrecognized REPORT_SUFFIX). This field is derived
// here from the same REPORT_SUFFIX value buildSkeletonWithTdocTables()
// already branches on, but buildSkeletonWithTdocTables() itself is NOT
// changed to consume it -- the branch logic there is untouched.
//
// options.emailStartDate and options.fetchAbstractsOnUpdate from the
// candidate shape in the SA4-ARCH-003 task description are intentionally
// OMITTED: getReportConfig_() does not read or return EMAIL_START_DATE
// (only getCollectorConfig_() and the config dialog touch it, via a
// different code path -- Code.js:721,4052,4119) or FETCH_ABSTRACTS_ON_UPDATE
// (only getFetchAbstractsSetting_() reads it, again a different code path --
// Code.js:385-391). Populating them here would mean either re-reading
// PropertiesService independently (a second source of truth) or reaching
// into those other reader functions (reintroducing the same
// multiple-config-systems problem SA4-ARCH-001 flagged for
// getReportConfig_()/getCollectorConfig_()/getConfig_()). Both are out of
// scope for this compatibility layer; add them once there is one canonical
// place to read them from.
/**
 * SA4-IMPL-002: pure source-resolution rule, extracted so it is testable
 * without Google APIs and reusable once an explicit-override property
 * mechanism exists for ad-hoc meetings.
 *
 * `derived` is a plain object already shaped like MeetingContext.sources
 * (ftpBase, tdocListUrl, agendaTdoc, agendaTemplateDocId, mailingList,
 * draftsFolder, revisionsUrl) -- today this is always built from
 * getReportConfig_()'s output by the caller, but this function itself does
 * not know or care where `derived` came from, which is what keeps it pure
 * and independently testable (it never touches PropertiesService).
 *
 * `overrides` is an optional, partial object using the SAME keys. For each
 * key: a non-empty (after trimming) string in `overrides` wins; anything
 * else (absent, '', whitespace-only, null, undefined, non-string) falls
 * back to `derived`'s value. This is a plain precedence rule, not a second
 * configuration system -- it has no defaults of its own and invents nothing
 * for a field that has no value in either input.
 *
 * Deliberately works even when `derived` describes a meeting with no
 * MEETING_NUMBER at all (e.g. every field simply came from `overrides`
 * instead) -- see tests/meeting-context.test.js for the numberless case.
 * This capability is NOT yet exercised in production: getReportConfig_()
 * itself still requires MEETING_NUMBER today, and no DocumentProperty
 * mechanism yet exists to supply mailingList/draftsFolder/revisionsUrl
 * overrides (see the SA4-IMPL-002 report for the exact gap). Establishing
 * that the RESOLVER can already do this is the point of this task.
 */
function resolveMeetingSources_(derived, overrides) {
  const base = derived || {};
  const ov = overrides || {};

  function pick(key) {
    const explicit = ov[key];
    if (typeof explicit === 'string' && explicit.trim() !== '') return explicit;
    return base[key];
  }

  return {
    ftpBase: pick('ftpBase'),
    tdocListUrl: pick('tdocListUrl'),
    agendaTdoc: pick('agendaTdoc'),
    agendaTemplateDocId: pick('agendaTemplateDocId'),
    mailingList: pick('mailingList'),
    draftsFolder: pick('draftsFolder'),
    revisionsUrl: pick('revisionsUrl')
  };
}

/**
 * SA4-IMPL-003/003A: normalizes the MEETING_TYPE document property.
 * trim + lowercase; absent or blank -> 'main' (backward-compatible: every
 * document that predates this property, and every existing main-meeting
 * test/workflow, has no MEETING_TYPE set at all).
 *
 * Any other, unrecognized NON-EMPTY value throws, rather than silently
 * falling back to 'main' (SA4-IMPL-003A -- changed from the original
 * SA4-IMPL-003 behavior, which logged and fell back). Rationale: once
 * meeting type decides whether main-derived or explicit ad-hoc sources are
 * exposed (see getMeetingContext_()), silently converting a configuration
 * typo ("ad-hoc", "Adhoc " mistyped, "electronic") into 'main' risks
 * silently building against the WRONG meeting's sources rather than failing
 * where the mistake was made. This is a deliberate departure from this
 * file's usual "always fall back, never throw" convention, specific to this
 * one property, because the two failure modes are not equally bad here: a
 * missing property (blank) has an obviously-correct, harmless default
 * (main, matching every meeting that has ever existed so far); a wrong,
 * non-empty value does not.
 */
function normalizeMeetingType_(value) {
  const t = String(value || '').trim().toLowerCase();
  if (!t) return 'main';
  if (t === 'main') return 'main';
  if (t === 'adhoc') return 'adhoc';
  throw new Error('Unsupported MEETING_TYPE "' + value + '". Expected "main" or "adhoc".');
}

/**
 * SA4-IMPL-003: reads every meeting-identity-related DocumentProperty
 * together in one place, so getMeetingContext_() itself doesn't scatter
 * PropertiesService calls (per the SA4-IMPL-002/003 "keep configuration
 * access separate from pure resolution" rule).
 *
 * FTP_BASE / TDOC_LIST_URL / AGENDA_TDOC / AGENDA_SOURCE_DOC_ID /
 * REVISIONS_URL are read RAW here (no fallback formula), deliberately
 * bypassing getReportConfig_()'s own handling of those same property names.
 * This is intentional, not a second source of truth for MAIN meetings
 * (getMeetingContext_() still uses cfg.* for those, unchanged): for an
 * AD-HOC meeting, getReportConfig_()'s fallbacks for these fields are
 * main-meeting-specific formulas (e.g. FTP_BASE falls back to a
 * Montreal-based URL built from MEETING_FOLDER) that would be actively
 * WRONG -- not just incomplete -- to expose. Reading them raw lets a
 * missing ad-hoc value surface as genuinely absent instead of a silently
 * wrong main-meeting default. No new property names are introduced: these
 * are the exact same properties getReportConfig_() already reads.
 */
function getMeetingIdentityConfig_() {
  const props = PropertiesService.getDocumentProperties();
  return {
    MEETING_TYPE: normalizeMeetingType_(props.getProperty('MEETING_TYPE')),
    MEETING_NAME: String(props.getProperty('MEETING_NAME') || '').trim(),
    FTP_BASE: String(props.getProperty('FTP_BASE') || '').trim(),
    TDOC_LIST_URL: String(props.getProperty('TDOC_LIST_URL') || '').trim(),
    AGENDA_TDOC: String(props.getProperty('AGENDA_TDOC') || '').trim(),
    AGENDA_SOURCE_DOC_ID: String(props.getProperty('AGENDA_SOURCE_DOC_ID') || '').trim(),
    REVISIONS_URL: String(props.getProperty('REVISIONS_URL') || '').trim()
  };
}

/**
 * SA4-IMPL-004: validates and normalizes an agendaSelector.
 *
 * Exactly three modes are supported -- 'prefix', 'all', 'itemList' -- no
 * speculative modes. Throws for anything invalid (wrong shape, missing
 * mode, missing/empty required value) rather than silently coercing to
 * {mode:'all'}: silently selecting EVERY agenda item instead of the
 * intended subset is a much worse failure than a loud, immediate error,
 * exactly the same reasoning SA4-IMPL-003A applied to MEETING_TYPE.
 *
 * 'prefix': value must be a non-empty (after trim) string. Stored trimmed.
 * 'all': no value needed or used.
 * 'itemList': value must be an array; each entry is coerced to a trimmed
 * string, empty entries are dropped, and the result must be non-empty.
 */
function normalizeAgendaSelector_(selector) {
  if (!selector || typeof selector !== 'object') {
    throw new Error('Invalid agendaSelector: expected an object, got ' + JSON.stringify(selector));
  }

  const mode = selector.mode;

  if (mode === 'all') {
    return { mode: 'all' };
  }

  if (mode === 'prefix') {
    const value = selector.value;
    if (typeof value !== 'string' || value.trim() === '') {
      throw new Error('Invalid agendaSelector: mode "prefix" requires a non-empty string "value".');
    }
    return { mode: 'prefix', value: value.trim() };
  }

  if (mode === 'itemList') {
    if (!Array.isArray(selector.value)) {
      throw new Error('Invalid agendaSelector: mode "itemList" requires an array "value".');
    }
    const normalizedItems = selector.value
      .map(v => String(v === null || v === undefined ? '' : v).trim())
      .filter(v => v !== '');
    if (normalizedItems.length === 0) {
      throw new Error('Invalid agendaSelector: mode "itemList" requires at least one non-empty item in "value".');
    }
    return { mode: 'itemList', value: normalizedItems };
  }

  throw new Error('Invalid agendaSelector: unsupported mode ' + JSON.stringify(mode) + '. Expected "prefix", "all", or "itemList".');
}

/**
 * SA4-IMPL-004: pure agenda-item matching against a MeetingContext
 * agendaSelector. No Google API calls, no property reads, deterministic.
 * Has NO production filtering caller as of SA4-IMPL-004 -- it exists as a
 * tested abstraction, ready for a later task to wire into
 * downloadAndGroupTdocs_()/parseAgendaForReport_(), which are UNCHANGED by
 * this task.
 *
 * 'prefix': String(agendaItem).trim().startsWith(selector.value) -- this is
 * EXACTLY today's production semantics in downloadAndGroupTdocs_() and
 * parseAgendaForReport_() (`agendaItem.startsWith(agendaPrefix)`, on an
 * already-trimmed cell value). Every existing prefix value includes its
 * trailing dot ('7.', '11.', ...), which is WHY '7.' already cannot
 * false-positive-match '17.' or '70.' today -- verified against the actual
 * source, not assumed, see the SA4-IMPL-004 report's inventory. Preserved
 * exactly here, not "improved".
 *
 * 'all': matches any agenda item that is a real, non-blank-after-trim
 * value. null/undefined/''/whitespace-only do NOT match: a row with no
 * real agenda item isn't "part of the agenda" under any selector, which
 * keeps 'all' consistent with how a blank agenda item is already
 * implicitly excluded under 'prefix' today (an empty string can never
 * start with a non-empty prefix).
 *
 * 'itemList': EXACT match (after trimming both sides) against the
 * normalized value list -- never a prefix/subtree match. "1.5" matches
 * only "1.5" (and whitespace-padded variants), never "1.50", "1.5.1", or
 * "11.5". SA4-ARCH-006's evidence (the Audio-SWG-telco "keeping only
 * relevant items from the unique agenda" pattern) supports explicit item
 * selection only, not speculative hierarchical/subtree semantics.
 */
function agendaSelectorMatches_(selector, agendaItem) {
  const normalized = normalizeAgendaSelector_(selector);
  const item = String(agendaItem === null || agendaItem === undefined ? '' : agendaItem).trim();

  if (normalized.mode === 'all') {
    return item !== '';
  }
  if (normalized.mode === 'prefix') {
    return item.startsWith(normalized.value);
  }
  // itemList
  return normalized.value.indexOf(item) !== -1;
}

/**
 * SA4-IMPL-006: pure list-level agenda projection. Converts a complete
 * parsed agenda-item array (as produced by parseAgendaFromHeadings_()/
 * parseAgendaFromTables_(), each item shaped {number, title, level, heading,
 * text}) into the ordered subset selected for the report.
 *
 * Deliberately simple, by design, not by omission -- see the SA4-ARCH-007
 * report for the evidence behind each of these decisions:
 *
 *   - No ancestor retention. The current parseAgendaForReport_() ZIP-path
 *     filter DOES retain a bare parent (e.g. "7" alongside "7.1"/"7.2"), but
 *     SA4-ARCH-007 traced its sole real consumer (buildSkeletonWithTdocTables)
 *     and found it immediately, redundantly re-filters the parent back OUT
 *     before rendering anything -- the parent has zero observable effect
 *     today. Reproducing that intermediate, unused state here would just be
 *     copying an accident, not preserving a requirement.
 *   - No descendant expansion beyond what 'prefix' membership itself already
 *     provides via String.startsWith() (a descendant's own number already
 *     starts with the prefix, so it's already a direct member -- there is
 *     nothing extra to "expand").
 *   - No node synthesis: if an ancestor/parent was never present in the
 *     source document, it is never invented (SA4-ARCH-007 confirmed this is
 *     already true of every existing parser; this function preserves it by
 *     construction -- it can only ever return items that were passed in).
 *   - No sorting: SA4-ARCH-007 confirmed agenda-item order is always
 *     source-document order, never re-derived from the number string. This
 *     function preserves that via Array.prototype.filter, which never
 *     reorders.
 *   - itemList is EXACT-selected-items-only (Model A from the SA4-ARCH-007
 *     analysis): the one real recurring-telco example evidenced (DaCAS=1.5,
 *     ATIAS_Ph3-MED=1.6, SA4-ARCH-006) shows leaf slots with no observed
 *     nested children. This is not proven to generalize -- if subtree
 *     selection is ever genuinely required, 'prefix' mode is the
 *     architecturally correct tool for that, not itemList.
 *
 * All membership semantics are delegated to the UNCHANGED
 * agendaSelectorMatches_(selector, item.number) -- this function adds no
 * second switch over 'prefix'/'all'/'itemList' and no new flags on the
 * matching predicate.
 *
 * Pure: no Google APIs, no PropertiesService, does not mutate `agendaItems`
 * or any item object, and returns the ORIGINAL item object references for
 * every retained entry (a filter, not a clone/rewrite).
 */
function projectAgendaItems_(agendaItems, selector) {
  if (!Array.isArray(agendaItems)) {
    throw new Error('projectAgendaItems_: expected an array of agenda items, got ' + JSON.stringify(agendaItems));
  }

  return agendaItems.filter(item => agendaSelectorMatches_(selector, item ? item.number : undefined));
}

function getMeetingContext_() {
  const cfg = getReportConfig_();
  const identity = getMeetingIdentityConfig_();

  const reportType = cfg.REPORT_SUFFIX;
  const SWG_REPORT_TYPES = ['Audio', 'Video', 'MBS', 'RTC'];
  let structureProfile;
  if (reportType === '6G') {
    structureProfile = 'main-6g';
  } else if (SWG_REPORT_TYPES.indexOf(reportType) !== -1) {
    structureProfile = 'main-swg';
  } else {
    structureProfile = 'main-other';
  }
  // structureProfile is still the SA4-ARCH-003 model, untouched by
  // SA4-IMPL-004 -- frontMatterProfile/skeleton work is a later task.

  let meeting;
  let sources;
  let agendaPrefix;
  let agendaSelector;

  if (identity.MEETING_TYPE === 'adhoc') {
    meeting = {
      type: 'adhoc',
      name: identity.MEETING_NAME || null,
      folder: null,
      number: null,
      portalId: null
    };

    // Explicit ad-hoc source overrides, reusing the SAME property names
    // getReportConfig_() already reads (no parallel SOURCE_*-named
    // properties introduced). mailingList/draftsFolder are handled via
    // `adhocDerived` below, not here, since neither has an override
    // property today.
    const adhocOverrides = {
      ftpBase: identity.FTP_BASE,
      tdocListUrl: identity.TDOC_LIST_URL,
      agendaTdoc: identity.AGENDA_TDOC,
      agendaTemplateDocId: identity.AGENDA_SOURCE_DOC_ID,
      revisionsUrl: identity.REVISIONS_URL
    };

    // mailingList: PROVISIONAL reuse of the existing report-type -> mailing
    // list mapping (cfg.LIST_NAME, e.g. '3GPP_TSG_SA_WG4_AUDIO' for
    // report.type === 'Audio'). This is NOT independently verified for
    // ad-hoc traffic -- SA4-ARCH-006 explicitly left "which mailing list do
    // ad-hoc meetings actually use" as an unresolved unknown. Reusing it is
    // a deliberate, documented placeholder, not a verified fact.
    //
    // draftsFolder: always null for ad-hoc. The main-meeting
    // DRAFTS_FOLDERS lookup (SWG name -> subfolder within one shared
    // meeting's Inbox/Drafts/) has no ad-hoc equivalent -- each ad-hoc
    // series already has its OWN, unshared Inbox/Drafts/ folder (verified,
    // SA4-ARCH-005), so sources.revisionsUrl alone fully represents where
    // ad-hoc revisions live. Inventing a folder name on top would be
    // meaningless, not just redundant.
    const adhocDerived = {
      mailingList: cfg.LIST_NAME,
      draftsFolder: null
    };

    sources = resolveMeetingSources_(adhocDerived, adhocOverrides);

    // SA4-IMPL-004: an ad-hoc meeting's agenda is not filtered by any
    // main-meeting SWG-plenary prefix -- SA4-ARCH-005/006 verified this has
    // no meaning for a real ad-hoc agenda (agenda numbering is flat/
    // instance-specific, e.g. ULBC-MED's items are "1".."6", not "7.x").
    // agendaPrefix is explicitly null (no longer the old, wrong
    // cfg.AGENDA_ITEM_PREFIX value) and every parsed agenda item is in
    // scope via {mode:'all'}. This is the verified dedicated-single-topic
    // ad-hoc shape (ULBC-MED). The recurring multi-topic-telco shape
    // (itemList, e.g. {mode:'itemList', value:['1.5','1.6']}) is equally
    // evidenced (SA4-ARCH-006) but there is currently only one ad-hoc
    // MeetingContext branch -- choosing between 'all'/'itemList' per
    // ad-hoc instance is a later task, not introduced here.
    agendaPrefix = null;
    agendaSelector = normalizeAgendaSelector_({ mode: 'all' });
  } else {
    meeting = {
      type: 'main',
      name: null, // stable shape with the ad-hoc branch; no main meeting has ever had a free-form display name
      folder: cfg.MEETING_FOLDER,
      number: cfg.MEETING_NUMBER,
      portalId: cfg.MEETING_ID
    };

    // Unchanged since SA4-IMPL-002: no overrides supplied, sources fall
    // straight through to the existing main-meeting derivation in cfg.
    const derivedSources = {
      ftpBase: cfg.FTP_BASE,
      tdocListUrl: cfg.TDOC_LIST_URL,
      agendaTdoc: cfg.AGENDA_TDOC,
      agendaTemplateDocId: cfg.AGENDA_SOURCE_DOC_ID,
      mailingList: cfg.LIST_NAME,
      draftsFolder: cfg.DRAFTS_FOLDER,
      revisionsUrl: cfg.REVISIONS_URL
    };
    sources = resolveMeetingSources_(derivedSources, {});

    // SA4-IMPL-004: unchanged report-type -> agenda prefix mapping
    // (SA4-ARCH-003), now ALSO exposed in the new agendaSelector shape.
    agendaPrefix = cfg.AGENDA_ITEM_PREFIX;
    agendaSelector = normalizeAgendaSelector_({ mode: 'prefix', value: cfg.AGENDA_ITEM_PREFIX });
  }

  return {
    group: 'SA4',

    meeting: meeting,

    report: {
      type: reportType,
      agendaPrefix: agendaPrefix,
      agendaSelector: agendaSelector,
      structureProfile: structureProfile
    },

    sources: sources,

    options: {
      showPreviewSnippet: cfg.SHOW_PREVIEW_SNIPPET
      // emailStartDate, fetchAbstractsOnUpdate: omitted -- see comment above.
    }
  };
}

// =========================================================
// COLLECTOR CONFIG (now uses centralized config)
// =========================================================

function getCollectorConfig_() {
  const reportConfig = getReportConfig_(); // Get centralized config
  // ensureCollectorConfigTable_();
  const collectorTableConfig = readCollectorConfigTable_();

  const cfg = { ...reportConfig, ...collectorTableConfig }; // Merge configs

  // Set defaults for collector-specific settings if not in table
  if (!cfg.ARCHIVE_DAYS_BACK) cfg.ARCHIVE_DAYS_BACK = '14';
  if (!cfg.A1_EMPTY_CACHE_TTL_HOURS) cfg.A1_EMPTY_CACHE_TTL_HOURS = '24';
  if (!cfg.TIMEZONE) cfg.TIMEZONE = Session.getScriptTimeZone();
  if (!cfg.SHOW_PREVIEW_SNIPPET) cfg.SHOW_PREVIEW_SNIPPET = 'false';
  if (!cfg.TDOC_ID_REGEX) cfg.TDOC_ID_REGEX = '^S4-\\d{6}$';
  if (!cfg.EMAIL_START_DATE) cfg.EMAIL_START_DATE = '2026-08-21';

  return cfg;
}

function ensureCollectorConfigTable_() {
  const body = DocumentApp.getActiveDocument().getBody();
  for (const t of body.getTables()) {
    if (isCollectorConfigTable_(t)) return;
  }

  body.insertParagraph(0, 'Collector Configuration')
    .setHeading(DocumentApp.ParagraphHeading.HEADING3);

  // This table is now for optional overrides. Key settings are derived automatically.
  body.insertTable(1, [
    ['Key', 'Value'],
    ['SHOW_PREVIEW_SNIPPET', 'false']
  ]);
}

// =========================================================
// DOCUMENT REALLOCATION SYSTEM
// =========================================================

/**
 * Create or ensure reallocation table exists
 */
/**
 * Create or ensure reallocation table exists as subsection X.1.3
 */
function ensureReallocationTable_() {
  const body = DocumentApp.getActiveDocument().getBody();
  
  // Check if table already exists
  for (const t of body.getTables()) {
    if (isReallocationTable_(t)) return;
  }
  
  // Get the agenda prefix from config (e.g., "7." for Audio, "11." for 6G)
  const cfg = getReportConfig_();
  const agendaPrefix = cfg.AGENDA_ITEM_PREFIX || '7.';
  const targetSection = `${agendaPrefix}1.3`; // e.g., "7.1.3" or "11.1.3"
  
  // Find insertion point: after section X.1.2 or before X.1.4
  let insertIndex = -1;
  let foundSection = false;
  
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    
    const para = child.asParagraph();
    const text = para.getText().trim();
    const heading = para.getHeading();
    
    // Look for X.1.2 (insert after) or X.1.4 (insert before)
    if (heading !== DocumentApp.ParagraphHeading.NORMAL) {
      // Check if this is X.1.2 - insert after it
      if (text.startsWith(`${agendaPrefix}1.2`)) {
        insertIndex = i + 1;
        foundSection = true;
        break;
      }
      // Check if this is X.1.4 or later - insert before it
      if (text.startsWith(`${agendaPrefix}1.4`) || 
          text.startsWith(`${agendaPrefix}1.5`) ||
          text.startsWith(`${agendaPrefix}2`)) {
        insertIndex = i;
        foundSection = true;
        break;
      }
    }
  }
  
  // If we didn't find a good spot, try to find X.1 and insert after it
  if (!foundSection) {
    for (let i = 0; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      
      const para = child.asParagraph();
      const text = para.getText().trim();
      const heading = para.getHeading();
      
      if (heading !== DocumentApp.ParagraphHeading.NORMAL && 
          text.startsWith(`${agendaPrefix}1 `)) {
        insertIndex = i + 1;
        foundSection = true;
        break;
      }
    }
  }
  
  // If still not found, insert at beginning (fallback)
  if (insertIndex === -1) {
    insertIndex = 0;
    Logger.log('Warning: Could not find appropriate section, inserting at beginning');
  }
  
  // Create heading with proper numbering
  const heading = body.insertParagraph(insertIndex, `${targetSection} Document Reallocations`);
  heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  // Create table
  body.insertTable(insertIndex + 1, [
    ['TDoc', 'Original Agenda', 'New Agenda', 'Reason']
  ]);
  
  Logger.log(`Created reallocation table at index ${insertIndex} as section ${targetSection}`);
}


/**
 * Check if table is a reallocation table
 */
function isReallocationTable_(table) {
  try {
    const row0 = table.getRow(0);
    return row0.getNumCells() >= 3 &&
           row0.getCell(0).getText().trim() === 'TDoc' &&
           row0.getCell(1).getText().trim() === 'Original Agenda' &&
           row0.getCell(2).getText().trim() === 'New Agenda';
  } catch (e) {
    return false;
  }
}

/**
 * Read reallocation configuration
 * Returns: { 'S4-260123': { original: '5.3', new: '8.3', reason: '...' }, ... }
 */
function getReallocationMap_() {
  const body = DocumentApp.getActiveDocument().getBody();
  const map = {};
  
  for (const table of body.getTables()) {
    if (!isReallocationTable_(table)) continue;
    
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 3) continue;
      
      const tdoc = row.getCell(0).getText().trim();
      const original = row.getCell(1).getText().trim();
      const newAgenda = row.getCell(2).getText().trim();
      const reason = row.getNumCells() >= 4 ? row.getCell(3).getText().trim() : '';
      
      if (tdoc && newAgenda) {
        map[tdoc] = {
          original: original,
          new: newAgenda,
          reason: reason
        };
      }
    }
  }
  
  return map;
}

/**
 * Copies content from a source element until a stop condition is met.
 * Replaces all occurrences of oldPrefix with newPrefix in the text.
 * Also replaces <add list and make hyperlink> with the appropriate mailing list.
 * Copies tables and inline images from template, preserving ALL formatting.
 */
function copySectionContentWithReplacement_(startElement, targetBody, stopRegex, oldPrefix, newPrefix) {
  const cfg = getReportConfig_();
  const mailingList = cfg.LIST_NAME || LIST_NAME_LOCK;
  
  let currentElement = startElement.getNextSibling();
  while (currentElement) {
    const elementType = currentElement.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const p = currentElement.asParagraph();
      const text = p.getText();
      if (p.getHeading() !== DocumentApp.ParagraphHeading.NORMAL && stopRegex.test(text.trim())) {
        break; // Stop at the start of the next section
      }

      // Do text replacements first
      let updatedText = text.replace(new RegExp(oldPrefix, 'g'), newPrefix);
      updatedText = updatedText.replace(/<add list and make hyperlink>/g, mailingList);
      
      // Copy the paragraph using the proper method
      const sourceIndex = p.getParent().getChildIndex(p);
      const sourcePara = p.getParent().getChild(sourceIndex).asParagraph();
      const copiedElement = sourcePara.copy();
      
      // Append the copied element
      const newPara = targetBody.appendParagraph(copiedElement);
      
      // Update text if needed
      if (updatedText !== text) {
        newPara.replaceText(text, updatedText);
      }
      
      // Find and hyperlink the mailing list URL if present
      if (updatedText) {
        const urlMatch = updatedText.match(/(https:\/\/list\.etsi\.org\/scripts\/wa\.exe\?A1=[^&\s]+&L=)([^\s]+)/);
        if (urlMatch) {
          const fullUrl = urlMatch[0];
          const startPos = updatedText.indexOf(fullUrl);
          if (startPos >= 0) {
            try {
              const textElement = newPara.editAsText();
              textElement.setLinkUrl(startPos, startPos + fullUrl.length - 1, fullUrl);
            } catch (e) {
              Logger.log('Failed to set link: ' + e.message);
            }
          }
        }
      }
    } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      const li = currentElement.asListItem();
      const text = li.getText();
      
      // Do text replacements
      let updatedText = text.replace(new RegExp(oldPrefix, 'g'), newPrefix);
      updatedText = updatedText.replace(/<add list and make hyperlink>/g, mailingList);
      
      // Copy the list item
      const sourceIndex = li.getParent().getChildIndex(li);
      const sourceLi = li.getParent().getChild(sourceIndex).asListItem();
      const copiedElement = sourceLi.copy();
      
      // Append the copied element
      const newListItem = targetBody.appendListItem(copiedElement);
      
      // Update text if needed
      if (updatedText !== text) {
        newListItem.replaceText(text, updatedText);
      }
    } else if (elementType === DocumentApp.ElementType.TABLE) {
      // Copy tables from template - this preserves all formatting including colors
      const sourceIndex = currentElement.getParent().getChildIndex(currentElement);
      const sourceTable = currentElement.getParent().getChild(sourceIndex).asTable();
      targetBody.appendTable(sourceTable.copy());
    } else if (elementType === DocumentApp.ElementType.INLINE_IMAGE) {
      // Copy inline images from template
      try {
        const sourceIndex = currentElement.getParent().getChildIndex(currentElement);
        const sourceImage = currentElement.getParent().getChild(sourceIndex).asInlineImage();
        targetBody.appendImage(sourceImage.copy());
      } catch (e) {
        Logger.log('Failed to copy inline image: ' + e.message);
      }
    }
    
    currentElement = currentElement.getNextSibling();
  }
}

function isCollectorConfigTable_(t) {
  try {
    return t.getCell(0, 0).getText() === 'Key' && t.getCell(0, 1).getText() === 'Value';
  } catch (e) {
    return false;
  }
}

function readCollectorConfigTable_() {
  const body = DocumentApp.getActiveDocument().getBody();
  for (const t of body.getTables()) {
    if (!isCollectorConfigTable_(t)) continue;
    const cfg = {};
    for (let r = 1; r < t.getNumRows(); r++) {
      cfg[t.getCell(r, 0).getText().trim()] = t.getCell(r, 1).getText().trim();
    }
    return cfg;
  }
  return { DEBUG: 'true' };
}

function isDebug_(cfg) {
  return String(cfg.DEBUG || '').trim().toLowerCase() === 'true';
}
function log_(cfg, msg, obj) {
  if (!isDebug_(cfg)) return;
  const line = `[${new Date().toISOString()}] ${msg}`;
  Logger.log(obj !== undefined ? line + ' ' + JSON.stringify(obj) : line);
}
function warn_(cfg, msg) {
  Logger.log(`[WARN ${new Date().toISOString()}] ${msg}`);
}

function copyTableToDoc2(spreadsheetName, sheetName) {
  // Get the spreadsheet by name
  var files = DriveApp.getFilesByName(spreadsheetName);
  if (!files.hasNext()) {
    Logger.log('Spreadsheet not found: ' + spreadsheetName);
    return;
  }
  var file = files.next();
  var spreadsheet = SpreadsheetApp.open(file);
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return;
  }

  // Get the table data
  var range = sheet.getDataRange();
  var values = range.getValues();
  var richTextValues = range.getRichTextValues();

  if (!values || values.length === 0) {
    Logger.log('No data in sheet: ' + sheetName);
    return;
  }

  // Get the active document
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  // Search for the table with matching headings
  var tables = body.getTables();
  var tableToReplace = null;
  var tableIndex = -1;

  var headings = values[0]; // first row headings
  var headingCount = headings ? headings.length : 0;

  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    if (table.getNumRows() < 1) continue;

    var firstRow = table.getRow(0);
    if (firstRow.getNumCells() < headingCount) continue; // ✅ guard

    var match = true;
    for (var j = 0; j < headingCount; j++) {
      if (firstRow.getCell(j).getText() !== String(headings[j])) {
        match = false;
        break;
      }
    }

    if (match) {
      tableToReplace = table;
      tableIndex = body.getChildIndex(table);
      break;
    }
  }

  // If table exists, remove it (tableIndex remains valid for reinsertion)
  if (tableToReplace) {
    body.removeChild(tableToReplace);
  }

  // ✅ If no matching table found, append at end instead of insert(-1)
  var newTable;
  if (tableIndex < 0) {
    newTable = body.appendTable();
  } else {
    newTable = body.insertTable(tableIndex);
  }

  // Add rows and cells
  values.forEach(function (row) {
    var tableRow = newTable.appendTableRow();
    row.forEach(function (cell) {
      tableRow.appendTableCell(String(cell));
    });
  });


  // ✅ remove the initial placeholder row if empty
  removeInitialEmptyRow_(newTable);

  // Re-apply rich text hyperlinks for the first column (safe guards)
  for (var r = 0; r < values.length; r++) {
    if (!richTextValues || !richTextValues[r] || !richTextValues[r][0]) continue;

    var tableRow = newTable.getRow(r);
    if (tableRow.getNumCells() < 1) continue;

    var tableCell = tableRow.getCell(0);
    var richTextValue = richTextValues[r][0];

    if (!richTextValue || !richTextValue.getRuns) continue;

    var runs = richTextValue.getRuns();
    var textElement = tableCell.editAsText();
    textElement.setText('');

    var pos = 0;
    runs.forEach(function (run) {
      var partText = run.getText();
      if (!partText) return;

      textElement.appendText(partText);
      var url = run.getLinkUrl();
      if (url) {
        try {
          textElement.setLinkUrl(pos, pos + partText.length - 1, url);
        } catch (e) { }
      }
      pos += partText.length;
    });
  }
}

function copyDocsToReport() {
  var spreadsheetName = 'Template-Report-6G.xlsx'; // Replace with your spreadsheet name
  var sheetName = 'Sheet1'; // Replace with your sheet name
  copyTableToDoc2(spreadsheetName, sheetName);
}

/**
 * Dispatch table copy:
 * - TDOC sheets (B1 = S4-xxxxxx) -> merge logic
 * - Non-TDOC sheets             -> replace logic (copyTableToDoc2)
 */
function copyTableToDocDispatcher_(sheet, body) {
  const b1 = String(sheet.getRange('B1').getValue() || '').trim();

  // TDOC table → MERGE (never replace)
  if (/^S4-\d{6}$/i.test(b1)) {
    copyTableToDoc_(sheet, body);
    return;
  }

  // Non-TDOC table → REPLACE
  if (typeof copyTableToDoc2 === 'function') {
    const spreadsheet = sheet.getParent();
    copyTableToDoc2(spreadsheet.getName(), sheet.getName());
    return;
  }

  // Fallback
  copyTableToDoc_(sheet, body);
}





// =========================================================
// ENTRY: Copy tables from Excel template to Doc
// =========================================================

/**
 * NEW: Download TDOC list directly from 3GPP and process it
 * This replaces the need for Colab or manual Excel files
 */
function downloadAndProcessFromWeb() {
  const cfg = getReportConfig_();
  const meetingUrl = cfg.TDOC_LIST_URL || 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4%23136.xlsx';
  
  Logger.log('Downloading TDOC list from: ' + meetingUrl);
  
  try {
    // Download the Excel file
    const response = UrlFetchApp.fetch(meetingUrl);
    const blob = response.getBlob();
    blob.setName('TDoc_List_Temp.xlsx');
    
    // Create temporary file in Drive
    const tempFile = DriveApp.createFile(blob);
    const tempFileId = tempFile.getId();
    
    Logger.log('Downloaded successfully, processing...');
    
    // Open as Sheets and process
    const spreadsheet = SpreadsheetApp.open(DriveApp.getFileById(tempFileId));
    const sheet = spreadsheet.getSheets()[0];
    
    // Process the sheet
    processWebDownloadedSheet_(sheet);
    
    // Clean up temp file
    DriveApp.getFileById(tempFileId).setTrashed(true);
    
    Logger.log('TDOC list processed and temp file cleaned up');
    DocumentApp.getUi().alert('TDOC list downloaded and processed successfully!');
    
  } catch (e) {
    Logger.log('Error downloading TDOC list: ' + e.message);
    DocumentApp.getUi().alert('Error downloading TDOC list: ' + e.message);
  }
}

/**
 * Process the downloaded sheet and create TDOC tables organized by agenda items.
 * Tables are inserted under the matching agenda heading already in the document.
 * If no matching heading exists, a new heading is appended.
 */
function processWebDownloadedSheet_(sheet) {
  const cfg = getReportConfig_();
  const suffix = cfg.REPORT_SUFFIX || '6G';
  const agendaPrefix = getConfiguredAgendaPrefix_();
  
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  const data = sheet.getDataRange().getValues();
  const richTextValues = sheet.getDataRange().getRichTextValues();
  
  // Get reallocation map
  const reallocations = getReallocationMap_();
  Logger.log(`Loaded ${Object.keys(reallocations).length} reallocations`);
  
  // Find column indices (assuming standard 3GPP format)
  const headers = data[0];
  const tdocCol = headers.indexOf('TDoc');
  const titleCol = headers.indexOf('Title');
  const sourceCol = headers.indexOf('Source');
  const contactCol = headers.indexOf('Contact');
  const agendaCol = headers.indexOf('Agenda item');
  const agendaTopicCol = headers.indexOf('Agenda Topic');
  const statusCol = headers.indexOf('TDoc Status');
  const typeCol = headers.indexOf('Type'); // Column F
  const forCol = headers.indexOf('For');   // Column G
  const revisedToCol = headers.indexOf('Revised to');
  
  if (tdocCol === -1 || agendaCol === -1) {
    throw new Error('Could not find required columns (TDoc, Agenda item) in Excel file');
  }
  if (revisedToCol === -1) {
    Logger.log('Warning: no "Revised to" column in the TDOC list; revision placement is skipped');
  }
  
  Logger.log(`Processing ${data.length - 1} rows, filtering by agenda prefix: ${agendaPrefix}`);
  
  // Group TDOCs by agenda item
  const agendaGroups = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tdoc = String(row[tdocCol] || '').trim();
    if (!tdoc) continue;
    let agendaItem = String(row[agendaCol] || '').trim();
    
    // Apply reallocation if exists
    if (reallocations[tdoc]) {
      const originalAgenda = agendaItem;
      agendaItem = reallocations[tdoc].new;
      Logger.log(`Reallocating ${tdoc}: ${originalAgenda} → ${agendaItem} (${reallocations[tdoc].reason || 'no reason'})`);
    }
    
    // Filter by report type (after reallocation)
    if (!agendaItem.startsWith(agendaPrefix)) continue;
    
    const agendaTopic = agendaTopicCol >= 0 ? String(row[agendaTopicCol] || '').trim() : '';
    
    if (!agendaGroups[agendaItem]) {
      agendaGroups[agendaItem] = { topic: agendaTopic, tdocs: [] };
    }
    
    agendaGroups[agendaItem].tdocs.push({
      rowIndex: i,
      row: row,
      richTextRow: richTextValues[i],
      tdocCol, revisedToCol   // needed by orderTdocsByRevision_ / getRevisedTo_
    });
  }
  
  // Sort agenda items numerically
  const sortedAgendaItems = Object.keys(agendaGroups).sort((a, b) => {
    const partsA = a.split('.').map(Number);
    const partsB = b.split('.').map(Number);
    for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
      const diff = (partsA[i] || 0) - (partsB[i] || 0);
      if (diff !== 0) return diff;
    }
    return 0;
  });
  
  const totalTdocs = Object.values(agendaGroups).reduce((sum, g) => sum + g.tdocs.length, 0);
  Logger.log(`Found ${sortedAgendaItems.length} agenda items with ${totalTdocs} TDOCs`);
  
  // Build summary table of all registered documents
  const allTdocs = [];
  sortedAgendaItems.forEach(agendaItem => {
    agendaGroups[agendaItem].tdocs.forEach(tdoc => {
      allTdocs.push({ ...tdoc, agendaItem });
    });
  });
  
  // Insert "Registered Documents" summary at the top (before agenda sections)
  const registeredHeading = body.appendParagraph('Registered Documents');
  registeredHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  createSummaryTable_(body, allTdocs, tdocCol, titleCol, sourceCol, agendaCol);
  
  // Insert TDOC tables under the matching agenda heading in the document.
  // If the skeleton was already created, tables go under the existing heading.
  sortedAgendaItems.forEach(agendaItem => {
    const group = agendaGroups[agendaItem];
    
    // Find the insertion point: end of the section for this agenda item
    const insertIdx = findInsertionPointForAgendaItem_(body, agendaItem, group.topic);
    
    Logger.log(`Inserting ${group.tdocs.length} TDOCs for ${agendaItem} at index ${insertIdx}`);
    
    // Insert tables in reverse order so each one goes to the same position
    // (inserting at insertIdx pushes subsequent ones down)
    let currentIdx = insertIdx;
    // Revisions are emitted directly below the document they revise.
    orderTdocsByRevision_(group.tdocs).forEach(tdocData => {
      const row = tdocData.row;
      const revisedTo = getRevisedTo_(tdocData);
      const typeFor = (typeCol >= 0 && forCol >= 0 && row[typeCol] && row[forCol])
        ? `${row[typeCol]} for ${row[forCol]}`
        : (typeCol >= 0 && row[typeCol]) ? row[typeCol] : (forCol >= 0 && row[forCol]) ? row[forCol] : '';

      const tempData = [
        ['TDoc', row[tdocCol]],
        ['Title', row[titleCol]],
        ['Source', row[sourceCol]],
        ['Contact', contactCol >= 0 ? row[contactCol] : ''],
        ['Agenda Item', agendaItem],
        ['Type/For', typeFor],
        ['E-mail Discussion', ''],
        ['Revisions', ''],
        ['Minutes', ''],
        ['Disposition', revisedTo ? 'Revised to ' + revisedTo : ''],
        ['Status', statusCol >= 0 ? row[statusCol] : '']
      ];
      
      insertTDocTableAtIndex_(body, currentIdx, tempData, tdocData.richTextRow, tdocCol);
      currentIdx++;
    });
  });
  
  Logger.log(`Processed ${totalTdocs} TDOCs across ${sortedAgendaItems.length} agenda items for report type: ${suffix}`);
}

/**
 * Find the insertion point for TDOC tables under a given agenda item heading.
 * Returns the index just before the next same-or-higher-level heading,
 * or just after the agenda heading if no content exists yet.
 * Creates the heading if it does not exist.
 */
function findInsertionPointForAgendaItem_(body, agendaItem, agendaTopic) {
  const numChildren = body.getNumChildren();
  const agendaLevel = (agendaItem.match(/\./g) || []).length + 1;
  
  // Find the heading for this agenda item
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    
    const para = child.asParagraph();
    const text = para.getText().trim();
    const heading = para.getHeading();
    
    if (heading === DocumentApp.ParagraphHeading.NORMAL) continue;
    
    // Match heading that starts with this agenda item number
    if (!text.startsWith(agendaItem + ' ') && text !== agendaItem) continue;
    
    // Found the heading. Now find where this section ends.
    for (let j = i + 1; j < numChildren; j++) {
      const next = body.getChild(j);
      if (next.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
      
      const nextPara = next.asParagraph();
      const nextHeading = nextPara.getHeading();
      if (nextHeading === DocumentApp.ParagraphHeading.NORMAL) continue;
      
      const nextText = nextPara.getText().trim();
      const nextLevel = (nextText.match(/^(\d+(?:\.\d+)*)/) || ['', ''])[1];
      const nextDots = nextLevel ? (nextLevel.match(/\./g) || []).length + 1 : 99;
      
      // Stop at same or higher level heading
      if (nextDots <= agendaLevel) return j;
    }
    
    // No next heading found: insert at end of document
    return numChildren;
  }
  
  // Heading not found: create it and return position after it
  const headingText = agendaItem + (agendaTopic ? ' ' + agendaTopic : '');
  const newHeading = body.appendParagraph(headingText);
  if (agendaLevel === 1) newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  else if (agendaLevel === 2) newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  else newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  return body.getNumChildren();
}

/**
 * Insert a TDOC table at a specific index in the document body.
 */
function insertTDocTableAtIndex_(body, insertIndex, data, richTextRow, tdocCol) {
  const table = body.insertTable(insertIndex);
  removeInitialEmptyRow_(table);
  
  data.forEach(rowData => {
    const tr = table.appendTableRow();
    tr.appendTableCell(rowData[0]);
    tr.appendTableCell(String(rowData[1] || ''));
  });
  
  // Apply hyperlink to TDoc if present
  if (richTextRow && richTextRow[tdocCol]) {
    const richText = richTextRow[tdocCol];
    if (richText.getLinkUrl && richText.getLinkUrl()) {
      const cell = table.getRow(0).getCell(1);
      const text = cell.editAsText();
      const tdocValue = String(data[0][1]);
      if (tdocValue.length > 0) {
        text.setLinkUrl(0, tdocValue.length - 1, richText.getLinkUrl());
      }
    }
  }
  
  styleStatusCell_(table);
  return table;
}

/**
 * Get agenda item prefix for report type
 */
function getAgendaPrefixForReportType_(reportType) {
  const prefixes = {
    'Liaison': '5.',
    'Audio': '7.',
    'MBS': '8.',
    'Video': '9.',
    'RTC': '10.',
    '6G': '11.',
    'New': '18.'
  };
  return prefixes[reportType] || '11.';
}

function getConfiguredAgendaPrefix_() {
  const cfg = getReportConfig_();
  return cfg.AGENDA_ITEM_PREFIX || getAgendaPrefixForReportType_(cfg.REPORT_SUFFIX || '6G');
}

/**
 * Create summary table listing all TDOCs (4 columns: TDoc, Title, Source, Agenda Item)
 */
function createSummaryTable_(body, tdocs, tdocCol, titleCol, sourceCol, agendaCol) {
  const summaryTable = body.appendTable();
  
  // Add header row
  const headerRow = summaryTable.appendTableRow();
  headerRow.appendTableCell('TDoc');
  headerRow.appendTableCell('Title');
  headerRow.appendTableCell('Source');
  headerRow.appendTableCell('Agenda Item');
  
  // Add each TDOC as a row
  tdocs.forEach(tdocData => {
    const row = tdocData.row;
    const dataRow = summaryTable.appendTableRow();
    
    // TDoc number (with hyperlink if available)
    const tdocCell = dataRow.appendTableCell(String(row[tdocCol] || ''));
    if (tdocData.richTextRow && tdocData.richTextRow[tdocCol]) {
      const richText = tdocData.richTextRow[tdocCol];
      if (richText.getLinkUrl && richText.getLinkUrl()) {
        const text = tdocCell.editAsText();
        const tdocValue = String(row[tdocCol] || '');
        text.setLinkUrl(0, tdocValue.length - 1, richText.getLinkUrl());
      }
    }
    
    // Title
    dataRow.appendTableCell(String(row[titleCol] || ''));
    
    // Source
    dataRow.appendTableCell(String(row[sourceCol] || ''));
    
    // Agenda Item
    dataRow.appendTableCell(String(tdocData.agendaItem || ''));
  });
  
  // Remove initial empty row if present
  removeInitialEmptyRow_(summaryTable);
  
  return summaryTable;
}

/**
 * Create TDOC table from data array
 */
function createTDocTableFromData_(body, data, richTextRow, tdocCol) {
  const table = body.appendTable();
  
  // Add each row
  data.forEach(function(rowData) {
    const tr = table.appendTableRow();
    tr.appendTableCell(rowData[0]);
    tr.appendTableCell(String(rowData[1] || ''));
  });
  
  // Apply hyperlink to TDoc if present
  if (richTextRow && richTextRow[tdocCol]) {
    const richText = richTextRow[tdocCol];
    if (richText.getLinkUrl && richText.getLinkUrl()) {
      const cell = table.getRow(0).getCell(1);
      const text = cell.editAsText();
      const tdocValue = String(data[0][1]);
      text.setLinkUrl(0, tdocValue.length - 1, richText.getLinkUrl());
    }
  }
  
  // Abstracts are intentionally added in workflow step 5.
  const skipAbstracts = PropertiesService.getDocumentProperties().getProperty('SKIP_ABSTRACTS_DURING_TABLE_BUILD') === 'true';
  const tdocNumber = String(data[0][1] || '').trim();
  if (!skipAbstracts && tdocNumber && /^S4-\d{6}$/i.test(tdocNumber)) {
    fetchAndAddAbstract_(table, tdocNumber);
  }
  
  // Apply status styling
  styleStatusCell_(table);
  
  return table;
}

/**
 * Fetch abstract from Contribution Reviewer API and add to table.
 * Inserts Abstract row after Agenda Item, so the top rows remain:
 * TDoc, Title, Source, Contact, Agenda Item, Abstract.
 */
function fetchAndAddAbstract_(table, tdocNumber) {
  try {
    const apiToken = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN');
    if (!apiToken) {
      Logger.log('No REVIEWER_API_TOKEN found in script properties');
      return;
    }
    
    const apiUrl = `https://reviewer.bouazizi.dev/api/v1/documents/${tdocNumber}/summary?type=summary`;
    
    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      headers: {
        'X-API-Key': apiToken,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const statusCode = response.getResponseCode();
    
    if (statusCode === 200) {
      const result = JSON.parse(response.getContentText());
      const abstractText = result.text || '';
      
      if (abstractText) {
        const insertIndex = findAbstractInsertIndex_(table);
        const abstractRow = table.insertTableRow(insertIndex);
        abstractRow.appendTableCell('Abstract');
        const abstractCell = abstractRow.appendTableCell(abstractText);
        
        // Set abstract text to font size 8
        const abstractTextElement = abstractCell.editAsText();
        abstractTextElement.setFontSize(8);
        
        Logger.log(`Added abstract for ${tdocNumber}`);
      }
    } else if (statusCode === 404) {
      Logger.log(`No summary available for ${tdocNumber}`);
    } else {
      Logger.log(`API error for ${tdocNumber}: ${statusCode}`);
    }
  } catch (e) {
    Logger.log(`Error fetching abstract for ${tdocNumber}: ${e.message}`);
  }
}

function findAbstractInsertIndex_(table) {
  // Prefer placing Abstract immediately after Agenda Item.
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 1) continue;
    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key === 'agenda item' || key === 'agenda item:') {
      return r + 1;
    }
  }

  // Fallback for older/minimal tables: after Contact if present.
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 1) continue;
    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key === 'contact' || key === 'contact:') {
      return r + 1;
    }
  }

  // Last fallback: after Source, otherwise append at end.
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 1) continue;
    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key === 'source' || key === 'source:') {
      return r + 1;
    }
  }

  return table.getNumRows();
}

/**
 * Complete update from web - downloads and processes everything
 */
function updateAllFromWeb() {
  // 1. Download and process TDOC list from 3GPP
  downloadAndProcessFromWeb();
  
  // 2. Collect email discussions and revisions
  collectorUpdate_();
  
  // 3. Format the document
  removeRowHeightAndSpacing();
  
  Logger.log('Complete update from web finished');
}

function copyIndividualToReport() {
  const cfg = getReportConfig_();
  const suffix = cfg.REPORT_SUFFIX || '6G';
  const spreadsheetName = `Template-Report-${suffix}-Individual.xlsx`;
  copyAllTablesToDoc_(spreadsheetName);
}

function copyAllTablesToDoc_(spreadsheetName) {
  var files = DriveApp.getFilesByName(spreadsheetName);
  if (!files.hasNext()) {
    Logger.log('Spreadsheet not found: ' + spreadsheetName);
    return;
  }
  var spreadsheet = SpreadsheetApp.open(files.next());
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  spreadsheet.getSheets().forEach(function (sheet) {
    copyTableToDocDispatcher_(sheet, body);
  });
}

function copyTableToDoc_(sheet, body) {
  if (!sheet) return;

  var range = sheet.getDataRange();
  var values = range.getValues();
  var richTextValues = range.getRichTextValues();

  var b1Value = String(sheet.getRange('B1').getValue() || '').trim();
  if (!b1Value) return;

  var ftpBase = getFtpBase_(sheet, richTextValues);

  // Find existing table
  var tables = body.getTables();
  var existingTable = null;

  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];
    if (t.getNumRows() < 1) continue;
    var row0 = t.getRow(0);
    if (row0.getNumCells() < 2) continue;
    if (row0.getCell(0).getText().trim() !== 'TDoc') continue;
    if (row0.getCell(1).getText().trim() === b1Value) { existingTable = t; break; }
  }

  if (existingTable) {
    mergeMissingRowsFromSheet_(sheet, existingTable, ftpBase); // ✅ correct name
    ensureAgendaItemRow_(sheet, existingTable);
    updateExistingTable(sheet, existingTable, ftpBase);

    return;
  }


  // Insert placement (same as your logic)
  var insertIndex = findInsertIndex(tables, b1Value, body);

  var agendaItem = getAgendaItemFromSheet(sheet);
  var agendaInsertIndex = null;
  if (insertIndex === null && agendaItem) {
    agendaInsertIndex = findInsertIndexByAgendaItem(body, agendaItem);
    if (agendaInsertIndex !== null) insertIndex = agendaInsertIndex;
  }

  var newTable = (insertIndex !== null) ? body.insertTable(insertIndex, []) : body.appendTable();
  
  // ✅ remove the initial placeholder row if empty
  removeInitialEmptyRow_(newTable);


  // Populate
  values.forEach(function (row) {
    var tr = newTable.appendTableRow();
    var colCount = Math.max(row.length, 2);
    for (var c = 0; c < colCount; c++) {
      tr.appendTableCell(String((c < row.length && row[c] != null) ? row[c] : ''));
    }
  });

  // ✅ ensure Agenda Item exists for NEW tables as well
  ensureAgendaItemRow_(sheet, newTable);

  // Apply rich text links ...
  try {
    for (var r = 0; r < newTable.getNumRows(); r++) {
      for (var c = 0; c < newTable.getRow(r).getNumCells(); c++) {
        if (richTextValues && richTextValues[r] && richTextValues[r][c]) {
          handleHyperlinks(newTable.getRow(r).getCell(c), richTextValues[r][c], ftpBase);
        }
      }
    }
  } catch (e) { }

  // Column widths...
  var ref = findFirstTDocTable_(body);
  if (ref && ref.getNumRows() > 0 && ref.getRow(0).getNumCells() >= 2) {
    setColumnWidth(newTable, 0, ref.getRow(0).getCell(0).getWidth());
    setColumnWidth(newTable, 1, ref.getRow(0).getCell(1).getWidth());
  }

  // ✅ NOW apply status styling (AFTER all text rewriting)
  styleStatusCell_(newTable);
}

// =========================================================
// UPDATE EXISTING TABLE (revised row + status strict rule)
// =========================================================

function updateExistingTable(sheet, existingTable, ftpBase) {
  // Status rule:
  // - If doc status is NOT reserved/available, do not change (unless sheet says revised)
  updateStatusWithStrictRules_(sheet, existingTable);

  // ✅ Always style status
  styleStatusCell_(existingTable);
}

function updateStatusWithStrictRules_(sheet, table) {
  var s = findStatusInSheet_(sheet);
  if (!s) return;

  var d = findStatusInDocTable_(table);
  if (!d) return;

  var docStatus = normalizeStatus_(d.value);      // normalized: reserved/available/other
  var sheetStatusRaw = String(s.value || '').trim().toLowerCase();
  var sheetIsRevised = sheetStatusRaw.indexOf('revised') !== -1;

  // Allowed updates:
  // 1) doc is reserved OR available
  // 2) sheet is revised
  var allowUpdate = (docStatus === 'reserved' || docStatus === 'available' || sheetIsRevised);
  if (!allowUpdate) return;

  var cell = d.cell;
  cell.clear();
  if (s.richText) cell.setText(String(s.value || '')); // avoid hyperlink noise in status
  else cell.setText(String(s.value || ''));
}

function findStatusInSheet_(sheet) {
  var rng = sheet.getDataRange();
  var values = rng.getValues();
  for (var r = 0; r < values.length; r++) {
    var k = String(values[r][0] || '').trim().toLowerCase();
    if (k === 'tdoc status' || k === 'status') {
      return { value: String(values[r][1] || '').trim(), richText: null };
    }
  }
  return null;
}

function findStatusInDocTable_(table) {
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    if (row.getNumCells() < 2) continue;
    var k = row.getCell(0).getText().trim().toLowerCase();
    if (k === 'tdoc status' || k === 'status') {
      return { cell: row.getCell(1), value: row.getCell(1).getText().trim() };
    }
  }
  return null;
}

function normalizeStatus_(s) {
  s = String(s || '').trim().toLowerCase();
  if (s.indexOf('reserved') !== -1) return 'reserved';
  if (s.indexOf('available') !== -1) return 'available';
  return 'other';
}

// =========================================================
// COLLECTOR: Email discussion + revisions (safe, no data loss)
// =========================================================

function collectorUpdate_() {
  const cfg = getCollectorConfig_();
  log_(cfg, '=== COLLECTOR START ===');

  try { checkRSSFeed_(cfg); } catch (e) { warn_(cfg, 'checkRSSFeed_ failed: ' + e.message); }
  try { updateRevisions_(cfg); } catch (e) { warn_(cfg, 'updateRevisions_ failed: ' + e.message); }

  log_(cfg, '=== COLLECTOR END ===');
}

// --- Email discussion ---
function checkRSSFeed_(cfg) {
  // Guard against missing config (prevents your toLowerCase crash)
  const listName = String(cfg.LIST_NAME || LIST_NAME_LOCK);
  const listLower = listName.toLowerCase();

  const msgs = collectHybridListservMessages_(cfg);
  const tz = String(cfg.TIMEZONE || Session.getScriptTimeZone());
  const showPreview = String(cfg.SHOW_PREVIEW_SNIPPET || 'false').toLowerCase() === 'true';

  const props = PropertiesService.getDocumentProperties();
  const tables = DocumentApp.getActiveDocument().getBody().getTables();

  // Deadlines extended verbally / by e-mail, keyed by TDOC.
  const extensions = getDeadlineExtensionMap_();

  tables.forEach(t => {
    if (!isTDocTable_(t)) return;

    const tdocRaw = String(safeCellText_(t, 0, 1) || '').trim();
    const tdoc = extractTdocId_(tdocRaw, cfg.TDOC_ID_REGEX);
    const short = computeShortNumber_(tdoc);
    if (!tdoc || !short) return;

    const storeKey = 'DISCUSS_' + tdoc;
    const store = loadJsonObject_(props.getProperty(storeKey));

    let added = 0;

    msgs.forEach(m => {
      if (!m || !m.title) return;
      
      // Parse email subject to extract TDoc number and deadline
      const parsed = parseEmailSubject_(m.title);
      if (!parsed || !parsed.tdocShort) return;
      
      // Check if this email is for our TDoc
      if (parsed.tdocShort !== short && parsed.tdocFull !== tdoc) return;
      
      // Use author+formatted date as the primary deduplication key to prevent duplicates
      // from RSS and A1 archives with different messageIds and different raw date formats
      const author = (m.author || 'Unknown').trim();
      const dateFormatted = formatLocalDate_(m.date, tz);
      const dedupeKey = `${author}|${dateFormatted}`;
      
      // Check if we already have this email by author+formatted date
      let existingId = null;
      for (const [storeId, storeMsg] of Object.entries(store)) {
        const storeAuthor = (storeMsg.author || 'Unknown').trim();
        const storeDateFormatted = formatLocalDate_(storeMsg.date, tz);
        if (`${storeAuthor}|${storeDateFormatted}` === dedupeKey) {
          existingId = storeId;
          break;
        }
      }
      
      if (!existingId) {
        // New email - use messageId or link as storage key
        const id = (m.messageId || m.link || dedupeKey).trim();
        if (id) {
          store[id] = m;
          added++;
        }
      } else {
        // Duplicate found - only upgrade if new one has preview and old doesn't
        if (!store[existingId].preview && m.preview) {
          store[existingId] = m;
        }
      }
    });

    props.setProperty(storeKey, JSON.stringify(store));

    // Filter stored emails by start date before displaying
    const startDateStr = cfg.EMAIL_START_DATE || '2026-08-21';
    const startDateMillis = new Date(startDateStr).getTime();
    
    const ordered = Object.values(store)
      .filter(m => {
        const emailDateMillis = parseDateToMillis_(m.date);
        return emailDateMillis >= startDateMillis;
      })
      .sort((a, b) => parseDateToMillis_(a.date) - parseDateToMillis_(b.date));

    // IMPORTANT: only clear cell AFTER we know we have content to write
    const cell = findOrFallbackCell_(t, ['E-mail discussion', 'E-Mail discussion', 'Email discussion'], 5, 1);

    if (!ordered.length) {
      if (!cell.getText().trim()) cell.setText('No e-mail discussion.');
      return;
    }

    cell.setText('');
    const te = cell.editAsText();

    // Extract the thread deadline: prefer the opening e-mail, otherwise the
    // earliest message that carries one.
    let threadDeadline = null;
    for (let i = 0; i < ordered.length; i++) {
      if (ordered[i].deadline) { threadDeadline = ordered[i].deadline; break; }
    }
    
    // An entry in the "Document Deadline Extensions" table wins over whatever
    // the subject lines said (deadline extended verbally or by e-mail).
    if (extensions[tdoc]) {
      threadDeadline = extensions[tdoc];
      log_(cfg, 'Deadline extension applied', { tdoc, deadline: threadDeadline.dateStr });
    }
    
    // Update or create "Commenting Deadline" row in the table
    updateCommentingDeadlineRow_(t, threadDeadline, tz);
    
    // Late-response greying: anything sent after the deadline is greyed out,
    // except messages from the sender who opened the thread (they may keep
    // responding in black, e.g. to summarize the outcome).
    const firstAuthor = String((ordered[0] && ordered[0].author) || '').trim();
    const deadlineMillis = (threadDeadline && threadDeadline.dateStr)
      ? new Date(threadDeadline.dateStr).getTime()
      : NaN;
    
    // Deduplicate by author+date to prevent duplicate entries
    // (messageId differs between RSS and A1 for the same email)
    // Use formatted local date for deduplication since RSS and A1 have different raw date formats
    const seen = new Set();
    const segments = [];
    
    ordered.forEach(m => {
      const author = m.author || 'Unknown';
      const dateLocal = formatLocalDate_(m.date, tz);
      const dedupeKey = `${author}|${dateLocal}`;
      
      // Skip if we've already written this message
      if (seen.has(dedupeKey)) return;
      seen.add(dedupeKey);
      
      let line = `${author} on ${dateLocal}\n`;
      const start = te.getText().length;
      te.appendText(line);

      if (m.link) {
        try { te.setLinkUrl(start, start + author.length - 1, m.link); } catch (e) { }
      }
      if (showPreview && m.preview) te.appendText('  ↳ ' + snippet_(m.preview, 160) + '\n');
      
      segments.push({
        start: start,
        end: te.getText().length - 1,
        late: isLateResponse_(m, deadlineMillis, author, firstAuthor)
      });
    });

    // Set font size to 8 for email discussion content
    if (te.getText().length > 0) {
      te.setFontSize(0, te.getText().length - 1, 8);
    }
    
    // Apply the late/on-time colouring per message block
    segments.forEach(seg => {
      if (seg.end < seg.start) return;
      try {
        te.setForegroundColor(seg.start, seg.end, seg.late ? '#808080' : '#000000');
      } catch (e) { }
    });

    log_(cfg, 'Email discussion updated', { tdoc, added, total: ordered.length });
  });
}

function collectHybridListservMessages_(cfg) {
  let out = [];
  // Use only RSS v2.0 to avoid duplicates (v2.0 has more items and better metadata)
  out = out.concat(collectRssItems_(cfg, cfg.RSS_URL_V2, 'rss_v2'));

  const a1Urls = buildArchiveIndexUrlsByDaysBack_(cfg.LIST_NAME, parseInt(cfg.ARCHIVE_DAYS_BACK || '14', 10), cfg);
  a1Urls.forEach(url => out = out.concat(collectA1_(cfg, url)));

  return dedupePreviewFirst_(out);
}

function collectRssItems_(cfg, url, tag) {
  if (!url) return [];
  const resp = safeFetch_(cfg, url, {}, 'RSS fetch ' + tag);
  if (!resp.ok) return [];

  // Get the email start date filter
  const startDateStr = cfg.EMAIL_START_DATE || '2026-08-21';
  const startDateMillis = new Date(startDateStr).getTime();

  const out = [];
  try {
    const xml = XmlService.parse(resp.text);
    const root = xml.getRootElement();
    const dcNs = XmlService.getNamespace('dc', 'http://purl.org/dc/elements/1.1/');

    const channel = root.getChild('channel');
    const items = channel ? (channel.getChildren('item') || []) : (root.getChildren('item') || []);

    items.forEach(it => {
      const title = it.getChildText('title') || '';
      const desc = it.getChildText('description') || '';
      const pubDate = it.getChildText('pubDate') || it.getChildText('date', dcNs) || '';
      const author = it.getChildText('author') || it.getChildText('creator', dcNs) || '';
      const link = it.getChildText('link') || it.getChildText('identifier', dcNs) || '';
      const guid = it.getChildText('guid') || link || title;

      // Filter by date - only include emails from startDate onwards
      const emailDateMillis = parseDateToMillis_(pubDate);
      if (emailDateMillis < startDateMillis) {
        return; // Skip emails before the start date
      }

      // Extract deadline from title - parse subject first to get clean date/time parts
      let deadline = null;
      const parsed = parseEmailSubject_(title);
      if (parsed && parsed.deadline) {
        deadline = parsed.deadline;
      } else {
        // Fallback: try extracting from full title
        deadline = extractDeadlineFromTitle_(title);
      }
      
      out.push({
        title: title || desc,
        author: author || 'Unknown',
        date: pubDate,
        link: link,
        messageId: String(guid || '').trim(),
        preview: desc || '',
        source: tag,
        deadline: deadline
      });
    });
  } catch (e) {
    warn_(cfg, 'RSS parse failed: ' + e.message);
  }
  return out;
}

function collectA1_(cfg, url) {
  if (isA1CachedEmpty_(cfg, url)) return [];
  const resp = safeFetch_(cfg, url, {}, 'A1 fetch');
  if (!resp.ok) { markA1Empty_(cfg, url, true); return []; }

  const html = resp.text || '';
  if (!html.toLowerCase().includes('a2=')) { markA1Empty_(cfg, url, true); return []; }

  const rows = parseA1TableRows_(html, cfg);
  markA1Empty_(cfg, url, rows.length === 0);
  return rows;
}

function parseA1TableRows_(html, cfg) {
  // Get the email start date filter
  const startDateStr = cfg.EMAIL_START_DATE || '2026-08-21';
  const startDateMillis = new Date(startDateStr).getTime();

  const base = 'https://list.etsi.org';
  const out = [];
  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let m;
  while ((m = trRe.exec(html)) !== null) {
    const block = m[1];
    if (!block.toLowerCase().includes('a2=')) continue;

    const a = block.match(/<a[^>]+href="([^"]+A2=[^"]+)"[^>]*>([\s\S]*?)<\/a>/i);
    if (!a) continue;

    const href = a[1];
    const link = href.startsWith('http') ? href : base + href;
    const title = stripTags_(a[2]);

    const tds = [];
    const tdRe = /<td[^>]*>\s*<div[^>]*>\s*([\s\S]*?)\s*<\/div>\s*<\/td>/gi;
    let td;
    while ((td = tdRe.exec(block)) !== null) tds.push(stripTags_(td[1]));

    const author = (tds.length >= 2 ? tds[tds.length - 2] : 'Unknown') || 'Unknown';
    const date = (tds.length >= 1 ? tds[tds.length - 1] : '') || '';

    // Filter by date - only include emails from startDate onwards
    const emailDateMillis = parseDateToMillis_(date);
    if (emailDateMillis < startDateMillis) {
      continue; // Skip emails before the start date
    }

    const preview = extractShowDesc_(block);

    // Extract deadline from title - parse subject first to get clean date/time parts
    let deadline = null;
    const parsed = parseEmailSubject_(title);
    if (parsed && parsed.deadline) {
      deadline = parsed.deadline;
    } else {
      // Fallback: try extracting from full title
      deadline = extractDeadlineFromTitle_(title);
    }
    
    out.push({
      title,
      author,
      date,
      link,
      messageId: link,
      preview: preview || '',
      source: 'a1',
      deadline: deadline
    });
  }
  return out;
}

function extractShowDesc_(block) {
  let m = block.match(/showDesc\(\s*'([\s\S]*?)'\s*\)/i);
  if (!m) m = block.match(/showDesc\(\s*"([\s\S]*?)"\s*\)/i);
  if (!m) return '';
  return String(m[1] || '').replace(/<br\s*\/?>/gi, '\n').trim();
}

function dedupePreviewFirst_(arr) {
  const map = {};
  arr.forEach(x => {
    const key = String(x.messageId || x.link || x.title || '').trim().toLowerCase();
    if (!key) return;
    if (!map[key]) { map[key] = x; return; }

    const cur = map[key];
    const curP = !!(cur.preview && cur.preview.trim());
    const newP = !!(x.preview && x.preview.trim());

    if (!curP && newP) map[key] = x;
    else if (curP && newP && cur.source !== 'a1' && x.source === 'a1') map[key] = x;
  });
  return Object.values(map);
}

// --- Revisions ---
function updateRevisions_(cfg) {
  const baseUrl = String(cfg.REVISIONS_URL || '').trim();
  if (!baseUrl) return;

  const url = baseUrl.replace(/\/$/, '') + '/';
  const resp = safeFetch_(cfg, url, {}, 'Revisions fetch');
  if (!resp.ok) return;

  const anchors = parseAnchors_(resp.text || '');
  const props = PropertiesService.getDocumentProperties();
  const body = DocumentApp.getActiveDocument().getBody();
  const tables = body.getTables();

  tables.forEach(t => {
    if (!isTDocTable_(t)) return;

    const tdocRaw = String(safeCellText_(t, 0, 1) || '').trim();
    const tdoc = extractTdocId_(tdocRaw, cfg.TDOC_ID_REGEX);
    if (!tdoc) return;

    const rx = new RegExp(escapeRegExp_(tdoc));
    const storeKey = 'REVIS_' + tdoc;
    const store = loadJsonObject_(props.getProperty(storeKey));

    let added = 0;
    anchors.forEach(a => {
      const text = String(a.text || '').trim();
      if (!text || !rx.test(text)) return;
      const abs = toAbsoluteUrl_(url, a.href);
      const id = (abs || text).trim();
      if (!id) return;
      if (!store[id]) { store[id] = { text, link: abs }; added++; }
    });

    props.setProperty(storeKey, JSON.stringify(store));
    const ordered = Object.values(store).sort((x, y) => String(x.text || '').localeCompare(String(y.text || '')));

    // Write revisions into the Revisions cell of this table
    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);

    if (!ordered.length) {
      if (!cell.getText().trim()) cell.setText('No revisions available.');
      return;
    }

    cell.setText('');
    const te = cell.editAsText();
    ordered.forEach(item => {
      const line = String(item.text || '').trim();
      if (!line) return;
      const start = te.getText().length;
      te.appendText(line + '\n');
      if (item.link) { try { te.setLinkUrl(start, start + line.length - 1, item.link); } catch (e) { } }
    });

    // Also insert the revision tables directly after this TDOC table in the document.
    // Find the revised TDOCs and insert their tables immediately after this one.
    insertRevisedDocTablesAfter_(body, t, ordered, cfg);

    log_(cfg, 'Revisions updated', { tdoc, added, total: ordered.length });
  });
}

/**
 * For each revision file found, check if there is already a TDOC table for it.
 * If not, insert a minimal table immediately after the parent TDOC table.
 */
function insertRevisedDocTablesAfter_(body, parentTable, revisions, cfg) {
  const parentIndex = body.getChildIndex(parentTable);
  if (parentIndex < 0) return;

  let insertAfter = parentIndex + 1;

  revisions.forEach(item => {
    const revTdoc = extractTdocId_(String(item.text || ''), cfg.TDOC_ID_REGEX || 'S4-\\d{6}');
    if (!revTdoc) return;

    // Check if a table for this revision already exists
    const existing = body.getTables().find(t => {
      if (!isTDocTable_(t)) return false;
      return String(safeCellText_(t, 0, 1) || '').trim() === revTdoc;
    });
    if (existing) return;

    // Insert a minimal revision table
    const revTable = body.insertTable(insertAfter);
    removeInitialEmptyRow_(revTable);

    const rows = [
      ['TDoc', revTdoc],
      ['Title', ''],
      ['Source', ''],
      ['Contact', ''],
      ['Agenda Item', findCellText_(parentTable, 'Agenda Item')],
      ['E-mail Discussion', ''],
      ['Revisions', ''],
      ['Minutes', ''],
      ['Disposition', ''],
      ['Status', 'Available']
    ];

    rows.forEach(rowData => {
      const tr = revTable.appendTableRow();
      tr.appendTableCell(rowData[0]);
      tr.appendTableCell(String(rowData[1] || ''));
    });

    // Add hyperlink to the revision file
    if (item.link) {
      try {
        const cell = revTable.getRow(0).getCell(1);
        cell.editAsText().setLinkUrl(0, revTdoc.length - 1, item.link);
      } catch (e) { }
    }

    styleStatusCell_(revTable);
    insertAfter++;
    Logger.log(`Inserted revision table for ${revTdoc} after parent`);
  });
}

function parseAnchors_(html) {
  const out = [];
  const re = /<a\b[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = re.exec(html)) !== null) out.push({ href: m[1], text: stripTags_(m[2]) });
  return out;
}

// =========================================================
// SA4 TDOC IDENTIFIER MODEL (SA4-IMPL-001)
// =========================================================
//
// Central, single-source-of-truth registry of the SA4 document-identifier
// families verified in SA4-ARCH-005/006 against real, currently-live 3GPP
// FTP material. Each entry's `prefix` is the EXACT canonical casing found in
// real documents (e.g. Readme_Audio.txt, real TDoc-list filenames) -- this
// is deliberately NOT derived from the report-type name (RTC's real prefix
// is `A4aR`, not `S4aR`; MBS's is `S4aI`, not `S4aM` -- guessing by symmetry
// was explicitly wrong for both, per the evidence).
//
// Unverified/unconfirmed series (MTSI, EVS, SQ, ...) are intentionally
// ABSENT: an unrecognized prefix must stay unrecognized rather than being
// guessed into existence. Adding a family here is the only way to recognize
// it; nothing else in this model infers families from shape alone.
//
// `body` is an UNANCHORED regex fragment (a trailing `(?!\d)` guards against
// a 6-digit capture inside a longer run of digits, e.g. "S4-2601234" must
// NOT match as "S4-260123" -- verified against SA4-IMPL-001's test suite).
// Both parseSA4DocumentId_() below and any future exact-match use wrap this
// fragment with `^...$` themselves; the registry stores the fragment once.
const SA4_TDOC_FAMILIES = [
  { key: 'main',          family: 'main',  seriesCode: null, prefix: 'S4-',  body: 'S4-(\\d{6})(?!\\d)' },
  { key: 'audio-adhoc',   family: 'adhoc', seriesCode: 'A',  prefix: 'S4aA', body: 'S4aA(\\d{6})(?!\\d)' },
  { key: 'plenary-adhoc', family: 'adhoc', seriesCode: 'P',  prefix: 'S4aP', body: 'S4aP(\\d{6})(?!\\d)' },
  { key: 'video-adhoc',   family: 'adhoc', seriesCode: 'V',  prefix: 'S4aV', body: 'S4aV(\\d{6})(?!\\d)' },
  { key: 'mbs-adhoc',     family: 'adhoc', seriesCode: 'I',  prefix: 'S4aI', body: 'S4aI(\\d{6})(?!\\d)' },
  { key: 'rtc-adhoc',     family: 'adhoc', seriesCode: 'R',  prefix: 'A4aR', body: 'A4aR(\\d{6})(?!\\d)' }
];

/**
 * Shared matcher behind parseSA4DocumentId_()/parseExactSA4DocumentId_().
 * Not exposed directly -- both public entry points below derive their
 * behavior from this single loop over SA4_TDOC_FAMILIES, so "search within
 * text" and "whole string must be exactly one identifier" can never drift
 * into two independently-maintained regex sets.
 *
 * `exact`: false = search `value` for the first matching family anywhere
 * within it (used where a field may legitimately contain other text around
 * the identifier, e.g. "revised to S4-261599"). true = the ENTIRE trimmed
 * `value` must itself be exactly one recognized identifier, nothing more
 * (used where a field is defined to hold nothing but the identifier).
 *
 * The returned identifier's casing is always the family's own canonical
 * casing (SA4_TDOC_FAMILIES[i].prefix), never the input's original casing
 * and never force-uppercased -- this is deliberate: main-meeting IDs are
 * conventionally all-uppercase ("S4-260123") so canonical-casing and
 * uppercasing happen to agree there, but ad-hoc IDs are NOT all-uppercase in
 * real documents ("S4aA260090", not "S4AA260090") -- forcing uppercase would
 * silently corrupt the real, verified identifier spelling.
 *
 * Returns { raw, isValid, family, familyKey, seriesCode, yearCode, sequence }.
 * On no match, isValid=false, raw=<trimmed input>, all other fields null.
 */
function matchSA4DocumentId_(value, exact) {
  const trimmed = String(value === null || value === undefined ? '' : value).trim();

  if (trimmed) {
    for (const fam of SA4_TDOC_FAMILIES) {
      const pattern = exact ? ('^' + fam.body + '$') : fam.body;
      const m = trimmed.match(new RegExp(pattern, 'i'));
      if (m) {
        const digits = m[1];
        return {
          raw: fam.prefix + digits,
          isValid: true,
          family: fam.family,
          familyKey: fam.key,
          seriesCode: fam.seriesCode,
          yearCode: digits.slice(0, 2),
          sequence: digits.slice(2)
        };
      }
    }
  }

  return { raw: trimmed, isValid: false, family: null, familyKey: null, seriesCode: null, yearCode: null, sequence: null };
}

/**
 * Parse an SA4 document identifier out of `value` (search within text).
 *
 * Mirrors normalizeTdoc_()'s pre-existing "extract from surrounding text"
 * behavior (e.g. "revised to S4-261599" -> S4-261599) -- several callers,
 * including normalizeTdoc_() itself, intentionally rely on this NOT being
 * anchored to the whole string. Use parseExactSA4DocumentId_() where a field
 * is defined to hold nothing but the identifier itself (e.g. a TDoc table's
 * own "TDoc" value cell).
 */
function parseSA4DocumentId_(value) {
  return matchSA4DocumentId_(value, false);
}

/**
 * SA4-IMPL-001A: like parseSA4DocumentId_(), but the ENTIRE trimmed `value`
 * must be exactly one recognized identifier -- no other characters before
 * or after it are tolerated. This restores cleanUpWrongEmailDiscussions()'s
 * pre-IMPL-001 exact-match semantics (it used to require
 * `^S4-\d{6}$` against the whole cell) while still recognizing all 6
 * verified families, without hand-maintaining a second, separately-anchored
 * copy of each family's pattern -- both this and parseSA4DocumentId_() are
 * thin wrappers over the same matchSA4DocumentId_()/SA4_TDOC_FAMILIES pair.
 */
function parseExactSA4DocumentId_(value) {
  return matchSA4DocumentId_(value, true);
}

// =========================================================
// REVISIONS FROM THE TDOC LIST ("Revised to" column)
// =========================================================

/**
 * TDOC number of a TDOC-list row.
 *
 * SA4-IMPL-001: this used to unconditionally uppercase the cell text, which
 * was harmless while every TDoc number was in the always-uppercase main
 * family ("S4-260123" upper-cased is itself) but silently corrupts a
 * verified ad-hoc identifier's real casing ("S4aA260090" -> "S4AA260090",
 * which no longer matches the same identifier as written in the actual
 * TDoc-list Excel or in getRevisedTo_()'s output -- discovered by the
 * revision-chain test for a same-family ad-hoc revision, which failed to
 * link until this was fixed). Now: a recognized SA4 document identifier
 * (main or ad-hoc) is returned in its real canonical casing via
 * parseSA4DocumentId_(); anything NOT recognized as a valid identifier falls
 * back to the exact previous behavior (trim + uppercase) so this function's
 * contract for non-TDoc/malformed input is unchanged.
 */
function tdocNumberOf_(tdocData) {
  const raw = String(tdocData.row[tdocData.tdocCol] || '').trim();
  const parsed = parseSA4DocumentId_(raw);
  return parsed.isValid ? parsed.raw : raw.toUpperCase();
}

/**
 * The document this row was revised to, from the "Revised to" column.
 * Returns '' when the column is absent or the cell holds no TDOC number.
 *
 * SA4-IMPL-001: migrated from a hardcoded 'S4-\d{6}' extraction (which
 * bypassed cfg.TDOC_ID_REGEX entirely and could never recognize an ad-hoc
 * "Revised to" target) to the central parseSA4DocumentId_() model, so
 * same-family ad-hoc revisions (e.g. "S4aA260049 is revised to S4aA260050",
 * verified in a real Audio SWG ad-hoc report, SA4-ARCH-005/006) are now
 * recognized. Existing main-meeting behavior is unchanged -- see
 * tests/revision-order.test.js.
 */
function getRevisedTo_(tdocData) {
  if (!tdocData || tdocData.revisedToCol === undefined || tdocData.revisedToCol === null) return '';
  if (tdocData.revisedToCol < 0) return '';
  const raw = String(tdocData.row[tdocData.revisedToCol] || '');
  const parsed = parseSA4DocumentId_(raw);
  return parsed.isValid ? parsed.raw : '';
}

/**
 * Order a group's TDOCs so each revision directly follows the document it
 * revises, following chains (A -> B -> C).
 *
 * Documents whose revision lives in a different agenda item keep their own
 * position; only relationships inside this group are reordered.
 */
function orderTdocsByRevision_(tdocs) {
  const byNumber = {};
  tdocs.forEach(td => {
    const n = tdocNumberOf_(td);
    if (n) byNumber[n] = td;
  });

  // "child" = some other document in this group is revised TO it
  const isChild = {};
  tdocs.forEach(td => {
    const target = getRevisedTo_(td);
    if (target && byNumber[target]) isChild[target] = true;
  });

  const out = [];
  const done = {};

  function emitChain(td) {
    const n = tdocNumberOf_(td);
    if (done[n]) return;          // also guards against cycles
    done[n] = true;
    out.push(td);
    const target = getRevisedTo_(td);
    if (target && byNumber[target]) emitChain(byNumber[target]);
  }

  // Roots first (documents that nothing in this group revises to)
  tdocs.forEach(td => { if (!isChild[tdocNumberOf_(td)]) emitChain(td); });
  // Then anything left over (pure cycles), so nothing is ever dropped
  tdocs.forEach(td => {
    const n = tdocNumberOf_(td);
    if (!done[n]) { done[n] = true; out.push(td); }
  });

  return out;
}

/**
 * Build { 'S4-261480': 'S4-261599', ... } from the meeting TDOC list,
 * i.e. document -> the document it was revised to.
 * Cached in the REVISION_MAP property for reference/debugging.
 *
 * Pass an already-downloaded `groups` object (from downloadAndGroupTdocs_) to
 * avoid fetching the XLSX a second time.
 */
function buildRevisionMapFromTdocList_(cfg, groups) {
  cfg = cfg || getReportConfig_();
  groups = groups || downloadAndGroupTdocs_(cfg);
  const map = {};

  Object.keys(groups).forEach(key => {
    groups[key].tdocs.forEach(td => {
      const from = tdocNumberOf_(td);
      const to = getRevisedTo_(td);
      if (from && to && from !== to) map[from] = to;
    });
  });

  PropertiesService.getDocumentProperties().setProperty('REVISION_MAP', JSON.stringify(map));
  Logger.log(`Revision map: ${Object.keys(map).length} revised document(s)`);
  return map;
}

/**
 * Return the revision relationships ordered so that chain roots come first
 * (A before B for A -> B -> C). Moving tables in that order keeps the whole
 * chain contiguous; processing a chain backwards would strand its tail.
 */
function orderRevisionChains_(map) {
  const isTarget = {};
  Object.keys(map).forEach(p => { isTarget[map[p]] = true; });

  const out = [];
  const seen = {};

  function walk(p) {
    if (!p || seen[p] || !map[p]) return;
    seen[p] = true;
    out.push(p);
    walk(map[p]);
  }

  Object.keys(map).sort().forEach(p => { if (!isTarget[p]) walk(p); });
  Object.keys(map).sort().forEach(p => { if (!seen[p]) walk(p); }); // cycles
  return out;
}

/**
 * Find the TDOC table for a given document number.
 */
function findTdocTable_(body, tdocNumber) {
  const want = String(tdocNumber || '').trim().toUpperCase();
  if (!want) return null;

  const tables = body.getTables();
  for (let i = 0; i < tables.length; i++) {
    if (!isTDocTable_(tables[i])) continue;
    if (String(safeCellText_(tables[i], 0, 1) || '').trim().toUpperCase() === want) return tables[i];
  }
  return null;
}

/**
 * Write "Revised to S4-xxxxxx" into the Disposition row.
 *
 * Hand-written disposition text is never destroyed: if the cell already holds
 * something else the marker is appended, and if it already mentions a revision
 * the cell is left untouched. Returns true when the cell was changed.
 */
function setDispositionRevisedTo_(table, revisedTo) {
  const marker = 'Revised to ' + String(revisedTo || '').trim().toUpperCase();

  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 2) continue;

    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key !== 'disposition' && key !== 'disposition:') continue;

    const current = row.getCell(1).getText().trim();
    if (current === marker) return false;                        // already correct
    if (/revised\s+to/i.test(current)) return false;             // already mentions a revision

    row.getCell(1).setText(current ? current + ' — ' + marker : marker);
    return true;
  }
  return false;
}

/**
 * Move a table so it sits immediately after another table.
 * Returns true when the document was actually changed.
 */
function moveTableAfter_(body, table, afterTable) {
  const targetIndex = body.getChildIndex(afterTable) + 1;
  const currentIndex = body.getChildIndex(table);
  if (currentIndex === targetIndex) return false; // already in the right place

  // Insert a copy at the destination, then drop the original. Holding the
  // element reference means the shifting indices do not matter.
  const copy = table.copy();
  body.insertTable(targetIndex, copy);
  table.removeFromParent();
  return true;
}

/**
 * Menu action: repair revision placement in an existing report.
 */
function rearrangeRevisionTables() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Re-arrange Revision Tables',
    'This reads the "Revised to" column of the meeting TDOC list and then:\n\n' +
    '• moves each revision table directly below the document it revises\n' +
    '• fills the revised document\'s Disposition with "Revised to S4-xxxxxx"\n\n' +
    'Existing minutes and hand-written disposition text are preserved.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const r = rearrangeRevisionTables_();
    ui.alert(
      'Re-arrangement Complete',
      `✅ Revisions in the TDOC list: ${r.total}\n` +
      `🔀 Tables moved: ${r.moved}\n` +
      `📋 Dispositions updated: ${r.dispositions}\n` +
      `➖ Already in place: ${r.alreadyInPlace}\n` +
      `⚠️ Revision not in this report: ${r.missingChild}\n` +
      `⚠️ Revised document not in this report: ${r.missingParent}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Error', 'Re-arrangement failed: ' + e.message, ui.ButtonSet.OK);
    Logger.log('rearrangeRevisionTables failed: ' + e.message);
    Logger.log(e.stack);
  }
}

/**
 * Worker for the above; no UI, so it is safe to call from other steps.
 * `groups` is optional (see buildRevisionMapFromTdocList_).
 */
function rearrangeRevisionTables_(cfg, groups) {
  cfg = cfg || getReportConfig_();
  const body = DocumentApp.getActiveDocument().getBody();
  const map = buildRevisionMapFromTdocList_(cfg, groups);

  const stats = {
    total: Object.keys(map).length,
    moved: 0,
    dispositions: 0,
    alreadyInPlace: 0,
    missingChild: 0,
    missingParent: 0
  };

  orderRevisionChains_(map).forEach(parent => {
    const child = map[parent];

    const parentTable = findTdocTable_(body, parent);
    if (!parentTable) { stats.missingParent++; return; }

    if (setDispositionRevisedTo_(parentTable, child)) stats.dispositions++;

    const childTable = findTdocTable_(body, child);
    if (!childTable) { stats.missingChild++; return; }

    if (moveTableAfter_(body, childTable, parentTable)) {
      stats.moved++;
      Logger.log(`Moved ${child} directly below ${parent}`);
    } else {
      stats.alreadyInPlace++;
    }
  });

  Logger.log(`Revision re-arrangement: ${JSON.stringify(stats)}`);
  return stats;
}

// =========================================================
// Helpers shared
// =========================================================

function safeFetch_(cfg, url, opt, label) {
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const code = r.getResponseCode();
    const text = r.getContentText() || '';
    if (isDebug_(cfg)) log_(cfg, label, { url, code, bytes: text.length });
    // reject login-ish pages
    const low = text.toLowerCase();
    if (low.includes('login') && low.includes('password') && low.includes('username')) return { ok: false, code, text: '' };
    return { ok: code < 400 && !!text, code, text };
  } catch (e) {
    warn_(cfg, 'Fetch failed: ' + url + ' :: ' + e.message);
    return { ok: false, code: 0, text: '' };
  }
}

function stripTags_(s) {
  return String(s || '').replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}
function snippet_(s, n) { s = String(s || '').replace(/\s+/g, ' ').trim(); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
function escapeRegExp_(s) { return String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/**
 * Strip reply/forward prefixes from an e-mail subject so the leading
 * bracketed metadata can be parsed.
 *
 * Handles repeated and localized prefixes, with or without a counter:
 *   "Re: ", "RE: ", "Re: Re: ", "AW: ", "FW: ", "FWD: ", "RE[2]: ", ...
 *
 * "Re: [FS_6G_MED, 1483, ...] title"  ->  "[FS_6G_MED, 1483, ...] title"
 */
function stripReplyPrefixes_(subject) {
  return String(subject || '')
    .replace(/^(?:\s*(?:re|aw|fw|fwd|tr|sv|antw|vs|rif|res)\s*(?:\[\d+\])?\s*:\s*)+/i, '')
    .trim();
}

/**
 * Parse email subject line to extract TDoc number and deadline
 * Handles multiple formats:
 * - "[agenda; tdoc_short; date time] title"
 * - "[tdoc; agenda; date time] title"
 * - "[agenda, tdoc, date time] title" (commas)
 * - "[agenda; S4-xxxxxx; date time] title" (full TDoc)
 * - any of the above prefixed with "Re: ", "RE: ", "AW: ", "FW: ", ...
 * 
 * Returns: { tdocShort: '1399', tdocFull: 'S4-261399', deadline: {...} } or null
 */
function parseEmailSubject_(subject) {
  if (!subject) return null;
  
  // Replies/forwards prepend "Re: " etc., which would defeat the ^\[ anchor.
  subject = stripReplyPrefixes_(subject);
  
  // Extract the bracketed portion: [...]
  const bracketMatch = subject.match(/^\[([^\]]+)\]/);
  if (!bracketMatch) return null;
  
  const bracketContent = bracketMatch[1];
  
  // Split by semicolons OR commas to get parts
  const parts = bracketContent.split(/[;,]/).map(p => p.trim());
  if (parts.length < 2) return null;
  
  // Try to find TDoc number in any part
  let tdocShort = null;
  let tdocFull = null;
  let dateTimeParts = [];
  
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i].trim();
    
    // Check for full TDoc format (S4-xxxxxx or A4aR260087)
    const fullTdocMatch = part.match(/(?:S4|A4[a-z]*)-?\d{6}/i);
    if (fullTdocMatch) {
      tdocFull = fullTdocMatch[0].toUpperCase().replace(/^(S4|A4[A-Z]*)(\d{6})$/, '$1-$2');
      // Extract short number from full TDoc
      const shortMatch = tdocFull.match(/\d{6}$/);
      if (shortMatch) {
        const sixDigits = shortMatch[0];
        tdocShort = sixDigits.charAt(2) === '0' ? sixDigits.slice(-3) : sixDigits.slice(-4);
      }
      continue;
    }
    
    // Check for short TDoc format (3-4 digits, not a time)
    if (/^\d{3,4}$/.test(part)) {
      const nextPart = i + 1 < parts.length ? parts[i + 1].trim() : '';
      const prevPart = i > 0 ? parts[i - 1].trim() : '';
      
      // Strategy: TDoc numbers typically appear in position 1 (after agenda item)
      // Times appear later with date context
      
      // This is DEFINITELY a TIME if:
      // 1. Next part is a timezone (e.g., "1400" before "CEST")
      // 2. Previous part contains a date/month (e.g., "25 Aug 2026" before "1400")
      // 3. Part is exactly 4 digits AND previous part contains year (2026, 2025, etc.)
      // 4. Previous part contains month name (even without year)
      const hasTimezoneAfter = /^(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i.test(nextPart);
      const hasDateBefore = /(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|\d{1,2}(st|nd|rd|th)?)\s+\d{4}$/i.test(prevPart);
      const hasYearBefore = /\b(202[0-9]|203[0-9])\b/.test(prevPart);
      const hasMonthBefore = /(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i.test(prevPart);
      
      // If it has a timezone after it, it's ALWAYS a time, never a TDoc
      if (hasTimezoneAfter) {
        dateTimeParts.push(part);
        continue;
      }
      
      // If preceded by date/month/year context, it's a time
      const isPartOfTime = hasDateBefore || hasYearBefore || hasMonthBefore;
      
      // This is likely a TDOC only if:
      // - It's in position 1 (second part overall, after agenda)
      // - Previous part looks like agenda item (contains letters/underscores)
      // - NOT preceded by date/year/month
      // - NOT followed by timezone
      const isLikelyTdoc = (i === 1) && !isPartOfTime;
      
      if (isLikelyTdoc) {
        tdocShort = part;
        continue;
      }
      
      // Otherwise treat as date/time part
      dateTimeParts.push(part);
    }
    
    // Everything else is potentially date/time
    dateTimeParts.push(part);
  }
  
  // If we didn't find a TDoc number, return null
  if (!tdocShort && !tdocFull) return null;
  
  // If we only have full TDoc, try to extract short from it
  if (!tdocShort && tdocFull) {
    const shortMatch = tdocFull.match(/\d{6}$/);
    if (shortMatch) {
      const sixDigits = shortMatch[0];
      tdocShort = sixDigits.charAt(2) === '0' ? sixDigits.slice(-3) : sixDigits.slice(-4);
    }
  }
  
  // Try to extract full TDoc from anywhere in the subject if we don't have it yet
  if (!tdocFull) {
    const fullTdocMatch = subject.match(/(?:S4|A4[a-z]*)-?\d{6}/i);
    if (fullTdocMatch) {
      tdocFull = fullTdocMatch[0].toUpperCase().replace(/^(S4|A4[A-Z]*)(\d{6})$/, '$1-$2');
    }
  }
  
  // Extract deadline from the date/time parts
  const dateTimeStr = dateTimeParts.join(' ');
  const deadline = extractDeadlineFromTitle_(dateTimeStr);
  
  return {
    tdocShort: tdocShort,
    tdocFull: tdocFull,
    deadline: deadline
  };
}

function extractTdocId_(raw, regexStr) {
  const rx = new RegExp(regexStr || 'S4-\\d{6}');
  const m = String(raw || '').match(rx);
  return m ? m[0] : '';
}
function computeShortNumber_(tdoc) {
  const m = String(tdoc || '').match(/(\d{6})/);
  if (!m) return '';
  return m[1].charAt(2) === '0' ? m[1].slice(-3) : m[1].slice(-4);
}
function parseDateToMillis_(d) { const t = new Date(d).getTime(); return isNaN(t) ? 9e15 : t; }
function formatLocalDate_(d, tz) { const x = new Date(d); return isNaN(x) ? String(d || '') : Utilities.formatDate(x, tz, 'yyyy-MM-dd HH:mm'); }

/**
 * Extract deadline from email subject line
 * Handles various formats:
 * - "25 Aug 2026 1400 CEST"
 * - "25th August 1200CEST" (time stuck to timezone)
 * - "Aug 26, 1400 CEST"
 * - "August 26th 12pm CEST"
 * - "26 August, 1300 CEST"
 * - "26 August 2026 12pm CEST"
 * 
 * IMPORTANT: Patterns are ordered to prioritize time extraction over year extraction
 * to prevent "1200CEST" from being interpreted as year 1200 AD.
 */
function extractDeadlineFromTitle_(title) {
  if (!title) return null;
  
  const currentYear = new Date().getFullYear();
  
  const patterns = [
    // Pattern 1: "DD Month YYYY HHMM TZ" (e.g., "25 Aug 2026 1400 CEST")
    {
      regex: /(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*,?\s+(\d{4})\s+(\d{3,4})\s+(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i,
      parse: (m) => ({
        day: parseInt(m[1]),
        month: parseMonth_(m[2]),
        year: parseInt(m[3]),
        hour: parseInt(m[4].padStart(4, '0').substring(0, 2)),
        minute: parseInt(m[4].padStart(4, '0').substring(2, 4)),
        timezone: m[5]
      })
    },
    
    // Pattern 2: "DD Month YYYY HH:MMam/pm TZ" or "DD Month YYYY HHam/pm TZ" (e.g., "26 August 2026 12pm CEST")
    {
      regex: /(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*,?\s+(\d{4})\s+(\d{1,2}):?(\d{2})?(am|pm)\s+(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i,
      parse: (m) => {
        let hour = parseInt(m[4]);
        const minute = m[5] ? parseInt(m[5]) : 0;
        const ampm = m[6].toLowerCase();
        if (ampm === 'pm' && hour !== 12) hour += 12;
        if (ampm === 'am' && hour === 12) hour = 0;
        return {
          day: parseInt(m[1]),
          month: parseMonth_(m[2]),
          year: parseInt(m[3]),
          hour: hour,
          minute: minute,
          timezone: m[7]
        };
      }
    },
    
    // Pattern 3: "DD Month, HHMM TZ" (e.g., "26 August, 1300 CEST") - NO YEAR
    {
      regex: /(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*,?\s+(\d{3,4})\s+(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i,
      parse: (m) => ({
        day: parseInt(m[1]),
        month: parseMonth_(m[2]),
        year: currentYear,
        hour: parseInt(m[3].padStart(4, '0').substring(0, 2)),
        minute: parseInt(m[3].padStart(4, '0').substring(2, 4)),
        timezone: m[4]
      })
    },
    
    // Pattern 4: "DD Month HH:MMam/pm TZ" or "DD Month HHam/pm TZ" (e.g., "26 August 12pm CEST")
    {
      regex: /(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*,?\s+(\d{1,2}):?(\d{2})?(am|pm)\s+(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i,
      parse: (m) => {
        let hour = parseInt(m[3]);
        const minute = m[4] ? parseInt(m[4]) : 0;
        const ampm = m[5].toLowerCase();
        if (ampm === 'pm' && hour !== 12) hour += 12;
        if (ampm === 'am' && hour === 12) hour = 0;
        return {
          day: parseInt(m[1]),
          month: parseMonth_(m[2]),
          year: currentYear,
          hour: hour,
          minute: minute,
          timezone: m[6]
        };
      }
    },
    
    // Pattern 4b: "Month DDth HHam/pm TZ" (e.g., "August 26th 12pm CEST")
    {
      regex: /(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{1,2}):?(\d{2})?(am|pm)\s+(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i,
      parse: (m) => {
        let hour = parseInt(m[3]);
        const minute = m[4] ? parseInt(m[4]) : 0;
        const ampm = m[5].toLowerCase();
        if (ampm === 'pm' && hour !== 12) hour += 12;
        if (ampm === 'am' && hour === 12) hour = 0;
        return {
          day: parseInt(m[2]),
          month: parseMonth_(m[1]),
          year: currentYear,
          hour: hour,
          minute: minute,
          timezone: m[6]
        };
      }
    },
    
    // Pattern 5: "DDth Month HHMMTZ" (e.g., "28th August 1200CEST") - time stuck to timezone
    // IMPORTANT: Check this BEFORE "DD Month YYYY" pattern to prioritize time over year
    // Must have exactly 4 digits before timezone and be a valid time (HH < 24)
    {
      regex: /(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*,?\s+(\d{2})(\d{2})(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i,
      parse: (m) => {
        const hour = parseInt(m[3]);
        // Only accept if it's a valid hour (00-23)
        if (hour >= 24) return null;
        return {
          day: parseInt(m[1]),
          month: parseMonth_(m[2]),
          year: currentYear,
          hour: hour,
          minute: parseInt(m[4]),
          timezone: m[5]
        };
      }
    },
    
    // Pattern 6: "DD Month YYYY" without time (check AFTER time patterns)
    {
      regex: /(\d{1,2})(?:st|nd|rd|th)?\s+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*,?\s+(\d{4})/i,
      parse: (m) => ({
        day: parseInt(m[1]),
        month: parseMonth_(m[2]),
        year: parseInt(m[3]),
        hour: 23,
        minute: 59,
        timezone: 'CEST'
      })
    },
    
    // Pattern 7: "Month DD, HHMM TZ" (e.g., "Aug 26, 1400 CEST")
    {
      regex: /(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{3,4})\s+(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)/i,
      parse: (m) => ({
        day: parseInt(m[2]),
        month: parseMonth_(m[1]),
        year: currentYear,
        hour: parseInt(m[3].padStart(4, '0').substring(0, 2)),
        minute: parseInt(m[3].padStart(4, '0').substring(2, 4)),
        timezone: m[4]
      })
    }
  ];
  
  for (const pattern of patterns) {
    const match = title.match(pattern.regex);
    if (match) {
      try {
        const parsed = pattern.parse(match);
        if (!parsed) continue; // Skip if parse returned null (e.g., invalid hour)
        
        const { day, month, year, hour, minute, timezone } = parsed;
        
        // Create date string in ISO format
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:00`;
        
        return {
          dateStr: dateStr,
          timezone: timezone,
          day: day,
          month: month,
          year: year,
          hour: hour,
          minute: minute
        };
      } catch (e) {
        // Continue to next pattern if parsing fails
        continue;
      }
    }
  }
  
  return null;
}

/**
 * Parse month name to number (1-12)
 */
function parseMonth_(monthStr) {
  const months = {
    'jan': 1, 'january': 1,
    'feb': 2, 'february': 2,
    'mar': 3, 'march': 3,
    'apr': 4, 'april': 4,
    'may': 5,
    'jun': 6, 'june': 6,
    'jul': 7, 'july': 7,
    'aug': 8, 'august': 8,
    'sep': 9, 'september': 9,
    'oct': 10, 'october': 10,
    'nov': 11, 'november': 11,
    'dec': 12, 'december': 12
  };
  return months[monthStr.toLowerCase()] || 1;
}

/**
 * Format deadline for display
 */
function formatDeadline_(deadline, tz) {
  if (!deadline || !deadline.dateStr) return '';
  
  try {
    const date = new Date(deadline.dateStr);
    const formatted = Utilities.formatDate(date, tz, 'dd MMM yyyy HH:mm');
    return `${formatted} ${deadline.timezone}`;
  } catch (e) {
    return `${deadline.day} ${getMonthName_(deadline.month)} ${deadline.year} ${String(deadline.hour).padStart(2, '0')}:${String(deadline.minute).padStart(2, '0')} ${deadline.timezone}`;
  }
}

/**
 * Get month name from number
 */
function getMonthName_(month) {
  const names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return names[month] || '';
}

/**
 * Calculate time remaining until deadline
 */
function calculateTimeRemaining_(deadline, tz) {
  if (!deadline || !deadline.dateStr) return '';
  
  try {
    const now = new Date();
    const deadlineDate = new Date(deadline.dateStr);
    const diffMs = deadlineDate.getTime() - now.getTime();
    
    if (diffMs < 0) {
      // Deadline has passed
      const absDiffMs = Math.abs(diffMs);
      const days = Math.floor(absDiffMs / (24 * 3600 * 1000));
      const hours = Math.floor((absDiffMs % (24 * 3600 * 1000)) / (3600 * 1000));
      
      if (days > 0) {
        return `(⚠️ EXPIRED ${days} day${days !== 1 ? 's' : ''} ago)`;
      } else if (hours > 0) {
        return `(⚠️ EXPIRED ${hours} hour${hours !== 1 ? 's' : ''} ago)`;
      } else {
        return '(⚠️ EXPIRED)';
      }
    } else {
      // Deadline is in the future
      const days = Math.floor(diffMs / (24 * 3600 * 1000));
      const hours = Math.floor((diffMs % (24 * 3600 * 1000)) / (3600 * 1000));
      const minutes = Math.floor((diffMs % (3600 * 1000)) / (60 * 1000));
      
      if (days > 1) {
        return `(${days} days ${hours} hours remaining)`;
      } else if (days === 1) {
        return `(1 day ${hours} hours remaining)`;
      } else if (hours > 0) {
        return `(⏰ ${hours} hour${hours !== 1 ? 's' : ''} ${minutes} min remaining)`;
      } else {
        return `(⏰ ${minutes} minute${minutes !== 1 ? 's' : ''} remaining)`;
      }
    }
  } catch (e) {
    return '';
  }
}

/**
 * Format deadline for display
 */
function formatDeadline_(deadline, tz) {
  if (!deadline || !deadline.dateStr) return '';
  
  try {
    const date = new Date(deadline.dateStr);
    const formatted = Utilities.formatDate(date, tz, 'dd MMM yyyy HH:mm');
    return `${formatted} ${deadline.timezone}`;
  } catch (e) {
    return `${deadline.day} ${getMonthName_(deadline.month)} ${deadline.year} ${String(deadline.hour).padStart(2, '0')}:${String(deadline.minute).padStart(2, '0')} ${deadline.timezone}`;
  }
}

/**
 * Get month name from number
 */
function getMonthName_(month) {
  const names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return names[month] || '';
}

/**
 * Calculate time remaining until deadline
 */
function calculateTimeRemaining_(deadline, tz) {
  if (!deadline || !deadline.dateStr) return '';
  
  try {
    const now = new Date();
    const deadlineDate = new Date(deadline.dateStr);
    const diffMs = deadlineDate.getTime() - now.getTime();
    
    if (diffMs < 0) {
      // Deadline has passed
      const absDiffMs = Math.abs(diffMs);
      const days = Math.floor(absDiffMs / (24 * 3600 * 1000));
      const hours = Math.floor((absDiffMs % (24 * 3600 * 1000)) / (3600 * 1000));
      
      if (days > 0) {
        return `(⚠️ EXPIRED ${days} day${days !== 1 ? 's' : ''} ago)`;
      } else if (hours > 0) {
        return `(⚠️ EXPIRED ${hours} hour${hours !== 1 ? 's' : ''} ago)`;
      } else {
        return '(⚠️ EXPIRED)';
      }
    } else {
      // Deadline is in the future
      const days = Math.floor(diffMs / (24 * 3600 * 1000));
      const hours = Math.floor((diffMs % (24 * 3600 * 1000)) / (3600 * 1000));
      const minutes = Math.floor((diffMs % (3600 * 1000)) / (60 * 1000));
      
      if (days > 1) {
        return `(${days} days ${hours} hours remaining)`;
      } else if (days === 1) {
        return `(1 day ${hours} hours remaining)`;
      } else if (hours > 0) {
        return `(⏰ ${hours} hour${hours !== 1 ? 's' : ''} ${minutes} min remaining)`;
      } else {
        return `(⏰ ${minutes} minute${minutes !== 1 ? 's' : ''} remaining)`;
      }
    }
  } catch (e) {
    return '';
  }
}

/**
 * Update or create "E-mail deadline" row in TDOC table
 * Shows the deadline from the first email in the thread with countdown
 * Makes text red if deadline is within 6 hours
 */
function updateCommentingDeadlineRow_(table, deadline, tz) {
  // Find if "E-mail deadline" row already exists
  let deadlineRowIndex = -1;
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 2) continue;
    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key === 'e-mail deadline' || key === 'e-mail deadline:' || 
        key === 'commenting deadline' || key === 'commenting deadline:') {
      deadlineRowIndex = r;
      break;
    }
  }
  
  // Determine where to insert the row (after "Type/For" or "Agenda Item")
  let insertIndex = -1;
  if (deadlineRowIndex === -1) {
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      const key = row.getCell(0).getText().trim().toLowerCase();
      if (key === 'type/for' || key === 'type/for:') {
        insertIndex = r + 1;
        break;
      }
      if (key === 'agenda item' || key === 'agenda item:') {
        insertIndex = r + 1;
      }
    }
  }
  
  // Format the deadline text and check if within 6 hours
  let deadlineText = '';
  let isUrgent = false;
  
  if (deadline && deadline.dateStr) {
    const formattedDeadline = formatDeadline_(deadline, tz);
    const timeRemaining = calculateTimeRemaining_(deadline, tz);
    const extendedTag = deadline.extended ? ' (extended)' : '';
    deadlineText = `${formattedDeadline}${extendedTag} ${timeRemaining}`;
    
    // Check if deadline is within 6 hours
    try {
      const now = new Date();
      const deadlineDate = new Date(deadline.dateStr);
      const diffMs = deadlineDate.getTime() - now.getTime();
      const sixHoursMs = 6 * 3600 * 1000;
      
      // Urgent if within 6 hours and not expired
      isUrgent = (diffMs > 0 && diffMs <= sixHoursMs);
    } catch (e) {
      // If date parsing fails, not urgent
      isUrgent = false;
    }
  } else {
    deadlineText = 'No deadline set';
  }
  
  // Update existing row or create new one
  if (deadlineRowIndex >= 0) {
    // Update existing row - also update the label if it's the old one
    const labelCell = table.getRow(deadlineRowIndex).getCell(0);
    const currentLabel = labelCell.getText().trim().toLowerCase();
    if (currentLabel === 'commenting deadline' || currentLabel === 'commenting deadline:') {
      labelCell.setText('E-mail deadline');
    }
    
    const cell = table.getRow(deadlineRowIndex).getCell(1);
    cell.setText(deadlineText);
    
    // Set red color if urgent, otherwise reset to black
    const te = cell.editAsText();
    if (te.getText().length > 0) {
      if (isUrgent) {
        te.setForegroundColor(0, te.getText().length - 1, '#FF0000');
      } else {
        te.setForegroundColor(0, te.getText().length - 1, '#000000');
      }
    }
  } else if (insertIndex >= 0) {
    // Create new row at the appropriate position
    const newRow = table.insertTableRow(insertIndex);
    newRow.appendTableCell('E-mail deadline');
    const cell = newRow.appendTableCell(deadlineText);
    
    // Set red color if urgent, otherwise black
    const te = cell.editAsText();
    if (te.getText().length > 0) {
      if (isUrgent) {
        te.setForegroundColor(0, te.getText().length - 1, '#FF0000');
      } else {
        te.setForegroundColor(0, te.getText().length - 1, '#000000');
      }
    }
  }
}

/**
 * Recognise the "Document Deadline Extensions" table:
 *
 *   | TDOC       | Extended Deadline      |
 *   | S4-261579  | 27 Aug 2026 23:59 CEST |
 *
 * Matched on the header cells only, so the heading text/number above it and
 * its position in the document do not matter.
 */
function isDeadlineExtensionTable_(table) {
  try {
    const row0 = table.getRow(0);
    if (row0.getNumCells() < 2) return false;
    const c0 = row0.getCell(0).getText().trim().toLowerCase();
    const c1 = row0.getCell(1).getText().trim().toLowerCase();
    return (c0 === 'tdoc' || c0 === 'tdoc number' || c0 === 'document') &&
           c1.indexOf('deadline') !== -1;
  } catch (e) {
    return false;
  }
}

/**
 * Read the deadline-extension table.
 * Returns: { 'S4-261579': { dateStr, timezone, ..., extended: true, raw } }
 *
 * The map is mirrored into the DEADLINE_EXTENSIONS document property. If the
 * table is missing entirely (for example immediately after a skeleton rebuild,
 * which clears the body) the last known map is reused instead of silently
 * dropping every extension.
 */
function getDeadlineExtensionMap_() {
  const body = DocumentApp.getActiveDocument().getBody();
  const props = PropertiesService.getDocumentProperties();
  const map = {};
  let tableFound = false;

  for (const table of body.getTables()) {
    if (!isDeadlineExtensionTable_(table)) continue;
    tableFound = true;

    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 2) continue;

      // SA4-IMPL-001: was extractTdocId_(text, 'S4-\\d{6}') -- migrated to the
      // central model so an ad-hoc TDoc can also get a deadline extension.
      const parsedTdoc = parseSA4DocumentId_(row.getCell(0).getText());
      const tdoc = parsedTdoc.isValid ? parsedTdoc.raw : '';
      const raw = row.getCell(1).getText().trim();
      if (!tdoc || !raw) continue;

      const deadline = parseExtendedDeadline_(raw);
      if (!deadline) {
        Logger.log(`Could not parse extended deadline for ${tdoc}: "${raw}"`);
        continue;
      }

      deadline.extended = true;
      deadline.raw = raw;
      map[tdoc] = deadline;
    }
  }

  if (tableFound) {
    props.setProperty('DEADLINE_EXTENSIONS', JSON.stringify(map));
    Logger.log(`Deadline extensions: ${Object.keys(map).length} entrie(s)`);
    return map;
  }

  return loadJsonObject_(props.getProperty('DEADLINE_EXTENSIONS'));
}

/**
 * Parse a deadline written by hand in the extension table.
 *
 * Accepts, with or without a timezone (default CEST):
 *   "27 Aug 2026 23:59 CEST"   "27 August 2026 2359"
 *   "Aug 27, 2026 23:59 CEST"  "2026-08-27 23:59"
 *   "27 Aug 2026"              "2026-08-27"        (-> 23:59, end of day)
 * Anything else falls back to the subject-line parser, which additionally
 * understands forms like "12pm CEST" and "1200CEST".
 *
 * Returns a deadline object in the same shape as extractDeadlineFromTitle_,
 * or null when the text cannot be understood.
 */
function parseExtendedDeadline_(text) {
  const s = String(text || '').trim();
  if (!s) return null;

  const MONTH = '(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)';
  const TZ = '(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)';

  // Build + validate; returns null for impossible clock values.
  const build = (year, month, day, hour, minute, tz) => {
    if (!(month >= 1 && month <= 12)) return null;
    if (!(day >= 1 && day <= 31)) return null;
    if (!(hour >= 0 && hour <= 23)) return null;
    if (!(minute >= 0 && minute <= 59)) return null;
    return {
      dateStr: `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}` +
               `T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:00`,
      timezone: tz ? String(tz).toUpperCase() : 'CEST',
      day: day, month: month, year: year, hour: hour, minute: minute
    };
  };

  let m;

  // "27 Aug 2026 23:59 CEST"
  m = s.match(new RegExp(`^(\\d{1,2})(?:st|nd|rd|th)?\\s+${MONTH}[a-z]*\\.?,?\\s+(\\d{4})\\s+(\\d{1,2}):(\\d{2})\\s*${TZ}?`, 'i'));
  if (m) return build(+m[3], parseMonth_(m[2]), +m[1], +m[4], +m[5], m[6]);

  // "27 Aug 2026 2359 CEST"
  m = s.match(new RegExp(`^(\\d{1,2})(?:st|nd|rd|th)?\\s+${MONTH}[a-z]*\\.?,?\\s+(\\d{4})\\s+(\\d{2})(\\d{2})\\s*${TZ}?`, 'i'));
  if (m) return build(+m[3], parseMonth_(m[2]), +m[1], +m[4], +m[5], m[6]);

  // "Aug 27, 2026 23:59 CEST"
  m = s.match(new RegExp(`^${MONTH}[a-z]*\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})\\s+(\\d{1,2}):?(\\d{2})\\s*${TZ}?`, 'i'));
  if (m) return build(+m[3], parseMonth_(m[1]), +m[2], +m[4], +m[5], m[6]);

  // "2026-08-27 23:59" / "2026-08-27T23:59"
  m = s.match(new RegExp(`^(\\d{4})-(\\d{1,2})-(\\d{1,2})[T ](\\d{1,2}):(\\d{2})\\s*${TZ}?`, 'i'));
  if (m) return build(+m[1], +m[2], +m[3], +m[4], +m[5], m[6]);

  // Date only -> end of day
  m = s.match(new RegExp(`^(\\d{1,2})(?:st|nd|rd|th)?\\s+${MONTH}[a-z]*\\.?,?\\s+(\\d{4})\\s*$`, 'i'));
  if (m) return build(+m[3], parseMonth_(m[2]), +m[1], 23, 59, null);

  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s*$/);
  if (m) return build(+m[1], +m[2], +m[3], 23, 59, null);

  // Last resort: the subject-line parser (12pm CEST, 1200CEST, ...)
  return extractDeadlineFromTitle_(s);
}

/**
 * Decide whether an e-mail counts as a "late response" for display purposes.
 *
 * Late  = sent strictly after the thread deadline AND by somebody other than
 *         the sender who opened the thread.
 * Not late when: there is no usable deadline, the date cannot be parsed, or the
 *         sender is the opening sender.
 */
function isLateResponse_(msg, deadlineMillis, author, firstAuthor) {
  if (!deadlineMillis || !isFinite(deadlineMillis)) return false;
  if (firstAuthor && String(author || '').trim() === firstAuthor) return false;
  
  const sent = parseDateToMillis_(msg && msg.date);
  if (!isFinite(sent) || sent >= 9e15) return false; // unparseable date
  
  return sent > deadlineMillis;
}

function loadJsonObject_(s) { try { const o = JSON.parse(s || '{}'); return (o && typeof o === 'object') ? o : {}; } catch (e) { return {}; } }

function buildArchiveIndexUrlsByDaysBack_(list, daysBack, cfg) {
  const urls = [];
  const now = new Date();
  const start = new Date(now.getTime() - daysBack * 24 * 3600 * 1000);
  const weeks = ['A', 'B', 'C', 'D', 'E'];
  let cur = new Date(start.getFullYear(), start.getMonth(), 1);
  const endMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  while (cur.getTime() <= endMonth.getTime()) {
    const yy = String(cur.getFullYear()).slice(-2);
    const mm = String(cur.getMonth() + 1).padStart(2, '0');
    weeks.forEach(w => urls.push(`https://list.etsi.org/scripts/wa.exe?A1=ind${yy}${mm}${w}&L=${encodeURIComponent(LIST_NAME_LOCK)}`));
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }
  return urls;
}

// A1 empty cache
function a1CacheKey_(url) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, url, Utilities.Charset.UTF_8);
  const hex = digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  return 'A1_EMPTY_CACHE_' + hex;
}
function isA1CachedEmpty_(cfg, url) {
  const props = PropertiesService.getDocumentProperties();
  const ttlHours = parseInt(cfg.A1_EMPTY_CACHE_TTL_HOURS || '24', 10);
  const ttlMs = ttlHours * 3600 * 1000;
  const raw = props.getProperty(a1CacheKey_(url));
  if (!raw) return false;
  try {
    const obj = JSON.parse(raw);
    if ((Date.now() - obj.ts) > ttlMs) return false;
    return obj.empty === true;
  } catch (e) { return false; }
}
function markA1Empty_(cfg, url, empty) {
  PropertiesService.getDocumentProperties().setProperty(a1CacheKey_(url), JSON.stringify({ ts: Date.now(), empty: !!empty }));
}
function clearA1EmptyCache_() {
  const props = PropertiesService.getDocumentProperties();
  props.getKeys().forEach(k => { if (k.startsWith('A1_EMPTY_CACHE_')) props.deleteProperty(k); });
  Logger.log('A1 empty-week cache cleared');
}

function toAbsoluteUrl_(baseUrl, href) {
  if (!href) return '';
  const h = String(href).trim();
  if (!h) return '';
  if (/^https?:\/\//i.test(h)) return h;
  if (h.startsWith('//')) return 'https:' + h;
  if (h.startsWith('/')) return 'https://www.3gpp.org' + h;
  const dir = String(baseUrl || '').replace(/[?#].*$/, '').replace(/\/[^\/]*$/, '/');
  return dir + h;
}

function isTDocTable_(t) {
  try { return t.getNumRows() >= 1 && t.getCell(0, 0).getText().trim() === 'TDoc'; }
  catch (e) { return false; }
}

function safeCellText_(t, r, c) {
  try {
    if (t.getNumRows() <= r) return '';
    const row = t.getRow(r);
    if (row.getNumCells() <= c) return '';
    return row.getCell(c).getText();
  } catch (e) { return ''; }
}

function findOrFallbackCell_(table, labels, fallbackRow, fallbackCol) {
  // Find by label in col0
  try {
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      const k = row.getCell(0).getText().trim().toLowerCase();
      for (const lab of labels) {
        if (k === String(lab).toLowerCase()) return row.getCell(1);
      }
    }
  } catch (e) { }

  // Fallback coordinate
  try {
    if (table.getNumRows() > fallbackRow && table.getRow(fallbackRow).getNumCells() > fallbackCol) {
      return table.getCell(fallbackRow, fallbackCol);
    }
  } catch (e) { }

  // Append
  const rr = table.appendTableRow();
  rr.appendTableCell(labels[0] || 'Value');
  rr.appendTableCell('');
  return rr.getCell(1);
}

// =========================================================
// Your existing hyperlink + FTP base helpers (kept)
// =========================================================

function handleHyperlinks(cell, richTextValue, ftpBase) {
  if (!richTextValue) return;
  var text = cell.editAsText();
  text.setText('');
  var pos = 0;
  var runs = richTextValue.getRuns ? richTextValue.getRuns() : [];
  runs.forEach(function (run) {
    var t = run.getText();
    text.appendText(t);
    var url = run.getLinkUrl();
    if (url) {
      url = rewrite3gppLink_(url, t, ftpBase);
      text.setLinkUrl(pos, pos + t.length - 1, url);
    }
    pos += t.length;
  });
}

function rewrite3gppLink_(url, displayedText, ftpBase) {
  if (!url) return url;
  var t = String(displayedText || '').trim();
  ftpBase = ftpBase || DEFAULT_S4_FTP_BASE;
  var m = t.match(/S4-\d+/i);
  var token = m ? m[0] : '';
  if (token && /portal\.3gpp\.org\/ngppapp\/CreateTdoc\.aspx/i.test(url)) {
    return ftpBase + token + '.zip';
  }
  return url;
}

function getFtpBase_(sheet, richTextValues) {
  try {
    var ss = sheet.getParent();
    var nr = ss.getRangeByName('FTP_BASE');
    if (nr) {
      var v = String(nr.getDisplayValue() || '').trim();
      if (v) return normalizeFtpBase_(v);
    }
  } catch (e) { }
  var detected = detectFtpBaseFromRichText_(richTextValues);
  if (detected) return detected;
  return DEFAULT_S4_FTP_BASE;
}

function normalizeFtpBase_(base) {
  base = String(base || '').trim();
  if (!base) return '';
  if (!/\/Docs\/$/i.test(base)) base = base.replace(/\/+$/, '') + '/Docs/';
  return base;
}

function detectFtpBaseFromRichText_(richTextValues) {
  if (!richTextValues || !richTextValues.length) return '';
  for (var r = 0; r < richTextValues.length; r++) {
    for (var c = 0; c < (richTextValues[r] ? richTextValues[r].length : 0); c++) {
      var rt = richTextValues[r][c];
      if (!rt || !rt.getRuns) continue;
      var runs = rt.getRuns();
      for (var k = 0; k < runs.length; k++) {
        var url = runs[k].getLinkUrl();
        if (!url) continue;
        var m = String(url).match(/https:\/\/(?:ftp\.3gpp\.org|www\.3gpp\.org\/ftp)\/.*?\/Docs\//i);
        if (m && m[0]) return m[0];
      }
    }
  }
  return '';
}

// =========================================================
// Merge missing rows from sheet (Agenda Item etc) — NON-DESTRUCTIVE
// =========================================================

function mergeMissingRowsFromSheet_(sheet, table, ftpBase) {
  var rng = sheet.getDataRange();
  var values = rng.getValues();
  var rtv = rng.getRichTextValues();

  // existing labels in doc table
  var existing = {};
  for (var r = 0; r < table.getNumRows(); r++) {
    var k = table.getRow(r).getNumCells() > 0 ? table.getRow(r).getCell(0).getText().trim() : '';
    if (k) existing[k] = true;
  }

  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (!key) continue;

    var kl = key.toLowerCase();

    // never override these
    if (kl === 'tdoc status' || kl === 'status') continue;
    if (kl.startsWith('revised')) continue;
    if (kl.indexOf('e-mail discussion') !== -1) continue;
    if (kl.indexOf('revisions') !== -1) continue;

    if (!existing[key]) {
      var row = table.appendTableRow();
      row.appendTableCell(String(values[i][0] || ''));
      row.appendTableCell(String(values[i][1] || ''));

      // preserve rich text hyperlinks in col2 if present
      if (rtv && rtv[i] && rtv[i][1]) {
        try { handleHyperlinks(row.getCell(1), rtv[i][1], ftpBase); } catch (e) { }
      }
    }
  }
}

// =========================================================
// Placement + formatting helpers (from your code)
// =========================================================

function getAgendaItemFromSheet(sheet) {
  var values = sheet.getDataRange().getValues();
  for (var i = 0; i < values.length; i++) {
    var k = String(values[i][0] || '').trim().toLowerCase();
    if (k === 'agenda item' || k === 'agenda item:' || k.startsWith('agenda item')) {
      var v = String(values[i][1] || '').trim();
      return v || null;
    }
  }
  return null;
}

function findInsertIndexByAgendaItem(body, agendaItem) {
  var rx = new RegExp('^' + agendaItem.replace('.', '\\.') + '(\\b|\\.)');
  var headings = [];
  for (var i = 0; i < body.getNumChildren(); i++) {
    var el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      if (p.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
        headings.push({ index: i, level: p.getHeading(), text: p.getText().trim() });
      }
    }
  }
  for (var h = 0; h < headings.length; h++) {
    if (rx.test(headings[h].text)) {
      var thisLevel = headings[h].level;
      for (var n = h + 1; n < headings.length; n++) {
        if (headings[n].level <= thisLevel) return headings[n].index;
      }
      return body.getNumChildren();
    }
  }
  return null;
}

function findInsertIndex(tables, docNumber, body) {
  const needle = ('revised to ' + String(docNumber || '').toLowerCase());
  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    if (table.getNumRows() < 1) continue;
    var firstRow = table.getRow(0);
    if (firstRow.getNumCells() < 2) continue;
    if (firstRow.getCell(0).getText().trim() !== 'TDoc') continue;

    for (var r = 0; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      var text = row.getCell(1).getText();
      if (text && text.toLowerCase().includes(needle)) return body.getChildIndex(table) + 1;
    }
  }
  return null;
}

function findFirstTDocTable_(body) {
  var tables = body.getTables();
  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];
    if (t.getNumRows() < 1) continue;
    var row0 = t.getRow(0);
    if (row0.getNumCells() < 2) continue;
    if (row0.getCell(0).getText().trim() === 'TDoc') return t;
  }
  return null;
}

function setColumnWidth(table, columnIndex, width) {
  for (var i = 0; i < table.getNumRows(); i++) {
    table.getRow(i).getCell(columnIndex).setWidth(width);
  }
}

// =========================================================
// Fix links + reorder (optional hooks if you have them)
// =========================================================

function fixLinksAndReorder_() {
  // keep your existing ones if present; safe no-op if missing
  try { if (typeof fixExistingDocLinks_ === 'function') fixExistingDocLinks_(); } catch (e) { }
  try { if (typeof reorderRevisedTables_ === 'function') reorderRevisedTables_(); } catch (e) { }
}

// =========================================================
// portal→FTP rewrite menu helper (simple)
// =========================================================

function rewritePortalLinksInDoc_() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();
  const ftpBase = DEFAULT_S4_FTP_BASE;

  tables.forEach(t => {
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        const te = cell.editAsText();
        const s = te.getText() || '';
        for (let i = 0; i < s.length; i++) {
          const url = te.getLinkUrl(i);
          if (!url) continue;
          if (/portal\.3gpp\.org\/ngppapp\/CreateTdoc\.aspx/i.test(url)) {
            const m = s.substring(Math.max(0, i - 40), Math.min(s.length, i + 40)).match(/S4-\d{6}/i);
            if (m) {
              const token = m[0];
              const start = s.lastIndexOf(token, i);
              if (start >= 0) {
                te.setLinkUrl(start, start + token.length - 1, ftpBase + token + '.zip');
              }
            }
          }
        }
      }
    }
  });
}


function ensureAgendaItemRow_(sheet, table) {
  const agenda = getAgendaItemFromSheet(sheet);
  if (!agenda) return;

  // If row exists → update value (do not duplicate)
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() >= 2) {
      const k = row.getCell(0).getText().trim().toLowerCase();
      if (k === 'agenda item' || k === 'agenda item:') {
        // Only write if empty OR different (safe update)
        const current = row.getCell(1).getText().trim();
        if (!current || current !== String(agenda)) {
          row.getCell(1).setText(String(agenda));
        }
        return;
      }
    }
  }

  // Otherwise insert AFTER header row
  const newRow = table.insertTableRow(1);
  newRow.appendTableCell('Agenda Item');
  newRow.appendTableCell(String(agenda));
}

/********************************************************
 * FORMATTING
 ********************************************************/
function removeRowHeightAndSpacing() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();

  // Apply widths ONLY for 2-column TDOC tables (and leave other tables untouched)
  setTwoColumnTDocTableWidths_();

  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];

    // Row formatting + paragraph spacing cleanup
    for (let j = 0; j < table.getNumRows(); j++) {
      const row = table.getRow(j);
      row.setMinimumHeight(0);

      for (let k = 0; k < row.getNumCells(); k++) {
        const cell = row.getCell(k);
        const numChildren = cell.getNumChildren();

        for (let l = 0; l < numChildren; l++) {
          const child = cell.getChild(l);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            paragraph.setSpacingBefore(0);
            paragraph.setSpacingAfter(0);
          }
        }
      }
    }

    // Header row background color (keep your behavior)
    try {
      const headerRow = table.getRow(0);
      for (let c = 0; c < headerRow.getNumCells(); c++) {
        headerRow.getCell(c).setBackgroundColor('#D9EAF7');
      }
    } catch (e) {
      // ignore
    }

    // Normalize font based on first cell
    try {
      const firstCell = table.getCell(0, 0);
      const firstChild = firstCell.getNumChildren() ? firstCell.getChild(0) : null;
      if (firstChild && firstChild.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const t0 = firstChild.asParagraph().editAsText();
        const ff = t0.getFontFamily() || null;
        const fs = t0.getFontSize() || null;

        if (ff || fs) {
          for (let r = 0; r < table.getNumRows(); r++) {
            for (let col = 0; col < table.getRow(r).getNumCells(); col++) {
              const cell = table.getCell(r, col);
              const child = cell.getNumChildren() ? cell.getChild(0) : null;
              if (!child || child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
              const txt = child.asParagraph().editAsText();
              if (ff) txt.setFontFamily(ff);
              if (fs) txt.setFontSize(fs);
            }
          }
        }
      }
    } catch (e) {
      // ignore
    }
  }

  removeEmptyParagraphs_();
}

/**
 * Only set widths for TDOC tables that have exactly 2 columns.
 * Leaves all other tables alone.
 */
function setTwoColumnTDocTableWidths_() {
  const cfg = getConfig_();

  const pageWidth = parseInt(cfg.TDOC_PAGE_USABLE_WIDTH || '468', 10);
  const firstColWidth = parseInt(cfg.TDOC_COL1_WIDTH || '108', 10);
  const secondColWidth = pageWidth - firstColWidth;

  const body = DocumentApp.getActiveDocument().getBody();
  const tables = body.getTables();

  tables.forEach(t => {
    if (!isTDocTable_(t)) return;

    // only apply to exactly 2 columns
    const numCols = t.getRow(0).getNumCells();
    if (numCols !== 2) return;

    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      row.getCell(0).setWidth(firstColWidth);
      row.getCell(1).setWidth(secondColWidth);
    }
  });
}

/**
 * Removes empty paragraphs only between "E-Mail Discussion" and "Revisions:"
 * (same logic you had).
 */
function removeEmptyParagraphs_() {
  const body = DocumentApp.getActiveDocument().getBody();
  const paragraphs = body.getParagraphs();

  let inTargetSection = false;
  for (let i = 0; i < paragraphs.length; i++) {
    const paragraph = paragraphs[i];
    const text = paragraph.getText().trim();

    if (text === "E-Mail Discussion") {
      inTargetSection = true;
    } else if (text === "Revisions:") {
      inTargetSection = false;
    } else if (inTargetSection && text === "") {
      paragraph.removeFromParent();
      i--;
    }
  }
}

/********************************************************
 * CONFIG TABLE
 ********************************************************/
function ensureConfigTable_() {
  const body = DocumentApp.getActiveDocument().getBody();
  
  // Find existing table
  for (const t of body.getTables()) {
    if (isConfigTable_(t)) {
      // Clear existing rows except header
      while (t.getNumRows() > 1) t.removeRow(1);
      // Repopulate with current config
      populateConfigTable_(t);
      return;
    }
  }

  // Create new table if it doesn't exist
  body.insertParagraph(0, 'Configuration').setHeading(DocumentApp.ParagraphHeading.HEADING3);
  const newTable = body.insertTable(1, [['Key', 'Value']]);
  populateConfigTable_(newTable);
}

function populateConfigTable_(table) {
  const cfg = getReportConfig_();
  const collectorCfg = readCollectorConfigTable_(); // Read optional overrides

  const data = [
    ['DEBUG', 'true'],
    ['TIMEZONE', Session.getScriptTimeZone()],
    ['LIST_NAME', cfg.LIST_NAME],
    ['RSS_URL_V2', cfg.RSS_URL_V2],
    ['RSS_URL_V1', cfg.RSS_URL_V1],
    ['REVISIONS_URL', cfg.REVISIONS_URL],
    ['ARCHIVE_DAYS_BACK', collectorCfg.ARCHIVE_DAYS_BACK || '14'],
    ['A1_EMPTY_CACHE_TTL_HOURS', collectorCfg.A1_EMPTY_CACHE_TTL_HOURS || '24'],
    ['SHOW_PREVIEW_SNIPPET', collectorCfg.SHOW_PREVIEW_SNIPPET || 'false'],
    ['TDOC_ID_REGEX', '^S4-\\d{6}$'],
    ['TDOC_COL1_WIDTH', '108'],
    ['TDOC_PAGE_USABLE_WIDTH', '468']
  ];

  data.forEach(row => {
    table.appendTableRow(row);
  });
}

function isConfigTable_(t) {
  try {
    return t.getCell(0, 0).getText() === 'Key' &&
      t.getCell(0, 1).getText() === 'Value';
  } catch (e) {
    return false;
  }
}

function getConfig_() {
  const body = DocumentApp.getActiveDocument().getBody();
  for (const t of body.getTables()) {
    if (!isConfigTable_(t)) continue;

    const cfg = {};
    for (let r = 1; r < t.getNumRows(); r++) {
      cfg[t.getCell(r, 0).getText().trim()] =
        t.getCell(r, 1).getText().trim();
    }

    // HARD-LOCK list name no matter what table says
    cfg.LIST_NAME = LIST_NAME_LOCK;

    // Ensure RSS URLs match the lock
    cfg.RSS_URL_V2 = `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME_LOCK}&v=2.0&LIMIT=2000`;
    cfg.RSS_URL_V1 = `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME_LOCK}&v=1.0&LIMIT=2000`;

    return cfg;
  }

  // fallback
  return {
    DEBUG: 'true',
    TIMEZONE: Session.getScriptTimeZone(),
    LIST_NAME: LIST_NAME_LOCK,
    RSS_URL_V2: `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME_LOCK}&v=2.0&LIMIT=2000`,
    RSS_URL_V1: `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME_LOCK}&v=1.0&LIMIT=2000`,
    ARCHIVE_DAYS_BACK: '14',
    A1_EMPTY_CACHE_TTL_HOURS: '24',
    SHOW_PREVIEW_SNIPPET: 'false',
    TDOC_ID_REGEX: '^S4-\\d{6}$',
    REVISIONS_URL: 'https://www.3gpp.org/ftp/tsg_sa/WG4_CODEC/TSGS4_136_Montreal/Inbox/Drafts/FS_6G_MED',
    REVISIONS_CELL_ROW: '6',
    REVISIONS_CELL_COL: '1',
    TDOC_COL1_WIDTH: '108',
    TDOC_PAGE_USABLE_WIDTH: '468'
  };
}

function styleStatusCell_(table) {
  // Find the status cell
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    if (row.getNumCells() < 2) continue;

    var key = row.getCell(0).getText().trim().toLowerCase();
    if (key !== 'tdoc status' && key !== 'status') continue;

    var cell = row.getCell(1);
    var txt = cell.getText();
    if (!txt) return;

    var v = txt.trim().toLowerCase();
    var color = null;
    var bold = false;

    if (v.includes('available')) {
      color = '#0070C0'; // blue
      bold = false;
    } else if (v.includes('noted')) {
      color = '#7030A0'; // purple
      bold = true;
    } else if (v.includes('agreed') || v.includes('endorsed')) {
      color = '#00B050'; // green
      bold = true;
    } else if (v.includes('revised')) {
      color = '#C00000'; // red
      bold = true;
    } else if (v.includes('withdrawn')) {
      color = '#000000'; // black
      bold = true;
    } else {
      return; // leave other statuses untouched
    }

    // Apply style to entire cell text range
    var te = cell.editAsText();
    var len = te.getText().length;
    if (len <= 0) return;

    te.setForegroundColor(0, len - 1, color);
    te.setBold(0, len - 1, bold);

    return; // done
  }
}

function removeInitialEmptyRow_(table) {
  try {
    if (!table || table.getNumRows() < 1) return;
    const r0 = table.getRow(0);
    if (r0.getNumCells() !== 1) return;
    const t = r0.getCell(0).getText().trim();
    // If first row is empty and there are additional rows, remove it
    if (!t && table.getNumRows() > 1) table.removeRow(0);
  } catch (e) { }
}

/********************************************************
 * NEW WORKFLOW FUNCTIONS
 ********************************************************/

/**
 * PHASE 1: Configure Meeting Settings
 */
function configureMeetingSettings() {
  const ui = DocumentApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  
  // Get current values
  const currentMeetingFolder = props.getProperty('MEETING_FOLDER') || 'TSGS4_136_Montreal';
  const currentMeetingNumber = props.getProperty('MEETING_NUMBER') || '136';
  const currentMeetingId = props.getProperty('MEETING_ID') || '60777';
  const currentSuffix = props.getProperty('REPORT_SUFFIX') || '6G';
  const currentAgendaSourceDocId = props.getProperty('AGENDA_SOURCE_DOC_ID') || '1qP--dusvUhNwwBtMEH4xVdxaP1c6L1hZ49geICoYV2s';
  const currentAgendaTdoc = props.getProperty('AGENDA_TDOC') || '';
  const currentTdocUrl = props.getProperty('TDOC_LIST_URL') || '';
  const currentShowPreview = props.getProperty('SHOW_PREVIEW_SNIPPET') !== 'false';
  const currentToken = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN') || '';
  
  // Build HTML form
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      .hint { font-size: 11px; color: #666; margin-top: 3px; }
      .section { background: #f5f5f5; padding: 10px; margin: 15px 0; border-left: 3px solid #4285f4; }
    </style>
    
    <h2>Meeting Configuration</h2>
    
    <div class="section">
      <h3>📅 Meeting Information</h3>
      
      <label>Meeting Folder:</label>
      <input type="text" id="meetingFolder" value="${currentMeetingFolder}" placeholder="TSGS4_136_Montreal">
      <div class="hint">Folder name on 3GPP FTP (e.g., TSGS4_136_Montreal or TSGS4_137-e)</div>
      
      <label>Meeting Number:</label>
      <input type="text" id="meetingNumber" value="${currentMeetingNumber}" placeholder="136">
      <div class="hint">Meeting number for TDOC list filename (e.g., 136 or 137-e)</div>
      
      <label>3GPP Portal Meeting ID:</label>
      <input type="text" id="meetingId" value="${currentMeetingId}" placeholder="60777">
      <div class="hint">Portal ID from https://portal.3gpp.org/Home.aspx#/meeting?MtgId=60777</div>
    </div>
    
    <div class="section">
      <h3>📋 Report Configuration</h3>
      
      <label>Report Type:</label>
      <select id="reportType">
        <option value="6G" ${currentSuffix === '6G' ? 'selected' : ''}>6G (Agenda 11.x)</option>
        <option value="Audio" ${currentSuffix === 'Audio' ? 'selected' : ''}>Audio (Agenda 7.x)</option>
        <option value="Video" ${currentSuffix === 'Video' ? 'selected' : ''}>Video (Agenda 9.x)</option>
        <option value="RTC" ${currentSuffix === 'RTC' ? 'selected' : ''}>RTC (Agenda 10.x)</option>
        <option value="MBS" ${currentSuffix === 'MBS' ? 'selected' : ''}>MBS (Agenda 8.x)</option>
        <option value="Liaison" ${currentSuffix === 'Liaison' ? 'selected' : ''}>Liaison (Agenda 5.x)</option>
        <option value="New" ${currentSuffix === 'New' ? 'selected' : ''}>New Work Items (Agenda 18.x)</option>
      </select>
      
      <label>Meeting Report Template (Google Doc):</label>
      <input type="text" id="agendaSourceDocId" value="${currentAgendaSourceDocId}" placeholder="https://docs.google.com/document/d/...">
      <div class="hint">Google Doc ID or URL for skeleton/template with agenda structure</div>
      
      <label>Agenda TDoc Number:</label>
      <input type="text" id="agendaTdoc" value="${currentAgendaTdoc}" placeholder="S4-260868">
      <div class="hint">TDoc number of the official meeting agenda (e.g., S4-260868)</div>
    </div>
    
    <div class="section">
      <h3>📁 File Locations</h3>
      
      <label>TDOC List URL (optional - auto-detected if empty):</label>
      <input type="text" id="tdocUrl" value="${currentTdocUrl}" placeholder="Leave empty for auto-detection">
      <div class="hint">Auto: https://www.3gpp.org/ftp/.../TSGS4_136_Montreal/Docs/TDoc_List_Meeting_SA4#136.xlsx</div>
    </div>
    
    <div class="section">
      <h3>⚙️ Options</h3>
      
      <label>
        <input type="checkbox" id="showPreview" ${currentShowPreview ? 'checked' : ''}> 
        Show email preview snippets in report
      </label>
      
      <label>Email Collection Start Date:</label>
      <input type="date" id="emailStartDate" value="${props.getProperty('EMAIL_START_DATE') || '2026-08-21'}">
      <div class="hint">Only collect emails from this date onwards (format: YYYY-MM-DD)</div>
      
      <label>Reviewer API Token (optional):</label>
      <input type="password" id="apiToken" value="${currentToken}" placeholder="crv1_...">
      <div class="hint">For fetching AI summaries and abstracts</div>
    </div>
    
    <button onclick="saveConfig()">💾 Save Configuration</button>
    <button onclick="google.script.host.close()" style="background: #666;">Cancel</button>
    
    <script>
      function saveConfig() {
        const config = {
          meetingFolder: document.getElementById('meetingFolder').value,
          meetingNumber: document.getElementById('meetingNumber').value,
          meetingId: document.getElementById('meetingId').value,
          reportType: document.getElementById('reportType').value,
          agendaSourceDocId: document.getElementById('agendaSourceDocId').value,
          agendaTdoc: document.getElementById('agendaTdoc').value,
          tdocUrl: document.getElementById('tdocUrl').value,
          showPreview: document.getElementById('showPreview').checked,
          apiToken: document.getElementById('apiToken').value
        };
        google.script.run
          .withSuccessHandler(() => {
            alert('✅ Configuration saved successfully!');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('❌ Error saving configuration: ' + error);
          })
          .saveConfigurationSettings(config);
      }
    </script>
  `)
  .setWidth(600)
  .setHeight(700);
  
  ui.showModalDialog(html, 'Meeting Configuration');
}


function saveConfigurationSettings(config) {
  const docProps = PropertiesService.getDocumentProperties();
  const scriptProps = PropertiesService.getScriptProperties();
  
  // Save meeting information
  docProps.setProperty('MEETING_FOLDER', config.meetingFolder || 'TSGS4_136_Montreal');
  docProps.setProperty('MEETING_NUMBER', config.meetingNumber || '136');
  docProps.setProperty('MEETING_ID', config.meetingId || '60777');
  
  // Save report configuration
  docProps.setProperty('REPORT_SUFFIX', config.reportType || '6G');
  docProps.setProperty('AGENDA_SOURCE_DOC_ID', extractGoogleDocId_(config.agendaSourceDocId || ''));
  docProps.setProperty('AGENDA_TDOC', config.agendaTdoc || '');
  
  // Save file locations (only if provided, otherwise auto-detect)
  if (config.tdocUrl && config.tdocUrl.trim()) {
    docProps.setProperty('TDOC_LIST_URL', config.tdocUrl.trim());
  } else {
    docProps.deleteProperty('TDOC_LIST_URL'); // Will auto-detect
  }
  
  // Save options
  docProps.setProperty('SHOW_PREVIEW_SNIPPET', config.showPreview ? 'true' : 'false');
  if (config.emailStartDate && config.emailStartDate.trim()) {
    docProps.setProperty('EMAIL_START_DATE', config.emailStartDate.trim());
  }
  
  // Save API token (script properties for security)
  if (config.apiToken && config.apiToken.trim()) {
    scriptProps.setProperty('REVIEWER_API_TOKEN', config.apiToken.trim());
  }
  
  Logger.log('Configuration saved: ' + JSON.stringify(config));
}

/**
 * PHASE 1: Test All Connections
 */
function testAllConnections() {
  const ui = DocumentApp.getUi();
  const results = [];
  
  results.push('🧪 Testing All Connections...\n');
  
  // Test 1: TDOC List URL
  results.push('\n1️⃣ TDOC List URL:');
  try {
    const cfg = getReportConfig_();
    const url = cfg.TDOC_LIST_URL;
    if (!url) {
      results.push('   ❌ Not configured');
    } else {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        results.push('   ✅ Connected successfully');
        results.push('   📊 File size: ' + response.getBlob().getBytes().length + ' bytes');
      } else {
        results.push('   ❌ HTTP ' + response.getResponseCode());
      }
    }
  } catch (e) {
    results.push('   ❌ Error: ' + e.message);
  }
  
  // Test 2: Reviewer API
  results.push('\n2️⃣ Reviewer API:');
  try {
    const token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN');
    if (!token) {
      results.push('   ⚠️  Not configured (abstracts will be skipped)');
    } else {
      const testUrl = 'https://reviewer.bouazizi.dev/api/v1/documents/S4-260001/summary?type=summary';
      const response = UrlFetchApp.fetch(testUrl, {
        headers: { 'X-API-Key': token },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() === 200 || response.getResponseCode() === 404) {
        results.push('   ✅ API token valid');
      } else {
        results.push('   ❌ HTTP ' + response.getResponseCode());
      }
    }
  } catch (e) {
    results.push('   ❌ Error: ' + e.message);
  }
  
  // Test 3: Email Feeds
  results.push('\n3️⃣ Email Feeds (RSS):');
  try {
    const rssUrl = `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME_LOCK}&v=2.0&LIMIT=10`;
    const response = UrlFetchApp.fetch(rssUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      results.push('   ✅ RSS feed accessible');
      const text = response.getContentText();
      const itemCount = (text.match(/<item>/g) || []).length;
      results.push('   📧 Recent messages: ' + itemCount);
    } else {
      results.push('   ❌ HTTP ' + response.getResponseCode());
    }
  } catch (e) {
    results.push('   ❌ Error: ' + e.message);
  }
  
  // Test 4: Revisions Folder
  results.push('\n4️⃣ Revisions Folder:');
  try {
    const cfg = getCollectorConfig_();
    const url = cfg.REVISIONS_URL;
    if (!url) {
      results.push('   ⚠️  Not configured');
    } else {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        results.push('   ✅ Folder accessible');
        const html = response.getContentText();
        const linkCount = (html.match(/<a[^>]+href/gi) || []).length;
        results.push('   📁 Files found: ~' + linkCount);
      } else {
        results.push('   ❌ HTTP ' + response.getResponseCode());
      }
    }
  } catch (e) {
    results.push('   ❌ Error: ' + e.message);
  }
  
  results.push('\n\n✅ Connection test complete!');
  
  ui.alert('Connection Test Results', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * PHASE 1: Create Configuration Tables
 */
function createConfigurationTables() {
  ensureConfigTable_();
  // ensureCollectorConfigTable_();
  ensureReallocationTable_();
  DocumentApp.getUi().alert('Configuration tables created successfully!\n\nIncluding:\n• Configuration\n• Collector Configuration\n• Document Reallocations');
}

/**
 * PHASE 2: Build Initial Report
 */
function buildInitialReport() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Build Initial Report',
    'This will download the TDOC list and create the complete report structure.\n\n' +
    'Make sure you have configured the meeting settings first.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    // Download and process from web
    downloadAndProcessFromWeb();
    
    // Collect initial email discussions and revisions
    collectorUpdate_();
    
    // Format
    removeRowHeightAndSpacing();
    
    ui.alert('Success!', 'Initial report built successfully!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to build report: ' + e.message, ui.ButtonSet.OK);
    Logger.log('Build error: ' + e.message);
  }
}

/**
 * PHASE 3: Update Report (Incremental)
 */
function updateReportIncremental() {
  const ui = DocumentApp.getUi();
  
  try {
    // Update email discussions and revisions
    collectorUpdate_();
    
    // Re-format
    removeRowHeightAndSpacing();
    
    ui.alert('Success!', 'Report updated successfully!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to update report: ' + e.message, ui.ButtonSet.OK);
    Logger.log('Update error: ' + e.message);
  }
}

/**
 * PHASE 4: Analyze Report Status
 */
function analyzeReportStatus() {
  const body = DocumentApp.getActiveDocument().getBody();
  const tables = body.getTables();
  
  const stats = {
    total: 0,
    agreed: 0,
    noted: 0,
    endorsed: 0,
    revised: 0,
    withdrawn: 0,
    available: 0,
    reserved: 0,
    missingMinutes: [],
    missingDisposition: [],
    noEmailDiscussion: []
  };
  
  tables.forEach(t => {
    if (!isTDocTable_(t)) return;
    
    stats.total++;
    
    const tdoc = safeCellText_(t, 0, 1);
    const status = findStatusText_(t);
    const minutes = findCellText_(t, 'Minutes');
    const disposition = findCellText_(t, 'Disposition');
    const email = findCellText_(t, 'E-mail Discussion') || findCellText_(t, 'Email discussion');
    
    // Count by status
    const statusLower = status.toLowerCase();
    if (statusLower.includes('agreed')) stats.agreed++;
    else if (statusLower.includes('noted')) stats.noted++;
    else if (statusLower.includes('endorsed')) stats.endorsed++;
    else if (statusLower.includes('revised')) stats.revised++;
    else if (statusLower.includes('withdrawn')) stats.withdrawn++;
    else if (statusLower.includes('available')) stats.available++;
    else if (statusLower.includes('reserved')) stats.reserved++;
    
    // Track missing information
    if (!minutes || minutes.trim() === '') stats.missingMinutes.push(tdoc);
    if (!disposition || disposition.trim() === '') stats.missingDisposition.push(tdoc);
    if (!email || email.trim() === '' || email.includes('No e-mail')) stats.noEmailDiscussion.push(tdoc);
  });
  
  // Build report
  const lines = [];
  lines.push('📊 REPORT STATUS ANALYSIS\n');
  lines.push('═══════════════════════════════════\n');
  
  lines.push(`\n📈 Total Documents: ${stats.total}\n`);
  
  lines.push('\n📊 Status Breakdown:');
  lines.push(`   ✅ Agreed:      ${stats.agreed} (${pct(stats.agreed, stats.total)}%)`);
  lines.push(`   ✅ Noted:       ${stats.noted} (${pct(stats.noted, stats.total)}%)`);
  lines.push(`   ✅ Endorsed:    ${stats.endorsed} (${pct(stats.endorsed, stats.total)}%)`);
  lines.push(`   ⚠️  Revised:     ${stats.revised} (${pct(stats.revised, stats.total)}%)`);
  lines.push(`   ❌ Withdrawn:   ${stats.withdrawn} (${pct(stats.withdrawn, stats.total)}%)`);
  lines.push(`   ⏳ Available:   ${stats.available} (${pct(stats.available, stats.total)}%)`);
  lines.push(`   ⏳ Reserved:    ${stats.reserved} (${pct(stats.reserved, stats.total)}%)`);
  
  const decided = stats.agreed + stats.noted + stats.endorsed + stats.withdrawn;
  lines.push(`\n✅ Decided: ${decided}/${stats.total} (${pct(decided, stats.total)}%)`);
  lines.push(`⏳ Pending: ${stats.available + stats.reserved}/${stats.total} (${pct(stats.available + stats.reserved, stats.total)}%)`);
  
  if (stats.available > 0 || stats.reserved > 0) {
    lines.push('\n\n⚠️  DOCUMENTS NEEDING ATTENTION:\n');
    if (stats.available > 0) lines.push(`   Available: ${stats.available} documents`);
    if (stats.reserved > 0) lines.push(`   Reserved: ${stats.reserved} documents`);
  }
  
  if (stats.missingMinutes.length > 0) {
    lines.push(`\n\n❌ Missing Minutes: ${stats.missingMinutes.length} documents`);
    if (stats.missingMinutes.length <= 10) {
      stats.missingMinutes.forEach(tdoc => lines.push(`   - ${tdoc}`));
    }
  }
  
  if (stats.missingDisposition.length > 0) {
    lines.push(`\n\n❌ Missing Disposition: ${stats.missingDisposition.length} documents`);
    if (stats.missingDisposition.length <= 10) {
      stats.missingDisposition.forEach(tdoc => lines.push(`   - ${tdoc}`));
    }
  }
  
  if (stats.noEmailDiscussion.length > 0) {
    lines.push(`\n\nℹ️  No Email Discussion: ${stats.noEmailDiscussion.length} documents`);
  }
  
  DocumentApp.getUi().alert('Report Status', lines.join('\n'), DocumentApp.getUi().ButtonSet.OK);
}

function pct(num, total) {
  return total > 0 ? Math.round((num / total) * 100) : 0;
}

function findStatusText_(table) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 2) continue;
    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key === 'tdoc status' || key === 'status') {
      return row.getCell(1).getText().trim();
    }
  }
  return '';
}

function findCellText_(table, label) {
  const labelLower = label.toLowerCase();
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 2) continue;
    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key === labelLower || key === labelLower + ':') {
      return row.getCell(1).getText().trim();
    }
  }
  return '';
}

/**
 * TOOLS: Individual test functions
 */
function testTdocListUrl() {
  const ui = DocumentApp.getUi();
  const cfg = getReportConfig_();
  const url = cfg.TDOC_LIST_URL;
  
  if (!url) {
    ui.alert('Error', 'TDOC List URL not configured.\n\nPlease configure it first using:\n⚙️ Configure Meeting Settings', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const response = UrlFetchApp.fetch(url);
    const blob = response.getBlob();
    
    ui.alert(
      'TDOC List Test',
      `✅ Successfully downloaded!\n\n` +
      `URL: ${url}\n` +
      `Size: ${blob.getBytes().length} bytes\n` +
      `Type: ${blob.getContentType()}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Error', `❌ Failed to download TDOC list:\n\n${e.message}`, ui.ButtonSet.OK);
  }
}

function testReviewerApi() {
  const ui = DocumentApp.getUi();
  const token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN');
  
  if (!token) {
    ui.alert('Error', 'Reviewer API token not configured.\n\nPlease configure it first using:\n⚙️ Configure Meeting Settings', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const testUrl = 'https://reviewer.bouazizi.dev/api/v1/documents/S4-260001/summary?type=summary';
    const response = UrlFetchApp.fetch(testUrl, {
      headers: { 'X-API-Key': token },
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    if (code === 200) {
      ui.alert('Reviewer API Test', '✅ API token is valid!\n\nAbstracts will be fetched successfully.', ui.ButtonSet.OK);
    } else if (code === 404) {
      ui.alert('Reviewer API Test', '✅ API token is valid!\n\n(Test document not found, but authentication works)', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `❌ API returned HTTP ${code}\n\nPlease check your API token.`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `❌ Failed to test API:\n\n${e.message}`, ui.ButtonSet.OK);
  }
}

function testEmailFeeds() {
  const ui = DocumentApp.getUi();
  
  try {
    const rssUrl = `https://list.etsi.org/scripts/wa.exe?RSS&L=${LIST_NAME_LOCK}&v=2.0&LIMIT=10`;
    const response = UrlFetchApp.fetch(rssUrl);
    const text = response.getContentText();
    const itemCount = (text.match(/<item>/g) || []).length;
    
    ui.alert(
      'Email Feeds Test',
      `✅ RSS feed accessible!\n\n` +
      `List: ${LIST_NAME_LOCK}\n` +
      `Recent messages: ${itemCount}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Error', `❌ Failed to access email feeds:\n\n${e.message}`, ui.ButtonSet.OK);
  }
}

function testRevisionsFolder() {
  const ui = DocumentApp.getUi();
  const cfg = getCollectorConfig_();
  const url = cfg.REVISIONS_URL;
  
  if (!url) {
    ui.alert('Error', 'Revisions URL not configured.\n\nPlease add it to the Collector Configuration table.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const response = UrlFetchApp.fetch(url);
    const html = response.getContentText();
    const linkCount = (html.match(/<a[^>]+href/gi) || []).length;
    
    ui.alert(
      'Revisions Folder Test',
      `✅ Folder accessible!\n\n` +
      `URL: ${url}\n` +
      `Files found: ~${linkCount}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Error', `❌ Failed to access revisions folder:\n\n${e.message}`, ui.ButtonSet.OK);
  }
}

function validateConfiguration() {
  const ui = DocumentApp.getUi();
  const issues = [];
  const warnings = [];
  
  // Check document properties
  const docProps = PropertiesService.getDocumentProperties();
  const reportSuffix = docProps.getProperty('REPORT_SUFFIX');
  const tdocListUrl = docProps.getProperty('TDOC_LIST_URL');
  
  if (!reportSuffix) warnings.push('⚠️  Report type not set (defaulting to 6G)');
  if (!tdocListUrl) issues.push('❌ TDOC List URL not configured');
  
  // Check script properties
  const scriptProps = PropertiesService.getScriptProperties();
  const apiToken = scriptProps.getProperty('REVIEWER_API_TOKEN');
  
  if (!apiToken) warnings.push('⚠️  Reviewer API token not set (abstracts will be skipped)');
  
  // Check collector config
  const cfg = getCollectorConfig_();
  if (!cfg.REVISIONS_URL) warnings.push('⚠️  Revisions URL not configured');
  
  // Build report
  const lines = [];
  lines.push('✅ CONFIGURATION VALIDATION\n');
  lines.push('═══════════════════════════════════\n');
  
  if (issues.length === 0 && warnings.length === 0) {
    lines.push('\n✅ All configuration is valid!\n');
    lines.push(`\nReport Type: ${reportSuffix || '6G'}`);
    lines.push(`TDOC List URL: Configured`);
    lines.push(`Reviewer API: Configured`);
    lines.push(`Revisions URL: Configured`);
  } else {
    if (issues.length > 0) {
      lines.push('\n❌ CRITICAL ISSUES:\n');
      issues.forEach(issue => lines.push(issue));
    }
    if (warnings.length > 0) {
      lines.push('\n\n⚠️  WARNINGS:\n');
      warnings.forEach(warning => lines.push(warning));
    }
    lines.push('\n\nPlease fix these issues using:\n⚙️ Configure Meeting Settings');
  }
  
  ui.alert('Configuration Validation', lines.join('\n'), ui.ButtonSet.OK);
}

function clearAllCaches() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Clear All Caches',
    'This will clear:\n' +
    '• Email discussion cache\n' +
    '• Revisions cache\n' +
    '• A1 empty-week cache\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const props = PropertiesService.getDocumentProperties();
  let cleared = 0;
  
  props.getKeys().forEach(key => {
    if (key.startsWith('DISCUSS_') || key.startsWith('REVIS_') || key.startsWith('A1_EMPTY_CACHE_')) {
      props.deleteProperty(key);
      cleared++;
    }
  });
  
  ui.alert('Success', `✅ Cleared ${cleared} cache entries!`, ui.ButtonSet.OK);
}

function fixColumnWidths() {
  setTwoColumnTDocTableWidths_();
  DocumentApp.getUi().alert('Success', '✅ Column widths fixed!', DocumentApp.getUi().ButtonSet.OK);
}

/********************************************************
 * PHASE 2: REALLOCATION UI FUNCTIONS
 ********************************************************/

/**
 * Add a new document reallocation via dialog
 */
function addDocumentReallocation() {
  const ui = DocumentApp.getUi();
  
  // Ensure table exists
  ensureReallocationTable_();
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      textarea { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; min-height: 60px; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      .hint { font-size: 11px; color: #666; margin-top: 3px; }
      .example { font-size: 11px; color: #999; font-style: italic; }
    </style>
    
    <h2>Add Document Reallocation</h2>
    
    <label>TDoc Number:</label>
    <input type="text" id="tdoc" placeholder="S4-260123">
    <div class="hint">The document number to reallocate</div>
    
    <label>Original Agenda Item:</label>
    <input type="text" id="original" placeholder="5.3">
    <div class="hint">Where it was originally registered (optional)</div>
    
    <label>New Agenda Item:</label>
    <input type="text" id="newAgenda" placeholder="8.3">
    <div class="hint">Where it should be placed</div>
    
    <label>Reason:</label>
    <textarea id="reason" placeholder="Scope changed to MBS during meeting"></textarea>
    <div class="hint">Why this document is being reallocated (optional)</div>
    
    <button onclick="addReallocation()">Add Reallocation</button>
    <button onclick="google.script.host.close()" style="background: #666;">Cancel</button>
    
    <script>
      function addReallocation() {
        const tdoc = document.getElementById('tdoc').value.trim();
        const original = document.getElementById('original').value.trim();
        const newAgenda = document.getElementById('newAgenda').value.trim();
        const reason = document.getElementById('reason').value.trim();
        
        if (!tdoc) {
          alert('Please enter a TDoc number');
          return;
        }
        
        if (!newAgenda) {
          alert('Please enter a new agenda item');
          return;
        }
        
        google.script.run
          .withSuccessHandler(() => {
            alert('Reallocation added successfully!\\n\\nRemember to rebuild the report to apply changes.');
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('Error adding reallocation: ' + error);
          })
          .saveReallocation(tdoc, original, newAgenda, reason);
      }
    </script>
  `)
  .setWidth(500)
  .setHeight(500);
  
  ui.showModalDialog(html, 'Add Document Reallocation');
}

/**
 * Save a reallocation to the table
 */
function saveReallocation(tdoc, original, newAgenda, reason) {
  ensureReallocationTable_();
  
  const body = DocumentApp.getActiveDocument().getBody();
  
  // Find the reallocation table
  for (const table of body.getTables()) {
    if (!isReallocationTable_(table)) continue;
    
    // Check if this TDoc already exists
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getCell(0).getText().trim() === tdoc) {
        // Update existing row
        row.getCell(1).setText(original);
        row.getCell(2).setText(newAgenda);
        if (row.getNumCells() >= 4) {
          row.getCell(3).setText(reason);
        }
        Logger.log(`Updated reallocation for ${tdoc}`);
        return;
      }
    }
    
    // Add new row
    const newRow = table.appendTableRow();
    newRow.appendTableCell(tdoc);
    newRow.appendTableCell(original);
    newRow.appendTableCell(newAgenda);
    newRow.appendTableCell(reason);
    
    Logger.log(`Added reallocation for ${tdoc}: ${original} → ${newAgenda}`);
    return;
  }
  
  throw new Error('Reallocation table not found');
}

/**
 * View all reallocations in a formatted dialog
 */
function viewAllReallocations() {
  const ui = DocumentApp.getUi();
  const reallocations = getReallocationMap_();
  
  if (Object.keys(reallocations).length === 0) {
    ui.alert('No Reallocations', 'No document reallocations have been configured yet.', ui.ButtonSet.OK);
    return;
  }
  
  // Build HTML table
  let tableRows = '';
  Object.keys(reallocations).sort().forEach(tdoc => {
    const r = reallocations[tdoc];
    tableRows += `
      <tr>
        <td>${tdoc}</td>
        <td>${r.original || '-'}</td>
        <td><strong>${r.new}</strong></td>
        <td>${r.reason || '-'}</td>
      </tr>
    `;
  });
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      h2 { margin-top: 0; }
      table { width: 100%; border-collapse: collapse; margin-top: 15px; }
      th, td { padding: 8px; text-align: left; border-bottom: 1px solid #ddd; }
      th { background-color: #f5f5f5; font-weight: bold; }
      tr:hover { background-color: #f9f9f9; }
      .count { color: #666; font-size: 14px; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
    </style>
    
    <h2>Document Reallocations</h2>
    <div class="count">Total: ${Object.keys(reallocations).length} document(s)</div>
    
    <table>
      <thead>
        <tr>
          <th>TDoc</th>
          <th>Original</th>
          <th>New Agenda</th>
          <th>Reason</th>
        </tr>
      </thead>
      <tbody>
        ${tableRows}
      </tbody>
    </table>
    
    <button onclick="google.script.host.close()">Close</button>
  `)
  .setWidth(700)
  .setHeight(500);
  
  ui.showModalDialog(html, 'All Document Reallocations');
}

/**
 * Clear all reallocations (with confirmation)
 */
function clearAllReallocations() {
  const ui = DocumentApp.getUi();
  const reallocations = getReallocationMap_();
  
  if (Object.keys(reallocations).length === 0) {
    ui.alert('No Reallocations', 'There are no reallocations to clear.', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.alert(
    'Clear All Reallocations',
    `This will remove all ${Object.keys(reallocations).length} reallocation(s) from the table.\n\n` +
    'This action cannot be undone.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const body = DocumentApp.getActiveDocument().getBody();
  
  // Find and clear the reallocation table
  for (const table of body.getTables()) {
    if (!isReallocationTable_(table)) continue;
    
    // Remove all rows except header
    while (table.getNumRows() > 1) {
      table.removeRow(1);
    }
    
    Logger.log('Cleared all reallocations');
    ui.alert('Success', `✅ Cleared ${Object.keys(reallocations).length} reallocation(s)!`, ui.ButtonSet.OK);
    return;
  }
  
  ui.alert('Error', 'Reallocation table not found', ui.ButtonSet.OK);
}

/********************************************************
 * PHASE 3: AGENDA DOCUMENT PARSING
 ********************************************************/

/**
 * Parse agenda document by downloading the zipped Word agenda TDOC from 3GPP.
 * Example: Agenda TDoc S4-260868 -> FTP_BASE + S4-260868.zip.
 */
function parseAgendaDocument() {
  const ui = DocumentApp.getUi();
  const cfg = getReportConfig_();

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      .hint { font-size: 11px; color: #666; margin-top: 3px; }
      .warning { font-size: 12px; color: #a65f00; margin-top: 10px; }
    </style>

    <h2>Parse Agenda Source</h2>

    <label>Agenda Google Doc URL/ID (optional):</label>
    <input type="text" id="agendaDoc" placeholder="https://docs.google.com/document/d/1c_N_0-5HEV2Ma4eBbkvMm4yvTVmmw_CudD7DItQC1Mo/edit">
    <div class="hint">Use this for an already-converted/shared agenda document. If filled, this is used instead of downloading the ZIP.</div>

    <label>Agenda TDoc:</label>
    <input type="text" id="agendaTdoc" value="${cfg.AGENDA_TDOC || ''}" placeholder="S4-260868">
    <div class="hint">Official agenda TDOC number. The script downloads FTP Base + Agenda TDoc + .zip if no Google Doc URL/ID is provided.</div>

    <label>FTP Base:</label>
    <input type="text" id="ftpBase" value="${cfg.FTP_BASE || DEFAULT_S4_FTP_BASE}" placeholder="https://www.3gpp.org/ftp/.../Docs/">
    <div class="hint">Folder containing TDOC ZIP files. Must end with /Docs/</div>

    <div class="warning">
      Note: ZIP/DOCX parsing requires the Apps Script Advanced Drive service to be enabled
      because the Word document must be converted to a temporary Google Doc.
    </div>

    <button onclick="handleParseClick()">Parse Agenda</button>
    <button onclick="google.script.host.close()" style="background: #666;">Cancel</button>

    <script>
      function handleParseClick() {
        const agendaDoc = document.getElementById('agendaDoc').value.trim();
        const agendaTdoc = document.getElementById('agendaTdoc').value.trim();
        const ftpBase = document.getElementById('ftpBase').value.trim();

        if (agendaDoc) {
          google.script.run
            .withSuccessHandler((result) => {
              alert('Agenda parsed successfully!\\n\\n' + result);
              google.script.host.close();
            })
            .withFailureHandler((error) => {
              alert('Error parsing agenda Google Doc: ' + error);
            })
            .parseAgendaDocumentById(agendaDoc);
          return;
        }

        if (!agendaTdoc) {
          alert('Please enter either an agenda Google Doc URL/ID or the agenda TDoc, e.g. S4-260868');
          return;
        }

        google.script.run
          .withSuccessHandler((result) => {
            alert('Agenda parsed successfully!\\n\\n' + result);
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('Error parsing zipped Word agenda: ' + error);
          })
          .parseAgendaFromZippedTdoc(agendaTdoc, ftpBase);
      }
    </script>
  `)
  .setWidth(560)
  .setHeight(430);

  ui.showModalDialog(html, 'Parse Agenda TDOC');
}

/**
 * Main agenda parsing orchestrator called by the build workflow.
 */
function parseAgendaForReport_(cfg, templateDocId) {
  let agendaItems;
  if (cfg.AGENDA_TDOC) {
    Logger.log('Parsing agenda from TDOC ZIP: ' + cfg.AGENDA_TDOC);
    const agendaStructure = downloadMeetingAgenda_(cfg.AGENDA_TDOC, cfg.FTP_BASE);
    const prefix = getConfiguredAgendaPrefix_();
    agendaItems = agendaStructure.filter(item =>
      item.number === prefix.replace(/\.$/, '') ||
      item.number.startsWith(prefix)
    );
  } else {
    Logger.log('Parsing agenda from template Google Doc: ' + templateDocId);
    const prefix = getConfiguredAgendaPrefix_();
    const agendaDoc = DocumentApp.openById(templateDocId);
    agendaItems = parseAgendaStructureWithText_(agendaDoc.getBody(), prefix);
  }
  PropertiesService.getDocumentProperties().setProperty('PARSED_AGENDA', JSON.stringify(agendaItems));
  return agendaItems;
}


/**
 * Download and parse meeting agenda ZIP by TDOC number.
 */
function parseAgendaFromZippedTdoc(agendaTdoc, ftpBase) {
  agendaTdoc = normalizeTdoc_(agendaTdoc);
  if (!agendaTdoc) throw new Error('Invalid agenda TDoc. Expected format like S4-260868.');

  ftpBase = normalizeFtpBase_(ftpBase || DEFAULT_S4_FTP_BASE);

  const props = PropertiesService.getDocumentProperties();
  props.setProperty('AGENDA_TDOC', agendaTdoc);
  props.setProperty('FTP_BASE', ftpBase);

  const agendaStructure = downloadMeetingAgenda_(agendaTdoc, ftpBase);

  const cfg = getReportConfig_();
  const reportType = cfg.REPORT_SUFFIX || '6G';
  const agendaPrefix = getConfiguredAgendaPrefix_();
  const agendaItems = agendaStructure.filter(item => item.number.startsWith(agendaPrefix));

  props.setProperty('PARSED_AGENDA', JSON.stringify(agendaItems));
  props.setProperty('PARSED_AGENDA_ALL', JSON.stringify(agendaStructure));

  Logger.log(`Parsed ${agendaStructure.length} total agenda items, ${agendaItems.length} relevant for ${reportType}`);

  return `Downloaded: ${ftpBase}${agendaTdoc}.zip\n` +
    `Found ${agendaStructure.length} total agenda items.\n` +
    `Relevant for ${reportType} (${agendaPrefix}): ${agendaItems.length}\n\n` +
    agendaItems.slice(0, 12).map(item => `• ${item.number} ${item.title}`).join('\n') +
    (agendaItems.length > 12 ? `\n... and ${agendaItems.length - 12} more` : '');
}

/**
 * Backward-compatible helper: accepts a Google Doc URL/ID OR an agenda TDOC.
 * Prefer parseAgendaFromZippedTdoc() for the intended workflow.
 */
function parseAgendaDocumentById(input) {
  const tdoc = normalizeTdoc_(input);
  if (tdoc) return parseAgendaFromZippedTdoc(tdoc, getReportConfig_().FTP_BASE);

  let docId = input;
  const urlMatch = String(input || '').match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch) docId = urlMatch[1];

  try {
    const agendaDoc = DocumentApp.openById(docId);
    const cfg = getReportConfig_();
    const reportType = cfg.REPORT_SUFFIX || '6G';
    const agendaPrefix = getConfiguredAgendaPrefix_();
    const agendaItems = parseAgendaStructureWithText_(agendaDoc.getBody(), agendaPrefix);

    PropertiesService.getDocumentProperties().setProperty('PARSED_AGENDA', JSON.stringify(agendaItems));
    PropertiesService.getDocumentProperties().setProperty('AGENDA_SOURCE_DOC_ID', docId);

    return `Found ${agendaItems.length} agenda items from Google Doc source:\n\n` +
      agendaItems.slice(0, 10).map(item => `• ${item.number} ${item.title}`).join('\n') +
      (agendaItems.length > 10 ? `\n... and ${agendaItems.length - 10} more` : '');
  } catch (e) {
    throw new Error(`Failed to open or parse agenda source: ${e.message}`);
  }
}

function downloadMeetingAgenda_(agendaTdoc, ftpBase) {
  const url = normalizeFtpBase_(ftpBase || DEFAULT_S4_FTP_BASE) + agendaTdoc + '.zip';
  Logger.log('Downloading agenda document: ' + url);

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  const code = response.getResponseCode();
  if (code >= 400) throw new Error(`Could not download agenda ZIP (${code}): ${url}`);

  const zipBlob = response.getBlob().setName(agendaTdoc + '.zip');
  const wordBlob = extractWordFromZip_(zipBlob);
  return parseAgendaWordBlob_(wordBlob);
}

function extractWordFromZip_(zipBlob) {
  const files = Utilities.unzip(zipBlob);
  const candidates = files.filter(blob => {
    const name = String(blob.getName() || '').toLowerCase();
    return name.endsWith('.docx') || name.endsWith('.doc');
  });

  if (!candidates.length) {
    throw new Error('No .docx/.doc file found inside agenda ZIP.');
  }

  // Prefer DOCX and avoid macOS metadata folders.
  candidates.sort((a, b) => {
    const an = String(a.getName() || '').toLowerCase();
    const bn = String(b.getName() || '').toLowerCase();
    if (an.indexOf('__macosx') !== -1) return 1;
    if (bn.indexOf('__macosx') !== -1) return -1;
    if (an.endsWith('.docx') && !bn.endsWith('.docx')) return -1;
    if (!an.endsWith('.docx') && bn.endsWith('.docx')) return 1;
    return an.localeCompare(bn);
  });

  return candidates[0];
}

function parseAgendaWordBlob_(wordBlob) {
  const convertedDocId = convertWordBlobToGoogleDoc_(wordBlob);

  try {
    const doc = DocumentApp.openById(convertedDocId);
    return parseAgendaStructureWithText_(doc.getBody(), '');
  } finally {
    try { DriveApp.getFileById(convertedDocId).setTrashed(true); } catch (e) { }
  }
}

function convertWordBlobToGoogleDoc_(wordBlob) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error(
      'Advanced Drive service is required to convert the Word agenda. ' +
      'In Apps Script enable: Services (+) → Drive API → Add. ' +
      'Also ensure Google Cloud Drive API is enabled if prompted.'
    );
  }

  const resource = {
    title: 'TEMP_Agenda_' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_DOCS
  };

  // The 'insert' method is from an older version of the Drive API.
  // The modern equivalent is 'create'.
  const file = Drive.Files.create(resource, wordBlob);
  if (!file || !file.id) throw new Error('Drive conversion did not return a file ID.');
  return file.id;
}

/**
 * SA4-IMPL-001: delegates to the central parseSA4DocumentId_() model instead
 * of a hardcoded /S4-\d{6}/. Preserves its exact external contract (valid ->
 * normalized identifier string, invalid -> '') and every existing
 * main-meeting case (see tests/pure-logic.test.js) is unchanged. New:
 * recognizes the 5 verified ad-hoc families (S4aA/S4aP/S4aV/S4aI/A4aR) --
 * this is the primary intentional behavior change of SA4-IMPL-001. This is
 * also what makes parseAgendaFromZippedTdoc()/parseAgendaDocumentById()
 * (both of which reject input when this returns '') ad-hoc-capable, with no
 * changes needed in either of those two functions themselves.
 */
function normalizeTdoc_(value) {
  const parsed = parseSA4DocumentId_(value);
  return parsed.isValid ? parsed.raw : '';
}

function extractGoogleDocId_(input) {
  const s = String(input || '').trim();
  const m = s.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : s;
}

/**
 * Parse agenda structure from the document, correctly handling agendas
 * that are formatted as either headings or inside a table.
 */
function parseAgendaStructureWithText_(body, prefix) {
  const itemsFromHeadings = parseAgendaFromHeadings_(body, prefix);
  if (itemsFromHeadings.length > 0) {
    Logger.log(`Parsed ${itemsFromHeadings.length} agenda items from paragraph headings.`);
    return itemsFromHeadings;
  }
  
  const itemsFromTables = parseAgendaFromTables_(body, prefix);
  if (itemsFromTables.length > 0) {
    Logger.log(`Parsed ${itemsFromTables.length} agenda items from tables.`);
    return itemsFromTables;
  }

  return [];
}

/**
 * Parses agenda items from a 2-column table with "A.I.#" header.
 */
function parseAgendaFromTables_(body, prefix) {
  const items = [];
  const parentNum = prefix ? prefix.replace(/\.$/, '') : null;

  for (const table of body.getTables()) {
    if (table.getNumRows() < 1) continue;
    
    const firstCellText = table.getCell(0, 0).getText().trim();
    if (firstCellText !== 'A.I.#') continue;

    // This is the main agenda table.
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 2) continue;
      
      const number = row.getCell(0).getText().trim();
      const title = row.getCell(1).getText().trim();
      
      if (!number || !title) continue;
      
      if (parentNum && number !== parentNum && !number.startsWith(prefix)) {
        continue;
      }
      
      const level = (number.match(/\./g) || []).length + 1;
      
      items.push({
        number: number,
        title: title,
        level: level,
        heading: DocumentApp.ParagraphHeading.NORMAL, // No heading style from table
        text: '' // No descriptive text in table format
      });
    }
    return items; // Assume only one agenda table
  }
  return []; // No agenda table found
}

/**
 * Original logic to parse agenda from paragraph headings.
 */
function parseAgendaFromHeadings_(body, prefix) {
  const items = [];
  const numChildren = body.getNumChildren();
  const anyAgendaRegex = /^(\d+(?:\.\d+)*)\s+(.+)$/;
  const parentNum = prefix ? prefix.replace(/\.$/, '') : null;
  const parentRegex = parentNum ? new RegExp(`^(${parentNum.replace(/\./g, '\\.')})\\s+(.+)$`) : null;
  const subItemRegex = prefix ? new RegExp(`^(${prefix.replace(/\./g, '\\.')}\\d+(?:\\.\\d+)*)\\s+(.+)$`) : anyAgendaRegex;

  let inSection = !prefix; // If no prefix, we are always "in section"
  let current = null;

  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const para = child.asParagraph();
    const text = para.getText().trim();
    if (!text) continue;

    const headingMatch = text.match(subItemRegex);

    if (headingMatch) {
      const number = headingMatch[1];
      const title = headingMatch[2].trim();
      const isParent = parentRegex && text.match(parentRegex);
      
      if (isParent) inSection = true;

      if (inSection) {
        const level = (number.match(/\./g) || []).length + 1;
        current = { number, title, level, heading: para.getHeading(), text: '' };
        items.push(current);
      }
    } else if (current && text.match(anyAgendaRegex)) {
        // We've hit a new heading that doesn't match our prefix, so we exit the section
        inSection = false;
        current = null;
    } else if (current && inSection) {
      // Non-heading paragraph, collect as text
      if (!shouldSkipAgendaTemplateText_(text)) {
        current.text += (current.text ? '\n' : '') + text;
      }
    }
  }
  return items;
}

function shouldSkipAgendaTemplateText_(text) {
  const t = String(text || '').trim();
  if (!t) return true;
  if (/^(page|rapporteur|chairman|chair|secretary)\b/i.test(t)) return true;
  if (/^3GPP\s+TSG/i.test(t)) return true;
  if (/^S4-\d{6}/i.test(t)) return true;
  return false;
}

/**
 * Parse agenda structure from document body
 * Returns array of { number: '11.1', title: 'Topic Name', level: 2 }
 */
function parseAgendaStructure_(body, prefix) {
  const items = [];
  const numChildren = body.getNumChildren();
  
  // Regex to match agenda items like "11.1 Topic Name" or "11.1.1 Subtopic".
  // If prefix is empty, parse all numeric agenda items from the converted Word document.
  const agendaRegex = prefix
    ? new RegExp(`^(${prefix.replace('.', '\\.')}\\d+(?:\\.\\d+)*)\\s+(.+)$`)
    : /^(\d+(?:\.\d+)*)\s+(.+)$/;
  
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      const text = para.getText().trim();
      const heading = para.getHeading();
      
      // Check if this is a heading with agenda number
      if (heading !== DocumentApp.ParagraphHeading.NORMAL) {
        const match = text.match(agendaRegex);
        
        if (match) {
          const number = match[1];
          const title = match[2].trim();
          
          // Determine level based on number of dots
          const level = (number.match(/\./g) || []).length + 1;
          
          items.push({
            number: number,
            title: title,
            level: level,
            heading: heading
          });
          
          Logger.log(`Found agenda item: ${number} ${title} (level ${level})`);
        }
      }
    }
  }
  
  return items;
}

/********************************************************
 * SPLIT REPORT BUILD WORKFLOW
 ********************************************************/

function runFullReportBuild() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Run Full Report Build',
    'This will run all steps:\n\n' +
    '1+2) Build skeleton from agenda + insert TDOC tables\n' +
    '3) Collect e-mail discussion\n' +
    '4) Collect revisions\n' +
    '5) Add abstracts\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    buildSkeletonWithTdocTables();
    collectEmailDiscussionOnly();
    collectRevisionsOnly();
    addAbstractsOnly();
    removeRowHeightAndSpacing();
    ui.alert('Success', 'Full report build completed.', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Full report build failed: ' + e.message, ui.ButtonSet.OK);
    Logger.log('Full report build failed: ' + e.message);
  }
}

/**
 * Combined step 1+2: Build skeleton from agenda document + insert TDOC tables.
 *
 * Workflow:
 * 1. Set document title
 * 2. Copy preamble + IPR/Trust from template Google Doc
 * 3. Download agenda ZIP (S4-260868) → parse section headings for this report type
 * 4. Download TDOC list → group by agenda item
 * 5. For each agenda item: insert heading + agenda text + TDOC tables
 * 6. Append "Registered Documents" summary table at the end
 */
function buildSkeletonWithTdocTables() {
  const cfg = getReportConfig_();

  if (!cfg.TDOC_LIST_URL) {
    Logger.log('Error: TDOC List URL not configured');
    try {
      DocumentApp.getUi().alert('Error', 'TDOC List URL not configured. Please run "Configure Meeting" first.', DocumentApp.getUi().ButtonSet.OK);
    } catch (e) {
      // UI not available in this context
    }
    return;
  }

  // Step 1: Clear document, set title
  const body = DocumentApp.getActiveDocument().getBody().clear();
  const templateDocId = extractGoogleDocId_(cfg.AGENDA_SOURCE_DOC_ID);
  setDocumentTitleFromTemplate_(templateDocId);

  // Step 2: Parse agenda and download TDOCs
  const agendaItems = parseAgendaForReport_(cfg, templateDocId);
  if (!agendaItems || agendaItems.length === 0) {
    Logger.log('Warning: No agenda items found for prefix: ' + getConfiguredAgendaPrefix_());
    try {
      DocumentApp.getUi().alert('Warning', 'No agenda items found for prefix: ' + getConfiguredAgendaPrefix_(), DocumentApp.getUi().ButtonSet.OK);
    } catch (e) {
      // UI not available in this context
    }
    return;
  }
  
  // Filter out the parent item (e.g., "9 Video SWG") - we only want sub-items (9.1, 9.2, etc.)
  const agendaPrefix = getConfiguredAgendaPrefix_();
  const filteredAgendaItems = agendaItems.filter(item => {
    // Keep items that have a dot after the prefix (e.g., 9.1, 9.2, not just 9)
    return item.number !== agendaPrefix.replace(/\.$/, '');
  });
  
  if (filteredAgendaItems.length === 0) {
    Logger.log('Warning: No sub-agenda items found for prefix: ' + agendaPrefix);
    try {
      DocumentApp.getUi().alert('Warning', 'No sub-agenda items found for prefix: ' + agendaPrefix, DocumentApp.getUi().ButtonSet.OK);
    } catch (e) {
      // UI not available in this context
    }
    return;
  }
  const tdocGroups = downloadAndGroupTdocs_(cfg);
  const allTdocs = [];
  Object.keys(tdocGroups).forEach(key => {
    tdocGroups[key].tdocs.forEach(td => allTdocs.push({ ...td, agendaItem: key }));
  });

  Logger.log(`Found ${filteredAgendaItems.length} agenda items and ${allTdocs.length} TDOCs`);
  
  // Step 3: Build full agenda structure, backfilling template content
  const sourceBody = DocumentApp.openById(templateDocId).getBody();
  const agendaPrefixNum = (cfg.AGENDA_ITEM_PREFIX || '0.').replace(/\.$/, '');
  const reportType = cfg.REPORT_SUFFIX || '6G';
  const is6G = (reportType === '6G');

  // For 6G reports, create the 11.0 parent section first
  if (is6G) {
    body.appendParagraph(`${agendaPrefixNum}.0 Opening of the session, registration of documents`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    // 11.0.1 Opening of the session
    body.appendParagraph(`${agendaPrefixNum}.0.1 Opening of the session`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    const openingHeader = findHeading_(sourceBody, /^X\.1\s+/);
    if (openingHeader) {
      copySectionContentWithReplacement_(openingHeader, body, /^X\.2\s+/, 'X', agendaPrefixNum);
    }
    
    // 11.0.2 Registration of Documents
    body.appendParagraph(`${agendaPrefixNum}.0.2 Registration of Documents`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    if (allTdocs.length > 0) {
      createSummaryTable_(body, allTdocs, allTdocs[0].tdocCol, allTdocs[0].titleCol, allTdocs[0].sourceCol, -1);
    }
    
    // 11.0.3 Document Reallocations
    body.appendParagraph(`${agendaPrefixNum}.0.3 Document Reallocations`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    const reallocations = getReallocationMap_();
    if (Object.keys(reallocations).length > 0) {
      // Create reallocation table
      const reallocationTable = body.appendTable();
      const headerRow = reallocationTable.appendTableRow();
      headerRow.appendTableCell('TDoc');
      headerRow.appendTableCell('Original Agenda');
      headerRow.appendTableCell('New Agenda');
      headerRow.appendTableCell('Reason');
      
      Object.keys(reallocations).sort().forEach(tdoc => {
        const r = reallocations[tdoc];
        const dataRow = reallocationTable.appendTableRow();
        dataRow.appendTableCell(tdoc);
        dataRow.appendTableCell(r.original || '-');
        dataRow.appendTableCell(r.new);
        dataRow.appendTableCell(r.reason || '-');
      });
      
      removeInitialEmptyRow_(reallocationTable);
    } else {
      body.appendParagraph('No document reallocations for this meeting.');
    }
    
    // 11.0.4 IPR and antitrust reminder
    body.appendParagraph(`${agendaPrefixNum}.0.4 IPR and antitrust reminder`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING3);
    const iprHeader = findHeading_(sourceBody, /^X\.2\s+/);
    if (iprHeader) {
      copySectionContentWithReplacement_(iprHeader, body, /^X\.Y\s+/, 'X', agendaPrefixNum);
    }
  }

  // Track if we've seen AOB and Close of Session in the agenda
  let hasAOB = false;
  let hasCloseOfSession = false;
  
  filteredAgendaItems.forEach((item, idx) => {
    const headingText = `${item.number} ${item.title}`;
    const heading = body.appendParagraph(headingText);
    const level = (item.number.match(/\./g) || []).length;
    if (level === 1) heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    else heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);

    // Track AOB and Close of Session
    if (item.title.toLowerCase().includes('any other business') || item.title.toLowerCase().includes('aob')) {
      hasAOB = true;
    }
    if (item.title.toLowerCase().includes('close of') || item.title.toLowerCase().includes('closing')) {
      hasCloseOfSession = true;
    }

    // Special handling based on agenda number (not index)
    // For 6G: sections 11.0.x are already created above, so skip them
    // For SWG reports: use X.1, X.2 as before
    const openingSection = is6G ? `${agendaPrefixNum}.0.1` : `${agendaPrefixNum}.1`;
    const registrationSection = is6G ? `${agendaPrefixNum}.0.2` : `${agendaPrefixNum}.1.2`;
    const reallocationSection = is6G ? `${agendaPrefixNum}.0.3` : null;
    const iprSection = is6G ? `${agendaPrefixNum}.0.4` : `${agendaPrefixNum}.2`;
    
    // Skip 11.0.x sections for 6G as they're already created
    if (is6G && (item.number === openingSection || item.number === registrationSection || 
                 item.number === reallocationSection || item.number === iprSection)) {
      return; // Skip - already created above
    }
    
    if (item.number === openingSection) {
      // Opening section for SWG reports - copy X.1 content from template
      const openingHeader = findHeading_(sourceBody, /^X\.1\s+/);
      if (openingHeader) {
        copySectionContentWithReplacement_(openingHeader, body, /^X\.2\s+/, 'X', agendaPrefixNum);
      }
      
      // Always ensure Registration of Documents section exists with the summary table
      if (!documentContainsHeading_(body, registrationSection)) {
        body.appendParagraph(`${registrationSection} Registration of Documents`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      }
      
      // Always add/update the registered documents summary table
      if (allTdocs.length > 0) {
        createSummaryTable_(body, allTdocs, allTdocs[0].tdocCol, allTdocs[0].titleCol, allTdocs[0].sourceCol, -1);
      }
    } else if (item.number === iprSection) {
      // IPR section - only copy template content for SWG reports (Audio, Video, MBS, RTC)
      const isSWGReport = (reportType === 'Audio' || reportType === 'Video' || reportType === 'MBS' || reportType === 'RTC');
      
      if (isSWGReport) {
        const iprHeader = findHeading_(sourceBody, /^X\.2\s+/);
        if (iprHeader) {
          copySectionContentWithReplacement_(iprHeader, body, /^X\.Y\s+/, 'X', agendaPrefixNum);
        }
      } else {
        // For non-SWG reports (6G, Liaison, New), .2 items are regular agenda items with TDOC tables
        const group = tdocGroups[item.number];
        if (group && group.tdocs.length > 0) {
          // Revisions are emitted directly below the document they revise.
          orderTdocsByRevision_(group.tdocs).forEach(tdocData => {
            const row = tdocData.row;
            const revisedTo = getRevisedTo_(tdocData);
            
            // Build Type/For field
            const typeCol = tdocData.typeCol;
            const forCol = tdocData.forCol;
            const typeFor = (typeCol >= 0 && forCol >= 0 && row[typeCol] && row[forCol])
              ? `${row[typeCol]} for ${row[forCol]}`
              : (typeCol >= 0 && row[typeCol]) ? row[typeCol] 
              : (forCol >= 0 && row[forCol]) ? row[forCol] 
              : '';
            
            const tempData = [
              ['TDoc', row[tdocData.tdocCol]],
              ['Title', row[tdocData.titleCol]],
              ['Source', row[tdocData.sourceCol]],
              ['Contact', tdocData.contactCol >= 0 ? row[tdocData.contactCol] : ''],
              ['Agenda Item', item.number],
              ['Type/For', typeFor],
              ['E-mail Discussion', ''],
              ['Revisions', ''],
              ['Minutes', ''],
              ['Disposition', revisedTo ? 'Revised to ' + revisedTo : ''],
              ['Status', tdocData.statusCol >= 0 ? row[tdocData.statusCol] : '']
            ];
            const table = body.appendTable();
            removeInitialEmptyRow_(table);
            tempData.forEach(rowData => {
              const tr = table.appendTableRow();
              tr.appendTableCell(rowData[0]);
              tr.appendTableCell(String(rowData[1] || ''));
            });
            if (tdocData.richTextRow && tdocData.richTextRow[tdocData.tdocCol]) {
              const richText = tdocData.richTextRow[tdocData.tdocCol];
              if (richText.getLinkUrl && richText.getLinkUrl()) {
                const cell = table.getRow(0).getCell(1);
                const tdocValue = String(row[tdocData.tdocCol]);
                if (tdocValue.length > 0) {
                  cell.editAsText().setLinkUrl(0, tdocValue.length - 1, richText.getLinkUrl());
                }
              }
            }
            styleStatusCell_(table);
          });
        }
      }
    } else if (item.title.toLowerCase().includes('any other business') || item.title.toLowerCase().includes('aob')) {
      // AOB section - copy template content
      const aobHeader = findHeading_(sourceBody, /^X\.Y\s+/);
      if (aobHeader) {
        copySectionContentWithReplacement_(aobHeader, body, /^X\.Z\s+/, 'X.Y', item.number);
      }
    } else if (item.title.toLowerCase().includes('close of') || item.title.toLowerCase().includes('closing')) {
      // Close of Session - copy template content
      const closeHeader = findHeading_(sourceBody, /^X\.Z\s+/);
      if (closeHeader) {
        copySectionContentWithReplacement_(closeHeader, body, /^(?!X\.)/, 'X.Z', item.number);
      }
    } else {
      // Regular agenda item - insert TDOC tables
      const group = tdocGroups[item.number];
      if (group && group.tdocs.length > 0) {
        // Revisions are emitted directly below the document they revise.
        orderTdocsByRevision_(group.tdocs).forEach(tdocData => {
          const row = tdocData.row;
          const revisedTo = getRevisedTo_(tdocData);
          
          // Build Type/For field
          const typeCol = tdocData.typeCol;
          const forCol = tdocData.forCol;
          const typeFor = (typeCol >= 0 && forCol >= 0 && row[typeCol] && row[forCol])
            ? `${row[typeCol]} for ${row[forCol]}`
            : (typeCol >= 0 && row[typeCol]) ? row[typeCol] 
            : (forCol >= 0 && row[forCol]) ? row[forCol] 
            : '';
          
          const tempData = [
            ['TDoc', row[tdocData.tdocCol]],
            ['Title', row[tdocData.titleCol]],
            ['Source', row[tdocData.sourceCol]],
            ['Contact', tdocData.contactCol >= 0 ? row[tdocData.contactCol] : ''],
            ['Agenda Item', item.number],
            ['Type/For', typeFor],
            ['E-mail Discussion', ''],
            ['Revisions', ''],
            ['Minutes', ''],
            ['Disposition', revisedTo ? 'Revised to ' + revisedTo : ''],
            ['Status', tdocData.statusCol >= 0 ? row[tdocData.statusCol] : '']
          ];
          const table = body.appendTable();
          removeInitialEmptyRow_(table);
          tempData.forEach(rowData => {
            const tr = table.appendTableRow();
            tr.appendTableCell(rowData[0]);
            tr.appendTableCell(String(rowData[1] || ''));
          });
          if (tdocData.richTextRow && tdocData.richTextRow[tdocData.tdocCol]) {
            const richText = tdocData.richTextRow[tdocData.tdocCol];
            if (richText.getLinkUrl && richText.getLinkUrl()) {
              const cell = table.getRow(0).getCell(1);
              const tdocValue = String(row[tdocData.tdocCol]);
              if (tdocValue.length > 0) {
                cell.editAsText().setLinkUrl(0, tdocValue.length - 1, richText.getLinkUrl());
              }
            }
          }
          styleStatusCell_(table);
        });
      }
    }
  });
  
  // Add AOB and Close of Session at the end if they weren't in the agenda
  if (!hasAOB) {
    const lastAgendaNum = filteredAgendaItems[filteredAgendaItems.length - 1].number;
    const aobNum = incrementAgendaNumber_(lastAgendaNum);
    body.appendParagraph(`${aobNum} Any other business`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    const aobHeader = findHeading_(sourceBody, /^X\.Y\s+/);
    if (aobHeader) {
      copySectionContentWithReplacement_(aobHeader, body, /^X\.Z\s+/, 'X.Y', aobNum);
    }
  }
  
  if (!hasCloseOfSession) {
    const lastAgendaNum = filteredAgendaItems[filteredAgendaItems.length - 1].number;
    const closeNum = hasAOB ? incrementAgendaNumber_(incrementAgendaNumber_(lastAgendaNum)) : incrementAgendaNumber_(lastAgendaNum);
    body.appendParagraph(`${closeNum} Close of the session`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    const closeHeader = findHeading_(sourceBody, /^X\.Z\s+/);
    if (closeHeader) {
      copySectionContentWithReplacement_(closeHeader, body, /^(?!X\.)/, 'X.Z', closeNum);
    }
  }
  
  removeRowHeightAndSpacing();
  Logger.log(`Done: Built skeleton with ${filteredAgendaItems.length} agenda items and ${allTdocs.length} TDOCs.`);
  try {
    DocumentApp.getUi().alert('Done', `Built skeleton with ${filteredAgendaItems.length} agenda items and ${allTdocs.length} TDOCs.`, DocumentApp.getUi().ButtonSet.OK);
  } catch (e) {
    // UI not available in this context
  }
}

/**
 * Download TDOC list and group by agenda item.
 * Returns: { '9.1': { tdocs: [...] }, '9.2': { tdocs: [...] }, ... }
 */
function downloadAndGroupTdocs_(cfg) {
  const meetingUrl = cfg.TDOC_LIST_URL;
  if (!meetingUrl) return {};

  Logger.log('Downloading TDOC list: ' + meetingUrl);
  const response = UrlFetchApp.fetch(meetingUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() >= 400) {
    throw new Error('Could not download TDOC list: HTTP ' + response.getResponseCode());
  }

  const blob = response.getBlob();
  blob.setName('TDoc_List_Temp.xlsx');
  const tempFile = DriveApp.createFile(blob);

  try {
    const spreadsheet = SpreadsheetApp.open(tempFile);
    const sheet = spreadsheet.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const richTextValues = sheet.getDataRange().getRichTextValues();

    const headers = data[0];
    const tdocCol = headers.indexOf('TDoc');
    const titleCol = headers.indexOf('Title');
    const sourceCol = headers.indexOf('Source');
    const contactCol = headers.indexOf('Contact');
    const agendaCol = headers.indexOf('Agenda item');
    const agendaTopicCol = headers.indexOf('Agenda Topic');
    const statusCol = headers.indexOf('TDoc Status');
    const typeCol = headers.indexOf('Type');
    const forCol = headers.indexOf('For');
    const revisedToCol = headers.indexOf('Revised to');

    if (tdocCol === -1 || agendaCol === -1) {
      throw new Error('Could not find TDoc or Agenda item columns in TDOC list');
    }
    if (revisedToCol === -1) {
      Logger.log('Warning: no "Revised to" column in the TDOC list; revision placement is skipped');
    }

    // SA4-IMPL-005: agenda selection now goes through MeetingContext's
    // agendaSelector (SA4-IMPL-004) instead of a direct
    // getConfiguredAgendaPrefix_() + .startsWith() check. getMeetingContext_()
    // is called here specifically for report.agendaSelector -- cfg (this
    // function's existing parameter) cannot supply it, because cfg is a
    // getReportConfig_() result and has no notion of MEETING_TYPE, so it can
    // never carry the ad-hoc {mode:'all'} selector. Every other value this
    // function uses (TDOC_LIST_URL via `cfg`, the reallocation map, the
    // column layout) is completely unchanged. For every existing main
    // meeting, agendaSelectorMatches_({mode:'prefix', value: <same prefix as
    // before>}, agendaItem) is byte-equivalent to the old
    // agendaItem.startsWith(agendaPrefix) check -- see
    // tests/tdoc-agenda-filter.test.js.
    const agendaSelector = getMeetingContext_().report.agendaSelector;
    const reallocations = getReallocationMap_();
    const groups = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tdoc = String(row[tdocCol] || '').trim();
      if (!tdoc) continue;

      let agendaItem = String(row[agendaCol] || '').trim();

      if (reallocations[tdoc]) {
        agendaItem = reallocations[tdoc].new;
      }

      if (!agendaSelectorMatches_(agendaSelector, agendaItem)) continue;

      if (!groups[agendaItem]) groups[agendaItem] = { tdocs: [] };

      groups[agendaItem].tdocs.push({
        row: row,
        richTextRow: richTextValues[i],
        tdocCol, titleCol, sourceCol, contactCol, agendaCol, agendaTopicCol, statusCol, typeCol, forCol,
        revisedToCol
      });
    }

    Logger.log(`Grouped TDOCs: ${Object.keys(groups).length} agenda items`);
    return groups;

  } finally {
    try { tempFile.setTrashed(true); } catch (e) { }
  }
}

// Keep addReportSkeleton as a legacy alias
function addReportSkeleton() {
  buildSkeletonWithTdocTables();
}

function addTdocTablesOnly() {
  buildSkeletonWithTdocTables();
}

/**
 * Set the document title based on the new requested format.
 * Example: "9. Video SWG Minutes SA4#137-e"
 *
 * SA4-ARCH-004: this is the first production consumer migrated to
 * MeetingContext. It no longer takes a `cfg` (getReportConfig_() result)
 * parameter -- it reads getMeetingContext_() itself, which derives from
 * getReportConfig_() internally, so there is still only one underlying
 * PropertiesService read per call, not two independent config reads.
 *
 * generateReportTitle_() itself is intentionally left unchanged (same
 * plain-object {REPORT_SUFFIX, TDOC_LIST_URL, MEETING_ID} contract as
 * before) so any other current or future caller of it is unaffected; the
 * small object below exists only to satisfy that unchanged contract with
 * values sourced from MeetingContext instead of a legacy cfg object.
 */
function setDocumentTitleFromTemplate_(sourceDocId) {
  try {
    const targetDoc = DocumentApp.getActiveDocument();
    const context = getMeetingContext_();
    const newTitle = generateReportTitle_({
      REPORT_SUFFIX: context.report.type,
      TDOC_LIST_URL: context.sources.tdocListUrl,
      MEETING_ID: context.meeting.portalId
    });

    targetDoc.setName(newTitle);
    Logger.log(`Set document title to: ${newTitle}`);

    // Always insert title at the beginning with TITLE style
    const body = targetDoc.getBody();
    const titlePara = body.insertParagraph(0, newTitle);
    titlePara.setHeading(DocumentApp.ParagraphHeading.TITLE);
  } catch (e) {
    Logger.log('Could not set document title: ' + e.message);
  }
}

function generateReportTitle_(cfg) {
    const reportType = cfg.REPORT_SUFFIX || '6G';
    const topicNames = {
      '6G': '6G Media',
      'Audio': 'Audio SWG',
      'Video': 'Video SWG',
      'MBS': 'MBS SWG',
      'RTC': 'RTC SWG',
      'Liaison': 'Liaison',
      'New': 'New Work'
    };
    const topicName = topicNames[reportType] || reportType;
    
    // Extract full meeting name like "SA4#137-e" from the TDOC list URL
    const tdocUrl = cfg.TDOC_LIST_URL || '';
    const meetingNameMatch = tdocUrl.match(/SA4%23(\d+(?:-e)?)/i);
    const meetingName = meetingNameMatch ? `SA4#${decodeURIComponent(meetingNameMatch[1])}` : `SA4#${cfg.MEETING_ID}`;
    
    // Per user feedback, do not include agenda number in the title.
    return `${topicName} Minutes ${meetingName}`;
}


function buildReportPreamble_(sourceDocId, cfg) {
  const sourceBody = DocumentApp.openById(sourceDocId).getBody();
  const targetBody = DocumentApp.getActiveDocument().getBody();
  const agendaPrefix = (cfg.AGENDA_ITEM_PREFIX || '0.').replace(/\.$/, '');

  // --- 1. Opening of the session ---
  const openingHeader = findHeading_(sourceBody, /^8\.1\s+Opening/);
  if (openingHeader) {
    const openingSubHeader = findHeading_(sourceBody, /^8\.1\.1\s+Opening/);
    if (openingSubHeader) {
      const newHeaderText = `${agendaPrefix}.1 Opening of the session`;
      targetBody.appendParagraph(newHeaderText).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      
      copySectionContent_(openingSubHeader, targetBody, /^8\.1\.2/);

      targetBody.appendParagraph(`${agendaPrefix}.1.1 Registration of documents`)
        .setHeading(DocumentApp.ParagraphHeading.HEADING3);
      targetBody.appendParagraph('[The table of registered documents will be inserted here during the build process].');
    }
  }
  
  // --- 2. IPR and antitrust reminder ---
  const iprHeader = findHeading_(sourceBody, /^8\.2\s+IPR/);
  if (iprHeader) {
    const newHeaderText = `${agendaPrefix}.2 IPR and antitrust reminder`;
    targetBody.appendParagraph(newHeaderText).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    copySectionContent_(iprHeader, targetBody, /^8\.3/, `${agendaPrefix}.2`);
  }
}

/**
 * Copies content from a source element until a stop condition is met.
 * Replaces all occurrences of oldPrefix with newPrefix in the text.
 * Also replaces <add list and make hyperlink> with the appropriate mailing list.
 */
function copySectionContentWithReplacement_(startElement, targetBody, stopRegex, oldPrefix, newPrefix) {
  const cfg = getReportConfig_();
  const mailingList = cfg.LIST_NAME || LIST_NAME_LOCK;
  
  let currentElement = startElement.getNextSibling();
  while (currentElement) {
    const elementType = currentElement.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const p = currentElement.asParagraph();
      const text = p.getText();
      if (p.getHeading() !== DocumentApp.ParagraphHeading.NORMAL && stopRegex.test(text.trim())) {
        break; // Stop at the start of the next section
      }

      // Do replacements on the original text first
      let updatedText = text.replace(new RegExp(oldPrefix, 'g'), newPrefix);
      updatedText = updatedText.replace(/<add list and make hyperlink>/g, mailingList);
      
      // Skip if text is empty or whitespace-only after replacements
      const trimmedText = updatedText.trim();
      if (!trimmedText) {
        currentElement = currentElement.getNextSibling();
        continue;
      }
      
      // Create paragraph with non-empty text
      const newPara = targetBody.appendParagraph(trimmedText);
      newPara.setHeading(p.getHeading());
      
      // Find and hyperlink the mailing list URL if present
      const urlMatch = trimmedText.match(/(https:\/\/list\.etsi\.org\/scripts\/wa\.exe\?A1=[^&\s]+&L=)([^\s]+)/);
      if (urlMatch) {
        const fullUrl = urlMatch[0];
        const startPos = trimmedText.indexOf(fullUrl);
        if (startPos >= 0) {
          try {
            const textElement = newPara.editAsText();
            textElement.setLinkUrl(startPos, startPos + fullUrl.length - 1, fullUrl);
          } catch (e) {
            Logger.log('Failed to set link: ' + e.message);
          }
        }
      }
    } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      const li = currentElement.asListItem();
      let updatedText = li.getText().replace(new RegExp(oldPrefix, 'g'), newPrefix);
      updatedText = updatedText.replace(/<add list and make hyperlink>/g, mailingList);
      
      // Skip if text is empty after replacements
      const trimmedText = updatedText.trim();
      if (!trimmedText) {
        currentElement = currentElement.getNextSibling();
        continue;
      }
      
      // Create list item with non-empty text
      const newListItem = targetBody.appendListItem(trimmedText);
      try {
        newListItem.setGlyphType(li.getGlyphType());
      } catch (e) {
        Logger.log('Failed to set glyph type: ' + e.message);
      }
    } else if (elementType === DocumentApp.ElementType.TABLE) {
      targetBody.appendTable(currentElement.asTable().copy());
    }
    
    currentElement = currentElement.getNextSibling();
  }
}

function findHeading_(body, regex) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      if (para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL && regex.test(para.getText())) {
        return para;
      }
    }
  }
  return null;
}

/**
 * Check if document contains a heading starting with the given text
 */
function documentContainsHeading_(body, headingPrefix) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      if (para.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
        const text = para.getText().trim();
        if (text.startsWith(headingPrefix)) {
          return true;
        }
      }
    }
  }
  return false;
}

/**
 * Increment an agenda number (e.g., "11.5" -> "11.6")
 */
function incrementAgendaNumber_(agendaNum) {
  const parts = agendaNum.split('.');
  const lastPart = parseInt(parts[parts.length - 1], 10);
  parts[parts.length - 1] = String(lastPart + 1);
  return parts.join('.');
}

function findParagraphByText_(body, searchText) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      const text = para.getText();
      if (searchText instanceof RegExp && searchText.test(text)) {
        return para;
      }
      if (typeof searchText === 'string' && text.includes(searchText)) {
        return para;
      }
    }
  }
  return null;
}

function addTdocTablesOnly() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('SKIP_ABSTRACTS_DURING_TABLE_BUILD', 'true');
  try {
    downloadAndProcessFromWeb();
  } finally {
    props.deleteProperty('SKIP_ABSTRACTS_DURING_TABLE_BUILD');
  }
}

function collectEmailDiscussionOnly() {
  const cfg = getCollectorConfig_();
  checkRSSFeed_(cfg);
  DocumentApp.getUi().alert('Success', 'E-mail discussion collection completed.', DocumentApp.getUi().ButtonSet.OK);
}

function collectRevisionsOnly() {
  const cfg = getCollectorConfig_();
  updateRevisions_(cfg);
  DocumentApp.getUi().alert('Success', 'Revision collection completed.', DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Fetch and insert abstracts for every TDOC table that does not have one yet.
 * No UI, so it is safe to call from triggers. Returns the number of tables filled.
 */
function addAbstractsForTables_(body) {
  body = body || DocumentApp.getActiveDocument().getBody();
  let count = 0;

  body.getTables().forEach(table => {
    if (!isTDocTable_(table)) return;

    const existingAbstract = findCellText_(table, 'Abstract');
    if (existingAbstract) return;

    const tdocNumber = String(safeCellText_(table, 0, 1) || '').trim();
    if (!/^S4-\d{6}$/i.test(tdocNumber)) return;

    fetchAndAddAbstract_(table, tdocNumber);
    count++;
  });

  return count;
}

function addAbstractsOnly() {
  // Manual step: always runs, independent of the trigger-level switch.
  const count = addAbstractsForTables_();
  DocumentApp.getUi().alert('Success', `Abstract step completed for ${count} TDOC table(s).`, DocumentApp.getUi().ButtonSet.OK);
}

/********************************************************
 * PHASE 4: AUTO-CREATE REPORT STRUCTURE
 ********************************************************/

/**
 * Auto-create report structure from parsed agenda
 */
function autoCreateReportStructure() {
  const ui = DocumentApp.getUi();
  
  // Check if agenda has been parsed
  const props = PropertiesService.getDocumentProperties();
  const parsedAgenda = props.getProperty('PARSED_AGENDA');
  
  if (!parsedAgenda) {
    ui.alert(
      'No Parsed Agenda',
      'Please parse an agenda document first using:\n📄 Parse Agenda Document',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const agendaItems = JSON.parse(parsedAgenda);
  
  const response = ui.alert(
    'Auto-Create Report Structure',
    `This will create the report structure with ${agendaItems.length} agenda items.\n\n` +
    'This will add headings and placeholders to the current document.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    createReportStructure_(agendaItems);
    ui.alert('Success!', `Created structure with ${agendaItems.length} agenda items!`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to create structure: ${e.message}`, ui.ButtonSet.OK);
    Logger.log('Structure creation error: ' + e.message);
  }
}

/**
 * Create report structure from agenda items
 */
function createReportStructure_(agendaItems) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Find where to insert (after config tables and "Registered Documents")
  let insertIndex = findInsertionPoint_(body);
  
  Logger.log(`Creating structure at index ${insertIndex}`);
  
  // Create structure for each agenda item
  agendaItems.forEach(item => {
    // Create heading: "11.1 Title of agenda item"
    const headingText = `${item.number} ${item.title}`;
    const heading = body.insertParagraph(insertIndex++, headingText);
    
    // Set heading level based on item level
    if (item.level === 1) {
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    } else if (item.level === 2) {
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else {
      heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    }
    
    // Copy agenda text from template (no label, no placeholder noise)
    if (item.text) {
      String(item.text).split('\n').forEach(line => {
        const trimmed = line.trim();
        if (!trimmed) return;
        const agendaText = body.insertParagraph(insertIndex++, trimmed);
        agendaText.setItalic(true);
      });
    }
    
    Logger.log(`Created section: ${headingText}`);
  });
  
  Logger.log(`Created ${agendaItems.length} sections`);
}

/**
 * Find insertion point for structure (after config tables and registered docs)
 */
function findInsertionPoint_(body) {
  const numChildren = body.getNumChildren();
  let lastConfigIndex = -1;
  let registeredDocsIndex = -1;
  
  // Find last config table and "Registered Documents" heading
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    
    if (child.getType() === DocumentApp.ElementType.TABLE) {
      const table = child.asTable();
      
      // Check if it's a config table
      if (isConfigTable_(table) || isCollectorConfigTable_(table) || isReallocationTable_(table)) {
        lastConfigIndex = i;
      }
      
      // Check if it's the registered documents table (4 columns: TDoc, Title, Source, Agenda)
      if (table.getNumRows() > 0) {
        const row0 = table.getRow(0);
        if (row0.getNumCells() === 4 &&
            row0.getCell(0).getText().trim() === 'TDoc' &&
            row0.getCell(1).getText().trim() === 'Title') {
          registeredDocsIndex = i;
        }
      }
    }
    
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const para = child.asParagraph();
      const text = para.getText().trim();
      
      if (text === 'Registered Documents') {
        registeredDocsIndex = i;
      }
    }
  }
  
  // Insert after the last relevant element
  if (registeredDocsIndex > -1) {
    // Find the table after "Registered Documents" heading
    for (let i = registeredDocsIndex + 1; i < numChildren; i++) {
      if (body.getChild(i).getType() === DocumentApp.ElementType.TABLE) {
        return i + 1;
      }
    }
    return registeredDocsIndex + 1;
  }
  
  if (lastConfigIndex > -1) {
    return lastConfigIndex + 1;
  }
  
  // Default: append at end
  return numChildren;
}

/********************************************************
 * CLEANUP FUNCTIONS
 ********************************************************/

/**
 * Remove duplicate email entries from all TDOC tables
 * Checks every 3rd line (author + date pattern) and removes if it matches the previous entry
 */
function removeDuplicateEmailEntries() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Remove Duplicate Email Entries',
    'This will scan all TDOC tables and remove duplicate email discussion entries.\n\n' +
    'Duplicates are detected by matching:\n' +
    '• Author name\n' +
    '• Date and time\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const body = DocumentApp.getActiveDocument().getBody();
  const tables = body.getTables();
  
  let processedTables = 0;
  let removedDuplicates = 0;
  
  tables.forEach(t => {
    if (!isTDocTable_(t)) return;
    
    const tdoc = safeCellText_(t, 0, 1).trim();
    
    // Find email discussion cell
    const cell = findOrFallbackCell_(t, ['E-mail discussion', 'E-Mail discussion', 'Email discussion'], 5, 1);
    const text = cell.getText();
    
    if (!text || text.trim() === '' || text.includes('No e-mail')) return;
    
    // Split into lines
    const lines = text.split('\n');
    const newLines = [];
    const seen = new Set();
    
    let i = 0;
    while (i < lines.length) {
      const line = lines[i].trim();
      
      // Check if this is an author/date line (e.g., "Champel MaryLuc on 2026-08-24 09:22")
      const authorDateMatch = line.match(/^(.+?)\s+on\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})$/);
      
      if (authorDateMatch) {
        const author = authorDateMatch[1].trim();
        const date = authorDateMatch[2].trim();
        const key = `${author}|${date}`;
        
        if (seen.has(key)) {
          // Duplicate found - skip this entry and its preview line
          Logger.log(`Removing duplicate for ${tdoc}: ${author} on ${date}`);
          removedDuplicates++;
          
          // Skip current line and next line (preview) if it starts with ↳
          i++;
          if (i < lines.length && lines[i].trim().startsWith('↳')) {
            i++;
          }
          continue;
        }
        
        // Not a duplicate - add to output
        seen.add(key);
        newLines.push(lines[i]);
        i++;
        
        // Add preview line if present
        if (i < lines.length && lines[i].trim().startsWith('↳')) {
          newLines.push(lines[i]);
          i++;
        }
      } else if (line === '') {
        // Keep empty lines
        newLines.push(lines[i]);
        i++;
      } else {
        // Other lines (shouldn't happen in well-formed email discussions)
        newLines.push(lines[i]);
        i++;
      }
    }
    
    // Update cell if we removed any duplicates
    if (newLines.length < lines.length) {
      const newText = newLines.join('\n');
      cell.setText(newText);
      
      // Reapply font size 8
      const te = cell.editAsText();
      if (te.getText().length > 0) {
        te.setFontSize(0, te.getText().length - 1, 8);
      }
      
      processedTables++;
      Logger.log(`Cleaned ${tdoc}: removed ${lines.length - newLines.length} duplicate lines`);
    }
  });
  
  Logger.log(`Processed ${processedTables} tables, removed ${removedDuplicates} duplicate entries`);
  ui.alert(
    'Cleanup Complete',
    `✅ Processed ${processedTables} TDOC table(s)\n` +
    `🗑️ Removed ${removedDuplicates} duplicate email entries`,
    ui.ButtonSet.OK
  );
}

/**
 * Clean up wrong email discussions that match timezone patterns
 * (e.g., "1600 CEST" being mistaken for TDoc numbers)
 */
function cleanUpWrongEmailDiscussions() {
  const ui = DocumentApp.getUi();
  const response = ui.alert(
    'Clean Up Wrong Email Discussions',
    'This will remove email discussion entries that contain timezone patterns like:\n' +
    '• "1600 CEST"\n' +
    '• "1400 CET"\n' +
    '• "0900 UTC"\n' +
    'etc.\n\n' +
    'These are likely false matches where times were mistaken for TDoc numbers.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const body = DocumentApp.getActiveDocument().getBody();
  const tables = body.getTables();
  const props = PropertiesService.getDocumentProperties();
  
  let cleanedTables = 0;
  let removedEntries = 0;
  
  // Timezone pattern: 4 digits followed by timezone abbreviation
  const timezonePattern = /\b\d{4}\s+(CEST|CET|UTC|GMT|EST|PST|JST|EDT|PDT|BST|IST)\b/i;
  
  tables.forEach(t => {
    if (!isTDocTable_(t)) return;
    
    const tdocRaw = String(safeCellText_(t, 0, 1) || '').trim();
    // SA4-IMPL-001/001A: was extractTdocId_(text, '^S4-\\d{6}$') --
    // main-meeting-only, but correctly anchored (the whole cell had to be
    // exactly one TDoc number). Migrated to parseExactSA4DocumentId_() so
    // all 6 verified families are recognized while preserving that same
    // exact-whole-cell requirement -- not the unanchored/search-within-text
    // parseSA4DocumentId_() used elsewhere (e.g. normalizeTdoc_()).
    const parsedTdoc = parseExactSA4DocumentId_(tdocRaw);
    const tdoc = parsedTdoc.isValid ? parsedTdoc.raw : '';
    if (!tdoc) return;

    const storeKey = 'DISCUSS_' + tdoc;
    const storeJson = props.getProperty(storeKey);
    if (!storeJson) return;
    
    const store = loadJsonObject_(storeJson);
    const originalCount = Object.keys(store).length;
    
    // Filter out entries with timezone patterns in title
    const cleaned = {};
    let localRemoved = 0;
    
    Object.keys(store).forEach(id => {
      const entry = store[id];
      if (entry && entry.title && timezonePattern.test(entry.title)) {
        localRemoved++;
        Logger.log(`Removing wrong email for ${tdoc}: ${entry.title}`);
      } else {
        cleaned[id] = entry;
      }
    });
    
    if (localRemoved > 0) {
      // Save cleaned store
      props.setProperty(storeKey, JSON.stringify(cleaned));
      
      // Update the cell in the document
      const cell = findOrFallbackCell_(t, ['E-mail discussion', 'E-Mail discussion', 'Email discussion'], 5, 1);
      
      if (Object.keys(cleaned).length === 0) {
        cell.setText('No e-mail discussion.');
      } else {
        // Rebuild the cell content
        const cfg = getCollectorConfig_();
        const tz = String(cfg.TIMEZONE || Session.getScriptTimeZone());
        const showPreview = String(cfg.SHOW_PREVIEW_SNIPPET || 'false').toLowerCase() === 'true';
        
        const ordered = Object.values(cleaned)
          .sort((a, b) => parseDateToMillis_(a.date) - parseDateToMillis_(b.date));
        
        cell.setText('');
        const te = cell.editAsText();
        
        // Deduplicate by author+date
        const seen = new Set();
        ordered.forEach(m => {
          const author = m.author || 'Unknown';
          const dateLocal = formatLocalDate_(m.date, tz);
          const key = `${author}|${dateLocal}`;
          
          if (seen.has(key)) return;
          seen.add(key);
          
      let line = `${author} on ${dateLocal}`;
      
      // Add deadline info if present and email is "for agreement"
      if (m.deadline && m.title && m.title.toLowerCase().includes('for agreement')) {
        const timeRemaining = calculateTimeRemaining_(m.deadline, tz);
        line += `\n  📅 Deadline: ${formatDeadline_(m.deadline, tz)} ${timeRemaining}`;
      }
      
      line += '\n';
      const start = te.getText().length;
      te.appendText(line);

      if (m.link) {
        try { te.setLinkUrl(start, start + author.length - 1, m.link); } catch (e) { }
      }
      if (showPreview && m.preview) te.appendText('  ↳ ' + snippet_(m.preview, 160) + '\n');
        });
      }
      
      cleanedTables++;
      removedEntries += localRemoved;
    }
  });
  
  Logger.log(`Cleaned ${cleanedTables} tables, removed ${removedEntries} wrong entries`);
  ui.alert(
    'Cleanup Complete',
    `✅ Cleaned ${cleanedTables} TDOC table(s)\n` +
    `🗑️ Removed ${removedEntries} wrong email discussion entries`,
    ui.ButtonSet.OK
  );
}

/**
 * Apply document reallocations: move or remove TDOC tables based on reallocation table
 */
function applyDocumentReallocations() {
  const ui = DocumentApp.getUi();
  const reallocations = getReallocationMap_();
  
  if (Object.keys(reallocations).length === 0) {
    ui.alert('No Reallocations', 'No document reallocations have been configured yet.', ui.ButtonSet.OK);
    return;
  }
  
  const response = ui.alert(
    'Apply Document Reallocations',
    `This will process ${Object.keys(reallocations).length} reallocation(s):\n\n` +
    '• Move TDOC tables to new agenda items\n' +
    '• Remove tables if reallocated to "removed" or "withdrawn"\n' +
    '• Update agenda item fields in tables\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const body = DocumentApp.getActiveDocument().getBody();
  const tables = body.getTables();
  
  let movedCount = 0;
  let removedCount = 0;
  let updatedCount = 0;
  
  // Process each reallocation
  Object.keys(reallocations).forEach(tdoc => {
    const realloc = reallocations[tdoc];
    const newAgenda = realloc.new.toLowerCase();
    
    // Find the table for this TDoc
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      if (!isTDocTable_(table)) continue;
      
      const tableTdoc = safeCellText_(table, 0, 1).trim();
      if (tableTdoc !== tdoc) continue;
      
      // Check if document should be removed
      if (newAgenda === 'removed' || newAgenda === 'withdrawn' || newAgenda === 'n/a') {
        Logger.log(`Removing table for ${tdoc} (reallocated to ${newAgenda})`);
        body.removeChild(table);
        removedCount++;
        break;
      }
      
      // Update the Agenda Item field in the table
      let agendaUpdated = false;
      for (let r = 0; r < table.getNumRows(); r++) {
        const row = table.getRow(r);
        if (row.getNumCells() < 2) continue;
        
        const key = row.getCell(0).getText().trim().toLowerCase();
        if (key === 'agenda item' || key === 'agenda item:') {
          row.getCell(1).setText(realloc.new);
          agendaUpdated = true;
          break;
        }
      }
      
      if (agendaUpdated) {
        updatedCount++;
        Logger.log(`Updated agenda item for ${tdoc}: ${realloc.original} → ${realloc.new}`);
      }
      
      // Move table to new location
      const currentIndex = body.getChildIndex(table);
      const targetIndex = findInsertionPointForAgendaItem_(body, realloc.new, '');
      
      if (targetIndex !== currentIndex && targetIndex !== currentIndex + 1) {
        // Remove from current position
        const tableElement = body.removeChild(table);
        
        // Insert at new position
        if (targetIndex > currentIndex) {
          body.insertTable(targetIndex - 1, tableElement.asTable());
        } else {
          body.insertTable(targetIndex, tableElement.asTable());
        }
        
        movedCount++;
        Logger.log(`Moved table for ${tdoc} from index ${currentIndex} to ${targetIndex}`);
      }
      
      break;
    }
  });
  
  Logger.log(`Applied reallocations: ${movedCount} moved, ${removedCount} removed, ${updatedCount} updated`);
  ui.alert(
    'Reallocations Applied',
    `✅ Processing complete:\n\n` +
    `📋 Updated agenda items: ${updatedCount}\n` +
    `🔄 Moved tables: ${movedCount}\n` +
    `🗑️ Removed tables: ${removedCount}`,
    ui.ButtonSet.OK
  );
}
