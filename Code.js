/*******************************
 * SA4 Report Generator + Email/Revisions Collector
 * Version: 2.18.1 (2026-10-05)
 * - NO global name collisions
 * - RSS/A1 + Revisions restored
 * - Agenda Item rows preserved/merged
 * - Central Google Docs Editor Add-on (ADDON-002..006): background
 *   scheduler, document context/state abstraction, live-verified
 *
 * CHANGELOG
 * 2.18.1 (2026-10-05)
 *   - Fixed: a TDoc that entered the report while it was only reserved
 *     never got the hyperlink of its TDoc number, in its table or in the
 *     registration table: the update wrote it only when it inserted the
 *     table. The update now adds a missing link once the TDoc list has
 *     one. A link that is already there is never replaced.
 *   - Fixed: such a TDoc did not get its abstract either. The Reviewer's
 *     "no summary" answer from before the upload blocked it for 24 hours,
 *     and with abstracts on update switched off it was never asked for.
 *     When a TDoc comes to be in the report as uploaded -- an existing
 *     table gets its link, or a table is inserted with one -- the update
 *     now asks the Reviewer for its abstract once, whatever that setting
 *     is; before that, for an existing table, the answer from before the
 *     upload is forgotten. "No summary" is cached as always, and no later
 *     update asks again on account of the upload.
 *   - Changed: the abstract sweep of an update (abstracts on update
 *     switched on) leaves out TDocs the list shows as not uploaded yet,
 *     and those already asked for in the same update. "Update Abstracts"
 *     and Full Build ask for every table without an abstract, as before.
 *   - Unchanged: the status rules, Minutes, Disposition, Title, Source,
 *     Contact, Type/For, Agenda Item and existing abstracts.
 * 2.18.0 (2026-10-02)
 *   - Added (ad-hoc reports only): sessions. An ad-hoc report can describe
 *     the meetings of its series as sessions in "Configure Sessions…". A
 *     session is one logical meeting and may span several days: a first
 *     day, optionally a last day, optional planned times (the start on the
 *     first day, the end on the last) and an optional label. Stored as
 *     ADHOC_SESSIONS with stable session ids; a session of one day has no
 *     end date. Everything below is active only in an ad-hoc report that
 *     has sessions configured.
 *   - Added: "Import Teams Attendance…" reads the attendance report
 *     Microsoft Teams exports for one meeting, shows a preview, and after
 *     confirmation stores the summary Teams states and a name / company /
 *     e-mail list. A session has one attendance record per day: an export
 *     is accepted when its date is one of the days of the session, and
 *     importing a day again replaces that day only. The report gets one
 *     generated "Attendance" section at the end of the closing section
 *     (statistics and an attendee table per session and day); Company
 *     cells completed by hand are kept. Participant IDs,
 *     dial-in numbers, join / leave times and individual durations are
 *     never stored, shown or logged. "Refresh Attendance Section" writes
 *     the section again from what is stored.
 *   - Added: TDoc sessions. Each TDoc is assigned to the first session
 *     whose cut-off (planned end on its last day, else the end of that day) is not before
 *     its "Uploaded" time in the TDoc list. The Portal records that time in
 *     UTC; it is converted to the time zone of the script. "Assign TDoc
 *     Sessions…" shows the result and takes manual additions and
 *     replacements (ADHOC_TDOC_SESSIONS). The registration table gets a
 *     fifth column "Session" when a report is built with sessions, or
 *     through the explicit action in that dialog -- never by an update.
 *     downloadAndGroupTdocs_() adds the upload time to each TDoc,
 *     createSummaryTable_() and updateRegisteredDocumentsTable_() accept
 *     the fifth column, continuousUpdateCore_() keeps it up to date.
 *   - Added: "Edit Opening Details…" -- Chair, minute taker(s) and an
 *     administrative note per session (ADHOC_SESSION_OPENING), written as
 *     one generated "Session administration" section before "Registration
 *     of Documents". With sessions configured, Build from Scratch writes
 *     that section and no longer the line "<Chair> opens the session on
 *     ... CEST."
 *   - Added: "Report Status Summary" ends with a "Sessions and attendance"
 *     block for such a report, and "Post-meeting Statistics…" shows the
 *     same per session: counts, and the values Teams states. Nothing is
 *     calculated from the attendance, nothing is stored for it, and both
 *     views are read-only. For this, analyzeReportStatus() calls
 *     adhocSessionStatusLines_() before its alert.
 *   - Added (template runtime only): the submenu "Sessions and Attendance"
 *     with these six actions, shown in an ad-hoc report. Build from Scratch
 *     keeps the imported attendance and the opening details, and its
 *     confirmation says so.
 *   - Adoption: ADHOC_SESSIONS and ADHOC_TDOC_SESSIONS are report state
 *     and are copied; attendance and opening details stay in the
 *     document's own properties.
 *   - Unchanged: main-meeting reports, and ad-hoc reports without
 *     sessions -- same build, same update, same status summary, same menu
 *     apart from the submenu of an ad-hoc report. No existing report gains
 *     a table, a column or text because of this release. Every function
 *     that existed is unchanged except the five named above and
 *     buildSkeletonWithTdocTables().
 * 2.17.4 (2026-10-01)
 *   - Changed (SA4 Report Template runtime only, TEMPLATE-003 stage 1): a
 *     report created from the template has its own menu, "SA4 Report",
 *     built in ReportCreator.js: Report (Build Report from Scratch, Update
 *     Report Now, the three partial updates, status summary), Prepare
 *     Discussion E-mails, Document Reallocation, Automatic Updates,
 *     Configure Meeting, Advanced and Repair, About. The LEGACY submenu and
 *     the technical / duplicate items are no longer shown there; every
 *     function behind them is still in this file. The master template's
 *     menu is called "SA4 Report" too.
 *   - Changed (template runtime only): the build confirmation says that the
 *     document's content is replaced and that minutes entered by hand are
 *     lost. Still one confirmation before any change, no UI call after it.
 *   - Added (template runtime only): "Update Report Now" runs the complete
 *     update of continuousUpdate() and fails visibly: no pop-up after the
 *     document was changed, a completed update ends normally, a failed one
 *     throws. For that, collectorUpdate_() now also RETURNS the failures of
 *     its two steps (it only logged them), and continuousUpdateCore_()
 *     passes them on in its result. Nothing else reads these values:
 *     continuousUpdate(), the trigger and the add-on scheduler behave as
 *     before.
 *   - Changed (all runtimes, text only): the hint under "Fetch abstracts
 *     during each update" in the trigger dialog no longer names a menu
 *     item ("5️⃣ Add Abstracts", which a template report does not have).
 *     It now reads "Abstracts can also be updated manually at any time."
 *   - Unchanged: the CENTRAL add-on / Legacy menu (the menu code below is
 *     again exactly what it was before the template runtime), and every
 *     existing function.
 * 2.17.3 (2026-10-01)
 *   - Fixed: a completed "Run Full Report Build" was reported as FAILED.
 *     Live (T-2026.10.2, report for 86178): all five phases finished in
 *     65 s and the report was saved, but the final "Success" pop-up never
 *     appeared -- that ui.alert() call hung for about 96 s and then threw
 *     "Service Documents failed while accessing document with id ...".
 *     Full Build no longer calls the UI after it has changed the
 *     document. A completed build simply ends (the result is in the
 *     [FULLBUILD] log lines; the execution is "Completed"). A build in
 *     which a phase fails still fails: it throws an error carrying the
 *     summary (failed phase, its message, the phases not run), which Docs
 *     shows and the execution records as "Failed". The confirmation
 *     before the build is unchanged. No build phase changed.
 * 2.17.2 (2026-10-01)
 *   - Fixed: the 6G report family derived the general list 3GPP_TSG_SA_WG4
 *     as its mailing list. Its list is 3GPP_TSG_SA4_FS_6G_MED
 *     (MAILING_LISTS, the one family -> list table). Found live: "Create
 *     New SA4 Report" pre-filled the general list for a 6G meeting. Every
 *     consumer derives from that table, so Configure Meeting, the RSS and
 *     A1 collectors and the discussion e-mail To / Reply-To now use the 6G
 *     list for a 6G report without a Mailing List override, ad-hoc or
 *     main. A saved Mailing List override still wins. The other families
 *     are unchanged.
 * 2.17.1 (2026-10-01)
 *   - Fixed (TEMPLATE-002C): "Run Full Report Build" ran the single-step
 *     menu functions, each of which ends with its own blocking pop-up
 *     ("Done", "E-mail discussion collection completed", ...). The build
 *     therefore stopped four times inside its one 6-minute execution and
 *     waited for a click; live on the first template report (86178) it
 *     ended with "Exceeded maximum execution time" although the work
 *     itself took about a minute. Full Build now runs non-interactive
 *     phases (runFullReportBuildCore_()): one confirmation before, no
 *     pop-up between phases, one summary at the end. A failed phase stops
 *     the build and is reported once. The single-step menu items keep
 *     their own completion pop-ups.
 *   - Fixed: Full Build formatted the whole document twice (inside the
 *     skeleton step and again at the end). It now formats once, after the
 *     enrichment phases; the single-step "Build Skeleton + TDOC Tables"
 *     still formats.
 *   - Fixed: the optional Reviewer API token is read once per execution.
 *     Without a token Full Build skips the abstracts phase (no table scan,
 *     no request) and says so once, instead of one property read and one
 *     "No REVIEWER_API_TOKEN found" line per TDoc table, twice. With a
 *     token the abstracts are fetched in their own phase.
 *   - Added: [FULLBUILD] log lines -- "<phase>: start" and "<phase>:
 *     done|skipped|failed in N ms", plus one total line -- so a timeout
 *     shows which phase was running.
 *   - Changed (decision 2026-10-01): generated discussion e-mails carry
 *     "Reply-To: <list>@list.etsi.org" -- the effective mailing list (an
 *     ad-hoc meeting's Mailing List, else the report-family list; a main
 *     meeting always the family list), normalized like the collector does.
 *     Replies used to go to the sender. From and To are unchanged. A list
 *     that cannot be normalized refuses the export.
 *   - Changed (decision 2026-10-01): the discussion subject no longer
 *     starts with the list/family tag. Old: "[<tag>,<agenda item>,
 *     <deadline>][<TDoc>] Discussion: <title>"; new: "[<agenda item>]
 *     [<deadline>][<TDoc>] Discussion: <title>" (written without spaces
 *     between the brackets). The tag still names the exported files and
 *     the default introduction. Replies to e-mails sent in the old form
 *     are still associated (the collector looks for the TDoc identifier).
 * 2.17.0 (2026-10-01)
 *   - Added (TEMPLATE-002B): this file is also the runtime of the "SA4
 *     Report Template" Google Doc and of the reports copied from it. A
 *     template release bundle adds Release.js (generated, defines
 *     SA4_RELEASE_) and ReportCreator.js. Everything template-specific
 *     here is switched by templateRuntimeRelease_() at call time; without
 *     Release.js (the CENTRAL add-on, Legacy copies) behaviour is
 *     unchanged.
 *   - Template runtime: the master template gets its own menu ("Create New
 *     SA4 Report", "Template Release Info") and refuses report operations
 *     (build, update, trigger, configuration). A report gets the existing
 *     menu without the CENTRAL add-on submenu, with the old "Legacy: Build
 *     Initial Report" / "Legacy: Update All" in a LEGACY submenu, plus
 *     "Finish Report Setup" (until done) and "About This Report".
 *     Configure Meeting and Run Full Report Build store the creator's
 *     setup information on first use.
 *   - Template runtime: Continuous Update every 15, 30 (default) or 60
 *     minutes (60 = everyHours(1)); the replacement trigger is created
 *     before the previous one is removed, and the installed interval is
 *     shown (Document Property CONTINUOUS_UPDATE_INTERVAL_MINUTES). The
 *     installer lives in ReportCreator.js; the add-on stays hourly only.
 *   - Template runtime (decision 2026-10-01): "Email Collection Start
 *     Date" is a Configure Meeting field again. It is stored in the
 *     existing EMAIL_START_DATE property; without a stored value the
 *     meeting start date is proposed, and the collector falls back to the
 *     meeting start date (else 2026-08-21). Outside the template runtime
 *     the dialog and the fallback are as in 2.16.0.
 * 2.16.0 (2026-10-01)
 *   - TEMPLATE-002A (Legacy parity): the fixes accepted in the Legacy
 *     bound script (sa4-report-legacy b95e06e, live on 86178) that this
 *     file did not have. Legacy's own regression suites now run here
 *     (tests/legacy-parity-*.test.js).
 *   - Fixed (Legacy BUGFIX-LEGACY-002): duplicate TDoc tables. A TDoc
 *     listed under two agenda groups in one Continuous Update run was
 *     inserted twice, and every further run of "Legacy: Build Initial
 *     Report" duplicated every TDoc table. A TDoc inserted in a run now
 *     counts as existing at once, and the web-sheet import skips a TDoc
 *     that already has a table. Existing duplicates are left alone.
 *   - Fixed (Legacy LEGACY-0099): the web-sheet import filed a TDoc
 *     reallocated to Removed/Withdrawn/N/A under a literal "Removed" item;
 *     it now skips it, like the build and Continuous Update already did
 *     (ADDON-008A1b).
 *   - Fixed (Legacy BUGFIX-LEGACY-003): e-mail collector. A stored
 *     EMAIL_START_DATE is used (2026-08-21 remains the fallback). The A1
 *     archive reads the same list as RSS instead of always the general
 *     SA4 list. A Collector Configuration table that only repeats the
 *     family default no longer masks an ad-hoc meeting's Mailing List; an
 *     explicit LIST_NAME also sets the RSS URL and an explicit RSS_URL_V2
 *     also sets the archive list. A Mailing List saved as
 *     "<list>@list.etsi.org" is read as that list.
 *   - Unchanged on purpose: the meeting-neutral discussion e-mail
 *     introduction, the subject tag derived from the mailing list, the
 *     Meeting Configuration dialog (no e-mail start date field), and the
 *     hourly-only trigger interval of the add-on runtime.
 * 2.15.2 (2026-09-30)
 *   - Fixed: "Manage Auto-Update Trigger" offered 15/30-minute intervals
 *     that CENTRAL cannot create -- add-on time-driven triggers run at most
 *     once per hour (ADDON-001B), so "Every 30 minutes" failed live. The
 *     dialog now offers only "Every hour (Slow meeting)" -> everyHours(1);
 *     15/30 are refused before the running trigger is touched.
 *   - Fixed: discussion e-mail export copied Google Docs links to places in
 *     the report itself (tabs, headings: "?tab=t...", "#heading=h...")
 *     verbatim, unusable outside the editor. They are now resolved against
 *     the report's URL (https://docs.google.com/document/d/<id>/edit);
 *     links with a scheme (https:, mailto:, ...) are unchanged.
 * 2.15.1 (2026-09-30)
 *   - Fixed: "Manage Auto-Update Trigger" -> "Every hour (Slow meeting)"
 *     failed with "The value you passed to everyMinutes was invalid": the
 *     dialog's 60 went straight to everyMinutes(60). Each offered interval
 *     now maps explicitly (CONTINUOUS_TRIGGER_INTERVALS_): 15/30 minutes ->
 *     everyMinutes(15/30), every hour -> everyHours(1). An unsupported,
 *     malformed or missing interval is refused before anything changes (the
 *     running trigger was previously deleted before the failing call).
 * 2.15.0 (2026-09-30)
 *   - Added (ADDON-009): "📧 EMAIL EXPORT" -> "Prepare TDoc Discussion
 *     E-mails", the Legacy discussion e-mail exporter (LEGACY-UPGRADE-006,
 *     Legacy 80ab081) ported unchanged: one Outlook-ready .eml per selected
 *     eligible TDoc (Approved/Agreed/reserved never offered) plus a ZIP of
 *     them in the Drive folder "SA4 Report Email Exports"; nothing is sent.
 *     Subject "[<tag>,<agenda item>,YY-MM-DD-HHmmTZ][<TDoc>] Discussion:
 *     <title>". From is the new "Discussion E-mail Sender" setting
 *     (DISCUSSION_EMAIL_SENDER, never derived); To is the meeting's mailing
 *     list as a list.etsi.org address (an ad-hoc meeting's saved Mailing
 *     List, else the report-family list). The subject tag is that list's
 *     name without the SA4 prefix (3GPP_TSG_SA4_FS_6G_MED -> FS_6G_MED,
 *     3GPP_TSG_SA_WG4_MBS -> MBS; the general SA4 list uses the report
 *     family's SWG name). Every selection is validated and built before any
 *     file is written, so a refused request leaves nothing in Drive.
 *     Meeting-neutral defaults: the introduction says "the upcoming
 *     meeting", and the discussion deadline has no prefilled date/time
 *     (Legacy's 2026-10-15 15:00) -- the user enters it.
 * 2.14.3 (2026-09-30)
 *   - Fixed (ADDON-008A2): incoming e-mail was never associated with ad-hoc
 *     TDocs. checkRSSFeed_() now identifies report tables with the
 *     registered SA4 families (parseExactSA4DocumentId_) instead of the
 *     legacy '^S4-\d{6}$' TDOC_ID_REGEX default, and associates a message
 *     with every recognized full identifier anywhere in its subject
 *     (findSA4DocumentIdsInText_, canonical, revision suffix -> base), so
 *     e.g. '[ AMD_ARCH_Ph2-MED] S4aI260082 26501-CR0124-B "..."' now
 *     matches S4aI260082. Only a subject without any full identifier falls
 *     back to the legacy short-number parse, and only for main-meeting S4-
 *     tables (a bare "082" no longer reaches S4aI260082/S4aA260082).
 *   - Fixed (ADDON-008A2): the collector reads the configured list -- an
 *     ad-hoc meeting's saved MAILING_LIST override (validated as a list
 *     name), else the report-family list; main meetings unchanged.
 * 2.14.2 (2026-09-30)
 *   - Fixed (ADDON-008A1b): saved Document Reallocations were lost by a
 *     full rebuild -- the build cleared the document before reading the
 *     reallocation table and never recreated it. The build now reads them
 *     first and treats them as report overrides: each TDoc's effective
 *     allocation (reallocation, else source) decides the agenda.csv
 *     sections and where the TDoc is placed; removed/withdrawn/n/a TDocs are
 *     not placed and never cause a section. Agenda parsing, the TDoc
 *     download and the reallocation check now run before the document is
 *     cleared, so an unusable destination refuses the build with the
 *     report unchanged. The table is restored after the rebuild. One shared
 *     interpretation (interpretReallocationTarget_) serves the build, the
 *     TDoc grouping and Apply Document Reallocations; Apply leaves a table
 *     that is already under its destination heading in place.
 * 2.14.1 (2026-09-30)
 *   - Fixed (ADDON-008A1): "Apply Document Reallocations" threw
 *     "BODY_SECTION can't be cast to TABLE." on every real table move
 *     (Body.removeChild() returns the body, not the removed table) -- after
 *     the table had already been removed. Latent since the initial commit;
 *     first hit live on 86172 (S4aI260081, 3.7 -> 2.7). Entries are now
 *     resolved and validated first (no change at all if any entry is
 *     invalid), and a move inserts an updated copy before removing the
 *     original. A destination with no heading in the report is an error
 *     instead of a heading appended at the document end; a TDoc with no
 *     table in the report is skipped and listed.
 * 2.14.0 (2026-09-30)
 *   - Added (ADDON-008A): ad-hoc agenda + TDoc-list discovery. Agenda
 *     structure precedence for ad-hoc meetings: saved AGENDA_TDOC >
 *     discovered unambiguous agenda TDoc > the series agenda.csv
 *     (<series>/Agenda/agenda.csv, new AGENDA_CSV_URL property), accepted
 *     only when every agenda item/description pair of the meeting's own
 *     TDoc list matches it; the report uses the CSV sections those TDocs
 *     fall under, through the existing itemList projection and skeleton
 *     path. The build re-validates the CSV before the document is cleared.
 *     "Discover Agenda / TDocs" also proposes the meeting-ID Portal
 *     document list (GenerateDocumentList.aspx) as the TDoc List URL after
 *     checking it is a workbook with the importer's columns; a saved URL
 *     still wins. The ad-hoc agenda readiness rule accepts an agenda TDoc
 *     or this meeting's own agenda.csv. Main meetings are unchanged.
 *     Real cases: 86172 (MBS, CSV fallback), 85916 (Audio, agenda TDoc
 *     S4aA260090 wins over a 0-byte CSV).
 * 2.13.0 (2026-09-23)
 *   - Added (ADDON-002..004): a document execution-context abstraction
 *     (getReportDocument_/getReportBody_/getReportDocumentId_) and a
 *     report-state-store abstraction (getReportStateStore_) with two
 *     backends -- Document Properties (legacy bound-script/interactive,
 *     unchanged default) and a central, documentId-namespaced Script
 *     Properties backend ('SA4_STATE|<documentId>|<key>') for background
 *     execution. Every background-reachable function in
 *     continuousUpdateCore_()'s call graph now threads this context
 *     instead of assuming an active document/Document Properties.
 *   - Added (ADDON-003): a central document registry
 *     ('SA4_REGISTRY_INDEX' + one 'SA4_REGISTRY_DOC|<documentId>' entry
 *     per document -- never one unbounded blob) and an explicit,
 *     idempotent adoption operation (adoptReportDocumentForAddon_())
 *     that copies a document's report state from Document Properties
 *     into the central backend, verifies the copy, and only then
 *     registers it. Document Properties are never modified or deleted by
 *     adoption.
 *   - Added (ADDON-004): continuousUpdateForDocument_(documentId) -- the
 *     per-document background unit of work -- and runAddonScheduler_(), a
 *     central scheduler handler (not yet wired to a live trigger at this
 *     point) that acquires a ScriptLock, computes due documents
 *     (never-run-first, then oldest-lastRunAt-first, stable tie-break),
 *     enforces a conservative runtime budget (deferring, not failing,
 *     documents once exhausted), and isolates each document's failure
 *     from the others. continuousUpdateCore_() now returns
 *     {success, error}; a background run's lastRunAt only advances when
 *     that reports success.
 *   - Added (ADDON-005): add-on entry points (onInstall(e)/onOpen(e), the
 *     existing bound-script menu fully preserved and unchanged), a
 *     separate "CENTRAL ADD-ON" menu (Enable/Disable Automatic Updates,
 *     Set Update Interval [1/2/4/6/12/24h only -- add-on time-driven
 *     triggers cannot run more frequently than hourly], Show Add-on
 *     Status), the central scheduler's own hourly trigger lifecycle
 *     (ensureAddonSchedulerTrigger_/getAddonSchedulerTriggerStatus_/
 *     deleteAddonSchedulerTrigger_, scoped strictly to its own handler
 *     name, never touching unrelated project triggers), and
 *     withAddonScriptLock_() serializing the new interactive
 *     enable/disable/set-interval actions against the scheduler on an
 *     adopted document. Deployed to a separate, new central Apps Script
 *     project (test-deployment install) -- the legacy production
 *     bound-script project and its existing reports were never touched.
 *   - Fixed (live acceptance, ADDON-005B): adoptReportDocumentForAddon_()
 *     now always reads Document Properties directly as its source
 *     (never through the mode-aware state-store seam), so re-adoption on
 *     an already-adopted document correctly re-syncs a changed value
 *     instead of copying central state onto itself.
 *   - Fixed (live acceptance, ADDON-005B): removeEmptyParagraphs_() --
 *     called from removeRowHeightAndSpacing()'s formatting stage -- now
 *     accepts and threads the execution context; it was the one
 *     get-active-document-body call site missed by the original
 *     background-reachability audit, and crashed a real background
 *     scheduler run with "Cannot read properties of null (reading
 *     'getBody')". A full re-audit of every DocumentApp.getActiveDocument()
 *     and getActiveDocumentBodyCounted_() call site found no further gaps.
 *   - Live acceptance evidence (manual scheduler invocation, real scratch
 *     Google Docs, no production report or meeting 85916 involved):
 *       * Failure path: an explicit TDoc-list URL pointing at a
 *         guaranteed-unreachable host reached the background execution
 *         correctly, failed with the expected error,
 *         continuousUpdateCore_() reported failure,
 *         continuousUpdateForDocument_() propagated it, the scheduler
 *         counted it as failed (not succeeded), lastRunAt was NOT
 *         advanced, and a second, independently-registered document was
 *         still processed afterward (failure isolation confirmed live,
 *         not just in unit tests).
 *       * Success path: a synthetic header-only XLSX (columns "TDoc",
 *         "Agenda item", "Revised to", zero data rows) was downloaded and
 *         parsed for real; zero TDocs were found, zero existing TDocs
 *         were found, zero document structural mutations occurred, the
 *         collector stage completed, continuousUpdateCore_() reached its
 *         normal COMPLETE path, the scheduler summary was
 *         {"considered":2,"due":1,"succeeded":1,"failed":0,"skipped":0},
 *         lastRunAt advanced to a real timestamp
 *         (2026-09-23T12:34:47.457Z), and exactly one central scheduler
 *         trigger existed throughout.
 *   - Distribution finding (ADDON-006, research only, no code impact): a
 *     personal (non-Workspace) Google account cannot install a private
 *     Editor Add-on once for all Docs -- confirmed against current Google
 *     documentation that only public Marketplace visibility is available
 *     to consumer accounts (private/unlisted both require a Workspace
 *     domain). Decision: keep the central project + one Apps Script test
 *     deployment, accepting one extra "add this document as a test
 *     document" step per new report document; defer public Marketplace
 *     publication. The central scheduler is unaffected by this choice --
 *     it operates on registered documents via DocumentApp.openById(),
 *     independent of whether the add-on's interactive menu is reachable
 *     in that document.
 * 2.12.0 (2026-09-23)
 *   - Added: a per-run TDoc-table index (buildTdocTableIndex_()) built once
 *     from the same table scan continuousUpdate() already does for
 *     existing-TDoc detection, and threaded through findTdocTable_(),
 *     updateTdocStatus_(), insertNewTdoc_() and rearrangeRevisionTables_().
 *     A no-change incremental update no longer re-scans the whole document
 *     once per existing TDoc; measured in production: body.getTables()
 *     calls dropped from 39 to 5, status-update time from 6,668ms to
 *     1,769ms for a 34-existing-TDoc run.
 *   - Added: shouldReformatAfterUpdate_() gates continuousUpdate()'s
 *     document-wide formatting pass (removeRowHeightAndSpacing()) so a run
 *     that only rewrote existing cell TEXT (status updates, e-mail
 *     discussion, revision rendering) skips it entirely; a run that
 *     inserted/moved a table, or inserted a row into an existing table
 *     (e.g. a newly-fetched Abstract row), still gets it. Measured in
 *     production: eliminated a 16,378ms unconditional cost on a no-change
 *     run. Every other call site (full report build, "Update All", manual
 *     formatting menu items) is untouched and still unconditional.
 *   - Added: LockService.getDocumentLock() concurrency protection.
 *     continuousUpdate() (both the "Continuous Update" menu item and the
 *     time-driven trigger) and "Update Report (During Meeting)" each
 *     acquire a short, non-blocking document lock before any work; a run
 *     that loses the race does no work at all and logs/alerts that another
 *     update is already in progress, instead of mutating the document
 *     concurrently with another run.
 *   - Added: a bounded 24-hour negative cache for the Reviewer API. A
 *     definitive "no summary exists" (404) response is no longer retried
 *     on every subsequent automatic run for the same TDoc; a later
 *     successful summary clears the cached entry. Auth failures, other
 *     4xx, 5xx, and empty-but-200 responses are never cached negative.
 *   - Added: revision-folder anchors are now parsed once per run into a
 *     canonical-TDoc-id bucket (buildRevisionAnchorIndex_()) instead of
 *     being re-parsed once per (TDoc table, anchor) pair; a 34-table/
 *     65-anchor workload goes from up to 2,210 parses down to 65. This is
 *     a complexity/code-quality fix (production measured the old matching
 *     loop at only ~9ms for that workload), not a runtime-bottleneck fix.
 *   - Added: lightweight [PERF] timing/counting instrumentation across
 *     continuousUpdate() and its call graph (perfTimed_/perfTimedAccum_/
 *     perfCount_ + a per-run summary line), used to measure every
 *     improvement above against real production runs.
 *   - Fixed: after a Reviewer negative-cache skip, the automatic-update log
 *     still read "Fetched abstracts for 1 table(s)" even though zero
 *     Reviewer requests occurred. addAbstractsForTables_() now returns a
 *     breakdown (candidate tables processed / Reviewer requests made /
 *     negative-cache skips / abstract rows inserted) so the log accurately
 *     reflects what actually happened. Diagnostics only.
 *   - Investigated, not implemented: positive caching of ETSI A1 mailing-
 *     list archive pages. The archive's actual moderation/indexing-delay
 *     behavior cannot be established from this codebase or from date/URL
 *     alone, so a closed-looking week page cannot be proven immutable;
 *     the existing 24-hour empty-page-only A1 cache is unchanged.
 * 2.11.0 (2026-09-22)
 *   - Fixed: revision-cell rendering could corrupt/duplicate filenames when
 *     an incremental update's text-mutation-then-hyperlink pattern ran
 *     against a live Text reference. renderRevisionsCellContent_() now
 *     builds the full cell text as one string, calls setText() once, then
 *     computes every hyperlink range by pure string arithmetic against the
 *     now-stable text. A corrupted cell from a prior run is automatically
 *     repaired the next time revisions are collected -- no manual cleanup.
 *   - Added: normalizeRevisionKey_() for REVIS_<tdoc> store-level dedup, so
 *     two stored entries that differ only by an incidental encoding/casing
 *     difference in how their URL was captured collapse into one rendered
 *     line instead of duplicating it.
 *   - Added: ad-hoc SA4 draft documents (e.g. "S4aP260071_QCOM.docx") are
 *     now recognized as revisions via the central SA4 TDoc identifier
 *     registry (parseDraftAnchor_()), not a main-meeting-only regex, and
 *     multiple distinct drafts for the same TDoc are all surfaced
 *     deterministically (no silent "latest file wins" behavior).
 *   - Fixed: the revisions/drafts source now reads from
 *     getMeetingContext_().sources.revisionsUrl (the resolved, possibly
 *     ad-hoc, drafts location) instead of the main-meeting-only
 *     cfg.REVISIONS_URL formula, so an ad-hoc meeting's revisions are found
 *     at its real location instead of a fabricated main-meeting URL.
 * 2.10.0 (2026-09-22)
 *   - Added: Reviewer (Contribution Reviewer) API support for every
 *     registered SA4 TDoc family, not just the main-meeting S4-xxxxxx
 *     pattern -- migrated to the same central parseExactSA4DocumentId_()/
 *     SA4_TDOC_FAMILIES registry every other TDoc-identity check uses.
 *     Automatic Abstract retrieval runs both when a new TDoc table is
 *     created and, when "Fetch abstracts during each update" is enabled,
 *     as a batch sweep over existing tables missing an Abstract cell
 *     (addAbstractsForTables_()).
 * 2.9.0 (2026-09-22)
 *   - Added: generic (not meeting-86178-specific) ad-hoc meeting support --
 *     SA4 ad-hoc source URLs, meeting labels/titles, opening/registration
 *     document structure, and a per-meeting mailing-list override, all
 *     resolved through MeetingContext rather than hardcoded to the main
 *     meeting.
 *   - Added: canonical IPR section generation, derived consistently from
 *     the resolved structureBranch for every report type (not just the
 *     plenary/6G template-copy case).
 *   - Added: ad-hoc drafts/revisions folder discovery from Meeting-ID
 *     metadata, feeding the same revision workflow a main meeting uses.
 * 2.8.0 (2026-09-22)
 *   - Added: the production Meeting-ID configuration workflow -- Meeting ID
 *     -> Resolve -> Discover Agenda/TDocs -> Review -> Save. "Resolve"
 *     fetches anonymous, publicly-available 3GPP meeting metadata (dates,
 *     venue, agenda/TDoc source locations) for the entered Meeting ID
 *     without requiring any prior manual configuration; "Discover Agenda/
 *     TDocs" is a separate, explicit step so a slow enrichment fetch never
 *     blocks the core resolve. Explicit configuration overrides remain
 *     fully possible at every step -- resolution only fills in what is not
 *     already configured.
 *   - Fixed: the Resolve/Discover Agenda RPCs never ran when invoked from
 *     the configuration dialog. Apps Script's google.script.run cannot
 *     invoke a function whose name ends in "_" (treated as private); the
 *     dialog was calling the "_"-suffixed internals directly. Added public
 *     wrapper functions (resolveMeetingForConfigDialog(),
 *     discoverAgendaForConfigDialog()) that delegate to the real
 *     implementations, and the dialog's client script now calls the public
 *     names.
 * 2.7.0 (2026-09-22)
 *   - Added: MeetingContext architecture -- a single, normalized read of
 *     "which meeting is this report for" (dates, agenda source, TDoc
 *     source, revisions source, mailing list) that every consumer now
 *     reads from instead of re-deriving main-meeting-only assumptions
 *     independently. Main-meeting behavior is fully preserved: for a main
 *     meeting, every resolved source is identical to the pre-existing
 *     hardcoded formula.
 *   - Added: central SA4 TDoc identifier-family registry
 *     (SA4_TDOC_FAMILIES / parseSA4DocumentId_() / parseExactSA4DocumentId_())
 *     recognizing all 6 verified families (S4-, S4aA, S4aP, S4aV, S4aI,
 *     A4aR), replacing scattered main-meeting-only regexes across TDoc
 *     table creation, status updates, abstracts, and revision matching.
 *   - Added: a MeetingContext source resolver, ad-hoc meeting identity, an
 *     agendaSelector model, and a pure agenda-projection model, so an
 *     ad-hoc meeting's agenda/TDoc/source data can be derived the same way
 *     a main meeting's is, without special-casing individual meetings.
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
  '6G': '3GPP_TSG_SA4_FS_6G_MED',
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

/**
 * ADDON-005: standard Editor Add-on install hook -- Google's documented
 * boilerplate is simply to delegate to onOpen(e), since a freshly-
 * installed add-on's first "open" event needs the same menu onOpen()
 * already builds. Does nothing else; no adoption, no registration, no
 * trigger creation happens here -- those remain explicit user actions
 * (Phase C below), never implicit side effects of installing/opening.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * ADDON-005: accepts the optional add-on event object Apps Script passes
 * to onOpen(e) for an installed add-on (unused by the body below -- the
 * menu itself is IDENTICAL whether this runs as a bound script's simple
 * trigger (e undefined) or an add-on's (e populated), which is exactly
 * why Phase B does not touch anything past this signature: preserving
 * "e" as an accepted-but-unused parameter is what lets the SAME onOpen()
 * serve both the legacy bound deployment and the central add-on project
 * without a second, parallel menu-building function to keep in sync.
 */
function onOpen(e) {
  const ui = DocumentApp.getUi();

  // TEMPLATE-002B / TEMPLATE-003: in the SA4 Report Template runtime
  // (Release.js present) the menus are built in ReportCreator.js -- the
  // master template gets its own minimal menu, a report gets the
  // "SA4 Report" menu. Without Release.js (CENTRAL, Legacy copies) neither
  // branch is taken and the menu below is unchanged.
  if (templateRuntimeRelease_()) {
    if (isTemplateMasterDocument_()) buildTemplateMasterMenu_(ui);
    else buildTemplateReportMenu_(ui);
    return;
  }

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
  
  // ADDON-005: add-on-specific automatic-update controls -- deliberately
  // SEPARATE from "⏰ Manage Auto-Update Trigger" above (the legacy
  // 15/30/60-minute bound-script trigger UI, untouched). This submenu is
  // additive only: every item it contains is new code calling new
  // functions (Phase C/D/E below); nothing here replaces or reinterprets
  // any existing menu item or its zero-context call.
  const addonMenu = ui.createMenu('☁️ CENTRAL ADD-ON (hourly)');
  addonMenu.addItem('▶️ Enable Automatic Updates (this doc)', 'enableAutomaticUpdatesForAddon');
  addonMenu.addItem('⏹️ Disable Automatic Updates (this doc)', 'disableAutomaticUpdatesForAddon');
  addonMenu.addItem('⏱️ Set Update Interval (this doc)', 'setAutomaticUpdateIntervalForAddon');
  addonMenu.addItem('📊 Show Add-on Status (this doc)', 'showAddonSchedulerStatusForAddon');

  // ADDON-009: the Legacy "📧 EMAIL EXPORT" submenu, unchanged (read-only
  // export; see prepareTdocDiscussionEmails()).
  const emailExportMenu = ui.createMenu('📧 EMAIL EXPORT');
  emailExportMenu.addItem('Prepare TDoc Discussion E-mails', 'prepareTdocDiscussionEmails');

  // Add all submenus to main menu
  menu.addSubMenu(setupMenu);
  menu.addSubMenu(reportMenu);
  menu.addSubMenu(docMenu);
  menu.addSubMenu(toolsMenu);
  menu.addSubMenu(formatMenu);
  menu.addSubMenu(emailExportMenu);
  menu.addSubMenu(addonMenu);

  // Legacy functions (for backward compatibility)
  menu.addSeparator();
  menu.addItem('⚠️ Legacy: Update All', 'updateAll');

  menu.addToUi();
}


// =========================================================
// TEMPLATE-002B -- SA4 REPORT TEMPLATE RUNTIME HOOKS
// =========================================================
//
// This file is also the bound script of the "SA4 Report Template" Google
// Doc and of every report copied from it. A template release bundle adds
// two files next to it: Release.js (generated; defines SA4_RELEASE_) and
// ReportCreator.js (creator, first run, template menus). Everything
// template-specific in THIS file asks templateRuntimeRelease_() at call
// time, so without Release.js (the CENTRAL add-on, a Legacy copy, the
// tests) every function below is inert and behaviour is unchanged.
//
// Three kinds of document can run this code:
//   - the master template: its id is SA4_RELEASE_.templateDocumentId. It
//     only creates reports; report operations are refused in it.
//   - a report created from the template: any other document, set up from
//     the creator's bootstrap information on first use.
//   - an ordinary document that happens to contain the code: any other
//     document without bootstrap information. It behaves like a report and
//     is configured by hand, as before.
// The role comes from the document id, never from inherited properties: a
// copy starts with empty property stores (TEMPLATE-001, live verified).

/** The template release this script copy was built from, or null. */
function templateRuntimeRelease_() {
  return typeof SA4_RELEASE_ !== 'undefined' && SA4_RELEASE_ && SA4_RELEASE_.flavor === 'template' ? SA4_RELEASE_ : null;
}

/** 'template' for the master template document, 'report' for anything else. */
function templateDocumentRole_(documentId, release) {
  return release && documentId && documentId === release.templateDocumentId ? 'template' : 'report';
}

function isTemplateMasterDocument_() {
  const release = templateRuntimeRelease_();
  return !!release && templateDocumentRole_(getActiveDocumentIdSafely_(), release) === 'template';
}

/**
 * Refuses a report operation in the master template, so the template never
 * acquires meeting configuration, report content or a trigger. Called by
 * the build/update guard, the trigger installer and the configuration
 * dialog and save.
 */
function assertNotTemplateMaster_() {
  if (isTemplateMasterDocument_()) {
    throw new Error('This is the SA4 Report Template itself. Use "Create New SA4 Report" -- ' +
      'reports are never built or configured in the template.');
  }
}

// =========================================================
// PERF-001 -- LIGHTWEIGHT PERFORMANCE INSTRUMENTATION
// =========================================================
//
// Manual-diagnostic only, purely additive: measures where continuousUpdate()
// actually spends time and how many times it repeats expensive DocumentApp/
// network operations, so a later controlled live run's Stackdriver/
// execution transcript can be read for real numbers. Every helper here
// either times/counts an existing call and returns ITS real, unchanged
// result (or rethrows its real error), or is a Logger.log-only summary --
// nothing here alters control flow, mutates the document differently, or
// changes what gets written/persisted. Counters/timings are per-execution
// (module-level state reset at the top of continuousUpdate(), never
// persisted to PropertiesService, never shown to the user).
// var (not const): Node's vm module (used by tests/helpers/load-code.js)
// only attaches top-level `var`/`function` declarations to the sandbox
// context object, not `const`/`let` -- this has no effect on real Apps
// Script behavior (which has no such distinction) and lets tests read
// PERF_TIMINGS_/PERF_COUNTERS_ directly to verify instrumentation state.
var PERF_TIMINGS_ = {};
var PERF_COUNTERS_ = {};

function perfResetState_() {
  Object.keys(PERF_TIMINGS_).forEach(k => delete PERF_TIMINGS_[k]);
  Object.keys(PERF_COUNTERS_).forEach(k => delete PERF_COUNTERS_[k]);
}

function perfAddTime_(label, ms) {
  PERF_TIMINGS_[label] = (PERF_TIMINGS_[label] || 0) + ms;
}

function perfCount_(label, n) {
  PERF_COUNTERS_[label] = (PERF_COUNTERS_[label] || 0) + (n === undefined ? 1 : n);
}

function perfCounterValue_(label) {
  return PERF_COUNTERS_[label] || 0;
}

/**
 * Times a single, one-shot stage: runs fn(), logs
 * "[PERF] <label>: <ms> ms" immediately (so a live execution transcript
 * shows progress stage-by-stage as it happens, not only at the end), and
 * also accumulates into PERF_TIMINGS_ for the final summary line. Use for
 * stages that run once per continuousUpdate() call. Always re-throws
 * fn()'s real error (via `finally`) -- timing never swallows a failure.
 */
function perfTimed_(label, fn) {
  const start = Date.now();
  try {
    return fn();
  } finally {
    const ms = Date.now() - start;
    perfAddTime_(label, ms);
    Logger.log('[PERF] ' + label + ': ' + ms + ' ms');
  }
}

/**
 * Times a stage that runs MANY times per continuousUpdate() call (e.g.
 * once per TDoc) -- accumulates into PERF_TIMINGS_ silently, WITHOUT a
 * Logger.log per call, so a report with dozens of TDocs doesn't produce
 * dozens of near-duplicate log lines. The accumulated total is included
 * in perfLogSummary_()'s single summary line at the end.
 */
function perfTimedAccum_(label, fn) {
  const start = Date.now();
  try {
    return fn();
  } finally {
    perfAddTime_(label, Date.now() - start);
  }
}

/**
 * Targeted counting wrapper around `body.getTables()` -- called
 * independently by many separate functions throughout this file (each
 * doing its OWN fresh full-document table scan: continuousUpdate() itself,
 * getReallocationMap_(), updateTdocStatus_(), findTdocTable_(),
 * updateRegisteredDocumentsTable_(), addAbstractsForTables_(),
 * removeRowHeightAndSpacing(), setTwoColumnTDocTableWidths_(),
 * checkRSSFeed_(), updateRevisions_()). Counting every call site here --
 * rather than monkey-patching DocumentApp globally -- directly answers
 * PERF-001's core question ("are we scanning the complete document once or
 * twenty times per update?") without touching the DocumentApp API surface
 * itself. Returns the exact same array `body.getTables()` would have.
 */
function getTablesCounted_(body, siteLabel) {
  perfCount_('body.getTables() calls (total)');
  if (siteLabel) perfCount_('body.getTables() call site: ' + siteLabel);
  return body.getTables();
}

/**
 * Targeted counting wrapper around `DocumentApp.getActiveDocument().getBody()`
 * -- several functions each independently re-fetch the live document body
 * rather than reusing continuousUpdate()'s own `body` reference. Returns
 * the exact same Body `getBody()` would have.
 *
 * ADDON-004: routed through getReportBody_(context) instead of calling
 * DocumentApp.getActiveDocument() directly, so every call site below stays
 * background-execution-safe once it threads a context through. With no
 * context (every pre-ADDON-004 call site, unchanged) getReportBody_()
 * falls back to DocumentApp.getActiveDocument().getBody() -- byte-identical
 * to this function's previous implementation.
 */
function getActiveDocumentBodyCounted_(siteLabel, context) {
  perfCount_('DocumentApp.getActiveDocument().getBody() calls (total)');
  if (siteLabel) perfCount_('getActiveDocument().getBody() call site: ' + siteLabel);
  return getReportBody_(context);
}

/**
 * Logs one single, machine-searchable summary line each for accumulated
 * timings and call counters -- call once, at the very end of
 * continuousUpdate() (in a `finally`, so it still runs after an error).
 */
function perfLogSummary_() {
  Logger.log('[PERF] timings summary (ms): ' + JSON.stringify(PERF_TIMINGS_));
  Logger.log('[PERF] counters summary: ' + JSON.stringify(PERF_COUNTERS_));
}

/**
 * CONTINUOUS UPDATE FUNCTION
 * Downloads latest TDOCs, adds new ones, updates status
 *
 * PERF-003B (Part 2): this is the public entry point -- the one the time-
 * driven trigger (ScriptApp.newTrigger('continuousUpdate'), see
 * createContinuousTrigger()) and the "Continuous Update" menu item both
 * invoke by name. It acquires a short, non-blocking, document-scoped lock
 * BEFORE any work at all (a losing run does nothing, not partial work),
 * then delegates to continuousUpdateCore_() for the actual logic, and
 * always releases the lock in `finally`. This guards against the exact
 * overlap risk flagged (but not fixed) in POST-MEETING-001 Task C: the
 * trigger firing again while a previous run -- trigger- or menu-invoked --
 * is still executing.
 */
function continuousUpdate() {
  // ADDON-004: legacy/interactive entry point -- UNCHANGED DocumentLock
  // semantics, exactly as before. This is the ONLY path a bound-script
  // document (v2.12.0-style) or an interactive add-on menu click ever
  // takes; it is never called by the background scheduler (see
  // continuousUpdateForDocument_(), which calls continuousUpdateCore_()
  // directly and owns its OWN concurrency guard via the scheduler's
  // ScriptLock instead -- see ADDON-004 report §5/§8 for why DocumentLock
  // is deliberately never acquired there).
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    Logger.log('continuousUpdate(): another incremental update is already in progress -- skipping this run.');
    return;
  }
  try {
    continuousUpdateCore_();
  } finally {
    lock.releaseLock();
  }
}

/**
 * The actual continuousUpdate() logic. Interactive callers (the bound
 * script's continuousUpdate(), the menu item, an interactive add-on
 * action) never call this directly -- always go through continuousUpdate(),
 * which holds the document lock for the duration of this call. The
 * background scheduler path (continuousUpdateForDocument_(), ADDON-004) is
 * the one exception: it calls this directly, with an explicit
 * addon-background context, and relies on the SCHEDULER's ScriptLock for
 * concurrency protection instead (see ADDON-004 report). (Kept as a
 * separate function, rather than inlined, purely so the lock-acquisition
 * wrapper above stays a small, easily-audited block instead of
 * interleaving lock logic into this function's existing try/catch/finally.)
 *
 * ADDON-004: accepts an optional execution context, threaded into every
 * background-reachable call below. With no context (continuousUpdate()'s
 * only call site, unchanged), every one of those calls resolves to EXACTLY
 * what it did before this parameter existed -- DocumentApp.getActiveDocument()
 * for document/body access, PropertiesService.getDocumentProperties() for
 * state. See getReportDocument_()/getReportBody_()/getReportStateStore_()
 * (ADDON-002) for why that fallback is guaranteed, not just intended.
 */
function continuousUpdateCore_(context) {
  perfResetState_();
  const perfTotalStart = Date.now();
  const cfg = perfTimed_('configuration/context (getReportConfig_)', () => getReportConfig_(context));
  const body = getReportBody_(context);
  perfCount_('DocumentApp.getActiveDocument().getBody() calls (total)');
  perfCount_('getActiveDocument().getBody() call site: continuousUpdate:initial');

  Logger.log('=== CONTINUOUS UPDATE START ===');

  // ADDON-004 (live-test follow-up): this result is new -- previously
  // this function returned nothing (implicit undefined) on EITHER path,
  // which meant continuousUpdateForDocument_() had no way to tell a real
  // completion apart from a failure this catch below swallowed. The
  // catch's own swallow-and-log behavior is UNCHANGED (still never
  // rethrows -- continuousUpdate(), the legacy/interactive caller, relies
  // on that and already ignores this function's return value entirely,
  // so this is purely additive for it). Only the NEW caller,
  // continuousUpdateForDocument_(), reads this.
  let result = { success: true, error: null };

  try {
    // ADDON-007B3: fail before any document change (a background run then
    // counts as failed and does not advance lastRunAt).
    assertMeetingReadyToBuild_(context);

    // Download latest TDOC list
    const tdocGroups = perfTimed_('TDoc-list fetch/download+parse (downloadAndGroupTdocs_)', () => downloadAndGroupTdocs_(cfg, context));
    const allTdocs = [];
    Object.keys(tdocGroups).forEach(key => {
      tdocGroups[key].tdocs.forEach(td => allTdocs.push({ ...td, agendaItem: key }));
    });
    perfCount_('TDocs downloaded (this run)', allTdocs.length);

    Logger.log(`Downloaded ${allTdocs.length} TDOCs`);

    // Get existing TDOCs. PERF-003 (Part A): also builds the canonical
    // TDoc-table index from this SAME scan (no extra body.getTables()
    // call) -- this is what lets updateTdocStatus_()/findTdocTable_() stop
    // rescanning the whole document once per TDoc/revision pair below.
    const existingTdocs = new Set();
    let tdocTableIndex;
    perfTimed_('existing-report/document scan (existing TDoc detection)', () => {
      const tables = getTablesCounted_(body, 'continuousUpdate:existingTdocScan');
      tables.forEach(t => {
        if (!isTDocTable_(t)) return;
        const tdoc = safeCellText_(t, 0, 1).trim();
        if (tdoc) existingTdocs.add(tdoc);
      });
      tdocTableIndex = buildTdocTableIndex_(tables);
    });

    Logger.log(`Found ${existingTdocs.size} existing TDOCs`);

    // Add new TDOCs and update status
    let newTdocsAdded = 0;
    let statusUpdated = 0;
    let linksAdded = 0;
    let uploadAbstractAttempts = 0;
    // TDocs whose abstract was asked for in this run on account of their
    // upload; the abstract sweep below does not ask for them again.
    const abstractAskedThisRun = {};

    allTdocs.forEach(tdocData => {
      const row = tdocData.row;
      const tdocNumber = String(row[tdocData.tdocCol] || '').trim();
      perfCount_('TDocs processed (loop iterations)');

      if (!existingTdocs.has(tdocNumber)) {
        perfTimedAccum_('new-TDoc insertion (insertNewTdoc_, accumulated)', () => insertNewTdoc_(body, tdocData, cfg, tdocTableIndex, context));
        // TEMPLATE-002A (port of Legacy BUGFIX-LEGACY-002, d1021cf):
        // `existingTdocs` was only populated from the document scan taken
        // BEFORE this loop, so a TDoc number appearing twice within this
        // same run's downloaded groups (e.g. under two agenda items) was
        // inserted twice. Recording it at insertion makes the repeat an
        // in-place update instead.
        existingTdocs.add(tdocNumber);
        newTdocsAdded++;
        // A TDoc that is uploaded when it is inserted has its link from the
        // insertion; its abstract is asked for once, here.
        if (perfTimedAccum_('upload completion (completeInsertedUploadedTdoc_, accumulated)', () => completeInsertedUploadedTdoc_(body, tdocNumber, tdocData, tdocTableIndex, context))) {
          uploadAbstractAttempts++;
          abstractAskedThisRun[tdocNumber] = true;
        }
      } else {
        if (perfTimedAccum_('status updates (updateTdocStatus_, accumulated)', () => updateTdocStatus_(body, tdocNumber, tdocData, tdocTableIndex))) {
          statusUpdated++;
        }
        // A TDoc that was inserted before it was uploaded is completed once
        // the list shows the upload: its link and one attempt at its
        // abstract. Independent of the status and of the abstracts setting.
        const completed = perfTimedAccum_('upload completion (refreshUploadedTdocMetadata_, accumulated)', () => refreshUploadedTdocMetadata_(body, tdocNumber, tdocData, tdocTableIndex, context));
        if (completed.linkAdded) linksAdded++;
        if (completed.abstractAttempted) {
          uploadAbstractAttempts++;
          abstractAskedThisRun[tdocNumber] = true;
        }
      }
    });

    Logger.log(`Added ${newTdocsAdded} new, updated ${statusUpdated} statuses`);
    const registrationLinksAdded = refreshRegistrationTableLinks_(body, allTdocs);
    Logger.log(`Upload completion: ${linksAdded} TDoc link(s) added, ${uploadAbstractAttempts} abstract attempt(s), ${registrationLinksAdded} registration-table link(s) added`);

    // Update summary table
    // Ad-hoc sessions (stage B): null unless this is an ad-hoc report with
    // sessions. With it, a report that HAS the Session column gets the
    // column brought up to date; a four-column report stays four-column.
    const tdocSessions = makeAdhocTdocSessionResolverSafely_(context);
    if (newTdocsAdded > 0) {
      perfTimed_('registered-documents summary (updateRegisteredDocumentsTable_)', () => updateRegisteredDocumentsTable_(body, allTdocs, tdocGroups, tdocSessions));
    }
    const sessionColumnFailure = refreshRegistrationSessionColumnSafely_(body, allTdocs, tdocSessions);

    // Revision placement: new TDOCs are appended at the end of their agenda
    // section, so move every revision back under the document it revises and
    // fill the revised document's Disposition. Reuses the TDOC list already
    // downloaded above, so this costs no extra fetch.
    const rev = perfTimed_('revision ordering/rearrangement (rearrangeRevisionTables_)', () => rearrangeRevisionTables_(cfg, tdocGroups, tdocTableIndex, context));
    Logger.log(`Revisions: ${rev.moved} moved, ${rev.dispositions} disposition(s) filled`);

    // Abstracts are optional on automatic updates (OFF by default).
    // Toggle via: REPORT OPERATIONS -> Manage Auto-Update Trigger.
    if (getFetchAbstractsSetting_(context)) {
      // PERF-006B: addAbstractsForTables_() now returns a breakdown, not a
      // bare count -- "candidate tables processed" is NOT the same as
      // "Reviewer requests actually made" (a candidate can resolve via a
      // negative-cache skip with zero fetches). See its own header comment.
      // A TDoc the list shows as not uploaded yet is left out: the Reviewer
      // has nothing for it, and asking would block it for a day (its "no
      // summary" answer is cached) just when the upload arrives. So is a
      // TDoc that was asked for above, in this run.
      const abstractsResult = perfTimed_('abstracts (addAbstractsForTables_)', () => addAbstractsForTables_(body, context, Object.assign(tdocsNotUploadedYet_(allTdocs), abstractAskedThisRun)));
      Logger.log(`Abstracts: ${abstractsResult.candidatesProcessed} candidate table(s) processed, ${abstractsResult.requestsMade} Reviewer request(s) made, ${abstractsResult.cacheSkips} negative-cache skip(s), ${abstractsResult.rowsInserted} abstract row(s) inserted`);
    } else {
      Logger.log('Abstract fetching is disabled (trigger configuration)');
    }

    // Collect emails and revisions
    // 2.17.4: the collector's own failures (logged there, never thrown) are
    // passed on in the result. `success` is unchanged by them.
    const collected = perfTimed_('collectorUpdate_ (RSS/mail + revisions, total)', () => collectorUpdate_(context));
    // Stage B: a Session column that could not be updated is reported the same way.
    const failures = (sessionColumnFailure ? [sessionColumnFailure] : []).concat((collected && collected.failures) || []);
    if (failures.length) result.collectorFailures = failures;

    // PERF-003 (Part B) / PERF-003B (Part 3): the unconditional,
    // full-document formatting pass is only actually needed when this run
    // inserted or moved a TABLE, or inserted a ROW into an existing table
    // -- never merely because cell TEXT changed (status updates, e-mail
    // discussion, revision cell rendering all rewrite existing cells in
    // place, never add/remove rows or tables). See
    // removeRowHeightAndSpacing()'s own header comment for the full
    // characterization, and shouldReformatAfterUpdate_() for the decision
    // itself (including why the 4th, abstract-row-insertion signal was
    // added). Every OTHER call site of removeRowHeightAndSpacing() (full
    // report build, "Update All", manual formatting menu items) is
    // untouched and still calls it unconditionally. All four counters are
    // read AFTER both addAbstractsForTables_() and collectorUpdate_() (the
    // only two call chains below this point that can still insert/move a
    // table) have already run, so nothing structural that happens later
    // in this same function is missed.
    const structuralChangeThisRun = shouldReformatAfterUpdate_(
      newTdocsAdded,
      rev.moved,
      perfCounterValue_('structural: new revision-linked tables inserted (insertRevisedDocTablesAfter_)'),
      perfCounterValue_('structural: abstract row inserted (fetchAndAddAbstract_)')
    );

    // Format
    if (structuralChangeThisRun) {
      perfTimed_('formatting (removeRowHeightAndSpacing)', () => removeRowHeightAndSpacing(context));
    } else {
      Logger.log('[PERF] formatting skipped: no structural change this run (no new/moved/inserted tables)');
    }

    Logger.log('=== COMPLETE ===');

  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    Logger.log(e.stack);
    result = { success: false, error: e.message };
  } finally {
    // PERF-002: this `finally` guarantees the summary runs after a normal
    // completion OR a catchable in-script exception (the `catch` above) --
    // it does NOT run after Apps Script's own hard execution-time-limit
    // termination, which kills the whole process from OUTSIDE the running
    // JavaScript and gives no guarantee that ANY further code (not even a
    // `finally` block, and not even an already-queued Logger.log call)
    // executes or flushes. That is exactly why every major stage above
    // also logs its own immediate "[PERF] <label>: <ms> ms" line via
    // perfTimed_() the moment it finishes, rather than relying solely on
    // this end-of-run summary -- for a genuine timeout, those per-stage
    // lines (as far as they got) are the only available evidence.
    perfAddTime_('TOTAL continuousUpdate', Date.now() - perfTotalStart);
    perfLogSummary_();
  }

  return result;
}

// =========================================================
// ADDON-004 -- BACKGROUND EXECUTION (per-document unit of work)
// =========================================================
//
// continuousUpdateForDocument_() is an INTERNAL unit of work, not the
// clock-trigger handler itself (that is runAddonScheduler_(), below,
// which is still not wired to any real trigger in this stage -- ADDON-004
// implements the handler only, per the task's explicit scope).
//
// It never acquires LockService.getDocumentLock() -- concurrency
// protection for the background path is the SCHEDULER's ScriptLock
// (runAddonScheduler_()), acquired once for the whole scheduler run, not
// per document (see that function's header comment for why). The
// legacy/interactive continuousUpdate() keeps its own, separate
// DocumentLock exactly as before -- the two paths never share a lock.
//
// Validates against the registry (getRegisteredReportDocument_()) BEFORE
// doing anything else -- a pure Script-Properties read, no DocumentApp/
// Document-Properties access at all -- so an invalid documentId or a
// disabled/unadopted document fails fast, loudly (throws), before any
// document is opened or any state is touched.
function continuousUpdateForDocument_(documentId) {
  if (!documentId) {
    throw new Error('continuousUpdateForDocument_: documentId is required.');
  }

  const registryEntry = getRegisteredReportDocument_(documentId);
  if (!registryEntry) {
    throw new Error('continuousUpdateForDocument_: document ' + documentId + ' is not registered (adopt it first).');
  }
  if (!registryEntry.enabled) {
    throw new Error('continuousUpdateForDocument_: document ' + documentId + ' is registered but not enabled.');
  }

  const context = { documentId: documentId, mode: 'addon-background' };
  const startedAt = Date.now();

  // continuousUpdateCore_() can fail two distinguishable ways, and only
  // one of them should advance lastRunAt:
  //
  //   (a) a SETUP-PHASE failure (invalid/unreadable central state, or the
  //       document itself failing to open via getReportBody_()) -- these
  //       happen BEFORE continuousUpdateCore_()'s own try block and so
  //       genuinely throw/propagate out of it, straight past this call;
  //   (b) an INTERNAL per-stage failure (e.g. the TDoc-list fetch itself
  //       failing) -- continuousUpdateCore_() catches and logs these
  //       itself (unchanged legacy behavior, relied on by continuousUpdate()),
  //       so they never throw here -- they surface only in the {success,
  //       error} result it now returns.
  //
  // Both must be treated as a failed run: lastRunAt is advanced ONLY when
  // continuousUpdateCore_() both returns AND reports success. A (b)-class
  // failure is converted into a thrown error here specifically so the
  // scheduler's per-document try/catch (runAddonScheduler_()) counts it
  // as failed, not succeeded -- exactly like an (a)-class failure already
  // did by simply propagating.
  const coreResult = continuousUpdateCore_(context);

  if (!coreResult || !coreResult.success) {
    throw new Error(
      'continuousUpdateForDocument_: continuousUpdateCore_() did not complete successfully for ' +
      documentId + (coreResult && coreResult.error ? (': ' + coreResult.error) : '.')
    );
  }

  const updated = updateRegisteredReportDocument_(documentId, { lastRunAt: new Date().toISOString() });

  return {
    documentId: documentId,
    success: true,
    durationMs: Date.now() - startedAt,
    lastRunAt: updated.lastRunAt
  };
}

// =========================================================
// ADDON-004 -- CENTRAL SCHEDULER (handler only -- no trigger created)
// =========================================================
//
// Due calculation: never-run (lastRunAt === null) is always due; otherwise
// due once intervalHours has elapsed since lastRunAt. Pure/testable in
// isolation from the scheduler loop itself.
function isReportDocumentDue_(entry, nowMs) {
  if (!entry || !entry.enabled) return false;
  if (!entry.lastRunAt) return true;

  const intervalMs = (entry.intervalHours || 1) * 3600 * 1000;
  const lastRunMs = new Date(entry.lastRunAt).getTime();
  // A malformed/unparseable lastRunAt fails safe as "due" -- never
  // permanently wedges a document out of consideration because of a
  // corrupted timestamp.
  if (isNaN(lastRunMs)) return true;

  return (nowMs - lastRunMs) >= intervalMs;
}

/**
 * Deterministic fair ordering for a batch of due documents: never-run
 * documents first (they have been waiting the longest, definitionally),
 * then oldest lastRunAt first, with a stable documentId tie-break so two
 * documents with identical timestamps always order the same way run to
 * run. Pure -- does not mutate its input.
 */
function orderDueReportDocuments_(entries) {
  return (entries || []).slice().sort(function (a, b) {
    const aNever = !a.lastRunAt;
    const bNever = !b.lastRunAt;
    if (aNever !== bNever) return aNever ? -1 : 1;

    if (!aNever) {
      const aTime = new Date(a.lastRunAt).getTime();
      const bTime = new Date(b.lastRunAt).getTime();
      const aValid = !isNaN(aTime);
      const bValid = !isNaN(bTime);
      if (aValid !== bValid) return aValid ? 1 : -1; // an unparseable timestamp sorts as if never-run (first)
      if (aValid && bValid && aTime !== bTime) return aTime - bTime;
    }

    return a.documentId < b.documentId ? -1 : (a.documentId > b.documentId ? 1 : 0);
  });
}

// Conservative runtime budget for one scheduler execution, well under
// Apps Script's documented per-execution ceiling (commonly cited as 6
// minutes for a triggered function on a standard account -- see the
// ADDON-004 report for the exact platform reference and its
// requires-live-verification caveat). Deliberately conservative: this
// budget must leave room for the CURRENTLY-RUNNING document's own update
// to finish (it is checked BEFORE starting the next document, not used to
// interrupt one already in progress) plus genuine headroom for network
// latency variance across TDoc-list/RSS/A1/revision fetches.
var REPORT_SCHEDULER_RUNTIME_BUDGET_MS_ = 4 * 60 * 1000; // 4 minutes

/**
 * The scheduler HANDLER only -- not wired to any real time-driven trigger
 * in this stage (ADDON-004 explicitly stops short of creating one). Owns
 * the ScriptLock for the whole run (not per document -- see
 * continuousUpdateForDocument_()'s header comment), lists/filters/orders
 * due documents, and processes them one at a time, isolating each
 * document's failure from the others and stopping (deferring, not
 * failing) once the runtime budget is exhausted.
 */
function runAddonScheduler_() {
  const startedAt = Date.now();
  const summary = {
    considered: 0,
    due: 0,
    succeeded: 0,
    failed: 0,
    skipped: 0,
    lockAcquired: false,
    documents: []
  };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('runAddonScheduler_(): could not acquire ScriptLock (another scheduler run or interactive update holds it) -- skipping this run.');
    return summary;
  }
  summary.lockAcquired = true;

  try {
    const registered = listRegisteredReportDocuments_();
    summary.considered = registered.length;

    const now = Date.now();
    const dueEntries = registered.filter(function (entry) { return isReportDocumentDue_(entry, now); });
    summary.due = dueEntries.length;

    const ordered = orderDueReportDocuments_(dueEntries);

    ordered.forEach(function (entry) {
      if ((Date.now() - startedAt) >= REPORT_SCHEDULER_RUNTIME_BUDGET_MS_) {
        summary.skipped++;
        summary.documents.push({ documentId: entry.documentId, outcome: 'deferred', reason: 'runtime budget exhausted' });
        return;
      }

      try {
        const result = continuousUpdateForDocument_(entry.documentId);
        summary.succeeded++;
        summary.documents.push({ documentId: entry.documentId, outcome: 'succeeded', durationMs: result.durationMs });
      } catch (e) {
        summary.failed++;
        summary.documents.push({ documentId: entry.documentId, outcome: 'failed', error: e.message });
        Logger.log('runAddonScheduler_(): ' + entry.documentId + ' failed: ' + e.message);
      }
    });

    Logger.log('runAddonScheduler_() summary: ' + JSON.stringify({
      considered: summary.considered, due: summary.due, succeeded: summary.succeeded,
      failed: summary.failed, skipped: summary.skipped
    }));

    return summary;
  } finally {
    lock.releaseLock();
  }
}

// =========================================================
// ADDON-004 -- CENTRAL STATE CAPACITY DIAGNOSTICS
// =========================================================
//
// Read-only monitoring helper: approximates how much of the shared 500KB
// Script Properties budget (ADDON-002's capacity analysis) is actual
// report state (SA4_STATE|<documentId>|<key>, Backend B), broken down per
// document, so old meetings can be identified for archiving
// (deleteCentralReportState_(), ADDON-003) before that shared ceiling
// becomes a real constraint. Never touches the registry
// (SA4_REGISTRY_*) or global keys (REVIEWER_API_TOKEN) -- only counts
// keys under the report-state namespace.
//
// Byte counts are an intentional APPROXIMATION (JavaScript string
// .length -- UTF-16 code units, not a true UTF-8 byte count) rather than
// a Utilities.newBlob() round trip per key -- adequate for a monitoring
// signal over what is, in practice, ASCII-heavy JSON, and keeps this
// diagnostic free of any extra Apps Script service dependency (so it also
// runs unmodified in the Node test sandbox).
function getCentralStateUsage_() {
  const scriptProps = PropertiesService.getScriptProperties();
  const prefix = 'SA4_STATE|';
  const perDocument = {};
  let totalBytesApprox = 0;
  let keyCount = 0;

  scriptProps.getKeys().forEach(function (key) {
    if (key.indexOf(prefix) !== 0) return;
    const rest = key.slice(prefix.length);
    const sep = rest.indexOf('|');
    if (sep === -1) return; // malformed/unexpected key shape -- ignore, don't throw

    const documentId = rest.slice(0, sep);
    const value = scriptProps.getProperty(key) || '';
    const bytesApprox = key.length + value.length;

    if (!perDocument[documentId]) {
      perDocument[documentId] = { documentId: documentId, keyCount: 0, bytesApprox: 0 };
    }
    perDocument[documentId].keyCount++;
    perDocument[documentId].bytesApprox += bytesApprox;

    totalBytesApprox += bytesApprox;
    keyCount++;
  });

  return {
    totalBytesApprox: totalBytesApprox,
    keyCount: keyCount,
    documentCount: Object.keys(perDocument).length,
    perDocument: Object.keys(perDocument)
      .map(function (id) { return perDocument[id]; })
      .sort(function (a, b) { return b.bytesApprox - a.bytesApprox; })
  };
}

// =========================================================
// ADDON-005 -- CENTRAL SCHEDULER TRIGGER LIFECYCLE
// =========================================================
//
// Exactly ONE hourly clock trigger for the whole central project, shared
// by every adopted document (ADDON-004's runAddonScheduler_() already
// loops the registry itself -- the trigger only needs to exist once, not
// once per document). Never scans or touches any trigger whose handler
// function isn't REPORT_SCHEDULER_TRIGGER_HANDLER_ -- deliberately not
// the legacy deleteContinuousTrigger() pattern of iterating every project
// trigger, which would be unsafe to reuse in a shared central project
// that could, in principle, host other unrelated triggers later.
//
// ADDON-005 Phase D: verified against community reports (not primary
// Google documentation, which does not explicitly confirm either way --
// see the ADDON-005 report) that an underscore-suffixed function name is
// NOT reliably documented as a valid installable-trigger handler, mirroring
// the exact RESOLVER-HOTFIX problem this codebase already hit once for
// google.script.run (a trailing "_" made a function silently uncallable
// as an RPC target). Rather than assume either way, the ACTUAL trigger
// handler is the public runAddonSchedulerTrigger() wrapper below, not
// runAddonScheduler_() itself -- same defensive pattern already
// established in this file, zero behavior risk either way.
var REPORT_SCHEDULER_TRIGGER_UID_KEY_ = 'SA4_SCHEDULER_TRIGGER_UID';
var REPORT_SCHEDULER_TRIGGER_HANDLER_ = 'runAddonSchedulerTrigger';

/**
 * The ACTUAL clock-trigger handler function name registered with
 * ScriptApp.newTrigger(). Public/non-underscore by construction (see
 * header comment above) -- a thin, otherwise-behavior-free delegation to
 * the tested, documented runAddonScheduler_().
 */
function runAddonSchedulerTrigger() {
  return runAddonScheduler_();
}

function findAddonSchedulerTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === REPORT_SCHEDULER_TRIGGER_HANDLER_;
  });
}

/**
 * Read-only status inspection -- never creates, deletes, or modifies
 * anything.
 */
function getAddonSchedulerTriggerStatus_() {
  const triggers = findAddonSchedulerTriggers_();
  return {
    exists: triggers.length > 0,
    count: triggers.length,
    triggerUids: triggers.map(function (t) { return t.getUniqueId(); }),
    storedUid: PropertiesService.getScriptProperties().getProperty(REPORT_SCHEDULER_TRIGGER_UID_KEY_)
  };
}

/**
 * Idempotent: creates the one hourly scheduler trigger if and only if
 * none with this handler already exists. Self-heals the stored trigger
 * UID (SA4_SCHEDULER_TRIGGER_UID) against whatever ACTUALLY exists in
 * ScriptApp.getProjectTriggers() -- the stored UID is a convenience/
 * diagnostic value, never the sole source of truth for "does a trigger
 * exist" (that question always goes through findAddonSchedulerTriggers_(),
 * a live platform read). Defensively de-duplicates if more than one
 * matching trigger is somehow found (should never happen given this
 * function is the only trigger-creation path, but "duplicate prevention"
 * is a hard requirement, not just a create-time check).
 */
function ensureAddonSchedulerTrigger_() {
  const existing = findAddonSchedulerTriggers_();
  const scriptProps = PropertiesService.getScriptProperties();

  if (existing.length > 1) {
    for (let i = 1; i < existing.length; i++) ScriptApp.deleteTrigger(existing[i]);
  }

  if (existing.length > 0) {
    const uid = existing[0].getUniqueId();
    scriptProps.setProperty(REPORT_SCHEDULER_TRIGGER_UID_KEY_, uid);
    return { created: false, triggerUid: uid, duplicatesRemoved: Math.max(0, existing.length - 1) };
  }

  const trigger = ScriptApp.newTrigger(REPORT_SCHEDULER_TRIGGER_HANDLER_)
    .timeBased()
    .everyHours(1)
    .create();
  const uid = trigger.getUniqueId();
  scriptProps.setProperty(REPORT_SCHEDULER_TRIGGER_UID_KEY_, uid);
  return { created: true, triggerUid: uid, duplicatesRemoved: 0 };
}

/**
 * Safe deletion: removes every trigger matching ONLY this handler name,
 * and the stored UID property. Never touches SA4_REGISTRY_ or SA4_STATE
 * keys -- registered documents and their central state survive the
 * scheduler trigger being deleted (they simply stop being serviced until
 * ensureAddonSchedulerTrigger_() is called again).
 */
function deleteAddonSchedulerTrigger_() {
  const existing = findAddonSchedulerTriggers_();
  existing.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  PropertiesService.getScriptProperties().deleteProperty(REPORT_SCHEDULER_TRIGGER_UID_KEY_);
  return { deleted: existing.length };
}

// =========================================================
// ADDON-005 -- INTERACTIVE/BACKGROUND LOCKING FOR ADOPTED DOCUMENTS
// =========================================================
//
// Resolves the ADDON-004 open question: an interactive add-on action that
// mutates an ADOPTED document's central state/registry could otherwise
// race the background scheduler touching the SAME state. Decision (this
// is a single-user tool -- Thomas is the only interactive actor, so
// script-wide serialization has no real downside): the three NEW
// interactive mutation entry points below (Phase C) acquire the SAME
// LockService.getScriptLock() the scheduler (runAddonScheduler_()) uses,
// with the same short, non-blocking tryLock() pattern the legacy
// DocumentLock path already established -- not a new locking philosophy,
// just the ScriptLock equivalent of it. Deliberately NOT applied to every
// menu item: the pre-existing legacy report-building functions
// (runFullReportBuild, buildSkeletonWithTdocTables, etc.) are unchanged
// by ADDON-005 and are a separate, already-characterized, NOT-yet-fixed
// residual gap (see the ADDON-005 report) -- retrofitting locks onto them
// is out of scope here, the same way ADDON-004 deferred this exact
// decision to ADDON-005 rather than over-engineering it early.
function withAddonScriptLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw new Error('Another central add-on operation (interactive or scheduled) is already in progress -- try again shortly.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// =========================================================
// ADDON-005 -- ADD-ON AUTOMATIC-UPDATE UI (separate from the legacy
// 15/30/60-minute bound-script trigger UI, which is untouched)
// =========================================================
//
// Every function below is menu-callable (public, zero-arg) and explicitly
// constructs { documentId, mode: 'addon-interactive' } for the active
// document -- this is the "make execution mode explicit" requirement:
// these are new, add-on-specific entry points, not a reinterpretation of
// any existing zero-context call (every pre-existing menu item is
// untouched and still resolves to legacy Document Properties exactly as
// before).
function enableAutomaticUpdatesForAddon() {
  const ui = DocumentApp.getUi();
  const documentId = DocumentApp.getActiveDocument().getId();
  const context = { documentId: documentId, mode: 'addon-interactive' };

  withAddonScriptLock_(function () {
    // ADDON-005 (live-test follow-up): always (re-)adopt, not just when
    // never-before-registered. adoptReportDocumentForAddon_() is
    // idempotent (ADDON-003) -- re-running it on an already-adopted
    // document simply re-copies the CURRENT Document Properties values
    // into central state. Without this, "Enable Automatic Updates" on an
    // already-adopted document silently kept the STALE central-state
    // snapshot from first adoption, with no way to pick up a config
    // change (e.g. TDOC_LIST_URL) saved afterward via "Configure Meeting
    // Settings" (which still writes Document Properties, unchanged --
    // see ADDON-004 report Section 5) -- discovered during the live
    // scratch-document acceptance test itself. "Enable" is the one
    // button that already both reads AND writes registration state, so
    // it's the natural place to also (re-)sync, without adding a
    // separate new menu item for it.
    const adoption = adoptReportDocumentForAddon_(context);
    if (!adoption.verified) {
      ui.alert('Could Not Enable Automatic Updates',
        'Adopting this document failed to verify: ' + adoption.mismatches.join(', ') + '.\n' +
        'No automatic updates were enabled.',
        ui.ButtonSet.OK);
      return;
    }

    const registryEntry = registerReportDocument_(documentId, { enabled: true });
    const triggerResult = ensureAddonSchedulerTrigger_();

    ui.alert('Automatic Updates Enabled',
      'This document will be updated automatically every ' + registryEntry.intervalHours + ' hour(s) ' +
      'by the central scheduler.\n\n' +
      'Central scheduler trigger: ' + (triggerResult.created ? 'created just now' : 'already running') + '.',
      ui.ButtonSet.OK);
  });
}

function disableAutomaticUpdatesForAddon() {
  const ui = DocumentApp.getUi();
  const documentId = DocumentApp.getActiveDocument().getId();

  withAddonScriptLock_(function () {
    const registryEntry = getRegisteredReportDocument_(documentId);
    if (!registryEntry) {
      ui.alert('Not Enabled', 'This document has no automatic-update registration to disable.', ui.ButtonSet.OK);
      return;
    }

    // enabled = false only -- report state and the registry entry itself
    // are deliberately preserved (see Phase C requirement: "do not purge
    // central state automatically"). The scheduler's own due-calculation
    // (isReportDocumentDue_()) already treats a disabled entry as never
    // due, so this alone is sufficient to stop future runs.
    registerReportDocument_(documentId, { enabled: false });
    ui.alert('Automatic Updates Disabled',
      'This document will no longer be updated automatically. Its saved state and settings were kept -- ' +
      're-enabling later resumes from where it left off.',
      ui.ButtonSet.OK);
  });
}

/**
 * Presents only the platform-supported interval choices (ADDON-001B: add-on
 * time-driven triggers cannot run more frequently than once per hour) --
 * deliberately excludes the legacy dialog's 15/30-minute options.
 */
function setAutomaticUpdateIntervalForAddon() {
  const ui = DocumentApp.getUi();
  const documentId = DocumentApp.getActiveDocument().getId();
  const registryEntry = getRegisteredReportDocument_(documentId);

  if (!registryEntry) {
    ui.alert('Not Enabled', 'Enable automatic updates for this document first (☁️ CENTRAL ADD-ON → Enable Automatic Updates).', ui.ButtonSet.OK);
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 10px; font-weight: bold; }
      select { width: 100%; padding: 8px; margin-top: 5px; }
      button { margin-top: 20px; padding: 10px 20px; border: none; cursor: pointer; color: white; background: #1a73e8; }
    </style>
    <h2>⏱️ Set Update Interval</h2>
    <label>Update every:</label>
    <select id="interval">
      ${REPORT_REGISTRY_ALLOWED_INTERVAL_HOURS_.map(h =>
        `<option value="${h}" ${h === registryEntry.intervalHours ? 'selected' : ''}>Every ${h} hour${h === 1 ? '' : 's'}</option>`
      ).join('')}
    </select>
    <div>
      <button onclick="save()">Save</button>
    </div>
    <script>
      function save() {
        const hours = parseInt(document.getElementById('interval').value, 10);
        google.script.run
          .withSuccessHandler(() => { alert('Saved.'); google.script.host.close(); })
          .withFailureHandler((err) => alert('Error: ' + err))
          .setAutomaticUpdateIntervalForAddonRpc(hours);
      }
    </script>
  `).setWidth(360).setHeight(260);

  ui.showModalDialog(html, 'Set Update Interval');
}

/**
 * The actual public google.script.run RPC target for the dialog above --
 * same RESOLVER-HOTFIX-informed naming discipline as the rest of this
 * file: the dialog never calls an underscore-suffixed function directly.
 */
function setAutomaticUpdateIntervalForAddonRpc(hours) {
  const documentId = DocumentApp.getActiveDocument().getId();
  return withAddonScriptLock_(function () {
    return registerReportDocument_(documentId, { intervalHours: hours });
  });
}

function showAddonSchedulerStatusForAddon() {
  const ui = DocumentApp.getUi();
  const documentId = DocumentApp.getActiveDocument().getId();
  const registryEntry = getRegisteredReportDocument_(documentId);
  const triggerStatus = getAddonSchedulerTriggerStatus_();

  const lines = [];
  if (!registryEntry) {
    lines.push('This document is not registered for automatic updates.');
  } else {
    lines.push('Enabled: ' + (registryEntry.enabled ? 'Yes' : 'No'));
    lines.push('Interval: every ' + registryEntry.intervalHours + ' hour(s)');
    lines.push('Last run: ' + (registryEntry.lastRunAt || 'never'));
    lines.push('Registered: ' + registryEntry.registeredAt);
  }
  lines.push('');
  lines.push('Central scheduler trigger: ' + (triggerStatus.exists ? 'running (' + triggerStatus.count + ')' : 'not running'));

  ui.alert('Add-on Status', lines.join('\n'), ui.ButtonSet.OK);
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
  // TEMPLATE-002B: the template runtime records the interval it installed.
  const storedInterval = continuousTrigger && templateRuntimeRelease_() ? describeStoredContinuousInterval_(props) : null;
  if (storedInterval) currentInterval = storedInterval.label;
  
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
      ${continuousTriggerIntervals_().map(function (o) {
        return '<option value="' + o.minutes + '"' + ((storedInterval ? o.minutes === storedInterval.minutes : o.selected) ? ' selected' : '') + '>' + o.label + '</option>';
      }).join('\n      ')}
    </select>
    
    <label>Abstracts:</label>
    <label class="check">
      <input type="checkbox" id="fetchAbstracts" ${fetchAbstracts ? 'checked' : ''} onchange="saveAbstracts()">
      Fetch abstracts during each update
    </label>
    <div class="hint">
      Off by default. When on, every automatic update calls the Reviewer API for
      each TDOC still missing an abstract, which makes updates noticeably slower.
      Abstracts can also be updated manually at any time.
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
 *
 * ADDON-004: accepts an optional context, routed through
 * getReportStateStore_() -- reachable from continuousUpdateCore_(), so a
 * background execution must read this from the adopted document's central
 * state, not always Document Properties. No context (unchanged) reads
 * Document Properties exactly as before.
 */
function getFetchAbstractsSetting_(context) {
  return getReportStateStore_(context)
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

/**
 * 2.15.1: the Continuous Update intervals the trigger dialog offers, each
 * with the TriggerBuilder call that implements it. The dialog's options are
 * rendered from this list.
 *
 * 2.15.2: CENTRAL runs as an Editor add-on, and add-on time-driven triggers
 * cannot run more often than once per hour (ADDON-001B; see
 * REPORT_REGISTRY_ALLOWED_INTERVAL_HOURS_) -- everyMinutes(15/30), valid in
 * a bound script, is refused live. Only the hourly interval is offered.
 */
var CONTINUOUS_TRIGGER_INTERVALS_ = [
  { minutes: 60, label: 'Every hour (Slow meeting)', everyHours: 1, selected: true }
];

// TEMPLATE-002B: Document Property in which the template runtime records
// the interval of the trigger it installed (Apps Script does not expose a
// trigger's interval).
var CONTINUOUS_UPDATE_INTERVAL_KEY_ = 'CONTINUOUS_UPDATE_INTERVAL_MINUTES';

/**
 * TEMPLATE-002B: the intervals this runtime offers. The bound template
 * runtime offers 15/30/60 minutes (its table and installer live in
 * ReportCreator.js, which only template bundles contain -- this file never
 * creates a sub-hourly trigger); the add-on offers hourly only.
 */
function continuousTriggerIntervals_() {
  return templateRuntimeRelease_() ? templateContinuousTriggerIntervals_() : CONTINUOUS_TRIGGER_INTERVALS_;
}

/** { minutes, label } for the recorded interval, or null when none is recorded. */
function describeStoredContinuousInterval_(props) {
  const minutes = parseInt(props.getProperty(CONTINUOUS_UPDATE_INTERVAL_KEY_) || '', 10);
  const known = continuousTriggerIntervals_().some(function (o) { return o.minutes === minutes; });
  return known ? { minutes: minutes, label: minutes + ' minutes' } : null;
}

/**
 * 2.15.1: the CONTINUOUS_TRIGGER_INTERVALS_ entry for a dialog interval (a
 * whole number of minutes, as a number or a digit string). Anything else --
 * missing, malformed, or not offered -- throws before ScriptApp is touched.
 */
function resolveContinuousTriggerInterval_(intervalMinutes) {
  const raw = typeof intervalMinutes === 'string' ? intervalMinutes.trim() : intervalMinutes;
  const minutes = (typeof raw === 'number' && Number.isInteger(raw)) ? raw
    : (typeof raw === 'string' && /^\d+$/.test(raw)) ? parseInt(raw, 10)
    : null;
  const entry = minutes === null ? null : continuousTriggerIntervals_().find(function (o) { return o.minutes === minutes; });
  if (!entry) {
    throw new Error('Unsupported update interval: ' + JSON.stringify(intervalMinutes === undefined ? null : intervalMinutes) +
      ' (supported: ' + continuousTriggerIntervals_().map(function (o) { return o.minutes; }).join(', ') + ' minutes).');
  }
  return entry;
}

function createContinuousTrigger(intervalMinutes, fetchAbstracts) {
  // 2.15.1: validate first -- an unsupported interval changes nothing (the
  // existing trigger and the abstracts switch are left as they are).
  const interval = resolveContinuousTriggerInterval_(intervalMinutes);

  // TEMPLATE-002B: the bound template runtime installs the replacement
  // before it removes the previous trigger (createTemplateContinuousTrigger_()
  // in ReportCreator.js).
  if (templateRuntimeRelease_()) return createTemplateContinuousTrigger_(interval, fetchAbstracts);

  // Delete existing trigger first
  deleteContinuousTrigger();

  // Persist the abstracts switch alongside the trigger
  if (fetchAbstracts !== undefined && fetchAbstracts !== null) {
    setFetchAbstractsSetting(fetchAbstracts);
  }

  // Create new trigger
  ScriptApp.newTrigger('continuousUpdate')
    .timeBased()
    .everyHours(interval.everyHours)
    .create();

  Logger.log(`Trigger created: everyHours(${interval.everyHours}) (abstracts: ${getFetchAbstractsSetting_()})`);
}

function deleteContinuousTrigger() {
  if (templateRuntimeRelease_()) PropertiesService.getDocumentProperties().deleteProperty(CONTINUOUS_UPDATE_INTERVAL_KEY_);
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
  
  // TEMPLATE-002B: the template runtime knows the interval it installed.
  const stored = templateRuntimeRelease_() ? describeStoredContinuousInterval_(PropertiesService.getDocumentProperties()) : null;
  return {
    active: true,
    interval: stored ? stored.label : 'Active' // Apps Script doesn't expose the exact interval
  };
}


/**
 * PERF-003 (Part A): `index`, if supplied, is threaded through to
 * findParentRevisedToTable_() (avoiding a fresh full-table scan there),
 * and is itself UPDATED in place with the newly-inserted table -- so a
 * LATER new TDoc in the same continuousUpdate() run that revises THIS one
 * can still find it via the index, without needing a full rebuild mid-run.
 */
function insertNewTdoc_(body, tdocData, cfg, index, context) {
  const row = tdocData.row;
  const agendaItem = tdocData.agendaItem;
  const revisedTo = getRevisedTo_(tdocData);

  // Prefer sitting directly below the document this one revises. Only if that
  // document is not in the report yet do we fall back to the end of the
  // agenda section (continuousUpdate's re-arrangement pass fixes it later).
  let insertIdx = -1;
  const parentTable = findParentRevisedToTable_(body, tdocNumberOf_(tdocData), index, context);
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

  const newTable = insertTDocTableAtIndex_(body, insertIdx, tempData, tdocData.richTextRow, tdocData.tdocCol);
  Logger.log(`Inserted ${row[tdocData.tdocCol]}`);

  if (index) {
    const parsed = parseExactSA4DocumentId_(String(row[tdocData.tdocCol] || '').trim());
    if (parsed.isValid) index.set(parsed.raw, newTable);
  }
}

/**
 * Find the table of the document that was revised INTO `tdocNumber`, using the
 * REVISION_MAP built from the TDOC list. Returns null when unknown or when that
 * document has no table in this report.
 */
function findParentRevisedToTable_(body, tdocNumber, index, context) {
  const want = String(tdocNumber || '').trim().toUpperCase();
  if (!want) return null;

  const map = loadJsonObject_(
    getReportStateStore_(context).getProperty('REVISION_MAP'));

  for (const parent of Object.keys(map)) {
    if (String(map[parent] || '').toUpperCase() !== want) continue;
    const t = findTdocTable_(body, parent, index);
    if (t) return t;
  }
  return null;
}

/**
 * PERF-003 (Part A): the actual "found the table -- decide/apply the
 * status update" logic, extracted so BOTH the fast (indexed) and fallback
 * (full-scan) paths in updateTdocStatus_() below share exactly one copy
 * of it -- byte-identical decision logic either way.
 */
function applyTdocStatusUpdate_(table, tdocNumber, newStatus) {
  const statusInfo = findStatusInDocTable_(table);
  if (!statusInfo) return false;

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

/**
 * PERF-003 (Part A): `index`, if supplied (a buildTdocTableIndex_()
 * result), is tried first for an O(1) lookup -- the measured cost this
 * optimization targets was exactly this function calling
 * body.getTables() once per EXISTING TDoc every run (32 times for a
 * 32-TDoc no-change run). The indexed candidate's raw TDoc cell text is
 * still verified to exactly equal `tdocNumber` before use (the SAME
 * case-sensitive, raw-string-equality check the original code always
 * used) -- if that verification ever fails, or no index was supplied,
 * this falls straight back to the original full scan, so behavior is
 * byte-identical in every case, not just the common one.
 */
function updateTdocStatus_(body, tdocNumber, tdocData, index) {
  const row = tdocData.row;
  const newStatus = tdocData.statusCol >= 0 ? String(row[tdocData.statusCol] || '').trim() : '';

  if (!newStatus) return false;

  if (index) {
    const candidate = lookupTdocTableInIndex_(index, tdocNumber);
    if (candidate && safeCellText_(candidate, 0, 1).trim() === tdocNumber) {
      return applyTdocStatusUpdate_(candidate, tdocNumber, newStatus);
    }
  }

  const tables = getTablesCounted_(body, 'updateTdocStatus_');
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (!isTDocTable_(table)) continue;

    const tableTdoc = safeCellText_(table, 0, 1).trim();
    if (tableTdoc !== tdocNumber) continue;

    return applyTdocStatusUpdate_(table, tdocNumber, newStatus);
  }
  return false;
}

function updateRegisteredDocumentsTable_(body, allTdocs, tdocGroups, sessionResolver) {
  const tables = getTablesCounted_(body, 'updateRegisteredDocumentsTable_');
  let summaryTable = null;
  // Ad-hoc sessions (stage B): the table is the four-column one of every
  // release, or the same with a fifth "Session" column. A row is appended
  // with as many cells as the table has -- a column is never added here.
  let columns = 0;

  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (table.getNumRows() < 1) continue;

    columns = registrationTableColumns_(table);
    if (columns) {
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
      if (columns === 5) dataRow.appendTableCell(sessionResolver ? sessionResolver.cell(tdocData) : '');

      added++;
    }
  });
  
  // New rows of a table with the Session column get the widths of its columns.
  if (columns === 5) applyRegistrationTableWidths_(summaryTable);

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
// ADDON-002 -- REPORT EXECUTION CONTEXT
// =========================================================
//
// Infrastructure only (per ADDON-001/ADDON-001B): a small, optional context
// shape so business logic can eventually run without an active document
// (a background/trigger execution), without changing how it runs TODAY.
// Nothing in this file constructs or passes a context yet -- every existing
// caller continues to call getReportDocument_()/getReportBody_()/
// getReportConfig_()/getMeetingContext_() with no arguments, which is
// guaranteed (see each function below) to behave EXACTLY as
// DocumentApp.getActiveDocument()/getBody() and the direct
// PropertiesService.getDocumentProperties() calls they replace.
//
// Shape (plain object, deliberately not a class -- nothing here needs
// inheritance, private fields, or identity beyond "a bag of optional
// values"):
//   {
//     documentId: string | null,   // required for addon-background mode
//     document:   Document | null, // an already-open Document, reused as-is
//     mode:       'bound' | 'addon-interactive' | 'addon-background'
//   }
//
// mode is informational/selects the state-store backend (see
// getReportStateStore_() below); getReportDocument_() itself only looks at
// document/documentId, not mode, because "does this execution have a
// Document to work with" and "where does its Properties state live" are
// two separate questions (ADDON-001B found DocumentApp.openById() does NOT
// establish a PropertiesService document context -- so a context can
// legitimately have a resolvable document while still needing the central,
// documentId-keyed Properties backend).
function getReportDocument_(context) {
  if (context && context.document) return context.document;
  if (context && context.documentId) {
    // Cache on the context object itself so a single execution that calls
    // getReportDocument_()/getReportBody_() many times (continuousUpdateCore_()
    // and friends do, today via getActiveDocumentBodyCounted_()) only pays
    // for one DocumentApp.openById() round trip, not one per call site.
    if (!context._resolvedDocument) {
      context._resolvedDocument = DocumentApp.openById(context.documentId);
    }
    return context._resolvedDocument;
  }
  return DocumentApp.getActiveDocument();
}

function getReportBody_(context) {
  return getReportDocument_(context).getBody();
}

function getReportDocumentId_(context) {
  if (context && context.documentId) return context.documentId;
  return getReportDocument_(context).getId();
}

// ---------------------------------------------------------------
// ADDON-002 -- REPORT STATE STORE (Document Properties abstraction)
// ---------------------------------------------------------------
//
// Two backends behind one Properties-compatible surface
// (getProperty/setProperty/deleteProperty/getKeys/getProperties/
// setProperties -- the exact subset of the real Properties API this file
// actually uses today, see the ADDON-002 report's property inventory).
//
// Backend A (default, unchanged): the real
// PropertiesService.getDocumentProperties() object, returned as-is -- not
// wrapped, not reimplemented. This is why "no context" is byte-for-byte
// identical to pre-ADDON-002 behavior: it IS the same call.
//
// Backend B (central/background, net-new, NOT wired into production
// state yet): one Script Property per ORIGINAL key, namespaced by
// documentId, not one combined JSON blob -- see the ADDON-002 report's
// capacity analysis for why (several existing per-document keys, e.g.
// PARSED_AGENDA_ALL, plus dozens-to-hundreds of per-TDoc DISCUSS_/REVIS_/
// cache keys, can together approach the 500KB per-store limit; a single
// JSON blob would additionally be capped at the 9KB per-value limit, which
// a meeting's full state can easily exceed even though no single existing
// key does).
function reportStateScriptPropertyPrefix_(documentId) {
  if (!documentId) {
    throw new Error('reportStateScriptPropertyPrefix_: documentId is required.');
  }
  return 'SA4_STATE|' + documentId + '|';
}

function makeCentralReportStateStore_(documentId) {
  const scriptProps = PropertiesService.getScriptProperties();
  const prefix = reportStateScriptPropertyPrefix_(documentId);

  return {
    getProperty: function (key) {
      return scriptProps.getProperty(prefix + key);
    },
    setProperty: function (key, value) {
      scriptProps.setProperty(prefix + key, value);
    },
    deleteProperty: function (key) {
      scriptProps.deleteProperty(prefix + key);
    },
    // Real Properties.getKeys() returns every key in the store; this
    // backend narrows that to just this document's namespace and strips
    // the prefix back off, so callers see the same bare key names Backend
    // A would give them (e.g. 'DISCUSS_S4-260123', not the namespaced form).
    getKeys: function () {
      return scriptProps.getKeys()
        .filter(function (k) { return k.indexOf(prefix) === 0; })
        .map(function (k) { return k.slice(prefix.length); });
    },
    getProperties: function () {
      const self = this;
      const result = {};
      this.getKeys().forEach(function (key) {
        result[key] = self.getProperty(key);
      });
      return result;
    },
    // Mirrors real Properties.setProperties(properties, deleteAllOthers):
    // deleteAllOthers, if true, clears every OTHER key in this document's
    // namespace first (never touches other documents' namespaced keys, and
    // never touches non-namespaced Script Properties like REVIEWER_API_TOKEN).
    setProperties: function (properties, deleteAllOthers) {
      const self = this;
      if (deleteAllOthers) {
        this.getKeys().forEach(function (key) { self.deleteProperty(key); });
      }
      Object.keys(properties || {}).forEach(function (key) {
        self.setProperty(key, properties[key]);
      });
    }
  };
}

/**
 * ADDON-002/ADDON-004: the single seam business logic should go through
 * instead of calling PropertiesService.getDocumentProperties() directly,
 * so a future background/trigger execution can supply a documentId-keyed
 * backend without every call site needing to know which backend it's
 * talking to.
 *
 * Backend selection -- exactly ONE authoritative backend per document,
 * never both (ADDON-004 requirement: no dual-write, no post-adoption
 * synchronization back into Document Properties):
 *
 *   - no context, or context.mode is 'bound'/omitted: Backend A
 *     (PropertiesService.getDocumentProperties()) -- the REAL current
 *     document's properties, exactly as today. Correct because a
 *     'bound'/no-context execution is only ever entered from that
 *     document's own real session (the legacy bound script) -- never
 *     constructed pointing at a document other than the one actually
 *     executing right now. This is also the exact behavior every existing
 *     zero-argument caller gets, unchanged.
 *
 *   - context.mode === 'addon-interactive': resolves the active
 *     document's id (context.documentId/context.document, or -- same as
 *     the legacy path -- DocumentApp.getActiveDocument() via
 *     getReportDocumentId_()), then looks it up with
 *     getRegisteredReportDocument_():
 *       - registered (adopted, via adoptReportDocumentForAddon_()):
 *         Backend B -- central state is authoritative for an adopted
 *         document, interactive or not, so an add-on menu/dialog action
 *         on an adopted document reads/writes the SAME state a
 *         background scheduler run would.
 *       - NOT registered: Backend A (Document Properties) -- explicit,
 *         not a fallback-by-accident. A document that has never been
 *         adopted has no central state to speak of; silently handing
 *         business logic an empty Backend B store here would look
 *         indistinguishable from "this document has no configuration",
 *         which is exactly the silent-configuration-loss ADDON-004
 *         forbids. Document Properties remains authoritative until
 *         adoptReportDocumentForAddon_() explicitly runs.
 *
 *   - context.mode === 'addon-background': Backend B, namespaced by
 *     context.documentId (falling back to context.document.getId() if
 *     only a Document object was supplied). No registration check here --
 *     that gate belongs to the caller (continuousUpdateForDocument_(),
 *     ADDON-004 §3), keeping this function a mechanical backend lookup.
 */
function getReportStateStore_(context) {
  if (context && context.mode === 'addon-background') {
    const documentId = context.documentId ||
      (context.document && context.document.getId());
    return makeCentralReportStateStore_(documentId);
  }
  if (context && context.mode === 'addon-interactive') {
    const documentId = context.documentId || getReportDocumentId_(context);
    if (getRegisteredReportDocument_(documentId)) {
      return makeCentralReportStateStore_(documentId);
    }
    // Not adopted: fall through to Backend A below, deliberately -- see
    // the header comment above.
  }
  return PropertiesService.getDocumentProperties();
}

// =========================================================
// ADDON-003 -- CENTRAL DOCUMENT REGISTRY
// =========================================================
//
// A separate Script-Properties namespace from Backend B's
// 'SA4_STATE|<documentId>|<key>' report state (ADDON-002) -- the registry
// records WHICH documents the future scheduler (ADDON-004) should look at;
// it never holds report state itself. Nothing in this file creates a
// trigger or reads this registry yet.
//
// Deliberately NOT one ever-growing JSON blob: a small index (just the
// list of documentIds) plus one small property PER registered document.
// Registering another document grows the index by one short string and
// adds one new, small, bounded-size property -- it never grows any
// EXISTING document's own property, and the whole registry stays far
// under the 500KB/store or 9KB/value limits (ADDON-002's capacity
// analysis) for any realistic number of meeting documents.
function reportRegistryIndexKey_() {
  return 'SA4_REGISTRY_INDEX';
}

function reportRegistryDocKey_(documentId) {
  if (!documentId) throw new Error('reportRegistryDocKey_: documentId is required.');
  return 'SA4_REGISTRY_DOC|' + documentId;
}

// Add-on time-driven triggers cannot run more frequently than once per
// hour (ADDON-001B, verified against Apps Script's Editor-add-on trigger
// documentation) -- so, unlike the bound-script UI's 15/30/60-minute
// options (Code.js manageTriggers(), unchanged here), no sub-hour value is
// ever valid for a registry entry's intervalHours.
var REPORT_REGISTRY_ALLOWED_INTERVAL_HOURS_ = [1, 2, 4, 6, 12, 24];

function loadReportRegistryIndex_() {
  const raw = PropertiesService.getScriptProperties().getProperty(reportRegistryIndexKey_());
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    // Malformed/unexpected-shape index data fails safe -- treated as an
    // empty registry rather than throwing, so a corrupted index never
    // blocks every other registry operation.
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(function (id) { return typeof id === 'string' && id; });
  } catch (e) {
    return [];
  }
}

function saveReportRegistryIndex_(documentIds) {
  PropertiesService.getScriptProperties().setProperty(reportRegistryIndexKey_(), JSON.stringify(documentIds));
}

/**
 * Registers a document, or updates its entry if already registered
 * (upsert -- this is also what updateRegisteredReportDocument_(),
 * repeat adoption, and -- as of ADDON-004 -- continuousUpdateForDocument_()
 * go through to advance lastRunAt after a successful run). registeredAt is
 * always preserved from any existing entry (set once, at first
 * registration, never reset). lastRunAt is preserved from the existing
 * entry UNLESS `details.lastRunAt` is explicitly provided, in which case
 * that explicit value wins -- this is the one field ADDON-004's scheduler
 * path needs to actively advance; every other caller that omits
 * `lastRunAt` continues to leave it untouched exactly as before.
 */
function registerReportDocument_(documentId, details) {
  if (!documentId) throw new Error('registerReportDocument_: documentId is required.');
  const info = details || {};

  if (info.intervalHours !== undefined &&
      REPORT_REGISTRY_ALLOWED_INTERVAL_HOURS_.indexOf(info.intervalHours) === -1) {
    throw new Error(
      'registerReportDocument_: intervalHours must be one of ' +
      REPORT_REGISTRY_ALLOWED_INTERVAL_HOURS_.join(', ') + ' (add-on time-driven ' +
      'triggers cannot run more frequently than once per hour) -- got ' + info.intervalHours + '.'
    );
  }

  const existing = getRegisteredReportDocument_(documentId);

  const entry = {
    documentId: documentId,
    enabled: info.enabled !== undefined ? !!info.enabled : (existing ? existing.enabled : false),
    meetingId: info.meetingId !== undefined ? info.meetingId : (existing ? existing.meetingId : ''),
    meetingName: info.meetingName !== undefined ? info.meetingName : (existing ? existing.meetingName : ''),
    registeredAt: existing ? existing.registeredAt : new Date().toISOString(),
    lastRunAt: info.lastRunAt !== undefined ? info.lastRunAt : (existing ? existing.lastRunAt : null),
    intervalHours: info.intervalHours !== undefined ? info.intervalHours : (existing ? existing.intervalHours : 1)
  };

  PropertiesService.getScriptProperties().setProperty(reportRegistryDocKey_(documentId), JSON.stringify(entry));

  const index = loadReportRegistryIndex_();
  if (index.indexOf(documentId) === -1) {
    index.push(documentId);
    saveReportRegistryIndex_(index);
  }

  return entry;
}

/**
 * Removes a document from the scheduler registry only. Does NOT touch
 * Document Properties (the bound-script source of truth is never affected
 * by registry membership) and does NOT touch the document's central
 * report state (SA4_STATE|<documentId>|... survives unregistration, so a
 * document can be re-registered later without re-adopting from scratch --
 * see deleteCentralReportState_() below for the separate, explicit,
 * destructive operation if that state ever needs to be purged).
 */
function unregisterReportDocument_(documentId) {
  if (!documentId) throw new Error('unregisterReportDocument_: documentId is required.');

  PropertiesService.getScriptProperties().deleteProperty(reportRegistryDocKey_(documentId));

  const index = loadReportRegistryIndex_().filter(function (id) { return id !== documentId; });
  saveReportRegistryIndex_(index);
}

function getRegisteredReportDocument_(documentId) {
  if (!documentId) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(reportRegistryDocKey_(documentId));
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    // Malformed entry fails safe -- treated as "not registered" rather
    // than throwing or returning corrupt data to a caller.
    return null;
  }
}

function listRegisteredReportDocuments_() {
  return loadReportRegistryIndex_()
    .map(function (id) { return getRegisteredReportDocument_(id); })
    .filter(function (entry) { return entry !== null; });
}

function updateRegisteredReportDocument_(documentId, patch) {
  const existing = getRegisteredReportDocument_(documentId);
  if (!existing) {
    throw new Error('updateRegisteredReportDocument_: document ' + documentId + ' is not registered.');
  }
  return registerReportDocument_(documentId, Object.assign({}, existing, patch || {}));
}

/**
 * Destructive, explicit-only: deletes a document's ENTIRE central report
 * state (Backend B, every 'SA4_STATE|<documentId>|...' key). Does NOT
 * touch the registry entry (a separate, orthogonal piece of state -- see
 * unregisterReportDocument_() above) and does NOT touch Document
 * Properties. Never called automatically anywhere in this file.
 */
function deleteCentralReportState_(documentId) {
  if (!documentId) throw new Error('deleteCentralReportState_: documentId is required.');
  const store = makeCentralReportStateStore_(documentId);
  store.getKeys().forEach(function (key) { store.deleteProperty(key); });
}

// ---------------------------------------------------------------
// ADDON-003 -- DOCUMENT ADOPTION (Document Properties -> Backend B)
// ---------------------------------------------------------------
//
// Property classification driving what gets copied (see the ADDON-003
// report for the full call-graph trace behind this):
//
//   REQUIRED for a future continuousUpdateForDocument_(documentId) to
//   behave correctly with NO Document Properties access:
//     - every getReportConfig_()/getMeetingIdentityConfig_() field
//       (meeting/report identity, paths, agenda, options)
//     - FETCH_ABSTRACTS_ON_UPDATE (the user's own automation choice --
//       omitting it would silently change what the user asked for)
//     - REVISION_MAP (genuinely cross-run state: continuousUpdateCore_()
//       reads last run's map via findParentRevisedToTable_() BEFORE
//       rearrangeRevisionTables_() recomputes and overwrites it for next
//       time -- not just a performance cache)
//     - DISCUSS_<tdoc>/REVIS_<tdoc> (the collector's per-TDoc dedupe
//       ledgers -- omitting these risks re-inserting emails/revisions
//       already present in the document body as apparent duplicates)
//
//   USEFUL BUT OPTIONAL (safe to omit -- self-healing, TTL'd negative
//   caches; omitting only costs a redundant fetch on the next run, no
//   correctness impact):
//     - REVIEWER_NO_SUMMARY_CACHE_<id>, A1_EMPTY_CACHE_<md5>
//     - DEADLINE_EXTENSIONS (mirrors an in-document table that is itself
//       already background-readable via getReportBody_(); the property is
//       only a fallback for when that table is temporarily missing)
//
//   INTERACTIVE-ONLY (never read by continuousUpdateCore_() or anything
//   it calls -- NOT copied):
//     - PARSED_AGENDA, PARSED_AGENDA_ALL (menu-driven skeleton building
//       only: parseAgendaForReport_(), autoCreateReportStructure(),
//       addTdocTablesOnly())
//     - SKIP_ABSTRACTS_DURING_TABLE_BUILD (createTDocTableFromData_(),
//       interactive skeleton build only)
//
// REVIEWER_API_TOKEN is intentionally never in scope here at all -- it
// already lives in Script Properties (global, not per-document) and stays
// exactly where it is under either backend (see ADDON-002's parity tests).
var ADDON003_ADOPTION_FIXED_KEYS_ = [
  'MEETING_FOLDER', 'MEETING_NUMBER', 'MEETING_ID',
  'FTP_BASE', 'TDOC_LIST_URL',
  'REPORT_SUFFIX', 'AGENDA_ITEM_PREFIX', 'AGENDA_SOURCE_DOC_ID', 'AGENDA_TDOC', 'AGENDA_CSV_URL', 'MEETING_DATE',
  'SHOW_PREVIEW_SNIPPET',
  'MEETING_TYPE', 'MEETING_NAME', 'REVISIONS_URL', 'MAILING_LIST',
  'FETCH_ABSTRACTS_ON_UPDATE',
  'REVISION_MAP',
  'DEADLINE_EXTENSIONS',
  // Ad-hoc sessions: read by every update and build of an ad-hoc report (the
  // Session column and the generated sections are computed from them).
  'ADHOC_SESSIONS',
  // Manual TDoc session assignments and the upload clock (stage B): read by
  // every update of a report that has the Session column.
  'ADHOC_TDOC_SESSIONS'
];

var ADDON003_ADOPTION_PREFIX_FAMILIES_ = [
  'DISCUSS_', 'REVIS_', 'A1_EMPTY_CACHE_', 'REVIEWER_NO_SUMMARY_CACHE_'
];

/**
 * The exact set of keys adoptReportDocumentForAddon_() will copy from a
 * given source store: every fixed key that is actually set (skips absent
 * ones -- nothing invents a value), plus every key belonging to one of the
 * per-TDoc/per-URL cache families, discovered via the store's own
 * getKeys() rather than hardcoded. Pure/read-only, exposed separately so
 * it's independently testable.
 */
function reportStateKeysToAdopt_(sourceStore) {
  const keys = [];

  ADDON003_ADOPTION_FIXED_KEYS_.forEach(function (key) {
    if (sourceStore.getProperty(key) !== null) keys.push(key);
  });

  sourceStore.getKeys().forEach(function (key) {
    if (keys.indexOf(key) !== -1) return;
    const inFamily = ADDON003_ADOPTION_PREFIX_FAMILIES_.some(function (prefix) {
      return key.indexOf(prefix) === 0;
    });
    if (inFamily) keys.push(key);
  });

  return keys;
}

/**
 * Explicit, interactive-only infrastructure operation (ADDON-003 scope --
 * NOT wired into onOpen() or any menu item). Copies the current
 * document's report state from Document Properties (Backend A) into the
 * central, documentId-namespaced backend (Backend B), verifies the copy,
 * and only then registers the document. The source is never modified and
 * never deleted -- Document Properties remains fully intact and usable by
 * the existing bound-script/interactive path exactly as before.
 *
 * The SOURCE is always the real PropertiesService.getDocumentProperties()
 * directly -- deliberately NOT routed through getReportStateStore_(context),
 * even though every other function in this file goes through that seam.
 * Adoption's entire purpose is "copy FROM the legacy per-document backend
 * INTO the central one"; once ADDON-004 made getReportStateStore_()'s
 * addon-interactive mode resolve an ALREADY-registered document to the
 * central backend itself, routing adoption's source through that same
 * seam would make re-running adoption on an already-adopted document copy
 * central state onto itself (a silent no-op) instead of picking up
 * whatever has changed in Document Properties since -- found live, during
 * the ADDON-005 scratch-document acceptance test, when "Enable Automatic
 * Updates" (which re-adopts, see below) stopped picking up a TDOC_LIST_URL
 * change made via the legacy "Configure Meeting Settings" dialog.
 *
 * Idempotent: calling this again on an already-adopted document re-copies
 * the CURRENT Document Properties values (a merge write, never
 * deleteAllOthers) and re-verifies; registerReportDocument_()'s own
 * upsert behavior means registeredAt is preserved from the first
 * adoption, not reset.
 *
 * On a verification failure, registerReportDocument_() is never called --
 * a document can never end up registered without its central state having
 * been verified first (no half-registered state).
 */
function adoptReportDocumentForAddon_(context) {
  if (context && context.mode === 'addon-background') {
    throw new Error(
      'adoptReportDocumentForAddon_: interactive-only -- must run in the ' +
      'document\'s real bound/addon-interactive session (source of truth is ' +
      'live Document Properties), not addon-background mode.'
    );
  }

  const documentId = getReportDocumentId_(context);
  const source = PropertiesService.getDocumentProperties();
  const target = getReportStateStore_({ documentId: documentId, mode: 'addon-background' });

  const keysToCopy = reportStateKeysToAdopt_(source);
  const toWrite = {};
  keysToCopy.forEach(function (key) { toWrite[key] = source.getProperty(key); });
  target.setProperties(toWrite);

  const mismatches = keysToCopy.filter(function (key) {
    return target.getProperty(key) !== source.getProperty(key);
  });
  const verified = mismatches.length === 0;

  let registered = false;
  let registryEntry = null;

  if (verified) {
    // Same reasoning as `source` above: read the just-copied identity
    // fields straight back from the source Document Properties (via
    // getReportConfig_({mode:'bound'}), which always uses Backend A),
    // never through the possibly-already-central addon-interactive path.
    const cfg = getReportConfig_({ mode: 'bound' });
    registryEntry = registerReportDocument_(documentId, {
      meetingId: cfg.MEETING_ID,
      meetingName: registryMeetingName_(source, cfg)
    });
    registered = true;
  }

  return {
    documentId: documentId,
    copiedKeys: keysToCopy,
    verified: verified,
    mismatches: mismatches,
    registered: registered,
    registryEntry: registryEntry
  };
}

// =========================================================
// CENTRALIZED CONFIGURATION
// =========================================================

function getReportConfig_(context) {
  // ADDON-002: routed through getReportStateStore_() instead of calling
  // PropertiesService.getDocumentProperties() directly, so a future
  // background execution can supply a documentId-keyed backend. With no
  // context (every existing caller, unchanged) getReportStateStore_()
  // returns PropertiesService.getDocumentProperties() itself -- same
  // object, same calls, same behavior as before this function took a
  // parameter at all.
  const props = getReportStateStore_(context);

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
  // SA4-PROD-007A: optional, generic -- no reliable meeting-date value is
  // available anywhere else during a build (no date field exists on a
  // parsed agenda item, a TDoc-list row, or MeetingContext today). Used
  // only by ad-hoc opening-content generation in
  // buildSkeletonWithTdocTables(); main meetings never read this. Empty
  // by default -- never invented, never defaulted to any specific date.
  const MEETING_DATE = props.getProperty('MEETING_DATE') || '';

  // ========================================
  // 6. OPTIONS
  // ========================================
  const SHOW_PREVIEW_SNIPPET = props.getProperty('SHOW_PREVIEW_SNIPPET') !== 'false';
  
  // ========================================
  // 7. API INTEGRATION (Optional)
  // ========================================
  const REVIEWER_API_TOKEN = getReviewerApiTokenForRun_(); // TEMPLATE-002C: one read per execution
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
    MEETING_DATE,

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
function getMeetingIdentityConfig_(context) {
  // ADDON-002: same seam as getReportConfig_() above -- no behavior change
  // for the existing no-argument call.
  const props = getReportStateStore_(context);
  return {
    MEETING_TYPE: normalizeMeetingType_(props.getProperty('MEETING_TYPE')),
    MEETING_NAME: String(props.getProperty('MEETING_NAME') || '').trim(),
    FTP_BASE: String(props.getProperty('FTP_BASE') || '').trim(),
    TDOC_LIST_URL: String(props.getProperty('TDOC_LIST_URL') || '').trim(),
    AGENDA_TDOC: String(props.getProperty('AGENDA_TDOC') || '').trim(),
    // ADDON-008A: ad-hoc agenda.csv structure source (separate from AGENDA_TDOC).
    AGENDA_CSV_URL: String(props.getProperty('AGENDA_CSV_URL') || '').trim(),
    AGENDA_SOURCE_DOC_ID: String(props.getProperty('AGENDA_SOURCE_DOC_ID') || '').trim(),
    REVISIONS_URL: String(props.getProperty('REVISIONS_URL') || '').trim(),
    // SA4-PROD-007A: generic ad-hoc mailing-list override, same raw-read/
    // override pattern as every other identity field above -- no
    // meeting-ID-specific value here. Absent by default; the exact
    // symbolic ETSI list identifier for a given ad-hoc meeting (e.g.
    // FS_6G_MED) must be confirmed externally and set explicitly, never
    // guessed by this function.
    MAILING_LIST: String(props.getProperty('MAILING_LIST') || '').trim()
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

function getMeetingContext_(context) {
  // ADDON-002: threads the same optional context into both underlying
  // reads. No existing caller passes one, so this remains the same two
  // PropertiesService.getDocumentProperties() reads it always was.
  const cfg = getReportConfig_(context);
  const identity = getMeetingIdentityConfig_(context);

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
    // properties introduced). draftsFolder is handled via `adhocDerived`
    // below, not here, since it has no override property today.
    //
    // mailingList: SA4-PROD-007A wires in the generic MAILING_LIST
    // override here, reusing resolveMeetingSources_()'s existing
    // precedence rule (a non-empty override wins; falls back to
    // adhocDerived.mailingList below otherwise). No meeting-ID-specific
    // value is introduced -- an ad-hoc meeting with no MAILING_LIST set
    // still falls through to the same provisional cfg.LIST_NAME reuse as
    // before, unchanged.
    const adhocOverrides = {
      ftpBase: identity.FTP_BASE,
      tdocListUrl: identity.TDOC_LIST_URL,
      agendaTdoc: identity.AGENDA_TDOC,
      agendaTemplateDocId: identity.AGENDA_SOURCE_DOC_ID,
      revisionsUrl: identity.REVISIONS_URL,
      mailingList: identity.MAILING_LIST
    };

    // mailingList (fallback, used only when identity.MAILING_LIST is
    // unset): PROVISIONAL reuse of the existing report-type -> mailing
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

function getCollectorConfig_(context) {
  const reportConfig = getReportConfig_(context); // Get centralized config
  // ensureCollectorConfigTable_();
  const collectorTableConfig = readCollectorConfigTable_(context);

  const cfg = { ...reportConfig, ...collectorTableConfig }; // Merge configs

  // Set defaults for collector-specific settings if not in table
  if (!cfg.ARCHIVE_DAYS_BACK) cfg.ARCHIVE_DAYS_BACK = '14';
  if (!cfg.A1_EMPTY_CACHE_TTL_HOURS) cfg.A1_EMPTY_CACHE_TTL_HOURS = '24';
  if (!cfg.TIMEZONE) cfg.TIMEZONE = Session.getScriptTimeZone();
  if (!cfg.SHOW_PREVIEW_SNIPPET) cfg.SHOW_PREVIEW_SNIPPET = 'false';
  if (!cfg.TDOC_ID_REGEX) cfg.TDOC_ID_REGEX = '^S4-\\d{6}$';
  if (!cfg.EMAIL_START_DATE) cfg.EMAIL_START_DATE = resolveCollectorStartDate_(context);

  // ADDON-008A2: read the list the report is configured for -- the same
  // precedence getMeetingContext_() already applies (an ad-hoc meeting's
  // saved MAILING_LIST override, else the report-family list; main meetings
  // keep the family list) -- instead of always the family list.
  //
  // TEMPLATE-002A (port of Legacy BUGFIX-LEGACY-003, 123369b): a Collector
  // Configuration table's own LIST_NAME/RSS_URL_V2 still wins -- unless it
  // merely repeats the derived family default (the snapshot "Create
  // Configuration Tables" writes), which is not an intentional override and
  // used to mask the meeting's MAILING_LIST.
  const tableList = String(collectorTableConfig.LIST_NAME || '').trim();
  const tableRss = String(collectorTableConfig.RSS_URL_V2 || '').trim();
  const explicitList = tableList && tableList !== reportConfig.LIST_NAME;
  const explicitRss = tableRss && tableRss !== reportConfig.RSS_URL_V2;
  if (explicitList || explicitRss) {
    // The A1 archive and RSS read the same list: derive whichever one the
    // table does not set from the one it does.
    if (explicitList && !explicitRss) {
      cfg.RSS_URL_V2 = buildCollectorRssUrl_(tableList, '2.0');
      cfg.RSS_URL_V1 = buildCollectorRssUrl_(tableList, '1.0');
    } else if (explicitRss && !explicitList) {
      const rssList = normalizeEtsiListName_((tableRss.match(/[?&]L=([^&]+)/i) || [])[1]);
      if (rssList) cfg.LIST_NAME = rssList;
    }
  } else {
    const listName = resolveCollectorListName_(context, reportConfig.LIST_NAME);
    cfg.LIST_NAME = listName;
    cfg.RSS_URL_V2 = buildCollectorRssUrl_(listName, '2.0');
    cfg.RSS_URL_V1 = buildCollectorRssUrl_(listName, '1.0');
  }

  return cfg;
}

function buildCollectorRssUrl_(listName, version) {
  return `https://list.etsi.org/scripts/wa.exe?RSS&L=${listName}&v=${version}&LIMIT=2000`;
}

/**
 * TEMPLATE-002A (port of Legacy BUGFIX-LEGACY-003): an ETSI list identifier
 * as used in list.etsi.org URLs, from either stored form:
 * "3GPP_TSG_SA4_FS_6G_MED" or "3gpp_tsg_sa4_fs_6g_med@list.etsi.org" (the
 * form deriveEmailExportRecipientFromMailingList_() also accepts). Returns
 * '' for anything else (CR/LF, spaces, other domains, URLs...).
 */
function normalizeEtsiListName_(value) {
  const raw = String(value === null || value === undefined ? '' : value).trim();
  if (!raw || /[\r\n]/.test(raw)) return '';
  const name = raw.replace(/@list\.etsi\.org$/i, '');
  return /^[A-Za-z0-9][A-Za-z0-9_-]{2,63}$/.test(name) ? name : '';
}

/**
 * ADDON-008A2: the mailing list the collector reads: the meeting's
 * configured MAILING_LIST (ad-hoc only, via getMeetingContext_()), else the
 * report-family list. A configured value that is not a valid list
 * identifier is not used: the family list is read instead and the rejection
 * is logged.
 */
function resolveCollectorListName_(context, familyListName) {
  const configured = String(getMeetingContext_(context).sources.mailingList || '').trim();
  if (!configured || configured === familyListName) return familyListName;
  const name = normalizeEtsiListName_(configured);
  if (!name) {
    Logger.log('Ignoring configured mailing list "' + configured.replace(/[\r\n]/g, ' ') + '" (not a valid ETSI list name); reading ' + familyListName);
    return familyListName;
  }
  return name;
}

/**
 * TEMPLATE-002A (port of Legacy BUGFIX-LEGACY-003): the collection start
 * date when no Collector Configuration table sets one -- a stored
 * EMAIL_START_DATE (YYYY-MM-DD) when valid, else the historical default.
 * Read through the report state store, so a background run reads the
 * adopted document's central state like every other setting.
 */
function resolveCollectorStartDate_(context) {
  const store = getReportStateStore_(context);
  const saved = String(store.getProperty('EMAIL_START_DATE') || '').trim();
  if (isValidCollectorStartDate_(saved)) return saved;
  // TEMPLATE-002B (decision 2026-10-01): in the template runtime a report
  // without a usable EMAIL_START_DATE collects from the meeting start date.
  const meetingStart = templateRuntimeRelease_() ? meetingStartDateIso_(store.getProperty('MEETING_DATE')) : '';
  const fallback = meetingStart || '2026-08-21';
  if (saved) Logger.log('Ignoring EMAIL_START_DATE "' + saved.replace(/[\r\n]/g, ' ') + '" (expected YYYY-MM-DD); using ' + fallback);
  return fallback;
}

/** The one rule for an e-mail collection start date: a real YYYY-MM-DD date. */
function isValidCollectorStartDate_(value) {
  const s = String(value === null || value === undefined ? '' : value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(new Date(s).getTime());
}

/**
 * TEMPLATE-002B: the meeting start date as YYYY-MM-DD, from MEETING_DATE as
 * the resolver stores it ("October 1, 2026", computeMeetingDateFromStartDate_())
 * or from a value that already starts with an ISO date. '' when it cannot
 * be read -- never a guess.
 */
function meetingStartDateIso_(meetingDateText) {
  const s = String(meetingDateText === null || meetingDateText === undefined ? '' : meetingDateText).trim();
  const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
  let iso = '';
  const isoMatch = s.match(/^(\d{4}-\d{2}-\d{2})(?:$|[^\d])/);
  const textMatch = s.match(/^([A-Za-z]+) (\d{1,2}), (\d{4})$/);
  if (isoMatch) {
    iso = isoMatch[1];
  } else if (textMatch && MONTHS.indexOf(textMatch[1].toLowerCase()) !== -1) {
    iso = textMatch[3] + '-' + String(MONTHS.indexOf(textMatch[1].toLowerCase()) + 1).padStart(2, '0') + '-' + textMatch[2].padStart(2, '0');
  }
  return isValidCollectorStartDate_(iso) ? iso : '';
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
 * ADDON-007A: reads the administrative anchors back out of an already-built
 * report body by finding the real "<number> Registration of Documents"
 * heading (which the build places as a child of the opening item). Returns
 * null when no such heading exists. No report-family knowledge involved.
 */
function deriveAdminAnchorsFromBuiltDocument_(body) {
  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = child.asParagraph();
    if (para.getHeading() === DocumentApp.ParagraphHeading.NORMAL) continue;
    const m = para.getText().trim().match(/^(\d+(?:\.\d+)*)\s+Registration of Documents\b/);
    if (!m) continue;
    const parts = m[1].split('.');
    if (parts.length < 2) continue;
    const last = parseInt(parts.pop(), 10);
    const parent = parts.join('.');
    return {
      source: 'adhoc-built-document',
      prefixNum: parent,
      openingSection: parent,
      registrationSection: m[1],
      reallocationSection: `${parent}.${last + 1}`,
      documentsSection: `${parent}.${last + 2}`,
      afterDocumentsSection: `${parent}.${last + 3}`,
      iprSection: null,
      iprSyntheticSection: null
    };
  }
  return null;
}

/**
 * Create or ensure reallocation table exists as subsection X.1.3
 */
function ensureReallocationTable_() {
  const body = DocumentApp.getActiveDocument().getBody();

  // Check if table already exists
  for (const t of body.getTables()) {
    if (isReallocationTable_(t)) return;
  }

  // SA4-PROD-002 added an unconditional early return here for any ad-hoc
  // meeting, on the assumption that an ad-hoc agenda never has an X.1.3
  // subsection. SA4-PROD-003 corrected that assumption (clarified
  // production requirement): meeting 86178 DOES want a real
  // "{agendaPrefixNum}.1.3 Document Reallocations" section, functioning
  // exactly as it does for a main SWG meeting -- this function's own
  // insertion-point search (below) already targets "X.1.2"/"X.1.4" as
  // generic boundaries, not anything meeting-type-specific, so no ad-hoc
  // guard belongs here. Reverted to unconditional, byte-identical to its
  // pre-PROD-002 form.

  // Get the agenda prefix from config (e.g., "7." for Audio, "11." for 6G)
  const cfg = getReportConfig_();
  const agendaPrefix = cfg.AGENDA_ITEM_PREFIX || '7.';
  // ADDON-007A: main meetings keep the configured-prefix strings exactly
  // ("7.1.3", "11.1.3", ...). An ad-hoc meeting never uses the report-family
  // prefix: its administrative anchor is read from what the build actually
  // produced (the real "... Registration of Documents" heading), with the real
  // parsed agenda (PARSED_AGENDA) supplying the IPR boundary when available.
  let parsedAgenda = [];
  try {
    const parsed = JSON.parse(PropertiesService.getDocumentProperties().getProperty('PARSED_AGENDA') || '[]');
    if (Array.isArray(parsed)) parsedAgenda = parsed;
  } catch (e) { /* treated as no parsed agenda */ }
  const meetingContext = getMeetingContext_();
  let anchors = getAdministrativeAgendaAnchors_(meetingContext, parsedAgenda, agendaPrefix);
  if (meetingContext.meeting.type === 'adhoc') {
    const built = deriveAdminAnchorsFromBuiltDocument_(body);
    if (built) {
      anchors = Object.assign({}, built, {
        iprSection: anchors.iprSection,
        iprSyntheticSection: anchors.iprSyntheticSection
      });
    }
  }
  if (!anchors.openingSection) {
    Logger.log('ensureReallocationTable_: ad-hoc meeting with no built/parsed administrative anchor -- Document Reallocations not created.');
    return;
  }
  const targetSection = anchors.reallocationSection; // e.g. "7.1.3" (main) or the real opening item's sibling (ad-hoc)
  const beforeMarkers = [anchors.documentsSection, anchors.afterDocumentsSection, anchors.iprSection || anchors.iprSyntheticSection].filter(Boolean);
  
  // Find insertion point: before X.1.4 (preferred) or after X.1.2 (fallback)
  //
  // SA4-PROD-004: this used to be a SINGLE forward scan that checked "is
  // this X.1.2?" before "is this X.1.4-or-later?" on each paragraph -- so
  // whenever X.1.2 (which always exists, and is always immediately
  // followed by its OWN Registration-of-Documents summary table, built by
  // buildSkeletonWithTdocTables()) appeared BEFORE X.1.4 in the document
  // (which it always does), the loop matched and broke on X.1.2 first,
  // computing "insert right after the X.1.2 HEADING paragraph" -- i.e.
  // BEFORE that heading's own summary table, not after the whole X.1.2
  // section. That wedged the reallocation heading+table between X.1.2's
  // heading and its own table, visually merging the reallocation table
  // with the unrelated X.1.2 summary table (and, once X.1.4 exists per
  // SA4-PROD-003, pushing X.1.4 and its own staged TDoc tables further
  // down but never actually separating them from that summary table
  // either). Fixed by searching for the "before X.1.4/X.1.5/X.2" anchor
  // FIRST, in its own complete pass -- it is the more specific, correct
  // anchor whenever it exists (X.1.4 now always exists once there is any
  // registration/staged content) -- and falling back to the original
  // "after X.1.2" pass only when no X.1.4-or-later heading is present at
  // all (a document with no X.1.4 section).
  let insertIndex = -1;
  let foundSection = false;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const para = child.asParagraph();
    const text = para.getText().trim();
    const heading = para.getHeading();

    if (heading !== DocumentApp.ParagraphHeading.NORMAL &&
        beforeMarkers.some(m => text.startsWith(m))) {
      insertIndex = i;
      foundSection = true;
      break;
    }
  }

  if (!foundSection) {
    for (let i = 0; i < body.getNumChildren(); i++) {
      const child = body.getChild(i);
      if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

      const para = child.asParagraph();
      const text = para.getText().trim();
      const heading = para.getHeading();

      if (heading !== DocumentApp.ParagraphHeading.NORMAL &&
          text.startsWith(anchors.registrationSection)) {
        insertIndex = i + 1;
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
          text.startsWith(`${anchors.openingSection} `)) {
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
function getReallocationMap_(context) {
  const body = getActiveDocumentBodyCounted_('getReallocationMap_', context);
  const map = {};

  for (const table of getTablesCounted_(body, 'getReallocationMap_')) {
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

function readCollectorConfigTable_(context) {
  const body = getReportBody_(context);
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
      // TEMPLATE-002A (port of Legacy LEGACY-0099, cb1208c): the same shared
      // interpretation as downloadAndGroupTdocs_() (ADDON-008A1b) -- a
      // removed/withdrawn/n/a TDoc is never imported.
      const target = interpretReallocationTarget_(reallocations[tdoc].new);
      if (target.kind === 'remove') {
        Logger.log(`Skipping ${tdoc}: reallocated to "${target.value}" (${reallocations[tdoc].reason || 'no reason'})`);
        continue;
      }
      agendaItem = target.kind === 'move' ? target.agendaItem : reallocations[tdoc].new;
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
  
  // TEMPLATE-002A (port of Legacy BUGFIX-LEGACY-002, d1021cf): this import
  // (the "📝 Legacy: Build Initial Report" menu item) never clears the
  // document and never checked whether a table for a TDoc already existed,
  // so every further run duplicated every TDoc table it processed. Existing
  // tables are scanned ONCE (the isTDocTable_()/safeCellText_() technique
  // continuousUpdateCore_() uses) and a TDoc already present is skipped.
  // Duplicate tables already in a document are left untouched -- this only
  // prevents new ones. Updated as tables are inserted, so a source list
  // repeating a TDoc number cannot duplicate it either.
  const existingTdocNumbers = {};
  body.getTables().forEach(function (t) {
    if (!isTDocTable_(t)) return;
    const existingTdoc = safeCellText_(t, 0, 1).trim();
    if (existingTdoc) existingTdocNumbers[existingTdoc] = true;
  });

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
      const tdocNumber = String(row[tdocCol] || '').trim();
      if (existingTdocNumbers[tdocNumber]) {
        Logger.log(`Skipping ${tdocNumber}: a table for this TDoc already exists in the document (BUGFIX-LEGACY-002 idempotency guard).`);
        return;
      }
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
      existingTdocNumbers[tdocNumber] = true;
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
  const existing = findAgendaSectionEndIndex_(body, agendaItem);
  if (existing !== -1) return existing;

  // Heading not found: create it and return position after it
  const agendaLevel = (agendaItem.match(/\./g) || []).length + 1;
  const headingText = agendaItem + (agendaTopic ? ' ' + agendaTopic : '');
  const newHeading = body.appendParagraph(headingText);
  if (agendaLevel === 1) newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  else if (agendaLevel === 2) newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  else newHeading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  return body.getNumChildren();
}

/**
 * ADDON-008A1: the read-only half of findInsertionPointForAgendaItem_() --
 * the body index just before the next same-or-higher-level heading after
 * this agenda item's heading (or the body end), or -1 when the report has
 * no heading for it. Never modifies the document.
 */
function findAgendaSectionEndIndex_(body, agendaItem) {
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

  return -1;
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
function createSummaryTable_(body, tdocs, tdocCol, titleCol, sourceCol, agendaCol, sessionResolver) {
  const summaryTable = body.appendTable();

  // Add header row
  const headerRow = summaryTable.appendTableRow();
  headerRow.appendTableCell('TDoc');
  headerRow.appendTableCell('Title');
  headerRow.appendTableCell('Source');
  headerRow.appendTableCell('Agenda Item');
  // Ad-hoc sessions (stage B): a fifth column, only when the caller passes
  // the sessions of an ad-hoc report (makeAdhocTdocSessionResolver_()).
  if (sessionResolver) headerRow.appendTableCell(ADHOC_SESSION_COLUMN_HEADER_);

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
    if (sessionResolver) dataRow.appendTableCell(sessionResolver.cell(tdocData));
  });

  // Remove initial empty row if present
  removeInitialEmptyRow_(summaryTable);
  // With the Session column the columns get explicit widths (the four-column table is left as it is).
  if (sessionResolver) applyRegistrationTableWidths_(summaryTable);
  
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
  // SA4-PROD-006: this used to gate on a hardcoded main-meeting-only
  // /^S4-\d{6}$/ regex, so no ad-hoc TDoc (e.g. "S4aP260098") could ever
  // reach fetchAndAddAbstract_(). Migrated to the same centralized,
  // registered-family identifier check the rest of the codebase already
  // uses (parseExactSA4DocumentId_() / SA4_TDOC_FAMILIES) -- no new regex
  // introduced. parsed.raw (not the raw cell text) is passed to
  // fetchAndAddAbstract_() so the API always receives the canonical
  // identifier spelling.
  //
  // PERF-003 (Part D, characterization only -- no behavior change): this
  // is a SEPARATE gate from FETCH_ABSTRACTS_ON_UPDATE (getFetchAbstractsSetting_(),
  // see the 2.3.0 changelog entry above and addAbstractsForTables_()).
  // FETCH_ABSTRACTS_ON_UPDATE only controls continuousUpdate()'s bulk sweep
  // over EXISTING tables that don't have an abstract yet. This call site is
  // reached every time ANY new TDoc table is created -- including from
  // inside continuousUpdate() when it inserts a brand-new TDoc this run --
  // regardless of the FETCH_ABSTRACTS_ON_UPDATE setting. It is only skipped
  // when SKIP_ABSTRACTS_DURING_TABLE_BUILD is explicitly set, which happens
  // solely inside addTdocTablesOnly()'s manual two-step "tables now,
  // abstracts later" workflow.
  //
  // PERF-003B (correction): the PERF-002 measured run logged "Added 0 new,
  // updated 0 statuses", so insertNewTdoc_() (and therefore this call
  // site) was never reached that run -- it cannot be the source of that
  // run's 1 Reviewer API request. That request traced instead to
  // continuousUpdate()'s `if (getFetchAbstractsSetting_())` branch calling
  // addAbstractsForTables_() (see its own call site), which means
  // FETCH_ABSTRACTS_ON_UPDATE's Document Property was actually 'true' for
  // that document at measurement time -- "default OFF" only describes the
  // value when the property was never set; once explicitly enabled via
  // the trigger dialog's checkbox (createContinuousTrigger's
  // fetchAbstracts argument / setFetchAbstractsSetting()), it stays on
  // until explicitly turned off. addAbstractsForTables_() then found
  // exactly 1 existing TDoc table with no Abstract cell yet and fetched
  // it, matching the measured "Fetched abstracts for 1 table(s)" /
  // "Reviewer API requests: 1". This call site's own behavior (fetching a
  // just-created table's abstract immediately, independent of the
  // trigger-level toggle) is still real and intentional, it just was not
  // what fired in that specific measured run.
  const skipAbstracts = SKIP_ABSTRACTS_DURING_BUILD_THIS_RUN_ ||
    PropertiesService.getDocumentProperties().getProperty('SKIP_ABSTRACTS_DURING_TABLE_BUILD') === 'true';
  const tdocNumber = String(data[0][1] || '').trim();
  const parsedTdoc = parseExactSA4DocumentId_(tdocNumber);
  if (!skipAbstracts && parsedTdoc.isValid) {
    fetchAndAddAbstract_(table, parsedTdoc.raw);
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
function fetchAndAddAbstract_(table, tdocNumber, context) {
  try {
    // TEMPLATE-002C: read once per execution; a missing token is logged once.
    const apiToken = getReviewerApiTokenForRun_();
    if (!apiToken) {
      if (!REVIEWER_TOKEN_RUN_STATE_.missingLogged) {
        Logger.log('No REVIEWER_API_TOKEN found in script properties');
        REVIEWER_TOKEN_RUN_STATE_.missingLogged = true;
      }
      return;
    }

    // PERF-006 (Part A): canonicalize via the central SA4 document-ID
    // registry (never a second parser) before touching the negative
    // cache. Both existing call sites already pass an already-canonical
    // tdocNumber (parsedTdoc.raw from parseExactSA4DocumentId_()), but
    // re-deriving it here makes the cache self-contained rather than
    // relying on every future caller to canonicalize first. If somehow
    // not a recognized identifier, the cache is simply not consulted --
    // the request still proceeds exactly as before (no behavior change).
    const canonical = parseExactSA4DocumentId_(tdocNumber);
    const cacheId = canonical.isValid ? canonical.raw : null;

    if (cacheId && isReviewerNoSummaryCached_(cacheId, context)) {
      perfCount_('Reviewer negative-cache hits');
      Logger.log(`Skipping Reviewer API for ${tdocNumber}: no summary was available on a previous run (cached, within TTL)`);
      return;
    }

    const apiUrl = `https://reviewer.bouazizi.dev/api/v1/documents/${tdocNumber}/summary?type=summary`;

    perfCount_('UrlFetchApp.fetch() calls (total)');
    perfCount_('fetch() call site: fetchAndAddAbstract_ (Reviewer API)');
    perfCount_('Reviewer API requests');
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
        // PERF-006 (Part A): a real summary just arrived -- clear any
        // stale negative-cache entry (e.g. Reviewer had not yet finished
        // processing this TDoc on an earlier run and returned 404 then).
        if (cacheId) clearReviewerNoSummaryCache_(cacheId, context);

        const insertIndex = findAbstractInsertIndex_(table);
        // PERF-003B (Part 3): this inserts a brand-new ROW into an
        // EXISTING table -- a genuine structural change (new row/cells
        // default to non-zero height/spacing, same reason
        // removeRowHeightAndSpacing() exists at all -- see its header
        // comment), reachable from continuousUpdate() via
        // addAbstractsForTables_() even when newTdocsAdded/rev.moved/the
        // revision-linked-table counter are all zero (exactly the
        // PERF-002 measured run: 0 new TDocs, 1 abstract row inserted).
        // shouldReformatAfterUpdate_()'s call site reads this counter as
        // its 4th signal so that case still gets formatted.
        perfCount_('structural: abstract row inserted (fetchAndAddAbstract_)');
        const abstractRow = table.insertTableRow(insertIndex);
        abstractRow.appendTableCell('Abstract');
        const abstractCell = abstractRow.appendTableCell(abstractText);
        
        // Set abstract text to font size 8
        const abstractTextElement = abstractCell.editAsText();
        abstractTextElement.setFontSize(8);
        
        Logger.log(`Added abstract for ${tdocNumber}`);
      }
    } else if (statusCode === 404) {
      // PERF-006 (Part A): 404 is the Reviewer API's own explicit "no
      // summary exists for this document" contract -- the one genuinely
      // definitive "does not exist" signal this endpoint returns, and
      // exactly the case PERF-005 measured recurring every run. This is
      // the ONLY status cached negative. Deliberately NOT cached:
      //  - any other 4xx (e.g. 401/403 auth failures, 400 bad request) --
      //    a client/auth/config problem, not "no summary", and caching it
      //    would silently hide a real misconfiguration on every future run;
      //  - any 5xx -- a server-side/transient failure, not a content fact;
      //  - a 200 with an empty body -- Reviewer accepted the request but
      //    may simply not have finished processing this TDoc yet, which is
      //    not proven permanent, so it is left exactly as before (no
      //    insertion, no cache write, retried next run, same as pre-PERF-006);
      //  - the catch block below (exceptions: network failures, JSON parse
      //    failures, malformed responses) -- none of these are a content
      //    fact about the document, so none are cached.
      Logger.log(`No summary available for ${tdocNumber}`);
      if (cacheId) {
        markReviewerNoSummary_(cacheId, statusCode, context);
        perfCount_('Reviewer negative-cache writes');
      }
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

function collectorUpdate_(context) {
  const cfg = getCollectorConfig_(context);
  log_(cfg, '=== COLLECTOR START ===');

  // 2.17.4: a failed step is still logged and still does not stop the
  // other one; it is now also returned, so that a caller which wants to
  // (Update Report Now in a template report) can report it.
  const failures = [];
  try { perfTimed_('  collector: RSS/mail (checkRSSFeed_)', () => checkRSSFeed_(cfg, context)); } catch (e) { warn_(cfg, 'checkRSSFeed_ failed: ' + e.message); failures.push({ step: 'e-mail discussions', error: e.message }); }
  try { perfTimed_('  collector: revisions (updateRevisions_)', () => updateRevisions_(cfg, context)); } catch (e) { warn_(cfg, 'updateRevisions_ failed: ' + e.message); failures.push({ step: 'TDoc revisions', error: e.message }); }

  log_(cfg, '=== COLLECTOR END ===');
  return { failures: failures };
}

// --- Email discussion ---
function checkRSSFeed_(cfg, context) {
  // Guard against missing config (prevents your toLowerCase crash)
  const listName = String(cfg.LIST_NAME || LIST_NAME_LOCK);
  const listLower = listName.toLowerCase();

  const msgs = perfTimed_('    RSS/mail: collectHybridListservMessages_ (fetch, all sources)', () => collectHybridListservMessages_(cfg, context));
  perfCount_('RSS/mail messages collected (this run)', msgs.length);
  const tz = String(cfg.TIMEZONE || Session.getScriptTimeZone());
  const showPreview = String(cfg.SHOW_PREVIEW_SNIPPET || 'false').toLowerCase() === 'true';

  const props = getReportStateStore_(context);
  const tables = getTablesCounted_(getActiveDocumentBodyCounted_('checkRSSFeed_', context), 'checkRSSFeed_');

  // Deadlines extended verbally / by e-mail, keyed by TDOC.
  const extensions = getDeadlineExtensionMap_(context);

  // ADDON-008A2: what each message is about, worked out once per message.
  // A subject naming any recognized full SA4 identifier is associated with
  // exactly those documents, whatever its layout (no bracket format or
  // agenda item required). Only a subject with no full identifier falls back
  // to the legacy short-number parse -- and only for main-meeting S4- tables,
  // since a bare "082" cannot tell S4aI260082 from S4aA260082.
  const msgKeys = msgs.map(m => {
    if (!m || !m.title) return null;
    const ids = findSA4DocumentIdsInText_(stripReplyPrefixes_(m.title));
    return { ids: ids, legacy: ids.length ? null : parseEmailSubject_(m.title) };
  });

  const mailProcessingStart = Date.now();
  tables.forEach(t => {
    if (!isTDocTable_(t)) return;

    // ADDON-008A2: the registered SA4 families (all ad-hoc series too), not
    // the legacy cfg.TDOC_ID_REGEX default '^S4-\d{6}$' that skipped every
    // ad-hoc TDoc table -- the same migration PROD-017 made for revisions.
    const tdocRaw = String(safeCellText_(t, 0, 1) || '').trim();
    const tdocParsed = parseExactSA4DocumentId_(tdocRaw);
    if (!tdocParsed.isValid) return;
    const tdoc = tdocParsed.raw;
    const short = computeShortNumber_(tdoc);
    const shortNumberMatchAllowed = tdocParsed.familyKey === 'main';

    const storeKey = 'DISCUSS_' + tdoc;
    const store = loadJsonObject_(props.getProperty(storeKey));

    let added = 0;

    msgs.forEach((m, msgIndex) => {
      const keys = msgKeys[msgIndex];
      if (!keys) return;

      // Check if this email is for our TDoc
      if (keys.ids.length > 0) {
        if (keys.ids.indexOf(tdoc) === -1) return;
      } else {
        const parsed = keys.legacy;
        if (!shortNumberMatchAllowed || !parsed || !parsed.tdocShort) return;
        if (parsed.tdocShort !== short && parsed.tdocFull !== tdoc) return;
      }
      
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
  perfAddTime_('    RSS/mail: per-table message matching + cell rendering', Date.now() - mailProcessingStart);
}

function collectHybridListservMessages_(cfg, context) {
  let out = [];
  // Use only RSS v2.0 to avoid duplicates (v2.0 has more items and better metadata)
  out = out.concat(collectRssItems_(cfg, cfg.RSS_URL_V2, 'rss_v2'));

  const a1Urls = buildArchiveIndexUrlsByDaysBack_(cfg.LIST_NAME, parseInt(cfg.ARCHIVE_DAYS_BACK || '14', 10), cfg);
  // PERF-001/PERF-003 (Part C, characterization): buildArchiveIndexUrlsByDaysBack_()
  // does not generate one URL per day -- ETSI's archive is indexed by
  // WEEK-LETTER (A-E) within each calendar MONTH the ARCHIVE_DAYS_BACK
  // window touches (e.g. "ind2609A".."ind2609E" for September 2026), so a
  // default 14-day window that stays inside one month produces exactly 5
  // URLs, not 14. collectA1_()'s own isA1CachedEmpty_()/markA1Empty_()
  // 24h-TTL cache already skips a re-fetch for any week-page previously
  // found empty and still within its TTL -- this is why a measured run
  // considered 5 URLs but only fetched 3: the other 2 were weeks already
  // known (within the last 24h) to have no new messages. Every
  // NOT-yet-known-empty week (necessarily including the current, still
  // possibly-changing week) is always fetched fresh, every run, by design
  // -- per PERF-003 Part C's explicit constraint, correctness must never
  // depend on skipping a week that could still contain new/current
  // messages, so no further safe reduction was found here; this network
  // cost is genuine ETSI round-trip latency, not redundant work, and is
  // left unchanged.
  perfCount_('A1 archive day-URLs considered (ARCHIVE_DAYS_BACK)', a1Urls.length);
  a1Urls.forEach(url => out = out.concat(collectA1_(cfg, url, context)));

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

function collectA1_(cfg, url, context) {
  if (isA1CachedEmpty_(cfg, url, context)) return [];
  const resp = safeFetch_(cfg, url, {}, 'A1 fetch');
  if (!resp.ok) { markA1Empty_(cfg, url, true, context); return []; }

  const html = resp.text || '';
  if (!html.toLowerCase().includes('a2=')) { markA1Empty_(cfg, url, true, context); return []; }

  const rows = parseA1TableRows_(html, cfg);
  markA1Empty_(cfg, url, rows.length === 0, context);
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
/**
 * PROD-017: the drafts/revisions folder source is now
 * getMeetingContext_().sources.revisionsUrl -- NOT cfg.REVISIONS_URL
 * (getReportConfig_()'s always-computed main-meeting formula, which
 * fabricates a URL for every ad-hoc meeting regardless of whether one is
 * actually configured; see testAllConnections()'s Test 4, which already
 * made exactly this same fix). getMeetingContext_() never touches the
 * network (deterministic from saved Document Properties, per its own
 * contract) -- this is still a pure "read configuration" call, not
 * Meeting-ID resolution. For a main meeting, sources.revisionsUrl is
 * IDENTICAL to cfg.REVISIONS_URL (see getMeetingContext_()'s own main-
 * meeting branch), so this is behavior-preserving there; for an ad-hoc
 * meeting it is the configured REVISIONS_URL override, or genuinely absent
 * (undefined) if none was ever set -- the exact same graceful "not
 * configured" early return below.
 */
/**
 * HOTFIX-001: normalizes a stored/candidate revision entry's identity key
 * for deduplication -- prefers its link (URL), falling back to its
 * filename only when no URL was resolved. `decodeURIComponent` is applied
 * so two encodings of the literal same URL (e.g. a space as "%20" vs a
 * literal space) collapse to the SAME key rather than being treated as two
 * different files; an undecodable string is used as-is rather than
 * throwing. Returns '' (never null/undefined) for a genuinely empty entry,
 * so callers can uniformly skip it.
 */
function normalizeRevisionKey_(item) {
  const raw = String((item && (item.link || item.text)) || '').trim();
  if (!raw) return '';
  try { return decodeURIComponent(raw); } catch (e) { return raw; }
}

/**
 * HOTFIX-001: root cause of the live meeting-86178 corruption (repeated/
 * concatenated filenames, broken hyperlinks in cells like S4aP260068's and
 * S4aP260075's, each of which legitimately has multiple real draft files)
 * was in the ORIGINAL rendering loop, not the matching/store logic itself:
 * it called `cell.editAsText()` ONCE, then repeatedly called
 * `te.appendText(line + '\n')` per revision, re-querying `te.getText().length`
 * for each new offset and calling `te.setLinkUrl()` immediately after each
 * append -- i.e. interleaving live text MUTATION with offset-range
 * CALCULATION and immediate hyperlink APPLICATION against a Text object
 * that was still changing. Apps Script's own documented behavior is that a
 * literal "\n" passed to Text.appendText() does not behave like a plain
 * appended character (it is a paragraph-affecting insertion, not a stable,
 * purely-additive text run) -- continuing to mutate/query the SAME `te`
 * reference afterward is exactly the kind of "stale reference" pattern
 * that produces the concatenation/corruption actually observed live.
 *
 * The fix: build the ENTIRE final cell text as a single, ordinary
 * JavaScript string FIRST (no Apps Script calls at all), replace the
 * cell's contents with exactly ONE cell.setText(fullText) call, and ONLY
 * THEN compute each line's start/end offsets via plain string arithmetic
 * against that now-STABLE, already-final text -- no further text mutation
 * happens after this point, so every setLinkUrl() call operates against
 * offsets that can never have shifted underneath it. This also REPAIRS an
 * already-corrupted cell from a prior run for free: the cell's contents
 * are always fully replaced, never appended to.
 */
function renderRevisionsCellContent_(cell, orderedRevisions) {
  const entries = orderedRevisions
    .map(item => ({ line: String(item.text || '').trim(), link: item.link }))
    .filter(e => e.line);

  if (!entries.length) {
    if (!cell.getText().trim()) cell.setText('No revisions available.');
    return;
  }

  const fullText = entries.map(e => e.line).join('\n');
  cell.setText(fullText);

  const te = cell.editAsText();
  let offset = 0;
  entries.forEach(e => {
    const start = offset;
    const end = start + e.line.length - 1; // setLinkUrl's endOffsetInclusive
    if (e.link) { try { te.setLinkUrl(start, end, e.link); } catch (err) { } }
    offset = end + 2; // skip past this line's '\n' separator
  });
}

function updateRevisions_(cfg, context) {
  const baseUrl = String(getMeetingContext_(context).sources.revisionsUrl || '').trim();
  if (!baseUrl) return;

  const url = baseUrl.replace(/\/$/, '') + '/';
  // PERF-001 stage: "revision folder fetch/listing" -- the network fetch
  // (self-timed inside safeFetch_) plus parsing the directory listing HTML
  // into anchors.
  const resp = safeFetch_(cfg, url, {}, 'Revisions fetch');
  if (!resp.ok) return;

  const anchors = perfTimed_('    revisions: parseAnchors_ (listing parse)', () => parseAnchors_(resp.text || ''));
  perfCount_('revision-folder anchors found (this run)', anchors.length);
  const props = getReportStateStore_(context);
  const body = getActiveDocumentBodyCounted_('updateRevisions_', context);
  const tables = getTablesCounted_(body, 'updateRevisions_');

  // PERF-006 (Part B): pre-parse every anchor into a
  // Map<canonicalTdocId, DraftEntry[]> ONCE, instead of the per-table
  // anchors.forEach(parseDraftAnchor_) loop below re-parsing the SAME
  // anchor once per TDoc table (O(tables*anchors) -> O(anchors)). See
  // buildRevisionAnchorIndex_()'s own header comment -- PERF-005 measured
  // the old matching loop at only ~9ms for this workload, so this is a
  // complexity/code-quality improvement, not a runtime-bottleneck fix.
  const anchorIndex = buildRevisionAnchorIndex_(anchors, url);
  perfCount_('revision anchors parsed', anchors.length);

  let revisionFilesProcessed = 0;
  let matchingMs = 0;
  let renderingMs = 0;

  tables.forEach(t => {
    if (!isTDocTable_(t)) return;

    const tdocRaw = String(safeCellText_(t, 0, 1) || '').trim();
    // PROD-017: the TDoc cell holds nothing but the identifier itself --
    // use the exact-match central parser (all 6 verified SA4 families),
    // not the legacy cfg.TDOC_ID_REGEX default ('^S4-\d{6}$', main-meeting
    // only).
    const tdocParsed = parseExactSA4DocumentId_(tdocRaw);
    if (!tdocParsed.isValid) return;
    const tdoc = tdocParsed.raw;

    const storeKey = 'REVIS_' + tdoc;
    const store = loadJsonObject_(props.getProperty(storeKey));

    // PERF-001 stage: "revision matching" -- per TDoc table, looking up
    // this TDoc's already-parsed drafts (PERF-006 Part B: an O(1) bucket
    // lookup into anchorIndex, instead of re-scanning+re-parsing every
    // anchor here). Canonical-identity match semantics are unchanged --
    // buildRevisionAnchorIndex_() bucketed by the exact same draft.tdocId
    // this loop used to compare against.
    const matchStart = Date.now();
    perfCount_('revision TDoc bucket lookups');
    const drafts = anchorIndex.get(tdoc) || [];
    drafts.forEach(draft => {
      const id = normalizeRevisionKey_({ link: draft.url, text: draft.fileName });
      if (!id) return;
      if (!store[id]) { store[id] = { text: draft.fileName, link: draft.url }; revisionFilesProcessed++; }
    });
    matchingMs += Date.now() - matchStart;

    // HOTFIX-001: REVIS_<tdoc> is an accumulating, cross-run CACHE (so a
    // single transient fetch failure never loses a previously-discovered
    // draft) -- it is intentionally NOT wiped/replaced wholesale here. But
    // it must never be allowed to RENDER a duplicate: re-key every stored
    // entry through the SAME normalizeRevisionKey_() used above, so two
    // entries that only differ by an incidental encoding/casing
    // difference in how their URL was captured on different runs collapse
    // into exactly one rendered line.
    const deduped = {};
    Object.values(store).forEach(item => {
      const key = normalizeRevisionKey_(item);
      if (!key) return;
      if (!deduped[key]) deduped[key] = item;
    });
    props.setProperty(storeKey, JSON.stringify(store));
    const ordered = Object.values(deduped).sort((x, y) => {
      const byText = String(x.text || '').localeCompare(String(y.text || ''));
      return byText !== 0 ? byText : String(x.link || '').localeCompare(String(y.link || ''));
    });

    // PERF-001 stage: "revision cell rendering" -- see
    // renderRevisionsCellContent_() for the HOTFIX-001 rendering fix
    // itself (unchanged, not rewritten here).
    const renderStart = Date.now();
    const cell = findOrFallbackCell_(t, ['Revisions', 'Revisions:', 'Revision'], 6, 1);
    renderRevisionsCellContent_(cell, ordered);
    renderingMs += Date.now() - renderStart;

    // Also insert the revision tables directly after this TDOC table in the document.
    // Find the revised TDOCs and insert their tables immediately after this one.
    insertRevisedDocTablesAfter_(body, t, ordered, cfg);

    log_(cfg, 'Revisions updated', { tdoc, total: ordered.length });
  });

  perfAddTime_('    revisions: matching (per-table anchor scan, accumulated)', matchingMs);
  perfAddTime_('    revisions: cell rendering (renderRevisionsCellContent_, accumulated)', renderingMs);
  perfCount_('revision files newly stored (this run)', revisionFilesProcessed);
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
    // PROD-017: central registry, same as updateRevisions_() -- item.text
    // is a filename (e.g. "S4aP260071_QCOM.docx"), so this searches within
    // it rather than requiring an exact match.
    const revParsed = parseSA4DocumentId_(String(item.text || ''));
    if (!revParsed.isValid) return;
    const revTdoc = revParsed.raw;

    // Check if a table for this revision already exists
    const existing = body.getTables().find(t => {
      if (!isTDocTable_(t)) return false;
      return String(safeCellText_(t, 0, 1) || '').trim() === revTdoc;
    });
    if (existing) return;

    // PERF-003 (Part B): this is the one place inside collectorUpdate_()'s
    // call chain that can insert a brand-new TABLE (structural change) --
    // continuousUpdate()'s end-of-run formatting-skip decision reads this
    // counter to know whether a genuinely structural change happened here,
    // even though nothing else that run needed reformatting.
    perfCount_('structural: new revision-linked tables inserted (insertRevisedDocTablesAfter_)');

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
  { key: 'main',          family: 'main',  seriesCode: null, prefix: 'S4-',  reportFamily: null, body: 'S4-(\\d{6})(?!\\d)' },
  { key: 'audio-adhoc',   family: 'adhoc', seriesCode: 'A',  prefix: 'S4aA', reportFamily: 'Audio', body: 'S4aA(\\d{6})(?!\\d)' },
  { key: 'plenary-adhoc', family: 'adhoc', seriesCode: 'P',  prefix: 'S4aP', reportFamily: null, body: 'S4aP(\\d{6})(?!\\d)' },
  { key: 'video-adhoc',   family: 'adhoc', seriesCode: 'V',  prefix: 'S4aV', reportFamily: 'Video', body: 'S4aV(\\d{6})(?!\\d)' },
  { key: 'mbs-adhoc',     family: 'adhoc', seriesCode: 'I',  prefix: 'S4aI', reportFamily: 'MBS', body: 'S4aI(\\d{6})(?!\\d)' },
  { key: 'rtc-adhoc',     family: 'adhoc', seriesCode: 'R',  prefix: 'A4aR', reportFamily: 'RTC', body: 'A4aR(\\d{6})(?!\\d)' }
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

/**
 * ADDON-008A2: EVERY recognized SA4 document identifier in free text (e.g.
 * an e-mail subject), canonical spelling, deduplicated, in order of first
 * appearance. Same SA4_TDOC_FAMILIES bodies and canonicalization as
 * matchSA4DocumentId_(), scanned globally with boundaries on both sides:
 * the identifier may not continue a word ("XS4aI260082") or run into
 * further letters/digits -- except a revision suffix, which maps to the
 * base document exactly as parseSA4DocumentId_() does ("S4aI260082r1" ->
 * S4aI260082). Nothing is guessed: an unregistered prefix or a look-alike
 * letter ("S4al260064") is not an identifier.
 */
function findSA4DocumentIdsInText_(text) {
  const src = String(text === null || text === undefined ? '' : text);
  const hits = [];
  SA4_TDOC_FAMILIES.forEach(function (fam) {
    const re = new RegExp('(?:^|[^A-Za-z0-9])' + fam.body + '(?=$|[^A-Za-z0-9]|r\\d)', 'gi');
    let m;
    while ((m = re.exec(src)) !== null) {
      hits.push({ index: m.index, id: fam.prefix + m[1] });
      re.lastIndex = m.index + m[0].length;
    }
  });
  hits.sort(function (a, b) { return a.index - b.index; });
  const ids = [];
  hits.forEach(function (h) { if (ids.indexOf(h.id) === -1) ids.push(h.id); });
  return ids;
}

/**
 * PROD-017: parses one <a> anchor from a 3GPP FTP drafts/revisions folder
 * listing (parseAnchors_() output: {href, text}) into a structured draft
 * document record, using the SAME central SA4 TDoc identifier registry
 * every other consumer uses (parseSA4DocumentId_() / SA4_TDOC_FAMILIES) --
 * never a second, isolated 'S4-\d{6}'-only regex. A real anchor's text is
 * a filename, e.g. "S4aP260071_QCOM.docx" -- parseSA4DocumentId_() SEARCHES
 * within it (not an exact-match), so the identifier is found regardless of
 * the surrounding "_QCOM"/extension. That suffix is never discarded: it
 * survives in `fileName` (and in `url`) alongside the canonical `tdocId` --
 * matching against a report's own TDoc elsewhere always compares `tdocId`
 * values by exact string equality, never filenames or substrings.
 *
 * Returns null (never a partially-filled object) when the anchor's text
 * contains no recognized SA4 document identifier -- an unrecognized/
 * unsupported family is never guessed into existence, matching this
 * project's existing "do not guess" convention.
 */
function parseDraftAnchor_(anchor, baseUrl) {
  const fileName = String(anchor && anchor.text || '').trim();
  if (!fileName) return null;
  const parsed = parseSA4DocumentId_(fileName);
  if (!parsed.isValid) return null;
  return {
    tdocId: parsed.raw,
    fileName: fileName,
    url: toAbsoluteUrl_(baseUrl, anchor && anchor.href)
  };
}

/**
 * PERF-006 (Part B): pre-parses every revision-folder anchor ONCE into a
 * Map<canonicalTdocId, DraftEntry[]> bucket, instead of updateRevisions_()
 * re-running parseDraftAnchor_() (and the parseSA4DocumentId_() regex
 * match it does internally) once per (table, anchor) pair.
 * parseDraftAnchor_(a, baseUrl) depends only on the anchor and the
 * (per-run-constant) baseUrl -- never on which TDoc table is currently
 * being processed -- so recomputing it per table was pure redundant work:
 * for a 34-table/65-anchor workload that is up to 2,210 parse calls
 * collapsed down to exactly 65. PERF-005 measured the OLD nested-loop
 * matching stage at only ~9ms for that same workload, so this is a
 * complexity/code-quality improvement (same class of fix as PERF-003
 * Part A's TDoc table index), not a runtime-bottleneck fix -- do not
 * expect a large wall-clock improvement from this alone.
 *
 * Uses the SAME parseDraftAnchor_() every other caller uses (never a
 * second filename/TDoc parser), and preserves each anchor's original
 * relative order within its TDoc's bucket -- iterating `anchors` in the
 * same order the old anchors.forEach() loop did, so
 * updateRevisions_()'s per-table dedup/ordering logic downstream (which
 * only depends on the SET and relative order of matched drafts, both
 * unchanged here) sees byte-identical input to before.
 */
function buildRevisionAnchorIndex_(anchors, baseUrl) {
  const index = new Map();
  (anchors || []).forEach(a => {
    const draft = parseDraftAnchor_(a, baseUrl);
    if (!draft) return;
    if (!index.has(draft.tdocId)) index.set(draft.tdocId, []);
    index.get(draft.tdocId).push(draft);
  });
  return index;
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
function buildRevisionMapFromTdocList_(cfg, groups, context) {
  cfg = cfg || getReportConfig_(context);
  groups = groups || downloadAndGroupTdocs_(cfg, context);
  const map = {};

  Object.keys(groups).forEach(key => {
    groups[key].tdocs.forEach(td => {
      const from = tdocNumberOf_(td);
      const to = getRevisedTo_(td);
      if (from && to && from !== to) map[from] = to;
    });
  });

  getReportStateStore_(context).setProperty('REVISION_MAP', JSON.stringify(map));
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
 * PERF-003 (Part A): builds a canonical-identity index of every TDoc table
 * currently in the document from a SINGLE table scan --
 * Map<canonicalTdocId, Table>. Uses the SAME central SA4 TDoc registry
 * (parseExactSA4DocumentId_) every other "is this a TDoc table, what's its
 * number" consumer in this file already uses -- not a second, parallel
 * identification scheme. A table whose own TDoc cell doesn't parse as a
 * valid SA4 identifier is skipped, never indexed under a guessed/partial
 * key. `tables` is expected to already be a `getTablesCounted_()` result
 * (this function does not fetch the document itself, so building the
 * index never costs an extra scan beyond whichever scan already produced
 * `tables`).
 */
function buildTdocTableIndex_(tables) {
  const index = new Map();
  tables.forEach(table => {
    if (!isTDocTable_(table)) return;
    const raw = String(safeCellText_(table, 0, 1) || '').trim();
    const parsed = parseExactSA4DocumentId_(raw);
    if (!parsed.isValid) return;
    index.set(parsed.raw, table);
  });
  return index;
}

/**
 * PERF-003 (Part A): O(1) canonical-identity lookup against a
 * buildTdocTableIndex_() result. Returns null (never throws, never
 * guesses) for an input that doesn't parse as a valid SA4 identifier or
 * that simply isn't in the index.
 */
function lookupTdocTableInIndex_(index, tdocNumber) {
  const parsed = parseExactSA4DocumentId_(tdocNumber);
  if (!parsed.isValid) return null;
  return index.get(parsed.raw) || null;
}

/**
 * Find the TDOC table for a given document number.
 *
 * PERF-003 (Part A): `index`, if supplied (a buildTdocTableIndex_()
 * result), is used for an O(1) lookup instead of a fresh full-document
 * table scan -- this is what lets findParentRevisedToTable_()/
 * rearrangeRevisionTables_() avoid repeating body.getTables() once per
 * TDoc/revision-pair. Omitting `index` preserves the exact original
 * behavior (a fresh scan every call) for any caller that hasn't been
 * updated to build/pass one.
 */
function findTdocTable_(body, tdocNumber, index) {
  const want = String(tdocNumber || '').trim().toUpperCase();
  if (!want) return null;

  if (index) return lookupTdocTableInIndex_(index, tdocNumber);

  const tables = getTablesCounted_(body, 'findTdocTable_');
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
/**
 * PERF-003 (Part A): now returns the NEWLY INSERTED copy (truthy) when a
 * move happened, or `null` (falsy, same as the old `false`) when the
 * table was already in place -- existing `if (moveTableAfter_(...))`
 * call sites branch identically either way. The return value lets a
 * caller maintaining a table-identity index (e.g. rearrangeRevisionTables_())
 * update that index to point at the new table object, since the ORIGINAL
 * one this function removes via removeFromParent() becomes a stale/
 * dangling reference the instant this function returns.
 */
function moveTableAfter_(body, table, afterTable) {
  const targetIndex = body.getChildIndex(afterTable) + 1;
  const currentIndex = body.getChildIndex(table);
  if (currentIndex === targetIndex) return null; // already in the right place

  // Insert a copy at the destination, then drop the original. Holding the
  // element reference means the shifting indices do not matter.
  const copy = table.copy();
  body.insertTable(targetIndex, copy);
  table.removeFromParent();
  return copy;
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
 *
 * PERF-003 (Part A): `index`, if supplied, is used for both findTdocTable_()
 * lookups per revision pair instead of a fresh full-table scan each time.
 * When moveTableAfter_() actually moves a table, the ORIGINAL table
 * object it removed becomes a stale reference -- the index entry for the
 * moved (child) TDoc is updated in place to point at the new copy, so a
 * LATER revision pair in the same call that also touches this TDoc still
 * gets a live, correct reference rather than a dangling one.
 */
function rearrangeRevisionTables_(cfg, groups, index, context) {
  cfg = cfg || getReportConfig_(context);
  const body = getActiveDocumentBodyCounted_('rearrangeRevisionTables_', context);
  const map = buildRevisionMapFromTdocList_(cfg, groups, context);
  perfCount_('revision pairs processed (rearrangeRevisionTables_)', Object.keys(map).length);

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

    const parentTable = perfTimedAccum_('  rearrangeRevisionTables_: findTdocTable_ (accumulated)', () => findTdocTable_(body, parent, index));
    if (!parentTable) { stats.missingParent++; return; }

    if (setDispositionRevisedTo_(parentTable, child)) stats.dispositions++;

    const childTable = perfTimedAccum_('  rearrangeRevisionTables_: findTdocTable_ (accumulated)', () => findTdocTable_(body, child, index));
    if (!childTable) { stats.missingChild++; return; }

    const movedTable = perfTimedAccum_('  rearrangeRevisionTables_: moveTableAfter_ (accumulated)', () => moveTableAfter_(body, childTable, parentTable));
    if (movedTable) {
      if (index) {
        const parsedChild = parseExactSA4DocumentId_(child);
        if (parsedChild.isValid) index.set(parsedChild.raw, movedTable);
      }
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
    // PERF-001: safeFetch_() is the shared choke point for RSS, A1
    // archive, and Revisions-folder fetches -- one counter/timer here
    // covers all three call sites without touching each individually.
    perfCount_('UrlFetchApp.fetch() calls (total)');
    perfCount_('fetch() call site: safeFetch_ (' + (label || 'unlabeled') + ')');
    const fetchStart = Date.now();
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    perfAddTime_('  safeFetch_ total UrlFetchApp.fetch time', Date.now() - fetchStart);
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
function getDeadlineExtensionMap_(context) {
  const body = getReportBody_(context);
  const props = getReportStateStore_(context);
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
  // TEMPLATE-002A (port of Legacy BUGFIX-LEGACY-003): the archive of the
  // list the collector resolved (the same one RSS reads); the SA4 default
  // only when none is given.
  const listName = String(list || '').trim() || LIST_NAME_LOCK;
  const urls = [];
  const now = new Date();
  const start = new Date(now.getTime() - daysBack * 24 * 3600 * 1000);
  const weeks = ['A', 'B', 'C', 'D', 'E'];
  let cur = new Date(start.getFullYear(), start.getMonth(), 1);
  const endMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  while (cur.getTime() <= endMonth.getTime()) {
    const yy = String(cur.getFullYear()).slice(-2);
    const mm = String(cur.getMonth() + 1).padStart(2, '0');
    weeks.forEach(w => urls.push(`https://list.etsi.org/scripts/wa.exe?A1=ind${yy}${mm}${w}&L=${encodeURIComponent(listName)}`));
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
function isA1CachedEmpty_(cfg, url, context) {
  const props = getReportStateStore_(context);
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
function markA1Empty_(cfg, url, empty, context) {
  getReportStateStore_(context).setProperty(a1CacheKey_(url), JSON.stringify({ ts: Date.now(), empty: !!empty }));
}
function clearA1EmptyCache_() {
  const props = PropertiesService.getDocumentProperties();
  props.getKeys().forEach(k => { if (k.startsWith('A1_EMPTY_CACHE_')) props.deleteProperty(k); });
  Logger.log('A1 empty-week cache cleared');
}

// =========================================================
// PERF-006 (Part A): Reviewer API "no summary" negative cache
// =========================================================
// PERF-005 found that fetchAndAddAbstract_() retries the exact same
// Reviewer API request every automatic run for a TDoc whose summary is
// permanently unavailable (a 404 response persists nothing, so the next
// run's addAbstractsForTables_() sweep -- which only skips a table that
// already HAS an Abstract cell -- always finds this one still missing and
// tries again). This mirrors the existing A1 empty-page cache above:
// same {ts, ...} + TTL shape, same Document Properties storage, same
// conservative 24h default -- but keyed by canonical TDoc identity (via
// the central parseExactSA4DocumentId_() registry, never a second parser)
// rather than by URL, and deliberately narrower in what it caches (see
// fetchAndAddAbstract_()'s own comment for exactly which response is
// treated as "definitive no summary" vs. left uncached).
var REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS = 24;

function reviewerNoSummaryCacheKey_(canonicalTdocId) {
  return 'REVIEWER_NO_SUMMARY_CACHE_' + canonicalTdocId;
}

function isReviewerNoSummaryCached_(canonicalTdocId, context) {
  const raw = getReportStateStore_(context).getProperty(reviewerNoSummaryCacheKey_(canonicalTdocId));
  if (!raw) return false;
  try {
    const obj = JSON.parse(raw);
    const ttlMs = REVIEWER_NO_SUMMARY_CACHE_TTL_HOURS * 3600 * 1000;
    if ((Date.now() - obj.ts) > ttlMs) return false;
    return true;
  } catch (e) { return false; }
}

function markReviewerNoSummary_(canonicalTdocId, statusCode, context) {
  getReportStateStore_(context).setProperty(
    reviewerNoSummaryCacheKey_(canonicalTdocId),
    JSON.stringify({ ts: Date.now(), statusCode: statusCode })
  );
}

function clearReviewerNoSummaryCache_(canonicalTdocId, context) {
  getReportStateStore_(context).deleteProperty(reviewerNoSummaryCacheKey_(canonicalTdocId));
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
/**
 * Re-applies this report's compact table style (zero row height, zero
 * paragraph spacing, fixed 2-column TDoc widths, header shading, font
 * normalization) across every table in the document, plus a stray-
 * empty-paragraph cleanup in the e-mail discussion section. It exists
 * because Apps Script's own table/row/cell/paragraph insertion APIs
 * (insertTable/appendTableRow/appendTableCell/insertParagraph) all default
 * to non-zero spacing/height -- every newly INSERTED table or paragraph
 * needs this pass to look consistent with the rest of the report; content
 * already formatted by a PRIOR run's pass needs it again only if its
 * structure changed (a row/cell/table was added), never merely because
 * its TEXT changed (status/e-mail/revision cell rewrites reuse the same
 * existing rows/cells).
 *
 * PERF-003 (Part B): measured as ~16s of the ~73s no-change baseline
 * (largely thousands of individual per-row/per-paragraph writes across
 * the WHOLE document, unconditionally, every run) -- see
 * continuousUpdate()'s own call site for the structural-change gate that
 * now skips this call specifically for a no-change incremental update.
 * This function itself, and every OTHER call site (full report build,
 * "Update All", manual formatting menu items), is completely unchanged
 * and still calls it unconditionally, exactly as before.
 */
/**
 * PERF-003 (Part B) / PERF-003B (Part 3): pure decision -- whether
 * continuousUpdate()'s unconditional, full-document formatting pass
 * (removeRowHeightAndSpacing()) is actually needed this run. True whenever
 * ANY table was inserted/moved, OR a row was inserted into an existing
 * table, this run (new TDocs, a moved revision table, a newly-inserted
 * revision-linked table, or an abstract row added to an existing table);
 * false when this run only rewrote EXISTING cell TEXT (status updates,
 * e-mail discussion, revision cell rendering), which never adds/removes a
 * row/cell/table and so never needs reformatting.
 *
 * abstractRowsInserted was added in PERF-003B after the call-graph audit
 * found that continuousUpdate() -> addAbstractsForTables_() ->
 * fetchAndAddAbstract_() inserts a table ROW into an EXISTING table
 * (table.insertTableRow(), see fetchAndAddAbstract_()) independently of
 * the other three signals -- exactly the PERF-002 measured run (0 new
 * TDocs, 0 moved revisions, 1 abstract row inserted), which the original
 * three-signal version of this function would have wrongly classified as
 * "no structural change" and skipped formatting for.
 */
function shouldReformatAfterUpdate_(newTdocsAdded, revisionsMoved, revisionLinkedTablesInserted, abstractRowsInserted) {
  return (newTdocsAdded || 0) > 0 || (revisionsMoved || 0) > 0 || (revisionLinkedTablesInserted || 0) > 0 || (abstractRowsInserted || 0) > 0;
}

function removeRowHeightAndSpacing(context) {
  const body = getReportBody_(context);
  const tables = getTablesCounted_(body, 'removeRowHeightAndSpacing');

  // Apply widths ONLY for 2-column TDOC tables (and leave other tables untouched)
  // PERF-001: this ALWAYS runs, unconditionally, on every continuousUpdate()
  // call, regardless of whether anything actually changed -- suspected
  // major cost driver (see this task's report). Timed as its own stage.
  perfTimed_('  formatting: setTwoColumnTDocTableWidths_', () => setTwoColumnTDocTableWidths_(context));

  const formattingWriteLoopStart = Date.now();
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];

    // Row formatting + paragraph spacing cleanup
    for (let j = 0; j < table.getNumRows(); j++) {
      const row = table.getRow(j);
      row.setMinimumHeight(0);
      perfCount_('DocumentApp writes: row.setMinimumHeight()');

      for (let k = 0; k < row.getNumCells(); k++) {
        const cell = row.getCell(k);
        const numChildren = cell.getNumChildren();

        for (let l = 0; l < numChildren; l++) {
          const child = cell.getChild(l);
          if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = child.asParagraph();
            paragraph.setSpacingBefore(0);
            paragraph.setSpacingAfter(0);
            perfCount_('DocumentApp writes: paragraph.setSpacingBefore/After()', 2);
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
  perfAddTime_('  formatting: per-table row/cell/paragraph write loop', Date.now() - formattingWriteLoopStart);

  perfTimed_('  formatting: removeEmptyParagraphs_ (own full-document paragraph scan)', () => removeEmptyParagraphs_(context));
}

/**
 * Only set widths for TDOC tables that have exactly 2 columns.
 * Leaves all other tables alone.
 */
function setTwoColumnTDocTableWidths_(context) {
  const cfg = getConfig_(context);

  const pageWidth = parseInt(cfg.TDOC_PAGE_USABLE_WIDTH || '468', 10);
  const firstColWidth = parseInt(cfg.TDOC_COL1_WIDTH || '108', 10);
  const secondColWidth = pageWidth - firstColWidth;

  const body = getActiveDocumentBodyCounted_('setTwoColumnTDocTableWidths_', context);
  const tables = getTablesCounted_(body, 'setTwoColumnTDocTableWidths_');

  tables.forEach(t => {
    if (!isTDocTable_(t)) return;

    // only apply to exactly 2 columns
    const numCols = t.getRow(0).getNumCells();
    if (numCols !== 2) return;

    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      row.getCell(0).setWidth(firstColWidth);
      row.getCell(1).setWidth(secondColWidth);
      perfCount_('DocumentApp writes: cell.setWidth()', 2);
    }
  });
}

/**
 * Removes empty paragraphs only between "E-Mail Discussion" and "Revisions:"
 * (same logic you had).
 */
function removeEmptyParagraphs_(context) {
  const body = getActiveDocumentBodyCounted_('removeEmptyParagraphs_', context);
  const paragraphs = body.getParagraphs();
  perfCount_('body.getParagraphs() calls (total)');
  perfCount_('body.getParagraphs() call site: removeEmptyParagraphs_');
  perfCount_('paragraphs scanned (removeEmptyParagraphs_)', paragraphs.length);

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

function getConfig_(context) {
  const body = getReportBody_(context);
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
 *
 * ADDON-007B1: the normal dialog shows Meeting / Report setup / Options;
 * implementation and override fields live under a collapsed "Advanced"
 * section. Nothing is presented from a hard-coded fresh-document default
 * (no SA4#136 folder/number, no fallback portal ID, no pre-selected 6G):
 * an unset value renders blank. The runtime fallbacks in getReportConfig_()
 * are untouched -- this only changes what the dialog shows and saves.
 */
function configureMeetingSettings() {
  const ui = DocumentApp.getUi();
  // TEMPLATE-002B: refused in the master template; in a report created from
  // it, the creator's setup information is stored first (once).
  if (templateRuntimeRelease_()) {
    assertNotTemplateMaster_();
    ensureReportBootstrapped_();
  }
  const props = PropertiesService.getDocumentProperties();

  const currentMeetingFolder = props.getProperty('MEETING_FOLDER') || '';
  const currentMeetingNumber = props.getProperty('MEETING_NUMBER') || '';
  const currentSuffix = props.getProperty('REPORT_SUFFIX') || '';
  const currentAgendaSourceDocId = props.getProperty('AGENDA_SOURCE_DOC_ID') || '';
  const currentTdocUrl = props.getProperty('TDOC_LIST_URL') || '';
  const currentShowPreview = props.getProperty('SHOW_PREVIEW_SNIPPET') !== 'false';
  const tokenConfigured = Boolean(PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN'));

  // TEMPLATE-002B (decision 2026-10-01): Email Collection Start Date, in the
  // template runtime only. A stored value is shown as it is; without one the
  // meeting start date is proposed (and replaced by the resolved start date
  // on Resolve, unless the user has typed a date). Outside the template
  // runtime the field is not rendered, and the dialog script, which only
  // looks the field up by id, finds nothing.
  const storedEmailStartDate = String(props.getProperty('EMAIL_START_DATE') || '').trim();
  const emailStartDateExplicit = isValidCollectorStartDate_(storedEmailStartDate);
  const emailStartDateValue = emailStartDateExplicit ? storedEmailStartDate : meetingStartDateIso_(props.getProperty('MEETING_DATE'));
  const emailStartDateFieldHtml = !templateRuntimeRelease_() ? '' :
    '<label>Email Collection Start Date:</label>' +
    '<input type="date" id="collectionStartField" name="emailStartDate" data-config-key="emailStartDate"' +
    ' data-explicit="' + (emailStartDateExplicit ? 'true' : 'false') + '" value="' + emailStartDateValue + '" oninput="this.touched = true">' +
    '<div class="hint">' + (emailStartDateExplicit
      ? 'Saved value. E-mails before this date are not collected.'
      : 'Default: the meeting start date. E-mails before this date are not collected; move it earlier or later if needed.') + '</div>';

  // ARCH-010: the SAME pure merge function the "Resolve" button's server
  // call uses (computeResolvedMeetingPreview_()), called here with
  // resolverResult=null -- since no resolution has happened yet on dialog
  // open, every field falls straight through to its existing Document
  // Property value. This is what makes an existing, legacy-configured
  // document (no MEETING_ID resolution ever performed) open pre-filled
  // correctly without forcing the user through Resolve first.
  const existingForPreview = {
    MEETING_ID: props.getProperty('MEETING_ID'),
    MEETING_TYPE: props.getProperty('MEETING_TYPE'),
    MEETING_NAME: props.getProperty('MEETING_NAME'),
    MEETING_DATE: props.getProperty('MEETING_DATE'),
    FTP_BASE: props.getProperty('FTP_BASE'),
    AGENDA_TDOC: props.getProperty('AGENDA_TDOC'),
    AGENDA_CSV_URL: props.getProperty('AGENDA_CSV_URL'),
    MAILING_LIST: props.getProperty('MAILING_LIST'),
    TDOC_LIST_URL: props.getProperty('TDOC_LIST_URL'),
    REVISIONS_URL: props.getProperty('REVISIONS_URL')
  };
  const initialPreview = computeResolvedMeetingPreview_(existingForPreview, null);
  const initiallyAdhoc = String(initialPreview.meetingType.value || '').trim().toLowerCase() === 'adhoc';
  const familyInfo = buildReportFamilyInfo_();
  // ADDON-007B3: the canonical rules and the evaluator's own source are
  // handed to the browser, so the live status uses exactly the server rules.
  const readinessRules = MEETING_READINESS_RULES_;
  const readinessEvaluatorSource = evaluateMeetingReadiness_.toString();

  function esc(v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function sourceLabel(source) {
    if (source === 'resolved') return '<span class="badge badge-resolved">found automatically</span>';
    if (source === 'existing') return '<span class="badge badge-existing">saved value</span>';
    if (source === 'candidate') return '<span class="badge badge-candidate">suggested — please review</span>';
    return '<span class="badge badge-unresolved">needs input</span>';
  }
  const familyOptions = Object.keys(familyInfo).map(function (family) {
    return '<option value="' + esc(family) + '"' + (currentSuffix === family ? ' selected' : '') + '>' + esc(familyInfo[family].label) + '</option>';
  }).join('');

  // Build HTML form
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      label { display: block; margin-top: 15px; font-weight: bold; }
      input, select { width: 100%; padding: 8px; margin-top: 5px; box-sizing: border-box; }
      input[type=checkbox] { width: auto; }
      button { margin-top: 20px; padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      button:disabled { background: #999; cursor: default; }
      .hint { font-size: 11px; color: #666; margin-top: 3px; }
      .section { background: #f5f5f5; padding: 10px; margin: 15px 0; border-left: 3px solid #4285f4; }
      .resolve-row { display: flex; gap: 8px; align-items: flex-start; }
      .resolve-row input { flex: 1; }
      .resolve-row button { margin-top: 5px; white-space: nowrap; }
      .badge { display: inline-block; font-size: 10px; font-weight: normal; padding: 2px 6px; border-radius: 3px; margin-left: 6px; vertical-align: middle; }
      .badge-resolved { background: #d4edda; color: #155724; }
      .badge-existing { background: #e2e3e5; color: #383d41; }
      .badge-candidate { background: #cce5ff; color: #004085; }
      .badge-unresolved { background: #fff3cd; color: #856404; }
      #resolveStatus { font-size: 12px; margin-top: 6px; }
      #resolveStatus.error { color: #a94442; }
      #resolveStatus.busy { color: #666; }
      #resolveStatus.ok { color: #2d7d2d; }
      #meetingSummary { margin-top: 10px; font-size: 13px; font-weight: bold; }
      #agendaCandidates { margin-top: 8px; font-size: 12px; }
      #readinessStatus { margin-top: 15px; padding: 10px; border-left: 3px solid #999; background: #f5f5f5; font-size: 13px; }
      #readinessStatus.ready { border-left-color: #2d7d2d; background: #eaf6ea; }
      #readinessStatus.attention { border-left-color: #c77c00; background: #fff6e0; }
      #readinessStatus .readiness-title { font-weight: bold; }
      #readinessStatus ul { margin: 6px 0 0 0; padding-left: 18px; }
      details.advanced { margin-top: 20px; }
      details.advanced summary { cursor: pointer; font-weight: bold; color: #4285f4; padding: 8px 0; }
    </style>

    <h2>Meeting Configuration</h2>

    <div class="section">
      <h3>1. Meeting</h3>
      <label>3GPP Portal Meeting ID:</label>
      <div class="resolve-row">
        <input type="text" id="meetingId" value="${esc(initialPreview.meetingId.value)}" placeholder="Portal meeting ID (digits)" oninput="updateDependentUi()">
        <button type="button" id="resolveBtn" onclick="resolveMeeting()">Resolve</button>
      </div>
      <div class="hint">Find the ID in the meeting's 3GPP portal address (...MtgId=NNNNN). Resolving reads public meeting details and does not save anything.</div>
      <div id="resolveStatus"></div>
      <div id="meetingSummary">${esc(initialPreview.summary)}</div>
    </div>

    <div class="section">
      <h3>2. Report setup</h3>

      <label>Report Family:</label>
      <select id="reportType" onchange="onFamilyChanged()">
        <option value=""${currentSuffix === '' ? ' selected' : ''}>Select report family...</option>
        ${familyOptions}
      </select>
      <div class="hint" id="familyStatus"></div>
      <div class="hint" id="agendaStructure"></div>

      <label>Agenda TDoc: <span id="agendaTdocBadge">${sourceLabel(initialPreview.agendaTdoc.source)}</span></label>
      <div class="resolve-row">
        <input type="text" id="agendaTdoc" value="${esc(initialPreview.agendaTdoc.value)}" placeholder="Agenda TDoc number (or enter manually)" oninput="updateDependentUi()">
        <button type="button" id="discoverBtn" onclick="discoverAgendaTdocs()">Discover Agenda / TDocs</button>
      </div>
      <div class="hint">Finds the agenda document and TDocs in the meeting's document list. This can take a moment.</div>
      <div id="discoverStatus"></div>
      <div id="agendaCandidates"></div>
      <div class="hint" id="agendaSourceStatus">${esc(initialPreview.agendaSourceStatus)}</div>
      <input type="hidden" id="agendaCsvUrl" value="${esc(initialPreview.agendaCsvUrl.value)}">

      <label>TDoc List URL:</label>
      <input type="text" id="tdocUrl" value="${esc(currentTdocUrl)}" placeholder="https://.../TDoc_List....xlsx" oninput="updateDependentUi()">
      <div class="hint" id="tdocUrlHint"></div>

      <label>Mailing List:</label>
      <input type="text" id="mailingList" value="${esc(initialPreview.mailingList.value)}" placeholder="Mailing list name" oninput="onMailingListEdited()">
      <div class="hint"><span id="mailingListHint"></span> <a href="#" id="mailingListReset" style="display:none" onclick="resetMailingList(); return false;">Use the default</a></div>

      <label>Discussion E-mail Sender:</label>
      <input type="email" id="discussionEmailSender" value="${esc(props.getProperty('DISCUSSION_EMAIL_SENDER') || '')}" placeholder="e.g. reporter@example.com">
      <div class="hint">The "From" address of generated discussion e-mails (Prepare TDoc Discussion E-mails); they are addressed To the Mailing List's reflector. Never derived from your Google account or from Mailing List -- leave blank and Discussion Export will refuse to generate until this is set.</div>

      <label>Revisions / Drafts URL: <span id="revisionsUrlBadge">${sourceLabel(initialPreview.revisionsUrl.source)}</span></label>
      <input type="text" id="revisionsUrl" value="${esc(initialPreview.revisionsUrl.value)}" placeholder="https://www.3gpp.org/ftp/.../inbox/drafts/" oninput="updateDependentUi()">
      <div class="hint">Suggested drafts/revisions folder. Review if needed.</div>
    </div>

    <div class="section">
      <h3>3. Options</h3>
      ${emailStartDateFieldHtml}
      <label>
        <input type="checkbox" id="showPreview" ${currentShowPreview ? 'checked' : ''}>
        Show email preview snippets in report
      </label>
      <div class="hint" id="tokenStatus">${tokenConfigured ? 'Reviewer API token configured' : 'No Reviewer API token configured'}</div>
    </div>

    <details class="advanced">
      <summary>Advanced</summary>

      <div class="section">
        <label>FTP Base: <span id="ftpBaseBadge">${sourceLabel(initialPreview.ftpBase.source)}</span></label>
        <input type="text" id="ftpBase" value="${esc(initialPreview.ftpBase.value)}" placeholder="https://www.3gpp.org/ftp/.../Docs/" oninput="updateDependentUi()">
        <div class="hint">Location of the meeting's documents. Found automatically when the meeting is resolved.</div>

        <label>Meeting Report Template (Google Doc):</label>
        <input type="text" id="agendaSourceDocId" value="${esc(currentAgendaSourceDocId)}" placeholder="Leave empty to use the shared default template">
        <div class="hint">Google Doc ID or URL of the template used for the report's opening and IPR sections.</div>

        <label>Meeting Type (adhoc / main):</label>
        <input type="text" id="meetingType" value="${esc(initialPreview.meetingType.value)}" placeholder="adhoc or main" oninput="updateDependentUi()">
        <div class="hint" id="portalTypeHint"></div>

        <label>Meeting Name:</label>
        <input type="text" id="meetingName" value="${esc(initialPreview.meetingName.value)}" placeholder="Meeting name" oninput="updateDependentUi()">

        <label>Meeting Date:</label>
        <input type="text" id="meetingDate" value="${esc(initialPreview.meetingDate.value)}" placeholder="September 22, 2026">
        <div class="hint" id="dateRangeHint"></div>

        <div id="mainMeetingFields"${initiallyAdhoc ? ' style="display:none"' : ''}>
          <label>Meeting Folder (main meetings):</label>
          <input type="text" id="meetingFolder" value="${esc(currentMeetingFolder)}" placeholder="Folder name on the 3GPP FTP" oninput="updateDependentUi()">
          <div class="hint">Folder name on the 3GPP FTP for a main meeting.</div>

          <label>Meeting Number (main meetings):</label>
          <input type="text" id="meetingNumber" value="${esc(currentMeetingNumber)}" placeholder="Meeting number" oninput="updateDependentUi()">
          <div class="hint">Meeting number used in the TDoc list filename of a main meeting.</div>
        </div>

        <label>Replace Reviewer API Token:</label>
        <input type="password" id="apiToken" value="" placeholder="Enter a new token to replace the current one" autocomplete="off">
        <label>
          <input type="checkbox" id="clearApiToken">
          Remove the saved Reviewer API token
        </label>
        <div class="hint">Used to fetch AI summaries and abstracts. The saved token is never displayed.</div>
      </div>
    </details>

    <input type="hidden" id="familyInfo" value="${esc(JSON.stringify(familyInfo))}">
    <input type="hidden" id="readinessRules" value="${esc(JSON.stringify(readinessRules))}">

    <div id="readinessStatus" aria-live="polite"></div>

    <button onclick="saveConfig()">Save Configuration</button>
    <button onclick="google.script.host.close()" style="background: #666;">Cancel</button>

    <script>
      ${readinessEvaluatorSource}

      function sourceBadgeHtml(source) {
        if (source === 'resolved') return '<span class="badge badge-resolved">found automatically</span>';
        if (source === 'existing') return '<span class="badge badge-existing">saved value</span>';
        if (source === 'candidate') return '<span class="badge badge-candidate">suggested \\u2014 please review</span>';
        return '<span class="badge badge-unresolved">needs input</span>';
      }

      function setFieldWithBadge(inputId, labelId, fieldPreview) {
        document.getElementById(inputId).value = fieldPreview.value;
        const badgeHost = document.getElementById(labelId);
        if (badgeHost) badgeHost.innerHTML = sourceBadgeHtml(fieldPreview.source);
      }

      function readFamilyInfo() {
        try {
          return JSON.parse(document.getElementById('familyInfo').value) || {};
        } catch (e) {
          return {};
        }
      }

      // ADDON-007B2: what the user has decided versus what was derived.
      // An explicit family choice (or a saved one) is never replaced by an
      // automatic suggestion; a mailing list is either the family's derived
      // default or an explicit override the user typed / saved earlier.
      let familyUserChosen = !!document.getElementById('reportType').value;
      let familyAutoApplied = null;
      let familyInference = null;
      let mailingListOverridden = !!document.getElementById('mailingList').value;
      let mailingListTouched = false;
      // ADDON-008A: the TDoc list URL Discover filled in (if any) and its status.
      let tdocListDiscovered = null;
      let tdocListStatusText = '';

      function familyLabel(family) {
        const entry = readFamilyInfo()[family];
        return entry ? entry.label : family;
      }

      function familySourceText(inf) {
        const parts = [];
        if (inf.sources.indexOf('meeting-name') !== -1) parts.push('meeting name');
        if (inf.sources.indexOf('tdoc-family') !== -1) parts.push((inf.tdocPrefixes || []).join('/') + ' documents');
        return parts.join(' and ');
      }

      function familyStatusText() {
        const inf = familyInference;
        if (!inf || !inf.applicable) return '';
        const family = document.getElementById('reportType').value;
        if (familyUserChosen) {
          if (family && inf.family && inf.family !== family) {
            return familyLabel(family) + ' \\u2014 selected (meeting information suggests ' + familyLabel(inf.family) + ')';
          }
          return '';
        }
        if (family && family === familyAutoApplied && inf.family === family) {
          return familyLabel(family) + (inf.confidence === 'confident'
            ? ' \\u2014 confirmed by ' + familySourceText(inf)
            : ' \\u2014 suggested from ' + familySourceText(inf));
        }
        if (inf.confidence === 'conflict') return 'Needs input \\u2014 conflicting family information found';
        return 'Needs input \\u2014 family could not be determined';
      }

      // Called with the server's inference after Resolve / Discover.
      function applyFamilyInference(inf) {
        familyInference = inf || null;
        const select = document.getElementById('reportType');
        if (!familyUserChosen) {
          if (inf && inf.applicable && inf.family) {
            select.value = inf.family;
            familyAutoApplied = inf.family;
          } else if (familyAutoApplied) {
            // earlier suggestion no longer holds (conflict / new meeting): never silently switch
            select.value = '';
            familyAutoApplied = null;
          }
        }
      }

      // The select's own change handler: a manual choice always wins.
      function onFamilyChanged() {
        familyUserChosen = !!document.getElementById('reportType').value;
        familyAutoApplied = null;
        updateDependentUi();
      }

      function onMailingListEdited() {
        mailingListOverridden = true;
        mailingListTouched = true;
        updateDependentUi();
      }

      function resetMailingList() {
        mailingListOverridden = false;
        mailingListTouched = true;
        updateDependentUi();
      }

      function readReadinessRules() {
        try {
          return JSON.parse(document.getElementById('readinessRules').value) || [];
        } catch (e) {
          return [];
        }
      }

      function readinessFields() {
        function val(id) {
          const el = document.getElementById(id);
          return el ? el.value : '';
        }
        return {
          meetingType: val('meetingType'), meetingId: val('meetingId'), meetingName: val('meetingName'),
          ftpBase: val('ftpBase'), reportFamily: val('reportType'), agendaTdoc: val('agendaTdoc'),
          agendaCsvUrl: val('agendaCsvUrl'), tdocListUrl: val('tdocUrl'), mailingList: val('mailingList'),
          meetingFolder: val('meetingFolder'), meetingNumber: val('meetingNumber')
        };
      }

      function escapeHtml(text) {
        return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }

      // Same rules and evaluator as the server (shipped with the dialog);
      // purely local, no server call.
      function updateReadiness() {
        const box = document.getElementById('readinessStatus');
        if (!box) return null;
        const result = evaluateMeetingReadiness_(readinessFields(), readReadinessRules());
        function bullets(list) {
          return '<ul>' + list.map(function (item) { return '<li>' + escapeHtml(item.message) + '</li>'; }).join('') + '</ul>';
        }
        if (result.ready) {
          box.className = 'ready';
          box.innerHTML = '<div class="readiness-title">Ready to build</div>All required meeting sources are configured.' +
            (result.warnings.length ? '<div style="margin-top:6px">Warnings:</div>' + bullets(result.warnings) : '');
        } else {
          box.className = 'attention';
          box.innerHTML = '<div class="readiness-title">Needs attention</div>' + bullets(result.issues);
        }
        return result;
      }

      // Family and agenda structure are separate concepts: the family only
      // names the report; the agenda structure depends on the meeting type.
      function updateDependentUi() {
        const type = String(document.getElementById('meetingType').value || '').trim().toLowerCase();
        const meetingId = String(document.getElementById('meetingId').value || '').trim();
        const family = document.getElementById('reportType').value;
        const info = readFamilyInfo();
        const familyEntry = family ? info[family] : null;

        let structure;
        if (type === 'adhoc') {
          structure = 'Ad-hoc \\u2014 use complete discovered agenda';
        } else if (!type && !meetingId) {
          structure = 'Resolve the meeting to see how the agenda is used.';
        } else if (familyEntry) {
          structure = 'Main meeting \\u2014 ' + familyEntry.label + ' section ' + familyEntry.section + 'x';
        } else {
          structure = 'Main meeting \\u2014 select a report family.';
        }
        document.getElementById('agendaStructure').textContent = structure;

        // ADDON-008A: name the discovered meeting-specific list as such.
        const tdocUrlValue = String(document.getElementById('tdocUrl').value || '').trim();
        let tdocHint;
        if (type !== 'adhoc') {
          tdocHint = 'Main meetings: leave empty to find the TDoc list automatically. Ad-hoc meetings: paste the TDoc list URL.';
        } else if (tdocListDiscovered && tdocUrlValue === tdocListDiscovered) {
          tdocHint = 'Meeting-specific Portal document list (found automatically).';
        } else if (tdocUrlValue) {
          tdocHint = 'TDoc list URL for this meeting.';
        } else {
          tdocHint = (tdocListStatusText ? tdocListStatusText + ' ' : 'The TDoc list could not be determined automatically for this meeting. ') +
            'Paste the meeting\\'s TDoc list URL.';
        }
        document.getElementById('tdocUrlHint').textContent = tdocHint;

        document.getElementById('familyStatus').textContent = familyStatusText();

        const mailingInput = document.getElementById('mailingList');
        const resetLink = document.getElementById('mailingListReset');
        if (!mailingListOverridden) mailingInput.value = familyEntry ? familyEntry.mailingList : '';
        if (mailingListOverridden) {
          document.getElementById('mailingListHint').textContent = familyEntry
            ? 'Custom value \\u2014 overrides the ' + familyEntry.label + ' default (' + familyEntry.mailingList + ').'
            : 'Custom value.';
          resetLink.style.display = '';
        } else {
          document.getElementById('mailingListHint').textContent = familyEntry
            ? 'Default for ' + familyEntry.label + ' (' + familyEntry.mailingList + ') \\u2014 derived from the report family. Edit to override.'
            : 'Used to collect meeting-related email. Derived from the report family \\u2014 select a family first.';
          resetLink.style.display = 'none';
        }

        updateReadiness();

        document.getElementById('mainMeetingFields').style.display = type === 'adhoc' ? 'none' : '';
      }

      function applyPreview(preview) {
        // TEMPLATE-002B: the template runtime's start-date field (rendered
        // only there). Without a saved or typed value it follows the
        // resolved meeting start date.
        const startField = document.getElementById('collectionStartField');
        if (startField && typeof startField.getAttribute === 'function' && startField.getAttribute('data-explicit') !== 'true' &&
            !startField.touched && /^[0-9]{4}-[0-9]{2}-[0-9]{2}/.test(String(preview.startDateRaw || ''))) {
          startField.value = String(preview.startDateRaw).slice(0, 10);
        }
        setFieldWithBadge('meetingName', 'meetingNameBadge', preview.meetingName);
        setFieldWithBadge('meetingType', 'meetingTypeBadge', preview.meetingType);
        setFieldWithBadge('meetingDate', 'meetingDateBadge', preview.meetingDate);
        setFieldWithBadge('ftpBase', 'ftpBaseBadge', preview.ftpBase);
        setFieldWithBadge('agendaTdoc', 'agendaTdocBadge', preview.agendaTdoc);
        // a saved value is an explicit override; otherwise the list is derived from the family
        if (!mailingListTouched && preview.mailingList.source === 'existing') {
          document.getElementById('mailingList').value = preview.mailingList.value;
          mailingListOverridden = true;
        }
        setFieldWithBadge('revisionsUrl', 'revisionsUrlBadge', preview.revisionsUrl);

        document.getElementById('meetingSummary').textContent = preview.summary || '';

        const portalHint = document.getElementById('portalTypeHint');
        if (preview.portalType && preview.meetingType.source !== 'resolved') {
          portalHint.textContent = 'Unrecognized meeting type code "' + preview.portalType + '" \\u2014 please confirm adhoc or main manually.';
        } else {
          portalHint.textContent = '';
        }

        const dateHint = document.getElementById('dateRangeHint');
        if (preview.startDateRaw || preview.endDateRaw) {
          dateHint.textContent = 'Meeting dates: ' + (preview.startDateRaw || '?') + ' \\u2192 ' + (preview.endDateRaw || '?');
        } else {
          dateHint.textContent = '';
        }

        // ADDON-008A: a discovered meeting-specific TDoc list only fills an
        // empty field -- a typed or saved URL is never replaced.
        const tdocInput = document.getElementById('tdocUrl');
        if (preview.tdocListUrl && preview.tdocListUrl.source === 'resolved' && tdocInput && !String(tdocInput.value || '').trim()) {
          tdocInput.value = preview.tdocListUrl.value;
          tdocListDiscovered = preview.tdocListUrl.value;
        }
        tdocListStatusText = preview.tdocListStatus || '';
        const csvInput = document.getElementById('agendaCsvUrl');
        if (csvInput && preview.agendaCsvUrl) csvInput.value = preview.agendaCsvUrl.value;
        const agendaSourceEl = document.getElementById('agendaSourceStatus');
        if (agendaSourceEl) agendaSourceEl.textContent = preview.agendaSourceStatus || '';

        const candBox = document.getElementById('agendaCandidates');
        if (preview.agendaCandidates && preview.agendaCandidates.length > 0) {
          candBox.textContent = 'Several possible agenda documents were found \\u2014 please choose one and enter it above: ' + preview.agendaCandidates.join(', ');
        } else {
          candBox.textContent = '';
        }

        applyFamilyInference(preview.familyInference);
        updateDependentUi();
      }

      // PROD-016: holds the CORE resolve result (GetMeetings only, no
      // TdocList.aspx) so "Discover Agenda / TDocs" can enrich it without
      // re-fetching GetMeetings. Cleared whenever Resolve is re-run or the
      // Meeting ID field changes, so a stale enrichment can never be
      // merged against a different meeting's core result.
      let lastResolvedCore = null;

      function resolveMeeting() {
        const meetingId = document.getElementById('meetingId').value;
        lastResolvedCore = null;
        const statusEl = document.getElementById('resolveStatus');
        const btn = document.getElementById('resolveBtn');
        const discoverStatusEl = document.getElementById('discoverStatus');
        discoverStatusEl.className = '';
        discoverStatusEl.textContent = '';
        btn.disabled = true;
        statusEl.className = 'busy';
        statusEl.textContent = 'Resolving from 3GPP\\u2026';
        // POST-MEETING-001 (Task A5): manual-diagnostic stage markers only
        // (browser DevTools console, never required for normal use) --
        // pairs with resolveMeetingForConfigDialog_()'s own Logger.log
        // stage markers so "response returned to client" / "client
        // success handler entered" / "client preview render completed"
        // can be distinguished from a genuine server-side stall.
        console.log('resolveMeeting(): client invoking google.script.run.resolveMeetingForConfigDialog');
        google.script.run
          .withSuccessHandler(function(result) {
            console.log('resolveMeeting(): response returned to client, success handler entered');
            btn.disabled = false;
            if (!result.ok) {
              statusEl.className = 'error';
              statusEl.textContent = '\\u274C ' + result.error;
              applyPreview(result.preview);
              console.log('resolveMeeting(): client preview render completed (ok:false path)');
              return;
            }
            lastResolvedCore = result.resolved;
            applyPreview(result.preview);
            console.log('resolveMeeting(): client preview render completed (ok:true path)');
            if (result.resolved.warnings && result.resolved.warnings.length > 0) {
              statusEl.className = 'error';
              statusEl.textContent = '\\u26A0\\uFE0F ' + result.resolved.warnings.join(' | ');
            } else {
              statusEl.className = 'ok';
              statusEl.textContent = '\\u2705 Meeting found. Use "Discover Agenda / TDocs" to find the agenda, then review and Save.';
            }
          })
          .withFailureHandler(function(error) {
            console.log('resolveMeeting(): response returned to client, failure handler entered: ' + error);
            btn.disabled = false;
            statusEl.className = 'error';
            statusEl.textContent = '\\u274C Resolve failed: ' + error;
            // Dialog stays open -- resolution failure never closes it or
            // saves anything.
          })
          .resolveMeetingForConfigDialog(meetingId);
      }

      // PROD-016: separate, explicit enrichment action -- fetches
      // TdocList.aspx (agenda TDoc, TDoc family). Requires a successful
      // Resolve first (needs lastResolvedCore); never runs automatically
      // after Resolve, and its own slowness/failure can never affect the
      // already-resolved core meeting fields above.
      function discoverAgendaTdocs() {
        const meetingId = document.getElementById('meetingId').value;
        const statusEl = document.getElementById('discoverStatus');
        const btn = document.getElementById('discoverBtn');
        if (!lastResolvedCore) {
          statusEl.className = 'error';
          statusEl.textContent = '\\u274C Resolve the meeting first.';
          return;
        }
        btn.disabled = true;
        statusEl.className = 'busy';
        statusEl.textContent = 'Fetching and scanning the TDoc list\\u2026';
        google.script.run
          .withSuccessHandler(function(result) {
            btn.disabled = false;
            if (!result.ok) {
              statusEl.className = 'error';
              statusEl.textContent = '\\u274C ' + result.error;
              return;
            }
            applyPreview(result.preview);
            if (result.resolved.warnings && result.resolved.warnings.length > 0) {
              statusEl.className = 'error';
              statusEl.textContent = '\\u26A0\\uFE0F ' + result.resolved.warnings.join(' | ');
            } else {
              statusEl.className = 'ok';
              statusEl.textContent = '\\u2705 Agenda/TDocs discovered.';
            }
          })
          .withFailureHandler(function(error) {
            btn.disabled = false;
            statusEl.className = 'error';
            statusEl.textContent = '\\u274C Discovery failed: ' + error + ' \\u2014 the meeting details above are unaffected.';
          })
          .discoverAgendaForConfigDialog(meetingId, lastResolvedCore);
      }

      function saveConfig() {
        const newToken = document.getElementById('apiToken').value;
        let apiTokenAction = 'keep';
        if (document.getElementById('clearApiToken').checked) {
          apiTokenAction = 'clear';
        } else if (newToken && newToken.trim()) {
          apiTokenAction = 'replace';
        }
        const config = {
          meetingFolder: document.getElementById('meetingFolder').value,
          meetingNumber: document.getElementById('meetingNumber').value,
          meetingId: document.getElementById('meetingId').value,
          meetingType: document.getElementById('meetingType').value,
          meetingName: document.getElementById('meetingName').value,
          meetingDate: document.getElementById('meetingDate').value,
          ftpBase: document.getElementById('ftpBase').value,
          mailingList: mailingListOverridden ? document.getElementById('mailingList').value.trim() : '',
          mailingListMode: mailingListOverridden && document.getElementById('mailingList').value.trim() ? 'override' : 'derived',
          discussionEmailSender: document.getElementById('discussionEmailSender') ? document.getElementById('discussionEmailSender').value : '',
          revisionsUrl: document.getElementById('revisionsUrl').value,
          reportType: document.getElementById('reportType').value,
          agendaSourceDocId: document.getElementById('agendaSourceDocId').value,
          agendaTdoc: document.getElementById('agendaTdoc').value,
          agendaCsvUrl: document.getElementById('agendaCsvUrl') ? document.getElementById('agendaCsvUrl').value : '',
          tdocUrl: document.getElementById('tdocUrl').value,
          showPreview: document.getElementById('showPreview').checked,
          apiTokenAction: apiTokenAction,
          apiToken: apiTokenAction === 'replace' ? newToken : ''
        };
        // TEMPLATE-002B: see applyPreview(); the field names its own config key.
        const startFieldToSave = document.getElementById('collectionStartField');
        if (startFieldToSave && typeof startFieldToSave.getAttribute === 'function') {
          config[startFieldToSave.getAttribute('data-config-key')] = startFieldToSave.value;
        }
        google.script.run
          .withSuccessHandler(() => {
            // Saved is not the same as ready to build.
            const readiness = updateReadiness();
            alert('\\u2705 Configuration saved.' + (readiness && !readiness.ready
              ? '\\n\\n\\u26A0\\uFE0F It is not ready to build yet:\\n' + readiness.issues.map(function (issue) { return '\\u2022 ' + issue.message; }).join('\\n')
              : ''));
            google.script.host.close();
          })
          .withFailureHandler((error) => {
            alert('\\u274C Error saving configuration: ' + error);
          })
          .saveConfigurationSettings(config);
      }

      updateDependentUi();
    </script>
  `)
  .setWidth(600)
  .setHeight(820);

  ui.showModalDialog(html, 'Meeting Configuration');
}


function saveConfigurationSettings(config) {
  assertNotTemplateMaster_(); // TEMPLATE-002B
  // ADDON-007B1: Save writes Document Properties (unchanged, and what every
  // interactive build reads). For a document registered with the central
  // add-on the background scheduler reads a SEPARATE central copy, so the
  // same save is mirrored into it -- under the add-on ScriptLock, before
  // anything is written, so a lock failure leaves both copies untouched.
  // An unregistered (legacy/bound) document never touches central state.
  const documentId = getActiveDocumentIdSafely_();
  const registered = documentId ? getRegisteredReportDocument_(documentId) : null;

  if (!registered) {
    persistConfigurationSettings_(config);
    return;
  }

  withAddonScriptLock_(function () {
    persistConfigurationSettings_(config);
    syncConfigurationDialogStateToCentral_(documentId);
  });
}

// Shared template Google Doc used by the runtime when no template is
// configured (same value as getReportConfig_()'s own fallback).
const DEFAULT_MEETING_REPORT_TEMPLATE_DOC_ID_ = '1qP--dusvUhNwwBtMEH4xVdxaP1c6L1hZ49geICoYV2s';

// Every Document Property the Meeting Configuration dialog can write --
// all of them are also adoption keys (ADDON003_ADOPTION_FIXED_KEYS_).
var CONFIG_DIALOG_MANAGED_KEYS_ = [
  'MEETING_FOLDER', 'MEETING_NUMBER', 'MEETING_ID',
  'REPORT_SUFFIX', 'AGENDA_SOURCE_DOC_ID', 'AGENDA_TDOC', 'AGENDA_CSV_URL', 'TDOC_LIST_URL',
  'SHOW_PREVIEW_SNIPPET',
  'MEETING_TYPE', 'MEETING_NAME', 'MEETING_DATE', 'FTP_BASE', 'MAILING_LIST', 'REVISIONS_URL'
];

function getActiveDocumentIdSafely_() {
  try {
    return DocumentApp.getActiveDocument().getId() || null;
  } catch (e) {
    return null;
  }
}

function setPropertyIfPresent_(store, key, value) {
  const v = String(value === null || value === undefined ? '' : value).trim();
  if (v) store.setProperty(key, v);
}

/**
 * Secrets never reach a log line: any token/secret/password-like field is
 * replaced by a marker before the config object is serialized.
 */
function redactConfigForLog_(config) {
  const safe = {};
  Object.keys(config || {}).forEach(function (key) {
    if (/token|secret|password|apikey/i.test(key) && key !== 'apiTokenAction') {
      safe[key] = config[key] ? '[redacted]' : '';
    } else {
      safe[key] = config[key];
    }
  });
  return safe;
}

function persistConfigurationSettings_(config) {
  const docProps = PropertiesService.getDocumentProperties();
  const scriptProps = PropertiesService.getScriptProperties();

  // TEMPLATE-002B (decision 2026-10-01): Email Collection Start Date is a
  // normal Configure Meeting field in the template runtime. Checked with the
  // collector's own rule BEFORE anything is written; a blank value never
  // erases a stored one. Outside the template runtime the dialog has no such
  // field and the value is not written (ADDON-007B1, unchanged).
  const emailStartDate = templateRuntimeRelease_() ? String(config.emailStartDate || '').trim() : '';
  if (emailStartDate && !isValidCollectorStartDate_(emailStartDate)) {
    throw new Error('Email Collection Start Date must be a date in the form YYYY-MM-DD.');
  }

  const submittedType = String(config.meetingType || '').trim().toLowerCase();
  const effectiveType = submittedType || String(docProps.getProperty('MEETING_TYPE') || '').trim().toLowerCase();

  // Meeting identity. Only an explicit MAIN meeting keeps the historical
  // "blank means the built-in default" write; for an ad-hoc meeting or a
  // not-yet-resolved fresh document a blank value is simply not written,
  // so no SA4#136 folder/number or fallback portal ID is ever persisted.
  if (effectiveType === 'main') {
    docProps.setProperty('MEETING_FOLDER', config.meetingFolder || 'TSGS4_136_Montreal');
    docProps.setProperty('MEETING_NUMBER', config.meetingNumber || '136');
    docProps.setProperty('MEETING_ID', config.meetingId || '60777');
  } else {
    setPropertyIfPresent_(docProps, 'MEETING_FOLDER', config.meetingFolder);
    setPropertyIfPresent_(docProps, 'MEETING_NUMBER', config.meetingNumber);
    setPropertyIfPresent_(docProps, 'MEETING_ID', config.meetingId);
  }

  // Report family is a user choice: never invent one (the runtime's own
  // 6G fallback in getReportConfig_() is untouched).
  setPropertyIfPresent_(docProps, 'REPORT_SUFFIX', config.reportType);

  const submittedTemplate = extractGoogleDocId_(config.agendaSourceDocId || '');
  if (submittedTemplate) {
    docProps.setProperty('AGENDA_SOURCE_DOC_ID', submittedTemplate);
  } else if (!docProps.getProperty('AGENDA_SOURCE_DOC_ID')) {
    docProps.setProperty('AGENDA_SOURCE_DOC_ID', DEFAULT_MEETING_REPORT_TEMPLATE_DOC_ID_);
  }

  // ARCH-010: skip-if-blank -- an existing manually-configured AGENDA_TDOC
  // is never erased by a blank submitted value (e.g. the resolver found no
  // agenda TDoc).
  if (config.agendaTdoc && config.agendaTdoc.trim()) {
    docProps.setProperty('AGENDA_TDOC', config.agendaTdoc.trim());
  }

  // Save file locations (only if provided, otherwise auto-detect)
  if (config.tdocUrl && config.tdocUrl.trim()) {
    docProps.setProperty('TDOC_LIST_URL', config.tdocUrl.trim());
  } else {
    docProps.deleteProperty('TDOC_LIST_URL'); // Will auto-detect
  }

  // Save options
  docProps.setProperty('SHOW_PREVIEW_SNIPPET', config.showPreview ? 'true' : 'false');

  // Reviewer API token (Script Properties). 'keep' leaves it alone,
  // 'replace' stores a new value, 'clear' removes it. A caller that sends
  // only a non-blank apiToken (no action) is treated as 'replace'.
  const tokenAction = config.apiTokenAction || 'replace';
  if (tokenAction === 'clear') {
    scriptProps.deleteProperty('REVIEWER_API_TOKEN');
  } else if (tokenAction === 'replace' && config.apiToken && config.apiToken.trim()) {
    scriptProps.setProperty('REVIEWER_API_TOKEN', config.apiToken.trim());
  }
  resetReviewerTokenRunState_(); // TEMPLATE-002C: the token may just have changed

  // ARCH-010: resolver-driven fields, skip-if-blank -- a blank submitted
  // value means "nothing new to say", never "erase what was configured".
  if (config.meetingType && config.meetingType.trim()) {
    docProps.setProperty('MEETING_TYPE', config.meetingType.trim());
  }
  if (config.meetingName && config.meetingName.trim()) {
    docProps.setProperty('MEETING_NAME', config.meetingName.trim());
  }
  if (config.meetingDate && config.meetingDate.trim()) {
    docProps.setProperty('MEETING_DATE', config.meetingDate.trim());
  }
  if (config.ftpBase && config.ftpBase.trim()) {
    docProps.setProperty('FTP_BASE', config.ftpBase.trim());
  }
  // ADDON-008A: the agenda.csv source is accepted only as the exact
  // candidate of this ad-hoc meeting's own series folder -- never an
  // arbitrary client-supplied URL. Skip-if-blank, like AGENDA_TDOC; the
  // build re-validates the CSV content before using it.
  const agendaCsvUrl = String(config.agendaCsvUrl || '').trim();
  if (agendaCsvUrl) {
    if (effectiveType === 'adhoc' && isAdhocAgendaCsvUrlFor_(agendaCsvUrl, docProps.getProperty('FTP_BASE'))) {
      docProps.setProperty('AGENDA_CSV_URL', agendaCsvUrl);
    } else {
      Logger.log('Ignored an agenda.csv URL that does not belong to this ad-hoc meeting: ' + agendaCsvUrl);
    }
  }
  // ADDON-007B2: an explicit dialog mode separates a user override from the
  // family-derived default. 'derived' removes the property so normal
  // derivation resumes; 'override' stores it. A caller sending no mode keeps
  // the historical skip-if-blank behaviour.
  const mailingListValue = String(config.mailingList || '').trim();
  if (config.mailingListMode === 'derived' || (config.mailingListMode === 'override' && !mailingListValue)) {
    docProps.deleteProperty('MAILING_LIST');
  } else if (mailingListValue) {
    docProps.setProperty('MAILING_LIST', mailingListValue);
  }
  // ARCH-012: same skip-if-blank protection for REVISIONS_URL.
  if (config.revisionsUrl && config.revisionsUrl.trim()) {
    docProps.setProperty('REVISIONS_URL', config.revisionsUrl.trim());
  }
  // ADDON-009 (ported from LEGACY-UPGRADE-006H): the discussion e-mail
  // sender, an independent value (never the Mailing List). Same
  // skip-if-blank protection -- a blank value never erases a saved sender.
  // Read only by the interactive exporter, so it is not a central-state key.
  if (config.discussionEmailSender && config.discussionEmailSender.trim()) {
    docProps.setProperty('DISCUSSION_EMAIL_SENDER', config.discussionEmailSender.trim());
  }

  if (emailStartDate) docProps.setProperty('EMAIL_START_DATE', emailStartDate);

  Logger.log('Configuration saved: ' + JSON.stringify(redactConfigForLog_(config)));
}

/**
 * Mirrors the dialog-managed Document Properties into the document's
 * central state (same 'SA4_STATE|<documentId>|<key>' namespace adoption
 * uses). Unlike adoption -- which copies only keys that are SET -- this
 * also removes a central key whose Document Property is absent, so a
 * cleared override (e.g. TDOC_LIST_URL) cannot survive centrally. Only the
 * dialog-managed keys are touched; caches and other state are left alone.
 * Document Properties are never modified here.
 */
function syncConfigurationDialogStateToCentral_(documentId) {
  const source = PropertiesService.getDocumentProperties();
  const target = getReportStateStore_({ documentId: documentId, mode: 'addon-background' });

  CONFIG_DIALOG_MANAGED_KEYS_.forEach(function (key) {
    const value = source.getProperty(key);
    if (value === null) {
      target.deleteProperty(key);
    } else {
      target.setProperty(key, value);
    }
  });

  const mismatches = CONFIG_DIALOG_MANAGED_KEYS_.filter(function (key) {
    return target.getProperty(key) !== source.getProperty(key);
  });
  if (mismatches.length > 0) {
    throw new Error('Your settings were saved, but could not be applied to the automatic updates (' + mismatches.join(', ') + '). Please try Save again.');
  }

  const cfg = getReportConfig_({ mode: 'bound' });
  registerReportDocument_(documentId, {
    meetingId: cfg.MEETING_ID,
    meetingName: registryMeetingName_(source, cfg)
  });
}

function registryMeetingName_(documentProperties, cfg) {
  return documentProperties.getProperty('MEETING_NAME') || cfg.MEETING_FOLDER;
}

const REPORT_FAMILY_LABELS_ = {
  '6G': '6G',
  'Audio': 'Audio',
  'Video': 'Video',
  'RTC': 'RTC',
  'MBS': 'MBS',
  'Liaison': 'Liaison',
  'New': 'New Work Items'
};

/**
 * Per report family: display label, the MAIN-meeting agenda section and the
 * default mailing list -- both taken from the existing lookups, so the
 * dialog never carries its own copy of them.
 */
function buildReportFamilyInfo_() {
  const info = {};
  Object.keys(REPORT_FAMILY_LABELS_).forEach(function (family) {
    info[family] = {
      label: REPORT_FAMILY_LABELS_[family],
      section: getAgendaPrefixForReportType_(family),
      mailingList: MAILING_LISTS[family] || LIST_NAME_LOCK
    };
  });
  return info;
}

// Report families that are never inferred from words in a meeting name or
// path ("New"/"Liaison" are ordinary words, not SWG identifiers).
const REPORT_FAMILIES_NOT_NAME_INFERRED_ = ['Liaison', 'New'];

/**
 * ADDON-007B2: splits free text into lower-case alphanumeric tokens
 * (underscores, dashes and slashes separate), so "FS_6G_MED" yields
 * fs / 6g / med and "Audio" never matches inside a longer word.
 */
function tokenizeFamilyEvidenceText_(text) {
  return String(text || '').toLowerCase().split(/[^a-z0-9]+/).filter(function (t) { return t; });
}

/** Report families whose own name appears as a whole token in `text`. */
function reportFamiliesNamedIn_(text) {
  const tokens = tokenizeFamilyEvidenceText_(text);
  return Object.keys(REPORT_FAMILY_LABELS_).filter(function (family) {
    return REPORT_FAMILIES_NOT_NAME_INFERRED_.indexOf(family) === -1 && tokens.indexOf(family.toLowerCase()) !== -1;
  });
}

/**
 * ADDON-007B2: pure, generic report-family inference for AD-HOC meetings.
 *
 * evidence: { meetingType, meetingName, ftpBase, tdocFamilyKeys }, where
 * tdocFamilyKeys is an array (or an object keyed by) SA4_TDOC_FAMILIES keys
 * seen in the meeting's TDoc list.
 *
 * Strong evidence: whole-word family tokens in the meeting name, and the
 * report family of the TDoc series seen (SA4_TDOC_FAMILIES.reportFamily --
 * null for S4aP, so plenary documents can never decide anything). The FTP
 * path only ever confirms a family the strong evidence already chose.
 *
 * confidence: 'confident' (name and TDocs agree), 'suggested' (one
 * unambiguous source), 'unresolved' (nothing usable) or 'conflict'
 * (incompatible evidence -- nothing is chosen). 'none' with
 * applicable:false for a non-ad-hoc meeting: one main meeting holds several
 * report families, so the family is never inferred from the meeting itself.
 */
function inferReportFamily_(evidence) {
  const e = evidence || {};
  const result = { applicable: true, family: null, confidence: 'unresolved', sources: [], tdocPrefixes: [], reason: null };
  if (String(e.meetingType || '').trim().toLowerCase() !== 'adhoc') {
    result.applicable = false;
    result.confidence = 'none';
    result.reason = 'not-adhoc';
    return result;
  }

  const nameFamilies = reportFamiliesNamedIn_(e.meetingName);

  const keys = Array.isArray(e.tdocFamilyKeys) ? e.tdocFamilyKeys : Object.keys(e.tdocFamilyKeys || {});
  const tdocFamilies = [];
  keys.forEach(function (key) {
    SA4_TDOC_FAMILIES.forEach(function (entry) {
      if (entry.key === key && entry.reportFamily) {
        if (tdocFamilies.indexOf(entry.reportFamily) === -1) tdocFamilies.push(entry.reportFamily);
        if (result.tdocPrefixes.indexOf(entry.prefix) === -1) result.tdocPrefixes.push(entry.prefix);
      }
    });
  });

  if (nameFamilies.length > 1 || tdocFamilies.length > 1 ||
      (nameFamilies.length === 1 && tdocFamilies.length === 1 && nameFamilies[0] !== tdocFamilies[0])) {
    result.confidence = 'conflict';
    result.reason = 'conflicting-evidence';
    return result;
  }

  const family = nameFamilies[0] || tdocFamilies[0] || null;
  if (!family) {
    result.reason = keys.length > 0 ? 'ambiguous-tdoc-series' : 'no-evidence';
    return result;
  }
  result.family = family;
  if (nameFamilies.length) result.sources.push('meeting-name');
  if (tdocFamilies.length) result.sources.push('tdoc-family');
  result.confidence = nameFamilies.length && tdocFamilies.length ? 'confident' : 'suggested';
  const ftpFamilies = reportFamiliesNamedIn_(e.ftpBase);
  if (ftpFamilies.length === 1 && ftpFamilies[0] === family) result.sources.push('ftp-path');
  return result;
}

/**
 * One-line, human-readable summary of the resolved (or saved) meeting,
 * built from an already-computed preview object.
 */
function buildMeetingSummary_(preview) {
  const name = preview.meetingName.value;
  if (!name && !preview.meetingId.value) return 'No meeting resolved yet.';

  const type = String(preview.meetingType.value || '').trim().toLowerCase();
  let typeLabel = '';
  if (type === 'adhoc') typeLabel = 'Ad-hoc meeting';
  else if (type === 'main') typeLabel = 'Main meeting';

  let dates = '';
  const start = computeMeetingDateFromStartDate_(preview.startDateRaw);
  const end = computeMeetingDateFromStartDate_(preview.endDateRaw);
  if (start && end && start !== end) dates = start + ' – ' + end;
  else if (start || end) dates = start || end;
  else dates = preview.meetingDate.value;

  return [name, dates, preview.location, typeLabel].filter(function (part) { return part; }).join(' · ');
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
  // SA4-PROD-006: probes the current meeting's own AGENDA_TDOC when it is
  // a valid, registered SA4 identifier (any family), falling back to the
  // original hardcoded main-meeting probe otherwise. This tests
  // authentication/connectivity only -- it does not prove any specific
  // document's abstract is available in Reviewer, ad-hoc or main.
  results.push('\n2️⃣ Reviewer API:');
  try {
    const token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN');
    if (!token) {
      results.push('   ⚠️  Not configured (abstracts will be skipped)');
    } else {
      const cfg = getReportConfig_();
      const parsedAgendaTdoc = parseExactSA4DocumentId_(cfg.AGENDA_TDOC);
      const probeTdoc = parsedAgendaTdoc.isValid ? parsedAgendaTdoc.raw : 'S4-260001';
      const testUrl = `https://reviewer.bouazizi.dev/api/v1/documents/${probeTdoc}/summary?type=summary`;
      const response = UrlFetchApp.fetch(testUrl, {
        headers: { 'X-API-Key': token },
        muteHttpExceptions: true
      });
      if (response.getResponseCode() === 200 || response.getResponseCode() === 404) {
        results.push(`   ✅ API token valid (probed ${probeTdoc}; connectivity/auth only)`);
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
  // SA4-PROD-002: this used to read getCollectorConfig_().REVISIONS_URL,
  // which is ALWAYS the main-meeting formula (`${INBOX_BASE}Drafts/
  // ${DRAFTS_FOLDER}`, Code.js:649) regardless of meeting.type -- it
  // fabricated and tested a URL for meeting 86178 that was never a real
  // ad-hoc revisions location (observed: HTTP 403). getMeetingContext_()
  // .sources.revisionsUrl is the value that already correctly represents
  // "genuinely absent" for an ad-hoc meeting with no REVISIONS_URL override
  // (undefined -- see resolveMeetingSources_(), unchanged by this task) and
  // is IDENTICAL to the old cfg.REVISIONS_URL value for every main meeting
  // (derivedSources.revisionsUrl: cfg.REVISIONS_URL there), so main-meeting
  // behavior here is unchanged.
  results.push('\n4️⃣ Revisions Folder:');
  try {
    const url = getMeetingContext_().sources.revisionsUrl;
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
 *
 * PERF-003B (Part 2): shares the same document-scoped lock as
 * continuousUpdate() (see its header comment) -- this is a manual menu
 * item that mutates the same document body (collectorUpdate_() +
 * removeRowHeightAndSpacing()), so it must not be allowed to run
 * concurrently with a trigger- or menu-invoked continuousUpdate() either.
 * Neither function calls the other, so there is no nested/double-lock
 * path here -- each acquires the lock itself, once, at its own entry.
 */
function updateReportIncremental() {
  const ui = DocumentApp.getUi();

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    Logger.log('updateReportIncremental(): another incremental update is already in progress -- skipping this run.');
    ui.alert('Busy', 'Another update is currently in progress. Please try again in a moment.', ui.ButtonSet.OK);
    return;
  }

  try {
    // Update email discussions and revisions
    collectorUpdate_();

    // Re-format
    removeRowHeightAndSpacing();

    ui.alert('Success!', 'Report updated successfully!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', 'Failed to update report: ' + e.message, ui.ButtonSet.OK);
    Logger.log('Update error: ' + e.message);
  } finally {
    lock.releaseLock();
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

  // Ad-hoc sessions (stage F): sessions, opening details, attendance and
  // TDoc sessions. No lines -- and nothing read -- unless this is an ad-hoc
  // report with sessions.
  adhocSessionStatusLines_().forEach(line => lines.push(line));
  
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
    // SA4-PROD-006: use the current meeting's own configured agenda TDoc
    // when it is a valid, registered SA4 identifier (any family), so the
    // test at least exercises a real document for THIS meeting; fall back
    // to the original hardcoded main-meeting probe otherwise. Either way
    // this proves authentication/connectivity to the Reviewer API only --
    // a 200/404 here does NOT prove any particular document (ad-hoc or
    // main) actually has a summary available in Reviewer.
    const cfg = getReportConfig_();
    const parsedAgendaTdoc = parseExactSA4DocumentId_(cfg.AGENDA_TDOC);
    const probeTdoc = parsedAgendaTdoc.isValid ? parsedAgendaTdoc.raw : 'S4-260001';
    const testUrl = `https://reviewer.bouazizi.dev/api/v1/documents/${probeTdoc}/summary?type=summary`;
    const response = UrlFetchApp.fetch(testUrl, {
      headers: { 'X-API-Key': token },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code === 200) {
      ui.alert('Reviewer API Test', `✅ API token is valid (probed ${probeTdoc}).\n\nThis confirms authentication/connectivity only -- it does not guarantee every document's abstract is available.`, ui.ButtonSet.OK);
    } else if (code === 404) {
      ui.alert('Reviewer API Test', `✅ API token is valid (probed ${probeTdoc}).\n\n(That document has no summary in Reviewer, but authentication works -- this confirms connectivity only.)`, ui.ButtonSet.OK);
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
function parseAgendaForReport_(cfg, templateDocId, preparedAgendaItems) {
  let agendaItems;
  if (preparedAgendaItems) {
    // ADDON-008A: already validated and projected by
    // prepareAdhocCsvAgendaForBuild_() (ad-hoc agenda.csv source only).
    agendaItems = preparedAgendaItems;
  } else if (cfg.AGENDA_TDOC) {
    Logger.log('Parsing agenda from TDOC ZIP: ' + cfg.AGENDA_TDOC);
    const agendaStructure = downloadMeetingAgenda_(cfg.AGENDA_TDOC, cfg.FTP_BASE);
    // SA4-IMPL-007: canonical ZIP-path filtering migrated from the old
    // inline parent-or-prefix filter to the MeetingContext-driven
    // projection. Note this drops the bare parent item (e.g. "7") that the
    // old filter retained -- SA4-ARCH-007 confirmed buildSkeletonWithTdocTables()
    // immediately, redundantly re-filters that parent back out before doing
    // anything else with the array, so this has no observable effect on the
    // canonical build. The Google-Doc fallback below is deliberately left
    // unmigrated -- see its comment.
    const context = getMeetingContext_();
    agendaItems = projectAgendaItems_(agendaStructure, context.report.agendaSelector);
  } else {
    Logger.log('Parsing agenda from template Google Doc: ' + templateDocId);
    // SA4-IMPL-007: NOT migrated to projectAgendaItems_(). This branch
    // passes a non-empty prefix into parseAgendaStructureWithText_(), which
    // pre-filters at heading/table-parse time via the known-buggy
    // parseAgendaFromHeadings_() (SA4-ARCH-007). Migrating this branch to
    // selector-aware projection would require first obtaining an unfiltered
    // agenda structure from that parser, which means fixing or restructuring
    // the heading-parser bug -- out of scope for this task. Left unchanged.
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
  // TEMPLATE-002B: see configureMeetingSettings().
  if (templateRuntimeRelease_()) {
    assertNotTemplateMaster_();
    ensureReportBootstrapped_();
  }
  // TEMPLATE-003: a template report asks with its own text, which says
  // that the document's content is replaced (ReportCreator.js).
  const response = templateRuntimeRelease_() ? confirmTemplateBuildFromScratch_(ui) : ui.alert(
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

  // TEMPLATE-002C: the phases run without any pop-up in between -- a pop-up
  // suspends the script until it is clicked, inside this execution's time
  // limit.
  const result = runFullReportBuildCore_();

  // 2.17.3: no UI call after the build has changed the document. Live, the
  // final "Success" ui.alert() never appeared: the call hung for ~96 s and
  // threw "Service Documents failed while accessing document ...", turning
  // a completed, saved build into a FAILED execution. A completed build
  // therefore just ends here; its result is in the [FULLBUILD] log lines.
  // A failed phase must stay visible, and needs no UI call for that: the
  // error thrown below is shown by Docs and recorded by the execution.
  if (!result.ok) {
    Logger.log('Full report build failed: ' + result.failedPhase + ': ' + result.error);
    throw new Error(formatFullBuildSummary_(result));
  }
  return result;
}

/**
 * TEMPLATE-002C: the Full Build itself -- no UI at all. Phases run in
 * order; the first one that fails stops the build (later phases are not
 * started) and is reported once in the result (runFullReportBuild() turns
 * that into one thrown error; a completed build ends silently). Each phase logs
 * "[FULLBUILD] <phase>: start" and "[FULLBUILD] <phase>: <status> in N ms",
 * so the last line before a timeout names the phase that was running.
 *
 * Formatting runs ONCE, at the end: removeRowHeightAndSpacing() formats
 * every table of the document from scratch, and the enrichment phases add
 * content (e-mail lines, revision tables, abstract rows) that needs it, so
 * a pass before them is repeated work.
 *
 * Abstracts need the optional Reviewer API token. Without one the phase is
 * skipped as a whole. With one, the tables are built without abstracts and
 * the abstracts phase fetches them, so their cost shows under its own name.
 */
function runFullReportBuildCore_() {
  const totalStart = Date.now();
  const phases = [];
  let failed = null;

  function phase(name, fn) {
    if (failed) return;
    const start = Date.now();
    const entry = { name: name, status: 'done', detail: '', ms: 0 };
    Logger.log('[FULLBUILD] ' + name + ': start');
    try {
      const outcome = fn() || {};
      if (outcome.skipped) entry.status = 'skipped';
      entry.detail = outcome.detail || '';
    } catch (e) {
      entry.status = 'failed';
      entry.detail = e.message;
      failed = entry;
    }
    entry.ms = Date.now() - start;
    phases.push(entry);
    Logger.log('[FULLBUILD] ' + name + ': ' + entry.status + ' in ' + entry.ms + ' ms' + (entry.detail ? ' -- ' + entry.detail : ''));
  }

  resetReviewerTokenRunState_();

  phase('skeleton', function () {
    SKIP_ABSTRACTS_DURING_BUILD_THIS_RUN_ = true;
    try {
      const built = buildSkeletonWithTdocTables({ nonInteractive: true, skipFormatting: true });
      if (!built.ok) throw new Error(built.message);
      return { detail: built.agendaItems + ' agenda items, ' + built.tdocs + ' TDocs' + (built.note || '') };
    } finally {
      SKIP_ABSTRACTS_DURING_BUILD_THIS_RUN_ = false;
    }
  });
  phase('e-mail', function () { collectEmailDiscussionCore_(); });
  phase('revisions', function () { collectRevisionsCore_(); });
  phase('abstracts', function () {
    if (!getReviewerApiTokenForRun_()) return { skipped: true, detail: 'no Reviewer API token configured' };
    const added = addAbstractsForTables_();
    return { detail: added.candidatesProcessed + ' candidate table(s), ' + added.rowsInserted + ' abstract(s) inserted' };
  });
  phase('formatting', function () {
    removeRowHeightAndSpacing();
  });

  const totalMs = Date.now() - totalStart;
  Logger.log('[FULLBUILD] total: ' + totalMs + ' ms -- ' + phases.map(function (p) { return p.name + ' ' + p.status + ' ' + p.ms + ' ms'; }).join(', '));
  return {
    ok: !failed,
    failedPhase: failed ? failed.name : null,
    error: failed ? failed.detail : null,
    phases: phases,
    totalMs: totalMs
  };
}

/** The one dialog text of a Full Build (pure). */
function formatFullBuildSummary_(result) {
  const LABELS = { 'skeleton': 'Skeleton + TDoc tables', 'e-mail': 'E-mail discussion', 'revisions': 'Revisions', 'abstracts': 'Abstracts', 'formatting': 'Formatting' };
  const seconds = function (ms) { return Math.round(ms / 1000) + ' s'; };
  const lines = [result.ok ? 'Full report build completed.' : 'Full report build failed: ' + result.error, ''];
  result.phases.forEach(function (p) {
    lines.push((p.status === 'failed' ? '✗ ' : p.status === 'skipped' ? '– ' : '✓ ') + LABELS[p.name] + ': ' +
      (p.status === 'failed' ? 'FAILED' : p.status === 'skipped' ? 'skipped' : 'done') +
      (p.detail && p.status !== 'failed' ? ' (' + p.detail + ')' : '') + ', ' + seconds(p.ms));
  });
  if (!result.ok) {
    const notRun = Object.keys(LABELS).filter(function (name) { return !result.phases.some(function (p) { return p.name === name; }); });
    if (notRun.length) lines.push('Not run: ' + notRun.map(function (name) { return LABELS[name]; }).join(', ') + '.');
  }
  lines.push('');
  lines.push('Total: ' + seconds(result.totalMs));
  return lines.join('\n');
}

// TEMPLATE-002C: the optional Reviewer API token, read once per execution
// (Apps Script starts every execution with fresh globals). A missing token
// is logged once, not once per TDoc table.
var REVIEWER_TOKEN_RUN_STATE_ = { read: false, token: '', missingLogged: false };
// Set by Full Build while it builds the tables: abstracts get their own phase.
var SKIP_ABSTRACTS_DURING_BUILD_THIS_RUN_ = false;

function getReviewerApiTokenForRun_() {
  if (!REVIEWER_TOKEN_RUN_STATE_.read) {
    REVIEWER_TOKEN_RUN_STATE_.token = PropertiesService.getScriptProperties().getProperty('REVIEWER_API_TOKEN') || '';
    REVIEWER_TOKEN_RUN_STATE_.read = true;
  }
  return REVIEWER_TOKEN_RUN_STATE_.token;
}

/** Forget the cached token: at the start of an operation, and when Configure Meeting changes it. */
function resetReviewerTokenRunState_() {
  REVIEWER_TOKEN_RUN_STATE_.read = false;
  REVIEWER_TOKEN_RUN_STATE_.token = '';
  REVIEWER_TOKEN_RUN_STATE_.missingLogged = false;
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
function buildSkeletonWithTdocTables(options) {
  // TEMPLATE-002C: the menu calls this without arguments (unchanged: its
  // own pop-ups, formatting at the end). Run Full Report Build passes
  // { nonInteractive: true, skipFormatting: true }: no pop-up may suspend
  // the build, and it formats once, after its enrichment phases. The
  // outcome is also returned: { ok, message } (+ agendaItems, tdocs, note).
  const quiet = !!(options && options.nonInteractive);
  // ADDON-007B3: an ad-hoc meeting with missing sources fails here, before
  // the document is cleared, instead of building an empty/wrong report.
  assertMeetingReadyToBuild_();
  const cfg = getReportConfig_();

  if (!cfg.TDOC_LIST_URL) {
    Logger.log('Error: TDOC List URL not configured');
    try {
      if (!quiet) DocumentApp.getUi().alert('Error', 'TDOC List URL not configured. Please run "Configure Meeting" first.', DocumentApp.getUi().ButtonSet.OK);
    } catch (e) {
      // UI not available in this context
    }
    return { ok: false, message: 'TDOC List URL not configured. Please run "Configure Meeting" first.' };
  }

  // ADDON-008A1b: the saved Document Reallocations live in this document,
  // so they are read now, before the document is cleared. They are REPORT
  // overrides: every agenda/TDoc decision below uses the effective
  // allocation (reallocation, else source), and the table is restored at
  // the end of the build.
  const savedReallocations = getReallocationMap_();

  // ADDON-008A: an ad-hoc agenda.csv source is fetched and re-validated
  // here, before the document is cleared, so a failed fallback never leaves
  // an emptied report behind. null = not the agenda source (unchanged path).
  // ADDON-008A1b: its sections follow the effective allocations.
  const preparedAgendaItems = prepareAdhocCsvAgendaForBuild_(cfg, savedReallocations);

  // Step 1: Parse agenda and download TDOCs -- ADDON-008A1b: BEFORE the
  // document is cleared, so the build can still refuse without having
  // changed anything.
  const templateDocId = extractGoogleDocId_(cfg.AGENDA_SOURCE_DOC_ID);
  const agendaItems = parseAgendaForReport_(cfg, templateDocId, preparedAgendaItems);
  if (!agendaItems || agendaItems.length === 0) {
    Logger.log('Warning: No agenda items found for prefix: ' + getConfiguredAgendaPrefix_());
    try {
      if (!quiet) DocumentApp.getUi().alert('Warning', 'No agenda items found for prefix: ' + getConfiguredAgendaPrefix_(), DocumentApp.getUi().ButtonSet.OK);
    } catch (e) {
      // UI not available in this context
    }
    return { ok: false, message: 'No agenda items found for prefix: ' + getConfiguredAgendaPrefix_() };
  }
  
  // Filter out the parent item (e.g., "9 Video SWG") - we only want sub-items (9.1, 9.2, etc.)
  const agendaPrefix = getConfiguredAgendaPrefix_();
  // ADDON-007A: this parent-removal rule is a MAIN-meeting convention (a main
  // agenda's SWG section "9 Video SWG" is a parent of 9.x). For an ad-hoc
  // meeting the real agenda numbering is independent of the report-family
  // prefix, so a real item that happens to be numbered like the family digit
  // must not be dropped. Main-meeting behavior is unchanged.
  const isAdhocForParentFilter = getMeetingContext_().meeting.type === 'adhoc';
  const filteredAgendaItems = agendaItems.filter(item => {
    // Keep items that have a dot after the prefix (e.g., 9.1, 9.2, not just 9)
    return isAdhocForParentFilter || item.number !== agendaPrefix.replace(/\.$/, '');
  });
  
  if (filteredAgendaItems.length === 0) {
    Logger.log('Warning: No sub-agenda items found for prefix: ' + agendaPrefix);
    try {
      if (!quiet) DocumentApp.getUi().alert('Warning', 'No sub-agenda items found for prefix: ' + agendaPrefix, DocumentApp.getUi().ButtonSet.OK);
    } catch (e) {
      // UI not available in this context
    }
    return { ok: false, message: 'No sub-agenda items found for prefix: ' + agendaPrefix };
  }
  // ADDON-008A1b: every saved destination must be usable before anything is
  // cleared. Ad-hoc: it must be an item of this report's agenda (the
  // agenda.csv path already checked it against the whole CSV). Main: any
  // agenda number -- a main-meeting TDoc may move to another SWG's report.
  const reallocationProblems = findReallocationBuildProblems_(savedReallocations, null,
    isAdhocForParentFilter && !preparedAgendaItems ? filteredAgendaItems.map(item => item.number) : null);
  if (reallocationProblems.length > 0) refuseBuildForReallocations_(reallocationProblems);

  const tdocGroups = downloadAndGroupTdocs_(cfg, undefined, savedReallocations);

  // Ad-hoc attendance (stage D): the Company cells typed into the attendee
  // tables are kept before the document is cleared. null, and nothing read
  // or written, unless this ad-hoc report has sessions and imported attendance.
  const attendanceRebuild = beginAdhocAttendanceRebuild_();
  // Ad-hoc sessions (stage B): what the Session column is computed from; null
  // unless this ad-hoc report has sessions.
  const tdocSessions = makeAdhocTdocSessionResolverSafely_();

  // Step 2: Clear document, set title
  const body = DocumentApp.getActiveDocument().getBody().clear();
  setDocumentTitleFromTemplate_(templateDocId);

  const allTdocs = [];
  Object.keys(tdocGroups).forEach(key => {
    tdocGroups[key].tdocs.forEach(td => allTdocs.push({ ...td, agendaItem: key }));
  });

  Logger.log(`Found ${filteredAgendaItems.length} agenda items and ${allTdocs.length} TDOCs`);
  
  // Step 3: Build full agenda structure, backfilling template content
  const sourceBody = DocumentApp.openById(templateDocId).getBody();
  const agendaPrefixNum = (cfg.AGENDA_ITEM_PREFIX || '0.').replace(/\.$/, '');
  const reportType = cfg.REPORT_SUFFIX || '6G';
  // SA4-PROD-001: the nested X.0.1-X.0.4 skeleton below (opening/registration/
  // reallocation/IPR bundled under a synthetic "X.0" parent) is the MAIN
  // 6G-plenary meeting's own template convention. Real evidence from 3GPP
  // meeting 86178 (SA4-e (AH) on FS_6G_MED, a REPORT_SUFFIX='6G' ad-hoc)
  // shows its actual agenda uses the SAME flat X.1 (Opening) / X.2 (IPR)
  // numbering as SWG reports, not X.0.1-X.0.4 -- reusing the nested
  // convention here would duplicate Opening/IPR (once as synthetic X.0.x,
  // once as the real X.1/X.2 items) and misnumber the report. This is an
  // isolated, transitional guard on meeting.type, not a general
  // frontMatterProfile resolution -- report.structureProfile is still
  // 'main-6g' and known-imperfect for this ad-hoc case, left as-is per the
  // same deferral SA4-IMPL-004/007 already established for ULBC-MED. Every
  // existing main-meeting 6G test is unaffected: context.meeting.type is
  // 'main' there, so is6G is unchanged for them.
  const context = getMeetingContext_();
  const is6G = (reportType === '6G') && context.meeting.type !== 'adhoc';

  // SA4-PROD-003: extract, from tdocGroups, every TDoc whose ORIGINAL
  // TDoc-list agenda assignment is before the normal agenda sections begin
  // (isBeforeRegistrationBoundary_(), generic/prefix-driven -- no literal
  // "5"). Gated on !is6G: the 6G-plenary nested convention (is6G branch,
  // just above) uses REAL topic items starting at "{prefix}.1" onward with
  // no reserved Opening/IPR slots, so this boundary concept does not apply
  // there and must not divert real 6G-main topic documents. For the
  // SWG-style branch (every SWG main meeting, and now every ad-hoc 6G-type
  // meeting per SA4-PROD-001), X.1/X.2 are ALWAYS special-cased below
  // (never a plain tdocGroups[...] lookup target), so removing entries
  // here has no effect on any EXISTING main-meeting rendering path -- it
  // only prevents a pre-agenda document from silently vanishing. Rendered
  // under {agendaPrefixNum}.1.4 "Documents" inside the openingSection
  // branch below.
  // ADDON-007A: administrative-section anchors. MAIN meetings: the exact
  // configured-prefix strings used before (byte-identical). ADHOC meetings:
  // derived from the REAL parsed agenda, never from the report-family
  // prefix -- see getAdministrativeAgendaAnchors_().
  const anchors = is6G ? null : getAdministrativeAgendaAnchors_(context, filteredAgendaItems, cfg.AGENDA_ITEM_PREFIX || '0.');

  const registrationDocs = [];
  if (!is6G) {
    Object.keys(tdocGroups).forEach(key => {
      if (isBeforeAdminBoundary_(key, anchors)) {
        registrationDocs.push(...tdocGroups[key].tdocs);
        delete tdocGroups[key];
      }
    });
  }

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
    // ADDON-008A1b: the snapshot read before clearing (re-reading the
    // cleared body here always found nothing).
    const reallocations = savedReallocations;
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
  
  // ADDON-007A: the opening/registration/documents block, emitted at the
  // anchor's emit-after item (the opening item itself for every main meeting
  // and for an ad-hoc agenda with no real opening children).
  const emitOpeningAdminBlock = () => {
    if (context.meeting.type === 'adhoc') {
      // SA4-PROD-007A: generated directly, not copied from the template.
      // No CHAIR_NAME/START_TIME property is introduced -- those remain
      // literal, editable placeholders for the chair to fill in by hand.
      // MEETING_DATE is optional and generic (no meeting-ID-specific
      // value baked in here); when unset, "<meeting date>" is used
      // instead of inventing or assuming any specific date.
      const openingSubSection = anchors.openingSubSection;
      body.appendParagraph(`${openingSubSection} Opening of the session`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      const meetingDateText = (cfg.MEETING_DATE || '').trim() || '<meeting date>';
      // Ad-hoc opening (stage E): with sessions configured the Session
      // administration section says who chaired and when, so this line is
      // not written as well. Every other report gets it as before.
      if (!adhocSessionsEnabled_(context)) {
        body.appendParagraph(`<Chair> opens the session on ${meetingDateText} at <start> CEST.`);
      }
    } else {
      // Opening section for SWG reports - copy X.1 content from template
      const openingHeader = findHeading_(sourceBody, /^X\.1\s+/);
      if (openingHeader) {
        copySectionContentWithReplacement_(openingHeader, body, /^X\.2\s+/, 'X', agendaPrefixNum);
      }
    }

    // Always ensure Registration of Documents section exists with the summary table
    if (!documentContainsHeading_(body, anchors.registrationSection)) {
      body.appendParagraph(`${anchors.registrationSection} Registration of Documents`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    }

    // Always add/update the registered documents summary table
    // Ad-hoc sessions (stage B): with sessions configured it is built with
    // the Session column (tdocSessions is null for every other report).
    if (allTdocs.length > 0) {
      createSummaryTable_(body, allTdocs, allTdocs[0].tdocCol, allTdocs[0].titleCol, allTdocs[0].sourceCol, -1, tdocSessions);
    }

    // SA4-PROD-003: {agendaPrefixNum}.1.4 "Documents" -- collects every
    // TDoc whose ORIGINAL TDoc-list agenda assignment is before the
    // normal agenda sections begin (numerically < "{agendaPrefixNum}.3":
    // e.g. "5", "5.0", "5.1", "5.2" for prefix "5"), extracted from
    // tdocGroups further above via isBeforeRegistrationBoundary_() into
    // `registrationDocs`. Such a document has no real numbered
    // subsection of its own to render under (X.1/X.2 are always
    // special-cased, never a plain TDoc-table target) and would
    // otherwise simply vanish from the report -- exactly the
    // S4aP260089/S4aP260098 problem SA4-PROD-001/002 found. Each row
    // preserves its OWN original agenda-item value (not this loop's
    // item.number) so the existing Document Reallocations workflow
    // (X.1.3, ensureReallocationTable_()/addDocumentReallocation()) can
    // read where it came from. Only created when non-empty: every
    // existing main-meeting report has nothing to put here (X.1/X.2 are
    // never a raw tdocGroups[...] lookup target for main meetings
    // either), so this is a pure no-op there -- no new heading appears.
    if (registrationDocs.length > 0) {
      const documentsSection = anchors.documentsSection;
      if (!documentContainsHeading_(body, documentsSection)) {
        body.appendParagraph(`${documentsSection} Documents`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      }
      orderTdocsByRevision_(registrationDocs).forEach(tdocData => {
        appendTdocDetailTable_(body, tdocData, tdocData.row[tdocData.agendaCol]);
      });
    }
    // ADDON-007A: an ad-hoc agenda with no real IPR item gets the standard
    // IPR content as a synthetic child of the opening item instead of losing it.
    if (anchors.iprSyntheticSection) {
      body.appendParagraph(`${anchors.iprSyntheticSection} IPR and antitrust reminder`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
      appendStandardIprSection_(body, agendaPrefixNum, anchors.iprSyntheticSection, 0);
    }
  };

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
    const openingSection = is6G ? `${agendaPrefixNum}.0.1` : anchors.openingSection;
    const registrationSection = is6G ? `${agendaPrefixNum}.0.2` : anchors.registrationSection;
    const reallocationSection = is6G ? `${agendaPrefixNum}.0.3` : null;
    const iprSection = is6G ? `${agendaPrefixNum}.0.4` : anchors.iprSection;
    
    // Skip 11.0.x sections for 6G as they're already created
    if (is6G && (item.number === openingSection || item.number === registrationSection || 
                 item.number === reallocationSection || item.number === iprSection)) {
      return; // Skip - already created above
    }
    
    // SA4-PROD-002 gated this whole branch off for ad-hoc meetings, on the
    // assumption that a real ad-hoc agenda never wants the X.1.1/X.1.2
    // subsections this branch produces. SA4-PROD-003 corrected that
    // assumption (clarified production requirement): meeting 86178 DOES
    // want the standard X.1.1 Opening / X.1.2 Registration / X.1.3
    // Document Reallocations / X.1.4 Documents structure under its real
    // X.1 item, identical to a main SWG meeting. The branch itself is
    // unconditional again -- only the OPENING CONTENT source differs now
    // (SA4-PROD-007A): a main meeting still copies the shared template's
    // X.1 body (unchanged, byte-identical to before PROD-002); an ad-hoc
    // meeting generates its own minimal X.1.1 content instead, because
    // that shared template's X.1 body is multi-day, main-meeting-specific
    // boilerplate (dated minute-taker assignments, etc) with no meaning
    // for a single ad-hoc call -- copying it produced stale August dates
    // and irrelevant text (SA4-PROD-007 finding). Registration/
    // Reallocation/Documents handling below is completely unchanged and
    // shared by both meeting types.
    if (item.number === openingSection) {
      // ADDON-007A: emitted after this item's subtree (see emitOpeningAdminBlock).
    } else if (item.number === iprSection) {
      // SA4-PROD-008: standard IPR/antitrust/consensus boilerplate,
      // generated directly for EVERY report type (previously: SWG reports
      // copied template content here; 6G/Liaison/New reports fell through
      // to rendering tdocGroups[iprSection] as a plain TDoc table instead).
      // That old TDoc-table fallback is safely removable: since
      // SA4-PROD-007's registration-boundary extraction (isBeforeRegistrationBoundary_,
      // above, gated on !is6G) already removes every tdocGroups entry whose
      // raw agenda item is < "{agendaPrefixNum}.3" -- which "{agendaPrefixNum}.2"
      // (iprSection) always is -- BEFORE this loop runs, tdocGroups[iprSection]
      // is already guaranteed empty by this point for every is6G-false
      // report type. No real TDoc visibility is lost; it was already
      // dead code post-PROD-007. The agenda's own "X.2 ..." heading
      // (already appended above from the real parsed agenda item) remains
      // the section anchor; this only appends its standard child content.
      if (anchors.iprEmitAfter === iprSection) {
        appendStandardIprSection_(body, agendaPrefixNum, anchors.source === 'main-prefix' ? undefined : iprSection, anchors.iprChildOffset);
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

    // ADDON-007A: administrative blocks are emitted at their emit-after item
    // (the anchor itself unless it has real child items, so a main meeting
    // and a childless ad-hoc anchor render exactly where they always did).
    if (!is6G && anchors.openingEmitAfter && item.number === anchors.openingEmitAfter) {
      emitOpeningAdminBlock();
    }
    if (!is6G && anchors.iprSection && anchors.iprEmitAfter !== anchors.iprSection && item.number === anchors.iprEmitAfter) {
      appendStandardIprSection_(body, agendaPrefixNum, anchors.iprSection, anchors.iprChildOffset);
    }
  });
  
  // Add AOB and Close of Session at the end if they weren't in the agenda
  // SA4-PROD-001: real evidence from meeting 86178's actual agenda (...5.10
  // "Other issues", 5.11 "Close of the session", no explicit AOB item) shows
  // a valid 3GPP ad-hoc agenda can end at Close with no separate AOB item.
  // Unconditionally auto-appending a synthetic AOB heading after an
  // already-present Close heading would put "Any other business" AFTER
  // "Close of the session" -- nonsensical. Every existing main-meeting
  // fixture agenda already has both AOB and Close as explicit items
  // (hasAOB is true there), so this added guard never changes their output.
  if (!hasAOB && !hasCloseOfSession) {
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
  
  // ADDON-008A1b: put the saved Document Reallocations back (the 6G branch
  // already wrote them into its own X.0.3 table above), through the same
  // table helpers the Add Document Reallocation dialog uses.
  let reallocationRestoreNote = '';
  if (!is6G && Object.keys(savedReallocations).length > 0) {
    try {
      Object.keys(savedReallocations).forEach(tdoc => {
        const r = savedReallocations[tdoc];
        saveReallocation(tdoc, r.original, r.new, r.reason);
      });
    } catch (e) {
      Logger.log('Could not restore the Document Reallocations table: ' + e.message);
      reallocationRestoreNote = '\n\n⚠️ The Document Reallocations table could not be restored (' + e.message + '). Re-enter: ' +
        Object.keys(savedReallocations).map(t => `${t} → ${savedReallocations[t].new}`).join(', ');
    }
  }

  // Ad-hoc attendance (stage D): the Attendance section is written again
  // from what is stored; a problem is added to the note of the result.
  reallocationRestoreNote += finishAdhocAttendanceRebuild_(body, attendanceRebuild);
  // Ad-hoc opening (stage E): the Session administration section, from the
  // sessions and the stored opening details ('' and nothing done without sessions).
  reallocationRestoreNote += finishAdhocOpeningRebuild_(body);

  // TEMPLATE-002C: Full Build formats once, after its enrichment phases.
  if (!(options && options.skipFormatting)) removeRowHeightAndSpacing();
  Logger.log(`Done: Built skeleton with ${filteredAgendaItems.length} agenda items and ${allTdocs.length} TDOCs.`);
  try {
    if (!quiet) DocumentApp.getUi().alert('Done', `Built skeleton with ${filteredAgendaItems.length} agenda items and ${allTdocs.length} TDOCs.` + reallocationRestoreNote, DocumentApp.getUi().ButtonSet.OK);
  } catch (e) {
    // UI not available in this context
  }
  return {
    ok: true,
    message: `Built skeleton with ${filteredAgendaItems.length} agenda items and ${allTdocs.length} TDOCs.` + reallocationRestoreNote,
    agendaItems: filteredAgendaItems.length,
    tdocs: allTdocs.length,
    note: reallocationRestoreNote
  };
}

/**
 * Download TDOC list and group by agenda item.
 * Returns: { '9.1': { tdocs: [...] }, '9.2': { tdocs: [...] }, ... }
 */
function downloadAndGroupTdocs_(cfg, context, reallocationsSnapshot) {
  const meetingUrl = cfg.TDOC_LIST_URL;
  if (!meetingUrl) return {};

  Logger.log('Downloading TDOC list: ' + meetingUrl);
  perfCount_('UrlFetchApp.fetch() calls (total)');
  perfCount_('fetch() call site: downloadAndGroupTdocs_:TDocListXlsx');
  const response = perfTimed_('  TDoc-list: UrlFetchApp.fetch (download xlsx)', () => UrlFetchApp.fetch(meetingUrl, { muteHttpExceptions: true }));
  if (response.getResponseCode() >= 400) {
    throw new Error('Could not download TDOC list: HTTP ' + response.getResponseCode());
  }

  const blob = response.getBlob();
  blob.setName('TDoc_List_Temp.xlsx');
  const tempFile = perfTimed_('  TDoc-list: DriveApp.createFile (upload temp xlsx)', () => DriveApp.createFile(blob));

  try {
    // Suspected primary cost/variability driver (PERF-001): converting an
    // uploaded .xlsx blob into an openable Google Sheet is a genuinely
    // slow, non-deterministic Apps Script/Drive backend operation --
    // timed separately from the two full-range reads that follow, which
    // read the ENTIRE sheet TWICE (getValues() and getRichTextValues()
    // are two independent full round trips over the same range).
    const spreadsheet = perfTimed_('  TDoc-list: SpreadsheetApp.open (xlsx->Sheet conversion)', () => SpreadsheetApp.open(tempFile));
    const sheet = spreadsheet.getSheets()[0];
    const data = perfTimed_('  TDoc-list: sheet.getDataRange().getValues()', () => sheet.getDataRange().getValues());
    const richTextValues = perfTimed_('  TDoc-list: sheet.getDataRange().getRichTextValues()', () => sheet.getDataRange().getRichTextValues());
    perfCount_('TDoc-list sheet rows read', data.length);

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
    // Ad-hoc sessions (stage B): the upload time of each TDoc. null, and
    // nothing read, unless this is an ad-hoc report with sessions.
    const uploadTimes = readTdocUploadTimesForSessions_(spreadsheet, sheet, headers, data, context);

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
    const agendaSelector = getMeetingContext_(context).report.agendaSelector;
    // ADDON-008A1b: a build passes the reallocations it read BEFORE clearing
    // the document (reading them here would find the cleared body).
    const reallocations = reallocationsSnapshot ||
      perfTimed_('  TDoc-list: getReallocationMap_ (own body.getTables() scan)', () => getReallocationMap_(context));
    const groups = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tdoc = String(row[tdocCol] || '').trim();
      if (!tdoc) continue;

      let agendaItem = String(row[agendaCol] || '').trim();

      if (reallocations[tdoc]) {
        // ADDON-008A1b: shared interpretation -- a removed/withdrawn/n/a TDoc
        // is not placed anywhere in the report (a main meeting's prefix
        // selector already excluded these words; an ad-hoc 'all' selector
        // used to file them under a literal "withdrawn" item).
        const target = interpretReallocationTarget_(reallocations[tdoc].new);
        if (target.kind === 'remove') continue;
        agendaItem = target.kind === 'move' ? target.agendaItem : reallocations[tdoc].new;
      }

      if (!agendaSelectorMatches_(agendaSelector, agendaItem)) continue;

      if (!groups[agendaItem]) groups[agendaItem] = { tdocs: [] };

      groups[agendaItem].tdocs.push({
        row: row,
        richTextRow: richTextValues[i],
        tdocCol, titleCol, sourceCol, contactCol, agendaCol, agendaTopicCol, statusCol, typeCol, forCol,
        revisedToCol
      });
      if (uploadTimes) groups[agendaItem].tdocs[groups[agendaItem].tdocs.length - 1].uploaded = uploadTimes[i];
    }

    Logger.log(`Grouped TDOCs: ${Object.keys(groups).length} agenda items`);
    return groups;

  } finally {
    const trashStart = Date.now();
    try { tempFile.setTrashed(true); } catch (e) { }
    perfAddTime_('  TDoc-list: tempFile.setTrashed (Drive cleanup)', Date.now() - trashStart);
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
 * generateReportTitle_()'s existing {REPORT_SUFFIX, TDOC_LIST_URL,
 * MEETING_ID} contract is unchanged and still drives every main-meeting
 * title exactly as before; the small object below exists only to satisfy
 * that unchanged contract with values sourced from MeetingContext instead
 * of a legacy cfg object.
 *
 * SA4-PROD-002: an ad-hoc meeting has no real portal meeting number
 * (context.meeting.portalId is always null there) and its TDOC_LIST_URL
 * does not follow the main-meeting "SA4%23NNN" filename convention, so the
 * old formula fell back to the literal string "SA4#null" (observed in
 * meeting 86178's real production run). meetingLabel is an additional,
 * OPTIONAL field on generateReportTitle_()'s cfg object -- every existing
 * caller that never sets it (including every main-meeting call) gets
 * byte-identical output to before. Only set here, for ad-hoc meetings,
 * from context.meeting.name -- no meeting number is invented.
 */
function setDocumentTitleFromTemplate_(sourceDocId) {
  try {
    const targetDoc = DocumentApp.getActiveDocument();
    const context = getMeetingContext_();
    const newTitle = generateReportTitle_({
      REPORT_SUFFIX: context.report.type,
      TDOC_LIST_URL: context.sources.tdocListUrl,
      MEETING_ID: context.meeting.portalId,
      meetingLabel: context.meeting.type === 'adhoc' ? context.meeting.name : undefined
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

    // SA4-PROD-002: optional explicit meeting label (e.g. an ad-hoc
    // meeting's real name, "SA4-e (AH) on FS_6G_MED") for meetings with no
    // real portal meeting number and no "SA4%23NNN"-style TDOC_LIST_URL to
    // extract one from. Absent for every existing caller today except the
    // ad-hoc branch of setDocumentTitleFromTemplate_(), so this is a pure
    // addition -- every other caller's output is unchanged.
    if (cfg.meetingLabel) {
      return `${topicName} Minutes – ${cfg.meetingLabel}`;
    }

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
 * SA4-PROD-003: true when `agendaItemValue` (a TDoc-list row's raw,
 * unprocessed "Agenda item" cell) is numerically BEFORE
 * "<agendaPrefixNum>.3" -- i.e. its first dotted component equals
 * agendaPrefixNum and its second component is absent or < 3. Generic and
 * prefix-driven, no literal "5": for prefix "5" this is 5 / 5.0 / 5.1 / 5.2
 * (and any 5.2.x); for prefix "11" it would be 11 / 11.0 / 11.1 / 11.2.
 *
 * Used only to route a document filed under a pre-agenda registration slot
 * into the {agendaPrefixNum}.1.4 "Documents" bucket instead of silently
 * vanishing from the report -- it does NOT change
 * agendaSelectorMatches_()/projectAgendaItems_() or any other canonical
 * agenda-membership decision; both remain untouched by this task.
 */
function isBeforeRegistrationBoundary_(agendaItemValue, agendaPrefixNum) {
  const parts = String(agendaItemValue || '').trim().split('.').map(Number);
  const prefixNum = Number(agendaPrefixNum);
  if (!parts.length || isNaN(parts[0]) || isNaN(prefixNum) || parts[0] !== prefixNum) return false;
  const second = parts.length > 1 ? parts[1] : -1;
  return isNaN(second) ? false : second < 3;
}

/**
 * ADDON-007A: where the report's ADMINISTRATIVE sections live -- the
 * Opening-of-the-session content, the Registration-of-Documents summary,
 * the Document Reallocations slot, the "Documents" bucket and the
 * IPR/antitrust reminder -- as full section numbers.
 *
 * Report family (Audio/Video/6G/...) and meeting structure (main/adhoc)
 * are separate concepts:
 *
 *   MAIN meeting: unchanged. Anchors are the configured report-type
 *   prefix + the fixed SWG layout (X.1 opening, X.1.1/X.1.2/X.1.3/X.1.4
 *   children, X.2 IPR) -- exactly the strings the build always used.
 *
 *   ADHOC meeting: the report-type prefix is NEVER consulted (an ad-hoc
 *   agenda's numbering is independent of it). Anchors come from the REAL
 *   parsed agenda:
 *     - opening: the shallowest (then first) real item whose title reads
 *       as an opening ("Opening ...", "... opening of the session/
 *       meeting"). Fallback (documented, generic): the first real
 *       top-level item -- the administrative block then hangs off the
 *       start of the real agenda.
 *     - IPR: the shallowest (then first) real item mentioning IPR /
 *       antitrust / competition law. If the real agenda has none, the
 *       standard IPR content is generated as a synthetic child of the
 *       opening item (iprSyntheticSection) rather than dropped.
 *   Synthetic children are numbered AFTER any real children of the anchor
 *   item (childOffset), so they never collide with real numbering, and
 *   are emitted after the last real descendant (openingEmitAfter /
 *   iprEmitAfter) so document order matches numeric order. With no real
 *   children the offsets are 0 and emission happens at the anchor itself,
 *   i.e. the same shape as the main layout.
 */
function getAdministrativeAgendaAnchors_(context, agendaItems, configuredPrefix) {
  const isAdhoc = !!(context && context.meeting && context.meeting.type === 'adhoc');

  if (!isAdhoc) {
    const n = String(configuredPrefix || '').replace(/\.$/, '');
    return {
      source: 'main-prefix',
      prefixNum: n,
      openingSection: `${n}.1`,
      openingSubSection: `${n}.1.1`,
      registrationSection: `${n}.1.2`,
      reallocationSection: `${n}.1.3`,
      documentsSection: `${n}.1.4`,
      afterDocumentsSection: `${n}.1.5`,
      iprSection: `${n}.2`,
      iprSyntheticSection: null,
      iprChildOffset: 0,
      openingEmitAfter: `${n}.1`,
      iprEmitAfter: `${n}.2`
    };
  }

  const items = (Array.isArray(agendaItems) ? agendaItems : [])
    .filter(it => it && typeof it.number === 'string' && it.number.trim() !== '');
  const levelOf = num => String(num).split('.').length;
  function shallowestFirst(predicate) {
    let best = null;
    items.forEach(it => {
      if (!predicate(it)) return;
      if (!best || levelOf(it.number) < levelOf(best.number)) best = it;
    });
    return best;
  }
  function childOffsetOf(num) {
    let max = 0;
    items.forEach(it => {
      if (it.number.indexOf(num + '.') !== 0) return;
      const rest = it.number.slice(num.length + 1);
      if (/^\d+$/.test(rest)) max = Math.max(max, parseInt(rest, 10));
    });
    return max;
  }
  function lastInSubtree(num) {
    let last = num;
    items.forEach(it => {
      if (it.number === num || it.number.indexOf(num + '.') === 0) last = it.number;
    });
    return last;
  }

  let source = 'adhoc-semantic';
  let opening = shallowestFirst(it =>
    /^\s*opening\b|\bopening of (the )?(session|meeting)\b/i.test(it.title || ''));
  if (!opening) {
    opening = shallowestFirst(() => true);
    source = 'adhoc-fallback';
  }
  if (!opening) {
    return {
      source: 'adhoc-none', prefixNum: null, openingSection: null, openingSubSection: null,
      registrationSection: null, reallocationSection: null, documentsSection: null,
      afterDocumentsSection: null, iprSection: null, iprSyntheticSection: null,
      iprChildOffset: 0, openingEmitAfter: null, iprEmitAfter: null
    };
  }

  const o = opening.number;
  const ipr = shallowestFirst(it =>
    it.number !== o && /\b(ipr|antitrust|competition law)\b/i.test(it.title || ''));
  const k = childOffsetOf(o);

  return {
    source: source,
    prefixNum: o,
    openingSection: o,
    openingSubSection: `${o}.${k + 1}`,
    registrationSection: `${o}.${k + 2}`,
    reallocationSection: `${o}.${k + 3}`,
    documentsSection: `${o}.${k + 4}`,
    afterDocumentsSection: `${o}.${k + 5}`,
    iprSection: ipr ? ipr.number : null,
    iprSyntheticSection: ipr ? null : `${o}.${k + 5}`,
    iprChildOffset: ipr ? childOffsetOf(ipr.number) : 0,
    openingEmitAfter: lastInSubtree(o),
    iprEmitAfter: ipr ? lastInSubtree(ipr.number) : null,
    _realNumbers: items.reduce((m, it) => { m[it.number] = true; return m; }, {})
  };
}

/**
 * ADDON-007A: the "pre-agenda" TDoc capture rule (see
 * isBeforeRegistrationBoundary_()) for the resolved anchors. Main meetings
 * keep the existing prefix rule exactly. For ad-hoc meetings a TDoc is
 * captured only when it is filed under an administrative anchor item
 * itself (the opening item, the IPR item, or a synthetic IPR child) --
 * items that are special-cased and can never be a plain TDoc-table target.
 * A TDoc filed under a REAL child item of the opening item stays with it.
 */
function isBeforeAdminBoundary_(agendaItemValue, anchors) {
  if (!anchors) return false;
  if (anchors.source === 'main-prefix') {
    return isBeforeRegistrationBoundary_(agendaItemValue, anchors.prefixNum);
  }
  const v = String(agendaItemValue || '').trim();
  if (!v || !anchors.openingSection) return false;
  if (v === anchors.openingSection) return true;
  if (anchors.iprSection && (v === anchors.iprSection || v.indexOf(anchors.iprSection + '.') === 0)) {
    return v === anchors.iprSection || !isRealItemNumber_(anchors, v);
  }
  return false;
}
function isRealItemNumber_(anchors, v) {
  return !!(anchors._realNumbers && anchors._realNumbers[v]);
}

/**
 * SA4-PROD-003: renders one TDoc's standard detail table (TDoc / Title /
 * Source / Contact / Agenda Item / Type-For / E-mail Discussion /
 * Revisions / Minutes / Disposition / Status) -- identical row schema and
 * cell logic to the two existing inline per-TDoc-table blocks inside
 * buildSkeletonWithTdocTables() (both left untouched by this task).
 * Extracted here only so the NEW {agendaPrefixNum}.1.4 "Documents" section
 * does not become a third verbatim copy of that block; the two existing
 * call sites are not migrated to it.
 *
 * `agendaItemLabel` is passed explicitly (not read from a loop item) so a
 * caller can preserve a TDoc's ORIGINAL agenda-item value even when the
 * section it is rendered under (here, "{agendaPrefixNum}.1.4") is not that
 * value -- required so the Document Reallocations workflow can see where
 * the document actually came from.
 */
function appendTdocDetailTable_(body, tdocData, agendaItemLabel) {
  const row = tdocData.row;
  const revisedTo = getRevisedTo_(tdocData);
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
    ['Agenda Item', agendaItemLabel],
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
}

/**
 * SA4-PROD-008: canonical, report-boilerplate IPR/antitrust/consensus
 * subsections (X.2.1-X.2.4), generated directly rather than copied from
 * any template document -- this exact wording is standard across every
 * SA4 report and is never meeting-specific. Prefix-independent via
 * `agendaPrefixNum` (never a hardcoded "5" or any other literal number).
 *
 * The caller is responsible for the "X.2 ..." heading itself (the real
 * agenda item's own title, already appended before this is called) -- this
 * function only appends the X.2.1-X.2.4 child content beneath it, so no
 * duplicate X.2 heading is ever created.
 *
 * Wording is preserved EXACTLY as supplied/approved, including its
 * internal inconsistency between a closing ASCII quote (Call for IPRs) and
 * closing curly quotes (Statement regarding competition law / Consensus
 * principles reminder) -- do not "fix" this without a separate, explicit
 * decision to do so.
 */
// ADDON-007A: optional iprSectionNumber/childOffset let an ad-hoc meeting
// anchor these children under its REAL IPR agenda item; omitted (every main
// meeting), the numbers are the unchanged `${agendaPrefixNum}.2.1`..`.2.4`.
function appendStandardIprSection_(body, agendaPrefixNum, iprSectionNumber, childOffset) {
  const base = iprSectionNumber || `${agendaPrefixNum}.2`;
  const off = childOffset || 0;
  body.appendParagraph(`${base}.${off + 1} Introduction`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('The chair read the antitrust and IPR clause at the opening of the meeting session.');

  body.appendParagraph(`${base}.${off + 2} Call for IPRs`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('“I draw your attention to your obligations under the 3GPP Partner Organizations’ IPR policies. Every Individual Member organization is obliged to declare to the Partner Organization or Organizations of which it is a member any IPR owned by the Individual Member or any other organization which is or is likely to become essential to the work of 3GPP.');
  body.appendParagraph('Delegates are asked to take note that they are thereby invited:');
  body.appendListItem('to investigate whether their organization or any other organization owns IPRs which were, or were likely to become Essential in respect of the work of 3GPP.').setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem('to notify their respective Organizational Partners of all potential IPRs, e.g., for ETSI, by means of the IPR Information Statement and the Licensing declaration forms"').setGlyphType(DocumentApp.GlyphType.BULLET);

  body.appendParagraph(`${base}.${off + 3} Statement regarding competition law`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('“I also draw your attention to the fact that 3GPP activities are subject to all applicable antitrust and competition laws and that compliance with said laws is therefore required of any participant of this TSG/WG/SWG meeting including the Chair and Vice Chairs. In case of question I recommend that you contact your legal counsel.');
  body.appendParagraph('The leadership shall conduct the present meeting with impartiality and in the interests of 3GPP.');
  body.appendParagraph('Furthermore, I would like to remind you that timely submission of work items in advance of TSG/WG/SWG meetings is important to allow for full and fair consideration of such matters.”');

  body.appendParagraph(`${base}.${off + 4} Consensus principles reminder`).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph('“I also draw your attention to the fact that 3GPP endeavours to reach consensus on all decisions and therefore depends on a cooperative spirit of the Individual Members. In particular, Individual Members are encouraged to seek a consensus-based solution and only to sustain objections as a very last resort, and where absolutely necessary and well justified. The leadership will conduct the present meeting in a manner whereby informal methods of reaching consensus are encouraged, whilst ensuring that well justified concerns are taken into account.”');
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

// TEMPLATE-002C: each single step is a non-interactive core (also a Full
// Build phase) and a menu wrapper that keeps its completion pop-up.
function collectEmailDiscussionCore_() {
  const cfg = getCollectorConfig_();
  checkRSSFeed_(cfg);
}

function collectEmailDiscussionOnly() {
  collectEmailDiscussionCore_();
  DocumentApp.getUi().alert('Success', 'E-mail discussion collection completed.', DocumentApp.getUi().ButtonSet.OK);
}

function collectRevisionsCore_() {
  const cfg = getCollectorConfig_();
  updateRevisions_(cfg);
}

function collectRevisionsOnly() {
  collectRevisionsCore_();
  DocumentApp.getUi().alert('Success', 'Revision collection completed.', DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Fetch and insert abstracts for every TDOC table that does not have one yet.
 * No UI, so it is safe to call from triggers. Returns the number of tables filled.
 */
/**
 * PERF-006B (diagnostics only, no behavior change): "candidate tables
 * processed" (this function's own loop-iteration count, unchanged from
 * the pre-existing `count`) is NOT the same thing as "Reviewer requests
 * actually made" -- a candidate can also resolve via a negative-cache
 * skip (isReviewerNoSummaryCached_(), zero fetches) without ever calling
 * the Reviewer API. Production validation of the PERF-006 negative cache
 * showed this concretely: a run with 1 candidate table and a fresh
 * negative-cache hit still logged the old "Fetched abstracts for 1
 * table(s)" at the call site below -- wrong, since zero Reviewer API
 * requests occurred that run. Returns a breakdown object instead of a
 * bare count so each of this function's two callers can log accurately
 * for what THEY care about, computed as a delta of the SAME global PERF
 * counters fetchAndAddAbstract_() already increments (Reviewer API
 * requests / Reviewer negative-cache hits / structural: abstract row
 * inserted), snapshotted immediately before this function's own loop so
 * an EARLIER-in-the-same-run fetchAndAddAbstract_() call from a different
 * call site (insertNewTdoc_() -> createTDocTableFromData_(), for a
 * brand-new TDoc table) is correctly excluded from this function's own
 * delta.
 */
function addAbstractsForTables_(body, context, leaveOut) {
  body = body || getActiveDocumentBodyCounted_('addAbstractsForTables_', context);
  let candidatesProcessed = 0;
  // Only the update passes `leaveOut`, { number: true }: TDocs that are not
  // uploaded yet (tdocsNotUploadedYet_()) or were asked for earlier in the
  // same run. Without it every table is a candidate, as before (Full Build,
  // "Update Abstracts").

  const requestsBefore = perfCounterValue_('Reviewer API requests');
  const cacheSkipsBefore = perfCounterValue_('Reviewer negative-cache hits');
  const rowsInsertedBefore = perfCounterValue_('structural: abstract row inserted (fetchAndAddAbstract_)');

  getTablesCounted_(body, 'addAbstractsForTables_').forEach(table => {
    if (!isTDocTable_(table)) return;

    const existingAbstract = findCellText_(table, 'Abstract');
    if (existingAbstract) return;

    // SA4-PROD-006: same migration as createTDocTableFromData_() -- was a
    // hardcoded main-meeting-only /^S4-\d{6}$/ regex, now the centralized
    // registered-family check, canonical spelling passed to Reviewer.
    const tdocNumber = String(safeCellText_(table, 0, 1) || '').trim();
    const parsedTdoc = parseExactSA4DocumentId_(tdocNumber);
    if (!parsedTdoc.isValid) return;
    if (leaveOut && leaveOut[tdocNumber] === true) return;

    perfTimedAccum_('Reviewer API abstract fetch (fetchAndAddAbstract_, accumulated)', () => fetchAndAddAbstract_(table, parsedTdoc.raw, context));
    candidatesProcessed++;
  });

  // Callers log their own message from this breakdown (see
  // continuousUpdateCore_() and addAbstractsOnly()) -- not logged here, to
  // avoid a duplicate line every run.
  return {
    candidatesProcessed: candidatesProcessed,
    requestsMade: perfCounterValue_('Reviewer API requests') - requestsBefore,
    cacheSkips: perfCounterValue_('Reviewer negative-cache hits') - cacheSkipsBefore,
    rowsInserted: perfCounterValue_('structural: abstract row inserted (fetchAndAddAbstract_)') - rowsInsertedBefore
  };
}

function addAbstractsOnly() {
  // Manual step: always runs, independent of the trigger-level switch.
  const result = addAbstractsForTables_();
  DocumentApp.getUi().alert('Success', `Abstract step completed: ${result.candidatesProcessed} candidate table(s) processed, ${result.rowsInserted} abstract(s) inserted.`, DocumentApp.getUi().ButtonSet.OK);
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
  const result = applyDocumentReallocationsToBody_(body, reallocations);

  Logger.log(`Applied reallocations: ${result.movedCount} moved, ${result.removedCount} removed, ${result.updatedCount} updated, ${result.notFound.length} not in report`);
  ui.alert(
    'Reallocations Applied',
    `✅ Processing complete:\n\n` +
    `📋 Updated agenda items: ${result.updatedCount}\n` +
    `🔄 Moved tables: ${result.movedCount}\n` +
    `🗑️ Removed tables: ${result.removedCount}` +
    (result.notFound.length ? `\n\nNot in this report (skipped): ${result.notFound.join(', ')}` : ''),
    ui.ButtonSet.OK
  );
}

const REALLOCATION_REMOVAL_TARGETS_ = ['removed', 'withdrawn', 'n/a'];

/**
 * ADDON-008A1b: the ONE interpretation of a saved reallocation's "New
 * Agenda" value, shared by Apply Document Reallocations, the TDoc grouping
 * (build and continuous update) and the build's agenda projection:
 *   'remove' -- removed / withdrawn / n/a: not placed in the report;
 *   'move'   -- an agenda item number: the TDoc's report allocation;
 *   'invalid'-- anything else.
 */
function interpretReallocationTarget_(value) {
  const v = String(value === null || value === undefined ? '' : value).trim();
  if (REALLOCATION_REMOVAL_TARGETS_.indexOf(v.toLowerCase()) !== -1) return { kind: 'remove', agendaItem: null, value: v };
  if (/^\d+(?:\.\d+)*$/.test(v)) return { kind: 'move', agendaItem: v, value: v };
  return { kind: 'invalid', agendaItem: null, value: v };
}

/**
 * ADDON-008A1b: effective REPORT allocation of every TDoc in a TDoc list
 * (sheet values): the saved reallocation when there is one, otherwise the
 * source "Agenda item". Source data is never changed. `effective` is null
 * for a removed TDoc.
 */
function computeEffectiveTdocAllocations_(tdocListValues, reallocations) {
  const rows = Array.isArray(tdocListValues) ? tdocListValues : [];
  const header = (rows[0] || []).map(function (h) { return String(h === null || h === undefined ? '' : h).trim(); });
  const tdocCol = header.indexOf('TDoc');
  const itemCol = header.indexOf('Agenda item');
  const out = [];
  if (tdocCol === -1 || itemCol === -1) return out;
  for (let i = 1; i < rows.length; i++) {
    const tdoc = String((rows[i] || [])[tdocCol] || '').trim();
    if (!tdoc) continue;
    const source = String(rows[i][itemCol] === null || rows[i][itemCol] === undefined ? '' : rows[i][itemCol]).trim();
    const realloc = reallocations ? reallocations[tdoc] : null;
    const target = realloc ? interpretReallocationTarget_(realloc.new) : null;
    out.push({
      tdoc: tdoc,
      source: source,
      reallocated: !!target,
      removed: !!target && target.kind === 'remove',
      effective: !target ? source : (target.kind === 'move' ? target.agendaItem : (target.kind === 'remove' ? null : target.value))
    });
  }
  return out;
}

/**
 * ADDON-008A1b: why saved reallocations cannot be honoured by a build, one
 * message per entry (TDoc, source allocation, requested destination,
 * reason). `agendaNumbers`, when given, are the agenda items a destination
 * must exist in (ad-hoc builds); main-meeting builds pass null, since a
 * main-meeting TDoc may legitimately move to another SWG's report.
 */
function findReallocationBuildProblems_(reallocations, sourceByTdoc, agendaNumbers) {
  const problems = [];
  Object.keys(reallocations || {}).forEach(function (tdoc) {
    const realloc = reallocations[tdoc];
    const target = interpretReallocationTarget_(realloc.new);
    const source = (sourceByTdoc && sourceByTdoc[tdoc]) || realloc.original || 'unknown';
    const label = `${tdoc} (${source} → ${target.value || 'blank'})`;
    if (target.kind === 'invalid') {
      problems.push(`${label}: "${target.value}" is not an agenda item number or removed/withdrawn/n/a.`);
    } else if (target.kind === 'move' && agendaNumbers && agendaNumbers.indexOf(target.agendaItem) === -1) {
      problems.push(`${label}: this meeting's agenda has no item ${target.agendaItem}.`);
    }
  });
  return problems;
}

function refuseBuildForReallocations_(problems) {
  throw new Error('Cannot build report -- the report was not changed.\n\nSaved document reallocations cannot be applied:\n• ' +
    problems.join('\n• ') + '\n\nCorrect the Document Reallocations table, then build again.');
}

/**
 * ADDON-008A1: applies the reallocation map to a built report body.
 *
 * The old in-place loop moved a table with
 * `body.insertTable(i, body.removeChild(table).asTable())` -- but
 * Body.removeChild() returns the BODY, not the removed element, so every
 * real move threw "BODY_SECTION can't be cast to TABLE." AFTER the table
 * had already been removed (and its Agenda Item row already rewritten).
 *
 * Now two phases:
 *   1. plan -- resolve every entry's source TDoc table and destination
 *      agenda section, validating the element types, WITHOUT touching the
 *      document; any unresolvable entry aborts the whole run with a
 *      message naming the TDoc, old and new agenda item and the reason;
 *   2. apply -- a move inserts an updated COPY at the destination first and
 *      only then removes the original, so a failed insert leaves it intact.
 * Unchanged semantics: one TDoc detail table per TDoc, moved to the end of
 * its new agenda item's section and its "Agenda Item" row set to the new
 * value; "removed"/"withdrawn"/"n/a" remove the table; a table already at
 * its destination is only relabelled; agenda headings are never added or
 * removed. A TDoc with no table in the report is skipped (it is placed by
 * the next build/update, which already applies the reallocation map).
 * A destination with no heading in the report is now an error rather than
 * a heading silently appended at the end of the document.
 */
function applyDocumentReallocationsToBody_(body, reallocations) {
  const TABLE = DocumentApp.ElementType.TABLE;
  const tdocTables = {};
  body.getTables().forEach(function (table) {
    if (!isTDocTable_(table)) return;
    const id = safeCellText_(table, 0, 1).trim();
    if (!id) return;
    (tdocTables[id] = tdocTables[id] || []).push(table);
  });

  // ---- phase 1: plan (read-only) ----
  const plan = [];
  const notFound = [];
  const problems = [];
  Object.keys(reallocations).forEach(function (tdoc) {
    const realloc = reallocations[tdoc];
    const target = String(realloc.new || '').trim();
    const label = `${tdoc} (${realloc.original || 'unknown'} → ${target || 'blank'})`;
    const tables = tdocTables[tdoc] || [];
    if (tables.length === 0) { notFound.push(tdoc); return; }
    if (tables.length > 1) { problems.push(`${label}: the report contains ${tables.length} tables for this TDoc.`); return; }
    const table = tables[0];
    const parent = table.getParent();
    if (table.getType() !== TABLE || !parent || parent.getType() !== body.getType()) {
      problems.push(`${label}: its table is not a top-level table of the report body.`);
      return;
    }
    const interpreted = interpretReallocationTarget_(target);
    if (interpreted.kind === 'remove') {
      plan.push({ tdoc: tdoc, realloc: realloc, table: table, remove: true });
      return;
    }
    if (interpreted.kind !== 'move') {
      problems.push(`${label}: "${target}" is not an agenda item number.`);
      return;
    }
    if (findAgendaSectionEndIndex_(body, target) === -1) {
      problems.push(`${label}: the report has no agenda item ${target} heading.`);
      return;
    }
    plan.push({ tdoc: tdoc, realloc: realloc, table: table, remove: false, target: target });
  });
  if (problems.length > 0) {
    throw new Error('No reallocations were applied -- the report was not changed.\n\n• ' + problems.join('\n• '));
  }

  // ---- phase 2: apply ----
  let movedCount = 0;
  let removedCount = 0;
  let updatedCount = 0;
  plan.forEach(function (step) {
    if (step.remove) {
      Logger.log(`Removing table for ${step.tdoc} (reallocated to ${step.realloc.new})`);
      step.table.removeFromParent();
      removedCount++;
      return;
    }
    // Section boundaries are re-read per step: earlier moves shift indices.
    // ADDON-008A1b: a table already under its destination heading (e.g.
    // placed there by a rebuild that applied this reallocation) stays put.
    const currentIndex = body.getChildIndex(step.table);
    const targetIndex = findAgendaSectionEndIndex_(body, step.target);
    const moving = !isUnderAgendaHeading_(body, currentIndex, step.target) &&
      targetIndex !== currentIndex && targetIndex !== currentIndex + 1;
    const placed = moving ? step.table.copy() : step.table;
    if (setTdocTableAgendaItem_(placed, step.realloc.new)) {
      updatedCount++;
      Logger.log(`Updated agenda item for ${step.tdoc}: ${step.realloc.original} → ${step.realloc.new}`);
    }
    if (moving) {
      body.insertTable(targetIndex, placed);
      step.table.removeFromParent();
      movedCount++;
      Logger.log(`Moved table for ${step.tdoc} from index ${currentIndex} to ${targetIndex}`);
    }
  });

  return { movedCount: movedCount, removedCount: removedCount, updatedCount: updatedCount, notFound: notFound };
}

/** True when the nearest heading above body child `index` is this agenda item's heading. */
function isUnderAgendaHeading_(body, index, agendaItem) {
  for (let i = index - 1; i >= 0; i--) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const para = child.asParagraph();
    if (para.getHeading() === DocumentApp.ParagraphHeading.NORMAL) continue;
    const text = para.getText().trim();
    return text === agendaItem || text.indexOf(agendaItem + ' ') === 0;
  }
  return false;
}

/** Sets a TDoc detail table's "Agenda Item" value; false when it has no such row. */
function setTdocTableAgendaItem_(table, value) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 2) continue;
    const key = row.getCell(0).getText().trim().toLowerCase();
    if (key === 'agenda item' || key === 'agenda item:') {
      row.getCell(1).setText(value);
      return true;
    }
  }
  return false;
}

// =========================================================
// TDOC UPLOAD COMPLETION -- what an existing TDoc gets once it is uploaded
// =========================================================
//
// A TDoc can enter the report while it is only reserved. The TDoc list has
// no hyperlink for it then, and the Reviewer has no summary. Its table is
// inserted without a link, its row of the registration table too, and no
// abstract can be had. Everything else in the table comes from the
// reservation and does not depend on the upload.
//
// A TDoc comes to be in the report as uploaded in one of two ways, and is
// completed in both:
//   A. its table exists and the list now shows the upload (the transition
//      below): refreshUploadedTdocMetadata_();
//   B. its table is inserted when the list already shows the upload: the
//      insertion writes the link, completeInsertedUploadedTdoc_() asks for
//      the abstract.
//
// The update completes an existing TDoc once the list shows the upload:
//   - THE SIGN OF AN UPLOAD is the hyperlink the TDoc list has on the TDoc
//     number (the Portal sets it when the file is there);
//   - THE TRANSITION is: the list has that link and the TDoc number in the
//     report has none. It is acted on once, because acting on it sets the
//     link. Nothing is stored for it.
// At the transition the link is added, a "no summary" answer the Reviewer
// gave before the upload is forgotten, and the abstract is asked for once,
// whatever the abstracts-on-update setting says. An answer "no summary" is cached
// as always, and no later update asks again on account of the upload.
//
// Nothing else is touched: not the Status, not Minutes or Disposition, not
// Title, Source, Contact, Type/For or Agenda Item, and never a link or an
// abstract that is already there.

/** Whether the TDoc list tells if this TDoc has a hyperlink (it does not when its rich text was not read). */
function tdocListLinkIsKnown_(tdocData) {
  const richText = tdocData && tdocData.richTextRow && tdocData.richTextRow[tdocData.tdocCol];
  return !!(richText && typeof richText.getLinkUrl === 'function');
}

/**
 * The hyperlink the TDoc list has on a TDoc's number, or '' when it has
 * none (a TDoc that is not uploaded yet). The same reading of the list
 * cell that insertTDocTableAtIndex_() uses for a new table.
 */
function tdocListLinkUrl_(tdocData) {
  return tdocListLinkIsKnown_(tdocData) ? (tdocData.richTextRow[tdocData.tdocCol].getLinkUrl() || '') : '';
}

/**
 * The TDocs the list shows as not uploaded yet: { number: true }. A TDoc
 * is only in it when the list is known to have no link for it; a list
 * whose links cannot be read names none.
 */
function tdocsNotUploadedYet_(allTdocs) {
  const out = {};
  (allTdocs || []).forEach(function (td) {
    const number = String(td.row[td.tdocCol] || '').trim();
    if (!number) return;
    if (tdocListLinkIsKnown_(td) && !tdocListLinkUrl_(td)) out[number] = true;
    else delete out[number];
  });
  return out;
}

/**
 * Sets `url` on the TDoc number in `cell` when the cell has no link at all.
 * A cell with a link anywhere in its text -- whichever link -- is left as
 * it is. Only the link is set: the text and its other formatting stay.
 * Returns true when the link was set.
 */
function addMissingTdocLink_(cell, tdocNumber, url) {
  const cellText = cell.getText();
  const start = cellText.indexOf(tdocNumber);
  if (!url || !tdocNumber || start === -1 || cellText.trim() !== tdocNumber) return false;

  const text = cell.editAsText();
  for (let i = 0; i < cellText.length; i++) {
    if (text.getLinkUrl(i)) return false;
  }
  text.setLinkUrl(start, start + tdocNumber.length - 1, url);
  return true;
}

/**
 * Completes an EXISTING TDoc table at its upload transition (see the
 * section comment): adds the link of the TDoc list to the TDoc number,
 * forgets a "no summary" answer from before the upload, and -- when the
 * table has no Abstract -- asks the Reviewer once, through
 * fetchAndAddAbstract_() and whatever the abstracts setting is.
 *
 * Does nothing when the list has no link or the report already has one.
 * Returns { linkAdded, abstractAttempted }. Never throws: what cannot be
 * done is logged and does not stop the update.
 */
function refreshUploadedTdocMetadata_(body, tdocNumber, tdocData, index, context) {
  const done = { linkAdded: false, abstractAttempted: false };
  try {
    const url = tdocListLinkUrl_(tdocData);
    if (!url || !tdocNumber) return done;

    const table = findTdocTable_(body, tdocNumber, index);
    if (!table || safeCellText_(table, 0, 1).trim() !== tdocNumber) return done;

    // The link first: it is what marks the transition as acted on, so the
    // Reviewer is asked on account of this upload at most once.
    if (!addMissingTdocLink_(table.getRow(0).getCell(1), tdocNumber, url)) return done;
    done.linkAdded = true;
    Logger.log(`Upload of ${tdocNumber}: link added (${url})`);

    const parsed = parseExactSA4DocumentId_(tdocNumber);
    if (!parsed.isValid) return done;
    clearReviewerNoSummaryCache_(parsed.raw, context);
    if (findCellText_(table, 'Abstract')) return done;
    done.abstractAttempted = true;
    fetchAndAddAbstract_(table, parsed.raw, context);
  } catch (e) {
    Logger.log(`Upload of ${tdocNumber}: could not be completed: ${e.message}`);
  }
  return done;
}

/**
 * Asks the Reviewer once for the abstract of a TDoc table that was just
 * INSERTED for a TDoc the list shows as uploaded (way B of the section
 * comment), through fetchAndAddAbstract_() and whatever the abstracts
 * setting is. A "no summary" answer that is cached is respected. Nothing
 * is asked for a TDoc that is not uploaded. Returns true when the abstract
 * was asked for. Never throws.
 */
function completeInsertedUploadedTdoc_(body, tdocNumber, tdocData, index, context) {
  try {
    if (!tdocNumber || !tdocListLinkUrl_(tdocData)) return false;
    const parsed = parseExactSA4DocumentId_(tdocNumber);
    if (!parsed.isValid) return false;

    const table = findTdocTable_(body, tdocNumber, index);
    if (!table || safeCellText_(table, 0, 1).trim() !== tdocNumber) return false;
    if (findCellText_(table, 'Abstract')) return false;
    fetchAndAddAbstract_(table, parsed.raw, context);
    return true;
  } catch (e) {
    Logger.log(`Upload of ${tdocNumber}: its abstract could not be asked for: ${e.message}`);
    return false;
  }
}

/**
 * Adds the link of the TDoc list to the rows of the registration table
 * whose TDoc number has none. No row is added, removed or moved, and no
 * other cell is touched. Returns the number of links added; never throws.
 */
function refreshRegistrationTableLinks_(body, allTdocs) {
  let added = 0;
  try {
    const found = findRegistrationTable_(body);
    if (!found) return 0;
    const urls = {};
    (allTdocs || []).forEach(function (td) {
      const url = tdocListLinkUrl_(td);
      if (url) urls[String(td.row[td.tdocCol] || '').trim()] = url;
    });
    const table = found.table;
    for (let r = 1; r < table.getNumRows(); r++) {
      const cell = table.getRow(r).getCell(0);
      const tdocNumber = cell.getText().trim();
      if (!tdocNumber || !urls[tdocNumber]) continue;
      if (addMissingTdocLink_(cell, tdocNumber, urls[tdocNumber])) added++;
    }
  } catch (e) {
    Logger.log('Registration table: its links could not be completed: ' + e.message);
  }
  return added;
}

// =========================================================
// AD-HOC SESSIONS (stage A) -- SESSION MODEL AND CONFIGURATION
// =========================================================
//
// docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md. An ad-hoc report can describe
// the meetings of its series as sessions. A session is one logical meeting
// and may span several calendar days. It is metadata: a stable id, a label,
// its first day (`date`), optionally its last day (`endDate`; absent for a
// session of one day) and optional planned times -- `start` on the first
// day, `end` on the last, where it is the cut-off for the TDocs of the
// session. All are wall-clock values of the report's time zone, stored as
// text. A session of one day is stored exactly as before `endDate` existed.
//
// Opt-in and ad-hoc only: sessions are active when the meeting type is
// adhoc AND the ADHOC_SESSIONS property holds a valid configuration with at
// least one session (adhocSessionsEnabled_()). Everything else of the
// feature -- the Session column, the Attendance and Session administration
// sections, the status -- reads them through getAdhocSessions_().
//
// The property is written by one function (saveAdhocSessionsWith_()) and
// only on Save. A value that cannot be read is treated as "no sessions" by
// every reader and never throws.

const ADHOC_SESSIONS_KEY_ = 'ADHOC_SESSIONS';
const ADHOC_SESSIONS_SCHEMA_VERSION_ = 1;
const ADHOC_SESSION_LABEL_MAX_LENGTH_ = 16;
// One property value holds at most 9 KB (see the report state store above).
const ADHOC_SESSIONS_MAX_CHARS_ = 8000;
const ADHOC_SESSION_MONTHS_ = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** A real calendar date written YYYY-MM-DD. */
function isValidAdhocSessionDate_(value) {
  const m = String(value === null || value === undefined ? '' : value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return false;
  const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
}

/** A time of day written HH:mm (24-hour clock). */
function isValidAdhocSessionTime_(value) {
  return /^([01]\d|2[0-3]):[0-5]\d$/.test(String(value === null || value === undefined ? '' : value));
}

/**
 * "22 Sep" for 2026-09-22; with a later last day "26–28 Oct", "30 Oct – 1
 * Nov" or "31 Dec – 2 Jan". '' for anything that is not a valid date.
 */
function defaultAdhocSessionLabel_(date, endDate) {
  if (!isValidAdhocSessionDate_(date)) return '';
  const day = function (d) { return String(parseInt(d.slice(8, 10), 10)); };
  const month = function (d) { return ADHOC_SESSION_MONTHS_[parseInt(d.slice(5, 7), 10) - 1]; };
  if (!isValidAdhocSessionDate_(endDate) || endDate <= date) return day(date) + ' ' + month(date);
  if (endDate.slice(0, 7) === date.slice(0, 7)) return day(date) + '–' + day(endDate) + ' ' + month(date);
  return day(date) + ' ' + month(date) + ' – ' + day(endDate) + ' ' + month(endDate);
}

/** The label shown for a session: the user's own, else the one derived from its dates. */
function adhocSessionLabel_(session) {
  return (session && session.label) || defaultAdhocSessionLabel_(session && session.date, session && session.endDate);
}

/** The last calendar day of a session: its endDate, else its date. */
function adhocSessionLastDay_(session) {
  return (session && session.endDate) || (session && session.date) || '';
}

/** True for a session that spans more than one calendar day. */
function isAdhocMultiDaySession_(session) {
  return !!(session && session.endDate && session.endDate > session.date);
}

/** How many calendar days a session spans (1 for a session of one day). */
function adhocSessionDayCount_(session) {
  const at = function (d) { const m = String(d).match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : NaN; };
  const days = Math.round((at(adhocSessionLastDay_(session)) - at(session.date)) / 86400000) + 1;
  return days >= 1 ? days : 1;
}

/**
 * THE way a range of days is written: "September 22, 2026", "October 26–28,
 * 2026", "October 30 – November 1, 2026", "December 31, 2026 – January 2,
 * 2027". `first` and `last` are YYYY-MM-DD; '' when `first` is not a date.
 */
function adhocDateRangeText_(first, last) {
  const a = adhocLongDate_(first);
  const b = adhocLongDate_(last);
  if (!a || !b || b === a || String(last) < String(first)) return a;
  const pa = a.match(/^(\S+) (\d+), (\d+)$/);
  const pb = b.match(/^(\S+) (\d+), (\d+)$/);
  if (!pa || !pb || pa[3] !== pb[3]) return a + ' – ' + b;
  if (pa[1] === pb[1]) return pa[1] + ' ' + pa[2] + '–' + pb[2] + ', ' + pa[3];
  return pa[1] + ' ' + pa[2] + ' – ' + pb[1] + ' ' + pb[2] + ', ' + pa[3];
}

/** The days of one session, written as a range. */
function adhocSessionDatesText_(session) {
  return adhocDateRangeText_(session.date, adhocSessionLastDay_(session));
}

/**
 * Chronological order: date, then planned start. The id decides only between
 * sessions the validation rejects anyway (same date, same or no start), so
 * that sorting is total and never depends on the input order.
 */
function compareAdhocSessions_(a, b) {
  if (a.date !== b.date) return a.date < b.date ? -1 : 1;
  if (a.start !== b.start) return a.start < b.start ? -1 : 1;
  return (parseInt(String(a.id).slice(1), 10) || 0) - (parseInt(String(b.id).slice(1), 10) || 0);
}

function sortAdhocSessions_(sessions) {
  return sessions.slice().sort(compareAdhocSessions_);
}

/**
 * Pure: the short code of the session at position `index` (from 0) of the
 * chronological order: "A01", "A02", ... "A10", ... "A100". It is what the
 * Session column of the registration table shows, where a label would not
 * fit. A code is derived from the order every time it is needed and is
 * never stored: sessions are stored and referred to by their ids only.
 */
function adhocSessionCode_(index) {
  const n = index + 1;
  return 'A' + (n < 10 ? '0' : '') + n;
}

/** Pure: { <session id>: code } for these sessions, by their chronological order. */
function adhocSessionCodes_(sessions) {
  const codes = {};
  sortAdhocSessions_(sessions).forEach(function (session, i) { codes[session.id] = adhocSessionCode_(i); });
  return codes;
}

/**
 * Normalizes and validates a list of sessions that already have their ids.
 * Returns { errors, sessions }: `sessions` is the normalized list in
 * chronological order and is meaningful only when `errors` is empty.
 *
 * Sessions on the same date must be orderable without guessing: each needs
 * a planned start, the starts differ, and every session but the last of
 * that day needs a planned end that is not after the next start.
 *
 * A session may span several days (endDate after date). Its planned start
 * is then on the first day and its planned end on the last, so the end need
 * not be later in the day than the start. The days of two sessions must not
 * overlap; a session may begin on the last day of the one before only when
 * that one's planned end and its own planned start say so.
 */
function validateAdhocSessions_(sessions) {
  const errors = [];
  if (!Array.isArray(sessions)) return { errors: ['The session list is missing.'], sessions: [] };
  const text = function (v) { return String(v === null || v === undefined ? '' : v).trim(); };

  const normalized = [];
  const seenIds = {};
  sessions.forEach(function (raw, i) {
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) {
      errors.push('Session ' + (i + 1) + ' is not a session.');
      return;
    }
    const s = { id: text(raw.id), label: text(raw.label).replace(/\s+/g, ' '), date: text(raw.date) };
    // The last day, only for a session of several days: one that ends on its first day is stored as before.
    const endDate = text(raw.endDate);
    if (endDate && endDate !== s.date) s.endDate = endDate;
    s.start = text(raw.start);
    s.end = text(raw.end);
    const name = s.label || s.date || 'Session ' + (i + 1);

    if (!/^s[1-9]\d*$/.test(s.id)) errors.push(name + ': the session id is missing or malformed.');
    else if (seenIds[s.id]) errors.push(name + ': the session id ' + s.id + ' is used twice.');
    else seenIds[s.id] = true;

    if (!s.date) errors.push(name + ': enter a date.');
    else if (!isValidAdhocSessionDate_(s.date)) errors.push(name + ': "' + s.date + '" is not a date (YYYY-MM-DD).');
    if (s.start && !isValidAdhocSessionTime_(s.start)) errors.push(name + ': the planned start "' + s.start + '" is not a time (HH:mm).');
    if (s.end && !isValidAdhocSessionTime_(s.end)) errors.push(name + ': the planned end "' + s.end + '" is not a time (HH:mm).');
    if (s.endDate && !isValidAdhocSessionDate_(s.endDate)) errors.push(name + ': the end date "' + s.endDate + '" is not a date (YYYY-MM-DD).');
    else if (s.endDate && isValidAdhocSessionDate_(s.date) && s.endDate < s.date) errors.push(name + ': the end date is before the start date.');
    // On one day the end is later than the start. Over several days the start is on the first and the end on the last.
    if (!s.endDate && isValidAdhocSessionTime_(s.start) && isValidAdhocSessionTime_(s.end) && s.end <= s.start) {
      errors.push(name + ': the planned end must be after the planned start.');
    }
    if (s.label.length > ADHOC_SESSION_LABEL_MAX_LENGTH_) {
      errors.push(name + ': the label is longer than ' + ADHOC_SESSION_LABEL_MAX_LENGTH_ + ' characters.');
    }
    normalized.push(s);
  });
  if (errors.length) return { errors: errors, sessions: [] };

  const sorted = sortAdhocSessions_(normalized);

  const seenLabels = {};
  sorted.forEach(function (s) {
    const label = adhocSessionLabel_(s);
    const key = label.toLowerCase();
    if (seenLabels[key]) errors.push('The label "' + label + '" is used by more than one session. Give each session its own label.');
    seenLabels[key] = true;
  });

  sorted.forEach(function (s, i) {
    const next = sorted[i + 1];
    if (!next) return;
    if (isAdhocMultiDaySession_(s) || isAdhocMultiDaySession_(next)) {
      // Sessions of several days: their days must not overlap.
      const last = adhocSessionLastDay_(s);
      if (next.date > last) return;
      if (next.date < last || next.date === s.date) {
        errors.push(adhocSessionLabel_(s) + ' and ' + adhocSessionLabel_(next) + ' overlap: their days must not overlap.');
      } else if (!s.end || !next.start) {
        errors.push(adhocSessionLabel_(next) + ' begins on the last day of ' + adhocSessionLabel_(s) + ': enter the planned end of the first and the planned start of the second, so that their order is known.');
      } else if (s.end > next.start) {
        errors.push(adhocSessionLabel_(s) + ': the planned end (' + s.end + ') is after the start of ' + adhocSessionLabel_(next) + ' on that day (' + next.start + ').');
      }
      return;
    }
    if (next.date !== s.date) return;
    const day = defaultAdhocSessionLabel_(s.date);
    if (!s.start || !next.start) {
      errors.push('Two sessions are on ' + day + ': enter a planned start for each of them, so that their order is known.');
    } else if (s.start === next.start) {
      errors.push('Two sessions on ' + day + ' have the same planned start (' + s.start + ').');
    } else if (!s.end) {
      errors.push(adhocSessionLabel_(s) + ': enter a planned end, because another session follows on the same day.');
    } else if (s.end > next.start) {
      errors.push(adhocSessionLabel_(s) + ': the planned end (' + s.end + ') is after the start of the next session on that day (' + next.start + ').');
    }
  });

  return { errors: errors.filter(function (e, i) { return errors.indexOf(e) === i; }), sessions: errors.length ? [] : sorted };
}

/**
 * Reads an ADHOC_SESSIONS property value.
 *   { status: 'absent' }                      no value
 *   { status: 'ok', config }                  { v, nextId, sessions } -- sessions in order
 *   { status: 'unsupported', error }          written by a newer schema version
 *   { status: 'invalid', error }              anything else that cannot be used
 * Never throws.
 */
function parseAdhocSessionsProperty_(raw) {
  if (raw === null || raw === undefined || String(raw).trim() === '') return { status: 'absent', config: null, error: null };
  let data;
  try {
    data = JSON.parse(String(raw));
  } catch (e) {
    return { status: 'invalid', config: null, error: 'The stored session configuration is not valid JSON.' };
  }
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    return { status: 'invalid', config: null, error: 'The stored session configuration is not an object.' };
  }
  if (data.v !== ADHOC_SESSIONS_SCHEMA_VERSION_) {
    return { status: 'unsupported', config: null, error: 'The stored session configuration has version ' + JSON.stringify(data.v) +
      '; this release reads version ' + ADHOC_SESSIONS_SCHEMA_VERSION_ + '.' };
  }
  const checked = validateAdhocSessions_(data.sessions);
  if (checked.errors.length) return { status: 'invalid', config: null, error: checked.errors.join(' ') };
  const highest = checked.sessions.reduce(function (max, s) { return Math.max(max, parseInt(s.id.slice(1), 10)); }, 0);
  if (typeof data.nextId !== 'number' || Math.floor(data.nextId) !== data.nextId || data.nextId <= highest) {
    return { status: 'invalid', config: null, error: 'The stored session configuration has no usable id counter.' };
  }
  return { status: 'ok', config: { v: ADHOC_SESSIONS_SCHEMA_VERSION_, nextId: data.nextId, sessions: checked.sessions }, error: null };
}

/** The one stored form: fixed key order, sessions in chronological order. */
function serializeAdhocSessions_(config) {
  return JSON.stringify({
    v: ADHOC_SESSIONS_SCHEMA_VERSION_,
    nextId: config.nextId,
    sessions: sortAdhocSessions_(config.sessions).map(adhocSessionRecord_)
  });
}

/** One session as it is stored and handed on: `endDate` only for a session of several days, after `date`. */
function adhocSessionRecord_(s) {
  const out = { id: s.id, label: s.label, date: s.date };
  if (s.endDate && s.endDate !== s.date) out.endDate = s.endDate;
  out.start = s.start;
  out.end = s.end;
  return out;
}

/**
 * The lowest id number that is safe to issue when a stored value cannot be
 * read: above every id and counter that can still be found in its text, so
 * that replacing a damaged configuration does not reuse an id either.
 */
function salvageAdhocSessionNextId_(raw) {
  let next = 1;
  const text = String(raw === null || raw === undefined ? '' : raw);
  const idRe = /"s([1-9]\d{0,8})"/g;
  let m;
  while ((m = idRe.exec(text)) !== null) next = Math.max(next, parseInt(m[1], 10) + 1);
  const counter = text.match(/"nextId"\s*:\s*(\d{1,9})/);
  if (counter) next = Math.max(next, parseInt(counter[1], 10));
  return next;
}

/**
 * Pure: the configuration that results from the dialog's rows.
 * `raw` is the current property value; `rows` is [{ id, label, date,
 * endDate, start, end }], where a blank id is a new session. A row keeps its id whatever
 * else changed; a new row gets the next id of the counter, which only ever
 * grows -- the id of a removed session is never issued again.
 * Returns { ok, errors, config }.
 */
function applyAdhocSessionsEdit_(raw, rows) {
  const stored = parseAdhocSessionsProperty_(raw);
  if (stored.status === 'unsupported') {
    return { ok: false, errors: [stored.error + ' It was written by a newer release and is not changed here.'], config: null };
  }
  if (!Array.isArray(rows)) return { ok: false, errors: ['The session list is missing.'], config: null };

  const known = {};
  if (stored.status === 'ok') stored.config.sessions.forEach(function (s) { known[s.id] = true; });
  let nextId = stored.status === 'ok' ? stored.config.nextId : salvageAdhocSessionNextId_(raw);

  const errors = [];
  const sessions = rows.map(function (row) {
    const r = row && typeof row === 'object' ? row : {};
    let id = String(r.id === null || r.id === undefined ? '' : r.id).trim();
    if (!id) {
      id = 's' + nextId;
      nextId++;
    } else if (!known[id]) {
      errors.push('A session of this dialog is no longer stored. Close the dialog and open Configure Sessions again.');
    }
    return { id: id, label: r.label, date: r.date, endDate: r.endDate, start: r.start, end: r.end };
  });
  const checked = validateAdhocSessions_(sessions);
  const all = errors.concat(checked.errors).filter(function (e, i, list) { return list.indexOf(e) === i; });
  if (all.length) return { ok: false, errors: all, config: null };

  const config = { v: ADHOC_SESSIONS_SCHEMA_VERSION_, nextId: nextId, sessions: checked.sessions };
  if (serializeAdhocSessions_(config).length > ADHOC_SESSIONS_MAX_CHARS_) {
    return { ok: false, errors: ['There are too many sessions to store.'], config: null };
  }
  return { ok: true, errors: [], config: config };
}

/** True for an ad-hoc meeting; false for a main meeting and for a meeting type that cannot be read. */
function isAdhocMeetingForSessions_(context) {
  try {
    return normalizeMeetingType_(getReportStateStore_(context).getProperty('MEETING_TYPE')) === 'adhoc';
  } catch (e) {
    return false;
  }
}

/**
 * The sessions of this report in chronological order, each with the label
 * to show (`displayLabel`); [] unless the meeting is ad-hoc and a valid
 * configuration with at least one session is stored. Never throws.
 */
function getAdhocSessions_(context) {
  if (!isAdhocMeetingForSessions_(context)) return [];
  let stored;
  try {
    stored = parseAdhocSessionsProperty_(getReportStateStore_(context).getProperty(ADHOC_SESSIONS_KEY_));
  } catch (e) {
    return [];
  }
  if (stored.status !== 'ok') return [];
  return stored.config.sessions.map(function (s) {
    const session = adhocSessionRecord_(s);
    session.displayLabel = adhocSessionLabel_(s);
    return session;
  });
}

/** The one switch of the session feature (design §3). */
function adhocSessionsEnabled_(context) {
  return getAdhocSessions_(context).length > 0;
}

/**
 * Saves the dialog's rows into `store`. Everything is validated first; a
 * refused save writes nothing. Saving what is already stored writes nothing
 * either, and neither does saving "no sessions" into a report that never
 * had any. Removing every session keeps the property, with its id counter.
 * Returns { ok, errors, changed, sessions }.
 *
 * Stage D: `attendanceSessionIds` are the sessions that have imported
 * attendance (null: that cannot be told). Such a session is not removed
 * here -- its attendance has to be removed first, explicitly -- so that
 * attendance is never deleted, or left without its session, by the way.
 *
 * Stage B: `tdocSessionReferences` are the sessions that manual TDoc
 * assignments name, { <session id>: [TDoc, ...] } (null: that cannot be
 * told). Such a session is not removed either; the assignments are never
 * stripped silently. A session that TDocs belong to only automatically can
 * be removed -- those TDocs are then placed among the remaining sessions.
 *
 * Stage E: `openingSessionIds` are the sessions that have opening details
 * (null: that cannot be told). Such a session is not removed until its
 * details were cleared.
 *
 * Attendance stays inside its session: `attendanceDays` is { <session id>:
 * [the days that have an imported attendance record] } (null: that cannot
 * be told; left out: not checked). The days of a session cannot be changed
 * so that one of those days falls outside it. Such a save is refused as a
 * whole and writes nothing; no attendance is removed or moved, and the
 * session is not extended to fit.
 */
function saveAdhocSessionsWith_(store, isAdhoc, rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays) {
  if (!isAdhoc) return { ok: false, errors: ['Sessions are available for ad-hoc reports only.'], changed: false, sessions: [] };
  const raw = store.getProperty(ADHOC_SESSIONS_KEY_);
  const edit = applyAdhocSessionsEdit_(raw, rows);
  if (!edit.ok) return { ok: false, errors: edit.errors, changed: false, sessions: [] };

  const before = parseAdhocSessionsProperty_(raw);
  const outside = adhocSessionsLeavingAttendanceOutside_(before.status === 'ok' ? before.config.sessions : [], edit.config.sessions, attendanceDays);
  if (outside.length) return { ok: false, errors: outside, changed: false, sessions: [] };
  const removed = (before.status === 'ok' ? before.config.sessions : []).filter(function (s) {
    return !edit.config.sessions.some(function (kept) { return kept.id === s.id; });
  });
  if (removed.length && attendanceSessionIds === null) {
    return { ok: false, errors: ['The stored attendance was written by a newer release, so it cannot be told which sessions have attendance. No session was removed.'], changed: false, sessions: [] };
  }
  const blocked = removed.filter(function (s) { return (attendanceSessionIds || []).indexOf(s.id) !== -1; });
  if (blocked.length) {
    return { ok: false, errors: blocked.map(function (s) {
      return adhocSessionLabel_(s) + ' has imported attendance and was not removed. Remove its attendance first: Sessions and Attendance > Import Teams Attendance… > Remove attendance….';
    }), changed: false, sessions: [] };
  }
  if (removed.length && tdocSessionReferences === null) {
    return { ok: false, errors: ['The manual TDoc session assignments were written by a newer release, so it cannot be told which sessions they name. No session was removed.'], changed: false, sessions: [] };
  }
  const named = removed.filter(function (s) { return ((tdocSessionReferences || {})[s.id] || []).length > 0; });
  if (named.length) {
    return { ok: false, errors: named.map(function (s) {
      const tdocs = tdocSessionReferences[s.id];
      return adhocSessionLabel_(s) + ' is named by manual TDoc session assignments (' + tdocs.slice(0, 3).join(', ') + (tdocs.length > 3 ? ' and ' + (tdocs.length - 3) + ' more' : '') +
        ') and was not removed. Change those assignments first: Sessions and Attendance > Assign TDoc Sessions….';
    }), changed: false, sessions: [] };
  }
  if (removed.length && openingSessionIds === null) {
    return { ok: false, errors: ['The opening details were written by a newer release, so it cannot be told which sessions have them. No session was removed.'], changed: false, sessions: [] };
  }
  const withOpening = removed.filter(function (s) { return (openingSessionIds || []).indexOf(s.id) !== -1; });
  if (withOpening.length) {
    return { ok: false, errors: withOpening.map(function (s) {
      return adhocSessionLabel_(s) + ' has opening details and was not removed. Clear them first: Sessions and Attendance > Edit Opening Details….';
    }), changed: false, sessions: [] };
  }

  const serialized = serializeAdhocSessions_(edit.config);
  const nothingToStore = (raw === null || raw === undefined) && edit.config.sessions.length === 0;
  const changed = !nothingToStore && serialized !== raw;
  if (changed) store.setProperty(ADHOC_SESSIONS_KEY_, serialized);
  return { ok: true, errors: [], changed: changed, sessions: edit.config.sessions };
}

/**
 * Pure: why the edited sessions cannot be saved because of the attendance
 * that is imported -- one message per session whose new days would leave a
 * day with attendance outside, in the order of the sessions; [] when
 * nothing stands in the way. `attendanceDays` as for saveAdhocSessionsWith_().
 */
function adhocSessionsLeavingAttendanceOutside_(storedSessions, editedSessions, attendanceDays) {
  if (attendanceDays === undefined) return [];
  const errors = [];
  if (attendanceDays === null) {
    // Which days have attendance is not known: the days of an existing session are then not changed.
    const moved = editedSessions.filter(function (s) {
      const was = storedSessions.filter(function (x) { return x.id === s.id; })[0];
      return was && (was.date !== s.date || adhocSessionLastDay_(was) !== adhocSessionLastDay_(s));
    });
    return moved.length ? ['The stored attendance was written by a newer release, so it cannot be told which days have attendance. The dates of ' +
      moved.map(adhocSessionLabel_).join(', ') + ' were not changed.'] : [];
  }
  editedSessions.forEach(function (s) {
    const last = adhocSessionLastDay_(s);
    const days = (attendanceDays[s.id] || []).filter(function (day) { return day < s.date || day > last; }).sort();
    if (days.length === 0) return;
    const which = days.map(adhocLongDate_).join(', ');
    errors.push(adhocSessionLabel_(s) + ' cannot be changed to ' + adhocSessionDatesText_(s) + ' because attendance for ' + which + ' is already imported. ' +
      'Remove that attendance first (Sessions and Attendance > Import Teams Attendance… > Remove attendance…) or keep ' + (days.length === 1 ? 'that day' : 'those days') + ' within the session.');
  });
  return errors;
}

/**
 * The days that have an attendance record, per session: { <session id>:
 * [YYYY-MM-DD, ...] }, each day once, in order. The day of a record is the
 * one adhocAttendanceRecordsOf_() gives: the date of its own Teams start
 * time, whatever key it is stored under. A record whose day cannot be told
 * (it cannot be read and its key names none) is left out.
 */
function adhocAttendanceDaysBySession_(state) {
  const days = {};
  Object.keys(state.sessions).concat(state.problems).forEach(function (key) {
    const parts = adhocAttendanceRecordParts_(key);
    if (!parts || days[parts.sessionId]) return;
    const list = [];
    adhocAttendanceRecordsOf_(state, parts.sessionId).forEach(function (r) { if (r.day && list.indexOf(r.day) === -1) list.push(r.day); });
    days[parts.sessionId] = list.sort();
  });
  return days;
}

/** RPC of the Configure Sessions dialog. */
function saveAdhocSessionsConfiguration(rows) {
  assertNotTemplateMaster_();
  let attendanceSessionIds = [];
  // The days that have attendance, per session: null when that cannot be told.
  let attendanceDays = {};
  try {
    attendanceSessionIds = adhocAttendanceSessionIds_(adhocAttendanceStore_());
    const attendance = readAdhocAttendance_(adhocAttendanceStore_());
    attendanceDays = attendance.status === 'unsupported' ? null : adhocAttendanceDaysBySession_(attendance);
  } catch (e) {
    attendanceSessionIds = null;
    attendanceDays = null;
  }
  let tdocSessionReferences = {};
  try {
    tdocSessionReferences = adhocTdocSessionReferences_(getReportStateStore_().getProperty(ADHOC_TDOC_SESSIONS_KEY_));
  } catch (e) {
    tdocSessionReferences = null;
  }
  let openingSessionIds = [];
  try {
    openingSessionIds = adhocOpeningSessionIds_(adhocOpeningStore_().getProperty(ADHOC_OPENING_KEY_));
  } catch (e) {
    openingSessionIds = null;
  }
  const result = saveAdhocSessionsWith_(getReportStateStore_(), isAdhocMeetingForSessions_(), rows, attendanceSessionIds, tdocSessionReferences, openingSessionIds, attendanceDays);

  // Stage E: the Session administration section says when each session is,
  // so it follows the sessions -- in a report that has the section, or has
  // opening details stored. A report that has neither is not touched.
  if (result.ok && result.changed) {
    let relevant = !!(openingSessionIds && openingSessionIds.length);
    try {
      relevant = relevant || !!findAdhocOpeningContainer_(DocumentApp.getActiveDocument().getBody());
    } catch (e) {
      // Without a document to look at there is nothing to bring in line.
    }
    if (relevant) {
      const openingNotice = renderAdhocOpeningAfterChange_('The sessions were saved.');
      if (openingNotice) result.notice = openingNotice;
    }
  }

  // Stage D: the Attendance section names the sessions and follows their
  // order, so it is written again when sessions with attendance changed.
  // The sessions are saved either way; a section that could not be written
  // is reported with the way to write it later.
  if (result.ok && result.changed && attendanceSessionIds && attendanceSessionIds.length && result.sessions.length) {
    const rendered = withAdhocAttendanceLock_(function () {
      try {
        refreshAdhocAttendanceSection_(DocumentApp.getActiveDocument().getBody());
        return { ok: true };
      } catch (e) {
        return { ok: false, error: e.message };
      }
    });
    if (!rendered.ok) {
      result.notice = (result.notice ? result.notice + '\n\n' : '') +
        'The sessions were saved. The Attendance section could not be updated (' + rendered.error + '). Use ' + ADHOC_ATTENDANCE_REFRESH_HINT_ + '.';
    }
  }
  // Stage B: the Session column follows the sessions with the next update of
  // the report (it needs the TDoc list); the dialog says so.
  if (result.ok && result.changed) {
    let hasColumn = false;
    try {
      const found = findRegistrationTable_(DocumentApp.getActiveDocument().getBody());
      hasColumn = !!found && found.columns === 5;
    } catch (e) {
      hasColumn = false;
    }
    if (hasColumn) {
      result.notice = (result.notice ? result.notice + '\n\n' : 'The sessions were saved. ') +
        'The Session column of the registration table follows with the next update: SA4 Report > Report > Update Report Now.';
    }
  }
  return result;
}

/**
 * What the Configure Sessions dialog shows (pure; reads nothing but its
 * arguments, writes nothing). Without a stored configuration one session is
 * proposed from the meeting date; it is stored only if the user saves.
 */
function buildAdhocSessionsDialogModel_(raw, meetingDateText) {
  const stored = parseAdhocSessionsProperty_(raw);
  const model = { sessions: [], notice: '', readOnly: false, labelMaxLength: ADHOC_SESSION_LABEL_MAX_LENGTH_ };
  if (stored.status === 'ok') {
    model.sessions = stored.config.sessions.map(adhocSessionRecord_);
  } else if (stored.status === 'unsupported') {
    model.readOnly = true;
    model.notice = stored.error + ' It was written by a newer release and cannot be changed here.';
  } else if (stored.status === 'invalid') {
    model.notice = 'The stored sessions cannot be read and are not used (' + stored.error + ') Saving replaces them.';
  } else {
    const date = meetingStartDateIso_(meetingDateText);
    if (date) {
      model.sessions = [{ id: '', label: '', date: date, start: '', end: '' }];
      model.notice = 'No sessions are configured yet. One session is proposed from the meeting date; nothing is stored until you save.';
    }
  }
  return model;
}

/** Menu (ad-hoc template report): Sessions and Attendance > Configure Sessions… */
function configureAdhocSessions() {
  assertNotTemplateMaster_();
  const ui = DocumentApp.getUi();
  if (!isAdhocMeetingForSessions_()) {
    ui.alert('Configure Sessions', 'Sessions are available for ad-hoc reports only.', ui.ButtonSet.OK);
    return;
  }
  const props = getReportStateStore_();
  const model = buildAdhocSessionsDialogModel_(props.getProperty(ADHOC_SESSIONS_KEY_), props.getProperty('MEETING_DATE'));
  // "<" is written as an escape so that no stored text can end the script element.
  const modelJson = JSON.stringify(model).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; font-size: 13px; }
      table { border-collapse: collapse; width: 100%; margin-top: 10px; }
      th { text-align: left; font-size: 12px; padding: 4px 6px 4px 0; }
      td { padding: 3px 6px 3px 0; vertical-align: middle; }
      input { padding: 6px; box-sizing: border-box; width: 100%; }
      button { margin-top: 16px; padding: 9px 18px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      button:disabled { background: #999; cursor: default; }
      button.remove { margin-top: 0; padding: 6px 10px; background: #666; }
      .hint { font-size: 11px; color: #666; margin-top: 4px; }
      #notice { margin-top: 10px; padding: 8px; background: #fff6e0; border-left: 3px solid #c77c00; display: none; }
      #errors { color: #a94442; margin-top: 12px; white-space: pre-wrap; }
    </style>

    <h2>Sessions</h2>
    <div class="hint">The meetings of this ad-hoc series. A session can span several days: enter its last day as End date, or leave End date empty for a session of one day. Times are optional and are meant in the report's time zone: the start time is on the first day, the end time is the final cut-off on the last day. An empty label is shown as the day and month of the dates.</div>
    <div id="notice"></div>

    <table>
      <thead><tr><th>Label</th><th>Start date</th><th>End date</th><th>Start time</th><th>End time (final cut-off)</th><th></th></tr></thead>
      <tbody id="rows"></tbody>
    </table>
    <div class="hint" id="empty"></div>
    <button type="button" id="addBtn" onclick="addSession()">Add session</button>

    <div id="errors"></div>
    <div>
      <button type="button" id="saveBtn" onclick="saveSessions()">Save</button>
      <button type="button" style="background: #666;" onclick="google.script.host.close()">Cancel</button>
    </div>

    <script>
      var MODEL = ${modelJson};
      var rows = [];
      function el(id) { return document.getElementById(id); }
      function field(type, value) {
        var input = document.createElement('input');
        input.type = type;
        input.value = value || '';
        return input;
      }
      function addRow(session) {
        var tr = document.createElement('tr');
        var row = {
          id: session.id || '', tr: tr,
          label: field('text', session.label), date: field('date', session.date), endDate: field('date', session.endDate),
          start: field('time', session.start), end: field('time', session.end)
        };
        row.label.placeholder = 'from the dates';
        row.label.maxLength = MODEL.labelMaxLength;
        row.endDate.title = 'The last day of a session of several days. Leave empty for a session of one day.';
        ['label', 'date', 'endDate', 'start', 'end'].forEach(function (name) {
          var td = document.createElement('td');
          td.appendChild(row[name]);
          tr.appendChild(td);
        });
        var remove = document.createElement('button');
        remove.type = 'button';
        remove.className = 'remove';
        remove.textContent = 'Remove';
        remove.disabled = MODEL.readOnly;
        remove.onclick = function () { removeRow(row); };
        var last = document.createElement('td');
        last.appendChild(remove);
        tr.appendChild(last);
        el('rows').appendChild(tr);
        rows.push(row);
        refresh();
      }
      function removeRow(row) {
        el('rows').removeChild(row.tr);
        rows.splice(rows.indexOf(row), 1);
        refresh();
      }
      function addSession() { addRow({}); }
      function refresh() {
        el('empty').textContent = rows.length ? '' : 'No sessions. Without sessions the report behaves as it does today.';
      }
      function collect() {
        return rows.map(function (r) {
          var out = { id: r.id, label: r.label.value, date: r.date.value, start: r.start.value, end: r.end.value };
          // Sent only for a session of several days, so that a session of one day is saved as it always was.
          if (r.endDate.value && r.endDate.value !== r.date.value) out.endDate = r.endDate.value;
          return out;
        });
      }
      function showErrors(list) { el('errors').textContent = list.join(String.fromCharCode(10)); }
      function saveSessions() {
        el('saveBtn').disabled = true;
        showErrors([]);
        google.script.run
          .withSuccessHandler(function (result) {
            if (result && result.ok) {
              if (result.notice) alert(result.notice);
              google.script.host.close();
              return;
            }
            el('saveBtn').disabled = false;
            showErrors((result && result.errors) || ['The sessions were not saved.']);
          })
          .withFailureHandler(function (error) {
            el('saveBtn').disabled = false;
            showErrors([error && error.message ? error.message : String(error)]);
          })
          .saveAdhocSessionsConfiguration(collect());
      }
      MODEL.sessions.forEach(addRow);
      refresh();
      if (MODEL.notice) {
        el('notice').textContent = MODEL.notice;
        el('notice').style.display = 'block';
      }
      if (MODEL.readOnly) {
        el('addBtn').disabled = true;
        el('saveBtn').disabled = true;
      }
    </script>
  `)
  .setWidth(780)
  .setHeight(480);

  ui.showModalDialog(html, 'Configure Sessions');
}

// =========================================================
// TEAMS ATTENDANCE (stage C) -- PARSER AND NORMALIZED MODEL
// =========================================================
//
// docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md §9-§11. Reads the attendance
// report Microsoft Teams exports for one meeting and returns a normalized
// model. Everything here is pure: no Google service, no property, no
// document, and no log line -- the input is personal data.
//
// Only the sections "Summary" and "Participants" are read. Reading stops at
// the section that follows them, so the content of "In-Meeting Activities"
// and "Meeting Engagement" is never looked at.
//
// The model (parseTeamsAttendanceReport_().attendance):
//
//   schemaVersion  1
//   source         { type: 'teams-attendance', delimiter: 'tab' | 'comma',
//                    dateOrder: 'month-first' | 'day-first' | null }
//
//   summary        WHAT TEAMS STATES. Never replaced by a calculated value.
//     meetingTitle             white space collapsed, otherwise as written
//     attendanceRecords        Teams' "Attended participants" (number)
//     start, end               'YYYY-MM-DDTHH:mm:ss', wall-clock time as
//                              exported; the export has no time zone and
//                              none is assumed
//     durationSeconds          Teams' "Meeting duration"
//     averageAttendanceSeconds Teams' "Average attendance time"
//     raw                      the five values above exactly as written
//     A value Teams did not supply, or that cannot be read, is null.
//
//   participants   NORMALIZED. One entry per attendee, in order of first
//                  appearance; records of the same attendee are merged.
//     name               display name without Teams' own markers
//     company            from a leading "[Company] Name" only, else ''
//     email              the explicit Email value when valid, else ''
//     firstJoin, lastLeave   like summary.start, or null
//     durationSeconds    sum of the merged records' durations, or null
//     records            how many participant rows this entry stands for
//
//   diagnostics    DERIVED. For checking and for the import preview only.
//     counts       participantRows, rowsWithJoinTime, attendees,
//                  mergedRecords, shortRows, external, unverified, dialIn,
//                  withoutEmail
//     calculated   totalDurationSeconds, rowsWithDuration,
//                  meanDurationSeconds, spanSeconds (end minus start)
//     notes        [{ level: 'info' | 'warning', code, message, rows? }];
//                  `rows` are row numbers within the Participants section
//
// Never part of the model: the Participant ID (UPN) column, which is not
// read at all; a telephone number used as a display name; the role; the
// engagement columns; any raw line.

const TEAMS_ATTENDANCE_SCHEMA_VERSION_ = 1;
const TEAMS_ATTENDANCE_MAX_CHARS_ = 2000000;
const TEAMS_ATTENDANCE_MAX_ROWS_ = 5000;
const TEAMS_DIAL_IN_NAME_ = 'Dial-in participant';

/**
 * Decodes the bytes of an export into text. The real export is UTF-16 LE
 * with a byte-order mark; UTF-16 BE and UTF-8 are accepted, and without a
 * mark UTF-16 is recognized by its zero bytes. `bytes` may be signed (Apps
 * Script) or unsigned. Returns { text, encoding }.
 */
function decodeTeamsAttendanceBytes_(bytes) {
  const n = bytes ? bytes.length : 0;
  const at = function (i) { return bytes[i] & 0xFF; };
  let encoding = 'utf-8';
  let start = 0;
  if (n >= 2 && at(0) === 0xFF && at(1) === 0xFE) {
    encoding = 'utf-16le'; start = 2;
  } else if (n >= 2 && at(0) === 0xFE && at(1) === 0xFF) {
    encoding = 'utf-16be'; start = 2;
  } else if (n >= 3 && at(0) === 0xEF && at(1) === 0xBB && at(2) === 0xBF) {
    start = 3;
  } else {
    const pairs = Math.floor(Math.min(n, 1024) / 2);
    let lowFirst = 0;
    let highFirst = 0;
    for (let p = 0; p < pairs; p++) {
      if (at(2 * p + 1) === 0 && at(2 * p) !== 0) lowFirst++;
      if (at(2 * p) === 0 && at(2 * p + 1) !== 0) highFirst++;
    }
    if (pairs >= 2 && lowFirst > pairs / 2) encoding = 'utf-16le';
    else if (pairs >= 2 && highFirst > pairs / 2) encoding = 'utf-16be';
  }

  const units = [];
  if (encoding === 'utf-8') {
    let i = start;
    while (i < n) {
      const b0 = at(i);
      let cp = 0xFFFD;
      let len = 1;
      const cont = function (k) { return i + k < n && (at(i + k) & 0xC0) === 0x80; };
      if (b0 < 0x80) {
        cp = b0;
      } else if (b0 >= 0xC2 && b0 <= 0xDF && cont(1)) {
        cp = ((b0 & 0x1F) << 6) | (at(i + 1) & 0x3F); len = 2;
      } else if (b0 >= 0xE0 && b0 <= 0xEF && cont(1) && cont(2)) {
        cp = ((b0 & 0x0F) << 12) | ((at(i + 1) & 0x3F) << 6) | (at(i + 2) & 0x3F); len = 3;
      } else if (b0 >= 0xF0 && b0 <= 0xF4 && cont(1) && cont(2) && cont(3)) {
        cp = ((b0 & 0x07) << 18) | ((at(i + 1) & 0x3F) << 12) | ((at(i + 2) & 0x3F) << 6) | (at(i + 3) & 0x3F); len = 4;
      }
      if (cp > 0xFFFF) {
        cp -= 0x10000;
        units.push(0xD800 + (cp >> 10), 0xDC00 + (cp & 0x3FF));
      } else {
        units.push(cp);
      }
      i += len;
    }
  } else {
    const low = encoding === 'utf-16le' ? 0 : 1;
    for (let i = start; i + 1 < n; i += 2) units.push(at(i + low) | (at(i + 1 - low) << 8));
  }

  let text = '';
  for (let i = 0; i < units.length; i += 8192) text += String.fromCharCode.apply(null, units.slice(i, i + 8192));
  return { text: text, encoding: encoding };
}

/**
 * Reads delimited text row by row and calls onRow(cells) for each row. A
 * field that begins with a double quote is quoted: it may contain the
 * delimiter and line breaks, and "" stands for one quote; after its closing
 * quote the cell must end. A quote elsewhere is an ordinary character.
 * onRow may return false to stop reading.
 * Returns false when the quoting is broken -- the text ends inside a quoted
 * field, or something follows a closing quote -- else true. Broken quoting
 * is never read leniently: a cell would swallow the rows that follow it.
 */
function scanDelimitedRows_(text, delimiter, onRow) {
  const src = String(text);
  let row = [];
  let field = '';
  let inQuotes = false;
  let atFieldStart = true;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (inQuotes) {
      if (ch !== '"') {
        field += ch;
      } else if (src[i + 1] === '"') {
        field += '"'; i++;
      } else {
        inQuotes = false;
        const next = src[i + 1];
        if (next !== undefined && next !== delimiter && next !== '\n' && next !== '\r') return false;
      }
    } else if (ch === '"' && atFieldStart) {
      inQuotes = true;
      atFieldStart = false;
    } else if (ch === delimiter) {
      row.push(field); field = ''; atFieldStart = true;
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i++;
      row.push(field);
      if (onRow(row) === false) return true;
      row = []; field = ''; atFieldStart = true;
    } else {
      field += ch;
      atFieldStart = false;
    }
  }
  if (inQuotes) return false;
  if (field !== '' || row.length > 0) { row.push(field); onRow(row); }
  return true;
}

/** All rows of a delimited text, or null when its quoting is broken. */
function splitDelimitedRows_(text, delimiter) {
  const rows = [];
  return scanDelimitedRows_(text, delimiter, function (row) { rows.push(row); }) ? rows : null;
}

function teamsTwoDigits_(v) { return (v < 10 ? '0' : '') + v; }

/** "3h 23m 20s", "1h 7s", "45m" -> seconds; null for anything else. */
function parseTeamsDuration_(value) {
  const s = String(value === null || value === undefined ? '' : value).trim();
  const m = s.match(/^(?:(\d+)\s*h)?\s*(?:(\d+)\s*m)?\s*(?:(\d+)\s*s)?$/i);
  if (!s || !m || (m[1] === undefined && m[2] === undefined && m[3] === undefined)) return null;
  return (parseInt(m[1] || '0', 10) * 3600) + (parseInt(m[2] || '0', 10) * 60) + parseInt(m[3] || '0', 10);
}

/** 12200 -> "3:23:20". */
function formatTeamsDuration_(seconds) {
  if (typeof seconds !== 'number' || !isFinite(seconds) || seconds < 0) return '';
  const total = Math.round(seconds);
  return Math.floor(total / 3600) + ':' + teamsTwoDigits_(Math.floor((total % 3600) / 60)) + ':' + teamsTwoDigits_(total % 60);
}

/**
 * The parts of a date and time as Teams writes it, without deciding what
 * they mean: { form: 'slash' | 'dot' | 'iso', a, b, year, time }, or null.
 * 'slash' ("9/22/26, 2:48:05 PM") is month-first or day-first depending on
 * the language of the Teams client; 'dot' is day-first; 'iso' is year-first.
 */
function readTeamsDateTimeParts_(value) {
  const s = String(value === null || value === undefined ? '' : value).replace(/[\s\u00a0\u2009\u202f]+/g, ' ').trim();
  const clock = '(\\d{1,2}):(\\d{2})(?::(\\d{2}))?(?: ?([AaPp])\\.?[Mm]\\.?)?';
  let form = 'slash';
  let m = s.match(new RegExp('^(\\d{1,2})/(\\d{1,2})/(\\d{2}|\\d{4}),? ' + clock + '$'));
  if (!m) { form = 'dot'; m = s.match(new RegExp('^(\\d{1,2})\\.(\\d{1,2})\\.(\\d{2}|\\d{4}),? ' + clock + '$')); }
  if (!m) {
    form = 'iso';
    const iso = s.match(new RegExp('^(\\d{4})-(\\d{2})-(\\d{2})[T, ]+' + clock + '$'));
    if (iso) m = [iso[0], iso[2], iso[3], iso[1], iso[4], iso[5], iso[6], iso[7]];
  }
  if (!m) return null;

  let hour = parseInt(m[4], 10);
  const minute = parseInt(m[5], 10);
  const second = parseInt(m[6] || '0', 10);
  if (m[7]) {
    if (hour < 1 || hour > 12) return null;
    hour = (hour % 12) + (m[7].toLowerCase() === 'p' ? 12 : 0);
  }
  if (hour > 23 || minute > 59 || second > 59) return null;
  return {
    form: form, a: parseInt(m[1], 10), b: parseInt(m[2], 10),
    year: m[3].length === 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10),
    time: teamsTwoDigits_(hour) + ':' + teamsTwoDigits_(minute) + ':' + teamsTwoDigits_(second)
  };
}

/** 'YYYY-MM-DDTHH:mm:ss' for parts read under `order`, or null when that is not a date. */
function teamsDateTimeFromParts_(parts, order) {
  if (!parts) return null;
  let month = parts.a;
  let day = parts.b;
  if (parts.form === 'dot' || (parts.form === 'slash' && order === 'day-first')) { month = parts.b; day = parts.a; }
  const date = parts.year + '-' + teamsTwoDigits_(month) + '-' + teamsTwoDigits_(day);
  return isValidAdhocSessionDate_(date) ? date + 'T' + parts.time : null;
}

/**
 * Decides how the slash-form dates of one export are to be read.
 * `values` are all its date cells, the Summary start time first;
 * `expectedDate` (YYYY-MM-DD, optional) is the date of the session the
 * export is meant for.
 *
 * With an expected date, the reading under which the start is that date is
 * taken; if there is none, the export is not for that session. Without one,
 * a reading is taken only if it is the only one under which every value is
 * a date. Nothing is guessed: two possible readings are an error.
 * Returns { order, error }; order is null when no value needs one.
 */
function resolveTeamsDateOrder_(values, expectedDate, expectedEndDate) {
  // The session the export is meant for: one day, or every day from expectedDate to expectedEndDate.
  const inSession = function (date) { return expectedEndDate ? (date >= expectedDate && date <= expectedEndDate) : date === expectedDate; };
  const parts = values.map(readTeamsDateTimeParts_);
  const start = parts[0];
  const slash = parts.filter(function (p) { return p && p.form === 'slash'; });
  const other = function (order) { return order === 'month-first' ? 'day-first' : 'month-first'; };
  // A reading is ruled out by a value that is a date only under the other one.
  const ruledOut = function (order) {
    return slash.some(function (p) { return teamsDateTimeFromParts_(p, order) === null && teamsDateTimeFromParts_(p, other(order)) !== null; });
  };
  const sameUnderBoth = slash.every(function (p) { return p.a === p.b; });
  const mismatch = function (dates) {
    return { order: null, error: { code: 'SESSION_DATE_MISMATCH',
      message: 'This attendance export is for ' + dates.join(' or ') + ', not for the session ' + (expectedEndDate ? 'of ' + expectedDate + ' to ' + expectedEndDate : 'on ' + expectedDate) + '.' } };
  };
  const decide = function (orders) {
    const left = orders.filter(function (order) { return !ruledOut(order); });
    if (left.length === 1) return { order: left[0], error: null };
    if (left.length === 2 && sameUnderBoth) return { order: 'month-first', error: null };
    return { order: null, error: { code: 'AMBIGUOUS_DATE',
      message: 'The dates of this export can be read month-first or day-first, and nothing in it decides which. Import it for a session with a date.' } };
  };

  if (!start || start.form !== 'slash') {
    // The start is written unambiguously (or cannot be read at all, which the caller reports).
    const at = teamsDateTimeFromParts_(start, null);
    if (expectedDate && at && !inSession(at.slice(0, 10))) return mismatch([at.slice(0, 10)]);
    return slash.length ? decide(['month-first', 'day-first']) : { order: null, error: null };
  }

  const readings = ['month-first', 'day-first']
    .map(function (order) { return { order: order, at: teamsDateTimeFromParts_(start, order) }; })
    .filter(function (r) { return r.at !== null; });
  if (!expectedDate) return decide(readings.map(function (r) { return r.order; }));

  const matching = readings.filter(function (r) { return inSession(r.at.slice(0, 10)); });
  if (matching.length === 0) {
    return mismatch(readings.map(function (r) { return r.at.slice(0, 10); }).filter(function (d, i, list) { return list.indexOf(d) === i; }));
  }
  if (matching.length === 1) return { order: matching[0].order, error: null };
  return decide(matching.map(function (r) { return r.order; }));
}

/** Seconds between two 'YYYY-MM-DDTHH:mm:ss' values (b minus a), or null. */
function teamsSecondsBetween_(a, b) {
  const at = function (v) {
    const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
    return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]) / 1000 : null;
  };
  const x = at(a);
  const y = at(b);
  return x === null || y === null ? null : y - x;
}

/**
 * A Teams display name, conservatively normalized. Removed: surrounding
 * white space, runs of white space, and Teams' own trailing "(External)" /
 * "(Unverified)". A leading "[Company] Name" gives the company and the name.
 * A name that is only a telephone number becomes "Dial-in participant" and
 * the number is dropped here. Everything else stays as the person wrote it:
 * commas, apostrophes, hyphens, other scripts, and any other brackets.
 * Returns { name, company, external, unverified, dialIn }.
 */
function normalizeTeamsDisplayName_(value) {
  let name = String(value === null || value === undefined ? '' : value).replace(/[\s\u00a0\u2009\u202f]+/g, ' ').trim();
  if (name.normalize) name = name.normalize('NFC');
  const result = { name: '', company: '', external: false, unverified: false, dialIn: false };
  let marker;
  while ((marker = name.match(/\s*\((External|Unverified)\)$/i)) !== null) {
    if (marker[1].toLowerCase() === 'external') result.external = true; else result.unverified = true;
    name = name.slice(0, marker.index).trim();
  }
  if (/^\+?[\d\s().*-]+$/.test(name) && (name.match(/\d/g) || []).length >= 5) {
    result.name = TEAMS_DIAL_IN_NAME_;
    result.dialIn = true;
    return result;
  }
  const prefixed = name.match(/^\[([^\[\]]*[^\[\]\s][^\[\]]*)\]\s*(\S.*)$/);
  if (prefixed) {
    result.company = prefixed[1].trim();
    name = prefixed[2].trim();
  }
  result.name = name;
  return result;
}

/** The explicit Email value when it is one usable address, else ''. */
function normalizeTeamsEmail_(value) {
  const s = String(value === null || value === undefined ? '' : value).trim();
  return /^[^\s@<>()\[\],;:"\\]+@[^\s@<>()\[\],;:"\\]+\.[A-Za-z0-9-]{2,}$/.test(s) && s.indexOf('..') === -1 ? s : '';
}

/**
 * What identifies an attendee across records and sessions: the e-mail when
 * there is one, else the name (with the company, when the name carried one).
 * Not part of the model; computed from an entry when it is needed.
 */
function teamsAttendeeKey_(participant) {
  if (participant.email) return 'email:' + participant.email.toLowerCase();
  return 'name:' + participant.name.toLowerCase() + '|' + String(participant.company || '').toLowerCase();
}

/**
 * Parses the text of a Teams attendance export.
 * options.expectedDate: the date (YYYY-MM-DD) of the session it is meant for.
 * options.expectedEndDate: the last day of that session when it spans several
 * days; the export may then be of any day from expectedDate to it.
 * Returns { ok, error, attendance }: on a fatal problem ok is false, error
 * is { code, message } and attendance is null. Everything that is not fatal
 * is a note in attendance.diagnostics. Throws nothing, logs nothing.
 */
function parseTeamsAttendanceReport_(text, options) {
  const fatal = function (code, message) { return { ok: false, error: { code: code, message: message }, attendance: null }; };
  const expectedDate = options && options.expectedDate ? String(options.expectedDate).trim() : '';
  if (expectedDate && !isValidAdhocSessionDate_(expectedDate)) return fatal('INVALID_EXPECTED_DATE', 'The session date is not a date (YYYY-MM-DD).');
  let expectedEndDate = options && options.expectedEndDate ? String(options.expectedEndDate).trim() : '';
  if (expectedEndDate && (!expectedDate || !isValidAdhocSessionDate_(expectedEndDate) || expectedEndDate < expectedDate)) {
    return fatal('INVALID_EXPECTED_DATE', 'The end date of the session is not a date on or after its first day (YYYY-MM-DD).');
  }
  if (expectedEndDate === expectedDate) expectedEndDate = '';
  if (typeof text !== 'string' || text.trim() === '') return fatal('NOT_A_TEAMS_EXPORT', 'The file is empty.');
  if (text.length > TEAMS_ATTENDANCE_MAX_CHARS_) return fatal('TOO_LARGE', 'The file is too large to be a Teams attendance export.');
  const src = text.replace(/^\uFEFF/, '');
  const trimmed = function (v) { return String(v === null || v === undefined ? '' : v).trim(); };

  // ---- rows of the two sections that are read
  const nameColumn = src.match(/^Name(\t|,)/m);
  const delimiter = nameColumn ? nameColumn[1] : (src.indexOf('\t') !== -1 ? '\t' : ',');
  const summaryRows = [];
  const participantRows = [];
  const seen = { summary: false, participants: false };
  let section = null;
  let rowCount = 0;
  let tooMany = false;
  const closed = scanDelimitedRows_(src, delimiter, function (row) {
    if (row.every(function (cell) { return cell.trim() === ''; })) return true;
    const title = row[0].trim().match(/^\d+\.\s+(\S.*)$/);
    if (title && row.slice(1).every(function (cell) { return cell.trim() === ''; })) {
      const name = title[1].trim().toLowerCase();
      if (seen.summary && seen.participants) return false;
      section = name === 'summary' || name === 'participants' ? name : 'other';
      if (section !== 'other') seen[section] = true;
      return true;
    }
    if (++rowCount > TEAMS_ATTENDANCE_MAX_ROWS_) { tooMany = true; return false; }
    if (section === 'summary') summaryRows.push(row);
    else if (section === 'participants') participantRows.push(row);
    return true;
  });
  if (tooMany) return fatal('TOO_LARGE', 'The file has too many rows to be a Teams attendance export.');
  if (!seen.summary && !seen.participants) {
    return fatal(closed ? 'NOT_A_TEAMS_EXPORT' : 'MALFORMED_QUOTING',
      closed ? 'This is not a Teams attendance export: it has no "Summary" and no "Participants" section.'
        : 'The file cannot be read: its quoting is broken.');
  }
  if (!closed) return fatal('MALFORMED_QUOTING', 'The file cannot be read: its quoting is broken.');
  if (!seen.summary) return fatal('SUMMARY_MISSING', 'The export has no "Summary" section.');
  if (!seen.participants) return fatal('PARTICIPANTS_MISSING', 'The export has no "Participants" section.');

  // ---- columns
  const header = (participantRows[0] || []).map(function (cell) { return cell.trim().toLowerCase(); });
  const column = function (name) { return header.indexOf(name); };
  const col = { name: column('name'), join: column('first join'), leave: column('last leave'), duration: column('in-meeting duration'), email: column('email') };
  if (col.name === -1) return fatal('HEADERS_UNUSABLE', 'The "Participants" section has no "Name" column.');
  const dataRows = participantRows.slice(1);

  // ---- summary values, as written
  const stated = {};
  summaryRows.forEach(function (row) {
    const label = trimmed(row[0]).toLowerCase();
    if (label && !Object.prototype.hasOwnProperty.call(stated, label)) stated[label] = trimmed(row[1]);
  });
  const raw = {
    attendedParticipants: stated['attended participants'] || '',
    startTime: stated['start time'] || '',
    endTime: stated['end time'] || '',
    meetingDuration: stated['meeting duration'] || '',
    averageAttendanceTime: stated['average attendance time'] || ''
  };
  if (!raw.startTime) return fatal('SUMMARY_START_MISSING', 'The "Summary" section has no start time.');
  if (!readTeamsDateTimeParts_(raw.startTime)) return fatal('SUMMARY_START_UNREADABLE', 'The start time in the "Summary" section is not written in a form this import reads.');

  // ---- how dates are to be read
  const cell = function (row, index) { return index === -1 || index >= row.length ? '' : trimmed(row[index]); };
  const dateCells = [raw.startTime, raw.endTime];
  dataRows.forEach(function (row) { dateCells.push(cell(row, col.join), cell(row, col.leave)); });
  const resolved = resolveTeamsDateOrder_(dateCells.filter(function (v, i) { return i === 0 || v !== ''; }), expectedDate, expectedEndDate);
  if (resolved.error) return { ok: false, error: resolved.error, attendance: null };
  const readTime = function (value) { return teamsDateTimeFromParts_(readTeamsDateTimeParts_(value), resolved.order); };

  const notes = [];
  const note = function (level, code, message, rows) {
    const entry = { level: level, code: code, message: message };
    if (rows && rows.length) entry.rows = rows;
    notes.push(entry);
  };
  const plural = function (n, one, many) { return n + ' ' + (n === 1 ? one : many); };

  // ---- summary
  const summary = {
    meetingTitle: (stated['meeting title'] || '').replace(/\s+/g, ' ').trim(),
    attendanceRecords: /^\d+$/.test(raw.attendedParticipants) ? parseInt(raw.attendedParticipants, 10) : null,
    start: readTime(raw.startTime),
    end: raw.endTime ? readTime(raw.endTime) : null,
    durationSeconds: parseTeamsDuration_(raw.meetingDuration),
    averageAttendanceSeconds: parseTeamsDuration_(raw.averageAttendanceTime),
    raw: raw
  };
  if (!summary.start) return fatal('SUMMARY_START_UNREADABLE', 'The start time in the "Summary" section is not a date and time.');
  [['attendanceRecords', 'Attended participants'], ['end', 'End time'], ['durationSeconds', 'Meeting duration'], ['averageAttendanceSeconds', 'Average attendance time']]
    .forEach(function (pair) {
      if (summary[pair[0]] === null) note('warning', 'SUMMARY_VALUE_MISSING', 'The Summary has no readable "' + pair[1] + '".');
    });
  if (!summary.meetingTitle) note('warning', 'SUMMARY_VALUE_MISSING', 'The Summary has no "Meeting title".');
  [['join', 'First Join'], ['leave', 'Last Leave'], ['duration', 'In-Meeting Duration'], ['email', 'Email']].forEach(function (pair) {
    if (col[pair[0]] === -1) note('warning', 'MISSING_COLUMN', 'The Participants section has no "' + pair[1] + '" column.');
  });

  // ---- participant records
  const counts = { participantRows: dataRows.length, rowsWithJoinTime: 0, attendees: 0, mergedRecords: 0, shortRows: 0, external: 0, unverified: 0, dialIn: 0, withoutEmail: 0 };
  const flagged = { nameless: [], invalidEmail: [], unreadableTime: [], unreadableDuration: [], incompleteTimes: [] };
  let totalDuration = 0;
  let rowsWithDuration = 0;
  const records = [];
  dataRows.forEach(function (row, index) {
    const rowNumber = index + 1;
    if (row.length < header.length) counts.shortRows++;
    const who = normalizeTeamsDisplayName_(cell(row, col.name));
    if (!who.name) { flagged.nameless.push(rowNumber); return; }
    if (who.external) counts.external++;
    if (who.unverified) counts.unverified++;
    if (who.dialIn) counts.dialIn++;

    const emailCell = cell(row, col.email);
    const email = normalizeTeamsEmail_(emailCell);
    if (emailCell && !email) flagged.invalidEmail.push(rowNumber);

    const joinCell = cell(row, col.join);
    const leaveCell = cell(row, col.leave);
    const join = joinCell ? readTime(joinCell) : null;
    const leave = leaveCell ? readTime(leaveCell) : null;
    if ((joinCell && !join) || (leaveCell && !leave)) flagged.unreadableTime.push(rowNumber);
    if (!join || !leave) flagged.incompleteTimes.push(rowNumber);
    if (join) counts.rowsWithJoinTime++;

    const durationCell = cell(row, col.duration);
    const duration = durationCell ? parseTeamsDuration_(durationCell) : null;
    if (durationCell && duration === null) flagged.unreadableDuration.push(rowNumber);
    if (duration !== null) { totalDuration += duration; rowsWithDuration++; }

    records.push({ row: rowNumber, name: who.name, company: who.company, email: email, join: join, leave: leave, duration: duration, dialIn: who.dialIn });
  });

  // ---- attendees: records of the same person are merged, conservatively.
  // The same e-mail is the same person. Without an e-mail, records are the
  // same person only when the name (and company) are the same, both have
  // their times, and they were never present at the same time -- a guest who
  // left and joined again. Anything less certain stays separate and is noted.
  const groups = [];
  const overlaps = function (a, b) { return a.join < b.leave && b.join < a.leave; };
  records.forEach(function (record) {
    const key = teamsAttendeeKey_(record);
    let target = null;
    if (record.email) {
      target = groups.filter(function (g) { return g.key === key; })[0] || null;
    } else if (!record.dialIn && record.join && record.leave) {
      target = groups.filter(function (g) {
        return g.key === key && g.members.every(function (m) { return m.join && m.leave && !overlaps(m, record); });
      })[0] || null;
    }
    if (target) target.members.push(record);
    else groups.push({ key: key, nameKey: record.name.toLowerCase(), members: [record] });
  });

  const participants = groups.map(function (g) {
    const first = g.members[0];
    const joins = g.members.map(function (m) { return m.join; }).filter(Boolean).sort();
    const leaves = g.members.map(function (m) { return m.leave; }).filter(Boolean).sort();
    const durations = g.members.map(function (m) { return m.duration; }).filter(function (d) { return d !== null; });
    return {
      name: first.name,
      company: first.company,
      email: first.email,
      firstJoin: joins.length ? joins[0] : null,
      lastLeave: leaves.length ? leaves[leaves.length - 1] : null,
      durationSeconds: durations.length ? durations.reduce(function (sum, d) { return sum + d; }, 0) : null,
      records: g.members.length
    };
  });
  counts.attendees = participants.length;
  counts.mergedRecords = records.length - participants.length;
  counts.withoutEmail = participants.filter(function (p) { return !p.email; }).length;

  // ---- notes
  const rowsOf = function (g) { return g.members.map(function (m) { return m.row; }); };
  groups.filter(function (g) { return g.members.length > 1; }).forEach(function (g) {
    note('info', 'RECORDS_MERGED', plural(g.members.length, 'record', 'records') + ' of one attendee were merged.', rowsOf(g));
  });
  const byName = {};
  groups.forEach(function (g) {
    if (g.members[0].dialIn) return;
    (byName[g.nameKey] = byName[g.nameKey] || []).push(g);
  });
  Object.keys(byName).forEach(function (nameKey) {
    if (byName[nameKey].length < 2) return;
    note('warning', 'DUPLICATE_CANDIDATE', 'Records with the same name were kept as separate attendees: nothing in the export shows that they are one person.',
      byName[nameKey].reduce(function (all, g) { return all.concat(rowsOf(g)); }, []).sort(function (a, b) { return a - b; }));
  });
  if (flagged.nameless.length) note('warning', 'ROW_WITHOUT_NAME', plural(flagged.nameless.length, 'participant row has', 'participant rows have') + ' no name and ' + (flagged.nameless.length === 1 ? 'was' : 'were') + ' left out.', flagged.nameless);
  if (flagged.invalidEmail.length) note('warning', 'INVALID_EMAIL', plural(flagged.invalidEmail.length, 'e-mail value is', 'e-mail values are') + ' not a usable address and ' + (flagged.invalidEmail.length === 1 ? 'was' : 'were') + ' left empty.', flagged.invalidEmail);
  if (flagged.unreadableTime.length) note('warning', 'UNREADABLE_TIME', plural(flagged.unreadableTime.length, 'record has', 'records have') + ' a join or leave time that cannot be read.', flagged.unreadableTime);
  if (flagged.unreadableDuration.length) note('warning', 'UNREADABLE_DURATION', plural(flagged.unreadableDuration.length, 'record has', 'records have') + ' a duration that cannot be read.', flagged.unreadableDuration);
  if (flagged.incompleteTimes.length) note('info', 'INCOMPLETE_TIMES', plural(flagged.incompleteTimes.length, 'record has', 'records have') + ' no complete join and leave time; ' + (flagged.incompleteTimes.length === 1 ? 'it is' : 'they are') + ' kept.', flagged.incompleteTimes);
  if (counts.shortRows) note('info', 'SHORT_ROWS', plural(counts.shortRows, 'participant row has', 'participant rows have') + ' fewer cells than the header.');

  const calculated = {
    totalDurationSeconds: totalDuration,
    rowsWithDuration: rowsWithDuration,
    meanDurationSeconds: rowsWithDuration ? Math.round(totalDuration / rowsWithDuration) : null,
    spanSeconds: teamsSecondsBetween_(summary.start, summary.end)
  };
  if (summary.attendanceRecords !== null && summary.attendanceRecords !== counts.participantRows) {
    note('warning', 'RECORD_COUNT_DIFFERS', 'Teams states ' + plural(summary.attendanceRecords, 'attended participant', 'attended participants') + '; the Participants section has ' + plural(counts.participantRows, 'row', 'rows') + '.');
  }
  if (summary.averageAttendanceSeconds !== null && calculated.meanDurationSeconds !== null && summary.averageAttendanceSeconds !== calculated.meanDurationSeconds) {
    note('info', 'AVERAGE_DIFFERS', 'Teams states an average attendance of ' + formatTeamsDuration_(summary.averageAttendanceSeconds) + '; the mean of the listed durations is ' + formatTeamsDuration_(calculated.meanDurationSeconds) + '. The value Teams states is used.');
  }
  if (summary.durationSeconds !== null && calculated.spanSeconds !== null && summary.durationSeconds !== calculated.spanSeconds) {
    note('info', 'DURATION_DIFFERS', 'Teams states a meeting duration of ' + formatTeamsDuration_(summary.durationSeconds) + '; end minus start is ' + formatTeamsDuration_(calculated.spanSeconds) + '. The value Teams states is used.');
  }

  return {
    ok: true,
    error: null,
    attendance: {
      schemaVersion: TEAMS_ATTENDANCE_SCHEMA_VERSION_,
      source: { type: 'teams-attendance', delimiter: delimiter === '\t' ? 'tab' : 'comma', dateOrder: resolved.order },
      summary: summary,
      participants: participants,
      diagnostics: { counts: counts, calculated: calculated, notes: notes }
    }
  };
}

// =========================================================
// AD-HOC ATTENDANCE (stage D) -- IMPORT, PERSISTENCE, RENDERING
// =========================================================
//
// docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md §4, §12, §18. For an ad-hoc
// report with configured sessions, a Teams attendance export can be imported
// for one session; the report then shows, at the end of its closing section,
// one generated "Attendance" container with a statistics block and an
// attendee table per imported session.
//
// STORED (Document Properties of the report, nowhere else):
//
//   ADHOC_SESSION_ATTENDANCE            the index:
//       { v: 1, gen: <counter>, entries: { <id>: { g, n, len, hash } } }
//     <id> is a session id, or "companies" for the manual Company values.
//   ADHOC_SESSION_ATTENDANCE_<id>_<g>_<i>   chunk i of n of generation g
//
//   A value is split because one property holds at most 9 KB; a chunk has at
//   most ADHOC_ATTENDANCE_CHUNK_CHARS_ characters (at most 3 bytes each).
//   A new value is written as a NEW generation: chunks first, then the
//   index, then the old chunks are deleted. Until the index is written the
//   old value stays valid; a value whose chunks do not add up to the length
//   and hash in the index is never used.
//
//   Per session, only what the report is generated from:
//       { t: title, n: records, s: start, e: end, d: duration, a: average,
//         p: [[name, company, email], ...] }
//   -- the Summary values as Teams states them and the normalized attendees
//   of the parser above. Not stored: join / leave times, per-person
//   durations, diagnostics, the file, its name, or anything the parser does
//   not return (Participant ID, telephone number, role, engagement).
//
// Background updates never read any of this, so none of it is an adoption
// key: it stays in the document's own properties.

const ADHOC_ATTENDANCE_KEY_ = 'ADHOC_SESSION_ATTENDANCE';
const ADHOC_ATTENDANCE_SCHEMA_VERSION_ = 1;
const ADHOC_ATTENDANCE_CHUNK_CHARS_ = 2500;
const ADHOC_ATTENDANCE_COMPANIES_ID_ = 'companies';
const ADHOC_ATTENDANCE_HEADING_ = 'Attendance';
const ADHOC_ATTENDEE_TABLE_HEADER_ = ['Name', 'Company', 'Email'];
// The share of the page width each column of an attendee table gets, and the header shading every table of a report has (removeRowHeightAndSpacing()).
const ADHOC_ATTENDEE_TABLE_WIDTHS_ = [0.34, 0.26, 0.40];
const ADHOC_TABLE_HEADER_BACKGROUND_ = '#D9EAF7';
const ADHOC_ATTENDANCE_COMPANY_MAX_LENGTH_ = 200;
const ADHOC_ATTENDANCE_MAX_BASE64_CHARS_ = 4000000;
const ADHOC_ATTENDANCE_REFRESH_HINT_ = 'SA4 Report > Sessions and Attendance > Refresh Attendance Section';

/** Attendance always lives in the document's own properties (see above). */
function adhocAttendanceStore_() {
  return PropertiesService.getDocumentProperties();
}

/** FNV-1a (32 bit) of a text, as 8 hex digits. An integrity check, not a secret. */
function adhocTextHash_(text) {
  let h = 0x811c9dc5;
  const s = String(text);
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return ('0000000' + h.toString(16)).slice(-8);
}

/** Splits a text into chunks of at most `size` characters, never between the two halves of one character. */
function splitAdhocAttendanceChunks_(text, size) {
  const chunks = [];
  let at = 0;
  while (at < text.length) {
    let end = Math.min(text.length, at + size);
    const last = text.charCodeAt(end - 1);
    if (end < text.length && last >= 0xD800 && last <= 0xDBFF && end - at > 1) end--;
    chunks.push(text.slice(at, end));
    at = end;
  }
  return chunks;
}

function adhocAttendanceChunkKey_(id, generation, index) {
  return ADHOC_ATTENDANCE_KEY_ + '_' + id + '_' + generation + '_' + index;
}

/**
 * The index: { status: 'absent' | 'ok' | 'unsupported' | 'invalid', index,
 * error }. Never throws.
 */
function readAdhocAttendanceIndex_(store) {
  const raw = store.getProperty(ADHOC_ATTENDANCE_KEY_);
  if (raw === null || raw === undefined || String(raw).trim() === '') return { status: 'absent', index: null, error: null };
  let data;
  try { data = JSON.parse(String(raw)); } catch (e) { data = null; }
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    return { status: 'invalid', index: null, error: 'The stored attendance index cannot be read.' };
  }
  if (data.v !== ADHOC_ATTENDANCE_SCHEMA_VERSION_) {
    return { status: 'unsupported', index: null, error: 'The stored attendance has version ' + JSON.stringify(data.v) +
      '; this release reads version ' + ADHOC_ATTENDANCE_SCHEMA_VERSION_ + '.' };
  }
  const entries = {};
  const whole = function (v) { return typeof v === 'number' && Math.floor(v) === v && v >= 0; };
  const source = data.entries && typeof data.entries === 'object' && !Array.isArray(data.entries) ? data.entries : null;
  if (!source || !whole(data.gen)) return { status: 'invalid', index: null, error: 'The stored attendance index cannot be read.' };
  const ids = Object.keys(source);
  for (let i = 0; i < ids.length; i++) {
    const e = source[ids[i]];
    if (!/^(s[1-9]\d*(d\d{8})?|companies)$/.test(ids[i]) || !e || !whole(e.g) || !whole(e.n) || !whole(e.len) || typeof e.hash !== 'string' || e.g > data.gen) {
      return { status: 'invalid', index: null, error: 'The stored attendance index cannot be read.' };
    }
    entries[ids[i]] = { g: e.g, n: e.n, len: e.len, hash: e.hash };
  }
  return { status: 'ok', index: { v: ADHOC_ATTENDANCE_SCHEMA_VERSION_, gen: data.gen, entries: entries }, error: null };
}

/** One stored value: { status: 'absent' | 'ok' | 'corrupt', text }. */
function readAdhocAttendanceValue_(store, index, id) {
  const entry = index && index.entries[id];
  if (!entry) return { status: 'absent', text: null };
  let text = '';
  for (let i = 0; i < entry.n; i++) {
    const chunk = store.getProperty(adhocAttendanceChunkKey_(id, entry.g, i));
    if (chunk === null || chunk === undefined) return { status: 'corrupt', text: null };
    text += chunk;
  }
  if (text.length !== entry.len || adhocTextHash_(text) !== entry.hash) return { status: 'corrupt', text: null };
  return { status: 'ok', text: text };
}

/**
 * Stores `text` under `id`, or removes the value when `text` is null.
 * Writes nothing when the stored value already is `text`. Order: new chunks,
 * then the index, then the old chunks -- see the section comment. When the
 * last value is removed, the index and every chunk go too.
 * Returns { changed }. Throws when the stored attendance cannot be extended
 * safely (a newer schema version).
 */
function writeAdhocAttendanceValue_(store, id, text) {
  const read = readAdhocAttendanceIndex_(store);
  if (read.status === 'unsupported') throw new Error(read.error + ' It was written by a newer release and is not changed here.');
  // An index that cannot be read holds nothing usable; a new one replaces it.
  // Its generation counter starts above every chunk that is still stored.
  let index = read.status === 'ok' ? read.index : null;
  const prefix = ADHOC_ATTENDANCE_KEY_ + '_';
  if (!index) {
    let highest = 0;
    store.getKeys().forEach(function (key) {
      const m = key.indexOf(prefix) === 0 ? key.slice(prefix.length).match(/^[a-z0-9]+_(\d+)_\d+$/) : null;
      if (m) highest = Math.max(highest, parseInt(m[1], 10));
    });
    index = { v: ADHOC_ATTENDANCE_SCHEMA_VERSION_, gen: highest, entries: {} };
  }

  const current = readAdhocAttendanceValue_(store, index, id);
  if (text === null ? (current.status === 'absent' && read.status !== 'invalid') : (current.status === 'ok' && current.text === text)) return { changed: false };

  const next = { v: index.v, gen: index.gen, entries: Object.assign({}, index.entries) };
  if (text === null) {
    delete next.entries[id];
  } else {
    const chunks = splitAdhocAttendanceChunks_(text, ADHOC_ATTENDANCE_CHUNK_CHARS_);
    next.gen = index.gen + 1;
    chunks.forEach(function (chunk, i) { store.setProperty(adhocAttendanceChunkKey_(id, next.gen, i), chunk); });
    next.entries[id] = { g: next.gen, n: chunks.length, len: text.length, hash: adhocTextHash_(text) };
  }

  const empty = Object.keys(next.entries).length === 0;
  if (empty) store.deleteProperty(ADHOC_ATTENDANCE_KEY_);
  else store.setProperty(ADHOC_ATTENDANCE_KEY_, JSON.stringify(next));

  // Chunks the index no longer names: the replaced generation, and anything
  // an interrupted earlier write left behind.
  const live = {};
  Object.keys(next.entries).forEach(function (entryId) {
    for (let i = 0; i < next.entries[entryId].n; i++) live[adhocAttendanceChunkKey_(entryId, next.entries[entryId].g, i)] = true;
  });
  store.getKeys().forEach(function (key) {
    if (key.indexOf(prefix) === 0 && !live[key]) store.deleteProperty(key);
  });
  return { changed: true };
}

/** The stored form of one session's attendance: what the report is generated from, and nothing else. */
function compactAdhocAttendance_(attendance) {
  const s = attendance.summary;
  return JSON.stringify({
    t: s.meetingTitle, n: s.attendanceRecords, s: s.start, e: s.end, d: s.durationSeconds, a: s.averageAttendanceSeconds,
    p: attendance.participants.map(function (p) { return [p.name, p.company, p.email]; })
  });
}

/** The reverse of compactAdhocAttendance_(); null for anything that is not that form. */
function expandAdhocAttendance_(text) {
  let d;
  try { d = JSON.parse(text); } catch (e) { return null; }
  const numberOrNull = function (v) { return v === null || (typeof v === 'number' && isFinite(v)); };
  const time = function (v) { return typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(v); };
  if (!d || typeof d !== 'object' || typeof d.t !== 'string' || !time(d.s) || !(d.e === null || time(d.e)) ||
      !numberOrNull(d.n) || !numberOrNull(d.d) || !numberOrNull(d.a) || !Array.isArray(d.p)) return null;
  const attendees = [];
  for (let i = 0; i < d.p.length; i++) {
    const row = d.p[i];
    if (!Array.isArray(row) || row.length !== 3 || typeof row[0] !== 'string' || typeof row[1] !== 'string' || typeof row[2] !== 'string') return null;
    attendees.push({ name: row[0], company: row[1], email: row[2] });
  }
  return { meetingTitle: d.t, attendanceRecords: d.n, start: d.s, end: d.e, durationSeconds: d.d, averageAttendanceSeconds: d.a, attendees: attendees };
}

// ---------------------------------------------------------------
// Attendance records
// ---------------------------------------------------------------
//
// A session may span several days, and Teams exports one attendance report
// per meeting. So a session has attendance RECORDS, one per day: the day is
// the date of the Teams start time of the export. Each record is stored
// under its own key:
//   <session id>               the record of a session of one day -- the
//                              only form there was before sessions could
//                              span several days, and still what a session
//                              of one day writes;
//   <session id>d<YYYYMMDD>    a day of a session of several days.
// A record stored under the plain key belongs to the day its own Teams
// start time says. Importing a day again replaces that day's record only.

/** 's2' -> { sessionId: 's2', day: '' }; 's2d20261027' -> { sessionId: 's2', day: '2026-10-27' }; null for anything else. */
function adhocAttendanceRecordParts_(key) {
  const m = String(key).match(/^(s[1-9]\d*)(?:d(\d{4})(\d{2})(\d{2}))?$/);
  if (!m) return null;
  return { sessionId: m[1], day: m[2] ? m[2] + '-' + m[3] + '-' + m[4] : '' };
}

/** The key of the record of one day of a session of several days. */
function adhocAttendanceRecordKey_(sessionId, day) {
  return sessionId + 'd' + String(day).replace(/-/g, '');
}

/**
 * The records of one session in a readAdhocAttendance_() result, by day:
 * [{ key, day, model }]. `model` is null for a record that cannot be read;
 * `day` is then the one its key names ('' when the key names none).
 */
function adhocAttendanceRecordsOf_(state, sessionId) {
  const records = [];
  Object.keys(state.sessions).concat(state.problems).forEach(function (key) {
    const parts = adhocAttendanceRecordParts_(key);
    if (!parts || parts.sessionId !== sessionId) return;
    const model = state.sessions[key] || null;
    records.push({ key: key, day: model ? model.start.slice(0, 10) : parts.day, model: model });
  });
  return records.sort(function (a, b) { return a.day !== b.day ? (a.day < b.day ? -1 : 1) : (a.key < b.key ? -1 : (a.key > b.key ? 1 : 0)); });
}

/** The session ids that have stored attendance; null when that cannot be told (a newer schema version). */
function adhocAttendanceSessionIds_(store) {
  const read = readAdhocAttendanceIndex_(store);
  if (read.status === 'unsupported') return null;
  if (read.status !== 'ok') return [];
  const ids = [];
  Object.keys(read.index.entries).forEach(function (key) {
    const parts = adhocAttendanceRecordParts_(key);
    if (parts && ids.indexOf(parts.sessionId) === -1) ids.push(parts.sessionId);
  });
  return ids;
}

/**
 * Everything stored: { status, error, sessions: { <record key>: model },
 * raw: { <record key>: text }, problems: [<record key>], companies: { <key>:
 * company } }. A record key is a session id (the record of a session of one
 * day) or <session id>d<YYYYMMDD> (see "Attendance records");
 * adhocAttendanceRecordsOf_() gives the records of one session. `problems`
 * are records whose stored value cannot be used. Never throws.
 */
function readAdhocAttendance_(store) {
  const result = { status: 'ok', error: null, sessions: {}, raw: {}, problems: [], companies: {} };
  const read = readAdhocAttendanceIndex_(store);
  if (read.status !== 'ok') {
    result.status = read.status;
    result.error = read.error;
    return result;
  }
  Object.keys(read.index.entries).forEach(function (id) {
    const value = readAdhocAttendanceValue_(store, read.index, id);
    if (id === ADHOC_ATTENDANCE_COMPANIES_ID_) {
      let map = null;
      try { map = value.status === 'ok' ? JSON.parse(value.text) : null; } catch (e) { map = null; }
      if (map && typeof map === 'object' && !Array.isArray(map)) {
        Object.keys(map).forEach(function (key) { if (typeof map[key] === 'string' && map[key]) result.companies[key] = map[key]; });
      }
      return;
    }
    const model = value.status === 'ok' ? expandAdhocAttendance_(value.text) : null;
    if (!model) { result.problems.push(id); return; }
    result.sessions[id] = model;
    result.raw[id] = value.text;
  });
  return result;
}

// ---------------------------------------------------------------
// Manual Company values
// ---------------------------------------------------------------
//
// Teams supplies no company, so the Company cells of the attendee tables are
// completed by hand. Before the container is generated again, those cells
// are read back and kept as corrections, per attendee:
//   - an attendee with an e-mail is identified by it;
//   - an attendee without one is identified by the exact name, and only
//     where that name stands for one such attendee -- a value is never
//     carried over between two people who merely share a name.
// A correction applies to that attendee in every session. A non-empty cell
// wins over the company the parser found. An empty cell removes a manual
// value, but never the company the parser found in a "[Company] Name".

/** 'email:<address>' or 'name:<name>', both in lower case. */
function adhocAttendeeCompanyKey_(name, email) {
  return email ? 'email:' + String(email).toLowerCase() : 'name:' + String(name).toLowerCase();
}

/** The names that stand for more than one attendee without an e-mail (lower case -> true). */
function ambiguousAdhocAttendeeNames_(attendees) {
  const seen = {};
  const twice = {};
  attendees.forEach(function (a) {
    if (a.email) return;
    const key = String(a.name).toLowerCase();
    if (seen[key]) twice[key] = true;
    seen[key] = true;
  });
  return twice;
}

/** The attendees of one session as they are shown: with the manual Company values applied. */
function effectiveAdhocAttendees_(model, companies) {
  const ambiguous = ambiguousAdhocAttendeeNames_(model.attendees);
  return model.attendees.map(function (a) {
    const usable = a.email || !ambiguous[String(a.name).toLowerCase()];
    const manual = usable ? companies[adhocAttendeeCompanyKey_(a.name, a.email)] : '';
    return { name: a.name, company: manual || a.company, email: a.email };
  });
}

function isAdhocAttendeeTable_(table) {
  try {
    const row = table.getRow(0);
    return row.getNumCells() === ADHOC_ATTENDEE_TABLE_HEADER_.length &&
      ADHOC_ATTENDEE_TABLE_HEADER_.every(function (label, i) { return row.getCell(i).getText().trim() === label; });
  } catch (e) {
    return false;
  }
}

/**
 * The attendee tables whose Company cells are read back: the ones of the
 * generated Attendance section. A table of the same shape anywhere else in
 * the report is somebody's own -- an attendance list made before this
 * feature, for instance -- and is not read once the section exists: its
 * values would otherwise override, at every refresh, what is typed into the
 * generated table. Only while the report has no generated section yet (the
 * first import) are such tables read, once, so that companies entered by
 * hand before are not lost.
 */
function adhocAttendeeTablesToHarvest_(body) {
  const container = findAdhocAttendanceContainer_(body);
  const tables = [];
  if (!container) {
    body.getTables().forEach(function (table) { if (isAdhocAttendeeTable_(table)) tables.push(table); });
    return tables;
  }
  for (let i = container.start; i < container.end; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.TABLE && isAdhocAttendeeTable_(child.asTable())) tables.push(child.asTable());
  }
  return tables;
}

/** How many tables shaped like an attendee table the report has outside the generated Attendance section. Never throws. */
function countAdhocAttendeeTablesElsewhere_(body) {
  try {
    const container = findAdhocAttendanceContainer_(body);
    let count = 0;
    for (let i = 0; i < body.getNumChildren(); i++) {
      if (container && i >= container.start && i < container.end) continue;
      const child = body.getChild(i);
      if (child.getType() === DocumentApp.ElementType.TABLE && isAdhocAttendeeTable_(child.asTable())) count++;
    }
    return count;
  } catch (e) {
    return 0;
  }
}

/** The Company cells of the attendee tables that are read back (adhocAttendeeTablesToHarvest_()): [{ key, company }], ambiguous rows left out. */
function harvestAdhocAttendeeCompanies_(body) {
  const rows = [];
  adhocAttendeeTablesToHarvest_(body).forEach(function (table) {
    const found = [];
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() < 3) continue;
      const name = row.getCell(0).getText().replace(/\s+/g, ' ').trim();
      const email = normalizeTeamsEmail_(row.getCell(2).getText());
      if (!name && !email) continue;
      found.push({ name: name, email: email, company: row.getCell(1).getText().replace(/\s+/g, ' ').trim().slice(0, ADHOC_ATTENDANCE_COMPANY_MAX_LENGTH_) });
    }
    const ambiguous = ambiguousAdhocAttendeeNames_(found);
    found.forEach(function (f) {
      if (!f.email && ambiguous[f.name.toLowerCase()]) return;
      rows.push({ key: adhocAttendeeCompanyKey_(f.name, f.email), company: f.company });
    });
  });
  return rows;
}

/**
 * Pure: the corrections after the document was read. `stored` are the
 * corrections so far, `sessions` the stored attendance, `harvested` the
 * rows of harvestAdhocAttendeeCompanies_(). Corrections of people who are
 * in no stored session are dropped.
 */
function mergeAdhocCompanyCorrections_(stored, sessions, harvested) {
  const parsed = {};   // key -> the companies the parser found for that attendee
  Object.keys(sessions).forEach(function (id) {
    const ambiguous = ambiguousAdhocAttendeeNames_(sessions[id].attendees);
    sessions[id].attendees.forEach(function (a) {
      if (!a.email && ambiguous[String(a.name).toLowerCase()]) return;
      const key = adhocAttendeeCompanyKey_(a.name, a.email);
      parsed[key] = parsed[key] || [];
      if (parsed[key].indexOf(a.company) === -1) parsed[key].push(a.company);
    });
  });

  const cells = {};
  harvested.forEach(function (row) {
    if (!Object.prototype.hasOwnProperty.call(parsed, row.key)) return;
    cells[row.key] = cells[row.key] || [];
    if (cells[row.key].indexOf(row.company) === -1) cells[row.key].push(row.company);
  });

  const result = {};
  Object.keys(parsed).forEach(function (key) {
    const before = stored[key] || '';
    if (!cells[key]) { if (before) result[key] = before; return; }
    const filled = cells[key].filter(function (v) { return v !== ''; });
    let chosen = before;
    if (filled.length === 0) {
      chosen = '';
    } else if (filled.length === 1) {
      chosen = filled[0];
    } else {
      // The same attendee in several tables with different values: the one
      // that was typed is the one that is neither the stored value nor what
      // the parser found. Two different typed values decide nothing.
      const typed = filled.filter(function (v) { return v !== before && parsed[key].indexOf(v) === -1; });
      if (typed.length === 1) chosen = typed[0];
    }
    if (chosen && !(parsed[key].length === 1 && parsed[key][0] === chosen)) result[key] = chosen;
  });
  return result;
}

// ---------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------

/** "September 22, 2026" for 2026-09-22; '' for anything else. */
function adhocLongDate_(date) {
  const text = String(date || '').slice(0, 10);
  return isValidAdhocSessionDate_(text) ? (computeMeetingDateFromStartDate_(text) || '') : '';
}

/**
 * Pure: what is rendered -- one block per attendance record, in the
 * chronological order of the sessions and, within a session, of the days.
 * Statistics are the values Teams states; a value Teams did not state is
 * left out, never calculated. Every block names its session with its days;
 * a block of a session that spans several days (or has several records)
 * also says which day it is.
 * Returns [{ sessionId, key, day, statistics: [[label, value]], attendees }].
 */
function buildAdhocAttendanceBlocks_(sessions, state) {
  const blocks = [];
  sessions.forEach(function (session) {
    const records = adhocAttendanceRecordsOf_(state, session.id).filter(function (r) { return !!r.model; });
    const dates = adhocSessionDatesText_(session);
    records.forEach(function (record) { blocks.push(block(session, record, dates, isAdhocMultiDaySession_(session) || records.length > 1)); });
  });
  return blocks;

  function block(session, record, dates, sayDay) {
    const model = record.model;
    const statistics = [];
    const add = function (label, value) { if (value !== null && value !== undefined && value !== '') statistics.push([label, String(value)]); };
    add('Meeting', model.meetingTitle);
    add('Session', session.label ? session.label + ', ' + dates : dates);
    if (sayDay) add('Date', adhocLongDate_(record.day));
    add('Start', model.start.slice(11));
    if (model.end) add('End', model.end.slice(0, 10) === model.start.slice(0, 10) ? model.end.slice(11) : model.end.replace('T', ' '));
    if (model.durationSeconds !== null) add('Duration', formatTeamsDuration_(model.durationSeconds));
    add('Attendance records', model.attendanceRecords);
    if (model.averageAttendanceSeconds !== null) add('Average attendance', formatTeamsDuration_(model.averageAttendanceSeconds));
    return { sessionId: session.id, key: record.key, day: record.day, statistics: statistics, attendees: effectiveAdhocAttendees_(model, state.companies) };
  }
}

/** The generated container: { start, end } (end exclusive), or null. It runs from its heading to the next heading. */
function findAdhocAttendanceContainer_(body) {
  const PARAGRAPH = DocumentApp.ElementType.PARAGRAPH;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const count = body.getNumChildren();
  for (let i = 0; i < count; i++) {
    const child = body.getChild(i);
    if (child.getType() !== PARAGRAPH) continue;
    const p = child.asParagraph();
    if (p.getHeading() === NORMAL || p.getText().trim() !== ADHOC_ATTENDANCE_HEADING_) continue;
    let end = i + 1;
    while (end < count) {
      const next = body.getChild(end);
      if (next.getType() === PARAGRAPH && next.asParagraph().getHeading() !== NORMAL) break;
      // Stage E: a table that is not an attendee table is never part of the
      // container. An update can put the table of a new TDoc after it (at the
      // end of the document); that table must not be removed with it.
      if (next.getType() === DocumentApp.ElementType.TABLE && !isAdhocAttendeeTable_(next.asTable())) break;
      end++;
    }
    return { start: i, end: end };
  }
  return null;
}

/**
 * Where a new container goes: the end of the closing section -- before the
 * heading that follows the last "Close of ..." heading, or at the end of the
 * document. Without such a heading: the end of the document.
 * Returns { index, inClosingSection }.
 */
function findAdhocAttendanceInsertIndex_(body) {
  const PARAGRAPH = DocumentApp.ElementType.PARAGRAPH;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const count = body.getNumChildren();
  let closing = -1;
  for (let i = 0; i < count; i++) {
    const child = body.getChild(i);
    if (child.getType() !== PARAGRAPH || child.asParagraph().getHeading() === NORMAL) continue;
    const text = child.asParagraph().getText().toLowerCase();
    if (text.indexOf('close of') !== -1 || text.indexOf('closing') !== -1) closing = i;
  }
  if (closing === -1) return { index: count, inClosingSection: false };
  let index = closing + 1;
  while (index < count) {
    const next = body.getChild(index);
    if (next.getType() === PARAGRAPH && next.asParagraph().getHeading() !== NORMAL) break;
    index++;
  }
  return { index: index, inClosingSection: true };
}

/**
 * Removes one child of the body. A document must end with a paragraph, so
 * the last paragraph is emptied instead of removed.
 */
function removeAdhocBodyChild_(body, child) {
  const isLast = body.getChildIndex(child) === body.getNumChildren() - 1;
  if (isLast && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
    child.asParagraph().setText('');
    child.asParagraph().setHeading(DocumentApp.ParagraphHeading.NORMAL);
    return;
  }
  body.removeChild(child);
}

/**
 * The look of a generated attendee table, set explicitly. A table that is
 * inserted takes its text attributes from the paragraph before it -- here
 * the bold "Attendees" line -- so without this the whole table is bold; and
 * a table only gets the report's usual look (no extra row height, no
 * paragraph spacing, a shaded header row) at the next formatting pass. So:
 * header row bold and shaded, every other cell not bold, rows as low as
 * their text, and column widths that give the e-mail addresses room.
 */
function formatAdhocAttendeeTable_(table) {
  let page = 468;
  try { page = parseInt(getConfig_().TDOC_PAGE_USABLE_WIDTH || '468', 10) || 468; } catch (e) { page = 468; }
  const widths = ADHOC_ATTENDEE_TABLE_WIDTHS_.map(function (share) { return Math.round(page * share); });
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    row.setMinimumHeight(0);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.editAsText().setBold(r === 0);
      if (r === 0) cell.setBackgroundColor(ADHOC_TABLE_HEADER_BACKGROUND_);
      if (c < widths.length) cell.setWidth(widths[c]);
      for (let k = 0; k < cell.getNumChildren(); k++) {
        const child = cell.getChild(k);
        if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
        child.asParagraph().setSpacingBefore(0);
        child.asParagraph().setSpacingAfter(0);
      }
    }
  }
}

/**
 * Writes the container for `blocks` and nothing else. An existing container
 * is replaced as a whole (its heading stays where it is); without blocks it
 * is removed. Text outside the container is never touched.
 * Returns { container: 'created' | 'replaced' | 'removed' | 'none', sessions, inClosingSection }.
 */
function renderAdhocAttendanceSection_(body, blocks) {
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const existing = findAdhocAttendanceContainer_(body);
  if (blocks.length === 0) {
    if (!existing) return { container: 'none', sessions: 0, inClosingSection: true };
    for (let i = existing.end - 1; i >= existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));
    return { container: 'removed', sessions: 0, inClosingSection: true };
  }

  let at;
  let inClosingSection = true;
  if (existing) {
    for (let i = existing.end - 1; i > existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));
    at = existing.start + 1;
  } else {
    const where = findAdhocAttendanceInsertIndex_(body);
    inClosingSection = where.inClosingSection;
    const heading = where.index >= body.getNumChildren() ? body.appendParagraph(ADHOC_ATTENDANCE_HEADING_) : body.insertParagraph(where.index, ADHOC_ATTENDANCE_HEADING_);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    at = where.index + 1;
  }

  const paragraph = function (text, bold) {
    const p = at >= body.getNumChildren() ? body.appendParagraph(text) : body.insertParagraph(at, text);
    p.setHeading(NORMAL);
    p.editAsText().setBold(bold);
    at++;
  };
  blocks.forEach(function (block) {
    paragraph('Statistics', true);
    block.statistics.forEach(function (line) { paragraph(line[0] + ': ' + line[1], false); });
    paragraph('Attendees', true);
    const cells = [ADHOC_ATTENDEE_TABLE_HEADER_.slice()].concat(block.attendees.map(function (a) { return [a.name, a.company, a.email]; }));
    const table = at >= body.getNumChildren() ? body.appendTable(cells) : body.insertTable(at, cells);
    at++;
    // The look never decides whether an import succeeds.
    try {
      formatAdhocAttendeeTable_(table);
    } catch (e) {
      Logger.log('Attendance: the attendee table could not be formatted: ' + e.message);
    }
    block.attendees.forEach(function (a, i) {
      // The parser accepts only a plain address, so "mailto:" + address is the whole link.
      if (a.email) table.getRow(i + 1).getCell(2).editAsText().setLinkUrl(0, a.email.length - 1, 'mailto:' + a.email);
    });
  });
  return { container: existing ? 'replaced' : 'created', sessions: blocks.length, inClosingSection: inClosingSection };
}

/**
 * Brings the Attendance container up to date with what is stored: reads the
 * Company cells back, keeps them, and writes the container again.
 * options.afterRebuild: the document was just rebuilt, so there is nothing
 * to read back and sessions whose stored attendance is unusable are skipped
 * (they are named in the result) instead of stopping the render.
 * Returns { container, sessions, inClosingSection, problems }.
 */
function refreshAdhocAttendanceSection_(body, options) {
  const afterRebuild = !!(options && options.afterRebuild);
  const sessions = getAdhocSessions_();
  if (sessions.length === 0) throw new Error('Attendance needs configured sessions. Use Configure Sessions first.');
  const store = adhocAttendanceStore_();
  let state = readAdhocAttendance_(store);
  if (state.status === 'unsupported' || state.status === 'invalid') throw new Error(state.error);
  // A record that cannot be read is named by its session, and by its day when its key says which.
  const labelOf = function (key) {
    const parts = adhocAttendanceRecordParts_(key) || { sessionId: key, day: '' };
    const s = sessions.filter(function (x) { return x.id === parts.sessionId; })[0];
    return (s ? s.displayLabel : parts.sessionId) + (parts.day ? ' (' + adhocLongDate_(parts.day) + ')' : '');
  };
  if (state.problems.length && !afterRebuild) {
    throw new Error('The stored attendance of ' + state.problems.map(labelOf).join(', ') + ' cannot be read. Import it again or remove it; the report was not changed.');
  }

  if (!afterRebuild) {
    const merged = mergeAdhocCompanyCorrections_(state.companies, state.sessions, harvestAdhocAttendeeCompanies_(body));
    writeAdhocAttendanceValue_(store, ADHOC_ATTENDANCE_COMPANIES_ID_, Object.keys(merged).length ? JSON.stringify(merged) : null);
    state.companies = merged;
  }
  const rendered = renderAdhocAttendanceSection_(body, buildAdhocAttendanceBlocks_(sessions, state));
  rendered.problems = state.problems.map(labelOf);
  return rendered;
}

/**
 * Build from Scratch, before the document is cleared: keeps the Company
 * cells of the attendee tables that are about to disappear. null -- and
 * nothing read or written -- unless this is an ad-hoc report with sessions
 * and stored attendance.
 */
function beginAdhocAttendanceRebuild_() {
  if (!adhocSessionsEnabled_()) return null;
  try {
    const store = adhocAttendanceStore_();
    const state = readAdhocAttendance_(store);
    if (state.status !== 'ok' || (Object.keys(state.sessions).length === 0 && state.problems.length === 0)) return null;
    const merged = mergeAdhocCompanyCorrections_(state.companies, state.sessions, harvestAdhocAttendeeCompanies_(DocumentApp.getActiveDocument().getBody()));
    writeAdhocAttendanceValue_(store, ADHOC_ATTENDANCE_COMPANIES_ID_, Object.keys(merged).length ? JSON.stringify(merged) : null);
    return { active: true };
  } catch (e) {
    Logger.log('Attendance: the Company cells could not be kept before the rebuild: ' + e.message);
    return { active: true };
  }
}

/** Build from Scratch, at the end: writes the container again. Returns a note for the build result, '' when all is well. */
function finishAdhocAttendanceRebuild_(body, rebuild) {
  if (!rebuild) return '';
  try {
    const result = refreshAdhocAttendanceSection_(body, { afterRebuild: true });
    return result.problems.length
      ? '\n\n⚠️ The stored attendance of ' + result.problems.join(', ') + ' cannot be read and was left out. Import it again.'
      : '';
  } catch (e) {
    Logger.log('Attendance: the section could not be written after the rebuild: ' + e.message);
    return '\n\n⚠️ The Attendance section could not be written (' + e.message + '). Use ' + ADHOC_ATTENDANCE_REFRESH_HINT_ + '.';
  }
}

// ---------------------------------------------------------------
// Import
// ---------------------------------------------------------------

/**
 * What an import of `text` for session `sessionId` would do. Reads only;
 * writes nothing. Returns { ok, error, plan }: plan = { session, key, day,
 * payload, token, unchanged, replaces, preview }. The export must be of a
 * day of the session; `day` is that day and `key` the record it is stored
 * as -- the record of that day when there is one, so that importing a day
 * again replaces that day and no other.
 */
function planTeamsAttendanceImport_(sessionId, text) {
  const refuse = function (message) { return { ok: false, error: message, plan: null }; };
  if (!isAdhocMeetingForSessions_()) return refuse('Attendance is available for ad-hoc reports only.');
  const sessions = getAdhocSessions_();
  if (sessions.length === 0) return refuse('Attendance needs configured sessions. Use Configure Sessions first.');
  const session = sessions.filter(function (s) { return s.id === sessionId; })[0];
  if (!session) return refuse('This session is no longer configured. Close the dialog and open it again.');

  const state = readAdhocAttendance_(adhocAttendanceStore_());
  if (state.status === 'unsupported') return refuse(state.error + ' It was written by a newer release and is not changed here.');

  const multiDay = isAdhocMultiDaySession_(session);
  const parsed = parseTeamsAttendanceReport_(text, multiDay ? { expectedDate: session.date, expectedEndDate: session.endDate } : { expectedDate: session.date });
  if (!parsed.ok) return refuse(parsed.error.message);
  const attendance = parsed.attendance;
  const payload = compactAdhocAttendance_(attendance);
  if (!expandAdhocAttendance_(payload)) return refuse('The attendance could not be prepared for storing.');

  const s = attendance.summary;
  const day = s.start.slice(0, 10);
  const records = adhocAttendanceRecordsOf_(state, sessionId);
  // The record this import is: the one of that day; for a session of one day also its record that cannot be read.
  const same = records.filter(function (r) { return r.day === day; })[0] || (multiDay ? null : records.filter(function (r) { return r.key === sessionId; })[0]) || null;
  const key = same ? same.key : (multiDay ? adhocAttendanceRecordKey_(sessionId, day) : sessionId);
  const stored = Object.prototype.hasOwnProperty.call(state.raw, key) ? state.raw[key] : null;
  const existing = stored !== null ? state.sessions[key] : null;
  const unreadable = state.problems.indexOf(key) !== -1;
  const dates = adhocSessionDatesText_(session);
  const others = records.filter(function (r) { return r.key !== key; }).map(function (r) { return r.day ? adhocLongDate_(r.day) : 'a record that cannot be read'; });
  const what = multiDay ? adhocLongDate_(day) + ' of this session' : 'this session';
  const messages = function (level) {
    return attendance.diagnostics.notes.filter(function (n) { return n.level === level; }).map(function (n) {
      return n.message + (n.rows ? ' (row' + (n.rows.length === 1 ? ' ' : 's ') + n.rows.join(', ') + ')' : '');
    });
  };
  return { ok: true, error: null, plan: {
    session: session,
    key: key,
    day: day,
    payload: payload,
    token: adhocTextHash_(sessionId + '\n' + payload),
    unchanged: stored === payload,
    replaces: (stored !== null && stored !== payload) || unreadable,
    preview: {
      session: session.label ? session.label + ', ' + dates : dates,
      meetingTitle: s.meetingTitle,
      date: adhocLongDate_(s.start),
      start: s.start.slice(11),
      end: s.end ? (s.end.slice(0, 10) === s.start.slice(0, 10) ? s.end.slice(11) : s.end.replace('T', ' ')) : '',
      duration: s.durationSeconds !== null ? formatTeamsDuration_(s.durationSeconds) : '',
      attendanceRecords: s.attendanceRecords,
      attendees: attendance.participants.length,
      averageAttendance: s.averageAttendanceSeconds !== null ? formatTeamsDuration_(s.averageAttendanceSeconds) : '',
      warnings: messages('warning'),
      notes: messages('info'),
      replaces: stored === payload ? ''
        : (existing ? 'This replaces the attendance already imported for ' + what + ' (' + existing.attendees.length + ' attendees, "' + existing.meetingTitle + '").'
          : (unreadable ? 'This replaces the attendance stored for ' + what + ', which cannot be read.' : '')),
      // The other days of the session that have attendance: they are not touched by this import.
      otherDays: others.join('; '),
      unchanged: stored === payload
    }
  } };
}

/** The text of an uploaded file. Returns { ok, error, text }. */
function teamsAttendanceTextFromBase64_(base64) {
  if (typeof base64 !== 'string' || base64 === '') return { ok: false, error: 'No file was received.', text: null };
  if (base64.length > ADHOC_ATTENDANCE_MAX_BASE64_CHARS_) return { ok: false, error: 'The file is too large to be a Teams attendance export.', text: null };
  try {
    return { ok: true, error: null, text: decodeTeamsAttendanceBytes_(Utilities.base64Decode(base64)).text };
  } catch (e) {
    return { ok: false, error: 'The file could not be read.', text: null };
  }
}

/** Runs fn while holding the document lock the updates use; a busy lock is an error result. */
function withAdhocAttendanceLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok: false, error: 'Nothing was changed: an update of this report is running right now. Try again in a minute.' };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** Renders after a change that is already stored; a failure is reported, the stored data stays. */
function renderAfterAdhocAttendanceChange_(done) {
  try {
    const body = DocumentApp.getActiveDocument().getBody();
    const rendered = refreshAdhocAttendanceSection_(body);
    // An attendance list the report had before (made by hand, or by another tool) is never changed or removed here: it is only pointed out.
    const elsewhere = rendered.container === 'removed' || rendered.container === 'none' ? 0 : countAdhocAttendeeTablesElsewhere_(body);
    const older = elsewhere === 0 ? '' : ' This report also has ' + (elsewhere === 1 ? 'another Name / Company / Email table' : elsewhere + ' other Name / Company / Email tables') +
      ' outside the generated Attendance section -- an attendance list made earlier. ' + (elsewhere === 1 ? 'It was' : 'They were') + ' not changed' +
      (rendered.container === 'created' ? '; the companies entered there were taken over' : '') + '. If ' + (elsewhere === 1 ? 'it shows' : 'they show') + ' the same attendance, remove ' + (elsewhere === 1 ? 'it' : 'them') + ' by hand.';
    return { ok: true, error: null, changed: true, rendered: true,
      message: done + (rendered.container === 'created' && !rendered.inClosingSection ? ' No closing section was found; the Attendance section is at the end of the document.' : '') + older };
  } catch (e) {
    Logger.log('Attendance: stored, but the section could not be written: ' + e.message);
    return { ok: true, error: null, changed: true, rendered: false,
      message: done + ' It is stored, but the Attendance section could not be written (' + e.message + '). Use ' + ADHOC_ATTENDANCE_REFRESH_HINT_ + '.' };
  }
}

/** RPC of the import dialog: the preview. Changes nothing. */
function previewTeamsAttendanceImport(sessionId, base64) {
  assertNotTemplateMaster_();
  const file = teamsAttendanceTextFromBase64_(base64);
  if (!file.ok) return { ok: false, error: file.error };
  const planned = planTeamsAttendanceImport_(String(sessionId || ''), file.text);
  if (!planned.ok) return { ok: false, error: planned.error };
  return { ok: true, error: null, token: planned.plan.token, preview: planned.plan.preview };
}

/**
 * RPC of the import dialog: the import, after the preview was confirmed.
 * `token` is the preview's; the file is parsed again and must give the same
 * result, so what is stored is what was shown.
 */
function confirmTeamsAttendanceImport(sessionId, base64, token) {
  assertNotTemplateMaster_();
  const file = teamsAttendanceTextFromBase64_(base64);
  if (!file.ok) return { ok: false, error: file.error };
  return withAdhocAttendanceLock_(function () {
    const planned = planTeamsAttendanceImport_(String(sessionId || ''), file.text);
    if (!planned.ok) return { ok: false, error: planned.error };
    const plan = planned.plan;
    // A session of several days is named with the day the attendance is of.
    const whose = plan.session.displayLabel + (isAdhocMultiDaySession_(plan.session) ? ' (' + adhocLongDate_(plan.day) + ')' : '');
    if (plan.token !== token) return { ok: false, error: 'Nothing was imported: the file or the session is not the one that was previewed. Preview again.' };

    if (plan.unchanged) {
      if (findAdhocAttendanceContainer_(DocumentApp.getActiveDocument().getBody())) {
        return { ok: true, error: null, changed: false, rendered: false, message: 'This attendance is already imported for ' + whose + '. Nothing was changed.' };
      }
      const restored = renderAfterAdhocAttendanceChange_('This attendance was already imported for ' + whose + '; the Attendance section was written again.');
      restored.changed = false;
      return restored;
    }
    try {
      writeAdhocAttendanceValue_(adhocAttendanceStore_(), plan.key, plan.payload);
    } catch (e) {
      return { ok: false, error: 'Nothing was imported: the attendance could not be stored (' + e.message + ').' };
    }
    return renderAfterAdhocAttendanceChange_('The attendance of ' + whose + ' was imported (' + plan.preview.attendees + ' attendees).');
  });
}

/**
 * RPC of the import dialog: removes ONE attendance record. The dialog asks
 * first. `day` (YYYY-MM-DD) says which record of the session; it can be
 * left out when the session has one record. The other days of the session
 * are not touched.
 */
function removeTeamsAttendanceImport(sessionId, day) {
  assertNotTemplateMaster_();
  const id = String(sessionId || '');
  const wanted = String(day === null || day === undefined ? '' : day).trim();
  if (!isAdhocMeetingForSessions_()) return { ok: false, error: 'Attendance is available for ad-hoc reports only.' };
  return withAdhocAttendanceLock_(function () {
    const store = adhocAttendanceStore_();
    const ids = adhocAttendanceSessionIds_(store);
    if (ids === null) return { ok: false, error: 'The stored attendance was written by a newer release and is not changed here.' };
    if (ids.indexOf(id) === -1) return { ok: false, error: 'This session has no imported attendance.' };
    const session = getAdhocSessions_().filter(function (s) { return s.id === id; })[0];
    const records = adhocAttendanceRecordsOf_(readAdhocAttendance_(store), id);
    let record = null;
    if (wanted) {
      record = records.filter(function (r) { return r.day === wanted; })[0] || null;
      if (!record) return { ok: false, error: 'This session has no imported attendance for that day.' };
    } else if (records.length === 1) {
      record = records[0];
    } else {
      // Several records and no day named: only a record whose day is not known (one that cannot be read) is meant.
      const unknown = records.filter(function (r) { return !r.day; });
      if (unknown.length !== 1) return { ok: false, error: 'This session has attendance for several days. Choose the day to remove.' };
      record = unknown[0];
    }
    try {
      writeAdhocAttendanceValue_(store, record.key, null);
    } catch (e) {
      return { ok: false, error: 'Nothing was removed (' + e.message + ').' };
    }
    if (getAdhocSessions_().length === 0) return { ok: true, error: null, changed: true, rendered: false, message: 'The attendance was removed.' };
    return renderAfterAdhocAttendanceChange_('The attendance of ' + (session ? session.displayLabel : id) + (records.length > 1 && record.day ? ' (' + adhocLongDate_(record.day) + ')' : '') + ' was removed.');
  });
}

/**
 * Menu: Sessions and Attendance > Refresh Attendance Section. Writes the
 * container again from what is stored, keeping the Company cells. Like
 * Update Report Now: no message when it worked, an error when it did not.
 */
function refreshAdhocAttendanceSection() {
  assertNotTemplateMaster_();
  if (!adhocSessionsEnabled_()) throw new Error('Attendance needs an ad-hoc report with configured sessions. Use Configure Sessions first.');
  const result = withAdhocAttendanceLock_(function () {
    return { ok: true, rendered: refreshAdhocAttendanceSection_(DocumentApp.getActiveDocument().getBody()) };
  });
  if (!result.ok) throw new Error(result.error);
  return result.rendered;
}

/** What the import dialog shows (reads only). */
function buildTeamsAttendanceDialogModel_() {
  const state = readAdhocAttendance_(adhocAttendanceStore_());
  return {
    notice: state.status === 'unsupported' || state.status === 'invalid' ? state.error : '',
    readOnly: state.status === 'unsupported',
    sessions: getAdhocSessions_().map(function (s) {
      const dates = adhocSessionDatesText_(s);
      const records = adhocAttendanceRecordsOf_(state, s.id).map(function (r) {
        return { day: r.day, label: r.day ? adhocLongDate_(r.day) : 'a record that cannot be read',
          text: r.model ? r.model.attendees.length + ' attendees, "' + r.model.meetingTitle + '"' : 'stored, but it cannot be read' };
      });
      // One record of a session of one day is described as it always was; several are listed by their day.
      const plain = records.length === 1 && !isAdhocMultiDaySession_(s);
      return {
        id: s.id,
        label: s.label ? s.label + ', ' + dates : dates,
        imported: records.length === 0 ? '' : (plain ? records[0].text : records.map(function (r) { return r.label + ': ' + r.text; }).join('; ')),
        records: records
      };
    })
  };
}

/** Menu (ad-hoc template report): Sessions and Attendance > Import Teams Attendance… */
function importTeamsAttendance() {
  assertNotTemplateMaster_();
  const ui = DocumentApp.getUi();
  if (!isAdhocMeetingForSessions_()) {
    ui.alert('Import Teams Attendance', 'Attendance is available for ad-hoc reports only.', ui.ButtonSet.OK);
    return;
  }
  if (!adhocSessionsEnabled_()) {
    ui.alert('Import Teams Attendance', 'Attendance is imported for a session. Configure the sessions of this report first:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…', ui.ButtonSet.OK);
    return;
  }
  // "<" is written as an escape so that no stored text can end the script element.
  const modelJson = JSON.stringify(buildTeamsAttendanceDialogModel_()).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; font-size: 13px; }
      label { display: block; margin-top: 14px; font-weight: bold; }
      select, input[type=file] { margin-top: 5px; padding: 6px; box-sizing: border-box; width: 100%; }
      button { margin-top: 16px; padding: 9px 18px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      button:disabled { background: #999; cursor: default; }
      button.grey { background: #666; }
      button.danger { background: #b3261e; }
      table { border-collapse: collapse; margin-top: 10px; }
      td { padding: 2px 12px 2px 0; vertical-align: top; }
      td.k { color: #555; white-space: nowrap; }
      .hint { font-size: 11px; color: #666; margin-top: 4px; }
      .box { margin-top: 10px; padding: 8px; background: #fff6e0; border-left: 3px solid #c77c00; display: none; white-space: pre-wrap; }
      #status { margin-top: 12px; white-space: pre-wrap; }
      #status.error { color: #a94442; }
      #status.done { color: #2d7d2d; }
      #infoNotes { color: #555; margin-top: 8px; white-space: pre-wrap; font-size: 12px; }
    </style>

    <h2>Import Teams Attendance</h2>
    <div id="notice" class="box"></div>

    <label>Session</label>
    <select id="session" onchange="sessionChanged()"></select>
    <div id="existing" class="hint"></div>
    <div id="removeRow" style="display:none">
      <select id="removeDay" style="display:none; width:auto" onchange="cancelRemove()"></select>
      <button type="button" class="danger" id="removeBtn" onclick="askRemove()">Remove attendance…</button>
      <span id="removeConfirm" style="display:none">
        <span id="removeQuestion"></span>
        <button type="button" class="danger" id="removeYes" onclick="removeAttendance()">Yes, remove</button>
        <button type="button" class="grey" id="removeNo" onclick="cancelRemove()">No</button>
      </span>
    </div>

    <label>Teams attendance report (.csv)</label>
    <input type="file" id="file" accept=".csv,text/csv" onchange="fileChanged()">
    <div class="hint">The file is read here and checked before anything is changed. Only the Summary and the Participants are used. For a session of several days, import the report of each day: each is kept as its own day of the session.</div>
    <button type="button" id="previewBtn" onclick="previewImport()" disabled>Preview</button>

    <div id="replaces" class="box"></div>
    <table id="preview"></table>
    <div id="warnings" class="box"></div>
    <div id="infoNotes"></div>

    <div id="status"></div>
    <div>
      <button type="button" id="importBtn" onclick="confirmImport()" style="display:none">Import</button>
      <button type="button" class="grey" id="closeBtn" onclick="google.script.host.close()">Cancel</button>
    </div>

    <script>
      var MODEL = ${modelJson};
      var NL = String.fromCharCode(10);
      var MAX_FILE_BYTES = 3000000;
      var pending = null;   // { sessionId, base64, token } of the preview that is shown
      function el(id) { return document.getElementById(id); }
      function show(id, visible, display) { el(id).style.display = visible ? (display || 'block') : 'none'; }
      function selectedSession() {
        var id = el('session').value;
        for (var i = 0; i < MODEL.sessions.length; i++) if (MODEL.sessions[i].id === id) return MODEL.sessions[i];
        return null;
      }
      function setStatus(text, kind) { el('status').textContent = text || ''; el('status').className = kind || ''; }
      function clearPreview() {
        pending = null;
        while (el('preview').firstChild) el('preview').removeChild(el('preview').firstChild);
        el('warnings').textContent = ''; show('warnings', false);
        el('replaces').textContent = ''; show('replaces', false);
        el('infoNotes').textContent = '';
        show('importBtn', false);
        setStatus('');
      }
      function sessionChanged() {
        clearPreview();
        cancelRemove();
        var s = selectedSession();
        el('existing').textContent = s && s.imported ? 'Imported: ' + s.imported : (s ? 'No attendance is imported for this session yet.' : '');
        show('removeRow', !!(s && s.imported) && !MODEL.readOnly);
        // A session with attendance for several days: the day to remove is chosen.
        while (el('removeDay').firstChild) el('removeDay').removeChild(el('removeDay').firstChild);
        var several = !!(s && s.records && s.records.length > 1);
        if (several) s.records.forEach(function (r) {
          var option = document.createElement('option');
          option.value = r.day;
          option.textContent = r.label;
          el('removeDay').appendChild(option);
        });
        show('removeDay', several, 'inline-block');
        updatePreviewButton();
      }
      function fileChanged() { clearPreview(); updatePreviewButton(); }
      function updatePreviewButton() {
        el('previewBtn').disabled = MODEL.readOnly || !selectedSession() || !(el('file').files && el('file').files.length);
      }
      function toBase64(buffer) {
        var bytes = new Uint8Array(buffer);
        var binary = '';
        for (var i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
        return btoa(binary);
      }
      function previewRow(label, value) {
        if (value === null || value === undefined || value === '') return;
        var tr = document.createElement('tr');
        var k = document.createElement('td');
        k.className = 'k';
        k.textContent = label;
        var v = document.createElement('td');
        v.textContent = String(value);
        tr.appendChild(k);
        tr.appendChild(v);
        el('preview').appendChild(tr);
      }
      function busy(on) {
        el('previewBtn').disabled = on || MODEL.readOnly || !(el('file').files && el('file').files.length);
        el('importBtn').disabled = on;
        el('session').disabled = on;
        el('file').disabled = on;
      }
      function failed(error) { busy(false); setStatus(error && error.message ? error.message : String(error), 'error'); }
      function previewImport() {
        clearPreview();
        var session = selectedSession();
        var file = el('file').files && el('file').files[0];
        if (!session || !file) return;
        if (file.size > MAX_FILE_BYTES) { setStatus('The file is too large to be a Teams attendance export.', 'error'); return; }
        busy(true);
        setStatus('Reading the file...');
        var reader = new FileReader();
        reader.onerror = function () { failed('The file could not be read.'); };
        reader.onload = function () {
          var base64 = toBase64(reader.result);
          google.script.run
            .withSuccessHandler(function (result) { showPreview(session.id, base64, result); })
            .withFailureHandler(failed)
            .previewTeamsAttendanceImport(session.id, base64);
        };
        reader.readAsArrayBuffer(file);
      }
      function showPreview(sessionId, base64, result) {
        busy(false);
        if (!result || !result.ok) { setStatus((result && result.error) || 'The file could not be checked.', 'error'); return; }
        var p = result.preview;
        setStatus('');
        previewRow('Session', p.session);
        previewRow('Teams meeting', p.meetingTitle);
        previewRow('Date', p.date);
        previewRow('Other days already imported', p.otherDays);
        previewRow('Start', p.start);
        previewRow('End', p.end);
        previewRow('Duration', p.duration);
        previewRow('Attendance records (Teams)', p.attendanceRecords);
        previewRow('Attendees in the table', p.attendees);
        previewRow('Average attendance', p.averageAttendance);
        previewRow('Warnings', p.warnings.length);
        if (p.warnings.length) { el('warnings').textContent = p.warnings.join(NL); show('warnings', true); }
        if (p.notes.length) el('infoNotes').textContent = p.notes.join(NL);
        if (p.unchanged) {
          setStatus('This attendance is already imported for this session. Nothing needs to be done.', 'done');
          return;
        }
        if (p.replaces) { el('replaces').textContent = p.replaces; show('replaces', true); }
        pending = { sessionId: sessionId, base64: base64, token: result.token };
        el('importBtn').textContent = p.replaces ? 'Replace the imported attendance' : 'Import';
        show('importBtn', true, 'inline-block');
      }
      function finished(result) {
        if (!result || !result.ok) { failed((result && result.error) || 'Nothing was changed.'); return; }
        // The report was changed: this dialog is finished.
        pending = null;
        ['previewBtn', 'importBtn', 'session', 'file', 'removeBtn', 'removeYes', 'removeNo', 'removeDay'].forEach(function (id) { el(id).disabled = true; });
        show('importBtn', false);
        show('removeRow', false);
        el('closeBtn').textContent = 'Close';
        setStatus(result.message, result.rendered === false && result.changed ? 'error' : 'done');
      }
      function confirmImport() {
        if (!pending) return;
        busy(true);
        setStatus('Importing...');
        google.script.run.withSuccessHandler(finished).withFailureHandler(failed)
          .confirmTeamsAttendanceImport(pending.sessionId, pending.base64, pending.token);
      }
      function askRemove() {
        var s = selectedSession();
        if (!s) return;
        var which = s.records && s.records.length > 1 ? ' for ' + el('removeDay').options[el('removeDay').selectedIndex].textContent : '';
        el('removeQuestion').textContent = 'Remove the imported attendance of ' + s.label + which + ' from the report?';
        show('removeBtn', false);
        show('removeConfirm', true, 'inline');
      }
      function cancelRemove() { show('removeConfirm', false); show('removeBtn', true, 'inline-block'); }
      function removeAttendance() {
        var s = selectedSession();
        if (!s) return;
        busy(true);
        setStatus('Removing...');
        if (s.records && s.records.length > 1) {
          google.script.run.withSuccessHandler(finished).withFailureHandler(failed).removeTeamsAttendanceImport(s.id, el('removeDay').value);
        } else {
          google.script.run.withSuccessHandler(finished).withFailureHandler(failed).removeTeamsAttendanceImport(s.id);
        }
      }
      MODEL.sessions.forEach(function (s) {
        var option = document.createElement('option');
        option.value = s.id;
        option.textContent = s.label + (s.imported ? ' (attendance imported)' : '');
        el('session').appendChild(option);
      });
      if (MODEL.notice) { el('notice').textContent = MODEL.notice; show('notice', true); }
      sessionChanged();
    </script>
  `)
  .setWidth(640)
  .setHeight(620);

  ui.showModalDialog(html, 'Import Teams Attendance');
}

// =========================================================
// AD-HOC TDOC SESSIONS (stage B) -- ASSIGNMENT AND THE SESSION COLUMN
// =========================================================
//
// docs/ADHOC_SESSIONS_ATTENDANCE_DESIGN.md §6-§8, §12.1. In an ad-hoc report
// with configured sessions every registered TDoc has the sessions it belongs
// to; the registration table can show them in a fifth column, "Session".
//
// AUTOMATIC: from the "Uploaded" column of the TDoc list -- the upload date
// and time the Portal exports for each TDoc (verified on real lists: a
// date-time cell, per TDoc number, empty while a TDoc is only reserved). A
// TDoc belongs to the first session whose cut-off is not before its upload;
// the cut-off is the planned end of the session, or the end of its day. At
// most one session; none after the last cut-off, and none without a
// readable upload time. Nothing else is ever used to place a TDoc: not its
// number, its row, its revisions, or when an update first saw it.
//
// THE CLOCK. The list gives its times without a time zone. They are UTC:
// checked against the public file server, where the HTTP Last-Modified time
// (GMT) of a TDoc's file is its "Uploaded" value to within seconds or a few
// minutes -- for meetings in February, May, August and September, held in
// India, Montreal and online, and never an hour or more apart. Each time is
// therefore converted to the report's time zone by Apps Script's own
// time-zone rules ("utc") before it is compared with the session times.
// For a list from another source whose times are already on the sessions'
// clock, a report can say so ("session"); that choice is stored with the
// assignments, and a stored choice is never reinterpreted.
//
// MANUAL, per TDoc number, by stable session id:
//   add  the automatic session and these as well (a TDoc that was taken up
//        again in a later session without a new revision);
//   set  these and no others; an empty list means "no session".
// A TDoc without an entry is automatic.
//
// STORED: one property, ADHOC_TDOC_SESSIONS, read by the update as well:
//   { v: 1, clock: 'session' | 'utc',
//     add: { '<ids>': [TDoc, ...] }, set: { '<ids>': [TDoc, ...] } }
// where <ids> are session ids joined by "+", in ascending order ('' in "set"
// is "no session"). TDocs are grouped by their sessions so that hundreds of
// manual entries stay far below the 9 KB of one property; a save that would
// not fit is refused. Labels and dates are never stored here.
//
// THE COLUMN is opt-in: a report gets it when it is built from scratch with
// sessions configured, or through the explicit action below -- never by an
// update. A report that has it is kept up to date by every update. A release
// before this one does not recognize a five-column registration table.

const ADHOC_TDOC_SESSIONS_KEY_ = 'ADHOC_TDOC_SESSIONS';
const ADHOC_TDOC_SESSIONS_SCHEMA_VERSION_ = 1;
const ADHOC_TDOC_SESSIONS_MAX_CHARS_ = 8000;
const ADHOC_SESSION_COLUMN_HEADER_ = 'Session';
const ADHOC_NO_SESSION_TEXT_ = '–';
const REGISTRATION_TABLE_HEADER_ = ['TDoc', 'Title', 'Source', 'Agenda Item'];
// The share of the page width each column of the registration table gets once it has the Session column
// (TDoc, Title, Source, Agenda Item, Session). The four-column table of every release is left as it is.
// The Session column holds codes ("A01"), so it is narrow and the titles get the room. The TDoc column is wide
// enough for a TDoc number ("S4aP260067") on one line: with 15 % of the page it broke before its last digit.
const ADHOC_REGISTRATION_TABLE_WIDTHS_ = [0.19, 0.42, 0.19, 0.12, 0.08];

/**
 * 'YYYY-MM-DDTHH:mm:ss': the latest upload time that still belongs to the
 * session -- its planned end on its LAST day, or the end of that day.
 */
function adhocSessionCutoff_(session) {
  return adhocSessionLastDay_(session) + 'T' + (session.end ? session.end + ':00' : '23:59:59');
}

/**
 * One value of the "Uploaded" column as a wall-clock time.
 * `display` is the cell as the sheet shows it, `raw` its value, and
 * `formatDate(date)` writes a Date as 'YYYY-MM-DDTHH:mm:ss' in the time zone
 * of the sheet -- used only when the shown text is not in that form.
 * Returns { at, dateOnly, unreadable }: `at` is a date and time; `dateOnly`
 * is set instead when the cell holds a date without a time; an empty cell is
 * all null / false; `unreadable` is true for anything else.
 */
function normalizeTdocUploadValue_(display, raw, formatDate) {
  const read = function (value) {
    const m = String(value === null || value === undefined ? '' : value).trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?)?$/);
    if (!m || !isValidAdhocSessionDate_(m[1] + '-' + m[2] + '-' + m[3])) return null;
    if (m[4] === undefined) return { at: null, dateOnly: m[1] + '-' + m[2] + '-' + m[3], unreadable: false };
    if (+m[4] > 23 || +m[5] > 59 || +(m[6] || '0') > 59) return null;
    return { at: m[1] + '-' + m[2] + '-' + m[3] + 'T' + m[4] + ':' + m[5] + ':' + (m[6] || '00'), dateOnly: null, unreadable: false };
  };
  const shown = read(display);
  if (shown) return shown;
  const isDate = Object.prototype.toString.call(raw) === '[object Date]';
  if (isDate && !isNaN(raw.getTime())) {
    let text = '';
    try { text = formatDate ? formatDate(raw) : ''; } catch (e) { text = ''; }
    return read(text) || { at: null, dateOnly: null, unreadable: true };
  }
  const blank = function (v) { return v === null || v === undefined || String(v).trim() === ''; };
  if (blank(display) && blank(raw)) return { at: null, dateOnly: null, unreadable: false };
  return (!isDate && read(raw)) || { at: null, dateOnly: null, unreadable: true };
}

/**
 * The upload times of a downloaded TDoc list, one per row of `data` (the
 * header row included), or null. null -- and nothing read -- unless the
 * report is ad-hoc with sessions and the list has an "Uploaded" column.
 * Never throws: a list whose times cannot be read gives null.
 */
function readTdocUploadTimesForSessions_(spreadsheet, sheet, headers, data, context) {
  try {
    if (!adhocSessionsEnabled_(context)) return null;
    const col = headers.indexOf('Uploaded');
    if (col === -1) return null;
    // What the sheet shows is the wall-clock time as the list has it; a Date
    // value would first have to be read back in the sheet's own time zone.
    let shown = null;
    try { shown = sheet.getRange(1, col + 1, data.length, 1).getDisplayValues(); } catch (e) { shown = null; }
    let zone = '';
    try { zone = spreadsheet.getSpreadsheetTimeZone() || ''; } catch (e) { zone = ''; }
    const formatDate = function (date) { return zone ? Utilities.formatDate(date, zone, "yyyy-MM-dd'T'HH:mm:ss") : ''; };
    return data.map(function (row, i) {
      return normalizeTdocUploadValue_(shown && shown[i] ? shown[i][0] : '', row[col], formatDate);
    });
  } catch (e) {
    Logger.log('TDoc sessions: the upload times of the TDoc list could not be read: ' + e.message);
    return null;
  }
}

/** A wall-clock time meant as UTC, as the wall-clock time of the report's time zone; '' when that cannot be done. */
function convertUtcWallClockToReportZone_(at) {
  const m = String(at || '').match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
  if (!m) return '';
  const text = Utilities.formatDate(new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6])), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/.test(String(text)) ? String(text) : '';
}

/**
 * Pure: the automatic session of one upload time.
 * `upload` is { at, dateOnly, unreadable } already on the sessions' clock
 * (or undefined when the list has no upload column); `sessions` are the
 * configured sessions in any order.
 * Returns { sessionId, reason }: reason is 'assigned', 'not-uploaded',
 * 'unreadable', 'no-upload-time', 'date-only' or 'after-last'.
 * A date without a time gets a session only when the whole day lies in one
 * session's window -- otherwise it is not known which side of a cut-off the
 * upload was on.
 */
function assignAdhocSession_(upload, sessions) {
  if (!upload) return { sessionId: null, reason: 'no-upload-time' };
  if (upload.unreadable) return { sessionId: null, reason: 'unreadable' };
  if (!upload.at && !upload.dateOnly) return { sessionId: null, reason: 'not-uploaded' };
  const ordered = sortAdhocSessions_(sessions);
  const first = function (at) {
    for (let i = 0; i < ordered.length; i++) if (at <= adhocSessionCutoff_(ordered[i])) return ordered[i].id;
    return null;
  };
  if (upload.at) {
    const id = first(upload.at);
    return { sessionId: id, reason: id ? 'assigned' : 'after-last' };
  }
  const morning = first(upload.dateOnly + 'T00:00:00');
  const night = first(upload.dateOnly + 'T23:59:59');
  if (morning !== night) return { sessionId: null, reason: 'date-only' };
  return { sessionId: morning, reason: morning ? 'assigned' : 'after-last' };
}

/**
 * Pure: the sessions of a TDoc, in the chronological order of the sessions.
 * `override` is { mode: 'add' | 'set', sessions: [ids] } or undefined; ids
 * of sessions that are not configured are left out.
 */
function effectiveAdhocTdocSessions_(automaticId, override, sessions) {
  let wanted = automaticId ? [automaticId] : [];
  if (override && override.mode === 'set') wanted = override.sessions.slice();
  else if (override && override.mode === 'add') wanted = wanted.concat(override.sessions);
  return sortAdhocSessions_(sessions).map(function (s) { return s.id; }).filter(function (id) { return wanted.indexOf(id) !== -1; });
}

/** Pure: what the Session cell shows for these session ids: the codes of the sessions ("A01", "A01, A02"), which the Session administration section explains. */
function adhocTdocSessionCellText_(ids, sessions) {
  const codes = adhocSessionCodes_(sessions);
  const shown = sortAdhocSessions_(sessions).filter(function (s) { return ids.indexOf(s.id) !== -1; }).map(function (s) { return codes[s.id]; });
  return shown.length ? shown.join(', ') : ADHOC_NO_SESSION_TEXT_;
}

/** "s2+s10" for ['s10', 's2']; '' for none. */
function adhocSessionIdsKey_(ids) {
  return ids.slice().sort(function (a, b) { return parseInt(a.slice(1), 10) - parseInt(b.slice(1), 10); }).join('+');
}

/**
 * Pure: checks and normalizes manual assignments.
 * `tdocs` is { <TDoc>: { mode, sessions } }. Returns { errors, tdocs }: TDoc
 * numbers in their canonical spelling, session ids unique and ascending.
 * An "add" without sessions is the same as no entry and is left out.
 */
function normalizeAdhocTdocOverrides_(tdocs) {
  const errors = [];
  const result = {};
  if (!tdocs || typeof tdocs !== 'object' || Array.isArray(tdocs)) return { errors: ['The manual assignments are missing.'], tdocs: {} };
  Object.keys(tdocs).forEach(function (key) {
    const parsed = parseExactSA4DocumentId_(String(key).trim());
    const entry = tdocs[key];
    if (!parsed.isValid) { errors.push('"' + String(key).slice(0, 40) + '" is not a TDoc number.'); return; }
    if (!entry || (entry.mode !== 'add' && entry.mode !== 'set') || !Array.isArray(entry.sessions) ||
        !entry.sessions.every(function (id) { return typeof id === 'string' && /^s[1-9]\d*$/.test(id); })) {
      errors.push(parsed.raw + ': the manual assignment cannot be read.');
      return;
    }
    if (Object.prototype.hasOwnProperty.call(result, parsed.raw)) { errors.push(parsed.raw + ' is assigned twice.'); return; }
    const ids = adhocSessionIdsKey_(entry.sessions.filter(function (id, i, list) { return list.indexOf(id) === i; }));
    if (entry.mode === 'add' && ids === '') return;
    result[parsed.raw] = { mode: entry.mode, sessions: ids ? ids.split('+') : [] };
  });
  return { errors: errors, tdocs: errors.length ? {} : result };
}

/** The one stored form: fixed key order, groups and TDocs in ascending order. */
function serializeAdhocTdocSessions_(model) {
  const groups = { add: {}, set: {} };
  Object.keys(model.tdocs).sort().forEach(function (tdoc) {
    const entry = model.tdocs[tdoc];
    const key = adhocSessionIdsKey_(entry.sessions);
    (groups[entry.mode][key] = groups[entry.mode][key] || []).push(tdoc);
  });
  const ordered = function (group) {
    const out = {};
    Object.keys(group).sort().forEach(function (key) { out[key] = group[key]; });
    return out;
  };
  return JSON.stringify({ v: ADHOC_TDOC_SESSIONS_SCHEMA_VERSION_, clock: model.clock, add: ordered(groups.add), set: ordered(groups.set) });
}

/**
 * Reads an ADHOC_TDOC_SESSIONS value: { status: 'absent' | 'ok' |
 * 'unsupported' | 'invalid', model: { clock, tdocs }, error }. For every
 * status but 'ok' the model is the automatic one: the Portal's clock (UTC)
 * and no manual assignment. Never throws.
 */
function parseAdhocTdocSessionsProperty_(raw) {
  const automatic = function (status, error) { return { status: status, model: { clock: 'utc', tdocs: {} }, error: error || null }; };
  if (raw === null || raw === undefined || String(raw).trim() === '') return automatic('absent');
  let data;
  try { data = JSON.parse(String(raw)); } catch (e) { data = null; }
  if (!data || typeof data !== 'object' || Array.isArray(data)) return automatic('invalid', 'The stored manual TDoc session assignments cannot be read.');
  if (data.v !== ADHOC_TDOC_SESSIONS_SCHEMA_VERSION_) {
    return automatic('unsupported', 'The stored manual TDoc session assignments have version ' + JSON.stringify(data.v) + '; this release reads version ' + ADHOC_TDOC_SESSIONS_SCHEMA_VERSION_ + '.');
  }
  const invalid = automatic('invalid', 'The stored manual TDoc session assignments cannot be read.');
  if (data.clock !== 'session' && data.clock !== 'utc') return invalid;
  const tdocs = {};
  const modes = ['add', 'set'];
  for (let m = 0; m < modes.length; m++) {
    const group = data[modes[m]];
    if (!group || typeof group !== 'object' || Array.isArray(group)) return invalid;
    const keys = Object.keys(group);
    for (let k = 0; k < keys.length; k++) {
      if (!Array.isArray(group[keys[k]]) || !/^(s[1-9]\d*(\+s[1-9]\d*)*)?$/.test(keys[k])) return invalid;
      for (let t = 0; t < group[keys[k]].length; t++) {
        const tdoc = group[keys[k]][t];
        if (typeof tdoc !== 'string' || Object.prototype.hasOwnProperty.call(tdocs, tdoc)) return invalid;
        tdocs[tdoc] = { mode: modes[m], sessions: keys[k] ? keys[k].split('+') : [] };
      }
    }
  }
  const checked = normalizeAdhocTdocOverrides_(tdocs);
  if (checked.errors.length || Object.keys(checked.tdocs).length !== Object.keys(tdocs).length) return invalid;
  return { status: 'ok', model: { clock: data.clock, tdocs: checked.tdocs }, error: null };
}

/** The session ids that manual assignments name; null when that cannot be told (a newer schema version). */
function adhocTdocSessionReferences_(raw) {
  const stored = parseAdhocTdocSessionsProperty_(raw);
  if (stored.status === 'unsupported') return null;
  const used = {};
  Object.keys(stored.model.tdocs).forEach(function (tdoc) {
    stored.model.tdocs[tdoc].sessions.forEach(function (id) { (used[id] = used[id] || []).push(tdoc); });
  });
  return used;
}

/**
 * Saves the manual assignments and the clock into `store`. Everything is
 * validated first; a refused save writes nothing, and so does saving what is
 * stored already (or "nothing manual, the Portal's clock" where nothing is
 * stored). `sessionIds` are the configured sessions.
 * Returns { ok, errors, changed }.
 */
function saveAdhocTdocSessionsWith_(store, sessionIds, input) {
  const raw = store.getProperty(ADHOC_TDOC_SESSIONS_KEY_);
  const stored = parseAdhocTdocSessionsProperty_(raw);
  if (stored.status === 'unsupported') return { ok: false, errors: [stored.error + ' They were written by a newer release and are not changed here.'], changed: false };
  const clock = input && input.clock;
  if (clock !== 'session' && clock !== 'utc') return { ok: false, errors: ['The clock of the upload times is missing. Close the dialog and open it again.'], changed: false };
  const checked = normalizeAdhocTdocOverrides_(input.tdocs);
  const errors = checked.errors.slice();
  Object.keys(checked.tdocs).forEach(function (tdoc) {
    checked.tdocs[tdoc].sessions.forEach(function (id) {
      if (sessionIds.indexOf(id) === -1) errors.push(tdoc + ' is assigned to a session that is not configured. Close the dialog and open it again.');
    });
  });
  if (errors.length) return { ok: false, errors: errors.filter(function (e, i) { return errors.indexOf(e) === i; }), changed: false };

  const serialized = serializeAdhocTdocSessions_({ clock: clock, tdocs: checked.tdocs });
  if (serialized.length > ADHOC_TDOC_SESSIONS_MAX_CHARS_) return { ok: false, errors: ['There are too many manual assignments to store.'], changed: false };
  const nothingToStore = (raw === null || raw === undefined) && clock === 'utc' && Object.keys(checked.tdocs).length === 0;
  const changed = !nothingToStore && serialized !== raw;
  if (changed) store.setProperty(ADHOC_TDOC_SESSIONS_KEY_, serialized);
  return { ok: true, errors: [], changed: changed };
}

/**
 * Everything the Session column is computed from, or null when the report is
 * not an ad-hoc report with sessions. For one TDoc of the downloaded list:
 *   uploaded(td)  -> the upload time on the clock of the sessions (for the
 *                    Assign TDoc Sessions dialog only: it is never stored)
 *   automatic(td) -> { sessionId, reason }
 *   effective(td) -> [session ids]
 *   cell(td)      -> the text of the Session cell
 * Manual assignments that cannot be read are ignored (status says so); the
 * automatic result is still given.
 */
function makeAdhocTdocSessionResolver_(context) {
  const sessions = getAdhocSessions_(context);
  if (sessions.length === 0) return null;
  const stored = parseAdhocTdocSessionsProperty_(getReportStateStore_(context).getProperty(ADHOC_TDOC_SESSIONS_KEY_));
  const model = stored.model;
  const onSessionClock = function (upload) {
    if (!upload || model.clock !== 'utc' || upload.unreadable || !upload.at) return upload;
    const at = convertUtcWallClockToReportZone_(upload.at);
    return at ? { at: at, dateOnly: null, unreadable: false } : { at: null, dateOnly: null, unreadable: true };
  };
  // The upload time of a TDoc on the clock of the sessions: THE value the automatic session is decided from.
  const uploaded = function (td) { return onSessionClock(td.uploaded); };
  const automatic = function (td) { return assignAdhocSession_(uploaded(td), sessions); };
  const effective = function (td) { return effectiveAdhocTdocSessions_(automatic(td).sessionId, model.tdocs[tdocNumberOf_(td)], sessions); };
  return {
    sessions: sessions, status: stored.status, error: stored.error, clock: model.clock, overrides: model.tdocs,
    uploaded: uploaded, automatic: automatic, effective: effective,
    cell: function (td) { return adhocTdocSessionCellText_(effective(td), sessions); }
  };
}

/** The resolver for a build or an update: null, and a log line, when it cannot be made. Never throws. */
function makeAdhocTdocSessionResolverSafely_(context) {
  try {
    const resolver = makeAdhocTdocSessionResolver_(context);
    if (resolver && resolver.error) Logger.log('TDoc sessions: ' + resolver.error + ' Only the automatic sessions are used.');
    return resolver;
  } catch (e) {
    Logger.log('TDoc sessions: not available for this run: ' + e.message);
    return null;
  }
}

/**
 * 4 for the registration table as every release writes it (TDoc | Title |
 * Source | Agenda Item), 5 for the same with a Session column, 0 for any
 * other table.
 */
function registrationTableColumns_(table) {
  try {
    if (table.getNumRows() < 1) return 0;
    const row = table.getRow(0);
    const n = row.getNumCells();
    if (n !== 4 && n !== 5) return 0;
    for (let i = 0; i < 4; i++) if (row.getCell(i).getText().trim() !== REGISTRATION_TABLE_HEADER_[i]) return 0;
    if (n === 5 && row.getCell(4).getText().trim() !== ADHOC_SESSION_COLUMN_HEADER_) return 0;
    return n;
  } catch (e) {
    return 0;
  }
}

/**
 * The column widths of a registration table that has the Session column.
 * Without them Google Docs gives five columns a fifth of the page each,
 * which leaves the titles far too little room. Only cells whose width
 * differs are written, so doing it again changes nothing; a table that is
 * not the five-column registration table is not touched. A problem is
 * logged and never thrown: the widths are looks, not content.
 * Returns the number of cells whose width was set.
 */
function applyRegistrationTableWidths_(table) {
  try {
    if (registrationTableColumns_(table) !== 5) return 0;
    let page = 468;
    try { page = parseInt(getConfig_().TDOC_PAGE_USABLE_WIDTH || '468', 10) || 468; } catch (e) { page = 468; }
    const widths = ADHOC_REGISTRATION_TABLE_WIDTHS_.map(function (share) { return Math.round(page * share); });
    let set = 0;
    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      if (row.getNumCells() !== widths.length) continue;
      for (let c = 0; c < widths.length; c++) {
        const cell = row.getCell(c);
        if (cell.getWidth() !== widths[c]) { cell.setWidth(widths[c]); set++; }
      }
    }
    return set;
  } catch (e) {
    Logger.log('TDoc sessions: the column widths of the registration table could not be set: ' + e.message);
    return 0;
  }
}

/** The registration table of the report: { table, columns }, or null. */
function findRegistrationTable_(body) {
  const tables = body.getTables();
  for (let i = 0; i < tables.length; i++) {
    const columns = registrationTableColumns_(tables[i]);
    if (columns) return { table: tables[i], columns: columns };
  }
  return null;
}

/**
 * Brings the Session cells of a five-column registration table up to date;
 * only cells whose text differs are written. A row whose TDoc is not in the
 * list keeps its cell. Does nothing for a four-column table.
 * Returns { rows, changed }.
 */
function refreshRegistrationSessionColumn_(body, allTdocs, resolver) {
  const found = findRegistrationTable_(body);
  if (!found || found.columns !== 5 || !resolver) return { rows: 0, changed: 0 };
  const byNumber = {};
  allTdocs.forEach(function (td) { byNumber[String(td.row[td.tdocCol] || '').trim()] = td; });
  let changed = 0;
  const table = found.table;
  for (let r = 1; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    if (row.getNumCells() < 5) continue;
    const td = byNumber[row.getCell(0).getText().trim()];
    if (!td) continue;
    const text = resolver.cell(td);
    if (row.getCell(4).getText() !== text) { row.getCell(4).setText(text); changed++; }
  }
  // The widths of the five columns, for a table that was made before they were set or whose rows are new.
  applyRegistrationTableWidths_(table);
  return { rows: table.getNumRows() - 1, changed: changed };
}

/** The same for an update: a problem is returned as { step, error } and never thrown; null when all is well. */
function refreshRegistrationSessionColumnSafely_(body, allTdocs, resolver) {
  if (!resolver) return null;
  try {
    refreshRegistrationSessionColumn_(body, allTdocs, resolver);
    return null;
  } catch (e) {
    Logger.log('TDoc sessions: the Session column could not be updated: ' + e.message);
    return { step: 'Session column', error: e.message };
  }
}

/** Every TDoc of a downloaded list, once, in the order of the TDoc numbers. */
function flattenTdocGroups_(tdocGroups) {
  const seen = {};
  const all = [];
  Object.keys(tdocGroups).forEach(function (key) {
    tdocGroups[key].tdocs.forEach(function (td) {
      const number = tdocNumberOf_(td);
      if (!number || seen[number]) return;
      seen[number] = true;
      all.push(Object.assign({}, td, { agendaItem: key }));
    });
  });
  return all.sort(function (a, b) { return tdocNumberOf_(a) < tdocNumberOf_(b) ? -1 : 1; });
}

/**
 * Pure: an upload time as the Assign TDoc Sessions dialog shows it, so that
 * it can be seen why a TDoc has its automatic session. `upload` is what the
 * resolver decides from ({ at, dateOnly, unreadable } on the clock of the
 * sessions). "22 Sep 15:42"; a date without a time "22 Sep"; the neutral
 * sign for a TDoc that is not uploaded or whose time cannot be read.
 *   withYear     the year is added ("22 Sep 2026 15:42"): asked for when the
 *                times and sessions of the dialog are not all of one year
 *   withSeconds  the seconds are shown ("22 Sep 18:00:01"): asked for when
 *                the minute is that of a session's cut-off, where they decide
 * This text is for the dialog only: it is never stored and never written
 * into the report, the status or the statistics.
 */
function adhocUploadedDisplayText_(upload, withYear, withSeconds) {
  if (!upload || upload.unreadable) return ADHOC_NO_SESSION_TEXT_;
  const value = upload.at || upload.dateOnly || '';
  const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}):(\d{2}))?$/);
  if (!m) return ADHOC_NO_SESSION_TEXT_;
  const day = String(parseInt(m[3], 10)) + ' ' + ADHOC_SESSION_MONTHS_[parseInt(m[2], 10) - 1] + (withYear ? ' ' + m[1] : '');
  if (m[4] === undefined) return day;
  return day + ' ' + m[4] + ':' + m[5] + (withSeconds ? ':' + m[6] : '');
}

/** What the Assign TDoc Sessions dialog shows (reads only; `tdocGroups` is a downloaded list). */
function buildAdhocTdocSessionsDialogModel_(resolver, tdocGroups, registrationColumns) {
  const REASONS = {
    'not-uploaded': 'not uploaded yet', 'unreadable': 'upload time cannot be read', 'no-upload-time': 'the TDoc list has no upload times',
    'date-only': 'upload date without a time, on the day of a cut-off', 'after-last': 'uploaded after the last session'
  };
  const listed = {};
  const all = flattenTdocGroups_(tdocGroups);
  // The upload times, on the clock of the sessions -- exactly what the automatic session is decided from.
  const uploads = all.map(function (td) { return resolver.uploaded(td); });
  const valueOf = function (u) { return u && !u.unreadable ? (u.at || u.dateOnly || '') : ''; };
  // The year is shown only when the upload times and the sessions are not all of one year.
  const years = {};
  uploads.forEach(function (u) { if (valueOf(u)) years[valueOf(u).slice(0, 4)] = true; });
  resolver.sessions.forEach(function (s) { years[String(s.date).slice(0, 4)] = true; years[adhocSessionLastDay_(s).slice(0, 4)] = true; });
  const withYear = Object.keys(years).length > 1;
  // The seconds are shown only in the minute of a session's cut-off, where they decide the session.
  const cutoffMinutes = {};
  resolver.sessions.forEach(function (s) { cutoffMinutes[adhocSessionCutoff_(s).slice(0, 16)] = true; });
  const tdocs = all.map(function (td, i) {
    const number = tdocNumberOf_(td);
    const auto = resolver.automatic(td);
    const manual = resolver.overrides[number];
    const upload = uploads[i];
    listed[number] = true;
    return {
      id: number,
      title: String(td.titleCol >= 0 ? td.row[td.titleCol] || '' : '').replace(/\s+/g, ' ').trim().slice(0, 90),
      uploaded: adhocUploadedDisplayText_(upload, withYear, !!(upload && upload.at && cutoffMinutes[upload.at.slice(0, 16)])),
      auto: auto.sessionId || '',
      why: auto.sessionId ? '' : (REASONS[auto.reason] || ''),
      mode: manual ? manual.mode : 'auto',
      sessions: manual ? manual.sessions.slice() : []
    };
  });
  // A manual assignment of a TDoc that is not in the list now is kept, and shown, so that a save does not drop it.
  Object.keys(resolver.overrides).sort().forEach(function (number) {
    if (listed[number]) return;
    tdocs.push({ id: number, title: '(not in the TDoc list now)', uploaded: ADHOC_NO_SESSION_TEXT_, auto: '', why: 'not in the TDoc list', mode: resolver.overrides[number].mode, sessions: resolver.overrides[number].sessions.slice() });
  });
  return {
    sessions: resolver.sessions.map(function (s) { return { id: s.id, label: s.displayLabel }; }),
    clock: resolver.clock,
    // What the Uploaded column shows: the time zone the upload times are converted to, or '' when they are shown as the list has them.
    uploadedZone: resolver.clock === 'utc' ? String(Session.getScriptTimeZone() || '') : '',
    notice: resolver.error ? resolver.error + ' They are not used; saving replaces them.' : '',
    readOnly: resolver.status === 'unsupported',
    column: registrationColumns,
    none: ADHOC_NO_SESSION_TEXT_,
    tdocs: tdocs
  };
}

/** The checks every TDoc-session action starts with; '' when the report can have TDoc sessions. */
function adhocTdocSessionsUnavailable_() {
  if (!isAdhocMeetingForSessions_()) return 'TDoc sessions are available for ad-hoc reports only.';
  if (!adhocSessionsEnabled_()) return 'TDoc sessions need configured sessions. Use Configure Sessions first.';
  return '';
}

/** Downloads the TDoc list and brings the Session column up to date. Returns { ok, error, changed }; never throws. */
function refreshSessionColumnFromList_(body) {
  try {
    const resolver = makeAdhocTdocSessionResolver_();
    const found = findRegistrationTable_(body);
    if (!resolver || !found || found.columns !== 5) return { ok: true, error: null, changed: 0 };
    const all = flattenTdocGroups_(downloadAndGroupTdocs_(getReportConfig_()));
    return { ok: true, error: null, changed: refreshRegistrationSessionColumn_(body, all, resolver).changed };
  } catch (e) {
    return { ok: false, error: e.message, changed: 0 };
  }
}

/** RPC of the Assign TDoc Sessions dialog: Save. `input` is { clock, tdocs: { <TDoc>: { mode, sessions } } }. */
function saveAdhocTdocSessions(input) {
  assertNotTemplateMaster_();
  const unavailable = adhocTdocSessionsUnavailable_();
  if (unavailable) return { ok: false, errors: [unavailable], changed: false };
  const result = saveAdhocTdocSessionsWith_(getReportStateStore_(), getAdhocSessions_().map(function (s) { return s.id; }), input);
  if (!result.ok || !result.changed) return result;

  // The assignments are saved. A report that has the Session column shows them at once.
  const body = DocumentApp.getActiveDocument().getBody();
  const found = findRegistrationTable_(body);
  if (found && found.columns === 5) {
    const refreshed = withAdhocAttendanceLock_(function () { return refreshSessionColumnFromList_(body); });
    if (!refreshed.ok) result.notice = 'The assignments were saved. The Session column could not be updated now (' + refreshed.error + '); the next update of the report does it.';
  }
  return result;
}

/**
 * RPC of the Assign TDoc Sessions dialog: adds the Session column to the
 * registration table of a report that was built without it. The dialog asks
 * first. Doing it again changes nothing.
 */
function enableAdhocSessionColumn() {
  assertNotTemplateMaster_();
  const unavailable = adhocTdocSessionsUnavailable_();
  if (unavailable) return { ok: false, error: unavailable };
  return withAdhocAttendanceLock_(function () {
    const body = DocumentApp.getActiveDocument().getBody();
    const found = findRegistrationTable_(body);
    if (!found) return { ok: false, error: 'This report has no registration table yet. Build the report first; a report built with sessions gets the Session column.' };
    if (found.columns === 5) return { ok: true, error: null, changed: false, message: 'This report already has the Session column.' };
    let resolver;
    let all;
    try {
      resolver = makeAdhocTdocSessionResolver_();
      all = flattenTdocGroups_(downloadAndGroupTdocs_(getReportConfig_()));
    } catch (e) {
      return { ok: false, error: 'Nothing was changed: the TDoc list could not be read (' + e.message + ').' };
    }
    const byNumber = {};
    all.forEach(function (td) { byNumber[String(td.row[td.tdocCol] || '').trim()] = td; });
    const table = found.table;
    // Every cell text is worked out before the first cell is added.
    const texts = [ADHOC_SESSION_COLUMN_HEADER_];
    for (let r = 1; r < table.getNumRows(); r++) {
      const td = byNumber[table.getRow(r).getCell(0).getText().trim()];
      texts.push(td ? resolver.cell(td) : ADHOC_NO_SESSION_TEXT_);
    }
    for (let r = 0; r < table.getNumRows(); r++) table.getRow(r).appendTableCell(texts[r]);
    applyRegistrationTableWidths_(table);
    return { ok: true, error: null, changed: true, message: 'The Session column was added to the registration table (' + (texts.length - 1) + ' TDocs).' };
  });
}

/** Menu (ad-hoc template report): Sessions and Attendance > Assign TDoc Sessions… */
function assignAdhocTdocSessions() {
  assertNotTemplateMaster_();
  const ui = DocumentApp.getUi();
  const unavailable = adhocTdocSessionsUnavailable_();
  if (unavailable) {
    ui.alert('Assign TDoc Sessions', unavailable === 'TDoc sessions are available for ad-hoc reports only.' ? unavailable
      : 'Sessions are assigned to TDocs once the sessions of this report are configured:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…', ui.ButtonSet.OK);
    return;
  }
  let model;
  try {
    const resolver = makeAdhocTdocSessionResolver_();
    const found = findRegistrationTable_(DocumentApp.getActiveDocument().getBody());
    model = buildAdhocTdocSessionsDialogModel_(resolver, downloadAndGroupTdocs_(getReportConfig_()), found ? found.columns : 0);
  } catch (e) {
    ui.alert('Assign TDoc Sessions', 'The TDoc list could not be read:\n\n' + e.message, ui.ButtonSet.OK);
    return;
  }
  // "<" is written as an escape so that no title can end the script element.
  const modelJson = JSON.stringify(model).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      table { border-collapse: collapse; width: 100%; margin-top: 8px; }
      th { text-align: left; font-size: 12px; padding: 4px 8px 4px 0; border-bottom: 1px solid #ccc; }
      td { padding: 3px 8px 3px 0; vertical-align: top; border-bottom: 1px solid #eee; }
      td.title { color: #555; font-size: 12px; max-width: 250px; }
      td.auto { white-space: nowrap; }
      td.up { white-space: nowrap; color: #555; font-size: 12px; }
      label.box { white-space: nowrap; margin-right: 8px; font-weight: normal; }
      select, input[type=text] { padding: 5px; box-sizing: border-box; }
      button { margin-top: 14px; padding: 9px 18px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      button:disabled { background: #999; cursor: default; }
      button.grey { background: #666; }
      .hint { font-size: 11px; color: #666; margin-top: 4px; }
      .box-note { margin-top: 8px; padding: 8px; background: #fff6e0; border-left: 3px solid #c77c00; display: none; white-space: pre-wrap; }
      #rowsWrap { max-height: 330px; overflow-y: auto; border: 1px solid #ddd; margin-top: 8px; padding: 0 6px; }
      #errors { color: #a94442; margin-top: 10px; white-space: pre-wrap; }
      #status { margin-top: 10px; white-space: pre-wrap; }
      .manual { font-weight: bold; }
    </style>

    <h2>TDoc Sessions</h2>
    <div class="hint">Each TDoc belongs automatically to the first session that was not over when it was uploaded. Change a TDoc only where that is not what happened.</div>
    <div class="hint" id="uploadedHint"></div>
    <div id="notice" class="box-note"></div>
    <div id="clockNote" class="box-note"></div>

    <div style="margin-top: 10px">
      <input type="text" id="filter" placeholder="Filter by TDoc number or title" style="width: 60%" oninput="applyFilter()">
      <span id="counts" class="hint"></span>
    </div>
    <div id="rowsWrap">
      <table>
        <thead><tr><th>TDoc</th><th>Title</th><th>Uploaded</th><th>Automatic</th><th>Assignment</th><th>Sessions</th><th>Result</th></tr></thead>
        <tbody id="rows"></tbody>
      </table>
    </div>

    <div id="columnRow" style="margin-top: 12px">
      <span id="columnState"></span>
      <button type="button" class="grey" id="enableBtn" onclick="askEnable()" style="display:none">Add the Session column…</button>
      <span id="enableConfirm" style="display:none">
        <span>Add a Session column to the registration table of this report? A report with this column needs this release or a later one.</span>
        <button type="button" id="enableYes" onclick="enableColumn()">Yes, add it</button>
        <button type="button" class="grey" id="enableNo" onclick="cancelEnable()">No</button>
      </span>
    </div>

    <details style="margin-top: 12px">
      <summary class="hint" style="cursor: pointer">Advanced</summary>
      <div style="margin-top: 6px">
        Upload times in the TDoc list are
        <select id="clock">
          <option value="utc">in UTC, as the 3GPP Portal records them (normal)</option>
          <option value="session">already on the same clock as the session times</option>
        </select>
        <div class="hint">The Portal records upload times in UTC; they are converted to the time zone of this report automatically. Change this only for a TDoc list that does not come from the Portal. A change takes effect when you save.</div>
      </div>
    </details>

    <div id="errors"></div>
    <div id="status"></div>
    <div>
      <button type="button" id="saveBtn" onclick="saveAssignments()">Save</button>
      <button type="button" class="grey" id="closeBtn" onclick="google.script.host.close()">Cancel</button>
    </div>

    <script>
      var MODEL = ${modelJson};
      var NL = String.fromCharCode(10);
      var rows = [];
      function el(id) { return document.getElementById(id); }
      function show(id, visible, display) { el(id).style.display = visible ? (display || 'block') : 'none'; }
      function labelOf(id) {
        for (var i = 0; i < MODEL.sessions.length; i++) if (MODEL.sessions[i].id === id) return MODEL.sessions[i].label;
        return '';
      }
      function chosen(row) {
        var ids = [];
        for (var i = 0; i < MODEL.sessions.length; i++) if (row.boxes[i].checked) ids.push(MODEL.sessions[i].id);
        return ids;
      }
      // The sessions a row results in: automatic, automatic and the ticked ones, or only the ticked ones.
      function effective(row) {
        var mode = row.mode.value;
        var ticked = mode === 'auto' ? [] : chosen(row);
        var ids = [];
        for (var i = 0; i < MODEL.sessions.length; i++) {
          var id = MODEL.sessions[i].id;
          var automatic = mode !== 'set' && row.data.auto === id;
          if (automatic || ticked.indexOf(id) !== -1) ids.push(id);
        }
        return ids;
      }
      function refreshRow(row) {
        var mode = row.mode.value;
        for (var i = 0; i < row.boxes.length; i++) row.boxes[i].disabled = MODEL.readOnly || mode === 'auto';
        var ids = effective(row);
        var labels = [];
        for (var k = 0; k < ids.length; k++) labels.push(labelOf(ids[k]));
        row.result.textContent = labels.length ? labels.join(', ') : MODEL.none;
        row.result.className = mode === 'auto' ? '' : 'manual';
        refreshCounts();
      }
      function refreshCounts() {
        var manual = 0;
        var none = 0;
        for (var i = 0; i < rows.length; i++) {
          if (rows[i].mode.value !== 'auto') manual++;
          if (effective(rows[i]).length === 0) none++;
        }
        el('counts').textContent = rows.length + ' TDocs, ' + manual + ' assigned by hand, ' + none + ' without a session';
      }
      function addRow(data) {
        var tr = document.createElement('tr');
        var cell = function (text, className) {
          var td = document.createElement('td');
          if (className) td.className = className;
          td.textContent = text;
          tr.appendChild(td);
          return td;
        };
        cell(data.id);
        cell(data.title, 'title');
        // Read-only, and not part of what is saved: when the TDoc was uploaded, which is what the automatic session comes from.
        cell(data.uploaded, 'up');
        cell(data.auto ? labelOf(data.auto) : MODEL.none + (data.why ? ' (' + data.why + ')' : ''), 'auto');
        var mode = document.createElement('select');
        [['auto', 'Automatic'], ['add', 'Automatic, and also'], ['set', 'Only these']].forEach(function (pair) {
          var option = document.createElement('option');
          option.value = pair[0];
          option.textContent = pair[1];
          mode.appendChild(option);
        });
        mode.value = data.mode;
        mode.disabled = MODEL.readOnly;
        var modeCell = cell('');
        modeCell.appendChild(mode);
        var boxCell = cell('');
        var boxes = [];
        MODEL.sessions.forEach(function (s) {
          var label = document.createElement('label');
          label.className = 'box';
          var box = document.createElement('input');
          box.type = 'checkbox';
          box.checked = data.sessions.indexOf(s.id) !== -1;
          label.appendChild(box);
          var text = document.createElement('span');
          text.textContent = ' ' + s.label;
          label.appendChild(text);
          boxCell.appendChild(label);
          boxes.push(box);
        });
        var row = { data: data, tr: tr, mode: mode, boxes: boxes, result: cell('') };
        mode.onchange = function () { refreshRow(row); };
        boxes.forEach(function (box) { box.onchange = function () { refreshRow(row); }; });
        el('rows').appendChild(tr);
        rows.push(row);
        refreshRow(row);
      }
      function applyFilter() {
        var wanted = el('filter').value.toLowerCase();
        for (var i = 0; i < rows.length; i++) {
          var text = (rows[i].data.id + ' ' + rows[i].data.title).toLowerCase();
          rows[i].tr.style.display = !wanted || text.indexOf(wanted) !== -1 ? '' : 'none';
        }
      }
      // Only the TDocs that are not automatic are sent.
      function collect() {
        var tdocs = {};
        for (var i = 0; i < rows.length; i++) {
          var mode = rows[i].mode.value;
          if (mode === 'auto') continue;
          var ids = chosen(rows[i]);
          if (mode === 'add' && ids.length === 0) continue;
          tdocs[rows[i].data.id] = { mode: mode, sessions: ids };
        }
        return { clock: el('clock').value, tdocs: tdocs };
      }
      function showErrors(list) { el('errors').textContent = list.join(NL); }
      function failed(error) { el('saveBtn').disabled = MODEL.readOnly; showErrors([error && error.message ? error.message : String(error)]); }
      function saveAssignments() {
        el('saveBtn').disabled = true;
        showErrors([]);
        google.script.run
          .withSuccessHandler(function (result) {
            if (result && result.ok) {
              if (result.notice) alert(result.notice);
              google.script.host.close();
              return;
            }
            el('saveBtn').disabled = false;
            showErrors((result && result.errors) || ['The assignments were not saved.']);
          })
          .withFailureHandler(failed)
          .saveAdhocTdocSessions(collect());
      }
      function showColumnState() {
        el('columnState').textContent = MODEL.column === 5 ? 'This report has the Session column.'
          : (MODEL.column === 4 ? 'This report has no Session column: assignments are stored, but not shown. ' : 'This report has no registration table yet. ');
        show('enableBtn', MODEL.column === 4 && !MODEL.readOnly, 'inline-block');
        show('enableConfirm', false);
      }
      function askEnable() { show('enableBtn', false); show('enableConfirm', true, 'inline'); }
      function cancelEnable() { show('enableConfirm', false); show('enableBtn', true, 'inline-block'); }
      function enableColumn() {
        el('enableYes').disabled = true;
        google.script.run
          .withSuccessHandler(function (result) {
            el('enableYes').disabled = false;
            if (!result || !result.ok) { show('enableConfirm', false); show('enableBtn', true, 'inline-block'); showErrors([(result && result.error) || 'The column was not added.']); return; }
            MODEL.column = 5;
            showColumnState();
            el('status').textContent = result.message;
          })
          .withFailureHandler(function (error) { el('enableYes').disabled = false; failed(error); })
          .enableAdhocSessionColumn();
      }
      el('uploadedHint').textContent = MODEL.uploadedZone
        ? 'Uploaded is the upload time on the 3GPP Portal, shown in the time zone of this report (' + MODEL.uploadedZone + '). It is shown here only and is not stored.'
        : 'Uploaded is the upload time as the TDoc list has it. It is shown here only and is not stored.';
      el('clock').value = MODEL.clock;
      el('clock').disabled = MODEL.readOnly;
      if (MODEL.clock === 'session') {
        el('clockNote').textContent = 'This report compares the upload times as they stand, without converting them from UTC (see Advanced).';
        show('clockNote', true);
      }
      MODEL.tdocs.forEach(addRow);
      refreshCounts();
      showColumnState();
      if (MODEL.notice) { el('notice').textContent = MODEL.notice; show('notice', true); }
      if (MODEL.readOnly) el('saveBtn').disabled = true;
    </script>
  `)
  .setWidth(940)
  .setHeight(640);

  ui.showModalDialog(html, 'Assign TDoc Sessions');
}

// =========================================================
// AD-HOC OPENING (stage E) -- SESSION ADMINISTRATION
// =========================================================
//
// For an ad-hoc report with configured sessions: who chaired each session,
// who took the minutes, and an optional administrative note. These are
// facts the rapporteur types in; nothing here is inferred, summarized or
// matched against the attendance.
//
// STORED (Document Properties of the report), by stable session id:
//   ADHOC_SESSION_OPENING
//     { v: 1, sessions: { <id>: { chair, minuteTakers, note } } }
//   Empty fields and empty sessions are left out; with nothing to store the
//   property is removed. The label, date and times of a session are NOT
//   stored here: they are always taken from the session configuration.
//   Background updates never read this, so it is not an adoption key.
//
// RENDERED as one generated container in the opening part of the report: a
// heading "Session administration" and, per session in chronological order,
// one line saying when it is, then Chair, Minute taker(s) and the note --
// each only when it has a value. The container goes at the end of the
// "Opening of the session" sub-section, before "Registration of Documents",
// and runs to the next heading or table. It is replaced as a whole, so the
// place to change it is the dialog; prose above it is never touched.
//
// The container is written by Build Report from Scratch, by saving the
// dialog, and -- where it exists already or details are stored -- by saving
// the sessions. An update of the report never writes it.

const ADHOC_OPENING_KEY_ = 'ADHOC_SESSION_OPENING';
const ADHOC_OPENING_SCHEMA_VERSION_ = 1;
const ADHOC_OPENING_HEADING_ = 'Session administration';
const ADHOC_OPENING_FIELDS_ = [
  { key: 'chair', label: 'Chair', max: 200, multiline: false },
  { key: 'minuteTakers', label: 'Minute taker(s)', max: 500, multiline: false },
  { key: 'note', label: 'Administrative note', max: 2000, multiline: true }
];
// One property value holds at most 9 KB; free text can take three bytes a character.
const ADHOC_OPENING_MAX_BYTES_ = 8500;

/** Opening details always live in the document's own properties. */
function adhocOpeningStore_() {
  return PropertiesService.getDocumentProperties();
}

/** The number of bytes a text takes when stored as UTF-8. */
function utf8ByteLength_(text) {
  let bytes = 0;
  const s = String(text);
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) bytes += 1;
    else if (c < 0x800) bytes += 2;
    else if (c >= 0xD800 && c <= 0xDBFF && i + 1 < s.length) { bytes += 4; i++; }
    else bytes += 3;
  }
  return bytes;
}

/**
 * One session's details as they are stored: line ends normalized, outer
 * white space removed, and nothing else changed. A single-line field has
 * its line breaks turned into spaces. Returns { chair, minuteTakers, note }.
 */
function normalizeAdhocOpeningEntry_(entry) {
  const out = {};
  ADHOC_OPENING_FIELDS_.forEach(function (field) {
    const raw = entry && typeof entry[field.key] === 'string' ? entry[field.key] : '';
    const text = raw.replace(/\r\n?/g, '\n');
    out[field.key] = (field.multiline ? text : text.replace(/\n+/g, ' ')).trim();
  });
  return out;
}

/** The one stored form, or '' when there is nothing to store. Sessions in ascending order, fields in a fixed order. */
function serializeAdhocOpening_(sessions) {
  const kept = {};
  Object.keys(sessions).sort(function (a, b) { return parseInt(a.slice(1), 10) - parseInt(b.slice(1), 10); }).forEach(function (id) {
    const entry = {};
    ADHOC_OPENING_FIELDS_.forEach(function (field) { if (sessions[id][field.key]) entry[field.key] = sessions[id][field.key]; });
    if (Object.keys(entry).length) kept[id] = entry;
  });
  return Object.keys(kept).length ? JSON.stringify({ v: ADHOC_OPENING_SCHEMA_VERSION_, sessions: kept }) : '';
}

/**
 * Reads an ADHOC_SESSION_OPENING value: { status: 'absent' | 'ok' |
 * 'unsupported' | 'invalid', sessions: { <id>: { chair, minuteTakers, note
 * } }, error }. For every status but 'ok' there are no details. Never throws.
 */
function parseAdhocOpeningProperty_(raw) {
  const none = function (status, error) { return { status: status, sessions: {}, error: error || null }; };
  if (raw === null || raw === undefined || String(raw).trim() === '') return none('absent');
  let data;
  try { data = JSON.parse(String(raw)); } catch (e) { data = null; }
  if (!data || typeof data !== 'object' || Array.isArray(data)) return none('invalid', 'The stored opening details cannot be read.');
  if (data.v !== ADHOC_OPENING_SCHEMA_VERSION_) {
    return none('unsupported', 'The stored opening details have version ' + JSON.stringify(data.v) + '; this release reads version ' + ADHOC_OPENING_SCHEMA_VERSION_ + '.');
  }
  if (!data.sessions || typeof data.sessions !== 'object' || Array.isArray(data.sessions)) return none('invalid', 'The stored opening details cannot be read.');
  const sessions = {};
  const ids = Object.keys(data.sessions);
  for (let i = 0; i < ids.length; i++) {
    const entry = data.sessions[ids[i]];
    if (!/^s[1-9]\d*$/.test(ids[i]) || !entry || typeof entry !== 'object' || Array.isArray(entry)) return none('invalid', 'The stored opening details cannot be read.');
    for (let f = 0; f < ADHOC_OPENING_FIELDS_.length; f++) {
      const value = entry[ADHOC_OPENING_FIELDS_[f].key];
      if (value !== undefined && typeof value !== 'string') return none('invalid', 'The stored opening details cannot be read.');
    }
    const normalized = normalizeAdhocOpeningEntry_(entry);
    if (ADHOC_OPENING_FIELDS_.some(function (field) { return normalized[field.key]; })) sessions[ids[i]] = normalized;
  }
  return { status: 'ok', sessions: sessions, error: null };
}

/** The sessions that have opening details; null when that cannot be told (a newer schema version). */
function adhocOpeningSessionIds_(raw) {
  const stored = parseAdhocOpeningProperty_(raw);
  return stored.status === 'unsupported' ? null : Object.keys(stored.sessions);
}

/**
 * Saves the dialog's details into `store`. `sessions` are the configured
 * sessions; `input` is { <session id>: { chair, minuteTakers, note } }.
 * Everything is validated first; a refused save writes nothing, and so does
 * saving what is stored already. With nothing left to store the property is
 * removed. Messages name a session and a field, never what was typed.
 * Returns { ok, errors, changed }.
 */
function saveAdhocOpeningWith_(store, sessions, input) {
  const raw = store.getProperty(ADHOC_OPENING_KEY_);
  const stored = parseAdhocOpeningProperty_(raw);
  if (stored.status === 'unsupported') return { ok: false, errors: [stored.error + ' They were written by a newer release and are not changed here.'], changed: false };
  if (!input || typeof input !== 'object' || Array.isArray(input)) return { ok: false, errors: ['The opening details are missing.'], changed: false };

  const errors = [];
  const details = {};
  Object.keys(input).forEach(function (id) {
    const session = sessions.filter(function (s) { return s.id === id; })[0];
    if (!session) { errors.push('A session of this dialog is no longer configured. Close the dialog and open it again.'); return; }
    const entry = normalizeAdhocOpeningEntry_(input[id]);
    ADHOC_OPENING_FIELDS_.forEach(function (field) {
      if (entry[field.key].length > field.max) errors.push(adhocSessionLabel_(session) + ': ' + field.label + ' is longer than ' + field.max + ' characters.');
    });
    details[id] = entry;
  });
  if (errors.length) return { ok: false, errors: errors.filter(function (e, i) { return errors.indexOf(e) === i; }), changed: false };

  const serialized = serializeAdhocOpening_(details);
  if (utf8ByteLength_(serialized) > ADHOC_OPENING_MAX_BYTES_) return { ok: false, errors: ['The opening details are too long to store together. Shorten the notes.'], changed: false };
  const absent = raw === null || raw === undefined;
  if (serialized === '') {
    if (absent) return { ok: true, errors: [], changed: false };
    store.deleteProperty(ADHOC_OPENING_KEY_);
    return { ok: true, errors: [], changed: true };
  }
  if (serialized === raw) return { ok: true, errors: [], changed: false };
  store.setProperty(ADHOC_OPENING_KEY_, serialized);
  return { ok: true, errors: [], changed: true };
}

/**
 * "Kick-off, September 22, 2026, 15:00–18:00": when a session is, from its
 * configuration. A session of several days reads "AHG 2, October 26–28,
 * 2026 (from 09:00 on the first day, until 17:00 on the last day)": its
 * times are never written as if they were on one day.
 */
function adhocSessionWhenText_(session) {
  const prefix = session.label ? session.label + ', ' : '';
  if (isAdhocMultiDaySession_(session)) {
    const parts = [];
    if (session.start) parts.push('from ' + session.start + ' on the first day');
    if (session.end) parts.push('until ' + session.end + ' on the last day');
    return prefix + adhocSessionDatesText_(session) + (parts.length ? ' (' + parts.join(', ') + ')' : '');
  }
  const date = adhocLongDate_(session.date);
  const time = session.start && session.end ? ', ' + session.start + '–' + session.end
    : (session.start ? ', ' + session.start : (session.end ? ', until ' + session.end : ''));
  return prefix + date + time;
}

/**
 * Pure: the paragraphs of the container, below its heading -- [{ text, bold
 * }]. Per session, in the order given: its code and when it is ("A01:
 * Kick-off, September 22, 2026, 15:00–18:00" -- the code is what the
 * Session column of the registration table shows); then Chair, Minute
 * taker(s) and the lines of the note, each only when it has a value.
 */
function buildAdhocOpeningLines_(sessions, details) {
  const lines = [];
  const codes = adhocSessionCodes_(sessions);
  sessions.forEach(function (session) {
    const entry = details[session.id] || {};
    lines.push({ text: codes[session.id] + ': ' + adhocSessionWhenText_(session), bold: true });
    if (entry.chair) lines.push({ text: 'Chair: ' + entry.chair, bold: false });
    if (entry.minuteTakers) lines.push({ text: 'Minute taker(s): ' + entry.minuteTakers, bold: false });
    String(entry.note || '').split('\n').forEach(function (line) {
      if (line.trim()) lines.push({ text: line.trim(), bold: false });
    });
  });
  return lines;
}

/**
 * The generated container: { start, end } (end exclusive), or null. It runs
 * from its heading to the next heading or table -- a table is never part of
 * it, so none can be removed with it.
 */
function findAdhocOpeningContainer_(body) {
  const PARAGRAPH = DocumentApp.ElementType.PARAGRAPH;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const count = body.getNumChildren();
  for (let i = 0; i < count; i++) {
    const child = body.getChild(i);
    if (child.getType() !== PARAGRAPH) continue;
    const p = child.asParagraph();
    if (p.getHeading() === NORMAL || p.getText().trim() !== ADHOC_OPENING_HEADING_) continue;
    let end = i + 1;
    while (end < count) {
      const next = body.getChild(end);
      if (next.getType() !== PARAGRAPH || next.asParagraph().getHeading() !== NORMAL) break;
      end++;
    }
    return { start: i, end: end };
  }
  return null;
}

/**
 * Where a new container goes: directly before the "<number> Registration of
 * Documents" heading the build writes, i.e. at the end of the opening
 * sub-section. Without that heading: the end of the "Opening of the session"
 * sub-section, else of the section of the first "Opening ..." heading. -1
 * when the report has none of them.
 */
function findAdhocOpeningInsertIndex_(body) {
  const PARAGRAPH = DocumentApp.ElementType.PARAGRAPH;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const count = body.getNumChildren();
  const headingAt = function (i) {
    const child = body.getChild(i);
    if (child.getType() !== PARAGRAPH || child.asParagraph().getHeading() === NORMAL) return null;
    return child.asParagraph().getText().trim();
  };
  let opening = -1;
  let openingOfSession = -1;
  for (let i = 0; i < count; i++) {
    const text = headingAt(i);
    if (text === null) continue;
    if (/^\d+(?:\.\d+)*\s+Registration of Documents\b/.test(text)) return i;
    if (opening === -1 && /^(?:\d+(?:\.\d+)*\s+)?opening\b/i.test(text)) opening = i;
    if (openingOfSession === -1 && /^(?:\d+(?:\.\d+)*\s+)?opening of the session\s*$/i.test(text)) openingOfSession = i;
  }
  // The sub-section the build writes ("<number> Opening of the session") is
  // preferred over the agenda item above it, whose title may start alike.
  if (openingOfSession !== -1) opening = openingOfSession;
  if (opening === -1) return -1;
  let index = opening + 1;
  while (index < count && headingAt(index) === null) index++;
  return index;
}

/**
 * Writes the container for `lines` and nothing else. A container that
 * already says exactly this is left alone; otherwise it is replaced as a
 * whole (its heading stays where it is). Without lines it is removed.
 * Returns { container: 'created' | 'replaced' | 'unchanged' | 'removed' |
 * 'none' | 'no-place' }.
 */
function renderAdhocOpeningSection_(body, lines) {
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const existing = findAdhocOpeningContainer_(body);
  if (lines.length === 0) {
    if (!existing) return { container: 'none' };
    for (let i = existing.end - 1; i >= existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));
    return { container: 'removed' };
  }

  let at;
  if (existing) {
    const shown = [];
    for (let i = existing.start + 1; i < existing.end; i++) shown.push(body.getChild(i).asParagraph().getText());
    if (shown.length === lines.length && shown.every(function (text, i) { return text === lines[i].text; })) return { container: 'unchanged' };
    for (let i = existing.end - 1; i > existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));
    at = existing.start + 1;
  } else {
    const index = findAdhocOpeningInsertIndex_(body);
    if (index === -1) return { container: 'no-place' };
    const heading = index >= body.getNumChildren() ? body.appendParagraph(ADHOC_OPENING_HEADING_) : body.insertParagraph(index, ADHOC_OPENING_HEADING_);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    at = index + 1;
  }
  lines.forEach(function (line) {
    // The text is passed as text: nothing in it is interpreted.
    const p = at >= body.getNumChildren() ? body.appendParagraph(line.text) : body.insertParagraph(at, line.text);
    p.setHeading(NORMAL);
    p.editAsText().setBold(line.bold);
    at++;
  });
  return { container: existing ? 'replaced' : 'created' };
}

// --- Online information: what the report says about the meeting as a whole.
//
// A main-meeting report copies its opening from the report template, where
// such a block is ordinary text. The opening of an ad-hoc report is
// generated, and until now it said nothing about the meeting. With sessions
// configured, Build Report from Scratch writes this block above the Session
// administration section, from what the report is configured with: nothing
// is looked up, and a value the report does not have is left out. Like the
// Session administration section it is a generated container (its heading,
// then plain paragraphs), replaced as a whole. Only a build creates it;
// later changes of the sessions keep an existing one up to date.
const ADHOC_MEETING_INFO_HEADING_ = 'Online information';
// The labels of its lines, in the order they are written. A line of the block is a paragraph that begins with one of them.
const ADHOC_MEETING_INFO_LABELS_ = ['Meeting name', 'Start Date', 'End Date', 'Portal meeting URL', 'TDoc List URL', 'Excel Docs URL', 'Docs Folder', 'Drafts Folder'];

/**
 * Pure: the lines of the block -- [{ text, link }], `link` the URL the
 * value of the line is linked to, or ''. `source` is { meetingName,
 * meetingId, tdocListUrl, docsFolder, draftsFolder }: values of the
 * report's configuration, '' when it has none. [] without sessions.
 */
function buildAdhocMeetingInformationLines_(sessions, source) {
  if (!sessions || sessions.length === 0) return [];
  const text = function (v) { return String(v === null || v === undefined ? '' : v).replace(/\s+/g, ' ').trim(); };
  const url = function (v) { const u = text(v); return /^https?:\/\/\S+$/i.test(u) ? u : ''; };
  const lines = [];
  const add = function (label, value, link) { if (value) lines.push({ text: label + ': ' + value, link: link ? value : '' }); };
  const ordered = sortAdhocSessions_(sessions);
  let last = '';
  ordered.forEach(function (session) { const day = adhocSessionLastDay_(session); if (day > last) last = day; });
  const id = /^\d+$/.test(text(source && source.meetingId)) ? text(source.meetingId) : '';
  const L = ADHOC_MEETING_INFO_LABELS_;
  add(L[0], text(source && source.meetingName), false);
  add(L[1], adhocLongDate_(ordered[0].date), false);
  add(L[2], adhocLongDate_(last), false);
  add(L[3], id ? 'https://portal.3gpp.org/Home.aspx#/meeting?MtgId=' + id : '', true);
  add(L[4], id ? 'https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=' + id : '', true);
  add(L[5], url(source && source.tdocListUrl), true);
  add(L[6], url(source && source.docsFolder), true);
  add(L[7], url(source && source.draftsFolder), true);
  return lines;
}

/**
 * What the block is written from: the report's own configuration, as it is
 * stored -- no default of a main meeting stands in for a missing value. The
 * meeting id is the one the configured TDoc list URL of the Portal names
 * ("...?meetingId=85916"); with another kind of list URL the two Portal
 * links are left out.
 */
function adhocMeetingInformationSource_() {
  const identity = getMeetingIdentityConfig_();
  const named = String(identity.TDOC_LIST_URL || '').match(/^https:\/\/portal\.3gpp\.org\/[^?#\s]*\?(?:[^#\s]*&)?meetingId=(\d+)(?:&|$)/i);
  return { meetingName: identity.MEETING_NAME, meetingId: named ? named[1] : '', tdocListUrl: identity.TDOC_LIST_URL, docsFolder: identity.FTP_BASE, draftsFolder: identity.REVISIONS_URL };
}

/**
 * The generated block: { start, end } (end exclusive), or null. It is its
 * heading and the lines directly below it that begin with one of its labels.
 * It ends at the first paragraph that does not: text typed below the block
 * -- between it and the Session administration section, where the opening
 * of a meeting is minuted -- is not part of it and is never replaced.
 */
function findAdhocMeetingInformationContainer_(body) {
  const PARAGRAPH = DocumentApp.ElementType.PARAGRAPH;
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const count = body.getNumChildren();
  for (let i = 0; i < count; i++) {
    const child = body.getChild(i);
    if (child.getType() !== PARAGRAPH) continue;
    const p = child.asParagraph();
    if (p.getHeading() === NORMAL || p.getText().trim() !== ADHOC_MEETING_INFO_HEADING_) continue;
    let end = i + 1;
    while (end < count) {
      const next = body.getChild(end);
      if (next.getType() !== PARAGRAPH || next.asParagraph().getHeading() !== NORMAL) break;
      const line = next.asParagraph().getText();
      if (!ADHOC_MEETING_INFO_LABELS_.some(function (label) { return line.indexOf(label + ': ') === 0; })) break;
      end++;
    }
    return { start: i, end: end };
  }
  return null;
}

/**
 * Writes the block for `lines` and nothing else. One that already says
 * exactly this is left alone; otherwise it is replaced as a whole. Without
 * lines it is removed. A new one is written only when `mayCreate` is set
 * (a build): directly above the Session administration section, else where
 * that section would go. Returns { container: 'created' | 'replaced' |
 * 'unchanged' | 'removed' | 'none' | 'absent' | 'no-place' }.
 */
function renderAdhocMeetingInformation_(body, lines, mayCreate) {
  const NORMAL = DocumentApp.ParagraphHeading.NORMAL;
  const existing = findAdhocMeetingInformationContainer_(body);
  if (lines.length === 0) {
    if (!existing) return { container: 'none' };
    for (let i = existing.end - 1; i >= existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));
    return { container: 'removed' };
  }
  let at;
  if (existing) {
    const shown = [];
    for (let i = existing.start + 1; i < existing.end; i++) shown.push(body.getChild(i).asParagraph().getText());
    if (shown.length === lines.length && shown.every(function (text, i) { return text === lines[i].text; })) return { container: 'unchanged' };
    for (let i = existing.end - 1; i > existing.start; i--) removeAdhocBodyChild_(body, body.getChild(i));
    at = existing.start + 1;
  } else {
    if (!mayCreate) return { container: 'absent' };
    const administration = findAdhocOpeningContainer_(body);
    const index = administration ? administration.start : findAdhocOpeningInsertIndex_(body);
    if (index === -1) return { container: 'no-place' };
    const heading = index >= body.getNumChildren() ? body.appendParagraph(ADHOC_MEETING_INFO_HEADING_) : body.insertParagraph(index, ADHOC_MEETING_INFO_HEADING_);
    heading.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    at = index + 1;
  }
  lines.forEach(function (line) {
    // The text is passed as text: nothing in it is interpreted.
    const p = at >= body.getNumChildren() ? body.appendParagraph(line.text) : body.insertParagraph(at, line.text);
    p.setHeading(NORMAL);
    const t = p.editAsText();
    t.setBold(false);
    if (line.link) t.setLinkUrl(line.text.length - line.link.length, line.text.length - 1, line.link);
    at++;
  });
  return { container: existing ? 'replaced' : 'created' };
}

/** The same, from the report's sessions and configuration. A problem is logged and never thrown: the block is no reason to fail a build or a save. */
function refreshAdhocMeetingInformationSafely_(body, sessions, mayCreate) {
  try {
    return renderAdhocMeetingInformation_(body, buildAdhocMeetingInformationLines_(sessions, adhocMeetingInformationSource_()), mayCreate);
  } catch (e) {
    Logger.log('Opening details: the Online information block could not be written: ' + e.message);
    return { container: 'failed' };
  }
}

/**
 * Brings the container up to date with the sessions and the stored details.
 * Stored details that cannot be used (unreadable, or of a newer version)
 * leave an existing report alone -- { container: 'skipped' } -- so that what
 * it shows is not replaced by less; after a rebuild (options.afterRebuild)
 * the sessions are written without details instead.
 */
function refreshAdhocOpeningSection_(body, options) {
  const sessions = getAdhocSessions_();
  // The report-level block above it: written by a build, kept up to date afterwards.
  refreshAdhocMeetingInformationSafely_(body, sessions, !!(options && options.afterRebuild));
  const stored = parseAdhocOpeningProperty_(adhocOpeningStore_().getProperty(ADHOC_OPENING_KEY_));
  if ((stored.status === 'unsupported' || stored.status === 'invalid') && !(options && options.afterRebuild)) return { container: 'skipped', error: stored.error };
  return renderAdhocOpeningSection_(body, buildAdhocOpeningLines_(sessions, stored.sessions));
}

/** Build from Scratch, at the end: writes the container again. Returns a note for the build result, '' when all is well. */
function finishAdhocOpeningRebuild_(body) {
  if (!adhocSessionsEnabled_()) return '';
  try {
    refreshAdhocOpeningSection_(body, { afterRebuild: true });
    return '';
  } catch (e) {
    Logger.log('Opening details: the section could not be written after the rebuild: ' + e.message);
    return '\n\n⚠️ The Session administration section could not be written (' + e.message + '). Save Edit Opening Details… again to write it.';
  }
}

/** The check every opening action starts with; '' when the report can have opening details. */
function adhocOpeningUnavailable_() {
  if (!isAdhocMeetingForSessions_()) return 'Opening details are available for ad-hoc reports only.';
  if (!adhocSessionsEnabled_()) return 'Opening details are entered per session. Use Configure Sessions first.';
  return '';
}

/** Renders after a change that is already stored. Returns '' or a notice saying how to write the section later. */
function renderAdhocOpeningAfterChange_(saved) {
  const outcome = withAdhocAttendanceLock_(function () {
    try {
      return { ok: true, rendered: refreshAdhocOpeningSection_(DocumentApp.getActiveDocument().getBody()) };
    } catch (e) {
      Logger.log('Opening details: stored, but the section could not be written: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
  if (!outcome.ok) return saved + ' The Session administration section could not be written (' + outcome.error + '). Save Edit Opening Details… again to write it.';
  if (outcome.rendered.container === 'no-place') return saved + ' This report has no opening section to put them in; they are written when the report is built.';
  if (outcome.rendered.container === 'skipped') return saved + ' The Session administration section was left as it is (' + outcome.rendered.error + ').';
  return '';
}

/** RPC of the Edit Opening Details dialog: Save. `input` is { <session id>: { chair, minuteTakers, note } }. */
function saveAdhocOpeningDetails(input) {
  assertNotTemplateMaster_();
  const unavailable = adhocOpeningUnavailable_();
  if (unavailable) return { ok: false, errors: [unavailable], changed: false };
  let result;
  try {
    result = saveAdhocOpeningWith_(adhocOpeningStore_(), getAdhocSessions_(), input);
  } catch (e) {
    return { ok: false, errors: ['Nothing was saved: the opening details could not be stored (' + e.message + ').'], changed: false };
  }
  if (!result.ok) return result;
  // Saved (or already as stored): the section is brought in line; one that says the same already is not touched.
  const notice = renderAdhocOpeningAfterChange_('The opening details were saved.');
  if (notice) result.notice = notice;
  return result;
}

/** What the Edit Opening Details dialog shows (reads only). */
function buildAdhocOpeningDialogModel_() {
  const stored = parseAdhocOpeningProperty_(adhocOpeningStore_().getProperty(ADHOC_OPENING_KEY_));
  return {
    notice: stored.status === 'unsupported' ? stored.error + ' They were written by a newer release and cannot be changed here.'
      : (stored.status === 'invalid' ? stored.error + ' They are not used; saving replaces them.' : ''),
    readOnly: stored.status === 'unsupported',
    fields: ADHOC_OPENING_FIELDS_.map(function (f) { return { key: f.key, label: f.label, max: f.max, multiline: f.multiline }; }),
    sessions: getAdhocSessions_().map(function (s) {
      const entry = stored.sessions[s.id] || {};
      return { id: s.id, when: adhocSessionWhenText_(s), chair: entry.chair || '', minuteTakers: entry.minuteTakers || '', note: entry.note || '' };
    })
  };
}

/** Menu (ad-hoc template report): Sessions and Attendance > Edit Opening Details… */
function editAdhocOpeningDetails() {
  assertNotTemplateMaster_();
  const ui = DocumentApp.getUi();
  const unavailable = adhocOpeningUnavailable_();
  if (unavailable) {
    ui.alert('Edit Opening Details', unavailable === 'Opening details are available for ad-hoc reports only.' ? unavailable
      : 'Opening details are entered per session. Configure the sessions of this report first:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…', ui.ButtonSet.OK);
    return;
  }
  // "<" is written as an escape so that no stored text can end the script element.
  const modelJson = JSON.stringify(buildAdhocOpeningDialogModel_()).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; font-size: 13px; }
      label { display: block; margin-top: 8px; font-weight: bold; }
      input[type=text], textarea { width: 100%; padding: 6px; margin-top: 3px; box-sizing: border-box; font-family: Arial, sans-serif; font-size: 13px; }
      textarea { min-height: 54px; }
      button { margin-top: 16px; padding: 9px 18px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:hover { background: #357ae8; }
      button:disabled { background: #999; cursor: default; }
      button.grey { background: #666; }
      .hint { font-size: 11px; color: #666; margin-top: 4px; }
      .session { margin-top: 14px; padding: 10px; background: #f5f5f5; border-left: 3px solid #4285f4; }
      .when { font-weight: bold; }
      #notice { margin-top: 10px; padding: 8px; background: #fff6e0; border-left: 3px solid #c77c00; display: none; }
      #errors { color: #a94442; margin-top: 12px; white-space: pre-wrap; }
    </style>

    <h2>Opening Details</h2>
    <div class="hint">Who chaired each session and who took the minutes. The section "Session administration" of the report is written from these fields and from the sessions; change it here, because text typed into that section is replaced.</div>
    <div id="notice"></div>
    <div id="sessions"></div>

    <div id="errors"></div>
    <div>
      <button type="button" id="saveBtn" onclick="saveDetails()">Save</button>
      <button type="button" class="grey" id="closeBtn" onclick="google.script.host.close()">Cancel</button>
    </div>

    <script>
      var MODEL = ${modelJson};
      var NL = String.fromCharCode(10);
      var blocks = [];
      function el(id) { return document.getElementById(id); }
      function addSession(session) {
        var box = document.createElement('div');
        box.className = 'session';
        var when = document.createElement('div');
        when.className = 'when';
        when.textContent = session.when;
        box.appendChild(when);
        var inputs = {};
        MODEL.fields.forEach(function (field) {
          var label = document.createElement('label');
          label.textContent = field.label;
          box.appendChild(label);
          var input = document.createElement(field.multiline ? 'textarea' : 'input');
          if (!field.multiline) input.type = 'text';
          input.maxLength = field.max;
          input.value = session[field.key] || '';
          input.disabled = MODEL.readOnly;
          box.appendChild(input);
          inputs[field.key] = input;
        });
        el('sessions').appendChild(box);
        blocks.push({ id: session.id, inputs: inputs });
      }
      function collect() {
        var out = {};
        blocks.forEach(function (block) {
          var entry = {};
          MODEL.fields.forEach(function (field) { entry[field.key] = block.inputs[field.key].value; });
          out[block.id] = entry;
        });
        return out;
      }
      function showErrors(list) { el('errors').textContent = list.join(NL); }
      function saveDetails() {
        el('saveBtn').disabled = true;
        showErrors([]);
        google.script.run
          .withSuccessHandler(function (result) {
            if (result && result.ok) {
              if (result.notice) alert(result.notice);
              google.script.host.close();
              return;
            }
            el('saveBtn').disabled = false;
            showErrors((result && result.errors) || ['The opening details were not saved.']);
          })
          .withFailureHandler(function (error) {
            el('saveBtn').disabled = false;
            showErrors([error && error.message ? error.message : String(error)]);
          })
          .saveAdhocOpeningDetails(collect());
      }
      MODEL.sessions.forEach(addSession);
      if (MODEL.notice) { el('notice').textContent = MODEL.notice; el('notice').style.display = 'block'; }
      if (MODEL.readOnly) el('saveBtn').disabled = true;
    </script>
  `)
  .setWidth(620)
  .setHeight(640);

  ui.showModalDialog(html, 'Edit Opening Details');
}

// =========================================================
// AD-HOC STATUS (stage F) -- STATUS SUMMARY AND STATISTICS
// =========================================================
//
// For an ad-hoc report with configured sessions: what is there and what is
// still missing -- the sessions, their opening details, the imported
// attendance and the sessions of the TDocs. One model, two read-only views:
//
//   - Report Status Summary gets a short "Sessions and attendance" block;
//   - Sessions and Attendance > Post-meeting Statistics… shows the same and
//     the figures of each session.
//
// Everything is derived when it is asked for, from ADHOC_SESSIONS,
// ADHOC_SESSION_OPENING, the stored attendance, ADHOC_TDOC_SESSIONS and the
// current TDoc list. Nothing is stored for it and nothing is written --
// neither a property nor the document.
//
// Only counts and the summary values Teams states are shown: no attendee, no
// Chair or minute taker, no note, no upload time. A value Teams did not
// state is left out; none is calculated. Each part is read on its own, so
// one that cannot be read does not hide the others. Something that is merely
// not there yet is a remark ("info"); only data that cannot be used is a
// warning.

/** "September 22, 2026", or "September 22–24, 2026": from the first day of the first session to the last day of the last. */
function adhocSessionDateRangeText_(sessions) {
  if (sessions.length === 0) return '';
  const first = sessions.map(function (s) { return s.date; }).sort()[0];
  const last = sessions.map(adhocSessionLastDay_).sort()[sessions.length - 1];
  return adhocDateRangeText_(first, last);
}

/**
 * Pure: the session counts of the TDocs of a downloaded list. `resolver` is
 * the one the Session column is computed from (stage B) -- nothing is
 * assigned here a second way. Every TDoc of the list is one row; a revision
 * is a TDoc of its own.
 *
 *   total, withSession, severalSessions, none, explicitNone
 *   automatic            no manual assignment, a session by upload time
 *   manualAdd, manualSet TDocs with a manual assignment
 *   noneByReason         why the others have no session
 *   staleReferences      manual assignments naming a session that is not configured
 *   notInList            manual assignments of TDocs the list does not have now
 *   perSession[id]       { assigned, automatic, manual }: a TDoc counts in
 *                        every session it is in; `automatic` when it is there
 *                        by its upload time, `manual` when a manual
 *                        assignment put it there.
 */
function summarizeAdhocTdocSessions_(resolver, allTdocs) {
  const configured = {};
  const perSession = {};
  resolver.sessions.forEach(function (s) { configured[s.id] = true; perSession[s.id] = { assigned: 0, automatic: 0, manual: 0 }; });
  const out = {
    total: allTdocs.length, withSession: 0, severalSessions: 0, none: 0, explicitNone: 0, automatic: 0, manualAdd: 0, manualSet: 0,
    noneByReason: { 'after-last': 0, 'not-uploaded': 0, 'unreadable': 0, 'no-upload-time': 0, 'date-only': 0, 'stale': 0 },
    staleReferences: 0, notInList: 0, perSession: perSession
  };
  const listed = {};
  allTdocs.forEach(function (td) {
    const number = tdocNumberOf_(td);
    listed[number] = true;
    const auto = resolver.automatic(td);
    const manual = resolver.overrides[number];
    const effective = resolver.effective(td);
    if (manual) out[manual.mode === 'set' ? 'manualSet' : 'manualAdd']++;
    else if (auto.sessionId) out.automatic++;

    if (effective.length === 0) {
      out.none++;
      if (manual && manual.mode === 'set' && manual.sessions.length === 0) out.explicitNone++;
      else if (manual && manual.mode === 'set') out.noneByReason.stale++;
      else if (Object.prototype.hasOwnProperty.call(out.noneByReason, auto.reason)) out.noneByReason[auto.reason]++;
      return;
    }
    out.withSession++;
    if (effective.length > 1) out.severalSessions++;
    effective.forEach(function (id) {
      const byUploadTime = auto.sessionId === id && (!manual || manual.mode === 'add');
      perSession[id].assigned++;
      perSession[id][byUploadTime ? 'automatic' : 'manual']++;
    });
  });
  Object.keys(resolver.overrides).forEach(function (number) {
    if (!listed[number]) out.notInList++;
    if (resolver.overrides[number].sessions.some(function (id) { return !configured[id]; })) out.staleReferences++;
  });
  return out;
}

/**
 * Pure: the status of the sessions, for both views.
 *   sessions    the configured sessions, chronological
 *   opening     a parseAdhocOpeningProperty_() result
 *   attendance  a readAdhocAttendance_() result
 *   tdocs       { available, reason, column, clock, storedStatus, storedError, summary }
 *               column: 5 with a Session column, 4 without, 0 no
 *               registration table, -1 unknown
 * Returns { count, range, overview: [[label, value]], notes: [{ level:
 * 'info' | 'warning', text }], sessions: [{ id, when, groups: [{ title,
 * rows: [[label, value]] }] }], opening, attendance, tdocs } -- the last
 * three hold the counts. No name, note or upload time gets in here.
 */
function buildAdhocSessionStatusModel_(sessions, opening, attendance, tdocs) {
  const total = sessions.length;
  const notes = [];
  const note = function (level, text) { notes.push({ level: level, text: text }); };
  const labels = function (list) { return list.map(function (s) { return s.displayLabel; }).join(', '); };
  const configured = {};
  sessions.forEach(function (s) { configured[s.id] = true; });
  const count = function (n, one, many) { return n + ' ' + (n === 1 ? one : many); };

  // --- sessions
  const incomplete = sessions.filter(function (s) { return !s.start || !s.end; });
  if (incomplete.length) {
    note('info', 'Planned times incomplete: ' + incomplete.map(function (s) {
      return s.displayLabel + ' (' + (!s.start && !s.end ? 'no start, no end' : (!s.start ? 'no start' : 'no end')) + ')';
    }).join(', ') + '. A session without an end takes the TDocs uploaded until the end of its day.');
  }

  // --- opening details: whether a field is filled in, never what it says
  const openingKnown = opening.status === 'ok' || opening.status === 'absent';
  const openingOf = function (s) {
    const entry = opening.sessions[s.id] || {};
    return { chair: !!entry.chair, minuteTakers: !!entry.minuteTakers, note: !!entry.note };
  };
  const openingFacts = { known: openingKnown, withDetails: 0, noChair: 0, noMinuteTakers: 0, stale: 0 };
  if (openingKnown) {
    const noChair = sessions.filter(function (s) { return !openingOf(s).chair; });
    const noMinuteTakers = sessions.filter(function (s) { return !openingOf(s).minuteTakers; });
    openingFacts.withDetails = sessions.filter(function (s) { const o = openingOf(s); return o.chair || o.minuteTakers || o.note; }).length;
    openingFacts.noChair = noChair.length;
    openingFacts.noMinuteTakers = noMinuteTakers.length;
    openingFacts.stale = Object.keys(opening.sessions).filter(function (id) { return !configured[id]; }).length;
    if (noChair.length) note('info', 'Chair not entered: ' + labels(noChair) + '.');
    if (noMinuteTakers.length) note('info', 'Minute taker(s) not entered: ' + labels(noMinuteTakers) + '.');
    if (openingFacts.stale) note('warning', 'Opening details are stored for ' + count(openingFacts.stale, 'session that is', 'sessions that are') + ' not configured.');
  } else {
    note('warning', (opening.error || 'The stored opening details cannot be read.') + ' The status of the opening details is not available.');
  }

  // --- attendance: the values Teams states, and how many attendees are listed
  const attendanceKnown = attendance.status === 'ok' || attendance.status === 'absent';
  // A session can have several attendance records, one per day ("Attendance records").
  const blocks = {};
  if (attendance.status === 'ok') {
    buildAdhocAttendanceBlocks_(sessions, attendance).forEach(function (block) {
      (blocks[block.sessionId] = blocks[block.sessionId] || []).push({ values: block.statistics.filter(function (line) { return line[0] !== 'Session'; }), attendees: block.attendees.length });
    });
  }
  const sessionOf = function (key) { const parts = adhocAttendanceRecordParts_(key); return parts ? parts.sessionId : String(key); };
  const damagedIds = attendance.status === 'ok' ? attendance.problems.map(sessionOf) : [];
  const unreadable = function (s) { return damagedIds.filter(function (id) { return id === s.id; }).length; };
  const attendanceState = function (s) {
    if (!attendanceKnown) return 'unknown';
    if (blocks[s.id]) return 'imported';
    return unreadable(s) ? 'damaged' : 'none';
  };
  // The rows of a session's attendance: the values of each record, one after the other, each with its number of attendees.
  const attendanceRows = function (s) {
    let rows = [];
    blocks[s.id].forEach(function (b) { rows = rows.concat(b.values, [['Attendees listed', String(b.attendees)]]); });
    return unreadable(s) ? rows.concat([['Records that cannot be read', String(unreadable(s))]]) : rows;
  };
  const attendanceFacts = { known: attendanceKnown, imported: 0, missing: 0, damaged: 0, stale: 0, records: 0 };
  if (attendanceKnown) {
    const missing = sessions.filter(function (s) { return attendanceState(s) === 'none'; });
    const damaged = sessions.filter(function (s) { return unreadable(s) > 0; });
    attendanceFacts.imported = Object.keys(blocks).length;
    Object.keys(blocks).forEach(function (id) { attendanceFacts.records += blocks[id].length; });
    attendanceFacts.missing = missing.length;
    attendanceFacts.damaged = damaged.length;
    attendanceFacts.stale = attendance.status === 'ok'
      ? Object.keys(attendance.sessions).concat(attendance.problems).filter(function (key) { return !configured[sessionOf(key)]; }).length : 0;
    if (missing.length) note('info', 'Attendance not imported: ' + labels(missing) + '.');
    sessions.forEach(function (s) {
      const days = adhocSessionDayCount_(s);
      if (blocks[s.id] && blocks[s.id].length < days) note('info', 'Attendance of ' + s.displayLabel + ' is imported for ' + blocks[s.id].length + ' of its ' + days + ' days.');
    });
    if (damaged.length) note('warning', 'The stored attendance of ' + labels(damaged) + ' cannot be read. Import it again or remove it.');
    if (attendanceFacts.stale) note('warning', 'Attendance is stored for ' + count(attendanceFacts.stale, 'session that is', 'sessions that are') + ' not configured.');
  } else {
    note('warning', (attendance.error || 'The stored attendance cannot be read.') + ' The status of the attendance is not available.');
  }

  // --- TDoc sessions
  const t = tdocs.available ? tdocs.summary : null;
  const columnText = tdocs.column === 5 ? 'in the registration table'
    : (tdocs.column === 4 ? 'not enabled (Assign TDoc Sessions… can add it)' : (tdocs.column === 0 ? 'the report has no registration table yet' : 'not known'));
  if (t) {
    if (tdocs.storedStatus === 'invalid' || tdocs.storedStatus === 'unsupported') {
      note('warning', (tdocs.storedError || 'The stored manual TDoc session assignments cannot be read.') + ' Only the sessions by upload time are counted.');
    }
    const r = t.noneByReason;
    if (r['after-last']) note('warning', count(r['after-last'], 'TDoc was', 'TDocs were') + ' uploaded after the end of the last session and ' + (r['after-last'] === 1 ? 'has' : 'have') + ' no session.');
    if (r['unreadable']) note('warning', 'The upload time of ' + count(r['unreadable'], 'TDoc', 'TDocs') + ' cannot be read; no session.');
    if (r['no-upload-time']) note('warning', 'The TDoc list has no upload times: ' + count(r['no-upload-time'], 'TDoc', 'TDocs') + ' without a session.');
    if (r['stale'] || t.staleReferences) note('warning', count(t.staleReferences, 'manual assignment names', 'manual assignments name') + ' a session that is not configured.');
    if (r['not-uploaded']) note('info', count(r['not-uploaded'], 'TDoc is', 'TDocs are') + ' not uploaded yet and ' + (r['not-uploaded'] === 1 ? 'has' : 'have') + ' no session.');
    if (r['date-only']) note('info', count(r['date-only'], 'TDoc has', 'TDocs have') + ' an upload date without a time, on the day of a cut-off; no session.');
    if (t.notInList) note('info', count(t.notInList, 'manual assignment is', 'manual assignments are') + ' for a TDoc that is not in the TDoc list now.');
    if (tdocs.clock === 'session') note('info', 'The upload times are read on the clock of the sessions, not as UTC (compatibility setting of Assign TDoc Sessions…).');
  } else {
    note('warning', 'The TDoc sessions are not available: the TDoc list could not be read (' + (tdocs.reason || 'unknown reason') + ').');
  }

  const overview = [['Sessions', total + ' (' + adhocSessionDateRangeText_(sessions) + ')']];
  overview.push(['Opening details', openingKnown ? openingFacts.withDetails + ' of ' + total + ' sessions' : 'not available']);
  overview.push(['Attendance imported', attendanceKnown ? attendanceFacts.imported + ' of ' + total + ' sessions' +
    (attendanceFacts.records > attendanceFacts.imported ? ' (' + attendanceFacts.records + ' attendance records)' : '') : 'not available']);
  if (t) {
    overview.push(['TDocs with a session', t.withSession + ' of ' + t.total]);
    overview.push(['– by upload time', String(t.automatic)]);
    overview.push(['– with a manual assignment', (t.manualAdd + t.manualSet) + ' (' + t.manualAdd + ' added, ' + t.manualSet + ' set)']);
    overview.push(['– in more than one session', String(t.severalSessions)]);
    overview.push(['TDocs without a session', String(t.none)]);
    overview.push(['– set to no session', String(t.explicitNone)]);
  } else {
    overview.push(['TDoc sessions', 'not available']);
  }
  overview.push(['Session column', columnText]);

  return {
    count: total,
    range: adhocSessionDateRangeText_(sessions),
    overview: overview,
    notes: notes,
    sessions: sessions.map(function (s) {
      const groups = [{ title: 'Planned', rows: isAdhocMultiDaySession_(s)
        ? [['Dates', adhocSessionDatesText_(s)], ['Start (first day)', s.start || 'not planned'], ['End (last day)', s.end || 'not planned']]
        : [['Date', adhocLongDate_(s.date)], ['Start', s.start || 'not planned'], ['End', s.end || 'not planned']] }];
      const o = openingOf(s);
      groups.push({ title: 'Opening', rows: openingKnown
        ? [['Chair', o.chair ? 'entered' : 'not entered'], ['Minute taker(s)', o.minuteTakers ? 'entered' : 'not entered'], ['Administrative note', o.note ? 'entered' : 'none']]
        : [['Opening details', 'not available']] });
      const state = attendanceState(s);
      groups.push({ title: 'Attendance (Microsoft Teams)', rows: state === 'imported'
        ? attendanceRows(s)
        : [['Attendance', state === 'none' ? 'not imported' : (state === 'damaged' ? 'stored, but cannot be read' : 'not available')]] });
      groups.push({ title: 'TDocs', rows: t
        ? [['In this session', String(t.perSession[s.id].assigned)], ['– by upload time', String(t.perSession[s.id].automatic)], ['– by manual assignment', String(t.perSession[s.id].manual)]]
        : [['TDoc sessions', 'not available']] });
      return { id: s.id, when: adhocSessionWhenText_(s), groups: groups };
    }),
    opening: openingFacts,
    attendance: attendanceFacts,
    tdocs: t
  };
}

/**
 * Reads everything the status is made of -- and writes nothing. null unless
 * this is an ad-hoc report with sessions. A part that cannot be read is
 * reported as such in the model; the TDoc list is the only download, and a
 * failed one leaves the rest of the status as it is.
 */
function collectAdhocSessionStatus_() {
  const sessions = getAdhocSessions_();
  if (sessions.length === 0) return null;

  let opening;
  try {
    opening = parseAdhocOpeningProperty_(adhocOpeningStore_().getProperty(ADHOC_OPENING_KEY_));
  } catch (e) {
    opening = { status: 'invalid', sessions: {}, error: 'The opening details could not be read (' + e.message + ').' };
  }
  let attendance;
  try {
    attendance = readAdhocAttendance_(adhocAttendanceStore_());
  } catch (e) {
    attendance = { status: 'invalid', error: 'The attendance could not be read (' + e.message + ').', sessions: {}, raw: {}, problems: [], companies: {} };
  }

  const tdocs = { available: false, reason: '', column: -1, clock: 'utc', storedStatus: 'absent', storedError: null, summary: null };
  try {
    const found = findRegistrationTable_(DocumentApp.getActiveDocument().getBody());
    tdocs.column = found ? found.columns : 0;
  } catch (e) {
    tdocs.column = -1;
  }
  try {
    const resolver = makeAdhocTdocSessionResolver_();
    const cfg = getReportConfig_();
    if (!cfg.TDOC_LIST_URL) throw new Error('no TDoc list is configured');
    tdocs.summary = summarizeAdhocTdocSessions_(resolver, flattenTdocGroups_(downloadAndGroupTdocs_(cfg)));
    tdocs.clock = resolver.clock;
    tdocs.storedStatus = resolver.status;
    tdocs.storedError = resolver.error;
    tdocs.available = true;
  } catch (e) {
    tdocs.reason = e.message;
    Logger.log('Session status: the TDoc sessions are not available: ' + e.message);
  }
  return buildAdhocSessionStatusModel_(sessions, opening, attendance, tdocs);
}

/** Pure: the block of the Report Status Summary, in its style. */
function formatAdhocSessionStatusLines_(model) {
  const lines = ['\n\n🗓 SESSIONS AND ATTENDANCE\n'];
  model.overview.forEach(function (row) { lines.push('   ' + row[0] + ': ' + row[1]); });
  const warnings = model.notes.filter(function (n) { return n.level === 'warning'; });
  const remarks = model.notes.filter(function (n) { return n.level !== 'warning'; });
  if (warnings.length) lines.push('');
  warnings.forEach(function (n) { lines.push('   ⚠️  ' + n.text); });
  if (remarks.length) lines.push('');
  remarks.forEach(function (n) { lines.push('   ℹ️  ' + n.text); });
  lines.push('\n   Per session: Sessions and Attendance > Post-meeting Statistics…');
  return lines;
}

/**
 * The lines Report Status Summary adds: none -- and nothing read beyond the
 * sessions -- unless this is an ad-hoc report with sessions. Never throws.
 */
function adhocSessionStatusLines_() {
  let enabled = false;
  try { enabled = adhocSessionsEnabled_(); } catch (e) { enabled = false; }
  if (!enabled) return [];
  try {
    return formatAdhocSessionStatusLines_(collectAdhocSessionStatus_());
  } catch (e) {
    Logger.log('Session status: not available: ' + e.message);
    return ['\n\n🗓 SESSIONS AND ATTENDANCE\n', '   ⚠️  The status of the sessions is not available (' + e.message + ').'];
  }
}

/** Menu (ad-hoc template report): Sessions and Attendance > Post-meeting Statistics… -- a read-only view. */
function showAdhocSessionStatistics() {
  assertNotTemplateMaster_();
  const ui = DocumentApp.getUi();
  if (!isAdhocMeetingForSessions_()) {
    ui.alert('Post-meeting Statistics', 'The statistics of the sessions are available for ad-hoc reports only.', ui.ButtonSet.OK);
    return;
  }
  const model = collectAdhocSessionStatus_();
  if (!model) {
    ui.alert('Post-meeting Statistics', 'The statistics are per session. Configure the sessions of this report first:\n\nSA4 Report > Sessions and Attendance > Configure Sessions…', ui.ButtonSet.OK);
    return;
  }
  // "<" is written as an escape so that no stored text can end the script element.
  const modelJson = JSON.stringify({ overview: model.overview, notes: model.notes, sessions: model.sessions }).replace(/</g, '\\u003c');

  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; font-size: 13px; }
      h3 { margin: 16px 0 4px 0; font-size: 13px; }
      button { margin-top: 16px; padding: 9px 18px; background: #666; color: white; border: none; cursor: pointer; }
      table { border-collapse: collapse; width: 100%; }
      td { padding: 2px 8px 2px 0; vertical-align: top; }
      td.label { width: 45%; color: #444; }
      .hint { font-size: 11px; color: #666; margin-top: 4px; }
      .session { margin-top: 14px; padding: 10px; background: #f5f5f5; border-left: 3px solid #4285f4; }
      .when { font-weight: bold; }
      .group { margin-top: 8px; font-weight: bold; color: #444; }
      .note { margin-top: 6px; padding: 6px 8px; border-left: 3px solid #999; background: #f5f5f5; }
      .note.warning { border-left-color: #c77c00; background: #fff6e0; }
    </style>

    <h2>Post-meeting Statistics</h2>
    <div class="hint">Counted from the sessions, the opening details, the imported attendance and the TDoc list as they are now. Nothing is stored or changed by opening this. The attendance values are the ones Microsoft Teams states.</div>
    <h3>All sessions</h3>
    <div id="overview"></div>
    <div id="notes"></div>
    <div id="sessions"></div>
    <div><button type="button" id="closeBtn" onclick="google.script.host.close()">Close</button></div>

    <script>
      var MODEL = ${modelJson};
      function el(id) { return document.getElementById(id); }
      function addRows(parent, rows) {
        var table = document.createElement('table');
        rows.forEach(function (row) {
          var tr = document.createElement('tr');
          var label = document.createElement('td');
          label.className = 'label';
          label.textContent = row[0];
          var value = document.createElement('td');
          value.textContent = row[1];
          tr.appendChild(label);
          tr.appendChild(value);
          table.appendChild(tr);
        });
        parent.appendChild(table);
      }
      addRows(el('overview'), MODEL.overview);
      MODEL.notes.forEach(function (note) {
        var box = document.createElement('div');
        box.className = note.level === 'warning' ? 'note warning' : 'note';
        box.textContent = note.text;
        el('notes').appendChild(box);
      });
      MODEL.sessions.forEach(function (session) {
        var box = document.createElement('div');
        box.className = 'session';
        var when = document.createElement('div');
        when.className = 'when';
        when.textContent = session.when;
        box.appendChild(when);
        session.groups.forEach(function (group) {
          var title = document.createElement('div');
          title.className = 'group';
          title.textContent = group.title;
          box.appendChild(title);
          addRows(box, group.rows);
        });
        el('sessions').appendChild(box);
      });
    </script>
  `)
  .setWidth(640)
  .setHeight(680);

  ui.showModalDialog(html, 'Post-meeting Statistics');
}

// =========================================================
// ARCH-009 -- MEETING-ID RESOLVER CORE
// =========================================================
//
// Built from the ARCH-007/ARCH-008 investigations. This section is
// DELIBERATELY ADDITIVE and NOT wired into getMeetingContext_(), the
// configuration dialog, or any other production consumer -- every function
// here is dead code from the rest of the app's point of view until a
// future task explicitly integrates it. Nothing here mutates
// PropertiesService.
//
// Three real, anonymous, official 3GPP/ETSI endpoints, verified during
// ARCH-007/008 against real meeting IDs (86178, 86174, 85916, 60778) with
// no cookies, no session, no auth header, and (for two of the three) no
// User-Agent at all:
//
//   - POST https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetMeetings
//     (primary identity/date/FTP source)
//   - GET  https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetiCal/{id}.ics
//     (secondary identity/date source)
//   - GET  https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId={id}
//     (TDoc/agenda discovery source)
//
// Network access (fetchMeetingMetadataById_/fetchMeetingIcalById_/
// fetchMeetingTdocListById_) is kept separate from parsing/normalization
// (everything else below), so the parsing/normalization logic is pure and
// independently testable without UrlFetchApp -- see
// tests/meeting-resolver.test.js.
//
// ARCH-011 adds a fourth, DERIVED (not Portal-provided) endpoint: a
// candidate revisions/drafts folder, one level up from the resolved
// FTP Docs/ directory (<meeting root>/inbox/drafts/), validated with its
// own probe (fetchRevisionsUrlCandidate_) before ever being reported as
// resolved. See deriveRevisionsUrlCandidate_()/
// validateRevisionsUrlCandidateResponse_() below and
// tests/meeting-revisions-folder.test.js.

/**
 * ARCH-009: validates and normalizes a Meeting ID input. Accepts a
 * positive integer or a string containing ONLY digits (after trimming) --
 * never extracts digits from arbitrary surrounding text (e.g. "id: 86178"
 * is rejected, not silently parsed). Rejects empty/blank, non-numeric,
 * zero, negative, and non-integer (e.g. 86178.5) values.
 */
function parseMeetingIdInput_(input) {
  if (input === null || input === undefined) {
    return { isValid: false, id: null, error: 'Meeting ID is required.' };
  }
  if (typeof input === 'number') {
    if (!Number.isFinite(input) || !Number.isInteger(input) || input <= 0) {
      return { isValid: false, id: null, error: 'Meeting ID must be a positive integer.' };
    }
    return { isValid: true, id: input, error: null };
  }
  if (typeof input === 'string') {
    const trimmed = input.trim();
    if (trimmed === '') {
      return { isValid: false, id: null, error: 'Meeting ID is required.' };
    }
    if (!/^[0-9]+$/.test(trimmed)) {
      return { isValid: false, id: null, error: 'Meeting ID must contain only digits.' };
    }
    const n = parseInt(trimmed, 10);
    if (n <= 0) {
      return { isValid: false, id: null, error: 'Meeting ID must be a positive integer.' };
    }
    return { isValid: true, id: n, error: null };
  }
  return { isValid: false, id: null, error: 'Meeting ID must be a number or a numeric string.' };
}

// ----------------------------------------------- network access (impure) --

function fetchMeetingMetadataById_(meetingId) {
  const url = 'https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetMeetings';
  const payload = JSON.stringify({
    getMeetingsInput: {
      MeetingId: meetingId,
      StartRow: 0,
      ResultsPerPage: 1,
      IncludeChildTbs: true,
      IncludeNonTBMeetings: true,
      Tbs: [0]
    }
  });
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  });
  return { statusCode: response.getResponseCode(), text: response.getContentText() };
}

function fetchMeetingIcalById_(meetingId) {
  const url = `https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetiCal/${meetingId}.ics`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return { statusCode: response.getResponseCode(), text: response.getContentText() };
}

function fetchMeetingTdocListById_(meetingId) {
  const url = `https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=${meetingId}`;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return { statusCode: response.getResponseCode(), text: response.getContentText() };
}

/**
 * ARCH-011: probes a derived revisions/drafts folder candidate (see
 * deriveRevisionsUrlCandidate_() below). `followRedirects: false` is
 * deliberate -- a candidate that redirects anywhere (including a
 * seemingly benign trailing-slash redirect) is treated as unvalidated
 * rather than silently followed, since this project has no way to inspect
 * the actual redirect target's content here.
 */
function fetchRevisionsUrlCandidate_(url) {
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: false });
  return { statusCode: response.getResponseCode(), text: response.getContentText() };
}

// -------------------------------------- pure parsing/normalization below --

/**
 * ARCH-009: parses a fetchMeetingMetadataById_() result. Handles: non-200
 * HTTP status, malformed JSON, a non-array response, and a valid-but-empty
 * array (a syntactically valid Meeting ID that Portal does not recognize --
 * NOT an error, `meeting` is simply null).
 */
function parseMeetingMetadataResponse_(fetchResult) {
  if (!fetchResult || fetchResult.statusCode !== 200) {
    return { ok: false, meeting: null, error: `GetMeetings returned HTTP ${fetchResult ? fetchResult.statusCode : 'unknown'}` };
  }
  let parsed;
  try {
    parsed = JSON.parse(fetchResult.text);
  } catch (e) {
    return { ok: false, meeting: null, error: 'GetMeetings response was not valid JSON: ' + e.message };
  }
  if (!Array.isArray(parsed)) {
    return { ok: false, meeting: null, error: 'GetMeetings response was not an array.' };
  }
  if (parsed.length === 0) {
    return { ok: true, meeting: null, error: null };
  }
  return { ok: true, meeting: parsed[0], error: null };
}

/**
 * ARCH-009: strips a literal leading "3GPP" marker only (e.g.
 * "3GPPSA4-e (AH) on FS_6G_MED" -> "SA4-e (AH) on FS_6G_MED"). Deliberately
 * narrow -- an explicit, tested helper, not broad title rewriting. The
 * original Portal title is always preserved separately by the caller
 * (resolveMeetingById_()'s `raw.meeting`).
 */
function normalizePortalMeetingTitle_(rawTitle) {
  const s = String(rawTitle || '').trim();
  if (s.indexOf('3GPP') === 0) {
    return s.slice(4);
  }
  return s;
}

// SA4-ARCH-009: only these two Portal `Type` codes are verified (against
// 4 real meetings: "AH" x3, "OR" x1). Do not add unverified codes here --
// an unrecognized code must surface as unresolved, never guessed.
const PORTAL_MEETING_TYPE_MAP_ = { AH: 'adhoc', OR: 'main' };

/**
 * ARCH-009: maps a raw Portal `Type` code to this project's
 * meeting.type vocabulary. Returns `recognized:false` (never a guess from
 * title text) for anything outside PORTAL_MEETING_TYPE_MAP_.
 */
function normalizePortalMeetingType_(rawType) {
  const key = String(rawType || '').trim();
  if (Object.prototype.hasOwnProperty.call(PORTAL_MEETING_TYPE_MAP_, key)) {
    return { type: PORTAL_MEETING_TYPE_MAP_[key], portalType: key, recognized: true };
  }
  return { type: null, portalType: key || null, recognized: false };
}

/**
 * ARCH-009: normalizes a Portal `MtgDocURL` into the actual TDoc Docs/
 * directory. Conservative by design:
 *   - collapses an accidental doubled slash right after the host (observed
 *     on meeting 86178: "https://ftp.3gpp.org//tsg_sa/...") without
 *     touching the "://" itself;
 *   - never changes path segment casing (observed to vary --
 *     "TSG_SA" vs "tsg_sa" -- across real meetings; this function does not
 *     "fix" that, it passes it through untouched);
 *   - if the path already ends in "/Docs/" (case-insensitive check), keeps
 *     it as-is;
 *   - otherwise appends literal "Docs/" -- this is the ONLY case observed
 *     (meeting 86178) and the only case handled; nothing else is inferred.
 */
function normalizeMtgDocUrlToFtpBase_(mtgDocUrl) {
  if (!mtgDocUrl || !String(mtgDocUrl).trim()) {
    return { ftpBase: null, error: 'MtgDocURL is missing.' };
  }
  let s = String(mtgDocUrl).trim();
  s = s.replace(/^(https?:\/\/[^/]+)\/{2,}/, '$1/');
  s = s.replace(/([^:])\/{2,}/g, '$1/');
  if (!/\/$/.test(s)) s += '/';
  if (/\/Docs\/$/i.test(s)) {
    return { ftpBase: s, error: null };
  }
  return { ftpBase: s + 'Docs/', error: null };
}

/**
 * ARCH-011: derives a revisions/drafts folder CANDIDATE from an already
 * normalized ftpBase (normalizeMtgDocUrlToFtpBase_() output), using ONLY
 * the authoritative MtgDocURL-derived directory evidence already resolved
 * -- never the meeting title or Meeting ID. Conceptually:
 *
 *   <meeting root>/Docs/           (ftpBase)
 *   <meeting root>/inbox/drafts/   (this candidate)
 *
 * so the "Docs/" segment is replaced with "inbox/drafts/" under the same
 * meeting-root parent. This is a PURE derivation only -- it does not mean
 * the candidate is real; see fetchRevisionsUrlCandidate_()/
 * validateRevisionsUrlCandidateResponse_() for that. Returns
 * `revisionsUrlCandidate: null` (with an `error`) when ftpBase is
 * missing/blank or does not end in a recognizable "Docs/" segment -- this
 * function never guesses a meeting root from anything else.
 */
function deriveRevisionsUrlCandidate_(ftpBase) {
  if (!ftpBase || !String(ftpBase).trim()) {
    return { revisionsUrlCandidate: null, error: 'ftpBase is missing.' };
  }
  let s = String(ftpBase).trim();
  s = s.replace(/^(https?:\/\/[^/]+)\/{2,}/, '$1/');
  s = s.replace(/([^:])\/{2,}/g, '$1/');
  if (!/\/$/.test(s)) s += '/';
  const m = s.match(/^(.*\/)Docs\/$/i);
  if (!m) {
    return { revisionsUrlCandidate: null, error: 'ftpBase does not end in a recognizable "Docs/" segment; cannot derive a meeting root.' };
  }
  return { revisionsUrlCandidate: `${m[1]}inbox/drafts/`, error: null };
}

/**
 * ARCH-011: decides whether a fetchRevisionsUrlCandidate_() result is
 * credible evidence that the candidate revisions/drafts folder actually
 * exists and is accessible -- a derived candidate is NEVER itself treated
 * as resolved (see deriveRevisionsUrlCandidate_() above). Requires HTTP
 * 200 (a 3xx/4xx/5xx, including a redirect -- see
 * fetchRevisionsUrlCandidate_()'s `followRedirects: false` -- is a
 * failure) AND a body that looks like a directory/file listing (an
 * Apache-style "Index of" autoindex title, or at least one `<a href=`
 * entry -- this project's real 3GPP FTP directory pages use exactly this
 * shape; a malformed or unexpectedly empty body is rejected, not
 * guessed-at).
 */
function validateRevisionsUrlCandidateResponse_(fetchResult) {
  if (!fetchResult || fetchResult.statusCode !== 200) {
    return { ok: false, reason: `HTTP ${fetchResult ? fetchResult.statusCode : 'unknown'}` };
  }
  const text = String(fetchResult.text || '');
  if (!text.trim()) {
    return { ok: false, reason: 'Empty response body.' };
  }
  const looksLikeIndex = /Index of/i.test(text);
  const hasLinks = /<a\s+href=/i.test(text);
  if (!looksLikeIndex && !hasLinks) {
    return { ok: false, reason: 'Response did not look like a directory/file listing (no "Index of" title, no <a href> entries).' };
  }
  return { ok: true, reason: null };
}

/**
 * ARCH-009: minimal, line-based ICS property extraction (UID, SUMMARY,
 * DTSTART, DTEND, LOCATION, DESCRIPTION) -- the properties this project's
 * GetiCal responses were observed to contain. No RFC-5545 folding/
 * unfolding, no timezone conversion -- this is a secondary/cross-check
 * source, not the primary one.
 */
function parseMeetingIcal_(icsText) {
  const text = String(icsText || '');
  function extractLine(name) {
    const m = text.match(new RegExp('^' + name + ':(.*)$', 'm'));
    return m ? m[1].trim() : null;
  }
  return {
    uid: extractLine('UID'),
    summary: extractLine('SUMMARY'),
    dtstart: extractLine('DTSTART'),
    dtend: extractLine('DTEND'),
    location: extractLine('LOCATION'),
    description: extractLine('DESCRIPTION')
  };
}

function decodeTdocListHtmlEntities_(s) {
  return String(s || '')
    .replace(/&amp;/g, '&')
    .replace(/&#39;/g, '\'')
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ');
}

function stripTdocListHtmlTags_(s) {
  return String(s || '').replace(/<[^>]*>/g, '').trim();
}

/**
 * ARCH-009: splits one TdocList.aspx grid row's HTML into its top-level
 * <td>...</td> cell contents, in column order. Verified against real
 * captured rows from meetings 86178 (32 TDocs) and 60778 (482 TDocs,
 * confirming this also scales to the largest real fixture observed) --
 * these RadGrid rows do not nest additional <table> markup inside a cell,
 * so a non-greedy <td>...</td> match is safe for this specific tool's
 * output. Column indices (0-based), confirmed by direct inspection of real
 * rows for BOTH an "agenda"-typed row and a revision pair:
 *   0 = details-icon cell (no text)      6 = For
 *   1 = TDoc link (id + FTP url)         7 = meeting name (constant per meeting)
 *   2 = Type                             8 = agenda allocation (span.agendaItem)
 *   3 = Title                            9 = Revision of (link or empty anchor)
 *   4 = Source                          10 = Revised To (link or empty)
 *   5 = Status                          11 = Extra info (usually empty)
 */
function extractTdocListRowCells_(rowHtml) {
  const cells = [];
  const re = /<td[^>]*>([\s\S]*?)<\/td>/g;
  let m;
  while ((m = re.exec(rowHtml))) {
    cells.push(m[1]);
  }
  return cells;
}

function extractTdocListRevisionLinkId_(cellHtml) {
  if (!cellHtml) return null;
  const m = cellHtml.match(/<a[^>]*>([^<]*)<\/a>/);
  if (!m) return null;
  const text = decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(m[1]));
  return text ? text : null;
}

/**
 * ARCH-009: parses TdocList.aspx's server-rendered HTML into one object per
 * TDoc row. Does not depend on any browser DOM API (Apps Script has none) --
 * pure regex/string parsing only. Rows that don't contain a recognizable
 * TDoc link (e.g. a header/pager row that also happens to carry an
 * rgRow-shaped class) are silently skipped rather than producing a
 * malformed entry.
 */
function parseMeetingTdocListHtml_(html) {
  const text = String(html || '');
  const rows = [];
  const rowRe = /<tr[^>]*class="(?:rgRow|rgAltRow)"[\s\S]*?<\/tr>/g;
  let rowMatch;
  while ((rowMatch = rowRe.exec(text))) {
    const rowHtml = rowMatch[0];
    const cells = extractTdocListRowCells_(rowHtml);
    const tdocLinkMatch = cells[1] && cells[1].match(/<a[^>]*href="([^"]*)"[^>]*>([^<]*)<\/a>/);
    if (!tdocLinkMatch) continue;

    const id = decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(tdocLinkMatch[2]));
    const url = decodeTdocListHtmlEntities_(tdocLinkMatch[1]);
    const agendaSpanMatch = rowHtml.match(/<span[^>]*title="([^"]*)"[^>]*class="agendaItem"[^>]*>([^<]*)<\/span>/);

    rows.push({
      id: id,
      url: url,
      type: decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(cells[2] || '')),
      title: decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(cells[3] || '')),
      source: decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(cells[4] || '')),
      status: decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(cells[5] || '')),
      forAction: decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(cells[6] || '')),
      meetingName: decodeTdocListHtmlEntities_(stripTdocListHtmlTags_(cells[7] || '')),
      agendaItem: agendaSpanMatch ? decodeTdocListHtmlEntities_(agendaSpanMatch[2].trim()) : null,
      agendaTopic: agendaSpanMatch ? decodeTdocListHtmlEntities_(agendaSpanMatch[1]) : null,
      revisionOf: extractTdocListRevisionLinkId_(cells[9]),
      revisedTo: extractTdocListRevisionLinkId_(cells[10])
    });
  }
  return rows;
}

/**
 * ARCH-009: pure agenda-candidate selection over already-parsed TDoc rows.
 * Algorithm (per the ARCH-008 evidence-based design):
 *   1. candidates = rows where type === "agenda" (case-insensitive);
 *   2. a candidate that is named as another agenda candidate's
 *      `revisionOf` target is superseded -- excluded from "current";
 *   3. if exactly one current candidate remains, it is the resolved
 *      agendaTdoc;
 *   4. zero candidates -> agendaTdoc: null, unresolved reason "no agenda
 *      TDoc found" (verified real case: meeting 86174);
 *   5. more than one still-current candidate -> agendaTdoc: null,
 *      ambiguous: true, every remaining candidate listed -- never a silent
 *      pick.
 * Never infers anything from title text; `status` and `revisionOf` are the
 * only signals used, matching what was actually evidenced (60778: the
 * "revised" S4-261375 / "approved" S4-261392 pair).
 */
function selectAgendaCandidate_(tdocRows) {
  const rows = Array.isArray(tdocRows) ? tdocRows : [];
  const agendaCandidates = rows.filter(r => String(r.type || '').trim().toLowerCase() === 'agenda');

  if (agendaCandidates.length === 0) {
    return { agendaTdoc: null, candidates: [], ambiguous: false, unresolvedReason: 'No TDoc with type "agenda" was found.' };
  }

  const supersededIds = {};
  agendaCandidates.forEach(c => {
    if (c.revisionOf) supersededIds[c.revisionOf] = true;
  });
  const current = agendaCandidates.filter(c => !supersededIds[c.id]);

  if (current.length === 1) {
    return { agendaTdoc: current[0].id, candidates: current, ambiguous: false, unresolvedReason: null };
  }
  if (current.length === 0) {
    // Every agenda-typed candidate was superseded by another -- shouldn't
    // normally happen (something must be "current"), but represent it
    // honestly rather than falling back to a guess.
    return { agendaTdoc: null, candidates: agendaCandidates, ambiguous: true, unresolvedReason: 'All agenda-typed TDocs appear superseded; none is clearly current.' };
  }
  return { agendaTdoc: null, candidates: current, ambiguous: true, unresolvedReason: 'Multiple current agenda-typed TDocs found.' };
}

/**
 * ARCH-009: derives TDoc-family evidence from already-parsed TDoc rows,
 * reusing the EXISTING centralized parseSA4DocumentId_()/SA4_TDOC_FAMILIES
 * registry -- no second family-recognition implementation. Surfaces a
 * mixed-family meeting explicitly (`consistent:false`) rather than
 * silently picking the first family seen.
 */
function deriveTdocFamilyEvidence_(tdocRows) {
  const rows = Array.isArray(tdocRows) ? tdocRows : [];
  const familiesSeen = {};
  const unrecognizedIds = [];
  rows.forEach(r => {
    const parsed = parseExactSA4DocumentId_(r.id);
    if (parsed.isValid) {
      familiesSeen[parsed.familyKey] = (familiesSeen[parsed.familyKey] || 0) + 1;
    } else {
      unrecognizedIds.push(r.id);
    }
  });
  const familyKeys = Object.keys(familiesSeen);
  return {
    family: familyKeys.length === 1 ? familyKeys[0] : null,
    familiesSeen: familiesSeen,
    consistent: familyKeys.length <= 1,
    unrecognizedIds: unrecognizedIds
  };
}

/**
 * PROD-014: decides whether the GetiCal fallback fetch is worth making at
 * all. iCal is ONLY ever consulted (see resolveMeetingById_() below) as a
 * fallback for meeting.name/startDate/endDate/location when GetMeetings
 * didn't supply them -- every real captured GetMeetings response (86178,
 * 86174, 85916, 60778) already supplies all four, so on the normal
 * success path this fetch was pure latency with no effect on the result.
 * Skipped entirely when GetMeetings has already conclusively determined
 * the meeting does not exist (an empty result) since the caller's early
 * "meeting not found" return never consults `ical` either. Still fetched
 * whenever GetMeetings itself failed outright (iCal is then the only
 * remaining source) or the meeting it found is missing any one of the
 * four fields iCal can substitute for.
 */
function shouldFetchIcalFallback_(metadataParsed) {
  if (!metadataParsed || !metadataParsed.ok) return true;
  if (metadataParsed.meeting === null) return false;
  const m = metadataParsed.meeting || {};
  return m.Title === undefined || m.StartDate === undefined || m.EndDate === undefined || m.Location === undefined;
}

/**
 * PROD-014: the ONLY place that still performs ARCH-011's revisions/
 * drafts folder network probe -- resolveMeetingById_() itself no longer
 * calls it synchronously (see that function's PROD-014 note below), since
 * a production smoke test on meeting 86178 found the Resolve action
 * hanging for 3+ minutes with no response, and this was the one network
 * call in the chain whose target host (www.3gpp.org's FTP paths, behind a
 * WAF/bot-challenge -- see ARCH-011's own commit message) had NOT
 * previously been confirmed fast/reliable from a plain HTTP client the
 * way portal.3gpp.org's REST API had been.
 *
 * This function is meant to be invoked as an explicit, SEPARATE action
 * (e.g. a future "Validate" step) against a candidate URL that
 * resolveMeetingById_() already derived (its `sources.revisionsUrlCandidate`)
 * -- never wired into the normal Resolve path. Never mutates
 * PropertiesService. Returns `{ ok, revisionsUrl, reason }`; `revisionsUrl`
 * is the candidate itself once validated, else null.
 */
function validateRevisionsUrlCandidate_(candidateUrl) {
  if (!candidateUrl || !String(candidateUrl).trim()) {
    return { ok: false, revisionsUrl: null, reason: 'No candidate URL supplied.' };
  }
  try {
    const fetchResult = fetchRevisionsUrlCandidate_(candidateUrl);
    const validation = validateRevisionsUrlCandidateResponse_(fetchResult);
    if (validation.ok) {
      return { ok: true, revisionsUrl: candidateUrl, reason: null };
    }
    return { ok: false, revisionsUrl: null, reason: validation.reason };
  } catch (e) {
    return { ok: false, revisionsUrl: null, reason: 'Revisions/drafts folder candidate request failed: ' + e.message };
  }
}

/**
 * PROD-016: CORE meeting resolution -- the ONLY network call is GetMeetings
 * (POST), plus a conditional GetiCal fallback via shouldFetchIcalFallback_()
 * (PROD-014, unchanged). NEVER fetches TdocList.aspx and NEVER fetches the
 * revisions/drafts folder candidate (still derive-only, per PROD-014) --
 * agenda/TDoc discovery is a SEPARATE, explicit operation, see
 * enrichMeetingFromTdocList_() below.
 *
 * This exists because PROD-014 (removing the synchronous revisions probe)
 * was NOT sufficient: a fresh live smoke test on meeting 86178 still hung
 * 90+ seconds, proving TdocList.aspx itself (or its combination with
 * GetiCal) was also part of the problem. This is what the configuration
 * dialog's Resolve button now calls directly (resolveMeetingForConfigDialog_()),
 * so a slow/hanging TdocList.aspx request can never make Meeting-ID
 * resolution itself appear hung again.
 *
 * `documents` is always null here. `unresolved` still lists
 * 'documents.agendaTdoc'/'documents.family' (not yet attempted, not
 * "failed") using the exact same unresolved-list convention the rest of
 * this module already uses, so a caller can't mistake "not yet
 * discovered" for "resolved".
 */
function resolveMeetingCoreById_(meetingId) {
  const idResult = parseMeetingIdInput_(meetingId);
  const warnings = [];
  const unresolved = [];

  if (!idResult.isValid) {
    return {
      id: null,
      meeting: null,
      sources: null,
      documents: null,
      unresolved: ['meeting'],
      warnings: [idResult.error],
      raw: {}
    };
  }

  const id = idResult.id;
  const raw = {};

  // --- Primary: GetMeetings -------------------------------------------
  // POST-MEETING-001 (Task A5): manual-diagnostic stage markers -- see
  // resolveMeetingForConfigDialog_()'s own header note.
  Logger.log('resolveMeetingCoreById_: before GetMeetings');
  let metadataParsed = { ok: false, meeting: null, error: 'GetMeetings was not called.' };
  try {
    const metadataFetch = fetchMeetingMetadataById_(id);
    metadataParsed = parseMeetingMetadataResponse_(metadataFetch);
  } catch (e) {
    metadataParsed = { ok: false, meeting: null, error: 'GetMeetings request failed: ' + e.message };
  }
  Logger.log('resolveMeetingCoreById_: after GetMeetings, ok=' + metadataParsed.ok);
  if (!metadataParsed.ok) {
    warnings.push(metadataParsed.error);
  }
  raw.meeting = metadataParsed.meeting;

  // --- Secondary: GetiCal (cross-check / fallback only) ----------------
  // PROD-014: only fetched when it can actually matter -- see
  // shouldFetchIcalFallback_(). Skipped on the normal success path (every
  // real captured GetMeetings response already supplies everything iCal
  // could otherwise substitute for).
  let ical = null;
  if (shouldFetchIcalFallback_(metadataParsed)) {
    try {
      const icalFetch = fetchMeetingIcalById_(id);
      if (icalFetch.statusCode === 200) {
        ical = parseMeetingIcal_(icalFetch.text);
      } else {
        warnings.push(`GetiCal returned HTTP ${icalFetch.statusCode}`);
      }
    } catch (e) {
      warnings.push('GetiCal request failed: ' + e.message);
    }
  }
  raw.ical = ical;

  // --- Handle "no meeting found" up front -------------------------------
  if (metadataParsed.ok && metadataParsed.meeting === null) {
    unresolved.push('meeting');
    return {
      id: id,
      meeting: null,
      sources: {
        portalMeetingUrl: `https://portal.3gpp.org/Home.aspx#/meeting?MtgId=${id}`,
        tdocListEndpoint: `https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=${id}`,
        icalEndpoint: `https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetiCal/${id}.ics`,
        revisionsUrl: null,
        revisionsUrlCandidate: null
      },
      documents: null,
      unresolved: unresolved,
      warnings: warnings,
      raw: raw
    };
  }

  const m = metadataParsed.meeting || {};
  const typeInfo = normalizePortalMeetingType_(m.Type);
  if (!typeInfo.recognized && m.Type !== undefined) {
    unresolved.push('meeting.type');
    warnings.push(`Unrecognized Portal meeting Type code: ${JSON.stringify(m.Type)}`);
  }

  const ftpInfo = m.MtgDocURL ? normalizeMtgDocUrlToFtpBase_(m.MtgDocURL) : { ftpBase: null, error: 'MtgDocURL is missing.' };
  if (!ftpInfo.ftpBase) {
    unresolved.push('sources.ftpBase');
    warnings.push(ftpInfo.error);
  }

  // --- Revisions/drafts folder discovery (ARCH-011 derivation; PROD-014
  // made this DERIVE-ONLY, no network). sources.revisionsUrl therefore
  // stays null from THIS function always; only
  // validateRevisionsUrlCandidate_(), called separately, can ever turn a
  // candidate into a validated "resolved" value.
  let revisionsUrlCandidate = null;
  if (!ftpInfo.ftpBase) {
    unresolved.push('sources.revisionsUrl');
  } else {
    const candidateResult = deriveRevisionsUrlCandidate_(ftpInfo.ftpBase);
    unresolved.push('sources.revisionsUrl');
    if (!candidateResult.revisionsUrlCandidate) {
      warnings.push('Could not derive a revisions/drafts folder candidate: ' + candidateResult.error);
    } else {
      revisionsUrlCandidate = candidateResult.revisionsUrlCandidate;
    }
  }

  const name = m.Title !== undefined ? normalizePortalMeetingTitle_(m.Title) : (ical && ical.summary ? normalizePortalMeetingTitle_(ical.summary) : null);
  if (!name) unresolved.push('meeting.name');

  // PROD-016: agenda/TDoc discovery has NOT been attempted yet -- it is
  // now a separate, explicit operation (enrichMeetingFromTdocList_()).
  // Marked unresolved here for exactly the same reason every other
  // not-yet-determined field is: absence must never be mistaken for a
  // negative result.
  unresolved.push('documents.agendaTdoc');
  unresolved.push('documents.family');

  // Mailing list is never derived -- see ARCH-008 §6/§9. Always unresolved.
  unresolved.push('mailingList');

  return {
    id: id,

    meeting: {
      name: name,
      type: typeInfo.type,
      portalType: typeInfo.portalType,
      group: 'SA4',
      tb: m.TB !== undefined ? m.TB : null,
      tbId: m.TBId !== undefined ? m.TBId : null,
      startDate: m.StartDate !== undefined ? m.StartDate : (ical ? ical.dtstart : null),
      endDate: m.EndDate !== undefined ? m.EndDate : (ical ? ical.dtend : null),
      timeZone: m.StartTimeZone !== undefined ? m.StartTimeZone : null,
      location: m.Location !== undefined ? m.Location : (ical ? ical.location : null)
    },

    sources: {
      ftpBase: ftpInfo.ftpBase,
      portalMeetingUrl: `https://portal.3gpp.org/Home.aspx#/meeting?MtgId=${id}`,
      tdocListEndpoint: `https://portal.3gpp.org/ngppapp/TdocList.aspx?meetingId=${id}`,
      icalEndpoint: `https://portal.3gpp.org/webservices/Rest/Meetings.svc/GetiCal/${id}.ics`,
      mailingList: null,
      revisionsUrl: null,
      revisionsUrlCandidate: revisionsUrlCandidate
    },

    // Always null from core resolution -- see enrichMeetingFromTdocList_().
    documents: null,

    unresolved: unresolved,
    warnings: warnings,
    raw: raw
  };
}

/**
 * PROD-016: the SEPARATE, explicit TDoc/agenda ENRICHMENT step -- fetches
 * TdocList.aspx (the ONLY network call this function makes) and merges
 * family/agenda evidence into a CLONE of `coreResult` (a previously
 * computed resolveMeetingCoreById_() result), returning a full result in
 * the SAME shape resolveMeetingById_() has always returned. Never mutates
 * `coreResult`; never touches PropertiesService.
 *
 * Safe to call independently, at any time after a core resolve -- its own
 * failure or latency can never affect or erase what core resolution
 * already established. On any failure the returned result keeps
 * `documents: null` (exactly as coreResult already had it) plus an added
 * warning, the same degrade-to-warning behavior every other source in
 * this module already has.
 */
function enrichMeetingFromTdocList_(meetingId, coreResult) {
  const idResult = parseMeetingIdInput_(meetingId);
  const base = coreResult ? JSON.parse(JSON.stringify(coreResult)) : null;

  if (!idResult.isValid || !base) {
    return base || {
      id: null,
      meeting: null,
      sources: null,
      documents: null,
      unresolved: ['meeting'],
      warnings: [idResult.error || 'No core result supplied to enrich.'],
      raw: {}
    };
  }

  if (base.id !== idResult.id) {
    base.warnings = (base.warnings || []).concat(['enrichMeetingFromTdocList_: meetingId does not match the supplied core result -- enrichment skipped.']);
    return base;
  }

  if (!base.meeting) {
    // Core resolution never found a meeting -- nothing to enrich.
    return base;
  }

  let tdocRows = [];
  let tdocListFetchOk = false;
  try {
    const tdocFetch = fetchMeetingTdocListById_(idResult.id);
    if (tdocFetch.statusCode === 200) {
      tdocRows = parseMeetingTdocListHtml_(tdocFetch.text);
      tdocListFetchOk = true;
    } else {
      base.warnings.push(`TdocList.aspx returned HTTP ${tdocFetch.statusCode}`);
    }
  } catch (e) {
    base.warnings.push('TdocList.aspx request failed: ' + e.message);
  }
  if (tdocListFetchOk && tdocRows.length === 0) {
    base.warnings.push('TdocList.aspx returned no recognizable TDoc rows.');
  }
  if (!tdocListFetchOk) {
    // Enrichment failed -- core/existing configuration is untouched;
    // documents stays exactly as coreResult already had it (null).
    return base;
  }

  const agendaResult = selectAgendaCandidate_(tdocRows);
  const unresolved = base.unresolved.filter(u => u !== 'documents.agendaTdoc' && u !== 'documents.family');
  if (!agendaResult.agendaTdoc) {
    unresolved.push('documents.agendaTdoc');
    if (agendaResult.unresolvedReason) base.warnings.push(agendaResult.unresolvedReason);
  }

  const familyEvidence = deriveTdocFamilyEvidence_(tdocRows);
  if (!familyEvidence.consistent) {
    unresolved.push('documents.family');
    base.warnings.push('Inconsistent TDoc families observed: ' + JSON.stringify(familyEvidence.familiesSeen));
  } else if (!familyEvidence.family && tdocRows.length > 0) {
    unresolved.push('documents.family');
  }

  const rawDocCount = base.raw && base.raw.meeting && base.raw.meeting.DocCount !== undefined ? base.raw.meeting.DocCount : undefined;
  base.documents = {
    count: rawDocCount !== undefined ? rawDocCount : tdocRows.length,
    family: familyEvidence.family,
    familiesSeen: familyEvidence.familiesSeen,
    agendaTdoc: agendaResult.agendaTdoc,
    agendaCandidates: agendaResult.candidates.map(c => c.id),
    agendaAmbiguous: agendaResult.ambiguous,
    agendaItemsObserved: tdocRows.map(r => r.agendaItem).filter(v => v !== null && v !== undefined && v !== '')
  };
  base.unresolved = unresolved;
  return base;
}

/**
 * ARCH-009 (kept for compatibility) / PROD-016: FULL meeting resolution --
 * resolveMeetingCoreById_() PLUS enrichMeetingFromTdocList_(), composed
 * together. This is NOT what the configuration dialog's Resolve button
 * calls any more -- it calls resolveMeetingCoreById_() alone, with
 * enrichment as a separate, explicit "Discover Agenda / TDocs" action
 * (discoverAgendaForConfigDialog_()). resolveMeetingById_() is kept, with
 * its full original GetMeetings+[GetiCal]+TdocList.aspx behavior and
 * result shape UNCHANGED, for any full-resolution consumer (and this
 * module's own regression tests) that wants the old all-in-one call.
 * NEVER mutates PropertiesService.
 */
function resolveMeetingById_(meetingId) {
  const core = resolveMeetingCoreById_(meetingId);
  if (!core.meeting) return core;
  return enrichMeetingFromTdocList_(meetingId, core);
}

// =========================================================
// ADDON-008A -- AD-HOC AGENDA.CSV FALLBACK + MEETING-SPECIFIC TDOC LIST
// =========================================================
//
// Three concepts, three separate properties:
//   - AGENDA_TDOC     the agenda TDoc reference -- unchanged, and still the
//                     preferred agenda STRUCTURE source.
//   - AGENDA_CSV_URL  the ad-hoc series' public agenda.csv
//                     (<series>/Agenda/agenda.csv, the Portal's "Agenda"
//                     link). Series-wide, so it is only used as the
//                     structure source for an ad-hoc meeting with no agenda
//                     TDoc, and only after it matches THIS meeting's own
//                     TDoc list (every agenda item/description pair).
//   - TDOC_LIST_URL   the contribution list -- unchanged; for ad-hoc
//                     meetings the Portal's meeting-ID document list is now
//                     discovered, an explicit value still wins.
//
// Ad-hoc agenda-structure precedence (main meetings are never affected):
//   1. saved AGENDA_TDOC, 2. discovered unambiguous agenda TDoc,
//   3. validated agenda.csv, 4. the existing readiness error.
//
// Verified live (2026-09-30): 86172 (MBS, no agenda TDoc) has a 53-row
// headerless "number","title" CSV whose rows match its TDoc list exactly
// (2.5/2.7/3.7, sort order = CSV row); 85916 (Audio) has agenda TDoc
// S4aA260090 and a 0-byte CSV.
//
// Remote content is untrusted: candidate URLs are rebuilt server-side from
// the resolved meeting (never taken from the client), CSV/XLSX content is
// parsed defensively, and the build re-validates the CSV itself before the
// document is touched.

const ADHOC_AGENDA_CSV_MAX_BYTES_ = 262144;
const ADHOC_AGENDA_CSV_MAX_ROWS_ = 1000;
const ADHOC_AGENDA_TITLE_MAX_LENGTH_ = 500;
const ADHOC_SERIES_DOCS_RE_ = /^https:\/\/(?:www\.3gpp\.org\/ftp|ftp\.3gpp\.org)\/tsg_sa\/wg4_codec\/3gpp_sa4_ahoc_mtgs\/([A-Za-z0-9_.-]+)\/docs\/$/i;
const ADHOC_NO_AGENDA_TDOC_WARNING_ = 'No TDoc with type "agenda" was found.';

/**
 * <series>/Docs/ -> <series>/Agenda/agenda.csv, same host. Only for the
 * ad-hoc series layout (…/3GPP_SA4_AHOC_MTGs/<series>/Docs/ on the 3GPP
 * FTP); any other folder (a main meeting, an unknown host) has no candidate.
 */
function deriveAdhocAgendaCsvCandidate_(ftpBase) {
  const s = String(ftpBase === null || ftpBase === undefined ? '' : ftpBase).trim();
  const m = s.match(ADHOC_SERIES_DOCS_RE_);
  if (!m || /^\.+$/.test(m[1])) {
    return { url: null, error: 'The meeting document folder is not an ad-hoc series folder (…/3GPP_SA4_AHOC_MTGs/<series>/Docs/).' };
  }
  return { url: s.slice(0, s.length - 'Docs/'.length) + 'Agenda/agenda.csv', error: null };
}

/** True only for the exact agenda.csv candidate of this FTP_BASE. */
function isAdhocAgendaCsvUrlFor_(url, ftpBase) {
  const candidate = deriveAdhocAgendaCsvCandidate_(ftpBase);
  return !!candidate.url && String(url === null || url === undefined ? '' : url).trim() === candidate.url;
}

/** The Portal's meeting-ID document list (anonymous XLSX download). */
function buildAdhocTdocListCandidateUrl_(meetingId) {
  const idResult = parseMeetingIdInput_(meetingId);
  if (!idResult.isValid) return null;
  return 'https://portal.3gpp.org/ngppapp/GenerateDocumentList.aspx?meetingId=' + idResult.id;
}

function normalizeAgendaTitleForMatch_(title) {
  return String(title === null || title === undefined ? '' : title)
    .replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * Splits CSV text into rows of cells (RFC 4180 quoting, "" escapes, CRLF/LF,
 * leading BOM). Returns null for unbalanced quotes.
 */
function splitCsvRows_(text) {
  const src = String(text).replace(/^﻿/, '');
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (inQuotes) {
      if (ch === '"') {
        if (src[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(field); field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && src[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else {
      field += ch;
    }
  }
  if (inQuotes) return null;
  if (field !== '' || row.length > 0) { row.push(field); rows.push(row); }
  return rows;
}

/**
 * Parses the headerless two-column ("number","title") agenda.csv into the
 * SAME item shape the agenda-table parser produces ({number, title, level,
 * heading, text}), in file order. Strict: any malformed row, duplicate
 * number, or a file with no numbered rows makes the whole CSV unusable.
 */
function parseAgendaCsv_(text) {
  function fail(reason) { return { ok: false, items: [], reason: reason }; }
  const src = String(text === null || text === undefined ? '' : text);
  if (src.trim() === '') return fail('agenda.csv is empty.');
  if (src.length > ADHOC_AGENDA_CSV_MAX_BYTES_) return fail('agenda.csv is unexpectedly large.');

  const rows = splitCsvRows_(src);
  if (!rows) return fail('agenda.csv is malformed (unbalanced quotes).');
  const meaningful = rows.filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });
  if (meaningful.length === 0) return fail('agenda.csv has no agenda rows.');
  if (meaningful.length > ADHOC_AGENDA_CSV_MAX_ROWS_) return fail('agenda.csv has unexpectedly many rows.');

  const items = [];
  const seen = {};
  for (let i = 0; i < meaningful.length; i++) {
    const cells = meaningful[i];
    const extra = cells.slice(2).some(function (c) { return String(c).trim() !== ''; });
    const number = String(cells[0] || '').trim();
    const title = String(cells[1] || '').replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim();
    if (cells.length < 2 || extra || !/^\d{1,3}(?:\.\d{1,3}){0,4}$/.test(number) || !title || title.length > ADHOC_AGENDA_TITLE_MAX_LENGTH_) {
      return fail('agenda.csv row ' + (i + 1) + ' is not a "number","title" agenda entry.');
    }
    if (seen[number]) return fail('agenda.csv lists agenda item ' + number + ' more than once.');
    seen[number] = true;
    items.push({
      number: number,
      title: title,
      level: (number.match(/\./g) || []).length + 1,
      heading: DocumentApp.ParagraphHeading.NORMAL,
      text: ''
    });
  }
  return { ok: true, items: items, reason: null };
}

/**
 * The existing importer needs "TDoc" and "Agenda item"; a plausible list has
 * at least one recognizable SA4 TDoc row.
 */
function validateTdocListWorkbookValues_(values) {
  const rows = Array.isArray(values) ? values : [];
  const header = (rows[0] || []).map(function (h) { return String(h === null || h === undefined ? '' : h).trim(); });
  const tdocCol = header.indexOf('TDoc');
  if (tdocCol === -1 || header.indexOf('Agenda item') === -1) {
    return { ok: false, tdocCount: 0, reason: 'The document list has no "TDoc"/"Agenda item" columns.' };
  }
  let tdocCount = 0;
  let sa4Count = 0;
  for (let i = 1; i < rows.length; i++) {
    const id = String((rows[i] || [])[tdocCol] || '').trim();
    if (!id) continue;
    tdocCount++;
    if (parseExactSA4DocumentId_(id).isValid) sa4Count++;
  }
  if (sa4Count === 0) {
    return { ok: false, tdocCount: tdocCount, reason: 'The document list contains no recognizable SA4 TDocs yet.' };
  }
  return { ok: true, tdocCount: tdocCount, reason: null };
}

/**
 * The meeting's own agenda item -> description pairs, from its TDoc list.
 * A list that describes the same item two different ways is unusable as
 * validation evidence.
 */
function extractTdocListAgendaPairs_(values) {
  const rows = Array.isArray(values) ? values : [];
  const header = (rows[0] || []).map(function (h) { return String(h === null || h === undefined ? '' : h).trim(); });
  const tdocCol = header.indexOf('TDoc');
  const itemCol = header.indexOf('Agenda item');
  const descCol = header.indexOf('Agenda item description');
  if (tdocCol === -1 || itemCol === -1 || descCol === -1) {
    return { ok: false, pairs: [], reason: 'The TDoc list has no "Agenda item description" column to check agenda.csv against.' };
  }
  const byNumber = {};
  const pairs = [];
  const conflicts = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (!String(row[tdocCol] || '').trim()) continue;
    const number = String(row[itemCol] === null || row[itemCol] === undefined ? '' : row[itemCol]).trim();
    const title = String(row[descCol] === null || row[descCol] === undefined ? '' : row[descCol]).trim();
    if (!number || !title) continue;
    if (!Object.prototype.hasOwnProperty.call(byNumber, number)) {
      byNumber[number] = title;
      pairs.push({ number: number, title: title });
    } else if (normalizeAgendaTitleForMatch_(byNumber[number]) !== normalizeAgendaTitleForMatch_(title) && conflicts.indexOf(number) === -1) {
      conflicts.push(number);
    }
  }
  if (conflicts.length > 0) {
    return { ok: false, pairs: [], reason: 'The TDoc list describes agenda item(s) ' + conflicts.join(', ') + ' in different ways.' };
  }
  return { ok: true, pairs: pairs, reason: null };
}

/**
 * agenda.csv is series-wide: accept it only when every agenda item this
 * meeting's TDocs use appears in it with the same title. No pairs to check
 * means no evidence, so the CSV is not accepted automatically. `sections`
 * are the top-level agenda sections those TDocs fall under, in CSV order.
 */
function validateAgendaCsvForMeeting_(csvItems, pairsResult) {
  if (!pairsResult || !pairsResult.ok) {
    return { ok: false, sections: [], reason: pairsResult ? pairsResult.reason : 'No TDoc list to check agenda.csv against.' };
  }
  if (pairsResult.pairs.length === 0) {
    return { ok: false, sections: [], reason: 'The meeting\'s TDocs have no agenda items yet, so agenda.csv cannot be checked against this meeting.' };
  }
  const titleByNumber = {};
  (csvItems || []).forEach(function (it) { titleByNumber[it.number] = it.title; });
  const missing = [];
  const conflicting = [];
  pairsResult.pairs.forEach(function (p) {
    if (!Object.prototype.hasOwnProperty.call(titleByNumber, p.number)) missing.push(p.number);
    else if (normalizeAgendaTitleForMatch_(titleByNumber[p.number]) !== normalizeAgendaTitleForMatch_(p.title)) conflicting.push(p.number);
  });
  if (missing.length > 0) {
    return { ok: false, sections: [], reason: 'agenda.csv does not contain agenda item(s) ' + missing.join(', ') + ' used by this meeting\'s TDocs.' };
  }
  if (conflicting.length > 0) {
    return { ok: false, sections: [], reason: 'agenda.csv titles differ from this meeting\'s TDoc list for agenda item(s) ' + conflicting.join(', ') + '.' };
  }
  const roots = {};
  pairsResult.pairs.forEach(function (p) { roots[p.number.split('.')[0]] = true; });
  const sections = [];
  (csvItems || []).forEach(function (it) {
    const root = it.number.split('.')[0];
    if (roots[root] && sections.indexOf(root) === -1) sections.push(root);
  });
  return { ok: true, sections: sections, matchedCount: pairsResult.pairs.length, reason: null };
}

/**
 * The report agenda from a validated CSV: every CSV item in the top-level
 * sections this meeting's TDocs use (a combined SA4 agenda otherwise pulls
 * in every subgroup), selected through the existing itemList projection.
 */
function selectAdhocCsvAgendaItems_(csvItems, sections) {
  const numbers = (csvItems || [])
    .filter(function (it) { return sections.indexOf(it.number.split('.')[0]) !== -1; })
    .map(function (it) { return it.number; });
  if (numbers.length === 0) return [];
  return projectAgendaItems_(csvItems, { mode: 'itemList', value: numbers });
}

function isXlsxSignature_(bytes) {
  return !!bytes && bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04;
}

function fetchAdhocAgendaCsv_(url) {
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  return { statusCode: response.getResponseCode(), text: response.getContentText() };
}

/**
 * Downloads a TDoc list and reads its first sheet, the same xlsx -> Sheet
 * conversion downloadAndGroupTdocs_() uses. Anything that is not an XLSX
 * (e.g. an HTML error page) is rejected before conversion.
 */
function readTdocListWorkbook_(url) {
  let response;
  try {
    response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {
    return { ok: false, values: null, reason: 'The document list could not be fetched: ' + e.message };
  }
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    return { ok: false, values: null, reason: 'The document list returned HTTP ' + code + '.' };
  }
  const blob = response.getBlob();
  if (!isXlsxSignature_(blob.getBytes())) {
    return { ok: false, values: null, reason: 'The document list is not an Excel workbook.' };
  }
  let tempFile = null;
  try {
    blob.setName('TDoc_List_Discovery_Temp.xlsx');
    tempFile = DriveApp.createFile(blob);
    const values = SpreadsheetApp.open(tempFile).getSheets()[0].getDataRange().getValues();
    return { ok: true, values: values, reason: null };
  } catch (e) {
    return { ok: false, values: null, reason: 'The document list could not be read as a workbook: ' + e.message };
  } finally {
    if (tempFile) { try { tempFile.setTrashed(true); } catch (e) { } }
  }
}

/**
 * Fetch + parse + validate one agenda.csv against the meeting's TDoc list
 * values. ADDON-008A1b: with `reallocations` (a build), the projected
 * sections follow each TDoc's EFFECTIVE report allocation -- a moved TDoc
 * counts at its destination, a removed one not at all -- and every saved
 * destination must exist in the CSV agenda (`reallocationProblems`). The
 * CSV itself is still validated against the unchanged source pairs.
 */
function evaluateAdhocAgendaCsvSource_(csvUrl, tdocListValues, reallocations) {
  function fail(reason) { return { ok: false, url: csvUrl, reason: reason, items: [] }; }
  let fetched;
  try {
    fetched = fetchAdhocAgendaCsv_(csvUrl);
  } catch (e) {
    return fail('agenda.csv could not be fetched: ' + e.message);
  }
  if (!fetched || fetched.statusCode < 200 || fetched.statusCode >= 300) {
    return fail('agenda.csv returned HTTP ' + (fetched ? fetched.statusCode : 'unknown') + '.');
  }
  const parsed = parseAgendaCsv_(fetched.text);
  if (!parsed.ok) return fail(parsed.reason);
  const pairsResult = extractTdocListAgendaPairs_(tdocListValues);
  const validation = validateAgendaCsvForMeeting_(parsed.items, pairsResult);
  if (!validation.ok) return fail(validation.reason);
  let sections = validation.sections;
  if (reallocations && Object.keys(reallocations).length > 0) {
    const allocations = computeEffectiveTdocAllocations_(tdocListValues, reallocations);
    const sourceByTdoc = {};
    allocations.forEach(function (a) { sourceByTdoc[a.tdoc] = a.source; });
    const problems = findReallocationBuildProblems_(reallocations, sourceByTdoc, parsed.items.map(function (it) { return it.number; }));
    if (problems.length > 0) {
      const failed = fail('Saved document reallocations cannot be applied.');
      failed.reallocationProblems = problems;
      return failed;
    }
    // Same evidence as the unreallocated case (items the source list
    // describes), with each reallocated TDoc counted at its destination.
    const described = {};
    pairsResult.pairs.forEach(function (p) { described[p.number] = true; });
    const roots = {};
    allocations.forEach(function (a) {
      if (a.removed || !a.effective) return;
      if (a.reallocated || described[a.effective]) roots[a.effective.split('.')[0]] = true;
    });
    sections = [];
    parsed.items.forEach(function (it) {
      const root = it.number.split('.')[0];
      if (roots[root] && sections.indexOf(root) === -1) sections.push(root);
    });
  }
  const items = selectAdhocCsvAgendaItems_(parsed.items, sections);
  return {
    ok: true, url: csvUrl, reason: null, items: items,
    totalCount: parsed.items.length, selectedCount: items.length,
    sections: sections, matchedCount: validation.matchedCount
  };
}

/**
 * Discover Agenda / TDocs, ad-hoc part: the meeting-ID document list, and --
 * only when no agenda TDoc is available -- the validated agenda.csv
 * fallback. Never throws; every failure is a reason string. The client only
 * ever receives a summary (no remote titles).
 */
function discoverAdhocMeetingSources_(meetingId, ftpBase, checkAgendaCsv) {
  const tdocList = { url: buildAdhocTdocListCandidateUrl_(meetingId), ok: false, tdocCount: 0, reason: null };
  let tdocListValues = null;
  if (!tdocList.url) {
    tdocList.reason = 'No valid meeting ID.';
  } else {
    const workbook = readTdocListWorkbook_(tdocList.url);
    if (!workbook.ok) {
      tdocList.reason = workbook.reason;
    } else {
      const validation = validateTdocListWorkbookValues_(workbook.values);
      tdocList.ok = validation.ok;
      tdocList.tdocCount = validation.tdocCount;
      tdocList.reason = validation.reason;
      if (validation.ok) tdocListValues = workbook.values;
    }
  }
  if (!tdocList.ok) tdocList.url = null;

  const agendaCsv = { checked: !!checkAgendaCsv, url: null, ok: false, reason: null, sections: [], selectedCount: 0, totalCount: 0 };
  if (checkAgendaCsv) {
    const candidate = deriveAdhocAgendaCsvCandidate_(ftpBase);
    if (!candidate.url) {
      agendaCsv.reason = candidate.error;
    } else if (!tdocListValues) {
      agendaCsv.reason = 'agenda.csv could not be checked because the meeting\'s document list is unavailable.';
    } else {
      const evaluated = evaluateAdhocAgendaCsvSource_(candidate.url, tdocListValues);
      agendaCsv.ok = evaluated.ok;
      agendaCsv.reason = evaluated.reason;
      if (evaluated.ok) {
        agendaCsv.url = candidate.url;
        agendaCsv.sections = evaluated.sections;
        agendaCsv.selectedCount = evaluated.selectedCount;
        agendaCsv.totalCount = evaluated.totalCount;
      }
    }
  }
  return { tdocList: tdocList, agendaCsv: agendaCsv };
}

/**
 * Build-time agenda for an ad-hoc meeting whose agenda comes from
 * agenda.csv: re-fetched and re-validated against the configured TDoc list
 * right here, BEFORE the document is cleared. Returns null whenever the CSV
 * is not the agenda source (main meeting, an agenda TDoc is configured, or
 * no CSV was accepted) so the existing path runs unchanged.
 */
function prepareAdhocCsvAgendaForBuild_(cfg, reallocations) {
  const identity = getMeetingIdentityConfig_();
  if (identity.MEETING_TYPE !== 'adhoc' || identity.AGENDA_TDOC || !identity.AGENDA_CSV_URL) return null;

  function refuse(reason) {
    throw new Error('Cannot build report yet.\n\n• ' + reason + '\n\nOpen Configure Meeting and run "Discover Agenda / TDocs" again.');
  }
  if (!isAdhocAgendaCsvUrlFor_(identity.AGENDA_CSV_URL, identity.FTP_BASE)) {
    refuse('The saved agenda.csv does not belong to this meeting\'s document folder.');
  }
  const workbook = readTdocListWorkbook_(cfg.TDOC_LIST_URL);
  if (!workbook.ok) refuse(workbook.reason);
  const evaluated = evaluateAdhocAgendaCsvSource_(identity.AGENDA_CSV_URL, workbook.values, reallocations);
  if (evaluated.reallocationProblems) refuseBuildForReallocations_(evaluated.reallocationProblems);
  if (!evaluated.ok) refuse(evaluated.reason);
  if (evaluated.items.length === 0) refuse('agenda.csv has no agenda items for this meeting.');
  Logger.log('Agenda structure from validated agenda.csv: ' + evaluated.selectedCount + ' of ' + evaluated.totalCount +
    ' items (sections ' + evaluated.sections.join(', ') + ')');
  return evaluated.items;
}

// =========================================================
// PROD-014 -- DIAGNOSTIC: TIME EACH RESOLVER NETWORK CALL INDEPENDENTLY
// =========================================================
//
// Manual/diagnostic only. Never called by resolveMeetingById_(), any
// dialog, or any menu item -- run it directly from the Apps Script editor
// (select diagnoseMeetingResolverTiming_ in the function dropdown, Run)
// when investigating resolver latency. Reads no Document Properties and
// writes none; entirely read-only against the live 3GPP/ETSI endpoints.
//
// VERIFIED PLATFORM LIMITATION: Google Apps Script's UrlFetchApp has NO
// per-request timeout parameter -- there is no `{ timeout: ... }` (or
// equivalent) option anywhere in its fetch() signature. A call either
// returns (success or an HTTP error status) or eventually throws once
// Google's own internal, undocumented fetch ceiling is hit, or the whole
// script is killed once the platform's total execution-time limit (6
// minutes for a consumer/free account) is reached. This function cannot
// impose a true timeout on any individual UrlFetchApp call -- it can only
// measure how long each one actually took (or that the whole diagnostic
// itself never finished, which is itself a measurement).
//
// Times each of the (up to) four network operations resolveMeetingById_()
// CAN perform (not what resolveMeetingCoreById_()/enrichMeetingFromTdocList_()
// actually do by default -- this diagnostic always probes every source it's
// asked to, regardless of the production conditional-skip logic), so a
// single slow/hanging source can be identified without being masked by, or
// blamed on, any other.
//
// PROD-016: accepts an optional `only` array to run just a subset (e.g.
// `diagnoseMeetingResolverTiming_(86178, ['metadata'])` for GetMeetings
// alone, or `['tdoc']` for TdocList.aspx alone) without running the full
// diagnostic -- valid labels: 'metadata', 'ical', 'tdoc', 'revisions'.
// Omit `only` (or pass null/undefined) to time all four, the original
// PROD-014 behavior.
function diagnoseMeetingResolverTiming_(meetingId, only) {
  const id = meetingId || 86178;
  const wanted = Array.isArray(only) && only.length > 0 ? only : ['metadata', 'ical', 'tdoc', 'revisions'];
  const report = [];

  function timed(label, fn) {
    const start = Date.now();
    try {
      const result = fn();
      const elapsedMs = Date.now() - start;
      const entry = { label: label, elapsedMs: elapsedMs, statusCode: result && result.statusCode !== undefined ? result.statusCode : null, error: null };
      report.push(entry);
      Logger.log(`${label}: ${elapsedMs} ms / HTTP ${entry.statusCode}`);
      return result;
    } catch (e) {
      const elapsedMs = Date.now() - start;
      report.push({ label: label, elapsedMs: elapsedMs, statusCode: null, error: e.message });
      Logger.log(`${label}: ${elapsedMs} ms / ERROR ${e.message}`);
      return null;
    }
  }

  let metadataFetch = null;
  if (wanted.indexOf('metadata') !== -1) {
    metadataFetch = timed('Meeting metadata (GetMeetings)', () => fetchMeetingMetadataById_(id));
  }
  if (wanted.indexOf('ical') !== -1) {
    timed('iCal (GetiCal)', () => fetchMeetingIcalById_(id));
  }
  if (wanted.indexOf('tdoc') !== -1) {
    timed('TDoc list (TdocList.aspx)', () => fetchMeetingTdocListById_(id));
  }

  // The revisions probe target depends on a successfully parsed FTP base
  // from the metadata fetch above -- timed separately here, using the
  // SAME derivation resolveMeetingCoreById_() uses, so this measures the
  // real production candidate URL, not a guess. Requires 'metadata' to
  // also have been requested (and to have succeeded) in this same call.
  if (wanted.indexOf('revisions') !== -1) {
    if (metadataFetch && metadataFetch.statusCode === 200) {
      const metadataParsed = parseMeetingMetadataResponse_(metadataFetch);
      const m = metadataParsed.meeting || {};
      if (m.MtgDocURL) {
        const ftpInfo = normalizeMtgDocUrlToFtpBase_(m.MtgDocURL);
        if (ftpInfo.ftpBase) {
          const candidateResult = deriveRevisionsUrlCandidate_(ftpInfo.ftpBase);
          if (candidateResult.revisionsUrlCandidate) {
            timed('Revisions/drafts probe (' + candidateResult.revisionsUrlCandidate + ')',
              () => fetchRevisionsUrlCandidate_(candidateResult.revisionsUrlCandidate));
          } else {
            Logger.log('Revisions/drafts probe: SKIPPED -- could not derive a candidate (' + candidateResult.error + ')');
          }
        } else {
          Logger.log('Revisions/drafts probe: SKIPPED -- could not normalize ftpBase from MtgDocURL');
        }
      } else {
        Logger.log('Revisions/drafts probe: SKIPPED -- metadata had no MtgDocURL');
      }
    } else {
      Logger.log('Revisions/drafts probe: SKIPPED -- \'metadata\' was not requested in this call, or did not return HTTP 200');
    }
  }

  Logger.log('--- diagnoseMeetingResolverTiming_ summary ---');
  Logger.log(JSON.stringify(report, null, 2));
  return report; // Execution-log/manual-inspection only -- never persisted.
}

// =========================================================
// ARCH-010 -- MEETING-ID RESOLVE -> PREVIEW -> SAVE CONFIGURATION
// =========================================================
//
// Wires resolveMeetingById_() into the Configure Meeting Settings dialog
// as an EXPLICIT, configuration-time-only operation:
//
//   Meeting ID -> Resolve -> Preview -> user reviews/edits -> Save ->
//   Document Properties -> existing getMeetingContext_()/report generation
//
// getMeetingContext_() itself is NOT changed and does NOT call the Portal
// -- resolution only ever happens when the user clicks "Resolve" inside
// this dialog. Once Document Properties are saved, everything downstream
// is exactly the existing, unchanged, deterministic
// getReportConfig_()/getMeetingContext_() pipeline.
//
// Three pure functions carry the actual merge/readiness POLICY (fully
// testable in Node, no PropertiesService/HtmlService/DocumentApp
// involved):
//   - computeMeetingDateFromStartDate_(startDate)
//   - computeResolvedMeetingPreview_(existingProps, resolverResult)
//   - evaluateMeetingReadiness_(fields, rules, options)
// plus one impure orchestrator (resolveMeetingForConfigDialog_) that reads
// current Document Properties and calls resolveMeetingById_() -- it NEVER
// writes. configureMeetingSettings()/saveConfigurationSettings() (below,
// existing functions) are extended, not replaced.

/**
 * ARCH-010: extracts only the CALENDAR DATE portion of a Portal
 * `StartDate` string ("YYYY-MM-DD HH:MM:SS", as returned by GetMeetings)
 * and formats it exactly like this project's existing MEETING_DATE
 * convention (e.g. "September 22, 2026"). Deliberately ignores time and
 * EndDate entirely -- per the ARCH-010 requirement, MEETING_DATE is only
 * ever derived from the start calendar date, and the unusual multi-day
 * 86178 StartDate/EndDate range is preserved untouched elsewhere (the
 * resolver's own raw/normalized startDate/endDate fields) and never
 * "explained" or collapsed into a single-day assumption here.
 */
function computeMeetingDateFromStartDate_(startDate) {
  const s = String(startDate || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const year = parseInt(m[1], 10);
  const monthIndex = parseInt(m[2], 10) - 1;
  const day = parseInt(m[3], 10);
  const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  if (monthIndex < 0 || monthIndex > 11 || day < 1 || day > 31) return null;
  return `${MONTH_NAMES[monthIndex]} ${day}, ${year}`;
}

/**
 * ARCH-010: pure merge of (current Document Properties, a
 * resolveMeetingById_() result) into a preview object the dialog can
 * render and pre-fill its editable inputs from -- one entry per primary
 * field, each `{ value, source }` where `source` is exactly one of
 * `"resolved"` (came from the Portal just now), `"existing"` (Portal had
 * nothing useful, the prior manually-configured value was kept), or
 * `"unresolved"` (neither source has anything -- the field is genuinely
 * blank and needs manual entry).
 *
 * The core rule, applied per field: a non-empty resolved value always
 * wins; ONLY when the resolver has nothing does the existing property
 * value survive into the preview. This is exactly what prevents a
 * resolver gap (e.g. meeting 86174's missing agenda TDoc, or ANY meeting's
 * always-missing mailing list) from ever proposing a blank value over an
 * already-configured manual one -- resolverResult may be null (Resolve was
 * never clicked, e.g. opening the dialog on an existing legacy document)
 * and every field then falls straight through to `existing`.
 *
 * agendaTdoc is a special case: pre-filled ONLY when the resolver found
 * exactly one current candidate. When it found several
 * (`documents.agendaAmbiguous`), NONE is pre-filled -- the candidate list
 * is surfaced separately (`agendaCandidates`) for the user to choose from
 * explicitly; this function never silently picks one.
 */
function computeResolvedMeetingPreview_(existingProps, resolverResult) {
  const props = existingProps || {};
  const resolved = resolverResult || null;

  function field(resolvedValue, existingValue) {
    if (resolvedValue !== null && resolvedValue !== undefined && String(resolvedValue).trim() !== '') {
      return { value: String(resolvedValue), source: 'resolved' };
    }
    if (existingValue !== null && existingValue !== undefined && String(existingValue).trim() !== '') {
      return { value: String(existingValue), source: 'existing' };
    }
    return { value: '', source: 'unresolved' };
  }

  const resolvedMeeting = resolved && resolved.meeting ? resolved.meeting : null;
  const resolvedSources = resolved && resolved.sources ? resolved.sources : null;
  const resolvedDocuments = resolved && resolved.documents ? resolved.documents : null;

  const meetingIdValue = resolved && resolved.id !== null && resolved.id !== undefined ? resolved.id : null;
  const meetingDateValue = resolvedMeeting && resolvedMeeting.startDate ? computeMeetingDateFromStartDate_(resolvedMeeting.startDate) : null;

  let agendaTdocResolvedValue = null;
  let agendaCandidates = [];
  if (resolvedDocuments) {
    if (resolvedDocuments.agendaTdoc) {
      agendaTdocResolvedValue = resolvedDocuments.agendaTdoc;
    } else if (resolvedDocuments.agendaAmbiguous && Array.isArray(resolvedDocuments.agendaCandidates) && resolvedDocuments.agendaCandidates.length > 0) {
      agendaCandidates = resolvedDocuments.agendaCandidates;
    }
  }

  /**
   * PROD-014: revisionsUrl has a FOURTH provenance state ("candidate") the
   * generic `field()` helper above doesn't have, because resolveMeetingById_()
   * no longer synchronously validates the candidate it derives (see that
   * function's PROD-014 note) -- so `resolvedSources.revisionsUrl` is now
   * always null from a normal Resolve, and only ever non-null via a
   * separate, explicit validateRevisionsUrlCandidate_() call. Priority:
   * an actually-validated value always wins ("resolved"); failing that, an
   * existing manually configured value is preserved ("existing" -- a real,
   * previously-confirmed value outranks a mere unvalidated guess); failing
   * that, a derived-but-unvalidated candidate is shown, clearly labeled as
   * such, NEVER as "resolved"; otherwise genuinely "unresolved". A
   * resolver warning/403/timeout can therefore never erase an already
   * configured REVISIONS_URL, and an unvalidated candidate can never be
   * mistaken for a confirmed one.
   */
  function revisionsUrlField() {
    const resolvedValue = resolvedSources ? resolvedSources.revisionsUrl : null;
    if (resolvedValue !== null && resolvedValue !== undefined && String(resolvedValue).trim() !== '') {
      return { value: String(resolvedValue), source: 'resolved' };
    }
    const existingValue = props.REVISIONS_URL;
    if (existingValue !== null && existingValue !== undefined && String(existingValue).trim() !== '') {
      return { value: String(existingValue), source: 'existing' };
    }
    const candidateValue = resolvedSources ? resolvedSources.revisionsUrlCandidate : null;
    if (candidateValue !== null && candidateValue !== undefined && String(candidateValue).trim() !== '') {
      return { value: String(candidateValue), source: 'candidate' };
    }
    return { value: '', source: 'unresolved' };
  }

  const preview = {
    meetingId: field(meetingIdValue, props.MEETING_ID),
    meetingType: field(resolvedMeeting ? resolvedMeeting.type : null, props.MEETING_TYPE),
    meetingName: field(resolvedMeeting ? resolvedMeeting.name : null, props.MEETING_NAME),
    meetingDate: field(meetingDateValue, props.MEETING_DATE),
    ftpBase: field(resolvedSources ? resolvedSources.ftpBase : null, props.FTP_BASE),
    agendaTdoc: field(agendaTdocResolvedValue, props.AGENDA_TDOC),
    // The resolver NEVER proposes a mailing list (ARCH-008 §6/§9/§11) --
    // this field can only ever be "existing" or "unresolved", never
    // "resolved". Explicit `field(null, ...)` documents that, rather than
    // omitting the field.
    mailingList: field(null, props.MAILING_LIST),
    revisionsUrl: revisionsUrlField(),

    // Supplementary evidence, not itself a Document Property:
    portalType: resolvedMeeting ? resolvedMeeting.portalType : null,
    agendaCandidates: agendaCandidates,
    startDateRaw: resolvedMeeting ? resolvedMeeting.startDate : null,
    endDateRaw: resolvedMeeting ? resolvedMeeting.endDate : null,
    warnings: resolved ? resolved.warnings : [],
    unresolved: resolved ? resolved.unresolved : [],
    location: resolvedMeeting && resolvedMeeting.location ? resolvedMeeting.location : null
  };
  // ADDON-008A: ad-hoc TDoc list and agenda-structure provenance. A saved
  // TDoc list URL always wins over the discovered one.
  const adhocSources = resolved && resolved.adhocSources ? resolved.adhocSources : null;
  if (props.TDOC_LIST_URL && String(props.TDOC_LIST_URL).trim()) {
    preview.tdocListUrl = { value: String(props.TDOC_LIST_URL).trim(), source: 'existing' };
  } else if (adhocSources && adhocSources.tdocList.ok) {
    preview.tdocListUrl = { value: adhocSources.tdocList.url, source: 'resolved' };
  } else {
    preview.tdocListUrl = { value: '', source: 'unresolved' };
  }
  if (adhocSources && adhocSources.agendaCsv.ok) {
    preview.agendaCsvUrl = { value: adhocSources.agendaCsv.url, source: 'resolved' };
  } else if (props.AGENDA_CSV_URL && String(props.AGENDA_CSV_URL).trim() && !(adhocSources && adhocSources.agendaCsv.checked)) {
    preview.agendaCsvUrl = { value: String(props.AGENDA_CSV_URL).trim(), source: 'existing' };
  } else {
    preview.agendaCsvUrl = { value: '', source: 'unresolved' };
  }
  preview.tdocListStatus = adhocSources
    ? (adhocSources.tdocList.ok
      ? 'Meeting-specific Portal document list (' + adhocSources.tdocList.tdocCount + ' TDocs).'
      : 'Meeting-specific Portal document list not usable: ' + adhocSources.tdocList.reason)
    : '';
  preview.agendaSourceStatus = describeAdhocAgendaSource_(preview.agendaTdoc.value, adhocSources, preview.agendaCsvUrl);

  preview.familyInference = inferReportFamily_({
    meetingType: preview.meetingType.value,
    meetingName: preview.meetingName.value,
    ftpBase: preview.ftpBase.value,
    tdocFamilyKeys: resolvedDocuments ? resolvedDocuments.familiesSeen : null
  });
  preview.summary = buildMeetingSummary_(preview);
  return preview;
}

/** ADDON-008A: one plain sentence on where the ad-hoc agenda structure comes from. */
function describeAdhocAgendaSource_(agendaTdoc, adhocSources, agendaCsvField) {
  if (agendaTdoc && String(agendaTdoc).trim()) return 'Agenda TDoc: ' + String(agendaTdoc).trim();
  const csv = adhocSources ? adhocSources.agendaCsv : null;
  if (csv && csv.ok) {
    return 'Agenda structure: validated agenda.csv fallback (no agenda TDoc yet) — ' + csv.selectedCount +
      ' of ' + csv.totalCount + ' items, section(s) ' + csv.sections.join(', ') + '.';
  }
  if (agendaCsvField && agendaCsvField.source === 'existing') return 'Agenda structure: saved agenda.csv fallback.';
  if (csv && csv.checked) return 'No agenda TDoc found, and the agenda.csv fallback is not usable: ' + csv.reason;
  return '';
}

/**
 * ADDON-007B3: the ONE canonical "is this meeting ready to build?" rule set.
 * Declarative on purpose, so the exact same rules drive the dialog's live
 * status (rendered into the client with the evaluator below), the save
 * feedback and the build boundary.
 *
 * Each rule: `anyOf` -- at least one of these fields must be non-blank;
 * `appliesTo` -- meeting types the rule is checked for ('unresolved' = no
 * meeting type known yet); `severity` -- 'blocking' or 'warning'; `buildFor`
 * -- meeting types for which the BUILD ENTRY POINTS refuse to run when the
 * rule fails. Checked against the real runtime: an ad-hoc build reads its
 * sources raw (no main-meeting fallbacks) and its family, agenda TDoc, TDoc
 * list and document folder have no working default, so those are enforced at
 * build. Main meetings keep every existing fallback (Meeting Folder/Number,
 * TDoc-list derivation, template agenda), so NO rule is enforced for main at
 * build -- the dialog is only stricter for a newly configured document.
 * Meeting identity/name, the mailing list, Revisions URL, Reviewer token and
 * email preview are never build-blocking.
 */
const MEETING_READINESS_RULES_ = [
  { code: 'MEETING_NOT_RESOLVED', appliesTo: ['unresolved'], severity: 'blocking', anyOf: [], buildFor: [],
    message: 'Resolve the 3GPP meeting first.' },
  { code: 'MEETING_NOT_RESOLVED', appliesTo: ['adhoc'], severity: 'blocking', anyOf: ['meetingId'], buildFor: [],
    message: 'Resolve the 3GPP meeting first.' },
  { code: 'REPORT_FAMILY_REQUIRED', appliesTo: ['unresolved', 'adhoc', 'main'], severity: 'blocking', anyOf: ['reportFamily'], buildFor: ['adhoc'],
    message: 'Select a report family.' },
  { code: 'MEETING_NAME_REQUIRED', appliesTo: ['adhoc'], severity: 'blocking', anyOf: ['meetingName'], buildFor: [],
    message: 'The meeting name is missing. Resolve the meeting or enter it under Advanced.' },
  // ADDON-008A: an agenda STRUCTURE source is required -- an agenda TDoc or
  // a validated agenda.csv (code and message kept for 007B3 compatibility).
  { code: 'AGENDA_TDOC_REQUIRED', appliesTo: ['adhoc'], severity: 'blocking', anyOf: ['agendaTdoc', 'agendaCsvUrl'], buildFor: ['adhoc'],
    message: 'Discover or select the agenda document.' },
  { code: 'TDOC_LIST_REQUIRED', appliesTo: ['adhoc'], severity: 'blocking', anyOf: ['tdocListUrl'], buildFor: ['adhoc'],
    message: 'Paste the meeting\'s TDoc list URL.' },
  { code: 'FTP_BASE_REQUIRED', appliesTo: ['adhoc'], severity: 'blocking', anyOf: ['ftpBase'], buildFor: ['adhoc'],
    message: 'The meeting document folder could not be determined. Review the meeting source under Advanced.' },
  { code: 'MAIN_MEETING_FOLDER_REQUIRED', appliesTo: ['main'], severity: 'blocking', anyOf: ['meetingFolder', 'ftpBase'], buildFor: [],
    message: 'Enter the meeting folder, or resolve the meeting.' },
  { code: 'MAIN_MEETING_NUMBER_REQUIRED', appliesTo: ['main'], severity: 'blocking', anyOf: ['meetingNumber', 'tdocListUrl'], buildFor: [],
    message: 'Enter the meeting number, or paste the TDoc list URL.' },
  { code: 'MAILING_LIST_MISSING', appliesTo: ['adhoc', 'main'], severity: 'warning', anyOf: ['mailingList'], buildFor: [],
    message: 'No mailing list is configured. Email collection will be unavailable.' }
];

/**
 * ADDON-007B3: pure evaluator for MEETING_READINESS_RULES_. Self-contained
 * (no helpers, no globals, no backslashes) because the dialog ships this
 * function's own source to the browser -- one implementation, not a client
 * copy of a server rule.
 *
 * fields: { meetingType, meetingId, meetingName, ftpBase, reportFamily,
 * agendaTdoc, agendaCsvUrl, tdocListUrl, mailingList, meetingFolder,
 * meetingNumber }.
 * A blank meetingType means a legacy main document only when it has a
 * folder or number; otherwise the meeting is simply not resolved yet.
 * options.forBuild: only rules enforced at the build boundary for this type.
 */
function evaluateMeetingReadiness_(fields, rules, options) {
  const f = fields || {};
  const forBuild = !!(options && options.forBuild);
  function has(name) {
    const v = f[name];
    return String(v === null || v === undefined ? '' : v).trim() !== '';
  }
  let type = String(f.meetingType || '').trim().toLowerCase();
  if (!type) type = (has('meetingFolder') || has('meetingNumber')) ? 'main' : 'unresolved';

  const issues = [];
  const warnings = [];
  (rules || []).forEach(function (rule) {
    if (rule.appliesTo.indexOf(type) === -1) return;
    if (forBuild && rule.buildFor.indexOf(type) === -1) return;
    if (rule.anyOf.some(has)) return;
    (rule.severity === 'warning' ? warnings : issues).push({ code: rule.code, message: rule.message });
  });
  return { ready: issues.length === 0, meetingType: type, issues: issues, warnings: warnings };
}

/**
 * ADDON-007B3: build-boundary readiness, read from the SAME state the build
 * itself reads (document properties, or the central copy for a background
 * context). Ad-hoc sources are read raw -- exactly like getMeetingContext_()
 * -- so a missing value is missing here instead of being masked by a
 * main-meeting fallback.
 */
function getBuildReadiness_(context) {
  const identity = getMeetingIdentityConfig_(context);
  const props = getReportStateStore_(context);
  return evaluateMeetingReadiness_({
    meetingType: identity.MEETING_TYPE,
    meetingName: identity.MEETING_NAME,
    ftpBase: identity.FTP_BASE,
    reportFamily: props.getProperty('REPORT_SUFFIX'),
    agendaTdoc: identity.AGENDA_TDOC,
    // ADDON-008A: only the agenda.csv of this meeting's own series folder
    // counts; its content is re-validated by the build itself.
    agendaCsvUrl: isAdhocAgendaCsvUrlFor_(identity.AGENDA_CSV_URL, identity.FTP_BASE) ? identity.AGENDA_CSV_URL : '',
    tdocListUrl: identity.TDOC_LIST_URL
  }, MEETING_READINESS_RULES_, { forBuild: true });
}

/**
 * ADDON-007B3: refuses to build/update an ad-hoc report whose required
 * sources are missing -- before anything in the document is touched. Main
 * meetings are never blocked here (legacy fallbacks keep working).
 */
function assertMeetingReadyToBuild_(context) {
  // TEMPLATE-002B: never build or update in the master template (a
  // background context is the CENTRAL scheduler, which has no template).
  if (!context) assertNotTemplateMaster_();
  const readiness = getBuildReadiness_(context);
  if (readiness.ready) return;
  throw new Error('Cannot build report yet.\n\n' +
    readiness.issues.map(function (issue) { return '• ' + issue.message; }).join('\n') +
    '\n\nOpen Configure Meeting to finish the setup.');
}

/**
 * PROD-016: reads current Document Properties (read-only) and calls
 * resolveMeetingCoreById_() -- NOT the full resolveMeetingById_() -- so
 * this returns as soon as GetMeetings (plus, rarely, GetiCal) succeeds,
 * WITHOUT ever touching TdocList.aspx. That fetch was found live to make
 * Resolve hang 90+ seconds even after PROD-014 removed the revisions
 * probe; agenda/TDoc discovery is now the separate "Discover Agenda /
 * TDocs" action (discoverAgendaForConfigDialog_(), below). NEVER writes to
 * PropertiesService -- resolution and persistence are deliberately
 * separate actions (clicking Resolve alone must never change saved
 * configuration).
 *
 * RESOLVER-HOTFIX: this is the internal implementation only -- the
 * dialog's "Resolve" button does NOT call this directly any more. Apps
 * Script's google.script.run can only invoke a PUBLIC top-level function;
 * a function name ending in "_" is treated as private by Apps Script
 * convention (hidden from the IDE's function selector, not assignable as
 * a trigger handler, and NOT invokable via google.script.run from client
 * HTML) -- this function's trailing underscore meant the client's RPC
 * call to it was silently never dispatched at all (confirmed live: Apps
 * Script Executions showed zero executions of this function after
 * clicking Resolve, only the unrelated configureMeetingSettings() call
 * that renders the dialog itself). See resolveMeetingForConfigDialog()
 * (no trailing underscore) below for the actual public RPC entry point;
 * this internal function, its name, and every existing test against it
 * are otherwise unchanged.
 */
function resolveMeetingForConfigDialog_(meetingIdInput) {
  // POST-MEETING-001 (Task A5): manual-diagnostic stage markers only --
  // visible in the Apps Script execution transcript when this is run
  // (from the dialog, or directly from the editor), never required for
  // normal operation. Pairs with tests/resolve-dialog-client-rendering.test.js's
  // client-side "client success handler entered"/"client preview render
  // completed" console.log markers, so a later manual investigation can
  // distinguish "server entered" / "before GetMeetings" / "after
  // GetMeetings" / "before response return" / "response returned to
  // client" / "client success handler entered" / "client preview render
  // completed" as seven distinct, independently-timestamped points.
  Logger.log('resolveMeetingForConfigDialog_: server entered, meetingIdInput=' + meetingIdInput);
  const props = PropertiesService.getDocumentProperties();
  const existing = {
    MEETING_ID: props.getProperty('MEETING_ID'),
    MEETING_TYPE: props.getProperty('MEETING_TYPE'),
    MEETING_NAME: props.getProperty('MEETING_NAME'),
    MEETING_DATE: props.getProperty('MEETING_DATE'),
    FTP_BASE: props.getProperty('FTP_BASE'),
    AGENDA_TDOC: props.getProperty('AGENDA_TDOC'),
    AGENDA_CSV_URL: props.getProperty('AGENDA_CSV_URL'),
    MAILING_LIST: props.getProperty('MAILING_LIST'),
    TDOC_LIST_URL: props.getProperty('TDOC_LIST_URL'),
    REVISIONS_URL: props.getProperty('REVISIONS_URL')
  };

  const idResult = parseMeetingIdInput_(meetingIdInput);
  if (!idResult.isValid) {
    Logger.log('resolveMeetingForConfigDialog_: invalid meeting id, before response return');
    return {
      ok: false,
      error: idResult.error,
      preview: computeResolvedMeetingPreview_(existing, null)
    };
  }

  Logger.log('resolveMeetingForConfigDialog_: before resolveMeetingCoreById_ (includes GetMeetings [+ conditional GetiCal])');
  const resolved = resolveMeetingCoreById_(idResult.id);
  Logger.log('resolveMeetingForConfigDialog_: after resolveMeetingCoreById_, before response return');
  return {
    ok: true,
    error: null,
    resolved: resolved,
    preview: computeResolvedMeetingPreview_(existing, resolved)
  };
}

/**
 * RESOLVER-HOTFIX: the actual public google.script.run entry point the
 * Configure Meeting Settings dialog's "Resolve" button calls -- see
 * resolveMeetingForConfigDialog_()'s own header comment above for why a
 * trailing underscore made the original direct RPC target silently
 * uncallable. A thin, otherwise-behavior-free delegation, so the tested,
 * documented internal implementation and its name are unchanged.
 *
 * MEETING-RESOLVER-FINAL: live-verified against meeting 86178 in the real
 * Google Doc -- Resolve, Discover Agenda/TDocs, Save, and a close/reopen
 * round-trip all passed, confirming this fix in production, not just
 * locally.
 */
function resolveMeetingForConfigDialog(meetingIdInput) {
  return resolveMeetingForConfigDialog_(meetingIdInput);
}

/**
 * PROD-016: the dialog's new, SEPARATE "Discover Agenda / TDocs" button
 * calls the public discoverAgendaForConfigDialog() wrapper below (see
 * RESOLVER-HOTFIX), which delegates here. Takes the meetingId AND the
 * previously-resolved CORE result (from resolveMeetingForConfigDialog_(),
 * which the client already has in hand) -- this avoids re-fetching
 * GetMeetings, and its own TdocList.aspx fetch (via
 * enrichMeetingFromTdocList_()) is the ONLY network call it makes. Reads
 * current Document Properties (read-only, same fields as
 * resolveMeetingForConfigDialog_()) and NEVER writes to PropertiesService
 * -- enrichment and persistence are separate actions, exactly like Resolve
 * itself. Refuses (without ever calling TdocList.aspx) if no valid core
 * result for the SAME meeting ID is supplied -- this action only ever
 * enriches an already-resolved meeting, it never resolves one from
 * scratch.
 */
function discoverAgendaForConfigDialog_(meetingIdInput, coreResolved) {
  const props = PropertiesService.getDocumentProperties();
  const existing = {
    MEETING_ID: props.getProperty('MEETING_ID'),
    MEETING_TYPE: props.getProperty('MEETING_TYPE'),
    MEETING_NAME: props.getProperty('MEETING_NAME'),
    MEETING_DATE: props.getProperty('MEETING_DATE'),
    FTP_BASE: props.getProperty('FTP_BASE'),
    AGENDA_TDOC: props.getProperty('AGENDA_TDOC'),
    AGENDA_CSV_URL: props.getProperty('AGENDA_CSV_URL'),
    MAILING_LIST: props.getProperty('MAILING_LIST'),
    TDOC_LIST_URL: props.getProperty('TDOC_LIST_URL'),
    REVISIONS_URL: props.getProperty('REVISIONS_URL')
  };

  const idResult = parseMeetingIdInput_(meetingIdInput);
  if (!idResult.isValid) {
    return {
      ok: false,
      error: idResult.error,
      preview: computeResolvedMeetingPreview_(existing, coreResolved || null)
    };
  }
  if (!coreResolved || !coreResolved.meeting || coreResolved.id !== idResult.id) {
    return {
      ok: false,
      error: 'Resolve the meeting first before discovering agenda/TDocs.',
      preview: computeResolvedMeetingPreview_(existing, coreResolved || null)
    };
  }

  const enriched = enrichMeetingFromTdocList_(idResult.id, coreResolved);
  applyAdhocSourceDiscovery_(enriched, idResult.id, existing);
  return {
    ok: true,
    error: null,
    resolved: enriched,
    preview: computeResolvedMeetingPreview_(existing, enriched)
  };
}

/**
 * ADDON-008A: ad-hoc only -- adds the meeting-ID document list and, when no
 * agenda TDoc is available (none saved, none discovered, no candidates),
 * the validated agenda.csv fallback to an enriched resolve result. When the
 * fallback validates, "no agenda TDoc" is no longer reported as a problem.
 */
function applyAdhocSourceDiscovery_(enriched, meetingId, existing) {
  if (!enriched || !enriched.meeting || enriched.meeting.type !== 'adhoc') return;
  const docs = enriched.documents || {};
  const hasAgendaTdoc = !!docs.agendaTdoc || (docs.agendaCandidates || []).length > 0 ||
    !!String((existing && existing.AGENDA_TDOC) || '').trim();
  const ftpBase = enriched.sources ? enriched.sources.ftpBase : null;
  const adhocSources = discoverAdhocMeetingSources_(meetingId, ftpBase, !hasAgendaTdoc);
  enriched.adhocSources = adhocSources;
  if (adhocSources.agendaCsv.ok) {
    enriched.warnings = (enriched.warnings || []).filter(function (w) { return w !== ADHOC_NO_AGENDA_TDOC_WARNING_; });
  }
}

/**
 * RESOLVER-HOTFIX: the actual public google.script.run entry point the
 * Configure Meeting Settings dialog's "Discover Agenda / TDocs" button
 * calls -- see resolveMeetingForConfigDialog()'s header comment for why a
 * trailing underscore makes a function uncallable via google.script.run.
 * A thin, otherwise-behavior-free delegation.
 */
function discoverAgendaForConfigDialog(meetingIdInput, coreResolved) {
  return discoverAgendaForConfigDialog_(meetingIdInput, coreResolved);
}

// =========================================================
// ADDON-009 -- DISCUSSION E-MAIL (.eml) EXPORTER
// (ported from Legacy LEGACY-UPGRADE-006, accepted at Legacy 80ab081)
// =========================================================
//
// ADDON-009: the Legacy exporter, ported as-is. The only CENTRAL
// adaptations are:
//   - the subject's list tag is no longer the constant 'FS_6G_MED' but is
//     derived from the meeting's mailing list (deriveEmailExportListTag_(),
//     which yields 'FS_6G_MED' for 3GPP_TSG_SA4_FS_6G_MED) and is threaded
//     through the builders as a parameter (also used for the .eml/ZIP file
//     names and the default introduction);
//   - the recipient is derived from getMeetingContext_().sources.mailingList
//     (an ad-hoc meeting's saved MAILING_LIST, else the report-family list;
//     main meetings unchanged) instead of the raw MAILING_LIST property;
//   - the TDoc identifier is the canonical parseExactSA4DocumentId_() form,
//     re-checked at Generate time;
//   - Generate validates and builds every selected e-mail in memory before
//     it writes anything to Drive, so a refused request leaves no files.
// Everything else below (status rules, deadline, subject, body, MIME, ZIP,
// dialog) is the Legacy code; the LEGACY-UPGRADE-006x notes describe it.
//
// Isolated, READ-ONLY export feature. It never clears/rebuilds/mutates the
// report document, never touches Document Properties, never creates a
// trigger, never calls continuousUpdate()/any build function, and never
// sends mail (Gmail/MailApp are never called) -- it only READS the
// already-built TDoc tables and existing meeting configuration, and WRITES
// new .eml files into a dedicated Drive folder. The Google Doc report
// remains the sole source of truth: TDoc-table content is copied into the
// generated e-mail's HTML, never regenerated, summarized, or reconstructed
// from any other source.
//
// Existing TDoc tables (see createTDocTableFromData_()/insertNewTdoc_()/
// appendTdocDetailTable_()) are always simple two-column label/value
// tables: row 0 is always ['TDoc', <tdocNumber>] (isTDocTable_()'s own
// check), followed by Title/Source/Contact/Agenda Item/Type-For/
// (optional Abstract)/E-mail Discussion/Revisions/Minutes/Disposition/
// Status -- there is no separate "Decision" row in this legacy codebase.
// "E-mail Discussion"/"Minutes"/"Revisions" cells accumulate their content
// as MULTIPLE LINES via Text.appendText() with embedded "\n" (see
// checkRSSFeed_()/updateRevisions_()) -- Apps Script represents each such
// line as its own paragraph inside the cell, so walking the cell's child
// elements (getNumChildren()/getChild(i)) already gives one paragraph per
// discussion entry, exactly preserving paragraph boundaries. Hyperlinks
// (author mailto/thread links, revision links) and bold (status styling,
// see styleStatusCell_()) are applied as per-character-offset TEXT
// attributes via Text.setLinkUrl()/setBold() -- read back here with the
// exact mirror-image per-offset query methods (Text.getLinkUrl(offset)/
// isBold(offset)/isItalic(offset)/isUnderline(offset)), segmented at each
// Text.getTextAttributeIndices() breakpoint. No merged/colspan cell is
// ever created by any TDoc-table builder in this file (every row is
// exactly two plain appendTableCell() calls) -- DocumentApp's own Table
// API has no colspan/merge-span accessor at all, so this is a genuine,
// but currently inapplicable, Apps Script platform limitation, documented
// here rather than silently ignored.
//
// Kept in Code.js (as in Legacy): tests/helpers/load-code.js loads exactly
// this one file into its sandbox.

// ---- centralized wording (change here, not scattered in the builders) ----

/**
 * ADDON-009: the reflector/list tag used in the canonical subject's first
 * bracket group (Legacy: the constant EMAIL_EXPORT_LIST_TAG_ = 'FS_6G_MED').
 * Derived from the already-validated recipient address -- the meeting's
 * own reflector -- by dropping the SA4 list prefix, upper-cased:
 *
 *   3gpp_tsg_sa4_fs_6g_med@list.etsi.org -> FS_6G_MED
 *   3gpp_tsg_sa_wg4_mbs@list.etsi.org    -> MBS
 *
 * The general SA4 list (3GPP_TSG_SA_WG4, main Liaison/New reports) names
 * no topic, so the report family's SWG name (getReportConfig_().DRAFTS_FOLDER,
 * e.g. 6G -> FS_6G_MED) is used instead. Returns null when nothing usable
 * remains; the tag may never contain a comma, bracket or whitespace.
 */
function deriveEmailExportListTag_(recipientAddress, reportSwgName) {
  const local = String(recipientAddress || '').trim().split('@')[0].toUpperCase();
  const topic = local.replace(/^3GPP_TSG_SA(?:4|_WG4)(?:_|$)/, '') ||
    String(reportSwgName || '').trim().toUpperCase();
  return /^[A-Z0-9_.+-]+$/.test(topic) ? topic : null;
}

// LEGACY-UPGRADE-006B (bulk text controls): the DEFAULT plain-text wording
// for the export dialog's two editable, GLOBAL text areas (Introduction /
// Discussion request). These are the single source of truth for the
// dialog's prefilled defaults AND for what generateTdocDiscussionEmails()
// falls back to when no override is supplied (e.g. an older/direct caller
// that predates this stage) -- never duplicated as a second HTML copy.
// Deliberately plain TEXT, not HTML: converted to safe HTML per email via
// plainTextToSafeHtmlParagraphs_() at generation time, exactly like a
// user's own edited wording is. A blank line (\n\n) is a paragraph break;
// a single \n is a <br> within a paragraph.
//
// ADDON-009: the Legacy wording, made meeting-neutral: its "FS_6G_MED" is
// the meeting's list tag (buildEmailExportDefaultIntroText_()) and "the
// October meeting" is "the upcoming meeting".
const EMAIL_EXPORT_DEFAULT_INTRO_TEXT_TEMPLATE_ =
  'Dear all,\n\n' +
  'As discussed during the {LIST_TAG} AHG, this email starts a technical discussion on the contribution below. ' +
  'The purpose is to collect comments, refine the proposal and, where appropriate, prepare a revision for the upcoming meeting.\n\n' +
  'This discussion is not an email agreement and does not constitute a formal SA4 decision.';

function buildEmailExportDefaultIntroText_(listTag) {
  return EMAIL_EXPORT_DEFAULT_INTRO_TEXT_TEMPLATE_.replace('{LIST_TAG}', String(listTag || '').trim());
}

const EMAIL_EXPORT_DISCUSSION_HEADING_HTML_ = 'Discussion';

// LEGACY-UPGRADE-006B (Part F, deadline in body): the closing
// "Formal decisions remain reserved..." sentence used to be the last line
// of the (user-editable) default Discussion-request text. It is now a
// FIXED, exporter-controlled sentence (EMAIL_EXPORT_FORMAL_DECISIONS_NOTE_HTML_
// below), placed after the deadline sentence, so a user editing the global
// Discussion-request text can never accidentally remove it (or the
// deadline sentence next to it) -- per this stage's explicit architectural
// requirement. The user-editable default text itself is unchanged apart
// from no longer repeating that sentence.
const EMAIL_EXPORT_DEFAULT_DISCUSSION_TEXT_ =
  'Comments are invited on the contribution and on the issues captured in the meeting minutes above.\n\n' +
  'Please provide comments, proposed changes and, where appropriate, proposed text for a revision on this thread.';

const EMAIL_EXPORT_FORMAL_DECISIONS_NOTE_HTML_ = 'Formal decisions remain reserved for the normal SA4 meeting process.';

const EMAIL_EXPORT_CLOSING_HTML_ = 'Best regards,<br>Thomas';

const EMAIL_EXPORT_DRIVE_FOLDER_NAME_ = 'SA4 Report Email Exports';

/**
 * LEGACY-UPGRADE-006 (status-filter follow-up): a TDoc whose report Status
 * is one of these is treated as already completed and is not offered for
 * a new FS_6G_MED discussion. Centralized here deliberately -- no scattered
 * string comparisons. Comparison is always case-insensitive and trimmed
 * (see isEmailExportStatusExcluded_()). Deliberately narrow: only the two
 * statuses actually specified are excluded; nothing else is inferred.
 */
const EMAIL_EXPORT_EXCLUDED_STATUSES_ = ['approved', 'agreed'];

/**
 * LEGACY-UPGRADE-006 (status-filter follow-up): true when `status` (the
 * TDoc table's own, UNMODIFIED Status cell text) matches one of
 * EMAIL_EXPORT_EXCLUDED_STATUSES_, case-insensitively and trimmed. A
 * blank/missing status is never excluded (nothing to match).
 */
function isEmailExportStatusExcluded_(status) {
  const normalized = String(status || '').trim().toLowerCase();
  if (!normalized) return false;
  return EMAIL_EXPORT_EXCLUDED_STATUSES_.indexOf(normalized) !== -1;
}

/**
 * LEGACY-UPGRADE-006B (reserved-status hard exclusion): "reserved" is an
 * EXISTING, code-defined status in this codebase, not invented for this
 * stage -- normalizeStatus_() (used by applyTdocStatusUpdate_(), the
 * status-update decision logic shared by the fast/fallback paths of
 * updateTdocStatus_()) already recognizes it as one of exactly two
 * "not yet finalized" states (the other being "available"): a "reserved"
 * TDoc slot has been allocated/registered in the tracking sheet but has no
 * real submitted content yet, which is why its Status cell is still
 * eligible to be freely overwritten on every update. That same reason
 * (no real content to discuss) is why a "reserved" TDoc must never be
 * offered, nor accepted, for a discussion e-mail -- distinct from
 * Approved/Agreed (which DO have content, but are already decided).
 * Comparison is case-insensitive and trimmed, matching exactly (not a
 * substring match like normalizeStatus_()'s own broader use) -- this
 * function's only job is "is the Status cell literally reserved", not
 * "does it contain the word reserved somewhere". A blank/missing status
 * is never treated as reserved (nothing to match).
 *
 * A codebase-wide check for OTHER existing "unavailable"/placeholder
 * status concepts (per this stage's own requirement) found exactly one:
 * "available" (normalizeStatus_()'s other recognized value) -- which
 * means the OPPOSITE of unavailable (a TDoc slot ready to receive
 * content) and is correctly never excluded. No other status-like
 * "not really there yet" concept exists anywhere else in this file.
 */
const EMAIL_EXPORT_RESERVED_STATUS_ = 'reserved';

function isEmailExportReserved_(status) {
  const normalized = String(status || '').trim().toLowerCase();
  if (!normalized) return false;
  return normalized === EMAIL_EXPORT_RESERVED_STATUS_;
}

// ---- discussion deadline (LEGACY-UPGRADE-006B, Parts D-H) ----

/**
 * LEGACY-UPGRADE-006B: default DISCUSSION deadline for the current
 * FS_6G_MED batch -- explicitly a request-for-comments deadline, NOT an
 * Email Agreement deadline (a separate, future, stricter workflow; see
 * this stage's own scope note above generateTdocDiscussionEmails()).
 * Stored as three separate, unambiguous fields (never a single free-text
 * string) so validation can check each independently. The human-visible
 * format built from these (formatEmailExportDeadline_()) is always
 * "YY-MM-DD HH:mm TZ" -- deliberately never a natural-language or
 * locale-dependent rendering.
 *
 * ADDON-009: no default date or time. Legacy prefilled its FS_6G_MED batch
 * deadline (2026-10-15 15:00); CENTRAL has no stored meeting end date
 * (MEETING_DATE is an optional display string, the Portal dates are only
 * shown during Resolve) and no approved rule for deriving a discussion
 * deadline from one, so the user enters it and Generate refuses a missing
 * one (validateEmailExportDeadline_()). The time zone keeps the only
 * supported value.
 */
const EMAIL_EXPORT_DEFAULT_DEADLINE_DATE_ = ''; // ISO YYYY-MM-DD, entered by the user
const EMAIL_EXPORT_DEFAULT_DEADLINE_TIME_ = ''; // 24-hour HH:mm, entered by the user
const EMAIL_EXPORT_DEFAULT_DEADLINE_TZ_ = 'CEST';

/**
 * LEGACY-UPGRADE-006B: the explicit, closed set of time zone abbreviations
 * this exporter accepts for a discussion deadline -- deliberately NOT
 * inferred from browser locale or the Apps Script server's own time zone
 * (Session.getScriptTimeZone()), since the deadline's time zone is part of
 * the deadline itself, not an environment detail. Only CEST is needed for
 * the current FS_6G_MED batch; extending this set is a deliberate later
 * decision, not something to guess at now.
 */
const EMAIL_EXPORT_SUPPORTED_TIMEZONES_ = ['CEST'];

function isSupportedEmailExportTimezone_(tz) {
  return EMAIL_EXPORT_SUPPORTED_TIMEZONES_.indexOf(String(tz || '').trim()) !== -1;
}

/**
 * LEGACY-UPGRADE-006B: true only for a real, existing calendar date in
 * strict YYYY-MM-DD form (rejects e.g. "2026-02-30", "2026-13-01", or any
 * non-numeric/mis-shaped input) -- never trusts a client-supplied date
 * string without checking it actually exists on the calendar.
 */
function isValidEmailExportDeadlineDate_(dateStr) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr || '').trim());
  if (!m) return false;
  const y = parseInt(m[1], 10), mo = parseInt(m[2], 10), d = parseInt(m[3], 10);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return false;
  const dt = new Date(Date.UTC(y, mo - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === mo - 1 && dt.getUTCDate() === d;
}

/**
 * LEGACY-UPGRADE-006B: true only for a strict 24-hour HH:mm value
 * ("15:00", "09:05") -- never "3pm", "15.00", or any locale-dependent form.
 */
function isValidEmailExportDeadlineTime_(timeStr) {
  return /^([01]\d|2[0-3]):[0-5]\d$/.test(String(timeStr || '').trim());
}

/**
 * LEGACY-UPGRADE-006B: the ONE validation gate a submitted
 * {date, time, tz} deadline must pass before it is used anywhere (Subject
 * or body) -- date must be a real calendar date, time must be strict
 * 24-hour HH:mm, and tz must be one of the explicit supported set. Never
 * silently substitutes a default for an invalid/missing field; the caller
 * (generateTdocDiscussionEmails()) is expected to refuse to generate that
 * TDoc's e-mail and surface `error` verbatim rather than guessing.
 */
function validateEmailExportDeadline_(deadline) {
  const d = deadline || {};
  const date = String(d.date || '').trim();
  const time = String(d.time || '').trim();
  const tz = String(d.tz || '').trim();
  if (!isValidEmailExportDeadlineDate_(date)) {
    return { valid: false, error: 'Invalid or missing deadline date (expected a real calendar date, YYYY-MM-DD): "' + date + '"' };
  }
  if (!isValidEmailExportDeadlineTime_(time)) {
    return { valid: false, error: 'Invalid or missing deadline time (expected 24-hour HH:mm): "' + time + '"' };
  }
  if (!isSupportedEmailExportTimezone_(tz)) {
    return { valid: false, error: 'Unsupported or missing deadline time zone: "' + tz + '" (supported: ' + EMAIL_EXPORT_SUPPORTED_TIMEZONES_.join(', ') + ')' };
  }
  return { valid: true, date: date, time: time, tz: tz };
}

/**
 * LEGACY-UPGRADE-006B: the ONE place that renders an already-validated
 * deadline into its human-visible "YY-MM-DD HH:mm TZ" form -- used for the
 * body's "Please provide your comments by..."/"Please upload revisions
 * by..." sentences (never the Subject any more -- see
 * formatEmailExportDeadlineForSubjectToken_() below, LEGACY-UPGRADE-006F).
 * Takes the object validateEmailExportDeadline_() returns on success;
 * never called with an unvalidated value.
 */
function formatEmailExportDeadline_(validated) {
  const yy = validated.date.slice(2, 4);
  const mm = validated.date.slice(5, 7);
  const dd = validated.date.slice(8, 10);
  return yy + '-' + mm + '-' + dd + ' ' + validated.time + ' ' + validated.tz;
}

/**
 * LEGACY-UPGRADE-006F: the ONE place that renders an already-validated
 * deadline into its compact canonical-Subject token form
 * "YY-MM-DD-HHmmTZ" (e.g. "26-10-15-1500CEST") -- no spaces, no colon,
 * distinct from formatEmailExportDeadline_()'s human/body "YY-MM-DD HH:mm
 * TZ" form. Both formatters read from the SAME validated {date, time, tz}
 * object (validateEmailExportDeadline_()'s output) -- one authoritative
 * deadline value, two representations, never two separate deadline
 * semantics. Never called with an unvalidated value.
 */
function formatEmailExportDeadlineForSubjectToken_(validated) {
  const yy = validated.date.slice(2, 4);
  const mm = validated.date.slice(5, 7);
  const dd = validated.date.slice(8, 10);
  const hhmm = validated.time.replace(':', '');
  return yy + '-' + mm + '-' + dd + '-' + hhmm + validated.tz;
}

// ---- revision-upload location (LEGACY-UPGRADE-006C) ----

/**
 * LEGACY-UPGRADE-006C: the discussion e-mail's revision-upload destination
 * is NOT a new concept or a new piece of configuration -- it reuses
 * getMeetingContext_().sources.revisionsUrl, the SAME authoritative,
 * already-used-elsewhere source that scanRevisionsFolder_() (the report's
 * own "where are submitted drafts?" scanner, PROD-017) and the
 * "Test All Connections" diagnostic already read. That value is:
 *
 *   - for a MAIN meeting: derived by formula, `${INBOX_BASE}Drafts/${DRAFTS_FOLDER}`,
 *     where DRAFTS_FOLDER comes from the report-family -> SWG-name map
 *     (DRAFTS_FOLDERS, e.g. '6G' -> 'FS_6G_MED') -- differs by report
 *     family, never by this exporter's own logic;
 *   - for an AD-HOC meeting (this script's own kind -- FS_6G_MED is
 *     MEETING_TYPE=adhoc): NO formula exists; it is READ RAW from the
 *     REVISIONS_URL Document Property (set via the "Revisions / Drafts
 *     URL" field in Configure Meeting Settings) and is genuinely absent
 *     (undefined) if that property was never set -- resolveMeetingSources_()
 *     never invents one. Each ad-hoc series has its own dedicated
 *     Inbox/Drafts/ folder, so no further per-family subfolder applies.
 *
 * This function adds NO new resolution logic of its own -- it only
 * re-reads the existing value fresh (never caches, never trusts a
 * client-supplied URL) and trims it. Never fetches the URL (no
 * UrlFetchApp call) -- resolving it is a pure Document Properties/config
 * read, exactly like every other consumer of this value.
 */
function resolveEmailExportRevisionUploadUrl_() {
  return String(getMeetingContext_().sources.revisionsUrl || '').trim();
}

/**
 * LEGACY-UPGRADE-006C: true only for an absolute http(s) URL -- rejects
 * "javascript:", "data:", "mailto:", a bare relative path, or anything
 * else that is not a plain, clickable external link. This is what stands
 * between "trusted meeting configuration" and "a real <a href>"; it is
 * never the ONLY check (the URL is also never taken from client input in
 * the first place -- see generateTdocDiscussionEmails()), but it defends
 * against a misconfigured or malformed REVISIONS_URL property value
 * producing something unsafe or non-clickable.
 */
function isSafeEmailExportUrl_(url) {
  return /^https?:\/\/[^\s"<>]+$/i.test(String(url || '').trim());
}

// ---- discussion e-mail sender and recipient (LEGACY-UPGRADE-006H) ----
//
// LEGACY-UPGRADE-006H: these are TWO DELIBERATELY SEPARATE concepts,
// resolved by two separate functions, on purpose -- 006F/006G's mistake
// was collapsing them into one address:
//
//   - SENDER (From): who the e-mail is actually sent as in Outlook. An
//     explicitly configured, independent value (DISCUSSION_EMAIL_SENDER,
//     e.g. "reporter@example.com" -- a person's real mailbox). Resolved
//     by resolveEmailExportSenderAddress_(). NEVER derived from Mailing
//     List, Session identity, or client input.
//   - RECIPIENT (To): the reflector this discussion is posted to. Derived
//     from the already-configured Mailing List (e.g.
//     "3GPP_TSG_SA4_FS_6G_MED" -> "3gpp_tsg_sa4_fs_6g_med@list.etsi.org").
//     Resolved by deriveEmailExportRecipientFromMailingList_(). NEVER the
//     same value as the sender, and never itself independently configured
//     (Mailing List is already the single source of truth for it).
//
// Both share the SAME final validator (isValidEmailExportSenderAddress_(),
// name kept generic on purpose -- it validates "is this a safe address to
// put in a MIME header", not "is this specifically a sender").

/**
 * LEGACY-UPGRADE-006E, restored by LEGACY-UPGRADE-006H (006F/006G had
 * incorrectly removed this and derived From from Mailing List instead):
 * the discussion e-mail's From address. Ordinary, non-secret configuration
 * -- read fresh, server-side, from the SAME Document Properties mechanism
 * every other Configure Meeting Settings field already uses
 * (DISCUSSION_EMAIL_SENDER, set via that dialog's own "E-mail
 * Configuration" section). Never derived from
 * Session.getActiveUser()/getEffectiveUser() or any other Google login
 * identity, never derived from Mailing List, and never read from
 * client-supplied request data (a `from`/`sender`/`discussionEmailSender`
 * field on a selection or globalText payload is simply never looked at).
 * Genuinely absent (empty string) if never configured -- never invented.
 * A stale value already present on an existing document (e.g. from before
 * 006F removed it) becomes active again automatically, with no migration
 * needed, simply because this function reads the property again.
 */
function resolveEmailExportSenderAddress_() {
  return String(PropertiesService.getDocumentProperties().getProperty('DISCUSSION_EMAIL_SENDER') || '').trim();
}

/**
 * LEGACY-UPGRADE-006F/006G, corrected naming by LEGACY-UPGRADE-006H (was
 * deriveEmailExportSenderFromMailingList_() -- renamed because its result
 * is the discussion RECIPIENT/To address, never the sender/From, and the
 * old name was directly responsible for the 006G regression that
 * conflated the two). Derives the discussion recipient from the
 * already-configured Mailing List, which is the single source of truth
 * for the reflector this batch belongs to. Transformation: trim, reject
 * CR/LF (defense against a malformed/tampered MAILING_LIST value
 * injecting extra MIME headers into the raw "To: " line
 * buildEmlContent_() writes), lowercase, then either:
 *
 *   - already a full "...@list.etsi.org" address (case-insensitive) --
 *     used exactly as-is (lowercased), NEVER double-suffixed
 *     ("...@list.etsi.org@list.etsi.org"), so an existing configuration
 *     that already stores the full address keeps working unchanged; or
 *   - a bare reflector identifier (letters/digits/underscore only, e.g.
 *     "3GPP_TSG_SA4_FS_6G_MED") -- "@list.etsi.org" is appended.
 *
 * Anything else (blank, containing "@" but not a "...@list.etsi.org"
 * address, containing whitespace/punctuation outside [A-Za-z0-9_], or
 * containing CR/LF) is rejected as unsuitable for safe derivation --
 * never guessed at, never silently accepted. Never read from
 * client-supplied request data (a `to`/`recipient`/`mailingList` field on
 * a selection or globalText payload is simply never looked at).
 */
function deriveEmailExportRecipientFromMailingList_(mailingList) {
  const raw = String(mailingList || '').trim();
  if (!raw) {
    return { valid: false, recipient: null, error: 'Mailing List is not configured.' };
  }
  if (/[\r\n]/.test(raw)) {
    return { valid: false, recipient: null, error: 'Mailing List contains invalid characters (CR/LF) and cannot be used to derive a discussion recipient.' };
  }
  const lower = raw.toLowerCase();
  if (/^[a-z0-9_.+-]+@list\.etsi\.org$/.test(lower)) {
    return { valid: true, recipient: lower, error: null };
  }
  if (lower.indexOf('@') !== -1) {
    return { valid: false, recipient: null, error: 'Mailing List "' + raw + '" looks like an e-mail address but is not a recognized "...@list.etsi.org" reflector address; cannot safely derive a discussion recipient.' };
  }
  if (!/^[a-z0-9_]+$/.test(lower)) {
    return { valid: false, recipient: null, error: 'Mailing List "' + raw + '" is not a recognized reflector identifier (expected letters/digits/underscore, e.g. "3GPP_TSG_SA4_FS_6G_MED"); cannot safely derive a discussion recipient.' };
  }
  return { valid: true, recipient: lower + '@list.etsi.org', error: null };
}

/**
 * LEGACY-UPGRADE-006E (kept; used by LEGACY-UPGRADE-006H for BOTH the
 * sender and the recipient): true only for a plausible, single,
 * injection-safe e-mail address -- a simple local-part@domain syntax
 * check, with an explicit rejection of any CR or LF. Used as a final,
 * defense-in-depth sanity gate on the resolved sender
 * (resolveEmailExportSenderAddress_()'s output) and separately on the
 * derived recipient (deriveEmailExportRecipientFromMailingList_()'s
 * output), immediately before each is used -- exactly like the existing
 * double-check pattern isSafeEmailExportUrl_() already uses for the
 * revision-upload URL. Generic on purpose: it validates "is this address
 * safe for a MIME header", not which header it goes into.
 */
function isValidEmailExportSenderAddress_(address) {
  const a = String(address || '').trim();
  if (!a) return false;
  if (/[\r\n]/.test(a)) return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a);
}

/**
 * LEGACY-UPGRADE-006C: the fixed, exporter-controlled "where to upload a
 * revision" sentence -- conceptually distinct from, and placed AFTER, the
 * copied TDoc table's own (unmodified) "Revisions" row, which only
 * reflects revisions currently known/already submitted. `deadlineText` is
 * the SAME already-validated/formatted string used for the Subject and
 * the "Please provide your comments by..." sentence (formatEmailExportDeadline_()),
 * so all three can never disagree. `revisionUrl` is expected to already
 * be resolved server-side (resolveEmailExportRevisionUploadUrl_()) and
 * validated (isSafeEmailExportUrl_()) by the caller; this function
 * re-validates it anyway before rendering an href, as a last defensive
 * layer, and renders nothing (returns '') if either input is missing/
 * unsafe -- it never emits a misleading instruction with no real
 * destination. The visible link text is a short, human-readable label
 * ("Revision upload folder"), not the raw (often very long) FTP URL.
 */
function buildEmailExportRevisionUploadSentenceHtml_(deadlineText, revisionUrl) {
  const dl = String(deadlineText || '').trim();
  const url = String(revisionUrl || '').trim();
  if (!dl || !url || !isSafeEmailExportUrl_(url)) return '';
  return '<p>Please upload revisions by ' + escapeHtmlForEmailExport_(dl) + ' to: ' +
    '<a href="' + escapeHtmlForEmailExport_(url) + '">Revision upload folder</a></p>';
}

/**
 * LEGACY-UPGRADE-006: standard HTML escaping for text pulled out of the
 * report document before it is placed into the generated e-mail's HTML.
 * A separate, fuller escaper than configureMeetingSettings()'s own local
 * esc() (which only escapes quotes, for an HTML *attribute* value) --
 * this one also escapes "<"/">"/"&" for safe HTML *body* content.
 */
function escapeHtmlForEmailExport_(text) {
  return String(text === null || text === undefined ? '' : text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * LEGACY-UPGRADE-006: filenames must never contain characters Windows/
 * Drive forbid or that would make the file hard to use; collapses
 * whitespace and strips anything outside a conservative safe set.
 */
function sanitizeEmailExportFilename_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '') || 'export';
}

/**
 * LEGACY-UPGRADE-006 (canonical subject + agenda-item follow-up), extended
 * by LEGACY-UPGRADE-006B (Part E, deadline) and tightened by
 * LEGACY-UPGRADE-006F (compact, no-space bracket): the canonical FS_6G_MED
 * discussion subject:
 *
 *   [FS_6G_MED,<agendaItem>,<deadlineText>][<tdoc>] Discussion: <title>
 *
 * e.g. "[FS_6G_MED,5.4,26-10-15-1500CEST][S4aP260069] Discussion: ...".
 * `deadlineText` here is the ALREADY-VALIDATED, already-formatted COMPACT
 * "YY-MM-DD-HHmmTZ" subject token (formatEmailExportDeadlineForSubjectToken_()
 * -- deliberately NOT the same string as the body's "YY-MM-DD HH:mm TZ"
 * form, formatEmailExportDeadline_(); both come from the same one
 * validated deadline, just rendered two ways for two different places).
 * This function never validates or formats a deadline itself, only places
 * whatever string it is given.
 *
 * LEGACY-UPGRADE-006F: the first bracket's comma separators have NO
 * surrounding spaces (",", never ", ") -- this applies ONLY to that first
 * bracket; the human-readable part after "Discussion: " (the title) is
 * completely unaffected and keeps its own natural spacing.
 *
 * The TDoc identifier (its OWN bracket group) is the authoritative,
 * machine-readable association key -- ADDON-009: CENTRAL's collector
 * (checkRSSFeed_(), ADDON-008A2: findSA4DocumentIdsInText_() on the
 * reply-prefix-stripped subject) recovers it regardless of a reply/forward
 * prefix and regardless of whether Agenda Item, deadline, or the title are
 * edited in a reply. (Legacy's own extractTdocFromEmailExportSubject_() is
 * therefore not ported -- one association path, not two.)
 * Agenda Item, deadline, and title exist purely for human readability;
 * none is required by any collector to stay unchanged. Agenda Item's
 * existing textual representation (e.g. "5.10") is used exactly as given
 * -- never renumbered or reinterpreted as a number (which would silently
 * turn "5.10" into "5.1"). None of Agenda Item, deadline, or title is ever
 * invented; each is simply omitted, in bracket order, when unavailable
 * (deadlineText is optional here purely for this function's own general
 * reusability/backward compatibility -- the live RPC always supplies one,
 * since every discussion e-mail requires a deadline):
 *
 *   no deadline:                      [FS_6G_MED,<agendaItem>][<tdoc>] Discussion: <title>
 *   no agendaItem:                    [FS_6G_MED,<deadlineText>][<tdoc>] Discussion: <title>
 *   no agendaItem or deadline:        [FS_6G_MED][<tdoc>] Discussion: <title>
 *   no title:                         [FS_6G_MED,<agendaItem>,<deadlineText>][<tdoc>] Discussion
 *   no agendaItem, deadline, or title: [FS_6G_MED][<tdoc>] Discussion
 *
 * ADDON-009: `listTag` (deriveEmailExportListTag_()) replaced Legacy's
 * constant FS_6G_MED as the first field.
 *
 * TEMPLATE-002C (decision 2026-10-01): that first field is gone. It was a
 * report-family / list label ("MBS", "FS_6G_MED"), not the work item, and
 * it is redundant next to the agenda item. Nothing replaces it -- no work
 * item code either. Each remaining part has its own bracket, in the old
 * order:
 *
 *   old   [<tag>,<agendaItem>,<deadlineText>][<tdoc>] Discussion: <title>
 *   new   [<agendaItem>][<deadlineText>][<tdoc>] Discussion: <title>
 *
 * e.g. "[2.5][26-10-15-1500CEST][S4aI260082] Discussion: ...". The
 * deadline token, the TDoc bracket and "Discussion: <title>" are exactly
 * as before; an unavailable agenda item or deadline is still simply
 * omitted:
 *
 *   no deadline:                       [<agendaItem>][<tdoc>] Discussion: <title>
 *   no agendaItem:                     [<deadlineText>][<tdoc>] Discussion: <title>
 *   no agendaItem or deadline:         [<tdoc>] Discussion: <title>
 *   no title:                          [<agendaItem>][<deadlineText>][<tdoc>] Discussion
 *
 * The collector is unaffected: it finds the TDoc identifier anywhere in
 * the subject (findSA4DocumentIdsInText_()).
 */
function buildEmailExportSubject_(tdocNumber, title, agendaItem, deadlineText) {
  const t = String(title || '').trim();
  const a = String(agendaItem || '').trim();
  const dl = String(deadlineText || '').trim();
  const prefix = (a ? '[' + a + ']' : '') + (dl ? '[' + dl + ']' : '');
  const base = `${prefix}[${tdocNumber}] Discussion`;
  return t ? `${base}: ${t}` : base;
}

/**
 * LEGACY-UPGRADE-006B (bulk text controls): converts a plain-text value --
 * as typed into the export dialog's Introduction/Discussion-request text
 * areas -- into safe HTML paragraphs. EVERY character is escaped first via
 * escapeHtmlForEmailExport_() (the same escaper already used everywhere
 * else in this file; never a second, hand-rolled escaper), so arbitrary
 * user text (including "<script>", "&", quotes, etc.) can never inject
 * markup. Structure is then reintroduced on the ESCAPED text only: a blank
 * line (one or more fully-blank lines) starts a new <p>; a single line
 * break within a paragraph becomes <br>. Leading/trailing blank paragraphs
 * are dropped. Blank/whitespace-only input returns '' (never an empty
 * <p></p> pair).
 */
function plainTextToSafeHtmlParagraphs_(text) {
  const raw = String(text === null || text === undefined ? '' : text).replace(/\r\n/g, '\n');
  const paragraphs = raw.split(/\n{2,}/)
    .map(function (p) { return p.trim(); })
    .filter(function (p) { return p.length > 0; });
  if (paragraphs.length === 0) return '';
  return paragraphs.map(function (p) {
    return '<p>' + escapeHtmlForEmailExport_(p).replace(/\n/g, '<br>') + '</p>';
  }).join('');
}

/**
 * LEGACY-UPGRADE-006: assembles the full HTML body -- intro, the COMPLETE
 * copied TDoc table (tableHtml, produced by docTableToHtml_() from the
 * REAL report table, never regenerated here), then the Discussion section
 * and closing. A minimal inline style keeps the table readable in Outlook
 * (which strips most <style> blocks) without depending on external CSS.
 *
 * LEGACY-UPGRADE-006B (bulk text controls): `introHtml`/`discussionHtml`
 * are OPTIONAL pre-converted HTML overrides (already run through
 * plainTextToSafeHtmlParagraphs_() by the caller) for the Introduction and
 * Discussion-request sections; when omitted (null/undefined/''), falls
 * back to this file's own default wording (buildEmailExportDefaultIntroText_()/
 * EMAIL_EXPORT_DEFAULT_DISCUSSION_TEXT_, converted the same way) so every
 * existing caller that predates this stage keeps producing the exact same
 * output as before.
 *
 * LEGACY-UPGRADE-006B (Part F, deadline in body): `deadlineText`, when
 * supplied, is the SAME already-validated/formatted "YY-MM-DD HH:mm TZ"
 * string passed to buildEmailExportSubject_() -- this is what guarantees
 * Subject and body can never disagree about a TDoc's deadline (one value,
 * two call sites, never re-derived). Renders as a fixed, exporter-
 * controlled sentence ("Please provide your comments by <deadlineText>.")
 * placed after the (now shorter) editable discussion text and before the
 * ALSO fixed EMAIL_EXPORT_FORMAL_DECISIONS_NOTE_HTML_ sentence -- neither
 * is part of `discussionHtml`, so editing the global Discussion-request
 * text can never remove either. Omitted, the sentence is simply skipped
 * (kept optional for this function's own general reusability/backward
 * compatibility; the live RPC always supplies one).
 *
 * LEGACY-UPGRADE-006C (revision-upload instruction): `revisionUploadUrl`,
 * when supplied alongside `deadlineText`, adds a second fixed sentence
 * right after the deadline sentence (buildEmailExportRevisionUploadSentenceHtml_()) --
 * also exporter-controlled, also never part of `discussionHtml`. Omitted,
 * or not a safe http(s) URL, the sentence is simply skipped (kept optional
 * here purely for reusability; generateTdocDiscussionEmails() always
 * resolves and validates one before calling this, refusing to generate
 * anything rather than silently omitting the instruction).
 *
 * ADDON-009: `listTag` only fills the default introduction's topic
 * (buildEmailExportDefaultIntroText_()) when no `introHtml` is supplied.
 */
function buildEmailExportHtmlBody_(tableHtml, introHtml, discussionHtml, deadlineText, revisionUploadUrl, listTag) {
  const intro = (introHtml || plainTextToSafeHtmlParagraphs_(buildEmailExportDefaultIntroText_(listTag)));
  const discussion = (discussionHtml || plainTextToSafeHtmlParagraphs_(EMAIL_EXPORT_DEFAULT_DISCUSSION_TEXT_));
  const dl = String(deadlineText || '').trim();
  const deadlineSentenceHtml = dl ? '<p>Please provide your comments by ' + escapeHtmlForEmailExport_(dl) + '.</p>' : '';
  const revisionUploadSentenceHtml = buildEmailExportRevisionUploadSentenceHtml_(deadlineText, revisionUploadUrl);
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#000;">' +
    intro +
    (tableHtml || '') +
    '<p><b>' + escapeHtmlForEmailExport_(EMAIL_EXPORT_DISCUSSION_HEADING_HTML_) + '</b></p>' +
    discussion +
    deadlineSentenceHtml +
    revisionUploadSentenceHtml +
    '<p>' + EMAIL_EXPORT_FORMAL_DECISIONS_NOTE_HTML_ + '</p>' +
    '<p>' + EMAIL_EXPORT_CLOSING_HTML_ + '</p>' +
    '</div>';
}

// ---- Google Docs (Table/Cell/Paragraph/ListItem/Text) -> HTML ----

/**
 * 2.15.2: the report document's own editor URL, the base for links to its
 * tabs/headings/bookmarks. Built from the document ID (Document.getUrl()
 * returns the "open?id=" form, which does not keep a tab/heading link).
 * '' when there is no document ID.
 */
function buildEmailExportDocumentUrl_(documentId) {
  const id = String(documentId || '').trim();
  return /^[A-Za-z0-9_-]+$/.test(id) ? 'https://docs.google.com/document/d/' + id + '/edit' : '';
}

/**
 * 2.15.2: the href a copied Google Docs link gets in the e-mail. Google Docs
 * stores a link to a place in the same document (a tab, heading or
 * bookmark) relative to the document -- "?tab=t.x", "#heading=h.x",
 * "?tab=t.x#heading=h.x" -- which means nothing outside the editor. Such a
 * link is resolved against `documentUrl` (which has no query or fragment of
 * its own, so the query/fragment is simply appended). Anything with a
 * scheme (https:, http:, mailto:, ...) is returned unchanged, and so is any
 * other value (never guessed into a Google Docs URL), as is a relative link
 * when no document URL is known.
 */
function resolveEmailExportLinkUrl_(link, documentUrl) {
  const url = String(link === null || link === undefined ? '' : link);
  const base = String(documentUrl || '');
  if (!base || !/^[?#][^\s"<>]*$/.test(url)) return url;
  return base + url;
}

/**
 * LEGACY-UPGRADE-006: converts one paragraph-or-list-item's TEXT into an
 * HTML fragment, preserving bold/italic/underline/hyperlinks as nested
 * inline tags, segmented at each real formatting-change offset
 * (Text.getTextAttributeIndices()) -- never a naive whole-paragraph
 * bold/italic guess. Line breaks WITHIN a single paragraph's own text
 * (rare here; every writer in this file inserts a new paragraph per line
 * instead, see this section's header comment) are preserved as <br>.
 */
function richTextToHtml_(text, documentUrl) {
  const full = text.getText();
  if (!full) return '';
  let indices = [0];
  try {
    const real = text.getTextAttributeIndices();
    if (Array.isArray(real) && real.length) indices = real;
  } catch (e) { /* fall back to the single [0] segment */ }
  if (indices[0] !== 0) indices = [0].concat(indices);

  let html = '';
  for (let i = 0; i < indices.length; i++) {
    const start = indices[i];
    const end = i + 1 < indices.length ? indices[i + 1] - 1 : full.length - 1;
    if (start > end) continue;
    const segment = full.slice(start, end + 1);
    let piece = escapeHtmlForEmailExport_(segment).replace(/\n/g, '<br>');
    let bold = false, italic = false, underline = false, link = null;
    try { bold = !!text.isBold(start); } catch (e) { }
    try { italic = !!text.isItalic(start); } catch (e) { }
    try { underline = !!text.isUnderline(start); } catch (e) { }
    try { link = text.getLinkUrl(start); } catch (e) { }
    // 2.15.2: a link into this document is made absolute.
    if (link) link = resolveEmailExportLinkUrl_(link, documentUrl);
    if (link) piece = '<a href="' + escapeHtmlForEmailExport_(link) + '">' + piece + '</a>';
    if (underline && !link) piece = '<u>' + piece + '</u>'; // a link is already visually distinct
    if (italic) piece = '<i>' + piece + '</i>';
    if (bold) piece = '<b>' + piece + '</b>';
    html += piece;
  }
  return html;
}

/**
 * LEGACY-UPGRADE-006: converts one report TableCell's full content
 * (every paragraph/list item it contains, in document order) into an HTML
 * fragment. Consecutive list items are grouped into one <ul>/<ol>
 * (BULLET-family glyphs -> <ul>, everything else -> <ol>); every other
 * paragraph becomes its own <p>. An empty cell becomes an empty string
 * (renders as a normal empty table cell, not a stray tag).
 */
/**
 * LEGACY-UPGRADE-006 (Outlook-ready follow-up): true for a glyph type that
 * reads as an ordered/numbered list ("NUMBER", "LATIN_UPPER"/"LATIN_LOWER",
 * "ROMAN_UPPER"/"ROMAN_LOWER" -- Apps Script's real GlyphType values other
 * than "BULLET"); everything else (including an unavailable glyph) is
 * treated as unordered. Used per list-item, not per document, since real
 * Google Docs list items each carry their own glyph.
 */
function isOrderedListGlyph_(glyph) {
  return /NUMBER|LATIN|ROMAN/.test(String(glyph || ''));
}

/**
 * LEGACY-UPGRADE-006 (Outlook-ready follow-up): converts one report
 * TableCell's full content into HTML, preserving the Google Doc's real
 * LIST NESTING rather than flattening every list item into one level.
 *
 * Uses ListItem.getNestingLevel() (0-based depth) and ListItem.getListId()
 * -- both real DocumentApp.ListItem methods -- to reconstruct proper
 * nested <ul>/<ol> structure: a stack of open list levels is maintained;
 * an item deeper than the current stack opens new nested list(s); an item
 * shallower closes back down to that level; a paragraph or a level-0 item
 * belonging to a DIFFERENT listId than the currently open level-0 list
 * closes everything first (paragraphs always interrupt a list in HTML;
 * two lists back-to-back with no paragraph between them but different
 * source lists are kept as two separate, independent <ul>/<ol> elements,
 * never silently merged). Ordered-vs-unordered is decided per item from
 * its own glyph type (isOrderedListGlyph_()), matching how Google Docs
 * itself assigns glyphs. Inline formatting (bold/italic/underline/links,
 * via richTextToHtml_()) works identically inside a nested item as at the
 * top level. Missing nesting-level/list-ID support degrades gracefully to
 * flat level-0 behavior (the previous, pre-nesting output), never throws.
 */
function docCellToHtml_(cell, documentUrl) {
  const PARAGRAPH = DocumentApp.ElementType.PARAGRAPH;
  const LIST_ITEM = DocumentApp.ElementType.LIST_ITEM;
  const n = cell.getNumChildren();
  let html = '';
  const stack = []; // one open frame per nesting level: { level, tag, listId }; stack[top] has an open, not-yet-closed <li>

  function closeFrame() {
    const frame = stack.pop();
    html += '</li></' + frame.tag + '>';
  }
  function closeDeeperThan(level) {
    while (stack.length > 0 && stack[stack.length - 1].level > level) closeFrame();
  }
  function closeAllLists() {
    while (stack.length > 0) closeFrame();
  }

  // Places one list item at `level`. Closes anything deeper first. If a
  // frame is already open at exactly this level, either continues it
  // (same tag/listId: close the previous <li>, open a new one in the SAME
  // wrapper) or -- a real, different list resuming at this level, with no
  // intervening paragraph -- closes that wrapper with its OWN original
  // tag (never the new item's tag, which would produce a mismatched
  // closing tag) and opens a genuinely new one. If no frame is open at
  // this level yet (first item, or a level was skipped -- Google Docs
  // nesting always increases one level at a time in practice, so a
  // larger jump is simply opened as a single new level rather than
  // guessing at intermediate ones), opens it fresh.
  function placeListItem(level, tag, listId, itemHtml) {
    closeDeeperThan(level);
    const top = stack.length ? stack[stack.length - 1] : null;
    if (top && top.level === level) {
      if (top.tag !== tag || (listId !== null && top.listId !== null && listId !== top.listId)) {
        closeFrame();
        html += '<' + tag + '><li>' + itemHtml;
        stack.push({ level: level, tag: tag, listId: listId });
      } else {
        html += '</li><li>' + itemHtml;
        top.listId = listId;
      }
    } else {
      html += '<' + tag + '><li>' + itemHtml;
      stack.push({ level: level, tag: tag, listId: listId });
    }
  }

  for (let i = 0; i < n; i++) {
    const child = cell.getChild(i);
    const type = child.getType();
    if (type === LIST_ITEM) {
      const li = child.asListItem ? child.asListItem() : child;
      const level = (typeof li.getNestingLevel === 'function') ? (li.getNestingLevel() || 0) : 0;
      const listId = (typeof li.getListId === 'function') ? li.getListId() : null;
      const tag = isOrderedListGlyph_(li.getGlyphType && li.getGlyphType()) ? 'ol' : 'ul';
      placeListItem(level, tag, listId, richTextToHtml_(li.editAsText(), documentUrl));
    } else if (type === PARAGRAPH) {
      closeAllLists();
      const p = child.asParagraph ? child.asParagraph() : child;
      const inner = richTextToHtml_(p.editAsText(), documentUrl);
      html += '<p style="margin:0 0 4px 0;">' + (inner || '&nbsp;') + '</p>';
    }
    // Any other child type (e.g. a nested table -- never produced by this
    // codebase's TDoc-table builders) is intentionally skipped rather than
    // guessed at.
  }
  closeAllLists();
  return html;
}

/**
 * LEGACY-UPGRADE-006: converts the COMPLETE report TDoc table into one
 * self-contained HTML <table> -- every row, every cell, in document
 * order. Basic inline borders/padding only (Outlook strips most CSS
 * classes/embedded <style> blocks, so every style that matters is inline).
 * This is a literal copy of the existing table's content; nothing is
 * reordered, summarized, or regenerated.
 */
function docTableToHtml_(table, documentUrl) {
  const rows = table.getNumRows();
  let html = '<table style="border-collapse:collapse;width:100%;font-family:Arial,Helvetica,sans-serif;font-size:12px;" border="1" cellpadding="6" cellspacing="0">';
  for (let r = 0; r < rows; r++) {
    const row = table.getRow(r);
    const cells = row.getNumCells();
    html += '<tr>';
    for (let c = 0; c < cells; c++) {
      const cell = row.getCell(c);
      const isLabelCol = c === 0;
      const cellStyle = 'border:1px solid #999;padding:6px;vertical-align:top;text-align:left;' + (isLabelCol ? 'font-weight:bold;background:#f2f2f2;white-space:nowrap;' : '');
      html += '<td style="' + cellStyle + '">' + docCellToHtml_(cell, documentUrl) + '</td>';
    }
    html += '</tr>';
  }
  html += '</table>';
  return html;
}

// ---- TDoc-table detection (read-only) ----

/**
 * LEGACY-UPGRADE-006: scans the CURRENT document body (read-only -- no
 * write of any kind) for TDoc tables (isTDocTable_(), the SAME check the
 * build/update paths already use) and returns one summary entry per table
 * for the dialog list. Table order in the returned array matches document
 * order; `tableIndex` is that table's position among getTables() results,
 * used later to re-locate the exact same table for export without ever
 * relying on TDoc-number uniqueness (defensive, though TDoc numbers are
 * expected unique in practice).
 */
/**
 * LEGACY-UPGRADE-006 (detection-fix follow-up): isTDocTable_() only checks
 * that row 0 / column 0 reads "TDoc" -- true for a genuine TDoc detail
 * table, but ALSO true for two other tables this same file creates with
 * "TDoc" as their own header label: the "Registered Documents" summary
 * table (createSummaryTable_(): header row ['TDoc', 'Title', 'Source',
 * 'Agenda Item']) and the "Document Reallocations" table
 * (isReallocationTable_()'s own check: header row ['TDoc',
 * 'Original Agenda', 'New Agenda', 'Reason']). Both pass isTDocTable_()'s
 * check, and their row-0/column-1 header text ("Title"/"Original Agenda")
 * was being read as if it were a TDoc number -- the exact live bug
 * observed. isTDocTable_() itself is deliberately left untouched (it is
 * used by the actual build/update paths -- this task does not touch
 * report generation); this is an ADDITIONAL, export-only validation layer
 * on top of it, reusing the SAME canonical whole-string SA4 identifier
 * check every other consumer in this file uses (parseExactSA4DocumentId_()/
 * SA4_TDOC_FAMILIES) -- never a second, hand-rolled regex, and never
 * limited to one series: every family that registry recognizes (S4-,
 * S4aA, S4aP, S4aV, S4aI, A4aR) is accepted here exactly as it already is
 * everywhere else in this codebase.
 */
function isPlausibleTdocIdentifier_(value) {
  return parseExactSA4DocumentId_(value).isValid;
}

function detectTdocTablesInDocument_(body) {
  const tables = body.getTables();
  const found = [];
  const countByTdoc = {};
  tables.forEach(function (t, idx) {
    if (!isTDocTable_(t)) return;
    // LEGACY-UPGRADE-006 (detection-fix follow-up): reject anything that
    // is not an actual SA4 TDoc identifier (see isPlausibleTdocIdentifier_()
    // above) -- this is what excludes the summary/reallocation tables'
    // own header rows, blank values, and arbitrary prose.
    // ADDON-009: the identifier is used in its canonical spelling.
    const parsedTdoc = parseExactSA4DocumentId_(safeCellText_(t, 0, 1));
    if (!parsedTdoc.isValid) return;
    const tdoc = parsedTdoc.raw;
    countByTdoc[tdoc] = (countByTdoc[tdoc] || 0) + 1;
    // LEGACY-UPGRADE-006 (status-filter follow-up): `status` is the
    // table's own UNMODIFIED Status cell -- read only, never written back.
    // `excluded` is derived purely for the dialog's selectable-list
    // filtering; the original status value is always preserved alongside
    // it for display/data purposes.
    //
    // LEGACY-UPGRADE-006B (reserved-status hard exclusion): `excluded` now
    // covers BOTH hard-exclusion reasons (Approved/Agreed and reserved) --
    // in both cases the TDoc must never be selectable or exportable, so
    // one boolean still correctly answers "can this be offered at all".
    // `hardExclusionReason` distinguishes WHICH rule fired, purely so the
    // dialog can report separate counts ("excluded by status
    // (Approved/Agreed)" vs. "reserved/unavailable") without guessing --
    // never a third catch-all reason, since no other hard-exclusion rule
    // exists in this codebase (see isEmailExportReserved_()'s own header
    // comment for the codebase-wide check that confirmed this).
    const status = findCellText_(t, 'Status').trim();
    const approvedOrAgreed = isEmailExportStatusExcluded_(status);
    const reserved = !approvedOrAgreed && isEmailExportReserved_(status);
    found.push({
      tableIndex: idx,
      tdoc: tdoc,
      title: findCellText_(t, 'Title').trim(),
      source: findCellText_(t, 'Source').trim(),
      agendaItem: findCellText_(t, 'Agenda Item').trim(),
      status: status,
      excluded: approvedOrAgreed || reserved,
      hardExclusionReason: reserved ? 'reserved' : (approvedOrAgreed ? 'approved-agreed' : null)
    });
  });
  // LEGACY-UPGRADE-006 (detection-fix follow-up): a genuinely valid TDoc
  // number appearing on MORE THAN ONE physical table is flagged, never
  // silently deduplicated or merged -- the tables may hold materially
  // different content (e.g. one left stale by an update that only found/
  // touched the other), and only a human can safely judge which is
  // current. Every occurrence remains its own separate, independently
  // selectable row.
  found.forEach(function (f) {
    f.duplicateCount = countByTdoc[f.tdoc];
  });
  return found;
}

// ---- quoted-printable + MIME (.eml) ----

/**
 * LEGACY-UPGRADE-006: quoted-printable encodes a UTF-8 string per RFC
 * 2045 -- every byte outside printable US-ASCII (and outside "=") is
 * escaped as "=XX", long lines are soft-wrapped at 76 characters with a
 * trailing "=" continuation, and CRLF is used throughout (required for
 * MIME/email; Apps Script string content is otherwise "\n"-only).
 */
function quotedPrintableEncode_(str) {
  const bytes = Utilities.newBlob(String(str || '')).getBytes();
  let out = '';
  let lineLen = 0;
  function put(chunk) {
    if (lineLen + chunk.length > 75) { out += '=\r\n'; lineLen = 0; }
    out += chunk;
    lineLen += chunk.length;
  }
  for (let i = 0; i < bytes.length; i++) {
    let b = bytes[i];
    if (b < 0) b += 256; // Apps Script bytes are signed
    const ch = String.fromCharCode(b);
    if (ch === '\r') continue; // normalize: CRLF is emitted explicitly below
    if (ch === '\n') { out += '\r\n'; lineLen = 0; continue; }
    if ((b >= 33 && b <= 126 && b !== 61) || b === 32 || b === 9) {
      put(ch);
    } else {
      put('=' + ('0' + b.toString(16).toUpperCase()).slice(-2));
    }
  }
  return out;
}

/**
 * LEGACY-UPGRADE-006: RFC 2047 encoded-word for a header value that may
 * contain non-ASCII characters (a title with an accented name, an en/em
 * dash, ...) -- "=?UTF-8?B?<base64>?=". Plain ASCII values are returned
 * unchanged (no unnecessary encoding).
 */
function encodeMimeHeaderValue_(value) {
  const v = String(value || '');
  if (/^[\x20-\x7E]*$/.test(v)) return v;
  const b64 = Utilities.base64Encode(Utilities.newBlob(v).getBytes());
  return '=?UTF-8?B?' + b64 + '?=';
}

/**
 * LEGACY-UPGRADE-006: builds a complete, standards-compatible .eml
 * (RFC 5322 message / MIME) text body: Subject/To/From (From/Date omitted
 * entirely when not available or not configured -- never a fabricated
 * address), Content-Type text/html; charset=UTF-8, quoted-printable
 * transfer encoding, CRLF line endings throughout.
 */
function buildEmlContent_(headers, htmlBody) {
  const h = headers || {};
  const lines = [];
  lines.push('MIME-Version: 1.0');
  // LEGACY-UPGRADE-006 (Outlook-ready follow-up): tells desktop Outlook to
  // open this .eml as an unsent/editable message (body editable, a real
  // Send button) rather than as a received message with only Reply/
  // Reply All/Forward. This never sends anything by itself -- the user
  // still edits and presses Send manually.
  lines.push('X-Unsent: 1');
  if (h.from) lines.push('From: ' + h.from);
  if (h.to) lines.push('To: ' + h.to);
  // TEMPLATE-002C: replies go to the meeting's mailing list, not to the
  // sender. Written only when the caller supplies it -- never derived from
  // From here.
  if (h.replyTo) lines.push('Reply-To: ' + h.replyTo);
  lines.push('Subject: ' + encodeMimeHeaderValue_(h.subject || ''));
  if (h.date) lines.push('Date: ' + h.date);
  lines.push('Content-Type: text/html; charset=UTF-8');
  lines.push('Content-Transfer-Encoding: quoted-printable');
  lines.push('');
  lines.push(quotedPrintableEncode_(htmlBody || ''));
  return lines.join('\r\n');
}

// ---- orchestration (still read-only against the Doc) ----

/**
 * LEGACY-UPGRADE-006: builds one export {fileName, eml} for a single
 * already-located TDoc table. `subjectOverride` is an optional parameter
 * kept for this function's own general reusability/testability -- the
 * live dialog's RPC (generateTdocDiscussionEmails()) never supplies one,
 * since the canonical Subject is part of the machine-readable TDoc-
 * association contract and is not user-editable (see that function's own
 * header comment). `mailingList` is the already-resolved (Document
 * Property) value, never invented here.
 *
 * LEGACY-UPGRADE-006B (bulk text controls): `introText`/`discussionText`
 * are OPTIONAL plain-text overrides for the Introduction/Discussion-
 * request sections (converted to safe HTML here via
 * plainTextToSafeHtmlParagraphs_() -- the caller never pre-converts them);
 * omitted, buildEmailExportHtmlBody_() falls back to this file's own
 * default wording, so every pre-existing call site is unaffected.
 *
 * LEGACY-UPGRADE-006B (Part E/F, deadline): `deadlineText`, when supplied,
 * is the SAME already-validated/formatted string placed into both the
 * canonical Subject (buildEmailExportSubject_()) and the body's deadline
 * sentence (buildEmailExportHtmlBody_()) -- one value, both call sites,
 * so the two can never disagree. Optional purely for this function's own
 * reusability/backward compatibility.
 *
 * LEGACY-UPGRADE-006C: `revisionUploadUrl`, when supplied, is passed
 * straight through to buildEmailExportHtmlBody_() for the revision-upload
 * sentence -- resolved/validated server-side by the caller, never taken
 * from client input.
 *
 * LEGACY-UPGRADE-006H: `senderAddress`, when supplied, is the From header
 * value -- resolved/validated server-side by the caller
 * (resolveEmailExportSenderAddress_()/isValidEmailExportSenderAddress_())
 * from the explicitly configured DISCUSSION_EMAIL_SENDER, NEVER from
 * Mailing List, NEVER from Session.getActiveUser()/getEffectiveUser() (the
 * Google login identity running this script), and never from
 * client-supplied request data. `mailingList` above (this function's own
 * pre-existing parameter) carries the SEPARATE derived recipient/To
 * address -- the two are deliberately different values; see
 * generateTdocDiscussionEmails()'s own header comment for why. If
 * `senderAddress` is omitted, no From header is emitted (matching this
 * file's existing "no fabricated header" convention for a missing value)
 * -- kept optional purely for this function's own reusability/backward
 * compatibility.
 *
 * LEGACY-UPGRADE-006F (compact subject deadline token): `subjectDeadlineToken`,
 * when supplied, is the compact "YY-MM-DD-HHmmTZ" form
 * (formatEmailExportDeadlineForSubjectToken_()) used ONLY in the Subject;
 * `deadlineText` (the human "YY-MM-DD HH:mm TZ" form) continues to drive
 * the body's deadline sentences, unchanged. Omitted, falls back to
 * `deadlineText` itself (this function's own general reusability/backward
 * compatibility) -- the live RPC always supplies both, derived from the
 * SAME validated deadline.
 *
 * ADDON-009: `listTag` (required, deriveEmailExportListTag_()) replaces
 * Legacy's constant FS_6G_MED in the subject, the default introduction and
 * the file name ("<listTag>_<tdoc>.eml").
 *
 * 2.15.2: `documentUrl` (buildEmailExportDocumentUrl_()) makes links into
 * the report document absolute (resolveEmailExportLinkUrl_()).
 */
function buildEmailExportForTdocTable_(table, meta, mailingList, subjectOverride, introText, discussionText, deadlineText, revisionUploadUrl, senderAddress, subjectDeadlineToken, listTag, documentUrl, replyToAddress) {
  const subject = (subjectOverride && subjectOverride.trim()) || buildEmailExportSubject_(meta.tdoc, meta.title, meta.agendaItem, subjectDeadlineToken || deadlineText);
  const tableHtml = docTableToHtml_(table, documentUrl);
  const introHtml = introText ? plainTextToSafeHtmlParagraphs_(introText) : null;
  const discussionHtml = discussionText ? plainTextToSafeHtmlParagraphs_(discussionText) : null;
  const htmlBody = buildEmailExportHtmlBody_(tableHtml, introHtml, discussionHtml, deadlineText, revisionUploadUrl, listTag);
  const from = String(senderAddress || '').trim();
  const eml = buildEmlContent_({ to: mailingList || '', from: from, replyTo: String(replyToAddress || '').trim(), subject: subject }, htmlBody);
  const fileName = sanitizeEmailExportFilename_(listTag + '_' + meta.tdoc) + '.eml';
  return { fileName: fileName, eml: eml, subject: subject };
}

/**
 * LEGACY-UPGRADE-006: this stage's internal data model is a LIST OF
 * GROUPS, where a group is one or more TDoc numbers that should share a
 * single generated e-mail -- deliberately future-proofed for "multiple
 * TDocs -> one combined discussion thread" (per this stage's own scope,
 * the UI only ever sends single-TDoc groups today; a later stage can add
 * a grouping UI without changing this function's contract). A group with
 * more than one TDoc concatenates each table's HTML in sequence, in the
 * order given, under one shared subject/intro/discussion/closing.
 * ADDON-009: `listTag` as for buildEmailExportForTdocTable_().
 */
function buildEmailExportForGroup_(tables, metas, mailingList, subjectOverride, introText, discussionText, deadlineText, revisionUploadUrl, senderAddress, subjectDeadlineToken, listTag, documentUrl, replyToAddress) {
  const primary = metas[0];
  const subject = (subjectOverride && subjectOverride.trim()) || buildEmailExportSubject_(primary.tdoc, primary.title, primary.agendaItem, subjectDeadlineToken || deadlineText);
  const tableHtml = tables.map(function (t) { return docTableToHtml_(t, documentUrl); }).join('<br>');
  const introHtml = introText ? plainTextToSafeHtmlParagraphs_(introText) : null;
  const discussionHtml = discussionText ? plainTextToSafeHtmlParagraphs_(discussionText) : null;
  const htmlBody = buildEmailExportHtmlBody_(tableHtml, introHtml, discussionHtml, deadlineText, revisionUploadUrl, listTag);
  const from = String(senderAddress || '').trim();
  const eml = buildEmlContent_({ to: mailingList || '', from: from, replyTo: String(replyToAddress || '').trim(), subject: subject }, htmlBody);
  const fileNameBase = metas.length === 1 ? metas[0].tdoc : metas.map(function (m) { return m.tdoc; }).join('_');
  const fileName = sanitizeEmailExportFilename_(listTag + '_' + fileNameBase) + '.eml';
  return { fileName: fileName, eml: eml, subject: subject };
}

/**
 * LEGACY-UPGRADE-006: finds (or creates, exactly once) the shared export
 * folder by name. Never overwrites or otherwise touches any unrelated
 * file; if the folder already exists, reuses it as-is.
 */
function ensureEmailExportDriveFolder_() {
  const existing = DriveApp.getFoldersByName(EMAIL_EXPORT_DRIVE_FOLDER_NAME_);
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder(EMAIL_EXPORT_DRIVE_FOLDER_NAME_);
}

/**
 * LEGACY-UPGRADE-006D: deterministic, filesystem-safe ZIP filename for one
 * export batch -- "FS_6G_MED_Discussion_Emails_YYYY-MM-DD_HHmm.zip", built
 * from the EXPORT-GENERATION timestamp (`now`), never the discussion
 * deadline (a completely different value with a completely different
 * format). Uses the script's own time zone (Session.getScriptTimeZone()),
 * never the caller's browser locale, so the name is never locale-
 * dependent. Reuses sanitizeEmailExportFilename_() -- the SAME helper
 * every individual .eml filename already goes through -- rather than a
 * second, parallel sanitizer.
 * ADDON-009: "<listTag>_Discussion_Emails_YYYY-MM-DD_HHmm.zip".
 */
function buildEmailExportZipFileName_(now, listTag) {
  const tz = Session.getScriptTimeZone();
  const stamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd_HHmm');
  return sanitizeEmailExportFilename_(listTag + '_Discussion_Emails_' + stamp) + '.zip';
}

/**
 * LEGACY-UPGRADE-006: menu entry point -- "📧 EMAIL EXPORT" -> "Prepare
 * TDoc Discussion E-mails". Read-only: scans the current document for
 * TDoc tables and shows them in a selection dialog. Nothing is written
 * (to the Doc, Drive, or Document Properties) until the dialog's own
 * "Generate" action runs generateTdocDiscussionEmails() below.
 */
function prepareTdocDiscussionEmails() {
  const ui = DocumentApp.getUi();
  const body = DocumentApp.getActiveDocument().getBody();
  const allDetected = detectTdocTablesInDocument_(body);
  // ADDON-009: the dialog needs the list tag for its default introduction.
  // A configuration problem is shown up front; Generate re-checks it and
  // refuses (resolveEmailExportConfiguration_()) before writing anything.
  let exportConfig = null;
  let configError = null;
  try {
    exportConfig = resolveEmailExportConfiguration_();
  } catch (e) {
    configError = e.message;
  }
  const defaultIntroText = buildEmailExportDefaultIntroText_(exportConfig ? exportConfig.listTag : '');
  // LEGACY-UPGRADE-006 (status-filter follow-up), extended by
  // LEGACY-UPGRADE-006B (reserved-status hard exclusion): a TDoc that is
  // already Approved/Agreed, OR reserved (no real content yet), is not
  // offered for a new discussion -- simply omitted from the selectable
  // list. Nothing about the source table/Status is ever changed;
  // `allDetected` (including hard-excluded entries, with their real
  // status) is only used for the summary counts below.
  const tdocs = allDetected.filter(function (t) { return !t.excluded; });
  const approvedAgreedCount = allDetected.filter(function (t) { return t.hardExclusionReason === 'approved-agreed'; }).length;
  const reservedCount = allDetected.filter(function (t) { return t.hardExclusionReason === 'reserved'; }).length;

  function esc(v) { return String(v === null || v === undefined ? '' : v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

  // LEGACY-UPGRADE-006B (status filter): the filter's own checkbox set is
  // built from the ACTUALLY-DETECTED distinct status strings among the
  // remaining (non-hard-excluded) TDocs -- never a hardcoded universal
  // status list. A blank status is grouped under its own "(no status)"
  // label rather than being hidden or merged into another value. This is
  // a pure client-side VISIBILITY filter on top of the already-eligible
  // list -- it can only narrow which of these rows are shown/selectable;
  // it can never bring back an Approved/Agreed or reserved TDoc (those are
  // never even present in `tdocs` to begin with).
  const distinctStatuses = [];
  tdocs.forEach(function (t) {
    const label = t.status || '(no status)';
    if (distinctStatuses.indexOf(label) === -1) distinctStatuses.push(label);
  });
  const statusFilterHtml = distinctStatuses.length <= 1 ? '' :
    '<div class="hint" style="margin-top:10px;">Show status: ' +
    distinctStatuses.map(function (s) {
      return '<label style="margin-right:10px;font-weight:normal;">' +
        '<input type="checkbox" class="statusPick" value="' + esc(s) + '" checked onchange="applyStatusFilter()"> ' + esc(s) + '</label>';
    }).join('') +
    '</div>';

  // LEGACY-UPGRADE-006 correctness fix: data-idx must be the TDoc table's
  // REAL position among body.getTables() (t.tableIndex) -- the earlier
  // implementation used this loop's own index into the (already-filtered)
  // `tdocs` array, which only happened to match when every table in the
  // document was itself a TDoc table (never true once a summary/config
  // table is also present).
  // LEGACY-UPGRADE-006 (detection-fix follow-up): a TDoc number seen on
  // more than one physical table is never silently merged -- each
  // occurrence stays its own row, flagged so the user can judge which one
  // (if either) is current before selecting it.
  // LEGACY-UPGRADE-006 (canonical/read-only subject follow-up): the
  // Subject is no longer an editable per-row field -- it is part of the
  // machine-readable contract that lets a reflector reply be associated
  // back with the right TDoc (ADDON-008A2 collector), so
  // the user must not be able to accidentally change/remove the TDoc
  // identifier, the Agenda Item prefix, or the canonical structure. The
  // canonical Subject itself is generated server-side (see
  // generateTdocDiscussionEmails() below), never taken from this UI.
  // LEGACY-UPGRADE-006B (status column): a Status column is added purely
  // for display/filtering -- the row's `data-status` attribute drives the
  // client-side filter above; the Status value itself is never editable
  // here and is always re-read fresh, server-side, at Generate time.
  //
  // LEGACY-UPGRADE-006B (Part G, per-row deadline): every eligible row also
  // gets its OWN date/time/time-zone inputs, prefilled with this file's own
  // default batch deadline (EMAIL_EXPORT_DEFAULT_DEADLINE_*_; ADDON-009:
  // empty date/time, so every row needs an entered deadline). A hard-excluded (Approved/Agreed/reserved) TDoc has no row at
  // all here, so it structurally can never receive or export a deadline.
  const timezoneOptionsHtml = EMAIL_EXPORT_SUPPORTED_TIMEZONES_.map(function (tz) {
    return '<option value="' + esc(tz) + '"' + (tz === EMAIL_EXPORT_DEFAULT_DEADLINE_TZ_ ? ' selected' : '') + '>' + esc(tz) + '</option>';
  }).join('');
  const rowsHtml = tdocs.length === 0
    ? ''
    : tdocs.map(function (t) {
      const dupNote = t.duplicateCount > 1 ? ' <span style="color:#a94442;">⚠ ' + t.duplicateCount + ' tables detected for this TDoc</span>' : '';
      const statusLabel = t.status || '(no status)';
      return '<tr class="tdocRow" data-status="' + esc(statusLabel) + '">' +
        '<td><input type="checkbox" class="tdocPick" data-idx="' + t.tableIndex + '"></td>' +
        '<td>' + esc(t.tdoc) + dupNote + '</td>' +
        '<td>' + esc(t.title) + '</td>' +
        '<td>' + esc(t.source) + '</td>' +
        '<td>' + esc(t.agendaItem) + '</td>' +
        '<td>' + esc(t.status) + '</td>' +
        '<td style="white-space:nowrap;">' +
          '<input type="date" class="deadlineDate" data-idx="' + t.tableIndex + '" value="' + esc(EMAIL_EXPORT_DEFAULT_DEADLINE_DATE_) + '" style="width:112px;"> ' +
          '<input type="time" class="deadlineTime" data-idx="' + t.tableIndex + '" value="' + esc(EMAIL_EXPORT_DEFAULT_DEADLINE_TIME_) + '" style="width:78px;"> ' +
          '<select class="deadlineTz" data-idx="' + t.tableIndex + '">' + timezoneOptionsHtml + '</select>' +
        '</td>' +
        '</tr>';
    }).join('');

  let summaryHtml;
  if (allDetected.length === 0) {
    summaryHtml = '<p>No TDoc tables were found in this document.</p>';
  } else if (tdocs.length === 0) {
    // LEGACY-UPGRADE-006 (status-filter follow-up) wording preserved
    // verbatim for the pure Approved/Agreed case; LEGACY-UPGRADE-006B
    // (reserved-status hard exclusion) extends it only when a reserved
    // TDoc is ALSO present, never changing the original phrase's meaning.
    summaryHtml = reservedCount > 0 && approvedAgreedCount === 0
      ? '<p>All ' + allDetected.length + ' detected TDoc(s) are reserved/unavailable — nothing is currently available for a new discussion.</p>'
      : reservedCount > 0
        ? '<p>All ' + allDetected.length + ' detected TDoc(s) are already Approved/Agreed or reserved — nothing is currently available for a new discussion.</p>'
        : '<p>All ' + allDetected.length + ' detected TDoc(s) are already Approved/Agreed — nothing is currently available for a new discussion.</p>';
  } else {
    summaryHtml = '<p>' + tdocs.length + ' TDoc' + (tdocs.length === 1 ? '' : 's') + ' available for discussion' +
      (approvedAgreedCount > 0 ? '<br>' + approvedAgreedCount + ' TDoc' + (approvedAgreedCount === 1 ? '' : 's') + ' excluded by status (Approved/Agreed)' : '') +
      (reservedCount > 0 ? '<br>' + reservedCount + ' TDoc' + (reservedCount === 1 ? '' : 's') + ' reserved/unavailable' : '') +
      '</p>';
  }

  const tableHtml = tdocs.length === 0 ? '' :
    '<table>' +
    '<tr><th></th><th>TDoc</th><th>Title</th><th>Source</th><th>Agenda Item</th><th>Status</th><th>Deadline (date / time / TZ)</th></tr>' +
    rowsHtml +
    '</table>';

  // LEGACY-UPGRADE-006B (Part G, batch deadline controls): a single
  // date/time/time-zone set the user can apply, in bulk, to either the
  // currently-selected rows or every currently-visible (not filtered out)
  // eligible row -- never to a hard-excluded TDoc, since those have no row
  // to apply to in the first place. Initialized to the same
  // EMAIL_EXPORT_DEFAULT_DEADLINE_*_ constants as every individual row.
  const batchDeadlineHtml = tdocs.length === 0 ? '' :
    '<div class="hint" style="margin-top:10px;">' +
    '<b>Batch deadline (required):</b> ' +
    'Date <input type="date" id="batchDate" value="' + esc(EMAIL_EXPORT_DEFAULT_DEADLINE_DATE_) + '"> ' +
    'Time <input type="time" id="batchTime" value="' + esc(EMAIL_EXPORT_DEFAULT_DEADLINE_TIME_) + '"> ' +
    'TZ <select id="batchTz">' + timezoneOptionsHtml + '</select> ' +
    '<button type="button" onclick="applyDeadlineToSelected()">Apply to selected</button> ' +
    '<button type="button" onclick="applyDeadlineToVisible()">Apply to all visible/eligible</button>' +
    '</div>';

  // LEGACY-UPGRADE-006B (bulk text controls): these two text areas are
  // GLOBAL -- one Introduction and one Discussion-request value apply to
  // EVERY e-mail generated by this one dialog operation, never per-row.
  // Prefilled with this file's own default wording (single source of
  // truth: buildEmailExportDefaultIntroText_()/EMAIL_EXPORT_DEFAULT_DISCUSSION_TEXT_)
  // so the dialog always reverts to the same defaults on reopen -- nothing
  // here is read from or written to Document Properties or any other
  // storage. Plain text only: the server converts it to safe HTML
  // (plainTextToSafeHtmlParagraphs_(), escaping every character) at
  // Generate time, so arbitrary typed text (including "<"/"&"/quotes)
  // can never inject markup into the generated e-mail. The canonical,
  // machine-readable Subject is a completely separate mechanism and is
  // NOT affected by, and has no input field alongside, this text.
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      table { border-collapse: collapse; width: 100%; margin-top: 10px; }
      th, td { border: 1px solid #ccc; padding: 6px; text-align: left; font-size: 12px; }
      th { background: #f5f5f5; }
      tr.tdocRow.hiddenByFilter { display: none; }
      button { margin-top: 16px; padding: 8px 16px; background: #4285f4; color: white; border: none; cursor: pointer; }
      button:disabled { background: #999; }
      #status { margin-top: 10px; font-size: 12px; }
      #results a { display: block; margin-top: 4px; }
      .hint { font-size: 11px; color: #666; margin-top: 6px; }
      #summary { font-size: 13px; margin-top: 4px; }
      textarea { width: 100%; box-sizing: border-box; font-family: inherit; font-size: 12px; margin-top: 4px; }
      label.fieldLabel { display: block; margin-top: 12px; font-weight: bold; }
    </style>
    <h2>📧 Prepare TDoc Discussion E-mails</h2>
    <div class="hint">Read-only: this only creates new .eml files in Drive ("${EMAIL_EXPORT_DRIVE_FOLDER_NAME_}"). It never modifies this report, never sends any e-mail.</div>
    ${exportConfig
      ? '<div class="hint">From: <b>' + esc(exportConfig.senderAddress) + '</b> &nbsp; To: <b>' + esc(exportConfig.recipientAddress) + '</b></div>'
      : '<p style="color:#a94442;">❌ ' + esc(configError) + '</p>'}
    <label class="fieldLabel" for="introText">Introduction (applies to all selected e-mails)</label>
    <textarea id="introText" rows="5">${esc(defaultIntroText)}</textarea>
    <label class="fieldLabel" for="discussionText">Discussion request (applies to all selected e-mails)</label>
    <textarea id="discussionText" rows="5">${esc(EMAIL_EXPORT_DEFAULT_DISCUSSION_TEXT_)}</textarea>
    <div id="summary">${summaryHtml}</div>
    ${statusFilterHtml}
    ${batchDeadlineHtml}
    ${tableHtml}
    <button id="genBtn" onclick="generate()" ${tdocs.length === 0 ? 'disabled' : ''}>Generate selected (one e-mail per TDoc)</button>
    <div id="status"></div>
    <div id="results"></div>
    <script>
      function applyStatusFilter() {
        const checked = Array.prototype.slice.call(document.querySelectorAll('.statusPick:checked')).map(function (b) { return b.value; });
        Array.prototype.slice.call(document.querySelectorAll('.tdocRow')).forEach(function (row) {
          const visible = checked.indexOf(row.getAttribute('data-status')) !== -1;
          row.classList.toggle('hiddenByFilter', !visible);
          if (!visible) {
            const box = row.querySelector('.tdocPick');
            if (box) box.checked = false;
          }
        });
      }
      // LEGACY-UPGRADE-006B (Part G, batch deadline): copies the batch
      // date/time/TZ controls into each given row's OWN per-row inputs --
      // this is purely a convenience that pre-fills the per-row fields the
      // server will actually read at Generate time; it never bypasses
      // server-side validation (validateEmailExportDeadline_()).
      function applyDeadlineToRows(rows) {
        const d = document.getElementById('batchDate').value;
        const t = document.getElementById('batchTime').value;
        const tz = document.getElementById('batchTz').value;
        rows.forEach(function (row) {
          const dateInput = row.querySelector('.deadlineDate');
          const timeInput = row.querySelector('.deadlineTime');
          const tzSelect = row.querySelector('.deadlineTz');
          if (dateInput) dateInput.value = d;
          if (timeInput) timeInput.value = t;
          if (tzSelect) tzSelect.value = tz;
        });
      }
      function applyDeadlineToSelected() {
        const rows = Array.prototype.slice.call(document.querySelectorAll('.tdocRow:not(.hiddenByFilter)')).filter(function (row) {
          const box = row.querySelector('.tdocPick');
          return box && box.checked;
        });
        applyDeadlineToRows(rows);
      }
      function applyDeadlineToVisible() {
        applyDeadlineToRows(Array.prototype.slice.call(document.querySelectorAll('.tdocRow:not(.hiddenByFilter)')));
      }
      function generate() {
        // LEGACY-UPGRADE-006B (status filter): only a CURRENTLY VISIBLE
        // (not filtered out) AND checked row is collected -- a row hidden
        // by the status filter is also force-unchecked above, so this
        // check is defense-in-depth, not the only guard.
        const boxes = Array.prototype.slice.call(document.querySelectorAll('.tdocRow:not(.hiddenByFilter) .tdocPick:checked'));
        if (boxes.length === 0) { document.getElementById('status').textContent = 'Select at least one TDoc.'; return; }
        // LEGACY-UPGRADE-006 (canonical/read-only subject follow-up): no
        // subject is collected here -- the canonical Subject is always
        // generated server-side from the selected TDoc's own table data.
        // LEGACY-UPGRADE-006B (Part G/H, deadline): each selection carries
        // its OWN row's current date/time/TZ inputs -- the server re-
        // validates every one of these regardless (validateEmailExportDeadline_()).
        const selections = boxes.map(function (b) {
          const row = b.closest('tr');
          return {
            tableIndex: parseInt(b.getAttribute('data-idx'), 10),
            deadline: {
              date: row.querySelector('.deadlineDate').value,
              time: row.querySelector('.deadlineTime').value,
              tz: row.querySelector('.deadlineTz').value
            }
          };
        });
        // LEGACY-UPGRADE-006B (bulk text controls): the SAME two global
        // values are sent alongside every selection in this one call.
        const globalText = {
          introText: document.getElementById('introText').value,
          discussionText: document.getElementById('discussionText').value
        };
        document.getElementById('genBtn').disabled = true;
        document.getElementById('status').textContent = 'Generating\\u2026';
        google.script.run
          .withSuccessHandler(function (result) {
            document.getElementById('genBtn').disabled = false;
            document.getElementById('status').textContent = result.ok ? ('\\u2705 ' + result.files.length + ' discussion e-mail(s) generated.') : ('\\u274C ' + result.error);
            const resultsBox = document.getElementById('results');
            resultsBox.innerHTML = '';
            // LEGACY-UPGRADE-006D (ZIP archive): shown first, above the
            // individual-file links, so it is the most visible result.
            if (result.ok && result.zip) {
              const zipP = document.createElement('p');
              zipP.innerHTML = 'ZIP archive: <b>' + result.zip.fileName + '</b>';
              resultsBox.appendChild(zipP);
              const zipA = document.createElement('a');
              zipA.href = result.zip.url; zipA.target = '_blank'; zipA.textContent = 'Open ZIP in Drive';
              resultsBox.appendChild(zipA);
              const indivP = document.createElement('p');
              indivP.className = 'hint';
              indivP.textContent = 'Individual .eml files are also available in the same Drive folder:';
              resultsBox.appendChild(indivP);
            } else if (result.ok && result.zipError) {
              const warn = document.createElement('p');
              warn.style.color = '#a94442';
              warn.textContent = '\\u26A0 ' + result.zipError;
              resultsBox.appendChild(warn);
            }
            (result.files || []).forEach(function (f) {
              const a = document.createElement('a');
              a.href = f.url; a.target = '_blank'; a.textContent = f.fileName;
              resultsBox.appendChild(a);
            });
          })
          .withFailureHandler(function (error) {
            document.getElementById('genBtn').disabled = false;
            document.getElementById('status').textContent = '\\u274C ' + error;
          })
          .generateTdocDiscussionEmails(selections, globalText);
      }
    </script>
  `).setWidth(820).setHeight(720);

  ui.showModalDialog(html, 'Prepare TDoc Discussion E-mails');
}

/**
 * ADDON-009: the export's configuration, resolved fresh and validated as a
 * whole -- Legacy's generateTdocDiscussionEmails() preflight, moved into one
 * function so the dialog can show it and Generate re-checks it. Throws the
 * Legacy error message for the first problem found:
 *
 *   - senderAddress (From): DISCUSSION_EMAIL_SENDER, never derived;
 *   - recipientAddress (To): the meeting-context mailing list
 *     (getMeetingContext_().sources.mailingList -- an ad-hoc meeting's saved
 *     MAILING_LIST, else the report-family list; main meetings unchanged),
 *     through deriveEmailExportRecipientFromMailingList_(). An invalid saved
 *     list is refused, never swapped for another list;
 *   - listTag: deriveEmailExportListTag_() of that recipient;
 *   - revisionUploadUrl: getMeetingContext_().sources.revisionsUrl.
 */
function resolveEmailExportConfiguration_() {
  const senderAddress = resolveEmailExportSenderAddress_();
  if (!isValidEmailExportSenderAddress_(senderAddress)) {
    throw new Error('Discussion E-mail Sender is not configured or is invalid. Open Configure Meeting Settings (E-mail Configuration) and set the Discussion E-mail Sender before generating discussion e-mails.');
  }
  const mailingList = String(getMeetingContext_().sources.mailingList || '').trim();
  const recipientDerivation = deriveEmailExportRecipientFromMailingList_(mailingList);
  if (!recipientDerivation.valid || !isValidEmailExportSenderAddress_(recipientDerivation.recipient)) {
    throw new Error((recipientDerivation.error || 'Could not derive a discussion recipient from the configured Mailing List.') + ' Open Configure Meeting Settings and set the Mailing List before generating discussion e-mails.');
  }
  const recipientAddress = recipientDerivation.recipient;
  // TEMPLATE-002C: Reply-To is the effective mailing list itself -- the same
  // getMeetingContext_() value the recipient comes from (an ad-hoc meeting's
  // Mailing List, else the report-family list; a main meeting always the
  // family list), normalized with the collector's rule. It is never the
  // sender, and a list that cannot be normalized refuses the export instead
  // of producing a misleading header.
  const replyToAddress = deriveEmailExportReplyToFromMailingList_(mailingList);
  if (!replyToAddress) {
    throw new Error('Could not derive the Reply-To address from the Mailing List "' + mailingList.replace(/[\r\n]/g, ' ') +
      '". Open Configure Meeting Settings and set the Mailing List to the ETSI list name (or its ...@list.etsi.org address) before generating discussion e-mails.');
  }
  // The list tag names the files and the default introduction; it is no
  // longer part of the subject (TEMPLATE-002C).
  const listTag = deriveEmailExportListTag_(recipientAddress, getReportConfig_().DRAFTS_FOLDER);
  if (!listTag) {
    throw new Error('Could not derive the discussion list tag from the Mailing List "' + mailingList + '".');
  }
  const revisionUploadUrl = resolveEmailExportRevisionUploadUrl_();
  if (!isSafeEmailExportUrl_(revisionUploadUrl)) {
    throw new Error('Revision upload location is not configured or is invalid (checked getMeetingContext_().sources.revisionsUrl / REVISIONS_URL) -- cannot generate discussion e-mails without a real upload destination.');
  }
  return { senderAddress: senderAddress, recipientAddress: recipientAddress, replyToAddress: replyToAddress, listTag: listTag, revisionUploadUrl: revisionUploadUrl };
}

/**
 * TEMPLATE-002C: "<list>@list.etsi.org" for an effective mailing list given
 * as "<list>" or "<list>@list.etsi.org" (normalizeEtsiListName_(), the
 * collector's rule; the list name keeps its spelling). '' for anything else.
 */
function deriveEmailExportReplyToFromMailingList_(mailingList) {
  const name = normalizeEtsiListName_(mailingList);
  return name ? name + '@list.etsi.org' : '';
}

/**
 * LEGACY-UPGRADE-006: the dialog's "Generate" RPC. `selections` is
 * [{ tableIndex }, ...] -- one group per selected TDoc for this stage
 * (see buildEmailExportForGroup_()'s own header comment for the
 * forward-compatible group data model). Re-scans the document fresh
 * (never trusts stale client-side data for the actual table content, only
 * for which index was picked) so the exported content is always the
 * document's current state. Read-only against the Doc; READS the
 * configured mailing list from Document Properties, WRITES only new files
 * into the dedicated Drive export folder.
 *
 * LEGACY-UPGRADE-006 (canonical/read-only subject follow-up): the Subject
 * is part of the machine-readable contract that associates a reflector
 * reply back with the right TDoc (the ADDON-008A2 collector),
 * so it is ALWAYS generated here, server-side, from the freshly-read
 * tdoc/title/agendaItem -- an `sel.subject` field is never read even if a
 * caller supplies one (removing the UI input alone would not be enough;
 * this is the actual enforcement point). This is deliberately NOT the
 * same as buildEmailExportForTdocTable_()/buildEmailExportForGroup_()'s
 * own optional subjectOverride parameter (kept there for those functions'
 * general reusability/testability) -- this RPC simply never passes one.
 *
 * LEGACY-UPGRADE-006B (bulk text controls): `globalText`, if supplied, is
 * `{ introText, discussionText }` -- plain text from the dialog's two
 * GLOBAL text areas, applied identically to EVERY e-mail generated in this
 * one call (never per-TDoc). Omitted/blank fields fall back to this file's
 * own default wording exactly as before this stage. Never persisted
 * anywhere (not Document Properties, not any other storage) -- it lives
 * only for the duration of this single RPC call.
 *
 * LEGACY-UPGRADE-006B (reserved-status hard exclusion): the server-side
 * re-check below now also rejects a "reserved" TDoc (isEmailExportReserved_()),
 * with the exact same defense-in-depth reasoning as the pre-existing
 * Approved/Agreed re-check -- the dialog never offers one as selectable,
 * but a tampered/direct RPC call must still be refused.
 *
 * LEGACY-UPGRADE-006B (Part H, deadline validation): each selection may
 * carry its own `deadline: {date, time, tz}` -- intentionally client-
 * supplied (the user sets/overrides it in the dialog), but NEVER trusted
 * blindly: validateEmailExportDeadline_() re-checks it here regardless of
 * what the dialog already validated client-side. A selection with no
 * deadline, or one that fails validation (bad calendar date, non-24h time,
 * unsupported time zone), is refused with a clear, actionable error naming
 * the TDoc and the problem -- no email is generated for it, and (matching
 * this RPC's existing defense-in-depth behavior for an excluded TDoc in a
 * mixed batch) no email is generated for any OTHER TDoc in the same call
 * either, rather than silently partially succeeding.
 *
 * LEGACY-UPGRADE-006C (revision-upload instruction): the revision-upload
 * URL is resolved ONCE here, fresh, from trusted meeting configuration
 * (resolveEmailExportRevisionUploadUrl_() -- getMeetingContext_().sources.revisionsUrl,
 * the SAME source scanRevisionsFolder_()/"Test All Connections" already
 * use) -- it is the SAME destination for every TDoc in this call, never
 * per-selection, and is NEVER read from `selections` or `globalText` (no
 * client-supplied revision URL is ever accepted). If it cannot be resolved
 * to a safe http(s) URL, the WHOLE request is refused up front, before any
 * table is even inspected -- this Discussion workflow always expects a
 * real revision-upload destination to exist, so a missing one blocks
 * generation with an actionable error rather than producing an e-mail with
 * a misleading or absent upload instruction.
 *
 * LEGACY-UPGRADE-006D (ZIP archive): after every selected TDoc's .eml has
 * been individually created exactly as before, this ALSO bundles the SAME
 * in-memory blobs from THIS run (never a Drive-folder rescan, so an older
 * .eml or a previous run's ZIP can never be included) into one ZIP archive
 * via the native Utilities.zip() -- no external dependency. The ZIP is
 * purely additive convenience output: every individual .eml is still
 * created in the same Drive folder exactly as before this stage. If any
 * validation/generation step above throws, nothing (no .eml, no ZIP) is
 * created, unchanged from before. If ZIP creation itself fails AFTER the
 * individual .eml files already succeeded, those files are NOT rolled
 * back or deleted -- the response still reports `ok: true` and the
 * created `files`, with `zip: null` and a `zipError` explaining that the
 * individual files exist but the archive could not be created.
 *
 * LEGACY-UPGRADE-006H (corrects 006F/006G's mistaken merge of sender and
 * recipient into one address): TWO INDEPENDENT values are resolved here,
 * fresh, before any table is even inspected --
 *
 *   - `senderAddress` (From): resolveEmailExportSenderAddress_() reads the
 *     explicitly configured DISCUSSION_EMAIL_SENDER Document Property,
 *     then a final isValidEmailExportSenderAddress_() sanity gate. NEVER
 *     derived from Mailing List, NEVER from
 *     Session.getActiveUser()/getEffectiveUser(), NEVER from
 *     `selections`/`globalText` (a client-supplied
 *     `from`/`sender`/`discussionEmailSender` field is simply never read).
 *   - `recipientAddress` (To): deriveEmailExportRecipientFromMailingList_()
 *     derives the reflector address from the meeting-context mailing
 *     list (ADDON-009; Legacy: the MAILING_LIST Document Property),
 *     then the SAME isValidEmailExportSenderAddress_() sanity gate. NEVER the same value as the sender, NEVER itself
 *     independently configured, and NEVER from client-supplied request
 *     data (a `to`/`recipient`/`mailingList` field is simply never read).
 *
 * If EITHER is missing or invalid, the WHOLE request is refused up front
 * with a clear, actionable error naming which one (Discussion E-mail
 * Sender vs. Mailing List) and pointing at Configure Meeting Settings --
 * this blocks ONLY Discussion Email Export, never report build/Continuous
 * Update/general meeting readiness, which never call this function or
 * read either property this way.
 */
function generateTdocDiscussionEmails(selections, globalText) {
  try {
    const exportConfig = resolveEmailExportConfiguration_();
    if (!Array.isArray(selections) || selections.length === 0) {
      throw new Error('Select at least one TDoc.');
    }
    const body = DocumentApp.getActiveDocument().getBody();
    const tables = body.getTables();
    // 2.15.2: base for the report's own tab/heading links.
    const documentUrl = buildEmailExportDocumentUrl_(getActiveDocumentIdSafely_());
    const introText = globalText && globalText.introText;
    const discussionText = globalText && globalText.discussionText;

    // ADDON-009: every selection is validated and its e-mail built in
    // memory first; nothing is written to Drive unless all of them succeed
    // (Legacy created each file inside this loop, so a later refusal left
    // the earlier files behind).
    const built = selections.map(function (sel) {
      const table = tables[sel.tableIndex];
      if (!table || !isTDocTable_(table)) {
        throw new Error('Selected table is no longer a valid TDoc table (index ' + sel.tableIndex + ').');
      }
      const tdocCell = safeCellText_(table, 0, 1).trim();
      const parsedTdoc = parseExactSA4DocumentId_(tdocCell);
      if (!parsedTdoc.isValid) {
        throw new Error('Selected table (index ' + sel.tableIndex + ') does not hold a recognized SA4 TDoc number: "' + tdocCell + '".');
      }
      const meta = {
        tdoc: parsedTdoc.raw,
        title: findCellText_(table, 'Title').trim(),
        // LEGACY-UPGRADE-006 (canonical subject + agenda-item follow-up):
        // read fresh from the CURRENT table, exactly like tdoc/title/status
        // above -- never taken from stale client-side data.
        agendaItem: findCellText_(table, 'Agenda Item').trim()
      };
      // LEGACY-UPGRADE-006 (status-filter follow-up): defensive server-side
      // re-check -- the dialog never offers an excluded TDoc as selectable,
      // but this re-verifies against the table's CURRENT status regardless
      // of what the client actually sent, so an excluded TDoc can never
      // enter an export (single or grouped) through this RPC either.
      const currentStatus = findCellText_(table, 'Status').trim();
      if (isEmailExportReserved_(currentStatus)) {
        throw new Error(meta.tdoc + ' is reserved (no content submitted yet) and is excluded from discussion e-mail export.');
      }
      if (isEmailExportStatusExcluded_(currentStatus)) {
        throw new Error(meta.tdoc + ' is already ' + currentStatus + ' and is excluded from discussion e-mail export.');
      }
      const deadlineValidation = validateEmailExportDeadline_(sel.deadline);
      if (!deadlineValidation.valid) {
        throw new Error(meta.tdoc + ': ' + deadlineValidation.error);
      }
      const deadlineText = formatEmailExportDeadline_(deadlineValidation);
      const subjectDeadlineToken = formatEmailExportDeadlineForSubjectToken_(deadlineValidation);
      // LEGACY-UPGRADE-006H: To (recipientAddress, derived from the Mailing
      // List) and From (senderAddress, configured DISCUSSION_EMAIL_SENDER)
      // are DELIBERATELY DIFFERENT values.
      const email = buildEmailExportForTdocTable_(table, meta, exportConfig.recipientAddress, null, introText, discussionText, deadlineText, exportConfig.revisionUploadUrl, exportConfig.senderAddress, subjectDeadlineToken, exportConfig.listTag, documentUrl, exportConfig.replyToAddress); // no subjectOverride: always canonical
      return { tdoc: meta.tdoc, fileName: email.fileName, eml: email.eml };
    });

    const folder = ensureEmailExportDriveFolder_();
    const blobsForZip = [];
    const files = built.map(function (email) {
      // LEGACY-UPGRADE-006D: the SAME blob is used for the individual file
      // AND, below, as this run's own ZIP member -- never regenerated or
      // re-derived, so the ZIP's content is byte-identical to the
      // individual .eml.
      const blob = Utilities.newBlob(email.eml, 'message/rfc822', email.fileName);
      const file = folder.createFile(blob);
      blobsForZip.push(blob);
      return { tdoc: email.tdoc, fileName: email.fileName, url: file.getUrl() };
    });

    let zip = null;
    let zipError = null;
    if (blobsForZip.length > 0) {
      try {
        const zipFileName = buildEmailExportZipFileName_(new Date(), exportConfig.listTag);
        const zipBlob = Utilities.zip(blobsForZip, zipFileName);
        const zipFile = folder.createFile(zipBlob);
        zip = { fileName: zipFileName, url: zipFile.getUrl() };
      } catch (zipEx) {
        zipError = 'Individual .eml files were created successfully, but the ZIP archive could not be created: ' + zipEx.message;
      }
    }

    return { ok: true, files: files, error: null, zip: zip, zipError: zipError };
  } catch (e) {
    return { ok: false, files: [], error: e.message, zip: null, zipError: null };
  }
}
