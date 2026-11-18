# Rollout of T-2026.10.31 (Code.js 2.26.15): hotfix plan, not yet approved

T-2026.10.31 is T-2026.10.30 (`729ebfa`, Code.js `b3e86e31…`) plus one fix: *Set Up Goes to Plenary*,
and an update of a report that has it, no longer stop when somebody else edits the report while
they write it (docs/GOES_TO_PLENARY.md § Incident).

Release files: Code.js and Release.js change. ReportCreator.js, HyperLink.js and appsscript.json are
those of T-2026.10.30.

Every step below needs approval. Nothing here has been done in production. The local listing agent
is not part of this release and is not touched.

## 0. Validation on a disposable Audio copy

**By you:**

1. Make a copy of "Audio SWG Minutes SA4#138" (*File › Make a copy*), named
   "DISPOSABLE TEST COPY - Audio SA4#138".
2. Send Claude the copy's document ID and its script ID (*Extensions › Apps Script › Project
   Settings*).

**By Claude:** push the T-2026.10.31 candidate into the copy's script only. The push is guarded:
the name must say "DISPOSABLE TEST COPY", the IDs must match, and the script must run
T-2026.10.30. Then read it back.

**By you, in the copy:**

3. *About This Report* shows T-2026.10.31 / 2.26.15.
4. *Check Report Integrity* (read-only): note the "Goes to Plenary" line and "Status dropdowns
   missing".
5. **The incident case.** Have a second person, or a second browser window, type in the minutes
   at the top of the copy without pause. Meanwhile run *Report › Set Up Goes to Plenary…*. Expected
   result, either of:
   - "Goes to Plenary is set up."; the execution log has "changed by someone else … reading it
     again";
   - "set up, but not completely: … edited by someone else … N Status row(s) not made compact yet
     … nothing was written in a wrong place …".

   It must never say "Disposition Summary: unreadable".
6. Stop typing. Run *Set Up Goes to Plenary…* again. Expected:
   - "Goes to Plenary is set up.";
   - every Status row reads `[status ▾] Plenary [☐ No ▾]` once;
   - the four tables are on landscape pages;
   - the typed text and the minutes are kept.
7. *Check Report Integrity*: "Plenary dropdown missing: none", "Status dropdowns missing: none",
   one summary heading.

## A. Production Audio: recovery (separate approval)

The Audio report is "switched on but not set up": it has `GOES_TO_PLENARY = 1`, no compact row, and
its three-table summary is unchanged.

1. *File › Version history › Name current version*: "before Goes to Plenary recovery".
2. Push T-2026.10.31 into the Audio script, guarded and read back as for T-2026.10.30.
3. **Between sessions, with nobody editing:** run *Report › Set Up Goes to Plenary…*. It resumes;
   nothing is done twice.
4. *Check Report Integrity*, then look at the compact rows, the four tables and the minutes.

**Alternative, without the hotfix:** step 3 alone, with T-2026.10.30, at a time when nobody else
has the report open. That is safe: what is done is not done twice, and a refused batch writes
nothing. It fails again, safely, if somebody types meanwhile.

**Rollback:** restore the named version, and delete the document property `GOES_TO_PLENARY`, or
push T-2026.10.30 again.

## B. The other four reports and the master template (separate approval)

Push T-2026.10.31 like T-2026.10.30, one report at a time. Each report needs a fresh backup equal
to its T-2026.10.30 read-back, and a read-back check after the push.

Until then, do not run *Set Up Goes to Plenary…* in them while a session is being minuted.
