# Audio SWG SA4#138 — large-build recovery (2026-11-11)

## Confirmed scope
Audio SWG is a subgroup report of SA4#138 Calgary. Agenda S4-261857 and the shared SA4#138 TDoc list are correct; do not change them. Other reports remain untouched.

## Observed execution
- Skeleton completed: 10 agenda items, 111 TDocs.
- Abstracts: 100 inserted out of 107 candidate tables.
- Revisions and formatting phases reported completion.
- Final status-dropdown phase reported: `Too many changes applied before saving document. Please save changes in smaller batches using Document.saveAndClose(), then reopen the document with Document.openById().`
- Execution ended in error after dropdown phase was skipped. Do not equate the logged phase completion with a verified persisted document.

## Critical existing design
`docs/STATUS_DROPDOWN_RELEASE_CANDIDATE.md` states that `finalizeStatusDropdowns_()` already calls `saveAndClose()` before Docs API read/write and processes dropdowns in batches of 50. The tests already simulate 130 TDocs. The failure may arise at the initial save, from accumulated DocumentApp mutations; another unconditional final save is not a demonstrated fix.

## Investigation and implementation checklist
1. Back up the existing Audio document and inspect whether its skeleton/abstracts and status text persisted. Record revision ID and document size.
2. Trace `runFullReportBuildCore_()`, `finalizeStatusDropdowns_()`, `buildSkeletonWithTdocTables()` and all uses of the active DocumentApp object; identify the exact throwing operation and whether the catch/finally path triggers a second save.
3. Reproduce the >100-TDoc mutation scenario; extend tests to exercise a DocumentApp save failure due to too many unsaved changes, not only Docs API batch limits.
4. Implement safe, tested DocumentApp checkpointing at phase boundaries **only after verifying all later functions reacquire fresh document/body/table references**. If not, refactor to explicit fresh-document access. Avoid stale element references.
5. Preserve manual minutes and dropdown selections; provide an idempotent repair-only status finalization path on the existing Audio report, with no skeleton clearing or duplicated TDoc tables.
6. Validate against a disposable copy of Audio. Verify saved content, statuses, dropdowns, and update idempotence.
7. PR + CI + review, guarded template release, deploy to Master Template and Audio bound script with hash/readback verification. Do not deploy to other reports.

## Safety gates
No Full Build on the original Audio document, no merge or deployment until review and approval. Keep separate from issue #4 (time zones) and TDoc revision enhancements.

References: issue #6; Audio document `1zawfIcbkzCJG-h2HYtnbQSEsm6MCd9OGhYzMgAFYENs`; Audio Apps Script `1TJAxpoSF2t6RS4FdO9TuCkHeM4eyaI7bDxcZk8v8GkaexZJLiOvKtTN1`.
