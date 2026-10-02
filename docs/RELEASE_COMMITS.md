# Template releases: deployed artifacts and public commits

The history of this repository was sanitized once, before its first publication on
2026-10-02: a personal e-mail address, the identifiers of real Apps Script projects and
Google documents, and one person's contact details in a captured test fixture were replaced
by synthetic values. Commits from that period therefore have different hashes here than they
had when the first five template releases were built and deployed.

For those five releases there are two things with the same release name. They must not be
confused.

| | A. Historical deployed artifact | B. Public reproducible release |
|---|---|---|
| What it is | The exact bundle that was generated and deployed on 2026-10-01 | The bundle the release tool builds from the public tag |
| Built from | The commit as it was before the sanitization | The public commit the tag points to |
| `Release.js` records | The pre-publication commit hash and the original build time | The public commit hash and its own build time |
| Where it is | Running in the template and in the reports created or migrated at the time; kept locally, outside Git | Reproducible by anyone from the tag |
| Validates against the public tag | No, and it is not meant to | Yes |
| Used by the release and adoption tool | No | Yes |

*About This Report* and *Template Release Info* in documents deployed at the time show the
hash of column A. That is correct historical information: it is the commit that release was
really built from. The public commit hash was never embedded in a deployed artifact.

## Mapping

| Release | `Code.js` | A. Commit recorded in the deployed `Release.js` | B. Public commit (tag `template-release/<release>`) |
|---|---|---|---|
| T-2026.10.0 | 2.17.0 | `d8cb0758139154966324153847c7309caaac649c` | `86bbcfdf4f6ce87a72557e05252a76a2c51ac068` |
| T-2026.10.1 | 2.17.1 | `6b65ebd02ec871168fa0282b6aa90a253712b2ff` | `c4aa524c5fb281c5aafdbf51bc83a7ac6ca4094e` |
| T-2026.10.2 | 2.17.2 | `77a23f3557346d0621a2d562471a4b69a9b68398` | `a5db7f855d2f55b83811b8c1b34bfc66e314bcc5` |
| T-2026.10.3 | 2.17.3 | `39dbfa1822c5e80a6f7759aa34e68477bdfb50be` | `3d7ab824767344e13da91474c77a74d357013cbf` |
| T-2026.10.4 | 2.17.4 | `962fb7de1f58302b54c51eeafa3c5c57d50f12c3` | `cb3eb1df8611bf82a1a761eb79b3d72c292fa5bf` |

Each pair is the same commit in the two histories: same position, same message, same date.

## What differs between A and B

Compared file by file for all five releases:

| File | Difference | Effect at run time |
|---|---|---|
| `Code.js` | Two lines: the example address shown as placeholder in the *Discussion E-mail Sender* field of Configure Meeting, and the same address in one comment | The hint text in an empty field reads `reporter@example.com`. No logic differs. |
| `Release.js` | `gitCommit` and `builtAt` | What *About This Report* displays |
| `ReportCreator.js` | Identical. In the deployed T-2026.10.1 to .3 bundles the file had LF line endings, in the rebuilt ones CRLF. | None |
| `HyperLink.js`, `appsscript.json`, both HTML files | Identical | None |

Releases built after the publication have only one form: A and B are the same.

## Building a public bundle from its tag

The release tool builds a release only from a checkout whose HEAD carries the release tag, so
that every bundle maps to exactly one commit. To build an older release, check out its tag in
a separate worktree and run that tag's own tool there:

```
git worktree add --detach ../sa4-report-T-2026.10.4 template-release/T-2026.10.4
cd ../sa4-report-T-2026.10.4
cp template/template-target.example.json template/template-target.json   # then fill in your template
node tools/template-release.js --release T-2026.10.4 --write
```

The bundle is written to `dist/template-release/T-2026.10.4/`. Placed in the same folder of
the main checkout, it is what the adoption mode (`--adopt`) uses and validates against the tag.

## Other notes

- Documents under `docs/` written before the publication quote commit hashes of column A, and
  short hashes of other commits from that time. The commit subjects are unchanged and can be
  used to find them in the public history.
- The release tool lists the projects it refuses to push to by the SHA-256 digest of their
  Script ID, not by the ID.
