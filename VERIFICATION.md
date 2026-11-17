# Verification and limits

Local measurement date: 17 November 2026. This file records build-time evidence and limits. Consult [Actions](https://github.com/zznabil/excel-zero-to-reporting/actions) and the [versioned release notes](https://github.com/zznabil/excel-zero-to-reporting/releases) for ref-specific CI and publication verdicts; a local pass alone is not release acceptance.

## Current observed evidence

| Boundary | Result | Evidence and scope |
|---|---|---|
| Formula examples | PASS | 60 initial and changed cases: 120 comparisons against Microsoft Excel 16.0 before save. Base results were checked again after reopening the saved workbook; changed-input results were not rechecked after reopen. Native source hashes and actual matrices: tests/excel-oracle.json. |
| Independent native projects | PASS | 78 checks: blanks/empty-text/space, text extraction, IF boundaries, named input, filtered/manual-hidden totals, Goal Seek, freeze panes, validation, conditional formatting, office/HR measures, real local-file Power Query, Table-backed Pivots, charts, source append, stale Pivot detection, independent copy/rebind/probe/move/recovery, save/reopen and PDF geometry. Actual results: tests/project-oracle.json. |
| Real GUI CSV import | PASS | Imported Staff-training.csv through Excel's actual From Text/CSV and Power Query controls. Removed the automatic lossy type step, set explicit types and loaded the Table. Then 25 native checks verified six text IDs through refresh and saved-copy reopen; numeric 1 formatted as 001 and text 1 were rejected. Recipe: tests/GUI-IMPORT.md. Results: tests/import-oracle.json. |
| Deterministic formula/data checks | PASS | Seven Node tests compare the native oracle, literal-criteria limitations, missing/type/zero-denominator cases and independent CSV totals. |
| Browser learner journeys | PASS | Fourteen browser tests, including all 60 lessons: initial/changed/reset outputs match native results, empty/wrong/right questions, specific worked checks, independent completion, navigation, reload/history; genuine backup/restore/cancel, invalid/foreign records, legacy isolation, corrupt/blocked/quota storage, phone/theme/print, offline inputs, copy fallback and schedule. Text extraction and IF boundaries match independent native results; wrong/missing answers and a wrong calculator are rejected. A forced History SecurityError failed before repair; the repaired restore preserves lesson, tick and checked answer through reload without History API access. Success checks cannot match "not restored". |
| Rendered report | PASS | Exported native asset PDF rendered and inspected: one page, all four departments, RM units, full chart and RM17700 total. Print-area bounds also checked against the actual chart geometry. |
| Independent critic | RELEASE-SCOPED | Exactly one retained Orca OMP worker, launched as GPT-6.1 Terra with max effort. Reviews found misleading wildcard output, missing worked answers, prerequisite gaps, ineffective copy recovery and licence conflict. These were repaired; the critic accepted the first exact archive locally. A later merged-CI restore failure caused a further repair and invalidated that archive's acceptance. Each replacement candidate needs scoped review; the versioned release records its exact archive and final verdict. |
| Package and publication | RELEASE-SCOPED | The recipe checks an explicit 40-file allowlist, deterministic bytes, source hashes, local linked assets and an embedded inventory. Five calibrated controls cover reproducibility, omission/corruption, private-path/token and forbidden-asset canaries with a safe control, duplicate entries and a required linked script omitted even from the allowlist. Candidate-specific builds, clean extraction, remote history/CI/merge/tag and actual public-download evidence belong to the versioned release notes; none is inferred from a local build. |

The native verifier creates only authored synthetic workbooks in its own output folder. It disables macros/events, uses local-file queries only and closes its own Excel application. No proprietary source workbooks or external data services are required.

## Failure sensitivity

- A deliberately wrong native formula (`=999999`) was rejected, including on reopen; process exit was nonzero.
- A deliberately zeroed capstone total was rejected as `0 != 17700`; process exit was nonzero.
- The project workflow observes the stale Pivot after query refresh and does not accept it as the new total.
- A copied workbook is explicitly rebound to a disposable copied source. A probe changes only the copy to 9 records / 14 units / RM12623; the original remains 8 / 13 / RM12500. Moving the copied source causes an observed refresh failure. Repairing Source restores the copy, including after save/reopen.
- Wrong and omitted worked answers are rejected by the same browser-to-native comparison that accepts the correct answer. A zero-output browser calculator is rejected by its native-result comparison. Omitted/modified ZIP members, private-path/token canaries, forbidden media/workbooks/nested archives, duplicate ZIP entries and a missing required linked script are rejected by the actual release verifier. Positive controls pass. Success-only markers appear only after assertions pass.

## What this does not prove

- No first-time human learner study or independent human mastery assessment was performed. Automated full-course interaction is not proof that a beginner has learnt every skill.
- Core examples and selected independent project workflows are verified. Not every possible input, locale, Excel build, keyboard layout, assistive technology or platform is covered.
- Browser demos implement only their displayed rule. They do not evaluate arbitrary Excel formulas or claim full wildcard, locale, blank-cell, sorting or number semantics. Unsupported cases are visibly rejected where applicable.
- The Windows native checks are not run in Linux CI, which has no licensed Excel installation. CI can check the recorded hash-bound native oracle and current browser/data/package behaviours; it is not a fresh Excel calculation.
- The course is account-free and offline; Excel licensing/activation is separate. Full Mac/Web/mobile workbook feature parity is not claimed.
- Selected keyboard, focus, phone, dark and print journeys were checked. No full WCAG audit, accessibility certification, Easy-to-Read intended-user review or standards-conformance claim is made.
- The one critic is independent of the builder but is reused across review rounds; final review is not a fresh second-model context. Effective launch model/effort were observed, not independently verified backend-serving telemetry.
- SHA256 and the inventory detect corruption and inconsistent assets. They do not establish publisher authenticity, a complete security audit, SLSA level or formal supply-chain certification.
- Software is licensed under MIT; original content and synthetic data use CC BY 4.0. LICENSE defines their separate scopes. Third-party dependencies retain their own licences; third-party media/exercises are not included.

## Reproduce

Use the exact commands in README.md. The public release archive is built from release-files.json only. Generated workbooks, progress backups, raw browser traces, local paths, videos, subtitles, node_modules and unrelated files are excluded. Recheck affected gates after changing source, data, verifier, dependencies, environment or release entrypoint. Never relabel a failed check as success.
