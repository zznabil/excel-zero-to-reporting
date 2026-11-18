# Verification and limits

Native Excel measurement date: 17 November 2026. Launcher and browser measurement date: 18 November 2026. This file records build-time evidence and limits. Consult [Actions](https://github.com/zznabil/excel-zero-to-reporting/actions) and the [versioned release notes](https://github.com/zznabil/excel-zero-to-reporting/releases) for ref-specific CI and publication verdicts; a local pass alone is not release acceptance.

## Current observed evidence

| Boundary | Result | Evidence and scope |
|---|---|---|
| Formula examples | PASS | 60 initial and changed cases: 120 comparisons against Microsoft Excel 16.0 before save. Base results were checked again after reopening the saved workbook; changed-input results were not rechecked after reopen. Native source hashes and actual matrices: tests/excel-oracle.json. |
| Independent native projects | PASS | 78 checks: blanks/empty-text/space, text extraction, IF boundaries, named input, filtered/manual-hidden totals, Goal Seek, freeze panes, validation, conditional formatting, office/HR measures, real local-file Power Query, Table-backed Pivots, charts, source append, stale Pivot detection, independent copy/rebind/probe/move/recovery, save/reopen and PDF geometry. Actual results: tests/project-oracle.json. |
| Real GUI CSV import | PASS | Imported Staff-training.csv through Excel's actual From Text/CSV and Power Query controls. Removed the automatic lossy type step, set explicit types and loaded the Table. Then 25 native checks verified six text IDs through refresh and saved-copy reopen; numeric 1 formatted as 001 and text 1 were rejected. Recipe: tests/GUI-IMPORT.md. Results: tests/import-oracle.json. |
| Deterministic formula/data checks | PASS | Seven Node tests compare the native oracle, literal-criteria limitations, missing/type/zero-denominator cases and independent CSV totals. |
| Browser learner journeys | PASS | Eighteen browser tests, including all 60 lessons: initial/changed/reset outputs match native results, empty/wrong/right questions, specific worked checks, independent completion, navigation, reload/history; genuine backup/restore/cancel, invalid/foreign records, legacy isolation, corrupt/blocked/quota storage, phone/theme/print, offline inputs, copy fallback and schedule. Text extraction and IF boundaries match independent native results; wrong/missing answers and a wrong calculator are rejected. History API rejection, native form restoration and real back-forward caching are covered. The cached-document guard must observe a real persisted pageshow and reject a deliberately omitted recovery handler. File preview preserves old saved bytes and exports current in-memory progress without claiming automatic saving. |
| Local launcher | PASS | Four real HTTP integration tests cover public CSV content and HEAD, private-path/traversal/Host/write rejection, exclusive port ownership and restart at the same origin. Windows needed exclusive socket binding; the duplicate-listener test failed before repair and passed after it. CLI help succeeds; unsupported network-binding arguments fail without a readiness marker. The Windows batch launcher was executed, and isolated Chrome and Brave learner views retained practice through reload. The default Brave window opened, but desktop lock prevented inspection of that particular window. |
| Rendered report | PASS | Exported native asset PDF rendered and inspected: one page, all four departments, RM units, full chart and RM17700 total. Print-area bounds also checked against the actual chart geometry. |
| Independent critic | RELEASE-SCOPED | Exactly one retained Orca OMP worker, launched as GPT-6.1 Terra with max effort. Reviews found misleading wildcard output, missing worked answers, prerequisite gaps, ineffective copy recovery and licence conflict. These were repaired; the critic accepted the first exact archive locally. A later merged-CI restore failure caused a further repair and invalidated that archive's acceptance. Each replacement candidate needs scoped review; the versioned release records its exact archive and final verdict. |
| Package and publication | RELEASE-SCOPED | The recipe checks an explicit 43-file allowlist, deterministic bytes, source hashes, local linked assets and an embedded inventory. Five calibrated controls cover reproducibility, omission/corruption, private-path/token and forbidden-asset canaries with a safe control, duplicate entries and a required linked script omitted even from the allowlist. Candidate-specific builds, clean extraction, remote history/CI/merge/tag and actual public-download evidence belong to the versioned release notes; none is inferred from a local build. |

The native verifier creates only authored synthetic workbooks in its own output folder. It disables macros/events, uses local-file queries only and closes its own Excel application. No proprietary source workbooks or external data services are required.

## Progress storage boundary

Start `START-COURSE.cmd` on Windows, or run `python -m tools.serve`, for the fixed `http://127.0.0.1:8765` origin. Python 3.10 or later is required; no extra Python packages are used. The launcher serves only allowlisted public course files over loopback. It does not serve personal workbooks or accept uploads.

Direct `file://` opening is a preview, not the supported automatic-saving mode. [MDN documents that file-origin localStorage behaviour is undefined](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage). The preview therefore preserves old saved bytes without writing new automatic progress. Export a JSON backup before leaving it, then restore that backup in the launched course. Origins and browser profiles do not share records automatically.

Native browser form history was separately reproduced restoring obsolete controls despite a correct saved record. App-owned controls now disable autofill, and a genuinely cached document reads the current saved state again on `pageshow`. CI uses [Playwright's full Chromium new-headless mode](https://playwright.dev/docs/browsers#chromium-new-headless-mode), with back-forward caching enabled, rather than assuming headless-shell behaviour proves this browser feature.

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
