# Excel: zero to reporting

A text-first offline Excel course. Start at the first cell. Finish with a report that imports new records, refreshes a PivotTable and explains its checks.

**57 core lessons + 3 optional advanced lessons.** Each has a written explanation, a small editable example, a checked question, independent work in real Excel and a specific worked check. Original synthetic data covers office requests, training bookings, sales and IT assets.

## Start learning

1. Download the course ZIP from [Releases](https://github.com/zznabil/excel-zero-to-reporting/releases).
2. Extract the whole ZIP. Do not open the course inside the archive.
3. On Windows with **Python 3.10+ installed**, double-click **START-COURSE.cmd**. On other systems, run `python -m tools.serve` from this folder. It opens the course at **http://127.0.0.1:8765/START-HERE.html**. Keep the console open while learning; stop it with Ctrl+C.
4. Create a separate **My-work** folder for your saved Excel workbooks.
5. Follow **Next** in order. Use [the complete schedule](SCHEDULE.html) to plan your week.

Real workbook practice requires a separately licensed and activated **Microsoft 365 desktop Excel on Windows**. The course does not supply or activate Excel. Browser examples are deliberately limited calculators, not a full spreadsheet editor, PivotTable engine, Power Query engine or workbook grader. Unsupported wildcard/comparison criteria are rejected visibly rather than calculated incorrectly.

The offline launcher uses Python's standard library only. No Node.js, package installation, administrator access, course account or subscription is needed. Install Python from [python.org](https://www.python.org/downloads/) first if required; the launcher does not install software or change security settings. If Windows shows a script warning, stop and review the source and your policy. Do not disable security checks.

Opening **START-HERE.html directly** is a **file preview**, not the recommended progress-saving mode. Lessons, examples and downloads work, but automatic saving is deliberately off: browsers do not guarantee storage for `file://` pages. Save a JSON backup before leaving a preview and restore it in the launched course. Existing saved file-preview records are not overwritten.

The local server listens only on 127.0.0.1. It serves the public release allowlist, not directory listings, My-work, test-results, .git or arbitrary files; it accepts no uploads. Keep your private workbooks separate. Use the same address, port, browser and profile each time.

### The learning loop

Read → predict → change the example → check the question → practise in Excel → inspect the worked check → save and reopen → explain the result → tick practice.

The default schedule is **17 weeks at six hours per week**, including review. Core lesson estimates total 71 hours 20 minutes; these are planning estimates, not measured guarantees. Repeat difficult skills. Three optional Later chapters cover LET/auditing, PMT assumptions and Goal Seek/MOS orientation.

## What you will build

- An office-request report with counts, conditional totals and an RM chart.
- An HR training-booking report that distinguishes bookings, people and hours.
- A sales report that combines monthly CSVs and refreshes from **RM10,200 to RM12,500**.
- An asset report that refreshes from **12 records / 15 devices / RM15,000** to **15 records / 19 devices / RM17,700**, with a visible missing-owner exception.
- A one-page handover explaining source files, record meaning, types, refresh order, checks, limitations and source-path recovery.

These are spreadsheet skills and portfolio exercises. They are not payroll/tax/legal advice, a qualification, or a salary/job guarantee. Python programming and AI can be later tools; neither is needed for Excel practice. The launcher requires only the Python runtime. Do not let an agent invent source data, silently replace errors with zero, or claim a report works without checking the underlying records.

## Progress and backups

Question answers and self-declared workbook practice are tracked separately. A quiz pass does not mark workbook practice complete. Navigation is not locked.

- Progress is local to the browser profile and origin. Different addresses, protocols, ports and profiles have separate records. There is no cloud or cross-profile sync, and backups do not include Excel workbooks.
- Wait for **Progress saved in this browser** before closing or reloading. **Saving progress...** means a save is still pending; a memory-only warning means it did not reach browser storage.
- Tabs of this release coordinate saves and update only the selected lesson/action. Unchecking removes that lesson only; confirmed restore and recovery reset intentionally replace the whole record.
- Save a **JSON backup weekly**, and before moving folders, changing browsers, clearing browser data or updating the course. Close old course tabs before updating.
- Restore requires confirmation. Cancelling or selecting invalid/foreign JSON leaves current progress unchanged.
- Automatic saving needs working IndexedDB storage, supported by current Chrome, Edge and Brave. If it is unavailable, blocked or full, progress remains in memory and the page tells you to save a backup before closing.
- An unreadable saved record is not overwritten. Export its raw bytes before starting fresh if you need them.
- If another tab recovers saved progress while this tab has unsaved work, tracking pauses to preserve both records. **Save this tab's JSON backup before reloading.** Reload adopts the saved record; it does not merge unsaved work. A restored backup replaces progress, so do not restore over newer work blindly.
- This public edition has a distinct storage key and backup identity. The previous private/local course remains unchanged; its backups are not misinterpreted as public-course completion.
- Before switching from a file preview, another address or another browser, export JSON in the old view and restore it in the new one. There is no automatic cross-origin migration. File preview requires a backup before every exit, not only weekly.
- Earlier public prototypes use a separate storage generation and cannot overwrite this release. Export JSON from the old public view and explicitly restore it here; its saved bytes are not automatically read, migrated or deleted. Private/local-course backups remain a different format.

## When something fails

| Failure sign | Next action | What stays safe |
|---|---|---|
| No lesson appears | Extract the whole ZIP; keep all files and folders together; restart START-COURSE.cmd. Re-download if a file is missing. | Separately saved Excel files |
| Formula shows `#NAME?` | Check spelling, function availability and your Excel build. Use your locale's formula separator if it is `;`, not `,`. | Raw source data |
| A date or number shows `#####` | Widen the column, then inspect the stored value. | The underlying value |
| `#SPILL!` | Inspect the destination; move valuable content rather than deleting it blindly. Put dynamic spills outside Tables. | Original inputs |
| Query fails after files move | Edit its Source setting to the new location. Refresh query, wait for loaded rows, then refresh PivotTable. | Last saved working workbook |
| Query succeeds but chart is old | Refresh the PivotTable after query load and reconcile its totals. | Source records |
| Copy is blocked | Use the selected-text fallback and press Ctrl+C. | All data remains local |
| Progress cannot save | Export JSON before closing. Save workbooks separately. | In-memory progress until the page closes |
| File preview says saving is off | Export JSON, start the launcher, then restore that JSON in the launched course. | Old saved preview record and separate workbooks |
| Python is missing or too old | Install Python 3.10+ from python.org, or use file preview with manual JSON backups. | Course source and workbooks |
| Port 8765 is already in use | Close a previous course console if it is yours. Do not kill unrelated applications or choose a random port: a different origin has different progress. | Saved browser progress |

Do not enable macros, weaken Trust Center/privacy settings, send diagnostics or refresh unrelated old external connections to complete this course. Mac, mobile and Excel for the Web are not verified for the complete native workbook workflow.

## Sources and scope

[Seven-source research and MOS gap route](SOURCES.md) records what was adopted, rejected and left unverified. Concepts and teaching conventions were researched; third-party lessons, screenshots, videos, workbooks and model answers are not bundled. No affiliation with the linked providers or Microsoft.

[Verification](VERIFICATION.md) distinguishes current evidence from limits. The course is not complete MO-210 or MO-211 exam preparation. Use official objectives for a separate exam gap checklist. No exam dumps are included.

## Develop and verify

Learners need Python 3.10+ only for the offline launcher. Maintainers also need Node.js 22+; native checks require Windows PowerShell and separately licensed desktop Excel. Browser tests start the same allowlisted loopback server as the learner launcher, not an unrelated filesystem server.

```sh
npm ci --ignore-scripts
npx playwright install chromium
npm test
npm run test:server
npm run test:browser
npm run test:package
npm run package
```

On Windows with Excel, run:

```powershell
npm run test:native
powershell -NoProfile -ExecutionPolicy Bypass -File tests/native-projects.ps1
```

Native scripts create their own Excel application and only authored synthetic workbooks under test-results. They disable macros/events and close only that owned application. Formula checks compare initial and changed results before save, then check base results after reopen. The project script exercises real Tables, validation, conditional formatting, freeze panes, Goal Seek, PivotTables, charts, Power Query refresh and moved-source recovery.

For the separately dogfooded Excel GUI import, follow [tests/GUI-IMPORT.md](tests/GUI-IMPORT.md). The procedure starts from the supplied CSV, removes Excel's automatic lossy type step and saves a real GUI-authored workbook. Run its guarded native verifier against that workbook to check refresh, stored text IDs and save/reopen. Do not substitute a workbook generated by a script for the GUI evidence.

The formula oracle is tied to hashes of lessons.js and demos.js. Project and GUI-import oracles also identify their native verifier hashes. Changed formula sources or native verifiers require new measurements. Independent text-extraction and IF-boundary answers are compared with native project results in the browser; wrong or omitted answers are rejected. Do not update expected values merely to make a failed test green. Recalculate in Excel and investigate discrepancies first. The exact tested recipe, negative controls and limitations belong in VERIFICATION.md.

## Contribute safely

Keep one prerequisite-respecting path. Prefer short explanations, labelled inputs, original synthetic records, a changed-input check and an independent task with a meaningful outcome. Keep reference and optional advanced work outside the core prerequisites.

Open an issue with the lesson, exact steps, expected and actual result, browser/Excel version and a **redacted** screenshot if useful. Never upload employer/customer/employee records, home addresses, tokens, progress backups or unredacted private paths. Do not contribute copied exercises or exam questions.

Software is licensed under **MIT**. Original lessons, documentation and synthetic data are licensed under **CC BY 4.0**. [LICENSE](LICENSE) defines the material boundaries, attribution and third-party exclusions. Contribute only material you are entitled to license under the applicable terms.

Run the affected deterministic checks. Browser tests use isolated contexts, not a learner's profile. Do not create a test that only checks a source string or echoes an expected answer.
