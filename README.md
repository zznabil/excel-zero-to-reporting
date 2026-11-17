# Excel: zero to reporting

A text-first offline Excel course. Start at the first cell. Finish with a report that imports new records, refreshes a PivotTable and explains its checks.

**57 core lessons + 3 optional advanced lessons.** Each has a written explanation, a small editable example, a checked question and independent work in real Excel. Original synthetic data covers office requests, training bookings, sales and IT assets.

## Start learning

1. Download the course ZIP from [Releases](https://github.com/zznabil/excel-zero-to-reporting/releases).
2. Extract the whole ZIP. Do not open the course inside the archive.
3. Open **START-HERE.html** in Chrome or Edge. No installation, server, account or subscription is needed for the course.
4. Create a separate **My-work** folder for your saved Excel workbooks.
5. Follow **Next** in order. Use [the complete schedule](SCHEDULE.html) to plan your week.

Real workbook practice requires a separately licensed and activated **Microsoft 365 desktop Excel on Windows**. The course does not supply or activate Excel. Browser examples are deliberately limited calculators, not a full spreadsheet editor, PivotTable engine, Power Query engine or workbook grader. Unsupported wildcard/comparison criteria are rejected visibly rather than calculated incorrectly.

### The learning loop

Read → predict → change the example → check the question → practise in Excel → inspect the worked check → save and reopen → explain the result → tick practice.

The default schedule is **17 weeks at six hours per week**, including review. Core lesson estimates total 71 hours 20 minutes; these are planning estimates, not measured guarantees. Repeat difficult skills. Three optional Later chapters cover LET/auditing, PMT assumptions and Goal Seek/MOS orientation.

## What you will build

- An office-request report with counts, conditional totals and an RM chart.
- An HR training-booking report that distinguishes bookings, people and hours.
- A sales report that combines monthly CSVs and refreshes from **RM10,200 to RM12,500**.
- An asset report that refreshes from **12 records / 15 devices / RM15,000** to **15 records / 19 devices / RM17,700**, with a visible missing-owner exception.
- A one-page handover explaining source files, record meaning, types, refresh order, checks, limitations and source-path recovery.

These are spreadsheet skills and portfolio exercises. They are not payroll/tax/legal advice, a qualification, or a salary/job guarantee. Python and AI can be later tools; neither is a prerequisite. Do not let an agent invent source data, silently replace errors with zero, or claim a report works without checking the underlying records.

## Progress and backups

Question answers and self-declared workbook practice are tracked separately. A quiz pass does not mark workbook practice complete. Navigation is not locked.

- Progress is local to the browser profile and origin. It does not sync or include Excel workbooks.
- Save a **JSON backup weekly**, and before moving folders, changing browsers, clearing browser data or updating the course.
- Restore requires confirmation. Cancelling or selecting invalid/foreign JSON leaves current progress unchanged.
- If storage is blocked or full, progress remains in memory and the page tells you to save a backup before closing.
- An unreadable saved record is not overwritten. Export its raw bytes before starting fresh if you need them.
- This public edition has a distinct storage key and backup identity. The previous private/local course remains unchanged; its backups are not misinterpreted as public-course completion.

## When something fails

| Failure sign | Next action | What stays safe |
|---|---|---|
| No lesson appears | Extract the whole ZIP; keep all files together; reopen START-HERE.html. Re-download if a script is missing. | Separately saved Excel files |
| Formula shows `#NAME?` | Check spelling, function availability and your Excel build. Use your locale's formula separator if it is `;`, not `,`. | Raw source data |
| A date or number shows `#####` | Widen the column, then inspect the stored value. | The underlying value |
| `#SPILL!` | Inspect the destination; move valuable content rather than deleting it blindly. Put dynamic spills outside Tables. | Original inputs |
| Query fails after files move | Edit its Source setting to the new location. Refresh query, wait for loaded rows, then refresh PivotTable. | Last saved working workbook |
| Query succeeds but chart is old | Refresh the PivotTable after query load and reconcile its totals. | Source records |
| Copy is blocked | Use the selected-text fallback and press Ctrl+C. | All data remains local |
| Progress cannot save | Export JSON before closing. Save workbooks separately. | In-memory progress until the page closes |

Do not enable macros, weaken Trust Center/privacy settings, send diagnostics or refresh unrelated old external connections to complete this course. Mac, mobile and Excel for the Web are not verified for the complete native workbook workflow.

## Sources and scope

[Seven-source research and MOS gap route](SOURCES.md) records what was adopted, rejected and left unverified. Concepts and teaching conventions were researched; third-party lessons, screenshots, videos, workbooks and model answers are not bundled. No affiliation with the linked providers or Microsoft.

[Verification](VERIFICATION.md) distinguishes current evidence from limits. The course is not complete MO-210 or MO-211 exam preparation. Use official objectives for a separate exam gap checklist. No exam dumps are included.

## Develop and verify

Learners need none of these tools. Maintainers need Node.js 22+ and Python 3.10+; native checks also require Windows PowerShell and separately licensed desktop Excel.

```sh
npm ci --ignore-scripts
npx playwright install chromium
npm test
npm run test:browser
npm run package
```

On Windows with Excel, run:

```powershell
npm run test:native
powershell -NoProfile -ExecutionPolicy Bypass -File tests/native-projects.ps1
```

Native scripts create their own Excel application and only authored synthetic workbooks under test-results. They disable macros/events and close only that owned application. Formula checks compare initial and changed results, save, reopen and recheck. The project script exercises real Tables, validation, conditional formatting, freeze panes, Goal Seek, PivotTables, charts, Power Query refresh and moved-source recovery.

The recorded native oracle is tied to hashes of lessons.js and demos.js. Changed formula sources invalidate it. Do not update expected values merely to make a failed test green. Recalculate in Excel and investigate discrepancies first. The exact tested recipe, negative controls and limitations belong in VERIFICATION.md.

## Contribute safely

Keep one prerequisite-respecting path. Prefer short explanations, labelled inputs, original synthetic records, a changed-input check and an independent task with a meaningful outcome. Keep reference and optional advanced work outside the core prerequisites.

Open an issue with the lesson, exact steps, expected and actual result, browser/Excel version and a **redacted** screenshot if useful. Never upload employer/customer/employee records, home addresses, tokens, progress backups or unredacted private paths. Do not contribute copied exercises or exam questions. Public visibility is not an open-source licence grant; see LICENSE before redistributing or proposing licensed contributions.

Run the affected deterministic checks. Browser tests use isolated contexts, not a learner's profile. Do not create a test that only checks a source string or echoes an expected answer.
