# Changes

## 1.0.0

- Separate portable public edition: 57 core lessons and three optional advanced lessons; no required local/private course files.
- W3Schools-inspired text-first layout with editable fixed-rule examples, reset, copy/fallback, checked questions, independent practice and free navigation.
- Original office-request, HR training, monthly-sales and asset datasets; refreshable reports, readable charts, worked checks and handover practice.
- Added freeze/find, blank versus empty-text semantics, named inputs, visible-row totals, text cleaning, weighted averages, LET/auditing, PMT assumptions and Goal Seek.
- Complete 17-week planning schedule with review and milestone evidence; estimates are not a mastery promise.
- Public-edition progress isolation, real JSON backup/restore, cancellation and corrupt/blocked-storage recovery. Legacy local-course records are not imported as completion.
- Unsupported wildcard/comparison-text criteria now fail visibly in guided calculators instead of showing misleading totals.
- Seven-source adoption/rights manifest and official MOS scope/policy cross-check; no proprietary videos, workbooks, exercise answers or exam dumps bundled.
- Native Excel formula/project checks, isolated browser journeys, deterministic packaging and calibrated failure checks. See VERIFICATION.md for current results and limitations.
- All 60 lessons have specific worked checks; independent text-extraction answers and IF boundaries are checked against native Excel, including wrong/missing-answer fixtures.
- Tables and typed CSV import are taught before first use. A real Excel GUI import was saved, refreshed and reopened with text IDs preserved.
- Disposable-copy recovery now proves the source binding with an isolated probe before moving and repairing it; the original report remains unchanged.
- MIT licence for software and CC BY 4.0 for original educational content and synthetic data, with explicit third-party exclusions.
- Restore/reset no longer depend on History API state changes. Restored state is saved before fragment navigation; import status is cleared while reading. A forced History rejection regression checks the restored lesson, practice tick and answer after reload, and success checks cannot match "not restored".
- Added a standard-library Python launcher and Windows start command. Automatic progress now uses a fixed loopback origin; the read-only server exposes only allowlisted public files and rejects duplicate listeners.
- Direct HTML opening is explicitly a preview: old saved bytes remain untouched, automatic saving is off, and current in-memory progress can be exported and restored in the launched course.
- App-owned controls disable native autofill. Real cached documents reload current saved state on return; normal document history and native back-forward cache tests check practice, answers and counts, including a broken-handler control.
- All progress writers now read and update the record inside one IndexedDB write transaction, with success shown only after commit. Live-tab notifications preserve unsent question choices and example edits. Confirmed replacement/reset remain explicit whole-record actions; unavailable storage falls back visibly to manual JSON, not an unsafe write.
- A new storage generation isolates older public prototype writers without deleting their records. Public-edition JSON backups remain explicitly restorable; private/local backups are still rejected.
- Recovery revalidates changed records, clears obsolete corruption only after a committed replacement, and preserves unsaved memory as an explicit conflict instead of silently discarding or overwriting either record.
