# Check the real beginner CSV import

For a maintainer using Microsoft 365 desktop Excel on Windows. Use only the supplied synthetic file and a new QA workbook. Do not open private recovered workbooks, change Office accounts, enable macros or lower security settings.

1. Create `test-results` if it does not exist. Create a blank Excel workbook.
2. Choose Data → From Text/CSV and browse to `Data/Staff-training.csv`. In a Windows file picker, use a normal Windows path or browse to the file; a pasted forward-slash path may be rejected.
3. Inspect the default preview. With automatic type detection it may show EmployeeID `1`, not `001`. This is the failure to prevent; do not choose Load and assume formatting will repair it.
4. Choose Transform Data. In Applied Steps remove the automatic Changed Type step using its X. Confirm the raw first EmployeeID is `001` again.
5. Select EmployeeID and choose Data Type → Text. Keep BookingID and other labels as Text. Set Hours to Whole Number, StartDate and DueDate to Date. Inspect `001` before loading.
6. Choose Close & Load → Close & Load To → Table → New worksheet. Verify six records and `001` in the actual worksheet.
7. Use Data → Refresh All. **Wait for “6 rows loaded” before saving.** If Excel warns that saving will cancel a pending refresh, cancel the save attempt, let the refresh finish, then save.
8. Save a new `test-results/ui-training-import-1.xlsx` through Save As, as an Excel Workbook. Use another numeric suffix if that file already exists; do not overwrite someone else's workbook.
9. Run the read-only/copy-based verifier against the file you actually created:

```powershell
powershell -NoProfile -File tests/native-import.ps1 -Workbook ./test-results/ui-training-import-1.xlsx
```

The verifier creates a separate Excel application, opens the GUI-authored workbook read-only, verifies every EmployeeID and the date type, refreshes the actual saved query, saves an owned verification copy and reopens it. It also rejects numeric `1` formatted to look like `001`, and stripped text `1`. A matching display alone must not pass.

Expected: process exit zero and `GUI_IMPORTED_TEXT_IDS_VERIFIED`, with all checks passed in `test-results/ui-import-results.json`. A missing/stripped ID, failed refresh or mismatched date is a failure, not completed practice. Fix the import steps and repeat; do not edit expected values.

This is a GUI-first verification procedure, not an unattended replacement for walking the user's ribbon/file-picker workflow. Raw screenshots, traces, workbooks and local paths stay in ignored test-results and are not release assets.
