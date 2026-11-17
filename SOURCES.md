# Source research and adoption

Reviewed 17 November 2026. The decision was to select high-transfer concepts for a beginner-to-reporting path, not mirror seven websites or reproduce examination questions. All lessons, tasks and bundled data are original public-edition material.

## Coverage and method

All **seven requested landing pages** were read. Selected deeper pages were ranked by missing prerequisites, everyday reporting value, technical uncertainty and redistribution risk. The manifest records **19 source-page URLs**, including terms, plus **five primary certification references**. This is bounded sampling, not a claim to have read every article, downloadable workbook or question.

Live relay observations: W3Schools' Tryit sheet rendered; Excel Practice Online's selected practice region stayed blank in this session while its download link remained visible; the OpenExamPrep question interface rendered but an interaction was blocked by an overlay and a later snapshot timed out. These are session observations, not a diagnosis that the sites fail for everyone. Certiport's rendered expiry-policy body was readable even though static extraction returned no body. The public course therefore does not depend on remote embeds, accounts, advertisements or provider availability.

The implementation itself is checked separately in real Excel and isolated browsers. An attractive website or a question count does not establish teaching accuracy or learner mastery.

## What was adopted

| Source | Valuable idea | Public implementation | Intentionally excluded |
|---|---|---|---|
| [W3Schools Excel](https://www.w3schools.com/excel/) | Short text topics; visible, editable examples; reset/copy; immediate feedback; free Previous/Next | All 60 lessons and the topic rail; x01 orientation/view skills; separate question and practice records | Copied text, graphics, branding, sample data and account/streak dependency |
| [Excel Practice Online](https://excel-practice-online.com/) | Explanation immediately followed by active practice; targeted text-function exercises | s33 and x05 delimiter/type checks; every lesson has independent Excel work | Required OneDrive embeds, copied workbooks and global skill labels from a short quiz |
| [Excel Easy](https://www.excel-easy.com/) | Small single-purpose examples; clear Goal Seek input/target/result | x02 blank distinctions; x09 original break-even Goal Seek exercise; separate initial/changed checks | Unqualified blank/ISBLANK equivalence; changing global iteration settings by default; copying images |
| [GeeksforGeeks Excel](https://www.geeksforgeeks.org/excel/excel-tutorial/) | Repeatable query cleaning and explicit source-path recovery | s43–s45 and s49; original office and monthly-sales data | Power View/ActiveX detours, unsafe identity 'corrections', malformed quoted M examples and magical portable-path claims |
| [Wise Owl exercises](https://www.wiseowl.co.uk/excel/exercises/standard/) | Independent business tasks, difficulty progression, saved deliverables and attempt-before-answer | s07/s12/s16/s24/s29/s41/s48 projects; folded worked checks; final report/handover | All Wise Owl exercise text, numerical scenarios, images, files and model answers: written reproduction permission is required |
| [MO-210 guide](https://open-exam-prep.com/study-guides/mo-210) | Objective-linked reference and import preview checks | s43 delimiter/headers/ID types; x01 freeze/find; x03 named inputs; x04 filtered totals | Copied question banks; treating knowledge checks as performance mastery; attributing Power Query to the expert exam without official support |
| [MO-211 guide](https://open-exam-prep.com/study-guides/mo-211) | Optional advanced path after the core: auditing, financial assumptions and what-if analysis | x07 LET/auditing, x08 PMT, x09 Goal Seek and explicit exam gap route | Non-expiry claim, short-hour expert guarantees, mandatory macro enabling and a full expert-coverage claim |

The raw list of selected pages and observations is in [source-manifest.json](source-manifest.json). Unselected catalogue pages, paid products, logins and answer downloads remain unreviewed. No Wise Owl workbook or model answer was downloaded for this public product.

## Corrections that changed the course

- [Excel Easy's blank-cell page](https://www.excel-easy.com/examples/if-cell-is-blank.html) says two approaches give the same result, while its own note distinguishes a formula returning empty text. Native Excel confirms: truly empty → ISBLANK TRUE / COUNTBLANK 1 / COUNTA 0; `=""` → FALSE / 1 / 1; single space → FALSE / 0 / 1. **x02 teaches the distinction.**
- A sampled [GeeksforGeeks path article](https://www.geeksforgeeks.org/excel/how-to-create-a-relative-file-path-with-power-query-in-excel/) mixes `filepath` prose with `CELL("filename",...)`, inconsistent cell addresses and quoted M fragments. **s49 teaches explicit Source-setting repair and verifies it after a folder move**, rather than promising that a copied formula makes every workbook portable.
- The [MO-210 import article](https://open-exam-prep.com/study-guides/mo-210/import-and-navigate/import-text-files) attributes advanced query transforms to MO-211. The official expert outline does not list Power Query. **The course teaches it as a job skill, not an official MO-211 objective.**
- The sampled [MO-211 date lesson](https://open-exam-prep.com/study-guides/mo-211/date-time-text-auditing/current-date-weekday-functions) warns about default WEEKDAY numbering but overstates which days exceed 5: with Sunday=1, Friday=6 and Saturday=7 both do. Do not copy a weekend test without checking its return type and Friday/Saturday/Sunday cases.
- OpenExamPrep's MO-211 landing says the credential does not expire. [Microsoft Learn](https://learn.microsoft.com/en-us/credentials/certifications/mos-excel-expert-m365-apps/) specifies **60-month renewal**. [Certiport's policy](https://certiport.pearsonvue.com/Educator-resources/Exam-policies/Certification-expiration.aspx) says MOS expiry started being applied after **23 September 2025** and earlier credentials are unaffected. Verify the exact policy for a credential before paying.
- Independent review found the inherited browser SUMIFS demo treated `*` as literal text, unlike Excel's wildcard criterion. The public demo now **visibly rejects unsupported wildcard/comparison-text criteria**; native Excel demonstrates the real wildcard behaviour. It does not return a misleading zero or pretend to support every Excel criterion.

## Official MOS route — separate from job readiness

| Route | Official scope | Course relationship |
|---|---|---|
| [MO-210 Associate](https://learn.microsoft.com/en-us/credentials/certifications/mos-excel-associate-m365-apps/) | Workbooks 25–30%; cells/ranges 25–30%; tables 10–15%; formulas/functions 15–20%; charts 15–20% | Substantial overlap. Not a full objective-by-objective exam preparation claim. |
| [MO-211 Expert](https://learn.microsoft.com/en-us/credentials/certifications/mos-excel-expert-m365-apps/) | Settings 10–15%; data 30–35%; advanced formulas/macros 25–30%; advanced charts/tables 25–30% | Later lessons sample useful skills. Macros, all advanced chart types, workbook collaboration/settings and several other objectives are not fully taught or assessed. |

Use the exact official outlines:

- [MO-210 skills measured](https://arch-center.azureedge.net/Learning/Credentials/microsoft-certified-security-operations-analyst-skills-measured.pdf). The filename is misleading; the retrieved document's title/body are the Excel MO-210 outline.
- [MO-211 skills measured](https://arch-center.azureedge.net/Learning/Credentials/MO-211_OD_MOS365_ExcelExpert.pdf).

For an exam decision, make your own checklist from the current outline: **can demonstrate unaided / practised with help / not practised**. Repeat performance tasks from a blank or supplied workbook. Multiple-choice scores and self-ticks are not equivalent to a proctored performance exam. Microsoft recommends approximately 150 hours of instruction and hands-on product experience; third-party short study-time estimates are not a validated replacement.

## Rights and evidence limits

A free-to-read site is not automatically free to redistribute. W3Schools' [terms](https://www.w3schools.com/about/about_copyright.asp) restrict mass/verbatim copying. Wise Owl's [sample exercise](https://www.wiseowl.co.uk/excel/exercises/standard/creating-formulae/4871/) expressly prohibits reproduction in whole or part without prior written consent. Other providers' redistribution permission was not established. This release uses links, general spreadsheet facts, original prose, original tasks and synthetic data—not their downloadable assets.

No controlled study of learning outcomes was performed. No first-time human learner acceptance trial or accessibility certification is claimed. Tests show the specified software and workbook behaviours in the recorded environment, not that every learner will become an expert or get a job.
