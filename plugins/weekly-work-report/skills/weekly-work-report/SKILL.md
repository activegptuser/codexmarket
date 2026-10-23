---
name: weekly-work-report
description: Create the Korean software weekly work report from the 동원정밀 OneNote Dashboard and linked task histories while preserving the existing Excel template, print settings, and one-page layout. Use for requests such as 주간업무보고 작성, 이번 주 업무보고 만들어줘, or 주간업무현황 작성. Do not use for unrelated summaries or live Excel/OneNote UI control.
---

# Weekly Work Report

Read [references/report-rules.md](references/report-rules.md) completely before collecting data or editing a workbook. Also use the spreadsheet artifact skill for workbook inspection and rendering.

## Workflow

1. Resolve the report anchor date, OneNote source, and template according to the reference rules. Explicit user choices override defaults.
2. Read the `동원정밀` notebook, `Dashboard` section, `Dashboard` page with the connected OneNote read tool. Never use PC automation or the Excel/OneNote desktop UI.
3. Locate the Dashboard tables by label. Read every linked Task page used in the report, including `주요 결과물`, start date, end date, and the complete `이력` list. Start with plain text. If a Dashboard hyperlink ID returns `0x80042005`, resolve the exact title with the OneNote search tool and use its local page ID. Apply the duplicate-title rule in the reference file.
4. If the OneNote read tool is unavailable, stop and request a connection or current source material. Do not generate a report from stale cached material.
5. Draft `Completed` as current-week work and `Doing` as next-week work. Dashboard membership decides the side. Use only reporting-period history for dated detail; when none exists, use the verified task title and `주요 결과물` without presenting older history as current-week work. Keep each side within 20 rows and use the hierarchy `▶ project`, `   ▷ task`, `      - detail`.
6. Copy the chosen healthy `.xlsx` template and patch only the anchor date and report text:

   ```powershell
   python scripts/patch_weekly_report.py --template <template.xlsx> --output <output.xlsx> --anchor-date YYYY-MM-DD --current-json <current.json> --next-json <next.json> --balance-columns
   ```

   Each JSON file is a UTF-8 array of strings. A blank string reserves a blank row. `--balance-columns` preserves the combined `C:Z` width while making `C:O` and `P:Z` equal in total width.
7. Validate the result against the exact template:

   ```powershell
   python scripts/validate_weekly_report.py --template <template.xlsx> --output <output.xlsx> --current-json <current.json> --next-json <next.json> --require-balanced-columns
   ```

8. Re-open the generated file with the spreadsheet artifact tool, inspect the date values, and render `B2:Z39`. Confirm no clipping, one-page composition, correct dates, and correct current/next-week classification. Do not export the workbook from the artifact tool. If its renderer shows `#VALUE!` only for the template's unchanged `TEXT(...)` formulas, rely on the validator's formula/cache checks and disclose the renderer limitation; never rewrite the template formulas to accommodate the renderer.
9. Save exactly one output workbook in the required output directory. In the final response, link the output once and list every Dashboard table, linked Task page, and template file used. Disclose every linked page that could not be read.

## Safety and Preservation

- Never reserialize `sheet1.xml`; the patcher performs byte-oriented XML substitutions.
- Never delete or regenerate `calcChain.xml`, printer settings, formulas, styles, merges, dimensions, print area, or print titles. The approved `--balance-columns` operation may change only the `C:Z` column definitions while preserving their combined width.
- Never use Excel UI automation as a validation method.
- Treat validation failure as a hard failure. Do not present the workbook as complete.
