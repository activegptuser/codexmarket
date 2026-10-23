# Weekly Work Report

Weekly Work Report is a personal Codex plugin for creating the Korean software weekly work report from local OneNote task data and an existing Excel template.

It provides the `weekly-work-report` skill, which:

- reads the `동원정밀` notebook, `Dashboard` section, and `Dashboard` page
- follows linked task pages and uses their result, date, and history fields
- maps `Completed` tasks to current-week work and `Doing` tasks to next-week work
- copies a healthy XLSX template and patches only the permitted cells and column definitions
- preserves formulas, styles, merges, print settings, calculation chains, and printer settings
- validates cell placement, package integrity, print ranges, and balanced report widths

## Requirements

- Windows with Python 3.10 or later
- The `onenote-local` plugin installed and connected to Microsoft OneNote desktop
- The Codex spreadsheet artifact capability for workbook inspection and rendering
- A healthy weekly-report `.xlsx` template in `D:\doc\주간업무보고`

The plugin does not include an MCP server. It uses the tools supplied by the separately installed `onenote-local` plugin.

## Usage

Ask Codex to create a weekly work report. The skill selects the report date and newest healthy template unless the request supplies them explicitly. It writes one file to `D:\doc\주간업무보고\outputs\YYYY-MM-DD\`.

The workflow does not use Excel, OneNote, browser, or desktop UI automation. If the OneNote read tools are unavailable, it stops instead of using stale source material.

## Validation

Run the behavior tests from the skill's scripts directory:

```powershell
python -m unittest -v test_weekly_report.py
```

The patcher and validator accept UTF-8 JSON arrays of report lines. See the skill instructions and `references/report-rules.md` for the exact commands and workbook invariants.
