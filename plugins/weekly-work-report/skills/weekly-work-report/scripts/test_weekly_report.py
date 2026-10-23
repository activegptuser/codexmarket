#!/usr/bin/env python3
"""Behavior-level regression tests for weekly report patching and validation."""

from __future__ import annotations

import unittest
from datetime import date
from xml.etree import ElementTree as ET

from patch_weekly_report import balance_report_columns, patch_sheet
from validate_weekly_report import inspect_sheet_structure, parse_shared_strings, report_column_widths


def worksheet(rows: str, columns: str = '<col min="3" max="26" width="10" customWidth="1"/>') -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<cols>{columns}</cols><sheetData>{rows}</sheetData></worksheet>'
    )


class WeeklyReportBehaviorTests(unittest.TestCase):
    def test_patch_places_each_line_in_its_exact_row(self) -> None:
        rows = [
            '<row r="3"><c r="X3" s="1"><v>1</v></c></row>',
            '<row r="4"><c r="H4" s="2"><f>TEXT(X3-6,"x")</f><v>old</v></c>'
            '<c r="T4" s="2"><f>TEXT(X3+1,"x")</f><v>old</v></c></row>',
        ]
        for row in range(8, 33):
            rows.append(f'<row r="{row}"><c r="C{row}" s="3"/><c r="P{row}" s="3"/></row>')
        current = ["▶ Project", "   ▷ Task", "      - Detail"]
        next_week = ["▶ Next", "", "   ▷ Planned"]
        values = [value for value in current + next_week if value]
        indices = {value: index for index, value in enumerate(values)}

        patched = patch_sheet(worksheet("".join(rows)), indices, date(2026, 10, 22), current, next_week)
        errors: list[str] = []
        _, cells = inspect_sheet_structure(patched.encode(), errors)

        self.assertEqual([], errors)
        self.assertEqual("0", next(node.text for node in cells["C8"] if node.tag.endswith("v")))
        self.assertEqual("1", next(node.text for node in cells["C9"] if node.tag.endswith("v")))
        self.assertEqual("2", next(node.text for node in cells["C10"] if node.tag.endswith("v")))
        self.assertEqual("3", next(node.text for node in cells["P8"] if node.tag.endswith("v")))
        self.assertNotIn("P9", {ref for ref, cell in cells.items() if cell.get("t") == "s"})
        self.assertEqual("4", next(node.text for node in cells["P10"] if node.tag.endswith("v")))

    def test_balancing_preserves_total_width_and_equalizes_halves(self) -> None:
        original = worksheet('<row r="8"/>')
        balanced = balance_report_columns(original)
        original_widths = report_column_widths(ET.fromstring(original))
        balanced_widths = report_column_widths(ET.fromstring(balanced))

        self.assertAlmostEqual(sum(original_widths), sum(balanced_widths), places=6)
        self.assertAlmostEqual(balanced_widths[0], balanced_widths[1], places=6)

    def test_structure_check_detects_cell_in_wrong_row(self) -> None:
        malformed = worksheet('<row r="8"><c r="C9"/></row>')
        errors: list[str] = []

        inspect_sheet_structure(malformed.encode(), errors)

        self.assertIn("cell C9 is stored in row 8", errors)

    def test_shared_strings_include_rich_text_runs(self) -> None:
        shared = (
            '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<si><r><t>first</t></r><r><t> second</t></r></si></sst>'
        )

        self.assertEqual(["first second"], parse_shared_strings(shared.encode()))


if __name__ == "__main__":
    unittest.main()
