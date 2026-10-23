#!/usr/bin/env python3
"""Validate a patched weekly report against its exact XLSX template."""

from __future__ import annotations

import argparse
import html
import json
import re
import zipfile
from datetime import date, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET


SHEET = "xl/worksheets/sheet1.xml"
SHARED_STRINGS = "xl/sharedStrings.xml"
WORKBOOK = "xl/workbook.xml"
ALLOWED_CHANGES = {SHEET, SHARED_STRINGS}
REQUIRED_PRINT_AREA = "$B$2:$Z$39"
REQUIRED_PRINT_TITLES = "$2:$6"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--current-json", type=Path)
    parser.add_argument("--next-json", type=Path)
    parser.add_argument("--require-balanced-columns", action="store_true")
    return parser.parse_args()


def read_zip(path: Path) -> tuple[zipfile.ZipFile, dict[str, bytes]]:
    archive = zipfile.ZipFile(path, "r")
    bad_member = archive.testzip()
    if bad_member:
        archive.close()
        raise ValueError(f"{path} contains a corrupt member: {bad_member}")
    return archive, {name: archive.read(name) for name in archive.namelist()}


def namespace_declarations(root_tag: str) -> dict[str, str]:
    return dict(re.findall(r'xmlns:([A-Za-z_][\w.-]*)="([^"]+)"', root_tag))


def validate_ignorable(sheet_xml: str, errors: list[str]) -> None:
    root = re.search(r'<worksheet\b[^>]*>', sheet_xml, re.DOTALL)
    if not root:
        errors.append("sheet1.xml has no worksheet root element")
        return
    declarations = namespace_declarations(root.group(0))
    ignorable = re.search(r'\bmc:Ignorable="([^"]+)"', root.group(0))
    if ignorable and "mc" not in declarations:
        errors.append("mc:Ignorable is present but the mc namespace is not declared")
    if ignorable:
        for prefix in ignorable.group(1).split():
            if prefix not in declarations:
                errors.append(f"mc:Ignorable prefix {prefix!r} has no namespace declaration")


def extract_formula(sheet_xml: str, reference: str) -> str | None:
    cell = re.search(
        rf'<c\b(?=[^>]*\br="{reference}")[^>]*>(.*?)</c>', sheet_xml, re.DOTALL
    )
    if not cell:
        return None
    formula = re.search(r'<f(?:\s[^>]*)?>(.*?)</f>', cell.group(1), re.DOTALL)
    return formula.group(1) if formula else None


def extract_value(sheet_xml: str, reference: str) -> str | None:
    cell = re.search(
        rf'<c\b(?=[^>]*\br="{reference}")[^>]*>(.*?)</c>', sheet_xml, re.DOTALL
    )
    if not cell:
        return None
    value = re.search(r'<v>(.*?)</v>', cell.group(1), re.DOTALL)
    return html.unescape(value.group(1)) if value else None


def expected_period(start: date, end: date) -> str:
    return f"({start.year}.{start.month:02d}.{start.day:02d} ~ {end.year}.{end.month:02d}.{end.day:02d})"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def parse_shared_strings(shared_xml: bytes) -> list[str]:
    root = ET.fromstring(shared_xml)
    return [
        "".join(node.text or "" for node in item.iter() if local_name(node.tag) == "t")
        for item in root
        if local_name(item.tag) == "si"
    ]


def inspect_sheet_structure(sheet_xml: bytes, errors: list[str]) -> tuple[ET.Element, dict[str, ET.Element]]:
    root = ET.fromstring(sheet_xml)
    cells: dict[str, ET.Element] = {}
    row_numbers: set[str] = set()
    for row in (node for node in root.iter() if local_name(node.tag) == "row"):
        row_number = row.get("r")
        if not row_number:
            errors.append("worksheet contains a row without an r attribute")
            continue
        if row_number in row_numbers:
            errors.append(f"duplicate row number: {row_number}")
        row_numbers.add(row_number)
        for cell in (node for node in row if local_name(node.tag) == "c"):
            reference = cell.get("r")
            if not reference:
                errors.append(f"row {row_number} contains a cell without an r attribute")
                continue
            if reference in cells:
                errors.append(f"duplicate cell reference: {reference}")
            cells[reference] = cell
            match = re.search(r"(\d+)$", reference)
            if not match or match.group(1) != row_number:
                errors.append(f"cell {reference} is stored in row {row_number}")
    return root, cells


def cell_text(cell: ET.Element | None, shared_strings: list[str], errors: list[str], reference: str) -> str:
    if cell is None:
        return ""
    value = next((node.text or "" for node in cell if local_name(node.tag) == "v"), "")
    if cell.get("t") != "s" or not value:
        return value
    try:
        return shared_strings[int(value)]
    except (ValueError, IndexError):
        errors.append(f"{reference} has an invalid shared-string index: {value!r}")
        return ""


def load_expected_lines(path: Path | None, label: str, errors: list[str]) -> list[str] | None:
    if path is None:
        return None
    try:
        value = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        errors.append(f"cannot read {label} JSON: {exc}")
        return None
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        errors.append(f"{label} JSON must be an array of strings")
        return None
    if len(value) > 20:
        errors.append(f"{label} JSON has {len(value)} rows; maximum is 20")
        return None
    return value


def validate_report_lines(cells: dict[str, ET.Element], shared_strings: list[str], column: str,
                          expected: list[str], label: str, errors: list[str]) -> None:
    padded = expected + [""] * (25 - len(expected))
    for offset, expected_text in enumerate(padded):
        reference = f"{column}{8 + offset}"
        actual = cell_text(cells.get(reference), shared_strings, errors, reference)
        if actual != expected_text:
            errors.append(f"{label} {reference} differs: expected={expected_text!r}, actual={actual!r}")


def report_column_widths(root: ET.Element) -> tuple[float, float]:
    widths: dict[int, float] = {}
    for column in (node for node in root.iter() if local_name(node.tag) == "col"):
        start = int(column.get("min", "0"))
        end = int(column.get("max", "0"))
        width = float(column.get("width", "0"))
        for number in range(max(start, 3), min(end, 26) + 1):
            widths[number] = width
    missing = [number for number in range(3, 27) if number not in widths]
    if missing:
        raise ValueError(f"missing width definitions for columns: {missing}")
    return sum(widths[number] for number in range(3, 16)), sum(widths[number] for number in range(16, 27))


def main() -> None:
    args = parse_args()
    errors: list[str] = []
    template_zip, template = read_zip(args.template)
    output_zip, output = read_zip(args.output)
    try:
        if set(template) != set(output):
            missing = sorted(set(template) - set(output))
            extra = sorted(set(output) - set(template))
            errors.append(f"ZIP member set differs; missing={missing}, extra={extra}")

        for name in sorted(set(template) & set(output)):
            if name not in ALLOWED_CHANGES and template[name] != output[name]:
                errors.append(f"unexpected package change: {name}")

        for name in ("xl/calcChain.xml", "xl/printerSettings/printerSettings1.bin"):
            if name in template and name in output and template[name] != output[name]:
                errors.append(f"preserved part changed: {name}")
            if (name in template) != (name in output):
                errors.append(f"preserved part presence changed: {name}")

        for required in (SHEET, SHARED_STRINGS, WORKBOOK):
            if required not in output:
                errors.append(f"missing required part: {required}")

        output_root = None
        output_cells: dict[str, ET.Element] = {}
        if SHEET in output:
            sheet_text = output[SHEET].decode("utf-8")
            validate_ignorable(sheet_text, errors)
            try:
                output_root, output_cells = inspect_sheet_structure(output[SHEET], errors)
            except ET.ParseError as exc:
                errors.append(f"sheet1.xml is not well-formed XML: {exc}")
            template_sheet = template[SHEET].decode("utf-8")
            for reference in ("H4", "T4"):
                if extract_formula(template_sheet, reference) != extract_formula(sheet_text, reference):
                    errors.append(f"formula changed in {reference}")
            anchor_serial = extract_value(sheet_text, "X3")
            try:
                anchor = date(1899, 12, 30) + timedelta(days=int(anchor_serial or ""))
            except ValueError:
                errors.append("X3 does not contain a valid Excel date serial")
            else:
                expected = {
                    "H4": expected_period(anchor - timedelta(days=6), anchor),
                    "T4": expected_period(anchor + timedelta(days=1), anchor + timedelta(days=8)),
                }
                for reference, expected_value in expected.items():
                    if extract_value(sheet_text, reference) != expected_value:
                        errors.append(f"cached date text in {reference} is incorrect")

        current_lines = load_expected_lines(args.current_json, "current", errors)
        next_lines = load_expected_lines(args.next_json, "next", errors)
        if (args.current_json is None) != (args.next_json is None):
            errors.append("--current-json and --next-json must be supplied together")
        if current_lines is not None and next_lines is not None and SHARED_STRINGS in output:
            try:
                shared_strings = parse_shared_strings(output[SHARED_STRINGS])
            except ET.ParseError as exc:
                errors.append(f"sharedStrings.xml is not well-formed XML: {exc}")
            else:
                validate_report_lines(output_cells, shared_strings, "C", current_lines, "current", errors)
                validate_report_lines(output_cells, shared_strings, "P", next_lines, "next", errors)

        if args.require_balanced_columns and output_root is not None:
            try:
                output_left, output_right = report_column_widths(output_root)
                template_root = ET.fromstring(template[SHEET])
                template_left, template_right = report_column_widths(template_root)
            except (ET.ParseError, ValueError) as exc:
                errors.append(f"cannot validate report column widths: {exc}")
            else:
                if abs(output_left - output_right) > 0.001:
                    errors.append(
                        f"report columns are not balanced: current={output_left}, next={output_right}"
                    )
                if abs((output_left + output_right) - (template_left + template_right)) > 0.001:
                    errors.append("combined C:Z width differs from the template")

        if WORKBOOK in output:
            workbook_text = output[WORKBOOK].decode("utf-8")
            if REQUIRED_PRINT_AREA not in workbook_text:
                errors.append(f"print area {REQUIRED_PRINT_AREA} is missing")
            if REQUIRED_PRINT_TITLES not in workbook_text:
                errors.append(f"print title rows {REQUIRED_PRINT_TITLES} are missing")

        summary = {
            "valid": not errors,
            "template": str(args.template.resolve()),
            "output": str(args.output.resolve()),
            "allowed_changed_parts": sorted(
                name for name in ALLOWED_CHANGES if template.get(name) != output.get(name)
            ),
            "errors": errors,
        }
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        if errors:
            raise SystemExit(1)
    finally:
        template_zip.close()
        output_zip.close()


if __name__ == "__main__":
    main()
