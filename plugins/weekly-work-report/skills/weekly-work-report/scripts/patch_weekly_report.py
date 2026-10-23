#!/usr/bin/env python3
"""Patch a weekly report without reserializing the worksheet XML."""

from __future__ import annotations

import argparse
import html
import json
import re
import shutil
import tempfile
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path


SHEET = "xl/worksheets/sheet1.xml"
SHARED_STRINGS = "xl/sharedStrings.xml"
MAX_LINES = 20
CONTENT_END_ROW = 32


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--anchor-date", required=True, type=date.fromisoformat)
    parser.add_argument("--current-json", required=True, type=Path)
    parser.add_argument("--next-json", required=True, type=Path)
    parser.add_argument("--balance-columns", action="store_true")
    return parser.parse_args()


def load_lines(path: Path) -> list[str]:
    value = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        raise ValueError(f"{path} must contain a JSON array of strings")
    if len(value) > MAX_LINES:
        raise ValueError(f"{path} contains {len(value)} rows; maximum is {MAX_LINES}")
    return value


def excel_serial(value: date) -> int:
    return (value - date(1899, 12, 30)).days


def column_number(reference: str) -> int:
    letters = re.match(r"[A-Z]+", reference)
    if not letters:
        raise ValueError(f"invalid cell reference: {reference}")
    result = 0
    for char in letters.group(0):
        result = result * 26 + ord(char) - ord("A") + 1
    return result


def cell_pattern(reference: str) -> re.Pattern[str]:
    return re.compile(
        rf'<c\b(?![^>]*/>)(?=[^>]*\br="{re.escape(reference)}")(?P<attrs>[^>]*)>'
        rf'(?P<body>.*?)</c>|<c\b(?=[^>]*\br="{re.escape(reference)}")(?P<selfattrs>[^>]*)/>',
        re.DOTALL,
    )


def get_cell_style(sheet_xml: str, reference: str, fallback: str | None = None) -> str | None:
    match = cell_pattern(reference).search(sheet_xml)
    if match:
        attrs = match.group("attrs") if match.group("attrs") is not None else match.group("selfattrs")
        style = re.search(r'\bs="(\d+)"', attrs or "")
        if style:
            return style.group(1)
    return fallback


def replace_existing_cell(sheet_xml: str, reference: str, replacement: str) -> tuple[str, bool]:
    pattern = cell_pattern(reference)
    updated, count = pattern.subn(replacement, sheet_xml, count=1)
    return updated, count == 1


def insert_cell(sheet_xml: str, reference: str, replacement: str) -> str:
    row_number = int(re.search(r"\d+$", reference).group(0))
    row_pattern = re.compile(
        rf'(<row\b(?=[^>]*\br="{row_number}")[^>]*>)(.*?)(</row>)', re.DOTALL
    )
    row_match = row_pattern.search(sheet_xml)
    if not row_match:
        sheet_data = re.search(r'(<sheetData>)(.*?)(</sheetData>)', sheet_xml, re.DOTALL)
        if not sheet_data:
            raise ValueError("worksheet has no sheetData element")
        body = sheet_data.group(2)
        insertion = len(body)
        for match in re.finditer(r'<row\b[^>]*\br="(\d+)"[^>]*(?:/>|>.*?</row>)', body, re.DOTALL):
            if int(match.group(1)) > row_number:
                insertion = match.start()
                break
        new_body = body[:insertion] + f'<row r="{row_number}">{replacement}</row>' + body[insertion:]
        return sheet_xml[: sheet_data.start(2)] + new_body + sheet_xml[sheet_data.end(2) :]
    body = row_match.group(2)
    target_column = column_number(reference)
    insertion = len(body)
    for match in re.finditer(r'<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:/>|>.*?</c>)', body, re.DOTALL):
        if column_number(match.group(1)) > target_column:
            insertion = match.start()
            break
    new_body = body[:insertion] + replacement + body[insertion:]
    return sheet_xml[: row_match.start(2)] + new_body + sheet_xml[row_match.end(2) :]


def make_cell(reference: str, style: str | None, shared_index: int | None, numeric: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    if numeric is not None:
        return f'<c r="{reference}"{style_attr}><v>{numeric}</v></c>'
    if shared_index is None:
        return f'<c r="{reference}"{style_attr}/>'
    return f'<c r="{reference}"{style_attr} t="s"><v>{shared_index}</v></c>'


def set_cell(sheet_xml: str, reference: str, replacement: str) -> str:
    updated, found = replace_existing_cell(sheet_xml, reference, replacement)
    return updated if found else insert_cell(sheet_xml, reference, replacement)


def append_shared_strings(shared_xml: str, values: list[str]) -> tuple[str, dict[str, int]]:
    existing = []
    for item in re.findall(r'<si(?:\s[^>]*)?>(.*?)</si>', shared_xml, re.DOTALL):
        parts = re.findall(r'<t(?:\s[^>]*)?>(.*?)</t>', item, re.DOTALL)
        existing.append("".join(html.unescape(part) for part in parts))
    indices = {value: index for index, value in enumerate(existing)}
    additions = []
    for value in values:
        if value and value not in indices:
            indices[value] = len(existing) + len(additions)
            escaped = html.escape(value, quote=False)
            additions.append(f'<si><t xml:space="preserve">{escaped}</t></si>')
    if additions:
        shared_xml = shared_xml.replace("</sst>", "".join(additions) + "</sst>", 1)
    unique_count = len(existing) + len(additions)
    shared_xml = re.sub(r'\buniqueCount="\d+"', f'uniqueCount="{unique_count}"', shared_xml, count=1)
    return shared_xml, indices


def update_shared_string_count(shared_xml: str, sheet_xml: str) -> str:
    reference_count = len(re.findall(r'<c\b[^>]*\bt="s"', sheet_xml))
    return re.sub(r'\bcount="\d+"', f'count="{reference_count}"', shared_xml, count=1)


def patch_formula_cache(sheet_xml: str, reference: str, cached_text: str) -> str:
    pattern = cell_pattern(reference)
    match = pattern.search(sheet_xml)
    if not match or match.group("body") is None:
        raise ValueError(f"formula cell {reference} is missing")
    attrs = match.group("attrs") or ""
    attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
    attrs += ' t="str"'
    body = match.group("body")
    escaped = html.escape(cached_text, quote=False)
    if re.search(r'<v>.*?</v>', body, re.DOTALL):
        body = re.sub(r'<v>.*?</v>', f'<v>{escaped}</v>', body, count=1, flags=re.DOTALL)
    else:
        body += f'<v>{escaped}</v>'
    replacement = f"<c{attrs}>{body}</c>"
    return sheet_xml[: match.start()] + replacement + sheet_xml[match.end() :]


def period_text(start: date, end: date) -> str:
    return f"({start.year}.{start.month:02d}.{start.day:02d} ~ {end.year}.{end.month:02d}.{end.day:02d})"


def format_width(value: float) -> str:
    return format(value, ".12g")


def balance_report_columns(sheet_xml: str) -> str:
    combined = re.compile(
        r'<col min="3" max="26" width="(?P<width>[^"]+)"(?P<attrs>[^>]*)/>'
    )
    match = combined.search(sheet_xml)
    if match:
        total_width = 24 * float(match.group("width"))
        attrs = match.group("attrs")
        replacement = (
            f'<col min="3" max="15" width="{format_width(total_width / 26)}"{attrs}/>'
            f'<col min="16" max="26" width="{format_width(total_width / 22)}"{attrs}/>'
        )
        return sheet_xml[:match.start()] + replacement + sheet_xml[match.end():]

    split = re.compile(
        r'<col min="3" max="15" width="(?P<left>[^"]+)"(?P<left_attrs>[^>]*)/>'
        r'<col min="16" max="26" width="(?P<right>[^"]+)"(?P<right_attrs>[^>]*)/>'
    )
    match = split.search(sheet_xml)
    if not match:
        raise ValueError("C:Z column definitions are not in a supported weekly-report shape")
    total_width = 13 * float(match.group("left")) + 11 * float(match.group("right"))
    replacement = (
        f'<col min="3" max="15" width="{format_width(total_width / 26)}"{match.group("left_attrs")}/>'
        f'<col min="16" max="26" width="{format_width(total_width / 22)}"{match.group("right_attrs")}/>'
    )
    return sheet_xml[:match.start()] + replacement + sheet_xml[match.end():]


def patch_sheet(sheet_xml: str, shared_indices: dict[str, int], anchor: date,
                current: list[str], next_week: list[str]) -> str:
    heading_styles = {
        "C": get_cell_style(sheet_xml, "C8"),
        "P": get_cell_style(sheet_xml, "P8"),
    }
    body_styles = {
        "C": get_cell_style(sheet_xml, "C9", heading_styles["C"]),
        "P": get_cell_style(sheet_xml, "P9", heading_styles["P"]),
    }
    x3_style = get_cell_style(sheet_xml, "X3")
    sheet_xml = set_cell(sheet_xml, "X3", make_cell("X3", x3_style, None, excel_serial(anchor)))

    for column, lines in (("C", current), ("P", next_week)):
        for row in range(8, CONTENT_END_ROW + 1):
            reference = f"{column}{row}"
            existing_style = get_cell_style(sheet_xml, reference, body_styles[column])
            sheet_xml = set_cell(sheet_xml, reference, make_cell(reference, existing_style, None))
        for offset, text in enumerate(lines):
            if not text:
                continue
            row = 8 + offset
            reference = f"{column}{row}"
            is_heading = text.lstrip().startswith("▶")
            style = heading_styles[column] if is_heading else body_styles[column]
            sheet_xml = set_cell(
                sheet_xml,
                reference,
                make_cell(reference, style, shared_indices[text]),
            )

    sheet_xml = patch_formula_cache(
        sheet_xml, "H4", period_text(anchor - timedelta(days=6), anchor)
    )
    sheet_xml = patch_formula_cache(
        sheet_xml, "T4", period_text(anchor + timedelta(days=1), anchor + timedelta(days=8))
    )
    return sheet_xml


def copy_and_patch(template: Path, output: Path, anchor: date,
                   current: list[str], next_week: list[str], balance_columns: bool = False) -> None:
    if template.suffix.lower() != ".xlsx" or output.suffix.lower() != ".xlsx":
        raise ValueError("template and output must be .xlsx files")
    if not template.is_file():
        raise FileNotFoundError(template)
    output.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(template, "r") as source:
        names = set(source.namelist())
        for required in (SHEET, SHARED_STRINGS):
            if required not in names:
                raise ValueError(f"template is missing {required}")
        sheet_xml = source.read(SHEET).decode("utf-8")
        shared_xml = source.read(SHARED_STRINGS).decode("utf-8")
        shared_xml, indices = append_shared_strings(shared_xml, current + next_week)
        sheet_xml = patch_sheet(sheet_xml, indices, anchor, current, next_week)
        if balance_columns:
            sheet_xml = balance_report_columns(sheet_xml)
        shared_xml = update_shared_string_count(shared_xml, sheet_xml)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=output.parent) as handle:
            temporary = Path(handle.name)
        try:
            with zipfile.ZipFile(temporary, "w") as target:
                for item in source.infolist():
                    data = source.read(item.filename)
                    if item.filename == SHEET:
                        data = sheet_xml.encode("utf-8")
                    elif item.filename == SHARED_STRINGS:
                        data = shared_xml.encode("utf-8")
                    target.writestr(item, data)
            shutil.move(str(temporary), output)
        finally:
            temporary.unlink(missing_ok=True)


def main() -> None:
    args = parse_args()
    current = load_lines(args.current_json)
    next_week = load_lines(args.next_json)
    copy_and_patch(
        args.template,
        args.output,
        args.anchor_date,
        current,
        next_week,
        balance_columns=args.balance_columns,
    )
    print(args.output)


if __name__ == "__main__":
    main()
