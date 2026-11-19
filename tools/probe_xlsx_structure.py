#!/usr/bin/env python3
"""Report anonymized worksheet shape without printing workbook values."""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MAX_UNCOMPRESSED_BYTES = 256 * 1024 * 1024


def _column_number(cell_ref: str) -> int:
    letters = re.match(r"([A-Z]+)", cell_ref.upper())
    if not letters:
        return 0
    result = 0
    for char in letters.group(1):
        result = result * 26 + ord(char) - ord("A") + 1
    return result


def _row_number(cell_ref: str) -> int:
    digits = re.search(r"(\d+)$", cell_ref)
    return int(digits.group(1)) if digits else 0


def _shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(node.itertext()) for node in root.findall(f"{{{MAIN_NS}}}si")]


def _cell_payload(cell: ET.Element, shared: list[str]) -> tuple[bool, str | None, bool]:
    """Return nonempty, text value when known, and formula-present."""
    formula = cell.find(f"{{{MAIN_NS}}}f") is not None
    kind = cell.attrib.get("t")
    value_node = cell.find(f"{{{MAIN_NS}}}v")
    inline_node = cell.find(f"{{{MAIN_NS}}}is")

    if kind == "s" and value_node is not None and value_node.text is not None:
        try:
            value = shared[int(value_node.text)]
        except (ValueError, IndexError):
            return True, None, formula
        return formula or value != "", value, formula
    if kind == "inlineStr" and inline_node is not None:
        value = "".join(inline_node.itertext())
        return formula or value != "", value, formula
    if kind == "str":
        value = value_node.text if value_node is not None else ""
        return formula or value != "", value, formula
    if value_node is not None and value_node.text not in (None, ""):
        return True, None, formula
    return formula, None, formula


def summarize_sheet(sheet_xml: bytes, shared: list[str], sheet_index: int) -> dict[str, object]:
    root = ET.fromstring(sheet_xml)
    dimension = root.find(f"{{{MAIN_NS}}}dimension")
    dimension_ref = dimension.attrib.get("ref", "") if dimension is not None else ""
    dimension_last_row = max((_row_number(part) for part in dimension_ref.split(":")), default=0)

    cells_by_row: dict[int, dict[int, tuple[bool, str | None, bool]]] = {}
    max_column = 0
    for cell in root.findall(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row/{{{MAIN_NS}}}c"):
        ref = cell.attrib.get("r", "")
        row = _row_number(ref)
        col = _column_number(ref)
        if not row or not col:
            continue
        max_column = max(max_column, col)
        cells_by_row.setdefault(row, {})[col] = _cell_payload(cell, shared)

    header = cells_by_row.get(1, {})
    product_columns: list[int] = []
    amg_columns: list[int] = []
    invalid_headers = 0
    for col, (nonempty, text, formula) in header.items():
        if not nonempty:
            continue
        if formula or text is None:
            invalid_headers += 1
        elif text.upper().startswith("AMG"):
            amg_columns.append(col)
        else:
            product_columns.append(col)

    anchors = [col for col in product_columns if col <= 5]
    data_rows: dict[int, tuple[bool, bool]] = {}
    last_anchor_by_col: dict[int, int] = {col: 0 for col in anchors}
    last_amg_by_col: dict[int, int] = {col: 0 for col in amg_columns}
    amg_value_kinds: Counter[str] = Counter()
    formula_cells_in_amg = 0

    for row, cells in cells_by_row.items():
        if row < 2:
            continue
        anchor_nonempty = False
        amg_nonempty = False
        for col in anchors:
            payload = cells.get(col, (False, None, False))
            if payload[0]:
                anchor_nonempty = True
                last_anchor_by_col[col] = max(last_anchor_by_col[col], row)
        for col in amg_columns:
            payload = cells.get(col, (False, None, False))
            if payload[0]:
                amg_nonempty = True
                last_amg_by_col[col] = max(last_amg_by_col[col], row)
                if payload[2]:
                    formula_cells_in_amg += 1
                if payload[1] is not None:
                    amg_value_kinds["text"] += 1
                else:
                    amg_value_kinds["non_text_or_formula"] += 1
        if anchor_nonempty or amg_nonempty:
            data_rows[row] = (anchor_nonempty, amg_nonempty)

    last_anchor = max(last_anchor_by_col.values(), default=0)
    last_amg = max(last_amg_by_col.values(), default=0)
    relevant_rows = sorted(data_rows)
    first_relevant = relevant_rows[0] if relevant_rows else 0
    last_relevant = relevant_rows[-1] if relevant_rows else 0
    internal_blank_rows = 0
    if first_relevant and last_relevant > first_relevant:
        internal_blank_rows = sum(
            1 for row in range(first_relevant + 1, last_relevant) if row not in data_rows
        )

    return {
        "sheet_index": sheet_index,
        "has_amg": bool(amg_columns),
        "amg_column_count": len(amg_columns),
        "amg_columns": sorted(amg_columns),
        "product_columns_a_to_e": sorted(anchors),
        "product_columns_after_e": sum(1 for col in product_columns if col > 5),
        "invalid_nonempty_header_count": invalid_headers,
        "last_anchor_row": last_anchor,
        "last_amg_row": last_amg,
        "amg_rows_without_any_anchor": sum(
            1 for anchor_nonempty, amg_nonempty in data_rows.values() if amg_nonempty and not anchor_nonempty
        ),
        "internal_fully_empty_rows": internal_blank_rows,
        "amg_value_kinds": dict(sorted(amg_value_kinds.items())),
        "formula_cells_in_amg": formula_cells_in_amg,
        "xml_dimension_last_row": dimension_last_row,
        "xml_max_column": max_column,
    }


def _worksheet_paths(archive: zipfile.ZipFile) -> list[str]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    paths = []
    for sheet in workbook.findall(f".//{{{MAIN_NS}}}sheet"):
        target = targets[sheet.attrib[f"{{{DOC_REL_NS}}}id"]]
        if target.startswith("/"):
            paths.append(posixpath.normpath(target.lstrip("/")))
        elif target.startswith("xl/"):
            paths.append(posixpath.normpath(target))
        else:
            paths.append(posixpath.normpath(posixpath.join("xl", target)))
    return paths


def inspect_workbook(path: Path) -> dict[str, object]:
    with zipfile.ZipFile(path) as archive:
        total_size = sum(item.file_size for item in archive.infolist())
        if total_size > MAX_UNCOMPRESSED_BYTES:
            raise ValueError("workbook exceeds the probe's uncompressed-size limit")
        shared = _shared_strings(archive)
        sheets = [
            summarize_sheet(archive.read(sheet_path), shared, index)
            for index, sheet_path in enumerate(_worksheet_paths(archive), start=1)
        ]
    return {
        "worksheet_count": len(sheets),
        "worksheets_with_amg": sum(1 for sheet in sheets if sheet["has_amg"]),
        "worksheets_without_a_to_e_anchor": sum(
            1 for sheet in sheets if sheet["has_amg"] and not sheet["product_columns_a_to_e"]
        ),
        "sheets": sheets,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbooks", nargs="+", type=Path)
    args = parser.parse_args()
    reports = []
    for index, path in enumerate(args.workbooks, start=1):
        reports.append({"sample_index": index, **inspect_workbook(path)})
    print(json.dumps(reports, indent=2, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
