"""Validate the structural deliverables for Phase 1."""

from __future__ import annotations

import os
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXPECTED_SHEETS = [
    "START",
    "QUOTE",
    "QUOTE_LINES",
    "COSTING",
    "OFFER",
    "MASTERS",
    "SETTINGS",
    "CUSTOMERS",
    "PLANTS",
    "DISTANCE_CACHE",
    "QUOTE_LOG",
]
EXPECTED_TABLES = [
    "tblProducts",
    "tblSizes",
    "tblThickness",
    "tblGrades",
    "tblExternalCoatings",
    "tblInternalCoatings",
    "tblPlants",
    "tblCustomers",
    "tblPaymentTerms",
    "tblSettings",
    "tblQuoteLines",
    "tblDistanceCache",
    "tblQuoteLog",
]
EXPECTED_START = {
    "A1": "Mega Pipes Quotation System",
    "A2": "Development Workbook",
    "A3": "Version: DEV",
    "A5": "Workbook architecture initialized.",
    "A6": "This is the development workbook. Do not use for production quotations.",
    "A8": "System Status",
    "A9": "Environment",
    "B9": "DEV",
    "A10": "Workbook Version",
    "B10": "DEV-0.2",
    "A11": "Architecture Status",
    "B11": "PHASE 5 - P1 VERIFIED",
    "A12": "VBA Status",
    "B12": "Source imported; compile and P1 runtime verified",
    "A13": "Calculation Engine",
    "B13": "Python reference layer pending VBA integration",
    "A14": "ORS",
    "B14": "Not implemented",
    "A15": "Production Status",
    "B15": "NOT READY",
}


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return ["".join(node.text or "" for node in si.iter(f"{{{NS['m']}}}t")) for si in root]


def cell_values(archive: zipfile.ZipFile, sheet_path: str) -> dict[str, str]:
    root = ET.fromstring(archive.read(sheet_path))
    strings = shared_strings(archive) if "xl/sharedStrings.xml" in archive.namelist() else []
    values: dict[str, str] = {}
    for cell in root.findall(".//m:c", NS):
        value = cell.find("m:v", NS)
        if value is None:
            continue
        text = value.text or ""
        if cell.attrib.get("t") == "s":
            text = strings[int(text)]
        values[cell.attrib["r"]] = text
    return values


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    xlsx_path = root / "DEV" / "Mega_Pipes_Quotation_DEV.xlsx"
    xlsm_path = root / "DEV" / "Mega_Pipes_Quotation_DEV.xlsm"
    workbook_path = xlsx_path if xlsx_path.is_file() else xlsm_path
    if not workbook_path.is_file() or workbook_path.stat().st_size == 0:
        fail("development workbook is missing or empty")

    with zipfile.ZipFile(workbook_path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = [
            sheet.attrib["name"]
            for sheet in workbook.find("m:sheets", NS)
        ]
        if sheets != EXPECTED_SHEETS:
            fail(f"sheet names/order mismatch: {sheets}")

        table_names: list[str] = []
        for name in archive.namelist():
            if not name.startswith("xl/tables/table") or not name.endswith(".xml"):
                continue
            table = ET.fromstring(archive.read(name))
            table_names.append(table.attrib["name"])
            columns = [
                column.attrib["name"]
                for column in table.findall("m:tableColumns/m:tableColumn", NS)
            ]
            if len(columns) != len(set(columns)):
                fail(f"duplicate headers in {table.attrib['name']}")
        if sorted(table_names) != sorted(EXPECTED_TABLES):
            fail(f"table names mismatch: {sorted(table_names)}")
        if len(table_names) != len(set(table_names)):
            fail("table names are not unique")

        start = cell_values(archive, "xl/worksheets/sheet1.xml")
        for address, expected in EXPECTED_START.items():
            if start.get(address) != expected:
                fail(f"START!{address} expected {expected!r}, got {start.get(address)!r}")

        formula_cells = 0
        for name in archive.namelist():
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                formula_cells += len(
                    ET.fromstring(archive.read(name)).findall(".//m:f", NS)
                )
        if formula_cells:
            fail(f"{formula_cells} formula cells found")

    print("PASS: Phase 1 workbook architecture checks completed")
    print(f"Workbook: {workbook_path}")
    print(f"Sheets: {len(EXPECTED_SHEETS)}")
    print(f"Tables: {len(EXPECTED_TABLES)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
