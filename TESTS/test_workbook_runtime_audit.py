from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET


NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "DEV" / "Mega_Pipes_Quotation_DEV.xlsm"
EXPECTED_SHEETS = [
    "START", "QUOTE", "QUOTE_LINES", "COSTING", "OFFER", "MASTERS",
    "SETTINGS", "CUSTOMERS", "PLANTS", "DISTANCE_CACHE", "QUOTE_LOG",
]
EXPECTED_TABLES = {
    "tblProducts", "tblSizes", "tblThickness", "tblGrades",
    "tblExternalCoatings", "tblInternalCoatings", "tblPlants",
    "tblCustomers", "tblPaymentTerms", "tblSettings", "tblQuoteLines",
    "tblDistanceCache", "tblQuoteLog",
}


def test_macro_enabled_package_and_structure():
    with zipfile.ZipFile(WORKBOOK) as archive:
        assert "xl/vbaProject.bin" in archive.namelist()
        content_types = archive.read("[Content_Types].xml").decode("utf-8")
        assert "application/vnd.ms-excel.sheet.macroEnabled.main+xml" in content_types
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = [sheet.attrib["name"] for sheet in workbook.find("m:sheets", NS)]
        assert sheets == EXPECTED_SHEETS
        table_names = {
            ET.fromstring(archive.read(name)).attrib["name"]
            for name in archive.namelist()
            if name.startswith("xl/tables/table") and name.endswith(".xml")
        }
        assert table_names == EXPECTED_TABLES


def test_customer_master_has_one_nonblank_customer():
    with zipfile.ZipFile(WORKBOOK) as archive:
        shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        strings = [
            "".join(t.text or "" for t in item.iter(f"{{{NS['m']}}}t"))
            for item in shared
        ]
        sheet = ET.fromstring(archive.read("xl/worksheets/sheet8.xml"))
        rows = sheet.findall(".//m:sheetData/m:row", NS)
        customer_ids = []
        for row in rows[1:]:
            values = []
            for cell in row.findall("m:c", NS):
                value = cell.find("m:v", NS)
                if value is None:
                    values.append("")
                elif cell.attrib.get("t") == "s":
                    values.append(strings[int(value.text)])
                else:
                    values.append(value.text or "")
            if values and values[0].strip():
                customer_ids.append(values[0])
        assert customer_ids == ["C001"]
