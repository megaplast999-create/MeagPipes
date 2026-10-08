"""Reproducibly build the DEV workbook structure and documented seed data."""

from __future__ import annotations

from pathlib import Path

import xlsxwriter


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "DEV" / "Mega_Pipes_Quotation_DEV.xlsx"

SHEETS = [
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


def table(worksheet, row, col, name, headers, data, widths=None):
    worksheet.add_table(
        row,
        col,
        row + len(data) + 1,
        col + len(headers) - 1,
        {
            "name": name,
            "style": "Table Style Medium 2",
            "columns": [{"header": header} for header in headers],
            "data": data or [[""] * len(headers)],
        },
    )
    if widths:
        for offset, width in enumerate(widths):
            worksheet.set_column(col + offset, col + offset, width)


def build() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(OUTPUT))
    workbook.set_properties(
        {
            "title": "Mega Pipes Quotation System - Development Workbook",
            "subject": "Master data and quotation architecture",
            "author": "Mega Pipes",
        }
    )
    navy = "#1F4E78"
    white = "#FFFFFF"
    light = "#F3F6F9"
    title = workbook.add_format(
        {"bold": True, "font_size": 16, "font_color": white, "bg_color": navy}
    )
    section = workbook.add_format(
        {"bold": True, "font_color": white, "bg_color": navy, "border": 1}
    )
    label = workbook.add_format({"bold": True, "bg_color": light, "border": 1})
    value = workbook.add_format({"border": 1})
    header = workbook.add_format(
        {"bold": True, "font_color": white, "bg_color": navy, "border": 1}
    )
    ws = {name: workbook.add_worksheet(name) for name in SHEETS}

    start = ws["START"]
    start.set_column("A:A", 28)
    start.set_column("B:B", 54)
    start.write("A1", "Mega Pipes Quotation System", title)
    start.write("A2", "Development Workbook")
    start.write("A3", "Version: DEV")
    start.write("A5", "Workbook architecture initialized.")
    start.write("A6", "This is the development workbook. Do not use for production quotations.")
    start.write("A8", "System Status", section)
    status = [
        ("Environment", "DEV"),
        ("Workbook Version", "DEV-0.2"),
        ("Architecture Status", "PHASE 4 - SEEDED"),
        ("VBA Status", "Source not embedded; import pending"),
        ("Calculation Engine", "Python reference layer pending VBA integration"),
        ("ORS", "Not implemented"),
        ("Production Status", "NOT READY"),
    ]
    for row, (key, val) in enumerate(status, 8):
        start.write(row, 0, key, label)
        start.write(row, 1, val, value)
    start.freeze_panes(8, 0)

    quote = ws["QUOTE"]
    quote.set_column("A:A", 30)
    quote.set_column("B:B", 34)
    quote.write("A1", "MEGA PIPES QUOTATION", title)
    quote.write("A3", "QUOTE HEADER", section)
    for row, field in enumerate(
        [
            "Quote Number",
            "Quote Date",
            "Enquiry Reference",
            "Enquiry Date",
            "Customer",
            "Contact Person",
            "Project / Tender",
            "Delivery Place",
            "Due Date",
            "Sales Person",
            "Plant",
            "Payment Terms",
            "Validity",
            "Price Basis",
        ],
        3,
    ):
        quote.write(row, 0, field, label)
        quote.write_blank(row, 1, None, value)
    quote.write("A20", "PIPE REQUIREMENT", section)
    for row, field in enumerate(
        [
            "Product",
            "Family",
            "Basis",
            "Size",
            "Thickness",
            "Grade",
            "External Coating",
            "Internal Coating",
            "Quantity",
            "Quantity Unit",
            "Pipe Length",
        ],
        20,
    ):
        quote.write(row, 0, field, label)
        quote.write_blank(row, 1, None, value)

    quote_line_headers = [
        "LineID", "QuoteID", "ProductID", "Family", "Basis", "SizeMM",
        "ThicknessMM", "GradeID", "CoatingExternalID", "CoatingInternalID",
        "Quantity", "QuantityUnit", "PipeLengthM", "CalculatedODMM",
        "CalculatedIDMM", "MassKgPerM", "Metres", "PipeCount", "WeightMT",
        "UnitPrice", "PriceUnit", "LineTotal", "PlantID", "Status",
    ]
    table(ws["QUOTE_LINES"], 2, 0, "tblQuoteLines", quote_line_headers, [[""] * 24], [16] * 24)

    costing = ws["COSTING"]
    costing.set_column("A:A", 30)
    costing.write("A1", "COSTING", title)
    for row, text in [(3, "Quote Information"), (10, "Material Cost"),
                      (20, "External Coating Cost"), (30, "Internal Coating Cost"),
                      (40, "Freight Cost"), (50, "Financial Cost"),
                      (60, "Margin"), (70, "Selling Price")]:
        costing.write(row - 1, 0, text, section)

    offer = ws["OFFER"]
    offer.set_column("A:A", 34)
    offer.write("A1", "OFFER", title)
    for row, text in enumerate(
        ["Offer Header", "Customer Information", "Manufacturer Information",
         "Pipe Schedule", "Pricing", "GST", "Commercial Terms",
         "Technical Scope", "Delivery", "Payment Terms", "Validity",
         "Other Terms", "Signatory"],
        3,
    ):
        offer.write(row, 0, text, section)

    masters = ws["MASTERS"]
    masters.set_column("A:A", 26)
    masters.set_column("B:B", 28)
    masters.set_column("C:C", 20)
    masters.set_column("D:D", 40)
    masters.write("A1", "MASTER DATA", title)
    masters.write_row("A3", ["Master Name", "Table Name", "Primary Key", "Description"], header)
    registry = [
        ["Products", "tblProducts", "ProductID", "Product definitions"],
        ["Sizes", "tblSizes", "SizeID", "Pipe sizes and basis"],
        ["Thickness", "tblThickness", "ThicknessID", "Pipe thickness definitions"],
        ["Grades", "tblGrades", "GradeID", "Material specifications"],
        ["External Coatings", "tblExternalCoatings", "CoatingID", "External coating definitions"],
        ["Internal Coatings", "tblInternalCoatings", "CoatingID", "Internal coating definitions"],
        ["Plants", "tblPlants", "PlantID", "Manufacturing plants"],
        ["Customers", "tblCustomers", "CustomerID", "Customer master"],
        ["Payment Terms", "tblPaymentTerms", "PaymentTermID", "Payment terms"],
        ["Settings", "tblSettings", "SettingKey", "System settings"],
    ]
    for row, record in enumerate(registry, 3):
        masters.write_row(row, 0, record)

    products = [
        ["P01", "Bare HSAW", "HSAW", "No", "No", "", "", "", "Yes"],
        ["P02", "Coated HSAW", "HSAW", "Yes", "Optional", "", "", "", "Yes"],
        ["P03", "Bare HFW", "HFW", "No", "No", "", "", "", "Yes"],
        ["P04", "Coated HFW", "HFW", "Yes", "Optional", "", "", "", "Yes"],
    ]
    table(masters, 16, 0, "tblProducts",
          ["ProductID", "ProductName", "Family", "TopCoatingRequired",
           "BackCoatingOptional", "PrimaryColor", "SecondaryColor",
           "DefaultStandard", "Active"], products, [16, 24, 12, 20, 22, 16, 16, 20, 12])

    sizes = []
    for family, basis, values in [
        ("HSAW", "ID", [1226, 1300, 1350, 1400]),
        ("HSAW", "OD", [457, 508, 610, 762, 914, 1016, 1219, 1422, 1524, 1829, 2134, 2540]),
        ("HFW", "OD", [168.3, 219.1, 273.1, 323.9, 355.6, 406.4, 457.2, 508]),
    ]:
        for index, size in enumerate(values, 1):
            sizes.append([f"{family}-{basis}-{index:03d}", family, basis, size, "Yes", "No", ""])
    table(masters, 23, 0, "tblSizes",
          ["SizeID", "Family", "Basis", "SizeMM", "Active", "Pending", "Notes"], sizes, [20, 12, 12, 14, 12, 12, 28])

    thickness = []
    for family, values, max_value in [
        ("HSAW", [6.4, 7.1, 7.9, 8.4, 9.5, 11.1, 12.7, 15.9, 19.1, 25.4], 25.4),
        ("HFW", [3.5, 4.0, 4.5, 5.0, 5.6, 6.4, 7.1, 8.0, 8.8, 10.0, 12.5, 14.2, 16.0], 17.0),
    ]:
        for index, thickness_mm in enumerate(values, 1):
            thickness.append([f"{family}-{index:03d}", family, thickness_mm, 5.0 if family == "HSAW" else 3.5, max_value, "Yes"])
    table(masters, 52, 0, "tblThickness",
          ["ThicknessID", "Family", "ThicknessMM", "MinMM", "MaxMM", "Active"], thickness, [20, 12, 16, 12, 12, 12])

    hsaw_grades = ["IS 3589 Fe 330", "IS 3589 Fe 410", "IS 3589 Fe 450", "API 5L B", "X42 PSL1", "X52 PSL1", "X60 PSL1", "X65 PSL1", "X70 PSL1", "X52 PSL2", "X65 PSL2", "X70 PSL2", "X80M PSL2"]
    hfw_grades = ["IS 3589 Fe 330", "IS 3589 Fe 410", "IS 3589 Fe 450", "IS 1239 Part-1", "IS 1161 YST210", "IS 1161 YST240", "IS 1161 YST310", "API 5L A25", "API 5L B", "API 5L X42", "API 5L X52", "API 5L X60", "API 5L X65", "API 5L X70", "API 5L X80M PSL2", "API 5CT H40", "API 5CT J55", "API 5CT K55"]
    grades = [[f"{family}-{index:03d}", family, name, "Yes"] for family, names in [("HSAW", hsaw_grades), ("HFW", hfw_grades)] for index, name in enumerate(names, 1)]
    table(masters, 79, 0, "tblGrades", ["GradeID", "Family", "GradeName", "Active"], grades, [18, 12, 28, 12])

    ext = [
        ["EXT-PU-500", "PU", "500 micron", 500, "Yes", "Yes"],
        ["EXT-PU-1000", "PU", "1000 micron", "", "No", "Yes"],
        ["EXT-3LPE", "3LPE", "", "", "No", "Yes"],
        ["EXT-3LPP", "3LPP", "", "", "No", "Yes"],
        ["EXT-FBE", "FBE", "", "", "No", "Yes"],
        ["EXT-DFBE", "DFBE", "", "", "No", "Yes"],
        ["EXT-ZRE", "Zinc Rich Epoxy", "", "", "No", "Yes"],
    ]
    table(masters, 114, 0, "tblExternalCoatings",
          ["CoatingID", "CoatingName", "Thickness", "RateRsPerM2", "RateConfirmed", "Active"], ext, [20, 24, 18, 16, 18, 12])
    inte = [
        ["INT-FGE-406", "Food Grade Epoxy", "406 micron", 275, "Yes", "Yes"],
        ["INT-SFE", "Solvent-Free Epoxy", "", "", "No", "Yes"],
        ["INT-PU", "PU", "", "", "No", "Yes"],
        ["INT-ZRE", "Zinc Rich Epoxy", "", "", "No", "Yes"],
        ["INT-CL", "Cement Lining", "", "", "No", "Yes"],
    ]
    table(masters, 124, 0, "tblInternalCoatings",
          ["CoatingID", "CoatingName", "Thickness", "RateRsPerM2", "RateConfirmed", "Active"], inte, [20, 24, 18, 16, 18, 12])

    settings = ws["SETTINGS"]
    settings.write("A1", "SETTINGS", title)
    settings_headers = ["SettingKey", "SettingValue", "DataType", "Description", "DefaultValue", "Editable", "Category"]
    setting_values = [
        ["GSTPct", 18, "Number", "GST percentage", 18, "Yes", "Pricing"],
        ["PIPct", 3.14, "Number", "Financial percentage", 3.14, "Yes", "Pricing"],
        ["FreightMethod", "B", "Text", "Freight calculation method", "B", "Yes", "Freight"],
        ["RatePerKm", 125.32, "Number", "Trip rate per kilometre", 125.32, "Yes", "Freight"],
        ["TripRateManual", 98000, "Number", "Manual legacy trip rate", 98000, "Yes", "Freight"],
        ["HandlingRsPerMT", 100, "Number", "Handling cost per MT", 100, "Yes", "Freight"],
        ["TrailerPayloadMT", 28, "Number", "Trailer payload", 28, "Yes", "Freight"],
        ["KmRounding", "Exact", "Text", "Distance rounding rule", "Exact", "Yes", "Freight"],
        ["TripDirection", "One-way", "Text", "Trip direction", "One-way", "Yes", "Freight"],
        ["UndercutRsPerMT", 500, "Number", "Competitor undercut amount", 500, "Yes", "Competitor"],
        ["MinMarginPct", 0, "Number", "Minimum margin", 0, "Yes", "Pricing"],
        ["MarketFreightRsPerMTperKm", 0, "Number", "Market freight override", 0, "Yes", "Freight"],
        ["ValidityDays", 5, "Integer", "Offer validity days", 5, "Yes", "Commercial"],
        ["PriceUnitRule", "Follow enquiry", "Text", "Price unit behavior", "Follow enquiry", "Yes", "Pricing"],
        ["StaleDays", 30, "Integer", "Stale data threshold", 30, "Yes", "System"],
        ["EnquirySerialStart", 595, "Integer", "Initial enquiry serial", 595, "Yes", "Sequence"],
        ["ORSProfile", "driving-hgv", "Text", "ORS routing profile", "driving-hgv", "Yes", "ORS"],
        ["TruckUpliftPct", 0, "Number", "Truck uplift", 0, "Yes", "Freight"],
        ["DefaultSteelRate", 61000, "Number", "Default steel rate per MT", 61000, "Yes", "Pricing"],
        ["DefaultGradeExtra", 0, "Number", "Default grade extra", 0, "Yes", "Pricing"],
        ["DefaultYieldPct", 2, "Number", "Default yield percentage", 2, "Yes", "Pricing"],
        ["DefaultConversion", 5000, "Number", "Default conversion cost", 5000, "Yes", "Pricing"],
        ["DefaultFinancialPct", 1.5, "Number", "Default financial percentage", 1.5, "Yes", "Pricing"],
        ["DefaultTesting", 0, "Number", "Default testing cost", 0, "Yes", "Pricing"],
        ["DefaultMargin", 0, "Number", "Default margin", 0, "Yes", "Pricing"],
        ["PipeLengthMin", 6, "Number", "Minimum pipe length", 6, "Yes", "Validation"],
        ["PipeLengthMax", 13, "Number", "Maximum pipe length", 13, "Yes", "Validation"],
    ]
    table(settings, 2, 0, "tblSettings", settings_headers, setting_values, [28, 20, 14, 36, 20, 12, 18])

    customers = ws["CUSTOMERS"]
    customer_headers = ["CustomerID", "CustomerName", "ContactPerson", "Email", "Phone", "Address", "City", "State", "Country", "PIN", "Latitude", "Longitude", "Active", "CreatedDate", "UpdatedDate"]
    table(customers, 2, 0, "tblCustomers", customer_headers, [["C001", "Unipro Techno Infrastructure Pvt Ltd", "Rakesh Gupta", "", "", "Krishna Nagar, Bagodar, Jharkhand - 825322", "Bagodar", "Jharkhand", "India", "825322", "", "", "Yes", "", ""]], [20] * len(customer_headers))
    plants = ws["PLANTS"]
    plant_headers = ["PlantID", "PlantName", "PlantType", "Company", "Address", "City", "State", "Country", "PIN", "Latitude", "Longitude", "Active", "Notes"]
    table(plants, 2, 0, "tblPlants", plant_headers, [["P01", "Mega Pipes Jabalpur (Maneri)", "OURS", "Mega Pipes", "", "Jabalpur", "Madhya Pradesh", "India", "", "", "", "Yes", ""], ["P02", "Mega Pipes Khopoli", "OURS", "Mega Pipes", "", "Khopoli", "Maharashtra", "India", "", "", "", "Yes", ""]], [20] * len(plant_headers))
    payment = [["PT01", "25% Advance + balance against LC", "Yes"], ["PT02", "100% Advance", "Yes"], ["PT03", "30 days credit", "Yes"]]
    table(masters, 133, 0, "tblPaymentTerms", ["PaymentTermID", "PaymentTerm", "Active"], payment, [20, 42, 12])

    table(ws["DISTANCE_CACHE"], 2, 0, "tblDistanceCache", ["CacheKey", "Profile", "PlantID", "CustomerPIN", "CustomerLatitude", "CustomerLongitude", "DistanceKM", "DurationMinutes", "Source", "CalculatedDate", "ValidUntil", "Status"], [[""] * 12], [20] * 12)
    table(ws["QUOTE_LOG"], 2, 0, "tblQuoteLog", ["LogID", "QuoteID", "Timestamp", "User", "Action", "OldValue", "NewValue", "SheetName", "Status", "Notes"], [[""] * 10], [20] * 10)
    for worksheet in ws.values():
        worksheet.hide_gridlines(2)
    workbook.close()


if __name__ == "__main__":
    build()
