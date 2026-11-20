"""Create sample input for the demo: old-format risk assessments and the new template.

Everything here is invented. The new template is saved as a macro-enabled document
(.docm); it carries no macros of its own because it is generated without Word.
"""

import shutil
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

DOCX_MAIN = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
DOCM_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml"

ASSESSMENTS = [
    {
        "file": "RA-001 Manual Handling - Warehouse",
        "Activity / Task": "Manual handling of stock in the warehouse",
        "Location": "Main warehouse, Unit 4",
        "Assessed by": "P. Hartley",
        "Date of assessment": "12/03/2024",
        "Review date": "12/03/2025",
        "hazards": [
            (
                "Lifting heavy boxes",
                "Warehouse staff",
                "Manual handling training\nTrolleys provided",
                3,
                3,
                9,
                "Introduce two-person lift rule above 20 kg",
            ),
            (
                "Trailing pallet wrap",
                "Staff, visitors",
                "Housekeeping checks each shift",
                2,
                2,
                4,
                "",
            ),
            (
                "Falling stock from racking",
                "Staff",
                "Racking inspected annually\nLoad limits displayed",
                2,
                5,
                10,
                "Add monthly visual inspection",
            ),
        ],
    },
    {
        "file": "RA-002 Working at Height - Roof Access",
        "Activity / Task": "Roof access for gutter clearing",
        "Location": "Office block roof",
        "Assessed by": "S. Okafor",
        "Date of assessment": "05/04/2024",
        "Review date": "05/04/2025",
        "hazards": [
            (
                "Fall from roof edge",
                "Maintenance contractors",
                "Edge protection fitted\nPermit to work required",
                2,
                5,
                10,
                "Review anchor point certification",
            ),
            (
                "Fragile roof lights",
                "Maintenance contractors",
                "Roof lights marked and covered",
                2,
                5,
                10,
                "",
            ),
            (
                "Falling tools",
                "Pedestrians below",
                "Exclusion zone at ground level\nTool lanyards",
                3,
                4,
                12,
                "",
            ),
            (
                "Adverse weather",
                "Maintenance contractors",
                "Work stopped in high wind or rain",
                3,
                3,
                9,
                "",
            ),
        ],
    },
    {
        "file": "RA-003 Forklift Operations",
        "Activity / Task": "Forklift truck operations in the yard",
        "Location": "Goods-in yard",
        "Assessed by": "P. Hartley",
        "Date of assessment": "19/04/2024",
        "Review date": "19/04/2025",
        "hazards": [
            (
                "Collision with pedestrians",
                "Staff, drivers, visitors",
                "Segregated walkways\nHi-vis mandatory\nSpeed limit 5 mph",
                4,
                4,
                12,
                "Install blue spot warning lights",
            ),
            (
                "Overturning truck",
                "Forklift operators",
                "Licensed operators only\nDaily pre-use checks",
                2,
                5,
                10,
                "",
            ),
            (
                "Battery charging",
                "Forklift operators",
                "Ventilated charging bay\nEye wash station",
                2,
                3,
                6,
                "",
            ),
        ],
    },
    {
        "file": "RA-004 Office DSE",
        "Activity / Task": "Display screen equipment use",
        "Location": "First floor offices",
        "Assessed by": "L. Brennan",
        "Date of assessment": "02/05/2024",
        "Review date": "02/05/2026",
        "hazards": [
            (
                "Poor posture",
                "Office staff",
                "Adjustable chairs\nDSE self-assessment",
                3,
                2,
                6,
                "",
            ),
            (
                "Eye strain",
                "Office staff",
                "Eye tests funded\nScreen breaks encouraged",
                3,
                2,
                6,
                "",
            ),
        ],
    },
    {
        "file": "RA-005 Cleaning Chemicals",
        "Activity / Task": "Use of cleaning chemicals",
        "Location": "All areas",
        "Assessed by": "S. Okafor",
        "Date of assessment": "21/05/2024",
        "Review date": "21/05/2025",
        "hazards": [
            (
                "Skin contact with bleach",
                "Cleaning staff",
                "Gloves and goggles issued\nCOSHH sheets on site",
                3,
                3,
                9,
                "Trial a lower-hazard product",
            ),
            (
                "Mixing incompatible products",
                "Cleaning staff",
                "Products stored separately\nStaff briefing",
                2,
                4,
                8,
                "",
            ),
            (
                "Wet floors",
                "Staff, visitors",
                "Wet floor signs\nCleaning outside busy hours",
                3,
                "",
                "",
                "Severity to be confirmed",
            ),
        ],
    },
]

OLD_COLUMNS = [
    "Hazard",
    "Who might be harmed",
    "Existing control measures",
    "Likelihood (1-5)",
    "Severity (1-5)",
    "Risk",
    "Further action required",
]
NEW_DETAILS = [
    ("Assessment title", "Site / location"),
    ("Assessor", "Document ref"),
    ("Assessment date", "Next review"),
]
NEW_COLUMNS = [
    "No.",
    "Hazard",
    "Persons at risk",
    "Controls in place",
    "L",
    "S",
    "Risk (L × S)",
    "Further controls",
]

NEW_WIDTHS = [0.5, 1.8, 1.4, 2.4, 0.5, 0.5, 0.9, 1.7]


def shade(cell, colour):
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), colour)
    cell._tc.get_or_add_tcPr().append(shading)


def write(cell, text, bold=False, colour=None, size=None):
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.bold = bold
    if colour:
        run.font.color.rgb = RGBColor.from_string(colour)
    if size:
        run.font.size = Pt(size)
    return paragraph


def make_old(folder, item):
    doc = Document()
    doc.add_heading("RISK ASSESSMENT", level=1)
    details = doc.add_table(rows=0, cols=2)
    details.style = "Table Grid"
    for label in (
        "Activity / Task",
        "Location",
        "Assessed by",
        "Date of assessment",
        "Review date",
    ):
        row = details.add_row()
        write(row.cells[0], label, bold=True)
        write(row.cells[1], item[label])
    doc.add_paragraph()
    table = doc.add_table(rows=1, cols=len(OLD_COLUMNS))
    table.style = "Table Grid"
    for cell, title in zip(table.rows[0].cells, OLD_COLUMNS):
        write(cell, title, bold=True)
    for hazard in item["hazards"]:
        row = table.add_row()
        for cell, value in zip(row.cells, hazard):
            lines = str(value).split("\n")
            write(cell, lines[0])
            for line in lines[1:]:
                cell.add_paragraph(line)
    doc.save(folder / f"{item['file']}.docx")


def make_template(path):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = section.right_margin = Inches(0.6)
    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style.font.size = Pt(10)
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Risk Assessment Record")
    run.bold = True
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor.from_string("1F3864")

    details = doc.add_table(rows=0, cols=4)
    details.style = "Table Grid"
    for left, right in NEW_DETAILS:
        row = details.add_row()
        for cell, label in ((row.cells[0], left), (row.cells[2], right)):
            write(cell, label, bold=True)
            shade(cell, "D9E2F3")
    doc.add_paragraph()

    table = doc.add_table(rows=2, cols=len(NEW_COLUMNS))
    table.style = "Table Grid"
    table.autofit = False
    for column, width in zip(table.columns, NEW_WIDTHS):
        column.width = Inches(width)
    for row in table.rows:
        for cell, width in zip(row.cells, NEW_WIDTHS):
            cell.width = Inches(width)
    for cell, heading in zip(table.rows[0].cells, NEW_COLUMNS):
        write(cell, heading, bold=True, colour="FFFFFF")
        shade(cell, "1F3864")
    for i, cell in enumerate(table.rows[1].cells):
        paragraph = write(cell, "", size=9)
        if NEW_COLUMNS[i] in ("No.", "L", "S", "Risk (L × S)"):
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    doc.add_paragraph("Risk = Likelihood (1-5) × Severity (1-5).")
    doc.add_paragraph("Approved by: ____________________    Date: ____________")

    plain = path.with_suffix(".docx")
    doc.save(plain)
    with (
        zipfile.ZipFile(plain) as src,
        zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst,
    ):
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "[Content_Types].xml":
                data = data.replace(DOCX_MAIN.encode(), DOCM_MAIN.encode())
            dst.writestr(info, data)
    plain.unlink()


def main():
    old, template = Path("old"), Path("template")
    for folder in (old, template):
        shutil.rmtree(folder, ignore_errors=True)
        folder.mkdir()
    for item in ASSESSMENTS:
        make_old(old, item)
    make_template(template / "Risk Assessment Template.docm")
    print(f"{len(ASSESSMENTS)} old documents and 1 template created")


if __name__ == "__main__":
    main()
