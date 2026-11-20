"""Move risk assessments from an old Word layout into a new macro-enabled template.

Every output starts from the template file, so its styles, layout and any macros are
carried over untouched. Only the field values and hazard rows are filled in. Each
output is then re-opened and compared with its source, and the result is written to
a checklist so a person knows exactly which documents need a closer look.
"""

import argparse
import copy
import re
import zipfile
from pathlib import Path

from docx.opc.part import PartFactory
from docx.package import Package
from docx.parts.document import DocumentPart
from openpyxl import Workbook
from openpyxl.styles import Font

DOCM_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml"
# python-docx only recognises plain .docx main parts unless told about the macro-enabled type.
PartFactory.part_type_for[DOCM_MAIN] = DocumentPart

# new template label -> old document label
DETAIL_MAP = {
    "Assessment title": "Activity / Task",
    "Site / location": "Location",
    "Assessor": "Assessed by",
    "Assessment date": "Date of assessment",
    "Next review": "Review date",
}
# new template column -> old document column
COLUMN_MAP = {
    "Hazard": "Hazard",
    "Persons at risk": "Who might be harmed",
    "Controls in place": "Existing control measures",
    "L": "Likelihood (1-5)",
    "S": "Severity (1-5)",
    "Further controls": "Further action required",
}
NUMBER_COLUMN, RISK_COLUMN, OLD_RISK_COLUMN = "No.", "Risk (L × S)", "Risk"


def open_word(path):
    return Package.open(str(path)).main_document_part.document


def norm(text):
    return re.sub(r"\s+", " ", text or "").strip().lower()


def cell_text(cell):
    return "\n".join(p.text.strip() for p in cell.paragraphs if p.text.strip())


def set_cell_text(cell, text):
    """Write text into a cell, reusing the cell's existing paragraph and run formatting."""
    lines = str(text).split("\n")
    first = cell.paragraphs[0]
    for extra in cell.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    if not first.runs:
        first.add_run("")
    for run in first.runs[1:]:
        run._r.getparent().remove(run._r)
    first.runs[0].text = lines[0]
    previous = first._p
    for line in lines[1:]:
        clone = copy.deepcopy(first._p)
        previous.addnext(clone)
        previous = clone
        cell.paragraphs[-1].runs[0].text = line


def to_int(text):
    return int(text) if re.fullmatch(r"\d+", text.strip()) else None


def read_old(path):
    """Returns ({label: value}, [hazard dicts], [notes])."""
    doc = open_word(path)
    details, hazards, notes = {}, [], []
    wanted_columns = {norm(v) for v in COLUMN_MAP.values()} | {norm(OLD_RISK_COLUMN)}
    for table in doc.tables:
        header = [norm(c.text) for c in table.rows[0].cells]
        if norm("Hazard") in header:
            unknown = [
                c.text.strip()
                for c in table.rows[0].cells
                if norm(c.text) not in wanted_columns
            ]
            if unknown:
                notes.append(
                    f"Old columns with no place in the new template: {', '.join(unknown)}"
                )
            for row in table.rows[1:]:
                values = {
                    header[i]: cell_text(cell) for i, cell in enumerate(row.cells)
                }
                if any(values.values()):
                    hazards.append(values)
        else:
            for row in table.rows:
                cells = row.cells
                if len(cells) >= 2 and cells[0].text.strip():
                    details[norm(cells[0].text)] = cell_text(cells[1])
    return details, hazards, notes


def fill_template(template_path, details, hazards):
    doc = open_word(template_path)
    notes = []
    details_table, hazard_table = doc.tables[0], doc.tables[1]

    for row in details_table.rows:
        cells = row.cells
        for i in range(0, len(cells) - 1, 2):
            label = cells[i].text.strip()
            old_label = DETAIL_MAP.get(label)
            if old_label is None:
                continue
            value = details.get(norm(old_label), "")
            if value:
                set_cell_text(cells[i + 1], value)
            else:
                notes.append(
                    f"'{label}' left blank: no '{old_label}' in the old document"
                )

    columns = [c.text.strip() for c in hazard_table.rows[0].cells]
    blank_row = hazard_table.rows[1]._tr
    for _ in range(len(hazards) - 1):
        blank_row.addnext(copy.deepcopy(blank_row))

    for number, (row, hazard) in enumerate(
        zip(hazard_table.rows[1:], hazards), start=1
    ):
        likelihood = to_int(hazard.get(norm(COLUMN_MAP["L"]), ""))
        severity = to_int(hazard.get(norm(COLUMN_MAP["S"]), ""))
        risk = (
            likelihood * severity
            if likelihood is not None and severity is not None
            else None
        )
        for cell, column in zip(row.cells, columns):
            if column == NUMBER_COLUMN:
                set_cell_text(cell, number)
            elif column == RISK_COLUMN:
                set_cell_text(cell, "" if risk is None else risk)
            elif column in COLUMN_MAP:
                set_cell_text(cell, hazard.get(norm(COLUMN_MAP[column]), ""))
        if risk is None:
            notes.append(
                f"Hazard {number}: likelihood or severity is missing, so the risk score is blank"
            )
        else:
            old_risk = to_int(hazard.get(norm(OLD_RISK_COLUMN), ""))
            if old_risk is not None and old_risk != risk:
                notes.append(
                    f"Hazard {number}: old document said {old_risk}, but {likelihood} × {severity} = {risk}; "
                    f"{risk} used"
                )
    return doc, notes


def verify(output_path, template_path, details, hazards):
    """Re-open the saved file and compare it with what the old document contained."""
    problems = []
    with zipfile.ZipFile(output_path) as out, zipfile.ZipFile(template_path) as tpl:
        if DOCM_MAIN.encode() not in out.read("[Content_Types].xml"):
            problems.append("Output is not a macro-enabled document")
        missing = set(tpl.namelist()) - set(out.namelist())
        if missing:
            problems.append(f"Template parts missing from output: {sorted(missing)}")
        if "word/vbaProject.bin" in tpl.namelist() and out.read(
            "word/vbaProject.bin"
        ) != tpl.read("word/vbaProject.bin"):
            problems.append("Macro project differs from the template")

    doc = open_word(output_path)
    filled = {}
    for row in doc.tables[0].rows:
        cells = row.cells
        for i in range(0, len(cells) - 1, 2):
            filled[cells[i].text.strip()] = cell_text(cells[i + 1])
    for label, old_label in DETAIL_MAP.items():
        if filled.get(label, "") != details.get(norm(old_label), ""):
            problems.append(f"'{label}' does not match the old document")

    table = doc.tables[1]
    columns = [c.text.strip() for c in table.rows[0].cells]
    rows = table.rows[1:]
    if len(rows) != len(hazards):
        problems.append(
            f"{len(hazards)} hazards in the old document but {len(rows)} rows in the new one"
        )
    for number, (row, hazard) in enumerate(zip(rows, hazards), start=1):
        values = {column: cell_text(cell) for column, cell in zip(columns, row.cells)}
        for column, old_column in COLUMN_MAP.items():
            if values.get(column, "") != hazard.get(norm(old_column), ""):
                problems.append(
                    f"Hazard {number}: '{column}' does not match the old document"
                )
        likelihood, severity = to_int(values.get("L", "")), to_int(values.get("S", ""))
        if (
            likelihood is not None
            and severity is not None
            and values.get(RISK_COLUMN) != str(likelihood * severity)
        ):
            problems.append(
                f"Hazard {number}: risk score is not {likelihood} × {severity}"
            )
    return problems


def write_checklist(path, results):
    wb = Workbook()
    ws = wb.active
    ws.title = "Checklist"
    ws.append(["Document", "Hazards moved", "Status", "Notes for the reviewer"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for res in results:
        ws.append([res["file"], res["hazards"], res["status"], "\n".join(res["notes"])])
    for column, width in zip("ABCD", (44, 15, 16, 100)):
        ws.column_dimensions[column].width = width
    wb.save(path)


def migrate(old_dir, template_path, out_dir):
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    for source in sorted(
        p for p in old_dir.iterdir() if p.suffix.lower() in (".docx", ".docm")
    ):
        details, hazards, notes = read_old(source)
        if not hazards:
            results.append(
                {
                    "file": source.name,
                    "hazards": 0,
                    "status": "NOT CONVERTED",
                    "notes": ["No hazard table found in the old document"],
                }
            )
            continue
        doc, fill_notes = fill_template(template_path, details, hazards)
        target = out_dir / f"{source.stem}.docm"
        doc.save(str(target))
        problems = verify(target, template_path, details, hazards)
        notes = problems + notes + fill_notes
        status = "FAILED CHECK" if problems else ("Review notes" if notes else "OK")
        results.append(
            {
                "file": target.name,
                "hazards": len(hazards),
                "status": status,
                "notes": notes,
            }
        )
    return results


def main(argv=None):
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--old", default="old", type=Path)
    ap.add_argument(
        "--template", default="template/Risk Assessment Template.docm", type=Path
    )
    ap.add_argument("--out", default="new", type=Path)
    args = ap.parse_args(argv)
    results = migrate(args.old, args.template, args.out)
    write_checklist(args.out / "migration_checklist.xlsx", results)
    for res in results:
        print(f"{res['status']:14} {res['file']} ({res['hazards']} hazards)")
        for note in res["notes"]:
            print(f"{'':14}   - {note}")
    return results


if __name__ == "__main__":
    main()
