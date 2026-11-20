import zipfile

import pytest

import make_samples
import migrate


@pytest.fixture(scope="module")
def run(tmp_path_factory):
    root = tmp_path_factory.mktemp("demo")
    old, template = root / "old", root / "template" / "Risk Assessment Template.docm"
    old.mkdir()
    template.parent.mkdir()
    for item in make_samples.ASSESSMENTS:
        make_samples.make_old(old, item)
    make_samples.make_template(template)
    results = migrate.migrate(old, template, root / "new")
    return root, template, {r["file"]: r for r in results}


def table(path, index):
    doc = migrate.open_word(path)
    return [[migrate.cell_text(c) for c in row.cells] for row in doc.tables[index].rows]


def test_every_document_is_converted_with_the_same_name(run):
    root, _, results = run
    expected = sorted(f"{item['file']}.docm" for item in make_samples.ASSESSMENTS)
    assert sorted(results) == expected
    assert sorted(p.name for p in (root / "new").glob("*.docm")) == expected


def test_outputs_are_macro_enabled_and_keep_every_template_part(run):
    root, template, _ = run
    with zipfile.ZipFile(template) as tpl:
        template_parts = set(tpl.namelist())
    for path in (root / "new").glob("*.docm"):
        with zipfile.ZipFile(path) as out:
            assert migrate.DOCM_MAIN.encode() in out.read("[Content_Types].xml")
            assert template_parts <= set(out.namelist())


def test_details_land_in_the_right_fields(run):
    root, _, _ = run
    rows = table(root / "new" / "RA-002 Working at Height - Roof Access.docm", 0)
    assert rows[0] == [
        "Assessment title",
        "Roof access for gutter clearing",
        "Site / location",
        "Office block roof",
    ]
    assert rows[1] == ["Assessor", "S. Okafor", "Document ref", ""]
    assert rows[2] == ["Assessment date", "05/04/2024", "Next review", "05/04/2025"]


def test_rows_are_added_and_numbered(run):
    root, _, _ = run
    rows = table(root / "new" / "RA-002 Working at Height - Roof Access.docm", 1)
    assert len(rows) == 1 + 4
    assert [r[0] for r in rows[1:]] == ["1", "2", "3", "4"]
    assert rows[3][1:4] == [
        "Falling tools",
        "Pedestrians below",
        "Exclusion zone at ground level\nTool lanyards",
    ]


def test_risk_is_likelihood_times_severity(run):
    root, _, _ = run
    for path in (root / "new").glob("*.docm"):
        for row in table(path, 1)[1:]:
            if row[4] and row[5]:
                assert row[6] == str(int(row[4]) * int(row[5]))


def test_wrong_score_in_old_document_is_corrected_and_flagged(run):
    root, _, results = run
    name = "RA-003 Forklift Operations.docm"
    assert table(root / "new" / name, 1)[1][4:7] == ["4", "4", "16"]
    assert results[name]["status"] == "Review notes"
    assert "old document said 12" in results[name]["notes"][0]


def test_missing_severity_leaves_score_blank_and_flags_it(run):
    root, _, results = run
    name = "RA-005 Cleaning Chemicals.docm"
    assert table(root / "new" / name, 1)[3][4:7] == ["3", "", ""]
    assert results[name]["status"] == "Review notes"


def test_added_rows_keep_the_template_row_formatting(run):
    root, _, _ = run
    doc = migrate.open_word(
        root / "new" / "RA-002 Working at Height - Roof Access.docm"
    )
    rows = doc.tables[1].rows[1:]
    sizes = {cell.paragraphs[0].runs[0].font.size for row in rows for cell in row.cells}
    alignments = {row.cells[0].paragraphs[0].alignment for row in rows}
    assert len(sizes) == 1 and None not in sizes
    assert len(alignments) == 1 and None not in alignments


def test_clean_documents_report_ok(run):
    _, _, results = run
    assert results["RA-001 Manual Handling - Warehouse.docm"]["status"] == "OK"
