"""Import compounds from a spreadsheet (SPEC §13y)."""

from __future__ import annotations

import openpyxl
import pytest

from si_assistant.core import si_import_apply as A
from si_assistant.core import si_import_table as T
from si_assistant.core.project import Project

HEADERS = ["Compound", "Name", "SMILES", "Scale (mmol)", "Mass (mg)", "Yield", "Appearance",
           "Rf", "1H NMR", "13C NMR", "NMR solvent", "1H MHz", "HRMS ion", "HRMS calcd",
           "HRMS found", "ee", "Comments"]


def workbook(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "N-aryl scope"
    ws.append(["Scope of the N-arylation"])            # a title row above the headers
    ws.append(HEADERS)
    ws.append(["3a", "1-Phenylpyridin-2(1H)-one", "O=c1ccccn1-c1ccccc1", 0.5, 57.4, 0.94,
               "White solid", "0.23 (4:1 hexanes/EtOAc)",
               "δ 7.49 – 7.30 (m, 6H), 6.65 (d, J = 9.2 Hz, 1H), 6.24 (t, J = 6.7 Hz, 1H)",
               "δ 162.3, 140.9, 139.9, 129.4, 128.6, 126.6, 121.9, 106.0", "CDCl3", 500,
               "[M+H]+", "172.0757", "172.0760", "94% ee", "check rotamers"])
    ws.append(["3b", "", "not a smiles", 0.5, "1.25 g", "89%", "colorless oil", "", "", "", "",
               "", "", "", "", "", ""])
    ws.append([None] * 3)
    ws["F3"].number_format = "0%"
    ws["D3"].number_format = "0.00"
    wb.save(path)
    return path


def test_headers_are_guessed():
    assert T.guess_roles(HEADERS) == [
        "id", "name", "smiles", "scale", "mass", "yield_percent", "physical_state", "rf",
        "nmr_1H", "nmr_13C", "nmr_solvent", "nmr_mhz_1H", "hrms_adduct", "hrms_calcd",
        "hrms_found", "ee", "notes"]
    assert T.guess_role("Product #") == "id"
    assert T.guess_role("13C{1H} NMR") == "nmr_13C"
    assert T.guess_role("Melting point (°C)") == "mp"
    assert T.guess_role("") == "ignore"


def test_rows_become_compounds(tmp_path):
    read = T.read_table(workbook(tmp_path / "scope.xlsx"))
    assert not read.error
    sheet = read.sheets[0]
    h = sheet.header_row()
    assert h == 1                                       # (the title row is skipped)
    plan = T.plan_from_table(sheet, h, T.guess_roles(sheet.rows[h]))
    assert [c.cid for c in plan.compounds] == ["3a", "3b"]
    assert plan.section_name == "N-aryl scope"
    a, b = plan.compounds
    v = a.values
    assert a.structure and a.use_structure
    assert v["scale"] == {"value": "0.50", "unit": "mmol"}          # as Excel shows it
    assert v["yield"] == {"percent": "94", "mass": "57.4", "mass_unit": "mg"}
    assert v["physical_state"] == {"text": "white solid"}
    assert v["rf"] == {"value": "0.23", "eluent": "4:1 hexanes/EtOAc"}
    assert v["nmr"][0]["text"].startswith("1H NMR (500 MHz, CDCl3) δ 7.49")
    assert v["nmr"][1]["text"].startswith("13C NMR (126 MHz, CDCl3) δ 162.3")
    assert v["hrms"] == {"adduct": "[M+H]+", "found": "172.0760", "ionization": "ESI"}
    assert v["ee"]["value"] == "94% ee"
    assert a.notes == ["Comments: check rotamers"] and not a.unrecognised
    assert b.values["yield"] == {"mass": "1.25", "mass_unit": "g", "percent": "89"}
    assert not b.use_structure and any("SMILES" in w for w in b.warnings)


def test_csv_and_errors(tmp_path):
    f = tmp_path / "scope.csv"
    f.write_text("ID;Name;Yield (%)\n4a;Thing;76\n", encoding="cp1252")
    sheet = T.read_table(f).sheets[0]
    plan = T.plan_from_table(sheet, 0, T.guess_roles(sheet.rows[0]))
    assert plan.compounds[0].values["yield"]["percent"] == "76"
    assert "older Excel" in T.read_table(tmp_path / "old.xls").error


def test_imported_into_a_project(tmp_path):
    read = T.read_table(workbook(tmp_path / "scope.xlsx"))
    sheet = read.sheets[0]
    plan = T.plan_from_table(sheet, 1, T.guess_roles(sheet.rows[1]))
    p = Project.create(tmp_path, "N-aryl SI")
    try:
        result = A.apply_plan(p, plan, "scope.xlsx")
        section = p.section(result.section_id)
        assert section.id_table[:2] == ["3a", "3b"]
        c = p.compound(result.created[0])
        assert p.read_structure(c) is not None
        assert c.values["yield"]["percent"] == "94"
        assert "check rotamers" in p.read_compound_text(c, "notes")
    finally:
        p.close()


def test_columns_window(tmp_path, monkeypatch):
    pytest.importorskip("PySide6")
    from PySide6.QtWidgets import QApplication
    from si_assistant.ui import si_import_table_ui as U
    from si_assistant.ui import si_import_ui
    from si_assistant.core import si_import_read
    QApplication.instance() or QApplication([])
    monkeypatch.setattr(U, "_remembered", lambda: {"comments": "ignore"})
    saved = {}
    monkeypatch.setattr(U, "_remember", lambda h, r: saved.update(zip(h, r)))
    read = T.read_table(workbook(tmp_path / "scope.xlsx"))
    dlg = U.ColumnsDialog(read)
    assert dlg.header.value() == 2
    roles = dlg.roles()
    assert roles[0] == "id" and roles[-1] == "ignore"          # (as remembered)
    assert "2 row(s)" in dlg.count.text() and dlg.next_btn.isEnabled()
    dlg.combos[0].setCurrentIndex(dlg.combos[0].findData("ignore"))
    dlg.combos[1].setCurrentIndex(dlg.combos[1].findData("ignore"))
    assert not dlg.next_btn.isEnabled()
    dlg.combos[0].setCurrentIndex(dlg.combos[0].findData("id"))
    dlg._accept()
    assert saved["Compound"] == "id"
    plan = dlg.plan()
    assert len(plan.compounds) == 2
    review = si_import_ui.ImportDialog(plan, si_import_read.ReadResult("scope.xlsx", "table"),
                                       None)
    assert review.import_btn.isEnabled()
    review.deleteLater()
    dlg.deleteLater()


def test_conditions_written_in_the_headers():
    headers = ["Product", "1H NMR (CDCl3, 400 MHz)", "13C{1H} NMR (CDCl3, 101 MHz)",
               "HRMS (APCI) calcd [M+Na]+", "HRMS found", "Rf solvent", "Rf"]
    roles = T.guess_roles(headers)
    assert roles == ["id", "nmr_1H", "nmr_13C", "hrms_calcd", "hrms_found", "rf_eluent", "rf"]
    c = T.compound_from_row(["5a", "δ 7.2 (m, 5H)", "δ 128.0", "195.1", "195.1003",
                             "4:1 hexanes/EtOAc", "0.4"], [""] * 7, headers, roles)
    assert c.values["nmr"][0]["text"] == "1H NMR (400 MHz, CDCl3) δ 7.2 (m, 5H)."
    assert c.values["nmr"][1]["text"] == "13C NMR (101 MHz, CDCl3) δ 128.0."
    assert c.values["hrms"] == {"found": "195.1003", "adduct": "[M+Na]+", "ionization": "APCI"}
    assert c.values["rf"] == {"value": "0.4", "eluent": "4:1 hexanes/EtOAc"}
    assert not c.warnings


def _png() -> bytes:
    import io
    from PIL import Image
    out = io.BytesIO()
    Image.new("RGB", (40, 30), "white").save(out, "PNG")
    return out.getvalue()


def _with_cell_picture(path, cell="B2"):
    """Add a picture placed in `cell` (Excel's Place in Cell: a rich value) to the
    workbook at `path`, as Excel writes it."""
    import re
    import shutil
    import zipfile
    tmp = path.with_suffix(".tmp")
    shutil.move(path, tmp)
    with zipfile.ZipFile(tmp) as zin, zipfile.ZipFile(path, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = re.sub(rf'<c r="{cell}"([^>]*?)(/?)>'.encode(),
                              rf'<c r="{cell}"\1 vm="1"\2>'.encode(), data, count=1)
            zout.writestr(item, data)
        ns = 'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"'
        zout.writestr("xl/metadata.xml", (
            '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">'
            '<metadataTypes count="1"><metadataType name="XLRICHVALUE"/></metadataTypes>'
            '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="x">'
            '<xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>'
            '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>'))
        zout.writestr("xl/richData/rdrichvalue.xml",
                      f'<rvData {ns} count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>')
        zout.writestr("xl/richData/richValueRel.xml",
                      '<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/'
                      '2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/'
                      'officeDocument/2006/relationships"><rel r:id="rId1"/></richValueRels>')
        zout.writestr("xl/richData/_rels/richValueRel.xml.rels",
                      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                      'relationships"><Relationship Id="rId1" Type="image" '
                      'Target="../media/image1.png"/></Relationships>')
        zout.writestr("xl/media/image1.png", _png())
    tmp.unlink()


def test_a_tracker_with_tables_pictures_and_no_ids(tmp_path):
    """Like a lab's scope tracker: an Excel table under a title row, product pictures in the
    cells, no ID column, fractions for ee, computed yields, failed formulas."""
    from openpyxl.worksheet.table import Table
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Scope Chi"
    ws["A1"] = "Data Table"
    rows = [["Boronic Acid", "Picture", "Yield", "HPLC Coulmn", "HPLC Solvent Conditions",
             "HPLC Rt", "ee%", "Percent Yield", "Mass Spec Calc", "Mass Spec Found",
             "Comments", "Column1"],
            ["2-Chlorophenylboronic acid (AB2)", "#VALUE!", 19.8, "IB",
             "IPA-Hex 2.5%, iso, 15 min, 220nm, F100", "4.104 (maj), 4.434 (min)", 0.96,
             73.94792273562499, 268.0893, 268.089, "done", None],
            ["Planned acid", None, None, None, None, None, None, "#DIV/0!", None, None, None,
             None],
            ["5-Indolylboronic acid (AB38)", None, 20.9, "IB", "IPA-Hex 20%", "7.524  9.484",
             0.89, 0, 273.1392, 273.1389, "", None]]
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            ws.cell(3 + i, 1 + j, v)
    ws.add_table(Table(displayName="Tracker", ref="A3:L6"))
    path = tmp_path / "tracker.xlsx"
    wb.save(path)
    _with_cell_picture(path, "B4")

    read = T.read_table(path)
    sheet = read.sheets[0]
    assert [a.name for a in sheet.areas] == ["Tracker"]
    view = sheet.view(sheet.areas[0])
    assert view.header_row() == 0 and view.rows[0][0] == "Boronic Acid"
    assert view.rows[2][7] == ""                       # (#DIV/0!: no value)
    assert (1, 1) in view.pictures
    roles = T.guess_roles(view.rows[0])
    assert roles == ["notes", "picture", "mass", "ee_column", "ee_eluent", "ee_times", "ee",
                     "yield_percent", "hrms_calcd", "hrms_found", "notes", "ignore"]
    plan = T.plan_from_table(view, 0, roles, first_id="3a")
    assert [c.cid for c in plan.compounds] == ["3a", "3b"]     # (the planned row: no data)
    a, b = plan.compounds
    assert a.picture_pdf and a.picture_ext == ".png" and not a.structure
    assert a.values["yield"] == {"mass": "19.8", "mass_unit": "mg", "percent": "74"}
    assert a.values["ee"] == {"value": "96% ee", "column": "IB", "wavelength": "220",
                              "eluent": "IPA-Hex 2.5%, iso, 15 min, F100",
                              "t_major": "4.104", "t_minor": "4.434"}
    assert a.notes == ["Boronic Acid: 2-Chlorophenylboronic acid (AB2)", "Comments: done"]
    assert "percent" not in b.values["yield"]                 # (0: not measured yet)
    assert any("(maj)/(min)" in w for w in b.warnings)

    p = Project.create(tmp_path, "N-aryl SI")
    try:
        result = A.apply_plan(p, plan, "tracker.xlsx")
        c = p.compound(result.created[0])
        assert c.values["structure"]["image_source"] == "user"
        assert p.section(result.section_id).id_table[:2] == ["3a", "3b"]
    finally:
        p.close()


def test_ids_from():
    assert T.ids_from("3a", 3) == ["3a", "3b", "3c"]
    assert T.ids_from("3y", 4) == ["3y", "3z", "3aa", "3ab"]
    assert T.ids_from("S1", 3, {"S2"}) == ["S1", "S3", "S4"]
