"""A Data section with general procedures and compounds: the procedures and the compounds
each under a sub-heading of their own (Project.group_headings), named per section."""

import re

from pypdf import PdfReader

from si_assistant.core import headings
from si_assistant.core.project import Project
from tests.test_compile_pdf import compile_sample, sample  # noqa: F401  (fixture)


def lines_of(p) -> list[str]:
    text = "\n".join(pg.extract_text() or "" for pg in PdfReader(str(compile_sample(p).pdf)).pages)
    return [line.strip() for line in text.split("\n")]


def data_section(p):
    return next(s for s in p.data_sections() if s.compounds)


def test_defaults_names_and_switching_off_are_stored(sample):
    p, _parts = sample
    s = data_section(p)
    assert p.group_headings(s) == ("General Procedures", "Characterization Data")
    p.set_group_headings(s.id, True, "Procedures", "Compounds")
    again = Project.open(p.folder, read_only=True)
    assert again.group_headings(again.section(s.id)) == ("Procedures", "Compounds")
    p.set_group_headings(s.id, True, "General Procedures", "")   # (the defaults: not stored)
    assert "gp_heading" not in s.data and "compounds_heading" not in s.data
    p.set_group_headings(s.id, False)
    assert s.data["group_headings"] is False and p.group_headings(s) is None


def test_procedures_and_compounds_under_their_own_numbered_headings(sample):
    p, _parts = sample
    s = data_section(p)
    gp = p.add_general_procedure(s.id, "General Procedure A")
    p.write_gp_text(gp["id"], "The amine (1.0 mmol) was stirred.")
    headings.store_heading_numbers(p.settings, "outline")
    lines = lines_of(p)
    gp_at = lines.index("A. General Procedures")
    cd_at = lines.index("B. Characterization Data")
    # (the body, after the table of contents: the first heading is that line's last use)
    gp_at = len(lines) - 1 - lines[::-1].index("A. General Procedures")
    cd_at = len(lines) - 1 - lines[::-1].index("B. Characterization Data")
    assert gp_at < next(i for i in range(gp_at, len(lines))
                        if lines[i].startswith("General Procedure A")) < cd_at
    # (both in the table of contents)
    assert any(re.match(r"A\. General Procedures \.", x) for x in lines)
    # Renamed, then switched off: the procedure follows the section heading as before.
    p.set_group_headings(s.id, True, "Procedures", "Compounds")
    lines = lines_of(p)
    assert "A. Procedures" in lines and "B. Compounds" in lines
    p.set_group_headings(s.id, False)
    lines = lines_of(p)
    assert "A. Procedures" not in lines and "B. Compounds" not in lines
    assert not any(x.startswith(("A. General Procedures", "B. Characterization")) for x in lines)


def test_no_headings_without_general_procedures(sample):
    p, _parts = sample
    headings.store_heading_numbers(p.settings, "outline")
    lines = lines_of(p)
    assert not any("General Procedures" in x or "Characterization Data" in x for x in lines)


def test_general_procedures_tab_edits_the_headings(qapp, sample):
    from si_assistant.ui.procedure_widgets import GeneralProceduresEditor
    p, _parts = sample
    s = data_section(p)
    ed = GeneralProceduresEditor(p, s)
    assert ed.group_box.isChecked()
    assert ed.gp_heading.placeholderText() == "General Procedures"
    assert ed.compounds_heading.placeholderText() == "Characterization Data"
    ed.compounds_heading.setText("Compound Data")
    ed.compounds_heading.editingFinished.emit()
    assert p.group_headings(s) == ("General Procedures", "Compound Data")
    ed.group_box.setChecked(False)
    assert p.group_headings(s) is None and not ed.compounds_heading.isEnabled()
    ed.group_box.setChecked(True)
    assert p.group_headings(s) == ("General Procedures", "Compound Data")
