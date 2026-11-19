"""Export to Word (compile/docx_export.py): the sample project as build/SI.docx, read back
with python-docx: headings (numbered as in the PDF), compound text, the table, pictures,
citations and references, page numbers, and the "Include spectra" choice."""

from docx import Document

from si_assistant.compile.docx_export import HeadingNumbers, export_docx, html_runs
from si_assistant.compile.document import CompileOptions
from si_assistant.compile.qt_images import convert_to_png, image_info
from si_assistant.core import headings
from si_assistant.core.markup import SUPER
from tests.test_compile_pdf import sample  # noqa: F401  (fixture)


def export(p, **kw):
    return export_docx(p, CompileOptions.from_settings(p.settings), image_info, convert_to_png,
                       **kw)


def texts(doc):
    return [(par.style.name, par.text) for par in doc.paragraphs]


def pictures(doc) -> int:
    return len(doc.element.body.xpath(".//w:drawing"))


def test_heading_numbers_follow_the_setting():
    n = HeadingNumbers({"heading_numbers": "outline_roman"}, "acs")
    got = [n.next(1, True), n.next(2, True), n.next(3, True), n.next(3, True),
           n.next(4, True), n.next(2, False), n.next(2, True), n.next(1, True)]
    assert got == ["1.", "A.", "i.", "ii.", "a.", "", "B.", "2."]
    n = HeadingNumbers({"heading_numbers": "outline"}, "acs")
    assert [n.next(1, True), n.next(2, True), n.next(3, True)] == ["1.", "A.", "a."]
    n = HeadingNumbers({}, "rsc")
    assert [n.next(1, True), n.next(2, True), n.next(1, True)] == ["1", "1.1", "2"]
    n = HeadingNumbers({}, "wiley")
    assert [n.next(1, True), n.next(2, True)] == ["1.", "1.1."]
    assert HeadingNumbers({}, "acs").next(1, True) == ""
    n = HeadingNumbers({"heading_numbers": "outline"}, "acs")
    assert [n.next(2, True) for _ in range(27)][-2:] == ["Z.", "AA."]


def test_html_citations_and_references_become_runs():
    runs = html_runs('<sup>[</sup><a href="#x"><sup>1</sup></a><sup>–</sup><sup>3]</sup>')
    assert "".join(r.text for r in runs) == "[1–3]" and all(r.valign == SUPER for r in runs)
    runs = html_runs('<span class="prefix"><a>(1)</a></span> Doe, J. <em>J. Am. Chem. Soc.</em> '
                     '<strong>2020</strong>, 1&amp;2.')
    assert "".join(r.text for r in runs).startswith("(1)  Doe, J. J. Am. Chem. Soc. 2020, 1&2.")
    assert any(r.italic and r.text == "J. Am. Chem. Soc." for r in runs)
    assert any(r.bold and r.text == "2020" for r in runs)


def test_sample_exports_to_word(sample):
    p, _parts = sample
    p.add_reference({"id": "r-doe", "type": "article", "authors": ["Doe, J."],
                     "title": "A paper", "journal": "J. Am. Chem. Soc.", "year": "2020",
                     "volume": "142", "pages": "1-5"})
    p.write_front_text("general", "Materials were bought.\n\nAs reported.@cite{r-doe}")
    headings.store_heading_numbers(p.settings, "outline_roman")
    result = export(p)
    assert result.path.name == "SI.docx" and result.path.exists()
    doc = Document(str(result.path))
    lines = texts(doc)
    flat = "\n".join(t for _s, t in lines)
    # cover, contents field, numbered section headings in Word heading styles
    assert lines[0][1] == "Supporting Information"
    assert 'TOC \\o' in doc.element.body.xml
    assert ("Heading 1", "1. General Information") in lines   # (front matter, then the sample's own section of that name)
    assert ("Heading 1", "3. Synthesis of Example Compounds") in lines
    # compound titles are not numbered and not in the contents (TOC depth: sections)
    assert any(s.startswith("SI Heading") and t.startswith("N-Benzyl-N-methylacetamide (2)")
               for s, t in lines)
    assert "164.1070" in flat and "CDCl3" in flat
    # the table is a Word table with its caption above
    assert len(doc.tables) == 1 and "Table S1." in flat
    assert doc.tables[0].cell(1, 0).text.strip() == "1"
    # citation and reference list
    assert ("Heading 1", "References") in lines
    ref = next(par for par in doc.paragraphs if "A paper" in par.text or "Doe" in par.text
               and par.style.name == "Normal" and par.text.startswith("("))
    assert "Doe" in ref.text
    cited = next(par for par in doc.paragraphs if par.text.startswith("As reported."))
    assert any(r.font.superscript and r.text.strip("()[]").startswith("1") for r in cited.runs)
    # pictures: structures, the scheme, the figure and the spectra pages
    with_spectra = pictures(doc)
    assert with_spectra >= 5
    # page numbers "S" + PAGE field in the footer
    footer = doc.sections[0].footer.paragraphs[0]
    assert "S" in footer.text and "PAGE" in footer._p.xml
    # without the spectra
    doc2 = Document(str(export(p, include_spectra=False).path))
    assert pictures(doc2) < with_spectra
    assert ("Heading 1", "5. Spectra") not in texts(doc2)


def test_compile_window_offers_word(sample):
    from si_assistant.ui.compile_ui import CompileDialog
    p, _parts = sample
    dlg = CompileDialog(p, word=True)
    assert dlg.as_word.isChecked() and dlg.word_spectra.isEnabled()
    assert dlg.windowTitle() == "Export to Word"
    dlg.word_spectra.setChecked(False)
    dlg.remember()
    assert p.settings["compile_output"] == "docx" and p.settings["docx_spectra"] is False
    again = CompileDialog(p)   # (as last time)
    assert again.as_word.isChecked() and not again.word_spectra.isChecked()
    again.as_pdf.setChecked(True)
    assert not again.word_spectra.isEnabled() and again.windowTitle() == "Compile SI"


def test_structures_beside_the_text_in_word(sample):
    p, parts = sample
    p.settings["docx_structures"] = "left"
    doc = Document(str(export(p).path))
    body = doc.element.body.xml
    assert "<wp:anchor" in body and '<wp:align>left</wp:align>' in body
    assert 'wrapText="right"' in body and 'w:clear="all"' in body
    title = next(par for par in doc.paragraphs if par.text.startswith("N-Benzyl-N-methyl"))
    assert title._p.xpath(".//wp:anchor")   # (anchored to the compound's title)
    doc = Document(str(export(p, wrap="center").path))
    assert "<wp:anchor" not in doc.element.body.xml


def test_spectra_printed_with_the_compound(sample):
    """Spectrum kinds ticked on the Spectra section's page (or one spectrum's own
    "Printed" choice) go under the compound's data instead of the Spectra section."""
    from si_assistant.core import spectra as S
    p, parts = sample
    c2 = parts["c2"]
    kinds = {i["type"] for i in c2.values["spectra"]}
    assert "13C" in kinds and len(S.spectrum_items(p, c2)) == len(c2.values["spectra"])
    d = p.spectra_defaults()
    d["inline_types"] = ["13C", "nonsense"]
    p.set_spectra_defaults(d)
    assert p.spectra_defaults()["inline_types"] == ["13C"]
    assert all(i["type"] != "13C" for _k, i in S.spectrum_items(p, c2))
    assert [i["type"] for i in S.inline_items(p, c2)] == ["13C"]
    # the PDF: the ¹³C spectrum is drawn in the compound's block, not on a spectra page
    from si_assistant.compile.document import Builder, CompileOptions
    b = Builder(p, CompileOptions.from_settings(p.settings), image_info)
    b.assets.mkdir(parents=True, exist_ok=True)
    block = b.compound(c2)
    assert "si-spectrum(" in block and "13C" in block
    assert "#si-fit-inline(" in block and "in * f" in block   # (shrunk to fit its page)
    # one spectrum says otherwise
    for item in c2.values["spectra"]:
        if item["type"] == "13C":
            item["placement"] = "appendix"
        elif item["type"] == "1H":
            item["placement"] = "inline"
    assert [i["type"] for i in S.inline_items(p, c2)] == ["1H"]
    assert any(i["type"] == "13C" for _k, i in S.spectrum_items(p, c2))
