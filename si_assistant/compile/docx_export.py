"""Export the SI as a Word document (build/SI.docx), for people who edit in Word.

What is editable text in Word, and what is a picture:

  text        the cover, General Information, custom sections, general procedures, every
              compound's characterisation data, captions, table and figure labels, and
              the references (formatted by the journal's citation style, the same as in
              the PDF), with bold, italic, sub- and superscripts kept;
  tables      real Word tables (merged cells, header rows repeated on each page, rules
              above and below the header and at the bottom); a picture in a cell stays a
              picture;
  headings    Word's own heading styles, numbered as in the PDF (Project settings ▸ Output
              ▸ Heading numbers), with a Word table of contents that Word fills in when the
              file is opened;
  pictures    structures, figures, reaction schemes, X-ray drawings, attachment pages and
              (when chosen) the spectra pages and HRMS reports, drawn by Typst exactly as in
              the PDF and placed as PNG pictures.

The page numbers ("S1", "S2"…) and header/footer lines follow Project settings ▸ Page.

How it works: the same Builder that writes the PDF's Typst source is used for everything
that is worked out (current compound IDs, labels, the characterisation text, structure
drawings, the spectra layout); this module walks the project in the same order as
Builder.whole_document and, instead of Typst text, lists Word items. The pictures and
drawn blocks are put into one Typst document, each starting a page of its own, and
compiled to PNG in one go; the citations and the reference list go through Typst's HTML
output so they are numbered and formatted as in the PDF.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import struct
from dataclasses import dataclass, field, replace
from html.parser import HTMLParser
from pathlib import Path

import typst

from ..core import front, headings
from ..core import page as page_setup
from ..core import scheme_editor
from ..core.markup import NORMAL, SUB, SUPER, Paragraph, Run, parse
from ..core.project import (BUILD_DIR, GENERAL_OWNER, SPECTRA_SECTION, Compound, Project,
                            SchemePage, Section, figure_width_cm)
from ..core.references import to_hayagriva
from ..core.tables import Table
from .document import (ATTACHMENT_MAX_HEIGHT_IN, CM_PER_IN, HRMS_HEADING, PAGE_BREAK,
                       SPECTRA_HEADING, STRUCTURE_MAX, STYLES_DIR, TABLE_FIT_DEFINITION,
                       TEXT_WIDTH_IN, Builder, CompileError, CompileOptions, editor_scheme)
from .text import format_chemistry, leave_alone

DOCX_NAME = "SI.docx"
RENDER_NAME = "SI-docx"            # build/SI-docx.typ: the pictures, drawn by Typst
PPI = 300                          # the pictures' resolution…
SPECTRA_PPI = 200                  # …and the spectra pages' and HRMS reports' (smaller file)
PAGE_TEXT_HEIGHT_IN = 9.0          # US letter with 1 in margins
MARKER = "si-docx-block"
STYLE_NUMBERING = {"acs": None, "rsc": "1.1", "wiley": "1.1."}   # each style's own
FONT = "Times New Roman"
# Where the compounds' structures go in the Word document (Compile SI window; stored as
# "docx_structures" in project.json).
STRUCTURE_WRAPS = {"center": "Centred on a line of their own (as in the PDF)",
                   "left": "On the left, with the text wrapping around them",
                   "right": "On the right, with the text wrapping around them"}


# ---------------------------------------------------------------------- what goes in

@dataclass
class Heading:
    level: int
    runs: list[Run]
    number: str = ""          # "2.", "A.", "1.1" … ("" = unnumbered)
    outlined: bool = True     # listed in the table of contents
    compound: bool = False    # a compound's title (bold, never italic)
    float_rid: int | None = None   # its structure, beside the title with the text around it


@dataclass
class ClearFloat:
    """What follows starts below a structure the text wraps around."""


@dataclass
class Para:
    runs: list[Run]
    align: str = ""           # "", "left", "center", "right", "justify"
    size: float = 0           # points (0 = the body text's)
    keep_next: bool = False
    hanging: float = 0        # inches (the reference list)


@dataclass
class Picture:
    rid: int                  # its picture(s) from the Typst render
    align: str = "center"
    keep_next: bool = False
    kind: str = "block"       # "pic", "block", "report", "portrait", "landscape"


@dataclass
class TableItem:
    table: Table
    folder: Path | None
    where: str


@dataclass
class PageBreak:
    pass


@dataclass
class Contents:
    title: list[Run]


@dataclass
class References:
    pass


# ---------------------------------------------------------------------- the Typst render

class Renderer:
    """Typst code for each picture or drawn block, each starting a page of its own; one
    compile makes them all (rid → its PNG pages)."""

    PAGES = {
        "pic": "#set page(width: auto, height: auto, margin: 0pt)",
        "block": f"#set page(width: {TEXT_WIDTH_IN}in, height: auto, margin: (x: 0pt, y: 2pt))",
        "report": f"#set page(width: {TEXT_WIDTH_IN}in, height: auto, margin: (x: 0pt, y: 2pt))",
        "portrait": f"#set page(width: {TEXT_WIDTH_IN}in, height: {PAGE_TEXT_HEIGHT_IN}in, "
                    "margin: 0pt)",
        "landscape": f"#set page(width: {PAGE_TEXT_HEIGHT_IN}in, height: {TEXT_WIDTH_IN}in, "
                     "margin: 0pt)",
    }

    def __init__(self):
        self.codes: list[tuple[str, str]] = []   # (kind, code)
        self.pages: dict[int, list[bytes]] = {}
        self.ppi: dict[int, int] = {}

    def add(self, code: str, kind: str = "block") -> int:
        self.codes.append((kind, code))
        return len(self.codes) - 1

    def source(self, b: Builder, rids: list[int]) -> str:
        parts = ['#import "style.typ": *', "#show: si-doc.with()",
                 headings.typst_set_rule(b.project.settings), TABLE_FIT_DEFINITION,
                 # (links in the PDF; here only their text)
                 "#let si-link(target, body) = body",
                 "#set page(header: none, footer: none, numbering: none, fill: white)"]
        for rid in rids:
            kind, code = self.codes[rid]
            parts += [self.PAGES[kind], "#pagebreak(weak: true)",
                      f"#context [#metadata((rid: {rid}, page: here().page())) <{MARKER}>]",
                      code]
        return "\n\n".join(x for x in parts if x) + "\n"

    def render(self, b: Builder, background=None) -> None:
        """`background(work)`: runs Typst's work (e.g. in a worker thread, so the window
        keeps drawing itself); by default here."""
        pages = [rid for rid, (kind, _c) in enumerate(self.codes)
                 if kind in ("portrait", "landscape", "report")]
        rest = [rid for rid in range(len(self.codes)) if rid not in set(pages)]
        self._render(b, rest, PPI, RENDER_NAME, background)
        self._render(b, pages, SPECTRA_PPI, RENDER_NAME + "-spectra", background)

    def _render(self, b: Builder, rids: list[int], ppi: int, name: str, background) -> None:
        if not rids:
            return
        typ = b.build / f"{name}.typ"
        typ.write_text(self.source(b, rids), encoding="utf-8")
        root = str(b.root)

        def work():
            marks = json.loads(typst.query(str(typ), f"<{MARKER}>", field="value", root=root))
            return marks, typst.compile(str(typ), format="png", ppi=ppi, root=root)
        try:
            marks, images = (background or (lambda f: f()))(work)
        except Exception as e:  # noqa: BLE001 — Typst reports problems as plain exceptions
            raise CompileError("The pictures for the Word file could not be drawn. This is "
                               "usually caused by an unusual file or character; details for "
                               f"troubleshooting:\n\n{e}") from e
        if isinstance(images, bytes):
            images = [images]
        starts = sorted((int(m["page"]), int(m["rid"])) for m in marks)
        for i, (page, rid) in enumerate(starts):
            end = starts[i + 1][0] if i + 1 < len(starts) else len(images) + 1
            self.pages[rid] = images[page - 1:max(page, end) - 1]
            self.ppi[rid] = ppi


def png_size(data: bytes) -> tuple[int, int]:
    return struct.unpack(">II", data[16:24])


# ---------------------------------------------------------------------- heading numbers

def _letters(n: int) -> str:
    out = ""
    while n > 0:
        n, rest = divmod(n - 1, 26)
        out = chr(65 + rest) + out
    return out


def _format(kind: str, n: int) -> str:
    if kind == "A":
        return _letters(n)
    if kind == "a":
        return _letters(n).lower()
    if kind == "i":
        return page_setup.roman(n).lower()
    return str(n)


class HeadingNumbers:
    """The numbers Typst gives the headings (only numbered headings count), worked out in
    the same order (Project settings ▸ Heading numbers, core/headings.py)."""

    def __init__(self, settings: dict, style: str):
        self.mode = headings.heading_numbers(settings)
        if self.mode == "style":
            self.pattern = STYLE_NUMBERING.get(style)
        elif self.mode == "decimal":
            self.pattern = "1.1."
        elif self.mode == "none":
            self.pattern = None
        else:
            self.pattern = "outline"
        self.counts: list[int] = []

    def next(self, level: int, numbered: bool) -> str:
        if not numbered or self.pattern is None:
            return ""
        self.counts = (self.counts + [0] * level)[:level]
        self.counts[-1] += 1
        if self.pattern == "outline":
            levels = headings.OUTLINE_LEVELS[self.mode]
            return _format(levels[min(level, len(levels)) - 1], self.counts[-1]) + "."
        text = ".".join(str(c) for c in self.counts)
        return text + "." if self.pattern.endswith(".") else text


# ---------------------------------------------------------------------- walking the project

class DocxWriter:
    """Lists the Word items of the SI, in Builder.whole_document's order."""

    def __init__(self, b: Builder, include_spectra: bool = True, wrap: str = "center"):
        """`wrap`: the compounds' structures centred on a line of their own ("center"), or
        on the left or right of the compound's title with the text wrapping around them."""
        self.b = b
        self.wrap = wrap
        self.project = b.project
        self.include_spectra = include_spectra
        self.items: list = []
        self.render = Renderer()
        style = str(self.project.settings.get("journal_style") or "ACS").strip().lower()
        self.numbers = HeadingNumbers(self.project.settings, style)
        self.cite_groups: list[tuple[str, ...]] = []   # in order of first appearance
        self.cite_text: dict[tuple[str, ...], list[Run]] = {}
        self.reference_runs: list[list[Run]] = []
        self.cell_pictures: dict[tuple[int, int, int], int] = {}   # see cell_pictures

    # -- small pieces

    def heading(self, level: int, runs: list[Run], outlined: bool = True,
                numbered: bool = True) -> None:
        self.items.append(Heading(level, runs, self.numbers.next(level, numbered), outlined))

    def section_heading(self, level: int, name: str) -> None:
        self.heading(level, [Run(name)], outlined=level == 1 or self.b.toc_rank >= 2)

    def note_cites(self, runs: list[Run]) -> list[Run]:
        for r in runs:
            if r.ref and r.ref.startswith("cite:"):
                key = tuple(r.citations)
                if key and key not in self.cite_groups:
                    self.cite_groups.append(key)
        return runs

    def para(self, runs: list[Run], **kw) -> None:
        if any(r.text or r.ref for r in runs):
            self.items.append(Para(self.note_cites(runs), **kw))

    def picture(self, code: str | None, kind: str = "pic", align: str = "center",
                keep_next: bool = False) -> bool:
        if not code:
            return False
        if kind == "pic":
            code = f"#{code}" if not code.startswith("#") else code
        self.items.append(Picture(self.render.add(code, kind), align or "center", keep_next,
                                  kind))
        return True

    def caption(self, label: str, caption: list[Run], above: bool) -> None:
        runs = [Run(f"{label}.", bold=True)]
        if any(r.text or r.ref for r in caption):
            runs += [Run(" ")] + caption
        self.items.append(Para(self.note_cites(runs), align="center", size=10,
                               keep_next=above))

    def formatted(self, runs: list[Run]) -> list[Run]:
        return format_chemistry(runs) if self.b.options.format_text else runs

    # -- text with figures and tables (Builder.rich_text)

    def rich_text(self, paragraphs: list[Paragraph], where: str, folder: Path | None = None,
                  heading_offset: int = 1, outlined_levels: int = 2, owner=None,
                  numbered_headings: bool = False) -> None:
        b = self.b
        paragraphs = b.link_paragraphs(paragraphs, where)
        index = 0
        for p in paragraphs:
            runs = self.formatted(p.runs)
            if not p.is_object:
                if p.heading:
                    level = p.heading + heading_offset
                    if any(r.text for r in runs):
                        self.heading(level, self.note_cites(runs),
                                     outlined=level <= outlined_levels,
                                     numbered=numbered_headings)
                else:
                    self.para(runs, align=p.stored_align)
                continue
            item = b.labels.at(owner.key, index) if owner is not None else None
            label = item.label if item is not None else ""
            index += 1
            if p.table:
                table = self.project.read_table(owner, p.table) if owner is not None else None
                if table is None or not table.n_rows or not table.n_cols:
                    b.warnings.append(f"{where}: a table could not be found or is empty "
                                      f"(tables/{p.table}.json); left out.")
                    continue
                self.table(table, label or "Table", runs, where, owner.folder)
                continue
            path, width_cm = None, None
            if p.figure and owner is not None:
                path = self.project.figure_path(owner, p.figure)
                width_cm = figure_width_cm(self.project.figures(owner).get(p.figure, {}))
                if path is None:
                    b.warnings.append(f"{where}: a figure's picture is no longer in the "
                                      "project; left out.")
            elif p.image and folder is not None:
                path = folder / p.image
            if path is None:
                continue
            if width_cm:
                img = b.image(path, (round(min(TEXT_WIDTH_IN, width_cm / CM_PER_IN), 3), 9.0),
                              natural=False)
            else:
                img = b.image(path, (TEXT_WIDTH_IN, 7.0))
            if img and self.picture(img[0], align=p.stored_align or "center", keep_next=True):
                self.caption(label or "Figure", runs, above=False)

    def table(self, table: Table, label: str, caption: list[Run], where: str,
              folder: Path | None) -> None:
        self.caption(label, caption, above=True)
        self.items.append(TableItem(table, folder, where))

    # -- the parts of the SI

    def cover(self) -> None:
        s = self.project.settings
        self.items.append(Para([Run(front.cover_heading(s), bold=True)], align="center",
                               size=16))
        title = str(s.get("title", "") or "")
        if title:
            self.items.append(Para([Run(title, bold=True)], align="center", size=14))
        runs: list[Run] = []
        authors = front.authors(s)
        for i, a in enumerate(authors):
            if i:
                runs.append(Run(", and " if i == len(authors) - 1 and len(authors) > 2
                                else " and " if i == len(authors) - 1 else ", "))
            runs.append(Run(str(a.get("name", "") or "")))
            marks = ",".join(a["affiliations"]) + ("*" if a.get("corresponding") else "")
            if marks:
                runs.append(Run(marks, valign=SUPER))
        self.para(runs, align="center")
        for a in front.affiliations(s):
            self.para([Run(str(a.get("marker", "") or ""), valign=SUPER),
                       Run(str(a.get("text", "") or ""))], align="center", size=10)
        emails = [str(a.get("email", "")) for a in authors if a.get("corresponding")
                  and a.get("email")]
        if emails:
            self.para([Run("*Email: " + "; ".join(emails))], align="center", size=10)
        extra = self.project.read_front_text("cover")
        if extra.strip():
            self.rich_text(parse(extra), "Cover page", self.project.front_folder,
                           outlined_levels=0)
        self.items.append(PageBreak())

    def general_information(self, level: int = 1) -> None:
        text = self.project.read_front_text("general")
        if not text.strip():
            return
        s = self.project.settings
        title = front.general_heading(s).strip()
        outline = headings.numbers_text_headings(s)
        if title:
            self.heading(level, [Run(title)], numbered=front.general_numbered(s) or outline)
        self.rich_text(parse(text), title or "General Information", self.project.front_folder,
                       heading_offset=level,
                       outlined_levels=level + 1 if self.b.toc_rank >= 2 else 0,
                       owner=self.project.text_owner(GENERAL_OWNER),
                       numbered_headings=outline)

    def text_section(self, section: Section, level: int) -> None:
        self.section_heading(level, section.name)
        self.rich_text(parse(self.project.read_section_text(section)), section.name,
                       section.folder, heading_offset=level,
                       outlined_levels=level + 1 if self.b.toc_rank >= 2 else 0,
                       owner=self.project.text_owner(section.id),
                       numbered_headings=headings.numbers_text_headings(self.project.settings))

    def file_section(self, section: Section, level: int) -> None:
        b = self.b
        heading = str(section.data.get("heading", "")).strip()
        if heading:
            self.section_heading(level, heading)
        path = self.project.section_file_path(section)
        if path is None:
            b.warnings.append(f"Attachment section “{section.name}” has no file attached.")
            return
        is_pdf = path.suffix.lower() == ".pdf"
        for page in (b.pdf_pages(path) if is_pdf else [None]):
            if is_pdf:
                img = b.image(path, (TEXT_WIDTH_IN, ATTACHMENT_MAX_HEIGHT_IN), pdf_page=page,
                              sizing=f'width: 100%, height: {ATTACHMENT_MAX_HEIGHT_IN:g}in, '
                                     'fit: "contain"')
                if img:
                    self.picture(f"#align(center, {img[0]})", kind="block")
            else:
                img = b.image(path, (TEXT_WIDTH_IN, ATTACHMENT_MAX_HEIGHT_IN))
                if img:
                    self.picture(img[0])
        caption = str(section.data.get("caption", "")).strip()
        if caption:
            self.para([Run(caption)], align="center", size=10)

    def compound(self, compound: Compound, show_structure: bool = True, level: int = 2) -> None:
        b = self.b
        # The same text as the PDF's (Builder.compound): the PDF's Typst is thrown away,
        # but making it gives the same notes in the warnings.
        b.compound(compound, show_structure=False, level=level)
        from .blocks import compound_block
        block = compound_block(self.project, compound, b.report.display_id(compound.id),
                               b.options.format_text, None)
        self.items.append(Heading(level, self.note_cites(block.title), "",
                                  outlined=b.toc_rank >= 4, compound=True))
        wrapped = False
        if show_structure:
            code = b.structure_image(compound, STRUCTURE_MAX)
            if code and self.wrap in ("left", "right"):
                self.items[-1].float_rid = self.render.add(f"#{code}", "pic")
                wrapped = True
            else:
                self.picture(code, keep_next=True)
        where = (f"{b.report.display_id(compound.id)} ({compound.name})" if compound.name
                 else b.report.display_id(compound.id))
        for p in b.link_paragraphs(block.paragraphs, where):
            if p.heading:
                self.heading(p.heading + level, self.note_cites(p.runs), outlined=False,
                             numbered=False)
            else:
                self.para(p.runs, align=p.stored_align)
        spectra = b.inline_spectra(compound)
        if wrapped and (spectra or block.figures):
            self.items.append(ClearFloat())   # (pictures across the page start below it)
            wrapped = False
        for code in spectra:   # (e.g. HPLC traces, as in the PDF)
            self.picture(code, kind="block")
        for index, fig in enumerate(block.figures):
            img = b.image(fig.path, (TEXT_WIDTH_IN, 4.0))
            item = b.labels.at(compound.id, index)
            if img and self.picture(img[0], keep_next=True):
                self.caption(item.label if item else "Figure", self.formatted([Run(fig.caption)]),
                             above=False)
        if wrapped:
            self.items.append(ClearFloat())

    def scheme(self, scheme: SchemePage, level: int) -> None:
        self.picture(editor_scheme(self.b, scheme), kind="block")
        drawn = set(scheme_editor.linked_compound_ids(scheme))
        for child in scheme.children:
            self.compound(child, show_structure=child.id not in drawn, level=level)
            self.anchored_xray(child, level)

    def anchored_xray(self, compound: Compound, level: int) -> None:
        for x in self.project.anchored.get(compound.id) or []:
            if x.id not in self.b.options.excluded:
                self.xray(x, level, heading=False)

    def xray(self, section: Section, level: int, heading: bool = True) -> None:
        from ..core import xray as xr
        b = self.b
        if heading:
            self.section_heading(level, section.name)
        where = f"X-ray structure “{section.name}”"
        loaded = xr.load(self.project, section)
        if loaded.grown is None:
            b.warnings.append(f"{where}: {loaded.error} Left out.")
            return
        view = xr.view_of(section, loaded.grown)
        svg = b.assets / f"{section.id}.svg"
        svg.parent.mkdir(parents=True, exist_ok=True)
        svg.write_text(xr.svg(loaded.grown, view), encoding="utf-8")
        img = b.image(svg, (TEXT_WIDTH_IN, 9.0))
        if img and self.picture(img[0], keep_next=True):
            item = b.labels.at(section.id, 0)
            caption = [r for p in parse(xr.caption(self.project, section, view, loaded.block))
                       for r in p.runs]
            self.caption(item.label if item is not None else "Figure",
                         b.links(caption, where), above=False)
        if section.data.get("table", True):
            data = xr.table_json(section, loaded.block)
            if data["rows"]:
                item = b.labels.at(section.id, 1)
                caption = [r for p in parse(xr.table_caption(self.project, section))
                           for r in p.runs]
                self.table(Table.from_json(data), item.label if item is not None else "Table",
                           b.links(caption, where), where, None)

    def general_procedures(self, section: Section, level: int) -> None:
        b = self.b
        for gp in self.project.general_procedures(section):
            name = str(gp.get("name", "")).strip() or "General Procedure"
            title = str(gp.get("title", "")).strip()
            self.heading(level, self.formatted([Run(name + (f": {title}" if title else ""))]),
                         outlined=b.toc_rank >= 3, numbered=False)
            text = self.project.read_gp_text(gp["id"])
            if not text.strip():
                b.warnings.append(f"“{name}” in “{section.name}” has no text yet.")
            self.rich_text([replace(p, align="") for p in parse(text)],
                           f"{section.name} ▸ {name}", heading_offset=level, outlined_levels=0)

    def data_section(self, section: Section, level: int) -> None:
        self.section_heading(level, section.name)
        inner = level + 1
        groups = self.project.group_headings(section)
        grouped = bool(groups and self.project.general_procedures(section) and section.items)
        if grouped:
            self.section_heading(level + 1, groups[0])
            inner = level + 2
        self.general_procedures(section, inner)
        if grouped:
            self.section_heading(level + 1, groups[1])
        for item in section.items:
            if isinstance(item, SchemePage):
                self.scheme(item, inner)
            else:
                self.compound(item, level=inner)
                self.anchored_xray(item, inner)

    def section_block(self, section: Section, level: int, skip: frozenset[str]) -> None:
        if section.id in skip:
            return
        if section.type == SPECTRA_SECTION:
            if front.SPECTRA_KEY not in skip and section is self.project.spectra_section:
                self.spectra(section.name, level, skip)
        elif section.type == "text":
            self.text_section(section, level)
        elif section.type == "file":
            self.file_section(section, level)
        elif section.type == "xray":
            self.xray(section, level, heading=not section.anchored_to)
        else:
            self.data_section(section, level)
        for child in section.children:
            self.section_block(child, level + 1, skip)

    def spectra(self, name: str, level: int, skip: frozenset[str]) -> None:
        """The spectra pages and HRMS reports (Builder.spectra_section), as pictures."""
        if not self.include_spectra:
            return
        b = self.b
        compounds = b.included_compounds(self.project.top_sections, skip)
        pages = [p for p in re.split(r"\n\n(?=#page\(flipped: )", b.spectra_pages(compounds))
                 if p.strip()]
        reports = []
        if front.HRMS_KEY not in skip:
            from ..core import ms_reports
            for compound in compounds:
                if not ms_reports.prints_reports(self.project, compound):
                    continue
                for item in ms_reports.report_items(compound):
                    code = b.hrms_report(compound, str(item.get("file", "")))
                    if code:
                        reports.append(code)
        if not pages and not reports:
            return
        self.items.append(PageBreak())
        self.section_heading(level, name)
        for i, code in enumerate(pages):
            m = re.fullmatch(r"#page\(flipped: (true|false)\)\[\n(.*)\n\]", code, re.S)
            if m is None:
                continue
            if i:
                self.items.append(PageBreak())
            self.picture(m.group(2), kind="landscape" if m.group(1) == "true" else "portrait")
        if reports:
            self.items.append(PageBreak())
            self.section_heading(level + 1, HRMS_HEADING)
            for code in reports:
                parts = [x for x in code.split(PAGE_BREAK) if x.strip()]
                for part in parts:
                    self.picture(part, kind="report")

    def document(self) -> None:
        b, p = self.b, self.project
        excluded = set(b.options.excluded)
        if front.front_page_included("cover", excluded):
            self.cover()
        if front.front_page_included("toc", excluded):
            self.items.append(Contents([Run(b.toc["title"])] if b.toc["title"].strip() else []))
            self.items.append(PageBreak())
        if front.front_page_included("general", excluded):
            self.general_information()
        for s in p.top_sections:
            self.section_block(s, 1, frozenset(excluded))
        if p.spectra_section is None and front.SPECTRA_KEY not in excluded:
            self.spectra(SPECTRA_HEADING, 1, frozenset(excluded))
        if self.cite_groups:
            self.items.append(References())

    # -- citations and the reference list, numbered and formatted by Typst

    def format_citations(self) -> None:
        if not self.cite_groups:
            return
        b = self.b
        style_file = b.build / "style.typ"
        m = re.search(r'#let si-citation-style = "([^"]+)"',
                      style_file.read_text(encoding="utf-8"))
        csl = m.group(1) if m else "american-chemical-society"
        yml = b.build / f"{RENDER_NAME}-references.yml"
        yml.write_text(to_hayagriva(self.project.references), encoding="utf-8")
        groups = ["".join(f"#cite(label({json.dumps(c)}))" for c in g) for g in self.cite_groups]
        typ = b.build / f"{RENDER_NAME}-references.typ"
        typ.write_text("\n\n".join(groups) + f'\n\n#bibliography("{yml.name}", style: '
                       f"{json.dumps(csl)})\n", encoding="utf-8")
        try:
            html = typst.compile(str(typ), format="html", root=str(b.root))
        except Exception as e:  # noqa: BLE001
            raise CompileError("The references could not be formatted; details for "
                               f"troubleshooting:\n\n{e}") from e
        html = html.decode("utf-8") if isinstance(html, bytes) else str(html)
        cites = re.findall(r"<p>(.*?)</p>", html, re.S)
        for key, text in zip(self.cite_groups, cites):
            self.cite_text[key] = html_runs(text)
        self.reference_runs = [html_runs(x) for x in re.findall(r"<li[^>]*>(.*?)</li>", html,
                                                                 re.S)]

    def cite_runs(self, run: Run) -> list[Run]:
        runs = self.cite_text.get(tuple(run.citations))
        if runs is None:
            return [Run("[?]")]
        return [replace(r, bold=r.bold or run.bold, italic=r.italic or run.italic) for r in runs]


class _HtmlRuns(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack: list[str] = []
        self.runs: list[Run] = []
        self.prefix = False

    def handle_starttag(self, tag, attrs):
        self.stack.append(tag)
        if tag == "span" and ("class", "prefix") in attrs:
            self.prefix = True

    def handle_endtag(self, tag):
        if tag in self.stack:
            while self.stack and self.stack.pop() != tag:
                pass
        if tag == "span" and self.prefix:
            self.prefix = False
            self.runs.append(Run(" "))

    def handle_data(self, data):
        if not data:
            return
        s = self.stack
        valign = SUPER if "sup" in s else SUB if "sub" in s else NORMAL
        run = Run(data, bold="strong" in s or "b" in s, italic="em" in s or "i" in s,
                  valign=valign)
        if self.runs and self.runs[-1].same_format(run):
            self.runs[-1] = replace(self.runs[-1], text=self.runs[-1].text + data)
        else:
            self.runs.append(run)


def html_runs(html: str) -> list[Run]:
    parser = _HtmlRuns()
    parser.feed(html)
    parser.close()
    return [r for r in parser.runs if r.text]


# ---------------------------------------------------------------------- writing the .docx

class _Word:
    """python-docx helpers for the items above."""

    def __init__(self, writer: DocxWriter):
        from docx import Document
        self.w = writer
        self.doc = Document()
        self.cover_page = False
        self.break_next = False
        self._setup()

    # -- styles and the page

    def _setup(self) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt, RGBColor
        doc = self.doc
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Inches(8.5), Inches(11)
        for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, side, Inches(1))
        normal = doc.styles["Normal"]
        self._font(normal, FONT, 11)
        pf = normal.paragraph_format
        pf.space_before, pf.space_after = Pt(0), Pt(8)
        pf.line_spacing = 1.0
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        looks = {1: (13, True, False), 2: (11.5, True, False), 3: (11, True, True)}
        for level in range(1, 10):
            size, bold, italic = looks.get(level, (11, False, True))
            for name in (f"Heading {level}", self._plain_heading_name(level)):
                style = self._heading_style(name, level)
                self._font(style, FONT, size)
                style.font.bold, style.font.italic = bold, italic
                style.font.color.rgb = RGBColor(0, 0, 0)
                style.paragraph_format.space_before = Pt(14 if level == 1 else 10)
                style.paragraph_format.space_after = Pt(6)
                style.paragraph_format.keep_with_next = True
                style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        self.settings_update_fields()

    @staticmethod
    def _plain_heading_name(level: int) -> str:
        return f"SI Heading {level} (not in contents)"

    def _heading_style(self, name: str, level: int):
        from docx.enum.style import WD_STYLE_TYPE
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        styles = self.doc.styles
        if name in [s.name for s in styles]:
            return styles[name]
        style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = styles["Normal"]
        style.next_paragraph_style = styles["Normal"]
        style.quick_style = True
        ppr = style.element.get_or_add_pPr()
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), "9")   # body text: not in the table of contents
        ppr.append(outline)
        return style

    @staticmethod
    def _font(style, name: str, size: float) -> None:
        from docx.oxml.ns import qn
        from docx.shared import Pt
        style.font.name = name
        style.font.size = Pt(size)
        rpr = style.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is not None:
            for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                fonts.attrib.pop(qn(attr), None)
            for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                fonts.set(qn(attr), name)

    def settings_update_fields(self) -> None:
        """Word fills in the table of contents (and page fields) when the file is opened."""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        settings = self.doc.settings.element
        el = OxmlElement("w:updateFields")
        el.set(qn("w:val"), "true")
        settings.append(el)

    # -- runs and fields

    def add_runs(self, paragraph, runs: list[Run], size: float = 0) -> None:
        from docx.shared import Pt
        for r in runs:
            if r.ref:
                if r.ref.startswith("cite:"):
                    self.add_runs(paragraph, self.w.cite_runs(r), size)
                continue
            if not r.text:
                continue
            run = paragraph.add_run(r.text)
            run.bold = bool(r.bold) or None
            run.italic = bool(r.italic) or None
            if r.underline:
                run.underline = True
            if r.valign == SUPER:
                run.font.superscript = True
            elif r.valign == SUB:
                run.font.subscript = True
            if size:
                run.font.size = Pt(size)

    @staticmethod
    def add_field(paragraph, instr: str, shown: str = "", bold: bool = False,
                  size: float = 0) -> None:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt

        def piece(kind: str | None = None, text: str | None = None, code: str | None = None):
            run = paragraph.add_run()
            if bold:
                run.bold = True
            if size:
                run.font.size = Pt(size)
            if kind:
                el = OxmlElement("w:fldChar")
                el.set(qn("w:fldCharType"), kind)
                run._r.append(el)
            if code is not None:
                el = OxmlElement("w:instrText")
                el.set(qn("xml:space"), "preserve")
                el.text = code
                run._r.append(el)
            if text is not None:
                run.text = text
        piece("begin")
        piece(code=f" {instr} ")
        piece("separate")
        piece(text=shown)
        piece("end")

    # -- items

    def write(self) -> None:
        from docx.enum.section import WD_ORIENT, WD_SECTION
        items = self.w.items
        # Landscape spectra pages share a landscape section (with the headings just before
        # the first one); anything else turns the page back.
        wide = [False] * len(items)
        ahead = False
        for i in range(len(items) - 1, -1, -1):
            item = items[i]
            if isinstance(item, Picture):
                ahead = item.kind == "landscape"
            elif not isinstance(item, (Heading, PageBreak, ClearFloat)):
                ahead = False
            wide[i] = ahead
        for i, item in enumerate(items):
            self._orientation(WD_ORIENT.LANDSCAPE if wide[i] else WD_ORIENT.PORTRAIT,
                              WD_SECTION)
            if isinstance(item, Heading):
                self.heading(item)
            elif isinstance(item, Para):
                self.para(item)
            elif isinstance(item, Picture):
                self.picture(item)
            elif isinstance(item, TableItem):
                self.table(item)
            elif isinstance(item, PageBreak):
                self.page_break()
            elif isinstance(item, ClearFloat):
                self.clear_float()
            elif isinstance(item, Contents):
                self.contents(item)
            elif isinstance(item, References):
                self.references()
        self.page_lines()

    def _last_is_break(self) -> bool:
        paras = self.doc.paragraphs
        if not paras:
            return True
        last = paras[-1]._p
        return bool(last.xpath("./w:pPr/w:sectPr"))

    def page_break(self) -> None:
        """The next paragraph starts a new page ("page break before": a page break
        paragraph after a full page would leave an empty page)."""
        if not self._last_is_break():
            self.break_next = True

    def add_paragraph(self, style=None):
        p = self.doc.add_paragraph(style=style)
        if self.break_next:
            p.paragraph_format.page_break_before = True
            self.break_next = False
        return p

    def heading(self, h: Heading) -> None:
        level = max(1, min(9, h.level))
        style = f"Heading {level}" if h.outlined else self._plain_heading_name(level)
        p = self.add_paragraph(style)
        runs = ([Run(h.number + " ")] if h.number else []) + h.runs
        self.add_runs(p, runs)
        if h.compound:   # (a compound's title looks like the PDF's: bold, upright, 11 pt)
            from docx.shared import Pt
            for run in p.runs:
                run.font.size = Pt(11)
                if run.bold is None:
                    run.bold = True
                run.italic = run.italic if run.italic else False
            p.paragraph_format.alignment = 3   # justify, as a title that runs over lines
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
        if h.float_rid is not None:
            self.floating_picture(p, h.float_rid, self.w.wrap)

    MAX_FLOAT_IN = 2.5   # a structure the text wraps around, at most this wide

    def floating_picture(self, paragraph, rid: int, side: str) -> None:
        """The picture anchored to `paragraph` on the left or right of the text column, with
        the text wrapping around it (Word: Wrap Text ▸ Square)."""
        import copy
        from io import BytesIO

        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        from docx.shared import Inches
        pages = self.w.render.pages.get(rid) or []
        if not pages:
            return
        data = pages[0]
        width = min(self._picture_width(data, "pic"), self.MAX_FLOAT_IN)
        run = paragraph.add_run()
        run.add_picture(BytesIO(data), width=Inches(width))
        inline = run._r.xpath(".//wp:inline")[0]
        gap = 114300   # 0.125 in between the picture and the text
        anchor = parse_xml(
            f'<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="0" distB="{gap // 2}" '
            f'distL="{gap}" distR="{gap}" simplePos="0" relativeHeight="251658240" '
            'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="0">'
            '<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="column"><wp:align>{side}</wp:align></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            '</wp:anchor>')
        for tag in ("wp:extent", "wp:effectExtent"):
            found = inline.find(tag.replace("wp:", "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"))
            if found is not None:
                anchor.append(copy.deepcopy(found))
        anchor.append(parse_xml(f'<wp:wrapSquare {nsdecls("wp")} '
                                f'wrapText="{"right" if side == "left" else "left"}"/>'))
        for tag in ("docPr", "cNvGraphicFramePr"):
            found = inline.find(f"{{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}}{tag}")
            if found is not None:
                anchor.append(copy.deepcopy(found))
        anchor.append(copy.deepcopy(inline.find("{http://schemas.openxmlformats.org/drawingml/2006/main}graphic")))
        inline.getparent().replace(inline, anchor)
        # (move the picture's run to the start, so the title starts beside it)
        paragraph._p.remove(run._r)
        first = paragraph._p.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r")
        if first is not None:
            first.addprevious(run._r)
        else:
            paragraph._p.append(run._r)

    def clear_float(self) -> None:
        """A line break that moves what follows below a wrapped structure."""
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt
        # (in a tiny paragraph of its own: at the end of a justified paragraph it would
        # stretch that paragraph's last line)
        p = self.add_paragraph()
        p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = Pt(1)
        run = p.add_run()
        run.font.size = Pt(1)
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "textWrapping")
        br.set(qn("w:clear"), "all")
        run._r.append(br)

    def para(self, item: Para) -> None:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches
        p = self.add_paragraph()
        aligns = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                  "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
        if item.align in aligns:
            p.paragraph_format.alignment = aligns[item.align]
        if item.keep_next:
            p.paragraph_format.keep_with_next = True
        if item.hanging:
            p.paragraph_format.left_indent = Inches(item.hanging)
            p.paragraph_format.first_line_indent = Inches(-item.hanging)
        self.add_runs(p, item.runs, item.size)

    def _picture_width(self, data: bytes, kind: str, ppi: int = PPI) -> float:
        w, h = png_size(data)
        width, height = w / ppi, h / ppi
        room_w, room_h = ((PAGE_TEXT_HEIGHT_IN, TEXT_WIDTH_IN - 0.6) if kind == "landscape"
                          else (TEXT_WIDTH_IN, PAGE_TEXT_HEIGHT_IN - 0.6))
        return min(width, room_w, room_h * width / height if height else room_w)

    def picture(self, item: Picture) -> None:
        from io import BytesIO

        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches, Pt
        pages = self.w.render.pages.get(item.rid) or []
        for i, data in enumerate(pages):
            if i:
                self.page_break()
            p = self.add_paragraph()
            p.paragraph_format.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                                            "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(
                item.align, WD_ALIGN_PARAGRAPH.CENTER)
            p.paragraph_format.space_after = Pt(4)
            if item.keep_next:
                p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(BytesIO(data), width=Inches(self._picture_width(
                data, item.kind, self.w.render.ppi.get(item.rid, PPI))))

    def _orientation(self, orient, WD_SECTION) -> None:
        from docx.shared import Inches
        current = self.doc.sections[-1]
        if current.orientation == orient:
            return
        self.break_next = False   # (the section break starts a new page)
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        sec.orientation = orient
        wide = orient != 0
        sec.page_width, sec.page_height = ((Inches(11), Inches(8.5)) if wide
                                           else (Inches(8.5), Inches(11)))
        sec.different_first_page_header_footer = False

    def table(self, item: TableItem) -> None:
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Pt
        b, w = self.w.b, self.w
        t = item.table.copy()
        header = t.header_rows
        for r, _c, cell in t.cells():
            if r < header:
                cell.rowspan = min(cell.rowspan, header - r)
        t.normalize()
        if self.break_next:   # (a table can't start a page by itself)
            self.add_paragraph().paragraph_format.space_after = Pt(0)
        table = self.doc.add_table(rows=t.n_rows, cols=t.n_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        aligns = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                  "right": WD_ALIGN_PARAGRAPH.RIGHT}
        for r, c, cell in t.cells():
            target = table.cell(r, c)
            if cell.rowspan > 1 or cell.colspan > 1:
                target = target.merge(table.cell(r + cell.rowspan - 1, c + cell.colspan - 1))
            target.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            align = aligns.get(cell.align or t.columns[c], WD_ALIGN_PARAGRAPH.CENTER)
            paragraphs = []
            picture = w.cell_pictures.get((id(item), r, c))
            if picture is not None:
                paragraphs.append(("picture", picture))
            for p in b.link_paragraphs(parse(cell.text), item.where):
                paragraphs.append(("text", w.formatted(p.runs)))
            first = True
            for kind, value in paragraphs:
                par = target.paragraphs[0] if first else target.add_paragraph()
                first = False
                par.paragraph_format.alignment = align
                par.paragraph_format.space_after = Pt(0)
                if kind == "picture":
                    self._cell_picture(par, value)
                else:
                    self.add_runs(par, value, 10)
            if first:
                target.paragraphs[0].paragraph_format.space_after = Pt(0)
        # rules: above the table, under the header rows, at the bottom
        for r, row in enumerate(table.rows):
            if r < header:
                tr_pr = row._tr.get_or_add_trPr()
                el = OxmlElement("w:tblHeader")
                tr_pr.append(el)
            for cell in row.cells:
                edges = {}
                if r == 0:
                    edges["top"] = 6
                if header and r == header - 1:
                    edges["bottom"] = 4
                if r == t.n_rows - 1:
                    edges["bottom"] = 6
                if edges:
                    self._borders(cell, edges)
        # Columns as wide as their contents (Word's AutoFit to contents), not equal
        # widths; the table kept whole on one page unless it may run over a page break.
        # (as si-fit-table in the PDF: no column narrower than its longest word or picture,
        # the rest of the width to the columns with the most text)
        widths = self._column_widths(item, t)
        tbl_pr = table._tbl.tblPr
        width = tbl_pr.find(qn("w:tblW"))
        if width is None:
            width = OxmlElement("w:tblW")
            tbl_pr.append(width)
        width.set(qn("w:type"), "dxa")
        width.set(qn("w:w"), str(int(sum(widths) * 1440)))
        for grid, w_in in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths):
            grid.set(qn("w:w"), str(int(w_in * 1440)))
        last = len(table.rows) - 1
        for r, row in enumerate(table.rows):
            col = 0
            for tc in row._tr.findall(qn("w:tc")):
                span = tc.tcPr.find(qn("w:gridSpan")) if tc.tcPr is not None else None
                n = int(span.get(qn("w:val"))) if span is not None else 1
                tc_w = tc.get_or_add_tcPr().find(qn("w:tcW"))
                if tc_w is not None:
                    tc_w.set(qn("w:type"), "dxa")
                    tc_w.set(qn("w:w"), str(int(sum(widths[col:col + n]) * 1440)))
                col += n
            for cell in row.cells:
                for par in cell.paragraphs:
                    par.paragraph_format.space_before = Pt(1)
                    if r < last and not t.breakable:
                        par.paragraph_format.keep_with_next = True
        notes = [w.formatted(p.runs) for n in t.footnotes
                 for p in b.link_paragraphs(parse(n), item.where)]
        for runs in notes:
            par = self.doc.add_paragraph()
            par.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            par.paragraph_format.space_after = Pt(0)
            self.add_runs(par, w.note_cites(runs), 9)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(4)

    PAD_IN = 0.2       # a cell's left and right padding together (with a little room)

    @staticmethod
    def _text_in(text: str) -> float:
        """About how wide `text` is in 10 pt Times (bold), in inches."""
        return sum(0.1 if ch.isupper() or ch in "mwMW%" else 0.035 if ch in "il.,;:'|!()[]"
                   else 0.07 for ch in text)

    def _column_widths(self, item: TableItem, t: Table) -> list[float]:
        """Each column's width (inches): at least its longest word or picture, then the
        width left over shared by how much text each column has, never wider than the
        text width in all."""
        b, w = self.w.b, self.w
        least = [0.3] * t.n_cols
        want = [0.3] * t.n_cols
        for r, c, cell in t.cells():
            if cell.colspan != 1:
                continue
            text = "\n".join("".join(x.text for x in p.runs)
                             for p in b.link_paragraphs(parse(cell.text), item.where))
            words = [self._text_in(x) for x in
                     re.split(r"\s+|(?<=[A-Za-z)\]]-)(?=[A-Za-z(\[])", text) if x]
            longest = max(words, default=0)
            whole = max((self._text_in(line) for line in text.split("\n")), default=0)
            rid = w.cell_pictures.get((id(item), r, c))
            for data in (self.w.render.pages.get(rid) or []) if rid is not None else []:
                pic = self._picture_width(data, "pic")
                longest, whole = max(longest, pic), max(whole, pic)
            least[c] = max(least[c], longest + self.PAD_IN)
            want[c] = max(want[c], whole + self.PAD_IN)
        if sum(want) <= TEXT_WIDTH_IN:
            return want
        if sum(least) >= TEXT_WIDTH_IN:   # (too wide even so: shrunk evenly)
            return [x * TEXT_WIDTH_IN / sum(least) for x in least]
        room = TEXT_WIDTH_IN - sum(least)
        extra = [wa - le for wa, le in zip(want, least)]
        return [le + room * ex / sum(extra) for le, ex in zip(least, extra)]

    def _cell_picture(self, paragraph, rid: int) -> None:
        from io import BytesIO
        from docx.shared import Inches
        for data in self.w.render.pages.get(rid) or []:
            paragraph.add_run().add_picture(BytesIO(data),
                                            width=Inches(self._picture_width(data, "pic")))

    @staticmethod
    def _borders(cell, edges: dict[str, int]) -> None:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        tc_pr = cell._tc.get_or_add_tcPr()
        borders = tc_pr.find(qn("w:tcBorders"))
        if borders is None:
            borders = OxmlElement("w:tcBorders")
            tc_pr.append(borders)
        for edge, size in edges.items():
            el = borders.find(qn(f"w:{edge}"))
            if el is None:
                el = OxmlElement(f"w:{edge}")
                borders.append(el)
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")

    def contents(self, item: Contents) -> None:
        if item.title:
            p = self.add_paragraph(self._plain_heading_name(1))
            self.add_runs(p, item.title)
        p = self.add_paragraph()
        self.add_field(p, 'TOC \\o "1-9" \\h \\z', "Right-click here and choose “Update Field” "
                       "to fill in the table of contents.")

    def references(self) -> None:
        before = self.w.items[-2] if len(self.w.items) > 1 else None
        if isinstance(before, Picture) and before.kind in ("portrait", "landscape", "report", "block"):
            self.page_break()   # (after a full page of spectra or reports)
        p = self.add_paragraph("Heading 1")
        self.add_runs(p, [Run("References")])
        for runs in self.w.reference_runs:
            self.para(Para(runs, align="left", size=10, hanging=0.35))

    # -- header and footer (Project settings ▸ Page)

    def page_lines(self) -> None:
        from docx.enum.text import WD_TAB_ALIGNMENT
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Inches
        s = self.w.project.settings
        v = page_setup.page_settings(s)
        known = {"title": str(s.get("title", "")), "authors": page_setup.author_names(s),
                 "date": page_setup.date_text(self.w.b.today)}
        fmt = {"1": "decimal", "i": "lowerRoman", "I": "upperRoman"}[v["number_style"]]
        first = self.doc.sections[0]
        first.different_first_page_header_footer = self.cover_page and not v["number_on_cover"]
        for sec in self.doc.sections:
            pg = sec._sectPr.find(qn("w:pgNumType"))
            if pg is None:
                pg = OxmlElement("w:pgNumType")
                sec._sectPr.append(pg)
            pg.set(qn("w:fmt"), fmt)

        def fill(part, texts: dict[str, str]) -> None:
            p = part.paragraphs[0]
            p.text = ""
            p.paragraph_format.tab_stops.add_tab_stop(Inches(TEXT_WIDTH_IN / 2),
                                                      WD_TAB_ALIGNMENT.CENTER)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(TEXT_WIDTH_IN),
                                                      WD_TAB_ALIGNMENT.RIGHT)
            size = v["number_size"]
            for i, pos in enumerate(page_setup.POSITIONS):
                if i:
                    p.add_run("\t")
                for kind, value in page_setup.split_fill_ins(texts.get(pos, "")):
                    if kind in ("page", "pages"):
                        if v["number_prefix"]:
                            p.add_run(v["number_prefix"]).bold = v["number_bold"] or None
                        self.add_field(p, "PAGE" if kind == "page" else "NUMPAGES", "1",
                                       bold=v["number_bold"], size=size)
                    else:
                        text = value if kind == "text" else known[kind]
                        runs = [Run(text)]
                        if self.w.b.options.format_text:
                            runs = format_chemistry(runs)
                        self.add_runs(p, runs)

        sec = self.doc.sections[0]
        if any(t.strip() for t in v["header"].values()):
            fill(sec.header, v["header"])
        fill(sec.footer, v["footer"])


# ---------------------------------------------------------------------- the export

@dataclass
class DocxResult:
    path: Path
    warnings: list[str] = field(default_factory=list)


def export_docx(project: Project, options: CompileOptions, image_info, converter=None,
                crop=None, include_spectra: bool = True,
                dest: Path | None = None, background=None,
                wrap: str | None = None) -> DocxResult:
    """Write the SI as build/SI.docx (or `dest`). Raises CompileError; everything else
    is a warning (never blocks). If SI.docx is open in Word (Windows locks it), the new
    file gets a numbered name instead (SI (2).docx…)."""
    options = replace(options, only=None)
    b = Builder(project, options, image_info, name=RENDER_NAME)
    b.converter = converter
    b.crops.detect = crop
    b.build.mkdir(parents=True, exist_ok=True)
    if b.assets.exists():
        shutil.rmtree(b.assets, ignore_errors=True)
    b.assets.mkdir(parents=True, exist_ok=True)
    style = str(project.settings.get("journal_style") or "ACS").strip().lower() or "acs"
    style_file = STYLES_DIR / f"{style}.typ"
    if not style_file.exists():
        b.warnings.append(f"The {project.settings.get('journal_style')} style is not available "
                          "yet; the ACS style was used.")
        style_file = STYLES_DIR / "acs.typ"
    shutil.copy2(style_file, b.build / "style.typ")
    if wrap is None:
        wrap = str(project.settings.get("docx_structures", "center"))
    writer = DocxWriter(b, include_spectra=include_spectra,
                        wrap=wrap if wrap in STRUCTURE_WRAPS else "center")
    with leave_alone(project.dictionary_words()):
        writer.document()
        writer.cell_pictures = cell_pictures(writer)
    b.crops.save()
    writer.render.render(b, background)
    writer.format_citations()
    word = _Word(writer)
    word.cover_page = front.front_page_included("cover", set(options.excluded))
    word.write()
    if any(b.toc.get(k) for k in front.TOC_LISTS):
        b.warnings.append("The Word file's table of contents lists the headings only (no "
                          "lists of figures, tables or schemes).")
    target = Path(dest) if dest is not None else project.folder / BUILD_DIR / DOCX_NAME
    return DocxResult(save_docx(word.doc, target), list(dict.fromkeys(b.warnings)))


def cell_pictures(writer: DocxWriter) -> dict[tuple[int, int, int], int]:
    """The pictures in table cells, drawn with the rest: (table item, row, column) → rid."""
    b = writer.b
    out = {}
    for item in writer.items:
        if not isinstance(item, TableItem):
            continue
        for r, c, cell in item.table.cells():
            code = b.cell_picture(cell, item.folder, item.where)
            if code:
                out[(id(item), r, c)] = writer.render.add(f"#{code}", "pic")
    return out


def save_docx(doc, target: Path) -> Path:
    """Saved through a temporary file; a target held open by Word gets a numbered name."""
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name(target.stem + ".part.docx")
    doc.save(str(tmp))
    for n in range(1, 50):
        dest = target if n == 1 else target.with_name(f"{target.stem} ({n}){target.suffix}")
        try:
            os.replace(tmp, dest)
            return dest
        except PermissionError:
            continue
    raise CompileError(f"{target.name} could not be saved (it may be open in Word). Close "
                       "it there and export again.")
