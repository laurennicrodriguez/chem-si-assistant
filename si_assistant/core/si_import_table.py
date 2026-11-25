"""Import compounds from a spreadsheet: one row per compound (SPEC §13y).

`read_table(path)` reads an Excel workbook (.xlsx/.xlsm) or a CSV/TSV file into sheets of
text cells. `guess_role(header)` says which field a column's header names (the user can
change every choice in the column window). `plan_from_table(sheet, header_row, roles)`
turns the rows below the header into an `ImportPlan` of compounds, which the Import SI
review window shows and core/si_import_apply.py writes, as for a Word or PDF SI.

Each cell is read as it is shown in Excel (a number formatted "0.00" gives "0.50"; a
percentage cell gives "94"), and its text is sorted into the field's parts with the same
readers as an imported SI (core/si_import.py), so "57.4 mg (94%)", "Rf = 0.23 (4:1
hexanes/EtOAc)" or a whole data paragraph all work. Nothing is dropped: a cell that can't
be sorted is kept in the compound's private Notes, as with Import SI.
"""

from __future__ import annotations

import csv
import datetime as _dt
import re
from dataclasses import dataclass, field
from pathlib import Path

from . import si_import as S
from .ids import new_id
from .si_import import ImportedCompound, ImportPlan
from .si_import_read import clean_space

TABLE_SUFFIXES = (".xlsx", ".xlsm", ".csv", ".tsv", ".txt")

# ---------------------------------------------------------------------- what a column holds

# role -> (label in the column window, help)
ROLES: dict[str, tuple[str, str]] = {
    "ignore": ("— not imported —", ""),
    "id": ("Compound ID (e.g. 3a)", "The compound's number as printed."),
    "name": ("Name", ""),
    "smiles": ("Structure (SMILES)", "Drawn and stored as the compound's structure."),
    "picture": ("Structure picture (pictures in the cells)",
                "A picture placed in each cell (Excel: Insert ▸ Pictures ▸ Place in Cell, or a "
                "picture over the cell), kept as the structure's picture until a structure is "
                "pasted from ChemDraw."),
    "procedure": ("Experimental procedure", "Its own procedure text."),
    "data": ("Characterization paragraph (several values)",
             "A whole data paragraph (\"White solid (57 mg, 94%). ¹H NMR …\"), sorted into "
             "the fields."),
    "physical_state": ("Physical state", "e.g. white solid"),
    "yield_percent": ("Yield (%)", "\"94\", \"94%\" or \"57.4 mg (94%)\"."),
    "mass": ("Mass obtained (mg, or with its unit)", "\"57.4\" (mg) or \"1.25 g\"."),
    "scale": ("Reaction scale (mmol, or with its unit)", "\"0.50\" (mmol) or \"0.2 mmol\"."),
    "nmr_1H": ("¹H NMR", "The report; \"¹H NMR (500 MHz, CDCl₃)\" is added when the cell "
                         "starts at δ and the frequency and solvent columns give them."),
    "nmr_13C": ("¹³C NMR", ""),
    "nmr_19F": ("¹⁹F NMR", ""),
    "nmr_31P": ("³¹P NMR", ""),
    "nmr_11B": ("¹¹B NMR", ""),
    "nmr_solvent": ("NMR solvent", "Used for NMR cells that start at δ."),
    "nmr_mhz_1H": ("¹H NMR frequency (MHz)", "Used for NMR cells that start at δ; the ¹³C and "
                                             "¹⁹F frequencies follow from it if not given."),
    "nmr_mhz_13C": ("¹³C NMR frequency (MHz)", ""),
    "nmr_mhz_19F": ("¹⁹F NMR frequency (MHz)", ""),
    "hrms": ("HRMS (whole sentence)", "\"HRMS (ESI) m/z: [M+H]+ calcd for … ; found …\"."),
    "hrms_found": ("HRMS found m/z", ""),
    "hrms_calcd": ("HRMS calcd m/z (checked, not stored)",
                   "The program calculates it from the structure; a typed value is compared."),
    "hrms_adduct": ("HRMS ion / adduct (e.g. [M+H]+)", ""),
    "hrms_ionization": ("HRMS ionization (ESI, APCI…)", ""),
    "hrms_formula": ("HRMS formula", "The formula of M, when it differs from the structure's."),
    "rf": ("Rf", "\"0.23\" or \"Rf = 0.23 (4:1 hexanes/EtOAc)\"."),
    "rf_eluent": ("Rf eluent", ""),
    "mp": ("Melting point (°C)", ""),
    "ir": ("IR (cm⁻¹)", ""),
    "rotation": ("Optical rotation [α]D", "\"+54.4\" or \"[α]D20 +54.4 (c 1.0, CHCl3)\"."),
    "ee": ("ee / er (with HPLC details if in one cell)", "\"94\", \"94% ee\", or the whole "
                                                        "HPLC sentence."),
    "ee_method": ("ee: method (HPLC/SFC/GC)", ""),
    "ee_column": ("ee: column", ""),
    "ee_eluent": ("ee: eluent", ""),
    "ee_flow": ("ee: flow (mL/min)", ""),
    "ee_wavelength": ("ee: wavelength (nm)", ""),
    "ee_times": ("ee: retention times (major and minor)",
                 "\"4.10 (maj), 4.43 (min)\"; without (maj)/(min) the first is taken as the major."),
    "t_major": ("ee: tR major (min)", ""),
    "t_minor": ("ee: tR minor (min)", ""),
    "dr": ("d.r.", ""),
    "ea": ("Elemental analysis", ""),
    "ccdc": ("CCDC number", ""),
    "purification": ("Purification (method / eluent)", "Kept in the Purification field's "
                                                       "eluent; the method can be chosen after."),
    "notes": ("Notes (private, never printed)", ""),
}

# role -> header patterns (lower case, spaces evened out); the first role that matches wins
_GUESSES: list[tuple[str, str]] = [
    ("ignore", r"^column\s*\d+$|^unnamed|^mw\b|molecular weight|^mw of|^m\.w\."),
    ("smiles", r"\bsmiles\b"),
    ("id", r"^(?:compound|cmpd|cpd|product|entry)?\s*(?:id|#|no\.?|number|label)$|^(?:compound|cmpd|cpd|product|substrate)$"),
    ("nmr_mhz_1H", r"(?:1h|¹h|proton).*(?:mhz|freq)|(?:mhz|freq).*(?:1h|¹h|proton)"),
    ("nmr_mhz_13C", r"(?:13c|¹³c|carbon).*(?:mhz|freq)|(?:mhz|freq).*(?:13c|¹³c|carbon)"),
    ("nmr_mhz_19F", r"(?:19f|¹⁹f|fluorine).*(?:mhz|freq)|(?:mhz|freq).*(?:19f|¹⁹f|fluorine)"),
    ("nmr_mhz_1H", r"^(?:nmr\s*)?(?:mhz|frequency|field)$|^nmr\s*(?:mhz|frequency)"),
    ("rf_eluent", r"(?:rf|tlc).*(?:eluent|solvent)"),
    ("nmr_solvent", r"nmr.*solvent|deuterated|^solvent$"),
    ("nmr_13C", r"(?:^|\W)(?:13c|¹³c|c-13|carbon)(?:\{1h\})?(?:\s*nmr)?(?:$|\W)|^13c"),
    ("nmr_19F", r"(?:^|\W)(?:19f|¹⁹f|f-19|fluorine)(?:\s*nmr)?(?:$|\W)|^19f"),
    ("nmr_31P", r"(?:^|\W)(?:31p|³¹p|p-31)(?:\s*nmr)?(?:$|\W)"),
    ("nmr_11B", r"(?:^|\W)(?:11b|¹¹b|b-11)(?:\s*nmr)?(?:$|\W)"),
    ("nmr_1H", r"(?:^|\W)(?:1h|¹h|h-1|proton)(?:\s*nmr)?(?:$|\W)|^1h"),
    ("hrms_calcd", r"(?:hrms|m/z|mass).*calc|^calc"),
    ("hrms_found", r"(?:hrms|m/z|mass).*found|^found"),
    ("hrms_adduct", r"adduct|\bion\b|(?:hrms|ms).*(?:species|ion)"),
    ("hrms_ionization", r"ioni[sz]ation|(?:hrms|ms)\s*(?:method|source|mode)"),
    ("hrms_formula", r"(?:hrms|ms|molecular)?\s*formula"),
    ("hrms", r"^(?:hrms|hr-?ms|ms|mass\s*spec.*)$"),
    ("picture", r"^(?:picture|image|structure|drawing|product structure)s?$"),
    ("ee_times", r"\brt\b|retention|\bt_?r\b"),
    ("t_major", r"major"),
    ("t_minor", r"minor"),
    ("ee_column", r"(?:hplc|sfc|chiral|ee)\s*(?:column|coulmn|colum|colunm)|^(?:column|coulmn)$"),
    ("ee_flow", r"flow"),
    ("ee_wavelength", r"wavelength|\bnm\b|λ|detect"),
    ("ee_method", r"(?:ee|er)\s*method|^(?:hplc|sfc|gc)\s*(?:method)?$"),
    ("ee_eluent", r"(?:hplc|sfc|ee).*(?:eluent|solvent|mobile)|mobile\s*phase"),
    ("ee", r"(?:^|\W)(?:ee|er|e\.e\.|e\.r\.|enantio\w*)(?:$|\W)|% ?ee"),
    ("dr", r"(?:^|\W)(?:dr|d\.r\.|diastereo\w*)(?:$|\W)"),
    ("rf", r"(?:^|\W)(?:rf|r\s?f|tlc)(?:$|\W)"),
    ("mp", r"(?:^|\W)(?:mp|m\.p\.|melting)"),
    ("ir", r"(?:^|\W)(?:ft-?)?ir(?:$|\W)|infrared"),
    ("rotation", r"rotation|\[α\]|\[a\]|alpha|optical"),
    ("ea", r"elemental|^anal"),
    ("ccdc", r"ccdc|x-?ray|crystal"),
    ("purification", r"purif|column\s*(?:eluent|cond)|chromatograph"),
    ("scale", r"scale|mmol|limiting"),
    ("mass", r"^(?:mass|amount|obtained|isolated mass|product mass)|\bmg\b|\(g\)|mass\s*\("),
    ("yield_percent", r"yield|%"),
    ("physical_state", r"state|appearance|physical|colou?r|description"),
    ("procedure", r"procedure|experimental|synthesis|prep"),
    ("data", r"characteri[sz]ation|^(?:analytical |spectroscopic |characterization )?data$|"
             r"spectroscopic"),
    ("name", r"name|iupac|compound"),
    ("notes", r"note|comment|remark|\bnb\b|notebook|boronic|pyridinium|reagent|starting|"
              r"substrate|^salt$|^method$|work-?up|isolation"),
]


def _norm(header: str) -> str:
    return clean_space(str(header)).lower()


def guess_role(header: str, taken: set[str] | None = None) -> str:
    """The role a column's header most likely names ("ignore" if none). A role already
    `taken` by an earlier column is not given twice (except the NMR frequency fallback)."""
    h = _norm(header)
    if not h:
        return "ignore"
    if "nmr" in h:     # "1H NMR (CDCl3, 500 MHz)": the report, with its conditions
        h = re.sub(r"\([^()]*\)", " ", h).strip() or h
    taken = taken or set()
    for role, pattern in _GUESSES:
        if role in taken and role not in MANY:
            continue
        if re.search(pattern, h):
            return role
    return "ignore"


MANY = ("ignore", "notes", "data")      # roles more than one column can have


def guess_roles(headers: list[str]) -> list[str]:
    taken: set[str] = set()
    # Two yield columns ("Yield" and "Percent Yield"): the one without % is the mass.
    yields = [i for i, h in enumerate(headers) if re.search(r"yield", _norm(h))]
    percent = [i for i in yields if re.search(r"%|percent|pct", _norm(h := headers[i]))]
    out = []
    for i, h in enumerate(headers):
        if len(yields) > 1 and percent and i in yields:
            role = "yield_percent" if i == percent[0] else ("mass" if "mass" not in taken
                                                             else "ignore")
        else:
            role = guess_role(h, taken)
        if role not in MANY:
            taken.add(role)
        out.append(role)
    return out


# ---------------------------------------------------------------------- reading the file


@dataclass
class Area:
    """An Excel table (Insert ▸ Table) on a sheet: rows and columns 0-based, inclusive."""
    name: str
    min_row: int
    min_col: int
    max_row: int
    max_col: int


@dataclass
class Sheet:
    name: str
    rows: list[list[str]] = field(default_factory=list)   # every cell as shown, as text
    markup: list[list[str]] = field(default_factory=list) # …with sub/superscripts kept
    pictures: dict = field(default_factory=dict)          # (row, col) -> PNG/JPEG bytes
    areas: list[Area] = field(default_factory=list)       # its Excel tables
    area: str = ""                                        # (a view of one of them)

    def view(self, area: Area) -> "Sheet":
        """Just one of its tables (its header row first)."""
        def cut(rows):
            return [r[area.min_col:area.max_col + 1] for r in rows[area.min_row:area.max_row + 1]]
        pics = {(r - area.min_row, c - area.min_col): v for (r, c), v in self.pictures.items()
                if area.min_row <= r <= area.max_row and area.min_col <= c <= area.max_col}
        return Sheet(self.name, cut(self.rows), cut(self.markup or self.rows), pics, [],
                     area.name)

    def header_row(self) -> int:
        """The likely header row: the first row with at least two filled cells, mostly
        text (not numbers)."""
        counts = [sum(1 for c in row if c.strip()) for row in self.rows[:30]]
        most = max(counts, default=0)
        for i, row in enumerate(self.rows[:30]):
            filled = [c for c in row if c.strip()]
            # (a title row above the headers, "Data Table", has fewer cells)
            if len(filled) >= max(2, most // 2) and \
                    sum(not _is_number(c) for c in filled) >= 0.6 * len(filled):
                return i
        return 0


@dataclass
class TableRead:
    file_name: str
    sheets: list[Sheet] = field(default_factory=list)
    error: str = ""


def _is_number(text: str) -> bool:
    return bool(re.fullmatch(r"\s*[-+]?\d+(?:[.,]\d+)?\s*%?\s*", text))


def _number_text(value: float, fmt: str) -> str:
    """A number as Excel shows it with format `fmt` (its decimals and percent)."""
    fmt = (fmt or "General").split(";")[0]
    percent = "%" in fmt
    if percent:
        value = value * 100
    m = re.search(r"0\.(0+)", fmt)
    if m:
        text = f"{value:.{len(m.group(1))}f}"
    elif re.search(r"(?<![.\d])0(?![.\d])", fmt) and "General" not in fmt:
        text = f"{value:.0f}"
    else:
        text = f"{value:.10g}"
        if "e" in text:
            text = repr(value)
    return text


def _rich_markup(value) -> str:
    """A cell with formatted text (openpyxl CellRichText) in the program's markup:
    subscripts, superscripts, italics and bold kept."""
    from .markup import NORMAL, SUB, SUPER, Paragraph, Run, serialize
    runs = []
    for block in value:
        if isinstance(block, str):
            runs.append(Run(block))
            continue
        font = getattr(block, "font", None)
        valign = getattr(font, "vertAlign", None) if font else None
        runs.append(Run(block.text, bold=bool(font and font.b), italic=bool(font and font.i),
                        valign={"subscript": SUB, "superscript": SUPER}.get(valign or "", NORMAL)))
    return serialize([Paragraph(runs)])


def _cell_text(cell) -> tuple[str, str]:
    """(plain text, markup) of a cell as Excel shows it."""
    v = cell.value
    if v is None:
        return "", ""
    if isinstance(v, bool):
        return str(v), str(v)
    if isinstance(v, (int, float)):
        t = _number_text(float(v), cell.number_format)
        return t, t
    if isinstance(v, (_dt.datetime, _dt.date)):
        t = v.strftime("%Y-%m-%d")
        return t, t
    if hasattr(v, "__iter__") and not isinstance(v, str):     # CellRichText
        plain = "".join(b if isinstance(b, str) else b.text for b in v)
        try:
            return plain, _rich_markup(v)
        except Exception:         # (an odd rich text: its plain text is enough)
            return plain, plain
    t = str(v)
    if t in EXCEL_ERRORS:         # (a formula that failed: no value)
        return "", ""
    return t, t


EXCEL_ERRORS = {"#DIV/0!", "#VALUE!", "#N/A", "#REF!", "#NAME?", "#NUM!", "#NULL!",
                "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#GETTING_DATA", "#UNKNOWN!"}


def read_table(path: Path) -> TableRead:
    path = Path(path)
    suffix = path.suffix.lower()
    out = TableRead(path.name)
    try:
        if suffix in (".xlsx", ".xlsm"):
            out.sheets = _read_xlsx(path)
        elif suffix in (".csv", ".tsv", ".txt"):
            out.sheets = [_read_csv(path)]
        elif suffix == ".xls":
            out.error = (f"“{path.name}” is an older Excel file (.xls). Open it in Excel and "
                         "save it as an Excel Workbook (.xlsx) (File ▸ Save As), then import "
                         "that.")
        else:
            out.error = f"“{path.name}” is not an Excel (.xlsx) or CSV file."
    except PermissionError:
        out.error = (f"“{path.name}” could not be opened: it may be open in Excel with "
                     "changes not yet saved, or still downloading from OneDrive. Save or close "
                     "it and try again.")
    except Exception as e:      # (a damaged or password-protected workbook)
        out.error = (f"“{path.name}” could not be read ({type(e).__name__}: {e}). If it is "
                     "password-protected, save an unprotected copy.")
    if not out.error and not any(s.rows for s in out.sheets):
        out.error = f"“{path.name}” has no filled cells."
    return out


def _read_xlsx(path: Path) -> list[Sheet]:
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False, rich_text=True)
    except TypeError:            # (an older openpyxl without rich text)
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    pictures = _cell_pictures(path)
    sheets = []
    for ws in wb.worksheets:
        if getattr(ws, "sheet_state", "visible") != "visible":
            continue
        # merged cells: every cell of a merged block shows its value (a header spread
        # across two columns names both)
        merged = {}
        for rng in ws.merged_cells.ranges:
            first = ws.cell(rng.min_row, rng.min_col)
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    merged[(r, c)] = first
        rows: list[list[str]] = []
        marks: list[list[str]] = []
        for r, row in enumerate(ws.iter_rows(), start=1):
            texts, mk = [], []
            for c, cell in enumerate(row, start=1):
                plain, markup = _cell_text(merged.get((r, c), cell))
                texts.append(plain)
                mk.append(markup)
            rows.append(texts)
            marks.append(mk)
        while rows and not any(t.strip() for t in rows[-1]):
            rows.pop()
            marks.pop()
        width = max((max((i + 1 for i, t in enumerate(r) if t.strip()), default=0)
                     for r in rows), default=0)
        sheet = Sheet(ws.title, [r[:width] + [""] * (width - len(r)) for r in rows],
                      [m[:width] + [""] * (width - len(m)) for m in marks])
        for t in ws.tables.values():
            try:
                from openpyxl.utils.cell import range_boundaries
                c1, r1, c2, r2 = range_boundaries(t.ref)
            except (ValueError, TypeError):
                continue
            sheet.areas.append(Area(t.displayName or t.name, r1 - 1, c1 - 1,
                                    min(r2, len(rows)) - 1, min(c2, width) - 1))
        sheet.areas.sort(key=lambda a: (a.min_row, a.min_col))
        sheet.pictures = pictures.get(ws.title, {})
        for img in getattr(ws, "_images", []):        # pictures over the cells
            try:
                anchor = img.anchor._from
                data = img._data()
            except Exception:
                continue
            sheet.pictures.setdefault((anchor.row, anchor.col), data)
        sheets.append(sheet)
    return sheets


def _cell_pictures(path: Path) -> dict[str, dict[tuple[int, int], bytes]]:
    """Pictures placed in cells (Excel's "Place in Cell" / IMAGE()), which openpyxl does not
    read, by sheet name and (row, col), 0-based. A cell with a picture has a `vm` index into
    xl/metadata.xml → rich value → richValueRel → the picture in xl/media."""
    import posixpath
    import zipfile
    import xml.etree.ElementTree as ET
    M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    RD = "{http://schemas.microsoft.com/office/spreadsheetml/2017/richdata}"
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    out: dict[str, dict[tuple[int, int], bytes]] = {}
    try:
        z = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile):
        return out
    with z:
        names = set(z.namelist())
        if "xl/metadata.xml" not in names or "xl/richData/rdrichvalue.xml" not in names:
            return out
        try:
            md = ET.fromstring(z.read("xl/metadata.xml"))
            types = [t.get("name") for t in md.find(M + "metadataTypes")]
            future = {f.get("name"): [bk.find(".//" + RD + "rvb") for bk in f.findall(M + "bk")]
                      for f in md.findall(M + "futureMetadata")}
            values = []
            vmeta = md.find(M + "valueMetadata")
            for bk in (vmeta.findall(M + "bk") if vmeta is not None else []):
                rc = bk.find(M + "rc")
                values.append((types[int(rc.get("t")) - 1], int(rc.get("v"))))
            rvs = ET.fromstring(z.read("xl/richData/rdrichvalue.xml")).findall(RD + "rv")
            rels = [r.get(R + "id") for r in
                    ET.fromstring(z.read("xl/richData/richValueRel.xml"))]
            targets = {r.get("Id"): r.get("Target") for r in ET.fromstring(
                z.read("xl/richData/_rels/richValueRel.xml.rels")).iter(PR + "Relationship")}
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            wb_rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(
                z.read("xl/_rels/workbook.xml.rels")).iter(PR + "Relationship")}
        except (KeyError, ValueError, TypeError, AttributeError, IndexError, ET.ParseError):
            return out

        def picture(vm: int) -> bytes | None:
            try:
                name, v = values[vm - 1]
                i = int(future[name][v].get("i"))
                rel = int(rvs[i].find(RD + "v").text)
                target = posixpath.normpath(posixpath.join("xl/richData", targets[rels[rel]]))
                return z.read(target)
            except (KeyError, ValueError, TypeError, AttributeError, IndexError):
                return None
        sheets_el = wb.find(M + "sheets")
        for sh in (sheets_el if sheets_el is not None else []):
            target = wb_rels.get(sh.get(R + "id"), "")
            part = posixpath.normpath(posixpath.join("xl", target)) if not target.startswith("/") \
                else target.lstrip("/")
            if part not in names:
                continue
            found: dict[tuple[int, int], bytes] = {}
            try:
                root = ET.fromstring(z.read(part))
            except ET.ParseError:
                continue
            for c in root.iter(M + "c"):
                vm = c.get("vm")
                if not vm:
                    continue
                m = re.fullmatch(r"([A-Z]+)(\d+)", c.get("r", ""))
                data = picture(int(vm)) if m else None
                if data:
                    col = 0
                    for ch in m.group(1):
                        col = col * 26 + ord(ch) - 64
                    found[(int(m.group(2)) - 1, col - 1)] = data
            if found:
                out[sh.get("name")] = found
    return out


def _read_csv(path: Path) -> Sheet:
    raw = path.read_bytes()
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        dialect = csv.Sniffer().sniff(text[:5000], delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel_tab if path.suffix.lower() == ".tsv" else csv.excel
    rows = [list(r) for r in csv.reader(text.splitlines(), dialect)]
    width = max((len(r) for r in rows), default=0)
    return Sheet(path.stem, [r + [""] * (width - len(r)) for r in rows])


# ---------------------------------------------------------------------- rows -> compounds


def _strip_unit(text: str, units: str) -> str:
    return re.sub(rf"\s*(?:{units})\s*$", "", text.strip(), flags=re.I).strip()


def _nmr_head(nucleus: str, text: str, mhz: dict, solvent: str) -> tuple[str, bool]:
    """An NMR report with its "¹H NMR (500 MHz, CDCl3)" head (added when the cell starts
    at δ or with the first shift), and whether the head lacks the frequency or solvent."""
    t = clean_space(text)
    if S.NMR_HEAD.search(t[:30]):
        return t, False
    t = re.sub(r"^[:\s]*", "", t)
    if not t.startswith("δ"):
        t = "δ " + t
    freq = mhz.get(nucleus, "")
    inside = ", ".join(x for x in (f"{freq} MHz" if freq else "", solvent) if x)
    head = f"{nucleus} NMR" + (f" ({inside})" if inside else "")
    return f"{head} {t}", not (freq and solvent)


def header_conditions(header: str) -> tuple[str, str]:
    """(MHz, solvent) written in an NMR column's header: "1H NMR (CDCl3, 500 MHz)"."""
    mhz = re.search(r"(\d{2,4})\s*MHz", header, re.I)
    solvent = ""
    inside = re.search(r"\(([^()]*)\)", header)
    for part in (inside.group(1).split(",") if inside else []):
        part = part.strip()
        if part and not re.search(r"(?i)mhz|^\d", part):
            solvent = part
    return (mhz.group(1) if mhz else ""), solvent


def _derived_mhz(mhz: dict) -> dict:
    """The ¹³C and ¹⁹F frequencies from the ¹H one when not given (e.g. 500 → 126, 470)."""
    out = dict(mhz)
    try:
        h = float(out.get("1H", ""))
    except ValueError:
        return out
    for nuc, ratio in (("13C", 0.25145), ("19F", 0.94077), ("31P", 0.40481), ("11B", 0.32084)):
        out.setdefault(nuc, f"{h * ratio:.0f}")
    return out


def has_data(cells: list[str], roles: list[str], pictures: dict | None = None) -> bool:
    """Whether a row holds anything for a field (notes and ignored columns aside)."""
    return any(role not in ("ignore", "notes") and (text.strip() or
               (role == "picture" and pictures and i in pictures))
               for i, (role, text) in enumerate(zip(roles, cells)))


def ids_from(start: str, count: int, skip: set[str] = frozenset()) -> list[str]:
    """`count` IDs from `start` on: "3a" → 3a, 3b, … 3z, 3aa; "S1" → S1, S2; skipping the
    IDs in `skip`."""
    from .compound_ids import index_for_letters, letters_for
    start = start.strip()
    m = re.fullmatch(r"(.*?\d)([a-z]+)", start)
    n = re.fullmatch(r"(.*?)(\d+)", start)
    out: list[str] = []
    i = 0
    while len(out) < count and i < count + len(skip) + 1:
        if m:
            cid = m.group(1) + letters_for(index_for_letters(m.group(2)) + i)
        elif n:
            cid = f"{n.group(1)}{int(n.group(2)) + i}"
        else:
            cid = f"{start}{i + 1}"
        if cid not in skip:
            out.append(cid)
        i += 1
    return out


def _round_percent(t: str) -> str:
    """A computed yield (73.94792273562499) as an SI prints it (74); a typed one stays."""
    if re.fullmatch(r"\d+\.\d{3,}", t):
        return f"{float(t):.0f}"
    return t


def _times(t: str) -> tuple[str, str, bool]:
    """(t major, t minor, guessed) from "4.104 (maj), 4.434 (min)", "5.779 (min), 6.024
    (maj)", "17.140 (maj), n/a (min)" or "10.308  14.031" (guessed: first = major)."""
    maj = re.search(r"(\d+(?:\.\d+)?)\s*\(?\s*maj", t, re.I)
    mnr = re.search(r"(\d+(?:\.\d+)?)\s*\(?\s*min(?:or)?\b\s*\)?", t, re.I)
    if maj or mnr:
        return (maj.group(1) if maj else "", mnr.group(1) if mnr else "", False)
    nums = re.findall(r"\d+(?:\.\d+)?", t)
    return (nums[0] if nums else "", nums[1] if len(nums) > 1 else "", len(nums) > 1)


def compound_from_row(cells: list[str], markup: list[str], headers: list[str],
                      roles: list[str], picture: bytes | None = None,
                      auto_id: str = "") -> ImportedCompound | None:
    """One row as a compound (None for an empty row or one with no ID and no name).
    `picture`: the row's picture in the structure-picture column; `auto_id`: the ID given
    to a row without one (then a row with any data is a compound)."""
    by: dict[str, list[tuple[str, str, str]]] = {}
    for header, role, text, mk in zip(headers, roles, cells, markup):
        if role == "ignore" or not text.strip():
            continue
        by.setdefault(role, []).append((clean_space(text), mk, header))

    def one(role: str) -> str:
        return by[role][0][0] if role in by else ""

    cid, name = one("id"), one("name")
    if not cid and not name and not auto_id:
        return None
    if not cid and auto_id:
        cid = auto_id
    cid = re.sub(r"^\*\*|\*\*$", "", cid).strip()
    if re.fullmatch(r"\d+\.0", cid):
        cid = cid[:-2]
    c = ImportedCompound(cid=cid or "?", name=name)
    if picture and "smiles" not in by:
        c.picture_pdf, c.picture_ext = picture, _picture_ext(picture)
        c.structure_note = ("the picture in the Excel cell (low resolution: paste the structure "
                            "from ChemDraw later, for its formula and a sharp drawing)")
    if not cid:
        c.warnings.append("This row has no compound ID: give it one here or in the ID table "
                          "after the import.")
    v = c.values
    lines: list[str] = []          # (the row as a text, for "exactly as imported")

    # structure
    if "smiles" in by:
        from . import chem
        smiles = one("smiles")
        try:
            mol = chem.parse_structure(smiles)
            c.structure = chem.mol_to_block(mol)
            c.structure_note = f"from the SMILES in “{by['smiles'][0][2]}”"
        except Exception:
            c.warnings.append(f"The SMILES “{smiles}” could not be read; the compound has no "
                              "structure (paste one from ChemDraw after the import).")
            c.use_structure = False

    # a whole data paragraph first, so that single columns can fill what it lacks…
    for text, _mk, _h in by.get("data", []):
        S.parse_data(text, c)
        lines.append(text)
    if "procedure" in by:
        c.procedure = "\n".join(mk for _t, mk, _h in by["procedure"])

    def setpart(key: str, part: str, value: str) -> None:
        if value:
            v.setdefault(key, {})[part] = value

    if "physical_state" in by:
        sm = S._STATE.match(one("physical_state"))
        if sm:
            S._set_state(v, sm.group("state"), sm.group("inside") or "")
        else:
            v["physical_state"] = {"text": S._clean_state(one("physical_state"))}
    # yield and mass: "94", "94%", "0.94" (a % cell already gives 94), "57.4 mg (94%)"
    for role in ("yield_percent", "ee"):        # (0: not measured yet, as a formula gives)
        if role in by and re.fullmatch(r"0+(?:\.0+)?\s*%?", one(role)):
            del by[role]
    if "yield_percent" in by:
        t = one("yield_percent")
        if _is_number(t) and "%" not in t and float(t.replace(",", ".")) <= 1:
            t = f"{float(t.replace(',', '.')) * 100:.0f}"       # (0.73: a fraction)
        t = _round_percent(t)
        got, state, _more = S.parse_yield(t if "%" in t or "mg" in t else t + "%")
        y = v.setdefault("yield", {})
        for k, val in got.items():
            y[k] = val
        if not got:
            if re.search(r"(?i)quant", t):
                y["percent"] = "quant."
            else:
                c.unrecognised.append(f"Yield: {t}")
        if state:
            v.setdefault("physical_state", {"text": S._clean_state(state)})
    if "mass" in by:
        t, header = one("mass"), by["mass"][0][2]
        m = S._MASS.search(t)
        y = v.setdefault("yield", {})
        if m:
            y["mass"], y["mass_unit"] = m.group("v"), m.group("u").replace("μ", "µ")
        elif _is_number(t):
            y["mass"] = t
            y["mass_unit"] = "g" if re.search(r"\(\s*g\s*\)|\bin g\b|grams", header, re.I) else "mg"
        else:
            c.unrecognised.append(f"Mass: {t}")
    if "scale" in by:
        t = one("scale")
        m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*(mmol|mol|µmol|μmol|mg|g)?\b", t)
        if m:
            v["scale"] = {"value": m.group(1), "unit": (m.group(2) or "mmol").replace("μ", "µ")}
        else:
            c.unrecognised.append(f"Scale: {t}")

    # NMR
    mhz = {k[len("nmr_mhz_"):]: _strip_unit(one(k), "mhz") for k in by if k.startswith("nmr_mhz_")}
    mhz_typed = set(mhz)
    mhz = _derived_mhz(mhz)
    solvent = one("nmr_solvent")
    missing_head = False
    for nuc in ("1H", "13C", "19F", "31P", "11B"):
        for text, _mk, header in by.get(f"nmr_{nuc}", []):
            h_mhz, h_solvent = header_conditions(header)
            report, lacks = _nmr_head(nuc, text, {**mhz, **({nuc: h_mhz} if h_mhz and nuc not in
                                                             mhz_typed else {})},
                                      solvent or h_solvent)
            missing_head |= lacks
            if not report.rstrip().endswith("."):
                report = report.rstrip() + "."
            v.setdefault("nmr", []).append({"id": new_id("r"), "nucleus": nuc, "label": "",
                                            "text": report})
            lines.append(report)
    if missing_head:
        c.warnings.append("An NMR report had no frequency or solvent (in its cell or in an NMR "
                          "frequency / solvent column): add them to its report.")

    # HRMS
    if "hrms" in by:
        before = len(c.unrecognised)
        S._parse_item("hrms", one("hrms") if re.search(r"(?i)hr-?ms|m/z", one("hrms"))
                      else "HRMS " + one("hrms"), c, "")
        if len(c.unrecognised) > before or "hrms" not in v:
            c.unrecognised.append(f"HRMS: {one('hrms')}")
    for role, part in (("hrms_found", "found"), ("hrms_adduct", "adduct"),
                       ("hrms_ionization", "ionization"), ("hrms_formula", "formula")):
        t = one(role)
        if not t:
            continue
        if role == "hrms_found":
            t = S._number_after(t) or t
        elif role == "hrms_adduct":
            t = S.adduct_text(t) if re.search(r"M", t) else t
        elif role == "hrms_formula":
            t = re.sub(r"\s+", "", t)
        v.setdefault("hrms", {})[part] = t
    if "hrms_calcd" in by:
        c.hrms_calcd = S._number_after(one("hrms_calcd"))
    if "hrms" in v:
        # the ion and ionization written in a header ("HRMS (ESI) calcd [M+H]+")
        heads = " ".join(h for r in ("hrms_found", "hrms_calcd", "hrms")
                         for _t, _m, h in by.get(r, []))
        ion = S.ADDUCT.search(heads)
        if ion and "hrms_adduct" not in by and not v["hrms"].get("adduct"):
            v["hrms"]["adduct"] = S.adduct_text(ion.group(0))
        src = re.search(r"\b(ESI|APCI|EI|CI|FAB|MALDI|DART|APPI)\b", heads)
        if src and "hrms_ionization" not in by and not v["hrms"].get("ionization"):
            v["hrms"]["ionization"] = src.group(1)
    if "hrms" in v:
        v["hrms"].setdefault("ionization", "ESI" if "found" in v["hrms"] else "")
        v["hrms"].setdefault("adduct", "[M+H]+" if "found" in v["hrms"] else "")
        if ("hrms_adduct" not in by and not S.ADDUCT.search(one("hrms") + " " + heads)
                and "found" in v["hrms"]):
            c.warnings.append("HRMS: no ion was given; [M+H]+ was assumed. Check it.")

    # single values: the cell alone ("0.23") or with its label ("Rf = 0.23 (…)")
    simple = {"rf": ("Rf = ", S.parse_rf), "mp": ("mp ", S.parse_mp),
              "ir": ("IR ", S.parse_ir), "rotation": ("[α]D ", S.parse_rotation),
              "dr": ("dr ", S.parse_dr), "ea": ("Anal. ", S.parse_ea),
              "ccdc": ("CCDC ", S.parse_ccdc)}
    for role, (prefix, parser) in simple.items():
        if role not in by:
            continue
        t = one(role)
        key = "xray" if role == "ccdc" else role
        got = parser(t) or parser(prefix + t)
        if role == "dr" and not got and re.fullmatch(r"[>≥]?\s*\d+(?:\.\d+)?\s*:\s*\d+(?:\.\d+)?", t):
            got = {"value": re.sub(r"\s+", "", t)}
        if role == "ir" and not got and re.search(r"\d{3}", t):
            got = {"bands": t}
        if got:
            v[key] = {**v.get(key, {}), **got}
        else:
            c.unrecognised.append(f"{by[role][0][2]}: {t}")
    if "rf_eluent" in by:
        v.setdefault("rf", {})["eluent"] = one("rf_eluent")
    if "purification" in by:
        v.setdefault("purification", {})["eluent"] = one("purification")

    # ee / er
    if "ee" in by:
        t = one("ee")
        if _is_number(t) and "%" not in t and 0 < float(t.replace(",", ".")) <= 1:
            t = f"{float(t.replace(',', '.')) * 100:.0f}%"      # (0.96: a fraction)
        got = S.parse_ee(t) or S.parse_ee(t + ("" if "%" in t or ":" in t else "%") + " ee")
        if not got and re.fullmatch(r"\d+(?:\.\d+)?\s*:\s*\d+(?:\.\d+)?", t):
            got = {"value": re.sub(r"\s+", "", t) + " er"}
        if got:
            v["ee"] = {**v.get("ee", {}), **got}
        else:
            c.unrecognised.append(f"ee: {t}")
    for role, part in (("ee_method", "method"), ("ee_column", "column"), ("ee_eluent", "eluent"),
                       ("ee_flow", "flow"), ("ee_wavelength", "wavelength"),
                       ("t_major", "t_major"), ("t_minor", "t_minor")):
        t = one(role)
        if t:
            if part in ("flow", "wavelength", "t_major", "t_minor"):
                t = S._number_after(t) or t
            elif part == "eluent":
                # "IPA-Hex 2.5%, iso, 15 min, 220nm, F100": the wavelength to its own box
                nm = re.search(r"[,;]?\s*(?:λ\s*=?\s*)?(\d{3})\s*nm\b", t)
                if nm:
                    v.setdefault("ee", {}).setdefault("wavelength", nm.group(1))
                    t = clean_space(t[:nm.start()] + t[nm.end():]).strip(" ,;")
            v.setdefault("ee", {})[part] = t
    if "ee_times" in by:
        major, minor, guessed = _times(one("ee_times"))
        ee = v.setdefault("ee", {})
        if major:
            ee.setdefault("t_major", major)
        if minor:
            ee.setdefault("t_minor", minor)
        if guessed and ee.get("value"):
            c.warnings.append("ee: the retention times are not marked (maj)/(min); the first "
                              "was taken as the major enantiomer. Check them.")

    for text, _mk, header in by.get("notes", []):
        c.notes.append(f"{header}: {text}")
    head = f"{c.name} ({c.cid})." if c.name else f"{c.cid}."
    c.original = "\n".join([head] + ([c.procedure] if c.procedure else []) + lines)
    return c


def _picture_ext(data: bytes) -> str:
    if data[:4] == b"\x89PNG":
        return ".png"
    if data[:3] == b"\xff\xd8\xff":
        return ".jpg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return ".gif"
    return ".png"


def plan_from_table(sheet: Sheet, header_row: int, roles: list[str],
                    first_id: str = "") -> ImportPlan:
    """The compounds in the rows below `header_row`, with columns read as `roles`.
    `first_id`: rows without an ID are numbered from it ("3a", "3b"…), and every row with
    data is a compound."""
    headers = sheet.rows[header_row] if header_row < len(sheet.rows) else []
    markup = sheet.markup or sheet.rows
    plan = ImportPlan(section_name=sheet.name if sheet.name.lower() not in
                      ("sheet1", "sheet 1", "") else "")
    seen: dict[str, int] = {}
    pic_col = roles.index("picture") if "picture" in roles else -1
    id_col = roles.index("id") if "id" in roles else -1
    rows = range(header_row + 1, len(sheet.rows))
    fresh: list[str] = []
    if first_id.strip():
        typed = {sheet.rows[r][id_col].strip() for r in rows if id_col >= 0}
        need = sum(1 for r in rows if has_data(sheet.rows[r], roles,
                                               {pic_col: 1} if (r, pic_col) in sheet.pictures
                                               else None)
                   and not (id_col >= 0 and sheet.rows[r][id_col].strip()))
        fresh = ids_from(first_id, need, typed)
    for r in rows:
        picture = sheet.pictures.get((r, pic_col)) if pic_col >= 0 else None
        auto = ""
        if fresh and has_data(sheet.rows[r], roles, {pic_col: 1} if picture else None) \
                and not (id_col >= 0 and sheet.rows[r][id_col].strip()):
            auto = fresh.pop(0)
        c = compound_from_row(sheet.rows[r], markup[r], headers, roles, picture, auto)
        if c is None:
            continue
        c.page = r + 1           # (its row in the sheet, shown in messages)
        if c.cid in seen and c.cid != "?":
            c.warnings.append(f"Row {r + 1} has the same ID as row {seen[c.cid]}; only one of "
                              "them can keep it (untick the other).")
        seen.setdefault(c.cid, r + 1)
        plan.compounds.append(c)
    if not plan.compounds:
        plan.warnings.append("No rows with a compound ID or name were found below the header "
                             "row. Check the header row and which column holds the IDs.")
    if "id" not in roles and not first_id.strip():
        plan.warnings.append("No column was chosen as the compound ID.")
    elif first_id.strip():
        plan.warnings.append(f"Rows without an ID were numbered from {first_id.strip()}, in the "
                             "sheet's order. Change any ID here, or in the ID table after the "
                             "import.")
    return plan

