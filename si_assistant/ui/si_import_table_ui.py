"""File ▸ Import SI… with an Excel or CSV file (SPEC §13y): the window that says which
column holds what, before the usual review window.

The source is a sheet, or one of its Excel tables (Insert ▸ Table), so a sheet holding
several tables side by side or one under another can be read one table at a time. Each
column is listed with its header, a "Read as" choice (guessed from the header, or as chosen
last time for a column with the same header) and its first values. Rows without an ID can
be numbered automatically.
"""

from __future__ import annotations

import json

from PySide6.QtCore import QSettings, Qt
from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QHBoxLayout, QHeaderView,
                               QLabel, QLineEdit, QSpinBox, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget)

from .. import APP_NAME
from ..core import si_import_table as T

SETTINGS_KEY = "table_import/roles"     # {normalised header: role} chosen before
SOURCE_KEY = "table_import/sources"     # {file name: source label} chosen before


def _load(key: str) -> dict:
    try:
        raw = QSettings(APP_NAME, APP_NAME).value(key, "")
        data = json.loads(raw) if raw else {}
        return data if isinstance(data, dict) else {}
    except (ValueError, TypeError):
        return {}


def _save(key: str, data: dict) -> None:
    try:
        QSettings(APP_NAME, APP_NAME).setValue(key, json.dumps(data))
    except Exception:  # noqa: BLE001 — remembering is a convenience only
        pass


def _remembered() -> dict:
    return _load(SETTINGS_KEY)


def _remember(headers: list[str], roles: list[str]) -> None:
    data = _remembered()
    for h, r in zip(headers, roles):
        if h.strip():
            data[T._norm(h)] = r
    _save(SETTINGS_KEY, data)


def initial_roles(headers: list[str]) -> list[str]:
    """As chosen last time for the same headers, else guessed from the header."""
    known = _remembered()
    guessed = T.guess_roles(headers)
    return [known.get(T._norm(h), g) if known.get(T._norm(h)) in T.ROLES else g
            for h, g in zip(headers, guessed)]


def sources(read: T.TableRead) -> list[tuple[str, T.Sheet]]:
    """(label, sheet or table view) for every sheet and every Excel table on it."""
    out = []
    for s in read.sheets:
        if not s.rows:
            continue
        for a in s.areas:
            view = s.view(a)
            n = max(0, len(view.rows) - 1)
            out.append((f"{s.name} ▸ table “{a.name}” ({n} row{'s' if n != 1 else ''})", view))
        out.append((f"{s.name}" + (" (the whole sheet)" if s.areas else ""), s))
    return out


def _score(sheet: T.Sheet) -> int:
    """How much of a compound table a source looks like: its recognised columns."""
    if not sheet.rows:
        return 0
    roles = T.guess_roles(sheet.rows[sheet.header_row()])
    return sum(r not in ("ignore", "notes", "data") for r in roles) + 5 * ("id" in roles)


class ColumnsDialog(QDialog):
    """Which sheet or table, which row holds the headers, and what each column is."""

    def __init__(self, read: T.TableRead, parent: QWidget | None = None):
        super().__init__(parent)
        self.read = read
        self.sources = sources(read)
        self.setWindowTitle(f"Import from a spreadsheet — {read.file_name}")
        self.resize(1000, 680)
        intro = QLabel(
            "One row per compound. Choose the sheet (or one of its tables) and say what each "
            "column holds: the choices below were guessed from the headers and are "
            "remembered for next time. Columns set to “not imported” are left out. Next, the "
            "usual import window shows every compound as read, to check before anything is "
            "made.")
        intro.setWordWrap(True)

        self.source = QComboBox()
        for label, _s in self.sources:
            self.source.addItem(label)
        self.header = QSpinBox()
        self.header.setMinimum(1)
        top = QHBoxLayout()
        top.addWidget(QLabel("Read from:"))
        top.addWidget(self.source, 1)
        top.addSpacing(16)
        top.addWidget(QLabel("Headers in row:"))
        top.addWidget(self.header)

        self.first_id = QLineEdit()
        self.first_id.setPlaceholderText("e.g. 3a")
        self.first_id.setMaximumWidth(90)
        self.first_id.setToolTip("Rows with no ID (or with no ID column) get IDs from this "
                                 "one on, in the sheet's order: 3a, 3b, … 3z, 3aa. Leave empty "
                                 "to import only rows that have an ID or a name.")
        ids = QHBoxLayout()
        ids.addWidget(QLabel("Rows without an ID: number them from"))
        ids.addWidget(self.first_id)
        ids.addWidget(QLabel("(empty: skip them)"))
        ids.addStretch()

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Column header", "Read as", "First values"])
        self.table.verticalHeader().setVisible(False)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.setColumnWidth(0, 200)
        self.table.setColumnWidth(1, 300)
        self.combos: list[QComboBox] = []
        self.count = QLabel("")
        self.count.setWordWrap(True)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel)
        self.next_btn = buttons.addButton("Next", QDialogButtonBox.ButtonRole.AcceptRole)
        self.next_btn.setDefault(True)
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(intro)
        layout.addLayout(top)
        layout.addLayout(ids)
        layout.addWidget(self.table, 1)
        layout.addWidget(self.count)
        layout.addWidget(buttons)

        # the source chosen last time for this file, or the one most like a compound table
        last = _load(SOURCE_KEY).get(read.file_name, "")
        labels = [label for label, _s in self.sources]
        if last in labels:
            best = labels.index(last)
        else:
            scores = [_score(s) for _l, s in self.sources]
            best = scores.index(max(scores)) if scores else 0
        self.source.setCurrentIndex(best)
        self._source_changed()
        self.source.currentIndexChanged.connect(self._source_changed)
        self.header.valueChanged.connect(self._fill)
        self.first_id.textChanged.connect(self._update_count)

    # ------------------------------------------------------------------
    @property
    def current(self) -> T.Sheet:
        return self.sources[self.source.currentIndex()][1]

    def _source_changed(self) -> None:
        s = self.current
        self.header.blockSignals(True)
        self.header.setMaximum(max(1, len(s.rows)))
        self.header.setValue(s.header_row() + 1)
        self.header.blockSignals(False)
        self._fill()

    def headers(self) -> list[str]:
        s = self.current
        r = self.header.value() - 1
        return [h.strip() for h in s.rows[r]] if r < len(s.rows) else []

    def _fill(self) -> None:
        s = self.current
        headers = self.headers()
        roles = initial_roles(headers)
        self.table.setRowCount(len(headers))
        self.combos = []
        start = self.header.value()
        for i, (h, role) in enumerate(zip(headers, roles)):
            item = QTableWidgetItem(h or f"(column {i + 1}, no header)")
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, item)
            combo = QComboBox()
            for key, (label, help_text) in T.ROLES.items():
                combo.addItem(label, key)
                if help_text:
                    combo.setItemData(combo.count() - 1, help_text, Qt.ItemDataRole.ToolTipRole)
            combo.setCurrentIndex(max(0, combo.findData(role)))
            combo.currentIndexChanged.connect(self._update_count)
            self.combos.append(combo)
            self.table.setCellWidget(i, 1, combo)
            examples = []
            for r in range(start, len(s.rows)):
                if (r, i) in s.pictures:
                    examples.append("[picture]")
                elif i < len(s.rows[r]) and s.rows[r][i].strip():
                    examples.append(s.rows[r][i])
                if len(examples) == 3:
                    break
            ex = QTableWidgetItem("  |  ".join(e if len(e) < 60 else e[:57] + "…"
                                               for e in examples))
            ex.setToolTip("\n\n".join(examples))
            ex.setFlags(ex.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 2, ex)
        self._update_count()

    def roles(self) -> list[str]:
        return [c.currentData() for c in self.combos]

    def _rows_imported(self) -> int:
        s, roles = self.current, self.roles()
        auto = bool(self.first_id.text().strip())
        idx = [i for i, r in enumerate(roles) if r in ("id", "name")]
        pic = roles.index("picture") if "picture" in roles else -1
        n = 0
        for r in range(self.header.value(), len(s.rows)):
            row = s.rows[r]
            if any(i < len(row) and row[i].strip() for i in idx):
                n += 1
            elif auto and T.has_data(row, roles, {pic: 1} if (r, pic) in s.pictures else None):
                n += 1
        return n

    def _update_count(self) -> None:
        roles = self.roles()
        auto = bool(self.first_id.text().strip())
        has_key = any(r in ("id", "name") for r in roles)
        n = self._rows_imported()
        used = {r for r in roles if r != "ignore"}
        twice = sorted({T.ROLES[r][0] for r in roles if r not in T.MANY and roles.count(r) > 1})
        if has_key or auto:
            bits = [f"{n} row(s) will be imported."]
        else:
            bits = ["Choose the column with the compound IDs (or names), or number the rows "
                    "without an ID."]
        if "smiles" not in used:
            bits.append("No SMILES column: structures can be pasted from ChemDraw after the "
                        "import" + (" (the pictures are kept until then)." if "picture" in used
                                    else "."))
        if twice:
            bits.append("Chosen for more than one column (only the first is used): "
                        + ", ".join(twice) + ".")
        self.count.setText(" ".join(bits))
        self.next_btn.setEnabled((has_key or auto) and n > 0)

    def _accept(self) -> None:
        _remember(self.headers(), self.roles())
        sources_seen = _load(SOURCE_KEY)
        sources_seen[self.read.file_name] = self.sources[self.source.currentIndex()][0]
        _save(SOURCE_KEY, sources_seen)
        self.accept()

    def plan(self) -> T.ImportPlan:
        return T.plan_from_table(self.current, self.header.value() - 1, self.roles(),
                                 self.first_id.text())
