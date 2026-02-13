# main.py
from __future__ import annotations

import sys
import os
import json
import re
from pathlib import Path
from typing import List, Dict, Optional, Tuple

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QLineEdit,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QMessageBox,
    QProgressBar, QGroupBox, QCheckBox, QSpinBox, QDialog, QScrollArea
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from wb_fill import FillParams, fill_wb_template


APP_NAME = "Sunglasses SEO PRO"


# -------------------------------
# DATA DIR + SETTINGS
# -------------------------------
def app_data_dir() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    p = base / APP_NAME / "data"
    p.mkdir(parents=True, exist_ok=True)
    return p


def settings_path() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    p = base / APP_NAME
    p.mkdir(parents=True, exist_ok=True)
    return p / "settings.json"


def load_settings() -> Dict:
    sp = settings_path()
    if sp.exists():
        try:
            return json.loads(sp.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(d: Dict) -> None:
    sp = settings_path()
    sp.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def _norm_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("&", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


# -------------------------------
# LIST FILES
# -------------------------------
def list_file(path: Path, defaults: List[str]) -> List[str]:
    if not path.exists():
        path.write_text("\n".join(defaults) + "\n", encoding="utf-8")
        return defaults[:]
    lines = []
    for ln in path.read_text(encoding="utf-8").splitlines():
        ln = ln.strip()
        if ln:
            lines.append(ln)
    # ensure defaults included
    for d in defaults:
        if d not in lines:
            lines.append(d)
    return lines


def add_to_list_file(path: Path, value: str) -> None:
    value = (value or "").strip()
    if not value:
        return
    lines = []
    if path.exists():
        lines = [x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()]
    if value not in lines:
        lines.append(value)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# -------------------------------
# BRANDS RU MAP
# -------------------------------
def brands_ru_map_path() -> Path:
    return app_data_dir() / "brands_ru.json"


def load_brands_ru_map() -> Dict[str, str]:
    p = brands_ru_map_path()
    if not p.exists():
        p.write_text("{}", encoding="utf-8")
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_brands_ru_map(m: Dict[str, str]) -> None:
    brands_ru_map_path().write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")


def brand_to_ru(brand_lat: str, m: Dict[str, str]) -> str:
    k = _norm_key(brand_lat)
    # try exact
    if k in m:
        return m[k]
    # try direct
    if brand_lat in m:
        return m[brand_lat]
    return brand_lat  # fallback


# -------------------------------
# THEMES (UI like screenshot)
# -------------------------------
THEMES = {
    "Graphite": {
        "bg": "#0b0f17",
        "card": "#111827",
        "card2": "#0f172a",
        "text": "#e5e7eb",
        "muted": "#9ca3af",
        "accent": "#3b82f6",
        "accent2": "#7c3aed",
        "border": "#1f2937",
        "input": "#0b1220",
    },
    "Midnight": {
        "bg": "#070a12",
        "card": "#0b1220",
        "card2": "#0a1020",
        "text": "#e5e7eb",
        "muted": "#a3a3a3",
        "accent": "#2563eb",
        "accent2": "#6d28d9",
        "border": "#111827",
        "input": "#0a1020",
    },
    "Sepia": {
        "bg": "#0f0d0b",
        "card": "#15120f",
        "card2": "#1b1713",
        "text": "#f5f5f4",
        "muted": "#d6d3d1",
        "accent": "#f59e0b",
        "accent2": "#a855f7",
        "border": "#292524",
        "input": "#1b1713",
    }
}


def make_stylesheet(theme_name: str) -> str:
    t = THEMES.get(theme_name, THEMES["Graphite"])
    return f"""
    QWidget {{
        background: {t["bg"]};
        color: {t["text"]};
        font-family: "Segoe UI";
        font-size: 11pt;
    }}
    QGroupBox {{
        border: 1px solid {t["border"]};
        border-radius: 14px;
        margin-top: 10px;
        padding: 10px;
        background: {t["card"]};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 12px;
        padding: 2px 8px;
        color: {t["text"]};
        background: transparent;
        font-weight: 600;
    }}
    QLineEdit, QComboBox, QSpinBox {{
        background: {t["input"]};
        border: 1px solid {t["border"]};
        border-radius: 10px;
        padding: 10px 12px;
        min-height: 22px;
        selection-background-color: {t["accent"]};
    }}
    QComboBox::drop-down {{
        border: 0px;
        width: 26px;
        subcontrol-origin: padding;
        subcontrol-position: top right;
        margin-right: 6px;
    }}
    QComboBox::down-arrow {{
        width: 10px;
        height: 10px;
        image: none;
        border-left: 2px solid {t["muted"]};
        border-bottom: 2px solid {t["muted"]};
        transform: rotate(-45deg);
        margin-top: 2px;
    }}
    QPushButton {{
        background: {t["card2"]};
        border: 1px solid {t["border"]};
        border-radius: 12px;
        padding: 10px 14px;
        font-weight: 600;
    }}
    QPushButton:hover {{
        border: 1px solid {t["accent"]};
    }}
    QPushButton#Primary {{
        background: {t["accent"]};
        border: 1px solid {t["accent"]};
        color: white;
    }}
    QPushButton#Primary:hover {{
        background: {t["accent2"]};
        border: 1px solid {t["accent2"]};
    }}
    QPushButton#Plus {{
        background: {t["accent"]};
        border: 1px solid {t["accent"]};
        color: white;
        min-width: 46px;
        max-width: 46px;
    }}
    QPushButton#Plus:hover {{
        background: {t["accent2"]};
        border: 1px solid {t["accent2"]};
    }}
    QLabel#Muted {{
        color: {t["muted"]};
    }}
    QProgressBar {{
        border: 1px solid {t["border"]};
        border-radius: 12px;
        text-align: center;
        background: {t["card2"]};
        height: 18px;
    }}
    QProgressBar::chunk {{
        background: {t["accent2"]};
        border-radius: 12px;
    }}
    QCheckBox {{
        spacing: 10px;
    }}
    """


# -------------------------------
# Worker thread
# -------------------------------
class Worker(QThread):
    progress = pyqtSignal(int)
    done = pyqtSignal(list, int, str)
    fail = pyqtSignal(str)

    def __init__(self, params: FillParams):
        super().__init__()
        self.params = params

    def run(self):
        try:
            def cb(p: int):
                self.progress.emit(int(p))
            self.params.progress_callback = cb
            outs, total, rep = fill_wb_template(self.params)
            self.done.emit(outs, total, rep)
        except Exception as e:
            self.fail.emit(str(e))


# -------------------------------
# Holiday multi dialog
# -------------------------------
class HolidaysDialog(QDialog):
    def __init__(self, holidays: List[str], selected: List[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбрать праздники")
        self.setModal(True)
        self.resize(420, 520)
        self._selected = set(selected or [])

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # scroll with checkboxes
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(8, 8, 8, 8)
        v.setSpacing(8)

        self.checks: List[QCheckBox] = []
        for h in holidays:
            h = (h or "").strip()
            if not h:
                continue
            cb = QCheckBox(h)
            cb.setChecked(h in self._selected)
            self.checks.append(cb)
            v.addWidget(cb)
        v.addStretch(1)
        scroll.setWidget(w)
        root.addWidget(scroll)

        row = QHBoxLayout()
        self.btn_cancel = QPushButton("Отмена")
        self.btn_ok = QPushButton("OK")
        self.btn_ok.setObjectName("Primary")
        row.addWidget(self.btn_cancel)
        row.addWidget(self.btn_ok)
        root.addLayout(row)

        self.btn_cancel.clicked.connect(self.reject)
        self.btn_ok.clicked.connect(self._on_ok)

    def _on_ok(self):
        picked = [c.text().strip() for c in self.checks if c.isChecked()]
        self._picked = picked
        self.accept()

    def picked(self) -> List[str]:
        return getattr(self, "_picked", [])


# -------------------------------
# App UI
# -------------------------------
class App(QWidget):
    def __init__(self):
        super().__init__()

        self.data_dir = app_data_dir()

        # files
        self.brands_file = self.data_dir / "brands.txt"
        self.shapes_file = self.data_dir / "shapes.txt"
        self.lenses_file = self.data_dir / "lenses.txt"
        self.holidays_file = self.data_dir / "holidays.txt"

        # defaults (you can extend anytime)
        self.brands = list_file(self.brands_file, ["Dior", "Gucci", "Prada", "Cazal", "Ray-Ban", "Balenciaga"])
        self.shapes = list_file(self.shapes_file, ["Кошачий глаз", "Квадратные", "Овальные", "Круглые", "Прямоугольные", "Авиаторы", "Вайфареры"])
        self.lenses = list_file(self.lenses_file, ["UV400", "Поляризационные", "Фотохромные (хамелеон)", "Градиентные", "Зеркальные"])
        self.holidays = list_file(self.holidays_file, ["8 Марта", "14 Февраля", "Новый год", "23 Февраля", "День рождения", "Выпускной", "День матери"])

        self.brand_map = load_brands_ru_map()

        self.selected_holidays: List[str] = []

        self.xlsx_path: Optional[str] = None
        self.out_dir: str = ""

        self.settings = load_settings()

        self._build_ui()
        self._restore_settings()

        # window sizing – prevent “tiny UI”
        self.setMinimumSize(980, 680)

    # ---------- UI build ----------
    def _build_ui(self):
        self.setWindowTitle(APP_NAME)

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        # Header card
        header = QGroupBox()
        h = QVBoxLayout(header)
        h.setContentsMargins(16, 14, 16, 12)
        h.setSpacing(4)

        title = QLabel("😎  Sunglasses SEO PRO")
        title.setStyleSheet("font-size: 20pt; font-weight: 800;")
        sub = QLabel("Живые SEO-описания • Выпадающие списки • Прогресс • Темы • Anti-повторы PRO")
        sub.setObjectName("Muted")
        h.addWidget(title)
        h.addWidget(sub)
        root.addWidget(header)

        # Top controls card
        top = QGroupBox()
        tl = QGridLayout(top)
        tl.setContentsMargins(14, 14, 14, 14)
        tl.setHorizontalSpacing(10)
        tl.setVerticalSpacing(10)

        # Theme
        tl.addWidget(QLabel("🎨  Тема"), 0, 0)
        self.cmb_theme = QComboBox()
        self.cmb_theme.addItems(list(THEMES.keys()))
        self.cmb_theme.currentTextChanged.connect(self._apply_theme)
        tl.addWidget(self.cmb_theme, 0, 1)

        # Data path (read-only)
        tl.addWidget(QLabel("📁  Data"), 0, 2)
        self.ed_data = QLineEdit(str(self.data_dir))
        self.ed_data.setReadOnly(True)
        tl.addWidget(self.ed_data, 0, 3)

        self.btn_open_data = QPushButton("Папка")
        self.btn_open_data.clicked.connect(self._open_data_folder)
        tl.addWidget(self.btn_open_data, 0, 4)

        # Output folder
        tl.addWidget(QLabel("📦  Папка вывода"), 1, 0)
        self.ed_out = QLineEdit("")
        self.ed_out.setPlaceholderText("Выбери папку, куда сохранять пачкой")
        tl.addWidget(self.ed_out, 1, 1, 1, 3)
        self.btn_pick_out = QPushButton("Выбрать")
        self.btn_pick_out.clicked.connect(self._pick_out_dir)
        tl.addWidget(self.btn_pick_out, 1, 4)

        # Load XLSX
        self.btn_load = QPushButton("⬇️  Загрузить XLSX")
        self.btn_load.setObjectName("Primary")
        self.btn_load.clicked.connect(self._pick_xlsx)
        tl.addWidget(self.btn_load, 2, 0, 1, 1)
        self.lb_file = QLabel("Файл не выбран")
        self.lb_file.setObjectName("Muted")
        tl.addWidget(self.lb_file, 2, 1, 1, 4)

        root.addWidget(top)

        # Main form card
        form = QGroupBox()
        gl = QGridLayout(form)
        gl.setContentsMargins(14, 14, 14, 14)
        gl.setHorizontalSpacing(10)
        gl.setVerticalSpacing(10)

        row = 0

        # Brand (latin) + plus
        gl.addWidget(QLabel("Бренд (латиницей)"), row, 0)
        self.cmb_brand = QComboBox()
        self.cmb_brand.setEditable(True)
        self.cmb_brand.addItems(self.brands)
        gl.addWidget(self.cmb_brand, row, 1, 1, 3)
        self.btn_add_brand = QPushButton("+")
        self.btn_add_brand.setObjectName("Plus")
        self.btn_add_brand.clicked.connect(lambda: self._add_item("brand"))
        gl.addWidget(self.btn_add_brand, row, 4)
        row += 1

        # Shape + plus
        gl.addWidget(QLabel("Форма оправы"), row, 0)
        self.cmb_shape = QComboBox()
        self.cmb_shape.setEditable(True)
        self.cmb_shape.addItems(self.shapes)
        gl.addWidget(self.cmb_shape, row, 1, 1, 3)
        self.btn_add_shape = QPushButton("+")
        self.btn_add_shape.setObjectName("Plus")
        self.btn_add_shape.clicked.connect(lambda: self._add_item("shape"))
        gl.addWidget(self.btn_add_shape, row, 4)
        row += 1

        # Lenses + plus
        gl.addWidget(QLabel("Линзы"), row, 0)
        self.cmb_lenses = QComboBox()
        self.cmb_lenses.setEditable(True)
        self.cmb_lenses.addItems(self.lenses)
        gl.addWidget(self.cmb_lenses, row, 1, 1, 3)
        self.btn_add_lenses = QPushButton("+")
        self.btn_add_lenses.setObjectName("Plus")
        self.btn_add_lenses.clicked.connect(lambda: self._add_item("lenses"))
        gl.addWidget(self.btn_add_lenses, row, 4)
        row += 1

        # Collection (editable)
        gl.addWidget(QLabel("Коллекция"), row, 0)
        self.cmb_collection = QComboBox()
        self.cmb_collection.setEditable(True)
        self.cmb_collection.addItems(["Весна–Лето 2026", "Весна–Лето 2025–2026"])
        gl.addWidget(self.cmb_collection, row, 1, 1, 4)
        row += 1

        # Holidays (multi)
        gl.addWidget(QLabel("Праздники (в описание)"), row, 0)
        self.ed_holidays = QLineEdit("")
        self.ed_holidays.setReadOnly(True)
        self.ed_holidays.setPlaceholderText("Нажми «Выбрать» и отметь праздники")
        gl.addWidget(self.ed_holidays, row, 1, 1, 2)
        self.btn_holidays = QPushButton("Выбрать")
        self.btn_holidays.clicked.connect(self._pick_holidays)
        gl.addWidget(self.btn_holidays, row, 3)

        self.cmb_holiday_pos = QComboBox()
        self.cmb_holiday_pos.addItems(["middle", "start", "end"])
        gl.addWidget(self.cmb_holiday_pos, row, 4)
        row += 1

        # SEO + style + brand ratio
        gl.addWidget(QLabel("SEO-плотность"), row, 0)
        self.cmb_seo = QComboBox()
        self.cmb_seo.addItems(["low", "normal", "high"])
        gl.addWidget(self.cmb_seo, row, 1)

        gl.addWidget(QLabel("Стиль"), row, 2)
        self.cmb_style = QComboBox()
        self.cmb_style.addItems(["neutral", "premium", "mass", "social"])
        gl.addWidget(self.cmb_style, row, 3)

        self.cmb_brand_ratio = QComboBox()
        self.cmb_brand_ratio.addItems(["50/50", "100/0", "0/100"])
        gl.addWidget(self.cmb_brand_ratio, row, 4)
        row += 1

        # Rows to fill + batch count
        gl.addWidget(QLabel("Строк заполнять"), row, 0)
        self.spin_rows = QSpinBox()
        self.spin_rows.setRange(1, 1000)
        self.spin_rows.setValue(6)
        gl.addWidget(self.spin_rows, row, 1)

        gl.addWidget(QLabel("Сколько Excel файлов"), row, 2)
        self.spin_batch = QSpinBox()
        self.spin_batch.setRange(1, 50)
        self.spin_batch.setValue(1)
        gl.addWidget(self.spin_batch, row, 3)

        gl.addWidget(QLabel("Уникализация"), row, 4)
        self.spin_uni = QSpinBox()
        self.spin_uni.setRange(0, 100)
        self.spin_uni.setValue(92)
        gl.addWidget(self.spin_uni, row, 5)
        row += 1

        # Skip first rows
        gl.addWidget(QLabel("Не трогать первые строк"), row, 0)
        self.spin_skip = QSpinBox()
        self.spin_skip.setRange(0, 50)
        self.spin_skip.setValue(4)
        gl.addWidget(self.spin_skip, row, 1)
        row += 1

        # WB modes
        self.chk_safe = QCheckBox("WB Safe Mode (заменяет риск-слова)")
        self.chk_strict = QCheckBox("WB Strict (убирает обещания/абсолюты/стоп-фразы)")
        self.chk_safe.setChecked(True)
        self.chk_strict.setChecked(True)
        gl.addWidget(self.chk_safe, row, 0, 1, 3)
        gl.addWidget(self.chk_strict, row, 3, 1, 3)
        row += 1

        root.addWidget(form)

        # Footer progress + generate
        foot = QGroupBox()
        fl = QHBoxLayout(foot)
        fl.setContentsMargins(14, 12, 14, 12)
        fl.setSpacing(10)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        fl.addWidget(self.progress, 1)

        self.btn_go = QPushButton("🚀  СГЕНЕРИРОВАТЬ")
        self.btn_go.setObjectName("Primary")
        self.btn_go.clicked.connect(self._run)
        fl.addWidget(self.btn_go, 0)

        root.addWidget(foot)

        # Apply theme now
        self._apply_theme(self.cmb_theme.currentText())

    # ---------- Theme ----------
    def _apply_theme(self, name: str):
        self.setStyleSheet(make_stylesheet(name))
        self.settings["theme"] = name
        save_settings(self.settings)

    # ---------- Data folder ----------
    def _open_data_folder(self):
        try:
            os.startfile(str(self.data_dir))  # windows
        except Exception:
            QMessageBox.information(self, "Data", str(self.data_dir))

    # ---------- Output dir ----------
    def _pick_out_dir(self):
        p = QFileDialog.getExistingDirectory(self, "Выбери папку вывода", self.ed_out.text().strip() or str(Path.home()))
        if p:
            self.ed_out.setText(p)
            self.settings["out_dir"] = p
            save_settings(self.settings)

    # ---------- XLSX ----------
    def _pick_xlsx(self):
        p, _ = QFileDialog.getOpenFileName(self, "Выбери XLSX", str(Path.home()), "Excel (*.xlsx)")
        if p:
            self.xlsx_path = p
            self.lb_file.setText(Path(p).name)
            self.settings["last_xlsx"] = p
            save_settings(self.settings)

    # ---------- Add items ----------
    def _add_item(self, kind: str):
        if kind == "brand":
            value = self.cmb_brand.currentText().strip()
            if not value:
                QMessageBox.warning(self, "Бренд", "Введи бренд и нажми +")
                return
            add_to_list_file(self.brands_file, value)
            self.brands = list_file(self.brands_file, self.brands)
            self._reload_combo(self.cmb_brand, self.brands, value)

            # ask RU brand for title
            ru = QMessageBox.question(
                self, "Кириллица",
                f"Добавить кириллицу для бренда '{value}'?\n\nЕсли да — введи русское написание в следующем окне.",
                QMessageBox.Yes | QMessageBox.No
            )
            if ru == QMessageBox.Yes:
                from PyQt5.QtWidgets import QInputDialog
                ru_val, ok = QInputDialog.getText(self, "Бренд на кириллице", f"{value} → (например: Миу Миу)")
                if ok and ru_val.strip():
                    m = load_brands_ru_map()
                    m[_norm_key(value)] = ru_val.strip()
                    save_brands_ru_map(m)
                    self.brand_map = m

        elif kind == "shape":
            value = self.cmb_shape.currentText().strip()
            if not value:
                QMessageBox.warning(self, "Форма", "Введи форму и нажми +")
                return
            add_to_list_file(self.shapes_file, value)
            self.shapes = list_file(self.shapes_file, self.shapes)
            self._reload_combo(self.cmb_shape, self.shapes, value)

        elif kind == "lenses":
            value = self.cmb_lenses.currentText().strip()
            if not value:
                QMessageBox.warning(self, "Линзы", "Введи линзы и нажми +")
                return
            add_to_list_file(self.lenses_file, value)
            self.lenses = list_file(self.lenses_file, self.lenses)
            self._reload_combo(self.cmb_lenses, self.lenses, value)

    def _reload_combo(self, cmb: QComboBox, items: List[str], select: str):
        cur = select
        cmb.blockSignals(True)
        cmb.clear()
        cmb.addItems(items)
        cmb.setCurrentText(cur)
        cmb.blockSignals(False)

    # ---------- Holidays ----------
    def _pick_holidays(self):
        dlg = HolidaysDialog(self.holidays, self.selected_holidays, self)
        if dlg.exec_() == QDialog.Accepted:
            self.selected_holidays = dlg.picked()
            self._sync_holidays_ui()
            self.settings["holidays_multi"] = self.selected_holidays
            save_settings(self.settings)

    def _sync_holidays_ui(self):
        if not self.selected_holidays:
            self.ed_holidays.setText("")
            return
        self.ed_holidays.setText(", ".join(self.selected_holidays))

    # ---------- Run ----------
    def _run(self):
        # validate xlsx
        if not self.xlsx_path or not Path(self.xlsx_path).exists():
            QMessageBox.warning(self, "XLSX", "Сначала выбери XLSX файл")
            return

        out_dir = self.ed_out.text().strip()
        if not out_dir:
            # default: рядом с файлом
            out_dir = str(Path(self.xlsx_path).parent)
            self.ed_out.setText(out_dir)

        brand_lat = self.cmb_brand.currentText().strip()
        if not brand_lat:
            QMessageBox.warning(self, "Бренд", "Введи/выбери бренд")
            return

        # title brand RU; desc brand LAT
        self.brand_map = load_brands_ru_map()
        brand_ru = brand_to_ru(brand_lat, self.brand_map)

        shape = self.cmb_shape.currentText().strip()
        lenses = self.cmb_lenses.currentText().strip()
        collection = self.cmb_collection.currentText().strip()

        holidays = "||".join([h.strip() for h in self.selected_holidays if h.strip()])

        params = FillParams(
            xlsx_path=self.xlsx_path,
            output_dir=out_dir,

            brand_lat=brand_lat,
            brand_ru=brand_ru,
            shape=shape,
            lenses=lenses,
            collection=collection,

            holidays=holidays,
            holiday_pos=self.cmb_holiday_pos.currentText().strip(),

            seo_level=self.cmb_seo.currentText().strip(),
            style=self.cmb_style.currentText().strip(),
            wb_safe_mode=self.chk_safe.isChecked(),
            wb_strict=self.chk_strict.isChecked(),

            brand_in_title_ratio=self.cmb_brand_ratio.currentText().strip(),
            rows_to_fill=int(self.spin_rows.value()),
            skip_first_rows=int(self.spin_skip.value()),
            batch_count=int(self.spin_batch.value()),

            uniqueness=int(self.spin_uni.value()),
        )

        # persist quick
        self._persist_current()

        # UI lock
        self.btn_go.setEnabled(False)
        self.progress.setValue(0)

        self.worker = Worker(params)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.done.connect(self._on_done)
        self.worker.fail.connect(self._on_fail)
        self.worker.start()

    def _on_done(self, outs: list, total: int, report: str):
        self.btn_go.setEnabled(True)
        self.progress.setValue(100)

        msg = f"Готово ✅\n\nФайлов: {len(outs)}\nСтрок заполнено: {total}\n\n"
        msg += "Выход:\n" + "\n".join(outs[:8]) + ("\n..." if len(outs) > 8 else "")
        QMessageBox.information(self, "Готово", msg)

    def _on_fail(self, err: str):
        self.btn_go.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", err)

    # ---------- Persist / Restore ----------
    def _persist_current(self):
        self.settings["theme"] = self.cmb_theme.currentText()
        self.settings["out_dir"] = self.ed_out.text().strip()
        self.settings["brand"] = self.cmb_brand.currentText().strip()
        self.settings["shape"] = self.cmb_shape.currentText().strip()
        self.settings["lenses"] = self.cmb_lenses.currentText().strip()
        self.settings["collection"] = self.cmb_collection.currentText().strip()
        self.settings["holiday_pos"] = self.cmb_holiday_pos.currentText().strip()
        self.settings["seo"] = self.cmb_seo.currentText().strip()
        self.settings["style"] = self.cmb_style.currentText().strip()
        self.settings["brand_ratio"] = self.cmb_brand_ratio.currentText().strip()
        self.settings["rows"] = int(self.spin_rows.value())
        self.settings["batch"] = int(self.spin_batch.value())
        self.settings["skip"] = int(self.spin_skip.value())
        self.settings["uni"] = int(self.spin_uni.value())
        self.settings["safe"] = bool(self.chk_safe.isChecked())
        self.settings["strict"] = bool(self.chk_strict.isChecked())
        self.settings["holidays_multi"] = self.selected_holidays
        save_settings(self.settings)

    def _restore_settings(self):
        # theme first
        theme = self.settings.get("theme", "Graphite")
        if theme in THEMES:
            self.cmb_theme.setCurrentText(theme)
        self._apply_theme(self.cmb_theme.currentText())

        out_dir = self.settings.get("out_dir", "")
        if out_dir:
            self.ed_out.setText(out_dir)

        def set_combo(cmb: QComboBox, key: str):
            v = (self.settings.get(key) or "").strip()
            if v:
                cmb.setCurrentText(v)

        set_combo(self.cmb_brand, "brand")
        set_combo(self.cmb_shape, "shape")
        set_combo(self.cmb_lenses, "lenses")
        set_combo(self.cmb_collection, "collection")
        set_combo(self.cmb_holiday_pos, "holiday_pos")
        set_combo(self.cmb_seo, "seo")
        set_combo(self.cmb_style, "style")
        set_combo(self.cmb_brand_ratio, "brand_ratio")

        self.spin_rows.setValue(int(self.settings.get("rows", 6)))
        self.spin_batch.setValue(int(self.settings.get("batch", 1)))
        self.spin_skip.setValue(int(self.settings.get("skip", 4)))
        self.spin_uni.setValue(int(self.settings.get("uni", 92)))

        self.chk_safe.setChecked(bool(self.settings.get("safe", True)))
        self.chk_strict.setChecked(bool(self.settings.get("strict", True)))

        saved_h = self.settings.get("holidays_multi", [])
        if isinstance(saved_h, list):
            self.selected_holidays = [str(x) for x in saved_h if str(x).strip()]
        else:
            self.selected_holidays = []
        self._sync_holidays_ui()

        last_xlsx = self.settings.get("last_xlsx", "")
        if last_xlsx and Path(last_xlsx).exists():
            self.xlsx_path = last_xlsx
            self.lb_file.setText(Path(last_xlsx).name)


def main():
    # Fix tiny UI on Windows High DPI
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    w = App()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
