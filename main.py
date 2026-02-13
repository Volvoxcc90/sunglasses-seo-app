# main.py
from __future__ import annotations

import os
import re
import json
import time
import random
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QLineEdit,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QMessageBox,
    QProgressBar, QSpinBox, QDoubleSpinBox
)

from wb_fill import FillParams, fill_wb_template


APP_NAME = "Sunglasses SEO PRO"


# ----------------------------
# Paths / data
# ----------------------------
def app_data_dir() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    p = base / APP_NAME / "data"
    p.mkdir(parents=True, exist_ok=True)
    return p


DATA_DIR = app_data_dir()
BRANDS_FILE = DATA_DIR / "brands.txt"          # латиница список
BRANDS_RU_FILE = DATA_DIR / "brands_ru.json"   # {"dior":"Диор", ...}
SHAPES_FILE = DATA_DIR / "shapes.txt"
LENSES_FILE = DATA_DIR / "lenses.txt"
COLLECTIONS_FILE = DATA_DIR / "collections.txt"
SETTINGS_FILE = DATA_DIR / "settings.json"


def _norm_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("&", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def load_list(file: Path, defaults: List[str]) -> List[str]:
    if not file.exists():
        file.write_text("\n".join(defaults), encoding="utf-8")
        return defaults[:]
    items = []
    for line in file.read_text(encoding="utf-8", errors="ignore").splitlines():
        t = line.strip()
        if t:
            items.append(t)
    return items or defaults[:]


def save_list_append(file: Path, value: str) -> None:
    value = (value or "").strip()
    if not value:
        return
    existing = set(load_list(file, []))
    if value in existing:
        return
    with file.open("a", encoding="utf-8") as f:
        if file.stat().st_size > 0:
            f.write("\n")
        f.write(value)


def load_brand_ru() -> Dict[str, str]:
    if not BRANDS_RU_FILE.exists():
        BRANDS_RU_FILE.write_text("{}", encoding="utf-8")
        return {}
    try:
        return json.loads(BRANDS_RU_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_brand_ru_map(m: Dict[str, str]) -> None:
    BRANDS_RU_FILE.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")


def load_settings() -> Dict:
    if not SETTINGS_FILE.exists():
        return {}
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(s: Dict) -> None:
    SETTINGS_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")


# ----------------------------
# Themes (UI как на скрине)
# ----------------------------
THEMES = {
    "Graphite": r"""
        QWidget { background:#0b0f17; color:#eaf0ff; }
        QLabel#Title { font-size:28px; font-weight:800; }
        QLabel#Subtitle { color:#9fb2d7; font-size:12px; }
        QLabel#Section { font-size:12px; color:#aab7d3; }
        QWidget#Card {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #141a24, stop:1 #0f141d);
            border:1px solid #1f2b46;
            border-radius:16px;
        }
        QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
            background:#0b1426;
            border:1px solid #1f2b46;
            border-radius:10px;
            padding:6px 12px;
            min-height:34px;
        }
        QComboBox::drop-down { border:0px; width:34px; }
        QComboBox::down-arrow { image:none; }
        QComboBox QAbstractItemView {
            background:#0b1426;
            border:1px solid #1f2b46;
            selection-background-color:#2a3a5f;
            padding:6px;
        }
        QPushButton {
            background:#22324f;
            border:1px solid #2a3a5f;
            padding:10px 14px;
            border-radius:12px;
            font-weight:700;
        }
        QPushButton:hover { background:#2a3a5f; }
        QPushButton#Primary {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #7a2cff, stop:1 #2d7cff);
            border:0px;
            padding:12px 18px;
            border-radius:14px;
            font-size:13px;
            font-weight:900;
        }
        QPushButton#Small {
            background:#2d7cff;
            border:0px;
            border-radius:12px;
            min-width:44px;
            min-height:34px;
            font-weight:900;
        }
        QPushButton#Ghost {
            background:#1a2232;
            border:1px solid #2a3a5f;
        }
        QProgressBar {
            border:1px solid #1f2b46;
            border-radius:12px;
            text-align:center;
            background:#0b1426;
            height:22px;
        }
        QProgressBar::chunk {
            border-radius:12px;
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #7a2cff, stop:1 #2d7cff);
        }
    """,
    "Midnight": r"""
        QWidget { background:#05070d; color:#eaf0ff; }
        QLabel#Title { font-size:28px; font-weight:800; }
        QLabel#Subtitle { color:#90a3c8; font-size:12px; }
        QLabel#Section { font-size:12px; color:#aab7d3; }
        QWidget#Card {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #0f1220, stop:1 #070a12);
            border:1px solid #1f2b46;
            border-radius:16px;
        }
        QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
            background:#0b1426;
            border:1px solid #1f2b46;
            border-radius:10px;
            padding:6px 12px;
            min-height:34px;
        }
        QComboBox::drop-down { border:0px; width:34px; }
        QComboBox::down-arrow { image:none; }
        QComboBox QAbstractItemView {
            background:#0b1426;
            border:1px solid #1f2b46;
            selection-background-color:#2a3a5f;
            padding:6px;
        }
        QPushButton {
            background:#1a2232;
            border:1px solid #2a3a5f;
            padding:10px 14px;
            border-radius:12px;
            font-weight:700;
        }
        QPushButton:hover { background:#2a3a5f; }
        QPushButton#Primary {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #7a2cff, stop:1 #2d7cff);
            border:0px;
            padding:12px 18px;
            border-radius:14px;
            font-size:13px;
            font-weight:900;
        }
        QPushButton#Small {
            background:#2d7cff;
            border:0px;
            border-radius:12px;
            min-width:44px;
            min-height:34px;
            font-weight:900;
        }
        QPushButton#Ghost { background:#0b1426; border:1px solid #1f2b46; }
        QProgressBar {
            border:1px solid #1f2b46;
            border-radius:12px;
            text-align:center;
            background:#0b1426;
            height:22px;
        }
        QProgressBar::chunk {
            border-radius:12px;
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #7a2cff, stop:1 #2d7cff);
        }
    """,
}


# ----------------------------
# Worker thread
# ----------------------------
class Worker(QThread):
    progress = pyqtSignal(int)
    done = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(
        self,
        xlsx_path: str,
        out_dir: str,
        files_count: int,
        rows_to_fill: int,
        brand_lat: str,
        brand_ru: str,
        shape: str,
        lens: str,
        collection: str,
        holiday: str,
        seo_level: str,
        style: str,
        gender: str,
        brand_ratio: float,
    ):
        super().__init__()
        self.xlsx_path = xlsx_path
        self.out_dir = out_dir
        self.files_count = files_count
        self.rows_to_fill = rows_to_fill
        self.brand_lat = brand_lat
        self.brand_ru = brand_ru
        self.shape = shape
        self.lens = lens
        self.collection = collection
        self.holiday = holiday
        self.seo_level = seo_level
        self.style = style
        self.gender = gender
        self.brand_ratio = brand_ratio

    def run(self):
        try:
            src = Path(self.xlsx_path)
            out_dir = Path(self.out_dir)
            out_dir.mkdir(parents=True, exist_ok=True)

            used_openers_global: Set[str] = set()  # анти-монотонность между файлами пачки

            total_steps = max(1, self.files_count)
            for i in range(self.files_count):
                out_name = f"{src.stem}_SEO_{i+1}.xlsx"
                out_path = str(out_dir / out_name)

                seed = int(time.time() * 1000) + i * 777

                params = FillParams(
                    xlsx_path=str(src),
                    out_path=out_path,
                    brand_lat=self.brand_lat,
                    brand_ru=self.brand_ru,
                    shape=self.shape,
                    lens=self.lens,
                    collection=self.collection,
                    seo_level=self.seo_level,
                    style=self.style,
                    gender=self.gender,
                    holiday=self.holiday,
                    rows_to_fill=self.rows_to_fill,
                    skip_top_rows=4,
                    brand_in_title_ratio=self.brand_ratio,
                    seed=seed,
                )

                def cb(p):
                    # прогресс внутри файла + прогресс по пачке
                    base = int(i / total_steps * 100)
                    add = int(p / total_steps)
                    self.progress.emit(min(100, base + add))

                fill_wb_template(params, used_openers_global=used_openers_global, progress_callback=cb)

                self.progress.emit(int((i + 1) / total_steps * 100))

            self.done.emit(str(out_dir))
        except Exception as e:
            self.failed.emit(str(e))


# ----------------------------
# Main UI
# ----------------------------
class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumSize(1100, 680)

        # крупнее, чтобы не было "маленькое окно/буквы"
        self.setFont(QFont("Segoe UI", 10))

        self.settings = load_settings()

        # data
        self.brands = load_list(BRANDS_FILE, ["Dior", "Gucci", "Prada", "Ray-Ban", "Cazal", "Miu Miu"])
        self.shapes = load_list(SHAPES_FILE, ["Квадратные", "Овальные", "Кошачий глаз", "Авиаторы", "Вайфареры", "Круглые"])
        self.lenses = load_list(LENSES_FILE, ["UV400", "Поляризационные", "Фотохромные", "Градиентные", "Зеркальные"])
        self.collections = load_list(COLLECTIONS_FILE, ["Весна–Лето 2026", "Весна–Лето 2025–2026"])
        self.brand_ru_map = load_brand_ru()

        self.xlsx_path: Optional[str] = None
        self.out_dir: Optional[str] = self.settings.get("out_dir")

        self._build_ui()
        self._apply_theme(self.settings.get("theme", "Graphite"))

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        # Header card
        card_header = QWidget()
        card_header.setObjectName("Card")
        lay_h = QVBoxLayout(card_header)
        lay_h.setContentsMargins(18, 14, 18, 14)

        title = QLabel("🕶️  Sunglasses SEO PRO")
        title.setObjectName("Title")
        subtitle = QLabel("Заполняет Наименование + Описание (ровно 6 строк), с живым SEO и реальным рандомом")
        subtitle.setObjectName("Subtitle")

        lay_h.addWidget(title)
        lay_h.addWidget(subtitle)
        root.addWidget(card_header)

        # Top controls card
        card_top = QWidget()
        card_top.setObjectName("Card")
        top = QGridLayout(card_top)
        top.setContentsMargins(18, 16, 18, 16)
        top.setHorizontalSpacing(12)
        top.setVerticalSpacing(12)

        # Theme
        top.addWidget(QLabel("🎨  Тема"), 0, 0)
        self.cmb_theme = QComboBox()
        self.cmb_theme.addItems(list(THEMES.keys()))
        self.cmb_theme.currentTextChanged.connect(self._on_theme_changed)
        top.addWidget(self.cmb_theme, 0, 1)

        # Data folder
        top.addWidget(QLabel("📁  Data"), 0, 2)
        self.ed_data = QLineEdit(str(DATA_DIR))
        self.ed_data.setReadOnly(True)
        top.addWidget(self.ed_data, 0, 3)

        self.btn_data = QPushButton("Папка")
        self.btn_data.setObjectName("Ghost")
        self.btn_data.clicked.connect(self._open_data_folder)
        top.addWidget(self.btn_data, 0, 4)

        # Output folder
        top.addWidget(QLabel("💾  Папка вывода"), 1, 0)
        self.ed_out = QLineEdit(self.out_dir or "")
        self.ed_out.setPlaceholderText("Выбери папку, куда сохранять пачкой")
        self.ed_out.setReadOnly(True)
        top.addWidget(self.ed_out, 1, 1, 1, 3)

        self.btn_out = QPushButton("Выбрать")
        self.btn_out.setObjectName("Ghost")
        self.btn_out.clicked.connect(self._choose_out_dir)
        top.addWidget(self.btn_out, 1, 4)

        # Load xlsx
        self.btn_load = QPushButton("⬇️  Загрузить XLSX")
        self.btn_load.clicked.connect(self._choose_xlsx)
        top.addWidget(self.btn_load, 2, 0, 1, 1)

        self.lb_xlsx = QLabel("Файл не выбран")
        self.lb_xlsx.setObjectName("Section")
        top.addWidget(self.lb_xlsx, 2, 1, 1, 4)

        root.addWidget(card_top)

        # Params card
        card_params = QWidget()
        card_params.setObjectName("Card")
        grid = QGridLayout(card_params)
        grid.setContentsMargins(18, 16, 18, 16)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)

        r = 0

        # Brand (latin) — editable combo
        grid.addWidget(QLabel("Бренд (латиницей)"), r, 0)
        self.cmb_brand = QComboBox()
        self.cmb_brand.setEditable(True)
        self.cmb_brand.addItems(self.brands)
        self.cmb_brand.setCurrentText(self.settings.get("brand", self.brands[0] if self.brands else "Dior"))
        grid.addWidget(self.cmb_brand, r, 1, 1, 2)

        self.btn_add_brand = QPushButton("+")
        self.btn_add_brand.setObjectName("Small")
        self.btn_add_brand.clicked.connect(self._add_brand)
        grid.addWidget(self.btn_add_brand, r, 3)

        r += 1

        grid.addWidget(QLabel("Форма оправы"), r, 0)
        self.cmb_shape = QComboBox()
        self.cmb_shape.setEditable(True)
        self.cmb_shape.addItems(self.shapes)
        self.cmb_shape.setCurrentText(self.settings.get("shape", self.shapes[0] if self.shapes else "Квадратные"))
        grid.addWidget(self.cmb_shape, r, 1, 1, 2)

        self.btn_add_shape = QPushButton("+")
        self.btn_add_shape.setObjectName("Small")
        self.btn_add_shape.clicked.connect(self._add_shape)
        grid.addWidget(self.btn_add_shape, r, 3)

        r += 1

        grid.addWidget(QLabel("Линзы"), r, 0)
        self.cmb_lens = QComboBox()
        self.cmb_lens.setEditable(True)
        self.cmb_lens.addItems(self.lenses)
        self.cmb_lens.setCurrentText(self.settings.get("lens", self.lenses[0] if self.lenses else "UV400"))
        grid.addWidget(self.cmb_lens, r, 1, 1, 2)

        self.btn_add_lens = QPushButton("+")
        self.btn_add_lens.setObjectName("Small")
        self.btn_add_lens.clicked.connect(self._add_lens)
        grid.addWidget(self.btn_add_lens, r, 3)

        r += 1

        grid.addWidget(QLabel("Коллекция"), r, 0)
        self.cmb_collection = QComboBox()
        self.cmb_collection.setEditable(True)
        self.cmb_collection.addItems(self.collections)
        self.cmb_collection.setCurrentText(self.settings.get("collection", self.collections[0] if self.collections else "Весна–Лето 2026"))
        grid.addWidget(self.cmb_collection, r, 1, 1, 3)

        r += 1

        grid.addWidget(QLabel("Праздник (в описание)"), r, 0)
        self.cmb_holiday = QComboBox()
        self.cmb_holiday.setEditable(True)
        self.cmb_holiday.addItems([
            "", "8 Марта", "14 Февраля", "День рождения", "Новый год",
            "Выпускной", "Подарок без повода", "Лето / отпуск"
        ])
        self.cmb_holiday.setCurrentText(self.settings.get("holiday", ""))
        grid.addWidget(self.cmb_holiday, r, 1, 1, 3)

        r += 1

        grid.addWidget(QLabel("SEO-плотность"), r, 0)
        self.cmb_seo = QComboBox()
        self.cmb_seo.addItems(["low", "normal", "high"])
        self.cmb_seo.setCurrentText(self.settings.get("seo", "normal"))
        grid.addWidget(self.cmb_seo, r, 1)

        grid.addWidget(QLabel("Стиль"), r, 2)
        self.cmb_style = QComboBox()
        self.cmb_style.addItems(["premium", "market", "social", "neutral"])
        self.cmb_style.setCurrentText(self.settings.get("style", "premium"))
        grid.addWidget(self.cmb_style, r, 3)

        r += 1

        grid.addWidget(QLabel("Пол"), r, 0)
        self.cmb_gender = QComboBox()
        self.cmb_gender.addItems(["auto", "female", "male", "unisex"])
        self.cmb_gender.setCurrentText(self.settings.get("gender", "auto"))
        grid.addWidget(self.cmb_gender, r, 1)

        grid.addWidget(QLabel("Бренд в названии"), r, 2)
        self.cmb_brand_title = QComboBox()
        self.cmb_brand_title.addItems(["0/100", "50/50", "100/0"])
        self.cmb_brand_title.setCurrentText(self.settings.get("brand_title", "50/50"))
        grid.addWidget(self.cmb_brand_title, r, 3)

        r += 1

        grid.addWidget(QLabel("Строк заполнять"), r, 0)
        self.spin_rows = QSpinBox()
        self.spin_rows.setRange(1, 50)
        self.spin_rows.setValue(int(self.settings.get("rows", 6)))
        grid.addWidget(self.spin_rows, r, 1)

        grid.addWidget(QLabel("Сколько Excel файлов"), r, 2)
        self.spin_files = QSpinBox()
        self.spin_files.setRange(1, 50)
        self.spin_files.setValue(int(self.settings.get("files", 1)))
        grid.addWidget(self.spin_files, r, 3)

        root.addWidget(card_params)

        # Bottom row
        bottom = QHBoxLayout()
        bottom.setSpacing(12)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        bottom.addWidget(self.progress, 1)

        self.btn_go = QPushButton("🚀  СГЕНЕРИРОВАТЬ")
        self.btn_go.setObjectName("Primary")
        self.btn_go.clicked.connect(self._generate)
        bottom.addWidget(self.btn_go)

        root.addLayout(bottom)

        # fix size
        for cb in (self.cmb_theme, self.cmb_brand, self.cmb_shape, self.cmb_lens, self.cmb_collection, self.cmb_holiday,
                   self.cmb_seo, self.cmb_style, self.cmb_gender, self.cmb_brand_title):
            cb.setMinimumHeight(36)

        # restore theme selection in UI
        self.cmb_theme.setCurrentText(self.settings.get("theme", "Graphite"))

    # --------------------------
    # Actions
    # --------------------------
    def _apply_theme(self, theme: str):
        theme = theme if theme in THEMES else "Graphite"
        self.setStyleSheet(THEMES[theme])
        self.cmb_theme.blockSignals(True)
        self.cmb_theme.setCurrentText(theme)
        self.cmb_theme.blockSignals(False)

    def _on_theme_changed(self, t: str):
        self._apply_theme(t)
        self._save_settings()

    def _open_data_folder(self):
        try:
            os.startfile(str(DATA_DIR))
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", str(e))

    def _choose_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выбрать XLSX", "", "Excel (*.xlsx)")
        if not path:
            return
        self.xlsx_path = path
        self.lb_xlsx.setText(path)
        self._save_settings()

    def _choose_out_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Папка вывода")
        if not d:
            return
        self.out_dir = d
        self.ed_out.setText(d)
        self._save_settings()

    def _add_brand(self):
        # просим латиницу и кириллицу
        lat, ok = QFileDialog.getSaveFileName  # just to avoid extra imports? no. We'll use QMessageBox+QLineEdit trick.
        from PyQt5.QtWidgets import QInputDialog

        brand_lat, ok1 = QInputDialog.getText(self, "Добавить бренд (латиница)", "Например: Miu Miu")
        if not ok1 or not brand_lat.strip():
            return
        brand_lat = brand_lat.strip()

        brand_ru, ok2 = QInputDialog.getText(self, "Кириллица для названия", "Например: Миу Миу")
        if not ok2 or not brand_ru.strip():
            return
        brand_ru = brand_ru.strip()

        save_list_append(BRANDS_FILE, brand_lat)
        self.brands = load_list(BRANDS_FILE, self.brands)

        key = _norm_key(brand_lat)
        self.brand_ru_map[key] = brand_ru
        save_brand_ru_map(self.brand_ru_map)

        # обновим комбобокс
        self._reload_combo(self.cmb_brand, self.brands, brand_lat)

        QMessageBox.information(self, "Готово", f"Добавлено:\n{brand_lat} → {brand_ru}")

    def _add_shape(self):
        from PyQt5.QtWidgets import QInputDialog
        val, ok = QInputDialog.getText(self, "Добавить форму оправы", "Например: Прямоугольные")
        if not ok or not val.strip():
            return
        val = val.strip()
        save_list_append(SHAPES_FILE, val)
        self.shapes = load_list(SHAPES_FILE, self.shapes)
        self._reload_combo(self.cmb_shape, self.shapes, val)

    def _add_lens(self):
        from PyQt5.QtWidgets import QInputDialog
        val, ok = QInputDialog.getText(self, "Добавить линзы", "Например: Хамелеон / Фотохромные")
        if not ok or not val.strip():
            return
        val = val.strip()
        save_list_append(LENSES_FILE, val)
        self.lenses = load_list(LENSES_FILE, self.lenses)
        self._reload_combo(self.cmb_lens, self.lenses, val)

    def _reload_combo(self, cb: QComboBox, items: List[str], current: str):
        cb.blockSignals(True)
        cb.clear()
        cb.addItems(items)
        cb.setEditable(True)
        cb.setCurrentText(current)
        cb.blockSignals(False)

    def _brand_ru_for(self, brand_lat: str) -> str:
        key = _norm_key(brand_lat)
        return self.brand_ru_map.get(key, "")  # если нет — пусто (бренд в названии может не вставляться)

    def _brand_ratio(self) -> float:
        v = self.cmb_brand_title.currentText()
        if v == "0/100":
            return 0.0
        if v == "100/0":
            return 1.0
        return 0.5

    def _generate(self):
        if not self.xlsx_path:
            QMessageBox.warning(self, "Ошибка", "Сначала выбери XLSX.")
            return
        if not self.out_dir:
            QMessageBox.warning(self, "Ошибка", "Выбери папку вывода.")
            return

        brand_lat = self.cmb_brand.currentText().strip()
        shape = self.cmb_shape.currentText().strip()
        lens = self.cmb_lens.currentText().strip()
        collection = self.cmb_collection.currentText().strip()
        holiday = self.cmb_holiday.currentText().strip()

        brand_ru = self._brand_ru_for(brand_lat)

        # если бренд в названии включён, но кириллицы нет — предупреждаем один раз
        if self._brand_ratio() > 0 and not brand_ru:
            QMessageBox.information(
                self,
                "Подсказка",
                "Для этого бренда нет кириллицы для названия.\n"
                "Нажми '+' рядом с брендом и добавь соответствие латиница→кириллица.\n"
                "В описании бренд останется латиницей."
            )

        self.progress.setValue(0)
        self.btn_go.setEnabled(False)

        self.worker = Worker(
            xlsx_path=self.xlsx_path,
            out_dir=self.out_dir,
            files_count=int(self.spin_files.value()),
            rows_to_fill=int(self.spin_rows.value()),
            brand_lat=brand_lat,
            brand_ru=brand_ru,
            shape=shape,
            lens=lens,
            collection=collection,
            holiday=holiday,
            seo_level=self.cmb_seo.currentText(),
            style=self.cmb_style.currentText(),
            gender=self.cmb_gender.currentText(),
            brand_ratio=self._brand_ratio(),
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.done.connect(self._done)
        self.worker.failed.connect(self._failed)
        self.worker.start()

        self._save_settings()

    def _done(self, out_dir: str):
        self.btn_go.setEnabled(True)
        QMessageBox.information(self, "Готово", f"Файлы сохранены в:\n{out_dir}")

    def _failed(self, err: str):
        self.btn_go.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", err)

    def _save_settings(self):
        s = {
            "theme": self.cmb_theme.currentText(),
            "brand": self.cmb_brand.currentText(),
            "shape": self.cmb_shape.currentText(),
            "lens": self.cmb_lens.currentText(),
            "collection": self.cmb_collection.currentText(),
            "holiday": self.cmb_holiday.currentText(),
            "seo": self.cmb_seo.currentText(),
            "style": self.cmb_style.currentText(),
            "gender": self.cmb_gender.currentText(),
            "brand_title": self.cmb_brand_title.currentText(),
            "rows": int(self.spin_rows.value()),
            "files": int(self.spin_files.value()),
            "out_dir": self.out_dir or "",
        }
        save_settings(s)


def main():
    app = QApplication([])
    w = App()
    w.show()
    app.exec_()


if __name__ == "__main__":
    main()
