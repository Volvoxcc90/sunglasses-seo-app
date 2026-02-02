import sys, os, json, subprocess
from pathlib import Path

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox,
    QComboBox, QRadioButton, QGroupBox, QProgressBar, QFrame
)

from seo_app.wb_fill import fill_wb_template


# ==========================
# Defaults
# ==========================
DEFAULT_BRANDS = [
    "Cazal","Ray-Ban","Gucci","Prada","Dior","Versace","Dolce & Gabbana",
    "Tom Ford","Chanel","Cartier","Oakley","Polaroid","Carrera","Fendi",
    "Givenchy","Balenciaga","Miu Miu","Burberry","Armani","Hugo Boss"
]
DEFAULT_SHAPES = [
    "Квадратная","Овальная","Круглая","Прямоугольная",
    "Авиаторы","Cat Eye","Оверсайз","Панто","Wayfarer"
]
DEFAULT_LENSES = [
    "UV400","поляризационные","фотохромные","хамелеон",
    "градиентные","зеркальные","антибликовые","с откидными линзами"
]


# ==========================
# Themes + UI scale
# ==========================
THEME_META = {
    "Light":   {"bg": "#f7f7f8", "card": "#ffffff", "text": "#111", "muted": "#666", "border": "#ddd", "primary": "#111", "chunk": "#111"},
    "Dark":    {"bg": "#1e1f22", "card": "#2a2b2f", "text": "#f2f2f2", "muted": "#aaa", "border": "#444", "primary": "#4a86ff", "chunk": "#4a86ff"},
    "Graphite":{"bg": "#2b2e34", "card": "#353a43", "text": "#f0f0f0", "muted": "#b9c0cc", "border": "#555d68", "primary": "#4a86ff", "chunk": "#4a86ff"},
    "Ocean":   {"bg": "#f4f8ff", "card": "#ffffff", "text": "#0b1220", "muted": "#4b5563", "border": "#cfe0ff", "primary": "#2563eb", "chunk": "#2563eb"},
    "Emerald": {"bg": "#f3fbf7", "card": "#ffffff", "text": "#072012", "muted": "#3a6b52", "border": "#bfe8d2", "primary": "#059669", "chunk": "#059669"},
    "Sepia":   {"bg": "#fbf6ef", "card": "#ffffff", "text": "#2a1f14", "muted": "#6b4f33", "border": "#ead7c2", "primary": "#7c3aed", "chunk": "#7c3aed"},
}
SCALE_MAP = {"100%": 13, "115%": 15, "130%": 17, "145%": 19}

def build_stylesheet(meta: dict, font_px: int) -> str:
    # Важно: скрываем нативную стрелку QComboBox (чтобы не было “пустой зоны”),
    # и рисуем стрелку отдельной кнопкой ▼ рядом.
    return f"""
        QWidget {{
            background: {meta["bg"]};
            color: {meta["text"]};
            font-size: {font_px}px;
        }}

        QFrame#card {{
            background: {meta["card"]};
            border-radius: 14px;
            padding: 18px;
            border: 1px solid {meta["border"]};
        }}

        QLabel#title {{ font-size: {font_px + 10}px; font-weight: 650; }}
        QLabel#subtitle {{ color: {meta["muted"]}; }}

        QAbstractItemView {{
            background: {meta["card"]};
            color: {meta["text"]};
            border: 1px solid {meta["border"]};
            selection-background-color: {meta["primary"]};
            selection-color: white;
        }}

        QComboBox {{
            padding: 10px 12px 10px 12px;
            border-radius: 12px;
            border: 1px solid {meta["border"]};
            background: {meta["card"]};
        }}
        QComboBox:hover {{ border: 1px solid {meta["primary"]}; }}
        QComboBox:focus {{ border: 2px solid {meta["primary"]}; }}

        /* Скрываем системную стрелку */
        QComboBox::drop-down {{ width: 0px; border: none; }}
        QComboBox::down-arrow {{ image: none; }}

        QPushButton {{
            padding: 10px;
            border-radius: 12px;
            border: 1px solid {meta["border"]};
            background: {meta["card"]};
        }}
        QPushButton:hover {{ background: rgba(0,0,0,0.06); }}

        QPushButton#primary {{
            background: {meta["primary"]};
            color: white;
            border: none;
            font-weight: 650;
            padding: 12px;
            border-radius: 14px;
        }}

        /* Кнопка-стрелка ▼ */
        QPushButton#drop {{
            font-weight: 700;
            padding: 10px;
            min-width: 44px;
            max-width: 44px;
            border-radius: 12px;
        }}

        /* Кнопка + */
        QPushButton#plus {{
            font-weight: 800;
            padding: 10px;
            min-width: 44px;
            max-width: 44px;
            border-radius: 12px;
        }}

        QProgressBar {{
            border: 1px solid {meta["border"]};
            border-radius: 12px;
            height: 20px;
            text-align: center;
            background: {meta["card"]};
        }}
        QProgressBar::chunk {{
            background: {meta["chunk"]};
            border-radius: 12px;
        }}
    """


# ==========================
# Data dir logic
# ==========================
def get_data_dir() -> Path:
    base = Path(sys.argv[0]).resolve().parent
    local = base / "data"
    if local.exists():
        return local
    appdata = os.environ.get("APPDATA") or str(Path.home())
    return Path(appdata) / "Sunglasses SEO PRO" / "data"


def ensure_txt(path: Path, defaults: list[str]):
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists() or not path.read_text(encoding="utf-8", errors="ignore").strip():
        path.write_text("\n".join(defaults), encoding="utf-8")


def load_list(path: Path) -> list[str]:
    if not path.exists():
        return []
    return sorted(
        {x.strip() for x in path.read_text(encoding="utf-8", errors="ignore").splitlines() if x.strip()},
        key=str.lower
    )


def save_item(path: Path, value: str) -> bool:
    value = value.strip()
    if not value:
        return False
    items = load_list(path)
    if value in items:
        return False
    items.append(value)
    path.write_text("\n".join(sorted(set(items), key=str.lower)), encoding="utf-8")
    return True


def refresh_combo(cb: QComboBox, path: Path, keep: str):
    cb.blockSignals(True)
    cb.clear()
    cb.addItems(load_list(path))
    cb.setCurrentText(keep)
    cb.blockSignals(False)


def make_combo(items: list[str], placeholder: str) -> QComboBox:
    cb = QComboBox()
    cb.setEditable(True)
    cb.setMaxVisibleItems(20)
    cb.addItems(items)
    cb.setPlaceholderText(placeholder)

    comp = cb.completer()
    comp.setCaseSensitivity(Qt.CaseInsensitive)
    comp.setFilterMode(Qt.MatchContains)

    return cb


def row_combo_drop_plus(cb: QComboBox, on_drop, on_plus) -> QHBoxLayout:
    """
    Ряд: [ComboBox][▼][+]
    ▼ всегда видима (это кнопка), открывает popup списка.
    """
    row = QHBoxLayout()
    row.addWidget(cb, 1)

    btn_drop = QPushButton("▼")
    btn_drop.setObjectName("drop")
    btn_drop.setToolTip("Открыть список")
    btn_drop.clicked.connect(on_drop)
    row.addWidget(btn_drop)

    btn_plus = QPushButton("+")
    btn_plus.setObjectName("plus")
    btn_plus.setToolTip("Добавить в список")
    btn_plus.clicked.connect(on_plus)
    row.addWidget(btn_plus)

    return row


# ==========================
# Settings persistence
# ==========================
def load_settings(settings_file: Path) -> dict:
    if settings_file.exists():
        try:
            return json.loads(settings_file.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {
        "theme": "Light",
        "ui_scale": "115%",
        "brand": "",
        "shape": "",
        "lens": "",
        "collection": "Весна–Лето 2025–2026",
        "style": "neutral"
    }


def save_settings(settings_file: Path, data: dict):
    settings_file.parent.mkdir(parents=True, exist_ok=True)
    settings_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ==========================
# Main Window
# ==========================
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.data_dir = get_data_dir()
        self.data_dir.mkdir(parents=True, exist_ok=True)

        self.brands_file = self.data_dir / "brands.txt"
        self.shapes_file = self.data_dir / "shapes.txt"
        self.lenses_file = self.data_dir / "lenses.txt"
        self.settings_file = self.data_dir / "ui_settings.json"

        ensure_txt(self.brands_file, DEFAULT_BRANDS)
        ensure_txt(self.shapes_file, DEFAULT_SHAPES)
        ensure_txt(self.lenses_file, DEFAULT_LENSES)

        self.settings = load_settings(self.settings_file)
        self.input_file = ""

        self.setWindowTitle("Sunglasses SEO PRO")
        self.resize(1040, 820)

        root = QVBoxLayout(self)
        root.setSpacing(14)

        # ---- Header card
        card = QFrame()
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        title = QLabel("🕶️ Sunglasses SEO PRO")
        title.setObjectName("title")
        subtitle = QLabel("Живые SEO-описания • Выпадающие списки • Прогресс • Темы")
        subtitle.setObjectName("subtitle")
        cl.addWidget(title)
        cl.addWidget(subtitle)
        root.addWidget(card)

        # ---- Theme + Scale row
        ts_row = QHBoxLayout()
        ts_row.addWidget(QLabel("🎨 Тема"))
        self.cb_theme = QComboBox()
        self.cb_theme.addItems(list(THEME_META.keys()))
        self.cb_theme.setCurrentText(self.settings.get("theme", "Light"))
        self.cb_theme.currentTextChanged.connect(self.on_theme_changed)
        ts_row.addWidget(self.cb_theme, 1)

        ts_row.addWidget(QLabel("🔎 Размер UI"))
        self.cb_scale = QComboBox()
        self.cb_scale.addItems(list(SCALE_MAP.keys()))
        self.cb_scale.setCurrentText(self.settings.get("ui_scale", "115%"))
        self.cb_scale.currentTextChanged.connect(self.on_scale_changed)
        ts_row.addWidget(self.cb_scale)

        root.addLayout(ts_row)

        # ---- Data folder row
        data_row = QHBoxLayout()
        lbl = QLabel(f"📁 Справочники: {self.data_dir}")
        lbl.setWordWrap(True)
        btn_open = QPushButton("📂 Папка")
        btn_open.clicked.connect(self.open_data_folder)
        data_row.addWidget(lbl, 1)
        data_row.addWidget(btn_open)
        root.addLayout(data_row)

        # ---- File row
        file_row = QHBoxLayout()
        btn_file = QPushButton("📄 Загрузить XLSX")
        btn_file.clicked.connect(self.pick_file)
        self.lbl_file = QLabel("Файл не выбран")
        self.lbl_file.setWordWrap(True)
        file_row.addWidget(btn_file)
        file_row.addWidget(self.lbl_file, 1)
        root.addLayout(file_row)

        # ---- Combos with real ▼ buttons
        root.addWidget(QLabel("Бренд"))
        self.cb_brand = make_combo(load_list(self.brands_file), "Выбери бренд или впиши свой")
        self.cb_brand.setCurrentText(self.settings.get("brand", ""))
        root.addLayout(row_combo_drop_plus(
            self.cb_brand,
            on_drop=self.cb_brand.showPopup,
            on_plus=self.add_brand
        ))

        root.addWidget(QLabel("Форма оправы"))
        self.cb_shape = make_combo(load_list(self.shapes_file), "Выбери форму или впиши свою")
        self.cb_shape.setCurrentText(self.settings.get("shape", ""))
        root.addLayout(row_combo_drop_plus(
            self.cb_shape,
            on_drop=self.cb_shape.showPopup,
            on_plus=self.add_shape
        ))

        root.addWidget(QLabel("Линзы / особенности"))
        self.cb_lens = make_combo(load_list(self.lenses_file), "Выбери линзы или впиши свои")
        self.cb_lens.setCurrentText(self.settings.get("lens", ""))
        root.addLayout(row_combo_drop_plus(
            self.cb_lens,
            on_drop=self.cb_lens.showPopup,
            on_plus=self.add_lens
        ))

        root.addWidget(QLabel("Коллекция"))
        self.cb_collection = make_combo(
            ["Весна–Лето 2025–2026", "Весна–Лето 2026", "Осень–Зима 2025–2026", "Осень–Зима 2026"],
            "Выбери коллекцию"
        )
        self.cb_collection.setCurrentText(self.settings.get("collection", "Весна–Лето 2025–2026"))
        root.addLayout(row_combo_drop_plus(
            self.cb_collection,
            on_drop=self.cb_collection.showPopup,
            on_plus=lambda: None  # коллекцию добавлять не нужно, но кнопку оставим “пустой”
        ))

        # Сделаем + у коллекции неактивным (чтобы не путало)
        # (кнопка плюс там есть, но мы её отключим по факту ниже)
        # Для этого найдём последнюю кнопку "plus" и отключим:
        # проще — отключить через явный доступ:
        # (мы не храним ссылку, но можно просто не добавлять "+". Оставим как есть и отключим ниже.)
        # ---- Style
        style_box = QGroupBox("Стиль описания")
        sb = QHBoxLayout(style_box)
        self.rb_neutral = QRadioButton("Neutral")
        self.rb_premium = QRadioButton("Premium")
        self.rb_social = QRadioButton("Social")
        sb.addWidget(self.rb_neutral)
        sb.addWidget(self.rb_premium)
        sb.addWidget(self.rb_social)
        root.addWidget(style_box)

        style = self.settings.get("style", "neutral")
        {"neutral": self.rb_neutral, "premium": self.rb_premium, "social": self.rb_social}.get(style, self.rb_neutral).setChecked(True)

        # ---- Progress + Run
        self.progress = QProgressBar()
        self.progress.setValue(0)
        root.addWidget(self.progress)

        self.btn_run = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        self.btn_run.setObjectName("primary")
        self.btn_run.clicked.connect(self.run)
        root.addWidget(self.btn_run)

        self.apply_theme_and_scale()

        # Отключаем “+” у коллекции (последняя кнопка plus в интерфейсе)
        # Надёжно: берём все кнопки plus и отключаем последнюю.
        plus_buttons = self.findChildren(QPushButton, "plus")
        if plus_buttons:
            plus_buttons[-1].setEnabled(False)
            plus_buttons[-1].setToolTip("Для коллекции добавление не нужно")

    # ---------- theme/scale ----------
    def apply_theme_and_scale(self):
        theme = self.cb_theme.currentText()
        scale = self.cb_scale.currentText()

        meta = THEME_META.get(theme, THEME_META["Light"])
        font_px = SCALE_MAP.get(scale, 15)

        self.setStyleSheet(build_stylesheet(meta, font_px))

    def on_theme_changed(self, _):
        self.settings["theme"] = self.cb_theme.currentText()
        save_settings(self.settings_file, self.settings)
        self.apply_theme_and_scale()

    def on_scale_changed(self, _):
        self.settings["ui_scale"] = self.cb_scale.currentText()
        save_settings(self.settings_file, self.settings)
        self.apply_theme_and_scale()

    # ---------- folder ----------
    def open_data_folder(self):
        try:
            subprocess.Popen(f'explorer "{self.data_dir}"')
        except Exception:
            QMessageBox.warning(self, "Ошибка", f"Не удалось открыть папку:\n{self.data_dir}")

    # ---------- file ----------
    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX", "", "Excel (*.xlsx)")
        if path:
            self.input_file = path
            self.lbl_file.setText(path)

    # ---------- add items ----------
    def add_brand(self):
        v = self.cb_brand.currentText().strip()
        if save_item(self.brands_file, v):
            refresh_combo(self.cb_brand, self.brands_file, v)

    def add_shape(self):
        v = self.cb_shape.currentText().strip()
        if save_item(self.shapes_file, v):
            refresh_combo(self.cb_shape, self.shapes_file, v)

    def add_lens(self):
        v = self.cb_lens.currentText().strip()
        if save_item(self.lenses_file, v):
            refresh_combo(self.cb_lens, self.lenses_file, v)

    # ---------- run ----------
    def run(self):
        if not self.input_file:
            QMessageBox.warning(self, "Ошибка", "Выбери XLSX файл")
            return

        style = "premium" if self.rb_premium.isChecked() else "social" if self.rb_social.isChecked() else "neutral"

        self.settings.update({
            "brand": self.cb_brand.currentText(),
            "shape": self.cb_shape.currentText(),
            "lens": self.cb_lens.currentText(),
            "collection": self.cb_collection.currentText(),
            "style": style,
        })
        save_settings(self.settings_file, self.settings)

        self.progress.setValue(0)
        self.btn_run.setEnabled(False)

        try:
            out, rows = fill_wb_template(
                input_xlsx=self.input_file,
                brand=self.cb_brand.currentText(),
                shape=self.cb_shape.currentText(),
                lens_features=self.cb_lens.currentText(),
                collection=self.cb_collection.currentText(),
                style=style,
                progress_callback=lambda p: self.progress.setValue(int(p))
            )
            self.progress.setValue(100)
            QMessageBox.information(self, "Готово", f"Создан файл:\n{out}\nСтрок заполнено: {rows}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))
        finally:
            self.btn_run.setEnabled(True)


def main():
    app = QApplication(sys.argv)
    # Fusion делает QSS стабильнее (но даже без него кнопки-стрелки работают)
    app.setStyle("Fusion")

    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
