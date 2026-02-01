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


def row_with_plus(cb: QComboBox, on_plus):
    row = QHBoxLayout()
    row.addWidget(cb, 1)
    btn = QPushButton("+")
    btn.setFixedWidth(38)
    btn.setToolTip("Добавить в список")
    btn.clicked.connect(on_plus)
    row.addWidget(btn)
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
        "ui_scale": "100%",
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
# Arrow SVG (always visible)
# ==========================
def ensure_arrow_svg(path: Path, color_hex: str):
    """
    Создаёт SVG-стрелку, чтобы она не пропадала на темах.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        return
    svg = f"""<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 12 12">
  <path d="M2 4.2 L6 8.2 L10 4.2" fill="none" stroke="{color_hex}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
</svg>"""
    path.write_text(svg, encoding="utf-8")


# ==========================
# Themes (Notion/Stripe feel) + arrow colors
# ==========================
THEME_META = {
    "Light":   {"bg": "#f7f7f8", "card": "#ffffff", "text": "#111", "muted": "#666", "border": "#ddd", "primary": "#111", "chunk": "#111", "arrow": "#111"},
    "Dark":    {"bg": "#1e1f22", "card": "#2a2b2f", "text": "#f2f2f2", "muted": "#aaa", "border": "#444", "primary": "#4a86ff", "chunk": "#4a86ff", "arrow": "#f2f2f2"},
    "Graphite":{"bg": "#2b2e34", "card": "#353a43", "text": "#f0f0f0", "muted": "#b9c0cc", "border": "#555d68", "primary": "#4a86ff", "chunk": "#4a86ff", "arrow": "#f0f0f0"},
    "Ocean":   {"bg": "#f4f8ff", "card": "#ffffff", "text": "#0b1220", "muted": "#4b5563", "border": "#cfe0ff", "primary": "#2563eb", "chunk": "#2563eb", "arrow": "#0b1220"},
    "Emerald": {"bg": "#f3fbf7", "card": "#ffffff", "text": "#072012", "muted": "#3a6b52", "border": "#bfe8d2", "primary": "#059669", "chunk": "#059669", "arrow": "#072012"},
    "Sepia":   {"bg": "#fbf6ef", "card": "#ffffff", "text": "#2a1f14", "muted": "#6b4f33", "border": "#ead7c2", "primary": "#7c3aed", "chunk": "#7c3aed", "arrow": "#2a1f14"},
}

SCALE_MAP = {
    "100%": 13,
    "115%": 15,
    "130%": 17,
}


def build_stylesheet(meta: dict, font_px: int, arrow_uri: str) -> str:
    # arrow_uri должен быть file:///...
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

        QLabel#title {{ font-size: {font_px + 9}px; font-weight: 650; }}
        QLabel#subtitle {{ color: {meta["muted"]}; }}

        QComboBox {{
            padding: 8px 34px 8px 10px;
            border-radius: 10px;
            border: 1px solid {meta["border"]};
            background: {meta["card"]};
        }}
        QComboBox:focus {{
            border: 1px solid {meta["primary"]};
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 28px;
            border-left: 1px solid {meta["border"]};
            border-top-right-radius: 10px;
            border-bottom-right-radius: 10px;
            background: {meta["card"]};
        }}
        QComboBox::down-arrow {{
            image: url("{arrow_uri}");
            width: 12px;
            height: 12px;
        }}

        QPushButton {{
            padding: 10px;
            border-radius: 12px;
            border: 1px solid {meta["border"]};
            background: {meta["card"]};
        }}
        QPushButton:hover {{
            background: rgba(0,0,0,0.05);
        }}

        QPushButton#primary {{
            background: {meta["primary"]};
            color: white;
            border: none;
            font-weight: 650;
            padding: 12px;
        }}
        QPushButton#primary:hover {{
            opacity: 0.95;
        }}

        QProgressBar {{
            border: 1px solid {meta["border"]};
            border-radius: 10px;
            height: 18px;
            text-align: center;
            background: {meta["card"]};
        }}
        QProgressBar::chunk {{
            background: {meta["chunk"]};
            border-radius: 10px;
        }}
    """


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
        self.resize(1020, 780)

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
        self.cb_scale.setCurrentText(self.settings.get("ui_scale", "100%"))
        self.cb_scale.currentTextChanged.connect(self.on_scale_changed)
        ts_row.addWidget(self.cb_scale)

        root.addLayout(ts_row)

        # ---- Data folder row
        data_row = QHBoxLayout()
        self.lbl_data = QLabel(f"📁 Справочники: {self.data_dir}")
        self.lbl_data.setWordWrap(True)
        btn_open = QPushButton("📂 Папка")
        btn_open.clicked.connect(self.open_data_folder)
        data_row.addWidget(self.lbl_data, 1)
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

        # ---- Combos
        root.addWidget(QLabel("Бренд"))
        self.cb_brand = make_combo(load_list(self.brands_file), "Выбери бренд или впиши свой")
        self.cb_brand.setCurrentText(self.settings.get("brand", ""))
        root.addLayout(row_with_plus(self.cb_brand, self.add_brand))

        root.addWidget(QLabel("Форма оправы"))
        self.cb_shape = make_combo(load_list(self.shapes_file), "Выбери форму или впиши свою")
        self.cb_shape.setCurrentText(self.settings.get("shape", ""))
        root.addLayout(row_with_plus(self.cb_shape, self.add_shape))

        root.addWidget(QLabel("Линзы / особенности"))
        self.cb_lens = make_combo(load_list(self.lenses_file), "Выбери линзы или впиши свои")
        self.cb_lens.setCurrentText(self.settings.get("lens", ""))
        root.addLayout(row_with_plus(self.cb_lens, self.add_lens))

        root.addWidget(QLabel("Коллекция"))
        self.cb_collection = make_combo(
            ["Весна–Лето 2025–2026", "Весна–Лето 2026", "Осень–Зима 2025–2026", "Осень–Зима 2026"],
            "Выбери коллекцию"
        )
        self.cb_collection.setCurrentText(self.settings.get("collection", "Весна–Лето 2025–2026"))
        root.addWidget(self.cb_collection)

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

        # Apply theme/scale at the end
        self.apply_theme_and_scale()

    # ---------- folder ----------
    def open_data_folder(self):
        try:
            subprocess.Popen(f'explorer "{self.data_dir}"')
        except Exception:
            QMessageBox.warning(self, "Ошибка", f"Не удалось открыть папку:\n{self.data_dir}")

    # ---------- theme/scale ----------
    def apply_theme_and_scale(self):
        theme = self.cb_theme.currentText()
        scale = self.cb_scale.currentText()

        meta = THEME_META.get(theme, THEME_META["Light"])
        font_px = SCALE_MAP.get(scale, 13)

        # Создаём SVG-стрелку под тему (в data_dir), чтобы она НЕ пропадала
        arrow_file = self.data_dir / f"arrow_{theme}.svg"
        ensure_arrow_svg(arrow_file, meta["arrow"])
        arrow_uri = arrow_file.resolve().as_uri()

        self.setStyleSheet(build_stylesheet(meta, font_px, arrow_uri))

    def on_theme_changed(self, _):
        self.settings["theme"] = self.cb_theme.currentText()
        save_settings(self.settings_file, self.settings)
        self.apply_theme_and_scale()

    def on_scale_changed(self, _):
        self.settings["ui_scale"] = self.cb_scale.currentText()
        save_settings(self.settings_file, self.settings)
        self.apply_theme_and_scale()

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

        # сохраняем выборы
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
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
