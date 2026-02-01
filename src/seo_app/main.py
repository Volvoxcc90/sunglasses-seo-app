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
# Themes (Notion / Stripe)
# ==========================
THEMES = {
    "Light": """
        QWidget { background: #f7f7f8; color: #111; font-size: 13px; }
        QFrame#card { background: #ffffff; border-radius: 14px; padding: 18px; }
        QLabel#title { font-size: 22px; font-weight: 600; }
        QLabel#subtitle { color: #666; }
        QComboBox { padding: 8px; border-radius: 10px; border: 1px solid #ddd; background: #fff; }
        QComboBox::drop-down { width: 28px; border-left: 1px solid #ddd; }
        QPushButton { padding: 10px; border-radius: 12px; border: 1px solid #ddd; background: #fff; }
        QPushButton:hover { background: #f1f1f2; }
        QPushButton#primary { background: #111; color: white; border: none; font-weight: 600; }
        QPushButton#primary:hover { background: #333; }
        QProgressBar { border: 1px solid #ddd; border-radius: 10px; height: 18px; text-align: center; }
        QProgressBar::chunk { background: #111; border-radius: 10px; }
    """,
    "Dark": """
        QWidget { background: #1e1f22; color: #f2f2f2; font-size: 13px; }
        QFrame#card { background: #2a2b2f; border-radius: 14px; padding: 18px; }
        QLabel#title { font-size: 22px; font-weight: 600; }
        QLabel#subtitle { color: #aaa; }
        QComboBox { padding: 8px; border-radius: 10px; border: 1px solid #444; background: #1e1f22; }
        QComboBox::drop-down { width: 28px; border-left: 1px solid #444; }
        QPushButton { padding: 10px; border-radius: 12px; border: 1px solid #444; background: #2a2b2f; color: #f2f2f2; }
        QPushButton:hover { background: #34363c; }
        QPushButton#primary { background: #4a86ff; color: white; border: none; font-weight: 600; }
        QPushButton#primary:hover { background: #3b78ff; }
        QProgressBar { border: 1px solid #444; border-radius: 10px; height: 18px; text-align: center; }
        QProgressBar::chunk { background: #4a86ff; border-radius: 10px; }
    """
}


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
    btn.clicked.connect(on_plus)
    row.addWidget(btn)
    return row


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

        self.settings = self.load_settings()
        self.input_file = ""

        self.setWindowTitle("Sunglasses SEO PRO")
        self.resize(980, 740)

        root = QVBoxLayout(self)
        root.setSpacing(14)

        # ---- Header card
        card = QFrame()
        card.setObjectName("card")
        cl = QVBoxLayout(card)
        title = QLabel("🕶️ Sunglasses SEO PRO")
        title.setObjectName("title")
        subtitle = QLabel("Живые SEO-описания • Выпадающие списки • Прогресс")
        subtitle.setObjectName("subtitle")
        cl.addWidget(title)
        cl.addWidget(subtitle)
        root.addWidget(card)

        # ---- Theme
        theme_row = QHBoxLayout()
        theme_row.addWidget(QLabel("🎨 Тема"))
        self.cb_theme = QComboBox()
        self.cb_theme.addItems(["Light", "Dark"])
        self.cb_theme.setCurrentText(self.settings.get("theme", "Light"))
        self.cb_theme.currentTextChanged.connect(self.on_theme)
        theme_row.addWidget(self.cb_theme, 1)
        root.addLayout(theme_row)

        # ---- Data folder row
        data_row = QHBoxLayout()
        self.lbl_data = QLabel(f"📁 Справочники: {self.data_dir}")
        self.lbl_data.setWordWrap(True)
        btn_open = QPushButton("📂 Папка")
        btn_open.clicked.connect(self.open_data_folder)
        data_row.addWidget(self.lbl_data, 1)
        data_row.addWidget(btn_open)
        root.addLayout(data_row)

        # ---- File
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
        root.addWidget(self.progress)

        self.btn_run = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        self.btn_run.setObjectName("primary")
        self.btn_run.clicked.connect(self.run)
        root.addWidget(self.btn_run)

        self.apply_theme(self.cb_theme.currentText())

    # ---------- helpers ----------
    def load_settings(self):
        if self.settings_file.exists():
            try:
                return json.loads(self.settings_file.read_text(encoding="utf-8"))
            except Exception:
                pass
        return {}

    def save_settings(self):
        self.settings_file.write_text(json.dumps(self.settings, ensure_ascii=False, indent=2), encoding="utf-8")

    def open_data_folder(self):
        try:
            subprocess.Popen(f'explorer "{self.data_dir}"')
        except Exception:
            QMessageBox.warning(self, "Ошибка", f"Не удалось открыть папку:\n{self.data_dir}")

    def apply_theme(self, theme):
        self.setStyleSheet(THEMES.get(theme, THEMES["Light"]))

    def on_theme(self, theme):
        self.settings["theme"] = theme
        self.save_settings()
        self.apply_theme(theme)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX", "", "Excel (*.xlsx)")
        if path:
            self.input_file = path
            self.lbl_file.setText(path)

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
            "style": style
        })
        self.save_settings()

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
