import sys
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox,
    QComboBox, QRadioButton, QGroupBox
)

from seo_app.wb_fill import fill_wb_template


# ======================
# Theme styles (QSS)
# ======================

THEMES = {
    "Светлая": """
        QWidget { background: #f5f6f8; color: #1e1e1e; font-size: 13px; }
        QComboBox, QLabel { padding: 6px; }
        QPushButton {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                        stop:0 #ffffff, stop:1 #dcdcdc);
            border: 1px solid #b5b5b5;
            border-radius: 8px;
            padding: 10px;
        }
        QPushButton:hover { background: #ffffff; }
        QPushButton:pressed {
            background: #cfcfcf;
            padding-top: 12px;
        }
    """,

    "Тёмная": """
        QWidget { background: #1f1f1f; color: #e6e6e6; font-size: 13px; }
        QComboBox, QLabel { padding: 6px; }
        QPushButton {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                        stop:0 #3a3a3a, stop:1 #2a2a2a);
            border: 1px solid #555;
            border-radius: 8px;
            padding: 10px;
        }
        QPushButton:hover { background: #444; }
        QPushButton:pressed {
            background: #1e1e1e;
            padding-top: 12px;
        }
    """,

    "Graphite": """
        QWidget { background: #2b2e34; color: #f0f0f0; font-size: 13px; }
        QComboBox, QLabel { padding: 6px; }
        QPushButton {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                        stop:0 #4a4f57, stop:1 #2f333a);
            border: 1px solid #666;
            border-radius: 10px;
            padding: 10px;
        }
        QPushButton:hover { background: #555b65; }
        QPushButton:pressed {
            background: #23262b;
            padding-top: 12px;
        }
    """
}


# ======================
# Helpers
# ======================

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
BRANDS_FILE = DATA_DIR / "brands.txt"
SHAPES_FILE = DATA_DIR / "shapes.txt"
LENSES_FILE = DATA_DIR / "lenses.txt"


def ensure_data():
    DATA_DIR.mkdir(exist_ok=True)
    for f in [BRANDS_FILE, SHAPES_FILE, LENSES_FILE]:
        if not f.exists():
            f.write_text("", encoding="utf-8")


def load_list(path: Path):
    if not path.exists():
        return []
    return sorted(
        {x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()},
        key=str.lower
    )


def save_item(path: Path, value: str):
    value = value.strip()
    if not value:
        return False
    items = load_list(path)
    if value in items:
        return False
    items.append(value)
    path.write_text("\n".join(sorted(items, key=str.lower)), encoding="utf-8")
    return True


def combo(items, placeholder):
    cb = QComboBox()
    cb.setEditable(True)
    cb.addItems(items)
    cb.setPlaceholderText(placeholder)
    cb.completer().setCaseSensitivity(Qt.CaseInsensitive)
    cb.completer().setFilterMode(Qt.MatchContains)
    return cb


def row_with_plus(cb, handler):
    row = QHBoxLayout()
    row.addWidget(cb, 1)
    btn = QPushButton("+")
    btn.setFixedWidth(36)
    btn.clicked.connect(handler)
    row.addWidget(btn)
    return row


# ======================
# Main Window
# ======================

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        ensure_data()

        self.setWindowTitle("Sunglasses SEO PRO")
        self.resize(920, 620)

        self.input_file = ""

        root = QVBoxLayout(self)
        root.setSpacing(12)

        # ---- File
        file_row = QHBoxLayout()
        btn_file = QPushButton("📂 Загрузить XLSX")
        btn_file.clicked.connect(self.pick_file)
        self.lbl_file = QLabel("Файл не выбран")
        self.lbl_file.setWordWrap(True)
        file_row.addWidget(btn_file)
        file_row.addWidget(self.lbl_file, 1)
        root.addLayout(file_row)

        # ---- Combos
        self.cb_brand = combo(load_list(BRANDS_FILE), "Бренд")
        root.addLayout(row_with_plus(self.cb_brand, lambda: save_item(BRANDS_FILE, self.cb_brand.currentText())))

        self.cb_shape = combo(load_list(SHAPES_FILE), "Форма оправы")
        root.addLayout(row_with_plus(self.cb_shape, lambda: save_item(SHAPES_FILE, self.cb_shape.currentText())))

        self.cb_lens = combo(load_list(LENSES_FILE), "Линзы")
        root.addLayout(row_with_plus(self.cb_lens, lambda: save_item(LENSES_FILE, self.cb_lens.currentText())))

        self.cb_collection = combo(
            ["Весна–Лето 2025–2026", "Весна–Лето 2026", "Осень–Зима 2025–2026"],
            "Коллекция"
        )
        root.addWidget(self.cb_collection)

        # ---- Style
        style_box = QGroupBox("Стиль описания")
        sb = QHBoxLayout(style_box)
        self.rb_neutral = QRadioButton("Нейтральный")
        self.rb_premium = QRadioButton("Премиум")
        self.rb_mass = QRadioButton("Масс")
        self.rb_social = QRadioButton("Соцсети")
        self.rb_neutral.setChecked(True)
        for rb in [self.rb_neutral, self.rb_premium, self.rb_mass, self.rb_social]:
            sb.addWidget(rb)
        root.addWidget(style_box)

        # ---- Theme
        self.cb_theme = QComboBox()
        self.cb_theme.addItems(THEMES.keys())
        self.cb_theme.currentTextChanged.connect(self.apply_theme)
        root.addWidget(QLabel("🎨 Тема интерфейса"))
        root.addWidget(self.cb_theme)

        # ---- Run
        self.btn_run = QPushButton("ГОТОВО")
        self.btn_run.clicked.connect(self.run)
        root.addWidget(self.btn_run)

        self.lbl_status = QLabel("")
        self.lbl_status.setWordWrap(True)
        root.addWidget(self.lbl_status)

        self.apply_theme("Светлая")

    def apply_theme(self, name):
        self.setStyleSheet(THEMES.get(name, ""))

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX", "", "Excel (*.xlsx)")
        if path:
            self.input_file = path
            self.lbl_file.setText(path)

    def run(self):
        if not self.input_file:
            QMessageBox.warning(self, "Ошибка", "Выберите XLSX файл")
            return

        style = (
            "premium" if self.rb_premium.isChecked()
            else "mass" if self.rb_mass.isChecked()
            else "social" if self.rb_social.isChecked()
            else "neutral"
        )

        out, rows = fill_wb_template(
            input_xlsx=self.input_file,
            brand=self.cb_brand.currentText(),
            shape=self.cb_shape.currentText(),
            lens_features=self.cb_lens.currentText(),
            collection=self.cb_collection.currentText(),
            style=style
        )

        self.lbl_status.setText(f"Готово. Заполнено строк: {rows}\n{out}")
        QMessageBox.information(self, "Готово", "Файл создан")


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
