import sys
import json
from pathlib import Path

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox,
    QComboBox, QRadioButton, QGroupBox, QProgressBar, QFrame
)

from seo_app.wb_fill import fill_wb_template


# ======================
# Paths / settings
# ======================

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

SETTINGS_FILE = DATA_DIR / "ui_settings.json"

BRANDS_FILE = DATA_DIR / "brands.txt"
SHAPES_FILE = DATA_DIR / "shapes.txt"
LENSES_FILE = DATA_DIR / "lenses.txt"


# ======================
# UI SETTINGS
# ======================

def load_ui_settings():
    if SETTINGS_FILE.exists():
        try:
            return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"theme": "Светлая"}


def save_ui_settings(settings: dict):
    SETTINGS_FILE.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )


# ======================
# Themes (QSS)
# ======================

THEMES = {
    "Светлая": """
        QWidget { background: #f5f6f8; color: #1e1e1e; font-size: 13px; }
        QFrame#card { background: white; border-radius: 14px; padding: 16px; }
        QLabel#title { font-size: 20px; font-weight: 600; }
        QLabel#subtitle { color: #666; }
        QComboBox, QLabel { padding: 6px; }

        QPushButton {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                        stop:0 #ffffff, stop:1 #dcdcdc);
            border: 1px solid #b5b5b5;
            border-radius: 10px;
            padding: 10px;
        }
        QPushButton:hover { background: #ffffff; }
        QPushButton:pressed { background: #cfcfcf; padding-top: 12px; }

        QPushButton#primary {
            background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                        stop:0 #4a86ff, stop:1 #2f6fff);
            color: white;
            font-weight: 600;
            border: none;
        }
    """,

    "Тёмная": """
        QWidget { background: #1f1f1f; color: #e6e6e6; font-size: 13px; }
        QFrame#card { background: #2a2a2a; border-radius: 14px; padding: 16px; }
        QLabel#title { font-size: 20px; font-weight: 600; }
        QLabel#subtitle { color: #aaa; }
        QComboBox, QLabel { padding: 6px; }

        QPushButton {
            background: #333;
            border: 1px solid #555;
            border-radius: 10px;
            padding: 10px;
        }
        QPushButton:hover { background: #444; }
        QPushButton:pressed { background: #222; padding-top: 12px; }

        QPushButton#primary {
            background: #4a86ff;
            color: white;
            font-weight: 600;
            border: none;
        }
    """,

    "Graphite": """
        QWidget { background: #2b2e34; color: #f0f0f0; font-size: 13px; }
        QFrame#card { background: #353a43; border-radius: 14px; padding: 16px; }
        QLabel#title { font-size: 20px; font-weight: 600; }
        QLabel#subtitle { color: #bbb; }
        QComboBox, QLabel { padding: 6px; }

        QPushButton {
            background: #4a4f57;
            border: 1px solid #666;
            border-radius: 10px;
            padding: 10px;
        }
        QPushButton:hover { background: #555b65; }
        QPushButton:pressed { background: #23262b; padding-top: 12px; }

        QPushButton#primary {
            background: #4a86ff;
            color: white;
            font-weight: 600;
            border: none;
        }
    """
}


# ======================
# Helpers
# ======================

def load_list(path: Path):
    if not path.exists():
        return []
    return sorted(
        {x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()},
        key=str.lower
    )


def make_combo(items, placeholder):
    cb = QComboBox()
    cb.setEditable(True)
    cb.addItems(items)
    cb.setPlaceholderText(placeholder)
    cb.completer().setCaseSensitivity(Qt.CaseInsensitive)
    cb.completer().setFilterMode(Qt.MatchContains)
    return cb


# ======================
# Main Window
# ======================

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.settings = load_ui_settings()
        self.setWindowTitle("Sunglasses SEO PRO")
        self.resize(960, 680)

        root = QVBoxLayout(self)
        root.setSpacing(14)

        # ---------- SaaS Header Card ----------
        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)

        title = QLabel("🕶️ Sunglasses SEO PRO")
        title.setObjectName("title")
        subtitle = QLabel("Генерация живых SEO-описаний для маркетплейсов")
        subtitle.setObjectName("subtitle")

        card_layout.addWidget(title)
        card_layout.addWidget(subtitle)
        root.addWidget(card)

        # ---------- Theme selector ----------
        theme_row = QHBoxLayout()
        theme_row.addWidget(QLabel("🎨 Тема интерфейса"))
        self.cb_theme = QComboBox()
        self.cb_theme.addItems(THEMES.keys())
        self.cb_theme.setCurrentText(self.settings.get("theme", "Светлая"))
        self.cb_theme.currentTextChanged.connect(self.change_theme)
        theme_row.addWidget(self.cb_theme)
        root.addLayout(theme_row)

        # ---------- File ----------
        file_row = QHBoxLayout()
        btn_file = QPushButton("📂 Загрузить XLSX")
        btn_file.clicked.connect(self.pick_file)
        self.lbl_file = QLabel("Файл не выбран")
        self.lbl_file.setWordWrap(True)
        file_row.addWidget(btn_file)
        file_row.addWidget(self.lbl_file, 1)
        root.addLayout(file_row)

        # ---------- Combos ----------
        self.cb_brand = make_combo(load_list(BRANDS_FILE), "Бренд")
        root.addWidget(self.cb_brand)

        self.cb_shape = make_combo(load_list(SHAPES_FILE), "Форма оправы")
        root.addWidget(self.cb_shape)

        self.cb_lens = make_combo(load_list(LENSES_FILE), "Линзы")
        root.addWidget(self.cb_lens)

        self.cb_collection = make_combo(
            ["Весна–Лето 2025–2026", "Весна–Лето 2026"],
            "Коллекция"
        )
        root.addWidget(self.cb_collection)

        # ---------- Progress ----------
        self.progress = QProgressBar()
        self.progress.setValue(0)
        root.addWidget(self.progress)

        # ---------- Run ----------
        self.btn_run = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        self.btn_run.setObjectName("primary")
        self.btn_run.clicked.connect(self.run)
        root.addWidget(self.btn_run)

        self.apply_theme(self.settings.get("theme", "Светлая"))

    # ---------- Logic ----------

    def apply_theme(self, name):
        self.setStyleSheet(THEMES.get(name, ""))

    def change_theme(self, name):
        self.settings["theme"] = name
        save_ui_settings(self.settings)
        self.apply_theme(name)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX", "", "Excel (*.xlsx)")
        if path:
            self.input_file = path
            self.lbl_file.setText(path)

    def run(self):
        if not hasattr(self, "input_file"):
            QMessageBox.warning(self, "Ошибка", "Выберите XLSX файл")
            return

        self.progress.setValue(0)
        self.btn_run.setEnabled(False)

        # имитация прогресса
        self.timer = QTimer()
        self.timer.timeout.connect(self.fake_progress)
        self.timer.start(80)

        try:
            out, rows = fill_wb_template(
                input_xlsx=self.input_file,
                brand=self.cb_brand.currentText(),
                shape=self.cb_shape.currentText(),
                lens_features=self.cb_lens.currentText(),
                collection=self.cb_collection.currentText(),
            )
            QMessageBox.information(self, "Готово", f"Создан файл:\n{out}")
        finally:
            self.timer.stop()
            self.progress.setValue(100)
            self.btn_run.setEnabled(True)

    def fake_progress(self):
        if self.progress.value() < 90:
            self.progress.setValue(self.progress.value() + 2)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
