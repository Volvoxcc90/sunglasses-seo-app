import sys
import os
import json
import subprocess
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt

from seo_app.wb_fill import fill_wb_template


# ==========================
# Пути и файлы
# ==========================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
SETTINGS_FILE = DATA_DIR / "ui_settings.json"

BRANDS_FILE = DATA_DIR / "brands.txt"
SHAPES_FILE = DATA_DIR / "shapes.txt"
LENSES_FILE = DATA_DIR / "lenses.txt"

DATA_DIR.mkdir(exist_ok=True)


# ==========================
# Утилиты
# ==========================
def load_list(path: Path, defaults):
    if not path.exists():
        path.write_text("\n".join(defaults), encoding="utf-8")
        return defaults
    items = [x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()]
    if not items:
        items = defaults
        path.write_text("\n".join(items), encoding="utf-8")
    return items


def add_to_list(path: Path, value: str):
    value = value.strip()
    if not value:
        return
    items = load_list(path, [])
    if value not in items:
        items.append(value)
        path.write_text("\n".join(items), encoding="utf-8")


def load_settings():
    if SETTINGS_FILE.exists():
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    return {}


def save_settings(data: dict):
    SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def open_folder(path: Path):
    if sys.platform.startswith("win"):
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.call(["open", path])
    else:
        subprocess.call(["xdg-open", path])


# ==========================
# GUI
# ==========================
class SeoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sunglasses SEO Generator")
        self.setMinimumWidth(520)

        self.settings = load_settings()

        self.brands = load_list(BRANDS_FILE, ["Cazal", "Gucci", "Prada", "Ray-Ban"])
        self.shapes = load_list(SHAPES_FILE, ["Квадратные", "Овальные", "Кошачий глаз"])
        self.lenses = load_list(LENSES_FILE, ["UV400", "Поляризационные", "Фотохромные"])

        self.build_ui()
        self.restore_settings()

    def build_ui(self):
        layout = QVBoxLayout(self)

        # ===== Бренд =====
        self.brand_box = self.combo_row("Бренд", self.brands, BRANDS_FILE)
        layout.addLayout(self.brand_box[0])

        # ===== Форма =====
        self.shape_box = self.combo_row("Форма оправы", self.shapes, SHAPES_FILE)
        layout.addLayout(self.shape_box[0])

        # ===== Линзы =====
        self.lens_box = self.combo_row("Линзы", self.lenses, LENSES_FILE)
        layout.addLayout(self.lens_box[0])

        # ===== SEO уровень =====
        self.seo_level = QComboBox()
        self.seo_level.addItems(["soft", "normal", "hard"])
        layout.addWidget(QLabel("SEO-плотность"))
        layout.addWidget(self.seo_level)

        # ===== Длина =====
        self.desc_length = QComboBox()
        self.desc_length.addItems(["short", "medium", "long"])
        layout.addWidget(QLabel("Длина описания"))
        layout.addWidget(self.desc_length)

        # ===== Кнопки =====
        btn_row = QHBoxLayout()

        open_btn = QPushButton("📂 Папка data")
        open_btn.clicked.connect(lambda: open_folder(DATA_DIR))
        btn_row.addWidget(open_btn)

        gen_btn = QPushButton("🚀 Готово")
        gen_btn.clicked.connect(self.run_generation)
        btn_row.addWidget(gen_btn)

        layout.addLayout(btn_row)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

    def combo_row(self, label_text, items, file_path):
        layout = QHBoxLayout()
        label = QLabel(label_text)
        combo = QComboBox()
        combo.setEditable(True)
        combo.addItems(items)

        add_btn = QPushButton("+")
        add_btn.setFixedWidth(32)

        def add_item():
            val = combo.currentText().strip()
            if not val:
                return
            add_to_list(file_path, val)
            combo.clear()
            combo.addItems(load_list(file_path, []))
            combo.setCurrentText(val)

        add_btn.clicked.connect(add_item)

        layout.addWidget(label)
        layout.addWidget(combo)
        layout.addWidget(add_btn)

        return layout, combo

    def restore_settings(self):
        self.brand_box[1].setCurrentText(self.settings.get("brand", ""))
        self.shape_box[1].setCurrentText(self.settings.get("shape", ""))
        self.lens_box[1].setCurrentText(self.settings.get("lens", ""))
        self.seo_level.setCurrentText(self.settings.get("seo_level", "normal"))
        self.desc_length.setCurrentText(self.settings.get("desc_length", "medium"))

    def run_generation(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выбери Excel", "", "Excel (*.xlsx)")
        if not file_path:
            return

        brand = self.brand_box[1].currentText()
        shape = self.shape_box[1].currentText()
        lens = self.lens_box[1].currentText()
        seo = self.seo_level.currentText()
        length = self.desc_length.currentText()

        self.settings.update({
            "brand": brand,
            "shape": shape,
            "lens": lens,
            "seo_level": seo,
            "desc_length": length
        })
        save_settings(self.settings)

        try:
            out, count = fill_wb_template(
                input_xlsx=file_path,
                brand=brand,
                shape=shape,
                lens_features=lens,
                collection="Весна–Лето 2026",
                seo_level=seo,
                desc_length=length,
                progress_callback=self.progress.setValue
            )
            QMessageBox.information(self, "Готово", f"Создан файл:\n{out}\nКарточек: {count}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


# ==========================
# RUN
# ==========================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = SeoApp()
    w.show()
    sys.exit(app.exec())
