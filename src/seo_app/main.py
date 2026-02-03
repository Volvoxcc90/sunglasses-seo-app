import sys
import os
import json
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QProgressBar,
    QCheckBox
)

from seo_app.wb_fill import fill_wb_template


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
SETTINGS_FILE = DATA_DIR / "ui_settings.json"

BRANDS_FILE = DATA_DIR / "brands.txt"
SHAPES_FILE = DATA_DIR / "shapes.txt"
LENSES_FILE = DATA_DIR / "lenses.txt"

DATA_DIR.mkdir(exist_ok=True)


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
    value = (value or "").strip()
    if not value:
        return False
    items = load_list(path, [])
    if value in items:
        return False
    items.append(value)
    path.write_text("\n".join(items), encoding="utf-8")
    return True


def load_settings():
    if SETTINGS_FILE.exists():
        try:
            return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(data: dict):
    SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def open_folder(path: Path):
    try:
        os.startfile(str(path))
    except Exception:
        QMessageBox.warning(None, "Ошибка", f"Не удалось открыть папку:\n{path}")


class SeoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sunglasses SEO Generator v6")
        self.setMinimumWidth(560)

        self.settings = load_settings()

        self.brands = load_list(BRANDS_FILE, ["Cazal", "Gucci", "Prada", "Ray-Ban"])
        self.shapes = load_list(SHAPES_FILE, ["Квадратные", "Овальные", "Кошачий глаз"])
        self.lenses = load_list(LENSES_FILE, ["UV400", "Поляризационные", "Фотохромные"])

        self.build_ui()
        self.restore_settings()

    def build_ui(self):
        layout = QVBoxLayout(self)

        # --- combos with +
        self.brand_row, self.brand_cb = self.combo_row("Бренд", self.brands, BRANDS_FILE)
        self.shape_row, self.shape_cb = self.combo_row("Форма оправы", self.shapes, SHAPES_FILE)
        self.lens_row, self.lens_cb = self.combo_row("Линзы", self.lenses, LENSES_FILE)

        layout.addLayout(self.brand_row)
        layout.addLayout(self.shape_row)
        layout.addLayout(self.lens_row)

        # --- SEO level
        layout.addWidget(QLabel("SEO-плотность"))
        self.seo_level = QComboBox()
        self.seo_level.addItems(["soft", "normal", "hard"])
        layout.addWidget(self.seo_level)

        # --- length
        layout.addWidget(QLabel("Длина описания"))
        self.desc_length = QComboBox()
        self.desc_length.addItems(["short", "medium", "long"])
        layout.addWidget(self.desc_length)

        # --- style
        layout.addWidget(QLabel("Стиль текста"))
        self.style = QComboBox()
        self.style.addItems(["neutral", "premium", "social"])
        layout.addWidget(self.style)

        # --- WB Safe Mode
        self.cb_safe = QCheckBox("WB Safe Mode (убирает риск-слова: реплика/копия/люкс и т.п.)")
        layout.addWidget(self.cb_safe)

        # --- buttons
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

    def combo_row(self, label_text, items, file_path: Path):
        row = QHBoxLayout()
        row.addWidget(QLabel(label_text))

        combo = QComboBox()
        combo.setEditable(True)
        combo.addItems(items)
        row.addWidget(combo, 1)

        add_btn = QPushButton("+")
        add_btn.setFixedWidth(34)

        def add_item():
            val = combo.currentText().strip()
            if not val:
                QMessageBox.warning(self, "Пусто", "Введи значение и нажми +")
                return
            added = add_to_list(file_path, val)
            combo.clear()
            combo.addItems(load_list(file_path, []))
            combo.setCurrentText(val)
            QMessageBox.information(self, "Ок", "Добавлено" if added else "Уже было в списке")

        add_btn.clicked.connect(add_item)
        row.addWidget(add_btn)

        return row, combo

    def restore_settings(self):
        self.brand_cb.setCurrentText(self.settings.get("brand", ""))
        self.shape_cb.setCurrentText(self.settings.get("shape", ""))
        self.lens_cb.setCurrentText(self.settings.get("lens", ""))

        self.seo_level.setCurrentText(self.settings.get("seo_level", "normal"))
        self.desc_length.setCurrentText(self.settings.get("desc_length", "medium"))
        self.style.setCurrentText(self.settings.get("style", "neutral"))

        self.cb_safe.setChecked(bool(self.settings.get("wb_safe_mode", True)))

    def run_generation(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выбери Excel", "", "Excel (*.xlsx)")
        if not file_path:
            return

        brand = self.brand_cb.currentText().strip()
        shape = self.shape_cb.currentText().strip()
        lens = self.lens_cb.currentText().strip()

        seo = self.seo_level.currentText()
        length = self.desc_length.currentText()
        style = self.style.currentText()
        safe = self.cb_safe.isChecked()

        # сохраняем настройки
        self.settings.update({
            "brand": brand,
            "shape": shape,
            "lens": lens,
            "seo_level": seo,
            "desc_length": length,
            "style": style,
            "wb_safe_mode": safe
        })
        save_settings(self.settings)

        self.progress.setValue(0)

        try:
            out, count, report_json = fill_wb_template(
                input_xlsx=file_path,
                brand=brand,
                shape=shape,
                lens_features=lens,
                collection="Весна–Лето 2026",
                style=style,
                seo_level=seo,
                desc_length=length,
                wb_safe_mode=safe,
                progress_callback=self.progress.setValue
            )

            # покажем краткий итог по отчёту
            try:
                rep = json.loads(Path(report_json).read_text(encoding="utf-8"))
                labels = [r["seo"]["label"] for r in rep.get("rows", [])]
                green = labels.count("🟢 сильная")
                yellow = labels.count("🟡 норм")
                red = labels.count("🔴 слабая")
                msg = (
                    f"Создан файл:\n{out}\n"
                    f"Карточек: {count}\n\n"
                    f"SEO итог: 🟢 {green} | 🟡 {yellow} | 🔴 {red}\n"
                    f"Отчёт:\n{report_json}\n"
                    f"(рядом будет .seo_report.txt)"
                )
            except Exception:
                msg = f"Создан файл:\n{out}\nКарточек: {count}\nОтчёт:\n{report_json}"

            QMessageBox.information(self, "Готово (v6)", msg)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = SeoApp()
    w.show()
    sys.exit(app.exec())
