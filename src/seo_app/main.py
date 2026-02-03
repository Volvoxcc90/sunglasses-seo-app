# main.py
import sys
import os
import json
import re
import inspect
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QLineEdit,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QMessageBox,
    QProgressBar, QGroupBox
)
from PyQt5.QtCore import QThread, pyqtSignal

from wb_fill import fill_wb_template


APP_NAME = "Sunglasses SEO PRO"


# -------------------------------
# DATA DIR
# -------------------------------

def app_data_dir() -> Path:
    base = Path(os.getenv("APPDATA", str(Path.home())))
    p = base / APP_NAME / "data"
    p.mkdir(parents=True, exist_ok=True)
    return p


# -------------------------------
# BRAND MAP (AUTO)
# -------------------------------

def normalize_brand_key(brand: str) -> str:
    b = (brand or "").strip().lower()
    b = b.replace("-", " ").replace("&", " ")
    b = re.sub(r"\s+", " ", b).strip()
    return b


TRANSLIT = [
    ("sch","ш"),("sh","ш"),("ch","ч"),("ya","я"),("yu","ю"),("yo","ё"),
    ("kh","х"),("ts","ц"),("ph","ф"),("th","т"),
    ("a","а"),("b","б"),("c","к"),("d","д"),("e","е"),("f","ф"),
    ("g","г"),("h","х"),("i","и"),("j","дж"),("k","к"),("l","л"),
    ("m","м"),("n","н"),("o","о"),("p","п"),("q","к"),("r","р"),
    ("s","с"),("t","т"),("u","у"),("v","в"),("w","в"),("x","кс"),
    ("y","и"),("z","з"),
]

def guess_ru(brand: str) -> str:
    if re.search(r"[А-Яа-яЁё]", brand):
        return brand
    key = normalize_brand_key(brand)
    out = []
    for w in key.split():
        ww = w
        for a,b in TRANSLIT:
            ww = ww.replace(a,b)
        out.append(ww)
    return " ".join(x.capitalize() for x in out)


def load_brands_ru() -> dict:
    p = app_data_dir() / "brands_ru.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_brands_ru(m: dict):
    p = app_data_dir() / "brands_ru.json"
    p.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")


def auto_update_brand_map(brands: list):
    m = load_brands_ru()
    changed = False
    for b in brands:
        key = normalize_brand_key(b)
        if key and key not in m:
            m[key] = guess_ru(b)
            changed = True
    if changed:
        save_brands_ru(m)


# -------------------------------
# LIST FILES
# -------------------------------

def ensure_list(filename: str, defaults: list) -> Path:
    p = app_data_dir() / filename
    if not p.exists():
        p.write_text("\n".join(defaults), encoding="utf-8")
    return p


def load_list(filename: str, defaults: list) -> list:
    p = ensure_list(filename, defaults)
    return [x.strip() for x in p.read_text(encoding="utf-8").splitlines() if x.strip()]


def add_to_list(filename: str, value: str):
    value = value.strip()
    if not value:
        return
    p = app_data_dir() / filename
    items = load_list(filename, [])
    if value not in items:
        with p.open("a", encoding="utf-8") as f:
            f.write("\n" + value)


# -------------------------------
# WORKER
# -------------------------------

class Worker(QThread):
    progress = pyqtSignal(int)
    done = pyqtSignal(str, int, str)
    error = pyqtSignal(str)

    def __init__(self, args: dict):
        super().__init__()
        self.args = args

    def run(self):
        try:
            sig = inspect.signature(fill_wb_template)
            allowed = sig.parameters.keys()
            safe_args = {k: v for k, v in self.args.items() if k in allowed}

            if "progress_callback" in allowed:
                safe_args["progress_callback"] = lambda p: self.progress.emit(int(p))

            result = fill_wb_template(**safe_args)
            out, count, report = result[0], result[1], result[2]
            self.done.emit(out, count, report)
        except Exception as e:
            self.error.emit(str(e))


# -------------------------------
# UI
# -------------------------------

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(900, 520)

        self.data_dir = app_data_dir()

        self.brands = load_list("brands.txt", ["Gucci", "Prada", "Miu Miu"])
        auto_update_brand_map(self.brands)

        self._build()

    def _build(self):
        layout = QVBoxLayout(self)

        title = QLabel("🕶 Sunglasses SEO PRO")
        title.setStyleSheet("font-size:22px;font-weight:700;")
        layout.addWidget(title)

        box = QGroupBox("Параметры")
        grid = QGridLayout(box)

        grid.addWidget(QLabel("Бренд"), 0, 0)
        self.brand = QComboBox()
        self.brand.setEditable(True)
        self.brand.addItems(self.brands)
        grid.addWidget(self.brand, 0, 1)

        plus = QPushButton("+")
        plus.clicked.connect(self.add_brand)
        grid.addWidget(plus, 0, 2)

        grid.addWidget(QLabel("Коллекция"), 1, 0)
        self.collection = QLineEdit("Весна–Лето 2026")
        grid.addWidget(self.collection, 1, 1, 1, 2)

        layout.addWidget(box)

        file_row = QHBoxLayout()
        self.file_lbl = QLabel("Файл не выбран")
        pick = QPushButton("📄 XLSX")
        pick.clicked.connect(self.pick_file)
        file_row.addWidget(pick)
        file_row.addWidget(self.file_lbl)
        layout.addLayout(file_row)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        run = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        run.clicked.connect(self.run)
        layout.addWidget(run)

    def add_brand(self):
        val = self.brand.currentText().strip()
        if not val:
            return
        add_to_list("brands.txt", val)
        auto_update_brand_map([val])
        self.brands = load_list("brands.txt", [])
        self.brand.clear()
        self.brand.addItems(self.brands)
        self.brand.setCurrentText(val)
        QMessageBox.information(self, "Готово", f"Добавлено: {val}")

    def pick_file(self):
        fp, _ = QFileDialog.getOpenFileName(self, "XLSX", "", "Excel (*.xlsx)")
        if fp:
            self.file = fp
            self.file_lbl.setText(fp)

    def run(self):
        if not hasattr(self, "file"):
            QMessageBox.warning(self, "Ошибка", "Выбери файл")
            return

        args = {
            "input_xlsx": self.file,
            "brand": self.brand.currentText(),
            "collection": self.collection.text()
        }

        self.worker = Worker(args)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.done.connect(lambda *_: QMessageBox.information(self, "Готово", "Файл создан"))
        self.worker.error.connect(lambda e: QMessageBox.critical(self, "Ошибка", e))
        self.worker.start()


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
