# main.py
import sys
import os
import json
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog, QLineEdit,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QMessageBox,
    QProgressBar, QFrame, QCheckBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt

# ВАЖНО: wb_fill.py должен лежать рядом с main.py
from wb_fill import fill_wb_template  # noqa


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


def load_settings() -> dict:
    p = settings_path()
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(d: dict):
    try:
        settings_path().write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def ensure_list_file(path: Path, defaults: list):
    if not path.exists():
        path.write_text("\n".join(defaults) + "\n", encoding="utf-8")


def load_list(path: Path, defaults: list) -> list:
    ensure_list_file(path, defaults)
    lines = []
    for s in path.read_text(encoding="utf-8").splitlines():
        s = s.strip()
        if s:
            lines.append(s)
    # уникализируем и сортируем слегка
    out = []
    seen = set()
    for x in lines:
        if x.lower() not in seen:
            seen.add(x.lower())
            out.append(x)
    return out


def add_to_list(path: Path, value: str):
    value = (value or "").strip()
    if not value:
        return
    ensure_list_file(path, [])
    items = load_list(path, [])
    if value.lower() in {x.lower() for x in items}:
        return
    with path.open("a", encoding="utf-8") as f:
        f.write(value + "\n")


# -------------------------------
# THEMES (QSS)
# -------------------------------
THEMES = {
    "Midnight": """
        QWidget { background:#0b1220; color:#e7eefc; font-size:13px; }
        QLabel#Title { font-size:22px; font-weight:900; }
        QLabel#Subtitle { color:#aab8d6; }
        QLabel#Muted { color:#97a6c7; }
        QLabel#Section { font-weight:800; }
        QLabel#Pill { background:#16213b; border:1px solid #24345c; border-radius:10px; padding:6px 10px; color:#cfe1ff; }
        QFrame#Card { background:#0f1a2e; border:1px solid #1f2b46; border-radius:14px; }
        QLineEdit, QComboBox {
            background:#0b1426; border:1px solid #1f2b46; border-radius:10px; padding:8px;
        }
        QComboBox::drop-down { border:0; width:26px; }
        QComboBox::down-arrow {
            image:none;
            border-left:6px solid transparent;
            border-right:6px solid transparent;
            border-top:8px solid #cfe1ff;
            margin-right:8px;
        }
        QPushButton#Primary {
            background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6d28d9, stop:1 #8b5cf6);
            border:0; border-radius:12px; padding:10px 14px; font-weight:800; color:#ffffff;
        }
        QPushButton#Ghost { background:#16213b; border:1px solid #24345c; border-radius:12px; padding:10px 14px; font-weight:700; }
        QPushButton#Plus { background:#16213b; border:1px solid #24345c; border-radius:10px; padding:8px 12px; font-weight:900; }
        QPushButton:hover { opacity:0.96; }
        QPushButton:disabled { background:#2a3350; color:#9aa7c6; }
        QProgressBar { background:#0b1426; border:1px solid #1f2b46; border-radius:10px; text-align:center; }
        QProgressBar::chunk { background:#8b5cf6; border-radius:10px; }
        QCheckBox { spacing:8px; }
    """,
    "Graphite": """
        QWidget { background:#0f0f12; color:#f2f2f2; font-size:13px; }
        QLabel#Title { font-size:22px; font-weight:900; }
        QLabel#Subtitle { color:#bdbdbd; }
        QLabel#Muted { color:#b0b0b0; }
        QLabel#Section { font-weight:800; }
        QLabel#Pill { background:#1a1a1f; border:1px solid #2a2a33; border-radius:10px; padding:6px 10px; }
        QFrame#Card { background:#17171c; border:1px solid #2a2a33; border-radius:14px; }
        QLineEdit, QComboBox { background:#121217; border:1px solid #2a2a33; border-radius:10px; padding:8px; }
        QComboBox::drop-down { border:0; width:26px; }
        QComboBox::down-arrow { image:none; border-left:6px solid transparent; border-right:6px solid transparent; border-top:8px solid #eaeaea; margin-right:8px; }
        QPushButton#Primary { background:#2d6cdf; border:0; border-radius:12px; padding:10px 14px; font-weight:800; color:#fff; }
        QPushButton#Ghost { background:#1a1a1f; border:1px solid #2a2a33; border-radius:12px; padding:10px 14px; font-weight:700; }
        QPushButton#Plus { background:#1a1a1f; border:1px solid #2a2a33; border-radius:10px; padding:8px 12px; font-weight:900; }
        QProgressBar { background:#121217; border:1px solid #2a2a33; border-radius:10px; text-align:center; }
        QProgressBar::chunk { background:#2d6cdf; border-radius:10px; }
    """,
    "Light": """
        QWidget { background:#f6f7fb; color:#12131a; font-size:13px; }
        QLabel#Title { font-size:22px; font-weight:900; }
        QLabel#Subtitle { color:#55607a; }
        QLabel#Muted { color:#5f6b84; }
        QLabel#Section { font-weight:800; }
        QLabel#Pill { background:#eef3ff; border:1px solid #d7e2ff; border-radius:10px; padding:6px 10px; color:#224cff; }
        QFrame#Card { background:#ffffff; border:1px solid #dfe5f1; border-radius:14px; }
        QLineEdit, QComboBox { background:#ffffff; border:1px solid #dfe5f1; border-radius:10px; padding:8px; }
        QComboBox::drop-down { border:0; width:26px; }
        QComboBox::down-arrow { image:none; border-left:6px solid transparent; border-right:6px solid transparent; border-top:8px solid #2b61ff; margin-right:8px; }
        QPushButton#Primary { background:#2b61ff; border:0; border-radius:12px; padding:10px 14px; font-weight:800; color:#fff; }
        QPushButton#Ghost { background:#eef3ff; border:1px solid #d7e2ff; border-radius:12px; padding:10px 14px; font-weight:700; }
        QPushButton#Plus { background:#eef3ff; border:1px solid #d7e2ff; border-radius:10px; padding:8px 12px; font-weight:900; }
        QProgressBar { background:#ffffff; border:1px solid #dfe5f1; border-radius:10px; text-align:center; }
        QProgressBar::chunk { background:#2b61ff; border-radius:10px; }
    """,
}


# -------------------------------
# Worker thread
# -------------------------------
class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, int, dict)
    failed = pyqtSignal(str)

    def __init__(self, args: dict):
        super().__init__()
        self.args = args

    def run(self):
        try:
            out, count, report = fill_wb_template(
                progress_callback=lambda p: self.progress.emit(int(p)),
                **self.args
            )
            self.finished.emit(out, count, report)
        except Exception as e:
            self.failed.emit(str(e))


# -------------------------------
# UI helpers
# -------------------------------
def card() -> QFrame:
    f = QFrame()
    f.setObjectName("Card")
    f.setFrameShape(QFrame.NoFrame)
    return f


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumWidth(860)

        self.data_dir = app_data_dir()
        self.settings = load_settings()

        # list files in data dir
        self.brands_file = self.data_dir / "brands.txt"
        self.shapes_file = self.data_dir / "shapes.txt"
        self.lenses_file = self.data_dir / "lenses.txt"

        ensure_list_file(self.brands_file, ["Ray-Ban", "Gucci", "Prada", "Cazal", "Miu Miu"])
        ensure_list_file(self.shapes_file, ["квадратные", "авиаторы", "овальные", "кошачий глаз", "круглые"])
        ensure_list_file(self.lenses_file, ["UV400", "поляризационные", "фотохромные", "градиентные"])

        self.xlsx_path = ""

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        # ---------- Header card ----------
        header = card()
        hl = QVBoxLayout(header)
        hl.setContentsMargins(16, 14, 16, 14)
        title = QLabel("🕶️ Sunglasses SEO PRO")
        title.setObjectName("Title")
        subtitle = QLabel("Живые SEO-описания • Выпадающие списки • Прогресс • Темы • WB Safe/Strict • AUTO-пол")
        subtitle.setObjectName("Subtitle")
        hl.addWidget(title)
        hl.addWidget(subtitle)
        root.addWidget(header)

        # ---------- Top bar card ----------
        top = card()
        tl = QGridLayout(top)
        tl.setContentsMargins(16, 14, 16, 14)
        tl.setHorizontalSpacing(12)
        tl.setVerticalSpacing(10)

        lbl_theme = QLabel("🎨 Тема")
        lbl_theme.setObjectName("Section")
        self.cmb_theme = QComboBox()
        self.cmb_theme.addItems(list(THEMES.keys()))
        saved_theme = self.settings.get("theme", "Midnight")
        if saved_theme in THEMES:
            self.cmb_theme.setCurrentText(saved_theme)
        self.cmb_theme.currentTextChanged.connect(self.apply_theme)

        lbl_dir = QLabel("📁 Справочники:")
        lbl_dir.setObjectName("Section")
        self.lbl_data = QLabel(str(self.data_dir))
        self.lbl_data.setObjectName("Pill")
        btn_folder = QPushButton("Папка")
        btn_folder.setObjectName("Ghost")
        btn_folder.clicked.connect(self.open_data_folder)

        btn_xlsx = QPushButton("⬆️ Загрузить XLSX")
        btn_xlsx.setObjectName("Ghost")
        btn_xlsx.clicked.connect(self.pick_xlsx)
        self.lbl_xlsx = QLabel("Файл не выбран")
        self.lbl_xlsx.setObjectName("Muted")

        tl.addWidget(lbl_theme, 0, 0)
        tl.addWidget(self.cmb_theme, 0, 1)
        tl.addWidget(lbl_dir, 0, 2)
        tl.addWidget(self.lbl_data, 0, 3)
        tl.addWidget(btn_folder, 0, 4)

        tl.addWidget(btn_xlsx, 1, 0, 1, 2)
        tl.addWidget(self.lbl_xlsx, 1, 2, 1, 3)

        root.addWidget(top)

        # ---------- Main controls card ----------
        main = card()
        ml = QGridLayout(main)
        ml.setContentsMargins(16, 14, 16, 14)
        ml.setHorizontalSpacing(12)
        ml.setVerticalSpacing(10)

        # dropdowns
        self.cmb_brand = QComboBox()
        self.cmb_shape = QComboBox()
        self.cmb_lens = QComboBox()

        self.reload_lists()

        self.ed_collection = QLineEdit()
        self.ed_collection.setText(self.settings.get("collection", "Весна–Лето 2026"))

        btn_add_brand = QPushButton("+")
        btn_add_brand.setObjectName("Plus")
        btn_add_brand.clicked.connect(lambda: self.add_item_dialog("Бренд", self.brands_file, self.cmb_brand))

        btn_add_shape = QPushButton("+")
        btn_add_shape.setObjectName("Plus")
        btn_add_shape.clicked.connect(lambda: self.add_item_dialog("Форма оправы", self.shapes_file, self.cmb_shape))

        btn_add_lens = QPushButton("+")
        btn_add_lens.setObjectName("Plus")
        btn_add_lens.clicked.connect(lambda: self.add_item_dialog("Линзы", self.lenses_file, self.cmb_lens))

        ml.addWidget(QLabel("Бренд"), 0, 0)
        ml.addWidget(self.cmb_brand, 0, 1)
        ml.addWidget(btn_add_brand, 0, 2)

        ml.addWidget(QLabel("Форма оправы"), 1, 0)
        ml.addWidget(self.cmb_shape, 1, 1)
        ml.addWidget(btn_add_shape, 1, 2)

        ml.addWidget(QLabel("Линзы"), 2, 0)
        ml.addWidget(self.cmb_lens, 2, 1)
        ml.addWidget(btn_add_lens, 2, 2)

        ml.addWidget(QLabel("Коллекция"), 3, 0)
        ml.addWidget(self.ed_collection, 3, 1, 1, 2)

        # seo controls
        self.cmb_seo = QComboBox()
        self.cmb_seo.addItems(["low", "normal", "high"])
        self.cmb_seo.setCurrentText(self.settings.get("seo_level", "normal"))

        self.cmb_gender = QComboBox()
        self.cmb_gender.addItems(["Auto", "Женские", "Мужские", "Унисекс"])
        self.cmb_gender.setCurrentText(self.settings.get("gender_mode", "Auto"))

        self.chk_safe = QCheckBox("WB Safe Mode (заменяет риск-слова)")
        self.chk_safe.setChecked(bool(self.settings.get("wb_safe_mode", True)))

        self.chk_strict = QCheckBox("WB Strict (убирает обещания/абсолюты)")
        self.chk_strict.setChecked(bool(self.settings.get("wb_strict", True)))

        ml.addWidget(QLabel("SEO-плотность"), 4, 0)
        ml.addWidget(self.cmb_seo, 4, 1, 1, 2)

        ml.addWidget(QLabel("AUTO-пол"), 5, 0)
        ml.addWidget(self.cmb_gender, 5, 1, 1, 2)

        ml.addWidget(self.chk_safe, 6, 0, 1, 3)
        ml.addWidget(self.chk_strict, 7, 0, 1, 3)

        root.addWidget(main)

        # ---------- Bottom bar card ----------
        bottom = card()
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(16, 14, 16, 14)
        bl.setSpacing(12)

        self.progress = QProgressBar()
        self.progress.setValue(0)

        self.btn_run = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        self.btn_run.setObjectName("Primary")
        self.btn_run.clicked.connect(self.start)

        bl.addWidget(self.progress, 2)
        bl.addWidget(self.btn_run, 1)

        root.addWidget(bottom)

        # apply theme
        self.apply_theme(self.cmb_theme.currentText())

    # ------------------
    def apply_theme(self, name: str):
        qss = THEMES.get(name, "")
        QApplication.instance().setStyleSheet(qss)
        self.settings["theme"] = name
        save_settings(self.settings)

    def open_data_folder(self):
        # Windows explorer open
        try:
            os.startfile(str(self.data_dir))
        except Exception:
            QMessageBox.information(self, "Папка", str(self.data_dir))

    def pick_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX", "", "Excel (*.xlsx)")
        if path:
            self.xlsx_path = path
            self.lbl_xlsx.setText(Path(path).name)

    def reload_lists(self):
        brands = load_list(self.brands_file, ["Ray-Ban"])
        shapes = load_list(self.shapes_file, ["квадратные"])
        lenses = load_list(self.lenses_file, ["UV400"])

        self.cmb_brand.clear()
        self.cmb_shape.clear()
        self.cmb_lens.clear()

        self.cmb_brand.addItems(brands)
        self.cmb_shape.addItems(shapes)
        self.cmb_lens.addItems(lenses)

        # restore last
        last_brand = self.settings.get("brand", "")
        last_shape = self.settings.get("shape", "")
        last_lens = self.settings.get("lens", "")
        if last_brand in brands:
            self.cmb_brand.setCurrentText(last_brand)
        if last_shape in shapes:
            self.cmb_shape.setCurrentText(last_shape)
        if last_lens in lenses:
            self.cmb_lens.setCurrentText(last_lens)

    def add_item_dialog(self, title: str, file_path: Path, combo: QComboBox):
        from PyQt5.QtWidgets import QInputDialog
        val, ok = QInputDialog.getText(self, f"Добавить: {title}", f"Введите {title.lower()}:")
        if not ok:
            return
        val = (val or "").strip()
        if not val:
            return
        add_to_list(file_path, val)
        self.reload_lists()
        combo.setCurrentText(val)

    # ------------------
    def start(self):
        if not self.xlsx_path:
            QMessageBox.warning(self, "Файл", "Сначала выбери XLSX файл.")
            return

        brand = self.cmb_brand.currentText().strip()
        shape = self.cmb_shape.currentText().strip()
        lens = self.cmb_lens.currentText().strip()
        collection = self.ed_collection.text().strip()
        seo_level = self.cmb_seo.currentText()
        gender_mode = self.cmb_gender.currentText()
        wb_safe = self.chk_safe.isChecked()
        wb_strict = self.chk_strict.isChecked()

        # persist choices
        self.settings.update({
            "brand": brand,
            "shape": shape,
            "lens": lens,
            "collection": collection,
            "seo_level": seo_level,
            "gender_mode": gender_mode,
            "wb_safe_mode": wb_safe,
            "wb_strict": wb_strict,
        })
        save_settings(self.settings)

        self.progress.setValue(0)
        self.btn_run.setEnabled(False)

        args = dict(
            input_xlsx=self.xlsx_path,
            brand_lat=brand,
            shape=shape,
            lens=lens,
            collection=collection,
            style="neutral",
            desc_length="medium",
            seo_level=seo_level,
            gender_mode=gender_mode,
            wb_safe_mode=wb_safe,
            wb_strict=wb_strict,
            uniq_strength=85 if seo_level == "high" else (75 if seo_level == "normal" else 65),
            data_dir=str(self.data_dir),
        )

        self.worker = Worker(args)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.done)
        self.worker.failed.connect(self.fail)
        self.worker.start()

    def done(self, out_path: str, count: int, report: dict):
        self.btn_run.setEnabled(True)
        self.progress.setValue(100)
        QMessageBox.information(
            self,
            "Готово",
            f"Заполнено строк: {count}\nФайл: {out_path}\n\nОтчёт: {json.dumps(report, ensure_ascii=False)}"
        )

    def fail(self, msg: str):
        self.btn_run.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", msg)


def main():
    app = QApplication(sys.argv)
    w = App()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
