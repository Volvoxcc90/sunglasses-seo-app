# main.py
import sys
import os
import json
import traceback
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QGridLayout, QComboBox, QMessageBox,
    QProgressBar, QFrame, QCheckBox, QLineEdit, QSpinBox, QTextEdit, QDialog
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt

from wb_fill import fill_wb_template, generate_preview


APP_NAME = "Sunglasses SEO PRO"


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
    out = []
    seen = set()
    for x in lines:
        k = x.lower().strip()
        if k not in seen:
            seen.add(k)
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


THEMES = {
    "Midnight": """
        QWidget { background:#0b1220; color:#e7eefc; font-size:13px; }
        QLabel#Title { font-size:22px; font-weight:900; }
        QLabel#Subtitle { color:#aab8d6; }
        QLabel#Muted { color:#97a6c7; }
        QFrame#Card { background:#0f1a2e; border:1px solid #1f2b46; border-radius:14px; }
        QLineEdit, QComboBox, QTextEdit, QSpinBox {
            background:#0b1426; border:1px solid #1f2b46; border-radius:10px; padding:8px;
        }
        QComboBox::drop-down { border:0; width:28px; }
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
        QPushButton#Mini { background:#0b1426; border:1px solid #1f2b46; border-radius:10px; padding:8px 10px; font-weight:800; }
        QProgressBar { background:#0b1426; border:1px solid #1f2b46; border-radius:10px; text-align:center; }
        QProgressBar::chunk { background:#8b5cf6; border-radius:10px; }
    """,
    "Light": """
        QWidget { background:#f6f7fb; color:#12131a; font-size:13px; }
        QLabel#Title { font-size:22px; font-weight:900; }
        QLabel#Subtitle { color:#55607a; }
        QLabel#Muted { color:#5f6b84; }
        QFrame#Card { background:#ffffff; border:1px solid #dfe5f1; border-radius:14px; }
        QLineEdit, QComboBox, QTextEdit, QSpinBox { background:#ffffff; border:1px solid #dfe5f1; border-radius:10px; padding:8px; }
        QComboBox::drop-down { border:0; width:28px; }
        QComboBox::down-arrow { image:none; border-left:6px solid transparent; border-right:6px solid transparent; border-top:8px solid #2b61ff; margin-right:8px; }
        QPushButton#Primary { background:#2b61ff; border:0; border-radius:12px; padding:10px 14px; font-weight:800; color:#fff; }
        QPushButton#Ghost { background:#eef3ff; border:1px solid #d7e2ff; border-radius:12px; padding:10px 14px; font-weight:700; }
        QPushButton#Mini { background:#ffffff; border:1px solid #dfe5f1; border-radius:10px; padding:8px 10px; font-weight:800; }
        QProgressBar { background:#ffffff; border:1px solid #dfe5f1; border-radius:10px; text-align:center; }
        QProgressBar::chunk { background:#2b61ff; border-radius:10px; }
    """
}


def card() -> QFrame:
    f = QFrame()
    f.setObjectName("Card")
    f.setFrameShape(QFrame.NoFrame)
    return f


class PreviewDialog(QDialog):
    def __init__(self, parent, items: list):
        super().__init__(parent)
        self.setWindowTitle("Примеры (3)")
        self.setMinimumWidth(820)
        self.setMinimumHeight(520)

        lay = QVBoxLayout(self)
        text = QTextEdit()
        text.setReadOnly(True)

        out = []
        for i, (t, d) in enumerate(items, 1):
            out.append(f"{i}) НАИМЕНОВАНИЕ:\n{t}\n\nОПИСАНИЕ:\n{d}\n\n" + ("-"*70))
        text.setPlainText("\n\n".join(out))

        lay.addWidget(text)
        btn = QPushButton("OK")
        btn.setObjectName("Mini")
        btn.clicked.connect(self.accept)
        lay.addWidget(btn, alignment=Qt.AlignRight)


class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, args: dict):
        super().__init__()
        self.args = args

    def run(self):
        try:
            def p_cb(p):
                self.progress.emit(int(p))
            out, _rows, _rep = fill_wb_template(progress_callback=p_cb, **self.args)
            self.finished.emit(out)
        except Exception as e:
            self.failed.emit(str(e))


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumWidth(920)

        self.data_dir = app_data_dir()
        self.settings = load_settings()
        self.xlsx_path = ""

        self.brands_file = self.data_dir / "brands.txt"
        self.shapes_file = self.data_dir / "shapes.txt"
        self.lenses_file = self.data_dir / "lenses.txt"

        ensure_list_file(self.brands_file, ["Ray-Ban", "Gucci", "Prada", "Cazal", "Miu Miu"])
        ensure_list_file(self.shapes_file, ["квадратные", "авиаторы", "овальные", "кошачий глаз", "круглые"])
        ensure_list_file(self.lenses_file, ["UV400", "поляризационные", "фотохромные", "градиентные"])

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(12)

        header = card()
        hl = QVBoxLayout(header)
        hl.setContentsMargins(16, 14, 16, 14)
        title = QLabel("🕶️ Sunglasses SEO PRO")
        title.setObjectName("Title")
        subtitle = QLabel("Заполняет Наименование + Описание (ровно 6 строк), с живым SEO и реальным рандомом")
        subtitle.setObjectName("Subtitle")
        hl.addWidget(title)
        hl.addWidget(subtitle)
        root.addWidget(header)

        top = card()
        tl = QGridLayout(top)
        tl.setContentsMargins(16, 14, 16, 14)
        tl.setHorizontalSpacing(12)
        tl.setVerticalSpacing(10)

        tl.addWidget(QLabel("🎨 Тема"), 0, 0)
        self.cmb_theme = QComboBox()
        self.cmb_theme.addItems(list(THEMES.keys()))
        self.cmb_theme.setCurrentText(self.settings.get("theme", "Midnight"))
        self.cmb_theme.currentTextChanged.connect(self.apply_theme)
        tl.addWidget(self.cmb_theme, 0, 1)

        tl.addWidget(QLabel("📁 Data"), 0, 2)
        self.lbl_data = QLabel(str(self.data_dir))
        self.lbl_data.setObjectName("Muted")
        tl.addWidget(self.lbl_data, 0, 3)

        btn_folder = QPushButton("Папка")
        btn_folder.setObjectName("Ghost")
        btn_folder.clicked.connect(self.open_data_folder)
        tl.addWidget(btn_folder, 0, 4)

        btn_xlsx = QPushButton("⬆️ Загрузить XLSX")
        btn_xlsx.setObjectName("Ghost")
        btn_xlsx.clicked.connect(self.pick_xlsx)
        tl.addWidget(btn_xlsx, 1, 0, 1, 2)

        self.lbl_xlsx = QLabel("Файл не выбран")
        self.lbl_xlsx.setObjectName("Muted")
        tl.addWidget(self.lbl_xlsx, 1, 2, 1, 2)

        btn_prev = QPushButton("👀 Примеры")
        btn_prev.setObjectName("Mini")
        btn_prev.clicked.connect(self.preview)
        tl.addWidget(btn_prev, 1, 4)

        root.addWidget(top)

        main = card()
        ml = QGridLayout(main)
        ml.setContentsMargins(16, 14, 16, 14)
        ml.setHorizontalSpacing(12)
        ml.setVerticalSpacing(10)

        self.cmb_brand = QComboBox()
        self.cmb_shape = QComboBox()
        self.cmb_lens = QComboBox()

        # ВАЖНО: editable — можно печатать вручную (как ты хочешь)
        self.cmb_brand.setEditable(True)
        self.cmb_shape.setEditable(True)
        self.cmb_lens.setEditable(True)

        self.reload_lists()

        ml.addWidget(QLabel("Бренд (латиницей)"), 0, 0)
        ml.addWidget(self.cmb_brand, 0, 1, 1, 2)

        ml.addWidget(QLabel("Форма оправы"), 1, 0)
        ml.addWidget(self.cmb_shape, 1, 1, 1, 2)

        ml.addWidget(QLabel("Линзы"), 2, 0)
        ml.addWidget(self.cmb_lens, 2, 1, 1, 2)

        ml.addWidget(QLabel("Коллекция"), 3, 0)
        self.ed_collection = QLineEdit(self.settings.get("collection", "Весна–Лето 2026"))
        ml.addWidget(self.ed_collection, 3, 1, 1, 2)

        ml.addWidget(QLabel("SEO"), 4, 0)
        self.cmb_seo = QComboBox()
        self.cmb_seo.addItems(["low", "normal", "high"])
        self.cmb_seo.setCurrentText(self.settings.get("seo_level", "high"))
        ml.addWidget(self.cmb_seo, 4, 1, 1, 2)

        ml.addWidget(QLabel("Пол"), 5, 0)
        self.cmb_gender = QComboBox()
        self.cmb_gender.addItems(["Auto", "Женские", "Мужские", "Унисекс"])
        self.cmb_gender.setCurrentText(self.settings.get("gender_mode", "Auto"))
        ml.addWidget(self.cmb_gender, 5, 1, 1, 2)

        ml.addWidget(QLabel("Бренд в названии"), 6, 0)
        self.cmb_brand_title = QComboBox()
        self.cmb_brand_title.addItems(["50/50", "Всегда", "Никогда"])
        self.cmb_brand_title.setCurrentText(self.settings.get("brand_title_ui", "50/50"))
        ml.addWidget(self.cmb_brand_title, 6, 1, 1, 2)

        # ВСЕГДА 6 строк (показывать можно, но UI не меняем визуально сильно)
        self.spin_rows = QSpinBox()
        self.spin_rows.setMinimum(6)
        self.spin_rows.setMaximum(6)
        self.spin_rows.setValue(6)
        ml.addWidget(QLabel("Строк заполнять"), 7, 0)
        ml.addWidget(self.spin_rows, 7, 1, 1, 2)

        self.spin_uniq = QSpinBox()
        self.spin_uniq.setMinimum(60)
        self.spin_uniq.setMaximum(95)
        self.spin_uniq.setValue(int(self.settings.get("uniq_strength", 90)))
        ml.addWidget(QLabel("Уникализация"), 8, 0)
        ml.addWidget(self.spin_uniq, 8, 1, 1, 2)

        root.addWidget(main)

        bottom = card()
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(16, 14, 16, 14)
        bl.setSpacing(12)

        self.progress = QProgressBar()
        self.progress.setValue(0)

        self.btn_run = QPushButton("🚀 ГОТОВО")
        self.btn_run.setObjectName("Primary")
        self.btn_run.clicked.connect(self.start)

        bl.addWidget(self.progress, 2)
        bl.addWidget(self.btn_run, 1)
        root.addWidget(bottom)

        self.apply_theme(self.cmb_theme.currentText())

    def apply_theme(self, name: str):
        QApplication.instance().setStyleSheet(THEMES.get(name, ""))
        self.settings["theme"] = name
        save_settings(self.settings)

    def open_data_folder(self):
        try:
            os.startfile(str(self.data_dir))
        except Exception:
            QMessageBox.information(self, "Data", str(self.data_dir))

    def pick_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX", "", "Excel (*.xlsx)")
        if path:
            self.xlsx_path = path
            self.lbl_xlsx.setText(Path(path).name)

    def reload_lists(self):
        brands = load_list(self.brands_file, ["Ray-Ban"])
        shapes = load_list(self.shapes_file, ["квадратные"])
        lenses = load_list(self.lenses_file, ["UV400"])

        self.cmb_brand.clear(); self.cmb_shape.clear(); self.cmb_lens.clear()
        self.cmb_brand.addItems(brands)
        self.cmb_shape.addItems(shapes)
        self.cmb_lens.addItems(lenses)

        # restore
        if self.settings.get("brand") in brands:
            self.cmb_brand.setCurrentText(self.settings["brand"])
        if self.settings.get("shape") in shapes:
            self.cmb_shape.setCurrentText(self.settings["shape"])
        if self.settings.get("lens") in lenses:
            self.cmb_lens.setCurrentText(self.settings["lens"])

    def _brand_title_mode(self) -> str:
        t = self.cmb_brand_title.currentText()
        if "Всегда" in t:
            return "always"
        if "Никогда" in t:
            return "never"
        return "smart50"

    def _autosave_user_values(self):
        # сохраняем то, что ты ввёл вручную в combo (без кнопки +)
        b = self.cmb_brand.currentText().strip()
        s = self.cmb_shape.currentText().strip()
        l = self.cmb_lens.currentText().strip()
        if b: add_to_list(self.brands_file, b)
        if s: add_to_list(self.shapes_file, s)
        if l: add_to_list(self.lenses_file, l)
        self.reload_lists()

    def preview(self):
        try:
            b = self.cmb_brand.currentText().strip()
            s = self.cmb_shape.currentText().strip()
            l = self.cmb_lens.currentText().strip()
            c = self.ed_collection.text().strip()
            items = generate_preview(
                brand_lat=b, shape=s, lens=l, collection=c,
                seo_level=self.cmb_seo.currentText(),
                gender_mode=self.cmb_gender.currentText(),
                uniq_strength=int(self.spin_uniq.value()),
                brand_in_title_mode=self._brand_title_mode(),
                data_dir=str(self.data_dir),
                count=3,
            )
            PreviewDialog(self, items).exec_()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка preview", str(e))

    def start(self):
        if not self.xlsx_path:
            QMessageBox.warning(self, "Файл", "Сначала выбери XLSX файл.")
            return

        # сохраняем ввод пользователя
        self._autosave_user_values()

        b = self.cmb_brand.currentText().strip()
        s = self.cmb_shape.currentText().strip()
        l = self.cmb_lens.currentText().strip()
        c = self.ed_collection.text().strip()

        self.settings.update({
            "brand": b, "shape": s, "lens": l, "collection": c,
            "seo_level": self.cmb_seo.currentText(),
            "gender_mode": self.cmb_gender.currentText(),
            "brand_title_ui": self.cmb_brand_title.currentText(),
            "uniq_strength": int(self.spin_uniq.value()),
        })
        save_settings(self.settings)

        args = dict(
            input_xlsx=self.xlsx_path,
            brand_lat=b,
            shape=s,
            lens=l,
            collection=c,
            seo_level=self.cmb_seo.currentText(),
            gender_mode=self.cmb_gender.currentText(),
            uniq_strength=int(self.spin_uniq.value()),
            brand_in_title_mode=self._brand_title_mode(),
            data_dir=str(self.data_dir),
            max_fill_rows=6,
            output_index=1,
            output_total=1,
            between_files_slogan_lock=True,
        )

        self.progress.setValue(0)
        self.btn_run.setEnabled(False)

        self.worker = Worker(args)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.done)
        self.worker.failed.connect(self.fail)
        self.worker.start()

    def done(self, out_path: str):
        self.btn_run.setEnabled(True)
        self.progress.setValue(100)
        QMessageBox.information(self, "Готово", f"Сохранено:\n{out_path}")

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
