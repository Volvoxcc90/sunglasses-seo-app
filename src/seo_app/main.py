import sys
import os
import json
from pathlib import Path

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QStandardPaths
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QFileDialog,
    QVBoxLayout, QHBoxLayout, QComboBox, QMessageBox, QProgressBar,
    QCheckBox, QFrame, QSizePolicy
)

from seo_app.wb_fill import fill_wb_template


# ==========================
# User data path (AppData/Roaming) — важно для EXE
# ==========================
APP_NAME = "Sunglasses SEO PRO"


def get_appdata_dir() -> Path:
    base = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation)
    # Обычно это .../AppData/Roaming/<AppName>
    p = Path(base)
    p.mkdir(parents=True, exist_ok=True)
    return p


APP_DIR = get_appdata_dir()
DATA_DIR = APP_DIR / "data"
SETTINGS_FILE = DATA_DIR / "ui_settings.json"

BRANDS_FILE = DATA_DIR / "brands.txt"
SHAPES_FILE = DATA_DIR / "shapes.txt"
LENSES_FILE = DATA_DIR / "lenses.txt"

DATA_DIR.mkdir(parents=True, exist_ok=True)


# ==========================
# Helpers
# ==========================
def load_list(path: Path, defaults):
    if not path.exists():
        path.write_text("\n".join(defaults), encoding="utf-8")
        return list(defaults)
    items = [x.strip() for x in path.read_text(encoding="utf-8").splitlines() if x.strip()]
    if not items:
        items = list(defaults)
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
    try:
        SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        # Это должно работать в AppData, но на всякий случай
        print("save_settings error:", e)


def open_folder(path: Path):
    try:
        os.startfile(str(path))
    except Exception:
        QMessageBox.warning(None, "Ошибка", f"Не удалось открыть папку:\n{path}")


# ==========================
# Themes
# ==========================
THEMES = {
    "Sepia": """
        QWidget { background: #f4f1ea; color: #1f1f1f; font-size: 13px; }
        QFrame#Card { background: rgba(255,255,255,0.75); border: 1px solid rgba(0,0,0,0.08); border-radius: 14px; }
        QLabel#Title { font-size: 22px; font-weight: 700; }
        QLabel#Subtitle { color: rgba(0,0,0,0.55); }
        QComboBox, QPushButton, QProgressBar, QCheckBox { border-radius: 10px; }
        QComboBox { padding: 8px 10px; border: 1px solid rgba(0,0,0,0.12); background: rgba(255,255,255,0.85); }
        QComboBox:focus { border: 1px solid rgba(124,58,237,0.55); }
        QPushButton { padding: 10px 14px; border: 1px solid rgba(0,0,0,0.10); background: rgba(255,255,255,0.85); }
        QPushButton:hover { background: rgba(255,255,255,1.0); }
        QPushButton#Primary {
            border: 0; color: white; font-weight: 700; padding: 12px 16px;
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6d28d9, stop:1 #9333ea);
        }
        QPushButton#Primary:hover {
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #5b21b6, stop:1 #7e22ce);
        }
        QPushButton#Tiny { padding: 8px 10px; font-weight: 700; min-width: 40px; }
        QProgressBar { height: 18px; border: 1px solid rgba(0,0,0,0.10); background: rgba(255,255,255,0.65); text-align: center; }
        QProgressBar::chunk { border-radius: 8px; background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6d28d9, stop:1 #9333ea); }
    """,
    "Midnight": """
        QWidget { background: #0b1020; color: #e9ecf1; font-size: 13px; }
        QFrame#Card { background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.10); border-radius: 14px; }
        QLabel#Title { font-size: 22px; font-weight: 700; }
        QLabel#Subtitle { color: rgba(233,236,241,0.60); }
        QComboBox, QPushButton, QProgressBar, QCheckBox { border-radius: 10px; }
        QComboBox { padding: 8px 10px; border: 1px solid rgba(255,255,255,0.16); background: rgba(255,255,255,0.08); }
        QComboBox:focus { border: 1px solid rgba(124,58,237,0.70); }
        QPushButton { padding: 10px 14px; border: 1px solid rgba(255,255,255,0.16); background: rgba(255,255,255,0.08); }
        QPushButton:hover { background: rgba(255,255,255,0.12); }
        QPushButton#Primary {
            border: 0; color: white; font-weight: 700; padding: 12px 16px;
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6d28d9, stop:1 #9333ea);
        }
        QPushButton#Tiny { padding: 8px 10px; font-weight: 700; min-width: 40px; }
        QProgressBar { height: 18px; border: 1px solid rgba(255,255,255,0.16); background: rgba(255,255,255,0.06); text-align: center; }
        QProgressBar::chunk { border-radius: 8px; background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #6d28d9, stop:1 #9333ea); }
    """,
    "Slate": """
        QWidget { background: #eef2f7; color: #0f172a; font-size: 13px; }
        QFrame#Card { background: #ffffff; border: 1px solid rgba(15,23,42,0.08); border-radius: 14px; }
        QLabel#Title { font-size: 22px; font-weight: 700; }
        QLabel#Subtitle { color: rgba(15,23,42,0.55); }
        QComboBox, QPushButton, QProgressBar, QCheckBox { border-radius: 10px; }
        QComboBox { padding: 8px 10px; border: 1px solid rgba(15,23,42,0.14); background: #ffffff; }
        QComboBox:focus { border: 1px solid rgba(124,58,237,0.55); }
        QPushButton { padding: 10px 14px; border: 1px solid rgba(15,23,42,0.14); background: #ffffff; }
        QPushButton#Primary {
            border:0; color:#fff; font-weight:700; padding:12px 16px;
            background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #1d4ed8, stop:1 #2563eb);
        }
        QPushButton#Tiny { padding: 8px 10px; font-weight: 700; min-width: 40px; }
        QProgressBar { height: 18px; border: 1px solid rgba(15,23,42,0.12); background: #ffffff; text-align:center; }
        QProgressBar::chunk { border-radius: 8px; background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #1d4ed8, stop:1 #2563eb); }
    """,
}


# ==========================
# Worker Thread
# ==========================
class FillWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str, int, str)
    failed = pyqtSignal(str)

    def __init__(
        self,
        *,
        input_xlsx,
        brand,
        shape,
        lens,
        collection,
        style,
        seo_level,
        desc_length,
        wb_safe_mode,
        wb_strict,
        gender_mode
    ):
        super().__init__()
        self.args = dict(
            input_xlsx=input_xlsx,
            brand=brand,
            shape=shape,
            lens_features=lens,
            collection=collection,
            style=style,
            seo_level=seo_level,
            desc_length=desc_length,
            wb_safe_mode=wb_safe_mode,
            wb_strict=wb_strict,
            gender_mode=gender_mode,
        )

    def run(self):
        try:
            import inspect

            sig = inspect.signature(fill_wb_template)
            allowed = sig.parameters.keys()

            safe_args = {k: v for k, v in self.args.items() if k in allowed}

            # прогресс — только если функция реально поддерживает
            if "progress_callback" in allowed:
                safe_args["progress_callback"] = lambda p: self.progress.emit(int(p))

            out, count, report_json = fill_wb_template(**safe_args)
            self.finished.emit(out, count, report_json)

        except Exception as e:
            self.failed.emit(str(e))


# ==========================
# UI
# ==========================
class SeoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumWidth(840)

        self.settings = load_settings()

        self.brands = load_list(BRANDS_FILE, ["Cazal", "Gucci", "Prada", "Ray-Ban"])
        self.shapes = load_list(SHAPES_FILE, ["Квадратные", "Овальные", "Кошачий глаз"])
        self.lenses = load_list(LENSES_FILE, ["UV400", "Поляризационные", "Фотохромные"])

        self.worker = None
        self.selected_file = ""

        self.build_ui()

        # ВАЖНО: сначала заполнить, потом восстановить значения
        self.restore_settings()
        self.apply_theme(self.theme_cb.currentText())

    def build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        # Header
        header = QFrame()
        header.setObjectName("Card")
        header_l = QVBoxLayout(header)
        header_l.setContentsMargins(18, 16, 18, 16)
        header_l.setSpacing(6)

        title_row = QHBoxLayout()
        title_row.setSpacing(10)
        icon = QLabel("🕶️")
        icon.setFixedWidth(28)
        icon.setAlignment(Qt.AlignVCenter)
        title_row.addWidget(icon)

        title = QLabel(APP_NAME)
        title.setObjectName("Title")
        title_row.addWidget(title, 1)
        header_l.addLayout(title_row)

        subtitle = QLabel("Живые SEO-описания • Выпадающие списки • Прогресс • Темы • WB Safe/Strict • AUTO-пол")
        subtitle.setObjectName("Subtitle")
        header_l.addWidget(subtitle)

        root.addWidget(header)

        # Top card
        top = QFrame()
        top.setObjectName("Card")
        top_l = QVBoxLayout(top)
        top_l.setContentsMargins(18, 16, 18, 16)
        top_l.setSpacing(10)

        theme_row = QHBoxLayout()
        theme_row.setSpacing(10)
        theme_row.addWidget(QLabel("🎨 Тема"))
        self.theme_cb = QComboBox()
        self.theme_cb.addItems(list(THEMES.keys()))
        self.theme_cb.currentTextChanged.connect(self.on_theme_changed)
        self.theme_cb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        theme_row.addWidget(self.theme_cb, 1)

        theme_row.addWidget(QLabel("📁 Справочники:"))
        self.data_path_lbl = QLabel(str(DATA_DIR))
        self.data_path_lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
        theme_row.addWidget(self.data_path_lbl, 2)

        btn_open_data = QPushButton("Папка")
        btn_open_data.setObjectName("Tiny")
        btn_open_data.clicked.connect(lambda: open_folder(DATA_DIR))
        theme_row.addWidget(btn_open_data)

        top_l.addLayout(theme_row)

        file_row = QHBoxLayout()
        file_row.setSpacing(10)

        self.btn_pick = QPushButton("📄 Загрузить XLSX")
        self.btn_pick.clicked.connect(self.pick_file)
        file_row.addWidget(self.btn_pick)

        self.file_lbl = QLabel("Файл не выбран")
        self.file_lbl.setStyleSheet("opacity: 0.85;")
        file_row.addWidget(self.file_lbl, 1)

        top_l.addLayout(file_row)
        root.addWidget(top)

        # Form card
        form = QFrame()
        form.setObjectName("Card")
        form_l = QVBoxLayout(form)
        form_l.setContentsMargins(18, 16, 18, 16)
        form_l.setSpacing(10)

        form_l.addLayout(self.combo_row("Бренд", self.brands, BRANDS_FILE, attr_name="brand_cb"))
        form_l.addLayout(self.combo_row("Форма оправы", self.shapes, SHAPES_FILE, attr_name="shape_cb"))
        form_l.addLayout(self.combo_row("Линзы", self.lenses, LENSES_FILE, attr_name="lens_cb"))

        row = QHBoxLayout()
        row.setSpacing(10)
        row.addWidget(QLabel("Коллекция"))
        self.collection_cb = QComboBox()
        self.collection_cb.addItems(["Весна–Лето 2025–2026", "Весна–Лето 2026", "Весна–Лето 2025"])
        row.addWidget(self.collection_cb, 1)
        form_l.addLayout(row)

        grid1 = QHBoxLayout()
        grid1.setSpacing(10)

        grid1.addWidget(QLabel("SEO-плотность"))
        self.seo_cb = QComboBox()
        self.seo_cb.addItems(["soft", "normal", "hard"])
        grid1.addWidget(self.seo_cb, 1)

        grid1.addWidget(QLabel("Длина"))
        self.len_cb = QComboBox()
        self.len_cb.addItems(["short", "medium", "long"])
        grid1.addWidget(self.len_cb, 1)

        grid1.addWidget(QLabel("Стиль"))
        self.style_cb = QComboBox()
        self.style_cb.addItems(["neutral", "premium", "social"])
        grid1.addWidget(self.style_cb, 1)

        form_l.addLayout(grid1)

        grid2 = QHBoxLayout()
        grid2.setSpacing(10)

        grid2.addWidget(QLabel("AUTO-пол"))
        self.gender_cb = QComboBox()
        self.gender_cb.addItems(["Auto", "Жен", "Муж", "Унисекс"])
        grid2.addWidget(self.gender_cb, 1)

        self.safe_cb = QCheckBox("WB Safe Mode (заменяет риск-слова)")
        grid2.addWidget(self.safe_cb, 2)

        self.strict_cb = QCheckBox("WB Strict (убирает обещания/абсолюты/стоп-фразы)")
        grid2.addWidget(self.strict_cb, 2)

        form_l.addLayout(grid2)

        root.addWidget(form)

        # Bottom
        bottom = QFrame()
        bottom.setObjectName("Card")
        bottom_l = QHBoxLayout(bottom)
        bottom_l.setContentsMargins(18, 14, 18, 14)
        bottom_l.setSpacing(12)

        self.progress = QProgressBar()
        self.progress.setValue(0)
        bottom_l.addWidget(self.progress, 3)

        self.btn_run = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        self.btn_run.setObjectName("Primary")
        self.btn_run.clicked.connect(self.run_generation)
        bottom_l.addWidget(self.btn_run, 2)

        root.addWidget(bottom)

    def combo_row(self, label_text, items, file_path: Path, attr_name: str):
        row = QHBoxLayout()
        row.setSpacing(10)

        row.addWidget(QLabel(label_text))
        cb = QComboBox()
        cb.setEditable(True)
        cb.addItems(items)
        cb.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        row.addWidget(cb, 1)

        btn = QPushButton("+")
        btn.setObjectName("Tiny")

        def add_item():
            val = cb.currentText().strip()
            if not val:
                QMessageBox.warning(self, "Пусто", "Введи значение и нажми +")
                return
            added = add_to_list(file_path, val)
            cb.clear()
            cb.addItems(load_list(file_path, []))
            cb.setCurrentText(val)
            QMessageBox.information(self, "Ок", "Добавлено" if added else "Уже было в списке")

        btn.clicked.connect(add_item)
        row.addWidget(btn)

        setattr(self, attr_name, cb)
        return row

    # ---- settings
    def capture_settings(self) -> dict:
        return {
            "theme": self.theme_cb.currentText(),
            "brand": self.brand_cb.currentText().strip(),
            "shape": self.shape_cb.currentText().strip(),
            "lens": self.lens_cb.currentText().strip(),
            "collection": self.collection_cb.currentText(),
            "seo_level": self.seo_cb.currentText(),
            "desc_length": self.len_cb.currentText(),
            "style": self.style_cb.currentText(),
            "wb_safe_mode": self.safe_cb.isChecked(),
            "wb_strict": self.strict_cb.isChecked(),
            "gender_mode": self.gender_cb.currentText(),
        }

    def restore_settings(self):
        s = self.settings or {}

        theme = s.get("theme", "Sepia")
        if theme in THEMES:
            self.theme_cb.setCurrentText(theme)
        else:
            self.theme_cb.setCurrentText("Sepia")

        # editable=True -> setCurrentText ок даже если значения нет
        self.brand_cb.setCurrentText(s.get("brand", ""))
        self.shape_cb.setCurrentText(s.get("shape", ""))
        self.lens_cb.setCurrentText(s.get("lens", ""))

        self.collection_cb.setCurrentText(s.get("collection", "Весна–Лето 2025–2026"))
        self.seo_cb.setCurrentText(s.get("seo_level", "normal"))
        self.len_cb.setCurrentText(s.get("desc_length", "medium"))
        self.style_cb.setCurrentText(s.get("style", "neutral"))

        self.safe_cb.setChecked(bool(s.get("wb_safe_mode", True)))
        self.strict_cb.setChecked(bool(s.get("wb_strict", True)))
        self.gender_cb.setCurrentText(s.get("gender_mode", "Auto"))

    def on_theme_changed(self, name: str):
        self.apply_theme(name)
        self.settings = self.capture_settings()
        save_settings(self.settings)

    def apply_theme(self, name: str):
        qss = THEMES.get(name, THEMES["Sepia"])
        self.setStyleSheet(qss)

    def pick_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выбери Excel", "", "Excel (*.xlsx)")
        if not file_path:
            return
        self.selected_file = file_path
        self.file_lbl.setText(file_path)

    def set_busy(self, busy: bool):
        self.btn_run.setEnabled(not busy)
        self.btn_pick.setEnabled(not busy)

    def run_generation(self):
        if not self.selected_file:
            QMessageBox.warning(self, "Файл", "Сначала выбери XLSX.")
            return

        # сохраняем настройки сразу
        self.settings = self.capture_settings()
        save_settings(self.settings)

        brand = self.settings["brand"]
        shape = self.settings["shape"]
        lens = self.settings["lens"]
        collection = self.settings["collection"]
        seo = self.settings["seo_level"]
        length = self.settings["desc_length"]
        style = self.settings["style"]
        safe = self.settings["wb_safe_mode"]
        strict = self.settings["wb_strict"]
        gender_mode = self.settings["gender_mode"]

        self.progress.setValue(0)
        self.set_busy(True)

        self.worker = FillWorker(
            input_xlsx=self.selected_file,
            brand=brand,
            shape=shape,
            lens=lens,
            collection=collection,
            style=style,
            seo_level=seo,
            desc_length=length,
            wb_safe_mode=safe,
            wb_strict=strict,
            gender_mode=gender_mode,
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_done)
        self.worker.failed.connect(self.on_fail)
        self.worker.start()

    def on_done(self, out_path: str, count: int, report_json: str):
        self.set_busy(False)
        QMessageBox.information(
            self,
            "Готово",
            f"Создан файл:\n{out_path}\nКарточек: {count}\n\nОтчёт:\n{report_json}\n(рядом будет .seo_report.txt)"
        )

    def on_fail(self, err: str):
        self.set_busy(False)
        QMessageBox.critical(self, "Ошибка", err)

    def closeEvent(self, event):
        # сохраняем при закрытии тоже
        try:
            self.settings = self.capture_settings()
            save_settings(self.settings)
        except Exception:
            pass
        super().closeEvent(event)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Чтобы AppData путь был стабилен и “красивый”
    app.setApplicationName(APP_NAME)
    w = SeoApp()
    w.show()
    sys.exit(app.exec())
