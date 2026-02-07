# main.py
from __future__ import annotations

import os
import sys
import traceback
from dataclasses import asdict
from typing import Optional

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFrame, QGridLayout,
    QHBoxLayout, QLabel, QLineEdit, QMainWindow, QMessageBox, QPushButton,
    QProgressBar, QSpinBox, QVBoxLayout, QWidget, QInputDialog
)

from wb_fill import FillParams, fill_wb_template, generate_preview


def app_data_dir(app_name: str = "Sunglasses SEO PRO") -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, app_name, "data")
    os.makedirs(path, exist_ok=True)
    return path


GRAPHITE_QSS = """
QMainWindow { background: #0f0f10; }
QWidget { color: #e8e8e8; font-size: 12px; }
QFrame#card { background: #151516; border: 1px solid #2b2b2c; border-radius: 14px; }
QLabel#title { font-size: 22px; font-weight: 700; }
QLabel#subtitle { color: #b8b8b8; }
QPushButton {
  background: #2d6cdf; border: none; padding: 10px 14px; border-radius: 12px; font-weight: 700;
}
QPushButton:hover { background: #3776ea; }
QPushButton:disabled { background: #2b2b2c; color: #888; }

QPushButton#btnSmall {
  padding: 8px 12px; border-radius: 12px; font-weight: 700;
}

QLineEdit, QComboBox {
  background: #101011; border: 1px solid #2b2b2c; border-radius: 10px; padding: 8px 10px;
}
QComboBox::drop-down { border: none; width: 24px; }
QComboBox::down-arrow { image: none; }

QCheckBox { spacing: 8px; }
QProgressBar {
  background: #101011; border: 1px solid #2b2b2c; border-radius: 12px; text-align: center;
  height: 22px;
}
QProgressBar::chunk { background: #2d6cdf; border-radius: 12px; }
"""


class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, in_path: str, out_dir: str, base_name: str, params: FillParams, batch_n: int = 1):
        super().__init__()
        self.in_path = in_path
        self.out_dir = out_dir
        self.base_name = base_name
        self.params = params
        self.batch_n = max(1, int(batch_n))

    def run(self):
        try:
            results = {"files": []}
            for i in range(1, self.batch_n + 1):
                if self.batch_n == 1:
                    out_name = f"{self.base_name}_ready.xlsx"
                else:
                    out_name = f"{self.base_name}_ready_{i}.xlsx"
                out_path = os.path.join(self.out_dir, out_name)

                # seed, чтобы реально отличались тексты между файлами
                p = self.params
                p.seed = (p.seed or 0) + i * 10007

                rep = fill_wb_template(self.in_path, out_path, p)
                results["files"].append(rep)

                pct = int(i / self.batch_n * 100)
                self.progress.emit(pct)

            self.finished.emit(results)
        except Exception as e:
            self.failed.emit(f"{e}\n\n{traceback.format_exc()}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sunglasses SEO PRO")
        self.setFixedSize(1104, 738)

        # DPI
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

        self.data_dir = app_data_dir("Sunglasses SEO PRO")
        self.in_xlsx: Optional[str] = None

        root = QWidget()
        self.setCentralWidget(root)
        root.setStyleSheet(GRAPHITE_QSS)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        # Header card
        header = QFrame()
        header.setObjectName("card")
        hl = QVBoxLayout(header)
        hl.setContentsMargins(18, 14, 18, 14)
        title = QLabel("Sunglasses SEO PRO")
        title.setObjectName("title")
        subtitle = QLabel("Живые SEO-описания • Выпадающие списки • Прогресс • Темы • WB Safe/Strict • AUTO-пол")
        subtitle.setObjectName("subtitle")
        hl.addWidget(title)
        hl.addWidget(subtitle)
        layout.addWidget(header)

        # Controls card (theme + data dir)
        ctrl = QFrame()
        ctrl.setObjectName("card")
        cl = QGridLayout(ctrl)
        cl.setContentsMargins(18, 14, 18, 14)
        cl.setHorizontalSpacing(12)
        cl.setVerticalSpacing(10)

        cl.addWidget(QLabel("🎨 Тема"), 0, 0)
        self.theme_cb = QComboBox()
        self.theme_cb.addItems(["Graphite"])
        cl.addWidget(self.theme_cb, 0, 1)

        cl.addWidget(QLabel("📁 Справочники:"), 0, 2)
        self.data_dir_le = QLineEdit(self.data_dir)
        self.data_dir_le.setReadOnly(True)
        cl.addWidget(self.data_dir_le, 0, 3)

        btn_folder = QPushButton("Папка")
        btn_folder.setObjectName("btnSmall")
        btn_folder.clicked.connect(self.pick_data_dir)
        cl.addWidget(btn_folder, 0, 4)

        layout.addWidget(ctrl)

        # XLSX card
        xcard = QFrame()
        xcard.setObjectName("card")
        xl = QHBoxLayout(xcard)
        xl.setContentsMargins(18, 14, 18, 14)
        btn_load = QPushButton("📄 Загрузить XLSX")
        btn_load.clicked.connect(self.pick_xlsx)
        self.x_label = QLabel("Файл не выбран")
        xl.addWidget(btn_load)
        xl.addWidget(self.x_label, 1)
        layout.addWidget(xcard)

        # Form card
        form = QFrame()
        form.setObjectName("card")
        fl = QGridLayout(form)
        fl.setContentsMargins(18, 14, 18, 14)
        fl.setHorizontalSpacing(12)
        fl.setVerticalSpacing(10)

        self.brand_cb = QComboBox(); self.brand_cb.addItems(["Balenciaga", "Gucci", "Prada", "Ray-Ban", "Dior", "Versace"])
        self.shape_cb = QComboBox(); self.shape_cb.addItems(["Вайфаеры", "Авиаторы", "Кошачий глаз", "Квадратные", "Круглые", "Овальные"])
        self.lens_cb = QComboBox(); self.lens_cb.addItems(["Поляризационные", "Градиентные", "Зеркальные", "Фотохромные", "УФ400"])
        self.collection_le = QLineEdit("Весна–Лето 2026")

        # left labels + combos
        fl.addWidget(QLabel("Бренд"), 0, 0); fl.addWidget(self.brand_cb, 0, 1)
        fl.addWidget(QLabel("Форма оправы"), 1, 0); fl.addWidget(self.shape_cb, 1, 1)
        fl.addWidget(QLabel("Линзы"), 2, 0); fl.addWidget(self.lens_cb, 2, 1)
        fl.addWidget(QLabel("Коллекция"), 3, 0); fl.addWidget(self.collection_le, 3, 1)

        # placeholders for "+" buttons to match screenshot look (без логики, чтобы не ломать UI)
        for r in range(0, 3):
            plus = QPushButton("+")
            plus.setFixedWidth(42)
            plus.clicked.connect(lambda: None)
            fl.addWidget(plus, r, 2)

        # seo density / length / style row
        fl.addWidget(QLabel("SEO-плотность"), 4, 0)
        self.seo_cb = QComboBox(); self.seo_cb.addItems(["low", "normal", "high"])
        fl.addWidget(self.seo_cb, 4, 1)

        fl.addWidget(QLabel("Длина"), 4, 2)
        self.len_cb = QComboBox(); self.len_cb.addItems(["short", "medium", "long"])
        fl.addWidget(self.len_cb, 4, 3)

        fl.addWidget(QLabel("Стиль"), 4, 4)
        self.style_cb = QComboBox(); self.style_cb.addItems(["premium", "basic", "sport"])
        fl.addWidget(self.style_cb, 4, 5)

        # gender + safe/strict
        fl.addWidget(QLabel("AUTO-пол"), 5, 0)
        self.gender_cb = QComboBox(); self.gender_cb.addItems(["Auto", "Male", "Female"])
        fl.addWidget(self.gender_cb, 5, 1)

        self.safe_ck = QCheckBox("WB Safe Mode (заменяет риск-слова)")
        self.safe_ck.setChecked(True)
        self.strict_ck = QCheckBox("WB Strict (убирает обещания/абсолюты)")
        self.strict_ck.setChecked(True)
        fl.addWidget(self.safe_ck, 5, 2, 1, 2)
        fl.addWidget(self.strict_ck, 5, 4, 1, 2)

        layout.addWidget(form)

        # Bottom bar card (progress + buttons)
        bottom = QFrame()
        bottom.setObjectName("card")
        bl = QHBoxLayout(bottom)
        bl.setContentsMargins(18, 14, 18, 14)
        bl.setSpacing(12)

        self.progress = QProgressBar()
        self.progress.setValue(0)

        self.btn_single = QPushButton("🚀 СГЕНЕРИРОВАТЬ")
        self.btn_single.clicked.connect(self.run_single)

        self.btn_batch = QPushButton("📦 СОЗДАТЬ ПАКЕТ XLSX")
        self.btn_batch.clicked.connect(self.run_batch)

        bl.addWidget(self.progress, 1)
        bl.addWidget(self.btn_batch)
        bl.addWidget(self.btn_single)

        layout.addWidget(bottom)

        # Live preview (не отображаем отдельным блоком, чтобы UI оставался как на скрине)
        # но generate_preview доступен — можно включить позже без риска импорта.

        self.worker: Optional[Worker] = None

    def pick_data_dir(self):
        path = QFileDialog.getExistingDirectory(self, "Выберите папку со справочниками", self.data_dir)
        if path:
            self.data_dir = path
            self.data_dir_le.setText(path)

    def pick_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX", "", "Excel (*.xlsx)")
        if path:
            self.in_xlsx = path
            self.x_label.setText(os.path.basename(path))

    def _build_params(self, seed: Optional[int] = None) -> FillParams:
        gender = self.gender_cb.currentText().lower()
        if gender == "auto":
            gm = "auto"
        elif gender == "male":
            gm = "male"
        else:
            gm = "female"

        return FillParams(
            brand=self.brand_cb.currentText(),
            shape=self.shape_cb.currentText(),
            lens=self.lens_cb.currentText(),
            collection=self.collection_le.text().strip() or "Коллекция",
            seo_density=self.seo_cb.currentText(),
            length_mode=self.len_cb.currentText(),
            style_mode=self.style_cb.currentText(),
            gender_mode=gm,
            safe_mode=self.safe_ck.isChecked(),
            strict_mode=self.strict_ck.isChecked(),
            data_dir=self.data_dir,
            seed=seed,
            rows_to_fill=6,          # <<< ВАЖНО: 6 строк
            fill_only_empty=True,
            uniq_strength=3,
        )

    def _ensure_ready(self) -> bool:
        if not self.in_xlsx or not os.path.exists(self.in_xlsx):
            QMessageBox.warning(self, "Нет файла", "Сначала нажми «Загрузить XLSX» и выбери файл.")
            return False
        return True

    def _run(self, batch_n: int):
        if not self._ensure_ready():
            return

        in_path = self.in_xlsx
        out_dir = os.path.dirname(in_path)
        base = os.path.splitext(os.path.basename(in_path))[0]

        params = self._build_params(seed=12345)

        self.progress.setValue(0)
        self.btn_single.setEnabled(False)
        self.btn_batch.setEnabled(False)

        self.worker = Worker(in_path, out_dir, base, params, batch_n=batch_n)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.on_done)
        self.worker.failed.connect(self.on_fail)
        self.worker.start()

    def run_single(self):
        self._run(batch_n=1)

    def run_batch(self):
        if not self._ensure_ready():
            return
        n, ok = QInputDialog.getInt(self, "Пакет XLSX", "Сколько XLSX создать разом?", 5, 2, 50, 1)
        if ok:
            self._run(batch_n=n)

    def on_done(self, result: dict):
        self.btn_single.setEnabled(True)
        self.btn_batch.setEnabled(True)
        self.progress.setValue(100)

        files = result.get("files") or []
        last = files[-1]["out_path"] if files else ""
        QMessageBox.information(self, "Готово", f"Создано файлов: {len(files)}\nПоследний файл:\n{last}")

    def on_fail(self, err: str):
        self.btn_single.setEnabled(True)
        self.btn_batch.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", err)


def main():
    app = QApplication(sys.argv)
    # аккуратный базовый шрифт
    app.setFont(QFont("Segoe UI", 10))
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
