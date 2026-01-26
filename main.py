import sys

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QMessageBox
)

from seo_app.wb_fill import fill_wb_template


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sunglasses SEO (WB)")

        self.input_path = ""

        layout = QVBoxLayout()

        row_file = QHBoxLayout()
        self.btn_pick = QPushButton("Загрузить XLSX (WB шаблон)")
        self.btn_pick.clicked.connect(self.pick_file)
        self.lbl_file = QLabel("Файл не выбран")
        self.lbl_file.setWordWrap(True)
        row_file.addWidget(self.btn_pick)
        row_file.addWidget(self.lbl_file)
        layout.addLayout(row_file)

        self.ed_brand = QLineEdit()
        self.ed_brand.setPlaceholderText("Бренд (например: PRADA)")

        self.ed_shape = QLineEdit()
        self.ed_shape.setPlaceholderText("Форма оправы (например: авиаторы)")

        self.ed_lens = QLineEdit()
        self.ed_lens.setPlaceholderText("Линзы (например: UV400, поляризационные)")

        self.ed_collection = QLineEdit()
        self.ed_collection.setPlaceholderText("Коллекция (например: Весна–Лето 2025–2026)")
        self.ed_collection.setText("Весна–Лето 2025–2026")

        layout.addWidget(QLabel("Параметры (одинаковые для всего файла):"))
        layout.addWidget(self.ed_brand)
        layout.addWidget(self.ed_shape)
        layout.addWidget(self.ed_lens)
        layout.addWidget(self.ed_collection)

        self.btn_run = QPushButton("Готово (заполнить Наименование и Описание)")
        self.btn_run.clicked.connect(self.run_fill)
        layout.addWidget(self.btn_run)

        self.lbl_status = QLabel("")
        self.lbl_status.setWordWrap(True)
        layout.addWidget(self.lbl_status)

        self.setLayout(layout)
        self.resize(780, 260)

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите XLSX файл WB", "", "Excel files (*.xlsx)")
        if path:
            self.input_path = path
            self.lbl_file.setText(path)
            self.lbl_status.setText("")

    def run_fill(self):
        if not self.input_path:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите XLSX файл.")
            return

        brand = self.ed_brand.text().strip()
        shape = self.ed_shape.text().strip()
        lens = self.ed_lens.text().strip()
        collection = self.ed_collection.text().strip()

        if not brand or not shape or not lens:
            QMessageBox.warning(self, "Ошибка", "Заполните: Бренд, Форма оправы, Линзы.")
            return

        try:
            self.btn_run.setEnabled(False)
            self.lbl_status.setText("Работаю...")

            out_path, rows_done = fill_wb_template(
                input_xlsx=self.input_path,
                brand=brand,
                shape=shape,
                lens_features=lens,
                collection=collection
            )

            msg = f"Готово. Заполнено строк: {rows_done}\nФайл сохранён: {out_path}"
            self.lbl_status.setText(msg)
            QMessageBox.information(self, "Готово", msg)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", str(e))
            self.lbl_status.setText(f"Ошибка: {e}")
        finally:
            self.btn_run.setEnabled(True)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
