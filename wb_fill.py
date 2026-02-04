# wb_fill.py
import random
import re
from pathlib import Path
from openpyxl import load_workbook


TITLE_MAX = 60
DESC_MAX = 2000


SLOGANS = [
    "Красивые", "Крутые", "Стильные", "Модные", "Молодёжные",
    "Дизайнерские", "Эффектные", "Трендовые", "Лаконичные",
    "Яркие", "Современные", "Премиальные", "Универсальные",
    "Актуальные", "Выразительные", "Элегантные", "Минималистичные",
    "Смелые", "Классные", "Городские", "Лёгкие", "Комфортные",
    "Популярные", "Эксклюзивные", "Фирменные"
]

SUN_TERMS = ["солнцезащитные очки", "солнечные очки"]


INTRO = [
    "Эта модель легко вписывается в повседневный гардероб и подчёркивает индивидуальность.",
    "Аксессуар создан для тех, кто ценит стиль и комфорт каждый день.",
    "Модель смотрится актуально и гармонично дополняет образ.",
    "Очки становятся заметным акцентом и завершают образ.",
]

SCENARIOS = [
    "город", "путешествия", "отпуск", "прогулки",
    "вождение", "пляж", "активный отдых", "повседневное использование"
]

ENDS = [
    "Отличный выбор для тёплого сезона.",
    "Подходят для города и отдыха.",
    "Хорошо сочетаются с разными стилями одежды.",
    "Актуальны на каждый день.",
]

SEO_KEYS = [
    "очки солнцезащитные",
    "солнечные очки",
    "брендовые очки",
    "очки из инстаграм",
    "очки из tiktok",
    "очки cat eye",
    "модные очки",
    "очки женские",
    "очки мужские"
]


def _cut(text: str, limit: int) -> str:
    if len(text) <= limit:
        return text
    return text[:limit].rsplit(" ", 1)[0]


def generate_title(brand_ru: str) -> str:
    parts = [
        random.choice(SLOGANS),
        random.choice(SUN_TERMS)
    ]

    if random.random() < 0.5 and brand_ru:
        parts.append(brand_ru)

    title = " ".join(parts)
    return _cut(title.capitalize(), TITLE_MAX)


def generate_description(brand_lat: str, collection: str) -> str:
    blocks = []

    blocks.append(random.choice(INTRO))

    if brand_lat:
        blocks.append(
            f"Очки {brand_lat} подчёркивают характер образа и подходят для разных ситуаций."
        )

    if collection:
        blocks.append(
            f"Модель актуальна для сезона {collection} и хорошо смотрится в городе и на отдыхе."
        )

    blocks.append(
        f"Подойдут для таких сценариев, как {', '.join(random.sample(SCENARIOS, 4))}."
    )

    blocks.append(random.choice(ENDS))

    blocks.append(
        " ".join(random.sample(SEO_KEYS, 4)) + "."
    )

    return _cut(" ".join(blocks), DESC_MAX)


def find_col(ws, names):
    for row in ws.iter_rows(min_row=1, max_row=20):
        for cell in row:
            if cell.value and str(cell.value).strip().lower() in names:
                return cell.column
    return None


def fill_wb_template(template_path, data, style=None):
    wb = load_workbook(input_xlsx)
    ws = wb.active

    col_title = find_col(ws, {"наименование", "название"})
    col_desc = find_col(ws, {"описание", "description"})

    if not col_title or not col_desc:
        raise RuntimeError("Не найдены колонки Наименование и/или Описание")

    rows = ws.max_row
    done = 0

    # кириллица для названия
    brand_ru = brand

    for r in range(2, rows + 1):
        title = generate_title(brand_ru)
        desc = generate_description(brand, collection)

        ws.cell(row=r, column=col_title).value = title
        ws.cell(row=r, column=col_desc).value = desc

        done += 1
        if progress_callback:
            progress_callback(done / rows * 100)

    out = Path(input_xlsx).with_name(Path(input_xlsx).stem + "_ready.xlsx")
    wb.save(out)

    return str(out), done
