# wb_fill.py
import random
import time
import json
from pathlib import Path
from copy import deepcopy

from openpyxl import load_workbook


# -------------------------------
# УТИЛИТЫ
# -------------------------------

def _seed():
    # Гарантируем уникальный рандом КАЖДЫЙ запуск
    random.seed(time.time_ns())


def _choice(seq):
    return random.choice(seq)


def _shuffle(seq):
    seq = list(seq)
    random.shuffle(seq)
    return seq


# -------------------------------
# СЛОВАРИ / БАЗА
# -------------------------------

SLOGANS = [
    "Красивые", "Стильные", "Крутые", "Модные", "Дизайнерские",
    "Эффектные", "Современные", "Трендовые", "Лаконичные",
    "Выразительные", "Актуальные", "Премиальные", "Яркие",
    "Минималистичные", "Элегантные", "Универсальные",
    "Молодёжные", "Городские", "Летние", "Комфортные",
    "Практичные", "Инстаграмные", "ТикТок-тренд", "Фэшн",
    "С характером", "Смелые", "Лёгкие", "Статусные"
]

SCENARIOS = [
    "городские прогулки",
    "отпуск и путешествия",
    "пляж и активный отдых",
    "повседневные образы",
    "вождение",
    "город и поездки",
]

INTRO_PHRASES = [
    "Эти очки легко вписываются в современные образы и подчёркивают стиль.",
    "Модель создана для тех, кто ценит комфорт и выразительный дизайн.",
    "Очки выглядят актуально и гармонично дополняют летний гардероб.",
    "Аксессуар, который делает образ завершённым и уверенным.",
    "Идеальный вариант для яркого солнца и активного дня.",
]

STYLE_ENDINGS = [
    "Подходят для повседневной носки и отдыха.",
    "Сочетаются с городским и курортным стилем.",
    "Удобны в течение всего дня.",
    "Станут заметным акцентом образа.",
    "Хорошо смотрятся в динамичном ритме города.",
]

SEO_KEYS = [
    "солнцезащитные очки",
    "солнечные очки",
    "брендовые очки",
    "очки для города",
    "очки для лета",
    "модные очки",
    "очки UV400",
]


# -------------------------------
# ГЕНЕРАЦИЯ ТЕКСТОВ
# -------------------------------

def generate_title(brand, shape, lens):
    """
    Бренд — РАНДОМНО:
    • в ~50% случаев есть
    • в ~50% отсутствует
    Бренд — ВСЕГДА на кириллице
    """
    slogan = _choice(SLOGANS)
    base = f"{slogan} солнцезащитные очки"

    parts = [base]

    if random.random() < 0.5 and brand:
        parts.append(brand)

    if lens:
        parts.append(lens)

    if shape and random.random() < 0.5:
        parts.append(shape.lower())

    title = " ".join(parts)
    return title[:60].rstrip()


def generate_description(brand, shape, lens, collection, style):
    blocks = []

    blocks.append(_choice(INTRO_PHRASES))

    if brand:
        blocks.append(
            f"Очки {brand} отличаются продуманным дизайном и вниманием к деталям, "
            f"что делает модель актуальной в сезоне {collection}."
        )

    if shape:
        blocks.append(
            f"Форма оправы {shape.lower()} подчёркивает черты лица и смотрится уместно "
            f"как в повседневных, так и в более выразительных образах."
        )

    if lens:
        blocks.append(
            f"Линзы {lens} обеспечивают защиту от яркого солнца и комфорт "
            f"при длительном использовании."
        )

    blocks.append(
        f"Очки подойдут для таких сценариев, как {_choice(SCENARIOS)}, "
        f"и легко адаптируются под разные стили."
    )

    blocks.append(_choice(STYLE_ENDINGS))

    # SEO — мягко, внутри текста
    seo_mix = _shuffle(SEO_KEYS)[:3]
    blocks.append(" ".join(seo_mix) + ".")

    text = " ".join(blocks)

    # Немного стилистики
    if style == "premium":
        text = text.replace("очки", "аксессуар").replace("модель", "изделие")
    elif style == "social":
        text += " Отличный вариант для фото и социальных сетей."

    return text


# -------------------------------
# ОСНОВНАЯ ФУНКЦИЯ
# -------------------------------

def fill_wb_template(
    input_xlsx,
    brand="",
    shape="",
    lens_features="",
    collection="Весна–Лето 2026",
    style="neutral",
    seo_level="normal",
    desc_length="medium",
    wb_safe_mode=True,
    progress_callback=None,
):
    """
    ВОЗВРАЩАЕТ СТРОГО:
    (out_xlsx_path, rows_count, report_json)
    """

    _seed()

    wb = load_workbook(input_xlsx)
    ws = wb.active

    title_col = None
    desc_col = None

    # ищем колонки
    for col in range(1, ws.max_column + 1):
        header = str(ws.cell(row=1, column=col).value).lower()
        if "наимен" in header:
            title_col = col
        if "описан" in header:
            desc_col = col

    if not title_col or not desc_col:
        raise ValueError("Не найдены колонки 'Наименование' и/или 'Описание'")

    rows_filled = 0

    for row in range(2, ws.max_row + 1):
        title = generate_title(brand, shape, lens_features)
        desc = generate_description(brand, shape, lens_features, collection, style)

        ws.cell(row=row, column=title_col).value = title
        ws.cell(row=row, column=desc_col).value = desc

        rows_filled += 1

        if progress_callback:
            progress_callback((row - 1) / (ws.max_row - 1) * 100)

    out_path = Path(input_xlsx).with_name(
        Path(input_xlsx).stem + "_SEO.xlsx"
    )

    wb.save(out_path)

    report = {
        "rows": rows_filled,
        "brand": brand,
        "style": style,
        "collection": collection,
    }

    return str(out_path), rows_filled, json.dumps(report, ensure_ascii=False)
