# wb_fill.py
import random
import time
import json
import re
from pathlib import Path

from openpyxl import load_workbook


TITLE_MAX = 60
DESC_MAX = 2000


# -------------------------------
# УТИЛИТЫ
# -------------------------------

def _seed():
    random.seed(time.time_ns())


def _norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _cut_no_word_break(text: str, max_len: int) -> str:
    text = _norm(text)
    if len(text) <= max_len:
        return text
    cut = text[:max_len]
    # не режем слово
    if " " in cut:
        cut = cut.rsplit(" ", 1)[0]
    return cut.strip() if cut else text[:max_len].strip()


def _find_header_row_and_cols(ws, max_scan_rows=20):
    """
    Ищем строку заголовков в первых max_scan_rows строках.
    Возвращаем: (header_row, col_name, col_desc)
    """
    for r in range(1, min(max_scan_rows, ws.max_row) + 1):
        name_col = None
        desc_col = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            lv = v.strip().lower()
            if "наимен" in lv:      # Наименование / Наимен / Наименование товара
                name_col = c
            if "описан" in lv:      # Описание
                desc_col = c
        if name_col and desc_col:
            return r, name_col, desc_col
    raise ValueError("Не найдены колонки 'Наименование' и/или 'Описание' (проверил первые 20 строк).")


# -------------------------------
# СЛОВАРИ
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

INTRO = [
    "Лёгкий акцент на тёплый сезон: модель подчёркивает стиль и помогает чувствовать себя комфортно при ярком солнце.",
    "Аксессуар, который сразу делает образ собраннее: уместно и в городе, и в поездках.",
    "Очки выглядят актуально и легко сочетаются с повседневной одеждой и летними образами.",
    "Вариант на каждый день: комфортная посадка и выразительная геометрия оправы без перегруза.",
    "Хорошо читаются в образе: добавляют уверенности и завершают стиль."
]

ENDS = [
    "Подойдут для города, отдыха и поездок — удобно, стильно и практично.",
    "Уместны в повседневной носке и в отпуске: образ становится более выразительным.",
    "Легко сочетаются с базовым гардеробом и летними луками.",
    "Хороший выбор, когда нужны и стиль, и комфорт на весь день.",
]

SCENARIOS = [
    "город", "прогулки", "отпуск", "пляж", "путешествия", "вождение", "поездки", "активный отдых"
]

SEO_KEYS = [
    "солнцезащитные очки", "солнечные очки", "очки солнцезащитные", "брендовые очки",
    "модные очки", "очки для города", "очки для вождения", "аксессуар на лето"
]


# -------------------------------
# ГЕНЕРАЦИЯ
# -------------------------------

def _sun_term():
    return random.choice(["солнцезащитные очки", "солнечные очки"])


def generate_title(brand: str, shape: str, lens: str) -> str:
    brand = _norm(brand)
    shape = _norm(shape)
    lens = _norm(lens)

    slogan = random.choice(SLOGANS)
    core = _sun_term()

    # бренд рандомно: в половине есть, в половине нет
    include_brand = (random.random() < 0.5)

    parts = [slogan, core]

    if include_brand and brand:
        parts.append(brand)

    # линзы чаще показываем, форма реже
    if lens and random.random() < 0.75:
        parts.append(lens)

    if shape and random.random() < 0.55:
        parts.append(shape.lower())

    title = " ".join(parts)
    return _cut_no_word_break(title, TITLE_MAX)


def generate_description(brand: str, shape: str, lens: str, collection: str, style: str) -> str:
    brand = _norm(brand)
    shape = _norm(shape)
    lens = _norm(lens)
    collection = _norm(collection)
    style = (_norm(style) or "neutral").lower()

    blocks = []
    blocks.append(random.choice(INTRO))

    # бренд/коллекция
    if brand and collection:
        blocks.append(f"{brand} — заметный акцент сезона {collection}: модель выглядит актуально и легко вписывается в разные стили.")
    elif brand:
        blocks.append(f"{brand} — стильный аксессуар, который подчёркивает индивидуальность и добавляет образу уверенности.")
    elif collection:
        blocks.append(f"Актуально на сезон {collection}: лёгкий аксессуар, который дополняет образ и делает его более выразительным.")

    # форма
    if shape:
        blocks.append(f"Форма оправы {shape.lower()} подчёркивает черты лица и помогает сбалансировать образ — от повседневного до более смелого.")

    # линзы
    if lens:
        blocks.append(f"Линзы {lens} подходят для яркого солнца и помогают чувствовать себя комфортнее в течение дня — особенно в городе и в поездках.")

    # сценарии (без слова “Сценарии:”)
    sc = ", ".join(random.sample(SCENARIOS, k=4))
    blocks.append(f"Подходит для ситуаций: {sc}.")

    blocks.append(random.choice(ENDS))

    # SEO внутри текста (без “SEO:”)
    seo = ", ".join(random.sample(SEO_KEYS, k=3))
    blocks.append(f"Поисковые запросы, по которым часто ищут: {seo}.")

    text = " ".join(blocks)

    # стили
    if style == "premium":
        text = text.replace("аксессуар", "премиальный аксессуар").replace("стильный", "премиальный")
    elif style == "social":
        text += " Смотрится эффектно в кадре — хороший вариант для фото и соцсетей."

    return _cut_no_word_break(text, DESC_MAX)


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
    Возвращает строго 3 значения:
    (out_xlsx_path, rows_count, report_json)
    """
    _seed()

    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, col_name, col_desc = _find_header_row_and_cols(ws, max_scan_rows=20)

    start_row = header_row + 1
    end_row = ws.max_row

    rows_filled = 0
    seen_titles = set()
    seen_desc_starts = set()

    total = max(1, end_row - start_row + 1)

    for i, r in enumerate(range(start_row, end_row + 1), start=1):
        # генерируем без дублей (несколько попыток)
        title = ""
        for _ in range(12):
            t = generate_title(brand, shape, lens_features)
            if t not in seen_titles:
                title = t
                break
        if not title:
            title = generate_title(brand, shape, lens_features)
        seen_titles.add(title)

        desc = ""
        for _ in range(12):
            d = generate_description(brand, shape, lens_features, collection, style)
            start7 = " ".join(d.split()[:7]).lower()
            if start7 not in seen_desc_starts:
                desc = d
                seen_desc_starts.add(start7)
                break
        if not desc:
            desc = generate_description(brand, shape, lens_features, collection, style)

        # перезаписываем наименование/описание
        ws.cell(r, col_name).value = title
        ws.cell(r, col_desc).value = desc

        rows_filled += 1

        if progress_callback:
            progress_callback(int(i * 100 / total))

    out_path = Path(input_xlsx).with_name(Path(input_xlsx).stem + "_SEO.xlsx")
    wb.save(out_path)

    report = {
        "rows_filled": rows_filled,
        "header_row": header_row,
        "name_col": col_name,
        "desc_col": col_desc,
        "brand": brand,
        "shape": shape,
        "lens": lens_features,
        "collection": collection,
        "style": style
    }

    return str(out_path), rows_filled, json.dumps(report, ensure_ascii=False, indent=2)
