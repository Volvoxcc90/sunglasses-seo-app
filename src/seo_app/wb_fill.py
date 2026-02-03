import re
import time
import json
import random
from pathlib import Path
from typing import Callable, Optional, Tuple, List, Dict, Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# =====================
# НАСТРОЙКИ
# =====================
TITLE_MAX_LEN = 60
DESC_MAX_LEN = 2000

DESC_LENGTH_RANGES = {
    "short":  (550, 850),
    "medium": (900, 1400),
    "long":   (1500, 2000),
}

SEO_LEVEL_COUNTS = {
    "soft":   {"core": 1, "tail": 1, "feature": 0, "extra": 0},
    "normal": {"core": 2, "tail": 1, "feature": 1, "extra": 0},
    "hard":   {"core": 2, "tail": 2, "feature": 1, "extra": 1},
}

FORBIDDEN_LABELS = [
    "сценарии", "ключевые слова:", "форма:", "линза:", "коллекция:"
]


# =====================
# СЛОГАНЫ
# =====================
SLOGANS = [
    "Красивые", "Крутые", "Стильные", "Модные", "Молодёжные",
    "Трендовые", "Эффектные", "Дизайнерские", "Лаконичные",
    "Яркие", "Премиальные", "Удобные", "Лёгкие", "Универсальные",
    "Городские", "Летние", "Актуальные", "Смелые", "Элегантные",
    "Современные", "Ультрамодные", "Хитовые"
]


# =====================
# SEO
# =====================
SEO_CORE = [
    "солнцезащитные очки",
    "солнечные очки",
    "очки солнцезащитные",
    "брендовые очки",
]

SEO_TAIL = [
    "очки для города",
    "очки для отпуска",
    "очки для вождения",
    "аксессуар на лето",
    "инста очки",
    "очки из tiktok",
]

SEO_FEATURES = [
    "UV400",
    "поляризационные очки",
    "фотохромные очки",
]


# =====================
# СЕМАНТИЧЕСКАЯ МАТРИЦА
# =====================
SEMANTIC_MATRIX = [
    {"focus": "город", "must_tail": ["очки для города"]},
    {"focus": "вождение", "must_tail": ["очки для вождения"]},
    {"focus": "отпуск", "must_tail": ["очки для отпуска"]},
    {"focus": "соцсети", "must_tail": ["инста очки"]},
    {"focus": "универсальность", "must_tail": ["аксессуар на лето"]},
    {"focus": "охват", "must_tail": []},
]


# =====================
# БРЕНДЫ (КИРИЛЛИЦА ТОЛЬКО ДЛЯ НАИМЕНОВАНИЯ)
# =====================
BRAND_RU_OVERRIDES = {
    # ТОП
    "gucci": "Гуччи",
    "prada": "Прада",
    "dior": "Диор",
    "chanel": "Шанель",
    "versace": "Версаче",
    "fendi": "Фенди",
    "balenciaga": "Баленсиага",

    # Ray-Ban
    "ray ban": "Рэй-Бэн",
    "ray-ban": "Рэй-Бэн",
    "rayban": "Рэй-Бэн",

    # Miu Miu
    "miu miu": "Миу Миу",
    "miu-miu": "Миу Миу",
    "miumiu": "Миу Миу",

    # ic! berlin
    "ic berlin": "Айс Берлин",
    "ic-berlin": "Айс Берлин",
    "icberlin": "Айс Берлин",
    "ic! berlin": "Айс Берлин",
    "ic!berlin": "Айс Берлин",

    # прочие
    "cazal": "Казал",
    "oakley": "Окли",
    "gentle monster": "Джентл Монстер",
    "gentlemonster": "Джентл Монстер",
    "dolce gabbana": "Дольче Габбана",
    "dolcegabbana": "Дольче Габбана",
    "saint laurent": "Сен-Лоран",
    "saintlaurent": "Сен-Лоран",
}


# =====================
# ВСПОМОГАТЕЛЬНЫЕ
# =====================
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def _contains_cyrillic(s: str) -> bool:
    return bool(re.search(r"[А-Яа-яЁё]", s or ""))


def _sun_term() -> str:
    return random.choice(["солнцезащитные очки", "солнечные очки"])


def _cut(text: str, max_len: int) -> str:
    if len(text) <= max_len:
        return text
    return text[:max_len].rsplit(" ", 1)[0]


# =====================
# БРЕНД ДЛЯ НАИМЕНОВАНИЯ
# =====================
def brand_title_ru(brand_raw: str) -> str:
    b = _norm(brand_raw)
    if not b:
        return ""

    if _contains_cyrillic(b):
        return b

    b = b.lower()
    b = b.replace("–", "-").replace("—", "-")
    b = re.sub(r"[®™©!]", "", b)
    b = re.sub(r"[^a-z0-9\- ]+", " ", b)
    b = re.sub(r"\s+", " ", b).strip()

    variants = {
        b,
        b.replace(" ", ""),
        b.replace(" ", "-"),
    }

    for v in variants:
        if v in BRAND_RU_OVERRIDES:
            return BRAND_RU_OVERRIDES[v]

    return brand_raw  # если неизвестный бренд — не портим


# =====================
# НАИМЕНОВАНИЕ
# =====================
def build_titles(brand_raw: str, shape: str, lens: str) -> List[str]:
    brand_ru = brand_title_ru(brand_raw)
    shape = _norm(shape)
    lens = _norm(lens)

    flags = [True, True, True, False, False, False]
    random.shuffle(flags)

    titles = []
    for i in range(6):
        slogan = random.choice(SLOGANS)
        core = _sun_term()
        brand_part = f"{brand_ru} " if flags[i] else ""
        title = f"{slogan} {core} {brand_part}{lens} {shape}"
        titles.append(_cut(_norm(title), TITLE_MAX_LEN))

    return titles


# =====================
# ОПИСАНИЕ (БРЕНД = ЛАТИНИЦА)
# =====================
def build_description(
    brand_raw: str,
    shape: str,
    lens: str,
    collection: str,
    slot: Dict[str, Any],
) -> str:
    parts = []

    parts.append(
        f"{brand_raw} — стильный аксессуар для солнечных дней, который подчёркивает образ и помогает чувствовать себя комфортно при ярком свете."
    )

    if shape:
        parts.append(
            f"Форма оправы {shape.lower()} выглядит актуально и хорошо сочетается с повседневными и летними образами."
        )

    if lens:
        parts.append(
            f"Линзы {lens} подходят для города, поездок и отдыха, когда важно снизить дискомфорт от солнца."
        )

    if collection:
        parts.append(
            f"Модель актуальна для сезона {collection} и легко вписывается в повседневный гардероб."
        )

    parts.append(
        f"Подходит для сценариев: город, прогулки, отпуск. "
        f"Если вы ищете {', '.join(slot.get('must_tail', SEO_CORE[:2]))}, эта модель сочетает внешний вид и практичность."
    )

    text = " ".join(parts)
    return _cut(text, DESC_MAX_LEN)


# =====================
# ЗАПОЛНЕНИЕ EXCEL
# =====================
def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str,
    progress_callback: Optional[Callable[[int], None]] = None,
) -> str:

    wb = load_workbook(input_xlsx)
    ws = wb.active

    # ищем заголовки
    header = None
    for r in range(1, 15):
        names = [str(ws.cell(r, c).value).lower() for c in range(1, ws.max_column + 1)]
        if "наименование" in names and "описание" in names:
            header = r
            break
    if not header:
        raise ValueError("Не найдены колонки Наименование / Описание")

    col_name = names.index("наименование") + 1
    col_desc = names.index("описание") + 1

    rows = ws.max_row - header
    for i, r in enumerate(range(header + 1, ws.max_row + 1)):
        slot = SEMANTIC_MATRIX[(i) % len(SEMANTIC_MATRIX)]

        title = random.choice(build_titles(brand, shape, lens_features))
        desc = build_description(brand, shape, lens_features, collection, slot)

        ws.cell(r, col_name).value = title
        ws.cell(r, col_desc).value = desc

        if progress_callback:
            progress_callback(int((i + 1) * 100 / rows))

    out = Path(input_xlsx).with_name(Path(input_xlsx).stem + "_FILLED.xlsx")
    wb.save(out)
    return str(out)
