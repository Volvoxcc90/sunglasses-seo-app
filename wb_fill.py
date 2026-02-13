# wb_fill.py
from __future__ import annotations

import re
import json
import random
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set

from openpyxl import load_workbook


# -----------------------------
# Helpers
# -----------------------------
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip()).lower()


def _cap_sentence(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return s
    return s[0].upper() + s[1:]


def _clamp_text(s: str, max_len: int) -> str:
    s = (s or "").strip()
    if len(s) <= max_len:
        return s
    # режем только по пробелу, чтобы не резать слова
    cut = s[:max_len].rstrip()
    if " " in cut:
        cut = cut.rsplit(" ", 1)[0].rstrip()
    return cut


def _maybe(val: str) -> bool:
    return bool((val or "").strip())


def _safe_join(parts: List[str]) -> str:
    parts = [p.strip() for p in parts if p and p.strip()]
    return " ".join(parts).strip()


def _find_header_row(ws, headers=("наименование", "описание")) -> Tuple[int, Dict[str, int]]:
    """
    Ищем строку заголовков в первых 1..20 строках и возвращаем:
    (row_index, {"name": col_index, "desc": col_index})
    """
    want = [_norm(h) for h in headers]
    for r in range(1, 21):
        mapping = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            nv = _norm(v)
            if want[0] in nv:
                mapping["name"] = c
            if want[1] in nv:
                mapping["desc"] = c
        if "name" in mapping and "desc" in mapping:
            return r, mapping
    raise ValueError("Не найдены колонки Наименование и/или Описание (заголовки не найдены в первых 20 строках).")


# -----------------------------
# Text engine
# -----------------------------
ADJECTIVES = [
    "Красивые", "Крутые", "Стильные", "Модные", "Яркие", "Эффектные", "Трендовые",
    "Удобные", "Лаконичные", "Дизайнерские", "Молодёжные", "Классические",
    "Элегантные", "Премиальные", "Актуальные", "Лёгкие", "Универсальные",
    "Дерзкие", "Винтажные", "Современные", "Городские", "Имиджевые",
    "Супер-стильные", "Топовые", "Сочные", "Невероятные", "Акцентные",
    "Минималистичные", "Роскошные", "Глянцевые"
]

SUN_WORDS = ["солнцезащитные", "солнечные"]

SCENARIOS = [
    "для города", "для отпуска", "для пляжа", "для поездок", "для прогулок",
    "для вождения", "для лета", "на каждый день", "для фоток", "для путешествий"
]

SEO_BUCKET = [
    "очки солнцезащитные", "солнцезащитные очки", "солнечные очки",
    "брендовые очки", "модные очки", "очки женские", "очки мужские", "очки унисекс",
    "инста очки", "очки из тиктока"
]

OPENERS = [
    "Очки отлично дополняют любой образ и сразу делают стиль собранным.",
    "Если нужен аксессуар «на каждый день», эти очки заходят идеально.",
    "С такими очками образ выглядит дороже и аккуратнее — без лишнего шума.",
    "Это тот случай, когда очки не просто от солнца, а реально про стиль.",
    "Лёгкий акцент, который заметен сразу: надеваешь — и образ готов.",
    "Очки смотрятся современно и легко сочетаются с любыми вещами.",
    "Удачный вариант, когда хочется и защиты, и красивого силуэта на лице.",
    "Эти очки хорошо сидят и не перегружают лицо — выглядят ровно."
]

MIDDLE_BLOCKS = [
    "Оправа {shape_lc} смотрится выразительно, подчёркивает черты лица и не выглядит громоздко.",
    "Форма {shape_lc} — универсальная: подходит под повседневные луки и под более нарядные образы.",
    "Оправа {shape_lc} добавляет характер: выглядит аккуратно, но при этом заметно.",
    "Форма {shape_lc} делает образ собранным и «дорогим» на фото и вживую.",
]

LENS_BLOCKS = [
    "Линзы {lens} дают комфорт при ярком солнце: меньше щуришься, глаза меньше устают.",
    "С {lens} проще в городе и в дороге — яркость ощущается мягче и комфортнее.",
    "Линзы {lens} — хороший выбор на лето: комфортно при дневном свете и в поездках.",
]

CLOSERS = [
    "Подойдёт как себе, так и на подарок — вариант универсальный и практичный.",
    "Берут и себе, и в подарок — вещь нужная и всегда в тему.",
    "Отличный вариант обновить аксессуары к сезону и собрать образ без усилий.",
    "Хорошая покупка на тёплый сезон: удобно, красиво и по делу.",
]

STYLE_FLAVORS = {
    "neutral": {
        "extra": ["Смотрятся аккуратно и легко сочетаются с базовой одеждой.", "Комфортная посадка на каждый день."],
    },
    "premium": {
        "extra": ["Выглядят дорого и чисто по стилю — без визуального шума.", "Акцент на детали: образ получается «люкс»."],
    },
    "social": {
        "extra": ["На фотках смотрятся огонь — прям тот самый вайб.", "Инста-образ собирается за минуту."],
    },
    "market": {
        "extra": ["Подходят для работы, учёбы, прогулок и отдыха.", "Универсальный вариант под разные сценарии."],
    },
}

SEO_LEVEL_MULT = {"low": 1, "normal": 2, "high": 3}


def make_title(
    rng: random.Random,
    brand_lat: str,
    brand_ru: str,
    shape: str,
    lens: str,
    collection: str,
    brand_in_title_ratio: float,
    used_adjs: Set[str],
) -> str:
    # уникальный лозунг/первое слово
    adj_pool = [a for a in ADJECTIVES if a not in used_adjs] or ADJECTIVES[:]
    adj = rng.choice(adj_pool)
    used_adjs.add(adj)

    sun = rng.choice(SUN_WORDS)
    parts = [adj, sun, "очки"]

    # бренд в названии (кириллица) — рандом по ratio
    if rng.random() < brand_in_title_ratio and _maybe(brand_ru):
        parts.append(brand_ru)

    # форма/линза — аккуратно
    if _maybe(shape):
        parts.append(shape.lower())
    if _maybe(lens):
        parts.append(lens.upper() if lens.lower().startswith("uv") else lens)

    # коллекцию можно коротко
    if _maybe(collection):
        # оставляем только год/сезон, если длинно
        parts.append(collection.replace("—", "-"))

    title = _safe_join(parts)
    return _clamp_text(title, 60)


def _seo_tail(rng: random.Random, level: str, gender: str) -> str:
    mult = SEO_LEVEL_MULT.get(level, 2)
    bag = SEO_BUCKET[:]
    # чуть адаптируем по полу
    if gender == "female":
        bag = [x for x in bag if x != "очки мужские"] + ["очки женские"]
    elif gender == "male":
        bag = [x for x in bag if x != "очки женские"] + ["очки мужские"]
    else:
        bag = [x for x in bag if x not in ("очки женские", "очки мужские")] + ["очки унисекс"]

    rng.shuffle(bag)
    take = min(len(bag), 2 * mult + 2)
    tail = ", ".join(bag[:take])
    # без "ключевые слова:" — просто естественным хвостом
    return tail


def make_description(
    rng: random.Random,
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    gender: str,
    style: str,
    seo_level: str,
    holiday: str,
    used_openers_global: Set[str],
) -> str:
    """
    Народная подача: как человек написал.
    Один абзац. Без меток "Коллекция:" и т.п.
    """
    opener_pool = [o for o in OPENERS if o not in used_openers_global] or OPENERS[:]
    opener = rng.choice(opener_pool)
    used_openers_global.add(opener)

    shape_lc = (shape or "").strip().lower()
    mid_tpl = rng.choice(MIDDLE_BLOCKS)
    mid = mid_tpl.format(shape_lc=shape_lc) if _maybe(shape_lc) else "Оправа выглядит аккуратно и хорошо садится на лицо."

    lens_blk = ""
    if _maybe(lens):
        lens_blk = rng.choice(LENS_BLOCKS).format(lens=lens)

    flavor = STYLE_FLAVORS.get(style, STYLE_FLAVORS["neutral"])
    extra = rng.choice(flavor["extra"])

    coll = ""
    if _maybe(collection):
        coll = f"Сезон {collection} — модель выглядит актуально и легко вписывается в летний стиль."

    scen = rng.sample(SCENARIOS, k=2)
    scen_txt = f"Подойдут {scen[0]} и {scen[1]}."

    holiday_txt = ""
    if _maybe(holiday):
        holiday_txt = f"Хороший вариант на {holiday}: и полезно, и выглядит красиво."

    closer = rng.choice(CLOSERS)

    # бренд (латиница) в тексте — 1-2 раза, не в начале двоеточием
    brand_txt = brand_lat.strip() if _maybe(brand_lat) else ""
    brand_line = f"Очки {brand_txt} смотрятся достойно и ощущаются как качественный аксессуар." if brand_txt else ""

    seo_tail = _seo_tail(rng, seo_level, gender)

    parts = [
        _cap_sentence(opener),
        brand_line,
        _cap_sentence(mid),
        _cap_sentence(lens_blk) if lens_blk else "",
        _cap_sentence(coll) if coll else "",
        _cap_sentence(extra),
        _cap_sentence(scen_txt),
        _cap_sentence(holiday_txt) if holiday_txt else "",
        _cap_sentence(closer),
        # SEO хвост
        f"По запросам: {seo_tail}."
    ]

    text = " ".join([p for p in parts if p and p.strip()])
    # подчистка двойных пробелов
    text = re.sub(r"\s+", " ", text).strip()

    # гарантируем нормальную первую букву
    text = _cap_sentence(text)

    # лимит WB до 2000
    return _clamp_text(text, 2000)


# -----------------------------
# Main fill function
# -----------------------------
@dataclass
class FillParams:
    xlsx_path: str
    out_path: str
    brand_lat: str
    brand_ru: str
    shape: str
    lens: str
    collection: str
    seo_level: str = "normal"  # low/normal/high
    style: str = "premium"     # neutral/premium/social/market
    gender: str = "auto"       # auto/female/male/unisex
    holiday: str = ""          # optional
    rows_to_fill: int = 6
    skip_top_rows: int = 4
    brand_in_title_ratio: float = 0.5
    seed: Optional[int] = None


def fill_wb_template(
    params: FillParams,
    used_openers_global: Optional[Set[str]] = None,
    progress_callback=None,
) -> Tuple[str, int]:
    """
    Заполняет ТОЛЬКО колонки Наименование/Описание.
    Не трогает первые skip_top_rows строк.
    Заполняет rows_to_fill строк подряд.
    """
    used_openers_global = used_openers_global if used_openers_global is not None else set()
    rng = random.Random(params.seed if params.seed is not None else random.randrange(1_000_000_000))

    wb = load_workbook(params.xlsx_path)
    ws = wb.active

    header_row, cols = _find_header_row(ws)
    col_name = cols["name"]
    col_desc = cols["desc"]

    start_row = max(header_row + 1, params.skip_top_rows + 1)
    end_row = start_row + max(1, int(params.rows_to_fill)) - 1

    used_adjs = set()  # чтобы названия не начинались одинаково в пределах файла

    # gender auto — оставим унисекс, если не задано
    gender = params.gender
    if gender == "auto":
        gender = "unisex"

    total = end_row - start_row + 1
    filled = 0

    for r in range(start_row, end_row + 1):
        # генерим title/desc с реальным анти-повтором
        title = make_title(
            rng=rng,
            brand_lat=params.brand_lat,
            brand_ru=params.brand_ru,
            shape=params.shape,
            lens=params.lens,
            collection=params.collection,
            brand_in_title_ratio=params.brand_in_title_ratio,
            used_adjs=used_adjs,
        )

        desc = make_description(
            rng=rng,
            brand_lat=params.brand_lat,
            shape=params.shape,
            lens=params.lens,
            collection=params.collection,
            gender=gender,
            style=params.style,
            seo_level=params.seo_level,
            holiday=params.holiday,
            used_openers_global=used_openers_global,
        )

        ws.cell(row=r, column=col_name).value = title
        ws.cell(row=r, column=col_desc).value = desc

        filled += 1
        if progress_callback:
            progress_callback(int(filled / total * 100))

    wb.save(params.out_path)
    return params.out_path, filled
