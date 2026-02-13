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


def _cap(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return s
    return s[0].upper() + s[1:]


def _clamp_text(s: str, max_len: int) -> str:
    s = (s or "").strip()
    if len(s) <= max_len:
        return s
    cut = s[:max_len].rstrip()
    if " " in cut:
        cut = cut.rsplit(" ", 1)[0].rstrip()
    return cut


def _tokens(text: str) -> Set[str]:
    text = _norm(text)
    # токены только из букв/цифр
    arr = re.findall(r"[a-zа-я0-9]+", text, flags=re.IGNORECASE)
    # фильтруем мусор
    stop = {
        "и", "а", "но", "или", "что", "это", "в", "на", "для", "как", "по", "из",
        "с", "к", "у", "о", "же", "то", "мы", "вы", "они", "он", "она", "оно",
        "этот", "эта", "эти", "там", "тут", "при", "всё", "все", "очень"
    }
    return {t for t in arr if len(t) >= 3 and t not in stop}


def jaccard(a: Set[str], b: Set[str]) -> float:
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / max(1, union)


def _safe_join(parts: List[str]) -> str:
    parts = [p.strip() for p in parts if p and p.strip()]
    return " ".join(parts).strip()


def _find_header_row(ws, headers=("наименование", "описание")) -> Tuple[int, Dict[str, int]]:
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
# Text engine (народная подача)
# -----------------------------
ADJECTIVES = [
    "Красивые", "Крутые", "Стильные", "Модные", "Яркие", "Эффектные", "Трендовые",
    "Удобные", "Лаконичные", "Дизайнерские", "Молодёжные", "Классические",
    "Элегантные", "Премиальные", "Актуальные", "Лёгкие", "Универсальные",
    "Дерзкие", "Винтажные", "Современные", "Городские", "Имиджевые",
    "Топовые", "Сочные", "Акцентные", "Минималистичные", "Роскошные"
]

SUN_WORDS = ["солнцезащитные", "солнечные"]

SCENARIOS = [
    "для города", "для отпуска", "для пляжа", "для поездок", "для прогулок",
    "для вождения", "на каждый день", "для фоток", "для путешествий"
]

OPENERS = [
    "Очки — это как финальный штрих: надел и образ сразу собран.",
    "Если нужен аксессуар «на каждый день», эти очки заходят идеально.",
    "Такие очки делают образ дороже и аккуратнее — без лишнего шума.",
    "Это тот случай, когда очки не просто от солнца, а реально про стиль.",
    "Лёгкий акцент, который заметен сразу: надеваешь — и всё выглядит ровно.",
    "Очки смотрятся современно и легко сочетаются с любыми вещами.",
    "Удачный вариант, когда хочется и защиты, и красивого силуэта на лице.",
    "Хорошая посадка + нормальный дизайн — и очки реально носишь каждый день."
]

MIDDLE_BLOCKS = {
    "neutral": [
        "Оправа {shape_lc} выглядит аккуратно и подходит под базовые образы.",
        "Форма {shape_lc} универсальная: и к джинсам, и к более нарядному стилю.",
        "Оправа {shape_lc} добавляет характер, но не перегружает лицо.",
    ],
    "premium": [
        "Оправа {shape_lc} смотрится «дорого»: чистые линии, аккуратный силуэт на лице.",
        "Форма {shape_lc} делает образ собранным — как будто всё продумано заранее.",
        "Оправа {shape_lc} даёт правильный акцент: заметно, но без перебора.",
    ],
    "social": [
        "Оправа {shape_lc} на фото выглядит огонь — прям тот самый вайб.",
        "Форма {shape_lc} делает лицо выразительнее, и кадр сразу сильнее.",
        "Оправа {shape_lc} — тот вариант, который любят за «вау-эффект».",
    ],
    "market": [
        "Оправа {shape_lc} удобная и понятная: под работу, прогулки и отдых.",
        "Форма {shape_lc} подходит и под повседневный стиль, и под выход.",
        "Оправа {shape_lc} хорошо сочетается с разной одеждой и обувью.",
    ],
}

LENS_BLOCKS = [
    "Линзы {lens} дают комфорт при ярком солнце: меньше щуришься, глаза меньше устают.",
    "С {lens} проще в городе и в дороге — яркость ощущается мягче и комфортнее.",
    "Линзы {lens} — хороший выбор на тёплый сезон: комфортно при дневном свете и в поездках.",
]

CLOSERS = [
    "Подойдёт и себе, и на подарок — вещь нужная и всегда в тему.",
    "Берут и себе, и в подарок — выглядит стильно и по делу.",
    "Отличный вариант обновить аксессуары к сезону и собрать образ без усилий.",
    "Хорошая покупка на тёплый сезон: удобно, красиво и практично.",
]

# SEO вставляем “внутрь текста”, без ярлыков
SEO_BASE = [
    "солнцезащитные очки", "солнечные очки", "брендовые очки", "модные очки",
    "очки унисекс", "инста очки", "очки для отпуска"
]


SEO_LEVEL_COUNT = {"low": 2, "normal": 4, "high": 6}


def _seo_pack(rng: random.Random, gender: str, level: str) -> List[str]:
    bag = SEO_BASE[:]
    if gender == "female":
        bag += ["очки женские"]
    elif gender == "male":
        bag += ["очки мужские"]
    else:
        bag += ["очки унисекс"]
    rng.shuffle(bag)
    n = SEO_LEVEL_COUNT.get(level, 4)
    return bag[:min(n, len(bag))]


def _inject_seo_naturally(rng: random.Random, text: str, seo_phrases: List[str]) -> str:
    """
    Вставляем ключи естественно: 2–3 в середине, 1–2 ближе к концу.
    Без “Ключевые слова:” и без хвоста-перечня.
    """
    if not seo_phrases:
        return text

    # гарантируем, что в тексте будет хотя бы один из “солнцезащитные очки / солнечные очки”
    core = rng.choice(["солнцезащитные очки", "солнечные очки"])
    if core not in text.lower():
        # вставим в первое предложение
        text = re.sub(r"^(Очки\s—\sэто|Если нужен|Такие очки|Это тот случай|Лёгкий акцент|Очки смотрятся|Удачный вариант|Хорошая посадка)",
                      lambda m: m.group(0) + f" {core} ", text, count=1)

    # остальные ключи — в 2 коротких вставках
    rest = [p for p in seo_phrases if p.lower() not in text.lower()]
    if not rest:
        return text

    rng.shuffle(rest)

    mid_take = rest[:max(1, len(rest)//2)]
    end_take = rest[len(mid_take):]

    # вставка в середину: “… — как {k1} / {k2}…”
    if mid_take:
        k = " / ".join(mid_take[:3])
        text += f" По сути это {k} — без лишних слов."

    # вставка ближе к концу: “Часто берут как …”
    if end_take:
        k2 = ", ".join(end_take[:3])
        text += f" Часто берут как {k2}."

    return re.sub(r"\s+", " ", text).strip()


def make_title(
    rng: random.Random,
    brand_ru: str,
    shape: str,
    lens: str,
    collection: str,
    brand_in_title_ratio: float,
    used_first_words: Set[str],
    used_titles: Set[str],
) -> str:
    # анти-монотонность: стараемся не повторять первое слово
    pool = [a for a in ADJECTIVES if a not in used_first_words] or ADJECTIVES[:]
    first = rng.choice(pool)
    used_first_words.add(first)

    sun = rng.choice(SUN_WORDS)
    parts = [first, sun, "очки"]

    if rng.random() < brand_in_title_ratio and brand_ru.strip():
        parts.append(brand_ru.strip())

    if shape.strip():
        parts.append(shape.strip().lower())
    if lens.strip():
        parts.append(lens.strip().upper() if lens.strip().lower().startswith("uv") else lens.strip())

    if collection.strip():
        parts.append(collection.replace("—", "-").strip())

    title = _clamp_text(_safe_join(parts), 60)

    # анти-дубли названий (перегенерим пару раз)
    tries = 0
    while title.lower() in used_titles and tries < 15:
        tries += 1
        first = rng.choice(ADJECTIVES)
        sun = rng.choice(SUN_WORDS)
        parts = [first, sun, "очки"]
        if rng.random() < brand_in_title_ratio and brand_ru.strip():
            parts.append(brand_ru.strip())
        if shape.strip():
            parts.append(shape.strip().lower())
        if lens.strip():
            parts.append(lens.strip().upper() if lens.strip().lower().startswith("uv") else lens.strip())
        if collection.strip():
            parts.append(collection.replace("—", "-").strip())
        title = _clamp_text(_safe_join(parts), 60)

    used_titles.add(title.lower())
    return title


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
    holiday_pos: str,  # "middle" | "end"
    used_openers_global: Set[str],
    used_desc_tokens: List[Set[str]],
) -> str:
    # анти-монотонность: не повторяем стартовые фразы в пачке
    opener_pool = [o for o in OPENERS if o not in used_openers_global] or OPENERS[:]
    opener = rng.choice(opener_pool)
    used_openers_global.add(opener)

    shape_lc = shape.strip().lower() if shape else ""
    mid_list = MIDDLE_BLOCKS.get(style, MIDDLE_BLOCKS["neutral"])
    mid = rng.choice(mid_list).format(shape_lc=shape_lc) if shape_lc else "Оправа выглядит аккуратно и хорошо садится на лицо."

    lens_blk = ""
    if lens.strip():
        lens_blk = rng.choice(LENS_BLOCKS).format(lens=lens.strip())

    coll_blk = ""
    if collection.strip():
        coll_blk = f"На сезон {collection.strip()} — вариант актуальный и легко сочетается с летними образами."

    scen = rng.sample(SCENARIOS, k=2)
    scen_blk = f"Подойдёт {scen[0]} и {scen[1]}."

    brand_blk = ""
    if brand_lat.strip():
        # НЕ начинаем с "бренд:" и без двоеточия
        brand_blk = f"Очки {brand_lat.strip()} выглядят аккуратно и добавляют образу статусный акцент."

    holiday_blk = ""
    if holiday.strip():
        holiday_blk = f"И ещё момент: на {holiday.strip()} — отличный вариант в подарок, потому что вещь полезная и выглядит красиво."

    closer = rng.choice(CLOSERS)

    base_parts_middle = [
        _cap(opener),
        brand_blk,
        _cap(mid),
        _cap(lens_blk) if lens_blk else "",
        _cap(coll_blk) if coll_blk else "",
        _cap(scen_blk),
    ]
    base_parts_end = [
        closer,
    ]

    if holiday_blk:
        if holiday_pos == "middle":
            base_parts_middle.append(_cap(holiday_blk))
        else:
            base_parts_end.append(_cap(holiday_blk))

    # SEO внутрь текста
    seo_phrases = _seo_pack(rng, gender, seo_level)

    text = " ".join([p for p in base_parts_middle if p and p.strip()])
    text = _inject_seo_naturally(rng, text, seo_phrases)
    text = " ".join([text] + [p for p in base_parts_end if p and p.strip()])
    text = re.sub(r"\s+", " ", text).strip()
    text = _cap(text)
    text = _clamp_text(text, 2000)

    # анти-дубли описаний: проверяем похожесть и перегенерим
    tok = _tokens(text)
    tries = 0
    while tries < 20:
        too_close = False
        for prev in used_desc_tokens:
            if jaccard(tok, prev) >= 0.55:  # порог “похожести”
                too_close = True
                break
        if not too_close:
            break

        tries += 1
        # перегенерация: меняем блоки + перемешиваем
        opener = rng.choice(OPENERS)
        mid = rng.choice(mid_list).format(shape_lc=shape_lc) if shape_lc else "Оправа выглядит аккуратно и хорошо садится на лицо."
        lens_blk = rng.choice(LENS_BLOCKS).format(lens=lens.strip()) if lens.strip() else ""
        closer = rng.choice(CLOSERS)
        scen = rng.sample(SCENARIOS, k=2)
        scen_blk = f"Подойдёт {scen[0]} и {scen[1]}."
        seo_phrases = _seo_pack(rng, gender, seo_level)

        base_parts_middle = [
            _cap(opener),
            brand_blk,
            _cap(mid),
            _cap(lens_blk) if lens_blk else "",
            _cap(coll_blk) if coll_blk else "",
            _cap(scen_blk),
        ]
        base_parts_end = [closer]
        if holiday_blk:
            if holiday_pos == "middle":
                base_parts_middle.append(_cap(holiday_blk))
            else:
                base_parts_end.append(_cap(holiday_blk))

        text = " ".join([p for p in base_parts_middle if p and p.strip()])
        text = _inject_seo_naturally(rng, text, seo_phrases)
        text = " ".join([text] + [p for p in base_parts_end if p and p.strip()])
        text = re.sub(r"\s+", " ", text).strip()
        text = _cap(text)
        text = _clamp_text(text, 2000)
        tok = _tokens(text)

    used_desc_tokens.append(tok)
    return text


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
    seo_level: str = "normal"        # low/normal/high
    style: str = "premium"           # premium/market/social/neutral
    gender: str = "auto"             # auto/female/male/unisex
    holiday: str = ""
    holiday_pos: str = "middle"      # middle/end
    rows_to_fill: int = 6
    skip_top_rows: int = 4
    brand_in_title_ratio: float = 0.5
    seed: Optional[int] = None


def fill_wb_template(
    params: FillParams,
    used_openers_global: Optional[Set[str]] = None,
    used_openers_pack: Optional[Set[str]] = None,
    progress_callback=None,
) -> Tuple[str, int]:
    """
    Заполняет ТОЛЬКО колонки Наименование/Описание.
    Не трогает первые skip_top_rows строк вообще.
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

    # gender auto → unisex (чтобы не лепить “женские/мужские” если не надо)
    gender = params.gender
    if gender == "auto":
        gender = "unisex"

    used_first_words: Set[str] = set()
    used_titles: Set[str] = set()
    used_desc_tokens: List[Set[str]] = []

    total = end_row - start_row + 1
    filled = 0

    for r in range(start_row, end_row + 1):
        title = make_title(
            rng=rng,
            brand_ru=params.brand_ru,
            shape=params.shape,
            lens=params.lens,
            collection=params.collection,
            brand_in_title_ratio=params.brand_in_title_ratio,
            used_first_words=used_first_words,
            used_titles=used_titles,
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
            holiday_pos=params.holiday_pos,
            used_openers_global=used_openers_global,
            used_desc_tokens=used_desc_tokens,
        )

        ws.cell(row=r, column=col_name).value = title
        ws.cell(row=r, column=col_desc).value = desc

        filled += 1
        if progress_callback:
            progress_callback(int(filled / total * 100))

    wb.save(params.out_path)
    return params.out_path, filled
