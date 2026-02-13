# wb_fill.py
from __future__ import annotations

import os
import re
import json
import time
import math
import random
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Callable, Set

from openpyxl import load_workbook


# ----------------------------
# Helpers
# ----------------------------
def _norm_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("&", " ").replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _safe_filename(name: str) -> str:
    name = (name or "").strip()
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:120] if name else "output"


def _cap_first(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return s
    return s[0].upper() + s[1:]


def _join_ru_list(items: List[str]) -> str:
    items = [x.strip() for x in items if x and x.strip()]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} и {items[1]}"
    return f"{', '.join(items[:-1])} и {items[-1]}"


def _jaccard(a: str, b: str) -> float:
    wa = set(re.findall(r"[a-zA-Zа-яА-Я0-9]+", (a or "").lower()))
    wb = set(re.findall(r"[a-zA-Zа-яА-Я0-9]+", (b or "").lower()))
    if not wa and not wb:
        return 1.0
    inter = len(wa & wb)
    uni = len(wa | wb)
    return inter / max(uni, 1)


# ----------------------------
# Parameters
# ----------------------------
@dataclass
class FillParams:
    xlsx_path: str
    output_dir: str

    brand_lat: str
    brand_ru: str
    shape: str
    lenses: str
    collection: str

    holidays: str            # "8 Марта||14 Февраля||Новый год"
    holiday_pos: str         # start/middle/end

    seo_level: str           # low/normal/high
    style: str               # neutral/premium/social/mass
    wb_safe_mode: bool
    wb_strict: bool

    brand_in_title_ratio: str  # "0/100", "50/50", "100/0"
    rows_to_fill: int
    skip_first_rows: int

    batch_count: int

    # uniqueness knobs
    uniqueness: int = 92     # 0..100
    progress_callback: Optional[Callable[[int], None]] = None


# ----------------------------
# “Live” text blocks
# ----------------------------
SLOGANS = [
    "Красивые", "Крутые", "Стильные", "Модные", "Молодёжные",
    "Дизайнерские", "Трендовые", "Эффектные", "Лаконичные", "Яркие",
    "Ультрамодные", "Новые", "Актуальные", "Премиальные", "Классные",
    "Элегантные", "Сочные", "Выразительные", "Нарядные", "Минималистичные",
    "Удобные", "Лёгкие", "Городские", "Летние", "Топовые",
    "Любимые", "Универсальные", "Супер-стильные", "Нереально красивые", "Свежие",
]

PRODUCT_WORDS = ["солнцезащитные очки", "солнечные очки"]

SHAPE_HINTS = {
    "кошачий глаз": ["кошачий глаз", "cat eye"],
    "квадратные": ["квадратные", "квадратная оправа"],
    "квадратная": ["квадратная оправа", "квадратные"],
    "вайфареры": ["вайфареры", "wayfarer"],
    "авиаторы": ["авиаторы", "aviator"],
    "овальные": ["овальные", "овальная оправа"],
    "круглые": ["круглые", "круглая оправа"],
    "прямоугольные": ["прямоугольные", "прямоугольная оправа"],
}

LENS_HINTS = {
    "uv400": ["UV400", "защита UV400"],
    "поляризационные": ["поляризация", "поляризационные линзы"],
    "фотохромные": ["фотохромные", "хамелеон"],
    "градиентные": ["градиентные", "градиент"],
    "зеркальные": ["зеркальные", "зеркалка"],
}

SCENARIOS = [
    "город", "путешествия", "прогулки", "пляж", "отпуск", "дорога", "поездки",
    "вождение", "парк", "летние прогулки", "дневные выходы", "выходные",
]

GIFTS = ["подарок", "подарок девушке", "подарок парню", "подарок жене", "подарок мужу", "подарок подруге"]

HOLIDAYS_DEFAULT = ["8 Марта", "14 Февраля", "Новый год", "День рождения", "Выпускной", "День матери", "23 Февраля"]

SEO_KEYS_COMMON = [
    "очки солнцезащитные", "солнечные очки", "брендовые очки", "модные очки",
    "очки женские", "очки мужские", "очки унисекс", "очки для отпуска",
    "очки для города", "очки UV400", "инста очки", "очки из TikTok",
]

RISK_WORDS = [
    r"\b100%\b", r"\bлучшие\b", r"\bсамые лучшие\b", r"\bгарантированно\b",
    r"\bвылечит\b", r"\bабсолютно\b", r"\bидеально\b",
]

STOP_PHRASES_STRICT = [
    r"\bпо факту\b", r"\bпрям\b", r"\bреально\b", r"\bтоп\b",
]


def _pick_shape_phrase(rnd: random.Random, shape: str) -> str:
    s = (shape or "").strip()
    k = _norm_key(s)
    for key, variants in SHAPE_HINTS.items():
        if key in k:
            return rnd.choice(variants)
    return s.lower() if s else ""


def _pick_lens_phrase(rnd: random.Random, lenses: str) -> str:
    s = (lenses or "").strip()
    k = _norm_key(s)
    for key, variants in LENS_HINTS.items():
        if key in k:
            return rnd.choice(variants)
    return s if s else ""


def _seo_pack(rnd: random.Random, level: str, lenses: str, shape: str) -> List[str]:
    level = (level or "normal").lower().strip()
    keys = SEO_KEYS_COMMON[:]

    # add hints
    lp = _pick_lens_phrase(rnd, lenses)
    sp = _pick_shape_phrase(rnd, shape)
    if lp:
        keys.append(f"очки {lp}".lower())
    if sp:
        keys.append(f"очки {sp}".lower())

    # unique shuffle
    rnd.shuffle(keys)

    if level == "low":
        return keys[:4]
    if level == "high":
        return keys[:9]
    return keys[:6]


def _insert_holidays_block(rnd: random.Random, holidays: str) -> str:
    raw = (holidays or "").strip()
    if raw:
        items = [x.strip() for x in raw.split("||") if x.strip()]
    else:
        # sometimes none
        items = []
    if not items:
        return ""

    joined = _join_ru_list(items)

    variants = [
        f"Часто берут {rnd.choice(GIFTS)} к {joined}: аксессуар заметный и полезный.",
        f"К {joined} — отличный вариант, если хочется подарок “и красивый, и нужный”.",
        f"На {joined} такие очки берут часто: и образ собирают, и глаза защищают.",
    ]
    return rnd.choice(variants)


def _apply_safe_mode(text: str) -> str:
    t = text
    for pat in RISK_WORDS:
        t = re.sub(pat, "", t, flags=re.IGNORECASE)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def _apply_strict(text: str) -> str:
    t = text
    for pat in STOP_PHRASES_STRICT:
        t = re.sub(pat, "", t, flags=re.IGNORECASE)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def _make_title(
    rnd: random.Random,
    brand_lat: str,
    brand_ru: str,
    shape: str,
    lenses: str,
    collection: str,
    ratio: str,
    used_titles: Set[str],
) -> str:
    # First word must be slogan (per your requirement)
    slogan = rnd.choice(SLOGANS)
    prod = rnd.choice(PRODUCT_WORDS)

    # choose include brand in title or not
    ratio = (ratio or "50/50").strip()
    include_brand = True
    if ratio == "0/100":
        include_brand = False
    elif ratio == "100/0":
        include_brand = True
    else:
        include_brand = rnd.random() < 0.5

    # shape/lens snippets (optional)
    sp = _pick_shape_phrase(rnd, shape)
    lp = _pick_lens_phrase(rnd, lenses)

    extras = []
    # add either lens or shape or both
    if lp and rnd.random() < 0.70:
        extras.append(lp)
    if sp and rnd.random() < 0.55:
        # do not shout in CAPS, use normal
        extras.append(sp)

    # collection snippet sometimes
    if collection and rnd.random() < 0.35:
        extras.append(collection.replace("–", "-"))

    # build candidates with small permutations
    parts = [slogan, prod]
    if include_brand and brand_ru:
        parts.append(brand_ru)  # TITLE uses RU
    parts.extend(extras)

    # keep <= 60 chars, avoid cut words: just drop last parts until fits
    def join(p: List[str]) -> str:
        return " ".join([x for x in p if x]).strip()

    t = join(parts)
    while len(t) > 60 and len(parts) > 2:
        parts.pop()  # drop tail
        t = join(parts)

    # if still too long, shorten by removing collection first
    if len(t) > 60:
        t = t[:60].rstrip()

    # anti-duplicate within generation
    base = _norm_key(t)
    if base in used_titles:
        # slight variation
        for _ in range(6):
            slogan2 = rnd.choice([s for s in SLOGANS if s != slogan] or SLOGANS)
            parts2 = parts[:]
            parts2[0] = slogan2
            t2 = join(parts2)
            while len(t2) > 60 and len(parts2) > 2:
                parts2.pop()
                t2 = join(parts2)
            if _norm_key(t2) not in used_titles:
                t = t2
                break

    used_titles.add(_norm_key(t))
    return t


def _make_description(
    rnd: random.Random,
    brand_lat: str,
    shape: str,
    lenses: str,
    collection: str,
    holidays: str,
    holiday_pos: str,
    seo_level: str,
    style: str,
    wb_safe: bool,
    wb_strict: bool,
    used_first_phrases: Set[str],
    used_descs: List[str],
    uniqueness: int,
) -> str:
    # We want “народная” подача, но логично, как в твоём примере.
    # No labels like "Коллекция:" "Сценарии:" etc.

    sp = _pick_shape_phrase(rnd, shape)
    lp = _pick_lens_phrase(rnd, lenses)

    # first phrase pool (anti-monotony)
    first_pool = [
        "Очки — отличный аксессуар на каждый день: и образ собирают, и глаза бережёт от яркого солнца.",
        "Эти очки легко вписываются в любой образ — от повседневного до более нарядного.",
        "Если хочется добавить образу акцент — такие очки делают это быстро и без лишнего шума.",
        "Очки смотрятся аккуратно и дорого: подходят и на каждый день, и на поездки, и на отпуск.",
        "Универсальный вариант: можно носить в городе, на отдыхе и просто на прогулках.",
        "Это тот самый аксессуар, который “делает” образ — спокойно, уверенно и со вкусом.",
    ]
    rnd.shuffle(first_pool)

    first = first_pool[0]
    # anti-duplicate starts
    for cand in first_pool:
        k = _norm_key(cand)
        if k not in used_first_phrases:
            first = cand
            used_first_phrases.add(k)
            break

    # middle blocks
    blocks = []

    # style shaping
    style = (style or "neutral").lower().strip()
    if style == "premium":
        blocks.append("Визуально очки выглядят собранно: линии ровные, посадка аккуратная, образ получается “дороже”.")
    elif style == "social":
        blocks.append("На фото смотрятся очень эффектно — прям тот аксессуар, который сразу цепляет.")
    elif style == "mass":
        blocks.append("Простой понятный вариант: носить удобно, выглядит хорошо, подходит под разные вещи.")
    else:
        blocks.append("Сидят комфортно, не перегружают лицо и подходят под разные стили одежды.")

    # shape paragraph
    if sp:
        blocks.append(
            f"Форма {sp} подчёркивает черты лица и добавляет образу выразительности. "
            f"Смотрится гармонично и в повседневном стиле, и в более нарядном."
        )

    # lens paragraph
    if lp:
        if "uv400" in _norm_key(lp):
            blocks.append("Линзы UV400 помогают чувствовать себя комфортно при ярком солнце — хороший вариант для города, дороги и отдыха.")
        elif "поляр" in _norm_key(lp):
            blocks.append("Поляризационные линзы уменьшают блики — удобно за рулём, у воды и в солнечные дни в городе.")
        elif "фотох" in _norm_key(lp) or "хамелеон" in _norm_key(lp):
            blocks.append("Фотохромные линзы (хамелеон) подстраиваются под свет — комфортнее, когда освещение меняется в течение дня.")
        else:
            blocks.append(f"Линзы: {lp}. Комфортно в солнечную погоду и в активных сценариях дня.")

    # scenarios
    sc = rnd.sample(SCENARIOS, k=4)
    blocks.append(
        f"Подходит для таких сценариев: {', '.join(sc)}. "
        f"Можно брать себе или на подарок — практично и красиво."
    )

    # collection mention (no label)
    if collection and rnd.random() < 0.75:
        blocks.append(f"Сезон {collection}: модель выглядит актуально и легко сочетается с летним гардеробом.")

    # holiday block
    hb = _insert_holidays_block(rnd, holidays)
    if hb:
        if (holiday_pos or "middle").lower() == "start":
            blocks.insert(0, hb)
        elif (holiday_pos or "middle").lower() == "end":
            blocks.append(hb)
        else:
            # middle
            blocks.insert(min(2, len(blocks)), hb)

    # SEO keys block (народно, но без “Ключевые слова:”)
    keys = _seo_pack(rnd, seo_level, lenses, shape)
    # make it look not like machine: weave in a sentence
    keys_sentence_templates = [
        "По запросам люди ищут так: {keys}.",
        "Если подбирать по поиску, обычно ищут: {keys}.",
        "В поиске часто пишут: {keys}.",
    ]
    keys_sentence = rnd.choice(keys_sentence_templates).format(keys=", ".join(keys))

    # final assemble
    desc_parts = [first] + blocks + [keys_sentence]
    text = " ".join([_cap_first(p).strip().rstrip(".") + "." for p in desc_parts if p and p.strip()])

    # remove double dots/spaces
    text = re.sub(r"\.\s*\.", ".", text)
    text = re.sub(r"\s{2,}", " ", text).strip()

    # brand in description should be LATIN (per your requirement)
    if brand_lat:
        # insert brand naturally (not "Brand:")
        inserts = [
            f"Модель {brand_lat} хорошо вписывается в базовый гардероб и в более яркие образы.",
            f"Очки {brand_lat} — удачный вариант, если нравится аккуратный брендовый стиль.",
            f"{brand_lat} смотрится уверенно: можно носить каждый день.",
        ]
        ins = rnd.choice(inserts)
        # put near the beginning, but not as "Brand:"
        text = re.sub(r"^([^\.]+\.)", r"\1 " + ins, text, count=1).strip()

    # strict/safe
    if wb_safe:
        text = _apply_safe_mode(text)
    if wb_strict:
        text = _apply_strict(text)

    # uniqueness check (anti near-duplicates)
    target = max(0.18, (100 - max(0, min(100, uniqueness))) / 100.0)  # uniqueness 92 => ~0.08..0.12
    for prev in used_descs[-12:]:
        if _jaccard(text, prev) > (0.55 - target):
            # mutate by shuffling blocks and changing first sentence
            rnd.shuffle(blocks)
            first2 = rnd.choice([x for x in first_pool if _norm_key(x) not in used_first_phrases] or first_pool)
            used_first_phrases.add(_norm_key(first2))
            desc_parts2 = [first2] + blocks + [keys_sentence]
            text2 = " ".join([_cap_first(p).strip().rstrip(".") + "." for p in desc_parts2 if p and p.strip()])
            text2 = re.sub(r"\s{2,}", " ", text2).strip()
            if wb_safe:
                text2 = _apply_safe_mode(text2)
            if wb_strict:
                text2 = _apply_strict(text2)
            text = text2
            break

    used_descs.append(text)
    return text


# ----------------------------
# Excel fill
# ----------------------------
def _find_col_by_header(ws, header_row: int, names: List[str]) -> Optional[int]:
    wanted = {_norm_key(n) for n in names}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=col).value
        if v is None:
            continue
        k = _norm_key(str(v))
        if k in wanted:
            return col
    return None


def _detect_header_row(ws, max_scan: int = 30) -> int:
    # find row that contains both "Наименование" and "Описание"
    for r in range(1, min(max_scan, ws.max_row) + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, min(ws.max_column, 50) + 1)]
        joined = " | ".join([str(x) for x in row_vals if x is not None])
        if not joined:
            continue
        j = _norm_key(joined)
        if "наимен" in j and "описан" in j:
            return r
    return 1


def fill_wb_template(params: FillParams) -> Tuple[List[str], int, str]:
    """
    Returns:
      (output_paths, rows_filled_total, report_json_str)
    """
    rnd_master = random.Random()
    rnd_master.seed(time.time_ns() ^ (hash(params.xlsx_path) & 0xFFFFFFFF))

    in_path = Path(params.xlsx_path)
    out_dir = Path(params.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    total_filled = 0
    outputs: List[str] = []

    # track anti-duplicates across the whole batch
    used_first_phrases: Set[str] = set()
    used_titles: Set[str] = set()
    used_descs: List[str] = []

    # for progress
    total_steps = max(1, params.batch_count)
    done_steps = 0

    for i in range(1, params.batch_count + 1):
        wb = load_workbook(in_path)
        ws = wb.active

        header_row = _detect_header_row(ws)
        name_col = _find_col_by_header(ws, header_row, ["Наименование", "Название", "Заголовок", "Наим-е"])
        desc_col = _find_col_by_header(ws, header_row, ["Описание", "Description", "Опис-е"])

        if not name_col or not desc_col:
            raise ValueError("Не найдены колонки Наименование и/или Описание (проверь заголовки в файле).")

        # rows start after header row
        start_row = header_row + 1

        # don't touch first N rows (absolute rows in sheet)
        skip_until = max(0, int(params.skip_first_rows))
        # eligible rows: >= start_row and > skip_until
        eligible_rows = [r for r in range(start_row, ws.max_row + 1) if r > skip_until]

        # fill only first N eligible rows
        rows_to_fill = max(0, int(params.rows_to_fill))
        eligible_rows = eligible_rows[:rows_to_fill]

        # If sheet is shorter, still fine
        rows_filled = 0

        # per-file random seed
        rnd = random.Random()
        rnd.seed((time.time_ns() & 0xFFFFFFFFFFFF) ^ (i * 99991) ^ (hash(params.brand_lat) & 0xFFFFFFFF))

        for r in eligible_rows:
            title = _make_title(
                rnd=rnd,
                brand_lat=params.brand_lat,
                brand_ru=params.brand_ru,
                shape=params.shape,
                lenses=params.lenses,
                collection=params.collection,
                ratio=params.brand_in_title_ratio,
                used_titles=used_titles,
            )

            desc = _make_description(
                rnd=rnd,
                brand_lat=params.brand_lat,   # description uses LATIN brand
                shape=params.shape,
                lenses=params.lenses,
                collection=params.collection,
                holidays=params.holidays,
                holiday_pos=params.holiday_pos,
                seo_level=params.seo_level,
                style=params.style,
                wb_safe=params.wb_safe_mode,
                wb_strict=params.wb_strict,
                used_first_phrases=used_first_phrases,
                used_descs=used_descs,
                uniqueness=params.uniqueness,
            )

            # overwrite always
            ws.cell(row=r, column=name_col).value = title
            ws.cell(row=r, column=desc_col).value = desc

            rows_filled += 1

        total_filled += rows_filled

        base = _safe_filename(in_path.stem)
        out_name = f"{base}_{i:02d}.xlsx" if params.batch_count > 1 else f"{base}_out.xlsx"
        out_path = out_dir / out_name
        wb.save(out_path)
        outputs.append(str(out_path))

        done_steps += 1
        if params.progress_callback:
            params.progress_callback(int(done_steps * 100 / total_steps))

    report = {
        "input": str(in_path),
        "outputs": outputs,
        "rows_total_filled": total_filled,
        "rows_per_file": int(params.rows_to_fill),
        "batch_count": int(params.batch_count),
        "brand_title_ru": params.brand_ru,
        "brand_desc_lat": params.brand_lat,
        "shape": params.shape,
        "lenses": params.lenses,
        "collection": params.collection,
        "holidays": params.holidays,
        "style": params.style,
        "seo_level": params.seo_level,
        "wb_safe_mode": params.wb_safe_mode,
        "wb_strict": params.wb_strict,
    }
    return outputs, total_filled, json.dumps(report, ensure_ascii=False, indent=2)
