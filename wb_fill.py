# wb_fill.py
from __future__ import annotations

import json
import os
import random
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import openpyxl

TITLE_MAX = 60
DESC_MAX = 1000

DEFAULT_BRANDS_RU = {
    "Balenciaga": "Balenciaga",
    "Gucci": "Gucci",
    "Prada": "Prada",
    "Ray-Ban": "Ray-Ban",
    "Dior": "Dior",
    "Versace": "Versace",
}

DEFAULT_SHAPES = [
    "Вайфаеры", "Авиаторы", "Кошачий глаз", "Квадратные", "Круглые", "Овальные"
]

DEFAULT_LENSES = [
    "Поляризационные", "Градиентные", "Зеркальные", "Фотохромные", "УФ400"
]

SLOGANS = [
    "Лаконичный дизайн", "Уверенный образ", "Комфорт на каждый день",
    "Акцент на деталях", "Современная эстетика", "Городской стиль",
]

SAFE_REPLACE = {
    "лечит": "помогает",
    "гарантия": "поддержка",
    "100%": "",
    "абсолютно": "",
    "лучший": "отличный",
    "идеальный": "удачный",
    "навсегда": "надолго",
    "никогда": "редко",
}

STRICT_PATTERNS = [
    (re.compile(r"\bсам(ый|ая|ое|ые)\b", re.I), ""),
    (re.compile(r"\bгарант(ия|ируем|ировано)\b", re.I), "поддержка"),
    (re.compile(r"\b100%\b"), ""),
    (re.compile(r"\bбезупречн(ый|ая|ое|ые)\b", re.I), "аккуратный"),
]


def _read_json(path: str) -> Optional[object]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def load_brands_ru_map(data_dir: str) -> Dict[str, str]:
    # ожидаем либо brands.json как dict {"Balenciaga":"Балenciaga"} или list объектов
    if data_dir:
        p = os.path.join(data_dir, "brands.json")
        obj = _read_json(p)
        if isinstance(obj, dict):
            return {str(k): str(v) for k, v in obj.items()}
        if isinstance(obj, list):
            m = {}
            for it in obj:
                if isinstance(it, dict):
                    lat = it.get("lat") or it.get("en") or it.get("brand") or it.get("key")
                    ru = it.get("ru") or it.get("name") or it.get("value")
                    if lat and ru:
                        m[str(lat)] = str(ru)
            if m:
                return m
    return dict(DEFAULT_BRANDS_RU)


def load_list(data_dir: str, filename: str, fallback: List[str]) -> List[str]:
    if data_dir:
        p = os.path.join(data_dir, filename)
        obj = _read_json(p)
        if isinstance(obj, list):
            out = []
            for it in obj:
                if isinstance(it, str):
                    out.append(it)
                elif isinstance(it, dict):
                    # поддержка форматов {"name": "..."} / {"ru": "..."} / {"value":"..."}
                    v = it.get("ru") or it.get("name") or it.get("value")
                    if v:
                        out.append(str(v))
            if out:
                return out
    return list(fallback)


def apply_safe(text: str) -> str:
    t = text
    for bad, good in SAFE_REPLACE.items():
        t = re.sub(re.escape(bad), good, t, flags=re.I)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def apply_strict(text: str) -> str:
    t = text
    for rx, repl in STRICT_PATTERNS:
        t = rx.sub(repl, t)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def clamp(s: str, max_len: int) -> str:
    s = s.strip()
    if len(s) <= max_len:
        return s
    # аккуратно режем по слову
    cut = s[: max_len]
    if " " in cut:
        cut = cut.rsplit(" ", 1)[0]
    return cut.strip()


def tokenize(text: str) -> List[str]:
    text = re.sub(r"[^a-zA-Zа-яА-Я0-9\s-]", " ", text)
    text = re.sub(r"\s{2,}", " ", text).lower().strip()
    return [t for t in text.split() if len(t) > 2]


def jaccard(a: str, b: str) -> float:
    sa = set(tokenize(a))
    sb = set(tokenize(b))
    if not sa or not sb:
        return 0.0
    inter = len(sa & sb)
    union = len(sa | sb)
    return inter / union if union else 0.0


def generate_title(
    brand_lat: str,
    shape: str,
    lens: str,
    brand_map: Dict[str, str],
    slogan_pool: List[str],
) -> str:
    brand_ru = brand_map.get(brand_lat, brand_lat)
    bits = [
        f"Солнцезащитные очки {brand_ru}",
        shape,
        lens,
        random.choice(slogan_pool) if slogan_pool else "",
    ]
    s = " • ".join([b for b in bits if b]).strip(" •")
    return clamp(s, TITLE_MAX)


def _gender_phrase(gender_mode: str) -> str:
    gm = (gender_mode or "").lower()
    if gm in ("male", "m", "men", "муж", "мужской"):
        return "мужские"
    if gm in ("female", "f", "women", "жен", "женский"):
        return "женские"
    return "унисекс"


def _seo_density_hint(seo_level: str) -> str:
    sl = (seo_level or "normal").lower()
    if sl == "low":
        return "без перегруза ключами"
    if sl == "high":
        return "с усиленным SEO-упоминанием"
    return "с естественным SEO"


def _len_hint(length_mode: str) -> str:
    lm = (length_mode or "medium").lower()
    if lm == "short":
        return "короткое описание"
    if lm == "long":
        return "подробное описание"
    return "средняя длина"


def _style_hint(style_mode: str) -> str:
    sm = (style_mode or "premium").lower()
    if sm == "basic":
        return "простым языком"
    if sm == "sport":
        return "в спортивном тоне"
    return "в премиальном тоне"


def build_description(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
    length_mode: str,
    style_mode: str,
) -> str:
    g = _gender_phrase(gender_mode)
    seo = _seo_density_hint(seo_level)
    ln = _len_hint(length_mode)
    st = _style_hint(style_mode)

    lines = [
        f"{g.capitalize()} солнцезащитные очки {brand_lat} — {shape.lower()} с линзами: {lens.lower()}.",
        f"Коллекция: {collection}.",
        f"Описание {st}, {ln}, {seo}.",
        f"Удобная посадка, аккуратные материалы, комфорт для города и путешествий.",
        f"Линзы помогают снизить блики и повысить визуальный комфорт в яркий день.",
        f"Подойдут к повседневному и деловому образу: лаконично, современно, уместно.",
    ]

    # немного вариативности
    extras = [
        "Защита от яркого света и продуманная форма оправы.",
        "Сбалансированный дизайн для разных типов лица.",
        "Стильно сочетаются с классическими и casual-образами.",
        "Практичный аксессуар на сезон и не только.",
    ]
    random.shuffle(extras)

    lm = (length_mode or "medium").lower()
    if lm == "short":
        lines = lines[:4]
    elif lm == "long":
        lines.extend(extras[:2])
    else:
        lines.append(extras[0])

    desc = " ".join(lines)
    return clamp(desc, DESC_MAX)


def generate_description_best_of(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
    length_mode: str,
    style_mode: str,
    used_desc: List[str],
    best_of: int = 8,
    uniq_strength: int = 3,
) -> Tuple[str, float]:
    """
    Возвращает (описание, score), где score — максимальный Jaccard с used_desc (чем меньше, тем лучше).
    uniq_strength 1..5: чем больше, тем сильнее штрафуем похожие варианты.
    """
    best_desc = ""
    best_score = 10.0

    for _ in range(max(2, best_of)):
        d = build_description(
            brand_lat=brand_lat,
            shape=shape,
            lens=lens,
            collection=collection,
            seo_level=seo_level,
            gender_mode=gender_mode,
            length_mode=length_mode,
            style_mode=style_mode,
        )

        score = 0.0
        for u in used_desc[-50:]:
            score = max(score, jaccard(d, u))

        # усиливаем "нелюбовь" к похожести
        score = score * (1.0 + (uniq_strength - 1) * 0.25)

        if score < best_score:
            best_score = score
            best_desc = d

    return best_desc, best_score


def find_header_col(ws, candidates: set, header_scan_rows: int = 25):
    candidates = {c.strip().lower() for c in candidates}
    for r in range(1, header_scan_rows + 1):
        for cell in ws[r]:
            if cell.value is None:
                continue
            val = str(cell.value).strip().lower()
            if val in candidates:
                return cell.column, r
            if any(c in val for c in candidates):
                return cell.column, r
    return None, None


@dataclass
class FillParams:
    brand: str
    shape: str
    lens: str
    collection: str
    seo_density: str = "normal"
    length_mode: str = "medium"
    style_mode: str = "premium"
    gender_mode: str = "auto"
    safe_mode: bool = True
    strict_mode: bool = True
    data_dir: str = ""
    seed: Optional[int] = None
    rows_to_fill: int = 6
    fill_only_empty: bool = True
    uniq_strength: int = 3


def fill_wb_template(
    in_path: str,
    out_path: str,
    params: FillParams,
) -> Dict[str, object]:
    """
    Заполняет XLSX: наименование + описание. Заполняет rows_to_fill строк.
    Возвращает небольшой отчёт.
    """
    if params.seed is not None:
        random.seed(params.seed)
    else:
        random.seed()

    wb = openpyxl.load_workbook(in_path)
    ws = wb.active

    # Находим колонки
    title_candidates = {"наименование", "название", "name", "title", "наименование товара"}
    desc_candidates = {"описание", "description", "описание товара", "текст"}

    title_col, header_row = find_header_col(ws, title_candidates)
    desc_col, header_row2 = find_header_col(ws, desc_candidates)

    if not header_row and header_row2:
        header_row = header_row2
    if header_row and not header_row2:
        header_row2 = header_row

    if not title_col or not desc_col or not header_row:
        raise RuntimeError("Не удалось найти колонки 'Наименование/Описание' в XLSX (заголовки).")

    data_dir = params.data_dir or ""
    brand_map = load_brands_ru_map(data_dir)
    slogan_pool = SLOGANS[:]
    random.shuffle(slogan_pool)

    used_desc: List[str] = []
    max_scores: List[float] = []

    start_row = header_row + 1
    end_row = start_row + max(1, params.rows_to_fill) - 1

    for r in range(start_row, end_row + 1):
        title_cell = ws.cell(row=r, column=title_col)
        desc_cell = ws.cell(row=r, column=desc_col)

        if params.fill_only_empty:
            if title_cell.value and str(title_cell.value).strip():
                # если заголовок уже есть — можно всё равно обновить описание или пропустить.
                pass
            if desc_cell.value and str(desc_cell.value).strip():
                continue

        title = generate_title(params.brand, params.shape, params.lens, brand_map, slogan_pool)

        desc, score = generate_description_best_of(
            brand_lat=params.brand,
            shape=params.shape,
            lens=params.lens,
            collection=params.collection,
            seo_level=params.seo_density,
            gender_mode=params.gender_mode,
            length_mode=params.length_mode,
            style_mode=params.style_mode,
            used_desc=used_desc,
            best_of=8,
            uniq_strength=params.uniq_strength,
        )

        if params.safe_mode:
            title = apply_safe(title)
            desc = apply_safe(desc)
        if params.strict_mode:
            title = apply_strict(title)
            desc = apply_strict(desc)

        title = clamp(title, TITLE_MAX)
        desc = clamp(desc, DESC_MAX)

        title_cell.value = title
        desc_cell.value = desc

        used_desc.append(desc)
        max_scores.append(score)

    wb.save(out_path)

    avg = (sum(max_scores) / len(max_scores)) if max_scores else 0.0
    return {
        "filled_rows": len(max_scores),
        "avg_max_jaccard": round(avg, 4),
        "out_path": out_path,
    }


def generate_preview(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_density: str = "normal",
    length_mode: str = "medium",
    style_mode: str = "premium",
    gender_mode: str = "auto",
    safe_mode: bool = True,
    strict_mode: bool = True,
    data_dir: str = "",
    used_desc: Optional[List[str]] = None,
    uniq_strength: int = 3,
) -> Tuple[str, str]:
    """
    Для live-preview: возвращает (title, description).
    """
    if used_desc is None:
        used_desc = []

    brand_map = load_brands_ru_map(data_dir or "")
    slogan_pool = SLOGANS[:]
    random.shuffle(slogan_pool)

    title = generate_title(brand_lat, shape, lens, brand_map, slogan_pool)
    desc, _ = generate_description_best_of(
        brand_lat=brand_lat,
        shape=shape,
        lens=lens,
        collection=collection,
        seo_level=seo_density,
        gender_mode=gender_mode,
        length_mode=length_mode,
        style_mode=style_mode,
        used_desc=used_desc,
        best_of=8,
        uniq_strength=uniq_strength,
    )

    if safe_mode:
        title = apply_safe(title)
        desc = apply_safe(desc)
    if strict_mode:
        title = apply_strict(title)
        desc = apply_strict(desc)

    return clamp(title, TITLE_MAX), clamp(desc, DESC_MAX)
