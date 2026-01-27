from __future__ import annotations

import re
import random
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# ======================
# Config (v6 PRO)
# ======================

_WORD_RE = re.compile(r"[a-zа-яё0-9]+", re.IGNORECASE)

MAX_TITLE_LEN = 60
MAX_DESC_LEN = 2000
MIN_DESC_TARGET = 1150

VARIANTS = 12

START_PHRASES = ["Солнцезащитные очки", "Солнечные очки"]

# Не трогаем строки 1..N вообще
PROTECT_TOP_ROWS = 4

# Премиум-бренды: более “дорогая” подача, меньше инста/тренд-лексики
PREMIUM_BRANDS = {
    "gucci", "dior", "prada", "cazal", "tom ford", "tomford", "chanel", "cartier",
    "balenciaga", "saint laurent", "yves saint laurent", "ysl", "loewe", "givenchy",
    "versace", "bottega veneta", "fendi", "burberry", "dolce", "dolce&gabbana",
    "dolce gabbana", "dg", "ray-ban", "rayban", "oakley", "persol"
}

# Ключевые токены (для адаптивного SEO-tail)
CORE_KW = [
    "солнцезащитные очки",
    "солнечные очки",
    "очки солнцезащитные",
    "брендовые очки",
]


LENS_MAP = {
    "uv400": "UV400",
    "uv 400": "UV400",
    "поляризация": "поляризационные",
    "поляризационные": "поляризационные",
    "фотохром": "фотохромные",
    "фотохромные": "фотохромные",
    "хамелеон": "фотохромные",
    "хамелеоны": "фотохромные",
    "откидные линзы": "с откидными линзами",
    "откидные": "с откидными линзами",
    "зеркальные": "зеркальные",
    "зеркало": "зеркальные",
    "антибликовые": "антибликовые",
    "антиблик": "антибликовые",
    "градиентные": "градиентные",
    "градиент": "градиентные",
}

LENS_PRIORITY = [
    "UV400",
    "поляризационные",
    "фотохромные",
    "с откидными линзами",
    "зеркальные",
    "антибликовые",
    "градиентные",
]

SYN = {
    # Для title: в премиуме меньше “маркетплейсных” прилагательных
    "title_adjs_common": ["трендовые", "стильные", "брендовые", "модные", "актуальные"],
    "title_adjs_premium": ["брендовые", "дизайнерские", "премиальные"],

    "intro_hooks_common": [
        "трендовый аксессуар на лето",
        "актуальный акцент сезона",
        "выразительная деталь образа",
        "яркое дополнение к гардеробу",
    ],
    "intro_hooks_premium": [
        "акцент сезона",
        "выверенная деталь образа",
        "элегантный штрих к гардеробу",
        "сдержанный люкс в деталях",
    ],

    "luxury_phrases": [
        "продуманная эстетика деталей",
        "баланс трендов и классики",
        "внимание к деталям и пропорциям",
        "выразительная геометрия линий",
    ],

    "tone_common": [
        "Эти очки — не просто защита, а выразительный элемент стиля.",
        "Это аксессуар, который мгновенно собирает образ.",
        "Форма оправы подчёркивает черты лица и добавляет уверенности.",
    ],
    "tone_premium": [
        "Это аксессуар с узнаваемой эстетикой и выверенной посадкой.",
        "Лаконичная подача и точные пропорции делают модель универсальной.",
        "Аккуратный акцент, который не спорит с образом, а усиливает его.",
    ],

    "benefits_uv": [
        "надежно защищают глаза от ультрафиолета",
        "делают использование при ярком свете заметно комфортнее",
    ],
    "benefits_polar": [
        "уменьшают блики от воды и асфальта",
        "повышают контраст и читаемость изображения",
    ],
    "benefits_photo": [
        "адаптируются к уровню освещения",
        "подстраиваются под яркость солнца в течение дня",
    ],
    "benefits_flip": [
        "дают дополнительную функциональность",
        "удобны, когда нужен быстрый вариант переключения",
    ],

    "scenarios": ["город", "путешествия", "отпуск", "пляж", "прогулки", "вождение", "активный отдых"],

    # Триггеры — только в 3–4 из 12, и мягкие
    "triggers": ["дизайнерские", "эксклюзивные", "уникальные", "молодёжные"],

    "kw_style": [
        "брендовые очки", "солнечные очки", "солнцезащитные очки",
        "очки солнцезащитные", "стильные очки"
    ],
    "kw_social": ["тренд из instagram", "инста очки", "очки из tiktok"],
    "kw_shapes": [
        "очки cat eye", "очки кошачий глаз", "очки квадратные", "очки авиаторы",
        "очки круглые", "очки прямоугольные", "очки оверсайз"
    ],
    "kw_lens": [
        "очки UV400", "очки поляризационные", "очки фотохромные",
        "очки хамелеон", "очки с откидными линзами"
    ],
}


# ======================
# Helpers
# ======================

def _clean(s: Optional[str]) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def _low(s: str) -> str:
    return _clean(s).lower()


def _tokenize(text: str) -> set:
    return set(_WORD_RE.findall(text.lower()))


def _jaccard(a: str, b: str) -> float:
    ta = _tokenize(a)
    tb = _tokenize(b)
    if not ta and not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


def _split_features(raw: str) -> List[str]:
    raw = _clean(raw)
    if not raw:
        return []
    parts = re.split(r"[;,/]+", raw)
    return [_clean(p).lower() for p in parts if _clean(p)]


def _is_premium_brand(brand: str) -> bool:
    b = _low(brand)
    b = b.replace("&", " ").replace("-", " ")
    b = re.sub(r"\s+", " ", b).strip()
    return b in PREMIUM_BRANDS


def _normalize_lens_tokens(raw: str) -> List[str]:
    feats = _split_features(raw)
    out: List[str] = []
    for f in feats:
        if f in LENS_MAP:
            out.append(LENS_MAP[f])
            continue
        f2 = f.replace(" ", "")
        if "uv400" in f2:
            out.append("UV400")
        elif "поляр" in f:
            out.append("поляризационные")
        elif "фотохром" in f or "хамелеон" in f:
            out.append("фотохромные")
        elif "откид" in f:
            out.append("с откидными линзами")
        elif "зерк" in f:
            out.append("зеркальные")
        elif "антиблик" in f:
            out.append("антибликовые")
        elif "градиент" in f:
            out.append("градиентные")

    seen = set()
    uniq = []
    for x in out:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)

    pidx = {k: i for i, k in enumerate(LENS_PRIORITY)}
    uniq.sort(key=lambda x: pidx.get(x, 10_000))
    return uniq


def _fit_title(parts: List[str]) -> str:
    """
    Собирает заголовок <= 60 символов. Не режет слова: если не помещается — удаляем хвостовые блоки.
    """
    parts = [_clean(p) for p in parts if _clean(p)]
    while parts:
        s = " ".join(parts)
        if len(s) <= MAX_TITLE_LEN:
            return s
        parts.pop()
    return ""


def _norm_mid_case(s: str) -> str:
    """
    Нормальный регистр для середины текста (не ломаем UV400 / цифры / бренды).
    """
    s = _clean(s)
    if not s:
        return ""
    if re.fullmatch(r"[A-Z0-9\-]+", s):
        return s
    # если это уже “Cazal” / “Gucci” / “Prada” — оставляем как есть
    if s[:1].isupper() and (len(s) == 1 or s[1:].islower()):
        return s
    return s[:1].upper() + s[1:].lower()


def _shape_for_title(shape: str) -> str:
    """
    Нормализация формы ОПРАВЫ для названия: убираем “Квадрат”, делаем ожидаемую форму.
    """
    s = _low(shape)
    if "кошач" in s or "cat" in s:
        return "cat eye"
    if "авиат" in s:
        return "авиаторы"
    if "квад" in s:
        return "квадратные"
    if "круг" in s:
        return "круглые"
    if "прямоуг" in s:
        return "прямоугольные"
    if "овер" in s:
        return "оверсайз"
    # дефолт
    return _norm_mid_case(shape)


def _shape_keywords(shape: str) -> List[str]:
    s = _low(shape)
    if "кошач" in s or "cat" in s:
        return ["очки cat eye", "очки кошачий глаз"]
    if "авиат" in s:
        return ["очки авиаторы"]
    if "квад" in s:
        return ["очки квадратные"]
    if "круг" in s:
        return ["очки круглые"]
    if "прямоуг" in s:
        return ["очки прямоугольные"]
    if "овер" in s:
        return ["очки оверсайз"]
    return []


def _lens_keywords(lens_tokens: List[str]) -> List[str]:
    out = []
    lt = set(lens_tokens)
    if "UV400" in lt:
        out.append("очки UV400")
    if "поляризационные" in lt:
        out.append("очки поляризационные")
    if "фотохромные" in lt:
        out.append("очки фотохромные")
        out.append("очки хамелеон")
    if "с откидными линзами" in lt:
        out.append("очки с откидными линзами")
    return out


def _balanced_starts(rng: random.Random) -> List[str]:
    half = VARIANTS // 2
    starts = [START_PHRASES[0]] * half + [START_PHRASES[1]] * (VARIANTS - half)
    rng.shuffle(starts)
    return starts


def _build_seo_pack(brand: str, shape: str, lens_tokens: List[str], rng: random.Random, premium: bool) -> Dict[str, str]:
    brand = _clean(brand)
    base = [f"{brand} очки", f"очки {brand}"]

    shape_k = _shape_keywords(shape)
    lens_k = _lens_keywords(lens_tokens)

    pool_mid = SYN["kw_style"][:] + base + shape_k + lens_k
    rng.shuffle(pool_mid)

    mid = []
    seen = set()
    for x in pool_mid:
        k = _clean(x)
        if not k:
            continue
        kl = k.lower()
        if kl in seen:
            continue
        seen.add(kl)
        mid.append(k)
        if len(mid) >= 3:
            break

    pool_tail = SYN["kw_style"][:] + base + (shape_k or SYN["kw_shapes"][:]) + (lens_k or SYN["kw_lens"][:])
    if (not premium) and rng.random() < 0.45:
        pool_tail += SYN["kw_social"]  # соцключи — реже, и только не для премиума
    rng.shuffle(pool_tail)

    tail = []
    seen = set()
    for x in pool_tail:
        k = _clean(x)
        if not k:
            continue
        kl = k.lower()
        if kl in seen:
            continue
        seen.add(kl)
        tail.append(k)
        if len(tail) >= rng.randint(6, 8):
            break

    return {
        "kw1": mid[0] if len(mid) > 0 else "солнцезащитные очки",
        "kw2": mid[1] if len(mid) > 1 else "солнечные очки",
        "kw3": mid[2] if len(mid) > 2 else "брендовые очки",
        "tail": ", ".join(tail),
    }


def _spec_inline(brand: str, shape: str, lens_line: str, collection: str, scenarios: str) -> str:
    # Одна строка, твой формат
    return (
        f"Характеристики: "
        f"• Бренд: {brand} "
        f"• Форма оправы: {shape} "
        f"• Линзы: {lens_line} "
        f"• Коллекция: {collection} "
        f"• Сценарии: {scenarios}."
    )


def _pick_trigger_indices(rng: random.Random) -> set:
    k = rng.choice([3, 4])
    return set(rng.sample(range(VARIANTS), k))


def _lens_benefits(lens_tokens: List[str], rng: random.Random) -> str:
    parts = [rng.choice(SYN["benefits_uv"])]
    if "поляризационные" in lens_tokens:
        parts.append(rng.choice(SYN["benefits_polar"]))
    if "фотохромные" in lens_tokens:
        parts.append(rng.choice(SYN["benefits_photo"]))
    if "с откидными линзами" in lens_tokens:
        parts.append(rng.choice(SYN["benefits_flip"]))

    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return parts[0] + ", а также " + parts[1]
    return parts[0] + ", " + parts[1] + " и " + parts[2]


def _count_core_kw(text: str) -> int:
    t = text.lower()
    return sum(1 for kw in CORE_KW if kw in t)


def _should_add_search_tail(desc: str) -> bool:
    """
    Адаптивное SEO:
    - если уже есть ≥2 core-ключа и есть брендовые/линзовые слова — tail можно не добавлять
    - иначе добавляем tail
    """
    core = _count_core_kw(desc)
    t = desc.lower()
    has_lens = ("uv400" in t) or ("поляр" in t) or ("фотохром" in t) or ("хамелеон" in t)
    has_brand_kw = ("брендовые очки" in t) or ("очки " in t)
    # если текст уже “насыщен” — не перегружаем хвостом
    if core >= 2 and (has_lens or has_brand_kw):
        return False
    return True


# ======================
# Title scoring (PRO)
# ======================

def _title_score(title: str, brand: str, lens_tokens: List[str]) -> float:
    """
    Скоринг названия (для выбора “лучшего”).
    """
    t = title.strip()
    tl = t.lower()
    brand_l = brand.lower()

    score = 0.0

    # старт
    if t.startswith("Солнцезащитные очки"):
        score += 3.0
    if t.startswith("Солнечные очки"):
        score += 2.2

    # бренд
    if brand_l in tl:
        score += 3.0
        # бренд ближе к началу — лучше
        pos = tl.find(brand_l)
        if pos >= 0:
            score += max(0.0, 1.5 - (pos / 25.0))
    else:
        score -= 6.0  # без бренда — плохо

    # линзы
    lt = [x.lower() for x in lens_tokens]
    if "uv400" in tl:
        score += 1.8
    if "поляр" in tl or "поляризацион" in tl:
        score += 1.8 if ("поляризационные" in lt) else 1.1
    if "фотохром" in tl or "хамелеон" in tl:
        score += 1.6 if ("фотохромные" in lt) else 1.0

    # длина (WB любит вменяемый size)
    L = len(t)
    if 48 <= L <= 58:
        score += 2.2
    elif 42 <= L <= 60:
        score += 1.2
    else:
        score -= 0.8

    # мусор-слова
    if "модель" in tl:
        score -= 1.0
    if "2025" in tl or "2026" in tl:
        score -= 0.4  # год можно, но лучше не злоупотреблять

    # лишнее “очки очки”
    if "очки очки" in tl:
        score -= 1.5

    return score


def _rank_titles(titles: List[str], brand: str, lens_tokens: List[str]) -> List[Tuple[str, float]]:
    ranked = [(t, _title_score(t, brand, lens_tokens)) for t in titles]
    ranked.sort(key=lambda x: x[1], reverse=True)
    return ranked


# ======================
# Title generation
# ======================

def _gen_titles_12(brand: str, shape: str, lens_raw: str, collection: str, rng: random.Random, premium: bool) -> List[str]:
    brand = _clean(brand)
    collection = _clean(collection)

    lens_tokens = _normalize_lens_tokens(lens_raw)
    lens_main = lens_tokens[0] if lens_tokens else "UV400"

    sh_title = _shape_for_title(shape)
    starts = _balanced_starts(rng)

    adjs = SYN["title_adjs_premium"] if premium else SYN["title_adjs_common"]
    adj = rng.choice(adjs)

    # Важный момент: делаем названия “внятными”, без случайных кусков
    patterns = [
        lambda s: [s, brand, lens_main],
        lambda s: [s, brand, lens_main, sh_title],
        lambda s: [s, brand, sh_title, lens_main],
        lambda s: [s, adj, brand, lens_main],
        lambda s: [s, adj, brand, sh_title],
        lambda s: [s, brand, sh_title],
        lambda s: [s, brand, lens_main, collection],
        lambda s: [s, brand, sh_title, collection],
        lambda s: [s, brand, lens_main, "очки", sh_title],
        lambda s: [s, brand, "очки", sh_title, lens_main],
        lambda s: [s, adj, brand, lens_main, sh_title],
        lambda s: [s, adj, brand, sh_title, lens_main],
    ]

    titles: List[str] = []
    for i in range(VARIANTS):
        t = _fit_title(patterns[i](starts[i]))
        # гарантируем: start + brand
        if not t.startswith(starts[i]) or brand.lower() not in t.lower():
            t = _fit_title([starts[i], brand, lens_main, sh_title])
        titles.append(t)

    # убрать точные дубли
    seen = set()
    out = []
    for t in titles:
        if t not in seen:
            seen.add(t)
            out.append(t)
        else:
            alt_start = START_PHRASES[0] if t.startswith(START_PHRASES[1]) else START_PHRASES[1]
            t2 = _fit_title([alt_start, brand, lens_main, sh_title])
            out.append(t2 if t2 and t2 not in seen else t)
            seen.add(out[-1])

    return out[:VARIANTS]


# ======================
# Description generation (one line)
# ======================

def _build_desc_one_line(
    brand: str,
    shape: str,
    lens_raw: str,
    collection: str,
    seed: int,
    add_trigger: bool,
    premium: bool,
    allow_tail: bool,
) -> str:
    rng = random.Random(seed)

    brand = _clean(brand)
    shape_mid = _norm_mid_case(shape)
    collection = _clean(collection) or "Весна–Лето 2025–2026"

    lens_tokens = _normalize_lens_tokens(lens_raw)
    lens_line = ", ".join(lens_tokens[:3]) if lens_tokens else "UV400"

    scenarios_list = rng.sample(SYN["scenarios"], k=4)
    scenarios = ", ".join(scenarios_list)

    hook = rng.choice(SYN["intro_hooks_premium"] if premium else SYN["intro_hooks_common"])
    luxury = rng.choice(SYN["luxury_phrases"])
    tone_line = rng.choice(SYN["tone_premium"] if premium else SYN["tone_common"])
    benefits = _lens_benefits(lens_tokens, rng)

    seo = _build_seo_pack(brand, shape_mid, lens_tokens, rng, premium=premium)

    trigger_sentence = ""
    if add_trigger and (not premium):
        trig = rng.choice(SYN["triggers"])
        trigger_sentence = rng.choice([
            f"Очки {brand} — {trig} акцент на сезон: заметные детали и уверенная подача.",
            f"{brand} часто выбирают как {trig} аксессуар — он работает и в кадре, и в жизни.",
        ])

    s1 = (
        f"Очки {brand} — {hook} {collection}. {tone_line} "
        f"Если вы ищете {seo['kw1']}, обратите внимание на эту модель: она выглядит современно и уместно в разных образах."
    )
    s2 = (
        f"Дизайн {brand} — это {luxury}: выразительные линии, гармоничные пропорции и внимание к деталям; "
        f"такие {seo['kw2']} легко становятся главным акцентом в гардеробе."
    )
    s3 = (
        f"Форма оправы {shape_mid} подчёркивает индивидуальность и балансирует образ; "
        f"по восприятию это {seo['kw3']}, которые «собирают» стилизацию."
    )
    spec = _spec_inline(brand, shape_mid, lens_line, collection, scenarios)
    s4 = (
        f"Линзы: {lens_line}. Они {benefits}, поэтому это не только модные, но и действительно комфортные солнцезащитные очки на каждый день."
    )
    s5 = (
        f"Сценарии использования: {scenarios}; модель удобна в городе и в поездках, на прогулках и за рулём — когда важны комфорт и защита зрения."
    )
    s6 = (
        f"Выбирайте солнечные очки {brand} как подарок или как обновление гардероба к тёплому сезону: стильно, удобно и функционально."
    )

    parts = [s1]
    if trigger_sentence:
        parts.append(trigger_sentence)
    parts += [s2, s3, spec, s4, s5, s6]

    text = " ".join(_clean(x) for x in parts if _clean(x))
    text = _clean(text)

    # Tail добавляем адаптивно (и только если allow_tail)
    if allow_tail and _should_add_search_tail(text):
        text = f"{text} Поисковые запросы: {seo['tail']}."
        text = _clean(text)

    # добивка длины
    if len(text) < MIN_DESC_TARGET:
        extra_pool = [
            f"Цветовые решения и детали {brand} помогают подчеркнуть характер образа — от лаконичных сочетаний до ярких акцентов.",
            f"Если важны бренд, форма и защита, такие солнцезащитные очки выглядят уместно и в городе, и на отдыхе.",
            f"Это удачный выбор для тех, кто ищет брендовые очки с понятной посадкой и современным силуэтом.",
        ]
        if premium:
            extra_pool = [
                f"Сдержанная подача и узнаваемая эстетика {brand} делают аксессуар универсальным для гардероба.",
                f"Точные пропорции и аккуратные детали подчёркивают стиль без лишней демонстративности.",
            ] + extra_pool[:1]
        rng.shuffle(extra_pool)
        for e in extra_pool:
            cand = text + " " + e
            if len(cand) <= MAX_DESC_LEN:
                text = cand
            if len(text) >= MIN_DESC_TARGET:
                break

    # лимит 2000
    if len(text) > MAX_DESC_LEN:
        text = text[:MAX_DESC_LEN]
        text = text.rsplit(" ", 1)[0].rstrip(".,;:") + "."

    return text


def _generate_12_pairs(
    brand: str,
    shape: str,
    lens_raw: str,
    collection: str,
    seed: int,
    allow_tail: bool,
) -> Tuple[List[Tuple[str, str]], List[Tuple[str, float]]]:
    """
    Возвращаем:
    - pairs: 12 (title, desc) с сильной уникализацией
    - ranked_titles: рейтинг названий (для логов/выбора)
    """
    premium = _is_premium_brand(brand)
    rng = random.Random(seed)

    lens_tokens = _normalize_lens_tokens(lens_raw)
    titles = _gen_titles_12(brand, shape, lens_raw, collection, rng, premium=premium)
    ranked = _rank_titles(titles, brand=brand, lens_tokens=lens_tokens)

    trigger_idx = _pick_trigger_indices(rng)

    pairs: List[Tuple[str, str]] = []
    descs: List[str] = []

    for i in range(VARIANTS):
        add_trigger = i in trigger_idx

        best_desc = None
        for attempt in range(1, 9):
            desc = _build_desc_one_line(
                brand=brand,
                shape=shape,
                lens_raw=lens_raw,
                collection=collection,
                seed=seed * 1000 + i * 97 + attempt * 10007,
                add_trigger=add_trigger,
                premium=premium,
                allow_tail=allow_tail,
            )
            if not descs or all(_jaccard(desc, prev) <= 0.58 for prev in descs):
                best_desc = desc
                break
        if best_desc is None:
            best_desc = desc

        descs.append(best_desc)
        pairs.append((titles[i], best_desc))

    return pairs, ranked


# ======================
# Excel / WB
# ======================

def _find_header_row_and_cols(
    ws,
    needed: List[str],
    max_rows: int = 30,
    max_cols: int = 600
) -> Tuple[int, Dict[str, int]]:
    for r in range(1, max_rows + 1):
        found: Dict[str, int] = {}
        for c in range(1, max_cols + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                v = v.strip()
                if v in needed:
                    found[v] = c
        if all(k in found for k in needed):
            return r, found
    raise RuntimeError(f"Не найдены колонки: {', '.join(needed)}")


def _try_find_col(ws, header: str, header_row: int, max_cols: int = 600) -> Optional[int]:
    for c in range(1, max_cols + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip() == header:
            return c
    return None


def _row_is_product(ws, r: int, signal_cols: List[int]) -> bool:
    for c in signal_cols:
        v = ws.cell(r, c).value
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        return True
    return False


def _write_safe(ws, row: int, col: int, value: str) -> None:
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                ws.cell(mr.min_row, mr.min_col).value = value
                return
        return
    cell.value = value


def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str,
    include_search_tail: bool = True,   # “разрешить хвост” (фактически адаптивный)
    overwrite_existing: bool = True,
) -> Tuple[str, int]:
    """
    v6 PRO:
    - не трогаем строки 1..PROTECT_TOP_ROWS
    - генерим 12 вариантов
    - для каждой строки выбираем “лучшее” название из топа, чтобы названия не были одинаковыми
      (идём по top-рангу и используем разные из топ-5)
    - описание берём из соответствующего варианта (с уникализацией)
    - создаём лог рядом с выходным файлом
    """
    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, cols = _find_header_row_and_cols(ws, ["Наименование", "Описание"])
    col_name = cols["Наименование"]
    col_desc = cols["Описание"]

    signal_cols = [col_name, col_desc]
    for h in ["Фото", "Артикул продавца", "Баркоды", "Размер", "Цвет"]:
        c = _try_find_col(ws, h, header_row)
        if c:
            signal_cols.append(c)
    if len(signal_cols) <= 2:
        signal_cols = list(range(1, 30))

    pairs12, ranked_titles = _generate_12_pairs(
        brand=_clean(brand),
        shape=_clean(shape),
        lens_raw=_clean(lens_features),
        collection=_clean(collection) or "Весна–Лето 2025–2026",
        seed=1234567,
        allow_tail=include_search_tail,
    )

    # Индекс “топовых” названий -> их позиции в исходных 12
    title_to_idx = {pairs12[i][0]: i for i in range(len(pairs12))}

    # Топ-5 названий — будем ротировать по строкам, чтобы не было одинаковых
    top_titles = [t for t, _s in ranked_titles[:5] if t in title_to_idx]
    if not top_titles:
        top_titles = [pairs12[0][0]]

    filled = 0
    counter = 0

    start_row = max(header_row + 1, PROTECT_TOP_ROWS + 1)

    # лог
    in_path = Path(input_xlsx)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = in_path.parent / f"{in_path.stem}_filled_{ts}.xlsx"
    log_path = in_path.parent / f"{in_path.stem}_filled_{ts}.log.txt"

    log_lines = []
    log_lines.append("v6 PRO log")
    log_lines.append(f"Input: {in_path.name}")
    log_lines.append(f"Brand: {brand} | Shape: {shape} | Lens: {lens_features} | Collection: {collection}")
    log_lines.append("Ranked titles (top 5):")
    for i, (t, sc) in enumerate(ranked_titles[:5], 1):
        log_lines.append(f"{i}. {t}  [score={sc:.2f}]")

    for r in range(start_row, ws.max_row + 1):
        if not _row_is_product(ws, r, signal_cols):
            continue

        if not overwrite_existing:
            existing_title = ws.cell(r, col_name).value
            existing_desc = ws.cell(r, col_desc).value
            if _clean(existing_title) or _clean(existing_desc):
                continue

        # выбираем название из топ-5 по кругу, но описание — из соответствующего варианта
        chosen_title = top_titles[counter % len(top_titles)]
        idx = title_to_idx.get(chosen_title, counter % VARIANTS)

        title, desc = pairs12[idx]

        _write_safe(ws, r, col_name, title)
        _write_safe(ws, r, col_desc, desc)

        log_lines.append(f"Row {r}: title_variant={idx+1} title='{title}'")
        filled += 1
        counter += 1

    wb.save(out_path)

    try:
        log_path.write_text("\n".join(log_lines), encoding="utf-8")
    except Exception:
        # лог не критичен
        pass

    return str(out_path), filled
