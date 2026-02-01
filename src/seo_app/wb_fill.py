from __future__ import annotations

import json
import re
import time
import random
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, List, Callable, Dict

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


MAX_TITLE_LEN = 60
MAX_DESC_LEN = 2000
TARGET_DESC_MIN = 1050
PROTECT_TOP_ROWS = 4

SIMILARITY_THRESHOLD = 0.72
REGEN_ATTEMPTS = 8


# ======================
# Helpers
# ======================

def _clean(s: Optional[str]) -> str:
    return re.sub(r"\s+", " ", str(s)).strip() if s else ""


def _tokenize(text: str) -> set:
    return set(re.findall(r"[a-zа-яё0-9]+", text.lower(), flags=re.IGNORECASE))


def _jaccard(a: str, b: str) -> float:
    A = _tokenize(a)
    B = _tokenize(b)
    if not A or not B:
        return 0.0
    return len(A & B) / len(A | B)


def _norm_shape(shape: str) -> str:
    s = _clean(shape).lower()
    if "квад" in s:
        return "квадратная"
    if "круг" in s:
        return "круглая"
    if "кошач" in s or "cat" in s:
        return "cat eye"
    if "авиат" in s:
        return "авиаторы"
    if "прямоуг" in s:
        return "прямоугольная"
    if "овер" in s:
        return "оверсайз"
    return _clean(shape) or "универсальная"


def _shape_for_title(shape: str) -> str:
    s = _norm_shape(shape)
    mapping = {
        "квадратная": "квадратные",
        "круглая": "круглые",
        "прямоугольная": "прямоугольные",
        "авиаторы": "авиаторы",
        "cat eye": "cat eye",
        "оверсайз": "оверсайз",
        "универсальная": "универсальные",
    }
    return mapping.get(s, s)


def _normalize_lens(raw: str) -> str:
    r = _clean(raw)
    if not r:
        return "UV400"

    parts = [p.strip() for p in re.split(r"[;,/]+", r) if p.strip()]
    out: List[str] = []

    for p in parts:
        pl = p.lower().replace(" ", "")
        if "uv400" in pl:
            out.append("UV400")
        elif "поляр" in p.lower():
            out.append("поляризационные")
        elif "фотохром" in p.lower() or "хамелеон" in p.lower():
            out.append("фотохромные")
        elif "откид" in p.lower():
            out.append("с откидными линзами")
        elif "зерк" in p.lower():
            out.append("зеркальные")
        elif "антиблик" in p.lower():
            out.append("антибликовые")
        elif "градиент" in p.lower():
            out.append("градиентные")
        elif "replica" in p.lower() or "реплик" in p.lower():
            out.append("replica")
        else:
            out.append(p.strip())

    seen = set()
    uniq = []
    for x in out:
        xl = x.lower()
        if xl not in seen:
            seen.add(xl)
            uniq.append(x)

    return ", ".join(uniq) if uniq else "UV400"


def _pick_feat_for_title(lens_features: str, shape: str) -> str:
    ln = _normalize_lens(lens_features).lower()
    if "replica" in ln:
        return "replica"
    if "uv400" in ln:
        return "UV400"
    if "поляр" in ln:
        return "поляризационные"
    if "фотохром" in ln or "хамелеон" in ln:
        return "хамелеон"
    if "откид" in ln:
        return "откидные"
    return _shape_for_title(shape)


def _fit_title_tokens(tokens: List[str]) -> str:
    tokens = [_clean(t) for t in tokens if _clean(t)]
    out = ""
    for t in tokens:
        test = (out + " " + t).strip()
        if len(test) <= MAX_TITLE_LEN:
            out = test
        else:
            break
    return out if out else "Солнцезащитные очки"


# ======================
# Brand Profiles
# ======================

def _profiles_path() -> Path:
    base_dir = Path(__file__).resolve().parent.parent.parent  # repo root
    return base_dir / "data" / "brand_profiles.json"


def load_brand_profiles() -> Dict:
    p = _profiles_path()
    if not p.exists():
        # если файла нет — не падаем, работаем дефолтом
        return {"default_profile": "classic", "profiles": {}, "brand_to_profile": {}}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {"default_profile": "classic", "profiles": {}, "brand_to_profile": {}}


def pick_profile(brand: str, profiles_data: Dict) -> Dict:
    brand = _clean(brand)
    default_key = profiles_data.get("default_profile", "classic")
    brand_map = profiles_data.get("brand_to_profile", {})
    profiles = profiles_data.get("profiles", {})

    key = brand_map.get(brand, default_key)

    prof = profiles.get(key, {})
    # дефолтные безопасные поля
    return {
        "key": key,
        "tone": prof.get("tone", "neutral"),
        "allow_words": prof.get("allow_words", []),
        "avoid_words": prof.get("avoid_words", []),
        "hooks": prof.get("hooks", []),
        "style_lines": prof.get("style_lines", []),
        "closings": prof.get("closings", []),
    }


def _filter_avoid(text: str, avoid_words: List[str]) -> str:
    # мягкая чистка: выкидываем “плохие” слова/фразы, чтобы не палилось
    if not avoid_words:
        return text
    out = text
    for w in avoid_words:
        if not w:
            continue
        out = re.sub(re.escape(w), "", out, flags=re.IGNORECASE)
    return _clean(out)


# ======================
# Title slogans (deck, no repeats per run)
# ======================

SLOGANS = [
    "Красивые", "Стильные", "Модные", "Крутые", "Молодёжные",
    "Трендовые", "Актуальные", "Эффектные", "Элегантные",
    "Лаконичные", "Дизайнерские", "Премиальные", "Современные",
    "Универсальные", "Практичные", "Лёгкие", "Комфортные",
    "Статусные", "Выразительные", "Минималистичные",
    "Городские", "Повседневные", "Сезонные", "Функциональные",
    "Изысканные", "Яркие", "Сдержанные", "Удобные",
    "Брендовые", "Коллекционные"
]

PRODUCT_PHRASES = ["солнцезащитные очки", "солнечные очки"]


class SloganDeck:
    def __init__(self, seed: int):
        self.rng = random.Random(seed)
        self.deck = SLOGANS[:]
        self.rng.shuffle(self.deck)
        self.i = 0

    def next(self) -> str:
        if self.i >= len(self.deck):
            self.rng.shuffle(self.deck)
            self.i = 0
        v = self.deck[self.i]
        self.i += 1
        return v


def build_title(deck: SloganDeck, brand: str, shape: str, lens_features: str, rng: random.Random) -> str:
    slogan = deck.next()
    product = rng.choice(PRODUCT_PHRASES)
    feat = _pick_feat_for_title(lens_features, shape)
    shape_tail = _shape_for_title(shape)

    tail_options = [[feat], [shape_tail], [feat, shape_tail]]
    rng.shuffle(tail_options)

    candidates = []
    for tail in tail_options:
        candidates.append(_fit_title_tokens([slogan, product, brand] + tail))

    def score(t: str) -> float:
        L = len(t)
        s = 0.0
        s += 3.0 if any(t.startswith(x + " ") for x in SLOGANS) else -100.0
        s += 2.0 if ("солнцезащитные очки" in t or "солнечные очки" in t) else -10.0
        s += 2.0 if _clean(brand).lower() in t.lower() else -10.0
        s += 1.4 if any(k in t.lower() for k in ["uv400", "replica", "поляр", "хамелеон", "откид", "квадрат", "кругл", "авиатор", "cat eye"]) else 0.0
        s += 2.0 if 48 <= L <= 60 else (1.0 if 40 <= L <= 60 else 0.0)
        return s

    candidates.sort(key=score, reverse=True)
    return candidates[0]


# ======================
# Description (brand-profile-driven, human-like)
# ======================

def _scenarios(rng: random.Random) -> str:
    picks = rng.sample(
        ["город", "путешествия", "отпуск", "пляж", "прогулки", "вождение", "активный отдых", "летние выходы"],
        k=4
    )
    return ", ".join(picks)


def _lens_benefit(lens_norm: str, rng: random.Random) -> str:
    ln = lens_norm.lower()
    variants = []

    if "поляр" in ln:
        variants += [
            "Поляризация помогает убрать блики — особенно приятно у воды и за рулём.",
            "С поляризационными линзами меньше бликов и визуально комфортнее в яркий день."
        ]
    if "uv400" in ln:
        variants += [
            "UV400 — базовая защита для солнечных дней: глазам комфортнее на улице.",
            "UV400 помогает меньше щуриться и спокойнее переносить яркий свет."
        ]
    if "фотохром" in ln or "хамелеон" in ln:
        variants += [
            "Фотохромные линзы подстраиваются под освещение — удобно, когда день меняется.",
            "Линзы-хамелеон реагируют на свет: на улице темнее, в тени мягче."
        ]
    if "откид" in ln:
        variants += [
            "Откидные линзы — быстро переключаешься под ситуацию прямо на ходу.",
            "Фишка с откидными линзами удобна, когда нужен другой режим за секунду."
        ]

    if not variants:
        variants = [
            "Линзы подобраны так, чтобы в солнечную погоду было комфортно и в городе, и в поездках.",
            "Комфорт в яркий день — главное: меньше напряжения для глаз."
        ]
    return rng.choice(variants)


def _shape_line(shape_norm: str, rng: random.Random) -> str:
    sn = shape_norm.lower()
    variants = []

    if "квадрат" in sn:
        variants += [
            "Квадратная форма смотрится собранно и добавляет образу структуру.",
            "Квадратная оправа делает образ более “чётким” и современным."
        ]
    if "кругл" in sn:
        variants += [
            "Круглая оправа смотрится мягче и добавляет лёгкости в образ.",
            "Круглая форма даёт лёгкий винтажный акцент."
        ]
    if "cat eye" in sn:
        variants += [
            "Форма cat eye добавляет женственный акцент и визуально “поднимает” образ.",
            "Cat eye — аккуратный вау-эффект без лишнего шума."
        ]
    if "авиат" in sn:
        variants += [
            "Авиаторы — классика, которая уместна и в городе, и в отпуске.",
            "Форма “авиаторы” обычно идёт почти всем — универсальная история."
        ]
    if not variants:
        variants = [
            f"Форма оправы {shape_norm} выглядит актуально и легко сочетается с базовым гардеробом.",
            f"Силуэт {shape_norm} спокойный и универсальный — для ежедневной носки."
        ]
    return rng.choice(variants)


def _seo_sentence(brand: str, rng: random.Random) -> str:
    return rng.choice([
        f"Если ищешь солнцезащитные очки {brand} на каждый день — это как раз про удобство и стиль.",
        "Внутри смысла здесь естественно звучат ключи: солнечные очки, солнцезащитные очки, брендовые очки.",
        "По запросам «очки солнцезащитные» и «солнечные очки» такие модели часто выбирают за форму и комфорт."
    ])


def _keywords_tail(brand: str, shape_norm: str, lens_norm: str, rng: random.Random) -> str:
    pool = [
        "солнцезащитные очки", "солнечные очки", "очки солнцезащитные", "брендовые очки",
        f"очки {brand}", f"{brand} очки",
        "очки UV400", "очки поляризационные", "очки фотохромные", "очки хамелеон",
        "очки квадратные", "очки круглые", "очки авиаторы", "очки cat eye",
        "инста очки", "очки из tiktok"
    ]
    rng.shuffle(pool)
    out, seen = [], set()
    for x in pool:
        xl = x.lower()
        if xl in seen:
            continue
        seen.add(xl)
        out.append(x)
        if len(out) >= 9:
            break
    return "Ключевые фразы: " + ", ".join(out) + "."


def _compose_desc(
    style_key: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str,
    seed: int,
    tail_mode: str,
    profile: Dict
) -> str:
    rng = random.Random(seed)

    brand = _clean(brand) or "Бренд"
    shape_n = _norm_shape(shape)
    lens_n = _normalize_lens(lens_features)
    collection = _clean(collection) or "Весна–Лето 2025–2026"

    # бренд-профильные строки
    hooks = profile.get("hooks") or []
    style_lines = profile.get("style_lines") or []
    closings = profile.get("closings") or []

    sc = _scenarios(rng)

    blocks = [
        rng.choice(hooks) if hooks else rng.choice([
            f"Очки {brand} — тот аксессуар, который быстро делает образ собраннее.",
            f"Солнцезащитные очки {brand} хорошо заходят в сезон {collection}: спокойно и уместно."
        ]),
        _shape_line(shape_n, rng),
        rng.choice(style_lines) if style_lines else rng.choice([
            "Посадка и силуэт выглядят аккуратно — не перегружают образ.",
            "Форма оправы читается в образе и сочетается с повседневной одеждой."
        ]),
        _lens_benefit(lens_n, rng),
        rng.choice([
            f"Сценарии: {sc}.",
            f"Лучше всего раскрываются в сценариях {sc} — когда день проходит на улице."
        ]),
        rng.choice([
            "По ощущениям это пара, которую легко носить каждый день: надел — и не думаешь о ней.",
            "Есть ощущение законченности образа: очки работают как аккуратный акцент."
        ]),
        _seo_sentence(brand, rng),
        rng.choice(closings) if closings else rng.choice([
            f"Хороший выбор на сезон {collection}: стильно и удобно.",
            "Подойдёт и себе, и в подарок — практично и красиво."
        ]),
        f"Форма: {shape_n}. Линзы: {lens_n}. Коллекция: {collection}."
    ]

    # тасуем середину, чтобы не было “одного шаблона”
    head = blocks[0]
    tail = blocks[-1]
    middle = blocks[1:-1]
    rng.shuffle(middle)

    text = _clean(" ".join([head] + middle + [tail]))

    # профильные “разрешённые” слова — иногда вставляем 1–2 мягко
    allow_words = profile.get("allow_words") or []
    if allow_words and rng.random() < 0.65:
        w = rng.choice(allow_words)
        inject = rng.choice([
            f"По настроению это ощущается {w}.",
            f"В образе это выглядит {w} — без лишней громкости."
        ])
        cand = _clean(text + " " + inject)
        if len(cand) <= MAX_DESC_LEN:
            text = cand

    # хвост ключей
    def need_tail(desc: str) -> bool:
        if tail_mode == "always":
            return True
        if tail_mode == "never":
            return False
        t = desc.lower()
        core = sum(1 for k in ["солнцезащитные очки", "солнечные очки", "очки солнцезащитные", "брендовые очки"] if k in t)
        has_lens = any(k in t for k in ["uv400", "поляр", "фотохром", "хамелеон"])
        return not (core >= 2 and has_lens)

    if need_tail(text):
        cand = _clean(text + " " + _keywords_tail(brand, shape_n, lens_n, rng))
        if len(cand) <= MAX_DESC_LEN:
            text = cand

    # добиваем до приятной длины (без одинакового ощущения)
    fillers = [
        "Эта модель не спорит с одеждой — она просто аккуратно усиливает образ.",
        "В солнечную погоду это реально ощущается: комфортнее глазам и спокойнее на улице.",
        "Сочетается с базовыми вещами и при этом добавляет заметный акцент.",
        "Хорошо выглядит вживую и на фото — без ощущения “шаблонного” аксессуара."
    ]
    rng.shuffle(fillers)
    i = 0
    while len(text) < TARGET_DESC_MIN and i < len(fillers):
        cand = _clean(text + " " + fillers[i])
        if len(cand) <= MAX_DESC_LEN:
            text = cand
        i += 1

    # чистим “запрещённые слова”
    text = _filter_avoid(text, profile.get("avoid_words") or [])

    if len(text) > MAX_DESC_LEN:
        text = text[:MAX_DESC_LEN].rsplit(" ", 1)[0].rstrip(".,;:") + "."

    return text


# ======================
# Excel helpers
# ======================

def _find_headers(ws) -> Tuple[int, int, int]:
    header_row = None
    col_name = None
    col_desc = None
    for r in range(1, 40):
        for c in range(1, 600):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                vv = v.strip()
                if vv == "Наименование":
                    header_row = r
                    col_name = c
                elif vv == "Описание":
                    col_desc = c
        if header_row and col_name and col_desc:
            return header_row, col_name, col_desc
    raise RuntimeError("Не найдены колонки 'Наименование' и 'Описание' в первых 40 строках.")


def _write_safe(ws, row: int, col: int, value: str) -> None:
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                ws.cell(mr.min_row, mr.min_col).value = value
                return
        return
    cell.value = value


# ======================
# Main
# ======================

def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str,
    style: str = "neutral",
    tail_mode: str = "adaptive",
    overwrite_existing: bool = True,
    include_search_tail: bool = True,   # legacy
    progress_callback: Optional[Callable[[int], None]] = None,
) -> Tuple[str, int]:

    profiles_data = load_brand_profiles()
    profile = pick_profile(brand, profiles_data)

    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, col_name, col_desc = _find_headers(ws)
    start_row = max(header_row + 1, PROTECT_TOP_ROWS + 1)

    deck_seed = int(time.time() * 1000) ^ (hash(Path(input_xlsx).name) & 0xFFFF_FFFF)
    slogan_deck = SloganDeck(deck_seed)

    # считаем “сколько строк реально будем трогать”
    target_rows = []
    for r in range(start_row, ws.max_row + 1):
        if all(ws.cell(r, c).value in (None, "") for c in range(1, 20)):
            continue
        target_rows.append(r)

    total = len(target_rows)
    filled = 0
    recent_descs: List[str] = []

    def report(p: int):
        if progress_callback:
            try:
                progress_callback(max(0, min(100, int(p))))
            except Exception:
                pass

    report(0)

    for idx, r in enumerate(target_rows, start=1):
        if not overwrite_existing:
            if _clean(ws.cell(r, col_name).value) or _clean(ws.cell(r, col_desc).value):
                continue

        row_seed_base = (r * 1315423911) ^ deck_seed
        rng = random.Random(row_seed_base)

        title = build_title(slogan_deck, brand, shape, lens_features, rng)

        desc = ""
        for attempt in range(REGEN_ATTEMPTS):
            seed = row_seed_base ^ (attempt * 2654435761)
            candidate = _compose_desc(
                style_key=style,
                brand=brand,
                shape=shape,
                lens_features=lens_features,
                collection=collection,
                seed=seed,
                tail_mode=tail_mode,
                profile=profile
            )
            too_similar = False
            for prev in recent_descs[-20:]:
                if _jaccard(candidate, prev) >= SIMILARITY_THRESHOLD:
                    too_similar = True
                    break
            if not too_similar:
                desc = candidate
                break

        if not desc:
            desc = _compose_desc(style, brand, shape, lens_features, collection, seed=row_seed_base, tail_mode=tail_mode, profile=profile)

        _write_safe(ws, r, col_name, title)
        _write_safe(ws, r, col_desc, desc)

        recent_descs.append(desc)
        filled += 1

        report((idx / max(1, total)) * 100)

    in_path = Path(input_xlsx)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = in_path.parent / f"{in_path.stem}_filled_{ts}.xlsx"
    wb.save(out_path)

    report(100)
    return str(out_path), filled
