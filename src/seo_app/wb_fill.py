from __future__ import annotations

import re
import time
import random
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, List

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# ======================
# Limits / Settings
# ======================

MAX_TITLE_LEN = 60
MAX_DESC_LEN = 2000
TARGET_DESC_MIN = 1100          # целимся в “живой” объём
PROTECT_TOP_ROWS = 4            # не трогаем строки 1..4
SIMILARITY_THRESHOLD = 0.72     # чем меньше — тем сильнее уникализация
REGEN_ATTEMPTS = 8              # пересборка при похожести


# ======================
# Helpers
# ======================

def _clean(s: Optional[str]) -> str:
    return re.sub(r"\s+", " ", str(s)).strip() if s else ""


def _norm_shape(shape: str) -> str:
    s = _clean(shape).lower()
    if "квад" in s:
        return "квадратная"
    if "круг" in s:
        return "круглая"
    if "кошач" in s or "cat" in s:
        return "cat eye кошачий глаз"
    if "авиат" in s:
        return "авиаторы"
    if "прямоуг" in s:
        return "прямоугольная"
    if "овер" in s:
        return "овальная"
    if not s:
        return "универсальная"
    return _clean(shape)


def _shape_for_title(shape: str) -> str:
    s = _norm_shape(shape)
    mapping = {
        "квадратная": "квадратные",
        "круглая": "круглые",
        "прямоугольная": "прямоугольные",
        "авиаторы": "авиаторы",
        "cat eye": "кошачий глаз",
        "овальные": "овальные",
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

    # unique preserve order
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

    # приоритеты для title
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


def _tokenize(text: str) -> set:
    # очень простая токенизация: слова/числа
    return set(re.findall(r"[a-zа-яё0-9]+", text.lower(), flags=re.IGNORECASE))


def _jaccard(a: str, b: str) -> float:
    A = _tokenize(a)
    B = _tokenize(b)
    if not A or not B:
        return 0.0
    return len(A & B) / len(A | B)


# ======================
# Title slogan pool (no duplicates per run)
# ======================

SLOGANS = [
    "Красивые", "Стильные", "Модные", "Крутые", "Молодёжные",
    "Трендовые", "Актуальные", "Эффектные", "Элегантные",
    "Лаконичные", "Дизайнерские", "Премиальные", "Современные",
    "Универсальные", "Практичные", "Лёгкие", "Комфортные",
    "Статусные", "Выразительные", "Минималистичные",
    "Городские", "Повседневные", "Сезонные", "Функциональные",
    "Изысканные", "Яркие", "Сдержанные", "Удобные",
    "Брендовые", "Коллекционные", "Стильные", "Классные"
]

PRODUCT_PHRASES = ["солнцезащитные очки", "солнечные очки"]


class SloganDeck:
    """
    Колода лозунгов: внутри одного файла лозунги не повторяются,
    пока колода не закончится. На новый запуск файла колода перемешивается.
    """
    def __init__(self, seed: int):
        self.rng = random.Random(seed)
        self.deck = SLOGANS[:]
        self.rng.shuffle(self.deck)
        self.i = 0

    def next(self) -> str:
        if self.i >= len(self.deck):
            self.rng.shuffle(self.deck)
            self.i = 0
        s = self.deck[self.i]
        self.i += 1
        return s


def build_title(deck: SloganDeck, brand: str, shape: str, lens_features: str, rng: random.Random) -> str:
    slogan = deck.next()
    product = rng.choice(PRODUCT_PHRASES)
    feat = _pick_feat_for_title(lens_features, shape)

    # иногда добавим форму, но только если помещается
    shape_tail = _shape_for_title(shape)
    tail_options = [
        [feat],
        [shape_tail],
        [feat, shape_tail],
    ]
    rng.shuffle(tail_options)

    candidates = []
    for tail in tail_options:
        title = _fit_title_tokens([slogan, product, brand] + tail)
        candidates.append(title)

    # выбираем самый “богатый” (длина ближе к лимиту и есть фича)
    def score(t: str) -> float:
        L = len(t)
        s = 0.0
        s += 4.0 if any(t.startswith(x + " ") for x in SLOGANS) else -100.0
        s += 2.0 if ("солнцезащитные очки" in t or "солнечные очки" in t) else -10.0
        s += 2.0 if _clean(brand).lower() in t.lower() else -10.0
        s += 1.5 if any(k in t.lower() for k in ["uv400", "replica", "поляр", "хамелеон", "откид", "квадрат", "кругл", "авиатор", "cat eye"]) else 0.0
        s += 2.0 if 48 <= L <= 60 else (1.0 if 40 <= L <= 60 else 0.0)
        return s

    candidates.sort(key=score, reverse=True)
    return candidates[0]


# ======================
# Description generator (human-like)
# ======================

def _scenarios(rng: random.Random) -> str:
    picks = rng.sample(
        ["город", "путешествия", "отпуск", "пляж", "прогулки", "вождение", "активный отдых", "летние выходы"],
        k=4
    )
    return ", ".join(picks)


def _lens_benefit_line(lens_norm: str, rng: random.Random) -> str:
    ln = lens_norm.lower()
    variants = []

    if "поляр" in ln:
        variants += [
            "Поляризация помогает убрать блики — особенно приятно у воды и за рулём.",
            "С поляризационными линзами меньше бликов и визуально комфортнее в яркий день.",
        ]
    if "uv400" in ln:
        variants += [
            "UV400 — базовая защита от солнца, когда хочется просто спокойно ходить по улице без дискомфорта.",
            "UV400 даёт нормальный уровень защиты для яркой погоды и долгих прогулок.",
        ]
    if "фотохром" in ln or "хамелеон" in ln:
        variants += [
            "Фотохромные линзы подстраиваются под освещение — удобно, когда день постоянно меняется.",
            "Линзы-хамелеон реагируют на свет: на улице темнее, в тени мягче.",
        ]
    if "откид" in ln:
        variants += [
            "Откидные линзы — это удобно: быстро переключился под ситуацию и поехал дальше.",
            "Фишка с откидными линзами спасает, когда нужен другой режим буквально на ходу.",
        ]
    if "replica" in ln:
        # ВАЖНО: я не добавляю текст, который помогает продавать контрафакт.
        # Просто нейтральная фраза, без “бренд/оригинал/копия”.
        variants += [
            "Если тебе важен визуальный стиль, тут как раз тот формат, который легко вписать в повседневные образы.",
        ]

    if not variants:
        variants = [
            "Линзы подобраны так, чтобы в солнечную погоду было комфортно и в городе, и в поездках.",
            "Комфорт в яркий день — главное: меньше щуриться, легче глазам.",
        ]
    return rng.choice(variants)


def _shape_style_line(shape_norm: str, rng: random.Random) -> str:
    sn = shape_norm.lower()
    variants = []
    if "квадрат" in sn:
        variants += [
            "Квадратная форма смотрится собранно и добавляет образу структуру.",
            "Квадратная оправа делает образ более “чётким” и современным.",
        ]
    if "кругл" in sn:
        variants += [
            "Круглая оправа смотрится мягче и часто даёт эффект “винтажного” акцента.",
            "Круглая форма добавляет лёгкости — особенно в летних образах.",
        ]
    if "cat eye" in sn:
        variants += [
            "Форма cat eye визуально “поднимает” образ и выглядит очень женственно.",
            "Cat eye — это аккуратный вау-эффект без лишнего шума.",
        ]
    if "авиат" in sn:
        variants += [
            "Авиаторы — классика, которая вечно выглядит уместно: и город, и отпуск.",
            "Форма “авиаторы” обычно идёт почти всем — универсальная история на каждый день.",
        ]
    if not variants:
        variants = [
            f"Форма оправы {shape_norm} выглядит актуально и легко сочетается с базовым гардеробом.",
            f"Оправa {shape_norm} — вариант, который спокойно носится каждый день.",
        ]
    return rng.choice(variants)


def _seo_sentence(brand: str, rng: random.Random) -> str:
    variants = [
        "Если ищешь солнцезащитные очки на каждый день, здесь совпадает и стиль, и практичность.",
        f"По запросам «солнечные очки» и «очки солнцезащитные» такие модели часто ищут по бренду {brand} и форме оправы.",
        "Это как раз тот случай, когда SEO-слова не отдельно, а внутри смысла: солнечные очки, брендовые очки, UV400 и удобство носки.",
    ]
    return rng.choice(variants)


def _closing(brand: str, collection: str, rng: random.Random) -> str:
    variants = [
        f"Хороший выбор на сезон {collection}: носится легко и не надоедает.",
        f"Эти очки {brand} — удобный способ обновить образ к {collection} без лишних экспериментов.",
        "Если хочется одну пару, которая “работает” и в городе, и на отдыхе — это оно.",
    ]
    return rng.choice(variants)


# “Голоса” — не стили UI, а манера речи
def _voice_blocks(style_key: str) -> List[str]:
    # style_key приходит из GUI: neutral/premium/mass/social
    # но мы всё равно оставляем вариативность, чтобы не было одинаковых текстов
    if style_key == "premium":
        return ["stylist", "observer", "seller"]
    if style_key == "mass":
        return ["practical", "seller", "observer"]
    if style_key == "social":
        return ["social", "stylist", "observer"]
    return ["stylist", "practical", "observer", "seller", "social"]


def _compose_desc(style_key: str, brand: str, shape: str, lens_features: str, collection: str, seed: int, tail_mode: str) -> str:
    rng = random.Random(seed)
    brand = _clean(brand)
    shape_n = _norm_shape(shape)
    lens_n = _normalize_lens(lens_features)
    collection = _clean(collection) or "Весна–Лето 2025–2026"

    sc = _scenarios(rng)

    # Библиотека смысловых блоков (живые, разные)
    blocks = {
        "hook": [
            f"Очки {brand} — тот аксессуар, который быстро делает образ собраннее.",
            f"Солнцезащитные очки {brand} хорошо заходят в сезон {collection}: без лишней “показухи”, но со стилем.",
            f"Если хочется солнечные очки, которые выглядят уместно и в городе, и на отдыхе — это хороший кандидат.",
            f"В тёплый сезон важно, чтобы очки были не только красивыми, но и удобными — у этой модели как раз такой характер.",
        ],
        "shape": [
            _shape_style_line(shape_n, rng),
            f"Форма оправы {shape_n} хорошо читается в образе и не выглядит случайно.",
            f"Посадка и силуэт {shape_n} — то, что обычно выбирают на каждый день: сочетается легко.",
        ],
        "lens": [
            _lens_benefit_line(lens_n, rng),
            f"Линзы {lens_n} дают комфорт при ярком свете — особенно в середине дня.",
            f"С {lens_n} проще переносить солнечную погоду: меньше напряжения для глаз.",
        ],
        "scenarios": [
            f"Сценарии, где они особенно хороши: {sc}.",
            f"Лучше всего раскрываются в сценариях {sc} — когда день проходит на улице.",
            f"Для {sc} — именно то, что нужно: удобно и по делу.",
        ],
        "season": [
            f"Коллекция/сезон: {collection}.",
            f"Актуально для сезона {collection} — когда хочется обновления без риска.",
            f"В {collection} такой силуэт смотрится особенно уместно.",
        ],
        "microhuman": [
            "По ощущениям это “простая в носке” пара: надел — и не думаешь о ней весь день.",
            "Есть ощущение законченности образа: не нужно добавлять ничего лишнего.",
            "Смотрится аккуратно и “дорого” без попытки кричать о себе.",
            "Та самая вещь, которая часто становится любимой: потому что просто удобно.",
        ],
        "seo": [
            _seo_sentence(brand, rng),
            "Ключевые категории, где это ищут: солнцезащитные очки, солнечные очки, очки UV400, брендовые очки.",
        ],
        "close": [
            _closing(brand, collection, rng),
            f"Если выбираешь солнечные очки на {collection}, эта модель — спокойный и надёжный вариант.",
            "Подойдёт и себе, и в подарок: выглядит хорошо и носится комфортно.",
        ],
    }

    # Определяем “голос” + схему сборки
    allowed_voices = _voice_blocks(style_key)
    voice = rng.choice(allowed_voices)

    # Схемы: 5–7 блоков, но в разном порядке (чтобы не палилось)
    if voice == "social":
        scheme = rng.choice([
            ["hook", "shape", "lens", "scenarios", "seo", "close"],
            ["hook", "lens", "shape", "microhuman", "close", "seo"],
        ])
    elif voice == "seller":
        scheme = rng.choice([
            ["hook", "season", "shape", "lens", "scenarios", "close", "seo"],
            ["hook", "lens", "season", "shape", "microhuman", "close", "seo"],
        ])
    elif voice == "practical":
        scheme = rng.choice([
            ["hook", "lens", "scenarios", "shape", "microhuman", "close", "seo"],
            ["hook", "scenarios", "lens", "shape", "close", "seo"],
        ])
    elif voice == "stylist":
        scheme = rng.choice([
            ["hook", "shape", "microhuman", "lens", "season", "close", "seo"],
            ["hook", "shape", "season", "microhuman", "lens", "close", "seo"],
        ])
    else:  # observer
        scheme = rng.choice([
            ["hook", "shape", "lens", "scenarios", "microhuman", "close", "seo"],
            ["hook", "microhuman", "shape", "lens", "season", "close", "seo"],
        ])

    # Собираем текст
    parts = []
    for key in scheme:
        parts.append(rng.choice(blocks[key]))

    text = _clean(" ".join(parts))

    # “Хвост” ключей (по режиму)
    def need_tail(desc: str) -> bool:
        if tail_mode == "always":
            return True
        if tail_mode == "never":
            return False
        # adaptive
        t = desc.lower()
        core = sum(1 for k in ["солнцезащитные очки", "солнечные очки", "очки солнцезащитные", "брендовые очки"] if k in t)
        has_lens = any(k in t for k in ["uv400", "поляр", "фотохром", "хамелеон"])
        return not (core >= 2 and has_lens)

    if need_tail(text):
        tail_pool = [
            "солнцезащитные очки", "солнечные очки", "очки солнцезащитные", "брендовые очки",
            f"очки {brand}", f"{brand} очки",
            "очки UV400", "очки поляризационные", "очки фотохромные", "очки хамелеон",
            "очки квадратные", "очки круглые", "очки авиаторы", "очки cat eye",
            "инста очки", "очки из tiktok",
        ]
        rng.shuffle(tail_pool)
        tail = []
        seen = set()
        for x in tail_pool:
            xl = x.lower()
            if xl in seen:
                continue
            seen.add(xl)
            tail.append(x)
            if len(tail) >= 9:
                break
        text = _clean(text + " " + "Ключевые фразы: " + ", ".join(tail) + ".")

    # Добиваем до “живой” длины (без одинаковых вставок)
    fillers = [
        "Важно, что модель не выглядит случайной: она именно дополняет стиль.",
        "Если ты любишь, когда аксессуар не спорит с одеждой — это сильный вариант.",
        "В таких очках обычно быстро привыкаешь ходить каждый день.",
        "Летом это реально выручает: вышел — и уже комфортнее глазам.",
        "Визуально форма делает образ более собранным — даже с простой футболкой.",
    ]
    rng.shuffle(fillers)
    i = 0
    while len(text) < TARGET_DESC_MIN and i < len(fillers):
        cand = _clean(text + " " + fillers[i])
        if len(cand) <= MAX_DESC_LEN:
            text = cand
        i += 1

    if len(text) > MAX_DESC_LEN:
        text = text[:MAX_DESC_LEN].rsplit(" ", 1)[0].rstrip(".,;:") + "."

    return text


# ======================
# Excel utils
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
# Main entry
# ======================

def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str,
    style: str = "neutral",        # neutral | premium | mass | social
    tail_mode: str = "adaptive",   # adaptive | always | never
    overwrite_existing: bool = True,
    include_search_tail: bool = True,  # legacy arg (ignored)
) -> Tuple[str, int]:

    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, col_name, col_desc = _find_headers(ws)
    start_row = max(header_row + 1, PROTECT_TOP_ROWS + 1)

    # новый запуск → новый deck seed
    deck_seed = int(time.time() * 1000) ^ (hash(Path(input_xlsx).name) & 0xFFFF_FFFF)
    slogan_deck = SloganDeck(deck_seed)

    filled = 0
    recent_descs: List[str] = []

    for r in range(start_row, ws.max_row + 1):
        # пустые строки пропускаем
        if all(ws.cell(r, c).value in (None, "") for c in range(1, 20)):
            continue

        if not overwrite_existing:
            if _clean(ws.cell(r, col_name).value) or _clean(ws.cell(r, col_desc).value):
                continue

        # rng на строку + соль на запуск
        row_seed_base = (r * 1315423911) ^ deck_seed
        rng = random.Random(row_seed_base)

        title = build_title(slogan_deck, brand, shape, lens_features, rng)

        # Генерация описания с регеном при похожести
        desc = ""
        for attempt in range(REGEN_ATTEMPTS):
            seed = row_seed_base ^ (attempt * 2654435761)
            candidate = _compose_desc(style, brand, shape, lens_features, collection, seed=seed, tail_mode=tail_mode)

            # проверка похожести с последними (чтобы быстро работало)
            too_similar = False
            for prev in recent_descs[-20:]:
                if _jaccard(candidate, prev) >= SIMILARITY_THRESHOLD:
                    too_similar = True
                    break

            if not too_similar:
                desc = candidate
                break

        if not desc:
            desc = _compose_desc(style, brand, shape, lens_features, collection, seed=row_seed_base, tail_mode=tail_mode)

        _write_safe(ws, r, col_name, title)
        _write_safe(ws, r, col_desc, desc)

        recent_descs.append(desc)
        filled += 1

    in_path = Path(input_xlsx)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = in_path.parent / f"{in_path.stem}_filled_{ts}.xlsx"
    wb.save(out_path)

    return str(out_path), filled
