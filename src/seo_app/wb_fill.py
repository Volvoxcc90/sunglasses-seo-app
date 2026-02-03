import re
import time
import random
from pathlib import Path
from typing import Callable, Optional, Tuple, List, Dict

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# ==========================
# Лимиты
# ==========================
TITLE_MAX_LEN = 60
DESC_MAX_LEN = 2000

# Режимы длины описания (ориентиры по символам)
DESC_LENGTH_RANGES = {
    "short":  (550, 850),
    "medium": (900, 1400),
    "long":   (1500, 2000),
}

# SEO плотность
# soft:   2-3 ключа
# normal: 4-6 ключей
# hard:   7-9 ключей (но всё равно “человечно”, без списков)
SEO_LEVEL_COUNTS = {
    "soft":   {"core": 1, "tail": 1, "feature": 0, "extra": 0},
    "normal": {"core": 2, "tail": 1, "feature": 1, "extra": 0},
    "hard":   {"core": 2, "tail": 2, "feature": 1, "extra": 1},
}

# Жёстко запрещаем ярлыки
FORBIDDEN_LABELS = [
    "сценарии:", "ключевые слова:", "форма:", "линза:", "коллекция:"
]

# ==========================
# Словари
# ==========================
SLOGANS = [
    "Красивые", "Крутые", "Стильные", "Модные", "Молодёжные", "Трендовые",
    "Эффектные", "Дизайнерские", "Лаконичные", "Яркие", "Премиальные",
    "Удобные", "Лёгкие", "Универсальные", "Городские", "Летние",
    "Актуальные", "Смелые", "Элегантные", "Минималистичные",
    "Современные", "Ультрамодные", "Хитовые", "Культовые", "Фирменные",
    "Супер", "Невероятные", "Топовые", "Сочные"
]

SCENARIOS = [
    "город", "прогулки", "отпуск", "пляж", "путешествия",
    "вождение", "активный отдых", "повседневные дела",
    "кафе и встречи", "поездки", "выходные", "летние фестивали"
]

SEO_CORE = [
    "солнцезащитные очки", "солнечные очки", "очки солнцезащитные",
    "брендовые очки", "модные очки"
]

SEO_TAIL = [
    "очки для города", "очки для отпуска", "очки для вождения",
    "очки для путешествий", "аксессуар на лето", "очки унисекс",
    "инста очки", "очки из tiktok"
]

SEO_FEATURES = [
    "UV400", "поляризационные очки", "фотохромные очки",
    "зеркальные линзы", "градиентные линзы"
]


# ==========================
# Утилиты
# ==========================
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def _cut_no_word_break(text: str, max_len: int) -> str:
    text = _norm(text)
    if len(text) <= max_len:
        return text
    cut = text[:max_len].rsplit(" ", 1)[0]
    return cut.strip() if cut else text[:max_len].strip()


def _sun_term() -> str:
    return random.choice(["солнцезащитные очки", "солнечные очки"])


def _contains_cyrillic(s: str) -> bool:
    return bool(re.search(r"[А-Яа-яЁё]", s or ""))


def _strip_forbidden(text: str) -> str:
    t = text
    for lab in FORBIDDEN_LABELS:
        t = re.sub(re.escape(lab), "", t, flags=re.IGNORECASE)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t


def _first_n_words(text: str, n: int = 7) -> str:
    w = re.sub(r"[^0-9A-Za-zА-Яа-яёЁ ]+", " ", (text or "")).split()
    return " ".join(w[:n]).lower()


def _jaccard(a: str, b: str) -> float:
    def tok(x: str) -> set:
        x = re.sub(r"[^0-9A-Za-zА-Яа-яёЁ ]+", " ", x.lower())
        return {p for p in x.split() if len(p) > 2}
    A, B = tok(a), tok(b)
    return len(A & B) / max(1, len(A | B)) if A and B else 0.0


def _clamp_modes(style: str, seo_level: str, desc_length: str) -> Tuple[str, str, str]:
    style = (style or "neutral").lower().strip()
    if style not in {"neutral", "premium", "social"}:
        style = "neutral"

    seo_level = (seo_level or "normal").lower().strip()
    if seo_level not in {"soft", "normal", "hard"}:
        seo_level = "normal"

    desc_length = (desc_length or "medium").lower().strip()
    if desc_length not in {"short", "medium", "long"}:
        desc_length = "medium"

    return style, seo_level, desc_length


# ==========================
# Бренд -> кириллица (в названии)
# ==========================
BRAND_RU_OVERRIDES = {
    "gucci": "Гуччи",
    "dior": "Диор",
    "prada": "Прада",
    "ray-ban": "Рэй-Бэн",
    "ray ban": "Рэй-Бэн",
    "cazal": "Казал",
    "versace": "Версаче",
    "chanel": "Шанель",
    "cartier": "Картье",
    "oakley": "Окли",
    "dolce gabbana": "Дольче Габбана",
    "dolce & gabbana": "Дольче Габбана",
    "armani": "Армани",
    "burberry": "Бёрберри",
    "balenciaga": "Баленсиага",
}

TRANSLIT_MAP = [
    ("sch", "ш"), ("ch", "ч"), ("sh", "ш"), ("ya", "я"), ("yu", "ю"),
    ("yo", "ё"), ("ye", "е"), ("kh", "х"), ("ts", "ц"),
    ("a", "а"), ("b", "б"), ("c", "к"), ("d", "д"), ("e", "е"),
    ("f", "ф"), ("g", "г"), ("h", "х"), ("i", "и"), ("j", "дж"),
    ("k", "к"), ("l", "л"), ("m", "м"), ("n", "н"), ("o", "о"),
    ("p", "п"), ("q", "к"), ("r", "р"), ("s", "с"), ("t", "т"),
    ("u", "у"), ("v", "в"), ("w", "в"), ("x", "кс"), ("y", "и"), ("z", "з"),
]

def brand_to_cyrillic(brand: str) -> str:
    brand = _norm(brand)
    if not brand:
        return ""
    if _contains_cyrillic(brand):
        return brand

    key = brand.lower().replace("&", " ").replace("-", " ").strip()
    key = re.sub(r"\s+", " ", key)
    if key in BRAND_RU_OVERRIDES:
        return BRAND_RU_OVERRIDES[key]

    b = key
    for latin, ru in TRANSLIT_MAP:
        b = b.replace(latin, ru)
    return b[:1].upper() + b[1:]


# ==========================
# Наименования: 6 вариантов, 3 с брендом / 3 без
# + анти-повторы лозунга по строкам
# ==========================
def build_titles_6(brand: str, shape: str, lens: str) -> List[str]:
    brand_ru = brand_to_cyrillic(brand)
    shape = _norm(shape)
    lens = _norm(lens)

    flags = [True, True, True, False, False, False]
    random.shuffle(flags)

    templates = [
        "{slogan} {core} {brand}{shape}{lens}",
        "{slogan} {core} {shape}{brand}{lens}",
        "{slogan} {core} {lens}{brand}{shape}",
        "{slogan} {core} {brand}{lens}{shape}",
        "{slogan} {core} {shape}{lens}{brand}",
        "{slogan} {core} {lens}{shape}{brand}",
    ]

    # берём 6 разных лозунгов, чтобы внутри одной строки не было дублей
    local_slogans = random.sample(SLOGANS, k=6) if len(SLOGANS) >= 6 else [random.choice(SLOGANS) for _ in range(6)]

    used = set()
    out = []

    for i in range(6):
        slogan = local_slogans[i]
        core = _sun_term()

        brand_part = (brand_ru + " ") if (flags[i] and brand_ru) else ""
        shape_part = (shape + " ") if (shape and random.random() < 0.55) else ""
        lens_part = (lens + " ") if (lens and random.random() < 0.70) else ""

        raw = templates[i].format(
            slogan=slogan,
            core=core,
            brand=brand_part,
            shape=shape_part,
            lens=lens_part
        )
        title = _cut_no_word_break(raw, TITLE_MAX_LEN)

        tries = 0
        while title in used and tries < 8:
            slogan = random.choice(SLOGANS)
            core = _sun_term()
            raw = templates[i].format(
                slogan=slogan,
                core=core,
                brand=brand_part,
                shape=shape_part,
                lens=lens_part
            )
            title = _cut_no_word_break(raw, TITLE_MAX_LEN)
            tries += 1

        used.add(title)
        out.append(title)

    return out


def pick_best_title(titles: List[str], last_slogan: str, recent_starts: List[str]) -> str:
    """
    Выбираем наименование:
    - ближе к 55 символам
    - НЕ повторяет лозунг прошлой строки (если можем)
    - НЕ повторяет старт (первые 4 слова) последних 3 строк (если можем)
    """
    def slogan_of(t: str) -> str:
        return (t.split(" ", 1)[0] if t else "").strip()

    def start4(t: str) -> str:
        return " ".join((t or "").split()[:4]).lower()

    scored = []
    for t in titles:
        L = len(t)
        score = -abs(55 - L)
        if last_slogan and slogan_of(t).lower() == last_slogan.lower():
            score -= 6
        if recent_starts and start4(t) in recent_starts:
            score -= 5
        scored.append((score, t))

    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1] if scored else (titles[0] if titles else "")


# ==========================
# SEO-факт по линзам (доверие + поиск, без ярлыков)
# ==========================
def _lens_fact(lens: str) -> str:
    l = (lens or "").lower()
    if "uv400" in l:
        return random.choice([
            "UV400 — популярный ориентир при выборе: комфортнее при ярком солнце и меньше хочется щуриться на улице.",
            "Защита UV400 обычно выбирается для уверенного использования в солнечную погоду — особенно в городе и в поездках.",
        ])
    if "поляр" in l:
        return random.choice([
            "Поляризация помогает уменьшить блики от воды, стекла и асфальта — особенно заметно в дороге и на открытых пространствах.",
            "Поляризационный эффект делает картинку более читаемой при ярком свете и снижает утомляемость глаз.",
        ])
    if "фото" in l or "хамеле" in l:
        return random.choice([
            "Фотохромный эффект удобен, когда освещение меняется: на улице темнее, в помещении спокойнее.",
            "Фотохромные линзы подходят тем, кто часто выходит из помещения на улицу и обратно.",
        ])
    return ""


def _choose_keywords(brand: str, lens: str, seo_level: str) -> Dict[str, List[str]]:
    cfg = SEO_LEVEL_COUNTS[seo_level]
    core = random.sample(SEO_CORE, k=cfg["core"])
    tail = random.sample(SEO_TAIL, k=cfg["tail"])

    features = []
    if cfg["feature"] > 0:
        l = (lens or "").lower()
        if "uv400" in l:
            features.append("UV400")
        elif "поляр" in l:
            features.append("поляризационные очки")
        elif "фото" in l or "хамеле" in l:
            features.append("фотохромные очки")
        else:
            features.append(random.choice(SEO_FEATURES))

    extra = []
    if cfg["extra"] > 0:
        # “мягкие” дополнительные хвосты
        extra.append(random.choice(["очки женские", "очки мужские", "очки унисекс", "брендовые очки"]))

    # гарантируем, что будет или “солнцезащитные”, или “солнечные”
    if not any(("солнцезащитные" in x or "солнечные" in x) for x in core):
        core[0] = _sun_term()

    return {"core": core, "tail": tail, "features": features, "extra": extra}


# ==========================
# Описание: логично + SEO + живо + режимы длины + анти-повторы
# ==========================
def generate_description(
    brand: str,
    shape: str,
    lens: str,
    collection: str,
    style: str,
    seo_level: str,
    desc_length: str,
    recent_desc_starts: List[str],
) -> str:
    brand = _norm(brand)
    shape = _norm(shape)
    lens = _norm(lens)
    collection = _norm(collection)
    style, seo_level, desc_length = _clamp_modes(style, seo_level, desc_length)

    min_len, max_len = DESC_LENGTH_RANGES[desc_length]
    kw = _choose_keywords(brand, lens, seo_level)
    scen = ", ".join(random.sample(SCENARIOS, k=4))

    # 3–4 “подачи” внутри стиля
    openers = []
    if brand:
        openers += [
            f"{brand} — это когда аксессуар работает на образ и на комфорт: удачный выбор на яркие дни и активный ритм.",
            f"Очки {brand} легко вписываются в гардероб: выглядят актуально и подходят, когда нужен уверенный летний акцент.",
            f"Если хочется обновить образ без лишней сложности — {brand} смотрятся выразительно и при этом остаются удобными.",
        ]
    else:
        openers += [
            "Эта модель выглядит актуально и подходит, когда нужен заметный, но аккуратный акцент на лице.",
            "Хороший вариант на яркие дни: удобно носить, легко сочетать и выглядит свежо в городской среде.",
        ]

    design = random.choice([
        f"Дизайн с {shape.lower()} линиями подчёркивает черты лица и делает образ собранным — от повседневных луков до более смелых стилизаций."
        if shape else
        "Дизайн подчёркивает черты лица и делает образ собранным — от повседневных луков до более смелых стилизаций.",
        f"Оправа выглядит современно и помогает сбалансировать пропорции лица — особенно хорошо смотрится в дневном свете."
    ])

    lenses = random.choice([
        f"Линзы {lens} дают комфорт при ярком солнце и подходят для дня “в движении”: прогулки, дорога, открытые пространства."
        if lens else
        "Линзы дают комфорт при ярком солнце и подходят для дня “в движении”: прогулки, дорога, открытые пространства.",
        f"С {lens} меньше хочется щуриться на улице, и дневной свет воспринимается спокойнее — это особенно ценится в городе и в поездках."
        if lens else
        "Дневной свет воспринимается спокойнее — это особенно ценится в городе и в поездках."
    ])

    fact = _lens_fact(lens)
    season = ""
    if collection and random.random() < 0.85:
        season = random.choice([
            f"Сезон {collection} — время лёгких деталей: модель выглядит свежо и не “устаревает” через месяц.",
            f"Актуально на {collection}: можно носить каждый день и при этом сохраняется ощущение трендовой вещи.",
        ])

    if style == "premium":
        vibe = random.choice([
            "Визуально модель выглядит дороже за счёт чистых линий и аккуратных пропорций — образ получается уверенным и собранным.",
            "Сдержанный премиальный акцент: не спорит с другими аксессуарами, но заметно усиливает общий стиль.",
        ])
    elif style == "social":
        vibe = random.choice([
            "В кадре смотрится эффектно: очки добавляют летний вайб и делают образ более выразительным буквально за секунду.",
            "Хорошо “заходит” в фото и сторис: простой апгрейд образа, который сразу считывается как тренд.",
        ])
    else:
        vibe = random.choice([
            "Универсальный вариант на каждый день: легко сочетать с базовой одеждой и не думать, подходит ли под образ.",
            "Практично и удобно: можно носить целый день, и при этом выглядеть аккуратно и актуально.",
        ])

    # Финал — SEO мягко в смысле, не списком
    # Собираем ключи в 1–2 предложения, чтобы не выглядело “набитым”
    core_str = ", ".join(kw["core"])
    tail_str = ", ".join(kw["tail"])
    feat_str = ""
    if kw["features"]:
        feat_str = f" Часто такие модели ищут по запросу “{kw['features'][0]}”."
    extra_str = ""
    if kw["extra"]:
        extra_str = f" Подойдёт как вариант {kw['extra'][0]} — в зависимости от образа и посадки."

    tail = (
        f"Подходит для {scen}. Если в поиске нужны {core_str} и {tail_str}, здесь это совпадает с реальным удобством, а не только с картинкой."
        f"{feat_str}{extra_str}"
    )

    # Блоки по длине
    # short:  opener + design + lenses + tail (+ иногда vibe)
    # medium: opener + design + lenses + (fact/season) + vibe + tail
    # long:   всё + ещё 1 доп. фраза в середине
    parts = [random.choice(openers), design, lenses]

    if desc_length in {"medium", "long"}:
        if fact and random.random() < 0.9:
            parts.append(fact)
        if season and random.random() < 0.85:
            parts.append(season)
        parts.append(vibe)

    if desc_length == "long":
        long_extra = random.choice([
            "Хорошо сочетаются с базовой одеждой и летними образами, когда хочется подчеркнуть стиль без перегруза деталями.",
            "Уместны и в городе, и в отпуске: добавляют уверенности и визуально “собирают” образ.",
            "Носить удобно: модель делает акцент на лице и помогает чувствовать себя спокойнее при ярком свете.",
        ])
        parts.append(long_extra)

    parts.append(tail)

    # Перемешиваем середину, чтобы не читалось шаблоном
    mid = parts[1:-1]
    random.shuffle(mid)
    text = " ".join([parts[0]] + mid + [parts[-1]])

    # убираем запрещённые ярлыки (на всякий)
    text = _strip_forbidden(text)

    # анти-повторы стартов (первые 7 слов)
    start = _first_n_words(text, 7)
    tries = 0
    while start in recent_desc_starts and tries < 8:
        # немного перетасуем середину + поменяем вступление
        parts[0] = random.choice(openers)
        mid = parts[1:-1]
        random.shuffle(mid)
        text = " ".join([parts[0]] + mid + [parts[-1]])
        text = _strip_forbidden(text)
        start = _first_n_words(text, 7)
        tries += 1

    # подгоняем в диапазон длины (без резки слов)
    # если слишком длинно — режем аккуратно
    if len(text) > max_len:
        text = _cut_no_word_break(text, max_len)

    # если слишком коротко — добавим короткую “человеческую” фразу (только если есть запас)
    if len(text) < min_len and desc_length != "short":
        add = random.choice([
            "Это тот аксессуар, который легко носить каждый день и который заметно усиливает стиль.",
            "Модель выглядит актуально и не требует сложных сочетаний — надел и пошёл.",
            "Хороший баланс: и про внешний вид, и про комфорт, без лишней показухи.",
        ])
        text = _cut_no_word_break(text + " " + add, max_len)

    # общий максимум
    return _cut_no_word_break(text, min(DESC_MAX_LEN, max_len))


def generate_unique_description(
    brand: str, shape: str, lens: str, collection: str,
    style: str, seo_level: str, desc_length: str,
    prev_desc: List[str], recent_desc_starts: List[str],
    min_jaccard: float = 0.48, tries: int = 18
) -> str:
    """
    Усиленная уникализация:
    - избегаем похожести по Jaccard
    - избегаем одинакового начала (первые 7 слов)
    """
    best = ""
    best_score = 1.0

    for _ in range(tries):
        d = generate_description(
            brand=brand, shape=shape, lens=lens, collection=collection,
            style=style, seo_level=seo_level, desc_length=desc_length,
            recent_desc_starts=recent_desc_starts,
        )
        if not prev_desc:
            return d

        score = max(_jaccard(d, p) for p in prev_desc)
        if score < min_jaccard:
            return d

        if score < best_score:
            best_score = score
            best = d

    return best or generate_description(
        brand, shape, lens, collection, style, seo_level, desc_length, recent_desc_starts
    )


# ==========================
# Excel / WB
# ==========================
def find_header_row_and_cols(ws: Worksheet) -> Tuple[int, int, int]:
    for r in range(1, 16):
        name_col = desc_col = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                lv = v.lower()
                if "наименование" in lv:
                    name_col = c
                if "описание" in lv:
                    desc_col = c
        if name_col and desc_col:
            return r, name_col, desc_col
    raise ValueError("Не найдены колонки Наименование / Описание (первые 15 строк).")


def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str,
    style: str = "neutral",
    progress_callback: Optional[Callable[[int], None]] = None,
    seo_level: str = "normal",      # <-- новое (не ломает GUI)
    desc_length: str = "medium",    # <-- новое (не ломает GUI)
) -> Tuple[str, int]:
    """
    1 строка = 1 карточка.
    Перезаписываем только Наименование/Описание.
    SEO усиливается параметрами seo_level/desc_length (опционально).
    """
    random.seed(time.time())

    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, col_name, col_desc = find_header_row_and_cols(ws)
    start = header_row + 1
    end = ws.max_row

    style, seo_level, desc_length = _clamp_modes(style, seo_level, desc_length)

    prev_desc: List[str] = []
    recent_desc_starts: List[str] = []
    recent_title_starts: List[str] = []
    last_title_slogan = ""

    total = max(1, end - start + 1)
    done = 0

    for r in range(start, end + 1):
        titles = build_titles_6(brand, shape, lens_features)
        title = pick_best_title(titles, last_title_slogan, recent_title_starts)

        # обновим анти-повторы названий
        last_title_slogan = (title.split(" ", 1)[0] if title else "")
        recent_title_starts.append(" ".join(title.split()[:4]).lower())
        if len(recent_title_starts) > 3:
            recent_title_starts.pop(0)

        desc = generate_unique_description(
            brand=brand,
            shape=shape,
            lens=lens_features,
            collection=collection,
            style=style,
            seo_level=seo_level,
            desc_length=desc_length,
            prev_desc=prev_desc,
            recent_desc_starts=recent_desc_starts,
            min_jaccard=0.48 if seo_level != "soft" else 0.52,
            tries=18
        )

        # обновим анти-повторы описаний
        recent_desc_starts.append(_first_n_words(desc, 7))
        if len(recent_desc_starts) > 5:
            recent_desc_starts.pop(0)

        prev_desc.append(desc)

        ws.cell(r, col_name).value = title
        ws.cell(r, col_desc).value = desc

        done += 1
        if progress_callback:
            progress_callback(int(done * 100 / total))

    out = Path(input_xlsx).with_name(Path(input_xlsx).stem + "_FILLED.xlsx")
    wb.save(out)
    return str(out), done
