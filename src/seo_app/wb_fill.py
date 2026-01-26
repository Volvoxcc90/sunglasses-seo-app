from __future__ import annotations

import re
import random
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# ======================
# Basic helpers
# ======================

_WORD_RE = re.compile(r"[a-zа-яё0-9]+", re.IGNORECASE)


def _clean(s: Optional[str]) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


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
    # Товарная строка, если есть данные в любой из сигнальных колонок
    for c in signal_cols:
        v = ws.cell(r, c).value
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        return True
    return False


def _write_safe(ws, row: int, col: int, value: str) -> None:
    """
    WB-шаблоны часто имеют merged cells. В MergedCell писать нельзя.
    Если целевая ячейка merged — пишем в верхнюю-левую ячейку диапазона.
    """
    cell = ws.cell(row, col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if cell.coordinate in mr:
                ws.cell(mr.min_row, mr.min_col).value = value
                return
        return
    cell.value = value


def _tokenize(text: str) -> set:
    return set(_WORD_RE.findall(text.lower()))


def _jaccard(a: str, b: str) -> float:
    ta = _tokenize(a)
    tb = _tokenize(b)
    if not ta and not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


# ======================
# SEO text engine (v4+)
# ======================

MAX_TITLE_LEN = 60
MAX_DESC_LEN = 2000
MIN_DESC_TARGET = 1200
VARIANTS = 12

START_PHRASES = ["Солнцезащитные очки", "Солнечные очки"]

# Нормализация линз (чтобы “хамелеон/фотохром” приводить к одной форме)
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
    "intro_hooks": [
        "трендовый аксессуар на лето",
        "актуальный акцент сезона",
        "выразительная деталь образа",
        "элегантный штрих к повседневным и вечерним лукам",
        "яркое дополнение к гардеробу",
    ],
    "luxury_phrases": [
        "роскошь и индивидуальность",
        "изысканный дизайн и мастерство",
        "баланс трендов и классики",
        "продуманная эстетика деталей",
        "выразительная геометрия линий",
    ],
    "benefits_uv": [
        "надежно защищают глаза от ультрафиолета",
        "помогают снизить зрительную нагрузку при ярком свете",
        "делают прогулки в солнечную погоду заметно комфортнее",
    ],
    "benefits_polar": [
        "уменьшают блики от воды и асфальта",
        "повышают контраст и читаемость изображения",
        "особенно комфортны при вождении и у воды",
    ],
    "benefits_photo": [
        "адаптируются к уровню освещения",
        "подстраиваются под яркость солнца в течение дня",
        "удобны при переменной погоде",
    ],
    "benefits_flip": [
        "дают дополнительную функциональность",
        "удобны, когда нужен быстрый вариант переключения",
        "подходят тем, кто любит практичные решения",
    ],
    "scenarios": [
        "город", "путешествия", "отпуск", "пляж", "прогулки", "вождение", "активный отдых"
    ],
    "tone": [
        "Эти очки — не просто защита, а выразительный элемент стиля.",
        "Это аксессуар, который мгновенно собирает образ.",
        "Модель выглядит эффектно в кадре и органично в жизни.",
        "Форма оправы подчёркивает черты лица и добавляет уверенности.",
    ],
    # 3–4 из 12 (не во все) — мягкие триггеры
    "triggers": ["дизайнерские", "эксклюзивные", "уникальные", "молодёжные"],
    # мягкие “ключи”, без мусорной метки SEO
    "keywords_tail": [
        "очки солнцезащитные женские",
        "очки солнцезащитные мужские",
        "очки cat eye",
        "очки кошачий глаз",
        "брендовые очки",
        "тренд из instagram",
        "инста очки",
        "очки из tiktok",
        "солнечные очки",
        "солнцезащитные очки",
    ]
}


def _split_features(raw: str) -> List[str]:
    raw = _clean(raw)
    if not raw:
        return []
    parts = re.split(r"[;,/]+", raw)
    return [_clean(p).lower() for p in parts if _clean(p)]


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

    # de-dup
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
    Собирает заголовок <= 60 символов.
    Не режет слова: если не помещается — удаляет хвостовые части.
    """
    parts = [_clean(p) for p in parts if _clean(p)]
    while parts:
        s = " ".join(parts)
        if len(s) <= MAX_TITLE_LEN:
            return s
        parts.pop()  # убираем последний блок
    return ""


def _balanced_starts(rng: random.Random) -> List[str]:
    half = VARIANTS // 2
    starts = [START_PHRASES[0]] * half + [START_PHRASES[1]] * (VARIANTS - half)
    rng.shuffle(starts)
    return starts


def _gen_titles_12(brand: str, shape: str, lens_raw: str, collection: str, rng: random.Random) -> List[str]:
    brand = _clean(brand)
    shape = _clean(shape)
    collection = _clean(collection)

    lens_tokens = _normalize_lens_tokens(lens_raw)
    main_lens = lens_tokens[0] if lens_tokens else "UV400"
    second = lens_tokens[1] if len(lens_tokens) > 1 else ""

    starts = _balanced_starts(rng)

    # 12 разных структур (смысловые, без мусора)
    structures = [
        lambda s: [s, brand, main_lens],
        lambda s: [s, brand, main_lens, shape],
        lambda s: [s, brand, shape, main_lens],
        lambda s: [s, main_lens, brand, shape],
        lambda s: [s, brand, f"{main_lens} {second}".strip()],
        lambda s: [s, brand, shape],
        lambda s: [s, brand, main_lens, collection],
        lambda s: [s, brand, shape, collection],
        lambda s: [s, brand, second or main_lens],
        lambda s: [s, main_lens, brand],
        lambda s: [s, brand, main_lens, "2026"],
        lambda s: [s, brand, shape, "2026"],
    ]

    titles: List[str] = []
    for i in range(VARIANTS):
        t = _fit_title(structures[i](starts[i]))
        # гарантируем: старт + бренд
        if not t.startswith(starts[i]) or brand not in t:
            t = _fit_title([starts[i], brand, main_lens, shape])
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
            t2 = _fit_title([alt_start, brand, main_lens, shape])
            out.append(t2 if t2 and t2 not in seen else t)
            seen.add(out[-1])
    return out[:VARIANTS]


def _pick_trigger_indices(rng: random.Random) -> set:
    # 3–4 из 12
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
    # красиво собрать
    if len(parts) == 1:
        return parts[0]
    if len(parts) == 2:
        return parts[0] + ", а также " + parts[1]
    return parts[0] + ", " + parts[1] + " и " + parts[2]


def _keywords_tail(gender_hint: str, rng: random.Random) -> str:
    # аккуратная строка в конце, без "SEO:"
    pool = SYN["keywords_tail"][:]
    if gender_hint == "жен":
        pool = ["очки солнцезащитные женские"] + pool
    elif gender_hint == "муж":
        pool = ["очки солнцезащитные мужские"] + pool

    picked = rng.sample(pool, k=6)
    # уберём дубли в строке
    uniq = []
    seen = set()
    for x in picked:
        x = _clean(x)
        if x.lower() not in seen:
            seen.add(x.lower())
            uniq.append(x)
    return ", ".join(uniq)


def _detect_gender_hint(title_or_brand: str) -> str:
    t = (title_or_brand or "").lower()
    if "жен" in t:
        return "жен"
    if "муж" in t:
        return "муж"
    return ""


def _build_desc(
    brand: str,
    shape: str,
    lens_raw: str,
    collection: str,
    title: str,
    seed: int,
    add_trigger: bool
) -> str:
    rng = random.Random(seed)

    brand = _clean(brand)
    shape = _clean(shape)
    collection = _clean(collection) or "лето 2026"

    lens_tokens = _normalize_lens_tokens(lens_raw)
    lens_line = ", ".join(lens_tokens[:3]) if lens_tokens else "UV400"
    hook = rng.choice(SYN["intro_hooks"])

    # 6–10 блоков, потом берём 5–7, чтобы было длинно и по-разному
    scenarios = ", ".join(rng.sample(SYN["scenarios"], k=4))
    luxury = rng.choice(SYN["luxury_phrases"])
    tone_line = rng.choice(SYN["tone"])
    benefits = _lens_benefits(lens_tokens, rng)

    trigger_line = ""
    if add_trigger:
        trig = rng.choice(SYN["triggers"])
        trigger_line = rng.choice([
            f"{brand} — {trig} подача для тех, кто ценит стиль и детали.",
            f"Эту пару часто выбирают как {trig} акцент на сезон.",
            f"По ощущению это {trig} аксессуар, который легко собирает образ.",
        ])

    blocks = [
        f"Очки {brand} — {hook} {collection} года. {tone_line}",
        trigger_line,
        f"Встречайте стильные брендовыми очки {brand}: они не только дополняют образ, но и помогают чувствовать себя уверенно при ярком солнце.",
        f"Форма оправы {shape} выглядит актуально и подчёркивает индивидуальность — от повседневных луков до более смелых стилизаций.",
        f"Линзы: {lens_line}. Они {benefits}.",
        f"Дизайн {brand} — это {luxury}: выразительные линии, гармоничные пропорции и внимание к деталям, которые заметны с первого взгляда.",
        f"Сценарии: {scenarios}. Модель особенно удобна в городе и в поездках, когда важны комфорт и защита зрения.",
        "Характеристики:\n"
        f"• Бренд: {brand}\n"
        f"• Форма оправы: {shape}\n"
        f"• Линзы: {lens_line}\n"
        f"• Коллекция: {collection}",
        f"Если нужен акцент в кадре — такие солнечные очки смотрятся эффектно и в Instagram, и в TikTok, при этом остаются практичными.",
    ]

    # перемешаем подачу
    blocks = [b for b in blocks if _clean(b)]
    rng.shuffle(blocks)

    # соберём 6–7 блоков, чтобы выйти на 1200–2000 символов
    count = rng.randint(6, 7)
    chosen = blocks[:count]

    gender_hint = _detect_gender_hint(title + " " + brand)
    tail = _keywords_tail(gender_hint, rng)

    # аккуратно: не "SEO:", а “Ключевые фразы:”
    chosen.append(f"Ключевые фразы: {tail}.")

    text = "\n\n".join(_clean(x) for x in chosen if _clean(x))
    text = _clean(text)

    # добивка до MIN_DESC_TARGET (если вдруг коротко) — добавим ещё 1–2 блока
    if len(text) < MIN_DESC_TARGET:
        extra = [
            f"Цветовые решения и детали {brand} помогают подчеркнуть характер образа — от лаконичных сочетаний до ярких акцентов.",
            f"Выбирайте солнцезащитные очки {brand} как подарок или как обновление гардероба к тёплому сезону: стильно, удобно и функционально.",
        ]
        rng.shuffle(extra)
        for e in extra:
            cand = text + "\n\n" + e
            if len(cand) <= MAX_DESC_LEN:
                text = cand

    # ограничение 2000
    if len(text) > MAX_DESC_LEN:
        text = text[:MAX_DESC_LEN]
        # не резать слово/строку грубо
        text = text.rsplit(" ", 1)[0].rstrip(".,;:") + "."

    return text


def _generate_12_pairs(brand: str, shape: str, lens_raw: str, collection: str, seed: int) -> List[Tuple[str, str]]:
    rng = random.Random(seed)
    titles = _gen_titles_12(brand, shape, lens_raw, collection, rng)
    trigger_idx = _pick_trigger_indices(rng)

    pairs: List[Tuple[str, str]] = []
    descs: List[str] = []

    for i in range(VARIANTS):
        title = titles[i]
        add_trigger = i in trigger_idx

        # регенерация, если слишком похоже на предыдущие
        best = None
        for attempt in range(1, 9):
            desc = _build_desc(
                brand=brand,
                shape=shape,
                lens_raw=lens_raw,
                collection=collection,
                title=title,
                seed=seed * 1000 + i * 97 + attempt * 10007,
                add_trigger=add_trigger
            )
            if not descs:
                best = desc
                break
            if all(_jaccard(desc, prev) <= 0.58 for prev in descs):
                best = desc
                break
        if best is None:
            best = desc

        descs.append(best)
        pairs.append((title, best))

    return pairs


# ======================
# Main entry
# ======================

def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str
) -> Tuple[str, int]:
    """
    Заполняет только "Наименование" и "Описание".
    Не зависит от остальных колонок.
    Умеет писать в merged-cells безопасно.
    Сохраняет новый файл рядом с исходником.
    """
    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, cols = _find_header_row_and_cols(ws, ["Наименование", "Описание"])
    col_name = cols["Наименование"]
    col_desc = cols["Описание"]

    # сигнальные колонки, чтобы понять, где реально товарные строки
    signal_cols = [col_name, col_desc]
    for h in ["Фото", "Артикул продавца", "Баркоды", "Размер", "Цвет"]:
        c = _try_find_col(ws, h, header_row)
        if c:
            signal_cols.append(c)
    if len(signal_cols) <= 2:
        signal_cols = list(range(1, 30))

    pairs12 = _generate_12_pairs(
        brand=_clean(brand),
        shape=_clean(shape),
        lens_raw=_clean(lens_features),
        collection=_clean(collection) or "лето 2026",
        seed=1234567
    )

    filled = 0
    counter = 0

    for r in range(header_row + 1, ws.max_row + 1):
        if not _row_is_product(ws, r, signal_cols):
            continue

        title, desc = pairs12[counter % VARIANTS]

        # микро-уникализация после 12 строк (если строк больше)
        if counter >= VARIANTS:
            mrng = random.Random(900000 + r * 17 + counter * 101)
            if mrng.random() < 0.35:
                desc = desc.replace("Встречайте", "Оцените", 1)
            if mrng.random() < 0.35:
                desc = desc.replace("стильные", "эффектные", 1)

        _write_safe(ws, r, col_name, title)
        _write_safe(ws, r, col_desc, desc)

        filled += 1
        counter += 1

    in_path = Path(input_xlsx)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = in_path.parent / f"{in_path.stem}_filled_{ts}.xlsx"
    wb.save(out_path)

    return str(out_path), filled
