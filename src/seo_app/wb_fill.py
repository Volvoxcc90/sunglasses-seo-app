from __future__ import annotations

import re
import random
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook


def _clean(s: Optional[str]) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def _desktop_path() -> Path:
    home = Path.home()
    p1 = home / "Desktop"
    if p1.exists():
        return p1
    p2 = home / "OneDrive" / "Desktop"
    if p2.exists():
        return p2
    return home


def _find_header_row_and_cols(ws, needed: List[str], max_rows: int = 20, max_cols: int = 400) -> Tuple[int, Dict[str, int]]:
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
    raise RuntimeError(f"Не нашёл строку заголовков с колонками: {', '.join(needed)}")


def _try_find_col(ws, header: str, header_row: int, max_cols: int = 400) -> Optional[int]:
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


MAX_TITLE_LEN = 60
MAX_DESC_LEN = 2000
VARIANTS = 12

START_PHRASES = ["Солнцезащитные очки", "Солнечные очки"]

LENS_TITLE_NORMALIZE = {
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
    "зеркало": "зеркальные",
    "зеркальные": "зеркальные",
    "антиблик": "антибликовые",
    "антибликовые": "антибликовые",
    "градиент": "градиентные",
    "градиентные": "градиентные",
    "ударопрочные": "ударопрочные",
}

LENS_PRIORITY = [
    "UV400",
    "поляризационные",
    "фотохромные",
    "с откидными линзами",
    "зеркальные",
    "антибликовые",
    "градиентные",
    "ударопрочные",
]

SYN = {
    "style_adj": ["современные", "актуальные", "универсальные", "стильные", "практичные"],
    "verbs": ["подчёркивают", "дополняют", "выделяют", "завершают", "усиливают"],
    "use": ["город", "отпуск", "пляж", "путешествия", "вождение", "прогулки", "активный отдых"],
    "cta": [
        "Подчеркните стиль и защитите зрение в солнечные дни.",
        "Универсальный аксессуар для повседневных образов и поездок.",
        "Комфортный выбор для яркого света и активного ритма.",
        "Практично для города и отдыха: удобно и функционально.",
    ],
    "uv_benefits": [
        "защищают глаза от ультрафиолета",
        "снижают нагрузку на зрение при ярком свете",
        "повышают визуальный комфорт в солнечную погоду",
    ],
    "polar_benefits": [
        "уменьшают блики от воды и асфальта",
        "делают картинку более контрастной",
        "комфортнее при вождении и у воды",
    ],
    "photo_benefits": [
        "адаптируются к уровню освещения",
        "подстраиваются под яркость солнца",
        "удобны при переменной погоде",
    ],
    "flip_benefits": [
        "дают вариативность и практичность",
        "удобны, когда нужен быстрый переключатель",
        "добавляют функциональности в повседневности",
    ],
    "triggers": ["дизайнерские", "молодёжные", "эксклюзивные", "уникальные"],
}

_word_re = re.compile(r"[a-zа-яё0-9]+", re.IGNORECASE)


def _split_features(raw: str) -> List[str]:
    raw = _clean(raw)
    if not raw:
        return []
    parts = re.split(r"[;,/]+", raw)
    return [_clean(p).lower() for p in parts if _clean(p)]


def _normalize_lens_tokens(features: List[str]) -> List[str]:
    normalized: List[str] = []
    for f in features:
        if f in LENS_TITLE_NORMALIZE:
            normalized.append(LENS_TITLE_NORMALIZE[f])
            continue

        f2 = f.replace(" ", "")
        if "uv400" in f2:
            normalized.append("UV400")
        elif "поляр" in f:
            normalized.append("поляризационные")
        elif "фотохром" in f or "хамелеон" in f:
            normalized.append("фотохромные")
        elif "откид" in f:
            normalized.append("с откидными линзами")
        elif "зерк" in f:
            normalized.append("зеркальные")
        elif "антиблик" in f:
            normalized.append("антибликовые")
        elif "градиент" in f:
            normalized.append("градиентные")
        elif "ударопр" in f:
            normalized.append("ударопрочные")

    seen = set()
    out = []
    for x in normalized:
        if x and x not in seen:
            seen.add(x)
            out.append(x)

    pidx = {k: i for i, k in enumerate(LENS_PRIORITY)}
    out.sort(key=lambda x: pidx.get(x, 10_000))
    return out


def _join_words_limit(words: List[str], max_len: int) -> str:
    result: List[str] = []
    for w in words:
        w = _clean(w)
        if not w:
            continue
        cand = " ".join(result + [w])
        if len(cand) <= max_len:
            result.append(w)
    return " ".join(result).strip()


def _balanced_starts(n: int, rng: random.Random) -> List[str]:
    half = n // 2
    starts = [START_PHRASES[0]] * half + [START_PHRASES[1]] * (n - half)
    rng.shuffle(starts)
    return starts


def _choose_trigger_indices(rng: random.Random) -> set:
    k = rng.choice([3, 4])
    return set(rng.sample(range(1, VARIANTS + 1), k))


def _build_12_templates_titles(brand: str, shape: str, lens_raw: str, rng: random.Random) -> List[str]:
    brand = _clean(brand)
    shape = _clean(shape)
    lens_tokens = _normalize_lens_tokens(_split_features(lens_raw))
    main_lens = lens_tokens[0] if lens_tokens else "UV400"
    second_lens = lens_tokens[1] if len(lens_tokens) > 1 else ""

    starts = _balanced_starts(VARIANTS, rng)

    structures = [
        lambda start: [start, brand, main_lens],
        lambda start: [start, brand, main_lens, shape],
        lambda start: [start, brand, shape, main_lens],
        lambda start: [start, main_lens, brand, shape],
        lambda start: [start, main_lens, brand],
        lambda start: [start, brand, second_lens or main_lens, shape],
        lambda start: [start, brand, main_lens, second_lens, shape],
        lambda start: [start, second_lens or main_lens, brand],
        lambda start: [start, brand, shape, second_lens or main_lens],
        lambda start: [start, second_lens or main_lens, brand, shape],
        lambda start: [start, brand, second_lens or main_lens],
        lambda start: [start, brand, shape, main_lens, second_lens],
    ]

    titles: List[str] = []
    for i in range(VARIANTS):
        start = starts[i]
        words = [w for w in structures[i](start) if _clean(w)]
        t = _join_words_limit(words, MAX_TITLE_LEN)
        if brand not in t or main_lens not in t:
            t = _join_words_limit([start, brand, main_lens, shape], MAX_TITLE_LEN)
        titles.append(t)

    return titles


def _pick_usecases(rng: random.Random) -> str:
    k = rng.randint(3, 5)
    items = rng.sample(SYN["use"], k=min(k, len(SYN["use"])))
    return ", ".join(items)


def _lens_sentence(lens_tokens: List[str], rng: random.Random) -> str:
    if not lens_tokens:
        return "защита от ультрафиолета"
    head = lens_tokens[:]
    if len(head) > 2 and rng.random() < 0.35:
        i = rng.randint(0, 2)
        j = rng.randint(0, 2)
        head[i], head[j] = head[j], head[i]
    return ", ".join(head[:3])


def _keywords(brand: str, shape: str, lens_tokens: List[str], collection: str) -> str:
    kw = ["солнцезащитные очки", "солнечные очки", f"{brand} очки"]
    if shape:
        kw.append(f"очки {shape}")
    for lt in lens_tokens[:2]:
        kw.append(f"очки {lt}")
    if "UV400" in lens_tokens:
        kw.append("очки UV400")
    if collection:
        kw.append(f"коллекция {collection}")

    seen = set()
    out = []
    for x in kw:
        x = _clean(x)
        if x and x.lower() not in seen:
            seen.add(x.lower())
            out.append(x)
    return ", ".join(out)


def _enforce_max_desc(text: str) -> str:
    text = _clean(text)
    if len(text) <= MAX_DESC_LEN:
        return text
    parts = re.split(r"(?<=[.!?])\s+|\n+", text)
    out: List[str] = []
    for p in parts:
        p = _clean(p)
        if not p:
            continue
        cand = (" ".join(out + [p])).strip()
        if len(cand) <= MAX_DESC_LEN:
            out.append(p)
        else:
            break
    return _clean(" ".join(out))[:MAX_DESC_LEN].rstrip()


def _tokenize(text: str) -> set:
    return set(_word_re.findall(text.lower()))


def _jaccard(a: str, b: str) -> float:
    ta = _tokenize(a)
    tb = _tokenize(b)
    if not ta and not tb:
        return 0.0
    return len(ta & tb) / len(ta | tb)


def _build_desc_variant(
    brand: str,
    shape: str,
    lens_raw: str,
    collection: str,
    title: str,
    rng: random.Random,
    add_trigger: bool
) -> str:
    brand = _clean(brand)
    shape = _clean(shape)
    collection = _clean(collection) or "Весна–Лето 2025–2026"

    lens_tokens = _normalize_lens_tokens(_split_features(lens_raw))
    lens_sent = _lens_sentence(lens_tokens, rng)
    use = _pick_usecases(rng)
    kw = _keywords(brand, shape, lens_tokens, collection)

    g = START_PHRASES[0] if title.startswith(START_PHRASES[0]) else START_PHRASES[1]

    blocks = [
        f"{g} {brand} — {rng.choice(SYN['style_adj'])} вариант из коллекции {collection}. Форма оправы {shape} {rng.choice(SYN['verbs'])} образ.",
        f"Линзы: {lens_sent}. Они {rng.choice(SYN['uv_benefits'])}.",
        f"Сценарии: {use}. Удобно в яркую погоду — акцент на {rng.choice(['комфорте', 'защите', 'практичности'])}.",
        f"Посадка продумана: {rng.choice(['удобно', 'комфортно', 'приятно'])} в течение дня.",
        f"Силуэт {shape} сочетается с {rng.choice(['повседневной', 'городской', 'курортной'])} одеждой.",
        f"Сезонность: {collection}.",
        f"Ключевые слова: {kw}.",
    ]

    rng.shuffle(blocks)
    paras = blocks[:rng.randint(4, 6)]

    if add_trigger:
        trig = rng.choice(SYN["triggers"])
        paras.insert(0, f"По подаче это {trig} решение, но при этом практичное.")

    if rng.random() < 0.7:
        paras.insert(2, (
            "Характеристики:\n"
            f"• Бренд: {brand}\n"
            f"• Форма оправы: {shape}\n"
            f"• Линзы: {lens_sent}\n"
            f"• Коллекция: {collection}\n"
        ).rstrip())

    return _enforce_max_desc("\n\n".join(_clean(p) for p in paras if _clean(p)))


def _generate_12_pairs_fixed_params(brand: str, shape: str, lens_raw: str, collection: str, seed: int) -> List[Tuple[str, str]]:
    rng = random.Random(seed)
    titles = _build_12_templates_titles(brand, shape, lens_raw, rng)
    trigger_idx = _choose_trigger_indices(rng)

    descs: List[str] = []
    pairs: List[Tuple[str, str]] = []

    for i in range(1, VARIANTS + 1):
        title = titles[i - 1]
        vrng = random.Random(seed * 1000 + i * 97)
        add_trigger = i in trigger_idx

        best_desc = None
        for attempt in range(1, 9):
            desc = _build_desc_variant(brand, shape, lens_raw, collection, title, vrng, add_trigger)
            if not descs or all(_jaccard(desc, prev) <= 0.58 for prev in descs):
                best_desc = desc
                break
            vrng = random.Random(seed * 1000 + i * 97 + attempt * 10007)

        if best_desc is None:
            best_desc = desc

        descs.append(best_desc)
        pairs.append((title, best_desc))

    return pairs


def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str
) -> Tuple[str, int]:
    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, cols = _find_header_row_and_cols(ws, ["Наименование", "Описание"], max_rows=20, max_cols=400)
    col_name = cols["Наименование"]
    col_desc = cols["Описание"]

    col_photo = _try_find_col(ws, "Фото", header_row)
    col_vendor = _try_find_col(ws, "Артикул продавца", header_row)
    col_barcodes = _try_find_col(ws, "Баркоды", header_row)

    signal_cols = [col_name, col_desc]
    for c in [col_photo, col_vendor, col_barcodes]:
        if c is not None:
            signal_cols.append(c)
    if len(signal_cols) <= 2:
        signal_cols = list(range(1, 30))

    pairs12 = _generate_12_pairs_fixed_params(brand, shape, lens_features, collection, seed=1234567)

    filled = 0
    variant_counter = 0

    for r in range(header_row + 1, ws.max_row + 1):
        if not _row_is_product(ws, r, signal_cols):
            continue

        idx = variant_counter % VARIANTS
        title, desc = pairs12[idx]

        ws.cell(r, col_name).value = title
        ws.cell(r, col_desc).value = desc

        filled += 1
        variant_counter += 1

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    in_path = Path(input_xlsx)
    out_name = f"{in_path.stem}_filled_{ts}.xlsx"
    out_path = _desktop_path() / out_name
    wb.save(out_path)

    return str(out_path), filled
