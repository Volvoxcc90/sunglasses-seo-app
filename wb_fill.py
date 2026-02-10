# wb_fill.py
import json
import os
import random
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.cell_range import MultiCellRange

TITLE_MAX = 60
DESC_MAX = 2000

# -----------------------------
# GLOBALS (для пачки файлов)
# -----------------------------
_GLOBAL_USED_SLOGANS: Set[str] = set()
_GLOBAL_USED_TITLE_SIGS: Set[str] = set()


# -----------------------------
# DATA MODEL (для совместимости)
# -----------------------------
@dataclass
class FillParams:
    input_xlsx: str
    brand_lat: str
    shape: str
    lens: str
    collection: str

    style: str = "neutral"
    desc_length: str = "medium"
    seo_level: str = "high"
    gender_mode: str = "Auto"

    wb_safe_mode: bool = True
    wb_strict: bool = True

    uniq_strength: int = 90
    brand_in_title_mode: str = "smart50"
    data_dir: str = ""

    max_fill_rows: int = 6
    output_index: int = 1
    output_total: int = 1
    between_files_slogan_lock: bool = True


# -----------------------------
# POOLS
# -----------------------------
SLOGANS = [
    "Красивые","Крутые","Стильные","Модные","Молодёжные","Дизайнерские","Эффектные","Трендовые","Лаконичные","Яркие",
    "Современные","Премиальные","Универсальные","Актуальные","Выразительные","Элегантные","Смелые","Классные","Городские",
    "Лёгкие","Комфортные","Популярные","Эксклюзивные","Фирменные","Изящные","Брутальные","Шикарные","Статусные",
    "Инстаграмные","Фотогеничные","Сочные","Практичные","Надёжные","Удобные","Качественные","В тренде сезона",
    "Собирают образ","Делают образ дороже","Сочетаются легко","Выглядят дорого"
]

SUN_TERMS = ["солнцезащитные очки", "солнечные очки"]
SEO_CORE = ["солнцезащитные очки", "солнечные очки", "очки солнцезащитные"]
SEO_STYLE = ["брендовые очки", "модные очки", "трендовые очки", "стильные очки", "имиджевые очки"]
SEO_SOC = ["инста очки", "очки из tiktok", "очки для фото"]
SEO_USE = ["очки для вождения", "очки для города", "очки для отпуска", "очки для пляжа", "очки для прогулок"]

SCENARIOS = [
    "город","прогулки","поездки","путешествия","отпуск","пляж","вождение","на каждый день","для фото","летние мероприятия"
]

# Слова/метки, которые НЕЛЬЗЯ видеть в описании
FORBIDDEN_LABELS_RE = re.compile(
    r"\b(Сценарии|Ключевые\s*слова|Форма|Линза|Линзы|Коллекция)\s*:\s*",
    flags=re.IGNORECASE
)

STOPWORDS_RU = {
    "и","в","во","на","а","но","что","это","как","для","по","из","к","с","со","при","от","до","у","же","не","без","над","под","про","или","то","ли"
}

# -----------------------------
# UTILS
# -----------------------------
def _seed_hard():
    # сильный seed каждый запуск/файл
    random.seed(int.from_bytes(os.urandom(16), "big"))

def _cut_no_break_words(text: str, limit: int) -> str:
    t = (text or "").strip()
    if len(t) <= limit:
        return t
    cut = t[:limit]
    if " " not in cut:
        return cut
    return cut.rsplit(" ", 1)[0].strip()

def _cap_first(text: str) -> str:
    t = (text or "").strip()
    if not t:
        return t
    # убираем невидимые/мусорные символы в начале
    t = re.sub(r"^[\s\-\–\—\•\·\.\,]+", "", t).strip()
    if not t:
        return ""
    return t[0].upper() + t[1:]

def normalize_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("-", " ").replace("&", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def load_json(path: Path) -> dict:
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def load_brands_ru_map(data_dir: str) -> Dict[str, str]:
    p = Path(data_dir) / "brands_ru.json"
    return load_json(p)

def brand_ru(brand_lat: str, brand_map: Dict[str, str]) -> str:
    key = normalize_key(brand_lat)
    return (brand_map.get(key) or brand_lat).strip()

def _normalize_plain(text: str) -> str:
    t = (text or "").lower()
    t = FORBIDDEN_LABELS_RE.sub("", t)
    t = re.sub(r"[^a-zа-яё0-9\s\-]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def _tokens(text: str) -> Set[str]:
    parts = _normalize_plain(text).split()
    return {w for w in parts if len(w) >= 3 and w not in STOPWORDS_RU}

def jaccard(a: str, b: str) -> float:
    A = _tokens(a); B = _tokens(b)
    if not A or not B:
        return 0.0
    return len(A & B) / max(1, len(A | B))

def uniqueness_threshold(uniq_strength: int) -> float:
    # 60..95 => 0.78..0.52 (агрессивнее)
    s = max(60, min(95, int(uniq_strength)))
    return 0.78 - (s - 60) * (0.26 / 35.0)

def _desc_signature(text: str) -> str:
    t = _normalize_plain(text)
    words = t.split()
    pref = " ".join(words[:22])
    # top words
    freq: Dict[str, int] = {}
    for w in words:
        if len(w) < 4 or w in STOPWORDS_RU:
            continue
        freq[w] = freq.get(w, 0) + 1
    top = " ".join([w for w, _ in sorted(freq.items(), key=lambda x: (-x[1], x[0]))[:6]])
    return (pref + " | " + top).strip()

def _pick_slogan(pool: List[str], lock_between_files: bool) -> str:
    global _GLOBAL_USED_SLOGANS

    if not pool:
        pool.extend(SLOGANS)
        random.shuffle(pool)

    if not lock_between_files:
        return pool.pop()

    # ищем неиспользованный между файлами
    for _ in range(len(pool) * 2):
        s = pool.pop()
        if s not in _GLOBAL_USED_SLOGANS:
            _GLOBAL_USED_SLOGANS.add(s)
            return s
        pool.insert(0, s)

    _GLOBAL_USED_SLOGANS.clear()
    random.shuffle(pool)
    s = pool.pop()
    _GLOBAL_USED_SLOGANS.add(s)
    return s

def _fix_merged_cells(ws):
    try:
        if isinstance(ws.merged_cells, MultiCellRange):
            return
        old = ws.merged_cells
        fixed = MultiCellRange()
        for r in list(old):
            fixed.add(str(r))
        ws.merged_cells = fixed
    except Exception:
        pass

def _norm_header(x: str) -> str:
    s = str(x).strip().lower()
    s = s.replace("ё", "е")
    s = re.sub(r"\s+", " ", s)
    return s

def find_header_col(ws, candidates: Set[str], header_scan_rows: int = 30) -> Tuple[Optional[int], Optional[int]]:
    cand = {c.lower() for c in candidates}
    for r in range(1, header_scan_rows + 1):
        for cell in ws[r]:
            if cell.value is None:
                continue
            val = _norm_header(cell.value)
            if val in cand:
                return cell.column, r
    for r in range(1, header_scan_rows + 1):
        for cell in ws[r]:
            if cell.value is None:
                continue
            val = _norm_header(cell.value)
            for c in cand:
                if c in val:
                    return cell.column, r
    return None, None

def read_output_dir(data_dir: str) -> Optional[str]:
    try:
        p = Path(data_dir) / "output_dir.txt"
        if not p.exists():
            return None
        out = p.read_text(encoding="utf-8").strip()
        if not out:
            return None
        d = Path(out)
        d.mkdir(parents=True, exist_ok=True)
        return str(d)
    except Exception:
        return None


# -----------------------------
# TITLE GENERATOR
# -----------------------------
def generate_title(
    brand_lat: str,
    shape: str,
    lens: str,
    brand_map: Dict[str, str],
    brand_in_title_mode: str,
    slogan_pool: List[str],
    lock_between_files: bool,
) -> str:
    b_ru = brand_ru(brand_lat, brand_map)
    slogan = _pick_slogan(slogan_pool, lock_between_files)
    sun = random.choice(SUN_TERMS)

    if brand_in_title_mode == "always":
        put_brand = True
    elif brand_in_title_mode == "never":
        put_brand = False
    else:
        put_brand = (random.random() < 0.5)

    parts = [slogan, sun]

    # бренд кириллицей только в названии (как ты просил)
    if put_brand and b_ru:
        parts.append(b_ru)

    # добавки (форма/линза) — но не одинаково каждый раз
    if shape and random.random() < 0.70:
        parts.append(shape)
    if lens and random.random() < 0.60:
        parts.append(lens)

    title = " ".join([p for p in parts if p]).strip()
    title = re.sub(r"\s{2,}", " ", title)
    title = _cap_first(title)
    return _cut_no_break_words(title, TITLE_MAX)


# -----------------------------
# “НАРОДНОЕ” ОПИСАНИЕ (SEO)
# -----------------------------
def build_desc(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
    variant_id: int,
) -> Tuple[str, str]:
    """
    Возвращает (description, struct_key)
    variant_id позволяет сделать 6 реально разных “подач”.
    """

    # SEO плотность
    if seo_level == "low":
        core_n, style_n, use_n, soc_n = 1, 1, 1, 0
    elif seo_level == "normal":
        core_n, style_n, use_n, soc_n = 2, 2, 2, 1
    else:
        core_n, style_n, use_n, soc_n = 3, 3, 3, 2

    core = random.sample(SEO_CORE, k=min(core_n, len(SEO_CORE)))
    style = random.sample(SEO_STYLE, k=min(style_n, len(SEO_STYLE)))
    use = random.sample(SEO_USE, k=min(use_n, len(SEO_USE)))
    soc = random.sample(SEO_SOC, k=min(soc_n, len(SEO_SOC))) if soc_n else []

    if gender_mode == "Женские":
        gender_kw = "очки женские"
    elif gender_mode == "Мужские":
        gender_kw = "очки мужские"
    elif gender_mode == "Унисекс":
        gender_kw = "очки унисекс"
    else:
        gender_kw = random.choice(["очки женские", "очки мужские", "очки унисекс"])

    scen = ", ".join(random.sample(SCENARIOS, 4))

    # Фразы (разные стили “как маркетплейсы”)
    openers = [
        f"{brand_lat} — очки, которые реально выручают в солнце и сразу собирают образ.",
        f"Если нужны {random.choice(core)} на каждый день — {brand_lat} прям в тему.",
        f"{brand_lat}: удобные {random.choice(core)}, которые выглядят аккуратно и дорого.",
        f"Очки {brand_lat} — стильный акцент на сезон, который легко носить каждый день.",
    ]
    shape_line = [
        f"Оправа {shape} — смотрится ровно, подчёркивает лицо и не выглядит громоздко.",
        f"Форма {shape} хорошо садится и подходит под базовый гардероб.",
        f"{shape} — трендовая форма, которая делает образ более собранным.",
    ] if shape else [
        "Оправа выглядит аккуратно и легко сочетается с одеждой.",
        "Посадка комфортная — можно носить целый день.",
    ]

    lens_line = [
        f"Линзы {lens} дают комфорт при ярком солнце — глаза меньше устают.",
        f"{lens} — хороший вариант для города и поездок, когда на улице ярко.",
        f"С линзами {lens} реально удобнее: меньше лишнего света и бликов.",
    ] if lens else [
        "Линзы комфортные в солнечную погоду.",
        "В яркий день носить удобно и спокойно.",
    ]

    coll_line = [
        f"Сезон {collection}: модель выглядит актуально и подходит под летний стиль.",
        f"На сезон {collection} — отличный вариант, чтобы обновить аксессуары.",
    ] if collection else []

    # “ключи” не списком, а “вшиты”
    seo_line = [
        f"Ищут так: {', '.join(core[:1] + style[:1] + [gender_kw])}.",
        f"По запросам: {', '.join([random.choice(core), random.choice(style), random.choice(use)])}.",
        f"Под фото и сторис: {', '.join(style[:1] + soc[:1] if soc else style[:1])}.",
    ]

    close_line = [
        f"Подойдут для: {scen}. Можно брать себе или на подарок.",
        f"Норм вариант и в город, и в отпуск: {scen}. Отличный подарок тоже.",
        f"Под ежедневный стиль и поездки: {scen}. Берут себе и как подарок.",
    ]

    # 6 разных структур, чтобы не было “3 одинаковых”
    structures = [
        ("A", [random.choice(openers), random.choice(shape_line), random.choice(lens_line)] + coll_line + [random.choice(close_line), random.choice(seo_line)]),
        ("B", [random.choice(openers), random.choice(lens_line), random.choice(shape_line), random.choice(seo_line), random.choice(close_line)] + coll_line),
        ("C", [random.choice(openers)] + coll_line + [random.choice(shape_line), random.choice(close_line), random.choice(seo_line), random.choice(lens_line)]),
        ("D", [random.choice(openers), random.choice(seo_line), random.choice(shape_line), random.choice(lens_line), random.choice(close_line)] + coll_line),
        ("E", [random.choice(openers), random.choice(close_line), random.choice(shape_line)] + coll_line + [random.choice(lens_line), random.choice(seo_line)]),
        ("F", [random.choice(openers), random.choice(shape_line), random.choice(seo_line), random.choice(close_line), random.choice(lens_line)] + coll_line),
    ]

    struct_key, parts = structures[variant_id % len(structures)]

    text = " ".join([p.strip() for p in parts if p and p.strip()]).strip()
    text = FORBIDDEN_LABELS_RE.sub("", text)  # ещё раз вычищаем метки
    text = re.sub(r"\s{2,}", " ", text).strip()
    text = _cap_first(text)

    return _cut_no_break_words(text, DESC_MAX), struct_key


def generate_unique_descs(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
    uniq_strength: int,
    need: int,
) -> List[str]:
    """
    Генерирует ровно need описаний:
      - разные структуры (A..F)
      - подпись + Jaccard против уже сгенеренных
      - запрещаем повтор одинаковых “подписей”
    """
    thr = uniqueness_threshold(uniq_strength)
    used: List[str] = []
    used_sigs: Set[str] = set()
    used_structs: Set[str] = set()

    for i in range(need):
        best = None
        best_mx = 1.0

        for _ in range(160):
            cand, struct = build_desc(brand_lat, shape, lens, collection, seo_level, gender_mode, variant_id=i)
            sig = _desc_signature(cand)

            if sig in used_sigs:
                continue

            # на 6 строк — стараемся держать разные структуры
            if struct in used_structs and len(used_structs) < need:
                continue

            if not used:
                used.append(cand)
                used_sigs.add(sig)
                used_structs.add(struct)
                best = cand
                best_mx = 0.0
                break

            mx = max(jaccard(cand, prev) for prev in used)
            if mx <= thr:
                used.append(cand)
                used_sigs.add(sig)
                used_structs.add(struct)
                best = cand
                best_mx = mx
                break

            if mx < best_mx:
                best = cand
                best_mx = mx

        # fallback (всё равно добавляем лучший найденный)
        if best is None:
            best, _ = build_desc(brand_lat, shape, lens, collection, seo_level, gender_mode, variant_id=i)
        used.append(best)
        used_sigs.add(_desc_signature(best))

    return used[:need]


# -----------------------------
# MAIN FILL
# -----------------------------
def fill_wb_template(
    input_xlsx: str,
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    style: str = "neutral",
    desc_length: str = "medium",
    seo_level: str = "high",
    gender_mode: str = "Auto",
    wb_safe_mode: bool = True,
    wb_strict: bool = True,
    uniq_strength: int = 90,
    brand_in_title_mode: str = "smart50",
    data_dir: str = "",
    progress_callback=None,
    max_fill_rows: int = 6,
    output_index: int = 1,
    output_total: int = 1,
    between_files_slogan_lock: bool = True,
) -> Tuple[str, int, dict]:

    # Поддержка старых вызовов: fill_wb_template(FillParams(...))
    if not isinstance(input_xlsx, str) and input_xlsx is not None and hasattr(input_xlsx, "__dict__"):
        return fill_wb_template(**dict(input_xlsx.__dict__))

    if not input_xlsx:
        raise RuntimeError("Файл XLSX не выбран")

    _seed_hard()

    wb = load_workbook(input_xlsx, data_only=False, keep_links=False)
    ws = wb.active
    _fix_merged_cells(ws)

    title_candidates = {"наименование","название","название товара","наименование товара","имя товара","title"}
    desc_candidates = {"описание","описание товара","текст","текст описания","description","desc"}

    col_title, hr1 = find_header_col(ws, title_candidates)
    col_desc, hr2 = find_header_col(ws, desc_candidates)

    if not col_title or not col_desc:
        raise RuntimeError("Не найдены колонки Наименование и/или Описание (проверь заголовки)")

    header_row = hr1 or hr2 or 1
    start_row = max(header_row + 1, 5)  # 1–4 не трогаем

    # РОВНО 6 строк (или сколько max_fill_rows), не больше
    max_fill_rows = int(max_fill_rows) if max_fill_rows else 6
    total_rows = max(1, min(max_fill_rows, max(1, ws.max_row - start_row + 1)))

    brand_map = load_brands_ru_map(data_dir) if data_dir else {}
    slogan_pool = SLOGANS[:]
    random.shuffle(slogan_pool)

    # генерим 6 уникальных описаний одним пакетом (чтобы точно не повторялись)
    descs = generate_unique_descs(
        brand_lat=brand_lat,
        shape=shape,
        lens=lens,
        collection=collection,
        seo_level=seo_level,
        gender_mode=gender_mode,
        uniq_strength=uniq_strength,
        need=total_rows,
    )

    used_titles: Set[str] = set()
    processed = 0

    for i in range(total_rows):
        r = start_row + i

        # title unique
        t = None
        for _ in range(260):
            tt = generate_title(
                brand_lat=brand_lat,
                shape=shape,
                lens=lens,
                brand_map=brand_map,
                brand_in_title_mode=brand_in_title_mode,
                slogan_pool=slogan_pool,
                lock_between_files=between_files_slogan_lock,
            )
            # ещё подпись, чтобы не было дубля по смыслу
            sig = _normalize_plain(tt)
            if tt not in used_titles and sig not in _GLOBAL_USED_TITLE_SIGS:
                t = tt
                used_titles.add(tt)
                _GLOBAL_USED_TITLE_SIGS.add(sig)
                break
        if t is None:
            t = generate_title(
                brand_lat, shape, lens, brand_map,
                brand_in_title_mode, slogan_pool, between_files_slogan_lock
            )

        d = descs[i]
        d = FORBIDDEN_LABELS_RE.sub("", d)  # финальная зачистка
        d = _cap_first(d)

        ws.cell(row=r, column=col_title).value = t
        ws.cell(row=r, column=col_desc).value = d

        processed += 1
        if progress_callback:
            progress_callback((processed / total_rows) * 100)

    _fix_merged_cells(ws)

    src = Path(input_xlsx)
    out_dir = read_output_dir(data_dir) if data_dir else None
    base_name = src.stem + "_ready"
    file_name = f"{base_name}_{output_index:02d}.xlsx" if output_total and output_total > 1 else f"{base_name}.xlsx"
    out_path = str((Path(out_dir) / file_name) if out_dir else src.with_name(file_name))

    wb.save(out_path)

    return out_path, processed, {
        "rows_filled": processed,
        "max_fill_rows": total_rows,
        "uniq_strength": uniq_strength,
        "output_dir": out_dir or "",
    }


def generate_preview(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str = "high",
    gender_mode: str = "Auto",
    uniq_strength: int = 90,
    brand_in_title_mode: str = "smart50",
    data_dir: str = "",
    count: int = 3,
) -> list:
    """
    UI обычно ждёт list[(title, desc)].
    """
    _seed_hard()
    brand_map = load_brands_ru_map(data_dir) if data_dir else {}
    slogan_pool = SLOGANS[:]
    random.shuffle(slogan_pool)

    descs = generate_unique_descs(
        brand_lat, shape, lens, collection,
        seo_level=seo_level, gender_mode=gender_mode,
        uniq_strength=uniq_strength, need=max(1, int(count))
    )

    used_titles = set()
    out = []
    for i in range(max(1, int(count))):
        t = generate_title(
            brand_lat, shape, lens, brand_map,
            brand_in_title_mode, slogan_pool, lock_between_files=False
        )
        if t in used_titles:
            t = t + " "
        used_titles.add(t)
        out.append((t, _cap_first(descs[i])))
    return out
