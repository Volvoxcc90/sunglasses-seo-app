# wb_fill.py
import json
import random
import re
import os
from pathlib import Path
from typing import Dict, List, Tuple, Set, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.cell_range import MultiCellRange

TITLE_MAX = 60
DESC_MAX = 2000

SLOGANS = [
    "Красивые","Крутые","Стильные","Модные","Молодёжные","Дизайнерские","Эффектные","Трендовые","Лаконичные","Яркие",
    "Современные","Премиальные","Универсальные","Актуальные","Выразительные","Элегантные","Минималистичные","Смелые","Классные","Городские",
    "Лёгкие","Комфортные","Популярные","Эксклюзивные","Фирменные","Изящные","Брутальные","Ультрамодные","Шикарные","Статусные",
    "Инстаграмные","Фотогеничные","Сочные","Новые","Практичные","Надёжные","Удобные","Качественные","В тренде сезона","На каждый день",
    "С характером","Стильный акцент","Сильный силуэт","Под базовый гардероб","Для города","Для отпуска","Для лета","Для поездок","Для фото",
    "Хит сезона","Топовые","Повседневные","Нарядные","С эффектной оправой","С модным вайбом","Лёгкий люкс-стиль","С современным силуэтом",
    "Вау-эффект","Прям в тему","Тот самый акцент","Собирают образ","Делают образ дороже","Сочетаются легко","Выглядят дорого",
]

SUN_TERMS = ["солнцезащитные очки", "солнечные очки"]

SCENARIOS = [
    "город","путешествия","отпуск","прогулки","вождение","пляж","активный отдых","поездки","летние мероприятия",
    "кафе","шопинг","свидание","на каждый день","для фото"
]

SEO_CORE = ["солнцезащитные очки", "солнечные очки", "очки солнцезащитные"]
SEO_STYLE = ["брендовые очки", "модные очки", "трендовые очки", "стильные очки", "имиджевые очки"]
SEO_USE = ["очки для вождения", "очки для города", "очки для отпуска", "очки для пляжа", "очки для прогулок"]
SEO_SOC = ["инста очки", "очки из tiktok", "очки для фото"]

STRICT_DROP = ["лучшие","самые лучшие","идеальные","100%","гарантия","гарантируем","абсолютно","безусловно","всегда","никогда","полностью"]
SAFE_REPLACE = {"реплика":"стиль в духе бренда", "копия":"вдохновлённый дизайн", "люкс":"премиальный стиль"}

STOPWORDS_RU = {"и","в","во","на","а","но","что","это","как","для","по","из","к","с","со","при","от","до","у","же","не","без","над","под","про","или","то","ли"}

BLOCK_OPEN = [
    "Очки — классное дополнение к любому образу",
    "Эти очки реально выручают в солнечную погоду",
    "Если нужен стильный аксессуар на каждый день — вот он",
    "Лёгкая модель: и в городе норм, и в отпуск — самое то",
    "Смотрятся аккуратно, но при этом заметно",
    "Подойдут под базовый гардероб и под яркий лук",
    "Модель выглядит трендово, но без перебора",
    "Это тот самый аксессуар, который собирает образ в одну линию",
    "С такими очками образ сразу выглядит дороже и аккуратнее",
]

BLOCK_FRAME = [
    "Оправа смотрится ровно и хорошо садится",
    "Форма оправы удачная — лицо смотрится более собранно",
    "Дизайн оправы делает образ дороже",
    "Оправу легко сочетать с одеждой — и кэжуал, и более нарядно",
    "Форма подчёркивает стиль и не «шумит» в образе",
    "Сидят комфортно, не давят и не раздражают в носке",
]

BLOCK_LENS = [
    "Линзы дают комфорт при ярком солнце",
    "Глаза меньше устают на улице и в дороге",
    "На солнце реально удобнее: меньше бликов и лишнего света",
    "Под яркий день — то, что нужно",
    "В солнечную погоду видно спокойнее и приятнее",
]

BLOCK_SCEN = [
    "Хорошо заходят для города, поездок и прогулок",
    "Удобно брать в отпуск, на пляж и на выходные",
    "Подойдут для вождения и повседневных дел",
    "Норм вариант для фото и сторис",
    "Если много двигаешься — удобный формат на каждый день",
]

BLOCK_GIFT = [
    "Можно брать себе или на подарок",
    "Отличный подарочный вариант для девушки или парня",
    "Если ищешь подарок — вариант рабочий",
]

BLOCK_MISC = [
    "Футляр/комплектация могут отличаться.",
    "Оттенок может немного отличаться из-за настроек экрана.",
    "Детали могут отличаться в зависимости от партии.",
]


def _cut_no_break_words(text: str, limit: int) -> str:
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    cut = text[:limit]
    if " " not in cut:
        return cut
    return cut.rsplit(" ", 1)[0].strip()

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

def apply_safe(text: str) -> str:
    t = text
    for a, b in SAFE_REPLACE.items():
        t = re.sub(rf"\b{re.escape(a)}\b", b, t, flags=re.IGNORECASE)
    return t

def apply_strict(text: str) -> str:
    t = text
    for w in STRICT_DROP:
        t = re.sub(rf"\b{re.escape(w)}\b", "", t, flags=re.IGNORECASE)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t

def _tokens(text: str) -> Set[str]:
    t = (text or "").lower()
    t = re.sub(r"[^a-zа-яё0-9\s\-]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return {w for w in t.split() if len(w) >= 3 and w not in STOPWORDS_RU}

def jaccard(a: str, b: str) -> float:
    A = _tokens(a); B = _tokens(b)
    if not A or not B:
        return 0.0
    return len(A & B) / max(1, len(A | B))

def uniqueness_threshold(uniq_strength: int) -> float:
    uniq_strength = max(60, min(95, uniq_strength))
    return 0.80 - (uniq_strength - 60) * (0.25 / 35.0)

def _sentence(s: str) -> str:
    s = re.sub(r"\s{2,}", " ", (s or "").strip())
    if not s:
        return ""
    if s[-1] not in ".!?":
        s += "."
    return s

def _pick_seo(seo_level: str, gender_mode: str) -> Dict[str, List[str]]:
    if seo_level == "low":
        k_core, k_style, k_use, k_soc = 1, 1, 1, 0
    elif seo_level == "normal":
        k_core, k_style, k_use, k_soc = 2, 2, 2, 1
    else:
        k_core, k_style, k_use, k_soc = 3, 3, 3, 2

    core = random.sample(SEO_CORE, k=min(k_core, len(SEO_CORE)))
    style = random.sample(SEO_STYLE, k=min(k_style, len(SEO_STYLE)))
    use = random.sample(SEO_USE, k=min(k_use, len(SEO_USE)))
    soc = random.sample(SEO_SOC, k=min(k_soc, len(SEO_SOC))) if k_soc > 0 else []

    if gender_mode == "Auto":
        gender = ["очки женские", "очки мужские"]
    elif gender_mode == "Женские":
        gender = ["очки женские"]
    elif gender_mode == "Мужские":
        gender = ["очки мужские"]
    elif gender_mode == "Унисекс":
        gender = ["очки унисекс"]
    else:
        gender = ["очки унисекс"]

    return {"core": core, "style": style, "use": use, "soc": soc, "gender": gender}

def generate_title(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    brand_map: Dict[str, str],
    slogan_pool: List[str],
    brand_in_title_mode: str = "smart50",
) -> str:
    b_ru = brand_ru(brand_lat, brand_map)

    if not slogan_pool:
        slogan_pool.extend(SLOGANS)
        random.shuffle(slogan_pool)

    slogan = slogan_pool.pop()
    sun = random.choice(SUN_TERMS)

    parts = [slogan, sun]

    if brand_in_title_mode == "always":
        put_brand = True
    elif brand_in_title_mode == "never":
        put_brand = False
    else:
        put_brand = (random.random() < 0.5)

    if put_brand and b_ru:
        parts.append(b_ru)

    if shape and random.random() < 0.70:
        parts.append(shape)
    if lens and random.random() < 0.55:
        parts.append(lens)

    title = " ".join([p for p in parts if p]).strip()
    title = re.sub(r"\s{2,}", " ", title)
    title = title[:1].upper() + title[1:]
    return _cut_no_break_words(title, TITLE_MAX)

def build_description_variant(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
) -> str:
    seo = _pick_seo(seo_level, gender_mode)
    scen = random.sample(SCENARIOS, 4)

    blocks: List[str] = []

    head_core = random.choice(seo["core"]) if seo["core"] else "солнцезащитные очки"
    blocks.append(_sentence(f"{head_core.capitalize()} {brand_lat} — {random.choice(BLOCK_OPEN).lower()}"))

    if shape:
        blocks.append(_sentence(f"Форма {shape} — {random.choice(BLOCK_FRAME).lower()}"))
    else:
        blocks.append(_sentence(random.choice(BLOCK_FRAME)))

    if lens:
        blocks.append(_sentence(f"Линзы {lens}: {random.choice(BLOCK_LENS).lower()}"))
    else:
        blocks.append(_sentence(random.choice(BLOCK_LENS)))

    if collection:
        blocks.append(_sentence(f"Под сезон {collection} — смотрятся актуально"))

    use_phrase = random.choice(seo["use"]) if seo["use"] else "очки для города"
    blocks.append(_sentence(f"{random.choice(BLOCK_SCEN)} — особенно если нужны {use_phrase}"))
    blocks.append(_sentence(f"Подойдут для: {', '.join(scen)}"))

    # ключи аккуратно, чтобы не выглядело “списком”
    tail_bits = []
    if seo["style"]:
        tail_bits.append(random.choice(seo["style"]))
    if seo["gender"]:
        tail_bits.append(random.choice(seo["gender"]))
    if seo["soc"] and random.random() < 0.6:
        tail_bits.append(random.choice(seo["soc"]))
    if seo["core"] and random.random() < 0.7:
        tail_bits.append(random.choice(seo["core"]))

    if tail_bits:
        blocks.append(_sentence("Часто ищут: " + ", ".join(tail_bits)))

    blocks.append(_sentence(random.choice(BLOCK_GIFT)))
    if random.random() < 0.35:
        blocks.append(_sentence(random.choice(BLOCK_MISC)))

    # перемешиваем середину чтобы не было шаблона
    first = blocks[0]
    middle = blocks[1:]
    random.shuffle(middle)

    text = " ".join([first] + middle).strip()
    text = re.sub(r"\s{2,}", " ", text)

    # запрещаем маркеры/слова вида "Сценарии:" и т.п.
    text = re.sub(r"\b(Сценарии|Ключевые слова|Форма|Линза|Коллекция)\s*:\s*", "", text, flags=re.IGNORECASE)
    return _cut_no_break_words(text, DESC_MAX)

def generate_best_description(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
    used_desc: List[str],
    uniq_strength: int,
    no_repeat_last_n: int,
    tries: int = 40,
) -> Tuple[str, float]:
    thr = uniqueness_threshold(uniq_strength)
    recent = used_desc[-max(0, no_repeat_last_n):] if no_repeat_last_n > 0 else used_desc[-30:]

    best_text = ""
    best_score = 1.0

    for _ in range(max(16, tries)):
        cand = build_description_variant(brand_lat, shape, lens, collection, seo_level, gender_mode)
        if not recent:
            return cand, 0.0
        mx = max(jaccard(cand, prev) for prev in recent)
        if mx <= thr:
            return cand, mx
        if mx < best_score:
            best_score = mx
            best_text = cand

    return (best_text or build_description_variant(brand_lat, shape, lens, collection, seo_level, gender_mode)), best_score

def _fix_merged_cells(ws):
    try:
        if isinstance(ws.merged_cells, MultiCellRange):
            return
        old = ws.merged_cells
        fixed = MultiCellRange()
        try:
            for r in list(old):
                fixed.add(str(r))
        except Exception:
            pass
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
    """
    UI НЕ меняем: путь к папке вывода читаем из data/output_dir.txt (1 строка).
    Если файла нет или путь пустой — вернём None → сохраняем рядом с исходным.
    """
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
    uniq_strength: int = 80,
    no_repeat_last_n: int = 40,
    brand_in_title_mode: str = "smart50",
    data_dir: str = "",
    progress_callback=None,
    max_fill_rows: int = 6,
    output_index: int = 1,
    output_total: int = 1,
) -> Tuple[str, int, dict]:
    if not input_xlsx:
        raise RuntimeError("Файл XLSX не выбран")

    # каждый файл → новый seed (чтобы пачка точно была разной)
    random.seed(int.from_bytes(os.urandom(8), "big"))

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
    start_row = max(header_row + 1, 5)  # не трогаем 1–4 строки

    total_available = ws.max_row - start_row + 1
    if total_available <= 0:
        raise RuntimeError("Нет строк для заполнения (после заголовка)")

    total_rows = min(max(1, int(max_fill_rows)), total_available)

    brand_map = load_brands_ru_map(data_dir) if data_dir else {}
    slogan_pool = SLOGANS[:]
    random.shuffle(slogan_pool)

    used_titles: Set[str] = set()
    used_desc: List[str] = []
    sum_mx = 0.0
    processed = 0

    for i in range(total_rows):
        r = start_row + i

        # уникальные названия в пределах файла
        t = None
        for _k in range(120):
            tt = generate_title(brand_lat, shape, lens, collection, brand_map, slogan_pool, brand_in_title_mode)
            if tt not in used_titles:
                t = tt
                used_titles.add(tt)
                break
        if t is None:
            t = generate_title(brand_lat, shape, lens, collection, brand_map, slogan_pool, brand_in_title_mode)

        d, mx = generate_best_description(
            brand_lat=brand_lat,
            shape=shape,
            lens=lens,
            collection=collection,
            seo_level=seo_level,
            gender_mode=gender_mode,
            used_desc=used_desc,
            uniq_strength=uniq_strength,
            no_repeat_last_n=no_repeat_last_n,
            tries=46 if seo_level == "high" else 34,
        )
        used_desc.append(d)
        sum_mx += float(mx)

        if wb_safe_mode:
            t = apply_safe(t)
            d = apply_safe(d)
        if wb_strict:
            t = apply_strict(t)
            d = apply_strict(d)

        ws.cell(row=r, column=col_title).value = t
        ws.cell(row=r, column=col_desc).value = d

        processed += 1
        if progress_callback:
            progress_callback((processed / total_rows) * 100)

    _fix_merged_cells(ws)

    src = Path(input_xlsx)
    out_dir = read_output_dir(data_dir) if data_dir else None

    base_name = src.stem + "_ready"
    if output_total and output_total > 1:
        file_name = f"{base_name}_{output_index:02d}.xlsx"
    else:
        file_name = f"{base_name}.xlsx"

    if out_dir:
        out_path = str(Path(out_dir) / file_name)
    else:
        out_path = str(src.with_name(file_name))

    wb.save(out_path)

    report = {
        "rows_filled": processed,
        "max_fill_rows": max_fill_rows,
        "avg_max_jaccard": round(sum_mx / max(1, processed), 3),
        "uniq_strength": uniq_strength,
        "no_repeat_last_n": no_repeat_last_n,
        "brand_in_title_mode": brand_in_title_mode,
        "file_index": output_index,
        "file_total": output_total,
        "output_dir": out_dir or "",
    }
    return out_path, processed, report
