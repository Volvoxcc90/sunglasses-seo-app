# wb_fill.py
import json
import random
import re
from pathlib import Path
from typing import Set, Dict, List, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.cell_range import MultiCellRange

TITLE_MAX = 60
DESC_MAX = 2000

# =========================
# 80 слоганов (для названия)
# =========================
SLOGANS = [
    "Красивые","Крутые","Стильные","Модные","Молодёжные","Дизайнерские","Эффектные","Трендовые","Лаконичные","Яркие",
    "Современные","Премиальные","Универсальные","Актуальные","Выразительные","Элегантные","Минималистичные","Смелые","Классные","Городские",
    "Лёгкие","Комфортные","Популярные","Эксклюзивные","Фирменные","Изящные","Брутальные","Ультрамодные","Шикарные","Статусные",
    "Фотогеничные","Инстаграмные","Сочные","Новые","Практичные","Надёжные","Удобные","Качественные","В тренде сезона","На каждый день",
    "С характером","Стильный акцент","Аккуратные","Сильный силуэт","Под базовый гардероб","Для города","Для отпуска","Для лета","Для поездок","Для фото",
    "Универсал","Хит сезона","Культовые","Легендарные","Смарт-кэжуал","Стрит-стайл","Выгодные","Топовые","Нарядные","Повседневные",
    "С чистой геометрией","С мягкими линиями","С выразительной формой","С классной посадкой","С актуальным дизайном","С модным вайбом","Лёгкий люкс-стиль","С тонким логотипом","С современным силуэтом","С эффектной оправой",
    "Вау-эффект","Стиль-ап","Для уверенного образа","Для яркого лука","Под любой стиль","Под кэжуал","Под классику","Для прогулок","Для пляжа","Для выходных",
    "Супер-стильные","Очень стильные","Прям в тему","Тот самый акцент","Собирают образ","Дают вау-эффект","Поднимают лук","Делают образ дороже","Сочетаются легко","Выглядят дорого"
]

SUN_TERMS = ["солнцезащитные очки", "солнечные очки"]

SCENARIOS = [
    "город","путешествия","отпуск","прогулки","вождение","пляж","активный отдых","поездки","летние мероприятия","выходные",
    "работа","учёба","кафе","шопинг","свидание","фото","на каждый день"
]

# SEO набор (не списком в конце; вшиваем внутрь)
SEO_CORE = ["солнцезащитные очки", "солнечные очки", "очки солнцезащитные"]
SEO_STYLE = ["брендовые очки", "модные очки", "трендовые очки", "стильные очки", "имиджевые очки"]
SEO_USE = ["очки для вождения", "очки для города", "очки для отпуска", "очки для пляжа", "очки для прогулок"]
SEO_SOC = ["инста очки", "очки из tiktok", "очки для фото"]

STRICT_DROP = ["лучшие","самые лучшие","идеальные","100%","гарантия","гарантируем","абсолютно","безусловно","всегда","никогда","полностью"]
SAFE_REPLACE = {"реплика":"стиль в духе бренда", "копия":"вдохновлённый дизайн", "люкс":"премиальный стиль"}

STOPWORDS_RU = {"и","в","во","на","а","но","что","это","как","для","по","из","к","с","со","при","от","до","у","же","не","без","над","под","про","или","то","же","ли"}

# ======= фразы (расширено) =======
OPENERS = [
    "Эти очки — удачное дополнение к любому образу",
    "Очки легко впишутся в повседневный и более нарядный образ",
    "Модель смотрится актуально и делает образ собранным",
    "Современный дизайн подойдёт на каждый день",
    "Очки добавляют стильный акцент и выделяют образ",
    "Универсальная модель: подходит и под базу, и под яркий лук",
    "Очки выглядят аккуратно и современно",
    "Модель хорошо смотрится в городском стиле и в отпуске",
]

BENEFITS = [
    "подчеркнут стиль и индивидуальность",
    "станут ярким акцентом в образе",
    "добавят уверенности и завершат лук",
    "сделают образ более выразительным",
    "легко сочетаются с одеждой",
    "подойдут на каждый день",
    "в кадре смотрятся особенно выигрышно",
    "подойдут под разные стили",
]

FRAME_PHRASES = [
    "оправа обращает на себя внимание и выглядит гармонично",
    "форма оправы хорошо балансирует черты лица",
    "дизайн оправы подчёркивает индивидуальность",
    "оправа смотрится современно и аккуратно",
    "форма оправы подходит под разные типы лица",
    "оправа делает образ более собранным",
]

LENS_PHRASES = [
    "линзы помогают чувствовать себя комфортно при ярком солнце",
    "линзы подходят для улицы, поездок и активного дня",
    "за счёт линз глаза меньше устают на солнце",
    "линзы дают комфорт при ярком свете и на улице, и в городе",
]

GIFT_PHRASES = [
    "отличный вариант на подарок",
    "удачный подарочный вариант для девушки или парня",
    "можно взять себе и в подарок",
    "подойдёт как подарок на праздник или просто без повода",
]

UNISEX_PHRASES = [
    "подойдут и девушкам, и мужчинам",
    "унисекс — хорошо смотрятся и на женском, и на мужском образе",
    "универсальная модель для разных стилей",
    "подойдут как для девушек, так и для мужчин",
]

DISCLAIMERS = [
    "Футляр и комплектация могут отличаться.",
    "Оттенок может немного отличаться из-за настроек экрана.",
    "Принт и детали могут отличаться в зависимости от партии.",
    "Комплектация может отличаться.",
]

# =========================
# Utils
# =========================
def _cut_no_break_words(text: str, limit: int) -> str:
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    return text[:limit].rsplit(" ", 1)[0].strip()

def normalize_key(s: str) -> str:
    s = (s or "").strip().lower()
    s = s.replace("-", " ").replace("&", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def load_brands_ru_map(data_dir: str) -> Dict[str, str]:
    p = Path(data_dir) / "brands_ru.json"
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

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
    uniq_strength = max(40, min(90, uniq_strength))
    return 0.86 - (uniq_strength - 40) * (0.26 / 50.0)

def gender_phrase(mode: str) -> str:
    if mode == "Женские":
        return "женские"
    if mode == "Мужские":
        return "мужские"
    if mode == "Унисекс":
        return "унисекс"
    return ""

def _sentence(s: str) -> str:
    s = re.sub(r"\s{2,}", " ", (s or "").strip())
    if not s:
        return ""
    if s[-1] not in ".!?":
        s += "."
    return s

def _pick_seo_inline(seo_level: str, gender_mode: str) -> Dict[str, str]:
    """
    Возвращает отдельные SEO-вставки, которые мы ВШИВАЕМ в смысл.
    """
    core_a = random.choice(SEO_CORE)
    core_b = random.choice([x for x in SEO_CORE if x != core_a] or SEO_CORE)

    k_style = 1 if seo_level == "low" else (2 if seo_level == "normal" else 3)
    style_pack = random.sample(SEO_STYLE, k=min(k_style, len(SEO_STYLE)))

    k_use = 1 if seo_level == "low" else (2 if seo_level == "normal" else 3)
    use_pack = random.sample(SEO_USE, k=min(k_use, len(SEO_USE)))

    soc_pack = []
    if seo_level != "low":
        soc_pack = random.sample(SEO_SOC, k=1 if seo_level == "normal" else 2)

    if gender_mode == "Auto":
        gender_key = "очки женские и мужские"
    elif gender_mode == "Женские":
        gender_key = "очки женские"
    elif gender_mode == "Мужские":
        gender_key = "очки мужские"
    else:
        gender_key = "очки унисекс"

    return {
        "core_a": core_a,
        "core_b": core_b,
        "style_1": style_pack[0] if len(style_pack) > 0 else "",
        "style_2": style_pack[1] if len(style_pack) > 1 else "",
        "style_3": style_pack[2] if len(style_pack) > 2 else "",
        "use_1": use_pack[0] if len(use_pack) > 0 else "",
        "use_2": use_pack[1] if len(use_pack) > 1 else "",
        "use_3": use_pack[2] if len(use_pack) > 2 else "",
        "soc_1": soc_pack[0] if len(soc_pack) > 0 else "",
        "soc_2": soc_pack[1] if len(soc_pack) > 1 else "",
        "gender": gender_key,
    }

# =========================
# Title generator
# =========================
def generate_title(
    brand_lat: str,
    shape: str,
    lens: str,
    brand_map: Dict[str, str],
    slogan_pool: List[str],
) -> str:
    b_ru = brand_ru(brand_lat, brand_map)

    if not slogan_pool:
        slogan_pool.extend(SLOGANS)
        random.shuffle(slogan_pool)

    slogan = slogan_pool.pop()
    parts = [slogan, random.choice(SUN_TERMS)]

    # бренд 50%
    if random.random() < 0.5 and b_ru:
        parts.append(b_ru)

    # форма/линзы
    if shape and random.random() < 0.65:
        parts.append(shape)
    if lens and random.random() < 0.55:
        parts.append(lens)

    title = " ".join([p for p in parts if p]).strip()
    title = re.sub(r"\s{2,}", " ", title)
    title = title[:1].upper() + title[1:]
    return _cut_no_break_words(title, TITLE_MAX)

# =========================
# Description templates (20+)
# =========================
def _build_desc_variant(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
) -> str:
    seo = _pick_seo_inline(seo_level, gender_mode)
    scen = random.sample(SCENARIOS, 4)

    opener = random.choice(OPENERS)
    benefit = random.choice(BENEFITS)
    unisex = random.choice(UNISEX_PHRASES)
    frame = random.choice(FRAME_PHRASES)
    lens_p = random.choice(LENS_PHRASES)
    gift = random.choice(GIFT_PHRASES)

    # чтобы не было одинаковой структуры — выбираем шаблон
    t = random.randint(1, 24)

    sents: List[str] = []

    if t == 1:
        sents.append(_sentence(f"{seo['core_a'].capitalize()} {brand_lat} — {opener.lower()}, они {benefit}"))
        sents.append(_sentence(unisex))
        if shape:
            sents.append(_sentence(f"Форма оправы {shape}: {frame}"))
        if lens:
            sents.append(_sentence(f"Линзы {lens} — {lens_p}"))
        sents.append(_sentence(f"Подойдут как {seo['use_1']}, а также для {', '.join(scen)}"))
        if collection:
            sents.append(_sentence(f"Коллекция {collection}"))
        if seo["style_1"]:
            sents.append(_sentence(f"Если ищете {seo['style_1']}, эта модель будет удачным выбором"))
        sents.append(_sentence(gift))

    elif t == 2:
        sents.append(_sentence(f"{opener}. {seo['core_a'].capitalize()} {brand_lat} {benefit}"))
        if shape:
            sents.append(_sentence(f"{shape.capitalize()} оправа смотрится актуально и подходит под разные стили"))
        if lens:
            sents.append(_sentence(f"Линзы {lens}: {lens_p}"))
        sents.append(_sentence(f"Хороши для {', '.join(scen[:3])} и когда нужно {seo['use_1']}"))
        sents.append(_sentence(f"{seo['gender']} — модель универсальная и удобная"))
        if seo["soc_1"]:
            sents.append(_sentence(f"Модель отлично смотрится на фото — часто берут как {seo['soc_1']}"))
        sents.append(_sentence(gift))

    elif t == 3:
        sents.append(_sentence(f"{seo['core_a'].capitalize()} {brand_lat} — современная модель на тёплый сезон"))
        sents.append(_sentence(f"Они {benefit} и легко вписываются в гардероб"))
        if shape:
            sents.append(_sentence(f"Форма оправы {shape} — {frame}"))
        if lens:
            sents.append(_sentence(f"Линзы {lens} — {lens_p}"))
        if collection:
            sents.append(_sentence(f"Актуальная коллекция: {collection}"))
        sents.append(_sentence(f"Подходят для {seo['use_1']} и для {', '.join(scen)}"))
        if seo["style_1"]:
            sents.append(_sentence(f"{seo['style_1'].capitalize()} — отличный вариант для тех, кто любит заметные аксессуары"))
        sents.append(_sentence(random.choice(DISCLAIMERS)))

    elif t == 4:
        sents.append(_sentence(f"Модель {brand_lat} — {seo['core_a']}, которые {benefit}"))
        if shape:
            sents.append(_sentence(f"Оправу {shape} выбирают за то, что она выглядит аккуратно и современно"))
        sents.append(_sentence(unisex))
        if lens:
            sents.append(_sentence(f"Линзы {lens} подходят для {seo['use_1']} и для активного дня"))
        sents.append(_sentence(f"Подойдут для {', '.join(scen[:4])}"))
        if seo["style_1"] and seo["style_2"]:
            sents.append(_sentence(f"Ищете {seo['style_1']} или {seo['style_2']} — присмотритесь к этой модели"))
        sents.append(_sentence(gift))

    elif t == 5:
        sents.append(_sentence(f"{seo['core_a'].capitalize()} {brand_lat} — аккуратный аксессуар, который {benefit}"))
        if shape:
            sents.append(_sentence(f"Форма {shape} подходит под разные типы лица и под разные образы"))
        if lens:
            sents.append(_sentence(f"Линзы {lens}: {lens_p}"))
        sents.append(_sentence(f"Удобны для {seo['use_1']}, прогулок и поездок"))
        if collection:
            sents.append(_sentence(collection))
        if seo["soc_1"]:
            sents.append(_sentence(f"Для фото и сторис — отличный вариант, многие ищут именно {seo['soc_1']}"))
        sents.append(_sentence(random.choice(DISCLAIMERS)))

    else:
        # остальные варианты: миксуем блоки и порядок
        blocks = []

        blocks.append(_sentence(f"{seo['core_a'].capitalize()} {brand_lat} {benefit} — {opener.lower()}"))
        if random.random() < 0.8:
            blocks.append(_sentence(unisex))
        if shape and random.random() < 0.9:
            blocks.append(_sentence(f"Форма оправы {shape} — {frame}"))
        if lens and random.random() < 0.9:
            blocks.append(_sentence(f"Линзы {lens} — {lens_p}"))
        blocks.append(_sentence(f"Подойдут для {seo['use_1']} и для {', '.join(scen)}"))
        if collection and random.random() < 0.7:
            blocks.append(_sentence(f"Коллекция: {collection}"))
        if seo["style_1"] and random.random() < 0.8:
            blocks.append(_sentence(f"{seo['style_1'].capitalize()} — хороший выбор на тёплый сезон"))
        if seo["style_2"] and random.random() < 0.55:
            blocks.append(_sentence(f"Также это {seo['style_2']} — модель легко сочетается с одеждой"))
        if seo["soc_1"] and random.random() < 0.5:
            blocks.append(_sentence(f"Часто берут как {seo['soc_1']} — смотрится выигрышно в кадре"))
        if random.random() < 0.7:
            blocks.append(_sentence(gift))
        if random.random() < 0.35:
            blocks.append(_sentence(random.choice(DISCLAIMERS)))

        random.shuffle(blocks)
        sents = blocks[:random.randint(7, 10)]  # длина 7–10 предложений

    text = " ".join([s for s in sents if s]).strip()
    text = re.sub(r"\s{2,}", " ", text)
    return _cut_no_break_words(text, DESC_MAX)

def generate_description_best_of(
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    seo_level: str,
    gender_mode: str,
    used_desc: List[str],
    uniq_strength: int,
    tries: int = 30,
) -> Tuple[str, float]:
    """
    Главное анти-дубли:
    генерим много кандидатов и выбираем самый "далёкий" от последних описаний.
    """
    thr = uniqueness_threshold(uniq_strength)
    recent = used_desc[-25:]  # сравниваем с последними
    best_text = ""
    best_score = 1.0  # чем меньше, тем менее похоже
    for _ in range(max(10, tries)):
        cand = _build_desc_variant(brand_lat, shape, lens, collection, seo_level, gender_mode)
        if not recent:
            return cand, 0.0
        mx = max(jaccard(cand, prev) for prev in recent)
        # если ниже порога — сразу берём
        if mx <= thr:
            return cand, mx
        # иначе оставляем самый непохожий (минимум mx)
        if mx < best_score:
            best_score = mx
            best_text = cand
    return (best_text or _build_desc_variant(brand_lat, shape, lens, collection, seo_level, gender_mode)), best_score

# =========================
# Excel helpers
# =========================
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

def find_header_col(ws, candidates: set, header_scan_rows: int = 25):
    for r in range(1, header_scan_rows + 1):
        for cell in ws[r]:
            if cell.value is None:
                continue
            val = str(cell.value).strip().lower()
            if val in candidates:
                return cell.column, r
    return None, None

# =========================
# Fill XLSX
# =========================
def fill_wb_template(
    input_xlsx: str,
    brand_lat: str,
    shape: str,
    lens: str,
    collection: str,
    style: str = "neutral",      # оставили для совместимости
    desc_length: str = "medium", # оставили для совместимости
    seo_level: str = "high",
    gender_mode: str = "Auto",
    wb_safe_mode: bool = True,
    wb_strict: bool = True,
    uniq_strength: int = 75,
    data_dir: str = "",
    progress_callback=None,
) -> Tuple[str, int, dict]:
    if not input_xlsx:
        raise RuntimeError("Файл XLSX не выбран")

    wb = load_workbook(input_xlsx, data_only=False, keep_links=False)
    ws = wb.active
    _fix_merged_cells(ws)

    col_title, header_row = find_header_col(ws, {"наименование", "название"})
    col_desc, header_row2 = find_header_col(ws, {"описание", "description"})
    if not col_title or not col_desc:
        raise RuntimeError("Не найдены колонки Наименование и/или Описание")

    header_row = header_row or header_row2 or 1
    start_row = max(header_row + 1, 5)  # не трогаем 1–4 строки

    total_rows = ws.max_row - start_row + 1
    if total_rows <= 0:
        raise RuntimeError("Нет строк для заполнения (после заголовка)")

    brand_map = load_brands_ru_map(data_dir) if data_dir else {}
    slogan_pool = SLOGANS[:]
    random.shuffle(slogan_pool)

    used_titles = set()
    used_desc: List[str] = []

    report = {
        "picked_best_of": 0,
        "avg_max_jaccard": 0.0,
        "uniq_strength": uniq_strength
    }

    processed = 0
    sum_mx = 0.0

    for r in range(start_row, ws.max_row + 1):
        # ---- Title unique
        t = None
        for _k in range(120):
            tt = generate_title(brand_lat, shape, lens, brand_map, slogan_pool)
            if tt not in used_titles:
                t = tt
                used_titles.add(tt)
                break
        if t is None:
            t = generate_title(brand_lat, shape, lens, brand_map, slogan_pool)

        # ---- Description (best-of-30 анти-дубль)
        d, mx = generate_description_best_of(
            brand_lat=brand_lat,
            shape=shape,
            lens=lens,
            collection=collection,
            seo_level=seo_level,
            gender_mode=gender_mode,
            used_desc=used_desc,
            uniq_strength=uniq_strength,
            tries=32 if seo_level == "high" else 24,
        )
        report["picked_best_of"] += 1
        sum_mx += float(mx)
        used_desc.append(d)

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

    report["avg_max_jaccard"] = round(sum_mx / max(1, processed), 3)

    _fix_merged_cells(ws)

    out_path = str(Path(input_xlsx).with_name(Path(input_xlsx).stem + "_ready.xlsx"))
    wb.save(out_path)
    return out_path, processed, report
