from __future__ import annotations

import re
import random
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell


# ======================
# helpers
# ======================

def _clean(s: Optional[str]) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def _find_header_row_and_cols(
    ws,
    needed: List[str],
    max_rows: int = 20,
    max_cols: int = 400
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


def _try_find_col(ws, header: str, header_row: int, max_cols: int = 400) -> Optional[int]:
    for c in range(1, max_cols + 1):
        v = ws.cell(header_row, c).value
        if isinstance(v, str) and v.strip() == header:
            return c
    return None


def _row_is_product(ws, r: int, signal_cols: List[int]) -> bool:
    for c in signal_cols:
        v = ws.cell(r, c).value
        if v not in (None, ""):
            return True
    return False


# ======================
# SEO generation
# ======================

STARTS = ["Солнцезащитные очки", "Солнечные очки"]
MAX_TITLE = 60
VARIANTS = 12


def _gen_titles(brand: str, shape: str, lens: str) -> List[str]:
    brand = _clean(brand)
    shape = _clean(shape)
    lens = _clean(lens)

    patterns = [
        "{s} {b} {l}",
        "{s} {b} {l} {sh}",
        "{s} {b} {sh} {l}",
        "{s} {l} {b}",
        "{s} {l} {b} {sh}",
        "{s} {b} {sh}",
        "{s} {b} {l} модель 2025",
        "{s} {b} {l} 2026",
        "{s} {b} {sh} {l} 2025",
        "{s} {l} {b} модель",
        "{s} {b} дизайнерские {l}",
        "{s} {b} {l} солнцезащитные",
    ]

    titles: List[str] = []
    for i in range(VARIANTS):
        s = STARTS[i % 2]
        t = patterns[i].format(s=s, b=brand, l=lens, sh=shape)
        t = _clean(t)
        if len(t) > MAX_TITLE:
            t = t[:MAX_TITLE].rsplit(" ", 1)[0]
        titles.append(t)

    return titles


def _gen_desc(brand: str, shape: str, lens: str, collection: str, seed: int) -> str:
    rng = random.Random(seed)

    blocks = [
        f"{rng.choice(STARTS)} {brand} — стильное решение из коллекции {collection}.",
        f"Форма оправы: {shape}. Линзы: {lens}.",
        f"Подходит для города, отдыха, поездок и яркой солнечной погоды.",
        f"Комфортная посадка и защита зрения при повседневном использовании.",
        f"Современный аксессуар, который подчёркивает образ.",
        f"SEO: солнцезащитные очки {brand}, очки {shape}, очки {lens}."
    ]

    rng.shuffle(blocks)
    return "\n\n".join(blocks[:4])


# ======================
# SAFE WRITE (merged cells)
# ======================

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
# MAIN FUNCTION
# ======================

def fill_wb_template(
    input_xlsx: str,
    brand: str,
    shape: str,
    lens_features: str,
    collection: str
) -> Tuple[str, int]:

    wb = load_workbook(input_xlsx)
    ws = wb.active

    header_row, cols = _find_header_row_and_cols(ws, ["Наименование", "Описание"])
    col_name = cols["Наименование"]
    col_desc = cols["Описание"]

    # сигнальные колонки (чтобы не писать в пустые строки)
    signal_cols = [col_name, col_desc]
    for h in ["Фото", "Артикул продавца", "Баркоды"]:
        c = _try_find_col(ws, h, header_row)
        if c:
            signal_cols.append(c)

    titles = _gen_titles(brand, shape, lens_features)

    filled = 0
    variant = 0

    for r in range(header_row + 1, ws.max_row + 1):
        if not _row_is_product(ws, r, signal_cols):
            continue

        title = titles[variant % VARIANTS]
        desc = _gen_desc(brand, shape, lens_features, collection, seed=1000 + r)

        _write_safe(ws, r, col_name, title)
        _write_safe(ws, r, col_desc, desc)

        filled += 1
        variant += 1

    in_path = Path(input_xlsx)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = in_path.parent / f"{in_path.stem}_filled_{ts}.xlsx"

    wb.save(out_path)
    return str(out_path), filled
