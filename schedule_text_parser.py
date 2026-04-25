# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import List

from schedule_core import ShiftRecord, export_to_excel, timestamped_output_path


SHIFT_HOURS = {"早": 4, "午": 5, "晚": 4}
WEEKDAY_MAP = "一二三四五六日"
NAME_VARIANTS = {
    "张盈慧": "張盈慧",
    "張盈慧": "張盈慧",
}


def _normalize_name(text: str) -> str:
    normalized = str(text).strip().replace(" ", "")
    return NAME_VARIANTS.get(normalized, normalized)


def _split_rows(text: str) -> List[List[str]]:
    rows: List[List[str]] = []
    width = 0
    for raw_line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if not raw_line.strip():
            continue
        cells = [cell.strip() for cell in raw_line.split("\t")]
        while cells and cells[-1] == "":
            cells.pop()
        width = max(width, len(cells))
        rows.append(cells)

    normalized_rows: List[List[str]] = []
    for row in rows:
        normalized_rows.append(row + [""] * (width - len(row)))
    return normalized_rows


def _extract_year_month(rows: List[List[str]]) -> tuple[int, int]:
    for row in rows:
        for cell in row:
            match = re.search(r"(\d{4})\s*/\s*(\d{1,2})", cell)
            if match:
                return int(match.group(1)), int(match.group(2))
    raise ValueError("無法從貼上的文字辨識出年份與月份。")


def parse_pasted_schedule_text(text: str, employee_name: str) -> tuple[list[ShiftRecord], dict]:
    rows = _split_rows(text)
    if not rows:
        raise ValueError("沒有貼上任何班表文字內容。")

    year, month = _extract_year_month(rows)
    target_name = _normalize_name(employee_name)
    records: list[ShiftRecord] = []

    date_row_indexes = [
        idx
        for idx, row in enumerate(rows)
        if any(re.fullmatch(r"\d{2}/\d{2}", cell) for cell in row if cell)
    ]

    for date_row_idx in date_row_indexes:
        if date_row_idx + 3 >= len(rows):
            continue

        early_row = rows[date_row_idx + 1]
        noon_row = rows[date_row_idx + 2]
        late_row = rows[date_row_idx + 3]

        for col_idx, cell in enumerate(rows[date_row_idx]):
            if not re.fullmatch(r"\d{2}/\d{2}", cell or ""):
                continue
            mm, dd = map(int, cell.split("/"))
            if mm != month:
                continue
            work_date = date(year, month, dd)

            for shift, row, hours in (
                ("早", early_row, SHIFT_HOURS["早"]),
                ("午", noon_row, SHIFT_HOURS["午"]),
                ("晚", late_row, SHIFT_HOURS["晚"]),
            ):
                if col_idx >= len(row):
                    continue
                if _normalize_name(row[col_idx]) != target_name:
                    continue
                records.append(
                    ShiftRecord(
                        year=work_date.year,
                        month=work_date.month,
                        day=work_date.day,
                        weekday=WEEKDAY_MAP[work_date.weekday()],
                        shift=shift,
                        hours=hours,
                    )
                )

    records.sort(key=lambda item: (item.year, item.month, item.day, ("早", "午", "晚").index(item.shift)))
    debug = {
        "source": "pasted-text",
        "rows": len(rows),
        "matched_records": len(records),
        "table_rows": rows,
    }
    return records, debug


def run_text_conversion_debug(
    text: str,
    employee_name: str,
    output_path: Path | None = None,
) -> tuple[Path, list[ShiftRecord], dict]:
    records, debug = parse_pasted_schedule_text(text, employee_name)
    pseudo_source = Path("pasted_schedule.txt")
    final_output = output_path or timestamped_output_path(pseudo_source, employee_name, records)
    saved_path = export_to_excel(records=records, employee_name=employee_name, output_path=final_output)
    return saved_path, records, debug
