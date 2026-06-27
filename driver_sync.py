#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ドライバー情報 自動反映スクリプト（Graph API版）

個人用OneDriveの共有Excelを読み取り専用で取得し、
会社OneDrive上の営業用Excel（閲覧リンク）へ反映する。
必要に応じて SharePoint リストへも洗い替え可能（設定でON/OFF）。
元の依頼書Excelには一切書き込みません。
"""

from __future__ import annotations

import argparse
import atexit
import base64
import json
import logging
import math
import unicodedata
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import date, datetime, time as dt_time, timedelta, timezone

JST = timezone(timedelta(hours=9))
from io import BytesIO
from pathlib import Path
from typing import Any

import jpholiday
import msal
import openpyxl
import requests
from openpyxl.styles import PatternFill

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "driver_sync_config.json"
TOKEN_CACHE_PATH = SCRIPT_DIR / "token_cache.bin"
SHEET_PASSWORD_PATH = SCRIPT_DIR / "sheet_password.txt"
SHEET_PASSWORD_ENV_VAR = "DRIVER_SYNC_SHEET_PASSWORD"
PROTECTED_HIDDEN_COLUMN_NAMES = ("携帯番号",)  # 非表示+シート保護で隠す列
LOG_PATH = SCRIPT_DIR / "driver_sync.log"
TASK_LOG_PATH = SCRIPT_DIR / "driver_sync_task.log"
STATUS_PATH = SCRIPT_DIR / "driver_sync_status.json"
LOCK_PATH = SCRIPT_DIR / "driver_sync.lock"
VIEW_LINK_PATH = SCRIPT_DIR / "営業用閲覧リンク.txt"
LINK_CACHE_PATH = SCRIPT_DIR / "onedrive_link_cache.json"

# 見出し行から列を自動判定（届け先名を案件名に、納入先住所は使わない）
HEADER_FIELD_ALIASES: dict[str, dict[str, Any]] = {
    "anNo": {"labels": ["案件no", "案件no.", "案件番号"], "exclude": []},
    "anName": {
        # 「納入先名」は「納入先住所」と誤マッチするため使わない
        "labels": ["届け先名", "案件名", "現場名"],
        "exclude": ["納入先住所", "納入先", "住所"],
    },
    "ship": {"labels": ["出荷日", "出庫日"], "exclude": []},
    "arr": {"labels": ["着日", "納入先着日"], "exclude": []},
    "model": {"labels": ["型式", "備考"], "exclude": []},
    # ★ドライバー情報★ 列の見出し（列挿入に追従させるため完全一致で検出）
    "company": {"labels": ["会社名", "運送会社", "協力会社", "傭車先"], "exclude": []},
    "driver": {"labels": ["乗務員", "運転手", "ドライバー", "担当ドライバー"], "exclude": []},
    "plate": {"labels": ["車番", "車輌番号", "車両番号", "ナンバー"], "exclude": []},
    "phone": {
        "labels": ["携帯番号", "携帯", "電話番号", "連絡先", "tel"],
        "exclude": ["会社"],
    },
}

ADDRESS_PATTERN = re.compile(
    r"(都|道|府|県).{1,20}(市|区|町|村|郡)|\d{3}-?\d{4}"
)

EXCEL_OUTPUT_COLUMNS = [
    "案件No",
    "案件名",
    "出荷日",
    "着日",
    "備考",
    "型式",
    "車型",
    "会社名",
    "乗務員",
    "車番",
    "携帯番号",
    "依頼先",
]
# ---------------------------------------------------------------------------
# 出力Excel 行色（A〜E=案件情報 / F〜K=★ドライバー情報★）
# ---------------------------------------------------------------------------
CASE_BLOCK_COLUMN_NAMES = ("案件No", "案件名", "出荷日", "着日", "備考", "型式")
DRIVER_BLOCK_COLUMN_NAMES = ("車型", "会社名", "乗務員", "車番", "携帯番号")
# 依頼先(L列)は色分け対象外。常に塗りつぶしなし（白）にする。
WHITE_FILL_COLUMN_NAMES = ("依頼先",)

HIGHLIGHT_CASE_NAMES = (
    "㈱丸運　羽田京浜物流センター",
    "司企業　鳥栖営業所　横持",
    "司企業　鳥栖営業所",
    "株式会社　中通(本社)",
    "福岡ロジテック　宇美倉庫",
)

# 強調3種: ほんの少し濃い青（A〜E） / アクセント2・明るめ60%（F〜K）
HIGHLIGHT_CASE_FILL_HEX = "#C8E3F3"
HIGHLIGHT_DRIVER_FILL_HEX = "#E6B9B8"
# その他: 薄い青（A〜E） / 薄い赤（F〜K）
DEFAULT_CASE_FILL_HEX = "#DAEEF3"
DEFAULT_DRIVER_FILL_HEX = "#F8D6D6"
CHANGED_DRIVER_FILL_HEX = "#FFFF00"

DEFAULT_CASE_NAME_ROW_COLORS: dict[str, str] = {
    name: HIGHLIGHT_CASE_FILL_HEX for name in HIGHLIGHT_CASE_NAMES
}

# 余剰行（書込範囲外）: 値は空、①強調3種と②その他の色を交互に適用

DEFAULT_PADDING_END_ROW = 80
DEFAULT_ROW_FILL_HEX = DEFAULT_CASE_FILL_HEX
DRIVER_BLOCK_FILL_HEX = DEFAULT_DRIVER_FILL_HEX
DRIVER_BLOCK_FILL_HIGHLIGHT_HEX = HIGHLIGHT_DRIVER_FILL_HEX

# 出荷確認ツールと同じ Graph 権限（アプリ登録済み）
GRAPH_SCOPES = [
    "https://graph.microsoft.com/Files.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

SOURCE_ORDER = ["matsuzaki", "nakadori", "fukuoka", "maruun"]
SOURCE_DISPLAY_NAMES = {
    "matsuzaki": "松崎運輸",
    "nakadori": "中通",
    "fukuoka": "福岡ロジテック",
    "maruun": "丸運",
}
DEBUG_OUTPUT_PATH = SCRIPT_DIR / "debug_extract_output.json"
API_DEBUG_PATH = SCRIPT_DIR / "debug_api_last.json"

# 案件名の列（中西さん指定・固定）
CASE_NAME_COLUMN: dict[tuple[str, str], str] = {
    ("matsuzaki", "車両依頼書"): "H",
    ("fukuoka", "車両依頼書"): "G",
    ("nakadori", "車両依頼書"): "G",
    ("nakadori", "滋賀管理シート"): "K",
    ("nakadori", "滋賀倉庫管理シート"): "K",
    ("maruun", "管理シート"): "L",
}

# Office Script と同じ列マッピング（固定・検証済み）
SHEET_CONFIG: dict[str, list[dict[str, Any]]] = {
    "matsuzaki": [
        {
            "sheet": "車両依頼書",
            "anNo": "B",
            "anName": "H",
            "ship": "D",
            "arr": "E",
            "model": "O",
            "car": ["K", "L", "M"],
            "arrTime": "F",
            "carNo": "N",
            "legs": [{"company": "S", "driver": "T", "plate": "U", "phone": "V"}],
        }
    ],
    "fukuoka": [
        {
            "sheet": "車両依頼書",
            "anNo": "B",
            "anName": "G",
            "ship": "C",
            "arr": "D",
            "model": "S",
            "car": ["J", "K", "L"],
            "arrTime": "E",
            "carNo": "M",
            "legs": [
                {"company": "Y", "driver": "Z", "plate": "AA", "phone": "AB"},
                {"company": "AC", "driver": "AD", "plate": "AE", "phone": "AF"},
            ],
        }
    ],
    "nakadori": [
        {
            "sheet": "車両依頼書",
            "anNo": "B",
            "anName": "G",
            "ship": "C",
            "arr": "D",
            "model": "R",
            "car": ["J", "L"],
            "arrTime": "E",
            "carNo": "M",
            "legs": [
                {"company": "W", "driver": "X", "plate": "Y", "phone": "Z"},
                {"company": "AA", "driver": "AB", "plate": "AC", "phone": "AD"},
            ],
        },
        {
            "sheet": "滋賀倉庫管理シート",
            "anNo": "B",
            "anName": "K",
            "ship": "E",
            "arr": "F",
            "model": "T",
            "car": ["M", "N", "O"],
            "arrTime": "G",
            "carNo": "P",
            "legs": [{"company": "U", "driver": "V", "plate": "W", "phone": "X"}],
        },
    ],
    "maruun": [
        {
            "sheet": "管理シート",
            "anNo": "B",
            "anName": "L",
            "ship": "E",
            "arr": "F",
            "model": "R",
            "car": ["N", "O", "P"],
            "arrTime": "G",
            "carNo": "Q",
            "legs": [{"company": "S", "driver": "T", "plate": "U", "phone": "V"}],
        }
    ],
}

OUTPUT_FIELD_MAP = {
    "案件No": "Title",
    "案件名": "案件名",
    "出荷日": "出荷日",
    "着日": "着日",
    "型式": "型式",
    "車型": "車型",
    "会社名": "会社名",
    "乗務員": "乗務員",
    "車番": "車番",
    "携帯番号": "携帯番号",
    "備考": "備考",
    "行キー": "行キー",
}


def is_unattended() -> bool:
    flag = os.environ.get("DRIVER_SYNC_UNATTENDED", "").strip().lower()
    if flag in ("1", "true", "yes", "on"):
        return True
    return not sys.stdin.isatty()


def rotate_log_if_needed(max_bytes: int = 2_097_152) -> None:
    for path in (LOG_PATH, TASK_LOG_PATH):
        if path.exists() and path.stat().st_size > max_bytes:
            backup = path.with_suffix(path.suffix + ".old")
            if backup.exists():
                backup.unlink()
            path.rename(backup)


def setup_logging(quiet_console: bool = False) -> None:
    rotate_log_if_needed()
    handlers: list[logging.Handler] = [
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
    ]
    console_handler = logging.StreamHandler(sys.stdout)
    if quiet_console:
        # 無人実行時は通常ログを抑制するが、警告・エラーはCI(GitHub Actions等)の
        # ログで原因調査できるよう必ず出力する。
        console_handler.setLevel(logging.WARNING)
    handlers.append(console_handler)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=handlers,
        force=True,
    )


def write_status(
    *,
    ok: bool,
    message: str,
    row_count: int | None = None,
    last_data_rows: int | None = None,
    last_patched_end_row: int | None = None,
    last_sheet_used_end_row: int | None = None,
    error: str | None = None,
) -> None:
    payload: dict[str, Any] = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "ok": ok,
        "message": message,
        "row_count": row_count,
        "last_data_rows": last_data_rows if last_data_rows is not None else row_count,
        "last_patched_end_row": last_patched_end_row,
        "last_sheet_used_end_row": last_sheet_used_end_row,
        "error": error,
        "unattended": is_unattended(),
    }
    STATUS_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


_lock_fd: int | None = None


def acquire_process_lock() -> None:
    global _lock_fd
    if LOCK_PATH.exists():
        try:
            age = time.time() - LOCK_PATH.stat().st_mtime
            if age > 3600:
                LOCK_PATH.unlink(missing_ok=True)
        except OSError:
            pass
    _lock_fd = os.open(str(LOCK_PATH), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    os.write(_lock_fd, str(os.getpid()).encode("ascii"))
    atexit.register(release_process_lock)


def release_process_lock() -> None:
    global _lock_fd
    if _lock_fd is not None:
        try:
            os.close(_lock_fd)
        except OSError:
            pass
        _lock_fd = None
    try:
        LOCK_PATH.unlink(missing_ok=True)
    except OSError:
        pass


def load_config() -> dict[str, Any]:
    with CONFIG_PATH.open(encoding="utf-8") as f:
        return json.load(f)


def col_idx(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def col_letter_from_index(index: int) -> str:
    result = ""
    n = index
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def normalize_header(text: str) -> str:
    t = str(text).strip().lower()
    for ch in ("．", "。", " ", "　", "\n", "\r"):
        t = t.replace(ch, "")
    return t


def build_header_grid(ws: Any, header_rows: int) -> list[list[Any]]:
    return [
        list(row)
        for row in ws.iter_rows(
            min_row=1,
            max_row=header_rows,
            max_col=39,
            values_only=True,
        )
    ]


def header_is_excluded(norm: str, exclude: set[str]) -> bool:
    if norm in exclude:
        return True
    return any(ex in norm or norm in ex for ex in exclude)


def find_column_by_headers_grid(
    header_grid: list[list[Any]],
    field_key: str,
) -> str | None:
    spec = HEADER_FIELD_ALIASES[field_key]
    labels = [normalize_header(x) for x in spec["labels"]]
    exclude = {normalize_header(x) for x in spec["exclude"]}
    best: tuple[float, str] | None = None
    for row_idx, row in enumerate(header_grid):
        for col_idx, val in enumerate(row):
            if val is None or str(val).strip() == "":
                continue
            norm = normalize_header(str(val))
            if header_is_excluded(norm, exclude):
                continue
            for priority, label in enumerate(labels):
                if norm == label:
                    score = priority + (len(header_grid) - row_idx) * 0.01
                elif label in norm or norm in label:
                    # 部分一致は優先度を下げる（納入先住所↔納入先名の誤爆防止）
                    score = priority + 50 + (len(header_grid) - row_idx) * 0.01
                else:
                    continue
                letter = col_letter_from_index(col_idx + 1)
                if best is None or score < best[0]:
                    best = (score, letter)
    return best[1] if best else None


def find_column_by_exact_header(
    header_grid: list[list[Any]],
    header_name: str,
) -> str | None:
    target = normalize_header(header_name)
    for row_idx, row in enumerate(header_grid):
        for col_idx, val in enumerate(row):
            if val is None:
                continue
            if normalize_header(str(val)) == target:
                return col_letter_from_index(col_idx + 1)
    return None


def detect_field_by_exact_label(
    header_grid: list[list[Any]],
    field_key: str,
) -> str | None:
    """見出しの完全一致のみで列を特定。

    完全一致に限定することで、元データに列が挿入されても見出しごと
    追従でき、かつ部分一致による誤爆（例: 会社名↔会社）を防げる。
    """
    spec = HEADER_FIELD_ALIASES.get(field_key)
    if not spec:
        return None
    exclude = {normalize_header(x) for x in spec.get("exclude", [])}
    for label in spec["labels"]:
        target = normalize_header(label)
        for row in header_grid:
            for col_index, val in enumerate(row):
                if val is None:
                    continue
                norm = normalize_header(str(val))
                if norm != target or header_is_excluded(norm, exclude):
                    continue
                return col_letter_from_index(col_index + 1)
    return None


def resolve_sheet_map(
    header_grid: list[list[Any]],
    sheet_map: dict[str, Any],
) -> dict[str, Any]:
    resolved = dict(sheet_map)
    # 案件No・出荷日・着日・型式は見出しのファジー一致で検出
    for field in ("anNo", "ship", "arr", "model"):
        detected = find_column_by_headers_grid(header_grid, field)
        if detected:
            resolved[field] = detected

    # 案件名(届け先名)は完全一致のみで上書き（納入先住所への誤爆を防止）
    name_col = detect_field_by_exact_label(header_grid, "anName")
    if name_col:
        resolved["anName"] = name_col
        resolved["_auto_anName"] = name_col

    # 単一legのシートのみ、ドライバー4項目を完全一致で上書きし列挿入に追従。
    # 複数leg（見出しが繰り返すシート）は曖昧になるため固定設定を維持する。
    legs = resolved.get("legs", [])
    if len(legs) == 1:
        new_leg = dict(legs[0])
        for field in ("company", "driver", "plate", "phone"):
            detected = detect_field_by_exact_label(header_grid, field)
            if detected:
                new_leg[field] = detected
        resolved["legs"] = [new_leg]

    return resolved


def looks_like_address(text: str) -> bool:
    if not text:
        return False
    return bool(ADDRESS_PATTERN.search(text))


def validate_row_quality(rows: list[dict[str, Any]]) -> None:
    if not rows:
        return
    named = [r.get("案件名", "") for r in rows if r.get("案件名")]
    if not named:
        return
    address_like = sum(1 for name in named if looks_like_address(name))
    ratio = address_like / len(named)
    if ratio >= 0.5:
        logging.warning(
            "案件名の約%.0f%%が住所らしい文字列です。"
            "見出し「届け先名」の列を使っていますが、元データが住所のみの可能性があります。"
            "サンプル: %s",
            ratio * 100,
            named[0][:60],
        )
    elif ratio >= 0.2:
        logging.info(
            "案件名の一部(%.0f%%)が住所形式です（届け先名＝現場住所の行がある想定内）",
            ratio * 100,
        )


def raw_cell(row: tuple[Any, ...], letter: str) -> Any:
    """型変換せずセルの値そのもの（datetime/time/数値など）を返す。"""
    if not letter:
        return None
    idx = col_idx(letter)
    if idx < 0 or idx >= len(row):
        return None
    return row[idx]


def cell(row: tuple[Any, ...], letter: str) -> str:
    if not letter:
        return ""
    idx = col_idx(letter)
    if idx < 0 or idx >= len(row):
        return ""
    value = row[idx]
    if value is None:
        return ""
    return str(value).strip()


def cell_lines(row: tuple[Any, ...], letter: str) -> list[str]:
    raw = cell(row, letter)
    if not raw:
        return []
    return [p.strip() for p in raw.replace("\r\n", "\n").split("\n") if p.strip()]


def first_line_field(row: tuple[Any, ...], letter: str) -> str:
    """セル内改行があっても先頭行だけ使う（ドライバー列の副行は無視）。"""
    lines = cell_lines(row, letter)
    if lines:
        return lines[0]
    return cell(row, letter)


def normalize_an_no_value(raw: str) -> str:
    text = str(raw or "").strip()
    if not text:
        return ""
    return text.replace("\r\n", "\n").split("\n")[0].strip()


def build_merged_cell_lookup(ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[tuple[int, int], Any]:
    """マージセル内の全セル → 左上セルの値（ドライバー列の取りこぼし防止）。"""
    lookup: dict[tuple[int, int], Any] = {}
    for merged_range in ws.merged_cells.ranges:
        top_value = ws.cell(merged_range.min_row, merged_range.min_col).value
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                lookup[(row, col)] = top_value
    return lookup


def read_row_with_merges(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    excel_row_num: int,
    max_col: int,
    merge_lookup: dict[tuple[int, int], Any],
) -> tuple[Any, ...]:
    values: list[Any] = []
    for col in range(1, max_col + 1):
        if (excel_row_num, col) in merge_lookup:
            values.append(merge_lookup[(excel_row_num, col)])
        else:
            values.append(ws.cell(excel_row_num, col).value)
    return tuple(values)


def case_name_column_letter(
    config_key: str,
    sheet_key: str,
    sheet_map: dict[str, Any],
) -> str:
    auto_name = sheet_map.get("_auto_anName")
    if auto_name:
        return auto_name
    letter = CASE_NAME_COLUMN.get((config_key, sheet_key))
    if letter is None and config_key == "nakadori" and "滋賀" in sheet_key and "管理" in sheet_key:
        letter = "K"
    if letter is None and config_key == "maruun" and sheet_key == "管理シート":
        letter = "L"
    if letter is None:
        letter = sheet_map["anName"]
    return letter


def physical_case_column_letters(
    sheet_map: dict[str, Any],
    case_name_letter: str,
) -> list[str]:
    letters = [
        sheet_map["anNo"],
        case_name_letter,
        sheet_map["ship"],
        sheet_map["arr"],
        sheet_map["model"],
        *sheet_map["car"],
    ]
    seen: set[str] = set()
    ordered: list[str] = []
    for letter in letters:
        if letter not in seen:
            seen.add(letter)
            ordered.append(letter)
    return ordered


def apply_physical_case_cells(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    excel_row_num: int,
    row_tuple: tuple[Any, ...],
    case_letters: list[str],
) -> tuple[Any, ...]:
    """案件側の列は当該行の物理セルのみ使う（上段マージの案件名流れ込みを防ぐ）。"""
    data = extend_row_tuple(row_tuple, max_col_for_letters(case_letters))
    for letter in case_letters:
        col = col_idx(letter) + 1
        data[col_idx(letter)] = ws.cell(excel_row_num, col).value
    return tuple(data)


def max_col_for_letters(letters: list[str]) -> int:
    if not letters:
        return 0
    return max(col_idx(letter) for letter in letters) + 1


def resolve_case_name_for_part(
    row_tuple: tuple[Any, ...],
    sheet_map: dict[str, Any],
    config_key: str,
    actual_sheet: str,
    *,
    part_index: int = 0,
) -> str:
    """案件名は同一行の名称列からそのまま読む（複数行セルは改行を空白に正規化）。
    中継倉庫名（「○○倉庫」のような行）は実際の届け先ではないため除外する。"""
    letter = case_name_column_letter(config_key, actual_sheet.strip(), sheet_map)
    raw = cell(row_tuple, letter)
    if not raw:
        return ""
    parts = [p.strip() for p in raw.replace("\r\n", "\n").split("\n") if p.strip()]
    parts = [p for p in parts if "倉庫" not in p] or parts
    return " ".join(parts)


def resolve_case_name(
    row_tuple: tuple[Any, ...],
    sheet_map: dict[str, Any],
    config_key: str,
    actual_sheet: str,
) -> str:
    return resolve_case_name_for_part(
        row_tuple, sheet_map, config_key, actual_sheet, part_index=0
    )


def row_output_signature(
    row_tuple: tuple[Any, ...],
    sheet_map: dict[str, Any],
    leg: dict[str, str],
) -> tuple[str, ...]:
    """連続行のマージセル重複出力を抑止するための比較キー。"""
    return (
        cell(row_tuple, sheet_map["anNo"]),
        cell(row_tuple, sheet_map["ship"]),
        cell(row_tuple, sheet_map["arr"]),
        *leg_fields_at(row_tuple, leg, 0).values(),
    )


def append_record_from_row_tuple(
    excel_row_num: int,
    row_tuple: tuple[Any, ...],
    *,
    effective_map: dict[str, Any],
    config_key: str,
    actual_name: str,
    today: date,
    arr_days_back: int,
    primary: dict[str, str],
    results: list[dict[str, Any]],
    part_index: int,
    an_no: str,
    model_value: str,
) -> None:
    driver = leg_fields_at(row_tuple, primary, 0)
    row_key = f"{config_key}|{actual_name}|{excel_row_num}"
    if part_index:
        row_key = f"{row_key}|{part_index}"

    has_relay = len(effective_map["legs"]) > 1 and leg_complete_at(
        row_tuple, effective_map["legs"][1], 0
    )
    remarks = build_remarks(
        cell(row_tuple, effective_map.get("carNo", "")),
        raw_cell(row_tuple, effective_map.get("arrTime", "")),
    )
    if has_relay:
        relay = leg_fields_at(row_tuple, effective_map["legs"][1], 0)
        relay_text = (
            f"2次配送（{relay['company']} / {relay['driver']} / "
            f"{relay['plate']} / {relay['phone']}）"
        )
        remarks = f"{remarks}\n{relay_text}" if remarks else relay_text

    results.append(
        {
            "案件No": an_no,
            "案件名": resolve_case_name_for_part(
                row_tuple,
                effective_map,
                config_key,
                actual_name,
                part_index=part_index,
            ),
            "出荷日": to_md(cell(row_tuple, effective_map["ship"]), today),
            "着日": to_md(cell(row_tuple, effective_map["arr"]), today),
            "型式": model_value,
            "車型": join_car_row(row_tuple, effective_map["car"]),
            "会社名": driver["company"],
            "乗務員": driver["driver"],
            "車番": driver["plate"],
            "携帯番号": driver["phone"],
            "備考": remarks,
            "中継あり": has_relay,
            "依頼先": config_key,
            "案件行": excel_row_num,
            "ドライバー行": excel_row_num,
            "行キー": row_key,
        }
    )


def append_records_from_physical_row(
    excel_row_num: int,
    row_tuple: tuple[Any, ...],
    *,
    effective_map: dict[str, Any],
    config_key: str,
    actual_name: str,
    today: date,
    arr_days_back: int,
    primary: dict[str, str],
    results: list[dict[str, Any]],
) -> None:
    """1 Excel 行を軸に、指定列から案件情報とドライバー情報をセットで抽出。"""
    if not has_case_on_row(row_tuple, effective_map):
        return

    an_no_lines = cell_lines(row_tuple, effective_map["anNo"])
    if not an_no_lines:
        return

    arr_date = to_date(cell(row_tuple, effective_map["arr"]), today)
    if not arr_date or not is_arr_in_sync_window(arr_date, today, arr_days_back):
        return

    if not leg_complete_at(row_tuple, primary, 0):
        logging.warning(
            "ドライバー4項目未入力: %s/%s 行%d 出荷日=%s 案件=%s",
            config_key,
            actual_name,
            excel_row_num,
            to_md(cell(row_tuple, effective_map["ship"]), today),
            " / ".join(an_no_lines)[:80],
        )
        return

    # 1物理行＝1レコード。複数案件Noは改行区切りで1セルにまとめる
    # （案件情報・ドライバー情報は必ず同一行から読むのでズレない）。
    an_no_value = "\n".join(an_no_lines)
    append_record_from_row_tuple(
        excel_row_num,
        row_tuple,
        effective_map=effective_map,
        config_key=config_key,
        actual_name=actual_name,
        today=today,
        arr_days_back=arr_days_back,
        primary=primary,
        results=results,
        part_index=0,
        an_no=an_no_value,
        model_value=model_text(row_tuple, effective_map["model"]),
    )


def row_has_mapped_content(row: tuple[Any, ...], sheet_map: dict[str, Any]) -> bool:
    letters = [
        sheet_map["anNo"],
        sheet_map["ship"],
        sheet_map["arr"],
        sheet_map["model"],
        *sheet_map["car"],
    ]
    for leg in sheet_map["legs"]:
        letters.extend([leg["company"], leg["driver"], leg["plate"], leg["phone"]])
    return any(cell(row, letter) for letter in letters)


def leg_complete_at(row: tuple[Any, ...], leg: dict[str, str], index: int) -> bool:
    return all(
        first_line_field(row, leg[k]) for k in ("company", "driver", "plate", "phone")
    )


def leg_fields_at(
    row: tuple[Any, ...], leg: dict[str, str], index: int
) -> dict[str, str]:
    return {k: first_line_field(row, leg[k]) for k in ("company", "driver", "plate", "phone")}


def model_text(row: tuple[Any, ...], letter: str) -> str:
    """型式はセル内の複数行をそのまま改行区切りで1フィールドにする。"""
    raw = cell(row, letter)
    if not raw:
        return ""
    parts = [p.strip() for p in raw.replace("\r\n", "\n").split("\n") if p.strip()]
    return "\n".join(parts)


def join_car_row(row: tuple[Any, ...], cols: list[str]) -> str:
    """車型列を改行分割せず結合。"""
    parts = [cell(row, c) for c in cols if cell(row, c)]
    return " ".join(parts)


def extend_row_tuple(row: tuple[Any, ...], min_len: int) -> list[Any]:
    data = list(row)
    if len(data) < min_len:
        data.extend([None] * (min_len - len(data)))
    return data


def has_case_on_row(row: tuple[Any, ...], sheet_map: dict[str, Any]) -> bool:
    return bool(cell(row, sheet_map["anNo"]))


def to_date(raw: Any, today: date) -> date | None:
    if raw is None or raw == "":
        return None
    if isinstance(raw, datetime):
        return raw.date()
    if isinstance(raw, date):
        return raw
    if isinstance(raw, (int, float)) and raw > 0:
        base = date(1899, 12, 30)
        return base + timedelta(days=int(raw))
    text = str(raw).strip()
    if not text:
        return None
    # 時刻部分・曜日表記を除去（例: "2025/5/1 0:00:00", "2025/5/1(木)"）
    text = text.split(" ")[0].split("T")[0]
    if "(" in text:
        text = text.split("(")[0]
    for ch in ("-", ".", "年", "月"):
        text = text.replace(ch, "/")
    text = text.replace("日", "")
    parts = [int(p) for p in text.split("/") if p.strip().isdigit()]
    if not parts:
        return None
    try:
        if len(parts) >= 3:
            return date(parts[0], parts[1], parts[2])
        if len(parts) == 2:
            return date(today.year, parts[0], parts[1])
    except ValueError:
        return None
    return None


def to_md(raw: Any, today: date) -> str:
    d = to_date(raw, today)
    if not d:
        return "" if raw is None else str(raw).strip()
    return f"{d.month}/{d.day}"


def format_time_value(raw: Any) -> str:
    """着時間セルの値を "HH:MM" 表記に揃える。"""
    if raw is None or raw == "":
        return ""
    if isinstance(raw, datetime):
        return raw.strftime("%H:%M")
    if isinstance(raw, dt_time):
        return raw.strftime("%H:%M")
    if isinstance(raw, (int, float)):
        total_minutes = round(raw * 24 * 60)
        hour, minute = divmod(total_minutes, 60)
        return f"{hour:02d}:{minute:02d}"
    text = str(raw).strip()
    match = re.match(r"^(\d{1,2}):(\d{2})", text)
    if match:
        return f"{int(match.group(1)):02d}:{match.group(2)}"
    return text


def build_remarks(car_no_raw: str, arr_time_raw: Any) -> str:
    """備考列: "1号車 / 10:00着" の形式で号車と着時間をまとめる。"""
    car_no = str(car_no_raw).strip() if car_no_raw else ""
    parts = []
    if car_no:
        parts.append(car_no if car_no.endswith("号車") else f"{car_no}号車")
    time_str = format_time_value(arr_time_raw)
    if time_str:
        parts.append(f"{time_str}着")
    return " / ".join(parts)


def arr_window_start(today: date, days_back: int) -> date:
    """着日下限: 今日から days_back 日前（暦日）。"""
    if days_back <= 0:
        return today
    return today - timedelta(days=days_back)


def arr_window_end(today: date) -> date:
    """着日上限: 明後日（暦日）。ただし今日が金曜日の場合は、土日を挟むため
    月曜（祝日なら火曜）まで延長する。"""
    if today.weekday() == 4:  # 0=月 ... 4=金
        monday = today + timedelta(days=3)
        if jpholiday.is_holiday(monday):
            return monday + timedelta(days=1)
        return monday
    return today + timedelta(days=2)


def is_arr_in_sync_window(
    arr_date: date,
    today: date,
    days_back: int = 7,
) -> bool:
    """着日が「N日前（暦日）〜 明日」の範囲内なら表示対象。"""
    window_start = arr_window_start(today, days_back)
    window_end = arr_window_end(today)
    return window_start <= arr_date <= window_end


def max_col_for_map(sheet_map: dict[str, Any]) -> int:
    letters = [
        sheet_map["anNo"],
        sheet_map["anName"],
        sheet_map["ship"],
        sheet_map["arr"],
        sheet_map["model"],
        *sheet_map["car"],
    ]
    if sheet_map.get("arrTime"):
        letters.append(sheet_map["arrTime"])
    if sheet_map.get("carNo"):
        letters.append(sheet_map["carNo"])
    for leg in sheet_map["legs"]:
        letters.extend([leg["company"], leg["driver"], leg["plate"], leg["phone"]])
    return max(col_idx(letter) for letter in letters) + 1


def parse_row_key_parts(row_key: str) -> tuple[str, int, int]:
    parts = row_key.split("|")
    src = parts[0] if parts else ""
    try:
        excel_row = int(parts[2]) if len(parts) > 2 else 99999
    except ValueError:
        excel_row = 99999
    try:
        part_idx = int(parts[3]) if len(parts) > 3 else 0
    except ValueError:
        part_idx = 0
    return src, excel_row, part_idx


def validate_output_rows_integrity(rows: list[dict[str, Any]]) -> None:
    """案件情報とドライバー情報が同一出力行に揃っているか検証。"""
    issues = 0
    for idx, row in enumerate(rows, start=1):
        required = ("案件No", "案件名", "会社名", "乗務員", "車番", "携帯番号")
        missing = [name for name in required if not str(row.get(name, "")).strip()]
        if missing:
            issues += 1
            logging.warning(
                "出力行%d 整合性NG(欠落=%s): %s %s",
                idx,
                ",".join(missing),
                row.get("案件No", ""),
                row.get("行キー", ""),
            )
        case_row = row.get("案件行")
        driver_row = row.get("ドライバー行")
        if (
            isinstance(case_row, int)
            and isinstance(driver_row, int)
            and case_row != driver_row
        ):
            issues += 1
            logging.warning(
                "出力行%d 行不一致(案件行=%d ドライバー行=%d): %s %s",
                idx,
                case_row,
                driver_row,
                row.get("案件No", ""),
                row.get("行キー", ""),
            )
    if issues:
        logging.warning("出力整合性チェック: %d件に問題あり", issues)
    else:
        logging.info("出力整合性チェック: 全%d件OK（案件行=ドライバー行）", len(rows))


def _md_to_sortable_date(md_str: str, today: date) -> date:
    """「6/15」のような月日表記を、todayに最も近い年で実際の日付に戻す（年境界対応）。"""
    if not md_str:
        return today
    try:
        month, day = (int(p) for p in md_str.split("/"))
    except (ValueError, AttributeError):
        return today
    candidates = [date(today.year + offset, month, day) for offset in (-1, 0, 1)]
    return min(candidates, key=lambda d: abs((d - today).days))


def prepare_rows_for_output(
    rows: list[dict[str, Any]], today: date | None = None
) -> list[dict[str, Any]]:
    """重複除去し、着日の新しい順に並べ替え（案件とドライバーは同じ行データのまま連動して動く）。"""
    if today is None:
        today = date.today()
    seen: set[str] = set()
    unique: list[dict[str, Any]] = []
    for row in rows:
        key = str(row.get("行キー", ""))
        if key and key in seen:
            continue
        if key:
            seen.add(key)
        unique.append(row)

    def sort_key(row: dict[str, Any]) -> tuple[date, date, int, int, str]:
        src, excel_row, part_idx = parse_row_key_parts(str(row.get("行キー", "")))
        src_idx = SOURCE_ORDER.index(src) if src in SOURCE_ORDER else 99
        arr_date = _md_to_sortable_date(str(row.get("着日", "")), today)
        ship_date = _md_to_sortable_date(str(row.get("出荷日", "")), today)
        # 着日→出荷日の新しい順（降順）。同日内は元の取得順（依頼先→Excel行）で安定させる。
        return (arr_date, ship_date, -src_idx, -excel_row, str(row.get("案件No", "")))

    return sorted(unique, key=sort_key, reverse=True)


def load_output_column_map(config: dict[str, Any]) -> dict[str, int]:
    od = config.get("onedrive_output", {})
    raw = od.get("output_column_map")
    if isinstance(raw, dict) and raw:
        mapping: dict[str, int] = {}
        for name, letter in raw.items():
            if name in EXCEL_OUTPUT_COLUMNS and isinstance(letter, str):
                mapping[name] = col_idx(letter) + 1
        if mapping:
            return mapping
    return {name: idx for idx, name in enumerate(EXCEL_OUTPUT_COLUMNS, start=1)}


def extract_rows_from_workbook(
    workbook: openpyxl.Workbook,
    config_key: str,
    header_rows: int,
    today: date,
    *,
    arr_days_back: int = 7,
    auto_detect_columns: bool = False,
) -> list[dict[str, Any]]:
    maps = SHEET_CONFIG.get(config_key, [])
    results: list[dict[str, Any]] = []

    for sheet_map in maps:
        sheet_name = sheet_map["sheet"]
        # シート名の前後空白ゆれを無視して照合（例: "滋賀管理シート "）
        actual_name = next(
            (s for s in workbook.sheetnames if s.strip() == sheet_name.strip()),
            None,
        )
        if actual_name is None and config_key == "nakadori" and "滋賀" in sheet_name:
            actual_name = next(
                (
                    s
                    for s in workbook.sheetnames
                    if "滋賀" in s and "管理" in s
                ),
                None,
            )
        if actual_name is None:
            logging.warning("シートが見つかりません: %s (%s)", sheet_name, config_key)
            continue

        ws = workbook[actual_name]
        effective_map = sheet_map
        sheet_key = actual_name.strip()
        case_col = CASE_NAME_COLUMN.get((config_key, sheet_key))
        if case_col is None and config_key == "nakadori" and "滋賀" in sheet_key and "管理" in sheet_key:
            case_col = "K"
        if case_col is None and config_key == "maruun" and sheet_key == "管理シート":
            case_col = "L"
        if case_col is None:
            case_col = effective_map.get("anName", "")
        logging.info(
            "列マッピング(固定): %s/%s 案件名=%s列 出荷=%s 着日=%s 型式=%s",
            config_key,
            actual_name,
            case_col,
            effective_map.get("ship"),
            effective_map.get("arr"),
            effective_map.get("model"),
        )
        if auto_detect_columns:
            header_grid = build_header_grid(ws, header_rows)
            effective_map = resolve_sheet_map(header_grid, sheet_map)
            if effective_map != sheet_map:
                logging.info(
                    "列マッピング自動調整: %s/%s 案件名=%s 出荷=%s 着日=%s",
                    config_key,
                    actual_name,
                    effective_map.get("anName"),
                    effective_map.get("ship"),
                    effective_map.get("arr"),
                )
        max_col = max_col_for_map(effective_map)
        start_row = header_rows + 1
        merge_lookup = build_merged_cell_lookup(ws)
        physical_rows: list[tuple[int, tuple[Any, ...]]] = []
        for excel_row_num in range(start_row, ws.max_row + 1):
            physical_rows.append(
                (
                    excel_row_num,
                    read_row_with_merges(ws, excel_row_num, max_col, merge_lookup),
                )
            )

        primary = effective_map["legs"][0]
        case_name_letter = case_name_column_letter(
            config_key, sheet_key, effective_map
        )
        case_letters = physical_case_column_letters(effective_map, case_name_letter)
        prev_excel_row = -1
        prev_signature: tuple[str, ...] | None = None

        for excel_row_num, row_tuple_merged in physical_rows:
            row_tuple = apply_physical_case_cells(
                ws, excel_row_num, row_tuple_merged, case_letters
            )
            if not row_has_mapped_content(row_tuple, effective_map):
                continue
            before = len(results)
            append_records_from_physical_row(
                excel_row_num,
                row_tuple,
                effective_map=effective_map,
                config_key=config_key,
                actual_name=actual_name,
                today=today,
                arr_days_back=arr_days_back,
                primary=primary,
                results=results,
            )
            if len(results) <= before:
                continue

            record = results[-1]
            signature = row_output_signature(row_tuple, effective_map, primary)
            if (
                prev_signature is not None
                and signature == prev_signature
                and excel_row_num == prev_excel_row + 1
            ):
                logging.info(
                    "マージセル重複行をスキップ: %s/%s Excel行%d 案件=%s",
                    config_key,
                    actual_name,
                    excel_row_num,
                    record.get("案件No", ""),
                )
                results.pop()
            else:
                prev_excel_row = excel_row_num
                prev_signature = signature

    return results


def encode_share_url(url: str) -> str:
    encoded = base64.b64encode(url.encode("utf-8")).decode("utf-8")
    encoded = encoded.rstrip("=").replace("/", "_").replace("+", "-")
    return f"u!{encoded}"


BROWSER_UA = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}


def download_share_file(share_url: str) -> bytes:
    """匿名共有リンク(1drv.ms)から直接ダウンロード（認証不要・元ファイルは読取のみ）。

    検証で確定した方式：リンク末尾に download=1 を付け、ブラウザ相当のUAで
    リダイレクトを追従すると、xlsx 本体が取得できる。
    """
    sep = "&" if "?" in share_url else "?"
    url = f"{share_url}{sep}download=1"
    res = requests.get(url, headers=BROWSER_UA, timeout=180, allow_redirects=True)
    if not res.ok:
        raise RuntimeError(f"ファイル取得失敗: {res.status_code} {res.text[:200]}")
    if res.content[:4] != b"PK\x03\x04":
        raise RuntimeError(
            "取得データがExcel(xlsx)ではありません。"
            f"先頭バイト={res.content[:8]!r} 最終URL={res.url[:120]}"
        )
    return res.content


def sharepoint_scopes(config: dict[str, Any]) -> list[str]:
    host = config["sharepoint_hostname"]
    return [f"https://{host}/AllSites.Write"]


def required_scopes(config: dict[str, Any]) -> list[str]:
    scopes: list[str] = []
    od = config.get("onedrive_output", {})
    if od.get("enabled", True):
        scopes.extend(GRAPH_SCOPES)
    if config.get("sync_sharepoint", False):
        for scope in sharepoint_scopes(config):
            if scope not in scopes:
                scopes.append(scope)
    return scopes or list(GRAPH_SCOPES)


def normalize_case_name_key(name: str) -> str:
    """案件名の空白ゆれを吸収（色分け・照合用）。"""
    text = str(name or "").replace("\r\n", "\n").split("\n")[0].strip()
    text = text.replace("\u3000", " ")
    return re.sub(r"\s+", " ", text)


def load_case_name_row_color_map(config: dict[str, Any]) -> dict[str, str]:
    """強調3種の案件名 → A〜E列の背景色。"""
    od = config.get("onedrive_output", {})
    raw = od.get("row_colors_by_case_name") or DEFAULT_CASE_NAME_ROW_COLORS
    mapping: dict[str, str] = {}
    for label, color in raw.items():
        key = normalize_case_name_key(label)
        if key and color:
            mapping[key] = str(color)
    return mapping


@dataclass(frozen=True)
class OutputRowColors:
    """1行分の案件ブロック色・ドライバーブロック色。"""

    default_case: str
    default_driver: str
    highlight_driver: str


def load_output_row_colors(config: dict[str, Any]) -> OutputRowColors:
    od = config.get("onedrive_output", {})
    return OutputRowColors(
        default_case=str(od.get("default_row_fill", DEFAULT_CASE_FILL_HEX)),
        default_driver=str(od.get("driver_block_fill", DEFAULT_DRIVER_FILL_HEX)),
        highlight_driver=str(
            od.get("driver_block_fill_highlight", HIGHLIGHT_DRIVER_FILL_HEX)
        ),
    )


def resolve_row_block_fills(
    case_name: str,
    color_map: dict[str, str],
    colors: OutputRowColors,
) -> tuple[str, str]:
    """案件名から (A〜E列の色, F〜K列の色) を返す。"""
    if is_highlight_case_name(case_name, color_map):
        case_fill = color_map.get(
            normalize_case_name_key(case_name), HIGHLIGHT_CASE_FILL_HEX
        )
        return case_fill, colors.highlight_driver
    return colors.default_case, colors.default_driver


def is_highlight_case_name(case_name: str, color_map: dict[str, str]) -> bool:
    return normalize_case_name_key(case_name) in color_map


def openpyxl_fill_from_hex(color_hex: str) -> PatternFill:
    hex6 = str(color_hex or "").lstrip("#").upper()
    if len(hex6) == 6:
        hex6 = f"FF{hex6}"
    return PatternFill(start_color=hex6, end_color=hex6, fill_type="solid")


def graph_patch_range_fill(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    address: str,
    color_hex: str,
    session_id: str,
) -> None:
    seg = worksheet_segment(sheet_name)
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"{seg}/range(address='{address}')/format/fill"
    )
    res = graph_request_with_retry(
        "PATCH",
        url,
        graph_token,
        session_id=session_id,
        json={"color": color_hex},
    )
    if not res.ok:
        raise RuntimeError(
            f"Excel塗りつぶし失敗({address}): {res.status_code} {res.text[:300]}"
        )


# $batch で一度に送るサブリクエスト上限（Graph 仕様の最大値）
GRAPH_BATCH_CHUNK = 20
# 一時エラー（ロック・スロットリング・サーバ側一時障害）として再試行する HTTP ステータス
GRAPH_BATCH_RETRY_STATUS = {423, 429, 500, 502, 503, 504}


def _post_fill_batch_with_retry(
    graph_token: str,
    requests_body: list[dict[str, Any]],
    session_id: str,
) -> None:
    """1 バッチ分の塗りつぶしを送信し、一時失敗のサブリクエストのみ再試行する。"""
    batch_url = f"{GRAPH_BASE}/$batch"
    pending = requests_body
    max_attempts = 8
    for attempt in range(1, max_attempts + 1):
        res = requests.post(
            batch_url,
            headers=graph_auth_headers(graph_token),
            json={"requests": pending},
            timeout=180,
        )
        if not res.ok:
            raise RuntimeError(
                f"$batch 送信失敗: {res.status_code} {res.text[:300]}"
            )
        responses = res.json().get("responses", [])
        failed = [r for r in responses if int(r.get("status", 200)) >= 400]
        if not failed:
            return
        hard = [
            r for r in failed
            if int(r.get("status", 0)) not in GRAPH_BATCH_RETRY_STATUS
        ]
        if hard:
            sample = hard[0]
            raise RuntimeError(
                f"塗りつぶし失敗(status={sample.get('status')}): "
                f"{str(sample.get('body'))[:200]}"
            )
        if attempt >= max_attempts:
            raise RuntimeError("塗りつぶしバッチが再試行上限に達しました")
        failed_ids = {r.get("id") for r in failed}
        wait_sec = 5 * attempt
        logging.warning(
            "塗りつぶし %d 件が一時失敗。%d 秒後に再試行します（%d/%d）。",
            len(failed_ids),
            wait_sec,
            attempt,
            max_attempts,
        )
        time.sleep(wait_sec)
        pending = [r for r in pending if r.get("id") in failed_ids]


def graph_batch_patch_fills(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    fills: list[tuple[str, str]],
    session_id: str,
) -> None:
    """複数セル範囲の背景色を $batch でまとめて適用する。

    1 範囲ごとに個別 PATCH すると API 往復が膨大になり遅い（余剰行の色塗りで
    顕著）。$batch（複数リクエストを 1 回の HTTP にまとめる Graph の仕組み）で
    最大 20 件ずつ送り、往復回数を約 1/20 に削減する。
    """
    if not fills:
        return
    seg = worksheet_segment(sheet_name)
    for start in range(0, len(fills), GRAPH_BATCH_CHUNK):
        chunk = fills[start:start + GRAPH_BATCH_CHUNK]
        requests_body = [
            {
                "id": str(i),
                "method": "PATCH",
                "url": (
                    f"/me/drive/items/{item_id}/workbook/"
                    f"{seg}/range(address='{address}')/format/fill"
                ),
                "headers": {
                    "Content-Type": "application/json",
                    "workbook-session-id": session_id,
                },
                "body": {"color": color_hex},
            }
            for i, (address, color_hex) in enumerate(chunk)
        ]
        _post_fill_batch_with_retry(graph_token, requests_body, session_id)


def graph_batch_clear_fills(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    addresses: list[str],
    session_id: str,
) -> None:
    """指定範囲の背景色を「塗りつぶしなし」に戻す（中継ありの一次配送行用）。"""
    if not addresses:
        return
    seg = worksheet_segment(sheet_name)
    for start in range(0, len(addresses), GRAPH_BATCH_CHUNK):
        chunk = addresses[start:start + GRAPH_BATCH_CHUNK]
        requests_body = [
            {
                "id": str(i),
                "method": "POST",
                "url": (
                    f"/me/drive/items/{item_id}/workbook/"
                    f"{seg}/range(address='{address}')/format/fill/clear"
                ),
                "headers": {
                    "Content-Type": "application/json",
                    "workbook-session-id": session_id,
                },
            }
            for i, address in enumerate(chunk)
        ]
        _post_fill_batch_with_retry(graph_token, requests_body, session_id)


def case_block_col_range(col_map: dict[str, int]) -> tuple[int, int]:
    """A列(案件No)〜E列(型式)の列番号範囲。"""
    cols = [col_map[name] for name in CASE_BLOCK_COLUMN_NAMES if name in col_map]
    if not cols:
        return 1, 5
    return min(cols), max(cols)


def driver_block_col_range(col_map: dict[str, int]) -> tuple[int, int]:
    """F列(車型)〜K列(中継あり)の列番号範囲。"""
    cols = [col_map[name] for name in DRIVER_BLOCK_COLUMN_NAMES if name in col_map]
    if not cols:
        return 6, 11
    return min(cols), max(cols)


DIFF_STATE_SHEET = "_diff_state"
DRIVER_SNAPSHOT_FIELDS = ("車型", "会社名", "乗務員", "車番", "携帯番号")


def driver_block_snapshot(row: dict[str, Any]) -> str:
    return "|".join(str(row.get(f, "") or "") for f in DRIVER_SNAPSHOT_FIELDS)


def diff_state_case_key(row: dict[str, Any]) -> str:
    # 同じ案件Noが複数の物理行（別配送）に存在することがあるため、
    # 案件Noではなく行単位で一意な「行キー」を使う（取り違えによる誤検出防止）。
    row_key = str(row.get("行キー", "") or "")
    if row_key:
        return row_key
    raw = str(row.get("案件No", "") or "")
    return raw.splitlines()[0].strip() if raw else ""


def graph_ensure_diff_state_sheet(
    graph_token: str, item_id: str, session_id: str
) -> str:
    """ドライバー情報の変更検出用の非表示シートが無ければ作成する。"""
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/worksheets"
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    names = [w["name"] for w in res.json().get("value", [])] if res.ok else []
    if DIFF_STATE_SHEET in names:
        return DIFF_STATE_SHEET
    res = graph_request_with_retry(
        "POST", url, graph_token, session_id=session_id,
        json={"name": DIFF_STATE_SHEET},
    )
    if not res.ok:
        raise RuntimeError(
            f"_diff_state シート作成失敗: {res.status_code} {res.text[:200]}"
        )
    hide_url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"worksheets('{DIFF_STATE_SHEET}')"
    )
    graph_request_with_retry(
        "PATCH", hide_url, graph_token, session_id=session_id,
        json={"visibility": "Hidden"},
    )
    return DIFF_STATE_SHEET


def graph_load_diff_state(
    graph_token: str, item_id: str, session_id: str
) -> dict[str, tuple[str, bool]]:
    sheet = graph_ensure_diff_state_sheet(graph_token, item_id, session_id)
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"worksheets('{sheet}')/usedRange(valuesOnly=true)"
    )
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    state: dict[str, tuple[str, bool]] = {}
    if not res.ok:
        return state
    for row in res.json().get("values") or []:
        if not row or not row[0]:
            continue
        key = str(row[0]).strip()
        if not key:
            continue
        snapshot = str(row[1]) if len(row) > 1 and row[1] is not None else ""
        changed = bool(len(row) > 2 and str(row[2]).strip() in ("1", "TRUE", "True"))
        state[key] = (snapshot, changed)
    return state


def graph_save_diff_state(
    graph_token: str,
    item_id: str,
    session_id: str,
    state: dict[str, tuple[str, bool]],
) -> None:
    sheet = graph_ensure_diff_state_sheet(graph_token, item_id, session_id)
    base = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/worksheets('{sheet}')"
    )
    graph_request_with_retry(
        "POST", f"{base}/range(address='A1:C5000')/clear", graph_token,
        session_id=session_id, json={"applyTo": "Contents"},
    )
    if not state:
        return
    rows_out = [
        [key, snapshot, "1" if changed else ""]
        for key, (snapshot, changed) in state.items()
    ]
    address = f"A1:C{len(rows_out)}"
    graph_request_with_retry(
        "PATCH", f"{base}/range(address='{address}')", graph_token,
        session_id=session_id, json={"values": rows_out},
    )


def apply_driver_change_detection(
    graph_token: str,
    item_id: str,
    session_id: str,
    rows: list[dict[str, Any]],
) -> None:
    """ドライバー情報(車型/会社名/乗務員/車番/携帯番号)が前回から変わった案件を検出し、
    一度変わった案件は以後ずっと黄色で表示し続けるためのフラグを各行に付与する。"""
    state = graph_load_diff_state(graph_token, item_id, session_id)
    new_state: dict[str, tuple[str, bool]] = {}
    changed_count = 0
    for row in rows:
        key = diff_state_case_key(row)
        if not key:
            row["_driver_changed"] = False
            continue
        snapshot = driver_block_snapshot(row)
        prev_snapshot, prev_changed = state.get(key, ("", False))
        changed = prev_changed or (bool(prev_snapshot) and prev_snapshot != snapshot)
        row["_driver_changed"] = changed
        new_state[key] = (snapshot, changed)
        if changed:
            changed_count += 1
    graph_save_diff_state(graph_token, item_id, session_id, new_state)
    if changed_count:
        logging.info("ドライバー情報の変更検出: %d件を黄色表示", changed_count)


def apply_output_row_colors(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
    rows: list[dict[str, Any]],
    *,
    data_start_row: int,
    col_map: dict[str, int],
    config: dict[str, Any],
) -> None:
    """データ行の背景色を塗る（A〜E=案件情報、F〜K=ドライバー情報）。"""
    color_map = load_case_name_row_color_map(config)
    colors = load_output_row_colors(config)
    case_min, case_max = case_block_col_range(col_map)
    driver_min, driver_max = driver_block_col_range(col_map)

    fills: list[tuple[str, str]] = []
    clear_addresses: list[str] = []
    for idx, row in enumerate(rows):
        excel_row = data_start_row + idx
        case_name = str(row.get("案件名", ""))
        case_fill, driver_fill = resolve_row_block_fills(case_name, color_map, colors)
        fills.append(
            (range_address(case_min, case_max, excel_row, excel_row), case_fill)
        )
        driver_address = range_address(driver_min, driver_max, excel_row, excel_row)
        if row.get("_driver_changed"):
            # ドライバー情報が前回から変わった案件は、中継ありの塗りなし表示より
            # 優先して黄色で目立たせる（一度変わったら以後ずっと黄色のまま）。
            fills.append((driver_address, CHANGED_DRIVER_FILL_HEX))
        elif row.get("中継あり"):
            # 中継あり（二次配送が別行に発生する）一次配送行は、ドライバー情報ブロックを
            # 塗りつぶしなしにして区別する。
            clear_addresses.append(driver_address)
        else:
            fills.append((driver_address, driver_fill))

        for white_col_name in WHITE_FILL_COLUMN_NAMES:
            white_col = col_map.get(white_col_name)
            if white_col:
                clear_addresses.append(
                    range_address(white_col, white_col, excel_row, excel_row)
                )

    graph_batch_patch_fills(graph_token, item_id, sheet_name, fills, session_id)
    graph_batch_clear_fills(graph_token, item_id, sheet_name, clear_addresses, session_id)

    highlight_case = next(iter(color_map.values()), HIGHLIGHT_CASE_FILL_HEX)
    logging.info(
        "行色適用: %d行（強調 A〜E=%s F〜K=%s / その他 A〜E=%s F〜K=%s / 中継あり%d行はF〜K塗りつぶしなし）",
        len(rows),
        highlight_case,
        colors.highlight_driver,
        colors.default_case,
        colors.default_driver,
        len(clear_addresses),
    )


def apply_row_fills_openpyxl(
    ws: Any,
    excel_row: int,
    case_name: str,
    *,
    case_min: int,
    case_max: int,
    driver_min: int,
    driver_max: int,
    color_map: dict[str, str],
    colors: OutputRowColors,
) -> None:
    case_hex, driver_hex = resolve_row_block_fills(case_name, color_map, colors)
    case_fill = openpyxl_fill_from_hex(case_hex)
    driver_fill = openpyxl_fill_from_hex(driver_hex)
    for col in range(case_min, case_max + 1):
        ws.cell(excel_row, col).fill = case_fill
    for col in range(driver_min, driver_max + 1):
        ws.cell(excel_row, col).fill = driver_fill


def rows_to_xlsx_bytes(rows: list[dict[str, Any]], config: dict[str, Any] | None = None) -> bytes:
    """新規ブックを作る（初回のみ・書式維持モード失敗時のフォールバック）。"""
    cfg = config or {}
    color_map = load_case_name_row_color_map(cfg)
    colors = load_output_row_colors(cfg)
    col_map = {name: idx for idx, name in enumerate(EXCEL_OUTPUT_COLUMNS, start=1)}
    case_min, case_max = case_block_col_range(col_map)
    driver_min, driver_max = driver_block_col_range(col_map)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ドライバー情報"
    ws.append(EXCEL_OUTPUT_COLUMNS)
    for row in rows:
        ws.append([row.get(col, "") for col in EXCEL_OUTPUT_COLUMNS])
        apply_row_fills_openpyxl(
            ws,
            ws.max_row,
            str(row.get("案件名", "")),
            case_min=case_min,
            case_max=case_max,
            driver_min=driver_min,
            driver_max=driver_max,
            color_map=color_map,
            colors=colors,
        )
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def cell_output_value(row: dict[str, Any], col_name: str) -> Any:
    value = row.get(col_name, "")
    if value is None:
        return ""
    if col_name == "依頼先":
        return SOURCE_DISPLAY_NAMES.get(str(value), str(value))
    return value


def load_last_data_rows() -> int:
    if not STATUS_PATH.exists():
        return 0
    try:
        data = json.loads(STATUS_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return 0
    return int(data.get("last_data_rows") or data.get("row_count") or 0)


def load_last_patched_end_row(data_start_row: int) -> int:
    """前回同期で書き込んだ最終行（Excel行番号）。未記録時は件数から推定。"""
    if not STATUS_PATH.exists():
        return data_start_row - 1
    try:
        data = json.loads(STATUS_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return data_start_row - 1
    if data.get("last_patched_end_row") is not None:
        return int(data["last_patched_end_row"])
    row_count = int(data.get("last_data_rows") or data.get("row_count") or 0)
    return output_data_end_row(data_start_row, row_count)


def load_last_sheet_used_end_row(data_start_row: int) -> int:
    """前回までに同期が触れた最終行（余剰行の背景色削除用）。"""
    if not STATUS_PATH.exists():
        return data_start_row - 1
    try:
        data = json.loads(STATUS_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return data_start_row - 1
    if data.get("last_sheet_used_end_row") is not None:
        return int(data["last_sheet_used_end_row"])
    return load_last_patched_end_row(data_start_row)


def output_data_end_row(data_start_row: int, row_count: int) -> int:
    """着日範囲内の件数から、データ最終行の Excel 行番号を求める。"""
    if row_count <= 0:
        return data_start_row - 1
    return data_start_row + row_count - 1


def resolve_padding_end_row(
    config: dict[str, Any],
    *,
    data_end_row: int,
    sheet_used_end_row: int,
    previous_end_row: int,
    used_last_row: int,
) -> int:
    """書込範囲外の交互色を塗る最終行（設定の padding_end_row まで最低限確保）。"""
    od = config.get("onedrive_output", {})
    min_end_row = int(od.get("padding_end_row", DEFAULT_PADDING_END_ROW))
    return max(
        sheet_used_end_row,
        previous_end_row,
        used_last_row,
        data_end_row,
        min_end_row,
    )


def graph_get_worksheet_used_last_row(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
    *,
    data_start_row: int,
) -> int:
    """シート usedRange の最終行（データ開始行未満なら 0）。"""
    seg = worksheet_segment(sheet_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/usedRange"
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    if not res.ok:
        return 0
    try:
        address = str(res.json().get("address", ""))
        _, _, _, max_row = parse_a1_range(address)
    except (ValueError, KeyError, TypeError):
        return 0
    if max_row < data_start_row:
        return 0
    return max_row


def highlight_style_fills(
    colors: OutputRowColors,
    color_map: dict[str, str],
) -> tuple[str, str]:
    """①強調3種セット（A〜E / F〜K）。"""
    case_fill = next(iter(color_map.values()), HIGHLIGHT_CASE_FILL_HEX)
    return case_fill, colors.highlight_driver


def default_style_fills(colors: OutputRowColors) -> tuple[str, str]:
    """②その他セット（A〜E / F〜K）。"""
    return colors.default_case, colors.default_driver


def resolve_padding_row_fills(
    excel_row: int,
    padding_start_row: int,
    colors: OutputRowColors,
    color_map: dict[str, str],
) -> tuple[str, str]:
    """書込範囲外の行: ①②を交互に返す（先頭行=①）。"""
    if (excel_row - padding_start_row) % 2 == 0:
        return highlight_style_fills(colors, color_map)
    return default_style_fills(colors)


def clear_output_row_values(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
    *,
    col_map: dict[str, int],
    clear_start_row: int,
    clear_end_row: int,
) -> str | None:
    """指定行のセル値を空にする（背景色は別途適用）。"""
    if clear_end_row < clear_start_row:
        return None
    clear_min = min(col_map.values())
    clear_max = max(col_map.values())
    row_count = clear_end_row - clear_start_row + 1
    width = clear_max - clear_min + 1
    clear_matrix = [[""] * width for _ in range(row_count)]
    clear_address = range_address(
        clear_min, clear_max, clear_start_row, clear_end_row
    )
    graph_patch_range_values(
        graph_token,
        item_id,
        sheet_name,
        clear_address,
        clear_matrix,
        session_id,
    )
    return clear_address


def apply_alternating_padding_row_colors(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
    *,
    col_map: dict[str, int],
    config: dict[str, Any],
    padding_start_row: int,
    padding_end_row: int,
) -> int:
    """書込範囲外の行に ①強調3種 / ②その他 の色を交互に塗る。"""
    if padding_end_row < padding_start_row:
        return 0
    colors = load_output_row_colors(config)
    color_map = load_case_name_row_color_map(config)
    case_min, case_max = case_block_col_range(col_map)
    driver_min, driver_max = driver_block_col_range(col_map)
    highlight_case, highlight_driver = highlight_style_fills(colors, color_map)
    default_case, default_driver = default_style_fills(colors)

    fills: list[tuple[str, str]] = []
    clear_addresses: list[str] = []
    for excel_row in range(padding_start_row, padding_end_row + 1):
        case_fill, driver_fill = resolve_padding_row_fills(
            excel_row, padding_start_row, colors, color_map
        )
        fills.append(
            (range_address(case_min, case_max, excel_row, excel_row), case_fill)
        )
        fills.append(
            (range_address(driver_min, driver_max, excel_row, excel_row), driver_fill)
        )
        for white_col_name in WHITE_FILL_COLUMN_NAMES:
            white_col = col_map.get(white_col_name)
            if white_col:
                clear_addresses.append(
                    range_address(white_col, white_col, excel_row, excel_row)
                )

    graph_batch_patch_fills(graph_token, item_id, sheet_name, fills, session_id)
    graph_batch_clear_fills(graph_token, item_id, sheet_name, clear_addresses, session_id)

    row_count = padding_end_row - padding_start_row + 1
    logging.info(
        "書込範囲外 %d 行に交互色適用（①強調 A〜E=%s F〜K=%s / ②その他 A〜E=%s F〜K=%s）",
        row_count,
        highlight_case,
        highlight_driver,
        default_case,
        default_driver,
    )
    return row_count


def graph_auth_headers(
    graph_token: str,
    session_id: str | None = None,
) -> dict[str, str]:
    headers = {
        "Authorization": f"Bearer {graph_token}",
        "Content-Type": "application/json",
    }
    if session_id:
        headers["workbook-session-id"] = session_id
    return headers


def graph_request_with_retry(
    method: str,
    url: str,
    graph_token: str,
    *,
    session_id: str | None = None,
    **kwargs: Any,
) -> requests.Response:
    max_attempts = 8
    last_res: requests.Response | None = None
    for attempt in range(1, max_attempts + 1):
        last_res = requests.request(
            method,
            url,
            headers=graph_auth_headers(graph_token, session_id),
            timeout=180,
            **kwargs,
        )
        if last_res.status_code == 423 and attempt < max_attempts:
            wait_sec = 5 * attempt
            logging.warning(
                "Excelがロック中（%d/%d）。%d秒後に再試行します。",
                attempt,
                max_attempts,
                wait_sec,
            )
            time.sleep(wait_sec)
            continue
        return last_res
    assert last_res is not None
    return last_res


def graph_get_drive_item(graph_token: str, remote_path: str) -> dict[str, Any]:
    url = f"{GRAPH_BASE}/me/drive/root:{remote_path}"
    res = graph_request_with_retry("GET", url, graph_token)
    if not res.ok:
        raise RuntimeError(
            f"OneDriveファイル情報取得失敗: {res.status_code} {res.text[:300]}"
        )
    return res.json()


def graph_create_workbook_session(graph_token: str, item_id: str) -> str:
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/createSession"
    res = graph_request_with_retry(
        "POST",
        url,
        graph_token,
        json={"persistChanges": True},
    )
    if not res.ok:
        raise RuntimeError(
            f"Excelセッション開始失敗: {res.status_code} {res.text[:300]}"
        )
    return res.json()["id"]


def graph_close_workbook_session(
    graph_token: str,
    item_id: str,
    session_id: str,
) -> None:
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/closeSession"
    res = graph_request_with_retry(
        "POST",
        url,
        graph_token,
        session_id=session_id,
        json={"persistChanges": True},
    )
    if not res.ok:
        logging.warning(
            "Excelセッション終了に失敗: %s %s", res.status_code, res.text[:200]
        )


def worksheet_segment(sheet_name: str) -> str:
    escaped = sheet_name.replace("'", "''")
    return f"worksheets('{escaped}')"


def graph_resolve_worksheet_name(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
) -> str:
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/worksheets"
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    if not res.ok:
        raise RuntimeError(
            f"ワークシート一覧取得失敗: {res.status_code} {res.text[:300]}"
        )
    sheets = res.json().get("value", [])
    target = sheet_name.strip()
    for sheet in sheets:
        if str(sheet.get("name", "")).strip() == target:
            return str(sheet["name"])
    if sheets:
        fallback = str(sheets[0]["name"])
        logging.warning(
            "シート '%s' が見つからないため '%s' を使用します",
            sheet_name,
            fallback,
        )
        return fallback
    raise RuntimeError(f"ワークシートが1つもありません: {sheet_name}")


def graph_read_range_values(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    address: str,
    session_id: str,
) -> list[list[Any]]:
    seg = worksheet_segment(sheet_name)
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"{seg}/range(address='{address}')"
    )
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    if not res.ok:
        raise RuntimeError(
            f"Excel範囲読取失敗({address}): {res.status_code} {res.text[:300]}"
        )
    return res.json().get("values", [])


def graph_patch_range_values(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    address: str,
    values: list[list[Any]],
    session_id: str,
) -> None:
    seg = worksheet_segment(sheet_name)
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"{seg}/range(address='{address}')"
    )
    res = graph_request_with_retry(
        "PATCH",
        url,
        graph_token,
        session_id=session_id,
        json={"values": values},
    )
    if not res.ok:
        raise RuntimeError(
            f"Excel範囲更新失敗({address}): {res.status_code} {res.text[:300]}"
        )


def graph_set_wrap_text(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    address: str,
    session_id: str,
) -> None:
    """セル内の改行を表示するため折り返し表示(wrapText)を有効化する。"""
    seg = worksheet_segment(sheet_name)
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"{seg}/range(address='{address}')/format"
    )
    res = graph_request_with_retry(
        "PATCH",
        url,
        graph_token,
        session_id=session_id,
        json={"wrapText": True},
    )
    if not res.ok:
        raise RuntimeError(
            f"折り返し表示設定失敗({address}): {res.status_code} {res.text[:300]}"
        )


SINGLE_LINE_ROW_HEIGHT = 21.0
ROW_HEIGHT_PADDING = 20.0
PER_EXTRA_LINE_HEIGHT = 20.5
ROW_FONT_SIZE = 13.0
ROW_LINE_COUNT_FIELDS = ("案件No", "案件名", "型式", "備考")
DEFAULT_COLUMN_WIDTH_PT = 200.0
# unicodedata.east_asian_width(): W/F=全角, H/Na/N=半角相当, A=曖昧(全角扱い)
NARROW_EAW_CATEGORIES = {"H", "Na", "N"}


def estimate_wrapped_line_count(text: str, column_width_pt: float) -> int:
    """折り返し表示時の見た目の行数を、全角/半角の文字幅から概算する。

    "リフト降ろしの為リン木必要です。" のような改行を含まない長文でも、
    列幅を超えれば自動的に折り返されるため、その分の高さも考慮する
    （半角カタカナなどを誤って全角幅と数えないよう east_asian_width で判定）。
    """
    if not text:
        return 1
    capacity = max(column_width_pt / ROW_FONT_SIZE, 1.0)
    total_units = sum(
        0.55 if unicodedata.east_asian_width(ch) in NARROW_EAW_CATEGORIES else 1.0
        for ch in text
    )
    if total_units <= 0:
        return 1
    return max(1, math.ceil(total_units / capacity))


def row_line_count(row: dict[str, Any], column_widths: dict[str, float]) -> int:
    """セル内改行＋折り返しを考慮した行の表示行数（最も行数の多い列に合わせる）。"""
    counts = []
    for field in ROW_LINE_COUNT_FIELDS:
        text = str(row.get(field, ""))
        width = column_widths.get(field, DEFAULT_COLUMN_WIDTH_PT)
        segments = text.split("\n") if text else [""]
        counts.append(
            sum(estimate_wrapped_line_count(seg, width) for seg in segments)
        )
    return max(counts, default=1)


def row_height_for(row: dict[str, Any], column_widths: dict[str, float]) -> float:
    """改行・折り返しが無い行にも少し余裕を持たせつつ、複数行ぶんの高さを確保する。"""
    lines = row_line_count(row, column_widths)
    return SINGLE_LINE_ROW_HEIGHT + ROW_HEIGHT_PADDING + (lines - 1) * PER_EXTRA_LINE_HEIGHT


def graph_get_column_widths(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
    field_to_letter: dict[str, str],
) -> dict[str, float]:
    """折り返し行数の推定に使う列幅(pt)を、実際のシートから取得する。"""
    seg = worksheet_segment(sheet_name)
    widths: dict[str, float] = {}
    for field, letter in field_to_letter.items():
        url = (
            f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
            f"{seg}/range(address='{letter}6:{letter}6')/format"
        )
        res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
        if res.ok:
            widths[field] = float(res.json().get("columnWidth") or DEFAULT_COLUMN_WIDTH_PT)
    return widths


def graph_batch_set_row_heights(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    heights: list[tuple[int, float]],
    session_id: str,
) -> None:
    """データ行ごとに行の高さを設定する（改行なしの行にも余白を確保）。"""
    if not heights:
        return
    seg = worksheet_segment(sheet_name)
    for start in range(0, len(heights), GRAPH_BATCH_CHUNK):
        chunk = heights[start:start + GRAPH_BATCH_CHUNK]
        requests_body = [
            {
                "id": str(i),
                "method": "PATCH",
                "url": (
                    f"/me/drive/items/{item_id}/workbook/"
                    f"{seg}/range(address='A{excel_row}')/format"
                ),
                "headers": {
                    "Content-Type": "application/json",
                    "workbook-session-id": session_id,
                },
                "body": {"rowHeight": height},
            }
            for i, (excel_row, height) in enumerate(chunk)
        ]
        _post_fill_batch_with_retry(graph_token, requests_body, session_id)


def parse_header_column_map(header_values: list[list[Any]]) -> dict[str, int]:
    header_row = header_values[0] if header_values else []
    mapping: dict[str, int] = {}
    for idx, val in enumerate(header_row, start=1):
        if val is None:
            continue
        name = str(val).strip()
        if name in EXCEL_OUTPUT_COLUMNS:
            mapping[name] = idx
    if not mapping:
        for idx, name in enumerate(EXCEL_OUTPUT_COLUMNS, start=1):
            mapping[name] = idx
    return mapping


def build_range_values(
    rows: list[dict[str, Any]],
    col_map: dict[str, int],
) -> tuple[list[list[Any]], int, int]:
    min_col = min(col_map.values())
    max_col = max(col_map.values())
    width = max_col - min_col + 1
    matrix: list[list[Any]] = []
    for row in rows:
        line = [""] * width
        for name, col_idx in col_map.items():
            line[col_idx - min_col] = cell_output_value(row, name)
        matrix.append(line)
    return matrix, min_col, max_col


def range_address(min_col: int, max_col: int, start_row: int, end_row: int) -> str:
    return (
        f"{col_letter_from_index(min_col)}{start_row}:"
        f"{col_letter_from_index(max_col)}{end_row}"
    )


def parse_a1_range(address: str) -> tuple[int, int, int, int]:
    """A1表記の範囲を (min_col, min_row, max_col, max_row) に分解（列は1始まり）。"""
    clean = address.split("!")[-1].replace("$", "")
    matched = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", clean.upper())
    if not matched:
        raise ValueError(f"範囲の解析に失敗: {address}")
    c1, r1, c2, r2 = matched.groups()
    return col_idx(c1) + 1, int(r1), col_idx(c2) + 1, int(r2)


def table_segment(table_name: str) -> str:
    escaped = table_name.replace("'", "''")
    return f"tables('{escaped}')"


def graph_list_tables_on_sheet(
    graph_token: str,
    item_id: str,
    ws_name: str,
    session_id: str,
) -> list[dict[str, Any]]:
    seg = worksheet_segment(ws_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/tables"
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    if not res.ok:
        raise RuntimeError(
            f"テーブル一覧取得失敗: {res.status_code} {res.text[:300]}"
        )
    return res.json().get("value", [])


def graph_get_table_range_address(
    graph_token: str,
    item_id: str,
    table_name: str,
    session_id: str,
) -> str:
    seg = table_segment(table_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/range"
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    if not res.ok:
        raise RuntimeError(
            f"テーブル範囲取得失敗: {res.status_code} {res.text[:300]}"
        )
    return str(res.json().get("address", ""))


def graph_get_table_column_map(
    graph_token: str,
    item_id: str,
    table_name: str,
    session_id: str,
    table_min_col: int,
) -> dict[str, int]:
    seg = table_segment(table_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/columns"
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    if not res.ok:
        raise RuntimeError(
            f"テーブル列取得失敗: {res.status_code} {res.text[:300]}"
        )
    mapping: dict[str, int] = {}
    for idx, col in enumerate(res.json().get("value", [])):
        name = str(col.get("name", "")).strip()
        if name in EXCEL_OUTPUT_COLUMNS:
            mapping[name] = table_min_col + idx
    return mapping


def graph_resize_table(
    graph_token: str,
    item_id: str,
    table_name: str,
    address: str,
    session_id: str,
) -> None:
    seg = table_segment(table_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/resize"
    res = graph_request_with_retry(
        "POST",
        url,
        graph_token,
        session_id=session_id,
        json={"address": address.split("!")[-1]},
    )
    if not res.ok:
        raise RuntimeError(
            f"テーブルリサイズ失敗: {res.status_code} {res.text[:300]}"
        )


def graph_get_table_column_names(
    graph_token: str,
    item_id: str,
    table_name: str,
    session_id: str,
) -> list[str]:
    seg = table_segment(table_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/columns"
    res = graph_request_with_retry("GET", url, graph_token, session_id=session_id)
    if not res.ok:
        raise RuntimeError(
            f"テーブル列取得失敗({table_name}): {res.status_code} {res.text[:200]}"
        )
    return [
        str(col.get("name", "")).strip()
        for col in res.json().get("value", [])
    ]


def pick_target_table(
    tables: list[dict[str, Any]],
    preferred_name: str,
    graph_token: str,
    item_id: str,
    session_id: str,
) -> dict[str, Any] | None:
    if not tables:
        return None
    if preferred_name:
        for table in tables:
            if str(table.get("name", "")).strip() == preferred_name.strip():
                return table

    scored: list[tuple[int, dict[str, Any], list[str]]] = []
    for table in tables:
        name = str(table.get("name", ""))
        try:
            columns = graph_get_table_column_names(
                graph_token, item_id, name, session_id
            )
        except RuntimeError as exc:
            logging.warning("テーブル %s の列取得スキップ: %s", name, exc)
            continue
        score = sum(1 for col in EXCEL_OUTPUT_COLUMNS if col in columns)
        if "案件No" in columns:
            score += 20
        scored.append((score, table, columns))

    if scored:
        scored.sort(key=lambda item: item[0], reverse=True)
        best_score, best_table, best_cols = scored[0]
        if len(scored) > 1:
            logging.info(
                "テーブル選択: %s（一致列 %d / %s）",
                best_table.get("name"),
                best_score,
                ", ".join(best_cols[:6]),
            )
        return best_table

    if len(tables) == 1:
        return tables[0]
    logging.warning(
        "ドライバー用テーブルを特定できません。先頭を使用: %s",
        ", ".join(str(t.get("name", "")) for t in tables),
    )
    return tables[0]


def resolve_table_write_target(
    graph_token: str,
    item_id: str,
    ws_name: str,
    session_id: str,
    config: dict[str, Any],
) -> tuple[str | None, int, int, int, int, dict[str, int]]:
    od = config.get("onedrive_output", {})
    tables = graph_list_tables_on_sheet(graph_token, item_id, ws_name, session_id)
    table = pick_target_table(
        tables,
        str(od.get("table_name", "")).strip(),
        graph_token,
        item_id,
        session_id,
    )

    table_header_row = int(od.get("table_header_row", 6))
    data_start_row = int(od.get("data_start_row", 7))
    if data_start_row <= table_header_row:
        raise RuntimeError(
            f"data_start_row({data_start_row}) は table_header_row({table_header_row}) より大きくしてください"
        )

    col_map = load_output_column_map(config)
    min_col = min(col_map.values())
    max_col = max(col_map.values())

    if table is None:
        logging.warning(
            "テーブルなし: 固定範囲 A%d:K に書き込みます", data_start_row
        )
        return None, min_col, max_col, table_header_row, data_start_row, col_map

    table_name = str(table["name"])
    return table_name, min_col, max_col, table_header_row, data_start_row, col_map


def update_onedrive_values_only(
    graph_token: str,
    config: dict[str, Any],
    rows: list[dict[str, Any]],
) -> int:
    """Graph Excel API で着日範囲内の件数だけ A7:K* に書込み、余剰行を自動削除。"""
    od = config.get("onedrive_output", {})
    remote_path = od.get("path", "/ドライバー情報/ドライバー情報_営業用.xlsx")
    sheet_name = od.get("sheet_name", "ドライバー情報")

    item = graph_get_drive_item(graph_token, remote_path)
    item_id = item["id"]
    session_id = graph_create_workbook_session(graph_token, item_id)
    try:
        ws_name = graph_resolve_worksheet_name(
            graph_token, item_id, sheet_name, session_id
        )

        data_start_row = int(od.get("data_start_row", 7))
        table_header_row = int(od.get("table_header_row", 6))
        col_map = load_output_column_map(config)

        sheet_password = load_sheet_password()
        if sheet_password:
            graph_unprotect_worksheet(
                graph_token, item_id, ws_name, session_id, sheet_password
            )

        row_count = len(rows)
        data_end_row = output_data_end_row(data_start_row, row_count)
        previous_end_row = load_last_patched_end_row(data_start_row)
        sheet_used_end_row = load_last_sheet_used_end_row(data_start_row)

        logging.info(
            "書式維持モード: シート=%s, 見出し行=%d(触らない), データ=%d行 → 行%d〜%d",
            ws_name,
            table_header_row,
            row_count,
            data_start_row if row_count else 0,
            data_end_row if row_count else 0,
        )

        logging.info(
            "列マッピング: %s",
            ", ".join(f"{k}={col_letter_from_index(v)}" for k, v in col_map.items()),
        )

        api_trace: dict[str, Any] = {
            "item_id": item_id,
            "worksheet": ws_name,
            "session": "persistent (workbook session)",
            "row_count": row_count,
            "data_start_row": data_start_row,
            "data_end_row": data_end_row,
            "previous_end_row": previous_end_row,
            "sheet_used_end_row": sheet_used_end_row,
            "steps": [],
        }
        get_item_url = f"{GRAPH_BASE}/me/drive/root:{remote_path}"
        api_trace["steps"].append({"method": "GET", "url": get_item_url, "purpose": "ファイルID取得"})
        api_trace["steps"].append(
            {
                "method": "POST",
                "url": f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/createSession",
                "purpose": "編集セッション開始",
            }
        )

        if rows:
            matrix, write_min_col, write_max_col = build_range_values(rows, col_map)
            data_address = range_address(
                write_min_col, write_max_col, data_start_row, data_end_row
            )
            patch_url = (
                f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/worksheets('{ws_name}')"
                f"/range(address='{data_address}')"
            )
            api_trace["steps"].append(
                {
                    "method": "PATCH",
                    "url": patch_url,
                    "range": data_address,
                    "row_count": row_count,
                    "sample_first_row": matrix[0] if matrix else [],
                    "sample_last_row": matrix[-1] if matrix else [],
                }
            )
            graph_patch_range_values(
                graph_token,
                item_id,
                ws_name,
                data_address,
                matrix,
                session_id,
            )
            logging.info(
                "Graph API PATCH: %s に %d 行（着日範囲内・例: 1行目=%s）",
                data_address,
                row_count,
                matrix[0] if matrix else [],
            )
            graph_set_wrap_text(graph_token, item_id, ws_name, data_address, session_id)
            field_to_letter = {
                field: col_letter_from_index(col_map[field])
                for field in ROW_LINE_COUNT_FIELDS
                if field in col_map
            }
            column_widths = graph_get_column_widths(
                graph_token, item_id, ws_name, session_id, field_to_letter
            )
            row_heights = [
                (data_start_row + idx, row_height_for(row, column_widths))
                for idx, row in enumerate(rows)
            ]
            graph_batch_set_row_heights(
                graph_token, item_id, ws_name, row_heights, session_id
            )
            apply_driver_change_detection(graph_token, item_id, session_id, rows)
            apply_output_row_colors(
                graph_token,
                item_id,
                ws_name,
                session_id,
                rows,
                data_start_row=data_start_row,
                col_map=col_map,
                config=config,
            )
            api_trace["steps"].append(
                {
                    "method": "PATCH",
                    "url": f".../format/fill (行色 {row_count} 行)",
                    "purpose": "案件名グループ別の行背景色",
                }
            )
        else:
            logging.info("着日範囲内の案件なし — データ行は書き込みません")

        used_last_row = graph_get_worksheet_used_last_row(
            graph_token,
            item_id,
            ws_name,
            session_id,
            data_start_row=data_start_row,
        )
        padding_start_row = data_end_row + 1
        padding_end_row = resolve_padding_end_row(
            config,
            data_end_row=data_end_row,
            sheet_used_end_row=sheet_used_end_row,
            previous_end_row=previous_end_row,
            used_last_row=used_last_row,
        )
        if padding_end_row >= padding_start_row and col_map:
            clear_address = clear_output_row_values(
                graph_token,
                item_id,
                ws_name,
                session_id,
                col_map=col_map,
                clear_start_row=padding_start_row,
                clear_end_row=padding_end_row,
            )
            padded = apply_alternating_padding_row_colors(
                graph_token,
                item_id,
                ws_name,
                session_id,
                col_map=col_map,
                config=config,
                padding_start_row=padding_start_row,
                padding_end_row=padding_end_row,
            )
            logging.info(
                "Graph API: 書込範囲外 %d 行を整理（値クリア %s）",
                padded,
                clear_address,
            )
            api_trace["steps"].append(
                {
                    "method": "PATCH",
                    "url": f".../range + format/fill (書込範囲外 {padded} 行・①②交互)",
                    "range": clear_address,
                    "purpose": "範囲外行の値削除と交互色",
                }
            )
            sheet_used_end_row = padding_end_row
        else:
            logging.info(
                "書込範囲外なし（今回最終行=%d, 前回最終行=%d, シート使用最終行=%d, usedRange=%d）",
                data_end_row,
                previous_end_row,
                sheet_used_end_row,
                used_last_row,
            )
            sheet_used_end_row = max(sheet_used_end_row, data_end_row)

        api_trace["data_end_row"] = data_end_row
        api_trace["sheet_used_end_row"] = sheet_used_end_row

        api_trace["steps"].append(
            {
                "method": "POST",
                "url": f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/closeSession",
                "purpose": "セッション終了",
            }
        )
        API_DEBUG_PATH.write_text(
            json.dumps(api_trace, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        logging.info("API詳細を保存: %s", API_DEBUG_PATH.name)

        if sheet_password:
            for hidden_col_name in PROTECTED_HIDDEN_COLUMN_NAMES:
                hidden_col = col_map.get(hidden_col_name)
                if hidden_col:
                    graph_set_column_width(
                        graph_token, item_id, ws_name,
                        col_letter_from_index(hidden_col), 0, session_id,
                    )
            graph_protect_worksheet(
                graph_token, item_id, ws_name, session_id, sheet_password
            )
            logging.info(
                "シート保護を再適用（非表示列: %s）",
                ", ".join(PROTECTED_HIDDEN_COLUMN_NAMES),
            )

        return data_end_row, sheet_used_end_row
    finally:
        graph_close_workbook_session(graph_token, item_id, session_id)


def graph_item_url(remote_path: str, action: str) -> str:
    return f"{GRAPH_BASE}/me/drive/root:{remote_path}:{action}"


def upload_onedrive_excel(graph_token: str, remote_path: str, content: bytes) -> None:
    url = graph_item_url(remote_path, "/content")
    headers = {
        "Authorization": f"Bearer {graph_token}",
        "Content-Type": (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    }
    max_attempts = 8
    for attempt in range(1, max_attempts + 1):
        res = requests.put(url, headers=headers, data=content, timeout=180)
        if res.ok:
            return
        if res.status_code == 423 and attempt < max_attempts:
            wait_sec = 5 * attempt
            logging.warning(
                "OneDriveファイルがロック中（%d/%d）。%d秒後に再試行します。"
                "ブラウザで営業用Excelを開いている場合はタブを閉じてください。",
                attempt,
                max_attempts,
                wait_sec,
            )
            time.sleep(wait_sec)
            continue
        if res.status_code == 423:
            raise RuntimeError(
                "OneDriveアップロード失敗: ファイルがロックされています。"
                "「ドライバー情報_営業用.xlsx」を Excel Online で開いているタブを"
                "すべて閉じてから、もう一度 python driver_sync.py を実行してください。"
            )
        raise RuntimeError(f"OneDriveアップロード失敗: {res.status_code} {res.text[:300]}")


def create_onedrive_view_link(
    graph_token: str,
    remote_path: str,
    scope: str,
) -> str:
    url = graph_item_url(remote_path, "/createLink")
    res = requests.post(
        url,
        headers={
            "Authorization": f"Bearer {graph_token}",
            "Content-Type": "application/json",
        },
        json={"type": "view", "scope": scope},
        timeout=60,
    )
    if not res.ok:
        raise RuntimeError(f"閲覧リンク作成失敗: {res.status_code} {res.text[:300]}")
    link = res.json().get("link", {})
    return link.get("webUrl") or link.get("shareUrl") or ""


def load_cached_view_link(remote_path: str) -> str | None:
    if not LINK_CACHE_PATH.exists():
        return None
    try:
        data = json.loads(LINK_CACHE_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None
    if data.get("path") == remote_path and data.get("url"):
        return data["url"]
    return None


def save_view_link(remote_path: str, url: str) -> None:
    LINK_CACHE_PATH.write_text(
        json.dumps({"path": remote_path, "url": url}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    VIEW_LINK_PATH.write_text(
        "営業向け閲覧リンク（ログイン不要・閲覧のみ）\n"
        f"{url}\n",
        encoding="utf-8",
    )


class WarningCountHandler(logging.Handler):
    """抽出・検証フェーズで出たWARNINGを集計（E2〜E4の要確認アラート用）。"""

    def __init__(self) -> None:
        super().__init__(level=logging.WARNING)
        self.count = 0
        self.messages: list[str] = []

    def emit(self, record: logging.LogRecord) -> None:
        self.count += 1
        self.messages.append(record.getMessage())


ISSUE_CATEGORY_KEYWORDS = [
    ("ドライバー4項目未入力", "ドライバー4項目未入力"),
    ("行整合性NG", "行整合性NG"),
    ("整合性NG", "出力整合性NG"),
    ("シートが見つかりません", "シート不明"),
]
ISSUE_CASE_NO_PATTERN = re.compile(r"\b[A-Z][A-Z0-9]{3,9}-\d{1,2}\b")
ISSUE_CASE_AFTER_LABEL_PATTERN = re.compile(r"案件=([^\s]+)")
ISSUE_SHIP_DATE_PATTERN = re.compile(r"出荷日=([\d/]+)")
ISSUE_SOURCE_KEY_PATTERN = re.compile(r"\b(matsuzaki|nakadori|fukuoka|maruun)\b")
ISSUE_SOURCE_ABBREV = {
    "matsuzaki": "松崎",
    "nakadori": "中通",
    "fukuoka": "福岡ロジ",
    "maruun": "丸運",
}
E3_MAX_LEN = 40  # セルに収まる目安の文字数。超えたらE4に続きを出す。


def extract_source_abbrev(msg: str) -> str:
    """警告メッセージから依頼先(略称)を取り出す。不明なら空文字。"""
    match = ISSUE_SOURCE_KEY_PATTERN.search(msg)
    if not match:
        return ""
    return ISSUE_SOURCE_ABBREV.get(match.group(1), "")


def extract_case_no(msg: str) -> str | None:
    """「案件=」以降の値を最優先で取り出し、なければ案件No風パターンで補完。"""
    match = ISSUE_CASE_AFTER_LABEL_PATTERN.search(msg)
    if match:
        return match.group(1)
    match = ISSUE_CASE_NO_PATTERN.search(msg)
    if match:
        return match.group(0)
    return None


def extract_issue_entry(msg: str) -> str | None:
    """「依頼先+出荷日 / 案件No」形式のエントリを作る（例: 中通6.29 / W25A392-01）。"""
    case_no = extract_case_no(msg)
    if not case_no:
        return None
    abbrev = extract_source_abbrev(msg)
    date_match = ISSUE_SHIP_DATE_PATTERN.search(msg)
    if date_match:
        ship_md = date_match.group(1).replace("/", ".")
        return f"{abbrev}{ship_md} / {case_no}"
    return f"{abbrev}{case_no}" if abbrev else case_no


IGNORE_LIST_RANGE = "G3:H4"  # G3〜H4の4セル。1セルに1案件Noを入力する。


def graph_get_ignored_case_numbers(
    graph_token: str, config: dict[str, Any]
) -> set[str]:
    """F2:H3（確認済み・無視リスト）の6セルに入力された案件Noを取得する。

    営業さんがOK判断した案件Noを1セルずつ入力しておくと、
    その案件は次回以降「要確認」アラートから除外される（読むだけで上書きしない）。
    """
    od = config.get("onedrive_output", {})
    remote_path = od.get("path", "/ドライバー情報/ドライバー情報_営業用.xlsx")
    sheet_name = od.get("sheet_name", "ドライバー情報")
    item = graph_get_drive_item(graph_token, remote_path)
    item_id = item["id"]
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"worksheets('{sheet_name}')/range(address='{IGNORE_LIST_RANGE}')"
    )
    res = requests.get(url, headers={"Authorization": f"Bearer {graph_token}"})
    if not res.ok:
        return set()
    values = res.json().get("values") or []
    result: set[str] = set()
    for row in values:
        for cell in row:
            if isinstance(cell, float) and cell.is_integer():
                text = str(int(cell))
            else:
                text = str(cell).strip()
            if text:
                result.add(text)
    return result


def categorize_issue_messages(
    messages: list[str], ignored_case_numbers: set[str] | None = None
) -> tuple[list[str], list[str]]:
    """WARNINGメッセージから (種類別件数の行, 「出荷日 / 案件No」一覧) を作る。
    確認済み・無視リストに入っている案件Noは集計から除外する。"""
    ignored = ignored_case_numbers or set()
    category_counts: dict[str, int] = {}
    entries: list[str] = []
    for msg in messages:
        case_no = extract_case_no(msg)
        if case_no and case_no in ignored:
            continue
        label = next(
            (lbl for kw, lbl in ISSUE_CATEGORY_KEYWORDS if kw in msg),
            "その他",
        )
        category_counts[label] = category_counts.get(label, 0) + 1
        entry = extract_issue_entry(msg)
        if entry and entry not in entries:
            entries.append(entry)
    summary_lines = [f"{label}:{count}件" for label, count in category_counts.items()]
    return summary_lines, entries


def split_entries_for_cells(entries: list[str], max_len: int) -> tuple[str, str]:
    """エントリ一覧を「、」区切りで詰め、最初のセルに収まらない分を2つ目に回す。"""
    primary: list[str] = []
    overflow: list[str] = []
    current_len = 0
    for entry in entries:
        added_len = len(entry) + (1 if primary else 0)
        if not overflow and current_len + added_len <= max_len:
            primary.append(entry)
            current_len += added_len
        else:
            overflow.append(entry)
    return "、".join(primary), "、".join(overflow)


def write_alert_label(
    graph_token: str,
    config: dict[str, Any],
    issue_count: int,
    issue_messages: list[str] | None = None,
) -> None:
    """E2に種類別件数、E3〜E4に「出荷日 / 案件No」を表示（問題なければ空欄）。"""
    od = config.get("onedrive_output", {})
    remote_path = od.get("path", "/ドライバー情報/ドライバー情報_営業用.xlsx")
    sheet_name = od.get("sheet_name", "ドライバー情報")
    item = graph_get_drive_item(graph_token, remote_path)
    item_id = item["id"]

    ignored_case_numbers = graph_get_ignored_case_numbers(graph_token, config)
    summary_lines, entries = categorize_issue_messages(
        issue_messages or [], ignored_case_numbers
    )
    e2_text = " / ".join(summary_lines)
    e3_text, e4_text = split_entries_for_cells(entries, E3_MAX_LEN)

    headers = {
        "Authorization": f"Bearer {graph_token}",
        "Content-Type": "application/json",
    }
    seg = f"worksheets('{sheet_name}')"
    base = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}"
    requests.patch(
        f"{base}/range(address='E2:E4')",
        headers=headers,
        json={"values": [[e2_text], [e3_text], [e4_text]]},
    )


def write_last_synced_label(graph_token: str, config: dict[str, Any]) -> None:
    """E1セルに最終更新時刻を表示（コメントBOX横の更新状況表示）。"""
    od = config.get("onedrive_output", {})
    remote_path = od.get("path", "/ドライバー情報/ドライバー情報_営業用.xlsx")
    sheet_name = od.get("sheet_name", "ドライバー情報")
    item = graph_get_drive_item(graph_token, remote_path)
    item_id = item["id"]
    timestamp = datetime.now(JST).strftime("%Y/%m/%d %H:%M")
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"worksheets('{sheet_name}')/range(address='E1')"
    )
    headers = {
        "Authorization": f"Bearer {graph_token}",
        "Content-Type": "application/json",
    }
    requests.patch(
        url, headers=headers, json={"values": [[f"最終更新: {timestamp}"]]}
    )


def publish_onedrive_for_sales(
    graph_token: str,
    config: dict[str, Any],
    rows: list[dict[str, Any]],
    issue_count: int = 0,
    issue_messages: list[str] | None = None,
) -> tuple[str, int, int]:
    od = config.get("onedrive_output", {})
    remote_path = od.get("path", "/ドライバー情報/ドライバー情報_営業用.xlsx")
    share_scope = od.get("share_scope", "anonymous")
    data_start_row = int(od.get("data_start_row", 7))
    data_end_row = output_data_end_row(data_start_row, len(rows))
    sheet_used_end_row = data_end_row

    update_mode = od.get("update_mode", "preserve_format")
    if update_mode == "preserve_format":
        data_end_row, sheet_used_end_row = update_onedrive_values_only(
            graph_token, config, rows
        )
        if rows:
            logging.info(
                "OneDrive Excel データ更新: %s (%d 行・A%d:K%d・書式維持)",
                remote_path,
                len(rows),
                data_start_row,
                data_end_row,
            )
        else:
            logging.info(
                "OneDrive Excel データ更新: %s (着日範囲内 0 件・書式維持)",
                remote_path,
            )
        try:
            write_last_synced_label(graph_token, config)
            write_alert_label(graph_token, config, issue_count, issue_messages)
        except requests.RequestException as exc:
            logging.warning("最終更新時刻/要確認アラートの書込に失敗: %s", exc)
    else:
        content = rows_to_xlsx_bytes(rows, config)
        upload_onedrive_excel(graph_token, remote_path, content)
        logging.info("OneDrive Excel 全置換: %s (%d 行)", remote_path, len(rows))

    view_url = load_cached_view_link(remote_path)
    if not view_url:
        try:
            view_url = create_onedrive_view_link(graph_token, remote_path, share_scope)
        except RuntimeError:
            if share_scope == "anonymous":
                logging.warning("匿名リンク作成失敗。organization スコープで再試行します。")
                view_url = create_onedrive_view_link(
                    graph_token, remote_path, "organization"
                )
            else:
                raise
        save_view_link(remote_path, view_url)
        logging.info("閲覧リンクを新規作成しました")

    logging.info("営業用閲覧リンク: %s", view_url)
    logging.info("リンクは %s にも保存しました", VIEW_LINK_PATH.name)
    return view_url, data_end_row, sheet_used_end_row


def sync_sharepoint_list(
    config: dict[str, Any],
    all_rows: list[dict[str, Any]],
    force_login: bool,
    use_device_code: bool,
) -> None:
    sp_token = acquire_token(
        config,
        scopes=sharepoint_scopes(config),
        force_login=force_login,
        use_device_code=use_device_code,
    )
    sp_client = SharePointRestClient(
        api_base=f"https://{config['sharepoint_hostname']}{config['sharepoint_site_path']}/_api",
        access_token=sp_token,
    )
    list_title = config["list_display_name"]
    entity_map = sp_client.get_entity_property_map(list_title)
    item_entity_type = sp_client.get_list_item_entity_type(list_title)
    digest = sp_client.get_digest()

    existing_ids = sp_client.list_all_item_ids(list_title)
    logging.info("既存項目を削除: %d 件", len(existing_ids))
    for item_id in existing_ids:
        sp_client.delete_item(list_title, item_id, digest)

    logging.info("新規項目を作成: %d 件", len(all_rows))
    for row in all_rows:
        fields = row_to_sp_fields(row, entity_map)
        sp_client.create_item(list_title, fields, digest, item_entity_type)


@dataclass
class SharePointRestClient:
    api_base: str
    access_token: str

    @property
    def headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Accept": "application/json;odata=verbose",
        }

    def get_digest(self) -> str:
        res = requests.post(
            f"{self.api_base}/contextinfo",
            headers=self.headers,
            timeout=60,
        )
        if not res.ok:
            raise RuntimeError(f"FormDigest取得失敗: {res.status_code} {res.text[:300]}")
        return res.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    def get_list_item_entity_type(self, list_title: str) -> str:
        safe_title = list_title.replace("'", "''")
        url = (
            f"{self.api_base}/web/lists/getbytitle('{safe_title}')"
            "?$select=ListItemEntityTypeFullName"
        )
        res = requests.get(url, headers=self.headers, timeout=60)
        if not res.ok:
            raise RuntimeError(f"リスト型取得失敗: {res.status_code} {res.text[:300]}")
        return res.json()["d"]["ListItemEntityTypeFullName"]

    def get_entity_property_map(self, list_title: str) -> dict[str, str]:
        safe_title = list_title.replace("'", "''")
        url = (
            f"{self.api_base}/web/lists/getbytitle('{safe_title}')/fields"
            "?$select=Title,EntityPropertyName"
        )
        res = requests.get(url, headers=self.headers, timeout=60)
        if not res.ok:
            raise RuntimeError(f"列定義取得失敗: {res.status_code} {res.text[:300]}")
        mapping: dict[str, str] = {"Title": "Title"}
        for field in res.json()["d"]["results"]:
            title = field.get("Title")
            entity = field.get("EntityPropertyName")
            if title and entity:
                mapping[title] = entity
        return mapping

    def list_all_item_ids(self, list_title: str) -> list[int]:
        safe_title = list_title.replace("'", "''")
        url = f"{self.api_base}/web/lists/getbytitle('{safe_title}')/items?$select=Id"
        ids: list[int] = []
        while url:
            res = requests.get(url, headers=self.headers, timeout=60)
            if not res.ok:
                raise RuntimeError(f"項目取得失敗: {res.status_code} {res.text[:300]}")
            data = res.json()["d"]
            ids.extend(int(item["Id"]) for item in data["results"])
            url = data.get("__next")
        return ids

    def delete_item(self, list_title: str, item_id: int, digest: str) -> None:
        safe_title = list_title.replace("'", "''")
        url = f"{self.api_base}/web/lists/getbytitle('{safe_title}')/items({item_id})"
        headers = {
            **self.headers,
            "X-RequestDigest": digest,
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
        }
        res = requests.post(url, headers=headers, timeout=60)
        if res.status_code not in (200, 204):
            raise RuntimeError(f"項目削除失敗: {res.status_code} {res.text[:300]}")

    def create_item(
        self,
        list_title: str,
        fields: dict[str, Any],
        digest: str,
        item_entity_type: str,
    ) -> None:
        safe_title = list_title.replace("'", "''")
        url = f"{self.api_base}/web/lists/getbytitle('{safe_title}')/items"
        headers = {
            **self.headers,
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": digest,
        }
        body = {"__metadata": {"type": item_entity_type}, **fields}
        res = requests.post(url, headers=headers, json=body, timeout=60)
        if not res.ok:
            raise RuntimeError(f"項目作成失敗: {res.status_code} {res.text[:300]}")


def load_sheet_password() -> str:
    """シート保護用パスワードを環境変数→ローカルファイルの順で読む。"""
    env_value = os.environ.get(SHEET_PASSWORD_ENV_VAR, "").strip()
    if env_value:
        return env_value
    if SHEET_PASSWORD_PATH.exists():
        return SHEET_PASSWORD_PATH.read_text(encoding="utf-8").strip()
    return ""


def graph_set_column_width(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    column_letter: str,
    width: float,
    session_id: str,
) -> None:
    seg = worksheet_segment(sheet_name)
    url = (
        f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/"
        f"{seg}/range(address='{column_letter}:{column_letter}')/format"
    )
    graph_request_with_retry(
        "PATCH", url, graph_token, session_id=session_id,
        json={"columnWidth": width},
    )


def graph_unprotect_worksheet(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
    password: str,
) -> None:
    seg = worksheet_segment(sheet_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/protection/unprotect"
    res = graph_request_with_retry(
        "POST", url, graph_token, session_id=session_id,
        json={"password": password} if password else {},
    )
    if not res.ok and res.status_code != 400:
        raise RuntimeError(
            f"シート保護解除失敗: {res.status_code} {res.text[:200]}"
        )


def graph_protect_worksheet(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    session_id: str,
    password: str,
) -> None:
    seg = worksheet_segment(sheet_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/protection/protect"
    body: dict[str, Any] = {"options": {"allowFormatColumns": False}}
    if password:
        body["password"] = password
    res = graph_request_with_retry(
        "POST", url, graph_token, session_id=session_id, json=body,
    )
    if not res.ok:
        raise RuntimeError(f"シート保護失敗: {res.status_code} {res.text[:200]}")


def build_token_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
    return cache


def save_token_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")


def acquire_token(
    config: dict[str, Any],
    scopes: list[str],
    force_login: bool = False,
    use_device_code: bool = False,
) -> str:
    cache = build_token_cache()
    app = msal.PublicClientApplication(
        config["client_id"],
        authority=f"https://login.microsoftonline.com/{config['tenant_id']}",
        token_cache=cache,
    )
    if not force_login:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                save_token_cache(cache)
                return result["access_token"]
        if is_unattended():
            raise RuntimeError(
                "無人実行: 認証トークンが無効または期限切れです。"
                "次回出社時に python driver_sync.py --login を1回実行してください。"
            )

    if use_device_code:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(
                f"デバイスコードフロー開始失敗: {json.dumps(flow, ensure_ascii=False)}"
            )
        print(flow["message"])
        print("※ コード入力が完了するまで、この画面は閉じないでください（Ctrl+C不要）")
        result = app.acquire_token_by_device_flow(flow)
    else:
        login_hint = config.get("login_hint", "t_nakanishi@showa.co.jp")
        print(f"ブラウザが開きます。{login_hint} でサインインしてください。")
        print("※ サインイン後は自動でこの画面に戻ります（ポートは自動割当）。")
        print("権限の同意画面が出たら「同意」を押してください。")
        result = app.acquire_token_interactive(
            scopes=scopes,
            login_hint=login_hint,
            prompt="select_account" if force_login else None,
        )

    save_token_cache(cache)

    if "access_token" not in result:
        error = result.get("error_description") or result.get("error") or result
        raise RuntimeError(
            "認証失敗: "
            f"{error}\n"
            "→ IT管理者に「SharePoint の AllSites.Write（委任）」または"
            "「Microsoft Graph の Sites.ReadWrite.All（委任）」の追加を依頼してください。"
        )

    logging.info("認証成功。token_cache.bin を保存しました。")
    return result["access_token"]


def row_to_sp_fields(
    row: dict[str, Any],
    entity_map: dict[str, str],
) -> dict[str, Any]:
    fields: dict[str, Any] = {}
    for src_key, display_name in OUTPUT_FIELD_MAP.items():
        entity_name = entity_map.get(display_name)
        if not entity_name:
            continue
        value = row.get(src_key)
        fields[entity_name] = "" if value is None else str(value)
    return fields


def run_sync(
    dry_run: bool = False,
    force_login: bool = False,
    use_device_code: bool = False,
    validate_headers_only: bool = False,
    debug_dump: bool = False,
) -> int:
    try:
        acquire_process_lock()
    except FileExistsError:
        logging.warning("前回の同期がまだ実行中のためスキップしました")
        write_status(ok=True, message="skipped: already running")
        return 0

    config = load_config()
    today = date.today()
    arr_days_back = int(config.get("arr_days_back", 7))
    header_rows = int(config.get("header_rows", 4))
    auto_detect_columns = bool(config.get("auto_detect_columns", False))
    arr_window_start_date = arr_window_start(today, arr_days_back)
    arr_window_end_date = arr_window_end(today)

    logging.info(
        "同期開始: today=%s, 着日範囲=%s〜%s（%d日前〜明日・暦日）",
        today,
        arr_window_start_date,
        arr_window_end_date,
        arr_days_back,
    )

    warning_counter = WarningCountHandler()
    logging.getLogger().addHandler(warning_counter)

    all_rows: list[dict[str, Any]] = []
    for source in config["sources"]:
        key = source["key"]
        share_url = source["share_url"]
        logging.info("取得中: %s", key)
        content = download_share_file(share_url)
        workbook = openpyxl.load_workbook(
            BytesIO(content),
            read_only=False,
            data_only=True,
        )
        rows = extract_rows_from_workbook(
            workbook,
            key,
            header_rows,
            today,
            arr_days_back=arr_days_back,
            auto_detect_columns=auto_detect_columns,
        )
        workbook.close()
        logging.info("抽出件数: %s = %d", key, len(rows))
        all_rows.extend(rows)

    all_rows = prepare_rows_for_output(all_rows, today)
    validate_row_quality(all_rows)
    validate_output_rows_integrity(all_rows)
    logging.getLogger().removeHandler(warning_counter)
    issue_count = warning_counter.count
    issue_messages = warning_counter.messages
    logging.info("合計抽出件数: %d（要確認警告: %d件）", len(all_rows), issue_count)

    data_start_row = int(
        config.get("onedrive_output", {}).get("data_start_row", 7)
    )
    data_end_row = output_data_end_row(data_start_row, len(all_rows))
    if all_rows:
        logging.info(
            "営業用Excel書込予定: 着日範囲内 %d 件 → A%d:K%d",
            len(all_rows),
            data_start_row,
            data_end_row,
        )
    else:
        logging.info(
            "営業用Excel: 着日範囲内 0 件 — 既存データ行をクリアします（開始行=%d）",
            data_start_row,
        )

    dump_payload = {
        "today": str(today),
        "arr_days_back": arr_days_back,
        "arr_window_start": str(arr_window_start_date),
        "arr_window_end": str(arr_window_end_date),
        "row_count": len(all_rows),
        "rows": all_rows,
    }
    DEBUG_OUTPUT_PATH.write_text(
        json.dumps(dump_payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    if debug_dump or validate_headers_only or dry_run:
        logging.info("抽出JSONを保存: %s", DEBUG_OUTPUT_PATH.name)
        for sample in all_rows[:8]:
            logging.info("行サンプル: %s", json.dumps(sample, ensure_ascii=False))

    if validate_headers_only:
        logging.info("ヘッダー検証のみ完了（アップロードはしません）")
        for sample in all_rows[:5]:
            logging.info("サンプル: %s", json.dumps(sample, ensure_ascii=False))
        write_status(ok=True, message="validate-headers", row_count=len(all_rows))
        return 0

    if dry_run:
        logging.info("dry-run のため OneDrive / SharePoint 反映はスキップしました")
        for sample in all_rows[:3]:
            logging.info("サンプル: %s", json.dumps(sample, ensure_ascii=False))
        write_status(ok=True, message="dry-run", row_count=len(all_rows))
        return 0

    od = config.get("onedrive_output", {})
    sheet_used_end_row = data_end_row
    if od.get("enabled", True):
        graph_token = acquire_token(
            config,
            scopes=GRAPH_SCOPES,
            force_login=force_login,
            use_device_code=use_device_code,
        )
        _, data_end_row, sheet_used_end_row = publish_onedrive_for_sales(
            graph_token, config, all_rows, issue_count, issue_messages
        )

    if config.get("sync_sharepoint", False):
        sync_sharepoint_list(config, all_rows, force_login, use_device_code)

    logging.info("同期完了")
    write_status(
        ok=True,
        message="sync-complete",
        row_count=len(all_rows),
        last_data_rows=len(all_rows),
        last_patched_end_row=data_end_row,
        last_sheet_used_end_row=sheet_used_end_row,
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="ドライバー情報をOneDrive Excel（営業用）へ反映"
    )
    parser.add_argument(
        "--login",
        action="store_true",
        help="OneDrive/SharePoint書き込み用の初回認証（ブラウザが開きます）",
    )
    parser.add_argument(
        "--device-code",
        action="store_true",
        help="デバイスコードで認証（ブラウザが使えない場合）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Excel取得・抽出のみ。OneDrive/SharePointは更新しない",
    )
    parser.add_argument(
        "--validate-headers",
        action="store_true",
        help="列マッピングと抽出サンプルを確認（アップロードしない）",
    )
    parser.add_argument(
        "--debug-dump",
        action="store_true",
        help="抽出結果を debug_extract_output.json に保存し詳細ログを出す",
    )
    args = parser.parse_args()

    setup_logging(quiet_console=is_unattended())

    try:
        if args.login and not args.dry_run:
            config = load_config()
            scopes = required_scopes(config)
            acquire_token(
                config,
                scopes=scopes,
                force_login=True,
                use_device_code=args.device_code,
            )
            logging.info(
                "認証完了。本番反映は python driver_sync.py を実行してください。"
            )
            return 0

        return run_sync(
            dry_run=args.dry_run,
            force_login=args.login,
            use_device_code=args.device_code,
            validate_headers_only=args.validate_headers,
            debug_dump=args.debug_dump,
        )
    except Exception as exc:
        logging.exception("同期失敗: %s", exc)
        write_status(ok=False, message="sync-failed", error=str(exc))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
