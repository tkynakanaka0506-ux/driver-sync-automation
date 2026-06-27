#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
出荷日別 案件サマリー 自動反映スクリプト（Graph API版）

driver_sync.py と同じ4社の元データ（車両依頼書・管理シート）を読み取り、
「出荷日」基準で当日〜数日先の案件数・案件名をOneDrive上の別Excelへ書き込む。
driver_sync.py の抽出ロジック・認証をそのまま再利用する（元の依頼書には書き込まない）。
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Any

import jpholiday
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from driver_sync import (
    GRAPH_SCOPES,
    JST,
    acquire_process_lock,
    acquire_token,
    download_share_file,
    extract_rows_from_workbook,
    is_unattended,
    load_config,
    prepare_rows_for_output,
    release_process_lock,
    rotate_log_if_needed,
    setup_logging,
    upload_onedrive_excel,
)

SCRIPT_DIR_LOG = "shipping_summary.log"

WEEKDAY_LABELS = ("月", "火", "水", "木", "金", "土", "日")

# 出荷日サマリー用の追加設定キー（driver_sync_config.json に追記して使う）。
# 未設定でも動くようにデフォルト値を持つ。
DEFAULT_OUTPUT_PATH = "/ドライバー情報/出荷日別案件サマリー.xlsx"
DEFAULT_SHEET_NAME = "出荷日サマリー"
DEFAULT_DAYS_AHEAD = 3

HEADER_FILL_HEX = "FFD9E1F2"
TODAY_FILL_HEX = "FFFFF2CC"
ZERO_FILL_HEX = "FFF2F2F2"


def ship_window_end(today: date, days_ahead: int = DEFAULT_DAYS_AHEAD) -> date:
    """出荷日表示範囲の終端日を求める。

    通常は today + days_ahead（暦日）。
    today が金曜日の場合は土日を挟むため +days_ahead+1（通常は火曜）まで延長する。
    さらに、表示範囲内のどこかの日が祝日なら、その祝日の翌日まで範囲を延長する
    （延長後の範囲にまた祝日が含まれる場合は安定するまで繰り返す）。
    """
    end = today + timedelta(days=days_ahead)
    if today.weekday() == 4:  # 0=月 ... 4=金
        end = today + timedelta(days=days_ahead + 1)

    while True:
        extended = end
        d = today
        while d <= end:
            if jpholiday.is_holiday(d):
                extended = max(extended, d + timedelta(days=1))
            d += timedelta(days=1)
        if extended == end:
            break
        end = extended
    return end


def md_to_date(md_str: str, today: date) -> date | None:
    """「6/27」形式の出荷日文字列を、todayに最も近い年の実際の日付に戻す。"""
    if not md_str:
        return None
    try:
        month, day = (int(p) for p in str(md_str).split("/"))
    except ValueError:
        return None
    candidates = [date(today.year + offset, month, day) for offset in (-1, 0, 1)]
    return min(candidates, key=lambda d: abs((d - today).days))


def collect_all_rows(config: dict[str, Any], today: date) -> list[dict[str, Any]]:
    """4社の元データから抽出する。着日フィルタは出荷日サマリーには使わないため、
    着日範囲を十分広く取り、出荷日側で別途フィルタする。"""
    header_rows = int(config.get("header_rows", 4))
    auto_detect_columns = bool(config.get("auto_detect_columns", False))

    all_rows: list[dict[str, Any]] = []
    for source in config["sources"]:
        key = source["key"]
        share_url = source["share_url"]
        logging.info("取得中: %s", key)
        content = download_share_file(share_url)
        workbook = openpyxl.load_workbook(BytesIO(content), read_only=False, data_only=True)
        rows = extract_rows_from_workbook(
            workbook,
            key,
            header_rows,
            today,
            arr_days_back=400,
            auto_detect_columns=auto_detect_columns,
        )
        workbook.close()
        logging.info("抽出件数: %s = %d", key, len(rows))
        all_rows.extend(rows)
    return prepare_rows_for_output(all_rows, today)


def filter_rows_by_ship_window(
    rows: list[dict[str, Any]], today: date, window_end: date
) -> list[dict[str, Any]]:
    filtered: list[dict[str, Any]] = []
    for row in rows:
        ship_date = md_to_date(str(row.get("出荷日", "")), today)
        if ship_date is None:
            continue
        if today <= ship_date <= window_end:
            row = dict(row)
            row["_出荷日付"] = ship_date
            filtered.append(row)
    return filtered


def group_by_ship_date(
    rows: list[dict[str, Any]], today: date, window_end: date
) -> list[dict[str, Any]]:
    """today〜window_end の全日付について（件数0の日も含めて）まとめる。"""
    by_date: dict[date, list[dict[str, Any]]] = {}
    for row in rows:
        by_date.setdefault(row["_出荷日付"], []).append(row)

    result: list[dict[str, Any]] = []
    d = today
    while d <= window_end:
        day_rows = by_date.get(d, [])
        case_lines = [
            f"{r.get('案件No', '')} {r.get('案件名', '')}".strip()
            for r in sorted(day_rows, key=lambda r: str(r.get("案件No", "")))
        ]
        result.append(
            {
                "date": d,
                "count": len(day_rows),
                "case_lines": case_lines,
            }
        )
        d += timedelta(days=1)
    return result


def format_date_label(d: date) -> str:
    return f"{d.month}/{d.day}({WEEKDAY_LABELS[d.weekday()]})"


def build_summary_xlsx_bytes(
    daily_summary: list[dict[str, Any]], today: date, window_end: date
) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = DEFAULT_SHEET_NAME

    ws["A1"] = (
        f"最終更新: {datetime.now(JST).strftime('%Y/%m/%d %H:%M')}　"
        f"（表示範囲: {format_date_label(today)}〜{format_date_label(window_end)}）"
    )
    ws["A1"].font = Font(bold=True)

    headers = ["出荷日", "件数", "案件一覧（案件No 案件名）"]
    header_row = 3
    for col, text in enumerate(headers, start=1):
        cell = ws.cell(header_row, col, text)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color=HEADER_FILL_HEX, end_color=HEADER_FILL_HEX, fill_type="solid")

    for idx, day in enumerate(daily_summary):
        row_num = header_row + 1 + idx
        d: date = day["date"]
        ws.cell(row_num, 1, format_date_label(d))
        ws.cell(row_num, 2, day["count"])
        ws.cell(row_num, 3, "\n".join(day["case_lines"]))
        ws.cell(row_num, 3).alignment = Alignment(wrap_text=True, vertical="top")

        if d == today:
            fill_hex = TODAY_FILL_HEX
        elif day["count"] == 0:
            fill_hex = ZERO_FILL_HEX
        else:
            fill_hex = None
        if fill_hex:
            fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
            for col in range(1, 4):
                ws.cell(row_num, col).fill = fill

        line_count = max(1, len(day["case_lines"]))
        ws.row_dimensions[row_num].height = max(20, 15 * line_count)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 80

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def run_summary(dry_run: bool = False, force_login: bool = False) -> int:
    try:
        acquire_process_lock()
    except FileExistsError:
        logging.warning("driver_sync.py が実行中のためスキップしました")
        return 0

    try:
        config = load_config()
        today = date.today()
        days_ahead = int(config.get("shipping_summary_days_ahead", DEFAULT_DAYS_AHEAD))
        window_end = ship_window_end(today, days_ahead)

        logging.info(
            "出荷日サマリー集計開始: today=%s, 表示範囲=%s〜%s",
            today,
            today,
            window_end,
        )

        all_rows = collect_all_rows(config, today)
        ship_rows = filter_rows_by_ship_window(all_rows, today, window_end)
        daily_summary = group_by_ship_date(ship_rows, today, window_end)

        for day in daily_summary:
            logging.info(
                "%s: %d件", format_date_label(day["date"]), day["count"]
            )

        if dry_run:
            logging.info("dry-run のため OneDrive 反映はスキップしました")
            return 0

        content = build_summary_xlsx_bytes(daily_summary, today, window_end)
        remote_path = config.get("shipping_summary_path", DEFAULT_OUTPUT_PATH)

        graph_token = acquire_token(
            config,
            scopes=GRAPH_SCOPES,
            force_login=force_login,
        )
        upload_onedrive_excel(graph_token, remote_path, content)
        logging.info("出荷日サマリーを更新しました: %s", remote_path)
        return 0
    except Exception as exc:
        logging.exception("出荷日サマリー更新失敗: %s", exc)
        return 1
    finally:
        release_process_lock()


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description="出荷日別案件サマリーをOneDrive Excelへ反映")
    parser.add_argument("--dry-run", action="store_true", help="集計のみ。OneDriveは更新しない")
    parser.add_argument("--login", action="store_true", help="初回認証（ブラウザが開きます）")
    args = parser.parse_args()

    rotate_log_if_needed()
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(SCRIPT_DIR_LOG, encoding="utf-8"),
        ]
        + ([] if is_unattended() else [logging.StreamHandler()]),
        force=True,
    )
    return run_summary(dry_run=args.dry_run, force_login=args.login)


if __name__ == "__main__":
    raise SystemExit(main())
