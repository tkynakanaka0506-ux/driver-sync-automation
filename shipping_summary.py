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
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from driver_sync import (
    GRAPH_SCOPES,
    JST,
    acquire_process_lock,
    acquire_token,
    cell_output_value,
    download_share_file,
    extract_rows_from_workbook,
    graph_get_drive_item,
    graph_read_range_values,
    load_config,
    openpyxl_fill_from_hex,
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

# 営業用Excelと同じ列構成（A〜G）をそのまま使う
SUMMARY_OUTPUT_COLUMNS = ["案件No", "案件名", "出荷日", "着日", "備考", "型式", "車型"]
HEADER_ROW = 6
DATA_START_ROW = 7


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


def sort_rows_by_ship_date(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """出荷日の昇順（同日内は案件No順）に並べ替える。"""
    return sorted(
        rows,
        key=lambda r: (r["_出荷日付"], str(r.get("案件No", ""))),
    )


def format_date_label(d: date) -> str:
    return f"{d.month}/{d.day}({WEEKDAY_LABELS[d.weekday()]})"


# 営業用Excel データ行(7行目)の実書式をGraph APIで調査した結果（2026-06-27確認）。
# 全列とも 中央揃え(横・縦)・折り返し表示・Meiryo 13pt・行高さ81.75pt で統一されている。
MAIN_FONT_NAME = "Meiryo"
MAIN_FONT_SIZE = 13.0
MAIN_ROW_HEIGHT = 81.75
DATA_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
# 列幅(pt)を実書式から取得し、openpyxlの文字幅単位に変換（目安: pt/7）
MAIN_COLUMN_WIDTH_PT = {"A": 113.25, "B": 264.0, "C": 74.25, "D": 78.0, "E": 119.25, "F": 355.5, "G": 126.75}
# 案件情報の塗り色は若干グレーで統一（営業用Excelの案件名別ハイライトは使わない）
CASE_INFO_FILL_HEX = "#EDEDED"
# 横持ち（案件名に「横持」を含む案件）は少し濃いグレーで区別する
YOKOMOCHI_FILL_HEX = "#BFBFBF"
YOKOMOCHI_KEYWORD = "横持"
# 出荷日ごとの区切り見出し行（黒背景・白文字）
DATE_HEADER_FILL_HEX = "FF000000"
DATE_HEADER_FONT = Font(name=MAIN_FONT_NAME, size=MAIN_FONT_SIZE, bold=True, color="FFFFFFFF")
DATE_HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
DATE_HEADER_ROW_HEIGHT = 24
SPACER_ROW_HEIGHT = 8
# 表全体に罫線を引く（薄いグレー）
TABLE_BORDER = Border(
    left=Side(style="thin", color="FFBFBFBF"),
    right=Side(style="thin", color="FFBFBFBF"),
    top=Side(style="thin", color="FFBFBFBF"),
    bottom=Side(style="thin", color="FFBFBFBF"),
)


def group_rows_by_date(
    rows: list[dict[str, Any]], today: date, window_end: date
) -> list[tuple[date, list[dict[str, Any]]]]:
    """today〜window_end の全日付を順に、その日の案件行とセットで返す（0件の日も含む）。"""
    by_date: dict[date, list[dict[str, Any]]] = {}
    for row in rows:
        by_date.setdefault(row["_出荷日付"], []).append(row)

    result: list[tuple[date, list[dict[str, Any]]]] = []
    d = today
    while d <= window_end:
        result.append((d, by_date.get(d, [])))
        d += timedelta(days=1)
    return result


def fetch_main_header_row(graph_token: str, config: dict[str, Any]) -> list[str]:
    """営業用Excelの6行目（A6:G6）の見出しをそのままコピーする。"""
    od = config.get("onedrive_output", {})
    remote_path = od.get("path", "/ドライバー情報/ドライバー情報_営業用.xlsx")
    sheet_name = od.get("sheet_name", "ドライバー情報")
    item = graph_get_drive_item(graph_token, remote_path)
    item_id = item["id"]
    address = f"A{HEADER_ROW}:G{HEADER_ROW}"
    values = graph_read_range_values(graph_token, item_id, sheet_name, address, session_id=None)
    if values and values[0]:
        return [str(v) if v is not None else "" for v in values[0]]
    return list(SUMMARY_OUTPUT_COLUMNS)


def build_summary_xlsx_bytes(
    rows: list[dict[str, Any]],
    header_values: list[str],
    today: date,
    window_end: date,
    config: dict[str, Any],
) -> bytes:
    """営業用Excelと同じA〜G列構成（6行目=見出し、7行目〜=データ）で出力する。"""
    case_min, case_max = 1, len(SUMMARY_OUTPUT_COLUMNS)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = DEFAULT_SHEET_NAME

    ws["A1"] = (
        f"最終更新: {datetime.now(JST).strftime('%Y/%m/%d %H:%M')}　"
        f"（表示範囲: {format_date_label(today)}〜{format_date_label(window_end)}・"
        f"{len(rows)}件）"
    )
    ws["A1"].font = Font(name=MAIN_FONT_NAME, size=MAIN_FONT_SIZE, bold=True)

    for col, text in enumerate(header_values[: len(SUMMARY_OUTPUT_COLUMNS)], start=1):
        cell = ws.cell(HEADER_ROW, col, text)
        cell.font = Font(name=MAIN_FONT_NAME, size=MAIN_FONT_SIZE, bold=True)
        cell.alignment = DATA_ALIGNMENT
        cell.border = TABLE_BORDER
    ws.row_dimensions[HEADER_ROW].height = MAIN_ROW_HEIGHT

    last_col_letter = openpyxl.utils.get_column_letter(case_max)
    row_num = DATA_START_ROW
    for day, day_rows in group_rows_by_date(rows, today, window_end):
        ws.row_dimensions[row_num].height = SPACER_ROW_HEIGHT
        row_num += 1

        ws.merge_cells(f"A{row_num}:{last_col_letter}{row_num}")
        header_cell = ws.cell(row_num, 1, f"{format_date_label(day)} 出荷数：{len(day_rows)}件")
        header_cell.font = DATE_HEADER_FONT
        header_cell.alignment = DATE_HEADER_ALIGNMENT
        header_fill = PatternFill(start_color=DATE_HEADER_FILL_HEX, end_color=DATE_HEADER_FILL_HEX, fill_type="solid")
        for col in range(case_min, case_max + 1):
            ws.cell(row_num, col).fill = header_fill
            ws.cell(row_num, col).border = TABLE_BORDER
        ws.row_dimensions[row_num].height = DATE_HEADER_ROW_HEIGHT
        row_num += 1

        for day_row in day_rows:
            is_yokomochi = YOKOMOCHI_KEYWORD in str(day_row.get("案件名", ""))
            for col, name in enumerate(SUMMARY_OUTPUT_COLUMNS, start=1):
                value = cell_output_value(day_row, name)
                cell = ws.cell(row_num, col, value)
                cell.font = Font(name=MAIN_FONT_NAME, size=MAIN_FONT_SIZE)
                cell.alignment = DATA_ALIGNMENT

            fill = openpyxl_fill_from_hex(YOKOMOCHI_FILL_HEX if is_yokomochi else CASE_INFO_FILL_HEX)
            for col in range(case_min, case_max + 1):
                ws.cell(row_num, col).fill = fill
                ws.cell(row_num, col).border = TABLE_BORDER

            ws.row_dimensions[row_num].height = MAIN_ROW_HEIGHT
            row_num += 1

    for col_letter, width_pt in MAIN_COLUMN_WIDTH_PT.items():
        ws.column_dimensions[col_letter].width = round(width_pt / 7, 1)

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
        ship_rows = sort_rows_by_ship_date(ship_rows)

        counts: dict[date, int] = {}
        for row in ship_rows:
            counts[row["_出荷日付"]] = counts.get(row["_出荷日付"], 0) + 1
        for d, count in sorted(counts.items()):
            logging.info("%s: %d件", format_date_label(d), count)

        if dry_run:
            logging.info("dry-run のため OneDrive 反映はスキップしました")
            return 0

        graph_token = acquire_token(
            config,
            scopes=GRAPH_SCOPES,
            force_login=force_login,
        )
        header_values = fetch_main_header_row(graph_token, config)
        content = build_summary_xlsx_bytes(ship_rows, header_values, today, window_end, config)
        remote_path = config.get("shipping_summary_path", DEFAULT_OUTPUT_PATH)
        upload_onedrive_excel(graph_token, remote_path, content)
        logging.info("出荷日サマリーを更新しました: %s (%d件)", remote_path, len(ship_rows))
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
            logging.StreamHandler(),
        ],
        force=True,
    )
    return run_summary(dry_run=args.dry_run, force_login=args.login)


if __name__ == "__main__":
    raise SystemExit(main())
