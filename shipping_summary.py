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
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from io import BytesIO
from typing import Any

import jpholiday
import openpyxl

from driver_sync import (
    GRAPH_BASE,
    GRAPH_SCOPES,
    JST,
    acquire_process_lock,
    acquire_token,
    cell,
    cell_output_value,
    col_letter_from_index,
    download_share_file,
    extract_rows_from_workbook,
    graph_batch_clear_fills,
    graph_batch_patch_fills,
    graph_batch_set_row_heights,
    graph_close_workbook_session,
    graph_create_workbook_session,
    graph_get_drive_item,
    graph_list_tables_on_sheet,
    graph_patch_range_values,
    graph_read_range_values,
    graph_request_with_retry,
    graph_resize_table,
    graph_resolve_worksheet_name,
    is_highlight_case_name,
    load_case_name_row_color_map,
    load_config,
    normalize_case_name_key,
    prepare_rows_for_output,
    release_process_lock,
    rotate_log_if_needed,
    setup_logging,
    to_md,
    upload_onedrive_excel,
    worksheet_segment,
)

SCRIPT_DIR_LOG = "shipping_summary.log"

WEEKDAY_LABELS = ("月", "火", "水", "木", "金", "土", "日")

# 出荷日サマリー用の追加設定キー（driver_sync_config.json に追記して使う）。
# 未設定でも動くようにデフォルト値を持つ。
DEFAULT_OUTPUT_PATH = "/ドライバー情報/出荷日別案件サマリー.xlsx"
DEFAULT_SHEET_NAME = "出荷日サマリー"
DEFAULT_DAYS_AHEAD = 3

# A列:案件No、B列:納入先住所（新規）、C列以降は営業用Excelと同じ構成
SUMMARY_OUTPUT_COLUMNS = ["案件No", "納入先住所", "案件名", "出荷日", "着日", "備考", "型式", "車型"]
HEADER_ROW = 6

# KSサポート（出荷日サマリー専用の追加抽出元。営業用Excel側の4社とは別枠・別構造のため
# driver_sync_config.json の sources / SHEET_CONFIG には追加せず、ここに直接持つ）。
# レイアウト: 4行目が見出し、5行目からデータ。A=ステータス B=案件No C=出荷日 D=着日
# E=着時間 F=納入先住所 G=届け先名(=案件名) H=出荷工場 I=製品種別。
KS_SOURCE_SHARE_URL = (
    "https://onedrive.live.com/:x:/g/personal/1331A7580E0E4466/"
    "IQBmRA4OWKcxIIAThgEAAAAAAU-iRPzHNuNnvea3gng6R4w"
    "?rtime=Cylv1YGT3kg&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy8xMzMxQTc1ODBFMEU0NDY2L0lRQm1SQTRPV0tjeElJQVRoZ0VBQUFBQUFVLWlSUHpITnVObnZlYTNnbmc2UjR3P2U9NDo2YVpUUEgmc2hhcmluZ3YyPXRydWUmZnJvbVNoYXJlPXRydWUmYXQ9OQ"
)
KS_SOURCE_KEY = "ks"
KS_SHEET_NAME = "車両依頼書"
KS_HEADER_ROWS = 4
KS_STATUS_COL = 1
KS_AN_NO_COL = 2
KS_SHIP_COL = 3
KS_ARR_COL = 4
KS_ADDRESS_COL = 6
KS_CASE_NAME_COL = 7
KS_PRODUCT_TYPE_COL = 9
KS_ACCEPT_STATUS = "配車確定"
KS_CAR_TYPE_LABEL = "KS"


def extract_ks_rows(workbook: openpyxl.Workbook, today: date) -> list[dict[str, Any]]:
    """KSサポートの依頼書から、ステータスが「配車確定」の行だけを抽出する。

    製品種別(I列)は備考②（出力上は「型式」キー）に、車型は固定で「KS」と表示する。
    """
    ws = workbook[KS_SHEET_NAME]
    fiscal_year_start = date(today.year if today.month >= 4 else today.year - 1, 4, 1)
    rows: list[dict[str, Any]] = []
    for r in range(KS_HEADER_ROWS + 1, ws.max_row + 1):
        status = str(ws.cell(row=r, column=KS_STATUS_COL).value or "").strip()
        if status != KS_ACCEPT_STATUS:
            continue
        an_no = str(ws.cell(row=r, column=KS_AN_NO_COL).value or "").strip()
        case_name = str(ws.cell(row=r, column=KS_CASE_NAME_COL).value or "").strip()
        if not an_no and not case_name:
            continue
        # 今年度より前の行はスキップ
        ship_val = ws.cell(row=r, column=KS_SHIP_COL).value
        if isinstance(ship_val, datetime):
            ship_val = ship_val.date()
        if isinstance(ship_val, date) and ship_val < fiscal_year_start:
            continue
        product_type = str(ws.cell(row=r, column=KS_PRODUCT_TYPE_COL).value or "").strip()
        remark2 = f"{product_type} / {KS_CAR_TYPE_LABEL}" if product_type else KS_CAR_TYPE_LABEL
        address = str(ws.cell(row=r, column=KS_ADDRESS_COL).value or "").strip()
        rows.append(
            {
                "案件No": an_no,
                "納入先住所": address,
                "案件名": case_name,
                "出荷日": to_md(ws.cell(row=r, column=KS_SHIP_COL).value, today),
                "着日": to_md(ws.cell(row=r, column=KS_ARR_COL).value, today),
                "備考": "",
                "型式": remark2,
                "車型": "",
                "依頼先": KS_SOURCE_KEY,
                "行キー": f"{KS_SOURCE_KEY}|{KS_SHEET_NAME}|{r}",
            }
        )
    return rows


# パーツ➔発送（出荷日サマリー専用の追加抽出元、KSサポートと同じ構造）。
# A〜G列はKSサポートと同じ。M列(運送業者)を車型に反映する。
PARTS_SOURCE_SHARE_URL = (
    "https://onedrive.live.com/:x:/g/personal/1331A7580E0E4466/"
    "IQBOFWrnEII9R4qy9nCbksTDAbGkrg8q69j8umLQjpZyNMU"
    "?resid=1331A7580E0E4466!se76a154e8210473d8ab2f6709b92c4c3&ithint=file%2Cxlsx&e=4%3Ald9U5q"
    "&sharingv2=true&fromShare=true&at=9&migratedtospo=true"
    "&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy8xMzMxQTc1ODBFMEU0NDY2L0lRQk9GV3JuRUlJOVI0cXk5bkNia3NUREFiR2tyZzhxNjlqOHVtTFFqcFp5Tk1VP2U9NDpsZDlVNXEmc2hhcmluZ3YyPXRydWUmZnJvbVNoYXJlPXRydWUmYXQ9OQ"
)
PARTS_SOURCE_KEY = "parts"
PARTS_SHEET_NAME = "車両依頼書"
PARTS_HEADER_ROWS = 4
PARTS_STATUS_COL = 1
PARTS_AN_NO_COL = 2
PARTS_SHIP_COL = 3
PARTS_ARR_COL = 4
PARTS_ADDRESS_COL = 6
PARTS_CASE_NAME_COL = 7
PARTS_CARRIER_COL = 13
PARTS_ACCEPT_STATUS = "配車確定"
PARTS_LABEL = "パーツ"


def extract_parts_rows(workbook: openpyxl.Workbook, today: date) -> list[dict[str, Any]]:
    """パーツ➔発送の依頼書から、ステータスが「配車確定」の行だけを抽出する。

    M列(運送業者)を車型に反映する（KSサポートと違い固定値ではなく実際の業者名）。
    """
    ws = workbook[PARTS_SHEET_NAME]
    fiscal_year_start = date(today.year if today.month >= 4 else today.year - 1, 4, 1)
    rows: list[dict[str, Any]] = []
    for r in range(PARTS_HEADER_ROWS + 1, ws.max_row + 1):
        status = str(ws.cell(row=r, column=PARTS_STATUS_COL).value or "").strip()
        if status != PARTS_ACCEPT_STATUS:
            continue
        an_no = str(ws.cell(row=r, column=PARTS_AN_NO_COL).value or "").strip()
        case_name = str(ws.cell(row=r, column=PARTS_CASE_NAME_COL).value or "").strip()
        if not an_no and not case_name:
            continue
        # 今年度より前の行はスキップ
        ship_val = ws.cell(row=r, column=PARTS_SHIP_COL).value
        if isinstance(ship_val, datetime):
            ship_val = ship_val.date()
        if isinstance(ship_val, date) and ship_val < fiscal_year_start:
            continue
        carrier = str(ws.cell(row=r, column=PARTS_CARRIER_COL).value or "").strip()
        remark2 = f"{PARTS_LABEL} / {carrier}" if carrier else PARTS_LABEL
        address = str(ws.cell(row=r, column=PARTS_ADDRESS_COL).value or "").strip()
        rows.append(
            {
                "案件No": an_no,
                "納入先住所": address,
                "案件名": case_name,
                "出荷日": to_md(ws.cell(row=r, column=PARTS_SHIP_COL).value, today),
                "着日": to_md(ws.cell(row=r, column=PARTS_ARR_COL).value, today),
                "備考": "",
                "型式": remark2,
                "車型": "",
                "依頼先": PARTS_SOURCE_KEY,
                "行キー": f"{PARTS_SOURCE_KEY}|{PARTS_SHEET_NAME}|{r}",
            }
        )
    return rows


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


# 各依頼先の元データにおける納入先住所の列番号（1始まり）
# matsuzaki のみ G列(7)、その他は F列(6)
SOURCE_ADDRESS_COL: dict[str, int] = {
    "matsuzaki": 7,
    "fukuoka": 6,
    "nakadori": 6,
    "maruun": 6,
}

# 各依頼先の元データにおける出荷日の列番号（1始まり）。SHEET_CONFIG の ship フィールドと対応。
# 実際の日付オブジェクトを読み返して今年度フィルターに使う。
SOURCE_SHIP_COL: dict[str, int] = {
    "matsuzaki": 4,   # D列
    "fukuoka": 3,     # C列
    "nakadori": 3,    # C列
    "maruun": 5,      # E列
}


def collect_all_rows(config: dict[str, Any], today: date, window_end: date) -> list[dict[str, Any]]:
    """4社の元データから抽出する。着日でなく出荷日基準で抽出範囲を判定する
    （window_date_field="ship"）。

    require_driver_info=False: 横持ち（中継）案件などドライバー4項目が未確定の行も
    営業用Excelとは異なり除外しない（出荷日の有無だけが知りたいため）。
    """
    header_rows = int(config.get("header_rows", 4))
    auto_detect_columns = bool(config.get("auto_detect_columns", False))

    # 今年度（4月1日〜翌年3月31日）の開始日
    fiscal_year_start = date(today.year if today.month >= 4 else today.year - 1, 4, 1)

    all_rows: list[dict[str, Any]] = []
    for source in config["sources"]:
        key = source["key"]
        share_url = source["share_url"]
        logging.info("取得中: %s", key)
        content = download_share_file(share_url)
        workbook = openpyxl.load_workbook(BytesIO(content), read_only=False, data_only=True)
        # 出荷日サマリーでは滋賀管理シート分は対象外（要望により除外。営業用Excel側は対象のまま）
        excluded_sheet_keywords = ("滋賀", "管理") if key == "nakadori" else ()
        # 福岡ロジテックのF列(納入先住所)に「御社積み」がある行は、外部倉庫から現場へ
        # 出庫する分（宇美工場からの出荷ではない）なので出荷日サマリーでは除外する
        # （要望により出荷日サマリーのみ。営業用Excel側は対象のまま）
        # 松崎運輸のG列(納入先住所)に「司鳥栖倉庫積込み」がある行も同様に除外する
        if key == "fukuoka":
            row_exclude_predicate = lambda row_tuple: "御社積み" in cell(row_tuple, "F")
        elif key == "matsuzaki":
            row_exclude_predicate = lambda row_tuple: "司鳥栖倉庫積込み" in cell(row_tuple, "G")
        else:
            row_exclude_predicate = None
        rows = extract_rows_from_workbook(
            workbook,
            key,
            header_rows,
            today,
            arr_days_back=400,
            auto_detect_columns=auto_detect_columns,
            arr_window_end_override=window_end,
            require_driver_info=False,
            window_date_field="ship",
            excluded_sheet_keywords=excluded_sheet_keywords,
            row_exclude_predicate=row_exclude_predicate,
        )
        # 行キーからシート名・行番号を逆引きして納入先住所を付加し、
        # 今年度（fiscal_year_start 以降）の行だけを残す。
        # to_md() が年情報を失うため md_to_date() が古い日付を今年の同日と誤認識する問題を
        # ここで実際のセル値（date オブジェクト）を確認することで解決する。
        addr_col = SOURCE_ADDRESS_COL.get(key, 6)
        ship_col = SOURCE_SHIP_COL.get(key, 3)
        valid_rows: list[dict[str, Any]] = []
        for row in rows:
            row_key_parts = row.get("行キー", "").split("|")
            if len(row_key_parts) >= 3:
                try:
                    ws_ref = workbook[row_key_parts[1]]
                    rnum = int(row_key_parts[2])
                    row["納入先住所"] = str(ws_ref.cell(row=rnum, column=addr_col).value or "").strip()
                    ship_val = ws_ref.cell(row=rnum, column=ship_col).value
                    if isinstance(ship_val, datetime):
                        ship_val = ship_val.date()
                    if isinstance(ship_val, date) and ship_val < fiscal_year_start:
                        continue  # 今年度より前の行はスキップ
                except Exception:
                    row["納入先住所"] = ""
            else:
                row["納入先住所"] = ""
            valid_rows.append(row)
        rows = valid_rows
        workbook.close()
        logging.info("抽出件数: %s = %d", key, len(rows))
        all_rows.extend(rows)

    logging.info("取得中: %s", KS_SOURCE_KEY)
    ks_content = download_share_file(KS_SOURCE_SHARE_URL)
    ks_workbook = openpyxl.load_workbook(BytesIO(ks_content), read_only=False, data_only=True)
    ks_rows = extract_ks_rows(ks_workbook, today)
    ks_workbook.close()
    logging.info("抽出件数: %s = %d", KS_SOURCE_KEY, len(ks_rows))
    all_rows.extend(ks_rows)

    logging.info("取得中: %s", PARTS_SOURCE_KEY)
    parts_content = download_share_file(PARTS_SOURCE_SHARE_URL)
    parts_workbook = openpyxl.load_workbook(BytesIO(parts_content), read_only=False, data_only=True)
    parts_rows = extract_parts_rows(parts_workbook, today)
    parts_workbook.close()
    # KSサポートと案件No(案件No)が重複する分は二重表示を避けるため除外する（要望により）
    ks_an_nos = {str(r.get("案件No", "")) for r in ks_rows if r.get("案件No")}
    parts_rows = [r for r in parts_rows if str(r.get("案件No", "")) not in ks_an_nos]
    logging.info("抽出件数: %s = %d", PARTS_SOURCE_KEY, len(parts_rows))
    all_rows.extend(parts_rows)

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


def row_sort_category(row: dict[str, Any], color_map: dict[str, str]) -> int:
    """同日内の表示順カテゴリ: KS→パーツ→横持ち→横持ち以外（直行）。"""
    source = row.get("依頼先")
    if source == KS_SOURCE_KEY:
        return 0
    if source == PARTS_SOURCE_KEY:
        return 1
    if is_relay_highlight_case_name(str(row.get("案件名", "")), color_map):
        return 2
    return 3


def sort_rows_by_ship_date(
    rows: list[dict[str, Any]], color_map: dict[str, str] | None = None
) -> list[dict[str, Any]]:
    """出荷日の昇順、同日内はKS→パーツ→横持ち→横持ち以外（直行）の順、
    さらに同カテゴリ内は案件No順に並べ替える。"""
    color_map = color_map or {}
    return sorted(
        rows,
        key=lambda r: (
            r["_出荷日付"],
            row_sort_category(r, color_map),
            str(r.get("案件No", "")),
        ),
    )


def format_date_label(d: date) -> str:
    return f"{d.month}/{d.day}({WEEKDAY_LABELS[d.weekday()]})"


# 営業用Excel データ行(7行目)の実書式をGraph APIで調査した結果（2026-06-27確認）。
# フォント・行高さは引き続き合わせる。色はExcelテーブル機能（バンド）に一本化したため、
# 案件名別の手動色分け（横持ち等）は廃止。
MAIN_FONT_NAME = "Meiryo"
MAIN_FONT_SIZE = 13.0
MAIN_ROW_HEIGHT = 81.75
# 見出し行は1行テキストのみなのでデータ行より低くする
HEADER_ROW_HEIGHT = 24.0
# 列幅(pt)を実書式から取得（Graph APIのcolumnWidthはptそのまま使える）
MAIN_COLUMN_WIDTH_PT = {
    "A": 113.25,  # 案件No
    "B": 355.5,   # 納入先住所（新規）
    "C": 264.0,   # 案件名（旧B）
    "D": 74.25,   # 出荷日（旧C）
    "E": 78.0,    # 着日（旧D）
    "F": 119.25,  # 備考①（旧E）
    "G": 355.5,   # 備考②（旧F）
    "H": 126.75,  # 車型（旧G）
}
# Excel組み込みテーブルスタイル「赤、テーブルスタイル（中間）17」の内部名
TABLE_STYLE_NAME = "TableStyleMedium17"
# 出荷日ごとの区切り見出し行（黒背景・白文字。テーブル内の通常行として挟む）
DATE_HEADER_FILL_HEX = "#000000"
DATE_HEADER_ROW_HEIGHT = 24.0
# 横持ち等の強調行は営業用Excel側の薄い青ではなく、サマリーでは少し濃めの灰色で区別する。
# それ以外の通常案件は薄い灰色で統一する（テーブルスタイルの交互配色を上書きする）。
HIGHLIGHT_CASE_FILL_HEX = "#E6E6E6"
DEFAULT_DATA_ROW_FILL_HEX = "#F7F7F7"
# KSサポート分は出荷元が異なるため薄い赤色で区別する
KS_ROW_FILL_HEX = "#FFE3E6"

# 法人格表記（株式会社/㈱ など）の有無が案件名表記でブレるため、driver_sync_config.json の
# row_colors_by_case_name には一致しない場合がある（例:「司企業株式会社　鳥栖営業所」と
# 「司企業　鳥栖営業所」）。サマリー側だけ、法人格表記を取り除いた上で再照合する。
CORPORATE_AFFIXES = ("株式会社", "㈱", "(本社)", "（本社）")


def normalize_case_name_loose(name: str) -> str:
    text = normalize_case_name_key(name)
    for affix in CORPORATE_AFFIXES:
        text = text.replace(affix, "")
    return re.sub(r"\s+", " ", text).strip()


def is_relay_highlight_case_name(case_name: str, color_map: dict[str, str]) -> bool:
    if is_highlight_case_name(case_name, color_map):
        return True
    loose_name = normalize_case_name_loose(case_name)
    return any(loose_name == normalize_case_name_loose(key) for key in color_map)


def graph_range_format_url(item_id: str, sheet_name: str, address: str) -> str:
    seg = worksheet_segment(sheet_name)
    return f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/range(address='{address}')/format"


def graph_unmerge_range(graph_token: str, item_id: str, sheet_name: str, address: str, session_id: str) -> None:
    seg = worksheet_segment(sheet_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/range(address='{address}')/unmerge"
    graph_request_with_retry("POST", url, graph_token, session_id=session_id)


def graph_clear_range(
    graph_token: str, item_id: str, sheet_name: str, address: str, session_id: str, apply_to: str = "All"
) -> None:
    seg = worksheet_segment(sheet_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/range(address='{address}')/clear"
    graph_request_with_retry("POST", url, graph_token, session_id=session_id, json={"applyTo": apply_to})


def graph_create_table(
    graph_token: str, item_id: str, sheet_name: str, address: str, session_id: str, *, has_headers: bool = True
) -> str:
    seg = worksheet_segment(sheet_name)
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/tables/add"
    res = graph_request_with_retry(
        "POST", url, graph_token, session_id=session_id, json={"address": address, "hasHeaders": has_headers}
    )
    if not res.ok:
        raise RuntimeError(f"テーブル作成失敗: {res.status_code} {res.text[:300]}")
    return str(res.json()["name"])


def graph_set_table_style(graph_token: str, item_id: str, table_name: str, session_id: str, style_name: str) -> None:
    url = f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/tables('{table_name}')"
    res = graph_request_with_retry("PATCH", url, graph_token, session_id=session_id, json={"style": style_name})
    if not res.ok:
        raise RuntimeError(f"テーブルスタイル設定失敗: {res.status_code} {res.text[:300]}")


def graph_set_range_alignment(
    graph_token: str,
    item_id: str,
    sheet_name: str,
    address: str,
    session_id: str,
    *,
    horizontal: str,
    vertical: str,
    wrap_text: bool,
) -> None:
    url = graph_range_format_url(item_id, sheet_name, address)
    graph_request_with_retry(
        "PATCH",
        url,
        graph_token,
        session_id=session_id,
        json={"horizontalAlignment": horizontal, "verticalAlignment": vertical, "wrapText": wrap_text},
    )


def graph_set_range_font(
    graph_token: str, item_id: str, sheet_name: str, address: str, session_id: str, **font_props: Any
) -> None:
    url = f"{graph_range_format_url(item_id, sheet_name, address)}/font"
    graph_request_with_retry("PATCH", url, graph_token, session_id=session_id, json=font_props)


BORDER_EDGES = ("EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal")


def graph_set_range_border_color(
    graph_token: str, item_id: str, sheet_name: str, address: str, session_id: str, color_hex: str
) -> None:
    seg = worksheet_segment(sheet_name)
    for edge in BORDER_EDGES:
        url = (
            f"{GRAPH_BASE}/me/drive/items/{item_id}/workbook/{seg}/"
            f"range(address='{address}')/format/borders('{edge}')"
        )
        graph_request_with_retry(
            "PATCH", url, graph_token, session_id=session_id,
            json={"color": color_hex, "style": "Continuous"},
        )


def graph_set_column_width(
    graph_token: str, item_id: str, sheet_name: str, address: str, session_id: str, width_pt: float
) -> None:
    url = graph_range_format_url(item_id, sheet_name, address)
    graph_request_with_retry("PATCH", url, graph_token, session_id=session_id, json={"columnWidth": width_pt})


def ensure_summary_file_exists(graph_token: str, remote_path: str) -> dict[str, Any]:
    """サマリーExcelが無ければ、シート名だけのブックを新規作成する。"""
    try:
        return graph_get_drive_item(graph_token, remote_path)
    except RuntimeError as exc:
        if "404" not in str(exc):
            raise
        logging.info("サマリーExcelが存在しないため新規作成します: %s", remote_path)
        wb = openpyxl.Workbook()
        wb.active.title = DEFAULT_SHEET_NAME
        buf = BytesIO()
        wb.save(buf)
        upload_onedrive_excel(graph_token, remote_path, buf.getvalue())
        return graph_get_drive_item(graph_token, remote_path)


# E列(備考)・F列(型式)・G列(車型)は営業用Excelの見出し文字のままだと意味が伝わらないため、
# サマリー側だけ表示ラベルを上書きする（列番号=SUMMARY_OUTPUT_COLUMNSのindexで指定、データ自体は変えない）。
HEADER_LABEL_OVERRIDES = {5: "備考①", 6: "備考②", 7: "車型"}


def fetch_main_header_row(graph_token: str, config: dict[str, Any]) -> list[str]:
    """営業用Excelの6行目（A6:G6）の見出しをコピーし、B列に納入先住所を挿入した8列分を返す。"""
    od = config.get("onedrive_output", {})
    remote_path = od.get("path", "/ドライバー情報/ドライバー情報_営業用.xlsm")
    sheet_name = od.get("sheet_name", "ドライバー情報")
    item = graph_get_drive_item(graph_token, remote_path)
    item_id = item["id"]
    address = f"A{HEADER_ROW}:G{HEADER_ROW}"
    values = graph_read_range_values(graph_token, item_id, sheet_name, address, session_id=None)
    if values and values[0]:
        orig = [str(v) if v is not None else "" for v in values[0]]
    else:
        # フォールバック: SUMMARY_OUTPUT_COLUMNS から納入先住所以外を使用
        orig = [c for c in SUMMARY_OUTPUT_COLUMNS if c != "納入先住所"]
    # 営業用ExcelはA〜G列（7列）なので、B位置に「納入先住所」を挿入して8列にする
    header_values = [orig[0], "納入先住所"] + orig[1:]
    for col_index, label in HEADER_LABEL_OVERRIDES.items():
        if col_index < len(header_values):
            header_values[col_index] = label
    return header_values


def group_rows_by_date(
    rows: list[dict[str, Any]], today: date, window_end: date
) -> list[tuple[date, list[dict[str, Any]]]]:
    """表示範囲内の全日付（0件の日も含む）について、出荷日昇順でグループ化する。"""
    by_date: dict[date, list[dict[str, Any]]] = {}
    for row in rows:
        by_date.setdefault(row["_出荷日付"], []).append(row)

    groups: list[tuple[date, list[dict[str, Any]]]] = []
    d = today
    while d <= window_end:
        groups.append((d, by_date.get(d, [])))
        d += timedelta(days=1)
    return groups


def build_summary_layout(
    rows: list[dict[str, Any]],
    header_values: list[str],
    today: date,
    window_end: date,
    color_map: dict[str, str] | None = None,
) -> dict[str, Any]:
    """営業用Excelと同じA〜G列構成で、出荷日ごとの「出荷数：N件」行をテーブル内の
    通常行として挟みつつ書き込む値・行高さ・塗り色を組み立てる。

    Tableは結合セルを含められないため結合はしない（A列のみに文言を入れ、他列は空）。
    見出し行は1つだけ（テーブル全体で1つのExcelテーブルとして扱う＝既存テーブルは
    削除せずリサイズのみで、罫線などの手動書式を保持する）。
    """
    case_max = len(SUMMARY_OUTPUT_COLUMNS)
    last_col_letter = col_letter_from_index(case_max)
    color_map = color_map or {}

    matrix: list[list[Any]] = []
    row_heights: list[tuple[int, float]] = []
    date_header_fill_rows: list[int] = []
    date_header_font_rows: list[int] = []
    data_row_fills: list[tuple[int, str]] = []

    matrix.append(
        [
            (
                f"最終更新: {datetime.now(JST).strftime('%Y/%m/%d %H:%M')}　"
                f"（表示範囲: {format_date_label(today)}〜{format_date_label(window_end)}・"
                f"{len(rows)}件）"
            ),
            *([""] * (case_max - 1)),
        ]
    )
    for _ in range(2, HEADER_ROW):
        matrix.append([""] * case_max)

    matrix.append((header_values[:case_max] + [""] * case_max)[:case_max])
    row_heights.append((HEADER_ROW, HEADER_ROW_HEIGHT))

    row_num = HEADER_ROW
    groups = group_rows_by_date(rows, today, window_end)
    for d, day_rows in groups:
        # 黒背景白文字の出荷数行（テーブル内の通常行として挟む。結合はしない）
        row_num += 1
        date_header_row = row_num
        matrix.append([f"{format_date_label(d)} 出荷数：{len(day_rows)}件"] + [""] * (case_max - 1))
        row_heights.append((date_header_row, DATE_HEADER_ROW_HEIGHT))
        date_header_fill_rows.append(date_header_row)
        date_header_font_rows.append(date_header_row)

        for row in day_rows:
            row_num += 1
            matrix.append([cell_output_value(row, name) for name in SUMMARY_OUTPUT_COLUMNS])
            row_heights.append((row_num, MAIN_ROW_HEIGHT))
            case_name = str(row.get("案件名", ""))
            if row.get("依頼先") in (KS_SOURCE_KEY, PARTS_SOURCE_KEY):
                data_row_fills.append((row_num, KS_ROW_FILL_HEX))
            elif is_relay_highlight_case_name(case_name, color_map):
                data_row_fills.append((row_num, HIGHLIGHT_CASE_FILL_HEX))
            else:
                data_row_fills.append((row_num, DEFAULT_DATA_ROW_FILL_HEX))

    return {
        "matrix": matrix,
        "row_heights": row_heights,
        "last_row": row_num,
        "last_col_letter": last_col_letter,
        "date_header_fill_rows": date_header_fill_rows,
        "date_header_font_rows": date_header_font_rows,
        "data_row_fills": data_row_fills,
    }


CLEAR_RANGE = "A1:H500"


def write_summary_via_graph(
    graph_token: str,
    remote_path: str,
    layout: dict[str, Any],
) -> None:
    """Excel Workbook API（範囲PATCH・本テーブル機能）で書き込む。ファイルが開かれていても更新できる。"""
    item = ensure_summary_file_exists(graph_token, remote_path)
    item_id = item["id"]
    session_id = graph_create_workbook_session(graph_token, item_id)
    try:
        sheet_name = graph_resolve_worksheet_name(graph_token, item_id, DEFAULT_SHEET_NAME, session_id)

        # 罫線など手動で編集した書式は壊さないよう、値だけクリアする（書式はクリアしない）。
        # ただし塗り色だけは前回実行分が残ると行数減少時に古い黒/灰色が残ってしまうため、
        # 罫線は触らずに塗り色のみリセットしてから今回分を塗り直す。
        graph_unmerge_range(graph_token, item_id, sheet_name, CLEAR_RANGE, session_id)
        graph_clear_range(graph_token, item_id, sheet_name, CLEAR_RANGE, session_id, apply_to="Contents")
        graph_batch_clear_fills(graph_token, item_id, sheet_name, [CLEAR_RANGE], session_id)
        # 文字色も塗り色と同様、前回実行分の白文字（出荷数行用）が行数減少時に残ってしまう
        # ため、罫線以外を触らない範囲で文字色だけ黒にリセットしてから塗り直す。
        # 太字も塗り色・文字色と同様、前回実行分のタイトル行/見出し行用の太字が
        # 行数減少時に残ってしまうため、罫線以外を触らない範囲で太字もリセットする。
        graph_set_range_font(
            graph_token, item_id, sheet_name, CLEAR_RANGE, session_id, color="#000000", bold=False,
        )

        last_row = layout["last_row"]
        last_col_letter = layout["last_col_letter"]
        data_address = f"A1:{last_col_letter}{last_row}"

        graph_patch_range_values(graph_token, item_id, sheet_name, data_address, layout["matrix"], session_id)

        # テーブルの作成・リサイズはテーブルスタイル（バンディング）や既定の配置・フォントを
        # 範囲全体に適用し直すため、手動の配置・塗り色・文字色より先に行う
        # （後から上書きすれば、テーブル側の処理で消されない）。
        # 既存テーブルがあればリサイズのみ（削除→再作成だと罫線などの手動編集が消える）。
        # スタイルは初回作成時だけ設定し、以降は触らない。
        table_address = f"A{HEADER_ROW}:{last_col_letter}{last_row}"
        existing_tables = graph_list_tables_on_sheet(graph_token, item_id, sheet_name, session_id)
        if existing_tables:
            graph_resize_table(graph_token, item_id, str(existing_tables[0]["name"]), table_address, session_id)
        else:
            table_name = graph_create_table(graph_token, item_id, sheet_name, table_address, session_id)
            graph_set_table_style(graph_token, item_id, table_name, session_id, TABLE_STYLE_NAME)

        # テーブル全体に黒の格子罫線を描く（出荷数行だけは後段で白に上書きする）
        graph_set_range_border_color(
            graph_token, item_id, sheet_name, table_address, session_id, "#000000",
        )

        graph_set_range_alignment(
            graph_token, item_id, sheet_name, data_address, session_id,
            horizontal="Center", vertical="Center", wrap_text=True,
        )
        graph_set_range_alignment(
            graph_token, item_id, sheet_name, f"A1:{last_col_letter}1", session_id,
            horizontal="Left", vertical="Center", wrap_text=False,
        )
        graph_set_range_font(
            graph_token, item_id, sheet_name, data_address, session_id,
            name=MAIN_FONT_NAME, size=MAIN_FONT_SIZE,
        )
        graph_set_range_font(
            graph_token, item_id, sheet_name, f"A1:{last_col_letter}1", session_id, bold=True,
        )
        graph_set_range_font(
            graph_token, item_id, sheet_name, f"A{HEADER_ROW}:{last_col_letter}{HEADER_ROW}", session_id,
            color="#FFFFFF", bold=True,
        )

        fills = [
            (f"A{r}:{last_col_letter}{r}", DATE_HEADER_FILL_HEX) for r in layout["date_header_fill_rows"]
        ]
        fills.extend(
            (f"A{r}:{last_col_letter}{r}", fill_hex) for r, fill_hex in layout["data_row_fills"]
        )
        graph_batch_patch_fills(graph_token, item_id, sheet_name, fills, session_id)

        for r in layout["date_header_font_rows"]:
            graph_set_range_font(
                graph_token, item_id, sheet_name, f"A{r}:{last_col_letter}{r}", session_id,
                color="#FFFFFF", bold=True,
            )
            # 出荷数行はA列のみに値が入っているが、Tableは結合セルを含められないため
            # 実際の結合はせず、複数セルの選択範囲内で中央に見せる
            # 「選択範囲内で中央」（CenterAcrossSelection）を使う。
            graph_set_range_alignment(
                graph_token, item_id, sheet_name, f"A{r}:{last_col_letter}{r}", session_id,
                horizontal="CenterAcrossSelection", vertical="Center", wrap_text=False,
            )
            graph_set_range_border_color(
                graph_token, item_id, sheet_name, f"A{r}:{last_col_letter}{r}", session_id, "#FFFFFF",
            )

        graph_batch_set_row_heights(graph_token, item_id, sheet_name, layout["row_heights"], session_id)

        for col_letter, width_pt in MAIN_COLUMN_WIDTH_PT.items():
            graph_set_column_width(
                graph_token, item_id, sheet_name, f"{col_letter}1:{col_letter}1", session_id, width_pt
            )
    finally:
        graph_close_workbook_session(graph_token, item_id, session_id)


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

        color_map = load_case_name_row_color_map(config)

        all_rows = collect_all_rows(config, today, window_end)
        ship_rows = filter_rows_by_ship_window(all_rows, today, window_end)
        ship_rows = sort_rows_by_ship_date(ship_rows, color_map)

        counts: dict[date, int] = {}
        for row in ship_rows:
            counts[row["_出荷日付"]] = counts.get(row["_出荷日付"], 0) + 1
        for d, count in sorted(counts.items()):
            logging.info("%s: %d件", format_date_label(d), count)

        source_counts: dict[str, int] = {}
        for row in ship_rows:
            src = str(row.get("依頼先", ""))
            source_counts[src] = source_counts.get(src, 0) + 1
        logging.info("依頼先別件数(表示範囲内): %s", source_counts)
        for row in ship_rows:
            logging.info(
                "出荷日サマリー対象行: 依頼先=%s 案件No=%s 出荷日=%s 案件名=%s",
                row.get("依頼先", ""), row.get("案件No", ""), row.get("出荷日", ""), row.get("案件名", ""),
            )

        if dry_run:
            logging.info("dry-run のため OneDrive 反映はスキップしました")
            return 0

        graph_token = acquire_token(
            config,
            scopes=GRAPH_SCOPES,
            force_login=force_login,
        )
        header_values = fetch_main_header_row(graph_token, config)
        layout = build_summary_layout(ship_rows, header_values, today, window_end, color_map)
        remote_path = config.get("shipping_summary_path", DEFAULT_OUTPUT_PATH)
        write_summary_via_graph(graph_token, remote_path, layout)
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
