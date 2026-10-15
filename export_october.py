#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""10/1〜10/31 の出荷日データを全ソースから抽出してデスクトップに保存する。

使い方: py -3 export_october.py
出力:   Desktop/ドライバー情報バックアップ/ドライバー情報_2026-10_全件.xlsx
"""

from __future__ import annotations

import json
import sys
from datetime import date
from io import BytesIO
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# driver_sync.py と同じディレクトリにあるので直接インポート
sys.path.insert(0, str(Path(__file__).resolve().parent))
from driver_sync import (  # noqa: E402
    download_share_file,
    extract_rows_from_workbook,
    SOURCE_ORDER,
    SOURCE_DISPLAY_NAMES,
)

CONFIG_PATH = Path(__file__).resolve().parent / "driver_sync_config.json"
DESKTOP = Path.home() / "Desktop"
BACKUP_FOLDER = DESKTOP / "ドライバー情報バックアップ"

EXPORT_START = date(2026, 10, 1)
EXPORT_END   = date(2026, 10, 31)

COLUMNS = ["出荷日", "着日", "案件No", "案件名", "型式", "車型", "会社名", "乗務員", "車番", "携帯番号", "備考", "依頼先"]

SOURCE_COLORS = {
    "matsuzaki": "D9EAD3",
    "nakadori":  "CFE2F3",
    "fukuoka":   "FFF2CC",
    "maruun":    "FCE5CD",
    "ks":        "EAD1DC",
}


def main() -> int:
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    sources = {s["key"]: s["share_url"] for s in config.get("sources", [])}

    all_rows: list[dict] = []

    for key in SOURCE_ORDER:
        share_url = sources.get(key)
        if not share_url:
            print(f"[{key}] share_url なし — スキップ")
            continue
        print(f"[{key}] ダウンロード中...")
        try:
            content = download_share_file(share_url)
        except Exception as e:
            print(f"[{key}] 取得失敗: {e}")
            continue

        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
        # 出荷日でフィルタ（着日ではなく出荷日基準）
        rows = extract_rows_from_workbook(
            wb,
            key,
            config.get("header_rows", 4),
            EXPORT_START,                       # today の代わりに開始日を渡す
            arr_days_back=0,
            arr_window_end_override=EXPORT_END,
            window_date_field="ship",           # 出荷日で絞り込み
            require_driver_info=False,
        )
        for row in rows:
            row["依頼先_key"] = key
        all_rows.extend(rows)
        print(f"  → {len(rows)} 件")

    if not all_rows:
        print("抽出データなし。")
        return 0

    # 出荷日でソート
    all_rows.sort(key=lambda r: (r.get("出荷日") or "", r.get("依頼先_key") or ""))

    # Excel 出力
    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = "10月全件"

    # ヘッダー
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font_white = Font(bold=True, color="FFFFFF")
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # データ行
    for row in all_rows:
        key = row.get("依頼先_key", "")
        color = SOURCE_COLORS.get(key, "FFFFFF")
        fill = PatternFill("solid", fgColor=color)
        values = [
            row.get("出荷日", ""),
            row.get("着日", ""),
            row.get("案件No", ""),
            row.get("案件名", ""),
            row.get("型式", ""),
            row.get("車型", ""),
            row.get("会社名", ""),
            row.get("乗務員", ""),
            row.get("車番", ""),
            row.get("携帯番号", ""),
            row.get("備考", ""),
            SOURCE_DISPLAY_NAMES.get(key, key),
        ]
        ws.append(values)
        for cell in ws[ws.max_row]:
            cell.fill = fill
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    # 列幅調整
    col_widths = [10, 10, 14, 28, 20, 12, 18, 12, 12, 16, 20, 12]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width
    ws.row_dimensions[1].height = 20

    BACKUP_FOLDER.mkdir(parents=True, exist_ok=True)
    out_path = BACKUP_FOLDER / "ドライバー情報_2026-10_全件.xlsx"
    out_wb.save(out_path)
    print(f"\n保存完了: {out_path}  (合計 {len(all_rows)} 件)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
