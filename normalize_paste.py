#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""【熱源】案件管理シート.xlsx 貼り付け行の正規化スクリプト。

A列にBoolean(TRUE/FALSE)が入っている行を「コピペ直後の行」と判定し、
貼り付け元(A~N列)のデータを正しい列(A~I)に並び替えて書き戻す。
貼り付け後に手動で1回実行する。

貼付元 → 書込先 マッピング:
  A(0) = 納期TRUE/FALSE  → A (TRUE→✅, FALSE→空白)
  B(1) = 案件No.         → B
  C(2) = 納入先名        → C
  D(3) = 空白            → 破棄
  E(4) = 型式            → D
  F(5) = 出荷日          → E
  G(6) = 着日            → F
  H(7) = 月テキスト      → 破棄
  I(8) = ●              → 破棄
  J(9) = ●              → 破棄
  K(10)= ●              → 破棄
  L(11)= 時間            → G
  M(12)= 指定車両        → H
  N(13)= 住所            → I
  O以降                  → 全て空白に
"""

from __future__ import annotations

import base64
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

EXCEL_BASE = datetime(1899, 12, 30)


def serial_to_date_str(val) -> str:
    """Excelの日付シリアル値はそのまま返す（書式はnumberFormatで設定）。"""
    return val


def time_serial_to_str(val) -> str:
    """Excelの時刻シリアル値（0.375 = 9:00）→ "H:MM" 文字列。"""
    if isinstance(val, float) and 0 <= val < 1:
        total_minutes = round(val * 24 * 60)
        h, m = divmod(total_minutes, 60)
        return f"{h}:{m:02d}"
    return val

import msal
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "driver_sync_config.json"
TOKEN_CACHE_PATH = SCRIPT_DIR / "token_cache.bin"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPES = [
    "https://graph.microsoft.com/Files.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]

SHARE_URL = (
    "https://showatekkou-my.sharepoint.com/:x:/r/personal/t_nakanishi_showa_co_jp"
    "/_layouts/15/Doc.aspx?sourcedoc=%7B8730BFF1-E41B-4966-959C-1AC3E69F6BBF%7D"
    "&file=%E3%80%90%E7%86%B1%E6%BA%90%E3%80%91%E6%A1%88%E4%BB%B6%E7%AE%A1%E7%90"
    "%86%E3%82%B7%E3%83%BC%E3%83%88.xlsx&action=default&mobileredirect=true"
)

DATA_START_ROW = 7
CHECK_MARK = "✅"
PASTE_WIDTH = 14  # 貼付元 A~N


def encode_share_url(url: str) -> str:
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{b64}"


def acquire_token() -> str:
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
    app = msal.PublicClientApplication(
        config["client_id"],
        authority=f"https://login.microsoftonline.com/{config['tenant_id']}",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    if not accounts:
        raise RuntimeError(
            "token_cache.bin にアカウントがありません。"
            "python driver_sync.py --login を実行してください。"
        )
    result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise RuntimeError("トークン取得失敗。再ログインが必要です。")
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


def remap(src: list) -> list:
    """貼付元14列 → 正規化後A~N列(14セル)の配列を返す。"""
    # 元の値を安全に取得
    def get(i):
        return src[i] if i < len(src) else ""

    mapped_a_to_i = [
        CHECK_MARK if get(0) is True else ("" if get(0) is False else get(0)),  # A: 納期
        get(1),   # B: 案件No.
        get(2),   # C: 納入先名
        get(4),                       # D: 型式 (src[3]は空白なので飛ばす)
        serial_to_date_str(get(5)),   # E: 出荷日（シリアル→日付文字列）
        serial_to_date_str(get(6)),   # F: 着日（シリアル→日付文字列）
        time_serial_to_str(get(11)),  # G: 時間（時刻シリアル→文字列）
        get(12),  # H: 指定車両
        get(13),  # I: 住所
    ]
    # J以降(src列は使わない)をN列まで空白で埋めて同じ幅に揃える
    return mapped_a_to_i + [""] * (PASTE_WIDTH - len(mapped_a_to_i))


def main() -> int:
    token = acquire_token()
    headers = {"Authorization": f"Bearer {token}"}

    share_id = encode_share_url(SHARE_URL)
    res = requests.get(f"{GRAPH_BASE}/shares/{share_id}/driveItem", headers=headers)
    if not res.ok:
        print("driveItem取得失敗:", res.status_code, res.text[:500])
        return 1
    item = res.json()
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]

    base = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/workbook"
    ws_res = requests.get(f"{base}/worksheets", headers=headers)
    ws_data = ws_res.json()
    if not ws_res.ok or "value" not in ws_data:
        print("worksheets取得失敗:", ws_res.status_code, ws_res.text[:500])
        return 1
    sheet_name = ws_data["value"][0]["name"]

    used_res = requests.get(
        f"{base}/worksheets('{sheet_name}')/usedRange(valuesOnly=true)",
        headers=headers,
    )
    if not used_res.ok:
        print("usedRange取得失敗:", used_res.status_code, used_res.text[:500])
        return 1
    used = used_res.json()
    last_row = used["rowCount"] + used["rowIndex"]

    # A列〜N列を一括取得
    addr = f"A{DATA_START_ROW}:N{last_row}"
    rng_res = requests.get(
        f"{base}/worksheets('{sheet_name}')/range(address='{addr}')",
        headers=headers,
    )
    if not rng_res.ok:
        print("範囲取得失敗:", rng_res.status_code, rng_res.text[:500])
        return 1
    rng = rng_res.json()
    all_values = rng["values"]
    all_types = rng["valueTypes"]

    changed = 0
    for i, (row_vals, row_types) in enumerate(zip(all_values, all_types)):
        row_num = DATA_START_ROW + i
        # A列がBoolean型、または文字列"TRUE"/"FALSE" = チェックボックス由来の貼り付け行と判定
        a_val = row_vals[0]
        a_type = row_types[0]
        is_checkbox_paste = (
            a_type == "Boolean"
            or (a_type == "String" and str(a_val).strip().upper() in ("TRUE", "FALSE"))
        )
        if not is_checkbox_paste:
            continue
        # remap用にBooleanに統一
        if a_type == "String":
            row_vals = list(row_vals)
            row_vals[0] = (str(row_vals[0]).strip().upper() == "TRUE")

        new_row = remap(row_vals)
        # 出荷日(E=index4)・着日(F=index5)をm/d書式、他はGeneral
        number_formats = [["General"] * PASTE_WIDTH]
        number_formats[0][4] = "m/d"
        number_formats[0][5] = "m/d"
        patch_res = requests.patch(
            f"{base}/worksheets('{sheet_name}')/range(address='A{row_num}:N{row_num}')",
            headers=headers,
            json={"values": [new_row], "numberFormat": number_formats},
        )
        if not patch_res.ok:
            print(f"{row_num}行目の書き込み失敗:", patch_res.status_code, patch_res.text[:300])
            return 1
        print(f"{row_num}行目: 正規化完了 → 案件No.={new_row[1]}, 型式={new_row[3]}")
        changed += 1

    if changed == 0:
        print("正規化対象の行はありませんでした（A列にBoolean型のセルなし）。")
    else:
        print(f"\n完了: {changed}行を正規化しました。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
