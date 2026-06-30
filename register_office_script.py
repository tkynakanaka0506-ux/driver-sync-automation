#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office Script を Graph API beta で登録するスクリプト。
"""
import base64
import json
import sys
from pathlib import Path

import msal
import requests

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = SCRIPT_DIR / "driver_sync_config.json"
TOKEN_CACHE_PATH = SCRIPT_DIR / "token_cache.bin"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"
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

SCRIPT_CODE = r"""function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getFirstWorksheet();
  const DATA_START_ROW = 6;
  const PASTE_COLS = 14;

  const usedRange = sheet.getUsedRange();
  if (!usedRange) return;

  const lastRow = usedRange.getRowIndex() + usedRange.getRowCount();
  if (lastRow <= DATA_START_ROW) return;

  const rowCount = lastRow - DATA_START_ROW;
  const range = sheet.getRangeByIndexes(DATA_START_ROW, 0, rowCount, PASTE_COLS);
  const values = range.getValues();
  const types = range.getValueTypes();

  let changed = 0;

  for (let i = 0; i < values.length; i++) {
    const aType = types[i][0];
    const aVal = values[i][0];

    const isCheckboxPaste =
      aType === ExcelScript.RangeValueType.boolean ||
      (aType === ExcelScript.RangeValueType.string &&
        String(aVal).trim().toUpperCase() === "TRUE");

    if (!isCheckboxPaste) continue;

    const src = values[i];
    const checkmark =
      aVal === true || String(aVal).trim().toUpperCase() === "TRUE" ? "✅" : "";

    const remapped: (string | number | boolean)[] = [
      checkmark, src[1], src[2], src[4], src[5], src[6],
      src[11], src[12], src[13],
      "", "", "", "", ""
    ];

    sheet.getRangeByIndexes(DATA_START_ROW + i, 0, 1, PASTE_COLS).setValues([remapped]);
    changed++;
  }

  console.log(`normalize_paste: ${changed}行を正規化しました。`);
}"""


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
        raise RuntimeError("token_cache.bin にアカウントがありません。")
    result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise RuntimeError("トークン取得失敗。")
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


def main() -> int:
    token = acquire_token()
    headers = {"Authorization": f"Bearer {token}"}

    # driveItem を取得
    share_id = encode_share_url(SHARE_URL)
    res = requests.get(f"{GRAPH_BASE}/shares/{share_id}/driveItem", headers=headers)
    if not res.ok:
        print("driveItem取得失敗:", res.status_code, res.text[:300])
        return 1
    item = res.json()
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    print(f"driveId: {drive_id}")
    print(f"itemId: {item_id}")

    # Office Scripts の一覧を取得（beta）
    scripts_url = f"{GRAPH_BETA}/drives/{drive_id}/items/{item_id}/workbook/scripts"
    res = requests.get(scripts_url, headers=headers)
    print(f"\nGET scripts → {res.status_code}")
    if res.ok:
        scripts = res.json().get("value", [])
        print(f"登録済みスクリプト数: {len(scripts)}")
        for s in scripts:
            print(f"  - {s.get('name')} (id: {s.get('id')})")
    else:
        print("レスポンス:", res.text[:500])
        print("\nOffice Scripts APIはGraph API beta経由では利用できない可能性があります。")
        print("Excelのウェブ版から手動で登録してください。")
        return 1

    # normalize_paste が既に存在するか確認
    existing = next((s for s in scripts if s.get("name") == "normalize_paste"), None)
    if existing:
        print(f"\n'normalize_paste' は既に登録されています (id: {existing['id']})")
        # コードを更新
        script_id = existing["id"]
        patch_url = f"{scripts_url}/{script_id}/content"
        res = requests.put(
            patch_url,
            headers={**headers, "Content-Type": "text/plain"},
            data=SCRIPT_CODE.encode("utf-8"),
        )
        print(f"コード更新 → {res.status_code}")
        if res.ok:
            print("normalize_paste スクリプトを更新しました。")
        else:
            print("更新失敗:", res.text[:300])
    else:
        # 新規作成
        res = requests.post(
            scripts_url,
            headers=headers,
            json={"name": "normalize_paste"},
        )
        print(f"\nスクリプト作成 → {res.status_code}")
        if not res.ok:
            print("作成失敗:", res.text[:300])
            return 1
        script_id = res.json().get("id")
        print(f"作成完了 id: {script_id}")

        # コードをアップロード
        content_url = f"{scripts_url}/{script_id}/content"
        res = requests.put(
            content_url,
            headers={**headers, "Content-Type": "text/plain"},
            data=SCRIPT_CODE.encode("utf-8"),
        )
        print(f"コードアップロード → {res.status_code}")
        if res.ok:
            print("normalize_paste スクリプトの登録が完了しました！")
        else:
            print("コードアップロード失敗:", res.text[:300])
            return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
