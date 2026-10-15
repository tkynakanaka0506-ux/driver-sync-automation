#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ドライバー情報_営業用.xlsx を毎月デスクトップのフォルダへバックアップする。

タスクスケジューラから月1回自動実行される。
保存先: Desktop/ドライバー情報バックアップ/YYYY-MM/ドライバー情報_YYYY-MM.xlsx
"""

from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path

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

DESKTOP = Path.home() / "Desktop"
BACKUP_FOLDER = DESKTOP / "ドライバー情報バックアップ"


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
    result = None
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        login_hint = config.get("login_hint", "t_nakanishi@showa.co.jp")
        print(f"ブラウザが開きます。{login_hint} でサインインしてください。")
        result = app.acquire_token_interactive(scopes=GRAPH_SCOPES, login_hint=login_hint)
    if not result or "access_token" not in result:
        error = result.get("error_description") or result.get("error") if result else "不明"
        raise RuntimeError(f"認証失敗: {error}")
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


def get_onedrive_file_bytes(token: str, remote_path: str) -> bytes:
    encoded = requests.utils.quote(remote_path, safe="/")
    url = f"{GRAPH_BASE}/me/drive/root:{encoded}:/content"
    res = requests.get(url, headers={"Authorization": f"Bearer {token}"}, allow_redirects=True)
    if not res.ok:
        raise RuntimeError(f"ダウンロード失敗 ({res.status_code}): {res.text[:300]}")
    return res.content


def main() -> int:
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    remote_path: str = config.get("onedrive_output", {}).get(
        "path", "/クロードコード/ドライバー情報_営業用.xlsx"
    )

    today = date.today()
    ym = today.strftime("%Y-%m")
    month_folder = BACKUP_FOLDER / ym
    month_folder.mkdir(parents=True, exist_ok=True)

    filename = f"ドライバー情報_{ym}.xlsx"
    save_path = month_folder / filename

    print(f"[monthly_backup] ダウンロード中: {remote_path}")
    token = acquire_token()
    data = get_onedrive_file_bytes(token, remote_path)
    save_path.write_bytes(data)
    print(f"[monthly_backup] 保存完了: {save_path}  ({len(data):,} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
