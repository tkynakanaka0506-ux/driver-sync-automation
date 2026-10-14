#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""毎月デスクトップのフォルダへ Excel をバックアップする。

バックアップ対象:
  1. ドライバー情報_営業用.xlsx  (OneDrive 出力ファイル)
  2. 中通 運行管理シート          (匿名共有リンク)
  3. 松崎運輸 運行管理シート      (匿名共有リンク)

保存先: Desktop/ドライバー情報バックアップ/YYYY-MM/
タスクスケジューラから月1回自動実行される。
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

BROWSER_UA = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

# 匿名共有リンクでバックアップするソースファイル
SOURCE_BACKUPS = [
    {"key": "nakadori", "label": "中通",     "filename": "中通_運行管理_{ym}.xlsx"},
    {"key": "matsuzaki","label": "松崎運輸",  "filename": "松崎運輸_運行管理_{ym}.xlsx"},
]


# ---------------------------------------------------------------------------
# 認証（OneDrive 出力ファイル用）
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# ダウンロード関数
# ---------------------------------------------------------------------------

def get_onedrive_file_bytes(token: str, remote_path: str) -> bytes:
    """自分の OneDrive から指定パスのファイルをダウンロード。"""
    encoded = requests.utils.quote(remote_path, safe="/")
    url = f"{GRAPH_BASE}/me/drive/root:{encoded}:/content"
    res = requests.get(url, headers={"Authorization": f"Bearer {token}"}, allow_redirects=True)
    if not res.ok:
        raise RuntimeError(f"ダウンロード失敗 ({res.status_code}): {res.text[:300]}")
    return res.content


def download_share_file(share_url: str) -> bytes:
    """匿名共有リンク(1drv.ms)から認証なしでダウンロード。"""
    sep = "&" if "?" in share_url else "?"
    url = f"{share_url}{sep}download=1"
    res = requests.get(url, headers=BROWSER_UA, timeout=180, allow_redirects=True)
    if not res.ok:
        raise RuntimeError(f"ファイル取得失敗: {res.status_code} {res.text[:200]}")
    if res.content[:4] != b"PK\x03\x04":
        raise RuntimeError(
            f"取得データがExcel(xlsx)ではありません。先頭={res.content[:8]!r}"
        )
    return res.content


# ---------------------------------------------------------------------------
# メイン
# ---------------------------------------------------------------------------

def main() -> int:
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    today = date.today()
    ym = today.strftime("%Y-%m")

    # 月別サブフォルダ作成: Desktop/ドライバー情報バックアップ/2026-10/
    month_folder = BACKUP_FOLDER / ym
    month_folder.mkdir(parents=True, exist_ok=True)
    print(f"保存先フォルダ: {month_folder}\n")

    errors: list[str] = []

    # --- 1. ドライバー情報_営業用.xlsx (OneDrive) ---
    remote_path: str = config.get("onedrive_output", {}).get(
        "path", "/クロードコード/ドライバー情報_営業用.xlsx"
    )
    filename = f"ドライバー情報_{ym}.xlsx"
    save_path = month_folder / filename
    try:
        print(f"[1/3] ドライバー情報_営業用.xlsx をダウンロード中...")
        token = acquire_token()
        data = get_onedrive_file_bytes(token, remote_path)
        save_path.write_bytes(data)
        print(f"      保存完了: {filename}  ({len(data):,} bytes)")
    except Exception as e:
        msg = f"ドライバー情報_営業用: {e}"
        print(f"      エラー: {msg}")
        errors.append(msg)

    # --- 2 & 3. 中通・松崎運輸（匿名共有リンク） ---
    sources_by_key = {s["key"]: s for s in config.get("sources", [])}
    for i, target in enumerate(SOURCE_BACKUPS, start=2):
        key = target["key"]
        label = target["label"]
        filename = target["filename"].format(ym=ym)
        save_path = month_folder / filename
        src = sources_by_key.get(key)
        if not src:
            msg = f"{label}: config に share_url が見つかりません"
            print(f"[{i}/3] {msg}")
            errors.append(msg)
            continue
        try:
            print(f"[{i}/3] {label} をダウンロード中...")
            data = download_share_file(src["share_url"])
            save_path.write_bytes(data)
            print(f"      保存完了: {filename}  ({len(data):,} bytes)")
        except Exception as e:
            msg = f"{label}: {e}"
            print(f"      エラー: {msg}")
            errors.append(msg)

    print()
    if errors:
        print(f"完了（{len(errors)}件エラー）:")
        for err in errors:
            print(f"  - {err}")
        return 1
    print("全て完了。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
