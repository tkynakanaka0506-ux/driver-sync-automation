# driver-sync-automation

OneDrive上の4社（matsuzaki/松崎運輸, nakadori/中通, fukuoka/福岡ロジテック, maruun/丸運）の
車両依頼書・管理シートから案件・ドライバー情報を集計し、社内の「ドライバー情報_営業用.xlsm」
（OneDrive上）へ書き込む自動同期システム。

## 仕組みの全体像

```
cron-job.org (15分ごとHTTP POST)
  → GitHub Actions (.github/workflows/sync.yml) を workflow_dispatch で起動
    → 一時的なLinux仮想マシンが起動
      → Secrets「TOKEN_CACHE_B64」を復号してtoken_cache.binを復元
      → python driver_sync.py を実行
          - Microsoft Graph APIでMSAL認証（トークンはtoken_cache.binから無人取得）
          - 4社のOneDrive上Excelを読み取り、SHEET_CONFIG（driver_sync.py内）の列定義で抽出
          - 出荷日が「過去7日〜翌日」の案件のみに絞る
          - 備考列に号車/着時間、中継ありなら2次配送情報を追記
          - ドライバー4項目未入力などの異常を警告として収集
          - 「ドライバー情報_営業用.xlsm」へGraph API PATCHで一括書き込み
          - 行の色分け（HIGHLIGHT_CASE_NAMES）、E1最終更新時刻（JST固定）、
            E2〜E4要確認アラート、F2:H4確認済み案件No除外を反映
      → 更新後のtoken_cache.binを再度暗号化してSecretsへ書き戻す（リフレッシュトークンローテーション対応）
      → 仮想マシンは破棄される
```

## なぜこの構成になっているか

- **GitHub Actions単体のcronは精度が低い**（無料枠は実行が1〜2時間ずれることがある）ため、
  外部の cron-job.org から `workflow_dispatch` をHTTP POSTで15分ごとに叩いて確実に起動させている。
- **ローカルのWindowsタスクスケジューラ「ドライバー情報反映」は無効化済み**
  （クラウドと同時に走るとOneDriveファイルロック競合が起きるため）。
- **認証情報はGitHub Secretsで暗号化保管**（`TOKEN_CACHE_B64`）。Azure ADはpublicクライアントの
  リフレッシュトークンを使うたびにローテーションするため、実行後に必ず書き戻す。
- E1のタイムスタンプは `datetime.now(JST)` で固定（GitHub Actionsランナーは UTC のため、
  素の `datetime.now()` を使うと表示が9時間ずれる）。

## ハマりやすい点（実際に起きた事故）

- **PowerShellのパイプでBOMが混入する**: `$str | gh secret set NAME` のように文字列を直接
  パイプすると先頭に1文字BOMが混入し、`base64: invalid input` で失敗する。
  `[System.Text.Encoding]::UTF8` (BOM無し) でファイルに書き出し、`cmd /c "... < file"` で
  リダイレクトする方式が安全。
- **cron-job.orgのヘッダー欄に複数ヘッダーを1つのValueへ誤って貼ると401になる**:
  Authorization の Value は `Bearer <token>` の1行のみにすること。
- **案件Noを数字のみで入力すると無視リスト（F2:H4）が効かない**: Excelが数値型(`1440.0`)として
  保存し、警告文字列の `"1440"` と一致しなくなるため。`driver_sync.py` の
  `graph_get_ignored_case_numbers` で整数化して吸収済み（修正済み、再発した場合はここを疑う）。
- **`token_cache.bin` はGit管理外**（`.gitignore`）。ローカルで再生成・再アップロードする際は
  `driver_sync_config.json` の認証情報（client_id/tenant_id）と一致しているか確認。
- **ファイル形式（拡張子）を変えると共有リンクが必ず変わる**: VBAマクロは`.xlsm`でしか保存できない
  というExcel自体の制限があり、`.xlsx`→`.xlsm`変換は新しいOneDriveアイテムとして作られるため、
  既存の匿名共有リンクは無効になり新しいリンクが発行される（復元不可・仕様上の制限）。
  ファイル形式を変える作業は、リンクの再共有が必要になることを前提に進めること。

## 携帯番号の保護（マスク表示＋VBAボタン）

`ドライバー情報_営業用.xlsm`（マクロ有効ブック）で運用中。

- K列（携帯番号）は常に表示されるが、値は`***`にマスクされる（`MASKED_COLUMN_NAMES`）
- 本物の番号はM列（`携帯番号秘`、列幅0で非表示）に保存
- シートはパスワード保護（`options.allowFormatColumns: false`等）。パスワードは
  環境変数`DRIVER_SYNC_SHEET_PASWORD`→ローカルファイル`sheet_password.txt`（Git管理外）の順で読む。
  GitHub Actionsでは`SHEET_PASSWORD`シークレットを`DRIVER_SYNC_SHEET_PASSWORD`として渡している。
- `driver_sync.py`は毎回 unprotect → 書き込み → M列非表示+K列幅復元 → reprotect を繰り返す
  （`graph_unprotect_worksheet` / `graph_protect_worksheet` / `graph_set_column_width`）。
- シート上に「携帯番号を表示」ボタン（VBA `ShowMobileNumbers`マクロ）があり、パスワードを
  入力するとM列を一時的に再表示できる。次回の自動同期（15分以内）で自動的に再び非表示に戻る。
- VBAプロジェクトの追加はGraph APIでは不可なため、ローカルでExcel COM自動化
  （`win32com.client`、要: Excelの「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」設定）
  を使って一度だけ手動で組み込んだ。再度VBAを変更する場合も同じ手順が必要。

## 関連シークレット（このリポジトリのGitHub Secrets）

- `TOKEN_CACHE_B64`: MSALトークンキャッシュのbase64（driver_sync.py実行用、自動更新される）
- `SECRETS_PAT`: ワークフロー内で`TOKEN_CACHE_B64`を書き戻すための fine-grained PAT
  （このリポジトリのみ、Secrets: Read and write）
- `SHEET_PASSWORD`: 携帯番号列のシート保護パスワード（`DRIVER_SYNC_SHEET_PASSWORD`として渡す）

## 外部サービス側の設定（cron-job.org、コード管理外）

- ジョブ名: `driver-sync`
- URL: `https://api.github.com/repos/<owner>/driver-sync-automation/actions/workflows/sync.yml/dispatches`
- Method: POST, 15分ごと
- Headers: `Authorization: Bearer <cron-trigger PAT、Actions: Read and write権限のみ>`,
  `Accept: application/vnd.github+json`, `X-GitHub-Api-Version: 2022-11-28`,
  `Content-Type: application/json`
- Body: `{"ref":"main"}`
- この `cron-trigger` PATは `SECRETS_PAT` とは別物（最小権限の原則で分離）。

## ローカル開発との関係

このリポジトリは `C:\Users\1229\Desktop\AI関連－仮保存フォルダ\driver_sync.py` のコピー。
ローカル側を直して動作確認したら、必ずこちらにもコピーしてpushすること
（クラウド側はこのリポジトリの内容だけを実行する）。`shipping_summary.py` も同様。

## 出荷日別案件サマリー（shipping_summary.py）

`driver_sync.py` と同じ4社の元データを再集計し、OneDrive上の別Excel
`/ドライバー情報/出荷日別案件サマリー.xlsx`（シート名: 出荷日サマリー）へ
1時間ごとに「出荷日・件数・案件一覧」を書き込む。`driver_sync.py` の抽出関数・
Graph API認証・token_cache.bin をそのまま再利用している（認証情報は共有）。

- 表示範囲: 当日から3日先まで（`driver_sync_config.json` の
  `shipping_summary_days_ahead` で変更可、未設定時は3）。
- 当日が金曜日の場合は土日を挟むため火曜まで延長。延長後の範囲内に祝日が
  あれば、その祝日の翌日まで再延長する（連休にも対応、`ship_window_end()`）。
- 出荷日の判定は `driver_sync.py` の着日フィルタとは独立（着日範囲は広く
  取った上で、出荷日側だけで表示範囲を絞り込む）。
- 出力レイアウトは営業用Excelと同じ列構成（A〜G = 案件No/案件名/出荷日/着日/
  備考/型式/車型）。見出し行（6行目）は営業用ExcelのA6:G6をGraph APIで読んで
  そのままコピーし、7行目から出荷日昇順でデータ行を書き込む
  （`fetch_main_header_row()` / `build_summary_xlsx_bytes()`）。
  件数は日付ごとの行数で読み取れる（個別の件数列は持たない）。
- 実行トリガー: `.github/workflows/shipping_summary.yml`（GitHub Actions
  `schedule: "0 * * * *"` + `workflow_dispatch`）。GitHub Actions単体のcronは
  最大1〜2時間遅延することがある（[[driver-sync-incidents]]参照）ため、
  既存と同様に外部cronサービス（cron-job.org等）から1時間ごとに
  `workflow_dispatch` をHTTP POSTで叩く設定を別途追加すること
  （URL: `.../actions/workflows/shipping_summary.yml/dispatches`、
  ヘッダー・PATは既存の`driver-sync`ジョブと同じものを使い回せる）。
