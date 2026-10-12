# monthly_backup.py をタスクスケジューラに毎月1日 9:00 で登録するスクリプト。
# 管理者権限不要。PowerShell から一度だけ実行してください。

$TaskName  = "DriverInfoMonthlyBackup"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe = (Get-Command python -ErrorAction SilentlyContinue).Source

if (-not $PythonExe) {
    Write-Error "python.exe が見つかりません。Python をインストールしてPATHに追加してください。"
    exit 1
}

$ScriptPath = Join-Path $ScriptDir "monthly_backup.py"
if (-not (Test-Path $ScriptPath)) {
    Write-Error "monthly_backup.py が見つかりません: $ScriptPath"
    exit 1
}

# 毎月1日 09:00 トリガー
$Trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At "09:00"

# python monthly_backup.py を ScriptDir をカレントディレクトリにして実行
$Action  = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "`"$ScriptPath`"" `
    -WorkingDirectory $ScriptDir

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable `   # PC が落ちていた場合、次回起動時に実行
    -RunOnlyIfNetworkAvailable

# 既存タスクがあれば上書き
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "既存タスクを削除しました: $TaskName"
}

Register-ScheduledTask `
    -TaskName $TaskName `
    -Trigger   $Trigger `
    -Action    $Action `
    -Settings  $Settings `
    -Description "ドライバー情報_営業用.xlsx を毎月1日にデスクトップへバックアップ" `
    -RunLevel Normal

Write-Host ""
Write-Host "登録完了: $TaskName"
Write-Host "  実行日時: 毎月1日 09:00（PC 起動後に実行される設定あり）"
Write-Host "  保存先  : $env:USERPROFILE\Desktop\ドライバー情報バックアップ\"
