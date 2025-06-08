# Setup.ps1 - 初期設定用スクリプト
# 管理者権限で実行してください

Write-Host "=== Outlook 15分前アラーム 初期設定 ===" -ForegroundColor Green

# 1. 実行ポリシーの設定
Write-Host "`n1. PowerShell実行ポリシーを設定中..." -ForegroundColor Yellow
try {
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    Write-Host "✅ 実行ポリシーを設定しました" -ForegroundColor Green
} catch {
    Write-Host "❌ 実行ポリシーの設定に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
}

# 2. スクリプトディレクトリの作成
$scriptDir = "C:\Scripts"
Write-Host "`n2. スクリプトディレクトリを作成中..." -ForegroundColor Yellow
try {
    if (-not (Test-Path $scriptDir)) {
        New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null
        Write-Host "✅ ディレクトリを作成しました: $scriptDir" -ForegroundColor Green
    } else {
        Write-Host "✅ ディレクトリは既に存在します: $scriptDir" -ForegroundColor Green
    }
} catch {
    Write-Host "❌ ディレクトリの作成に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
}

# 3. メインスクリプトのコピー
$sourceScript = ".\OutlookReminder.ps1"
$targetScript = "$scriptDir\OutlookReminder.ps1"
Write-Host "`n3. スクリプトファイルをコピー中..." -ForegroundColor Yellow
try {
    if (Test-Path $sourceScript) {
        Copy-Item $sourceScript $targetScript -Force
        Write-Host "✅ スクリプトをコピーしました: $targetScript" -ForegroundColor Green
    } else {
        Write-Host "❌ 元のスクリプトファイルが見つかりません: $sourceScript" -ForegroundColor Red
    }
} catch {
    Write-Host "❌ スクリプトのコピーに失敗しました: $($_.Exception.Message)" -ForegroundColor Red
}

# 4. テスト実行
Write-Host "`n4. テスト実行中..." -ForegroundColor Yellow
try {
    & $targetScript
    Write-Host "✅ テスト実行完了" -ForegroundColor Green
} catch {
    Write-Host "❌ テスト実行に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== 次のステップ ===" -ForegroundColor Cyan
Write-Host "1. タスクスケジューラを開いてください (Win+R → taskschd.msc)"
Write-Host "2. 基本タスクの作成で以下を設定:"
Write-Host "   - 名前: OutlookReminder"
Write-Host "   - トリガー: 毎日"
Write-Host "   - 開始時刻: 08:00 (お好みの時刻)"
Write-Host "   - 繰り返し: 5分間隔、12時間継続"
Write-Host "   - プログラム: powershell.exe"
Write-Host "   - 引数: -ExecutionPolicy Bypass -File `"C:\Scripts\OutlookReminder.ps1`""

Write-Host "`n設定完了です！" -ForegroundColor Green
