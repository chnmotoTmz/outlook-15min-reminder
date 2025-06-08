# Test.ps1 - 手動テスト用スクリプト
# Outlookに予定がある状態でテストしてください

Write-Host "=== Outlook 15分前アラーム テスト ===" -ForegroundColor Green

# 現在の時刻を表示
$now = Get-Date
Write-Host "現在時刻: $($now.ToString('yyyy/MM/dd HH:mm:ss'))" -ForegroundColor Yellow

# テスト予定の作成を提案
Write-Host "`n📝 テスト手順:" -ForegroundColor Cyan
Write-Host "1. Outlookを開いてください"
Write-Host "2. 現在時刻から15分後に予定を作成してください"
Write-Host "   推奨時刻: $($now.AddMinutes(15).ToString('HH:mm'))"
Write-Host "3. 予定を保存してください"
Write-Host "4. Enterキーを押してテストを開始してください"

Read-Host "`nEnterキーを押してテスト開始"

# メインスクリプトを実行
Write-Host "`n🔍 アラームスクリプトを実行中..." -ForegroundColor Yellow
try {
    $scriptPath = ".\OutlookReminder.ps1"
    if (Test-Path $scriptPath) {
        & $scriptPath
    } else {
        Write-Host "❌ OutlookReminder.ps1が見つかりません" -ForegroundColor Red
        Write-Host "現在のディレクトリ: $(Get-Location)"
    }
} catch {
    Write-Host "❌ スクリプト実行エラー: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n✅ テスト完了" -ForegroundColor Green
Write-Host "メッセージボックスが表示されましたか？" -ForegroundColor Yellow
