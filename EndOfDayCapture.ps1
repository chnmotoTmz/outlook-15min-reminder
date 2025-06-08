# EndOfDayCapture.ps1 - 終業時自動記録スクリプト
# 実行方法: powershell.exe -ExecutionPolicy Bypass -File "EndOfDayCapture.ps1"

# 共通ライブラリの読み込み
$taskTrackerPath = Join-Path $PSScriptRoot "TaskTracker.ps1"
if (Test-Path $taskTrackerPath) {
    . $taskTrackerPath
} else {
    Write-Error "TaskTracker.ps1が見つかりません: $taskTrackerPath"
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Media

try {
    Write-Host "=== 終業時記録システム開始 ==="
    Write-Host "時刻: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
    
    # 現在の作業コンテキストを取得
    Write-Host "作業コンテキストを取得中..."
    $workContext = Get-WorkContext
    
    # 作業コンテキストを保存
    Save-WorkContext $workContext
    
    # 通知音を鳴らす
    [System.Media.SystemSounds]::Asterisk.Play()
    
    # 昨日の未完了タスクを確認
    $yesterdayTasks = Get-YesterdaysTasks
    $incompleteTasks = $yesterdayTasks | Where-Object { $_.Status -eq "未完了" }
    
    # デフォルトタスクを準備
    $defaultTask = ""
    if ($incompleteTasks.Count -gt 0) {
        $defaultTask = "✅ 昨日の継続:`n"
        foreach ($task in $incompleteTasks) {
            $defaultTask += "- $($task.Content)`n"
        }
        $defaultTask += "`n新しいタスク:"
    } else {
        # 最後のOutlook予定がある場合はそれをヒントとして表示
        if ($workContext.lastOutlookAppointment) {
            $lastMeeting = $workContext.lastOutlookAppointment
            $defaultTask = "✅ 最後の予定: $($lastMeeting.subject)`n明日への継続作業:"
        }
    }
    
    # 情報メッセージを作成
    $infoMessage = @"
🕚 本日の作業お疑れ様でした！

📊 今日の作業状況:
- アクティブなアプリ: $($workContext.activeApps.Count)個
- 最後の予定: $(
    if ($workContext.lastOutlookAppointment) {
        "$($workContext.lastOutlookAppointment.subject) ($($workContext.lastOutlookAppointment.start))"
    } else {
        "予定なし"
    }
)

📋 明日のタスクを設定しましょう。
朝、スムーズに作業を開始できます！
"@
    
    # 情報表示
    $result = [System.Windows.Forms.MessageBox]::Show(
        $infoMessage, 
        "🌅 終業時タスク記録", 
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # タスク入力フォームを表示
        $inputResult = Show-TaskInputForm -defaultTask $defaultTask
        
        if ($inputResult) {
            $successMessage = @"
✅ 明日のタスクを記録しました！

🌅 明日の朝、スムーズに作業を開始できます。
🚀 良い一日をお過ごしください！
"@
            
            [System.Windows.Forms.MessageBox]::Show(
                $successMessage, 
                "✨ 記録完了", 
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            [System.Media.SystemSounds]::Exclamation.Play()
        }
    } else {
        Write-Host "ユーザーがキャンセルしました。"
    }
    
    # 結果レポート
    Write-Host "=== 結果レポート ==="
    Write-Host "作業コンテキスト保存: 完了"
    Write-Host "アクティブアプリ数: $($workContext.activeApps.Count)"
    Write-Host "終業時記録システム終了: $(Get-Date -Format 'HH:mm:ss')"
    
} catch {
    $errorMessage = "エラーが発生しました:`n$($_.Exception.Message)"
    Write-Host $errorMessage
    
    # エラーもメッセージボックスで表示
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage, 
        "終業時記録エラー", 
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
