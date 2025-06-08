# TestTaskSystem.ps1 - タスク管理システムテスト
# 実行方法: powershell.exe -ExecutionPolicy Bypass -File "TestTaskSystem.ps1"

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

Write-Host "=== タスク管理システム テスト ===" -ForegroundColor Cyan
Write-Host "時刻: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"

try {
    # テストメニューの作成
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "🧪 タスク管理システム テスト"
    $form.Size = New-Object System.Drawing.Size(500, 600)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    
    # タイトル
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "🧪 タスク管理システム テスト"
    $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Font = New-Object System.Drawing.Font("メイリオ", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($titleLabel)
    
    # 説明
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "各機能をテストして動作を確認してください。"
    $descLabel.Size = New-Object System.Drawing.Size(450, 20)
    $descLabel.Location = New-Object System.Drawing.Point(20, 60)
    $descLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($descLabel)
    
    # ボタン配置用変数
    $buttonY = 100
    $buttonHeight = 40
    $buttonWidth = 400
    $buttonSpacing = 50
    
    # ボタン1: 作業コンテキスト取得テスト
    $contextButton = New-Object System.Windows.Forms.Button
    $contextButton.Text = "📁 作業コンテキスト取得テスト"
    $contextButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $contextButton.Location = New-Object System.Drawing.Point(50, $buttonY)
    $contextButton.BackColor = [System.Drawing.Color]::LightBlue
    $form.Controls.Add($contextButton)
    
    $contextButton.Add_Click({
        Write-Host "作業コンテキスト取得テスト実行中..."
        $context = Get-WorkContext
        
        $contextInfo = @"
📁 作業コンテキスト情報:

• 取得時刻: $($context.timestamp)
• アクティブアプリ数: $($context.activeApps.Count)
• Outlook最新予定: $(
    if ($context.lastOutlookAppointment) {
        "$($context.lastOutlookAppointment.subject) ($($context.lastOutlookAppointment.start))"
    } else {
        "予定なし"
    }
)

🖥️ アクティブアプリ一覧:
$(
    if ($context.activeApps.Count -gt 0) {
        $appList = ""
        $displayApps = $context.activeApps | Select-Object -First 10
        foreach ($app in $displayApps) {
            $appList += "- $($app.name): $($app.title)`n"
        }
        if ($context.activeApps.Count -gt 10) {
            $appList += "... 他 $($context.activeApps.Count - 10) 件"
        }
        $appList
    } else {
        "アクティブアプリなし"
    }
)
"@
        
        Show-InfoDisplay -title "作業コンテキスト取得結果" -content $contextInfo -width 600 -height 500
        Write-Host "作業コンテキスト取得テスト完了"
    })
    
    # ボタン2: タスク入力テスト
    $buttonY += $buttonSpacing
    $taskInputButton = New-Object System.Windows.Forms.Button
    $taskInputButton.Text = "📋 タスク入力フォームテスト"
    $taskInputButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $taskInputButton.Location = New-Object System.Drawing.Point(50, $buttonY)
    $taskInputButton.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($taskInputButton)
    
    $taskInputButton.Add_Click({
        Write-Host "タスク入力フォームテスト実行中..."
        $testTask = "テストタスク: $(Get-Date -Format 'HH:mm:ss') に作成"
        $result = Show-TaskInputForm -defaultTask $testTask
        
        if ($result) {
            [System.Windows.Forms.MessageBox]::Show(
                "✅ タスク入力テスト成功！", 
                "テスト結果", 
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "⚠ タスク入力がキャンセルされました。", 
                "テスト結果", 
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
        Write-Host "タスク入力テスト完了"
    })
    
    # ボタン3: 今日のタスク表示テスト
    $buttonY += $buttonSpacing
    $todayTasksButton = New-Object System.Windows.Forms.Button
    $todayTasksButton.Text = "📅 今日のタスク表示テスト"
    $todayTasksButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $todayTasksButton.Location = New-Object System.Drawing.Point(50, $buttonY)
    $todayTasksButton.BackColor = [System.Drawing.Color]::LightYellow
    $form.Controls.Add($todayTasksButton)
    
    $todayTasksButton.Add_Click({
        Write-Host "今日のタスク表示テスト実行中..."
        $todayTasks = Get-TodaysTasks
        
        $taskInfo = @"
📅 今日のタスク一覧:

タスク数: $($todayTasks.Count) 件

$(
    if ($todayTasks.Count -gt 0) {
        $taskList = ""
        foreach ($task in $todayTasks) {
            $priorityIcon = switch ($task.Priority) {
                "高" { "🔴" }
                "中" { "🟡" }
                "低" { "🔵" }
                default { "⚪" }
            }
            $continueIcon = if ($task.IsContinuation -eq "True") { "↪️" } else { "✨" }
            $taskList += "$priorityIcon $continueIcon $($task.Content)`n"
            $taskList += "  作成日時: $($task.Date) $($task.Time)`n"
            $taskList += "  状態: $($task.Status)`n`n"
        }
        $taskList
    } else {
        "今日のタスクはまだありません。`n上の「タスク入力フォームテスト」でタスクを作成してみてください。"
    }
)
"@
        
        Show-InfoDisplay -title "今日のタスク一覧" -content $taskInfo -width 600 -height 400
        Write-Host "今日のタスク表示テスト完了"
    })
    
    # ボタン4: 終業時シミュレーション
    $buttonY += $buttonSpacing
    $endOfDayButton = New-Object System.Windows.Forms.Button
    $endOfDayButton.Text = "🌅 終業時シミュレーション"
    $endOfDayButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $endOfDayButton.Location = New-Object System.Drawing.Point(50, $buttonY)
    $endOfDayButton.BackColor = [System.Drawing.Color]::LightCoral
    $form.Controls.Add($endOfDayButton)
    
    $endOfDayButton.Add_Click({
        Write-Host "終業時シミュレーション実行中..."
        $endOfDayScript = Join-Path $PSScriptRoot "EndOfDayCapture.ps1"
        
        if (Test-Path $endOfDayScript) {
            & $endOfDayScript
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "EndOfDayCapture.ps1 が見つかりません。`nパス: $endOfDayScript", 
                "ファイルエラー", 
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        Write-Host "終業時シミュレーション完了"
    })
    
    # ボタン5: 始業時シミュレーション
    $buttonY += $buttonSpacing
    $startOfDayButton = New-Object System.Windows.Forms.Button
    $startOfDayButton.Text = "🌅 始業時シミュレーション"
    $startOfDayButton.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
    $startOfDayButton.Location = New-Object System.Drawing.Point(50, $buttonY)
    $startOfDayButton.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($startOfDayButton)
    
    $startOfDayButton.Add_Click({
        Write-Host "始業時シミュレーション実行中..."
        $startOfDayScript = Join-Path $PSScriptRoot "StartOfDayReminder.ps1"
        
        if (Test-Path $startOfDayScript) {
            & $startOfDayScript
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "StartOfDayReminder.ps1 が見つかりません。`nパス: $startOfDayScript", 
                "ファイルエラー", 
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        Write-Host "始業時シミュレーション完了"
    })
    
    # 閉じるボタン
    $buttonY += $buttonSpacing + 20
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "🚪 テスト終了"
    $closeButton.Size = New-Object System.Drawing.Size(200, $buttonHeight)
    $closeButton.Location = New-Object System.Drawing.Point(150, $buttonY)
    $closeButton.BackColor = [System.Drawing.Color]::LightGray
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($closeButton)
    
    # フォーム表示
    $form.ShowDialog() | Out-Null
    $form.Dispose()
    
    Write-Host "=== テスト完了 ===" -ForegroundColor Cyan
    
} catch {
    $errorMessage = "テスト中にエラーが発生しました:`n$($_.Exception.Message)"
    Write-Host $errorMessage -ForegroundColor Red
    
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage, 
        "テストエラー", 
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
