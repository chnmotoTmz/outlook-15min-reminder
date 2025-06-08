# StartOfDayReminder.ps1 - 始業時タスク表示スクリプト
# 実行方法: powershell.exe -ExecutionPolicy Bypass -File "StartOfDayReminder.ps1"

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
    Write-Host "=== 始業時タスク表示システム開始 ==="
    Write-Host "時刻: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
    
    # 昨日の作業コンテキストを取得
    $savedContext = Get-SavedWorkContext
    
    # 今日のタスクを取得
    $todayTasks = Get-TodaysTasks
    
    # 今日のOutlook予定を取得
    $todayAppointments = @()
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $calendar = $namespace.GetDefaultFolder(9)
        
        $appointments = $calendar.Items
        $appointments.Sort("[Start]")
        
        $today = Get-Date -Format "yyyy-MM-dd"
        
        foreach ($appointment in $appointments) {
            $appointmentDate = $appointment.Start.ToString("yyyy-MM-dd")
            if ($appointmentDate -eq $today) {
                $todayAppointments += @{
                    subject = $appointment.Subject
                    start = $appointment.Start.ToString("HH:mm")
                    location = if ($appointment.Location) { $appointment.Location } else { "場所未設定" }
                    duration = if ($appointment.Duration) { "$($appointment.Duration)分" } else { "時間未設定" }
                }
            }
        }
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    } catch {
        Write-Warning "Outlookの予定取得に失敗しました: $($_.Exception.Message)"
    }
    
    # 情報を整理して表示用テキストを作成
    $displayContent = @"
🌅 おはようございます！今日も一日よろしくお願いします。

📋 今日のタスク一覧:
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
        }
        $taskList
    } else {
        "今日のタスクは未設定です。下記の「最優先タスク設定」で設定してください。`n"
    }
)

📅 今日の予定:
$(
    if ($todayAppointments.Count -gt 0) {
        $appointmentList = ""
        foreach ($appt in $todayAppointments) {
            $appointmentList += "⏰ $($appt.start) - $($appt.subject)"
            if ($appt.location -ne "場所未設定") {
                $appointmentList += " @ $($appt.location)"
            }
            $appointmentList += " ($($appt.duration))`n"
        }
        $appointmentList
    } else {
        "今日は予定がありません。`n"
    }
)

$(
    if ($savedContext) {
        "📁 昨日の作業コンテキスト:`n"
        if ($savedContext.lastOutlookAppointment) {
            "• 最後の予定: $($savedContext.lastOutlookAppointment.subject)`n"
        }
        if ($savedContext.activeApps -and $savedContext.activeApps.Count -gt 0) {
            $importantApps = $savedContext.activeApps | Where-Object { 
                $_.name -in @("WINWORD", "EXCEL", "POWERPNT", "notepad", "Code", "chrome", "msedge") 
            } | Select-Object -First 5
            if ($importantApps.Count -gt 0) {
                "• 作業中のアプリ:`n"
                foreach ($app in $importantApps) {
                    "  - $($app.name): $($app.title)`n"
                }
            }
        }
        "`n"
    } else {
        ""
    }
)
⚙️ 朝のルーチンガイド:
1. ✅ Outlookメール確認 (Copilotで重要メールチェック)
2. ✅ 会社ポータルサイト一巡
3. ✅ Teamsタイムライン一巡

ルーチン完了後、「最優先タスク設定」ボタンを押して今日の作業を始めましょう！
"@
    
    # 通知音を鳴らす
    [System.Media.SystemSounds]::Asterisk.Play()
    
    # メインフォームを作成
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "🌅 朝のタスクガイド"
    $form.Size = New-Object System.Drawing.Size(700, 600)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Topmost = $true
    
    # テキスト表示エリア
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size(660, 480)
    $textBox.Location = New-Object System.Drawing.Point(20, 20)
    $textBox.Multiline = $true
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $textBox.ReadOnly = $true
    $textBox.Font = New-Object System.Drawing.Font("メイリオ", 10)
    $textBox.Text = $displayContent
    $form.Controls.Add($textBox)
    
    # ボタンエリア
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Size = New-Object System.Drawing.Size(660, 50)
    $buttonPanel.Location = New-Object System.Drawing.Point(20, 510)
    $form.Controls.Add($buttonPanel)
    
    # 最優先タスク設定ボタン
    $priorityButton = New-Object System.Windows.Forms.Button
    $priorityButton.Text = "🎯 最優先タスク設定"
    $priorityButton.Size = New-Object System.Drawing.Size(180, 35)
    $priorityButton.Location = New-Object System.Drawing.Point(10, 10)
    $priorityButton.BackColor = [System.Drawing.Color]::LightBlue
    $buttonPanel.Controls.Add($priorityButton)
    
    # ルーチン完了ボタン
    $routineButton = New-Object System.Windows.Forms.Button
    $routineButton.Text = "✅ ルーチン完了・作業開始"
    $routineButton.Size = New-Object System.Drawing.Size(180, 35)
    $routineButton.Location = New-Object System.Drawing.Point(200, 10)
    $routineButton.BackColor = [System.Drawing.Color]::LightGreen
    $buttonPanel.Controls.Add($routineButton)
    
    # 閉じるボタン
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "閉じる"
    $closeButton.Size = New-Object System.Drawing.Size(100, 35)
    $closeButton.Location = New-Object System.Drawing.Point(550, 10)
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $buttonPanel.Controls.Add($closeButton)
    
    # 最優先タスク設定イベント
    $priorityButton.Add_Click({
        $form.Hide()
        
        # 今日のタスクがない場合は新規作成を促す
        if ($todayTasks.Count -eq 0) {
            $inputResult = Show-TaskInputForm -defaultTask "今日の最優先タスク:"
            if ($inputResult) {
                [System.Windows.Forms.MessageBox]::Show(
                    "✅ 最優先タスクを設定しました！`n🚀 集中して作業を進めましょう。", 
                    "タスク設定完了", 
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        } else {
            # 既存タスクから最優先を選択
            $highPriorityTasks = $todayTasks | Where-Object { $_.Priority -eq "高" }
            if ($highPriorityTasks.Count -gt 0) {
                $taskSummary = "🔴 今日の高優先タスク:`n`n"
                foreach ($task in $highPriorityTasks) {
                    $taskSummary += "- $($task.Content)`n"
                }
                $taskSummary += "`n🚀 これらのタスクに集中しましょう！"
                
                [System.Windows.Forms.MessageBox]::Show(
                    $taskSummary, 
                    "高優先タスク確認", 
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            } else {
                $allTasksSummary = "📋 今日のタスク一覧:`n`n"
                foreach ($task in $todayTasks) {
                    $priorityIcon = switch ($task.Priority) {
                        "高" { "🔴" }
                        "中" { "🟡" }
                        "低" { "🔵" }
                        default { "⚪" }
                    }
                    $allTasksSummary += "$priorityIcon $($task.Content)`n"
                }
                $allTasksSummary += "`n🚀 上から順番に進めましょう！"
                
                [System.Windows.Forms.MessageBox]::Show(
                    $allTasksSummary, 
                    "タスク一覧", 
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
        
        $form.Show()
    })
    
    # ルーチン完了イベント
    $routineButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    })
    
    # 閉じるイベント
    $closeButton.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })
    
    # フォーム表示
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # ルーチン完了時のメッセージ
        $startMessage = @"
🎉 ルーチン完了お疑れ様でした！

🚀 今日も集中して作業を進めていきましょう。
✨ 良い一日になりますように！
"@
        
        [System.Windows.Forms.MessageBox]::Show(
            $startMessage, 
            "🚀 作業開始", 
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        [System.Media.SystemSounds]::Exclamation.Play()
    }
    
    $form.Dispose()
    
    # 結果レポート
    Write-Host "=== 結果レポート ==="
    Write-Host "今日のタスク数: $($todayTasks.Count)"
    Write-Host "今日の予定数: $($todayAppointments.Count)"
    Write-Host "始業時タスク表示システム終了: $(Get-Date -Format 'HH:mm:ss')"
    
} catch {
    $errorMessage = "エラーが発生しました:`n$($_.Exception.Message)"
    Write-Host $errorMessage
    
    # エラーもメッセージボックスで表示
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage, 
        "始業時タスク表示エラー", 
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
