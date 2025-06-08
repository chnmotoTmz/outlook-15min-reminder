# TaskTracker.ps1 - タスク管理ライブラリ
# 作業記録とタスク管理の共通機能

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Media

# データディレクトリの設定
$script:DataDir = Join-Path $PSScriptRoot "Data"
$script:TasksFile = Join-Path $DataDir "daily_tasks.csv"
$script:ContextFile = Join-Path $DataDir "work_context.json"
$script:SettingsFile = Join-Path $DataDir "settings.json"

# データディレクトリの作成
if (-not (Test-Path $DataDir)) {
    New-Item -ItemType Directory -Path $DataDir -Force | Out-Null
    Write-Host "データディレクトリを作成しました: $DataDir"
}

# 設定ファイルの初期化
function Initialize-Settings {
    if (-not (Test-Path $SettingsFile)) {
        $defaultSettings = @{
            endOfDayTime = "17:30"
            startOfDayTime = "09:00"
            routineReminderInterval = 30
            monitoredApps = @("WINWORD", "EXCEL", "POWERPNT", "notepad", "Code")
            priorityLevels = @("高", "中", "低")
        }
        $defaultSettings | ConvertTo-Json -Depth 3 | Out-File $SettingsFile -Encoding UTF8
        Write-Host "設定ファイルを作成しました: $SettingsFile"
    }
}

# 設定の読み込み
function Get-Settings {
    Initialize-Settings
    try {
        return Get-Content $SettingsFile -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Warning "設定ファイルの読み込みに失敗しました。デフォルト設定を使用します。"
        return @{
            endOfDayTime = "17:30"
            startOfDayTime = "09:00"
            routineReminderInterval = 30
        }
    }
}

# 現在の作業コンテキストを取得
function Get-WorkContext {
    $context = @{
        timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        openFiles = @()
        activeApps = @()
        lastOutlookAppointment = $null
    }
    
    try {
        # 実行中のプロセスを取得
        $processes = Get-Process | Where-Object { $_.MainWindowTitle -ne "" }
        $settings = Get-Settings
        
        foreach ($process in $processes) {
            $appInfo = @{
                name = $process.ProcessName
                title = $process.MainWindowTitle
                id = $process.Id
            }
            $context.activeApps += $appInfo
        }
        
        # Outlookの最新予定を取得
        try {
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            $calendar = $namespace.GetDefaultFolder(9)
            
            $appointments = $calendar.Items
            $appointments.Sort("[Start]")
            
            $today = Get-Date -Format "yyyy-MM-dd"
            $todayAppointments = @()
            
            foreach ($appointment in $appointments) {
                $appointmentDate = $appointment.Start.ToString("yyyy-MM-dd")
                if ($appointmentDate -eq $today) {
                    $todayAppointments += @{
                        subject = $appointment.Subject
                        start = $appointment.Start.ToString("HH:mm")
                        location = $appointment.Location
                    }
                }
            }
            
            if ($todayAppointments.Count -gt 0) {
                $context.lastOutlookAppointment = $todayAppointments[-1]
            }
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        } catch {
            Write-Warning "Outlookの予定取得に失敗しました: $($_.Exception.Message)"
        }
        
    } catch {
        Write-Warning "作業コンテキストの取得に失敗しました: $($_.Exception.Message)"
    }
    
    return $context
}

# 作業コンテキストを保存
function Save-WorkContext {
    param($context)
    
    try {
        $context | ConvertTo-Json -Depth 4 | Out-File $ContextFile -Encoding UTF8
        Write-Host "作業コンテキストを保存しました"
    } catch {
        Write-Error "作業コンテキストの保存に失敗しました: $($_.Exception.Message)"
    }
}

# 保存された作業コンテキストを読み込み
function Get-SavedWorkContext {
    if (Test-Path $ContextFile) {
        try {
            return Get-Content $ContextFile -Encoding UTF8 | ConvertFrom-Json
        } catch {
            Write-Warning "作業コンテキストの読み込みに失敗しました"
        }
    }
    return $null
}

# タスクをCSVに保存
function Save-Task {
    param(
        [string]$content,
        [string]$priority = "中",
        [bool]$isContinuation = $false,
        [string]$status = "未完了"
    )
    
    $taskData = [PSCustomObject]@{
        Date = Get-Date -Format "yyyy-MM-dd"
        Time = Get-Date -Format "HH:mm:ss"
        Content = $content
        Priority = $priority
        Status = $status
        IsContinuation = $isContinuation
        CreatedFor = (Get-Date).AddDays(1).ToString("yyyy-MM-dd")
    }
    
    # CSVファイルが存在しない場合はヘッダー付きで作成
    if (-not (Test-Path $TasksFile)) {
        $taskData | Export-Csv $TasksFile -NoTypeInformation -Encoding UTF8
    } else {
        $taskData | Export-Csv $TasksFile -NoTypeInformation -Encoding UTF8 -Append
    }
    
    Write-Host "タスクを保存しました: $content"
}

# 指定日のタスクを取得
function Get-TasksForDate {
    param([string]$date)
    
    if (Test-Path $TasksFile) {
        try {
            $allTasks = Import-Csv $TasksFile -Encoding UTF8
            return $allTasks | Where-Object { $_.CreatedFor -eq $date }
        } catch {
            Write-Warning "タスクファイルの読み込みに失敗しました"
        }
    }
    return @()
}

# 今日のタスクを取得
function Get-TodaysTasks {
    $today = Get-Date -Format "yyyy-MM-dd"
    return Get-TasksForDate $today
}

# 昨日のタスクを取得
function Get-YesterdaysTasks {
    $yesterday = (Get-Date).AddDays(-1).ToString("yyyy-MM-dd")
    return Get-TasksForDate $yesterday
}

# タスク入力フォームを表示
function Show-TaskInputForm {
    param([string]$defaultTask = "")
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "明日のタスク入力"
    $form.Size = New-Object System.Drawing.Size(500, 400)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Topmost = $true
    
    # タイトルラベル
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "📋 明日の最重要タスクを入力してください"
    $titleLabel.Size = New-Object System.Drawing.Size(450, 30)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Font = New-Object System.Drawing.Font("メイリオ", 12, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($titleLabel)
    
    # タスク入力エリア
    $taskLabel = New-Object System.Windows.Forms.Label
    $taskLabel.Text = "タスク内容:"
    $taskLabel.Size = New-Object System.Drawing.Size(100, 20)
    $taskLabel.Location = New-Object System.Drawing.Point(20, 60)
    $form.Controls.Add($taskLabel)
    
    $taskTextBox = New-Object System.Windows.Forms.TextBox
    $taskTextBox.Size = New-Object System.Drawing.Size(400, 120)
    $taskTextBox.Location = New-Object System.Drawing.Point(20, 85)
    $taskTextBox.Multiline = $true
    $taskTextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $taskTextBox.Text = $defaultTask
    $form.Controls.Add($taskTextBox)
    
    # 優先度選択
    $priorityLabel = New-Object System.Windows.Forms.Label
    $priorityLabel.Text = "優先度:"
    $priorityLabel.Size = New-Object System.Drawing.Size(100, 20)
    $priorityLabel.Location = New-Object System.Drawing.Point(20, 220)
    $form.Controls.Add($priorityLabel)
    
    $priorityCombo = New-Object System.Windows.Forms.ComboBox
    $priorityCombo.Size = New-Object System.Drawing.Size(100, 20)
    $priorityCombo.Location = New-Object System.Drawing.Point(120, 218)
    $priorityCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $priorityCombo.Items.AddRange(@("高", "中", "低"))
    $priorityCombo.SelectedIndex = 1  # デフォルトは「中」
    $form.Controls.Add($priorityCombo)
    
    # 継続タスクチェック
    $continueCheckBox = New-Object System.Windows.Forms.CheckBox
    $continueCheckBox.Text = "今日の作業の続き"
    $continueCheckBox.Size = New-Object System.Drawing.Size(200, 20)
    $continueCheckBox.Location = New-Object System.Drawing.Point(250, 220)
    $form.Controls.Add($continueCheckBox)
    
    # ボタン
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "保存"
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Location = New-Object System.Drawing.Point(250, 280)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "キャンセル"
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Location = New-Object System.Drawing.Point(340, 280)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    
    # フォーカス設定
    $taskTextBox.Select()
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    
    # フォーム表示
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if (-not [string]::IsNullOrWhiteSpace($taskTextBox.Text)) {
            Save-Task -content $taskTextBox.Text -priority $priorityCombo.SelectedItem -isContinuation $continueCheckBox.Checked
            [System.Media.SystemSounds]::Asterisk.Play()
            return $true
        } else {
            [System.Windows.Forms.MessageBox]::Show("タスク内容を入力してください。", "入力エラー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return $false
        }
    }
    
    $form.Dispose()
    return $false
}

# 情報表示フォーム
function Show-InfoDisplay {
    param(
        [string]$title,
        [string]$content,
        [int]$width = 600,
        [int]$height = 500
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size($width, $height)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.Topmost = $true
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Size = New-Object System.Drawing.Size($width - 40, $height - 100)
    $textBox.Location = New-Object System.Drawing.Point(20, 20)
    $textBox.Multiline = $true
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $textBox.ReadOnly = $true
    $textBox.Font = New-Object System.Drawing.Font("メイリオ", 10)
    $textBox.Text = $content
    $form.Controls.Add($textBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Location = New-Object System.Drawing.Point(($width / 2 - 40), ($height - 70))
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    
    $form.AcceptButton = $okButton
    $form.ShowDialog() | Out-Null
    $form.Dispose()
}

# エクスポート関数
Export-ModuleMember -Function @(
    'Get-Settings',
    'Get-WorkContext',
    'Save-WorkContext', 
    'Get-SavedWorkContext',
    'Save-Task',
    'Get-TasksForDate',
    'Get-TodaysTasks',
    'Get-YesterdaysTasks',
    'Show-TaskInputForm',
    'Show-InfoDisplay'
)
