# TaskSetupExtended.ps1 - 拡張システムセットアップ
# 実行方法: 管理者権限で powershell.exe -ExecutionPolicy Bypass -File "TaskSetupExtended.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Media

Write-Host "=== タスク管理拡張システムセットアップ ===" -ForegroundColor Green
Write-Host "時刻: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"

try {
    # スクリプトの存在確認
    $scriptDir = $PSScriptRoot
    $requiredFiles = @(
        "TaskTracker.ps1",
        "EndOfDayCapture.ps1", 
        "StartOfDayReminder.ps1"
    )
    
    Write-Host "
1. 必要ファイルの確認..." -ForegroundColor Yellow
    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $scriptDir $file
        if (Test-Path $filePath) {
            Write-Host "  ✓ $file - 存在" -ForegroundColor Green
        } else {
            Write-Host "  ✗ $file - 未発見" -ForegroundColor Red
            throw "$file が見つかりません。セットアップを中止します。"
        }
    }
    
    # スクリプトディレクトリの設定
    $targetDir = "C:\Scripts"
    Write-Host "
2. スクリプトディレクトリの設定..." -ForegroundColor Yellow
    
    if (-not (Test-Path $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        Write-Host "  ✓ ディレクトリ作成: $targetDir" -ForegroundColor Green
    } else {
        Write-Host "  ✓ ディレクトリ存在: $targetDir" -ForegroundColor Green
    }
    
    # ファイルのコピー
    Write-Host "
3. スクリプトファイルのコピー..." -ForegroundColor Yellow
    $allFiles = $requiredFiles + @("OutlookReminder.ps1")
    
    foreach ($file in $allFiles) {
        $sourcePath = Join-Path $scriptDir $file
        $targetPath = Join-Path $targetDir $file
        
        if (Test-Path $sourcePath) {
            Copy-Item $sourcePath $targetPath -Force
            Write-Host "  ✓ $file をコピー" -ForegroundColor Green
        } else {
            Write-Host "  ⚠ $file が見つかりません" -ForegroundColor Yellow
        }
    }
    
    # TaskTracker.ps1の初期化
    Write-Host "
4. データフォルダの初期化..." -ForegroundColor Yellow
    . (Join-Path $targetDir "TaskTracker.ps1")
    Write-Host "  ✓ データフォルダと設定ファイル初期化完了" -ForegroundColor Green
    
    # タスクスケジューラの設定
    Write-Host "
5. タスクスケジューラの設定..." -ForegroundColor Yellow
    
    # タスク定義
    $tasks = @(
        @{
            Name = "EndOfDayCapture"
            Description = "終業時タスク記録システム"
            Script = "EndOfDayCapture.ps1"
            Time = "17:30"
        },
        @{
            Name = "StartOfDayReminder"
            Description = "始業時タスク表示システム"
            Script = "StartOfDayReminder.ps1"
            Time = "09:00"
        }
    )
    
    foreach ($task in $tasks) {
        $taskName = $task.Name
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        
        if ($existingTask) {
            Write-Host "  ⚠ タスク '$taskName' は既に存在します。更新します。" -ForegroundColor Yellow
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        }
        
        # アクションの作成
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -File `"$targetDir\$($task.Script)`""
        
        # トリガーの作成
        $trigger = New-ScheduledTaskTrigger -Daily -At $task.Time
        
        # 設定の作成
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        
        # タスクの登録
        Register-ScheduledTask -TaskName $taskName -Description $task.Description -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
        
        Write-Host "  ✓ タスク '$taskName' を登録 (毎日 $($task.Time))" -ForegroundColor Green
    }
    
    # 既存のOutlookReminderタスクの確認
    $outlookTask = Get-ScheduledTask -TaskName "OutlookReminder" -ErrorAction SilentlyContinue
    if (-not $outlookTask) {
        Write-Host "  ⚠ OutlookReminderタスクが見つかりません。手動で設定してください。" -ForegroundColor Yellow
    } else {
        Write-Host "  ✓ OutlookReminderタスク確認済み" -ForegroundColor Green
    }
    
    # 権限の設定
    Write-Host "
6. ファイル権限の設定..." -ForegroundColor Yellow
    try {
        $acl = Get-Acl $targetDir
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $env:USERNAME, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow"
        )
        $acl.SetAccessRule($accessRule)
        Set-Acl $targetDir $acl
        Write-Host "  ✓ ファイル権限設定完了" -ForegroundColor Green
    } catch {
        Write-Host "  ⚠ 権限設定でエラー: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    # テスト実行
    Write-Host "
7. テスト実行..." -ForegroundColor Yellow
    $testMessage = [System.Windows.Forms.MessageBox]::Show(
        "セットアップが完了しました。`nテスト実行しますか？", 
        "セットアップ完了", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($testMessage -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "  TaskTrackerライブラリのテスト..." -ForegroundColor Cyan
        
        # テストタスクの作成
        Save-Task -content "テストタスク: システムセットアップ確認" -priority "中"
        
        # テストフォームの表示
        Show-InfoDisplay -title "テスト結果" -content "✅ システムのセットアップが正常に完了しました！`n`n今日の 17:30 に終業時記録が、`n明日の 09:00 に始業時タスク表示が自動実行されます。"
        
        Write-Host "  ✓ テスト完了" -ForegroundColor Green
    }
    
    # 最終確認メッセージ
    $finalMessage = @"
🎉 タスク管理拡張システムのセットアップが完了しました！

🕰️ 自動実行スケジュール:
• 毎日 17:30 - 終業時記録 (明日のタスク入力)
• 毎日 09:00 - 始業時タスク表示

📁 ファイル場所: $targetDir
📊 データ保存: $targetDir\Data\

🚀 今日から早速使えます！
"@
    
    [System.Windows.Forms.MessageBox]::Show(
        $finalMessage, 
        "✨ セットアップ完了", 
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    [System.Media.SystemSounds]::Exclamation.Play()
    
    Write-Host "
=== セットアップ完了 ===" -ForegroundColor Green
    Write-Host "終了時刻: $(Get-Date -Format 'yyyy/MM/dd HH:mm:ss')"
    
} catch {
    $errorMessage = "セットアップ中にエラーが発生しました:`n$($_.Exception.Message)"
    Write-Host $errorMessage -ForegroundColor Red
    
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage, 
        "セットアップエラー", 
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

Write-Host "
Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
