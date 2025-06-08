# OutlookReminder.ps1 - 15分前アラーム
# 実行方法: powershell.exe -ExecutionPolicy Bypass -File "OutlookReminder.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Media

try {
    # Outlookアプリケーションに接続
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $calendar = $namespace.GetDefaultFolder(9) # カレンダーフォルダ
    
    # 現在時刻と15分後の時刻を取得
    $now = Get-Date
    $checkTime = $now.AddMinutes(15)
    
    Write-Host "チェック時刻: $($now.ToString('yyyy/MM/dd HH:mm:ss'))"
    Write-Host "15分後まで: $($checkTime.ToString('yyyy/MM/dd HH:mm:ss'))"
    
    # 予定を取得
    $appointments = $calendar.Items
    $appointments.IncludeRecurrences = $true
    $appointments.Sort("[Start]")
    
    $foundAppointments = @()
    
    foreach ($appointment in $appointments) {
        $startTime = $appointment.Start
        
        # 15分後以内に開始される予定をチェック
        if ($startTime -ge $now -and $startTime -le $checkTime) {
            $foundAppointments += $appointment
            
            # 通知音を鳴らす
            [System.Media.SystemSounds]::Exclamation.Play()
            
            # 詳細な予定情報を作成
            $subject = if ($appointment.Subject) { $appointment.Subject } else { "（件名なし）" }
            $location = if ($appointment.Location) { $appointment.Location } else { "（場所未設定）" }
            $duration = if ($appointment.Duration) { "$($appointment.Duration)分" } else { "時間未設定" }
            
            $message = @"
🔔 まもなく予定があります！

📅 件名: $subject
⏰ 開始時間: $($startTime.ToString('HH:mm'))
📍 場所: $location
⏱️ 所要時間: $duration

あと$([math]::Round(($startTime - $now).TotalMinutes, 1))分で開始です。
"@
            
            # 大きなメッセージボックスを表示
            $result = [System.Windows.Forms.MessageBox]::Show(
                $message, 
                "📢 Outlook予定通知", 
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            Write-Host "通知表示: $subject ($($startTime.ToString('HH:mm')))"
        }
    }
    
    if ($foundAppointments.Count -eq 0) {
        Write-Host "15分以内に予定はありません。"
    } else {
        Write-Host "通知した予定数: $($foundAppointments.Count)"
    }
    
} catch {
    $errorMessage = "エラーが発生しました:`n$($_.Exception.Message)"
    Write-Host $errorMessage
    
    # エラーもメッセージボックスで表示
    [System.Windows.Forms.MessageBox]::Show(
        $errorMessage, 
        "Outlook通知エラー", 
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

# COMオブジェクトの解放
if ($outlook) {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
}
