# 📅 Outlook 15分前アラーム システム

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![Windows](https://img.shields.io/badge/Windows-10/11-green.svg)
![Outlook](https://img.shields.io/badge/Outlook-2016+-orange.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## 概要
Outlookの予定を監視し、開始15分前に大きなメッセージボックスでアラートを表示するシステムです。

## ファイル構成
- `OutlookReminder.ps1` - メインのアラームスクリプト
- `Setup.ps1` - 初期設定用スクリプト
- `Test.ps1` - テスト用スクリプト
- `README.md` - このファイル

## インストール手順

### 1. 準備
1. Outlookがインストールされていることを確認
2. PowerShellを管理者権限で開く

### 2. 初期設定
```powershell
# Setup.ps1を実行
.\Setup.ps1
```

### 3. タスクスケジューラ設定
1. `Win + R` → `taskschd.msc` → Enter
2. 右側「基本タスクの作成」をクリック
3. 以下の設定を入力:

**基本設定:**
- 名前: `OutlookReminder`
- 説明: `Outlook 15分前アラーム`

**トリガー:**
- 毎日
- 開始時刻: `08:00` (業務開始時刻)
- 詳細設定で「繰り返し間隔」をチェック
- 間隔: `5分間`
- 継続時間: `12時間`

**操作:**
- プログラムの開始
- プログラム: `powershell.exe`
- 引数: `-ExecutionPolicy Bypass -File "C:\Scripts\OutlookReminder.ps1"`

### 4. テスト
```powershell
# テストスクリプトを実行
.\Test.ps1
```

## 動作仕様

### チェック間隔
- 5分おきに自動実行
- 15分前～直前の予定を検出

### 表示内容
- 📅 件名
- ⏰ 開始時間
- 📍 場所
- ⏱️ 所要時間
- 残り時間

### 通知方法
- 大きなメッセージボックス
- システム音（警告音）
- 画面中央に表示

## カスタマイズ

### 通知タイミングの変更
`OutlookReminder.ps1`の15行目を編集:
```powershell
$checkTime = $now.AddMinutes(15)  # 15分前 → 他の分数に変更
```

### チェック間隔の変更
タスクスケジューラの繰り返し間隔を変更

### 通知音の変更
25行目の音声を変更:
```powershell
[System.Media.SystemSounds]::Exclamation.Play()  # 他の音に変更可能
# 例: Beep, Hand, Question, Asterisk
```

## トラブルシューティング

### Outlookに接続できない
- Outlookが起動していることを確認
- Outlookのセキュリティ設定を確認

### スクリプトが実行されない
- PowerShellの実行ポリシーを確認
- タスクスケジューラの設定を再確認

### 通知が表示されない
- 予定が15分以内にあることを確認
- Test.ps1でテスト実行

## 手動実行
```powershell
# 現在のディレクトリから実行
.\OutlookReminder.ps1

# フルパスで実行
C:\Scripts\OutlookReminder.ps1
```

## ログ確認
スクリプト実行時にコンソールにログが出力されます。タスクスケジューラの履歴からも確認可能です。
