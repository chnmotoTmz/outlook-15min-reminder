# 🚀 GitHubアップロード手順

## 手順1: GitHubで新しいリポジトリを作成
1. https://github.com にアクセス
2. 右上の「+」→「New repository」
3. Repository name: `outlook-15min-reminder`
4. Description: `PowerShell + タスクスケジューラでOutlookの予定15分前にアラート通知するシステム`
5. Public を選択
6. 「Create repository」をクリック

## 手順2: ローカルからプッシュ
PowerShellまたはコマンドプロンプトで以下を実行:

```bash
# outlook-reminderフォルダに移動
cd outlook-reminder

# Git初期化
git init

# ファイルをステージング
git add .

# 初回コミット
git commit -m "🎉 初回コミット: Outlook 15分前アラームシステム"

# リモートリポジトリを追加（YOUR_USERNAMEは自分のGitHubユーザー名に変更）
git remote add origin https://github.com/YOUR_USERNAME/outlook-15min-reminder.git

# メインブランチに設定
git branch -M main

# GitHubにプッシュ
git push -u origin main
```

## 手順3: 認証
初回プッシュ時にGitHubの認証を求められる場合があります:
- Personal Access Token を使用することを推奨
- Settings → Developer settings → Personal access tokens で作成

## 完了！
リポジトリが正常にアップロードされ、以下のファイルが公開されます:
- ✅ OutlookReminder.ps1
- ✅ Setup.ps1  
- ✅ Test.ps1
- ✅ README.md
- ✅ LICENSE
- ✅ .gitignore

## クローン方法
他の人がこのシステムを使う場合:
```bash
git clone https://github.com/YOUR_USERNAME/outlook-15min-reminder.git
cd outlook-15min-reminder
```

他のユーザーは `README.md` の手順に従って簡単にセットアップできます！
