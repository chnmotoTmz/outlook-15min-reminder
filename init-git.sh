#!/bin/bash
# Git初期化とコミット用スクリプト

echo "=== Git リポジトリ初期化 ==="

# Git初期化
git init

# 全ファイルをステージング
git add .

# 初回コミット
git commit -m "🎉 初回コミット: Outlook 15分前アラームシステム

✨ 機能:
- PowerShellスクリプトでOutlook予定監視
- 15分前に大きなメッセージボックス表示
- タスクスケジューラで5分間隔自動実行
- 通知音とカスタマイズ対応

📁 ファイル構成:
- OutlookReminder.ps1 (メインスクリプト)
- Setup.ps1 (自動セットアップ)
- Test.ps1 (テスト用)
- README.md (詳細な手順書)"

echo "✅ Git初期化完了"
echo ""
echo "次のコマンドでGitHubにプッシュしてください:"
echo "git remote add origin https://github.com/YOUR_USERNAME/outlook-15min-reminder.git"
echo "git branch -M main"
echo "git push -u origin main"
