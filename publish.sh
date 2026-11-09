#!/bin/zsh
# Excelを保存したあとに実行：閲覧ページを再生成してGitHubへ反映
set -e
cd "$(dirname "$0")"
python3 build.py "$@"
git add index.html
if git diff --cached --quiet; then echo "変更なし"; exit 0; fi
git commit -q -m "Update countdown $(date '+%Y-%m-%d %H:%M')"
git push -q
echo "反映しました（1〜2分後にページ更新）"
