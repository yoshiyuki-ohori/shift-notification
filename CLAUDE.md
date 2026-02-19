# プロジェクトルール

## Git Push ルール（必須）

### 基本方針
- **デフォルトは非公開**: `git add` する前に必ず対象ファイルを確認する
- `git add -A` や `git add .` は**禁止**。常に個別ファイルを指定する
- push 前に `git diff --cached --name-only` で追加ファイルを一覧表示して確認する

### 機密情報の定義（push 禁止）
以下を含むファイルは push してはならない:
- API キー・トークン（LINE_TOKEN, ADMIN_API_KEY 等）
- チャネルシークレット
- スプレッドシート ID
- 個人情報（氏名、LINE UserId、チャット履歴）
- .env ファイル
- .clasp.json（Script ID を含む）
- credentials/ ディレクトリ
- data/ ディレクトリ全体

### 承認が必要な操作
以下の場合はユーザーに確認を取ってから実行する:
- 新規ファイルを git に追加する場合
- .gitignore を変更する場合
- `--force` push を行う場合
- 公開リポジトリの設定を変更する場合

### push 前チェックリスト
1. `git diff --cached --name-only` で追加ファイルを確認
2. 機密情報パターンを検索: API キー、トークン、パスワード、個人名、UserId
3. .gitignore に新しい除外パターンが必要かチェック
4. ユーザーに push 対象ファイル一覧を提示して承認を得る

## 定期セキュリティチェック（1日1回）

リモートリポジトリに以下が含まれていないか確認する:
```bash
gh api repos/yoshiyuki-ohori/shift-notification/git/trees/main?recursive=1 --jq '.tree[].path'
```
チェック対象:
- data/ ディレクトリ配下のファイル
- .env, .clasp.json, credentials/ 等の機密ファイル
- *.csv, *.xlsx, *.tsv 等のデータファイル

## 技術スタック
- Google Apps Script (clasp でデプロイ)
- LINE Messaging API / LIFF
- GitHub Pages（LIFF フロントエンド）

## デプロイ手順
1. `npx clasp push --force` で GAS にアップロード
2. ユーザーが Apps Script エディタでデプロイバージョン更新
3. GitHub Pages は `docs/` ディレクトリから自動配信
4. `git add <specific files>` → `git commit` → `git push origin main`
