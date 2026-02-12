# シフト通知システム - セットアップ手順

## 前提条件

- Node.js (v18+)
- Google Workspace アカウント
- LINE Developers アカウント + Messaging API チャネル

## Step 1: GASプロジェクト作成

```bash
# claspにログイン
make login

# GASプロジェクト作成（Spreadsheet連携）
make create
```

これにより Google Drive に「シフト通知管理」スプレッドシートが作成され、
GASプロジェクトが紐づきます。

## Step 2: コードをアップロード

```bash
make push
```

## Step 3: GASエディタで初期セットアップ

```bash
make open
```

GASエディタで以下を実行:

1. `setupSheets` 関数を実行 → 5タブが自動生成される
2. Script Properties に以下を設定:
   - `LINE_CHANNEL_ACCESS_TOKEN`: LINE Messaging APIのチャネルアクセストークン
   - `LINE_CHANNEL_SECRET`: LINE Messaging APIのチャネルシークレット

## Step 4: 従業員マスタ投入

```bash
make master
```

`data/employee-master.tsv` が生成されるので、
スプレッドシートの「従業員マスタ」シートにTSVを貼り付け。

## Step 5: 設定シートの値を入力

「設定」シートに以下を入力:

| キー | 値 | 例 |
|------|-----|-----|
| 対象年月 | 送信対象月 | 2026-03 |
| 練馬シフトSS_ID | 練馬シフトSpreadsheetのID | 1BxC...xyz |
| 世田谷シフトSS_ID | 世田谷シフトSpreadsheetのID | 1AbC...xyz |
| 送信モード | テスト or 本番 | テスト |
| テスト送信先UserId | テスト用LINE UserId | Uxxxx... |

SpreadsheetのIDはURLから取得:
`https://docs.google.com/spreadsheets/d/【ここがID】/edit`

## Step 6: Webhook デプロイ（LINE UserID登録用）

```bash
make deploy
```

デプロイ後、表示されるWebアプリURLを LINE Developers Console の
Webhook URL に設定。

## Step 7: 動作確認

1. スプレッドシートをリロード → 「シフト通知」メニューが表示される
2. 「シフトデータ取込 (練馬)」→ 練馬シフトが読み込まれる
3. 「名寄せ実行」→ マッチ結果が表示される
4. 「未マッチ確認」→ 未マッチの名前を確認、別名リストで解消
5. 「テスト送信」→ 管理者のLINEに Flex Message が届く

## LINE UserID 登録方法（職員向け）

1. LINE公式アカウントを友だち追加
2. トーク画面で社員番号を送信（例: `072`）
3. 「石井 祐一さん、LINE連携が完了しました！」と返信が届く

## トラブルシューティング

### シフトデータが0件
- 設定シートの「対象年月」が正しいか確認
- シフトSpreadsheet IDが正しいか確認
- シフトスプレッドシートへの閲覧権限があるか確認

### 名寄せ未マッチが多い
- 「未マッチ確認」タブの候補を確認
- 従業員マスタの「別名リスト」にカンマ区切りで略称を追加
- 再度「名寄せ実行」

### LINE送信失敗
- Script Properties の LINE_CHANNEL_ACCESS_TOKEN を確認
- LINE公式アカウントの Messaging API が有効か確認
- エラー詳細は「送信ログ」タブに記録される
