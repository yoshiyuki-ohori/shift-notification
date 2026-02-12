# シフト通知システム - Makefile
# Google Apps Script デプロイ & テスト

.PHONY: test deploy push pull create login setup help

help: ## ヘルプ表示
	@echo "シフト通知システム - コマンド一覧"
	@echo ""
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'

test: ## ローカルテスト実行
	node tests/local-test.js

create: ## GASプロジェクト新規作成 (初回のみ)
	clasp create --type sheets --title "シフト通知管理" --rootDir src
	@echo ""
	@echo "作成完了。.clasp.json が生成されました。"
	@echo "次に make push でコードをアップロードしてください。"

login: ## claspにログイン
	clasp login

push: ## GASにコードをプッシュ
	clasp push

pull: ## GASからコードをプル
	clasp pull

deploy: push ## コードプッシュ + Webアプリデプロイ
	clasp deploy --description "$(shell date '+%Y-%m-%d %H:%M')"
	@echo ""
	@echo "デプロイ完了。LINE DevelopersのWebhook URLを更新してください。"

open: ## GASエディタをブラウザで開く
	clasp open

master: ## 従業員マスタTSV生成
	node tools/generate-master.js

logs: ## GASログを表示
	clasp logs --watch
