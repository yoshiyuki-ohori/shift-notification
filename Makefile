# シフト通知システム - Makefile
# Google Apps Script デプロイ & テスト

# 本番デプロイID (.env の WEBAPP_URL と一致させること)
DEPLOY_ID := AKfycbzfdn-x1nc0yIj5Yc5k7WYVlgZR1bjBmWjlnC4-_vi6984hPqandVLSPPA3i4ea27J1

.PHONY: test deploy deploy-new push pull create login setup help deployments

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

deploy: push ## プッシュ + 本番デプロイ更新 (URLそのまま)
	clasp deploy -i $(DEPLOY_ID) -d "$(shell date '+%Y-%m-%d %H:%M')"
	@echo ""
	@echo "デプロイ更新完了。Webhook URLはそのままです。"

deploy-new: push ## プッシュ + 新規デプロイ作成 (新しいURLが発行される)
	clasp deploy -d "$(shell date '+%Y-%m-%d %H:%M')"
	@echo ""
	@echo "新規デプロイ作成完了。LINE DevelopersのWebhook URLを更新してください。"

deployments: ## デプロイ一覧表示
	clasp deployments

open: ## GASエディタをブラウザで開く
	clasp open

master: ## 従業員マスタTSV生成
	node tools/generate-master.js

logs: ## GASログを表示
	clasp logs --watch
