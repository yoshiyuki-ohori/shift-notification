#!/bin/bash
# sheets-api.sh - Google Sheets APIヘルパー
# claspの認証情報を使ってSpreadsheetを操作

SPREADSHEET_ID="1IB6gK5bEgbo6Q6kbP27y4mPPZ4bYCMFUFJrp3ZTCKs0"
CLASPRC="$HOME/.clasprc.json"

# アクセストークン取得（必要ならリフレッシュ）
get_access_token() {
  local refresh_token=$(python3 -c "import json; d=json.load(open('$CLASPRC')); print(d['tokens']['default']['refresh_token'])")
  local client_id=$(python3 -c "import json; d=json.load(open('$CLASPRC')); print(d['tokens']['default']['client_id'])")
  local client_secret=$(python3 -c "import json; d=json.load(open('$CLASPRC')); print(d['tokens']['default']['client_secret'])")

  # リフレッシュしてアクセストークン取得
  local response=$(curl -s -X POST "https://oauth2.googleapis.com/token" \
    -d "client_id=$client_id" \
    -d "client_secret=$client_secret" \
    -d "refresh_token=$refresh_token" \
    -d "grant_type=refresh_token")

  echo "$response" | python3 -c "import json,sys; print(json.load(sys.stdin)['access_token'])"
}

# シートの内容を取得
# Usage: read_sheet "シート名" ["A1:Z100"]
read_sheet() {
  local sheet_name="$1"
  local range="${2:-}"
  local token=$(get_access_token)

  local encoded_range
  if [ -n "$range" ]; then
    encoded_range=$(python3 -c "import urllib.parse; print(urllib.parse.quote('${sheet_name}!${range}'))")
  else
    encoded_range=$(python3 -c "import urllib.parse; print(urllib.parse.quote('${sheet_name}'))")
  fi

  curl -s "https://sheets.googleapis.com/v4/spreadsheets/$SPREADSHEET_ID/values/$encoded_range" \
    -H "Authorization: Bearer $token"
}

# シートに値を書き込む
# Usage: write_sheet "シート名!A1" '{"values": [["val1","val2"]]}'
write_sheet() {
  local range="$1"
  local body="$2"
  local token=$(get_access_token)

  local encoded_range=$(python3 -c "import urllib.parse; print(urllib.parse.quote('$range'))")

  curl -s -X PUT "https://sheets.googleapis.com/v4/spreadsheets/$SPREADSHEET_ID/values/$encoded_range?valueInputOption=USER_ENTERED" \
    -H "Authorization: Bearer $token" \
    -H "Content-Type: application/json" \
    -d "$body"
}

# シートに値を追記
# Usage: append_sheet "シート名" '{"values": [["val1","val2"]]}'
append_sheet() {
  local sheet_name="$1"
  local body="$2"
  local token=$(get_access_token)

  local encoded_range=$(python3 -c "import urllib.parse; print(urllib.parse.quote('$sheet_name'))")

  curl -s -X POST "https://sheets.googleapis.com/v4/spreadsheets/$SPREADSHEET_ID/values/$encoded_range:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS" \
    -H "Authorization: Bearer $token" \
    -H "Content-Type: application/json" \
    -d "$body"
}

# スプレッドシートのメタ情報（シート一覧）
get_sheet_info() {
  local token=$(get_access_token)
  curl -s "https://sheets.googleapis.com/v4/spreadsheets/$SPREADSHEET_ID?fields=sheets.properties" \
    -H "Authorization: Bearer $token"
}

# コマンドルーター
case "$1" in
  read)   read_sheet "$2" "$3" ;;
  write)  write_sheet "$2" "$3" ;;
  append) append_sheet "$2" "$3" ;;
  info)   get_sheet_info ;;
  token)  get_access_token ;;
  *)
    echo "Usage: $0 {read|write|append|info|token} [args...]"
    echo "  read  <シート名> [範囲]     - シート読み取り"
    echo "  write <シート名!範囲> <JSON> - セル書き込み"
    echo "  append <シート名> <JSON>     - 行追記"
    echo "  info                        - シート一覧"
    echo "  token                       - アクセストークン取得"
    ;;
esac
