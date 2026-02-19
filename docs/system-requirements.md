# 統合シフト・人事管理システム 要件定義書

**プロジェクト名:** shift-notification (拡張)
**作成日:** 2026-02-18
**ステータス:** Phase 1 実装中 / Phase 2-5 要件定義

---

## 1. システム全体像

### 1.1 フェーズ構成

| Phase | 領域 | 概要 | 依存 |
|-------|------|------|------|
| **Phase 1** | シフト管理 | 作成→希望→調整→確定→配布 | なし (現在実装中) |
| **Phase 2** | 労務管理 | 勤怠記録・残業・有給・36協定 | Phase 1 |
| **Phase 3** | 給与・インセンティブ | 時給計算・手当・報奨 | Phase 1, 2 |
| **Phase 4** | 評価 | 定期評価・KPI・フィードバック | Phase 1, 2 |
| **Phase 5** | 教育・研修 | 研修管理・資格・スキルマップ | Phase 1 |

### 1.2 既存基盤（活用するもの）

| レイヤー | 技術 | 用途 |
|----------|------|------|
| アプリ基盤 | Google Apps Script (GAS) | メインロジック・UI |
| サーバレス | Cloud Functions (Node.js) | Webhook・AI分類 |
| DB | Google Sheets + Firestore | データストア |
| メッセージ | LINE Messaging API | 職員通知・双方向通信 |
| AI | Gemini 2.5 Flash | メッセージ分類・自動応答 |
| ファイル | Google Drive | ファイル管理 |
| 可視化 | 静的HTML (dashboard) | 管理者/職員向けダッシュボード |

### 1.3 データモデル概要

```
従業員 (Employee)
├── 基本情報: No, 氏名, フリガナ, LINE_UserId, ステータス
├── 所属: エリア, 主担当施設
├── シフト (Phase 1)
│   ├── シフト希望 (ShiftPreference)
│   ├── 確定シフト (ShiftAssignment)
│   └── シフト実績 (ShiftActual)
├── 労務 (Phase 2)
│   ├── 勤怠記録 (Attendance)
│   ├── 有給管理 (PaidLeave)
│   └── 残業記録 (Overtime)
├── 給与 (Phase 3)
│   ├── 給与マスタ (PayRate)
│   ├── 手当 (Allowance)
│   ├── インセンティブ (Incentive)
│   └── 月次給与 (MonthlyPay)
├── 評価 (Phase 4)
│   ├── 評価記録 (Evaluation)
│   └── KPI実績 (KpiRecord)
└── 教育 (Phase 5)
    ├── 研修記録 (Training)
    ├── 資格 (Certification)
    └── スキル (Skill)

施設 (Facility)
├── 基本情報: ID, 名称, 住所, 部屋, エリア
├── 利用者 (CareUser)
│   └── 利用者特性 (CareUserProfile)
└── 必要配置 (StaffingRequirement)
```

---

## 2. Phase 1: シフト管理

### 2.1 ワークフロー

```
[1. 作成] → [2. 希望収集] → [3. 調整] → [4. 確定] → [5. 配布]
                                  ↑
                            [差し戻し]
```

### 2.2 機能一覧

#### 2.2.1 シフト作成（管理者）

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| S-01 | テンプレート生成 | 月次シフト表の雛形を施設×日×時間帯で自動生成 | GAS + Sheets |
| S-02 | 前月コピー | 前月シフトをベースに翌月の初期配置を生成 | GAS |
| S-03 | 必要人数設定 | 施設×時間帯ごとの必要配置人数を設定 | Sheets マスタ |
| S-04 | 施設グループ管理 | 同一建物内施設（同一①②等）のグループ管理 | Config |

#### 2.2.2 シフト希望収集（職員→管理者）

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| S-05 | 希望入力フォーム | LINEリッチメニューから希望日・NGを入力 | LINE Flex + Webhook |
| S-06 | 希望収集期間管理 | 開始日・締切日の設定と自動リマインド | GAS Trigger |
| S-07 | 希望一覧表示 | 収集済み希望をダッシュボードで可視化 | HTML Dashboard |
| S-08 | 未提出者リマインド | 締切前に未提出者へLINE通知 | LINE Push + Trigger |
| S-09 | 希望集計 | 日別の希望/NG人数を集計 | GAS |

#### 2.2.3 シフト調整（管理者）

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| S-10 | 自動仮配置 | 希望・スキル・公平性を考慮した自動割当 | GAS or Cloud Functions |
| S-11 | 過不足警告 | 必要人数に対する過不足をハイライト | Dashboard |
| S-12 | 連勤チェック | 6連勤以上・夜勤→日勤等の労基違反チェック | GAS バリデーション |
| S-13 | 公平性レポート | 職員間の勤務日数・夜勤回数の偏りを表示 | Dashboard |
| S-14 | 手動調整UI | ドラッグ&ドロップ or セル編集でシフト変更 | Sheets + Dashboard |
| S-15 | 調整履歴 | 誰がいつ何を変更したかのログ | Sheets ログ |

#### 2.2.4 シフト確定（管理者）

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| S-16 | 確定アクション | シフトを「確定」ステータスに変更 | GAS メニュー |
| S-17 | 確定前バリデーション | 未配置・労基違反が残っていないかチェック | GAS |
| S-18 | 確定ロック | 確定後のシフト変更を制限（変更は要承認） | Sheets 保護 |

#### 2.2.5 シフト配布（管理者→職員）

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| S-19 | LINE一括配信 | 確定シフトをFlex Messageで全職員に送信 | 既存: BatchSender |
| S-20 | 個人シフト表示 | マイシフト画面で自分のシフト確認 | 既存: Dashboard |
| S-21 | Googleカレンダー連携 | 確定シフトをカレンダーに自動登録 | 既存: CalendarExport |
| S-22 | 施設詳細リンク | 施設情報・利用者特性ページへのジャンプ | 既存: facilities/*.html |
| S-23 | 変更通知 | 確定後の変更時に該当者へ差分通知 | LINE Push |

#### 2.2.6 シフト変更・交換（職員）

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| S-24 | 変更依頼 | LINEから変更依頼を送信 | LINE + AI分類 (既存カテゴリB) |
| S-25 | 交換リクエスト | 職員間のシフト交換申請 | LINE Flex + Webhook |
| S-26 | 管理者承認 | 変更・交換の承認/却下 | LINE Quick Reply |

### 2.3 Phase 1 データ構造

#### シフト希望 (ShiftPreference) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| YEAR_MONTH | string | 対象年月 (2026-03) |
| EMPLOYEE_NO | string | 従業員番号 |
| NAME | string | 氏名 |
| DATE | date | 希望日 |
| TYPE | string | 希望/NG/どちらでも |
| TIME_SLOT | string | 希望時間帯（空=終日） |
| REASON | string | 理由（任意） |
| SUBMITTED_AT | datetime | 提出日時 |

#### 必要配置 (StaffingRequirement) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| FACILITY_ID | string | 施設ID |
| TIME_SLOT | string | 時間帯 |
| DAY_TYPE | string | 平日/土曜/日祝 |
| MIN_STAFF | number | 最低人数 |
| PREFERRED_STAFF | number | 推奨人数 |

---

## 3. Phase 2: 労務管理

### 3.1 機能一覧

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| L-01 | 出退勤打刻 | LINE or ダッシュボードから出退勤記録 | LINE Beacon / GPS / 手動 |
| L-02 | シフト実績記録 | 確定シフト vs 実績の差分管理 | Sheets |
| L-03 | 残業自動計算 | 所定時間超過分の自動算出 | GAS |
| L-04 | 36協定チェック | 月45h/年360h超過警告 | GAS バリデーション |
| L-05 | 有給管理 | 付与・取得・残日数の管理 | Sheets |
| L-06 | 有給申請 | LINEから有給申請→管理者承認 | LINE + Webhook |
| L-07 | 勤怠月次サマリー | 月別の勤務時間・残業・有給を集計 | Dashboard |
| L-08 | 遅刻・早退管理 | 予定 vs 実績の乖離を記録 | GAS |
| L-09 | 欠勤管理 | 急な欠勤の記録と代替手配フロー | LINE (既存AI分類カテゴリB) |

### 3.2 データ構造

#### 勤怠記録 (Attendance) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| DATE | date | 日付 |
| EMPLOYEE_NO | string | 従業員番号 |
| FACILITY_ID | string | 施設ID |
| CLOCK_IN | datetime | 出勤時刻 |
| CLOCK_OUT | datetime | 退勤時刻 |
| BREAK_MINUTES | number | 休憩時間（分） |
| WORK_HOURS | number | 実労働時間 |
| OVERTIME_HOURS | number | 残業時間 |
| STATUS | string | 通常/遅刻/早退/欠勤/有給 |
| NOTE | string | 備考 |

#### 有給管理 (PaidLeave) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| EMPLOYEE_NO | string | 従業員番号 |
| FISCAL_YEAR | string | 年度 |
| GRANTED | number | 付与日数 |
| USED | number | 取得日数 |
| REMAINING | number | 残日数 |
| CARRY_OVER | number | 繰越日数 |

---

## 4. Phase 3: 給与・インセンティブ

### 4.1 機能一覧

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| P-01 | 時給マスタ | 職員別・時間帯別の時給設定 | Sheets マスタ |
| P-02 | 手当マスタ | 夜勤手当・休日手当・交通費等 | Sheets マスタ |
| P-03 | 月次給与計算 | シフト実績×時給+手当の自動計算 | GAS |
| P-04 | インセンティブ計算 | 条件ベースの報奨金自動算出 | GAS |
| P-05 | 給与明細生成 | 月次給与明細をPDF or LINE配信 | GAS + LINE |
| P-06 | 給与集計ダッシュボード | 全体・個人別の給与集計表示 | Dashboard |
| P-07 | 源泉徴収管理 | 税額の自動計算 | GAS |

### 4.2 インセンティブ条件（例）

| 条件 | 報奨金 | 備考 |
|------|--------|------|
| 月間勤務 N回以上 | ¥X,000 | 出勤回数ベース |
| 夜勤 N回以上 | ¥X,000/回 | 夜勤手当に加算 |
| 急な欠員カバー | ¥X,000/回 | 48h以内の代替出勤 |
| 無遅刻無欠勤（月間） | ¥X,000 | 皆勤手当 |
| 新人教育担当 | ¥X,000/回 | OJT実施 |
| 資格取得 | ¥X,000（一時金） | 対象資格リスト |

### 4.3 データ構造

#### 給与マスタ (PayRate) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| EMPLOYEE_NO | string | 従業員番号 |
| EFFECTIVE_FROM | date | 適用開始日 |
| BASE_HOURLY | number | 基本時給 |
| NIGHT_PREMIUM | number | 夜勤割増率 |
| HOLIDAY_PREMIUM | number | 休日割増率 |
| TRANSPORT_ALLOWANCE | number | 交通費（1回あたり） |

#### 月次給与 (MonthlyPay) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| YEAR_MONTH | string | 対象年月 |
| EMPLOYEE_NO | string | 従業員番号 |
| WORK_DAYS | number | 出勤日数 |
| TOTAL_HOURS | number | 総労働時間 |
| OVERTIME_HOURS | number | 残業時間 |
| BASE_PAY | number | 基本給 |
| OVERTIME_PAY | number | 残業手当 |
| NIGHT_PAY | number | 夜勤手当 |
| HOLIDAY_PAY | number | 休日手当 |
| INCENTIVE | number | インセンティブ |
| TRANSPORT | number | 交通費 |
| GROSS_PAY | number | 総支給額 |
| DEDUCTIONS | number | 控除合計 |
| NET_PAY | number | 手取り |
| STATUS | string | 計算済/確定/支払済 |

---

## 5. Phase 4: 評価

### 5.1 機能一覧

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| E-01 | 評価期間設定 | 四半期/半期/年次の評価サイクル | Sheets 設定 |
| E-02 | 自己評価入力 | LINEフォーム or ダッシュボードから入力 | LINE Flex / Dashboard |
| E-03 | 上長評価入力 | 管理者による評価記入 | Sheets / Dashboard |
| E-04 | KPI自動算出 | シフト実績から定量指標を自動計算 | GAS |
| E-05 | 評価面談記録 | 面談内容・合意事項の記録 | Sheets |
| E-06 | 評価履歴参照 | 過去の評価推移を表示 | Dashboard |

### 5.2 KPI指標（自動算出可能）

| KPI | 算出元 | 計算方法 |
|-----|--------|----------|
| 出勤率 | 勤怠記録 | 出勤日数 / シフト確定日数 |
| 遅刻率 | 勤怠記録 | 遅刻回数 / 出勤日数 |
| 夜勤担当率 | シフト実績 | 夜勤回数 / 全勤務回数 |
| 施設カバー数 | シフト実績 | ユニーク施設数 |
| 緊急カバー回数 | シフト変更履歴 | 急な代替出勤回数 |
| 研修受講数 | 研修記録 | 期間内の受講回数 |

### 5.3 データ構造

#### 評価記録 (Evaluation) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| PERIOD | string | 評価期間 (2026-Q1) |
| EMPLOYEE_NO | string | 従業員番号 |
| SELF_SCORE | number | 自己評価スコア (1-5) |
| MANAGER_SCORE | number | 上長評価スコア (1-5) |
| KPI_ATTENDANCE | number | 出勤率 |
| KPI_PUNCTUALITY | number | 遅刻率 |
| KPI_FLEXIBILITY | number | 施設カバー数 |
| KPI_EMERGENCY | number | 緊急カバー回数 |
| STRENGTHS | string | 強み・コメント |
| IMPROVEMENTS | string | 改善点 |
| GOALS | string | 次期目標 |
| FINAL_GRADE | string | 最終評価 (S/A/B/C/D) |
| EVALUATED_AT | datetime | 評価完了日 |

---

## 6. Phase 5: 教育・研修

### 6.1 機能一覧

| # | 機能 | 説明 | 実装方式 |
|---|------|------|----------|
| T-01 | 研修マスタ | 研修プログラム一覧の管理 | Sheets マスタ |
| T-02 | 研修参加記録 | 受講日・内容・評価の記録 | Sheets |
| T-03 | 資格管理 | 保有資格・有効期限の管理 | Sheets |
| T-04 | 資格期限アラート | 更新期限30日前にLINE通知 | GAS Trigger + LINE |
| T-05 | スキルマップ | 職員×スキルのマトリクス表示 | Dashboard |
| T-06 | OJT管理 | 新人教育の進捗トラッキング | Sheets |
| T-07 | 研修リマインド | 必須研修の未受講者通知 | LINE Push |

### 6.2 データ構造

#### 研修記録 (Training) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| TRAINING_ID | string | 研修ID |
| TRAINING_NAME | string | 研修名 |
| CATEGORY | string | カテゴリ (必須/任意/OJT) |
| EMPLOYEE_NO | string | 受講者 |
| DATE | date | 受講日 |
| HOURS | number | 研修時間 |
| SCORE | number | テストスコア（任意） |
| STATUS | string | 受講済/不合格/未受講 |
| INSTRUCTOR | string | 講師名 |

#### 資格 (Certification) - 新規Sheet

| カラム | 型 | 説明 |
|--------|-----|------|
| EMPLOYEE_NO | string | 従業員番号 |
| CERT_NAME | string | 資格名 |
| CERT_NUMBER | string | 資格番号 |
| ACQUIRED_DATE | date | 取得日 |
| EXPIRY_DATE | date | 有効期限 |
| STATUS | string | 有効/期限切れ/更新中 |

---

## 7. 技術設計方針

### 7.1 アーキテクチャ判断

| 判断事項 | 決定 | 理由 |
|----------|------|------|
| メインDB | Google Sheets (継続) | 管理者が直接編集可能、GAS統合が既に成熟 |
| 大量データ | Firestore (補助) | 施設マスタ等の構造化データ、月間実績の長期保存 |
| フロントエンド | 静的HTML生成 (継続) | サーバ不要、モバイル対応済、LINE内ブラウザ対応 |
| 職員UI | LINE + Dashboard HTML | 追加アプリ不要、既存LINE連携を活用 |
| 管理者UI | Sheets + Dashboard HTML | Sheets=入力・編集、HTML=可視化・レポート |
| バッチ処理 | GAS Trigger (継続) | 月次給与計算、定期リマインド |
| 認証 | LINE UserId + 従業員番号 | 既存の社員登録フローを活用 |

### 7.2 Google Sheets 構成（最終形）

| Sheet名 | Phase | 用途 |
|----------|-------|------|
| 従業員マスタ | 既存 | 基本情報 + LINE連携 |
| シフトデータ | 既存 | 確定シフト |
| 送信ログ | 既存 | LINE送信ログ |
| 設定 | 既存 | システム設定 |
| 名寄せ未マッチ | 既存 | 名寄せ結果 |
| シフト希望 | Phase 1 | 職員のシフト希望 |
| 必要配置 | Phase 1 | 施設×時間帯の必要人数 |
| 勤怠記録 | Phase 2 | 出退勤・実績 |
| 有給管理 | Phase 2 | 有給付与・取得 |
| 給与マスタ | Phase 3 | 時給・手当設定 |
| 月次給与 | Phase 3 | 月次給与計算結果 |
| インセンティブ | Phase 3 | 報奨金条件・実績 |
| 評価記録 | Phase 4 | 定期評価 |
| 研修記録 | Phase 5 | 研修受講履歴 |
| 資格管理 | Phase 5 | 保有資格・期限 |

### 7.3 ダッシュボード拡張計画

| 対象者 | 現在 | Phase 1完了後 | 全Phase完了後 |
|--------|------|---------------|---------------|
| **管理者** | 概要/施設別/職員/問題/カバレッジ | + シフト調整/希望一覧/確定操作 | + 労務集計/給与/評価/研修 |
| **職員** | マイシフト + 施設詳細 | + 希望入力/変更依頼 | + 給与明細/自己評価/研修 |

### 7.4 LINE連携拡張

| 機能 | Phase | メッセージ種別 |
|------|-------|----------------|
| シフト配信 | 既存 | Flex Message (Push) |
| 希望入力 | Phase 1 | Quick Reply + Flex Form |
| 変更リクエスト | Phase 1 | AI分類 (既存カテゴリB) |
| 出退勤打刻 | Phase 2 | リッチメニュー + Postback |
| 有給申請 | Phase 2 | Flex Form + 承認フロー |
| 給与明細 | Phase 3 | Flex Message (Push) |
| 評価リマインド | Phase 4 | Push + Quick Reply |
| 研修案内 | Phase 5 | Push + 申込Flex |

---

## 8. Phase 1 実装ロードマップ

### 8.1 現在の完了状況

- [x] Excel → シフトデータパース (parseSheet)
- [x] 名寄せエンジン (matchName + aliases)
- [x] 時間帯統合 (mergeOvernightShifts)
- [x] LINE Flex Message配信 (pushMessage + BatchSender)
- [x] Googleカレンダー連携 (CalendarExport)
- [x] 管理者ダッシュボード (HTML: 5タブ)
- [x] 職員マイシフト表示 (HTML: モバイル対応)
- [x] 施設詳細ページ生成 (facilities/*.html)
- [x] AI分類による変更依頼受付 (Cloud Functions)

### 8.2 Phase 1 残タスク

| # | タスク | 優先度 | 工数目安 |
|---|--------|--------|----------|
| 1 | シフト希望入力 LINE UI | 高 | LINEリッチメニュー + Flex Form設計 |
| 2 | 希望収集 Webhook処理 | 高 | Cloud Functions / GAS 拡張 |
| 3 | 希望データ → Sheets書込 | 高 | GAS |
| 4 | 希望一覧ダッシュボード | 中 | Dashboard タブ追加 |
| 5 | 必要配置マスタ設定 | 中 | Sheets + Config |
| 6 | 自動仮配置ロジック | 中 | GAS / Cloud Functions |
| 7 | 過不足・連勤チェック | 中 | GAS バリデーション |
| 8 | 確定ワークフロー | 高 | GAS メニュー + ステータス管理 |
| 9 | 変更通知（差分配信） | 低 | LINE Push |
| 10 | シフト交換機能 | 低 | LINE Flex + Webhook |

---

## 9. セキュリティ・運用

### 9.1 アクセス制御

| 対象 | 認証方式 | 権限 |
|------|----------|------|
| 管理者 | Google Account + ADMIN_API_KEY | 全機能 |
| 職員 | LINE UserId + 従業員番号 | 自分のデータのみ |
| Dashboard | 職員選択ドロップダウン | 選択した職員のシフトのみ表示 |

### 9.2 将来的なセキュリティ強化

- Dashboard HTML にパスワード or トークン認証追加
- 個人情報を含むページのアクセス制限
- 給与データは暗号化 or 別Sheet with 制限共有

### 9.3 バックアップ

| 対象 | 方式 | 頻度 |
|------|------|------|
| Google Sheets | Google自動バージョン管理 | 自動 |
| Firestore | GCP自動バックアップ | 日次 |
| Excel原本 | Google Drive保存 | シフト取込時 |
| 生成HTML | Git管理対象外 (再生成可能) | - |

---

## 10. 用語集

| 用語 | 説明 |
|------|------|
| GH | グループホーム (施設IDプレフィックス) |
| 名寄せ | Excel上の表記揺れを正式名に統合する処理 |
| 時間帯統合 | 17-22時+22時～+翌6-9時 → 17時～翌9時 に統合 |
| Flex Message | LINEのリッチUI形式メッセージ |
| Quick Reply | LINEメッセージ下部の選択ボタン |
| Postback | LINEボタンタップ時のサーバ通知イベント |
| GAS | Google Apps Script |
| 36協定 | 時間外労働の上限規制 (月45h/年360h) |
