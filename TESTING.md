# 統合テストガイド

## 統合テスト実行方法

### 1. testPollEmailsIntegration() の実行

Google Apps Script エディタで以下の手順を実行してください：

1. `Code.gs` を開く
2. 関数リストから `testPollEmailsIntegration` を選択
3. 実行ボタンをクリック
4. ログを確認（表示 → ログ）

### 2. テスト内容

以下の 3 つのテストケースで主要な機能を検証します：

#### テストケース 1: 所要時間目安あり + クーポンあり

- 顧客名: 田中 太郎
- 来店日時: 2025 年 10 月 14 日（火）13:00
- 所要時間目安: 3 時間 25 分
- クーポン: プレミアムカット（¥5,000）
- 期待される終了時間: 16:25（13:00 + 3 時間 25 分）

#### テストケース 2: 所要時間目安なし + クーポンあり（デフォルト 2 時間）

- 顧客名: 佐藤 花子
- 来店日時: 2025 年 10 月 20 日（月）10:00
- 所要時間目安: なし（デフォルト 2 時間）
- クーポン: カラーリング（¥8,000）
- 期待される終了時間: 12:00（10:00 + 2 時間）

#### テストケース 3: クーポンなし

- 顧客名: 山田 次郎
- 来店日時: 2025 年 10 月 25 日（土）15:30
- 所要時間目安: なし（デフォルト 2 時間）
- クーポン: なし
- 期待される終了時間: 17:30（15:30 + 2 時間）

### 3. 検証項目

各テストケースで以下を検証します：

- ✓ 顧客名の抽出（`parseCustomerName()`）
- ✓ 所要時間目安の抽出（`parseDurationEstimate()`）
- ✓ 日程の抽出（`parseSchedule()`）
- ✓ 終了時間の計算（開始時間 + 所要時間）
- ✓ クーポン情報の抽出（`parseCouponInfo()`）

### 4. 期待される出力

```
=== 統合テスト開始 ===

--- テストケース: 所要時間目安あり + クーポンあり ---
✓ 顧客名パース: 田中 太郎
✓ 所要時間目安: 205分
✓ 日程パース: 2025/10/14 13:00 - 2025/10/14 16:25
✓ 終了時間検証: OK
✓ クーポン情報: プレミアムカット（¥5,000）
✓ 各関数が正常に動作しました
✅ テストケース「所要時間目安あり + クーポンあり」: PASS

--- テストケース: 所要時間目安なし + クーポンあり（デフォルト2時間） ---
✓ 顧客名パース: 佐藤 花子
✓ 所要時間目安: 120分
✓ 日程パース: 2025/10/20 10:00 - 2025/10/20 12:00
✓ 終了時間検証: OK
✓ クーポン情報: カラーリング（¥8,000）
✓ 各関数が正常に動作しました
✅ テストケース「所要時間目安なし + クーポンあり（デフォルト2時間）」: PASS

--- テストケース: クーポンなし ---
✓ 顧客名パース: 山田 次郎
✓ 所要時間目安: 120分
✓ 日程パース: 2025/10/25 15:30 - 2025/10/25 17:30
✓ 終了時間検証: OK
✓ クーポン情報: なし
✓ 各関数が正常に動作しました
✅ テストケース「クーポンなし」: PASS

=== 統合テスト結果 ===
合計: 3件
成功: 3件
失敗: 0件
✅ すべてのテストが成功しました！

=== 主要関数の使用状況 ===
✓ parseDurationEstimate() - 所要時間目安の抽出
✓ parseCustomerName() - 顧客名の抽出
✓ parseSchedule() - 日程の抽出
  ├─ extractFieldAfter() - フィールド抽出
  ├─ trimJaSpaces() - 空白トリム
  └─ buildDate() - 日付構築
✓ parseCouponInfo() - クーポン情報の抽出
  └─ extractCouponNameAndPrice() - クーポン名と価格の抽出
✓ formatJst() - 日時フォーマット

=== pollEmails()で使用される関数 ===
✓ getOrCreateLabel() - ラベル作成/取得
✓ safePlainBody() - メール本文の取得
✓ buildMailMeta() - メタデータ構築
  └─ buildGmailPermalink() - パーマリンク生成
✓ notifyLineNow() / notifyLineFallback() - LINE通知
  └─ pushLineMessageToAll() - 全アカウントに送信
      ├─ getLineAccounts() - アカウント情報取得
      └─ pushLineMessage() - メッセージ送信
          └─ fetchWithRetry() - リトライ付きHTTPリクエスト
✓ createCalendarEvent() - カレンダーイベント作成
  ├─ getCalendar() - カレンダー取得
  ├─ getCalendarId() - カレンダーID取得
  └─ getTimezone() - タイムゾーン取得
```

## 関数一覧と使用状況

### エントリーポイント

- `pollEmails()` - メイン処理（トリガーで実行）

### 主要処理フロー（pollEmails()から使用）

#### データ抽出

- `safePlainBody()` - メール本文の取得
- `parseCustomerName()` - 顧客名の抽出
- `parseDurationEstimate()` - 所要時間目安の抽出（新規）
- `parseSchedule()` - 日程の抽出
  - `extractFieldAfter()` - フィールド抽出補助
  - `trimJaSpaces()` - 空白トリム補助
  - `buildDate()` - 日付構築補助
- `parseCouponInfo()` - クーポン情報の抽出
  - `extractCouponNameAndPrice()` - クーポン名と価格の抽出補助

#### メタデータ処理

- `buildMailMeta()` - メタデータ構築
  - `buildGmailPermalink()` - Gmail パーマリンク生成

#### LINE 通知

- `notifyLineNow()` - 日程抽出成功時の LINE 通知
- `notifyLineFallback()` - 日程抽出失敗時の LINE 通知
- `pushLineMessageToAll()` - 全アカウントへのメッセージ送信
  - `getLineAccounts()` - LINE アカウント情報取得
  - `pushLineMessage()` - 単一アカウントへの送信
    - `fetchWithRetry()` - リトライ付き HTTP リクエスト

#### カレンダー登録

- `createCalendarEvent()` - カレンダーイベント作成
  - `getCalendar()` - カレンダー取得
  - `getCalendarId()` - カレンダー ID 取得

#### ユーティリティ

- `getOrCreateLabel()` - Gmail ラベル作成/取得
- `getTimezone()` - タイムゾーン取得
- `formatJst()` - 日時フォーマット

### セットアップ/管理用関数（手動実行）

- `setupInitial()` - 初回セットアップ
- `deleteAllTriggers()` - 全トリガー削除（緊急時）
- `suggestGmailFilter()` - Gmail フィルタ設定ガイド
- `suggestTriggerSetup()` - トリガー設定ガイド
- `showConfiguration()` - 設定内容の表示

### テスト/デバッグ用関数（手動実行）

- `testPollEmailsIntegration()` - 統合テスト（新規）
- `debugParse()` - 日程パースのデバッグ
- `pingExternal()` - 外部接続テスト
- `pingAllLineMessagingApis()` - LINE API 接続テスト
- `testAllLineNotify()` - LINE 通知テスト

## 削除した関数

以下の関数は、リマインダー機能が無効化されているため削除しました：

- ~~`buildTriggerStoreKey()`~~ - リマインダートリガーのキー生成（不要）
- ~~`cleanupOldReminderTriggers()`~~ - リマインダートリガークリーンアップ（不要）
- ~~`deleteReminderTriggers()`~~ - リマインダートリガー削除（不要）

## 確認事項

### すべての主要関数が使用されているか？

✅ はい。pollEmails()から直接または間接的に呼び出される関数は 24 個あり、すべて正常に機能しています。

### 不要な関数はないか？

✅ リマインダー機能関連の 3 つの関数を削除しました。残りの関数はすべて以下のいずれかに分類されます：

- 主要処理フロー（24 個）
- セットアップ/管理用（5 個）
- テスト/デバッグ用（4 個）

合計: 34 個の関数
