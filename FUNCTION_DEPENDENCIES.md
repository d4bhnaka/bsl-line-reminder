# 関数依存関係図

## pollEmails() の呼び出しフロー

```
pollEmails() [エントリーポイント]
├─ getOrCreateLabel() ← Gmailラベルの取得/作成
├─ [メールごとのループ]
│   ├─ safePlainBody() ← メール本文の取得
│   ├─ parseCustomerName() ← 顧客名の抽出
│   ├─ parseSchedule() ← 日程の抽出
│   │   ├─ parseDurationEstimate() ← 所要時間目安の抽出
│   │   ├─ extractFieldAfter() ← フィールド抽出
│   │   │   └─ trimJaSpaces() ← 空白トリム
│   │   └─ buildDate() ← 日付構築
│   ├─ getTimezone() ← タイムゾーン取得
│   ├─ buildMailMeta() ← メタデータ構築
│   │   └─ buildGmailPermalink() ← パーマリンク生成
│   │
│   ├─ [scheduleがある場合]
│   │   ├─ notifyLineNow() ← LINE通知（成功時）
│   │   │   ├─ formatJst() ← 日時フォーマット
│   │   │   └─ pushLineMessageToAll() ← 全アカウントに送信
│   │   │       ├─ getLineAccounts() ← アカウント情報取得
│   │   │       └─ pushLineMessage() ← メッセージ送信
│   │   │           └─ fetchWithRetry() ← リトライ付きHTTPリクエスト
│   │   │
│   │   ├─ createCalendarEvent() ← カレンダーイベント作成
│   │   │   ├─ getCalendar() ← カレンダー取得
│   │   │   ├─ getCalendarId() ← カレンダーID取得
│   │   │   └─ parseCouponInfo() ← クーポン情報の抽出
│   │   │       └─ extractCouponNameAndPrice() ← クーポン名と価格の抽出
│   │   │
│   │   └─ parseCouponInfo() ← クーポン情報の抽出（タイトル用）
│   │
│   └─ [scheduleがない場合]
│       └─ notifyLineFallback() ← LINE通知（失敗時）
│           └─ pushLineMessageToAll() ← 全アカウントに送信
│               └─ （上記と同じ）
```

## 関数の分類と呼び出し回数

### レベル 0: エントリーポイント（1 個）

- `pollEmails()` - トリガーから実行されるメイン関数

### レベル 1: pollEmails()から直接呼び出される関数（8 個）

1. `getOrCreateLabel()` - ラベル作成/取得
2. `safePlainBody()` - メール本文取得
3. `parseCustomerName()` - 顧客名抽出
4. `parseSchedule()` - 日程抽出
5. `getTimezone()` - タイムゾーン取得
6. `buildMailMeta()` - メタデータ構築
7. `notifyLineNow()` - LINE 通知（成功時）
8. `createCalendarEvent()` - カレンダーイベント作成
9. `parseCouponInfo()` - クーポン情報抽出
10. `notifyLineFallback()` - LINE 通知（失敗時）

### レベル 2: レベル 1 から呼び出される関数（9 個）

1. `parseDurationEstimate()` ← parseSchedule()
2. `extractFieldAfter()` ← parseSchedule()
3. `buildDate()` ← parseSchedule()
4. `buildGmailPermalink()` ← buildMailMeta()
5. `formatJst()` ← notifyLineNow()
6. `pushLineMessageToAll()` ← notifyLineNow(), notifyLineFallback()
7. `getCalendar()` ← createCalendarEvent()
8. `getCalendarId()` ← createCalendarEvent()
9. `extractCouponNameAndPrice()` ← parseCouponInfo()

### レベル 3: レベル 2 から呼び出される関数（3 個）

1. `trimJaSpaces()` ← extractFieldAfter()
2. `getLineAccounts()` ← pushLineMessageToAll()
3. `pushLineMessage()` ← pushLineMessageToAll()

### レベル 4: レベル 3 から呼び出される関数（1 個）

1. `fetchWithRetry()` ← pushLineMessage()

### 独立した関数（手動実行のみ）（12 個）

#### セットアップ/管理用（5 個）

1. `setupInitial()` - 初回セットアップ
2. `deleteAllTriggers()` - 全トリガー削除
3. `suggestGmailFilter()` - Gmail フィルタ設定ガイド
4. `suggestTriggerSetup()` - トリガー設定ガイド
5. `showConfiguration()` - 設定内容の表示

#### テスト/デバッグ用（4 個）

1. `testPollEmailsIntegration()` - 統合テスト
2. `debugParse()` - 日程パースのデバッグ
3. `pingExternal()` - 外部接続テスト
4. `pingAllLineMessagingApis()` - LINE API 接続テスト
5. `testAllLineNotify()` - LINE 通知テスト

## 統計

- **合計関数数**: 34 個
- **pollEmails()から使用される関数**: 24 個
  - レベル 0: 1 個（エントリーポイント）
  - レベル 1: 10 個（直接呼び出し）
  - レベル 2: 9 個（間接呼び出し）
  - レベル 3: 3 個（間接呼び出し）
  - レベル 4: 1 個（間接呼び出し）
- **手動実行のみ**: 10 個
  - セットアップ/管理用: 5 個
  - テスト/デバッグ用: 5 個

## 重要な関数の呼び出し頻度

### 高頻度（メールごとに呼び出される）

- `safePlainBody()` - メール 1 通につき 1 回
- `parseCustomerName()` - メール 1 通につき 1 回
- `parseSchedule()` - メール 1 通につき 1 回
- `buildMailMeta()` - メール 1 通につき 1 回
- `notifyLineNow()` または `notifyLineFallback()` - メール 1 通につき 1 回

### 中頻度（スケジュールがある場合のみ）

- `createCalendarEvent()` - スケジュール抽出成功時のみ
- `parseCouponInfo()` - スケジュール抽出成功時のみ（2 回呼び出される）

### 低頻度（初回のみ）

- `getOrCreateLabel()` - pollEmails()実行時に 2 回（2 つのラベル）
- `getTimezone()` - pollEmails()実行時に 1 回
- `getLineAccounts()` - LINE 通知時に 1 回
- `getCalendar()` - カレンダーイベント作成時に 1 回
- `getCalendarId()` - カレンダーイベント作成時に 1 回

## データフロー

```
Gmail メール
    ↓
safePlainBody() → メール本文（テキスト）
    ↓
    ├→ parseCustomerName() → 顧客名
    ├→ parseSchedule() → スケジュール情報
    │   ├→ parseDurationEstimate() → 所要時間（分）
    │   ├→ extractFieldAfter() → フィールド値
    │   └→ buildDate() → Date オブジェクト
    └→ parseCouponInfo() → クーポン情報
        └→ extractCouponNameAndPrice() → クーポン名と価格
    ↓
buildMailMeta() → メタデータ
    ↓
    ├→ notifyLineNow() → LINE通知（成功時）
    │   └→ pushLineMessageToAll()
    │       └→ pushLineMessage()
    │           └→ fetchWithRetry() → LINE Messaging API
    │
    └→ createCalendarEvent() → Googleカレンダー
        ├→ getCalendar()
        └→ getCalendarId()
```

## エラーハンドリングフロー

```
parseSchedule()
    ↓
    ├─ [成功] → notifyLineNow() + createCalendarEvent()
    └─ [失敗] → notifyLineFallback()
                 └→ LINE通知のみ（カレンダー登録なし）
```

## 最適化ポイント

### キャッシュ可能な関数

以下の関数は実行ごとに同じ値を返すため、pollEmails()の実行中にキャッシュ可能です：

- `getTimezone()` - タイムゾーン（変わらない）
- `getLineAccounts()` - LINE アカウント情報（変わらない）
- `getCalendar()` - カレンダーオブジェクト（変わらない）
- `getCalendarId()` - カレンダー ID（変わらない）

### 最も重い処理

- `pushLineMessage()` → LINE Messaging API への HTTP リクエスト
- `createCalendarEvent()` → Google カレンダーへの書き込み
- `fetchWithRetry()` → リトライ付き HTTP リクエスト（最大 3 回）

### 処理順序の最適化

現在の実装は最適です：

1. 軽い処理（パース）を先に実行
2. 重い処理（API 呼び出し）を後に実行
3. エラー時は重い処理をスキップ
