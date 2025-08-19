/**
 * Apps Script (JavaScript) — Gmail→LINE通知＋日程抽出→Calendar登録＋前日LINEリマインド
 *
 * 事前準備:
 * - スクリプト プロパティ（File > Project properties > Script properties）に以下を設定
 *   - LINE_CHANNEL_ACCESS_TOKEN: Messaging API チャネルアクセストークン（長期）
 *   - TARGET_EMAIL: 監視対象の自アドレス（例: me@example.com）
 *   - CALENDAR_ID: 予定登録先カレンダーID（未設定ならデフォルト）
 *   - TIMEZONE: Asia/Tokyo （未設定でもAsia/Tokyoを既定）
 * - Gmail フィルタ: to:me@example.com → ラベル 'to-line' を自動付与
 * - トリガ: 時間主導（1〜2分毎）で pollEmails を実行
 */

const LABELS = Object.freeze({
  target: "to-line",
  processed: "processed-line",
});

const PROP_KEYS = Object.freeze({
  channelAccessToken: "LINE_CHANNEL_ACCESS_TOKEN",
  channelSecret: "LINE_CHANNEL_SECRET",
  targetEmail: "TARGET_EMAIL",
  calendarId: "CALENDAR_ID",
  timezone: "TIMEZONE",
});

// 既定値（スクリプトプロパティ未設定時のフォールバック）
const CONFIG_DEFAULTS = Object.freeze({
  targetEmail: "besol4b.shop@gmail.com",
  calendarId: "besol4b.shop@gmail.com",
  timezone: "Asia/Tokyo",
});

function pollEmails() {
  const props = PropertiesService.getScriptProperties();
  const targetEmail = (
    props.getProperty(PROP_KEYS.targetEmail) ||
    CONFIG_DEFAULTS.targetEmail ||
    ""
  ).toLowerCase();
  getOrCreateLabel(LABELS.target);
  getOrCreateLabel(LABELS.processed);

  // SALON BOARD 送信者かつ対象宛先のみ。処理済みラベルは除外。
  // 最適化: newer_than を短縮し、処理負荷を軽減
  const query = [
    "from:yoyaku_system@salonboard.com",
    targetEmail ? `to:${targetEmail}` : "",
    `-label:${LABELS.processed}`,
    "is:unread",
    "newer_than:3d", // 3日以内に短縮（予約メールは通常即座に処理されるため）
  ]
    .filter(Boolean)
    .join(" ");
  const threads = GmailApp.search(query, 0, 30); // 検索件数も30件に制限

  for (const thread of threads) {
    const messages = thread.getMessages().filter((m) => m.isUnread());
    for (const msg of messages) {
      // 宛先フィルタ（To/CCいずれかに含まれているか）
      const to = (msg.getTo() || "").toLowerCase();
      const cc = (msg.getCc() || "").toLowerCase();
      if (
        targetEmail &&
        !(to.includes(targetEmail) || cc.includes(targetEmail))
      ) {
        continue;
      }

      const body = safePlainBody(msg);
      const name = parseCustomerName(body) || "ご予約";
      const schedule = parseSchedule(body, getTimezone());
      const meta = buildMailMeta(thread, msg);

      if (schedule) {
        notifyLineNow(meta, schedule, name);
        const eventId = createCalendarEvent(schedule, meta, name, body);
        // クーポン情報を含むタイトルを作成
        const couponInfo = parseCouponInfo(body);
        const titleForReminder = couponInfo
          ? `${name}さま - ${couponInfo}`
          : `${name}さま予約`;

        scheduleReminder(schedule, {
          eventId,
          calendarId: getCalendarId(),
          title: titleForReminder,
          start: schedule.start.toISOString(),
          threadPermalink: meta.permalink,
          metaId: `${meta.threadId}:${
            meta.messageId
          }:${schedule.start.getTime()}`,
        });
      } else {
        // 抽出失敗でも最低限の通知（運用に応じてオフ可）
        notifyLineFallback(meta);
      }
    }

    // スレッド単位で処理済みラベル付与＆既読化（運用に合わせて調整）
    thread.addLabel(GmailApp.getUserLabelByName(LABELS.processed));
    thread.markRead();
  }
}

/**
 * 本文から日時を抽出（日本語表現に対応した簡易ルール）
 * - 優先度: YYYY/M/D HH:mm → M/D HH:mm → M月D日 HH時mm分 → 日付のみ（既定10:00）
 * - 年未指定: 今年、かつ過去なら翌年
 * - 時刻未指定: 10:00
 * - 範囲表現: 14:00-15:00 / 14時〜15時 → endに反映（分省略=00）
 */
function parseSchedule(text, tz) {
  const now = new Date();
  const currentYear = now.getFullYear();

  // 先に「■来店日時」セクションを優先抽出
  const visitLine = extractFieldAfter(text, "■来店日時");
  if (visitLine) {
    const v = visitLine;
    // a) YYYY年MM月DD日（曜）HH:mm（〜HH:mm任意）
    let m = v.match(
      /(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日(?:\s*[（(][^)）]+[)）])?\s*(\d{1,2})[:：](\d{1,2})(?:\s*[〜~\-]\s*(\d{1,2})[:：]?(\d{1,2})?)?/
    );
    if (m) {
      const ys = Number(m[1]);
      const ms = Number(m[2]);
      const ds = Number(m[3]);
      const hs = Number(m[4]);
      const mins = m[5] ? Number(m[5]) : 0;
      const he = m[6] ? Number(m[6]) : undefined;
      const mine = m[7] ? Number(m[7]) : 0;
      const start = buildDate(ys, ms, ds, hs, mins, tz);
      const end =
        typeof he === "number"
          ? buildDate(ys, ms, ds, he, mine, tz)
          : undefined;
      return {
        start,
        end: end && end > start ? end : undefined,
        isAllDay: false,
        sourceText: m[0],
      };
    }
    // b) YYYY/MM/DD HH:mm 等にフォールバック
    m = v.match(
      /(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})\s+(\d{1,2})(?::(\d{1,2}))?/
    );
    if (m) {
      const ys = Number(m[1]);
      const ms = Number(m[2]);
      const ds = Number(m[3]);
      const hs = Number(m[4]);
      const mins = m[5] ? Number(m[5]) : 0;
      const start = buildDate(ys, ms, ds, hs, mins, tz);
      return { start, isAllDay: false, sourceText: m[0] };
    }
    // c) M月D日 HH:mm / HH時mm分
    m = v.match(
      /(\d{1,2})月(\d{1,2})日(?:\s*[（(][^)）]+[)）])?\s*(\d{1,2})(?:[:：時](\d{1,2}))?/
    );
    if (m) {
      let year = currentYear;
      const ms = Number(m[1]);
      const ds = Number(m[2]);
      const hs = Number(m[3]);
      const mins = m[4] ? Number(m[4]) : 0;
      const startCandidate = buildDate(year, ms, ds, hs, mins, tz);
      if (startCandidate.getTime() < now.getTime() - 60 * 60 * 1000) year += 1;
      const start = buildDate(year, ms, ds, hs, mins, tz);
      return { start, isAllDay: false, sourceText: m[0] };
    }
  }

  // 1) YYYY/M/D HH:mm(-HH:mm)
  {
    const rx =
      /(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})[^\d\n]*?(\d{1,2})(?::(\d{1,2}))?(?:\s*[〜~\-]\s*(\d{1,2})(?::(\d{1,2}))?)?/;
    const m = text.match(rx);
    if (m) {
      const ys = Number(m[1]);
      const ms = Number(m[2]);
      const ds = Number(m[3]);
      const hs = Number(m[4]);
      const mins = m[5] ? Number(m[5]) : 0;
      const he = m[6] ? Number(m[6]) : undefined;
      const mine = m[7] ? Number(m[7]) : 0;
      const start = buildDate(ys, ms, ds, hs, mins, tz);
      const end =
        typeof he === "number"
          ? buildDate(ys, ms, ds, he, mine, tz)
          : undefined;
      return {
        start,
        end: end && end > start ? end : undefined,
        isAllDay: false,
        sourceText: m[0],
      };
    }
  }

  // 2) M/D (曜日) HH:mm(-HH:mm)
  {
    const rx =
      /(\d{1,2})[\/\-\.](\d{1,2})(?:\s*\([^)]+\))?\s+(\d{1,2})(?::(\d{1,2}))?(?:\s*[〜~\-]\s*(\d{1,2})(?::(\d{1,2}))?)?/;
    const m = text.match(rx);
    if (m) {
      const ms = Number(m[1]);
      const ds = Number(m[2]);
      const hs = Number(m[3]);
      const mins = m[4] ? Number(m[4]) : 0;
      const he = m[5] ? Number(m[5]) : undefined;
      const mine = m[6] ? Number(m[6]) : 0;
      let year = currentYear;
      const startCandidate = buildDate(year, ms, ds, hs, mins, tz);
      if (startCandidate.getTime() < now.getTime() - 60 * 60 * 1000) {
        year += 1; // 過去なら翌年
      }
      const start = buildDate(year, ms, ds, hs, mins, tz);
      const end =
        typeof he === "number"
          ? buildDate(year, ms, ds, he, mine, tz)
          : undefined;
      return {
        start,
        end: end && end > start ? end : undefined,
        isAllDay: false,
        sourceText: m[0],
      };
    }
  }

  // 3) M月D日 HH時mm分(-HH時mm分)
  {
    const rx =
      /(\d{1,2})月(\d{1,2})日(?:\s*\([^)]+\))?\s*(\d{1,2})時(?:(\d{1,2})分)?(?:\s*[〜~\-]\s*(\d{1,2})時(?:(\d{1,2})分)?)?/;
    const m = text.match(rx);
    if (m) {
      const ms = Number(m[1]);
      const ds = Number(m[2]);
      const hs = Number(m[3]);
      const mins = m[4] ? Number(m[4]) : 0;
      const he = m[5] ? Number(m[5]) : undefined;
      const mine = m[6] ? Number(m[6]) : 0;
      let year = currentYear;
      const startCandidate = buildDate(year, ms, ds, hs, mins, tz);
      if (startCandidate.getTime() < new Date().getTime() - 60 * 60 * 1000) {
        year += 1;
      }
      const start = buildDate(year, ms, ds, hs, mins, tz);
      const end =
        typeof he === "number"
          ? buildDate(year, ms, ds, he, mine, tz)
          : undefined;
      return {
        start,
        end: end && end > start ? end : undefined,
        isAllDay: false,
        sourceText: m[0],
      };
    }
  }

  // 4) 日付のみ（M/D or M月D日）→ 既定時刻(10:00)
  {
    const rx1 = /(\d{1,2})[\/\-\.](\d{1,2})(?:\s*\([^)]+\))?/;
    const rx2 = /(\d{1,2})月(\d{1,2})日(?:\s*\([^)]+\))?/;
    const m1 = text.match(rx1);
    const m2 = text.match(rx2);
    const m = m1 || m2;
    if (m) {
      const ms = Number(m[1]);
      const ds = Number(m[2]);
      let year = currentYear;
      const startCandidate = buildDate(year, ms, ds, 10, 0, tz);
      if (startCandidate.getTime() < new Date().getTime() - 60 * 60 * 1000) {
        year += 1;
      }
      const start = buildDate(year, ms, ds, 10, 0, tz);
      const end = buildDate(year, ms, ds, 11, 0, tz);
      return { start, end, isAllDay: false, sourceText: m[0] };
    }
  }

  return null;
}

function createCalendarEvent(schedule, meta, customerName, emailBody) {
  const calendar = getCalendar();
  const safeName = (customerName || "ご予約").replace(/[\s　]+/g, " ").trim();

  // メール本文からクーポン情報を抽出
  const couponInfo = parseCouponInfo(emailBody || "");

  // タイトルを氏名とクーポン内容で構成
  let title = `${safeName}さま予約`;
  if (couponInfo) {
    title = `${safeName}さま 予約 - ${couponInfo}`;
  }
  const descriptionLines = [
    `From: ${meta.from}`,
    `To: ${meta.to}`,
    meta.cc ? `Cc: ${meta.cc}` : "",
    `Received: ${formatJst(meta.receivedAt)}`,
    `Thread: ${meta.permalink}`,
    "",
    `Source: ${schedule.sourceText}`,
  ].filter(Boolean);
  const description = descriptionLines.join("\n");

  let event;
  if (schedule.isAllDay && schedule.end) {
    event = calendar.createAllDayEvent(title, schedule.start, schedule.end, {
      description,
      guests: "",
      sendInvites: false,
    });
  } else if (schedule.isAllDay) {
    event = calendar.createAllDayEvent(title, schedule.start, {
      description,
      guests: "",
      sendInvites: false,
    });
  } else {
    const start = schedule.start;
    const end =
      schedule.end && schedule.end > start
        ? schedule.end
        : new Date(start.getTime() + 60 * 60 * 1000);
    event = calendar.createEvent(title, start, end, {
      description,
      guests: "",
      sendInvites: false,
    });
  }

  // リマインダー: 前日メール＋10分前ポップアップ
  event.removeAllReminders();
  event.addEmailReminder(24 * 60);
  event.addPopupReminder(10);

  return event.getId();
}

function notifyLineNow(meta, schedule, customerName) {
  const safeName = (customerName || "ご予約").replace(/[\s　]+/g, " ").trim();
  const text = [
    "【新着予約】SALON BOARD",
    `お名前: ${safeName}さま`,
    `件名: ${meta.subject}`,
    `差出人: ${meta.from}`,
    `日程: ${formatJst(schedule.start)}${
      schedule.end ? ` - ${formatJst(schedule.end)}` : ""
    }`,
    `本文抜粋: ${meta.snippet}`,
    `Gmail: ${meta.permalink}`,
  ].join("\n");
  pushLineMessage({ type: "text", text });
}

function notifyLineFallback(meta) {
  const text = [
    "【新着メール】日程抽出に失敗しました",
    `件名: ${meta.subject}`,
    `差出人: ${meta.from}`,
    `本文抜粋: ${meta.snippet}`,
    `Gmail: ${meta.permalink}`,
  ].join("\n");
  pushLineMessage({ type: "text", text });
}

function scheduleReminder(schedule, ctx) {
  // トリガ数制限対策: 既存のsendReminderトリガを削除
  cleanupOldReminderTriggers();

  const remindAt = new Date(schedule.start.getTime() - 24 * 60 * 60 * 1000);
  const now = new Date();
  const fireAt =
    remindAt.getTime() > now.getTime() + 60 * 1000
      ? remindAt
      : new Date(now.getTime() + 2 * 60 * 1000);

  const trigger = ScriptApp.newTrigger("sendReminder")
    .timeBased()
    .at(fireAt)
    .create();

  // triggerUid をキーにコンテキストを保存
  const uid =
    typeof trigger.getUniqueId === "function" ? trigger.getUniqueId() : "";
  const storeKey = buildTriggerStoreKey(uid);
  const props = PropertiesService.getScriptProperties();
  props.setProperty(storeKey, JSON.stringify(ctx));
}

function sendReminder(e) {
  try {
    const uid = (e && e.triggerUid) || "";
    if (!uid) return;

    const storeKey = buildTriggerStoreKey(uid);
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(storeKey);
    if (!raw) return;

    const ctx = JSON.parse(raw);
    const title = ctx.title;
    const start = new Date(ctx.start);

    const text = [
      "【前日リマインド】",
      `タイトル: ${title}`,
      `開始: ${formatJst(start)}`,
      `Gmail: ${ctx.threadPermalink}`,
    ].join("\n");
    pushLineMessage({ type: "text", text });

    // 使い終わったら削除
    props.deleteProperty(storeKey);
  } catch (err) {
    console.error(err);
  }
}

/* ===== Helpers ===== */

function safePlainBody(msg) {
  const body = msg.getPlainBody();
  if (body && body.trim()) return body;
  // HTMLからテキスト抽出の簡易フォールバック
  const html = msg.getBody() || "";
  return html
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildMailMeta(thread, msg) {
  return {
    threadId: thread.getId(),
    messageId: String(msg.getId()),
    subject: msg.getSubject() || "(no subject)",
    from: msg.getFrom() || "",
    to: msg.getTo() || "",
    cc: msg.getCc() || "",
    snippet: (msg.getPlainBody() || "").split("\n").join(" "),
    permalink: buildGmailPermalink(thread.getId()),
    receivedAt: msg.getDate(),
  };
}

function buildGmailPermalink(threadId) {
  // 0番目アカウントの受信トレイスレッドリンク（必要に応じて u/1 等に変更）
  return `https://mail.google.com/mail/u/0/#all/${threadId}`;
}

function getCalendar() {
  const id = getCalendarId();
  return id
    ? CalendarApp.getCalendarById(id) || CalendarApp.getDefaultCalendar()
    : CalendarApp.getDefaultCalendar();
}

function getCalendarId() {
  const props = PropertiesService.getScriptProperties();
  const id =
    props.getProperty(PROP_KEYS.calendarId) || CONFIG_DEFAULTS.calendarId || "";
  return id || undefined;
}

function getTimezone() {
  const propTz = PropertiesService.getScriptProperties().getProperty(
    PROP_KEYS.timezone
  );
  return propTz || CONFIG_DEFAULTS.timezone || "Asia/Tokyo";
}

function formatJst(d) {
  return Utilities.formatDate(d, getTimezone(), "yyyy/MM/dd (E) HH:mm");
}

function buildDate(year, month, day, hour, minute, tz) {
  const d = new Date();
  d.setFullYear(year);
  d.setMonth(month - 1);
  d.setDate(day);
  d.setHours(hour, minute, 0, 0);
  return d;
}

function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

// Messaging API: broadcastメッセージ送信
function pushLineMessage(message) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_KEYS.channelAccessToken);
  if (!token) throw new Error("LINE_CHANNEL_ACCESS_TOKEN is not set");
  const url = "https://api.line.me/v2/bot/message/broadcast";
  const payload = { messages: [message] };
  const options = {
    method: "post",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
  };
  const res = fetchWithRetry(url, options, 3);
  const code = res.getResponseCode();
  if (code >= 300) {
    console.error(
      "LINE Messaging API broadcast failed",
      code,
      res.getContentText()
    );
  }
}

function fetchWithRetry(url, options, maxRetries) {
  let lastError = null;
  for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      return res;
    } catch (err) {
      lastError = err;
      const delayMs = Math.min(1000 * Math.pow(2, attempt), 8000);
      Utilities.sleep(delayMs);
    }
  }
  throw lastError;
}

function buildTriggerStoreKey(uid) {
  return `TRIGGER_CTX:${uid}`;
}

/* ===== Utility: 初回セットアップ補助 ===== */

function setupInitial() {
  getOrCreateLabel(LABELS.target);
  getOrCreateLabel(LABELS.processed);
  console.log("Labels ensured.");
}

// トリガクリーンアップ: 過去のsendReminderトリガを削除
function cleanupOldReminderTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const now = new Date().getTime();
  let deletedCount = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "sendReminder") {
      // 過去のトリガ（既に実行時刻を過ぎた）または古い（7日以上前に作成）トリガを削除
      const triggerTime =
        trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK
          ? new Date(trigger.getTriggerSourceId()).getTime()
          : 0;
      if (triggerTime < now || now - triggerTime > 7 * 24 * 60 * 60 * 1000) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }
  }

  if (deletedCount > 0) {
    console.log(`Cleaned up ${deletedCount} old reminder triggers`);
  }
}

// 手動でトリガを全削除（緊急時用）
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  console.log(`Deleted ${triggers.length} triggers`);
}

// Gmail フィルタ設定補助: SALON BOARDメール用の推奨フィルタ
function suggestGmailFilter() {
  const targetEmail = CONFIG_DEFAULTS.targetEmail;
  const filterCondition = `from:yoyaku_system@salonboard.com to:${targetEmail}`;
  const labelName = LABELS.target;

  console.log("=== Gmail フィルタ設定の推奨 ===");
  console.log("1. Gmailで「設定」→「フィルタとブロック中のアドレス」");
  console.log("2. 「新しいフィルタを作成」");
  console.log(`3. 検索条件: ${filterCondition}`);
  console.log(`4. アクション: ラベル「${labelName}」を付ける（新規作成）`);
  console.log("5. 「○○件の一致するスレッドにもフィルタを適用する」をチェック");
  console.log(
    "=== これにより、対象メールが自動でラベル付けされ、検索が高速化されます ==="
  );
}

// トリガ設定補助: 推奨トリガ設定
function suggestTriggerSetup() {
  console.log("=== 推奨トリガ設定 ===");
  console.log("1. GASエディタで「トリガー」タブを開く");
  console.log("2. 「トリガーを追加」");
  console.log("3. 実行する関数: pollEmails");
  console.log("4. イベントのソース: 時間主導型");
  console.log("5. 時間ベースのトリガーのタイプ: 分タイマー");
  console.log("6. 時間の間隔を選択: 1分おき（または2分おき）");
  console.log("=== 予約メールは即座に処理されるため、1-2分間隔で十分です ===");
}

/* ===== 手動テスト用 ===== */

function debugParse() {
  const tz = getTimezone();
  const samples = [
    "■来店日時\n　2025年08月25日（月）14:30",
    "■来店日時\n2025/08/12 14:30 面談のご案内",
    "■来店日時\n8/12(火) 14時〜15時",
    "■来店日時\n8月12日 14:30",
    "予約受付日時：2025年08月11日（月）17:19",
  ];
  for (const s of samples) {
    const sch = parseSchedule(s, tz);
    console.log(
      s,
      "=>",
      sch
        ? `${formatJst(sch.start)}${sch.end ? " - " + formatJst(sch.end) : ""}`
        : "null"
    );
  }
}

// 手動診断: GAS から外部URLへ到達できるかのヘルスチェック
function pingExternal() {
  const res = UrlFetchApp.fetch("https://httpbin.org/get", {
    muteHttpExceptions: true,
  });
  console.log(res.getResponseCode(), res.getContentText().slice(0, 200));
}

// LINE Notify ステータスAPIで到達性を確認
function pingLineMessagingApi() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(PROP_KEYS.channelAccessToken);
  if (!token) throw new Error("LINE_CHANNEL_ACCESS_TOKEN is not set");
  const res = UrlFetchApp.fetch("https://api.line.me/v2/bot/info", {
    method: "get",
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
  });
  console.log(res.getResponseCode(), res.getContentText());
}

// 送信テスト（承認フローの強制）
function testLineNotify() {
  sendLine("LINE Notify connection OK");
}

// 予約メール本文から「■氏名」の次行を抽出
function parseCustomerName(text) {
  const raw = extractFieldAfter(text, "■氏名");
  if (!raw) return null;
  const name = trimJaSpaces(raw)
    .replace(/[\s\u3000]+/g, " ")
    .trim();
  return name || null;
}

// 予約メール本文から「■ご利用クーポン」の内容を抽出（クーポン名と料金のみ）
function parseCouponInfo(text) {
  const lines = text.split(/\r?\n/);
  let foundCoupon = false;

  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes("■ご利用クーポン")) {
      foundCoupon = true;
      // 同じ行にクーポン情報がある場合
      const inline = lines[i].replace(/^.*?[:：]\s*/, "").trim();
      if (inline && inline !== "■ご利用クーポン") {
        return extractCouponNameAndPrice(inline);
      }
      continue;
    }

    if (foundCoupon) {
      const line = trimJaSpaces(lines[i]);
      // 次のセクション（■で始まる）が来たら終了
      if (/^■/.test(line)) {
        break;
      }
      // 最初の非空行のみを処理（クーポン名と料金が含まれる行）
      if (line) {
        return extractCouponNameAndPrice(line);
      }
    }
  }

  return null;
}

// クーポン名と料金部分のみを抽出するヘルパー関数
function extractCouponNameAndPrice(line) {
  // ★で囲まれたクーポン名と料金部分を抽出
  // 例: ★人気No.2★パーソナルカラー診断+顔診断+骨格診断　¥29800→¥17000
  const match = line.match(/★([^★]+)★([^¥]*¥[^¥]*¥[^\s\u3000]+)/);
  if (match) {
    return `★${match[1]}★${match[2]}`.trim();
  }

  // ★がない場合は、料金部分（¥記号を含む）までを抽出
  const priceMatch = line.match(/^([^¥]*¥[^¥]*¥[^\s\u3000]+)/);
  if (priceMatch) {
    return priceMatch[1].trim();
  }

  // 料金が見つからない場合は、最初の文や単語のみを返す
  const firstPart = line.split(/[\u3000\s]{2,}/)[0]; // 2つ以上の空白で区切る
  return firstPart ? firstPart.trim() : line.trim();
}

// ラベル行の直後の非空行を取り出す（値が同一行にある場合も考慮）
function extractFieldAfter(text, label) {
  const lines = text.split(/\r?\n/);
  for (let i = 0; i < lines.length; i += 1) {
    if (lines[i].includes(label)) {
      const inline = lines[i].replace(/^.*?[:：]\s*/, "").trim();
      if (inline && inline !== label) return inline;
      // 次の非空行
      for (let j = i + 1; j < lines.length; j += 1) {
        const cand = trimJaSpaces(lines[j]);
        if (cand) {
          // 次のセクション見出し（■始まり）は値ではない
          if (/^■/.test(cand)) break;
          return cand;
        }
      }
    }
  }
  return null;
}

function trimJaSpaces(s) {
  return (s || "").replace(/[\u3000\s]+$/g, "").replace(/^[\u3000\s]+/g, "");
}
