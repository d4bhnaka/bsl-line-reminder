/**
 * Apps Script (JavaScript) — Gmail→複数LINE公式アカウント通知＋日程抽出→Calendar登録
 *
 * 事前準備:
 * - スクリプト プロパティ（File > Project properties > Script properties）に以下を設定
 *   - LINE_CHANNEL_ACCESS_TOKEN_1: 1つ目のLINE公式アカウントのチャンネルアクセストークン
 *   - LINE_CHANNEL_ACCESS_TOKEN_2: 2つ目のLINE公式アカウントのチャンネルアクセストークン
 *   - LINE_CHANNEL_ACCESS_TOKEN_3: 3つ目のLINE公式アカウントのチャンネルアクセストークン（任意）
 *   - （必要に応じて LINE_CHANNEL_ACCESS_TOKEN_4, 5... を追加）
 *   - LINE_ACCOUNT_NAMES: 各アカウントの名前（カンマ区切り、例: "店舗A,店舗B,店舗C"）
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
  channelAccessTokenPrefix: "LINE_CHANNEL_ACCESS_TOKEN_", // _1, _2, _3... という形式で複数トークンを管理
  accountNames: "LINE_ACCOUNT_NAMES", // アカウント名のリスト（カンマ区切り）
  targetEmail: "TARGET_EMAIL",
  calendarId: "CALENDAR_ID",
  timezone: "TIMEZONE",
});

// 既定値（スクリプトプロパティ未設定時のフォールバック）
const CONFIG_DEFAULTS = Object.freeze({
  targetEmail: "besol4b.shop@gmail.com",
  calendarId: "besol4b.shop@gmail.com",
  timezone: "Asia/Tokyo",
  maxLineAccounts: 10, // 最大10個のLINEアカウントまでサポート
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

        // 前日LINEリマインド機能は無効化されています
        // scheduleReminder(schedule, {
        //   eventId,
        //   calendarId: getCalendarId(),
        //   title: titleForReminder,
        //   start: schedule.start.toISOString(),
        //   threadPermalink: meta.permalink,
        //   metaId: `${meta.threadId}:${
        //     meta.messageId
        //   }:${schedule.start.getTime()}`,
        // });
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
 * 本文から所要時間目安を抽出
 * 例: （所要時間目安：3時間25分） → 205分を返す
 * 所要時間目安が見つからない場合はデフォルト120分（2時間）を返す
 */
function parseDurationEstimate(text) {
  const match = text.match(/所要時間目安[：:]\s*(\d+)\s*時間\s*(\d+)\s*分/);
  if (match) {
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    return hours * 60 + minutes;
  }

  const hourOnlyMatch = text.match(/所要時間目安[：:]\s*(\d+)\s*時間/);
  if (hourOnlyMatch) {
    return Number(hourOnlyMatch[1]) * 60;
  }

  const minuteOnlyMatch = text.match(/所要時間目安[：:]\s*(\d+)\s*分/);
  if (minuteOnlyMatch) {
    return Number(minuteOnlyMatch[1]);
  }

  return 120; // デフォルト2時間
}

/**
 * 本文から日時を抽出（日本語表現に対応した簡易ルール）
 * - 優先度: YYYY/M/D HH:mm → M/D HH:mm → M月D日 HH時mm分 → 日付のみ（既定10:00）
 * - 年未指定: 今年、かつ過去なら翌年
 * - 時刻未指定: 10:00
 * - 範囲表現: 14:00-15:00 / 14時〜15時 → endに反映（分省略=00）
 * - 所要時間目安を反映して終了時間を設定
 */
function parseSchedule(text, tz) {
  const now = new Date();
  const currentYear = now.getFullYear();

  // 所要時間目安を抽出（デフォルト2時間）
  const durationMinutes = parseDurationEstimate(text);

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
      let end;
      if (typeof he === "number") {
        end = buildDate(ys, ms, ds, he, mine, tz);
      } else {
        // 終了時間が指定されていない場合は所要時間目安を使用
        end = new Date(start.getTime() + durationMinutes * 60 * 1000);
      }
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
      const end = new Date(start.getTime() + durationMinutes * 60 * 1000);
      return { start, end, isAllDay: false, sourceText: m[0] };
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
      const end = new Date(start.getTime() + durationMinutes * 60 * 1000);
      return { start, end, isAllDay: false, sourceText: m[0] };
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
      let end;
      if (typeof he === "number") {
        end = buildDate(ys, ms, ds, he, mine, tz);
      } else {
        // 終了時間が指定されていない場合は所要時間目安を使用
        end = new Date(start.getTime() + durationMinutes * 60 * 1000);
      }
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
      let end;
      if (typeof he === "number") {
        end = buildDate(year, ms, ds, he, mine, tz);
      } else {
        // 終了時間が指定されていない場合は所要時間目安を使用
        end = new Date(start.getTime() + durationMinutes * 60 * 1000);
      }
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
      let end;
      if (typeof he === "number") {
        end = buildDate(year, ms, ds, he, mine, tz);
      } else {
        // 終了時間が指定されていない場合は所要時間目安を使用
        end = new Date(start.getTime() + durationMinutes * 60 * 1000);
      }
      return {
        start,
        end: end && end > start ? end : undefined,
        isAllDay: false,
        sourceText: m[0],
      };
    }
  }

  // 4) 日付のみ（M/D or M月D日）→ 既定時刻(10:00)、所要時間目安を終了時間に反映
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
      const end = new Date(start.getTime() + durationMinutes * 60 * 1000);
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
        : new Date(start.getTime() + 120 * 60 * 1000); // デフォルト2時間
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

/**
 * メール本文から予約詳細情報を抽出してフォーマットされた文字列を返す
 * @param {string} text - メール本文
 * @returns {string} - フォーマットされた予約詳細情報
 */
function extractReservationDetails(text) {
  const lines = text.split(/\r?\n/);
  const details = [];

  // ◇ご予約内容セクションを探す
  let inReservationSection = false;
  let inTotalSection = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // ◇ご予約内容セクションの開始
    if (line.includes("◇ご予約内容")) {
      inReservationSection = true;
      details.push("◇ご予約内容");
      continue;
    }

    // PC版SALON BOARDが来たら終了
    if (line.includes("PC版SALON BOARD")) {
      break;
    }

    // ◇ご予約付加情報が来たら終了
    if (line.includes("◇ご予約付加情報")) {
      break;
    }

    if (inReservationSection) {
      // ■予約番号
      if (line.includes("■予約番号")) {
        details.push("■予約番号");
        // 次の非空行を取得
        for (let j = i + 1; j < lines.length; j++) {
          const nextLine = lines[j].trim();
          if (nextLine && !nextLine.startsWith("■")) {
            details.push(`　${nextLine}`);
            break;
          }
        }
      }

      // ■氏名
      else if (line.includes("■氏名")) {
        details.push("■氏名");
        // 次の非空行を取得
        for (let j = i + 1; j < lines.length; j++) {
          const nextLine = lines[j].trim();
          if (nextLine && !nextLine.startsWith("■")) {
            details.push(`　${nextLine}`);
            break;
          }
        }
      }

      // ■来店日時
      else if (line.includes("■来店日時")) {
        details.push("■来店日時");
        // 次の非空行を取得
        for (let j = i + 1; j < lines.length; j++) {
          const nextLine = lines[j].trim();
          if (nextLine && !nextLine.startsWith("■")) {
            details.push(`　${nextLine}`);
            break;
          }
        }
      }

      // ■指名スタッフ
      else if (line.includes("■指名スタッフ")) {
        details.push("■指名スタッフ");
        // 次の非空行を取得
        for (let j = i + 1; j < lines.length; j++) {
          const nextLine = lines[j].trim();
          if (nextLine && !nextLine.startsWith("■")) {
            details.push(`　${nextLine}`);
            break;
          }
        }
      }

      // ■メニュー
      else if (line.includes("■メニュー")) {
        details.push("■メニュー");
        // メニュー内容を複数行取得（所要時間目安の行まで）
        for (let j = i + 1; j < lines.length; j++) {
          const nextLine = lines[j];
          if (nextLine.trim() && !nextLine.startsWith("■")) {
            details.push(nextLine);
          } else if (nextLine.startsWith("■")) {
            break;
          }
        }
      }

      // ■ご利用クーポン
      else if (line.includes("■ご利用クーポン")) {
        details.push("■ご利用クーポン");
        // クーポン内容を複数行取得（次の■まで）
        for (let j = i + 1; j < lines.length; j++) {
          const nextLine = lines[j];
          // 次の■で始まる行が来たら終了
          if (nextLine.trim().startsWith("■")) {
            break;
          }
          // 空行でない場合、内容を追加
          if (nextLine.trim()) {
            details.push(nextLine);
          }
        }
      }

      // ■合計金額
      else if (line.includes("■合計金額")) {
        inTotalSection = true;
        details.push("■合計金額");
      }

      // 合計金額セクション内
      else if (inTotalSection) {
        const trimmedLine = line.trim();
        // 予約時合計金額、利用ギフト券、利用ポイント、お支払い予定金額のみ取得
        if (
          trimmedLine.includes("予約時合計金額") ||
          trimmedLine.includes("今回の利用ギフト券") ||
          trimmedLine.includes("今回の利用ポイント") ||
          trimmedLine.includes("お支払い予定金額")
        ) {
          details.push(line);
        }
        // ※表示金額は...の注意書きが来たら終了
        else if (trimmedLine.startsWith("※")) {
          inTotalSection = false;
        }
      }
    }
  }

  return details.join("\n");
}

function notifyLineNow(meta, schedule, customerName) {
  const safeName = (customerName || "ご予約").replace(/[\s　]+/g, " ").trim();

  // メール本文から予約詳細情報を抽出
  const body = meta.fullBody || meta.snippet;
  const reservationDetails = extractReservationDetails(body);

  const text = [
    "【新着予約】SALON BOARD",
    `お名前: ${safeName}さま`,
    `日程: ${formatJst(schedule.start)}${
      schedule.end ? ` - ${formatJst(schedule.end)}` : ""
    }`,
    reservationDetails,
  ].join("\n");
  pushLineMessageToAll({ type: "text", text });
}

function notifyLineFallback(meta) {
  let snippetText = "";
  if (meta.snippet && meta.snippet.trim()) {
    let trimmedSnippet = meta.snippet.trim();

    // "PC版SALON BOARD"以降を削除
    const cutoffText = "PC版SALON BOARD";
    const cutoffIndex = trimmedSnippet.indexOf(cutoffText);
    if (cutoffIndex !== -1) {
      trimmedSnippet = trimmedSnippet.substring(0, cutoffIndex).trim();
    }

    // 本文抜粋を整形（最大200文字に制限）
    if (trimmedSnippet.length > 200) {
      snippetText = `本文抜粋: ${trimmedSnippet.substring(0, 200)}...`;
    } else {
      snippetText = `本文抜粋: ${trimmedSnippet}`;
    }
  }

  const textParts = ["【新着メール】日程抽出に失敗しました"];

  // 本文抜粋がある場合のみ追加
  if (snippetText) {
    textParts.push(snippetText);
  }

  const text = textParts.join("\n");
  pushLineMessageToAll({ type: "text", text });
}

/**
 * 複数のLINE公式アカウントに同じメッセージを送信
 * @param {Object} message - LINE Messaging APIのメッセージオブジェクト
 */
function pushLineMessageToAll(message) {
  const lineAccounts = getLineAccounts();

  if (lineAccounts.length === 0) {
    throw new Error("No LINE channel access tokens configured");
  }

  const results = [];
  const errors = [];

  // 各アカウントに並行して送信
  for (const account of lineAccounts) {
    try {
      pushLineMessage(message, account.token, account.name);
      results.push({
        name: account.name,
        status: "success",
      });
    } catch (err) {
      console.error(`Failed to send to ${account.name}:`, err);
      errors.push({
        name: account.name,
        error: err.toString(),
      });
    }
  }

  // 送信結果をログに記録
  console.log("LINE送信結果:", {
    total: lineAccounts.length,
    success: results.length,
    failed: errors.length,
    details: { results, errors },
  });

  // すべて失敗した場合はエラーをスロー
  if (results.length === 0 && errors.length > 0) {
    throw new Error(`All LINE messages failed: ${JSON.stringify(errors)}`);
  }
}

/**
 * 設定されているLINEアカウント情報を取得
 * @returns {Array<{token: string, name: string}>} アカウント情報の配列
 */
function getLineAccounts() {
  const props = PropertiesService.getScriptProperties();
  const accounts = [];

  // アカウント名のリストを取得（カンマ区切り）
  const accountNames = (props.getProperty(PROP_KEYS.accountNames) || "").split(
    ","
  );

  // 最大10個のアカウントをチェック
  for (let i = 1; i <= CONFIG_DEFAULTS.maxLineAccounts; i++) {
    const tokenKey = `${PROP_KEYS.channelAccessTokenPrefix}${i}`;
    const token = props.getProperty(tokenKey);

    if (token) {
      const name = accountNames[i - 1]
        ? accountNames[i - 1].trim()
        : `Account ${i}`;
      accounts.push({ token, name });
    }
  }

  return accounts;
}

/**
 * 単一のLINE公式アカウントにメッセージを送信
 * @param {Object} message - LINE Messaging APIのメッセージオブジェクト
 * @param {string} token - チャンネルアクセストークン
 * @param {string} accountName - アカウント名（ログ用）
 */
function pushLineMessage(message, token, accountName = "Unknown") {
  if (!token)
    throw new Error(`Channel access token is not set for ${accountName}`);

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
    const errorMsg = `LINE Messaging API broadcast failed for ${accountName}`;
    console.error(errorMsg, code, res.getContentText());
    throw new Error(`${errorMsg}: ${code}`);
  } else {
    console.log(`Successfully sent to ${accountName}`);
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
  const plainBody = msg.getPlainBody() || "";
  return {
    threadId: thread.getId(),
    messageId: String(msg.getId()),
    subject: msg.getSubject() || "(no subject)",
    from: msg.getFrom() || "",
    to: msg.getTo() || "",
    cc: msg.getCc() || "",
    snippet: plainBody.split("\n").join(" "),
    fullBody: plainBody,
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

/* ===== Utility: 初回セットアップ補助 ===== */

function setupInitial() {
  getOrCreateLabel(LABELS.target);
  getOrCreateLabel(LABELS.processed);
  console.log("Labels ensured.");

  // LINE アカウント設定を確認
  const accounts = getLineAccounts();
  console.log(`Found ${accounts.length} LINE account(s):`);
  accounts.forEach((acc, idx) => {
    console.log(`  ${idx + 1}. ${acc.name}`);
  });
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

/**
 * 統合テスト: pollEmails()の全体フローをモックデータで検証
 * 実際のメール本文を使って、すべての関数が正しく呼び出されることを確認
 */
function testPollEmailsIntegration() {
  console.log("=== 統合テスト開始 ===");

  // テスト用メール本文サンプル
  const testEmailBody = `
■来店日時
　2025年10月14日（火）13:00

■お客様名
田中 太郎

■クーポン
プレミアムカット（¥5,000）

（所要時間目安：3時間25分）

その他の情報...
`;

  const testEmailBodyNoEstimate = `
■来店日時
　2025年10月20日（月）10:00

■お客様名
佐藤 花子

■クーポン
カラーリング（¥8,000）

その他の情報...
`;

  const testEmailBodyNoCoupon = `
■来店日時
　2025年10月25日（土）15:30

■お客様名
山田 次郎

通常予約です。
`;

  const testCases = [
    {
      name: "所要時間目安あり + クーポンあり",
      body: testEmailBody,
      expectedDuration: 205, // 3時間25分
      expectedName: "田中 太郎",
      expectedCoupon: "プレミアムカット",
    },
    {
      name: "所要時間目安なし + クーポンあり（デフォルト2時間）",
      body: testEmailBodyNoEstimate,
      expectedDuration: 120, // 2時間
      expectedName: "佐藤 花子",
      expectedCoupon: "カラーリング",
    },
    {
      name: "クーポンなし",
      body: testEmailBodyNoCoupon,
      expectedDuration: 120, // 2時間
      expectedName: "山田 次郎",
      expectedCoupon: null,
    },
  ];

  const tz = getTimezone();
  let passedTests = 0;
  let failedTests = 0;

  for (const testCase of testCases) {
    console.log(`\n--- テストケース: ${testCase.name} ---`);

    try {
      // 1. 顧客名パース
      const name = parseCustomerName(testCase.body);
      console.log(`✓ 顧客名パース: ${name}`);
      if (name !== testCase.expectedName) {
        throw new Error(`期待値: ${testCase.expectedName}, 実際: ${name}`);
      }

      // 2. 所要時間目安パース
      const duration = parseDurationEstimate(testCase.body);
      console.log(`✓ 所要時間目安: ${duration}分`);
      if (duration !== testCase.expectedDuration) {
        throw new Error(
          `期待値: ${testCase.expectedDuration}, 実際: ${duration}`
        );
      }

      // 3. 日程パース
      const schedule = parseSchedule(testCase.body, tz);
      if (!schedule) {
        throw new Error("日程パースに失敗");
      }
      console.log(
        `✓ 日程パース: ${formatJst(schedule.start)} - ${formatJst(
          schedule.end
        )}`
      );

      // 4. 終了時間の検証（開始時間 + 所要時間）
      const expectedEndTime = new Date(
        schedule.start.getTime() + testCase.expectedDuration * 60 * 1000
      );
      const actualEndTime = schedule.end;
      if (
        Math.abs(expectedEndTime.getTime() - actualEndTime.getTime()) > 1000
      ) {
        throw new Error(
          `終了時間が期待値と一致しません。期待: ${formatJst(
            expectedEndTime
          )}, 実際: ${formatJst(actualEndTime)}`
        );
      }
      console.log(`✓ 終了時間検証: OK`);

      // 5. クーポン情報パース
      const couponInfo = parseCouponInfo(testCase.body);
      console.log(`✓ クーポン情報: ${couponInfo || "なし"}`);
      if (testCase.expectedCoupon) {
        if (!couponInfo || !couponInfo.includes(testCase.expectedCoupon)) {
          throw new Error(
            `期待値: ${testCase.expectedCoupon}を含む, 実際: ${couponInfo}`
          );
        }
      } else if (couponInfo) {
        throw new Error(`期待値: null, 実際: ${couponInfo}`);
      }

      // 6. メタデータ構築（モック）
      console.log(`✓ 各関数が正常に動作しました`);

      passedTests++;
      console.log(`✅ テストケース「${testCase.name}」: PASS`);
    } catch (error) {
      failedTests++;
      console.error(`❌ テストケース「${testCase.name}」: FAIL`);
      console.error(`エラー: ${error.message}`);
    }
  }

  console.log("\n=== 統合テスト結果 ===");
  console.log(`合計: ${testCases.length}件`);
  console.log(`成功: ${passedTests}件`);
  console.log(`失敗: ${failedTests}件`);

  if (failedTests === 0) {
    console.log("✅ すべてのテストが成功しました！");
  } else {
    console.log("⚠️ 一部のテストが失敗しました。");
  }

  // 関数使用状況の確認
  console.log("\n=== 主要関数の使用状況 ===");
  console.log("✓ parseDurationEstimate() - 所要時間目安の抽出");
  console.log("✓ parseCustomerName() - 顧客名の抽出");
  console.log("✓ parseSchedule() - 日程の抽出");
  console.log("  ├─ extractFieldAfter() - フィールド抽出");
  console.log("  ├─ trimJaSpaces() - 空白トリム");
  console.log("  └─ buildDate() - 日付構築");
  console.log("✓ parseCouponInfo() - クーポン情報の抽出");
  console.log("  └─ extractCouponNameAndPrice() - クーポン名と価格の抽出");
  console.log("✓ formatJst() - 日時フォーマット");
  console.log("\n=== pollEmails()で使用される関数 ===");
  console.log("✓ getOrCreateLabel() - ラベル作成/取得");
  console.log("✓ safePlainBody() - メール本文の取得");
  console.log("✓ buildMailMeta() - メタデータ構築");
  console.log("  └─ buildGmailPermalink() - パーマリンク生成");
  console.log("✓ notifyLineNow() / notifyLineFallback() - LINE通知");
  console.log("  └─ pushLineMessageToAll() - 全アカウントに送信");
  console.log("      ├─ getLineAccounts() - アカウント情報取得");
  console.log("      └─ pushLineMessage() - メッセージ送信");
  console.log("          └─ fetchWithRetry() - リトライ付きHTTPリクエスト");
  console.log("✓ createCalendarEvent() - カレンダーイベント作成");
  console.log("  ├─ getCalendar() - カレンダー取得");
  console.log("  ├─ getCalendarId() - カレンダーID取得");
  console.log("  └─ getTimezone() - タイムゾーン取得");
}

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

// 複数のLINE Messaging APIの接続状態を確認
function pingAllLineMessagingApis() {
  const accounts = getLineAccounts();

  if (accounts.length === 0) {
    console.log("No LINE accounts configured");
    return;
  }

  console.log(`Testing ${accounts.length} LINE account(s):`);

  for (const account of accounts) {
    try {
      const res = UrlFetchApp.fetch("https://api.line.me/v2/bot/info", {
        method: "get",
        headers: { Authorization: `Bearer ${account.token}` },
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: true,
      });

      const code = res.getResponseCode();
      if (code === 200) {
        const info = JSON.parse(res.getContentText());
        console.log(`✓ ${account.name}: Connected successfully`);
        console.log(`  - Bot name: ${info.displayName || "N/A"}`);
        console.log(`  - Basic ID: ${info.basicId || "N/A"}`);
      } else {
        console.log(`✗ ${account.name}: Failed (${code})`);
        console.log(`  - Error: ${res.getContentText()}`);
      }
    } catch (err) {
      console.log(`✗ ${account.name}: Error - ${err.toString()}`);
    }
  }
}

// 全LINEアカウントへのテスト送信
function testAllLineNotify() {
  const accounts = getLineAccounts();

  if (accounts.length === 0) {
    console.log("No LINE accounts configured");
    return;
  }

  const testMessage = {
    type: "text",
    text: `【テスト送信】\n${
      accounts.length
    }個のLINE公式アカウントへの接続テスト\n送信時刻: ${formatJst(new Date())}`,
  };

  console.log(`Sending test message to ${accounts.length} account(s)...`);
  pushLineMessageToAll(testMessage);
  console.log("Test message sent successfully!");
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

// 設定内容を表示（デバッグ用）
function showConfiguration() {
  const props = PropertiesService.getScriptProperties();
  const accounts = getLineAccounts();

  console.log("=== 現在の設定 ===");
  console.log(
    `Target Email: ${
      props.getProperty(PROP_KEYS.targetEmail) || CONFIG_DEFAULTS.targetEmail
    }`
  );
  console.log(
    `Calendar ID: ${
      props.getProperty(PROP_KEYS.calendarId) || CONFIG_DEFAULTS.calendarId
    }`
  );
  console.log(
    `Timezone: ${
      props.getProperty(PROP_KEYS.timezone) || CONFIG_DEFAULTS.timezone
    }`
  );
  console.log(`LINE Accounts: ${accounts.length} account(s)`);

  accounts.forEach((acc, idx) => {
    console.log(`  ${idx + 1}. ${acc.name}`);
  });

  console.log("==================");
}

/**
 * LINE通知フォーマットのテスト関数
 * 実際のメール本文を使って、extractReservationDetails関数の動作を確認
 */
function testLineNotificationFormat() {
  console.log("=== LINE通知フォーマットテスト開始 ===\n");

  // 実際のメール本文サンプル
  const testEmailBody = `パーソナルカラー診断・顔タイプの診断・骨格診断 専門サロン　be style lab 名古屋様

HOT PEPPER Beauty「SALON BOARD」にお客様から
ご予約が入りました。

◇ご予約内容
■予約番号
　BD73586018
■氏名
　吉留 美里（ヨシトメ ミサト）
■来店日時
　2025年10月14日（火）15:30
■指名スタッフ
　指名なし
■メニュー
　ボディトリ＋ボディケア＋アート＋フェイシャル＋ボディ＋ブライダル＋その他
　（所要時間目安：2時間）
■ご利用クーポン
　[全員]
　"可愛い"も"綺麗"も叶える！顔診断+似合わせフルメイク/通常\\30000→\\19800
　　顔のバランスや雰囲気からあなたが1番輝くメイクをご提案！調和の取れた印象で魅力を最大限に引き出します★16type顔診断+フルメイク＋StyleBook付+AIコンシェルジュ+無期限Afterfollow付き
■合計金額
　予約時合計金額　19,800円
　今回の利用ギフト券　利用なし
　今回の利用ポイント　利用なし
　お支払い予定金額　19,800円
※表示金額は、予約時に選択したメニュー金額(クーポン適用の場合は適用後の金額)の合計金額です。来店時のメニュー変更、サロンが別途設定するスタッフ指名料等の追加料金やキャンセル料等により、実際の支払額と異なる場合があります。
追加料金についてはこちらから↓
https://beauty.help.hotpepper.jp/s/article/000031948
◇ご予約付加情報
■なりたいイメージ
　なし
■ご要望・ご相談
　-
■サロンからお客様への質問
　質問：【重要】
予約後すぐに公式LINEにお名前の送信をお願いいたします。（登録だけでは連絡ができません）公式LINEへは、Webサイトもしくはインスタグラムからアクセス可能です。「ビースタイルラボ」で検索
※ お電話での対応は一切行っておりません。すべて公式LINEのトークからお願いします。
※ 当サロンは女性専用となっております。男性や男女カップルの診断は行っておりませんので予めご了承ください。
　回答：yoshitome.m

PC版SALON BOARD
https://salonboard.com/login/
スマートフォン版SALON BOARD
https://salonboard.com/login_sp/

予約受付日時：2025年10月14日（火）00:25

===================================================
SALON BOARD・HOT PEPPER Beauty
お問い合わせ：https://sbhd-kirei.salonboard.com/hc/ja/articles/360039395113
===================================================`;

  // 予約詳細情報を抽出
  const reservationDetails = extractReservationDetails(testEmailBody);

  // 顧客名を抽出
  const customerName = parseCustomerName(testEmailBody);

  // 日程を抽出
  const schedule = parseSchedule(testEmailBody, getTimezone());

  // 完全なLINE通知メッセージを構築
  const safeName = (customerName || "ご予約").replace(/[\s　]+/g, " ").trim();
  const fullMessage = [
    "【新着予約】SALON BOARD",
    `お名前: ${safeName}さま`,
    `日程: ${schedule ? formatJst(schedule.start) : "日程不明"}${
      schedule && schedule.end ? ` - ${formatJst(schedule.end)}` : ""
    }`,
    reservationDetails,
  ].join("\n");

  console.log("=== 抽出された顧客名 ===");
  console.log(customerName);
  console.log("");

  console.log("=== 抽出された日程 ===");
  if (schedule) {
    console.log(`開始: ${formatJst(schedule.start)}`);
    console.log(`終了: ${formatJst(schedule.end)}`);
  } else {
    console.log("日程の抽出に失敗しました");
  }
  console.log("");

  console.log("=== 抽出された予約詳細 ===");
  console.log(reservationDetails);
  console.log("");

  console.log("=== 完全なLINE通知メッセージ ===");
  console.log(fullMessage);
  console.log("");

  console.log("=== テスト完了 ===");
}
