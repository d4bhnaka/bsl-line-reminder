/**
 * Apps Script (JavaScript) - Shopping Email to LINE Notification
 *
 * Setup:
 * - Script Properties (File > Project properties > Script properties):
 *   - LINE_CHANNEL_ACCESS_TOKEN_1: LINE channel access token #1
 *   - LINE_CHANNEL_ACCESS_TOKEN_2: LINE channel access token #2
 *   - LINE_CHANNEL_ACCESS_TOKEN_3: LINE channel access token #3 (optional)
 *   - (Add LINE_CHANNEL_ACCESS_TOKEN_4, 5... as needed)
 *   - LINE_ACCOUNT_NAMES: Account names (comma-separated)
 *   - SHOPPING_TARGET_EMAIL: Target email address (default: besol4b.shop@gmail.com)
 *   - TIMEZONE: Asia/Tokyo (default if not set)
 * - Gmail Filter: from:noreply@cromi.co to:besol4b.shop@gmail.com subject:shopping -> auto-apply label 'shopping-emails'
 * - Trigger: Time-driven (every 1-2 minutes) to run pollShoppingEmails
 */

const SHOPPING_LABELS = Object.freeze({
  target: "shopping-emails",
  processed: "processed-shopping",
});

const SHOPPING_PROP_KEYS = Object.freeze({
  channelAccessTokenPrefix: "LINE_CHANNEL_ACCESS_TOKEN_",
  accountNames: "LINE_ACCOUNT_NAMES",
  targetEmail: "SHOPPING_TARGET_EMAIL",
  timezone: "TIMEZONE",
});

const SHOPPING_CONFIG_DEFAULTS = Object.freeze({
  targetEmail: "besol4b.shop@gmail.com",
  timezone: "Asia/Tokyo",
  maxLineAccounts: 10,
});

/**
 * Main function: Monitor shopping emails and notify LINE
 */
function pollShoppingEmails() {
  const props = PropertiesService.getScriptProperties();
  const targetEmail = (
    props.getProperty(SHOPPING_PROP_KEYS.targetEmail) ||
    SHOPPING_CONFIG_DEFAULTS.targetEmail ||
    ""
  ).toLowerCase();

  getOrCreateShoppingLabel(SHOPPING_LABELS.target);
  getOrCreateShoppingLabel(SHOPPING_LABELS.processed);

  // Search query for shopping emails
  // From: noreply@cromi.co
  // To: besol4b.shop@gmail.com
  // Subject: shopping
  const query = [
    "from:noreply@cromi.co",
    targetEmail ? `to:${targetEmail}` : "",
    "subject:ショッピング同行",
    `-label:${SHOPPING_LABELS.processed}`,
    "is:unread",
    "newer_than:7d",
  ]
    .filter(Boolean)
    .join(" ");

  const threads = GmailApp.search(query, 0, 30);

  for (const thread of threads) {
    const messages = thread.getMessages().filter((m) => m.isUnread());
    for (const msg of messages) {
      // Filter by recipient (To/CC)
      const to = (msg.getTo() || "").toLowerCase();
      const cc = (msg.getCc() || "").toLowerCase();
      if (
        targetEmail &&
        !(to.includes(targetEmail) || cc.includes(targetEmail))
      ) {
        continue;
      }

      // Get email info
      const meta = buildShoppingMailMeta(thread, msg);
      const body = getShoppingMailBody(msg);

      // Send LINE notification
      notifyShoppingLine(meta, body);
    }

    // Mark as processed
    thread.addLabel(GmailApp.getUserLabelByName(SHOPPING_LABELS.processed));
    thread.markRead();
  }
}

/**
 * Send LINE notification for shopping email
 */
function notifyShoppingLine(meta, body) {
  // Format email body (max 1000 chars)
  let bodyText = body.trim();
  if (bodyText.length > 1000) {
    bodyText = bodyText.substring(0, 1000) + "...";
  }

  const text = [
    "【ショッピング同行】新着メール",
    `件名: ${meta.subject}`,
    `差出人: ${meta.from}`,
    `受信日時: ${formatShoppingJst(meta.receivedAt)}`,
    "",
    "--- メール本文 ---",
    bodyText,
    "",
    `Gmail: ${meta.permalink}`,
  ].join("\n");

  pushShoppingLineMessageToAll({ type: "text", text });
}

/**
 * Send message to all LINE accounts
 */
function pushShoppingLineMessageToAll(message) {
  const lineAccounts = getShoppingLineAccounts();

  if (lineAccounts.length === 0) {
    throw new Error("No LINE channel access tokens configured");
  }

  const results = [];
  const errors = [];

  for (const account of lineAccounts) {
    try {
      pushShoppingLineMessage(message, account.token, account.name);
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

  console.log("LINE送信結果:", {
    total: lineAccounts.length,
    success: results.length,
    failed: errors.length,
    details: { results, errors },
  });

  if (results.length === 0 && errors.length > 0) {
    throw new Error(`All LINE messages failed: ${JSON.stringify(errors)}`);
  }
}

/**
 * Get configured LINE accounts
 */
function getShoppingLineAccounts() {
  const props = PropertiesService.getScriptProperties();
  const accounts = [];

  const accountNames = (
    props.getProperty(SHOPPING_PROP_KEYS.accountNames) || ""
  ).split(",");

  for (let i = 1; i <= SHOPPING_CONFIG_DEFAULTS.maxLineAccounts; i++) {
    const tokenKey = `${SHOPPING_PROP_KEYS.channelAccessTokenPrefix}${i}`;
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
 * Send message to single LINE account
 */
function pushShoppingLineMessage(message, token, accountName = "Unknown") {
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

  const res = fetchShoppingWithRetry(url, options, 3);
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

function getShoppingMailBody(msg) {
  const body = msg.getPlainBody();
  if (body && body.trim()) return body;
  const html = msg.getBody() || "";
  return html
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildShoppingMailMeta(thread, msg) {
  return {
    threadId: thread.getId(),
    messageId: String(msg.getId()),
    subject: msg.getSubject() || "(no subject)",
    from: msg.getFrom() || "",
    to: msg.getTo() || "",
    cc: msg.getCc() || "",
    permalink: buildShoppingGmailPermalink(thread.getId()),
    receivedAt: msg.getDate(),
  };
}

function buildShoppingGmailPermalink(threadId) {
  return `https://mail.google.com/mail/u/0/#all/${threadId}`;
}

function getShoppingTimezone() {
  const propTz = PropertiesService.getScriptProperties().getProperty(
    SHOPPING_PROP_KEYS.timezone
  );
  return propTz || SHOPPING_CONFIG_DEFAULTS.timezone || "Asia/Tokyo";
}

function formatShoppingJst(d) {
  return Utilities.formatDate(d, getShoppingTimezone(), "yyyy/MM/dd (E) HH:mm");
}

function getOrCreateShoppingLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function fetchShoppingWithRetry(url, options, maxRetries) {
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

/* ===== Utility Functions ===== */

function setupShoppingInitial() {
  getOrCreateShoppingLabel(SHOPPING_LABELS.target);
  getOrCreateShoppingLabel(SHOPPING_LABELS.processed);
  console.log("Labels ensured.");

  const accounts = getShoppingLineAccounts();
  console.log(`Found ${accounts.length} LINE account(s):`);
  accounts.forEach((acc, idx) => {
    console.log(`  ${idx + 1}. ${acc.name}`);
  });
}

function suggestShoppingGmailFilter() {
  const targetEmail = SHOPPING_CONFIG_DEFAULTS.targetEmail;
  const filterCondition = `from:noreply@cromi.co to:${targetEmail} subject:ショッピング同行`;
  const labelName = SHOPPING_LABELS.target;

  console.log("=== Gmail Filter Setup ===");
  console.log("1. Gmail Settings -> Filters and Blocked Addresses");
  console.log("2. Create a new filter");
  console.log(`3. Search criteria: ${filterCondition}`);
  console.log(`4. Action: Apply label "${labelName}" (create if needed)`);
  console.log("5. Check 'Also apply filter to matching conversations'");
}

function suggestShoppingTriggerSetup() {
  console.log("=== Trigger Setup ===");
  console.log("1. GAS Editor -> Triggers");
  console.log("2. Add Trigger");
  console.log("3. Function: pollShoppingEmails");
  console.log("4. Event source: Time-driven");
  console.log("5. Type: Minutes timer");
  console.log("6. Interval: Every 1-2 minutes");
}

/* ===== Test Functions ===== */

function testShoppingLineNotify() {
  const accounts = getShoppingLineAccounts();

  if (accounts.length === 0) {
    console.log("No LINE accounts configured");
    return;
  }

  const testMessage = {
    type: "text",
    text: `【ショッピング同行テスト送信】\n${
      accounts.length
    }個のLINE公式アカウントへの接続テスト\n送信時刻: ${formatShoppingJst(
      new Date()
    )}`,
  };

  console.log(`Sending test message to ${accounts.length} account(s)...`);
  pushShoppingLineMessageToAll(testMessage);
  console.log("Test message sent successfully!");
}

function pingAllShoppingLineMessagingApis() {
  const accounts = getShoppingLineAccounts();

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

function showShoppingConfiguration() {
  const props = PropertiesService.getScriptProperties();
  const accounts = getShoppingLineAccounts();

  console.log("=== Current Configuration ===");
  console.log(
    `Target Email: ${
      props.getProperty(SHOPPING_PROP_KEYS.targetEmail) ||
      SHOPPING_CONFIG_DEFAULTS.targetEmail
    }`
  );
  console.log(
    `Timezone: ${
      props.getProperty(SHOPPING_PROP_KEYS.timezone) ||
      SHOPPING_CONFIG_DEFAULTS.timezone
    }`
  );
  console.log(`LINE Accounts: ${accounts.length} account(s)`);

  accounts.forEach((acc, idx) => {
    console.log(`  ${idx + 1}. ${acc.name}`);
  });

  console.log("==================");
}
