/* global Office */

const NOTIFICATION_KEY = "unsubscribe-status";
const ONE_CLICK_VALUE = "List-Unsubscribe=One-Click";

// Strong unsubscribe phrases used when scanning links embedded in the message
// body. This list includes the same kinds of multilingual terms used by
// Microsoft's legacy Outlook Unsubscribe add-in, plus a few common English
// variants. The scanner only considers http/https links.
const UNSUBSCRIBE_TERMS = [
  "unsubscribe",
  "opt out",
  "opt-out",
  "cancel subscription",
  "cancel my subscription",
  "manage subscription",
  "email preferences",
  "manage email preferences",
  "stop emails",
  "darse de baja",
  "cancelar la suscripción",
  "désinscrire",
  "desinscrire",
  "désabonner",
  "desabonner",
  "abbestellen",
  "abmelden",
  "austragen",
  "annullare l'iscrizione",
  "cancellarsi",
  "отписаться",
  "отменить подписку",
  "отказаться от подписки",
  "退会",
  "登録解除",
  "取消订阅",
  "终止订阅",
  "退订",
  "取消訂閱",
  "終止訂閱",
  "退訂",
  "descadastrar",
  "구독을 중단",
  "구독을 취소",
  "수신 가입을 취소",
  "가입을 취소",
  "إلغاء الاشتراك",
  "الغاء الاشتراك",
  "إلغاء الإشتراك",
  "לבטל את המנוי",
  "לבטל את הרישום",
  "לבטל את ההרשמה",
];

const UNSUBSCRIBE_URL_TERMS = [
  "unsubscribe",
  "unsub",
  "optout",
  "opt-out",
  "opt_out",
  "email-preferences",
  "email_preferences",
  "subscription-preferences",
];

Office.onReady(() => {
  // Office.js is initialized and the function command can now be associated.
});

/**
 * Ribbon command invoked when the user clicks Unsubscribe.
 *
 * Priority:
 *   1. RFC 8058 one-click POST when List-Unsubscribe-Post is present and
 *      an HTTPS List-Unsubscribe URI is available.
 *   2. HTTPS List-Unsubscribe web URL.
 *   3. HTTPS unsubscribe link embedded in the message body.
 *   4. HTTP List-Unsubscribe web URL.
 *   5. HTTP unsubscribe link embedded in the message body.
 *   6. mailto List-Unsubscribe, composed in Outlook.
 */
async function unsubscribe(event) {
  let stage = "initializing";

  try {
    stage = "getting current message";
    const item = Office.context.mailbox.item;
    if (!item || typeof item.getAllInternetHeadersAsync !== "function") {
      safeShowError("This Outlook client can't read message internet headers.");
      return;
    }

    stage = "reading internet headers";
    const headers = await getAllInternetHeaders(item);

    stage = "parsing List-Unsubscribe headers";
    const listUnsubscribe = getHeaderValue(headers, "List-Unsubscribe");
    const listUnsubscribePost = getHeaderValue(headers, "List-Unsubscribe-Post");

    let httpsUrl = null;
    let httpUrl = null;
    let mailto = null;

    if (listUnsubscribe) {
      stage = "parsing unsubscribe methods";
      const methods = parseUnsubscribeMethods(listUnsubscribe);
      httpsUrl = methods.find((value) => /^https:\/\//i.test(value)) || null;
      httpUrl = methods.find((value) => /^http:\/\//i.test(value)) || null;
      mailto = methods.find((value) => /^mailto:/i.test(value)) || null;
    }

    if (httpsUrl && isOneClickPost(listUnsubscribePost)) {
      stage = "submitting RFC 8058 one-click POST";
      safeClearStatus();
      await performOneClickUnsubscribe(httpsUrl);
      return;
    }

    if (httpsUrl) {
      stage = "opening HTTPS List-Unsubscribe page";
      safeClearStatus();
      await openWebUnsubscribe(httpsUrl);
      return;
    }

    // Before falling back to HTTP or mailto, look for a usable web unsubscribe
    // link in the message body. A visible HTTPS link is preferable to either.
    stage = "scanning message body for unsubscribe links";
    const bodyLinks = await findBodyUnsubscribeLinks(item);

    if (bodyLinks.httpsUrl) {
      stage = "opening HTTPS unsubscribe link from message body";
      safeClearStatus();
      await openWebUnsubscribe(bodyLinks.httpsUrl);
      return;
    }

    if (httpUrl) {
      stage = "opening HTTP List-Unsubscribe page";
      safeClearStatus();
      await openWebUnsubscribe(httpUrl);
      return;
    }

    if (bodyLinks.httpUrl) {
      stage = "opening HTTP unsubscribe link from message body";
      safeClearStatus();
      await openWebUnsubscribe(bodyLinks.httpUrl);
      return;
    }

    if (mailto) {
      stage = "composing mailto unsubscribe message";
      safeClearStatus();
      composeMailtoUnsubscribe(mailto, headers, item);
      return;
    }

    if (listUnsubscribe) {
      safeShowInfo("A List-Unsubscribe header exists, but no usable unsubscribe method was found.");
    } else {
      safeShowInfo("No unsubscribe method was found in this message.");
    }
  } catch (error) {
    console.error("Outlook Unsubscribe error at stage:", stage, error);
    const detail = getErrorText(error);
    const message = (`Unsubscribe failed while ${stage}: ${detail}`).slice(0, 145);
    safeShowError(message);
  } finally {
    try {
      if (event && typeof event.completed === "function") {
        event.completed();
      }
    } catch (completionError) {
      console.error("Unable to complete Outlook command event:", completionError);
    }
  }
}

function isOneClickPost(value) {
  return typeof value === "string" &&
    value.trim().toLowerCase() === ONE_CLICK_VALUE.toLowerCase();
}

/**
 * Sends the RFC 8058 one-click POST.
 *
 * no-cors is intentional: unsubscribe endpoints generally don't expose CORS
 * response headers. The request can still be sent, but the response is opaque,
 * so the add-in cannot truthfully claim that the remote server confirmed it.
 * credentials:"omit" ensures cookies and HTTP credentials are not sent.
 */
async function performOneClickUnsubscribe(url) {
  if (!/^https:\/\//i.test(url)) {
    throw new Error("RFC 8058 one-click unsubscribe requires HTTPS.");
  }

  if (typeof fetch !== "function") {
    console.warn("fetch() is unavailable; falling back to the unsubscribe page.");
    await openWebUnsubscribe(url);
    return;
  }

  try {
    // RFC 8058 permits application/x-www-form-urlencoded and requires that
    // cookies / HTTP credentials not be sent.  Cross-origin unsubscribe
    // endpoints generally don't expose CORS headers, so no-cors is used.
    // IMPORTANT: Fetch requires redirect mode "follow" for no-cors requests;
    // therefore we intentionally leave redirect at its default value.
    await fetch(url, {
      method: "POST",
      mode: "no-cors",
      credentials: "omit",
      referrerPolicy: "no-referrer",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: ONE_CLICK_VALUE,
    });

    safeShowInfo("One-click unsubscribe request submitted.");
  } catch (error) {
    // Some Outlook/WebView environments or remote endpoints may still block
    // a background cross-origin POST.  A manual GET to the same URI is the
    // RFC 8058 fallback path for an ordinary unsubscribe operation.
    console.error("One-click unsubscribe POST failed; opening unsubscribe page:", error);
    safeShowInfo("One-click unsubscribe could not be submitted automatically. Opening the unsubscribe page instead.");
    await openWebUnsubscribe(url);
  }
}

function openWebUnsubscribe(url) {
  const canOpenExternalBrowser =
    Office.context.requirements &&
    Office.context.requirements.isSetSupported(
      "OpenBrowserWindowApi",
      "1.1"
    );

  if (canOpenExternalBrowser && Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(url);
    return Promise.resolve();
  }

  // New Outlook on Windows and Outlook on the web currently don't support
  // OpenBrowserWindowApi 1.1. For HTTPS links, use an Office dialog whose
  // first page is same-origin with the add-in, then immediately redirect.
  if (/^https:\/\//i.test(url)) {
    return openInOfficeDialog(url);
  }

  safeShowInfo(
    "This Outlook client can't open this HTTP unsubscribe link automatically. Copy/open it manually."
  );
  return Promise.resolve();
}

// Do not add promptbeforeOpen:false as it will fail
function openInOfficeDialog(targetUrl) {
  return new Promise((resolve) => {
    const redirectUrl = new URL("redirect.html", window.location.href);
    redirectUrl.searchParams.set("target", targetUrl);

    Office.context.ui.displayDialogAsync(
      redirectUrl.toString(),
      { height: 70, width: 55, displayInIframe: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Unable to open unsubscribe dialog:", result.error);
          safeShowError("Outlook couldn't open the unsubscribe page.");
        }
        resolve();
      }
    );
  });
}

/**
 * Opens a new Outlook compose form for a mailto: List-Unsubscribe method.
 * The mailto URI controls the recipient and, when present, the subject.
 * If subject= is absent, "UNSUBSCRIBE" is used. A sender-provided body is
 * preserved; otherwise the body identifies the original recipient address.
 *
 * Outlook creates the draft in the mailbox context of the message being read.
 * Office.js doesn't expose Outlook's COM SentOnBehalfOfName property or any
 * other writable From field. If Outlook doesn't automatically select an alias,
 * the user must select it in the draft before sending.
 */
function composeMailtoUnsubscribe(mailtoUri, rawHeaders, item) {
  const message = parseMailtoUri(mailtoUri);

  if (!message.toRecipients.length) {
    throw new Error("The mailto unsubscribe method does not contain a recipient.");
  }

  const originalRecipient = getOriginalRecipientAddress(rawHeaders, item);
  if (!originalRecipient) {
    throw new Error("The address that received the original message could not be determined.");
  }

  const form = {
    toRecipients: message.toRecipients,
    subject: message.subject !== null ? message.subject : "UNSUBSCRIBE",
    htmlBody: textToSafeHtml(
      message.body !== null
        ? message.body
        : `Please UNSUBSCRIBE ${originalRecipient}`
    ),
  };

  Office.context.mailbox.displayNewMessageForm(form);
}

/**
 * Returns the original recipient address. The Internet To header is preferred
 * because it preserves an alias that Outlook may resolve to a mailbox's primary
 * address. Outlook's resolved To collection and user profile are fallbacks.
 */
function getOriginalRecipientAddress(rawHeaders, item) {
  const headerAddress = extractFirstEmailAddress(getHeaderValue(rawHeaders, "To"));
  if (headerAddress) return headerAddress;

  if (item && Array.isArray(item.to)) {
    for (const recipient of item.to) {
      const address = recipient && (recipient.emailAddress || recipient.address);
      if (address) return String(address).trim();
    }
  }

  const profile = Office.context.mailbox.userProfile;
  return profile && profile.emailAddress
    ? String(profile.emailAddress).trim()
    : null;
}

function extractFirstEmailAddress(value) {
  if (!value) return null;

  const match = String(value).match(
    /(?:<\s*)?([A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,})(?:\s*>)?/i
  );
  return match ? match[1] : null;
}

/**
 * Parses the RFC 6068 pieces we need from a mailto URI.
 * Supports recipients in both the URI path and a ?to= field, plus subject
 * and body. Other headers are intentionally ignored.
 */
function parseMailtoUri(uri) {
  if (!/^mailto:/i.test(uri)) {
    throw new Error("Not a mailto URI.");
  }

  const raw = uri.slice(uri.indexOf(":") + 1);
  const questionMark = raw.indexOf("?");
  const rawTo = questionMark >= 0 ? raw.slice(0, questionMark) : raw;
  const rawQuery = questionMark >= 0 ? raw.slice(questionMark + 1) : "";

  const recipients = decodeRecipientList(rawTo);
  let subject = null;
  let body = null;

  if (rawQuery) {
    for (const field of rawQuery.split("&")) {
      if (!field) continue;

      const equals = field.indexOf("=");
      const rawName = equals >= 0 ? field.slice(0, equals) : field;
      const rawValue = equals >= 0 ? field.slice(equals + 1) : "";
      const name = safeDecodeURIComponent(rawName).toLowerCase();
      const value = safeDecodeURIComponent(rawValue);

      if (name === "to") {
        recipients.push(...decodeRecipientList(rawValue));
      } else if (name === "subject" && subject === null) {
        subject = value;
      } else if (name === "body" && body === null) {
        body = value;
      }
    }
  }

  return {
    toRecipients: dedupeStrings(recipients),
    subject,
    body,
  };
}

function decodeRecipientList(rawValue) {
  if (!rawValue) return [];

  return rawValue
    .split(",")
    .map((value) => safeDecodeURIComponent(value.trim()))
    .filter(Boolean);
}

function safeDecodeURIComponent(value) {
  try {
    // decodeURIComponent correctly leaves literal '+' characters unchanged,
    // which is important for mailto addresses such as user+tag@example.com.
    return decodeURIComponent(value);
  } catch (error) {
    console.warn("Unable to decode mailto component:", value, error);
    return value;
  }
}

function dedupeStrings(values) {
  const seen = new Set();
  const output = [];

  for (const value of values) {
    const key = value.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      output.push(value);
    }
  }

  return output;
}

function textToSafeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/\r\n|\r|\n/g, "<br>");
}


/**
 * Looks for unsubscribe links embedded in the current message body.
 *
 * Failure to read/parse the body is deliberately non-fatal. Header-based
 * mailto unsubscribe can still be used if body inspection isn't available.
 */
async function findBodyUnsubscribeLinks(item) {
  try {
    if (!item.body || typeof item.body.getAsync !== "function") {
      return { httpsUrl: null, httpUrl: null };
    }

    const html = await getMessageBodyHtml(item);
    return scoreBodyUnsubscribeLinks(html);
  } catch (error) {
    console.warn("Unable to scan message body for unsubscribe links:", error);
    return { httpsUrl: null, httpUrl: null };
  }
}

function getMessageBodyHtml(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(result.error || new Error("body.getAsync failed"));
      }
    });
  });
}

/**
 * Parse all http/https anchors in the message and score their likelihood of
 * being an unsubscribe control. The link text, accessibility labels, title,
 * descendant image alt text, and URL itself are all considered.
 */
function scoreBodyUnsubscribeLinks(html) {
  if (!html || typeof DOMParser === "undefined") {
    return { httpsUrl: null, httpUrl: null };
  }

  const document = new DOMParser().parseFromString(String(html), "text/html");
  const anchors = Array.from(document.querySelectorAll("a[href]"));
  const candidates = [];

  for (const anchor of anchors) {
    const rawHref = (anchor.getAttribute("href") || "").trim();
    if (!/^https?:\/\//i.test(rawHref)) {
      continue;
    }

    const signalText = collectAnchorSignalText(anchor);
    const score = scoreUnsubscribeCandidate(signalText, rawHref);
    if (score <= 0) {
      continue;
    }

    candidates.push({
      url: rawHref,
      score,
      isHttps: /^https:\/\//i.test(rawHref),
    });
  }

  // Highest confidence first. For equal scores, prefer HTTPS.
  candidates.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (a.isHttps !== b.isHttps) return a.isHttps ? -1 : 1;
    return 0;
  });

  const bestHttps = candidates.find((candidate) => candidate.isHttps);
  const bestHttp = candidates.find((candidate) => !candidate.isHttps);

  return {
    httpsUrl: bestHttps ? bestHttps.url : null,
    httpUrl: bestHttp ? bestHttp.url : null,
  };
}

function collectAnchorSignalText(anchor) {
  const parts = [
    anchor.textContent || "",
    anchor.getAttribute("aria-label") || "",
    anchor.getAttribute("title") || "",
  ];

  for (const image of Array.from(anchor.querySelectorAll("img[alt]"))) {
    parts.push(image.getAttribute("alt") || "");
  }

  return normalizeSearchText(parts.join(" "));
}

function scoreUnsubscribeCandidate(signalText, href) {
  const text = normalizeSearchText(signalText);
  const urlText = normalizeSearchText(safeDecodeURIComponent(String(href)));
  let score = 0;

  for (const term of UNSUBSCRIBE_TERMS) {
    const normalizedTerm = normalizeSearchText(term);
    if (!normalizedTerm) continue;

    if (text === normalizedTerm) {
      score = Math.max(score, 120);
    } else if (text.includes(normalizedTerm)) {
      score = Math.max(score, 100);
    }
  }

  for (const term of UNSUBSCRIBE_URL_TERMS) {
    if (urlText.includes(normalizeSearchText(term))) {
      score = Math.max(score, 60);
    }
  }

  // Favor HTTPS when everything else is equal without allowing protocol alone
  // to turn an unrelated link into a candidate.
  if (score > 0 && /^https:\/\//i.test(href)) {
    score += 5;
  }

  return score;
}

function normalizeSearchText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
}

function getAllInternetHeaders(item) {
  return new Promise((resolve, reject) => {
    item.getAllInternetHeadersAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(result.error || new Error("getAllInternetHeadersAsync failed"));
      }
    });
  });
}

/**
 * Returns an unfolded RFC-style header value. Continuation lines that begin
 * with SP/HTAB are joined to the previous line.
 */
function getHeaderValue(rawHeaders, headerName) {
  const unfolded = String(rawHeaders).replace(/\r?\n[ \t]+/g, " ");
  const escaped = headerName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = unfolded.match(new RegExp(`^${escaped}:\\s*(.+)$`, "im"));
  return match ? match[1].trim() : null;
}

/**
 * List-Unsubscribe normally contains comma-separated URI references inside
 * angle brackets, for example:
 *   <mailto:list@example.com>, <https://example.com/unsubscribe/123>
 *
 * This parser follows that form but also tolerates a bare HTTP(S)/mailto URI.
 */
function parseUnsubscribeMethods(headerValue) {
  const values = [];
  const bracketed = /<([^>]+)>/g;
  let match;

  while ((match = bracketed.exec(headerValue)) !== null) {
    const value = match[1].trim();
    if (isAllowedMethod(value)) {
      values.push(value);
    }
  }

  if (values.length === 0) {
    for (const part of headerValue.split(",")) {
      const value = part.trim().replace(/^<|>$/g, "");
      if (isAllowedMethod(value)) {
        values.push(value);
      }
    }
  }

  return values;
}

function isAllowedMethod(value) {
  return /^(https?:\/\/|mailto:)/i.test(value);
}

function getErrorText(error) {
  if (!error) return "unknown error";
  if (typeof error === "string") return error;
  if (error.message) return String(error.message);
  if (error.name) return String(error.name);
  try {
    return JSON.stringify(error);
  } catch (_) {
    return "unknown error";
  }
}

function safeShowInfo(message) {
  try {
    showInfo(message);
  } catch (error) {
    console.error("Unable to display informational notification:", error);
  }
}

function safeShowError(message) {
  try {
    showError(message);
  } catch (error) {
    console.error("Unable to display error notification:", error);
  }
}

function safeClearStatus() {
  try {
    clearStatus();
  } catch (error) {
    console.error("Unable to clear prior notification:", error);
  }
}

function showInfo(message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    NOTIFICATION_KEY,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message,
      icon: "Icon.16x16",
      persistent: false,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Unable to display notification:", result.error);
      }
    }
  );
}

function showError(message) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    NOTIFICATION_KEY,
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message,
    },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Unable to display error notification:", result.error);
      }
    }
  );
}

function clearStatus() {
  Office.context.mailbox.item.notificationMessages.removeAsync(
    NOTIFICATION_KEY,
    () => {}
  );
}

Office.actions.associate("unsubscribe", unsubscribe);
