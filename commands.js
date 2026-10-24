/* Outlook Unsubscribe - v1.0.0.13
 *
 * Behavior:
 *  - First click:
 *      RFC 8058 One-Click + HTTPS -> submit POST.
 *      Otherwise prefer List-Unsubscribe web URL, then a body link,
 *      then mailto.
 *  - If BOTH RFC 8058 One-Click and a body unsubscribe/preferences link
 *    are available, successful submission of the first One-Click request
 *    records that this message has been acted on.
 *  - Second and subsequent clicks on that same message re-scan the message
 *    and display a chooser between:
 *        1. One-Click
 *        2. Embedded Link
 *    The chooser is shown only while both methods are available.
 */

const NOTIFICATION_KEY = "unsubscribe-status";
const ONE_CLICK_VALUE = "List-Unsubscribe=One-Click";
const FIRST_ONE_CLICK_PREFIX = "outlook-unsubscribe-oneclick-used:";
const firstOneClickInMemory = new Set();

const TEXT_SIGNALS = [
  ["unsubscribe", 120],
  ["opt out", 110],
  ["opt-out", 110],
  ["email preferences", 105],
  ["email preference", 102],
  ["manage email preferences", 105],
  ["manage preferences", 95],
  ["subscription preferences", 95],
  ["manage subscription", 90],
  ["cancel subscription", 90],
  ["remove me", 85],

  // Signals inspired by Microsoft's legacy Unsubscribe add-in.
  ["désinscrire", 90],
  ["désabonner", 90],
  ["abbestellen", 90],
  ["abmelden", 85],
  ["austragen", 85],
  ["cancelar la suscripción", 90],
  ["darse de baja", 90],
  ["annullare l'iscrizione", 90],
  ["cancellarsi", 85],
  ["отписаться", 90],
  ["отменить подписку", 90],
  ["取消订阅", 90],
  ["取消訂閱", 90],
  ["退会", 85],
  ["購読中止", 85],
  ["구독을 취소", 90],
  ["구독을 중단", 90],
  ["إلغاء الاشتراك", 90],
  ["לבטל את המנוי", 90],
  ["cancelar assinatura", 90],
  ["anular a subscrição", 90],
];

const URL_SIGNALS = [
  ["unsubscribe", 100],
  ["unsub", 85],
  ["optout", 90],
  ["opt-out", 90],
  ["opt_out", 90],
  ["email-preferences", 85],
  ["email_preferences", 85],
  ["emailpreferences", 85],
  ["manage-preferences", 80],
  ["manage_preferences", 80],
  ["preferences", 55],
  ["subscription-preferences", 80],
];

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("unsubscribe", unsubscribe);
  }
});

async function unsubscribe(event) {
  let stage = "starting";

  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) {
      safeShowError("No Outlook message is currently selected.");
      return;
    }

    stage = "reading internet headers";
    const headers = await getInternetHeaders(item);

    const listUnsubscribe = getHeaderValue(headers, "List-Unsubscribe");
    const listUnsubscribePost = getHeaderValue(
      headers,
      "List-Unsubscribe-Post"
    );

    const methods = parseListUnsubscribe(listUnsubscribe);
    const httpsUrl = methods.httpsUrl;
    const httpUrl = methods.httpUrl;
    const mailtoUrl = methods.mailtoUrl;
    const hasOneClick = Boolean(
      httpsUrl && isOneClickPost(listUnsubscribePost)
    );

    // We need the body early when One-Click exists, because on subsequent
    // clicks it may become the second choice in the chooser.
    let bodyLinks = { httpsUrl: null, httpUrl: null };

    if (hasOneClick) {
      stage = "finding unsubscribe link in message body";
      bodyLinks = await findBodyUnsubscribeLinks(item);
      const embeddedUrl = bodyLinks.httpsUrl || bodyLinks.httpUrl || null;

      // Second and subsequent click:
      // Show the selection dialog ONLY if both methods are still available.
      if (embeddedUrl && hasFirstOneClickRun(item)) {
        stage = "choosing unsubscribe method";
        const choice = await showUnsubscribeChoiceDialog();

        if (choice === "one-click") {
          stage = "submitting RFC 8058 one-click POST";
          safeClearStatus();

          const submitted = await performOneClickUnsubscribe(httpsUrl);
          if (submitted) {
            safeShowInfo("One-click unsubscribe request submitted.");
          }
          return;
        }

        if (choice === "embedded") {
          stage = "opening embedded unsubscribe link";
          safeClearStatus();
          await openWebUnsubscribe(embeddedUrl);
          return;
        }

        // Dialog was closed/cancelled.
        return;
      }

      // First click, or One-Click exists without an embedded body option.
      stage = "submitting RFC 8058 one-click POST";
      safeClearStatus();

      const submitted = await performOneClickUnsubscribe(httpsUrl);

      if (submitted) {
        if (embeddedUrl) {
          markFirstOneClickRun(item);
          safeShowInfo(
            "One-click unsubscribe request submitted. Click Unsubscribe again to choose an unsubscribe method."
          );
        } else {
          safeShowInfo("One-click unsubscribe request submitted.");
        }
      }

      return;
    }

    // Header web links remain preferred over heuristic body scanning.
    if (httpsUrl) {
      stage = "opening HTTPS List-Unsubscribe link";
      safeClearStatus();
      await openWebUnsubscribe(httpsUrl);
      return;
    }

    // If we did not scan the body above, do it now.
    stage = "finding unsubscribe link in message body";
    bodyLinks = await findBodyUnsubscribeLinks(item);

    if (bodyLinks.httpsUrl) {
      stage = "opening HTTPS unsubscribe link from message body";
      safeClearStatus();
      await openWebUnsubscribe(bodyLinks.httpsUrl);
      return;
    }

    if (httpUrl) {
      stage = "opening HTTP List-Unsubscribe link";
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

    if (mailtoUrl) {
      stage = "composing unsubscribe email";
      safeClearStatus();
      composeMailtoUnsubscribe(mailtoUrl);
      return;
    }

    safeShowInfo("No unsubscribe method was found in this message.");
  } catch (error) {
    console.error("Outlook Unsubscribe failed:", stage, error);
    safeShowError(
      `Unable to process this message's unsubscribe information (${stage}).`
    );
  } finally {
    try {
      if (event && typeof event.completed === "function") {
        event.completed();
      }
    } catch (error) {
      console.warn("Unable to complete Outlook command event:", error);
    }
  }
}

function getInternetHeaders(item) {
  return new Promise((resolve, reject) => {
    if (!item || typeof item.getAllInternetHeadersAsync !== "function") {
      reject(new Error("getAllInternetHeadersAsync is not available."));
      return;
    }

    item.getAllInternetHeadersAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(result.error || new Error("Unable to read Internet headers."));
      }
    });
  });
}

function getHeaderValue(rawHeaders, headerName) {
  if (!rawHeaders || !headerName) return null;

  // RFC header unfolding: a continuation line begins with whitespace.
  const unfolded = String(rawHeaders).replace(/\r?\n[ \t]+/g, " ");
  const escaped = headerName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = unfolded.match(
    new RegExp(`^${escaped}\\s*:\\s*(.*)$`, "im")
  );

  return match ? match[1].trim() : null;
}

function parseListUnsubscribe(value) {
  const result = {
    httpsUrl: null,
    httpUrl: null,
    mailtoUrl: null,
  };

  if (!value) return result;

  const bracketed = [...String(value).matchAll(/<([^>]+)>/g)].map(
    (match) => match[1].trim()
  );

  const candidates =
    bracketed.length > 0
      ? bracketed
      : String(value)
          .split(",")
          .map((part) => part.trim().replace(/^["'<\s]+|[>"'\s]+$/g, ""))
          .filter(Boolean);

  for (const candidate of candidates) {
    if (!result.httpsUrl && /^https:\/\//i.test(candidate)) {
      result.httpsUrl = candidate;
    } else if (!result.httpUrl && /^http:\/\//i.test(candidate)) {
      result.httpUrl = candidate;
    } else if (!result.mailtoUrl && /^mailto:/i.test(candidate)) {
      result.mailtoUrl = candidate;
    }
  }

  return result;
}

function isOneClickPost(value) {
  return Boolean(
    value &&
      /(?:^|[;,]\s*)List-Unsubscribe\s*=\s*One-Click(?:\s*(?:[;,]|$))/i.test(
        String(value)
      )
  );
}

async function performOneClickUnsubscribe(url) {
  try {
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

    // no-cors produces an opaque response, so "submitted" is intentionally
    // used instead of claiming the remote server confirmed success.
    return true;
  } catch (error) {
    console.error("RFC 8058 POST failed; opening URL instead:", error);
    safeShowInfo(
      "The one-click request could not be submitted. Opening the unsubscribe page instead."
    );

    await openWebUnsubscribe(url);
    return false;
  }
}

async function findBodyUnsubscribeLinks(item) {
  try {
    const html = await getBodyHtml(item);
    if (!html) return { httpsUrl: null, httpUrl: null };

    return scoreBodyLinks(html);
  } catch (error) {
    // Body scanning is a fallback. It must never prevent header/mailto logic.
    console.warn("Unable to inspect the message body:", error);
    return { httpsUrl: null, httpUrl: null };
  }
}

function getBodyHtml(item) {
  return new Promise((resolve, reject) => {
    if (
      !item ||
      !item.body ||
      typeof item.body.getAsync !== "function"
    ) {
      reject(new Error("Message body API is not available."));
      return;
    }

    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        reject(result.error || new Error("Unable to read the message body."));
      }
    });
  });
}

function scoreBodyLinks(html) {
  const result = { httpsUrl: null, httpUrl: null };

  let doc;
  try {
    doc = new DOMParser().parseFromString(html, "text/html");
  } catch (error) {
    console.warn("Unable to parse message HTML:", error);
    return result;
  }

  const candidates = [];

  for (const anchor of doc.querySelectorAll("a[href]")) {
    const rawHref = (anchor.getAttribute("href") || "").trim();

    if (!/^https?:\/\//i.test(rawHref)) {
      continue;
    }

    const visibleText = [
      anchor.textContent || "",
      anchor.getAttribute("aria-label") || "",
      anchor.getAttribute("title") || "",
      ...[...anchor.querySelectorAll("img[alt]")].map(
        (img) => img.getAttribute("alt") || ""
      ),
    ]
      .join(" ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();

    let decodedHref = rawHref.toLowerCase();
    try {
      decodedHref = decodeURIComponent(decodedHref);
    } catch (_) {
      // Keep original href if percent-decoding fails.
    }

    let score = 0;

    for (const [term, weight] of TEXT_SIGNALS) {
      if (visibleText.includes(term)) {
        score = Math.max(score, weight);
      }
    }

    for (const [term, weight] of URL_SIGNALS) {
      if (decodedHref.includes(term)) {
        score = Math.max(score, weight);
      }
    }

    // Exact/near-exact labels are especially strong signals.
    if (/^\s*unsubscribe\s*[.!]?\s*$/i.test(visibleText)) {
      score += 40;
    } else if (/^\s*email preferences\s*[.!]?\s*$/i.test(visibleText)) {
      score += 35;
    }

    if (score > 0) {
      candidates.push({
        url: rawHref,
        protocol: /^https:\/\//i.test(rawHref) ? "https" : "http",
        score,
      });
    }
  }

  candidates.sort((a, b) => {
    // Prefer HTTPS, then the stronger semantic signal.
    if (a.protocol !== b.protocol) {
      return a.protocol === "https" ? -1 : 1;
    }
    return b.score - a.score;
  });

  for (const candidate of candidates) {
    if (candidate.protocol === "https" && !result.httpsUrl) {
      result.httpsUrl = candidate.url;
    } else if (candidate.protocol === "http" && !result.httpUrl) {
      result.httpUrl = candidate.url;
    }

    if (result.httpsUrl && result.httpUrl) break;
  }

  return result;
}

function composeMailtoUnsubscribe(mailtoUrl) {
  let parsed;

  try {
    parsed = new URL(mailtoUrl);
  } catch (error) {
    throw new Error("The mailto unsubscribe address is invalid.");
  }

  const recipients = decodeURIComponent(parsed.pathname || "")
    .split(/[;,]/)
    .map((value) => value.trim())
    .filter(Boolean);

  if (recipients.length === 0) {
    throw new Error("The mailto unsubscribe address has no recipient.");
  }

  const subject = parsed.searchParams.get("subject") || "";
  const body = parsed.searchParams.get("body") || "";

  Office.context.mailbox.displayNewMessageForm({
    toRecipients: recipients,
    subject,
    htmlBody: escapeHtml(body).replace(/\r?\n/g, "<br>"),
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function openWebUnsubscribe(url) {
  if (!/^https?:\/\//i.test(String(url || ""))) {
    throw new Error("Unsupported unsubscribe URL.");
  }

  // Classic Outlook for Windows supports openBrowserWindow.
  if (
    Office.context.ui &&
    typeof Office.context.ui.openBrowserWindow === "function"
  ) {
    try {
      Office.context.ui.openBrowserWindow(url);
      return;
    } catch (error) {
      console.warn("openBrowserWindow failed; using Office dialog:", error);
    }
  }

  await openViaRedirectDialog(url);
}

function openViaRedirectDialog(url) {
  return new Promise((resolve, reject) => {
    if (
      !Office.context.ui ||
      typeof Office.context.ui.displayDialogAsync !== "function"
    ) {
      reject(new Error("Office dialog API is unavailable."));
      return;
    }

    const redirectUrl = new URL("redirect.html", window.location.href);
    redirectUrl.searchParams.set("target", url);
    redirectUrl.searchParams.set("v", "1.0.0.13");

    Office.context.ui.displayDialogAsync(
      redirectUrl.toString(),
      {
        height: 70,
        width: 55,
        displayInIframe: false,
        promptBeforeOpen: false,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(
            result.error || new Error("Unable to open the unsubscribe page.")
          );
        }
      }
    );
  });
}

function getMessageStateKey(item) {
  if (!item) return null;
  const id = item.itemId || item.internetMessageId;
  return id ? `${FIRST_ONE_CLICK_PREFIX}${String(id)}` : null;
}

function hasFirstOneClickRun(item) {
  const key = getMessageStateKey(item);
  if (!key) return false;

  if (firstOneClickInMemory.has(key)) return true;

  try {
    return window.localStorage.getItem(key) === "1";
  } catch (error) {
    console.warn("Unable to read one-click state from localStorage:", error);
    return false;
  }
}

function markFirstOneClickRun(item) {
  const key = getMessageStateKey(item);
  if (!key) return;

  firstOneClickInMemory.add(key);

  try {
    window.localStorage.setItem(key, "1");
  } catch (error) {
    console.warn("Unable to persist one-click state:", error);
  }
}

function showUnsubscribeChoiceDialog() {
  return new Promise((resolve, reject) => {
    if (
      !Office.context.ui ||
      typeof Office.context.ui.displayDialogAsync !== "function"
    ) {
      reject(new Error("Office dialog API is unavailable."));
      return;
    }

    const dialogUrl = new URL("unsubscribe-options.html", window.location.href);
    dialogUrl.searchParams.set("v", "1.0.0.13");

    Office.context.ui.displayDialogAsync(
      dialogUrl.toString(),
      {
        height: 45,
        width: 35,
        displayInIframe: true,
        promptBeforeOpen: false,
      },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(
            result.error ||
              new Error("Unable to display the unsubscribe method chooser.")
          );
          return;
        }

        const dialog = result.value;
        let settled = false;

        const finish = (choice) => {
          if (settled) return;
          settled = true;

          try {
            dialog.close();
          } catch (_) {
            // Ignore close failures.
          }

          resolve(choice);
        };

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (args) => {
            try {
              const message = JSON.parse(args.message || "{}");

              if (
                message &&
                message.type === "unsubscribe-choice" &&
                (message.choice === "one-click" ||
                  message.choice === "embedded" ||
                  message.choice === "cancel")
              ) {
                finish(
                  message.choice === "cancel" ? null : message.choice
                );
                return;
              }
            } catch (error) {
              console.warn("Unable to parse dialog message:", error);
            }

            finish(null);
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          () => finish(null)
        );
      }
    );
  });
}

function safeShowInfo(message) {
  safeReplaceNotification(
    Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message
  );
}

function safeShowError(message) {
  safeReplaceNotification(
    Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message
  );
}

function safeReplaceNotification(type, message) {
  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;

    if (
      !item ||
      !item.notificationMessages ||
      typeof item.notificationMessages.replaceAsync !== "function"
    ) {
      console.log(message);
      return;
    }

    const details = {
      type,
      message: String(message).slice(0, 150),
    };

    if (
      type ===
      Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage
    ) {
      details.icon = "Icon.16x16";
      details.persistent = false;
    }

    item.notificationMessages.replaceAsync(
      NOTIFICATION_KEY,
      details,
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("Unable to display Outlook notification:", result.error);
        }
      }
    );
  } catch (error) {
    console.warn("Unable to display Outlook notification:", error);
  }
}

function safeClearStatus() {
  try {
    const item = Office.context.mailbox && Office.context.mailbox.item;

    if (
      item &&
      item.notificationMessages &&
      typeof item.notificationMessages.removeAsync === "function"
    ) {
      item.notificationMessages.removeAsync(NOTIFICATION_KEY);
    }
  } catch (error) {
    console.warn("Unable to clear Outlook notification:", error);
  }
}

// Keep a global symbol as a compatibility fallback for hosts that resolve the
// ExecuteFunction function name directly instead of Office.actions.associate().
window.unsubscribe = unsubscribe;
