/* global Office */

const NOTIFICATION_KEY = "unsubscribe-status";
const ONE_CLICK_VALUE = "List-Unsubscribe=One-Click";

Office.onReady(() => {
  // Office.js is initialized and the function command can now be associated.
});

/**
 * Ribbon command invoked when the user clicks Unsubscribe.
 *
 * Priority:
 *   1. RFC 8058 one-click POST when List-Unsubscribe-Post is present and
 *      an HTTPS List-Unsubscribe URI is available.
 *   2. HTTPS web unsubscribe.
 *   3. HTTP web unsubscribe.
 *   4. mailto unsubscribe, composed in Outlook.
 */
async function unsubscribe(event) {
  try {
    const item = Office.context.mailbox.item;
    if (!item || typeof item.getAllInternetHeadersAsync !== "function") {
      showError("This Outlook client can't read message internet headers.");
      return;
    }

    const headers = await getAllInternetHeaders(item);
    const listUnsubscribe = getHeaderValue(headers, "List-Unsubscribe");
    const listUnsubscribePost = getHeaderValue(headers, "List-Unsubscribe-Post");

    if (!listUnsubscribe) {
      showInfo("No List-Unsubscribe header was found in this message.");
      return;
    }

    const methods = parseUnsubscribeMethods(listUnsubscribe);
    const httpsUrl = methods.find((value) => /^https:\/\//i.test(value));
    const httpUrl = methods.find((value) => /^http:\/\//i.test(value));
    const mailto = methods.find((value) => /^mailto:/i.test(value));

    // RFC 8058 one-click unsubscribe requires an HTTPS URI and the exact
    // List-Unsubscribe-Post instruction. The user's button click is the
    // explicit consent required before sending the POST.
    if (httpsUrl && isOneClickPost(listUnsubscribePost)) {
      clearStatus();
      await performOneClickUnsubscribe(httpsUrl);
      return;
    }

    if (httpsUrl) {
      clearStatus();
      await openWebUnsubscribe(httpsUrl);
      return;
    }

    if (httpUrl) {
      clearStatus();
      await openWebUnsubscribe(httpUrl);
      return;
    }

    // Only use email-based unsubscribe when there is no web alternative.
    if (mailto) {
      clearStatus();
      composeMailtoUnsubscribe(mailto);
      return;
    }

    showInfo("A List-Unsubscribe header exists, but it contains no supported unsubscribe method.");
  } catch (error) {
    console.error("Outlook Unsubscribe error:", error);
    showError("Unable to process this message's unsubscribe information.");
  } finally {
    // Outlook requires every ExecuteFunction command to signal completion.
    event.completed();
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

    showInfo("One-click unsubscribe request submitted.");
  } catch (error) {
    // Some Outlook/WebView environments or remote endpoints may still block
    // a background cross-origin POST.  A manual GET to the same URI is the
    // RFC 8058 fallback path for an ordinary unsubscribe operation.
    console.error("One-click unsubscribe POST failed; opening unsubscribe page:", error);
    showInfo("One-click unsubscribe could not be submitted automatically. Opening the unsubscribe page instead.");
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

  showInfo(
    "This Outlook client can't open this HTTP unsubscribe link automatically. Copy/open it manually."
  );
  return Promise.resolve();
}

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
          showError("Outlook couldn't open the unsubscribe page.");
        }
        resolve();
      }
    );
  });
}

/**
 * Opens a new Outlook compose form for a mailto: List-Unsubscribe method.
 * The mailto URI controls the recipient, subject and body. We intentionally
 * don't invent subject/body text when the URI doesn't provide it because
 * automated list processors may rely on the exact values supplied by the
 * sender.
 */
function composeMailtoUnsubscribe(mailtoUri) {
  const message = parseMailtoUri(mailtoUri);

  if (!message.toRecipients.length) {
    throw new Error("The mailto unsubscribe method does not contain a recipient.");
  }

  const form = {
    toRecipients: message.toRecipients,
  };

  if (message.subject !== null) {
    form.subject = message.subject;
  }

  if (message.body !== null) {
    form.htmlBody = textToSafeHtml(message.body);
  }

  Office.context.mailbox.displayNewMessageForm(form);
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
