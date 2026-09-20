/* global Office */

const NOTIFICATION_KEY = "unsubscribe-status";

Office.onReady(() => {
  // Office.js is initialized and the function command can now be associated.
});

/**
 * Ribbon command invoked when the user clicks Unsubscribe.
 * Reads the complete MIME headers from the current message and opens the
 * first HTTP(S) URI advertised by List-Unsubscribe.
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

    if (!listUnsubscribe) {
      showInfo("No List-Unsubscribe header was found in this message.");
      return;
    }

    const methods = parseUnsubscribeMethods(listUnsubscribe);
    const webUrl = methods.find((value) => /^https:\/\//i.test(value)) ||
                   methods.find((value) => /^http:\/\//i.test(value));

    if (webUrl) {
      clearStatus();
      await openWebUnsubscribe(webUrl);
      return;
    }

    const mailto = methods.find((value) => /^mailto:/i.test(value));
    if (mailto) {
      showInfo("This message offers email-based unsubscribe, but no web unsubscribe link.");
      return;
    }

    showInfo("A List-Unsubscribe header exists, but it contains no usable web link.");
  } catch (error) {
    console.error("Outlook Unsubscribe error:", error);
    showError("Unable to inspect this message for an unsubscribe link.");
  } finally {
    // Outlook requires every ExecuteFunction command to signal completion.
    event.completed();
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
