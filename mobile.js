/* global Office */

const ONE_CLICK_VALUE = "List-Unsubscribe=One-Click";

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
  document.getElementById("closeAction").addEventListener("click", closePane);
  initializeMobile().catch((error) => {
    console.error("Mobile unsubscribe initialization failed:", error);
    setStatus(`Unable to inspect this message: ${getErrorText(error)}`);
  });
});

async function initializeMobile() {
  const item = Office.context.mailbox.item;
  if (!item) {
    throw new Error("No message is currently selected.");
  }

  const headers = await tryGetInternetHeaders(item);
  const originalRecipient = getOriginalRecipientAddress(headers, item);
  const headerChoice = chooseHeaderMethod(headers, originalRecipient);

  if (headerChoice) {
    renderChoice(headerChoice);
    return;
  }

  const html = await getMessageBodyHtml(item);
  const bodyChoice = chooseBodyMethod(html, originalRecipient);

  if (bodyChoice) {
    renderChoice(bodyChoice);
    return;
  }

  setStatus(
    headers
      ? "No supported unsubscribe method was found in this message."
      : "No unsubscribe link was found in the message body. Outlook mobile doesn't expose the full message headers used by the desktop add-in."
  );
}

function renderChoice(choice) {
  const link = document.getElementById("linkAction");
  const postButton = document.getElementById("postAction");

  link.hidden = true;
  postButton.hidden = true;

  if (choice.type === "one-click") {
    setStatus("This message supports one-click unsubscribe.");
    postButton.hidden = false;
    postButton.onclick = async () => {
      postButton.disabled = true;
      setStatus("Submitting the unsubscribe request…");

      try {
        await performOneClickUnsubscribe(choice.url);
        setStatus("The unsubscribe request was submitted.");
      } catch (error) {
        console.error("One-click unsubscribe failed:", error);
        setStatus("Automatic unsubscribe failed. Open the unsubscribe page instead.");
        postButton.hidden = true;
        renderLink(choice.url, "Open unsubscribe page");
      } finally {
        postButton.disabled = false;
      }
    };
    return;
  }

  if (choice.type === "mailto") {
    setStatus("This sender uses an email-based unsubscribe method.");
    renderLink(choice.url, "Create unsubscribe email");
    return;
  }

  setStatus("An unsubscribe page was found.");
  renderLink(choice.url, "Open unsubscribe page");
}

function renderLink(url, label) {
  const link = document.getElementById("linkAction");
  link.href = url;
  link.textContent = label;
  link.target = /^mailto:/i.test(url) ? "_self" : "_blank";
  link.rel = /^mailto:/i.test(url) ? "" : "noopener noreferrer";
  link.hidden = false;
}

function chooseHeaderMethod(headers, originalRecipient) {
  if (!headers) return null;

  const listUnsubscribe = getHeaderValue(headers, "List-Unsubscribe");
  if (!listUnsubscribe) return null;

  const methods = parseUnsubscribeMethods(listUnsubscribe);
  const httpsUrl = methods.find((value) => /^https:\/\//i.test(value));
  const httpUrl = methods.find((value) => /^http:\/\//i.test(value));
  const mailto = methods.find((value) => /^mailto:/i.test(value));
  const oneClick = getHeaderValue(headers, "List-Unsubscribe-Post");

  if (httpsUrl && String(oneClick || "").trim().toLowerCase() === ONE_CLICK_VALUE.toLowerCase()) {
    return { type: "one-click", url: httpsUrl };
  }
  if (httpsUrl) return { type: "web", url: httpsUrl };
  if (httpUrl) return { type: "web", url: httpUrl };
  if (mailto) return { type: "mailto", url: prepareMailtoUri(mailto, originalRecipient) };
  return null;
}

function chooseBodyMethod(html, originalRecipient) {
  if (!html || typeof DOMParser === "undefined") return null;

  const parsed = new DOMParser().parseFromString(String(html), "text/html");
  const candidates = [];

  for (const anchor of Array.from(parsed.querySelectorAll("a[href]"))) {
    const href = (anchor.getAttribute("href") || "").trim();
    if (!/^(https?:\/\/|mailto:)/i.test(href)) continue;

    const signalText = collectAnchorSignalText(anchor);
    const score = scoreUnsubscribeCandidate(signalText, href);
    if (score <= 0) continue;

    candidates.push({ href, score, rank: protocolRank(href) });
  }

  candidates.sort((a, b) => b.score - a.score || b.rank - a.rank);
  if (!candidates.length) return null;

  const best = candidates[0].href;
  return /^mailto:/i.test(best)
    ? { type: "mailto", url: prepareMailtoUri(best, originalRecipient) }
    : { type: "web", url: best };
}

function protocolRank(url) {
  if (/^https:/i.test(url)) return 3;
  if (/^mailto:/i.test(url)) return 2;
  return 1;
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
    if (text === normalizedTerm) score = Math.max(score, 120);
    else if (text.includes(normalizedTerm)) score = Math.max(score, 100);
  }

  for (const term of UNSUBSCRIBE_URL_TERMS) {
    if (urlText.includes(normalizeSearchText(term))) {
      score = Math.max(score, 60);
    }
  }

  if (score > 0 && /^https:/i.test(href)) score += 5;
  return score;
}

function normalizeSearchText(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFKC")
    .replace(/\s+/g, " ")
    .trim();
}

function prepareMailtoUri(uri, originalRecipient) {
  const questionMark = uri.indexOf("?");
  const base = questionMark >= 0 ? uri.slice(0, questionMark) : uri;
  const rawQuery = questionMark >= 0 ? uri.slice(questionMark + 1) : "";
  const fields = rawQuery ? rawQuery.split("&").filter(Boolean) : [];
  let hasSubject = false;
  let hasBody = false;

  for (const field of fields) {
    const equals = field.indexOf("=");
    const rawName = equals >= 0 ? field.slice(0, equals) : field;
    const name = safeDecodeURIComponent(rawName).toLowerCase();
    if (name === "subject") hasSubject = true;
    if (name === "body") hasBody = true;
  }

  if (!hasSubject) fields.push(`subject=${encodeURIComponent("UNSUBSCRIBE")}`);
  if (!hasBody) {
    const body = originalRecipient
      ? `Please UNSUBSCRIBE ${originalRecipient}`
      : "Please UNSUBSCRIBE me from this list.";
    fields.push(`body=${encodeURIComponent(body)}`);
  }

  return fields.length ? `${base}?${fields.join("&")}` : base;
}

function getOriginalRecipientAddress(headers, item) {
  const headerAddress = extractFirstEmailAddress(getHeaderValue(headers, "To"));
  if (headerAddress) return headerAddress;

  if (item && Array.isArray(item.to)) {
    for (const recipient of item.to) {
      const address = recipient && (recipient.emailAddress || recipient.address);
      if (address) return String(address).trim();
    }
  }

  const profile = Office.context.mailbox.userProfile;
  return profile && profile.emailAddress ? String(profile.emailAddress).trim() : null;
}

function extractFirstEmailAddress(value) {
  if (!value) return null;
  const match = String(value).match(
    /(?:<\s*)?([A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,})(?:\s*>)?/i
  );
  return match ? match[1] : null;
}

function getMessageBodyHtml(item) {
  return new Promise((resolve, reject) => {
    if (!item.body || typeof item.body.getAsync !== "function") {
      reject(new Error("This Outlook client can't read the message body."));
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

function tryGetInternetHeaders(item) {
  if (typeof item.getAllInternetHeadersAsync !== "function") {
    return Promise.resolve("");
  }

  return new Promise((resolve) => {
    item.getAllInternetHeadersAsync((result) => {
      resolve(
        result.status === Office.AsyncResultStatus.Succeeded
          ? result.value || ""
          : ""
      );
    });
  });
}

function getHeaderValue(rawHeaders, headerName) {
  if (!rawHeaders) return null;
  const unfolded = String(rawHeaders).replace(/\r?\n[ \t]+/g, " ");
  const escaped = headerName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = unfolded.match(new RegExp(`^${escaped}:\\s*(.+)$`, "im"));
  return match ? match[1].trim() : null;
}

function parseUnsubscribeMethods(headerValue) {
  const values = [];
  const bracketed = /<([^>]+)>/g;
  let match;

  while ((match = bracketed.exec(headerValue)) !== null) {
    const value = match[1].trim();
    if (/^(https?:\/\/|mailto:)/i.test(value)) values.push(value);
  }

  if (!values.length) {
    for (const part of headerValue.split(",")) {
      const value = part.trim().replace(/^<|>$/g, "");
      if (/^(https?:\/\/|mailto:)/i.test(value)) values.push(value);
    }
  }

  return values;
}

function safeDecodeURIComponent(value) {
  try {
    return decodeURIComponent(value);
  } catch (_) {
    return value;
  }
}

function performOneClickUnsubscribe(url) {
  if (!/^https:\/\//i.test(url)) {
    return Promise.reject(new Error("One-click unsubscribe requires HTTPS."));
  }

  return fetch(url, {
    method: "POST",
    mode: "no-cors",
    credentials: "omit",
    referrerPolicy: "no-referrer",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: ONE_CLICK_VALUE,
  });
}

function setStatus(message) {
  document.getElementById("status").textContent = message;
}

function closePane() {
  if (Office.context.ui && typeof Office.context.ui.closeContainer === "function") {
    Office.context.ui.closeContainer();
    return;
  }
  window.close();
}

function getErrorText(error) {
  if (!error) return "unknown error";
  if (typeof error === "string") return error;
  return error.message || error.name || "unknown error";
}
