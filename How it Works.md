# Outlook Unsubscribe: How it Works
The UNSUBSCRIBE plug-in is implemented in JavaScript. When a message is opened, the UNSUBSCRIBE menu option either appears in the banner or (in the case of the Mobile Outlook version, you click the three dots to access).

## What the add-in does

When you open a message and select **Unsubscribe**, the add-in reads the message's Internet headers using the Outlook Office.js API and looks for the standard `List-Unsubscribe` and `List-Unsubscribe-Post` headers.

It chooses an unsubscribe method in this order:

1. **RFC 8058 one-click unsubscribe**  
   If the message contains both:

   ```text
   List-Unsubscribe: <https://example.com/unsubscribe/...>
   List-Unsubscribe-Post: List-Unsubscribe=One-Click
   ```

   the add-in sends the required HTTPS `POST` request with:

   ```text
   List-Unsubscribe=One-Click
   ```

   The current implementation recognizes the RFC 8058 headers and submits the request, but it does **not independently validate the message's DKIM signature**.

2. **HTTPS unsubscribe link**  
   If a normal HTTPS unsubscribe URL is available, the add-in opens it. Classic Outlook can open it in the system browser. New Outlook and Outlook on the web use an Office dialog fallback when necessary.

3. **HTTP unsubscribe link**  
   An HTTP link is used only when no HTTPS link is available and the Outlook client supports opening it.

4. **Email-based (`mailto:`) unsubscribe**  
   If there is no web-based method but the message provides a `mailto:` unsubscribe method, the add-in opens a new Outlook message addressed to the sender-provided recipient. Sender-provided `subject=` and `body=` values are preserved. When either parameter is absent, the subject defaults to `UNSUBSCRIBE` and the body defaults to `Please UNSUBSCRIBE xxx`, where `xxx` is the original recipient address. The draft is **not sent automatically**; you review it and click **Send** yourself.

5. **No supported unsubscribe method**  
   If the message does not contain a usable unsubscribe header, Outlook displays a notification explaining that no unsubscribe method was found.


### Behavior Flow

                   User clicks Unsubscribe
                           │
                           ▼
               Read Internet headers
                           │
          ┌────────────────┴────────────────┐
          │                                 │
 List-Unsubscribe-Post?                    No
          │
         Yes
          │
     HTTPS URL?
          │
         Yes ───────► RFC 8058 POST
          │
          No
          ▼
 List-Unsubscribe HTTPS?
          │
         Yes ───────► Open URL
          │
          No
          ▼
 Scan BodyAsHTML
          │
 unsubscribe HTTPS link?
          │
         Yes ───────► Open URL
          │
          No
          ▼
 List-Unsubscribe mailto?
          │
         Yes ───────► Compose Outlook email
          │
          No
          ▼
 "No unsubscribe method found"


The current priority order is:

1. List-Unsubscribe-Post + HTTPS → RFC 8058 POST
1. HTTPS List-Unsubscribe → open URL
1. HTTPS unsubscribe link found in message HTML → open URL
1. HTTP List-Unsubscribe
1. HTTP unsubscribe link found in message HTML
1. mailto: → compose unsubscribe email
1. Otherwise → “No unsubscribe method was found”

The body scanner examines anchor text, aria-label, title, image alt text, URL strings, and up to 150 characters immediately before and after each hyperlink. Adjacent text stops at neighboring hyperlinks, so wording such as “To unsubscribe, click here” identifies the nearby link without treating generic “click here” text as an unsubscribe signal. The scanner incorporates multilingual unsubscribe terminology modeled on Microsoft's legacy add-in manifest. It uses body.getAsync(...Html...), which is supported in Read mode with Mailbox 1.3 and ReadItem permissions, so the Mailbox 1.5 manifest requirement covers it.

### Example

A message might contain:

```text
List-Unsubscribe: <mailto:list@example.com?subject=unsubscribe>,
 <https://example.com/unsubscribe/12345>
```

Because a web link is available, the add-in uses the HTTPS link rather than the email method.

If the message contains only:

```text
List-Unsubscribe: <mailto:list@example.com?subject=unsubscribe&body=Sender%20text>
```

Outlook opens a new draft similar to:

```text
To: list@example.com
Subject: unsubscribe

Sender text
```

If the `mailto:` URI doesn't contain a `subject=` parameter, the subject is
`UNSUBSCRIBE`. If it doesn't contain a `body=` parameter, the body is `Please
UNSUBSCRIBE original-recipient@example.com`. Explicitly empty `subject=` and
`body=` parameters remain empty.

### Technical Overview

This is an **Outlook web add-in**, not a Windows COM/VSTO plug-in. Outlook installs the XML manifest, then loads the HTML and JavaScript runtime over HTTPS from this GitHub Pages site.

The add-in:

- Uses `getAllInternetHeadersAsync()` to read the Internet/MIME headers of the message currently being viewed.
- Uses a VersionOverrides 1.1 mobile command and a full-screen mobile task pane.
- Uses Mailbox 1.8 header APIs when the client exposes them and falls back to scanning the message body when those APIs aren't available. The manifest's 1.0 and 1.1 overrides both declare Mailbox 1.5 so the nested mobile override satisfies Microsoft's validation constraint.
- Uses only the Outlook **ReadItem** permission.
- Uses `displayNewMessageFormAsync()` where supported to create a `mailto:` unsubscribe draft.
- Uses `Office.context.ui.openBrowserWindow()` where supported, with an Office dialog fallback for clients such as new Outlook on Windows.
- Has no separate application server or database.
- Does not upload or store your message contents in a project-controlled backend.
- Contacts the sender's unsubscribe endpoint only when you explicitly click **Unsubscribe** and a web/one-click method is selected.

The hosted runtime files are published from:

```text
https://awsles.github.io/unsubscribe/
```

## Requirements

You need:

- A supported Outlook client, such as:
  - New Outlook for Windows
  - Classic Outlook for Windows
  - Outlook on the web
  - Outlook on iOS
  - Outlook on Android
- A Microsoft mailbox supported by Outlook add-ins, such as:
  - Microsoft 365 / Exchange Online work or school account
  - Outlook.com / Hotmail account
- Internet access, because Outlook web add-ins load their runtime files over HTTPS.

> **Mobile note:** Outlook mobile add-ins support Microsoft 365 and Outlook.com
> accounts, not non-Microsoft accounts such as Gmail. The mobile task pane scans
> links embedded in the message body. If a mobile client exposes the full-header
> API, it also processes `List-Unsubscribe`; otherwise header-only unsubscribe
> methods require a future Microsoft Graph-authenticated implementation.


## Source code

Project repository:
https://github.com/awsles/unsubscribe/

Hosted add-in files:
https://awsles.github.io/unsubscribe/

## Microsoft documentation

For additional details, see Microsoft's documentation:

- Sideload Outlook add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing
- Outlook add-ins overview: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/read-scenario
- Internet header APIs: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/internet-headers
- Outlook mobile add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/outlook-mobile-addins
