# Outlook Unsubscribe add-in

A small Outlook web add-in that reads the current message's `List-Unsubscribe` MIME header. If the header advertises an HTTP or HTTPS URI, clicking **Unsubscribe** opens that URI in the system browser. If no web URI is available, the add-in displays an Outlook notification instead.

## Ready-to-host build

A prebuilt copy of the static web assets is included in `dist\`. You can host that folder directly over HTTPS, or regenerate it later with `npm run build`.

## Windows 11 instructions

- Development/build: see `BUILD-WINDOWS11.md`
- Installation/sideloading: see `INSTALL-WINDOWS11.md`

## What the add-in does

1. Adds an **Unsubscribe** command to received-message read surfaces.
2. Calls `getAllInternetHeadersAsync()` on the current message.
3. Unfolds and finds `List-Unsubscribe`.
4. Prefers an HTTPS URI, then HTTP.
5. In classic Outlook for Windows, opens the URI with `Office.context.ui.openBrowserWindow()`.
6. In new Outlook for Windows / Outlook on the web, falls back to a same-origin Office dialog redirect for HTTPS links.
7. If no web URI is present, displays an Outlook notification.

## What it intentionally does not do yet

- It does not perform RFC 8058 `List-Unsubscribe-Post: List-Unsubscribe=One-Click` POST requests.
- It does not automatically send mail to a `mailto:` unsubscribe address.
- It does not dynamically grey/change the Outlook ribbon button based on the current message header; Outlook does not provide the required dynamic command-state API for this scenario.

## Security behavior

The parser only accepts `https://`, `http://`, and `mailto:` methods from the header. Only HTTP(S) methods are opened in the browser. Other URI schemes are ignored.
