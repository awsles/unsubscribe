# Outlook Unsubscribe v1.0.0.11

## What changed

This version improves message-body unsubscribe detection on desktop, web, and mobile:

- Examines up to 150 characters immediately before and after each hyperlink.
- Associates adjacent unsubscribe wording only with the nearest hyperlink by
  stopping at neighboring links.
- Detects patterns such as `To unsubscribe, click here` and `Click here to
  unsubscribe` without using generic `click here` text as a signal by itself.
- Gives direct unsubscribe anchor text precedence over adjacent-text matches.

The Outlook mobile support added in v1.0.0.10 is retained:

- Sets both the VersionOverrides 1.0 and nested VersionOverrides 1.1 minimum
  Mailbox requirement to 1.5, as required by Microsoft's manifest validator.
- Adds a VersionOverrides 1.1 `<MobileFormFactor>` with a
  `MobileMessageReadCommandSurface` button.
- Adds `mobile.html` and `mobile.js`, which provide a full-screen mobile task
  pane and user-initiated unsubscribe actions.
- Scans message-body links using the same multilingual unsubscribe terms as the
  desktop command, including HTTPS, HTTP, and `mailto:` links.
- Preserves sender-provided `mailto:` subject/body values and supplies the same
  defaults used by the desktop add-in when they are absent.
- Tries the Internet-header API when a client exposes it, but continues with
  body scanning when it isn't available.
- Keeps the existing desktop command and Outlook-on-the-web compose fix.

The v1.0.0.8 email draft behavior is retained:

- The recipient continues to come from the `mailto:` URI.
- A sender-provided `subject=` value is preserved, including an explicitly
  empty value.
- If the URI doesn't contain `subject=`, the subject is `UNSUBSCRIBE`.
- A sender-provided `body=` value is preserved, including an explicitly empty
  value.
- If the URI doesn't contain `body=`, the body is `Please UNSUBSCRIBE xxx`,
  where `xxx` is the first email address in the original message's Internet
  `To:` header. Outlook's resolved recipient and mailbox profile addresses are
  used as fallbacks.

Outlook creates the draft in the mailbox context of the message being read.
This Office.js web add-in can't use Outlook COM's `SentOnBehalfOfName` property
and Office.js doesn't expose another writable From field. If Outlook doesn't
automatically choose an original alias, select it in the draft before sending.

Unsubscribe priority remains:

1. `List-Unsubscribe-Post: List-Unsubscribe=One-Click` + HTTPS URL -> RFC 8058 POST.
2. HTTPS URL in `List-Unsubscribe` -> open the unsubscribe page.
3. HTTPS unsubscribe/opt-out link found in the message HTML -> open the link.
4. HTTP URL in `List-Unsubscribe` -> open where the Outlook client permits it.
5. HTTP unsubscribe/opt-out link found in the message HTML -> open where permitted.
6. `mailto:` in `List-Unsubscribe` -> open a pre-addressed Outlook draft.
7. Otherwise -> display `No unsubscribe method was found in this message.`

The HTML scanner examines anchor text, `aria-label`, `title`, descendant image `alt` text, and URL tokens. It only considers `http://` and `https://` links and prefers HTTPS. It includes multilingual unsubscribe terms patterned after Microsoft's legacy Outlook Unsubscribe add-in.

Failure to retrieve or parse the message body is non-fatal. The add-in continues to any available `mailto:` fallback.

## Deploy

Replace the repository-root files:

- `commands.js`
- `commands.html`
- `mobile.js`
- `mobile.html`
- `README.md`
- `UPDATE.md`

Then update `manifest.xml`:

```xml
<Version>1.0.0.11</Version>
```

For each manifest URL that loads JavaScript, use a v1.0.0.11 cache buster, for example:

```xml
https://awsles.github.io/unsubscribe/commands.html?v=1.0.0.11
https://awsles.github.io/unsubscribe/mobile.html?v=1.0.0.11
```

Commit/push the changes, wait for GitHub Pages to publish them, remove the currently sideloaded add-in, and sideload the updated manifest again.
