# Outlook Unsubscribe v1.0.0.9

## What changed

This version fixes email-based (`mailto:`) unsubscribe drafts in Outlook on the
web. The add-in now uses `displayNewMessageFormAsync()` when available and waits
for Outlook to report success or failure before completing the ribbon command.
Older clients retain the synchronous fallback.

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
- `README.md`
- `UPDATE.md`

Then update `manifest.xml`:

```xml
<Version>1.0.0.9</Version>
```

For each manifest URL that loads `commands.html`, use a v1.0.0.9 cache buster, for example:

```xml
https://awsles.github.io/unsubscribe/commands.html?v=1.0.0.9
```

Commit/push the changes, wait for GitHub Pages to publish them, remove the currently sideloaded add-in, and sideload the updated manifest again.
