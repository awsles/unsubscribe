# Outlook Unsubscribe v1.0.0.5

## What changed

This version adds an HTML-body fallback inspired by Microsoft's legacy Outlook Unsubscribe add-in.

Unsubscribe priority is now:

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

Then update `manifest.xml`:

```xml
<Version>1.0.0.5</Version>
```

For each manifest URL that loads `commands.html`, use a v1.0.0.5 cache buster, for example:

```xml
https://awsles.github.io/unsubscribe/commands.html?v=1.0.0.5
```

Commit/push the changes, wait for GitHub Pages to publish them, remove the currently sideloaded add-in, and sideload the updated manifest again.
