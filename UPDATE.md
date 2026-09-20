# Outlook Unsubscribe update 1.0.0.2

Replace the repository-root `commands.js` and `commands.html` with the files in this package.

Then update the existing `manifest.xml` (do not replace it with an older copy):

1. Change `<Version>1.0.0.1</Version>` to `<Version>1.0.0.2</Version>`.
2. Recommended cache-buster: change the `Commands.Url` value to:
   `https://awsles.github.io/unsubscribe/commands.html?v=1.0.0.2`
3. Commit/push the changes and wait for GitHub Pages to publish them.
4. Remove the sideloaded 1.0.0.1 add-in and sideload the updated manifest.xml.

Behavior priority:
1. `List-Unsubscribe-Post: List-Unsubscribe=One-Click` + HTTPS URI -> HTTPS POST.
2. HTTPS List-Unsubscribe URI -> open web unsubscribe.
3. HTTP List-Unsubscribe URI -> open where supported.
4. mailto List-Unsubscribe URI -> open a pre-addressed Outlook compose draft.

The mailto draft preserves sender-provided recipient, subject, and body. It does not invent subject/body text when omitted.

Note: RFC 8058 also specifies DKIM validation requirements. This client-only update implements the one-click POST mechanics but does not independently cryptographically validate DKIM signatures.
