# Outlook Unsubscribe
This scrappy Outlook add-in allows you to quickly unsubscribe from mailing lists.

## How does it work?
Email originating via a subscription often includes a hidden **List-Unsubscribe** header that contains a link for unsubscribing from the associated subscription.
This may be in addition to a *Click _here_ to unsubscribe* which appears in the email body.

This Outlook add-in provides an **Unsubscribe** button that appears when reading email in Outlook. When clicked, this button
looks for the List-Unsubscribe header and, if present, will launch a mini-browswer to that web page. In some cases, the Unsubscribe
does not display a page but simpy just accepts the unsubscribe action.

You should only have to do this for each unique sender & subscription that you want to unsubscribe from.
IT is best to choose the most recent email from that sender to use the unsubscribe button in.

## How long does it take to be unsubscribed?
Many provides claim it can take days for an unsubscribe to take effect. And sadly, there are a few that simply ignore the requests.


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

## Installation Instructions
Below are the installation instructions.

1. **Download `manifest.xml` to your Windows 11 PC.** In GitHub, open `manifest.xml` → click **Raw/Download raw file** and save it somewhere convenient such as Downloads.

2. In your browser, open Microsoft's Outlook sideload page:
   [Open Outlook add-in sideloading](https://aka.ms/olksideload)
   Sign in with the same Microsoft account/mailbox you use in Outlook. Microsoft currently recommends this route for manual sideloading in both new and classic Outlook. ([Microsoft Learn][2])

3. In the **Add-Ins for Outlook** window, select **My add-ins**.

4. Scroll down to **Custom Addins**.

5. Click **Add a custom add-in → Add from File**. Microsoft has removed the old **Add from URL** option, so you need the local XML file. ([Microsoft Learn][2])

6. Select the downloaded:

   ```text
   manifest.xml
   ```

7. Accept the installation/security prompts.

8. Open Outlook and then open a received email message. Your **Unsubscribe** command should become available on the message's add-in/ribbon surface.

9. Test it with a newsletter that has a `List-Unsubscribe` header. Clicking **Unsubscribe** should either open its unsubscribe link or display the appropriate "no unsubscribe" message.
## I don't see the Unsubscribe
Once you installed the add-in, you should refresh your Outlook web.
The Outlook Unsubscribe only appears when you are reading an email.
If you don't see the icon, then look for an icon that looks like a stove top with 4 burners (basically a square with 4 small circles within).
Click that and you should then see Outlook Unsubscribe.

