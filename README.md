# Outlook Unsubscribe

**Outlook Unsubscribe** is a lightweight Outlook web add-in that adds an **Unsubscribe** command to received email messages. It reads the standard mailing-list unsubscribe headers in the current message and uses the best unsubscribe method advertised by the sender.

The add-in is designed for Microsoft Outlook on Windows, including both **new Outlook** and **classic Outlook**, and can also run in Outlook on the web where the required Office.js APIs are supported.

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


# Installing the add-in
The add-in is currently installed by **sideloading** its `manifest.xml` file into Outlook.
This is done by installing it using the WEB version of Outlook (often called OWA for Outlook Web Access). IF you install it here,
it will eventually propagate to your tabel and desktop versions of Outlook.

## 1. Download `manifest.xml`
Download the current manifest from the project repository in GitHub:

https://github.com/awsles/unsubscribe/blob/main/manifest.xml

CLick the link to open the file and choose **Download raw file**, then save it somewhere convenient on your PC, such as your **Downloads** folder.

You can also view the deployed copy here:

https://awsles.github.io/unsubscribe/manifest.xml

Outlook requires a local XML file for manual sideloading; Microsoft's current sideloading interface no longer provides an **Add from URL** option.

## 2. Open the Outlook add-in sideloading page
In a web browser, open:

https://aka.ms/olksideload

Sign in with the **same Microsoft account/mailbox that you use in Outlook**.
Give the page a moment to fully load the "Add-Ins for Outlook" popup window.

The **Add-Ins for Outlook** dialog should open.

## 3. Open My add-ins
In the Add-Ins for Outlook dialog:

1. Select **My add-ins**.
2. Scroll to the **Custom Addins** section near the bottom.
3. Select **Add a custom add-in**.
4. Select **Add from File**.

## 4. Select the manifest
Choose the `manifest.xml` file you downloaded in Step 1.

Accept the installation/security prompts.

The add-in is now associated with your Outlook mailbox. Microsoft states that an add-in sideloaded through this dialog should also become available in supported Outlook desktop clients that use the same mailbox.

> **Classic Outlook note:** A manually sideloaded add-in can occasionally take up to 24 hours to appear in classic Outlook because of client caching. New Outlook and Outlook on the web normally reflect the installation much sooner.

## 5. Find the Unsubscribe command
Open a **received email message** in Outlook.

### New Outlook for Windows / Outlook on the web

1. Select the message.
2. In the message action bar, select **Apps**.
3. Look for **Outlook Unsubscribe** or **Unsubscribe**.

You may be able to pin the command depending on your Outlook configuration.

### Classic Outlook for Windows
Open a received message and look for the **Unsubscribe** command on the message's add-in/ribbon surface. If it does not appear immediately after installation, restart Outlook and allow time for its add-in cache to refresh.

## 6. Test it
For the best test, open a legitimate newsletter or mailing-list message that provides a `List-Unsubscribe` header (i.e., look for some junk mail).

Select **Unsubscribe**.

Depending on the message, the add-in will:

- submit an RFC 8058 one-click unsubscribe request,
- open the sender's unsubscribe webpage,
- open a pre-addressed unsubscribe email draft, or
- display a message saying that no supported unsubscribe method was found.

## I don't see the Unsubscribe
Once you installed the add-in, you should refresh your Outlook web.
The Outlook Unsubscribe only appears when you are reading an email.
If you don't see the icon, then look for an icon that looks like a stove top with 4 burners (basically a square with 4 small circles within).
Click that and you should then see Outlook Unsubscribe.

## Updating the add-in

Most JavaScript/HTML changes are loaded from the hosted GitHub Pages site automatically. However, if `manifest.xml` itself changes, Outlook may continue using the copy that was previously sideloaded.

When the manifest version changes:

1. Download the new `manifest.xml`.
2. Remove the existing custom add-in from **My add-ins** if necessary.
3. Sideload the new manifest using **Add from File** again.

The manifest's `<Version>` value can be used to verify that you have the expected release.

## Removing the add-in

To uninstall the sideloaded add-in:

1. Open https://aka.ms/olksideload 
2. Select **My add-ins**.
3. Locate **Outlook Unsubscribe** under your custom add-ins.
4. Remove the add-in.


---
## TECHNICAL INFO

### What the add-in does

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
   If there is no web-based method but the message provides a `mailto:` unsubscribe method, the add-in opens a new Outlook message populated with the recipient, subject, and body specified by the sender. The draft is **not sent automatically**; you review it and click **Send** yourself.

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


The v1.0.0.5 priority order is:

1. List-Unsubscribe-Post + HTTPS → RFC 8058 POST
1. HTTPS List-Unsubscribe → open URL
1. HTTPS unsubscribe link found in message HTML → open URL
1. HTTP List-Unsubscribe
1. HTTP unsubscribe link found in message HTML
1. mailto: → compose unsubscribe email
1. Otherwise → “No unsubscribe method was found”

The body scanner examines anchor text, aria-label, title, image alt text, and URL strings, and incorporates multilingual unsubscribe terminology modeled on Microsoft's legacy add-in manifest. It uses body.getAsync(...Html...), which is supported in Read mode with Mailbox 1.3 and ReadItem permissions, so your existing Mailbox 1.8 manifest already covers it.

### Example

A message might contain:

```text
List-Unsubscribe: <mailto:list@example.com?subject=unsubscribe>,
 <https://example.com/unsubscribe/12345>
```

Because a web link is available, the add-in uses the HTTPS link rather than the email method.

If the message contains only:

```text
List-Unsubscribe: <mailto:list@example.com?subject=unsubscribe&body=Please%20remove%20me>
```

Outlook opens a new draft similar to:

```text
To: list@example.com
Subject: unsubscribe

Please remove me
```

### Technical Overview

This is an **Outlook web add-in**, not a Windows COM/VSTO plug-in. Outlook installs the XML manifest, then loads the HTML and JavaScript runtime over HTTPS from this GitHub Pages site.

The add-in:

- Uses `getAllInternetHeadersAsync()` to read the Internet/MIME headers of the message currently being viewed.
- Requires Outlook **Mailbox requirement set 1.8 or later**.
- Uses only the Outlook **ReadItem** permission.
- Uses `displayNewMessageForm()` to create a `mailto:` unsubscribe draft.
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
- A Microsoft mailbox supported by Outlook add-ins, such as:
  - Microsoft 365 / Exchange Online work or school account
  - Outlook.com / Hotmail account
- Internet access, because Outlook web add-ins load their runtime files over HTTPS.

> **Note:** Outlook can connect to accounts such as Gmail and Yahoo, but Outlook add-ins are not supported for those non-Microsoft accounts in Outlook on Windows or Outlook on the web.


## Source code

Project repository:
https://github.com/awsles/unsubscribe/

Hosted add-in files:
https://awsles.github.io/unsubscribe/

### Update History

- v1.0.0.4 -- *Initial Version*
- v1.0.0.5 -- Add support for extracting Unsubscribe link from message body

## Microsoft documentation

For additional details, see Microsoft's documentation:

- Sideload Outlook add-ins: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing
- Outlook add-ins overview: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/read-scenario
- Internet header APIs: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/internet-headers

