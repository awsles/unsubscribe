# Install / sideload instructions — Windows 11

There are two installation modes: development installation from localhost, and normal installation from an HTTPS-hosted production build.

## Windows 11 client behavior

- **Classic Outlook for Windows:** the add-in can use `Office.context.ui.openBrowserWindow()` and launch the unsubscribe URL in the system browser.
- **New Outlook for Windows:** Microsoft currently does not support `OpenBrowserWindowApi 1.1`; this project falls back to an Office dialog for HTTPS unsubscribe links. The dialog first loads this add-in's same-origin `redirect.html` page and then redirects to the unsubscribe site.


## Option A — Development installation from this source tree

Use this while changing or debugging the code.

1. Open PowerShell in the project folder.
2. Run:

   ```powershell
   npm install
   npm start
   ```

3. Accept the localhost certificate and Edge WebView loopback prompts if they appear.
4. Sign into the Microsoft 365 account used by Outlook if Microsoft's sideloading tooling asks you to authenticate.
5. Open Outlook and then a received message.
6. Look for **Unsubscribe** in the message's add-in/app command area.

When finished testing, run:

```powershell
npm run stop
```

### Manual sideload fallback

If automatic sideloading does not work, keep the local server running:

```powershell
npm run dev-server
```

Then open this Microsoft shortcut in your browser:

```text
https://aka.ms/olksideload
```

In the **Add-Ins for Outlook** dialog:

1. Select **My add-ins**.
2. Scroll to **Custom Addins**.
3. Select **Add a custom add-in**.
4. Select **Add from File**.
5. Choose this project's `manifest.xml`.
6. Accept the installation prompts.

The local HTTPS server must remain running because `manifest.xml` points Outlook to `https://localhost:3000`.

## Option B — Install a hosted production build

Use this when you want the add-in to work without running Node.js on your PC.

### 1. Host the web files

Deploy everything under `dist\` to a public HTTPS location. Example:

```text
https://mailtools.example.com/outlook-unsubscribe/
```

Confirm these URLs work in a browser:

```text
https://mailtools.example.com/outlook-unsubscribe/commands.html
https://mailtools.example.com/outlook-unsubscribe/commands.js
https://mailtools.example.com/outlook-unsubscribe/assets/icon-32.png
```

### 2. Generate the production manifest

From PowerShell in the project folder:

```powershell
.\scripts\Create-ProductionManifest.ps1 `
  -BaseUrl "https://mailtools.example.com/outlook-unsubscribe"
```

This creates `manifest.production.xml`.

### 3. Install the production manifest

Open:

```text
https://aka.ms/olksideload
```

Then:

1. Choose **My add-ins**.
2. Under **Custom Addins**, choose **Add a custom add-in**.
3. Choose **Add from File**.
4. Select `manifest.production.xml`.
5. Accept the prompts.

The add-in is associated with the mailbox/account. It can then appear in supported Outlook clients signed into that account.

### Classic Outlook note

Microsoft notes that manually sideloaded add-ins can take substantially longer to appear in classic Outlook for Windows because of caching. New Outlook and Outlook on the web are generally better development/test targets when you want changes to appear quickly.

## Removing the add-in

For a development session started using `npm start`, run:

```powershell
npm run stop
```

For a manually installed custom add-in, return to **Add-Ins for Outlook → My add-ins → Custom Addins** and remove the custom add-in there.


---

### Sideload into Outlook (UPDATED)

1. **Download `manifest.xml` to your Windows 11 PC.** In GitHub, open `manifest.xml` → click **Raw/Download raw file** and save it somewhere convenient such as Downloads.

2. In your browser, open Microsoft's Outlook sideload page:
   [Open Outlook add-in sideloading](https://aka.ms/olksideload?utm_source=chatgpt.com)
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

The installation applies to the mailbox/account, so although you sideload through Outlook on the web, Microsoft says the add-in should also become available in supported Outlook desktop clients. In **classic Outlook for Windows**, Microsoft notes that a manually sideloaded add-in can occasionally take up to **24 hours** to appear because of caching. ([Microsoft Learn][2])

[1]: https://github.com/awsles/unsubscribe/blob/main/manifest.xml "unsubscribe/manifest.xml at main · awsles/unsubscribe · GitHub"
[2]: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?utm_source=chatgpt.com "Sideload Outlook add-ins for testing - Office Add-ins | Microsoft Learn"
