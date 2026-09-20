# Vibe Coding Notes
Conversation with ChatGPT to build this quickly.

---
Yes. An Outlook web add-in can read the raw Internet headers of the message currently being viewed, look for `List-Unsubscribe`, extract an HTTPS unsubscribe URL, and—after the user clicks your command—open that URL in the system browser.

Microsoft exposes `Office.context.mailbox.item.getAllInternetHeadersAsync()` for messages in **Read mode**. It requires Outlook Mailbox requirement set **1.8+** and the `ReadItem` permission. ([Microsoft Learn][1])

A basic implementation could look like this:

```javascript
function unsubscribe(event) {
    const item = Office.context.mailbox.item;

    item.getAllInternetHeadersAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(result.error);
            event.completed();
            return;
        }

        const headers = result.value;

        // Find List-Unsubscribe, including a folded continuation line.
        const match = headers.match(
            /^List-Unsubscribe:\s*(.*(?:\r?\n[ \t]+.*)*)/im
        );

        if (!match) {
            console.log("No List-Unsubscribe header.");
            event.completed();
            return;
        }

        // Unfold MIME header.
        const value = match[1].replace(/\r?\n[ \t]+/g, " ");

        // Prefer an HTTP/HTTPS unsubscribe method over mailto:
        const urls = [...value.matchAll(/<([^>]+)>/g)]
            .map(m => m[1]);

        const unsubscribeUrl =
            urls.find(url => /^https:\/\//i.test(url)) ??
            urls.find(url => /^http:\/\//i.test(url));

        if (unsubscribeUrl) {
            Office.context.ui.openBrowserWindow(unsubscribeUrl);
        } else {
            console.log("List-Unsubscribe exists, but has no web URL.");
        }

        event.completed();
    });
}
```

`Office.context.ui.openBrowserWindow()` is specifically intended for opening an external HTTPS/HTTP URL from an Office add-in; Microsoft recommends it instead of `window.open()`. ([Microsoft Learn][2])

For example, given:

```text
List-Unsubscribe: <mailto:unsubscribe@example.com>,
 <https://example.com/unsubscribe?id=123456>
```

your add-in would extract:

```text
https://example.com/unsubscribe?id=123456
```

and clicking your **Unsubscribe** command could launch that page.

There is one important UI distinction. **Putting an Unsubscribe button on the Outlook ribbon and checking the header when the user clicks it is straightforward.** Making that native Outlook ribbon button automatically appear/disappear or enable/disable *solely according to the presence of `List-Unsubscribe`* is a different matter. Outlook's manifest activation rules don't provide an arbitrary-MIME-header condition, and the current programmatic ribbon enable/disable mechanism documented by Microsoft isn't generally available to Outlook the way it is to Word/Excel/PowerPoint. ([Microsoft Learn][3])

So I would architect it as:

**Outlook message → Unsubscribe command → read MIME headers → find `List-Unsubscribe` → open HTTPS URL**

You can make the experience nicer by having the command display a small task pane/status message saying **“This message does not provide an unsubscribe link”** when the header isn't present.

There's also a useful second phase to this project: support **RFC 8058 one-click unsubscribe**. Messages may contain:

```text
List-Unsubscribe: <https://example.com/unsubscribe/123>
List-Unsubscribe-Post: List-Unsubscribe=One-Click
```

That case is subtly different: a standards-compliant one-click implementation sends a **POST** rather than simply navigating the user to the URL. So I would distinguish between **“Open unsubscribe page”** and a future **“One-click unsubscribe”** implementation rather than treating them as the same operation.

For the project you're describing, I would start with the browser-launch version because it is simple, requires no backend service, and doesn't require Graph/EWS access—the Outlook JavaScript API can read the header directly. ([Microsoft Learn][1])

If you want, I can next lay out the complete minimal Outlook add-in—**manifest + ribbon button + JavaScript/TypeScript implementation**—for exactly this feature.

[1]: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/internet-headers?utm_source=chatgpt.com "Get and set internet headers on a message in an Outlook add-in - Office Add-ins | Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins?utm_source=chatgpt.com "Browsers and webview controls used by Office Add-ins - Office Add-ins | Microsoft Learn"
[3]: https://learn.microsoft.com/en-us/office/dev/add-ins/design/disable-add-in-commands?utm_source=chatgpt.com "Enable and Disable Add-in Commands - Office Add-ins | Microsoft Learn"


---

Not in the exact way you’re describing, at least not with a standard Outlook add-in ribbon command.

For Outlook, the ribbon APIs that can programmatically enable/disable commands are not supported; Microsoft currently lists `RibbonApi 1.1` support for Excel, Word, and PowerPoint, but not Outlook. ([Microsoft Learn][1]) Also, even where the Ribbon API is supported, the runtime control object only exposes `enabled`, `visible`, and `id`; it does not let you dynamically change the button label or icon. ([Microsoft Learn][2])

So in Outlook, the practical design is:

* Keep an **Unsubscribe** button visible and clickable all the time.
* When clicked, call `getAllInternetHeadersAsync()`.
* If `List-Unsubscribe` contains a web URL, open it.
* If it doesn't, show a notification such as **“No unsubscribe link was found in this message.”**

That last part works quite nicely with Outlook's notification APIs. The user doesn't have to open a task pane; you can show the message directly on the email item.

For example:

```javascript
function unsubscribe(event) {
    const item = Office.context.mailbox.item;

    item.getAllInternetHeadersAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            showMessage("Unable to inspect this message.");
            event.completed();
            return;
        }

        const headers = result.value;

        const match = headers.match(
            /^List-Unsubscribe:\s*(.*(?:\r?\n[ \t]+.*)*)/im
        );

        if (!match) {
            showMessage("No unsubscribe link was found in this message.");
            event.completed();
            return;
        }

        const value = match[1].replace(/\r?\n[ \t]+/g, " ");

        const links = [...value.matchAll(/<([^>]+)>/g)]
            .map(m => m[1]);

        const url = links.find(x => /^https?:\/\//i.test(x));

        if (!url) {
            showMessage(
                "This message provides an unsubscribe method, but not a web link."
            );
            event.completed();
            return;
        }

        Office.context.ui.openBrowserWindow(url);
        event.completed();
    });
}

function showMessage(message) {
    Office.context.mailbox.item.notificationMessages.replaceAsync(
        "unsubscribeStatus",
        {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: message,
            icon: "Icon.16x16",
            persistent: false
        }
    );
}
```

Visually, it would behave approximately like:

```text
┌────────────────────────────── Outlook Ribbon ──────────────────────────┐
│ Reply   Reply All   Forward      🛑 Unsubscribe                       │
└────────────────────────────────────────────────────────────────────────┘

                         ↓ click

Message HAS List-Unsubscribe
        → browser opens unsubscribe page

Message DOES NOT
        ↓

┌────────────────────────────────────────────────────────────────────────┐
│ ℹ No unsubscribe link was found in this message.                      │
└────────────────────────────────────────────────────────────────────────┘
```

There is one possible way to get closer to the UI you want: use **two different commands defined in the manifest**, for example a normal `Unsubscribe` and a visually muted `No Unsubscribe` command. But Outlook doesn't give us a way to switch between those commands based on an arbitrary MIME header, so it doesn't really solve the problem.

I therefore wouldn't try to make the ribbon state convey the result. I'd make the button static and make the **click behavior intelligent**. An additional benefit is that it handles more than just “header/no header”: it can distinguish a web unsubscribe URL, a `mailto:` unsubscribe, RFC 8058 one-click unsubscribe, or malformed headers. That could make the button much more useful than Outlook's UI state alone.

[1]: https://learn.microsoft.com/en-us/javascript/api/requirement-sets/common/ribbon-api-requirement-sets?view=word-js-preview&utm_source=chatgpt.com "Ribbon API requirement sets - Office Add-ins | Microsoft Learn"
[2]: https://learn.microsoft.com/nb-no/javascript/api/office/office.control?view=word-js-preview&viewFallbackFrom=visio-js-1.1 "Office.Control interface - Office Add-ins | Microsoft Learn"


---

Yes. For this feature, I’d build an **Outlook Web Add-in** rather than an old-style Windows COM plug-in. It can work in classic Outlook, new Outlook, Outlook on the web, and Mac from the same JavaScript codebase. ([Microsoft Learn][1])

The development path is:

1. **Install the development tools.** Install the current Node.js LTS, then run:

   ```bash
   npm install -g yo generator-office
   ```

   Microsoft’s current Outlook quickstart uses the Yeoman Office generator for this. ([Microsoft Learn][2])

2. **Generate an Outlook project.**

   ```bash
   yo office
   ```

   For the initial project, choose something like:

   ```text
   Office Add-in Task Pane project
   JavaScript
   Outlook Unsubscribe
   Outlook
   Add-in only manifest
   ```

   Although the generator starts you with a task-pane example, we would change it into a **UI-less function command**: the ribbon button runs JavaScript directly without opening a pane. Outlook explicitly supports this model. ([Microsoft Learn][3])

3. **Define an Unsubscribe button in `manifest.xml`.** The important part would look roughly like this:

   ```xml
   <ExtensionPoint xsi:type="MessageReadCommandSurface">
     <OfficeTab id="TabDefault">
       <Group id="unsubscribeGroup">
         <Label resid="GroupLabel"/>

         <Control xsi:type="Button" id="unsubscribeButton">
           <Label resid="Unsubscribe.Label"/>

           <Supertip>
             <Title resid="Unsubscribe.Label"/>
             <Description resid="Unsubscribe.Description"/>
           </Supertip>

           <Icon>
             <bt:Image size="16" resid="Icon.16x16"/>
             <bt:Image size="32" resid="Icon.32x32"/>
             <bt:Image size="80" resid="Icon.80x80"/>
           </Icon>

           <Action xsi:type="ExecuteFunction">
             <FunctionName>unsubscribe</FunctionName>
           </Action>
         </Control>
       </Group>
     </OfficeTab>
   </ExtensionPoint>
   ```

   `MessageReadCommandSurface` puts the command on the message-reading UI/ribbon. ([Microsoft Learn][4])

4. **Give the add-in permission to inspect the message.** The manifest needs:

   ```xml
   <Permissions>ReadItem</Permissions>
   ```

   and the requirements should specify Mailbox 1.8 or later because `getAllInternetHeadersAsync()` was introduced in Mailbox 1.8. Microsoft documents it specifically for **Message Read** mode with minimum `ReadItem` permission. ([Microsoft Learn][5])

5. **Implement the button.** Your `commands.js` can be quite small:

   ```javascript
   Office.onReady(() => {
       Office.actions.associate("unsubscribe", unsubscribe);
   });

   function unsubscribe(event) {
       const item = Office.context.mailbox.item;

       item.getAllInternetHeadersAsync((result) => {

           if (result.status !== Office.AsyncResultStatus.Succeeded) {
               showStatus("Unable to inspect this message.");
               event.completed();
               return;
           }

           const headers = result.value;

           // Extract List-Unsubscribe, including folded MIME lines.
           const match = headers.match(
               /^List-Unsubscribe:\s*(.*(?:\r?\n[ \t]+.*)*)/im
           );

           if (!match) {
               showStatus("No unsubscribe link was found in this message.");
               event.completed();
               return;
           }

           const value =
               match[1].replace(/\r?\n[ \t]+/g, " ");

           const links =
               [...value.matchAll(/<([^>]+)>/g)]
                   .map(m => m[1]);

           // Prefer HTTPS.
           const url =
               links.find(x => /^https:\/\//i.test(x)) ??
               links.find(x => /^http:\/\//i.test(x));

           if (!url) {
               showStatus(
                   "An unsubscribe method exists, but no web link was provided."
               );
               event.completed();
               return;
           }

           Office.context.ui.openBrowserWindow(url);

           event.completed();
       });
   }

   function showStatus(text) {
       Office.context.mailbox.item.notificationMessages.replaceAsync(
           "unsubscribe-status",
           {
               type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
               message: text
           }
       );
   }
   ```

   `openBrowserWindow()` opens the URL in the user's external browser and is the Office-supported approach instead of relying on `window.open()`. ([Microsoft Learn][6])

6. **Run it locally.** From the project directory:

   ```bash
   npm start
   ```

   The generated project starts an HTTPS development server, installs/trusts a development certificate if necessary, and attempts to sideload the add-in into Outlook. Microsoft specifically recommends this route for Yeoman-generated Outlook add-ins. ([Microsoft Learn][7])

7. **Install it permanently.** During development, the files can live at something like:

   ```text
   https://localhost:3000/commands.html
   ```

   But for a version you want installed without running Node on your PC, you host those static HTML/JavaScript/image files on any HTTPS web server, for example:

   ```text
   https://yourdomain.com/outlook-unsubscribe/commands.html
   ```

   Then change the URLs in `manifest.xml` to point there.

   To manually install an add-in-only manifest, Microsoft currently directs you to the Outlook Add-Ins dialog, then **My add-ins → Custom Addins → Add a custom add-in → Add from File**, and select `manifest.xml`. The same installed add-in can then appear in supported Outlook desktop and web clients. ([Microsoft Learn][7])

The resulting architecture is extremely small:

```text
Outlook
   │
   │ click "Unsubscribe"
   ▼
manifest.xml
   │
   ▼
commands.html
   │
   ▼
commands.js
   │
   ├── getAllInternetHeadersAsync()
   │
   ├── no List-Unsubscribe
   │      └── show Outlook notification
   │
   └── List-Unsubscribe URL
          └── openBrowserWindow(url)
```

And importantly, **you don't need a server-side application, Microsoft Graph, OAuth flow, database, or Exchange administrator permissions** for this basic version. The add-in is simply a few static files plus the manifest. The currently selected message is inspected locally through Office.js. ([Microsoft Learn][5])

For your particular project, I would actually make the first release slightly smarter than the code above. I would recognize all three cases:

```text
List-Unsubscribe:
    HTTPS URL       → Open unsubscribe page

    mailto: only    → Explain that email-based unsubscribe is available

    nothing         → "No unsubscribe found"

List-Unsubscribe-Post:
    One-Click       → recognize RFC 8058 separately
```

That gives us a clean foundation for adding true **one-click unsubscribe** later without confusing it with simply visiting the unsubscribe webpage.

If you want to proceed, I can build the **complete ready-to-run project** for you—including `manifest.xml`, `commands.html`, `commands.js`, icons, `package.json`, and a README with the exact `npm start` and Outlook installation instructions—so you can download it, run it, and sideload it directly.

[1]: https://learn.microsoft.com/office/dev/add-ins/outlook/read-scenario?utm_source=chatgpt.com "Outlook add-ins overview - Office Add-ins | Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview?utm_source=chatgpt.com "Create Office Add-in projects using the Yeoman Generator - Office Add-ins | Microsoft Learn"
[3]: https://learn.microsoft.com/en-us/javascript/api/manifest/functionfile?view=word-js-preview&utm_source=chatgpt.com "FunctionFile element in the manifest file - Office Add-ins | Microsoft Learn"
[4]: https://learn.microsoft.com/en-us/javascript/api/manifest/extensionpoint?view=word-js-preview&utm_source=chatgpt.com "ExtensionPoint element in the manifest file - Office Add-ins | Microsoft Learn"
[5]: https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread?view=outlook-js-preview&utm_source=chatgpt.com "Office.MessageRead interface - Office Add-ins | Microsoft Learn"
[6]: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins?utm_source=chatgpt.com "Use the Office dialog API in your Office Add-ins - Office Add-ins | Microsoft Learn"
[7]: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?utm_source=chatgpt.com "Sideload Outlook add-ins for testing - Office Add-ins | Microsoft Learn"


---

Done. I built the complete Windows 11 Outlook add-in project.

[Download the complete Outlook Unsubscribe project](sandbox:/mnt/data/outlook-unsubscribe-addin.zip)

The package includes:

* `manifest.xml` — sideloadable Outlook add-in manifest.
* `src/commands.js` — reads `List-Unsubscribe`, parses it safely, and handles the button click.
* `src/commands.html` — UI-less Outlook function-command page.
* `src/redirect.html` — compatibility fallback for new Outlook.
* Ribbon icons in all required sizes.
* `dist\` — a prebuilt, ready-to-host static version.
* `package.json` and `webpack.config.js` — Windows development environment.
* PowerShell script for generating a production manifest.
* Separate build and installation guides.

You can also open the instructions directly:

[Build & development instructions for Windows 11](sandbox:/mnt/data/outlook-unsubscribe-addin/BUILD-WINDOWS11.md)

[Installation & sideload instructions for Windows 11](sandbox:/mnt/data/outlook-unsubscribe-addin/INSTALL-WINDOWS11.md)

[Project README](sandbox:/mnt/data/outlook-unsubscribe-addin/README.md)

### What it does

When you open a received message and click **Unsubscribe**, the add-in calls `getAllInternetHeadersAsync()`, which is available beginning with Outlook Mailbox requirement set 1.8. ([Microsoft Learn][1])

It handles messages such as:

```text
List-Unsubscribe: <mailto:remove@example.com>,
 <https://example.com/unsubscribe?id=12345>
```

and prefers the HTTPS web URL.

If there is no `List-Unsubscribe` header, Outlook displays:

> No List-Unsubscribe header was found in this message.

If there is only a `mailto:` method, it tells you an email-based unsubscribe method exists rather than automatically sending mail.

### Important Windows 11 Outlook difference

I changed the implementation slightly after checking Microsoft's current API support.

In **classic Outlook for Windows**, the add-in uses `Office.context.ui.openBrowserWindow()`, so the unsubscribe page opens in your normal Windows browser. Microsoft's current support table lists Outlook classic on Windows for `OpenBrowserWindowApi 1.1`. ([Microsoft Learn][2])

In **new Outlook for Windows**, that API is currently not supported. The project therefore falls back to an Office dialog for HTTPS unsubscribe URLs. It opens a same-origin `redirect.html` first and then redirects that dialog to the unsubscribe site, which follows Microsoft's documented dialog security model. ([Microsoft Learn][3])

### First thing to do on your Windows 11 machine

Extract the ZIP, open **PowerShell** in the extracted directory, and run:

```powershell
npm install
npm run validate
npm start
```

`npm start` uses Microsoft's Office add-in debugging tooling to start the local HTTPS server and attempt to sideload the add-in. Microsoft documents this as the recommended development workflow; manual sideloading through **My add-ins → Custom Addins → Add from File** is also supported. ([Microsoft Learn][4])

I syntax-checked the JavaScript, ran parser smoke tests against folded `List-Unsubscribe` headers, and verified the manifest is well-formed XML. I wasn't able to run Microsoft's npm-based manifest validator in this environment because the dependency installation timed out, so `npm run validate` is the first validation step I'd run on your Windows machine.

ZIP SHA-256:

```text
1b62c8e7bf87327d1fed0f14fd7fe59be19c79700863b13fd71860381941f21d
```

The logical next enhancement would be **RFC 8058 one-click unsubscribe**, where the add-in recognizes `List-Unsubscribe-Post: List-Unsubscribe=One-Click` and performs the appropriate POST instead of merely opening the webpage.

[1]: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/internet-headers?utm_source=chatgpt.com "Get and set internet headers on a message in an Outlook add-in - Office Add-ins | Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/javascript/api/office/office.ui?view=common-js-preview&utm_source=chatgpt.com "Office.UI interface - Office Add-ins | Microsoft Learn"
[3]: https://learn.microsoft.com/es-es/office/dev/add-ins/develop/dialog-api-in-office-add-ins?utm_source=chatgpt.com "Usar la API de diálogo de Office en los complementos de Office - Office Add-ins | Microsoft Learn"
[4]: https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing?utm_source=chatgpt.com "Sideload Outlook add-ins for testing - Office Add-ins | Microsoft Learn"






