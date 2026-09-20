# Build and development instructions — Windows 11

These instructions are for developing and testing the add-in from source on Windows 11.

## 1. Prerequisites

Install:

- Windows 11
- A supported Outlook client signed into the mailbox you will use for testing (new Outlook, classic Outlook, or Outlook on the web)
- Node.js LTS, including npm
- Optional but recommended: Visual Studio Code

After installing Node.js, open **PowerShell** and verify:

```powershell
node --version
npm --version
```

## 2. Install the project dependencies

Open PowerShell in this project folder and run:

```powershell
npm install
```

This installs the local development web server, Microsoft's Office add-in debugging tools, Microsoft's local HTTPS development-certificate utility, and the manifest validator.

## 3. Validate the Outlook manifest

Run:

```powershell
npm run validate
```

The manifest declares:

- Outlook/Mailbox as the host
- Mailbox requirement set 1.8
- `ReadItem` permission
- a `MessageReadCommandSurface` ribbon command
- an `ExecuteFunction` action named `unsubscribe`
- the local HTTPS function page at `https://localhost:3000/commands.html`

## 4. Start the add-in for development

Run:

```powershell
npm start
```

On the first run, Microsoft's tooling may ask you to:

- trust/install a localhost HTTPS development certificate;
- allow a Microsoft Edge WebView localhost loopback exemption;
- sign into your Microsoft 365 account for sideloading.

Accept the certificate and loopback prompts. If the Microsoft 365 sign-in flow does not appear and sideloading fails, use Microsoft's current Office add-in authentication tooling as prompted by `office-addin-debugging`, then run `npm start` again.

The tool starts an HTTPS development server on port 3000 and attempts to sideload the manifest into Outlook.

## 5. Test the command

Open a received email message. Find **Unsubscribe** under the add-in/app commands for the message and select it.

Expected behavior:

- In **classic Outlook for Windows**, an HTTP/HTTPS unsubscribe URI opens in the system browser.
- In **new Outlook for Windows** (and Outlook on the web), HTTPS unsubscribe links open in an Office dialog that immediately redirects to the unsubscribe site because `OpenBrowserWindowApi 1.1` is not currently supported there.
- Header exists but only contains `mailto:` → Outlook displays an informational message.
- Header does not exist → Outlook displays `No List-Unsubscribe header was found in this message.`
- Header cannot be read → Outlook displays an error notification.

The button itself remains static because Outlook does not expose a supported dynamic ribbon API for changing this command's label/icon/state based on an arbitrary MIME header.

## 6. Stop the development session

Always stop the Office debugging registration cleanly:

```powershell
npm run stop
```

Do this instead of only closing Outlook or the terminal. Microsoft's debugging tooling registers the add-in for the session and `npm run stop` removes that registration.

## 7. Create a production build

Run:

```powershell
npm run build
```

The deployable web files are written to:

```text
dist\
    commands.html
    commands.js
    support.html
    redirect.html
    assets\
```

These are static files. They can be hosted on IIS, Azure Static Web Apps, Azure Storage static website hosting, GitHub Pages with HTTPS, or another HTTPS web host.

## 8. Create a production manifest

After deploying the contents of `dist` to a public HTTPS URL, for example:

```text
https://mailtools.example.com/outlook-unsubscribe
```

create a manifest using:

```powershell
.\scripts\Create-ProductionManifest.ps1 `
  -BaseUrl "https://mailtools.example.com/outlook-unsubscribe"
```

This creates:

```text
manifest.production.xml
```

Verify that all generated URLs are reachable by the Outlook client before installing that manifest.

## Project structure

```text
outlook-unsubscribe-addin\
│   manifest.xml
│   package.json
│   webpack.config.js
│   BUILD-WINDOWS11.md
│   INSTALL-WINDOWS11.md
│   README.md
│
├── scripts\
│   └── Create-ProductionManifest.ps1
│
└── src\
    │   commands.html
    │   commands.js
    │   support.html
    │   redirect.html
    └── assets\
        icon-16.png
        icon-32.png
        icon-64.png
        icon-80.png
        icon-128.png
```
