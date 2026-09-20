# Outlook Unsubscribe — GitHub Pages deployment

This folder contains the static files required by the Outlook add-in manifest and its runtime dependencies.

Expected GitHub Pages base URL:

    https://awsles.github.io/unsubscribe/

Files referenced directly by `manifest.xml`:

- `commands.html`
- `support.html`
- `assets/icon-16.png`
- `assets/icon-32.png`
- `assets/icon-64.png`
- `assets/icon-80.png`
- `assets/icon-128.png`

Additional runtime files required by the implementation:

- `commands.js` — loaded by `commands.html`
- `redirect.html` — used as the HTTPS redirect bridge in clients without `OpenBrowserWindowApi 1.1`

## Publish with GitHub Pages

1. Copy these files to the root of the `awsles/unsubscribe` repository.
2. In GitHub, open **Settings > Pages**.
3. Under **Build and deployment**, select **Deploy from a branch**.
4. Select the branch you use (normally `main`) and folder `/ (root)`.
5. Save and wait for Pages to publish.
6. Verify that this URL loads in a browser:

       https://awsles.github.io/unsubscribe/commands.html

7. Sideload the included `manifest.xml` into Outlook.

If you deploy the static files somewhere other than GitHub Pages, replace every occurrence of `https://awsles.github.io/unsubscribe` in `manifest.xml` with your HTTPS origin/path.
