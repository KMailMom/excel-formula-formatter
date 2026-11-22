# Security and trust

This page explains what Formula Formatter for Excel can and cannot do, and how you can check that
yourself instead of taking our word for it.

## What the add-in does

- It reads the formulas of the selected cells (or, for Beautify All / Minify All, of the whole
  workbook), reformats the whitespace, and writes them back. A safety check refuses to write a
  formula that differs from the original in anything but whitespace.
- It stores your preferences in the add-in's browser storage on your device, and custom keyboard
  shortcuts in your Office settings.

## What it does not do

- **No network traffic.** No formula, cell value or file content leaves Excel. There is no
  server, no analytics, no tracking and no cookies.
- The only external resource is Microsoft's Office.js library
  (`https://appsforoffice.microsoft.com/lib/1/hosted/office.js`), which every Office add-in loads.
- It cannot access files on your computer or other programs. Office add-ins run in a sandboxed
  browser control, and the manifest only requests `ReadWriteDocument` (the open workbook).

## How you can verify this

1. **Automated network audit.** [`scripts/audit-bundle.mjs`](scripts/audit-bundle.mjs) scans the
   built add-in and fails if it uses any API that can send data (`fetch`, `XMLHttpRequest`,
   `WebSocket`, `EventSource`, `sendBeacon`, `RTCPeerConnection`, `importScripts`, `eval`,
   `new Function`) or references any URL other than Office.js. It runs in
   [the deploy workflow](.github/workflows/deploy.yml) before every deploy; a failure stops the
   deploy. Run it yourself with `npm run build && npm run audit:network`. It is a pattern
   check: it catches network code added by mistake or by a careless dependency, not code
   deliberately disguised to evade it. That is what points 2 and 3 are for.
2. **Public builds.** The hosted add-in is built from this repository's `main` branch by GitHub
   Actions and published to the `gh-pages` branch. Build logs are public under the Actions tab,
   and you can compare the `gh-pages` branch with your own `npm run build`.
3. **Readable source.** All code is in [`src/`](src). The hosted site also serves
   `taskpane.js.map`, so browser developer tools show the original source.
4. **Short installers.** [`install/install.ps1`](install/install.ps1) (Windows) writes one file to
   `%LOCALAPPDATA%\ExcelFormulaFormatter\` and one registry value under
   `HKCU\Software\Microsoft\Office\16.0\Wef\Developer`, and checks that the downloaded manifest
   has this add-in's Id. [`install/install-mac.sh`](install/install-mac.sh) copies the manifest
   into Excel's `wef` folder. [`install/uninstall.ps1`](install/uninstall.ps1) removes both.
   The double-clickable `install.cmd` / `uninstall.cmd` on the install site are these same
   scripts with a small header ([`install/cmd-wrapper.txt`](install/cmd-wrapper.txt)).
   No admin rights, no compiled executables.

   Why no `.exe`? An unsigned executable triggers stronger SmartScreen and antivirus warnings
   than a script, and nobody can read what it does. Signing removes the warnings but costs a
   yearly certificate tied to a verified legal identity. A readable script is the more
   verifiable option.

## What you are trusting

Excel loads the add-in code from GitHub Pages each time, which is how updates arrive
automatically. It also means you trust whoever can push to this repository's `main` branch.
If you would rather control updates yourself, fork the repository, enable GitHub Pages on the
fork (Settings, Pages, source: `gh-pages` branch) and install from your fork's site. The build
derives all URLs from the repository name, so a fork works without code changes; only generate
a new `<Id>` in `manifest.xml` if both versions should be installable side by side.

## Reporting a vulnerability

Please do not open a public issue for security problems. Use GitHub's
**Security > Report a vulnerability** on this repository (private vulnerability reporting).
