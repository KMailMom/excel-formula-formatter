# Excel Formula Formatter

An Excel add-in that beautifies or minifies the formula in the selected cell, at the
click of a ribbon button. The formatting engine is the MIT-licensed
[excelFormulaUtilitiesJS](https://github.com/joshbtn/excelFormulaUtilitiesJS), run
directly inside Excel through an Office.js add-in.

The ribbon adds a **Formula Tools** tab with two groups:

- **Selected cell**
  - **Beautify** expands the formula in the selected cell(s) onto indented multiple lines.
  - **Minify** collapses it back to a single line.
  - **Options** opens a task pane to choose indentation (spaces vs tab, size) and preview
    changes before writing them back.
- **Whole workbook**
  - **Beautify All** / **Minify All** reformat every formula across all sheets at once.
    Each asks for confirmation first and reports how many formulas changed. The whole
    operation is undoable with Ctrl+Z. Sheets with an unusually large used range
    (over 200,000 cells) are skipped and reported rather than processed.

Your formula data never leaves your PC. The add-in code is static; it reads and writes
cells locally through the Excel JavaScript API.

## How it works

Excel treats whitespace and newlines *between tokens* in a formula as insignificant
formatting that it stores and shows in the (expanded) formula bar. Beautify inserts
that formatting; minify strips it. The add-in reads and writes `range.formulas` (the
canonical en-US, comma-separated form), so it works regardless of your local list
separator (for example a Dutch `;`).

- **Beautify** uses the library's formatter plus a safety check: a beautified formula
  must reduce back to the same formula once whitespace is stripped, otherwise the cell
  is left unchanged. Beautify never silently corrupts a formula.
- **Minify** is a quote-aware whitespace stripper that never rewrites token contents, so
  string literals, sheet names, and array constants survive intact.

**Known limitation:** an implicit-intersection space between two ranges
(`=SUM(A1:A10 A1:C1)`) is removed by minify. This operator is very rarely used.

## Prerequisites

- Windows 11 with Microsoft 365 Excel (desktop)
- Node.js 18+ and npm

## Develop and try it (auto-sideload)

```
npm install
npm start
```

`npm start` builds the project, starts an HTTPS dev server on `https://localhost:3000`,
and sideloads the add-in into desktop Excel. On first run, accept the dev-certificate
trust prompt and the WebView2 loopback exemption (answer `Y`). A **Formula Tools** tab
appears on the ribbon with the Beautify, Minify, and Options buttons.

Stop with:

```
npm stop
```

Type-check and validate the manifest at any time:

```
npm run typecheck
npm run validate
```

## Install permanently (GitHub Pages + trusted folder)

The dev server only runs while `npm start` is active. For an always-on install, host the
built files on GitHub Pages and register the manifest from a trusted folder.

1. **Set your production URL.** In [`webpack.config.js`](webpack.config.js), change
   `urlProd` to your GitHub Pages URL, for example
   `https://<your-user>.github.io/excel-formula-formatter/` (keep the trailing slash).
2. **Build:**
   ```
   npm run build
   ```
   The `dist/` folder now contains the add-in with production URLs baked into
   `dist/manifest.xml`.
3. **Publish `dist/` to GitHub Pages.** For example, from a repo whose Pages source is the
   `gh-pages` branch:
   ```
   npx gh-pages -d dist
   ```
   Confirm `https://<your-user>.github.io/excel-formula-formatter/commands.html` and
   `.../excel-formula.js` load over HTTPS.
4. **Register the manifest from a trusted folder:**
   - Create a folder and share it (right-click, Properties, Sharing, Share, Read/Write).
     Note its UNC path, for example `\\PC-NAME\AddinCatalog`.
   - In Excel: File, Options, Trust Center, Trust Center Settings, Trusted Add-in
     Catalogs. Paste the UNC path, click Add catalog, tick "Show in Menu", OK, then
     restart Excel.
   - Copy `dist/manifest.xml` (the one with your Pages URLs) into that shared folder.
   - In Excel: Home, Add-ins, More Add-ins, Shared Folder tab, select
     "Excel Formula Formatter", Add.

When you change code, re-run `npm run build` and re-publish `dist/`. When you change the
manifest (icons, labels, buttons), bump `<Version>` in `manifest.xml`, clear the Office
cache (`%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`), and restart Excel.

## Project layout

```
manifest.xml                 add-in manifest (ribbon buttons, resources)
webpack.config.js            build + dev server (edit urlProd before deploying)
src/lib/formatter.ts         beautify / minify logic (verified core)
src/lib/excel-formula.js     vendored MIT library (window.excelFormulaUtilities)
src/types/excel-formula.d.ts ambient types for the vendored global
src/commands/                ribbon function-command handlers + host page
src/taskpane/                Options pane (indent settings + preview)
src/dialog/dialog.html       small message dialog for warnings/errors
assets/                      placeholder ribbon icons (replace as you like)
```

## License

This project is MIT licensed. It bundles excelFormulaUtilitiesJS (MIT); see
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
