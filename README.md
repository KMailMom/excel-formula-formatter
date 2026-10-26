# Excel Formula Formatter

An Excel add-in that beautifies or minifies the formula in the selected cell, from the
ribbon or with a keyboard shortcut.

The ribbon adds a **Formula Formatter** tab with two groups:

- **Selected cell**
  - **Beautify** (`Ctrl+Alt+Shift+B`) spreads the formula over indented lines.
  - **Minify** (`Ctrl+Alt+Shift+M`) collapses it to a single line.
  - **Options** opens the task pane with two tabs:
    - **Settings**: layout density (Expanded, Spacious, Balanced, Compact, Dense) with a live
      preview, indentation, single-line spacing (space after separators, spaces inside
      brackets), the argument separator for the editor, and the keyboard shortcuts.
    - **Editor**: load a formula, beautify/minify it, review, and write it back.
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
canonical en-US, comma-separated form), so cells always end up with Excel's own
separator (for example a Dutch `;`). The separator setting only affects how formulas
are shown in the task pane.

- The formatter (`src/lib/formatter.ts`) tokenizes the formula, keeping string
  literals, quoted sheet names, structured references such as `[@[Termijn van]]` and
  array constants verbatim, builds a small tree, and renders it. With a density other
  than Expanded, a function call whose one-line form fits the width limit stays on one
  line.
- Safety check: the output must reduce to exactly the same formula as the input once
  insignificant whitespace is stripped, otherwise the cell is left unchanged.
- Whitespace between two references (the intersection operator, `=SUM(A1:A10 B1:B10)`)
  is significant and is kept.

## Cell-edit mode

Excel blocks all add-in access to the workbook while a cell is being edited, and the
text you are typing is not exposed to add-ins at all. Beautify/Minify therefore cannot
work inside edit mode. Press `Ctrl+Enter` to confirm the edit (the cell stays
selected), then press the shortcut.

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
trust prompt and the WebView2 loopback exemption (answer `Y`). A **Formula Formatter**
tab appears on the ribbon.

Stop with:

```
npm stop
```

Type-check and validate the manifest at any time:

```
npm run typecheck
npm test
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
   Confirm `https://<your-user>.github.io/excel-formula-formatter/taskpane.html` and
   `.../shortcuts.json` load over HTTPS.
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
manifest.xml                 add-in manifest (ribbon, shared runtime, resources)
shortcuts.json               keyboard shortcut actions and defaults
webpack.config.js            build + dev server (edit urlProd before deploying)
src/lib/formatter.ts         tokenizer, beautify / minify, separator conversion
src/lib/settings.ts          stored preferences and density levels
src/lib/shortcuts.ts         keyboard shortcut helpers
src/lib/dialog.ts            message / confirmation dialog helper
src/commands/commands.ts     ribbon + shortcut handlers (run in the shared runtime)
src/taskpane/                task pane (Settings + Editor), entry of the shared runtime
src/dialog/dialog.html       dialog page
assets/                      ribbon icons; regenerate with `npm run icons`
test/formatter.test.ts       formatter tests (`npm test`)
```

## License

This project is MIT licensed.
