// Renders the ribbon and add-in icons from the inline SVG designs below into PNGs
// at every size the manifest references. Run with: npm run icons
// Needs Playwright with Chromium (npm i -D playwright && npx playwright install chromium).
import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";

const { chromium } = createRequire(import.meta.url)("playwright");
const dir = path.dirname(fileURLToPath(import.meta.url));

// Colors are mid-tones so the icons read on both the light and the dark ribbon.
const GREEN = "#1f9d55";
const AMBER = "#f2a33a";
const BLUE = "#3b8fe0";
const GRAY = "#8a8a8a";

// All designs use a 32x32 grid with shapes on even coordinates, so they stay crisp at 16 px.
const bar = (x, y, w, color = GREEN) => `<rect x="${x}" y="${y}" width="${w}" height="4" rx="1" fill="${color}"/>`;
const sparkle = (cx, cy, r, color = AMBER) =>
  `<path d="M${cx} ${cy - r} Q${cx + r * 0.18} ${cy - r * 0.18} ${cx + r} ${cy} Q${cx + r * 0.18} ${cy + r * 0.18} ${cx} ${cy + r} Q${cx - r * 0.18} ${cy + r * 0.18} ${cx - r} ${cy} Q${cx - r * 0.18} ${cy - r * 0.18} ${cx} ${cy - r}Z" fill="${color}"/>`;

const beautifyGlyph = () => bar(2, 6, 12) + bar(8, 14, 16) + bar(2, 22, 8) + sparkle(25, 7, 6);

const minifyGlyph = () =>
  bar(2, 14, 28) +
  `<path d="M10 3 L16 9 L22 3" fill="none" stroke="${BLUE}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>` +
  `<path d="M10 29 L16 23 L22 29" fill="none" stroke="${BLUE}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>`;

function gear() {
  const cx = 16, cy = 16, teeth = 8, rOut = 14, rIn = 10.5;
  let d = "";
  for (let i = 0; i < teeth; i++) {
    const a = (i / teeth) * Math.PI * 2;
    const half = Math.PI / teeth;
    const pts = [
      [a - half * 0.95, rIn], [a - half * 0.45, rOut], [a + half * 0.45, rOut], [a + half * 0.95, rIn],
    ];
    for (const [ang, r] of pts) {
      const x = (cx + r * Math.cos(ang)).toFixed(2), y = (cy + r * Math.sin(ang)).toFixed(2);
      d += (d ? " L" : "M") + x + " " + y;
    }
  }
  d += " Z";
  return `<path d="${d}" fill="${GRAY}"/><circle cx="16" cy="16" r="10.5" fill="${GRAY}"/>` +
    `<circle cx="16" cy="16" r="4.5" fill="none" stroke="#ffffff" stroke-width="3"/>`;
}

// Workbook badge for the "All" variants: a green sheet with a white grid.
const badge = () =>
  `<rect x="17" y="17" width="15" height="15" rx="2" fill="#107c41"/>` +
  `<path d="M17 22h15M17 27h15M22.5 17v15" stroke="#ffffff" stroke-width="1.5"/>`;

const scaled = (glyph, s) => `<g transform="scale(${s})">${glyph}</g>`;

const logo = () =>
  `<rect x="0" y="0" width="32" height="32" rx="6" fill="#107c41"/>` +
  `<g transform="translate(4 4) scale(0.75)">${bar(2, 6, 12, "#fff") + bar(8, 14, 16, "#fff") + bar(2, 22, 8, "#fff") + sparkle(25, 7, 6)}</g>`;

const ICONS = {
  beautify: beautifyGlyph(),
  minify: minifyGlyph(),
  options: gear(),
  "beautify-all": scaled(beautifyGlyph(), 0.7) + badge(),
  "minify-all": scaled(minifyGlyph(), 0.7) + badge(),
};
const RIBBON_SIZES = [16, 20, 24, 32, 40, 48, 64, 80];
const LOGO_SIZES = [16, 32, 64, 80, 128];

const svg = (body, size) =>
  `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 32 32">${body}</svg>`;

const browser = await chromium.launch();
const page = await browser.newPage();
async function render(body, size, file) {
  await page.setViewportSize({ width: size, height: size });
  await page.setContent(`<html><body style="margin:0;background:transparent">${svg(body, size)}</body></html>`);
  await page.screenshot({ path: path.join(dir, file), omitBackground: true, clip: { x: 0, y: 0, width: size, height: size } });
}
for (const [name, body] of Object.entries(ICONS)) {
  for (const size of RIBBON_SIZES) await render(body, size, `${name}-${size}.png`);
}
for (const size of LOGO_SIZES) await render(logo(), size, `logo-${size}.png`);

// Contact sheet on light and dark ribbon backgrounds, for review only (not deployed).
if (process.env.ICON_PREVIEW) {
  const row = (bg, fg) =>
    `<div style="background:${bg};color:${fg};padding:10px;display:flex;gap:18px;font:12px Segoe UI,sans-serif">` +
    Object.entries(ICONS)
      .map(([n, b]) => `<div style="text-align:center">${svg(b, 32)}<br>${svg(b, 16)}<div>${n}</div></div>`)
      .join("") +
    `<div style="text-align:center">${svg(logo(), 32)}<div>logo</div></div></div>`;
  await page.setViewportSize({ width: 560, height: 180 });
  await page.setContent(`<html><body style="margin:0">${row("#f3f3f3", "#242424")}${row("#292929", "#ffffff")}</body></html>`);
  await page.screenshot({ path: process.env.ICON_PREVIEW, fullPage: true });
}
await browser.close();
console.log("icons written to", dir);
