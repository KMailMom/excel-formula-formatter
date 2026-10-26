/* global Office, Excel, document */
/*
 * Entry point of the shared runtime: the Options task pane UI plus (via the import
 * below) the ribbon / keyboard command handlers.
 */
import "../commands/commands";
import { beautifyFormula, minifyFormulaSafe, toDisplayFormula, fromDisplayFormula } from "../lib/formatter";
import { loadSettings, saveSettings, toFormatOptions, Settings, DENSITY_LEVELS } from "../lib/settings";
import { DEFAULT_SHORTCUTS, ShortcutAction, shortcutsSupported, getShortcuts, setShortcuts, shortcutsInUse, comboFromEvent } from "../lib/shortcuts";

const SAMPLE_FORMULA =
  '=IF([@Soort]<>"Extra","",LET(k,[@Werknr]&"|"&LEFT(LOWER(TRIM([@Klant]&"")),80)&"|"&N([@[Termijn van]]),IF(k="","-",k)))';

let settings: Settings;
/** Separator detected from Excel's regional settings (used when the setting is "auto"). */
let detectedSeparator: "," | ";" = ",";
let previewEl: HTMLTextAreaElement;
let statusEl: HTMLElement;

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

function displaySeparator(): "," | ";" {
  return settings.separator === "auto" ? detectedSeparator : settings.separator;
}

// ---- theme -----------------------------------------------------------------

function applyTheme(): void {
  let dark = typeof matchMedia === "function" && matchMedia("(prefers-color-scheme: dark)").matches;
  try {
    const bg = Office.context.officeTheme?.bodyBackgroundColor;
    if (bg && /^#?[0-9a-f]{6}$/i.test(bg)) {
      const hex = bg.replace("#", "");
      const [r, g, b] = [0, 2, 4].map((k) => parseInt(hex.slice(k, k + 2), 16));
      dark = 0.2126 * r + 0.7152 * g + 0.0722 * b < 128;
    }
  } catch {
    /* keep the system preference */
  }
  document.documentElement.classList.toggle("dark", dark);
}

// ---- tabs ------------------------------------------------------------------

function selectTab(name: "settings" | "editor"): void {
  for (const n of ["settings", "editor"]) {
    $(`tab-${n}`).setAttribute("aria-selected", String(n === name));
    $(`panel-${n}`).classList.toggle("active", n === name);
  }
}

// ---- settings UI -----------------------------------------------------------

function readSettingsFromUi(): Settings {
  const kind = (document.querySelector('input[name="indentKind"]:checked') as HTMLInputElement)?.value;
  const size = parseInt($<HTMLInputElement>("indentSize").value, 10);
  const sep = $<HTMLSelectElement>("separator").value;
  return {
    indent: { useSpaces: kind !== "tab", size: Number.isFinite(size) ? Math.min(8, Math.max(1, size)) : settings.indent.size },
    spacing: {
      afterSeparator: $<HTMLInputElement>("spaceAfterSep").checked,
      insideParens: $<HTMLInputElement>("spaceInParens").checked,
    },
    density: parseInt($<HTMLInputElement>("density").value, 10),
    separator: sep === ";" || sep === "," ? sep : "auto",
  };
}

function applySettingsToUi(s: Settings): void {
  const radio = document.querySelector(`input[name="indentKind"][value="${s.indent.useSpaces ? "spaces" : "tab"}"]`) as HTMLInputElement | null;
  if (radio) radio.checked = true;
  $<HTMLInputElement>("indentSize").value = String(s.indent.size);
  $<HTMLInputElement>("indentSize").disabled = !s.indent.useSpaces;
  $<HTMLInputElement>("spaceAfterSep").checked = s.spacing.afterSeparator;
  $<HTMLInputElement>("spaceInParens").checked = s.spacing.insideParens;
  $<HTMLInputElement>("density").value = String(s.density);
  $<HTMLSelectElement>("separator").value = s.separator;
}

function renderSettingsInfo(): void {
  const level = DENSITY_LEVELS[settings.density];
  $("densityName").textContent = level.label;
  $("densityDesc").textContent = level.description;
  $<HTMLInputElement>("density").setAttribute("aria-valuetext", level.label);
  $<HTMLInputElement>("indentSize").disabled = !settings.indent.useSpaces;

  const autoLabel = detectedSeparator === ";" ? "Automatic ( ; )" : "Automatic ( , )";
  ($<HTMLSelectElement>("separator").options[0] as HTMLOptionElement).textContent = autoLabel;

  // Preview the user's own formula when the editor holds one, else a sample.
  const source = toCanonical(previewEl.value.trim()) ?? SAMPLE_FORMULA;
  const result = beautifyFormula(source, toFormatOptions(settings));
  $("sample").textContent = toDisplayFormula(result.value, displaySeparator());
}

function onSettingsChanged(): void {
  const before = displaySeparator();
  settings = readSettingsFromUi();
  saveSettings(settings);
  // Keep the editor text in the newly chosen separator.
  const after = displaySeparator();
  if (before !== after && previewEl.value.trim()) {
    previewEl.value = toDisplayFormula(fromDisplayFormula(previewEl.value, before), after);
  }
  renderSettingsInfo();
}

async function detectSeparator(): Promise<void> {
  try {
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.11")) return;
    // Office.js does not expose the list separator. Excel cannot use "," for both
    // decimals and arguments, so a comma decimal separator implies ";" arguments.
    await Excel.run(async (context) => {
      const nf = context.application.cultureInfo.numberFormat;
      nf.load("numberDecimalSeparator");
      await context.sync();
      detectedSeparator = nf.numberDecimalSeparator === "," ? ";" : ",";
    });
  } catch {
    /* keep the default */
  }
}

// ---- keyboard shortcuts ----------------------------------------------------

let recording: HTMLButtonElement | null = null;

function shortcutStatus(text: string, cls = ""): void {
  const el = $("shortcutStatus");
  el.textContent = text;
  el.className = "note " + cls;
}

function prettyCombo(combo: string | null): string {
  return combo ? combo.split("+").join(" + ") : "Not assigned (overridden)";
}

async function refreshShortcuts(): Promise<void> {
  try {
    const current = await getShortcuts();
    $("kbdBeautify").textContent = prettyCombo(current.Beautify);
    $("kbdMinify").textContent = prettyCombo(current.Minify);
  } catch {
    $("kbdBeautify").textContent = prettyCombo(DEFAULT_SHORTCUTS.Beautify);
    $("kbdMinify").textContent = prettyCombo(DEFAULT_SHORTCUTS.Minify);
  }
}

function stopRecording(): void {
  if (recording) recording.classList.remove("recording");
  recording = null;
}

async function onShortcutKey(e: KeyboardEvent): Promise<void> {
  if (!recording) return;
  e.preventDefault();
  e.stopPropagation();
  if (e.key === "Escape") {
    stopRecording();
    await refreshShortcuts();
    shortcutStatus("Cancelled.");
    return;
  }
  const combo = comboFromEvent(e);
  if (!combo) return; // modifiers only, keep listening
  if (combo.startsWith("!")) {
    shortcutStatus(combo.slice(1) + " Try again, or press Esc.", "msg-warn");
    return;
  }
  const action = recording.dataset.action as ShortcutAction;
  stopRecording();
  try {
    const inUse = await shortcutsInUse([combo]).catch(() => [] as string[]);
    await setShortcuts({ [action]: combo });
    await refreshShortcuts();
    shortcutStatus(
      inUse.length
        ? `Saved ${combo}. Note: Excel or another add-in also uses it; Excel will ask which one to run the first time you press it.`
        : `Saved: ${action} is now ${combo}.`,
      inUse.length ? "msg-warn" : "msg-ok"
    );
  } catch (err) {
    await refreshShortcuts();
    shortcutStatus("Could not save that shortcut: " + ((err as Error)?.message ?? String(err)), "msg-warn");
  }
}

function initShortcuts(): void {
  if (!shortcutsSupported()) {
    document.querySelectorAll<HTMLButtonElement>("#shortcutCard button").forEach((b) => (b.disabled = true));
    $("kbdBeautify").textContent = prettyCombo(DEFAULT_SHORTCUTS.Beautify);
    $("kbdMinify").textContent = prettyCombo(DEFAULT_SHORTCUTS.Minify);
    shortcutStatus("Changing shortcuts needs a newer Excel version (KeyboardShortcuts 1.1). The defaults shown here still work.");
    return;
  }
  void refreshShortcuts();
  document.querySelectorAll<HTMLButtonElement>(".kbd").forEach((btn) =>
    btn.addEventListener("click", () => {
      stopRecording();
      recording = btn;
      btn.classList.add("recording");
      btn.textContent = "Press keys...";
      shortcutStatus("Press the new combination (Ctrl and/or Alt, optionally Shift, plus a letter or digit). Esc cancels.");
    })
  );
  document.querySelectorAll<HTMLButtonElement>("[data-reset]").forEach((btn) =>
    btn.addEventListener("click", async () => {
      stopRecording();
      const action = btn.dataset.reset as ShortcutAction;
      try {
        await setShortcuts({ [action]: null });
        await refreshShortcuts();
        shortcutStatus(`${action} restored to ${DEFAULT_SHORTCUTS[action]}.`, "msg-ok");
      } catch (err) {
        shortcutStatus("Could not reset: " + ((err as Error)?.message ?? String(err)), "msg-warn");
      }
    })
  );
  document.addEventListener("keydown", (e) => void onShortcutKey(e), true);
  document.addEventListener("mousedown", (e) => {
    if (recording && e.target !== recording) {
      stopRecording();
      void refreshShortcuts();
      shortcutStatus("Click a shortcut, then press the new key combination. Esc cancels.");
    }
  });
}

// ---- editor ----------------------------------------------------------------

function setStatus(message: string, isError = false): void {
  statusEl.textContent = message;
  statusEl.classList.toggle("err", isError);
}

/** A formula typed in the display separator (or the other one) in canonical form, or null. */
function toCanonical(src: string): string | null {
  if (src.charAt(0) !== "=") return null;
  const sep = displaySeparator();
  const canonical = fromDisplayFormula(src, sep);
  // Accept a pasted formula in either separator style.
  if (sep === "," && !minifyFormulaSafe(canonical).ok) {
    const alt = fromDisplayFormula(src, ";");
    if (minifyFormulaSafe(alt).ok) return alt;
  }
  return canonical;
}

/** Editor text converted to Excel's canonical (comma) form, or null with a status message. */
function editorFormula(emptyMsg: string): string | null {
  const src = previewEl.value.trim();
  if (!src) {
    setStatus(emptyMsg);
    return null;
  }
  const f = toCanonical(src);
  if (!f) setStatus('A formula must start with "=".', true);
  return f;
}

async function loadFromCell(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load(["formulas", "address"]);
      await context.sync();
      const value = cell.formulas[0][0];
      const text = typeof value === "string" ? value : String(value ?? "");
      previewEl.value = toDisplayFormula(text, displaySeparator());
      setStatus(text.charAt(0) === "=" ? `Loaded ${cell.address}.` : `${cell.address} has no formula.`);
    });
  } catch (err) {
    setStatus((err as { code?: string })?.code === "InvalidOperationInCellEditMode" ? "Finish editing the cell first (Enter or Ctrl+Enter)." : (err as Error).message, true);
  }
  renderSettingsInfo();
}

function doBeautify(): void {
  const f = editorFormula("Nothing to beautify.");
  if (!f) return;
  const result = beautifyFormula(f, toFormatOptions(settings));
  if (result.ok) {
    previewEl.value = toDisplayFormula(result.value, displaySeparator());
    setStatus("Beautified. Review, then write back to the cell.");
  } else {
    setStatus("Left unchanged: " + (result.reason ?? "unsupported syntax") + ".", true);
  }
}

function doMinify(): void {
  const f = editorFormula("Nothing to minify.");
  if (!f) return;
  const result = minifyFormulaSafe(f, settings.spacing);
  if (result.ok) {
    previewEl.value = toDisplayFormula(result.value, displaySeparator());
    setStatus("Minified. Review, then write back to the cell.");
  } else {
    setStatus("Left unchanged: " + (result.reason ?? "unsupported syntax") + ".", true);
  }
}

async function applyToCell(): Promise<void> {
  const value = editorFormula("Nothing to write.");
  if (!value) return;
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.formulas = [[value]];
      cell.load("address");
      await context.sync();
      setStatus(`Written to ${cell.address}.`);
    });
  } catch (err) {
    setStatus((err as Error).message, true);
  }
}

// ---- init ------------------------------------------------------------------

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) return;

  previewEl = $<HTMLTextAreaElement>("preview");
  statusEl = $("status");
  settings = loadSettings();
  applyTheme();
  try {
    Office.addin.onVisibilityModeChanged(() => applyTheme());
  } catch {
    /* not in a shared runtime */
  }

  applySettingsToUi(settings);
  renderSettingsInfo();

  $("tab-settings").addEventListener("click", () => selectTab("settings"));
  $("tab-editor").addEventListener("click", () => selectTab("editor"));
  document.querySelectorAll('input[name="indentKind"], #indentSize, #spaceAfterSep, #spaceInParens, #separator').forEach((el) =>
    el.addEventListener("change", onSettingsChanged)
  );
  $("density").addEventListener("input", onSettingsChanged);

  $("load").addEventListener("click", () => void loadFromCell());
  $("beautify").addEventListener("click", doBeautify);
  $("minify").addEventListener("click", doMinify);
  $("apply").addEventListener("click", () => void applyToCell());

  initShortcuts();

  await detectSeparator();
  renderSettingsInfo();
});
