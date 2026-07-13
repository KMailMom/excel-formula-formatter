/* global Office, Excel, document */
import { beautifyFormula, minifyFormula, IndentOptions, DEFAULT_INDENT } from "../lib/formatter";
import { loadIndent, saveIndent } from "../lib/settings";

let previewEl: HTMLTextAreaElement;
let statusEl: HTMLElement;

function setStatus(message: string, isError = false): void {
  statusEl.textContent = message;
  statusEl.classList.toggle("err", isError);
}

function readIndentFromUi(): IndentOptions {
  const kind = (document.querySelector('input[name="indentKind"]:checked') as HTMLInputElement)?.value;
  const size = parseInt((document.getElementById("indentSize") as HTMLInputElement).value, 10);
  return { useSpaces: kind !== "tab", size: Number.isFinite(size) ? size : DEFAULT_INDENT.size };
}

function applyIndentToUi(indent: IndentOptions): void {
  const kind = indent.useSpaces ? "spaces" : "tab";
  const radio = document.querySelector(`input[name="indentKind"][value="${kind}"]`) as HTMLInputElement | null;
  if (radio) radio.checked = true;
  (document.getElementById("indentSize") as HTMLInputElement).value = String(indent.size);
}

function persistIndent(): void {
  saveIndent(readIndentFromUi());
}

function loadPersistedIndent(): void {
  applyIndentToUi(loadIndent());
}

async function loadFromCell(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.load(["formulas", "address"]);
      await context.sync();
      const value = cell.formulas[0][0];
      previewEl.value = typeof value === "string" ? value : String(value ?? "");
      setStatus(previewEl.value ? `Loaded ${cell.address}.` : `${cell.address} has no formula.`);
    });
  } catch (err) {
    setStatus((err as Error).message, true);
  }
}

function doBeautify(): void {
  const src = previewEl.value.trim();
  if (!src) return setStatus("Nothing to beautify.");
  if (src.charAt(0) !== "=") return setStatus('A formula must start with "=".', true);
  const result = beautifyFormula(src, readIndentFromUi());
  if (result.ok) {
    previewEl.value = result.value;
    setStatus("Beautified. Review, then write back to the cell.");
  } else {
    setStatus("Left unchanged: " + (result.reason ?? "unsupported syntax") + ".", true);
  }
}

function doMinify(): void {
  const src = previewEl.value.trim();
  if (!src) return setStatus("Nothing to minify.");
  if (src.charAt(0) !== "=") return setStatus('A formula must start with "=".', true);
  previewEl.value = minifyFormula(src);
  setStatus("Minified. Review, then write back to the cell.");
}

async function applyToCell(): Promise<void> {
  const value = previewEl.value.trim();
  if (!value) return setStatus("Nothing to write.");
  if (value.charAt(0) !== "=") return setStatus('A formula must start with "=".', true);
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

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) return;

  previewEl = document.getElementById("preview") as HTMLTextAreaElement;
  statusEl = document.getElementById("status") as HTMLElement;

  loadPersistedIndent();

  document.getElementById("load")!.addEventListener("click", () => void loadFromCell());
  document.getElementById("beautify")!.addEventListener("click", () => {
    persistIndent();
    doBeautify();
  });
  document.getElementById("minify")!.addEventListener("click", doMinify);
  document.getElementById("apply")!.addEventListener("click", () => void applyToCell());
  document.querySelectorAll('input[name="indentKind"], #indentSize').forEach((el) =>
    el.addEventListener("change", () => persistIndent())
  );

  // Prime the preview with the current selection for convenience.
  void loadFromCell();
});
