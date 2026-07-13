/* global Office, Excel, console */
import { beautifyFormula, minifyFormula, isFormula, IndentOptions } from "../lib/formatter";
import { loadIndent } from "../lib/settings";

// Safety cap for a single selection (e.g. guards against a whole-column select).
const MAX_SELECTION_CELLS = 50000;
// Per-sheet cap for the whole-workbook commands. Sheets with a larger used range
// are skipped and reported rather than silently ignored.
const MAX_SHEET_CELLS = 200000;

Office.onReady(() => {
  // Handlers are wired up via Office.actions.associate below.
});

type CellValue = string | number | boolean;

interface CellOutcome {
  value: CellValue;
  applied: boolean;
  reason?: string;
}
type CellTransform = (formula: string) => CellOutcome;

function beautifyTransform(indent: IndentOptions): CellTransform {
  return (formula) => {
    const r = beautifyFormula(formula, indent);
    return { value: r.value, applied: r.ok, reason: r.reason };
  };
}

const minifyTransform: CellTransform = (formula) => {
  return { value: minifyFormula(formula), applied: true };
};

// ---- shared grid logic -----------------------------------------------------

interface GridResult {
  next: CellValue[][];
  formulaCells: number;
  changed: number;
  firstSkipReason?: string;
  anyChange: boolean;
}

function applyToGrid(grid: CellValue[][], transform: CellTransform): GridResult {
  const result: GridResult = { next: grid, formulaCells: 0, changed: 0, anyChange: false };
  result.next = grid.map((row) =>
    row.map((cell) => {
      if (!isFormula(cell)) return cell;
      result.formulaCells++;
      const outcome = transform(cell);
      if (outcome.applied && outcome.value !== cell) {
        result.changed++;
        result.anyChange = true;
        return outcome.value;
      }
      if (!outcome.applied && !result.firstSkipReason && outcome.reason) {
        result.firstSkipReason = outcome.reason;
      }
      return cell;
    })
  );
  return result;
}

// ---- selection scope -------------------------------------------------------

interface RunSummary {
  formulaCells: number;
  changed: number;
  firstSkipReason?: string;
}

async function transformSelection(transform: CellTransform): Promise<RunSummary> {
  const summary: RunSummary = { formulaCells: 0, changed: 0 };
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["cellCount"]);
    await context.sync();

    if (range.cellCount > MAX_SELECTION_CELLS) {
      throw new Error(`The selection (${range.cellCount} cells) is too large. Select up to ${MAX_SELECTION_CELLS} cells.`);
    }

    // Always .formulas (canonical en-US commas), never .formulasLocal.
    range.load("formulas");
    await context.sync();

    const res = applyToGrid(range.formulas as CellValue[][], transform);
    summary.formulaCells = res.formulaCells;
    summary.changed = res.changed;
    summary.firstSkipReason = res.firstSkipReason;
    if (res.anyChange) {
      range.formulas = res.next;
      await context.sync();
    }
  });
  return summary;
}

// ---- workbook scope --------------------------------------------------------

interface WorkbookSummary {
  sheetsChanged: number;
  formulaCells: number;
  changed: number;
  skippedSheets: string[];
  firstSkipReason?: string;
}

async function transformWorkbook(transform: CellTransform): Promise<WorkbookSummary> {
  const summary: WorkbookSummary = { sheetsChanged: 0, formulaCells: 0, changed: 0, skippedSheets: [] };
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const used = sheets.items.map((ws) => {
      const range = ws.getUsedRangeOrNullObject(true);
      range.load(["isNullObject", "cellCount"]);
      return { name: ws.name, range };
    });
    await context.sync();

    const active: typeof used = [];
    for (const u of used) {
      if (u.range.isNullObject) continue;
      if (u.range.cellCount > MAX_SHEET_CELLS) {
        summary.skippedSheets.push(u.name);
        continue;
      }
      u.range.load("formulas");
      active.push(u);
    }
    if (active.length > 0) await context.sync();

    for (const u of active) {
      const res = applyToGrid(u.range.formulas as CellValue[][], transform);
      summary.formulaCells += res.formulaCells;
      summary.changed += res.changed;
      if (!summary.firstSkipReason && res.firstSkipReason) summary.firstSkipReason = res.firstSkipReason;
      if (res.anyChange) {
        u.range.formulas = res.next;
        summary.sheetsChanged++;
      }
    }
    if (summary.sheetsChanged > 0) await context.sync();
  });
  return summary;
}

// ---- dialogs (feedback + confirmation) -------------------------------------

/**
 * Open the small dialog page. When confirm is true it shows Cancel/Continue and
 * resolves with "ok" or "cancel"; otherwise it shows OK. Resolves "" on failure
 * or if the user dismisses the dialog.
 */
function showDialog(title: string, message: string, confirm = false): Promise<string> {
  return new Promise((resolve) => {
    let url: string;
    try {
      url =
        new URL("dialog.html", location.href).href +
        `?title=${encodeURIComponent(title)}&msg=${encodeURIComponent(message)}${confirm ? "&confirm=1" : ""}`;
    } catch {
      resolve("");
      return;
    }
    Office.context.ui.displayDialogAsync(url, { height: 28, width: 34, displayInIframe: true }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn("dialog failed:", result.error);
        resolve("");
        return;
      }
      const dialog = result.value;
      let settled = false;
      const finish = (value: string) => {
        if (settled) return;
        settled = true;
        try {
          dialog.close();
        } catch {
          /* already closed */
        }
        resolve(value);
      };
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const message = (arg as { message?: string }).message ?? "";
        finish(message);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => finish(""));
    });
  });
}

function notify(title: string, message: string): Promise<void> {
  return showDialog(title, message, false).then(() => undefined);
}

function confirmDialog(title: string, message: string): Promise<boolean> {
  return showDialog(title, message, true).then((r) => r === "ok");
}

// ---- command handlers ------------------------------------------------------

async function runBeautify(event: Office.AddinCommands.Event) {
  try {
    const s = await transformSelection(beautifyTransform(loadIndent()));
    if (s.formulaCells === 0) {
      await notify("Beautify Formula", "No formula found in the selected cell(s).");
    } else if (s.changed === 0) {
      await notify("Beautify Formula", s.firstSkipReason ? "Left unchanged: " + s.firstSkipReason + "." : "The formula is already beautified.");
    }
  } catch (err) {
    await notify("Beautify Formula error", (err as Error)?.message ?? String(err));
  } finally {
    event.completed();
  }
}

async function runMinify(event: Office.AddinCommands.Event) {
  try {
    const s = await transformSelection(minifyTransform);
    if (s.formulaCells === 0) {
      await notify("Minify Formula", "No formula found in the selected cell(s).");
    }
  } catch (err) {
    await notify("Minify Formula error", (err as Error)?.message ?? String(err));
  } finally {
    event.completed();
  }
}

function summarizeWorkbook(verb: string, s: WorkbookSummary): string {
  const skipNote = s.skippedSheets.length > 0 ? ` Skipped ${s.skippedSheets.length} large sheet(s): ${s.skippedSheets.join(", ")}.` : "";
  if (s.formulaCells === 0) return "No formulas found in the workbook." + skipNote;
  if (s.changed === 0) {
    const base = s.firstSkipReason ? "No changes made (" + s.firstSkipReason + ")." : "All formulas were already in that form.";
    return base + skipNote;
  }
  const cells = `${s.changed} formula${s.changed === 1 ? "" : "s"}`;
  const sheets = `${s.sheetsChanged} sheet${s.sheetsChanged === 1 ? "" : "s"}`;
  return `${verb} ${cells} across ${sheets}.` + skipNote;
}

async function runBeautifyAll(event: Office.AddinCommands.Event) {
  try {
    const go = await confirmDialog("Beautify all formulas", "This reformats every formula in all sheets of this workbook. You can undo with Ctrl+Z. Continue?");
    if (!go) return;
    const s = await transformWorkbook(beautifyTransform(loadIndent()));
    await notify("Beautify all formulas", summarizeWorkbook("Beautified", s));
  } catch (err) {
    await notify("Beautify all error", (err as Error)?.message ?? String(err));
  } finally {
    event.completed();
  }
}

async function runMinifyAll(event: Office.AddinCommands.Event) {
  try {
    const go = await confirmDialog("Minify all formulas", "This collapses every formula in all sheets of this workbook to a single line. You can undo with Ctrl+Z. Continue?");
    if (!go) return;
    const s = await transformWorkbook(minifyTransform);
    await notify("Minify all formulas", summarizeWorkbook("Minified", s));
  } catch (err) {
    await notify("Minify all error", (err as Error)?.message ?? String(err));
  } finally {
    event.completed();
  }
}

// FunctionName values in manifest.xml must match these action ids exactly.
Office.actions.associate("Beautify", runBeautify);
Office.actions.associate("Minify", runMinify);
Office.actions.associate("BeautifyAll", runBeautifyAll);
Office.actions.associate("MinifyAll", runMinifyAll);
