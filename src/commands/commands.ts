/* global Office, Excel */
/*
 * Command handlers for the ribbon buttons and keyboard shortcuts. This module runs in
 * the shared runtime (it is bundled into taskpane.html) and registers its actions with
 * Office.actions.associate when imported.
 */
import { beautifyFormula, minifyFormulaSafe, isFormula, FormatOptions, SpacingOptions } from "../lib/formatter";
import { loadSettings, toFormatOptions } from "../lib/settings";
import { showDialog, DialogKind } from "../lib/dialog";

// Safety cap for a single selection (e.g. guards against a whole-column select).
const MAX_SELECTION_CELLS = 50000;
// Per-sheet cap for the whole-workbook commands. Sheets with a larger used range
// are skipped and reported rather than silently ignored.
const MAX_SHEET_CELLS = 200000;

type CellValue = string | number | boolean;

interface CellOutcome {
  value: CellValue;
  applied: boolean;
  reason?: string;
}
type CellTransform = (formula: string) => CellOutcome;

function beautifyTransform(options: FormatOptions): CellTransform {
  return (formula) => {
    const r = beautifyFormula(formula, options);
    return { value: r.value, applied: r.ok, reason: r.reason };
  };
}

function minifyTransform(spacing: SpacingOptions): CellTransform {
  return (formula) => {
    const r = minifyFormulaSafe(formula, spacing);
    return { value: r.value, applied: r.ok, reason: r.reason };
  };
}

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

function notify(title: string, message: string, kind: DialogKind = "info"): Promise<void> {
  return showDialog({ title, message, kind }).then(() => undefined);
}

function confirmDialog(title: string, message: string, okLabel: string): Promise<boolean> {
  return showDialog({ title, message, kind: "confirm", okLabel }).then((r) => r === "ok");
}

/** Friendly text for errors; cell-edit mode gets an explanation instead of Excel's raw message. */
function errorMessage(err: unknown): string {
  const code = (err as { code?: string })?.code;
  if (code === "InvalidOperationInCellEditMode") {
    return (
      "Excel is still editing a cell, and add-ins cannot read or change a cell while you are typing in it.\n\n" +
      "Press Ctrl+Enter to confirm the edit (the cell stays selected), then run the command again."
    );
  }
  return (err as Error)?.message ?? String(err);
}

// ---- command handlers ------------------------------------------------------

// Ribbon buttons pass an event that must be completed; keyboard shortcuts pass none.
type CommandEvent = Office.AddinCommands.Event | undefined;

function complete(event: CommandEvent): void {
  try {
    event?.completed();
  } catch {
    /* not a ribbon invocation */
  }
}

async function runBeautify(event?: CommandEvent) {
  try {
    const s = await transformSelection(beautifyTransform(toFormatOptions(loadSettings())));
    if (s.formulaCells === 0) {
      await notify("Beautify", "No formula found in the selected cell(s).");
    } else if (s.changed === 0 && s.firstSkipReason) {
      await notify("Beautify", "The formula was left unchanged: " + s.firstSkipReason + ".", "warning");
    }
  } catch (err) {
    await notify("Beautify failed", errorMessage(err), "error");
  } finally {
    complete(event);
  }
}

async function runMinify(event?: CommandEvent) {
  try {
    const s = await transformSelection(minifyTransform(loadSettings().spacing));
    if (s.formulaCells === 0) {
      await notify("Minify", "No formula found in the selected cell(s).");
    } else if (s.changed === 0 && s.firstSkipReason) {
      await notify("Minify", "The formula was left unchanged: " + s.firstSkipReason + ".", "warning");
    }
  } catch (err) {
    await notify("Minify failed", errorMessage(err), "error");
  } finally {
    complete(event);
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

async function runBeautifyAll(event?: CommandEvent) {
  try {
    const go = await confirmDialog(
      "Beautify all formulas?",
      "Every formula on every sheet of this workbook will be reformatted with your current settings. You can undo this with Ctrl+Z.",
      "Beautify all"
    );
    if (!go) return;
    const s = await transformWorkbook(beautifyTransform(toFormatOptions(loadSettings())));
    await notify("Beautify all", summarizeWorkbook("Beautified", s), s.changed > 0 ? "success" : "info");
  } catch (err) {
    await notify("Beautify all failed", errorMessage(err), "error");
  } finally {
    complete(event);
  }
}

async function runMinifyAll(event?: CommandEvent) {
  try {
    const go = await confirmDialog(
      "Minify all formulas?",
      "Every formula on every sheet of this workbook will be collapsed to a single line. You can undo this with Ctrl+Z.",
      "Minify all"
    );
    if (!go) return;
    const s = await transformWorkbook(minifyTransform(loadSettings().spacing));
    await notify("Minify all", summarizeWorkbook("Minified", s), s.changed > 0 ? "success" : "info");
  } catch (err) {
    await notify("Minify all failed", errorMessage(err), "error");
  } finally {
    complete(event);
  }
}

// FunctionName values in manifest.xml and action ids in shortcuts.json must match
// these ids exactly.
Office.actions.associate("Beautify", runBeautify);
Office.actions.associate("Minify", runMinify);
Office.actions.associate("BeautifyAll", runBeautifyAll);
Office.actions.associate("MinifyAll", runMinifyAll);
