/*
 * Persist the user's indent preference. Uses localStorage, which is shared across
 * the add-in's same-origin runtimes (the commands host page and the task pane), so a
 * preference set in the Options pane is picked up by the ribbon buttons.
 */
import { IndentOptions, DEFAULT_INDENT } from "./formatter";

const KEY = "excelFormulaFormatter.indent";

export function loadIndent(): IndentOptions {
  try {
    const raw = typeof localStorage !== "undefined" ? localStorage.getItem(KEY) : null;
    if (raw) {
      const parsed = JSON.parse(raw) as Partial<IndentOptions>;
      return {
        useSpaces: parsed.useSpaces !== false,
        size: typeof parsed.size === "number" ? parsed.size : DEFAULT_INDENT.size,
      };
    }
  } catch {
    /* fall through to default */
  }
  return DEFAULT_INDENT;
}

export function saveIndent(indent: IndentOptions): void {
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(KEY, JSON.stringify(indent));
    }
  } catch {
    /* non-fatal */
  }
}
