/* global Office */
/*
 * Keyboard shortcuts for the Beautify / Minify actions. Defaults live in
 * shortcuts.json (referenced from manifest.xml). User changes go through
 * Office.actions.replaceShortcuts, which Office stores in the user's roaming settings.
 */

export type ShortcutAction = "Beautify" | "Minify";
export type ShortcutMap = Record<ShortcutAction, string | null>;

/** Must match shortcuts.json. */
export const DEFAULT_SHORTCUTS: Record<ShortcutAction, string> = {
  Beautify: "Ctrl+Alt+Shift+B",
  Minify: "Ctrl+Alt+Shift+M",
};

export function shortcutsSupported(): boolean {
  try {
    return Office.context.requirements.isSetSupported("KeyboardShortcuts", "1.1");
  } catch {
    return false;
  }
}

/** Current combinations; null means the shortcut was overridden in a conflict prompt. */
export async function getShortcuts(): Promise<ShortcutMap> {
  const all = (await Office.actions.getShortcuts()) as Record<string, string | null>;
  return { Beautify: all.Beautify ?? null, Minify: all.Minify ?? null };
}

/** Pass null for an action to restore its default from shortcuts.json. */
export async function setShortcuts(map: Partial<ShortcutMap>): Promise<void> {
  await Office.actions.replaceShortcuts(map as { [actionId: string]: string });
}

/** Combinations from the list that Excel or another add-in already uses. */
export async function shortcutsInUse(combos: string[]): Promise<string[]> {
  const report = await Office.actions.areShortcutsInUse(combos);
  return report.filter((r) => r.inUse).map((r) => r.shortcut);
}

/**
 * Turn a keydown event into an Office key combination string ("Ctrl+Alt+Shift+B").
 * Returns "" while only modifiers are held, or an error message prefixed with "!"
 * when the combination is not allowed.
 */
export function comboFromEvent(e: KeyboardEvent): string {
  let key = "";
  if (/^Key[A-Z]$/.test(e.code)) key = e.code.slice(3);
  else if (/^Digit[0-9]$/.test(e.code)) key = e.code.slice(5);
  else if (e.code === "Minus") key = "-";
  else if (e.code === "Equal") key = "+";
  if (["Control", "Alt", "Shift", "Meta", "AltGraph"].includes(e.key)) return "";
  if (!key) return "!Use a letter, a digit, - or +.";
  if (!e.ctrlKey && !e.altKey) return "!Include Ctrl and/or Alt.";
  return (e.ctrlKey ? "Ctrl+" : "") + (e.altKey ? "Alt+" : "") + (e.shiftKey ? "Shift+" : "") + key;
}
