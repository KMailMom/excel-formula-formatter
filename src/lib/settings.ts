/*
 * User preferences. Stored in localStorage, which the shared runtime (task pane and
 * ribbon/keyboard commands run in the same page) reads on every command, so a change
 * in the Options pane applies to the next Beautify/Minify immediately.
 */
import { IndentOptions, SpacingOptions, FormatOptions, DEFAULT_INDENT, DEFAULT_SPACING } from "./formatter";

/** How the argument separator is shown in the Options pane editor. */
export type SeparatorSetting = "auto" | ";" | ",";

export interface Settings {
  indent: IndentOptions;
  spacing: SpacingOptions;
  /** index into DENSITY_LEVELS */
  density: number;
  separator: SeparatorSetting;
}

export interface DensityLevel {
  label: string;
  description: string;
  /** see FormatOptions.inlineWidth */
  inlineWidth: number;
}

export const DENSITY_LEVELS: DensityLevel[] = [
  { label: "Expanded", description: "Every function gets its own lines. Maximum structure, long output.", inlineWidth: 0 },
  { label: "Spacious", description: "Very short calls such as N([@Tot]) stay on one line (up to 25 characters).", inlineWidth: 25 },
  { label: "Balanced", description: "Calls up to 50 characters stay on one line. Short formulas are not split.", inlineWidth: 50 },
  { label: "Compact", description: "Calls up to 80 characters stay on one line. Only long formulas are split.", inlineWidth: 80 },
  { label: "Dense", description: "Calls up to 120 characters stay on one line. Splits only very long formulas.", inlineWidth: 120 },
];

export const DEFAULT_SETTINGS: Settings = {
  indent: DEFAULT_INDENT,
  spacing: DEFAULT_SPACING,
  density: 2,
  separator: "auto",
};

const KEY = "excelFormulaFormatter.settings";
/** Pre-1.2 key that only held the indent preference. */
const LEGACY_INDENT_KEY = "excelFormulaFormatter.indent";

function read(key: string): unknown {
  try {
    const raw = typeof localStorage !== "undefined" ? localStorage.getItem(key) : null;
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function normalizeIndent(v: Partial<IndentOptions> | null | undefined): IndentOptions {
  const size = typeof v?.size === "number" && v.size >= 1 && v.size <= 8 ? Math.round(v.size) : DEFAULT_INDENT.size;
  return { useSpaces: v?.useSpaces !== false, size };
}

export function loadSettings(): Settings {
  const s = read(KEY) as Partial<Settings> | null;
  const legacy = read(LEGACY_INDENT_KEY) as Partial<IndentOptions> | null;
  const density = typeof s?.density === "number" && DENSITY_LEVELS[s.density] ? s.density : DEFAULT_SETTINGS.density;
  const separator = s?.separator === ";" || s?.separator === "," ? s.separator : "auto";
  return {
    indent: normalizeIndent(s?.indent ?? legacy),
    spacing: {
      afterSeparator: s?.spacing?.afterSeparator === true,
      insideParens: s?.spacing?.insideParens === true,
    },
    density,
    separator,
  };
}

export function saveSettings(settings: Settings): void {
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(KEY, JSON.stringify(settings));
    }
  } catch {
    /* non-fatal */
  }
}

export function toFormatOptions(settings: Settings): FormatOptions {
  return {
    indent: settings.indent,
    spacing: settings.spacing,
    inlineWidth: DENSITY_LEVELS[settings.density].inlineWidth,
  };
}
