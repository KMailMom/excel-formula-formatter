/*
 * Core beautify / minify logic. Host-agnostic and pure: depends only on the
 * vendored global `window.excelFormulaUtilities` (loaded via a <script> tag in
 * commands.html and taskpane.html). This module never imports the vendored file,
 * because webpack would wrap its top-level global assignment in a module scope
 * and `window.excelFormulaUtilities` would never be set.
 *
 * Design notes (verified against the real library in Node before shipping):
 *  - Beautify reuses the library's formatFormula, but the library has a bug where
 *    it does not re-double `"` inside string literals. We fix that with a
 *    customTokenRender hook. We then run a safety check: a beautified formula must
 *    minify back to the same canonical form as the original, otherwise we leave the
 *    formula unchanged. This makes beautify incapable of silently corrupting a
 *    formula (e.g. the library mangles array constants like {1,2;3,4}).
 *  - Minify is a self-contained, quote-aware whitespace stripper. It never rewrites
 *    token values, so arrays, sheet names, and string contents survive intact.
 *
 * Known limitation: an implicit-intersection space between two ranges
 * (e.g. =SUM(A1:A10 A1:C1)) is removed by minify. This operator is very rarely used.
 */

export interface IndentOptions {
  /** true = indent with spaces, false = indent with a tab character */
  useSpaces: boolean;
  /** number of spaces per indent level when useSpaces is true */
  size: number;
}

export const DEFAULT_INDENT: IndentOptions = { useSpaces: true, size: 4 };

export interface BeautifyResult {
  /** true if the formula was reformatted; false if it was left unchanged */
  ok: boolean;
  /** the beautified formula when ok, otherwise the original formula */
  value: string;
  /** why the formula was left unchanged (only set when ok is false) */
  reason?: string;
}

interface FToken {
  value: string;
  type: string;
  subtype: string;
}

/** True when the value is a formula we should try to transform. */
export function isFormula(value: unknown): value is string {
  return typeof value === "string" && value.charAt(0) === "=";
}

function stripLeadingEquals(formula: string): string {
  return formula.charAt(0) === "=" ? formula.slice(1) : formula;
}

/**
 * customTokenRender hook: re-double `"` inside string literals so the library's
 * `"{{token}}"` template stays valid (the tokenizer un-doubles them on the way in).
 */
function quoteFix(tokenString: string, token: FToken): { tokenString: string; useTemplate: boolean } {
  if (token.type === "operand" && token.subtype === "text") {
    return { tokenString: String(tokenString).replace(/"/g, '""'), useTemplate: true };
  }
  return { tokenString: tokenString, useTemplate: true };
}

/**
 * Remove insignificant whitespace (spaces, tabs, newlines) that sits OUTSIDE of
 * double-quoted string literals and single-quoted sheet names. Doubled quotes
 * (`""` and `''`) are treated as escaped and stay inside their literal.
 */
export function minifyFormula(formula: string): string {
  const s = String(formula);
  let out = "";
  let i = 0;
  let inDouble = false;
  let inSingle = false;
  while (i < s.length) {
    const c = s[i];
    if (inDouble) {
      out += c;
      if (c === '"') {
        if (s[i + 1] === '"') { out += '"'; i += 2; continue; }
        inDouble = false;
      }
      i++;
      continue;
    }
    if (inSingle) {
      out += c;
      if (c === "'") {
        if (s[i + 1] === "'") { out += "'"; i += 2; continue; }
        inSingle = false;
      }
      i++;
      continue;
    }
    if (c === '"') { inDouble = true; out += c; i++; continue; }
    if (c === "'") { inSingle = true; out += c; i++; continue; }
    if (/\s/.test(c)) { i++; continue; }
    out += c;
    i++;
  }
  return out;
}

/**
 * Expand a formula onto indented multiple lines. Returns a BeautifyResult; when
 * `ok` is false the original formula is returned unchanged (safe fallback).
 */
export function beautifyFormula(formula: string, indent: IndentOptions = DEFAULT_INDENT): BeautifyResult {
  const efu = window.excelFormulaUtilities;
  if (!efu || typeof efu.formatFormula !== "function") {
    return { ok: false, value: formula, reason: "formatter library not loaded" };
  }

  const body = stripLeadingEquals(formula);
  const indentUnit = indent.useSpaces ? " ".repeat(Math.max(1, indent.size)) : "\t";

  let out: string;
  try {
    out = efu.formatFormula(body, {
      tmplIndentTab: indentUnit,
      newLine: "\n",
      customTokenRender: quoteFix,
    });
  } catch (e) {
    return { ok: false, value: formula, reason: "formatter error: " + (e as Error).message };
  }

  // The library emits no leading "=" (prefix defaults to ""); add exactly one.
  const beautified = "=" + out.replace(/^\s*=/, "");

  // Safety net: the beautified formula must be semantically identical to the
  // original (same content once insignificant whitespace is stripped).
  if (minifyFormula(beautified) !== minifyFormula(formula)) {
    return { ok: false, value: formula, reason: "formula uses syntax the formatter cannot safely expand" };
  }

  return { ok: true, value: beautified };
}
