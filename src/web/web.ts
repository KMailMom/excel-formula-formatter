/*
 * Browser version of the formatter (online.html). Uses the same formatter code as the
 * add-in. Everything runs locally in the page; nothing is sent anywhere.
 */
import { formatFormula, flattenFormulaSafe, toDisplayFormula, fromDisplayFormula, FormatResult } from "../lib/formatter";
import { DENSITY_LEVELS } from "../lib/settings";

type Mode = "format" | "flatten";
type Sep = "," | ";";

interface WebState {
  mode: Mode;
  density: number;
  indent: "2" | "4" | "tab";
  sep: Sep;
  afterSeparator: boolean;
  insideParens: boolean;
}

const KEY = "formulaFormatter.web";
const SAMPLE =
  '=IF([@Status]="Closed",0,LET(days,TODAY()-[@[Start Date]],rate,XLOOKUP([@Customer],Customers[Name],Customers[Rate],0),IF(days>30,days*rate*1.1,days*rate)))';

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;

/** Argument separator for the visitor's locale: a comma decimal separator implies ";". */
function localeSeparator(): Sep {
  try {
    const decimal = new Intl.NumberFormat().formatToParts(1.5).find((p) => p.type === "decimal")?.value;
    return decimal === "," ? ";" : ",";
  } catch {
    return ",";
  }
}

function defaults(): WebState {
  return { mode: "format", density: 2, indent: "4", sep: localeSeparator(), afterSeparator: false, insideParens: false };
}

function load(): WebState {
  const d = defaults();
  try {
    const s = JSON.parse(localStorage.getItem(KEY) || "null") as Partial<WebState> | null;
    if (!s) return d;
    return {
      mode: s.mode === "flatten" ? "flatten" : "format",
      density: typeof s.density === "number" && DENSITY_LEVELS[s.density] ? s.density : d.density,
      indent: s.indent === "2" || s.indent === "tab" ? s.indent : "4",
      sep: s.sep === ";" || s.sep === "," ? s.sep : d.sep,
      afterSeparator: s.afterSeparator === true,
      insideParens: s.insideParens === true,
    };
  } catch {
    return d;
  }
}

function save(s: WebState): void {
  try {
    localStorage.setItem(KEY, JSON.stringify(s));
  } catch {
    /* storage unavailable: settings just aren't remembered */
  }
}

let state: WebState;

function readUi(): WebState {
  return {
    mode: ($<HTMLInputElement>("modeFlatten").checked ? "flatten" : "format") as Mode,
    density: parseInt($<HTMLSelectElement>("density").value, 10),
    indent: $<HTMLSelectElement>("indent").value as WebState["indent"],
    sep: ($<HTMLSelectElement>("sep").value === ";" ? ";" : ",") as Sep,
    afterSeparator: $<HTMLInputElement>("afterSep").checked,
    insideParens: $<HTMLInputElement>("inParens").checked,
  };
}

function applyUi(s: WebState): void {
  $<HTMLInputElement>(s.mode === "flatten" ? "modeFlatten" : "modeFormat").checked = true;
  $<HTMLSelectElement>("density").value = String(s.density);
  $<HTMLSelectElement>("indent").value = s.indent;
  $<HTMLSelectElement>("sep").value = s.sep;
  $<HTMLInputElement>("afterSep").checked = s.afterSeparator;
  $<HTMLInputElement>("inParens").checked = s.insideParens;
  // Density and indent only matter for Format.
  $("formatOnly").toggleAttribute("hidden", s.mode !== "format");
  $("densityDesc").textContent = DENSITY_LEVELS[s.density].description;
}

function status(text: string, kind: "" | "ok" | "err" = ""): void {
  const el = $("status");
  el.textContent = text;
  el.className = "status " + kind;
}

function run(): void {
  const output = $<HTMLTextAreaElement>("output");
  let raw = $<HTMLTextAreaElement>("input").value.trim();
  if (!raw) {
    output.value = "";
    status("Paste a formula above.");
    return;
  }
  const addedEquals = raw.charAt(0) !== "=";
  if (addedEquals) raw = "=" + raw;

  const canonical = fromDisplayFormula(raw, state.sep);
  const spacing = { afterSeparator: state.afterSeparator, insideParens: state.insideParens };
  let result: FormatResult;
  if (state.mode === "format") {
    const indent = state.indent === "tab" ? { useSpaces: false, size: 4 } : { useSpaces: true, size: parseInt(state.indent, 10) };
    result = formatFormula(canonical, { indent, spacing, inlineWidth: DENSITY_LEVELS[state.density].inlineWidth });
  } else {
    result = flattenFormulaSafe(canonical, spacing);
  }

  if (!result.ok) {
    output.value = "";
    const reason = result.reason ?? "could not read the formula";
    status(`${reason.charAt(0).toUpperCase()}${reason.slice(1)}. Check the separator setting and the brackets.`, "err");
    return;
  }
  output.value = toDisplayFormula(result.value, state.sep);
  const lines = output.value.split("\n").length;
  const what = state.mode === "format" ? `Formatted over ${lines} line${lines === 1 ? "" : "s"}` : "Flattened to one line";
  status(what + (addedEquals ? ' (added the leading "=")' : "") + ".", "ok");
}

function onChange(): void {
  const before = state.sep;
  state = readUi();
  // Keep the input readable when the separator changes.
  const input = $<HTMLTextAreaElement>("input");
  if (before !== state.sep && input.value.trim()) {
    input.value = toDisplayFormula(fromDisplayFormula(input.value, before), state.sep);
  }
  save(state);
  applyUi(state);
  run();
}

async function copyOutput(): Promise<void> {
  const out = $<HTMLTextAreaElement>("output");
  if (!out.value) return;
  try {
    await navigator.clipboard.writeText(out.value);
  } catch {
    out.select();
    document.execCommand("copy");
  }
  const btn = $<HTMLButtonElement>("copy");
  btn.textContent = "Copied";
  setTimeout(() => (btn.textContent = "Copy"), 1500);
}

function init(): void {
  const densitySelect = $<HTMLSelectElement>("density");
  DENSITY_LEVELS.forEach((l, i) => densitySelect.add(new Option(l.label, String(i))));
  state = load();
  applyUi(state);

  for (const id of ["modeFormat", "modeFlatten", "density", "indent", "sep", "afterSep", "inParens"]) {
    $(id).addEventListener("change", onChange);
  }
  $("input").addEventListener("input", run);
  $("copy").addEventListener("click", copyOutput);
  $("sample").addEventListener("click", () => {
    $<HTMLTextAreaElement>("input").value = toDisplayFormula(SAMPLE, state.sep);
    run();
  });
  $("clear").addEventListener("click", () => {
    $<HTMLTextAreaElement>("input").value = "";
    run();
    $("input").focus();
  });
  run();
}

document.addEventListener("DOMContentLoaded", init);
