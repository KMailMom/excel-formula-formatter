/*
 * Core beautify / minify logic. Host-agnostic and pure (no Office.js, no DOM), so it
 * can be unit-tested in Node.
 *
 * Pipeline: tokenize -> parse into a small tree -> render.
 *  - The tokenizer is quote/bracket aware: string literals, quoted sheet names,
 *    structured references ([@[Termijn van]]) and array constants ({1,2;3,4}) are
 *    kept verbatim, including any spaces inside them.
 *  - Beautify renders the tree with a width budget: a function call whose one-line
 *    form fits within `layout.inlineWidth` stays on one line, otherwise its arguments
 *    go on their own indented lines. inlineWidth 0 expands every call.
 *  - Minify renders everything on one line.
 *  - Safety net: any output must reduce to the same canonical form as the input once
 *    insignificant whitespace is stripped, otherwise the formula is left unchanged.
 *
 * Formulas are handled in the canonical en-US form that Excel's `range.formulas`
 * uses (comma separators, dot decimals). toDisplayFormula / fromDisplayFormula convert
 * to and from a semicolon-separated form for showing formulas to the user.
 *
 * Whitespace between two operands (the intersection operator, =SUM(A1:A10 B1:B10))
 * is significant and is preserved as a single space.
 */

export interface IndentOptions {
  /** true = indent with spaces, false = indent with a tab character */
  useSpaces: boolean;
  /** number of spaces per indent level when useSpaces is true */
  size: number;
}

export interface SpacingOptions {
  /** put a space after every argument separator on single-line parts */
  afterSeparator: boolean;
  /** put a space after "(" and before ")" on single-line parts */
  insideParens: boolean;
}

export interface FormatOptions {
  indent: IndentOptions;
  spacing: SpacingOptions;
  /**
   * Beautify only: a function call or bracketed group whose single-line form is at
   * most this many characters stays on one line. 0 = always expand.
   */
  inlineWidth: number;
}

export const DEFAULT_INDENT: IndentOptions = { useSpaces: true, size: 4 };
export const DEFAULT_SPACING: SpacingOptions = { afterSeparator: false, insideParens: false };
export const DEFAULT_FORMAT: FormatOptions = { indent: DEFAULT_INDENT, spacing: DEFAULT_SPACING, inlineWidth: 50 };

export interface FormatResult {
  /** true if the formula could be processed; false if it was left unchanged */
  ok: boolean;
  /** the formatted formula when ok, otherwise the original formula */
  value: string;
  /** why the formula was left unchanged (only set when ok is false) */
  reason?: string;
}

/** Back-compat alias. */
export type BeautifyResult = FormatResult;

/** True when the value is a formula we should try to transform. */
export function isFormula(value: unknown): value is string {
  return typeof value === "string" && value.charAt(0) === "=";
}

// ---- tokenizer -------------------------------------------------------------

type TokType = "atom" | "op" | "open" | "close" | "sep" | "ws";

interface Tok {
  type: TokType;
  text: string;
  /** for "open": true when this "(" starts a function call (follows a name or ")") */
  call?: boolean;
}

const OP_CHARS = "+-*/^&=<>%:";
const WS = /\s/;

class FormulaSyntaxError extends Error {}

/**
 * Split a formula body (without the leading "=") into tokens. `sep` is the argument
 * separator: "," for canonical formulas, ";" for the display form.
 */
function tokenize(src: string, sep: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  const n = src.length;

  // Read a quoted run starting at i (src[i] is the quote); doubled quotes escape.
  const readQuoted = (q: string): string => {
    let j = i + 1;
    while (j < n) {
      if (src[j] === q) {
        if (src[j + 1] === q) { j += 2; continue; }
        return src.slice(i, (i = j + 1));
      }
      j++;
    }
    throw new FormulaSyntaxError("unterminated " + (q === '"' ? "text" : "sheet name"));
  };

  // Read a [...] run (nested brackets; ' escapes the next character).
  const readBracket = (): string => {
    let depth = 0;
    let j = i;
    while (j < n) {
      const c = src[j];
      if (c === "'") { j += 2; continue; }
      if (c === "[") depth++;
      else if (c === "]") {
        depth--;
        if (depth === 0) return src.slice(i, (i = j + 1));
      }
      j++;
    }
    throw new FormulaSyntaxError("unterminated [ ] reference");
  };

  // Read a {...} array constant verbatim (strings inside may contain braces).
  const readArray = (): string => {
    let j = i + 1;
    while (j < n) {
      const c = src[j];
      if (c === '"') {
        j++;
        while (j < n && !(src[j] === '"' && src[j + 1] !== '"')) j += src[j] === '"' ? 2 : 1;
        j++;
        continue;
      }
      if (c === "}") return src.slice(i, (i = j + 1));
      j++;
    }
    throw new FormulaSyntaxError("unterminated { } array");
  };

  while (i < n) {
    const c = src[i];
    if (WS.test(c)) {
      let j = i;
      while (j < n && WS.test(src[j])) j++;
      toks.push({ type: "ws", text: src.slice(i, j) });
      i = j;
      continue;
    }
    if (c === "(") {
      // A "(" glued to a name (SUM() or to ")" (LAMBDA(...)(x)) is a call.
      const last = toks[toks.length - 1];
      const call = !!last && (last.type === "atom" && !/^["{]/.test(last.text) || last.type === "close");
      toks.push({ type: "open", text: "(", call });
      i++;
      continue;
    }
    if (c === ")") { toks.push({ type: "close", text: ")" }); i++; continue; }
    if (c === sep) { toks.push({ type: "sep", text: c }); i++; continue; }
    if (c === '"') { toks.push({ type: "atom", text: readQuoted('"') }); continue; }
    if (c === "{") { toks.push({ type: "atom", text: readArray() }); continue; }
    if (OP_CHARS.includes(c)) {
      const two = src.slice(i, i + 2);
      const text = two === "<>" || two === "<=" || two === ">=" ? two : c;
      toks.push({ type: "op", text });
      i += text.length;
      continue;
    }
    // Atom: names, references, numbers, errors, booleans, structured references.
    let text = "";
    while (i < n) {
      const d = src[i];
      if (d === "'") { text += readQuoted("'"); continue; }
      if (d === "[") { text += readBracket(); continue; }
      if (WS.test(d) || d === "(" || d === ")" || d === sep || d === '"' || d === "{" || d === "}") break;
      if (OP_CHARS.includes(d)) {
        // Keep exponent signs inside numbers: 1E+5, 2.5e-3.
        if ((d === "+" || d === "-") && /^(\d+[.,]?\d*|[.,]\d+)[eE]$/.test(text)) { text += d; i++; continue; }
        break;
      }
      text += d;
      i++;
    }
    if (!text) throw new FormulaSyntaxError("unexpected character '" + src[i] + "'");
    toks.push({ type: "atom", text });
  }
  return toks;
}

// ---- parser ----------------------------------------------------------------

/** One operand with its unary prefix (-, +, @...) and postfix (%) operators. */
interface Term {
  prefix: string;
  node: Node;
  postfix: string;
}

/** An expression: terms joined by binary operators (ops.length === terms.length - 1). */
interface Expr {
  terms: Term[];
  ops: string[];
}

type Node =
  | { kind: "atom"; text: string }
  | { kind: "call"; name: string; args: Expr[] }
  | { kind: "group"; items: Expr[] };

class Parser {
  private pos = 0;
  constructor(private toks: Tok[]) {}

  private peek(): Tok | undefined {
    return this.toks[this.pos];
  }

  private skipWs(): boolean {
    let skipped = false;
    while (this.peek()?.type === "ws") { this.pos++; skipped = true; }
    return skipped;
  }

  /** Comma-separated list of expressions, up to (not including) ")" or end. */
  parseList(): Expr[] {
    const items: Expr[] = [this.parseExpr()];
    while (this.peek()?.type === "sep") {
      this.pos++;
      items.push(this.parseExpr());
    }
    return items;
  }

  parseExpr(): Expr {
    const expr: Expr = { terms: [], ops: [] };
    this.skipWs();
    const t = this.peek();
    // Empty argument, e.g. IF(A1,,1) or LAMBDA's trailing empty parameter.
    if (!t || t.type === "sep" || t.type === "close") return expr;
    expr.terms.push(this.parseTerm());
    for (;;) {
      const hadWs = this.skipWs();
      const next = this.peek();
      if (!next || next.type === "sep" || next.type === "close") break;
      if (next.type === "op" && next.text !== "%") {
        this.pos++;
        expr.ops.push(next.text);
        this.skipWs();
        expr.terms.push(this.parseTerm());
        continue;
      }
      if (hadWs && (next.type === "atom" || next.type === "open")) {
        // Two operands separated only by whitespace: the intersection operator.
        expr.ops.push(" ");
        expr.terms.push(this.parseTerm());
        continue;
      }
      throw new FormulaSyntaxError("unexpected '" + next.text + "'");
    }
    return expr;
  }

  private parseTerm(): Term {
    let prefix = "";
    for (;;) {
      this.skipWs();
      const t = this.peek();
      if (t?.type === "op" && (t.text === "-" || t.text === "+")) { prefix += t.text; this.pos++; continue; }
      break;
    }
    const node = this.parsePrimary();
    let postfix = "";
    // Postfix % may only be glued (whitespace before it would be ambiguous, keep it simple).
    while (this.peek()?.type === "op" && this.peek()!.text === "%") { postfix += "%"; this.pos++; }
    return { prefix, node, postfix };
  }

  private parsePrimary(): Node {
    const t = this.peek();
    if (!t) throw new FormulaSyntaxError("formula ends unexpectedly");
    if (t.type === "atom") {
      this.pos++;
      const after = this.peek();
      if (after?.type === "open" && after.call) {
        this.pos++;
        return this.finishCall(t.text);
      }
      return { kind: "atom", text: t.text };
    }
    if (t.type === "open") {
      this.pos++;
      this.skipWs();
      const items = this.peek()?.type === "close" ? [] : this.parseList();
      this.expectClose();
      return this.maybeChainedCall({ kind: "group", items });
    }
    throw new FormulaSyntaxError("unexpected '" + t.text + "'");
  }

  private finishCall(name: string): Node {
    this.skipWs();
    const args = this.peek()?.type === "close" ? [] : this.parseList();
    this.expectClose();
    return this.maybeChainedCall({ kind: "call", name, args });
  }

  /** LAMBDA(x, x+1)(5): a "(" glued to ")" calls the result. */
  private maybeChainedCall(node: Node): Node {
    const after = this.peek();
    if (after?.type === "open" && after.call) {
      this.pos++;
      // Render the callee verbatim as the "name" of the chained call.
      return this.finishCall(renderInline(node, MINIFY_STYLE));
    }
    return node;
  }

  private expectClose(): void {
    this.skipWs();
    const t = this.peek();
    if (t?.type !== "close") throw new FormulaSyntaxError("missing ')'");
    this.pos++;
  }

  done(): boolean {
    this.skipWs();
    return this.pos >= this.toks.length;
  }
}

function parse(body: string): Expr[] {
  const p = new Parser(tokenize(body, ","));
  const list = p.parseList();
  if (!p.done()) throw new FormulaSyntaxError("unbalanced ')'");
  return list;
}

// ---- rendering -------------------------------------------------------------

interface InlineStyle {
  sep: string;
  open: string;
  close: string;
  /** spaces around binary operators */
  opSpace: boolean;
}

const MINIFY_STYLE: InlineStyle = { sep: ",", open: "(", close: ")", opSpace: false };

function inlineStyle(spacing: SpacingOptions, opSpace: boolean): InlineStyle {
  return {
    sep: spacing.afterSeparator ? ", " : ",",
    open: spacing.insideParens ? "( " : "(",
    close: spacing.insideParens ? " )" : ")",
    opSpace,
  };
}

function joinOp(op: string, style: InlineStyle): string {
  // ":" (range) and " " (intersection) are always tight.
  if (op === ":" || op === " " || !style.opSpace) return op;
  return " " + op + " ";
}

function renderInline(node: Node, style: InlineStyle): string {
  if (node.kind === "atom") return node.text;
  const parts = node.kind === "call" ? node.args : node.items;
  const name = node.kind === "call" ? node.name : "";
  if (parts.length === 0 || (parts.length === 1 && parts[0].terms.length === 0)) return name + "()";
  return name + style.open + parts.map((e) => renderExprInline(e, style)).join(style.sep) + style.close;
}

function renderTermInline(t: Term, style: InlineStyle): string {
  return t.prefix + renderInline(t.node, style) + t.postfix;
}

function renderExprInline(e: Expr, style: InlineStyle): string {
  let s = "";
  e.terms.forEach((t, k) => {
    if (k > 0) s += joinOp(e.ops[k - 1], style);
    s += renderTermInline(t, style);
  });
  return s;
}

const MIN_FLOW_WIDTH = 60;

class Beautifier {
  private style: InlineStyle;
  private unit: string;

  constructor(private opts: FormatOptions) {
    this.style = inlineStyle(opts.spacing, true);
    this.unit = opts.indent.useSpaces ? " ".repeat(Math.max(1, opts.indent.size)) : "\t";
  }

  private fits(text: string): boolean {
    return this.opts.inlineWidth > 0 && text.length <= this.opts.inlineWidth;
  }

  private ind(level: number): string {
    return this.unit.repeat(level);
  }

  /** True when a node will be rendered on more than one line. */
  private breaks(node: Node): boolean {
    if (node.kind === "atom") return false;
    const parts = node.kind === "call" ? node.args : node.items;
    if (parts.length === 0 || (parts.length === 1 && parts[0].terms.length === 0)) return false;
    return !this.fits(renderInline(node, this.style));
  }

  node(node: Node, level: number): string {
    if (!this.breaks(node)) return renderInline(node, this.style);
    if (node.kind === "atom") return node.text;
    const parts = node.kind === "call" ? node.args : node.items;
    const name = node.kind === "call" ? node.name : "";
    const inner = parts.map((e) => this.ind(level + 1) + this.expr(e, level + 1)).join(",\n");
    return name + "(\n" + inner + "\n" + this.ind(level) + ")";
  }

  expr(e: Expr, level: number): string {
    const inline = renderExprInline(e, this.style);
    if (e.terms.length <= 1 || this.fits(inline)) {
      if (e.terms.length === 0) return "";
      if (e.terms.length === 1) {
        const t = e.terms[0];
        return t.prefix + this.node(t.node, level) + t.postfix;
      }
      return inline;
    }
    // Flow terms on lines; a term that breaks, or that would overflow the line,
    // starts on a new line after the operator.
    const indent = this.ind(level);
    let out = "";
    let lineLen = 0;
    e.terms.forEach((t, k) => {
      const multi = this.breaks(t.node);
      const text = multi ? t.prefix + this.node(t.node, level) + t.postfix : renderTermInline(t, this.style);
      if (k === 0) {
        out = text;
      } else {
        const op = e.ops[k - 1];
        if (op === ":" || op === " ") {
          out += op + text;
        } else {
          out += " " + op;
          // Expanded mode: every call/group starts a new line. Otherwise wrap long lines
          // (never narrower than MIN_FLOW_WIDTH, so a short comparison isn't split).
          const tooLong =
            this.opts.inlineWidth > 0
              ? lineLen + op.length + 2 + text.length > Math.max(this.opts.inlineWidth, MIN_FLOW_WIDTH)
              : t.node.kind !== "atom";
          if (multi || tooLong) out += "\n" + indent + text;
          else out += " " + text;
        }
      }
      const lastNl = out.lastIndexOf("\n");
      lineLen = lastNl < 0 ? out.length : out.length - lastNl - 1 - indent.length;
    });
    return out;
  }
}

// ---- public API ------------------------------------------------------------

function stripLeadingEquals(formula: string): string {
  return formula.charAt(0) === "=" ? formula.slice(1) : formula;
}

/** The whitespace-free (except intersections) form used for the safety check. */
function canonical(list: Expr[]): string {
  return list.map((e) => renderExprInline(e, MINIFY_STYLE)).join(",");
}

function failure(formula: string, e: unknown): FormatResult {
  const msg = e instanceof FormulaSyntaxError ? "could not read the formula (" + e.message + ")" : "formatter error: " + (e as Error).message;
  return { ok: false, value: formula, reason: msg };
}

function checked(formula: string, out: string, list: Expr[]): FormatResult {
  try {
    if (canonical(parse(stripLeadingEquals(out))) !== canonical(list)) {
      return { ok: false, value: formula, reason: "formula uses syntax the formatter cannot safely rewrite" };
    }
  } catch {
    return { ok: false, value: formula, reason: "formula uses syntax the formatter cannot safely rewrite" };
  }
  return { ok: true, value: out };
}

/**
 * Expand a formula onto indented lines. When `ok` is false the original formula is
 * returned unchanged (safe fallback).
 */
export function beautifyFormula(formula: string, options: Partial<FormatOptions> = {}): FormatResult {
  const opts: FormatOptions = { ...DEFAULT_FORMAT, ...options };
  let list: Expr[];
  try {
    list = parse(stripLeadingEquals(formula));
  } catch (e) {
    return failure(formula, e);
  }
  const b = new Beautifier(opts);
  const body = list.length === 1 ? b.expr(list[0], 0) : list.map((e) => b.expr(e, 0)).join(",\n");
  return checked(formula, "=" + body, list);
}

/** Collapse a formula to a single line, with optional readability spacing. */
export function minifyFormulaSafe(formula: string, spacing: SpacingOptions = DEFAULT_SPACING): FormatResult {
  let list: Expr[];
  try {
    list = parse(stripLeadingEquals(formula));
  } catch (e) {
    return failure(formula, e);
  }
  const style = inlineStyle(spacing, false);
  return checked(formula, "=" + list.map((e) => renderExprInline(e, style)).join(style.sep), list);
}

/** Convenience wrapper: minified formula, or the original when it cannot be read. */
export function minifyFormula(formula: string, spacing: SpacingOptions = DEFAULT_SPACING): string {
  return minifyFormulaSafe(formula, spacing).value;
}

// ---- display conversion ------------------------------------------------------

const NUMBER_DOT = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
const NUMBER_COMMA = /^(\d+,?\d*|,\d+)([eE][+-]?\d+)?$/;

/**
 * Convert a canonical (comma) formula to the semicolon display form: argument
 * separators become ";" and decimal points become ",". Array constants are left
 * as they are. Whitespace and line breaks are preserved.
 */
export function toDisplayFormula(formula: string, sep: "," | ";"): string {
  if (sep === "," || !isFormula(formula)) return formula;
  try {
    return "=" + tokenize(stripLeadingEquals(formula), ",")
      .map((t) => {
        if (t.type === "sep") return ";";
        if (t.type === "atom" && NUMBER_DOT.test(t.text)) return t.text.replace(".", ",");
        return t.text;
      })
      .join("");
  } catch {
    return formula;
  }
}

/** Inverse of toDisplayFormula. */
export function fromDisplayFormula(formula: string, sep: "," | ";"): string {
  if (sep === "," || !isFormula(formula)) return formula;
  try {
    return "=" + tokenize(stripLeadingEquals(formula), ";")
      .map((t) => {
        if (t.type === "sep") return ",";
        if (t.type === "atom" && NUMBER_COMMA.test(t.text)) return t.text.replace(",", ".");
        return t.text;
      })
      .join("");
  } catch {
    return formula;
  }
}
