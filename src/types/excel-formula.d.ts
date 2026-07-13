/*
 * Ambient declaration for the vendored global from
 * github.com/joshbtn/excelFormulaUtilitiesJS (src/lib/excel-formula.js).
 * The dist file is a plain browser script that assigns window.excelFormulaUtilities.
 */

export {};

interface EFU_Token {
  value: string;
  type: string; // "operand" | "function" | "operator-infix" | "white-space" | ...
  subtype: string; // "text" | "number" | "range" | "start" | "stop" | ...
}

interface EFU_Tokens {
  items: EFU_Token[];
}

interface EFU_CustomRenderResult {
  tokenString: string;
  useTemplate: boolean;
}

interface EFU_FormatOptions {
  tmplIndentTab?: string;
  tmplIndentSpace?: string;
  newLine?: string;
  autoLineBreak?: string;
  prefix?: string;
  postfix?: string;
  customTokenRender?:
    | ((tokenString: string, token: EFU_Token, indent?: string, lineBreak?: string) => EFU_CustomRenderResult)
    | null;
  [key: string]: unknown;
}

interface ExcelFormulaUtilities {
  formatFormula(formula: string, options?: EFU_FormatOptions): string;
  formatFormulaHTML(formula: string, options?: EFU_FormatOptions): string;
  getTokens(formula: string): EFU_Tokens;
}

declare global {
  interface Window {
    excelFormulaUtilities: ExcelFormulaUtilities;
  }
}
