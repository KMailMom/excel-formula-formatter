/* Run with: npm test (compiles src/lib/formatter.ts + this file to .test-build and runs it). */
import * as assert from "assert";
import { beautifyFormula, minifyFormula, minifyFormulaSafe, toDisplayFormula, fromDisplayFormula } from "../src/lib/formatter";

let failed = 0;
function test(name: string, fn: () => void): void {
  try {
    fn();
    console.log("ok   " + name);
  } catch (e) {
    failed++;
    console.log("FAIL " + name + "\n" + (e as Error).message);
  }
}

const USER =
  '=IF([@Soort]<>"Nevenwerkzaamheid","",LET(k,[@Werknr]&"|"&LEFT(LOWER(TRIM([@Opdrachtgever]&"")),80)&"|"&N([@[Termijn van]])&"|"&N([@Tot])&"|"&LEFT(LOWER(TRIM([@Omschrijving]&"")),80),k))';
const indent = { useSpaces: true, size: 4 };
const spacing = { afterSeparator: false, insideParens: false };

test("minify keeps spaces inside structured references", () => {
  const b = beautifyFormula(USER, { inlineWidth: 0, indent, spacing });
  assert.ok(b.ok, b.reason);
  assert.strictEqual(minifyFormula(b.value), USER);
  assert.ok(minifyFormula(b.value).includes("[@[Termijn van]]"));
});

test("minify keeps spaces in strings, sheet names, [#This Row]", () => {
  const f = "=SUM('My Sheet'!A1:A3, Table1[[#This Row],[Col A]]) & \" a  b \"";
  assert.strictEqual(minifyFormula(f), "=SUM('My Sheet'!A1:A3,Table1[[#This Row],[Col A]])&\" a  b \"");
});

test("minify keeps intersection space", () => {
  assert.strictEqual(minifyFormula("=SUM( A1:A10   B1:B10 )"), "=SUM(A1:A10 B1:B10)");
});

test("minify spacing options", () => {
  assert.strictEqual(minifyFormula("=IF(A1>5,SUM(A1,B1),TODAY())", { afterSeparator: true, insideParens: false }), "=IF(A1>5, SUM(A1, B1), TODAY())");
  assert.strictEqual(minifyFormula("=IF(A1>5,SUM(A1,B1),TODAY())", { afterSeparator: false, insideParens: true }), "=IF( A1>5,SUM( A1,B1 ),TODAY() )");
  assert.strictEqual(minifyFormula("=IF(A1>5,SUM(A1,B1),\"\")", { afterSeparator: true, insideParens: true }), '=IF( A1>5, SUM( A1, B1 ), "" )');
});

test("expanded mode breaks every call", () => {
  const b = beautifyFormula("=IF(A1>5,N(B1),0)", { inlineWidth: 0, indent, spacing });
  assert.strictEqual(b.value, "=IF(\n    A1 > 5,\n    N(\n        B1\n    ),\n    0\n)");
});

test("width mode keeps short calls on one line", () => {
  const b = beautifyFormula(USER, { inlineWidth: 25, indent, spacing });
  assert.ok(b.ok, b.reason);
  assert.ok(b.value.includes("N([@[Termijn van]])"));
});

test("short formula stays on one line", () => {
  assert.strictEqual(beautifyFormula("=SUM(A1:A3,B1)", { inlineWidth: 50, indent, spacing }).value, "=SUM(A1:A3,B1)");
  assert.strictEqual(
    beautifyFormula("=SUM(A1:A3,B1)", { inlineWidth: 50, indent, spacing: { afterSeparator: true, insideParens: true } }).value,
    "=SUM( A1:A3, B1 )"
  );
});

test("balanced on the user formula", () => {
  const b = beautifyFormula(USER, { inlineWidth: 50, indent, spacing: { afterSeparator: true, insideParens: false } });
  assert.ok(b.ok, b.reason);
});

test("tricky syntax round-trips", () => {
  const cases = [
    "=-A1%+--B2",
    "=1E+5*2.5e-3",
    "=SUM({1,2;3,4})*{\"a}\",\"b\"}",
    "=LAMBDA(x,x+1)(5)",
    "=IF(A1,,1)",
    "=INDEX(A:A,1):INDEX(A:A,3)",
    "=SUM((A1:A2,B1:B2))",
    "=Sheet1!$A$1+'O''Brien'!B2",
    "=A1#",
    "=@A1:A10",
    "=#N/A",
    "=IFERROR(1/0,#DIV/0!)",
    "=\"he said \"\"hi\"\"\"",
    "=[1]Sheet1!A1",
    "=Table1[@[Col]:[Col2]]",
    "=LET(a,1,b,2,a+b)",
    "=TRUE<>FALSE",
    "=A1>=B1",
  ];
  for (const f of cases) {
    for (const w of [0, 20, 80]) {
      const b = beautifyFormula(f, { inlineWidth: w, indent, spacing });
      assert.ok(b.ok, f + " -> " + b.reason);
      assert.strictEqual(minifyFormula(b.value), minifyFormula(f), f);
    }
    assert.strictEqual(minifyFormula(f), f, "minify of already-minified " + f);
  }
});

test("bad formulas are left unchanged", () => {
  const r = minifyFormulaSafe("=SUM(A1");
  assert.strictEqual(r.ok, false);
  assert.strictEqual(r.value, "=SUM(A1");
});

test("display conversion", () => {
  const f = '=IF(A1>1.5,ROUND(B1,2),"a,b")+SUM({1,2})';
  const d = toDisplayFormula(f, ";");
  assert.strictEqual(d, '=IF(A1>1,5;ROUND(B1;2);"a,b")+SUM({1,2})');
  assert.strictEqual(fromDisplayFormula(d, ";"), f);
  const multi = "=IF(\n    A1,\n    1.5\n)";
  assert.strictEqual(toDisplayFormula(multi, ";"), "=IF(\n    A1;\n    1,5\n)");
});

if (failed) {
  console.log(failed + " test(s) failed");
  process.exit(1);
}
