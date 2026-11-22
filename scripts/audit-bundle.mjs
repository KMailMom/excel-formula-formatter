// Network audit of the built add-in (run after `npm run build`).
//
// Fails if the code that runs inside Excel (taskpane.html, taskpane.js, dialog.html):
//   - uses any browser API that can send data over the network, or
//   - references any external URL other than Microsoft's Office.js library.
// The GitHub Action runs this before every deploy, so the hosted add-in can't
// silently start "phoning home".
import { readFileSync } from "node:fs";
import { join } from "node:path";

const dist = new URL("../dist/", import.meta.url).pathname;
const runtimeFiles = ["taskpane.html", "taskpane.js", "dialog.html"];

const forbiddenApis = [
  /\bfetch\s*\(/,
  /\bXMLHttpRequest\b/,
  /\bWebSocket\b/,
  /\bEventSource\b/,
  /\bsendBeacon\b/,
  /\bimportScripts\b/,
  /\bRTCPeerConnection\b/,
  /\beval\s*\(/,
  /\bnew\s+Function\s*\(/,
];
const allowedUrlPrefixes = ["https://appsforoffice.microsoft.com/lib/1/hosted/office.js"];

let problems = 0;
for (const file of runtimeFiles) {
  const text = readFileSync(join(dist, file), "utf8");
  for (const re of forbiddenApis) {
    if (re.test(text)) {
      console.error(`FAIL ${file}: uses ${re.source}`);
      problems++;
    }
  }
  for (const [url] of text.matchAll(/(?:https?|wss?):\/\/[^\s"'`<>)]+/g)) {
    if (!allowedUrlPrefixes.some((p) => url.startsWith(p))) {
      console.error(`FAIL ${file}: external URL ${url}`);
      problems++;
    }
  }
}

if (problems) {
  console.error(`\nNetwork audit failed with ${problems} problem(s).`);
  process.exit(1);
}
console.log(`Network audit passed: ${runtimeFiles.join(", ")} make no network calls and load nothing but Office.js.`);
