/**
 * TOC-slot guard — structural OOXML checks for the `tocHtml` option.
 *
 * `tocHtml` is a caller-provided table of contents: an HTML fragment placed after
 * the cover page (if any) and before the body. The caller owns the markup and
 * styling; in-page links (`<a href="#id">`) resolve to `id` bookmarks in the body
 * (via the internal-href support). This asserts the slot renders in the right
 * place, its links become internal hyperlinks pointing at real body bookmarks, and
 * the full cover → toc → body document is schema-valid.
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { convertHtmlToDocx, type ConvertOptions } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "toc-slot");

const BODY = `
  <h2 id="intro">Introduction</h2>
  <p>Opening.</p>
  <h2 id="results">Results</h2>
  <p>Findings.</p>
`;

const TOC = `
  <div style="background:#f8fafc;border:1px solid #e2e8f0;padding:12px">
    <p style="font-weight:bold;color:#64748b">ON THIS PAGE</p>
    <ol>
      <li><a href="#intro">Introduction</a></li>
      <li><a href="#results">Results</a></li>
    </ol>
  </div>
`;

let failures = 0;
let checksRun = 0;
function check(name: string, cond: boolean, detail?: string): void {
  checksRun += 1;
  if (cond) console.log(`  ✓ ${name}`);
  else {
    console.error(`  ✗ ${name}${detail ? ` — ${detail}` : ""}`);
    failures += 1;
  }
}

async function build(options: ConvertOptions | undefined): Promise<{ doc: string; buf: Buffer }> {
  const buf = await convertHtmlToDocx(BODY, options);
  const files = unzipSync(new Uint8Array(buf));
  return { doc: new TextDecoder().decode(files["word/document.xml"] ?? new Uint8Array()), buf };
}

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });
  console.log("TOC-slot guard — structural checks:");

  // ---- Omitted: no slot content ----
  console.log("\ntocHtml omitted (baseline):");
  const off = await build(undefined);
  check("no ON THIS PAGE marker", !off.doc.includes("ON THIS PAGE"));

  // ---- tocHtml alone: renders before the body ----
  console.log("\ntocHtml (no cover):");
  const slot = await build({ tocHtml: TOC });
  check("slot content present", slot.doc.includes("ON THIS PAGE"));
  check(
    "slot precedes body",
    slot.doc.indexOf("ON THIS PAGE") < slot.doc.indexOf("Introduction"),
  );
  check("in-page links are internal hyperlinks", /<w:hyperlink[^>]*w:anchor="intro"/.test(slot.doc));
  check("body headings carry matching bookmarks", /<w:bookmarkStart[^>]*w:name="intro"/.test(slot.doc));
  const anchors = [...slot.doc.matchAll(/<w:hyperlink[^>]*w:anchor="([^"]+)"/g)].map((m) => m[1]);
  check("both TOC entries are internal hyperlinks", anchors.includes("intro") && anchors.includes("results"));
  check(
    "every slot anchor has a matching body bookmark",
    anchors.length > 0 && anchors.every((a) => slot.doc.includes(`w:name="${a}"`)),
  );

  // ---- cover + tocHtml: order is cover < toc slot < body ----
  console.log("\ncoverHtml + tocHtml (Cover > TOC > Content):");
  const full = await build({
    coverHtml: "<h1>Report Cover</h1>",
    tocHtml: TOC,
  });
  const iCover = full.doc.indexOf("Report Cover");
  const iBreak = full.doc.indexOf('<w:br w:type="page"/>');
  const iToc = full.doc.indexOf("ON THIS PAGE");
  const iBody = full.doc.indexOf("Introduction");
  check(
    "order: cover < page break < toc slot < body",
    iCover >= 0 && iCover < iBreak && iBreak < iToc && iToc < iBody,
  );

  // ---- OOXML schema validity ----
  console.log("\nOOXML schema:");
  const docxPath = path.join(OUT_DIR, "output.docx");
  await writeFile(docxPath, full.buf);
  await writeFile(path.join(OUT_DIR, "source.html"), `${TOC}\n${BODY}`.trim(), "utf-8");
  const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
  check(
    "cover + toc slot document is schema-valid",
    validation.ok,
    validation.errors.slice(0, 2).map((e) => e.description).join("; "),
  );

  const ok = failures === 0;
  await writeGuardResult({
    id: "toc-slot",
    label: "TOC slot",
    passed: checksRun - failures,
    total: checksRun,
    ok,
    unit: "OOXML slot placement + internal links + schema",
    command: "npm run guard:toc-slot",
  });

  console.log(ok ? "\nAll TOC-slot checks passed." : `\n${failures} check(s) failed.`);
  if (!ok) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
