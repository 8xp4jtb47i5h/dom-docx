/**
 * Table-of-contents guard — structural OOXML checks for the `tableOfContents`
 * config option.
 *
 * A TOC is a Word field: page numbers are computed by the word processor when the
 * document is opened. Word refreshes the field on open (via `w:updateFields`), but
 * LibreOffice and other viewers do not — so we also ship *cached entries* (the
 * heading titles) inside the field, or those viewers would show an empty TOC.
 *
 * These are structural checks: the `w:sdt` block, the `TOC` instruction with the
 * right `\o` heading range and `\h` hyperlink switch, the `w:dirty` flag, the
 * `w:updateFields` setting, the cached entry titles (honoring `headingRange`), and
 * OOXML schema validity. Critically, the heading styles must carry explicit
 * `w:outlineLvl` values: LibreOffice collects TOC entries strictly by outline
 * level, so without them "Update Table of Contents" regenerates to an empty table
 * (Word infers the levels from the built-in Heading styles; LibreOffice does not).
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { convertHtmlToDocx, type ConvertOptions } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "toc");

const HTML = `
  <h1>Introduction</h1>
  <p>Opening paragraph.</p>
  <h2>Background</h2>
  <p>Some background.</p>
  <h3>Details</h3>
  <p>Fine print.</p>
  <h4>Appendix</h4>
  <p>Aside.</p>
  <h1>Conclusion</h1>
  <p>Closing paragraph.</p>
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

interface Parts {
  doc: string;
  settings: string;
  styles: string;
  buf: Buffer;
}

async function build(options: ConvertOptions | undefined): Promise<Parts> {
  const buf = await convertHtmlToDocx(HTML, options);
  const files = unzipSync(new Uint8Array(buf));
  const dec = (p: string): string =>
    files[p] ? new TextDecoder().decode(files[p]) : "";
  return {
    doc: dec("word/document.xml"),
    settings: dec("word/settings.xml"),
    styles: dec("word/styles.xml"),
    buf,
  };
}

/** The `<w:style>` definition for a given style id, or "" if absent. */
function styleDef(styles: string, id: string): string {
  return styles.match(new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`))?.[0] ?? "";
}

/** The TOC field's own markup (`<w:sdt>…</w:sdt>`) — its cached entries live here. */
function tocField(doc: string): string {
  const start = doc.indexOf("<w:sdt>");
  const end = doc.indexOf("</w:sdt>");
  return start >= 0 && end > start ? doc.slice(start, end) : "";
}

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });
  console.log("Table-of-contents guard — structural checks:");

  // ---- Omitted: no TOC field, no updateFields, headings intact ----
  console.log("\ntableOfContents omitted (baseline):");
  const off = await build(undefined);
  check("no w:sdt / TOC field", !/<w:sdt>/.test(off.doc) && !/\bTOC\b/.test(off.doc));
  check("no updateFields in settings", !/<w:updateFields/.test(off.settings));
  check("headings still map to Heading styles", /w:pStyle w:val="Heading1"/.test(off.doc));
  check("no outline levels added without a TOC", !/<w:outlineLvl/.test(styleDef(off.styles, "Heading1")));

  // ---- tableOfContents: true → defaults (levels 1-3, hyperlinked) ----
  console.log("\ntableOfContents: true (defaults):");
  const on = await build({ tableOfContents: true });
  check("emits a w:sdt block", /<w:sdt>/.test(on.doc));
  check("emits a TOC field instruction", /<w:instrText[^>]*>\s*TOC\b/.test(on.doc));
  check('default heading range \\o "1-3"', /TOC\b[^<]*\\o\s+&quot;1-3&quot;/.test(on.doc));
  check("hyperlinked by default (\\h switch)", /TOC\b[^<]*\\h\b/.test(on.doc));
  check("field flagged dirty", /w:fldCharType="begin"\s+w:dirty="true"/.test(on.doc));
  check("settings enables updateFields", /<w:updateFields\s*\/>/.test(on.settings));
  check(
    "TOC precedes first body heading",
    on.doc.indexOf("TOC") < on.doc.indexOf('w:pStyle w:val="Heading1"'),
  );
  check("no title paragraph by default", !on.doc.includes("Table of Contents</w:t>"));
  check("no trailing page break by default", !/<w:br w:type="page"\/>/.test(on.doc));

  // Regression guard: heading styles MUST carry explicit outline levels, or
  // LibreOffice's "create from outline" TOC finds nothing and empties itself when
  // the user hits Update Table of Contents (Word infers levels, LibreOffice doesn't).
  check('Heading1 style has outlineLvl "0"', /<w:outlineLvl w:val="0"\/>/.test(styleDef(on.styles, "Heading1")));
  check('Heading2 style has outlineLvl "1"', /<w:outlineLvl w:val="1"\/>/.test(styleDef(on.styles, "Heading2")));
  check('Heading3 style has outlineLvl "2"', /<w:outlineLvl w:val="2"\/>/.test(styleDef(on.styles, "Heading3")));

  // Cached entries: the field must ship visible entries so viewers that don't
  // auto-update fields (LibreOffice on open, Google Docs, PDF/preview) still show
  // a populated TOC. Default range 1-3 includes h1/h2/h3 but not the h4 sibling.
  const onField = tocField(on.doc);
  check("cached entry: Introduction (h1) in field", />Introduction</.test(onField));
  check("cached entry: Background (h2) in field", />Background</.test(onField));
  check("cached entry: Details (h3) in field", />Details</.test(onField));
  check("h4 'Appendix' excluded from default 1-3 range", !/>Appendix</.test(onField));

  // ---- Full config: title + custom range + no hyperlink + page break ----
  console.log("\ntableOfContents: { title, headingRange, hyperlink: false, pageBreakAfter }:");
  const full = await build({
    tableOfContents: {
      title: "Contents",
      headingRange: "1-2",
      hyperlink: false,
      pageBreakAfter: true,
    },
  });
  check("title rendered as bold paragraph", /<w:b\/>[\s\S]*?<w:t[^>]*>Contents<\/w:t>/.test(full.doc));
  check(
    "title is not itself a heading (no self-referential TOC entry)",
    full.doc.indexOf("Contents</w:t>") < full.doc.indexOf("<w:sdt>") &&
      !/Heading\d"[^>]*\/>[\s\S]{0,80}Contents<\/w:t>/.test(full.doc),
  );
  check('custom heading range \\o "1-2"', /TOC\b[^<]*\\o\s+&quot;1-2&quot;/.test(full.doc));
  check("hyperlink: false omits \\h switch", !/TOC\b[^<]*\\h\b/.test(full.doc));
  // Cached entries honor headingRange: h1/h2 in, h3 "Details" + h4 "Appendix" out.
  const fullField = tocField(full.doc);
  check("cached entry: Introduction (h1) in field", />Introduction</.test(fullField));
  check("cached entry: Background (h2) in field", />Background</.test(fullField));
  check("cached entry: Conclusion (h1) in field", />Conclusion</.test(fullField));
  check("h3 'Details' excluded from 1-2 range", !/>Details</.test(fullField));
  check("h4 'Appendix' excluded from 1-2 range", !/>Appendix</.test(fullField));
  check("pageBreakAfter emits a page break", /<w:br w:type="page"\/>/.test(full.doc));
  check(
    "page break sits after the TOC field, before the body",
    full.doc.indexOf('<w:br w:type="page"/>') > full.doc.indexOf("</w:sdt>") &&
      full.doc.indexOf('<w:br w:type="page"/>') < full.doc.indexOf('w:pStyle w:val="Heading1"'),
  );

  // ---- OOXML schema validity of a TOC document ----
  console.log("\nOOXML schema:");
  const docxPath = path.join(OUT_DIR, "output.docx");
  await writeFile(docxPath, full.buf);
  await writeFile(path.join(OUT_DIR, "source.html"), HTML.trim(), "utf-8");
  const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
  check(
    "TOC document is schema-valid",
    validation.ok,
    validation.errors.slice(0, 2).map((e) => e.description).join("; "),
  );

  const ok = failures === 0;
  await writeGuardResult({
    id: "toc",
    label: "Table of contents",
    passed: checksRun - failures,
    total: checksRun,
    ok,
    unit: "OOXML field structure + schema",
    command: "npm run guard:toc",
  });

  console.log(ok ? "\nAll table-of-contents checks passed." : `\n${failures} check(s) failed.`);
  if (!ok) process.exitCode = 1;
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
