/**
 * Internal-href guard — structural OOXML checks for `href="#id"` links.
 *
 * Same-document fragments become Word internal hyperlinks (`w:hyperlink
 * w:anchor="…"`) and matching `id` / legacy `a[name]` targets become bookmarks
 * (`w:bookmarkStart`). External URLs stay as relationship-based hyperlinks.
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { convertHtmlToDocx } from "../src/converter.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "internal-href");

const HTML = `
  <p>Jump to <a href="#section-two">section two</a> or
     <a href="#note">the note</a>.</p>
  <h2 id="section-two">Section Two</h2>
  <p>Body of section two.</p>
  <p>Inline target: <span id="note">important note</span> in a paragraph.</p>
  <p>Also see <a href="https://example.com">Example Domain</a>
     and a no-op <a href="#">hash-only</a> link.</p>
  <p>Encoded fragment: <a href="#caf%C3%A9">café</a>.</p>
  <h3 id="café">Café heading</h3>
  <a id="legacy-empty"></a>
  <p>After empty named anchor.</p>
  <a name="legacy-name">Legacy named anchor text</a>
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

function documentXml(buffer: Buffer): string {
  const files = unzipSync(new Uint8Array(buffer));
  const data = files["word/document.xml"];
  return data ? new TextDecoder().decode(data) : "";
}

function relationshipXml(buffer: Buffer): string {
  const files = unzipSync(new Uint8Array(buffer));
  const data = files["word/_rels/document.xml.rels"];
  return data ? new TextDecoder().decode(data) : "";
}

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });
  await writeFile(path.join(OUT_DIR, "source.html"), HTML.trim(), "utf-8");

  console.log("Internal-href guard — structural checks:");

  const buf = await convertHtmlToDocx(HTML);
  const docxPath = path.join(OUT_DIR, "output.docx");
  await writeFile(docxPath, buf);
  check("convertHtmlToDocx succeeds", buf.length > 500, `${buf.length} bytes`);

  const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
  check(
    "OOXML schema valid",
    validation.ok,
    validation.errors.slice(0, 2).map((e) => e.description).join("; "),
  );

  const doc = documentXml(buf);
  const rels = relationshipXml(buf);

  check(
    "internal link to section-two",
    /<w:hyperlink[^>]*w:anchor="section-two"/.test(doc),
  );
  check(
    "bookmark for section-two",
    /<w:bookmarkStart[^>]*w:name="section-two"/.test(doc),
  );
  check(
    "internal link to note",
    /<w:hyperlink[^>]*w:anchor="note"/.test(doc),
  );
  check(
    "bookmark for note (inline span)",
    /<w:bookmarkStart[^>]*w:name="note"/.test(doc),
  );
  check(
    "decoded fragment café → internal link",
    /<w:hyperlink[^>]*w:anchor="café"/.test(doc) || /<w:hyperlink[^>]*w:anchor="caf&#233;"/.test(doc) || /w:anchor="caf&eacute;"/.test(doc),
  );
  check(
    "bookmark for café heading",
    /<w:bookmarkStart[^>]*w:name="café"/.test(doc) || /w:name="caf&#233;"/.test(doc),
  );
  check(
    "empty a[id] emits bookmark",
    /<w:bookmarkStart[^>]*w:name="legacy-empty"/.test(doc),
  );
  check(
    "legacy a[name] emits bookmark",
    /<w:bookmarkStart[^>]*w:name="legacy-name"/.test(doc),
  );
  check(
    "external https link still uses relationship",
    /Target="https:\/\/example\.com"/.test(rels) && /<w:hyperlink[^>]*r:id=/.test(doc),
  );
  check(
    "bare href=\"#\" is not an internal anchor",
    !/<w:hyperlink[^>]*w:anchor=""/.test(doc),
  );

  // Every internal anchor we emit should have a matching bookmark.
  const anchors = [...doc.matchAll(/<w:hyperlink[^>]*w:anchor="([^"]+)"/g)].map((m) => m[1]);
  check("at least one internal hyperlink", anchors.length >= 2, `found ${anchors.length}`);
  for (const anchor of anchors) {
    const decoded = anchor.replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
    check(
      `anchor "${decoded}" has matching bookmark`,
      doc.includes(`w:name="${anchor}"`) || doc.includes(`w:name="${decoded}"`),
    );
  }

  await writeGuardResult({
    id: "internal-href",
    label: "Internal hrefs",
    passed: checksRun - failures,
    total: checksRun,
    ok: failures === 0,
    unit: "structural checks",
    command: "npm run guard:internal-href",
  });

  if (failures > 0) {
    console.error(`\n${failures} check(s) failed`);
    process.exit(1);
  }
  console.log(`\nAll ${checksRun} checks passed`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
