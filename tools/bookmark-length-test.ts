/**
 * Bookmark-length guard — `w:bookmarkStart/@w:name` must satisfy the OOXML
 * ST_String MaxLength of 40.
 *
 * Real documentation pages carry section ids well past that (the Red Hat
 * storage page produced 71 schema errors from ids like
 * `deleting-a-file-system-from-a-stratis-pool-by-using-the-web-console`).
 * Long names are truncated with a hash suffix derived from the full id, so
 * `href="#long-id"` still resolves to the bookmark the `id` emitted and two ids
 * sharing a long prefix stay distinct.
 */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { unzipSync } from "fflate";
import { validateFile } from "@xarsh/ooxml-validator";
import { convertHtmlToDocx } from "../src/converter.js";
import { BOOKMARK_NAME_MAX_LENGTH, normalizeBookmarkId } from "../src/converter/bookmarks.js";
import { writeGuardResult } from "./guard-result.js";
import { GUARDS_OUTPUT } from "./output-paths.js";

const OUT_DIR = path.join(GUARDS_OUTPUT, "bookmark-length");

/** The id from the RHEL page that motivated the fix (67 code points). */
const RHEL_ID = "deleting-a-file-system-from-a-stratis-pool-by-using-the-web-console";

/** Two ids identical for the first 40 characters — must not collapse into one. */
const SHARED_PREFIX = "configuring-an-encrypted-storage-pool-on-";
const SIBLING_A = `${SHARED_PREFIX}a-local-disk`;
const SIBLING_B = `${SHARED_PREFIX}a-remote-target`;

const HTML = `
  <p>Jump to <a href="#${RHEL_ID}">the long section</a>.</p>
  <h2 id="${RHEL_ID}">Long section</h2>
  <p>Body of the long section.</p>

  <p>Siblings: <a href="#${SIBLING_A}">A</a> and <a href="#${SIBLING_B}">B</a>.</p>
  <h2 id="${SIBLING_A}">Sibling A</h2>
  <h2 id="${SIBLING_B}">Sibling B</h2>

  <p>Short target: <a href="#intro">intro</a>.</p>
  <h2 id="intro">Intro</h2>

  <p>Encoded long fragment: <a href="#${encodeURIComponent(RHEL_ID)}">encoded</a>.</p>
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

function codePoints(value: string): number {
  return [...value].length;
}

async function main(): Promise<void> {
  await mkdir(OUT_DIR, { recursive: true });
  await writeFile(path.join(OUT_DIR, "source.html"), HTML.trim(), "utf-8");

  console.log("Bookmark-length guard — normalizeBookmarkId:");

  // --- Pure-function checks -------------------------------------------------
  const long = normalizeBookmarkId(RHEL_ID)!;
  check(
    "long id truncated to the schema limit",
    codePoints(long) <= BOOKMARK_NAME_MAX_LENGTH,
    `${codePoints(long)} code points: ${long}`,
  );
  check("truncated name keeps a readable prefix", long.startsWith("deleting-a-file-system"), long);
  check(
    "truncation is deterministic",
    normalizeBookmarkId(RHEL_ID) === long,
    "two calls disagreed",
  );
  check(
    "href fragment normalizes to the same name as the id",
    normalizeBookmarkId(`${RHEL_ID}`) === long,
  );
  check(
    "percent-encoded fragment normalizes to the same name",
    normalizeBookmarkId(encodeURIComponent(RHEL_ID)) === long,
  );

  const a = normalizeBookmarkId(SIBLING_A)!;
  const b = normalizeBookmarkId(SIBLING_B)!;
  check("ids sharing a >40-char prefix stay distinct", a !== b, `${a} vs ${b}`);
  check(
    "both siblings within the limit",
    codePoints(a) <= BOOKMARK_NAME_MAX_LENGTH && codePoints(b) <= BOOKMARK_NAME_MAX_LENGTH,
    `${codePoints(a)} / ${codePoints(b)}`,
  );

  // Boundary + no-regression: anything at or under the limit is untouched.
  const exactly40 = "a".repeat(BOOKMARK_NAME_MAX_LENGTH);
  const over40 = "a".repeat(BOOKMARK_NAME_MAX_LENGTH + 1);
  check("id of exactly 40 chars is unchanged", normalizeBookmarkId(exactly40) === exactly40);
  check("id of 41 chars is truncated", codePoints(normalizeBookmarkId(over40)!) === BOOKMARK_NAME_MAX_LENGTH);
  check("short id is unchanged", normalizeBookmarkId("intro") === "intro");
  check("whitespace-only id is undefined", normalizeBookmarkId("   ") === undefined);
  check("empty id is undefined", normalizeBookmarkId("") === undefined);
  check("null id is undefined", normalizeBookmarkId(null) === undefined);

  // Astral characters: truncation counts code points and must not split a pair.
  const astral = "🚀".repeat(50);
  const astralName = normalizeBookmarkId(astral)!;
  check(
    "astral id truncated by code point, no lone surrogate",
    codePoints(astralName) <= BOOKMARK_NAME_MAX_LENGTH && !/[\uD800-\uDFFF]/.test(
      astralName.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, ""),
    ),
    `${codePoints(astralName)} code points`,
  );
  // 40 code points of non-ASCII is still under the limit and must pass through.
  const accented = "é".repeat(BOOKMARK_NAME_MAX_LENGTH);
  check("40 non-ASCII chars pass through unchanged", normalizeBookmarkId(accented) === accented);

  // --- End-to-end structural checks ----------------------------------------
  console.log("\nBookmark-length guard — emitted OOXML:");

  const buf = await convertHtmlToDocx(HTML);
  const docxPath = path.join(OUT_DIR, "output.docx");
  await writeFile(docxPath, buf);
  check("convertHtmlToDocx succeeds", buf.length > 500, `${buf.length} bytes`);

  const validation = await validateFile(docxPath, { officeVersion: "Office2019" });
  check(
    "OOXML schema valid (no MaxLength errors)",
    validation.ok,
    validation.errors.slice(0, 3).map((e) => e.description).join("; "),
  );

  const doc = documentXml(buf);

  const names = [...doc.matchAll(/<w:bookmarkStart[^>]*w:name="([^"]*)"/g)].map((m) => m[1]);
  check("document emits bookmarks", names.length >= 4, `found ${names.length}`);
  const tooLong = names.filter((n) => codePoints(n) > BOOKMARK_NAME_MAX_LENGTH);
  check(
    "every emitted bookmark name is within the limit",
    tooLong.length === 0,
    tooLong.join(", "),
  );

  // The whole point of the hash suffix: links must still land on their target.
  const anchors = [...doc.matchAll(/<w:hyperlink[^>]*w:anchor="([^"]+)"/g)].map((m) => m[1]);
  check("document emits internal hyperlinks", anchors.length >= 4, `found ${anchors.length}`);
  for (const anchor of anchors) {
    check(`anchor "${anchor}" has a matching bookmark`, names.includes(anchor));
  }

  check("truncated long section is linkable", anchors.includes(long) && names.includes(long));
  check(
    "sibling anchors resolve to different bookmarks",
    anchors.includes(a) && anchors.includes(b) && a !== b,
  );
  check("short bookmark unaffected end-to-end", names.includes("intro"));

  await writeGuardResult({
    id: "bookmark-length",
    label: "Bookmark name length",
    passed: checksRun - failures,
    total: checksRun,
    ok: failures === 0,
    unit: "structural checks",
    command: "npm run guard:bookmark-length",
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
