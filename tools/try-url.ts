#!/usr/bin/env -S npx tsx
// Ad-hoc: convert a real, live webpage to .docx for manual inspection.
// Not part of the test suite — just a way to point the converter at arbitrary
// URLs (like the Red Hat docs page from the bug report) and see what happens.
//
// Usage (run with tsx — it imports TypeScript from src/):
//   npx tsx tools/try-url.ts <url> [contentSelector] [outFile]
//
// Examples:
//   npx tsx tools/try-url.ts https://docs.redhat.com/.../managing_storage_devices/index ".content"
//   npx tsx tools/try-url.ts https://example.com "main" out.docx
import { writeFile } from "node:fs/promises";
import { chromium } from "playwright";
import { convertHtmlToDocx, type ImageResolver } from "../src/index.js";

const [, , url, contentSelector = "body", outFile = "try-url-output.docx"] = process.argv;
if (!url) {
  console.error("Usage: npx tsx tools/try-url.ts <url> [contentSelector] [outFile]");
  process.exit(1);
}

// Headed: some sites block/stall the headless UA. This also uses a real Chrome
// user-agent (no "HeadlessChrome"), which gets past most bot checks.
const browser = await chromium.launch({ headless: false });
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  // `networkidle` often never settles on chatty docs sites (analytics, long-poll)
  // and times out — wait for `load` (all resources fetched) with a generous cap.
  await page.goto(url, { waitUntil: "load", timeout: 60_000 });

  const root = await page.$(contentSelector);
  if (!root) throw new Error(`no element matched selector "${contentSelector}" on ${url}`);

  // Unsanitized resolver — fine for local, one-off inspection.
  // Do NOT ship this as-is: no host allowlist, no size cap, no private-IP guard.
  //
  // Fetch through the browser's request context (page.request), NOT a bare Node
  // `fetch`: it reuses the page's cookies, real user-agent, and origin, so image
  // CDNs behind a bot filter (e.g. Akamai on access.redhat.com) serve the bytes
  // instead of a challenge/403. A plain fetch from Node silently returned nothing,
  // so every image dropped to alt text and the figures came out empty.
  const imageResolver: ImageResolver = async (src) => {
    try {
      const absolute = new URL(src, url).href;
      const res = await page.request.get(absolute, { headers: { referer: url } });
      if (!res.ok()) return null;
      const buf = new Uint8Array(await res.body());
      const ct = res.headers()["content-type"] ?? "";
      const type = ct.includes("png") ? "png" : ct.includes("gif") ? "gif" : ct.includes("bmp") ? "bmp" : "jpg";
      return { data: buf, type };
    } catch {
      return null;
    }
  };

  const docx = await convertHtmlToDocx(await root.innerHTML(), {
    styleSource: "computed",
    page,
    rootSelector: contentSelector,
    imageResolver,
    // default onWarning (console.warn) — leave it on so degradation is visible
  });

  await writeFile(outFile, docx);
  console.error(`wrote ${outFile} (${docx.length} bytes)`);
} finally {
  await browser.close();
}
