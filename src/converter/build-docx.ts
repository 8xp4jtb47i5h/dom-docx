import {
  AlignmentType,
  Document,
  Footer,
  Header,
  PageBreak,
  PageNumber,
  Packer,
  PageOrientation,
  Paragraph,
  TableOfContents,
  TextRun,
  convertInchesToTwip,
  type FileChild,
} from "docx";
import * as cheerio from "cheerio";
import type { CheerioAPI } from "cheerio";
import { unzipSync, zipSync } from "fflate";
import { BODY_FONT, BODY_FONT_HALF_POINTS, NUMBERING_CONFIG, PAGE_MARGIN_TWIPS } from "./constants.js";
import { patchDocumentXml, patchHeadingOutlineLevels, patchNumberingXml } from "./ooxml-patch.js";
import { applyImageResolver, resetImageDocPrIds, type ImageResolver } from "./image.js";
import { INLINE_STYLE_RESOLVER, type StyleResolver } from "./style-resolver.js";
import { htmlToDocxBlocks } from "./visitor.js";

/** Table-of-contents field options (see `DocumentConfig.tableOfContents`). */
export interface TableOfContentsConfig {
  /**
   * Heading text rendered above the TOC (e.g. `"Contents"`). Rendered as a plain
   * bold paragraph — not a Word heading — so it never appears as its own entry.
   * Omit for no title.
   */
  title?: string;
  /**
   * Heading levels to include, as a Word range, e.g. `"1-3"` (default) or `"1-2"`.
   * `h1`–`h6` map to Word Heading 1–6.
   */
  headingRange?: string;
  /** Render entries as clickable hyperlinks to their headings (default `true`). */
  hyperlink?: boolean;
  /** Insert a page break after the TOC so body content starts on a new page. */
  pageBreakAfter?: boolean;
}

/** Page/font/metadata options (Tier 1 `ConvertOptions`). All lengths in inches / points. */
export interface DocumentConfig {
  /** `"letter"` (default), `"a4"`, or a custom size in inches. */
  pageSize?: "letter" | "a4" | { width: number; height: number };
  orientation?: "portrait" | "landscape";
  /** Page margins in inches (each side defaults to 1). */
  margins?: { top?: number; right?: number; bottom?: number; left?: number };
  /** Default body font family and size (points). */
  defaultFont?: { family?: string; sizePt?: number };
  /** Core document properties written to `docProps/core.xml`. */
  metadata?: {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string[];
    description?: string;
  };
  /** HTML fragment rendered as the page header (its own inline-styled fragment). */
  headerHtml?: string;
  /** HTML fragment rendered as the page footer. */
  footerHtml?: string;
  /** Append a centered `Page N` field to the footer (created if `footerHtml` is absent). */
  pageNumber?: boolean;
  /** Document language (spell-check locale), e.g. `"en-US"`, `"ar-SA"`. */
  lang?: string;
  /** Text direction; `"rtl"` sets right-to-left paragraphs. */
  direction?: "ltr" | "rtl";
  /**
   * Insert a Table of Contents field at the top of the document, built from the
   * headings present (`h1`–`h6` become Word Heading 1–6). The heading titles are
   * cached into the field so the TOC is visible immediately in every viewer; page
   * numbers are computed by the word processor on open (the field is emitted
   * "dirty" with `updateFields`, so Word fills them in — and any reader can refresh
   * via right-click → Update Field / Tools → Update). `true` uses defaults (levels
   * 1–3, hyperlinked).
   */
  tableOfContents?: boolean | TableOfContentsConfig;
}

// Portrait dimensions in twips. Letter matches convertInchesToTwip(8.5)×(11).
const PAGE_PRESETS_TWIPS = {
  letter: { width: 12240, height: 15840 },
  a4: { width: 11906, height: 16838 },
} as const;

interface ResolvedConfig {
  size: { width: number; height: number; orientation?: (typeof PageOrientation)[keyof typeof PageOrientation] };
  margin: { top: number; right: number; bottom: number; left: number };
  font: string;
  fontHalfPoints: number;
  metadata: {
    title?: string;
    subject?: string;
    creator?: string;
    keywords?: string;
    description?: string;
  };
  lang?: string;
  rtl: boolean;
  toc?: TableOfContentsConfig;
}

function resolveTableOfContents(
  toc: DocumentConfig["tableOfContents"],
): TableOfContentsConfig | undefined {
  if (!toc) return undefined;
  return toc === true ? {} : toc;
}

function resolveDocumentConfig(config?: DocumentConfig): ResolvedConfig {
  const ps = config?.pageSize;
  const base =
    !ps || ps === "letter"
      ? PAGE_PRESETS_TWIPS.letter
      : ps === "a4"
        ? PAGE_PRESETS_TWIPS.a4
        : { width: convertInchesToTwip(ps.width), height: convertInchesToTwip(ps.height) };

  // docx swaps width/height itself for landscape, so pass portrait dims + the flag.
  const size =
    config?.orientation === "landscape"
      ? { width: base.width, height: base.height, orientation: PageOrientation.LANDSCAPE }
      : { width: base.width, height: base.height };

  const m = config?.margins;
  const marginIn = (v: number | undefined): number =>
    v !== undefined ? convertInchesToTwip(v) : PAGE_MARGIN_TWIPS;

  const meta = config?.metadata;
  const metadata: ResolvedConfig["metadata"] = {};
  if (meta?.title) metadata.title = meta.title;
  if (meta?.subject) metadata.subject = meta.subject;
  if (meta?.creator) metadata.creator = meta.creator;
  if (meta?.keywords?.length) metadata.keywords = meta.keywords.join(", ");
  if (meta?.description) metadata.description = meta.description;

  return {
    size,
    margin: { top: marginIn(m?.top), right: marginIn(m?.right), bottom: marginIn(m?.bottom), left: marginIn(m?.left) },
    font: config?.defaultFont?.family ?? BODY_FONT,
    fontHalfPoints:
      config?.defaultFont?.sizePt !== undefined
        ? Math.round(config.defaultFont.sizePt * 2)
        : BODY_FONT_HALF_POINTS,
    metadata,
    lang: config?.lang,
    rtl: config?.direction === "rtl",
    toc: resolveTableOfContents(config?.tableOfContents),
  };
}

/** Convert a standalone HTML fragment (header/footer) to DOCX blocks via the inline resolver. */
function fragmentToBlocks(html: string, sizeHalfPoints: number): FileChild[] {
  const $ = cheerio.load(`<body>${html.trim()}</body>`, { xml: false });
  return htmlToDocxBlocks($, INLINE_STYLE_RESOLVER, sizeHalfPoints);
}

function pageNumberParagraph(): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })],
  });
}

/** A pre-extracted TOC entry: heading text and its level (1–6). */
interface TocEntry {
  title: string;
  level: number;
}

const DEFAULT_HEADING_RANGE = "1-3";

/** Parse a Word heading range (`"1-3"`, `"2"`) into an inclusive `[min, max]`. */
function parseHeadingRange(range: string): [number, number] {
  const span = range.match(/(\d)\s*-\s*(\d)/);
  if (span) {
    const lo = Number(span[1]);
    const hi = Number(span[2]);
    return lo <= hi ? [lo, hi] : [hi, lo];
  }
  const single = range.match(/\d/);
  const n = single ? Number(single[0]) : 1;
  return [n, n];
}

/**
 * Heading text + level for every `h1`–`h6` in document order that falls inside
 * `headingRange`. These become the TOC's *cached* entries — the entries a reader
 * shows before the field is refreshed. Word regenerates them (with page numbers)
 * on open via `updateFields`, but LibreOffice and other viewers do not update
 * fields automatically, so without a cache they would show an empty TOC.
 */
function tocEntriesFromDom($: CheerioAPI, headingRange: string): TocEntry[] {
  const [min, max] = parseHeadingRange(headingRange);
  const selector = [];
  for (let level = min; level <= max; level++) selector.push(`h${level}`);
  if (!selector.length) return [];

  const entries: TocEntry[] = [];
  $(selector.join(",")).each((_, el) => {
    const tag = String($(el).prop("tagName") ?? "");
    const level = Number(tag[1]);
    if (!level) return;
    const title = $(el).text().replace(/\s+/g, " ").trim();
    if (title) entries.push({ title, level });
  });
  return entries;
}

/**
 * TOC field + optional title, prepended to the body. The field carries the TOC
 * instruction (`\o` heading range, `\h` hyperlinks) and is emitted dirty so Word
 * refreshes it on open. `entries` pre-populate the field's cached content so the
 * TOC is visible in viewers that do not auto-update fields (LibreOffice, Google
 * Docs, PDF/preview panes). Returns `[]` when no TOC is configured.
 */
function buildTableOfContents(resolved: ResolvedConfig, entries: TocEntry[]): FileChild[] {
  const toc = resolved.toc;
  if (!toc) return [];

  const blocks: FileChild[] = [];
  if (toc.title) {
    blocks.push(
      new Paragraph({
        children: [
          new TextRun({
            text: toc.title,
            bold: true,
            font: resolved.font,
            size: Math.round(resolved.fontHalfPoints * 1.5),
          }),
        ],
      }),
    );
  }

  blocks.push(
    new TableOfContents("Table of Contents", {
      hyperlink: toc.hyperlink ?? true,
      headingStyleRange: toc.headingRange ?? DEFAULT_HEADING_RANGE,
      // Cached entries carry titles only — real page numbers need layout, so we
      // leave them for the reader to fill on field update rather than fake them.
      cachedEntries: entries,
      beginDirty: true,
    }),
  );

  if (toc.pageBreakAfter) {
    blocks.push(new Paragraph({ children: [new PageBreak()] }));
  }
  return blocks;
}

function buildFooter(config: DocumentConfig | undefined, resolved: ResolvedConfig): Footer | undefined {
  const hasFooterHtml = Boolean(config?.footerHtml);
  if (!hasFooterHtml && !config?.pageNumber) return undefined;
  const children: FileChild[] = hasFooterHtml
    ? fragmentToBlocks(config!.footerHtml!, resolved.fontHalfPoints)
    : [];
  if (config?.pageNumber) children.push(pageNumberParagraph());
  return new Footer({ children });
}

function buildHeader(config: DocumentConfig | undefined, resolved: ResolvedConfig): Header | undefined {
  if (!config?.headerHtml) return undefined;
  return new Header({ children: fragmentToBlocks(config.headerHtml, resolved.fontHalfPoints) });
}

async function packDocxToUint8Array(
  children: FileChild[],
  resolved: ResolvedConfig,
  chrome: { header?: Header; footer?: Footer },
  tocEntries: TocEntry[],
): Promise<Uint8Array> {
  const listStyleRun = { font: resolved.font, size: resolved.fontHalfPoints };
  const doc = new Document({
    ...resolved.metadata,
    numbering: NUMBERING_CONFIG,
    styles: {
      default: {
        document: {
          run: {
            font: resolved.font,
            size: resolved.fontHalfPoints,
            ...(resolved.lang ? { language: { value: resolved.lang } } : {}),
            ...(resolved.rtl ? { rightToLeft: true } : {}),
          },
        },
      },
      paragraphStyles: [
        {
          id: "ListNumber",
          name: "List Number",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: listStyleRun,
        },
        {
          id: "ListBullet",
          name: "List Bullet",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: listStyleRun,
        },
      ],
    },
    // Flag every field (incl. the TOC) dirty so Word/LibreOffice populate it on open.
    ...(resolved.toc ? { features: { updateFields: true } } : {}),
    sections: [
      {
        properties: {
          page: {
            size: resolved.size,
            margin: resolved.margin,
          },
        },
        ...(chrome.header ? { headers: { default: chrome.header } } : {}),
        ...(chrome.footer ? { footers: { default: chrome.footer } } : {}),
        children: [...buildTableOfContents(resolved, tocEntries), ...children],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  return new Uint8Array(await blob.arrayBuffer());
}

function patchPackedDocx(packed: Uint8Array, withTocOutline: boolean): Uint8Array {
  const files = unzipSync(packed);
  const documentXml = new TextDecoder().decode(files["word/document.xml"]);
  files["word/document.xml"] = new TextEncoder().encode(patchDocumentXml(documentXml));
  if (files["word/numbering.xml"]) {
    const numberingXml = new TextDecoder().decode(files["word/numbering.xml"]);
    files["word/numbering.xml"] = new TextEncoder().encode(patchNumberingXml(numberingXml));
  }
  // A TOC collects by outline level; give the heading styles explicit levels so
  // LibreOffice can rebuild the field (Word already infers them). See the patch.
  if (withTocOutline && files["word/styles.xml"]) {
    const stylesXml = new TextDecoder().decode(files["word/styles.xml"]);
    files["word/styles.xml"] = new TextEncoder().encode(patchHeadingOutlineLevels(stylesXml));
  }
  return zipSync(files);
}

/** Platform-neutral DOCX bytes from an HTML body fragment and style resolver. */
export async function buildDocxUint8Array(
  html: string,
  styleResolver: StyleResolver,
  imageResolver?: ImageResolver,
  documentConfig?: DocumentConfig,
): Promise<Uint8Array> {
  resetImageDocPrIds();
  const resolved = resolveDocumentConfig(documentConfig);
  const $ = cheerio.load(`<body>${html.trim()}</body>`, { xml: false });
  if (imageResolver) await applyImageResolver($, imageResolver);
  const children = htmlToDocxBlocks($, styleResolver, resolved.fontHalfPoints);
  const tocEntries = resolved.toc
    ? tocEntriesFromDom($, resolved.toc.headingRange ?? DEFAULT_HEADING_RANGE)
    : [];
  const chrome = {
    header: buildHeader(documentConfig, resolved),
    footer: buildFooter(documentConfig, resolved),
  };
  const packed = await packDocxToUint8Array(children, resolved, chrome, tocEntries);
  return patchPackedDocx(packed, Boolean(resolved.toc));
}

/** Browser entry — returns a `.docx` Blob. */
export async function buildDocxBlob(
  html: string,
  styleResolver: StyleResolver,
  imageResolver?: ImageResolver,
  documentConfig?: DocumentConfig,
): Promise<Blob> {
  const bytes = await buildDocxUint8Array(html, styleResolver, imageResolver, documentConfig);
  return new Blob([bytes.slice()], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}
