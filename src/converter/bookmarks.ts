import {
  BookmarkEnd,
  BookmarkStart,
  LineRuleType,
  Paragraph,
  Table,
  bookmarkUniqueNumericIdGen,
  type ParagraphChild,
} from "docx";

/** Shared across the document so `w:bookmarkStart/@w:id` values stay unique.
 *  (The higher-level `Bookmark` class in `docx` mistakenly creates a fresh
 *  counter per instance, which duplicates id="1" and fails OOXML validation.) */
const nextBookmarkNumericId = bookmarkUniqueNumericIdGen();

/**
 * Normalize an HTML id / fragment for use as a Word bookmark name (`w:name`)
 * and internal hyperlink anchor (`w:anchor`). Returns `undefined` when empty.
 */
export function normalizeBookmarkId(raw: string | undefined | null): string | undefined {
  if (raw == null) return undefined;
  let value = raw.trim();
  if (!value) return undefined;
  try {
    value = decodeURIComponent(value);
  } catch {
    // Keep the raw fragment if it isn't valid %-encoding.
  }
  value = value.trim();
  return value || undefined;
}

/**
 * If `href` is a same-document fragment (`#id`), return the bookmark id.
 * Bare `#`, empty string, and external/relative URLs return `undefined`.
 */
export function internalAnchorFromHref(href: string): string | undefined {
  if (!href.startsWith("#")) return undefined;
  return normalizeBookmarkId(href.slice(1));
}

/** Paragraph children that may include bookmark markers (not in `ParagraphChild`). */
type BookmarkChild = ParagraphChild | BookmarkStart | BookmarkEnd;

/** Wrap runs in a `w:bookmarkStart`/`w:bookmarkEnd` pair (empty children allowed). */
export function wrapWithBookmark(
  id: string | undefined,
  children: readonly ParagraphChild[],
): BookmarkChild[] {
  const name = normalizeBookmarkId(id);
  if (!name) return [...children];
  const linkId = nextBookmarkNumericId();
  return [new BookmarkStart(name, linkId), ...children, new BookmarkEnd(linkId)];
}

/**
 * Zero-height paragraph that only places a bookmark. Used when an element with
 * `id` produces blocks so the jump target sits at the start of that content.
 */
export function bookmarkMarkerParagraph(id: string): Paragraph {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
    children: wrapWithBookmark(id, []) as ParagraphChild[],
  });
}

/** Prepend a bookmark marker when `idAttr` is a non-empty HTML id. */
export function prependElementBookmark(
  idAttr: string | undefined,
  blocks: Array<Paragraph | Table>,
): Array<Paragraph | Table> {
  const id = normalizeBookmarkId(idAttr);
  if (!id || blocks.length === 0) return blocks;
  return [bookmarkMarkerParagraph(id), ...blocks];
}
