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

/** `w:bookmarkStart/@w:name` is ST_String with a schema MaxLength of 40. */
export const BOOKMARK_NAME_MAX_LENGTH = 40;

/** Base36 chars of the disambiguating hash appended to a truncated name. */
const BOOKMARK_HASH_LENGTH = 6;

/**
 * FNV-1a (32-bit) over UTF-16 code units, as base36. Only needs to be stable
 * and well-spread, not cryptographic: it disambiguates ids that share a
 * truncated prefix, so the same long id always yields the same short name.
 */
function shortHash(value: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(36).padStart(BOOKMARK_HASH_LENGTH, "0").slice(-BOOKMARK_HASH_LENGTH);
}

/**
 * Shorten a bookmark name to the OOXML 40-character limit. Real docs pages
 * carry section ids well past it (e.g. Red Hat's
 * `deleting-a-file-system-from-a-stratis-pool-by-using-the-web-console`),
 * which Word/LibreOffice open but schema validation rejects.
 *
 * Must stay a pure function of `value`: `normalizeBookmarkId` (the `id` target)
 * and `internalAnchorFromHref` (the `href="#id"` link) run independently on
 * different nodes, so a counter-based suffix would desynchronize them and break
 * the jump. A hash of the full id keeps both sides in agreement and keeps two
 * ids that share the first 33 characters distinct.
 *
 * Length is counted in code points to match XSD `maxLength` semantics and to
 * avoid slicing a surrogate pair in half.
 */
function truncateBookmarkName(value: string): string {
  const codePoints = [...value];
  if (codePoints.length <= BOOKMARK_NAME_MAX_LENGTH) return value;
  const keep = BOOKMARK_NAME_MAX_LENGTH - BOOKMARK_HASH_LENGTH - 1;
  return `${codePoints.slice(0, keep).join("")}-${shortHash(value)}`;
}

/**
 * Normalize an HTML id / fragment for use as a Word bookmark name (`w:name`)
 * and internal hyperlink anchor (`w:anchor`). Returns `undefined` when empty.
 * Names longer than the schema limit are truncated (see
 * `truncateBookmarkName`).
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
  if (!value) return undefined;
  return truncateBookmarkName(value);
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
