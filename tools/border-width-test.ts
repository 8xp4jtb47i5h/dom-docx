/**
 * `border-width` guard — the CSS `border-width` / `border-{side}-width`
 * properties weren't handled by `parseInlineStyle` at all, so they fell
 * through to the `default:` case and left `borderTop`/etc. undefined
 * (no border emitted, even though the author declared one explicitly).
 */
import { cssToBlockLayout, parseInlineStyle } from "../src/converter/css.js";
import { writeGuardResult } from "./guard-result.js";

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

async function main(): Promise<void> {
  console.log("border-width guard — parseInlineStyle:");

  const uniform = parseInlineStyle("border-width: 2px");
  check("uniform border-width sets all four sides", uniform.borderTop?.widthPx === 2, JSON.stringify(uniform));
  check(
    "uniform border-width: same widthPx on every side",
    [uniform.borderRight, uniform.borderBottom, uniform.borderLeft].every((s) => s?.widthPx === 2),
  );

  const perSide = parseInlineStyle("border-width: 1px 2px 3px 4px");
  check(
    "4-value border-width follows CSS top/right/bottom/left order",
    perSide.borderTop?.widthPx === 1 &&
      perSide.borderRight?.widthPx === 2 &&
      perSide.borderBottom?.widthPx === 3 &&
      perSide.borderLeft?.widthPx === 4,
    JSON.stringify(perSide),
  );

  const keyword = parseInlineStyle("border-top-width: thin");
  check("keyword width (thin) resolves to 1px", keyword.borderTop?.widthPx === 1, JSON.stringify(keyword));
  check("border-top-width only sets the top side", keyword.borderRight === undefined && keyword.borderLeft === undefined);

  const zero = parseInlineStyle("border-width: 0");
  check(
    "border-width: 0 leaves borders undefined (matches border:0 / computed-style convention)",
    zero.borderTop === undefined && zero.borderRight === undefined,
    JSON.stringify(zero),
  );

  const longhand = parseInlineStyle(
    "border-left-width: 3px; border-right-width: medium; border-bottom-width: thick",
  );
  check(
    "individual *-width longhands each set their own side",
    longhand.borderLeft?.widthPx === 3 &&
      longhand.borderRight?.widthPx === 3 &&
      longhand.borderBottom?.widthPx === 5 &&
      longhand.borderTop === undefined,
    JSON.stringify(longhand),
  );

  console.log("\nborder-width guard — block border color fallback:");

  const noColor = cssToBlockLayout(parseInlineStyle("border-width: 2px; border-color: #cc3333"));
  check(
    "standalone border-color applies to a block border with no per-side color",
    noColor.borders?.top?.color === "cc3333" && noColor.borders?.left?.color === "cc3333",
    JSON.stringify(noColor.borders),
  );

  const shorthandColorWins = cssToBlockLayout(
    parseInlineStyle("border: 2px solid #2a6f2a; border-color: #cc3333"),
  );
  check(
    "color embedded in the border shorthand still wins over border-color",
    shorthandColorWins.borders?.top?.color === "2a6f2a",
    JSON.stringify(shorthandColorWins.borders),
  );

  const noOverride = cssToBlockLayout(parseInlineStyle("border-width: 2px"));
  check(
    "no border-color declared falls back to black",
    noOverride.borders?.top?.color === "000000",
    JSON.stringify(noOverride.borders),
  );

  console.log("\nborder-width guard — border-style default width:");

  const styleOnly = parseInlineStyle("border-style: solid");
  check(
    "border-style: solid alone defaults all four sides to medium (3px)",
    [styleOnly.borderTop, styleOnly.borderRight, styleOnly.borderBottom, styleOnly.borderLeft].every(
      (s) => s?.widthPx === 3,
    ),
    JSON.stringify(styleOnly),
  );

  const styleNone = parseInlineStyle("border-style: none");
  check(
    "border-style: none does not synthesize a border",
    styleNone.borderTop === undefined && styleNone.borderLeft === undefined,
    JSON.stringify(styleNone),
  );

  const styleWithExplicitWidth = parseInlineStyle("border-top-width: thick; border-style: solid");
  check(
    "an explicit width on one side is kept, medium default fills the rest",
    styleWithExplicitWidth.borderTop?.widthPx === 5 &&
      styleWithExplicitWidth.borderRight?.widthPx === 3 &&
      styleWithExplicitWidth.borderBottom?.widthPx === 3 &&
      styleWithExplicitWidth.borderLeft?.widthPx === 3,
    JSON.stringify(styleWithExplicitWidth),
  );

  const styleLonghandOneSide = parseInlineStyle("border-left-style: dashed");
  check(
    "border-left-style longhand only defaults the left side",
    styleLonghandOneSide.borderLeft?.widthPx === 3 &&
      styleLonghandOneSide.borderTop === undefined &&
      styleLonghandOneSide.borderRight === undefined &&
      styleLonghandOneSide.borderBottom === undefined,
    JSON.stringify(styleLonghandOneSide),
  );

  const styleWithShorthandBorder = parseInlineStyle("border: 1px solid #000; border-style: solid");
  check(
    "border-style default does not clobber a side already covered by the border shorthand",
    styleWithShorthandBorder.borderTop === undefined && styleWithShorthandBorder.border?.widthPx === 1,
    JSON.stringify(styleWithShorthandBorder),
  );

  const explicitZeroWithStyle = parseInlineStyle("border-width: 3px 0px; border-style: solid");
  check(
    "an explicit 0 width is not backfilled by the border-style medium default",
    explicitZeroWithStyle.borderTop?.widthPx === 3 &&
      explicitZeroWithStyle.borderBottom?.widthPx === 3 &&
      explicitZeroWithStyle.borderRight === undefined &&
      explicitZeroWithStyle.borderLeft === undefined,
    JSON.stringify(explicitZeroWithStyle),
  );

  console.log("\nborder-width guard — shorthand edge cases:");

  // A 0 component in the multi-value shorthand must not bail out the whole
  // declaration — only that side should end up with no border.
  const zeroInShorthand = parseInlineStyle("border-width: 1px 0px");
  check(
    "border-width: 1px 0px keeps the 1px top/bottom sides",
    zeroInShorthand.borderTop?.widthPx === 1 && zeroInShorthand.borderBottom?.widthPx === 1,
    JSON.stringify(zeroInShorthand),
  );
  check(
    "border-width: 1px 0px leaves the 0px left/right sides undefined (no border)",
    zeroInShorthand.borderRight === undefined && zeroInShorthand.borderLeft === undefined,
    JSON.stringify(zeroInShorthand),
  );

  // border-width/border-*-width only touch the width component in real CSS —
  // a color already declared via `border`/`border-top`/etc on that side must
  // survive a later width-only update, in either declaration order.
  const colorThenWidth = cssToBlockLayout(
    parseInlineStyle("border: 2px solid #ff0000; border-width: 6px;"),
  );
  check(
    "border-width after a colored border shorthand keeps that color",
    colorThenWidth.borders?.top?.color === "ff0000",
    JSON.stringify(colorThenWidth.borders),
  );

  const widthThenColor = cssToBlockLayout(
    parseInlineStyle("border-top-width: 6px; border-top: 2px solid #ff0000;"),
  );
  check(
    "a later colored border-top still applies its own color",
    widthThenColor.borders?.top?.color === "ff0000",
    JSON.stringify(widthThenColor.borders),
  );

  console.log("\nborder-width guard — explicit zero vs generic border fallback:");

  // buildBlockBorders (and the table border builders) resolve a side as
  // `css.borderX ?? css.border` — an explicitly zeroed side must not inherit
  // the generic `border` shorthand's value just because it resolves to the
  // same `undefined` a never-declared side would.
  const zeroSideKeepsOthers = cssToBlockLayout(
    parseInlineStyle("border: 2px solid #1a6fb0; border-right-width: 0;"),
  );
  check(
    "border-right-width: 0 removes only the right side, others keep the generic border",
    zeroSideKeepsOthers.borders?.top?.color === "1a6fb0" &&
      zeroSideKeepsOthers.borders?.bottom?.color === "1a6fb0" &&
      zeroSideKeepsOthers.borders?.left?.color === "1a6fb0" &&
      zeroSideKeepsOthers.borders?.right === undefined,
    JSON.stringify(zeroSideKeepsOthers.borders),
  );

  const plainFallbackStillWorks = cssToBlockLayout(parseInlineStyle("border: 2px solid #1a6fb0;"));
  check(
    "border alone (no per-side overrides) still falls back on all four sides",
    [
      plainFallbackStillWorks.borders?.top,
      plainFallbackStillWorks.borders?.right,
      plainFallbackStillWorks.borders?.bottom,
      plainFallbackStillWorks.borders?.left,
    ].every((s) => s?.color === "1a6fb0"),
    JSON.stringify(plainFallbackStillWorks.borders),
  );

  console.log("\nborder-width guard — zero and non-px widths in the border shorthand:");

  // parseBorderShorthand only recognized a width immediately followed by
  // "px" — a bare "0" (the one CSS length valid without a unit) fell through
  // to the "no width token" default of 1, producing a visible border from a
  // declaration that explicitly asked for none.
  const bareZeroBorder = parseInlineStyle("border: 0");
  check("border: 0 (no unit) means no border", bareZeroBorder.border === undefined, JSON.stringify(bareZeroBorder));

  const bareZeroSide = cssToBlockLayout(
    parseInlineStyle("border: 2px solid #1a6fb0; border-right: 0;"),
  );
  check(
    "border-right: 0 (no unit) removes only the right side",
    bareZeroSide.borders?.top?.color === "1a6fb0" && bareZeroSide.borders?.right === undefined,
    JSON.stringify(bareZeroSide.borders),
  );

  // Same "no width found" fallback also missed a zero in any other absolute
  // unit (0pt, 0in, …), not just a unitless zero.
  const nonPxZeroBorder = parseInlineStyle("border: 0pt");
  check("border: 0pt (non-px zero) means no border", nonPxZeroBorder.border === undefined, JSON.stringify(nonPxZeroBorder));

  const nonPxZeroSide = cssToBlockLayout(
    parseInlineStyle("border: 2px solid #1a6fb0; border-right: 0in;"),
  );
  check(
    "border-right: 0in (non-px zero) removes only the right side",
    nonPxZeroSide.borders?.top?.color === "1a6fb0" && nonPxZeroSide.borders?.right === undefined,
    JSON.stringify(nonPxZeroSide.borders),
  );

  const nonPxWidth = parseInlineStyle("border: 2pt solid #1a6fb0");
  check(
    "border: 2pt (non-px, non-zero) converts to px instead of defaulting to 1",
    Math.abs((nonPxWidth.border?.widthPx ?? 0) - (2 * 96) / 72) < 0.001,
    JSON.stringify(nonPxWidth),
  );

  const widthInMiddle = parseInlineStyle("border: solid 2px #1a6fb0");
  check(
    "the width token is found regardless of its position among style/color",
    widthInMiddle.border?.widthPx === 2,
    JSON.stringify(widthInMiddle),
  );

  const rgbaZeroDoesNotFalseTrigger = parseInlineStyle("border: solid rgba(0,0,0,1)");
  check(
    "a 0 color channel inside rgba() doesn't get mistaken for a bare-zero width",
    rgbaZeroDoesNotFalseTrigger.border?.widthPx === 1,
    JSON.stringify(rgbaZeroDoesNotFalseTrigger),
  );

  await writeGuardResult({
    id: "border-width",
    label: "border-width shorthand",
    passed: checksRun - failures,
    total: checksRun,
    ok: failures === 0,
    unit: "pure-function checks",
    command: "npm run guard:border-width",
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
