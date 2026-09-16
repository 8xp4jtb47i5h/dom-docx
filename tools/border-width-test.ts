/**
 * `border-width` guard — the CSS `border-width` / `border-{side}-width`
 * properties weren't handled by `parseInlineStyle` at all, so they fell
 * through to the `default:` case and left `borderTop`/etc. undefined
 * (no border emitted, even though the author declared one explicitly).
 */
import { parseInlineStyle } from "../src/converter/css.js";
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
