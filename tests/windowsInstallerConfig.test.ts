import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const packagingSourcePath = fileURLToPath(new URL("../scripts/subcreator-package-windows-exe.mjs", import.meta.url));

describe("Windows installer restart behavior", () => {
  it("suppresses unnecessary computer restart prompts in both generated installers", () => {
    // // The extension and private runtime only require Premiere Pro to restart, not Windows.
    const packagingSource = readFileSync(packagingSourcePath, "utf8");
    const directiveMatches = packagingSource.match(/"RestartIfNeededByRun=no"/g) ?? [];

    expect(directiveMatches).toHaveLength(2);
  });
});
