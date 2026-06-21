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

  it("requires the private runtime version marker before a Light installer reuses it", () => {
    // // A few executable names are not enough to prove that Whisper imports still work.
    const packagingSource = readFileSync(packagingSourcePath, "utf8");

    expect(packagingSource).toContain('"  Result := FileExists(VersionFile) and"');
    expect(packagingSource).not.toContain("Accept the compatible runtime installed by the previous all-in-one EXE");
  });

  it("embeds a verified base model in the Full installer", () => {
    // // A first offline install must include one usable transcription model as well as Python and FFmpeg.
    const packagingSource = readFileSync(packagingSourcePath, "utf8");

    expect(packagingSource).toContain("async function prepareBundledBaseModel()");
    expect(packagingSource).toContain("Bundled base model SHA-256 mismatch");
    expect(packagingSource).toContain('model.id === "base" && bundledBaseModelPath');
  });

  it("fails Setup when a runtime or dependency validation process fails", () => {
    // // Inno's [Run] section ignores child exit codes, so the generated Pascal code must enforce them.
    const packagingSource = readFileSync(packagingSourcePath, "utf8");

    expect(packagingSource).not.toContain('"[Run]"');
    expect(packagingSource).toContain("AfterInstall: InstallSubCreator");
    expect(packagingSource).toContain("if ResultCode <> 0 then");
    expect(packagingSource).toContain("RaiseException(Format('Sub Creator dependency validation failed");
  });
});
