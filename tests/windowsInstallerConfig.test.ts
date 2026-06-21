import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const packagingSourcePath = fileURLToPath(new URL("../scripts/subcreator-package-windows-exe.mjs", import.meta.url));
const fontInstallerPath = fileURLToPath(new URL("../installers/subcreator_install_windows_fonts.ps1", import.meta.url));
const privateInstallerPath = fileURLToPath(
  new URL("../installers/subcreator_install_windows_private_runtime.ps1", import.meta.url)
);
const batchInstallerPath = fileURLToPath(new URL("../installers/subcreator_install_windows.bat", import.meta.url));

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

describe("Windows bundled font installation", () => {
  it("uses internal font names and content-addressed files without overwriting loaded fonts", () => {
    // // Stable hashed files prevent Adobe from losing a font while Windows still has the previous file mapped.
    const fontInstaller = readFileSync(fontInstallerPath, "utf8");

    expect(fontInstaller).toContain('ExtendedProperty("System.Title")');
    expect(fontInstaller).toContain('targetName = "SubCreator-$safeBaseName-$($sourceHash.Substring(0, 12))');
    expect(fontInstaller).not.toContain("Copy-Item -LiteralPath $sourceFile.FullName -Destination $destination -Force");
  });

  it("loads fonts into the current session and broadcasts the Windows font change", () => {
    // // Registry persistence alone is insufficient for Adobe applications opened in the current logon session.
    const fontInstaller = readFileSync(fontInstallerPath, "utf8");

    expect(fontInstaller).toContain("AddFontResourceEx($destination, 0");
    expect(fontInstaller).toContain("0x001D");
    expect(fontInstaller).toContain("SendMessageTimeout");
  });

  it("routes batch, private-runtime, and EXE installs through the shared font installer", () => {
    // // Every Windows package format must use the same registration and notification behavior.
    const batchInstaller = readFileSync(batchInstallerPath, "utf8");
    const privateInstaller = readFileSync(privateInstallerPath, "utf8");
    const packagingSource = readFileSync(packagingSourcePath, "utf8");

    expect(batchInstaller).toContain("subcreator_install_windows_fonts.ps1");
    expect(privateInstaller).toContain("subcreator_install_windows_fonts.ps1");
    expect(packagingSource).toContain("subcreator_install_windows_fonts.ps1");
  });
});
