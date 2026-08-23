// // Keep the translation flow local-key-only and ensure host duplication never removes the source selection.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const projectRoot = resolve(import.meta.dirname, "..");
const bridgeSource = readFileSync(resolve(projectRoot, "src/panel/cepBridge.ts"), "utf8");
const hostSource = readFileSync(resolve(projectRoot, "src/host/SubCreatorHost.jsx"), "utf8");

describe("translation integration", () => {
  it("selects the appropriate DeepL endpoint and sends the user key only as an authorization header", () => {
    expect(bridgeSource).toContain('authKey.endsWith(":fx") ? "api-free.deepl.com" : "api.deepl.com"');
    expect(bridgeSource).toContain("Authorization: `DeepL-Auth-Key ${authKey}`");
    expect(bridgeSource).not.toContain("auth_key:");
  });

  it("duplicates translated MOGRTs above the source without entering the removal path", () => {
    expect(hostSource).toContain("var duplicateSelection = Boolean(payload && payload.duplicateSelection === true);");
    expect(hostSource).toContain("if (!duplicateSelection) {");
    expect(hostSource).toContain("subcreator_get_or_create_video_track_above_index(sequence, targetTrackIndex)");
  });
});
