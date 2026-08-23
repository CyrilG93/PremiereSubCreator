// // Keep the translation flow local-key-only and ensure host duplication never removes the source selection.
import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const projectRoot = resolve(import.meta.dirname, "..");
const bridgeSource = readFileSync(resolve(projectRoot, "src/panel/cepBridge.ts"), "utf8");
const hostSource = readFileSync(resolve(projectRoot, "src/host/SubCreatorHost.jsx"), "utf8");
const panelSource = readFileSync(resolve(projectRoot, "src/panel/main.ts"), "utf8");

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

  it("keeps correction inputs and native SRT output tied to existing cue timings", () => {
    expect(panelSource).toContain('input.addEventListener("input"');
    expect(panelSource).toContain('getTranslationInputMode() === "srt"');
    expect(panelSource).toContain("applyNativeSubtitlePlan({");
  });

  it("persists the DeepL key locally without adding it to the translation request payload", () => {
    expect(panelSource).toContain("deeplApiKey?: string;");
    expect(panelSource).toContain("deeplApiKey: String(elements.deeplApiKey?.value || \"\")");
    expect(panelSource).toContain('elements.deeplApiKey?.addEventListener("input"');
  });

  it("loads plan-specific DeepL source and target language lists with the user's key", () => {
    expect(bridgeSource).toContain('path: `/v2/languages?type=${type}`');
    expect(bridgeSource).toContain('Promise.all([readType("source"), readType("target")])');
    expect(panelSource).toContain("refreshDeepLSupportedLanguages()");
    expect(panelSource).toContain("replaceTranslationLanguageOptions(elements.translationSourceLanguage");
  });

  it("maps DeepL failures to stable user-facing error codes", () => {
    expect(bridgeSource).toContain('return "SUBCREATOR_DEEPL_API_KEY_INVALID";');
    expect(bridgeSource).toContain('return "SUBCREATOR_DEEPL_QUOTA_EXCEEDED";');
    expect(bridgeSource).toContain('return "SUBCREATOR_DEEPL_RATE_LIMITED";');
    expect(panelSource).toContain("getDeepLUserErrorMessage(error)");
  });
});
