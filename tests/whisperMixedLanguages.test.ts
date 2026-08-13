import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const cepBridgeSourcePath = fileURLToPath(new URL("../src/panel/cepBridge.ts", import.meta.url));
const hostSourcePath = fileURLToPath(new URL("../src/host/SubCreatorHost.jsx", import.meta.url));
const panelSourcePath = fileURLToPath(new URL("../src/panel/main.ts", import.meta.url));
const pythonHelperSourcePath = fileURLToPath(new URL("../src/python/subcreator_align_corrected.py", import.meta.url));

describe("Whisper mixed-language preservation", () => {
  it("forces transcription mode and forwards initial prompts through the CEP Node path", () => {
    // // Keep mixed-language mode from accidentally falling back to Whisper translation defaults or losing its prompt.
    const source = readFileSync(cepBridgeSourcePath, "utf8");

    expect(source).toContain('"--task",');
    expect(source).toContain('"transcribe",');
    expect(source).toContain('"--initial_prompt", initialPrompt');
    expect(source).toContain('"--initial-prompt"');
  });

  it("keeps the host fallback aligned with the CEP Node Whisper arguments", () => {
    // // The ExtendScript fallback should behave the same way on hosts where CEP Node cannot launch Whisper.
    const source = readFileSync(hostSourcePath, "utf8");

    expect(source).toContain('" --task transcribe"');
    expect(source).toContain('" --initial_prompt "');
    expect(source).toContain("payload.initialPrompt");
  });

  it("passes the prompt into the WhisperX transcription seed", () => {
    // // WhisperX transcription uses the Python helper, so it needs the same mixed-language prompt path.
    const source = readFileSync(pythonHelperSourcePath, "utf8");

    expect(source).toContain('"task": "transcribe"');
    expect(source).toContain('"initial_prompt"');
    expect(source).toContain('"--initial-prompt"');
  });

  it("does not send a generic prompt when no dictionary terms are provided", () => {
    // // An empty dictionary must not add context that could destabilize multilingual transcription.
    const source = readFileSync(panelSourcePath, "utf8").replace(/\r\n/g, "\n");

    expect(source).toContain("return buildWhisperGlossaryPrompt(sanitizeWhisperGlossary(options.mixedLanguagePrompt));");
    expect(source).not.toContain("Keep English words in English Latin spelling");
  });

  it("exposes the global dictionary and applies it after Whisper transcription", () => {
    // // The dictionary combines prompt guidance with deterministic cue correction for exact spellings.
    const source = readFileSync(panelSourcePath, "utf8");

    expect(source).not.toContain("MIXED_LANGUAGE_FEATURE_ENABLED");
    expect(source).toContain("applyWhisperGlossaryToCues(fallbackCues, options.mixedLanguagePrompt)");
    expect(source).toContain("readWhisperGlossaryStore()");
    expect(source).toContain("writeWhisperGlossaryStore(");
  });

  it("stores the dictionary outside the extension for cross-project persistence", () => {
    // // Installer updates must not erase the user dictionary on either supported platform.
    const source = readFileSync(cepBridgeSourcePath, "utf8");

    expect(source).toContain('modules.path.join(appData, "SubCreator", "glossary.json")');
    expect(source).toContain('"Library", "Application Support", "SubCreator", "glossary.json"');
    expect(source).toContain("export async function readWhisperGlossaryStore");
    expect(source).toContain("export async function writeWhisperGlossaryStore");
  });
});
