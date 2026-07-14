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

  it("does not send a generic mixed-language prompt when no glossary terms are provided", () => {
    // // A generic prompt without concrete terms caused missing subtitle segments in Hindi/English tests.
    const source = readFileSync(panelSourcePath, "utf8").replace(/\r\n/g, "\n");

    expect(source).toContain("if (!glossary) {\n    return \"\";\n  }");
    expect(source).not.toContain("Keep English words in English Latin spelling");
  });

  it("keeps the experimental mixed-language UI disabled until it is reliable enough to expose again", () => {
    // // Older localStorage values must not be able to send a Preserve prompt while the feature flag is disabled.
    const source = readFileSync(panelSourcePath, "utf8");

    expect(source).toContain("const MIXED_LANGUAGE_FEATURE_ENABLED = false;");
    expect(source).toContain("if (!MIXED_LANGUAGE_FEATURE_ENABLED || !options.preserveMixedLanguages)");
    expect(source).toContain("preserveMixedLanguages: MIXED_LANGUAGE_FEATURE_ENABLED && Boolean(elements.preserveMixedLanguages");
  });
});
