// // Verify parsing, prompting, exact spelling corrections, and timing preservation for the global Whisper dictionary.
import { describe, expect, it } from "vitest";
import {
  applyWhisperGlossaryToCues,
  buildWhisperGlossaryPrompt,
  parseWhisperGlossary
} from "../src/core/whisperGlossary";

describe("Whisper global dictionary", () => {
  it("parses canonical spellings, aliases, comments, and legacy comma-separated terms", () => {
    const entries = parseWhisperGlossary(
      "# Marques\nAdobe Premiere Pro\nserial plugin | cyril plug-in => Cyril Plugin\nWhisperX, MOGRT"
    );

    expect(entries).toEqual([
      { canonical: "Adobe Premiere Pro", variants: ["Adobe Premiere Pro"] },
      { canonical: "Cyril Plugin", variants: ["Cyril Plugin", "serial plugin", "cyril plug-in"] },
      { canonical: "WhisperX", variants: ["WhisperX"] },
      { canonical: "MOGRT", variants: ["MOGRT"] }
    ]);
  });

  it("builds a compact prompt from complete canonical spellings only", () => {
    expect(buildWhisperGlossaryPrompt("adobe premiere => Adobe Premiere Pro\nMOGRT", 23)).toBe("Adobe Premiere Pro");
    expect(buildWhisperGlossaryPrompt("", 420)).toBe("");
  });

  it("corrects multi-word aliases and canonical capitalization while preserving timing spans", () => {
    const corrected = applyWhisperGlossaryToCues(
      [
        {
          id: "whisper-1",
          startSeconds: 0,
          endSeconds: 2,
          text: "Utilise adobe premiere et whisperx.",
          words: [
            { text: "Utilise", startSeconds: 0, endSeconds: 0.4 },
            { text: "adobe", startSeconds: 0.4, endSeconds: 0.8 },
            { text: "premiere", startSeconds: 0.8, endSeconds: 1.2 },
            { text: "et", startSeconds: 1.2, endSeconds: 1.5 },
            { text: "whisperx.", startSeconds: 1.5, endSeconds: 2 }
          ]
        }
      ],
      "adobe premiere => Adobe Premiere Pro\nWhisperX"
    );

    expect(corrected[0].text).toBe("Utilise Adobe Premiere Pro et WhisperX.");
    expect(corrected[0].words.map((word) => word.text)).toEqual(["Utilise", "Adobe", "Premiere", "Pro", "et", "WhisperX."]);
    expect(corrected[0].words[1].startSeconds).toBe(0.4);
    expect(corrected[0].words[3].endSeconds).toBe(1.2);
    expect(corrected[0].words[5].startSeconds).toBe(1.5);
    expect(corrected[0].words[5].endSeconds).toBe(2);
  });

  it("does not replace fragments inside unrelated words", () => {
    const corrected = applyWhisperGlossaryToCues(
      [
        {
          id: "whisper-1",
          startSeconds: 0,
          endSeconds: 1,
          text: "premièrement",
          words: [{ text: "premièrement", startSeconds: 0, endSeconds: 1 }]
        }
      ],
      "première => Premiere"
    );

    expect(corrected[0].text).toBe("premièrement");
  });
});
