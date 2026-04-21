// // Validate Whisper JSON parsing and fallback cue construction.
import { describe, expect, it } from "vitest";
import { parseWhisperJson } from "../src/core/whisper";

describe("parseWhisperJson", () => {
  it("preserves precise word timestamps from Whisper json", () => {
    const cues = parseWhisperJson(
      JSON.stringify({
        segments: [
          {
            id: 0,
            start: 0,
            end: 2.2,
            text: " Bonjour tout le monde",
            words: [
              { word: " Bonjour", start: 0, end: 0.5 },
              { word: " tout", start: 0.5, end: 0.9 },
              { word: " le", start: 0.9, end: 1.1 },
              { word: " monde", start: 1.1, end: 2.2 }
            ]
          }
        ]
      })
    );

    expect(cues).toHaveLength(1);
    expect(cues[0].text).toBe("Bonjour tout le monde");
    expect(cues[0].words).toHaveLength(4);
    expect(cues[0].words[0].startSeconds).toBe(0);
    expect(cues[0].words[3].endSeconds).toBe(2.2);
  });

  it("falls back to weighted synthetic words when Whisper json has no word array", () => {
    const cues = parseWhisperJson(
      JSON.stringify({
        segments: [{ id: 0, start: 1, end: 4, text: "go international" }]
      })
    );

    expect(cues).toHaveLength(1);
    expect(cues[0].words).toHaveLength(2);

    const shortWordDuration = cues[0].words[0].endSeconds - cues[0].words[0].startSeconds;
    const longWordDuration = cues[0].words[1].endSeconds - cues[0].words[1].startSeconds;
    expect(longWordDuration).toBeGreaterThan(shortWordDuration);
  });

  it("keeps split hyphen Whisper tokens available for later normalization", () => {
    const cues = parseWhisperJson(
      JSON.stringify({
        segments: [
          {
            id: 0,
            start: 0,
            end: 1.2,
            text: "peut -etre",
            words: [
              { word: "peut", start: 0, end: 0.3 },
              { word: "-", start: 0.3, end: 0.35 },
              { word: "etre", start: 0.35, end: 1.2 }
            ]
          }
        ]
      })
    );

    expect(cues).toHaveLength(1);
    expect(cues[0].text).toBe("peut -etre");
    expect(cues[0].words.map((word) => word.text)).toEqual(["peut", "-", "etre"]);
  });
});
