// // Verify WhisperX display cue safety rules before timeline subtitle creation.
import { describe, expect, it } from "vitest";
import { normalizeWhisperXCuesForDisplay } from "../src/core/whisperxDisplay";
import type { CaptionCue } from "../src/core/types";

function cue(id: string, startSeconds: number, endSeconds: number, text: string): CaptionCue {
  // // Build compact test cues with synthetic word timings matching the cue span.
  const words = text.split(/\s+/).filter(Boolean);
  return {
    id,
    startSeconds,
    endSeconds,
    text,
    words: words.map((word, index) => ({
      text: word,
      startSeconds: startSeconds + ((endSeconds - startSeconds) / Math.max(1, words.length)) * index,
      endSeconds: startSeconds + ((endSeconds - startSeconds) / Math.max(1, words.length)) * (index + 1)
    }))
  };
}

describe("normalizeWhisperXCuesForDisplay", () => {
  it("does not merge two tiny speaker turns when there is a real gap", () => {
    const normalized = normalizeWhisperXCuesForDisplay([cue("a", 1, 1.25, "Oui"), cue("b", 1.37, 1.66, "Non")]);

    expect(normalized).toHaveLength(2);
    expect(normalized.map((item) => item.text)).toEqual(["Oui", "Non"]);
  });

  it("does not merge overlapping tiny groups that may belong to different speakers", () => {
    const normalized = normalizeWhisperXCuesForDisplay([cue("a", 1, 1.35, "Moi"), cue("b", 1.31, 1.62, "Toi")]);

    expect(normalized).toHaveLength(2);
    expect(normalized.map((item) => item.text)).toEqual(["Moi", "Toi"]);
  });

  it("keeps short question replies separate instead of joining separate groups", () => {
    const normalized = normalizeWhisperXCuesForDisplay([cue("a", 2, 2.25, "Name?"), cue("b", 2.33, 2.75, "Elise")]);

    expect(normalized).toHaveLength(2);
    expect(normalized.map((item) => item.text)).toEqual(["Name?", "Elise"]);
  });

  it("still merges near-contiguous micro fragments that are effectively one phrase", () => {
    const normalized = normalizeWhisperXCuesForDisplay([cue("a", 3, 3.24, "je"), cue("b", 3.255, 3.62, "sais pas")]);

    expect(normalized).toHaveLength(1);
    expect(normalized[0].text).toBe("je sais pas");
  });
});
