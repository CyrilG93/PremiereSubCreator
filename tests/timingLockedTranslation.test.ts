// // Verify that forced-language text cannot alter the source subtitle boundaries.
import { describe, expect, it } from "vitest";
import { lockTranslatedCuesToSourceTiming } from "../src/core/timingLockedTranslation";

describe("lockTranslatedCuesToSourceTiming", () => {
  it("keeps every source cue boundary while replacing its text", () => {
    const result = lockTranslatedCuesToSourceTiming(
      [
        { id: "source-1", startSeconds: 1, endSeconds: 2, text: "One two", words: [] },
        { id: "source-2", startSeconds: 2, endSeconds: 4, text: "Three four", words: [] }
      ],
      [{ id: "translated-1", startSeconds: 0, endSeconds: 8, text: "Un deux trois quatre cinq six", words: [] }]
    );

    expect(result.map((cue) => [cue.startSeconds, cue.endSeconds])).toEqual([
      [1, 2],
      [2, 4]
    ]);
    expect(result.map((cue) => cue.text).join(" ")).toBe("Un deux trois quatre cinq six");
    expect(result.flatMap((cue) => cue.words).every((word) => word.startSeconds >= 1 && word.endSeconds <= 4)).toBe(true);
  });
});
