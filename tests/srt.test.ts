// // Validate SRT parsing and derived word timing behavior.
import { describe, expect, it } from "vitest";
import { parseSrt, serializeSrt, shiftCaptionCues, trimSrtCuesToRange } from "../src/core/srt";

describe("parseSrt", () => {
  it("parses numbered blocks", () => {
    const cues = parseSrt(`1\n00:00:01,000 --> 00:00:03,000\nBonjour tout le monde\n\n2\n00:00:03,500 --> 00:00:05,000\nDeuxieme ligne`);

    expect(cues).toHaveLength(2);
    expect(cues[0].text).toBe("Bonjour tout le monde");
    expect(cues[0].startSeconds).toBe(1);
    expect(cues[0].endSeconds).toBe(3);
    expect(cues[0].words).toHaveLength(4);
  });

  it("supports non-indexed cues", () => {
    const cues = parseSrt(`00:00:01,000 --> 00:00:02,000\nHello there`);
    expect(cues).toHaveLength(1);
    expect(cues[0].text).toBe("Hello there");
  });

  it("assigns more synthetic duration to longer words", () => {
    const cues = parseSrt(`1\n00:00:00,000 --> 00:00:04,000\ngo international`);

    expect(cues).toHaveLength(1);
    expect(cues[0].words).toHaveLength(2);

    const shortWordDuration = cues[0].words[0].endSeconds - cues[0].words[0].startSeconds;
    const longWordDuration = cues[0].words[1].endSeconds - cues[0].words[1].startSeconds;
    expect(longWordDuration).toBeGreaterThan(shortWordDuration);
  });

  it("keeps hyphenated words together for synthetic timing", () => {
    const cues = parseSrt(`1\n00:00:00,000 --> 00:00:02,000\npeut -etre`);

    expect(cues).toHaveLength(1);
    expect(cues[0].words).toEqual([{ text: "peut-etre", startSeconds: 0, endSeconds: 2 }]);
  });

  it("trims SRT cues to a requested range", () => {
    const cues = parseSrt(
      `1\n00:00:01,000 --> 00:00:03,000\nAlpha\n\n2\n00:00:04,000 --> 00:00:06,000\nBeta\n\n3\n00:00:07,000 --> 00:00:09,000\nGamma`
    );

    const trimmed = trimSrtCuesToRange(cues, 2, 7.5);

    expect(trimmed).toHaveLength(3);
    expect(trimmed[0].startSeconds).toBe(2);
    expect(trimmed[0].endSeconds).toBe(3);
    expect(trimmed[2].startSeconds).toBe(7);
    expect(trimmed[2].endSeconds).toBe(7.5);
  });

  it("shifts cue and word timings by a fixed offset", () => {
    const cues = parseSrt(`1\n00:00:00,000 --> 00:00:02,000\nHello there`);
    const shifted = shiftCaptionCues(cues, 15.25);

    expect(shifted[0].startSeconds).toBe(15.25);
    expect(shifted[0].endSeconds).toBe(17.25);
    expect(shifted[0].words[0].startSeconds).toBe(15.25);
  });

  it("serializes cues back into SRT text", () => {
    const cues = parseSrt(`1\n00:00:01,250 --> 00:00:03,500\nSerialized cue`);
    const serialized = serializeSrt(cues);

    expect(serialized).toContain("00:00:01,250 --> 00:00:03,500");
    expect(serialized).toContain("Serialized cue");
  });
});
