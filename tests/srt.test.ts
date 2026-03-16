// // Validate SRT parsing and derived word timing behavior.
import { describe, expect, it } from "vitest";
import { parseSrt } from "../src/core/srt";

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
});
