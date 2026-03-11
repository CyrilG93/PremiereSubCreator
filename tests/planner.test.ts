// // Verify text wrapping and uppercase behavior in caption planning.
import { describe, expect, it } from "vitest";
import { buildCaptionPlan } from "../src/core/planner";
import type { CaptionBuildOptions } from "../src/core/types";

const baseOptions: CaptionBuildOptions = {
  sourceMode: "srt",
  languageCode: "fr",
  style: {
    presetId: "clean",
    fontSize: 78,
    maxCharsPerLine: 12,
    animationMode: "line",
    uppercase: false,
    linesPerCaption: 2
  },
  extensionRootPath: "",
  mogrtPath: "",
  mogrtTemplateRelativePath: "",
  whisperAudioPath: "",
  whisperModel: "base",
  videoTrackIndex: 0,
  audioTrackIndex: 0
};

describe("buildCaptionPlan", () => {
  it("splits long cues into multiple parts", () => {
    const planned = buildCaptionPlan(
      [
        {
          id: "cue-1",
          startSeconds: 0,
          endSeconds: 6,
          text: "bonjour comment vas tu aujourd hui",
          words: []
        }
      ],
      baseOptions
    );

    expect(planned.length).toBeGreaterThan(1);
    expect(planned[0].startSeconds).toBe(0);
    expect(planned[planned.length - 1].endSeconds).toBe(6);
    expect(planned.every((cue) => cue.text.split("\n").length <= baseOptions.style.linesPerCaption)).toBe(true);
  });

  it("applies uppercase when enabled", () => {
    const planned = buildCaptionPlan(
      [
        {
          id: "cue-2",
          startSeconds: 0,
          endSeconds: 1,
          text: "mini phrase",
          words: []
        }
      ],
      {
        ...baseOptions,
        style: {
          ...baseOptions.style,
          uppercase: true
        }
      }
    );

    expect(planned[0].text).toBe("MINI PHRASE");
  });

  it("keeps chunk timing aligned to contiguous word timing", () => {
    const planned = buildCaptionPlan(
      [
        {
          id: "cue-3",
          startSeconds: 0,
          endSeconds: 4,
          text: "alpha beta gamma delta",
          words: [
            { text: "alpha", startSeconds: 0, endSeconds: 1 },
            { text: "beta", startSeconds: 1, endSeconds: 2 },
            { text: "gamma", startSeconds: 2, endSeconds: 3 },
            { text: "delta", startSeconds: 3, endSeconds: 4 }
          ]
        }
      ],
      {
        ...baseOptions,
        style: {
          ...baseOptions.style,
          maxCharsPerLine: 10,
          linesPerCaption: 1
        }
      }
    );

    expect(planned.length).toBeGreaterThan(1);
    expect(planned[0].startSeconds).toBe(0);
    expect(planned[0].endSeconds).toBe(2);
    expect(planned[planned.length - 1].endSeconds).toBe(4);
  });
});
