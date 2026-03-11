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
  mogrtPath: "",
  mogrtTemplateRelativePath: "",
  captionSourcePath: "",
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
});
