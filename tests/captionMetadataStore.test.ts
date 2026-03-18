// // Verify persisted caption metadata matching stays robust after safe subtitle track moves.
import { describe, expect, it } from "vitest";
import { findBestCaptionMetadataMatchForItem, type CaptionMetadataClipMatchSource } from "../src/panel/captionMetadataStore";
import type { SelectedMogrtTextItem } from "../src/panel/cepBridge";

describe("captionMetadataStore matching", () => {
  it("matches persisted word timing metadata even after the subtitle clip moved to another video track", () => {
    const clips: CaptionMetadataClipMatchSource[] = [
      {
        trackIndex: 6,
        startSeconds: 10,
        endSeconds: 12,
        text: "hello there",
        words: [
          { text: "hello", startSeconds: 10, endSeconds: 10.8 },
          { text: "there", startSeconds: 10.8, endSeconds: 12 }
        ]
      }
    ];
    const item: SelectedMogrtTextItem = {
      selectionIndex: 0,
      videoTrackIndex: 7,
      startSeconds: 10.01,
      endSeconds: 11.99,
      text: "hello there",
      clipName: "Subtitle"
    };

    const match = findBestCaptionMetadataMatchForItem(clips, item);
    expect(match).not.toBeNull();
    expect(match?.trackIndex).toBe(6);
  });

  it("prefers the same-track metadata match when both same-track and moved-track candidates exist", () => {
    const clips: CaptionMetadataClipMatchSource[] = [
      {
        trackIndex: 6,
        startSeconds: 10,
        endSeconds: 12,
        text: "hello there",
        words: [
          { text: "hello", startSeconds: 10, endSeconds: 10.8 },
          { text: "there", startSeconds: 10.8, endSeconds: 12 }
        ]
      },
      {
        trackIndex: 7,
        startSeconds: 10.02,
        endSeconds: 11.98,
        text: "hello there",
        words: [
          { text: "hello", startSeconds: 10.02, endSeconds: 10.85 },
          { text: "there", startSeconds: 10.85, endSeconds: 11.98 }
        ]
      }
    ];
    const item: SelectedMogrtTextItem = {
      selectionIndex: 0,
      videoTrackIndex: 7,
      startSeconds: 10.01,
      endSeconds: 11.99,
      text: "hello there",
      clipName: "Subtitle"
    };

    const match = findBestCaptionMetadataMatchForItem(clips, item);
    expect(match).not.toBeNull();
    expect(match?.trackIndex).toBe(7);
  });
});
