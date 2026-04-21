// // Verify subtitle-token normalization shared by generation and text editing workflows.
import { describe, expect, it } from "vitest";
import { normalizeCaptionWords, normalizeInlineSubtitleText, tokenizeSubtitleText } from "../src/core/textNormalization";

describe("textNormalization", () => {
  it("re-attaches apostrophes and terminal punctuation in plain text", () => {
    expect(normalizeInlineSubtitleText("c 'est d 'accord ! mais pourquoi ?")).toBe("c'est d'accord! mais pourquoi?");
    expect(tokenizeSubtitleText("Salut ! c 'est bon ?")).toEqual(["Salut!", "c'est", "bon?"]);
  });

  it("re-attaches hyphenated words when spacing is broken", () => {
    expect(normalizeInlineSubtitleText("peut -etre fin - de - phrase")).toBe("peut-etre fin-de-phrase");
    expect(tokenizeSubtitleText("peut -etre fin - de - phrase")).toEqual(["peut-etre", "fin-de-phrase"]);
  });

  it("merges timed apostrophe and punctuation tokens while preserving timing spans", () => {
    const normalized = normalizeCaptionWords([
      { text: "c", startSeconds: 0, endSeconds: 0.1 },
      { text: "'", startSeconds: 0.1, endSeconds: 0.15 },
      { text: "est", startSeconds: 0.15, endSeconds: 0.4 },
      { text: "vrai", startSeconds: 0.4, endSeconds: 0.7 },
      { text: "!", startSeconds: 0.7, endSeconds: 0.8 }
    ]);

    expect(normalized).toEqual([
      { text: "c'est", startSeconds: 0, endSeconds: 0.4 },
      { text: "vrai!", startSeconds: 0.4, endSeconds: 0.8 }
    ]);
  });

  it("merges timed hyphen fragments while preserving timing spans", () => {
    const normalized = normalizeCaptionWords([
      { text: "peut", startSeconds: 0, endSeconds: 0.2 },
      { text: "-", startSeconds: 0.2, endSeconds: 0.25 },
      { text: "etre", startSeconds: 0.25, endSeconds: 0.7 },
      { text: "semi", startSeconds: 0.7, endSeconds: 0.9 },
      { text: "-transparent", startSeconds: 0.9, endSeconds: 1.2 }
    ]);

    expect(normalized).toEqual([
      { text: "peut-etre", startSeconds: 0, endSeconds: 0.7 },
      { text: "semi-transparent", startSeconds: 0.7, endSeconds: 1.2 }
    ]);
  });
});
