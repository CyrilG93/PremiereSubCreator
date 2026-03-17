// // Verify text-tab editing helpers for moving, splitting, merging, and retiming subtitle blocks.
import { describe, expect, it } from "vitest";
import {
  buildTextEditorBlocks,
  mergeTextEditorBlocks,
  moveTextEditorWord,
  retimeTextEditorBlocks,
  sanitizeTextEditorBlocksForApply,
  splitTextEditorBlock,
  updateTextEditorBlockText
} from "../src/core/textEditor";

describe("textEditor helpers", () => {
  it("moves one word to the beginning of another subtitle block", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "Je suis alle"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "puis rentre"
      }
    ]);

    const moved = moveTextEditorWord(blocks, 0, 2, 1, 0);
    expect(moved[0].text).toBe("Je suis");
    expect(moved[1].text).toBe("alle puis rentre");
  });

  it("splits one subtitle block at the selected word boundary", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 2,
        text: "bonjour tout le monde"
      }
    ]);

    const split = splitTextEditorBlock(blocks, 0, 2);
    expect(split).toHaveLength(2);
    expect(split[0].text).toBe("bonjour tout");
    expect(split[1].text).toBe("le monde");
  });

  it("merges one subtitle block with the previous block", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "salut"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "les amis"
      }
    ]);

    const merged = mergeTextEditorBlocks(blocks, 1, "previous");
    expect(merged).toHaveLength(1);
    expect(merged[0].text).toBe("salut les amis");
    expect(merged[0].sourceSelectionIndex).toBe(0);
  });

  it("retimes longer edited blocks with more duration", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 2,
        text: "go"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 2,
        endSeconds: 4,
        text: "international subtitle"
      }
    ]);

    const retimed = retimeTextEditorBlocks(blocks);
    const firstDuration = retimed[0].endSeconds - retimed[0].startSeconds;
    const secondDuration = retimed[1].endSeconds - retimed[1].startSeconds;
    expect(secondDuration).toBeGreaterThan(firstDuration);
    expect(retimed[0].startSeconds).toBe(0);
    expect(retimed[1].endSeconds).toBe(4);
  });

  it("drops empty edited blocks before apply and keeps remaining timing span", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 2,
        text: "bonjour"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 2,
        endSeconds: 4,
        text: "tout le monde"
      }
    ]);

    const edited = updateTextEditorBlockText(blocks, 0, "");
    const sanitized = sanitizeTextEditorBlocksForApply(edited);
    expect(sanitized).toHaveLength(1);
    expect(sanitized[0].text).toBe("tout le monde");
    expect(sanitized[0].startSeconds).toBe(0);
    expect(sanitized[0].endSeconds).toBe(4);
  });
});
