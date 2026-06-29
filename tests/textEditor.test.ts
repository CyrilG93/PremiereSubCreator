// // Verify text-tab editing helpers for moving, splitting, merging, and retiming subtitle blocks.
import { describe, expect, it } from "vitest";
import {
  buildTextEditorApplyPlan,
  buildTextEditorApplyPlans,
  buildTextEditorSafeApplyPlans,
  buildTextEditorBlocks,
  mergeTextEditorBlocks,
  moveTextEditorWord,
  prepareTextEditorBlocksForApply,
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

  it("normalizes apostrophes and punctuation when building editor blocks", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "Salut ! c 'est d 'accord ?"
      }
    ]);

    expect(blocks[0].text).toBe("Salut! c'est d'accord?");
    expect(blocks[0].words).toEqual(["Salut!", "c'est", "d'accord?"]);
  });

  it("preserves timed words when moving one word across subtitle blocks", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one two",
        timedWords: [
          { text: "one", startSeconds: 0, endSeconds: 0.4 },
          { text: "two", startSeconds: 0.4, endSeconds: 1 }
        ]
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "three four",
        timedWords: [
          { text: "three", startSeconds: 1, endSeconds: 1.5 },
          { text: "four", startSeconds: 1.5, endSeconds: 2 }
        ]
      }
    ]);

    const moved = moveTextEditorWord(blocks, 0, 1, 1, 0);
    expect(moved[0].timedWords?.map((word) => word.text)).toEqual(["one"]);
    expect(moved[1].timedWords?.map((word) => word.text)).toEqual(["two", "three", "four"]);
    expect(moved[1].timedWords?.[0]?.startSeconds).toBe(0.4);
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

  it("preserves timed words when splitting one subtitle block", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 2,
        text: "bonjour tout le monde",
        timedWords: [
          { text: "bonjour", startSeconds: 0, endSeconds: 0.4 },
          { text: "tout", startSeconds: 0.4, endSeconds: 0.8 },
          { text: "le", startSeconds: 0.8, endSeconds: 1.2 },
          { text: "monde", startSeconds: 1.2, endSeconds: 2 }
        ]
      }
    ]);

    const split = splitTextEditorBlock(blocks, 0, 2);
    expect(split[0].timedWords?.map((word) => word.text)).toEqual(["bonjour", "tout"]);
    expect(split[1].timedWords?.map((word) => word.text)).toEqual(["le", "monde"]);
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

  it("preserves timed words when merging subtitle blocks", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "salut",
        timedWords: [{ text: "salut", startSeconds: 0, endSeconds: 1 }]
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "les amis",
        timedWords: [
          { text: "les", startSeconds: 1, endSeconds: 1.5 },
          { text: "amis", startSeconds: 1.5, endSeconds: 2 }
        ]
      }
    ]);

    const merged = mergeTextEditorBlocks(blocks, 1, "previous");
    expect(merged[0].timedWords?.map((word) => word.text)).toEqual(["salut", "les", "amis"]);
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

  it("reuses timed words instead of weighted timing when precise metadata is available", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1.1,
        text: "go now",
        timedWords: [
          { text: "go", startSeconds: 0, endSeconds: 0.2 },
          { text: "now", startSeconds: 0.2, endSeconds: 1.1 }
        ]
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1.1,
        endSeconds: 3,
        text: "international subtitle",
        timedWords: [
          { text: "international", startSeconds: 1.1, endSeconds: 2.6 },
          { text: "subtitle", startSeconds: 2.6, endSeconds: 3 }
        ]
      }
    ]);

    const retimed = retimeTextEditorBlocks(blocks, {
      startSeconds: 0,
      endSeconds: 3
    });

    expect(retimed[0].startSeconds).toBe(0);
    expect(retimed[0].endSeconds).toBe(1.1);
    expect(retimed[1].startSeconds).toBe(1.1);
    expect(retimed[1].endSeconds).toBe(3);
  });

  it("keeps compatible timed-word blocks precise while retiming only the edited gap blocks", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1.1,
        text: "go now",
        timedWords: [
          { text: "go", startSeconds: 0, endSeconds: 0.2 },
          { text: "now", startSeconds: 0.2, endSeconds: 1.1 }
        ]
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1.1,
        endSeconds: 2,
        text: "edited words"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "final block",
        timedWords: [
          { text: "final", startSeconds: 2, endSeconds: 2.6 },
          { text: "block", startSeconds: 2.6, endSeconds: 3 }
        ]
      }
    ]);

    const retimed = retimeTextEditorBlocks(blocks, {
      startSeconds: 0,
      endSeconds: 3
    });

    expect(retimed[0].startSeconds).toBe(0);
    expect(retimed[0].endSeconds).toBe(1.1);
    expect(retimed[1].startSeconds).toBe(1.1);
    expect(retimed[1].endSeconds).toBe(2);
    expect(retimed[2].startSeconds).toBe(2);
    expect(retimed[2].endSeconds).toBe(3);
    expect(retimed[2].timedWords?.map((word) => word.text)).toEqual(["final", "block"]);
  });

  it("keeps the original selected span when a merge removes the last block", () => {
    const blocks = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 10,
        endSeconds: 11,
        text: "one"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 11,
        endSeconds: 12,
        text: "two"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 12,
        endSeconds: 14,
        text: "three four"
      }
    ]);

    const merged = mergeTextEditorBlocks(blocks, 2, "previous");
    const retimed = retimeTextEditorBlocks(merged, {
      startSeconds: 10,
      endSeconds: 14
    });

    expect(retimed).toHaveLength(2);
    expect(retimed[0].startSeconds).toBe(10);
    expect(retimed[1].endSeconds).toBe(14);
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

  it("keeps the original selected span when the last block becomes empty before apply", () => {
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
        endSeconds: 5,
        text: "tout"
      }
    ]);

    const edited = updateTextEditorBlockText(blocks, 1, "");
    const sanitized = sanitizeTextEditorBlocksForApply(edited, {
      startSeconds: 0,
      endSeconds: 5
    });

    expect(sanitized).toHaveLength(1);
    expect(sanitized[0].text).toBe("bonjour");
    expect(sanitized[0].startSeconds).toBe(0);
    expect(sanitized[0].endSeconds).toBe(5);
  });

  it("builds an apply plan only for the changed middle range", () => {
    const original = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "two"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 3,
        endSeconds: 4,
        text: "four"
      }
    ]);

    const edited = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "two updated"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three updated"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 3,
        endSeconds: 4,
        text: "four"
      }
    ]);

    const plan = buildTextEditorApplyPlan(original, edited);
    expect(plan).not.toBeNull();
    expect(plan?.selectionStartIndex).toBe(1);
    expect(plan?.selectionEndIndex).toBe(2);
    expect(plan?.blocks).toHaveLength(2);
    expect(plan?.blocks[0].text).toBe("two updated");
    expect(plan?.blocks[1].text).toBe("three updated");
    expect(plan?.timingRange.startSeconds).toBe(1);
    expect(plan?.timingRange.endSeconds).toBe(3);
  });

  it("builds separate apply plans for disjoint changed ranges", () => {
    const original = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "two"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 3,
        endSeconds: 4,
        text: "four"
      }
    ]);

    const edited = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one updated"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "two"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 3,
        endSeconds: 4,
        text: "four updated"
      }
    ]);

    const plans = buildTextEditorApplyPlans(original, edited);
    expect(plans).toHaveLength(2);
    expect(plans[0]?.selectionStartIndex).toBe(0);
    expect(plans[0]?.selectionEndIndex).toBe(0);
    expect(plans[0]?.blocks.map((block) => block.text)).toEqual(["one updated"]);
    expect(plans[1]?.selectionStartIndex).toBe(3);
    expect(plans[1]?.selectionEndIndex).toBe(3);
    expect(plans[1]?.blocks.map((block) => block.text)).toEqual(["four updated"]);
  });

  it("maps apply plan indexes back to real Premiere selection indexes after filtered non-subtitle entries", () => {
    const original = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three"
      }
    ]);

    const edited = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one updated"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three"
      }
    ]);

    const plan = buildTextEditorApplyPlan(original, edited);
    expect(plan).not.toBeNull();
    expect(plan?.selectionStartIndex).toBe(0);
    expect(plan?.selectionEndIndex).toBe(0);
  });

  it("keeps unchanged middle blocks out of a split disjoint apply plan", () => {
    const original = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "alpha"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 3,
        text: "beta gamma"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 3,
        endSeconds: 4,
        text: "delta"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 4,
        endSeconds: 5,
        text: "epsilon"
      }
    ]);

    const edited = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "alpha"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "beta"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 2,
        endSeconds: 3,
        text: "gamma"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 3,
        endSeconds: 4,
        text: "delta"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 4,
        endSeconds: 5,
        text: "epsilon updated"
      }
    ]);

    const plans = buildTextEditorApplyPlans(original, edited);
    expect(plans).toHaveLength(2);
    expect(plans[0]?.selectionStartIndex).toBe(1);
    expect(plans[0]?.selectionEndIndex).toBe(1);
    expect(plans[0]?.blocks.map((block) => block.text)).toEqual(["beta", "gamma"]);
    expect(plans[1]?.selectionStartIndex).toBe(3);
    expect(plans[1]?.selectionEndIndex).toBe(3);
    expect(plans[1]?.blocks.map((block) => block.text)).toEqual(["epsilon updated"]);
  });

  it("collapses disjoint changed ranges into one safe combined apply span", () => {
    const original = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "two"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 3,
        endSeconds: 4,
        text: "four"
      }
    ]);

    const edited = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "one updated"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "two"
      },
      {
        sourceSelectionIndex: 2,
        clipName: "Clip C",
        startSeconds: 2,
        endSeconds: 3,
        text: "three"
      },
      {
        sourceSelectionIndex: 3,
        clipName: "Clip D",
        startSeconds: 3,
        endSeconds: 4,
        text: "four updated"
      }
    ]);

    const plans = buildTextEditorSafeApplyPlans(original, edited);
    expect(plans).toHaveLength(1);
    expect(plans[0]?.selectionStartIndex).toBe(0);
    expect(plans[0]?.selectionEndIndex).toBe(3);
    expect(plans[0]?.blocks.map((block) => block.text)).toEqual(["one updated", "two", "three", "four updated"]);
  });

  it("keeps one merged Text editor block even when it exceeds Creation word limits", () => {
    const original = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 2,
        text: "one two three"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 2,
        endSeconds: 4,
        text: "four five six seven"
      }
    ]);
    const merged = mergeTextEditorBlocks(original, 1, "previous");
    const plan = buildTextEditorApplyPlan(original, merged);

    expect(plan).not.toBeNull();
    const preparedBlocks = prepareTextEditorBlocksForApply(plan?.blocks || []);
    expect(preparedBlocks).toHaveLength(1);
    expect(preparedBlocks[0].text).toBe("one two three four five six seven");
    expect(preparedBlocks[0].startSeconds).toBe(0);
    expect(preparedBlocks[0].endSeconds).toBe(4);
  });

  it("returns no apply plan when nothing changed", () => {
    const original = buildTextEditorBlocks([
      {
        sourceSelectionIndex: 0,
        clipName: "Clip A",
        startSeconds: 0,
        endSeconds: 1,
        text: "same"
      },
      {
        sourceSelectionIndex: 1,
        clipName: "Clip B",
        startSeconds: 1,
        endSeconds: 2,
        text: "content"
      }
    ]);

    expect(buildTextEditorApplyPlan(original, original)).toBeNull();
  });
});
