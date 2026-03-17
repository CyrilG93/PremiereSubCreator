// // Provide pure text-editing helpers for the Text tab subtitle workflow.
import { buildWeightedCaptionWords } from "./wordTiming";

// // Describe one editable subtitle block shown in the Text tab.
export interface TextEditorBlock {
  sourceSelectionIndex: number;
  clipName: string;
  startSeconds: number;
  endSeconds: number;
  text: string;
  words: string[];
}

function normalizeTextEditorWords(text: string): string[] {
  // // Split one subtitle text into compact word tokens while preserving punctuation on each token.
  return String(text || "")
    .replace(/\r/g, "\n")
    .replace(/\n+/g, " ")
    .split(/\s+/)
    .map((value) => value.trim())
    .filter(Boolean);
}

function syncTextEditorBlock(block: TextEditorBlock): TextEditorBlock {
  // // Keep `text` and `words` synchronized after one editing operation.
  const words = Array.isArray(block.words) && block.words.length > 0 ? block.words.slice() : normalizeTextEditorWords(block.text);
  return {
    sourceSelectionIndex: Number.isFinite(Number(block.sourceSelectionIndex)) ? Number(block.sourceSelectionIndex) : 0,
    clipName: String(block.clipName || "").trim(),
    startSeconds: Number(block.startSeconds || 0),
    endSeconds: Number(block.endSeconds || 0),
    text: words.join(" "),
    words
  };
}

function cloneTextEditorBlocks(blocks: TextEditorBlock[]): TextEditorBlock[] {
  // // Clone editable blocks so UI reducers can stay immutable and predictable.
  return blocks.map((block) => syncTextEditorBlock(block));
}

function filterNonEmptyTextEditorBlocks(blocks: TextEditorBlock[]): TextEditorBlock[] {
  // // Remove empty subtitle blocks once an explicit move/merge/apply operation has consumed their content.
  return cloneTextEditorBlocks(blocks).filter((block) => block.words.length > 0);
}

function clampInsertIndex(targetLength: number, targetWordIndex?: number): number {
  // // Normalize one insertion index so drag-and-drop can safely prepend, append, or insert before one chip.
  if (!Number.isFinite(Number(targetWordIndex))) {
    return targetLength;
  }

  return Math.max(0, Math.min(targetLength, Math.floor(Number(targetWordIndex))));
}

export function buildTextEditorBlocks(items: Array<Omit<TextEditorBlock, "words"> & { words?: string[] }>): TextEditorBlock[] {
  // // Normalize host-read subtitle items into editor-ready text blocks with explicit word arrays.
  return items.map((item) =>
    syncTextEditorBlock({
      ...item,
      words: Array.isArray(item.words) ? item.words.slice() : normalizeTextEditorWords(item.text)
    })
  );
}

export function updateTextEditorBlockText(blocks: TextEditorBlock[], blockIndex: number, nextText: string): TextEditorBlock[] {
  // // Update one subtitle text field while keeping its words list synchronized for later drag/split actions.
  return cloneTextEditorBlocks(blocks).map((block, index) => {
    if (index !== blockIndex) {
      return block;
    }
    return syncTextEditorBlock({
      ...block,
      text: String(nextText || ""),
      words: normalizeTextEditorWords(nextText)
    });
  });
}

export function moveTextEditorWord(
  blocks: TextEditorBlock[],
  sourceBlockIndex: number,
  sourceWordIndex: number,
  targetBlockIndex: number,
  targetWordIndex?: number
): TextEditorBlock[] {
  // // Move one word between subtitle blocks, or within one block, using drag-and-drop insertion semantics.
  const nextBlocks = cloneTextEditorBlocks(blocks);
  const sourceBlock = nextBlocks[sourceBlockIndex];
  const targetBlock = nextBlocks[targetBlockIndex];

  if (!sourceBlock || !targetBlock) {
    return nextBlocks;
  }
  if (sourceWordIndex < 0 || sourceWordIndex >= sourceBlock.words.length) {
    return nextBlocks;
  }

  const movedWord = sourceBlock.words[sourceWordIndex];
  if (!movedWord) {
    return nextBlocks;
  }

  sourceBlock.words.splice(sourceWordIndex, 1);

  if (sourceBlockIndex === targetBlockIndex) {
    const adjustedTargetIndex = clampInsertIndex(
      targetBlock.words.length,
      typeof targetWordIndex === "number" && targetWordIndex > sourceWordIndex ? targetWordIndex - 1 : targetWordIndex
    );
    targetBlock.words.splice(adjustedTargetIndex, 0, movedWord);
    return nextBlocks.map((block) => syncTextEditorBlock(block));
  }

  const insertionIndex = clampInsertIndex(targetBlock.words.length, targetWordIndex);
  targetBlock.words.splice(insertionIndex, 0, movedWord);
  return filterNonEmptyTextEditorBlocks(nextBlocks);
}

export function splitTextEditorBlock(blocks: TextEditorBlock[], blockIndex: number, splitWordIndex: number): TextEditorBlock[] {
  // // Split one subtitle block into two consecutive blocks before the chosen word index.
  const nextBlocks = cloneTextEditorBlocks(blocks);
  const block = nextBlocks[blockIndex];
  if (!block) {
    return nextBlocks;
  }
  if (splitWordIndex <= 0 || splitWordIndex >= block.words.length) {
    return nextBlocks;
  }

  const leftWords = block.words.slice(0, splitWordIndex);
  const rightWords = block.words.slice(splitWordIndex);
  nextBlocks.splice(
    blockIndex,
    1,
    syncTextEditorBlock({
      ...block,
      text: leftWords.join(" "),
      words: leftWords
    }),
    syncTextEditorBlock({
      ...block,
      text: rightWords.join(" "),
      words: rightWords
    })
  );
  return nextBlocks;
}

export function mergeTextEditorBlocks(
  blocks: TextEditorBlock[],
  blockIndex: number,
  direction: "previous" | "next"
): TextEditorBlock[] {
  // // Merge one subtitle block with its previous or next neighbor while preserving the surviving source style index.
  const nextBlocks = cloneTextEditorBlocks(blocks);
  if (direction === "previous") {
    if (blockIndex <= 0 || blockIndex >= nextBlocks.length) {
      return nextBlocks;
    }
    const currentBlock = nextBlocks[blockIndex];
    const previousBlock = nextBlocks[blockIndex - 1];
    previousBlock.words = previousBlock.words.concat(currentBlock.words);
    nextBlocks.splice(blockIndex, 1);
    return nextBlocks.map((block) => syncTextEditorBlock(block));
  }

  if (blockIndex < 0 || blockIndex >= nextBlocks.length - 1) {
    return nextBlocks;
  }
  const block = nextBlocks[blockIndex];
  const nextBlock = nextBlocks[blockIndex + 1];
  block.words = block.words.concat(nextBlock.words);
  nextBlocks.splice(blockIndex + 1, 1);
  return nextBlocks.map((item) => syncTextEditorBlock(item));
}

export function retimeTextEditorBlocks(blocks: TextEditorBlock[]): TextEditorBlock[] {
  // // Redistribute the selected subtitle time span across edited blocks using weighted per-word timing.
  const sourceBlocks = cloneTextEditorBlocks(blocks);
  const normalizedBlocks = filterNonEmptyTextEditorBlocks(sourceBlocks);
  if (normalizedBlocks.length < 1) {
    return [];
  }

  const startSeconds = sourceBlocks.reduce(
    (lowestValue, block) => Math.min(lowestValue, Number(block.startSeconds || 0)),
    Number.POSITIVE_INFINITY
  );
  const endSeconds = sourceBlocks.reduce(
    (highestValue, block) => Math.max(highestValue, Number(block.endSeconds || 0)),
    Number.NEGATIVE_INFINITY
  );
  if (!(endSeconds > startSeconds)) {
    return normalizedBlocks;
  }

  const allWords = normalizedBlocks.flatMap((block) => block.words);
  if (allWords.length < 1) {
    return normalizedBlocks;
  }

  const weightedWords = buildWeightedCaptionWords(allWords, startSeconds, endSeconds);
  let wordCursor = 0;
  return normalizedBlocks.map((block, blockIndex) => {
    const wordCount = block.words.length;
    const blockWords = weightedWords.slice(wordCursor, wordCursor + wordCount);
    wordCursor += wordCount;
    if (blockWords.length < 1) {
      return block;
    }

    return {
      ...block,
      startSeconds: blockWords[0].startSeconds,
      endSeconds: blockIndex === normalizedBlocks.length - 1 ? endSeconds : blockWords[blockWords.length - 1].endSeconds,
      text: block.words.join(" "),
      words: block.words.slice()
    };
  });
}

export function sanitizeTextEditorBlocksForApply(blocks: TextEditorBlock[]): TextEditorBlock[] {
  // // Remove empty blocks and retime the remaining subtitles before they are rebuilt on the Premiere timeline.
  return retimeTextEditorBlocks(blocks);
}
