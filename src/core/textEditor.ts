// // Provide pure text-editing helpers for the Text tab subtitle workflow.
import type { CaptionWord } from "./types";
import { tokenizeSubtitleText } from "./textNormalization";
import { buildWeightedCaptionWords } from "./wordTiming";

// // Describe one editable subtitle block shown in the Text tab.
export interface TextEditorBlock {
  sourceSelectionIndex: number;
  clipName: string;
  startSeconds: number;
  endSeconds: number;
  text: string;
  words: string[];
  timedWords?: CaptionWord[];
}

export interface TextEditorTimingRange {
  startSeconds: number;
  endSeconds: number;
}

export interface TextEditorApplyPlan {
  selectionStartIndex: number;
  selectionEndIndex: number;
  timingRange: TextEditorTimingRange;
  blocks: TextEditorBlock[];
}

interface TextEditorChangedRange {
  originalStartIndex: number;
  originalEndIndexExclusive: number;
  editedStartIndex: number;
  editedEndIndexExclusive: number;
}

function resolveSourceSelectionRange(
  blocks: TextEditorBlock[],
  originalStartIndex: number,
  originalEndIndexExclusive: number
): { startIndex: number; endIndex: number } {
  // // Map one changed block slice back to the real Premiere selection indexes so filtered-out non-subtitle entries do not offset apply ranges.
  const slice = blocks.slice(
    Math.max(0, originalStartIndex),
    Math.max(Math.max(0, originalStartIndex) + 1, originalEndIndexExclusive)
  );
  if (slice.length < 1) {
    return {
      startIndex: Math.max(0, originalStartIndex),
      endIndex: Math.max(0, originalStartIndex)
    };
  }

  const sourceIndexes = slice
    .map((block) => Number(block.sourceSelectionIndex))
    .filter((value) => Number.isFinite(value))
    .sort((left, right) => left - right);

  if (sourceIndexes.length < 1) {
    return {
      startIndex: Math.max(0, originalStartIndex),
      endIndex: Math.max(0, originalStartIndex)
    };
  }

  return {
    startIndex: sourceIndexes[0] || 0,
    endIndex: sourceIndexes[sourceIndexes.length - 1] || sourceIndexes[0] || 0
  };
}

function normalizeTextEditorWords(text: string): string[] {
  // // Split one subtitle text into compact word tokens while preserving punctuation on each token.
  return tokenizeSubtitleText(text);
}

function syncTextEditorBlock(block: TextEditorBlock): TextEditorBlock {
  // // Keep `text` and `words` synchronized after one editing operation.
  const words = Array.isArray(block.words) && block.words.length > 0 ? block.words.slice() : normalizeTextEditorWords(block.text);
  const timedWords = areTimedWordsCompatible(words, block.timedWords) ? cloneTimedWords(block.timedWords || []) : undefined;
  return {
    sourceSelectionIndex: Number.isFinite(Number(block.sourceSelectionIndex)) ? Number(block.sourceSelectionIndex) : 0,
    clipName: String(block.clipName || "").trim(),
    startSeconds: Number(block.startSeconds || 0),
    endSeconds: Number(block.endSeconds || 0),
    text: words.join(" "),
    words,
    timedWords
  };
}

function areTextEditorBlocksEquivalent(left: TextEditorBlock, right: TextEditorBlock): boolean {
  // // Match unchanged prefix/suffix rows by original source index plus normalized text content.
  const normalizedLeft = syncTextEditorBlock(left);
  const normalizedRight = syncTextEditorBlock(right);
  return (
    normalizedLeft.sourceSelectionIndex === normalizedRight.sourceSelectionIndex &&
    normalizedLeft.text === normalizedRight.text
  );
}

function cloneTextEditorBlocks(blocks: TextEditorBlock[]): TextEditorBlock[] {
  // // Clone editable blocks so UI reducers can stay immutable and predictable.
  return blocks.map((block) => syncTextEditorBlock(block));
}

function normalizeTimedWordText(word: CaptionWord | undefined): string {
  // // Normalize one timed-word token so editor text and metadata can be compared safely.
  return String(word?.text || "").trim();
}

function cloneTimedWords(words: CaptionWord[]): CaptionWord[] {
  // // Clone word timing entries to preserve immutability across editor operations.
  return words.map((word) => ({
    text: String(word.text || "").trim(),
    startSeconds: Number(word.startSeconds || 0),
    endSeconds: Number(word.endSeconds || 0)
  }));
}

function areTimedWordsCompatible(words: string[], timedWords?: CaptionWord[]): boolean {
  // // Keep precise timings only when they still map one-to-one to the visible editor words.
  if (!Array.isArray(timedWords)) {
    return false;
  }
  if (timedWords.length !== words.length) {
    return false;
  }

  for (let index = 0; index < words.length; index += 1) {
    if (normalizeTimedWordText(timedWords[index]) !== String(words[index] || "").trim()) {
      return false;
    }
  }

  return true;
}

function buildPreservedTimedWordsOrNull(block: TextEditorBlock): CaptionWord[] | null {
  // // Reuse precise timings when one block still has a valid timed-word payload.
  return areTimedWordsCompatible(block.words, block.timedWords) ? cloneTimedWords(block.timedWords || []) : null;
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
    const nextWords = normalizeTextEditorWords(nextText);
    return syncTextEditorBlock({
      ...block,
      text: String(nextText || ""),
      words: nextWords,
      timedWords: areTimedWordsCompatible(nextWords, block.timedWords) ? block.timedWords : undefined
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

  const sourceTimedWords = buildPreservedTimedWordsOrNull(sourceBlock);
  const targetTimedWords = buildPreservedTimedWordsOrNull(targetBlock);
  const movedTimedWord = sourceTimedWords ? sourceTimedWords[sourceWordIndex] : null;

  sourceBlock.words.splice(sourceWordIndex, 1);
  if (sourceTimedWords && movedTimedWord) {
    sourceTimedWords.splice(sourceWordIndex, 1);
  }

  if (sourceBlockIndex === targetBlockIndex) {
    const adjustedTargetIndex = clampInsertIndex(
      targetBlock.words.length,
      typeof targetWordIndex === "number" && targetWordIndex > sourceWordIndex ? targetWordIndex - 1 : targetWordIndex
    );
    targetBlock.words.splice(adjustedTargetIndex, 0, movedWord);
    if (sourceTimedWords && movedTimedWord) {
      sourceTimedWords.splice(adjustedTargetIndex, 0, movedTimedWord);
      targetBlock.timedWords = sourceTimedWords;
    } else {
      targetBlock.timedWords = undefined;
    }
    return nextBlocks.map((block) => syncTextEditorBlock(block));
  }

  const insertionIndex = clampInsertIndex(targetBlock.words.length, targetWordIndex);
  targetBlock.words.splice(insertionIndex, 0, movedWord);
  if (sourceTimedWords && targetTimedWords && movedTimedWord) {
    sourceBlock.timedWords = sourceTimedWords;
    targetTimedWords.splice(insertionIndex, 0, movedTimedWord);
    targetBlock.timedWords = targetTimedWords;
  } else {
    sourceBlock.timedWords = undefined;
    targetBlock.timedWords = undefined;
  }
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
  const timedWords = buildPreservedTimedWordsOrNull(block);
  const leftTimedWords = timedWords ? timedWords.slice(0, splitWordIndex) : undefined;
  const rightTimedWords = timedWords ? timedWords.slice(splitWordIndex) : undefined;
  nextBlocks.splice(
    blockIndex,
    1,
    syncTextEditorBlock({
      ...block,
      text: leftWords.join(" "),
      words: leftWords,
      timedWords: leftTimedWords
    }),
    syncTextEditorBlock({
      ...block,
      text: rightWords.join(" "),
      words: rightWords,
      timedWords: rightTimedWords
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
    const previousTimedWords = buildPreservedTimedWordsOrNull(previousBlock);
    const currentTimedWords = buildPreservedTimedWordsOrNull(currentBlock);
    previousBlock.words = previousBlock.words.concat(currentBlock.words);
    previousBlock.timedWords = previousTimedWords && currentTimedWords ? previousTimedWords.concat(currentTimedWords) : undefined;
    nextBlocks.splice(blockIndex, 1);
    return nextBlocks.map((block) => syncTextEditorBlock(block));
  }

  if (blockIndex < 0 || blockIndex >= nextBlocks.length - 1) {
    return nextBlocks;
  }
  const block = nextBlocks[blockIndex];
  const nextBlock = nextBlocks[blockIndex + 1];
  const blockTimedWords = buildPreservedTimedWordsOrNull(block);
  const nextTimedWords = buildPreservedTimedWordsOrNull(nextBlock);
  block.words = block.words.concat(nextBlock.words);
  block.timedWords = blockTimedWords && nextTimedWords ? blockTimedWords.concat(nextTimedWords) : undefined;
  nextBlocks.splice(blockIndex + 1, 1);
  return nextBlocks.map((item) => syncTextEditorBlock(item));
}

function resolveTextEditorTimingRange(
  blocks: TextEditorBlock[],
  timingRange?: TextEditorTimingRange
): TextEditorTimingRange | null {
  // // Prefer one explicit selection range so merges keep the full original span even after edge blocks disappear.
  if (
    timingRange &&
    Number.isFinite(Number(timingRange.startSeconds)) &&
    Number.isFinite(Number(timingRange.endSeconds)) &&
    Number(timingRange.endSeconds) > Number(timingRange.startSeconds)
  ) {
    return {
      startSeconds: Number(timingRange.startSeconds),
      endSeconds: Number(timingRange.endSeconds)
    };
  }

  if (blocks.length < 1) {
    return null;
  }

  const startSeconds = blocks.reduce(
    (lowestValue, block) => Math.min(lowestValue, Number(block.startSeconds || 0)),
    Number.POSITIVE_INFINITY
  );
  const endSeconds = blocks.reduce(
    (highestValue, block) => Math.max(highestValue, Number(block.endSeconds || 0)),
    Number.NEGATIVE_INFINITY
  );
  if (!(endSeconds > startSeconds)) {
    return null;
  }

  return {
    startSeconds,
    endSeconds
  };
}

function buildWeightedRetimedTextEditorBlocks(
  blocks: TextEditorBlock[],
  startSeconds: number,
  endSeconds: number
): TextEditorBlock[] {
  // // Rebuild one contiguous block range with weighted word timings when precise timed-word anchors are unavailable.
  const allWords = blocks.flatMap((block) => block.words);
  if (allWords.length < 1) {
    return blocks;
  }

  const weightedWords = buildWeightedCaptionWords(allWords, startSeconds, endSeconds);
  let wordCursor = 0;
  return blocks.map((block, blockIndex) => {
    const wordCount = block.words.length;
    const blockWords = weightedWords.slice(wordCursor, wordCursor + wordCount);
    wordCursor += wordCount;
    if (blockWords.length < 1) {
      return block;
    }

    return {
      ...block,
      startSeconds: blockWords[0].startSeconds,
      endSeconds: blockIndex === blocks.length - 1 ? endSeconds : blockWords[blockWords.length - 1].endSeconds,
      text: block.words.join(" "),
      words: block.words.slice(),
      timedWords: blockWords
    };
  });
}

export function retimeTextEditorBlocks(blocks: TextEditorBlock[], timingRange?: TextEditorTimingRange): TextEditorBlock[] {
  // // Redistribute the selected subtitle time span across edited blocks using weighted per-word timing.
  const sourceBlocks = cloneTextEditorBlocks(blocks);
  const normalizedBlocks = filterNonEmptyTextEditorBlocks(sourceBlocks);
  if (normalizedBlocks.length < 1) {
    return [];
  }

  const resolvedTimingRange = resolveTextEditorTimingRange(sourceBlocks, timingRange);
  if (!resolvedTimingRange) {
    return normalizedBlocks;
  }
  const { startSeconds, endSeconds } = resolvedTimingRange;

  const timedWordPayloads = normalizedBlocks.map((block) =>
    areTimedWordsCompatible(block.words, block.timedWords) ? cloneTimedWords(block.timedWords || []) : null
  );
  const compatibleTimedBlockCount = timedWordPayloads.filter((payload) => Array.isArray(payload) && payload.length > 0).length;
  if (compatibleTimedBlockCount < 1) {
    return buildWeightedRetimedTextEditorBlocks(normalizedBlocks, startSeconds, endSeconds);
  }

  if (compatibleTimedBlockCount === normalizedBlocks.length) {
    return normalizedBlocks.map((block, blockIndex) => {
      const timedWords = cloneTimedWords(block.timedWords || []);
      if (timedWords.length < 1) {
        return block;
      }

      return {
        ...block,
        startSeconds: blockIndex === 0 ? startSeconds : timedWords[0].startSeconds,
        endSeconds: blockIndex === normalizedBlocks.length - 1 ? endSeconds : timedWords[timedWords.length - 1].endSeconds,
        text: block.words.join(" "),
        words: block.words.slice(),
        timedWords
      };
    });
  }

  const retimedBlocks = normalizedBlocks.map((block) => ({
    ...block,
    text: block.words.join(" "),
    words: block.words.slice(),
    timedWords: undefined as CaptionWord[] | undefined
  }));
  let previousAnchorEnd = startSeconds;
  let incompatibleRunStartIndex = -1;

  const applyWeightedRun = (runStartIndex: number, runEndIndex: number, rangeStart: number, rangeEnd: number): boolean => {
    // // Rebuild only the incompatible block run between two timing anchors while keeping surrounding precise timings intact.
    if (runStartIndex > runEndIndex) {
      return true;
    }
    if (!(rangeEnd > rangeStart)) {
      return false;
    }

    const weightedRun = buildWeightedRetimedTextEditorBlocks(
      normalizedBlocks.slice(runStartIndex, runEndIndex + 1),
      rangeStart,
      rangeEnd
    );
    for (let runIndex = 0; runIndex < weightedRun.length; runIndex += 1) {
      retimedBlocks[runStartIndex + runIndex] = weightedRun[runIndex];
    }
    return true;
  };

  for (let blockIndex = 0; blockIndex < normalizedBlocks.length; blockIndex += 1) {
    const timedWords = timedWordPayloads[blockIndex];
    if (!timedWords || timedWords.length < 1) {
      if (incompatibleRunStartIndex < 0) {
        incompatibleRunStartIndex = blockIndex;
      }
      continue;
    }

    if (incompatibleRunStartIndex >= 0) {
      if (!applyWeightedRun(incompatibleRunStartIndex, blockIndex - 1, previousAnchorEnd, timedWords[0].startSeconds)) {
        return buildWeightedRetimedTextEditorBlocks(normalizedBlocks, startSeconds, endSeconds);
      }
      incompatibleRunStartIndex = -1;
    }

    retimedBlocks[blockIndex] = {
      ...normalizedBlocks[blockIndex],
      startSeconds: blockIndex === 0 ? startSeconds : timedWords[0].startSeconds,
      endSeconds: blockIndex === normalizedBlocks.length - 1 ? endSeconds : timedWords[timedWords.length - 1].endSeconds,
      text: normalizedBlocks[blockIndex].words.join(" "),
      words: normalizedBlocks[blockIndex].words.slice(),
      timedWords
    };
    previousAnchorEnd = timedWords[timedWords.length - 1].endSeconds;
  }

  if (incompatibleRunStartIndex >= 0) {
    if (!applyWeightedRun(incompatibleRunStartIndex, normalizedBlocks.length - 1, previousAnchorEnd, endSeconds)) {
      return buildWeightedRetimedTextEditorBlocks(normalizedBlocks, startSeconds, endSeconds);
    }
  }

  if (retimedBlocks.length > 0) {
    retimedBlocks[0].startSeconds = startSeconds;
    retimedBlocks[retimedBlocks.length - 1].endSeconds = endSeconds;
  }

  return retimedBlocks;
}

export function sanitizeTextEditorBlocksForApply(
  blocks: TextEditorBlock[],
  timingRange?: TextEditorTimingRange
): TextEditorBlock[] {
  // // Remove empty blocks and retime the remaining subtitles before they are rebuilt on the Premiere timeline.
  return retimeTextEditorBlocks(blocks, timingRange);
}

function buildTextEditorChangedRanges(
  originalBlocks: TextEditorBlock[],
  editedBlocks: TextEditorBlock[]
): TextEditorChangedRange[] {
  // // Compute minimal changed block ranges so disjoint edits can be rebuilt independently from right to left.
  const normalizedOriginalBlocks = cloneTextEditorBlocks(originalBlocks);
  const normalizedEditedBlocks = filterNonEmptyTextEditorBlocks(editedBlocks);
  if (normalizedOriginalBlocks.length < 1 && normalizedEditedBlocks.length < 1) {
    return [];
  }

  const originalLength = normalizedOriginalBlocks.length;
  const editedLength = normalizedEditedBlocks.length;
  const lcsLengths: number[][] = Array.from({ length: originalLength + 1 }, () => Array<number>(editedLength + 1).fill(0));

  for (let originalIndex = originalLength - 1; originalIndex >= 0; originalIndex -= 1) {
    for (let editedIndex = editedLength - 1; editedIndex >= 0; editedIndex -= 1) {
      if (areTextEditorBlocksEquivalent(normalizedOriginalBlocks[originalIndex], normalizedEditedBlocks[editedIndex])) {
        lcsLengths[originalIndex][editedIndex] = lcsLengths[originalIndex + 1][editedIndex + 1] + 1;
        continue;
      }

      lcsLengths[originalIndex][editedIndex] = Math.max(
        lcsLengths[originalIndex + 1][editedIndex],
        lcsLengths[originalIndex][editedIndex + 1]
      );
    }
  }

  const changedRanges: TextEditorChangedRange[] = [];
  let currentRange: TextEditorChangedRange | null = null;
  let originalCursor = 0;
  let editedCursor = 0;

  while (originalCursor < originalLength || editedCursor < editedLength) {
    if (
      originalCursor < originalLength &&
      editedCursor < editedLength &&
      areTextEditorBlocksEquivalent(normalizedOriginalBlocks[originalCursor], normalizedEditedBlocks[editedCursor])
    ) {
      if (
        currentRange &&
        (currentRange.originalEndIndexExclusive > currentRange.originalStartIndex ||
          currentRange.editedEndIndexExclusive > currentRange.editedStartIndex)
      ) {
        changedRanges.push(currentRange);
      }
      currentRange = null;
      originalCursor += 1;
      editedCursor += 1;
      continue;
    }

    if (!currentRange) {
      currentRange = {
        originalStartIndex: originalCursor,
        originalEndIndexExclusive: originalCursor,
        editedStartIndex: editedCursor,
        editedEndIndexExclusive: editedCursor
      };
    }

    const preferInsertion =
      editedCursor < editedLength &&
      (originalCursor >= originalLength || lcsLengths[originalCursor][editedCursor + 1] >= lcsLengths[originalCursor + 1][editedCursor]);
    if (preferInsertion) {
      editedCursor += 1;
      currentRange.editedEndIndexExclusive = editedCursor;
      continue;
    }

    if (originalCursor < originalLength) {
      originalCursor += 1;
      currentRange.originalEndIndexExclusive = originalCursor;
      continue;
    }
  }

  if (
    currentRange &&
    (currentRange.originalEndIndexExclusive > currentRange.originalStartIndex ||
      currentRange.editedEndIndexExclusive > currentRange.editedStartIndex)
  ) {
    changedRanges.push(currentRange);
  }

  return changedRanges;
}

export function buildTextEditorApplyPlans(
  originalBlocks: TextEditorBlock[],
  editedBlocks: TextEditorBlock[]
): TextEditorApplyPlan[] {
  // // Build one rebuild plan per contiguous changed range so disjoint edits do not force one large replacement.
  const normalizedOriginalBlocks = cloneTextEditorBlocks(originalBlocks);
  const normalizedEditedBlocks = filterNonEmptyTextEditorBlocks(editedBlocks);
  if (normalizedOriginalBlocks.length < 1) {
    return [];
  }

  const plans: TextEditorApplyPlan[] = [];
  const changedRanges = buildTextEditorChangedRanges(normalizedOriginalBlocks, normalizedEditedBlocks);
  for (const changedRange of changedRanges) {
    if (changedRange.originalStartIndex >= normalizedOriginalBlocks.length) {
      continue;
    }

    const originalSliceStartIndex = Math.max(0, Math.min(normalizedOriginalBlocks.length - 1, changedRange.originalStartIndex));
    const originalSliceEndIndex = Math.max(
      originalSliceStartIndex,
      Math.min(normalizedOriginalBlocks.length - 1, changedRange.originalEndIndexExclusive - 1)
    );
    const selectionRange = resolveSourceSelectionRange(
      normalizedOriginalBlocks,
      changedRange.originalStartIndex,
      changedRange.originalEndIndexExclusive
    );
    const changedEditedBlocks = normalizedEditedBlocks.slice(
      changedRange.editedStartIndex,
      changedRange.editedEndIndexExclusive
    );
    const timingRange = {
      startSeconds: Number(normalizedOriginalBlocks[originalSliceStartIndex].startSeconds || 0),
      endSeconds: Number(normalizedOriginalBlocks[originalSliceEndIndex].endSeconds || 0)
    };

    plans.push({
      selectionStartIndex: selectionRange.startIndex,
      selectionEndIndex: selectionRange.endIndex,
      timingRange,
      blocks: retimeTextEditorBlocks(changedEditedBlocks, timingRange)
    });
  }

  return plans;
}

export function buildTextEditorSafeApplyPlans(
  originalBlocks: TextEditorBlock[],
  editedBlocks: TextEditorBlock[]
): TextEditorApplyPlan[] {
  // // Collapse disjoint edits into one safe combined replacement span until the host can guarantee transactional multi-range rebuilds.
  const normalizedOriginalBlocks = cloneTextEditorBlocks(originalBlocks);
  const normalizedEditedBlocks = filterNonEmptyTextEditorBlocks(editedBlocks);
  const changedRanges = buildTextEditorChangedRanges(normalizedOriginalBlocks, normalizedEditedBlocks);
  if (changedRanges.length < 2) {
    return buildTextEditorApplyPlans(originalBlocks, editedBlocks);
  }

  const firstChangedRange = changedRanges[0];
  const lastChangedRange = changedRanges[changedRanges.length - 1];
  const originalSliceStartIndex = Math.max(
    0,
    Math.min(normalizedOriginalBlocks.length - 1, firstChangedRange.originalStartIndex)
  );
  const originalSliceEndIndex = Math.max(
    originalSliceStartIndex,
    Math.min(normalizedOriginalBlocks.length - 1, lastChangedRange.originalEndIndexExclusive - 1)
  );
  const selectionRange = resolveSourceSelectionRange(
    normalizedOriginalBlocks,
    firstChangedRange.originalStartIndex,
    lastChangedRange.originalEndIndexExclusive
  );
  const combinedEditedBlocks = normalizedEditedBlocks.slice(
    firstChangedRange.editedStartIndex,
    lastChangedRange.editedEndIndexExclusive
  );
  const timingRange = {
    startSeconds: Number(normalizedOriginalBlocks[originalSliceStartIndex]?.startSeconds || 0),
    endSeconds: Number(normalizedOriginalBlocks[originalSliceEndIndex]?.endSeconds || 0)
  };

  return [
    {
      selectionStartIndex: selectionRange.startIndex,
      selectionEndIndex: selectionRange.endIndex,
      timingRange,
      blocks: retimeTextEditorBlocks(combinedEditedBlocks, timingRange)
    }
  ];
}

export function buildTextEditorApplyPlan(
  originalBlocks: TextEditorBlock[],
  editedBlocks: TextEditorBlock[]
): TextEditorApplyPlan | null {
  // // Preserve the legacy single-range helper by returning the first changed plan when callers only expect one range.
  const plans = buildTextEditorApplyPlans(originalBlocks, editedBlocks);
  if (plans.length < 1) {
    return null;
  }

  return plans[0];
}
