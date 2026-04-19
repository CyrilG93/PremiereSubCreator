// // Build subtitle cues from parsed content based on style constraints.
import type { CaptionBuildOptions, CaptionCue, CaptionWord } from "./types";
import { normalizeCaptionWords, tokenizeSubtitleText } from "./textNormalization";
import { buildWeightedCaptionWords } from "./wordTiming";

function normalizeWords(text: string): string[] {
  // // Keep contiguous word order so chunk timing stays coherent with speech rhythm.
  return tokenizeSubtitleText(text);
}

interface LineWrapConstraints {
  maxCharsPerLine: number;
  maxWordsPerLine: number;
}

function sanitizeMaxWordsPerLine(maxWordsPerLine: number): number {
  // // Treat invalid or missing word limits as "no extra limit" so legacy behavior stays intact by default.
  const normalizedValue = Number(maxWordsPerLine || 0);
  if (!Number.isFinite(normalizedValue) || normalizedValue < 1) {
    return Number.MAX_SAFE_INTEGER;
  }

  return Math.max(1, Math.floor(normalizedValue));
}

function wrapWordsWithConstraints(words: string[], constraints: LineWrapConstraints): string[] {
  // // Wrap words greedily while respecting both the character limit and the optional max-words-per-line limit.
  const lines: string[] = [];
  let currentLineWords: string[] = [];

  for (const word of words) {
    const candidateWords = [...currentLineWords, word];
    const candidate = candidateWords.join(" ");
    if (
      currentLineWords.length > 0 &&
      (candidate.length > constraints.maxCharsPerLine || candidateWords.length > constraints.maxWordsPerLine)
    ) {
      lines.push(currentLineWords.join(" "));
      currentLineWords = [word];
      continue;
    }

    currentLineWords = candidateWords;
  }

  if (currentLineWords.length > 0) {
    lines.push(currentLineWords.join(" "));
  }

  return lines;
}

function rebalanceWrappedLines(lines: string[], constraints: LineWrapConstraints): string[] {
  // // Balance neighboring lines so we avoid one long line and one very short line without breaking the wrap limits.
  if (lines.length < 2) {
    return lines;
  }

  const balanced = [...lines];
  let changed = true;
  let safety = 0;

  while (changed && safety < 32) {
    changed = false;
    safety += 1;

    for (let index = 0; index < balanced.length - 1; index += 1) {
      const currentWords = balanced[index].split(/\s+/).filter(Boolean);
      const nextWords = balanced[index + 1].split(/\s+/).filter(Boolean);
      if (currentWords.length <= 1 || nextWords.length < 1) {
        continue;
      }

      const movedWord = currentWords[currentWords.length - 1];
      const candidateCurrent = currentWords.slice(0, -1).join(" ");
      const candidateNext = [movedWord, ...nextWords].join(" ");
      if (
        candidateNext.length > constraints.maxCharsPerLine ||
        nextWords.length + 1 > constraints.maxWordsPerLine ||
        candidateCurrent.length < 1
      ) {
        continue;
      }

      const beforeMax = Math.max(balanced[index].length, balanced[index + 1].length);
      const afterMax = Math.max(candidateCurrent.length, candidateNext.length);
      if (afterMax < beforeMax) {
        balanced[index] = candidateCurrent;
        balanced[index + 1] = candidateNext;
        changed = true;
      }
    }
  }

  return balanced;
}

function findChunkEndIndex(words: string[], startIndex: number, constraints: LineWrapConstraints, linesPerCaption: number): number {
  // // Find the largest contiguous word range that still fits the configured line count.
  let bestEnd = Math.min(startIndex + 1, words.length);

  for (let endIndex = startIndex + 1; endIndex <= words.length; endIndex += 1) {
    const chunkWords = words.slice(startIndex, endIndex);
    const wrapped = wrapWordsWithConstraints(chunkWords, constraints);
    if (wrapped.length > linesPerCaption) {
      break;
    }

    bestEnd = endIndex;
  }

  return bestEnd;
}

interface ChunkRange {
  start: number;
  end: number;
}

function wordHasBoundaryPunctuation(word: string): boolean {
  // // Detect words that should not start a caption chunk by themselves (e.g. "time,").
  return /[,;:!?]$/.test(word.trim());
}

function wordStartsWithPunctuation(word: string): boolean {
  // // Detect leading punctuation that should stay attached to previous words.
  return /^[,;:!?)\]}]/.test(word.trim());
}

function normalizeBoundaryWord(word: string): string {
  // // Strip punctuation around boundary words to evaluate linguistic glue words.
  return word
    .trim()
    .toLowerCase()
    .replace(/^[^\p{L}\p{N}]+/gu, "")
    .replace(/[^\p{L}\p{N}]+$/gu, "");
}

function wordIsBoundaryConnector(word: string): boolean {
  // // Keep lightweight connector words attached to neighboring phrases instead of isolating them around commas.
  const normalized = normalizeBoundaryWord(word);
  return (
    normalized === "and" ||
    normalized === "or" ||
    normalized === "but" ||
    normalized === "so" ||
    normalized === "because" ||
    normalized === "since" ||
    normalized === "that" ||
    normalized === "which" ||
    normalized === "who" ||
    normalized === "when" ||
    normalized === "while" ||
    normalized === "if" ||
    normalized === "than" ||
    normalized === "then" ||
    normalized === "to" ||
    normalized === "for" ||
    normalized === "with" ||
    normalized === "from" ||
    normalized === "at" ||
    normalized === "on" ||
    normalized === "in" ||
    normalized === "of" ||
    normalized === "et" ||
    normalized === "ou" ||
    normalized === "mais" ||
    normalized === "donc" ||
    normalized === "car" ||
    normalized === "puis" ||
    normalized === "alors" ||
    normalized === "que" ||
    normalized === "qui" ||
    normalized === "quand" ||
    normalized === "comme" ||
    normalized === "pour" ||
    normalized === "avec" ||
    normalized === "sans" ||
    normalized === "dans" ||
    normalized === "sur" ||
    normalized === "en" ||
    normalized === "de" ||
    normalized === "du" ||
    normalized === "des" ||
    normalized === "au" ||
    normalized === "aux"
  );
}

function wordIsWeakEnding(word: string): boolean {
  // // Avoid finishing lines with connectors that read better when attached to next words.
  return wordIsBoundaryConnector(word);
}

function chunkFits(words: string[], range: ChunkRange, constraints: LineWrapConstraints, linesPerCaption: number): boolean {
  // // Validate whether a chunk range still respects the max-lines wrapping rule.
  if (range.end <= range.start) {
    return false;
  }

  const wrapped = wrapWordsWithConstraints(words.slice(range.start, range.end), constraints);
  return wrapped.length <= linesPerCaption;
}

function buildInitialChunkRanges(words: string[], constraints: LineWrapConstraints, linesPerCaption: number): ChunkRange[] {
  // // Build first-pass chunk ranges using the widest valid contiguous groups.
  const ranges: ChunkRange[] = [];
  let cursor = 0;

  while (cursor < words.length) {
    const chunkEnd = findChunkEndIndex(words, cursor, constraints, linesPerCaption);
    const safeEnd = Math.max(cursor + 1, chunkEnd);
    ranges.push({ start: cursor, end: safeEnd });
    cursor = safeEnd;
  }

  return ranges;
}

function rebalanceChunkRanges(ranges: ChunkRange[], words: string[], constraints: LineWrapConstraints, linesPerCaption: number): ChunkRange[] {
  // // Rebalance neighbor chunks to avoid tiny trailing chunks like a single-word caption.
  if (ranges.length < 2) {
    return ranges;
  }

  const balanced = ranges.map((range) => ({ ...range }));
  let pass = 0;
  while (pass < 8) {
    let changed = false;
    pass += 1;

    for (let index = 0; index < balanced.length - 1; index += 1) {
      const left = balanced[index];
      const right = balanced[index + 1];
      let leftSize = left.end - left.start;
      let rightSize = right.end - right.start;

      while (rightSize < 2 && leftSize > 1) {
        const candidateBoundary = right.start - 1;
        const candidateRight = { start: candidateBoundary, end: right.end };
        if (!chunkFits(words, candidateRight, constraints, linesPerCaption)) {
          break;
        }

        left.end = candidateBoundary;
        right.start = candidateBoundary;
        leftSize -= 1;
        rightSize += 1;
        changed = true;
      }

      while (leftSize - rightSize > 2 && leftSize > 1) {
        const candidateBoundary = right.start - 1;
        const candidateRight = { start: candidateBoundary, end: right.end };
        if (!chunkFits(words, candidateRight, constraints, linesPerCaption)) {
          break;
        }

        left.end = candidateBoundary;
        right.start = candidateBoundary;
        leftSize -= 1;
        rightSize += 1;
        changed = true;
      }

      // // Keep punctuation-attached words with the previous phrase when possible.
      const rightFirstWord = words[right.start] ?? "";
      if ((wordHasBoundaryPunctuation(rightFirstWord) || wordStartsWithPunctuation(rightFirstWord)) && rightSize > 1) {
        const moveRightBoundary = right.start + 1;
        const leftIfMoveRight = { start: left.start, end: moveRightBoundary };
        const rightIfMoveRight = { start: moveRightBoundary, end: right.end };
        if (
          rightIfMoveRight.end - rightIfMoveRight.start > 0 &&
          chunkFits(words, leftIfMoveRight, constraints, linesPerCaption) &&
          chunkFits(words, rightIfMoveRight, constraints, linesPerCaption)
        ) {
          left.end = moveRightBoundary;
          right.start = moveRightBoundary;
          leftSize = left.end - left.start;
          rightSize = right.end - right.start;
          changed = true;
        } else if (leftSize > 1) {
          const moveLeftBoundary = right.start - 1;
          const leftIfMoveLeft = { start: left.start, end: moveLeftBoundary };
          const rightIfMoveLeft = { start: moveLeftBoundary, end: right.end };
          const rightIfMoveLeftSize = rightIfMoveLeft.end - rightIfMoveLeft.start;
          if (
            rightIfMoveLeftSize > 1 &&
            chunkFits(words, leftIfMoveLeft, constraints, linesPerCaption) &&
            chunkFits(words, rightIfMoveLeft, constraints, linesPerCaption)
          ) {
            left.end = moveLeftBoundary;
            right.start = moveLeftBoundary;
            leftSize = left.end - left.start;
            rightSize = right.end - right.start;
            changed = true;
          }
        }
      }

      // // Avoid ending a chunk with weak connector words when right chunk can absorb them.
      const leftLastWord = words[left.end - 1] ?? "";
      if (wordIsWeakEnding(leftLastWord) && leftSize > 1) {
        const moveWeakEnding = right.start - 1;
        const leftIfMoveWeak = { start: left.start, end: moveWeakEnding };
        const rightIfMoveWeak = { start: moveWeakEnding, end: right.end };
        if (
          rightIfMoveWeak.end - rightIfMoveWeak.start > 1 &&
          chunkFits(words, leftIfMoveWeak, constraints, linesPerCaption) &&
          chunkFits(words, rightIfMoveWeak, constraints, linesPerCaption)
        ) {
          left.end = moveWeakEnding;
          right.start = moveWeakEnding;
          leftSize = left.end - left.start;
          rightSize = right.end - right.start;
          changed = true;
        }
      }

      // // Keep short connector words after commas with the previous phrase when the right chunk can still stand on its own.
      const refreshedLeftLastWord = words[left.end - 1] ?? "";
      const refreshedRightFirstWord = words[right.start] ?? "";
      if (/[,:;]$/.test(refreshedLeftLastWord.trim()) && wordIsBoundaryConnector(refreshedRightFirstWord) && rightSize > 1) {
        const absorbConnectorBoundary = right.start + 1;
        const leftIfAbsorbConnector = { start: left.start, end: absorbConnectorBoundary };
        const rightIfAbsorbConnector = { start: absorbConnectorBoundary, end: right.end };
        if (
          rightIfAbsorbConnector.end - rightIfAbsorbConnector.start > 0 &&
          chunkFits(words, leftIfAbsorbConnector, constraints, linesPerCaption) &&
          chunkFits(words, rightIfAbsorbConnector, constraints, linesPerCaption)
        ) {
          left.end = absorbConnectorBoundary;
          right.start = absorbConnectorBoundary;
          changed = true;
        }
      }
    }

    if (!changed) {
      break;
    }
  }

  return balanced;
}

function uppercaseIfNeeded(text: string, forceUppercase: boolean): string {
  // // Optionally convert text to uppercase for aggressive social-media subtitle styles.
  return forceUppercase ? text.toUpperCase() : text;
}

function ensureCueWords(cue: CaptionCue, forceUppercase: boolean): CaptionWord[] {
  // // Guarantee word-level timing exists, synthesizing it when source lacks per-word data.
  if (cue.words.length > 0) {
    return normalizeCaptionWords(cue.words).map((word) => {
      return {
        ...word,
        text: uppercaseIfNeeded(word.text, forceUppercase)
      };
    });
  }

  const words = normalizeWords(uppercaseIfNeeded(cue.text, forceUppercase));
  if (words.length < 1) {
    return [];
  }

  // // Estimate synthetic timings from weighted word lengths so chunk durations better follow reading/speaking density.
  return buildWeightedCaptionWords(words, cue.startSeconds, cue.endSeconds);
}

function renderChunkText(words: CaptionWord[], constraints: LineWrapConstraints): string {
  // // Render final chunk text with explicit line breaks to stabilize MOGRT layout.
  const wrapped = wrapWordsWithConstraints(words.map((word) => word.text), constraints);
  const balanced = rebalanceWrappedLines(wrapped, constraints);
  return balanced.join("\n");
}

export function buildCaptionPlan(cues: CaptionCue[], options: CaptionBuildOptions): CaptionCue[] {
  // // Split cues by contiguous word groups with timing proportional to word distribution.
  const plannedCues: CaptionCue[] = [];
  const constraints: LineWrapConstraints = {
    maxCharsPerLine: Math.max(6, Number(options.style.maxCharsPerLine || 28)),
    maxWordsPerLine: sanitizeMaxWordsPerLine(Number(options.style.maxWordsPerLine || 12))
  };
  const linesPerCaption = Math.max(1, Number(options.style.linesPerCaption || 2));

  for (const cue of cues) {
    const normalizedWords = ensureCueWords(cue, options.style.uppercase);
    if (normalizedWords.length < 1) {
      const normalizedText = uppercaseIfNeeded(cue.text, options.style.uppercase);
      if (!normalizedText.trim()) {
        continue;
      }

      plannedCues.push({
        ...cue,
        text: normalizedText
      });
      continue;
    }

    const cueWordTexts = normalizedWords.map((word) => word.text);
    const chunkRanges = rebalanceChunkRanges(
      buildInitialChunkRanges(cueWordTexts, constraints, linesPerCaption),
      cueWordTexts,
      constraints,
      linesPerCaption
    );

    chunkRanges.forEach((range, chunkIndex) => {
      const chunkWords = normalizedWords.slice(range.start, range.end);
      const firstWord = chunkWords[0];
      const lastWord = chunkWords[chunkWords.length - 1];
      const startSeconds = chunkIndex === 0 ? cue.startSeconds : firstWord.startSeconds;
      const endSeconds = range.end >= normalizedWords.length ? cue.endSeconds : lastWord.endSeconds;

      plannedCues.push({
        id: `${cue.id}-part-${chunkIndex + 1}`,
        startSeconds,
        endSeconds,
        text: renderChunkText(chunkWords, constraints),
        words: chunkWords
      });
    });
  }

  return plannedCues;
}
