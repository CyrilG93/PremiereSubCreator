// // Build subtitle cues from parsed content based on style constraints.
import type { CaptionBuildOptions, CaptionCue, CaptionWord } from "./types";

function normalizeWords(text: string): string[] {
  // // Keep contiguous word order so chunk timing stays coherent with speech rhythm.
  return text
    .replace(/\r/g, "\n")
    .replace(/\n+/g, " ")
    .split(/\s+/)
    .map((word) => word.trim())
    .filter(Boolean);
}

function wrapWordsByChars(words: string[], maxCharsPerLine: number): string[] {
  // // Wrap words greedily before optional balancing pass.
  const lines: string[] = [];
  let currentLine = "";

  for (const word of words) {
    const candidate = currentLine ? `${currentLine} ${word}` : word;
    if (currentLine && candidate.length > maxCharsPerLine) {
      lines.push(currentLine);
      currentLine = word;
      continue;
    }

    currentLine = candidate;
  }

  if (currentLine) {
    lines.push(currentLine);
  }

  return lines;
}

function rebalanceWrappedLines(lines: string[], maxCharsPerLine: number): string[] {
  // // Balance neighboring lines so we avoid one long line and one very short line.
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
      if (candidateNext.length > maxCharsPerLine || candidateCurrent.length < 1) {
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

function findChunkEndIndex(words: string[], startIndex: number, maxCharsPerLine: number, linesPerCaption: number): number {
  // // Find the largest contiguous word range that still fits the configured line count.
  let bestEnd = Math.min(startIndex + 1, words.length);

  for (let endIndex = startIndex + 1; endIndex <= words.length; endIndex += 1) {
    const chunkWords = words.slice(startIndex, endIndex);
    const wrapped = wrapWordsByChars(chunkWords, maxCharsPerLine);
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
    .replace(/^[^a-z0-9]+/i, "")
    .replace(/[^a-z0-9]+$/i, "");
}

function wordIsWeakEnding(word: string): boolean {
  // // Avoid finishing lines with connectors that read better when attached to next words.
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
    normalized === "of"
  );
}

function chunkFits(words: string[], range: ChunkRange, maxCharsPerLine: number, linesPerCaption: number): boolean {
  // // Validate whether a chunk range still respects the max-lines wrapping rule.
  if (range.end <= range.start) {
    return false;
  }

  const wrapped = wrapWordsByChars(words.slice(range.start, range.end), maxCharsPerLine);
  return wrapped.length <= linesPerCaption;
}

function buildInitialChunkRanges(words: string[], maxCharsPerLine: number, linesPerCaption: number): ChunkRange[] {
  // // Build first-pass chunk ranges using the widest valid contiguous groups.
  const ranges: ChunkRange[] = [];
  let cursor = 0;

  while (cursor < words.length) {
    const chunkEnd = findChunkEndIndex(words, cursor, maxCharsPerLine, linesPerCaption);
    const safeEnd = Math.max(cursor + 1, chunkEnd);
    ranges.push({ start: cursor, end: safeEnd });
    cursor = safeEnd;
  }

  return ranges;
}

function rebalanceChunkRanges(ranges: ChunkRange[], words: string[], maxCharsPerLine: number, linesPerCaption: number): ChunkRange[] {
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
        if (!chunkFits(words, candidateRight, maxCharsPerLine, linesPerCaption)) {
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
        if (!chunkFits(words, candidateRight, maxCharsPerLine, linesPerCaption)) {
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
          chunkFits(words, leftIfMoveRight, maxCharsPerLine, linesPerCaption) &&
          chunkFits(words, rightIfMoveRight, maxCharsPerLine, linesPerCaption)
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
            chunkFits(words, leftIfMoveLeft, maxCharsPerLine, linesPerCaption) &&
            chunkFits(words, rightIfMoveLeft, maxCharsPerLine, linesPerCaption)
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
          chunkFits(words, leftIfMoveWeak, maxCharsPerLine, linesPerCaption) &&
          chunkFits(words, rightIfMoveWeak, maxCharsPerLine, linesPerCaption)
        ) {
          left.end = moveWeakEnding;
          right.start = moveWeakEnding;
          leftSize = left.end - left.start;
          rightSize = right.end - right.start;
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
    return cue.words.map((word) => {
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

  const totalDuration = Math.max(cue.endSeconds - cue.startSeconds, 0.01);
  const wordDuration = totalDuration / words.length;

  return words.map((word, index) => {
    const startSeconds = cue.startSeconds + index * wordDuration;
    const endSeconds = index === words.length - 1 ? cue.endSeconds : startSeconds + wordDuration;
    return {
      text: word,
      startSeconds,
      endSeconds
    };
  });
}

function renderChunkText(words: CaptionWord[], maxCharsPerLine: number): string {
  // // Render final chunk text with explicit line breaks to stabilize MOGRT layout.
  const wrapped = wrapWordsByChars(
    words.map((word) => word.text),
    maxCharsPerLine
  );
  const balanced = rebalanceWrappedLines(wrapped, maxCharsPerLine);
  return balanced.join("\n");
}

export function buildCaptionPlan(cues: CaptionCue[], options: CaptionBuildOptions): CaptionCue[] {
  // // Split cues by contiguous word groups with timing proportional to word distribution.
  const plannedCues: CaptionCue[] = [];
  const maxCharsPerLine = Math.max(6, Number(options.style.maxCharsPerLine || 28));
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
      buildInitialChunkRanges(cueWordTexts, maxCharsPerLine, linesPerCaption),
      cueWordTexts,
      maxCharsPerLine,
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
        text: renderChunkText(chunkWords, maxCharsPerLine),
        words: chunkWords
      });
    });
  }

  return plannedCues;
}
