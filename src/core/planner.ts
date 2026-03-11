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

    let cursor = 0;
    let chunkIndex = 0;
    while (cursor < normalizedWords.length) {
      const chunkEnd = findChunkEndIndex(
        normalizedWords.map((word) => word.text),
        cursor,
        maxCharsPerLine,
        linesPerCaption
      );
      const safeEnd = Math.max(cursor + 1, chunkEnd);
      const chunkWords = normalizedWords.slice(cursor, safeEnd);
      const firstWord = chunkWords[0];
      const lastWord = chunkWords[chunkWords.length - 1];
      const startSeconds = chunkIndex === 0 ? cue.startSeconds : firstWord.startSeconds;
      const endSeconds = safeEnd >= normalizedWords.length ? cue.endSeconds : lastWord.endSeconds;

      plannedCues.push({
        id: `${cue.id}-part-${chunkIndex + 1}`,
        startSeconds,
        endSeconds,
        text: renderChunkText(chunkWords, maxCharsPerLine),
        words: chunkWords
      });
      cursor = safeEnd;
      chunkIndex += 1;
    }
  }

  return plannedCues;
}
