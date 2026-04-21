// // Build synthetic per-word timings that favor longer words and punctuation pauses over uniform timing.
import type { CaptionWord } from "./types";
import { tokenizeSubtitleText } from "./textNormalization";

function normalizeWeightedTimingWord(word: string): string {
  // // Strip punctuation around a word so length-based weighting follows spoken content instead of commas or quotes.
  return String(word || "")
    .trim()
    .replace(/^[^\p{L}\p{N}]+/gu, "")
    .replace(/[^\p{L}\p{N}]+$/gu, "");
}

function estimateWeightedTimingScore(word: string): number {
  // // Give slightly more duration to longer words and to words that naturally carry a pause at the end.
  const trimmedWord = String(word || "").trim();
  const normalizedWord = normalizeWeightedTimingWord(trimmedWord);
  const lengthScore = Math.max(1, normalizedWord.length || 1);

  let weight = 0.85 + Math.pow(Math.min(lengthScore, 16), 0.8) * 0.24;
  if (/[,;:]$/.test(trimmedWord)) {
    weight += 0.35;
  }
  if (/[.!?…]$/.test(trimmedWord)) {
    weight += 0.55;
  }

  return Math.max(0.5, Number(weight.toFixed(6)));
}

export function buildWeightedCaptionWords(words: string[], startSeconds: number, endSeconds: number): CaptionWord[] {
  // // Distribute one cue duration across words using weighted scores instead of uniform per-word slices.
  const normalizedWords = words.map((word) => String(word || "").trim()).filter(Boolean);
  if (normalizedWords.length < 1) {
    return [];
  }

  const totalDuration = Math.max(endSeconds - startSeconds, 0.01);
  const weights = normalizedWords.map((word) => estimateWeightedTimingScore(word));
  const totalWeight = Math.max(
    weights.reduce((sum, weight) => sum + weight, 0),
    0.000001
  );

  let cursor = startSeconds;
  return normalizedWords.map((word, index) => {
    const sliceDuration = totalDuration * (weights[index] / totalWeight);
    const wordStart = cursor;
    const wordEnd = index === normalizedWords.length - 1 ? endSeconds : Math.min(endSeconds, wordStart + sliceDuration);
    cursor = wordEnd;
    return {
      text: word,
      startSeconds: wordStart,
      endSeconds: wordEnd
    };
  });
}

export function buildWeightedCaptionWordsFromText(text: string, startSeconds: number, endSeconds: number): CaptionWord[] {
  // // Split a cue text into words, then assign weighted synthetic timings across the cue duration.
  const words = tokenizeSubtitleText(text);
  return buildWeightedCaptionWords(words, startSeconds, endSeconds);
}
