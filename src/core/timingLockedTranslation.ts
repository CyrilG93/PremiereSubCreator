// // Keep forced-language Whisper text on the timing obtained from the original audio transcription.
import { tokenizeSubtitleText } from "./textNormalization";
import type { CaptionCue } from "./types";
import { buildWeightedCaptionWords } from "./wordTiming";

export function lockTranslatedCuesToSourceTiming(sourceCues: CaptionCue[], translatedCues: CaptionCue[]): CaptionCue[] {
  // // Redistribute translated words across the source cue durations, preserving every source start and end exactly.
  const translatedWords = translatedCues.flatMap((cue) => tokenizeSubtitleText(cue.text));
  if (sourceCues.length < 1 || translatedWords.length < 1) {
    return [];
  }

  const totalDuration = Math.max(
    sourceCues.reduce((sum, cue) => sum + Math.max(0.01, cue.endSeconds - cue.startSeconds), 0),
    0.01
  );
  const canGiveEveryCueAWord = translatedWords.length >= sourceCues.length;
  let consumedDuration = 0;
  let wordCursor = 0;

  return sourceCues.flatMap((sourceCue, index) => {
    // // Allocate words by elapsed source duration so translated text remains in chronological order even when Whisper retimes it.
    consumedDuration += Math.max(0.01, sourceCue.endSeconds - sourceCue.startSeconds);
    const remainingCues = sourceCues.length - index - 1;
    const proportionalEnd = Math.round((consumedDuration / totalDuration) * translatedWords.length);
    const minimumEnd = wordCursor + (canGiveEveryCueAWord ? 1 : 0);
    const maximumEnd = translatedWords.length - (canGiveEveryCueAWord ? remainingCues : 0);
    const wordEnd = index === sourceCues.length - 1 ? translatedWords.length : Math.min(maximumEnd, Math.max(minimumEnd, proportionalEnd));
    const words = translatedWords.slice(wordCursor, wordEnd);
    wordCursor = wordEnd;
    if (words.length < 1) {
      return [];
    }

    return [{
      id: `${sourceCue.id}-timing-locked`,
      startSeconds: sourceCue.startSeconds,
      endSeconds: sourceCue.endSeconds,
      text: words.join(" "),
      words: buildWeightedCaptionWords(words, sourceCue.startSeconds, sourceCue.endSeconds)
    }];
  });
}
