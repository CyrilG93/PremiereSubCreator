// // Parse Whisper JSON output so caption planning can reuse exact word timestamps when available.
import type { CaptionCue, CaptionWord } from "./types";
import { buildWeightedCaptionWordsFromText } from "./wordTiming";

interface WhisperWordPayload {
  word?: unknown;
  start?: unknown;
  end?: unknown;
}

interface WhisperSegmentPayload {
  id?: unknown;
  text?: unknown;
  start?: unknown;
  end?: unknown;
  words?: unknown;
}

function normalizeWhisperText(value: unknown): string {
  // // Normalize Whisper text fragments into clean single-space strings for cue and word rendering.
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeWhisperWord(word: WhisperWordPayload): CaptionWord | null {
  // // Convert one Whisper word payload into a caption word only when timestamps are complete and valid.
  const text = normalizeWhisperText(word.word);
  const startSeconds = Number(word.start);
  const endSeconds = Number(word.end);
  if (!text || !Number.isFinite(startSeconds) || !Number.isFinite(endSeconds) || endSeconds < startSeconds) {
    return null;
  }

  return {
    text,
    startSeconds,
    endSeconds
  };
}

export function parseWhisperJson(input: string): CaptionCue[] {
  // // Parse Whisper JSON segments and preserve word-level timing whenever the CLI emitted it.
  const normalized = String(input || "").trim();
  if (!normalized) {
    return [];
  }

  const payload = JSON.parse(normalized) as { segments?: unknown };
  const segments = Array.isArray(payload?.segments) ? (payload.segments as WhisperSegmentPayload[]) : [];
  const cues: CaptionCue[] = [];

  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index];
    const segmentStart = Number(segment.start);
    const segmentEnd = Number(segment.end);
    const segmentText = normalizeWhisperText(segment.text);
    if (!Number.isFinite(segmentStart) || !Number.isFinite(segmentEnd) || segmentEnd < segmentStart) {
      continue;
    }

    const words = Array.isArray(segment.words)
      ? segment.words
          .map((word) => normalizeWhisperWord((word as WhisperWordPayload) || {}))
          .filter((word): word is CaptionWord => Boolean(word))
      : [];

    const cueText =
      segmentText ||
      words
        .map((word) => word.text)
        .join(" ")
        .replace(/\s+/g, " ")
        .trim();
    if (!cueText) {
      continue;
    }

    cues.push({
      id: `whisper-${index + 1}`,
      startSeconds: words.length > 0 ? words[0].startSeconds : segmentStart,
      endSeconds: words.length > 0 ? words[words.length - 1].endSeconds : segmentEnd,
      text: cueText,
      words: words.length > 0 ? words : buildWeightedCaptionWordsFromText(cueText, segmentStart, segmentEnd)
    });
  }

  return cues;
}
