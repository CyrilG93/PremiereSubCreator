// // Parse SRT timestamps and cues into strongly typed caption entries.
import type { CaptionCue, CaptionWord } from "./types";

const TIMECODE_SEPARATOR = "-->";

function parseTimestamp(timestamp: string): number {
  // // Convert HH:MM:SS,mmm into seconds with millisecond precision.
  const cleaned = timestamp.trim().replace(",", ".");
  const parts = cleaned.split(":");
  if (parts.length !== 3) {
    throw new Error(`Invalid timestamp format: ${timestamp}`);
  }

  const hours = Number(parts[0]);
  const minutes = Number(parts[1]);
  const seconds = Number(parts[2]);

  if ([hours, minutes, seconds].some((value) => Number.isNaN(value))) {
    throw new Error(`Invalid numeric timestamp: ${timestamp}`);
  }

  return hours * 3600 + minutes * 60 + seconds;
}

function splitWordsWithTiming(text: string, startSeconds: number, endSeconds: number): CaptionWord[] {
  // // Distribute timing evenly per word to support word-level animation.
  const words = text
    .split(/\s+/)
    .map((value) => value.trim())
    .filter(Boolean);

  if (words.length === 0) {
    return [];
  }

  const totalDuration = Math.max(endSeconds - startSeconds, 0.01);
  const wordDuration = totalDuration / words.length;

  return words.map((word, index) => {
    const wordStart = startSeconds + index * wordDuration;
    const wordEnd = index === words.length - 1 ? endSeconds : wordStart + wordDuration;
    return {
      text: word,
      startSeconds: wordStart,
      endSeconds: wordEnd
    };
  });
}

export function parseSrt(input: string): CaptionCue[] {
  // // Parse an SRT file by blocks separated with blank lines.
  const normalized = input.replace(/\r\n/g, "\n").trim();
  if (!normalized) {
    return [];
  }

  const blocks = normalized.split(/\n\s*\n/);
  const cues: CaptionCue[] = [];

  for (let i = 0; i < blocks.length; i += 1) {
    const lines = blocks[i]
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    if (lines.length < 2) {
      continue;
    }

    const hasNumericId = /^\d+$/.test(lines[0]);
    const timeLine = hasNumericId ? lines[1] : lines[0];
    const textLines = hasNumericId ? lines.slice(2) : lines.slice(1);

    if (!timeLine.includes(TIMECODE_SEPARATOR) || textLines.length === 0) {
      continue;
    }

    const [startRaw, endRaw] = timeLine.split(TIMECODE_SEPARATOR);
    const startSeconds = parseTimestamp(startRaw);
    const endSeconds = parseTimestamp(endRaw);
    const text = textLines.join(" ").replace(/\s+/g, " ").trim();

    if (!text) {
      continue;
    }

    cues.push({
      id: `cue-${i + 1}`,
      startSeconds,
      endSeconds,
      text,
      words: splitWordsWithTiming(text, startSeconds, endSeconds)
    });
  }

  return cues;
}
