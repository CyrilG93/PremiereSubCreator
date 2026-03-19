// // Parse SRT timestamps and cues into strongly typed caption entries.
import type { CaptionCue, CaptionWord } from "./types";
import { buildWeightedCaptionWordsFromText } from "./wordTiming";

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
  // // Build synthetic word timings from weighted text length so long words can occupy more of the cue duration.
  return buildWeightedCaptionWordsFromText(text, startSeconds, endSeconds);
}

function formatTimestamp(totalSeconds: number): string {
  // // Convert seconds back into SRT HH:MM:SS,mmm format so filtered corrected transcripts can be re-serialized.
  const safeSeconds = Math.max(0, Number(totalSeconds) || 0);
  const totalMilliseconds = Math.round(safeSeconds * 1000);
  const hours = Math.floor(totalMilliseconds / 3600000);
  const minutes = Math.floor((totalMilliseconds % 3600000) / 60000);
  const seconds = Math.floor((totalMilliseconds % 60000) / 1000);
  const milliseconds = totalMilliseconds % 1000;
  return [
    String(hours).padStart(2, "0"),
    String(minutes).padStart(2, "0"),
    String(seconds).padStart(2, "0")
  ].join(":") + `,${String(milliseconds).padStart(3, "0")}`;
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

export function shiftCaptionCues(cues: CaptionCue[], offsetSeconds: number): CaptionCue[] {
  // // Shift cue and word timings by a fixed offset so In/Out-generated captions land back at timeline time.
  const offset = Number(offsetSeconds) || 0;
  return cues.map((cue) => ({
    ...cue,
    startSeconds: cue.startSeconds + offset,
    endSeconds: cue.endSeconds + offset,
    words: cue.words.map((word) => ({
      ...word,
      startSeconds: word.startSeconds + offset,
      endSeconds: word.endSeconds + offset
    }))
  }));
}

export function trimSrtCuesToRange(cues: CaptionCue[], rangeStartSeconds: number, rangeEndSeconds: number): CaptionCue[] {
  // // Keep only cues that intersect the requested sequence range and clamp them so corrected align ignores the rest.
  const safeStart = Number(rangeStartSeconds);
  const safeEnd = Number(rangeEndSeconds);
  if (!Number.isFinite(safeStart) || !Number.isFinite(safeEnd) || safeEnd <= safeStart) {
    return [...cues];
  }

  return cues
    .filter((cue) => cue.endSeconds > safeStart && cue.startSeconds < safeEnd)
    .map((cue, index) => {
      const startSeconds = Math.max(cue.startSeconds, safeStart);
      const endSeconds = Math.min(cue.endSeconds, safeEnd);
      return {
        id: `trimmed-cue-${index + 1}`,
        startSeconds,
        endSeconds,
        text: cue.text,
        words: splitWordsWithTiming(cue.text, startSeconds, endSeconds)
      };
    });
}

export function serializeSrt(cues: CaptionCue[]): string {
  // // Serialize normalized cue data back into SRT text for filtered corrected-align inputs.
  return cues
    .map((cue, index) => `${index + 1}\n${formatTimestamp(cue.startSeconds)} ${TIMECODE_SEPARATOR} ${formatTimestamp(cue.endSeconds)}\n${cue.text}`)
    .join("\n\n");
}
