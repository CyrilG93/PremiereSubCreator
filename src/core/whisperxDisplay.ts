// // Normalize WhisperX-aligned cue timing before the generic caption planner creates timeline clips.
import { tokenizeSubtitleText } from "./textNormalization";
import type { CaptionCue } from "./types";

function countCueWords(cue: CaptionCue): number {
  // // Estimate how much text is visible so short-cue safeguards work even without word payloads.
  if (Array.isArray(cue.words) && cue.words.length > 0) {
    return cue.words.length;
  }
  return tokenizeSubtitleText(cue.text).length;
}

function getCueDurationSeconds(cue: CaptionCue): number {
  // // Clamp invalid durations to zero so merge decisions do not treat bad timestamps as long speech.
  return Math.max(0, Number(cue.endSeconds || 0) - Number(cue.startSeconds || 0));
}

function getCueGapSeconds(left: CaptionCue, right: CaptionCue): number {
  // // Measure the silence or overlap between two adjacent aligned cues.
  return Number(right.startSeconds || 0) - Number(left.endSeconds || 0);
}

function cueEndsWithHardBoundary(cue: CaptionCue): boolean {
  // // Avoid merging across sentence-like endings because short replies often end with these markers.
  return /[.!?…]$/.test(String(cue.text || "").trim());
}

function mergeCaptionCuesForDisplay(left: CaptionCue, right: CaptionCue): CaptionCue {
  // // Combine adjacent aligned cues only after the merge guard has accepted them.
  const leftText = String(left.text || "").trim();
  const rightText = String(right.text || "").trim();
  return {
    ...left,
    id: `${left.id}+${right.id}`,
    endSeconds: Math.max(Number(left.endSeconds || 0), Number(right.endSeconds || 0)),
    text: [leftText, rightText].filter(Boolean).join(" "),
    words: [...(left.words || []), ...(right.words || [])]
  };
}

function shouldMergeWhisperXCues(left: CaptionCue, right: CaptionCue): boolean {
  // // Merge only near-contiguous micro-cues; larger gaps usually mean a new phrase or another speaker.
  const gapSeconds = getCueGapSeconds(left, right);
  if (gapSeconds < -0.04 || gapSeconds > 0.14) {
    return false;
  }

  const leftWordCount = countCueWords(left);
  const rightWordCount = countCueWords(right);
  const bothGroupsAreTiny = leftWordCount <= 2 && rightWordCount <= 2;
  if (bothGroupsAreTiny && Math.abs(gapSeconds) > 0.02) {
    return false;
  }

  if (gapSeconds > 0.03 && cueEndsWithHardBoundary(left)) {
    return false;
  }

  const leftDuration = getCueDurationSeconds(left);
  const rightDuration = getCueDurationSeconds(right);
  return leftDuration < 0.72 || rightDuration < 0.72 || gapSeconds <= 0.03;
}

export function normalizeWhisperXCuesForDisplay(cues: CaptionCue[]): CaptionCue[] {
  // // WhisperX can be word-tight; pad cues for readability without over-merging short speaker turns.
  const sortedCues = cues
    .filter((cue) => String(cue.text || "").trim())
    .map((cue) => ({ ...cue, words: Array.isArray(cue.words) ? cue.words.map((word) => ({ ...word })) : [] }))
    .sort((left, right) => Number(left.startSeconds || 0) - Number(right.startSeconds || 0));
  const mergedCues: CaptionCue[] = [];
  const minDurationSeconds = 1.05;
  const leadPaddingSeconds = 0.08;
  const tailPaddingSeconds = 0.18;

  for (const cue of sortedCues) {
    const lastCue = mergedCues[mergedCues.length - 1];
    if (!lastCue) {
      mergedCues.push(cue);
      continue;
    }

    if (shouldMergeWhisperXCues(lastCue, cue)) {
      mergedCues[mergedCues.length - 1] = mergeCaptionCuesForDisplay(lastCue, cue);
    } else {
      mergedCues.push(cue);
    }
  }

  return mergedCues.map((cue, index) => {
    // // Expand each display cue into nearby silence, but never overlap neighboring cues.
    const previousCue = mergedCues[index - 1];
    const nextCue = mergedCues[index + 1];
    const minStart = previousCue ? Number(previousCue.endSeconds || 0) + 0.01 : 0;
    const maxEnd = nextCue ? Number(nextCue.startSeconds || 0) - 0.01 : Number.POSITIVE_INFINITY;
    const startSeconds = Math.max(minStart, Math.max(0, Number(cue.startSeconds || 0) - leadPaddingSeconds));
    const desiredEnd = Math.max(Number(cue.endSeconds || 0) + tailPaddingSeconds, startSeconds + minDurationSeconds);
    const endSeconds = Math.max(startSeconds + 0.12, Math.min(maxEnd, desiredEnd));
    return {
      ...cue,
      startSeconds,
      endSeconds
    };
  });
}
