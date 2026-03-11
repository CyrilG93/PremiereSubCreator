// // Build subtitle cues from parsed content based on style constraints.
import type { CaptionBuildOptions, CaptionCue } from "./types";

function wrapTextByChars(text: string, maxCharsPerLine: number, linesPerCaption: number): string[] {
  // // Wrap words to keep each line under the configured character budget.
  const words = text.split(/\s+/).filter(Boolean);
  const lines: string[] = [];
  let currentLine = "";

  for (const word of words) {
    const candidate = currentLine ? `${currentLine} ${word}` : word;
    if (candidate.length <= maxCharsPerLine || !currentLine) {
      currentLine = candidate;
      continue;
    }

    lines.push(currentLine);
    currentLine = word;
  }

  if (currentLine) {
    lines.push(currentLine);
  }

  // // Group lines into caption chunks that respect the maximum lines per cue.
  const chunks: string[] = [];
  for (let i = 0; i < lines.length; i += linesPerCaption) {
    const slice = lines.slice(i, i + linesPerCaption);
    chunks.push(slice.join("\n"));
  }

  return chunks;
}

function uppercaseIfNeeded(text: string, forceUppercase: boolean): string {
  // // Optionally convert text to uppercase for aggressive social-media subtitle styles.
  return forceUppercase ? text.toUpperCase() : text;
}

export function buildCaptionPlan(cues: CaptionCue[], options: CaptionBuildOptions): CaptionCue[] {
  // // Split incoming cues by line-length and preserve timing proportionally.
  const plannedCues: CaptionCue[] = [];

  for (const cue of cues) {
    const wrapped = wrapTextByChars(
      uppercaseIfNeeded(cue.text, options.style.uppercase),
      options.style.maxCharsPerLine,
      options.style.linesPerCaption
    );

    if (wrapped.length <= 1) {
      plannedCues.push({
        ...cue,
        text: wrapped[0] ?? cue.text,
        words: cue.words.map((word) => ({
          ...word,
          text: uppercaseIfNeeded(word.text, options.style.uppercase)
        }))
      });
      continue;
    }

    const totalDuration = Math.max(cue.endSeconds - cue.startSeconds, 0.01);
    const chunkDuration = totalDuration / wrapped.length;

    wrapped.forEach((chunkText, index) => {
      const startSeconds = cue.startSeconds + index * chunkDuration;
      const endSeconds = index === wrapped.length - 1 ? cue.endSeconds : startSeconds + chunkDuration;
      plannedCues.push({
        id: `${cue.id}-part-${index + 1}`,
        startSeconds,
        endSeconds,
        text: chunkText,
        words: []
      });
    });
  }

  return plannedCues;
}
