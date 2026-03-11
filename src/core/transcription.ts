// // Keep transcription support isolated so we can swap implementation later.
import type { CaptionCue } from "./types";

export interface TranscriptionResult {
  cues: CaptionCue[];
  provider: "premiere" | "external";
  warning?: string;
}

export async function transcribeActiveSequence(languageCode: string): Promise<TranscriptionResult> {
  // // Return a clear placeholder until Adobe exposes stable scripting hooks for transcript extraction.
  return {
    cues: [],
    provider: "premiere",
    warning: `Transcription active indisponible pour le moment (langue demandee: ${languageCode}). Utilise la source SRT pour generer maintenant.`
  };
}
