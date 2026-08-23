// // Define strongly typed structures shared by the panel and host bridge.
export type SourceMode = "srt" | "whisper_sequence" | "whisperx_sequence" | "corrected_align";
export type OutputMode = "mogrt" | "premiere_subtitles";
export type WhisperSequenceRangeMode = "entire_sequence" | "in_out";

// // Support both per-word and per-line animation strategies.
export type AnimationMode = "word" | "line" | "none";

// // Describe a single animated word timing segment.
export interface CaptionWord {
  text: string;
  startSeconds: number;
  endSeconds: number;
}

// // Describe one caption cue in timeline time.
export interface CaptionCue {
  id: string;
  startSeconds: number;
  endSeconds: number;
  text: string;
  words: CaptionWord[];
  mogrtPathOverride?: string;
  skipTextApply?: boolean;
}

// // Describe the style and animation configuration selected in the UI.
export interface CaptionStyleConfig {
  // // Keep backwards-compatible style preset metadata accepted by older tests or saved panel states.
  presetId?: string;
  maxCharsPerLine: number;
  maxWordsPerLine: number;
  animationMode: AnimationMode;
  uppercase: boolean;
  removePunctuation: boolean;
  linesPerCaption: number;
}

// // Carry raw Premiere-authored text-document payloads extracted from the `.mogrt` package so host-side writes can preserve style.
export interface PremiereTemplateTextPayload {
  displayName: string;
  initialText: string;
  sourcePayloadBase64: string;
  sourcePayloadXml?: string;
}

// // Describe full generation options for a build request.
export interface CaptionBuildOptions {
  sourceMode: SourceMode;
  outputMode: OutputMode;
  languageCode: string;
  style: CaptionStyleConfig;
  extensionRootPath: string;
  mogrtPath: string;
  mogrtTemplateRelativePath: string;
  correctedTranscriptPath?: string;
  whisperModel: string;
  whisperSequenceRange: WhisperSequenceRangeMode;
  preserveMixedLanguages: boolean;
  mixedLanguagePrompt: string;
  preserveTranslationTiming?: boolean;
  premiereTemplateTextPayloads?: PremiereTemplateTextPayload[];
  videoTrackIndex: number;
  audioTrackIndex: number;
}

// // Represent the payload sent to ExtendScript for timeline creation.
export interface HostApplyPayload {
  options: CaptionBuildOptions;
  cues: CaptionCue[];
}

// // Describe one discoverable MOGRT template displayed in the gallery.
export interface MogrtTemplateItem {
  id: string;
  name: string;
  aspect: string;
  relativePath: string;
  previewClass: string;
  previewImagePath?: string;
  previewVideoPath?: string;
}
