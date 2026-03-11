// // Define strongly typed structures shared by the panel and host bridge.
export type SourceMode = "srt" | "transcription";

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
}

// // Describe the style and animation configuration selected in the UI.
export interface CaptionStyleConfig {
  presetId: string;
  fontSize: number;
  maxCharsPerLine: number;
  animationMode: AnimationMode;
  uppercase: boolean;
  linesPerCaption: number;
}

// // Describe full generation options for a build request.
export interface CaptionBuildOptions {
  sourceMode: SourceMode;
  languageCode: string;
  style: CaptionStyleConfig;
  mogrtPath: string;
  videoTrackIndex: number;
  audioTrackIndex: number;
}

// // Represent the payload sent to ExtendScript for timeline creation.
export interface HostApplyPayload {
  options: CaptionBuildOptions;
  cues: CaptionCue[];
}

// // Define a user-selectable style preset with visual defaults.
export interface StylePreset {
  id: string;
  labelKey: string;
  defaultFontSize: number;
  defaultMaxCharsPerLine: number;
  defaultAnimationMode: AnimationMode;
}
