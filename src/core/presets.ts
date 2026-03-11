// // Expose style presets inspired by dynamic subtitle workflows.
import type { StylePreset } from "./types";

export const STYLE_PRESETS: StylePreset[] = [
  {
    id: "clean",
    labelKey: "preset.clean",
    defaultFontSize: 78,
    defaultMaxCharsPerLine: 28,
    defaultAnimationMode: "line"
  },
  {
    id: "punch",
    labelKey: "preset.punch",
    defaultFontSize: 92,
    defaultMaxCharsPerLine: 18,
    defaultAnimationMode: "word"
  },
  {
    id: "minimal",
    labelKey: "preset.minimal",
    defaultFontSize: 64,
    defaultMaxCharsPerLine: 34,
    defaultAnimationMode: "none"
  }
];
