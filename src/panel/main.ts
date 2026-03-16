// // Drive the Sub Creator panel UI and connect it to subtitle generation logic.
import { buildCaptionPlan } from "../core/planner";
import { parseWhisperJson } from "../core/whisper";
import { parseSrt } from "../core/srt";
import type {
  AnimationMode,
  CaptionBuildOptions,
  CaptionCue,
  HostApplyPayload,
  MogrtTemplateItem,
  SourceMode,
  WhisperSequenceRangeMode
} from "../core/types";
import {
  applyCaptionPlan,
  applyVisualPropertiesToSelectedMogrts,
  deleteTemporaryWhisperAudio,
  exportActiveSequenceAudioForWhisper,
  getSelectedMogrtCount,
  getWhisperRuntimeStatus,
  openExternalUrl,
  openInstalledMogrtFolder,
  pickSrtPath,
  readInstalledMogrtCatalog,
  readSelectedMogrtVisualProperties,
  readSystemFontCatalog,
  readTextFileFromHost,
  transcribeWithWhisper
} from "./cepBridge";
import type { InstalledMogrtCatalog, SystemFontCatalog, WhisperProgressUpdate } from "./cepBridge";

type LocaleMap = Record<string, string>;

interface MogrtCatalog {
  generatedAt: string;
  templateCount: number;
  templates: MogrtTemplateItem[];
}

interface PanelMeta {
  version: string;
  repository: string;
  releaseApiUrl: string;
  releasePageUrl: string;
}

interface UpdateState {
  visible: boolean;
  latestVersion: string;
  downloadUrl: string;
}

type PanelMode = "generate" | "visual";

interface HostVisualProperty {
  path: string;
  displayName: string;
  groupPath: string;
  valueType: "number" | "boolean" | "string" | "json";
  controlKind: "slider" | "number" | "checkbox" | "color" | "text" | "string" | "json" | "vector" | "select";
  options?: Array<{ value: number | string; label: string }>;
  styleOptionsByFamily?: Record<string, string[]>;
  vectorScale?: number[];
  vectorMode?: string;
  value: string | number | boolean;
  minValue?: number;
  maxValue?: number;
  stepValue?: number;
}

interface PanelStateSnapshot {
  languageCode: string;
  activeMode: PanelMode;
  sourceMode: SourceMode;
  srtPath: string;
  whisperModel: string;
  whisperSequenceRange: WhisperSequenceRangeMode;
  animationMode: AnimationMode;
  maxCharsPerLine: number;
  linesPerCaption: number;
  fontSize: number;
  mogrtAspectFilter: string;
  selectedMogrtId: string;
  visualLiveUpdate: boolean;
  logExpanded: boolean;
  verboseLogs: boolean;
}

interface RgbColor {
  red: number;
  green: number;
  blue: number;
}

interface CepHostSkinInfo {
  baseFontFamily?: string;
  panelBackgroundColor?: Partial<RgbColor>;
  panelBackgroundColorSRGB?: Partial<RgbColor>;
  systemHighlightColor?: Partial<RgbColor>;
}

interface CepHostEnvironment {
  appSkinInfo?: CepHostSkinInfo;
}

const elements = {
  languageSelect: document.querySelector<HTMLSelectElement>("#languageSelect"),
  appVersion: document.querySelector<HTMLSpanElement>("#appVersion"),
  updateBanner: document.querySelector<HTMLElement>("#updateBanner"),
  updateLink: document.querySelector<HTMLAnchorElement>("#updateLink"),
  tabGenerate: document.querySelector<HTMLButtonElement>("#tabGenerate"),
  tabVisual: document.querySelector<HTMLButtonElement>("#tabVisual"),
  modeGenerate: document.querySelector<HTMLElement>("#modeGenerate"),
  modeVisual: document.querySelector<HTMLElement>("#modeVisual"),
  sourceMode: document.querySelector<HTMLSelectElement>("#sourceMode"),
  srtInputField: document.querySelector<HTMLElement>("#srtInputField"),
  srtPath: document.querySelector<HTMLInputElement>("#srtPath"),
  srtBrowseButton: document.querySelector<HTMLButtonElement>("#srtBrowseButton"),
  whisperField: document.querySelector<HTMLElement>("#whisperField"),
  whisperModelRow: document.querySelector<HTMLElement>("#whisperModelRow"),
  whisperModel: document.querySelector<HTMLSelectElement>("#whisperModel"),
  whisperSequenceRange: document.querySelector<HTMLSelectElement>("#whisperSequenceRange"),
  whisperSequenceHint: document.querySelector<HTMLElement>("#whisperSequenceHint"),
  animationMode: document.querySelector<HTMLSelectElement>("#animationMode"),
  maxChars: document.querySelector<HTMLInputElement>("#maxChars"),
  linesPerCaption: document.querySelector<HTMLInputElement>("#linesPerCaption"),
  fontSize: document.querySelector<HTMLInputElement>("#fontSize"),
  mogrtAspectFilter: document.querySelector<HTMLSelectElement>("#mogrtAspectFilter"),
  mogrtFolderButton: document.querySelector<HTMLButtonElement>("#mogrtFolderButton"),
  mogrtRefreshButton: document.querySelector<HTMLButtonElement>("#mogrtRefreshButton"),
  mogrtGallery: document.querySelector<HTMLElement>("#mogrtGallery"),
  mogrtSelectedLabel: document.querySelector<HTMLParagraphElement>("#mogrtSelectedLabel"),
  visualReadButton: document.querySelector<HTMLButtonElement>("#visualReadButton"),
  visualApplyButton: document.querySelector<HTMLButtonElement>("#visualApplyButton"),
  visualLiveUpdateButton: document.querySelector<HTMLButtonElement>("#visualLiveUpdateButton"),
  visualApplyProgress: document.querySelector<HTMLElement>("#visualApplyProgress"),
  visualApplyProgressBar: document.querySelector<HTMLProgressElement>("#visualApplyProgressBar"),
  visualApplyProgressText: document.querySelector<HTMLElement>("#visualApplyProgressText"),
  visualSelectionSummary: document.querySelector<HTMLParagraphElement>("#visualSelectionSummary"),
  visualPropertyList: document.querySelector<HTMLElement>("#visualPropertyList"),
  generateButton: document.querySelector<HTMLButtonElement>("#generateButton"),
  generateProgress: document.querySelector<HTMLElement>("#generateProgress"),
  generateProgressBar: document.querySelector<HTMLProgressElement>("#generateProgressBar"),
  generateProgressText: document.querySelector<HTMLElement>("#generateProgressText"),
  logPanel: document.querySelector<HTMLElement>("#logPanel"),
  logToggleButton: document.querySelector<HTMLButtonElement>("#logToggleButton"),
  logVerbosityButton: document.querySelector<HTMLButtonElement>("#logVerbosityButton"),
  logOutput: document.querySelector<HTMLPreElement>("#logOutput")
};

interface PanelLogState {
  plainText: string;
  structuredTitle: string;
  structuredPayload: unknown;
  isError: boolean;
}

let currentLocale: LocaleMap = {};
let availableMogrts: MogrtTemplateItem[] = [];
let selectedMogrt: MogrtTemplateItem | null = null;
let pendingMogrtAspectFilter = "";
const FALLBACK_PANEL_META: PanelMeta = {
  version: "0.0.0",
  repository: "CyrilG93/PremiereSubCreator",
  releaseApiUrl: "https://api.github.com/repos/CyrilG93/PremiereSubCreator/releases/latest",
  releasePageUrl: "https://github.com/CyrilG93/PremiereSubCreator/releases/latest"
};
let panelMeta: PanelMeta = { ...FALLBACK_PANEL_META };
const updateState: UpdateState = {
  visible: false,
  latestVersion: "",
  downloadUrl: ""
};
const PANEL_STATE_STORAGE_KEY = "subcreator.panelState.v1";
let pendingSelectedMogrtId = "";
let activeMode: PanelMode = "generate";
let loadedVisualProperties: HostVisualProperty[] = [];
let loadedVisualPropertySignature = "";
const visualOriginalValuesByPath = new Map<string, string>();
const visualOpenGroups = new Set<string>();
const visualTextStyleTokenMapByBasePath = new Map<string, Record<string, string>>();
let visualLiveUpdateTimer: number | null = null;
let visualLiveUpdateQueued = false;
let visualLiveUpdateInFlight = false;
let visualApplyInProgress = false;
let visualLiveUpdateEnabled = false;
let systemFontCatalogLoadPromise: Promise<void> | null = null;
let logPanelExpanded = true;
let verboseLogsEnabled = false;
let currentLogState: PanelLogState | null = null;
let passiveMogrtRefreshTimer: number | null = null;
let lastPassiveMogrtCatalogRefreshAt = 0;
let generateInProgress = false;
let hostThemeListenerBound = false;
let systemFontCatalog: SystemFontCatalog = {
  available: false,
  source: "unavailable",
  details: "",
  families: [],
  stylesByFamily: {},
  fontTokensByFamilyStyle: {}
};

const CEP_THEME_COLOR_CHANGED_EVENT = "com.adobe.csxs.events.ThemeColorChanged";
const GENERATE_PROGRESS_MAX = 100;

function parseTextStyleVirtualPath(path: string): { basePath: string; styleKey: string } | null {
  // // Decode synthetic text-style editor paths like `4::textstyle.fontStyle`.
  const marker = "::textstyle.";
  const markerIndex = String(path || "").indexOf(marker);
  if (markerIndex <= 0) {
    return null;
  }
  const basePath = String(path || "").slice(0, markerIndex).trim();
  const styleKey = String(path || "").slice(markerIndex + marker.length).trim();
  if (!basePath || !styleKey) {
    return null;
  }
  return { basePath, styleKey };
}

function normalizeFontLookupKey(value: string): string {
  // // Normalize font display values for case/spacing-insensitive lookups.
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function normalizeCompactFontLookupKey(value: string): string {
  // // Normalize font values for alias matches like `Al Bayan` vs `AlBayan`.
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function listFontFamilyLookupKeys(value: string): string[] {
  // // Build all lookup keys used for family alias matching in style/token maps.
  const keys = [normalizeFontLookupKey(value), normalizeCompactFontLookupKey(value)].filter(Boolean);
  return Array.from(new Set(keys));
}

function listFontStyleLookupKeys(value: string): string[] {
  // // Build style lookup aliases so `Plain` and `Regular` resolve to the same token bucket.
  const normalized = normalizeFontLookupKey(value);
  const compact = normalizeCompactFontLookupKey(value);
  const keys = [normalized, compact].filter(Boolean);
  if (normalized === "regular") {
    keys.push("plain", "roman");
  }
  if (normalized === "plain" || normalized === "roman") {
    keys.push("regular");
  }
  return Array.from(new Set(keys.filter(Boolean)));
}

function normalizeFontStyleDisplayKey(value: string): string {
  // // Collapse display-only style aliases to one select matching key.
  const normalized = normalizeFontLookupKey(value);
  if (normalized === "plain" || normalized === "roman") {
    return "regular";
  }
  return normalized;
}

function findMatchingFontStyleOption(options: string[], targetStyle: string): string {
  // // Resolve one requested style against available options using the same alias rules as token lookup.
  const requestedKeys = new Set(listFontStyleLookupKeys(targetStyle));
  if (requestedKeys.size < 1) {
    return "";
  }
  for (const option of options) {
    const optionKeys = listFontStyleLookupKeys(option);
    if (optionKeys.some((entry) => requestedKeys.has(entry))) {
      return option;
    }
  }
  return "";
}

function getFontStylePriority(style: string): number {
  // // Rank neutral/default styles before heavier variants so family switches land on safer defaults.
  const normalizedKeys = new Set(listFontStyleLookupKeys(style));
  const priorityBuckets = [
    ["regular", "plain", "roman"],
    ["book"],
    ["medium"],
    ["semibold", "demibold"],
    ["bold"],
    ["italic", "oblique"],
    ["bold italic", "bolditalic"],
    ["black", "heavy"],
    ["extrabold", "ultrabold"]
  ];
  for (let index = 0; index < priorityBuckets.length; index += 1) {
    if (priorityBuckets[index].some((entry) => normalizedKeys.has(entry))) {
      return index;
    }
  }
  return priorityBuckets.length;
}

function sortFontStyleOptions(options: string[]): string[] {
  // // Keep style lists stable and neutral-first so dropdown defaults stay predictable across families.
  return options
    .map((entry) => String(entry || "").trim())
    .filter(Boolean)
    .sort((left, right) => {
      const priorityDelta = getFontStylePriority(left) - getFontStylePriority(right);
      if (priorityDelta !== 0) {
        return priorityDelta;
      }
      return left.localeCompare(right, undefined, { sensitivity: "base" });
    });
}

function pickPreferredFontStyleOption(options: string[], preferredStyles?: string[]): string {
  // // Pick the best style for a family, preferring neutral defaults like `Regular` before any fallback.
  const normalizedOptions = sortFontStyleOptions(options);
  if (normalizedOptions.length < 1) {
    return "";
  }
  const requestedStyles =
    preferredStyles && preferredStyles.length > 0
      ? preferredStyles
      : ["Regular", "Book", "Roman", "Plain", "Medium", "Semibold"];
  for (const requestedStyle of requestedStyles) {
    const matchedOption = findMatchingFontStyleOption(normalizedOptions, requestedStyle);
    if (matchedOption) {
      return matchedOption;
    }
  }
  return normalizedOptions[0];
}

function assertDomBindings(): void {
  // // Guard against missing panel DOM ids during development/build changes.
  const missing = Object.entries(elements)
    .filter(([, value]) => !value)
    .map(([key]) => key);

  if (missing.length > 0) {
    throw new Error(`Missing DOM elements: ${missing.join(", ")}`);
  }
}

function clampColorChannel(value: unknown): number {
  // // Normalize unknown numeric channel values into valid RGB integer channels.
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) {
    return 0;
  }

  return Math.max(0, Math.min(255, Math.round(numericValue)));
}

function readRgbColor(value: unknown): RgbColor | null {
  // // Parse CEP skin color payloads which expose `{ red, green, blue }` channel objects.
  if (!value || typeof value !== "object") {
    return null;
  }

  const payload = value as Record<string, unknown>;
  return {
    red: clampColorChannel(payload.red),
    green: clampColorChannel(payload.green),
    blue: clampColorChannel(payload.blue)
  };
}

function mixRgbColor(left: RgbColor, right: RgbColor, rightWeight: number): RgbColor {
  // // Blend two colors so panel surfaces can stay close to Premiere base shades.
  const clampedWeight = Math.max(0, Math.min(1, rightWeight));
  const leftWeight = 1 - clampedWeight;
  return {
    red: clampColorChannel(left.red * leftWeight + right.red * clampedWeight),
    green: clampColorChannel(left.green * leftWeight + right.green * clampedWeight),
    blue: clampColorChannel(left.blue * leftWeight + right.blue * clampedWeight)
  };
}

function offsetRgbColor(color: RgbColor, delta: number): RgbColor {
  // // Brighten or darken one color uniformly while keeping channels inside RGB bounds.
  return {
    red: clampColorChannel(color.red + delta),
    green: clampColorChannel(color.green + delta),
    blue: clampColorChannel(color.blue + delta)
  };
}

function normalizeHostPanelBackground(color: RgbColor): RgbColor {
  // // Keep Premiere skin variants readable while still following host light/dark/darkest appearance changes.
  const luminance = rgbColorLuminance(color);
  if (luminance <= 0.16) {
    return mixRgbColor(color, { red: 52, green: 52, blue: 52 }, 0.9);
  }
  if (luminance <= 0.32) {
    return mixRgbColor(color, { red: 58, green: 58, blue: 58 }, 0.42);
  }
  if (luminance >= 0.7) {
    return mixRgbColor(color, { red: 198, green: 198, blue: 198 }, 0.24);
  }
  if (luminance >= 0.55) {
    return mixRgbColor(color, { red: 184, green: 184, blue: 184 }, 0.12);
  }

  return color;
}

function rgbColorLuminance(color: RgbColor): number {
  // // Estimate perceived brightness to switch between light/dark Premiere skin variants.
  return (0.2126 * color.red + 0.7152 * color.green + 0.0722 * color.blue) / 255;
}

function setRootRgbVariable(variableName: string, color: RgbColor): void {
  // // Publish one RGB color both as `rgb(...)` and raw `r, g, b` triplet for CSS reuse.
  const root = document.documentElement;
  root.style.setProperty(variableName, `rgb(${color.red}, ${color.green}, ${color.blue})`);
  root.style.setProperty(`${variableName}-rgb`, `${color.red}, ${color.green}, ${color.blue}`);
}

function readHostEnvironmentSkin(): CepHostEnvironment | null {
  // // Read Premiere CEP host skin information without depending on an external CSInterface bundle.
  const rawEnvironment = window.__adobe_cep__?.getHostEnvironment?.();
  if (!rawEnvironment) {
    return null;
  }

  try {
    const parsed = JSON.parse(rawEnvironment) as CepHostEnvironment;
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function applyHostPanelTheme(): void {
  // // Derive neutral Premiere-like surfaces from the current CEP skin and react to host theme changes.
  const hostEnvironment = readHostEnvironmentSkin();
  if (!hostEnvironment?.appSkinInfo) {
    return;
  }

  const skinInfo = hostEnvironment.appSkinInfo;
  const panelBackground =
    readRgbColor(skinInfo.panelBackgroundColorSRGB) ||
    readRgbColor(skinInfo.panelBackgroundColor) || {
      red: 48,
      green: 48,
      blue: 48
    };
  const highlightColor = readRgbColor(skinInfo.systemHighlightColor) || {
    red: 70,
    green: 137,
    blue: 255
  };
  const hostLuminance = rgbColorLuminance(panelBackground);
  const normalizedPanelBackground = normalizeHostPanelBackground(panelBackground);
  const isLightTheme = hostLuminance >= 0.55;
  const isDarkestTheme = hostLuminance <= 0.18;
  const textPrimary = isLightTheme
    ? { red: 36, green: 36, blue: 36 }
    : { red: 236, green: 236, blue: 236 };
  const textDim = mixRgbColor(textPrimary, normalizedPanelBackground, isLightTheme ? 0.48 : 0.38);
  const bgPrimary = normalizedPanelBackground;
  const bgSurface = offsetRgbColor(normalizedPanelBackground, isLightTheme ? 8 : isDarkestTheme ? 8 : 6);
  const bgSoft = offsetRgbColor(normalizedPanelBackground, isLightTheme ? 13 : isDarkestTheme ? 12 : 10);
  const bgInput = offsetRgbColor(normalizedPanelBackground, isLightTheme ? -3 : isDarkestTheme ? -2 : -1);
  const bgCard = offsetRgbColor(normalizedPanelBackground, isLightTheme ? 5 : isDarkestTheme ? 5 : 4);
  const accent = mixRgbColor(highlightColor, normalizedPanelBackground, 0.15);
  const accentSoft = mixRgbColor(highlightColor, textPrimary, isLightTheme ? 0.22 : 0.18);
  const border = offsetRgbColor(normalizedPanelBackground, isLightTheme ? -28 : 16);
  const borderStrong = offsetRgbColor(normalizedPanelBackground, isLightTheme ? -42 : 24);
  const buttonPrimary = mixRgbColor(accent, bgSurface, 0.28);
  const buttonPrimaryAlt = mixRgbColor(accent, bgPrimary, 0.2);
  const buttonPrimaryText = rgbColorLuminance(buttonPrimary) >= 0.5
    ? { red: 16, green: 16, blue: 16 }
    : { red: 246, green: 246, blue: 246 };
  const root = document.documentElement;

  root.dataset.themeVariant = isLightTheme ? "light" : isDarkestTheme ? "darkest" : "dark";
  setRootRgbVariable("--bg-primary", bgPrimary);
  setRootRgbVariable("--bg-surface", bgSurface);
  setRootRgbVariable("--bg-soft", bgSoft);
  setRootRgbVariable("--bg-input", bgInput);
  setRootRgbVariable("--bg-card", bgCard);
  setRootRgbVariable("--text-primary", textPrimary);
  setRootRgbVariable("--text-dim", textDim);
  setRootRgbVariable("--accent", accent);
  setRootRgbVariable("--accent-soft", accentSoft);
  setRootRgbVariable("--border", border);
  setRootRgbVariable("--border-strong", borderStrong);
  setRootRgbVariable("--button-primary-bg", buttonPrimary);
  setRootRgbVariable("--button-primary-bg-alt", buttonPrimaryAlt);
  setRootRgbVariable("--button-primary-text", buttonPrimaryText);
  root.style.setProperty("--shadow", isLightTheme ? "0 6px 18px rgba(0, 0, 0, 0.08)" : "0 6px 18px rgba(0, 0, 0, 0.24)");

  const baseFontFamily = String(skinInfo.baseFontFamily || "").trim();
  if (baseFontFamily) {
    root.style.setProperty("--ui-font-family", `"${baseFontFamily}", "Avenir Next", "Helvetica Neue", sans-serif`);
  }
}

function bindHostThemeListener(): void {
  // // Subscribe once to CEP theme changes so the panel follows Premiere appearance switches live.
  if (hostThemeListenerBound || typeof window.__adobe_cep__?.addEventListener !== "function") {
    return;
  }

  hostThemeListenerBound = true;
  window.__adobe_cep__.addEventListener(CEP_THEME_COLOR_CHANGED_EVENT, () => {
    applyHostPanelTheme();
  });
}

function waitForNextPaint(): Promise<void> {
  // // Yield one frame so progress-bar/state changes paint before lengthy async work continues.
  return new Promise((resolve) => {
    window.requestAnimationFrame(() => resolve());
  });
}

function translate(key: string): string {
  // // Resolve translated labels and fallback to key when missing.
  return currentLocale[key] ?? key;
}

function translateTemplate(key: string, values: Record<string, string>): string {
  // // Apply simple token replacement for localized UI messages.
  const base = translate(key);
  return Object.keys(values).reduce((output, token) => {
    const matcher = new RegExp(`\\{${token}\\}`, "g");
    return output.replace(matcher, values[token]);
  }, base);
}

function hasSelectOption(select: HTMLSelectElement | null | undefined, value: string): boolean {
  // // Validate that a select option exists before restoring persisted values.
  if (!select || !value) {
    return false;
  }

  return Array.from(select.options).some((option) => option.value === value);
}

function readPersistedPanelState(): Partial<PanelStateSnapshot> {
  // // Restore the previous panel configuration from localStorage when available.
  try {
    const raw = window.localStorage.getItem(PANEL_STATE_STORAGE_KEY);
    if (!raw) {
      return {};
    }

    const parsed = JSON.parse(raw) as Partial<PanelStateSnapshot>;
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function persistPanelState(): void {
  // // Persist current panel configuration so reopening keeps user preferences.
  if (
    !elements.languageSelect ||
    !elements.tabGenerate ||
    !elements.sourceMode ||
    !elements.srtPath ||
    !elements.whisperModel ||
    !elements.whisperSequenceRange ||
    !elements.animationMode ||
    !elements.maxChars ||
    !elements.linesPerCaption ||
    !elements.fontSize ||
    !elements.mogrtAspectFilter
  ) {
    return;
  }

  const snapshot: PanelStateSnapshot = {
    languageCode: elements.languageSelect.value || "en",
    activeMode,
    sourceMode: getSourceMode(),
    srtPath: elements.srtPath.value || "",
    whisperModel: elements.whisperModel.value || "base",
    whisperSequenceRange: (elements.whisperSequenceRange.value as WhisperSequenceRangeMode) || "entire_sequence",
    animationMode: (elements.animationMode.value as AnimationMode) || "line",
    maxCharsPerLine: Number(elements.maxChars.value),
    linesPerCaption: Number(elements.linesPerCaption.value),
    fontSize: Number(elements.fontSize.value),
    mogrtAspectFilter: pendingMogrtAspectFilter || elements.mogrtAspectFilter.value || "all",
    selectedMogrtId: selectedMogrt?.id || pendingSelectedMogrtId || "",
    visualLiveUpdate: visualLiveUpdateEnabled,
    logExpanded: logPanelExpanded,
    verboseLogs: verboseLogsEnabled
  };

  try {
    window.localStorage.setItem(PANEL_STATE_STORAGE_KEY, JSON.stringify(snapshot));
  } catch {
    // // Ignore storage errors to keep panel usable in restricted environments.
  }
}

function applyPersistedPanelState(snapshot: Partial<PanelStateSnapshot>): void {
  // // Apply persisted values on startup with safe option validation.
  if (!snapshot || typeof snapshot !== "object") {
    return;
  }

  if (elements.sourceMode && snapshot.sourceMode && hasSelectOption(elements.sourceMode, snapshot.sourceMode)) {
    elements.sourceMode.value = snapshot.sourceMode;
  }

  if (elements.srtPath && typeof snapshot.srtPath === "string") {
    elements.srtPath.value = snapshot.srtPath;
  }

  if (elements.whisperModel && snapshot.whisperModel && hasSelectOption(elements.whisperModel, snapshot.whisperModel)) {
    elements.whisperModel.value = snapshot.whisperModel;
  }

  if (
    elements.whisperSequenceRange &&
    snapshot.whisperSequenceRange &&
    hasSelectOption(elements.whisperSequenceRange, snapshot.whisperSequenceRange)
  ) {
    elements.whisperSequenceRange.value = snapshot.whisperSequenceRange;
  }

  if (elements.animationMode && snapshot.animationMode && hasSelectOption(elements.animationMode, snapshot.animationMode)) {
    elements.animationMode.value = snapshot.animationMode;
  }

  if (elements.maxChars && Number.isFinite(Number(snapshot.maxCharsPerLine))) {
    elements.maxChars.value = String(snapshot.maxCharsPerLine);
  }

  if (elements.linesPerCaption && Number.isFinite(Number(snapshot.linesPerCaption))) {
    elements.linesPerCaption.value = String(snapshot.linesPerCaption);
  }

  if (elements.fontSize && Number.isFinite(Number(snapshot.fontSize))) {
    elements.fontSize.value = String(snapshot.fontSize);
  }

  if (elements.mogrtAspectFilter && snapshot.mogrtAspectFilter) {
    if (hasSelectOption(elements.mogrtAspectFilter, snapshot.mogrtAspectFilter)) {
      elements.mogrtAspectFilter.value = snapshot.mogrtAspectFilter;
    } else {
      pendingMogrtAspectFilter = snapshot.mogrtAspectFilter;
    }
  }

  if (typeof snapshot.visualLiveUpdate === "boolean") {
    setVisualLiveUpdateEnabled(snapshot.visualLiveUpdate, true);
  }

  if (typeof snapshot.logExpanded === "boolean") {
    setLogPanelExpanded(snapshot.logExpanded, true);
  }

  if (typeof snapshot.verboseLogs === "boolean") {
    setVerboseLogsEnabled(snapshot.verboseLogs, true);
  }

  if (typeof snapshot.selectedMogrtId === "string" && snapshot.selectedMogrtId.length > 0) {
    pendingSelectedMogrtId = snapshot.selectedMogrtId;
  }

  if (snapshot.activeMode === "visual") {
    activeMode = "visual";
  } else {
    activeMode = "generate";
  }
}

function refreshLiveUpdateButtonState(): void {
  // // Reflect live update toggle state through button color/text and ARIA state.
  if (!elements.visualLiveUpdateButton) {
    return;
  }

  elements.visualLiveUpdateButton.classList.toggle("is-active", visualLiveUpdateEnabled);
  elements.visualLiveUpdateButton.classList.toggle("button--secondary", !visualLiveUpdateEnabled);
  elements.visualLiveUpdateButton.setAttribute("aria-pressed", visualLiveUpdateEnabled ? "true" : "false");
  elements.visualLiveUpdateButton.textContent = translate(
    visualLiveUpdateEnabled ? "action.liveUpdateOn" : "action.liveUpdateOff"
  );
}

function setVisualLiveUpdateEnabled(enabled: boolean, skipPersist = false): void {
  // // Toggle live-update behavior and keep button state synchronized.
  visualLiveUpdateEnabled = enabled === true;
  refreshLiveUpdateButtonState();

  if (!skipPersist) {
    persistPanelState();
  }
}

function buildCompactLogValue(value: unknown, depth = 0, fieldName = ""): unknown {
  // // Reduce oversized structured log payloads so the panel stays readable during day-to-day use.
  if (value === null || typeof value === "undefined") {
    return value;
  }

  if (typeof value === "string") {
    return value.length > 1200 ? `${value.slice(0, 1200)}…` : value;
  }

  if (typeof value !== "object") {
    return value;
  }

  if (depth >= 4) {
    return Array.isArray(value) ? [`… ${value.length} item(s)`] : "[Object]";
  }

  if (Array.isArray(value)) {
    const limit = fieldName === "properties" ? 4 : fieldName === "debug" ? 8 : 6;
    const compactItems = value.slice(0, limit).map((entry) => buildCompactLogValue(entry, depth + 1, fieldName));
    if (value.length > limit) {
      compactItems.push(`… ${value.length - limit} more item(s)`);
    }
    return compactItems;
  }

  const compactObject: Record<string, unknown> = {};
  for (const [key, entryValue] of Object.entries(value as Record<string, unknown>)) {
    if (key === "debug" && Array.isArray(entryValue) && entryValue.length > 8) {
      compactObject[key] = buildCompactLogValue(entryValue, depth + 1, key);
      continue;
    }
    if (key === "properties" && Array.isArray(entryValue) && entryValue.length > 4) {
      compactObject[key] = buildCompactLogValue(entryValue, depth + 1, key);
      compactObject.propertiesCount = entryValue.length;
      continue;
    }
    compactObject[key] = buildCompactLogValue(entryValue, depth + 1, key);
  }
  return compactObject;
}

function renderCurrentLog(): void {
  // // Re-render the current log entry whenever locale, verbosity, or visibility changes.
  if (!elements.logOutput) {
    return;
  }

  if (!currentLogState) {
    elements.logOutput.textContent = "";
    elements.logOutput.classList.remove("log--error");
    return;
  }

  let outputText = currentLogState.plainText;
  if (currentLogState.structuredTitle) {
    const payloadToRender = verboseLogsEnabled
      ? currentLogState.structuredPayload
      : buildCompactLogValue(currentLogState.structuredPayload);
    outputText = `${currentLogState.structuredTitle}\n${JSON.stringify(payloadToRender, null, 2)}`;
  }

  elements.logOutput.textContent = outputText;
  elements.logOutput.classList.toggle("log--error", currentLogState.isError);
}

function refreshLogControlsState(): void {
  // // Keep log-panel buttons synchronized with expanded/verbosity state and current locale.
  if (elements.logPanel) {
    elements.logPanel.classList.toggle("is-collapsed", !logPanelExpanded);
  }
  if (elements.logOutput) {
    elements.logOutput.hidden = !logPanelExpanded;
  }
  if (elements.logToggleButton) {
    elements.logToggleButton.setAttribute("aria-pressed", logPanelExpanded ? "true" : "false");
    elements.logToggleButton.textContent = translate(logPanelExpanded ? "action.hideLogs" : "action.showLogs");
  }
  if (elements.logVerbosityButton) {
    elements.logVerbosityButton.setAttribute("aria-pressed", verboseLogsEnabled ? "true" : "false");
    elements.logVerbosityButton.textContent = translate(verboseLogsEnabled ? "action.fullLogs" : "action.compactLogs");
  }
}

function setLogPanelExpanded(expanded: boolean, skipPersist = false): void {
  // // Allow the user to collapse logs entirely when they do not need runtime traces on screen.
  logPanelExpanded = expanded === true;
  refreshLogControlsState();
  if (!skipPersist) {
    persistPanelState();
  }
}

function setVerboseLogsEnabled(enabled: boolean, skipPersist = false): void {
  // // Let the panel switch between compact logs and full debug payloads without losing the raw data.
  verboseLogsEnabled = enabled === true;
  refreshLogControlsState();
  renderCurrentLog();
  if (!skipPersist) {
    persistPanelState();
  }
}

function setStructuredLog(title: string, payload: unknown, isError = false): void {
  // // Store one structured log payload so verbosity changes can re-render without recomputing host calls.
  currentLogState = {
    plainText: "",
    structuredTitle: title,
    structuredPayload: payload,
    isError
  };
  renderCurrentLog();
}

function setStructuredLogFromRaw(title: string, rawPayload: string, isError = false): void {
  // // Parse JSON-ish host output when possible so the compact/full log toggle can work on the same entry.
  const payloadText = String(rawPayload || "").trim();
  try {
    setStructuredLog(title, JSON.parse(payloadText), isError);
  } catch {
    setLog(`${title}\n${payloadText}`, isError);
  }
}

function panelAssetPath(relativeOrAbsolute: string): string {
  // // Normalize extension-local asset paths for fetch/image usage.
  if (!relativeOrAbsolute) {
    return "";
  }

  if (/^(https?:|file:)?\/\//i.test(relativeOrAbsolute) || relativeOrAbsolute.startsWith("./")) {
    return relativeOrAbsolute;
  }

  const normalizedPath = String(relativeOrAbsolute || "").replace(/\\/g, "/");
  if (/^[a-zA-Z]:\//.test(normalizedPath)) {
    return `file:///${encodeURI(normalizedPath)}`;
  }
  if (normalizedPath.startsWith("/")) {
    return `file://${encodeURI(normalizedPath)}`;
  }

  return `./${relativeOrAbsolute.replace(/^\/+/, "")}`;
}

function normalizeVersion(input: string): string {
  // // Normalize semantic version strings like v1.2.3 into 1.2.3.
  const match = String(input || "").trim().match(/^v?(\d+)\.(\d+)\.(\d+)$/i);
  if (!match) {
    return "";
  }

  return `${match[1]}.${match[2]}.${match[3]}`;
}

function compareVersions(left: string, right: string): number {
  // // Compare normalized semver values and return 1, 0, or -1.
  const leftParts = normalizeVersion(left).split(".");
  const rightParts = normalizeVersion(right).split(".");

  if (leftParts.length !== 3 || rightParts.length !== 3) {
    return 0;
  }

  for (let index = 0; index < 3; index += 1) {
    const l = Number(leftParts[index]);
    const r = Number(rightParts[index]);
    if (l > r) {
      return 1;
    }
    if (l < r) {
      return -1;
    }
  }

  return 0;
}

function resolveReleaseZipUrl(release: { assets?: Array<{ name?: string; browser_download_url?: string }> }): string {
  // // Prefer zip release assets for one-click update downloads.
  const assets = Array.isArray(release.assets) ? release.assets : [];

  for (const asset of assets) {
    const name = String(asset.name || "").toLowerCase();
    const downloadUrl = String(asset.browser_download_url || "");
    if (name.endsWith(".zip") && downloadUrl) {
      return downloadUrl;
    }
  }

  return "";
}

function refreshVersionLabel(): void {
  // // Display the current extension version next to the panel title.
  if (!elements.appVersion) {
    return;
  }

  const normalized = normalizeVersion(panelMeta.version);
  const displayVersion = normalized || String(panelMeta.version || "0.0.0");
  elements.appVersion.textContent = `v${displayVersion}`;
}

function refreshUpdateBanner(): void {
  // // Render release-update banner whenever locale/version state changes.
  if (!elements.updateBanner || !elements.updateLink) {
    return;
  }

  if (!updateState.visible || !updateState.latestVersion || !updateState.downloadUrl) {
    elements.updateBanner.hidden = true;
    elements.updateLink.href = "#";
    return;
  }

  const currentVersion = normalizeVersion(panelMeta.version) || panelMeta.version;
  elements.updateBanner.hidden = false;
  elements.updateLink.href = updateState.downloadUrl;
  elements.updateLink.textContent = translateTemplate("update.downloadNotice", {
    latest: updateState.latestVersion,
    current: currentVersion
  });
}

async function openUpdateDownload(): Promise<void> {
  // // Route update-banner clicks through CEP browser opening so release downloads work inside Premiere panels.
  if (!updateState.visible || !updateState.downloadUrl) {
    return;
  }

  await openExternalUrl(updateState.downloadUrl);
}

function setLog(message: string, isError = false): void {
  // // Provide a single visible place for runtime status and error traces.
  currentLogState = {
    plainText: message,
    structuredTitle: "",
    structuredPayload: null,
    isError
  };
  renderCurrentLog();
}

async function loadPanelMeta(): Promise<void> {
  // // Load version/update metadata emitted by build step.
  try {
    const response = await fetch("./assets/subcreator-meta.json", { cache: "no-store" });
    if (!response.ok) {
      panelMeta = { ...FALLBACK_PANEL_META };
      return;
    }

    const parsed = (await response.json()) as Partial<PanelMeta>;
    panelMeta = {
      version: String(parsed.version || FALLBACK_PANEL_META.version),
      repository: String(parsed.repository || FALLBACK_PANEL_META.repository),
      releaseApiUrl: String(parsed.releaseApiUrl || FALLBACK_PANEL_META.releaseApiUrl),
      releasePageUrl: String(parsed.releasePageUrl || FALLBACK_PANEL_META.releasePageUrl)
    };
  } catch {
    panelMeta = { ...FALLBACK_PANEL_META };
  }
}

async function checkForUpdates(): Promise<void> {
  // // Query latest GitHub release and show a banner if a newer version exists.
  if (!window.fetch) {
    updateState.visible = false;
    refreshUpdateBanner();
    return;
  }

  updateState.visible = false;
  updateState.latestVersion = "";
  updateState.downloadUrl = "";
  refreshUpdateBanner();

  const currentVersion = normalizeVersion(panelMeta.version);
  if (!currentVersion) {
    return;
  }

  try {
    const response = await fetch(panelMeta.releaseApiUrl, { cache: "no-store" });
    if (!response.ok) {
      return;
    }

    const release = (await response.json()) as {
      tag_name?: string;
      html_url?: string;
      assets?: Array<{ name?: string; browser_download_url?: string }>;
    };

    const latestVersion = normalizeVersion(String(release.tag_name || ""));
    if (!latestVersion || compareVersions(latestVersion, currentVersion) <= 0) {
      return;
    }

    const downloadUrl = resolveReleaseZipUrl(release) || String(release.html_url || panelMeta.releasePageUrl || "");
    if (!downloadUrl) {
      return;
    }

    updateState.visible = true;
    updateState.latestVersion = latestVersion;
    updateState.downloadUrl = downloadUrl;
    refreshUpdateBanner();
  } catch {
    updateState.visible = false;
    refreshUpdateBanner();
  }
}

async function loadLocale(languageCode: string): Promise<void> {
  // // Fetch locale dictionaries from extension-local JSON files.
  const response = await fetch(`./locales/${languageCode}.json`);
  if (!response.ok) {
    throw new Error(`Cannot load locale '${languageCode}'`);
  }

  currentLocale = (await response.json()) as LocaleMap;

  document.querySelectorAll<HTMLElement>("[data-i18n]").forEach((node) => {
    const key = node.dataset.i18n;
    if (!key) {
      return;
    }

    node.textContent = translate(key);
  });

  document.querySelectorAll<HTMLElement>("[data-i18n-placeholder]").forEach((node) => {
    const key = node.dataset.i18nPlaceholder;
    if (!key) {
      return;
    }

    if (node instanceof HTMLInputElement || node instanceof HTMLTextAreaElement) {
      node.placeholder = translate(key);
    }
  });

  if (elements.visualApplyProgress && !elements.visualApplyProgress.hidden && elements.visualApplyProgressBar) {
    setVisualApplyProgressState(true, Number(elements.visualApplyProgressBar.value || 0), Number(elements.visualApplyProgressBar.max || 0));
  }
  refreshLiveUpdateButtonState();
  refreshLogControlsState();
  renderCurrentLog();
  refreshMogrtAspectFilterOptions();

  refreshUpdateBanner();
}

function getSourceMode(): SourceMode {
  // // Normalize source mode value from UI select control.
  return (elements.sourceMode?.value as SourceMode) || "srt";
}

function resolveExtensionRootPath(): string {
  // // Resolve extension root from current file:// URL in CEP panel context.
  try {
    const currentUrl = new URL(window.location.href);
    if (currentUrl.protocol !== "file:") {
      return "";
    }

    let pathname = decodeURIComponent(currentUrl.pathname);
    if (/^\/[a-zA-Z]:\//.test(pathname)) {
      pathname = pathname.slice(1);
    }

    return pathname.replace(/\/index\.html.*$/i, "");
  } catch {
    return "";
  }
}

function buildAbsoluteMogrtPath(extensionRootPath: string, templateRelativePath: string): string {
  // // Compose absolute bundled MOGRT path so host can import without guessing.
  if (!extensionRootPath || !templateRelativePath) {
    return "";
  }

  const rootNormalized = extensionRootPath.replace(/\/$/, "");
  const relNormalized = templateRelativePath.replace(/^\/+/, "");
  return `${rootNormalized}/templates/mogrt/${relNormalized}`;
}

function toggleSourceFields(): void {
  // // Show only the source-related controls needed for current workflow.
  const mode = getSourceMode();
  const whisperModeActive = mode === "whisper_sequence";

  if (elements.srtInputField) {
    elements.srtInputField.style.display = mode === "srt" ? "grid" : "none";
  }

  if (elements.whisperField) {
    elements.whisperField.style.display = whisperModeActive ? "grid" : "none";
  }

  if (elements.whisperSequenceHint) {
    elements.whisperSequenceHint.style.display = mode === "whisper_sequence" ? "block" : "none";
  }

  if (elements.whisperSequenceRange) {
    elements.whisperSequenceRange.style.display = mode === "whisper_sequence" ? "block" : "none";
    elements.whisperSequenceRange.disabled = mode !== "whisper_sequence";
  }
  if (elements.whisperModelRow) {
    elements.whisperModelRow.classList.toggle("is-single", mode !== "whisper_sequence");
  }
}

async function enforceWhisperSourceAvailability(): Promise<void> {
  // // Hide Whisper source option when local runtime is unavailable on this machine.
  if (!elements.sourceMode) {
    return;
  }

  const whisperOptions = Array.from(
    elements.sourceMode.querySelectorAll<HTMLOptionElement>('option[value="whisper_sequence"]')
  );
  if (!whisperOptions.length) {
    return;
  }

  try {
    const status = await getWhisperRuntimeStatus();
    if (status.available) {
      return;
    }

    for (const whisperOption of whisperOptions) {
      whisperOption.remove();
    }
    if (elements.sourceMode.value === "whisper_sequence") {
      elements.sourceMode.value = "srt";
    }
  } catch {
    // // Keep Whisper visible when detection fails unexpectedly to avoid hiding a usable source.
  }
}

async function loadSystemFontCatalogFallback(): Promise<void> {
  // // Load OS font families/styles so font selectors can offer non-MOGRT fonts when possible.
  try {
    const catalog = await readSystemFontCatalog();
    if (!catalog || !catalog.available) {
      systemFontCatalog = {
        available: false,
        source: catalog?.source || "unavailable",
        details: catalog?.details || "",
        families: [],
        stylesByFamily: {},
        fontTokensByFamilyStyle: {}
      };
      return;
    }

    systemFontCatalog = {
      available: true,
      source: String(catalog.source || "system-fonts"),
      details: String(catalog.details || ""),
      families: Array.isArray(catalog.families) ? catalog.families.slice() : [],
      stylesByFamily:
        catalog.stylesByFamily && typeof catalog.stylesByFamily === "object"
          ? Object.entries(catalog.stylesByFamily).reduce<Record<string, string[]>>((accumulator, [family, styles]) => {
              accumulator[String(family)] = Array.isArray(styles) ? styles.slice() : [];
              return accumulator;
            }, {})
          : {},
      fontTokensByFamilyStyle:
        catalog.fontTokensByFamilyStyle && typeof catalog.fontTokensByFamilyStyle === "object"
          ? Object.entries(catalog.fontTokensByFamilyStyle).reduce<Record<string, Record<string, string>>>(
              (accumulator, [family, tokenMap]) => {
                accumulator[String(family)] = tokenMap && typeof tokenMap === "object"
                  ? Object.entries(tokenMap).reduce<Record<string, string>>((styleAccumulator, [style, token]) => {
                      styleAccumulator[String(style)] = String(token || "");
                      return styleAccumulator;
                    }, {})
                  : {};
                return accumulator;
              },
              {}
            )
          : {}
    };
  } catch {
    systemFontCatalog = {
      available: false,
      source: "error",
      details: "",
      families: [],
      stylesByFamily: {},
      fontTokensByFamilyStyle: {}
    };
  }
}

function ensureSystemFontCatalogLoaded(): Promise<void> {
  // // Defer the expensive system-font scan until the visual editor actually needs font-family/style expansion.
  if (systemFontCatalog.available || systemFontCatalog.source !== "unavailable" || systemFontCatalog.details) {
    return Promise.resolve();
  }

  if (systemFontCatalogLoadPromise) {
    return systemFontCatalogLoadPromise;
  }

  systemFontCatalogLoadPromise = loadSystemFontCatalogFallback().finally(() => {
    systemFontCatalogLoadPromise = null;
  });
  return systemFontCatalogLoadPromise;
}

function setActiveMode(mode: PanelMode): void {
  // // Toggle tab state and active mode container visibility.
  activeMode = mode;
  if (mode !== "visual") {
    visualLiveUpdateQueued = false;
    if (visualLiveUpdateTimer !== null) {
      window.clearTimeout(visualLiveUpdateTimer);
      visualLiveUpdateTimer = null;
    }
  }

  if (elements.tabGenerate) {
    const isActive = mode === "generate";
    elements.tabGenerate.classList.toggle("is-active", isActive);
    elements.tabGenerate.setAttribute("aria-selected", isActive ? "true" : "false");
  }

  if (elements.tabVisual) {
    const isActive = mode === "visual";
    elements.tabVisual.classList.toggle("is-active", isActive);
    elements.tabVisual.setAttribute("aria-selected", isActive ? "true" : "false");
  }

  if (elements.modeGenerate) {
    elements.modeGenerate.hidden = mode !== "generate";
  }

  if (elements.modeVisual) {
    elements.modeVisual.hidden = mode !== "visual";
  }
}

function formatVisualValue(valueType: HostVisualProperty["valueType"], value: string | number | boolean): string {
  // // Normalize host values for text/textarea fields in the visual editor.
  if (valueType === "json") {
    if (typeof value === "string") {
      return value;
    }

    try {
      return JSON.stringify(value, null, 2);
    } catch {
      return String(value);
    }
  }

  return String(value);
}

function normalizeColorHex(value: string | number | boolean): string {
  // // Normalize possible color values into lowercase #rrggbb.
  const text = String(value || "").trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(text)) {
    return text;
  }
  return "";
}

function looksLikeGuidList(value: string): boolean {
  // // Detect Premiere internal GUID-list artifacts to avoid exposing them as group labels.
  return /^(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12};)+$/i.test(String(value || "").trim());
}

function parseVectorValues(value: string | number | boolean): number[] {
  // // Parse vector payloads from host JSON strings into editable number fields.
  if (typeof value === "number") {
    return [value];
  }

  if (typeof value === "boolean") {
    return [value ? 1 : 0];
  }

  const text = String(value || "").trim();
  if (!text) {
    return [0, 0];
  }

  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      const numbers = parsed
        .map((item) => Number(item))
        .filter((item) => Number.isFinite(item))
        .slice(0, 4);
      if (numbers.length > 0) {
        return numbers;
      }
    }
  } catch {
    // // Fall through to compact CSV parsing when payload is not strict JSON.
  }

  const csvNumbers = text
    .replace(/^\[|\]$/g, "")
    .split(",")
    .map((item) => Number(item.trim()))
    .filter((item) => Number.isFinite(item))
    .slice(0, 4);
  return csvNumbers.length > 0 ? csvNumbers : [0, 0];
}

function vectorAxisLabel(index: number): string {
  // // Label vector components to mirror X/Y presentation from Premiere Properties.
  if (index === 0) {
    return "X";
  }
  if (index === 1) {
    return "Y";
  }
  if (index === 2) {
    return "Z";
  }
  return "W";
}

function canonicalizeVisualValue(
  controlKind: HostVisualProperty["controlKind"],
  valueType: HostVisualProperty["valueType"],
  value: string | number | boolean
): string {
  // // Build stable comparable value strings so apply only sends modified controls.
  if (controlKind === "vector") {
    const vectorValues = parseVectorValues(value).map((item) => Number(item.toFixed(6)));
    return JSON.stringify(vectorValues);
  }

  if (controlKind === "color") {
    return normalizeColorHex(value) || String(value || "").trim().toLowerCase();
  }

  if (valueType === "boolean") {
    return value === true ? "true" : "false";
  }

  if (valueType === "number") {
    const numericValue = Number(value);
    return Number.isFinite(numericValue) ? String(Number(numericValue.toFixed(6))) : "0";
  }

  return String(value ?? "");
}

function captureOpenVisualGroupsFromDom(): void {
  // // Persist current expand/collapse state before re-rendering the visual editor list.
  if (!elements.visualPropertyList) {
    return;
  }

  const groups = elements.visualPropertyList.querySelectorAll<HTMLDetailsElement>("details.visual-group[data-group-name]");
  groups.forEach((groupNode) => {
    const groupName = String(groupNode.dataset.groupName || "").trim();
    if (!groupName) {
      return;
    }

    if (groupNode.open) {
      visualOpenGroups.add(groupName);
    } else {
      visualOpenGroups.delete(groupName);
    }
  });
}

function buildVisualPropertySignature(properties: HostVisualProperty[]): string {
  // // Track which host-property set is currently rendered so deferred refreshes never overwrite a newer selection.
  return properties
    .map((property) => `${property.path}|${property.controlKind}|${property.valueType}`)
    .sort()
    .join("\n");
}

function hasPendingVisualEditorEdits(): boolean {
  // // Avoid background rerenders once the user has started editing rendered visual controls.
  if (!elements.visualPropertyList) {
    return false;
  }

  const controls = elements.visualPropertyList.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(
    '[data-visual-role="value"]'
  );
  for (const control of controls) {
    const path = String(control.dataset.visualPath || "").trim();
    if (!path || !visualOriginalValuesByPath.has(path)) {
      continue;
    }

    const valueType = String(control.dataset.visualType || "string") as HostVisualProperty["valueType"];
    const controlKind = String(control.dataset.visualControlKind || "string") as HostVisualProperty["controlKind"];
    const currentValue =
      control instanceof HTMLInputElement && control.type === "checkbox"
        ? control.checked
        : control.value;
    const currentCanonical = canonicalizeVisualValue(controlKind, valueType, currentValue);
    const originalCanonical = String(visualOriginalValuesByPath.get(path) || "");
    if (currentCanonical !== originalCanonical) {
      return true;
    }
  }

  return false;
}

function updateVisualSelectionSummary(message: string): void {
  // // Keep selection summary centralized for clearer visual-editor feedback.
  if (!elements.visualSelectionSummary) {
    return;
  }

  elements.visualSelectionSummary.textContent = message;
}

function setVisualApplyButtonsBusy(isBusy: boolean): void {
  // // Prevent concurrent apply/read actions while host changes are being processed.
  if (elements.visualApplyButton) {
    elements.visualApplyButton.disabled = isBusy;
  }
  if (elements.visualReadButton) {
    elements.visualReadButton.disabled = isBusy;
  }
}

function setGenerateButtonsBusy(isBusy: boolean): void {
  // // Prevent duplicate generate runs while export/transcription/apply is already active.
  if (elements.generateButton) {
    elements.generateButton.disabled = isBusy;
  }
  if (elements.languageSelect) {
    elements.languageSelect.disabled = isBusy;
  }
  if (elements.srtBrowseButton) {
    elements.srtBrowseButton.disabled = isBusy;
  }
  if (elements.srtPath) {
    elements.srtPath.disabled = isBusy;
  }
  if (elements.sourceMode) {
    elements.sourceMode.disabled = isBusy;
  }
  if (elements.whisperModel) {
    elements.whisperModel.disabled = isBusy;
  }
  if (elements.whisperSequenceRange) {
    elements.whisperSequenceRange.disabled = isBusy || getSourceMode() !== "whisper_sequence";
  }
  if (elements.animationMode) {
    elements.animationMode.disabled = isBusy;
  }
  if (elements.maxChars) {
    elements.maxChars.disabled = isBusy;
  }
  if (elements.linesPerCaption) {
    elements.linesPerCaption.disabled = isBusy;
  }
  if (elements.fontSize) {
    elements.fontSize.disabled = isBusy;
  }
  if (elements.mogrtAspectFilter) {
    elements.mogrtAspectFilter.disabled = isBusy;
  }
  if (elements.mogrtFolderButton) {
    elements.mogrtFolderButton.disabled = isBusy;
  }
  if (elements.mogrtRefreshButton) {
    elements.mogrtRefreshButton.disabled = isBusy;
  }
}

function setGenerateProgressState(visible: boolean, done = 0, total = GENERATE_PROGRESS_MAX, label = ""): void {
  // // Render generation/transcription progress using the same compact progress component style as visual apply.
  if (!elements.generateProgress || !elements.generateProgressBar || !elements.generateProgressText) {
    return;
  }

  if (!visible || total < 1) {
    elements.generateProgress.hidden = true;
    elements.generateProgressBar.max = GENERATE_PROGRESS_MAX;
    elements.generateProgressBar.value = 0;
    elements.generateProgressText.textContent = "0%";
    return;
  }

  const clampedTotal = Math.max(1, total);
  const clampedDone = Math.max(0, Math.min(clampedTotal, done));
  elements.generateProgress.hidden = false;
  elements.generateProgressBar.max = clampedTotal;
  elements.generateProgressBar.value = clampedDone;
  elements.generateProgressText.textContent = label || `${Math.round((clampedDone / clampedTotal) * 100)}%`;
}

function setVisualApplyProgressState(visible: boolean, done = 0, total = 0): void {
  // // Render visual apply progress feedback for multi-MOGRT updates.
  if (!elements.visualApplyProgress || !elements.visualApplyProgressBar || !elements.visualApplyProgressText) {
    return;
  }

  if (!visible || total < 1) {
    elements.visualApplyProgress.hidden = true;
    elements.visualApplyProgressBar.max = 1;
    elements.visualApplyProgressBar.value = 0;
    elements.visualApplyProgressText.textContent = "0 / 0";
    return;
  }

  const clampedDone = Math.max(0, Math.min(total, done));
  const remaining = Math.max(0, total - clampedDone);
  elements.visualApplyProgress.hidden = false;
  elements.visualApplyProgressBar.max = total;
  elements.visualApplyProgressBar.value = clampedDone;
  elements.visualApplyProgressText.textContent = translateTemplate("visual.applyProgress", {
    done: String(clampedDone),
    total: String(total),
    remaining: String(remaining)
  });
}

function renderVisualPropertyEditor(properties: HostVisualProperty[]): void {
  // // Render editable controls from selected MOGRT property metadata returned by host.
  if (!elements.visualPropertyList) {
    return;
  }

  captureOpenVisualGroupsFromDom();
  elements.visualPropertyList.innerHTML = "";
  visualOriginalValuesByPath.clear();
  visualTextStyleTokenMapByBasePath.clear();

  if (!properties.length) {
    return;
  }

  const textStyleFamilySelectByBasePath = new Map<string, HTMLSelectElement>();
  const textStyleStyleSelectByBasePath = new Map<string, HTMLSelectElement>();
  const textStyleStylesByFamilyByBasePath = new Map<string, Record<string, string[]>>();
  const textStyleFlagCheckboxesByBasePath = new Map<
    string,
    { bold?: HTMLInputElement; italic?: HTMLInputElement; allCaps?: HTMLInputElement; smallCaps?: HTMLInputElement }
  >();

  const normalizeStyleMap = (value: unknown): Record<string, string[]> => {
    // // Normalize host style-map payload into lowercase-keyed arrays for quick lookups.
    const normalized: Record<string, string[]> = {};
    if (!value || typeof value !== "object") {
      return normalized;
    }
    for (const [rawFamily, rawStyles] of Object.entries(value as Record<string, unknown>)) {
      const family = String(rawFamily || "").trim();
      if (!family || !Array.isArray(rawStyles)) {
        continue;
      }
      const uniqueStyles: string[] = [];
      for (const styleValue of rawStyles) {
        const styleText = String(styleValue || "").trim();
        if (!styleText) {
          continue;
        }
        if (!uniqueStyles.some((item) => item.toLowerCase() === styleText.toLowerCase())) {
          uniqueStyles.push(styleText);
        }
      }
      if (uniqueStyles.length > 0) {
        for (const familyLookupKey of listFontFamilyLookupKeys(family)) {
          normalized[familyLookupKey] = uniqueStyles.slice();
        }
      }
    }
    return normalized;
  };

  const mergeStyleMaps = (...maps: Array<Record<string, string[]>>): Record<string, string[]> => {
    // // Merge style maps with case-insensitive family/style dedupe.
    const merged: Record<string, string[]> = {};

    for (const map of maps) {
      if (!map || typeof map !== "object") {
        continue;
      }

      for (const [rawFamily, rawStyles] of Object.entries(map)) {
        const family = String(rawFamily || "").trim();
        if (!family || !Array.isArray(rawStyles)) {
          continue;
        }

        const familyKey = family.toLowerCase();
        const existingFamily =
          Object.keys(merged).find((entry) => entry.toLowerCase() === familyKey) || family;
        if (!Array.isArray(merged[existingFamily])) {
          merged[existingFamily] = [];
        }
        for (const styleValue of rawStyles) {
          const styleText = String(styleValue || "").trim();
          if (!styleText) {
            continue;
          }
          if (!merged[existingFamily].some((entry) => entry.toLowerCase() === styleText.toLowerCase())) {
            merged[existingFamily].push(styleText);
          }
        }
      }
    }

    return merged;
  };

  const normalizeTokenMap = (value: unknown): Record<string, string> => {
    // // Normalize family/style -> exact font token map for panel-side apply payloads.
    const normalized: Record<string, string> = {};
    if (!value || typeof value !== "object") {
      return normalized;
    }
    for (const [rawFamily, rawStyleMap] of Object.entries(value as Record<string, unknown>)) {
      const family = String(rawFamily || "").trim();
      if (!family || !rawStyleMap || typeof rawStyleMap !== "object") {
        continue;
      }
      for (const [rawStyle, rawToken] of Object.entries(rawStyleMap as Record<string, unknown>)) {
        const style = String(rawStyle || "").trim();
        const token = String(rawToken || "").trim();
        if (!style || !token) {
          continue;
        }
        for (const familyLookupKey of listFontFamilyLookupKeys(family)) {
          for (const styleLookupKey of listFontStyleLookupKeys(style)) {
            normalized[`${familyLookupKey}::${styleLookupKey}`] = token;
          }
        }
      }
    }
    return normalized;
  };

  const mergeTokenMaps = (...maps: Array<Record<string, string>>): Record<string, string> => {
    // // Merge exact font-token lookup tables without overwriting earlier matches.
    const merged: Record<string, string> = {};
    for (const map of maps) {
      if (!map || typeof map !== "object") {
        continue;
      }
      for (const [lookupKey, token] of Object.entries(map)) {
        const normalizedLookupKey = String(lookupKey || "").trim();
        const normalizedToken = String(token || "").trim();
        if (!normalizedLookupKey || !normalizedToken || merged[normalizedLookupKey]) {
          continue;
        }
        merged[normalizedLookupKey] = normalizedToken;
      }
    }
    return merged;
  };

  const normalizedSystemStyleMap = systemFontCatalog.available
    ? normalizeStyleMap(systemFontCatalog.stylesByFamily)
    : {};
  const normalizedSystemTokenMap = systemFontCatalog.available
    ? normalizeTokenMap(systemFontCatalog.fontTokensByFamilyStyle)
    : {};
  const systemFamilies = systemFontCatalog.available && Array.isArray(systemFontCatalog.families)
    ? systemFontCatalog.families
    : [];
  const commonSyntheticFontStyles = ["Regular", "Medium", "Semibold", "Bold", "Italic", "Bold Italic", "Black", "ExtraBold"];

  const resolveStyleOptionsForFamily = (family: string, styleMap?: Record<string, string[]>): string[] => {
    // // Resolve styles for one family using host hints first, then the local system font catalog.
    const mergedMap = mergeStyleMaps(styleMap || {}, normalizedSystemStyleMap);
    for (const familyLookupKey of listFontFamilyLookupKeys(family)) {
      if (Array.isArray(mergedMap[familyLookupKey]) && mergedMap[familyLookupKey].length > 0) {
        return sortFontStyleOptions(mergedMap[familyLookupKey].slice());
      }
    }
    return [];
  };

  const ensureTextStyleControls = (sourceProperties: HostVisualProperty[]): HostVisualProperty[] => {
    // // Inject a synthetic `Font Style` select when host payload exposes only `Font Family`.
    const propertiesWithStyleControls: HostVisualProperty[] = [];
    const styleBasePaths = new Set<string>();

    for (const property of sourceProperties) {
      const textStylePath = parseTextStyleVirtualPath(property.path);
      if (textStylePath?.styleKey === "fontStyle") {
        styleBasePaths.add(textStylePath.basePath);
      }
    }

    for (const property of sourceProperties) {
      propertiesWithStyleControls.push(property);

      const textStylePath = parseTextStyleVirtualPath(property.path);
      if (!textStylePath || textStylePath.styleKey !== "fontFamily" || styleBasePaths.has(textStylePath.basePath)) {
        continue;
      }

      const currentFamily = String(property.value || "").trim();
      const resolvedStyleOptions = resolveStyleOptionsForFamily(currentFamily, property.styleOptionsByFamily);
      const syntheticStyleOptions = resolvedStyleOptions.length > 0 ? resolvedStyleOptions : commonSyntheticFontStyles.slice();
      const syntheticStyleValue = pickPreferredFontStyleOption(syntheticStyleOptions) || "Regular";

      propertiesWithStyleControls.push({
        path: `${textStylePath.basePath}::textstyle.fontStyle`,
        displayName: "Font Style",
        groupPath: property.groupPath,
        valueType: "string",
        controlKind: "select",
        options: syntheticStyleOptions.map((styleOption) => ({
          value: styleOption,
          label: styleOption
        })),
        styleOptionsByFamily: currentFamily ? { [currentFamily]: syntheticStyleOptions.slice() } : {},
        value: syntheticStyleValue
      });
    }

    return propertiesWithStyleControls;
  };

  const renderProperties = ensureTextStyleControls(properties);
  loadedVisualProperties = renderProperties.slice();

  const replaceSelectOptions = (
    select: HTMLSelectElement,
    options: string[],
    preferredValue: string,
    dedupeKeyBuilder?: (value: string) => string
  ): void => {
    // // Replace select items while preserving currently selected value when possible.
    const deduped: string[] = [];
    for (const optionText of options) {
      const normalized = String(optionText || "").trim();
      if (!normalized) {
        continue;
      }
      const normalizedKey = dedupeKeyBuilder ? dedupeKeyBuilder(normalized) : normalized.toLowerCase();
      if (!deduped.some((item) => (dedupeKeyBuilder ? dedupeKeyBuilder(item) : item.toLowerCase()) === normalizedKey)) {
        deduped.push(normalized);
      }
    }

    const currentValue = String(preferredValue || select.value || "").trim();
    const previousValue = select.value;
    select.innerHTML = "";
    for (const optionText of deduped) {
      const option = document.createElement("option");
      option.value = optionText;
      option.textContent = optionText;
      select.appendChild(option);
    }

    const targetValue = currentValue || previousValue;
    if (!targetValue) {
      return;
    }
    const targetValueKey = dedupeKeyBuilder ? dedupeKeyBuilder(targetValue) : targetValue.toLowerCase();
    for (const option of Array.from(select.options)) {
      const optionValueKey = dedupeKeyBuilder ? dedupeKeyBuilder(String(option.value || "")) : String(option.value).toLowerCase();
      if (optionValueKey === targetValueKey) {
        select.value = option.value;
        return;
      }
    }
    if (!select.value && select.options.length > 0) {
      select.selectedIndex = 0;
    }
  };

  const refreshStyleSelectForFamily = (
    basePath: string,
    options?: {
      preserveCurrent?: boolean;
      preferredStyles?: string[];
    }
  ): void => {
    // // Keep font-style options aligned with selected family and optionally reset to neutral defaults on family changes.
    const familySelect = textStyleFamilySelectByBasePath.get(basePath);
    const styleSelect = textStyleStyleSelectByBasePath.get(basePath);
    const styleMap = textStyleStylesByFamilyByBasePath.get(basePath);
    if (!familySelect || !styleSelect || !styleMap) {
      return;
    }

    const selectedFamilyKeys = listFontFamilyLookupKeys(familySelect.value);
    if (selectedFamilyKeys.length < 1) {
      return;
    }

    let mappedOptions: string[] = [];
    for (const selectedFamilyKey of selectedFamilyKeys) {
      if (Object.prototype.hasOwnProperty.call(styleMap, selectedFamilyKey) && Array.isArray(styleMap[selectedFamilyKey])) {
        mappedOptions = styleMap[selectedFamilyKey];
        break;
      }
    }
    const preserveCurrent = options?.preserveCurrent !== false;
    const currentStyle = String(styleSelect.value || "").trim();
    if (mappedOptions.length > 0) {
      const orderedOptions = sortFontStyleOptions(mappedOptions);
      const nextStyle = preserveCurrent
        ? currentStyle
        : pickPreferredFontStyleOption(orderedOptions, options?.preferredStyles);
      replaceSelectOptions(styleSelect, orderedOptions, nextStyle, normalizeFontStyleDisplayKey);
      return;
    }
    if (preserveCurrent && currentStyle) {
      replaceSelectOptions(styleSelect, [currentStyle], currentStyle, normalizeFontStyleDisplayKey);
      return;
    }
    const fallbackStyle = pickPreferredFontStyleOption(["Regular"], options?.preferredStyles) || "Regular";
    replaceSelectOptions(styleSelect, ["Regular"], fallbackStyle, normalizeFontStyleDisplayKey);
  };

  const bindLiveUpdateEvent = (
    control: HTMLElement | HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement,
    eventName: "input" | "change" = "change"
  ): void => {
    // // Queue a debounced live apply whenever a visual control is edited.
    control.addEventListener(eventName, () => {
      scheduleLiveVisualApply();
    });
  };

  const selectRowHeightPx = 28;
  const selectMinRows = 4;
  const selectMaxRows = 10;

  const findScrollableAncestor = (node: HTMLElement): HTMLElement | null => {
    // // Find nearest scrollable container so we can reveal dropdowns near panel bottom.
    let current: HTMLElement | null = node.parentElement;
    while (current) {
      const style = window.getComputedStyle(current);
      const overflowY = String(style.overflowY || "").toLowerCase();
      if ((overflowY === "auto" || overflowY === "scroll") && current.scrollHeight > current.clientHeight) {
        return current;
      }
      current = current.parentElement;
    }
    return null;
  };

  const getClippingSpaceBelow = (select: HTMLSelectElement): number => {
    // // Compute visible space below select, constrained by viewport and clipping ancestors.
    const rect = select.getBoundingClientRect();
    const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
    const viewportBelow = Math.max(0, viewportHeight - rect.bottom - 8);
    let clippedBelow = viewportBelow;
    let current: HTMLElement | null = select.parentElement;
    while (current) {
      const style = window.getComputedStyle(current);
      const overflowY = String(style.overflowY || "").toLowerCase();
      if (overflowY === "hidden" || overflowY === "auto" || overflowY === "scroll") {
        const bounds = current.getBoundingClientRect();
        clippedBelow = Math.min(clippedBelow, Math.max(0, bounds.bottom - rect.bottom - 6));
      }
      current = current.parentElement;
    }
    return Math.max(0, clippedBelow);
  };

  const collapseExpandedSelect = (select: HTMLSelectElement): void => {
    // // Restore normal compact select after an inline expanded list interaction.
    select.size = 1;
    select.style.removeProperty("max-height");
    select.classList.remove("visual-select-expanded");
    select.removeAttribute("data-expanded-inline");
  };

  const expandSelectInline = (select: HTMLSelectElement, rows: number, maxHeightPx: number): void => {
    // // Show dropdown as inline list with bounded rows to avoid clipping outside panel.
    const boundedRows = Math.max(2, Math.min(rows, Math.max(2, select.options.length || 2)));
    if (select.getAttribute("data-expanded-inline") === "1" && Number(select.size || 1) === boundedRows) {
      return;
    }
    select.size = boundedRows;
    select.style.maxHeight = `${Math.max(84, Math.floor(maxHeightPx))}px`;
    select.classList.add("visual-select-expanded");
    select.setAttribute("data-expanded-inline", "1");
    const collapse = (): void => {
      collapseExpandedSelect(select);
      select.removeEventListener("blur", collapse);
      select.removeEventListener("change", collapse);
      select.removeEventListener("keydown", onKeydown);
    };
    const onKeydown = (event: KeyboardEvent): void => {
      if (event.key === "Escape" || event.key === "Enter") {
        collapse();
      }
    };
    select.addEventListener("blur", collapse);
    select.addEventListener("change", collapse);
    select.addEventListener("keydown", onKeydown);
  };

  const ensureSelectViewportSpace = (select: HTMLSelectElement): boolean => {
    // // Keep long font lists visible by constraining inline dropdown height to panel space.
    const desiredRows = Math.min(Math.max(selectMinRows, Math.min(select.options.length || selectMinRows, selectMaxRows)), selectMaxRows);
    const desiredHeight = desiredRows * selectRowHeightPx + 12;
    let spaceBelow = getClippingSpaceBelow(select);
    const missingSpace = desiredHeight - spaceBelow;
    if (missingSpace > 0) {
      const scrollable = findScrollableAncestor(select);
      if (scrollable) {
        const maxScrollable = Math.max(0, scrollable.scrollHeight - scrollable.clientHeight - scrollable.scrollTop);
        if (maxScrollable > 0) {
          const delta = Math.min(missingSpace, maxScrollable);
          if (delta > 0) {
            scrollable.scrollTop += delta;
          }
          spaceBelow = getClippingSpaceBelow(select);
        }
      }
    }

    const forceInline = (select.options.length || 0) > selectMaxRows;
    const shouldUseInline = forceInline || spaceBelow < desiredHeight;
    if (!shouldUseInline) {
      collapseExpandedSelect(select);
      return false;
    }

    const rowsFromSpace = Math.floor((Math.max(spaceBelow, 96) - 12) / selectRowHeightPx);
    const fallbackRows = Math.max(2, Math.min(select.options.length || selectMinRows, Math.min(selectMaxRows, Math.max(selectMinRows, rowsFromSpace))));
    const maxHeightPx = fallbackRows * selectRowHeightPx + 12;
    expandSelectInline(select, fallbackRows, Math.min(maxHeightPx, Math.max(spaceBelow - 4, 84)));
    return true;
  };

  const tryOpenConstrainedSelect = (select: HTMLSelectElement, event?: Event): void => {
    // // Expand only from collapsed state; keep option-click behavior intact while inline list is already open.
    if (select.getAttribute("data-expanded-inline") === "1" || Number(select.size || 1) > 1) {
      return;
    }
    if (ensureSelectViewportSpace(select)) {
      if (event) {
        event.preventDefault();
      }
      select.focus();
    }
  };

  const grouped = new Map<string, HostVisualProperty[]>();
  for (const property of renderProperties) {
    if (property.controlKind === "text" || property.controlKind === "json") {
      continue;
    }

    const rawGroup = String(property.groupPath || "").trim();
    const groupKey = rawGroup && !looksLikeGuidList(rawGroup) ? rawGroup : "General";
    visualOriginalValuesByPath.set(
      property.path,
      canonicalizeVisualValue(property.controlKind, property.valueType, property.value)
    );
    const existing = grouped.get(groupKey);
    if (existing) {
      existing.push(property);
    } else {
      grouped.set(groupKey, [property]);
    }
  }

  for (const [groupName, groupProperties] of grouped.entries()) {
    const groupNode = document.createElement("details");
    groupNode.className = "visual-group";
    groupNode.dataset.groupName = groupName;
    groupNode.open = visualOpenGroups.has(groupName);
    groupNode.addEventListener("toggle", () => {
      if (groupNode.open) {
        visualOpenGroups.add(groupName);
      } else {
        visualOpenGroups.delete(groupName);
      }
    });

    const summary = document.createElement("summary");
    summary.className = "visual-group__title";
    summary.textContent = groupName;
    groupNode.appendChild(summary);

    const groupBody = document.createElement("div");
    groupBody.className = "visual-group__body";

    for (const property of groupProperties) {
      const row = document.createElement("div");
      row.className = "visual-property-item";

      const label = document.createElement("label");
      label.className = "visual-property-label";
      label.textContent = property.displayName;

      if (property.controlKind === "checkbox") {
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = Boolean(property.value);
        checkbox.dataset.visualPath = property.path;
        checkbox.dataset.visualType = property.valueType;
        checkbox.dataset.visualControlKind = property.controlKind;
        checkbox.dataset.visualRole = "value";
        bindLiveUpdateEvent(checkbox, "change");

        const textStylePath = parseTextStyleVirtualPath(property.path);
        if (
          textStylePath &&
          (textStylePath.styleKey === "fontFsBold" ||
            textStylePath.styleKey === "fontFsItalic" ||
            textStylePath.styleKey === "fontFsAllCaps" ||
            textStylePath.styleKey === "fontFsSmallCaps")
        ) {
          // // Track text-style faux-style toggles so family/style changes can keep them coherent.
          const existing = textStyleFlagCheckboxesByBasePath.get(textStylePath.basePath) || {};
          if (textStylePath.styleKey === "fontFsBold") {
            existing.bold = checkbox;
          } else if (textStylePath.styleKey === "fontFsItalic") {
            existing.italic = checkbox;
          } else if (textStylePath.styleKey === "fontFsAllCaps") {
            existing.allCaps = checkbox;
          } else {
            existing.smallCaps = checkbox;
          }
          textStyleFlagCheckboxesByBasePath.set(textStylePath.basePath, existing);
          if (textStylePath.styleKey === "fontFsAllCaps" || textStylePath.styleKey === "fontFsSmallCaps") {
            checkbox.addEventListener("change", () => {
              if (!checkbox.checked) {
                return;
              }
              const pair = textStyleFlagCheckboxesByBasePath.get(textStylePath.basePath);
              if (!pair) {
                return;
              }
              if (textStylePath.styleKey === "fontFsAllCaps" && pair.smallCaps) {
                pair.smallCaps.checked = false;
              }
              if (textStylePath.styleKey === "fontFsSmallCaps" && pair.allCaps) {
                pair.allCaps.checked = false;
              }
            });
          }
        }

        row.classList.add("visual-property-item--checkbox");
        row.append(checkbox, label);
        groupBody.appendChild(row);
        continue;
      }

      row.appendChild(label);

      const controlWrap = document.createElement("div");
      controlWrap.className = "visual-property-control";

      if (property.controlKind === "text") {
        const textarea = document.createElement("textarea");
        textarea.rows = 2;
        textarea.value = formatVisualValue(property.valueType, property.value);
        textarea.dataset.visualPath = property.path;
        textarea.dataset.visualType = property.valueType;
        textarea.dataset.visualControlKind = property.controlKind;
        textarea.dataset.visualRole = "value";
        controlWrap.appendChild(textarea);
      } else if (property.controlKind === "json") {
        const textarea = document.createElement("textarea");
        textarea.rows = 3;
        textarea.value = formatVisualValue(property.valueType, property.value);
        textarea.dataset.visualPath = property.path;
        textarea.dataset.visualType = property.valueType;
        textarea.dataset.visualControlKind = property.controlKind;
        textarea.dataset.visualRole = "value";
        controlWrap.appendChild(textarea);
      } else if (property.controlKind === "color") {
        const colorWrap = document.createElement("div");
        colorWrap.className = "visual-color-row";

        const colorSwatch = document.createElement("button");
        colorSwatch.type = "button";
        colorSwatch.className = "visual-color-swatch";
        colorSwatch.setAttribute("aria-label", `${property.displayName} color`);

        const hexInput = document.createElement("input");
        hexInput.type = "text";
        hexInput.className = "visual-color-hex";
        hexInput.maxLength = 7;
        hexInput.spellcheck = false;
        hexInput.autocapitalize = "off";
        hexInput.autocomplete = "off";

        const nativeColorInput = document.createElement("input");
        nativeColorInput.type = "color";
        nativeColorInput.className = "visual-color-native";
        nativeColorInput.tabIndex = -1;
        nativeColorInput.setAttribute("aria-hidden", "true");

        const hiddenInput = document.createElement("input");
        hiddenInput.type = "hidden";
        hiddenInput.dataset.visualPath = property.path;
        hiddenInput.dataset.visualType = property.valueType;
        hiddenInput.dataset.visualControlKind = property.controlKind;
        hiddenInput.dataset.visualRole = "value";

        let syncing = false;

        const setColorState = (nextHex: string): void => {
          // // Keep swatch/native picker/hex controls synchronized from one canonical hex color.
          if (syncing) {
            return;
          }
          syncing = true;
          const normalized = normalizeColorHex(nextHex) || "#ffffff";
          hiddenInput.value = normalized;
          hexInput.value = normalized;
          colorSwatch.style.backgroundColor = normalized;
          nativeColorInput.value = normalized;
          syncing = false;
        };

        const initialHex = normalizeColorHex(property.value) || "#ffffff";
        setColorState(initialHex);

        hexInput.addEventListener("input", () => {
          const normalized = normalizeColorHex(hexInput.value);
          if (!normalized) {
            return;
          }
          setColorState(normalized);
          scheduleLiveVisualApply();
        });
        hexInput.addEventListener("blur", () => {
          setColorState(hiddenInput.value || initialHex);
        });

        colorSwatch.addEventListener("click", () => {
          // // Re-anchor hidden color input so native picker can open even near lower viewport edge.
          const swatchRect = colorSwatch.getBoundingClientRect();
          const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
          const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
          const estimatedPaletteHeight = 320;
          const estimatedPaletteWidth = 320;
          const margin = 8;
          const hasRoomBelow = viewportHeight - swatchRect.bottom >= estimatedPaletteHeight;
          const targetTop = hasRoomBelow
            ? Math.min(viewportHeight - margin, swatchRect.bottom + margin)
            : Math.max(margin, swatchRect.top - estimatedPaletteHeight - margin);
          const targetLeft = Math.min(Math.max(margin, swatchRect.left), Math.max(margin, viewportWidth - estimatedPaletteWidth));

          nativeColorInput.style.left = `${Math.round(targetLeft)}px`;
          nativeColorInput.style.top = `${Math.round(targetTop)}px`;

          // // Prefer showPicker when available; fallback to click for older CEP runtimes.
          const picker = nativeColorInput as HTMLInputElement & { showPicker?: () => void };
          if (typeof picker.showPicker === "function") {
            picker.showPicker();
            return;
          }
          nativeColorInput.click();
        });
        nativeColorInput.addEventListener("input", () => {
          setColorState(nativeColorInput.value || hiddenInput.value || initialHex);
          scheduleLiveVisualApply();
        });
        nativeColorInput.addEventListener("change", () => {
          setColorState(nativeColorInput.value || hiddenInput.value || initialHex);
          scheduleLiveVisualApply();
        });

        colorWrap.append(colorSwatch, hexInput, nativeColorInput);
        controlWrap.append(colorWrap, hiddenInput);
      } else if (property.controlKind === "slider") {
        const sliderWrap = document.createElement("div");
        sliderWrap.className = "visual-slider-row";

        const rangeInput = document.createElement("input");
        rangeInput.type = "range";
        const fallbackValue = Number(property.value || 0);
        const minValue = Number.isFinite(Number(property.minValue))
          ? Number(property.minValue)
          : Math.floor(fallbackValue - Math.max(Math.abs(fallbackValue), 50));
        const maxValue = Number.isFinite(Number(property.maxValue))
          ? Number(property.maxValue)
          : Math.ceil(fallbackValue + Math.max(Math.abs(fallbackValue), 50));
        const stepValue = Number.isFinite(Number(property.stepValue))
          ? Number(property.stepValue)
          : Number.isInteger(fallbackValue)
            ? 1
            : 0.1;
        rangeInput.min = String(minValue);
        rangeInput.max = String(maxValue);
        rangeInput.step = String(stepValue);
        rangeInput.value = String(fallbackValue);

        const numberInput = document.createElement("input");
        numberInput.type = "number";
        numberInput.step = String(stepValue);
        numberInput.min = String(minValue);
        numberInput.max = String(maxValue);
        numberInput.value = String(fallbackValue);
        numberInput.dataset.visualPath = property.path;
        numberInput.dataset.visualType = property.valueType;
        numberInput.dataset.visualControlKind = property.controlKind;
        numberInput.dataset.visualRole = "value";

        rangeInput.addEventListener("input", () => {
          numberInput.value = rangeInput.value;
          scheduleLiveVisualApply();
        });
        numberInput.addEventListener("input", () => {
          rangeInput.value = numberInput.value;
          scheduleLiveVisualApply();
        });

        sliderWrap.append(rangeInput, numberInput);
        controlWrap.appendChild(sliderWrap);
      } else if (property.controlKind === "vector") {
        const vectorWrap = document.createElement("div");
        vectorWrap.className = "visual-vector-row";

        const vectorValues = parseVectorValues(property.value);
        const hiddenInput = document.createElement("input");
        hiddenInput.type = "hidden";
        hiddenInput.dataset.visualPath = property.path;
        hiddenInput.dataset.visualType = property.valueType;
        hiddenInput.dataset.visualControlKind = property.controlKind;
        hiddenInput.dataset.visualRole = "value";
        if (Array.isArray(property.vectorScale) && property.vectorScale.length > 0) {
          hiddenInput.dataset.visualVectorScale = JSON.stringify(property.vectorScale);
        }
        hiddenInput.value = JSON.stringify(vectorValues);

        const componentInputs: HTMLInputElement[] = [];
        vectorValues.forEach((vectorValue, index) => {
          const vectorCell = document.createElement("label");
          vectorCell.className = "visual-vector-cell";

          const axis = document.createElement("span");
          axis.className = "visual-vector-axis";
          axis.textContent = vectorAxisLabel(index);

          const component = document.createElement("input");
          component.type = "number";
          component.step = Number.isInteger(vectorValue) ? "1" : "0.01";
          component.value = String(vectorValue);
          component.className = "visual-vector-input";
          componentInputs.push(component);
          vectorCell.append(axis, component);
          vectorWrap.appendChild(vectorCell);
        });

        const syncVector = (): void => {
          const nextValues = componentInputs.map((input) => Number(input.value)).filter((item) => Number.isFinite(item));
          hiddenInput.value = JSON.stringify(nextValues.length > 0 ? nextValues : [0, 0]);
        };
        componentInputs.forEach((input) => {
          input.addEventListener("input", syncVector);
          bindLiveUpdateEvent(input, "input");
        });
        syncVector();

        controlWrap.append(vectorWrap, hiddenInput);
      } else if (property.controlKind === "select" && Array.isArray(property.options) && property.options.length > 0) {
        const select = document.createElement("select");
        const currentValue = String(property.value ?? "");
        const currentValueNormalized = currentValue.toLowerCase();
        property.options.forEach((option) => {
          const node = document.createElement("option");
          node.value = String(option.value);
          node.textContent = option.label;
          if (String(option.value).toLowerCase() === currentValueNormalized) {
            node.selected = true;
          }
          select.appendChild(node);
        });
        select.dataset.visualPath = property.path;
        select.dataset.visualType = property.valueType;
        select.dataset.visualControlKind = property.controlKind;
        select.dataset.visualRole = "value";

        const textStylePath = parseTextStyleVirtualPath(property.path);
        if (textStylePath) {
          const hostMap = property.styleOptionsByFamily ? normalizeStyleMap(property.styleOptionsByFamily) : {};
          const existingMap = textStyleStylesByFamilyByBasePath.get(textStylePath.basePath) || {};
          const combinedMap = mergeStyleMaps(existingMap, hostMap, normalizedSystemStyleMap);
          if (Object.keys(combinedMap).length > 0) {
            textStyleStylesByFamilyByBasePath.set(textStylePath.basePath, combinedMap);
          }
          const existingTokenMap = visualTextStyleTokenMapByBasePath.get(textStylePath.basePath) || {};
          const combinedTokenMap = mergeTokenMaps(existingTokenMap, normalizedSystemTokenMap);
          if (Object.keys(combinedTokenMap).length > 0) {
            visualTextStyleTokenMapByBasePath.set(textStylePath.basePath, combinedTokenMap);
          }
        }
        if (textStylePath?.styleKey === "fontFamily") {
          const currentFamilyValue = String(select.value || currentValue || "").trim();
          const selectFamilies = Array.from(select.options).map((option) => String(option.value || ""));
          const dedupedFamilies: string[] = [];
          const seenFamilyKeys = new Set<string>();
          for (const familyName of [...systemFamilies, ...selectFamilies]) {
            const normalizedFamily = String(familyName || "").trim();
            if (!normalizedFamily) {
              continue;
            }
            const dedupeKey = normalizeCompactFontLookupKey(normalizedFamily) || normalizeFontLookupKey(normalizedFamily);
            if (!dedupeKey || seenFamilyKeys.has(dedupeKey)) {
              continue;
            }
            seenFamilyKeys.add(dedupeKey);
            dedupedFamilies.push(normalizedFamily);
          }
          replaceSelectOptions(select, dedupedFamilies, currentFamilyValue, normalizeCompactFontLookupKey);
          select.addEventListener("mousedown", (event) => {
            tryOpenConstrainedSelect(select, event);
          });
          select.addEventListener("touchstart", (event) => {
            tryOpenConstrainedSelect(select, event);
          });
          select.addEventListener("keydown", (event) => {
            if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
              tryOpenConstrainedSelect(select, event);
            }
          });
          textStyleFamilySelectByBasePath.set(textStylePath.basePath, select);
          select.addEventListener("change", () => {
            // // Reset family changes to a neutral style so old `Bold`/`Italic` values do not bleed into the new font.
            refreshStyleSelectForFamily(textStylePath.basePath, {
              preserveCurrent: false,
              preferredStyles: ["Regular", "Book", "Roman", "Plain", "Medium", "Semibold"]
            });
            scheduleLiveVisualApply();
          });
        } else if (textStylePath?.styleKey === "fontStyle") {
          textStyleStyleSelectByBasePath.set(textStylePath.basePath, select);
          const relatedMap = textStyleStylesByFamilyByBasePath.get(textStylePath.basePath);
          if (relatedMap && Object.keys(relatedMap).length > 0) {
            const selectStyles = Array.from(select.options).map((option) => String(option.value || ""));
            const relatedFamilySelect = textStyleFamilySelectByBasePath.get(textStylePath.basePath);
            const selectedFamilyKeys = listFontFamilyLookupKeys(String(relatedFamilySelect?.value || ""));
            let mappedStyles: string[] = [];
            for (const selectedFamilyKey of selectedFamilyKeys) {
              if (Array.isArray(relatedMap[selectedFamilyKey])) {
                mappedStyles = relatedMap[selectedFamilyKey];
                break;
              }
            }
            if (mappedStyles.length > 0) {
              replaceSelectOptions(
                select,
                sortFontStyleOptions(mappedStyles),
                String(select.value || currentValue || ""),
                normalizeFontStyleDisplayKey
              );
            } else if (selectedFamilyKeys.length > 0) {
              replaceSelectOptions(
                select,
                sortFontStyleOptions(selectStyles),
                String(select.value || currentValue || ""),
                normalizeFontStyleDisplayKey
              );
            } else {
              replaceSelectOptions(
                select,
                sortFontStyleOptions([...selectStyles, ...Object.values(relatedMap).flat()]),
                String(select.value || currentValue || ""),
                normalizeFontStyleDisplayKey
              );
            }
          }
          select.addEventListener("mousedown", (event) => {
            tryOpenConstrainedSelect(select, event);
          });
          select.addEventListener("touchstart", (event) => {
            tryOpenConstrainedSelect(select, event);
          });
          select.addEventListener("keydown", (event) => {
            if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
              tryOpenConstrainedSelect(select, event);
            }
          });
          bindLiveUpdateEvent(select, "change");
        } else {
          bindLiveUpdateEvent(select, "change");
        }

        controlWrap.appendChild(select);
      } else {
        const input = document.createElement("input");
        input.type = property.controlKind === "number" ? "number" : "text";
        if (property.controlKind === "number") {
          input.step = Number.isInteger(Number(property.value || 0)) ? "1" : "0.1";
        }
        input.value = formatVisualValue(property.valueType, property.value);
        input.dataset.visualPath = property.path;
        input.dataset.visualType = property.valueType;
        input.dataset.visualControlKind = property.controlKind;
        input.dataset.visualRole = "value";
        bindLiveUpdateEvent(input, "input");
        controlWrap.appendChild(input);
      }

      row.appendChild(controlWrap);
      groupBody.appendChild(row);
    }

    groupNode.appendChild(groupBody);
    elements.visualPropertyList.appendChild(groupNode);
  }

  for (const basePath of textStyleStyleSelectByBasePath.keys()) {
    // // Run one initial sync so style list follows currently selected family on first render.
    refreshStyleSelectForFamily(basePath, { preserveCurrent: true });
  }
}

type VisualPropertyChange = {
  path: string;
  valueType: HostVisualProperty["valueType"];
  controlKind: HostVisualProperty["controlKind"];
  vectorScale?: number[];
  fontToken?: string;
  value: string | number | boolean;
};

function findRenderedVisualControl(path: string): HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement | null {
  // // Resolve another rendered visual control by exact virtual path for text-style lookups.
  if (!elements.visualPropertyList) {
    return null;
  }

  const controls = elements.visualPropertyList.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(
    '[data-visual-role="value"]'
  );
  for (const control of controls) {
    if (String(control.dataset.visualPath || "") === path) {
      return control;
    }
  }
  return null;
}

function resolveVisualTextStyleToken(basePath: string): string {
  // // Resolve the exact system font token for the currently selected family/style pair when known.
  const tokenMap = visualTextStyleTokenMapByBasePath.get(basePath);
  if (!tokenMap) {
    return "";
  }

  const familyControl = findRenderedVisualControl(`${basePath}::textstyle.fontFamily`);
  const styleControl = findRenderedVisualControl(`${basePath}::textstyle.fontStyle`);
  const familyValue = String(familyControl instanceof HTMLSelectElement || familyControl instanceof HTMLInputElement ? familyControl.value : "")
    .trim();
  const styleValue = String(styleControl instanceof HTMLSelectElement || styleControl instanceof HTMLInputElement ? styleControl.value : "")
    .trim();
  if (!familyValue) {
    return "";
  }

  const styleLookupValues = styleValue ? listFontStyleLookupKeys(styleValue) : ["regular", "roman", "plain"];
  for (const familyLookupKey of listFontFamilyLookupKeys(familyValue)) {
    for (const styleLookupKey of styleLookupValues) {
      const token = tokenMap[`${familyLookupKey}::${styleLookupKey}`];
      if (token) {
        return token;
      }
    }
    const familyPrefix = `${familyLookupKey}::`;
    const anyFamilyEntry = Object.entries(tokenMap).find(([lookupKey]) => lookupKey.startsWith(familyPrefix));
    if (anyFamilyEntry?.[1]) {
      return anyFamilyEntry[1];
    }
  }
  return "";
}

function collectVisualPropertyChanges(): VisualPropertyChange[] {
  // // Build payload from rendered editor controls for host-side property updates.
  // // Always include current values so styles can be re-applied to any newly selected MOGRT clips.
  if (!elements.visualPropertyList) {
    return [];
  }

  const changes: VisualPropertyChange[] = [];
  const controls = elements.visualPropertyList.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(
    '[data-visual-role="value"]'
  );
  controls.forEach((control) => {
    const path = String(control.dataset.visualPath || "");
    const valueType = (String(control.dataset.visualType || "string") as HostVisualProperty["valueType"]) || "string";
    const controlKind =
      (String(control.dataset.visualControlKind || "string") as HostVisualProperty["controlKind"]) || "string";
    const vectorScaleRaw = String(control.dataset.visualVectorScale || "");
    const vectorScale = vectorScaleRaw
      ? (() => {
          try {
            const parsed = JSON.parse(vectorScaleRaw) as unknown;
            return Array.isArray(parsed) ? parsed.map((entry) => Number(entry)).filter((entry) => Number.isFinite(entry)) : undefined;
          } catch {
            return undefined;
          }
        })()
      : undefined;
    if (!path) {
      return;
    }

    let value: string | number | boolean = "";
    if (valueType === "boolean" && control instanceof HTMLInputElement) {
      value = control.checked;
    } else if (valueType === "number") {
      value = Number(control.value);
    } else {
      value = control.value;
    }

    const textStylePath = parseTextStyleVirtualPath(path);
    const fontToken =
      textStylePath && (textStylePath.styleKey === "fontFamily" || textStylePath.styleKey === "fontStyle")
        ? resolveVisualTextStyleToken(textStylePath.basePath)
        : "";

    changes.push({
      path,
      valueType,
      controlKind,
      vectorScale,
      fontToken: fontToken || undefined,
      value
    });
  });

  return changes;
}

function selectionHasFontControls(properties: HostVisualProperty[]): boolean {
  // // Load OS font metadata only for selections that actually expose font family/style controls.
  return properties.some((property) => {
    const textStylePath = parseTextStyleVirtualPath(property.path);
    if (!textStylePath) {
      return false;
    }
    return textStylePath.styleKey === "fontFamily" || textStylePath.styleKey === "fontStyle";
  });
}

async function loadVisualPropertiesFromSelection(emitHostLog = false): Promise<void> {
  // // Read selected MOGRT editable controls from host and refresh visual editor UI.
  const result = await readSelectedMogrtVisualProperties();
  const propertySignature = buildVisualPropertySignature(result.properties);
  loadedVisualPropertySignature = propertySignature;
  renderVisualPropertyEditor(result.properties);
  if (selectionHasFontControls(result.properties)) {
    void ensureSystemFontCatalogLoaded()
      .then(() => {
        if (activeMode !== "visual" || loadedVisualProperties.length < 1) {
          return;
        }
        if (loadedVisualPropertySignature !== propertySignature) {
          return;
        }
        if (hasPendingVisualEditorEdits()) {
          return;
        }
        renderVisualPropertyEditor(result.properties);
      })
      .catch(() => {
        // // Ignore deferred font-catalog failures so visual property reading never breaks on startup or slow systems.
      });
  }
  if (emitHostLog) {
    setStructuredLog(translate("log.hostResult"), result);
  }
  if (result.properties.length > 0) {
    updateVisualSelectionSummary(
      translateTemplate("visual.selectionSummary", {
        clips: String(result.selectedCount),
        props: String(result.editableCount)
      })
    );
  } else if (result.selectedCount > 0) {
    updateVisualSelectionSummary(translate("visual.noProperties"));
  } else {
    updateVisualSelectionSummary(translate("visual.selectionDefault"));
  }
}

function isVisualLiveUpdateEnabled(): boolean {
  // // Read live-update toggle value with safe fallback when UI is not ready.
  return visualLiveUpdateEnabled;
}

async function applyVisualChangesToSelection(options?: { liveUpdate?: boolean }): Promise<void> {
  // // Apply edited visual values; use progressive per-clip mode for multi-selection manual apply.
  const useLiveUpdate = options?.liveUpdate === true;
  const changes = collectVisualPropertyChanges();
  if (!changes.length) {
    if (useLiveUpdate) {
      return;
    }
    throw new Error(translate("visual.noChanges"));
  }

  if (visualApplyInProgress) {
    if (useLiveUpdate) {
      visualLiveUpdateQueued = true;
      return;
    }
    throw new Error("Visual apply already in progress.");
  }

  visualApplyInProgress = true;
  if (!useLiveUpdate) {
    setVisualApplyButtonsBusy(true);
  }

  try {
    const selectedCount = await getSelectedMogrtCount();
    if (!useLiveUpdate && selectedCount > 1) {
      let updatedCount = 0;
      let failedCount = 0;
      const debugLines: string[] = [];
      setVisualApplyProgressState(true, 0, selectedCount);

      for (let clipIndex = 0; clipIndex < selectedCount; clipIndex += 1) {
        const step = await applyVisualPropertiesToSelectedMogrts(changes, {
          clipStartIndex: clipIndex,
          clipEndIndex: clipIndex + 1
        });
        updatedCount += Number(step.updatedCount || 0);
        failedCount += Number(step.failedCount || 0);
        if (Array.isArray(step.debug)) {
          debugLines.push(...step.debug);
        }
        setVisualApplyProgressState(true, clipIndex + 1, selectedCount);
      }

      setStructuredLog(translate("log.visualApplyDone"), {
        selectedCount,
        processedClipCount: selectedCount,
        updatedCount,
        failedCount,
        debug: debugLines
      });
      setVisualApplyProgressState(false);
      await loadVisualPropertiesFromSelection();
      return;
    }

    const response = await applyVisualPropertiesToSelectedMogrts(changes);
    if (!useLiveUpdate) {
      setStructuredLog(translate("log.visualApplyDone"), response);
      await loadVisualPropertiesFromSelection();
    }
  } finally {
    visualApplyInProgress = false;
    if (!useLiveUpdate) {
      setVisualApplyButtonsBusy(false);
      setVisualApplyProgressState(false);
    }
  }
}

function scheduleLiveVisualApply(): void {
  // // Debounce live updates so rapid UI edits do not flood host apply calls.
  if (!isVisualLiveUpdateEnabled() || activeMode !== "visual") {
    return;
  }

  visualLiveUpdateQueued = true;
  if (visualLiveUpdateTimer !== null) {
    window.clearTimeout(visualLiveUpdateTimer);
  }

  visualLiveUpdateTimer = window.setTimeout(() => {
    void runQueuedLiveVisualApply();
  }, 220);
}

async function runQueuedLiveVisualApply(): Promise<void> {
  // // Execute one queued live apply pass and re-run if edits happened during host call.
  visualLiveUpdateTimer = null;
  if (!isVisualLiveUpdateEnabled() || activeMode !== "visual") {
    visualLiveUpdateQueued = false;
    return;
  }
  if (visualLiveUpdateInFlight || visualApplyInProgress) {
    if (visualLiveUpdateQueued) {
      if (visualLiveUpdateTimer !== null) {
        window.clearTimeout(visualLiveUpdateTimer);
      }
      visualLiveUpdateTimer = window.setTimeout(() => {
        void runQueuedLiveVisualApply();
      }, 220);
    }
    return;
  }
  if (!visualLiveUpdateQueued) {
    return;
  }

  visualLiveUpdateQueued = false;
  visualLiveUpdateInFlight = true;
  try {
    await applyVisualChangesToSelection({ liveUpdate: true });
  } catch (error) {
    setLog(String(error), true);
  } finally {
    visualLiveUpdateInFlight = false;
    if (visualLiveUpdateQueued) {
      scheduleLiveVisualApply();
    }
  }
}

function normalizeMogrtRelativePathKey(value: string): string {
  // // Build a stable relative-path key so static build catalog and installed scan can be merged safely.
  return String(value || "")
    .trim()
    .replace(/\\/g, "/")
    .replace(/^\/+/, "")
    .toLowerCase();
}

async function readBundledMogrtCatalog(): Promise<MogrtCatalog | null> {
  // // Read build-time catalog so bundled preview assets remain available for shipped templates.
  try {
    const response = await fetch("./assets/mogrt-catalog.json", { cache: "no-store" });
    if (!response.ok) {
      return null;
    }

    return (await response.json()) as MogrtCatalog;
  } catch {
    return null;
  }
}

function mergeInstalledMogrtTemplates(
  bundledTemplates: MogrtTemplateItem[],
  installedCatalog: InstalledMogrtCatalog | null
): MogrtTemplateItem[] {
  // // Merge build-time previews with runtime-installed templates so manual additions appear without rebuild.
  const bundledByPath = new Map<string, MogrtTemplateItem>();
  for (const template of bundledTemplates) {
    bundledByPath.set(normalizeMogrtRelativePathKey(template.relativePath), template);
  }

  if (!installedCatalog || !installedCatalog.available || installedCatalog.templates.length < 1) {
    return bundledTemplates.slice();
  }

  const mergedTemplates: MogrtTemplateItem[] = installedCatalog.templates.map((template) => {
    const bundledMatch = bundledByPath.get(normalizeMogrtRelativePathKey(template.relativePath));
    return {
      id: bundledMatch?.id || template.id,
      name: template.name,
      aspect: template.aspect,
      relativePath: template.relativePath,
      previewClass: template.previewClass || bundledMatch?.previewClass || "default",
      previewImagePath: template.previewImagePath || bundledMatch?.previewImagePath || "",
      previewVideoPath: template.previewVideoPath || bundledMatch?.previewVideoPath || ""
    };
  });

  const mergedPathKeys = new Set(mergedTemplates.map((template) => normalizeMogrtRelativePathKey(template.relativePath)));
  for (const bundledTemplate of bundledTemplates) {
    const bundledKey = normalizeMogrtRelativePathKey(bundledTemplate.relativePath);
    if (!mergedPathKeys.has(bundledKey)) {
      mergedTemplates.push({ ...bundledTemplate });
    }
  }

  mergedTemplates.sort((left, right) => {
    const groupCompare = String(left.aspect || "").localeCompare(String(right.aspect || ""), undefined, {
      sensitivity: "base"
    });
    if (groupCompare !== 0) {
      return groupCompare;
    }
    return String(left.name || "").localeCompare(String(right.name || ""), undefined, { sensitivity: "base" });
  });

  return mergedTemplates;
}

function refreshMogrtAspectFilterOptions(): void {
  // // Rebuild gallery filter choices from real installed top-level folders instead of hardcoded aspect presets.
  if (!elements.mogrtAspectFilter) {
    return;
  }

  const previousValue = pendingMogrtAspectFilter || elements.mogrtAspectFilter.value || "all";
  const folders = Array.from(
    new Set(
      availableMogrts
        .map((template) => String(template.aspect || "").trim())
        .filter((templateAspect) => templateAspect.length > 0)
    )
  ).sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));

  elements.mogrtAspectFilter.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = translate("gallery.allFormats");
  elements.mogrtAspectFilter.appendChild(allOption);

  for (const folder of folders) {
    const option = document.createElement("option");
    option.value = folder;
    option.textContent = folder;
    elements.mogrtAspectFilter.appendChild(option);
  }

  if (hasSelectOption(elements.mogrtAspectFilter, previousValue)) {
    elements.mogrtAspectFilter.value = previousValue;
    pendingMogrtAspectFilter = "";
  } else {
    elements.mogrtAspectFilter.value = "all";
    if (folders.length > 0 || previousValue === "all") {
      pendingMogrtAspectFilter = "";
    }
  }
}

async function loadMogrtCatalog(): Promise<void> {
  // // Load bundled catalog, then overlay runtime-installed templates so user-added MOGRTs show in the gallery.
  const bundledCatalog = await readBundledMogrtCatalog();
  const bundledTemplates = Array.isArray(bundledCatalog?.templates) ? bundledCatalog?.templates : [];
  const extensionRootPath = resolveExtensionRootPath();
  const installedCatalog = await readInstalledMogrtCatalog(extensionRootPath);

  availableMogrts = mergeInstalledMogrtTemplates(bundledTemplates, installedCatalog);
  refreshMogrtAspectFilterOptions();

  if (!availableMogrts.length && !bundledCatalog && !installedCatalog.available) {
    throw new Error(translate("error.mogrtCatalogMissing"));
  }

  if (!selectedMogrt && availableMogrts.length > 0) {
    selectedMogrt = availableMogrts[0];
  }
}

async function reloadMogrtCatalogPreservingSelection(): Promise<void> {
  // // Refresh gallery content after manual filesystem edits while keeping current filter/selection when possible.
  const previousSelectionId = selectedMogrt?.id || pendingSelectedMogrtId;
  const previousSelectionPath = selectedMogrt?.relativePath || "";
  pendingMogrtAspectFilter = elements.mogrtAspectFilter?.value || pendingMogrtAspectFilter;

  await loadMogrtCatalog();

  if (previousSelectionPath) {
    const restoredByPath = availableMogrts.find(
      (template) => normalizeMogrtRelativePathKey(template.relativePath) === normalizeMogrtRelativePathKey(previousSelectionPath)
    );
    if (restoredByPath) {
      selectedMogrt = restoredByPath;
    }
  }

  if ((!selectedMogrt || !availableMogrts.some((template) => template.id === selectedMogrt?.id)) && previousSelectionId) {
    const restoredById = availableMogrts.find((template) => template.id === previousSelectionId);
    if (restoredById) {
      selectedMogrt = restoredById;
    }
  }
}

function schedulePassiveMogrtCatalogRefresh(): void {
  // // Throttle passive filesystem refreshes so focus changes do not repeatedly rescan the gallery tree.
  if (passiveMogrtRefreshTimer !== null) {
    window.clearTimeout(passiveMogrtRefreshTimer);
  }

  passiveMogrtRefreshTimer = window.setTimeout(() => {
    passiveMogrtRefreshTimer = null;
    const now = Date.now();
    if (now - lastPassiveMogrtCatalogRefreshAt < 1500) {
      return;
    }
    lastPassiveMogrtCatalogRefreshAt = now;

    void reloadMogrtCatalogPreservingSelection()
      .then(() => {
        renderMogrtGallery();
        persistPanelState();
      })
      .catch(() => {
        // // Ignore passive refresh failures so gallery browsing never blocks the panel.
      });
  }, 250);
}

function updateSelectedMogrtLabel(): void {
  // // Keep current template selection visible to user before generation.
  if (!elements.mogrtSelectedLabel) {
    return;
  }

  if (!selectedMogrt) {
    elements.mogrtSelectedLabel.textContent = translate("gallery.noneSelected");
    return;
  }

  elements.mogrtSelectedLabel.textContent = `${translate("gallery.selectedPrefix")} ${selectedMogrt.name} (${selectedMogrt.aspect})`;
}

function updateMogrtGallerySelectionState(): void {
  // // Update only card active states so selecting a template does not rebuild the whole gallery DOM.
  if (!elements.mogrtGallery) {
    return;
  }

  const selectedTemplateId = selectedMogrt?.id || "";
  elements.mogrtGallery.querySelectorAll<HTMLElement>(".mogrt-card").forEach((card) => {
    const isActive = String(card.dataset.templateId || "") === selectedTemplateId;
    card.classList.toggle("is-active", isActive);
  });
  updateSelectedMogrtLabel();
}

function selectMogrt(templateId: string): void {
  // // Save selected template without recreating gallery cards on every click.
  const found = availableMogrts.find((template) => template.id === templateId);
  if (!found) {
    return;
  }

  selectedMogrt = found;
  updateMogrtGallerySelectionState();
  persistPanelState();
}

function renderMogrtGallery(): void {
  // // Render gallery cards with lightweight visual previews and aspect filtering.
  if (!elements.mogrtGallery || !elements.mogrtAspectFilter) {
    return;
  }

  const selectedAspect = elements.mogrtAspectFilter.value;
  const filtered = availableMogrts.filter((template) => {
    return selectedAspect === "all" || template.aspect === selectedAspect;
  });

  if (!selectedMogrt && filtered.length > 0) {
    selectedMogrt = filtered[0];
  }

  if (selectedMogrt && filtered.length > 0 && !filtered.some((item) => item.id === selectedMogrt?.id)) {
    selectedMogrt = filtered[0];
  }

  elements.mogrtGallery.innerHTML = "";

  if (filtered.length === 0) {
    elements.mogrtGallery.textContent = translate("gallery.empty");
    updateSelectedMogrtLabel();
    return;
  }

  for (const template of filtered) {
    const card = document.createElement("button");
    card.type = "button";
    card.className = "mogrt-card";
    card.dataset.templateId = template.id;
    if (selectedMogrt?.id === template.id) {
      card.classList.add("is-active");
    }

    const preview = document.createElement("div");
    preview.className = "mogrt-card__preview";
    preview.dataset.preview = template.previewClass;

    if (template.previewImagePath) {
      const previewImage = document.createElement("img");
      previewImage.className = "mogrt-card__preview-image";
      previewImage.src = panelAssetPath(template.previewImagePath);
      previewImage.loading = "lazy";
      previewImage.alt = `${template.name} preview`;
      preview.appendChild(previewImage);
    } else if (template.previewVideoPath) {
      const previewVideo = document.createElement("video");
      previewVideo.className = "mogrt-card__preview-video";
      previewVideo.src = panelAssetPath(template.previewVideoPath);
      previewVideo.muted = true;
      previewVideo.loop = true;
      previewVideo.autoplay = false;
      previewVideo.playsInline = true;
      previewVideo.preload = "none";
      previewVideo.setAttribute("aria-hidden", "true");
      const playPreview = () => {
        // // Play preview videos only on interaction to avoid burning CPU/GPU on large galleries.
        void previewVideo.play().catch(() => {
          // // Ignore autoplay/playback failures because previews are non-blocking UI sugar.
        });
      };
      const stopPreview = () => {
        // // Pause and reset preview videos when cards are not being inspected anymore.
        previewVideo.pause();
        try {
          previewVideo.currentTime = 0;
        } catch {
          // // Ignore seek reset errors from partially loaded previews.
        }
      };
      card.addEventListener("mouseenter", playPreview);
      card.addEventListener("focus", playPreview);
      card.addEventListener("mouseleave", stopPreview);
      card.addEventListener("blur", stopPreview);
      preview.appendChild(previewVideo);
    } else {
      preview.textContent = translate("gallery.previewText");
    }

    const name = document.createElement("div");
    name.className = "mogrt-card__name";
    name.textContent = template.name;

    const meta = document.createElement("div");
    meta.className = "mogrt-card__meta";
    meta.textContent = `${template.aspect}`;

    card.append(preview, name, meta);
    card.addEventListener("click", () => {
      selectMogrt(template.id);
    });

    elements.mogrtGallery.appendChild(card);
  }

  updateSelectedMogrtLabel();
}

async function browseSrtPath(): Promise<void> {
  // // Pick an SRT file path via host-native file chooser.
  if (!elements.srtPath) {
    return;
  }

  const selectedPath = await pickSrtPath();
  if (selectedPath) {
    elements.srtPath.value = selectedPath;
    persistPanelState();
  }
}

function collectBuildOptions(): CaptionBuildOptions {
  // // Collect, normalize, and validate all panel options into a single object.
  if (
    !elements.sourceMode ||
    !elements.languageSelect ||
    !elements.fontSize ||
    !elements.maxChars ||
    !elements.linesPerCaption ||
    !elements.animationMode ||
    !elements.whisperModel ||
    !elements.whisperSequenceRange
  ) {
    throw new Error("Panel bindings not initialized.");
  }

  if (!selectedMogrt && availableMogrts.length > 0) {
    selectedMogrt = availableMogrts[0];
  }

  const extensionRootPath = resolveExtensionRootPath();
  const templateRelativePath = selectedMogrt?.relativePath ?? "";

  return {
    sourceMode: getSourceMode(),
    languageCode: elements.languageSelect.value,
    style: {
      fontSize: Number(elements.fontSize.value),
      maxCharsPerLine: Number(elements.maxChars.value),
      animationMode: elements.animationMode.value as AnimationMode,
      uppercase: false,
      linesPerCaption: Number(elements.linesPerCaption.value)
    },
    extensionRootPath,
    mogrtPath: buildAbsoluteMogrtPath(extensionRootPath, templateRelativePath),
    mogrtTemplateRelativePath: templateRelativePath,
    whisperModel: elements.whisperModel.value,
    whisperSequenceRange: (elements.whisperSequenceRange.value as WhisperSequenceRangeMode) || "entire_sequence",
    videoTrackIndex: 0,
    audioTrackIndex: 0
  };
}

function mapWhisperPercentToGenerateProgress(progress: WhisperProgressUpdate): number {
  // // Reserve most of the generate bar for Whisper analysis while keeping room for planning/apply steps.
  const clampedPercent = Math.max(0, Math.min(100, Number(progress.percent || 0)));
  return 22 + Math.round((clampedPercent / 100) * 58);
}

function buildWhisperProgressLabel(progress: WhisperProgressUpdate): string {
  // // Keep transcription feedback explicit so a long Whisper run looks active instead of blocked.
  return translateTemplate("progress.whisperAnalysis", {
    percent: String(Math.max(0, Math.min(100, Math.round(Number(progress.percent || 0)))))
  });
}

async function updateGenerateProgress(done: number, label: string, waitForPaint = false): Promise<void> {
  // // Centralize generate progress updates and optionally yield a frame before the next expensive step.
  setGenerateProgressState(true, done, GENERATE_PROGRESS_MAX, label);
  if (waitForPaint) {
    await waitForNextPaint();
  }
}

async function loadCuesFromSelectedSource(
  options: CaptionBuildOptions,
  onProgress?: (done: number, label: string, waitForPaint?: boolean) => Promise<void>
): Promise<CaptionCue[]> {
  // // Build cues from the currently selected source mode.
  if (options.sourceMode === "srt") {
    if (!elements.srtPath || !elements.srtPath.value.trim()) {
      throw new Error(translate("error.missingSrtPath"));
    }

    if (onProgress) {
      await onProgress(10, translate("progress.readSrt"), true);
    }
    const srtText = await readTextFileFromHost(elements.srtPath.value.trim());
    if (onProgress) {
      await onProgress(28, translate("progress.parseSrt"));
    }
    const cues = parseSrt(srtText);
    if (!cues.length) {
      throw new Error(translate("error.emptySrt"));
    }

    return cues;
  }

  setLog(translate("log.whisperSequenceExport"));
  if (onProgress) {
    await onProgress(10, translate("progress.exportSequence"), true);
  }
  const exportResult = await exportActiveSequenceAudioForWhisper(options.whisperSequenceRange);
  const whisperAudioPath = exportResult.audioPath;
  const cleanupAudioPath = exportResult.audioPath;

  if (!whisperAudioPath) {
    throw new Error(translate("error.missingActiveSequenceAudio"));
  }

  try {
    if (onProgress) {
      await onProgress(22, translate("progress.whisperAnalyzing"), true);
    }
    const whisperResult = await transcribeWithWhisper({
      audioPath: whisperAudioPath,
      languageCode: options.languageCode,
      model: options.whisperModel
    }, (progress) => {
      void updateGenerateProgress(mapWhisperPercentToGenerateProgress(progress), buildWhisperProgressLabel(progress));
    });

    let cues: CaptionCue[] = [];
    if (onProgress) {
      await onProgress(84, translate("progress.parseWhisper"));
    }
    if (whisperResult.jsonText) {
      try {
        cues = parseWhisperJson(whisperResult.jsonText);
      } catch {
        // // Fall back to SRT parsing when Whisper JSON is unavailable or malformed on this host/runtime.
        cues = [];
      }
    }
    const fallbackCues = cues.length > 0 ? cues : parseSrt(whisperResult.srtText);
    if (!fallbackCues.length) {
      throw new Error(translate("error.emptyWhisper"));
    }

    setLog(`${translate("log.whisperDone")} ${whisperResult.model}`);
    return fallbackCues;
  } finally {
    if (cleanupAudioPath) {
      await deleteTemporaryWhisperAudio(cleanupAudioPath);
    }
  }
}

async function generate(): Promise<void> {
  // // Build the caption plan from selected source and push to Premiere host.
  if (generateInProgress) {
    return;
  }

  generateInProgress = true;
  setGenerateButtonsBusy(true);
  try {
    const options = collectBuildOptions();
    setLog(translate("log.processing"));
    await updateGenerateProgress(4, translate("progress.prepareGeneration"), true);
    const cues = await loadCuesFromSelectedSource(options, updateGenerateProgress);
    await updateGenerateProgress(90, translate("progress.planCaptions"), true);
    const plannedCues = buildCaptionPlan(cues, options);

    const payload: HostApplyPayload = {
      options,
      cues: plannedCues
    };

    await updateGenerateProgress(98, translate("progress.applyCaptions"), true);
    const hostResultRaw = await applyCaptionPlan(payload);
    setStructuredLogFromRaw(translate("log.hostResult"), hostResultRaw);
  } finally {
    generateInProgress = false;
    setGenerateButtonsBusy(false);
    setGenerateProgressState(false);
  }
}

async function initialize(): Promise<void> {
  // // Initialize locale, controls, and event listeners once panel is loaded.
  assertDomBindings();
  applyHostPanelTheme();
  bindHostThemeListener();
  await loadPanelMeta();
  refreshVersionLabel();
  const persistedState = readPersistedPanelState();

  const defaultLanguage =
    typeof persistedState.languageCode === "string" && persistedState.languageCode.length > 0
      ? persistedState.languageCode
      : navigator.language?.startsWith("fr")
        ? "fr"
        : "en";
  if (elements.languageSelect) {
    elements.languageSelect.value = hasSelectOption(elements.languageSelect, defaultLanguage) ? defaultLanguage : "en";
  }

  await loadLocale(elements.languageSelect?.value ?? "en");
  setVisualLiveUpdateEnabled(false, true);
  applyPersistedPanelState(persistedState);
  setActiveMode(activeMode);
  toggleSourceFields();
  renderMogrtGallery();
  persistPanelState();

  elements.languageSelect?.addEventListener("change", async () => {
    await loadLocale(elements.languageSelect?.value ?? "en");
    renderMogrtGallery();
    if (!loadedVisualProperties.length) {
      updateVisualSelectionSummary(translate("visual.selectionDefault"));
    }
    persistPanelState();
  });

  elements.tabGenerate?.addEventListener("click", () => {
    setActiveMode("generate");
    persistPanelState();
  });
  elements.tabVisual?.addEventListener("click", () => {
    setActiveMode("visual");
    persistPanelState();
  });

  elements.sourceMode?.addEventListener("change", () => {
    toggleSourceFields();
    persistPanelState();
  });

  elements.srtBrowseButton?.addEventListener("click", async () => {
    try {
      await browseSrtPath();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.mogrtAspectFilter?.addEventListener("change", () => {
    renderMogrtGallery();
    persistPanelState();
  });
  elements.mogrtFolderButton?.addEventListener("click", async () => {
    try {
      const extensionRootPath = resolveExtensionRootPath();
      await openInstalledMogrtFolder(extensionRootPath);
      persistPanelState();
    } catch (error) {
      setLog(String(error), true);
    }
  });
  elements.mogrtRefreshButton?.addEventListener("click", async () => {
    try {
      await reloadMogrtCatalogPreservingSelection();
      renderMogrtGallery();
      persistPanelState();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.animationMode?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.maxChars?.addEventListener("input", () => {
    persistPanelState();
  });
  elements.linesPerCaption?.addEventListener("input", () => {
    persistPanelState();
  });
  elements.fontSize?.addEventListener("input", () => {
    persistPanelState();
  });
  elements.srtPath?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.whisperModel?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.whisperSequenceRange?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.visualLiveUpdateButton?.addEventListener("click", () => {
    setVisualLiveUpdateEnabled(!visualLiveUpdateEnabled);
    if (visualLiveUpdateEnabled) {
      scheduleLiveVisualApply();
    }
  });
  elements.logToggleButton?.addEventListener("click", () => {
    setLogPanelExpanded(!logPanelExpanded);
  });
  elements.logVerbosityButton?.addEventListener("click", () => {
    setVerboseLogsEnabled(!verboseLogsEnabled);
  });
  elements.updateBanner?.addEventListener("click", (event) => {
    const targetNode = event.target as HTMLElement | null;
    if (targetNode?.closest("#updateLink")) {
      return;
    }
    event.preventDefault();
    void openUpdateDownload().catch((error) => {
      setLog(String(error), true);
    });
  });
  elements.updateLink?.addEventListener("click", (event) => {
    event.preventDefault();
    void openUpdateDownload().catch((error) => {
      setLog(String(error), true);
    });
  });
  window.addEventListener("focus", () => {
    schedulePassiveMogrtCatalogRefresh();
  });

  elements.generateButton?.addEventListener("click", async () => {
    try {
      persistPanelState();
      await generate();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.visualReadButton?.addEventListener("click", async () => {
    try {
      if (visualApplyInProgress) {
        return;
      }
      await loadVisualPropertiesFromSelection(true);
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.visualApplyButton?.addEventListener("click", async () => {
    try {
      if (visualApplyInProgress) {
        return;
      }
      await applyVisualChangesToSelection();
      persistPanelState();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  setLog(translate("log.ready"));

  void enforceWhisperSourceAvailability()
    .then(() => {
      toggleSourceFields();
      persistPanelState();
    })
    .catch(() => {
      // // Ignore Whisper detection failures at startup so the panel stays interactive immediately.
    });

  void reloadMogrtCatalogPreservingSelection()
    .then(() => {
      if (pendingSelectedMogrtId) {
        const restoredTemplate = availableMogrts.find((template) => template.id === pendingSelectedMogrtId);
        if (restoredTemplate) {
          selectedMogrt = restoredTemplate;
        }
        pendingSelectedMogrtId = "";
      }
      renderMogrtGallery();
      persistPanelState();
    })
    .catch((error) => {
      setLog(String(error), true);
    });

  void checkForUpdates().catch(() => {
    // // Ignore release-check failures at startup so network latency never blocks panel readiness.
  });
}

initialize().catch((error) => {
  setLog(String(error), true);
});
