// // Drive the Sub Creator panel UI and connect it to subtitle generation logic.
import { buildCaptionPlan } from "../core/planner";
import { parseWhisperJson } from "../core/whisper";
import { applyWhisperGlossaryToCues, buildWhisperGlossaryPrompt } from "../core/whisperGlossary";
import { normalizeWhisperXCuesForDisplay } from "../core/whisperxDisplay";
import { lockTranslatedCuesToSourceTiming } from "../core/timingLockedTranslation";
import { parseSrt, shiftCaptionCues, trimSrtCuesToRange } from "../core/srt";
import {
  buildTextEditorSafeApplyPlans,
  mergeTextEditorBlocks,
  moveTextEditorWord,
  prepareTextEditorBlocksForApply,
  retimeTextEditorBlocks,
  splitTextEditorBlock,
  updateTextEditorBlockText,
  type TextEditorBlock,
  type TextEditorTimingRange
} from "../core/textEditor";
import { tokenizeSubtitleText } from "../core/textNormalization";
import type {
  AnimationMode,
  CaptionBuildOptions,
  CaptionCue,
  CaptionWord,
  HostApplyPayload,
  MogrtTemplateItem,
  OutputMode,
  SourceMode,
  WhisperSequenceRangeMode
} from "../core/types";
import {
  applyCaptionPlan,
  applyNativeSubtitlePlan,
  buildPremiereTemplateCueMogrts,
  cancelCurrentJob,
  alignCorrectedTranscript,
  addVisualSelectionChangedListener,
  applyVisualPropertiesToSelectedMogrts,
  deleteTemporaryWhisperAudio,
  exportActiveSequenceAudioForWhisper,
  getActiveSequenceRange,
  getDeepLSupportedLanguages,
  getSelectedMogrtCount,
  getWhisperRuntimeStatus,
  openExternalUrl,
  openInstalledMogrtFolder,
  openWhisperModelsFolder,
  pickCorrectedTranscriptPath,
  pickSrtPath,
  readSelectedMogrtVisualSignature,
  readSelectedMogrtTextItems,
  readInstalledMogrtCatalog,
  readPremiereTemplateTextPayloads,
  readSelectedMogrtVisualProperties,
  readTextFileFromHost,
  readWhisperGlossaryStore,
  registerVisualSelectionWatcher,
  setVisualSelectionWatcherEnabled,
  transcribeWithWhisper,
  transcribeWithWhisperX,
  translateWithDeepL,
  writeWhisperGlossaryStore,
  applySelectedMogrtTextItems,
  isCancelledJobError
} from "./cepBridge";
import { applyPremierePanelTheme, bindPremiereThemeListener } from "./cepTheme";
import type {
  ApplySelectedMogrtTextResult,
  InstalledMogrtCatalog,
  SelectedMogrtVisualComponentDebug,
  TextEditorApplyPayload,
  WhisperProgressUpdate
} from "./cepBridge";
import {
  persistGeneratedCaptionMetadata,
  persistTextEditorCaptionMetadata,
  resolveCaptionMetadataForSelection,
  type CaptionMetadataIdentity
} from "./captionMetadataStore";

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

type PanelMode = "generate" | "visual" | "text" | "translate";

interface OutputModeGenerationSettings {
  sourceMode: SourceMode;
  srtPath: string;
  correctedTranscriptPath: string;
  whisperModel: string;
  whisperLanguageCode: string;
  whisperSequenceRange: WhisperSequenceRangeMode;
  preserveTranslationTiming: boolean;
  preserveMixedLanguages: boolean;
  mixedLanguagePrompt: string;
  removePunctuation: boolean;
  animationMode: AnimationMode;
  maxCharsPerLine: number;
  maxWordsPerLine: number;
  linesPerCaption: number;
  mogrtAspectFilter: string;
  mogrtSearchQuery: string;
  selectedMogrtId: string;
}

interface HostVisualProperty {
  path: string;
  displayName: string;
  groupPath: string;
  valueType: "number" | "boolean" | "string" | "json";
  controlKind: "slider" | "number" | "checkbox" | "color" | "text" | "string" | "json" | "vector" | "select";
  cloneOnlyWhenDirty?: boolean;
  excludeFromClone?: boolean;
  fontToken?: string;
  options?: Array<{ value: number | string; label: string }>;
  styleOptionsByFamily?: Record<string, string[]>;
  vectorScale?: number[];
  vectorMode?: string;
  value: string | number | boolean;
  minValue?: number;
  maxValue?: number;
  stepValue?: number;
}

interface VisualEditorSectionNode {
  key: string;
  label: string;
  properties: HostVisualProperty[];
  children: VisualEditorSectionNode[];
}

interface VisualEditorMutableSectionNode extends Omit<VisualEditorSectionNode, "children"> {
  children: VisualEditorMutableSectionNode[];
  encounterOrder: number;
}

interface PanelStateSnapshot {
  languageCode: string;
  activeMode: PanelMode;
  outputMode: OutputMode;
  outputSettings?: Partial<Record<OutputMode, Partial<OutputModeGenerationSettings>>>;
  whisperLanguageCode?: string;
  sourceMode?: SourceMode;
  srtPath: string;
  correctedTranscriptPath: string;
  whisperModel: string;
  whisperSequenceRange: WhisperSequenceRangeMode;
  preserveTranslationTiming?: boolean;
  preserveMixedLanguages?: boolean;
  mixedLanguagePrompt?: string;
  whisperGlossaryEnabled?: boolean;
  whisperGlossary?: string;
  removePunctuation?: boolean;
  animationMode: AnimationMode;
  maxCharsPerLine: number;
  maxWordsPerLine: number;
  linesPerCaption: number;
  mogrtAspectFilter: string;
  mogrtSearchQuery: string;
  selectedMogrtId: string;
  visualLiveUpdate: boolean;
  logExpanded: boolean;
  verboseLogs: boolean;
  deeplApiKey?: string;
  translationAutoLoadGeneratedNativeSrt?: boolean;
}

interface TextEditorBlockState extends TextEditorBlock {
  editorId: string;
  selectedWordIndex: number;
}

interface ApplyCaptionPlanHostResult {
  ok?: boolean;
  error?: string;
  insertedMogrt?: number;
  insertedNativeSubtitles?: number;
  nativeSubtitleTrackCreated?: boolean;
  nativeSubtitleSrtPath?: string;
  videoTrackUsed?: number;
  projectDocumentId?: string;
  projectPath?: string;
  sequenceID?: string;
  sequenceName?: string;
}

const elements = {
  languageSelect: document.querySelector<HTMLSelectElement>("#languageSelect"),
  appVersion: document.querySelector<HTMLButtonElement>("#appVersion"),
  updateBanner: document.querySelector<HTMLElement>("#updateBanner"),
  updateLink: document.querySelector<HTMLAnchorElement>("#updateLink"),
  tabGenerate: document.querySelector<HTMLButtonElement>("#tabGenerate"),
  tabVisual: document.querySelector<HTMLButtonElement>("#tabVisual"),
  tabText: document.querySelector<HTMLButtonElement>("#tabText"),
  tabTranslate: document.querySelector<HTMLButtonElement>("#tabTranslate"),
  modeGenerate: document.querySelector<HTMLElement>("#modeGenerate"),
  modeVisual: document.querySelector<HTMLElement>("#modeVisual"),
  modeText: document.querySelector<HTMLElement>("#modeText"),
  modeTranslate: document.querySelector<HTMLElement>("#modeTranslate"),
  translationReadButton: document.querySelector<HTMLButtonElement>("#translationReadButton"),
  translationTranslateButton: document.querySelector<HTMLButtonElement>("#translationTranslateButton"),
  translationDuplicateButton: document.querySelector<HTMLButtonElement>("#translationDuplicateButton"),
  translationSelectionSummary: document.querySelector<HTMLElement>("#translationSelectionSummary"),
  translationSourceLanguage: document.querySelector<HTMLSelectElement>("#translationSourceLanguage"),
  translationTargetLanguage: document.querySelector<HTMLSelectElement>("#translationTargetLanguage"),
  translationLanguagesRefreshButton: document.querySelector<HTMLButtonElement>("#translationLanguagesRefreshButton"),
  translationInputMode: document.querySelector<HTMLSelectElement>("#translationInputMode"),
  translationSrtField: document.querySelector<HTMLElement>("#translationSrtField"),
  translationSrtPath: document.querySelector<HTMLInputElement>("#translationSrtPath"),
  translationSrtBrowseButton: document.querySelector<HTMLButtonElement>("#translationSrtBrowseButton"),
  translationAutoLoadGeneratedNativeSrt: document.querySelector<HTMLInputElement>("#translationAutoLoadGeneratedNativeSrt"),
  deeplApiKey: document.querySelector<HTMLInputElement>("#deeplApiKey"),
  deeplApiKeyLink: document.querySelector<HTMLAnchorElement>("#deeplApiKeyLink"),
  translationPreview: document.querySelector<HTMLElement>("#translationPreview"),
  sourceMode: document.querySelector<HTMLSelectElement>("#sourceMode"),
  outputMode: document.querySelector<HTMLSelectElement>("#outputMode"),
  srtInputField: document.querySelector<HTMLElement>("#srtInputField"),
  srtPath: document.querySelector<HTMLInputElement>("#srtPath"),
  srtBrowseButton: document.querySelector<HTMLButtonElement>("#srtBrowseButton"),
  correctedAlignField: document.querySelector<HTMLElement>("#correctedAlignField"),
  correctedTranscriptPath: document.querySelector<HTMLInputElement>("#correctedTranscriptPath"),
  correctedTranscriptBrowseButton: document.querySelector<HTMLButtonElement>("#correctedTranscriptBrowseButton"),
  correctedAlignHint: document.querySelector<HTMLElement>("#correctedAlignHint"),
  whisperField: document.querySelector<HTMLElement>("#whisperField"),
  whisperModelRow: document.querySelector<HTMLElement>("#whisperModelRow"),
  whisperModel: document.querySelector<HTMLSelectElement>("#whisperModel"),
  whisperModelFolderButton: document.querySelector<HTMLButtonElement>("#whisperModelFolderButton"),
  whisperModelHint: document.querySelector<HTMLElement>("#whisperModelHint"),
  sequenceAudioField: document.querySelector<HTMLElement>("#sequenceAudioField"),
  whisperLanguageField: document.querySelector<HTMLElement>("#whisperLanguageField"),
  whisperLanguage: document.querySelector<HTMLSelectElement>("#whisperLanguage"),
  whisperSequenceRange: document.querySelector<HTMLSelectElement>("#whisperSequenceRange"),
  translationTimingField: document.querySelector<HTMLElement>("#translationTimingField"),
  preserveTranslationTiming: document.querySelector<HTMLInputElement>("#preserveTranslationTiming"),
  translationTimingHint: document.querySelector<HTMLElement>("#translationTimingHint"),
  mixedLanguageField: document.querySelector<HTMLElement>("#mixedLanguageField"),
  preserveMixedLanguages: document.querySelector<HTMLInputElement>("#preserveMixedLanguages"),
  mixedLanguagePromptField: document.querySelector<HTMLElement>("#mixedLanguagePromptField"),
  mixedLanguagePrompt: document.querySelector<HTMLTextAreaElement>("#mixedLanguagePrompt"),
  removePunctuation: document.querySelector<HTMLInputElement>("#removePunctuation"),
  animationField: document.querySelector<HTMLElement>("#animationField"),
  animationMode: document.querySelector<HTMLSelectElement>("#animationMode"),
  maxChars: document.querySelector<HTMLInputElement>("#maxChars"),
  maxWords: document.querySelector<HTMLInputElement>("#maxWords"),
  linesPerCaption: document.querySelector<HTMLInputElement>("#linesPerCaption"),
  mogrtAspectFilter: document.querySelector<HTMLSelectElement>("#mogrtAspectFilter"),
  mogrtSearchInput: document.querySelector<HTMLInputElement>("#mogrtSearchInput"),
  mogrtFolderButton: document.querySelector<HTMLButtonElement>("#mogrtFolderButton"),
  mogrtRefreshButton: document.querySelector<HTMLButtonElement>("#mogrtRefreshButton"),
  mogrtGalleryField: document.querySelector<HTMLElement>("#mogrtGalleryField"),
  mogrtGallery: document.querySelector<HTMLElement>("#mogrtGallery"),
  mogrtSelectedLabel: document.querySelector<HTMLParagraphElement>("#mogrtSelectedLabel"),
  outputModeHint: document.querySelector<HTMLParagraphElement>("#outputModeHint"),
  visualCopyButton: document.querySelector<HTMLButtonElement>("#visualCopyButton"),
  visualApplyButton: document.querySelector<HTMLButtonElement>("#visualApplyButton"),
  visualApplyProgress: document.querySelector<HTMLElement>("#visualApplyProgress"),
  visualApplyProgressBar: document.querySelector<HTMLProgressElement>("#visualApplyProgressBar"),
  visualApplyProgressText: document.querySelector<HTMLElement>("#visualApplyProgressText"),
  visualSelectionSummary: document.querySelector<HTMLParagraphElement>("#visualSelectionSummary"),
  visualPropertyList: document.querySelector<HTMLElement>("#visualPropertyList"),
  textReadButton: document.querySelector<HTMLButtonElement>("#textReadButton"),
  textApplyButton: document.querySelector<HTMLButtonElement>("#textApplyButton"),
  textApplyProgress: document.querySelector<HTMLElement>("#textApplyProgress"),
  textApplyProgressBar: document.querySelector<HTMLProgressElement>("#textApplyProgressBar"),
  textApplyProgressText: document.querySelector<HTMLElement>("#textApplyProgressText"),
  textSelectionSummary: document.querySelector<HTMLParagraphElement>("#textSelectionSummary"),
  textEditorList: document.querySelector<HTMLElement>("#textEditorList"),
  generateButton: document.querySelector<HTMLButtonElement>("#generateButton"),
  generateStopButton: document.querySelector<HTMLButtonElement>("#generateStopButton"),
  generateProgress: document.querySelector<HTMLElement>("#generateProgress"),
  generateProgressBar: document.querySelector<HTMLProgressElement>("#generateProgressBar"),
  generateProgressText: document.querySelector<HTMLElement>("#generateProgressText"),
  logPanel: document.querySelector<HTMLElement>("#logPanel"),
  logToggleButton: document.querySelector<HTMLButtonElement>("#logToggleButton"),
  logVerbosityButton: document.querySelector<HTMLButtonElement>("#logVerbosityButton"),
  logOutput: document.querySelector<HTMLPreElement>("#logOutput")
};

interface PanelLogState {
  timestamp: string;
  plainText: string;
  structuredTitle: string;
  structuredPayload: unknown;
  isError: boolean;
}

let currentLocale: LocaleMap = {};
let availableMogrts: MogrtTemplateItem[] = [];
let selectedMogrt: MogrtTemplateItem | null = null;
let pendingMogrtAspectFilter = "";
let pendingMogrtSearchQuery = "";
const FALLBACK_PANEL_META: PanelMeta = {
  version: "0.0.0",
  repository: "CyrilG93/PremiereSubCreator",
  releaseApiUrl: "https://api.github.com/repos/CyrilG93/PremiereSubCreator/releases/latest",
  releasePageUrl: "https://github.com/CyrilG93/PremiereSubCreator/releases/latest"
};
const PRODUCT_PAGE_URL = "https://www.cyrilplugin.com/subcreator";
const DEEPL_API_KEY_PAGE_URL = "https://www.deepl.com/your-account/keys";
let panelMeta: PanelMeta = { ...FALLBACK_PANEL_META };
const updateState: UpdateState = {
  visible: false,
  latestVersion: "",
  downloadUrl: ""
};
const PANEL_STATE_STORAGE_KEY = "subcreator.panelState.v1";
let pendingSelectedMogrtId = "";
let activeMode: PanelMode = "generate";
let activeOutputMode: OutputMode = "mogrt";
let outputSettingsByMode: Record<OutputMode, OutputModeGenerationSettings> = {
  mogrt: createDefaultOutputModeSettings("mogrt"),
  premiere_subtitles: createDefaultOutputModeSettings("premiere_subtitles")
};
let loadedVisualProperties: HostVisualProperty[] = [];
let loadedVisualComponents: SelectedMogrtVisualComponentDebug[] = [];
let loadedVisualSelectionCount = 0;
let visualSelectionSummaryBase = "";
let visualLoadedRevision = 0;
let copiedVisualChanges: VisualPropertyChange[] = [];
let copiedVisualSourceRevision = 0;
let copiedVisualSourceClipCount = 0;
let copiedVisualSourcePropertyCount = 0;
const visualOriginalValuesByPath = new Map<string, string>();
const visualOpenGroups = new Set<string>();
const visualTextStyleTokenMapByBasePath = new Map<string, Record<string, string>>();
const visualDirtyPaths = new Set<string>();
let visualLiveUpdateTimer: number | null = null;
let visualLiveUpdateQueued = false;
let visualLiveUpdateInFlight = false;
let visualApplyInProgress = false;
let visualSelectionAutoRefreshTimer: number | null = null;
let visualSelectionPollTimer: number | null = null;
let visualSelectionRefreshInFlight = false;
let visualSelectionWatcherCleanup: (() => void) | null = null;
let lastVisualSelectionSignature = "";
let pendingVisualSelectionChangeNotice = false;
let logPanelExpanded = true;
let verboseLogsEnabled = false;
const logHistory: PanelLogState[] = [];
const MAX_LOG_HISTORY_ENTRIES = 200;
let passiveMogrtRefreshTimer: number | null = null;
let lastPassiveMogrtCatalogRefreshAt = 0;
let generateInProgress = false;
let generateCancelRequested = false;
let visualReadInProgress = false;
let textApplyInProgress = false;
let textReadInProgress = false;
let availableWhisperModels: string[] = [];
let whisperModelCachePaths: string[] = [];
let whisperRuntimeAvailable = false;
let whisperRuntimeDetails = "";
let correctedAlignAvailable = false;
let correctedAlignRuntimeDetails = "";
let pendingWhisperModelValue = "base";
let pendingWhisperLanguageValue = "auto";
let textEditorBlocks: TextEditorBlockState[] = [];
let textEditorOriginalBlocks: TextEditorBlock[] = [];
let textEditorSelectionSignature = "";
let textEditorSameTrack = true;
let textEditorVideoTrackIndex = -1;
let textEditorBlockIdCounter = 0;
let textEditorSelectionStartSeconds = 0;
let textEditorSelectionEndSeconds = 0;
let textEditorSelectionMetadataIdentity: CaptionMetadataIdentity | null = null;
let translationBlocks: TextEditorBlock[] = [];
let translationSelectionSignature = "";
let translationSameTrack = false;
let translatedSubtitleTexts: string[] = [];
let translationLanguagesLoadInProgress = false;
const textEditorPendingCommitTimers = new Map<string, number>();

const GENERATE_PROGRESS_MAX = 100;
const SUBCREATOR_GENERATE_CANCELLED_CODE = "SUBCREATOR_GENERATE_CANCELLED";
const VISUAL_SELECTION_AUTO_REFRESH_DEBOUNCE_MS = 350;
const WHISPER_GLOSSARY_MAX_LENGTH = 12000;
const WHISPER_GLOSSARY_SAVE_DELAY_MS = 300;
const VISUAL_SELECTION_POLL_INTERVAL_MS = 1000;
const VISUAL_LIVE_UPDATE_DEBOUNCE_MS = 220;
const VISUAL_COLOR_LIVE_UPDATE_DEBOUNCE_MS = 40;
const FLOATING_SELECT_ROW_HEIGHT_PX = 28;
const FLOATING_SELECT_MIN_ROWS = 4;
const FLOATING_SELECT_MAX_ROWS = 10;
let activeFloatingPanelSelect: {
  sourceSelect: HTMLSelectElement;
  overlayRoot: HTMLDivElement;
  cleanup: (restoreFocus?: boolean) => void;
} | null = null;
let whisperGlossarySaveTimer: number | null = null;
const SUBCREATOR_WHISPER_LANGUAGE_DEFINITIONS: Array<{ code: string; label: string }> = [
  ["af", "afrikaans"],
  ["am", "amharic"],
  ["ar", "arabic"],
  ["as", "assamese"],
  ["az", "azerbaijani"],
  ["ba", "bashkir"],
  ["be", "belarusian"],
  ["bg", "bulgarian"],
  ["bn", "bengali"],
  ["bo", "tibetan"],
  ["br", "breton"],
  ["bs", "bosnian"],
  ["ca", "catalan"],
  ["cs", "czech"],
  ["cy", "welsh"],
  ["da", "danish"],
  ["de", "german"],
  ["el", "greek"],
  ["en", "english"],
  ["es", "spanish"],
  ["et", "estonian"],
  ["eu", "basque"],
  ["fa", "persian"],
  ["fi", "finnish"],
  ["fo", "faroese"],
  ["fr", "french"],
  ["gl", "galician"],
  ["gu", "gujarati"],
  ["ha", "hausa"],
  ["haw", "hawaiian"],
  ["he", "hebrew"],
  ["hi", "hindi"],
  ["hr", "croatian"],
  ["ht", "haitian creole"],
  ["hu", "hungarian"],
  ["hy", "armenian"],
  ["id", "indonesian"],
  ["is", "icelandic"],
  ["it", "italian"],
  ["ja", "japanese"],
  ["jw", "javanese"],
  ["ka", "georgian"],
  ["kk", "kazakh"],
  ["km", "khmer"],
  ["kn", "kannada"],
  ["ko", "korean"],
  ["la", "latin"],
  ["lb", "luxembourgish"],
  ["ln", "lingala"],
  ["lo", "lao"],
  ["lt", "lithuanian"],
  ["lv", "latvian"],
  ["mg", "malagasy"],
  ["mi", "maori"],
  ["mk", "macedonian"],
  ["ml", "malayalam"],
  ["mn", "mongolian"],
  ["mr", "marathi"],
  ["ms", "malay"],
  ["mt", "maltese"],
  ["my", "myanmar"],
  ["ne", "nepali"],
  ["nl", "dutch"],
  ["nn", "nynorsk"],
  ["no", "norwegian"],
  ["oc", "occitan"],
  ["pa", "punjabi"],
  ["pl", "polish"],
  ["ps", "pashto"],
  ["pt", "portuguese"],
  ["ro", "romanian"],
  ["ru", "russian"],
  ["sa", "sanskrit"],
  ["sd", "sindhi"],
  ["si", "sinhala"],
  ["sk", "slovak"],
  ["sl", "slovenian"],
  ["sn", "shona"],
  ["so", "somali"],
  ["sq", "albanian"],
  ["sr", "serbian"],
  ["su", "sundanese"],
  ["sv", "swedish"],
  ["sw", "swahili"],
  ["ta", "tamil"],
  ["te", "telugu"],
  ["tg", "tajik"],
  ["th", "thai"],
  ["tk", "turkmen"],
  ["tl", "tagalog"],
  ["tr", "turkish"],
  ["tt", "tatar"],
  ["uk", "ukrainian"],
  ["ur", "urdu"],
  ["uz", "uzbek"],
  ["vi", "vietnamese"],
  ["yi", "yiddish"],
  ["yo", "yoruba"],
  ["yue", "cantonese"],
  ["zh", "chinese"]
]
  .map(([code, label]) => ({
    code,
    label: String(label).replace(/\b\w/g, (character) => character.toUpperCase())
  }))
  .sort((left, right) => left.label.localeCompare(right.label));

function createGenerateCancelledError(): Error {
  // // Normalize panel-side cancellation into one stable marker so generate can treat it as a non-error path.
  const error = new Error(SUBCREATOR_GENERATE_CANCELLED_CODE);
  error.name = SUBCREATOR_GENERATE_CANCELLED_CODE;
  return error;
}

function isGenerateCancelledError(error: unknown): boolean {
  // // Treat both panel-side and CEP-bridge cancellation markers as graceful stop events.
  if (isCancelledJobError(error)) {
    return true;
  }
  const errorName = error instanceof Error ? String(error.name || "").trim() : "";
  const errorMessage = error instanceof Error ? String(error.message || "").trim() : String(error ?? "").trim();
  return errorName === SUBCREATOR_GENERATE_CANCELLED_CODE || errorMessage === SUBCREATOR_GENERATE_CANCELLED_CODE;
}

function assertGenerateNotCancelled(): void {
  // // Abort the generate pipeline as soon as the user requested a stop from the panel UI.
  if (generateCancelRequested) {
    throw createGenerateCancelledError();
  }
}

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

function assertDomBindings(): void {
  // // Guard against missing panel DOM ids during development/build changes.
  const missing = Object.entries(elements)
    .filter(([, value]) => !value)
    .map(([key]) => key);

  if (missing.length > 0) {
    throw new Error(`Missing DOM elements: ${missing.join(", ")}`);
  }
}

function waitForNextPaint(): Promise<void> {
  // // Yield past one animation frame so progress-bar/state changes can paint before lengthy async work continues.
  return new Promise((resolve) => {
    window.requestAnimationFrame(() => window.setTimeout(resolve, 0));
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

function getFloatingSelectViewportSpaceBelow(select: HTMLSelectElement): number {
  // // Measure the visible space below one select so long CEP dropdowns can decide whether to open downward.
  const rect = select.getBoundingClientRect();
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
  return Math.max(0, viewportHeight - rect.bottom - 8);
}

function getFloatingSelectViewportSpaceAbove(select: HTMLSelectElement): number {
  // // Measure the visible space above one select so the fallback popover can open upward when needed.
  const rect = select.getBoundingClientRect();
  return Math.max(0, rect.top - 8);
}

function getDesiredFloatingSelectRows(select: HTMLSelectElement): number {
  // // Keep floating select lists compact and predictable even for very long option sets.
  return Math.min(
    Math.max(FLOATING_SELECT_MIN_ROWS, Math.min(select.options.length || FLOATING_SELECT_MIN_ROWS, FLOATING_SELECT_MAX_ROWS)),
    FLOATING_SELECT_MAX_ROWS
  );
}

function getDesiredFloatingSelectHeight(rows: number): number {
  // // Convert row counts into one stable popover height so placement logic stays simple.
  return rows * FLOATING_SELECT_ROW_HEIGHT_PX + 12;
}

function shouldUseFloatingSelect(select: HTMLSelectElement): boolean {
  // // Prefer the custom popover for long lists and for selects that do not have enough room below in CEP.
  const desiredHeight = getDesiredFloatingSelectHeight(getDesiredFloatingSelectRows(select));
  return (select.options.length || 0) > FLOATING_SELECT_MAX_ROWS || getFloatingSelectViewportSpaceBelow(select) < desiredHeight;
}

function getFloatingSelectRowsForHeight(availableHeight: number, optionCount: number): number {
  // // Fit as many rows as safely visible while keeping the fallback list usable in tight panels.
  const rowsFromHeight = Math.floor((Math.max(availableHeight, 84) - 12) / FLOATING_SELECT_ROW_HEIGHT_PX);
  return Math.max(2, Math.min(optionCount || FLOATING_SELECT_MIN_ROWS, Math.min(FLOATING_SELECT_MAX_ROWS, Math.max(2, rowsFromHeight))));
}

function closeFloatingPanelSelect(restoreFocus = false): void {
  // // Destroy the active floating select popover before opening another one or when viewport context changes.
  if (!activeFloatingPanelSelect) {
    return;
  }

  const currentOverlay = activeFloatingPanelSelect;
  activeFloatingPanelSelect = null;
  currentOverlay.cleanup(restoreFocus);
}

function syncNativeSelectValue(sourceSelect: HTMLSelectElement, nextValue: string): void {
  // // Mirror one floating-select choice back into the real select so existing handlers keep working unchanged.
  if (sourceSelect.value === nextValue) {
    return;
  }

  sourceSelect.value = nextValue;
  sourceSelect.dispatchEvent(new Event("change", { bubbles: true }));
}

function openFloatingPanelSelect(sourceSelect: HTMLSelectElement): void {
  // // Render one viewport-fixed popover so CEP never depends on unreliable native dropdown placement.
  closeFloatingPanelSelect(false);

  const optionCount = Math.max(1, sourceSelect.options.length || 1);
  const desiredRows = getDesiredFloatingSelectRows(sourceSelect);
  const desiredHeight = getDesiredFloatingSelectHeight(desiredRows);
  const spaceBelow = getFloatingSelectViewportSpaceBelow(sourceSelect);
  const spaceAbove = getFloatingSelectViewportSpaceAbove(sourceSelect);
  const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
  const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
  const rect = sourceSelect.getBoundingClientRect();

  let placement: "below" | "above" | "center" = "below";
  let overlayRows = desiredRows;
  if (spaceBelow >= desiredHeight) {
    placement = "below";
  } else if (spaceAbove >= desiredHeight) {
    placement = "above";
  } else if (spaceAbove > spaceBelow) {
    placement = "above";
    overlayRows = getFloatingSelectRowsForHeight(spaceAbove, optionCount);
  } else if (spaceBelow > 0) {
    placement = "below";
    overlayRows = getFloatingSelectRowsForHeight(spaceBelow, optionCount);
  } else {
    placement = "center";
    overlayRows = getFloatingSelectRowsForHeight(Math.max(spaceAbove, spaceBelow, viewportHeight - 24), optionCount);
  }

  const overlayHeight = getDesiredFloatingSelectHeight(overlayRows);
  const overlayWidth = Math.max(rect.width, 160);
  let overlayTop = rect.bottom + 4;
  if (placement === "above") {
    overlayTop = rect.top - overlayHeight - 4;
  } else if (placement === "center") {
    overlayTop = rect.top + rect.height / 2 - overlayHeight / 2;
  }

  overlayTop = Math.max(8, Math.min(overlayTop, Math.max(8, viewportHeight - overlayHeight - 8)));
  const overlayLeft = Math.max(8, Math.min(rect.left, Math.max(8, viewportWidth - overlayWidth - 8)));

  const overlayRoot = document.createElement("div");
  overlayRoot.className = "panel-floating-select";
  overlayRoot.tabIndex = -1;
  overlayRoot.style.left = `${Math.round(overlayLeft)}px`;
  overlayRoot.style.top = `${Math.round(overlayTop)}px`;
  overlayRoot.style.width = `${Math.round(overlayWidth)}px`;
  overlayRoot.style.maxHeight = `${Math.round(overlayHeight)}px`;

  const overlayList = document.createElement("div");
  overlayList.className = "panel-floating-select__list";
  overlayRoot.appendChild(overlayList);

  const sourceOptions = Array.from(sourceSelect.options);
  const selectedIndex = Math.max(0, sourceOptions.findIndex((option) => option.value === sourceSelect.value));
  let highlightedIndex = selectedIndex;
  const overlayItems: HTMLButtonElement[] = [];

  const renderOverlayState = (): void => {
    // // Keep keyboard and pointer navigation in sync with the currently highlighted floating option.
    for (let itemIndex = 0; itemIndex < overlayItems.length; itemIndex += 1) {
      const button = overlayItems[itemIndex];
      const sourceOption = sourceOptions[itemIndex];
      button.classList.toggle("is-highlighted", itemIndex === highlightedIndex);
      button.classList.toggle("is-selected", Boolean(sourceOption?.value === sourceSelect.value));
    }

    const highlightedItem = overlayItems[highlightedIndex];
    if (highlightedItem) {
      highlightedItem.scrollIntoView({ block: "nearest" });
    }
  };

  const chooseIndex = (itemIndex: number): void => {
    // // Commit one floating option back into the native field before closing the popover.
    const sourceOption = sourceOptions[itemIndex];
    if (!sourceOption) {
      return;
    }

    syncNativeSelectValue(sourceSelect, sourceOption.value);
    closeFloatingPanelSelect(true);
  };

  for (let optionIndex = 0; optionIndex < sourceOptions.length; optionIndex += 1) {
    const sourceOption = sourceOptions[optionIndex];
    const optionButton = document.createElement("button");
    optionButton.type = "button";
    optionButton.className = "panel-floating-select__option";
    optionButton.textContent = sourceOption.textContent || sourceOption.label || sourceOption.value;
    optionButton.dataset.optionValue = sourceOption.value;
    optionButton.addEventListener("mouseenter", () => {
      highlightedIndex = optionIndex;
      renderOverlayState();
    });
    optionButton.addEventListener("click", () => {
      chooseIndex(optionIndex);
    });
    overlayItems.push(optionButton);
    overlayList.appendChild(optionButton);
  }

  document.body.appendChild(overlayRoot);
  const openedAt = Date.now();
  renderOverlayState();

  const cleanup = (restoreFocus = false): void => {
    // // Remove the temporary popover and every listener tied to this one floating select instance.
    document.removeEventListener("mousedown", onDocumentPointerDown, true);
    document.removeEventListener("touchstart", onDocumentPointerDown, true);
    window.removeEventListener("resize", onViewportChanged, true);
    document.removeEventListener("scroll", onViewportChanged, true);
    overlayRoot.removeEventListener("keydown", onKeydown);
    if (overlayRoot.parentElement) {
      overlayRoot.parentElement.removeChild(overlayRoot);
    }
    if (restoreFocus) {
      sourceSelect.focus();
    }
  };

  const onViewportChanged = (event?: Event): void => {
    const scrollTarget = event?.target as Node | null;
    if (scrollTarget && overlayRoot.contains(scrollTarget)) {
      return;
    }
    closeFloatingPanelSelect(false);
  };

  const onDocumentPointerDown = (event: Event): void => {
    // // Ignore the opening click itself; CEP can otherwise close the popover right after opening it.
    if (Date.now() - openedAt < 120) {
      return;
    }
    const target = event.target as Node | null;
    if (target && (overlayRoot.contains(target) || sourceSelect.contains(target))) {
      return;
    }
    closeFloatingPanelSelect(false);
  };

  const onKeydown = (event: KeyboardEvent): void => {
    if (event.key === "Escape") {
      event.preventDefault();
      closeFloatingPanelSelect(true);
      return;
    }
    if (event.key === "ArrowDown") {
      event.preventDefault();
      highlightedIndex = Math.min(sourceOptions.length - 1, highlightedIndex + 1);
      renderOverlayState();
      return;
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      highlightedIndex = Math.max(0, highlightedIndex - 1);
      renderOverlayState();
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      chooseIndex(highlightedIndex);
      return;
    }
    if (event.key === "Home") {
      event.preventDefault();
      highlightedIndex = 0;
      renderOverlayState();
      return;
    }
    if (event.key === "End") {
      event.preventDefault();
      highlightedIndex = Math.max(0, sourceOptions.length - 1);
      renderOverlayState();
    }
  };

  document.addEventListener("mousedown", onDocumentPointerDown, true);
  document.addEventListener("touchstart", onDocumentPointerDown, true);
  window.addEventListener("resize", onViewportChanged, true);
  document.addEventListener("scroll", onViewportChanged, true);
  overlayRoot.addEventListener("keydown", onKeydown);

  activeFloatingPanelSelect = {
    sourceSelect,
    overlayRoot,
    cleanup
  };

  overlayRoot.focus();
}

function tryOpenFloatingPanelSelect(select: HTMLSelectElement, event?: Event): void {
  // // Intercept native dropdown opening only when the viewport or option count would make it unreliable in CEP.
  if ((activeFloatingPanelSelect && activeFloatingPanelSelect.sourceSelect === select) || Number(select.size || 1) > 1) {
    return;
  }

  if (!shouldUseFloatingSelect(select)) {
    return;
  }

  if (event) {
    event.preventDefault();
    event.stopPropagation();
  }
  select.blur();
  openFloatingPanelSelect(select);
}

function bindFloatingPanelSelect(select: HTMLSelectElement | null | undefined): void {
  // // Reuse the same floating-select fallback for static panel controls and dynamic visual-editor selects.
  if (!select || select.dataset.floatingPanelSelectBound === "1") {
    return;
  }

  select.dataset.floatingPanelSelectBound = "1";

  select.addEventListener("mousedown", (event) => {
    tryOpenFloatingPanelSelect(select, event);
  });
  select.addEventListener("touchstart", (event) => {
    tryOpenFloatingPanelSelect(select, event);
  });
  select.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
      tryOpenFloatingPanelSelect(select, event);
    }
  });
  select.addEventListener("click", (event) => {
    if (!shouldUseFloatingSelect(select)) {
      return;
    }
    event.preventDefault();
    event.stopPropagation();
    if (!activeFloatingPanelSelect || activeFloatingPanelSelect.sourceSelect !== select) {
      openFloatingPanelSelect(select);
    }
  });
}

function sortWhisperModels(models: string[]): string[] {
  // // Keep Whisper models in a stable, user-facing order instead of cache-discovery order.
  const preferredOrder = ["tiny", "base", "small", "medium", "large-v3", "turbo"];
  return models
    .slice()
    .sort((left, right) => {
      const leftIndex = preferredOrder.indexOf(left);
      const rightIndex = preferredOrder.indexOf(right);
      if (leftIndex >= 0 && rightIndex >= 0) {
        return leftIndex - rightIndex;
      }
      if (leftIndex >= 0) {
        return -1;
      }
      if (rightIndex >= 0) {
        return 1;
      }
      return left.localeCompare(right);
    });
}

function refreshWhisperModelUi(preferredValue = pendingWhisperModelValue): void {
  // // Rebuild the Whisper model select from locally installed models only and expose a README hint when none exist.
  if (!elements.whisperModel) {
    return;
  }

  const installedModels = sortWhisperModels(availableWhisperModels);
  const desiredValue = String(preferredValue || elements.whisperModel.value || "").trim();
  elements.whisperModel.innerHTML = "";

  if (installedModels.length < 1) {
    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = translate("whisper.noInstalledModels");
    elements.whisperModel.appendChild(emptyOption);
    elements.whisperModel.value = "";
  } else {
    for (const model of installedModels) {
      const option = document.createElement("option");
      option.value = model;
      option.textContent = model;
      elements.whisperModel.appendChild(option);
    }
    elements.whisperModel.value = hasSelectOption(elements.whisperModel, desiredValue)
      ? desiredValue
      : installedModels.includes("base")
        ? "base"
        : installedModels[0];
  }

  if (elements.whisperModelHint) {
    const mode = getSourceMode();
    const whisperModeActive = mode === "whisper_sequence";
    const whisperxModeActive = mode === "whisperx_sequence";
    const shouldShowMissingModels = (whisperModeActive || whisperxModeActive) && installedModels.length < 1;
    const shouldShowMissingRuntime = whisperModeActive && !whisperRuntimeAvailable;
    const shouldShowMissingWhisperx = whisperxModeActive && !correctedAlignAvailable;
    elements.whisperModelHint.hidden = !shouldShowMissingModels && !shouldShowMissingRuntime && !shouldShowMissingWhisperx;
    elements.whisperModelHint.textContent = shouldShowMissingRuntime
      ? translate("help.whisper")
      : shouldShowMissingWhisperx
        ? translate("help.whisperxMissing")
        : translate("help.whisperModelsMissing");
  }

  pendingWhisperModelValue = String(elements.whisperModel.value || "").trim() || pendingWhisperModelValue;
}

function getSelectedWhisperLanguageCode(): string {
  // // Keep Whisper language selection separate from UI locale and default to auto-detect when unset.
  return String(elements.whisperLanguage?.value || pendingWhisperLanguageValue || "auto").trim() || "auto";
}

function sanitizeWhisperGlossary(value: string): string {
  // // Preserve one-entry-per-line dictionary formatting while bounding user-profile storage.
  return String(value || "")
    .replace(/\r/g, "\n")
    .replace(/[\t ]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
    .slice(0, WHISPER_GLOSSARY_MAX_LENGTH);
}

function buildWhisperInitialPrompt(options: CaptionBuildOptions): string {
  // // Build one concise list of canonical spellings for Whisper's vocabulary guidance.
  if (!options.preserveMixedLanguages) {
    return "";
  }

  return buildWhisperGlossaryPrompt(sanitizeWhisperGlossary(options.mixedLanguagePrompt));
}

function scheduleWhisperGlossaryStoreSave(): void {
  // // Debounce user-profile writes while keeping localStorage updates immediate.
  if (whisperGlossarySaveTimer !== null) {
    window.clearTimeout(whisperGlossarySaveTimer);
  }

  whisperGlossarySaveTimer = window.setTimeout(() => {
    whisperGlossarySaveTimer = null;
    void writeWhisperGlossaryStore(
      Boolean(elements.preserveMixedLanguages?.checked),
      sanitizeWhisperGlossary(elements.mixedLanguagePrompt?.value || "")
    ).catch((error) => {
      setLog(String(error), true);
    });
  }, WHISPER_GLOSSARY_SAVE_DELAY_MS);
}

function resolveSpellcheckLanguageCode(): string {
  // // Prefer the explicit subtitle language when available, otherwise fall back to the current UI locale.
  const subtitleLanguageCode = getSelectedWhisperLanguageCode();
  if (subtitleLanguageCode && subtitleLanguageCode !== "auto") {
    return subtitleLanguageCode;
  }

  return String(elements.languageSelect?.value || document.documentElement.lang || "en").trim() || "en";
}

function refreshWhisperLanguageUi(preferredValue = pendingWhisperLanguageValue): void {
  // // Rebuild the Whisper language select from the official Whisper language-code list plus one auto-detect option.
  if (!elements.whisperLanguage) {
    return;
  }

  const desiredValue = String(preferredValue || elements.whisperLanguage.value || "auto").trim() || "auto";
  elements.whisperLanguage.innerHTML = "";

  const autoOption = document.createElement("option");
  autoOption.value = "auto";
  autoOption.textContent = translate("whisper.languageAuto");
  elements.whisperLanguage.appendChild(autoOption);

  for (const language of SUBCREATOR_WHISPER_LANGUAGE_DEFINITIONS) {
    const option = document.createElement("option");
    option.value = language.code;
    option.textContent = `${language.label} (${language.code})`;
    elements.whisperLanguage.appendChild(option);
  }

  elements.whisperLanguage.value = hasSelectOption(elements.whisperLanguage, desiredValue) ? desiredValue : "auto";
  pendingWhisperLanguageValue = getSelectedWhisperLanguageCode();
}

function refreshCorrectedAlignUi(): void {
  // // Keep corrected-align availability explicit because it depends on Python WhisperX rather than Whisper CLI/models.
  if (!elements.correctedAlignHint) {
    return;
  }

  const correctedModeActive = getSourceMode() === "corrected_align";
  elements.correctedAlignHint.hidden = !correctedModeActive || correctedAlignAvailable;
  elements.correctedAlignHint.textContent = translate("help.correctedAlignMissing");
}

function isCurrentSourceReady(): boolean {
  // // Keep Generate clickable so missing Whisper dependencies produce a visible diagnostic instead of a silent disabled button.
  const mode = getSourceMode();
  if (mode === "corrected_align") {
    return getSelectedWhisperLanguageCode() !== "auto";
  }
  return true;
}

function isOutputModeValue(value: unknown): value is OutputMode {
  // // Validate persisted output mode strings before using them as storage keys.
  return value === "mogrt" || value === "premiere_subtitles";
}

function isSourceModeValue(value: unknown): value is SourceMode {
  // // Validate persisted source mode strings before restoring them into the select.
  return value === "srt" || value === "whisper_sequence" || value === "whisperx_sequence" || value === "corrected_align";
}

function isAnimationModeValue(value: unknown): value is AnimationMode {
  // // Validate persisted animation mode strings before restoring MOGRT-specific behavior.
  return value === "word" || value === "line" || value === "none";
}

function isWhisperSequenceRangeValue(value: unknown): value is WhisperSequenceRangeMode {
  // // Validate persisted sequence range strings before restoring Whisper/In-Out behavior.
  return value === "entire_sequence" || value === "in_out";
}

function sanitizePersistedNumber(value: unknown, fallback: number): number {
  // // Keep older or corrupted localStorage values from pushing NaN into number fields.
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : fallback;
}

function createDefaultOutputModeSettings(mode: OutputMode): OutputModeGenerationSettings {
  // // Provide independent defaults so first-time MOGRT and native output states can diverge safely later.
  return {
    sourceMode: "srt",
    srtPath: "",
    correctedTranscriptPath: "",
    whisperModel: "base",
    whisperLanguageCode: "auto",
    whisperSequenceRange: "entire_sequence",
    preserveTranslationTiming: false,
    preserveMixedLanguages: true,
    mixedLanguagePrompt: "",
    removePunctuation: false,
    animationMode: mode === "premiere_subtitles" ? "none" : "line",
    maxCharsPerLine: 28,
    maxWordsPerLine: 12,
    linesPerCaption: 2,
    mogrtAspectFilter: "all",
    mogrtSearchQuery: "",
    selectedMogrtId: ""
  };
}

function normalizeOutputModeSettings(
  rawSettings: Partial<OutputModeGenerationSettings> | undefined,
  fallback: OutputModeGenerationSettings
): OutputModeGenerationSettings {
  // // Merge stored partial settings with defaults so migrations and future fields remain backwards-compatible.
  const raw = rawSettings && typeof rawSettings === "object" ? rawSettings : {};
  return {
    sourceMode: isSourceModeValue(raw.sourceMode) ? raw.sourceMode : fallback.sourceMode,
    srtPath: typeof raw.srtPath === "string" ? raw.srtPath : fallback.srtPath,
    correctedTranscriptPath:
      typeof raw.correctedTranscriptPath === "string" ? raw.correctedTranscriptPath : fallback.correctedTranscriptPath,
    whisperModel: typeof raw.whisperModel === "string" && raw.whisperModel.trim() ? raw.whisperModel.trim() : fallback.whisperModel,
    whisperLanguageCode:
      typeof raw.whisperLanguageCode === "string" && raw.whisperLanguageCode.trim()
        ? raw.whisperLanguageCode.trim()
        : fallback.whisperLanguageCode,
    whisperSequenceRange: isWhisperSequenceRangeValue(raw.whisperSequenceRange)
      ? raw.whisperSequenceRange
      : fallback.whisperSequenceRange,
    preserveTranslationTiming:
      typeof raw.preserveTranslationTiming === "boolean" ? raw.preserveTranslationTiming : fallback.preserveTranslationTiming,
    preserveMixedLanguages:
      typeof raw.preserveMixedLanguages === "boolean" ? raw.preserveMixedLanguages : fallback.preserveMixedLanguages,
    mixedLanguagePrompt:
      typeof raw.mixedLanguagePrompt === "string"
        ? sanitizeWhisperGlossary(raw.mixedLanguagePrompt)
        : fallback.mixedLanguagePrompt,
    removePunctuation: typeof raw.removePunctuation === "boolean" ? raw.removePunctuation : fallback.removePunctuation,
    animationMode: isAnimationModeValue(raw.animationMode) ? raw.animationMode : fallback.animationMode,
    maxCharsPerLine: sanitizePersistedNumber(raw.maxCharsPerLine, fallback.maxCharsPerLine),
    maxWordsPerLine: sanitizePersistedNumber(raw.maxWordsPerLine, fallback.maxWordsPerLine),
    linesPerCaption: sanitizePersistedNumber(raw.linesPerCaption, fallback.linesPerCaption),
    mogrtAspectFilter: typeof raw.mogrtAspectFilter === "string" && raw.mogrtAspectFilter ? raw.mogrtAspectFilter : fallback.mogrtAspectFilter,
    mogrtSearchQuery: typeof raw.mogrtSearchQuery === "string" ? raw.mogrtSearchQuery : fallback.mogrtSearchQuery,
    selectedMogrtId: typeof raw.selectedMogrtId === "string" ? raw.selectedMogrtId : fallback.selectedMogrtId
  };
}

function captureOutputModeSettingsFromControls(mode: OutputMode = activeOutputMode): OutputModeGenerationSettings {
  // // Snapshot all generation controls for the output mode that is currently being edited.
  const fallback = outputSettingsByMode[mode] || createDefaultOutputModeSettings(mode);
  return normalizeOutputModeSettings(
    {
      sourceMode: getSourceMode(),
      srtPath: elements.srtPath?.value || "",
      correctedTranscriptPath: elements.correctedTranscriptPath?.value || "",
      whisperModel: elements.whisperModel?.value || pendingWhisperModelValue || fallback.whisperModel,
      whisperLanguageCode: getSelectedWhisperLanguageCode(),
      whisperSequenceRange: (elements.whisperSequenceRange?.value as WhisperSequenceRangeMode) || fallback.whisperSequenceRange,
      preserveTranslationTiming: Boolean(elements.preserveTranslationTiming?.checked),
      preserveMixedLanguages: Boolean(elements.preserveMixedLanguages?.checked),
      mixedLanguagePrompt: sanitizeWhisperGlossary(elements.mixedLanguagePrompt?.value || ""),
      removePunctuation: Boolean(elements.removePunctuation?.checked),
      animationMode: (elements.animationMode?.value as AnimationMode) || fallback.animationMode,
      maxCharsPerLine: Number(elements.maxChars?.value),
      maxWordsPerLine: Number(elements.maxWords?.value),
      linesPerCaption: Number(elements.linesPerCaption?.value),
      mogrtAspectFilter: pendingMogrtAspectFilter || elements.mogrtAspectFilter?.value || fallback.mogrtAspectFilter,
      mogrtSearchQuery: pendingMogrtSearchQuery || elements.mogrtSearchInput?.value || "",
      selectedMogrtId: selectedMogrt?.id || pendingSelectedMogrtId || ""
    },
    fallback
  );
}

function captureActiveOutputModeSettings(): void {
  // // Save the visible controls into the in-memory bucket before switching output modes or persisting.
  outputSettingsByMode[activeOutputMode] = captureOutputModeSettingsFromControls(activeOutputMode);
}

function applyOutputModeSettingsToControls(mode: OutputMode): void {
  // // Restore the generation controls that belong to the selected output mode.
  const settings = outputSettingsByMode[mode] || createDefaultOutputModeSettings(mode);
  if (elements.sourceMode && hasSelectOption(elements.sourceMode, settings.sourceMode)) {
    elements.sourceMode.value = settings.sourceMode;
  }
  if (elements.srtPath) {
    elements.srtPath.value = settings.srtPath;
  }
  if (elements.correctedTranscriptPath) {
    elements.correctedTranscriptPath.value = settings.correctedTranscriptPath;
  }

  pendingWhisperModelValue = settings.whisperModel || pendingWhisperModelValue || "base";
  if (elements.whisperModel && hasSelectOption(elements.whisperModel, pendingWhisperModelValue)) {
    elements.whisperModel.value = pendingWhisperModelValue;
  }

  pendingWhisperLanguageValue = settings.whisperLanguageCode || "auto";
  if (elements.whisperLanguage && hasSelectOption(elements.whisperLanguage, pendingWhisperLanguageValue)) {
    elements.whisperLanguage.value = pendingWhisperLanguageValue;
  }

  if (elements.whisperSequenceRange && hasSelectOption(elements.whisperSequenceRange, settings.whisperSequenceRange)) {
    elements.whisperSequenceRange.value = settings.whisperSequenceRange;
  }
  if (elements.preserveTranslationTiming) {
    elements.preserveTranslationTiming.checked = Boolean(settings.preserveTranslationTiming);
  }
  // // The Whisper dictionary is global and must not change when switching MOGRT/native output settings.
  if (elements.removePunctuation) {
    elements.removePunctuation.checked = Boolean(settings.removePunctuation);
  }
  if (elements.animationMode && hasSelectOption(elements.animationMode, settings.animationMode)) {
    elements.animationMode.value = settings.animationMode;
  }
  if (elements.maxChars) {
    elements.maxChars.value = String(settings.maxCharsPerLine);
  }
  if (elements.maxWords) {
    elements.maxWords.value = String(settings.maxWordsPerLine);
  }
  if (elements.linesPerCaption) {
    elements.linesPerCaption.value = String(settings.linesPerCaption);
  }

  pendingMogrtAspectFilter = settings.mogrtAspectFilter || "all";
  if (elements.mogrtAspectFilter && hasSelectOption(elements.mogrtAspectFilter, pendingMogrtAspectFilter)) {
    elements.mogrtAspectFilter.value = pendingMogrtAspectFilter;
  }
  pendingMogrtSearchQuery = settings.mogrtSearchQuery || "";
  if (elements.mogrtSearchInput) {
    elements.mogrtSearchInput.value = pendingMogrtSearchQuery;
  }

  pendingSelectedMogrtId = settings.selectedMogrtId || "";
  if (pendingSelectedMogrtId) {
    const restoredTemplate = availableMogrts.find((template) => template.id === pendingSelectedMogrtId);
    if (restoredTemplate) {
      selectedMogrt = restoredTemplate;
    } else {
      selectedMogrt = null;
    }
  } else {
    selectedMogrt = null;
  }
}

function buildLegacyOutputSettings(snapshot: Partial<PanelStateSnapshot>, fallback: OutputModeGenerationSettings): OutputModeGenerationSettings {
  // // Convert the previous flat localStorage schema into the new per-output bucket for the active output mode.
  return normalizeOutputModeSettings(
    {
      sourceMode: snapshot.sourceMode,
      srtPath: snapshot.srtPath,
      correctedTranscriptPath: snapshot.correctedTranscriptPath,
      whisperModel: snapshot.whisperModel,
      whisperLanguageCode: snapshot.whisperLanguageCode,
      whisperSequenceRange: snapshot.whisperSequenceRange,
      preserveTranslationTiming: snapshot.preserveTranslationTiming,
      preserveMixedLanguages: snapshot.preserveMixedLanguages,
      mixedLanguagePrompt: snapshot.mixedLanguagePrompt,
      removePunctuation: snapshot.removePunctuation,
      animationMode: snapshot.animationMode,
      maxCharsPerLine: snapshot.maxCharsPerLine,
      maxWordsPerLine: snapshot.maxWordsPerLine,
      linesPerCaption: snapshot.linesPerCaption,
      mogrtAspectFilter: snapshot.mogrtAspectFilter,
      mogrtSearchQuery: snapshot.mogrtSearchQuery,
      selectedMogrtId: snapshot.selectedMogrtId
    },
    fallback
  );
}

function hydrateOutputModeSettings(snapshot: Partial<PanelStateSnapshot>): void {
  // // Initialize both output-mode buckets, migrating old flat settings only into the currently selected mode.
  const activeModeFromSnapshot = isOutputModeValue(snapshot.outputMode) ? snapshot.outputMode : "mogrt";
  const storedSettings = snapshot.outputSettings && typeof snapshot.outputSettings === "object" ? snapshot.outputSettings : {};
  const legacyActiveSettings = buildLegacyOutputSettings(snapshot, createDefaultOutputModeSettings(activeModeFromSnapshot));

  outputSettingsByMode = {
    mogrt: normalizeOutputModeSettings(
      storedSettings.mogrt || (activeModeFromSnapshot === "mogrt" ? legacyActiveSettings : undefined),
      createDefaultOutputModeSettings("mogrt")
    ),
    premiere_subtitles: normalizeOutputModeSettings(
      storedSettings.premiere_subtitles || (activeModeFromSnapshot === "premiere_subtitles" ? legacyActiveSettings : undefined),
      createDefaultOutputModeSettings("premiere_subtitles")
    )
  };

  activeOutputMode = activeModeFromSnapshot;
  if (elements.outputMode && hasSelectOption(elements.outputMode, activeOutputMode)) {
    elements.outputMode.value = activeOutputMode;
  }
  applyOutputModeSettingsToControls(activeOutputMode);
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
    !elements.outputMode ||
    !elements.srtPath ||
    !elements.whisperModel ||
    !elements.whisperLanguage ||
    !elements.whisperSequenceRange ||
    !elements.preserveMixedLanguages ||
    !elements.mixedLanguagePrompt ||
    !elements.removePunctuation ||
    !elements.animationMode ||
    !elements.maxChars ||
    !elements.maxWords ||
    !elements.linesPerCaption ||
    !elements.mogrtAspectFilter ||
    !elements.mogrtSearchInput
  ) {
    return;
  }

  captureActiveOutputModeSettings();
  const snapshot: PanelStateSnapshot = {
    languageCode: elements.languageSelect.value || "en",
    activeMode,
    outputMode: activeOutputMode,
    outputSettings: {
      mogrt: outputSettingsByMode.mogrt,
      premiere_subtitles: outputSettingsByMode.premiere_subtitles
    },
    whisperLanguageCode: outputSettingsByMode[activeOutputMode].whisperLanguageCode,
    sourceMode: outputSettingsByMode[activeOutputMode].sourceMode,
    srtPath: outputSettingsByMode[activeOutputMode].srtPath,
    correctedTranscriptPath: outputSettingsByMode[activeOutputMode].correctedTranscriptPath,
    whisperModel: outputSettingsByMode[activeOutputMode].whisperModel,
    whisperSequenceRange: outputSettingsByMode[activeOutputMode].whisperSequenceRange,
    preserveTranslationTiming: outputSettingsByMode[activeOutputMode].preserveTranslationTiming,
    preserveMixedLanguages: outputSettingsByMode[activeOutputMode].preserveMixedLanguages,
    mixedLanguagePrompt: outputSettingsByMode[activeOutputMode].mixedLanguagePrompt,
    whisperGlossaryEnabled: Boolean(elements.preserveMixedLanguages.checked),
    whisperGlossary: sanitizeWhisperGlossary(elements.mixedLanguagePrompt.value),
    removePunctuation: outputSettingsByMode[activeOutputMode].removePunctuation,
    animationMode: outputSettingsByMode[activeOutputMode].animationMode,
    maxCharsPerLine: outputSettingsByMode[activeOutputMode].maxCharsPerLine,
    maxWordsPerLine: outputSettingsByMode[activeOutputMode].maxWordsPerLine,
    linesPerCaption: outputSettingsByMode[activeOutputMode].linesPerCaption,
    mogrtAspectFilter: outputSettingsByMode[activeOutputMode].mogrtAspectFilter,
    mogrtSearchQuery: outputSettingsByMode[activeOutputMode].mogrtSearchQuery,
    selectedMogrtId: outputSettingsByMode[activeOutputMode].selectedMogrtId,
    visualLiveUpdate: true,
    logExpanded: logPanelExpanded,
    verboseLogs: verboseLogsEnabled,
    // // Keep the user's DeepL key in the local CEP profile so it survives a Premiere restart.
    deeplApiKey: String(elements.deeplApiKey?.value || ""),
    // // Keep automatic native-SRT preparation opt-in and enabled by default for new profiles.
    translationAutoLoadGeneratedNativeSrt: elements.translationAutoLoadGeneratedNativeSrt?.checked !== false
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

  hydrateOutputModeSettings(snapshot);

  const activeSettings = outputSettingsByMode[activeOutputMode];
  if (elements.preserveMixedLanguages) {
    elements.preserveMixedLanguages.checked =
      typeof snapshot.whisperGlossaryEnabled === "boolean" ? snapshot.whisperGlossaryEnabled : true;
  }
  if (elements.mixedLanguagePrompt) {
    const legacyGlossary =
      typeof snapshot.mixedLanguagePrompt === "string"
        ? snapshot.mixedLanguagePrompt
        : activeSettings?.mixedLanguagePrompt || "";
    elements.mixedLanguagePrompt.value = sanitizeWhisperGlossary(
      typeof snapshot.whisperGlossary === "string" ? snapshot.whisperGlossary : legacyGlossary
    );
  }

  if (typeof snapshot.logExpanded === "boolean") {
    setLogPanelExpanded(snapshot.logExpanded, true);
  }

  if (typeof snapshot.verboseLogs === "boolean") {
    setVerboseLogsEnabled(snapshot.verboseLogs, true);
  }

  if (elements.deeplApiKey && typeof snapshot.deeplApiKey === "string") {
    // // Restore the locally stored key without ever writing it to logs or host payloads.
    elements.deeplApiKey.value = snapshot.deeplApiKey;
  }

  if (elements.translationAutoLoadGeneratedNativeSrt) {
    // // Preserve the default enabled state when an older stored profile has no preference yet.
    elements.translationAutoLoadGeneratedNativeSrt.checked = snapshot.translationAutoLoadGeneratedNativeSrt !== false;
  }

  if (snapshot.activeMode === "visual" || snapshot.activeMode === "text" || snapshot.activeMode === "translate") {
    activeMode = snapshot.activeMode;
  } else {
    activeMode = "generate";
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
  // // Re-render the complete session history whenever locale, verbosity, or visibility changes.
  if (!elements.logOutput) {
    return;
  }

  if (logHistory.length < 1) {
    elements.logOutput.textContent = "";
    elements.logOutput.classList.remove("log--error");
    return;
  }

  elements.logOutput.textContent = logHistory
    .map((entry) => {
      let outputText = entry.plainText;
      if (entry.structuredTitle) {
        const payloadToRender = verboseLogsEnabled
          ? entry.structuredPayload
          : buildCompactLogValue(entry.structuredPayload);
        outputText = `${entry.structuredTitle}\n${JSON.stringify(payloadToRender, null, 2)}`;
      }
      return `[${entry.timestamp}] ${outputText}`;
    })
    .join("\n\n");
  elements.logOutput.classList.toggle("log--error", Boolean(logHistory[logHistory.length - 1]?.isError));
  elements.logOutput.scrollTop = elements.logOutput.scrollHeight;
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
    elements.logVerbosityButton.hidden = !logPanelExpanded;
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
  // // Append structured payloads so verbosity changes can re-render the complete operation history.
  logHistory.push({
    timestamp: new Date().toLocaleTimeString(),
    plainText: "",
    structuredTitle: title,
    structuredPayload: payload,
    isError
  });
  if (logHistory.length > MAX_LOG_HISTORY_ENTRIES) {
    logHistory.splice(0, logHistory.length - MAX_LOG_HISTORY_ENTRIES);
  }
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
  elements.appVersion.setAttribute("aria-label", `Open Sub Creator page for version ${displayVersion}`);
}

async function openProductPage(): Promise<void> {
  // // Route version-badge clicks through the same CEP-safe external URL opener as update links.
  await openExternalUrl(PRODUCT_PAGE_URL);
}

async function openDeepLApiKeyPage(): Promise<void> {
  // // Open DeepL's account-key page in the default browser because CEP panels cannot safely navigate there in place.
  await openExternalUrl(DEEPL_API_KEY_PAGE_URL);
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
  // // Append runtime status and error traces so every generation stage remains visible.
  logHistory.push({
    timestamp: new Date().toLocaleTimeString(),
    plainText: message,
    structuredTitle: "",
    structuredPayload: null,
    isError
  });
  if (logHistory.length > MAX_LOG_HISTORY_ENTRIES) {
    logHistory.splice(0, logHistory.length - MAX_LOG_HISTORY_ENTRIES);
  }
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
  document.documentElement.lang = String(languageCode || "en").trim() || "en";

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
  if (elements.textApplyProgress && !elements.textApplyProgress.hidden && elements.textApplyProgressBar) {
    setTextApplyProgressState(true, Number(elements.textApplyProgressBar.value || 0), Number(elements.textApplyProgressBar.max || 0));
  }
  if (textEditorBlocks.length > 0) {
    renderTextEditor();
  } else {
    setTextSelectionSummary(translate("text.selectionDefault"));
  }
  refreshLogControlsState();
  renderCurrentLog();
  refreshMogrtAspectFilterOptions();
  refreshWhisperModelUi();
  refreshWhisperLanguageUi(getSelectedWhisperLanguageCode());
  refreshCorrectedAlignUi();

  refreshUpdateBanner();
}

function getSourceMode(): SourceMode {
  // // Normalize source mode value from UI select control.
  return (elements.sourceMode?.value as SourceMode) || "srt";
}

function getOutputMode(): OutputMode {
  // // Normalize output mode so generation can switch between animated MOGRTs and native Premiere captions.
  const value = elements.outputMode?.value;
  return value === "premiere_subtitles" ? "premiere_subtitles" : "mogrt";
}

function isNativeSubtitleOutputMode(): boolean {
  // // Keep native-caption UI checks centralized because this mode skips all MOGRT-specific controls.
  return getOutputMode() === "premiere_subtitles";
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

function normalizeMogrtTemplateNameKey(value: string): string {
  // // Normalize template labels so the Text editor can map timeline clip names back to installed MOGRT files.
  return String(value || "")
    .trim()
    .replace(/\.mogrt$/i, "")
    .toLowerCase();
}

function resolveTextEditorTemplateFromSelection(): MogrtTemplateItem | null {
  // // Prefer the actual template read from the selected subtitle clips so safe text rebuild does not depend on gallery state.
  const uniqueClipNames = Array.from(
    new Set(
      textEditorOriginalBlocks
        .map((block) => normalizeMogrtTemplateNameKey(block.clipName))
        .filter((clipName) => clipName.length > 0)
    )
  );
  if (uniqueClipNames.length !== 1) {
    return null;
  }

  const clipNameKey = uniqueClipNames[0];
  if (selectedMogrt && normalizeMogrtTemplateNameKey(selectedMogrt.name) === clipNameKey) {
    return selectedMogrt;
  }

  const matches = availableMogrts.filter((template) => {
    const templateNameKey = normalizeMogrtTemplateNameKey(template.name);
    const relativePathLeaf = String(template.relativePath || "")
      .split("/")
      .pop()
      ?.replace(/\.[^.]+$/, "");
    const relativePathLeafKey = normalizeMogrtTemplateNameKey(relativePathLeaf || "");
    return templateNameKey === clipNameKey || relativePathLeafKey === clipNameKey;
  });

  return matches.length === 1 ? matches[0] || null : null;
}

function collectTextEditorBuildOptions(): CaptionBuildOptions {
  // // Override gallery selection with the template actually present on the timeline whenever it can be resolved safely.
  const options = collectBuildOptions();
  const textEditorTemplate = resolveTextEditorTemplateFromSelection();
  if (!textEditorTemplate) {
    return options;
  }

  return {
    ...options,
    outputMode: "mogrt",
    mogrtTemplateRelativePath: textEditorTemplate.relativePath,
    mogrtPath: buildAbsoluteMogrtPath(options.extensionRootPath, textEditorTemplate.relativePath)
  };
}

function toggleSourceFields(): void {
  // // Show only the source-related controls needed for current workflow.
  const mode = getSourceMode();
  const srtModeActive = mode === "srt";
  const whisperModeActive = mode === "whisper_sequence";
  const whisperxModeActive = mode === "whisperx_sequence";
  const correctedAlignModeActive = mode === "corrected_align";
  const sequenceRangeModeActive = srtModeActive || whisperModeActive || whisperxModeActive || correctedAlignModeActive;

  if (elements.srtInputField) {
    elements.srtInputField.style.display = mode === "srt" ? "grid" : "none";
  }

  if (elements.correctedAlignField) {
    elements.correctedAlignField.style.display = correctedAlignModeActive ? "grid" : "none";
  }

  if (elements.whisperField) {
    elements.whisperField.style.display = whisperModeActive || whisperxModeActive ? "grid" : "none";
  }

  if (elements.sequenceAudioField) {
    elements.sequenceAudioField.style.display = sequenceRangeModeActive ? "grid" : "none";
  }

  if (elements.mixedLanguageField) {
    elements.mixedLanguageField.style.display = whisperModeActive || whisperxModeActive ? "grid" : "none";
  }

  if (elements.mixedLanguagePromptField) {
    elements.mixedLanguagePromptField.style.display = elements.preserveMixedLanguages?.checked ? "grid" : "none";
  }

  if (elements.whisperLanguageField) {
    elements.whisperLanguageField.style.display = whisperModeActive || whisperxModeActive || correctedAlignModeActive ? "grid" : "none";
  }

  const canLockForcedLanguageTiming = whisperModeActive && getSelectedWhisperLanguageCode().toLowerCase() !== "auto";
  if (elements.translationTimingField) {
    elements.translationTimingField.style.display = canLockForcedLanguageTiming ? "flex" : "none";
  }
  if (elements.translationTimingHint) {
    elements.translationTimingHint.style.display = canLockForcedLanguageTiming ? "block" : "none";
  }
  if (elements.preserveTranslationTiming) {
    elements.preserveTranslationTiming.disabled = !canLockForcedLanguageTiming || generateInProgress;
    if (!canLockForcedLanguageTiming) {
      elements.preserveTranslationTiming.checked = false;
    }
  }

  if (elements.whisperSequenceRange) {
    elements.whisperSequenceRange.disabled = !sequenceRangeModeActive;
  }
  if (elements.whisperLanguage) {
    elements.whisperLanguage.disabled = !(whisperModeActive || whisperxModeActive || correctedAlignModeActive) || generateInProgress;
  }
  if (elements.whisperModelFolderButton) {
    elements.whisperModelFolderButton.disabled = !(whisperModeActive || whisperxModeActive) || generateInProgress;
  }
  if (elements.whisperModelRow) {
    elements.whisperModelRow.classList.toggle("is-single", !(whisperModeActive || whisperxModeActive));
  }
  refreshWhisperModelUi();
  refreshCorrectedAlignUi();
  toggleOutputFields();
  if (elements.generateButton && !generateInProgress) {
    elements.generateButton.disabled = !isCurrentSourceReady();
  }
}

function toggleOutputFields(): void {
  // // Hide MOGRT-only controls when creating native Premiere subtitle tracks from generated SRT data.
  const nativeSubtitleModeActive = isNativeSubtitleOutputMode();
  if (elements.animationField) {
    elements.animationField.style.display = nativeSubtitleModeActive ? "none" : "grid";
  }
  if (elements.mogrtGalleryField) {
    elements.mogrtGalleryField.style.display = nativeSubtitleModeActive ? "none" : "grid";
  }
  if (elements.outputModeHint) {
    elements.outputModeHint.textContent = translate(nativeSubtitleModeActive ? "help.nativeTracks" : "help.mogrtTracks");
  }
}

async function enforceWhisperSourceAvailability(): Promise<void> {
  // // Detect Whisper/WhisperX availability and refresh source-specific UI without removing source modes from the panel.
  try {
    const status = await getWhisperRuntimeStatus();
    whisperRuntimeAvailable = Boolean(status.available);
    whisperRuntimeDetails = String(status.details || "");
    availableWhisperModels = Array.isArray(status.installedModels) ? status.installedModels.slice() : [];
    whisperModelCachePaths = Array.isArray(status.modelCachePaths) ? status.modelCachePaths.slice() : [];
    correctedAlignAvailable = Boolean(status.alignmentAvailable);
    correctedAlignRuntimeDetails = String(status.alignmentDetails || "");
    refreshWhisperModelUi(pendingWhisperModelValue);
    refreshCorrectedAlignUi();
    toggleSourceFields();
  } catch (error) {
    whisperRuntimeAvailable = false;
    whisperRuntimeDetails = `Runtime detection failed: ${String(error)}`;
    availableWhisperModels = [];
    whisperModelCachePaths = [];
    correctedAlignAvailable = false;
    correctedAlignRuntimeDetails = "";
    refreshWhisperModelUi(pendingWhisperModelValue);
    refreshCorrectedAlignUi();
  }
}

function buildWhisperDiagnostic(): string {
  // // Build a compact shareable snapshot for customer machines where runtime detection differs from development.
  const selectedModel = String(elements.whisperModel?.value || pendingWhisperModelValue || "").trim();
  return [
    `Sub Creator: ${String(panelMeta.version || FALLBACK_PANEL_META.version || "unknown")}`,
    `Source: ${getSourceMode()}`,
    `Platform: ${String(navigator?.platform || "unknown")}`,
    `Whisper runtime: ${whisperRuntimeAvailable ? "available" : "unavailable"}`,
    `Runtime details: ${whisperRuntimeDetails || "none"}`,
    `WhisperX runtime: ${correctedAlignAvailable ? "available" : "unavailable"}`,
    `WhisperX details: ${correctedAlignRuntimeDetails || "none"}`,
    `Selected model: ${selectedModel || "none"}`,
    `Detected models: ${availableWhisperModels.join(", ") || "none"}`,
    `Model cache paths: ${whisperModelCachePaths.join(" | ") || "none"}`
  ].join("\n");
}

function setActiveMode(mode: PanelMode): void {
  // // Toggle tab state and active mode container visibility.
  activeMode = mode;
  if (mode !== "visual") {
    visualLiveUpdateQueued = false;
    clearVisualSelectionAutoRefreshTimer();
    stopVisualSelectionPolling();
    if (visualLiveUpdateTimer !== null) {
      window.clearTimeout(visualLiveUpdateTimer);
      visualLiveUpdateTimer = null;
    }
  } else if (isVisualSelectionMonitoringAllowed()) {
    startVisualSelectionPolling();
    scheduleVisualSelectionAutoRefresh("tab");
  }
  void syncVisualSelectionWatcherState();

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

  if (elements.tabText) {
    const isActive = mode === "text";
    elements.tabText.classList.toggle("is-active", isActive);
    elements.tabText.setAttribute("aria-selected", isActive ? "true" : "false");
  }
  if (elements.tabTranslate) {
    const isActive = mode === "translate";
    elements.tabTranslate.classList.toggle("is-active", isActive);
    elements.tabTranslate.setAttribute("aria-selected", isActive ? "true" : "false");
  }

  if (elements.modeGenerate) {
    elements.modeGenerate.hidden = mode !== "generate";
  }

  if (elements.modeVisual) {
    elements.modeVisual.hidden = mode !== "visual";
  }

  if (elements.modeText) {
    elements.modeText.hidden = mode !== "text";
  }
  if (elements.modeTranslate) {
    elements.modeTranslate.hidden = mode !== "translate";
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

function parseVisualComponentIndex(path: string): number {
  // // Recover the host component index from `c4|...` visual paths so duplicate Premiere groups do not collapse together.
  const match = /^c(\d+)\|/.exec(String(path || "").trim());
  return match ? Number(match[1] || 0) : 0;
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

function updateVisualSelectionSummary(message: string): void {
  // // Store the current selection summary and re-render with any active copied snapshot context.
  visualSelectionSummaryBase = String(message || "").trim();
  renderVisualSelectionSummary();
}

function renderVisualSelectionSummary(): void {
  // // Combine the live selection summary with copied-snapshot status so cross-selection clone state stays visible.
  if (!elements.visualSelectionSummary) {
    return;
  }

  const parts: string[] = [];
  if (visualSelectionSummaryBase) {
    parts.push(visualSelectionSummaryBase);
  }
  if (copiedVisualChanges.length > 0) {
    parts.push(
      translateTemplate("visual.copySummary", {
        clips: String(Math.max(1, copiedVisualSourceClipCount || 1)),
        props: String(Math.max(1, copiedVisualSourcePropertyCount || copiedVisualChanges.length))
      })
    );
  }
  if (pendingVisualSelectionChangeNotice) {
    parts.push(translate("visual.selectionChangedPending"));
  }

  elements.visualSelectionSummary.textContent = parts.join(" ").trim() || translate("visual.selectionDefault");
}

function cloneVisualPropertyChange(change: VisualPropertyChange): VisualPropertyChange {
  // // Deep-clone one change entry so copied snapshots remain immutable across future UI edits.
  return {
    ...change,
    vectorScale: Array.isArray(change.vectorScale) ? change.vectorScale.slice() : undefined
  };
}

function cloneVisualPropertyChanges(changes: VisualPropertyChange[]): VisualPropertyChange[] {
  // // Clone change arrays before storing or merging them to avoid accidental shared references.
  return changes.map((change) => cloneVisualPropertyChange(change));
}

function mergeVisualPropertyChanges(baseChanges: VisualPropertyChange[], overrideChanges: VisualPropertyChange[]): VisualPropertyChange[] {
  // // Overlay explicit UI edits on top of a copied snapshot while keeping one entry per property path.
  const mergedByPath = new Map<string, VisualPropertyChange>();
  baseChanges.forEach((change) => {
    mergedByPath.set(change.path, cloneVisualPropertyChange(change));
  });
  overrideChanges.forEach((change) => {
    mergedByPath.set(change.path, cloneVisualPropertyChange(change));
  });
  return Array.from(mergedByPath.values());
}

function updateCopiedVisualSnapshot(changes: VisualPropertyChange[]): void {
  // // Persist the last copied visual snapshot so Apply can target a new clip selection later on.
  copiedVisualChanges = cloneVisualPropertyChanges(changes);
  copiedVisualSourceRevision = visualLoadedRevision;
  copiedVisualSourceClipCount = loadedVisualSelectionCount;
  copiedVisualSourcePropertyCount = copiedVisualChanges.length;
  refreshVisualButtonsBusyState();
  renderVisualSelectionSummary();
}

function refreshVisualButtonsBusyState(): void {
  // // Prevent concurrent read/copy/apply actions while either selection loading or host writes are in progress.
  const isBusy = visualReadInProgress || visualApplyInProgress;
  if (elements.visualApplyButton) {
    elements.visualApplyButton.disabled = isBusy;
  }
  if (elements.visualCopyButton) {
    // // Keep `Copy properties` available even before a manual read because it now refreshes the current selection on click.
    elements.visualCopyButton.disabled = isBusy;
    elements.visualCopyButton.setAttribute("aria-pressed", copiedVisualChanges.length > 0 ? "true" : "false");
  }
}

function setGenerateButtonsBusy(isBusy: boolean): void {
  // // Prevent duplicate generate runs while export/transcription/apply is already active.
  if (elements.generateButton) {
    elements.generateButton.disabled = isBusy || !isCurrentSourceReady();
  }
  if (elements.generateStopButton) {
    elements.generateStopButton.hidden = !isBusy;
    elements.generateStopButton.disabled = !isBusy || generateCancelRequested;
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
  if (elements.correctedTranscriptPath) {
    elements.correctedTranscriptPath.disabled = isBusy;
  }
  if (elements.correctedTranscriptBrowseButton) {
    elements.correctedTranscriptBrowseButton.disabled = isBusy;
  }
  if (elements.sourceMode) {
    elements.sourceMode.disabled = isBusy;
  }
  if (elements.outputMode) {
    elements.outputMode.disabled = isBusy;
  }
  if (elements.whisperModel) {
    elements.whisperModel.disabled =
      isBusy || (getSourceMode() !== "whisper_sequence" && getSourceMode() !== "whisperx_sequence") || availableWhisperModels.length < 1;
  }
  if (elements.whisperModelFolderButton) {
    elements.whisperModelFolderButton.disabled =
      isBusy || (getSourceMode() !== "whisper_sequence" && getSourceMode() !== "whisperx_sequence");
  }
  if (elements.whisperSequenceRange) {
    elements.whisperSequenceRange.disabled =
      isBusy ||
      (getSourceMode() !== "whisper_sequence" && getSourceMode() !== "whisperx_sequence" && getSourceMode() !== "corrected_align");
  }
  if (elements.whisperLanguage) {
    elements.whisperLanguage.disabled =
      isBusy ||
      (getSourceMode() !== "whisper_sequence" && getSourceMode() !== "whisperx_sequence" && getSourceMode() !== "corrected_align");
  }
  if (elements.preserveTranslationTiming) {
    elements.preserveTranslationTiming.disabled =
      isBusy || getSourceMode() !== "whisper_sequence" || getSelectedWhisperLanguageCode().toLowerCase() === "auto";
  }
  if (elements.preserveMixedLanguages) {
    elements.preserveMixedLanguages.disabled =
      isBusy ||
      (getSourceMode() !== "whisper_sequence" && getSourceMode() !== "whisperx_sequence");
  }
  if (elements.mixedLanguagePrompt) {
    elements.mixedLanguagePrompt.disabled =
      isBusy ||
      !elements.preserveMixedLanguages?.checked ||
      (getSourceMode() !== "whisper_sequence" && getSourceMode() !== "whisperx_sequence");
  }
  if (elements.animationMode) {
    elements.animationMode.disabled = isBusy || isNativeSubtitleOutputMode();
  }
  if (elements.removePunctuation) {
    elements.removePunctuation.disabled = isBusy;
  }
  if (elements.maxChars) {
    elements.maxChars.disabled = isBusy;
  }
  if (elements.maxWords) {
    elements.maxWords.disabled = isBusy;
  }
  if (elements.linesPerCaption) {
    elements.linesPerCaption.disabled = isBusy;
  }
  if (elements.mogrtAspectFilter) {
    elements.mogrtAspectFilter.disabled = isBusy || isNativeSubtitleOutputMode();
  }
  if (elements.mogrtFolderButton) {
    elements.mogrtFolderButton.disabled = isBusy || isNativeSubtitleOutputMode();
  }
  if (elements.mogrtRefreshButton) {
    elements.mogrtRefreshButton.disabled = isBusy || isNativeSubtitleOutputMode();
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

function setVisualApplyProgressState(visible: boolean, done = 0, total = 0, label = ""): void {
  // // Render visual apply progress feedback for both queued startup delay and multi-MOGRT updates.
  if (!elements.visualApplyProgress || !elements.visualApplyProgressBar || !elements.visualApplyProgressText) {
    return;
  }

  if (!visible) {
    elements.visualApplyProgress.hidden = true;
    elements.visualApplyProgressBar.max = 1;
    elements.visualApplyProgressBar.value = 0;
    elements.visualApplyProgressText.textContent = "0 / 0";
    elements.visualApplyProgressBar.classList.remove("is-indeterminate");
    return;
  }

  if (total < 1) {
    elements.visualApplyProgress.hidden = false;
    elements.visualApplyProgressBar.max = 1;
    elements.visualApplyProgressBar.removeAttribute("value");
    elements.visualApplyProgressBar.classList.add("is-indeterminate");
    elements.visualApplyProgressText.textContent = label || translate("progress.visualApplyPending");
    return;
  }

  const clampedDone = Math.max(0, Math.min(total, done));
  const remaining = Math.max(0, total - clampedDone);
  elements.visualApplyProgress.hidden = false;
  elements.visualApplyProgressBar.max = total;
  elements.visualApplyProgressBar.value = clampedDone;
  elements.visualApplyProgressBar.classList.remove("is-indeterminate");
  elements.visualApplyProgressText.textContent = label || translateTemplate("visual.applyProgress", {
    done: String(clampedDone),
    total: String(total),
    remaining: String(remaining)
  });
}

function setTextApplyProgressState(visible: boolean, done = 0, total = 0, label = ""): void {
  // // Keep text-editor apply visibly busy even before the host rebuild starts returning progress-like milestones.
  if (!elements.textApplyProgress || !elements.textApplyProgressBar || !elements.textApplyProgressText) {
    return;
  }

  if (!visible) {
    elements.textApplyProgress.hidden = true;
    elements.textApplyProgressBar.max = 1;
    elements.textApplyProgressBar.value = 0;
    elements.textApplyProgressText.textContent = translate("progress.textApplyPending");
    elements.textApplyProgressBar.classList.remove("is-indeterminate");
    return;
  }

  if (total < 1) {
    elements.textApplyProgress.hidden = false;
    elements.textApplyProgressBar.max = 1;
    elements.textApplyProgressBar.removeAttribute("value");
    elements.textApplyProgressBar.classList.add("is-indeterminate");
    elements.textApplyProgressText.textContent = label || translate("progress.textApplyPending");
    return;
  }

  const clampedDone = Math.max(0, Math.min(total, done));
  elements.textApplyProgress.hidden = false;
  elements.textApplyProgressBar.max = Math.max(1, total);
  elements.textApplyProgressBar.value = clampedDone;
  elements.textApplyProgressBar.classList.remove("is-indeterminate");
  elements.textApplyProgressText.textContent =
    label ||
    translateTemplate("text.applyProgress", {
      done: String(clampedDone),
      total: String(total)
    });
}

function nextTextEditorBlockId(): string {
  // // Generate stable-enough UI ids for subtitle blocks inside one panel session.
  textEditorBlockIdCounter += 1;
  return `text-block-${textEditorBlockIdCounter}`;
}

function mapTextEditorBlocksToState(blocks: TextEditorBlock[]): TextEditorBlockState[] {
  // // Rebuild UI state from pure text-editor helper output after one editing operation.
  return blocks.map((block) => ({
    ...block,
    editorId: nextTextEditorBlockId(),
    selectedWordIndex: -1
  }));
}

function buildTextEditorBlocksFromState(blocks: TextEditorBlockState[]): TextEditorBlock[] {
  // // Convert panel block state back to pure text-editor blocks before diff/retime/apply helpers run.
  return blocks.map((block) => ({
    sourceSelectionIndex: block.sourceSelectionIndex,
    clipName: block.clipName,
    startSeconds: block.startSeconds,
    endSeconds: block.endSeconds,
    text: block.text,
    words: block.words.slice(),
    timedWords: block.timedWords ? block.timedWords.map((word) => ({ ...word })) : undefined
  }));
}

function getTextEditorSelectionTimingRange(): TextEditorTimingRange | undefined {
  // // Keep one stable selection-wide timing range so merges still fill the full original subtitle span.
  if (!(textEditorSelectionEndSeconds > textEditorSelectionStartSeconds)) {
    return undefined;
  }
  return {
    startSeconds: textEditorSelectionStartSeconds,
    endSeconds: textEditorSelectionEndSeconds
  };
}

function resolveCaptionMetadataIdentityFromHostPayload(payload: {
  projectDocumentId?: string;
  projectPath?: string;
  sequenceID?: string;
  sequenceName?: string;
}): CaptionMetadataIdentity | null {
  // // Normalize host identity payloads before using them as metadata-store keys.
  const projectDocumentId = String(payload.projectDocumentId || "").trim();
  const sequenceID = String(payload.sequenceID || "").trim();
  if (!projectDocumentId || !sequenceID) {
    return null;
  }

  return {
    projectDocumentId,
    projectPath: String(payload.projectPath || "").trim(),
    sequenceID,
    sequenceName: String(payload.sequenceName || "").trim()
  };
}

function setTextSelectionSummary(message: string): void {
  // // Centralize Text tab status feedback like selection count and unsupported multi-track warnings.
  if (!elements.textSelectionSummary) {
    return;
  }
  elements.textSelectionSummary.textContent = message;
}

function refreshTextButtonsBusyState(): void {
  // // Prevent concurrent read/apply actions while the Text tab is loading or rebuilding subtitle clips.
  const isBusy = textReadInProgress || textApplyInProgress;
  if (elements.textReadButton) {
    elements.textReadButton.disabled = isBusy;
  }
  if (elements.textApplyButton) {
    elements.textApplyButton.disabled = isBusy;
  }
}

function clearPendingTextEditorBlockCommit(editorId: string): void {
  // // Cancel a queued textarea normalization commit when a later input or action supersedes it.
  const pendingTimer = textEditorPendingCommitTimers.get(editorId);
  if (typeof pendingTimer === "number") {
    window.clearTimeout(pendingTimer);
    textEditorPendingCommitTimers.delete(editorId);
  }
}

function clearAllPendingTextEditorBlockCommits(): void {
  // // Drop stale deferred commits before merge/split/apply actions mutate the block list.
  Array.from(textEditorPendingCommitTimers.keys()).forEach((editorId) => {
    clearPendingTextEditorBlockCommit(editorId);
  });
}

function scheduleTextEditorBlockCommit(editorId: string, nextText: string): void {
  // // Defer blur/change normalization one tick so action-button clicks are not swallowed by an eager rerender on Windows.
  clearPendingTextEditorBlockCommit(editorId);
  const timer = window.setTimeout(() => {
    textEditorPendingCommitTimers.delete(editorId);
    const blockIndex = textEditorBlocks.findIndex((block) => block.editorId === editorId);
    if (blockIndex < 0) {
      return;
    }
    commitTextEditorBlockInput(blockIndex, nextText);
  }, 0);
  textEditorPendingCommitTimers.set(editorId, timer);
}

function selectTextEditorWord(blockIndex: number, wordIndex: number): void {
  // // Track which word the user picked so split actions can cut before that exact chip.
  textEditorBlocks = textEditorBlocks.map((block, index) => ({
    ...block,
    selectedWordIndex: index === blockIndex && block.selectedWordIndex === wordIndex ? -1 : index === blockIndex ? wordIndex : -1
  }));
  renderTextEditor();
}

function applyTextEditorBlocks(blocks: TextEditorBlock[]): void {
  // // Commit one pure text-edit result back into UI state and re-render the Text tab.
  clearAllPendingTextEditorBlockCommits();
  textEditorBlocks = mapTextEditorBlocksToState(retimeTextEditorBlocks(blocks, getTextEditorSelectionTimingRange()));
  renderTextEditor();
}

function updateTextEditorBlockInput(blockIndex: number, nextText: string): void {
  // // Keep Text tab state synchronized with freeform text edits without hitting Premiere yet.
  textEditorBlocks = textEditorBlocks.map((block, index) => {
    if (index !== blockIndex) {
      return block;
    }
    const words = tokenizeSubtitleText(nextText);
    const preserveTimedWords =
      Array.isArray(block.timedWords) &&
      block.timedWords.length === words.length &&
      block.timedWords.every((word, wordIndex) => String(word.text || "").trim() === String(words[wordIndex] || "").trim());
    return {
      ...block,
      text: String(nextText || ""),
      words,
      timedWords: preserveTimedWords ? block.timedWords?.map((word) => ({ ...word })) : undefined,
      selectedWordIndex:
        block.selectedWordIndex >= 0 && block.selectedWordIndex < words.length ? block.selectedWordIndex : -1
    };
  });
}

function commitTextEditorBlockInput(blockIndex: number, nextText: string): void {
  // // Normalize one edited subtitle row after blur/change and refresh its chip list/timing display.
  clearPendingTextEditorBlockCommit(textEditorBlocks[blockIndex]?.editorId || "");
  const updatedBlocks = updateTextEditorBlockText(buildTextEditorBlocksFromState(textEditorBlocks), blockIndex, nextText);
  textEditorBlocks = mapTextEditorBlocksToState(updatedBlocks);
  renderTextEditor();
}

function parseTextEditorDragPayload(event: DragEvent): { sourceBlockIndex: number; sourceWordIndex: number } | null {
  // // Decode one dragged word origin so drops can move chips across subtitle blocks.
  const rawPayload = event.dataTransfer?.getData("application/x-subcreator-word");
  if (!rawPayload) {
    return null;
  }
  try {
    const parsed = JSON.parse(rawPayload) as { sourceBlockIndex?: unknown; sourceWordIndex?: unknown };
    const sourceBlockIndex = Number(parsed.sourceBlockIndex);
    const sourceWordIndex = Number(parsed.sourceWordIndex);
    if (!Number.isFinite(sourceBlockIndex) || !Number.isFinite(sourceWordIndex)) {
      return null;
    }
    return {
      sourceBlockIndex,
      sourceWordIndex
    };
  } catch {
    return null;
  }
}

function moveTextEditorWordByDrop(targetBlockIndex: number, targetWordIndex?: number): (event: DragEvent) => void {
  // // Build one drop handler that inserts the dragged word before one chip or appends it to the target block.
  return (event: DragEvent) => {
    event.preventDefault();
    const payload = parseTextEditorDragPayload(event);
    if (!payload) {
      return;
    }
    const updatedBlocks = moveTextEditorWord(
      buildTextEditorBlocksFromState(textEditorBlocks),
      payload.sourceBlockIndex,
      payload.sourceWordIndex,
      targetBlockIndex,
      targetWordIndex
    );
    applyTextEditorBlocks(updatedBlocks);
  };
}

function bindTextEditorDropTarget(node: HTMLElement, targetBlockIndex: number, targetWordIndex?: number): void {
  // // Attach drag-over/drop handlers to chips and block containers for intuitive word moves.
  node.addEventListener("dragover", (event) => {
    event.preventDefault();
    node.classList.add("text-chip--drop-target");
  });
  node.addEventListener("dragleave", () => {
    node.classList.remove("text-chip--drop-target");
  });
  node.addEventListener("drop", (event) => {
    node.classList.remove("text-chip--drop-target");
    moveTextEditorWordByDrop(targetBlockIndex, targetWordIndex)(event);
  });
}

function setTextEditorActionPreview(targetBlockIndexes: number[]): void {
  // // Highlight the subtitle blocks affected by one merge/split action so users can preview the operation safely.
  if (!elements.textEditorList) {
    return;
  }

  const highlightedIndexes = new Set(targetBlockIndexes);
  const cards = elements.textEditorList.querySelectorAll<HTMLElement>(".text-block");
  cards.forEach((card) => {
    const blockIndex = Number(card.dataset.textBlockIndex || "-1");
    card.classList.toggle("is-action-target", highlightedIndexes.has(blockIndex));
  });
}

function bindTextEditorActionPreview(button: HTMLButtonElement, targetBlockIndexes: number[]): void {
  // // Mirror action hover/focus states onto the affected subtitle cards for clearer merge/split intent.
  if (button.disabled) {
    return;
  }

  button.addEventListener("mouseenter", () => {
    setTextEditorActionPreview(targetBlockIndexes);
  });
  button.addEventListener("mouseleave", () => {
    setTextEditorActionPreview([]);
  });
  button.addEventListener("focus", () => {
    setTextEditorActionPreview(targetBlockIndexes);
  });
  button.addEventListener("blur", () => {
    setTextEditorActionPreview([]);
  });
}

function renderTextEditor(): void {
  // // Render subtitle blocks, editable text, and draggable word chips for the Text tab.
  if (!elements.textEditorList) {
    return;
  }

  elements.textEditorList.innerHTML = "";
  if (textEditorBlocks.length < 1) {
    elements.textEditorList.textContent = translate("text.emptyState");
    return;
  }

  textEditorBlocks.forEach((block, blockIndex) => {
    const card = document.createElement("article");
    card.className = "text-block";
    card.dataset.textBlockIndex = String(blockIndex);

    const textarea = document.createElement("textarea");
    textarea.className = "text-block__textarea";
    textarea.value = block.text;
    textarea.placeholder = translate("text.textPlaceholder");
    textarea.spellcheck = true;
    textarea.lang = resolveSpellcheckLanguageCode();
    textarea.disabled = textApplyInProgress;
    textarea.addEventListener("input", () => {
      clearPendingTextEditorBlockCommit(block.editorId);
      updateTextEditorBlockInput(blockIndex, textarea.value);
    });
    textarea.addEventListener("change", () => {
      scheduleTextEditorBlockCommit(block.editorId, textarea.value);
    });
    textarea.addEventListener("blur", () => {
      scheduleTextEditorBlockCommit(block.editorId, textarea.value);
    });

    const chips = document.createElement("div");
    chips.className = "text-chip-list";
    bindTextEditorDropTarget(chips, blockIndex);

    block.words.forEach((word, wordIndex) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "text-chip";
      if (block.selectedWordIndex === wordIndex) {
        chip.classList.add("is-selected");
      }
      chip.textContent = word;
      chip.draggable = true;
      chip.disabled = textApplyInProgress;
      chip.addEventListener("click", () => {
        clearAllPendingTextEditorBlockCommits();
        selectTextEditorWord(blockIndex, wordIndex);
      });
      chip.addEventListener("dragstart", (event) => {
        chip.classList.add("is-dragging");
        event.dataTransfer?.setData(
          "application/x-subcreator-word",
          JSON.stringify({
            sourceBlockIndex: blockIndex,
            sourceWordIndex: wordIndex
          })
        );
        if (event.dataTransfer) {
          event.dataTransfer.effectAllowed = "move";
        }
      });
      chip.addEventListener("dragend", () => {
        chip.classList.remove("is-dragging");
      });
      bindTextEditorDropTarget(chip, blockIndex, wordIndex);
      chips.appendChild(chip);
    });

    const actions = document.createElement("div");
    actions.className = "text-block__actions";

    const mergePreviousButton = document.createElement("button");
    mergePreviousButton.type = "button";
    mergePreviousButton.className = "button button--secondary";
    mergePreviousButton.textContent = translate("action.mergePrevious");
    mergePreviousButton.disabled = textApplyInProgress || blockIndex < 1;
    if (blockIndex < 1) {
      // // Keep the three-button layout aligned while hiding impossible edge merges.
      mergePreviousButton.classList.add("is-placeholder");
      mergePreviousButton.setAttribute("aria-hidden", "true");
      mergePreviousButton.tabIndex = -1;
    }
    bindTextEditorActionPreview(mergePreviousButton, [blockIndex - 1, blockIndex]);
    mergePreviousButton.addEventListener("click", () => {
      clearAllPendingTextEditorBlockCommits();
      applyTextEditorBlocks(
        mergeTextEditorBlocks(
          buildTextEditorBlocksFromState(textEditorBlocks),
          blockIndex,
          "previous"
        )
      );
    });

    const splitButton = document.createElement("button");
    splitButton.type = "button";
    splitButton.className = "button button--secondary";
    splitButton.textContent = translate("action.splitSelectedWord");
    splitButton.disabled = textApplyInProgress || block.selectedWordIndex <= 0 || block.selectedWordIndex >= block.words.length;
    bindTextEditorActionPreview(splitButton, [blockIndex]);
    splitButton.addEventListener("click", () => {
      clearAllPendingTextEditorBlockCommits();
      if (block.selectedWordIndex <= 0) {
        return;
      }
      applyTextEditorBlocks(
        splitTextEditorBlock(
          buildTextEditorBlocksFromState(textEditorBlocks),
          blockIndex,
          block.selectedWordIndex
        )
      );
    });

    const mergeNextButton = document.createElement("button");
    mergeNextButton.type = "button";
    mergeNextButton.className = "button button--secondary";
    mergeNextButton.textContent = translate("action.mergeNext");
    mergeNextButton.disabled = textApplyInProgress || blockIndex >= textEditorBlocks.length - 1;
    if (blockIndex >= textEditorBlocks.length - 1) {
      // // Keep the three-button layout aligned while hiding impossible edge merges.
      mergeNextButton.classList.add("is-placeholder");
      mergeNextButton.setAttribute("aria-hidden", "true");
      mergeNextButton.tabIndex = -1;
    }
    bindTextEditorActionPreview(mergeNextButton, [blockIndex, blockIndex + 1]);
    mergeNextButton.addEventListener("click", () => {
      clearAllPendingTextEditorBlockCommits();
      applyTextEditorBlocks(
        mergeTextEditorBlocks(
          buildTextEditorBlocksFromState(textEditorBlocks),
          blockIndex,
          "next"
        )
      );
    });

    actions.append(mergePreviousButton, splitButton, mergeNextButton);
    card.append(textarea, chips, actions);
    elements.textEditorList?.appendChild(card);
  });
}

function isTextEditorPathLikeValue(text: string): boolean {
  // // Ignore media-path strings that some MOGRT controls expose so the Text tab only shows real subtitle text blocks.
  const normalized = String(text || "").trim();
  if (!normalized) {
    return false;
  }

  const looksAbsolutePath =
    normalized.startsWith("/Volumes/") ||
    normalized.startsWith("/") ||
    normalized.startsWith("~/") ||
    /^[a-z]:[\\/]/i.test(normalized) ||
    /^\\\\/.test(normalized);
  const looksMediaFile = /\.(mp4|mov|mxf|mp3|wav|aif|aiff|m4a|avi|mkv|jpg|jpeg|png|webp|psd|mogrt)$/i.test(normalized);
  return looksAbsolutePath && looksMediaFile;
}

function buildTextEditorTextFromTimedWords(words: CaptionWord[] | null | undefined): string {
  // // Rebuild the visible subtitle text from persisted word timings when Premiere only returns one opaque placeholder glyph.
  if (!Array.isArray(words) || words.length < 1) {
    return "";
  }

  return tokenizeSubtitleText(
    words
      .map((word) => String(word?.text || "").trim())
      .filter(Boolean)
      .join(" ")
  ).join(" ");
}

function cleanupTemporaryMogrtPaths(paths: string[]): void {
  // // Remove temporary baked Premiere `.mogrt` files after the host has imported them into the sequence.
  const modules = window.cep_node?.require?.("fs") as { unlinkSync?: (path: string) => void } | undefined;
  if (!modules?.unlinkSync) {
    return;
  }

  for (const cleanupPath of paths) {
    try {
      modules.unlinkSync(cleanupPath);
    } catch {
      // // Ignore cleanup failures because subtitle rebuild already completed.
    }
  }
}

async function loadTextItemsFromSelection(emitHostLog = false, showLoadingState = false): Promise<void> {
  // // Read selected subtitle MOGRTs into the Text tab so users can edit text blocks safely.
  if (showLoadingState) {
    if (textReadInProgress || textApplyInProgress) {
      return;
    }
    textReadInProgress = true;
    refreshTextButtonsBusyState();
    setTextApplyProgressState(true, 0, 0, translate("progress.textReadPending"));
    await waitForNextPaint();
  }

  try {
    clearAllPendingTextEditorBlockCommits();
    const result = await readSelectedMogrtTextItems();
    const filteredSelectionItems = result.items.filter((item) => !isTextEditorPathLikeValue(String(item.text || "").trim()));
    textEditorSelectionMetadataIdentity = resolveCaptionMetadataIdentityFromHostPayload(result);
    textEditorSelectionSignature = result.signature;
    textEditorSameTrack = result.sameTrack !== false;
    textEditorVideoTrackIndex = Number.isFinite(Number(result.videoTrackIndex)) ? Number(result.videoTrackIndex) : -1;
    const metadataWordsByItem = resolveCaptionMetadataForSelection(textEditorSelectionMetadataIdentity, filteredSelectionItems);
    textEditorOriginalBlocks = filteredSelectionItems.map((item, itemIndex) => {
      const resolvedTimedWords = Array.isArray(metadataWordsByItem[itemIndex]) ? metadataWordsByItem[itemIndex] || undefined : undefined;
      const resolvedText = buildTextEditorTextFromTimedWords(resolvedTimedWords) || String(item.text || "").trim();

      return {
        sourceSelectionIndex: Number(item.selectionIndex || 0),
        clipName: String(item.clipName || "").trim(),
        startSeconds: Number(item.startSeconds || 0),
        endSeconds: Number(item.endSeconds || 0),
        text: resolvedText,
        words: tokenizeSubtitleText(resolvedText),
        timedWords: resolvedTimedWords
      };
    });
    textEditorBlocks = mapTextEditorBlocksToState(textEditorOriginalBlocks);
    textEditorSelectionStartSeconds =
      filteredSelectionItems.length > 0
        ? filteredSelectionItems.reduce(
            (lowestValue, item) => Math.min(lowestValue, Number(item.startSeconds || 0)),
            Number.POSITIVE_INFINITY
          )
        : 0;
    textEditorSelectionEndSeconds =
      filteredSelectionItems.length > 0
        ? filteredSelectionItems.reduce(
            (highestValue, item) => Math.max(highestValue, Number(item.endSeconds || 0)),
            Number.NEGATIVE_INFINITY
          )
        : 0;
    renderTextEditor();

    if (emitHostLog) {
      setStructuredLog(translate("log.hostResult"), result);
    }

    if (filteredSelectionItems.length < 1) {
      setTextSelectionSummary(translate("text.selectionDefault"));
      return;
    }

    if (!textEditorSameTrack) {
      setTextSelectionSummary(
        translateTemplate("text.selectionMixedTracks", {
          clips: String(filteredSelectionItems.length)
        })
      );
      return;
    }

    setTextSelectionSummary(
      translateTemplate("text.selectionSummary", {
        clips: String(filteredSelectionItems.length),
        track: String(Math.max(0, textEditorVideoTrackIndex) + 1)
      })
    );
  } finally {
    if (showLoadingState) {
      textReadInProgress = false;
      refreshTextButtonsBusyState();
      setTextApplyProgressState(false);
    }
  }
}

function setTranslationSelectionSummary(message: string): void {
  // // Keep the dedicated translation tab clear about which Premiere clips will be duplicated.
  if (elements.translationSelectionSummary) {
    elements.translationSelectionSummary.textContent = message;
  }
}

function renderTranslationPreview(): void {
  // // Render one editable field per cue so manual corrections keep their original subtitle boundaries.
  if (!elements.translationPreview) {
    return;
  }
  elements.translationPreview.innerHTML = "";
  if (translatedSubtitleTexts.length < 1) {
    elements.translationPreview.textContent = translate("placeholder.translationPreview");
    return;
  }
  translatedSubtitleTexts.forEach((text, index) => {
    const item = document.createElement("label");
    item.className = "translation-preview__item";
    const label = document.createElement("span");
    label.textContent = `${index + 1}. ${translationBlocks[index]?.startSeconds.toFixed(2) ?? "0.00"}`;
    const input = document.createElement("textarea");
    input.value = text;
    input.rows = 2;
    input.addEventListener("input", () => {
      // // Store each correction directly by cue index, without recalculating any timeline time.
      translatedSubtitleTexts[index] = input.value;
    });
    item.append(label, input);
    elements.translationPreview?.appendChild(item);
  });
}

function getTranslationInputMode(): "mogrt" | "srt" {
  // // Keep translation sources explicit because Premiere cannot reliably expose native caption text through CEP.
  return elements.translationInputMode?.value === "srt" ? "srt" : "mogrt";
}

function toggleTranslationInputMode(): void {
  // // Show the appropriate source picker while keeping the translated preview untouched until it is reloaded.
  const srtMode = getTranslationInputMode() === "srt";
  if (elements.translationSrtField) {
    elements.translationSrtField.hidden = !srtMode;
  }
  if (elements.translationReadButton) {
    elements.translationReadButton.textContent = translate(srtMode ? "action.translationReadSrt" : "action.translationRead");
  }
}

function replaceTranslationLanguageOptions(
  select: HTMLSelectElement | null,
  languages: Array<{ language: string; name: string }>,
  includeAutoDetect: boolean,
  fallbackValue: string
): void {
  // // Rebuild the select from DeepL's plan-specific response while retaining a valid previous choice when possible.
  if (!select) {
    return;
  }
  const selectedValue = String(select.value || "").trim().toUpperCase();
  select.replaceChildren();
  if (includeAutoDetect) {
    const autoOption = document.createElement("option");
    autoOption.value = "AUTO";
    autoOption.textContent = translate("translation.languageAuto");
    select.appendChild(autoOption);
  }
  languages.forEach((language) => {
    const option = document.createElement("option");
    option.value = language.language;
    option.textContent = language.name;
    select.appendChild(option);
  });
  const availableValues = Array.from(select.options).map((option) => option.value.toUpperCase());
  const nextValue = availableValues.includes(selectedValue)
    ? selectedValue
    : availableValues.includes(fallbackValue.toUpperCase())
      ? fallbackValue
      : select.options[0]?.value || "";
  select.value = nextValue;
}

async function refreshDeepLSupportedLanguages(): Promise<void> {
  // // Fetch the exact DeepL source/target language lists for the stored personal API key.
  const authKey = String(elements.deeplApiKey?.value || "").trim();
  if (!authKey || translationLanguagesLoadInProgress) {
    return;
  }
  translationLanguagesLoadInProgress = true;
  if (elements.translationLanguagesRefreshButton) {
    elements.translationLanguagesRefreshButton.disabled = true;
  }
  try {
    const languages = await getDeepLSupportedLanguages(authKey);
    replaceTranslationLanguageOptions(elements.translationSourceLanguage, languages.source, true, "AUTO");
    replaceTranslationLanguageOptions(elements.translationTargetLanguage, languages.target, false, "FR");
    setLog(translate("log.deeplLanguagesLoaded"));
  } finally {
    translationLanguagesLoadInProgress = false;
    if (elements.translationLanguagesRefreshButton) {
      elements.translationLanguagesRefreshButton.disabled = false;
    }
  }
}

function getDeepLUserErrorMessage(error: unknown): string {
  // // Present common DeepL failures in the panel language instead of leaking HTTP or network implementation details.
  const code = String(error || "");
  if (code.includes("SUBCREATOR_DEEPL_API_KEY_INVALID")) {
    return translate("error.deeplApiKeyInvalid");
  }
  if (code.includes("SUBCREATOR_DEEPL_QUOTA_EXCEEDED")) {
    return translate("error.deeplQuotaExceeded");
  }
  if (code.includes("SUBCREATOR_DEEPL_RATE_LIMITED")) {
    return translate("error.deeplRateLimited");
  }
  if (code.includes("SUBCREATOR_DEEPL_NETWORK_UNAVAILABLE")) {
    return translate("error.deeplNetworkUnavailable");
  }
  if (code.includes("SUBCREATOR_DEEPL_RESPONSE_INVALID") || code.includes("SUBCREATOR_DEEPL_REQUEST_FAILED")) {
    return translate("error.deeplRequestFailed");
  }
  return code;
}

async function loadTranslationSelection(): Promise<void> {
  // // Read selected Sub Creator MOGRTs without altering the Text editor's active selection state.
  if (getTranslationInputMode() === "srt") {
    const srtPath = String(elements.translationSrtPath?.value || "").trim();
    if (!srtPath) {
      throw new Error(translate("error.translationSrtMissing"));
    }
    const cues = parseSrt(await readTextFileFromHost(srtPath));
    translationBlocks = cues.map((cue, index) => ({
      sourceSelectionIndex: index,
      clipName: `Caption ${index + 1}`,
      startSeconds: cue.startSeconds,
      endSeconds: cue.endSeconds,
      text: cue.text,
      words: cue.words.map((word) => word.text),
      timedWords: cue.words
    }));
    translationSelectionSignature = `srt:${srtPath}`;
    translationSameTrack = true;
    translatedSubtitleTexts = [];
    renderTranslationPreview();
    if (elements.translationDuplicateButton) {
      elements.translationDuplicateButton.disabled = true;
    }
    setTranslationSelectionSummary(
      translationBlocks.length > 0
        ? translateTemplate("translation.selectionReady", { count: String(translationBlocks.length) })
        : translate("translation.selectionDefault")
    );
    return;
  }
  const result = await readSelectedMogrtTextItems();
  const selectedItems = result.items.filter((item) => !isTextEditorPathLikeValue(String(item.text || "").trim()));
  const identity = resolveCaptionMetadataIdentityFromHostPayload(result);
  const metadataWordsByItem = resolveCaptionMetadataForSelection(identity, selectedItems);
  translationBlocks = selectedItems.map((item, index) => {
    const timedWords = Array.isArray(metadataWordsByItem[index]) ? metadataWordsByItem[index] : undefined;
    const text = buildTextEditorTextFromTimedWords(timedWords) || String(item.text || "").trim();
    return {
      sourceSelectionIndex: Number(item.selectionIndex || 0),
      clipName: String(item.clipName || "").trim(),
      startSeconds: Number(item.startSeconds || 0),
      endSeconds: Number(item.endSeconds || 0),
      text,
      words: tokenizeSubtitleText(text),
      timedWords
    };
  }).filter((block) => Boolean(block.text));
  translationSelectionSignature = String(result.signature || "").trim();
  translationSameTrack = result.sameTrack !== false;
  translatedSubtitleTexts = [];
  renderTranslationPreview();
  if (elements.translationDuplicateButton) {
    elements.translationDuplicateButton.disabled = true;
  }
  if (translationBlocks.length < 1) {
    setTranslationSelectionSummary(translate("translation.selectionDefault"));
    return;
  }
  if (!translationSameTrack) {
    setTranslationSelectionSummary(translate("translation.selectionMixedTracks"));
    return;
  }
  setTranslationSelectionSummary(translateTemplate("translation.selectionReady", { count: String(translationBlocks.length) }));
}

async function prepareGeneratedNativeSrtForTranslation(srtPath: string): Promise<void> {
  // // Reuse the exact SRT written before Premiere imports the native track, avoiding unreliable caption-text reads from CEP.
  if (!elements.translationAutoLoadGeneratedNativeSrt?.checked || !elements.translationInputMode || !elements.translationSrtPath) {
    return;
  }

  const normalizedPath = String(srtPath || "").trim();
  if (!normalizedPath) {
    return;
  }

  elements.translationInputMode.value = "srt";
  elements.translationSrtPath.value = normalizedPath;
  toggleTranslationInputMode();
  await loadTranslationSelection();
  persistPanelState();
  setLog(translate("log.translationGeneratedNativeSrtReady"));
}

async function translateLoadedSelection(): Promise<void> {
  // // Translate selected text through the user's DeepL key while retaining the source cue boundaries locally.
  if (!translationSelectionSignature || translationBlocks.length < 1) {
    throw new Error(translate("error.translationSelectionMissing"));
  }
  if (!translationSameTrack) {
    throw new Error(translate("error.translationMixedTracks"));
  }
  const authKey = String(elements.deeplApiKey?.value || "").trim();
  if (!authKey) {
    throw new Error(translate("error.deeplApiKeyMissing"));
  }
  const sourceLanguage = String(elements.translationSourceLanguage?.value || "AUTO").trim();
  const targetLanguage = String(elements.translationTargetLanguage?.value || "FR").trim();
  const context = translationBlocks.map((block) => block.text).join(" ");
  setTranslationSelectionSummary(translate("translation.translating"));
  translatedSubtitleTexts = await translateWithDeepL({
    authKey,
    texts: translationBlocks.map((block) => block.text),
    sourceLanguage,
    targetLanguage,
    context
  });
  renderTranslationPreview();
  if (elements.translationDuplicateButton) {
    elements.translationDuplicateButton.disabled = false;
  }
  setTranslationSelectionSummary(translateTemplate("translation.translatedReady", { count: String(translatedSubtitleTexts.length) }));
}

async function duplicateTranslatedSelection(): Promise<void> {
  // // Rebuild translated captions on a safe track above the source while preserving each original timing range.
  if (!translationSelectionSignature || translatedSubtitleTexts.length !== translationBlocks.length) {
    throw new Error(translate("error.translationRequired"));
  }
  if (translatedSubtitleTexts.some((text) => !String(text || "").trim())) {
    throw new Error(translate("error.translationRequired"));
  }
  if (getTranslationInputMode() === "srt") {
    // // Reuse the existing native-caption importer to add a second Premiere subtitle track from translated SRT-equivalent cues.
    const rawResult = await applyNativeSubtitlePlan({
      options: { outputMode: "premiere_subtitles" },
      cues: translationBlocks.map((block, index): CaptionCue => ({
        id: `translated-native-${index + 1}`,
        startSeconds: block.startSeconds,
        endSeconds: block.endSeconds,
        text: translatedSubtitleTexts[index] || "",
        words: []
      }))
    });
    setStructuredLogFromRaw(translate("log.hostResult"), rawResult);
    assertHostApplySucceeded(rawResult);
    setTranslationSelectionSummary(translateTemplate("translation.duplicateDone", { count: String(translationBlocks.length) }));
    return;
  }
  const options = collectTextEditorBuildOptions();
  const payload: TextEditorApplyPayload = {
    selectionSignature: translationSelectionSignature,
    replaceSelectionStartIndex: 0,
    replaceSelectionEndIndex: translationBlocks.length - 1,
    duplicateSelection: true,
    items: translationBlocks.map((block, index) => ({
      sourceSelectionIndex: block.sourceSelectionIndex,
      startSeconds: block.startSeconds,
      endSeconds: block.endSeconds,
      text: translatedSubtitleTexts[index] || ""
    })),
    options
  };
  setTranslationSelectionSummary(translate("translation.duplicating"));
  const result = await applySelectedMogrtTextItems(payload);
  if (Number(result.failedCount || 0) !== 0 || Number(result.rebuiltCount || 0) !== translationBlocks.length) {
    throw new Error(translate("error.translationDuplicateFailed"));
  }
  setStructuredLog(translate("log.hostResult"), result);
  setTranslationSelectionSummary(translateTemplate("translation.duplicateDone", { count: String(result.rebuiltCount || 0) }));
}

async function applyTextEditorChanges(): Promise<void> {
  // // Rebuild selected subtitle MOGRTs from edited Text tab blocks while preserving precise timing metadata when available.
  if (textReadInProgress || textApplyInProgress) {
    return;
  }
  if (!textEditorSelectionSignature) {
    throw new Error(translate("error.textSelectionMissing"));
  }
  if (!textEditorSameTrack) {
    throw new Error(translate("error.textMixedTracks"));
  }

  clearAllPendingTextEditorBlockCommits();
  const editableBlocks = buildTextEditorBlocksFromState(textEditorBlocks);
  // // Use one safe combined span for disjoint edits so the host does not partially rebuild the selection.
  const applyPlans = buildTextEditorSafeApplyPlans(textEditorOriginalBlocks, editableBlocks);
  if (applyPlans.length < 1) {
    setLog(translate("log.textNoChanges"));
    return;
  }

  const options = collectTextEditorBuildOptions();
  const premiereTemplateTextPayloads = await readPremiereTemplateTextPayloads(options.mogrtPath);

  textApplyInProgress = true;
  refreshTextButtonsBusyState();
  setTextApplyProgressState(true, 0, 0, translate("progress.textApplyPending"));
  try {
    const orderedPlans = applyPlans.slice().sort((left, right) => right.selectionStartIndex - left.selectionStartIndex);
    let currentSelectionSignature = textEditorSelectionSignature;
    let lastIdentity = textEditorSelectionMetadataIdentity;
    let lastSourceTrackIndex = textEditorVideoTrackIndex;
    let lastRebuildTrackIndex = textEditorVideoTrackIndex;
    const rangeResults: ApplySelectedMogrtTextResult[] = [];

    for (const applyPlan of orderedPlans) {
      let cleanupMogrtPaths: string[] = [];
      try {
        const plannedBlocks = prepareTextEditorBlocksForApply(applyPlan.blocks);
        if (plannedBlocks.length < 1) {
          throw new Error(translate("error.textEmptyBlocks"));
        }

        let applyItems = plannedBlocks.map((block) => ({
          sourceSelectionIndex: block.sourceSelectionIndex,
          startSeconds: block.startSeconds,
          endSeconds: block.endSeconds,
          text: block.text
        }));
        setTextApplyProgressState(
          true,
          orderedPlans.length - rangeResults.length - 1,
          orderedPlans.length,
          translate("progress.textApplyPending")
        );
        await waitForNextPaint();

        if (premiereTemplateTextPayloads.length > 0) {
          const preparedTemplateBuild = await buildPremiereTemplateCueMogrts(
            options.mogrtPath,
            plannedBlocks.map((block, blockIndex) => ({
              id: `text-apply-${blockIndex}`,
              startSeconds: block.startSeconds,
              endSeconds: block.endSeconds,
              text: block.text,
              words: []
            })),
            premiereTemplateTextPayloads
          );
          cleanupMogrtPaths = preparedTemplateBuild.cleanupPaths;
          applyItems = applyItems.map((item, itemIndex) => ({
            ...item,
            mogrtPathOverride: preparedTemplateBuild.cues[itemIndex]?.mogrtPathOverride,
            skipTextApply: preparedTemplateBuild.cues[itemIndex]?.skipTextApply
          }));
        }

        const payload: TextEditorApplyPayload = {
          selectionSignature: currentSelectionSignature,
          replaceSelectionStartIndex: applyPlan.selectionStartIndex,
          replaceSelectionEndIndex: applyPlan.selectionEndIndex,
          items: applyItems,
          options: {
            ...options,
            premiereTemplateTextPayloads
          }
        };

        const result: ApplySelectedMogrtTextResult = await applySelectedMogrtTextItems(payload);
        rangeResults.unshift(result);
        if (Number(result.failedCount || 0) !== 0 || Number(result.rebuiltCount || 0) !== plannedBlocks.length) {
          throw new Error(translate("error.textApplyPartialFailure"));
        }
        setTextApplyProgressState(true, orderedPlans.length - rangeResults.length, orderedPlans.length);

        currentSelectionSignature = String(result.selectionSignature || currentSelectionSignature || "").trim();
        lastIdentity = resolveCaptionMetadataIdentityFromHostPayload(result) || lastIdentity;
        lastSourceTrackIndex = Number.isFinite(Number(result.sourceTrackIndex))
          ? Number(result.sourceTrackIndex)
          : lastSourceTrackIndex;
        lastRebuildTrackIndex = Number.isFinite(Number(result.rebuildTrackIndex))
          ? Number(result.rebuildTrackIndex)
          : lastRebuildTrackIndex;
        persistTextEditorCaptionMetadata(
          lastIdentity,
          lastSourceTrackIndex,
          lastRebuildTrackIndex,
          applyPlan.timingRange,
          plannedBlocks
        );

        const refreshedSelection = await readSelectedMogrtTextItems();
        currentSelectionSignature = String(refreshedSelection.signature || currentSelectionSignature || "").trim();
        lastIdentity = resolveCaptionMetadataIdentityFromHostPayload(refreshedSelection) || lastIdentity;
        if (Number.isFinite(Number(refreshedSelection.videoTrackIndex))) {
          lastSourceTrackIndex = Number(refreshedSelection.videoTrackIndex);
          lastRebuildTrackIndex = Number(refreshedSelection.videoTrackIndex);
        }
      } finally {
        cleanupTemporaryMogrtPaths(cleanupMogrtPaths);
      }
    }

    setStructuredLog(translate("log.textApplyDone"), {
      rangeCount: rangeResults.length,
      rebuiltCount: rangeResults.reduce((total, result) => total + Number(result.rebuiltCount || 0), 0),
      failedCount: rangeResults.reduce((total, result) => total + Number(result.failedCount || 0), 0),
      ranges: rangeResults
    });
    await loadTextItemsFromSelection();
  } finally {
    textApplyInProgress = false;
    refreshTextButtonsBusyState();
    setTextApplyProgressState(false);
  }
}

function renderVisualPropertyEditor(properties: HostVisualProperty[]): void {
  // // Render editable controls from selected MOGRT property metadata returned by host.
  if (!elements.visualPropertyList) {
    return;
  }

  visualLoadedRevision += 1;
  captureOpenVisualGroupsFromDom();
  elements.visualPropertyList.innerHTML = "";
  loadedVisualProperties = [];
  visualOriginalValuesByPath.clear();
  visualTextStyleTokenMapByBasePath.clear();
  visualDirtyPaths.clear();

  if (!properties.length) {
    refreshVisualButtonsBusyState();
    renderVisualSelectionSummary();
    return;
  }

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

  const resolveStyleOptionsForFamily = (family: string, styleMap?: Record<string, string[]>): string[] => {
    // // Resolve host-exposed styles for one family without triggering any system-font scan.
    const mergedMap = normalizeStyleMap(styleMap || {});
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
      if (resolvedStyleOptions.length !== 1) {
        continue;
      }
      const syntheticStyleValue = resolvedStyleOptions[0];

      propertiesWithStyleControls.push({
        path: `${textStylePath.basePath}::textstyle.fontStyle`,
        displayName: "Font Style",
        groupPath: property.groupPath,
        valueType: "string",
        controlKind: "string",
        styleOptionsByFamily: currentFamily ? { [currentFamily]: resolvedStyleOptions.slice() } : {},
        value: syntheticStyleValue
      });
    }

    return propertiesWithStyleControls;
  };

  const renderProperties = ensureTextStyleControls(properties);
  loadedVisualProperties = renderProperties.slice();

  const bindLiveUpdateEvent = (
    control: HTMLElement | HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement,
    eventName: "input" | "change" = "change",
    dirtyPathOverride = ""
  ): void => {
    // // Track explicit user edits so single-clip apply only touches changed controls.
    control.addEventListener(eventName, () => {
      const dirtyPath =
        String(dirtyPathOverride || ((control as HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement).dataset?.visualPath || "")).trim();
      if (dirtyPath) {
        visualDirtyPaths.add(dirtyPath);
      }
      scheduleLiveVisualApply();
    });
  };

  const clipLevelSectionOrder = new Map([
    ["motion", 0],
    ["vector motion", 1],
    ["opacity", 2]
  ]);
  const clipRootVisualNode: VisualEditorSectionNode = {
    key: "clip-root",
    label: "Clip",
    properties: [],
    children: []
  };

  const splitVisualGroupPath = (value: string): string[] =>
    String(value || "")
      .split("/")
      .map((segment) => String(segment || "").trim())
      .filter((segment) => segment.length > 0);

  const componentNameByIndex = new Map<number, string>();
  for (const component of loadedVisualComponents) {
    const componentIndex = Number(component.index);
    if (!Number.isFinite(componentIndex)) {
      continue;
    }
    const componentName = String(component.name || "").trim();
    if (!componentName) {
      continue;
    }
    componentNameByIndex.set(componentIndex, componentName);
  }

  const resolveVisualComponentName = (property: HostVisualProperty): string => {
    // // Prefer host component labels so duplicated Premiere `Group`/`Text` components can stay distinct in the panel.
    const componentIndex = parseVisualComponentIndex(property.path);
    const knownComponentName = String(componentNameByIndex.get(componentIndex) || "").trim();
    if (knownComponentName) {
      return knownComponentName;
    }
    const groupSegments = splitVisualGroupPath(String(property.groupPath || ""));
    if (groupSegments.length > 0) {
      return groupSegments[0];
    }
    return `Component ${componentIndex + 1}`;
  };

  const isClipLevelComponent = (componentName: string): boolean =>
    clipLevelSectionOrder.has(String(componentName || "").trim().toLowerCase());

  const componentEntries = new Map<
    number,
    {
      componentIndex: number;
      componentName: string;
      encounterOrder: number;
      properties: HostVisualProperty[];
    }
  >();

  for (const property of renderProperties) {
    if (property.controlKind === "text" || property.controlKind === "json") {
      continue;
    }

    visualOriginalValuesByPath.set(
      property.path,
      canonicalizeVisualValue(property.controlKind, property.valueType, property.value)
    );

    const componentIndex = parseVisualComponentIndex(property.path);
    const existing = componentEntries.get(componentIndex);
    if (existing) {
      existing.properties.push(property);
      continue;
    }

    componentEntries.set(componentIndex, {
      componentIndex,
      componentName: resolveVisualComponentName(property),
      encounterOrder: componentEntries.size,
      properties: [property]
    });
  }

  const orderedComponentEntries = Array.from(componentEntries.values()).sort((left, right) => {
    if (left.componentIndex !== right.componentIndex) {
      return left.componentIndex - right.componentIndex;
    }
    return left.encounterOrder - right.encounterOrder;
  });

  const duplicateComponentCounts = new Map<string, number>();
  for (const componentEntry of orderedComponentEntries) {
    const componentKey = String(componentEntry.componentName || "").trim().toLowerCase();
    duplicateComponentCounts.set(componentKey, (duplicateComponentCounts.get(componentKey) || 0) + 1);
  }

  const duplicateComponentRemainders = new Map(duplicateComponentCounts);
  const componentLabelByIndex = new Map<number, string>();
  for (const componentEntry of orderedComponentEntries) {
    const componentKey = String(componentEntry.componentName || "").trim().toLowerCase();
    const duplicateCount = duplicateComponentCounts.get(componentKey) || 0;
    const duplicateOffset = duplicateComponentRemainders.get(componentKey) || duplicateCount;
    const baseLabel = componentEntry.componentName;
    const displayLabel =
      duplicateCount > 1 ? `${baseLabel} ${String(duplicateOffset).padStart(2, "0")}` : baseLabel;
    componentLabelByIndex.set(componentEntry.componentIndex, displayLabel);
    duplicateComponentRemainders.set(componentKey, Math.max(0, duplicateOffset - 1));
  }

  const buildVisualSectionTree = (
    properties: HostVisualProperty[],
    rootKey: string,
    rootLabel: string,
    rootGroupPathPrefix: string
  ): VisualEditorSectionNode => {
    // // Rebuild nested `Properties` sections from slash-separated group paths so Premiere-authored groups stay readable.
    const rootNode: VisualEditorMutableSectionNode = {
      key: rootKey,
      label: rootLabel,
      properties: [],
      children: [],
      encounterOrder: 0
    };
    const childNodesByParentKey = new Map<string, Map<string, VisualEditorMutableSectionNode>>();

    const ensureChildNode = (
      parentNode: VisualEditorMutableSectionNode,
      childLabel: string,
      childKeySegment: string
    ): VisualEditorMutableSectionNode => {
      // // Keep one stable child node per label path so repeated properties merge into the same subsection.
      const parentChildren = childNodesByParentKey.get(parentNode.key) || new Map<string, VisualEditorMutableSectionNode>();
      childNodesByParentKey.set(parentNode.key, parentChildren);
      const normalizedChildKey = String(childKeySegment || "").trim().toLowerCase() || "__section__";
      const existingChild = parentChildren.get(normalizedChildKey);
      if (existingChild) {
        return existingChild;
      }

      const createdChild: VisualEditorMutableSectionNode = {
        key: `${parentNode.key}:section:${normalizedChildKey}`,
        label: childLabel,
        properties: [],
        children: [],
        encounterOrder: parentNode.children.length
      };
      parentNode.children.push(createdChild);
      parentChildren.set(normalizedChildKey, createdChild);
      return createdChild;
    };

    for (const property of properties) {
      const groupSegments = splitVisualGroupPath(String(property.groupPath || ""));
      const relativeSegments =
        rootGroupPathPrefix &&
        groupSegments.length > 0 &&
        groupSegments[0].trim().toLowerCase() === rootGroupPathPrefix.toLowerCase()
          ? groupSegments.slice(1)
          : groupSegments.slice();

      if (relativeSegments.length < 1) {
        rootNode.properties.push(property);
        continue;
      }

      let targetNode = rootNode;
      for (const relativeSegment of relativeSegments) {
        targetNode = ensureChildNode(targetNode, relativeSegment, relativeSegment);
      }
      targetNode.properties.push(property);
    }

    const finalizeVisualSectionNode = (node: VisualEditorMutableSectionNode): VisualEditorSectionNode => {
      // // Strip editor-only ordering metadata once the tree is ready for render.
      const orderedChildren = node.children
        .slice()
        .sort((leftNode, rightNode) => leftNode.encounterOrder - rightNode.encounterOrder)
        .map((childNode) => finalizeVisualSectionNode(childNode));
      return {
        key: node.key,
        label: node.label,
        properties: node.properties.slice(),
        children: orderedChildren
      };
    };

    return finalizeVisualSectionNode(rootNode);
  };

  const buildVisualComponentNode = (
    componentEntry: {
      componentIndex: number;
      componentName: string;
      properties: HostVisualProperty[];
    },
    displayLabel: string
  ): VisualEditorSectionNode => {
    // // Drop the synthetic `Settings` wrapper and keep real nested section names only when Premiere exposes them.
    return buildVisualSectionTree(
      componentEntry.properties,
      `component:${componentEntry.componentIndex}`,
      displayLabel,
      String(componentEntry.componentName || "").trim()
    );
  };

  const normalizeVisualNodeForRender = (node: VisualEditorSectionNode): VisualEditorSectionNode => {
    // // Normalize nested Premiere sections recursively so inner `Shape -> Align and Transform` wrappers do not survive just because they sit under a Group.
    const normalizedChildren = node.children.map((childNode) => normalizeVisualNodeForRender(childNode));
    const nextNode: VisualEditorSectionNode = {
      key: node.key,
      label: node.label,
      properties: node.properties.slice(),
      children: normalizedChildren
    };

    if (nextNode.children.length !== 1) {
      return nextNode;
    }

    const onlyChild = nextNode.children[0];
    const normalizedChildLabel = String(onlyChild.label || "").trim().toLowerCase();
    if (normalizedChildLabel !== "settings" && normalizedChildLabel !== "align and transform") {
      return nextNode;
    }

    // // Fold one low-signal subsection directly into its parent so the panel stays closer to Premiere's visible structure.
    return {
      key: nextNode.key,
      label: nextNode.label,
      properties: nextNode.properties.concat(onlyChild.properties),
      children: onlyChild.children.slice()
    };
  };

  const topLevelVisualNodes: VisualEditorSectionNode[] = [];
  let activeGroupNode: VisualEditorSectionNode | null = null;

  for (const componentEntry of orderedComponentEntries) {
    const componentName = String(componentEntry.componentName || "").trim();
    const displayLabel = String(componentLabelByIndex.get(componentEntry.componentIndex) || componentName).trim() || "Component";
    const componentNode = buildVisualComponentNode(componentEntry, displayLabel);
    const normalizedComponentName = componentName.toLowerCase();

    if (isClipLevelComponent(componentName)) {
      clipRootVisualNode.children.push(componentNode);
      continue;
    }

    if (normalizedComponentName === "group") {
      topLevelVisualNodes.push(componentNode);
      activeGroupNode = componentNode;
      continue;
    }

    if (activeGroupNode) {
      activeGroupNode.children.push(componentNode);
      continue;
    }

    topLevelVisualNodes.push(componentNode);
  }

  if (clipRootVisualNode.children.length > 0) {
    // // Keep clip-level controls grouped together at the bottom so layer/effect sections stay closer to Premiere ordering.
    clipRootVisualNode.children.sort((leftNode, rightNode) => {
      const leftKey = String(leftNode.label || "").replace(/^Clip\s+/i, "").trim().toLowerCase();
      const rightKey = String(rightNode.label || "").replace(/^Clip\s+/i, "").trim().toLowerCase();
      const leftOrder = clipLevelSectionOrder.get(leftKey);
      const rightOrder = clipLevelSectionOrder.get(rightKey);
      if (typeof leftOrder === "number" && typeof rightOrder === "number" && leftOrder !== rightOrder) {
        return leftOrder - rightOrder;
      }
      if (typeof leftOrder === "number") {
        return -1;
      }
      if (typeof rightOrder === "number") {
        return 1;
      }
      return String(leftNode.label || "").localeCompare(String(rightNode.label || ""));
    });
    topLevelVisualNodes.push(clipRootVisualNode);
  }

  const appendVisualPropertyRows = (targetBody: HTMLElement, groupProperties: HostVisualProperty[]): void => {
    // // Render one subsection worth of editable controls while preserving existing control-specific behaviors.
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
        if (property.cloneOnlyWhenDirty) {
          checkbox.dataset.visualCloneOnlyWhenDirty = "1";
        }
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
        targetBody.appendChild(row);
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
        if (property.cloneOnlyWhenDirty) {
          hiddenInput.dataset.visualCloneOnlyWhenDirty = "1";
        }
        if (property.fontToken) {
          // // Preserve the exact host font token so copied font writes do not degrade to fallback style guesses.
          hiddenInput.dataset.visualFontToken = property.fontToken;
        }

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
          visualDirtyPaths.add(property.path);
          scheduleLiveVisualApply(VISUAL_COLOR_LIVE_UPDATE_DEBOUNCE_MS);
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
          visualDirtyPaths.add(property.path);
          scheduleLiveVisualApply(VISUAL_COLOR_LIVE_UPDATE_DEBOUNCE_MS);
        });
        nativeColorInput.addEventListener("change", () => {
          setColorState(nativeColorInput.value || hiddenInput.value || initialHex);
          visualDirtyPaths.add(property.path);
          scheduleLiveVisualApply(VISUAL_COLOR_LIVE_UPDATE_DEBOUNCE_MS);
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
        if (property.cloneOnlyWhenDirty) {
          numberInput.dataset.visualCloneOnlyWhenDirty = "1";
        }

        rangeInput.addEventListener("input", () => {
          numberInput.value = rangeInput.value;
          visualDirtyPaths.add(property.path);
          scheduleLiveVisualApply();
        });
        numberInput.addEventListener("input", () => {
          rangeInput.value = numberInput.value;
          visualDirtyPaths.add(property.path);
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
        if (property.cloneOnlyWhenDirty) {
          hiddenInput.dataset.visualCloneOnlyWhenDirty = "1";
        }
        if (property.fontToken) {
          // // Preserve the exact host font token so copied font writes do not degrade to fallback style guesses.
          hiddenInput.dataset.visualFontToken = property.fontToken;
        }
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
          bindLiveUpdateEvent(input, "input", property.path);
        });
        syncVector();

        controlWrap.append(vectorWrap, hiddenInput);
      } else if (
        parseTextStyleVirtualPath(property.path)?.styleKey === "fontFamily" ||
        parseTextStyleVirtualPath(property.path)?.styleKey === "fontStyle"
      ) {
        // // Keep detected font family/style visible for reference and clone payloads, but do not expose manual edits here.
        const textStylePath = parseTextStyleVirtualPath(property.path);
        const readonlyValue = document.createElement("div");
        readonlyValue.className = "visual-readonly-value";
        readonlyValue.textContent = String(property.value ?? "").trim() || "—";

        const hiddenInput = document.createElement("input");
        hiddenInput.type = "hidden";
        hiddenInput.value = String(property.value ?? "");
        hiddenInput.dataset.visualPath = property.path;
        hiddenInput.dataset.visualType = property.valueType;
        hiddenInput.dataset.visualControlKind = property.controlKind;
        hiddenInput.dataset.visualRole = "value";
        if (property.cloneOnlyWhenDirty) {
          hiddenInput.dataset.visualCloneOnlyWhenDirty = "1";
        }
        if (property.fontToken) {
          // // Keep the raw font token attached to readonly font rows because those values are cloned, not manually edited.
          hiddenInput.dataset.visualFontToken = property.fontToken;
        }

        controlWrap.append(readonlyValue, hiddenInput);
        if (textStylePath?.styleKey === "fontFamily") {
          // // Point users to Premiere Properties for font edits now that system-font scanning is disabled.
          const readonlyHint = document.createElement("p");
          readonlyHint.className = "visual-readonly-hint";
          readonlyHint.textContent = translate("visual.fontReadonlyHint");
          controlWrap.appendChild(readonlyHint);
        }
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
        if (property.cloneOnlyWhenDirty) {
          select.dataset.visualCloneOnlyWhenDirty = "1";
        }
        if (property.fontToken) {
          // // Pass through exact host tokens for any future editable text-style selects.
          select.dataset.visualFontToken = property.fontToken;
        }
        bindFloatingPanelSelect(select);
        bindLiveUpdateEvent(select, "change");

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
        if (property.cloneOnlyWhenDirty) {
          input.dataset.visualCloneOnlyWhenDirty = "1";
        }
        if (property.fontToken) {
          // // Keep exact host tokens on plain inputs when text-style fields are rendered this way.
          input.dataset.visualFontToken = property.fontToken;
        }
        bindLiveUpdateEvent(input, "input");
        controlWrap.appendChild(input);
      }

      row.appendChild(controlWrap);
      targetBody.appendChild(row);
    }
  };

  const renderVisualNode = (node: VisualEditorSectionNode, depth: number): HTMLDetailsElement => {
    // // Render one collapsible hierarchy node so Premiere group/components can stay visually nested.
    const groupNode = document.createElement("details");
    groupNode.className = "visual-group";
    if (depth > 0) {
      groupNode.classList.add("visual-group--nested");
    }
    groupNode.dataset.groupName = node.key;
    groupNode.dataset.groupDepth = String(depth);
    groupNode.open = visualOpenGroups.has(node.key) || depth === 0;
    groupNode.addEventListener("toggle", () => {
      if (groupNode.open) {
        visualOpenGroups.add(node.key);
      } else {
        visualOpenGroups.delete(node.key);
      }
    });

    const summary = document.createElement("summary");
    summary.className = "visual-group__title";
    summary.textContent = node.label;
    groupNode.appendChild(summary);

    const groupBody = document.createElement("div");
    groupBody.className = "visual-group__body";
    appendVisualPropertyRows(groupBody, node.properties);
    for (const childNode of node.children) {
      groupBody.appendChild(renderVisualNode(childNode, depth + 1));
    }
    groupNode.appendChild(groupBody);
    return groupNode;
  };

  for (const node of topLevelVisualNodes) {
    const renderedNode = normalizeVisualNodeForRender(node);
    elements.visualPropertyList.appendChild(renderVisualNode(renderedNode, 0));
  }
  refreshVisualButtonsBusyState();
  renderVisualSelectionSummary();
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

function collectVisualPropertyChanges(options?: { includeUnchanged?: boolean }): VisualPropertyChange[] {
  // // Build host payload from rendered editor controls and skip unchanged values unless a full cross-clip clone is requested.
  if (!elements.visualPropertyList) {
    return [];
  }

  const includeUnchanged = options?.includeUnchanged === true;
  const restrictToDirtyPaths = !includeUnchanged && loadedVisualSelectionCount <= 1;
  const loadedPropertyByPath = new Map(loadedVisualProperties.map((property) => [property.path, property]));
  const changes: VisualPropertyChange[] = [];
  const controls = elements.visualPropertyList.querySelectorAll<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(
    '[data-visual-role="value"]'
  );
  controls.forEach((control) => {
    const path = String(control.dataset.visualPath || "");
    const valueType = (String(control.dataset.visualType || "string") as HostVisualProperty["valueType"]) || "string";
    const controlKind =
      (String(control.dataset.visualControlKind || "string") as HostVisualProperty["controlKind"]) || "string";
    const cloneOnlyWhenDirty = String(control.dataset.visualCloneOnlyWhenDirty || "") === "1";
    const excludeFromClone = loadedPropertyByPath.get(path)?.excludeFromClone === true;
    const explicitFontToken = String(control.dataset.visualFontToken || "").trim();
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
    if (restrictToDirtyPaths && !visualDirtyPaths.has(path)) {
      return;
    }
    if (includeUnchanged && excludeFromClone) {
      // // Clip Duration and future timing-only controls must remain specific to each target clip during Copy/Apply.
      return;
    }
    if (includeUnchanged && cloneOnlyWhenDirty && !visualDirtyPaths.has(path)) {
      // // Keep ambiguous Premiere-only controls visible, but never clone them across clips unless the user changed them explicitly.
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

    const currentCanonical = canonicalizeVisualValue(controlKind, valueType, value);
    const originalCanonical = String(visualOriginalValuesByPath.get(path) || "");
    if (!includeUnchanged && currentCanonical === originalCanonical) {
      return;
    }

    const textStylePath = parseTextStyleVirtualPath(path);
    const fontToken =
      explicitFontToken ||
      (textStylePath && (textStylePath.styleKey === "fontFamily" || textStylePath.styleKey === "fontStyle")
        ? resolveVisualTextStyleToken(textStylePath.basePath)
        : "");

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

function commitAppliedVisualChanges(changes: VisualPropertyChange[]): void {
  // // Advance the live-update baseline after a successful apply so reverting a slider back to 0 still counts as a fresh change.
  for (const change of changes) {
    visualOriginalValuesByPath.set(
      change.path,
      canonicalizeVisualValue(change.controlKind, change.valueType, change.value)
    );
    visualDirtyPaths.delete(change.path);
  }
}

function copyLoadedVisualProperties(): void {
  // // Snapshot the currently loaded visual controls so Apply can clone them onto a later selection, including font tokens.
  if (!loadedVisualProperties.length) {
    throw new Error(translate("visual.noProperties"));
  }

  const copiedChanges = collectVisualPropertyChanges({ includeUnchanged: true });
  if (!copiedChanges.length) {
    throw new Error(translate("visual.noProperties"));
  }

  updateCopiedVisualSnapshot(copiedChanges);
  setStructuredLog(translate("log.visualCopyDone"), {
    selectedCount: loadedVisualSelectionCount,
    copiedCount: copiedChanges.length,
    sourceRevision: copiedVisualSourceRevision
  });
}

function clearVisualSelectionAutoRefreshTimer(): void {
  // // Cancel a queued selection refresh when the Visual editor leaves the foreground or another refresh supersedes it.
  if (visualSelectionAutoRefreshTimer !== null) {
    window.clearTimeout(visualSelectionAutoRefreshTimer);
    visualSelectionAutoRefreshTimer = null;
  }
}

function hasBlockingVisualEditorChanges(): boolean {
  // // Live update is always enabled now, so selection changes can refresh the editor without preserving stale manual edits.
  return false;
}

function isVisualSelectionMonitoringAllowed(): boolean {
  // // Avoid host calls while the Visual tab or the docked CEP document is not actually visible.
  return activeMode === "visual" && document.visibilityState !== "hidden";
}

async function syncVisualSelectionWatcherState(): Promise<void> {
  // // Keep Premiere's bound selection callback silent whenever this panel cannot use its events.
  try {
    await setVisualSelectionWatcherEnabled(isVisualSelectionMonitoringAllowed());
  } catch (error) {
    if (verboseLogsEnabled) {
      setStructuredLog(translate("log.hostResult"), {
        visualSelectionWatcherStateFailed: String(error)
      });
    }
  }
}

async function refreshVisualPropertiesIfSelectionChanged(reason: string): Promise<void> {
  // // Compare a lightweight host signature before doing the expensive full Visual editor read.
  if (
    !isVisualSelectionMonitoringAllowed() ||
    visualSelectionRefreshInFlight ||
    visualReadInProgress ||
    visualApplyInProgress
  ) {
    return;
  }

  visualSelectionRefreshInFlight = true;
  try {
    const selectionSignature = await readSelectedMogrtVisualSignature();
    const nextSignature = String(selectionSignature.signature || "");
    if (!nextSignature) {
      return;
    }
    if (nextSignature && nextSignature === lastVisualSelectionSignature) {
      return;
    }

    if (hasBlockingVisualEditorChanges()) {
      pendingVisualSelectionChangeNotice = true;
      renderVisualSelectionSummary();
      return;
    }

    lastVisualSelectionSignature = nextSignature;
    pendingVisualSelectionChangeNotice = false;
    await loadVisualPropertiesFromSelection(false, false);
  } catch (error) {
    if (reason === "poll") {
      return;
    }
    if (verboseLogsEnabled) {
      setStructuredLog(translate("log.hostResult"), {
        visualSelectionAutoRefreshFailed: String(error)
      });
    }
  } finally {
    visualSelectionRefreshInFlight = false;
  }
}

function scheduleVisualSelectionAutoRefresh(reason: string): void {
  // // Debounce noisy Premiere selection events before comparing the current selection signature.
  if (!isVisualSelectionMonitoringAllowed()) {
    return;
  }

  clearVisualSelectionAutoRefreshTimer();
  visualSelectionAutoRefreshTimer = window.setTimeout(() => {
    visualSelectionAutoRefreshTimer = null;
    void refreshVisualPropertiesIfSelectionChanged(reason);
  }, VISUAL_SELECTION_AUTO_REFRESH_DEBOUNCE_MS);
}

function startVisualSelectionPolling(): void {
  // // Keep a light fallback active because some CEP reloads or host versions can miss custom selection events.
  if (visualSelectionPollTimer !== null) {
    return;
  }

  visualSelectionPollTimer = window.setInterval(() => {
    if (isVisualSelectionMonitoringAllowed()) {
      void refreshVisualPropertiesIfSelectionChanged("poll");
    }
  }, VISUAL_SELECTION_POLL_INTERVAL_MS);
}

function stopVisualSelectionPolling(): void {
  // // Stop fallback polling outside the Visual editor so other tabs stay quiet.
  if (visualSelectionPollTimer !== null) {
    window.clearInterval(visualSelectionPollTimer);
    visualSelectionPollTimer = null;
  }
}

async function initializeVisualSelectionWatcher(): Promise<void> {
  // // Bridge Premiere timeline-selection events into the panel and keep polling as a safety net.
  if (!visualSelectionWatcherCleanup) {
    visualSelectionWatcherCleanup = addVisualSelectionChangedListener(() => {
      scheduleVisualSelectionAutoRefresh("event");
    });
  }

  try {
    await registerVisualSelectionWatcher();
    await syncVisualSelectionWatcherState();
  } catch (error) {
    if (verboseLogsEnabled) {
      setStructuredLog(translate("log.hostResult"), {
        visualSelectionWatcherUnavailable: String(error)
      });
    }
  }

  if (isVisualSelectionMonitoringAllowed()) {
    startVisualSelectionPolling();
    scheduleVisualSelectionAutoRefresh("startup");
  }
}

async function loadVisualPropertiesFromSelection(emitHostLog = false, showLoadingState = false): Promise<void> {
  // // Read selected MOGRT editable controls from host and refresh visual editor UI.
  if (showLoadingState) {
    if (visualReadInProgress || visualApplyInProgress) {
      return;
    }
    visualReadInProgress = true;
    refreshVisualButtonsBusyState();
    setVisualApplyProgressState(true, 0, 0, translate("progress.visualReadPending"));
    await waitForNextPaint();
  }

  try {
    const result = await readSelectedMogrtVisualProperties({
      // // Avoid raw color/text payload diagnostics during normal automatic selection refreshes.
      includeDebug: verboseLogsEnabled
    });
    pendingVisualSelectionChangeNotice = false;
    lastVisualSelectionSignature = String(result.signature || lastVisualSelectionSignature || "");
    loadedVisualSelectionCount = Number(result.selectedCount || 0);
    loadedVisualComponents = Array.isArray(result.debug?.components) ? result.debug.components.slice() : [];
    renderVisualPropertyEditor(result.properties);
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
    refreshVisualButtonsBusyState();
  } finally {
    if (showLoadingState) {
      visualReadInProgress = false;
      refreshVisualButtonsBusyState();
      setVisualApplyProgressState(false);
    }
  }
}

function isVisualLiveUpdateEnabled(): boolean {
  // // Visual editor writes are always live; the old toggle is intentionally ignored.
  return true;
}

async function applyVisualChangesToSelection(options?: { liveUpdate?: boolean }): Promise<void> {
  // // Apply edited visual values in one host batch so large subtitle selections do not repeat CEP work per clip.
  const useLiveUpdate = options?.liveUpdate === true;

  if (visualReadInProgress) {
    return;
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
    refreshVisualButtonsBusyState();
    setVisualApplyProgressState(true, 0, 0, translate("progress.visualApplyPending"));
  }

  try {
    const selectedCount = await getSelectedMogrtCount();
    const dirtyChanges = collectVisualPropertyChanges();
    let changes = dirtyChanges;
    const canMergeDirtyIntoCopiedSnapshot =
      copiedVisualChanges.length > 0 && copiedVisualSourceRevision === visualLoadedRevision;

    if (!useLiveUpdate && copiedVisualChanges.length > 0) {
      // // Let an explicit Copy properties snapshot drive Apply, even when the target selection changed or only one clip is selected.
      changes = canMergeDirtyIntoCopiedSnapshot
        ? mergeVisualPropertyChanges(copiedVisualChanges, dirtyChanges)
        : cloneVisualPropertyChanges(copiedVisualChanges);
    }

    if (!changes.length) {
      if (useLiveUpdate) {
        return;
      }

      // // Keep the one-click "copy first selected clip to the rest" workflow by sending a full clone only for explicit multi-selection applies.
      if (selectedCount > 1) {
        changes = collectVisualPropertyChanges({ includeUnchanged: true });
      }
    }

    if (!changes.length) {
      throw new Error(translate("visual.noChanges"));
    }

    if (!useLiveUpdate) {
      setVisualApplyProgressState(true, 0, Math.max(1, selectedCount), translate("progress.visualApplyPending"));
      await waitForNextPaint();
    }
    const response = await applyVisualPropertiesToSelectedMogrts(changes, {
      // // Return costly host-level diagnostics only when the panel's verbose log mode is enabled.
      includeDebug: verboseLogsEnabled
    });
    if (useLiveUpdate && Number(response.failedCount || 0) === 0) {
      commitAppliedVisualChanges(changes);
    }
    if (!useLiveUpdate) {
      setVisualApplyProgressState(true, Math.max(1, selectedCount), Math.max(1, selectedCount));
      setStructuredLog(translate("log.visualApplyDone"), response);
      await loadVisualPropertiesFromSelection();
    }
  } finally {
    visualApplyInProgress = false;
    if (!useLiveUpdate) {
      refreshVisualButtonsBusyState();
      setVisualApplyProgressState(false);
    }
  }
}

function scheduleLiveVisualApply(delayMs = VISUAL_LIVE_UPDATE_DEBOUNCE_MS): void {
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
  }, Math.max(0, delayMs));
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
      }, VISUAL_LIVE_UPDATE_DEBOUNCE_MS);
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
  const previousAspectFilter = pendingMogrtAspectFilter || elements.mogrtAspectFilter?.value || "all";
  const previousSearchQuery = pendingMogrtSearchQuery || elements.mogrtSearchInput?.value || "";

  // // Keep startup-restored gallery state authoritative so the default `all` option does not wipe the persisted folder filter.
  pendingMogrtAspectFilter = previousAspectFilter;
  pendingMogrtSearchQuery = previousSearchQuery;

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

function getNormalizedMogrtSearchQuery(): string {
  // // Keep gallery search matching case-insensitive and whitespace-tolerant across sessions.
  return String(elements.mogrtSearchInput?.value || pendingMogrtSearchQuery || "")
    .trim()
    .toLocaleLowerCase();
}

function renderMogrtGallery(): void {
  // // Render gallery cards with lightweight visual previews and aspect filtering.
  if (!elements.mogrtGallery || !elements.mogrtAspectFilter) {
    return;
  }

  const selectedAspect = elements.mogrtAspectFilter.value;
  const normalizedSearchQuery = getNormalizedMogrtSearchQuery();
  const filtered = availableMogrts.filter((template) => {
    const matchesAspect = selectedAspect === "all" || template.aspect === selectedAspect;
    const matchesSearch =
      normalizedSearchQuery.length < 1 || String(template.name || "").toLocaleLowerCase().includes(normalizedSearchQuery);
    return matchesAspect && matchesSearch;
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

async function browseCorrectedTranscriptPath(): Promise<void> {
  // // Pick a corrected transcript file used by WhisperX alignment.
  if (!elements.correctedTranscriptPath) {
    return;
  }

  const selectedPath = await pickCorrectedTranscriptPath();
  if (selectedPath) {
    elements.correctedTranscriptPath.value = selectedPath;
    persistPanelState();
  }
}

function collectBuildOptions(): CaptionBuildOptions {
  // // Collect, normalize, and validate all panel options into a single object.
  if (
    !elements.sourceMode ||
    !elements.languageSelect ||
    !elements.whisperLanguage ||
    !elements.maxChars ||
    !elements.maxWords ||
    !elements.linesPerCaption ||
    !elements.animationMode ||
    !elements.correctedTranscriptPath ||
    !elements.whisperModel ||
    !elements.whisperSequenceRange ||
    !elements.preserveTranslationTiming ||
    !elements.preserveMixedLanguages ||
    !elements.mixedLanguagePrompt
    || !elements.removePunctuation
  ) {
    throw new Error("Panel bindings not initialized.");
  }

  if (!selectedMogrt && availableMogrts.length > 0) {
    selectedMogrt = availableMogrts[0];
  }

  if (
    (getSourceMode() === "whisper_sequence" || getSourceMode() === "whisperx_sequence") &&
    !String(elements.whisperModel.value || "").trim()
  ) {
    throw new Error(`${translate("error.whisperModelMissing")}\n\n${buildWhisperDiagnostic()}`);
  }
  if (getSourceMode() === "whisper_sequence" && !whisperRuntimeAvailable) {
    throw new Error(`${translate("error.whisperUnavailable")}\n\n${buildWhisperDiagnostic()}`);
  }
  if (getSourceMode() === "whisperx_sequence" && !correctedAlignAvailable) {
    throw new Error(`${translate("error.whisperxUnavailable")}\n\n${buildWhisperDiagnostic()}`);
  }
  if (getSourceMode() === "corrected_align" && !String(elements.correctedTranscriptPath.value || "").trim()) {
    throw new Error(translate("error.missingCorrectedTranscriptPath"));
  }
  if (getSourceMode() === "corrected_align" && !correctedAlignAvailable) {
    const runtimeDetail = correctedAlignRuntimeDetails ? ` ${correctedAlignRuntimeDetails}` : "";
    throw new Error(`${translate("error.correctedAlignUnavailable")}${runtimeDetail ? ` (${runtimeDetail})` : ""}`);
  }
  if (getSourceMode() === "corrected_align" && getSelectedWhisperLanguageCode() === "auto") {
    throw new Error(translate("error.correctedAlignLanguageRequired"));
  }

  const extensionRootPath = resolveExtensionRootPath();
  const templateRelativePath = selectedMogrt?.relativePath ?? "";

  return {
    sourceMode: getSourceMode(),
    outputMode: getOutputMode(),
    languageCode: getSelectedWhisperLanguageCode(),
    style: {
      maxCharsPerLine: Number(elements.maxChars.value),
      maxWordsPerLine: Number(elements.maxWords.value),
      animationMode: elements.animationMode.value as AnimationMode,
      uppercase: false,
      removePunctuation: Boolean(elements.removePunctuation.checked),
      linesPerCaption: Number(elements.linesPerCaption.value)
    },
    extensionRootPath,
    mogrtPath: buildAbsoluteMogrtPath(extensionRootPath, templateRelativePath),
    mogrtTemplateRelativePath: templateRelativePath,
    correctedTranscriptPath: String(elements.correctedTranscriptPath.value || "").trim(),
    whisperModel: elements.whisperModel.value,
    whisperSequenceRange: (elements.whisperSequenceRange.value as WhisperSequenceRangeMode) || "entire_sequence",
    preserveTranslationTiming: Boolean(elements.preserveTranslationTiming.checked),
    preserveMixedLanguages: Boolean(elements.preserveMixedLanguages.checked),
    mixedLanguagePrompt: sanitizeWhisperGlossary(elements.mixedLanguagePrompt.value),
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
  const percent = String(Math.max(0, Math.min(100, Math.round(Number(progress.percent || 0)))));
  const remaining = String(progress.remaining || "").trim();
  if (remaining) {
    return translateTemplate("progress.whisperAnalysisRemaining", {
      percent,
      remaining
    });
  }

  return translateTemplate("progress.whisperAnalysis", {
    percent
  });
}

function formatProgressDuration(totalSeconds: number): string {
  // // Format generated progress durations like Whisper/tqdm so elapsed/remaining text stays compact.
  const safeSeconds = Math.max(0, Math.floor(Number(totalSeconds) || 0));
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const seconds = safeSeconds % 60;
  if (hours > 0) {
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }
  return `${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

function estimateWhisperRuntimeSeconds(audioDurationSeconds: number, modelName: string): number {
  // // Estimate runtime only as a fallback when Whisper does not stream tqdm progress from CEP.
  const duration = Math.max(1, Number(audioDurationSeconds) || 0);
  const normalizedModel = String(modelName || "").toLowerCase();
  let multiplier = 0.75;
  if (normalizedModel.includes("tiny")) {
    multiplier = 0.35;
  } else if (normalizedModel.includes("base")) {
    multiplier = 0.6;
  } else if (normalizedModel.includes("small")) {
    multiplier = 1.0;
  } else if (normalizedModel.includes("medium")) {
    multiplier = 1.8;
  } else if (normalizedModel.includes("large")) {
    multiplier = 3.0;
  } else if (normalizedModel.includes("turbo")) {
    multiplier = 0.55;
  }

  return Math.max(20, duration * multiplier);
}

function buildEstimatedWhisperProgress(
  elapsedSeconds: number,
  audioDurationSeconds: number,
  modelName: string
): { percent: number; remainingSeconds: number } {
  // // Keep fallback progress moving conservatively and leave the final stretch for real process completion.
  const expectedSeconds = estimateWhisperRuntimeSeconds(audioDurationSeconds, modelName);
  const percent = Math.max(1, Math.min(95, Math.round((Math.max(0, elapsedSeconds) / expectedSeconds) * 100)));
  const remainingSeconds = Math.max(0, Math.round(expectedSeconds - Math.max(0, elapsedSeconds)));
  return { percent, remainingSeconds };
}

function buildEstimatedWhisperProgressLabel(percent: number, remainingSeconds: number, elapsedSeconds: number): string {
  // // Tell the user when progress is estimated so it is not confused with Whisper's real tqdm percentage.
  if (remainingSeconds > 0) {
    return translateTemplate("progress.whisperAnalysisEstimatedRemaining", {
      percent: String(percent),
      remaining: formatProgressDuration(remainingSeconds)
    });
  }

  return translateTemplate("progress.whisperAnalysisElapsed", {
    elapsed: formatProgressDuration(elapsedSeconds)
  });
}

function mapCorrectedAlignPercentToGenerateProgress(progress: WhisperProgressUpdate): number {
  // // Reserve the middle of the generate bar for corrected transcript alignment while keeping room for export and apply.
  const clampedPercent = Math.max(0, Math.min(100, Number(progress.percent || 0)));
  return 30 + Math.round((clampedPercent / 100) * 52);
}

function buildCorrectedAlignProgressLabel(progress: WhisperProgressUpdate): string {
  // // Make corrected transcript alignment progress explicit so the user sees real activity during WhisperX processing.
  return translateTemplate("progress.correctedAlignAnalysis", {
    percent: String(Math.max(0, Math.min(100, Math.round(Number(progress.percent || 0)))))
  });
}

function buildWhisperXProgressLabel(progress: WhisperProgressUpdate): string {
  // // Show WhisperX transcription and forced-alignment progress as one precision-timing analysis step.
  return translateTemplate("progress.whisperxAnalysis", {
    percent: String(Math.max(0, Math.min(100, Math.round(Number(progress.percent || 0)))))
  });
}

function assertHostApplySucceeded(rawResult: string): ApplyCaptionPlanHostResult {
  // // Surface host-side apply failures immediately instead of leaving them only in the expanded log payload.
  const parsed = JSON.parse(String(rawResult || "{}")) as ApplyCaptionPlanHostResult;
  if (parsed && parsed.ok === false) {
    throw new Error(String(parsed.error || "Premiere did not apply the generated subtitles."));
  }
  return parsed;
}

async function resolveRequestedSequenceRange(
  options: CaptionBuildOptions
): Promise<{
  rangeStartSeconds?: number;
  rangeEndSeconds?: number;
  sequenceName?: string;
  fallbackReason?: string;
  hostError?: string;
  debug?: unknown;
}> {
  // // Read the active sequence In/Out only when the user explicitly selected range-limited generation.
  // // Fall back to the full sequence when Premiere has no valid In/Out range set yet.
  if (options.whisperSequenceRange !== "in_out") {
    return {};
  }

  const range = await getActiveSequenceRange();
  if (range.fallbackReason || range.hostError) {
    return {
      fallbackReason: range.fallbackReason,
      hostError: range.hostError,
      debug: range.debug
    };
  }

  const rangeStartSeconds = Number(range.rangeStartSeconds);
  const rangeEndSeconds = Number(range.rangeEndSeconds);
  if (!Number.isFinite(rangeStartSeconds) || !Number.isFinite(rangeEndSeconds) || rangeEndSeconds <= rangeStartSeconds) {
    return {
      fallbackReason: "No valid In/Out range is set; using entire sequence.",
      sequenceName: range.sequenceName
    };
  }

  return {
    rangeStartSeconds,
    rangeEndSeconds,
    sequenceName: range.sequenceName
  };
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
  onProgress?: (done: number, label: string, waitForPaint?: boolean) => Promise<void>,
  assertNotCancelled: () => void = () => {}
): Promise<CaptionCue[]> {
  // // Build cues from the currently selected source mode.
  assertNotCancelled();
  const requestedSequenceRange = await resolveRequestedSequenceRange(options);
  assertNotCancelled();
  setStructuredLog(translate("log.sequenceRange"), requestedSequenceRange);

  if (options.sourceMode === "srt") {
    if (!elements.srtPath || !elements.srtPath.value.trim()) {
      throw new Error(translate("error.missingSrtPath"));
    }

    if (onProgress) {
      await onProgress(10, translate("progress.readSrt"), true);
    }
    assertNotCancelled();
    const srtText = await readTextFileFromHost(elements.srtPath.value.trim());
    assertNotCancelled();
    if (onProgress) {
      await onProgress(28, translate("progress.parseSrt"));
    }
    const cues = parseSrt(srtText);
    if (!cues.length) {
      throw new Error(translate("error.emptySrt"));
    }

    return trimSrtCuesToRange(
      cues,
      Number.isFinite(Number(requestedSequenceRange.rangeStartSeconds))
        ? Number(requestedSequenceRange.rangeStartSeconds)
        : Number.NaN,
      Number.isFinite(Number(requestedSequenceRange.rangeEndSeconds))
        ? Number(requestedSequenceRange.rangeEndSeconds)
        : Number.NaN
    );
  }

  if (options.sourceMode === "corrected_align") {
    if (!options.correctedTranscriptPath) {
      throw new Error(translate("error.missingCorrectedTranscriptPath"));
    }

    if (onProgress) {
      await onProgress(10, translate("progress.readCorrectedTranscript"), true);
    }
    assertNotCancelled();
    const correctedTranscriptText = await readTextFileFromHost(options.correctedTranscriptPath);
    assertNotCancelled();
    if (!String(correctedTranscriptText || "").trim()) {
      throw new Error(translate("error.emptyCorrectedTranscript"));
    }

    setLog(translate("log.correctedAlignExport"));
    if (onProgress) {
      await onProgress(18, translate("progress.exportSequence"), true);
    }
    const exportResult = await exportActiveSequenceAudioForWhisper(options.whisperSequenceRange, options.extensionRootPath);
    assertNotCancelled();
    const cleanupAudioPath = exportResult.audioPath;
    if (!cleanupAudioPath) {
      throw new Error(translate("error.missingActiveSequenceAudio"));
    }

    try {
      if (onProgress) {
        await onProgress(30, translate("progress.correctedAligning"), true);
      }
      assertNotCancelled();
      const alignmentResult = await alignCorrectedTranscript(
        {
          audioPath: cleanupAudioPath,
          transcriptPath: options.correctedTranscriptPath,
          languageCode: options.languageCode,
          extensionRootPath: options.extensionRootPath,
          rangeStartSeconds: requestedSequenceRange.rangeStartSeconds,
          rangeEndSeconds: requestedSequenceRange.rangeEndSeconds
        },
        (progress) => {
          void updateGenerateProgress(
            mapCorrectedAlignPercentToGenerateProgress(progress),
            buildCorrectedAlignProgressLabel(progress)
          );
        }
      );
      assertNotCancelled();

      if (onProgress) {
        await onProgress(84, translate("progress.parseWhisper"));
      }
      const cues = parseWhisperJson(alignmentResult.jsonText);
      if (!cues.length) {
        throw new Error(translate("error.emptyCorrectedAlign"));
      }

      setLog(translate("log.correctedAlignDone"));
      return shiftCaptionCues(cues, Number(requestedSequenceRange.rangeStartSeconds) || 0);
    } finally {
      await deleteTemporaryWhisperAudio(cleanupAudioPath);
    }
  }

  const whisperxModeActive = options.sourceMode === "whisperx_sequence";
  setLog(translate(whisperxModeActive ? "log.whisperxSequenceExport" : "log.whisperSequenceExport"));
  if (onProgress) {
    await onProgress(10, translate("progress.exportSequence"), true);
  }
  const exportResult = await exportActiveSequenceAudioForWhisper(options.whisperSequenceRange, options.extensionRootPath);
  assertNotCancelled();
  setStructuredLog(translate("log.audioExportDone"), exportResult);
  const whisperAudioPath = exportResult.audioPath;
  const cleanupAudioPath = exportResult.audioPath;

  if (!whisperAudioPath) {
    throw new Error(translate("error.missingActiveSequenceAudio"));
  }

  try {
    if (onProgress) {
      await onProgress(22, translate(whisperxModeActive ? "progress.whisperxAnalyzing" : "progress.whisperAnalyzing"), true);
    }
    assertNotCancelled();
    let latestRealWhisperProgressAt = 0;
    let latestFallbackPercent = 0;
    const whisperStartedAt = Date.now();
    const audioDurationSeconds = Number(exportResult.audioDurationSeconds || 0);
    const fallbackProgressTimer =
      Number.isFinite(audioDurationSeconds) && audioDurationSeconds > 0
        ? window.setInterval(() => {
            if (Date.now() - latestRealWhisperProgressAt < 2500) {
              return;
            }
            const elapsedSeconds = Math.floor((Date.now() - whisperStartedAt) / 1000);
            const estimatedProgress = buildEstimatedWhisperProgress(elapsedSeconds, audioDurationSeconds, options.whisperModel);
            if (estimatedProgress.percent < latestFallbackPercent) {
              return;
            }
            latestFallbackPercent = estimatedProgress.percent;
            void updateGenerateProgress(
              mapWhisperPercentToGenerateProgress({ percent: estimatedProgress.percent, detail: "estimated" }),
              buildEstimatedWhisperProgressLabel(
                estimatedProgress.percent,
                estimatedProgress.remainingSeconds,
                elapsedSeconds
              )
            );
          }, 1000)
        : window.setInterval(() => {
            if (Date.now() - latestRealWhisperProgressAt < 2500) {
              return;
            }
            const elapsedSeconds = Math.floor((Date.now() - whisperStartedAt) / 1000);
            void updateGenerateProgress(22, translateTemplate("progress.whisperAnalysisElapsed", {
              elapsed: formatProgressDuration(elapsedSeconds)
            }));
          }, 1000);
    let whisperResult;
    let sourceTimingResult: Awaited<ReturnType<typeof transcribeWithWhisper>> | null = null;
    try {
      const initialPrompt = buildWhisperInitialPrompt(options);
      const preserveTranslationTiming =
        options.preserveTranslationTiming &&
        options.sourceMode === "whisper_sequence" &&
        options.languageCode.toLowerCase() !== "auto";
      if (preserveTranslationTiming) {
        // // Run auto-detect first because its timestamps follow the spoken language instead of the forced-language output.
        if (onProgress) {
          await onProgress(22, translate("progress.whisperSourceTiming"), true);
        }
        setLog(translate("log.whisperSourceTiming"));
        sourceTimingResult = await transcribeWithWhisper({
          audioPath: whisperAudioPath,
          languageCode: "auto",
          model: options.whisperModel,
          initialPrompt
        });
        assertNotCancelled();
        if (onProgress) {
          await onProgress(50, translate("progress.whisperForcedLanguage"), true);
        }
      }
      const transcriptionRequest = {
        audioPath: whisperAudioPath,
        languageCode: options.languageCode,
        model: options.whisperModel,
        initialPrompt,
        extensionRootPath: options.extensionRootPath
      };
      setStructuredLog(translate("log.whisperStarted"), {
        mode: options.sourceMode,
        model: options.whisperModel,
        languageCode: options.languageCode,
        preserveMixedLanguages: options.preserveMixedLanguages,
        initialPromptUsed: Boolean(initialPrompt),
        mixedLanguagePromptLength: sanitizeWhisperGlossary(options.mixedLanguagePrompt).length,
        audioPath: whisperAudioPath
      });
      whisperResult = await (whisperxModeActive
        ? transcribeWithWhisperX(transcriptionRequest, (progress) => {
            latestRealWhisperProgressAt = Date.now();
            void updateGenerateProgress(mapCorrectedAlignPercentToGenerateProgress(progress), buildWhisperXProgressLabel(progress));
          })
        : transcribeWithWhisper(transcriptionRequest, (progress) => {
            latestRealWhisperProgressAt = Date.now();
            const mappedProgress = preserveTranslationTiming
              ? 50 + Math.round((mapWhisperPercentToGenerateProgress(progress) - 22) * 0.5)
              : mapWhisperPercentToGenerateProgress(progress);
            void updateGenerateProgress(mappedProgress, buildWhisperProgressLabel(progress));
          }));
    } finally {
      window.clearInterval(fallbackProgressTimer);
    }
    assertNotCancelled();
    setStructuredLog(translate("log.whisperResult"), {
      model: whisperResult.model,
      audioPath: whisperResult.audioPath,
      hasJson: Boolean(whisperResult.jsonText),
      srtLength: String(whisperResult.srtText || "").length,
      commandOutput: whisperResult.commandOutput || ""
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
    const glossaryCues = options.preserveMixedLanguages
      ? applyWhisperGlossaryToCues(fallbackCues, options.mixedLanguagePrompt)
      : fallbackCues;
    let displayCues = whisperxModeActive ? normalizeWhisperXCuesForDisplay(glossaryCues) : glossaryCues;
    if (sourceTimingResult) {
      let sourceTimingCues: CaptionCue[] = [];
      if (sourceTimingResult.jsonText) {
        try {
          sourceTimingCues = parseWhisperJson(sourceTimingResult.jsonText);
        } catch {
          // // Fall back to SRT timing when the auto-detect Whisper JSON is unavailable on this runtime.
          sourceTimingCues = [];
        }
      }
      sourceTimingCues = sourceTimingCues.length > 0 ? sourceTimingCues : parseSrt(sourceTimingResult.srtText);
      const timingLockedCues = lockTranslatedCuesToSourceTiming(sourceTimingCues, displayCues);
      if (timingLockedCues.length > 0) {
        displayCues = timingLockedCues;
        setLog(translate("log.whisperTimingLocked"));
      }
    }
    if (!displayCues.length) {
      throw new Error(translate("error.emptyWhisper"));
    }

    setStructuredLog(translate("log.cuesReady"), {
      count: displayCues.length,
      firstStartSeconds: displayCues[0]?.startSeconds,
      lastEndSeconds: displayCues[displayCues.length - 1]?.endSeconds
    });
    setLog(`${translate(whisperxModeActive ? "log.whisperxDone" : "log.whisperDone")} ${whisperResult.model}`);
    return shiftCaptionCues(displayCues, Number(requestedSequenceRange.rangeStartSeconds) || 0);
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
  generateCancelRequested = false;
  setGenerateButtonsBusy(true);
  try {
    setLog(translate("log.generateRequested"));
    await updateGenerateProgress(1, translate("progress.prepareGeneration"), true);
    const requestedSourceMode = getSourceMode();
    if (
      requestedSourceMode === "whisper_sequence" ||
      requestedSourceMode === "whisperx_sequence" ||
      requestedSourceMode === "corrected_align"
    ) {
      setLog(translate("log.whisperRuntimeCheck"));
      await enforceWhisperSourceAvailability();
    }
    const options = collectBuildOptions();
    setStructuredLog(translate("log.generateConfiguration"), {
      sourceMode: options.sourceMode,
      outputMode: options.outputMode,
      languageCode: options.languageCode,
      whisperModel: options.whisperModel,
      whisperSequenceRange: options.whisperSequenceRange,
      preserveMixedLanguages: options.preserveMixedLanguages,
      mixedLanguageInitialPromptUsed: Boolean(buildWhisperInitialPrompt(options)),
      mixedLanguagePromptLength: sanitizeWhisperGlossary(options.mixedLanguagePrompt).length,
      mogrtTemplateRelativePath: options.mogrtTemplateRelativePath
    });
    setLog(translate("log.processing"));
    await updateGenerateProgress(4, translate("progress.prepareGeneration"), true);
    assertGenerateNotCancelled();
    const cues = await loadCuesFromSelectedSource(options, updateGenerateProgress, assertGenerateNotCancelled);
    assertGenerateNotCancelled();
    await updateGenerateProgress(90, translate("progress.planCaptions"), true);
    let plannedCues = buildCaptionPlan(cues, options);
    setStructuredLog(translate("log.captionPlanReady"), {
      sourceCueCount: cues.length,
      plannedCueCount: plannedCues.length,
      outputMode: options.outputMode
    });
    assertGenerateNotCancelled();
    if (options.outputMode === "premiere_subtitles") {
      const payload: HostApplyPayload = {
        options,
        cues: plannedCues
      };
      await updateGenerateProgress(98, translate("progress.applyCaptions"), true);
      const hostResultRaw = await applyNativeSubtitlePlan(payload);
      setStructuredLogFromRaw(translate("log.hostResult"), hostResultRaw);
      const hostResult = assertHostApplySucceeded(hostResultRaw);
      try {
        // // Never turn a successful Premiere subtitle creation into a failure when optional DeepL preparation cannot read the file.
        await prepareGeneratedNativeSrtForTranslation(String(hostResult.nativeSubtitleSrtPath || ""));
      } catch (error) {
        setLog(String(error), true);
      }
      return;
    }

    const premiereTemplateTextPayloads = await readPremiereTemplateTextPayloads(options.mogrtPath);
    let cleanupMogrtPaths: string[] = [];
    if (premiereTemplateTextPayloads.length) {
      const preparedTemplateBuild = await buildPremiereTemplateCueMogrts(
        options.mogrtPath,
        plannedCues,
        premiereTemplateTextPayloads
      );
      plannedCues = preparedTemplateBuild.cues;
      cleanupMogrtPaths = preparedTemplateBuild.cleanupPaths;
    }

    try {
      assertGenerateNotCancelled();
      const payload: HostApplyPayload = {
        options: {
          ...options,
          premiereTemplateTextPayloads
        },
        cues: plannedCues
      };

      await updateGenerateProgress(98, translate("progress.applyCaptions"), true);
      const hostResultRaw = await applyCaptionPlan(payload);
      setStructuredLogFromRaw(translate("log.hostResult"), hostResultRaw);
      const hostResult = assertHostApplySucceeded(hostResultRaw);
      try {
        if (
          Number(hostResult.insertedMogrt || 0) === plannedCues.length &&
          Number.isFinite(Number(hostResult.videoTrackUsed)) &&
          plannedCues.length > 0
        ) {
          persistGeneratedCaptionMetadata(
            resolveCaptionMetadataIdentityFromHostPayload(hostResult),
            Number(hostResult.videoTrackUsed),
            plannedCues
          );
        }
      } catch {
        // // Ignore host-result metadata persistence when the host payload is not parseable.
      }
    } finally {
      const modules = window.cep_node?.require?.("fs") as { unlinkSync?: (path: string) => void } | undefined;
      if (modules?.unlinkSync) {
        for (const cleanupPath of cleanupMogrtPaths) {
          try {
            modules.unlinkSync(cleanupPath);
          } catch {
            // // Ignore temporary-template cleanup failures because subtitle generation already completed.
          }
        }
      }
    }
  } catch (error) {
    if (isGenerateCancelledError(error)) {
      setLog(translate("log.generateCancelled"));
      return;
    }
    throw error;
  } finally {
    generateInProgress = false;
    generateCancelRequested = false;
    setGenerateButtonsBusy(false);
    setGenerateProgressState(false);
  }
}

async function stopCurrentGenerateJob(): Promise<void> {
  // // Stop the current Whisper or Whisper + SRT job and mark the whole generate flow as cancelled.
  if (!generateInProgress || generateCancelRequested) {
    return;
  }

  generateCancelRequested = true;
  setGenerateButtonsBusy(true);
  setLog(translate("log.generateCancelRequested"));

  try {
    await cancelCurrentJob();
  } catch (error) {
    setLog(String(error), true);
  }
}

async function initialize(): Promise<void> {
  // // Initialize locale, controls, and event listeners once panel is loaded.
  assertDomBindings();
  document.querySelectorAll<HTMLSelectElement>("select").forEach((select) => {
    // // Bind the durable floating dropdown fallback once for all static panel selects, including Whisper language.
    bindFloatingPanelSelect(select);
  });
  applyPremierePanelTheme();
  bindPremiereThemeListener();
  await loadPanelMeta();
  refreshVersionLabel();
  const persistedState = readPersistedPanelState();
  let whisperGlossaryLoadError = "";
  let shouldMigrateWhisperGlossaryToFile = false;
  try {
    const glossaryStore = await readWhisperGlossaryStore();
    shouldMigrateWhisperGlossaryToFile = glossaryStore.available && !glossaryStore.exists;
    if (glossaryStore.exists) {
      // // The user-profile file is canonical because it persists independently of Premiere projects and output modes.
      persistedState.whisperGlossaryEnabled = glossaryStore.enabled;
      persistedState.whisperGlossary = glossaryStore.text;
    }
  } catch (error) {
    // // Keep the localStorage fallback usable when a damaged or locked profile file cannot be read.
    whisperGlossaryLoadError = String(error);
  }

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
  refreshWhisperLanguageUi();
  toggleTranslationInputMode();
  applyPersistedPanelState(persistedState);
  setActiveMode(activeMode);
  toggleSourceFields();
  renderMogrtGallery();
  persistPanelState();
  if (shouldMigrateWhisperGlossaryToFile) {
    scheduleWhisperGlossaryStoreSave();
  }

  elements.languageSelect?.addEventListener("change", async () => {
    await loadLocale(elements.languageSelect?.value ?? "en");
    refreshWhisperLanguageUi(getSelectedWhisperLanguageCode());
    toggleSourceFields();
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
  elements.tabText?.addEventListener("click", () => {
    setActiveMode("text");
    persistPanelState();
  });
  elements.tabTranslate?.addEventListener("click", () => {
    setActiveMode("translate");
    persistPanelState();
  });

  elements.sourceMode?.addEventListener("change", () => {
    toggleSourceFields();
    persistPanelState();
  });
  elements.outputMode?.addEventListener("change", () => {
    captureActiveOutputModeSettings();
    activeOutputMode = getOutputMode();
    applyOutputModeSettingsToControls(activeOutputMode);
    toggleSourceFields();
    renderMogrtGallery();
    setGenerateButtonsBusy(generateInProgress);
    persistPanelState();
  });

  elements.srtBrowseButton?.addEventListener("click", async () => {
    try {
      await browseSrtPath();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.correctedTranscriptBrowseButton?.addEventListener("click", async () => {
    try {
      await browseCorrectedTranscriptPath();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.mogrtAspectFilter?.addEventListener("change", () => {
    pendingMogrtAspectFilter = elements.mogrtAspectFilter?.value || "all";
    renderMogrtGallery();
    persistPanelState();
  });
  elements.mogrtSearchInput?.addEventListener("input", () => {
    pendingMogrtSearchQuery = elements.mogrtSearchInput?.value || "";
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
  elements.whisperModelFolderButton?.addEventListener("click", async () => {
    try {
      await openWhisperModelsFolder(whisperModelCachePaths);
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.animationMode?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.removePunctuation?.addEventListener("change", () => {
    // // Persist the per-output punctuation preference immediately after the user toggles it.
    persistPanelState();
  });
  elements.maxChars?.addEventListener("input", () => {
    persistPanelState();
  });
  elements.maxWords?.addEventListener("input", () => {
    persistPanelState();
  });
  elements.linesPerCaption?.addEventListener("input", () => {
    persistPanelState();
  });
  elements.srtPath?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.correctedTranscriptPath?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.whisperModel?.addEventListener("change", () => {
    pendingWhisperModelValue = String(elements.whisperModel?.value || "").trim() || pendingWhisperModelValue;
    persistPanelState();
  });
  elements.whisperSequenceRange?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.whisperLanguage?.addEventListener("change", () => {
    pendingWhisperLanguageValue = getSelectedWhisperLanguageCode();
    toggleSourceFields();
    if (elements.generateButton && !generateInProgress) {
      elements.generateButton.disabled = !isCurrentSourceReady();
    }
    if (textEditorBlocks.length > 0) {
      renderTextEditor();
    }
    persistPanelState();
  });
  elements.preserveTranslationTiming?.addEventListener("change", () => {
    // // Remember the opt-in because timing locking performs a second local Whisper pass.
    persistPanelState();
  });
  elements.preserveMixedLanguages?.addEventListener("change", () => {
    toggleSourceFields();
    setGenerateButtonsBusy(generateInProgress);
    persistPanelState();
    scheduleWhisperGlossaryStoreSave();
  });
  elements.mixedLanguagePrompt?.addEventListener("input", () => {
    persistPanelState();
    scheduleWhisperGlossaryStoreSave();
  });
  elements.deeplApiKey?.addEventListener("input", () => {
    // // Persist the private key locally as it is typed so closing Premiere cannot lose it.
    persistPanelState();
  });
  elements.deeplApiKey?.addEventListener("change", () => {
    // // Refresh options after the key is complete, without sending it through Premiere's host bridge.
    void refreshDeepLSupportedLanguages().catch((error) => {
      setLog(getDeepLUserErrorMessage(error), true);
    });
  });
  window.addEventListener("beforeunload", () => {
    // // Flush the last keystrokes synchronously inside the bridge before CEP closes the panel.
    void writeWhisperGlossaryStore(
      Boolean(elements.preserveMixedLanguages?.checked),
      sanitizeWhisperGlossary(elements.mixedLanguagePrompt?.value || "")
    );
  });
  elements.logToggleButton?.addEventListener("click", () => {
    setLogPanelExpanded(!logPanelExpanded);
  });
  elements.logVerbosityButton?.addEventListener("click", () => {
    setVerboseLogsEnabled(!verboseLogsEnabled);
  });
  elements.appVersion?.addEventListener("click", () => {
    void openProductPage().catch((error) => {
      setLog(String(error), true);
    });
  });
  elements.deeplApiKeyLink?.addEventListener("click", (event) => {
    // // Keep the account page in the external browser instead of replacing the Premiere panel.
    event.preventDefault();
    void openDeepLApiKeyPage().catch((error) => {
      setLog(String(error), true);
    });
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
  document.addEventListener("visibilitychange", () => {
    // // A docked CEP panel can stay open while hidden behind another panel, so suspend its fallback host polling until it is visible again.
    void syncVisualSelectionWatcherState();
    if (!isVisualSelectionMonitoringAllowed()) {
      clearVisualSelectionAutoRefreshTimer();
      stopVisualSelectionPolling();
      return;
    }

    startVisualSelectionPolling();
    scheduleVisualSelectionAutoRefresh("visibility");
  });

  elements.generateButton?.addEventListener("click", async () => {
    try {
      persistPanelState();
      await generate();
    } catch (error) {
      setLog(String(error), true);
    }
  });
  elements.generateStopButton?.addEventListener("click", async () => {
    await stopCurrentGenerateJob();
  });

  elements.visualCopyButton?.addEventListener("click", async () => {
    try {
      if (visualReadInProgress || visualApplyInProgress) {
        return;
      }
      // // Refresh from the current Premiere selection first so copy never uses stale panel data.
      await loadVisualPropertiesFromSelection(false, true);
      copyLoadedVisualProperties();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.visualApplyButton?.addEventListener("click", async () => {
    try {
      if (visualReadInProgress || visualApplyInProgress) {
        return;
      }
      await applyVisualChangesToSelection();
      persistPanelState();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.textReadButton?.addEventListener("click", async () => {
    try {
      if (textReadInProgress || textApplyInProgress) {
        return;
      }
      await loadTextItemsFromSelection(true, true);
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.textApplyButton?.addEventListener("click", async () => {
    try {
      if (textReadInProgress || textApplyInProgress) {
        return;
      }
      await applyTextEditorChanges();
    } catch (error) {
      setLog(String(error), true);
    }
  });
  elements.translationReadButton?.addEventListener("click", async () => {
    try {
      await loadTranslationSelection();
    } catch (error) {
      setLog(String(error), true);
    }
  });
  elements.translationLanguagesRefreshButton?.addEventListener("click", () => {
    void refreshDeepLSupportedLanguages().catch((error) => {
      setLog(getDeepLUserErrorMessage(error), true);
    });
  });
  elements.translationInputMode?.addEventListener("change", () => {
    toggleTranslationInputMode();
    translationBlocks = [];
    translationSelectionSignature = "";
    translationSameTrack = false;
    translatedSubtitleTexts = [];
    renderTranslationPreview();
    setTranslationSelectionSummary(translate("translation.selectionDefault"));
    if (elements.translationDuplicateButton) {
      elements.translationDuplicateButton.disabled = true;
    }
  });
  elements.translationAutoLoadGeneratedNativeSrt?.addEventListener("change", () => {
    // // Persist the user's choice without changing the currently loaded translation source.
    persistPanelState();
  });
  elements.translationSrtBrowseButton?.addEventListener("click", async () => {
    try {
      const srtPath = await pickSrtPath();
      if (elements.translationSrtPath && srtPath) {
        elements.translationSrtPath.value = srtPath;
      }
    } catch (error) {
      setLog(String(error), true);
    }
  });
  elements.translationTranslateButton?.addEventListener("click", async () => {
    try {
      await translateLoadedSelection();
    } catch (error) {
      setLog(getDeepLUserErrorMessage(error), true);
    }
  });
  elements.translationDuplicateButton?.addEventListener("click", async () => {
    try {
      await duplicateTranslatedSelection();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  refreshVisualButtonsBusyState();
  setLog(translate("log.ready"));
  if (whisperGlossaryLoadError) {
    setLog(whisperGlossaryLoadError, true);
  }

  void initializeVisualSelectionWatcher();

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
  void refreshDeepLSupportedLanguages().catch((error) => {
    // // Keep the built-in fallback list usable if the stored key is invalid or DeepL is offline.
    setLog(getDeepLUserErrorMessage(error), true);
  });
}

initialize().catch((error) => {
  setLog(String(error), true);
});
