// // Drive the Sub Creator panel UI and connect it to subtitle generation logic.
import { buildCaptionPlan } from "../core/planner";
import { parseSrt } from "../core/srt";
import type { AnimationMode, CaptionBuildOptions, CaptionCue, HostApplyPayload, MogrtTemplateItem } from "../core/types";
import {
  applyCaptionPlan,
  applyVisualPropertiesToSelectedMogrts,
  getSelectedMogrtCount,
  readSystemFontCatalog,
  getWhisperRuntimeStatus,
  pickSrtPath,
  pickWhisperAudioPath,
  readSelectedMogrtVisualProperties,
  readTextFileFromHost,
  transcribeWithWhisper
} from "./cepBridge";
import type { SystemFontCatalog } from "./cepBridge";

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
  sourceMode: "srt" | "whisper_local";
  srtPath: string;
  whisperAudioPath: string;
  whisperModel: string;
  animationMode: AnimationMode;
  maxCharsPerLine: number;
  linesPerCaption: number;
  fontSize: number;
  mogrtAspectFilter: string;
  selectedMogrtId: string;
  visualLiveUpdate: boolean;
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
  whisperAudioPath: document.querySelector<HTMLInputElement>("#whisperAudioPath"),
  whisperBrowseButton: document.querySelector<HTMLButtonElement>("#whisperBrowseButton"),
  whisperModel: document.querySelector<HTMLSelectElement>("#whisperModel"),
  animationMode: document.querySelector<HTMLSelectElement>("#animationMode"),
  maxChars: document.querySelector<HTMLInputElement>("#maxChars"),
  linesPerCaption: document.querySelector<HTMLInputElement>("#linesPerCaption"),
  fontSize: document.querySelector<HTMLInputElement>("#fontSize"),
  mogrtAspectFilter: document.querySelector<HTMLSelectElement>("#mogrtAspectFilter"),
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
  logOutput: document.querySelector<HTMLPreElement>("#logOutput")
};

let currentLocale: LocaleMap = {};
let availableMogrts: MogrtTemplateItem[] = [];
let selectedMogrt: MogrtTemplateItem | null = null;
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
const visualOriginalValuesByPath = new Map<string, string>();
const visualOpenGroups = new Set<string>();
let visualLiveUpdateTimer: number | null = null;
let visualLiveUpdateQueued = false;
let visualLiveUpdateInFlight = false;
let visualApplyInProgress = false;
let visualLiveUpdateEnabled = false;
let systemFontCatalog: SystemFontCatalog = {
  available: false,
  source: "unavailable",
  details: "",
  families: [],
  stylesByFamily: {}
};

function assertDomBindings(): void {
  // // Guard against missing panel DOM ids during development/build changes.
  const missing = Object.entries(elements)
    .filter(([, value]) => !value)
    .map(([key]) => key);

  if (missing.length > 0) {
    throw new Error(`Missing DOM elements: ${missing.join(", ")}`);
  }
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
    !elements.whisperAudioPath ||
    !elements.whisperModel ||
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
    whisperAudioPath: elements.whisperAudioPath.value || "",
    whisperModel: elements.whisperModel.value || "base",
    animationMode: (elements.animationMode.value as AnimationMode) || "line",
    maxCharsPerLine: Number(elements.maxChars.value),
    linesPerCaption: Number(elements.linesPerCaption.value),
    fontSize: Number(elements.fontSize.value),
    mogrtAspectFilter: elements.mogrtAspectFilter.value || "all",
    selectedMogrtId: selectedMogrt?.id || "",
    visualLiveUpdate: visualLiveUpdateEnabled
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

  if (elements.whisperAudioPath && typeof snapshot.whisperAudioPath === "string") {
    elements.whisperAudioPath.value = snapshot.whisperAudioPath;
  }

  if (elements.whisperModel && snapshot.whisperModel && hasSelectOption(elements.whisperModel, snapshot.whisperModel)) {
    elements.whisperModel.value = snapshot.whisperModel;
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

  if (elements.mogrtAspectFilter && snapshot.mogrtAspectFilter && hasSelectOption(elements.mogrtAspectFilter, snapshot.mogrtAspectFilter)) {
    elements.mogrtAspectFilter.value = snapshot.mogrtAspectFilter;
  }

  if (typeof snapshot.visualLiveUpdate === "boolean") {
    setVisualLiveUpdateEnabled(snapshot.visualLiveUpdate, true);
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

function panelAssetPath(relativeOrAbsolute: string): string {
  // // Normalize extension-local asset paths for fetch/image usage.
  if (!relativeOrAbsolute) {
    return "";
  }

  if (/^(https?:)?\/\//i.test(relativeOrAbsolute) || relativeOrAbsolute.startsWith("./")) {
    return relativeOrAbsolute;
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

function setLog(message: string, isError = false): void {
  // // Provide a single visible place for runtime status and error traces.
  if (!elements.logOutput) {
    return;
  }

  elements.logOutput.textContent = message;
  elements.logOutput.classList.toggle("log--error", isError);
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

  refreshUpdateBanner();
}

function getSourceMode(): "srt" | "whisper_local" {
  // // Normalize source mode value from UI select control.
  return (elements.sourceMode?.value as "srt" | "whisper_local") || "srt";
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

  if (elements.srtInputField) {
    elements.srtInputField.style.display = mode === "srt" ? "grid" : "none";
  }

  if (elements.whisperField) {
    elements.whisperField.style.display = mode === "whisper_local" ? "grid" : "none";
  }
}

async function enforceWhisperSourceAvailability(): Promise<void> {
  // // Hide Whisper source option when local runtime is unavailable on this machine.
  if (!elements.sourceMode) {
    return;
  }

  const whisperOption = elements.sourceMode.querySelector<HTMLOptionElement>('option[value="whisper_local"]');
  if (!whisperOption) {
    return;
  }

  try {
    const status = await getWhisperRuntimeStatus();
    if (status.available) {
      return;
    }

    whisperOption.remove();
    if (elements.sourceMode.value === "whisper_local") {
      elements.sourceMode.value = "srt";
    }

    if (elements.whisperAudioPath) {
      elements.whisperAudioPath.value = "";
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
        stylesByFamily: {}
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
          : {}
    };
  } catch {
    systemFontCatalog = {
      available: false,
      source: "error",
      details: "",
      families: [],
      stylesByFamily: {}
    };
  }
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
  loadedVisualProperties = properties.slice();
  visualOriginalValuesByPath.clear();

  if (!properties.length) {
    return;
  }

  const textStyleFamilySelectByBasePath = new Map<string, HTMLSelectElement>();
  const textStyleStyleSelectByBasePath = new Map<string, HTMLSelectElement>();
  const textStyleStylesByFamilyByBasePath = new Map<string, Record<string, string[]>>();
  const textStyleFlagCheckboxesByBasePath = new Map<string, { allCaps?: HTMLInputElement; smallCaps?: HTMLInputElement }>();

  const parseTextStyleVirtualPath = (path: string): { basePath: string; styleKey: string } | null => {
    // // Decode virtual text-style control paths emitted by host (`4::textstyle.fontStyle`).
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
  };

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
        normalized[family.toLowerCase()] = uniqueStyles;
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

  const normalizedSystemStyleMap = systemFontCatalog.available
    ? normalizeStyleMap(systemFontCatalog.stylesByFamily)
    : {};
  const systemFamilies = systemFontCatalog.available && Array.isArray(systemFontCatalog.families)
    ? systemFontCatalog.families
    : [];

  const replaceSelectOptions = (select: HTMLSelectElement, options: string[], preferredValue: string): void => {
    // // Replace select items while preserving currently selected value when possible.
    const deduped: string[] = [];
    for (const optionText of options) {
      const normalized = String(optionText || "").trim();
      if (!normalized) {
        continue;
      }
      if (!deduped.some((item) => item.toLowerCase() === normalized.toLowerCase())) {
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
    for (const option of Array.from(select.options)) {
      if (String(option.value).toLowerCase() === targetValue.toLowerCase()) {
        select.value = option.value;
        return;
      }
    }
    if (!select.value && select.options.length > 0) {
      select.selectedIndex = 0;
    }
  };

  const refreshStyleSelectForFamily = (basePath: string): void => {
    // // Keep font-style options aligned with currently selected font family when map is available.
    const familySelect = textStyleFamilySelectByBasePath.get(basePath);
    const styleSelect = textStyleStyleSelectByBasePath.get(basePath);
    const styleMap = textStyleStylesByFamilyByBasePath.get(basePath);
    if (!familySelect || !styleSelect || !styleMap) {
      return;
    }

    const selectedFamily = String(familySelect.value || "").trim().toLowerCase();
    if (!selectedFamily) {
      return;
    }

    const mappedOptions = styleMap[selectedFamily];
    if (!Array.isArray(mappedOptions) || mappedOptions.length === 0) {
      return;
    }

    const currentStyle = String(styleSelect.value || "").trim();
    replaceSelectOptions(styleSelect, mappedOptions, currentStyle);
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

  const ensureSelectViewportSpace = (select: HTMLSelectElement): void => {
    // // Pre-scroll the visual panel before opening long font dropdown lists.
    const rect = select.getBoundingClientRect();
    const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
    const rowHeight = 28;
    const desiredRows = Math.min(Math.max(6, select.options.length > 0 ? Math.min(select.options.length, 10) : 6), 10);
    const desiredHeight = desiredRows * rowHeight + 12;
    const initialBelow = Math.max(0, viewportHeight - rect.bottom - 8);
    let refreshedBelow = initialBelow;
    const missingSpace = desiredHeight - initialBelow;
    if (missingSpace > 0) {
      const scrollable = findScrollableAncestor(select);
      if (scrollable) {
        const maxScrollable = Math.max(0, scrollable.scrollHeight - scrollable.clientHeight - scrollable.scrollTop);
        if (maxScrollable > 0) {
          const delta = Math.min(missingSpace, maxScrollable);
          if (delta > 0) {
            scrollable.scrollTop += delta;
          }
          const refreshedRect = select.getBoundingClientRect();
          refreshedBelow = Math.max(0, viewportHeight - refreshedRect.bottom - 8);
        }
      }
    }

    if (refreshedBelow < 120) {
      // // Fallback to inline expanded select when native popup still lacks vertical room.
      const fallbackRows = Math.max(4, Math.min(select.options.length || 6, Math.floor(Math.max(refreshedBelow, 120) / rowHeight)));
      if (fallbackRows > 1) {
        select.size = fallbackRows;
        select.classList.add("visual-select-expanded");
        const collapse = (): void => {
          select.size = 1;
          select.classList.remove("visual-select-expanded");
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
      }
    }
  };

  const grouped = new Map<string, HostVisualProperty[]>();
  for (const property of properties) {
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
        if (textStylePath && (textStylePath.styleKey === "fontFsAllCaps" || textStylePath.styleKey === "fontFsSmallCaps")) {
          // // Mirror Premiere mutual exclusivity between All Caps and Small Caps in editor state.
          const existing = textStyleFlagCheckboxesByBasePath.get(textStylePath.basePath) || {};
          if (textStylePath.styleKey === "fontFsAllCaps") {
            existing.allCaps = checkbox;
          } else {
            existing.smallCaps = checkbox;
          }
          textStyleFlagCheckboxesByBasePath.set(textStylePath.basePath, existing);
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
        }
        if (textStylePath?.styleKey === "fontFamily") {
          const currentFamilyValue = String(select.value || currentValue || "").trim();
          const selectFamilies = Array.from(select.options).map((option) => String(option.value || ""));
          replaceSelectOptions(select, [...selectFamilies, ...systemFamilies], currentFamilyValue);
          select.addEventListener("mousedown", () => {
            ensureSelectViewportSpace(select);
          });
          select.addEventListener("touchstart", () => {
            ensureSelectViewportSpace(select);
          });
          select.addEventListener("keydown", (event) => {
            if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
              ensureSelectViewportSpace(select);
            }
          });
          textStyleFamilySelectByBasePath.set(textStylePath.basePath, select);
          select.addEventListener("change", () => {
            refreshStyleSelectForFamily(textStylePath.basePath);
            scheduleLiveVisualApply();
          });
        } else if (textStylePath?.styleKey === "fontStyle") {
          textStyleStyleSelectByBasePath.set(textStylePath.basePath, select);
          const relatedMap = textStyleStylesByFamilyByBasePath.get(textStylePath.basePath);
          if (relatedMap && Object.keys(relatedMap).length > 0) {
            const selectStyles = Array.from(select.options).map((option) => String(option.value || ""));
            const relatedFamilySelect = textStyleFamilySelectByBasePath.get(textStylePath.basePath);
            const selectedFamilyKey = String(relatedFamilySelect?.value || "").trim().toLowerCase();
            const stylesFromMap =
              (selectedFamilyKey ? relatedMap[selectedFamilyKey] : undefined) ||
              Object.values(relatedMap).flat();
            replaceSelectOptions(select, [...selectStyles, ...stylesFromMap], String(select.value || currentValue || ""));
          }
          select.addEventListener("mousedown", () => {
            ensureSelectViewportSpace(select);
          });
          select.addEventListener("touchstart", () => {
            ensureSelectViewportSpace(select);
          });
          select.addEventListener("keydown", (event) => {
            if (event.key === "ArrowDown" || event.key === "Enter" || event.key === " ") {
              ensureSelectViewportSpace(select);
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
    refreshStyleSelectForFamily(basePath);
  }
}

type VisualPropertyChange = {
  path: string;
  valueType: HostVisualProperty["valueType"];
  controlKind: HostVisualProperty["controlKind"];
  vectorScale?: number[];
  value: string | number | boolean;
};

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

    changes.push({
      path,
      valueType,
      controlKind,
      vectorScale,
      value
    });
  });

  return changes;
}

async function loadVisualPropertiesFromSelection(emitHostLog = false): Promise<void> {
  // // Read selected MOGRT editable controls from host and refresh visual editor UI.
  const result = await readSelectedMogrtVisualProperties();
  renderVisualPropertyEditor(result.properties);
  if (emitHostLog) {
    setLog(`${translate("log.hostResult")}\n${JSON.stringify(result, null, 2)}`);
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

      setLog(
        `${translate("log.visualApplyDone")}\n${JSON.stringify(
          {
            selectedCount,
            processedClipCount: selectedCount,
            updatedCount,
            failedCount,
            debug: debugLines
          },
          null,
          2
        )}`
      );
      setVisualApplyProgressState(false);
      await loadVisualPropertiesFromSelection();
      return;
    }

    const response = await applyVisualPropertiesToSelectedMogrts(changes);
    if (!useLiveUpdate) {
      setLog(`${translate("log.visualApplyDone")}\n${JSON.stringify(response, null, 2)}`);
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

async function loadMogrtCatalog(): Promise<void> {
  // // Load generated MOGRT catalog emitted by the build script.
  const response = await fetch("./assets/mogrt-catalog.json");
  if (!response.ok) {
    throw new Error(translate("error.mogrtCatalogMissing"));
  }

  const catalog = (await response.json()) as MogrtCatalog;
  availableMogrts = Array.isArray(catalog.templates) ? catalog.templates : [];

  if (!selectedMogrt && availableMogrts.length > 0) {
    selectedMogrt = availableMogrts[0];
  }
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

function selectMogrt(templateId: string): void {
  // // Save selected template and rerender cards to reflect active state.
  const found = availableMogrts.find((template) => template.id === templateId);
  if (!found) {
    return;
  }

  selectedMogrt = found;
  renderMogrtGallery();
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
      previewVideo.autoplay = true;
      previewVideo.playsInline = true;
      previewVideo.preload = "metadata";
      previewVideo.setAttribute("aria-hidden", "true");
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

async function browseWhisperAudio(): Promise<void> {
  // // Pick local media file from host picker and populate Whisper audio path.
  if (!elements.whisperAudioPath) {
    return;
  }

  const path = await pickWhisperAudioPath();
  if (path) {
    elements.whisperAudioPath.value = path;
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
    !elements.whisperAudioPath ||
    !elements.whisperModel
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
    whisperAudioPath: elements.whisperAudioPath.value.trim(),
    whisperModel: elements.whisperModel.value,
    videoTrackIndex: 0,
    audioTrackIndex: 0
  };
}

async function loadCuesFromSelectedSource(options: CaptionBuildOptions): Promise<CaptionCue[]> {
  // // Build cues from the currently selected source mode.
  if (options.sourceMode === "srt") {
    if (!elements.srtPath || !elements.srtPath.value.trim()) {
      throw new Error(translate("error.missingSrtPath"));
    }

    const srtText = await readTextFileFromHost(elements.srtPath.value.trim());
    const cues = parseSrt(srtText);
    if (!cues.length) {
      throw new Error(translate("error.emptySrt"));
    }

    return cues;
  }

  if (!options.whisperAudioPath) {
    throw new Error(translate("error.missingWhisperAudio"));
  }

  const whisperResult = await transcribeWithWhisper({
    audioPath: options.whisperAudioPath,
    languageCode: options.languageCode,
    model: options.whisperModel
  });

  const cues = parseSrt(whisperResult.srtText);
  if (!cues.length) {
    throw new Error(translate("error.emptyWhisper"));
  }

  setLog(`${translate("log.whisperDone")} ${whisperResult.model}`);
  return cues;
}

async function generate(): Promise<void> {
  // // Build the caption plan from selected source and push to Premiere host.
  const options = collectBuildOptions();
  setLog(translate("log.processing"));

  const cues = await loadCuesFromSelectedSource(options);
  const plannedCues = buildCaptionPlan(cues, options);

  const payload: HostApplyPayload = {
    options,
    cues: plannedCues
  };

  const hostResultRaw = await applyCaptionPlan(payload);
  setLog(`${translate("log.hostResult")}\n${hostResultRaw}`);
}

async function initialize(): Promise<void> {
  // // Initialize locale, controls, and event listeners once panel is loaded.
  assertDomBindings();
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
  await enforceWhisperSourceAvailability();
  await loadSystemFontCatalogFallback();
  setVisualLiveUpdateEnabled(false, true);
  applyPersistedPanelState(persistedState);
  setActiveMode(activeMode);
  toggleSourceFields();

  await loadMogrtCatalog();
  if (pendingSelectedMogrtId) {
    const restoredTemplate = availableMogrts.find((template) => template.id === pendingSelectedMogrtId);
    if (restoredTemplate) {
      selectedMogrt = restoredTemplate;
    }
    pendingSelectedMogrtId = "";
  }
  renderMogrtGallery();
  persistPanelState();
  await checkForUpdates();

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

  elements.whisperBrowseButton?.addEventListener("click", async () => {
    try {
      await browseWhisperAudio();
    } catch (error) {
      setLog(String(error), true);
    }
  });

  elements.mogrtAspectFilter?.addEventListener("change", () => {
    renderMogrtGallery();
    persistPanelState();
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
  elements.whisperAudioPath?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.whisperModel?.addEventListener("change", () => {
    persistPanelState();
  });
  elements.visualLiveUpdateButton?.addEventListener("click", () => {
    setVisualLiveUpdateEnabled(!visualLiveUpdateEnabled);
    if (visualLiveUpdateEnabled) {
      scheduleLiveVisualApply();
    }
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
}

initialize().catch((error) => {
  setLog(String(error), true);
});
