// // Drive the Sub Creator panel UI and connect it to subtitle generation logic.
import { buildCaptionPlan } from "../core/planner";
import { STYLE_PRESETS } from "../core/presets";
import { parseSrt } from "../core/srt";
import type { AnimationMode, CaptionBuildOptions, CaptionCue, HostApplyPayload, HostCaptionCue, MogrtTemplateItem } from "../core/types";
import {
  applyCaptionPlan,
  getWhisperRuntimeStatus,
  pickSrtPath,
  pickWhisperAudioPath,
  pingHost,
  readActiveCaptionTrackCues,
  readTextFileFromHost,
  transcribeWithWhisper
} from "./cepBridge";

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

interface PanelStateSnapshot {
  languageCode: string;
  sourceMode: "srt" | "premiere_caption" | "whisper_local";
  srtPath: string;
  whisperAudioPath: string;
  whisperModel: string;
  presetId: string;
  animationMode: AnimationMode;
  maxCharsPerLine: number;
  linesPerCaption: number;
  fontSize: number;
  uppercase: boolean;
  mogrtAspectFilter: string;
  selectedMogrtId: string;
}

const elements = {
  languageSelect: document.querySelector<HTMLSelectElement>("#languageSelect"),
  appVersion: document.querySelector<HTMLSpanElement>("#appVersion"),
  updateBanner: document.querySelector<HTMLElement>("#updateBanner"),
  updateLink: document.querySelector<HTMLAnchorElement>("#updateLink"),
  sourceMode: document.querySelector<HTMLSelectElement>("#sourceMode"),
  srtInputField: document.querySelector<HTMLElement>("#srtInputField"),
  srtPath: document.querySelector<HTMLInputElement>("#srtPath"),
  srtBrowseButton: document.querySelector<HTMLButtonElement>("#srtBrowseButton"),
  premiereCaptionField: document.querySelector<HTMLElement>("#premiereCaptionField"),
  whisperField: document.querySelector<HTMLElement>("#whisperField"),
  whisperAudioPath: document.querySelector<HTMLInputElement>("#whisperAudioPath"),
  whisperBrowseButton: document.querySelector<HTMLButtonElement>("#whisperBrowseButton"),
  whisperModel: document.querySelector<HTMLSelectElement>("#whisperModel"),
  presetSelect: document.querySelector<HTMLSelectElement>("#presetSelect"),
  animationMode: document.querySelector<HTMLSelectElement>("#animationMode"),
  maxChars: document.querySelector<HTMLInputElement>("#maxChars"),
  linesPerCaption: document.querySelector<HTMLInputElement>("#linesPerCaption"),
  fontSize: document.querySelector<HTMLInputElement>("#fontSize"),
  uppercase: document.querySelector<HTMLInputElement>("#uppercase"),
  mogrtAspectFilter: document.querySelector<HTMLSelectElement>("#mogrtAspectFilter"),
  mogrtGallery: document.querySelector<HTMLElement>("#mogrtGallery"),
  mogrtSelectedLabel: document.querySelector<HTMLParagraphElement>("#mogrtSelectedLabel"),
  pingButton: document.querySelector<HTMLButtonElement>("#pingButton"),
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
    !elements.sourceMode ||
    !elements.srtPath ||
    !elements.whisperAudioPath ||
    !elements.whisperModel ||
    !elements.presetSelect ||
    !elements.animationMode ||
    !elements.maxChars ||
    !elements.linesPerCaption ||
    !elements.fontSize ||
    !elements.uppercase ||
    !elements.mogrtAspectFilter
  ) {
    return;
  }

  const snapshot: PanelStateSnapshot = {
    languageCode: elements.languageSelect.value || "en",
    sourceMode: getSourceMode(),
    srtPath: elements.srtPath.value || "",
    whisperAudioPath: elements.whisperAudioPath.value || "",
    whisperModel: elements.whisperModel.value || "base",
    presetId: elements.presetSelect.value || STYLE_PRESETS[0].id,
    animationMode: (elements.animationMode.value as AnimationMode) || "line",
    maxCharsPerLine: Number(elements.maxChars.value),
    linesPerCaption: Number(elements.linesPerCaption.value),
    fontSize: Number(elements.fontSize.value),
    uppercase: elements.uppercase.checked,
    mogrtAspectFilter: elements.mogrtAspectFilter.value || "all",
    selectedMogrtId: selectedMogrt?.id || ""
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

  if (elements.presetSelect && snapshot.presetId && hasSelectOption(elements.presetSelect, snapshot.presetId)) {
    elements.presetSelect.value = snapshot.presetId;
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

  if (elements.uppercase && typeof snapshot.uppercase === "boolean") {
    elements.uppercase.checked = snapshot.uppercase;
  }

  if (elements.mogrtAspectFilter && snapshot.mogrtAspectFilter && hasSelectOption(elements.mogrtAspectFilter, snapshot.mogrtAspectFilter)) {
    elements.mogrtAspectFilter.value = snapshot.mogrtAspectFilter;
  }

  if (typeof snapshot.selectedMogrtId === "string" && snapshot.selectedMogrtId.length > 0) {
    pendingSelectedMogrtId = snapshot.selectedMogrtId;
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

  refreshUpdateBanner();
}

function renderPresetSelect(): void {
  // // Populate style presets from static configuration.
  if (!elements.presetSelect) {
    return;
  }

  const selectedId = elements.presetSelect.value;
  elements.presetSelect.innerHTML = "";

  STYLE_PRESETS.forEach((preset) => {
    const option = document.createElement("option");
    option.value = preset.id;
    option.textContent = translate(preset.labelKey);
    elements.presetSelect?.appendChild(option);
  });

  elements.presetSelect.value = selectedId || STYLE_PRESETS[0].id;
}

function applyPresetDefaults(presetId: string): void {
  // // Sync numeric controls whenever user chooses a different style preset.
  const preset = STYLE_PRESETS.find((item) => item.id === presetId);
  if (!preset || !elements.fontSize || !elements.maxChars || !elements.animationMode) {
    return;
  }

  elements.fontSize.value = String(preset.defaultFontSize);
  elements.maxChars.value = String(preset.defaultMaxCharsPerLine);
  elements.animationMode.value = preset.defaultAnimationMode;
}

function getSourceMode(): "srt" | "premiere_caption" | "whisper_local" {
  // // Normalize source mode value from UI select control.
  return (elements.sourceMode?.value as "srt" | "premiere_caption" | "whisper_local") || "srt";
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

  if (elements.premiereCaptionField) {
    elements.premiereCaptionField.style.display = mode === "premiere_caption" ? "grid" : "none";
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
    !elements.presetSelect ||
    !elements.fontSize ||
    !elements.maxChars ||
    !elements.linesPerCaption ||
    !elements.animationMode ||
    !elements.uppercase ||
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
      presetId: elements.presetSelect.value,
      fontSize: Number(elements.fontSize.value),
      maxCharsPerLine: Number(elements.maxChars.value),
      animationMode: elements.animationMode.value as AnimationMode,
      uppercase: elements.uppercase.checked,
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

function normalizeHostCaptionCues(hostCues: HostCaptionCue[]): CaptionCue[] {
  // // Convert host caption cue payload into planner-compatible cue objects.
  return hostCues.map((cue, index) => {
    const startSeconds = Number(cue.startSeconds);
    const endSeconds = Number(cue.endSeconds);
    const text = String(cue.text || "").trim();
    const words = text
      .replace(/\r/g, "\n")
      .replace(/\n+/g, " ")
      .split(/\s+/)
      .map((value) => value.trim())
      .filter(Boolean);
    const totalDuration = Math.max(endSeconds - startSeconds, 0.01);
    const wordDuration = words.length > 0 ? totalDuration / words.length : 0;

    return {
      id: `host-cue-${index + 1}`,
      startSeconds,
      endSeconds,
      text,
      words: words.map((word, wordIndex) => {
        const wordStart = startSeconds + wordDuration * wordIndex;
        const wordEnd = wordIndex === words.length - 1 ? endSeconds : wordStart + wordDuration;
        return {
          text: word,
          startSeconds: wordStart,
          endSeconds: wordEnd
        };
      })
    };
  });
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

  if (options.sourceMode === "premiere_caption") {
    const hostCues = await readActiveCaptionTrackCues();
    const cues = normalizeHostCaptionCues(hostCues).filter((cue) => cue.text.length > 0 && cue.endSeconds > cue.startSeconds);
    if (!cues.length) {
      throw new Error(translate("error.noActiveCaptionTrack"));
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
  renderPresetSelect();
  applyPresetDefaults(elements.presetSelect?.value ?? STYLE_PRESETS[0].id);
  applyPersistedPanelState(persistedState);
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
    renderPresetSelect();
    renderMogrtGallery();
    persistPanelState();
  });

  elements.presetSelect?.addEventListener("change", () => {
    applyPresetDefaults(elements.presetSelect?.value ?? STYLE_PRESETS[0].id);
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
  elements.uppercase?.addEventListener("change", () => {
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

  elements.pingButton?.addEventListener("click", async () => {
    try {
      const result = await pingHost();
      setLog(`${translate("log.pingOk")}\n${result}`);
    } catch (error) {
      setLog(String(error), true);
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

  setLog(translate("log.ready"));
}

initialize().catch((error) => {
  setLog(String(error), true);
});
