// // Drive the Sub Creator panel UI and connect it to subtitle generation logic.
import { buildCaptionPlan } from "../core/planner";
import { STYLE_PRESETS } from "../core/presets";
import { parseSrt } from "../core/srt";
import type { AnimationMode, CaptionBuildOptions, CaptionCue, HostApplyPayload, HostCaptionCue, MogrtTemplateItem } from "../core/types";
import {
  applyCaptionPlan,
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

const elements = {
  languageSelect: document.querySelector<HTMLSelectElement>("#languageSelect"),
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

function setLog(message: string, isError = false): void {
  // // Provide a single visible place for runtime status and error traces.
  if (!elements.logOutput) {
    return;
  }

  elements.logOutput.textContent = message;
  elements.logOutput.classList.toggle("log--error", isError);
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
    preview.textContent = translate("gallery.previewText");

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
    return {
      id: `host-cue-${index + 1}`,
      startSeconds: Number(cue.startSeconds),
      endSeconds: Number(cue.endSeconds),
      text: String(cue.text || "").trim(),
      words: []
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

  const defaultLanguage = navigator.language?.startsWith("fr") ? "fr" : "en";
  if (elements.languageSelect) {
    elements.languageSelect.value = defaultLanguage;
  }

  await loadLocale(defaultLanguage);
  renderPresetSelect();
  applyPresetDefaults(STYLE_PRESETS[0].id);
  toggleSourceFields();

  await loadMogrtCatalog();
  renderMogrtGallery();

  elements.languageSelect?.addEventListener("change", async () => {
    await loadLocale(elements.languageSelect?.value ?? "en");
    renderPresetSelect();
    renderMogrtGallery();
    applyPresetDefaults(elements.presetSelect?.value ?? STYLE_PRESETS[0].id);
  });

  elements.presetSelect?.addEventListener("change", () => {
    applyPresetDefaults(elements.presetSelect?.value ?? STYLE_PRESETS[0].id);
  });

  elements.sourceMode?.addEventListener("change", () => {
    toggleSourceFields();
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
