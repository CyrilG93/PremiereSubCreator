// // Drive the Sub Creator panel UI and connect it to subtitle generation logic.
import { buildCaptionPlan } from "../core/planner";
import { STYLE_PRESETS } from "../core/presets";
import { parseSrt } from "../core/srt";
import { transcribeActiveSequence } from "../core/transcription";
import type { AnimationMode, CaptionBuildOptions, HostApplyPayload } from "../core/types";
import { applyCaptionPlan, pingHost } from "./cepBridge";

type LocaleMap = Record<string, string>;

const elements = {
  languageSelect: document.querySelector<HTMLSelectElement>("#languageSelect"),
  sourceMode: document.querySelector<HTMLSelectElement>("#sourceMode"),
  srtInputField: document.querySelector<HTMLElement>("#srtInputField"),
  srtFile: document.querySelector<HTMLInputElement>("#srtFile"),
  presetSelect: document.querySelector<HTMLSelectElement>("#presetSelect"),
  animationMode: document.querySelector<HTMLSelectElement>("#animationMode"),
  maxChars: document.querySelector<HTMLInputElement>("#maxChars"),
  linesPerCaption: document.querySelector<HTMLInputElement>("#linesPerCaption"),
  fontSize: document.querySelector<HTMLInputElement>("#fontSize"),
  uppercase: document.querySelector<HTMLInputElement>("#uppercase"),
  mogrtPath: document.querySelector<HTMLInputElement>("#mogrtPath"),
  videoTrackIndex: document.querySelector<HTMLInputElement>("#videoTrackIndex"),
  audioTrackIndex: document.querySelector<HTMLInputElement>("#audioTrackIndex"),
  pingButton: document.querySelector<HTMLButtonElement>("#pingButton"),
  generateButton: document.querySelector<HTMLButtonElement>("#generateButton"),
  logOutput: document.querySelector<HTMLPreElement>("#logOutput")
};

let currentLocale: LocaleMap = {};

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

    const translated = translate(key);
    if (node.tagName === "OPTION") {
      node.textContent = translated;
      return;
    }

    node.textContent = translated;
  });
}

function renderPresetSelect(): void {
  // // Populate style presets from static configuration.
  if (!elements.presetSelect) {
    return;
  }

  elements.presetSelect.innerHTML = "";
  STYLE_PRESETS.forEach((preset) => {
    const option = document.createElement("option");
    option.value = preset.id;
    option.textContent = translate(preset.labelKey);
    elements.presetSelect?.appendChild(option);
  });
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

function toggleSourceFields(): void {
  // // Show file input only when SRT mode is selected.
  if (!elements.sourceMode || !elements.srtInputField) {
    return;
  }

  elements.srtInputField.style.display = elements.sourceMode.value === "srt" ? "grid" : "none";
}

async function readSrtFile(fileInput: HTMLInputElement): Promise<string> {
  // // Read the selected SRT text content from disk.
  const file = fileInput.files?.[0];
  if (!file) {
    throw new Error(translate("error.missingSrt"));
  }

  return file.text();
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
    !elements.mogrtPath ||
    !elements.videoTrackIndex ||
    !elements.audioTrackIndex
  ) {
    throw new Error("Panel bindings not initialized.");
  }

  return {
    sourceMode: elements.sourceMode.value as "srt" | "transcription",
    languageCode: elements.languageSelect.value,
    style: {
      presetId: elements.presetSelect.value,
      fontSize: Number(elements.fontSize.value),
      maxCharsPerLine: Number(elements.maxChars.value),
      animationMode: elements.animationMode.value as AnimationMode,
      uppercase: elements.uppercase.checked,
      linesPerCaption: Number(elements.linesPerCaption.value)
    },
    mogrtPath: elements.mogrtPath.value.trim(),
    videoTrackIndex: Number(elements.videoTrackIndex.value),
    audioTrackIndex: Number(elements.audioTrackIndex.value)
  };
}

async function generate(): Promise<void> {
  // // Build the caption plan from selected source and push to Premiere host.
  if (!elements.srtFile) {
    throw new Error("SRT input not initialized.");
  }

  const options = collectBuildOptions();
  setLog(translate("log.processing"));

  let cues;
  if (options.sourceMode === "srt") {
    const srtText = await readSrtFile(elements.srtFile);
    cues = parseSrt(srtText);
  } else {
    const transcription = await transcribeActiveSequence(options.languageCode);
    if (transcription.warning) {
      setLog(transcription.warning, true);
    }
    cues = transcription.cues;
  }

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

  elements.languageSelect?.addEventListener("change", async () => {
    await loadLocale(elements.languageSelect?.value ?? "en");
    renderPresetSelect();
    applyPresetDefaults(elements.presetSelect?.value ?? STYLE_PRESETS[0].id);
  });

  elements.presetSelect?.addEventListener("change", () => {
    applyPresetDefaults(elements.presetSelect?.value ?? STYLE_PRESETS[0].id);
  });

  elements.sourceMode?.addEventListener("change", toggleSourceFields);

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
