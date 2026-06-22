// // Wrap CEP evalScript calls and provide a browser fallback for local testing.
import type {
  CaptionCue,
  CaptionBuildOptions,
  HostApplyPayload,
  MogrtTemplateItem,
  PremiereTemplateTextPayload,
  WhisperSequenceRangeMode
} from "../core/types";
import { gunzipSync, gzipSync, strFromU8, strToU8, unzipSync, zipSync } from "fflate";

declare global {
  interface Window {
    __adobe_cep__?: {
      evalScript: (script: string, callback: (result: string) => void) => void;
      getHostEnvironment?: () => string;
      addEventListener?: (eventName: string, listener: (event?: unknown) => void) => void;
      removeEventListener?: (eventName: string, listener: (event?: unknown) => void) => void;
    };
    cep?: {
      util?: {
        openURLInDefaultBrowser?: (url: string) => void;
      };
    };
    require?: (moduleName: string) => unknown;
    cep_node?: {
      require?: (moduleName: string) => unknown;
    };
  }
}

interface HostJsonResponse<T> {
  ok: boolean;
  error?: string;
  data?: T;
  [key: string]: unknown;
}

interface WhisperTranscriptionRequest {
  audioPath: string;
  languageCode: string;
  model: string;
}

interface CorrectedAlignmentRequest {
  audioPath: string;
  transcriptPath: string;
  languageCode: string;
  extensionRootPath: string;
  rangeStartSeconds?: number;
  rangeEndSeconds?: number;
}

interface WhisperXTranscriptionRequest extends WhisperTranscriptionRequest {
  extensionRootPath: string;
}

interface WhisperTranscriptionResult {
  srtText: string;
  jsonText?: string;
  model: string;
  audioPath: string;
  commandOutput?: string;
}

interface CorrectedAlignmentResult {
  jsonText: string;
  audioPath: string;
  transcriptPath: string;
  commandOutput?: string;
}

interface CepSpawnedChild {
  stdout?: { on: (eventName: "data", listener: (chunk: string | Uint8Array) => void) => void };
  stderr?: { on: (eventName: "data", listener: (chunk: string | Uint8Array) => void) => void };
  on: (eventName: "error" | "close", listener: (value: unknown) => void) => void;
  kill?: (signal?: string) => void;
  pid?: number;
}

export interface WhisperProgressUpdate {
  percent: number;
  detail: string;
  remaining?: string;
}

interface WhisperSequenceExportResult {
  audioPath: string;
  presetPath: string;
  exportMethod?: "premiere_direct" | "media_encoder";
  sequenceName?: string;
  rangeStartSeconds?: number;
  rangeEndSeconds?: number;
  audioDurationSeconds?: number;
  debug?: unknown;
}

interface ActiveSequenceRangeResult {
  rangeStartSeconds?: number;
  rangeEndSeconds?: number;
  sequenceName?: string;
  fallbackReason?: string;
  hostError?: string;
  debug?: unknown;
}

export interface WhisperRuntimeStatus {
  available: boolean;
  details: string;
  installedModels: string[];
  modelCachePaths: string[];
  alignmentAvailable: boolean;
  alignmentDetails: string;
}

export interface SystemFontCatalog {
  available: boolean;
  source: string;
  details: string;
  families: string[];
  stylesByFamily: Record<string, string[]>;
  fontTokensByFamilyStyle: Record<string, Record<string, string>>;
}

export interface InstalledMogrtCatalog {
  available: boolean;
  source: string;
  details: string;
  templatesRoot: string;
  groups: string[];
  templates: MogrtTemplateItem[];
}

export interface SelectedMogrtVisualProperty {
  path: string;
  displayName: string;
  groupPath: string;
  valueType: "number" | "boolean" | "string" | "json";
  controlKind: "slider" | "number" | "checkbox" | "color" | "text" | "string" | "json" | "vector" | "select";
  cloneOnlyWhenDirty?: boolean;
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

export interface SelectedMogrtVisualComponentDebug {
  index: number;
  name: string;
  propertyCount: number;
}

export interface SelectedMogrtVisualDebug {
  sequenceWidth?: number;
  sequenceHeight?: number;
  componentCount?: number;
  components?: SelectedMogrtVisualComponentDebug[];
  vectorCount?: number;
  colorCount?: number;
  selectCount?: number;
  sample?: unknown[];
  textStyleCandidates?: unknown[];
  [key: string]: unknown;
}

export interface SelectedMogrtVisualPropertyList {
  selectedCount: number;
  editableCount: number;
  properties: SelectedMogrtVisualProperty[];
  debug?: SelectedMogrtVisualDebug;
}

export interface ApplyVisualPropertiesResult {
  selectedCount: number;
  processedClipCount: number;
  clipStartIndex?: number;
  clipEndIndex?: number;
  updatedCount: number;
  failedCount: number;
  debug?: string[];
}

export interface SelectedMogrtTextItem {
  selectionIndex: number;
  videoTrackIndex: number;
  startSeconds: number;
  endSeconds: number;
  text: string;
  clipName: string;
}

export interface SelectedMogrtTextItemList {
  selectedCount: number;
  sameTrack: boolean;
  videoTrackIndex?: number;
  projectDocumentId?: string;
  projectPath?: string;
  sequenceID?: string;
  sequenceName?: string;
  signature: string;
  items: SelectedMogrtTextItem[];
}

export interface TextEditorApplyItemPayload {
  sourceSelectionIndex: number;
  startSeconds: number;
  endSeconds: number;
  text: string;
  mogrtPathOverride?: string;
  skipTextApply?: boolean;
}

export interface TextEditorApplyPayload {
  selectionSignature: string;
  replaceSelectionStartIndex: number;
  replaceSelectionEndIndex: number;
  items: TextEditorApplyItemPayload[];
  options: CaptionBuildOptions;
}

export interface ApplySelectedMogrtTextResult {
  selectedCount: number;
  rebuiltCount: number;
  failedCount: number;
  selectionSignature?: string;
  sourceTrackIndex?: number;
  rebuildTrackIndex?: number;
  projectDocumentId?: string;
  projectPath?: string;
  sequenceID?: string;
  sequenceName?: string;
  debug?: string[];
}

interface CepNodeModules {
  childProcess: {
    spawn?: (
      command: string,
      args: string[],
      options: {
        shell?: boolean;
        detached?: boolean;
        env?: Record<string, string | undefined>;
      }
    ) => CepSpawnedChild;
    spawnSync: (
      command: string,
      args: string[],
      options: {
        encoding: string;
        shell?: boolean;
        timeout?: number;
        maxBuffer?: number;
        env?: Record<string, string | undefined>;
      }
    ) => { status: number | null; stdout?: string; stderr?: string; error?: { message?: string; code?: string } };
  };
  fs: {
    existsSync: (path: string) => boolean;
    mkdirSync: (path: string, options: { recursive: boolean }) => void;
    readdirSync: (path: string) => string[];
    readFileSync: (path: string, encoding?: string) => string | Uint8Array;
    openSync?: (path: string, flags: string) => number;
    readSync?: (fd: number, buffer: Uint8Array, offset: number, length: number, position: number) => number;
    closeSync?: (fd: number) => void;
    statSync: (path: string) => {
      size?: number;
      mtimeMs?: number;
      isDirectory?: () => boolean;
      isFile?: () => boolean;
    };
    writeFileSync: (path: string, data: Uint8Array) => void;
    unlinkSync: (path: string) => void;
  };
  os: {
    tmpdir: () => string;
    homedir: () => string;
  };
  path: {
    join: (...parts: string[]) => string;
    basename: (value: string) => string;
    dirname: (value: string) => string;
  };
  process: {
    env: Record<string, string | undefined>;
    kill?: (pid: number, signal?: string | number) => boolean;
  };
}

interface WhisperCommandCandidate {
  command: string;
  args: string[];
  label: string;
}

interface WhisperModelDefinition {
  value: string;
  filenames: string[];
}

interface PythonLauncherCandidate {
  command: string;
  argsPrefix: string[];
  label: string;
}

interface SubcreatorRuntimeConfig {
  sourcePath: string;
  pythonCommand: string;
  pythonPath: string;
  pythonVersion: string;
  whisperPath: string;
  ffmpegPath: string;
  pathHints: string[];
}

interface ActiveCepJob {
  label: string;
  child: CepSpawnedChild;
  detached: boolean;
  cancelRequested: boolean;
}

let subcreatorRuntimeConfigCache: SubcreatorRuntimeConfig | null | undefined;
let subcreatorSystemFontCatalogCache: SystemFontCatalog | null | undefined;
let subcreatorInstalledMogrtCatalogCache:
  | {
      templatesRoot: string;
      signature: string;
      catalog: InstalledMogrtCatalog;
    }
  | undefined;
let activeCepJob: ActiveCepJob | null = null;
const SUBCREATOR_CANCELLED_JOB_CODE = "SUBCREATOR_JOB_CANCELLED";
const SUBCREATOR_SUPPORTED_WHISPER_MODELS: WhisperModelDefinition[] = [
  { value: "tiny", filenames: ["tiny.pt"] },
  { value: "base", filenames: ["base.pt"] },
  { value: "small", filenames: ["small.pt"] },
  { value: "medium", filenames: ["medium.pt"] },
  { value: "large-v3", filenames: ["large-v3.pt"] },
  { value: "turbo", filenames: ["turbo.pt", "large-v3-turbo.pt"] }
];

function escapeForJsx(input: string): string {
  // // Escape special characters before embedding text into evalScript call strings.
  return input
    .replace(/\\/g, "\\\\")
    .replace(/"/g, "\\\"")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "\\r");
}

function evalScript(script: string): Promise<string> {
  // // Route script execution through Premiere CEP host when available.
  if (window.__adobe_cep__) {
    return new Promise((resolve) => {
      window.__adobe_cep__?.evalScript(script, (result) => resolve(result));
    });
  }

  return Promise.resolve(
    JSON.stringify({
      ok: true,
      mocked: true,
      message: "CEP host unavailable, running in browser fallback mode."
    })
  );
}

function getHostFunctionName(script: string): string {
  // // Extract the host function name so guarded evalScript errors point to the failing Premiere bridge call.
  const match = String(script || "").trim().match(/^([A-Za-z_$][\w$]*)\s*\(/);
  return match ? match[1] : "unknown host function";
}

function previewHostResponse(value: unknown, maxLength = 800): string {
  // // Keep malformed host responses shareable in logs without flooding the CEP panel.
  const text = String(value ?? "");
  if (text.length <= maxLength) {
    return text;
  }
  return `${text.slice(0, maxLength)}...`;
}

function buildGuardedHostJsonScript(script: string): string {
  // // Wrap JSON-returning ExtendScript calls so missing host functions return actionable JSON instead of raw "EvalScript error".
  const hostFunctionName = getHostFunctionName(script);
  const missingError = JSON.stringify(
    `Premiere host function is missing: ${hostFunctionName}. Restart Premiere Pro, then reinstall Sub Creator with Premiere closed.`
  );
  const exceptionPrefix = JSON.stringify(`Premiere host error in ${hostFunctionName}: `);
  const quotedHostFunctionName = JSON.stringify(hostFunctionName);
  const missingGuard =
    hostFunctionName !== "unknown host function"
      ? `if (typeof ${hostFunctionName} !== "function") { return JSON.stringify({ ok: false, error: ${missingError} }); }`
      : "";

  return `(function(){try{${missingGuard}return ${script};}catch(error){var message=String(error);var details={hostFunction:${quotedHostFunctionName},name:"",message:message,line:"",fileName:""};try{details.name=String(error&&error.name?error.name:"");details.message=String(error&&error.message?error.message:message);details.line=String(error&&error.line?error.line:"");details.fileName=String(error&&error.fileName?error.fileName:"");}catch(detailError){}return JSON.stringify({ ok: false, error: ${exceptionPrefix} + details.message, debug: details });}})()`;
}

function buildInvalidHostJsonResponse<T>(
  script: string,
  guardedScript: string,
  raw: unknown,
  parseError: unknown
): HostJsonResponse<T> {
  // // Convert CEP's opaque EvalScript failures into actionable JSON that users can paste from the debug log.
  const hostFunctionName = getHostFunctionName(script);
  const rawText = String(raw ?? "");
  const rawPreview = previewHostResponse(rawText);
  const looksLikeEvalScriptError = /^EvalScript error\.?$/i.test(rawText.trim());
  const hint = looksLikeEvalScriptError
    ? "Premiere rejected the ExtendScript call before Sub Creator could read an error. Restart Premiere Pro, reinstall Sub Creator with Premiere closed, then retry. If it persists, share this debug payload."
    : "Premiere returned a non-JSON response. Share this debug payload so the failing host call can be identified.";

  return {
    ok: false,
    error: `Invalid host response from ${hostFunctionName}: ${String(parseError)}. ${hint}`,
    debug: {
      hostFunction: hostFunctionName,
      rawPreview,
      rawLength: rawText.length,
      scriptLength: script.length,
      guardedScriptLength: guardedScript.length,
      cepHostAvailable: Boolean(window.__adobe_cep__),
      generatedAt: new Date().toISOString()
    }
  };
}

function readCepFileSnapshot(modules: CepNodeModules, filePath: string): Record<string, unknown> {
  // // Capture the temporary WAV state so export failures show whether Premiere created or locked a file.
  try {
    if (!filePath || !modules.fs.existsSync(filePath)) {
      return { path: filePath, exists: false };
    }

    const stats = modules.fs.statSync(filePath);
    return {
      path: filePath,
      exists: true,
      size: Number(stats.size || 0),
      mtimeMs: Number(stats.mtimeMs || 0),
      isFile: typeof stats.isFile === "function" ? stats.isFile() : undefined
    };
  } catch (error) {
    return { path: filePath, exists: "unknown", error: String(error) };
  }
}

function buildWhisperExportErrorMessage(message: string, debug: Record<string, unknown>): string {
  // // Attach a compact JSON payload to user-shareable errors from machine-specific Premiere export failures.
  return `${message}\nDebug payload:\n${JSON.stringify(debug, null, 2)}`;
}

function readCepHostEnvironmentDebug(): Record<string, unknown> {
  // // Read CEP host metadata when available so export logs identify the Premiere build and CEP runtime.
  try {
    const rawEnvironment = window.__adobe_cep__?.getHostEnvironment?.();
    if (!rawEnvironment) {
      return {};
    }
    const parsed = JSON.parse(rawEnvironment) as Record<string, unknown>;
    return {
      appName: parsed.appName,
      appVersion: parsed.appVersion,
      appLocale: parsed.appLocale,
      extensionId: parsed.extensionId,
      rawLength: rawEnvironment.length
    };
  } catch (error) {
    return { error: String(error) };
  }
}

async function evalHostJson<T>(script: string): Promise<HostJsonResponse<T>> {
  // // Parse JSON returned by host-side ExtendScript function calls.
  const guardedScript = buildGuardedHostJsonScript(script);
  const raw = await evalScript(guardedScript);

  try {
    return JSON.parse(raw) as HostJsonResponse<T>;
  } catch (error) {
    return buildInvalidHostJsonResponse<T>(script, guardedScript, raw, error);
  }
}

async function evalHostJsonRaw(script: string): Promise<string> {
  // // Preserve the historical raw-string API while still guarding host calls and malformed CEP responses.
  const response = await evalHostJson<Record<string, unknown>>(script);
  return JSON.stringify(response);
}

function resolveCepNodeModules(): CepNodeModules | null {
  // // Resolve Node modules from CEP mixed-context runtime when available.
  const nodeRequire =
    (window.cep_node && typeof window.cep_node.require === "function" ? window.cep_node.require : null) ||
    (typeof window.require === "function" ? window.require : null);
  if (!nodeRequire) {
    return null;
  }

  try {
    return {
      childProcess: nodeRequire("child_process") as CepNodeModules["childProcess"],
      fs: nodeRequire("fs") as CepNodeModules["fs"],
      os: nodeRequire("os") as CepNodeModules["os"],
      path: nodeRequire("path") as CepNodeModules["path"],
      process: nodeRequire("process") as CepNodeModules["process"]
    };
  } catch {
    return null;
  }
}

function createCancelledJobError(): Error {
  // // Normalize bridge-side cancellation into one stable error marker for the panel.
  const error = new Error(SUBCREATOR_CANCELLED_JOB_CODE);
  error.name = SUBCREATOR_CANCELLED_JOB_CODE;
  return error;
}

export function isCancelledJobError(error: unknown): boolean {
  // // Detect the stable cancellation marker no matter whether it comes from `name` or `message`.
  const errorName = error instanceof Error ? String(error.name || "").trim() : "";
  const errorMessage = error instanceof Error ? String(error.message || "").trim() : String(error ?? "").trim();
  return errorName === SUBCREATOR_CANCELLED_JOB_CODE || errorMessage === SUBCREATOR_CANCELLED_JOB_CODE;
}

function registerActiveCepJob(child: CepSpawnedChild, label: string, detached: boolean): ActiveCepJob {
  // // Track the currently running CEP child so the panel can stop Whisper/WhisperX jobs on demand.
  const job: ActiveCepJob = {
    label,
    child,
    detached,
    cancelRequested: false
  };
  activeCepJob = job;
  return job;
}

function clearActiveCepJob(child?: CepSpawnedChild): void {
  // // Clear only the matching tracked child so overlapping cleanup does not wipe a newer job accidentally.
  if (!activeCepJob) {
    return;
  }

  if (child && activeCepJob.child !== child) {
    return;
  }

  activeCepJob = null;
}

function terminateTrackedCepJob(modules: CepNodeModules, job: ActiveCepJob): boolean {
  // // Stop the active Whisper/WhisperX process tree with the safest strategy available on the current OS.
  const pid = Number(job.child.pid || 0);
  const env = modules.process.env || {};

  if (detectWindowsRuntime()) {
    if (pid > 0) {
      const taskkillResult = modules.childProcess.spawnSync("taskkill", ["/pid", String(pid), "/t", "/f"], {
        encoding: "utf8",
        timeout: 15000,
        env
      });
      if (!taskkillResult.error) {
        return true;
      }
    }

    try {
      job.child.kill?.();
      return true;
    } catch {
      return false;
    }
  }

  const canUseProcessKill = typeof modules.process.kill === "function" && pid > 0;
  if (canUseProcessKill && job.detached) {
    try {
      modules.process.kill?.(-pid, "SIGTERM");
      return true;
    } catch {
      // // Fall through to direct child termination when process-group kill is unavailable.
    }
  }

  try {
    job.child.kill?.("SIGTERM");
    return true;
  } catch {
    // // Fall through to SIGKILL fallback for stubborn Python children.
  }

  if (canUseProcessKill && job.detached) {
    try {
      modules.process.kill?.(-pid, "SIGKILL");
      return true;
    } catch {
      // // Fall through to direct child termination when process-group SIGKILL also fails.
    }
  }

  try {
    job.child.kill?.("SIGKILL");
    return true;
  } catch {
    return false;
  }
}

export async function cancelCurrentJob(): Promise<boolean> {
  // // Expose one panel-level stop hook for the currently running Whisper or WhisperX CEP job.
  const modules = resolveCepNodeModules();
  if (!modules || !activeCepJob) {
    return false;
  }

  if (activeCepJob.cancelRequested) {
    return true;
  }

  activeCepJob.cancelRequested = true;
  return terminateTrackedCepJob(modules, activeCepJob);
}

function pushUniqueString(list: string[], value: string): void {
  // // Keep small runtime lists unique while preserving insertion order and ignoring case-only duplicates.
  const normalized = String(value || "").trim();
  if (!normalized) {
    return;
  }

  const lookup = normalized.toLowerCase();
  for (const item of list) {
    if (String(item || "").trim().toLowerCase() === lookup) {
      return;
    }
  }

  list.push(normalized);
}

function listWhisperModelCacheDirectories(modules: CepNodeModules): string[] {
  // // Probe standard Whisper cache locations so the panel can show only the models already available locally.
  const directories: string[] = [];
  const env = modules.process.env || {};
  const homeDir = typeof modules.os.homedir === "function" ? String(modules.os.homedir() || "").trim() : "";
  const xdgCacheHome = String(env.XDG_CACHE_HOME || "").trim();
  const userProfile = String(env.USERPROFILE || "").trim();
  const windowsHome = `${String(env.HOMEDRIVE || "").trim()}${String(env.HOMEPATH || "").trim()}`.trim();

  if (xdgCacheHome) {
    pushUniqueString(directories, modules.path.join(xdgCacheHome, "whisper"));
  }
  if (detectWindowsRuntime() && userProfile) {
    pushUniqueString(directories, modules.path.join(userProfile, ".cache", "whisper"));
  }
  if (detectWindowsRuntime() && windowsHome) {
    pushUniqueString(directories, modules.path.join(windowsHome, ".cache", "whisper"));
  }
  if (homeDir) {
    pushUniqueString(directories, modules.path.join(homeDir, ".cache", "whisper"));
    pushUniqueString(directories, modules.path.join(homeDir, "Library", "Caches", "whisper"));
  }
  if (!detectWindowsRuntime() && userProfile) {
    pushUniqueString(directories, modules.path.join(userProfile, ".cache", "whisper"));
  }
  if (!detectWindowsRuntime() && windowsHome) {
    pushUniqueString(directories, modules.path.join(windowsHome, ".cache", "whisper"));
  }

  return directories.filter((directory) => modules.fs.existsSync(directory));
}

function resolvePreferredWhisperModelCacheDirectory(modules: CepNodeModules, preferredPaths?: string[]): string {
  // // Reuse an existing Whisper cache path when possible, otherwise fall back to the standard per-user cache location.
  const explicitCandidates = Array.isArray(preferredPaths)
    ? preferredPaths.map((value) => String(value || "").trim()).filter(Boolean)
    : [];
  const discoveredCandidates = listWhisperModelCacheDirectories(modules);
  const allCandidates = explicitCandidates.concat(discoveredCandidates);
  if (allCandidates.length > 0) {
    return allCandidates[0];
  }

  const env = modules.process.env || {};
  const xdgCacheHome = String(env.XDG_CACHE_HOME || "").trim();
  const userProfile = String(env.USERPROFILE || "").trim();
  const windowsHome = `${String(env.HOMEDRIVE || "").trim()}${String(env.HOMEPATH || "").trim()}`.trim();
  const homeDir = typeof modules.os.homedir === "function" ? String(modules.os.homedir() || "").trim() : "";
  if (xdgCacheHome) {
    return modules.path.join(xdgCacheHome, "whisper");
  }
  if (userProfile) {
    return modules.path.join(userProfile, ".cache", "whisper");
  }
  if (windowsHome) {
    return modules.path.join(windowsHome, ".cache", "whisper");
  }
  return modules.path.join(homeDir, ".cache", "whisper");
}

function folderOpenCommandSucceeded(
  result: { error?: { code?: string } | null; status?: number | null },
  windowsRuntime: boolean
): boolean {
  // // Windows Explorer commonly opens the folder successfully while returning exit code 1 through CEP.
  if (result.error) {
    return false;
  }
  return windowsRuntime || result.status === 0 || result.status === null;
}

function detectInstalledWhisperModelsViaCepNode(modules: CepNodeModules): {
  installedModels: string[];
  modelCachePaths: string[];
} {
  // // Match cached `.pt` files against the subset of Whisper models exposed in the panel.
  const modelCachePaths = listWhisperModelCacheDirectories(modules);
  const installedModels: string[] = [];

  for (const cachePath of modelCachePaths) {
    for (const modelDefinition of SUBCREATOR_SUPPORTED_WHISPER_MODELS) {
      if (
        modelDefinition.filenames.some((filename) =>
          modules.fs.existsSync(modules.path.join(cachePath, filename))
        )
      ) {
        pushUniqueString(installedModels, modelDefinition.value);
      }
    }
  }

  return {
    installedModels,
    modelCachePaths
  };
}

function buildWhisperArgs(request: WhisperTranscriptionRequest, outputDir: string): string[] {
  // // Build CLI arguments for local Whisper invocation.
  const model = request.model?.trim() || "base";
  const args = [
    request.audioPath,
    "--model",
    model,
    "--output_format",
    "all",
    "--output_dir",
    outputDir,
    "--fp16",
    "False",
    "--word_timestamps",
    "True"
  ];

  const language = request.languageCode?.trim();
  if (language && language.toLowerCase() !== "auto") {
    args.push("--language", language);
  }

  return args;
}

function buildWhisperOutputPath(modules: CepNodeModules, outputDir: string, audioPath: string, extension: string): string {
  // // Resolve one Whisper output file by preferred basename first, then by directory scan fallback.
  const normalizedExtension = String(extension || "").replace(/^\./, "").toLowerCase();
  const baseName = modules.path.basename(audioPath).replace(/\.[^/.]+$/, "");
  const direct = modules.path.join(outputDir, `${baseName}.${normalizedExtension}`);
  if (modules.fs.existsSync(direct)) {
    return direct;
  }

  const entries = modules.fs.readdirSync(outputDir);
  for (const entry of entries) {
    const lower = String(entry).toLowerCase();
    if (lower.endsWith(`.${normalizedExtension}`) && lower.startsWith(baseName.toLowerCase())) {
      return modules.path.join(outputDir, entry);
    }
  }

  return "";
}

function cleanupWhisperOutputFiles(modules: CepNodeModules, outputDir: string, audioPath: string): void {
  // // Delete temporary Whisper output artifacts once the panel has loaded them into memory.
  const extensions = ["json", "srt", "txt", "tsv", "vtt"];
  for (const extension of extensions) {
    const filePath = buildWhisperOutputPath(modules, outputDir, audioPath, extension);
    if (!filePath || !modules.fs.existsSync(filePath)) {
      continue;
    }

    try {
      modules.fs.unlinkSync(filePath);
    } catch {
      // // Ignore cleanup failures because transcription already succeeded.
    }
  }
}

function findWhisperSequencePresetInSystemPresets(modules: CepNodeModules, systemPresetsRoot: string): string {
  // // Scan one level under AME system presets to find a usable WAV audio-only preset across app versions.
  if (!systemPresetsRoot || !modules.fs.existsSync(systemPresetsRoot)) {
    return "";
  }

  const preferredNames = [/^waveform audio 48khz 16-bit\.epr$/i, /^wav 48khz 16 bit\.epr$/i];

  function matchEntry(entryPath: string): string {
    if (!modules.fs.existsSync(entryPath)) {
      return "";
    }

    const stats = modules.fs.statSync(entryPath);
    if (stats?.isFile && stats.isFile()) {
      const fileName = modules.path.basename(entryPath);
      for (const pattern of preferredNames) {
        if (pattern.test(fileName)) {
          return entryPath;
        }
      }
      return "";
    }

    if (!stats?.isDirectory || !stats.isDirectory()) {
      return "";
    }

    for (const childEntry of modules.fs.readdirSync(entryPath)) {
      const matchedChild = matchEntry(modules.path.join(entryPath, childEntry));
      if (matchedChild) {
        return matchedChild;
      }
    }

    return "";
  }

  return matchEntry(systemPresetsRoot);
}

function detectWhisperSequencePresetPathViaCepNode(modules: CepNodeModules, extensionRootPath: string): string {
  // // Prefer the bundled WAV preset so active-sequence export is independent from Adobe language and preset installation.
  const normalizedExtensionRoot = String(extensionRootPath || "").trim();
  if (normalizedExtensionRoot) {
    const bundledPresetPath = modules.path.join(
      normalizedExtensionRoot,
      "assets",
      "presets",
      "SubCreator-WAV-48kHz-16bit.epr"
    );
    if (modules.fs.existsSync(bundledPresetPath)) {
      return bundledPresetPath;
    }
  }

  // // Keep Adobe system presets as a fallback for older or incomplete extension bundles.
  const candidates: Array<{ path: string; score: number }> = [];

  function pushCandidate(candidatePath: string, score: number): void {
    if (!candidatePath || !modules.fs.existsSync(candidatePath)) {
      return;
    }
    candidates.push({ path: candidatePath, score });
  }

  if (detectWindowsRuntime()) {
    const roots = [modules.process.env.ProgramFiles, modules.process.env["ProgramFiles(x86)"]].filter(Boolean) as string[];
    for (const root of roots) {
      const adobeRoot = modules.path.join(root, "Adobe");
      if (!modules.fs.existsSync(adobeRoot)) {
        continue;
      }

      for (const entry of modules.fs.readdirSync(adobeRoot)) {
        const normalizedEntry = String(entry || "");
        const versionMatch = normalizedEntry.match(/^Adobe (?:Media Encoder|Premiere Pro) (\d{4})$/);
        const betaMatch = /^Adobe Media Encoder(?: \(Beta\)| Beta)$/i.test(normalizedEntry);
        if (!versionMatch && !betaMatch) {
          continue;
        }

        const versionScore = versionMatch ? Number(versionMatch[1]) : 0;
        pushCandidate(
          findWhisperSequencePresetInSystemPresets(
            modules,
            modules.path.join(adobeRoot, normalizedEntry, "MediaIO", "systempresets")
          ),
          versionScore
        );
      }
    }
  } else {
    const applicationsRoot = "/Applications";
    if (modules.fs.existsSync(applicationsRoot)) {
      for (const entry of modules.fs.readdirSync(applicationsRoot)) {
        const normalizedEntry = String(entry || "");
        const versionMatch = normalizedEntry.match(/^Adobe (?:Media Encoder|Premiere Pro) (\d{4})$/);
        const betaMatch = /^Adobe Media Encoder(?: \(Beta\)| Beta)$/i.test(normalizedEntry);
        if (!versionMatch && !betaMatch) {
          continue;
        }

        const versionScore = versionMatch ? Number(versionMatch[1]) : 0;
        pushCandidate(
          findWhisperSequencePresetInSystemPresets(
            modules,
            modules.path.join(applicationsRoot, normalizedEntry, `${normalizedEntry}.app`, "Contents", "MediaIO", "systempresets")
          ),
          versionScore
        );
        pushCandidate(
          findWhisperSequencePresetInSystemPresets(
            modules,
            modules.path.join(applicationsRoot, `${normalizedEntry}.app`, "Contents", "MediaIO", "systempresets")
          ),
          versionScore
        );
      }
    }
  }

  candidates.sort((left, right) => right.score - left.score || left.path.localeCompare(right.path));
  return candidates[0]?.path || "";
}

function detectWindowsRuntime(): boolean {
  // // Detect Windows from CEP browser runtime for CLI fallback ordering.
  return /win/i.test(String(navigator?.platform || ""));
}

function normalizeMogrtPathText(value: string): string {
  // // Normalize MOGRT labels and path fragments into compact UI-safe text.
  return String(value || "")
    .replace(/[_]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeMogrtFileSystemPath(value: string): string {
  // // Normalize file-system paths so runtime catalog matching stays stable on macOS and Windows.
  return String(value || "")
    .replace(/\\/g, "/")
    .replace(/\/+/g, "/")
    .trim();
}

function buildMogrtRelativePath(rootPath: string, fullPath: string): string {
  // // Derive extension-relative MOGRT path without relying on optional Node `path.relative` helpers.
  const normalizedRoot = normalizeMogrtFileSystemPath(rootPath).replace(/\/+$/, "");
  const normalizedFullPath = normalizeMogrtFileSystemPath(fullPath);
  const lowerRoot = normalizedRoot.toLowerCase();
  const lowerFullPath = normalizedFullPath.toLowerCase();
  if (lowerFullPath.startsWith(`${lowerRoot}/`)) {
    return normalizedFullPath.slice(normalizedRoot.length + 1);
  }
  return normalizedFullPath;
}

function buildFileUrlFromSystemPath(filePath: string): string {
  // // Convert absolute system paths into `file://` URLs for panel image/video previews.
  const normalizedPath = normalizeMogrtFileSystemPath(filePath);
  if (!normalizedPath) {
    return "";
  }

  if (/^[a-zA-Z]:\//.test(normalizedPath)) {
    return `file:///${encodeURI(normalizedPath)}`;
  }

  if (normalizedPath.startsWith("/")) {
    return `file://${encodeURI(normalizedPath)}`;
  }

  return normalizedPath;
}

function subcreatorSlugifyMogrtId(input: string): string {
  // // Build stable runtime ids from relative paths so gallery selection survives rescans.
  return String(input || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 120);
}

function detectRuntimeMogrtPreviewClass(name: string): string {
  // // Reuse lightweight preview themes for dynamically discovered MOGRT cards.
  const lower = String(name || "").toLowerCase();
  if (lower.includes("clean")) {
    return "clean";
  }
  if (lower.includes("comic")) {
    return "comic";
  }
  if (lower.includes("glitch")) {
    return "glitch";
  }
  if (lower.includes("karaoke")) {
    return "karaoke";
  }
  if (lower.includes("typewriter")) {
    return "typewriter";
  }
  if (lower.includes("mr beast")) {
    return "mrbeast";
  }
  if (lower.includes("tiktok")) {
    return "tiktok";
  }
  if (lower.includes("akira")) {
    return "akira";
  }
  if (lower.includes("motion blur")) {
    return "motionblur";
  }
  if (lower.includes("marker")) {
    return "marker";
  }
  if (lower.includes("slide")) {
    return "slide";
  }
  if (lower.includes("slant")) {
    return "slant";
  }
  if (lower.includes("spinning")) {
    return "spinning";
  }
  if (lower.includes("block")) {
    return "block";
  }
  if (lower.includes("emphasis")) {
    return "emphasis";
  }
  if (lower.includes("obviously")) {
    return "obviously";
  }
  if (lower.includes("arch")) {
    return "arch";
  }
  return "default";
}

function pickRuntimeArchiveEntry(entryNames: string[], ruleMatchers: RegExp[]): string {
  // // Resolve preferred embedded thumbnail assets while preserving original archive entry casing.
  const normalizedEntries = entryNames.map((entryName) => ({
    raw: String(entryName || ""),
    normalized: String(entryName || "").replace(/\\/g, "/").toLowerCase()
  }));

  for (const rule of ruleMatchers) {
    const matchedEntry = normalizedEntries.find((entry) => rule.test(entry.normalized));
    if (matchedEntry) {
      return matchedEntry.raw;
    }
  }

  return "";
}

function normalizeRuntimeArchiveExt(entryName: string, fallbackExt: string): string {
  // // Keep extracted preview filenames stable even when archive assets use variant extensions.
  const matchedExt = String(entryName || "").toLowerCase().match(/\.[^.]+$/);
  const ext = matchedExt ? matchedExt[0] : fallbackExt;
  if (ext === ".jpeg") {
    return ".jpg";
  }
  return ext || fallbackExt;
}

function readRuntimeFileStat(
  modules: CepNodeModules,
  filePath: string
): { size: number; mtimeMs: number; isDirectory: boolean; isFile: boolean } | null {
  // // Normalize Node stat payloads so cache checks stay simple across CEP runtimes.
  try {
    const stat = modules.fs.statSync(filePath);
    return {
      size: Number(stat.size || 0),
      mtimeMs: Number(stat.mtimeMs || 0),
      isDirectory: typeof stat.isDirectory === "function" ? stat.isDirectory() : false,
      isFile: typeof stat.isFile === "function" ? stat.isFile() : false
    };
  } catch {
    return null;
  }
}

function isRuntimePreviewSidecarFile(entryName: string): boolean {
  // // Restrict gallery cache signatures to preview sidecars that can affect rendered thumbnails.
  return /\.(png|jpg|jpeg|webp|mp4|mov|webm)$/i.test(String(entryName || ""));
}

function resolveCachedRuntimePreviewFileUrl(
  modules: CepNodeModules,
  previewRoot: string,
  previewStem: string,
  mogrtStat: { size: number; mtimeMs: number; isDirectory: boolean; isFile: boolean } | null,
  extensions: string[]
): string {
  // // Reuse already extracted preview files while the source `.mogrt` timestamp stays unchanged.
  for (const extension of extensions) {
    const candidatePath = modules.path.join(previewRoot, `${previewStem}${extension}`);
    const candidateStat = readRuntimeFileStat(modules, candidatePath);
    if (!candidateStat || !candidateStat.isFile || candidateStat.size < 1) {
      continue;
    }
    if (mogrtStat && candidateStat.mtimeMs + 1 < mogrtStat.mtimeMs) {
      continue;
    }
    return buildFileUrlFromSystemPath(candidatePath);
  }
  return "";
}

function extractRuntimeMogrtPreviewFiles(
  modules: CepNodeModules,
  mogrtPath: string,
  relativePath: string
): { imageFileUrl: string; videoFileUrl: string } {
  // // Extract embedded `thumb.*` preview assets from a `.mogrt` archive for manually added gallery items.
  const previewRoot = modules.path.join(modules.os.tmpdir(), "subcreator-mogrt-previews");
  const previewStem = subcreatorSlugifyMogrtId(relativePath || modules.path.basename(mogrtPath));
  modules.fs.mkdirSync(previewRoot, { recursive: true });

  const mogrtStat = readRuntimeFileStat(modules, mogrtPath);
  const cachedImageFileUrl = resolveCachedRuntimePreviewFileUrl(modules, previewRoot, previewStem, mogrtStat, [
    ".png",
    ".jpg",
    ".webp"
  ]);
  const cachedVideoFileUrl = resolveCachedRuntimePreviewFileUrl(modules, previewRoot, previewStem, mogrtStat, [
    ".mp4",
    ".mov",
    ".webm"
  ]);
  if (cachedImageFileUrl || cachedVideoFileUrl) {
    return {
      imageFileUrl: cachedImageFileUrl,
      videoFileUrl: cachedVideoFileUrl
    };
  }

  let archiveMap: Record<string, Uint8Array> = {};

  try {
    const archiveBytes = modules.fs.readFileSync(mogrtPath);
    const archiveBuffer =
      typeof archiveBytes === "string"
        ? new TextEncoder().encode(archiveBytes)
        : archiveBytes instanceof Uint8Array
          ? archiveBytes
          : new Uint8Array();
    if (archiveBuffer.length < 1) {
      return { imageFileUrl: "", videoFileUrl: "" };
    }
    archiveMap = unzipSync(new Uint8Array(archiveBuffer));
  } catch {
    return { imageFileUrl: "", videoFileUrl: "" };
  }

  const entryNames = Object.keys(archiveMap);
  if (entryNames.length < 1) {
    return { imageFileUrl: "", videoFileUrl: "" };
  }

  const imageEntry = pickRuntimeArchiveEntry(entryNames, [/\/thumb\.png$/i, /thumb\.png$/i, /thumb\.(jpg|jpeg|webp)$/i, /\.(png|jpg|jpeg|webp)$/i]);
  const videoEntry = pickRuntimeArchiveEntry(entryNames, [/\/thumb\.mp4$/i, /thumb\.mp4$/i, /\.mp4$/i, /\.mov$/i, /\.webm$/i]);
  if (!imageEntry && !videoEntry) {
    return { imageFileUrl: "", videoFileUrl: "" };
  }

  let imageFileUrl = "";
  let videoFileUrl = "";

  if (imageEntry && archiveMap[imageEntry]) {
    const imageFilePath = modules.path.join(previewRoot, `${previewStem}${normalizeRuntimeArchiveExt(imageEntry, ".png")}`);
    try {
      modules.fs.writeFileSync(imageFilePath, archiveMap[imageEntry]);
      imageFileUrl = buildFileUrlFromSystemPath(imageFilePath);
    } catch {
      imageFileUrl = "";
    }
  }

  if (videoEntry && archiveMap[videoEntry]) {
    const videoFilePath = modules.path.join(previewRoot, `${previewStem}${normalizeRuntimeArchiveExt(videoEntry, ".mp4")}`);
    try {
      modules.fs.writeFileSync(videoFilePath, archiveMap[videoEntry]);
      videoFileUrl = buildFileUrlFromSystemPath(videoFilePath);
    } catch {
      videoFileUrl = "";
    }
  }

  return { imageFileUrl, videoFileUrl };
}

function decodePremiereTemplateTextBytes(bytes: Uint8Array): string {
  // // Mirror the host-side heuristic so CEP can recover the default visible text embedded in Premiere flatbuffer-style Source Text payloads.
  let bestText = "";
  let bestScore = Number.NEGATIVE_INFINITY;

  for (let index = 0; index <= bytes.length - 5; index += 1) {
    const byteLength =
      (bytes[index] ?? 0) |
      ((bytes[index + 1] ?? 0) << 8) |
      ((bytes[index + 2] ?? 0) << 16) |
      ((bytes[index + 3] ?? 0) << 24);

    if (!Number.isFinite(byteLength) || byteLength < 1 || byteLength > 2000) {
      continue;
    }

    const textStart = index + 4;
    const textEnd = textStart + byteLength;
    if (textEnd >= bytes.length || bytes[textEnd] !== 0) {
      continue;
    }

    const candidateBytes = bytes.slice(textStart, textEnd);
    let candidateText = "";
    try {
      candidateText = Buffer.from(candidateBytes).toString("utf8");
    } catch {
      continue;
    }

    if (!candidateText.trim()) {
      continue;
    }

    let printableCount = 0;
    for (const candidateByte of candidateBytes) {
      if ((candidateByte >= 32 && candidateByte <= 126) || candidateByte === 9 || candidateByte === 10 || candidateByte === 13 || candidateByte >= 128) {
        printableCount += 1;
      }
    }
    if (printableCount / candidateBytes.length < 0.8) {
      continue;
    }

    let hasOnlyZeroPaddingAfter = true;
    for (let suffixIndex = textEnd + 1; suffixIndex < bytes.length; suffixIndex += 1) {
      if (bytes[suffixIndex] !== 0) {
        hasOnlyZeroPaddingAfter = false;
        break;
      }
    }

    let score = (index / Math.max(bytes.length, 1)) * 5;
    score += candidateText.split(/\s+/).filter(Boolean).length * 2;
    score += Math.min(candidateText.length, 120) / 20;
    if (/[.,!?;:]/.test(candidateText)) {
      score += 1;
    }
    if (/[\r\n]/.test(candidateText)) {
      score += 1;
    }
    if (hasOnlyZeroPaddingAfter) {
      score += 3;
    }
    if (candidateText.length < 3) {
      score -= 10;
    }

    if (score > bestScore) {
      bestScore = score;
      bestText = candidateText.replace(/\r/g, "\n");
    }
  }

  return bestText;
}

function extractPremiereTemplateTextPayloads(
  modules: CepNodeModules,
  mogrtPath: string
): PremiereTemplateTextPayload[] {
  // // Read Premiere-authored `.mogrt` packages directly so the host can reuse the original text-document payload instead of flattening style to plain text.
  try {
    const archiveSource = modules.fs.readFileSync(mogrtPath);
    const archiveBytes = typeof archiveSource === "string" ? strToU8(archiveSource) : new Uint8Array(archiveSource);
    const outerArchive = unzipSync(archiveBytes);
    const definitionEntry = outerArchive["definition.json"];
    const projectEntry = outerArchive["project.prgraphic"];
    if (!definitionEntry || !projectEntry) {
      return [];
    }

    const definitionText = strFromU8(definitionEntry);
    const definition = JSON.parse(definitionText) as { authorApp?: unknown };
    if (String(definition?.authorApp || "").toLowerCase() !== "ppro") {
      return [];
    }

    const innerArchive = unzipSync(projectEntry);
    const projectKey = Object.keys(innerArchive).find((entryName) => /\.prproj$/i.test(String(entryName || "")));
    if (!projectKey) {
      return [];
    }

    const projectXml = strFromU8(gunzipSync(innerArchive[projectKey]));
    const payloads: PremiereTemplateTextPayload[] = [];
    const pattern = /<ArbVideoComponentParam[\s\S]*?<Name>([^<]+)<\/Name>[\s\S]*?<StartKeyframeValue[^>]*>([A-Za-z0-9+/=\s]+)<\/StartKeyframeValue>/g;

    for (const match of projectXml.matchAll(pattern)) {
      const displayName = String(match[1] || "").trim();
      const sourcePayloadXml = String(match[2] || "");
      const sourcePayloadBase64 = sourcePayloadXml.replace(/\s+/g, "");
      if (!displayName || !sourcePayloadBase64) {
        continue;
      }

      const lowerName = displayName.toLowerCase();
      if (lowerName.indexOf("source text") === -1 && lowerName.indexOf("text") === -1) {
        continue;
      }

      const initialText = decodePremiereTemplateTextBytes(Uint8Array.from(Buffer.from(sourcePayloadBase64, "base64")));
      if (!initialText) {
        continue;
      }

      payloads.push({
        displayName,
        initialText,
        sourcePayloadBase64,
        sourcePayloadXml
      });
    }

    return payloads;
  } catch {
    return [];
  }
}

function patchPremiereFlatbufferStringByAppending(bytes: Uint8Array, stringOffset: number, replacementBytes: Uint8Array): Uint8Array | null {
  const pointerOffsets: number[] = [];
  for (let offset = 0; offset <= stringOffset - 4; offset += 4) {
    const pointerValue =
      (bytes[offset] ?? 0) |
      ((bytes[offset + 1] ?? 0) << 8) |
      ((bytes[offset + 2] ?? 0) << 16) |
      ((bytes[offset + 3] ?? 0) << 24);
    if (pointerValue > 0 && offset + pointerValue === stringOffset) {
      pointerOffsets.push(offset);
    }
  }

  if (!pointerOffsets.length) {
    return null;
  }

  const patched = Array.from(bytes);
  while (patched.length % 4 !== 0) {
    patched.push(0);
  }

  const appendedStringOffset = patched.length;
  patched.push(replacementBytes.length & 255);
  patched.push((replacementBytes.length >> 8) & 255);
  patched.push((replacementBytes.length >> 16) & 255);
  patched.push((replacementBytes.length >> 24) & 255);
  for (const replacementByte of replacementBytes) {
    patched.push(replacementByte);
  }
  patched.push(0);
  while (patched.length % 4 !== 0) {
    patched.push(0);
  }

  for (const pointerOffset of pointerOffsets) {
    const relativeOffset = appendedStringOffset - pointerOffset;
    if (relativeOffset < 1) {
      return null;
    }
    patched[pointerOffset] = relativeOffset & 255;
    patched[pointerOffset + 1] = (relativeOffset >> 8) & 255;
    patched[pointerOffset + 2] = (relativeOffset >> 16) & 255;
    patched[pointerOffset + 3] = (relativeOffset >> 24) & 255;
  }

  return Uint8Array.from(patched);
}

function patchPremiereTemplatePayloadBase64(sourcePayloadBase64: string, nextText: string): string {
  // // Rewrite the UTF-8 text segment inside one Premiere Source Text payload while keeping the rest of the binary style document untouched.
  const bytes = Uint8Array.from(Buffer.from(String(sourcePayloadBase64 || "").replace(/\s+/g, ""), "base64"));
  if (!bytes.length) {
    return sourcePayloadBase64;
  }

  let bestOffset = -1;
  let bestByteLength = 0;
  let bestScore = Number.NEGATIVE_INFINITY;
  let bestHasOnlyZeroPaddingAfter = false;
  let bestZeroPaddingLength = 0;

  for (let index = 0; index <= bytes.length - 5; index += 1) {
    const byteLength =
      (bytes[index] ?? 0) |
      ((bytes[index + 1] ?? 0) << 8) |
      ((bytes[index + 2] ?? 0) << 16) |
      ((bytes[index + 3] ?? 0) << 24);
    if (!Number.isFinite(byteLength) || byteLength < 1 || byteLength > 2000) {
      continue;
    }

    const textStart = index + 4;
    const textEnd = textStart + byteLength;
    if (textEnd >= bytes.length || bytes[textEnd] !== 0) {
      continue;
    }

    const candidateBytes = bytes.slice(textStart, textEnd);
    const candidateText = Buffer.from(candidateBytes).toString("utf8");
    if (!candidateText.trim()) {
      continue;
    }

    let printableCount = 0;
    for (const candidateByte of candidateBytes) {
      if ((candidateByte >= 32 && candidateByte <= 126) || candidateByte === 9 || candidateByte === 10 || candidateByte === 13 || candidateByte >= 128) {
        printableCount += 1;
      }
    }
    if (printableCount / candidateBytes.length < 0.8) {
      continue;
    }

    let hasOnlyZeroPaddingAfter = true;
    let zeroPaddingLength = 0;
    for (let suffixIndex = textEnd + 1; suffixIndex < bytes.length; suffixIndex += 1) {
      if (bytes[suffixIndex] !== 0) {
        hasOnlyZeroPaddingAfter = false;
        break;
      }
      zeroPaddingLength += 1;
    }

    let score = (index / Math.max(bytes.length, 1)) * 5;
    score += candidateText.split(/\s+/).filter(Boolean).length * 2;
    score += Math.min(candidateText.length, 120) / 20;
    if (/[.,!?;:]/.test(candidateText)) {
      score += 1;
    }
    if (/[\r\n]/.test(candidateText)) {
      score += 1;
    }
    if (hasOnlyZeroPaddingAfter) {
      score += 3;
    }
    if (candidateText.length < 3) {
      score -= 10;
    }

    if (score > bestScore) {
      bestScore = score;
      bestOffset = index;
      bestByteLength = byteLength;
      bestHasOnlyZeroPaddingAfter = hasOnlyZeroPaddingAfter;
      bestZeroPaddingLength = zeroPaddingLength;
    }
  }

  if (bestOffset < 0) {
    return sourcePayloadBase64;
  }

  const replacementBytes = Uint8Array.from(Buffer.from(String(nextText || "").replace(/\r\n/g, "\n").replace(/\n/g, "\r"), "utf8"));
  if (!bestHasOnlyZeroPaddingAfter && replacementBytes.length !== bestByteLength) {
    const retargetedPayload = patchPremiereFlatbufferStringByAppending(bytes, bestOffset, replacementBytes);
    if (retargetedPayload) {
      return Buffer.from(retargetedPayload).toString("base64");
    }
    return sourcePayloadBase64;
  }

  const patched: number[] = [];
  const textStart = bestOffset + 4;
  const suffixStart = textStart + bestByteLength + 1;

  for (let index = 0; index < bestOffset; index += 1) {
    patched.push(bytes[index]);
  }

  patched.push(replacementBytes.length & 255);
  patched.push((replacementBytes.length >> 8) & 255);
  patched.push((replacementBytes.length >> 16) & 255);
  patched.push((replacementBytes.length >> 24) & 255);

  for (const replacementByte of replacementBytes) {
    patched.push(replacementByte);
  }
  patched.push(0);

  if (bestHasOnlyZeroPaddingAfter) {
    for (let zeroIndex = 0; zeroIndex < bestZeroPaddingLength; zeroIndex += 1) {
      patched.push(0);
    }
  } else {
    for (let padIndex = 0; padIndex < bestByteLength - replacementBytes.length; padIndex += 1) {
      patched.push(0);
    }
    for (let index = suffixStart; index < bytes.length; index += 1) {
      patched.push(bytes[index]);
    }
  }

  return Buffer.from(Uint8Array.from(patched)).toString("base64");
}

function patchPremiereDefinitionJsonText(definitionText: string, nextText: string): string {
  // // Keep definition-side client control text aligned with the patched project payload so Premiere metadata stays coherent.
  try {
    const definition = JSON.parse(String(definitionText || "")) as {
      authorApp?: unknown;
      clientControls?: Array<{
        type?: unknown;
        uiName?: { strDB?: Array<{ str?: unknown }> };
        value?: { strDB?: Array<{ str?: unknown }> };
      }>;
    };
    if (String(definition.authorApp || "").toLowerCase() !== "ppro" || !Array.isArray(definition.clientControls)) {
      return definitionText;
    }

    for (const control of definition.clientControls) {
      if (Number(control?.type) !== 6 || !Array.isArray(control?.value?.strDB)) {
        continue;
      }
      for (const localizedValue of control.value.strDB) {
        localizedValue.str = nextText;
      }
    }

    return JSON.stringify(definition);
  } catch {
    return definitionText;
  }
}

function patchPremiereProjectGraphicText(
  projectArchiveBytes: Uint8Array,
  payloads: PremiereTemplateTextPayload[],
  nextText: string
): Uint8Array {
  // // Rebuild the inner `.prgraphic` archive with patched Source Text payloads so imported Premiere-authored MOGRTs already contain the right text.
  const innerArchive = unzipSync(projectArchiveBytes);
  const prprojKey = Object.keys(innerArchive).find((entryName) => /\.prproj$/i.test(String(entryName || "")));
  if (!prprojKey) {
    return projectArchiveBytes;
  }

  let projectXml = strFromU8(gunzipSync(innerArchive[prprojKey]));
  for (const payload of payloads) {
    const sourcePayloadBase64 = String(payload.sourcePayloadBase64 || "").replace(/\s+/g, "");
    if (!sourcePayloadBase64) {
      continue;
    }
    const patchedPayloadBase64 = patchPremiereTemplatePayloadBase64(sourcePayloadBase64, nextText);
    const sourcePayloadXml = String(payload.sourcePayloadXml || "");
    if (sourcePayloadXml && projectXml.indexOf(sourcePayloadXml) !== -1) {
      projectXml = projectXml.replace(sourcePayloadXml, patchedPayloadBase64);
    } else {
      projectXml = projectXml.replace(sourcePayloadBase64, patchedPayloadBase64);
    }
  }

  innerArchive[prprojKey] = gzipSync(strToU8(projectXml));
  return zipSync(innerArchive);
}

export async function buildPremiereTemplateCueMogrts(
  mogrtPath: string,
  cues: CaptionCue[],
  payloads: PremiereTemplateTextPayload[]
): Promise<{ cues: CaptionCue[]; cleanupPaths: string[] }> {
  // // Generate one temporary `.mogrt` per cue for Premiere-authored templates, then still let the host rewrite text as a safety pass after import.
  const modules = resolveCepNodeModules();
  if (!modules || !payloads.length || !cues.length) {
    return {
      cues,
      cleanupPaths: []
    };
  }

  const sourceBytes = new Uint8Array(modules.fs.readFileSync(mogrtPath) as Uint8Array);
  if (!sourceBytes.length) {
    return {
      cues,
      cleanupPaths: []
    };
  }

  const outerArchive = unzipSync(sourceBytes);
  const definitionEntry = outerArchive["definition.json"];
  const projectEntry = outerArchive["project.prgraphic"];
  if (!definitionEntry || !projectEntry) {
    return {
      cues,
      cleanupPaths: []
    };
  }

  const tempRoot = modules.path.join(modules.os.tmpdir(), "subcreator-premiere-text-mogrts");
  modules.fs.mkdirSync(tempRoot, { recursive: true });

  const baseName = modules.path.basename(mogrtPath).replace(/\.mogrt$/i, "");
  const cleanupPaths: string[] = [];
  const patchedCues: CaptionCue[] = [];

  for (let cueIndex = 0; cueIndex < cues.length; cueIndex += 1) {
    const cue = cues[cueIndex];
    const patchedArchive = { ...outerArchive };
    patchedArchive["definition.json"] = strToU8(patchPremiereDefinitionJsonText(strFromU8(definitionEntry), cue.text));
    patchedArchive["project.prgraphic"] = patchPremiereProjectGraphicText(projectEntry, payloads, cue.text);

    const tempPath = modules.path.join(tempRoot, `${baseName}-${Date.now()}-${cueIndex}.mogrt`);
    modules.fs.writeFileSync(tempPath, zipSync(patchedArchive));
    cleanupPaths.push(tempPath);
    patchedCues.push({
      ...cue,
      mogrtPathOverride: tempPath,
      skipTextApply: true
    });
  }

  return {
    cues: patchedCues,
    cleanupPaths
  };
}

function resolveRuntimeMogrtPreviewFiles(
  modules: CepNodeModules,
  mogrtPath: string,
  relativePath: string,
  directoryPath: string,
  mogrtBaseName: string,
  directoryEntries: string[]
): { imageFileUrl: string; videoFileUrl: string } {
  // // Detect optional sidecar preview files so manually added MOGRTs can show custom gallery thumbnails.
  const normalizedBaseName = String(mogrtBaseName || "").trim().toLowerCase();
  const imageExtensions = [".png", ".jpg", ".jpeg", ".webp"];
  const videoExtensions = [".mp4", ".mov", ".webm"];

  const findEntry = (candidateNames: string[]): string => {
    for (const candidateName of candidateNames) {
      const normalizedCandidate = String(candidateName || "").trim().toLowerCase();
      const matchedEntry = directoryEntries.find((entry) => String(entry || "").trim().toLowerCase() === normalizedCandidate);
      if (matchedEntry) {
        return matchedEntry;
      }
    }
    return "";
  };

  const imageEntry = findEntry(
    [
      ...imageExtensions.map((extension) => `${normalizedBaseName}${extension}`),
      ...imageExtensions.map((extension) => `thumb${extension}`)
    ].filter(Boolean)
  );
  const videoEntry = findEntry(
    [
      ...videoExtensions.map((extension) => `${normalizedBaseName}${extension}`),
      ...videoExtensions.map((extension) => `thumb${extension}`)
    ].filter(Boolean)
  );

  const sidecarPreviewFiles = {
    imageFileUrl: imageEntry ? buildFileUrlFromSystemPath(modules.path.join(directoryPath, imageEntry)) : "",
    videoFileUrl: videoEntry ? buildFileUrlFromSystemPath(modules.path.join(directoryPath, videoEntry)) : ""
  };
  if (sidecarPreviewFiles.imageFileUrl || sidecarPreviewFiles.videoFileUrl) {
    return sidecarPreviewFiles;
  }

  return extractRuntimeMogrtPreviewFiles(modules, mogrtPath, relativePath);
}

function buildInstalledMogrtCatalogSignature(modules: CepNodeModules, templatesRoot: string): string {
  // // Build a cheap filesystem signature so passive gallery refreshes can skip full rescan/extraction when unchanged.
  const visited = new Set<string>();
  const queue: string[] = [templatesRoot];
  const rows: string[] = [];

  while (queue.length > 0) {
    const currentDirectory = queue.shift();
    if (!currentDirectory) {
      continue;
    }

    const normalizedDirectory = normalizeMogrtFileSystemPath(currentDirectory);
    if (!normalizedDirectory || visited.has(normalizedDirectory)) {
      continue;
    }
    visited.add(normalizedDirectory);

    let entries: string[] = [];
    try {
      entries = modules.fs.readdirSync(currentDirectory).slice().sort((left, right) => left.localeCompare(right));
    } catch {
      continue;
    }

    for (const entryName of entries) {
      const fullPath = modules.path.join(currentDirectory, entryName);
      const entryStat = readRuntimeFileStat(modules, fullPath);
      if (!entryStat) {
        continue;
      }

      if (entryStat.isDirectory) {
        queue.push(fullPath);
        continue;
      }

      if (!/\.mogrt$/i.test(String(entryName || "")) && !isRuntimePreviewSidecarFile(entryName)) {
        continue;
      }

      const relativePath = buildMogrtRelativePath(templatesRoot, fullPath);
      rows.push(`${relativePath}|${entryStat.size}|${Math.floor(entryStat.mtimeMs)}`);
    }
  }

  return rows.join("\n");
}

function readInstalledMogrtCatalogViaCepNode(extensionRootPath: string): InstalledMogrtCatalog | null {
  // // Scan installed `templates/mogrt` folders so the panel sees bundled and manually added templates.
  const modules = resolveCepNodeModules();
  if (!modules) {
    return null;
  }

  const normalizedExtensionRoot = String(extensionRootPath || "").trim();
  if (!normalizedExtensionRoot) {
    return null;
  }

  const templatesRoot = modules.path.join(normalizedExtensionRoot, "templates", "mogrt");
  if (!modules.fs.existsSync(templatesRoot)) {
    modules.fs.mkdirSync(templatesRoot, { recursive: true });
  }

  const signature = buildInstalledMogrtCatalogSignature(modules, templatesRoot);
  if (
    subcreatorInstalledMogrtCatalogCache &&
    subcreatorInstalledMogrtCatalogCache.templatesRoot === templatesRoot &&
    subcreatorInstalledMogrtCatalogCache.signature === signature
  ) {
    return subcreatorInstalledMogrtCatalogCache.catalog;
  }

  const groups: string[] = [];
  const templates: MogrtTemplateItem[] = [];
  const visited = new Set<string>();
  const queue: string[] = [templatesRoot];

  while (queue.length > 0) {
    const currentDirectory = queue.shift();
    if (!currentDirectory) {
      continue;
    }

    const normalizedDirectory = normalizeMogrtFileSystemPath(currentDirectory);
    if (!normalizedDirectory || visited.has(normalizedDirectory)) {
      continue;
    }
    visited.add(normalizedDirectory);

    let entries: string[] = [];
    try {
      entries = modules.fs.readdirSync(currentDirectory);
    } catch {
      continue;
    }

    for (const entryName of entries) {
      const fullPath = modules.path.join(currentDirectory, entryName);
      if (/\.mogrt$/i.test(String(entryName || ""))) {
        const relativePath = buildMogrtRelativePath(templatesRoot, fullPath);
        const pathParts = relativePath.split("/").filter(Boolean);
        const groupName = normalizeMogrtPathText(pathParts.length > 1 ? pathParts[0] : "General") || "General";
        const fileBaseName = String(entryName || "").replace(/\.[^.]+$/i, "");
        const previewFiles = resolveRuntimeMogrtPreviewFiles(modules, fullPath, relativePath, currentDirectory, fileBaseName, entries);

        pushUniqueString(groups, groupName);
        templates.push({
          id: `runtime-${subcreatorSlugifyMogrtId(relativePath || fileBaseName)}`,
          name: normalizeMogrtPathText(fileBaseName) || fileBaseName,
          aspect: groupName,
          relativePath,
          previewClass: detectRuntimeMogrtPreviewClass(fileBaseName),
          previewImagePath: previewFiles.imageFileUrl,
          previewVideoPath: previewFiles.videoFileUrl
        });
        continue;
      }

      try {
        modules.fs.readdirSync(fullPath);
        queue.push(fullPath);
      } catch {
        // // Ignore non-directory sidecar files while scanning the gallery tree.
      }
    }
  }

  templates.sort((left, right) => {
    const groupCompare = String(left.aspect || "").localeCompare(String(right.aspect || ""), undefined, {
      sensitivity: "base"
    });
    if (groupCompare !== 0) {
      return groupCompare;
    }
    return String(left.name || "").localeCompare(String(right.name || ""), undefined, { sensitivity: "base" });
  });
  groups.sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));

  const signatureItemCount = signature ? signature.split("\n").filter(Boolean).length : 0;
  const catalog = {
    available: true,
    source: "cep-node-installed-templates",
    details: `templates=${templates.length} groups=${groups.length} signatureItems=${signatureItemCount}`,
    templatesRoot,
    groups,
    templates
  };
  subcreatorInstalledMogrtCatalogCache = {
    templatesRoot,
    signature,
    catalog
  };
  return catalog;
}

function normalizeFontText(value: string): string {
  // // Normalize font fragments from filenames into user-facing labels.
  return String(value || "")
    .replace(/[_]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function splitFontFamilyAndStyle(rawName: string): { family: string; style: string } {
  // // Infer family/style from common font filename conventions (`Family-BoldItalic`, `Family Bold`).
  const cleaned = normalizeFontText(rawName);
  if (!cleaned) {
    return { family: "", style: "" };
  }

  const firstHyphen = cleaned.indexOf("-");
  if (firstHyphen > 0) {
    const familyPart = normalizeFontText(cleaned.slice(0, firstHyphen));
    const stylePart = normalizeFontText(cleaned.slice(firstHyphen + 1));
    if (familyPart) {
      return {
        family: familyPart,
        style: stylePart || "Regular"
      };
    }
  }

  const styleKeywords = new Set([
    "thin",
    "hairline",
    "extralight",
    "ultralight",
    "light",
    "book",
    "regular",
    "roman",
    "medium",
    "semibold",
    "demibold",
    "bold",
    "extrabold",
    "ultrabold",
    "black",
    "heavy",
    "italic",
    "oblique",
    "condensed",
    "narrow",
    "expanded",
    "extended",
    "display",
    "caps",
    "smallcaps"
  ]);
  const words = cleaned.split(/\s+/).filter(Boolean);
  if (words.length < 2) {
    return {
      family: cleaned,
      style: "Regular"
    };
  }

  const styleWords: string[] = [];
  let cursor = words.length - 1;
  while (cursor >= 0) {
    const probe = words[cursor].toLowerCase();
    if (!styleKeywords.has(probe)) {
      break;
    }
    styleWords.unshift(words[cursor]);
    cursor -= 1;
  }

  if (styleWords.length < 1 || cursor < 0) {
    return {
      family: cleaned,
      style: "Regular"
    };
  }

  const family = normalizeFontText(words.slice(0, cursor + 1).join(" "));
  const style = normalizeFontText(styleWords.join(" "));
  return {
    family: family || cleaned,
    style: style || "Regular"
  };
}

function mergeStyleMapEntry(target: Record<string, string[]>, family: string, style: string): void {
  // // Store one style under one family with case-insensitive dedupe.
  const normalizedFamily = normalizeFontText(family);
  const normalizedStyle = normalizeFontSystemDisplayStyle(style);
  if (!normalizedFamily) {
    return;
  }

  const familyLookupKey = normalizedFamily.toLowerCase();
  const existingKey = Object.keys(target).find((entry) => entry.toLowerCase() === familyLookupKey) || normalizedFamily;
  if (!Array.isArray(target[existingKey])) {
    target[existingKey] = [];
  }
  pushUniqueString(target[existingKey], normalizedStyle);
}

function normalizeFontSystemDisplayStyle(style: string): string {
  // // Collapse system-only labels like `Plain` into panel-friendly style names.
  const normalizedStyle = normalizeFontText(style) || "Regular";
  const normalizedKey = normalizedStyle.toLowerCase();
  if (normalizedKey === "plain" || normalizedKey === "roman") {
    return "Regular";
  }
  return normalizedStyle;
}

function listFontSystemStyleAliases(style: string): string[] {
  // // Register equivalent style aliases so token lookup survives host/system naming differences.
  const displayStyle = normalizeFontSystemDisplayStyle(style);
  const aliases = [displayStyle];
  if (displayStyle.toLowerCase() === "regular") {
    aliases.push("Plain", "Roman");
  }
  return Array.from(new Set(aliases.map((entry) => normalizeFontText(entry)).filter(Boolean)));
}

function mergeFontTokenEntry(
  target: Record<string, Record<string, string>>,
  family: string,
  style: string,
  token: string
): void {
  // // Store one exact host font token for one display family/style pair.
  const normalizedFamily = normalizeFontText(family);
  const normalizedToken = normalizeFontText(token);
  const normalizedStyleAliases = listFontSystemStyleAliases(style);
  if (!normalizedFamily || normalizedStyleAliases.length < 1 || !normalizedToken) {
    return;
  }

  const familyLookupKey = normalizedFamily.toLowerCase();
  const existingFamily = Object.keys(target).find((entry) => entry.toLowerCase() === familyLookupKey) || normalizedFamily;
  if (!target[existingFamily] || typeof target[existingFamily] !== "object") {
    target[existingFamily] = {};
  }

  for (const normalizedStyle of normalizedStyleAliases) {
    const styleLookupKey = normalizedStyle.toLowerCase();
    const existingStyle =
      Object.keys(target[existingFamily]).find((entry) => entry.toLowerCase() === styleLookupKey) || normalizedStyle;
    if (!target[existingFamily][existingStyle]) {
      target[existingFamily][existingStyle] = normalizedToken;
    }
  }
}

function detectMacSystemFontCatalogViaSystemProfiler(modules: CepNodeModules): SystemFontCatalog | null {
  // // Read authoritative macOS font metadata so family/style values match the system text engine.
  const maxProfilerBuffer = 128 * 1024 * 1024;
  const commands = ["/usr/sbin/system_profiler", "system_profiler"];
  let result:
    | { status: number | null; stdout?: string; stderr?: string; error?: { message?: string; code?: string } }
    | null = null;
  for (const command of commands) {
    result = modules.childProcess.spawnSync(command, ["SPFontsDataType", "-json"], {
      encoding: "utf8",
      timeout: 60000,
      maxBuffer: maxProfilerBuffer,
      env: modules.process.env
    });
    if (!result.error && result.status === 0 && result.stdout) {
      break;
    }
  }
  if (!result || result.error || result.status !== 0 || !result.stdout) {
    return null;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(String(result.stdout || ""));
  } catch {
    return null;
  }

  const payload = parsed && typeof parsed === "object" ? (parsed as Record<string, unknown>) : null;
  const fontEntries = payload && Array.isArray(payload.SPFontsDataType) ? payload.SPFontsDataType : [];
  const stylesByFamily: Record<string, string[]> = {};
  const fontTokensByFamilyStyle: Record<string, Record<string, string>> = {};
  let typefaceCount = 0;

  for (const fontEntry of fontEntries) {
    const fontRecord = fontEntry && typeof fontEntry === "object" ? (fontEntry as Record<string, unknown>) : null;
    if (!fontRecord || String(fontRecord.enabled || "yes").toLowerCase() === "no") {
      continue;
    }

    const typefaces = Array.isArray(fontRecord.typefaces) ? fontRecord.typefaces : [];
    for (const typefaceEntry of typefaces) {
      const typeface = typefaceEntry && typeof typefaceEntry === "object" ? (typefaceEntry as Record<string, unknown>) : null;
      if (!typeface || String(typeface.enabled || "yes").toLowerCase() === "no") {
        continue;
      }

      const family = normalizeFontText(String(typeface.family || ""));
      const style = normalizeFontText(String(typeface.style || "")) || "Regular";
      const token = normalizeFontText(String(typeface._name || ""));
      if (!family || family.startsWith(".")) {
        continue;
      }

      mergeStyleMapEntry(stylesByFamily, family, style);
      mergeFontTokenEntry(fontTokensByFamilyStyle, family, style, token);
      typefaceCount += 1;
    }
  }

  const families = Object.keys(stylesByFamily).sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));
  if (families.length < 1) {
    return null;
  }

  for (const family of families) {
    stylesByFamily[family] = stylesByFamily[family]
      .slice()
      .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));
  }

  return {
    available: true,
    source: "mac-system-profiler",
    details: `families=${families.length} typefaces=${typefaceCount}`,
    families,
    stylesByFamily,
    fontTokensByFamilyStyle
  };
}

function isFontFileName(entryName: string): boolean {
  // // Filter files likely to be installable fonts.
  return /\.(ttf|otf|ttc|dfont)$/i.test(String(entryName || ""));
}

function listSystemFontDirectories(modules: CepNodeModules): string[] {
  // // Build platform-specific directories where user/system fonts are typically installed.
  const directories: string[] = [];
  const home = modules.os.homedir();

  if (detectWindowsRuntime()) {
    const winDir = String(modules.process.env.WINDIR || "C:\\Windows");
    const localAppData = String(modules.process.env.LOCALAPPDATA || "");
    pushUniqueString(directories, modules.path.join(winDir, "Fonts"));
    if (localAppData) {
      pushUniqueString(directories, modules.path.join(localAppData, "Microsoft", "Windows", "Fonts"));
    }
    return directories;
  }

  pushUniqueString(directories, modules.path.join(home, "Library", "Fonts"));
  pushUniqueString(directories, "/Library/Fonts");
  pushUniqueString(directories, "/System/Library/Fonts");
  return directories;
}

function detectSystemFontCatalogViaCepNode(): SystemFontCatalog | null {
  // // Scan local font folders from CEP Node runtime and build family/style catalog for dropdown fallback.
  const modules = resolveCepNodeModules();
  if (!modules) {
    subcreatorSystemFontCatalogCache = null;
    return null;
  }

  if (typeof subcreatorSystemFontCatalogCache !== "undefined") {
    if (
      subcreatorSystemFontCatalogCache &&
      !detectWindowsRuntime() &&
      subcreatorSystemFontCatalogCache.source === "mac-font-dirs"
    ) {
      // // Allow one lazy upgrade from filename fallback cache to authoritative `system_profiler` metadata.
      const upgradedCatalog = detectMacSystemFontCatalogViaSystemProfiler(modules);
      if (upgradedCatalog && upgradedCatalog.available) {
        subcreatorSystemFontCatalogCache = upgradedCatalog;
      }
    }
    return subcreatorSystemFontCatalogCache;
  }

  if (!detectWindowsRuntime()) {
    const macCatalog = detectMacSystemFontCatalogViaSystemProfiler(modules);
    if (macCatalog && macCatalog.available) {
      subcreatorSystemFontCatalogCache = macCatalog;
      return subcreatorSystemFontCatalogCache;
    }
  }

  const directories = listSystemFontDirectories(modules);
  const stylesByFamily: Record<string, string[]> = {};
  const fontTokensByFamilyStyle: Record<string, Record<string, string>> = {};
  const visited = new Set<string>();
  const queue: Array<{ path: string; depth: number }> = [];
  const maxDepth = 2;
  const maxFiles = 8000;
  let fileCount = 0;

  for (const directory of directories) {
    if (!directory || !modules.fs.existsSync(directory)) {
      continue;
    }
    queue.push({ path: directory, depth: 0 });
  }

  while (queue.length > 0 && fileCount < maxFiles) {
    const current = queue.shift();
    if (!current) {
      continue;
    }

    const normalizedPath = String(current.path || "");
    if (!normalizedPath || visited.has(normalizedPath)) {
      continue;
    }
    visited.add(normalizedPath);

    let entries: string[] = [];
    try {
      entries = modules.fs.readdirSync(normalizedPath);
    } catch {
      continue;
    }

    for (const entryName of entries) {
      if (fileCount >= maxFiles) {
        break;
      }

      const fullPath = modules.path.join(normalizedPath, entryName);
      if (isFontFileName(entryName)) {
        const baseName = String(entryName).replace(/\.[^.]+$/i, "");
        const parsed = splitFontFamilyAndStyle(baseName);
        if (parsed.family) {
          mergeStyleMapEntry(stylesByFamily, parsed.family, parsed.style || "Regular");
          mergeFontTokenEntry(fontTokensByFamilyStyle, parsed.family, parsed.style || "Regular", baseName);
          fileCount += 1;
        }
        continue;
      }

      if (current.depth >= maxDepth) {
        continue;
      }

      try {
        // // Probe subfolders without statSync by attempting to read directory entries.
        modules.fs.readdirSync(fullPath);
        queue.push({ path: fullPath, depth: current.depth + 1 });
      } catch {
        // // Ignore non-directory entries.
      }
    }
  }

  const families = Object.keys(stylesByFamily).sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));
  for (const family of families) {
    stylesByFamily[family] = stylesByFamily[family]
      .slice()
      .sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));
  }

  subcreatorSystemFontCatalogCache = {
    available: families.length > 0,
    source: detectWindowsRuntime() ? "windows-font-dirs" : "mac-font-dirs",
    details: `families=${families.length} scannedFiles=${fileCount} maxFiles=${maxFiles}`,
    families,
    stylesByFamily,
    fontTokensByFamilyStyle
  };

  return subcreatorSystemFontCatalogCache;
}

function normalizeRuntimeConfigString(value: unknown): string {
  // // Normalize optional runtime config values into clean strings.
  return typeof value === "string" ? value.trim() : "";
}

function normalizeRuntimeConfigPathHints(value: unknown): string[] {
  // // Normalize optional path hint arrays from installer-generated config JSON.
  if (!Array.isArray(value)) {
    return [];
  }

  const hints: string[] = [];
  for (const item of value) {
    pushUniqueString(hints, normalizeRuntimeConfigString(item));
  }

  return hints;
}

function splitCommandString(value: string): { command: string; args: string[] } | null {
  // // Split command text into executable + args for spawnSync when installers store launcher labels.
  const normalized = String(value || "").trim();
  if (!normalized) {
    return null;
  }

  const tokens = normalized.split(/\s+/).filter(Boolean);
  if (tokens.length < 1) {
    return null;
  }

  return {
    command: tokens[0],
    args: tokens.slice(1)
  };
}

function resolveRuntimeConfigPathCandidates(modules: CepNodeModules): string[] {
  // // Resolve user-local runtime config paths for macOS/Windows installer outputs.
  const home = modules.os.homedir();
  const candidates: string[] = [];

  if (detectWindowsRuntime()) {
    const appData = String(modules.process.env.APPDATA || modules.path.join(home, "AppData", "Roaming"));
    pushUniqueString(candidates, modules.path.join(appData, "SubCreator", "subcreator-runtime.json"));
    pushUniqueString(candidates, modules.path.join(appData, "PremiereSubCreator", "subcreator-runtime.json"));
    return candidates;
  }

  pushUniqueString(candidates, modules.path.join(home, "Library", "Application Support", "SubCreator", "subcreator-runtime.json"));
  pushUniqueString(candidates, modules.path.join(home, "Library", "Application Support", "PremiereSubCreator", "subcreator-runtime.json"));
  return candidates;
}

function buildWindowsPrivateRuntimeConfig(modules: CepNodeModules): SubcreatorRuntimeConfig | null {
  // // Recover the standard private runtime directly when an older installer wrote an unreadable or missing config file.
  if (!detectWindowsRuntime()) {
    return null;
  }

  const localAppData = String(modules.process.env.LOCALAPPDATA || "").trim();
  if (!localAppData) {
    return null;
  }

  const runtimeRoot = modules.path.join(localAppData, "SubCreator", "runtime");
  const pythonPath = modules.path.join(runtimeRoot, "python", "python.exe");
  const whisperPath = modules.path.join(runtimeRoot, "python", "Scripts", "whisper.exe");
  const ffmpegPath = modules.path.join(runtimeRoot, "ffmpeg", "bin", "ffmpeg.exe");
  if (!modules.fs.existsSync(pythonPath)) {
    return null;
  }

  const pathHints: string[] = [];
  pushUniqueString(pathHints, modules.path.dirname(pythonPath));
  if (modules.fs.existsSync(whisperPath)) {
    pushUniqueString(pathHints, modules.path.dirname(whisperPath));
  }
  if (modules.fs.existsSync(ffmpegPath)) {
    pushUniqueString(pathHints, modules.path.dirname(ffmpegPath));
  }
  pushUniqueString(pathHints, modules.path.join(String(modules.process.env.SystemRoot || "C:\\Windows"), "System32"));

  return {
    sourcePath: `${runtimeRoot} (automatic recovery)`,
    pythonCommand: pythonPath,
    pythonPath,
    pythonVersion: "",
    whisperPath: modules.fs.existsSync(whisperPath) ? whisperPath : "",
    ffmpegPath: modules.fs.existsSync(ffmpegPath) ? ffmpegPath : "",
    pathHints
  };
}

function readRuntimeConfigFromDisk(modules: CepNodeModules): SubcreatorRuntimeConfig | null {
  // // Read installer-generated runtime config so CEP can use exact binary paths reliably.
  const candidates = resolveRuntimeConfigPathCandidates(modules);
  for (const candidatePath of candidates) {
    if (!modules.fs.existsSync(candidatePath)) {
      continue;
    }

    try {
      const rawText = String(modules.fs.readFileSync(candidatePath, "utf8") || "").replace(/^\uFEFF/, "");
      if (!rawText.trim()) {
        continue;
      }

      const parsed = JSON.parse(rawText) as Record<string, unknown>;
      const runtimeConfig: SubcreatorRuntimeConfig = {
        sourcePath: candidatePath,
        pythonCommand: normalizeRuntimeConfigString(parsed.pythonCommand),
        pythonPath: normalizeRuntimeConfigString(parsed.pythonPath),
        pythonVersion: normalizeRuntimeConfigString(parsed.pythonVersion),
        whisperPath: normalizeRuntimeConfigString(parsed.whisperPath),
        ffmpegPath: normalizeRuntimeConfigString(parsed.ffmpegPath),
        pathHints: normalizeRuntimeConfigPathHints(parsed.pathHints)
      };

      if (
        runtimeConfig.pythonCommand ||
        runtimeConfig.pythonPath ||
        runtimeConfig.whisperPath ||
        runtimeConfig.ffmpegPath ||
        runtimeConfig.pathHints.length > 0
      ) {
        return runtimeConfig;
      }
    } catch {
      // // Ignore malformed runtime config files and continue fallback probing.
    }
  }

  return buildWindowsPrivateRuntimeConfig(modules);
}

function getRuntimeConfig(modules: CepNodeModules): SubcreatorRuntimeConfig | null {
  // // Re-read when cache is still null so the panel can pick up a runtime config written after the first failed probe.
  if (typeof subcreatorRuntimeConfigCache !== "undefined" && subcreatorRuntimeConfigCache !== null) {
    return subcreatorRuntimeConfigCache;
  }

  subcreatorRuntimeConfigCache = readRuntimeConfigFromDisk(modules);
  return subcreatorRuntimeConfigCache;
}

function discoverUserWhisperExecutables(modules: CepNodeModules, runtimeConfig: SubcreatorRuntimeConfig | null): string[] {
  // // Probe common user-local installation locations where PATH may be incomplete in CEP.
  const discovered: string[] = [];
  if (runtimeConfig && runtimeConfig.whisperPath && modules.fs.existsSync(runtimeConfig.whisperPath)) {
    pushUniqueString(discovered, runtimeConfig.whisperPath);
  }

  const home = modules.os.homedir();
  const directCandidates = [modules.path.join(home, ".local", "bin", "whisper"), modules.path.join(home, "bin", "whisper")];

  for (const candidate of directCandidates) {
    if (modules.fs.existsSync(candidate)) {
      pushUniqueString(discovered, candidate);
    }
  }

  const pythonRoot = modules.path.join(home, "Library", "Python");
  if (modules.fs.existsSync(pythonRoot)) {
    try {
      const versions = modules.fs.readdirSync(pythonRoot);
      for (const versionName of versions) {
        const candidate = modules.path.join(pythonRoot, versionName, "bin", "whisper");
        if (modules.fs.existsSync(candidate)) {
          pushUniqueString(discovered, candidate);
        }
      }
    } catch {
      // // Ignore inaccessible folders and continue other command candidates.
    }
  }

  return discovered;
}

function resolveWhisperInterpreterFromScript(modules: CepNodeModules, executablePath: string): string {
  // // Resolve interpreter from shebang so we can run `-m whisper` on the matching Python runtime.
  try {
    const scriptText = String(modules.fs.readFileSync(executablePath, "utf8") || "");
    const firstLine = scriptText.split(/\r?\n/)[0] || "";
    if (!firstLine.startsWith("#!")) {
      return "";
    }

    const shebang = firstLine.slice(2).trim();
    const envMatch = shebang.match(/^\/usr\/bin\/env\s+(\S+)/);
    if (envMatch && envMatch[1]) {
      return envMatch[1];
    }

    const directMatch = shebang.match(/^(\S+)/);
    return directMatch && directMatch[1] ? directMatch[1] : "";
  } catch {
    return "";
  }
}

function buildSpawnEnv(
  modules: CepNodeModules,
  userExecutables: string[],
  runtimeConfig: SubcreatorRuntimeConfig | null
): Record<string, string | undefined> {
  // // Extend PATH for CEP-spawned subprocesses so ffmpeg/python locations are discoverable.
  const delimiter = detectWindowsRuntime() ? ";" : ":";
  const currentPath = String(modules.process.env.PATH || "");
  const segments = currentPath.length > 0 ? currentPath.split(delimiter).filter(Boolean) : [];
  const lowerSegments = segments.map((segment) => segment.toLowerCase());

  const extraSegments = detectWindowsRuntime()
    ? [
        "C:\\Program Files\\ffmpeg\\bin",
        "C:\\ffmpeg\\bin",
        modules.path.join(String(modules.process.env.SystemRoot || "C:\\Windows"), "System32")
      ]
    : ["/opt/homebrew/bin", "/usr/local/bin", "/usr/bin", "/bin"];

  if (runtimeConfig) {
    for (const hintPath of runtimeConfig.pathHints) {
      extraSegments.push(hintPath);
    }

    if (runtimeConfig.whisperPath) {
      extraSegments.push(modules.path.dirname(runtimeConfig.whisperPath));
    }
    if (runtimeConfig.pythonPath) {
      extraSegments.push(modules.path.dirname(runtimeConfig.pythonPath));
    }
    if (runtimeConfig.ffmpegPath) {
      extraSegments.push(modules.path.dirname(runtimeConfig.ffmpegPath));
    }
  }

  for (const executablePath of userExecutables) {
    extraSegments.push(modules.path.dirname(executablePath));
  }

  for (const extra of extraSegments) {
    if (!extra) {
      continue;
    }

    if (lowerSegments.indexOf(extra.toLowerCase()) !== -1) {
      continue;
    }

    segments.push(extra);
    lowerSegments.push(extra.toLowerCase());
  }

  return {
    ...modules.process.env,
    PATH: segments.join(delimiter)
  };
}

function buildWhisperCommandCandidates(
  modules: CepNodeModules,
  request: WhisperTranscriptionRequest,
  outputDir: string,
  userExecutables: string[],
  runtimeConfig: SubcreatorRuntimeConfig | null
): WhisperCommandCandidate[] {
  // // Build ordered command fallbacks for diverse Whisper install methods.
  const baseArgs = buildWhisperArgs(request, outputDir);
  const candidates: WhisperCommandCandidate[] = [];
  const isWindows = detectWindowsRuntime();

  function pushCandidate(command: string, args: string[], label: string): void {
    if (!command) {
      return;
    }

    for (const existing of candidates) {
      if (existing.command === command && existing.args.join("\u0001") === args.join("\u0001")) {
        return;
      }
    }

    candidates.push({ command, args, label });
  }

  if (runtimeConfig) {
    if (runtimeConfig.whisperPath) {
      pushCandidate(runtimeConfig.whisperPath, baseArgs, runtimeConfig.whisperPath);
    }

    if (runtimeConfig.pythonPath) {
      pushCandidate(runtimeConfig.pythonPath, ["-m", "whisper", ...baseArgs], `${runtimeConfig.pythonPath} -m whisper`);
    }

    const pythonCommand = splitCommandString(runtimeConfig.pythonCommand);
    if (pythonCommand) {
      pushCandidate(
        pythonCommand.command,
        [...pythonCommand.args, "-m", "whisper", ...baseArgs],
        `${runtimeConfig.pythonCommand} -m whisper`
      );
    }
  }

  for (const executablePath of userExecutables) {
    pushCandidate(executablePath, baseArgs, executablePath);
    const interpreter = resolveWhisperInterpreterFromScript(modules, executablePath);
    if (interpreter) {
      pushCandidate(interpreter, ["-m", "whisper", ...baseArgs], `${interpreter} -m whisper`);
    }
  }

  const versionedPython = ["python3.11", "python3.12", "python3.10", "python3.13", "python3.9", "python3.8"];
  for (const pythonCommand of versionedPython) {
    pushCandidate(pythonCommand, ["-m", "whisper", ...baseArgs], `${pythonCommand} -m whisper`);
  }

  pushCandidate("whisper", baseArgs, "whisper");
  pushCandidate("python3", ["-m", "whisper", ...baseArgs], "python3 -m whisper");
  pushCandidate("python", ["-m", "whisper", ...baseArgs], "python -m whisper");

  if (isWindows) {
    pushCandidate("py", ["-3.11", "-m", "whisper", ...baseArgs], "py -3.11 -m whisper");
    pushCandidate("py", ["-3.10", "-m", "whisper", ...baseArgs], "py -3.10 -m whisper");
    pushCandidate("py", ["-3.12", "-m", "whisper", ...baseArgs], "py -3.12 -m whisper");
    pushCandidate("py", ["-3.13", "-m", "whisper", ...baseArgs], "py -3.13 -m whisper");
    pushCandidate("py", ["-3.9", "-m", "whisper", ...baseArgs], "py -3.9 -m whisper");
    pushCandidate("py", ["-3.8", "-m", "whisper", ...baseArgs], "py -3.8 -m whisper");
    pushCandidate("py", ["-3", "-m", "whisper", ...baseArgs], "py -3 -m whisper");
    pushCandidate("py", ["-m", "whisper", ...baseArgs], "py -m whisper");
  }

  return candidates;
}

function buildPythonLauncherCandidates(
  modules: CepNodeModules,
  userExecutables: string[],
  runtimeConfig: SubcreatorRuntimeConfig | null
): PythonLauncherCandidate[] {
  // // Build Python launcher candidates used to detect `openai-whisper` module availability.
  const candidates: PythonLauncherCandidate[] = [];
  const isWindows = detectWindowsRuntime();

  function pushCandidate(command: string, argsPrefix: string[], label: string): void {
    if (!command) {
      return;
    }

    for (const existing of candidates) {
      if (existing.command === command && existing.argsPrefix.join("\u0001") === argsPrefix.join("\u0001")) {
        return;
      }
    }

    candidates.push({ command, argsPrefix, label });
  }

  if (runtimeConfig) {
    if (runtimeConfig.pythonPath) {
      pushCandidate(runtimeConfig.pythonPath, [], runtimeConfig.pythonPath);
    }

    const pythonCommand = splitCommandString(runtimeConfig.pythonCommand);
    if (pythonCommand) {
      pushCandidate(
        pythonCommand.command,
        pythonCommand.args,
        runtimeConfig.pythonCommand || pythonCommand.command
      );
    }
  }

  for (const executablePath of userExecutables) {
    const interpreter = resolveWhisperInterpreterFromScript(modules, executablePath);
    if (interpreter) {
      pushCandidate(interpreter, [], interpreter);
    }
  }

  if (isWindows) {
    // // Prefer the Windows Python launcher before generic aliases so Store stubs or unrelated `python.exe` do not block WhisperX.
    pushCandidate("py", ["-3.11"], "py -3.11");
    pushCandidate("py", ["-3.10"], "py -3.10");
    pushCandidate("py", ["-3.12"], "py -3.12");
    pushCandidate("py", ["-3.13"], "py -3.13");
    pushCandidate("py", ["-3.9"], "py -3.9");
    pushCandidate("py", ["-3.8"], "py -3.8");
    pushCandidate("py", ["-3"], "py -3");
    pushCandidate("py", [], "py");
    pushCandidate("python", [], "python");
    pushCandidate("python3", [], "python3");
    return candidates;
  }

  const posixVersioned = ["python3.11", "python3.12", "python3.10", "python3.13", "python3.9", "python3.8"];
  for (const command of posixVersioned) {
    pushCandidate(command, [], command);
  }

  pushCandidate("python3", [], "python3");
  pushCandidate("python", [], "python");

  return candidates;
}

function runSpawn(
  modules: CepNodeModules,
  command: string,
  args: string[],
  spawnEnv: Record<string, string | undefined>
): { ok: boolean; status: number; stdout: string; stderr: string; errorCode: string; errorMessage: string } {
  // // Execute a command synchronously with a safe timeout and normalized result shape.
  const run = modules.childProcess.spawnSync(command, args, {
    encoding: "utf8",
    shell: false,
    timeout: 12000,
    env: spawnEnv
  });

  const status = typeof run.status === "number" ? run.status : -1;
  const stdout = String(run.stdout || "");
  const stderr = String(run.stderr || "");
  const errorCode = String(run.error?.code || "");
  const errorMessage = String(run.error?.message || "");

  return {
    ok: !run.error && status === 0,
    status,
    stdout,
    stderr,
    errorCode,
    errorMessage
  };
}

function detectPythonModuleAvailabilityViaCepNode(
  modules: CepNodeModules,
  moduleName: string,
  runtimeConfig: SubcreatorRuntimeConfig | null,
  userExecutables: string[],
  spawnEnv: Record<string, string | undefined>
): { available: boolean; details: string } {
  // // Probe one Python module through the same launcher discovery used for Whisper/WhisperX commands.
  const checks: string[] = [];
  const pythonLaunchers = buildPythonLauncherCandidates(modules, userExecutables, runtimeConfig);
  const importStatement = `import ${moduleName}`;

  for (const launcher of pythonLaunchers) {
    const probe = runSpawn(modules, launcher.command, [...launcher.argsPrefix, "-c", importStatement], spawnEnv);
    if (probe.ok) {
      return {
        available: true,
        details: `Python module detected via ${launcher.label}${runtimeConfig ? ` (config: ${runtimeConfig.sourcePath})` : ""}`
      };
    }

    if (probe.errorCode === "ENOENT") {
      checks.push(`${launcher.label}: missing`);
    } else if (probe.errorMessage) {
      checks.push(`${launcher.label}: ${probe.errorMessage}`);
    } else if (probe.stderr.trim()) {
      checks.push(`${launcher.label}: ${probe.stderr.trim().split("\n")[0]}`);
    } else {
      checks.push(`${launcher.label}: exit ${probe.status}`);
    }
  }

  return {
    available: false,
    details: `${checks.join(" | ")}${runtimeConfig ? ` | config=${runtimeConfig.sourcePath}` : ""}`
  };
}

function isRuntimeConfigPythonLauncher(launcher: PythonLauncherCandidate, runtimeConfig: SubcreatorRuntimeConfig | null): boolean {
  // // Treat only explicit installer/runtime Python choices as authoritative enough to stop fallback probing.
  if (!runtimeConfig) {
    return false;
  }

  if (runtimeConfig.pythonPath && launcher.command === runtimeConfig.pythonPath && launcher.argsPrefix.length === 0) {
    return true;
  }

  const pythonCommand = splitCommandString(runtimeConfig.pythonCommand);
  if (!pythonCommand || launcher.command !== pythonCommand.command) {
    return false;
  }

  return launcher.argsPrefix.join("\u0001") === pythonCommand.args.join("\u0001");
}

function shouldContinueAfterPythonHelperFailure(summary: string): boolean {
  // // Keep probing alternate Python launchers when the helper ran under the wrong Python installation.
  return /no module named ['"]?whisperx|python was not found|microsoft store|app execution aliases/i.test(String(summary || ""));
}

function detectWhisperAvailabilityViaCepNode(): WhisperRuntimeStatus {
  // // Detect local Whisper runtime availability to decide if Whisper source should be shown.
  const modules = resolveCepNodeModules();
  if (!modules) {
    return {
      available: false,
      details: "CEP Node runtime unavailable",
      installedModels: [],
      modelCachePaths: [],
      alignmentAvailable: false,
      alignmentDetails: "CEP Node runtime unavailable"
    };
  }

  const runtimeConfig = getRuntimeConfig(modules);
  const userExecutables = discoverUserWhisperExecutables(modules, runtimeConfig);
  const spawnEnv = buildSpawnEnv(modules, userExecutables, runtimeConfig);
  const installedModels = detectInstalledWhisperModelsViaCepNode(modules);
  const whisperModuleStatus = detectPythonModuleAvailabilityViaCepNode(
    modules,
    "whisper",
    runtimeConfig,
    userExecutables,
    spawnEnv
  );
  const alignmentStatus = detectPythonModuleAvailabilityViaCepNode(
    modules,
    "whisperx",
    runtimeConfig,
    userExecutables,
    spawnEnv
  );

  if (whisperModuleStatus.available) {
    return {
      available: true,
      details: whisperModuleStatus.details,
      installedModels: installedModels.installedModels,
      modelCachePaths: installedModels.modelCachePaths,
      alignmentAvailable: alignmentStatus.available,
      alignmentDetails: alignmentStatus.details
    };
  }

  const checks: string[] = [];
  const cliCandidates = [...userExecutables, "whisper"];
  for (const command of cliCandidates) {
    const probe = runSpawn(modules, command, ["--help"], spawnEnv);
    if (probe.ok) {
      return {
        available: true,
        details: `CLI detected via ${command}${runtimeConfig ? ` (config: ${runtimeConfig.sourcePath})` : ""}`,
        installedModels: installedModels.installedModels,
        modelCachePaths: installedModels.modelCachePaths,
        alignmentAvailable: alignmentStatus.available,
        alignmentDetails: alignmentStatus.details
      };
    }

    if (probe.errorCode === "ENOENT") {
      checks.push(`${command}: missing`);
    } else if (probe.errorMessage) {
      checks.push(`${command}: ${probe.errorMessage}`);
    } else if (probe.stderr.trim()) {
      checks.push(`${command}: ${probe.stderr.trim().split("\n")[0]}`);
    } else {
      checks.push(`${command}: exit ${probe.status}`);
    }
  }

  return {
    available: false,
    details: `${checks.join(" | ")}${runtimeConfig ? ` | config=${runtimeConfig.sourcePath}` : ""}`,
    installedModels: installedModels.installedModels,
    modelCachePaths: installedModels.modelCachePaths,
    alignmentAvailable: alignmentStatus.available,
    alignmentDetails: alignmentStatus.details
  };
}

function resolveWhisperSrtPath(modules: CepNodeModules, outputDir: string, audioPath: string): string {
  // // Resolve Whisper output SRT from expected filename or directory scan fallback.
  return buildWhisperOutputPath(modules, outputDir, audioPath, "srt");
}

function readWavDurationSeconds(modules: CepNodeModules, wavPath: string): number | undefined {
  // // Estimate exported WAV duration from the RIFF byte rate so Whisper progress can move even when stderr is quiet.
  const normalizedPath = String(wavPath || "").trim();
  if (!normalizedPath || !modules.fs.existsSync(normalizedPath)) {
    return undefined;
  }

  try {
    let header = new Uint8Array();
    if (modules.fs.openSync && modules.fs.readSync && modules.fs.closeSync) {
      const fd = modules.fs.openSync(normalizedPath, "r");
      try {
        const buffer = new Uint8Array(65536);
        const bytesRead = modules.fs.readSync(fd, buffer, 0, buffer.length, 0);
        header = buffer.slice(0, bytesRead);
      } finally {
        modules.fs.closeSync(fd);
      }
    } else {
      const rawHeader = modules.fs.readFileSync(normalizedPath) as unknown as Uint8Array;
      header = rawHeader.slice(0, 65536);
    }

    const readAscii = (offset: number, length: number): string =>
      Array.from(header.slice(offset, offset + length))
        .map((byte) => String.fromCharCode(byte))
        .join("");
    const readUInt32Le = (offset: number): number =>
      ((header[offset] || 0) |
        ((header[offset + 1] || 0) << 8) |
        ((header[offset + 2] || 0) << 16) |
        ((header[offset + 3] || 0) << 24)) >>>
      0;

    if (!header || header.length < 44 || readAscii(0, 4) !== "RIFF" || readAscii(8, 4) !== "WAVE") {
      return undefined;
    }

    let byteRate = 0;
    let dataSize = 0;
    let offset = 12;
    while (offset + 8 <= header.length) {
      const chunkId = readAscii(offset, 4);
      const chunkSize = readUInt32Le(offset + 4);
      const chunkStart = offset + 8;
      if (chunkId === "fmt " && chunkStart + 12 <= header.length) {
        byteRate = readUInt32Le(chunkStart + 8);
      } else if (chunkId === "data") {
        dataSize = chunkSize;
        break;
      }
      offset = chunkStart + chunkSize + (chunkSize % 2);
    }

    const duration = byteRate > 0 && dataSize > 0 ? dataSize / byteRate : 0;
    return Number.isFinite(duration) && duration > 0 ? duration : undefined;
  } catch {
    return undefined;
  }
}

function resolveWhisperJsonPath(modules: CepNodeModules, outputDir: string, audioPath: string): string {
  // // Resolve Whisper JSON output so the planner can reuse precise word timestamps when available.
  return buildWhisperOutputPath(modules, outputDir, audioPath, "json");
}

function summarizeWhisperErrorOutput(output: string): string {
  // // Extract a meaningful root-cause line from Whisper tracebacks for actionable panel logs.
  const normalized = String(output || "").trim();
  if (!normalized) {
    return "";
  }

  const lines = normalized
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (!lines.length) {
    return "";
  }

  const preferredPatterns = [
    /SSLCertVerificationError/i,
    /CERTIFICATE_VERIFY_FAILED/i,
    /urllib\.error\.URLError/i,
    /No module named whisper/i,
    /No such file or directory.*ffmpeg/i,
    /\bffmpeg\b.*(not found|missing|failed)/i,
    /Permission denied/i,
    /File not found/i
  ];

  for (const pattern of preferredPatterns) {
    for (const line of lines) {
      if (pattern.test(line)) {
        return line;
      }
    }
  }

  return lines[lines.length - 1] || lines[0] || "";
}

function isWhisperFatalDownloadError(summary: string): boolean {
  // // Detect SSL download failures where retrying alternate launchers is noise-only.
  const normalized = String(summary || "").toLowerCase();
  if (!normalized) {
    return false;
  }
  return normalized.indexOf("sslcertverificationerror") !== -1 || normalized.indexOf("certificate_verify_failed") !== -1;
}

function normalizeWhisperOutputChunk(chunk: string | Uint8Array): string {
  // // Convert Node stream chunks into comparable text for progress/error parsing.
  if (typeof chunk === "string") {
    return chunk;
  }

  try {
    const decoder = new TextDecoder("utf-8");
    return decoder.decode(chunk);
  } catch {
    return String(chunk || "");
  }
}

function resolveBundledPythonScriptPath(
  modules: CepNodeModules,
  extensionRootPath: string,
  scriptFileName: string
): string {
  // // Resolve one bundled helper script from the installed extension payload so CEP Node can launch it via Python.
  const normalizedRoot = String(extensionRootPath || "").trim();
  if (!normalizedRoot) {
    return "";
  }

  const candidatePath = modules.path.join(normalizedRoot, "python", scriptFileName);
  return modules.fs.existsSync(candidatePath) ? candidatePath : "";
}

function normalizeWhisperEta(rawValue: string): string {
  // // Normalize tqdm ETA strings so the panel can show stable `MM:SS` or `HH:MM:SS` remaining values.
  const parts = String(rawValue || "")
    .trim()
    .split(":")
    .map((part) => Number(part));

  if (parts.length < 2 || parts.length > 3 || parts.some((part) => !Number.isFinite(part) || part < 0)) {
    return "";
  }

  const paddedParts = parts.map((part) => String(Math.floor(part)).padStart(2, "0"));
  return paddedParts.join(":");
}

function extractWhisperRemainingTime(output: string, progressMatchIndex: number): string {
  // // Read the latest tqdm ETA segment, for example `[00:42<01:18, ...]`.
  const tail = String(output || "").slice(Math.max(0, progressMatchIndex));
  const matches = Array.from(tail.matchAll(/\[[^\]\r\n<]*<\s*([0-9]+(?::[0-9]{1,2}){1,2})\s*(?:,|\])/g));
  if (!matches.length) {
    return "";
  }

  return normalizeWhisperEta(String(matches[matches.length - 1][1] || ""));
}

function extractWhisperProgressUpdate(output: string): WhisperProgressUpdate | null {
  // // Parse Whisper/tqdm stderr progress like ` 42%|...` into a panel-friendly percentage update.
  const normalized = String(output || "");
  if (!normalized) {
    return null;
  }

  const matches = Array.from(normalized.matchAll(/(\d{1,3})%\|/g));
  if (!matches.length) {
    return null;
  }

  const lastMatch = matches[matches.length - 1];
  const percent = Math.max(0, Math.min(100, Number(lastMatch[1] || 0)));
  const remaining = extractWhisperRemainingTime(normalized, Number(lastMatch.index || 0));
  return {
    percent,
    detail: `Whisper ${percent}%`,
    ...(remaining ? { remaining } : {})
  };
}

function extractCorrectedAlignProgressUpdate(output: string): WhisperProgressUpdate | null {
  // // Parse helper-script progress markers so corrected alignment can drive the same panel progress bar.
  const normalized = String(output || "");
  if (!normalized) {
    return null;
  }

  const matches = Array.from(normalized.matchAll(/SUBCREATOR_ALIGN_PROGRESS\t(\d{1,3})\t([^\r\n]*)/g));
  if (!matches.length) {
    return null;
  }

  const lastMatch = matches[matches.length - 1];
  const percent = Math.max(0, Math.min(100, Number(lastMatch[1] || 0)));
  const detail = String(lastMatch[2] || "").trim();
  return {
    percent,
    detail
  };
}

function transcribeWithWhisperViaCepNode(request: WhisperTranscriptionRequest): WhisperTranscriptionResult | null {
  // // Run Whisper with CEP Node runtime to avoid ExtendScript `system.callSystem` availability issues.
  const modules = resolveCepNodeModules();
  if (!modules) {
    return null;
  }

  const outputDir = modules.path.join(
    modules.os.tmpdir(),
    "SubCreatorWhisper",
    `run-${Date.now()}-${Math.floor(Math.random() * 100000)}`
  );
  modules.fs.mkdirSync(outputDir, { recursive: true });

  const runtimeConfig = getRuntimeConfig(modules);
  const userExecutables = discoverUserWhisperExecutables(modules, runtimeConfig);
  const spawnEnv = buildSpawnEnv(modules, userExecutables, runtimeConfig);
  const attempts: string[] = [];
  const commandCandidates = buildWhisperCommandCandidates(modules, request, outputDir, userExecutables, runtimeConfig);
  let collectedOutput = "";
  let rootCauseSummary = "";

  for (const candidate of commandCandidates) {
    const run = modules.childProcess.spawnSync(candidate.command, candidate.args, {
      encoding: "utf8",
      shell: false,
      env: spawnEnv
    });

    if (run.error) {
      const code = String(run.error.code || "");
      attempts.push(`${candidate.label}: ${String(run.error.message || run.error)}`);
      if (code === "ENOENT") {
        continue;
      }
      continue;
    }

    const attemptOutput = [String(run.stdout || ""), String(run.stderr || "")].filter(Boolean).join("\n").trim();
    if (attemptOutput && !collectedOutput) {
      collectedOutput = attemptOutput;
    }

    if (typeof run.status === "number" && run.status !== 0) {
      const summary = summarizeWhisperErrorOutput(attemptOutput);
      if (summary && !rootCauseSummary) {
        rootCauseSummary = summary;
      }
      attempts.push(`${candidate.label}: exit ${run.status}${summary ? ` (${summary})` : ""}`);
      if (isWhisperFatalDownloadError(summary)) {
        break;
      }
      continue;
    }

    const srtPath = resolveWhisperSrtPath(modules, outputDir, request.audioPath);
    if (!srtPath) {
      attempts.push(`${candidate.label}: no srt output`);
      continue;
    }

    const srtText = String(modules.fs.readFileSync(srtPath, "utf8") || "");
    if (!srtText.trim()) {
      attempts.push(`${candidate.label}: empty srt output`);
      continue;
    }

    const jsonPath = resolveWhisperJsonPath(modules, outputDir, request.audioPath);
    const jsonText = jsonPath && modules.fs.existsSync(jsonPath) ? String(modules.fs.readFileSync(jsonPath, "utf8") || "") : "";

    const result = {
      srtText,
      jsonText: jsonText.trim() ? jsonText : undefined,
      model: request.model?.trim() || "base",
      audioPath: request.audioPath,
      commandOutput: attemptOutput
    };

    cleanupWhisperOutputFiles(modules, outputDir, request.audioPath);
    return result;
  }

  let installHint = "";
  let runtimeHint = "";
  if (detectWindowsRuntime()) {
    if (runtimeConfig?.pythonPath) {
      installHint = `Install command: ${runtimeConfig.pythonPath} -m pip install --user -U openai-whisper`;
    } else if (runtimeConfig?.pythonCommand) {
      installHint = `Install command: ${runtimeConfig.pythonCommand} -m pip install --user -U openai-whisper`;
    } else {
      installHint = "Install command: py -m pip install --user -U openai-whisper";
    }
  } else {
    let interpreterHint = "";
    if (runtimeConfig?.pythonPath) {
      interpreterHint = runtimeConfig.pythonPath;
    } else if (runtimeConfig?.pythonCommand) {
      interpreterHint = runtimeConfig.pythonCommand;
    }
    if (!interpreterHint) {
      for (const executablePath of userExecutables) {
        interpreterHint = resolveWhisperInterpreterFromScript(modules, executablePath);
        if (interpreterHint) {
          break;
        }
      }
    }

    if (interpreterHint) {
      installHint = `Install command: ${interpreterHint} -m pip install --user -U openai-whisper`;
    } else {
      installHint = "Install command: python3 -m pip install --user -U openai-whisper";
    }
  }
  if (rootCauseSummary && /sslcertverificationerror|certificate_verify_failed|urllib\.error\.urlerror/i.test(rootCauseSummary)) {
    runtimeHint =
      "Whisper model access failed due TLS/SSL certificate validation. Configure trusted certs/proxy for Python, or copy the model into the Whisper cache manually.";
  }
  throw new Error(
    `Unable to execute Whisper CLI from CEP runtime. Attempts: ${attempts.join(" | ") || "none"}. ${installHint}. ${
      runtimeConfig ? `Runtime config: ${runtimeConfig.sourcePath}. ` : ""
    }${runtimeHint ? `${runtimeHint}. ` : ""}${collectedOutput || ""}`
  );
}

async function transcribeWithWhisperViaCepNodeAsync(
  request: WhisperTranscriptionRequest,
  onProgress?: (update: WhisperProgressUpdate) => void
): Promise<WhisperTranscriptionResult | null> {
  // // Stream Whisper execution through CEP Node so the panel can update progress while transcription is running.
  const modules = resolveCepNodeModules();
  if (!modules) {
    return null;
  }

  if (typeof modules.childProcess.spawn !== "function") {
    return transcribeWithWhisperViaCepNode(request);
  }

  const outputDir = modules.path.join(
    modules.os.tmpdir(),
    "SubCreatorWhisper",
    `run-${Date.now()}-${Math.floor(Math.random() * 100000)}`
  );
  modules.fs.mkdirSync(outputDir, { recursive: true });

  const runtimeConfig = getRuntimeConfig(modules);
  const userExecutables = discoverUserWhisperExecutables(modules, runtimeConfig);
  const spawnEnv = buildSpawnEnv(modules, userExecutables, runtimeConfig);
  const attempts: string[] = [];
  const commandCandidates = buildWhisperCommandCandidates(modules, request, outputDir, userExecutables, runtimeConfig);
  let collectedOutput = "";
  let rootCauseSummary = "";

  for (const candidate of commandCandidates) {
    let latestProgressPercent = -1;
    const attemptChunks: string[] = [];
    let attemptTailOutput = "";
    const attemptResult = await new Promise<{ code: number | null; error?: { message?: string; code?: string } }>((resolve) => {
      // // Attach stdout/stderr listeners before Whisper starts so tqdm progress can feed the panel immediately.
      let settled = false;
      const detached = !detectWindowsRuntime();
      const child = modules.childProcess.spawn?.(candidate.command, candidate.args, {
        shell: false,
        detached,
        env: spawnEnv
      });
      if (!child) {
        resolve({
          code: null,
          error: {
            message: "child_process.spawn unavailable"
          }
        });
        return;
      }
      const activeJob = registerActiveCepJob(child, candidate.label, detached);

      const handleChunk = (chunk: string | Uint8Array): void => {
        // // Keep the latest progress percentage from Whisper stderr without flooding the UI with duplicate values.
        const normalizedChunk = normalizeWhisperOutputChunk(chunk);
        if (!normalizedChunk) {
          return;
        }
        attemptChunks.push(normalizedChunk);
        attemptTailOutput = `${attemptTailOutput}${normalizedChunk}`.slice(-4096);
        const progress = extractWhisperProgressUpdate(attemptTailOutput);
        if (!progress || progress.percent <= latestProgressPercent) {
          return;
        }
        latestProgressPercent = progress.percent;
        onProgress?.(progress);
      };

      child.stdout?.on("data", handleChunk);
      child.stderr?.on("data", handleChunk);
      child.on("error", (error) => {
        clearActiveCepJob(child);
        if (settled) {
          return;
        }
        settled = true;
        if (activeJob.cancelRequested) {
          resolve({
            code: null,
            error: {
              message: SUBCREATOR_CANCELLED_JOB_CODE,
              code: SUBCREATOR_CANCELLED_JOB_CODE
            }
          });
          return;
        }
        resolve({
          code: null,
          error: error && typeof error === "object" ? (error as { message?: string; code?: string }) : { message: String(error) }
        });
      });
      child.on("close", (value) => {
        clearActiveCepJob(child);
        if (settled) {
          return;
        }
        settled = true;
        if (activeJob.cancelRequested) {
          resolve({
            code: null,
            error: {
              message: SUBCREATOR_CANCELLED_JOB_CODE,
              code: SUBCREATOR_CANCELLED_JOB_CODE
            }
          });
          return;
        }
        resolve({
          code: typeof value === "number" ? value : null
        });
      });
    });

    if (attemptResult.error) {
      if (isCancelledJobError(attemptResult.error.message || attemptResult.error.code || "")) {
        throw createCancelledJobError();
      }
      const code = String(attemptResult.error.code || "");
      attempts.push(`${candidate.label}: ${String(attemptResult.error.message || attemptResult.error)}`);
      if (code === "ENOENT") {
        continue;
      }
      continue;
    }

    const attemptOutput = attemptChunks.join("").trim();
    if (attemptOutput && !collectedOutput) {
      collectedOutput = attemptOutput;
    }

    if (typeof attemptResult.code === "number" && attemptResult.code !== 0) {
      const summary = summarizeWhisperErrorOutput(attemptOutput);
      if (summary && !rootCauseSummary) {
        rootCauseSummary = summary;
      }
      attempts.push(`${candidate.label}: exit ${attemptResult.code}${summary ? ` (${summary})` : ""}`);
      if (isWhisperFatalDownloadError(summary)) {
        break;
      }
      continue;
    }

    const srtPath = resolveWhisperSrtPath(modules, outputDir, request.audioPath);
    if (!srtPath) {
      attempts.push(`${candidate.label}: no srt output`);
      continue;
    }

    const srtText = String(modules.fs.readFileSync(srtPath, "utf8") || "");
    if (!srtText.trim()) {
      attempts.push(`${candidate.label}: empty srt output`);
      continue;
    }

    const jsonPath = resolveWhisperJsonPath(modules, outputDir, request.audioPath);
    const jsonText = jsonPath && modules.fs.existsSync(jsonPath) ? String(modules.fs.readFileSync(jsonPath, "utf8") || "") : "";

    onProgress?.({
      percent: 100,
      detail: "Whisper 100%"
    });

    const result = {
      srtText,
      jsonText: jsonText.trim() ? jsonText : undefined,
      model: request.model?.trim() || "base",
      audioPath: request.audioPath,
      commandOutput: attemptOutput
    };

    cleanupWhisperOutputFiles(modules, outputDir, request.audioPath);
    return result;
  }

  let installHint = "";
  let runtimeHint = "";
  if (detectWindowsRuntime()) {
    if (runtimeConfig?.pythonPath) {
      installHint = `Install command: ${runtimeConfig.pythonPath} -m pip install --user -U openai-whisper`;
    } else if (runtimeConfig?.pythonCommand) {
      installHint = `Install command: ${runtimeConfig.pythonCommand} -m pip install --user -U openai-whisper`;
    } else {
      installHint = "Install command: py -m pip install --user -U openai-whisper";
    }
  } else {
    let interpreterHint = "";
    if (runtimeConfig?.pythonPath) {
      interpreterHint = runtimeConfig.pythonPath;
    } else if (runtimeConfig?.pythonCommand) {
      interpreterHint = runtimeConfig.pythonCommand;
    }
    if (!interpreterHint) {
      for (const executablePath of userExecutables) {
        interpreterHint = resolveWhisperInterpreterFromScript(modules, executablePath);
        if (interpreterHint) {
          break;
        }
      }
    }

    if (interpreterHint) {
      installHint = `Install command: ${interpreterHint} -m pip install --user -U openai-whisper`;
    } else {
      installHint = "Install command: python3 -m pip install --user -U openai-whisper";
    }
  }
  if (rootCauseSummary && /sslcertverificationerror|certificate_verify_failed|urllib\.error\.urlerror/i.test(rootCauseSummary)) {
    runtimeHint =
      "Whisper model access failed due TLS/SSL certificate validation. Configure trusted certs/proxy for Python, or copy the model into the Whisper cache manually.";
  }
  throw new Error(
    `Unable to execute Whisper CLI from CEP runtime. Attempts: ${attempts.join(" | ") || "none"}. ${installHint}. ${
      runtimeConfig ? `Runtime config: ${runtimeConfig.sourcePath}. ` : ""
    }${runtimeHint ? `${runtimeHint}. ` : ""}${collectedOutput || ""}`
  );
}

export async function getWhisperRuntimeStatus(): Promise<WhisperRuntimeStatus> {
  // // Expose runtime detection to UI so unavailable Whisper source can be hidden safely.
  return detectWhisperAvailabilityViaCepNode();
}

export async function pingHost(): Promise<string> {
  // // Validate the bridge wiring with a lightweight host call.
  return evalScript("subcreator_ping()");
}

export async function applyCaptionPlan(payload: HostApplyPayload): Promise<string> {
  // // Send JSON payload as URI-encoded text to avoid quote escaping edge-cases.
  const encodedPayload = encodeURIComponent(JSON.stringify(payload));
  return evalHostJsonRaw(`subcreator_apply_captions("${escapeForJsx(encodedPayload)}")`);
}

export async function applyNativeSubtitlePlan(payload: HostApplyPayload): Promise<string> {
  // // Send planned cues to ExtendScript so Premiere can import them as one native subtitle track from SRT.
  const encodedPayload = encodeURIComponent(JSON.stringify(payload));
  return evalHostJsonRaw(`subcreator_apply_native_subtitles("${escapeForJsx(encodedPayload)}")`);
}

export async function readTextFileFromHost(filePath: string): Promise<string> {
  // // Read subtitle files through host to avoid browser file access limitations.
  const encoded = encodeURIComponent(filePath);
  const response = await evalHostJson<{ text: string }>(`subcreator_read_text_file("${escapeForJsx(encoded)}")`);

  if (!response.ok) {
    throw new Error(response.error ?? "Unable to read file from host.");
  }

  return String(response.data?.text ?? "");
}

export async function pickSrtPath(): Promise<string> {
  // // Open host-native picker for selecting an SRT file path.
  const response = await evalHostJson<{ path: string }>("subcreator_pick_srt_file()");
  if (!response.ok) {
    throw new Error(response.error ?? "SRT picker failed.");
  }

  return String(response.data?.path ?? "");
}

export async function getActiveSequenceRange(): Promise<ActiveSequenceRangeResult> {
  // // Read the active sequence In/Out values directly from ExtendScript so non-Whisper sources can respect the same range control.
  const response = await evalHostJson<ActiveSequenceRangeResult>("subcreator_get_active_sequence_range()");
  if (!response.ok) {
    return {
      fallbackReason: "Unable to read active sequence In/Out range; using entire sequence.",
      hostError: response.error ?? "Unable to read active sequence range.",
      debug: response.debug
    };
  }

  return {
    rangeStartSeconds: Number.isFinite(Number(response.data?.rangeStartSeconds))
      ? Number(response.data?.rangeStartSeconds)
      : undefined,
    rangeEndSeconds: Number.isFinite(Number(response.data?.rangeEndSeconds))
      ? Number(response.data?.rangeEndSeconds)
      : undefined,
    sequenceName: String(response.data?.sequenceName ?? "")
  };
}

export async function pickCorrectedTranscriptPath(): Promise<string> {
  // // Open host-native picker for corrected transcript files used by WhisperX alignment.
  const response = await evalHostJson<{ path: string }>("subcreator_pick_corrected_transcript_file()");
  if (!response.ok) {
    throw new Error(response.error ?? "Corrected transcript picker failed.");
  }

  return String(response.data?.path ?? "");
}

export async function exportActiveSequenceAudioForWhisper(
  rangeMode: WhisperSequenceRangeMode = "entire_sequence",
  extensionRootPath = ""
): Promise<WhisperSequenceExportResult> {
  // // Export the active sequence audible mix to a temporary WAV file so Whisper can analyze the current edit directly.
  const modules = resolveCepNodeModules();
  if (!modules) {
    throw new Error("CEP Node runtime unavailable. Unable to export active sequence audio for Whisper.");
  }

  const presetPath = detectWhisperSequencePresetPathViaCepNode(modules, extensionRootPath);
  if (!presetPath) {
    throw new Error("Unable to locate the WAV preset for Whisper sequence export.");
  }

  const outputDir = modules.path.join(modules.os.tmpdir(), "SubCreatorWhisperSequence");
  modules.fs.mkdirSync(outputDir, { recursive: true });

  const outputBase = `subcreator-sequence-${Date.now()}`;
  const exportDebug: Record<string, unknown> = {
    rangeMode,
    extensionRootPath,
    presetPath,
    outputDir,
    platform: typeof navigator !== "undefined" ? navigator.platform : "",
    userAgent: typeof navigator !== "undefined" ? navigator.userAgent : "",
    hostEnvironment: readCepHostEnvironmentDebug(),
    cepHostAvailable: Boolean(window.__adobe_cep__),
    cepNodeAvailable: Boolean(modules),
    attempts: []
  };
  const runExport = async (
    exportMethod: "premiere_direct" | "media_encoder"
  ): Promise<{
    response: HostJsonResponse<WhisperSequenceExportResult>;
    outputPath: string;
    elapsedMs: number;
    fileSnapshot: Record<string, unknown>;
  }> => {
    // // Invoke each exporter independently so Premiere cannot abort the fallback with the direct evalScript call.
    const outputPath = modules.path.join(outputDir, `${outputBase}-${exportMethod}.wav`);
    const encodedPayload = encodeURIComponent(
      JSON.stringify({
        outputPath,
        presetPath,
        rangeMode,
        exportMode: exportMethod
      })
    );
    const startedAt = Date.now();
    const response = await evalHostJson<WhisperSequenceExportResult>(
      `subcreator_export_active_sequence_audio("${escapeForJsx(encodedPayload)}")`
    );
    const elapsedMs = Date.now() - startedAt;
    const fileSnapshot = readCepFileSnapshot(modules, outputPath);
    const attempts = Array.isArray(exportDebug.attempts) ? exportDebug.attempts : [];
    attempts.push({
      exportMethod,
      outputPath,
      elapsedMs,
      ok: Boolean(response.ok),
      error: response.error,
      hostDebug: response.debug ?? response.data?.debug,
      fileSnapshot
    });
    exportDebug.attempts = attempts;
    return { response, outputPath, elapsedMs, fileSnapshot };
  };
  try {
    // // Probe export capabilities separately so a fatal export call still leaves useful Premiere API diagnostics.
    const capabilities = await evalHostJson<Record<string, unknown>>("subcreator_get_sequence_export_capabilities()");
    exportDebug.capabilities = capabilities.ok ? capabilities.data : { error: capabilities.error, debug: capabilities.debug };
  } catch (error) {
    exportDebug.capabilities = { error: String(error) };
  }

  const recoverCompletedExport = async (
    exportMethod: "premiere_direct" | "media_encoder",
    candidatePath: string,
    timeoutMs: number
  ): Promise<HostJsonResponse<WhisperSequenceExportResult> | null> => {
    // // Some Premiere export APIs can start a WAV export and still make evalScript return "EvalScript error".
    if (!(await waitForStableCepFile(modules, candidatePath, timeoutMs, 3))) {
      return null;
    }

    const activeRange = await getActiveSequenceRange().catch((): ActiveSequenceRangeResult => ({}));
    const attempts = Array.isArray(exportDebug.attempts) ? exportDebug.attempts : [];
    attempts.push({
      exportMethod,
      recoveredAfterInvalidHostResponse: true,
      recoveryWaitMs: timeoutMs,
      fileSnapshot: readCepFileSnapshot(modules, candidatePath)
    });
    exportDebug.attempts = attempts;

    return {
      ok: true,
      data: {
        audioPath: candidatePath,
        presetPath,
        exportMethod,
        sequenceName: String(activeRange.sequenceName ?? ""),
        rangeStartSeconds: rangeMode === "in_out" ? activeRange.rangeStartSeconds : undefined,
        rangeEndSeconds: rangeMode === "in_out" ? activeRange.rangeEndSeconds : undefined
      }
    };
  };

  const directAttempt = await runExport("premiere_direct");
  let response = directAttempt.response;
  let outputPath = directAttempt.outputPath;
  if (!response.ok) {
    const recoveredDirect = await recoverCompletedExport("premiere_direct", outputPath, 30000);
    if (recoveredDirect) {
      response = recoveredDirect;
    }
  }

  if (!response.ok) {
    // // Fall back to AME on its own WAV path when Premiere direct export fails or returns an opaque evalScript error.
    const mediaEncoderAttempt = await runExport("media_encoder");
    response = mediaEncoderAttempt.response;
    outputPath = mediaEncoderAttempt.outputPath;
    if (!response.ok) {
      const recoveredMediaEncoder = await recoverCompletedExport("media_encoder", outputPath, 120000);
      if (recoveredMediaEncoder) {
        response = recoveredMediaEncoder;
      }
    }
  }

  if (!response.ok) {
    exportDebug.finalFileSnapshot = readCepFileSnapshot(modules, outputPath);
    throw new Error(
      buildWhisperExportErrorMessage(response.error ?? "Unable to export active sequence audio for Whisper.", exportDebug)
    );
  }

  const resolvedAudioPath = String(response.data?.audioPath ?? outputPath);
  return {
    audioPath: resolvedAudioPath,
    presetPath: String(response.data?.presetPath ?? presetPath),
    exportMethod: response.data?.exportMethod ?? "premiere_direct",
    sequenceName: String(response.data?.sequenceName ?? ""),
    rangeStartSeconds: Number.isFinite(Number(response.data?.rangeStartSeconds))
      ? Number(response.data?.rangeStartSeconds)
      : undefined,
    rangeEndSeconds: Number.isFinite(Number(response.data?.rangeEndSeconds)) ? Number(response.data?.rangeEndSeconds) : undefined,
    audioDurationSeconds:
      Number.isFinite(Number(response.data?.audioDurationSeconds)) && Number(response.data?.audioDurationSeconds) > 0
        ? Number(response.data?.audioDurationSeconds)
        : readWavDurationSeconds(modules, resolvedAudioPath),
    debug: {
      ...exportDebug,
      hostDebug: response.data?.debug,
      finalFileSnapshot: readCepFileSnapshot(modules, resolvedAudioPath)
    }
  };
}

async function waitForStableCepFile(
  modules: CepNodeModules,
  filePath: string,
  timeoutMs: number,
  stablePasses: number
): Promise<boolean> {
  // // Confirm a non-empty export has stopped growing before accepting a WAV whose host response was lost.
  const deadline = Date.now() + Math.max(1000, timeoutMs);
  let lastSize = -1;
  let stableCount = 0;

  while (Date.now() < deadline) {
    try {
      if (modules.fs.existsSync(filePath)) {
        const currentSize = Number(modules.fs.statSync(filePath).size || 0);
        if (currentSize > 0 && currentSize === lastSize) {
          stableCount += 1;
          if (stableCount >= Math.max(2, stablePasses)) {
            return true;
          }
        } else {
          lastSize = currentSize;
          stableCount = 0;
        }
      }
    } catch {
      // // Retry transient file locks while Premiere finalizes its direct WAV export.
    }

    await new Promise<void>((resolve) => setTimeout(resolve, 250));
  }

  return false;
}

export async function readPremiereTemplateTextPayloads(mogrtPath: string): Promise<PremiereTemplateTextPayload[]> {
  // // Extract Premiere-authored Source Text payloads from the selected template so the host can preserve font/style when replacing subtitle text.
  const modules = resolveCepNodeModules();
  if (!modules) {
    return [];
  }
  return extractPremiereTemplateTextPayloads(modules, mogrtPath);
}

export async function deleteTemporaryWhisperAudio(filePath: string): Promise<void> {
  // // Remove temporary exported audio files once Whisper analysis has consumed them.
  const normalizedPath = String(filePath || "").trim();
  if (!normalizedPath) {
    return;
  }

  const modules = resolveCepNodeModules();
  if (!modules) {
    return;
  }

  if (!modules.fs.existsSync(normalizedPath)) {
    return;
  }

  try {
    modules.fs.unlinkSync(normalizedPath);
  } catch {
    // // Ignore cleanup failures because they should not block subtitle generation.
  }
}

function normalizeVisualPropertyList(data: unknown): SelectedMogrtVisualPropertyList {
  // // Sanitize host payload shape before rendering dynamic visual controls.
  const payload = (data && typeof data === "object" ? (data as Record<string, unknown>) : {}) || {};
  const rawProperties = Array.isArray(payload.properties) ? payload.properties : [];
  const properties: SelectedMogrtVisualProperty[] = [];

  for (const rawProperty of rawProperties) {
    const item = rawProperty && typeof rawProperty === "object" ? (rawProperty as Record<string, unknown>) : null;
    if (!item) {
      continue;
    }

    const path = String(item.path || "").trim();
    const displayName = String(item.displayName || "").trim();
    const groupPath = String(item.groupPath || "").trim();
    const valueTypeRaw = String(item.valueType || "string").trim().toLowerCase();
    const valueType: SelectedMogrtVisualProperty["valueType"] =
      valueTypeRaw === "number" || valueTypeRaw === "boolean" || valueTypeRaw === "json" ? valueTypeRaw : "string";
    const controlKindRaw = String(item.controlKind || valueType).trim().toLowerCase();
    const controlKind: SelectedMogrtVisualProperty["controlKind"] =
      controlKindRaw === "slider" ||
      controlKindRaw === "number" ||
      controlKindRaw === "checkbox" ||
      controlKindRaw === "color" ||
      controlKindRaw === "vector" ||
      controlKindRaw === "select" ||
      controlKindRaw === "text" ||
      controlKindRaw === "json"
        ? controlKindRaw
        : "string";
    if (!path || !displayName) {
      continue;
    }

    let value: string | number | boolean = "";
    if (valueType === "number") {
      value = Number(item.value || 0);
    } else if (valueType === "boolean") {
      value = Boolean(item.value);
    } else {
      value = String(item.value ?? "");
    }

    properties.push({
      path,
      displayName,
      groupPath,
      valueType,
      controlKind,
      cloneOnlyWhenDirty: item.cloneOnlyWhenDirty === true,
      fontToken: typeof item.fontToken === "string" ? item.fontToken : undefined,
      options: Array.isArray(item.options)
        ? item.options
            .map((option) => {
              const optionRecord = option && typeof option === "object" ? (option as Record<string, unknown>) : null;
              if (!optionRecord) {
                return null;
              }

              const rawOptionValue = optionRecord.value;
              const optionValue =
                typeof rawOptionValue === "number" || typeof rawOptionValue === "string"
                  ? rawOptionValue
                  : String(rawOptionValue ?? "");
              const optionLabel = String(optionRecord.label ?? optionValue).trim() || String(optionValue);
              return { value: optionValue, label: optionLabel };
            })
            .filter((option): option is { value: number | string; label: string } => Boolean(option))
        : undefined,
      styleOptionsByFamily:
        item.styleOptionsByFamily && typeof item.styleOptionsByFamily === "object"
          ? Object.entries(item.styleOptionsByFamily as Record<string, unknown>).reduce<Record<string, string[]>>(
              (accumulator, [family, rawStyles]) => {
                if (!Array.isArray(rawStyles)) {
                  return accumulator;
                }
                const styles = rawStyles
                  .map((entry) => String(entry || "").trim())
                  .filter((entry) => entry.length > 0);
                if (styles.length > 0) {
                  accumulator[String(family)] = styles;
                }
                return accumulator;
              },
              {}
            )
          : undefined,
      vectorScale: Array.isArray(item.vectorScale)
        ? item.vectorScale.map((entry) => Number(entry)).filter((entry) => Number.isFinite(entry))
        : undefined,
      vectorMode: typeof item.vectorMode === "string" ? item.vectorMode : undefined,
      minValue: Number.isFinite(Number(item.minValue)) ? Number(item.minValue) : undefined,
      maxValue: Number.isFinite(Number(item.maxValue)) ? Number(item.maxValue) : undefined,
      stepValue: Number.isFinite(Number(item.stepValue)) ? Number(item.stepValue) : undefined,
      value
    });
  }

  const rawDebug = payload.debug && typeof payload.debug === "object" ? (payload.debug as Record<string, unknown>) : null;
  const normalizedDebug: SelectedMogrtVisualDebug | undefined = rawDebug
    ? {
        ...rawDebug,
        sequenceWidth: Number.isFinite(Number(rawDebug.sequenceWidth)) ? Number(rawDebug.sequenceWidth) : undefined,
        sequenceHeight: Number.isFinite(Number(rawDebug.sequenceHeight)) ? Number(rawDebug.sequenceHeight) : undefined,
        componentCount: Number.isFinite(Number(rawDebug.componentCount)) ? Number(rawDebug.componentCount) : undefined,
        vectorCount: Number.isFinite(Number(rawDebug.vectorCount)) ? Number(rawDebug.vectorCount) : undefined,
        colorCount: Number.isFinite(Number(rawDebug.colorCount)) ? Number(rawDebug.colorCount) : undefined,
        selectCount: Number.isFinite(Number(rawDebug.selectCount)) ? Number(rawDebug.selectCount) : undefined,
        components: Array.isArray(rawDebug.components)
          ? rawDebug.components
              .map((entry) => {
                const item = entry && typeof entry === "object" ? (entry as Record<string, unknown>) : null;
                if (!item) {
                  return null;
                }
                const index = Number(item.index);
                const name = String(item.name || "").trim();
                const propertyCount = Number(item.propertyCount);
                if (!Number.isFinite(index)) {
                  return null;
                }
                return {
                  index,
                  name: name || `Component ${index + 1}`,
                  propertyCount: Number.isFinite(propertyCount) ? propertyCount : 0
                };
              })
              .filter((entry): entry is SelectedMogrtVisualComponentDebug => Boolean(entry))
          : undefined
      }
    : undefined;

  return {
    selectedCount: Number(payload.selectedCount || 0),
    editableCount: Number(payload.editableCount || properties.length),
    properties,
    debug: normalizedDebug
  };
}

function normalizeSelectedMogrtTextItemList(data: unknown): SelectedMogrtTextItemList {
  // // Sanitize host text-selection payload before rendering editable subtitle blocks in the Text tab.
  const payload = (data && typeof data === "object" ? (data as Record<string, unknown>) : {}) || {};
  const rawItems = Array.isArray(payload.items) ? payload.items : [];
  const items: SelectedMogrtTextItem[] = [];

  for (const rawItem of rawItems) {
    const item = rawItem && typeof rawItem === "object" ? (rawItem as Record<string, unknown>) : null;
    if (!item) {
      continue;
    }

    const startSeconds = Number(item.startSeconds);
    const endSeconds = Number(item.endSeconds);
    items.push({
      selectionIndex: Number(item.selectionIndex || 0),
      videoTrackIndex: Number(item.videoTrackIndex || 0),
      startSeconds: Number.isFinite(startSeconds) ? startSeconds : 0,
      endSeconds: Number.isFinite(endSeconds) ? endSeconds : 0,
      text: String(item.text || "").trim(),
      clipName: String(item.clipName || "").trim()
    });
  }

  return {
    selectedCount: Number(payload.selectedCount || items.length),
    sameTrack: payload.sameTrack !== false,
    videoTrackIndex: Number.isFinite(Number(payload.videoTrackIndex)) ? Number(payload.videoTrackIndex) : undefined,
    projectDocumentId: typeof payload.projectDocumentId === "string" ? payload.projectDocumentId : undefined,
    projectPath: typeof payload.projectPath === "string" ? payload.projectPath : undefined,
    sequenceID: typeof payload.sequenceID === "string" ? payload.sequenceID : undefined,
    sequenceName: typeof payload.sequenceName === "string" ? payload.sequenceName : undefined,
    signature: String(payload.signature || ""),
    items
  };
}

export async function readSelectedMogrtVisualProperties(): Promise<SelectedMogrtVisualPropertyList> {
  // // Request editable MOGRT properties from selected timeline clips.
  const response = await evalHostJson<SelectedMogrtVisualPropertyList>("subcreator_list_selected_mogrt_properties()");
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to read selected MOGRT properties.");
  }

  return normalizeVisualPropertyList(response.data);
}

export async function readSelectedMogrtTextItems(): Promise<SelectedMogrtTextItemList> {
  // // Request selected MOGRT text blocks from host for Text tab editing.
  const response = await evalHostJson<SelectedMogrtTextItemList>("subcreator_list_selected_mogrt_text_items()");
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to read selected MOGRT text items.");
  }

  return normalizeSelectedMogrtTextItemList(response.data);
}

export async function getSelectedMogrtCount(): Promise<number> {
  // // Return selected MOGRT count to drive panel-side apply progress steps.
  const response = await evalHostJson<{ selectedCount: number }>("subcreator_get_selected_mogrt_count()");
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to read selected MOGRT count.");
  }

  return Number(response.data?.selectedCount || 0);
}

export async function readSystemFontCatalog(): Promise<SystemFontCatalog> {
  // // Read local system font families/styles from CEP Node runtime for dropdown fallback.
  const detected = detectSystemFontCatalogViaCepNode();
  if (!detected) {
    return {
      available: false,
      source: "unavailable",
      details: "CEP Node runtime unavailable",
      families: [],
      stylesByFamily: {},
      fontTokensByFamilyStyle: {}
    };
  }

  return {
    available: detected.available,
    source: detected.source,
    details: detected.details,
    families: detected.families.slice(),
    stylesByFamily: Object.keys(detected.stylesByFamily).reduce<Record<string, string[]>>((accumulator, family) => {
      accumulator[family] = detected.stylesByFamily[family].slice();
      return accumulator;
    }, {}),
    fontTokensByFamilyStyle: Object.keys(detected.fontTokensByFamilyStyle || {}).reduce<Record<string, Record<string, string>>>(
      (accumulator, family) => {
        accumulator[family] = Object.entries(detected.fontTokensByFamilyStyle[family] || {}).reduce<Record<string, string>>(
          (styleAccumulator, [style, token]) => {
            styleAccumulator[style] = String(token || "");
            return styleAccumulator;
          },
          {}
        );
        return accumulator;
      },
      {}
    )
  };
}

export async function readInstalledMogrtCatalog(extensionRootPath: string): Promise<InstalledMogrtCatalog> {
  // // Read installed gallery templates from extension disk so user-added folders/files appear without rebuild.
  const detected = readInstalledMogrtCatalogViaCepNode(extensionRootPath);
  if (!detected) {
    return {
      available: false,
      source: "unavailable",
      details: "CEP Node runtime unavailable",
      templatesRoot: "",
      groups: [],
      templates: []
    };
  }

  return {
    available: detected.available,
    source: detected.source,
    details: detected.details,
    templatesRoot: String(detected.templatesRoot || ""),
    groups: detected.groups.slice(),
    templates: detected.templates.map((template) => ({
      ...template
    }))
  };
}

export async function openInstalledMogrtFolder(extensionRootPath: string): Promise<string> {
  // // Open installed `templates/mogrt` folder in Finder/Explorer so users can drop new templates manually.
  const modules = resolveCepNodeModules();
  if (!modules) {
    throw new Error("CEP Node runtime unavailable. Unable to open MOGRT folder.");
  }

  const normalizedExtensionRoot = String(extensionRootPath || "").trim();
  if (!normalizedExtensionRoot) {
    throw new Error("Extension root path unavailable. Unable to open MOGRT folder.");
  }

  const templatesRoot = modules.path.join(normalizedExtensionRoot, "templates", "mogrt");
  modules.fs.mkdirSync(templatesRoot, { recursive: true });

  const commandCandidates = detectWindowsRuntime()
    ? [{ command: "explorer", args: [templatesRoot] }]
    : [
        { command: "/usr/bin/open", args: [templatesRoot] },
        { command: "open", args: [templatesRoot] }
      ];

  for (const candidate of commandCandidates) {
    const result = modules.childProcess.spawnSync(candidate.command, candidate.args, {
      encoding: "utf8",
      timeout: 15000,
      env: modules.process.env
    });
    if (folderOpenCommandSucceeded(result, detectWindowsRuntime())) {
      return templatesRoot;
    }
  }

  throw new Error(`Unable to open installed MOGRT folder: ${templatesRoot}`);
}

export async function openWhisperModelsFolder(modelCachePaths: string[] = []): Promise<string> {
  // // Open the local Whisper model-cache folder in Finder/Explorer so users can add or inspect installed `.pt` files.
  const modules = resolveCepNodeModules();
  if (!modules) {
    throw new Error("CEP Node runtime unavailable. Unable to open Whisper models folder.");
  }

  const modelsRoot = resolvePreferredWhisperModelCacheDirectory(modules, modelCachePaths);
  modules.fs.mkdirSync(modelsRoot, { recursive: true });

  const commandCandidates = detectWindowsRuntime()
    ? [{ command: "explorer", args: [modelsRoot] }]
    : [
        { command: "/usr/bin/open", args: [modelsRoot] },
        { command: "open", args: [modelsRoot] }
      ];

  for (const candidate of commandCandidates) {
    const result = modules.childProcess.spawnSync(candidate.command, candidate.args, {
      encoding: "utf8",
      timeout: 15000,
      env: modules.process.env
    });
    if (folderOpenCommandSucceeded(result, detectWindowsRuntime())) {
      return modelsRoot;
    }
  }

  throw new Error(`Unable to open Whisper models folder: ${modelsRoot}`);
}

export async function openExternalUrl(url: string): Promise<void> {
  // // Open release/download links through CEP first because regular HTML anchors are unreliable inside Premiere panels.
  const normalizedUrl = String(url || "").trim();
  if (!normalizedUrl) {
    throw new Error("External URL unavailable.");
  }

  const cepOpenUrl = window.cep?.util?.openURLInDefaultBrowser;
  if (typeof cepOpenUrl === "function") {
    cepOpenUrl(normalizedUrl);
    return;
  }

  const modules = resolveCepNodeModules();
  if (modules) {
    const commandCandidates = detectWindowsRuntime()
      ? [
          { command: "cmd", args: ["/c", "start", "", normalizedUrl] },
          { command: "explorer", args: [normalizedUrl] }
        ]
      : [
          { command: "/usr/bin/open", args: [normalizedUrl] },
          { command: "open", args: [normalizedUrl] },
          { command: "xdg-open", args: [normalizedUrl] }
        ];

    for (const candidate of commandCandidates) {
      const result = modules.childProcess.spawnSync(candidate.command, candidate.args, {
        encoding: "utf8",
        timeout: 15000,
        env: modules.process.env
      });
      if (!result.error && (result.status === 0 || result.status === null)) {
        return;
      }
    }
  }

  const openedWindow = typeof window.open === "function" ? window.open(normalizedUrl, "_blank", "noopener") : null;
  if (openedWindow) {
    return;
  }

  window.location.href = normalizedUrl;
}

export async function applyVisualPropertiesToSelectedMogrts(
  changes: Array<{
    path: string;
    valueType: SelectedMogrtVisualProperty["valueType"];
    controlKind: SelectedMogrtVisualProperty["controlKind"];
    vectorScale?: number[];
    fontToken?: string;
    value: string | number | boolean;
  }>,
  options?: {
    clipStartIndex?: number;
    clipEndIndex?: number;
  }
): Promise<ApplyVisualPropertiesResult> {
  // // Send edited property payload to host and apply values on selected MOGRT clips.
  const encodedPayload = encodeURIComponent(
    JSON.stringify({
      changes,
      clipStartIndex: Number.isFinite(Number(options?.clipStartIndex)) ? Number(options?.clipStartIndex) : undefined,
      clipEndIndex: Number.isFinite(Number(options?.clipEndIndex)) ? Number(options?.clipEndIndex) : undefined
    })
  );
  const response = await evalHostJson<ApplyVisualPropertiesResult>(
    `subcreator_apply_selected_mogrt_properties("${escapeForJsx(encodedPayload)}")`
  );
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to apply selected MOGRT properties.");
  }

  return {
    selectedCount: Number(response.data?.selectedCount || 0),
    processedClipCount: Number(response.data?.processedClipCount || 0),
    clipStartIndex: Number.isFinite(Number(response.data?.clipStartIndex)) ? Number(response.data?.clipStartIndex) : undefined,
    clipEndIndex: Number.isFinite(Number(response.data?.clipEndIndex)) ? Number(response.data?.clipEndIndex) : undefined,
    updatedCount: Number(response.data?.updatedCount || 0),
    failedCount: Number(response.data?.failedCount || 0),
    debug: Array.isArray(response.data?.debug) ? response.data.debug.map((line) => String(line)) : undefined
  };
}

export async function applySelectedMogrtTextItems(payload: TextEditorApplyPayload): Promise<ApplySelectedMogrtTextResult> {
  // // Send edited subtitle text blocks to host so it can rebuild and retime the selected MOGRT clips.
  const encodedPayload = encodeURIComponent(JSON.stringify(payload));
  const response = await evalHostJson<ApplySelectedMogrtTextResult>(
    `subcreator_apply_selected_mogrt_text_items("${escapeForJsx(encodedPayload)}")`
  );
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to apply selected MOGRT text items.");
  }

  return {
    selectedCount: Number(response.data?.selectedCount || 0),
    rebuiltCount: Number(response.data?.rebuiltCount || 0),
    failedCount: Number(response.data?.failedCount || 0),
    selectionSignature: typeof response.data?.selectionSignature === "string" ? response.data.selectionSignature : undefined,
    sourceTrackIndex: Number.isFinite(Number(response.data?.sourceTrackIndex))
      ? Number(response.data?.sourceTrackIndex)
      : undefined,
    rebuildTrackIndex: Number.isFinite(Number(response.data?.rebuildTrackIndex))
      ? Number(response.data?.rebuildTrackIndex)
      : undefined,
    projectDocumentId: typeof response.data?.projectDocumentId === "string" ? response.data.projectDocumentId : undefined,
    projectPath: typeof response.data?.projectPath === "string" ? response.data.projectPath : undefined,
    sequenceID: typeof response.data?.sequenceID === "string" ? response.data.sequenceID : undefined,
    sequenceName: typeof response.data?.sequenceName === "string" ? response.data.sequenceName : undefined,
    debug: Array.isArray(response.data?.debug) ? response.data.debug.map((line) => String(line)) : undefined
  };
}

async function transcribeWithWhisperXViaCepNodeAsync(
  request: WhisperXTranscriptionRequest,
  onProgress?: (update: WhisperProgressUpdate) => void
): Promise<WhisperTranscriptionResult> {
  // // Run WhisperX as a transcription + forced-alignment pass and return Whisper-compatible JSON for the existing planner.
  const modules = resolveCepNodeModules();
  if (!modules) {
    throw new Error("CEP Node runtime unavailable. WhisperX transcription requires CEP Node and Python whisperx.");
  }

  if (typeof modules.childProcess.spawn !== "function") {
    throw new Error("CEP Node child_process.spawn unavailable. WhisperX transcription cannot start.");
  }

  const scriptPath = resolveBundledPythonScriptPath(modules, request.extensionRootPath, "subcreator_align_corrected.py");
  if (!scriptPath) {
    throw new Error("Bundled WhisperX helper is missing from the installed extension.");
  }

  const outputDir = modules.path.join(
    modules.os.tmpdir(),
    "SubCreatorWhisperX",
    `run-${Date.now()}-${Math.floor(Math.random() * 100000)}`
  );
  modules.fs.mkdirSync(outputDir, { recursive: true });

  const outputPath = modules.path.join(outputDir, "whisperx.json");
  const runtimeConfig = getRuntimeConfig(modules);
  const userExecutables = discoverUserWhisperExecutables(modules, runtimeConfig);
  const spawnEnv = buildSpawnEnv(modules, userExecutables, runtimeConfig);
  const pythonLaunchers = buildPythonLauncherCandidates(modules, userExecutables, runtimeConfig);
  const attempts: string[] = [];
  let collectedOutput = "";

  for (const launcher of pythonLaunchers) {
    let latestProgressPercent = -1;
    const attemptChunks: string[] = [];
    let attemptTailOutput = "";
    if (modules.fs.existsSync(outputPath)) {
      try {
        modules.fs.unlinkSync(outputPath);
      } catch {
        // // Ignore stale-output cleanup failures and let the attempt report the real execution problem.
      }
    }

    const args = [
      ...launcher.argsPrefix,
      scriptPath,
      "--audio",
      request.audioPath,
      "--language",
      request.languageCode,
      "--output",
      outputPath,
      "--transcribe-model",
      request.model
    ];

    const attemptResult = await new Promise<{ code: number | null; error?: { message?: string; code?: string } }>((resolve) => {
      // // Stream helper progress markers and keep cancellation routed through the existing CEP job registry.
      let settled = false;
      const detached = !detectWindowsRuntime();
      const child = modules.childProcess.spawn?.(launcher.command, args, {
        shell: false,
        detached,
        env: spawnEnv
      });
      if (!child) {
        resolve({
          code: null,
          error: {
            message: "child_process.spawn unavailable"
          }
        });
        return;
      }
      const activeJob = registerActiveCepJob(child, launcher.label, detached);

      const handleChunk = (chunk: string | Uint8Array): void => {
        // // Parse helper progress output without flooding the panel with duplicate percentages.
        const normalizedChunk = normalizeWhisperOutputChunk(chunk);
        if (!normalizedChunk) {
          return;
        }
        attemptChunks.push(normalizedChunk);
        attemptTailOutput = `${attemptTailOutput}${normalizedChunk}`.slice(-4096);
        const progress = extractCorrectedAlignProgressUpdate(attemptTailOutput);
        if (!progress || progress.percent <= latestProgressPercent) {
          return;
        }
        latestProgressPercent = progress.percent;
        onProgress?.(progress);
      };

      child.stdout?.on("data", handleChunk);
      child.stderr?.on("data", handleChunk);
      child.on("error", (error) => {
        clearActiveCepJob(child);
        if (settled) {
          return;
        }
        settled = true;
        if (activeJob.cancelRequested) {
          resolve({
            code: null,
            error: {
              message: SUBCREATOR_CANCELLED_JOB_CODE,
              code: SUBCREATOR_CANCELLED_JOB_CODE
            }
          });
          return;
        }
        resolve({
          code: null,
          error: error && typeof error === "object" ? (error as { message?: string; code?: string }) : { message: String(error) }
        });
      });
      child.on("close", (value) => {
        clearActiveCepJob(child);
        if (settled) {
          return;
        }
        settled = true;
        if (activeJob.cancelRequested) {
          resolve({
            code: null,
            error: {
              message: SUBCREATOR_CANCELLED_JOB_CODE,
              code: SUBCREATOR_CANCELLED_JOB_CODE
            }
          });
          return;
        }
        resolve({
          code: typeof value === "number" ? value : null
        });
      });
    });

    if (attemptResult.error) {
      if (isCancelledJobError(attemptResult.error.message || attemptResult.error.code || "")) {
        throw createCancelledJobError();
      }
      attempts.push(`${launcher.label}: ${String(attemptResult.error.message || attemptResult.error)}`);
      continue;
    }

    const attemptOutput = attemptChunks.join("").trim();
    if (attemptOutput && !collectedOutput) {
      collectedOutput = attemptOutput;
    }
    const authoritativeAttempt =
      isRuntimeConfigPythonLauncher(launcher, runtimeConfig) && !shouldContinueAfterPythonHelperFailure(attemptOutput);

    if (attemptResult.code !== 0) {
      const summary = summarizeWhisperErrorOutput(attemptOutput);
      attempts.push(`${launcher.label}: exit ${attemptResult.code}${summary ? ` (${summary})` : ""}`);
      if (authoritativeAttempt) {
        break;
      }
      continue;
    }

    if (!modules.fs.existsSync(outputPath)) {
      attempts.push(`${launcher.label}: no WhisperX json output`);
      continue;
    }

    const jsonText = String(modules.fs.readFileSync(outputPath, "utf8") || "").trim();
    if (!jsonText) {
      attempts.push(`${launcher.label}: empty WhisperX json output`);
      continue;
    }

    onProgress?.({
      percent: 100,
      detail: "WhisperX transcription complete"
    });
    return {
      srtText: "",
      jsonText,
      model: request.model?.trim() || "base",
      audioPath: request.audioPath,
      commandOutput: attemptOutput
    };
  }

  let installHint = "";
  if (runtimeConfig?.pythonPath) {
    installHint = `Install command: ${runtimeConfig.pythonPath} -m pip install --user --upgrade whisperx certifi`;
  } else if (runtimeConfig?.pythonCommand) {
    installHint = `Install command: ${runtimeConfig.pythonCommand} -m pip install --user --upgrade whisperx certifi`;
  } else if (detectWindowsRuntime()) {
    installHint = "Install command: py -3.11 -m pip install --user --upgrade whisperx certifi";
  } else {
    installHint = "Install command: python3.11 -m pip install --user --upgrade whisperx certifi";
  }

  throw new Error(
    `Unable to execute WhisperX transcription. Attempts: ${attempts.join(" | ") || "none"}. ${installHint}. ${
      runtimeConfig ? `Runtime config: ${runtimeConfig.sourcePath}. ` : ""
    }${collectedOutput || ""}`
  );
}

async function alignCorrectedTranscriptViaCepNodeAsync(
  request: CorrectedAlignmentRequest,
  onProgress?: (update: WhisperProgressUpdate) => void
): Promise<CorrectedAlignmentResult> {
  // // Run the bundled WhisperX helper script with CEP Node so corrected transcripts can be aligned to active-sequence audio.
  const modules = resolveCepNodeModules();
  if (!modules) {
    throw new Error("CEP Node runtime unavailable. Corrected transcript align requires CEP Node and Python whisperx.");
  }

  if (typeof modules.childProcess.spawn !== "function") {
    throw new Error("CEP Node child_process.spawn unavailable. Corrected transcript align cannot start.");
  }

  const scriptPath = resolveBundledPythonScriptPath(modules, request.extensionRootPath, "subcreator_align_corrected.py");
  if (!scriptPath) {
    throw new Error("Bundled corrected transcript align helper is missing from the installed extension.");
  }

  const outputDir = modules.path.join(
    modules.os.tmpdir(),
    "SubCreatorCorrectedAlign",
    `run-${Date.now()}-${Math.floor(Math.random() * 100000)}`
  );
  modules.fs.mkdirSync(outputDir, { recursive: true });

  const outputPath = modules.path.join(outputDir, "aligned.json");
  const runtimeConfig = getRuntimeConfig(modules);
  const userExecutables = discoverUserWhisperExecutables(modules, runtimeConfig);
  const spawnEnv = buildSpawnEnv(modules, userExecutables, runtimeConfig);
  const pythonLaunchers = buildPythonLauncherCandidates(modules, userExecutables, runtimeConfig);
  const attempts: string[] = [];
  let collectedOutput = "";

  for (const launcher of pythonLaunchers) {
    let latestProgressPercent = -1;
    const attemptChunks: string[] = [];
    let attemptTailOutput = "";
    if (modules.fs.existsSync(outputPath)) {
      try {
        modules.fs.unlinkSync(outputPath);
      } catch {
        // // Ignore stale-output cleanup failures and let the attempt report the real execution problem.
      }
    }
    const args = [
      ...launcher.argsPrefix,
      scriptPath,
      "--audio",
      request.audioPath,
      "--transcript",
      request.transcriptPath,
      "--language",
      request.languageCode,
      "--output",
      outputPath
    ];
    if (Number.isFinite(Number(request.rangeStartSeconds))) {
      args.push("--range-start-seconds", String(Number(request.rangeStartSeconds)));
    }
    if (Number.isFinite(Number(request.rangeEndSeconds))) {
      args.push("--range-end-seconds", String(Number(request.rangeEndSeconds)));
    }

    const attemptResult = await new Promise<{ code: number | null; error?: { message?: string; code?: string } }>((resolve) => {
      // // Attach stderr listeners before process start so helper progress markers can feed the UI immediately.
      let settled = false;
      const detached = !detectWindowsRuntime();
      const child = modules.childProcess.spawn?.(launcher.command, args, {
        shell: false,
        detached,
        env: spawnEnv
      });
      if (!child) {
        resolve({
          code: null,
          error: {
            message: "child_process.spawn unavailable"
          }
        });
        return;
      }
      const activeJob = registerActiveCepJob(child, launcher.label, detached);

      const handleChunk = (chunk: string | Uint8Array): void => {
        // // Parse helper progress output without flooding the panel with duplicate percentages.
        const normalizedChunk = normalizeWhisperOutputChunk(chunk);
        if (!normalizedChunk) {
          return;
        }
        attemptChunks.push(normalizedChunk);
        attemptTailOutput = `${attemptTailOutput}${normalizedChunk}`.slice(-4096);
        const progress = extractCorrectedAlignProgressUpdate(attemptTailOutput);
        if (!progress || progress.percent <= latestProgressPercent) {
          return;
        }
        latestProgressPercent = progress.percent;
        onProgress?.(progress);
      };

      child.stdout?.on("data", handleChunk);
      child.stderr?.on("data", handleChunk);
      child.on("error", (error) => {
        clearActiveCepJob(child);
        if (settled) {
          return;
        }
        settled = true;
        if (activeJob.cancelRequested) {
          resolve({
            code: null,
            error: {
              message: SUBCREATOR_CANCELLED_JOB_CODE,
              code: SUBCREATOR_CANCELLED_JOB_CODE
            }
          });
          return;
        }
        resolve({
          code: null,
          error: error && typeof error === "object" ? (error as { message?: string; code?: string }) : { message: String(error) }
        });
      });
      child.on("close", (value) => {
        clearActiveCepJob(child);
        if (settled) {
          return;
        }
        settled = true;
        if (activeJob.cancelRequested) {
          resolve({
            code: null,
            error: {
              message: SUBCREATOR_CANCELLED_JOB_CODE,
              code: SUBCREATOR_CANCELLED_JOB_CODE
            }
          });
          return;
        }
        resolve({
          code: typeof value === "number" ? value : null
        });
      });
    });

    if (attemptResult.error) {
      if (isCancelledJobError(attemptResult.error.message || attemptResult.error.code || "")) {
        throw createCancelledJobError();
      }
      attempts.push(`${launcher.label}: ${String(attemptResult.error.message || attemptResult.error)}`);
      continue;
    }

    const attemptOutput = attemptChunks.join("").trim();
    if (attemptOutput && !collectedOutput) {
      collectedOutput = attemptOutput;
    }
    const authoritativeAttempt =
      isRuntimeConfigPythonLauncher(launcher, runtimeConfig) && !shouldContinueAfterPythonHelperFailure(attemptOutput);

    if (attemptResult.code !== 0) {
      const summary = summarizeWhisperErrorOutput(attemptOutput);
      attempts.push(`${launcher.label}: exit ${attemptResult.code}${summary ? ` (${summary})` : ""}`);
      if (authoritativeAttempt) {
        break;
      }
      continue;
    }

    if (!modules.fs.existsSync(outputPath)) {
      attempts.push(`${launcher.label}: no aligned json output`);
      continue;
    }

    const jsonText = String(modules.fs.readFileSync(outputPath, "utf8") || "").trim();
    if (!jsonText) {
      attempts.push(`${launcher.label}: empty aligned json output`);
      continue;
    }

    onProgress?.({
      percent: 100,
      detail: "Corrected transcript alignment complete"
    });
    return {
      jsonText,
      audioPath: request.audioPath,
      transcriptPath: request.transcriptPath,
      commandOutput: attemptOutput
    };
  }

  let installHint = "";
  if (runtimeConfig?.pythonPath) {
    installHint = `Install command: ${runtimeConfig.pythonPath} -m pip install --user --upgrade whisperx certifi`;
  } else if (runtimeConfig?.pythonCommand) {
    installHint = `Install command: ${runtimeConfig.pythonCommand} -m pip install --user --upgrade whisperx certifi`;
  } else if (detectWindowsRuntime()) {
    installHint = "Install command: py -3.11 -m pip install --user --upgrade whisperx certifi";
  } else {
    installHint = "Install command: python3.11 -m pip install --user --upgrade whisperx certifi";
  }

  throw new Error(
    `Unable to execute corrected transcript align via WhisperX. Attempts: ${attempts.join(" | ") || "none"}. ${installHint}. ${
      runtimeConfig ? `Runtime config: ${runtimeConfig.sourcePath}. ` : ""
    }${collectedOutput || ""}`
  );
}

export async function transcribeWithWhisper(
  request: WhisperTranscriptionRequest,
  onProgress?: (update: WhisperProgressUpdate) => void
): Promise<WhisperTranscriptionResult> {
  // // Prefer CEP Node runtime for Whisper CLI, fallback to host ExtendScript bridge.
  const nodeResult = await transcribeWithWhisperViaCepNodeAsync(request, onProgress);
  if (nodeResult) {
    return nodeResult;
  }

  const encodedPayload = encodeURIComponent(JSON.stringify(request));
  const response = await evalHostJson<WhisperTranscriptionResult>(
    `subcreator_transcribe_whisper("${escapeForJsx(encodedPayload)}")`
  );

  if (!response.ok) {
    throw new Error(response.error ?? "Whisper transcription failed.");
  }

  return {
    srtText: String(response.data?.srtText ?? ""),
    jsonText: typeof response.data?.jsonText === "string" ? response.data.jsonText : undefined,
    model: String(response.data?.model ?? request.model),
    audioPath: String(response.data?.audioPath ?? request.audioPath),
    commandOutput: String(response.data?.commandOutput ?? "")
  };
}

export async function alignCorrectedTranscript(
  request: CorrectedAlignmentRequest,
  onProgress?: (update: WhisperProgressUpdate) => void
): Promise<CorrectedAlignmentResult> {
  // // Corrected transcript alignment depends on the bundled WhisperX helper and has no ExtendScript fallback path.
  return alignCorrectedTranscriptViaCepNodeAsync(request, onProgress);
}

export async function transcribeWithWhisperX(
  request: WhisperXTranscriptionRequest,
  onProgress?: (update: WhisperProgressUpdate) => void
): Promise<WhisperTranscriptionResult> {
  // // WhisperX transcription depends on the bundled helper and reuses the existing Whisper JSON parser.
  return transcribeWithWhisperXViaCepNodeAsync(request, onProgress);
}
