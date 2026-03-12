// // Wrap CEP evalScript calls and provide a browser fallback for local testing.
import type { HostApplyPayload } from "../core/types";

declare global {
  interface Window {
    __adobe_cep__?: {
      evalScript: (script: string, callback: (result: string) => void) => void;
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

interface WhisperTranscriptionResult {
  srtText: string;
  model: string;
  audioPath: string;
  commandOutput?: string;
}

export interface WhisperRuntimeStatus {
  available: boolean;
  details: string;
}

export interface SelectedMogrtVisualProperty {
  path: string;
  displayName: string;
  groupPath: string;
  valueType: "number" | "boolean" | "string" | "json";
  controlKind: "slider" | "number" | "checkbox" | "color" | "text" | "string" | "json" | "vector";
  value: string | number | boolean;
  minValue?: number;
  maxValue?: number;
  stepValue?: number;
}

export interface SelectedMogrtVisualPropertyList {
  selectedCount: number;
  editableCount: number;
  properties: SelectedMogrtVisualProperty[];
}

export interface ApplyVisualPropertiesResult {
  selectedCount: number;
  updatedCount: number;
  failedCount: number;
}

interface CepNodeModules {
  childProcess: {
    spawnSync: (
      command: string,
      args: string[],
      options: { encoding: string; shell?: boolean; timeout?: number; env?: Record<string, string | undefined> }
    ) => { status: number | null; stdout?: string; stderr?: string; error?: { message?: string; code?: string } };
  };
  fs: {
    existsSync: (path: string) => boolean;
    mkdirSync: (path: string, options: { recursive: boolean }) => void;
    readdirSync: (path: string) => string[];
    readFileSync: (path: string, encoding: string) => string;
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
  };
}

interface WhisperCommandCandidate {
  command: string;
  args: string[];
  label: string;
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

let subcreatorRuntimeConfigCache: SubcreatorRuntimeConfig | null | undefined;

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

async function evalHostJson<T>(script: string): Promise<HostJsonResponse<T>> {
  // // Parse JSON returned by host-side ExtendScript function calls.
  const raw = await evalScript(script);

  try {
    return JSON.parse(raw) as HostJsonResponse<T>;
  } catch (error) {
    return {
      ok: false,
      error: `Invalid host response: ${String(error)} | raw=${raw}`
    };
  }
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

function buildWhisperArgs(request: WhisperTranscriptionRequest, outputDir: string): string[] {
  // // Build CLI arguments for local Whisper invocation.
  const model = request.model?.trim() || "base";
  const args = [request.audioPath, "--model", model, "--output_format", "srt", "--output_dir", outputDir, "--fp16", "False"];

  const language = request.languageCode?.trim();
  if (language && language.toLowerCase() !== "auto") {
    args.push("--language", language);
  }

  return args;
}

function detectWindowsRuntime(): boolean {
  // // Detect Windows from CEP browser runtime for CLI fallback ordering.
  return /win/i.test(String(navigator?.platform || ""));
}

function pushUniqueString(target: string[], value: string): void {
  // // Append a value only once while preserving initial order.
  const normalized = String(value || "").trim();
  if (!normalized) {
    return;
  }

  const lookup = normalized.toLowerCase();
  for (const item of target) {
    if (String(item || "").trim().toLowerCase() === lookup) {
      return;
    }
  }

  target.push(normalized);
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

function readRuntimeConfigFromDisk(modules: CepNodeModules): SubcreatorRuntimeConfig | null {
  // // Read installer-generated runtime config so CEP can use exact binary paths reliably.
  const candidates = resolveRuntimeConfigPathCandidates(modules);
  for (const candidatePath of candidates) {
    if (!modules.fs.existsSync(candidatePath)) {
      continue;
    }

    try {
      const rawText = String(modules.fs.readFileSync(candidatePath, "utf8") || "");
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

  return null;
}

function getRuntimeConfig(modules: CepNodeModules): SubcreatorRuntimeConfig | null {
  // // Cache runtime config lookup per session to avoid repeated disk reads.
  if (typeof subcreatorRuntimeConfigCache !== "undefined") {
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

  const versionedPython = ["python3.13", "python3.12", "python3.11", "python3.10", "python3.9", "python3.8"];
  for (const pythonCommand of versionedPython) {
    pushCandidate(pythonCommand, ["-m", "whisper", ...baseArgs], `${pythonCommand} -m whisper`);
  }

  pushCandidate("whisper", baseArgs, "whisper");
  pushCandidate("python3", ["-m", "whisper", ...baseArgs], "python3 -m whisper");
  pushCandidate("python", ["-m", "whisper", ...baseArgs], "python -m whisper");

  if (isWindows) {
    pushCandidate("py", ["-3.13", "-m", "whisper", ...baseArgs], "py -3.13 -m whisper");
    pushCandidate("py", ["-3.12", "-m", "whisper", ...baseArgs], "py -3.12 -m whisper");
    pushCandidate("py", ["-3.11", "-m", "whisper", ...baseArgs], "py -3.11 -m whisper");
    pushCandidate("py", ["-3.10", "-m", "whisper", ...baseArgs], "py -3.10 -m whisper");
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

  const posixVersioned = ["python3.13", "python3.12", "python3.11", "python3.10", "python3.9", "python3.8"];
  for (const command of posixVersioned) {
    pushCandidate(command, [], command);
  }

  pushCandidate("python3", [], "python3");
  pushCandidate("python", [], "python");

  if (detectWindowsRuntime()) {
    pushCandidate("py", ["-3.13"], "py -3.13");
    pushCandidate("py", ["-3.12"], "py -3.12");
    pushCandidate("py", ["-3.11"], "py -3.11");
    pushCandidate("py", ["-3.10"], "py -3.10");
    pushCandidate("py", ["-3.9"], "py -3.9");
    pushCandidate("py", ["-3.8"], "py -3.8");
    pushCandidate("py", ["-3"], "py -3");
    pushCandidate("py", [], "py");
  }

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

function detectWhisperAvailabilityViaCepNode(): WhisperRuntimeStatus {
  // // Detect local Whisper runtime availability to decide if Whisper source should be shown.
  const modules = resolveCepNodeModules();
  if (!modules) {
    return {
      available: false,
      details: "CEP Node runtime unavailable"
    };
  }

  const checks: string[] = [];
  const runtimeConfig = getRuntimeConfig(modules);
  const userExecutables = discoverUserWhisperExecutables(modules, runtimeConfig);
  const spawnEnv = buildSpawnEnv(modules, userExecutables, runtimeConfig);

  const pythonLaunchers = buildPythonLauncherCandidates(modules, userExecutables, runtimeConfig);
  for (const launcher of pythonLaunchers) {
    const probe = runSpawn(modules, launcher.command, [...launcher.argsPrefix, "-c", "import whisper"], spawnEnv);
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

  const cliCandidates = [...userExecutables, "whisper"];
  for (const command of cliCandidates) {
    const probe = runSpawn(modules, command, ["--help"], spawnEnv);
    if (probe.ok) {
      return {
        available: true,
        details: `CLI detected via ${command}${runtimeConfig ? ` (config: ${runtimeConfig.sourcePath})` : ""}`
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
    details: `${checks.join(" | ")}${runtimeConfig ? ` | config=${runtimeConfig.sourcePath}` : ""}`
  };
}

function resolveWhisperSrtPath(modules: CepNodeModules, outputDir: string, audioPath: string): string {
  // // Resolve Whisper output SRT from expected filename or directory scan fallback.
  const baseName = modules.path.basename(audioPath).replace(/\.[^/.]+$/, "");
  const direct = modules.path.join(outputDir, `${baseName}.srt`);
  if (modules.fs.existsSync(direct)) {
    return direct;
  }

  const entries = modules.fs.readdirSync(outputDir);
  for (const entry of entries) {
    const lower = String(entry).toLowerCase();
    if (lower.endsWith(".srt") && lower.startsWith(baseName.toLowerCase())) {
      return modules.path.join(outputDir, entry);
    }
  }

  return "";
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

function transcribeWithWhisperViaCepNode(request: WhisperTranscriptionRequest): WhisperTranscriptionResult | null {
  // // Run Whisper with CEP Node runtime to avoid ExtendScript `system.callSystem` availability issues.
  const modules = resolveCepNodeModules();
  if (!modules) {
    return null;
  }

  const outputDir = modules.path.join(modules.os.tmpdir(), "SubCreatorWhisper");
  if (!modules.fs.existsSync(outputDir)) {
    modules.fs.mkdirSync(outputDir, { recursive: true });
  }

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

    return {
      srtText,
      model: request.model?.trim() || "base",
      audioPath: request.audioPath,
      commandOutput: attemptOutput
    };
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
      "Model download failed due TLS/SSL certificate validation. Configure trusted certs/proxy for Python, or pre-download Whisper models.";
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
  return evalScript(`subcreator_apply_captions("${escapeForJsx(encodedPayload)}")`);
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

export async function pickWhisperAudioPath(): Promise<string> {
  // // Open native file picker from host for Whisper transcription input.
  const response = await evalHostJson<{ path: string }>("subcreator_pick_audio_file()");
  if (!response.ok) {
    throw new Error(response.error ?? "Audio picker failed.");
  }

  return String(response.data?.path ?? "");
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
      minValue: Number.isFinite(Number(item.minValue)) ? Number(item.minValue) : undefined,
      maxValue: Number.isFinite(Number(item.maxValue)) ? Number(item.maxValue) : undefined,
      stepValue: Number.isFinite(Number(item.stepValue)) ? Number(item.stepValue) : undefined,
      value
    });
  }

  return {
    selectedCount: Number(payload.selectedCount || 0),
    editableCount: Number(payload.editableCount || properties.length),
    properties
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

export async function applyVisualPropertiesToSelectedMogrts(
  changes: Array<{
    path: string;
    valueType: SelectedMogrtVisualProperty["valueType"];
    controlKind: SelectedMogrtVisualProperty["controlKind"];
    value: string | number | boolean;
  }>
): Promise<ApplyVisualPropertiesResult> {
  // // Send edited property payload to host and apply values on selected MOGRT clips.
  const encodedPayload = encodeURIComponent(JSON.stringify({ changes }));
  const response = await evalHostJson<ApplyVisualPropertiesResult>(
    `subcreator_apply_selected_mogrt_properties("${escapeForJsx(encodedPayload)}")`
  );
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to apply selected MOGRT properties.");
  }

  return {
    selectedCount: Number(response.data?.selectedCount || 0),
    updatedCount: Number(response.data?.updatedCount || 0),
    failedCount: Number(response.data?.failedCount || 0)
  };
}

export async function transcribeWithWhisper(request: WhisperTranscriptionRequest): Promise<WhisperTranscriptionResult> {
  // // Prefer CEP Node runtime for Whisper CLI, fallback to host ExtendScript bridge.
  const nodeResult = transcribeWithWhisperViaCepNode(request);
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
    model: String(response.data?.model ?? request.model),
    audioPath: String(response.data?.audioPath ?? request.audioPath),
    commandOutput: String(response.data?.commandOutput ?? "")
  };
}
