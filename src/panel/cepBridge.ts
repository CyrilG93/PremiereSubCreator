// // Wrap CEP evalScript calls and provide a browser fallback for local testing.
import type { HostApplyPayload, HostCaptionCue } from "../core/types";

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

interface CepNodeModules {
  childProcess: {
    spawnSync: (
      command: string,
      args: string[],
      options: { encoding: string; shell?: boolean; timeout?: number }
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
      path: nodeRequire("path") as CepNodeModules["path"]
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

function discoverUserWhisperExecutables(modules: CepNodeModules): string[] {
  // // Probe common user-local installation locations where PATH may be incomplete in CEP.
  const discovered: string[] = [];
  const home = modules.os.homedir();
  const directCandidates = [modules.path.join(home, ".local", "bin", "whisper"), modules.path.join(home, "bin", "whisper")];

  for (const candidate of directCandidates) {
    if (modules.fs.existsSync(candidate)) {
      discovered.push(candidate);
    }
  }

  const pythonRoot = modules.path.join(home, "Library", "Python");
  if (modules.fs.existsSync(pythonRoot)) {
    try {
      const versions = modules.fs.readdirSync(pythonRoot);
      for (const versionName of versions) {
        const candidate = modules.path.join(pythonRoot, versionName, "bin", "whisper");
        if (modules.fs.existsSync(candidate)) {
          discovered.push(candidate);
        }
      }
    } catch {
      // // Ignore inaccessible folders and continue other command candidates.
    }
  }

  return discovered;
}

function buildWhisperCommandCandidates(
  modules: CepNodeModules,
  request: WhisperTranscriptionRequest,
  outputDir: string
): WhisperCommandCandidate[] {
  // // Build ordered command fallbacks for diverse Whisper install methods.
  const baseArgs = buildWhisperArgs(request, outputDir);
  const candidates: WhisperCommandCandidate[] = [];
  const isWindows = detectWindowsRuntime();
  const userExecutables = discoverUserWhisperExecutables(modules);

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

  for (const executablePath of userExecutables) {
    pushCandidate(executablePath, baseArgs, executablePath);
  }

  pushCandidate("whisper", baseArgs, "whisper");
  pushCandidate("python3", ["-m", "whisper", ...baseArgs], "python3 -m whisper");
  pushCandidate("python", ["-m", "whisper", ...baseArgs], "python -m whisper");

  if (isWindows) {
    pushCandidate("py", ["-3", "-m", "whisper", ...baseArgs], "py -3 -m whisper");
    pushCandidate("py", ["-m", "whisper", ...baseArgs], "py -m whisper");
  }

  return candidates;
}

function buildPythonLauncherCandidates(): PythonLauncherCandidate[] {
  // // Build Python launcher candidates used to detect `openai-whisper` module availability.
  const candidates: PythonLauncherCandidate[] = [
    { command: "python3", argsPrefix: [], label: "python3" },
    { command: "python", argsPrefix: [], label: "python" }
  ];

  if (detectWindowsRuntime()) {
    candidates.unshift({ command: "py", argsPrefix: ["-3"], label: "py -3" });
    candidates.push({ command: "py", argsPrefix: [], label: "py" });
  }

  return candidates;
}

function runSpawn(
  modules: CepNodeModules,
  command: string,
  args: string[]
): { ok: boolean; status: number; stdout: string; stderr: string; errorCode: string; errorMessage: string } {
  // // Execute a command synchronously with a safe timeout and normalized result shape.
  const run = modules.childProcess.spawnSync(command, args, {
    encoding: "utf8",
    shell: false,
    timeout: 12000
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

  const pythonLaunchers = buildPythonLauncherCandidates();
  for (const launcher of pythonLaunchers) {
    const probe = runSpawn(modules, launcher.command, [...launcher.argsPrefix, "-c", "import whisper"]);
    if (probe.ok) {
      return {
        available: true,
        details: `Python module detected via ${launcher.label}`
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

  const userExecutables = discoverUserWhisperExecutables(modules);
  const cliCandidates = [...userExecutables, "whisper"];
  for (const command of cliCandidates) {
    const probe = runSpawn(modules, command, ["--help"]);
    if (probe.ok) {
      return {
        available: true,
        details: `CLI detected via ${command}`
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
    details: checks.join(" | ")
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

  const attempts: string[] = [];
  const commandCandidates = buildWhisperCommandCandidates(modules, request, outputDir);
  let collectedOutput = "";
  let transcribed = false;

  for (const candidate of commandCandidates) {
    const run = modules.childProcess.spawnSync(candidate.command, candidate.args, {
      encoding: "utf8",
      shell: false
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
    if (attemptOutput) {
      collectedOutput = attemptOutput;
    }

    if (typeof run.status === "number" && run.status !== 0) {
      attempts.push(`${candidate.label}: exit ${run.status}`);
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

    transcribed = true;
    return {
      srtText,
      model: request.model?.trim() || "base",
      audioPath: request.audioPath,
      commandOutput: attemptOutput
    };
  }

  if (!transcribed) {
    const installHint = detectWindowsRuntime()
      ? "Install command: py -m pip install -U openai-whisper"
      : "Install command: python3 -m pip install -U openai-whisper";
    throw new Error(
      `Unable to execute Whisper CLI from CEP runtime. Attempts: ${attempts.join(" | ") || "none"}. ${installHint}. ${
        collectedOutput || ""
      }`
    );
  }

  return null;
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

export async function readActiveCaptionTrackCues(): Promise<HostCaptionCue[]> {
  // // Extract caption text/timing directly from the active caption track in Premiere.
  const response = await evalHostJson<HostCaptionCue[]>("subcreator_extract_active_caption_track()");
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to read active caption track.");
  }

  return Array.isArray(response.data) ? response.data : [];
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
