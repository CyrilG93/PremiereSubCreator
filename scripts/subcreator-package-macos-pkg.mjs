// // Build the complete Apple Silicon macOS Installer with its private runtime embedded.
import { createReadStream } from "node:fs";
import { cp, mkdir, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const stagingRoot = path.join(projectRoot, ".subcreator-macos-staging");
const downloadsDir = path.join(stagingRoot, "downloads");
const runtimeRoot = path.join(stagingRoot, "runtime");
const packagesDir = path.join(stagingRoot, "packages");
const coreScriptsDir = path.join(stagingRoot, "core-scripts");
const releasesDir = path.join(projectRoot, "Releases");
const runtimeManifestPath = path.join(projectRoot, "installers", "macos-runtime.json");
const pythonVersion = process.env.SUBCREATOR_PRIVATE_PYTHON_VERSION || "3.11.9";
const ffmpegVersion = process.env.SUBCREATOR_FFMPEG_VERSION || "8.0.2";
const ffmpegSourceUrl =
  process.env.SUBCREATOR_FFMPEG_SOURCE_URL || `https://ffmpeg.org/releases/ffmpeg-${ffmpegVersion}.tar.xz`;
const macArch = "arm64";
const rebuildRuntime = process.env.SUBCREATOR_REBUILD_RUNTIME === "1";
const reuseStaging = process.env.SUBCREATOR_REUSE_STAGING === "1";
const npmCommand = "npm";
const whisperModels = [
  {
    id: "tiny",
    title: "Whisper Tiny",
    description: "Fastest model, lower accuracy. About 75 MB.",
    fileName: "tiny.pt",
    sha256: "65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9",
    url: "https://openaipublic.azureedge.net/main/whisper/models/65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9/tiny.pt",
    defaultSelected: false
  },
  {
    id: "base",
    title: "Whisper Base",
    description: "Recommended starter model. About 142 MB.",
    fileName: "base.pt",
    sha256: "ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e",
    url: "https://openaipublic.azureedge.net/main/whisper/models/ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e/base.pt",
    defaultSelected: true
  },
  {
    id: "small",
    title: "Whisper Small",
    description: "Good quality and speed balance. About 466 MB.",
    fileName: "small.pt",
    sha256: "9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794",
    url: "https://openaipublic.azureedge.net/main/whisper/models/9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794/small.pt",
    defaultSelected: false
  },
  {
    id: "medium",
    title: "Whisper Medium",
    description: "Better for difficult audio. About 1.5 GB.",
    fileName: "medium.pt",
    sha256: "345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1",
    url: "https://openaipublic.azureedge.net/main/whisper/models/345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1/medium.pt",
    defaultSelected: false
  },
  {
    id: "large-v3",
    title: "Whisper Large v3",
    description: "Best accuracy, slowest model. About 2.9 GB.",
    fileName: "large-v3.pt",
    sha256: "e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb",
    url: "https://openaipublic.azureedge.net/main/whisper/models/e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb/large-v3.pt",
    defaultSelected: false
  }
];

function runCommand(command, args, options = {}) {
  // // Execute native packaging and runtime-build commands with visible progress and strict exit handling.
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options.cwd || projectRoot,
      env: {
        ...process.env,
        ...(options.env || {})
      },
      stdio: options.stdio || "inherit"
    });

    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`${command} exited with code ${code}`));
    });
  });
}

async function pathExists(targetPath) {
  // // Probe optional staging and release assets without turning a normal cache miss into an exception.
  try {
    await stat(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function prunePackagingMetadata(targetDir) {
  // // Remove Finder and AppleDouble metadata so installer payloads contain only intentional project files.
  const entries = await readdir(targetDir, { withFileTypes: true });
  for (const entry of entries) {
    const entryPath = path.join(targetDir, entry.name);
    if (entry.name === ".DS_Store" || entry.name.startsWith("._")) {
      await rm(entryPath, { recursive: entry.isDirectory(), force: true });
      continue;
    }
    if (entry.isDirectory()) {
      await prunePackagingMetadata(entryPath);
    }
  }
}

async function hashFile(targetPath) {
  // // Stream large runtime archives through SHA-256 without loading them into Node memory.
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    const stream = createReadStream(targetPath);
    stream.on("error", reject);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
  });
}

async function downloadFile(url, targetPath) {
  // // Download immutable build or release assets with curl and keep successful files for later packaging runs.
  if (await pathExists(targetPath)) {
    return;
  }

  await mkdir(path.dirname(targetPath), { recursive: true });
  await runCommand("curl", [
    "--fail",
    "--location",
    "--retry",
    "3",
    "--connect-timeout",
    "30",
    url,
    "--output",
    targetPath
  ]);
}

async function readPackageVersion() {
  // // Use package.json as the single installer version source.
  const raw = await readFile(path.join(projectRoot, "package.json"), "utf8");
  return String(JSON.parse(raw).version || "").trim();
}

async function readRuntimeManifest() {
  // // Read the Apple Silicon runtime metadata used to assemble the Full installer.
  const raw = await readFile(runtimeManifestPath, "utf8");
  const parsed = JSON.parse(raw);
  return {
    version: String(parsed.version || "").trim(),
    releaseTag: String(parsed.releaseTag || "").trim(),
    assets: parsed.assets && typeof parsed.assets === "object" ? parsed.assets : {}
  };
}

async function writeRuntimeManifest(manifest) {
  // // Persist the generated runtime hash so later Full builds can reuse the exact archive.
  await writeFile(runtimeManifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
}

async function findCommand(command) {
  // // Resolve required external build tools through the developer shell PATH.
  return new Promise((resolve) => {
    const child = spawn("/usr/bin/which", [command], { stdio: ["ignore", "pipe", "ignore"] });
    let output = "";
    child.stdout.on("data", (chunk) => {
      output += chunk.toString();
    });
    child.on("exit", (code) => resolve(code === 0 ? output.trim() : ""));
  });
}

async function preparePrivatePython() {
  // // Install a self-contained python-build-standalone distribution and add Whisper packages directly into it.
  const uvPath = process.env.SUBCREATOR_UV_PATH || (await findCommand("uv"));
  if (!uvPath) {
    throw new Error("uv is required to build the private macOS runtime. Install it from https://docs.astral.sh/uv/.");
  }

  const pythonInstallDir = path.join(stagingRoot, "python-install");
  await rm(pythonInstallDir, { recursive: true, force: true });
  await mkdir(pythonInstallDir, { recursive: true });
  await runCommand(
    uvPath,
    [
      "--native-tls",
      "python",
      "install",
      pythonVersion,
      "--install-dir",
      pythonInstallDir,
      "--no-bin",
      "--no-progress"
    ],
    {
      env: {
        UV_PYTHON_PREFERENCE: "only-managed"
      }
    }
  );

  const entries = await readdir(pythonInstallDir, { withFileTypes: true });
  const pythonEntry = entries.find(
    (entry) => entry.isDirectory() && entry.name.startsWith(`cpython-${pythonVersion}-macos-`)
  );
  if (!pythonEntry) {
    throw new Error(`uv did not create the expected Python ${pythonVersion} installation.`);
  }

  const sourcePythonDir = path.join(pythonInstallDir, pythonEntry.name);
  const targetPythonDir = path.join(runtimeRoot, "python");
  await rm(targetPythonDir, { recursive: true, force: true });
  // // Preserve the distribution's relative executable symlinks instead of converting them to staging paths.
  await runCommand("ditto", [sourcePythonDir, targetPythonDir]);
  const pythonPath = path.join(targetPythonDir, "bin", "python3");
  const runtimePackages = [
    "torch==2.8.0",
    "torchaudio==2.8.0",
    "torchvision==0.23.0",
    "openai-whisper",
    "whisperx",
    "requests",
    "nltk",
    "certifi"
  ];

  await runCommand(
    uvPath,
    [
      "--native-tls",
      "pip",
      "install",
      "--python",
      pythonPath,
      "--break-system-packages",
      "--upgrade",
      "--no-cache",
      ...runtimePackages
    ],
    {
      env: {
        PYTHONNOUSERSITE: "1",
        PYTHONPATH: ""
      }
    }
  );

  await prunePrivatePython(targetPythonDir);
  await runCommand(pythonPath, [
    "-c",
    "import sys, whisper, whisperx; print('Private macOS Python runtime OK'); print(sys.executable)"
  ], {
    env: {
      PYTHONNOUSERSITE: "1",
      PYTHONPATH: ""
    }
  });
  await writeFile(path.join(runtimeRoot, ".subcreator-python-validated"), `${macArch}:${pythonVersion}\n`, "ascii");
}

async function prunePrivatePython(pythonDir) {
  // // Remove development-only artifacts that add substantial size but are not needed for transcription.
  const pruneScript = [
    `root=${shellQuote(pythonDir)}`,
    'find "$root" -type d -name "__pycache__" -prune -exec rm -rf {} +',
    'find "$root" -type f \\( -name "*.a" -o -name "*.pyc" -o -name "*.pyo" \\) -delete',
    'rm -rf "$root/lib/python3.11/site-packages/torch/include"',
    'rm -rf "$root/lib/python3.11/site-packages/torch/share/cmake"',
    'rm -rf "$root/lib/python3.11/test" "$root/lib/python3.11/idlelib" "$root/lib/python3.11/tkinter"'
  ].join("\n");
  await runCommand("/bin/bash", ["-c", pruneScript]);
}

async function preparePrivateFfmpeg() {
  // // Compile an LGPL-only FFmpeg from official source so the runtime has no Homebrew or GPL dependency.
  const sourceArchive = path.join(downloadsDir, `ffmpeg-${ffmpegVersion}.tar.xz`);
  const sourceParent = path.join(stagingRoot, "ffmpeg-source");
  const sourceDir = path.join(sourceParent, `ffmpeg-${ffmpegVersion}`);
  const installDir = path.join(runtimeRoot, "ffmpeg");
  await downloadFile(ffmpegSourceUrl, sourceArchive);
  await rm(sourceParent, { recursive: true, force: true });
  await mkdir(sourceParent, { recursive: true });
  await runCommand("tar", ["-xJf", sourceArchive, "-C", sourceParent]);
  await rm(installDir, { recursive: true, force: true });

  const cpuCount = Math.max(1, Number.parseInt(process.env.SUBCREATOR_BUILD_JOBS || "", 10) || 4);
  const configureArgs = [
      `--prefix=${installDir}`,
      "--disable-debug",
      "--disable-doc",
      "--disable-ffplay",
      "--disable-ffprobe",
      "--disable-avdevice",
      "--disable-network",
      "--disable-autodetect",
      "--disable-shared",
      "--disable-everything",
      "--enable-static",
      "--enable-ffmpeg",
      "--enable-avcodec",
      "--enable-avformat",
      "--enable-avfilter",
      "--enable-swresample",
      "--enable-protocol=file,pipe",
      "--enable-demuxer=wav,aiff,mov,mp3,flac,ogg,matroska",
      "--enable-decoder=pcm_s16le,pcm_s16be,pcm_s24le,pcm_s24be,pcm_s32le,pcm_f32le,pcm_f64le,aac,aac_fixed,mp3,mp3float,flac,vorbis,opus,alac",
      "--enable-parser=aac,mpegaudio,flac,vorbis,opus",
      "--enable-encoder=pcm_s16le",
      "--enable-muxer=pcm_s16le",
      "--enable-filter=aresample,aformat,anull",
      "--disable-gpl",
      "--disable-nonfree"
    ];
  await runCommand(path.join(sourceDir, "configure"), configureArgs, { cwd: sourceDir });
  await runCommand("make", [`-j${cpuCount}`], { cwd: sourceDir });
  await runCommand("make", ["install"], { cwd: sourceDir });

  const licenseFiles = ["COPYING.LGPLv2.1", "COPYING.LGPLv3", "LICENSE.md"];
  for (const fileName of licenseFiles) {
    const sourcePath = path.join(sourceDir, fileName);
    if (await pathExists(sourcePath)) {
      await cp(sourcePath, path.join(installDir, fileName));
    }
  }

  await rm(path.join(installDir, "include"), { recursive: true, force: true });
  await rm(path.join(installDir, "lib"), { recursive: true, force: true });
  await rm(path.join(installDir, "share"), { recursive: true, force: true });
  await runCommand(path.join(installDir, "bin", "ffmpeg"), ["-version"]);
}

async function prepareRuntimePayload(runtimeVersion) {
  // // Build and validate the complete private runtime before creating its immutable archive.
  if (!reuseStaging) {
    await rm(runtimeRoot, { recursive: true, force: true });
  }
  await mkdir(runtimeRoot, { recursive: true });
  const pythonPath = path.join(runtimeRoot, "python", "bin", "python3");
  const pythonValidationPath = path.join(runtimeRoot, ".subcreator-python-validated");
  const expectedPythonValidation = `${macArch}:${pythonVersion}`;
  let canReusePython = false;
  if (reuseStaging && (await pathExists(pythonPath)) && (await pathExists(pythonValidationPath))) {
    const validationValue = String(await readFile(pythonValidationPath, "utf8")).trim();
    canReusePython = validationValue === expectedPythonValidation;
  }
  if (canReusePython) {
    // // Reuse a Python payload that already passed the expensive WhisperX import validation.
    await runCommand(pythonPath, ["--version"]);
  } else if (reuseStaging && (await pathExists(pythonPath))) {
    await runCommand(pythonPath, ["-c", "import whisper; import whisperx; print('Reusing private macOS Python runtime')"], {
      env: {
        PYTHONNOUSERSITE: "1",
        PYTHONPATH: ""
      }
    });
    await writeFile(pythonValidationPath, `${expectedPythonValidation}\n`, "ascii");
  } else {
    await preparePrivatePython();
  }

  const ffmpegPath = path.join(runtimeRoot, "ffmpeg", "bin", "ffmpeg");
  if (reuseStaging && (await pathExists(ffmpegPath))) {
    await runCommand(ffmpegPath, ["-version"]);
  } else {
    await preparePrivateFfmpeg();
  }
  await writeFile(path.join(runtimeRoot, ".subcreator-runtime-version"), `${runtimeVersion}\n`, "ascii");
}

async function createRuntimeArchive(runtimeManifest) {
  // // Reuse a valid runtime release asset or build a new one for the current architecture.
  const asset = runtimeManifest.assets[macArch];
  if (!asset || !asset.assetName) {
    throw new Error(`No macOS runtime asset is configured for architecture ${macArch}.`);
  }

  const outputPath = path.join(releasesDir, asset.assetName);
  const expectedHash = String(asset.sha256 || "").trim().toLowerCase();
  if (!rebuildRuntime && expectedHash && (await pathExists(outputPath))) {
    const localHash = await hashFile(outputPath);
    if (localHash !== expectedHash) {
      throw new Error(`Local runtime asset hash does not match ${runtimeManifestPath}.`);
    }
    return runtimeManifest;
  }

  if (!rebuildRuntime && expectedHash && !(await pathExists(outputPath))) {
    const publishedUrl =
      process.env.SUBCREATOR_RUNTIME_DOWNLOAD_URL ||
      `https://github.com/CyrilG93/PremiereSubCreator/releases/download/${runtimeManifest.releaseTag}/${asset.assetName}`;
    await downloadFile(publishedUrl, outputPath);
    const publishedHash = await hashFile(outputPath);
    if (publishedHash !== expectedHash) {
      throw new Error(`Published runtime asset hash does not match ${runtimeManifestPath}.`);
    }
    return runtimeManifest;
  }

  if (process.arch !== macArch) {
    throw new Error("Build the Apple Silicon runtime on an arm64 Mac.");
  }

  await prepareRuntimePayload(runtimeManifest.version);
  await mkdir(releasesDir, { recursive: true });
  await rm(outputPath, { force: true });
  await runCommand("tar", ["-czf", outputPath, "runtime"], { cwd: stagingRoot });
  const sha256 = await hashFile(outputPath);
  const finalizedManifest = {
    ...runtimeManifest,
    assets: {
      ...runtimeManifest.assets,
      [macArch]: {
        ...asset,
        sha256
      }
    }
  };
  await writeRuntimeManifest(finalizedManifest);
  process.stdout.write(`macOS runtime asset created at ${outputPath}\n`);
  return finalizedManifest;
}

function shellQuote(value) {
  // // Quote generated shell values so URLs and paths cannot alter package-script syntax.
  return `'${String(value).replace(/'/g, `'\\''`)}'`;
}

function xmlEscape(value) {
  // // Escape generated distribution strings for productbuild XML.
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function stageFullPayload(runtimeManifest, version) {
  // // Stage the extension, fonts, installer script, private runtime, and ARM64 metadata inside the core package.
  await rm(coreScriptsDir, { recursive: true, force: true });
  await mkdir(path.join(coreScriptsDir, "payload", "dist"), { recursive: true });
  await cp(
    path.join(projectRoot, "dist", "com.cyrilplugin.subcreator"),
    path.join(coreScriptsDir, "payload", "dist", "com.cyrilplugin.subcreator"),
    { recursive: true }
  );
  if (await pathExists(path.join(projectRoot, "Fonts"))) {
    await cp(path.join(projectRoot, "Fonts"), path.join(coreScriptsDir, "payload", "Fonts"), { recursive: true });
  }

  const installScriptSource = path.join(
    projectRoot,
    "installers",
    "subcreator_install_macos_private_runtime.sh"
  );
  const installScriptTarget = path.join(coreScriptsDir, "postinstall");
  await cp(installScriptSource, installScriptTarget);
  await runCommand("chmod", ["755", installScriptTarget]);

  const asset = runtimeManifest.assets[macArch];
  const runtimeAssetPath = path.join(releasesDir, asset.assetName);
  if (!(await pathExists(runtimeAssetPath))) {
    throw new Error(`Runtime asset is missing before PKG creation: ${runtimeAssetPath}`);
  }
  // // Always embed the verified runtime so the public package never acts as a connected updater.
  const bundledRuntimeDir = path.join(coreScriptsDir, "runtime");
  await mkdir(bundledRuntimeDir, { recursive: true });
  await cp(runtimeAssetPath, path.join(bundledRuntimeDir, asset.assetName));
  const runtimeEnv = [
    "# // Generated by subcreator-package-macos-pkg.mjs.",
    `SUBCREATOR_EXTENSION_VERSION=${shellQuote(version)}`,
    `SUBCREATOR_RUNTIME_VERSION=${shellQuote(runtimeManifest.version)}`,
    `SUBCREATOR_RUNTIME_ARCH=${shellQuote(macArch)}`,
    `SUBCREATOR_RUNTIME_ASSET_NAME=${shellQuote(asset.assetName)}`,
    `SUBCREATOR_RUNTIME_SHA256=${shellQuote(String(asset.sha256 || "").toLowerCase())}`,
    ""
  ].join("\n");
  await writeFile(path.join(coreScriptsDir, "runtime.env"), runtimeEnv, "utf8");
  await prunePackagingMetadata(coreScriptsDir);
  await runCommand("xattr", ["-cr", coreScriptsDir]);
}

function createModelInstallScript(model) {
  // // Generate a self-contained package script that downloads one selected model with SHA-256 verification.
  return `#!/bin/bash
set -eu

# // Install the selected Whisper model into the graphical user's cache without removing any other model.
SUBCREATOR_USER="$(stat -f "%Su" /dev/console 2>/dev/null || true)"
if [ -z "\${SUBCREATOR_USER}" ] || [ "\${SUBCREATOR_USER}" = "root" ] || [ "\${SUBCREATOR_USER}" = "loginwindow" ]; then
  SUBCREATOR_USER="\${SUDO_USER:-}"
fi
if [ -z "\${SUBCREATOR_USER}" ] || [ "\${SUBCREATOR_USER}" = "root" ]; then
  echo "Unable to resolve the macOS login user." >&2
  exit 1
fi
SUBCREATOR_UID="$(id -u "\${SUBCREATOR_USER}")"
SUBCREATOR_GID="$(id -g "\${SUBCREATOR_USER}")"
SUBCREATOR_HOME="$(dscl . -read "/Users/\${SUBCREATOR_USER}" NFSHomeDirectory 2>/dev/null | awk '{$1=""; sub(/^ /, ""); print}' || true)"
if [ -z "\${SUBCREATOR_HOME}" ]; then
  SUBCREATOR_HOME="$(eval echo "~\${SUBCREATOR_USER}")"
fi

MODEL_URL=${shellQuote(model.url)}
MODEL_NAME=${shellQuote(model.fileName)}
MODEL_SHA256=${shellQuote(model.sha256)}
MODEL_DIR="\${SUBCREATOR_HOME}/.cache/whisper"
MODEL_PATH="\${MODEL_DIR}/\${MODEL_NAME}"

# // Keep an existing valid model and replace only missing or damaged selected files.
if [ -f "\${MODEL_PATH}" ]; then
  EXISTING_HASH="$(shasum -a 256 "\${MODEL_PATH}" | awk '{print tolower($1)}')"
  if [ "\${EXISTING_HASH}" = "\${MODEL_SHA256}" ]; then
    echo "\${MODEL_NAME} is already available."
    exit 0
  fi
fi

TEMP_PATH="$(mktemp "\${TMPDIR:-/tmp}/subcreator-model.XXXXXX")"
trap 'rm -f "\${TEMP_PATH}"' EXIT
echo "Downloading \${MODEL_NAME}..."
curl --fail --location --retry 3 --connect-timeout 30 "\${MODEL_URL}" --output "\${TEMP_PATH}"
DOWNLOADED_HASH="$(shasum -a 256 "\${TEMP_PATH}" | awk '{print tolower($1)}')"
if [ "\${DOWNLOADED_HASH}" != "\${MODEL_SHA256}" ]; then
  echo "SHA-256 mismatch for \${MODEL_NAME}." >&2
  exit 1
fi

mkdir -p "\${MODEL_DIR}"
mv "\${TEMP_PATH}" "\${MODEL_PATH}"
chown -R "\${SUBCREATOR_UID}:\${SUBCREATOR_GID}" "\${SUBCREATOR_HOME}/.cache"
chmod 644 "\${MODEL_PATH}"
echo "\${MODEL_NAME} installed to \${MODEL_PATH}."
`;
}

async function createComponentPackages(version) {
  // // Build one mandatory core package and optional model packages for the Installer choices page.
  await rm(packagesDir, { recursive: true, force: true });
  await mkdir(packagesDir, { recursive: true });
  const corePackagePath = path.join(packagesDir, "SubCreatorCore.pkg");
  await runCommand("pkgbuild", [
    "--nopayload",
    "--scripts",
    coreScriptsDir,
    "--identifier",
    "com.cyrilplugin.subcreator.installer.core",
    "--version",
    version,
    corePackagePath
  ], {
    env: {
      COPYFILE_DISABLE: "1"
    }
  });

  for (const model of whisperModels) {
    const scriptsDir = path.join(stagingRoot, `model-${model.id}-scripts`);
    await rm(scriptsDir, { recursive: true, force: true });
    await mkdir(scriptsDir, { recursive: true });
    const scriptPath = path.join(scriptsDir, "postinstall");
    await writeFile(scriptPath, createModelInstallScript(model), "utf8");
    await runCommand("chmod", ["755", scriptPath]);
    await runCommand("xattr", ["-cr", scriptsDir]);
    await runCommand("pkgbuild", [
      "--nopayload",
      "--scripts",
      scriptsDir,
      "--identifier",
      `com.cyrilplugin.subcreator.installer.model.${model.id}`,
      "--version",
      version,
      path.join(packagesDir, `SubCreatorModel-${model.id}.pkg`)
    ], {
      env: {
        COPYFILE_DISABLE: "1"
      }
    });
  }
}

async function createDistribution(version) {
  // // Expose optional Whisper model downloads in the standard macOS Installer customization page.
  const distributionPath = path.join(stagingRoot, "Distribution.xml");
  const modelChoices = whisperModels
    .map(
      (model) =>
        `    <line choice="model-${xmlEscape(model.id)}"/>`
    )
    .join("\n");
  const modelDefinitions = whisperModels
    .map(
      (model) => `  <choice id="model-${xmlEscape(model.id)}" title="${xmlEscape(model.title)}" description="${xmlEscape(model.description)}" start_selected="${model.defaultSelected ? "true" : "false"}" selected="subcreatorSelectPreviouslyInstalledModel('${xmlEscape(model.id)}', '${xmlEscape(model.fileName)}')">
    <pkg-ref id="com.cyrilplugin.subcreator.installer.model.${xmlEscape(model.id)}"/>
  </choice>`
    )
    .join("\n");
  const packageRefs = whisperModels
    .map(
      (model) =>
        `  <pkg-ref id="com.cyrilplugin.subcreator.installer.model.${xmlEscape(model.id)}" version="${xmlEscape(version)}" onConclusion="none">SubCreatorModel-${xmlEscape(model.id)}.pkg</pkg-ref>`
    )
    .join("\n");
  const distribution = `<?xml version="1.0" encoding="utf-8"?>
<!-- // Generated by subcreator-package-macos-pkg.mjs. -->
<installer-gui-script minSpecVersion="2">
  <title>Sub Creator</title>
  <organization>com.cyrilplugin.subcreator</organization>
  <domains enable_localSystem="true"/>
  <options customize="always" require-scripts="true" hostArchitectures="${xmlEscape(macArch)}"/>
  <script><![CDATA[
// // Resolve the graphical console user so model detection targets the same Whisper cache used by Sub Creator.
var subcreatorModelSelectionInitialized = {};
function subcreatorConsoleUserName() {
  try {
    var environmentUser = system.env && system.env.USER ? system.env.USER : '';
    if (environmentUser && environmentUser != 'root' && environmentUser != 'loginwindow') {
      return environmentUser;
    }

    var registryRoot = system.ioregistry.fromPath('IOService:/');
    var consoleUsers = registryRoot ? registryRoot['IOConsoleUsers'] : null;
    if (!consoleUsers) {
      return '';
    }

    for (var index = 0; index < consoleUsers.length; index++) {
      var consoleUser = consoleUsers[index];
      var userName = consoleUser ? consoleUser['kCGSSessionUserNameKey'] : '';
      if (userName && userName != 'root' && userName != 'loginwindow') {
        return userName;
      }
    }
  } catch (error) {
    // // Fall back to PKG receipts below when Installer cannot inspect the current console session.
  }

  return '';
}

function subcreatorWhisperModelIsCached(modelFileName) {
  var consoleUser = subcreatorConsoleUserName();
  if (!consoleUser) {
    return false;
  }

  return system.files.fileExistsAtPath('/Users/' + consoleUser + '/.cache/whisper/' + modelFileName);
}

function subcreatorSelectPreviouslyInstalledModel(modelId, modelFileName) {
  if (subcreatorModelSelectionInitialized[modelId]) {
    return my.choice.selected;
  }

  subcreatorModelSelectionInitialized[modelId] = true;
  var upgradeAction = my.choice.packageUpgradeAction;
  var previouslyInstalled = subcreatorWhisperModelIsCached(modelFileName) || upgradeAction == 'installed' || upgradeAction == 'upgrade' || upgradeAction == 'downgrade' || upgradeAction == 'mixed';
  if (previouslyInstalled) {
    return true;
  }

  return my.choice.selected;
}
  ]]></script>
  <choices-outline>
    <line choice="core"/>
${modelChoices}
  </choices-outline>
  <choice id="core" visible="false" start_selected="true">
    <pkg-ref id="com.cyrilplugin.subcreator.installer.core"/>
  </choice>
${modelDefinitions}
  <pkg-ref id="com.cyrilplugin.subcreator.installer.core" version="${xmlEscape(version)}" onConclusion="none">SubCreatorCore.pkg</pkg-ref>
${packageRefs}
</installer-gui-script>
`;
  await writeFile(distributionPath, distribution, "utf8");
  return distributionPath;
}

async function createFullPackage(version, runtimeManifest) {
  // // Build the complete Apple Silicon product package and optionally sign and notarize it.
  await stageFullPayload(runtimeManifest, version);
  await createComponentPackages(version);
  const distributionPath = await createDistribution(version);
  const outputPath = path.join(releasesDir, `SubCreator-v${version}-macOS-Installer-${macArch}.pkg`);
  const unsignedPath = path.join(stagingRoot, `SubCreator-v${version}-macOS-Installer-${macArch}-unsigned.pkg`);
  const signingIdentity = process.env.SUBCREATOR_MAC_INSTALLER_IDENTITY || "";
  const productArgs = [
    "--distribution",
    distributionPath,
    "--package-path",
    packagesDir
  ];

  await mkdir(releasesDir, { recursive: true });
  await rm(outputPath, { force: true });
  await rm(unsignedPath, { force: true });
  if (signingIdentity) {
    productArgs.push("--sign", signingIdentity, outputPath);
  } else {
    productArgs.push(outputPath);
  }
  await runCommand("productbuild", productArgs, {
    env: {
      COPYFILE_DISABLE: "1"
    }
  });

  const notaryProfile = process.env.SUBCREATOR_NOTARY_PROFILE || "";
  if (notaryProfile) {
    await runCommand("xcrun", [
      "notarytool",
      "submit",
      outputPath,
      "--keychain-profile",
      notaryProfile,
      "--wait"
    ]);
    await runCommand("xcrun", ["stapler", "staple", outputPath]);
  }

  if (signingIdentity) {
    // // Verify the Developer ID signature when release signing was requested.
    await runCommand("pkgutil", ["--check-signature", outputPath]);
  } else {
    // // Expand an unsigned local package to confirm its product archive is structurally readable.
    const validationDir = path.join(stagingRoot, "package-validation");
    await rm(validationDir, { recursive: true, force: true });
    await runCommand("pkgutil", ["--expand", outputPath, validationDir]);
    await rm(validationDir, { recursive: true, force: true });
  }
  process.stdout.write(`Full Apple Silicon macOS installer created at ${outputPath}\n`);
}

async function main() {
  // // Build the extension, Apple Silicon runtime asset, and Full macOS Installer package.
  if (process.platform !== "darwin") {
    throw new Error("macOS PKG packaging must run on macOS.");
  }
  if (process.arch !== macArch) {
    throw new Error("macOS PKG packaging requires an Apple Silicon arm64 Mac.");
  }

  const version = await readPackageVersion();
  const runtimeManifest = await readRuntimeManifest();
  if (!runtimeManifest.version || !runtimeManifest.releaseTag) {
    throw new Error(`Incomplete runtime manifest: ${runtimeManifestPath}`);
  }

  if (!reuseStaging) {
    await rm(stagingRoot, { recursive: true, force: true });
  }
  await mkdir(downloadsDir, { recursive: true });
  await mkdir(releasesDir, { recursive: true });
  if (process.env.SUBCREATOR_SKIP_EXTENSION_BUILD !== "1") {
    // // Build the architecture-independent CEP payload unless a verified dist folder is reused.
    await runCommand(npmCommand, ["run", "subcreator:build"]);
  } else if (!(await pathExists(path.join(projectRoot, "dist", "com.cyrilplugin.subcreator")))) {
    throw new Error("SUBCREATOR_SKIP_EXTENSION_BUILD=1 requires an existing dist/com.cyrilplugin.subcreator build.");
  }

  const finalizedRuntimeManifest = await createRuntimeArchive(runtimeManifest);
  const finalizedAsset = finalizedRuntimeManifest.assets[macArch];
  if (!finalizedAsset?.sha256) {
    throw new Error(`Runtime SHA-256 is missing for ${macArch}.`);
  }
  await createFullPackage(version, finalizedRuntimeManifest);
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exit(1);
});
