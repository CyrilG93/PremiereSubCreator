// // Build light and full Windows installers plus a separately downloadable private runtime.
import { createReadStream, createWriteStream } from "node:fs";
import { cp, mkdir, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const stagingRoot = path.join(projectRoot, ".subcreator-exe-staging");
const downloadsDir = path.join(stagingRoot, "downloads");
const payloadRoot = path.join(stagingRoot, "payload");
const runtimeRoot = path.join(payloadRoot, "runtime");
const installerRoot = path.join(stagingRoot, "installer");
const bundledModelsRoot = path.join(stagingRoot, "bundled-models");
const releasesDir = path.join(projectRoot, "Releases");
const runtimeManifestPath = path.join(projectRoot, "installers", "windows-runtime.json");
const pythonVersion = process.env.SUBCREATOR_PRIVATE_PYTHON_VERSION || "3.11.9";
const pythonShortVersion = pythonVersion.split(".").slice(0, 2).join("");
const pythonEmbedUrl =
  process.env.SUBCREATOR_PYTHON_EMBED_URL ||
  `https://www.python.org/ftp/python/${pythonVersion}/python-${pythonVersion}-embed-amd64.zip`;
const getPipUrl = process.env.SUBCREATOR_GET_PIP_URL || "https://bootstrap.pypa.io/get-pip.py";
const ffmpegZipUrl =
  process.env.SUBCREATOR_FFMPEG_ZIP_URL ||
  "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-lgpl.zip";
const innoSetupUrl =
  process.env.SUBCREATOR_INNO_SETUP_URL ||
  "https://github.com/jrsoftware/issrc/releases/download/is-6_7_3/innosetup-6.7.3.exe";
const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";
const reuseStaging = process.env.SUBCREATOR_REUSE_STAGING === "1";
const rebuildRuntime = process.env.SUBCREATOR_REBUILD_RUNTIME === "1";
const skipRuntimeAssetDownload = process.env.SUBCREATOR_SKIP_RUNTIME_ASSET_DOWNLOAD === "1";
const lightOnly = process.env.SUBCREATOR_LIGHT_ONLY === "1";
const privatePythonEnv = {
  PYTHONUTF8: "1",
  PYTHONNOUSERSITE: "1",
  PYTHONPATH: ""
};
const whisperModels = [
  {
    id: "tiny",
    label: "Tiny - fastest, lower accuracy (about 75 MB)",
    fileName: "tiny.pt",
    sha256: "65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9",
    url: "https://openaipublic.azureedge.net/main/whisper/models/65147644a518d12f04e32d6f3b26facc3f8dd46e5390956a9424a650c0ce22b9/tiny.pt",
    defaultSelected: false
  },
  {
    id: "base",
    label: "Base - recommended starter model (about 142 MB)",
    fileName: "base.pt",
    sha256: "ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e",
    url: "https://openaipublic.azureedge.net/main/whisper/models/ed3a0b6b1c0edf879ad9b11b1af5a0e6ab5db9205f891f668f8b0e6c6326e34e/base.pt",
    defaultSelected: true
  },
  {
    id: "small",
    label: "Small - good quality and speed balance (about 466 MB)",
    fileName: "small.pt",
    sha256: "9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794",
    url: "https://openaipublic.azureedge.net/main/whisper/models/9ecf779972d90ba49c06d968637d720dd632c55bbf19d441fb42bf17a411e794/small.pt",
    defaultSelected: false
  },
  {
    id: "medium",
    label: "Medium - better for difficult audio (about 1.5 GB)",
    fileName: "medium.pt",
    sha256: "345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1",
    url: "https://openaipublic.azureedge.net/main/whisper/models/345ae4da62f9b3d59415adc60127b97c714f32e89e936602e85993674d08dcb1/medium.pt",
    defaultSelected: false
  },
  {
    id: "largeV3",
    label: "Large v3 - best accuracy, slowest (about 2.9 GB)",
    fileName: "large-v3.pt",
    sha256: "e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb",
    url: "https://openaipublic.azureedge.net/main/whisper/models/e5b1a55b89c1367dacf97e3e19bfd829a01529dbfdeefa8caeb59b3f1b81dadb/large-v3.pt",
    defaultSelected: false
  }
];

function runCommand(command, args, options = {}) {
  // // Execute packaging tools with inherited output so long downloads and compiles remain visible.
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: options.cwd || projectRoot,
      env: {
        ...process.env,
        ...(options.env || {})
      },
      shell: Boolean(options.shell),
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
  // // Probe optional local downloads and payload folders without throwing.
  try {
    await stat(targetPath);
    return true;
  } catch {
    return false;
  }
}

async function hashFile(targetPath) {
  // // Stream large release assets through SHA-256 without loading them fully into memory.
  return new Promise((resolve, reject) => {
    const hash = createHash("sha256");
    const stream = createReadStream(targetPath);
    stream.on("error", reject);
    stream.on("data", (chunk) => hash.update(chunk));
    stream.on("end", () => resolve(hash.digest("hex")));
  });
}

async function prepareBundledBaseModel() {
  // // Assemble and validate the repository's split base model for the offline Full installer.
  const modelsRoot = path.join(projectRoot, "Models");
  const directModelPath = path.join(modelsRoot, "base.pt");
  const targetPath = path.join(bundledModelsRoot, "base.pt");
  await mkdir(bundledModelsRoot, { recursive: true });

  if (await pathExists(directModelPath)) {
    await cp(directModelPath, targetPath);
  } else {
    const parts = (await readdir(modelsRoot))
      .filter((name) => /^base\.pt\.part-\d+$/i.test(name))
      .sort((left, right) => left.localeCompare(right, "en", { numeric: true }));
    if (parts.length === 0) {
      throw new Error("The Full installer requires Models/base.pt or its split base.pt.part-* files.");
    }

    const output = createWriteStream(targetPath);
    try {
      for (const part of parts) {
        // // Stream each part in order so the 142 MB model is never buffered fully in Node.js memory.
        for await (const chunk of createReadStream(path.join(modelsRoot, part))) {
          if (!output.write(chunk)) {
            await new Promise((resolve) => output.once("drain", resolve));
          }
        }
      }
      await new Promise((resolve, reject) => {
        output.once("error", reject);
        output.end(resolve);
      });
    } catch (error) {
      output.destroy();
      throw error;
    }
  }

  const baseModel = whisperModels.find((model) => model.id === "base");
  const actualHash = await hashFile(targetPath);
  if (!baseModel || actualHash !== baseModel.sha256) {
    throw new Error(`Bundled base model SHA-256 mismatch: ${actualHash}`);
  }
  return targetPath;
}

async function downloadFile(url, targetPath) {
  // // Download third-party runtime archives only when the local cache does not already contain them.
  if (await pathExists(targetPath)) {
    return;
  }

  await mkdir(path.dirname(targetPath), { recursive: true });
  await runCommand("powershell", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    `$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri ${JSON.stringify(url)} -OutFile ${JSON.stringify(targetPath)}`
  ]);
}

async function expandArchive(zipPath, targetDir) {
  // // Use PowerShell's native ZIP extraction so the build has no extra archive dependency.
  await rm(targetDir, { recursive: true, force: true });
  await mkdir(targetDir, { recursive: true });
  await runCommand("powershell", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    `Expand-Archive -LiteralPath ${JSON.stringify(zipPath)} -DestinationPath ${JSON.stringify(targetDir)} -Force`
  ]);
}

async function configureEmbeddedPython(runtimePythonDir) {
  // // Enable pip installation while keeping user site-packages disabled through the build environment.
  const pthPath = path.join(runtimePythonDir, `python${pythonShortVersion}._pth`);
  const pthLines = [`python${pythonShortVersion}.zip`, ".", "Lib\\site-packages", "import site", ""];
  await mkdir(path.join(runtimePythonDir, "Lib", "site-packages"), { recursive: true });
  await mkdir(path.join(runtimePythonDir, "Scripts"), { recursive: true });
  await writeFile(pthPath, pthLines.join("\r\n"), "utf8");
}

async function lockEmbeddedPythonRuntime(runtimePythonDir) {
  // // Remove `import site` after packaging so the installed runtime never reads the user's Python profile.
  const pthPath = path.join(runtimePythonDir, `python${pythonShortVersion}._pth`);
  const pthLines = [`python${pythonShortVersion}.zip`, ".", "Lib\\site-packages", ""];
  await writeFile(pthPath, pthLines.join("\r\n"), "utf8");
}

async function prunePythonRuntime(runtimePythonDir) {
  // // Remove build-time library/import artifacts from wheels; runtime DLLs and Python modules stay in place.
  await runCommand("powershell", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    [
      `$root = ${JSON.stringify(runtimePythonDir)};`,
      "Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in '.lib', '.pdb' } | Remove-Item -Force;",
      "$torchInclude = Join-Path $root 'Lib\\site-packages\\torch\\include';",
      "if (Test-Path -LiteralPath $torchInclude) { Remove-Item -LiteralPath $torchInclude -Recurse -Force; }",
      "Get-ChildItem -LiteralPath $root -Recurse -Directory -Filter __pycache__ -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue"
    ].join(" ")
  ]);
}

async function preparePythonRuntime() {
  // // Create a clean private Python runtime instead of copying the developer machine's global Python packages.
  const pythonZip = path.join(downloadsDir, `python-${pythonVersion}-embed-amd64.zip`);
  const getPipPath = path.join(downloadsDir, "get-pip.py");
  const runtimePythonDir = path.join(runtimeRoot, "python");
  const pythonExe = path.join(runtimePythonDir, "python.exe");

  await downloadFile(pythonEmbedUrl, pythonZip);
  await expandArchive(pythonZip, runtimePythonDir);
  await configureEmbeddedPython(runtimePythonDir);

  await downloadFile(getPipUrl, getPipPath);
  await runCommand(pythonExe, [getPipPath, "--no-warn-script-location"], {
    env: privatePythonEnv
  });
  await runCommand(pythonExe, ["-m", "pip", "install", "--upgrade", "--no-cache-dir", "setuptools", "wheel"], {
    env: privatePythonEnv
  });
  await runCommand(pythonExe, [
    "-m",
    "pip",
    "install",
    "--upgrade",
    "--no-cache-dir",
    "--no-warn-script-location",
    "torch==2.8.0+cpu",
    "torchaudio==2.8.0+cpu",
    "torchvision==0.23.0+cpu",
    "--index-url",
    "https://download.pytorch.org/whl/cpu"
  ], {
    env: privatePythonEnv
  });
  await runCommand(pythonExe, [
    "-m",
    "pip",
    "install",
    "--upgrade",
    "--no-cache-dir",
    "--no-warn-script-location",
    "openai-whisper==20250625",
    "whisperx==3.8.6",
    "transformers==4.57.6",
    "requests==2.34.2",
    "nltk==3.9.4",
    "certifi==2026.6.17",
    "--extra-index-url",
    "https://download.pytorch.org/whl/cpu"
  ], {
    env: privatePythonEnv
  });

  await prunePythonRuntime(runtimePythonDir);
  await lockEmbeddedPythonRuntime(runtimePythonDir);
  await validatePythonRuntime();
}

async function validatePythonRuntime() {
  // // Confirm the private Python runtime imports core packages without leaking user site-packages.
  const pythonExe = path.join(runtimeRoot, "python", "python.exe");
  const validationCode =
    "import sys, whisper, whisperx; print('Private Python runtime OK'); print('\\n'.join(sys.path))";
  await runCommand("powershell", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    [
      `$python = ${JSON.stringify(pythonExe)};`,
      "if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { throw \"Private Python executable is missing: $python\" }",
      "Unblock-File -LiteralPath $python -ErrorAction SilentlyContinue;",
      "& $python -c $env:SUBCREATOR_PYTHON_VALIDATION_CODE;",
      "exit $LASTEXITCODE"
    ].join(" ")
  ], {
    env: {
      ...privatePythonEnv,
      SUBCREATOR_PYTHON_VALIDATION_CODE: validationCode
    }
  });
}

async function prepareFfmpegRuntime() {
  // // Bundle a private LGPL FFmpeg binary so Whisper does not depend on the system PATH.
  const localFfmpegZip = process.env.SUBCREATOR_FFMPEG_ZIP || "";
  const ffmpegZip = localFfmpegZip || path.join(downloadsDir, path.basename(new URL(ffmpegZipUrl).pathname));
  const extractedDir = path.join(stagingRoot, "ffmpeg-extracted");
  const runtimeFfmpegDir = path.join(runtimeRoot, "ffmpeg");

  if (!localFfmpegZip) {
    await downloadFile(ffmpegZipUrl, ffmpegZip);
  }

  await expandArchive(ffmpegZip, extractedDir);
  await runCommand("powershell", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    [
      `$source = Get-ChildItem -LiteralPath ${JSON.stringify(extractedDir)} -Recurse -File -Filter ffmpeg.exe | Select-Object -First 1;`,
      "if (-not $source) { throw 'ffmpeg.exe not found in archive' }",
      "$root = Split-Path -Parent (Split-Path -Parent $source.FullName);",
      `$target = ${JSON.stringify(runtimeFfmpegDir)};`,
      "New-Item -ItemType Directory -Path (Join-Path $target 'bin') -Force | Out-Null;",
      "Copy-Item -Path (Join-Path $root 'bin\\*') -Destination (Join-Path $target 'bin') -Recurse -Force;",
      "Get-ChildItem -LiteralPath $root -File | Where-Object { $_.Name -match '^(LICENSE|COPYING|README|VERSION)' } | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $target -Force }"
    ].join(" ")
  ]);

  await validateFfmpegRuntime();
}

async function validateFfmpegRuntime() {
  // // Validate FFmpeg from a temp copy because some Windows policies deny execution from hidden staging paths.
  const runtimeFfmpegExe = path.join(runtimeRoot, "ffmpeg", "bin", "ffmpeg.exe");
  const tempFfmpegExe = path.join(process.env.TEMP || stagingRoot, "subcreator-ffmpeg-lgpl-test.exe");

  if (!(await pathExists(runtimeFfmpegExe))) {
    throw new Error(`FFmpeg executable missing from runtime payload: ${runtimeFfmpegExe}`);
  }

  await cp(runtimeFfmpegExe, tempFfmpegExe);
  await runCommand("powershell", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    `Unblock-File -LiteralPath ${JSON.stringify(tempFfmpegExe)} -ErrorAction SilentlyContinue`
  ]);
  await runCommand(tempFfmpegExe, ["-version"]);
}

async function copyConnectedPayload() {
  // // Stage only the extension, README, and runtime installer script inside the lightweight EXE.
  await mkdir(path.join(payloadRoot, "dist"), { recursive: true });
  await mkdir(path.join(payloadRoot, "installers"), { recursive: true });

  await cp(path.join(projectRoot, "README.md"), path.join(payloadRoot, "README.md"));
  await rm(path.join(payloadRoot, "dist", "com.cyrilplugin.subcreator"), { recursive: true, force: true });
  await cp(path.join(projectRoot, "dist", "com.cyrilplugin.subcreator"), path.join(payloadRoot, "dist", "com.cyrilplugin.subcreator"), {
    recursive: true
  });
  await cp(
    path.join(projectRoot, "installers", "subcreator_install_windows_private_runtime.ps1"),
    path.join(payloadRoot, "installers", "subcreator_install_windows_private_runtime.ps1")
  );
}

async function findExistingInnoCompiler() {
  // // Prefer an explicit compiler path, then the private staging install, then common machine installs.
  const candidates = [
    process.env.SUBCREATOR_ISCC_PATH || "",
    path.join(stagingRoot, "tools", "Inno Setup 6", "ISCC.exe"),
    path.join(stagingRoot, "tools", "Inno", "ISCC.exe"),
    "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe",
    "C:\\Program Files\\Inno Setup 6\\ISCC.exe"
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (await pathExists(candidate)) {
      return candidate;
    }
  }

  return "";
}

async function prepareInnoCompiler() {
  // // Install Inno Setup into staging so packaging does not depend on a global developer installation.
  const existingCompiler = await findExistingInnoCompiler();
  if (existingCompiler) {
    return existingCompiler;
  }

  const installerPath = path.join(downloadsDir, "innosetup.exe");
  const installDir = path.join(stagingRoot, "tools", "Inno");
  await downloadFile(innoSetupUrl, installerPath);
  await mkdir(installDir, { recursive: true });
  await runCommand("powershell", [
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    [
      `$installer = ${JSON.stringify(installerPath)};`,
      `$installDir = ${JSON.stringify(installDir)};`,
      "$args = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/CURRENTUSER', ('/DIR=' + $installDir));",
      "$process = Start-Process -FilePath $installer -ArgumentList $args -Wait -PassThru -WindowStyle Hidden;",
      "exit $process.ExitCode"
    ].join(" ")
  ]);

  const compilerPath = (await findExistingInnoCompiler()) || path.join(installDir, "ISCC.exe");
  if (!(await pathExists(compilerPath))) {
    throw new Error(`Inno Setup compiler missing after install: ${compilerPath}`);
  }
  return compilerPath;
}

function escapeInnoString(value) {
  // // Escape double quotes for Inno Setup string literals.
  return String(value || "").replace(/"/g, '""');
}

function escapePascalString(value) {
  // // Escape apostrophes for generated Inno Setup Pascal string literals.
  return String(value || "").replace(/'/g, "''");
}

async function readRuntimeManifest() {
  // // Keep the runtime asset URL and immutable hash stable across lightweight plugin releases.
  const raw = await readFile(runtimeManifestPath, "utf8");
  const parsed = JSON.parse(raw);
  return {
    version: String(parsed.version || "").trim(),
    releaseTag: String(parsed.releaseTag || "").trim(),
    assetName: String(parsed.assetName || "").trim(),
    sha256: String(parsed.sha256 || "").trim().toLowerCase()
  };
}

async function writeRuntimeManifest(manifest) {
  // // Persist the first compiled runtime hash so later installers reuse exactly the published binary.
  await writeFile(runtimeManifestPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
}

async function createRuntimeInstaller(compilerPath, runtimeManifest) {
  // // Package the large private runtime separately so normal extension updates stay small.
  const outputPath = path.join(releasesDir, runtimeManifest.assetName);
  if (!rebuildRuntime && runtimeManifest.sha256 && skipRuntimeAssetDownload) {
    // // CI can compile from verified immutable metadata even if an unrelated local runtime artifact is stale.
    process.stdout.write(`Reusing published Windows runtime metadata for ${runtimeManifest.assetName}.\n`);
    return runtimeManifest;
  }

  if (!rebuildRuntime && runtimeManifest.sha256 && (await pathExists(outputPath))) {
    const localHash = await hashFile(outputPath);
    if (localHash === runtimeManifest.sha256) {
      return runtimeManifest;
    }
    throw new Error(`Local runtime asset hash does not match ${runtimeManifestPath}.`);
  }

  if (!rebuildRuntime && runtimeManifest.sha256 && !(await pathExists(outputPath))) {
    const publishedUrl =
      process.env.SUBCREATOR_RUNTIME_DOWNLOAD_URL ||
      `https://github.com/CyrilG93/PremiereSubCreator/releases/download/${runtimeManifest.releaseTag}/${runtimeManifest.assetName}`;
    await downloadFile(publishedUrl, outputPath);
    const publishedHash = await hashFile(outputPath);
    if (publishedHash !== runtimeManifest.sha256) {
      throw new Error(`Published runtime asset hash does not match ${runtimeManifestPath}.`);
    }
    return runtimeManifest;
  }

  await writeFile(path.join(runtimeRoot, ".subcreator-runtime-version"), `${runtimeManifest.version}\r\n`, "ascii");
  const scriptPath = path.join(installerRoot, "SubCreatorRuntime.iss");
  const runtimeOutputBaseName = path.parse(runtimeManifest.assetName).name;
  const iss = [
    "; // Generated by subcreator-package-windows-exe.mjs.",
    "[Setup]",
    "AppId={{7387927E-4878-49FC-A762-93B85C6013A1}",
    "AppName=Sub Creator Private Runtime",
    `AppVersion=${runtimeManifest.version}`,
    "AppPublisher=Cyril Plugin",
    "DefaultDirName={localappdata}\\SubCreator\\runtime",
    "DisableDirPage=yes",
    "DisableProgramGroupPage=yes",
    "Uninstallable=no",
    "PrivilegesRequired=lowest",
    "RestartIfNeededByRun=no",
    "ArchitecturesAllowed=x64compatible",
    "ArchitecturesInstallIn64BitMode=x64compatible",
    "Compression=lzma2/ultra64",
    "SolidCompression=yes",
    "WizardStyle=modern dynamic",
    `OutputDir=${escapeInnoString(releasesDir)}`,
    `OutputBaseFilename=${runtimeOutputBaseName}`,
    "",
    "[InstallDelete]",
    'Type: filesandordirs; Name: "{localappdata}\\SubCreator\\runtime"',
    "",
    "[Files]",
    `Source: "${escapeInnoString(path.join(runtimeRoot, "*"))}"; DestDir: "{localappdata}\\SubCreator\\runtime"; Flags: recursesubdirs createallsubdirs ignoreversion`,
    ""
  ].join("\r\n");

  await mkdir(installerRoot, { recursive: true });
  await mkdir(releasesDir, { recursive: true });
  await rm(outputPath, { force: true });
  await writeFile(scriptPath, iss, "utf8");
  await runCommand(compilerPath, ["/Qp", scriptPath]);

  const sha256 = await hashFile(outputPath);
  const finalizedManifest = { ...runtimeManifest, sha256 };
  await writeRuntimeManifest(finalizedManifest);
  process.stdout.write(`Windows runtime asset created at ${outputPath}\n`);
  return finalizedManifest;
}

function createModelPascalDefinitions(bundledBaseModelPath) {
  // // Generate explicit model functions because Inno Setup check callbacks cannot capture dynamic arrays.
  return whisperModels.flatMap((model, index) => {
    const suffix = model.id[0].toUpperCase() + model.id.slice(1);
    return [
      `function ShouldInstall${suffix}Model: Boolean;`,
      "begin",
      `  Result := Download${suffix}Model;`,
      "end;",
      "",
      `function ${suffix}ModelPath: String;`,
      "begin",
      `  Result := ExpandConstant('{%USERPROFILE}\\.cache\\whisper\\${model.fileName}');`,
      "end;",
      "",
      `procedure Prepare${suffix}ModelDownload;`,
      "begin",
      `  Download${suffix}Model := ModelPage.Values[${index}] and (not FileHasHash(${suffix}ModelPath, '${model.sha256}'));`,
      `  if Download${suffix}Model then`,
      ...(model.id === "base" && bundledBaseModelPath
        ? ["    Log('The verified base model is embedded in this Full installer.');"]
        : [`    DownloadPage.Add('${escapePascalString(model.url)}', '${model.fileName}', '${model.sha256}');`]),
      "end;",
      ""
    ];
  });
}

async function createUserInstaller(compilerPath, version, runtimeManifest, mode) {
  // // Create either a connected lightweight installer or a complete installer with the runtime embedded.
  const includeRuntime = mode === "full";
  const bundledBaseModelPath = includeRuntime ? await prepareBundledBaseModel() : "";
  const outputBaseName = `SubCreator-v${version}-Windows-${includeRuntime ? "Full" : "Light"}-Installer`;
  const outputPath = path.join(releasesDir, `${outputBaseName}.exe`);
  const scriptPath = path.join(installerRoot, `SubCreator${includeRuntime ? "Full" : "Light"}.iss`);
  const runtimeUrl =
    process.env.SUBCREATOR_RUNTIME_DOWNLOAD_URL ||
    `https://github.com/CyrilG93/PremiereSubCreator/releases/download/${runtimeManifest.releaseTag}/${runtimeManifest.assetName}`;
  const modelVariables = whisperModels.map((model) => {
    const suffix = model.id[0].toUpperCase() + model.id.slice(1);
    return `  Download${suffix}Model: Boolean;`;
  });
  const modelPageItems = whisperModels.flatMap((model, index) => [
    `  ModelPage.Add('${escapePascalString(model.label)}${model.id === "base" && bundledBaseModelPath ? " - included" : ""}');`,
    `  ModelPage.Values[${index}] := ${model.defaultSelected ? "True" : `FileExists(ExpandConstant('{%USERPROFILE}\\.cache\\whisper\\${model.fileName}'))`};`
  ]);
  const prepareModelDownloads = whisperModels.map((model) => {
    const suffix = model.id[0].toUpperCase() + model.id.slice(1);
    return `    Prepare${suffix}ModelDownload;`;
  });
  const networkDownloadConditions = [
    ...(!includeRuntime ? ["DownloadRuntime"] : []),
    ...whisperModels
      .filter((model) => !(model.id === "base" && bundledBaseModelPath))
      .map((model) => `Download${model.id[0].toUpperCase() + model.id.slice(1)}Model`)
  ];
  const modelFileEntries = whisperModels.map((model) => {
    const suffix = model.id[0].toUpperCase() + model.id.slice(1);
    if (model.id === "base" && bundledBaseModelPath) {
      return `Source: "${escapeInnoString(bundledBaseModelPath)}"; DestDir: "{%USERPROFILE}\\.cache\\whisper"; Flags: ignoreversion; Check: ShouldInstall${suffix}Model`;
    }
    return `Source: "{tmp}\\${model.fileName}"; DestDir: "{%USERPROFILE}\\.cache\\whisper"; Flags: external ignoreversion; Check: ShouldInstall${suffix}Model`;
  });
  const iss = [
    "; // Generated by subcreator-package-windows-exe.mjs.",
    "[Setup]",
    "AppId={{1A10D6BA-247F-4E13-A48F-5AB0C0E11300}",
    "AppName=Sub Creator",
    `AppVersion=${version}`,
    "AppPublisher=Cyril Plugin",
    "DefaultDirName={localappdata}\\SubCreator\\InstallerPayload",
    "CreateAppDir=no",
    "DisableDirPage=yes",
    "DisableProgramGroupPage=yes",
    "Uninstallable=no",
    "PrivilegesRequired=lowest",
    "RestartIfNeededByRun=no",
    "ArchitecturesAllowed=x64compatible",
    "ArchitecturesInstallIn64BitMode=x64compatible",
    "Compression=lzma2/ultra64",
    "SolidCompression=yes",
    "WizardStyle=modern dynamic",
    `OutputDir=${escapeInnoString(releasesDir)}`,
    `OutputBaseFilename=${outputBaseName}`,
    "",
    "[Files]",
    `Source: "${escapeInnoString(path.join(payloadRoot, "README.md"))}"; DestDir: "{tmp}\\SubCreatorPayload"; Flags: ignoreversion`,
    `Source: "${escapeInnoString(path.join(payloadRoot, "dist", "com.cyrilplugin.subcreator", "*"))}"; DestDir: "{tmp}\\SubCreatorPayload\\dist\\com.cyrilplugin.subcreator"; Flags: recursesubdirs createallsubdirs ignoreversion`,
    ...(includeRuntime
      ? [`Source: "${escapeInnoString(path.join(runtimeRoot, "*"))}"; DestDir: "{tmp}\\SubCreatorPayload\\runtime"; Flags: recursesubdirs createallsubdirs ignoreversion`]
      : []),
    ...modelFileEntries,
    `Source: "${escapeInnoString(path.join(payloadRoot, "installers", "subcreator_install_windows_private_runtime.ps1"))}"; DestDir: "{tmp}\\SubCreatorPayload\\installers"; Flags: ignoreversion; AfterInstall: InstallSubCreator`,
    "",
    "[Code]",
    "var",
    "  ModelPage: TInputOptionWizardPage;",
    "  DownloadPage: TDownloadWizardPage;",
    "  ReadyWarningLabel: TNewStaticText;",
    "  InstallEstimateLabel: TNewStaticText;",
    "  InstallWarningLabel: TNewStaticText;",
    "  DownloadRuntime: Boolean;",
    ...modelVariables,
    "",
    "function FileHasHash(const FileName, ExpectedHash: String): Boolean;",
    "begin",
    "  Result := FileExists(FileName);",
    "  if Result then",
    "  begin",
    "    try",
    "      Result := CompareText(GetSHA256OfFile(FileName), ExpectedHash) = 0;",
    "    except",
    "      Result := False;",
    "    end;",
    "  end;",
    "end;",
    "",
    "function RuntimeIsCurrent: Boolean;",
    "var",
    "  InstalledVersion: AnsiString;",
    "  RuntimeRoot: String;",
    "  VersionFile: String;",
    "begin",
    "  RuntimeRoot := ExpandConstant('{localappdata}\\SubCreator\\runtime');",
    "  VersionFile := RuntimeRoot + '\\.subcreator-runtime-version';",
    "  Result := FileExists(VersionFile) and",
    "    LoadStringFromFile(VersionFile, InstalledVersion) and",
    `      (Trim(String(InstalledVersion)) = '${escapePascalString(runtimeManifest.version)}') and`,
    "      FileExists(RuntimeRoot + '\\python\\python.exe') and",
    "      FileExists(RuntimeRoot + '\\python\\Scripts\\whisper.exe') and",
    "      FileExists(RuntimeRoot + '\\ffmpeg\\bin\\ffmpeg.exe');",
    "end;",
    "",
    ...createModelPascalDefinitions(bundledBaseModelPath),
    "procedure InstallSubCreator;",
    "var",
    "  ResultCode: Integer;",
    "  RuntimeInstallerPath: String;",
    "  PowerShellPath: String;",
    "  InstallerScriptPath: String;",
    "  PayloadPath: String;",
    "  Parameters: String;",
    "begin",
    ...(includeRuntime
      ? []
      : [
          "  if DownloadRuntime then",
          "  begin",
          `    RuntimeInstallerPath := ExpandConstant('{tmp}\\${escapePascalString(runtimeManifest.assetName)}');`,
          "    WizardForm.StatusLabel.Caption := 'Preparing the private Whisper runtime. Windows security checks can take a moment...';",
          "    if not Exec(RuntimeInstallerPath, '/SILENT /SUPPRESSMSGBOXES /NORESTART /CURRENTUSER', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then",
          "      RaiseException('Windows could not start the private runtime installer.');",
          "    if ResultCode <> 0 then",
          "      RaiseException(Format('The private runtime installer failed with exit code %d.', [ResultCode]));",
          "  end;",
          ""
        ]),
    "  WizardForm.StatusLabel.Caption := 'Installing Sub Creator and validating its dependencies...';",
    "  PowerShellPath := ExpandConstant('{sys}\\WindowsPowerShell\\v1.0\\powershell.exe');",
    "  InstallerScriptPath := ExpandConstant('{tmp}\\SubCreatorPayload\\installers\\subcreator_install_windows_private_runtime.ps1');",
    "  PayloadPath := ExpandConstant('{tmp}\\SubCreatorPayload');",
    `  Parameters := '-NoProfile -ExecutionPolicy Bypass -File "' + InstallerScriptPath + '" -PayloadRoot "' + PayloadPath + '"${includeRuntime ? "" : " -SkipRuntimeInstall"} -RuntimeVersion "${escapePascalString(runtimeManifest.version)}"';`,
    "  if not Exec(PowerShellPath, Parameters, '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then",
    "    RaiseException('Windows could not start the Sub Creator installation script.');",
    "  if ResultCode <> 0 then",
    "    RaiseException(Format('Sub Creator dependency validation failed with exit code %d. Please keep this message for support.', [ResultCode]));",
    "end;",
    "",
    "procedure InitializeWizard;",
    "begin",
    "  { // Let users choose extra models without deleting models they already have. }",
    "  ModelPage := CreateInputOptionPage(wpWelcome,",
    "    'Whisper models', 'Choose the models to keep available offline',",
    "    'Only missing or damaged selected models will be downloaded. Existing models are never removed.',",
    "    False, False);",
    ...modelPageItems,
    "",
    "  DownloadPage := CreateDownloadPage('Downloading Sub Creator files',",
    "    'The first installation can take several minutes. Plugin-only updates stay lightweight.', nil);",
    "  DownloadPage.ShowBaseNameInsteadOfUrl := True;",
    "",
    "  { // Replace the scrollable Ready memo with one integrated, wrapped warning label. }",
    "  WizardForm.ReadyMemo.Visible := False;",
    "  ReadyWarningLabel := TNewStaticText.Create(WizardForm);",
    "  ReadyWarningLabel.Parent := WizardForm.ReadyPage;",
    "  ReadyWarningLabel.Left := WizardForm.ReadyMemo.Left;",
    "  ReadyWarningLabel.Top := WizardForm.ReadyMemo.Top;",
    "  ReadyWarningLabel.Width := WizardForm.ReadyMemo.Width;",
    "  ReadyWarningLabel.AutoSize := False;",
    "  ReadyWarningLabel.WordWrap := True;",
    "  ReadyWarningLabel.Height := ScaleY(58);",
    "  ReadyWarningLabel.Font.Style := [fsBold];",
    "  ReadyWarningLabel.Caption := 'Important: after you click Install, this window may appear frozen for a few seconds while Windows checks the installation files. This is normal; please wait.';",
    "",
    "  { // Explain long Windows checks and visible helper terminals during the install phase. }",
    "  InstallEstimateLabel := TNewStaticText.Create(WizardForm);",
    "  InstallEstimateLabel.Parent := WizardForm.InstallingPage;",
    "  InstallEstimateLabel.Left := WizardForm.ProgressGauge.Left;",
    "  InstallEstimateLabel.Top := WizardForm.ProgressGauge.Top + WizardForm.ProgressGauge.Height + ScaleY(18);",
    "  InstallEstimateLabel.Width := WizardForm.ProgressGauge.Width;",
    "  InstallEstimateLabel.AutoSize := False;",
    "  InstallEstimateLabel.WordWrap := True;",
    "  InstallEstimateLabel.Height := ScaleY(32);",
    "  InstallEstimateLabel.Caption := 'Estimated time remaining: usually 2 to 5 minutes. Windows security checks may make this longer.';",
    "",
    "  InstallWarningLabel := TNewStaticText.Create(WizardForm);",
    "  InstallWarningLabel.Parent := WizardForm.InstallingPage;",
    "  InstallWarningLabel.Left := WizardForm.ProgressGauge.Left;",
    "  InstallWarningLabel.Top := InstallEstimateLabel.Top + InstallEstimateLabel.Height + ScaleY(8);",
    "  InstallWarningLabel.Width := WizardForm.ProgressGauge.Width;",
    "  InstallWarningLabel.AutoSize := False;",
    "  InstallWarningLabel.WordWrap := True;",
    "  InstallWarningLabel.Height := ScaleY(42);",
    "  InstallWarningLabel.Font.Style := [fsBold];",
    "  InstallWarningLabel.Caption := 'Important: do not close any Command Prompt or PowerShell windows that open. They are required to complete the installation.';",
    "end;",
    "",
    "function NextButtonClick(CurPageID: Integer): Boolean;",
    "var",
    "  Error: String;",
    "begin",
    "  if CurPageID = wpReady then",
    "  begin",
    "    DownloadPage.Clear;",
    `    DownloadRuntime := ${includeRuntime ? "False" : "not RuntimeIsCurrent"};`,
    ...(includeRuntime
      ? []
      : [
          "    if DownloadRuntime then",
          `      DownloadPage.Add('${escapePascalString(runtimeUrl)}', '${runtimeManifest.assetName}', '${runtimeManifest.sha256}');`
        ]),
    ...prepareModelDownloads,
    "",
    `    if ${networkDownloadConditions.join(" or ")} then`,
    "    begin",
    "      DownloadPage.Show;",
    "      try",
    "        try",
    "          DownloadPage.Download;",
    "          Result := True;",
    "        except",
    "          if DownloadPage.AbortedByUser then",
    "            Log('Download aborted by user.')",
    "          else",
    "          begin",
    "            Error := Format('%s: %s', [DownloadPage.LastBaseNameOrUrl, GetExceptionMessage]);",
    "            SuppressibleMsgBox(AddPeriod(Error), mbCriticalError, MB_OK, IDOK);",
    "          end;",
    "          Result := False;",
    "        end;",
    "      finally",
    "        DownloadPage.Hide;",
    "      end;",
    "    end",
    "    else",
    "      Result := True;",
    "  end",
    "  else",
    "    Result := True;",
    "end;",
    ""
  ].join("\r\n");

  await mkdir(installerRoot, { recursive: true });
  await mkdir(releasesDir, { recursive: true });
  await rm(outputPath, { force: true });
  await writeFile(scriptPath, iss, "utf8");
  await runCommand(compilerPath, ["/Qp", scriptPath]);
  process.stdout.write(`${includeRuntime ? "Full" : "Light"} Windows installer created at ${outputPath}\n`);
}

async function readPackageVersion() {
  // // Use package.json as the installer version source so artifacts match the extension metadata.
  const raw = await readFile(path.join(projectRoot, "package.json"), "utf8");
  const parsed = JSON.parse(raw);
  return String(parsed.version || "").trim();
}

async function main() {
  // // Build both release assets, reusing the immutable runtime whenever its manifest already has a hash.
  if (process.platform !== "win32") {
    throw new Error("Windows EXE packaging must run on Windows.");
  }

  const version = await readPackageVersion();
  const runtimeManifest = await readRuntimeManifest();
  if (!runtimeManifest.version || !runtimeManifest.releaseTag || !runtimeManifest.assetName) {
    throw new Error(`Incomplete runtime manifest: ${runtimeManifestPath}`);
  }

  if (!reuseStaging) {
    await rm(stagingRoot, { recursive: true, force: true });
  }
  await mkdir(downloadsDir, { recursive: true });
  await mkdir(runtimeRoot, { recursive: true });
  await mkdir(releasesDir, { recursive: true });

  await runCommand(npmCommand, ["run", "subcreator:build"], { shell: process.platform === "win32" });
  await copyConnectedPayload();

  if (!runtimeManifest.sha256 || rebuildRuntime) {
    if (reuseStaging && (await pathExists(path.join(runtimeRoot, "python", "python.exe")))) {
      await validatePythonRuntime();
    } else {
      await preparePythonRuntime();
    }
    if (reuseStaging && (await pathExists(path.join(runtimeRoot, "ffmpeg", "bin", "ffmpeg.exe")))) {
      await validateFfmpegRuntime();
    } else {
      await prepareFfmpegRuntime();
    }
  }

  const compilerPath = await prepareInnoCompiler();
  const finalizedRuntimeManifest = await createRuntimeInstaller(compilerPath, runtimeManifest);
  await createUserInstaller(compilerPath, version, finalizedRuntimeManifest, "light");
  if (!lightOnly && (await pathExists(path.join(runtimeRoot, "python", "python.exe")))) {
    await createUserInstaller(compilerPath, version, finalizedRuntimeManifest, "full");
  } else if (!lightOnly) {
    process.stdout.write("Full Windows installer skipped because the unpacked runtime is not available.\n");
  } else {
    process.stdout.write("Full Windows installer skipped because SUBCREATOR_LIGHT_ONLY=1.\n");
  }
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exit(1);
});
