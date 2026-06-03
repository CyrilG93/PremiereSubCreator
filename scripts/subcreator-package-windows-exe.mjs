// // Build a Windows self-extracting installer with a private Python and FFmpeg runtime.
import { cp, mkdir, readFile, rm, stat, writeFile } from "node:fs/promises";
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
const releasesDir = path.join(projectRoot, "Releases");
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
const privatePythonEnv = {
  PYTHONUTF8: "1",
  PYTHONNOUSERSITE: "1",
  PYTHONPATH: ""
};

function runCommand(command, args, options = {}) {
  // // Execute packaging tools with inherited output so long downloads and pip installs remain visible.
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
      "Get-ChildItem -LiteralPath $root -Recurse -File -Include *.lib,*.pdb -ErrorAction SilentlyContinue | Remove-Item -Force;",
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
    "openai-whisper",
    "whisperx",
    "requests",
    "nltk",
    "certifi",
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
  await runCommand(pythonExe, ["-c", "import sys, whisper, whisperx; print('Private Python runtime OK'); print('\\n'.join(sys.path))"], {
    env: privatePythonEnv
  });
}

async function prepareFfmpegRuntime() {
  // // Bundle a private FFmpeg binary so Whisper can run without a system PATH dependency.
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
      `$root = Split-Path -Parent (Split-Path -Parent $source.FullName);`,
      `$target = ${JSON.stringify(runtimeFfmpegDir)};`,
      "New-Item -ItemType Directory -Path (Join-Path $target 'bin') -Force | Out-Null;",
      "Copy-Item -Path (Join-Path $root 'bin\\*') -Destination (Join-Path $target 'bin') -Recurse -Force;",
      "Get-ChildItem -LiteralPath $root -File | Where-Object { $_.Name -match '^(LICENSE|COPYING|README|VERSION)' } | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $target -Force }"
    ].join(" ")
  ]);

  await validateFfmpegRuntime();
}

async function validateFfmpegRuntime() {
  // // Validate FFmpeg from a temp copy because some Windows policies deny direct execution from hidden staging paths.
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

async function copyReleasePayload() {
  // // Stage the extension, models, fonts, installer script, and private runtime payload.
  await mkdir(path.join(payloadRoot, "dist"), { recursive: true });
  await mkdir(path.join(payloadRoot, "installers"), { recursive: true });

  await cp(path.join(projectRoot, "README.md"), path.join(payloadRoot, "README.md"));
  await cp(path.join(projectRoot, "dist", "com.cyrilplugin.subcreator"), path.join(payloadRoot, "dist", "com.cyrilplugin.subcreator"), {
    recursive: true
  });
  await cp(
    path.join(projectRoot, "installers", "subcreator_install_windows_private_runtime.ps1"),
    path.join(payloadRoot, "installers", "subcreator_install_windows_private_runtime.ps1")
  );

  if (await pathExists(path.join(projectRoot, "Models"))) {
    await cp(path.join(projectRoot, "Models"), path.join(payloadRoot, "Models"), { recursive: true });
  }

  if (await pathExists(path.join(projectRoot, "Fonts"))) {
    await cp(path.join(projectRoot, "Fonts"), path.join(payloadRoot, "Fonts"), { recursive: true });
  }
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
  // // Install Inno Setup into the build staging folder so packaging does not depend on a global developer install.
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

async function createInnoInstaller(version) {
  // // Generate a single Windows installer EXE using Inno Setup for reliable large-payload handling.
  const compilerPath = await prepareInnoCompiler();
  const outputBaseName = `SubCreator-v${version}-Windows-PrivateRuntime`;
  const scriptPath = path.join(installerRoot, "SubCreatorPrivateRuntime.iss");

  await mkdir(installerRoot, { recursive: true });
  await mkdir(releasesDir, { recursive: true });
  await rm(path.join(releasesDir, `${outputBaseName}.exe`), { force: true });

  const iss = [
    "; // Generated by subcreator-package-windows-exe.mjs.",
    "[Setup]",
    "AppId={{1A10D6BA-247F-4E13-A48F-5AB0C0E11300}",
    "AppName=Sub Creator",
    `AppVersion=${version}`,
    "AppPublisher=Cyril Plugin",
    "DefaultDirName={localappdata}\\SubCreator\\InstallerPayload",
    "DisableDirPage=yes",
    "DisableProgramGroupPage=yes",
    "Uninstallable=no",
    "PrivilegesRequired=lowest",
    "ArchitecturesAllowed=x64compatible",
    "ArchitecturesInstallIn64BitMode=x64compatible",
    "Compression=lzma2/ultra64",
    "SolidCompression=yes",
    "WizardStyle=modern",
    `OutputDir=${escapeInnoString(releasesDir)}`,
    `OutputBaseFilename=${outputBaseName}`,
    "",
    "[Files]",
    `Source: "${escapeInnoString(path.join(payloadRoot, "*"))}"; DestDir: "{tmp}\\SubCreatorPayload"; Flags: recursesubdirs createallsubdirs ignoreversion`,
    "",
    "[Run]",
    'Filename: "{sys}\\WindowsPowerShell\\v1.0\\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File ""{tmp}\\SubCreatorPayload\\installers\\subcreator_install_windows_private_runtime.ps1"" -PayloadRoot ""{tmp}\\SubCreatorPayload"""; StatusMsg: "Installing Sub Creator private runtime..."; Flags: waituntilterminated',
    ""
  ].join("\r\n");

  await writeFile(scriptPath, iss, "utf8");
  await runCommand(compilerPath, ["/Qp", scriptPath]);
  process.stdout.write(`Windows EXE installer created at ${path.join(releasesDir, `${outputBaseName}.exe`)}\n`);
}

async function readPackageVersion() {
  // // Use package.json as the installer version source so artifacts match the extension metadata.
  const raw = await readFile(path.join(projectRoot, "package.json"), "utf8");
  const parsed = JSON.parse(raw);
  return String(parsed.version || "").trim();
}

async function main() {
  // // Build the extension and all private runtime pieces before creating the final EXE.
  if (process.platform !== "win32") {
    throw new Error("Windows EXE packaging must run on Windows.");
  }

  const version = await readPackageVersion();
  if (!reuseStaging) {
    await rm(stagingRoot, { recursive: true, force: true });
  }
  await mkdir(downloadsDir, { recursive: true });
  await mkdir(runtimeRoot, { recursive: true });

  await runCommand(npmCommand, ["run", "subcreator:build"], { shell: process.platform === "win32" });
  await copyReleasePayload();
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
  await createInnoInstaller(version);
}

main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exit(1);
});
