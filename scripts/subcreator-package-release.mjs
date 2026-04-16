// // Create a local zip package in Releases/ with mandatory installer files.
import { cp, mkdir, readdir, readFile, rm, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const distExtensionDir = path.join(projectRoot, "dist", "com.cyrilg93.subcreator");
const distMetaPath = path.join(distExtensionDir, "assets", "subcreator-meta.json");
const distManifestPath = path.join(distExtensionDir, "CSXS", "manifest.xml");
const bundledModelsDir = path.join(projectRoot, "Models");
const bundledFontsDir = path.join(projectRoot, "Fonts");
const releasesDir = path.join(projectRoot, "Releases");
const stagingRoot = path.join(projectRoot, ".subcreator-release-staging");

function runCommand(command, args, commandCwd = projectRoot) {
  // // Execute platform-specific archive tooling and capture failures.
  return new Promise((resolve, reject) => {
    const processHandle = spawn(command, args, {
      cwd: commandCwd,
      stdio: "inherit"
    });

    processHandle.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`${command} exited with code ${code}`));
    });
  });
}

async function createZipFromDirectory(sourceDir, outputZip) {
  // // Use native archivers on each OS to avoid extra dependencies.
  const sourceParent = path.dirname(sourceDir);
  const sourceName = path.basename(sourceDir);
  if (process.platform === "darwin") {
    await runCommand("zip", ["-r", "-X", outputZip, sourceName], sourceParent);
    return;
  }

  if (process.platform === "win32") {
    const escapedSource = sourceDir.replace(/\\/g, "\\\\");
    const escapedOutput = outputZip.replace(/\\/g, "\\\\");
    await runCommand("powershell", [
      "-NoProfile",
      "-Command",
      `Compress-Archive -Path '${escapedSource}\\*' -DestinationPath '${escapedOutput}' -Force`
    ]);
    return;
  }

  await runCommand("zip", ["-r", "-X", outputZip, sourceName], sourceParent);
}

async function subcreatorPruneReleaseMetadata(targetDir) {
  // // Remove macOS metadata files so the release archive only contains install payload.
  const entries = await readdir(targetDir, { withFileTypes: true });
  for (const entry of entries) {
    const entryPath = path.join(targetDir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === "__MACOSX") {
        await rm(entryPath, { recursive: true, force: true });
        continue;
      }
      await subcreatorPruneReleaseMetadata(entryPath);
      continue;
    }

    if (entry.name === ".DS_Store" || entry.name.startsWith("._")) {
      await rm(entryPath, { force: true });
    }
  }
}

async function copyBundledModelPayload(sourceDir, targetDir) {
  // // Copy only release-safe bundled model assets so large local source files do not exceed GitHub's regular blob limit.
  const entries = await readdir(sourceDir, { withFileTypes: true });
  const chunkedModelNames = new Set();

  for (const entry of entries) {
    if (!entry.isFile()) {
      continue;
    }

    const chunkMatch = entry.name.match(/^(.*\.pt)\.part-\d+$/i);
    if (chunkMatch) {
      chunkedModelNames.add(chunkMatch[1]);
    }
  }

  const copyTasks = [];
  for (const entry of entries) {
    if (!entry.isFile()) {
      continue;
    }

    if (entry.name === ".DS_Store") {
      continue;
    }

    const sourcePath = path.join(sourceDir, entry.name);
    const targetPath = path.join(targetDir, entry.name);
    if (/\.pt\.part-\d+$/i.test(entry.name)) {
      copyTasks.push(cp(sourcePath, targetPath));
      continue;
    }

    if (/\.pt$/i.test(entry.name)) {
      if (chunkedModelNames.has(entry.name)) {
        continue;
      }

      const entryStat = await stat(sourcePath);
      if (entryStat.size >= 100 * 1024 * 1024) {
        continue;
      }

      copyTasks.push(cp(sourcePath, targetPath));
    }
  }

  if (copyTasks.length < 1) {
    return;
  }

  await mkdir(targetDir, { recursive: true });
  await Promise.all(copyTasks);
}

async function subcreatorPackageRelease() {
  // // Validate build output exists before packaging release assets.
  await stat(distExtensionDir);

  const packageJsonRaw = await readFile(path.join(projectRoot, "package.json"), "utf8");
  const packageJson = JSON.parse(packageJsonRaw);
  const version = packageJson.version;

  // // Refuse to package if `dist` still carries another version than the current project metadata.
  const distMetaRaw = await readFile(distMetaPath, "utf8");
  const distMeta = JSON.parse(distMetaRaw);
  const distMetaVersion = String(distMeta.version || "").trim();
  if (distMetaVersion !== version) {
    throw new Error(`dist metadata version mismatch: package.json=${version} dist=${distMetaVersion || "<empty>"}`);
  }

  const distManifestRaw = await readFile(distManifestPath, "utf8");
  const distManifestMatch = distManifestRaw.match(/ExtensionBundleVersion="([^"]+)"/);
  const distManifestVersion = String(distManifestMatch?.[1] || "").trim();
  if (distManifestVersion !== version) {
    throw new Error(`dist manifest version mismatch: package.json=${version} dist=${distManifestVersion || "<empty>"}`);
  }

  const bundleName = `SubCreator-v${version}`;
  const stagingBundleDir = path.join(stagingRoot, bundleName);
  const zipPath = path.join(releasesDir, `${bundleName}.zip`);

  await rm(stagingRoot, { recursive: true, force: true });
  await mkdir(stagingBundleDir, { recursive: true });
  await mkdir(releasesDir, { recursive: true });

  // // Copy only mandatory installation payload: extension, installers, README, and bundled Whisper models when available.
  const copyTasks = [
    cp(path.join(projectRoot, "README.md"), path.join(stagingBundleDir, "README.md")),
    cp(path.join(projectRoot, "installers"), path.join(stagingBundleDir, "installers"), { recursive: true }),
    cp(path.join(projectRoot, "dist"), path.join(stagingBundleDir, "dist"), { recursive: true })
  ];
  try {
    const bundledModelsStat = await stat(bundledModelsDir);
    if (bundledModelsStat.isDirectory()) {
      copyTasks.push(copyBundledModelPayload(bundledModelsDir, path.join(stagingBundleDir, "Models")));
    }
  } catch {
    // // Ignore missing bundled-model folders so packaging stays compatible with older working copies.
  }
  try {
    const bundledFontsStat = await stat(bundledFontsDir);
    if (bundledFontsStat.isDirectory()) {
      // // Include optional bundled font archives in the release package when the workspace provides them.
      copyTasks.push(cp(bundledFontsDir, path.join(stagingBundleDir, "Fonts"), { recursive: true }));
    }
  } catch {
    // // Ignore missing bundled-font folders so packaging stays compatible with older working copies.
  }

  await Promise.all(copyTasks);

  await subcreatorPruneReleaseMetadata(stagingBundleDir);

  await rm(zipPath, { force: true });
  await createZipFromDirectory(stagingBundleDir, zipPath);
  await rm(stagingRoot, { recursive: true, force: true });

  process.stdout.write(`Release zip created at ${zipPath}\n`);
}

subcreatorPackageRelease().catch((error) => {
  process.stderr.write(`${error}\n`);
  process.exit(1);
});
