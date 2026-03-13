// // Create a local zip package in Releases/ with mandatory installer files.
import { cp, mkdir, readdir, readFile, rm, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const distExtensionDir = path.join(projectRoot, "dist", "com.cyrilg93.subcreator");
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

async function subcreatorPackageRelease() {
  // // Validate build output exists before packaging release assets.
  await stat(distExtensionDir);

  const packageJsonRaw = await readFile(path.join(projectRoot, "package.json"), "utf8");
  const packageJson = JSON.parse(packageJsonRaw);
  const version = packageJson.version;

  const bundleName = `SubCreator-v${version}`;
  const stagingBundleDir = path.join(stagingRoot, bundleName);
  const zipPath = path.join(releasesDir, `${bundleName}.zip`);

  await rm(stagingRoot, { recursive: true, force: true });
  await mkdir(stagingBundleDir, { recursive: true });
  await mkdir(releasesDir, { recursive: true });

  // // Copy only mandatory installation payload: extension, installers, and README.
  await Promise.all([
    cp(path.join(projectRoot, "README.md"), path.join(stagingBundleDir, "README.md")),
    cp(path.join(projectRoot, "installers"), path.join(stagingBundleDir, "installers"), { recursive: true }),
    cp(path.join(projectRoot, "dist"), path.join(stagingBundleDir, "dist"), { recursive: true })
  ]);

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
