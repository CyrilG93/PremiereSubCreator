// // Create a local zip package in Releases/ with mandatory installer files.
import { cp, mkdir, readFile, rm, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawn } from "node:child_process";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const distExtensionDir = path.join(projectRoot, "dist", "com.cyrilg93.subcreator");
const releasesDir = path.join(projectRoot, "Releases");
const stagingRoot = path.join(projectRoot, ".subcreator-release-staging");

function runCommand(command, args) {
  // // Execute platform-specific archive tooling and capture failures.
  return new Promise((resolve, reject) => {
    const processHandle = spawn(command, args, {
      cwd: projectRoot,
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
  if (process.platform === "darwin") {
    await runCommand("ditto", ["-c", "-k", "--sequesterRsrc", "--keepParent", sourceDir, outputZip]);
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

  await runCommand("zip", ["-r", outputZip, path.basename(sourceDir)]);
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

  await rm(zipPath, { force: true });
  await createZipFromDirectory(stagingBundleDir, zipPath);
  await rm(stagingRoot, { recursive: true, force: true });

  process.stdout.write(`Release zip created at ${zipPath}\n`);
}

subcreatorPackageRelease().catch((error) => {
  process.stderr.write(`${error}\n`);
  process.exit(1);
});
