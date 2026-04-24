// // Install the local build into CEP extension folders for quick testing.
import { cp, mkdir, rename, rm, stat } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const extensionSource = path.join(projectRoot, "dist", "com.cyrilplugin.subcreator");

async function subcreatorPathExists(targetPath) {
  // // Check optional legacy install locations without turning missing folders into install failures.
  try {
    await stat(targetPath);
    return true;
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return false;
    }
    throw error;
  }
}

function getCepExtensionDir() {
  // // Resolve OS-specific CEP extensions path for user-level installation.
  const homeDir = os.homedir();
  if (process.platform === "darwin") {
    return path.join(homeDir, "Library", "Application Support", "Adobe", "CEP", "extensions");
  }

  if (process.platform === "win32") {
    const appData = process.env.APPDATA || path.join(homeDir, "AppData", "Roaming");
    return path.join(appData, "Adobe", "CEP", "extensions");
  }

  throw new Error("Unsupported platform for CEP installation.");
}

async function subcreatorInstallDev() {
  // // Verify build exists before attempting installation.
  await stat(extensionSource);

  const destinationRoot = getCepExtensionDir();
  const destination = path.join(destinationRoot, "com.cyrilplugin.subcreator");
  const legacyDestination = path.join(destinationRoot, "com.cyrilg93.subcreator");

  await mkdir(destinationRoot, { recursive: true });
  if ((await subcreatorPathExists(legacyDestination)) && !(await subcreatorPathExists(destination))) {
    // // Rename the old CEP folder once so local updates move to the new bundle id.
    await rename(legacyDestination, destination);
  }
  await rm(legacyDestination, { recursive: true, force: true });
  // // Remove the previous dev install first so deleted build files cannot linger in CEP.
  await rm(destination, { recursive: true, force: true });
  await cp(extensionSource, destination, { recursive: true, force: true });

  process.stdout.write(`Installed extension to ${destination}\n`);
}

subcreatorInstallDev().catch((error) => {
  process.stderr.write(`${error}\n`);
  process.exit(1);
});
