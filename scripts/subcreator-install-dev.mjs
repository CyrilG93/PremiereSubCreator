// // Install the local build into CEP extension folders for quick testing.
import { cp, mkdir, stat } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const extensionSource = path.join(projectRoot, "dist", "com.cyrilg93.subcreator");

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
  const destination = path.join(destinationRoot, "com.cyrilg93.subcreator");

  await mkdir(destinationRoot, { recursive: true });
  await cp(extensionSource, destination, { recursive: true, force: true });

  process.stdout.write(`Installed extension to ${destination}\n`);
}

subcreatorInstallDev().catch((error) => {
  process.stderr.write(`${error}\n`);
  process.exit(1);
});
