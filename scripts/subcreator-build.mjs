// // Build the CEP extension payload into dist/com.cyrilg93.subcreator.
import { mkdir, rm, cp } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const sourceRoot = path.join(projectRoot, "src");
const distRoot = path.join(projectRoot, "dist", "com.cyrilg93.subcreator");

async function subcreatorBuild() {
  // // Clean old build output to avoid stale files in extension bundles.
  await rm(distRoot, { recursive: true, force: true });

  // // Prepare final extension folders expected by CEP.
  await mkdir(path.join(distRoot, "js"), { recursive: true });
  await mkdir(path.join(distRoot, "host"), { recursive: true });
  await mkdir(path.join(distRoot, "CSXS"), { recursive: true });
  await mkdir(path.join(distRoot, "locales"), { recursive: true });

  // // Bundle panel TypeScript into a single browser-ready module.
  await build({
    entryPoints: [path.join(sourceRoot, "panel", "main.ts")],
    bundle: true,
    format: "esm",
    platform: "browser",
    target: ["chrome92"],
    outfile: path.join(distRoot, "js", "main.js")
  });

  // // Copy static panel resources and host scripts.
  await Promise.all([
    cp(path.join(sourceRoot, "panel", "index.html"), path.join(distRoot, "index.html")),
    cp(path.join(sourceRoot, "panel", "styles.css"), path.join(distRoot, "styles.css")),
    cp(path.join(sourceRoot, "host", "SubCreatorHost.jsx"), path.join(distRoot, "host", "SubCreatorHost.jsx")),
    cp(path.join(sourceRoot, "host", "manifest.xml"), path.join(distRoot, "CSXS", "manifest.xml")),
    cp(path.join(sourceRoot, "locales", "fr.json"), path.join(distRoot, "locales", "fr.json")),
    cp(path.join(sourceRoot, "locales", "en.json"), path.join(distRoot, "locales", "en.json"))
  ]);

  // // Provide explicit build completion output for scripts and CI.
  process.stdout.write(`Built extension at ${distRoot}\n`);
}

subcreatorBuild().catch((error) => {
  process.stderr.write(`${error}\n`);
  process.exit(1);
});
