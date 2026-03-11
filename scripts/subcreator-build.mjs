// // Build the CEP extension payload into dist/com.cyrilg93.subcreator.
import { mkdir, rm, cp, readdir, writeFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const sourceRoot = path.join(projectRoot, "src");
const distRoot = path.join(projectRoot, "dist", "com.cyrilg93.subcreator");
const templatesRoot = path.join(projectRoot, "templates", "mogrt");

function subcreatorSlugify(input) {
  // // Build stable ids safe for JSON and DOM usage.
  return input
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}

function subcreatorDetectPreviewClass(name) {
  // // Map template names to visual preview styles in the gallery.
  const lower = name.toLowerCase();
  if (lower.includes("clean")) {
    return "clean";
  }
  if (lower.includes("comic")) {
    return "comic";
  }
  if (lower.includes("glitch")) {
    return "glitch";
  }
  if (lower.includes("karaoke")) {
    return "karaoke";
  }
  if (lower.includes("typewriter")) {
    return "typewriter";
  }
  if (lower.includes("mr beast")) {
    return "mrbeast";
  }
  if (lower.includes("tiktok")) {
    return "tiktok";
  }
  if (lower.includes("akira")) {
    return "akira";
  }
  if (lower.includes("motion blur")) {
    return "motionblur";
  }
  if (lower.includes("marker")) {
    return "marker";
  }
  if (lower.includes("slide")) {
    return "slide";
  }
  if (lower.includes("slant")) {
    return "slant";
  }
  if (lower.includes("spinning")) {
    return "spinning";
  }
  if (lower.includes("block")) {
    return "block";
  }
  if (lower.includes("emphasis")) {
    return "emphasis";
  }
  if (lower.includes("obviously")) {
    return "obviously";
  }
  if (lower.includes("arch")) {
    return "arch";
  }
  return "default";
}

async function subcreatorPathExists(targetPath) {
  // // Keep missing template folders non-fatal for early project setup.
  try {
    await stat(targetPath);
    return true;
  } catch (error) {
    return false;
  }
}

async function subcreatorScanMogrt(dir, rootDir, collector) {
  // // Recursively discover all .mogrt files for gallery generation.
  const entries = await readdir(dir, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await subcreatorScanMogrt(fullPath, rootDir, collector);
      continue;
    }

    if (!entry.isFile() || !entry.name.toLowerCase().endsWith(".mogrt")) {
      continue;
    }

    const relativePath = path.relative(rootDir, fullPath).split(path.sep).join("/");
    const pathParts = relativePath.split("/");
    const aspect = pathParts.length > 1 ? pathParts[0] : "General";
    const name = path.basename(entry.name, ".mogrt");
    const id = `${subcreatorSlugify(aspect)}-${subcreatorSlugify(name)}-${collector.length + 1}`;

    collector.push({
      id,
      name,
      aspect,
      relativePath,
      previewClass: subcreatorDetectPreviewClass(name)
    });
  }
}

async function subcreatorBuildMogrtCatalog(distAssetsDir, distTemplatesDir) {
  // // Copy templates into extension bundle and emit catalog consumed by panel UI.
  const templates = [];
  const templateRootExists = await subcreatorPathExists(templatesRoot);

  if (templateRootExists) {
    await cp(templatesRoot, distTemplatesDir, { recursive: true });
    await subcreatorScanMogrt(templatesRoot, templatesRoot, templates);
  }

  templates.sort((left, right) => {
    const aspectCompare = left.aspect.localeCompare(right.aspect);
    if (aspectCompare !== 0) {
      return aspectCompare;
    }
    return left.name.localeCompare(right.name);
  });

  const catalog = {
    generatedAt: new Date().toISOString(),
    templateCount: templates.length,
    templates
  };

  await writeFile(path.join(distAssetsDir, "mogrt-catalog.json"), JSON.stringify(catalog, null, 2), "utf8");
}

async function subcreatorBuild() {
  // // Clean old build output to avoid stale files in extension bundles.
  await rm(distRoot, { recursive: true, force: true });

  // // Prepare final extension folders expected by CEP.
  await Promise.all([
    mkdir(path.join(distRoot, "js"), { recursive: true }),
    mkdir(path.join(distRoot, "host"), { recursive: true }),
    mkdir(path.join(distRoot, "CSXS"), { recursive: true }),
    mkdir(path.join(distRoot, "locales"), { recursive: true }),
    mkdir(path.join(distRoot, "assets"), { recursive: true }),
    mkdir(path.join(distRoot, "templates", "mogrt"), { recursive: true })
  ]);

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

  await subcreatorBuildMogrtCatalog(path.join(distRoot, "assets"), path.join(distRoot, "templates", "mogrt"));

  // // Provide explicit build completion output for scripts and CI.
  process.stdout.write(`Built extension at ${distRoot}\n`);
}

subcreatorBuild().catch((error) => {
  process.stderr.write(`${error}\n`);
  process.exit(1);
});
