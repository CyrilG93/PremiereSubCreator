// // Build the CEP extension payload into dist/com.cyrilg93.subcreator.
import { mkdir, rm, cp, readdir, writeFile, stat, copyFile, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { unzipSync } from "fflate";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, "..");
const sourceRoot = path.join(projectRoot, "src");
const distRoot = path.join(projectRoot, "dist", "com.cyrilg93.subcreator");
const templatesRoot = path.join(projectRoot, "templates", "mogrt");
const packageJsonPath = path.join(projectRoot, "package.json");
const releaseRepoSlug = "CyrilG93/PremiereSubCreator";

function subcreatorSlugify(input) {
  // // Build stable ids safe for JSON and DOM usage.
  return input
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}

function subcreatorSanitizeSegment(input) {
  // // Keep packaged template paths ASCII-safe for cross-platform importMGT.
  const ascii = String(input)
    .normalize("NFKD")
    .replace(/[^\x00-\x7F]/g, "")
    .trim();

  const cleaned = ascii
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "");

  return cleaned || "template";
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

function subcreatorPickArchiveEntry(entryNames, ruleMatchers) {
  // // Resolve preferred archive asset order while keeping original entry casing.
  const pairs = entryNames.map((name) => {
    return {
      raw: name,
      lower: String(name).replace(/\\/g, "/").toLowerCase()
    };
  });

  for (const matcher of ruleMatchers) {
    const hit = pairs.find((entry) => matcher.test(entry.lower));
    if (hit) {
      return hit.raw;
    }
  }

  return "";
}

function subcreatorNormalizeArchiveExt(entryName, fallbackExt) {
  // // Normalize archive extension for stable output filenames.
  const ext = path.extname(entryName || "").toLowerCase();
  if (!ext) {
    return fallbackExt;
  }

  if (ext === ".jpeg") {
    return ".jpg";
  }

  return ext;
}

async function subcreatorExtractMogrtPreviewAssets(sourcePath, outputDir, outputStem) {
  // // Extract embedded thumbnail assets from .mogrt (zip) for accurate gallery previews.
  let archiveMap = {};

  try {
    const archiveBytes = await readFile(sourcePath);
    archiveMap = unzipSync(new Uint8Array(archiveBytes));
  } catch {
    return {
      imageFileName: "",
      videoFileName: ""
    };
  }

  const entryNames = Object.keys(archiveMap);
  if (!entryNames.length) {
    return {
      imageFileName: "",
      videoFileName: ""
    };
  }

  const imageEntry = subcreatorPickArchiveEntry(entryNames, [/\/thumb\.png$/i, /thumb\.png$/i, /thumb\.(jpg|jpeg|webp)$/i, /\.(png|jpg|jpeg|webp)$/i]);
  const videoEntry = subcreatorPickArchiveEntry(entryNames, [/\/thumb\.mp4$/i, /thumb\.mp4$/i, /\.mp4$/i]);

  let imageFileName = "";
  let videoFileName = "";

  if (imageEntry && archiveMap[imageEntry]) {
    const imageExt = subcreatorNormalizeArchiveExt(imageEntry, ".png");
    imageFileName = `${outputStem}${imageExt}`;
    await writeFile(path.join(outputDir, imageFileName), Buffer.from(archiveMap[imageEntry]));
  }

  if (videoEntry && archiveMap[videoEntry]) {
    const videoExt = subcreatorNormalizeArchiveExt(videoEntry, ".mp4");
    videoFileName = `${outputStem}${videoExt}`;
    await writeFile(path.join(outputDir, videoFileName), Buffer.from(archiveMap[videoEntry]));
  }

  return {
    imageFileName,
    videoFileName
  };
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

    collector.push({
      sourcePath: fullPath,
      name,
      aspect,
      previewClass: subcreatorDetectPreviewClass(name)
    });
  }
}

async function subcreatorBuildMogrtCatalog(distAssetsDir, distTemplatesDir) {
  // // Copy templates into extension bundle with ASCII-safe paths and emit gallery catalog.
  const discovered = [];
  const templates = [];
  const previewRootDir = path.join(distAssetsDir, "mogrt-previews");
  const previewRootRelativePath = "assets/mogrt-previews";
  const templateRootExists = await subcreatorPathExists(templatesRoot);

  if (templateRootExists) {
    await subcreatorScanMogrt(templatesRoot, templatesRoot, discovered);
  }

  discovered.sort((left, right) => {
    const aspectCompare = left.aspect.localeCompare(right.aspect);
    if (aspectCompare !== 0) {
      return aspectCompare;
    }
    return left.name.localeCompare(right.name);
  });

  const dedupe = new Map();

  for (let index = 0; index < discovered.length; index += 1) {
    const item = discovered[index];
    const safeAspect = subcreatorSanitizeSegment(item.aspect);
    const safeName = subcreatorSanitizeSegment(item.name).toLowerCase();
    const dedupeKey = `${safeAspect}/${safeName}`;
    const nextCount = (dedupe.get(dedupeKey) ?? 0) + 1;
    dedupe.set(dedupeKey, nextCount);

    const safeFileName = `${safeName}-${nextCount}.mogrt`;
    const relativePath = `${safeAspect}/${safeFileName}`;
    const destDir = path.join(distTemplatesDir, safeAspect);
    const destPath = path.join(destDir, safeFileName);
    const previewDir = path.join(previewRootDir, safeAspect);
    const previewStem = `${safeName}-${nextCount}`;

    await mkdir(destDir, { recursive: true });
    await mkdir(previewDir, { recursive: true });
    await copyFile(item.sourcePath, destPath);
    const previewAssets = await subcreatorExtractMogrtPreviewAssets(item.sourcePath, previewDir, previewStem);

    templates.push({
      id: `${subcreatorSlugify(item.aspect)}-${subcreatorSlugify(item.name)}-${index + 1}`,
      name: item.name,
      aspect: item.aspect,
      relativePath,
      previewClass: item.previewClass,
      previewImagePath: previewAssets.imageFileName ? `${previewRootRelativePath}/${safeAspect}/${previewAssets.imageFileName}` : "",
      previewVideoPath: previewAssets.videoFileName ? `${previewRootRelativePath}/${safeAspect}/${previewAssets.videoFileName}` : ""
    });
  }

  const catalog = {
    generatedAt: new Date().toISOString(),
    templateCount: templates.length,
    templates
  };

  await writeFile(path.join(distAssetsDir, "mogrt-catalog.json"), JSON.stringify(catalog, null, 2), "utf8");
}

async function subcreatorReadPackageVersion() {
  // // Read extension version from package.json so panel UI stays in sync.
  try {
    const raw = await readFile(packageJsonPath, "utf8");
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed.version === "string" && parsed.version.trim().length > 0) {
      return parsed.version.trim();
    }
  } catch {
    return "0.0.0";
  }

  return "0.0.0";
}

async function subcreatorBuildPanelMeta(distAssetsDir) {
  // // Emit panel metadata used for version label and update notifications.
  const version = await subcreatorReadPackageVersion();
  const meta = {
    generatedAt: new Date().toISOString(),
    version,
    repository: releaseRepoSlug,
    releaseApiUrl: `https://api.github.com/repos/${releaseRepoSlug}/releases/latest`,
    releasePageUrl: `https://github.com/${releaseRepoSlug}/releases/latest`
  };

  await writeFile(path.join(distAssetsDir, "subcreator-meta.json"), JSON.stringify(meta, null, 2), "utf8");
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
  await subcreatorBuildPanelMeta(path.join(distRoot, "assets"));

  // // Provide explicit build completion output for scripts and CI.
  process.stdout.write(`Built extension at ${distRoot}\n`);
}

subcreatorBuild().catch((error) => {
  process.stderr.write(`${error}\n`);
  process.exit(1);
});
