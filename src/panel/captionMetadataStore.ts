// // Persist subtitle word-timing metadata locally so the Text editor can reuse precise timings after generation.
import type { CaptionCue, CaptionWord } from "../core/types";
import { buildWeightedCaptionWords } from "../core/wordTiming";
import type { TextEditorBlock } from "../core/textEditor";
import type { SelectedMogrtTextItem } from "./cepBridge";

interface CepNodeModules {
  fs: {
    existsSync: (path: string) => boolean;
    mkdirSync: (path: string, options: { recursive: boolean }) => void;
    readFileSync: (path: string, encoding?: string) => string;
    writeFileSync: (path: string, data: string) => void;
  };
  os: {
    homedir: () => string;
  };
  path: {
    join: (...parts: string[]) => string;
  };
}

export interface CaptionMetadataIdentity {
  projectDocumentId: string;
  projectPath: string;
  sequenceID: string;
  sequenceName: string;
}

interface PersistedCaptionClipMetadata {
  trackIndex: number;
  startSeconds: number;
  endSeconds: number;
  text: string;
  words: CaptionWord[];
}

export interface CaptionMetadataClipMatchSource {
  trackIndex: number;
  startSeconds: number;
  endSeconds: number;
  text: string;
  words: CaptionWord[];
}

interface PersistedSequenceCaptionMetadata {
  identity: CaptionMetadataIdentity;
  updatedAtIso: string;
  clips: PersistedCaptionClipMetadata[];
}

interface PersistedCaptionMetadataStore {
  version: 1;
  sequences: Record<string, PersistedSequenceCaptionMetadata>;
}

const CAPTION_METADATA_STORE_FILE_NAME = "subcreator-caption-metadata.json";

function resolveCepNodeModules(): CepNodeModules | null {
  // // Resolve CEP Node modules directly from the panel runtime so metadata can be saved outside Premiere projects.
  const nodeRequire =
    (window.cep_node && typeof window.cep_node.require === "function" ? window.cep_node.require : null) ||
    (typeof window.require === "function" ? window.require : null);
  if (!nodeRequire) {
    return null;
  }

  try {
    return {
      fs: nodeRequire("fs") as CepNodeModules["fs"],
      os: nodeRequire("os") as CepNodeModules["os"],
      path: nodeRequire("path") as CepNodeModules["path"]
    };
  } catch {
    return null;
  }
}

function normalizeMetadataIdentity(identity: CaptionMetadataIdentity | null | undefined): CaptionMetadataIdentity | null {
  // // Accept only identities with stable project and sequence IDs before touching the metadata store.
  const projectDocumentId = String(identity?.projectDocumentId || "").trim();
  const sequenceID = String(identity?.sequenceID || "").trim();
  if (!projectDocumentId || !sequenceID) {
    return null;
  }

  return {
    projectDocumentId,
    projectPath: String(identity?.projectPath || "").trim(),
    sequenceID,
    sequenceName: String(identity?.sequenceName || "").trim()
  };
}

function buildCaptionMetadataSequenceKey(identity: CaptionMetadataIdentity): string {
  // // Namespace persisted metadata by project and sequence so multiple timelines can coexist safely.
  return `${identity.projectDocumentId}::${identity.sequenceID}`;
}

function getCaptionMetadataStorePath(modules: CepNodeModules): string {
  // // Keep metadata in the user profile so it survives panel reloads without polluting project files.
  const supportFolder = modules.path.join(modules.os.homedir(), ".subcreator");
  modules.fs.mkdirSync(supportFolder, { recursive: true });
  return modules.path.join(supportFolder, CAPTION_METADATA_STORE_FILE_NAME);
}

function readCaptionMetadataStore(modules: CepNodeModules): PersistedCaptionMetadataStore {
  // // Load the store defensively and fall back to one empty schema when the file is missing or malformed.
  const storePath = getCaptionMetadataStorePath(modules);
  if (!modules.fs.existsSync(storePath)) {
    return {
      version: 1,
      sequences: {}
    };
  }

  try {
    const rawText = modules.fs.readFileSync(storePath, "utf8");
    const parsed = JSON.parse(String(rawText || "")) as PersistedCaptionMetadataStore;
    if (parsed && parsed.version === 1 && parsed.sequences && typeof parsed.sequences === "object") {
      return parsed;
    }
  } catch {
    // // Ignore malformed store content and rebuild from scratch.
  }

  return {
    version: 1,
    sequences: {}
  };
}

function writeCaptionMetadataStore(modules: CepNodeModules, store: PersistedCaptionMetadataStore): void {
  // // Persist the full metadata store atomically enough for single-panel usage.
  modules.fs.writeFileSync(getCaptionMetadataStorePath(modules), JSON.stringify(store, null, 2));
}

function normalizeMetadataText(value: string): string {
  // // Normalize caption text so matching survives trivial whitespace differences.
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function buildMetadataWordSignatureFromText(text: string): string {
  // // Reduce text to one normalized word signature so persisted timings can survive harmless whitespace and track moves.
  return normalizeMetadataText(text)
    .toLowerCase()
    .split(" ")
    .filter(Boolean)
    .join(" ");
}

function buildMetadataWordSignatureFromWords(words: CaptionWord[]): string {
  // // Derive the same normalized signature from stored word-timing payloads for fallback metadata matching.
  return words
    .map((word) => normalizeMetadataText(String(word.text || "")).toLowerCase())
    .filter(Boolean)
    .join(" ");
}

function cloneCaptionWords(words: CaptionWord[]): CaptionWord[] {
  // // Clone word timings to keep store reads/writes immutable and predictable.
  return words.map((word) => ({
    text: String(word.text || "").trim(),
    startSeconds: Number(word.startSeconds || 0),
    endSeconds: Number(word.endSeconds || 0)
  }));
}

function buildPersistedCaptionClipMetadata(trackIndex: number, cue: CaptionCue): PersistedCaptionClipMetadata {
  // // Convert one generated cue into persisted metadata on the target track.
  return {
    trackIndex,
    startSeconds: Number(cue.startSeconds || 0),
    endSeconds: Number(cue.endSeconds || 0),
    text: normalizeMetadataText(cue.text),
    words: cloneCaptionWords(Array.isArray(cue.words) ? cue.words : [])
  };
}

function buildPersistedCaptionClipFromTextBlock(trackIndex: number, block: TextEditorBlock): PersistedCaptionClipMetadata {
  // // Persist rebuilt Text editor blocks with their exact timings when available, otherwise with synthetic weighted words.
  const normalizedText = normalizeMetadataText(block.text);
  const normalizedWords =
    Array.isArray(block.timedWords) && block.timedWords.length === block.words.length
      ? cloneCaptionWords(block.timedWords)
      : buildWeightedCaptionWords(block.words, Number(block.startSeconds || 0), Number(block.endSeconds || 0));

  return {
    trackIndex,
    startSeconds: Number(block.startSeconds || 0),
    endSeconds: Number(block.endSeconds || 0),
    text: normalizedText,
    words: normalizedWords
  };
}

function rangesOverlap(leftStart: number, leftEnd: number, rightStart: number, rightEnd: number): boolean {
  // // Compare cue ranges with a tiny tolerance so store updates replace only the intended slice.
  return leftStart < rightEnd - 0.0005 && leftEnd > rightStart + 0.0005;
}

function upsertSequenceCaptionMetadata(
  store: PersistedCaptionMetadataStore,
  identity: CaptionMetadataIdentity
): PersistedSequenceCaptionMetadata {
  // // Reuse one sequence entry when present or initialize it lazily.
  const sequenceKey = buildCaptionMetadataSequenceKey(identity);
  const existing = store.sequences[sequenceKey];
  if (existing) {
    existing.identity = identity;
    return existing;
  }

  const created: PersistedSequenceCaptionMetadata = {
    identity,
    updatedAtIso: new Date().toISOString(),
    clips: []
  };
  store.sequences[sequenceKey] = created;
  return created;
}

export function persistGeneratedCaptionMetadata(
  identity: CaptionMetadataIdentity | null | undefined,
  trackIndex: number,
  cues: CaptionCue[]
): void {
  // // Save freshly generated subtitle timing metadata so later text edits can reuse exact word timing.
  const modules = resolveCepNodeModules();
  const normalizedIdentity = normalizeMetadataIdentity(identity);
  if (!modules || !normalizedIdentity || !Array.isArray(cues) || cues.length < 1 || !Number.isFinite(Number(trackIndex))) {
    return;
  }

  const store = readCaptionMetadataStore(modules);
  const sequenceEntry = upsertSequenceCaptionMetadata(store, normalizedIdentity);
  const clipMetadata = cues.map((cue) => buildPersistedCaptionClipMetadata(Number(trackIndex), cue));
  const rangeStart = clipMetadata.reduce((lowestValue, clip) => Math.min(lowestValue, clip.startSeconds), Number.POSITIVE_INFINITY);
  const rangeEnd = clipMetadata.reduce((highestValue, clip) => Math.max(highestValue, clip.endSeconds), Number.NEGATIVE_INFINITY);
  sequenceEntry.clips = sequenceEntry.clips
    .filter(
      (clip) => clip.trackIndex !== Number(trackIndex) || !rangesOverlap(clip.startSeconds, clip.endSeconds, rangeStart, rangeEnd)
    )
    .concat(clipMetadata);
  sequenceEntry.updatedAtIso = new Date().toISOString();
  store.sequences[buildCaptionMetadataSequenceKey(normalizedIdentity)] = sequenceEntry;
  writeCaptionMetadataStore(modules, store);
}

export function findBestCaptionMetadataMatchForItem(
  clips: CaptionMetadataClipMatchSource[],
  item: SelectedMogrtTextItem
): CaptionMetadataClipMatchSource | null {
  // // Match one timeline subtitle item against persisted metadata with fallback across tracks so timing survives safe track moves.
  const targetText = normalizeMetadataText(item.text);
  const targetWordSignature = buildMetadataWordSignatureFromText(item.text);
  const targetTrackIndex = Number(item.videoTrackIndex);
  const targetStart = Number(item.startSeconds || 0);
  const targetEnd = Number(item.endSeconds || 0);
  let bestMatch: CaptionMetadataClipMatchSource | null = null;
  let bestScore = Number.POSITIVE_INFINITY;

  for (const clip of clips) {
    const normalizedClipText = normalizeMetadataText(clip.text);
    const clipWordSignature = buildMetadataWordSignatureFromWords(Array.isArray(clip.words) ? clip.words : []);
    const sameTrack = clip.trackIndex === targetTrackIndex;
    const sameText = normalizedClipText === targetText;
    const sameWordSignature = clipWordSignature.length > 0 && clipWordSignature === targetWordSignature;
    if (!sameText && !sameWordSignature) {
      continue;
    }

    const startDelta = Math.abs(Number(clip.startSeconds || 0) - targetStart);
    const endDelta = Math.abs(Number(clip.endSeconds || 0) - targetEnd);
    if (startDelta > 0.6 || endDelta > 0.6) {
      continue;
    }

    let score = startDelta + endDelta;
    if (!sameTrack) {
      score += 0.2;
    }
    if (!sameText) {
      score += 0.25;
    }

    if (score < bestScore) {
      bestScore = score;
      bestMatch = clip;
    }
  }

  return bestMatch;
}

export function resolveCaptionMetadataForSelection(
  identity: CaptionMetadataIdentity | null | undefined,
  items: SelectedMogrtTextItem[]
): Array<CaptionWord[] | null> {
  // // Enrich each selected subtitle with persisted word timings when the current sequence has matching metadata.
  const modules = resolveCepNodeModules();
  const normalizedIdentity = normalizeMetadataIdentity(identity);
  if (!modules || !normalizedIdentity || !Array.isArray(items) || items.length < 1) {
    return items.map(() => null);
  }

  const store = readCaptionMetadataStore(modules);
  const sequenceEntry = store.sequences[buildCaptionMetadataSequenceKey(normalizedIdentity)];
  if (!sequenceEntry || !Array.isArray(sequenceEntry.clips)) {
    return items.map(() => null);
  }

  return items.map((item) => {
    const match = findBestCaptionMetadataMatchForItem(sequenceEntry.clips, item);
    return match && Array.isArray(match.words) && match.words.length > 0 ? cloneCaptionWords(match.words) : null;
  });
}

export function persistTextEditorCaptionMetadata(
  identity: CaptionMetadataIdentity | null | undefined,
  sourceTrackIndex: number,
  rebuildTrackIndex: number,
  replacedRange: { startSeconds: number; endSeconds: number },
  blocks: TextEditorBlock[]
): void {
  // // Update persisted metadata after text-editor rebuilds so future edits keep precise timing continuity.
  const modules = resolveCepNodeModules();
  const normalizedIdentity = normalizeMetadataIdentity(identity);
  if (
    !modules ||
    !normalizedIdentity ||
    !Array.isArray(blocks) ||
    blocks.length < 1 ||
    !Number.isFinite(Number(sourceTrackIndex)) ||
    !Number.isFinite(Number(rebuildTrackIndex))
  ) {
    return;
  }

  const rangeStart = Number(replacedRange.startSeconds || 0);
  const rangeEnd = Number(replacedRange.endSeconds || 0);
  if (!(rangeEnd > rangeStart)) {
    return;
  }

  const store = readCaptionMetadataStore(modules);
  const sequenceEntry = upsertSequenceCaptionMetadata(store, normalizedIdentity);
  sequenceEntry.clips = sequenceEntry.clips.filter((clip) => {
    if (clip.trackIndex === Number(sourceTrackIndex) && rangesOverlap(clip.startSeconds, clip.endSeconds, rangeStart, rangeEnd)) {
      return false;
    }
    if (clip.trackIndex === Number(rebuildTrackIndex) && rangesOverlap(clip.startSeconds, clip.endSeconds, rangeStart, rangeEnd)) {
      return false;
    }
    return true;
  });

  for (const block of blocks) {
    sequenceEntry.clips.push(buildPersistedCaptionClipFromTextBlock(Number(rebuildTrackIndex), block));
  }
  sequenceEntry.updatedAtIso = new Date().toISOString();
  store.sequences[buildCaptionMetadataSequenceKey(normalizedIdentity)] = sequenceEntry;
  writeCaptionMetadataStore(modules, store);
}
