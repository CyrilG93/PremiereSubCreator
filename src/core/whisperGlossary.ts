// // Parse and apply a persistent Whisper vocabulary without losing subtitle timing spans.
import type { CaptionCue, CaptionWord } from "./types";
import { normalizeCaptionWords, normalizeInlineSubtitleText, tokenizeSubtitleText } from "./textNormalization";
import { buildWeightedCaptionWords, buildWeightedCaptionWordsFromText } from "./wordTiming";

export interface WhisperGlossaryEntry {
  canonical: string;
  variants: string[];
}

interface WhisperGlossaryMatcher {
  canonical: string;
  canonicalWords: string[];
  sourceWords: string[];
}

const GLOSSARY_ALIAS_SEPARATOR = "=>";
const GLOSSARY_VARIANT_SEPARATOR = "|";
const GLOSSARY_PROMPT_MAX_LENGTH = 420;

function normalizeGlossaryValue(value: string): string {
  // // Keep one compact human-readable term while preserving its requested capitalization.
  return normalizeInlineSubtitleText(String(value || ""));
}

function normalizeGlossaryMatchWord(value: string): string {
  // // Ignore sentence punctuation and letter case while preserving meaningful symbols such as C++ or C#.
  return String(value || "")
    .trim()
    .replace(/^[([{"'“”‘’]+/u, "")
    .replace(/[)\]}"'“”‘’,.!?;:…]+$/u, "")
    .toLocaleLowerCase();
}

function pushUniqueValue(values: string[], candidate: string): void {
  // // De-duplicate glossary values without changing the spelling of the first occurrence.
  const normalizedCandidate = normalizeGlossaryValue(candidate);
  if (!normalizedCandidate) {
    return;
  }

  const lookup = normalizedCandidate.toLocaleLowerCase();
  if (!values.some((value) => value.toLocaleLowerCase() === lookup)) {
    values.push(normalizedCandidate);
  }
}

function parseStandaloneGlossaryTerms(line: string): string[] {
  // // Accept the former comma-separated field while recommending one canonical term per line.
  return line
    .split(/[,;]/)
    .map((value) => normalizeGlossaryValue(value))
    .filter(Boolean);
}

export function parseWhisperGlossary(input: string): WhisperGlossaryEntry[] {
  // // Parse canonical terms and optional `variant | variant => Canonical spelling` aliases.
  const entriesByCanonical = new Map<string, WhisperGlossaryEntry>();
  const lines = String(input || "")
    .replace(/\r/g, "\n")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter((line) => Boolean(line) && !line.startsWith("#"));

  for (const line of lines) {
    const separatorIndex = line.indexOf(GLOSSARY_ALIAS_SEPARATOR);
    const canonicalCandidates =
      separatorIndex >= 0 ? [normalizeGlossaryValue(line.slice(separatorIndex + GLOSSARY_ALIAS_SEPARATOR.length))] : parseStandaloneGlossaryTerms(line);

    for (const canonical of canonicalCandidates) {
      if (!canonical) {
        continue;
      }

      const lookup = canonical.toLocaleLowerCase();
      const entry = entriesByCanonical.get(lookup) || { canonical, variants: [] };
      pushUniqueValue(entry.variants, canonical);

      if (separatorIndex >= 0) {
        line
          .slice(0, separatorIndex)
          .split(GLOSSARY_VARIANT_SEPARATOR)
          .forEach((variant) => pushUniqueValue(entry.variants, variant));
      }

      entriesByCanonical.set(lookup, entry);
    }
  }

  return Array.from(entriesByCanonical.values());
}

export function buildWhisperGlossaryPrompt(input: string, maxLength = GLOSSARY_PROMPT_MAX_LENGTH): string {
  // // Send only complete canonical spellings because Whisper treats its prompt as compact vocabulary guidance.
  const availableLength = Math.max(0, Math.floor(Number(maxLength) || 0));
  if (availableLength < 1) {
    return "";
  }

  const selectedTerms: string[] = [];
  for (const entry of parseWhisperGlossary(input)) {
    const candidate = [...selectedTerms, entry.canonical].join(", ");
    if (candidate.length > availableLength) {
      break;
    }
    selectedTerms.push(entry.canonical);
  }

  return selectedTerms.join(", ");
}

function buildGlossaryMatchers(entries: WhisperGlossaryEntry[]): WhisperGlossaryMatcher[] {
  // // Prefer longer aliases first so a brand phrase wins over one of its individual words.
  const matchers: WhisperGlossaryMatcher[] = [];

  for (const entry of entries) {
    const canonicalWords = tokenizeSubtitleText(entry.canonical);
    if (canonicalWords.length < 1) {
      continue;
    }

    for (const variant of entry.variants) {
      const sourceWords = tokenizeSubtitleText(variant).map(normalizeGlossaryMatchWord).filter(Boolean);
      if (sourceWords.length < 1) {
        continue;
      }
      matchers.push({ canonical: entry.canonical, canonicalWords, sourceWords });
    }
  }

  return matchers.sort((left, right) => {
    // // Resolve phrases before short aliases, then use source length as a stable secondary priority.
    const wordCountDifference = right.sourceWords.length - left.sourceWords.length;
    return wordCountDifference || right.sourceWords.join(" ").length - left.sourceWords.join(" ").length;
  });
}

function readLeadingWrapperPunctuation(value: string): string {
  // // Preserve quotes or brackets surrounding a corrected phrase.
  return String(value || "").match(/^[([{"'“”‘’]+/u)?.[0] || "";
}

function readTrailingWrapperPunctuation(value: string): string {
  // // Preserve terminal punctuation surrounding a corrected phrase.
  return String(value || "").match(/[)\]}"'“”‘’,.!?;:…]+$/u)?.[0] || "";
}

function matcherMatchesAt(words: CaptionWord[], startIndex: number, matcher: WhisperGlossaryMatcher): boolean {
  // // Compare one alias against consecutive timed words at the current cue position.
  if (startIndex + matcher.sourceWords.length > words.length) {
    return false;
  }

  return matcher.sourceWords.every((sourceWord, offset) => {
    return normalizeGlossaryMatchWord(words[startIndex + offset]?.text || "") === sourceWord;
  });
}

function buildReplacementWords(sourceWords: CaptionWord[], matcher: WhisperGlossaryMatcher): CaptionWord[] {
  // // Re-distribute only the matched span so all unrelated Whisper word timestamps stay untouched.
  const firstSourceWord = sourceWords[0];
  const lastSourceWord = sourceWords[sourceWords.length - 1];
  const replacementWords = buildWeightedCaptionWords(
    matcher.canonicalWords,
    firstSourceWord.startSeconds,
    lastSourceWord.endSeconds
  );
  if (replacementWords.length < 1) {
    return sourceWords.map((word) => ({ ...word }));
  }

  const leadingPunctuation = readLeadingWrapperPunctuation(firstSourceWord.text);
  const trailingPunctuation = readTrailingWrapperPunctuation(lastSourceWord.text);
  if (leadingPunctuation && !replacementWords[0].text.startsWith(leadingPunctuation)) {
    replacementWords[0].text = `${leadingPunctuation}${replacementWords[0].text}`;
  }
  if (trailingPunctuation && !replacementWords[replacementWords.length - 1].text.endsWith(trailingPunctuation)) {
    replacementWords[replacementWords.length - 1].text = `${replacementWords[replacementWords.length - 1].text}${trailingPunctuation}`;
  }

  return replacementWords;
}

function applyMatchersToWords(sourceWords: CaptionWord[], matchers: WhisperGlossaryMatcher[]): CaptionWord[] {
  // // Replace aliases left-to-right while preventing a corrected span from being rewritten twice.
  const words = normalizeCaptionWords(sourceWords);
  const correctedWords: CaptionWord[] = [];

  for (let index = 0; index < words.length; ) {
    const matcher = matchers.find((candidate) => matcherMatchesAt(words, index, candidate));
    if (!matcher) {
      correctedWords.push({ ...words[index] });
      index += 1;
      continue;
    }

    const matchedWords = words.slice(index, index + matcher.sourceWords.length);
    correctedWords.push(...buildReplacementWords(matchedWords, matcher));
    index += matcher.sourceWords.length;
  }

  return correctedWords;
}

export function applyWhisperGlossaryToCues(cues: CaptionCue[], input: string): CaptionCue[] {
  // // Correct Whisper and WhisperX cues while preserving cue boundaries and matched timing spans.
  const matchers = buildGlossaryMatchers(parseWhisperGlossary(input));
  if (matchers.length < 1) {
    return cues.map((cue) => ({ ...cue, words: cue.words.map((word) => ({ ...word })) }));
  }

  return cues.map((cue) => {
    const sourceWords =
      cue.words.length > 0 ? cue.words : buildWeightedCaptionWordsFromText(cue.text, cue.startSeconds, cue.endSeconds);
    const words = applyMatchersToWords(sourceWords, matchers);
    return {
      ...cue,
      text: normalizeInlineSubtitleText(words.map((word) => word.text).join(" ")),
      words
    };
  });
}
