// // Share subtitle-token normalization across generation and editor workflows.
import type { CaptionWord } from "./types";

const SUBCREATOR_WORD_CHAR_PATTERN = "[\\p{L}\\p{N}]";
const SUBCREATOR_APOSTROPHE_PATTERN = "['’]";
const SUBCREATOR_HYPHEN_PATTERN = "[-‑‒–—]";

export function normalizeInlineSubtitleText(text: string): string {
  // // Collapse whitespace and re-attach apostrophes / terminal punctuation so word counts stay stable.
  let normalized = String(text || "")
    .replace(/\r/g, "\n")
    .replace(/\n+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (!normalized) {
    return "";
  }

  const beforeApostrophePattern = new RegExp(
    `(${SUBCREATOR_WORD_CHAR_PATTERN}+)\\s+(${SUBCREATOR_APOSTROPHE_PATTERN})`,
    "gu"
  );
  const afterApostrophePattern = new RegExp(
    `(${SUBCREATOR_APOSTROPHE_PATTERN})\\s+(${SUBCREATOR_WORD_CHAR_PATTERN}+)`,
    "gu"
  );
  const beforeHyphenPattern = new RegExp(`(${SUBCREATOR_WORD_CHAR_PATTERN}+)\\s+(${SUBCREATOR_HYPHEN_PATTERN})`, "gu");
  const afterHyphenPattern = new RegExp(`(${SUBCREATOR_HYPHEN_PATTERN})\\s+(${SUBCREATOR_WORD_CHAR_PATTERN}+)`, "gu");

  let previous = "";
  while (normalized !== previous) {
    previous = normalized;
    normalized = normalized
      .replace(beforeApostrophePattern, "$1$2")
      .replace(afterApostrophePattern, "$1$2")
      .replace(beforeHyphenPattern, "$1$2")
      .replace(afterHyphenPattern, "$1$2")
      .replace(/(\p{N})\s+%/gu, "$1%")
      .replace(/\s+([!?]+)/g, "$1")
      .replace(/\s+/g, " ")
      .trim();
  }

  return normalized;
}

export function tokenizeSubtitleText(text: string): string[] {
  // // Tokenize subtitle text after punctuation normalization so editor/generation counts agree.
  const normalized = normalizeInlineSubtitleText(text);
  return normalized ? normalized.split(/\s+/).filter(Boolean) : [];
}

export function normalizeCaptionWords(words: CaptionWord[]): CaptionWord[] {
  // // Merge split apostrophe / hyphen / punctuation tokens while preserving the original timing span.
  const normalized: CaptionWord[] = [];

  const appendToPrevious = (suffix: string, endSeconds: number): void => {
    const previousWord = normalized[normalized.length - 1];
    if (!previousWord) {
      return;
    }
    previousWord.text = `${previousWord.text}${suffix}`;
    previousWord.endSeconds = Number(endSeconds || previousWord.endSeconds);
  };

  for (let index = 0; index < words.length; index += 1) {
    const sourceWord = words[index];
    const currentText = String(sourceWord?.text || "")
      .trim()
      .replace(/(\p{N})\s+%/gu, "$1%");
    if (!currentText) {
      continue;
    }

    if (normalized.length > 0 && /^[!?]+$/.test(currentText)) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    if (normalized.length > 0 && currentText === "%" && /\p{N}$/u.test(normalized[normalized.length - 1]?.text || "")) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    if (normalized.length > 0 && /^['’][\p{L}\p{N}]+$/u.test(currentText)) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    if (normalized.length > 0 && /^[-‑‒–—][\p{L}\p{N}]+$/u.test(currentText)) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    if (/^['’]$/u.test(currentText) && normalized.length > 0) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    if (/^[-‑‒–—]$/u.test(currentText) && normalized.length > 0) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    if (normalized.length > 0 && /['’]$/u.test(normalized[normalized.length - 1]?.text || "") && /^[\p{L}\p{N}]+$/u.test(currentText)) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    if (normalized.length > 0 && /[-‑‒–—]$/u.test(normalized[normalized.length - 1]?.text || "") && /^[\p{L}\p{N}]+$/u.test(currentText)) {
      appendToPrevious(currentText, Number(sourceWord.endSeconds || sourceWord.startSeconds || 0));
      continue;
    }

    normalized.push({
      text: currentText,
      startSeconds: Number(sourceWord?.startSeconds || 0),
      endSeconds: Number(sourceWord?.endSeconds || 0)
    });
  }

  return normalized;
}
