#!/usr/bin/env python3
# // Align a corrected transcript or SRT against audio with WhisperX and emit Whisper-compatible JSON.

from __future__ import annotations

import argparse
import importlib
import json
import math
import os
import re
import sys
from typing import Any, Dict, List, Optional

SAMPLE_RATE = 16000.0
PROGRESS_PREFIX = "SUBCREATOR_ALIGN_PROGRESS"


def subcreator_log_progress(percent: int, message: str) -> None:
    # // Stream progress markers to stderr so the CEP panel can update its progress bar while alignment runs.
    bounded_percent = max(0, min(100, int(percent)))
    sys.stderr.write(f"{PROGRESS_PREFIX}\t{bounded_percent}\t{message}\n")
    sys.stderr.flush()


def subcreator_fail(message: str) -> int:
    # // Print one explicit error line for the bridge and return a failing exit code.
    sys.stderr.write(f"ERROR: {message}\n")
    sys.stderr.flush()
    return 1


def subcreator_normalize_text(value: Any) -> str:
    # // Collapse repeated whitespace so alignment payloads stay stable and comparable.
    return re.sub(r"\s+", " ", str(value or "")).strip()


def subcreator_to_float(value: Any) -> Optional[float]:
    # // Convert unknown numeric payloads into safe finite floats.
    try:
        parsed = float(value)
    except Exception:
        return None
    if not math.isfinite(parsed):
        return None
    return parsed


def subcreator_parse_srt_timestamp(value: str) -> float:
    # // Convert HH:MM:SS,mmm timestamps into float seconds.
    cleaned = str(value or "").strip().replace(",", ".")
    parts = cleaned.split(":")
    if len(parts) != 3:
        raise ValueError(f"Invalid SRT timestamp: {value}")
    hours = float(parts[0])
    minutes = float(parts[1])
    seconds = float(parts[2])
    return hours * 3600.0 + minutes * 60.0 + seconds


def subcreator_parse_srt_segments(text: str) -> List[Dict[str, Any]]:
    # // Keep corrected SRT cue boundaries as the initial seed before WhisperX refines the timing.
    blocks = re.split(r"\n\s*\n", text.replace("\r\n", "\n").strip())
    segments: List[Dict[str, Any]] = []
    for index, block in enumerate(blocks):
        lines = [line.strip() for line in block.split("\n") if line.strip()]
        if len(lines) < 2:
            continue
        time_line = lines[1] if re.fullmatch(r"\d+", lines[0] or "") else lines[0]
        text_lines = lines[2:] if re.fullmatch(r"\d+", lines[0] or "") else lines[1:]
        if "-->" not in time_line:
            continue
        start_raw, end_raw = [part.strip() for part in time_line.split("-->", 1)]
        start_seconds = subcreator_parse_srt_timestamp(start_raw)
        end_seconds = subcreator_parse_srt_timestamp(end_raw)
        segment_text = subcreator_normalize_text(" ".join(text_lines))
        if not segment_text or end_seconds <= start_seconds:
            continue
        segments.append({
            "id": index,
            "start": start_seconds,
            "end": end_seconds,
            "text": segment_text,
        })
    return segments


def subcreator_split_plaintext_segments(text: str) -> List[str]:
    # // Prefer user line breaks first, then sentence-like punctuation, so corrected TXT remains reasonably segmented.
    normalized_text = text.replace("\r\n", "\n")
    blocks: List[str] = []
    for paragraph in re.split(r"\n\s*\n+", normalized_text):
        stripped_paragraph = paragraph.strip()
        if not stripped_paragraph:
            continue
        raw_lines = [subcreator_normalize_text(line) for line in stripped_paragraph.split("\n") if subcreator_normalize_text(line)]
        if len(raw_lines) > 1:
            blocks.extend(raw_lines)
            continue
        sentence_chunks = [
            subcreator_normalize_text(chunk)
            for chunk in re.split(r"(?<=[.!?])\s+", stripped_paragraph)
            if subcreator_normalize_text(chunk)
        ]
        if len(sentence_chunks) > 1:
            blocks.extend(sentence_chunks)
        else:
            blocks.append(subcreator_normalize_text(stripped_paragraph))
    return [block for block in blocks if block]


def subcreator_compute_text_weight(text: str) -> float:
    # // Give more weight to denser lines so rough TXT pre-segmentation is less uniform than plain word-count splitting.
    alnum_count = len(re.sub(r"[^0-9A-Za-zÀ-ÿ]+", "", text))
    punctuation_bonus = len(re.findall(r"[,;:.!?]", text)) * 0.35
    return max(float(alnum_count) + punctuation_bonus, 1.0)


def subcreator_build_plaintext_segments(text: str, total_duration: float) -> List[Dict[str, Any]]:
    # // Build approximate segment spans for corrected TXT so WhisperX can refine them against the audio.
    blocks = subcreator_split_plaintext_segments(text)
    if not blocks:
        return []
    if total_duration <= 0:
        total_duration = float(len(blocks))

    total_weight = sum(subcreator_compute_text_weight(block) for block in blocks)
    cursor = 0.0
    segments: List[Dict[str, Any]] = []
    for index, block in enumerate(blocks):
        if index == len(blocks) - 1:
            end_seconds = total_duration
        else:
            share = subcreator_compute_text_weight(block) / total_weight if total_weight > 0 else 1.0 / len(blocks)
            duration = max(total_duration * share, 0.12)
            end_seconds = min(total_duration, cursor + duration)
        if end_seconds <= cursor:
            end_seconds = min(total_duration, cursor + 0.12)
        segments.append({
            "id": index,
            "start": cursor,
            "end": end_seconds,
            "text": block,
        })
        cursor = end_seconds
    if segments:
        segments[-1]["end"] = max(float(segments[-1].get("end", total_duration) or total_duration), total_duration)
    return segments


class SubcreatorSentenceSplitter:
    # // Replace WhisperX's NLTK punkt dependency with a lightweight local splitter to avoid SSL/download issues.
    def span_tokenize(self, text: str) -> List[tuple[int, int]]:
        spans: List[tuple[int, int]] = []
        for match in re.finditer(r"[^.!?]+[.!?]*\s*", text or ""):
            start = match.start()
            end = match.end()
            while start < end and text[start].isspace():
                start += 1
            while end > start and text[end - 1].isspace():
                end -= 1
            if start < end:
                spans.append((start, end))
        if not spans and text:
            spans.append((0, len(text)))
        return spans


def subcreator_patch_whisperx_sentence_tokenizer(whisperx_module: Any) -> None:
    # // Monkeypatch WhisperX sentence splitting so corrected align never depends on downloading NLTK punkt_tab.
    try:
        alignment_module = importlib.import_module("whisperx.alignment")
    except Exception:
        return

    def subcreator_nltk_load(_resource_name: str) -> SubcreatorSentenceSplitter:
        # // Return the local splitter for every requested punkt resource.
        return SubcreatorSentenceSplitter()

    try:
        alignment_module.nltk_load = subcreator_nltk_load
    except Exception:
        pass

    try:
        alignment_module.nltk.download = lambda *args, **kwargs: True
    except Exception:
        pass


def subcreator_load_corrected_segments(transcript_path: str, total_duration: float) -> List[Dict[str, Any]]:
    # // Parse corrected transcript input and return one seed segment list for WhisperX alignment.
    with open(transcript_path, "r", encoding="utf-8-sig") as handle:
        transcript_text = handle.read()
    if not transcript_text.strip():
        raise ValueError("Corrected transcript file is empty.")

    extension = os.path.splitext(transcript_path)[1].lower()
    if extension == ".srt":
        segments = subcreator_parse_srt_segments(transcript_text)
        if not segments:
            raise ValueError("Corrected SRT contains no valid subtitle cues.")
        return segments

    if extension == ".txt":
        segments = subcreator_build_plaintext_segments(transcript_text, total_duration)
        if not segments:
            raise ValueError("Corrected transcript contains no usable text.")
        return segments

    raise ValueError("Corrected transcript align supports .srt or .txt files only.")


def subcreator_clean_word_payload(word: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    # // Normalize word payloads so panel-side JSON parsing never sees tensor or numpy values.
    text = subcreator_normalize_text(word.get("word") or word.get("text") or "")
    if not text:
        return None

    cleaned_word: Dict[str, Any] = {"word": text}
    start_value = subcreator_to_float(word.get("start"))
    end_value = subcreator_to_float(word.get("end"))
    if start_value is not None:
        cleaned_word["start"] = start_value
    if end_value is not None:
        cleaned_word["end"] = end_value
    return cleaned_word


def subcreator_clean_segment_payload(index: int, segment: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    # // Convert WhisperX segment payloads into plain JSON objects compatible with the existing Whisper parser.
    words = []
    for raw_word in segment.get("words") or []:
        if not isinstance(raw_word, dict):
            continue
        cleaned_word = subcreator_clean_word_payload(raw_word)
        if cleaned_word:
            words.append(cleaned_word)

    text = subcreator_normalize_text(segment.get("text") or " ".join(word["word"] for word in words))
    if not text:
        return None

    start_value = subcreator_to_float(segment.get("start"))
    end_value = subcreator_to_float(segment.get("end"))
    if start_value is None and words:
        start_value = subcreator_to_float(words[0].get("start"))
    if end_value is None and words:
        end_value = subcreator_to_float(words[-1].get("end"))
    if start_value is None or end_value is None or end_value < start_value:
        return None

    cleaned_segment: Dict[str, Any] = {
        "id": int(segment.get("id", index) if isinstance(segment.get("id"), int) else index),
        "start": start_value,
        "end": end_value,
        "text": text,
        "words": words,
    }
    return cleaned_segment


def subcreator_detect_device() -> str:
    # // Prefer CUDA when torch reports it, otherwise stay on CPU for broad workstation compatibility.
    try:
        import torch  # type: ignore

        if hasattr(torch, "cuda") and callable(getattr(torch.cuda, "is_available", None)) and torch.cuda.is_available():
            return "cuda"
    except Exception:
        return "cpu"
    return "cpu"


def main() -> int:
    # // Run one corrected-alignment pass and emit Whisper-compatible JSON to the requested output file.
    parser = argparse.ArgumentParser(description="Align corrected transcript text with WhisperX.")
    parser.add_argument("--audio", required=True, help="Path to temporary WAV audio file")
    parser.add_argument("--transcript", required=True, help="Path to corrected .srt or .txt transcript file")
    parser.add_argument("--language", required=True, help="Language code used to load the alignment model")
    parser.add_argument("--output", required=True, help="Path to the output JSON file")
    args = parser.parse_args()

    audio_path = os.path.abspath(args.audio)
    transcript_path = os.path.abspath(args.transcript)
    output_path = os.path.abspath(args.output)
    language_code = subcreator_normalize_text(args.language or "").lower()

    if not os.path.isfile(audio_path):
        return subcreator_fail(f"Audio file not found: {audio_path}")
    if not os.path.isfile(transcript_path):
        return subcreator_fail(f"Corrected transcript file not found: {transcript_path}")
    if not language_code or language_code == "auto":
        return subcreator_fail("Corrected transcript align requires an explicit language code.")

    try:
        subcreator_log_progress(8, "Loading WhisperX dependencies")
        import whisperx  # type: ignore
        subcreator_patch_whisperx_sentence_tokenizer(whisperx)
    except Exception as error:
        return subcreator_fail(f"Unable to import whisperx. Install it with `python -m pip install --user --upgrade whisperx`. Detail: {error}")

    try:
        subcreator_log_progress(18, "Loading sequence audio")
        audio = whisperx.load_audio(audio_path)
        total_duration = max(float(len(audio)) / SAMPLE_RATE, 0.0)

        subcreator_log_progress(28, "Loading corrected transcript")
        transcript_segments = subcreator_load_corrected_segments(transcript_path, total_duration)

        device = subcreator_detect_device()
        subcreator_log_progress(42, f"Loading {language_code} alignment model on {device}")
        model_align, metadata = whisperx.load_align_model(language_code=language_code, device=device)

        subcreator_log_progress(62, "Aligning corrected transcript")
        aligned_result = whisperx.align(
            transcript_segments,
            model_align,
            metadata,
            audio,
            device,
            return_char_alignments=False,
        )

        raw_segments = aligned_result.get("segments") if isinstance(aligned_result, dict) else []
        cleaned_segments = []
        for index, segment in enumerate(raw_segments or []):
            if not isinstance(segment, dict):
                continue
            cleaned_segment = subcreator_clean_segment_payload(index, segment)
            if cleaned_segment:
                cleaned_segments.append(cleaned_segment)

        if not cleaned_segments:
            return subcreator_fail("WhisperX returned no usable aligned segments for the corrected transcript.")

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        payload = {"segments": cleaned_segments, "language": language_code}
        subcreator_log_progress(92, "Writing aligned timing data")
        with open(output_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False)

        subcreator_log_progress(100, "Corrected transcript alignment complete")
        return 0
    except Exception as error:
        return subcreator_fail(str(error))


if __name__ == "__main__":
    raise SystemExit(main())
