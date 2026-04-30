#!/usr/bin/env python3
# // Align a corrected transcript or SRT against audio with WhisperX and emit Whisper-compatible JSON.

from __future__ import annotations

import argparse
import importlib
import json
import math
import os
import re
import ssl
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


def subcreator_configure_certifi_ssl() -> bool:
    # // Point urllib/torch downloads at certifi so macOS Python installs with missing local CAs can still fetch alignment models.
    try:
        import certifi  # type: ignore

        cafile = certifi.where()
    except Exception:
        return False

    if not cafile or not os.path.isfile(cafile):
        return False

    os.environ.setdefault("SSL_CERT_FILE", cafile)
    os.environ.setdefault("REQUESTS_CA_BUNDLE", cafile)

    def subcreator_create_certifi_https_context(*args: Any, **kwargs: Any) -> ssl.SSLContext:
        # // Preserve explicit caller trust settings while defaulting missing CA paths to certifi.
        if not kwargs.get("cafile") and not kwargs.get("capath") and not kwargs.get("cadata"):
            kwargs["cafile"] = cafile
        return ssl.create_default_context(*args, **kwargs)

    try:
        ssl._create_default_https_context = subcreator_create_certifi_https_context  # type: ignore[attr-defined]
    except Exception:
        return False
    return True


def subcreator_is_ssl_certificate_error(error: Exception) -> bool:
    # // Recognize the certificate failure emitted by urllib when torch downloads WhisperX alignment models.
    normalized = str(error or "").lower()
    return (
        "certificate_verify_failed" in normalized
        or "sslcertverificationerror" in normalized
        or "unable to get local issuer certificate" in normalized
    )


def subcreator_format_runtime_error(error: Exception, certifi_ssl_enabled: bool) -> str:
    # // Add an actionable SSL hint without hiding the original WhisperX or torch error message.
    message = str(error)
    if subcreator_is_ssl_certificate_error(error):
        if certifi_ssl_enabled:
            return (
                f"{message}. SSL certificate validation still failed after using certifi. "
                "Check the network proxy/corporate certificate settings for this Python runtime."
            )
        return f"{message}. Install or update certifi with `python -m pip install --user --upgrade certifi`, then retry."
    return message


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


def subcreator_trim_srt_segments_to_range(
    segments: List[Dict[str, Any]],
    range_start_seconds: Optional[float],
    range_end_seconds: Optional[float]
) -> List[Dict[str, Any]]:
    # // Keep only corrected SRT cues that intersect the exported In/Out range, then rebase them to the audio excerpt.
    if range_start_seconds is None or range_end_seconds is None or range_end_seconds <= range_start_seconds:
        return segments

    trimmed_segments: List[Dict[str, Any]] = []
    for index, segment in enumerate(segments):
        start_seconds = subcreator_to_float(segment.get("start"))
        end_seconds = subcreator_to_float(segment.get("end"))
        if start_seconds is None or end_seconds is None:
            continue
        if end_seconds <= range_start_seconds or start_seconds >= range_end_seconds:
            continue

        trimmed_segments.append({
            "id": index,
            "start": max(start_seconds, range_start_seconds) - range_start_seconds,
            "end": min(end_seconds, range_end_seconds) - range_start_seconds,
            "text": subcreator_normalize_text(segment.get("text") or ""),
        })
    return trimmed_segments


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


def subcreator_load_corrected_segments(
    transcript_path: str,
    total_duration: float,
    range_start_seconds: Optional[float] = None,
    range_end_seconds: Optional[float] = None
) -> List[Dict[str, Any]]:
    # // Parse corrected transcript input and return one seed segment list for WhisperX alignment.
    with open(transcript_path, "r", encoding="utf-8-sig") as handle:
        transcript_text = handle.read()
    if not transcript_text.strip():
        raise ValueError("Corrected transcript file is empty.")

    extension = os.path.splitext(transcript_path)[1].lower()
    if extension == ".srt":
        segments = subcreator_parse_srt_segments(transcript_text)
        segments = subcreator_trim_srt_segments_to_range(segments, range_start_seconds, range_end_seconds)
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


def subcreator_transcribe_audio_with_local_whisper(
    audio_path: str,
    model_name: str,
    device: str,
    language_code: str,
) -> Dict[str, Any]:
    # // Use the locally cached openai-whisper model as the transcript seed so WhisperX mode does not download faster-whisper models.
    try:
        import whisper  # type: ignore
    except Exception as error:
        raise ValueError(f"Unable to import openai-whisper for WhisperX transcription seed. Detail: {error}") from error

    normalized_language = subcreator_normalize_text(language_code).lower()
    language_arg = None if not normalized_language or normalized_language == "auto" else normalized_language
    transcribe_kwargs: Dict[str, Any] = {
        "fp16": str(device or "").lower() == "cuda",
        "word_timestamps": False,
        "verbose": False,
    }
    if language_arg:
        transcribe_kwargs["language"] = language_arg

    model = whisper.load_model(model_name, device=device)
    result = model.transcribe(audio_path, **transcribe_kwargs)
    if not isinstance(result, dict):
        raise ValueError("openai-whisper returned an invalid payload.")
    return result


def subcreator_align_segments_with_whisperx(
    whisperx_module: Any,
    segments: List[Dict[str, Any]],
    audio: Any,
    language_code: str,
    device: str,
) -> List[Dict[str, Any]]:
    # // Run forced alignment and return cleaned Whisper-compatible segments for the panel parser.
    model_align, metadata = whisperx_module.load_align_model(language_code=language_code, device=device)
    aligned_result = whisperx_module.align(
        segments,
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
    return cleaned_segments


def main() -> int:
    # // Run one corrected-alignment pass and emit Whisper-compatible JSON to the requested output file.
    parser = argparse.ArgumentParser(description="Align corrected transcript text with WhisperX.")
    parser.add_argument("--audio", required=True, help="Path to temporary WAV audio file")
    parser.add_argument("--transcript", required=False, default="", help="Path to corrected .srt or .txt transcript file")
    parser.add_argument("--language", required=True, help="Language code used to load the alignment model")
    parser.add_argument("--output", required=True, help="Path to the output JSON file")
    parser.add_argument("--transcribe-model", default="", help="Optional WhisperX model name used to transcribe before alignment")
    parser.add_argument("--range-start-seconds", type=float, default=None, help="Optional sequence in-point for corrected SRT rebasing")
    parser.add_argument("--range-end-seconds", type=float, default=None, help="Optional sequence out-point for corrected SRT rebasing")
    args = parser.parse_args()

    audio_path = os.path.abspath(args.audio)
    transcript_path = os.path.abspath(args.transcript) if args.transcript else ""
    output_path = os.path.abspath(args.output)
    language_code = subcreator_normalize_text(args.language or "").lower()
    transcribe_model = subcreator_normalize_text(args.transcribe_model or "")

    if not os.path.isfile(audio_path):
        return subcreator_fail(f"Audio file not found: {audio_path}")
    if not transcribe_model and not transcript_path:
        return subcreator_fail("WhisperX helper requires either --transcribe-model or --transcript.")
    if transcript_path and not os.path.isfile(transcript_path):
        return subcreator_fail(f"Corrected transcript file not found: {transcript_path}")
    if not transcribe_model and (not language_code or language_code == "auto"):
        return subcreator_fail("Corrected transcript align requires an explicit language code.")

    certifi_ssl_enabled = subcreator_configure_certifi_ssl()

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

        device = subcreator_detect_device()
        if transcribe_model:
            subcreator_log_progress(30, f"Transcribing with local Whisper {transcribe_model} model on {device}")
            transcribed_result = subcreator_transcribe_audio_with_local_whisper(
                audio_path,
                transcribe_model,
                device,
                language_code,
            )
            transcript_segments = transcribed_result.get("segments") if isinstance(transcribed_result, dict) else []
            if not isinstance(transcript_segments, list) or not transcript_segments:
                return subcreator_fail("Local Whisper transcription returned no usable segments.")

            language_code = subcreator_normalize_text(transcribed_result.get("language") or language_code).lower()
            if not language_code or language_code == "auto":
                return subcreator_fail("WhisperX transcription did not detect a usable language for alignment.")

            subcreator_log_progress(72, f"Loading {language_code} alignment model on {device}")
            cleaned_segments = subcreator_align_segments_with_whisperx(
                whisperx,
                [segment for segment in transcript_segments if isinstance(segment, dict)],
                audio,
                language_code,
                device,
            )
        else:
            subcreator_log_progress(28, "Loading corrected transcript")
            transcript_segments = subcreator_load_corrected_segments(
                transcript_path,
                total_duration,
                args.range_start_seconds,
                args.range_end_seconds
            )

            subcreator_log_progress(42, f"Loading {language_code} alignment model on {device}")
            subcreator_log_progress(62, "Aligning corrected transcript")
            cleaned_segments = subcreator_align_segments_with_whisperx(
                whisperx,
                transcript_segments,
                audio,
                language_code,
                device,
            )

        if not cleaned_segments:
            return subcreator_fail("WhisperX returned no usable aligned segments.")

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
        return subcreator_fail(subcreator_format_runtime_error(error, certifi_ssl_enabled))


if __name__ == "__main__":
    raise SystemExit(main())
