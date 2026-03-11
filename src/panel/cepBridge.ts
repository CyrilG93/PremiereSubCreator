// // Wrap CEP evalScript calls and provide a browser fallback for local testing.
import type { CaptionSourceItem, HostApplyPayload } from "../core/types";

declare global {
  interface Window {
    __adobe_cep__?: {
      evalScript: (script: string, callback: (result: string) => void) => void;
    };
  }
}

interface HostJsonResponse<T> {
  ok: boolean;
  error?: string;
  data?: T;
  [key: string]: unknown;
}

interface WhisperTranscriptionRequest {
  audioPath: string;
  languageCode: string;
  model: string;
}

interface WhisperTranscriptionResult {
  srtText: string;
  model: string;
  audioPath: string;
  commandOutput?: string;
}

function escapeForJsx(input: string): string {
  // // Escape special characters before embedding text into evalScript call strings.
  return input
    .replace(/\\/g, "\\\\")
    .replace(/"/g, "\\\"")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "\\r");
}

function evalScript(script: string): Promise<string> {
  // // Route script execution through Premiere CEP host when available.
  if (window.__adobe_cep__) {
    return new Promise((resolve) => {
      window.__adobe_cep__?.evalScript(script, (result) => resolve(result));
    });
  }

  return Promise.resolve(
    JSON.stringify({
      ok: true,
      mocked: true,
      message: "CEP host unavailable, running in browser fallback mode."
    })
  );
}

async function evalHostJson<T>(script: string): Promise<HostJsonResponse<T>> {
  // // Parse JSON returned by host-side ExtendScript function calls.
  const raw = await evalScript(script);

  try {
    return JSON.parse(raw) as HostJsonResponse<T>;
  } catch (error) {
    return {
      ok: false,
      error: `Invalid host response: ${String(error)} | raw=${raw}`
    };
  }
}

export async function pingHost(): Promise<string> {
  // // Validate the bridge wiring with a lightweight host call.
  return evalScript("subcreator_ping()");
}

export async function applyCaptionPlan(payload: HostApplyPayload): Promise<string> {
  // // Send JSON payload as URI-encoded text to avoid quote escaping edge-cases.
  const encodedPayload = encodeURIComponent(JSON.stringify(payload));
  return evalScript(`subcreator_apply_captions("${escapeForJsx(encodedPayload)}")`);
}

export async function listCaptionSources(): Promise<CaptionSourceItem[]> {
  // // Ask Premiere host to list subtitle files already imported in project bins.
  const response = await evalHostJson<CaptionSourceItem[]>("subcreator_list_caption_sources()");
  if (!response.ok) {
    throw new Error(response.error ?? "Unable to list caption sources from host.");
  }

  return Array.isArray(response.data) ? response.data : [];
}

export async function readTextFileFromHost(filePath: string): Promise<string> {
  // // Read subtitle files through host to avoid browser file access limitations.
  const encoded = encodeURIComponent(filePath);
  const response = await evalHostJson<{ text: string }>(`subcreator_read_text_file("${escapeForJsx(encoded)}")`);

  if (!response.ok) {
    throw new Error(response.error ?? "Unable to read file from host.");
  }

  return String(response.data?.text ?? "");
}

export async function pickWhisperAudioPath(): Promise<string> {
  // // Open native file picker from host for Whisper transcription input.
  const response = await evalHostJson<{ path: string }>("subcreator_pick_audio_file()");
  if (!response.ok) {
    throw new Error(response.error ?? "Audio picker failed.");
  }

  return String(response.data?.path ?? "");
}

export async function transcribeWithWhisper(request: WhisperTranscriptionRequest): Promise<WhisperTranscriptionResult> {
  // // Trigger local Whisper CLI through ExtendScript host and return generated SRT.
  const encodedPayload = encodeURIComponent(JSON.stringify(request));
  const response = await evalHostJson<WhisperTranscriptionResult>(
    `subcreator_transcribe_whisper("${escapeForJsx(encodedPayload)}")`
  );

  if (!response.ok) {
    throw new Error(response.error ?? "Whisper transcription failed.");
  }

  return {
    srtText: String(response.data?.srtText ?? ""),
    model: String(response.data?.model ?? request.model),
    audioPath: String(response.data?.audioPath ?? request.audioPath),
    commandOutput: String(response.data?.commandOutput ?? "")
  };
}
