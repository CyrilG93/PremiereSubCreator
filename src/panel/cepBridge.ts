// // Wrap CEP evalScript calls and provide a browser fallback for local testing.
import type { HostApplyPayload } from "../core/types";

declare global {
  interface Window {
    __adobe_cep__?: {
      evalScript: (script: string, callback: (result: string) => void) => void;
    };
  }
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

export async function pingHost(): Promise<string> {
  // // Validate the bridge wiring with a lightweight host call.
  return evalScript("subcreator_ping()");
}

export async function applyCaptionPlan(payload: HostApplyPayload): Promise<string> {
  // // Send JSON payload as URI-encoded text to avoid quote escaping edge-cases.
  const encodedPayload = encodeURIComponent(JSON.stringify(payload));
  return evalScript(`subcreator_apply_captions("${escapeForJsx(encodedPayload)}")`);
}
