import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const hostSourcePath = fileURLToPath(new URL("../src/host/SubCreatorHost.jsx", import.meta.url));
const hostSource = readFileSync(hostSourcePath, "utf8");

function readDurationHelperSource(): string {
  // // Isolate the razor helper so the regression check ignores unrelated timeline edits elsewhere in the host.
  const match = hostSource.match(
    /function subcreator_try_razor_mogrt_duration[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_push_unique_path/
  );

  expect(match).not.toBeNull();
  return match?.[0] ?? "";
}

describe("MOGRT duration trimming", () => {
  it("uses a razor cut instead of shortening the clip or source out point", () => {
    // // A real timeline cut preserves the remaining source duration without activating Time Remapping.
    const helperSource = readDurationHelperSource();

    expect(helperSource).toContain("qeTrack.razor(razorTimecode)");
    expect(helperSource).toContain("subcreator_remove_track_item_without_ripple(rightFragment)");
    expect(helperSource).not.toMatch(/trackItem\.end\s*=/);
    expect(helperSource).not.toMatch(/trackItem\.outPoint\s*=/);
  });

  it("cuts only after MOGRT controls are initialized", () => {
    // // Text and layout writes must finish before razor invalidates the imported TrackItem reference.
    expect(hostSource).toMatch(/subcreator_try_set_mogrt_controls[\s\S]*?subcreator_try_razor_mogrt_duration\(/);
    expect(hostSource).not.toContain("subcreator_disable_mogrt_time_remapping");
    expect(hostSource).not.toContain("qeClip.setSpeed");
  });
});

describe("Whisper sequence audio export", () => {
  it("uses Adobe Media Encoder before the legacy direct export path", () => {
    // // AME avoids Premiere 26.x direct-export failures that can abort CEP evalScript before JSON is returned.
    const exportSource = hostSource.match(
      /function subcreator_export_active_sequence_audio[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_runtime_push_unique/
    );

    expect(exportSource).not.toBeNull();
    expect(exportSource?.[0]).toMatch(
      /subcreator_try_encoder_sequence_export\(sequence,[\s\S]*?subcreator_try_direct_sequence_export\(sequence,/
    );
  });

  it("skips exportAsMediaDirect on Premiere 26 and newer", () => {
    // // The direct exporter remains available only for older hosts where it is less likely to break evalScript.
    expect(hostSource).toMatch(/premiereMajorVersion >= 26/);
    expect(hostSource).toContain("exportAsMediaDirect skipped on Premiere");
  });
});
