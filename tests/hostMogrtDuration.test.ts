import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const hostSourcePath = fileURLToPath(new URL("../src/host/SubCreatorHost.jsx", import.meta.url));
const hostSource = readFileSync(hostSourcePath, "utf8");
const panelSourcePath = fileURLToPath(new URL("../src/panel/cepBridge.ts", import.meta.url));
const panelSource = readFileSync(panelSourcePath, "utf8");
const mainSourcePath = fileURLToPath(new URL("../src/panel/main.ts", import.meta.url));
const mainSource = readFileSync(mainSourcePath, "utf8");

function readDurationHelperSource(): string {
  // // Isolate the razor helper so the regression check ignores unrelated timeline edits elsewhere in the host.
  const match = hostSource.match(
    /function subcreator_try_razor_mogrt_duration[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_push_unique_path/
  );

  expect(match).not.toBeNull();
  return match?.[0] ?? "";
}

function readNearbyGapHelper(): (endSeconds: number, nextStartSeconds: number, frameDurationSeconds: number) => number {
  // // Execute the pure ExtendScript-compatible helper so frame-boundary behavior is tested with real values.
  const match = hostSource.match(
    /function subcreator_snap_nearby_caption_end[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_track_item_starts_near_seconds/
  );

  expect(match).not.toBeNull();
  const helperSource = (match?.[0] ?? "").replace(/\r?\nfunction subcreator_track_item_starts_near_seconds[\s\S]*$/, "");
  return new Function(`${helperSource}; return subcreator_snap_nearby_caption_end;`)() as (
    endSeconds: number,
    nextStartSeconds: number,
    frameDurationSeconds: number
  ) => number;
}

function readFrameSnapHelper(): (seconds: number, frameDurationSeconds: number) => number {
  // // Execute the frame quantizer against timestamps observed in real Premiere host logs.
  const match = hostSource.match(
    /function subcreator_snap_seconds_to_nearest_frame[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_track_item_starts_near_seconds/
  );

  expect(match).not.toBeNull();
  const helperSource = (match?.[0] ?? "").replace(/\r?\nfunction subcreator_track_item_starts_near_seconds[\s\S]*$/, "");
  return new Function(`${helperSource}; return subcreator_snap_seconds_to_nearest_frame;`)() as (
    seconds: number,
    frameDurationSeconds: number
  ) => number;
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

  it("closes gaps up to one sequence frame before applying the razor", () => {
    // // A timestamp just before the next cue must share its boundary instead of rounding one frame early.
    const snapNearbyEnd = readNearbyGapHelper();

    expect(snapNearbyEnd(1.96, 2, 1 / 25)).toBe(2);
    expect(snapNearbyEnd(1.98, 2, 1 / 25)).toBe(2);
    expect(snapNearbyEnd(1.97, 2, 1001 / 30000)).toBe(2);
    expect(snapNearbyEnd(1.99, 2, 1 / 60)).toBe(2);
  });

  it("preserves real gaps and overlaps", () => {
    // // Pauses longer than one frame and overlapping cues retain their authored timing.
    const snapNearbyEnd = readNearbyGapHelper();

    expect(snapNearbyEnd(1.95, 2, 1 / 25)).toBe(1.95);
    expect(snapNearbyEnd(2.01, 2, 1 / 25)).toBe(2.01);
  });

  it("uses the normalized end for generation and Text editor rebuilds", () => {
    // // Both MOGRT creation paths must feed the same adjusted boundary to duration controls and the razor helper.
    expect(hostSource).toMatch(/nextCue[\s\S]*?subcreator_snap_nearby_caption_end[\s\S]*?subcreator_clone_style_config_with_clip_duration/);
    expect(hostSource).toMatch(/nextEditedItem[\s\S]*?subcreator_snap_nearby_caption_end[\s\S]*?rebuiltDurationStyleConfig/);
  });

  it("aligns logged 24 fps boundaries before both import and razor operations", () => {
    // // Premiere previously rounded these starts up/nearest while QE rounded the matching ends down.
    const snapToFrame = readFrameSnapHelper();
    const frameDuration = 1 / 24;

    expect(snapToFrame(1.4, frameDuration)).toBeCloseTo(1.41666666666667, 12);
    expect(snapToFrame(2.44, frameDuration)).toBeCloseTo(2.45833333333333, 12);
    expect(snapToFrame(3.14, frameDuration)).toBeCloseTo(3.125, 12);
  });

  it("biases exact QE razor boundaries away from the preceding frame", () => {
    // // The tiny in-frame offset prevents getFormatted from flooring an exact floating-point frame boundary.
    expect(hostSource).toContain("Number(seconds) + frameDurationSeconds * 0.001");
  });
});

describe("Whisper sequence audio export", () => {
  it("does not depend on Premiere preloading JSON in ExtendScript", () => {
    // // Fresh Premiere sessions may not expose JSON until another Adobe panel polyfills it.
    expect(hostSource).toContain('if (typeof JSON !== "object")');
    expect(panelSource).toContain("buildExtendScriptJsonBootstrap");
  });

  it("falls back to the entire sequence when the host cannot read In/Out", () => {
    // // A malformed Premiere range response should not block Whisper generation before audio export can run.
    expect(panelSource).toContain('fallbackReason: "Unable to read active sequence In/Out range; using entire sequence."');
    expect(mainSource).toMatch(/range\.fallbackReason \|\| range\.hostError[\s\S]*?return \{[\s\S]*?hostError: range\.hostError/);
  });

  it("selects exactly one exporter per isolated host call", () => {
    // // Separate calls let the panel recover from Premiere 26.x aborting the direct evalScript response.
    const exportSource = hostSource.match(
      /function subcreator_export_active_sequence_audio[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_runtime_push_unique/
    );

    expect(exportSource).not.toBeNull();
    expect(exportSource?.[0]).toContain('exportMode === "media_encoder"');
    expect(exportSource?.[0]).toMatch(/\? subcreator_try_encoder_sequence_export\([\s\S]*?: subcreator_try_direct_sequence_export\(/);
  });

  it("allows Premiere direct export on current host versions", () => {
    // // CEP now owns the AME fallback, so Premiere 26 is no longer excluded from the faster direct path.
    expect(hostSource).not.toMatch(/premiereMajorVersion >= 26/);
    expect(hostSource).toContain("sequence.exportAsMediaDirect(outputPath, presetPath, workAreaType)");
  });

  it("tries Premiere before falling back to Media Encoder in a separate call", () => {
    // // The direct-first order keeps normal exports inside Premiere while preserving a recoverable AME path.
    const exportSource = panelSource.match(
      /export async function exportActiveSequenceAudioForWhisper[\s\S]*?\r?\n}\r?\n\r?\nasync function waitForStableCepFile/
    );

    expect(exportSource).not.toBeNull();
    expect(exportSource?.[0]).toMatch(/runExport\("premiere_direct"\)[\s\S]*?runExport\("media_encoder"\)/);
    expect(exportSource?.[0]).toContain('recoverCompletedExport("premiere_direct", outputPath, 30000)');
    expect(exportSource?.[0]).toContain("waitForStableCepFile(modules, candidatePath, timeoutMs, 3)");
    expect(exportSource?.[0]).toContain("`${outputBase}-${exportMethod}.wav`");
  });
});
