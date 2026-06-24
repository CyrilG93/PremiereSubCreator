import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const panelSourcePath = fileURLToPath(new URL("../src/panel/main.ts", import.meta.url));

describe("generate progress startup", () => {
  it("shows the progress bar before slow generation preflight checks", () => {
    // // The first visible feedback must happen before Whisper runtime detection can make the panel feel frozen.
    const panelSource = readFileSync(panelSourcePath, "utf8");
    const initialProgressIndex = panelSource.indexOf('await updateGenerateProgress(1, translate("progress.prepareGeneration"), true);');
    const whisperPreflightIndex = panelSource.indexOf("await enforceWhisperSourceAvailability();");

    expect(initialProgressIndex).toBeGreaterThan(0);
    expect(whisperPreflightIndex).toBeGreaterThan(initialProgressIndex);
  });

  it("waits until after a paint opportunity before continuing expensive panel work", () => {
    // // Resolving directly inside requestAnimationFrame can continue JavaScript before the browser paints the new progress state.
    const panelSource = readFileSync(panelSourcePath, "utf8");
    const waitForPaintIndex = panelSource.indexOf("function waitForNextPaint()");
    const frameIndex = panelSource.indexOf("window.requestAnimationFrame", waitForPaintIndex);
    const timeoutIndex = panelSource.indexOf("window.setTimeout(resolve, 0)", frameIndex);

    expect(waitForPaintIndex).toBeGreaterThan(0);
    expect(frameIndex).toBeGreaterThan(waitForPaintIndex);
    expect(timeoutIndex).toBeGreaterThan(frameIndex);
  });
});
