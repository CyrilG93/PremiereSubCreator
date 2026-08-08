import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const hostSourcePath = fileURLToPath(new URL("../src/host/SubCreatorHost.jsx", import.meta.url));
const hostSource = readFileSync(hostSourcePath, "utf8");
const panelSourcePath = fileURLToPath(new URL("../src/panel/main.ts", import.meta.url));
const panelSource = readFileSync(panelSourcePath, "utf8");

type HostComponent = {
  displayName?: string;
  matchName?: string;
  properties: {
    numItems: number;
  };
};

type HostTrackItem = {
  components: {
    numItems: number;
    [index: number]: HostComponent;
  };
  getMGTComponent?: () => HostComponent | null;
};

function createComponent(matchName: string, displayName: string): HostComponent {
  // // Build the minimum ExtendScript component shape used by the real graphic-selection helper.
  return {
    matchName,
    displayName,
    properties: {
      numItems: 1
    }
  };
}

function createTrackItem(components: HostComponent[], mgtComponent: HostComponent | null = null): HostTrackItem {
  // // Mirror Premiere's indexed ComponentCollection while keeping tests independent from the Adobe host.
  const collection = {
    numItems: components.length
  } as HostTrackItem["components"];

  components.forEach((component, index) => {
    collection[index] = component;
  });

  return {
    components: collection,
    getMGTComponent: () => mgtComponent
  };
}

function readMogrtComponentCollector(): (trackItem: HostTrackItem) => HostComponent[] {
  // // Execute the actual ExtendScript-compatible detection helpers instead of duplicating their filtering rules in tests.
  const match = hostSource.match(
    /function subcreator_is_graphic_component[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_collect_selected_mogrt_items/
  );

  expect(match).not.toBeNull();
  const helperSource = (match?.[0] ?? "").replace(
    /\r?\nfunction subcreator_collect_selected_mogrt_items[\s\S]*$/,
    ""
  );
  return new Function(`${helperSource}; return subcreator_get_mogrt_components_from_track_item;`)() as (
    trackItem: HostTrackItem
  ) => HostComponent[];
}

describe("MOGRT and graphic selection filtering", () => {
  it("ignores ordinary video clips with only intrinsic components and effects", () => {
    // // Motion, Opacity, and normal effects exist on media clips and must not trigger Visual editor loading.
    const collectComponents = readMogrtComponentCollector();
    const motion = createComponent("AE.ADBE Motion", "Motion");
    const opacity = createComponent("AE.ADBE Opacity", "Opacity");
    const lumetri = createComponent("AE.ADBE Lumetri", "Lumetri Color");

    expect(collectComponents(createTrackItem([motion, opacity, lumetri]))).toEqual([]);
  });

  it("keeps all controls for an After Effects MOGRT identified by getMGTComponent", () => {
    // // AE MOGRT identity comes from the dedicated API even when the returned component has a custom name.
    const collectComponents = readMogrtComponentCollector();
    const motion = createComponent("AE.ADBE Motion", "Motion");
    const opacity = createComponent("AE.ADBE Opacity", "Opacity");
    const exposedControls = createComponent("Vendor.Custom Controls", "Controls");

    expect(collectComponents(createTrackItem([motion, opacity, exposedControls], exposedControls))).toEqual([
      exposedControls,
      motion,
      opacity
    ]);
  });

  it("recognizes Premiere-authored text graphics when getMGTComponent returns null", () => {
    // // Premiere graphics expose stable text/vector match names through the normal component chain.
    const collectComponents = readMogrtComponentCollector();
    const motion = createComponent("AE.ADBE Motion", "Motion");
    const text = createComponent("AE.ADBE Text", "Text");

    expect(collectComponents(createTrackItem([motion, text]))).toEqual([motion, text]);
  });

  it("recognizes Premiere graphic parameter capsules without matching ordinary effects", () => {
    // // Recent Premiere versions expose native graphic controls as an AE.ADBE Capsule component.
    const collectComponents = readMogrtComponentCollector();
    const capsule = createComponent("AE.ADBE Capsule", "Graphic Parameters");

    expect(collectComponents(createTrackItem([capsule]))).toEqual([capsule]);
  });
});

describe("Visual selection watcher lifecycle", () => {
  it("silences host selection events and polling while a docked CEP document is hidden", () => {
    // // A hidden panel must not keep competing with other extensions for Premiere's single ExtendScript host.
    expect(hostSource).toMatch(
      /function subcreator_notify_visual_selection_changed\(\)[\s\S]*?if \(!subcreator_visual_selection_watcher_enabled\)/
    );
    expect(panelSource).toContain('document.visibilityState !== "hidden"');
    expect(panelSource).toContain('document.addEventListener("visibilitychange"');
    expect(panelSource).toContain("setVisualSelectionWatcherEnabled(isVisualSelectionMonitoringAllowed())");
  });
});

describe("Visual property bulk apply", () => {
  it("sends one host batch for Copy properties Apply and reuses resolved component paths", () => {
    // // Large subtitle selections must avoid one CEP evaluation, selection scan, and redraw per timeline clip.
    const bridgeSourcePath = fileURLToPath(new URL("../src/panel/cepBridge.ts", import.meta.url));
    const bridgeSource = readFileSync(bridgeSourcePath, "utf8");

    expect(panelSource).not.toContain("for (let clipIndex = 0; clipIndex < selectedCount; clipIndex += 1)");
    expect(panelSource).toContain("includeDebug: verboseLogsEnabled");
    expect(bridgeSource).toContain("includeDebug: options?.includeDebug === true");
    expect(hostSource).toContain("resolvedPropertiesByPath");
    expect(hostSource).toContain("subcreator_visual_resolve_property_from_track_item(clip, resolvedPath, clipComponents)");
    expect(hostSource).toContain("property.setValue(hostVector, false)");
    expect(hostSource).toContain("subcreator_force_color_apply_visual_refresh(sequence, mogrtItems)");
  });
});
