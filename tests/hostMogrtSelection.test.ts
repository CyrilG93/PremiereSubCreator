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
  nodeId?: string;
  projectItem?: {
    nodeId?: string;
    name?: string;
  };
  components: {
    numItems: number;
    [index: number]: HostComponent;
  };
  getMGTComponent?: () => HostComponent | null;
};

type HostSelectionAnalysis = {
  selectedItems: HostTrackItem[];
  mogrtItems: HostTrackItem[];
  firstTrackItem: HostTrackItem | null;
  firstComponents: HostComponent[];
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

function readSelectedMogrtAnalyzer(
  collectComponents: (trackItem: HostTrackItem) => HostComponent[],
  onComponentRead: () => void
): (sequence: { selectedItems: HostTrackItem[] }) => HostSelectionAnalysis {
  // // Execute the real selection grouping helper with controlled host-selection and component-read adapters.
  const match = hostSource.match(
    /function subcreator_get_track_item_project_identity[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_collect_selected_mogrt_items/
  );

  expect(match).not.toBeNull();
  const helperSource = (match?.[0] ?? "").replace(
    /\r?\nfunction subcreator_collect_selected_mogrt_items[\s\S]*$/,
    ""
  );
  return new Function(
    "subcreator_trim_string",
    "subcreator_get_selected_track_items",
    "subcreator_get_mogrt_components_from_track_item",
    `${helperSource}; return subcreator_analyze_selected_mogrt_items;`
  )(
    (value: unknown) => String(value ?? "").trim(),
    (sequence: { selectedItems: HostTrackItem[] }) => sequence.selectedItems,
    (trackItem: HostTrackItem) => {
      // // Count only real component classifications so same-project cache reuse stays directly testable.
      onComponentRead();
      return collectComponents(trackItem);
    }
  ) as (sequence: { selectedItems: HostTrackItem[] }) => HostSelectionAnalysis;
}

function readVisualSelectionPulse(): (trackItems: Array<{ setSelected: (selected: boolean, updateUi: boolean) => void }>) => boolean {
  // // Execute the real color-refresh selection pulse with a no-delay ExtendScript sleep adapter.
  const match = hostSource.match(
    /function subcreator_try_set_track_item_selected[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_force_color_apply_visual_refresh/
  );

  expect(match).not.toBeNull();
  const helperSource = (match?.[0] ?? "").replace(
    /\r?\nfunction subcreator_force_color_apply_visual_refresh[\s\S]*$/,
    ""
  );
  return new Function("$", `${helperSource}; return subcreator_pulse_selected_track_items_for_refresh;`)({
    // // Keep the unit test synchronous while preserving the host helper's expected API shape.
    sleep: () => undefined
  }) as (trackItems: Array<{ setSelected: (selected: boolean, updateUi: boolean) => void }>) => boolean;
}

function readVisualDescriptorNormalizer(): (descriptor: Record<string, unknown>) => Record<string, unknown> {
  // // Execute the host descriptor normalizer with lightweight label/group adapters for clone metadata tests.
  const match = hostSource.match(
    /function subcreator_visual_normalize_descriptor[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_visual_should_hide_descriptor/
  );

  expect(match).not.toBeNull();
  const helperSource = (match?.[0] ?? "").replace(
    /\r?\nfunction subcreator_visual_should_hide_descriptor[\s\S]*$/,
    ""
  );
  return new Function(
    "subcreator_trim_string",
    "subcreator_visual_normalize_label_key",
    "subcreator_visual_group_mentions",
    "subcreator_visual_append_group_suffix",
    `${helperSource}; return subcreator_visual_normalize_descriptor;`
  )(
    (value: unknown) => String(value ?? "").trim(),
    (value: unknown) =>
      String(value ?? "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-z0-9]/g, ""),
    (groupPath: unknown, token: unknown) => String(groupPath ?? "").toLowerCase().includes(String(token ?? "").toLowerCase()),
    (groupPath: unknown, suffix: unknown) => `${String(groupPath ?? "").trim()} / ${String(suffix ?? "").trim()}`
  ) as (descriptor: Record<string, unknown>) => Record<string, unknown>;
}

function readNativeColorWriter(): (
  property: { setColorValue: (...channels: Array<number | boolean>) => void },
  rgb: { red: number; green: number; blue: number },
  alpha: number
) => boolean {
  // // Execute the real native color writer with a minimal clamp adapter to verify Premiere's positional channel order.
  const match = hostSource.match(
    /function subcreator_visual_try_set_native_color_value[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_visual_extract_alpha_channel/
  );

  expect(match).not.toBeNull();
  const helperSource = (match?.[0] ?? "").replace(
    /\r?\nfunction subcreator_visual_extract_alpha_channel[\s\S]*$/,
    ""
  );
  return new Function(
    "subcreator_visual_clamp",
    `${helperSource}; return subcreator_visual_try_set_native_color_value;`
  )((value: number, minValue: number, maxValue: number) => Math.min(maxValue, Math.max(minValue, value))) as (
    property: { setColorValue: (...channels: Array<number | boolean>) => void },
    rgb: { red: number; green: number; blue: number },
    alpha: number
  ) => boolean;
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

  it("classifies timeline instances from the same project item only once", () => {
    // // Repeated subtitle instances from one source MOGRT should share the expensive getMGTComponent identity result.
    const collectComponents = readMogrtComponentCollector();
    const exposedControls = createComponent("Vendor.Custom Controls", "Controls");
    let componentReadCount = 0;
    const analyzeSelection = readSelectedMogrtAnalyzer(collectComponents, () => {
      componentReadCount += 1;
    });
    const selectedItems = [0, 1, 2].map((index) => ({
      ...createTrackItem([], exposedControls),
      nodeId: `track-${index}`,
      projectItem: {
        nodeId: "shared-mogrt-project-item",
        name: "Shared Subtitle"
      }
    }));

    const result = analyzeSelection({ selectedItems });

    expect(componentReadCount).toBe(1);
    expect(result.mogrtItems).toEqual(selectedItems);
    expect(result.firstTrackItem).toBe(selectedItems[0]);
    expect(result.firstComponents).toEqual([exposedControls]);
  });

  it("retries a shared project item after an unavailable first component read", () => {
    // // A transient empty Premiere proxy must not cache a false result for every later MOGRT instance.
    const collectComponents = readMogrtComponentCollector();
    const motion = createComponent("AE.ADBE Motion", "Motion");
    const exposedControls = createComponent("Vendor.Custom Controls", "Controls");
    let componentReadCount = 0;
    const analyzeSelection = readSelectedMogrtAnalyzer(collectComponents, () => {
      componentReadCount += 1;
    });
    const unavailableItem = {
      ...createTrackItem([motion]),
      nodeId: "track-unavailable",
      projectItem: {
        nodeId: "shared-transient-project-item",
        name: "Transient Subtitle"
      }
    };
    const availableItem = {
      ...createTrackItem([], exposedControls),
      nodeId: "track-available",
      projectItem: {
        nodeId: "shared-transient-project-item",
        name: "Transient Subtitle"
      }
    };

    const result = analyzeSelection({ selectedItems: [unavailableItem, availableItem] });

    expect(componentReadCount).toBe(2);
    expect(result.mogrtItems).toEqual([availableItem]);
    expect(result.firstTrackItem).toBe(availableItem);
    expect(result.firstComponents).toEqual([exposedControls]);
  });

  it("keeps the Visual selection signature free of component and track scans", () => {
    // // Polling must use TrackItem identity and timing only, leaving MOGRT inspection to the full read.
    const signatureSource = hostSource.match(
      /function subcreator_build_selected_mogrt_visual_signature[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_get_selected_mogrt_visual_signature/
    )?.[0];

    expect(signatureSource).toContain("trackItem.nodeId");
    expect(signatureSource).toContain("parts.sort()");
    expect(signatureSource).not.toContain("subcreator_get_mogrt_components_from_track_item(trackItem)");
    expect(signatureSource).not.toContain("subcreator_find_track_item_video_track_index(sequence, trackItem)");
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

  it("keeps expensive Visual read diagnostics opt-in and reuses first-clip components", () => {
    // // Normal auto-refresh reads should avoid diagnostic traversal and repeated component collection.
    const bridgeSourcePath = fileURLToPath(new URL("../src/panel/cepBridge.ts", import.meta.url));
    const bridgeSource = readFileSync(bridgeSourcePath, "utf8");
    const listSource = hostSource.match(
      /function subcreator_list_selected_mogrt_properties[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_get_selected_mogrt_count/
    )?.[0];

    expect(listSource).toContain("subcreator_analyze_selected_mogrt_items(sequence)");
    expect(listSource).toContain("var rawComponents = selectionAnalysis.firstComponents");
    expect(listSource).toContain("includeFullDebug &&");
    expect(listSource).toContain("subcreator_visual_resolve_property_from_track_item(firstTrackItem, item.path, rawComponents)");
    expect(bridgeSource).toContain("subcreator_list_selected_mogrt_properties(${options?.includeDebug === true");
    expect(panelSource).toContain("includeDebug: verboseLogsEnabled");
  });

  it("repaints the timeline only after restoring the complete MOGRT selection", () => {
    // // Premiere must not snapshot the UI after only the first selected clip has been restored.
    const pulseSelection = readVisualSelectionPulse();
    const selectedState = [true, true, true];
    const uiSnapshots: number[][] = [];
    const items = selectedState.map((_selected, itemIndex) => ({
      setSelected: (selected: boolean, updateUi: boolean) => {
        // // Mirror Premiere state changes and capture only moments where the host requests an interface repaint.
        selectedState[itemIndex] = selected;
        if (updateUi) {
          uiSnapshots.push(
            selectedState.reduce<number[]>((selectedIndexes, itemSelected, selectedIndex) => {
              if (itemSelected) {
                selectedIndexes.push(selectedIndex);
              }
              return selectedIndexes;
            }, [])
          );
        }
      }
    }));

    expect(pulseSelection(items)).toBe(true);
    expect(selectedState).toEqual([true, true, true]);
    expect(uiSnapshots).toEqual([[], [0, 1, 2]]);
  });

  it("excludes Clip Duration from full Copy/Apply property clones", () => {
    // // Duration drives each clip's animation timing and must never be copied as a shared visual style value.
    const normalizeDescriptor = readVisualDescriptorNormalizer();
    const bridgeSourcePath = fileURLToPath(new URL("../src/panel/cepBridge.ts", import.meta.url));
    const bridgeSource = readFileSync(bridgeSourcePath, "utf8");

    expect(
      normalizeDescriptor({
        path: "c0|4",
        displayName: "Clip Duration",
        groupPath: "Graphic Parameters",
        valueType: "number",
        controlKind: "slider",
        value: 1.5
      }).excludeFromClone
    ).toBe(true);
    expect(
      normalizeDescriptor({
        path: "c0|5",
        displayName: "Position",
        groupPath: "Graphic Parameters",
        valueType: "json",
        controlKind: "vector",
        value: "[0,0]"
      }).excludeFromClone
    ).toBeUndefined();
    expect(bridgeSource).toContain("excludeFromClone: item.excludeFromClone === true");
    expect(panelSource).toContain("if (includeUnchanged && excludeFromClone)");
  });

  it("writes native colors in Premiere's documented ARGB order on the first attempt", () => {
    // // A red picker value must not be interpreted as a green or blue channel before a second Apply.
    const writeNativeColor = readNativeColorWriter();
    const calls: Array<Array<number | boolean>> = [];
    const property = {
      setColorValue: (...channels: Array<number | boolean>) => {
        calls.push(channels);
      }
    };

    expect(writeNativeColor(property, { red: 255, green: 0, blue: 0 }, 1)).toBe(true);
    expect(calls).toEqual([[255, 255, 0, 0, false]]);
    expect(hostSource).not.toContain("setColorValue(fallbackRgbValue.red, fallbackRgbValue.green");
  });
});
