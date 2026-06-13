import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const hostSourcePath = fileURLToPath(new URL("../src/host/SubCreatorHost.jsx", import.meta.url));
const hostSource = readFileSync(hostSourcePath, "utf8");

function readDurationHelperSource(): string {
  // // Isolate the duration helper so the regression check ignores unrelated outPoint handling elsewhere in the host.
  const match = hostSource.match(/function subcreator_try_set_mogrt_duration[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_try_set_mogrt_start/);

  expect(match).not.toBeNull();
  return match?.[0] ?? "";
}

describe("MOGRT duration trimming", () => {
  it("does not shorten the source out point", () => {
    // // Timeline end trimming keeps the generated cue duration while preserving manual extension up to the template duration.
    const helperSource = readDurationHelperSource();

    expect(helperSource).toContain("trackItem.end = endTime");
    expect(helperSource).not.toMatch(/trackItem\.outPoint\s*=/);
  });

  it("disables only the intrinsic time-remapping keyframes after trimming", () => {
    // // Generated clips should remain extendable without disabling animation controls exposed by the MOGRT itself.
    const helperMatch = hostSource.match(
      /function subcreator_disable_mogrt_time_remapping[\s\S]*?\r?\n}\r?\n\r?\nfunction subcreator_push_unique_path/
    );
    const helperSource = helperMatch?.[0] ?? "";
    const durationSource = readDurationHelperSource();

    expect(helperMatch).not.toBeNull();
    expect(helperSource).toContain("subcreator_is_time_remapping_component(component)");
    expect(helperSource).toContain("property.setTimeVarying(false)");
    expect(durationSource).toContain("subcreator_disable_mogrt_time_remapping(trackItem)");
  });
});
