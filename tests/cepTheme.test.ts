// // Verify that CEP theme tokens follow the host skin payload shapes used by Adobe panels.
import { afterEach, describe, expect, it, vi } from "vitest";
import { applyPremierePanelTheme, bindPremiereThemeListener } from "../src/panel/cepTheme";

function installFakeDom(hostEnvironment: unknown): Map<string, string> {
  // // Provide the tiny DOM and CEP surface needed by the theme helper in node tests.
  const variables = new Map<string, string>();
  const documentElement = {
    dataset: {} as Record<string, string>,
    style: {
      setProperty: (name: string, value: string) => {
        variables.set(name, value);
      }
    }
  };

  vi.stubGlobal("document", {
    documentElement
  });
  vi.stubGlobal("window", {
    __adobe_cep__: {
      getHostEnvironment: () => JSON.stringify(hostEnvironment),
      addEventListener: vi.fn()
    }
  });

  return variables;
}

afterEach(() => {
  // // Restore globals so CEP theme tests do not leak into unrelated node tests.
  vi.unstubAllGlobals();
});

describe("applyPremierePanelTheme", () => {
  it("reads nested UIColor payloads for Premiere light mode", () => {
    const variables = installFakeDom({
      appSkinInfo: {
        panelBackgroundColorSRGB: {
          color: { red: 210, green: 210, blue: 210 }
        },
        systemHighlightColor: { red: 66, green: 137, blue: 245 },
        baseFontFamily: "Adobe Clean"
      }
    });

    applyPremierePanelTheme();

    expect(document.documentElement.dataset.themeVariant).toBe("light");
    expect(variables.get("--bg-primary")).toBe("rgb(236, 236, 236)");
    expect(variables.get("--bg-surface")).toBe("rgb(229, 229, 229)");
    expect(variables.get("--bg-soft")).toBe("rgb(234, 234, 234)");
    expect(variables.get("--text-primary")).toBe("rgb(36, 36, 36)");
    expect(variables.get("--text")).toBe("rgb(36, 36, 36)");
    expect(variables.get("--ui-font-family")).toContain("Adobe Clean");
  });

  it("falls back from malformed sRGB UIColor to direct RGB colors", () => {
    const variables = installFakeDom({
      appSkinInfo: {
        panelBackgroundColorSRGB: {
          color: {}
        },
        panelBackgroundColor: { red: 29, green: 29, blue: 29 }
      }
    });

    applyPremierePanelTheme();

    expect(document.documentElement.dataset.themeVariant).toBe("darkest");
    expect(variables.get("--bg-primary")).toBe("rgb(28, 28, 28)");
  });
});

describe("bindPremiereThemeListener", () => {
  it("subscribes to Adobe's CEP theme-change event", () => {
    installFakeDom({
      appSkinInfo: {
        panelBackgroundColorSRGB: {
          color: { red: 58, green: 58, blue: 58 }
        }
      }
    });

    bindPremiereThemeListener();

    expect(window.__adobe_cep__?.addEventListener).toHaveBeenCalledWith(
      "com.adobe.csxs.events.ThemeColorChanged",
      expect.any(Function)
    );
  });
});
