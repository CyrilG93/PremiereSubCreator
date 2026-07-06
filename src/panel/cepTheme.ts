// // Provide a reusable Premiere CEP theme bridge for panels that should follow the host UI brightness.
export interface RgbColor {
  red: number;
  green: number;
  blue: number;
}

interface CepUIColor {
  color?: Partial<RgbColor>;
}

interface CepHostSkinInfo {
  baseFontFamily?: string;
  panelBackgroundColor?: Partial<RgbColor> | CepUIColor;
  panelBackgroundColorSRGB?: Partial<RgbColor> | CepUIColor;
  systemHighlightColor?: Partial<RgbColor> | CepUIColor;
}

interface CepHostEnvironment {
  appSkinInfo?: CepHostSkinInfo;
}

declare global {
  interface Window {
    __adobe_cep__?: {
      evalScript: (script: string, callback: (result: string) => void) => void;
      getHostEnvironment?: () => string;
      addEventListener?: (eventName: string, listener: (event?: unknown) => void) => void;
      removeEventListener?: (eventName: string, listener: (event?: unknown) => void) => void;
    };
  }
}

const CEP_THEME_COLOR_CHANGED_EVENT = "com.adobe.csxs.events.ThemeColorChanged";
let hostThemeListenerBound = false;

function clampColorChannel(value: unknown): number | null {
  // // Normalize known numeric RGB channels and reject missing channels instead of turning them into black.
  const numericValue = Number(value);
  if (!Number.isFinite(numericValue)) {
    return null;
  }

  return Math.max(0, Math.min(255, Math.round(numericValue)));
}

function readRgbTriplet(value: unknown): RgbColor | null {
  // // Parse direct `{ red, green, blue }` payloads used by some CEP color fields.
  if (!value || typeof value !== "object") {
    return null;
  }

  const payload = value as Record<string, unknown>;
  const red = clampColorChannel(payload.red);
  const green = clampColorChannel(payload.green);
  const blue = clampColorChannel(payload.blue);
  if (red === null || green === null || blue === null) {
    return null;
  }

  return { red, green, blue };
}

function readCepRgbColor(value: unknown): RgbColor | null {
  // // Parse both CEP `RGBColor` and `UIColor.color` shapes documented by Adobe for appSkinInfo.
  const directColor = readRgbTriplet(value);
  if (directColor) {
    return directColor;
  }

  if (!value || typeof value !== "object") {
    return null;
  }

  return readRgbTriplet((value as Record<string, unknown>).color);
}

function mixRgbColor(left: RgbColor, right: RgbColor, rightWeight: number): RgbColor {
  // // Blend two colors so panel surfaces stay close to Premiere base shades.
  const clampedWeight = Math.max(0, Math.min(1, rightWeight));
  const leftWeight = 1 - clampedWeight;
  return {
    red: Math.round(left.red * leftWeight + right.red * clampedWeight),
    green: Math.round(left.green * leftWeight + right.green * clampedWeight),
    blue: Math.round(left.blue * leftWeight + right.blue * clampedWeight)
  };
}

function offsetRgbColor(color: RgbColor, delta: number): RgbColor {
  // // Brighten or darken one color uniformly while keeping channels inside RGB bounds.
  return {
    red: Math.max(0, Math.min(255, Math.round(color.red + delta))),
    green: Math.max(0, Math.min(255, Math.round(color.green + delta))),
    blue: Math.max(0, Math.min(255, Math.round(color.blue + delta)))
  };
}

function rgbColorLuminance(color: RgbColor): number {
  // // Estimate perceived brightness to switch between light, dark, and darkest Premiere skin variants.
  return (0.2126 * color.red + 0.7152 * color.green + 0.0722 * color.blue) / 255;
}

function buildHostAccentColor(baseColor: RgbColor, backgroundColor: RgbColor, isLightTheme: boolean): RgbColor {
  // // Anchor the accent around Adobe's UI blue while still reacting to the host skin highlight color.
  const adobeBlue = { red: 0, green: 100, blue: 203 };
  const accentSeed = mixRgbColor(baseColor, adobeBlue, 0.72);
  const mixedColor = mixRgbColor(accentSeed, backgroundColor, isLightTheme ? 0.04 : 0.01);
  return isLightTheme ? offsetRgbColor(mixedColor, -6) : offsetRgbColor(mixedColor, 10);
}

function normalizeHostPanelBackground(color: RgbColor): RgbColor {
  // // Keep Premiere skin variants readable while still following host light/dark/darkest appearance changes.
  const luminance = rgbColorLuminance(color);
  if (luminance <= 0.16) {
    return mixRgbColor(color, { red: 52, green: 52, blue: 52 }, 0.9);
  }
  if (luminance <= 0.32) {
    return mixRgbColor(color, { red: 58, green: 58, blue: 58 }, 0.42);
  }
  if (luminance >= 0.7) {
    return mixRgbColor(color, { red: 198, green: 198, blue: 198 }, 0.24);
  }
  if (luminance >= 0.55) {
    return mixRgbColor(color, { red: 184, green: 184, blue: 184 }, 0.12);
  }

  return color;
}

function setRootRgbVariable(variableName: string, color: RgbColor): void {
  // // Publish one RGB color both as `rgb(...)` and raw `r, g, b` triplet for CSS reuse.
  const root = document.documentElement;
  root.style.setProperty(variableName, `rgb(${color.red}, ${color.green}, ${color.blue})`);
  root.style.setProperty(`${variableName}-rgb`, `${color.red}, ${color.green}, ${color.blue}`);
}

function readHostEnvironmentSkin(): CepHostEnvironment | null {
  // // Read the latest CEP host skin payload so live Premiere preference changes are reflected.
  const rawEnvironment = window.__adobe_cep__?.getHostEnvironment?.();
  if (!rawEnvironment) {
    return null;
  }

  try {
    const parsed = JSON.parse(rawEnvironment) as CepHostEnvironment;
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

export function applyPremierePanelTheme(): void {
  // // Derive neutral Premiere-like CSS tokens from the current host theme.
  const hostEnvironment = readHostEnvironmentSkin();
  if (!hostEnvironment?.appSkinInfo) {
    return;
  }

  const skinInfo = hostEnvironment.appSkinInfo;
  const panelBackground =
    readCepRgbColor(skinInfo.panelBackgroundColorSRGB) ||
    readCepRgbColor(skinInfo.panelBackgroundColor) || {
      red: 48,
      green: 48,
      blue: 48
    };
  const highlightColor = readCepRgbColor(skinInfo.systemHighlightColor) || {
    red: 70,
    green: 137,
    blue: 255
  };
  const hostLuminance = rgbColorLuminance(panelBackground);
  const normalizedPanelBackground = normalizeHostPanelBackground(panelBackground);
  const isLightTheme = hostLuminance >= 0.55;
  const isDarkestTheme = hostLuminance <= 0.18;
  const textPrimary = isLightTheme
    ? { red: 36, green: 36, blue: 36 }
    : { red: 236, green: 236, blue: 236 };
  const textDim = mixRgbColor(textPrimary, normalizedPanelBackground, isLightTheme ? 0.48 : 0.38);
  const bgPrimary = normalizedPanelBackground;
  const bgSurface = offsetRgbColor(normalizedPanelBackground, isLightTheme ? 8 : isDarkestTheme ? 8 : 6);
  const bgSoft = offsetRgbColor(normalizedPanelBackground, isLightTheme ? 13 : isDarkestTheme ? 12 : 10);
  const bgInput = offsetRgbColor(normalizedPanelBackground, isLightTheme ? -3 : isDarkestTheme ? -2 : -1);
  const bgCard = offsetRgbColor(normalizedPanelBackground, isLightTheme ? 5 : isDarkestTheme ? 5 : 4);
  const accent = buildHostAccentColor(highlightColor, normalizedPanelBackground, isLightTheme);
  const accentSoft = isLightTheme ? offsetRgbColor(accent, -4) : offsetRgbColor(accent, 18);
  const border = offsetRgbColor(normalizedPanelBackground, isLightTheme ? -28 : 16);
  const borderStrong = offsetRgbColor(normalizedPanelBackground, isLightTheme ? -42 : 24);
  const buttonPrimary = mixRgbColor(accent, bgSurface, 0.28);
  const buttonPrimaryAlt = mixRgbColor(accent, bgPrimary, 0.2);
  const buttonPrimaryText = rgbColorLuminance(buttonPrimary) >= 0.5
    ? { red: 16, green: 16, blue: 16 }
    : { red: 246, green: 246, blue: 246 };
  const root = document.documentElement;

  root.dataset.themeVariant = isLightTheme ? "light" : isDarkestTheme ? "darkest" : "dark";
  setRootRgbVariable("--bg-primary", bgPrimary);
  setRootRgbVariable("--bg-surface", bgSurface);
  setRootRgbVariable("--bg-soft", bgSoft);
  setRootRgbVariable("--bg-input", bgInput);
  setRootRgbVariable("--bg-card", bgCard);
  setRootRgbVariable("--text-primary", textPrimary);
  setRootRgbVariable("--text", textPrimary);
  setRootRgbVariable("--text-dim", textDim);
  setRootRgbVariable("--text-muted", textDim);
  setRootRgbVariable("--accent", accent);
  setRootRgbVariable("--accent-soft", accentSoft);
  setRootRgbVariable("--border", border);
  setRootRgbVariable("--border-strong", borderStrong);
  setRootRgbVariable("--button-primary-bg", buttonPrimary);
  setRootRgbVariable("--button-primary-bg-alt", buttonPrimaryAlt);
  setRootRgbVariable("--button-primary-text", buttonPrimaryText);
  root.style.setProperty("--shadow", isLightTheme ? "0 6px 18px rgba(0, 0, 0, 0.08)" : "0 6px 18px rgba(0, 0, 0, 0.24)");

  const baseFontFamily = String(skinInfo.baseFontFamily || "").trim();
  if (baseFontFamily) {
    root.style.setProperty("--ui-font-family", `"${baseFontFamily}", "Avenir Next", "Helvetica Neue", sans-serif`);
  }
}

export function bindPremiereThemeListener(): void {
  // // Subscribe once to CEP theme changes so the panel follows Premiere appearance switches live.
  if (hostThemeListenerBound || typeof window.__adobe_cep__?.addEventListener !== "function") {
    return;
  }

  hostThemeListenerBound = true;
  window.__adobe_cep__.addEventListener(CEP_THEME_COLOR_CHANGED_EVENT, () => {
    applyPremierePanelTheme();
  });
}
