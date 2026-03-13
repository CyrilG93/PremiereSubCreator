// // Provide ExtendScript entry points used by the CEP panel.

function subcreator_ping() {
  // // Return a deterministic host status response.
  return JSON.stringify({ ok: true, message: "Sub Creator host online" });
}

function subcreator_decode_payload(input) {
  // // Decode payload string from URI component format.
  try {
    return decodeURIComponent(input);
  } catch (error) {
    return unescape(input);
  }
}

function subcreator_ok(data) {
  // // Normalize successful host responses for panel-side parsing.
  return JSON.stringify({ ok: true, data: data });
}

function subcreator_error(message) {
  // // Normalize failure responses for panel-side parsing.
  return JSON.stringify({ ok: false, error: String(message) });
}

function subcreator_is_windows() {
  // // Detect Windows platform to build shell commands correctly.
  return $.os && String($.os).toLowerCase().indexOf("windows") !== -1;
}

function subcreator_quote_posix(value) {
  // // Escape shell arguments for POSIX systems.
  return "'" + String(value).replace(/'/g, "'\"'\"'") + "'";
}

function subcreator_quote_cmd(value) {
  // // Escape shell arguments for Windows command line.
  return '"' + String(value).replace(/"/g, '""') + '"';
}

function subcreator_read_file_text(fileRef) {
  // // Read text content from ExtendScript File object.
  if (!fileRef || !fileRef.exists) {
    return "";
  }

  if (!fileRef.open("r")) {
    return "";
  }

  var content = fileRef.read();
  fileRef.close();
  return content;
}

function subcreator_trim_string(value) {
  // // Trim whitespace safely without relying on ES5 String.trim support.
  return String(value || "").replace(/^\s+|\s+$/g, "");
}

function subcreator_read_text_file(encodedPath) {
  // // Read a text file from disk and return content to the panel.
  try {
    var filePath = subcreator_decode_payload(encodedPath || "");
    var file = new File(filePath);
    if (!file.exists) {
      return subcreator_error("File not found: " + filePath);
    }

    if (!file.open("r")) {
      return subcreator_error("Unable to open file: " + filePath);
    }

    var text = file.read();
    file.close();

    return subcreator_ok({ text: text });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_pick_srt_file() {
  // // Open native picker restricted to .srt subtitle files.
  try {
    var selected = File.openDialog("Select SRT subtitle file", function (candidate) {
      if (candidate instanceof Folder) {
        return true;
      }
      return /\.srt$/i.test(String(candidate.name || ""));
    });
    if (!selected) {
      return subcreator_ok({ path: "" });
    }

    return subcreator_ok({ path: selected.fsName });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_pick_audio_file() {
  // // Open native picker to select local media for Whisper transcription.
  try {
    var selected = File.openDialog("Select audio or video file for Whisper transcription");
    if (!selected) {
      return subcreator_ok({ path: "" });
    }

    return subcreator_ok({ path: selected.fsName });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_runtime_push_unique(list, value) {
  // // Push unique string values while keeping ExtendScript compatibility.
  var normalized = subcreator_trim_string(String(value || ""));
  if (!normalized) {
    return;
  }

  var normalizedLower = normalized.toLowerCase();
  for (var i = 0; i < list.length; i += 1) {
    if (String(list[i] || "").toLowerCase() === normalizedLower) {
      return;
    }
  }

  list.push(normalized);
}

function subcreator_runtime_dirname(pathValue) {
  // // Resolve parent folder for Windows or POSIX path strings.
  var normalized = String(pathValue || "").replace(/\\/g, "/");
  var slashIndex = normalized.lastIndexOf("/");
  if (slashIndex < 1) {
    return "";
  }
  return normalized.substring(0, slashIndex);
}

function subcreator_resolve_runtime_config_paths() {
  // // Build user-local runtime config candidates written by installers.
  var candidates = [];

  if (subcreator_is_windows()) {
    var appData = "";
    try {
      appData = subcreator_trim_string($.getenv("APPDATA"));
    } catch (error) {}

    if (!appData && Folder.userData) {
      appData = subcreator_trim_string(Folder.userData.fsName);
    }

    if (appData) {
      subcreator_runtime_push_unique(candidates, appData + "/SubCreator/subcreator-runtime.json");
      subcreator_runtime_push_unique(candidates, appData + "/PremiereSubCreator/subcreator-runtime.json");
    }

    return candidates;
  }

  var homePath = Folder.home ? subcreator_trim_string(Folder.home.fsName) : "";
  if (homePath) {
    subcreator_runtime_push_unique(candidates, homePath + "/Library/Application Support/SubCreator/subcreator-runtime.json");
    subcreator_runtime_push_unique(candidates, homePath + "/Library/Application Support/PremiereSubCreator/subcreator-runtime.json");
  }

  return candidates;
}

function subcreator_read_runtime_config() {
  // // Read installer-generated runtime config to recover exact binary paths.
  if (typeof JSON === "undefined" || !JSON || typeof JSON.parse !== "function") {
    return null;
  }

  var candidatePaths = subcreator_resolve_runtime_config_paths();
  for (var i = 0; i < candidatePaths.length; i += 1) {
    var candidatePath = candidatePaths[i];
    var fileRef = new File(candidatePath);
    if (!fileRef.exists) {
      continue;
    }

    if (!fileRef.open("r")) {
      continue;
    }

    var payload = fileRef.read();
    fileRef.close();
    if (!payload || !subcreator_trim_string(payload)) {
      continue;
    }

    try {
      var parsed = JSON.parse(payload);
      if (parsed && typeof parsed === "object") {
        parsed.__sourcePath = candidatePath;
        return parsed;
      }
    } catch (error) {}
  }

  return null;
}

function subcreator_collect_runtime_path_hints(runtimeConfig) {
  // // Collect PATH additions from config + known defaults so Whisper can find ffmpeg.
  var hints = [];
  if (!runtimeConfig || typeof runtimeConfig !== "object") {
    return hints;
  }

  if (runtimeConfig.pathHints && typeof runtimeConfig.pathHints.length === "number") {
    for (var i = 0; i < runtimeConfig.pathHints.length; i += 1) {
      subcreator_runtime_push_unique(hints, runtimeConfig.pathHints[i]);
    }
  }

  subcreator_runtime_push_unique(hints, subcreator_runtime_dirname(runtimeConfig.whisperPath));
  subcreator_runtime_push_unique(hints, subcreator_runtime_dirname(runtimeConfig.pythonPath));
  subcreator_runtime_push_unique(hints, subcreator_runtime_dirname(runtimeConfig.ffmpegPath));

  if (subcreator_is_windows()) {
    subcreator_runtime_push_unique(hints, "C:/Program Files/ffmpeg/bin");
    subcreator_runtime_push_unique(hints, "C:/ffmpeg/bin");
    subcreator_runtime_push_unique(hints, "C:/Windows/System32");
  } else {
    subcreator_runtime_push_unique(hints, "/opt/homebrew/bin");
    subcreator_runtime_push_unique(hints, "/usr/local/bin");
    subcreator_runtime_push_unique(hints, "/usr/bin");
    subcreator_runtime_push_unique(hints, "/bin");
  }

  return hints;
}

function subcreator_build_runtime_env_prefix(runtimeConfig) {
  // // Build shell prefix that injects runtime PATH hints before Whisper command execution.
  var hints = subcreator_collect_runtime_path_hints(runtimeConfig);
  if (!hints.length) {
    return "";
  }

  if (subcreator_is_windows()) {
    var windowsHints = [];
    for (var i = 0; i < hints.length; i += 1) {
      windowsHints.push(String(hints[i] || "").replace(/\//g, "\\"));
    }
    return 'set "PATH=' + windowsHints.join(";") + ';%PATH%" && ';
  }

  return "PATH=" + subcreator_quote_posix(hints.join(":")) + ":$PATH ";
}

function subcreator_build_whisper_command(audioPath, outputDir, model, languageCode) {
  // // Build CLI command string for local Whisper execution.
  var runtimeConfig = subcreator_read_runtime_config();
  var pathPrefix = subcreator_build_runtime_env_prefix(runtimeConfig);
  var whisperBinary = "whisper";
  var usePythonModule = false;
  var pythonCommandText = "";
  if (runtimeConfig && typeof runtimeConfig === "object") {
    var configuredWhisperPath = subcreator_trim_string(runtimeConfig.whisperPath || "");
    var configuredPythonPath = subcreator_trim_string(runtimeConfig.pythonPath || "");
    var configuredPythonCommand = subcreator_trim_string(runtimeConfig.pythonCommand || "");

    if (configuredWhisperPath) {
      whisperBinary = configuredWhisperPath;
    } else if (configuredPythonPath) {
      whisperBinary = configuredPythonPath;
      usePythonModule = true;
    } else if (configuredPythonCommand) {
      pythonCommandText = configuredPythonCommand;
      usePythonModule = true;
    }
  }
  var modelArg = model && model.length > 0 ? model : "base";
  var languageArg = languageCode && languageCode.length > 0 ? languageCode : "";

  if (subcreator_is_windows()) {
    var launcherPrefix = "";
    if (usePythonModule && pythonCommandText) {
      launcherPrefix = pythonCommandText + " -m whisper ";
    } else if (usePythonModule) {
      launcherPrefix = subcreator_quote_cmd(whisperBinary) + " -m whisper ";
    } else {
      launcherPrefix = subcreator_quote_cmd(whisperBinary) + " ";
    }

    var cmd =
      pathPrefix +
      launcherPrefix +
      subcreator_quote_cmd(audioPath) +
      " --model " +
      subcreator_quote_cmd(modelArg) +
      " --output_format srt --output_dir " +
      subcreator_quote_cmd(outputDir) +
      " --fp16 False";

    if (languageArg && languageArg.toLowerCase() !== "auto") {
      cmd += " --language " + subcreator_quote_cmd(languageArg);
    }

    return cmd;
  }

  var launcher = "";
  if (usePythonModule && pythonCommandText) {
    launcher = pythonCommandText + " -m whisper ";
  } else if (usePythonModule) {
    launcher = subcreator_quote_posix(whisperBinary) + " -m whisper ";
  } else {
    launcher = subcreator_quote_posix(whisperBinary) + " ";
  }

  var shellCmd =
    pathPrefix +
    launcher +
    subcreator_quote_posix(audioPath) +
    " --model " +
    subcreator_quote_posix(modelArg) +
    " --output_format srt --output_dir " +
    subcreator_quote_posix(outputDir) +
    " --fp16 False";

  if (languageArg && languageArg.toLowerCase() !== "auto") {
    shellCmd += " --language " + subcreator_quote_posix(languageArg);
  }

  return shellCmd;
}

function subcreator_find_whisper_srt_file(tempFolder, baseName) {
  // // Resolve the SRT file created by Whisper in temporary output directory.
  var directPath = new File(tempFolder.fsName + "/" + baseName + ".srt");
  if (directPath.exists) {
    return directPath;
  }

  var files = tempFolder.getFiles("*.srt");
  for (var i = 0; i < files.length; i += 1) {
    var candidate = files[i];
    if (candidate instanceof File) {
      var candidateName = String(candidate.name || "").toLowerCase();
      if (candidateName.indexOf(String(baseName).toLowerCase()) === 0) {
        return candidate;
      }
    }
  }

  return null;
}

function subcreator_transcribe_whisper(payloadEncoded) {
  // // Run local Whisper CLI and return generated SRT text.
  try {
    var payloadText = subcreator_decode_payload(payloadEncoded || "");
    var payload = JSON.parse(payloadText);

    var audioPath = String(payload.audioPath || "");
    if (!audioPath) {
      return subcreator_error("Missing audioPath for Whisper transcription.");
    }

    var audioFile = new File(audioPath);
    if (!audioFile.exists) {
      return subcreator_error("Audio file not found: " + audioPath);
    }

    var tempFolder = new Folder(Folder.temp.fsName + "/SubCreatorWhisper");
    if (!tempFolder.exists) {
      tempFolder.create();
    }

    var model = String(payload.model || "base");
    var languageCode = String(payload.languageCode || "");
    var command = subcreator_build_whisper_command(audioFile.fsName, tempFolder.fsName, model, languageCode);
    if (typeof system === "undefined" || !system || typeof system.callSystem !== "function") {
      return subcreator_error("Host system.callSystem indisponible. Active le mode Node CEP pour Whisper local.");
    }

    var commandOutput = system.callSystem(command);

    var baseName = String(audioFile.name || "").replace(/\.[^\.]+$/, "");
    var srtFile = subcreator_find_whisper_srt_file(tempFolder, baseName);
    if (!srtFile || !srtFile.exists) {
      return subcreator_error(
        "Whisper did not produce an SRT file. Ensure Whisper CLI is installed and available in PATH. Output: " + commandOutput
      );
    }

    var srtText = subcreator_read_file_text(srtFile);
    if (!srtText || srtText.length === 0) {
      return subcreator_error("Whisper produced an empty SRT file: " + srtFile.fsName);
    }

    return subcreator_ok({
      srtText: srtText,
      model: model,
      audioPath: audioFile.fsName,
      commandOutput: commandOutput
    });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_to_seconds(value) {
  // // Convert unknown time objects (Time/ticks/numeric) to seconds.
  if (value === undefined || value === null) {
    return NaN;
  }

  if (typeof value === "number") {
    return Number(value);
  }

  if (typeof value === "string") {
    return Number(value);
  }

  if (typeof value.seconds !== "undefined") {
    return Number(value.seconds);
  }

  if (typeof value.ticks !== "undefined") {
    return Number(value.ticks) / 254016000000;
  }

  return NaN;
}

function subcreator_collection_to_array(collection) {
  // // Convert ExtendScript collections and JS arrays into simple arrays.
  var result = [];
  if (!collection) {
    return result;
  }

  if (typeof collection.length === "number") {
    for (var i = 0; i < collection.length; i += 1) {
      result.push(collection[i]);
    }
    return result;
  }

  if (typeof collection.numItems === "number") {
    for (var j = 0; j < collection.numItems; j += 1) {
      result.push(collection[j]);
    }
    return result;
  }

  return result;
}

function subcreator_get_selected_track_items(sequence) {
  // // Read current timeline selection and normalize to a plain array.
  if (!sequence || typeof sequence.getSelection !== "function") {
    return [];
  }

  try {
    return subcreator_collection_to_array(sequence.getSelection());
  } catch (error) {
    return [];
  }
}

function subcreator_get_mogrt_component_from_track_item(trackItem) {
  // // Resolve Essential Graphics component from a track item when available.
  if (!trackItem) {
    return null;
  }

  var component = null;
  try {
    if (typeof trackItem.getMGTComponent === "function") {
      component = trackItem.getMGTComponent();
    }
  } catch (mgtError) {}

  if (!component && trackItem.components && trackItem.components.numItems > 0) {
    component = trackItem.components[0];
  }

  if (!component || !component.properties || typeof component.properties.numItems !== "number") {
    return null;
  }

  return component;
}

function subcreator_detect_visual_property_type(rawValue) {
  // // Categorize host property values so panel can render matching input controls.
  if (typeof rawValue === "number") {
    return "number";
  }
  if (typeof rawValue === "boolean") {
    return "boolean";
  }
  if (typeof rawValue === "string") {
    return "string";
  }
  return "json";
}

function subcreator_visual_is_guid_list_string(value) {
  // // Detect Premiere internal GUID lists used by synthetic group metadata payloads.
  var text = subcreator_trim_string(String(value || ""));
  if (!text) {
    return false;
  }

  return /^(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12};)+$/i.test(text);
}

function subcreator_visual_is_group_metadata_value(rawValue) {
  // // Identify container-only values that should not be shown as editable controls.
  if (rawValue === undefined || rawValue === null) {
    return true;
  }

  if (typeof rawValue === "string") {
    var normalized = subcreator_trim_string(rawValue);
    if (!normalized) {
      return true;
    }

    if (subcreator_visual_is_guid_list_string(normalized)) {
      return true;
    }
  }

  if (typeof rawValue === "object") {
    if (typeof rawValue.length === "number" && rawValue.length < 1) {
      return true;
    }
  }

  return false;
}

function subcreator_visual_to_number(rawValue) {
  // // Convert unknown scalar values to number while preserving NaN on failure.
  var parsed = Number(rawValue);
  return isNaN(parsed) ? NaN : parsed;
}

function subcreator_visual_is_numeric_string(rawValue) {
  // // Detect numeric strings so slider/select controls can stay numeric instead of color/string.
  if (typeof rawValue !== "string") {
    return false;
  }

  return /^-?\d+(?:[.,]\d+)?$/.test(subcreator_trim_string(rawValue));
}

function subcreator_visual_clamp(value, minValue, maxValue) {
  // // Clamp numeric values so color/range conversions stay within valid bounds.
  var numericValue = Number(value);
  if (isNaN(numericValue)) {
    return Number(minValue);
  }
  if (numericValue < minValue) {
    return Number(minValue);
  }
  if (numericValue > maxValue) {
    return Number(maxValue);
  }
  return numericValue;
}

function subcreator_visual_channel_to_hex(value) {
  // // Convert one RGB channel to a 2-char hexadecimal value.
  var clamped = Math.round(subcreator_visual_clamp(value, 0, 255));
  var hex = clamped.toString(16);
  return hex.length < 2 ? "0" + hex : hex;
}

function subcreator_visual_rgb_to_hex(red, green, blue) {
  // // Build CSS hex color from RGB channels.
  return (
    "#" +
    subcreator_visual_channel_to_hex(red) +
    subcreator_visual_channel_to_hex(green) +
    subcreator_visual_channel_to_hex(blue)
  );
}

function subcreator_visual_extract_rgb_triplet(rawRed, rawGreen, rawBlue) {
  // // Normalize RGB channels from 0..1 or 0..255 formats.
  var red = subcreator_visual_to_number(rawRed);
  var green = subcreator_visual_to_number(rawGreen);
  var blue = subcreator_visual_to_number(rawBlue);
  if (isNaN(red) || isNaN(green) || isNaN(blue)) {
    return null;
  }

  var useUnitScale = red <= 1 && green <= 1 && blue <= 1;
  return {
    red: useUnitScale ? subcreator_visual_clamp(red * 255, 0, 255) : subcreator_visual_clamp(red, 0, 255),
    green: useUnitScale ? subcreator_visual_clamp(green * 255, 0, 255) : subcreator_visual_clamp(green, 0, 255),
    blue: useUnitScale ? subcreator_visual_clamp(blue * 255, 0, 255) : subcreator_visual_clamp(blue, 0, 255),
    unitScale: useUnitScale
  };
}

function subcreator_visual_detect_color_array_layout(rawArray) {
  // // Detect whether a color array uses RGB, RGBA, or ARGB channel order.
  if (!rawArray || typeof rawArray.length !== "number") {
    return "unknown";
  }

  var size = Number(rawArray.length || 0);
  if (size < 3) {
    return "unknown";
  }

  if (size === 3) {
    return "rgb";
  }

  var alpha = subcreator_visual_to_number(rawArray[0]);
  var red = subcreator_visual_to_number(rawArray[1]);
  var green = subcreator_visual_to_number(rawArray[2]);
  var blue = subcreator_visual_to_number(rawArray[3]);

  if (isNaN(alpha) || isNaN(red) || isNaN(green) || isNaN(blue)) {
    return "unknown";
  }

  var firstIsUnitAlpha = alpha >= 0 && alpha <= 1.0001 && (red > 1 || green > 1 || blue > 1);
  if (firstIsUnitAlpha) {
    return "argb";
  }

  var lastIsUnitAlpha = blue >= 0 && blue <= 1.0001 && (alpha > 1 || red > 1 || green > 1);
  if (lastIsUnitAlpha) {
    return "rgba";
  }

  var firstLooksLikeAlphaMarker = alpha === 255 || alpha === 1 || alpha === 0;
  var lastLooksLikeAlphaMarker = blue === 255 || blue === 1 || blue === 0;

  if (firstLooksLikeAlphaMarker && lastLooksLikeAlphaMarker) {
    // // Premiere color arrays are frequently `[A,R,G,B]`; prefer ARGB when both edges look like alpha markers.
    return "argb";
  }

  if (firstLooksLikeAlphaMarker && !lastLooksLikeAlphaMarker) {
    return "argb";
  }

  if (lastLooksLikeAlphaMarker && !firstLooksLikeAlphaMarker) {
    return "rgba";
  }

  if (alpha === 255 || alpha === 1 || alpha === 0) {
    return "argb";
  }

  if (blue === 255 || blue === 1) {
    return "rgba";
  }

  return "rgb";
}

function subcreator_visual_get_color_layout_hint(displayName, groupPath) {
  // // Keep layout hints neutral; runtime calibration and array detection decide final mapping.
  return "";
}

var subcreator_visual_color_read_layout_cache = {};
var subcreator_visual_color_write_layout_cache = {};

function subcreator_visual_get_color_cache_key(displayName) {
  // // Keep calibration cache scoped by color control display name.
  return subcreator_trim_string(String(displayName || "")).toLowerCase();
}

function subcreator_visual_get_cached_color_layout(displayName, mode) {
  // // Read in-memory layout calibration for color controls (read and write kept separate).
  var cacheKey = subcreator_visual_get_color_cache_key(displayName);
  if (!cacheKey) {
    return "";
  }
  var cacheMode = subcreator_trim_string(String(mode || "read")).toLowerCase();
  var sourceCache = cacheMode === "write" ? subcreator_visual_color_write_layout_cache : subcreator_visual_color_read_layout_cache;
  return String(sourceCache[cacheKey] || "");
}

function subcreator_visual_set_cached_color_layout(displayName, layout, mode) {
  // // Persist successful color layout calibrations for current CEP host session.
  var cacheKey = subcreator_visual_get_color_cache_key(displayName);
  if (!cacheKey) {
    return;
  }

  var normalizedLayout = subcreator_trim_string(String(layout || "")).toLowerCase();
  if (!normalizedLayout) {
    return;
  }

  var cacheMode = subcreator_trim_string(String(mode || "read")).toLowerCase();
  if (cacheMode === "write") {
    subcreator_visual_color_write_layout_cache[cacheKey] = normalizedLayout;
    return;
  }

  subcreator_visual_color_read_layout_cache[cacheKey] = normalizedLayout;
}

function subcreator_visual_build_color_layout_candidates(displayName, groupPath, mode) {
  // // Build ordered layout candidates (cache first, then hint, then fallbacks) for auto calibration.
  var candidates = [];
  var cacheMode = subcreator_trim_string(String(mode || "read")).toLowerCase();

  function pushCandidate(layout) {
    var normalizedLayout = subcreator_trim_string(String(layout || "")).toLowerCase();
    if (!normalizedLayout) {
      return;
    }
    for (var index = 0; index < candidates.length; index += 1) {
      if (candidates[index] === normalizedLayout) {
        return;
      }
    }
    candidates.push(normalizedLayout);
  }

  if (cacheMode === "write") {
    // // Prefer previously calibrated write layout, then currently known read layout as secondary hint.
    pushCandidate(subcreator_visual_get_cached_color_layout(displayName, "write"));
    pushCandidate(subcreator_visual_get_cached_color_layout(displayName, "read"));
    pushCandidate(subcreator_visual_get_color_layout_hint(displayName, groupPath));
  } else {
    // // Prefer previously calibrated read layout when decoding getColorValue payloads.
    pushCandidate(subcreator_visual_get_cached_color_layout(displayName, "read"));
    pushCandidate(subcreator_visual_get_color_layout_hint(displayName, groupPath));
    pushCandidate(subcreator_visual_get_cached_color_layout(displayName, "write"));
  }
  pushCandidate("argb");
  pushCandidate("rgba");
  pushCandidate("bgra");
  pushCandidate("abgr");
  pushCandidate("rgb");
  return candidates;
}

function subcreator_visual_color_layout_indices(layout, size) {
  // // Resolve channel index map for supported array color layouts.
  var normalizedLayout = String(layout || "").toLowerCase();
  if (normalizedLayout === "argb") {
    return { red: 1, green: 2, blue: 3, alpha: 0 };
  }
  if (normalizedLayout === "bgra") {
    return { red: 2, green: 1, blue: 0, alpha: 3 };
  }
  if (normalizedLayout === "abgr") {
    return { red: 3, green: 2, blue: 1, alpha: 0 };
  }
  if (normalizedLayout === "rgba") {
    return { red: 0, green: 1, blue: 2, alpha: 3 };
  }

  if (Number(size || 0) >= 4) {
    return { red: 0, green: 1, blue: 2, alpha: 3 };
  }

  return { red: 0, green: 1, blue: 2, alpha: -1 };
}

function subcreator_visual_extract_rgb_from_array_with_layout(rawArray, layout) {
  // // Extract RGB from one explicit array layout.
  if (!rawArray || typeof rawArray.length !== "number" || rawArray.length < 3) {
    return null;
  }

  var indices = subcreator_visual_color_layout_indices(layout, rawArray.length);
  return subcreator_visual_extract_rgb_triplet(rawArray[indices.red], rawArray[indices.green], rawArray[indices.blue]);
}

function subcreator_visual_is_alpha_first_color_array(rawArray) {
  // // Backward-compatible helper for existing call sites that need ARGB detection.
  return subcreator_visual_detect_color_array_layout(rawArray) === "argb";
}

function subcreator_visual_extract_rgb_from_packed_number(rawNumber) {
  // // Decode packed numeric color payloads used by some Essential Graphics controls.
  var numericColor = Math.floor(Math.abs(Number(rawNumber)));
  if (isNaN(numericColor)) {
    return null;
  }

  if (numericColor <= 1) {
    var grayUnit = subcreator_visual_clamp(numericColor * 255, 0, 255);
    return {
      red: grayUnit,
      green: grayUnit,
      blue: grayUnit,
      unitScale: true
    };
  }

  // // 64-bit packed shape stores color words in Blue/Red/Green order in many MOGRT controls.
  var rawHex = numericColor.toString(16);
  if (rawHex.length > 8) {
    while (rawHex.length < 16) {
      rawHex = "0" + rawHex;
    }
    if (rawHex.length >= 16) {
      var r16 = parseInt(rawHex.substring(0, 4), 16);
      var g16 = parseInt(rawHex.substring(4, 8), 16);
      var b16 = parseInt(rawHex.substring(8, 12), 16);
      if (!isNaN(r16) && !isNaN(g16) && !isNaN(b16)) {
        var channelBlue = r16 > 255 ? Math.floor(r16 / 256) : r16;
        var channelRed = g16 > 255 ? Math.floor(g16 / 256) : g16;
        var channelGreen = b16 > 255 ? Math.floor(b16 / 256) : b16;
        return {
          red: channelRed,
          green: channelGreen,
          blue: channelBlue,
          unitScale: false
        };
      }
    }
  }

  if (numericColor > 4294967295) {
    return null;
  }

  // // Compact packed numbers are also interpreted as Blue/Red/Green channel order.
  var packed = numericColor % 16777216;
  var packedBlue = Math.floor(packed / 65536) % 256;
  var packedRed = Math.floor(packed / 256) % 256;
  var packedGreen = packed % 256;
  return {
    red: packedRed,
    green: packedGreen,
    blue: packedBlue,
    unitScale: false
  };
}

function subcreator_visual_extract_rgb_from_value(rawValue, allowPackedNumbers, preferredArrayLayout) {
  // // Read RGB channels from known color payload shapes.
  if (rawValue === undefined || rawValue === null) {
    return null;
  }

  if (typeof rawValue === "number") {
    if (!allowPackedNumbers) {
      return null;
    }
    return subcreator_visual_extract_rgb_from_packed_number(rawValue);
  }

  if (typeof rawValue === "string") {
    var text = subcreator_trim_string(String(rawValue || ""));
    if (/^#[0-9a-f]{6}$/i.test(text)) {
      return {
        red: parseInt(text.substring(1, 3), 16),
        green: parseInt(text.substring(3, 5), 16),
        blue: parseInt(text.substring(5, 7), 16),
        unitScale: false
      };
    }

    if (/^#[0-9a-f]{3}$/i.test(text)) {
      return {
        red: parseInt(text.charAt(1) + text.charAt(1), 16),
        green: parseInt(text.charAt(2) + text.charAt(2), 16),
        blue: parseInt(text.charAt(3) + text.charAt(3), 16),
        unitScale: false
      };
    }

    if (/^\d+$/.test(text) && allowPackedNumbers) {
      var asNumber = Number(text);
      if (!isNaN(asNumber)) {
        return subcreator_visual_extract_rgb_from_packed_number(asNumber);
      }
      return null;
    }

    if (text.indexOf("{") !== -1 || text.indexOf("[") !== -1) {
      try {
        var parsed = JSON.parse(text);
        return subcreator_visual_extract_rgb_from_value(parsed, allowPackedNumbers, preferredArrayLayout);
      } catch (jsonError) {}
    }

    return null;
  }

  if (typeof rawValue === "object") {
    if (typeof rawValue.length === "number" && rawValue.length >= 3) {
      if (preferredArrayLayout) {
        var fromPreferredArray = subcreator_visual_extract_rgb_from_array_with_layout(rawValue, preferredArrayLayout);
        if (fromPreferredArray) {
          return fromPreferredArray;
        }
      }

      var arrayLayout = subcreator_visual_detect_color_array_layout(rawValue);
      var fromDetectedLayout = subcreator_visual_extract_rgb_from_array_with_layout(rawValue, arrayLayout);
      if (fromDetectedLayout) {
        return fromDetectedLayout;
      }

      var fromArray = subcreator_visual_extract_rgb_triplet(rawValue[0], rawValue[1], rawValue[2]);
      if (fromArray) {
        return fromArray;
      }
    }

    if (
      typeof rawValue.red !== "undefined" &&
      typeof rawValue.green !== "undefined" &&
      typeof rawValue.blue !== "undefined"
    ) {
      var fromRgbKeys = subcreator_visual_extract_rgb_triplet(rawValue.red, rawValue.green, rawValue.blue);
      if (fromRgbKeys) {
        return fromRgbKeys;
      }
    }

    if (typeof rawValue.r !== "undefined" && typeof rawValue.g !== "undefined" && typeof rawValue.b !== "undefined") {
      var fromShortKeys = subcreator_visual_extract_rgb_triplet(rawValue.r, rawValue.g, rawValue.b);
      if (fromShortKeys) {
        return fromShortKeys;
      }
    }

    if (rawValue.color && typeof rawValue.color === "object") {
      var fromNestedColor = subcreator_visual_extract_rgb_from_value(rawValue.color, allowPackedNumbers, preferredArrayLayout);
      if (fromNestedColor) {
        return fromNestedColor;
      }
    }
  }

  return null;
}

function subcreator_visual_extract_color_hex(rawValue, allowPackedNumbers, preferredArrayLayout) {
  // // Convert color payloads to CSS hex for panel color inputs.
  var rgb = subcreator_visual_extract_rgb_from_value(rawValue, allowPackedNumbers === true, preferredArrayLayout);
  if (!rgb) {
    return "";
  }
  return subcreator_visual_rgb_to_hex(rgb.red, rgb.green, rgb.blue);
}

function subcreator_visual_try_read_property_color_hex(property, rawFallbackValue, allowPackedNumbers, preferredArrayLayout) {
  // // Read color directly from color-capable APIs when available to avoid numeric payload ambiguity.
  if (property && typeof property.getColorValue === "function") {
    try {
      var colorValue = property.getColorValue();
      var fromColorMethod = subcreator_visual_extract_color_hex(colorValue, true, preferredArrayLayout);
      if (fromColorMethod) {
        return fromColorMethod;
      }
    } catch (colorReadError) {}
  }

  return subcreator_visual_extract_color_hex(rawFallbackValue, allowPackedNumbers, preferredArrayLayout);
}

function subcreator_visual_is_likely_color_payload(rawValue) {
  // // Detect payload shapes that genuinely look like color values.
  if (rawValue === undefined || rawValue === null) {
    return false;
  }

  if (typeof rawValue === "string") {
    var text = subcreator_trim_string(rawValue);
    if (/^#[0-9a-f]{3,6}$/i.test(text)) {
      return true;
    }
    if (text.indexOf("{") !== -1 || text.indexOf("[") !== -1) {
      var parsed = null;
      try {
        parsed = JSON.parse(text);
      } catch (parseError) {
        parsed = null;
      }
      if (parsed) {
        return subcreator_visual_is_likely_color_payload(parsed);
      }
    }
    return false;
  }

  if (typeof rawValue === "object") {
    if (typeof rawValue.length === "number" && rawValue.length >= 3) {
      return true;
    }
    if (
      typeof rawValue.red !== "undefined" ||
      typeof rawValue.green !== "undefined" ||
      typeof rawValue.blue !== "undefined" ||
      typeof rawValue.r !== "undefined" ||
      typeof rawValue.g !== "undefined" ||
      typeof rawValue.b !== "undefined"
    ) {
      return true;
    }
  }

  return false;
}

function subcreator_visual_is_color_label(displayName) {
  // // Detect color-like labels so panel can render native color pickers.
  var key = String(displayName || "").toLowerCase();
  if (
    key.indexOf("width") !== -1 ||
    key.indexOf("size") !== -1 ||
    key.indexOf("amount") !== -1 ||
    key.indexOf("opacity") !== -1 ||
    key.indexOf("position") !== -1 ||
    key.indexOf("offset") !== -1
  ) {
    return false;
  }

  return (
    key.indexOf("color") !== -1 ||
    key.indexOf("couleur") !== -1 ||
    key.indexOf("fill") !== -1 ||
    key.indexOf("stroke") !== -1 ||
    key.indexOf("outline") !== -1 ||
    key.indexOf("tint") !== -1 ||
    key.indexOf("shadow") !== -1
  );
}

function subcreator_visual_group_suggests_color(groupPath) {
  // // Detect color-oriented groups so numeric packed colors can be shown as color pickers.
  var key = String(groupPath || "").toLowerCase();
  return (
    key.indexOf("fill") !== -1 ||
    key.indexOf("stroke") !== -1 ||
    key.indexOf("highlight") !== -1 ||
    key.indexOf("color") !== -1 ||
    key.indexOf("couleur") !== -1 ||
    key.indexOf("outline") !== -1 ||
    key.indexOf("shadow") !== -1
  );
}

function subcreator_visual_is_discrete_numeric_label(displayName) {
  // // Detect numeric menu-like fields where a raw number input is safer than slider.
  var key = String(displayName || "").toLowerCase();
  return (
    key.indexOf("mode") !== -1 ||
    key.indexOf("type") !== -1 ||
    key.indexOf("style") !== -1 ||
    key.indexOf("preset") !== -1 ||
    key.indexOf("family") !== -1 ||
    key.indexOf("based on") !== -1 ||
    key.indexOf("align") !== -1 ||
    key.indexOf("justif") !== -1 ||
    key.indexOf("case") !== -1
  );
}

function subcreator_visual_try_read_number_member(property, key) {
  // // Read numeric properties/methods from host controls when available.
  if (!property || !key) {
    return NaN;
  }

  var value = NaN;

  try {
    if (typeof property[key] === "function") {
      value = Number(property[key]());
    } else if (typeof property[key] !== "undefined") {
      value = Number(property[key]);
    }
  } catch (error) {
    value = NaN;
  }

  return isNaN(value) ? NaN : value;
}

function subcreator_visual_guess_numeric_range(displayName, rawValue) {
  // // Guess ergonomic slider ranges when host metadata does not expose min/max.
  var key = String(displayName || "").toLowerCase();
  var numericValue = subcreator_visual_to_number(rawValue);

  if (isNaN(numericValue)) {
    numericValue = 0;
  }

  if (key.indexOf("opacity") !== -1 || key.indexOf("opacite") !== -1) {
    return { minValue: 0, maxValue: 100, stepValue: 1 };
  }

  if (
    key === "x" ||
    key === "y" ||
    key.indexOf("anchor") !== -1 ||
    key.indexOf("start") !== -1 ||
    key.indexOf("end") !== -1 ||
    key.indexOf("progress") !== -1 ||
    key.indexOf("delay") !== -1
  ) {
    return { minValue: 0, maxValue: 100, stepValue: 1 };
  }

  if (key.indexOf("offset") !== -1 || key.indexOf("position") !== -1) {
    return { minValue: -100, maxValue: 100, stepValue: 1 };
  }

  if (key.indexOf("scale") !== -1 || key.indexOf("size") !== -1 || key.indexOf("taille") !== -1) {
    return { minValue: 0, maxValue: 400, stepValue: 1 };
  }

  if (key.indexOf("rotation") !== -1 || key.indexOf("angle") !== -1) {
    return { minValue: -360, maxValue: 360, stepValue: 1 };
  }

  if (key.indexOf("line") !== -1 && key.indexOf("max") !== -1) {
    return { minValue: 1, maxValue: 6, stepValue: 1 };
  }

  if (key.indexOf("character") !== -1 || key.indexOf("chars") !== -1 || key.indexOf("letter") !== -1) {
    return { minValue: 4, maxValue: 120, stepValue: 1 };
  }

  var delta = Math.max(Math.abs(numericValue), 50);
  return {
    minValue: Math.floor(numericValue - delta),
    maxValue: Math.ceil(numericValue + delta),
    stepValue: Number(Math.abs(numericValue % 1) > 0 ? 0.1 : 1)
  };
}

function subcreator_visual_read_numeric_range(property, displayName, rawValue) {
  // // Resolve numeric ranges using host hints first, then name-based heuristics.
  var minCandidates = ["getMinValue", "getMinimum", "getMin", "minValue", "minimum", "min"];
  var maxCandidates = ["getMaxValue", "getMaximum", "getMax", "maxValue", "maximum", "max"];
  var stepCandidates = ["getStepValue", "getStep", "stepValue", "step"];

  var minValue = NaN;
  var maxValue = NaN;
  var stepValue = NaN;
  var i = 0;

  for (i = 0; i < minCandidates.length; i += 1) {
    minValue = subcreator_visual_try_read_number_member(property, minCandidates[i]);
    if (!isNaN(minValue)) {
      break;
    }
  }

  for (i = 0; i < maxCandidates.length; i += 1) {
    maxValue = subcreator_visual_try_read_number_member(property, maxCandidates[i]);
    if (!isNaN(maxValue)) {
      break;
    }
  }

  for (i = 0; i < stepCandidates.length; i += 1) {
    stepValue = subcreator_visual_try_read_number_member(property, stepCandidates[i]);
    if (!isNaN(stepValue)) {
      break;
    }
  }

  var guessed = subcreator_visual_guess_numeric_range(displayName, rawValue);

  if (isNaN(minValue)) {
    minValue = guessed.minValue;
  }

  if (isNaN(maxValue)) {
    maxValue = guessed.maxValue;
  }

  if (isNaN(stepValue) || stepValue <= 0) {
    stepValue = guessed.stepValue;
  }

  if (maxValue <= minValue) {
    maxValue = minValue + Math.max(1, Number(stepValue || 1));
  }

  return {
    minValue: minValue,
    maxValue: maxValue,
    stepValue: stepValue
  };
}

function subcreator_visual_build_select_options(displayName, rawValue) {
  // // Build known dropdown option sets for menu-like numeric controls.
  var key = String(displayName || "").toLowerCase();
  var numericValue = Number(rawValue);
  if (isNaN(numericValue)) {
    return null;
  }

  function buildLabeledRange(startValue, labels) {
    var options = [];
    for (var optionIndex = 0; optionIndex < labels.length; optionIndex += 1) {
      options.push({
        value: startValue + optionIndex,
        label: labels[optionIndex]
      });
    }
    return options;
  }

  if (key.indexOf("based on") !== -1 || key.indexOf("highlight based on") !== -1) {
    if (numericValue >= 0 && numericValue <= 1) {
      return buildLabeledRange(0, ["Words", "Lines"]);
    }
    if (numericValue >= 1 && numericValue <= 2) {
      return buildLabeledRange(1, ["Words", "Lines"]);
    }
  }

  if (key.indexOf("paragraph") !== -1 || key.indexOf("align") !== -1 || key.indexOf("alignment") !== -1) {
    if (numericValue >= 0 && numericValue <= 3) {
      return buildLabeledRange(0, ["Left", "Center", "Right", "Justify"]);
    }
    if (numericValue >= 1 && numericValue <= 4) {
      return buildLabeledRange(1, ["Left", "Center", "Right", "Justify"]);
    }
  }

  return null;
}

function subcreator_visual_extract_numeric_vector(rawValue) {
  // // Extract compact numeric vectors used by offset/size controls.
  if (!rawValue || typeof rawValue !== "object" || typeof rawValue.length !== "number") {
    return null;
  }

  var size = Number(rawValue.length || 0);
  if (size < 2 || size > 4) {
    return null;
  }

  var values = [];
  for (var index = 0; index < size; index += 1) {
    var numericValue = Number(rawValue[index]);
    if (isNaN(numericValue)) {
      return null;
    }
    values.push(numericValue);
  }

  return values;
}

function subcreator_visual_read_sequence_dimensions() {
  // // Read active sequence dimensions for converting internal vector units to UI-friendly values.
  var width = 1920;
  var height = 1080;

  try {
    if (app && app.project && app.project.activeSequence) {
      var sequence = app.project.activeSequence;

      if (typeof sequence.frameSizeHorizontal !== "undefined") {
        var frameWidth = Number(sequence.frameSizeHorizontal);
        if (!isNaN(frameWidth) && frameWidth > 0) {
          width = frameWidth;
        }
      }

      if (typeof sequence.frameSizeVertical !== "undefined") {
        var frameHeight = Number(sequence.frameSizeVertical);
        if (!isNaN(frameHeight) && frameHeight > 0) {
          height = frameHeight;
        }
      }

      if (typeof sequence.getSettings === "function") {
        var settings = sequence.getSettings();
        if (settings) {
          var settingsWidth = Number(settings.videoFrameWidth || settings.frameWidth || settings.width);
          if (!isNaN(settingsWidth) && settingsWidth > 0) {
            width = settingsWidth;
          }

          var settingsHeight = Number(settings.videoFrameHeight || settings.frameHeight || settings.height);
          if (!isNaN(settingsHeight) && settingsHeight > 0) {
            height = settingsHeight;
          }
        }
      }
    }
  } catch (readSequenceError) {}

  try {
    if (typeof app.enableQE === "function") {
      app.enableQE();
    }
    if (typeof qe !== "undefined" && qe.project && typeof qe.project.getActiveSequence === "function") {
      var qeSequence = qe.project.getActiveSequence();
      if (qeSequence) {
        var qeWidth = Number(qeSequence.videoFrameWidth);
        if (!isNaN(qeWidth) && qeWidth > 0) {
          width = qeWidth;
        }

        var qeHeight = Number(qeSequence.videoFrameHeight);
        if (!isNaN(qeHeight) && qeHeight > 0) {
          height = qeHeight;
        }

        if (qeSequence.sequence) {
          var nestedWidth = Number(qeSequence.sequence.videoFrameWidth);
          if (!isNaN(nestedWidth) && nestedWidth > 0) {
            width = nestedWidth;
          }
          var nestedHeight = Number(qeSequence.sequence.videoFrameHeight);
          if (!isNaN(nestedHeight) && nestedHeight > 0) {
            height = nestedHeight;
          }
        }
      }
    }
  } catch (readQeError) {}

  return {
    width: width,
    height: height
  };
}

var subcreator_visual_group_sequence_axis_preferences = {};

function subcreator_visual_reset_group_sequence_axis_preferences() {
  // // Reset per-group vector scaling hints before reading a new selection.
  subcreator_visual_group_sequence_axis_preferences = {};
}

function subcreator_visual_group_sequence_axis_key(groupPath) {
  // // Normalize group key used for cross-property scale inference.
  return subcreator_trim_string(String(groupPath || "")).toLowerCase();
}

function subcreator_visual_mark_group_sequence_axis(groupPath) {
  // // Remember that a group uses sequence-axis normalized units (for Position/Scale consistency).
  var key = subcreator_visual_group_sequence_axis_key(groupPath);
  if (!key) {
    return;
  }
  subcreator_visual_group_sequence_axis_preferences[key] = true;
}

function subcreator_visual_group_prefers_sequence_axis(groupPath) {
  // // Check whether previous properties in the group proved sequence-axis normalized behavior.
  var key = subcreator_visual_group_sequence_axis_key(groupPath);
  if (!key) {
    return false;
  }
  return !!subcreator_visual_group_sequence_axis_preferences[key];
}

function subcreator_visual_detect_vector_mode(displayName, groupPath) {
  // // Detect vector unit convention so panel can show human-friendly values.
  var displayKey = String(displayName || "").toLowerCase();
  var groupKey = String(groupPath || "").toLowerCase();

  if (displayKey.indexOf("size") !== -1 || displayKey.indexOf("scale") !== -1) {
    return "size_percent";
  }

  if (displayKey.indexOf("offset") !== -1 || displayKey.indexOf("position") !== -1) {
    return "offset_scaled";
  }

  var key = groupKey + " " + displayKey;

  if (key.indexOf("offset") !== -1 || key.indexOf("position") !== -1) {
    return "offset_scaled";
  }

  if (key.indexOf("size") !== -1 || key.indexOf("scale") !== -1) {
    return "size_percent";
  }

  return "raw";
}

function subcreator_visual_vector_looks_normalized_position(vectorValues) {
  // // Detect normalized position vectors (0..1-ish) that should be displayed in sequence pixels.
  if (!vectorValues || vectorValues.length < 2) {
    return false;
  }

  var x = Number(vectorValues[0]);
  var y = Number(vectorValues[1]);
  if (isNaN(x) || isNaN(y)) {
    return false;
  }

  return x >= -0.2 && x <= 1.2 && y >= -0.2 && y <= 1.2;
}

function subcreator_visual_score_vector_candidate(panelValues, minPreferred, maxPreferred, idealValue) {
  // // Score candidate panel-unit vectors and keep values in practical edit ranges.
  if (!panelValues || !panelValues.length) {
    return 999999;
  }

  var score = 0;
  for (var index = 0; index < panelValues.length; index += 1) {
    var value = Number(panelValues[index]);
    if (isNaN(value)) {
      score += 10000;
      continue;
    }

    var absValue = Math.abs(value);
    if (absValue > maxPreferred * 20) {
      score += 500;
    } else if (absValue > maxPreferred * 3) {
      score += 60;
    } else if (absValue > maxPreferred) {
      score += 20;
    }

    if (absValue < minPreferred) {
      score += 15;
    }

    score += Math.abs(absValue - idealValue) / Math.max(idealValue, 1) * 0.5;
  }

  return score;
}

function subcreator_visual_apply_vector_scale(vectorValues, scales) {
  // // Apply per-component scalar conversion.
  var output = [];
  for (var index = 0; index < vectorValues.length; index += 1) {
    var value = Number(vectorValues[index]);
    var scale = Number(scales[index]);
    if (isNaN(value) || isNaN(scale)) {
      output.push(0);
      continue;
    }
    output.push(value * scale);
  }
  return output;
}

function subcreator_visual_choose_vector_scale(displayName, groupPath, vectorValues, sequenceSize) {
  // // Infer best per-axis conversion scales for vector values (offset/size/raw).
  var width = Math.max(Number(sequenceSize && sequenceSize.width) || 1920, 1);
  var height = Math.max(Number(sequenceSize && sequenceSize.height) || 1080, 1);
  var vectorMode = subcreator_visual_detect_vector_mode(displayName, groupPath);
  var displayKey = String(displayName || "").toLowerCase();

  var candidates = [];
  if (vectorMode === "offset_scaled") {
    var looksLikePosition = displayKey.indexOf("position") !== -1;
    if (looksLikePosition && subcreator_visual_vector_looks_normalized_position(vectorValues)) {
      // // Position controls often report normalized coordinates; expose them as absolute sequence pixels in panel.
      subcreator_visual_mark_group_sequence_axis(groupPath);
      var normalizedScale = [width, height, 1, 1];
      var normalizedFinalScales = [];
      for (var normalizedIndex = 0; normalizedIndex < vectorValues.length; normalizedIndex += 1) {
        normalizedFinalScales.push(Number(normalizedScale[normalizedIndex] || 1));
      }
      return {
        mode: vectorMode,
        scale: normalizedFinalScales,
        candidateId: "position_normalized_axis",
        score: 0
      };
    }

    candidates.push({ id: "offset_raw", scales: [1, 1, 1, 1], minPreferred: 0.05, maxPreferred: 200, idealValue: 35 });
    candidates.push({
      id: "offset_div_axis",
      scales: [1 / width, 1 / height, 1, 1],
      minPreferred: 0.05,
      maxPreferred: 200,
      idealValue: 35
    });
    candidates.push({
      id: "offset_mul_axis",
      scales: [width, height, 1, 1],
      minPreferred: 0.05,
      maxPreferred: 200,
      idealValue: 35
    });
  } else if (vectorMode === "size_percent") {
    if (
      displayKey.indexOf("scale") !== -1 &&
      subcreator_visual_group_prefers_sequence_axis(groupPath) &&
      subcreator_visual_vector_looks_normalized_position(vectorValues)
    ) {
      // // Keep Scale consistent with Position when the same group uses sequence-normalized units.
      var groupedScale = [width, height, 1, 1];
      var groupedFinalScales = [];
      for (var groupedIndex = 0; groupedIndex < vectorValues.length; groupedIndex += 1) {
        groupedFinalScales.push(Number(groupedScale[groupedIndex] || 1));
      }
      return {
        mode: vectorMode,
        scale: groupedFinalScales,
        candidateId: "scale_group_sequence_axis",
        score: 0
      };
    }

    candidates.push({ id: "size_raw", scales: [1, 1, 1, 1], minPreferred: 1, maxPreferred: 400, idealValue: 100 });
    candidates.push({
      id: "size_fixed_1920",
      scales: [1920, 1080, 1, 1],
      minPreferred: 1,
      maxPreferred: 400,
      idealValue: 100
    });
    candidates.push({
      id: "size_axis",
      scales: [width, height, 1, 1],
      minPreferred: 1,
      maxPreferred: 400,
      idealValue: 100
    });
    candidates.push({
      id: "size_axis_half",
      scales: [width * 0.5, height * 0.5, 1, 1],
      minPreferred: 1,
      maxPreferred: 400,
      idealValue: 100
    });
    candidates.push({
      id: "size_axis_x2",
      scales: [width * 2, height * 2, 1, 1],
      minPreferred: 1,
      maxPreferred: 400,
      idealValue: 100
    });
    candidates.push({
      id: "size_percent_100",
      scales: [100, 100, 1, 1],
      minPreferred: 1,
      maxPreferred: 400,
      idealValue: 100
    });
    candidates.push({
      id: "size_area",
      scales: [width * height, width * height, 1, 1],
      minPreferred: 1,
      maxPreferred: 400,
      idealValue: 100
    });
  } else {
    candidates.push({ id: "raw", scales: [1, 1, 1, 1], minPreferred: 0.05, maxPreferred: 5000, idealValue: 50 });
  }

  var bestCandidate = candidates[0];
  var bestScore = 999999;

  for (var candidateIndex = 0; candidateIndex < candidates.length; candidateIndex += 1) {
    var candidate = candidates[candidateIndex];
    var projected = subcreator_visual_apply_vector_scale(vectorValues, candidate.scales);
    var score = subcreator_visual_score_vector_candidate(
      projected,
      candidate.minPreferred,
      candidate.maxPreferred,
      candidate.idealValue
    );
    if (score < bestScore) {
      bestScore = score;
      bestCandidate = candidate;
    }
  }

  var finalScales = [];
  for (var valueIndex = 0; valueIndex < vectorValues.length; valueIndex += 1) {
    finalScales.push(Number(bestCandidate.scales[valueIndex] || 1));
  }

  return {
    mode: vectorMode,
    scale: finalScales,
    candidateId: bestCandidate.id,
    score: bestScore
  };
}

function subcreator_visual_vector_to_panel_units(vectorValues, vectorScale) {
  // // Convert host vector values into panel units using inferred scale.
  if (!vectorValues || !vectorValues.length) {
    return [];
  }

  return subcreator_visual_apply_vector_scale(vectorValues, vectorScale || []);
}

function subcreator_visual_vector_to_host_units(vectorValues, vectorScale) {
  // // Convert panel vector values back to host units using inverse inferred scale.
  if (!vectorValues || !vectorValues.length) {
    return [];
  }
  var converted = [];
  for (var index = 0; index < vectorValues.length; index += 1) {
    var numericValue = Number(vectorValues[index]);
    var scale = Number(vectorScale && vectorScale[index] ? vectorScale[index] : 1);
    if (isNaN(numericValue) || isNaN(scale) || scale === 0) {
      converted.push(0);
      continue;
    }
    converted.push(numericValue / scale);
  }

  return converted;
}

function subcreator_serialize_visual_property_value(rawValue, valueType) {
  // // Serialize complex property values into text payloads usable in panel controls.
  if (valueType === "json") {
    try {
      return JSON.stringify(rawValue);
    } catch (error) {
      return String(rawValue);
    }
  }

  return rawValue;
}

function subcreator_visual_normalize_text_style_key(key) {
  // // Normalize style-field keys for resilient matching across MOGRT JSON variants.
  return String(key || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function subcreator_visual_is_font_family_key(normalizedKey) {
  // // Identify known JSON keys that carry font family/name values.
  return (
    normalizedKey === "fontname" ||
    normalizedKey === "mfontname" ||
    normalizedKey === "fontfamily" ||
    normalizedKey === "mfontfamily"
  );
}

function subcreator_visual_is_font_style_key(normalizedKey) {
  // // Identify known JSON keys that carry font style values.
  return (
    normalizedKey === "fontstyle" ||
    normalizedKey === "mfontstyle" ||
    normalizedKey === "fontstylename" ||
    normalizedKey === "mfontstylename"
  );
}

function subcreator_visual_is_font_size_key(normalizedKey) {
  // // Identify known JSON keys that carry font size values.
  return normalizedKey === "fontsize" || normalizedKey === "mfontsize";
}

function subcreator_visual_extract_text_style_from_value(rawValue) {
  // // Extract editable text style fields from text-document JSON payloads.
  var payload = rawValue;
  if (typeof payload === "string") {
    if (payload.indexOf("{") === -1) {
      return null;
    }
    try {
      payload = JSON.parse(payload);
    } catch (parseError) {
      return null;
    }
  }

  if (!payload || typeof payload !== "object") {
    return null;
  }

  var result = {
    fontFamily: "",
    fontStyle: "",
    fontSize: NaN
  };

  function scanNode(node, depth) {
    if (!node || typeof node !== "object" || depth > 12) {
      return;
    }

    if (typeof node.length === "number") {
      for (var arrIndex = 0; arrIndex < node.length; arrIndex += 1) {
        scanNode(node[arrIndex], depth + 1);
      }
      return;
    }

    for (var key in node) {
      if (!node.hasOwnProperty(key)) {
        continue;
      }
      var value = node[key];
      var normalizedKey = subcreator_visual_normalize_text_style_key(key);

      if (!result.fontFamily && subcreator_visual_is_font_family_key(normalizedKey)) {
        var familyValue = subcreator_trim_string(String(value || ""));
        if (familyValue) {
          result.fontFamily = familyValue;
        }
      }

      if (!result.fontStyle && subcreator_visual_is_font_style_key(normalizedKey)) {
        var styleValue = subcreator_trim_string(String(value || ""));
        if (styleValue) {
          result.fontStyle = styleValue;
        }
      }

      if (isNaN(result.fontSize) && subcreator_visual_is_font_size_key(normalizedKey)) {
        var sizeValue = Number(value);
        if (!isNaN(sizeValue) && sizeValue > 0 && sizeValue < 2000) {
          result.fontSize = sizeValue;
        }
      }

      if (value && typeof value === "object") {
        scanNode(value, depth + 1);
      }
    }
  }

  scanNode(payload, 0);

  if (!result.fontFamily && !result.fontStyle && isNaN(result.fontSize)) {
    return null;
  }

  return result;
}

function subcreator_visual_build_text_style_entries(rawValue, currentPath, groupPath) {
  // // Build synthetic visual-editor entries for font family/style/size from text payloads.
  var styleValues = subcreator_visual_extract_text_style_from_value(rawValue);
  if (!styleValues) {
    return [];
  }

  var entries = [];
  var targetGroup = groupPath || "General";

  if (styleValues.fontFamily) {
    entries.push({
      path: currentPath + "::textstyle.fontFamily",
      displayName: "Font Family",
      groupPath: targetGroup,
      valueType: "string",
      controlKind: "string",
      value: styleValues.fontFamily
    });
  }

  if (styleValues.fontStyle) {
    entries.push({
      path: currentPath + "::textstyle.fontStyle",
      displayName: "Font Style",
      groupPath: targetGroup,
      valueType: "string",
      controlKind: "string",
      value: styleValues.fontStyle
    });
  }

  if (!isNaN(styleValues.fontSize)) {
    entries.push({
      path: currentPath + "::textstyle.fontSize",
      displayName: "Font Size",
      groupPath: targetGroup,
      valueType: "number",
      controlKind: "slider",
      minValue: 1,
      maxValue: 500,
      stepValue: 0.1,
      value: styleValues.fontSize
    });
  }

  return entries;
}

function subcreator_build_visual_property_entry(property, currentPath, displayName, groupPath, hasChildren, rawValueOverride) {
  // // Build one panel-ready visual property entry with inferred control metadata.
  var rawValue = typeof rawValueOverride !== "undefined" ? rawValueOverride : undefined;
  if (typeof rawValue === "undefined") {
    try {
      rawValue = property.getValue();
    } catch (readError) {
      return null;
    }
  }

  if (typeof rawValue === "undefined") {
    return null;
  }

  if (subcreator_visual_is_guid_list_string(rawValue)) {
    return null;
  }

  if (hasChildren && subcreator_visual_is_group_metadata_value(rawValue)) {
    return null;
  }

  var detectedType = subcreator_detect_visual_property_type(rawValue);
  var key = String(displayName || "").toLowerCase();
  var shouldTreatAsText = subcreator_should_try_text_property(displayName, rawValue);
  if (detectedType === "string" && subcreator_visual_is_numeric_string(rawValue)) {
    detectedType = "number";
  }

  var hasColorApi = !!(property && (typeof property.getColorValue === "function" || typeof property.setColorValue === "function"));
  var groupSuggestsColor = subcreator_visual_group_suggests_color(groupPath);
  var colorCandidate = subcreator_visual_is_color_label(displayName);
  var colorLayoutCandidates = subcreator_visual_build_color_layout_candidates(displayName, groupPath, "read");
  var colorLayoutHint = colorLayoutCandidates.length ? colorLayoutCandidates[0] : "";
  var allowPackedColor = colorCandidate || groupSuggestsColor;
  var colorHex = allowPackedColor
    ? subcreator_visual_try_read_property_color_hex(property, rawValue, true, colorLayoutHint)
    : "";
  var colorBlocked =
    key.indexOf("width") !== -1 ||
    key.indexOf("size") !== -1 ||
    key.indexOf("amount") !== -1 ||
    key.indexOf("opacity") !== -1 ||
    key.indexOf("based on") !== -1 ||
    key.indexOf("paragraph") !== -1 ||
    key.indexOf("align") !== -1 ||
    key.indexOf("start") !== -1 ||
    key.indexOf("end") !== -1 ||
    key.indexOf("feather") !== -1;
  var looksLikeColor = !!(
    colorHex &&
    !colorBlocked &&
    (colorCandidate ||
      (groupSuggestsColor && (subcreator_visual_is_likely_color_payload(rawValue) || hasColorApi)) ||
      key.indexOf("rgb") !== -1)
  );
  var vectorValue = subcreator_visual_extract_numeric_vector(rawValue);

  // // Do not expose subtitle text in visual editor to avoid overriding all generated captions.
  if (shouldTreatAsText) {
    return null;
  }

  if (looksLikeColor) {
    if (colorLayoutHint) {
      // // Keep the decode layout used during property listing for later apply verification.
      subcreator_visual_set_cached_color_layout(displayName, colorLayoutHint, "read");
    }
    return {
      path: currentPath,
      displayName: displayName,
      groupPath: groupPath || "General",
      valueType: "string",
      controlKind: "color",
      value: colorHex
    };
  }

  if (detectedType === "boolean") {
    return {
      path: currentPath,
      displayName: displayName,
      groupPath: groupPath || "General",
      valueType: "boolean",
      controlKind: "checkbox",
      value: !!rawValue
    };
  }

  if (detectedType === "number") {
    var range = subcreator_visual_read_numeric_range(property, displayName, rawValue);
    var value = subcreator_visual_to_number(rawValue);
    if (isNaN(value)) {
      value = 0;
    }

    var selectOptions = subcreator_visual_build_select_options(displayName, value);
    if (selectOptions && selectOptions.length > 0) {
      return {
        path: currentPath,
        displayName: displayName,
        groupPath: groupPath || "General",
        valueType: "number",
        controlKind: "select",
        options: selectOptions,
        value: value
      };
    }

    var useSlider = !subcreator_visual_is_discrete_numeric_label(displayName);
    var descriptor = {
      path: currentPath,
      displayName: displayName,
      groupPath: groupPath || "General",
      valueType: "number",
      controlKind: useSlider ? "slider" : "number",
      value: value
    };

    if (useSlider) {
      descriptor.minValue = range.minValue;
      descriptor.maxValue = range.maxValue;
      descriptor.stepValue = range.stepValue;
    }

    return descriptor;
  }

  if (vectorValue) {
    var sequenceSize = subcreator_visual_read_sequence_dimensions();
    var vectorMeta = subcreator_visual_choose_vector_scale(displayName, groupPath, vectorValue, sequenceSize);
    var displayVector = subcreator_visual_vector_to_panel_units(vectorValue, vectorMeta.scale);
    return {
      path: currentPath,
      displayName: displayName,
      groupPath: groupPath || "General",
      valueType: "json",
      controlKind: "vector",
      value: JSON.stringify(displayVector),
      vectorScale: vectorMeta.scale,
      vectorMode: vectorMeta.mode,
      debugVector: {
        mode: vectorMeta.mode,
        candidateId: vectorMeta.candidateId,
        score: vectorMeta.score,
        raw: vectorValue,
        scale: vectorMeta.scale,
        panel: displayVector,
        sequenceWidth: sequenceSize.width,
        sequenceHeight: sequenceSize.height
      }
    };
  }

  if (detectedType === "json") {
    // // Skip unsupported JSON blobs to keep the editor compact and practical.
    return null;
  }

  return {
    path: currentPath,
    displayName: displayName,
    groupPath: groupPath || "General",
    valueType: "string",
    controlKind: "string",
    value: String(rawValue || "")
  };
}

function subcreator_collect_mogrt_visual_properties_recursive(
  propertyCollection,
  pathPrefix,
  groupPathPrefix,
  collector
) {
  // // Traverse nested Essential Graphics properties and capture editable entries with group paths.
  if (!propertyCollection || typeof propertyCollection.numItems !== "number") {
    return;
  }

  var activeSiblingGroupPath = String(groupPathPrefix || "");

  for (var index = 0; index < propertyCollection.numItems; index += 1) {
    var property = propertyCollection[index];
    if (!property) {
      continue;
    }

    var currentPath = pathPrefix ? pathPrefix + "." + String(index) : String(index);
    var displayName = subcreator_trim_string(String(property.displayName || ""));
    if (!displayName) {
      displayName = "Property " + currentPath;
    }

    var rawValue = undefined;
    var hasValue = false;
    if (typeof property.getValue === "function") {
      try {
        rawValue = property.getValue();
        hasValue = true;
      } catch (readValueError) {
        hasValue = false;
      }
    }

    var hasChildren = !!(
      property.properties &&
      typeof property.properties.numItems === "number" &&
      property.properties.numItems > 0
    );
    if (subcreator_visual_is_guid_list_string(displayName)) {
      displayName = "Group " + String(index + 1);
    }

    // // Group markers are represented as GUID-list payload strings in many MOGRT templates.
    if (hasValue && subcreator_visual_is_guid_list_string(rawValue)) {
      activeSiblingGroupPath = groupPathPrefix ? groupPathPrefix + " / " + displayName : displayName;

      if (hasChildren) {
        subcreator_collect_mogrt_visual_properties_recursive(
          property.properties,
          currentPath,
          activeSiblingGroupPath,
          collector
        );
      }
      continue;
    }

    var resolvedGroupPath = activeSiblingGroupPath || groupPathPrefix || "General";

    if (typeof property.getValue === "function" && typeof property.setValue === "function") {
      if (hasValue) {
        // // Expose style-only controls from text payloads while keeping actual caption text hidden.
        var textStyleEntries = subcreator_visual_build_text_style_entries(rawValue, currentPath, resolvedGroupPath);
        for (var styleIndex = 0; styleIndex < textStyleEntries.length; styleIndex += 1) {
          collector.push(textStyleEntries[styleIndex]);
        }
        if (textStyleEntries.length > 0) {
          continue;
        }
      }

      var descriptor = subcreator_build_visual_property_entry(
        property,
        currentPath,
        displayName,
        resolvedGroupPath,
        hasChildren,
        hasValue ? rawValue : undefined
      );
      if (descriptor) {
        collector.push(descriptor);
      }
    }

    if (hasChildren) {
      var nextGroupPath = resolvedGroupPath ? resolvedGroupPath + " / " + displayName : displayName;
      subcreator_collect_mogrt_visual_properties_recursive(
        property.properties,
        currentPath,
        nextGroupPath,
        collector
      );
    }
  }
}

function subcreator_find_property_by_path(propertyCollection, pathValue) {
  // // Resolve nested property by index-path notation (`0.2.4`) from panel payload.
  var pathText = subcreator_trim_string(String(pathValue || ""));
  if (!pathText) {
    return null;
  }

  var chunks = pathText.split(".");
  var collection = propertyCollection;
  var property = null;

  for (var index = 0; index < chunks.length; index += 1) {
    if (!collection || typeof collection.numItems !== "number") {
      return null;
    }

    var itemIndex = Number(chunks[index]);
    if (isNaN(itemIndex) || itemIndex < 0 || itemIndex >= collection.numItems) {
      return null;
    }

    property = collection[itemIndex];
    if (!property) {
      return null;
    }

    if (index === chunks.length - 1) {
      return property;
    }

    collection = property.properties;
  }

  return null;
}

function subcreator_visual_parse_text_style_virtual_path(pathValue) {
  // // Decode synthetic visual-editor paths like `4::textstyle.fontSize`.
  var pathText = subcreator_trim_string(String(pathValue || ""));
  var marker = "::textstyle.";
  var markerIndex = pathText.indexOf(marker);
  if (markerIndex <= 0) {
    return null;
  }

  var basePath = subcreator_trim_string(pathText.substring(0, markerIndex));
  var styleKey = subcreator_trim_string(pathText.substring(markerIndex + marker.length));
  if (!basePath || !styleKey) {
    return null;
  }

  return {
    basePath: basePath,
    styleKey: styleKey
  };
}

function subcreator_visual_normalize_text_style_change(styleKey, value) {
  // // Normalize incoming style values before patching text-document payloads.
  var normalizedStyleKey = subcreator_trim_string(String(styleKey || ""));
  if (!normalizedStyleKey) {
    return null;
  }

  if (normalizedStyleKey === "fontSize") {
    var sizeValue = Number(value);
    if (isNaN(sizeValue) || sizeValue <= 0 || sizeValue > 2000) {
      return null;
    }
    return sizeValue;
  }

  var textValue = subcreator_trim_string(String(value || ""));
  if (!textValue) {
    return null;
  }
  return textValue;
}

function subcreator_visual_apply_text_style_to_payload(payload, styleKey, styleValue) {
  // // Apply one style field recursively to known text JSON keys.
  if (!payload || typeof payload !== "object") {
    return false;
  }

  var updated = false;

  function patchNode(node, depth) {
    if (!node || typeof node !== "object" || depth > 12) {
      return;
    }

    if (typeof node.length === "number") {
      for (var arrIndex = 0; arrIndex < node.length; arrIndex += 1) {
        patchNode(node[arrIndex], depth + 1);
      }
      return;
    }

    for (var key in node) {
      if (!node.hasOwnProperty(key)) {
        continue;
      }

      var value = node[key];
      var normalizedKey = subcreator_visual_normalize_text_style_key(key);

      if (styleKey === "fontFamily" && subcreator_visual_is_font_family_key(normalizedKey)) {
        node[key] = String(styleValue);
        updated = true;
      } else if (styleKey === "fontStyle" && subcreator_visual_is_font_style_key(normalizedKey)) {
        node[key] = String(styleValue);
        updated = true;
      } else if (styleKey === "fontSize" && subcreator_visual_is_font_size_key(normalizedKey)) {
        node[key] = Number(styleValue);
        updated = true;
      }

      if (value && typeof value === "object") {
        patchNode(value, depth + 1);
      }
    }
  }

  patchNode(payload, 0);
  return updated;
}

function subcreator_try_patch_text_style_json_string(rawValue, styleKey, styleValue) {
  // // Fallback patch for JSON-like strings when `JSON.parse` is unavailable on host payloads.
  var raw = String(rawValue || "");
  if (!raw || raw.indexOf("{") === -1) {
    return "";
  }

  var patched = raw;
  var styleString = JSON.stringify(String(styleValue));
  var keyList = [];

  if (styleKey === "fontFamily") {
    keyList = ["fontName", "mFontName", "fontFamily", "mFontFamily"];
  } else if (styleKey === "fontStyle") {
    keyList = ["fontStyle", "mFontStyle", "fontStyleName", "mFontStyleName"];
  } else if (styleKey === "fontSize") {
    keyList = ["fontSize", "mFontSize"];
  }

  for (var keyIndex = 0; keyIndex < keyList.length; keyIndex += 1) {
    var keyName = keyList[keyIndex];
    if (styleKey === "fontSize") {
      var numericRegex = new RegExp('"' + keyName + '"\\s*:\\s*("([^"\\\\]|\\\\.)*"|-?\\d+(?:\\.\\d+)?)', "g");
      patched = patched.replace(numericRegex, '"' + keyName + '":' + String(Number(styleValue)));
    } else {
      var stringRegex = new RegExp('"' + keyName + '"\\s*:\\s*"([^"\\\\]|\\\\.)*"', "g");
      patched = patched.replace(stringRegex, '"' + keyName + '":' + styleString);
    }
  }

  if (patched === raw) {
    return "";
  }

  return patched;
}

function subcreator_try_set_mogrt_text_style_property(property, styleKey, styleValue) {
  // // Apply editable style-only text controls without mutating subtitle content.
  if (!property || typeof property.setValue !== "function") {
    return false;
  }

  var normalizedStyleKey = subcreator_trim_string(String(styleKey || ""));
  if (!normalizedStyleKey) {
    return false;
  }

  var normalizedStyleValue = subcreator_visual_normalize_text_style_change(normalizedStyleKey, styleValue);
  if (normalizedStyleValue === null) {
    return false;
  }

  var rawValue = "";
  if (typeof property.getValue === "function") {
    try {
      rawValue = property.getValue();
    } catch (getError) {
      rawValue = "";
    }
  }

  if (
    !subcreator_should_try_text_property(property.displayName || "", rawValue) &&
    !subcreator_visual_extract_text_style_from_value(rawValue)
  ) {
    return false;
  }

  if (rawValue && typeof rawValue === "object") {
    try {
      var objectCopy = JSON.parse(JSON.stringify(rawValue));
      if (subcreator_visual_apply_text_style_to_payload(objectCopy, normalizedStyleKey, normalizedStyleValue)) {
        property.setValue(objectCopy, true);
        return true;
      }
    } catch (copyError) {}

    try {
      if (subcreator_visual_apply_text_style_to_payload(rawValue, normalizedStyleKey, normalizedStyleValue)) {
        property.setValue(rawValue, true);
        return true;
      }
    } catch (directError) {}
  }

  if (typeof rawValue === "string" && rawValue.indexOf("{") !== -1) {
    try {
      var parsed = JSON.parse(rawValue);
      if (subcreator_visual_apply_text_style_to_payload(parsed, normalizedStyleKey, normalizedStyleValue)) {
        property.setValue(JSON.stringify(parsed), true);
        return true;
      }
    } catch (jsonError) {}

    try {
      var patchedRaw = subcreator_try_patch_text_style_json_string(rawValue, normalizedStyleKey, normalizedStyleValue);
      if (patchedRaw) {
        property.setValue(patchedRaw, true);
        return true;
      }
    } catch (patchError) {}
  }

  return false;
}

function subcreator_normalize_visual_payload_value(valueType, rawValue) {
  // // Convert panel-sent values to host-friendly types before property.setValue.
  if (valueType === "number") {
    return Number(rawValue);
  }

  if (valueType === "boolean") {
    if (typeof rawValue === "boolean") {
      return rawValue;
    }
    var text = subcreator_trim_string(String(rawValue || "")).toLowerCase();
    return text === "true" || text === "1" || text === "yes";
  }

  if (valueType === "json") {
    if (typeof rawValue === "string") {
      try {
        return JSON.parse(rawValue);
      } catch (jsonError) {
        return rawValue;
      }
    }
    return rawValue;
  }

  return String(rawValue || "");
}

function subcreator_visual_parse_hex_color(value) {
  // // Parse CSS hex color strings into RGB channels.
  var text = subcreator_trim_string(String(value || ""));
  if (!text) {
    return null;
  }

  if (/^#[0-9a-f]{3}$/i.test(text)) {
    return {
      red: parseInt(text.charAt(1) + text.charAt(1), 16),
      green: parseInt(text.charAt(2) + text.charAt(2), 16),
      blue: parseInt(text.charAt(3) + text.charAt(3), 16)
    };
  }

  if (/^#[0-9a-f]{6}$/i.test(text)) {
    return {
      red: parseInt(text.substring(1, 3), 16),
      green: parseInt(text.substring(3, 5), 16),
      blue: parseInt(text.substring(5, 7), 16)
    };
  }

  return null;
}

function subcreator_visual_try_read_property_rgb(property, allowPackedFallback, preferredArrayLayout) {
  // // Read RGB channels from getColorValue first, then getValue fallback when needed.
  if (!property) {
    return null;
  }

  if (typeof property.getColorValue === "function") {
    try {
      var colorValue = property.getColorValue();
      var fromColorApi = subcreator_visual_extract_rgb_from_value(colorValue, true, preferredArrayLayout);
      if (fromColorApi) {
        return fromColorApi;
      }
    } catch (colorApiReadError) {}
  }

  if (typeof property.getValue === "function") {
    try {
      var rawValue = property.getValue();
      return subcreator_visual_extract_rgb_from_value(rawValue, allowPackedFallback === true, preferredArrayLayout);
    } catch (valueReadError) {}
  }

  return null;
}

function subcreator_visual_color_distance(leftRgb, rightRgb) {
  // // Compute per-channel absolute distance to validate color writes.
  if (!leftRgb || !rightRgb) {
    return 9999;
  }

  return (
    Math.abs(Number(leftRgb.red) - Number(rightRgb.red)) +
    Math.abs(Number(leftRgb.green) - Number(rightRgb.green)) +
    Math.abs(Number(leftRgb.blue) - Number(rightRgb.blue))
  );
}

function subcreator_visual_apply_rgb_to_payload(payload, rgb) {
  // // Patch object/array color payloads while preserving their original numeric scale.
  if (!payload || typeof payload !== "object" || !rgb) {
    return false;
  }

  var updated = false;

  function setTriplet(target, redKey, greenKey, blueKey) {
    if (
      typeof target[redKey] === "undefined" ||
      typeof target[greenKey] === "undefined" ||
      typeof target[blueKey] === "undefined"
    ) {
      return false;
    }

    var redValue = Number(target[redKey]);
    var greenValue = Number(target[greenKey]);
    var blueValue = Number(target[blueKey]);
    var useUnitScale = !isNaN(redValue) && !isNaN(greenValue) && !isNaN(blueValue) && redValue <= 1 && greenValue <= 1 && blueValue <= 1;

    target[redKey] = useUnitScale ? rgb.red / 255 : rgb.red;
    target[greenKey] = useUnitScale ? rgb.green / 255 : rgb.green;
    target[blueKey] = useUnitScale ? rgb.blue / 255 : rgb.blue;
    return true;
  }

  if (typeof payload.length === "number" && payload.length >= 3) {
    var c0 = Number(payload[0]);
    var c1 = Number(payload[1]);
    var c2 = Number(payload[2]);
    var unitArrayScale = !isNaN(c0) && !isNaN(c1) && !isNaN(c2) && c0 <= 1 && c1 <= 1 && c2 <= 1;
    payload[0] = unitArrayScale ? rgb.red / 255 : rgb.red;
    payload[1] = unitArrayScale ? rgb.green / 255 : rgb.green;
    payload[2] = unitArrayScale ? rgb.blue / 255 : rgb.blue;
    updated = true;
  }

  if (setTriplet(payload, "red", "green", "blue")) {
    updated = true;
  }

  if (setTriplet(payload, "r", "g", "b")) {
    updated = true;
  }

  if (payload.color && typeof payload.color === "object") {
    if (subcreator_visual_apply_rgb_to_payload(payload.color, rgb)) {
      updated = true;
    }
  }

  if (payload.value && typeof payload.value === "object") {
    if (subcreator_visual_apply_rgb_to_payload(payload.value, rgb)) {
      updated = true;
    }
  }

  return updated;
}

function subcreator_try_set_mogrt_color_property(property, value) {
  // // Apply color values from panel hex input to color-capable MOGRT controls.
  if (!property || (typeof property.setValue !== "function" && typeof property.setColorValue !== "function")) {
    return false;
  }

  var rgb = subcreator_visual_parse_hex_color(value);
  if (!rgb) {
    return false;
  }
  var colorDisplayName = subcreator_trim_string(String(property.displayName || ""));
  var colorWriteLayoutCandidates = subcreator_visual_build_color_layout_candidates(colorDisplayName, "", "write");
  var colorReadLayoutCandidates = subcreator_visual_build_color_layout_candidates(colorDisplayName, "", "read");
  var colorLayoutHint = colorReadLayoutCandidates.length ? colorReadLayoutCandidates[0] : "";

  var fallbackRgb = {
    red: rgb.blue,
    green: rgb.green,
    blue: rgb.red
  };

  var colorOrders = [rgb, fallbackRgb];
  var colorDistanceThreshold = 8;

  function applyAndVerify(applyCallback, readLayout) {
    // // Apply one write strategy and verify readback when host API can expose a color.
    var attempted = false;
    try {
      attempted = applyCallback() !== false;
    } catch (applyError) {
      return false;
    }

    if (!attempted) {
      return false;
    }

    var readbackLayout = subcreator_trim_string(String(readLayout || colorLayoutHint || ""));
    var readback = subcreator_visual_try_read_property_rgb(property, true, readbackLayout);
    if (!readback) {
      return true;
    }

    return subcreator_visual_color_distance(readback, rgb) <= colorDistanceThreshold;
  }

  var rawValue = "";
  if (typeof property.getValue === "function") {
    try {
      rawValue = property.getValue();
    } catch (getError) {
      rawValue = "";
    }
  }

  var colorApiValue = null;
  var hasColorApiValue = false;
  if (typeof property.getColorValue === "function") {
    try {
      colorApiValue = property.getColorValue();
      hasColorApiValue = true;
    } catch (getColorError) {
      hasColorApiValue = false;
      colorApiValue = null;
    }
  }

  function trySetColorByApiShape(referenceValue, candidateRgb, layoutOverride) {
    // // Match native setColorValue payload shape to avoid unsupported host writes.
    if (typeof property.setColorValue !== "function") {
      return false;
    }

    if (!referenceValue) {
      return false;
    }

    try {
      if (typeof referenceValue.length === "number" && referenceValue.length >= 3) {
        var v0 = Number(referenceValue[0]);
        var v1 = Number(referenceValue[1]);
        var v2 = Number(referenceValue[2]);
        var v3 = Number(referenceValue[3]);
        var arrayLayout =
          subcreator_trim_string(String(layoutOverride || "")) || colorLayoutHint || subcreator_visual_detect_color_array_layout(referenceValue);
        var hasFourChannels = typeof referenceValue.length === "number" && referenceValue.length >= 4;
        var channelsUseUnit = false;
        var alphaSource = 1;

        if (arrayLayout === "argb") {
          channelsUseUnit = !isNaN(v1) && !isNaN(v2) && !isNaN(v3) && v1 <= 1 && v2 <= 1 && v3 <= 1;
          alphaSource = !isNaN(v0) ? v0 : 1;
        } else if (arrayLayout === "rgba" || arrayLayout === "bgra") {
          channelsUseUnit = !isNaN(v0) && !isNaN(v1) && !isNaN(v2) && v0 <= 1 && v1 <= 1 && v2 <= 1;
          alphaSource = !isNaN(v3) ? v3 : 1;
        } else if (arrayLayout === "abgr") {
          channelsUseUnit = !isNaN(v1) && !isNaN(v2) && !isNaN(v3) && v1 <= 1 && v2 <= 1 && v3 <= 1;
          alphaSource = !isNaN(v0) ? v0 : 1;
        } else {
          channelsUseUnit = !isNaN(v0) && !isNaN(v1) && !isNaN(v2) && v0 <= 1 && v1 <= 1 && v2 <= 1;
          alphaSource = 1;
        }

        var alphaAsUnit = alphaSource <= 1;
        var alphaUnit = alphaAsUnit ? alphaSource : alphaSource / 255;
        var alpha255 = alphaAsUnit ? Math.round(alphaSource * 255) : alphaSource;
        var redPayload = channelsUseUnit ? candidateRgb.red / 255 : candidateRgb.red;
        var greenPayload = channelsUseUnit ? candidateRgb.green / 255 : candidateRgb.green;
        var bluePayload = channelsUseUnit ? candidateRgb.blue / 255 : candidateRgb.blue;
        var payload = [redPayload, greenPayload, bluePayload];

        if (hasFourChannels) {
          var indices = subcreator_visual_color_layout_indices(arrayLayout, 4);
          payload = [0, 0, 0, 0];
          payload[indices.red] = redPayload;
          payload[indices.green] = greenPayload;
          payload[indices.blue] = bluePayload;
          if (indices.alpha >= 0 && indices.alpha < payload.length) {
            payload[indices.alpha] = channelsUseUnit ? alphaUnit : alpha255;
          }
        }

        try {
          property.setColorValue(payload, true);
          return true;
        } catch (arrayUiError) {}

        try {
          property.setColorValue(payload);
          return true;
        } catch (arrayError) {}

        try {
          if (payload.length >= 4) {
            property.setColorValue(payload[0], payload[1], payload[2], payload[3]);
          } else {
            property.setColorValue(payload[0], payload[1], payload[2]);
          }
          return true;
        } catch (positionalError) {}
      }
    } catch (arrayShapeError) {}

    try {
      if (
        typeof referenceValue.red !== "undefined" ||
        typeof referenceValue.green !== "undefined" ||
        typeof referenceValue.blue !== "undefined"
      ) {
        var red = Number(referenceValue.red);
        var green = Number(referenceValue.green);
        var blue = Number(referenceValue.blue);
        var alpha = Number(referenceValue.alpha);
        var objectUsesUnit = !isNaN(red) && !isNaN(green) && !isNaN(blue) && red <= 1 && green <= 1 && blue <= 1;

        if (objectUsesUnit) {
          var objectUnitPayload = {
            red: candidateRgb.red / 255,
            green: candidateRgb.green / 255,
            blue: candidateRgb.blue / 255,
            alpha: isNaN(alpha) ? 1 : alpha
          };
          try {
            property.setColorValue(objectUnitPayload, true);
            return true;
          } catch (objectUnitUiError) {}

          try {
            property.setColorValue(objectUnitPayload);
            return true;
          } catch (objectUnitError) {}
        } else {
          var objectBytePayload = {
            red: candidateRgb.red,
            green: candidateRgb.green,
            blue: candidateRgb.blue,
            alpha: isNaN(alpha) ? 255 : alpha
          };
          try {
            property.setColorValue(objectBytePayload, true);
            return true;
          } catch (objectByteUiError) {}

          try {
            property.setColorValue(objectBytePayload);
            return true;
          } catch (objectByteError) {}
        }
      }
    } catch (objectShapeError) {}

    return false;
  }

  function tryApplyStructuredPayload(structuredValue, stringifyJson) {
    // // Try object/json payload rewrites with RGB then BGR fallback.
    if (!structuredValue || typeof structuredValue !== "object") {
      return false;
    }

    for (var colorOrderIndex = 0; colorOrderIndex < colorOrders.length; colorOrderIndex += 1) {
      var candidateRgb = colorOrders[colorOrderIndex];
      try {
        var payloadCopy = JSON.parse(JSON.stringify(structuredValue));
        if (!subcreator_visual_apply_rgb_to_payload(payloadCopy, candidateRgb)) {
          continue;
        }

        if (
          applyAndVerify(function () {
            if (typeof property.setValue !== "function") {
              return false;
            }
            property.setValue(stringifyJson ? JSON.stringify(payloadCopy) : payloadCopy, true);
            return true;
          })
        ) {
          return true;
        }
      } catch (payloadError) {}
    }

    return false;
  }

  if (rawValue && typeof rawValue === "object") {
    if (tryApplyStructuredPayload(rawValue, false)) {
      return true;
    }
  }

  if (typeof rawValue === "string" && (rawValue.indexOf("{") !== -1 || rawValue.indexOf("[") !== -1)) {
    try {
      var parsed = JSON.parse(rawValue);
      if (tryApplyStructuredPayload(parsed, true)) {
        return true;
      }
    } catch (jsonError) {}
  }

  if (hasColorApiValue) {
    for (var layoutIndex = 0; layoutIndex < colorWriteLayoutCandidates.length; layoutIndex += 1) {
      var layoutCandidate = colorWriteLayoutCandidates[layoutIndex];
      for (var apiOrderIndex = 0; apiOrderIndex < colorOrders.length; apiOrderIndex += 1) {
        var apiRgb = colorOrders[apiOrderIndex];
        if (colorLayoutHint) {
          if (
            applyAndVerify(
              function () {
                return trySetColorByApiShape(colorApiValue, apiRgb, layoutCandidate);
              },
              colorLayoutHint
            )
          ) {
            subcreator_visual_set_cached_color_layout(colorDisplayName, layoutCandidate, "write");
            subcreator_visual_set_cached_color_layout(colorDisplayName, colorLayoutHint, "read");
            return true;
          }
        }

        for (var readLayoutIndex = 0; readLayoutIndex < colorReadLayoutCandidates.length; readLayoutIndex += 1) {
          var readLayoutCandidate = colorReadLayoutCandidates[readLayoutIndex];
          if (readLayoutCandidate && readLayoutCandidate === colorLayoutHint) {
            continue;
          }

          if (
            applyAndVerify(
              function () {
                return trySetColorByApiShape(colorApiValue, apiRgb, layoutCandidate);
              },
              readLayoutCandidate
            )
          ) {
            subcreator_visual_set_cached_color_layout(colorDisplayName, layoutCandidate, "write");
            subcreator_visual_set_cached_color_layout(colorDisplayName, readLayoutCandidate, "read");
            return true;
          }
        }
      }
    }
  }

  for (var fallbackOrderIndex = 0; fallbackOrderIndex < colorOrders.length; fallbackOrderIndex += 1) {
    var fallbackRgbValue = colorOrders[fallbackOrderIndex];

    if (
      applyAndVerify(function () {
        if (typeof property.setColorValue !== "function") {
          return false;
        }
        property.setColorValue(fallbackRgbValue.red, fallbackRgbValue.green, fallbackRgbValue.blue, 255);
        return true;
      })
    ) {
      return true;
    }

    if (
      applyAndVerify(function () {
        if (typeof property.setColorValue !== "function") {
          return false;
        }
        property.setColorValue(
          [fallbackRgbValue.red / 255, fallbackRgbValue.green / 255, fallbackRgbValue.blue / 255, 1],
          true
        );
        return true;
      })
    ) {
      return true;
    }

    if (
      applyAndVerify(function () {
        if (typeof property.setColorValue !== "function") {
          return false;
        }
        property.setColorValue([fallbackRgbValue.red, fallbackRgbValue.green, fallbackRgbValue.blue, 255], true);
        return true;
      })
    ) {
      return true;
    }

    if (
      applyAndVerify(function () {
        if (typeof property.setColorValue !== "function") {
          return false;
        }
        property.setColorValue(
          {
            red: fallbackRgbValue.red / 255,
            green: fallbackRgbValue.green / 255,
            blue: fallbackRgbValue.blue / 255,
            alpha: 1
          },
          true
        );
        return true;
      })
    ) {
      return true;
    }

    if (
      applyAndVerify(function () {
        if (typeof property.setValue !== "function") {
          return false;
        }
        property.setValue([fallbackRgbValue.red / 255, fallbackRgbValue.green / 255, fallbackRgbValue.blue / 255, 1], true);
        return true;
      })
    ) {
      return true;
    }

    if (
      applyAndVerify(function () {
        if (typeof property.setValue !== "function") {
          return false;
        }
        property.setValue([fallbackRgbValue.red, fallbackRgbValue.green, fallbackRgbValue.blue, 255], true);
        return true;
      })
    ) {
      return true;
    }

    if (
      applyAndVerify(function () {
        if (typeof property.setValue !== "function") {
          return false;
        }
        property.setValue(
          subcreator_visual_rgb_to_hex(fallbackRgbValue.red, fallbackRgbValue.green, fallbackRgbValue.blue),
          true
        );
        return true;
      })
    ) {
      return true;
    }

    if (typeof rawValue === "number") {
      var packedRgb = fallbackRgbValue.red * 65536 + fallbackRgbValue.green * 256 + fallbackRgbValue.blue;
      var packedBrg = fallbackRgbValue.blue * 65536 + fallbackRgbValue.red * 256 + fallbackRgbValue.green;

      if (
        applyAndVerify(function () {
          if (typeof property.setValue !== "function") {
            return false;
          }
          property.setValue(packedRgb, true);
          return true;
        })
      ) {
        return true;
      }

      if (
        applyAndVerify(function () {
          if (typeof property.setValue !== "function") {
            return false;
          }
          property.setValue(255 * 16777216 + packedRgb, true);
          return true;
        })
      ) {
        return true;
      }

      if (
        applyAndVerify(function () {
          if (typeof property.setValue !== "function") {
            return false;
          }
          property.setValue(packedBrg, true);
          return true;
        })
      ) {
        return true;
      }

      if (
        applyAndVerify(function () {
          if (typeof property.setValue !== "function") {
            return false;
          }
          property.setValue(255 * 16777216 + packedBrg, true);
          return true;
        })
      ) {
        return true;
      }
    }
  }

  return false;
}

function subcreator_force_sequence_visual_refresh(sequence) {
  // // Force Program Monitor redraw by nudging and restoring the playhead position.
  if (!sequence || typeof sequence.getPlayerPosition !== "function" || typeof sequence.setPlayerPosition !== "function") {
    return false;
  }

  try {
    var currentPosition = sequence.getPlayerPosition();
    var currentSeconds = subcreator_to_seconds(currentPosition);
    if (isNaN(currentSeconds)) {
      return false;
    }

    var currentTicks = "";
    if (currentPosition && typeof currentPosition.ticks !== "undefined") {
      currentTicks = String(currentPosition.ticks || "");
    }

    var sequenceEndSeconds = subcreator_to_seconds(sequence.end);
    var nudgeSeconds = 1 / 30;
    var targetSeconds = currentSeconds + nudgeSeconds;
    if (!isNaN(sequenceEndSeconds) && targetSeconds > sequenceEndSeconds) {
      targetSeconds = Math.max(0, currentSeconds - nudgeSeconds);
    }

    var nudgeTime = new Time();
    nudgeTime.seconds = targetSeconds;
    sequence.setPlayerPosition(String(nudgeTime.ticks));

    if (currentTicks) {
      sequence.setPlayerPosition(currentTicks);
      return true;
    }

    var restoreTime = new Time();
    restoreTime.seconds = currentSeconds;
    sequence.setPlayerPosition(String(restoreTime.ticks));
    return true;
  } catch (refreshError) {}

  try {
    if (app && typeof app.refresh === "function") {
      app.refresh();
      return true;
    }
  } catch (appRefreshError) {}

  return false;
}

function subcreator_list_selected_mogrt_properties() {
  // // Return editable visual properties from selected MOGRT clips in active sequence.
  try {
    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence in Premiere.");
    }

    var sequence = app.project.activeSequence;
    var selectedItems = subcreator_get_selected_track_items(sequence);
    var mogrtItems = [];

    for (var index = 0; index < selectedItems.length; index += 1) {
      var item = selectedItems[index];
      if (subcreator_get_mogrt_component_from_track_item(item)) {
        mogrtItems.push(item);
      }
    }

    if (!mogrtItems.length) {
      return subcreator_ok({
        selectedCount: 0,
        editableCount: 0,
        properties: []
      });
    }

    var firstComponent = subcreator_get_mogrt_component_from_track_item(mogrtItems[0]);
    var properties = [];
    var sequenceSize = subcreator_visual_read_sequence_dimensions();
    subcreator_visual_reset_group_sequence_axis_preferences();
    subcreator_collect_mogrt_visual_properties_recursive(
      firstComponent ? firstComponent.properties : null,
      "",
      "",
      properties
    );

    var debug = {
      sequenceWidth: sequenceSize.width,
      sequenceHeight: sequenceSize.height,
      vectorCount: 0,
      colorCount: 0,
      selectCount: 0,
      sample: []
    };
    for (var propertyIndex = 0; propertyIndex < properties.length; propertyIndex += 1) {
      var item = properties[propertyIndex];
      if (!item) {
        continue;
      }
      if (item.controlKind === "vector") {
        debug.vectorCount += 1;
      } else if (item.controlKind === "color") {
        debug.colorCount += 1;
      } else if (item.controlKind === "select") {
        debug.selectCount += 1;
      }

      if (
        debug.sample.length < 20 &&
        (item.controlKind === "vector" || item.controlKind === "color" || item.controlKind === "select")
      ) {
        var sampleEntry = {
          path: item.path,
          name: item.displayName,
          group: item.groupPath,
          kind: item.controlKind,
          value: item.value,
          vectorScale: item.vectorScale || null,
          vectorMode: item.vectorMode || null
        };

        if (item.controlKind === "color" && firstComponent && firstComponent.properties) {
          // // Include raw color API/value snapshots to troubleshoot host channel-order inconsistencies.
          var sampleProperty = subcreator_find_property_by_path(firstComponent.properties, item.path);
          if (sampleProperty) {
            try {
              var sampleColorApiValue =
                typeof sampleProperty.getColorValue === "function" ? sampleProperty.getColorValue() : "<no getColorValue>";
              sampleEntry.colorApiRaw =
                typeof sampleColorApiValue === "string" ? sampleColorApiValue : JSON.stringify(sampleColorApiValue);
            } catch (sampleColorApiError) {
              sampleEntry.colorApiRaw = "<error " + String(sampleColorApiError) + ">";
            }

            try {
              var sampleRawValue = typeof sampleProperty.getValue === "function" ? sampleProperty.getValue() : "<no getValue>";
              sampleEntry.valueRaw = typeof sampleRawValue === "string" ? sampleRawValue : JSON.stringify(sampleRawValue);
            } catch (sampleRawError) {
              sampleEntry.valueRaw = "<error " + String(sampleRawError) + ">";
            }
          }
        }

        debug.sample.push(sampleEntry);
      }
    }

    return subcreator_ok({
      selectedCount: mogrtItems.length,
      editableCount: properties.length,
      properties: properties,
      debug: debug
    });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_apply_selected_mogrt_properties(payloadEncoded) {
  // // Apply visual property changes from panel payload to each selected MOGRT clip.
  try {
    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence in Premiere.");
    }

    var decodedPayload = subcreator_decode_payload(payloadEncoded || "");
    var payload = JSON.parse(decodedPayload || "{}");
    var changes = payload && payload.changes && typeof payload.changes.length === "number" ? payload.changes : [];
    var sequence = app.project.activeSequence;
    var selectedItems = subcreator_get_selected_track_items(sequence);
    var mogrtItems = [];
    for (var itemIndex = 0; itemIndex < selectedItems.length; itemIndex += 1) {
      var trackItem = selectedItems[itemIndex];
      if (subcreator_get_mogrt_component_from_track_item(trackItem)) {
        mogrtItems.push(trackItem);
      }
    }

    if (!mogrtItems.length) {
      return subcreator_ok({
        selectedCount: 0,
        updatedCount: 0,
        failedCount: 0
      });
    }

    var updatedCount = 0;
    var failedCount = 0;
    var debugLines = [];
    var applySequenceSize = subcreator_visual_read_sequence_dimensions();
    debugLines.push("sequence=" + applySequenceSize.width + "x" + applySequenceSize.height);

    for (var clipIndex = 0; clipIndex < mogrtItems.length; clipIndex += 1) {
      var clip = mogrtItems[clipIndex];
      var component = subcreator_get_mogrt_component_from_track_item(clip);
      if (!component || !component.properties) {
        failedCount += changes.length;
        continue;
      }

      for (var changeIndex = 0; changeIndex < changes.length; changeIndex += 1) {
        var change = changes[changeIndex] || {};
        var path = subcreator_trim_string(String(change.path || ""));
        var valueType = subcreator_trim_string(String(change.valueType || "string")).toLowerCase();
        var controlKind = subcreator_trim_string(String(change.controlKind || "")).toLowerCase();
        var virtualTextStyleTarget = subcreator_visual_parse_text_style_virtual_path(path);
        var resolvedPath = virtualTextStyleTarget ? virtualTextStyleTarget.basePath : path;
        var value = change.value;
        var vectorScale = null;
        if (change.vectorScale && Object.prototype.toString.call(change.vectorScale) === "[object Array]") {
          vectorScale = change.vectorScale;
        }
        if (!path) {
          failedCount += 1;
          continue;
        }

        var property = subcreator_find_property_by_path(component.properties, resolvedPath);
        if (!property || typeof property.setValue !== "function") {
          failedCount += 1;
          continue;
        }

        var applied = false;
        var displayName = subcreator_trim_string(String(property.displayName || ""));
        if (virtualTextStyleTarget) {
          displayName += " (" + virtualTextStyleTarget.styleKey + ")";
        }
        if (
          controlKind === "vector" ||
          controlKind === "color" ||
          controlKind === "select" ||
          String(displayName || "").toLowerCase().indexOf("size") !== -1 ||
          !!virtualTextStyleTarget
        ) {
          debugLines.push(
            "change path=" +
              path +
              " name=" +
              displayName +
              " kind=" +
              controlKind +
              " in=" +
              String(value) +
              (virtualTextStyleTarget ? " virtualStyle=" + virtualTextStyleTarget.styleKey : "") +
              (vectorScale ? " scale=" + String(vectorScale) : "")
          );
        }

        if (virtualTextStyleTarget) {
          try {
            applied = subcreator_try_set_mogrt_text_style_property(property, virtualTextStyleTarget.styleKey, value);
          } catch (textStyleError) {}
        } else if (controlKind === "text") {
          try {
            applied = subcreator_try_set_mogrt_text_property(property, String(value || ""));
          } catch (textError) {}
        } else if (controlKind === "color") {
          try {
            applied = subcreator_try_set_mogrt_color_property(property, value);
          } catch (colorError) {}
        } else if (controlKind === "vector") {
          try {
            var parsedVector = subcreator_normalize_visual_payload_value("json", value);
            if (parsedVector && typeof parsedVector.length === "number") {
              var sourceVector = [];
              for (var vectorIndex = 0; vectorIndex < parsedVector.length; vectorIndex += 1) {
                sourceVector.push(Number(parsedVector[vectorIndex]));
              }

              var hostVector = subcreator_visual_vector_to_host_units(sourceVector, vectorScale || [1, 1, 1, 1]);
              property.setValue(hostVector, true);
              debugLines.push("vector out=" + String(hostVector));
              applied = true;
            }
          } catch (vectorError) {}
        }

        if (!applied && controlKind !== "color" && !virtualTextStyleTarget) {
          try {
            var normalizedValue = subcreator_normalize_visual_payload_value(valueType, value);
            property.setValue(normalizedValue, true);
            applied = true;
          } catch (setError) {
            applied = false;
          }
        } else if (!applied && controlKind === "color") {
          debugLines.push("color apply failed without generic setValue fallback");
        } else if (!applied && virtualTextStyleTarget) {
          debugLines.push("text style apply failed without generic setValue fallback");
        }

        if (applied) {
          if (controlKind === "color") {
            try {
              var afterColorValue = typeof property.getColorValue === "function" ? property.getColorValue() : "<no getColorValue>";
              var afterRawValue = typeof property.getValue === "function" ? property.getValue() : "<no getValue>";
              var afterColorText = "";
              var afterRawText = "";
              try {
                afterColorText = typeof afterColorValue === "string" ? afterColorValue : JSON.stringify(afterColorValue);
              } catch (afterColorSerializeError) {
                afterColorText = String(afterColorValue);
              }
              try {
                afterRawText = typeof afterRawValue === "string" ? afterRawValue : JSON.stringify(afterRawValue);
              } catch (afterRawSerializeError) {
                afterRawText = String(afterRawValue);
              }
              var cachedReadLayout = subcreator_visual_get_cached_color_layout(displayName, "read");
              var cachedWriteLayout = subcreator_visual_get_cached_color_layout(displayName, "write");
              debugLines.push("color readback color=" + afterColorText + " raw=" + afterRawText);
              debugLines.push("color layout read=" + String(cachedReadLayout || "<none>") + " write=" + String(cachedWriteLayout || "<none>"));
            } catch (colorReadbackError) {
              debugLines.push("color readback failed: " + String(colorReadbackError));
            }
          }
          updatedCount += 1;
        } else {
          debugLines.push("failed path=" + path + " name=" + displayName + " kind=" + controlKind);
          failedCount += 1;
        }
      }
    }

    var refreshTriggered = subcreator_force_sequence_visual_refresh(sequence);
    debugLines.push("ui_refresh=" + (refreshTriggered ? "forced" : "not_available"));

    return subcreator_ok({
      selectedCount: mogrtItems.length,
      updatedCount: updatedCount,
      failedCount: failedCount,
      debug: debugLines
    });
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_is_default_caption_label(text) {
  // // Detect synthetic/default caption names returned by some Premiere APIs.
  var normalized = subcreator_trim_string(String(text || "")).toLowerCase().replace(/\s+/g, "");
  return normalized === "syntheticcaption";
}

function subcreator_decode_xml_entities(text) {
  // // Decode common XML entities found in Premiere metadata blobs.
  return String(text || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function subcreator_extract_text_from_metadata_blob(metadataText) {
  // // Extract readable caption candidates from XMP/XML metadata payloads.
  var metadata = String(metadataText || "");
  if (!metadata) {
    return "";
  }

  var candidates = [];

  function pushCandidate(value) {
    var normalized = subcreator_trim_string(subcreator_decode_xml_entities(String(value || "")).replace(/\s+/g, " "));
    if (!normalized) {
      return;
    }

    if (subcreator_is_default_caption_label(normalized)) {
      return;
    }

    if (!/[A-Za-z0-9]/.test(normalized)) {
      return;
    }

    if (normalized.length > 300) {
      return;
    }

    var lower = normalized.toLowerCase();
    if (lower.indexOf("http://") === 0 || lower.indexOf("https://") === 0) {
      return;
    }

    for (var index = 0; index < candidates.length; index += 1) {
      if (candidates[index] === normalized) {
        return;
      }
    }

    candidates.push(normalized);
  }

  var prioritizedTagPattern = /<(?:[^>]*)(?:caption|subtitle|transcript|spoken|dialog|text|logcomment)[^>]*>([\s\S]*?)<\/[^>]+>/gi;
  var prioritizedMatch = null;
  while ((prioritizedMatch = prioritizedTagPattern.exec(metadata))) {
    pushCandidate(prioritizedMatch[1]);
  }

  var attributePattern = /\b(?:caption|subtitle|transcript|spoken|dialog|text|logcomment)[\w:-]*\s*=\s*"([^"]+)"/gi;
  var attributeMatch = null;
  while ((attributeMatch = attributePattern.exec(metadata))) {
    pushCandidate(attributeMatch[1]);
  }

  var nodePattern = />([^<]+)</g;
  var nodeMatch = null;
  while ((nodeMatch = nodePattern.exec(metadata))) {
    pushCandidate(nodeMatch[1]);
  }

  var bestCandidate = "";
  var bestScore = -9999;
  for (var candidateIndex = 0; candidateIndex < candidates.length; candidateIndex += 1) {
    var candidate = candidates[candidateIndex];
    var wordCount = candidate.split(/\s+/).filter(Boolean).length;
    var score = wordCount * 4 + Math.min(candidate.length, 120) / 20;
    if (/[.,!?;:]/.test(candidate)) {
      score += 1;
    }

    if (candidate.length < 2) {
      score -= 10;
    }

    if (score > bestScore) {
      bestScore = score;
      bestCandidate = candidate;
    }
  }

  return bestCandidate;
}

function subcreator_extract_text_from_json_payload(payload) {
  // // Read text from JSON payload shapes used by caption and MOGRT controls.
  if (!payload || typeof payload !== "object") {
    return "";
  }

  if (typeof payload.textEditValue === "string" && payload.textEditValue.length > 0) {
    return String(payload.textEditValue);
  }

  if (typeof payload.mText === "string" && payload.mText.length > 0) {
    return String(payload.mText);
  }

  if (payload.styleSheet && typeof payload.styleSheet === "object" && typeof payload.styleSheet.mText === "string") {
    return String(payload.styleSheet.mText);
  }

  if (payload.mStyleSheet && typeof payload.mStyleSheet === "object" && typeof payload.mStyleSheet.mText === "string") {
    return String(payload.mStyleSheet.mText);
  }

  if (payload.mTextParam && typeof payload.mTextParam === "object") {
    var nestedText = subcreator_extract_text_from_json_payload(payload.mTextParam);
    if (nestedText) {
      return nestedText;
    }
  }

  return "";
}

function subcreator_extract_text_from_property_value(rawValue) {
  // // Convert property values (plain/object/JSON string) into readable caption text.
  if (rawValue === undefined || rawValue === null) {
    return "";
  }

  if (typeof rawValue === "string") {
    var rawText = String(rawValue || "");
    if (rawText.indexOf("{") !== -1) {
      try {
        var parsed = JSON.parse(rawText);
        var parsedText = subcreator_extract_text_from_json_payload(parsed);
        if (parsedText) {
          return parsedText;
        }
      } catch (jsonError) {}
    }

    return rawText;
  }

  if (typeof rawValue === "object") {
    try {
      var payloadText = subcreator_extract_text_from_json_payload(rawValue);
      if (payloadText) {
        return payloadText;
      }
    } catch (payloadError) {}

    try {
      var serialized = JSON.stringify(rawValue);
      if (serialized && serialized.indexOf("{") !== -1) {
        var parsedSerialized = JSON.parse(serialized);
        var extracted = subcreator_extract_text_from_json_payload(parsedSerialized);
        if (extracted) {
          return extracted;
        }
      }
    } catch (serializeError) {}
  }

  return "";
}

function subcreator_extract_text_from_component_properties(propertyCollection) {
  // // Traverse component properties recursively to find editable caption text fields.
  if (!propertyCollection || typeof propertyCollection.numItems !== "number") {
    return "";
  }

  var syntheticFallback = "";

  for (var i = 0; i < propertyCollection.numItems; i += 1) {
    var property = propertyCollection[i];
    if (!property) {
      continue;
    }

    if (typeof property.getValue === "function") {
      try {
        var rawValue = property.getValue();
        if (subcreator_should_try_text_property(property.displayName || "", rawValue)) {
          var extracted = subcreator_extract_text_from_property_value(rawValue);
          var normalized = subcreator_trim_string(String(extracted || "").replace(/\r/g, "\n"));
          if (normalized) {
            if (!subcreator_is_default_caption_label(normalized)) {
              return normalized;
            }

            if (!syntheticFallback) {
              syntheticFallback = normalized;
            }
          }
        }
      } catch (propertyValueError) {}
    }

    if (property.properties && typeof property.properties.numItems === "number" && property.properties.numItems > 0) {
      var nested = subcreator_extract_text_from_component_properties(property.properties);
      if (nested) {
        return nested;
      }
    }
  }

  return syntheticFallback;
}

function subcreator_extract_text_from_item_components(item) {
  // // Try to read caption text from component/control payloads on track items.
  if (!item) {
    return "";
  }

  var component = null;
  try {
    if (typeof item.getMGTComponent === "function") {
      component = item.getMGTComponent();
    }
  } catch (mgtError) {}

  if (!component && item.components && item.components.numItems > 0) {
    component = item.components[0];
  }

  if (!component || !component.properties) {
    return "";
  }

  return subcreator_extract_text_from_component_properties(component.properties);
}

function subcreator_extract_text_from_item(item) {
  // // Read caption text from known item methods/properties.
  if (!item) {
    return "";
  }

  function rememberTextCandidate(value, state) {
    var normalizedValue = subcreator_trim_string(String(value || "").replace(/\r/g, "\n"));
    if (!normalizedValue) {
      return "";
    }

    if (!subcreator_is_default_caption_label(normalizedValue)) {
      return normalizedValue;
    }

    if (!state.syntheticFallback) {
      state.syntheticFallback = normalizedValue;
    }

    return "";
  }

  function extractTextFromUnknownValue(rawValue, state) {
    if (rawValue === undefined || rawValue === null) {
      return "";
    }

    if (typeof rawValue === "string") {
      return rememberTextCandidate(rawValue, state);
    }

    if (typeof rawValue === "object") {
      var fromPayload = subcreator_extract_text_from_property_value(rawValue);
      var rememberPayload = rememberTextCandidate(fromPayload, state);
      if (rememberPayload) {
        return rememberPayload;
      }

      if (typeof rawValue.text !== "undefined") {
        var fromTextField = rememberTextCandidate(rawValue.text, state);
        if (fromTextField) {
          return fromTextField;
        }
      }

      try {
        for (var valueKey in rawValue) {
          if (!rawValue.hasOwnProperty(valueKey)) {
            continue;
          }

          var keyValue = rawValue[valueKey];
          if (keyValue === undefined || keyValue === null) {
            continue;
          }

          var normalizedKey = String(valueKey || "").toLowerCase();
          var keyLooksTextual =
            normalizedKey.indexOf("text") !== -1 ||
            normalizedKey.indexOf("caption") !== -1 ||
            normalizedKey.indexOf("subtitle") !== -1 ||
            normalizedKey.indexOf("transcript") !== -1 ||
            normalizedKey.indexOf("content") !== -1 ||
            normalizedKey.indexOf("comment") !== -1 ||
            normalizedKey.indexOf("metadata") !== -1;

          if (typeof keyValue === "string") {
            if (keyLooksTextual || keyValue.indexOf(" ") !== -1) {
              var keyString = rememberTextCandidate(keyValue, state);
              if (keyString) {
                return keyString;
              }
            }
            continue;
          }

          if (typeof keyValue === "object" && keyLooksTextual) {
            var nestedString = extractTextFromUnknownValue(keyValue, state);
            if (nestedString) {
              return nestedString;
            }
          }
        }
      } catch (rawValueKeyError) {}
    }

    return "";
  }

  function extractTextViaReflection(rawItem, state) {
    if (!rawItem || !rawItem.reflect) {
      return "";
    }

    try {
      if (rawItem.reflect.methods && typeof rawItem.reflect.methods.length === "number") {
        for (var methodIdx = 0; methodIdx < rawItem.reflect.methods.length; methodIdx += 1) {
          var reflectedMethod = rawItem.reflect.methods[methodIdx];
          var methodName = reflectedMethod ? String(reflectedMethod.name || "") : "";
          if (!methodName) {
            continue;
          }

          var methodKey = methodName.toLowerCase();
          if (
            methodKey.indexOf("text") === -1 &&
            methodKey.indexOf("caption") === -1 &&
            methodKey.indexOf("transcript") === -1 &&
            methodKey.indexOf("content") === -1 &&
            methodKey.indexOf("comment") === -1 &&
            methodKey.indexOf("metadata") === -1 &&
            methodKey.indexOf("name") === -1 &&
            methodKey.indexOf("get") !== 0
          ) {
            continue;
          }

          try {
            if (typeof rawItem[methodName] === "function") {
              var reflectedValue = rawItem[methodName]();
              var reflectedText = extractTextFromUnknownValue(reflectedValue, state);
              if (reflectedText) {
                return reflectedText;
              }
            }
          } catch (reflectedMethodError) {}
        }
      }
    } catch (reflectMethodsError) {}

    try {
      if (rawItem.reflect.properties && typeof rawItem.reflect.properties.length === "number") {
        for (var propIdx = 0; propIdx < rawItem.reflect.properties.length; propIdx += 1) {
          var reflectedProp = rawItem.reflect.properties[propIdx];
          var propName = reflectedProp ? String(reflectedProp.name || "") : "";
          if (!propName) {
            continue;
          }

          var propKey = propName.toLowerCase();
          if (
            propKey.indexOf("text") === -1 &&
            propKey.indexOf("caption") === -1 &&
            propKey.indexOf("transcript") === -1 &&
            propKey.indexOf("content") === -1 &&
            propKey.indexOf("comment") === -1 &&
            propKey.indexOf("metadata") === -1 &&
            propKey.indexOf("name") === -1
          ) {
            continue;
          }

          try {
            if (typeof rawItem[propName] !== "undefined") {
              var reflectedPropText = extractTextFromUnknownValue(rawItem[propName], state);
              if (reflectedPropText) {
                return reflectedPropText;
              }
            }
          } catch (reflectedPropError) {}
        }
      }
    } catch (reflectPropsError) {}

    return "";
  }

  var syntheticFallback = "";
  var localState = { syntheticFallback: "" };
  var methodNames = [
    "getCaptionText",
    "getText",
    "getSourceText",
    "getFormattedText",
    "getTranscriptText",
    "getComment",
    "getMetadata",
    "getProjectMetadata",
    "getXMPMetadata"
  ];
  for (var methodIndex = 0; methodIndex < methodNames.length; methodIndex += 1) {
    var methodName = methodNames[methodIndex];
    try {
      if (typeof item[methodName] === "function") {
        var methodText = rememberTextCandidate(item[methodName](), localState);
        if (methodText) {
          return methodText;
        }
      }
    } catch (methodError) {}
  }

  if (localState.syntheticFallback) {
    syntheticFallback = localState.syntheticFallback;
  }

  var reflectedText = extractTextViaReflection(item, localState);
  if (reflectedText) {
    return reflectedText;
  }

  if (!syntheticFallback && localState.syntheticFallback) {
    syntheticFallback = localState.syntheticFallback;
  }

  var componentText = subcreator_extract_text_from_item_components(item);
  if (componentText) {
    if (!subcreator_is_default_caption_label(componentText)) {
      return componentText;
    }

    if (!syntheticFallback) {
      syntheticFallback = componentText;
    }
  }

  var propNames = ["captionText", "sourceText", "subtitleText", "text", "value"];
  for (var propIndex = 0; propIndex < propNames.length; propIndex += 1) {
    var propName = propNames[propIndex];
    try {
      if (typeof item[propName] !== "undefined") {
        var propText = subcreator_trim_string(String(item[propName] || "").replace(/\r/g, "\n"));
        if (propText) {
          if (!subcreator_is_default_caption_label(propText)) {
            return propText;
          }

          if (!syntheticFallback) {
            syntheticFallback = propText;
          }
        }
      }
    } catch (propError) {}
  }

  try {
    if (item.projectItem && typeof item.projectItem.getProjectMetadata === "function") {
      var metadata = String(item.projectItem.getProjectMetadata() || "");
      var metadataText = subcreator_extract_text_from_metadata_blob(metadata);
      if (metadataText) {
        return metadataText;
      }
    }
  } catch (metadataError) {}

  try {
    if (item.projectItem && typeof item.projectItem.getXMPMetadata === "function") {
      var xmpMetadata = String(item.projectItem.getXMPMetadata() || "");
      var xmpText = subcreator_extract_text_from_metadata_blob(xmpMetadata);
      if (xmpText) {
        return xmpText;
      }
    }
  } catch (xmpMetadataError) {}

  if (item.projectItem && item.projectItem.name) {
    var projectItemName = subcreator_trim_string(String(item.projectItem.name || ""));
    if (projectItemName) {
      if (!subcreator_is_default_caption_label(projectItemName)) {
        return projectItemName;
      }

      if (!syntheticFallback) {
        syntheticFallback = projectItemName;
      }
    }
  }

  if (item.name) {
    var itemName = subcreator_trim_string(String(item.name || ""));
    if (itemName) {
      if (!subcreator_is_default_caption_label(itemName)) {
        return itemName;
      }

      if (!syntheticFallback) {
        syntheticFallback = itemName;
      }
    }
  }

  return syntheticFallback;
}

function subcreator_drop_synthetic_cues(cues) {
  // // Keep cues that expose readable text and drop synthetic placeholder labels.
  var filtered = [];
  for (var index = 0; index < cues.length; index += 1) {
    var cue = cues[index];
    if (!cue) {
      continue;
    }

    if (!subcreator_is_default_caption_label(cue.text || "")) {
      filtered.push(cue);
    }
  }

  return filtered;
}

function subcreator_collect_track_items(track) {
  // // Collect caption items from multiple possible collection properties.
  var items = [];

  function appendCollection(collection) {
    var values = subcreator_collection_to_array(collection);
    for (var i = 0; i < values.length; i += 1) {
      items.push(values[i]);
    }
  }

  appendCollection(track ? track.clips : null);
  appendCollection(track ? track.items : null);
  appendCollection(track ? track.captions : null);

  return items;
}

function subcreator_extract_cues_from_items(items) {
  // // Convert caption-like track items into generic cue payloads.
  var cues = [];

  for (var i = 0; i < items.length; i += 1) {
    var item = items[i];
    var startSeconds = subcreator_to_seconds(item.start || item.inPoint || item.startTime);
    var endSeconds = subcreator_to_seconds(item.end || item.outPoint || item.endTime);
    var text = subcreator_trim_string(String(subcreator_extract_text_from_item(item) || "").replace(/\s+/g, " "));

    if (isNaN(startSeconds) || isNaN(endSeconds) || endSeconds <= startSeconds || !text) {
      continue;
    }

    cues.push({
      text: text,
      startSeconds: startSeconds,
      endSeconds: endSeconds
    });
  }

  cues.sort(function (left, right) {
    if (left.startSeconds < right.startSeconds) {
      return -1;
    }
    if (left.startSeconds > right.startSeconds) {
      return 1;
    }
    return 0;
  });

  return cues;
}

function subcreator_extract_active_caption_track() {
  // // Try to read cues from the active caption track of current sequence.
  try {
    if (!app || !app.project || !app.project.activeSequence) {
      return subcreator_error("No active sequence in Premiere.");
    }

    var sequence = app.project.activeSequence;

    if (sequence.captionTracks) {
      var tracks = subcreator_collection_to_array(sequence.captionTracks);
      if (tracks.length > 0) {
        var selectedTrack = tracks[0];

        for (var i = 0; i < tracks.length; i += 1) {
          var track = tracks[i];
          try {
            if (
              (typeof track.isTargeted === "function" && track.isTargeted()) ||
              (typeof track.isActive === "function" && track.isActive()) ||
              track.targeted === true ||
              track.active === true
            ) {
              selectedTrack = track;
              break;
            }
          } catch (trackStateError) {}
        }

        var trackItems = subcreator_collect_track_items(selectedTrack);
        var selectedTrackCues = subcreator_drop_synthetic_cues(subcreator_extract_cues_from_items(trackItems));
        if (selectedTrackCues.length > 0) {
          return subcreator_ok(selectedTrackCues);
        }

        var bestTrackCues = [];
        for (var trackIndex = 0; trackIndex < tracks.length; trackIndex += 1) {
          var candidateTrack = tracks[trackIndex];
          var candidateItems = subcreator_collect_track_items(candidateTrack);
          var candidateCues = subcreator_drop_synthetic_cues(subcreator_extract_cues_from_items(candidateItems));
          if (candidateCues.length > bestTrackCues.length) {
            bestTrackCues = candidateCues;
          }
        }

        if (bestTrackCues.length > 0) {
          return subcreator_ok(bestTrackCues);
        }
      }
    }

    // // Fallback: try selected timeline items when captionTracks API is unavailable.
    if (typeof sequence.getSelection === "function") {
      var selection = subcreator_collection_to_array(sequence.getSelection());
      var selectionCues = subcreator_drop_synthetic_cues(subcreator_extract_cues_from_items(selection));
      if (selectionCues.length > 0) {
        return subcreator_ok(selectionCues);
      }
    }

    return subcreator_error(
      "Impossible de lire un texte caption exploitable (Premiere renvoie uniquement des labels SyntheticCaption via cette API CEP). Selectionne les clips caption ou utilise la source SRT."
    );
  } catch (error) {
    return subcreator_error(error);
  }
}

function subcreator_should_try_text_property(displayName, rawValue) {
  // // Identify likely text controls from label and raw value shape.
  var key = String(displayName || "").toLowerCase();
  if (
    key.indexOf("source text") !== -1 ||
    key.indexOf("texte source") !== -1 ||
    key.indexOf("caption text") !== -1 ||
    key.indexOf("subtitle text") !== -1 ||
    key === "text" ||
    key === "texte"
  ) {
    return true;
  }

  var raw = String(rawValue || "");
  return (
    raw.indexOf("\"textEditValue\"") !== -1 ||
    raw.indexOf("\"mText\"") !== -1 ||
    raw.indexOf("\"fontTextRunLength\"") !== -1
  );
}

function subcreator_try_set_json_text_payload(payload, textValue) {
  // // Update known text fields in MOGRT JSON payloads.
  if (!payload || typeof payload !== "object") {
    return false;
  }

  var updated = false;

  if (typeof payload.textEditValue !== "undefined") {
    payload.textEditValue = textValue;
    updated = true;
  }

  if (payload.styleSheet && typeof payload.styleSheet === "object" && typeof payload.styleSheet.mText !== "undefined") {
    payload.styleSheet.mText = textValue;
    updated = true;
  }

  if (payload.mStyleSheet && typeof payload.mStyleSheet === "object" && typeof payload.mStyleSheet.mText !== "undefined") {
    payload.mStyleSheet.mText = textValue;
    updated = true;
  }

  if (payload.mTextParam && typeof payload.mTextParam === "object") {
    if (typeof payload.mTextParam.mText !== "undefined") {
      payload.mTextParam.mText = textValue;
      updated = true;
    }
    if (
      payload.mTextParam.mStyleSheet &&
      typeof payload.mTextParam.mStyleSheet === "object" &&
      typeof payload.mTextParam.mStyleSheet.mText !== "undefined"
    ) {
      payload.mTextParam.mStyleSheet.mText = textValue;
      updated = true;
    }
  }

  if (typeof payload.fontTextRunLength !== "undefined") {
    payload.fontTextRunLength = [String(textValue).length];
    updated = true;
  }

  return updated;
}

function subcreator_normalize_caption_text(textValue) {
  // // Match Premiere text payload conventions by using CR line-breaks.
  return String(textValue || "").replace(/\r\n/g, "\n").replace(/\n/g, "\r");
}

function subcreator_try_patch_text_json_string(rawValue, textValue) {
  // // Patch text fields directly in JSON-like strings when parsing is not supported.
  var raw = String(rawValue || "");
  if (raw.length < 1) {
    return "";
  }

  var escapedText = JSON.stringify(String(textValue || ""));
  var patched = raw;

  patched = patched.replace(/"textEditValue"\s*:\s*"([^"\\]|\\.)*"/g, '"textEditValue":' + escapedText);
  patched = patched.replace(/"mText"\s*:\s*"([^"\\]|\\.)*"/g, '"mText":' + escapedText);
  patched = patched.replace(/"fontTextRunLength"\s*:\s*\[[^\]]*\]/g, '"fontTextRunLength":[' + String(textValue).length + "]");

  if (patched === raw) {
    return "";
  }

  return patched;
}

function subcreator_try_set_mogrt_text_property(property, textValue) {
  // // Apply text to a property, supporting strings, JSON strings, and object payloads.
  var displayName = property.displayName || "";
  var rawValue = "";

  if (typeof property.getValue === "function") {
    try {
      rawValue = property.getValue();
    } catch (getError) {
      rawValue = "";
    }
  }

  if (!subcreator_should_try_text_property(displayName, rawValue)) {
    return false;
  }

  var textString = subcreator_normalize_caption_text(textValue);

  if (rawValue && typeof rawValue === "object") {
    try {
      var objectCopy = JSON.parse(JSON.stringify(rawValue));
      if (subcreator_try_set_json_text_payload(objectCopy, textString)) {
        property.setValue(objectCopy, true);
        return true;
      }
    } catch (objectJsonError) {}

    try {
      if (subcreator_try_set_json_text_payload(rawValue, textString)) {
        property.setValue(rawValue, true);
        return true;
      }
    } catch (objectDirectError) {}
  }

  if (typeof rawValue === "string" && rawValue.indexOf("{") !== -1) {
    try {
      var parsed = JSON.parse(rawValue);
      if (subcreator_try_set_json_text_payload(parsed, textString)) {
        property.setValue(JSON.stringify(parsed), true);
        return true;
      }
    } catch (jsonError) {}

    try {
      var patchedRaw = subcreator_try_patch_text_json_string(rawValue, textString);
      if (patchedRaw) {
        property.setValue(patchedRaw, true);
        return true;
      }
    } catch (patchError) {}
  }

  try {
    property.setValue(textString, true);
    return true;
  } catch (setError) {}

  return false;
}

function subcreator_try_set_animation_mode_property(property, animationMode) {
  // // Drive common MOGRT controls like "Animation" and "Highlight Based On".
  if (!property || typeof property.setValue !== "function") {
    return false;
  }

  var mode = String(animationMode || "").toLowerCase();
  if (mode !== "word" && mode !== "line" && mode !== "none") {
    return false;
  }

  var key = String(property.displayName || "").toLowerCase();
  if (key.indexOf("highlight based on") !== -1 || key.indexOf("based on") !== -1) {
    if (mode === "none") {
      return false;
    }

    // // Most MOGRT menu controls are 1-based: 1=Words, 2=Lines.
    var highlightValue = mode === "word" ? 1 : 2;
    try {
      property.setValue(highlightValue, true);
      return true;
    } catch (highlightError) {}
  }

  if (key === "animation") {
    try {
      property.setValue(mode !== "none", true);
      return true;
    } catch (animError) {}
  }

  return false;
}

function subcreator_try_set_layout_property(property, styleConfig) {
  // // Apply layout controls (characters/lines/font size) when available in a template.
  if (!property || typeof property.setValue !== "function" || !styleConfig) {
    return false;
  }

  var key = String(property.displayName || "").toLowerCase();
  var maxChars = Number(styleConfig.maxCharsPerLine || 0);
  var maxLines = Number(styleConfig.linesPerCaption || 0);
  var fontSize = Number(styleConfig.fontSize || 0);

  if ((key.indexOf("character") !== -1 && key.indexOf("line") !== -1) || key.indexOf("chars per line") !== -1) {
    if (!isNaN(maxChars) && maxChars > 0) {
      try {
        property.setValue(maxChars, true);
        return true;
      } catch (charsError) {}
    }
  }

  if ((key.indexOf("max") !== -1 && key.indexOf("line") !== -1) || key.indexOf("lines per") !== -1) {
    if (!isNaN(maxLines) && maxLines > 0) {
      try {
        property.setValue(maxLines, true);
        return true;
      } catch (linesError) {}
    }
  }

  if (key.indexOf("font size") !== -1 || key.indexOf("taille") !== -1) {
    if (!isNaN(fontSize) && fontSize > 0) {
      try {
        property.setValue(fontSize, true);
        return true;
      } catch (fontSizeError) {}
    }
  }

  return false;
}

function subcreator_try_set_controls_recursively(propertyCollection, textValue, animationMode, styleConfig, stats) {
  // // Traverse nested Essential Graphics property groups and apply text/animation updates.
  if (!propertyCollection || typeof propertyCollection.numItems !== "number") {
    return;
  }

  for (var i = 0; i < propertyCollection.numItems; i += 1) {
    var property = propertyCollection[i];
    if (!property) {
      continue;
    }

    if (typeof property.setValue === "function") {
      if (subcreator_try_set_mogrt_text_property(property, textValue)) {
        stats.textUpdates += 1;
      }

      if (subcreator_try_set_animation_mode_property(property, animationMode)) {
        stats.animationUpdates += 1;
      }

      if (subcreator_try_set_layout_property(property, styleConfig)) {
        stats.layoutUpdates += 1;
      }
    }

    if (property.properties && typeof property.properties.numItems === "number" && property.properties.numItems > 0) {
      subcreator_try_set_controls_recursively(property.properties, textValue, animationMode, styleConfig, stats);
    }
  }
}

function subcreator_try_set_mogrt_controls(trackItem, textValue, animationMode, styleConfig) {
  // // Update text + animation related controls on inserted MOGRT components.
  if (!trackItem || !textValue) {
    return {
      textUpdates: 0,
      animationUpdates: 0,
      layoutUpdates: 0
    };
  }

  var component = null;
  if (typeof trackItem.getMGTComponent === "function") {
    component = trackItem.getMGTComponent();
  }

  if (!component && trackItem.components && trackItem.components.numItems > 0) {
    component = trackItem.components[0];
  }

  if (!component || !component.properties || component.properties.numItems < 1) {
    return {
      textUpdates: 0,
      animationUpdates: 0,
      layoutUpdates: 0
    };
  }

  var stats = {
    textUpdates: 0,
    animationUpdates: 0,
    layoutUpdates: 0
  };
  subcreator_try_set_controls_recursively(component.properties, textValue, animationMode, styleConfig, stats);
  return stats;
}

function subcreator_resolve_extension_root() {
  // // Resolve extension root from current host script location.
  var scriptFile = new File($.fileName);
  if (!scriptFile || !scriptFile.exists) {
    return "";
  }

  return scriptFile.parent.parent.fsName;
}

function subcreator_find_first_mogrt_in_folder(folderRef) {
  // // Recursively return first .mogrt file path found under a folder.
  if (!folderRef || !folderRef.exists) {
    return "";
  }

  var entries = folderRef.getFiles();
  if (!entries || entries.length < 1) {
    return "";
  }

  for (var i = 0; i < entries.length; i += 1) {
    var entry = entries[i];
    if (entry instanceof File) {
      if (/\.mogrt$/i.test(String(entry.name || ""))) {
        return String(entry.fsName || "");
      }
      continue;
    }

    if (entry instanceof Folder) {
      var nested = subcreator_find_first_mogrt_in_folder(entry);
      if (nested && nested.length > 0) {
        return nested;
      }
    }
  }

  return "";
}

function subcreator_resolve_mogrt_path(options) {
  // // Prioritize explicit absolute path, then bundled template path, then first bundled fallback.
  var manualPath = options.mogrtPath || "";
  if (manualPath && manualPath.length > 0) {
    var manualVariants = [
      String(manualPath),
      decodeURI(String(manualPath)),
      String(manualPath).replace(/\\/g, "/")
    ];

    for (var variantIndex = 0; variantIndex < manualVariants.length; variantIndex += 1) {
      var manualFile = new File(manualVariants[variantIndex]);
      if (manualFile.exists) {
        return manualFile.fsName;
      }
    }
  }

  var extensionRoot = options.extensionRootPath || subcreator_resolve_extension_root();
  var templateRelativePath = options.mogrtTemplateRelativePath || "";
  if (templateRelativePath && templateRelativePath.length > 0) {
    if (extensionRoot && extensionRoot.length > 0) {
      var normalizedRelative = String(templateRelativePath).replace(/\\/g, "/");
      var bundledTemplate = new File(extensionRoot + "/templates/mogrt/" + normalizedRelative);
      if (bundledTemplate.exists) {
        return bundledTemplate.fsName;
      }
    }
  }

  if (extensionRoot && extensionRoot.length > 0) {
    // // Hard fallback when UI does not pass template path: take first bundled template.
    var templateFolder = new Folder(extensionRoot + "/templates/mogrt");
    var discoveredTemplate = subcreator_find_first_mogrt_in_folder(templateFolder);
    if (discoveredTemplate && discoveredTemplate.length > 0) {
      return discoveredTemplate;
    }
  }

  return "";
}

function subcreator_seconds_to_ticks(seconds) {
  // // Convert seconds to Premiere ticks for importMGT API.
  try {
    var time = new Time();
    time.seconds = Number(seconds);
    return String(time.ticks);
  } catch (error) {
    return String(Math.round(Number(seconds) * 254016000000));
  }
}

function subcreator_push_unique_path(list, value) {
  // // Keep only distinct non-empty candidate paths for import attempts.
  if (!value) {
    return;
  }

  var normalized = String(value);
  for (var i = 0; i < list.length; i += 1) {
    if (list[i] === normalized) {
      return;
    }
  }

  list.push(normalized);
}

function subcreator_build_mogrt_path_candidates(mogrtPath) {
  // // Build multiple path formats to maximize importMGT compatibility across OS versions.
  var candidates = [];
  if (!mogrtPath) {
    return candidates;
  }

  var fileRef = new File(mogrtPath);
  subcreator_push_unique_path(candidates, fileRef.fsName);
  subcreator_push_unique_path(candidates, fileRef.fullName);
  subcreator_push_unique_path(candidates, decodeURI(String(fileRef.fullName || "")));
  subcreator_push_unique_path(candidates, String(fileRef.fsName || "").replace(/\\/g, "/"));
  subcreator_push_unique_path(candidates, String(fileRef.fullName || "").replace(/\\/g, "/"));
  return candidates;
}

function subcreator_try_import_mogrt(sequence, pathCandidates, startSeconds, videoTrackIndex, audioTrackIndex) {
  // // Try importMGT with both tick and second timing modes and multiple path formats.
  var importResult = {
    trackItem: null,
    attempted: 0,
    usedPath: "",
    usedTimeMode: ""
  };

  var startTicks = subcreator_seconds_to_ticks(startSeconds);
  var timeModes = [
    { mode: "ticks", value: startTicks },
    { mode: "seconds", value: Number(startSeconds) }
  ];

  for (var pathIndex = 0; pathIndex < pathCandidates.length; pathIndex += 1) {
    var pathCandidate = pathCandidates[pathIndex];

    for (var timeIndex = 0; timeIndex < timeModes.length; timeIndex += 1) {
      var timeMode = timeModes[timeIndex];
      importResult.attempted += 1;

      try {
        var insertedItem = sequence.importMGT(pathCandidate, timeMode.value, videoTrackIndex, audioTrackIndex);
        if (insertedItem) {
          importResult.trackItem = insertedItem;
          importResult.usedPath = pathCandidate;
          importResult.usedTimeMode = timeMode.mode;
          return importResult;
        }
      } catch (importError) {}
    }
  }

  return importResult;
}

function subcreator_try_set_mogrt_duration(trackItem, startSeconds, endSeconds) {
  // // Apply cue-specific end time to imported MOGRT clip when API allows it.
  if (!trackItem) {
    return false;
  }

  var safeStart = Number(startSeconds);
  var safeEnd = Number(endSeconds);
  if (isNaN(safeStart) || isNaN(safeEnd) || safeEnd <= safeStart) {
    return false;
  }

  var applied = false;
  var endTime = null;
  try {
    endTime = new Time();
    endTime.seconds = safeEnd;
  } catch (createTimeError) {
    endTime = null;
  }

  if (endTime) {
    try {
      trackItem.end = endTime;
      applied = true;
    } catch (endAssignError) {}

    try {
      if (trackItem.end && typeof trackItem.end.seconds !== "undefined") {
        trackItem.end.seconds = safeEnd;
        applied = true;
      }
    } catch (endSecondsError) {}
  }

  try {
    if (typeof trackItem.outPoint !== "undefined") {
      var outPointTime = new Time();
      outPointTime.seconds = Math.max(safeEnd - safeStart, 0.01);
      trackItem.outPoint = outPointTime;
      applied = true;
    }
  } catch (outPointError) {}

  return applied;
}

function subcreator_get_video_track_clip_count(track) {
  // // Read clip count in a defensive way across Premiere/QE collection variants.
  if (!track || !track.clips) {
    return 0;
  }

  if (typeof track.clips.numItems === "number") {
    return Number(track.clips.numItems || 0);
  }

  if (typeof track.clips.length === "number") {
    return Number(track.clips.length || 0);
  }

  return 0;
}

function subcreator_find_highest_empty_video_track_index(trackCollection) {
  // // Return the top-most empty track to avoid touching existing media clips.
  if (!trackCollection || typeof trackCollection.numTracks !== "number") {
    return -1;
  }

  var totalTracks = Number(trackCollection.numTracks || 0);
  if (totalTracks < 1) {
    return -1;
  }

  for (var trackIndex = totalTracks - 1; trackIndex >= 0; trackIndex -= 1) {
    var track = trackCollection[trackIndex];
    if (subcreator_get_video_track_clip_count(track) < 1) {
      return trackIndex;
    }
  }

  return -1;
}

function subcreator_get_or_create_top_video_track_index(sequence) {
  // // Reuse top empty track when possible, otherwise append a new top track via QE.
  var currentTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0;
  if (currentTracks > 0) {
    var reusableTopEmpty = subcreator_find_highest_empty_video_track_index(sequence.videoTracks);
    if (reusableTopEmpty >= 0) {
      return {
        index: reusableTopEmpty,
        created: false,
        beforeTracks: currentTracks,
        afterTracks: currentTracks
      };
    }
  }

  var created = false;
  var inserted = false;
  try {
    if (typeof app.enableQE === "function") {
      app.enableQE();
      if (typeof qe !== "undefined" && qe.project && typeof qe.project.getActiveSequence === "function") {
        var qeSequence = qe.project.getActiveSequence();
        if (qeSequence && typeof qeSequence.addTracks === "function") {
          if (!inserted && currentTracks > 0) {
            try {
              // // Append one video track at the top so existing tracks are untouched.
              qeSequence.addTracks(1, currentTracks, 0, 0, 0);
              inserted = true;
            } catch (signatureErrorAppendFull) {}
          }

          if (!inserted && currentTracks > 0) {
            try {
              qeSequence.addTracks(1, currentTracks, 0, 0);
              inserted = true;
            } catch (signatureErrorAppendShort) {}
          }

          if (!inserted && currentTracks > 0) {
            try {
              qeSequence.addTracks(1, currentTracks, 0);
              inserted = true;
            } catch (signatureErrorAppendMinimal) {}
          }

          if (!inserted && currentTracks > 0) {
            try {
              qeSequence.addTracks(1, currentTracks);
              inserted = true;
            } catch (signatureErrorAppendTwoArgs) {}
          }

          try {
            if (!inserted) {
              qeSequence.addTracks(1);
              inserted = true;
            }
          } catch (signatureErrorSingleArg) {}
        }
      }
    }
  } catch (error) {}

  var updatedTracks = sequence && sequence.videoTracks ? Number(sequence.videoTracks.numTracks || 0) : 0;
  created = updatedTracks > currentTracks;
  var highestEmptyAfter = subcreator_find_highest_empty_video_track_index(sequence.videoTracks);
  var fallbackTop = updatedTracks > 0 ? updatedTracks - 1 : 0;

  return {
    index: highestEmptyAfter >= 0 ? highestEmptyAfter : fallbackTop,
    created: created,
    beforeTracks: currentTracks,
    afterTracks: updatedTracks
  };
}

function subcreator_apply_captions(payloadEncoded) {
  // // Insert MOGRT instances or fallback timeline markers from generated caption plan.
  try {
    var payloadText = subcreator_decode_payload(payloadEncoded);
    var payload = JSON.parse(payloadText);

    if (!app || !app.project || !app.project.activeSequence) {
      return JSON.stringify({ ok: false, error: "No active sequence in Premiere." });
    }

    var sequence = app.project.activeSequence;
    var options = payload.options || {};
    var cues = payload.cues || [];

    var mogrtPath = subcreator_resolve_mogrt_path(options);
    var pathCandidates = subcreator_build_mogrt_path_candidates(mogrtPath);
    var hasMogrt = pathCandidates.length > 0;

    var videoTrackInfo = subcreator_get_or_create_top_video_track_index(sequence);
    var videoTrackIndex = videoTrackInfo.index;
    var audioTrackIndex = 0;

    var insertedMogrt = 0;
    var insertedMarkers = 0;
    var updatedText = 0;
    var updatedAnimation = 0;
    var updatedLayout = 0;
    var durationAdjusted = 0;
    var mogrtAttempted = 0;
    var lastImportMode = "";
    var lastImportPath = "";

    for (var i = 0; i < cues.length; i += 1) {
      var cue = cues[i];
      var startSeconds = Number(cue.startSeconds);
      var endSeconds = Number(cue.endSeconds);
      var text = cue.text || "";

      if (hasMogrt && typeof sequence.importMGT === "function") {
        var importAttempt = subcreator_try_import_mogrt(sequence, pathCandidates, startSeconds, videoTrackIndex, audioTrackIndex);
        mogrtAttempted += importAttempt.attempted;
        if (importAttempt.trackItem) {
          insertedMogrt += 1;
          lastImportMode = importAttempt.usedTimeMode;
          lastImportPath = importAttempt.usedPath;
          var controlStats = subcreator_try_set_mogrt_controls(
            importAttempt.trackItem,
            text,
            options.style ? options.style.animationMode : "line",
            options.style || {}
          );
          if (controlStats.textUpdates > 0) {
            updatedText += controlStats.textUpdates;
          }
          if (controlStats.animationUpdates > 0) {
            updatedAnimation += controlStats.animationUpdates;
          }
          if (controlStats.layoutUpdates > 0) {
            updatedLayout += controlStats.layoutUpdates;
          }
          if (subcreator_try_set_mogrt_duration(importAttempt.trackItem, startSeconds, endSeconds)) {
            durationAdjusted += 1;
          }
          continue;
        }
      }

      if (sequence.markers && typeof sequence.markers.createMarker === "function") {
        var marker = sequence.markers.createMarker(startSeconds);
        if (marker) {
          marker.end = endSeconds;
          marker.name = "SubCreator";
          marker.comments = text;
          insertedMarkers += 1;
        }
      }
    }

    return JSON.stringify({
      ok: true,
      totalCues: cues.length,
      insertedMogrt: insertedMogrt,
      insertedMarkers: insertedMarkers,
      mogrtTextUpdated: updatedText,
      mogrtAnimationUpdated: updatedAnimation,
      mogrtLayoutUpdated: updatedLayout,
      mogrtDurationAdjusted: durationAdjusted,
      mogrtUsed: hasMogrt,
      mogrtPathResolved: mogrtPath,
      mogrtPathCandidates: pathCandidates,
      mogrtImportAttempts: mogrtAttempted,
      mogrtLastImportMode: lastImportMode,
      mogrtLastImportPath: lastImportPath,
      videoTrackCreated: videoTrackInfo.created,
      videoTracksBefore: videoTrackInfo.beforeTracks,
      videoTracksAfter: videoTrackInfo.afterTracks,
      videoTrackUsed: videoTrackIndex,
      audioTrackUsed: audioTrackIndex
    });
  } catch (error) {
    return JSON.stringify({ ok: false, error: error.toString() });
  }
}
